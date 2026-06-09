// mobile-cockpit / pwa / app.js — Stage A2 (read-write).
//
// What works in Stage A2:
//   - Load config.json
//   - MSAL.js v4 PKCE auth (silent first; redirect on cache miss)
//   - GET cursor-cockpit/state.json from Graph (read path)
//   - PUT cursor-cockpit/state.json with If-Match ETag (write path,
//     one retry on 412 — mirrors daemon/append-test-session.mjs)
//   - createSession() from the composer (view-new form)
//   - approveSession() / cancelSession() from the detail view
//   - Render sessions in #session-list
//   - Refresh button + auto-refresh every CONFIG.pwa.pollIntervalSeconds
//   - setView('list'|'detail'|'new') with back navigation
//
// What is NOT wired (deferred to later phases):
//   - Follow-up button (placeholder; needs Cursor SDK resume on daemon side)
//   - Push notifications (Phase 2, Power Automate / Teams push)
//   - Service worker / offline cache (Phase 2)
//
// Reference order while reading this file:
//   1. ../design.md §3-§6 — OneDrive Graph schema + ETag conflict resolution
//   2. ../daemon/append-test-session.mjs — canonical write-path blueprint
//   3. ./write-helpers.mjs — pure validators / mergers / id-gen (unit-tested)
//
// Style: vanilla JS, no framework, no bundler. ES2020. Single file. MSAL
// is loaded from ./vendor/msal-browser.min.js (defer-ordered before this).
// Pure helpers live in a sibling ESM module ./write-helpers.mjs; we pull
// them in via dynamic import() inside bootstrap() so this file stays a
// classic script and the MSAL UMD bundle keeps its source-order guarantee.

"use strict";

// =============================================================================
// 0. Build stamp + module-level state
// =============================================================================
//
// BUILD_STAMP is replaced by the deploy script before upload (sed on
// `2026-06-10 00:20 CEST aa9c782`). Keep the string literal — index.html cache-busts on it.
const BUILD_STAMP = "2026-06-10 00:20 CEST aa9c782";

/** Loaded asynchronously from ./config.json at boot. See pwa/config.json. */
let CONFIG = null;

/** Loaded once by initMsal(). Reused for every subsequent token acquisition. */
let msalClient = null;

/** Cached after the first successful sign-in. */
let activeAccount = null;

/** Cached last-known state.json (for instant render on refresh). */
let cachedState = null;

/** Cached driveItem ETag for state.json. Required by PUT (If-Match). */
let cachedStateEtag = null;

/** Auto-refresh handle from setInterval(). */
let refreshTimerId = null;

/**
 * Dynamically imported write-helpers module. Populated by bootstrap()
 * before any user-triggered write path can fire. We do this lazily so the
 * <script> tag for app.js can stay classic (UMD MSAL must load first); a
 * top-level static `import` would force ESM-module ordering and miss MSAL.
 */
let WRITE_HELPERS = null;

/** Dynamically imported pure helpers for the IDE-tabs view (M2.1). */
let IDE_HELPERS = null;

/** Cached last-known ide-tabs.json snapshot. Refreshed by loadIdeTabs(). */
let cachedIdeSnapshot = null;

/** Auto-refresh handle for the ide-tabs view (independent of sessions). */
let ideRefreshTimerId = null;

// =============================================================================
// 1. Auth — MSAL.js v4 PKCE flow
// =============================================================================
//
// Flow:
//   - On boot, instantiate PublicClientApplication.
//   - handleRedirectPromise() to consume any redirect response.
//   - getActiveAccount() — if present, silently acquire a token.
//   - Else, loginRedirect() (mobile-friendly; popups blocked on iOS Safari).
//   - acquireTokenSilent() on every Graph call; fallback to acquireTokenRedirect().
//
// Token cache: localStorage (survives mobile-Edge tab close).

async function initMsal() {
  if (typeof msal === "undefined") {
    throw new Error(
      "msal-browser not loaded — check ./vendor/msal-browser.min.js exists " +
        "and the <script> tag in index.html runs BEFORE app.js (defer order matters)",
    );
  }
  const config = {
    auth: {
      clientId: CONFIG.azure.clientId,
      authority: CONFIG.azure.authority,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  };
  msalClient = new msal.PublicClientApplication(config);
  await msalClient.initialize();
  const redirectResult = await msalClient.handleRedirectPromise();
  if (redirectResult && redirectResult.account) {
    activeAccount = redirectResult.account;
    msalClient.setActiveAccount(activeAccount);
    return;
  }
  const accounts = msalClient.getAllAccounts();
  if (accounts.length > 0) {
    activeAccount = accounts[0];
    msalClient.setActiveAccount(activeAccount);
  }
}

async function ensureSignedIn() {
  if (activeAccount) return activeAccount;
  await msalClient.loginRedirect({ scopes: CONFIG.graph.scopes });
  // loginRedirect navigates away; control does not return.
  throw new Error("loginRedirect did not navigate — unexpected");
}

async function getAccessToken() {
  await ensureSignedIn();
  try {
    const result = await msalClient.acquireTokenSilent({
      scopes: CONFIG.graph.scopes,
      account: activeAccount,
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      await msalClient.acquireTokenRedirect({ scopes: CONFIG.graph.scopes });
      throw new Error("acquireTokenRedirect did not navigate — unexpected");
    }
    throw err;
  }
}

// =============================================================================
// 2. Graph helpers (read + write)
// =============================================================================

async function graphFetch(path, init = {}) {
  const token = await getAccessToken();
  const headers = new Headers(init.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  const res = await fetch(`${CONFIG.graph.base}${path}`, { ...init, headers });
  return res;
}

/**
 * Load the OneDrive state.json. Returns { state, etag } on success.
 * Throws on HTTP errors other than 404 (treats 404 as empty state).
 *
 * Side-effect: refreshes the module-level cachedState / cachedStateEtag.
 * Callers that only need a one-shot snapshot can use the returned object
 * directly; the cache is for instant-render-on-refresh.
 */
async function loadState() {
  // First call: get driveItem (with eTag); second call: content stream.
  const meta = await graphFetch(`${CONFIG.state.endpoint}`);
  if (meta.status === 404) {
    const fresh = { state: { schemaVersion: 1, sessions: [] }, etag: null };
    cachedState = fresh.state;
    cachedStateEtag = fresh.etag;
    return fresh;
  }
  if (!meta.ok) {
    throw new Error(`Graph driveItem GET failed: ${meta.status} ${meta.statusText}`);
  }
  const metaJson = await meta.json();
  const etag = metaJson.eTag || metaJson["@odata.etag"] || null;
  const contentRes = await graphFetch(`${CONFIG.state.endpoint}:/content`);
  if (!contentRes.ok) {
    throw new Error(`Graph content GET failed: ${contentRes.status} ${contentRes.statusText}`);
  }
  const state = await contentRes.json();
  cachedState = state;
  cachedStateEtag = etag;
  return { state, etag };
}

/**
 * Write the OneDrive state.json with optimistic-concurrency If-Match.
 * Returns the parsed driveItem JSON (which includes the fresh eTag).
 * Throws `{ code: "PRECONDITION_FAILED", status: 412 }` on stale ETag —
 * the caller's retry loop is responsible for re-reading and re-merging.
 *
 * Why two-arg signature instead of a single options bag: matches
 * daemon/lib/graph-state.mjs#writeState() shape so the two surfaces stay
 * trivially comparable in code review.
 */
async function putState(stateObj, etagOrNull) {
  const headers = new Headers({ "Content-Type": "application/json" });
  if (etagOrNull) headers.set("If-Match", etagOrNull);
  const body = JSON.stringify(stateObj, null, 2) + "\n";
  const res = await graphFetch(`${CONFIG.state.endpoint}:/content`, {
    method: "PUT",
    body,
    headers,
  });
  if (res.status === 412) {
    const err = new Error("state.json changed since last read (412)");
    err.code = "PRECONDITION_FAILED";
    err.status = 412;
    throw err;
  }
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`putState: PUT failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  return res.json().catch(() => null);
}

/**
 * Load the OneDrive `cursor-cockpit/ide-tabs.json` snapshot written by the
 * read-only mirror daemon (flows/mobile-cockpit/ide-mirror/poll.mjs).
 *
 * GET-only — the PWA never writes this file in M2.1. Returns the parsed
 * snapshot envelope `{schemaVersion, snapshotAt, workspaceKey, workspacePath,
 * tabs[]}`, or an empty stub `{schemaVersion:1, tabs:[]}` on 404 (daemon has
 * not run yet). Refreshes the module-level cachedIdeSnapshot for instant
 * re-render. Throws on any non-404 HTTP error.
 *
 * Skips the eTag dance entirely (no write-back side, no need for
 * If-Match). The daemon writes ~every 10s when changed; the PWA polls
 * `ideTabs.pollIntervalSeconds` (default 20s) -- always read-fresh, never
 * If-None-Match (so we always see the latest write, eTag churn doesn't
 * matter here).
 */
async function loadIdeTabs() {
  const endpoint = CONFIG && CONFIG.ideTabs && CONFIG.ideTabs.endpoint;
  if (!endpoint) {
    throw new Error("config.ideTabs.endpoint missing -- update pwa/config.json");
  }
  const contentRes = await graphFetch(`${endpoint}:/content`);
  if (contentRes.status === 404) {
    const stub = {
      schemaVersion: 1,
      snapshotAt: null,
      workspaceKey: null,
      workspacePath: null,
      tabs: [],
    };
    cachedIdeSnapshot = stub;
    return stub;
  }
  if (!contentRes.ok) {
    throw new Error(`ide-tabs.json GET failed: ${contentRes.status} ${contentRes.statusText}`);
  }
  const snapshot = await contentRes.json();
  cachedIdeSnapshot = snapshot;
  return snapshot;
}

// =============================================================================
// 3. Pure helpers (sortable / testable)
// =============================================================================

/** Return sessions sorted by `lastUpdated` desc, capped to limit. */
function sortSessions(sessions, limit) {
  if (!Array.isArray(sessions)) return [];
  const sorted = [...sessions].sort((a, b) => {
    const aT = Date.parse(a.lastUpdated || a.createdAt || a.created || "") || 0;
    const bT = Date.parse(b.lastUpdated || b.createdAt || b.created || "") || 0;
    return bT - aT;
  });
  return sorted.slice(0, Math.max(0, limit | 0));
}

/** "12 s ago", "4 min ago", "2 h ago", "3 d ago" — for the list row. */
function relativeTime(isoString) {
  if (!isoString) return "?";
  const then = Date.parse(isoString);
  if (Number.isNaN(then)) return "?";
  const delta = Math.max(0, Date.now() - then);
  if (delta < 60_000) return `${Math.round(delta / 1000)} s ago`;
  if (delta < 3_600_000) return `${Math.round(delta / 60_000)} min ago`;
  if (delta < 86_400_000) return `${Math.round(delta / 3_600_000)} h ago`;
  return `${Math.round(delta / 86_400_000)} d ago`;
}

/** Status → CSS class hint. Pure mapping; renderer applies the class. */
function statusClass(status) {
  const known = ["pending", "running", "approved", "done", "failed", "cancelled"];
  return known.includes(status) ? `status-${status}` : "status-unknown";
}

// =============================================================================
// 4. Write actions — read-modify-write with one retry on 412
// =============================================================================
//
// All three actions follow the same shape (mirror of daemon/append-test-session.mjs):
//   1. loadState() to get { state, etag }.
//   2. Pure merge via WRITE_HELPERS (mergeAppendSession / mergeUpdateStatus).
//   3. putState(next, etag) → on 412, retry ONCE: re-load + re-merge + re-put.
//   4. On second 412 → throw; caller surfaces error.
//
// The pure-helper layer (write-helpers.mjs) is unit-tested without DOM;
// these wrappers add the side-effect plumbing (Graph I/O, status badge,
// view transitions) and live ONLY here.

/**
 * Validate + create a new session, persist via PUT-with-If-Match.
 * Returns { sessionId } on success. Throws on validation / network / 412×2.
 */
async function createSession({ prompt, cwd, model }) {
  if (!WRITE_HELPERS) throw new Error("write-helpers module not loaded yet (bootstrap order bug)");
  const allowedCwds = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  const validation = WRITE_HELPERS.validateCreateInputs({ prompt, cwd, model }, allowedCwds);
  if (!validation.valid) {
    throw new Error(`Invalid input: ${validation.errors.join("; ")}`);
  }

  setStatusBadge("saving…", "saving");

  for (let attempt = 0; attempt < 2; attempt++) {
    const { state, etag } = await loadState();
    const sessionId = WRITE_HELPERS.generateSessionId(Date.now(), WRITE_HELPERS.cryptoRandomBytes);
    const nowIso = new Date().toISOString();
    const session = {
      sessionId,
      status: "pending",
      title: (prompt || "").slice(0, 60),
      prompt,
      model: model || null,
      cwd: cwd || null,
      createdAt: nowIso,
      lastUpdated: nowIso,
      createdBy: "pwa/app.js",
    };
    const next = WRITE_HELPERS.mergeAppendSession(state, session);
    try {
      await putState(next, etag);
      flashSavedBadge();
      return { sessionId };
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) {
        setStatusBadge(`save failed: ${err.message}`, "error");
        throw err;
      }
      // fall through to retry — re-read state in next iteration
    }
  }
  // Defensive: should be unreachable; the throw above covers both legs.
  setStatusBadge("save failed: retries exhausted", "error");
  throw new Error("createSession: retries exhausted");
}

/**
 * Approve a pending session: status → "approved" + lastUpdated → now.
 * Same retry-once-on-412 shape as createSession.
 */
async function approveSession(sessionId) {
  if (!WRITE_HELPERS) throw new Error("write-helpers module not loaded yet (bootstrap order bug)");
  return updateSessionStatus(sessionId, "approved");
}

/**
 * Cancel an in-flight session: status → "cancelled" + lastUpdated → now.
 * Same retry-once-on-412 shape as createSession.
 */
async function cancelSession(sessionId) {
  if (!WRITE_HELPERS) throw new Error("write-helpers module not loaded yet (bootstrap order bug)");
  return updateSessionStatus(sessionId, "cancelled");
}

/** Shared retry skeleton for approve / cancel. NOT exported; consumers go
 *  through the typed wrappers above so the daemon-visible status set stays
 *  centralised in this file. */
async function updateSessionStatus(sessionId, newStatus) {
  setStatusBadge("saving…", "saving");
  for (let attempt = 0; attempt < 2; attempt++) {
    const { state, etag } = await loadState();
    const next = WRITE_HELPERS.mergeUpdateStatus(state, sessionId, newStatus, new Date());
    try {
      await putState(next, etag);
      flashSavedBadge();
      return { sessionId, status: newStatus };
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) {
        setStatusBadge(`save failed: ${err.message}`, "error");
        throw err;
      }
    }
  }
  setStatusBadge("save failed: retries exhausted", "error");
  throw new Error(`updateSessionStatus(${sessionId}, ${newStatus}): retries exhausted`);
}

// =============================================================================
// 5. Views
// =============================================================================

function setView(viewId, payload) {
  document.body.dataset.view = viewId;
  for (const section of document.querySelectorAll(".cockpit-view")) {
    section.hidden = section.dataset.viewId !== viewId;
  }
  // Sync the mode-toggle aria-current so the active pill matches the
  // current top-level view (the toggle is hidden by CSS in sub-views, so
  // this only matters when we're in list / ide-tabs).
  for (const btn of document.querySelectorAll(".cockpit-mode-btn")) {
    btn.setAttribute(
      "aria-current",
      btn.dataset.targetView === viewId ? "true" : "false",
    );
  }
  if (viewId === "list") {
    renderList().catch((err) => showListError(err.message));
  } else if (viewId === "detail" && payload && payload.sessionId) {
    renderDetail(payload.sessionId);
  } else if (viewId === "new") {
    renderNew();
  } else if (viewId === "ide-tabs") {
    renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
  } else if (viewId === "ide-tab-detail" && payload && payload.composerId) {
    renderIdeTabDetail(payload.composerId);
  }
}

function showListError(message) {
  const el = document.getElementById("list-error-state");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}

function clearListError() {
  const el = document.getElementById("list-error-state");
  if (el) el.hidden = true;
}

function showDetailError(message) {
  const el = document.getElementById("detail-error-state");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}

function clearDetailError() {
  const el = document.getElementById("detail-error-state");
  if (el) el.hidden = true;
}

function showNewError(message) {
  const el = document.getElementById("new-error-state");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}

function clearNewError() {
  const el = document.getElementById("new-error-state");
  if (el) el.hidden = true;
}

async function renderList() {
  clearListError();
  const ul = document.getElementById("session-list");
  const empty = document.getElementById("list-empty-state");
  if (!ul || !empty) return;

  try {
    await loadState(); // populates cachedState / cachedStateEtag
  } catch (err) {
    showListError(err.message);
    return;
  }

  const sessions = sortSessions(cachedState.sessions, CONFIG.pwa.recentSessionsCount);
  ul.innerHTML = "";
  if (sessions.length === 0) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;
  for (const s of sessions) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.dataset.sessionId = s.sessionId || "";
    li.tabIndex = 0;
    li.addEventListener("click", () => setView("detail", { sessionId: s.sessionId }));
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        setView("detail", { sessionId: s.sessionId });
      }
    });

    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = s.title || s.prompt || "(untitled)";
    li.appendChild(title);

    const status = document.createElement("span");
    status.className = `cockpit-row-status ${statusClass(s.status)}`;
    status.dataset.status = s.status || "unknown";
    status.textContent = s.status || "unknown";
    li.appendChild(status);

    const time = document.createElement("time");
    time.className = "cockpit-row-time";
    const stamp = s.lastUpdated || s.createdAt || s.created;
    if (stamp) time.dateTime = stamp;
    time.textContent = relativeTime(stamp);
    li.appendChild(time);

    ul.appendChild(li);
  }
}

function renderDetail(sessionId) {
  clearDetailError();
  if (!cachedState || !Array.isArray(cachedState.sessions)) return;
  const s = cachedState.sessions.find((x) => x.sessionId === sessionId);
  if (!s) {
    showListError(`Session not found locally: ${sessionId}`);
    setView("list");
    return;
  }
  const set = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text == null ? "—" : String(text);
  };
  set("detail-title", s.title || s.prompt || "(untitled)");
  set("detail-status", s.status);
  set("detail-session-id", s.sessionId);
  set("detail-agent-id", s.cursorAgentId);
  set("detail-model", s.model);
  set("detail-created", s.createdAt || s.created);
  set("detail-updated", s.lastUpdated);
  set("detail-prompt-text", s.prompt);
  set("detail-output-text", s.output);

  // Stage A2: action buttons. Approve only when pending AND autoApprove off.
  // Cancel when status is pending / approved / running (NOT done/cancelled/failed).
  const autoApprove = !!(CONFIG.session && CONFIG.session.autoApprove);
  const btnApprove = document.getElementById("btn-approve");
  const btnCancel = document.getElementById("btn-cancel");
  const btnFollowUp = document.getElementById("btn-follow-up");

  if (btnApprove) {
    const showApprove = s.status === "pending" && !autoApprove;
    btnApprove.hidden = !showApprove;
    btnApprove.disabled = !showApprove;
    btnApprove.onclick = showApprove
      ? () => handleApproveClick(s.sessionId)
      : null;
  }
  if (btnCancel) {
    const cancellable = ["pending", "approved", "running"].includes(s.status);
    btnCancel.hidden = !cancellable;
    btnCancel.disabled = !cancellable;
    btnCancel.onclick = cancellable
      ? () => handleCancelClick(s.sessionId)
      : null;
  }
  if (btnFollowUp) {
    // Follow-up = "send a new prompt that uses Agent.resume(agentId)".
    // Daemon-side feature; PWA Stage A2 leaves the button hidden until the
    // daemon wires resume semantics. Surfaced here so we don't lose track.
    btnFollowUp.hidden = true;
    btnFollowUp.disabled = true;
  }
}

function renderNew() {
  clearNewError();
  populateCwdSelect();
  // Clear stale form state when re-entering the composer.
  const form = document.getElementById("new-session-form");
  if (form) form.reset();
}

function populateCwdSelect() {
  const select = document.getElementById("new-cwd");
  if (!select) return;
  const allowed = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  // Rebuild: first option = "(daemon default)" (empty value), then one
  // <option> per allowedCwd. Idempotent — safe to call on every render.
  select.innerHTML = "";
  const defaultOpt = document.createElement("option");
  defaultOpt.value = "";
  defaultOpt.textContent = "(daemon default)";
  defaultOpt.selected = true;
  select.appendChild(defaultOpt);
  for (const cwd of allowed) {
    const opt = document.createElement("option");
    opt.value = cwd;
    opt.textContent = cwd;
    select.appendChild(opt);
  }
}

// -----------------------------------------------------------------------------
// IDE-tabs view (M2.1, read-only)
// -----------------------------------------------------------------------------

function showIdeTabsError(message) {
  const el = document.getElementById("ide-tabs-error-state");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}
function clearIdeTabsError() {
  const el = document.getElementById("ide-tabs-error-state");
  if (el) el.hidden = true;
}
function showIdeDetailError(message) {
  const el = document.getElementById("ide-detail-error-state");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}
function clearIdeDetailError() {
  const el = document.getElementById("ide-detail-error-state");
  if (el) el.hidden = true;
}

async function renderIdeTabsList() {
  if (!IDE_HELPERS) {
    showIdeTabsError("ide-helpers module not loaded yet (bootstrap order bug)");
    return;
  }
  clearIdeTabsError();
  const ul = document.getElementById("ide-tabs-list");
  const empty = document.getElementById("ide-tabs-empty-state");
  const summary = document.getElementById("ide-summary");
  if (!ul || !empty || !summary) return;

  try {
    await loadIdeTabs(); // populates cachedIdeSnapshot
  } catch (err) {
    showIdeTabsError(err.message);
    return;
  }

  const snapshot = cachedIdeSnapshot || { tabs: [] };
  const tabs = Array.isArray(snapshot.tabs) ? snapshot.tabs : [];
  // Defensive cap; sortIdeTabs also caps, but pulling the limit from config
  // keeps the two paths consistent. recentSessionsCount is reused as the
  // "max tabs to show in the list" cap.
  const cap = (CONFIG.pwa && CONFIG.pwa.recentSessionsCount) || 50;
  const sorted = IDE_HELPERS.sortIdeTabs(tabs, cap);

  const bucket = IDE_HELPERS.summarizeWaitingOn(sorted);
  const snapshotAtRel = IDE_HELPERS.relativeIdeTime(snapshot.snapshotAt, Date.now());
  summary.innerHTML = "";
  const total = document.createElement("span");
  total.className = "cockpit-ide-summary-bucket";
  total.innerHTML = `<strong>${bucket.total}</strong> tab${bucket.total === 1 ? "" : "s"}`;
  summary.appendChild(total);
  if (bucket.agent > 0) {
    const b = document.createElement("span");
    b.className = "cockpit-ide-summary-bucket";
    b.innerHTML = `<strong>${bucket.agent}</strong> agent thinking`;
    summary.appendChild(b);
  }
  if (bucket.user > 0) {
    const b = document.createElement("span");
    b.className = "cockpit-ide-summary-bucket";
    b.innerHTML = `<strong>${bucket.user}</strong> your turn`;
    summary.appendChild(b);
  }
  if (bucket.none > 0) {
    const b = document.createElement("span");
    b.className = "cockpit-ide-summary-bucket";
    b.innerHTML = `<strong>${bucket.none}</strong> idle`;
    summary.appendChild(b);
  }
  const ts = document.createElement("span");
  ts.className = "cockpit-ide-summary-bucket";
  ts.textContent = `mirrored ${snapshotAtRel}`;
  summary.appendChild(ts);

  ul.innerHTML = "";
  if (sorted.length === 0) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;

  for (const t of sorted) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.dataset.composerId = t.composerId || "";
    li.tabIndex = 0;
    li.addEventListener("click", () => setView("ide-tab-detail", { composerId: t.composerId }));
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        setView("ide-tab-detail", { composerId: t.composerId });
      }
    });

    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = IDE_HELPERS.formatTabTitle(t.title, 60);
    li.appendChild(title);

    const status = document.createElement("span");
    status.className = "cockpit-row-status";
    status.dataset.waitingOn = t.waitingOn || "none";
    status.textContent = IDE_HELPERS.waitingOnLabel(t.waitingOn);
    li.appendChild(status);

    const time = document.createElement("time");
    time.className = "cockpit-row-time";
    if (t.lastActivityAt) time.dateTime = t.lastActivityAt;
    time.textContent = IDE_HELPERS.relativeIdeTime(t.lastActivityAt, Date.now());
    li.appendChild(time);

    ul.appendChild(li);
  }
}

function renderIdeTabDetail(composerId) {
  if (!IDE_HELPERS) {
    showIdeDetailError("ide-helpers module not loaded yet (bootstrap order bug)");
    return;
  }
  clearIdeDetailError();
  if (!cachedIdeSnapshot || !Array.isArray(cachedIdeSnapshot.tabs)) {
    showIdeTabsError("No cached IDE snapshot — refresh the IDE-tabs list first.");
    setView("ide-tabs");
    return;
  }
  const tab = cachedIdeSnapshot.tabs.find((x) => x.composerId === composerId);
  if (!tab) {
    showIdeTabsError(`IDE tab not in current snapshot: ${composerId}`);
    setView("ide-tabs");
    return;
  }

  const set = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text == null ? "—" : String(text);
  };
  set("ide-detail-title", IDE_HELPERS.formatTabTitle(tab.title, 120));
  set("ide-detail-waiting", IDE_HELPERS.waitingOnLabel(tab.waitingOn));
  set("ide-detail-composer-id", tab.composerId);
  set("ide-detail-last-activity",
      tab.lastActivityAt
        ? `${tab.lastActivityAt} (${IDE_HELPERS.relativeIdeTime(tab.lastActivityAt, Date.now())})`
        : "—");
  set("ide-detail-message-count", tab.messageCount);
  // Friendly KB formatting; falls back to bytes for sub-1KB.
  if (typeof tab.transcriptSizeBytes === "number" && tab.transcriptSizeBytes >= 0) {
    const kb = tab.transcriptSizeBytes / 1024;
    set("ide-detail-transcript-size",
        kb >= 1
          ? `${kb.toFixed(1)} KB (${tab.transcriptSizeBytes} bytes)`
          : `${tab.transcriptSizeBytes} bytes`);
  } else {
    set("ide-detail-transcript-size", "—");
  }

  const threadOl = document.getElementById("ide-detail-thread");
  const threadEmpty = document.getElementById("ide-detail-thread-empty");
  if (!threadOl || !threadEmpty) return;
  threadOl.innerHTML = "";
  const thread = Array.isArray(tab.thread) ? tab.thread : [];
  if (thread.length === 0) {
    threadEmpty.hidden = false;
    return;
  }
  threadEmpty.hidden = true;
  // Most-recent first — `thread` is chronological in the snapshot so we
  // reverse for the UI. Avoid mutating the source array (cached snapshot).
  const reversed = [...thread].reverse();
  for (const raw of reversed) {
    const entry = IDE_HELPERS.formatThreadEntry(raw);
    const li = document.createElement("li");
    li.className = "cockpit-ide-turn";
    li.dataset.role = entry.role;

    const header = document.createElement("header");
    const label = document.createElement("span");
    label.className = "cockpit-ide-turn-label";
    label.textContent = entry.label;
    header.appendChild(label);
    if (entry.tools.length > 0) {
      const tools = document.createElement("span");
      tools.className = "cockpit-ide-turn-tools";
      tools.textContent = entry.tools.length === 1
        ? `1 tool: ${entry.tools[0]}`
        : `${entry.tools.length} tools: ${entry.tools.slice(0, 3).join(", ")}${entry.tools.length > 3 ? "…" : ""}`;
      header.appendChild(tools);
    }
    li.appendChild(header);

    if (entry.text.length > 0) {
      const p = document.createElement("p");
      p.className = "cockpit-ide-turn-text";
      // Cap on-screen text per turn so long agent responses stay scannable.
      const MAX = 1200;
      p.textContent = entry.text.length > MAX ? entry.text.slice(0, MAX) + "…" : entry.text;
      li.appendChild(p);
    } else if (entry.tools.length === 0) {
      const p = document.createElement("p");
      p.className = "cockpit-ide-turn-empty";
      p.textContent = "(no content)";
      li.appendChild(p);
    }

    threadOl.appendChild(li);
  }
}

// =============================================================================
// 6. Status badge + auto-refresh
// =============================================================================

function setStatusBadge(text, status) {
  const el = document.getElementById("status-badge");
  if (!el) return;
  el.textContent = text;
  el.dataset.status = status;
}

/**
 * Brief "saved" flash after a successful write, then revert to the
 * standard signed-in (read-write) badge. 2-second hold by default.
 */
function flashSavedBadge(holdMs = 2000) {
  setStatusBadge("saved", "ok");
  setTimeout(() => {
    if (activeAccount) {
      setStatusBadge(`signed in: ${activeAccount.username} (read-write)`, "ok");
    }
  }, holdMs);
}

function startAutoRefresh() {
  if (refreshTimerId === null) {
    const intervalMs = Math.max(5, (CONFIG.pwa.pollIntervalSeconds | 0)) * 1000;
    refreshTimerId = setInterval(() => {
      if (document.body.dataset.view === "list") {
        renderList().catch((err) => showListError(err.message));
      }
    }, intervalMs);
  }
  // Independent timer for the read-only IDE-tabs view — typically faster
  // (CONFIG.ideTabs.pollIntervalSeconds = 20 vs 30) because the GET path
  // is lighter (no ETag dance, single content stream, daemon batches
  // writes via fingerprint dedup so PUTs are sparse).
  if (ideRefreshTimerId === null && CONFIG.ideTabs && CONFIG.ideTabs.pollIntervalSeconds) {
    const ideIntervalMs = Math.max(5, (CONFIG.ideTabs.pollIntervalSeconds | 0)) * 1000;
    ideRefreshTimerId = setInterval(() => {
      const v = document.body.dataset.view;
      // Refresh either the list OR the detail (so the open thread stays
      // live if the user is reading it while the agent posts new turns).
      if (v === "ide-tabs") {
        renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
      } else if (v === "ide-tab-detail") {
        // Re-fetch and re-render the same detail; if the tab disappeared
        // from the snapshot, renderIdeTabDetail bounces back to the list.
        loadIdeTabs()
          .then(() => {
            const openId = document.getElementById("ide-detail-composer-id");
            if (openId && openId.textContent) {
              renderIdeTabDetail(openId.textContent);
            }
          })
          .catch((err) => showIdeDetailError(err.message));
      }
    }, ideIntervalMs);
  }
}

// =============================================================================
// 7. UI handlers (button → action glue)
// =============================================================================

async function handleNewSubmit(ev) {
  ev.preventDefault();
  clearNewError();
  const promptEl = document.getElementById("new-prompt");
  const cwdEl = document.getElementById("new-cwd");
  const modelEl = document.getElementById("new-model");
  const submitBtn = document.getElementById("btn-new-submit");
  const prompt = promptEl ? promptEl.value : "";
  const cwd = cwdEl ? cwdEl.value : "";
  const model = modelEl ? modelEl.value : "";

  if (submitBtn) submitBtn.disabled = true;
  try {
    await createSession({ prompt, cwd, model });
    setView("list");
  } catch (err) {
    showNewError(err.message);
  } finally {
    if (submitBtn) submitBtn.disabled = false;
  }
}

async function handleApproveClick(sessionId) {
  clearDetailError();
  try {
    await approveSession(sessionId);
    await renderList();
    renderDetail(sessionId);
  } catch (err) {
    showDetailError(err.message);
  }
}

async function handleCancelClick(sessionId) {
  clearDetailError();
  // Soft confirm — single tap is too easy to miss-click on mobile.
  if (typeof window !== "undefined" && typeof window.confirm === "function") {
    const ok = window.confirm("Cancel this session? The daemon will stop it on its next poll.");
    if (!ok) return;
  }
  try {
    await cancelSession(sessionId);
    await renderList();
    renderDetail(sessionId);
  } catch (err) {
    showDetailError(err.message);
  }
}

// =============================================================================
// 8. Bootstrap
// =============================================================================

async function bootstrap() {
  const buildStampEl = document.getElementById("build-stamp");
  if (buildStampEl) buildStampEl.textContent = BUILD_STAMP;
  const connEl = document.getElementById("conn-state");
  if (connEl) connEl.textContent = "loading…";

  try {
    const configRes = await fetch("./config.json", { cache: "no-store" });
    if (!configRes.ok) {
      throw new Error(`config.json fetch failed: ${configRes.status}`);
    }
    CONFIG = await configRes.json();
  } catch (err) {
    setStatusBadge(`config error: ${err.message}`, "error");
    if (connEl) connEl.textContent = "offline";
    return;
  }

  // Pull in the pure helpers BEFORE we wire any write handlers. Dynamic
  // import keeps this script as a classic <script defer> while letting
  // the helpers live in their own ESM module (so the Node-side unit test
  // can import them cleanly without dragging in MSAL / DOM). The
  // ide-helpers module is sibling; both are loaded in parallel because
  // they have no inter-dependency.
  try {
    [WRITE_HELPERS, IDE_HELPERS] = await Promise.all([
      import("./write-helpers.mjs"),
      import("./ide-helpers.mjs"),
    ]);
  } catch (err) {
    setStatusBadge(`helpers import error: ${err.message}`, "error");
    return;
  }

  try {
    await initMsal();
  } catch (err) {
    setStatusBadge(`auth init error: ${err.message}`, "error");
    return;
  }

  try {
    await ensureSignedIn();
  } catch (err) {
    setStatusBadge(`sign-in error: ${err.message}`, "error");
    return;
  }

  setStatusBadge(`signed in: ${activeAccount.username} (read-write)`, "ok");
  if (connEl) connEl.textContent = "online";

  // Wire navigation + write-path buttons.
  const btnRefresh = document.getElementById("btn-refresh");
  if (btnRefresh) btnRefresh.addEventListener("click", () => renderList());
  const btnNewSession = document.getElementById("btn-new-session");
  if (btnNewSession) {
    btnNewSession.addEventListener("click", () => setView("new"));
  }
  const btnNewCancel = document.getElementById("btn-new-cancel");
  if (btnNewCancel) btnNewCancel.addEventListener("click", () => setView("list"));
  const form = document.getElementById("new-session-form");
  if (form) form.addEventListener("submit", handleNewSubmit);
  // Back buttons use their `data-target-view` attribute so the IDE-tab
  // detail returns to the IDE-tabs list (not to sessions).
  for (const back of document.querySelectorAll(".cockpit-back-btn")) {
    const target = back.dataset.targetView || "list";
    back.addEventListener("click", () => setView(target));
  }

  // Mode toggle (Sessions / IDE tabs) — read data-target-view so we don't
  // hard-code the mapping here.
  for (const btn of document.querySelectorAll(".cockpit-mode-btn")) {
    btn.addEventListener("click", () => {
      const target = btn.dataset.targetView;
      if (target) setView(target);
    });
  }

  // IDE-tabs refresh button (mirror of the sessions refresh button).
  const btnIdeRefresh = document.getElementById("btn-ide-refresh");
  if (btnIdeRefresh) {
    btnIdeRefresh.addEventListener("click", () => {
      renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
    });
  }

  setView("list");
  startAutoRefresh();
}

if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", () => {
    bootstrap().catch((err) => {
      setStatusBadge(`fatal: ${err.message}`, "error");
    });
  });
}

// =============================================================================
// 9. Test surface
// =============================================================================
//
// Pure helpers used by app.js live in two sibling ESM modules and ARE Node-
// importable:
//   - ./write-helpers.mjs    (Stage A2 write-path; unit-tested by
//                             tests/flows/mobile-cockpit/pwa-write-helpers-unit.sh)
//   - ./ide-helpers.mjs      (M2.1 read-only IDE-tabs view; unit-tested by
//                             tests/flows/mobile-cockpit/pwa-ide-helpers-unit.sh)
//
// The local pure helpers in this file (sortSessions, relativeTime,
// statusClass) are intentionally NOT module-exported here — they stay
// internal to the browser script. Promote them to write-helpers.mjs if
// you ever want to assert them from Node.
//
// The DOM-coupled write paths (createSession / approveSession /
// cancelSession + their button handlers + setView + renderList /
// renderDetail / renderNew + renderIdeTabsList / renderIdeTabDetail) are
// NOT unit-testable in isolation; they are covered by live-validation
// runs against the OneDrive state.json + ide-tabs.json (see
// START_HERE.md §8 once that section lands).
