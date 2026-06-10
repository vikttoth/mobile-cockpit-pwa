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
// Phase 2 (2026-06-10):
//   - Faster detail poll while status=running (streaming output from daemon)
//   - Follow-up / resume on done sessions (mergeQueueFollowUp + daemon --resume)
//   - Hash routing (#list / #new / #detail/<id>) for Teams deep links + bookmarks
//   - Teams push is daemon-side (MC_TEAMS_NOTIFY_WEBHOOK_URL); not in PWA yet
//   - Service worker / offline cache still deferred
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
// `2026-06-10 23:58 CEST f270f68`). Keep the string literal — index.html cache-busts on it.
const BUILD_STAMP = "2026-06-10 23:58 CEST f270f68";

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

/** Faster poll while viewing a running session detail (Phase 2 streaming). */
let runningDetailTimerId = null;

/** Session id currently open in detail view (for running poll). */
let activeDetailSessionId = null;

/**
 * Dynamically imported write-helpers module. Populated by bootstrap()
 * before any user-triggered write path can fire. We do this lazily so the
 * <script> tag for app.js can stay classic (UMD MSAL must load first); a
 * top-level static `import` would force ESM-module ordering and miss MSAL.
 */
let WRITE_HELPERS = null;

/** Dynamically imported pure helpers for the IDE-tabs view (M2.1). */
let IDE_HELPERS = null;

/** Compose draft preservation across ~20s IDE-tab auto-refresh. */
let COMPOSE_DRAFT_HELPERS = null;

/** AskQuestion mobile UI helpers (pure). */
let PENDING_QUESTION_HELPERS = null;

/** In-progress multi-select / bundle answers keyed by composerId. */
const idePendingAnswerDrafts = new Map();

/** Per-composerId in-progress message text (survives background re-render). */
const ideComposeDrafts = new Map();

/** Pure helpers for refresh-signals.json nudge + wait logic. */
let REFRESH_HELPERS = null;

/** True while a manual ↻ refresh is in flight (prevents double-tap). */
let refreshInFlight = false;

/** Cached last-known ide-tabs.json snapshot. Refreshed by loadIdeTabs(). */
let cachedIdeSnapshot = null;

/** IDE tabs sub-view within the read-only mirror: live open vs chat history. */
let ideListMode = "open";

/** Auto-refresh handle for the ide-tabs view (independent of sessions). */
let ideRefreshTimerId = null;

/**
 * Dynamically imported pure helpers for the M2.2.3 write-back actions
 * (`pwa/ide-actions-helpers.mjs`). Populated by bootstrap() before any
 * action button is wired.
 */
let IDE_ACTION_HELPERS = null;

/**
 * Composer ID currently shown in `view-ide-tab-detail`. Cached so the
 * "Close tab" confirmation modal and the "Send" button handler know
 * which tab to target without re-parsing the DOM. Cleared when leaving
 * the detail view.
 */
let activeIdeTabComposerId = null;

/**
 * Active workspace path that the extension is currently bound to (read
 * from `cachedIdeSnapshot.workspacePath`). Every M2.2.3 action carries
 * the workspace so the extension can refuse cross-workspace targeting.
 */
let activeIdeWorkspacePath = null;

/** When true, setView skips hash sync (hashchange handler is driving navigation). */
let suppressHashSync = false;

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
// 2b. Manual refresh (↻) — nudge desktop daemons + wait for fresh OneDrive data
// =============================================================================

function sleepMs(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Write a timestamp into cursor-cockpit/refresh-signals.json so
 * daemon/poll.mjs and ide-mirror/poll.mjs wake early.
 *
 * @param {'sessions' | 'ideTabs' | 'both'} scope
 */
async function writeRefreshNudge(scope) {
  const cfg = CONFIG && CONFIG.refreshSignals;
  if (!cfg || !cfg.endpoint) return;
  const endpoint = cfg.endpoint;
  let existing = REFRESH_HELPERS.emptyRefreshSignals();
  const metaRes = await graphFetch(endpoint);
  if (metaRes.ok) {
    const contentRes = await graphFetch(`${endpoint}:/content`);
    if (contentRes.ok) {
      try {
        existing = REFRESH_HELPERS.parseRefreshSignals(await contentRes.json());
      } catch {
        /* treat corrupt file as empty */
      }
    }
  } else if (metaRes.status !== 404) {
    throw new Error(`refresh-signals GET failed: ${metaRes.status}`);
  }
  const merged = REFRESH_HELPERS.applyNudge(
    existing,
    scope,
    new Date().toISOString(),
  );
  const putRes = await graphFetch(`${endpoint}:/content`, {
    method: "PUT",
    body: JSON.stringify(merged, null, 2) + "\n",
  });
  if (!putRes.ok) {
    throw new Error(`refresh-signals PUT failed: ${putRes.status}`);
  }
}

/**
 * Poll state.json until ETag changes or wait budget elapses.
 * @param {string|null} beforeEtag
 */
async function waitForFreshSessions(beforeEtag) {
  const cfg = CONFIG && CONFIG.refreshSignals;
  const maxMs = (cfg && cfg.waitMaxMs) || 15000;
  const pollMs = (cfg && cfg.waitPollMs) || 500;
  const minMs = (cfg && cfg.waitMinMs) || 2000;
  const start = Date.now();
  const deadline = start + maxMs;
  while (Date.now() < deadline) {
    const { etag } = await loadState();
    if (etag !== beforeEtag) return true;
    if (Date.now() - start >= minMs) return false;
    await sleepMs(pollMs);
  }
  return false;
}

/**
 * Poll ide-tabs.json until snapshotAt changes or wait budget elapses.
 * @param {string|null} beforeSnapshotAt
 */
async function waitForFreshIdeTabs(beforeSnapshotAt, beforeFingerprint) {
  const cfg = CONFIG && CONFIG.refreshSignals;
  const maxMs = (cfg && cfg.waitMaxMs) || 15000;
  const pollMs = (cfg && cfg.waitPollMs) || 500;
  const minMs = (cfg && cfg.waitMinMs) || 2000;
  const start = Date.now();
  const deadline = start + maxMs;
  while (Date.now() < deadline) {
    await loadIdeTabs();
    const snap = cachedIdeSnapshot;
    const afterAt = snap && snap.snapshotAt;
    const afterFp = snap && snap.contentFingerprint;
    if (afterAt && afterAt !== beforeSnapshotAt) return true;
    if (
      beforeFingerprint &&
      afterFp &&
      afterFp !== beforeFingerprint
    ) {
      return true;
    }
    if (Date.now() - start >= minMs) return false;
    await sleepMs(pollMs);
  }
  return false;
}

function setRefreshButtonsBusy(busy) {
  refreshInFlight = busy;
  for (const id of [
    "btn-refresh",
    "btn-detail-refresh",
    "btn-ide-refresh",
    "btn-ide-detail-refresh",
  ]) {
    const el = document.getElementById(id);
    if (!el) continue;
    el.disabled = busy;
    el.setAttribute("aria-busy", busy ? "true" : "false");
    el.classList.toggle("cockpit-btn-refreshing", busy);
  }
}

/**
 * User tapped ↻ on the current view — nudge the matching desktop daemon,
 * wait briefly for a OneDrive write, then re-render the active view.
 */
async function refreshCurrentView() {
  if (refreshInFlight) return;
  const view = document.body.dataset.view;
  setRefreshButtonsBusy(true);
  try {
    if (view === "list") {
      const beforeEtag = cachedStateEtag;
      await writeRefreshNudge("sessions");
      await waitForFreshSessions(beforeEtag);
      await renderList();
    } else if (view === "detail" && activeDetailSessionId) {
      const beforeEtag = cachedStateEtag;
      await writeRefreshNudge("sessions");
      await waitForFreshSessions(beforeEtag);
      await loadState();
      renderDetail(activeDetailSessionId);
    } else if (view === "ide-tabs") {
      const beforeAt = cachedIdeSnapshot && cachedIdeSnapshot.snapshotAt;
      const beforeFp = cachedIdeSnapshot && cachedIdeSnapshot.contentFingerprint;
      await writeRefreshNudge("ideTabs");
      await waitForFreshIdeTabs(beforeAt, beforeFp);
      await renderIdeTabsList();
    } else if (view === "ide-tab-detail") {
      const composerId =
        activeIdeTabComposerId ||
        document.getElementById("ide-detail-composer-id")?.textContent;
      const beforeAt = cachedIdeSnapshot && cachedIdeSnapshot.snapshotAt;
      const beforeFp = cachedIdeSnapshot && cachedIdeSnapshot.contentFingerprint;
      await writeRefreshNudge("ideTabs");
      await waitForFreshIdeTabs(beforeAt, beforeFp);
      await loadIdeTabs();
      if (composerId) renderIdeTabDetail(composerId, { preserveCompose: true });
    }
  } catch (err) {
    if (view === "detail") showDetailError(err.message);
    else if (view === "ide-tabs" || view === "ide-tab-detail") {
      showIdeTabsError(err.message);
    } else {
      showListError(err.message);
    }
  } finally {
    setRefreshButtonsBusy(false);
  }
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

/**
 * Queue a follow-up on a finished session (resume same cursor-agent chat).
 * Returns { sessionId, status } on success.
 */
async function queueFollowUp(sessionId, prompt) {
  if (!WRITE_HELPERS) throw new Error("write-helpers module not loaded yet (bootstrap order bug)");
  const trimmed = typeof prompt === "string" ? prompt.trim() : "";
  if (!trimmed) throw new Error("Follow-up prompt is required");

  const autoApprove = !!(CONFIG.session && CONFIG.session.autoApprove);
  setStatusBadge("saving…", "saving");

  for (let attempt = 0; attempt < 2; attempt++) {
    const { state, etag } = await loadState();
    const next = WRITE_HELPERS.mergeQueueFollowUp(state, sessionId, {
      prompt: trimmed,
      now: new Date(),
      autoApprove,
    });
    try {
      await putState(next, etag);
      flashSavedBadge();
      const row = next.sessions.find((x) => x.sessionId === sessionId);
      return { sessionId, status: row ? row.status : "pending" };
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) {
        setStatusBadge(`save failed: ${err.message}`, "error");
        throw err;
      }
    }
  }
  setStatusBadge("save failed: retries exhausted", "error");
  throw new Error("queueFollowUp: retries exhausted");
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

function syncHashForView(viewId, payload) {
  if (typeof window === "undefined" || suppressHashSync || !WRITE_HELPERS) return;
  const target = WRITE_HELPERS.formatViewHash(viewId, payload);
  if (target == null) return;
  const want = `#${target}`;
  if (window.location.hash !== want) {
    suppressHashSync = true;
    try {
      window.location.hash = target;
    } finally {
      suppressHashSync = false;
    }
  }
}

function applyHashRoute() {
  if (!WRITE_HELPERS || typeof window === "undefined") return;
  const route = WRITE_HELPERS.parseLocationHash(window.location.hash);
  if (!route) return;
  suppressHashSync = true;
  try {
    if (route.view === "list") setView("list");
    else if (route.view === "new") setView("new");
    else if (route.view === "detail" && route.sessionId) {
      setView("detail", { sessionId: route.sessionId });
    }
  } finally {
    suppressHashSync = false;
  }
}

function setView(viewId, payload) {
  const prevView = document.body.dataset.view;
  if (
    prevView === "ide-tab-detail" &&
    viewId !== "ide-tab-detail" &&
    viewId !== "close-confirm"
  ) {
    const taLeave = document.getElementById("ide-detail-compose-text");
    if (activeIdeTabComposerId && taLeave) {
      saveComposeDraft(activeIdeTabComposerId, taLeave.value);
    }
  }
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
    activeDetailSessionId = payload.sessionId;
    renderDetail(payload.sessionId);
    syncRunningDetailPoll();
  } else if (viewId === "new") {
    renderNew();
  } else if (viewId === "ide-tabs") {
    syncIdeListModeToggle();
    renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
  } else if (viewId === "ide-tab-detail" && payload && payload.composerId) {
    renderIdeTabDetail(payload.composerId);
  } else if (viewId === "new-agent") {
    renderNewAgentModal();
  } else if (viewId === "close-confirm" && payload && payload.composerId) {
    renderCloseConfirmModal(payload.composerId, payload.title || "");
  }
  // Drop the cached composer ID when navigating away from the IDE detail
  // view so a stale value can't accidentally target the wrong tab on a
  // later send/close click. The close-confirm + new-agent modals overlay
  // the detail view and intentionally KEEP the cached ID.
  if (viewId !== "ide-tab-detail" && viewId !== "close-confirm") {
    activeIdeTabComposerId = null;
  }
  if (viewId !== "detail") {
    activeDetailSessionId = null;
    stopRunningDetailPoll();
  }
  if (viewId === "list" || viewId === "new" || (viewId === "detail" && payload && payload.sessionId)) {
    syncHashForView(viewId, payload);
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
  set("detail-agent-id", s.cursorAgentId || s.agentId);
  set("detail-model", s.model);
  set("detail-created", s.createdAt || s.created);
  set("detail-updated", s.lastUpdated);
  set("detail-prompt-text", s.prompt);
  set("detail-output-text", s.output || (s.status === "running" ? "(streaming…)" : ""));

  const outputHeading = document.getElementById("detail-output-heading");
  if (outputHeading) {
    outputHeading.textContent = s.status === "running" ? "Output (streaming…)" : "Output";
  }

  const followUpPanel = document.getElementById("detail-follow-up-panel");
  if (followUpPanel) followUpPanel.hidden = true;

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
    const agentId = s.cursorAgentId || s.agentId;
    const showFollowUp = s.status === "done" && !!agentId;
    btnFollowUp.hidden = !showFollowUp;
    btnFollowUp.disabled = !showFollowUp;
    btnFollowUp.onclick = showFollowUp
      ? () => showFollowUpPanel(s.sessionId)
      : null;
  }

  syncRunningDetailPoll();
}

function showFollowUpPanel(sessionId) {
  const panel = document.getElementById("detail-follow-up-panel");
  const ta = document.getElementById("follow-up-prompt");
  if (panel) {
    panel.hidden = false;
    panel.dataset.sessionId = sessionId;
  }
  if (ta) {
    ta.value = "";
    ta.focus();
  }
}

function hideFollowUpPanel() {
  const panel = document.getElementById("detail-follow-up-panel");
  if (panel) panel.hidden = true;
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

function syncIdeListModeToggle() {
  for (const btn of document.querySelectorAll(".cockpit-ide-list-btn")) {
    btn.setAttribute(
      "aria-current",
      btn.dataset.ideListMode === ideListMode ? "true" : "false",
    );
  }
}

function setIdeListMode(mode) {
  if (mode !== "open" && mode !== "history") return;
  ideListMode = mode;
  syncIdeListModeToggle();
  renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
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

  const snapshot = cachedIdeSnapshot || { openTabs: [], historyTabs: [] };
  const tabs = IDE_HELPERS.pickIdeTabList(snapshot, ideListMode);
  const cap = (CONFIG.pwa && CONFIG.pwa.recentSessionsCount) || 50;
  const sorted = IDE_HELPERS.orderIdeTabsForDisplay(tabs, cap, {
    openTabsSource: ideListMode === "open" ? snapshot.openTabsSource : null,
  });

  const sourceWarn = document.getElementById("ide-tabs-source-warning");
  if (sourceWarn) {
    const hint =
      ideListMode === "open"
        ? IDE_HELPERS.openTabsSourceHint(snapshot.openTabsSource)
        : null;
    if (hint) {
      sourceWarn.textContent = hint;
      sourceWarn.hidden = false;
    } else {
      sourceWarn.textContent = "";
      sourceWarn.hidden = true;
    }
  }

  const bucket = IDE_HELPERS.summarizeWaitingOn(sorted);
  const snapshotAtRel = IDE_HELPERS.relativeIdeTime(snapshot.snapshotAt, Date.now());
  summary.innerHTML = "";
  const total = document.createElement("span");
  total.className = "cockpit-ide-summary-bucket";
  const listLabel = ideListMode === "history" ? "in history" : "open";
  total.innerHTML =
    `<strong>${bucket.total}</strong> ${listLabel}` +
    (bucket.total === 1 ? " tab" : " tabs");
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
    if (ideListMode === "history") {
      empty.textContent =
        "No recent chat history in the mirror yet. Older conversations appear here after the next poll.";
    } else {
      empty.textContent =
        "No open IDE tabs in the mirror. Open chats in Cursor, ensure the mobile-cockpit extension is active, then refresh.";
    }
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

function renderIdeTabDetail(composerId, options = {}) {
  if (!IDE_HELPERS) {
    showIdeDetailError("ide-helpers module not loaded yet (bootstrap order bug)");
    return;
  }
  clearIdeDetailError();
  if (!cachedIdeSnapshot) {
    showIdeTabsError("No cached IDE snapshot — refresh the IDE-tabs list first.");
    setView("ide-tabs");
    return;
  }
  const tab = IDE_HELPERS.findIdeTab(cachedIdeSnapshot, composerId);
  if (!tab) {
    showIdeTabsError(`IDE tab not in current snapshot: ${composerId}`);
    setView("ide-tabs");
    return;
  }

  const taBefore = document.getElementById("ide-detail-compose-text");
  const prevComposer = activeIdeTabComposerId;
  if (prevComposer && prevComposer !== tab.composerId && taBefore) {
    saveComposeDraft(prevComposer, taBefore.value);
  }

  // Cache for the send / close handlers; cleared by setView when navigating away.
  activeIdeTabComposerId = tab.composerId;
  activeIdeWorkspacePath = cachedIdeSnapshot.workspacePath || null;

  const draftForTab = ideComposeDrafts.get(tab.composerId) || "";
  const preserveCompose = COMPOSE_DRAFT_HELPERS
    ? COMPOSE_DRAFT_HELPERS.shouldPreserveComposeOnRefresh({
        preserveCompose: options.preserveCompose === true,
        textareaFocused: taBefore === document.activeElement,
        textareaValue: taBefore ? taBefore.value : "",
        draftText: draftForTab,
      })
    : options.preserveCompose === true;

  if (!preserveCompose) {
    resetComposeUi({ clearDraft: true, composerId: tab.composerId });
  } else {
    const err = document.getElementById("ide-detail-compose-error");
    if (err) err.hidden = true;
  }

  // Wire the Close-tab button to the confirm modal. Idempotent on every
  // re-render — we re-attach because the modal payload (composer ID +
  // title) depends on which tab is open.
  const btnClose = document.getElementById("btn-ide-close-tab");
  if (btnClose) {
    btnClose.onclick = () => setView("close-confirm", {
      composerId: tab.composerId,
      title: IDE_HELPERS.formatTabTitle(tab.title, 60),
    });
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

  renderIdeTabPendingQuestion(tab);

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

  if (preserveCompose) {
    const restored =
      (taBefore && taBefore.value.length > 0 ? taBefore.value : null) ||
      draftForTab ||
      "";
    if (restored) {
      applyComposeUiFromText(restored);
      saveComposeDraft(tab.composerId, restored);
    }
  }
}

/**
 * Send a user message into the active IDE tab (shared by compose + AskQuestion UI).
 * @param {string} text
 * @returns {Promise<void>}
 */
async function sendIdeTabMessage(text) {
  const err = document.getElementById("ide-detail-compose-error");
  if (!activeIdeTabComposerId) {
    if (err) {
      err.textContent = "No active IDE tab — re-open the tab from the list.";
      err.hidden = false;
    }
    throw new Error("no active IDE tab");
  }
  if (!IDE_ACTION_HELPERS) {
    if (err) {
      err.textContent = "Action helpers not loaded.";
      err.hidden = false;
    }
    throw new Error("action helpers not loaded");
  }
  const trimmed = String(text || "").trim();
  if (!trimmed) {
    throw new Error("empty message");
  }
  const action = IDE_ACTION_HELPERS.buildSendMessageAction({
    tabId: activeIdeTabComposerId,
    text: trimmed,
    now: new Date(),
    rngFn: WRITE_HELPERS && WRITE_HELPERS.cryptoRandomBytes,
  });
  const allowedCwds = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  const validation = IDE_ACTION_HELPERS.validateActionInputs(action, allowedCwds);
  if (!validation.valid) {
    const msg = validation.errors.join("; ");
    if (err) {
      err.textContent = msg;
      err.hidden = false;
    }
    throw new Error(msg);
  }
  if (err) err.hidden = true;
  await submitIdeAction(action);
  if (activeIdeTabComposerId) ideComposeDrafts.delete(activeIdeTabComposerId);
  resetComposeUi();
}

function syncPendingSubmitButtons(body, questions, answers) {
  if (!PENDING_QUESTION_HELPERS || !body) return;
  const submitAll = body.querySelector(".cockpit-ide-pending-submit-all");
  if (submitAll) {
    submitAll.disabled = !PENDING_QUESTION_HELPERS.allQuestionsAnswered(questions, answers);
  }
  for (const block of body.querySelectorAll(".cockpit-ide-pending-q")) {
    const submitOne = block.querySelector(".cockpit-ide-pending-submit-one");
    if (!submitOne) continue;
    const key = submitOne.dataset.questionKey || "";
    const q = questions.find((item) => PENDING_QUESTION_HELPERS.questionKey(item) === key);
    if (!q) continue;
    const v = answers[key];
    const ok = q.allowMultiple
      ? Array.isArray(v) && v.length > 0
      : v != null && String(v).trim() !== "";
    submitOne.disabled = !ok;
  }
}

function appendPendingFreeTextRow(block, q, questions, answers, onAnswerChange) {
  const H = PENDING_QUESTION_HELPERS;
  const key = H.questionKey(q);
  const row = document.createElement("div");
  row.className = "cockpit-ide-pending-freetext";
  const input = document.createElement("textarea");
  input.className = "cockpit-compose-textarea cockpit-ide-pending-freetext-input";
  input.rows = 2;
  input.placeholder = "Type your answer…";
  input.value = typeof answers[key] === "string" ? answers[key] : "";
  input.addEventListener("input", () => {
    answers[key] = input.value;
    onAnswerChange();
  });
  row.appendChild(input);
  const sendBtn = document.createElement("button");
  sendBtn.type = "button";
  sendBtn.className = "cockpit-btn cockpit-btn-primary cockpit-ide-pending-submit-one";
  sendBtn.dataset.questionKey = key;
  sendBtn.textContent = questions.length > 1 ? "Save answer" : "Send answer";
  sendBtn.addEventListener("click", async () => {
    if (questions.length > 1) {
      onAnswerChange();
      input.focus();
      return;
    }
    const text = H.formatAnswersForSend([q], answers);
    if (!text) return;
    sendBtn.disabled = true;
    try {
      await sendIdeTabMessage(text);
      if (activeIdeTabComposerId) idePendingAnswerDrafts.delete(activeIdeTabComposerId);
    } catch {
      sendBtn.disabled = false;
      onAnswerChange();
    }
  });
  row.appendChild(sendBtn);
  block.appendChild(row);
}

/**
 * Show AskQuestion at the top: tap options, multi-select, or free-text + send.
 */
function renderIdeTabPendingQuestion(tab) {
  const section = document.getElementById("ide-detail-pending");
  const body = document.getElementById("ide-detail-pending-body");
  if (!section || !body || !PENDING_QUESTION_HELPERS) return;
  body.innerHTML = "";
  const composerId = tab && tab.composerId;
  const pq = tab && tab.pendingQuestion;
  const questions = pq && Array.isArray(pq.questions) ? pq.questions : [];
  if (questions.length === 0) {
    section.hidden = true;
    if (composerId) idePendingAnswerDrafts.delete(composerId);
    return;
  }
  section.hidden = false;
  const H = PENDING_QUESTION_HELPERS;
  const fingerprint = H.pendingQuestionsFingerprint(questions);
  let draft = composerId ? idePendingAnswerDrafts.get(composerId) : null;
  if (!draft || draft.fingerprint !== fingerprint) {
    draft = { fingerprint, answers: {} };
    if (composerId) idePendingAnswerDrafts.set(composerId, draft);
  }
  const answers = draft.answers;
  const multiBundle =
    questions.length > 1 || questions.some((q) => q.allowMultiple);

  const onAnswerChange = () => syncPendingSubmitButtons(body, questions, answers);

  for (const q of questions) {
    const block = document.createElement("div");
    block.className = "cockpit-ide-pending-q";
    const prompt = document.createElement("p");
    prompt.className = "cockpit-ide-pending-prompt";
    prompt.textContent = q.prompt || "Choose an option:";
    block.appendChild(prompt);

    if (q.allowMultiple) {
      const hint = document.createElement("p");
      hint.className = "cockpit-ide-pending-hint";
      hint.textContent = "Select one or more, then submit.";
      block.appendChild(hint);
    }

    if (H.showInlineFreeTextInput(q)) {
      appendPendingFreeTextRow(block, q, questions, answers, onAnswerChange);
      body.appendChild(block);
      continue;
    }

    const key = H.questionKey(q);
    const opts = document.createElement("div");
    opts.className = "cockpit-ide-pending-options";

    for (const o of q.options || []) {
      if (H.isFreeTextEscapeOption(o)) {
        const escapeBtn = document.createElement("button");
        escapeBtn.type = "button";
        escapeBtn.className = "cockpit-btn cockpit-ide-pending-opt cockpit-ide-pending-opt-escape";
        escapeBtn.textContent = o.label || o.id || "Other…";
        escapeBtn.addEventListener("click", () => {
          opts.hidden = true;
          const existing = block.querySelector(".cockpit-ide-pending-freetext");
          if (existing) existing.remove();
          appendPendingFreeTextRow(block, q, questions, answers, onAnswerChange);
          const input = block.querySelector(".cockpit-ide-pending-freetext-input");
          if (input) input.focus();
        });
        opts.appendChild(escapeBtn);
        continue;
      }

      const optBtn = document.createElement("button");
      optBtn.type = "button";
      optBtn.className = "cockpit-btn cockpit-ide-pending-opt";
      const label = o.label || o.id || "?";
      optBtn.textContent = label;

      if (q.allowMultiple) {
        optBtn.addEventListener("click", () => {
          const cur = Array.isArray(answers[key]) ? answers[key] : [];
          const idx = cur.indexOf(label);
          if (idx >= 0) {
            cur.splice(idx, 1);
            optBtn.classList.remove("cockpit-ide-pending-opt-selected");
          } else {
            cur.push(label);
            optBtn.classList.add("cockpit-ide-pending-opt-selected");
          }
          answers[key] = cur;
          onAnswerChange();
        });
      } else if (multiBundle) {
        optBtn.addEventListener("click", () => {
          answers[key] = label;
          for (const btn of opts.querySelectorAll(".cockpit-ide-pending-opt")) {
            btn.classList.toggle("cockpit-ide-pending-opt-selected", btn === optBtn);
          }
          onAnswerChange();
        });
      } else {
        optBtn.addEventListener("click", async () => {
          optBtn.disabled = true;
          try {
            await sendIdeTabMessage(label);
            if (composerId) idePendingAnswerDrafts.delete(composerId);
          } catch {
            optBtn.disabled = false;
          }
        });
      }
      opts.appendChild(optBtn);
    }

    if ((q.options || []).length > 0) block.appendChild(opts);

    if (questions.length === 1 && q.allowMultiple) {
      const submitOne = document.createElement("button");
      submitOne.type = "button";
      submitOne.className = "cockpit-btn cockpit-btn-primary cockpit-ide-pending-submit-one";
      submitOne.dataset.questionKey = key;
      submitOne.textContent = "Submit answer";
      submitOne.addEventListener("click", async () => {
        const text = H.formatAnswersForSend(questions, answers);
        if (!text) return;
        submitOne.disabled = true;
        try {
          await sendIdeTabMessage(text);
          if (composerId) idePendingAnswerDrafts.delete(composerId);
        } catch {
          submitOne.disabled = false;
          onAnswerChange();
        }
      });
      block.appendChild(submitOne);
    }

    body.appendChild(block);
  }

  if (questions.length > 1) {
    const submitAll = document.createElement("button");
    submitAll.type = "button";
    submitAll.className = "cockpit-btn cockpit-btn-primary cockpit-ide-pending-submit-all";
    submitAll.textContent = "Submit all answers";
    submitAll.addEventListener("click", async () => {
      const text = H.formatAnswersForSend(questions, answers);
      if (!text) return;
      submitAll.disabled = true;
      try {
        await sendIdeTabMessage(text);
        if (composerId) idePendingAnswerDrafts.delete(composerId);
      } catch {
        submitAll.disabled = false;
        onAnswerChange();
      }
    });
    body.appendChild(submitAll);
  }

  onAnswerChange();
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

function stopRunningDetailPoll() {
  if (runningDetailTimerId !== null) {
    clearInterval(runningDetailTimerId);
    runningDetailTimerId = null;
  }
}

function syncRunningDetailPoll() {
  stopRunningDetailPoll();
  if (document.body.dataset.view !== "detail" || !activeDetailSessionId || !cachedState) return;
  const s = (cachedState.sessions || []).find((x) => x.sessionId === activeDetailSessionId);
  if (!s || s.status !== "running") return;
  const sec = (CONFIG.pwa && CONFIG.pwa.runningPollIntervalSeconds) || 5;
  const intervalMs = Math.max(3, sec | 0) * 1000;
  runningDetailTimerId = setInterval(() => {
    if (document.body.dataset.view !== "detail" || !activeDetailSessionId) return;
    loadState()
      .then(() => {
        renderDetail(activeDetailSessionId);
      })
      .catch((err) => showDetailError(err.message));
  }, intervalMs);
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
              renderIdeTabDetail(openId.textContent, { preserveCompose: true });
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

async function handleFollowUpSubmit(ev) {
  ev.preventDefault();
  clearDetailError();
  const panel = document.getElementById("detail-follow-up-panel");
  const sessionId = panel && panel.dataset.sessionId;
  const ta = document.getElementById("follow-up-prompt");
  const submitBtn = document.getElementById("btn-follow-up-submit");
  const prompt = ta ? ta.value : "";
  if (!sessionId) {
    showDetailError("No session selected for follow-up");
    return;
  }
  if (submitBtn) submitBtn.disabled = true;
  try {
    await queueFollowUp(sessionId, prompt);
    hideFollowUpPanel();
    await renderList();
    renderDetail(sessionId);
  } catch (err) {
    showDetailError(err.message);
  } finally {
    if (submitBtn) submitBtn.disabled = false;
  }
}

// =============================================================================
// 7b. M2.2.3 write-back: actions file I/O + toast + modal renderers
// =============================================================================
//
// Shape (mirrors lib/config.mjs + extension/src/lib/types.ts):
//   - GET  cursor-cockpit/ide-actions.json         (action queue)
//   - PUT  cursor-cockpit/ide-actions.json         (append a new action)
//   - GET  cursor-cockpit/ide-actions-results.json (poll for the result)
//
// The PWA is the only writer for ide-actions.json (single-writer); the
// extension is the only writer for ide-actions-results.json. Therefore
// we DO NOT need optimistic-concurrency ETag handshakes here — the
// flows are append-only on a single writer per file. Compare with
// state.json (multi-writer between PWA + daemon) which keeps the
// If-Match dance.
//
// Result polling: after a successful PUT, we kick off a short polling
// loop that GET-fetches ide-actions-results.json every
// CONFIG.ideActions.resultPollIntervalSeconds until the matching
// actionId appears or CONFIG.ideActions.resultTimeoutSeconds elapses.
// The current toast reflects the outcome (in_progress → done → success
// toast; error → error toast).

/**
 * Load ide-actions.json from OneDrive. Returns a normalized object
 * (`{ schemaVersion, snapshotAt, actions }`) — on 404 we synthesize an
 * empty stub so the very first PUT can seed the file.
 */
async function loadIdeActionsFile() {
  const endpoint = CONFIG && CONFIG.ideActions && CONFIG.ideActions.endpoint;
  if (!endpoint) {
    throw new Error("config.ideActions.endpoint missing -- update pwa/config.json");
  }
  const res = await graphFetch(`${endpoint}:/content`);
  if (res.status === 404) {
    return { schemaVersion: 1, snapshotAt: null, actions: [] };
  }
  if (!res.ok) {
    throw new Error(`ide-actions.json GET failed: ${res.status} ${res.statusText}`);
  }
  return await res.json();
}

/**
 * PUT a fresh ide-actions.json. The action file is single-writer (PWA
 * only), so no If-Match handshake. On any non-2xx response the caller
 * surfaces an error toast.
 */
async function putIdeActionsFile(nextFile) {
  const endpoint = CONFIG && CONFIG.ideActions && CONFIG.ideActions.endpoint;
  if (!endpoint) {
    throw new Error("config.ideActions.endpoint missing -- update pwa/config.json");
  }
  const res = await graphFetch(`${endpoint}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(nextFile),
  });
  if (!res.ok) {
    throw new Error(`ide-actions.json PUT failed: ${res.status} ${res.statusText}`);
  }
}

/**
 * GET ide-actions-results.json. Returns null on 404 (extension never
 * wrote a result yet); otherwise returns the parsed file.
 */
async function loadIdeActionsResultsFile() {
  const endpoint = CONFIG && CONFIG.ideActions && CONFIG.ideActions.resultsEndpoint;
  if (!endpoint) return null;
  const res = await graphFetch(`${endpoint}:/content`);
  if (res.status === 404) return null;
  if (!res.ok) {
    throw new Error(`ide-actions-results.json GET failed: ${res.status} ${res.statusText}`);
  }
  return await res.json();
}

/**
 * Submit an action: append + PUT. On success, kicks off the result
 * polling loop in the background and returns immediately so the caller
 * can update the UI optimistically.
 */
async function submitIdeAction(action) {
  if (!IDE_ACTION_HELPERS) {
    throw new Error("ide-actions-helpers not loaded yet (bootstrap order bug)");
  }
  const prev = await loadIdeActionsFile();
  const next = IDE_ACTION_HELPERS.mergeAppendAction(prev, action, new Date());
  await putIdeActionsFile(next);
  // Fire-and-forget the result poll; UI toast is updated inside.
  pollIdeActionResult(action).catch((err) => {
    showToast(`Action ${action.kind} (${action.actionId.slice(0, 8)}) result poll failed: ${err.message}`, "error");
  });
}

/**
 * Poll ide-actions-results.json until a result for `action.actionId`
 * shows up OR the configured timeout elapses. Surfaces a toast on
 * terminal states (`done`, `error`, `skipped`). Pure infinite loops are
 * avoided via the timeout guard.
 */
async function pollIdeActionResult(action) {
  const intervalMs = Math.max(1, (CONFIG.ideActions.resultPollIntervalSeconds | 0)) * 1000;
  const timeoutMs = Math.max(5, (CONFIG.ideActions.resultTimeoutSeconds | 0)) * 1000;
  const startedAt = Date.now();
  showToast(`${humanizeKind(action.kind)} queued (${action.actionId.slice(0, 8)})…`, "info");
  while (Date.now() - startedAt < timeoutMs) {
    await new Promise((r) => setTimeout(r, intervalMs));
    let file;
    try {
      file = await loadIdeActionsResultsFile();
    } catch (err) {
      // Network blip — keep trying; final timeout still applies.
      continue;
    }
    const entry = IDE_ACTION_HELPERS.findResultForAction(file, action.actionId);
    if (!entry) continue;
    if (entry.status === "in_progress") {
      // Extension picked it up but hasn't finished — keep polling.
      continue;
    }
    if (entry.status === "done") {
      showToast(`${humanizeKind(action.kind)} done.`, "success");
      // Trigger a fresh IDE-tabs render so the user sees the new tab /
      // closed tab. The detail view also auto-refreshes via the IDE
      // refresh timer; we just give it a head start.
      // Best-effort head-start refresh. The auto-refresh timer will
      // retry on its next tick if either path fails; we log to console
      // instead of toasting so a transient blip doesn't compete with
      // the just-shown success banner.
      loadIdeTabs().then(() => {
        if (document.body.dataset.view === "ide-tabs") {
          renderIdeTabsList().catch((e) => console.warn("post-action ide-tabs render failed:", e));
        }
      }).catch((e) => console.warn("post-action ide-tabs reload failed:", e));
      return;
    }
    if (entry.status === "error") {
      showToast(`${humanizeKind(action.kind)} failed: ${entry.error || "unknown"}`, "error");
      return;
    }
    if (entry.status === "skipped") {
      showToast(`${humanizeKind(action.kind)} skipped: ${entry.error || "no reason"}`, "info");
      return;
    }
  }
  showToast(`${humanizeKind(action.kind)} timed out waiting for extension result.`, "error");
}

function humanizeKind(kind) {
  switch (kind) {
    case "send_message": return "Send message";
    case "new_agent":    return "New agent";
    case "close_tab":    return "Close tab";
    default:             return kind;
  }
}

/**
 * Append a toast banner to the bottom-center stack; auto-dismiss after
 * ~5 s. `kind` ∈ {"info", "success", "error"} — styled via CSS.
 */
function showToast(message, kind) {
  const container = document.getElementById("cockpit-toast-container");
  if (!container) return;
  const el = document.createElement("div");
  el.className = "cockpit-toast";
  el.dataset.kind = kind || "info";
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => {
    el.style.transition = "opacity 200ms ease-out";
    el.style.opacity = "0";
    setTimeout(() => el.remove(), 250);
  }, 5000);
}

// -----------------------------------------------------------------------------
// Compose UI (per-tab Send) — wired by renderIdeTabDetail()
// -----------------------------------------------------------------------------

function saveComposeDraft(composerId, text) {
  if (!COMPOSE_DRAFT_HELPERS || !COMPOSE_DRAFT_HELPERS.isValidComposeDraftKey(composerId)) {
    return;
  }
  const t = typeof text === "string" ? text : "";
  if (t.length > 0) ideComposeDrafts.set(composerId, t);
  else ideComposeDrafts.delete(composerId);
}

function applyComposeUiFromText(text) {
  const ta = document.getElementById("ide-detail-compose-text");
  const btn = document.getElementById("btn-ide-send-message");
  const counter = document.getElementById("ide-detail-compose-counter");
  const max = IDE_ACTION_HELPERS ? IDE_ACTION_HELPERS.MAX_TEXT_LEN : 20000;
  const safe = typeof text === "string" ? text : "";
  const state = COMPOSE_DRAFT_HELPERS
    ? COMPOSE_DRAFT_HELPERS.composeUiStateFromText(safe, max)
    : {
        counterText: `${safe.length} / ${max}`,
        over: safe.length > max,
        sendDisabled: safe.length === 0 || safe.length > max,
      };
  if (ta) ta.value = safe;
  if (counter) {
    counter.textContent = state.counterText;
    counter.dataset.over = state.over ? "true" : "false";
  }
  if (btn) btn.disabled = state.sendDisabled;
}

function resetComposeUi({ clearDraft = false, composerId = null } = {}) {
  if (
    clearDraft &&
    composerId &&
    COMPOSE_DRAFT_HELPERS &&
    COMPOSE_DRAFT_HELPERS.isValidComposeDraftKey(composerId)
  ) {
    ideComposeDrafts.delete(composerId);
  }
  applyComposeUiFromText("");
  const err = document.getElementById("ide-detail-compose-error");
  if (err) err.hidden = true;
}

function wireComposeUi() {
  const ta = document.getElementById("ide-detail-compose-text");
  const btn = document.getElementById("btn-ide-send-message");
  const counter = document.getElementById("ide-detail-compose-counter");
  if (!ta || !btn || !counter) return;
  const max = IDE_ACTION_HELPERS ? IDE_ACTION_HELPERS.MAX_TEXT_LEN : 20000;
  ta.addEventListener("input", () => {
    const len = ta.value.length;
    counter.textContent = `${len} / ${max}`;
    counter.dataset.over = len > max ? "true" : "false";
    btn.disabled = len === 0 || len > max;
    if (activeIdeTabComposerId) saveComposeDraft(activeIdeTabComposerId, ta.value);
  });
  btn.addEventListener("click", () => handleSendMessageClick());
}

async function handleSendMessageClick() {
  const ta = document.getElementById("ide-detail-compose-text");
  const btn = document.getElementById("btn-ide-send-message");
  if (!ta || !btn) return;
  btn.disabled = true;
  try {
    await sendIdeTabMessage(ta.value);
  } catch (e) {
    btn.disabled = ta.value.length > 0;
  }
}

// -----------------------------------------------------------------------------
// New-agent modal — toolbar button → modal → action
// -----------------------------------------------------------------------------

function renderNewAgentModal() {
  // Populate the workspace dropdown from CONFIG.session.allowedCwds.
  const select = document.getElementById("new-agent-workspace");
  const ta = document.getElementById("new-agent-text");
  const err = document.getElementById("new-agent-error");
  if (!select) return;
  const allowed = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  select.innerHTML = "";
  // Pre-select the current IDE workspace if it's in the allow-list — that
  // matches the user's mental model ("open another agent in this project").
  const current = (cachedIdeSnapshot && cachedIdeSnapshot.workspacePath) || null;
  for (const cwd of allowed) {
    const opt = document.createElement("option");
    opt.value = cwd;
    opt.textContent = cwd;
    if (cwd === current) opt.selected = true;
    select.appendChild(opt);
  }
  if (ta) ta.value = "";
  if (err) err.hidden = true;
}

async function handleNewAgentSubmit(ev) {
  ev.preventDefault();
  const select = document.getElementById("new-agent-workspace");
  const ta = document.getElementById("new-agent-text");
  const err = document.getElementById("new-agent-error");
  const btn = document.getElementById("btn-new-agent-create");
  if (!select || !ta) return;
  if (!IDE_ACTION_HELPERS) {
    if (err) { err.textContent = "Action helpers not loaded."; err.hidden = false; }
    return;
  }
  const workspace = select.value;
  const text = (ta.value || "").trim();
  const action = IDE_ACTION_HELPERS.buildNewAgentAction({
    workspace,
    text: text || null,
    now: new Date(),
    rngFn: WRITE_HELPERS && WRITE_HELPERS.cryptoRandomBytes,
  });
  const allowedCwds = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  const validation = IDE_ACTION_HELPERS.validateActionInputs(action, allowedCwds);
  if (!validation.valid) {
    if (err) { err.textContent = validation.errors.join("; "); err.hidden = false; }
    return;
  }
  if (err) err.hidden = true;
  if (btn) btn.disabled = true;
  try {
    await submitIdeAction(action);
    setView("ide-tabs");
  } catch (e) {
    if (err) { err.textContent = e.message; err.hidden = false; }
  } finally {
    if (btn) btn.disabled = false;
  }
}

// -----------------------------------------------------------------------------
// Close-tab confirmation modal — wired on render so the payload is fresh
// -----------------------------------------------------------------------------

function renderCloseConfirmModal(composerId, title) {
  const titleEl = document.getElementById("close-confirm-title-name");
  const idEl = document.getElementById("close-confirm-tab-id");
  const btnConfirm = document.getElementById("btn-close-confirm");
  if (titleEl) titleEl.textContent = title || "(untitled)";
  if (idEl) idEl.textContent = composerId || "—";
  if (btnConfirm) {
    btnConfirm.onclick = () => handleCloseTabConfirm(composerId);
  }
}

async function handleCloseTabConfirm(composerId) {
  if (!IDE_ACTION_HELPERS) {
    showToast("Action helpers not loaded.", "error");
    return;
  }
  const action = IDE_ACTION_HELPERS.buildCloseTabAction({
    tabId: composerId,
    now: new Date(),
    rngFn: WRITE_HELPERS && WRITE_HELPERS.cryptoRandomBytes,
  });
  const allowedCwds = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  const validation = IDE_ACTION_HELPERS.validateActionInputs(action, allowedCwds);
  if (!validation.valid) {
    showToast(`Close-tab validation failed: ${validation.errors.join("; ")}`, "error");
    return;
  }
  try {
    await submitIdeAction(action);
    // Pop the user back to the IDE-tabs list — the detail view will be
    // dead once the extension confirms the close anyway, and showing
    // an empty thread is worse than dropping to the list.
    setView("ide-tabs");
  } catch (e) {
    showToast(`Close-tab failed: ${e.message}`, "error");
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
    [
      WRITE_HELPERS,
      IDE_HELPERS,
      IDE_ACTION_HELPERS,
      REFRESH_HELPERS,
      COMPOSE_DRAFT_HELPERS,
      PENDING_QUESTION_HELPERS,
    ] = await Promise.all([
      import("./write-helpers.mjs?v=f270f68"),
      import("./ide-helpers.mjs?v=f270f68"),
      import("./ide-actions-helpers.mjs?v=f270f68"),
      import("./refresh-helpers.mjs"),
      import("./compose-draft.mjs"),
      import("./pending-question.mjs"),
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
  if (btnRefresh) {
    btnRefresh.addEventListener("click", () => {
      refreshCurrentView().catch((err) => showListError(err.message));
    });
  }
  const btnDetailRefresh = document.getElementById("btn-detail-refresh");
  if (btnDetailRefresh) {
    btnDetailRefresh.addEventListener("click", () => {
      refreshCurrentView().catch((err) => showDetailError(err.message));
    });
  }
  const btnNewSession = document.getElementById("btn-new-session");
  if (btnNewSession) {
    btnNewSession.addEventListener("click", () => setView("new"));
  }
  const btnNewCancel = document.getElementById("btn-new-cancel");
  if (btnNewCancel) btnNewCancel.addEventListener("click", () => setView("list"));
  const form = document.getElementById("new-session-form");
  if (form) form.addEventListener("submit", handleNewSubmit);
  const followUpForm = document.getElementById("follow-up-form");
  if (followUpForm) followUpForm.addEventListener("submit", handleFollowUpSubmit);
  const btnFollowUpCancel = document.getElementById("btn-follow-up-cancel");
  if (btnFollowUpCancel) btnFollowUpCancel.addEventListener("click", hideFollowUpPanel);
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
      refreshCurrentView().catch((err) => showIdeTabsError(err.message));
    });
  }
  const btnIdeDetailRefresh = document.getElementById("btn-ide-detail-refresh");
  if (btnIdeDetailRefresh) {
    btnIdeDetailRefresh.addEventListener("click", () => {
      refreshCurrentView().catch((err) => showIdeDetailError(err.message));
    });
  }
  for (const btn of document.querySelectorAll(".cockpit-ide-list-btn")) {
    btn.addEventListener("click", () => {
      const mode = btn.dataset.ideListMode;
      if (mode) setIdeListMode(mode);
    });
  }
  syncIdeListModeToggle();

  // M2.2.3 wireup: + New agent toolbar button, compose textarea + Send,
  // close-confirm modal cancel buttons, new-agent modal form + cancels.
  const btnNewAgent = document.getElementById("btn-ide-new-agent");
  if (btnNewAgent) {
    btnNewAgent.addEventListener("click", () => setView("new-agent"));
  }
  wireComposeUi();
  const newAgentForm = document.getElementById("form-new-agent");
  if (newAgentForm) newAgentForm.addEventListener("submit", handleNewAgentSubmit);
  for (const id of ["btn-new-agent-cancel", "btn-new-agent-cancel-bottom"]) {
    const el = document.getElementById(id);
    if (el) el.addEventListener("click", () => setView("ide-tabs"));
  }
  const btnCloseCancel = document.getElementById("btn-close-cancel");
  if (btnCloseCancel) btnCloseCancel.addEventListener("click", () => {
    // Modal overlays the detail view; restore the detail without losing
    // scroll position by simply hiding the modal and re-showing the
    // existing detail view.
    setView("ide-tab-detail", { composerId: activeIdeTabComposerId });
  });

  window.addEventListener("hashchange", () => {
    if (suppressHashSync) return;
    applyHashRoute();
  });
  applyHashRoute();
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
