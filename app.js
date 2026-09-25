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
// `2026-09-25 08:10 CEST 4c877ad`). Keep the string literal — index.html cache-busts on it.
const BUILD_STAMP = "2026-09-25 08:10 CEST 4c877ad";

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

/** Faster poll while an IDE tab detail shows waitingOn=agent (mirror lag). */
let ideDetailFastTimerId = null;

/**
 * Composer ID currently shown in `view-ide-tab-detail`. Cleared when
 * leaving the detail view.
 */
let activeIdeTabComposerId = null;

/** When true, setView skips hash sync (hashchange handler is driving navigation). */
let suppressHashSync = false;

// -----------------------------------------------------------------------------
// Mobile follow-along (2026-09-24): v2 (chat-model) state -- a separate mode
// alongside v1 Sessions / IDE tabs / App, not a replacement. See SPEC.md's
// "Mobile PWA v2 groundwork" section.
// -----------------------------------------------------------------------------

/** Dynamically imported pure v2 model helpers -- ./transcript-model.mjs, a
 *  byte-for-byte mirror of lib/transcript-model.mjs (see
 *  pwa-transcript-model-coherence.sh). Populated by bootstrap(). */
let V2_MODEL = null;

/** Cached last-known v2 sessions.json index. */
let cachedV2Index = null;
let cachedV2IndexEtag = null;

/** Session id currently open in v2 detail view (for the fast poll + composer). */
let activeV2DetailSessionId = null;

/** Faster poll while viewing a v2 session detail (mirrors runningDetailTimerId). */
let v2DetailTimerId = null;

/** Auto-refresh handle for the v2 list view (independent of v1's refreshTimerId). */
let v2RefreshTimerId = null;

/** Dynamically imported pure scrollback helpers (already existed, unwired
 *  until now) -- orderMessagesForDisplay / formatSessionMessage / AC-019. */
let SCROLLBACK_HELPERS = null;

function orderMessagesForDisplay(messages) {
  return SCROLLBACK_HELPERS ? SCROLLBACK_HELPERS.orderMessagesForDisplay(messages) : [];
}
function formatSessionMessage(message) {
  return SCROLLBACK_HELPERS
    ? SCROLLBACK_HELPERS.formatSessionMessage(message)
    : { role: "unknown", label: "System", text: "", ts: null };
}

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

async function graphFetch(path, init = {}, opts = {}) {
  const timeoutMs = opts.timeoutMs ?? 45000;
  const token = await getAccessToken();
  const headers = new Headers(init.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(`${CONFIG.graph.base}${path}`, {
      ...init,
      headers,
      signal: controller.signal,
    });
    return res;
  } catch (e) {
    if (e && e.name === "AbortError") {
      throw new Error(
        `Graph request timed out after ${Math.round(timeoutMs / 1000)}s — check network/VPN`,
      );
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
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
// 2c. v2 Graph helpers — generic JSON read/write (mobile follow-along, 2026-09-24)
// =============================================================================
//
// Generalizes loadState/putState's exact pattern (two-call metadata+content
// read, If-Match write, 412 -> PRECONDITION_FAILED) to an arbitrary Graph
// endpoint, so sessions.json (the light index) and sessions/<id>.json (each
// full record) can share one retry/error shape instead of duplicating it.
// loadState/putState above are UNTOUCHED -- v1 is unaffected by this.

async function loadJson(endpoint) {
  const meta = await graphFetch(`${endpoint}`);
  if (meta.status === 404) return { json: null, etag: null };
  if (!meta.ok) {
    throw new Error(`Graph driveItem GET failed: ${meta.status} ${meta.statusText}`);
  }
  const metaJson = await meta.json();
  const etag = metaJson.eTag || metaJson["@odata.etag"] || null;
  const contentRes = await graphFetch(`${endpoint}:/content`);
  if (!contentRes.ok) {
    throw new Error(`Graph content GET failed: ${contentRes.status} ${contentRes.statusText}`);
  }
  const json = await contentRes.json();
  return { json, etag };
}

async function putJson(endpoint, json, etagOrNull) {
  const headers = new Headers({ "Content-Type": "application/json" });
  if (etagOrNull) headers.set("If-Match", etagOrNull);
  const body = JSON.stringify(json, null, 2) + "\n";
  const res = await graphFetch(`${endpoint}:/content`, { method: "PUT", body, headers });
  if (res.status === 412) {
    const err = new Error("changed since last read (412)");
    err.code = "PRECONDITION_FAILED";
    err.status = 412;
    throw err;
  }
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`putJson: PUT failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  return res.json().catch(() => null);
}

/** Mirrors lib/config.mjs#sessionRecordRelativePath(id), endpoint-shaped. */
function v2RecordEndpoint(id) {
  return `${CONFIG.sessionsDir.endpoint}/${id}.json`;
}

async function loadV2Index() {
  const { json, etag } = await loadJson(CONFIG.sessionsIndex.endpoint);
  cachedV2Index = json && Array.isArray(json.sessions) ? json : { schemaVersion: 1, sessions: [] };
  cachedV2IndexEtag = etag;
  return { index: cachedV2Index, etag };
}

async function putV2Index(indexObj, etagOrNull) {
  return putJson(CONFIG.sessionsIndex.endpoint, indexObj, etagOrNull);
}

async function loadV2Record(id) {
  return loadJson(v2RecordEndpoint(id)).then(({ json, etag }) => ({ record: json, etag }));
}

async function putV2Record(id, recordObj, etagOrNull) {
  return putJson(v2RecordEndpoint(id), recordObj, etagOrNull);
}

// =============================================================================
// 2d. v2 write actions — read-modify-write against sessions.json / sessions/<id>.json
// =============================================================================
//
// Mirrors lib/session-store.mjs's exact shape, using the SAME pure transforms
// (V2_MODEL == pwa/transcript-model.mjs, the byte-for-byte browser mirror of
// lib/transcript-model.mjs) -- this IS the browser-side session-store.

async function v2WriteRecordWithRetry(id, transformFn) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const { record, etag } = await loadV2Record(id);
    if (!record) {
      const err = new Error(`v2 session not found: ${id}`);
      err.code = "SESSION_NOT_FOUND";
      throw err;
    }
    const next = transformFn(record);
    try {
      await putV2Record(id, next, etag);
      return next;
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) throw err;
    }
  }
  throw new Error(`v2WriteRecordWithRetry(${id}): retries exhausted`);
}

/** Same as v2WriteRecordWithRetry, but for intents that ALSO change a field
 *  in buildIndexEntry's shape (owner, archived, chatId, parentId, status,
 *  title) -- mirrors lib/session-store.mjs#writeRecordAndMirrorIndex. */
async function v2WriteRecordAndMirrorIndex(id, transformFn) {
  const next = await v2WriteRecordWithRetry(id, transformFn);
  for (let attempt = 0; attempt < 2; attempt++) {
    const { index, etag } = await loadV2Index();
    const nextIndex = V2_MODEL.upsertIndexEntry(index, V2_MODEL.buildIndexEntry(next));
    try {
      await putV2Index(nextIndex, etag);
      return next;
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) throw err;
    }
  }
  throw new Error(`v2WriteRecordAndMirrorIndex(${id}): index retries exhausted`);
}

/**
 * Create a brand-new v2 session directly against Graph. `chatId` starts
 * null -- AC-001's real mint needs `cursor-agent create-chat`, which only
 * the daemon's machine can run; `daemon/lib/v2-action-tick.mjs`'s
 * provisioning step fills it in within one `--v2-tick` cycle (see SPEC.md's
 * "Mobile PWA v2 groundwork"). `parentId` is the sub-agent wiring: set when
 * this session is a delegated parallel task started from an existing
 * session's detail view.
 */
async function v2CreateSession({ id, cwd, model, mode, parentId, firstMessage }) {
  const now = Date.now();
  const record = V2_MODEL.buildSessionRecord({
    id,
    chatId: null,
    model: model || null,
    mode: mode || null,
    cwd: cwd || null,
    worktree: null,
    parentId: parentId || null,
    now,
  });
  // Fresh id -- no retry needed, mirrors session-store.mjs#createSession.
  await putV2Record(id, record, null);
  for (let attempt = 0; attempt < 2; attempt++) {
    const { index, etag } = await loadV2Index();
    const nextIndex = V2_MODEL.upsertIndexEntry(index, V2_MODEL.buildIndexEntry(record));
    try {
      await putV2Index(nextIndex, etag);
      break;
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) throw err;
    }
  }
  if (typeof firstMessage === "string" && firstMessage.trim()) {
    // Nothing is running yet -- Queue is the correct primitive here
    // (Force's "kill the in-flight turn" has nothing to kill on a
    // brand-new session).
    await v2WriteRecordWithRetry(id, (r) =>
      V2_MODEL.enqueueMessage(r, { text: firstMessage.trim(), now: Date.now() }),
    );
  }
  return record;
}

async function v2EnqueueMessage(id, text) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.enqueueMessage(record, { text, now: Date.now() }));
}

async function v2RequestForce(id, text) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.requestForce(record, { text, now: Date.now() }));
}

async function v2RequestStop(id) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.requestStop(record, { now: Date.now() }));
}

async function v2SetModel(id, model) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.setModel(record, { model, now: Date.now() }));
}

async function v2SetMode(id, mode) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.setMode(record, { mode, now: Date.now() }));
}

// v2TakeOverLease/v2HandBackLease/v2RenewLeaseHeartbeat were removed
// 2026-09-24 along with the whole lease feature -- see
// lib/transcript-model.mjs's note for why.

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
    } else if (view === "app-info") {
      // Static help view — nothing to refresh.
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
    else if (route.view === "app-info") setView("app-info");
    else if (route.view === "detail" && route.sessionId) {
      setView("detail", { sessionId: route.sessionId });
    } else if (route.view === "v2-list") setView("v2-list");
    else if (route.view === "v2-new") setView("v2-new");
    else if (route.view === "v2-detail" && route.sessionId) {
      setView("v2-detail", { sessionId: route.sessionId });
    }
  } finally {
    suppressHashSync = false;
  }
}

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
    activeDetailSessionId = payload.sessionId;
    renderDetail(payload.sessionId);
    syncRunningDetailPoll();
  } else if (viewId === "new") {
    renderNew();
  } else if (viewId === "ide-tabs") {
    syncIdeListModeToggle();
    renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
  } else if (viewId === "app-info") {
    // Static copy in index.html — no network render.
  } else if (viewId === "ide-tab-detail" && payload && payload.composerId) {
    renderIdeTabDetail(payload.composerId);
  } else if (viewId === "v2-list") {
    renderV2List().catch((err) => showV2ListError(err.message));
  } else if (viewId === "v2-detail" && payload && payload.sessionId) {
    activeV2DetailSessionId = payload.sessionId;
    renderV2Detail(payload.sessionId).catch((err) => showV2DetailError(err.message));
  } else if (viewId === "v2-new") {
    renderV2New();
  }
  // Drop the cached composer ID when navigating away from the IDE detail
  // view so a stale value can't accidentally target the wrong tab on the
  // next render.
  if (viewId !== "ide-tab-detail") {
    activeIdeTabComposerId = null;
  }
  if (viewId !== "detail") {
    activeDetailSessionId = null;
    stopRunningDetailPoll();
  }
  if (viewId !== "ide-tab-detail") {
    stopIdeDetailFastPoll();
  }
  if (viewId !== "v2-detail") {
    activeV2DetailSessionId = null;
    stopV2DetailPoll();
  }
  if (
    viewId === "list" ||
    viewId === "new" ||
    viewId === "app-info" ||
    viewId === "v2-list" ||
    viewId === "v2-new" ||
    (viewId === "detail" && payload && payload.sessionId) ||
    (viewId === "v2-detail" && payload && payload.sessionId)
  ) {
    syncHashForView(viewId, payload);
  }
}

/**
 * Turns a known technical error into a short, plain-English sentence for
 * someone who isn't going to recognize "AADSTS50005" or "PRECONDITION_
 * FAILED" (Viktor's ask, 2026-09-24, after the raw MSAL error string
 * showed up in the UI verbatim during tonight's live testing). Matches on
 * substrings actually seen coming out of this app's own error paths --
 * NEVER swallows an unrecognized message, always falls back to the raw
 * text unchanged, so a new/unexpected error is still fully visible, just
 * not translated.
 */
function translateErrorMessage(raw) {
  const message = typeof raw === "string" ? raw : String(raw ?? "");
  const rules = [
    [/MSAL silent auth failed|AADSTS/i,
      "Can't connect to your Microsoft account right now. This usually clears up within a few minutes on its own; tell Viktor if it keeps happening."],
    [/PRECONDITION_FAILED|412/,
      "Someone else (or another tab) changed this at the same moment. Reload and try again."],
    [/SESSION_NOT_FOUND/,
      "This session doesn't exist anymore — it may have been deleted."],
    [/failed to fetch|networkerror|load failed/i,
      "Can't reach the server. Check your connection and try again."],
  ];
  for (const [pattern, friendly] of rules) {
    if (pattern.test(message)) return friendly;
  }
  return message;
}

function showListError(message) {
  const el = document.getElementById("list-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}

function clearListError() {
  const el = document.getElementById("list-error-state");
  if (el) el.hidden = true;
}

function showDetailError(message) {
  const el = document.getElementById("detail-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}

function clearDetailError() {
  const el = document.getElementById("detail-error-state");
  if (el) el.hidden = true;
}

function showNewError(message) {
  const el = document.getElementById("new-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
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
  // Clear stale form state BEFORE populating -- form.reset() reverts a
  // <select> to its first <option> in DOM order, which would silently
  // undo populateModelSelect's last-used-model default if it ran first.
  const form = document.getElementById("new-session-form");
  if (form) form.reset();
  populateCwdSelect();
  populateModelSelect("new-model");
  const modelSelect = document.getElementById("new-model");
  if (modelSelect) modelSelect.value = getLastUsedModel();
}

function populateCwdSelect() {
  const select = document.getElementById("new-cwd");
  if (!select) return;
  const allowed = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  // Rebuild: one <option> per allowedCwd (first one pre-selected -- a
  // concrete, always-valid path beats an ambiguous "(daemon default)"
  // that gives no visible confirmation it resolved to anything real,
  // 2026-09-24), then "(daemon default)" last for anyone who explicitly
  // wants to defer to the daemon's own fallback. Idempotent -- safe to
  // call on every render.
  select.innerHTML = "";
  for (const cwd of allowed) {
    const opt = document.createElement("option");
    opt.value = cwd;
    opt.textContent = cwd;
    select.appendChild(opt);
  }
  const defaultOpt = document.createElement("option");
  defaultOpt.value = "";
  defaultOpt.textContent = "(daemon default)";
  select.appendChild(defaultOpt);
  if (allowed.length > 0) select.value = allowed[0];
}

// Model picker (2026-09-24): Viktor asked for "the full list of all
// selectable models, like it is in the [Cursor] UI" rather than a free-text
// guess-the-id field, defaulting to "Auto" or whatever was picked last.
// CONFIG.session.modelOptions is the full `cursor-agent --list-models`
// catalog (mirrored from lib/config.mjs#MODEL_OPTIONS).
const LAST_MODEL_STORAGE_KEY = "mc-last-model";

function getLastUsedModel() {
  try {
    return localStorage.getItem(LAST_MODEL_STORAGE_KEY) || "auto";
  } catch (_err) {
    return "auto";
  }
}

function setLastUsedModel(model) {
  try {
    if (model) localStorage.setItem(LAST_MODEL_STORAGE_KEY, model);
  } catch (_err) {
    // best-effort only -- a blocked localStorage must not break model switching
  }
}

/**
 * Populate a <select> with CONFIG.session.modelOptions. Idempotent (skips
 * the rebuild once the option count already matches) so calling this on
 * every poll-driven re-render of the detail view doesn't fight an open
 * dropdown or discard the caller's subsequent `.value` assignment.
 */
/**
 * Purely derived from the id string, not a schema field -- adding a
 * `group` to every MODEL_OPTIONS entry would mean typing it 57 times
 * across 3 mirrored copies (lib/config.mjs, pwa/config.json,
 * local-ui/app.js) for something a 6-line prefix check already gives for
 * free. Mirrored as-is in local-ui/app.js (same reasoning as that file's
 * other small duplicated helpers -- it depends on nothing under pwa/).
 */
function modelGroupFor(id) {
  if (id === "auto") return null; // ungrouped, always first
  if (id.startsWith("claude-")) return "Claude";
  if (id.startsWith("gpt-")) return "GPT";
  if (id.startsWith("cursor-grok-")) return "Grok";
  if (id.startsWith("gemini-")) return "Gemini";
  return "Other";
}

function populateModelSelect(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;
  const options = (CONFIG.session && CONFIG.session.modelOptions) || [];
  if (select.options.length === options.length) return;
  select.innerHTML = "";
  const groups = new Map();
  for (const { id, label } of options) {
    const opt = document.createElement("option");
    opt.value = id;
    opt.textContent = label;
    const groupName = modelGroupFor(id);
    if (!groupName) {
      select.appendChild(opt);
      continue;
    }
    if (!groups.has(groupName)) {
      const og = document.createElement("optgroup");
      og.label = groupName;
      groups.set(groupName, og);
      select.appendChild(og);
    }
    groups.get(groupName).appendChild(opt);
  }
}

// -----------------------------------------------------------------------------
// IDE-tabs view (M2.1, read-only)
// -----------------------------------------------------------------------------

function showIdeTabsError(message) {
  const el = document.getElementById("ide-tabs-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearIdeTabsError() {
  const el = document.getElementById("ide-tabs-error-state");
  if (el) el.hidden = true;
}
function showIdeDetailError(message) {
  const el = document.getElementById("ide-detail-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
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
    const emptyTab = IDE_HELPERS.isEmptyIdeTab(t);
    status.dataset.waitingOn = emptyTab ? "none" : (t.waitingOn || "none");
    if (emptyTab) status.dataset.emptyTab = "true";
    status.textContent = IDE_HELPERS.ideTabStatusLabel(t);
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

  // Cache so a background re-render (fast poll / refresh) can re-target
  // the same tab; cleared by setView when navigating away.
  activeIdeTabComposerId = tab.composerId;

  syncIdeDetailFastPoll();

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

  renderIdeTabThread(tab);
}

function renderIdeTabThread(tab) {
  const threadOl = document.getElementById("ide-detail-thread");
  const threadEmpty = document.getElementById("ide-detail-thread-empty");
  if (!threadOl || !threadEmpty || !IDE_HELPERS) return;
  threadOl.innerHTML = "";
  const thread = Array.isArray(tab.thread) ? tab.thread : [];
  if (thread.length === 0) {
    threadEmpty.hidden = false;
    return;
  }
  threadEmpty.hidden = true;
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

// -----------------------------------------------------------------------------
// v2 (chat-model) views (mobile follow-along, 2026-09-24)
// -----------------------------------------------------------------------------

function showV2ListError(message) {
  const el = document.getElementById("v2-list-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearV2ListError() {
  const el = document.getElementById("v2-list-error-state");
  if (el) el.hidden = true;
}
function showV2DetailError(message) {
  const el = document.getElementById("v2-detail-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearV2DetailError() {
  const el = document.getElementById("v2-detail-error-state");
  if (el) el.hidden = true;
}
function showV2NewError(message) {
  const el = document.getElementById("v2-new-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearV2NewError() {
  const el = document.getElementById("v2-new-error-state");
  if (el) el.hidden = true;
}

async function renderV2List() {
  clearV2ListError();
  const ul = document.getElementById("v2-session-list");
  const empty = document.getElementById("v2-list-empty-state");
  if (!ul || !empty) return;
  try {
    await loadV2Index();
  } catch (err) {
    showV2ListError(err.message);
    return;
  }
  const all = (cachedV2Index && cachedV2Index.sessions) || [];
  const visible = all.filter((s) => s && !s.archived);
  const sorted = [...visible].sort((a, b) => (Date.parse(b.updatedAt) || 0) - (Date.parse(a.updatedAt) || 0));
  ul.innerHTML = "";
  if (sorted.length === 0) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;
  for (const s of sorted) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.dataset.sessionId = s.id || "";
    li.tabIndex = 0;
    li.addEventListener("click", () => setView("v2-detail", { sessionId: s.id }));
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        setView("v2-detail", { sessionId: s.id });
      }
    });

    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = (s.parentId ? "↳ " : "") + (s.title || "(untitled)");
    li.appendChild(title);

    const status = document.createElement("span");
    status.className = `cockpit-row-status ${statusClass(s.status)}`;
    status.dataset.status = s.status || "unknown";
    status.textContent = s.chatId ? (s.status || "unknown") : "provisioning…";
    li.appendChild(status);

    const time = document.createElement("time");
    time.className = "cockpit-row-time";
    if (s.updatedAt) time.dateTime = s.updatedAt;
    time.textContent = relativeTime(s.updatedAt);
    li.appendChild(time);

    ul.appendChild(li);
  }
}

// applyV2LeaseGating was removed 2026-09-24 along with the whole lease
// feature -- see lib/transcript-model.mjs's note for why.

/**
 * Viktor's ask (2026-09-24 evening): while an action is in flight, show a
 * working indicator and keep the next action from firing until it's clear
 * what state the session is actually in. Every caller re-syncs via a fresh
 * `renderV2Detail` in its own `finally` block (success OR error), which
 * re-derives the CORRECT disabled state from the real record afterward --
 * this function only owns the "busy right now" span, never the
 * after-the-fact state.
 */
function setV2Busy(busy) {
  const note = document.getElementById("v2-detail-busy");
  if (note) note.hidden = !busy;
  const ids = [
    "v2-detail-model-input",
    "v2-detail-mode-select",
    "v2-composer-text",
    "btn-v2-composer-send",
    "btn-v2-composer-force",
    "btn-v2-stop",
  ];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) el.disabled = busy;
  }
}

async function renderV2Detail(sessionId) {
  clearV2DetailError();
  // Always start collapsed when (re-)entering a session's detail view --
  // it staying open from a PREVIOUS session would be confusing, and the
  // poll tick below never touches this itself.
  const menuPanel = document.getElementById("v2-detail-menu");
  const menuBtn = document.getElementById("btn-v2-detail-menu");
  if (menuPanel) menuPanel.hidden = true;
  if (menuBtn) menuBtn.setAttribute("aria-expanded", "false");
  activeV2DetailSessionId = sessionId;
  let record;
  try {
    const [{ record: r }] = await Promise.all([loadV2Record(sessionId), loadV2Index()]);
    record = r;
  } catch (err) {
    showV2DetailError(err.message);
    return;
  }
  if (!record) {
    showV2ListError(`v2 session not found: ${sessionId}`);
    setView("v2-list");
    return;
  }

  const set = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text == null ? "—" : String(text);
  };
  set("v2-detail-title", record.id);
  set("v2-detail-status", record.status);
  set("v2-detail-chat-id", record.chatId || "(provisioning…)");

  populateModelSelect("v2-detail-model-input");
  const modelInput = document.getElementById("v2-detail-model-input");
  if (modelInput) modelInput.value = record.model || "auto";
  const modeSelect = document.getElementById("v2-detail-mode-select");
  if (modeSelect) modeSelect.value = record.mode || "agent";

  const btnStop = document.getElementById("btn-v2-stop");
  if (btnStop) {
    const stoppable = record.status === "running";
    btnStop.hidden = !stoppable;
    btnStop.onclick = stoppable ? () => handleV2StopClick(record.id) : null;
  }
  const btnForce = document.getElementById("btn-v2-composer-force");
  if (btnForce) btnForce.hidden = record.status !== "running";
  renderV2Messages(record);

  syncV2DetailPoll();
}

function renderV2Messages(record) {
  const container = document.getElementById("v2-messages");
  if (!container) return;
  container.innerHTML = "";
  const messages = orderMessagesForDisplay(record.messages);
  for (const raw of messages) {
    const m = formatSessionMessage(raw);
    const bubble = document.createElement("div");
    bubble.className = `v2-message v2-role-${m.role}`;
    const label = document.createElement("div");
    label.className = "v2-message-role-label";
    label.textContent = m.label;
    bubble.appendChild(label);
    const text = document.createElement("div");
    text.className = "v2-message-text";
    text.textContent = m.text;
    bubble.appendChild(text);
    container.appendChild(bubble);
  }
  if (messages.length === 0 && !record.streaming) {
    const empty = document.createElement("div");
    empty.className = "v2-message v2-role-system";
    empty.textContent = "No messages yet.";
    container.appendChild(empty);
  }
  // Mobile follow-along: the agent's reply only lands in messages[] once a
  // turn closes -- this renders whatever text the daemon's v2-action-tick
  // has flushed so far, refreshed by syncV2DetailPoll while open.
  if (record.streaming && record.streaming.text) {
    const bubble = document.createElement("div");
    bubble.className = "v2-message v2-role-assistant v2-message-streaming";
    const label = document.createElement("div");
    label.className = "v2-message-role-label";
    label.textContent = "Agent (typing…)";
    bubble.appendChild(label);
    const text = document.createElement("div");
    text.className = "v2-message-text";
    text.textContent = record.streaming.text;
    bubble.appendChild(text);
    container.appendChild(bubble);
  }
  if (Array.isArray(record.queue) && record.queue.length > 0) {
    const panel = document.createElement("div");
    panel.className = "v2-queue-panel";
    const title = document.createElement("div");
    title.className = "v2-queue-title";
    title.textContent = `Queued (${record.queue.length})`;
    panel.appendChild(title);
    for (const item of record.queue) {
      const row = document.createElement("div");
      row.className = "v2-queue-item";
      row.textContent = item.text;
      panel.appendChild(row);
    }
    container.appendChild(panel);
  }
}

// renderV2SubAgents was removed 2026-09-24: cursor-agent's own native
// local subagents (https://cursor.com/docs/subagents) make mobile-cockpit's
// separate, sequential-only "+ Start sub-agent" mechanism pointless -- see
// SPEC.md's dated note.

function renderV2New() {
  clearV2NewError();
  const idInput = document.getElementById("v2-new-id");
  const cwdSelect = document.getElementById("v2-new-cwd");
  const modelInput = document.getElementById("v2-new-model");
  const modeSelect = document.getElementById("v2-new-mode");
  const messageInput = document.getElementById("v2-new-message");

  if (idInput) idInput.value = `mcv2-${Date.now().toString(36)}`;
  populateModelSelect("v2-new-model");
  if (modelInput) modelInput.value = getLastUsedModel();
  if (modeSelect) modeSelect.value = "agent";
  if (messageInput) messageInput.value = "";
  populateV2CwdSelect();
}

function populateV2CwdSelect() {
  const select = document.getElementById("v2-new-cwd");
  if (!select) return;
  const allowed = (CONFIG.session && CONFIG.session.allowedCwds) || [];
  // Same "concrete path pre-selected, ambiguous default demoted to an
  // explicit opt-in" fix as populateCwdSelect (2026-09-24) -- see its
  // comment for why.
  select.innerHTML = "";
  for (const cwd of allowed) {
    const opt = document.createElement("option");
    opt.value = cwd;
    opt.textContent = cwd;
    select.appendChild(opt);
  }
  const defaultOpt = document.createElement("option");
  defaultOpt.value = "";
  defaultOpt.textContent = "(daemon default)";
  select.appendChild(defaultOpt);
  if (allowed.length > 0) select.value = allowed[0];
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

function stopIdeDetailFastPoll() {
  if (ideDetailFastTimerId !== null) {
    clearInterval(ideDetailFastTimerId);
    ideDetailFastTimerId = null;
  }
}

/** ~5s refresh while agent is thinking — default IDE poll is 20s. */
function syncIdeDetailFastPoll() {
  stopIdeDetailFastPoll();
  if (document.body.dataset.view !== "ide-tab-detail" || !cachedIdeSnapshot || !activeIdeTabComposerId) {
    return;
  }
  if (!IDE_HELPERS) return;
  const tab = IDE_HELPERS.findIdeTab(cachedIdeSnapshot, activeIdeTabComposerId);
  if (!tab || tab.waitingOn !== "agent") return;
  const sec = (CONFIG.pwa && CONFIG.pwa.runningPollIntervalSeconds) || 5;
  const intervalMs = Math.max(3, sec | 0) * 1000;
  ideDetailFastTimerId = setInterval(() => {
    if (document.body.dataset.view !== "ide-tab-detail" || !activeIdeTabComposerId) return;
    loadIdeTabs()
      .then(() => {
        renderIdeTabDetail(activeIdeTabComposerId, { preserveCompose: true });
      })
      .catch((err) => console.warn("ide detail fast poll failed:", err));
  }, intervalMs);
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

function stopV2DetailPoll() {
  if (v2DetailTimerId !== null) {
    clearInterval(v2DetailTimerId);
    v2DetailTimerId = null;
  }
}

/** Mirrors syncRunningDetailPoll, but v2's "is anything about to change?"
 *  is broader than v1's status==="running": also true while a chatId is
 *  still being provisioned (AC-001) or a queued message hasn't started
 *  yet, so the phone notices both promptly. */
function syncV2DetailPoll() {
  stopV2DetailPoll();
  if (document.body.dataset.view !== "v2-detail" || !activeV2DetailSessionId) return;
  const sec = (CONFIG.pwa && CONFIG.pwa.runningPollIntervalSeconds) || 5;
  const intervalMs = Math.max(3, sec | 0) * 1000;
  v2DetailTimerId = setInterval(() => {
    if (document.body.dataset.view !== "v2-detail" || !activeV2DetailSessionId) return;
    const id = activeV2DetailSessionId;
    Promise.all([loadV2Record(id), loadV2Index()])
      .then(([{ record }]) => {
        if (!record) return; // deleted elsewhere -- next manual nav will bounce to the list
        renderV2Messages(record);
        const set = (elId, text) => {
          const el = document.getElementById(elId);
          if (el) el.textContent = text == null ? "—" : String(text);
        };
        set("v2-detail-status", record.status);
        set("v2-detail-chat-id", record.chatId || "(provisioning…)");
        const btnStop = document.getElementById("btn-v2-stop");
        if (btnStop) {
          const stoppable = record.status === "running";
          btnStop.hidden = !stoppable;
          btnStop.onclick = stoppable ? () => handleV2StopClick(id) : null;
        }
        const btnForce = document.getElementById("btn-v2-composer-force");
        if (btnForce) btnForce.hidden = record.status !== "running";
        const stillChanging =
          record.status === "running" ||
          !record.chatId ||
          (Array.isArray(record.queue) && record.queue.length > 0);
        if (!stillChanging) stopV2DetailPoll();
      })
      .catch((err) => showV2DetailError(err.message));
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
  // Independent timer for the v2 session list (mobile follow-along,
  // 2026-09-24) -- same cadence as v1's list poll, separate handle so
  // switching modes never starts/stops the wrong one.
  if (v2RefreshTimerId === null) {
    const v2IntervalMs = Math.max(5, (CONFIG.pwa.pollIntervalSeconds | 0)) * 1000;
    v2RefreshTimerId = setInterval(() => {
      if (document.body.dataset.view === "v2-list") {
        renderV2List().catch((err) => showV2ListError(err.message));
      }
    }, v2IntervalMs);
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
    setLastUsedModel(model);
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

// -----------------------------------------------------------------------------
// v2 UI handlers (mobile follow-along, 2026-09-24)
// -----------------------------------------------------------------------------

async function handleV2NewSubmit(ev) {
  ev.preventDefault();
  clearV2NewError();
  const idEl = document.getElementById("v2-new-id");
  const cwdEl = document.getElementById("v2-new-cwd");
  const modelEl = document.getElementById("v2-new-model");
  const modeEl = document.getElementById("v2-new-mode");
  const messageEl = document.getElementById("v2-new-message");
  const submitBtn = document.getElementById("btn-v2-new-submit");
  const id = idEl ? idEl.value.trim() : "";
  if (!id) {
    showV2NewError("Session id is required");
    return;
  }
  if (submitBtn) submitBtn.disabled = true;
  try {
    const record = await v2CreateSession({
      id,
      cwd: cwdEl ? cwdEl.value : "",
      model: modelEl ? modelEl.value.trim() : "",
      mode: modeEl ? modeEl.value : "",
      parentId: null,
      firstMessage: messageEl ? messageEl.value : "",
    });
    setLastUsedModel(modelEl ? modelEl.value.trim() : "");
    setView("v2-detail", { sessionId: record.id });
  } catch (err) {
    showV2NewError(err.message);
  } finally {
    if (submitBtn) submitBtn.disabled = false;
  }
}

// Send always queues (safe default -- never interrupts a running turn
// without being asked). Force is a SEPARATE button, only ever visible
// while a turn is actually running (see renderV2Detail/the poll tick),
// mirroring how Claude's own composer behaves instead of a permanent
// Queue/Force mode selector (Viktor's ask, 2026-09-24).
async function handleV2ComposerSend() {
  const textEl = document.getElementById("v2-composer-text");
  const id = activeV2DetailSessionId;
  if (!id || !textEl) return;
  const text = textEl.value.trim();
  if (!text) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2EnqueueMessage(id, text);
    textEl.value = "";
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

async function handleV2ComposerForce() {
  const textEl = document.getElementById("v2-composer-text");
  const id = activeV2DetailSessionId;
  if (!id || !textEl) return;
  const text = textEl.value.trim();
  if (!text) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2RequestForce(id, text);
    textEl.value = "";
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

async function handleV2StopClick(id) {
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2RequestStop(id);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}


async function handleV2ModelChange() {
  const id = activeV2DetailSessionId;
  const modelInput = document.getElementById("v2-detail-model-input");
  if (!id || !modelInput) return;
  const model = modelInput.value.trim();
  if (!model) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2SetModel(id, model);
    setLastUsedModel(model);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

async function handleV2ModeChange() {
  const id = activeV2DetailSessionId;
  const modeSelect = document.getElementById("v2-detail-mode-select");
  if (!id || !modeSelect) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2SetMode(id, modeSelect.value);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
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
    [WRITE_HELPERS, IDE_HELPERS, REFRESH_HELPERS, V2_MODEL, SCROLLBACK_HELPERS] = await Promise.all([
      import("./write-helpers.mjs?v=4c877ad"),
      import("./ide-helpers.mjs?v=4c877ad"),
      import("./refresh-helpers.mjs"),
      import("./transcript-model.mjs"),
      import("./scrollback-helpers.mjs"),
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

  // v2 (chat-model) button wiring (mobile follow-along, 2026-09-24). Back
  // buttons + the mode-toggle pill are already generic (data-target-view),
  // handled by the loops above.
  const btnV2Refresh = document.getElementById("btn-v2-refresh");
  if (btnV2Refresh) {
    btnV2Refresh.addEventListener("click", () => {
      renderV2List().catch((err) => showV2ListError(err.message));
    });
  }
  const btnV2DetailRefresh = document.getElementById("btn-v2-detail-refresh");
  if (btnV2DetailRefresh) {
    btnV2DetailRefresh.addEventListener("click", () => {
      if (activeV2DetailSessionId) {
        renderV2Detail(activeV2DetailSessionId).catch((err) => showV2DetailError(err.message));
      }
    });
  }
  const btnV2NewSession = document.getElementById("btn-v2-new-session");
  if (btnV2NewSession) {
    btnV2NewSession.addEventListener("click", () => setView("v2-new"));
  }
  const btnV2NewCancel = document.getElementById("btn-v2-new-cancel");
  if (btnV2NewCancel) btnV2NewCancel.addEventListener("click", () => setView("v2-list"));
  const v2NewForm = document.getElementById("v2-new-session-form");
  if (v2NewForm) v2NewForm.addEventListener("submit", handleV2NewSubmit);
  const btnV2ComposerSend = document.getElementById("btn-v2-composer-send");
  if (btnV2ComposerSend) {
    btnV2ComposerSend.addEventListener("click", () => {
      handleV2ComposerSend().catch((err) => showV2DetailError(err.message));
    });
  }
  const btnV2ComposerForce = document.getElementById("btn-v2-composer-force");
  if (btnV2ComposerForce) {
    btnV2ComposerForce.addEventListener("click", () => {
      handleV2ComposerForce().catch((err) => showV2DetailError(err.message));
    });
  }
  const btnV2DetailMenu = document.getElementById("btn-v2-detail-menu");
  const v2DetailMenuPanel = document.getElementById("v2-detail-menu");
  if (btnV2DetailMenu && v2DetailMenuPanel) {
    btnV2DetailMenu.addEventListener("click", () => {
      const open = v2DetailMenuPanel.hidden;
      v2DetailMenuPanel.hidden = !open;
      btnV2DetailMenu.setAttribute("aria-expanded", String(open));
    });
  }
  const v2ModelInput = document.getElementById("v2-detail-model-input");
  if (v2ModelInput) {
    v2ModelInput.addEventListener("change", () => {
      handleV2ModelChange().catch((err) => showV2DetailError(err.message));
    });
  }
  const v2ModeSelect = document.getElementById("v2-detail-mode-select");
  if (v2ModeSelect) {
    v2ModeSelect.addEventListener("change", () => {
      handleV2ModeChange().catch((err) => showV2DetailError(err.message));
    });
  }

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
