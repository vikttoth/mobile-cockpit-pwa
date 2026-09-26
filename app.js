// mobile-cockpit / pwa / app.js
//
// v1's "Sessions" (list/detail/new, state.json-backed) and "App" (static
// cost-notice) surfaces were removed 2026-09-26
// (SPEC-DELTA-2026-09-26-ui-cleanup, item A) -- no in-flight v1 session
// needed migrating, so this was a clean deletion, same shape as the
// 2026-09-21 extension/ amputation (see SPEC.md). What's live now:
//   - Load config.json
//   - MSAL.js v4 PKCE auth (silent first; redirect on cache miss)
//   - The v2 (chat-model) surface: cursor-cockpit/sessions.json (light
//     index) + cursor-cockpit/sessions/<id>.json (full record), scrollback,
//     live streaming, queue/force/stop, model/mode switch, sharing,
//     archive/unarchive, chat naming + rename (lib/transcript-model.mjs)
//   - The read-only IDE-tabs mirror (M2.1, AC-022)
//   - "Shared with me" (SPEC-DELTA-2026-09-25-session-sharing-stage1)
//   - Hash routing (#v2-list / #v2-new / #v2-detail/<id> / ...) for
//     deep links + bookmarks
//   - Refresh button + auto-refresh, per-view poll intervals
//
// Reference order while reading this file:
//   1. ../SPEC.md — state model, acceptance criteria, scenario history
//   2. ./transcript-model.mjs — the pure v2 record/index model (byte-mirror
//      of ../lib/transcript-model.mjs)
//   3. ./write-helpers.mjs — pure hash-routing helpers + cryptoRandomBytes
//
// Style: vanilla JS, no framework, no bundler. ES2020. Single file. MSAL
// is loaded from ./vendor/msal-browser.min.js (defer-ordered before this).
// Pure helpers live in sibling ESM modules (./write-helpers.mjs,
// ./transcript-model.mjs, ./scrollback-helpers.mjs, ./ide-helpers.mjs); we
// pull them in via dynamic import() inside bootstrap() so this file stays a
// classic script and the MSAL UMD bundle keeps its source-order guarantee.

"use strict";

// =============================================================================
// 0. Build stamp + module-level state
// =============================================================================
//
// BUILD_STAMP is replaced by the deploy script before upload (sed on
// `2026-09-26 16:32 CEST 9bed028`). Keep the string literal — index.html cache-busts on it.
const BUILD_STAMP = "2026-09-26 16:32 CEST 9bed028";

/** Loaded asynchronously from ./config.json at boot. See pwa/config.json. */
let CONFIG = null;

/** Loaded once by initMsal(). Reused for every subsequent token acquisition. */
let msalClient = null;

/** Cached after the first successful sign-in. */
let activeAccount = null;

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

/** Last-rendered full record for the open v2 detail view -- used by the
 *  rename control (handleV2RenameClick) to read the current title without
 *  an extra Graph round-trip. */
let activeV2DetailRecord = null;

/** Faster poll while viewing a v2 session detail (mirrors ideDetailFastTimerId). */
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

// SPEC-DELTA-2026-09-25-session-sharing-stage1: item-level sharing, the
// browser-side mirror of lib/graph-state.mjs#invitePath/listPermissionsPath/
// revokePermissionPath. Same path-addressed pattern as loadJson/putJson
// above (`${endpoint}:/invite` etc.), not a JSON-file read/write.

async function graphInvite(endpoint, { email, sendInvitation }) {
  const res = await graphFetch(`${endpoint}:/invite`, {
    method: "POST",
    body: JSON.stringify({
      recipients: [{ email }],
      requireSignIn: true,
      sendInvitation: sendInvitation ?? false,
      roles: ["read"],
    }),
  });
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`graphInvite: POST failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  return res.json().catch(() => null);
}

async function graphListPermissions(endpoint) {
  const res = await graphFetch(`${endpoint}:/permissions`);
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`graphListPermissions: GET failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  const json = await res.json().catch(() => null);
  return Array.isArray(json?.value) ? json.value : [];
}

async function graphRevokePermission(endpoint, permissionId) {
  const res = await graphFetch(`${endpoint}:/permissions/${encodeURIComponent(permissionId)}`, { method: "DELETE" });
  if (!res.ok && res.status !== 404) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`graphRevokePermission: DELETE failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
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
 *
 * `id` (SPEC-DELTA-2026-09-26-ui-cleanup, item C): no longer a required
 * user-typed field -- an internal id is generated automatically
 * (generateInternalSessionId()) when the caller doesn't supply one. The
 * DISPLAYED name is always the derived/sequential/custom title, never this
 * internal id.
 */
async function v2CreateSession({ id, cwd, model, mode, parentId, firstMessage }) {
  const now = Date.now();
  const finalId = id && id.trim() ? id.trim() : generateInternalSessionId(now);
  // SPEC-DELTA-2026-09-26-ui-cleanup: "Chat1/Chat2/..." default naming needs
  // a 1-based position at creation time -- read the index once, up front,
  // purely to count. Best-effort, not a strict global counter (see the
  // delta's open question 1: a rare race across hosts/deletions can produce
  // a duplicate or non-contiguous number; cosmetic only, title is never a
  // unique key).
  const { index: indexForCount } = await loadV2Index();
  const sequenceNumber = Array.isArray(indexForCount?.sessions) ? indexForCount.sessions.length + 1 : 1;
  const record = V2_MODEL.buildSessionRecord({
    id: finalId,
    chatId: null,
    model: model || null,
    mode: mode || null,
    cwd: cwd || null,
    worktree: null,
    parentId: parentId || null,
    sequenceNumber,
    now,
  });
  // Fresh id -- no retry needed, mirrors session-store.mjs#createSession.
  await putV2Record(finalId, record, null);
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
    await v2WriteRecordWithRetry(finalId, (r) =>
      V2_MODEL.enqueueMessage(r, { text: firstMessage.trim(), now: Date.now() }),
    );
  }
  return record;
}

/**
 * Internal session id, never shown to the user (item C) -- same shape as
 * write-helpers.mjs#generateSessionId's `pwa-` prefix convention, but kept
 * local to this file since it has no validation/merge counterpart to share
 * a module with anymore (v1's write-helpers.mjs exports were removed
 * alongside the rest of the v1 surface).
 */
function generateInternalSessionId(now) {
  const t = Math.floor(now).toString(36).padStart(8, "0");
  const rand = Math.random().toString(36).slice(2, 8);
  return `mcv2-${t}-${rand}`;
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

// SPEC-DELTA-2026-09-26-ui-cleanup: `archived` and `title` (derived from
// `customTitle`) are BOTH index-mirrored fields (buildIndexEntry), so these
// two use v2WriteRecordAndMirrorIndex -- the same helper v2SetModel/v2SetMode
// above deliberately do NOT need, since model/mode never leak into the index.
async function v2SetArchived(id, archived) {
  return v2WriteRecordAndMirrorIndex(id, (record) =>
    archived
      ? V2_MODEL.archiveSession(record, { now: Date.now() })
      : V2_MODEL.unarchiveSession(record, { now: Date.now() }),
  );
}

async function v2SetCustomTitle(id, customTitle) {
  return v2WriteRecordAndMirrorIndex(id, (record) =>
    V2_MODEL.setCustomTitle(record, { customTitle, now: Date.now() }),
  );
}

// v2TakeOverLease/v2HandBackLease/v2RenewLeaseHeartbeat were removed
// 2026-09-24 along with the whole lease feature -- see
// lib/transcript-model.mjs's note for why.

// SPEC-DELTA-2026-09-25-session-sharing-stage1: mirrors
// lib/session-store.mjs#requestSessionInvite/requestSessionSetSharingEnabled
// exactly -- grant/revoke the real Graph permission FIRST, then update the
// record via the pure transform, same "never claim an invite that didn't
// actually happen" ordering.

async function v2InviteToSession(id, email) {
  await graphInvite(v2RecordEndpoint(id), { email });
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.addSharedWithEntry(record, { email, now: Date.now() }));
}

async function v2SetSharingEnabled(id, enabled) {
  const endpoint = v2RecordEndpoint(id);
  if (enabled) {
    const { record } = await loadV2Record(id);
    if (!record) {
      const err = new Error(`v2 session not found: ${id}`);
      err.code = "SESSION_NOT_FOUND";
      throw err;
    }
    for (const entry of record.sharedWith ?? []) {
      await graphInvite(endpoint, { email: entry.email });
    }
  } else {
    const permissions = (await graphListPermissions(endpoint)).filter((p) => !p.roles?.includes("owner"));
    for (const permission of permissions) {
      await graphRevokePermission(endpoint, permission.id);
    }
  }
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.setSharingEnabled(record, { enabled, now: Date.now() }));
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
  for (const id of ["btn-ide-refresh", "btn-ide-detail-refresh"]) {
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
    if (view === "ide-tabs") {
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
    showIdeTabsError(err.message);
  } finally {
    setRefreshButtonsBusy(false);
  }
}

// =============================================================================
// 3. Pure helpers (sortable / testable)
// =============================================================================

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
    if (route.view === "v2-list") setView("v2-list");
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
  if (viewId === "ide-tabs") {
    syncIdeListModeToggle();
    renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
  } else if (viewId === "ide-tab-detail" && payload && payload.composerId) {
    renderIdeTabDetail(payload.composerId);
  } else if (viewId === "v2-list") {
    renderV2List().catch((err) => showV2ListError(err.message));
  } else if (viewId === "v2-detail" && payload && payload.sessionId) {
    activeV2DetailSessionId = payload.sessionId;
    renderV2Detail(payload.sessionId).catch((err) => showV2DetailError(err.message));
  } else if (viewId === "v2-new") {
    renderV2New();
  } else if (viewId === "shared-list") {
    renderSharedList().catch((err) => showSharedListError(err.message));
  } else if (viewId === "shared-detail" && payload && payload.driveId && payload.itemId) {
    renderSharedDetail(payload.driveId, payload.itemId).catch((err) => showSharedDetailError(err.message));
  }
  // Drop the cached composer ID when navigating away from the IDE detail
  // view so a stale value can't accidentally target the wrong tab on the
  // next render.
  if (viewId !== "ide-tab-detail") {
    activeIdeTabComposerId = null;
  }
  if (viewId !== "ide-tab-detail") {
    stopIdeDetailFastPoll();
  }
  if (viewId !== "v2-detail") {
    activeV2DetailSessionId = null;
    activeV2DetailRecord = null;
    stopV2DetailPoll();
  }
  if (viewId !== "shared-detail") {
    activeSharedItem = null;
  }
  if (
    viewId === "v2-list" ||
    viewId === "v2-new" ||
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
 * The exact "IDE models" Cursor's own IDE model picker offers today
 * (SPEC-DELTA-2026-09-26-ui-cleanup, item F -- captured from a live
 * screenshot Viktor sent, labelled "Cursor Models" with a "NEW" badge on
 * Grok 4.7 High). Everything else in CONFIG.session.modelOptions is
 * `cursor-agent`-only, from `cursor-agent --list-models` -- the CLI list is
 * meant to be the superset of what the IDE offers. Hardcoded here (not
 * config-driven) for the same reason the old modelGroupFor() was purely
 * derived rather than a schema field: this is a small, rarely-changing,
 * UI-only classification, not worth typing into 3 mirrored config copies.
 * Order matters -- this is the exact display order for the "IDE models"
 * optgroup (AC-049). Mirror in local-ui/app.js if that surface ever gets
 * the same two-group treatment (out of scope for this delta).
 */
const IDE_APPROVED_MODEL_IDS = ["auto", "composer-2.5", "cursor-grok-4.6-xhigh", "cursor-grok-4.7-high"];

function isIdeApprovedModel(id) {
  return IDE_APPROVED_MODEL_IDS.includes(id);
}

/**
 * Populate a <select> with CONFIG.session.modelOptions, split into two
 * `<optgroup>`s: "IDE models" first (AC-049's exact order), then "CLI
 * models (not IDE-approved)" for everything else, in the config's own
 * order. Idempotent (skips the rebuild once the option count already
 * matches) so calling this on every poll-driven re-render of the detail
 * view doesn't fight an open dropdown or discard the caller's subsequent
 * `.value` assignment.
 */
function populateModelSelect(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;
  const options = (CONFIG.session && CONFIG.session.modelOptions) || [];
  if (select.options.length === options.length) return;
  select.innerHTML = "";

  const byId = new Map(options.map((o) => [o.id, o]));
  const ideGroup = document.createElement("optgroup");
  ideGroup.label = "IDE models";
  for (const id of IDE_APPROVED_MODEL_IDS) {
    const o = byId.get(id);
    if (!o) continue; // config drift -- don't render a phantom option
    const opt = document.createElement("option");
    opt.value = o.id;
    opt.textContent = o.label;
    ideGroup.appendChild(opt);
  }
  if (ideGroup.childElementCount > 0) select.appendChild(ideGroup);

  const cliGroup = document.createElement("optgroup");
  cliGroup.label = "CLI models (not IDE-approved)";
  for (const { id, label } of options) {
    if (isIdeApprovedModel(id)) continue;
    const opt = document.createElement("option");
    opt.value = id;
    opt.textContent = label;
    cliGroup.appendChild(opt);
  }
  if (cliGroup.childElementCount > 0) select.appendChild(cliGroup);
}

/**
 * Change-guard for the model <select> (AC-050): picking a model from the
 * "CLI models (not IDE-approved)" group asks for confirmation before
 * committing; declining reverts the select to its prior value. Callers
 * track "prior value" themselves via `select.dataset.priorValue` (set on
 * every render and on every accepted change) rather than this function
 * owning any state, so it stays a pure DOM helper.
 *
 * @param {HTMLSelectElement|null} selectEl
 * @returns {boolean} true if the caller should proceed with the new value
 */
function guardModelSelectionChange(selectEl) {
  if (!selectEl) return true;
  const prior = selectEl.dataset.priorValue ?? "";
  const next = selectEl.value;
  if (next === prior || isIdeApprovedModel(next)) {
    selectEl.dataset.priorValue = next;
    return true;
  }
  const ok =
    typeof window !== "undefined" && typeof window.confirm === "function"
      ? window.confirm("This model isn't officially approved for the IDE — are you sure?")
      : true;
  if (!ok) {
    selectEl.value = prior;
    return false;
  }
  selectEl.dataset.priorValue = next;
  return true;
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

/**
 * Builds one <li> chat-list row, shared by the Active and Archive sections
 * (SPEC-DELTA-2026-09-26-ui-cleanup, item D) and previously duplicated
 * inline. `archived` picks the row's own archive/unarchive control label.
 */
function buildV2ListRow(s, { archived }) {
  const li = document.createElement("li");
  li.className = "cockpit-session-row";
  li.dataset.sessionId = s.id || "";
  li.tabIndex = 0;
  const goToDetail = () => setView("v2-detail", { sessionId: s.id });
  li.addEventListener("click", (ev) => {
    if (ev.target.closest("button")) return; // let the archive button handle its own click
    goToDetail();
  });
  li.addEventListener("keydown", (ev) => {
    if (ev.target.closest("button")) return;
    if (ev.key === "Enter" || ev.key === " ") {
      ev.preventDefault();
      goToDetail();
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

  const archiveBtn = document.createElement("button");
  archiveBtn.type = "button";
  archiveBtn.className = "cockpit-btn cockpit-row-archive-btn";
  archiveBtn.textContent = archived ? "Unarchive" : "Archive";
  archiveBtn.setAttribute("aria-label", `${archived ? "Unarchive" : "Archive"} ${s.title || "chat"}`);
  archiveBtn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    handleV2ArchiveToggle(s.id, !archived).catch((err) => showV2ListError(err.message));
  });
  li.appendChild(archiveBtn);

  return li;
}

async function renderV2List() {
  clearV2ListError();
  const ul = document.getElementById("v2-session-list");
  const empty = document.getElementById("v2-list-empty-state");
  const archiveUl = document.getElementById("v2-archive-list");
  const archiveEmpty = document.getElementById("v2-archive-empty-state");
  if (!ul || !empty) return;
  try {
    await loadV2Index();
  } catch (err) {
    showV2ListError(err.message);
    return;
  }
  const all = (cachedV2Index && cachedV2Index.sessions) || [];
  const byRecency = (a, b) => (Date.parse(b.updatedAt) || 0) - (Date.parse(a.updatedAt) || 0);

  // Active/Archive grouping (item D): the backend already fully supports
  // this (archived: boolean, archiveSession/unarchiveSession, PATCH
  // .../archived) -- this used to just filter archived sessions out
  // entirely (dead end, no way to see/restore them).
  const active = all.filter((s) => s && !s.archived).sort(byRecency);
  ul.innerHTML = "";
  empty.hidden = active.length > 0;
  for (const s of active) ul.appendChild(buildV2ListRow(s, { archived: false }));

  if (archiveUl) {
    const archived = all.filter((s) => s && s.archived).sort(byRecency);
    archiveUl.innerHTML = "";
    if (archiveEmpty) archiveEmpty.hidden = archived.length > 0;
    for (const s of archived) archiveUl.appendChild(buildV2ListRow(s, { archived: true }));
  }
}

/** Per-row Archive/Unarchive control handler (AC-047). */
async function handleV2ArchiveToggle(id, archived) {
  clearV2ListError();
  try {
    await v2SetArchived(id, archived);
  } catch (err) {
    showV2ListError(err.message);
    return;
  }
  await renderV2List().catch((err) => showV2ListError(err.message));
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
  closeV2SettingsSheet();
  closeChatSwitcher();
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
  activeV2DetailRecord = record;

  const set = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text == null ? "—" : String(text);
  };
  // AC-039: the friendly title, never the raw record.id/chatId slug, is the
  // primary displayed text -- both in the header and the settings sheet's
  // rename readout. "Chat id" (below) is a labeled diagnostic field, same
  // disclosure pattern as the IDE-tab detail view's Composer ID row.
  const title = V2_MODEL ? V2_MODEL.deriveTitle(record) : record.id;
  set("v2-detail-title", title);
  set("v2-settings-title-readout", title);
  set("v2-detail-status", record.status);
  set("v2-detail-chat-id", record.chatId || "(provisioning…)");

  populateModelSelect("v2-detail-model-input");
  const modelInput = document.getElementById("v2-detail-model-input");
  if (modelInput) {
    modelInput.value = record.model || "auto";
    modelInput.dataset.priorValue = modelInput.value;
  }
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
  // AC-042: the settings-gear shows a dot whenever sharing is on, visible
  // without opening the sheet.
  const settingsDot = document.getElementById("v2-settings-dot");
  if (settingsDot) settingsDot.hidden = !record.sharingEnabled;
  renderV2Messages(record);
  renderSharePanel(record);

  syncV2DetailPoll();
}

/** Collapses the settings sheet (AC-041) and resets its toggle affordance. */
function closeV2SettingsSheet() {
  const sheet = document.getElementById("v2-settings-sheet");
  const btn = document.getElementById("btn-v2-detail-settings");
  if (sheet) sheet.hidden = true;
  if (btn) btn.setAttribute("aria-expanded", "false");
}

/** Collapses the chat switcher (AC-040) and resets its toggle affordance. */
function closeChatSwitcher() {
  const panel = document.getElementById("v2-chat-switcher");
  const btn = document.getElementById("btn-v2-detail-title");
  if (panel) panel.hidden = true;
  if (btn) btn.setAttribute("aria-expanded", "false");
}

/**
 * Tapping the title (AC-040): a list of the host's OTHER open/recent v2
 * chats to jump to. Viktor explicitly chose this over a horizontal
 * tab-strip, which he judged too cramped on a phone screen.
 */
function renderChatSwitcher() {
  const list = document.getElementById("v2-chat-switcher-list");
  const empty = document.getElementById("v2-chat-switcher-empty");
  if (!list) return;
  list.innerHTML = "";
  const all = (cachedV2Index && cachedV2Index.sessions) || [];
  const others = all
    .filter((s) => s && s.id !== activeV2DetailSessionId && !s.archived)
    .sort((a, b) => (Date.parse(b.updatedAt) || 0) - (Date.parse(a.updatedAt) || 0));
  if (empty) empty.hidden = others.length > 0;
  for (const s of others) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.tabIndex = 0;
    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = (s.parentId ? "↳ " : "") + (s.title || "(untitled)");
    li.appendChild(title);
    const goToChat = () => {
      closeChatSwitcher();
      setView("v2-detail", { sessionId: s.id });
    };
    li.addEventListener("click", goToChat);
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        goToChat();
      }
    });
    list.appendChild(li);
  }
}

/** Simple rename control (AC-044): window.prompt(), same "simple" posture
 *  as the existing window.confirm() usage elsewhere in this file -- no new
 *  modal component for a rarely-used control. Submitting a blank value
 *  clears the override (setCustomTitle's own contract), reverting display
 *  to the derived title. */
async function handleV2RenameClick() {
  const id = activeV2DetailSessionId;
  if (!id || typeof window === "undefined" || typeof window.prompt !== "function") return;
  const current = activeV2DetailRecord ? V2_MODEL.deriveTitle(activeV2DetailRecord) : "";
  const next = window.prompt("Rename chat", current);
  if (next === null) return; // cancelled
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2SetCustomTitle(id, next);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

/** SPEC-DELTA-2026-09-25-session-sharing-stage1: renders the Share panel's
 * toggle + remembered invite list from whatever the record currently says.
 * Never throws -- a rendering-only function, same posture as renderV2Messages. */
function renderSharePanel(record) {
  const toggle = document.getElementById("v2-detail-sharing-toggle");
  if (toggle) toggle.checked = !!record.sharingEnabled;
  const list = document.getElementById("v2-detail-shared-list");
  const empty = document.getElementById("v2-share-empty-state");
  if (!list) return;
  list.innerHTML = "";
  const sharedWith = Array.isArray(record.sharedWith) ? record.sharedWith : [];
  if (empty) empty.hidden = sharedWith.length > 0;
  for (const entry of sharedWith) {
    const li = document.createElement("li");
    li.textContent = `${entry.email} — invited ${relativeTime(entry.invitedAt)}`;
    list.appendChild(li);
  }
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
  const cwdSelect = document.getElementById("v2-new-cwd");
  const modelInput = document.getElementById("v2-new-model");
  const modeSelect = document.getElementById("v2-new-mode");
  const messageInput = document.getElementById("v2-new-message");

  // Item C: no more "Session id" field -- v2CreateSession generates an
  // internal id automatically; the displayed name is the Chat1/Chat2/...
  // sequential default (or the first-message-derived title once one exists).
  populateModelSelect("v2-new-model");
  if (modelInput) {
    modelInput.value = getLastUsedModel();
    modelInput.dataset.priorValue = modelInput.value;
  }
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

// -----------------------------------------------------------------------------
// v2 UI handlers (mobile follow-along, 2026-09-24)
// -----------------------------------------------------------------------------

async function handleV2NewSubmit(ev) {
  ev.preventDefault();
  clearV2NewError();
  const cwdEl = document.getElementById("v2-new-cwd");
  const modelEl = document.getElementById("v2-new-model");
  const modeEl = document.getElementById("v2-new-mode");
  const messageEl = document.getElementById("v2-new-message");
  const submitBtn = document.getElementById("btn-v2-new-submit");
  // Item C: no "Session id" field to validate -- v2CreateSession mints an
  // internal id automatically.
  if (submitBtn) submitBtn.disabled = true;
  try {
    const record = await v2CreateSession({
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
  if (!guardModelSelectionChange(modelInput)) return; // AC-050: reverted, no-op
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

// SPEC-DELTA-2026-09-25-session-sharing-stage1.
async function handleV2ShareInviteSubmit(evt) {
  evt.preventDefault();
  const id = activeV2DetailSessionId;
  const emailInput = document.getElementById("v2-share-invite-email");
  if (!id || !emailInput) return;
  const email = emailInput.value.trim();
  if (!email) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2InviteToSession(id, email);
    emailInput.value = "";
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

async function handleV2SharingToggleChange() {
  const id = activeV2DetailSessionId;
  const toggle = document.getElementById("v2-detail-sharing-toggle");
  if (!id || !toggle) return;
  const enabled = toggle.checked;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2SetSharingEnabled(id, enabled);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

// =============================================================================
// 7c. Health badge (SPEC task 12, AC-008 -- wired 2026-09-25)
// =============================================================================
//
// GET .../cursor-cockpit/health.json -- the SAME file the daemon publishes
// to on every tick where silent MSAL auth just succeeded (see
// daemon/poll.mjs#publishDaemonHealth). Read-only here; the phone never
// writes this file. Never more than one status shown, matching
// buildHealthStatus's own "exactly one {state, remediation, checkedAt}"
// contract -- this function only renders whatever that one object says.

/** Renders one health status into #health-badge. Never throws. */
function renderHealthBadge(status) {
  const el = document.getElementById("health-badge");
  if (!el) return;
  if (!status || !status.state || status.state === "unknown") {
    el.hidden = true;
    return;
  }
  el.dataset.state = status.state;
  el.title = status.remediation || "";
  const labels = {
    ok: "auth ok",
    cache_missing: "auth: not signed in on laptop",
    cache_corrupt: "auth: cache corrupt",
    auth_failing: "auth: sign-in failing",
    expired: "auth: token expired",
    expiring_soon: "auth: expiring soon",
  };
  el.textContent = labels[status.state] || `auth: ${status.state}`;
  el.hidden = status.state === "ok";
}

/** Fetches the daemon-published health status and renders it. Never throws. */
async function loadHealth() {
  if (!CONFIG.health) return;
  try {
    const { json } = await loadJson(CONFIG.health.endpoint);
    renderHealthBadge(json);
  } catch {
    // Read-only, best-effort -- a failed fetch just leaves the badge as it
    // was (or hidden, if it never loaded), same "don't crash the rest of
    // the app over a secondary signal" posture as the IDE-tabs mirror.
  }
}

// =============================================================================
// 7d. "Shared with me" -- read-only guest view (SPEC-DELTA-2026-09-25-
//     session-sharing-stage1, AC-034/AC-035)
// =============================================================================
//
// A completely separate read path from the host's own v2 code above: the
// host reads their OWN drive (/me/drive/root:/...); a guest reads whatever
// Graph's own sharedWithMe surfaced, addressed by the OWNER's driveId/itemId
// (/drives/{driveId}/items/{itemId}/content), never the guest's own root.
// No composer, no model/mode/stop, no overflow menu anywhere in this
// section -- reusing v2-detail's write controls here would be a real
// access-control bug, not a UX slip.

let activeSharedItem = null;

function showSharedListError(message) {
  const el = document.getElementById("shared-list-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearSharedListError() {
  const el = document.getElementById("shared-list-error-state");
  if (el) el.hidden = true;
}
function showSharedDetailError(message) {
  const el = document.getElementById("shared-detail-error-state");
  if (!el) return;
  el.textContent = translateErrorMessage(message);
  el.hidden = false;
}
function clearSharedDetailError() {
  const el = document.getElementById("shared-detail-error-state");
  if (el) el.hidden = true;
}

/**
 * GET /me/drive/sharedWithMe, filtered to mobile-cockpit v2 session records
 * (AC-034: never any OTHER kind of file someone might have shared with this
 * account). `.json` is a loose filter, but this app has no other sharing
 * feature to collide with it.
 */
async function loadSharedWithMe() {
  const res = await graphFetch("/me/drive/sharedWithMe");
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`loadSharedWithMe: GET failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  const json = await res.json().catch(() => null);
  const items = Array.isArray(json?.value) ? json.value : [];
  return items
    .filter((item) => item.remoteItem && typeof item.remoteItem.name === "string" && item.remoteItem.name.endsWith(".json"))
    .map((item) => ({
      driveId: item.remoteItem.parentReference?.driveId,
      itemId: item.remoteItem.id,
      name: item.remoteItem.name,
    }))
    .filter((item) => item.driveId && item.itemId);
}

async function renderSharedList() {
  clearSharedListError();
  const ul = document.getElementById("shared-session-list");
  const empty = document.getElementById("shared-list-empty-state");
  if (!ul || !empty) return;
  let items;
  try {
    items = await loadSharedWithMe();
  } catch (err) {
    showSharedListError(err.message);
    return;
  }
  ul.innerHTML = "";
  if (items.length === 0) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;
  for (const item of items) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.tabIndex = 0;
    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    // AC-039: never the raw OneDrive filename (== the session id in
    // disguise) as the row's primary text -- fetch the real record and show
    // its derived/custom title instead, filled in asynchronously so the
    // list still paints immediately.
    title.textContent = "…";
    li.appendChild(title);
    loadSharedRecord(item.driveId, item.itemId)
      .then((record) => {
        title.textContent = (V2_MODEL && V2_MODEL.deriveTitle(record)) || "(shared session)";
      })
      .catch(() => {
        title.textContent = "(shared session)";
      });
    const goToDetail = () => setView("shared-detail", { driveId: item.driveId, itemId: item.itemId });
    li.addEventListener("click", goToDetail);
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        goToDetail();
      }
    });
    ul.appendChild(li);
  }
}

/** GET .../content on the OWNER's drive item -- never the guest's own /me/drive/root. */
async function loadSharedRecord(driveId, itemId) {
  const res = await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`);
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`loadSharedRecord: GET failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  return res.json();
}

async function renderSharedDetail(driveId, itemId) {
  clearSharedDetailError();
  activeSharedItem = { driveId, itemId };
  let record;
  try {
    record = await loadSharedRecord(driveId, itemId);
  } catch (err) {
    showSharedDetailError(err.message);
    return;
  }
  const titleEl = document.getElementById("shared-detail-title");
  // AC-039: never the raw record.id -- the derived/custom title (or a
  // neutral fallback) is always the primary displayed text.
  if (titleEl) titleEl.textContent = (V2_MODEL && V2_MODEL.deriveTitle(record)) || "(shared session)";
  renderSharedMessages(record);
}

/** Read-only render -- deliberately NOT renderV2Messages (no streaming bubble,
 * no write-path coupling of any kind; a shared record never has a composer). */
function renderSharedMessages(record) {
  const container = document.getElementById("shared-messages");
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
  if (messages.length === 0) {
    const empty = document.createElement("div");
    empty.className = "v2-message v2-role-system";
    empty.textContent = "No messages yet.";
    container.appendChild(empty);
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
      import("./write-helpers.mjs?v=9bed028"),
      import("./ide-helpers.mjs?v=9bed028"),
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

  // SPEC task 12 (AC-008): fire-and-forget, does not block the rest of
  // bootstrap -- a slow/failed health fetch must never delay sign-in.
  loadHealth();
  if (CONFIG.health && Number.isFinite(CONFIG.health.pollIntervalSeconds)) {
    setInterval(loadHealth, CONFIG.health.pollIntervalSeconds * 1000);
  }

  // Wire navigation + write-path buttons.
  // Back buttons use their `data-target-view` attribute so the IDE-tab
  // detail returns to the IDE-tabs list (not to the chat list).
  for (const back of document.querySelectorAll(".cockpit-back-btn")) {
    const target = back.dataset.targetView || "v2-list";
    back.addEventListener("click", () => setView(target));
  }

  // Mode toggle (Chat / Shared / IDE tabs) — read data-target-view so we
  // don't hard-code the mapping here.
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
  const btnSharedRefresh = document.getElementById("btn-shared-refresh");
  if (btnSharedRefresh) {
    btnSharedRefresh.addEventListener("click", () => {
      renderSharedList().catch((err) => showSharedListError(err.message));
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
  const v2NewModelInput = document.getElementById("v2-new-model");
  if (v2NewModelInput) {
    // AC-050: same confirm-before-switch guard as the settings sheet's model
    // select -- nothing to persist yet (the value is only read at submit),
    // so this just guards the <select>'s own value.
    v2NewModelInput.addEventListener("change", () => {
      guardModelSelectionChange(v2NewModelInput);
    });
  }
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
  // Header redesign (item B): settings-gear opens #v2-settings-sheet;
  // tapping the title opens the chat switcher. Mutually exclusive -- opening
  // one closes the other, same posture as the old single overflow sheet.
  const btnV2Settings = document.getElementById("btn-v2-detail-settings");
  const v2SettingsSheet = document.getElementById("v2-settings-sheet");
  if (btnV2Settings && v2SettingsSheet) {
    btnV2Settings.addEventListener("click", () => {
      const open = v2SettingsSheet.hidden;
      closeChatSwitcher();
      v2SettingsSheet.hidden = !open;
      btnV2Settings.setAttribute("aria-expanded", String(open));
    });
  }
  const btnV2Title = document.getElementById("btn-v2-detail-title");
  const v2ChatSwitcher = document.getElementById("v2-chat-switcher");
  if (btnV2Title && v2ChatSwitcher) {
    btnV2Title.addEventListener("click", () => {
      const open = v2ChatSwitcher.hidden;
      closeV2SettingsSheet();
      if (open) renderChatSwitcher();
      v2ChatSwitcher.hidden = !open;
      btnV2Title.setAttribute("aria-expanded", String(open));
    });
  }
  const btnV2Rename = document.getElementById("btn-v2-rename");
  if (btnV2Rename) {
    btnV2Rename.addEventListener("click", () => {
      handleV2RenameClick().catch((err) => showV2DetailError(err.message));
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
  const v2ShareInviteForm = document.getElementById("v2-share-invite-form");
  if (v2ShareInviteForm) {
    v2ShareInviteForm.addEventListener("submit", (evt) => {
      handleV2ShareInviteSubmit(evt).catch((err) => showV2DetailError(err.message));
    });
  }
  const v2SharingToggle = document.getElementById("v2-detail-sharing-toggle");
  if (v2SharingToggle) {
    v2SharingToggle.addEventListener("change", () => {
      handleV2SharingToggleChange().catch((err) => showV2DetailError(err.message));
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
// Pure helpers used by app.js live in several sibling ESM modules and ARE
// Node-importable:
//   - ./write-helpers.mjs        (hash routing + cryptoRandomBytes; unit-tested
//                                 by tests/flows/mobile-cockpit/pwa-write-helpers-unit.sh)
//   - ./ide-helpers.mjs          (M2.1 read-only IDE-tabs view; unit-tested by
//                                 tests/flows/mobile-cockpit/pwa-ide-helpers-unit.sh)
//   - ./transcript-model.mjs     (the v2 record/index model; unit-tested by
//                                 tests/flows/mobile-cockpit/transcript-model-unit.sh)
//   - ./scrollback-helpers.mjs   (message ordering/formatting)
//
// The local pure helpers in this file (relativeTime, statusClass,
// modelGroupFor-successor isIdeApprovedModel) are intentionally NOT
// module-exported here — they stay internal to the browser script. Promote
// them to write-helpers.mjs if you ever want to assert them from Node.
//
// The DOM-coupled paths (setView / renderV2List / renderV2Detail /
// renderIdeTabsList / renderIdeTabDetail / renderSharedList /
// renderSharedDetail + their button handlers) are NOT unit-testable in
// isolation; they are covered by structural presence/wiring assertions
// (tests/flows/mobile-cockpit/pwa-v2-structural.sh, pwa-html-structural.sh)
// and, for the write paths, by live-validation runs against the real
// OneDrive sessions.json / sessions/<id>.json.
