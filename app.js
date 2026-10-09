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
// `2026-10-09 08:22 CEST 7d8a394`). Keep the string literal — index.html cache-busts on it.
const BUILD_STAMP = "2026-10-09 08:22 CEST 7d8a394";

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

/** Dynamically imported pure helpers for the daemon start/stop control
 *  strip (SPEC-DELTA-2026-09-27-daemon-control-watchdog) -- byte-for-byte
 *  mirror of ../lib/daemon-control-model.mjs (see
 *  pwa-daemon-control-coherence.sh). */
let DAEMON_CONTROL_MODEL = null;

/** Dynamically imported pure helpers for the global diagnostics panel
 *  (SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal) --
 *  byte-for-byte mirror of ../lib/cockpit-health-model.mjs. */
let COCKPIT_HEALTH_MODEL = null;

/** Pure helpers for the v2 composer buttons + status chip
 *  (SPEC-DELTA-2026-09-29-composer-contextual-buttons). */
let COMPOSER_STATE = null;

/** Pure helpers for the header account menu (SPEC-DELTA-2026-09-29-app-menu). */
let APP_MENU_STATE = null;

/** PWA-wide Graph back-off after a 429/503 (SPEC-DELTA-2026-09-29-pwa-polling-hygiene);
 *  null until the helper module loads. */
let graphBackoff = null;

/** Last-rendered v2 session status + in-flight flag, the two inputs
 *  applyComposerButtons() combines with the textbox content. */
let v2ComposerStatus = null;
let v2Busy = false;

/** True while a Start/Stop request write is in flight (prevents double-tap). */
let daemonControlRequestInFlight = false;

/** True while a manual ↻ refresh is in flight (prevents double-tap). */
let refreshInFlight = false;

/** Cached last-known ide-tabs.json snapshot. Refreshed by loadIdeTabs(). */
let cachedIdeSnapshot = null;

/** IDE tabs sub-view within the read-only mirror: live open vs chat history. */
let ideListMode = "open";

/**
 * IDE Tracker source: which read-only mirror the current view shows.
 * SPEC-DELTA-2026-09-28-claude-code-ide-tracker.md. "cursor" reads
 * CONFIG.ideTabs (unchanged); "claude-code" reads CONFIG.claudeCodeTabs (new).
 * Both write into the SAME cachedIdeSnapshot, so every existing render
 * function (renderIdeTabsList, renderIdeTabDetail, ...) needs no changes at
 * all -- only loadIdeTabs() picks which endpoint to fetch.
 */
let ideTrackerSource = "cursor";

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

/** Pure "is this model id one the CLI accepts?" helpers -- ./model-choice.mjs
 *  (SPEC-DELTA-2026-10-08-model-catalog-drift-guard). Populated by bootstrap(). */
let MODEL_CHOICE = null;

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
  const backoffLeft = graphBackoff ? graphBackoff.remainingMs() : 0;
  if (backoffLeft > 0) {
    const err = new Error(
      `OneDrive is throttling requests — retrying in ${Math.ceil(backoffLeft / 1000)}s`,
    );
    err.code = "GRAPH_BACKOFF";
    throw err;
  }
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
    if (graphBackoff && (res.status === 429 || res.status === 503)) {
      // OneDrive puts retryAfterSeconds in the JSON body; read a clone so
      // the caller still gets an unconsumed response.
      const body = await res.clone().text().catch(() => "");
      graphBackoff.note(res.status, res.headers.get("retry-after"), body);
    }
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
/**
 * SPEC-DELTA-2026-10-07-cockpit-pin-density-unified-view.md: extracted out
 * of loadIdeTabs() so the Tracker view can fetch all three IDE-mirror
 * endpoints independently, without touching the single-source-at-a-time
 * cachedIdeSnapshot/ideTrackerSource globals the Cursor/Claude Code/Browser
 * sub-views depend on.
 */
async function fetchSnapshotByConfigKey(configKey) {
  const endpoint = CONFIG && CONFIG[configKey] && CONFIG[configKey].endpoint;
  if (!endpoint) {
    throw new Error(`config.${configKey}.endpoint missing -- update pwa/config.json`);
  }
  const contentRes = await graphFetch(`${endpoint}:/content`);
  if (contentRes.status === 404) {
    return { schemaVersion: 1, snapshotAt: null, workspaceKey: null, workspacePath: null, tabs: [] };
  }
  if (!contentRes.ok) {
    throw new Error(`${configKey} GET failed: ${contentRes.status} ${contentRes.statusText}`);
  }
  return contentRes.json();
}

async function loadIdeTabs() {
  const configKey =
    ideTrackerSource === "claude-code" ? "claudeCodeTabs" : ideTrackerSource === "browser" ? "browserTabs" : "ideTabs";
  const snapshot = await fetchSnapshotByConfigKey(configKey);
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
  // 409 too, like lib/graph-state.mjs#isConflictStatus: OneDrive answers a
  // racing If-Match write with either, and a 409 used to surface as a raw
  // error instead of the retry (SPEC-DELTA-2026-09-29-session-sharing-stage2).
  if (res.status === 412 || res.status === 409) {
    const err = new Error(`changed since last read (${res.status})`);
    err.code = "PRECONDITION_FAILED";
    err.status = res.status;
    throw err;
  }
  if (!res.ok) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`putJson: PUT failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  return res.json().catch(() => null);
}

// The Stage 1 per-file Graph invite/permissions helpers were removed with the
// Stage 1 host path (SPEC-DELTA-2026-09-29-session-sharing-stage2). The only
// remaining invite is the GUEST's Connect step, which grants the host write
// access to a folder in the guest's own drive -- guest-app.mjs#inviteHost.

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

/**
 * Enqueue + Force mirror the index: the daemon tick only re-reads a record
 * whose index `updatedAt` changed (SPEC-DELTA-2026-09-29-graph-load-and-
 * backoff), so without the mirror a new message would wait for the tick's
 * once-a-minute full sweep.
 */
async function v2EnqueueMessage(id, text) {
  return v2WriteRecordAndMirrorIndex(id, (record) => V2_MODEL.enqueueMessage(record, { text, now: Date.now() }));
}

/**
 * SPEC-DELTA-2026-09-27-queue-remove-and-session-delete (S-006, AC-015
 * "remove"): remove one not-yet-sent message from `queue[]`. No index
 * mirroring -- removing never needs the daemon to pick anything up.
 */
async function v2RemoveQueuedMessage(id, queueItemId) {
  return v2WriteRecordWithRetry(id, (record) => V2_MODEL.removeQueuedMessage(record, { id: queueItemId }));
}

async function v2RequestForce(id, text) {
  return v2WriteRecordAndMirrorIndex(id, (record) => V2_MODEL.requestForce(record, { text, now: Date.now() }));
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

/**
 * SPEC-DELTA-2026-09-27-queue-remove-and-session-delete (AC-030): permanent
 * whole-session delete, browser-side mirror of
 * lib/session-store.mjs#deleteSession. Deletes the record's own driveItem
 * FIRST (a plain Graph DELETE, like the removed Stage 1 permission revoke, but against the
 * record endpoint itself rather than a `:/permissions/<id>` sub-resource --
 * same as scripts/delete-onedrive-pwa.mjs's `DELETE /me/drive/items/<id>`
 * pattern for removing a whole item, not its content), tolerating an
 * already-gone file (404) as success, THEN removes the row from the light
 * index with the same read-modify-write-retry-on-412 shape every other
 * index mutation here uses. Order matters: never claim a session is deleted
 * if the record file delete itself failed.
 *
 * @param {string} id
 * @returns {Promise<void>}
 */
async function v2DeleteSession(id) {
  const res = await graphFetch(v2RecordEndpoint(id), { method: "DELETE" });
  if (!res.ok && res.status !== 404) {
    const bodyText = await res.text().catch(() => "");
    throw new Error(`v2DeleteSession: DELETE failed ${res.status} ${res.statusText}: ${bodyText.slice(0, 300)}`);
  }
  for (let attempt = 0; attempt < 2; attempt++) {
    const { index, etag } = await loadV2Index();
    let nextIndex;
    try {
      nextIndex = V2_MODEL.removeIndexEntry(index, id);
    } catch (err) {
      if (err.code === "SESSION_NOT_FOUND") return; // already gone from the index too
      throw err;
    }
    try {
      await putV2Index(nextIndex, etag);
      return;
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt > 0) throw err;
    }
  }
  throw new Error(`v2DeleteSession(${id}): index retries exhausted`);
}

// SPEC task 15 (S-138, AC-163): the phone's "Take over" when the laptop holds
// the lease via mc. The phone cannot run turns itself, so taking over hands
// the lease back to the daemon (open question 1's default in
// SPEC-DELTA-2026-09-29-session-sharing-stage2). Mirrors the index so the
// daemon's v2 tick sees the owner change at once.
async function v2TakeOverFromDevice(id) {
  return v2WriteRecordAndMirrorIndex(id, (record) =>
    record.owner === "daemon" ? record : V2_MODEL.handBackLease(record, { now: Date.now() }),
  );
}

// SPEC-DELTA-2026-09-29-session-sharing-stage2: the sharing registry
// (cursor-cockpit/shares.json), same read-modify-write-with-retry shape as
// lib/share-store.mjs#mutateShares. The Stage 1 per-file Graph invite path
// (v2InviteToSession / v2SetSharingEnabled) was removed: in this tenant a
// guest can never read the host's file, see the SPEC-DELTA.
let SHARE_MODEL = null;
let SHARE_UI = null;
let cachedShares = null;
let cachedRelayStatus = null;

async function loadShares() {
  const endpoint = CONFIG.sharing && CONFIG.sharing.sharesEndpoint;
  if (!endpoint || !SHARE_MODEL) return { shares: null, etag: null };
  const { json, etag } = await loadJson(endpoint);
  cachedShares = json && Array.isArray(json.items) ? json : SHARE_MODEL.emptyShares(Date.now());
  return { shares: cachedShares, etag };
}

async function mutateSharesPwa(transformFn) {
  const endpoint = CONFIG.sharing.sharesEndpoint;
  for (let attempt = 0; attempt < 3; attempt++) {
    const { shares, etag } = await loadShares();
    const next = transformFn(shares);
    try {
      await putJson(endpoint, next, etag);
      cachedShares = next;
      return next;
    } catch (err) {
      if (err.code !== "PRECONDITION_FAILED" || attempt >= 2) throw err;
    }
  }
  throw new Error("mutateSharesPwa: retries exhausted");
}

/**
 * shares.json + relay status at most every 30 s from views that poll (chat
 * detail, IDE tab detail) -- the host's own share actions refresh the cache
 * directly, so a slow cadence here only affects the relay's connected flags.
 */
let sharesLoadedAt = 0;
async function refreshSharesIfStale(maxAgeMs = 30_000) {
  if (!SHARE_MODEL || Date.now() - sharesLoadedAt < maxAgeMs) return;
  sharesLoadedAt = Date.now();
  await Promise.all([loadShares(), loadRelayStatus().catch(() => null)]);
}

async function loadRelayStatus() {
  const endpoint = CONFIG.sharing && CONFIG.sharing.relayStatusEndpoint;
  if (!endpoint) return null;
  const res = await graphFetch(`${endpoint}:/content`);
  if (!res.ok) return null;
  cachedRelayStatus = await res.json().catch(() => null);
  return cachedRelayStatus;
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
  } else if (viewId === "shared-detail" && payload && payload.kind && payload.id) {
    renderSharedDetail(payload).catch((err) => showSharedDetailError(err.message));
  } else if (viewId === "tracker") {
    renderTrackerView().catch((err) => showTrackerError(err.message));
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
  if (viewId !== "shared-detail" && GUEST_APP) {
    GUEST_APP.leaveItem();
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

/** Ids currently offered by CONFIG.session.modelOptions (== lib/config.mjs#MODEL_OPTIONS). */
function modelOptionIds() {
  return ((CONFIG && CONFIG.session && CONFIG.session.modelOptions) || []).map((o) => o.id);
}

function getLastUsedModel() {
  let stored = "auto";
  try {
    stored = localStorage.getItem(LAST_MODEL_STORAGE_KEY) || "auto";
  } catch (_err) {
    stored = "auto";
  }
  // AC-275 (SPEC-DELTA-2026-10-08-model-catalog-drift-guard): a stored id the
  // catalog no longer offers must not preselect nothing -- it falls back to auto.
  return MODEL_CHOICE ? MODEL_CHOICE.resolveModelChoice(stored, modelOptionIds()).model : stored;
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
  // The stale-model option (syncStaleModelOption) is extra, not part of the catalog.
  const catalogCount = Array.from(select.options).filter((o) => !o.dataset.stale).length;
  if (catalogCount === options.length) return;
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
 * AC-277 (SPEC-DELTA-2026-10-08-model-catalog-drift-guard): a session whose
 * recorded model the catalog no longer offers would render a blank select.
 * Show it instead as one disabled, selected option ("<id> (unavailable — runs
 * as Auto)"); the daemon runs such a session as `auto` (AC-278). Idempotent,
 * and removes the option again once the record's model is a listed one.
 *
 * @param {HTMLSelectElement|null} select
 * @param {string|null|undefined} recordModel
 */
function syncStaleModelOption(select, recordModel) {
  if (!select) return;
  const stale = MODEL_CHOICE ? MODEL_CHOICE.describeStaleModelOption(recordModel, modelOptionIds()) : null;
  for (const o of Array.from(select.options)) {
    if (o.dataset.stale && (!stale || o.value !== stale.value)) o.remove();
  }
  if (!stale || Array.from(select.options).some((o) => o.dataset.stale)) return;
  const opt = document.createElement("option");
  opt.value = stale.value;
  opt.textContent = stale.label;
  opt.disabled = stale.disabled;
  opt.dataset.stale = "1";
  select.insertBefore(opt, select.firstChild);
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

/** SPEC-DELTA-2026-09-28-claude-code-ide-tracker.md: sync the Cursor/Claude
 *  Code switcher's aria-current to the in-memory ideTrackerSource. */
function syncIdeTrackerSourceToggle() {
  for (const btn of document.querySelectorAll(".cockpit-ide-tracker-btn")) {
    btn.setAttribute(
      "aria-current",
      btn.dataset.ideTrackerSource === ideTrackerSource ? "true" : "false",
    );
  }
}

function setIdeTrackerSource(source) {
  if (source !== "cursor" && source !== "claude-code" && source !== "browser") return;
  if (source === ideTrackerSource) return;
  ideTrackerSource = source;
  syncIdeTrackerSourceToggle();
  // SPEC-DELTA-2026-10-07-browser-tab-copilot-mirror.md: no Open/Archive
  // concept for browser tabs -- there's no transcript, just whatever is
  // currently open. Hide the toggle rather than leaving a no-op control.
  const openHistoryToggle = document.getElementById("btn-ide-list-open")?.closest("nav");
  if (openHistoryToggle) openHistoryToggle.hidden = source === "browser";
  renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
}

async function renderIdeTabsList() {
  if (ideTrackerSource === "browser") {
    return renderBrowserTabsList();
  }
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

  const workspaceLabel = document.getElementById("ide-tabs-workspace-label");
  if (workspaceLabel) {
    workspaceLabel.textContent = snapshot.workspacePath ? `Watching: ${snapshot.workspacePath}` : "";
  }

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
    // SPEC-DELTA-2026-10-06-ide-tracker-summary-and-open-link.md (AC-198/AC-199):
    // a row with a known deep link (today: claude-code-mirror only) jumps straight
    // to the real tool + real session instead of the cockpit's own read-only
    // preview. A row with no link (Cursor, today) first tries the local,
    // on-demand UI-Automation activation (SPEC-DELTA-2026-10-07-cursor-tab-
    // activate-via-ui-automation.md) -- only succeeds when physically at the
    // laptop with the daemon's local server running -- and falls back to
    // today's in-app preview otherwise.
    const hasLink = typeof t.link === "string" && t.link.length > 0;
    const activateRow = async () => {
      if (hasLink) {
        window.open(t.link, "_blank", "noopener");
        return;
      }
      const activated = await tryActivateCursorTab(t.title);
      if (!activated) setView("ide-tab-detail", { composerId: t.composerId });
    };
    li.addEventListener("click", activateRow);
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        activateRow();
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

    // AC-197: plain-language summary, shown in a fixed-size popover on hover
    // (desktop) or tap (touch) of this info affordance -- not a native tooltip.
    if (typeof t.summary === "string" && t.summary.length > 0) {
      const info = document.createElement("button");
      info.type = "button";
      info.className = "cockpit-row-info";
      info.setAttribute("aria-label", "Show summary");
      info.textContent = "ⓘ"; // circled small "i"
      info.addEventListener("click", (ev) => {
        ev.stopPropagation();
        toggleIdeRowSummaryPopover(info, t.summary);
      });
      info.addEventListener("mouseenter", () => showIdeRowSummaryPopover(info, t.summary));
      info.addEventListener("mouseleave", () => hideIdeRowSummaryPopover());
      li.appendChild(info);
    }

    ul.appendChild(li);
  }
}

// SPEC-DELTA-2026-10-07-cursor-tab-activate-via-ui-automation.md: the daemon's
// local control surface (daemon/poll.mjs#LOCAL_SERVER_PORT -- keep this port
// in lockstep with that constant, same cross-file-comment convention as
// CLAUDE_CODE_TABS_RELATIVE_PATH elsewhere in this flow). Reached directly by
// fetch() regardless of which origin served this page -- modern browsers
// exempt 127.0.0.1 from mixed-content blocking, and local-server.mjs answers
// CORS for this PWA's own known hosted origins.
const CURSOR_ACTIVATE_URL = "http://127.0.0.1:4127/api/cursor-tabs/activate";
const CURSOR_ACTIVATE_TIMEOUT_MS = 500;
// Tracker only (AC-241). Measured 2026-10-08 with a title matching no tab:
// Cursor activation ~2.7 s (UI Automation priming + tree walk), browser-tab
// activation ~0.75 s. A 500 ms budget made every Tracker click BOTH fall
// back (in-app view / open the url in a new tab) AND, a moment later,
// really jump -- a duplicate Copilot tab every time. The browser budget
// stays under Chromium's ~5 s user-activation window, so the window.open
// fallback is never popup-blocked. The IDE Tracker list views keep the
// 500 ms default (Viktor's Tracker-only rule).
const TRACKER_CURSOR_ACTIVATE_TIMEOUT_MS = 40000;
const TRACKER_BROWSER_ACTIVATE_TIMEOUT_MS = 4000;

/**
 * Best-effort: true only when the daemon's local server is up AND found
 * exactly one live Cursor tab matching `title`. Any failure (daemon not
 * running, not at the laptop, ambiguous/not-found match, `timeoutMs` passed)
 * resolves to false, so the caller can fall back to the in-app preview.
 * The default budget is short (IDE Tracker list views); the Tracker passes a
 * realistic one (AC-241).
 */
async function tryActivateCursorTab(title, timeoutMs = CURSOR_ACTIVATE_TIMEOUT_MS) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let res;
    try {
      res = await fetch(CURSOR_ACTIVATE_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timer);
    }
    if (!res.ok) return false;
    const data = await res.json();
    // "window_only": Cursor was restored to the front but its tab tree is still rebuilding;
    // the app is up, so do NOT also open the in-app thread view.
    return data.status === "activated" || data.status === "window_only";
  } catch {
    return false;
  }
}

// SPEC-DELTA-2026-10-07-browser-tab-copilot-mirror.md: same local on-demand
// activation shape as tryActivateCursorTab, pointed at the Browser (Copilot/
// Cowork) activation endpoint.
const BROWSER_ACTIVATE_URL = "http://127.0.0.1:4127/api/browser-tabs/activate";

async function tryActivateBrowserTab(title, timeoutMs = CURSOR_ACTIVATE_TIMEOUT_MS, target = null) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let res;
    try {
      res = await fetch(BROWSER_ACTIVATE_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(target ? { title, ...target } : { title }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timer);
    }
    if (!res.ok) return false;
    const data = await res.json();
    return data.status === "activated";
  } catch {
    return false;
  }
}

// -----------------------------------------------------------------------------
// Window pin (SPEC-DELTA-2026-10-07-cockpit-pin-density-unified-view.md).
// Same local-only, silent-off-the-laptop shape as tryActivateCursorTab/
// tryActivateBrowserTab.
// -----------------------------------------------------------------------------

const WINDOW_PIN_URL = "http://127.0.0.1:4127/api/window/pin";
// Measured 2026-10-08: a pin round trip is ~0.9-1.1 s (the daemon spawns
// powershell.exe across WSL interop each time). Reusing the 500 ms Cursor
// activation budget aborted EVERY request before the answer came back -- the
// window still toggled, but the button never updated and looked dead.
const WINDOW_PIN_TIMEOUT_MS = 5000;

async function tryToggleWindowPin(desired) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), WINDOW_PIN_TIMEOUT_MS);
    let res;
    try {
      res = await fetch(WINDOW_PIN_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ desired }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timer);
    }
    if (!res.ok) return null;
    return res.json();
  } catch {
    return null;
  }
}

// AC-237/AC-238: every .cockpit-pin-btn (header + Tracker topbar) shows ONE
// shared state, and that state is always the daemon's answer about the REAL
// window (set-window-topmost.ps1), never a local guess -- so a change made in
// any view, any open cockpit window, or outside the PWA shows up everywhere.
const PIN_TITLES = {
  on: "Always on top. Click to unpin.",
  off: "Not on top. Click to keep this window always on top.",
  unknown: "Click to keep this window always on top.",
  unavailable: "Pin unavailable: the cockpit's local service isn't answering (daemon not running, or not on this laptop). Click to retry.",
  nowindow: "Pin works in the standalone cockpit window (desktop shortcut), not in a browser tab.",
};
let pinChannel = null;
// Latest-request-wins: clicking the pin in an unfocused window fires a focus
// refresh ("query") and the click ("toggle") together; whichever answer
// arrives last must not overwrite a newer request's result.
let pinRequestSeq = 0;

function applyPinState(state, titleKey = state) {
  document.querySelectorAll(".cockpit-pin-btn").forEach((btn) => {
    btn.dataset.pinState = state;
    btn.setAttribute("aria-pressed", state === "on" ? "true" : "false");
    btn.title = PIN_TITLES[titleKey] || PIN_TITLES.unknown;
  });
}

function syncPinButton(pinned) {
  applyPinState(pinned ? "on" : "off");
}

function requestPinState(desired, { broadcast = false } = {}) {
  const seq = ++pinRequestSeq;
  return tryToggleWindowPin(desired).then((result) => {
    if (seq !== pinRequestSeq) return;
    if (!result) {
      applyPinState("unavailable");
      return;
    }
    if (result.status !== "ok") {
      applyPinState("unavailable", "nowindow");
      return;
    }
    syncPinButton(result.pinned);
    if (broadcast && pinChannel) pinChannel.postMessage({ pinned: !!result.pinned });
  });
}

function wirePinButtons() {
  if (typeof BroadcastChannel === "function") {
    pinChannel = new BroadcastChannel("cockpit-window-pin");
    pinChannel.onmessage = (ev) => {
      if (ev.data && typeof ev.data.pinned === "boolean") syncPinButton(ev.data.pinned);
    };
  }
  // A toggle takes ~1 s; a second click inside that window would toggle the
  // window straight back. Ignore clicks while one is in flight, and show it.
  let pinToggleInFlight = false;
  const setPinBusy = (busy) => {
    document.querySelectorAll(".cockpit-pin-btn").forEach((btn) => {
      if (busy) btn.dataset.pinBusy = "true";
      else delete btn.dataset.pinBusy;
    });
  };
  document.querySelectorAll(".cockpit-pin-btn").forEach((btnPin) => {
    btnPin.addEventListener("click", () => {
      if (pinToggleInFlight) return;
      pinToggleInFlight = true;
      setPinBusy(true);
      requestPinState("toggle", { broadcast: true }).finally(() => {
        pinToggleInFlight = false;
        setPinBusy(false);
      });
    });
  });
  window.addEventListener("focus", () => requestPinState("query"));
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) requestPinState("query");
  });
  // AC-214: explicit "true", not a toggle -- a page reload while already
  // pinned must never un-pin it.
  requestPinState("true", { broadcast: true });
}

// -----------------------------------------------------------------------------
// Tracker view (SPEC-DELTA-2026-10-07-cockpit-pin-density-unified-view.md):
// a dedicated, chrome-free view with THREE SEPARATE groups (Viktor asked for
// grouped, not one interleaved/sorted list) -- Claude Code (sidebar "Active"
// group only), Cursor (ide-mirror GUI tabs + non-archived v2/CLI sessions),
// Copilot (every browser-tab-mirror entry, unfiltered).
// -----------------------------------------------------------------------------

function showTrackerError(message) {
  const el = document.getElementById("tracker-error-state");
  if (el) {
    el.hidden = false;
    el.textContent = message;
  }
}

/**
 * v2 session `status` -> the same {agent, user, none} vocabulary the IDE
 * mirrors' `waitingOn` uses, PLUS a genuine "problem" kind for `"failed"` --
 * a real signal v2 sessions carry (unlike the IDE mirrors, which have no
 * error/failure field at all, so they never get a fabricated 4th state).
 */
function v2StatusToTrackerStatus(status) {
  if (status === "running") return { statusKind: "agent", statusLabel: "running" };
  if (status === "pending") return { statusKind: "user", statusLabel: "waiting on you" };
  if (status === "failed") return { statusKind: "problem", statusLabel: "problem" };
  if (status === "done" || status === "stopped") return { statusKind: "none", statusLabel: "done" };
  return { statusKind: null, statusLabel: status || "" };
}

// Tracker hover card: the shared summary popover, anchored to the whole row
// (Viktor asked for detail on hovering the ROW, not via a separate button).
// data-context lets style.css give it multi-line + click-through behavior
// in the Tracker only, leaving the IDE-tabs list's popover untouched.
function showTrackerRowPopover(anchorEl, text) {
  showIdeRowSummaryPopover(anchorEl, text);
  const popover = document.getElementById("ide-row-summary-popover");
  if (popover) popover.dataset.context = "tracker";
}

function hideTrackerRowPopover() {
  const popover = document.getElementById("ide-row-summary-popover");
  if (popover) delete popover.dataset.context;
  hideIdeRowSummaryPopover();
}

// Claude app sidebar groups the Tracker shows, in this order, followed by
// Routines (Viktor, 2026-10-08: "csak az Active group es a Routines group
// alattiakat"). Ungrouped and any other group (e.g. Dependent) stay hidden.
const TRACKER_CLAUDE_GROUPS = ["Active"];

// Same limit as the mirror's RUNNING_STALE_AFTER_MS (claude-code-mirror/lib/
// tracker-model.mjs). The mirror applies it when it publishes, but a stopped
// daemon (WSL shut down mid-turn happens on this machine) leaves the last
// "running" frozen in the file -- so the page applies it again at render time.
const TRACKER_RUNNING_STALE_MS = 30 * 60 * 1000;

function claudeTrackerRow(r) {
  const staleRunning =
    r.statusKind === "agent" && r.lastActivityAt && Date.now() - Date.parse(r.lastActivityAt) > TRACKER_RUNNING_STALE_MS;
  return {
    title: r.title,
    sessionId: r.sessionId || null,
    lastActivityAt: r.lastActivityAt,
    link: r.link || null,
    statusKind: staleRunning ? "none" : r.statusKind || null,
    statusLabel: staleRunning ? "done" : r.statusLabel || null,
    summary: r.summary || null,
    activity: r.activity || null,
    needsAction: r.needsAction || null,
  };
}

// `subgroups` ([{label, rows}]) renders labeled sub-lists under one source
// header -- the Claude Code section's sidebar groups + Routines (AC-243).
// `hint` is a small muted line under the header (e.g. why a section is empty).
function appendTrackerGroup(container, label, rows, onActivate, sourceKey, subgroups = null, hint = null) {
  const section = document.createElement("section");
  section.className = "cockpit-tracker-group";
  section.dataset.trackerSource = sourceKey;
  container.appendChild(section);

  const total = subgroups ? subgroups.reduce((n, sg) => n + sg.rows.length, 0) : rows.length;
  const title = document.createElement("h2");
  title.className = "cockpit-tracker-group-title";
  title.dataset.trackerSource = sourceKey;
  title.textContent = `${label} (${total})`;
  section.appendChild(title);

  if (hint) {
    const p = document.createElement("p");
    p.className = "cockpit-tracker-hint";
    p.textContent = hint;
    section.appendChild(p);
  }

  if (!subgroups) {
    section.appendChild(buildTrackerList(rows, onActivate));
    return;
  }
  for (const sg of subgroups) {
    const sub = document.createElement("h3");
    sub.className = "cockpit-tracker-subgroup-title";
    sub.textContent = `${sg.label} (${sg.rows.length})`;
    section.appendChild(sub);
    section.appendChild(buildTrackerList(sg.rows, onActivate));
  }
}

function buildTrackerList(rows, onActivate) {
  const ul = document.createElement("ul");
  ul.className = "cockpit-session-list";
  if (rows.length === 0) {
    const li = document.createElement("li");
    li.className = "cockpit-hint";
    li.textContent = "Nothing open.";
    ul.appendChild(li);
  }
  for (const row of rows) {
    const li = document.createElement("li");
    // cockpit-tracker-row overrides cockpit-session-row's grid (designed for
    // a fixed, always-present column set) with a flex layout where the title
    // always gets the flexible space regardless of which optional siblings
    // (dot / badge / time / info) are present on THIS row -- fixes rows
    // looking left- vs right-aligned depending on source (AC-226).
    li.className = "cockpit-session-row cockpit-tracker-row";
    li.tabIndex = 0;
    // A jump can take ~1-3 s (AC-241): mark the row busy meanwhile and
    // ignore repeat clicks on it, so one click is exactly one jump.
    const activate = () => {
      if (li.dataset.activating) return;
      const pending = onActivate(row);
      if (pending && typeof pending.finally === "function") {
        li.dataset.activating = "true";
        pending.finally(() => {
          delete li.dataset.activating;
        });
      }
    };
    li.addEventListener("click", activate);
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        activate();
      }
    });

    if (row.statusKind) {
      const dot = document.createElement("span");
      dot.className = "cockpit-row-status-dot";
      dot.dataset.status = row.statusKind;
      dot.title = row.statusLabel || "";
      li.appendChild(dot);
    }

    const titleEl = document.createElement("span");
    titleEl.className = "cockpit-row-title";
    titleEl.textContent = IDE_HELPERS.formatTabTitle(row.title, 50);
    li.appendChild(titleEl);

    // AC-227: Copilot vs Cowork in the row text itself -- they're otherwise
    // indistinguishable (see ide-helpers.mjs#classifyCopilotKind).
    if (row.sourceBadge) {
      const badge = document.createElement("span");
      badge.className = "cockpit-row-source-badge";
      badge.dataset.kind = row.sourceBadge.toLowerCase();
      badge.textContent = row.sourceBadge;
      li.appendChild(badge);
    }
    // Where this tab lives ("Chrome · tab 9", "2/3"): rows can share a title.
    if (row.where) {
      const where = document.createElement("span");
      where.className = "cockpit-row-where";
      where.textContent = row.where;
      li.appendChild(where);
    }

    if (row.lastActivityAt) {
      const time = document.createElement("time");
      time.className = "cockpit-row-time";
      time.dateTime = row.lastActivityAt;
      time.textContent = IDE_HELPERS.relativeIdeTime(row.lastActivityAt, Date.now());
      li.appendChild(time);
    }

    // AC-234: hovering the row itself shows its detail card. No separate info
    // button -- its 32px touch-target min-height was what kept every
    // summary-carrying row about twice as tall as its text.
    const hoverText = IDE_HELPERS.trackerHoverText(row, Date.now());
    li.addEventListener("mouseenter", () => showTrackerRowPopover(li, hoverText));
    li.addEventListener("mouseleave", hideTrackerRowPopover);

    ul.appendChild(li);
  }
  return ul;
}

async function renderTrackerView() {
  const groupsEl = document.getElementById("tracker-groups");
  const empty = document.getElementById("tracker-empty-state");
  const errorEl = document.getElementById("tracker-error-state");
  const summaryEl = document.getElementById("tracker-summary");
  if (!groupsEl || !empty || !errorEl || !IDE_HELPERS) return;
  errorEl.hidden = true;
  errorEl.textContent = "";

  let claudeSnap, cursorSnap, browserSnap, v2IndexResult, digestsDoc;
  try {
    [claudeSnap, cursorSnap, browserSnap, v2IndexResult, digestsDoc] = await Promise.all([
      fetchSnapshotByConfigKey("claudeCodeTabs"),
      fetchSnapshotByConfigKey("ideTabs"),
      fetchSnapshotByConfigKey("browserTabs"),
      loadV2Index().catch(() => ({ index: { sessions: [] } })),
      // Layer B (optional LLM digest, SPEC-DELTA-2026-10-08-tracker-hover-
      // digest.md): absent / failing / not configured must never affect the
      // Tracker -- the hover card then shows the deterministic text.
      fetchSnapshotByConfigKey("digests").catch(() => null),
    ]);
  } catch (err) {
    showTrackerError(err.message);
    return;
  }

  // AC-243 (2026-10-08, supersedes AC-225 at Viktor's ask): the Claude app's
  // own sidebar -- only the groups in TRACKER_CLAUDE_GROUPS, then Routines --
  // from the mirror's claudeTracker block, so moving a session to ungrouped
  // in Claude makes it leave the Tracker. When the block is present it is
  // authoritative -- even empty, even with an error: falling back to the
  // recency list would show exactly the ungrouped sessions Viktor moved out.
  // That list is only for a daemon too old to publish the block at all.
  const claudeTracker = claudeSnap.claudeTracker;
  const hasClaudeTracker = !!claudeTracker && Array.isArray(claudeTracker.groups) && Array.isArray(claudeTracker.routines);
  const claudeSubgroups = hasClaudeTracker
    ? [
        // Routines first, as in the Claude app's own sidebar (Viktor, 2026-10-08).
        { label: "Routines", rows: (claudeTracker.routines || []).map(claudeTrackerRow) },
        ...TRACKER_CLAUDE_GROUPS.map((name) => ({
          label: name,
          rows: (claudeTracker.groups.find((g) => g.name === name)?.sessions || []).map(claudeTrackerRow),
        })),
      ]
    : null;
  const claudeRows = hasClaudeTracker
    ? []
    : (claudeSnap.openTabs || []).map((t) => ({
        title: t.title,
        lastActivityAt: t.lastActivityAt,
        composerId: t.composerId,
        link: t.link || null,
        statusKind: IDE_HELPERS.isEmptyIdeTab(t) ? "none" : t.waitingOn || "none",
        statusLabel: IDE_HELPERS.ideTabStatusLabel(t),
        summary: t.summary || null,
        activity: t.activity || null,
      }));
  const claudeCount = claudeSubgroups ? claudeSubgroups.reduce((n, sg) => n + sg.rows.length, 0) : claudeRows.length;
  // Only worth saying when there is nothing to show; with last-good data the
  // mirror's error is transient and the rows are right.
  const claudeHint = hasClaudeTracker && claudeTracker.error && claudeCount === 0 ? `Claude data unavailable: ${claudeTracker.error}` : null;

  // AC-219: ide-mirror's own open GUI tabs...
  const cursorGuiRows = (cursorSnap.openTabs || []).map((t) => ({
    title: t.title,
    lastActivityAt: t.lastActivityAt,
    composerId: t.composerId,
    statusKind: IDE_HELPERS.isEmptyIdeTab(t) ? "none" : t.waitingOn || "none",
    statusLabel: IDE_HELPERS.ideTabStatusLabel(t),
    summary: t.summary || null,
    activity: t.activity || null,
    kind: "gui",
  }));
  // ...plus this cockpit's own non-archived v2 (CLI) sessions. `title` is
  // already the index-mirrored derived/custom title (see lib/transcript-
  // model.mjs#buildIndexEntry) -- no need to re-derive it here.
  const v2Sessions = (v2IndexResult.index && v2IndexResult.index.sessions) || [];
  const cursorCliRows = v2Sessions
    .filter((s) => !s.archived)
    .map((s) => {
      const { statusKind, statusLabel } = v2StatusToTrackerStatus(s.status);
      return {
        title: s.title || s.id,
        lastActivityAt: s.updatedAt,
        sessionId: s.id,
        statusKind,
        statusLabel,
        kind: "cli",
      };
    });
  const cursorRows = [...cursorGuiRows, ...cursorCliRows];

  // AC-220: every browser-tab-mirror entry, unfiltered. No status signal
  // exists for these (no agent/running concept for a browser tab) -- no
  // dot, not a fabricated one. sourceBadge distinguishes Copilot vs Cowork
  // in the row text itself (Viktor's own ask -- they look identical
  // otherwise), best-guess from the captured url, see
  // ide-helpers.mjs#classifyCopilotKind.
  const browserRows = (Array.isArray(browserSnap.tabs) ? browserSnap.tabs : []).map((t) => ({
    title: t.title,
    link: t.url || null,
    // CDP-sourced tabs (copilot-watch Edge, SPEC-DELTA-2026-10-08) carry a
    // real status + the last request; UIA-only tabs have neither.
    statusKind: t.statusKind || null,
    statusLabel: t.statusLabel || null,
    summary: t.request ? `Request: ${t.request}` : null,
    activity: t.activity || null,
    sourceBadge: IDE_HELPERS.classifyCopilotKind(t.url) === "cowork" ? "Cowork" : "Copilot",
    tracked: t.source === "cdp",
    // Precise identity of THIS tab (SPEC-DELTA-2026-10-08-tracker-activation-fixes.md): a CDP page
    // id, or the plain browser's window handle + tab position -- identical titles are common.
    targetId: t.targetId || null,
    windowHandle: t.windowHandle ?? null,
    tabIndex: t.tabIndex ?? null,
    where: t.windowHandle != null && t.tabIndex != null ? `${t.browser || "Browser"} · tab ${t.tabIndex}` : null,
  }));
  // Rows with the very same title and no position (e.g. two CDP "New chat" tabs) get "n/m".
  const sameTitle = new Map();
  for (const r of browserRows) if (!r.where) sameTitle.set(r.title, [...(sameTitle.get(r.title) || []), r]);
  for (const group of sameTitle.values()) {
    if (group.length > 1) group.forEach((r, i) => { r.where = `${i + 1}/${group.length}`; });
  }
  // Tabs in a plain Chrome/Edge are title-only (no status, no request): they
  // go to their own group at the very bottom (Viktor, 2026-10-08).
  const trackedBrowserRows = browserRows.filter((r) => r.tracked);
  const untrackedBrowserRows = browserRows.filter((r) => !r.tracked);

  // Layer B: join each row's stored LLM digest (shown only while it still
  // matches the row's current activity -- pickDigest drops stale ones).
  const attachDigests = (source, rows) => {
    for (const r of rows) r.digest = IDE_HELPERS.pickDigest(digestsDoc, IDE_HELPERS.digestRowKey(source, r), r.activity);
  };
  attachDigests("claude", claudeRows);
  for (const sg of claudeSubgroups || []) attachDigests("claude", sg.rows);
  attachDigests("cursor", cursorGuiRows);
  attachDigests("browser", browserRows);

  groupsEl.innerHTML = "";
  const totalRows = claudeCount + cursorRows.length + browserRows.length;
  if (summaryEl) summaryEl.textContent = `${totalRows} session${totalRows === 1 ? "" : "s"}`;

  if (totalRows === 0) {
    empty.hidden = false;
    groupsEl.hidden = true;
    return;
  }
  empty.hidden = true;
  groupsEl.hidden = false;

  appendTrackerGroup(groupsEl, "Claude Code", claudeRows, (row) => {
    if (row.link) window.open(row.link, "_blank", "noopener");
  }, "claude", claudeSubgroups, claudeHint);
  // Open GUI tabs first, then this cockpit's CLI sessions (Viktor, 2026-10-08).
  const cursorSubgroups = [
    { label: "GUI tabs", rows: cursorGuiRows },
    { label: "CLI sessions", rows: cursorCliRows },
  ];
  appendTrackerGroup(groupsEl, "Cursor", cursorRows, (row) => {
    if (row.kind === "cli") {
      setView("v2-detail", { sessionId: row.sessionId });
      return;
    }
    // AC-222: on a failed/unavailable local activation, point the shared
    // cachedIdeSnapshot/ideTrackerSource globals at THIS render's own
    // cursorSnap before opening the in-app detail view -- otherwise it could
    // show stale or wrong-source data left over from whatever the Cursor/
    // Claude Code/Browser sub-views last loaded.
    // AC-241: wait for the daemon's real answer (~2.7 s for Cursor) before
    // falling back -- see TRACKER_CURSOR_ACTIVATE_TIMEOUT_MS.
    return tryActivateCursorTab(row.title, TRACKER_CURSOR_ACTIVATE_TIMEOUT_MS).then((activated) => {
      if (activated) return;
      ideTrackerSource = "cursor";
      cachedIdeSnapshot = cursorSnap;
      setView("ide-tab-detail", { composerId: row.composerId });
    });
  }, "cursor", cursorSubgroups);
  const activateBrowserRow = (row) => {
    // AC-241: only open the url when the existing tab really couldn't be
    // activated -- falling back early is what opened duplicate Copilot tabs.
    const target = row.tracked
      ? { url: row.link, targetId: row.targetId }
      : row.windowHandle != null && row.tabIndex != null
        ? { windowHandle: row.windowHandle, tabIndex: row.tabIndex }
        : null;
    return tryActivateBrowserTab(row.title, TRACKER_BROWSER_ACTIVATE_TIMEOUT_MS, target).then((activated) => {
      if (!activated && row.link) window.open(row.link, "_blank", "noopener");
    });
  };
  // One section, two named sub-groups (same shape as Claude Code's Routines/Active).
  const copilotSubgroups = [{ label: "Tracked (Playwright Edge)", rows: trackedBrowserRows }];
  if (untrackedBrowserRows.length > 0) {
    copilotSubgroups.push({ label: "Not trackable (plain Chrome/Edge)", rows: untrackedBrowserRows });
  }
  appendTrackerGroup(groupsEl, "Copilot / Cowork", [], activateBrowserRow, "copilot", copilotSubgroups);
}

// -----------------------------------------------------------------------------
// IDE Tracker row summary popover (SPEC-DELTA-2026-10-06-ide-tracker-summary-
// and-open-link.md, AC-197). One shared, fixed-size element, repositioned next
// to whichever row's info button triggered it -- same "single shared panel"
// idiom as #v2-settings-sheet / #app-menu, not one popover per row.
// -----------------------------------------------------------------------------

let ideRowSummaryPopoverAnchor = null;

function positionIdeRowSummaryPopover(popover, anchorEl) {
  const rect = anchorEl.getBoundingClientRect();
  const margin = 8;
  let left = rect.left;
  const maxLeft = window.innerWidth - popover.offsetWidth - margin;
  if (left > maxLeft) left = Math.max(margin, maxLeft);
  let top = rect.bottom + margin;
  if (top + popover.offsetHeight > window.innerHeight - margin) {
    top = Math.max(margin, rect.top - popover.offsetHeight - margin);
  }
  popover.style.left = `${left}px`;
  popover.style.top = `${top}px`;
}

function showIdeRowSummaryPopover(anchorEl, text) {
  const popover = document.getElementById("ide-row-summary-popover");
  if (!popover) return;
  popover.textContent = text;
  popover.hidden = false;
  ideRowSummaryPopoverAnchor = anchorEl;
  positionIdeRowSummaryPopover(popover, anchorEl);
}

function hideIdeRowSummaryPopover() {
  const popover = document.getElementById("ide-row-summary-popover");
  if (popover) popover.hidden = true;
  ideRowSummaryPopoverAnchor = null;
}

function toggleIdeRowSummaryPopover(anchorEl, text) {
  const popover = document.getElementById("ide-row-summary-popover");
  if (popover && !popover.hidden && ideRowSummaryPopoverAnchor === anchorEl) {
    hideIdeRowSummaryPopover();
  } else {
    showIdeRowSummaryPopover(anchorEl, text);
  }
}

/** Wires the one shared popover's dismiss-on-outside-click / Escape behavior. */
function wireIdeRowSummaryPopover() {
  const popover = document.getElementById("ide-row-summary-popover");
  if (!popover) return;
  document.addEventListener("click", (ev) => {
    if (popover.hidden) return;
    if (popover.contains(ev.target)) return;
    if (ev.target.closest && ev.target.closest(".cockpit-row-info")) return;
    hideIdeRowSummaryPopover();
  });
  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape" && !popover.hidden) hideIdeRowSummaryPopover();
  });
}

/**
 * SPEC-DELTA-2026-10-07-browser-tab-copilot-mirror.md: a much simpler flat
 * list than renderIdeTabsList's Cursor/Claude Code rendering -- there is no
 * transcript, no waitingOn, no Open/Archive split, just a title and a
 * best-effort url. Reuses the same #ide-tabs-list / cockpit-session-row DOM.
 */
async function renderBrowserTabsList() {
  const ul = document.getElementById("ide-tabs-list");
  const empty = document.getElementById("ide-tabs-empty-state");
  const summary = document.getElementById("ide-summary");
  if (!ul || !empty || !summary) return;

  clearIdeTabsError();
  let snapshot;
  try {
    snapshot = await loadIdeTabs();
  } catch (err) {
    showIdeTabsError(err.message);
    return;
  }

  const tabs = Array.isArray(snapshot.tabs) ? snapshot.tabs : [];

  const workspaceLabel = document.getElementById("ide-tabs-workspace-label");
  if (workspaceLabel) workspaceLabel.textContent = "";

  summary.innerHTML = "";
  const total = document.createElement("span");
  total.className = "cockpit-ide-summary-bucket";
  total.innerHTML = `<strong>${tabs.length}</strong> ${tabs.length === 1 ? "tab" : "tabs"}`;
  summary.appendChild(total);
  if (snapshot.snapshotAt && IDE_HELPERS) {
    const ts = document.createElement("span");
    ts.className = "cockpit-ide-summary-bucket";
    ts.textContent = `mirrored ${IDE_HELPERS.relativeIdeTime(snapshot.snapshotAt, Date.now())}`;
    summary.appendChild(ts);
  }

  ul.innerHTML = "";
  if (tabs.length === 0) {
    empty.hidden = false;
    empty.textContent = "No Copilot/Cowork browser tabs detected yet. Open one, then refresh.";
    return;
  }
  empty.hidden = true;

  for (const t of tabs) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.tabIndex = 0;
    const hasUrl = typeof t.url === "string" && t.url.length > 0;
    // Try local on-demand activation first (same two-tier fallback shape as
    // Cursor rows); if that doesn't succeed, fall back to the known url
    // when there is one (AC-211). No url and no activation -> inert row,
    // never a false "opened" signal.
    const activateRow = async () => {
      const activated = await tryActivateBrowserTab(t.title);
      if (!activated && hasUrl) window.open(t.url, "_blank", "noopener");
    };
    li.addEventListener("click", activateRow);
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        activateRow();
      }
    });

    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = IDE_HELPERS ? IDE_HELPERS.formatTabTitle(t.title, 60) : t.title;
    li.appendChild(title);

    const status = document.createElement("span");
    status.className = "cockpit-row-status";
    status.textContent = hasUrl ? "link known" : "no link yet";
    li.appendChild(status);

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
  // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-160): this tab's Share panel.
  refreshSharesIfStale()
    .then(() => renderSharePanel(currentShareTarget("ide")))
    .catch(() => renderSharePanel(currentShareTarget("ide")));
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
    handleV2ArchiveToggle(s.id, !archived, s.status).catch((err) => showV2ListError(err.message));
  });
  li.appendChild(archiveBtn);

  // SPEC-DELTA-2026-09-27-queue-remove-and-session-delete: reachable straight
  // from the list, same directness as Archive above -- Viktor's whole
  // complaint was list clutter, so cleanup can't be buried a tap deeper.
  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "cockpit-btn cockpit-btn-danger cockpit-row-archive-btn";
  deleteBtn.textContent = "Delete";
  deleteBtn.setAttribute("aria-label", `Delete ${s.title || "chat"} permanently`);
  deleteBtn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    handleV2ListDeleteClick(s.id, s.title).catch((err) => showV2ListError(err.message));
  });
  li.appendChild(deleteBtn);

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

/**
 * Per-row Archive/Unarchive control handler (AC-047). AC-029: archiving a
 * RUNNING session requires an explicit confirmation first -- the backend
 * (`archiveSession`) already stops it (`pendingAction: "stop"`) once
 * confirmed, this is just the "ask first" half. Unarchiving, and archiving a
 * non-running session, need no confirmation (matches AC-050's "only guard
 * when it matters" precedent).
 */
async function handleV2ArchiveToggle(id, archived, status) {
  if (archived && status === "running") {
    const proceed = window.confirm(
      "This chat is currently running. Archiving it will stop the agent first. Continue?",
    );
    if (!proceed) return;
  }
  clearV2ListError();
  try {
    await v2SetArchived(id, archived);
  } catch (err) {
    showV2ListError(err.message);
    return;
  }
  await renderV2List().catch((err) => showV2ListError(err.message));
}

/**
 * Per-row Delete control handler (AC-030). Unlike Archive, this is
 * IRREVERSIBLE, so it always confirms first -- same UX approach as the
 * daemon-control-watchdog delta's confirm-before-stop dialog, adapted
 * wording (see SPEC-DELTA-2026-09-27-queue-remove-and-session-delete).
 */
async function handleV2ListDeleteClick(id, titleText) {
  const proceed = window.confirm(
    `Delete "${titleText || "(untitled)"}" permanently? This cannot be undone.`,
  );
  if (!proceed) return;
  clearV2ListError();
  try {
    await v2DeleteSession(id);
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
  ];
  for (const id of ids) {
    const el = document.getElementById(id);
    // Never re-enable over a read-only chat (a guest holds Control, or the
    // lease is elsewhere) -- SPEC-DELTA-2026-09-29-session-sharing-stage2.
    if (el) el.disabled = busy || v2HostReadOnly;
  }
  // Composer buttons are re-derived, never blanket-enabled (AC-094): an
  // empty box must keep Send disabled after the action finishes.
  v2Busy = busy;
  applyComposerButtons();
}

/**
 * Show/hide/enable Stop, Force and Send/Queue from the current status, box
 * content and busy flag (SPEC-DELTA-2026-09-29-composer-contextual-buttons).
 */
function applyComposerButtons() {
  if (!COMPOSER_STATE) return;
  const textEl = document.getElementById("v2-composer-text");
  const b = COMPOSER_STATE.deriveComposerButtons({
    status: v2ComposerStatus,
    text: textEl ? textEl.value : "",
    busy: v2Busy,
    readOnly: v2HostReadOnly,
  });
  const btnSend = document.getElementById("btn-v2-composer-send");
  if (btnSend) {
    btnSend.disabled = !b.send.enabled;
    btnSend.dataset.mode = b.send.mode;
    btnSend.title = b.send.title;
    btnSend.setAttribute("aria-label", b.send.label);
  }
  const btnStop = document.getElementById("btn-v2-stop");
  if (btnStop) {
    btnStop.hidden = !b.stop.visible;
    btnStop.disabled = !b.stop.enabled;
  }
  const btnForce = document.getElementById("btn-v2-composer-force");
  if (btnForce) {
    btnForce.hidden = !b.force.visible;
    btnForce.disabled = !b.force.enabled;
  }
}

/** Auto-grow the composer box from 1 row up to ~5 rows. */
function autoGrowComposer() {
  const textEl = document.getElementById("v2-composer-text");
  if (!textEl || !COMPOSER_STATE) return;
  textEl.style.height = "auto";
  // scrollHeight excludes the border; box-sizing is border-box, so add it
  // back or a 3-line message already shows a scrollbar.
  const border = textEl.offsetHeight - textEl.clientHeight;
  const px = COMPOSER_STATE.clampComposerHeight(textEl.scrollHeight + border, { minPx: 40, maxPx: 140 });
  textEl.style.height = `${px}px`;
}

/** Status chip next to the chat title (AC-095). */
function applyStatusChip(record) {
  const chip = document.getElementById("v2-detail-status-chip");
  if (!chip) return;
  const c = COMPOSER_STATE ? COMPOSER_STATE.deriveStatusChip(record) : null;
  chip.hidden = !c;
  if (c) {
    chip.textContent = c.label;
    chip.dataset.tone = c.tone;
  }
}

/** Apply a freshly read record to the composer + chip. */
function syncComposerFromRecord(record) {
  v2ComposerStatus = record.status;
  const btnStop = document.getElementById("btn-v2-stop");
  if (btnStop) {
    btnStop.onclick = record.status === "running" ? () => handleV2StopClick(record.id) : null;
  }
  applyComposerButtons();
  applyStatusChip(record);
}

async function renderV2Detail(sessionId) {
  clearV2DetailError();
  // Always start collapsed when (re-)entering a session's detail view --
  // it staying open from a PREVIOUS session would be confusing, and the
  // poll tick below never touches this itself.
  closeV2SettingsSheet();
  closeChatSwitcher();
  activeV2DetailSessionId = sessionId;
  // (Re-)entering a chat lands on the newest message (AC-102).
  v2ScrollState = { sessionId: null, messageCount: 0 };
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
  syncStaleModelOption(modelInput, record.model);
  if (modelInput) {
    modelInput.value = record.model || "auto";
    modelInput.dataset.priorValue = modelInput.value;
  }
  const modeSelect = document.getElementById("v2-detail-mode-select");
  if (modeSelect) modeSelect.value = record.mode || "agent";

  syncComposerFromRecord(record);
  autoGrowComposer();
  renderV2Messages(record);
  // SPEC-DELTA-2026-09-29-session-sharing-stage2: Share panel + host
  // read-only state (AC-157/AC-163). AC-042's settings-gear dot now means
  // "this chat has an active guest" (applyHostAccess).
  applyHostAccess(record);
  sharesLoadedAt = 0;
  refreshSharesIfStale()
    .catch(() => null)
    .then(() => {
      if (activeV2DetailSessionId !== record.id) return;
      renderSharePanel(currentShareTarget("chat"));
      applyHostAccess(activeV2DetailRecord || record);
    });

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

/**
 * SPEC-DELTA-2026-09-29-session-sharing-stage2: renders one Share panel (the
 * chat's in the settings sheet, or an IDE tab's) from the cached registry and
 * relay status. One row per guest: name + connection, Off | Read | Control |
 * Concurrent (tabs: Off | Read), Copy link, remove. Rendering only -- never
 * throws; data comes from loadShares()/loadRelayStatus().
 */
function renderSharePanel(target) {
  if (!target || !SHARE_UI) return;
  const isIde = target.panel === "ide";
  const list = document.getElementById(isIde ? "ide-detail-shared-list" : "v2-detail-shared-list");
  const empty = document.getElementById(isIde ? "ide-share-empty-state" : "v2-share-empty-state");
  if (!list) return;
  const rows = SHARE_UI.deriveShareRows({ shares: cachedShares, kind: target.kind, id: target.id, relayStatus: cachedRelayStatus });
  list.innerHTML = "";
  if (empty) empty.hidden = rows.length > 0;
  for (const row of rows) {
    const li = document.createElement("li");
    li.className = "v2-share-row";
    const who = document.createElement("div");
    who.className = "v2-share-who";
    const name = document.createElement("span");
    name.className = "v2-share-name";
    name.textContent = row.name;
    name.title = row.email;
    who.appendChild(name);
    if (row.connectionLabel) {
      // SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal
      // (AC-186/AC-187): an ERR state is tappable and opens the same
      // diagnostics sheet the VT-menu entry point uses -- every ERR
      // anywhere in the app opens one shared view, not a one-off.
      const conn = document.createElement(row.connection === "error" ? "button" : "span");
      conn.className = "v2-share-conn";
      conn.dataset.state = row.connection;
      conn.textContent = row.connectionLabel;
      if (row.connection === "error") {
        conn.type = "button";
        conn.addEventListener("click", () => openDiagnosticsSheet());
      }
      who.appendChild(conn);
    }
    li.appendChild(who);
    const seg = document.createElement("div");
    seg.className = "v2-share-modes";
    seg.setAttribute("role", "radiogroup");
    seg.setAttribute("aria-label", `Access for ${row.name}`);
    for (const m of row.modes) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "v2-share-mode-btn";
      b.textContent = m.label;
      b.dataset.mode = m.id;
      b.setAttribute("role", "radio");
      b.setAttribute("aria-checked", String(m.selected));
      if (!m.selected) b.addEventListener("click", () => handleShareModeClick(target.panel, row.email, m.id));
      seg.appendChild(b);
    }
    li.appendChild(seg);
    const actions = document.createElement("div");
    actions.className = "v2-share-actions";
    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "cockpit-btn cockpit-btn-small";
    copy.textContent = "Copy link";
    copy.addEventListener("click", () => handleShareCopyLinkClick(target.panel));
    actions.appendChild(copy);
    const del = document.createElement("button");
    del.type = "button";
    del.className = "cockpit-btn cockpit-btn-icon v2-share-remove";
    del.textContent = "🗑";
    del.setAttribute("aria-label", `Remove ${row.email}`);
    del.addEventListener("click", () => handleShareRemoveClick(target.panel, row.email));
    actions.appendChild(del);
    li.appendChild(actions);
    list.appendChild(li);
  }
  const stopAll = document.getElementById("btn-v2-share-stop-all");
  if (stopAll) {
    const st = SHARE_UI.deriveStopAllState(cachedShares);
    stopAll.textContent = st.label;
    stopAll.dataset.stopped = String(st.stopped);
  }
}

/**
 * Host read-only state for the open chat (AC-157/AC-163): a guest holds
 * Control, or another device holds the lease. Drives the access banner, the
 * take-over row, and -- via v2HostReadOnly -- deriveComposerButtons.
 */
let v2HostReadOnly = false;
function applyHostAccess(record) {
  if (!SHARE_UI) return;
  const access = SHARE_UI.deriveHostChatAccess({ shares: cachedShares, record });
  v2HostReadOnly = access.readOnly;
  const banner = document.getElementById("v2-access-banner");
  const text = document.getElementById("v2-access-banner-text");
  const action = document.getElementById("btn-v2-access-action");
  if (banner) banner.hidden = !access.readOnly;
  if (text) text.textContent = access.banner || "";
  if (action) {
    action.hidden = !access.action;
    action.textContent = access.action ? access.action.label : "";
    action.dataset.action = access.action ? access.action.id : "";
  }
  for (const id of ["v2-detail-model-input", "v2-detail-mode-select", "v2-composer-text"]) {
    const el = document.getElementById(id);
    if (el) el.disabled = v2Busy || access.readOnly;
  }
  const row = SHARE_UI.deriveTakeOverRow(record);
  const rowText = document.getElementById("v2-takeover-text");
  const rowBtn = document.getElementById("btn-v2-take-over");
  if (rowText) rowText.textContent = row.text;
  if (rowBtn) rowBtn.hidden = !row.action;
  const settingsDot = document.getElementById("v2-settings-dot");
  if (settingsDot) {
    const item = cachedShares && SHARE_MODEL ? SHARE_MODEL.findItem(cachedShares, "session", record.id) : null;
    settingsDot.hidden = !(item && cachedShares.sharingEnabled !== false && item.guests.some((g) => g.mode !== "off"));
  }
  applyComposerButtons();
}

function renderV2Messages(record) {
  const container = document.getElementById("v2-messages");
  if (!container) return;
  // Measure BEFORE the rebuild: clearing innerHTML resets a scroll
  // container to the top (SPEC-DELTA-2026-09-29-chat-bottom-panel-and-autoscroll).
  const wasAtBottom = SCROLLBACK_HELPERS
    ? SCROLLBACK_HELPERS.shouldAutoScrollToBottom({
        scrollTop: container.scrollTop,
        scrollHeight: container.scrollHeight,
        clientHeight: container.clientHeight,
      })
    : true;
  const prevScrollTop = container.scrollTop;
  container.innerHTML = "";
  try {
    renderV2MessageList(container, record);
  } finally {
    applyV2ScrollAfterRender(container, record, wasAtBottom, prevScrollTop);
  }
}

/** Per-session scroll bookkeeping for applyV2ScrollAfterRender. */
let v2ScrollState = { sessionId: null, messageCount: 0 };
/** Set by the user's own Send/Queue/Force so the next render follows the bottom (AC-106). */
let v2ForceScrollBottom = false;

function applyV2ScrollAfterRender(container, record, wasAtBottom, prevScrollTop) {
  if (!SCROLLBACK_HELPERS) return;
  // Only the agent's/system's messages count as "new" -- the user's own
  // message moving out of the queue must not raise the button (S-086).
  const count = SCROLLBACK_HELPERS.countIncomingMessages(record.messages);
  const isFirstRender = v2ScrollState.sessionId !== record.id;
  const newMessageArrived = !isFirstRender && count > v2ScrollState.messageCount;
  v2ScrollState = { sessionId: record.id, messageCount: count };
  const pill = document.getElementById("btn-v2-new-message");
  const d = SCROLLBACK_HELPERS.decideScrollAfterRender({
    isFirstRender,
    forceBottom: v2ForceScrollBottom,
    wasAtBottom,
    newMessageArrived,
    pillVisible: pill ? !pill.hidden : false,
  });
  v2ForceScrollBottom = false;
  if (pill) pill.hidden = !d.showNewMessagePill;
  if (d.scrollToBottom) {
    container.scrollTop = container.scrollHeight;
    // Again after layout: on a first render the view may only just have
    // become visible, with no scrollHeight yet.
    requestAnimationFrame(() => {
      container.scrollTop = container.scrollHeight;
    });
  } else {
    container.scrollTop = prevScrollTop;
  }
}

function renderV2MessageList(container, record) {
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
      const text = document.createElement("span");
      text.className = "v2-queue-item-text";
      text.textContent = item.text;
      row.appendChild(text);
      // S-006 / AC-015 "remove" -- no confirmation needed, fully reversible
      // (the user can just retype it), unlike Force/Stop/Delete.
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "cockpit-btn cockpit-btn-icon v2-queue-remove-btn";
      removeBtn.textContent = "✕";
      removeBtn.setAttribute("aria-label", "Remove queued message");
      removeBtn.addEventListener("click", () => {
        handleV2QueueRemoveClick(record.id, item.id).catch((err) => showV2DetailError(err.message));
      });
      row.appendChild(removeBtn);
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
    if (document.hidden) return; // AC-131: no polling while the page is hidden
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
    if (document.hidden) return; // AC-131: no polling while the page is hidden
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
        activeV2DetailRecord = record;
        syncComposerFromRecord(record);
        refreshSharesIfStale()
          .catch(() => null)
          .then(() => applyHostAccess(record));
        // A chat a guest can write to may change at any time, so it keeps
        // polling (at the same 5 s cadence) even when it looks idle.
        const guestCanWrite = SHARE_MODEL && cachedShares
          ? (SHARE_MODEL.findItem(cachedShares, "session", record.id)?.guests || []).some((g) => g.mode === "control" || g.mode === "concurrent")
          : false;
        const stillChanging =
          record.status === "running" ||
          !record.chatId ||
          (Array.isArray(record.queue) && record.queue.length > 0) ||
          (guestCanWrite && cachedShares.sharingEnabled !== false);
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
      if (document.hidden) return; // AC-131: no polling while the page is hidden
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
      } else if (v === "tracker") {
        // Tracker merges 3 already-auto-refreshing background mirrors; this
        // re-render is what makes the glance-dashboard itself live while
        // it's the open view, same cadence as the IDE-tabs list.
        renderTrackerView().catch((err) => showTrackerError(err.message));
      }
    }, ideIntervalMs);
  }
  // Independent timer for the v2 session list (mobile follow-along,
  // 2026-09-24) -- same cadence as v1's list poll, separate handle so
  // switching modes never starts/stops the wrong one.
  if (v2RefreshTimerId === null) {
    const v2IntervalMs = Math.max(5, (CONFIG.pwa.pollIntervalSeconds | 0)) * 1000;
    v2RefreshTimerId = setInterval(() => {
      if (document.hidden) return; // AC-131: no polling while the page is hidden
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
  // AC-276: an empty/unlisted value is never sent -- resolveModelChoice falls back to auto.
  const model = MODEL_CHOICE
    ? MODEL_CHOICE.resolveModelChoice(modelEl ? modelEl.value : "", modelOptionIds()).model
    : (modelEl ? modelEl.value.trim() : "") || "auto";
  try {
    const record = await v2CreateSession({
      cwd: cwdEl ? cwdEl.value : "",
      model,
      mode: modeEl ? modeEl.value : "",
      parentId: null,
      firstMessage: messageEl ? messageEl.value : "",
    });
    setLastUsedModel(model);
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
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
  const textEl = document.getElementById("v2-composer-text");
  const id = activeV2DetailSessionId;
  if (!id || !textEl) return;
  const text = textEl.value.trim();
  if (!text) return;
  clearV2DetailError();
  setV2Busy(true);
  v2ForceScrollBottom = true;
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
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
  const textEl = document.getElementById("v2-composer-text");
  const id = activeV2DetailSessionId;
  if (!id || !textEl) return;
  const text = textEl.value.trim();
  if (!text) return;
  clearV2DetailError();
  setV2Busy(true);
  v2ForceScrollBottom = true;
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
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
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

/** Per-queue-item remove control handler (S-006, AC-015 "remove"). */
async function handleV2QueueRemoveClick(id, queueItemId) {
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2RemoveQueuedMessage(id, queueItemId);
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
    setV2Busy(false);
  }
}

/**
 * Settings-sheet Delete control handler (AC-030) -- symmetry with the list
 * row's own Delete, for a user already inside a session. Reads
 * activeV2DetailSessionId/activeV2DetailRecord the same way
 * handleV2RenameClick does. On success, navigates back to the list (this
 * session's own detail view obviously can't be re-rendered) and refreshes it.
 */
async function handleV2DetailDeleteClick() {
  const id = activeV2DetailSessionId;
  if (!id) return;
  const title = activeV2DetailRecord ? V2_MODEL.deriveTitle(activeV2DetailRecord) : "";
  const proceed = window.confirm(`Delete "${title || "(untitled)"}" permanently? This cannot be undone.`);
  if (!proceed) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    await v2DeleteSession(id);
  } catch (err) {
    showV2DetailError(err.message);
    setV2Busy(false);
    return;
  }
  setV2Busy(false);
  setView("v2-list");
  await renderV2List().catch((err) => showV2ListError(err.message));
}


async function handleV2ModelChange() {
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
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
  if (v2HostReadOnly) return; // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-157/AC-163)
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
// SPEC-DELTA-2026-09-29-session-sharing-stage2: Share panel actions. Every
// action is one shares.json read-modify-write; share-relay/ picks the change
// up within one cycle (mirror, delete, stop ingesting).

/** Which item the currently visible Share panel belongs to. */
function currentShareTarget(panel) {
  if (panel === "ide") {
    if (!activeIdeTabComposerId) return null;
    const kind = ideTrackerSource === "claude-code" ? "claude-tab" : "cursor-tab";
    const tab = cachedIdeSnapshot && IDE_HELPERS
      ? IDE_HELPERS.findIdeTab(cachedIdeSnapshot, activeIdeTabComposerId)
      : null;
    return { kind, id: activeIdeTabComposerId, title: (tab && tab.title) || null, panel };
  }
  if (!activeV2DetailSessionId) return null;
  const title = activeV2DetailRecord && V2_MODEL ? V2_MODEL.deriveTitle(activeV2DetailRecord) : null;
  return { kind: "session", id: activeV2DetailSessionId, title, panel: "chat" };
}

function shareErrorEl(panel) {
  return document.getElementById(panel === "ide" ? "ide-share-error-state" : "v2-share-error-state");
}

async function runShareAction(panel, fn) {
  const errEl = shareErrorEl(panel);
  if (errEl) errEl.hidden = true;
  try {
    await fn();
  } catch (err) {
    if (errEl) {
      errEl.textContent = translateErrorMessage(err.message);
      errEl.hidden = false;
    }
  }
  const target = currentShareTarget(panel);
  if (target) renderSharePanel(target);
  if (panel === "chat" && activeV2DetailRecord) applyHostAccess(activeV2DetailRecord);
}

async function handleShareInviteSubmit(evt, panel) {
  evt.preventDefault();
  const target = currentShareTarget(panel);
  const emailInput = document.getElementById(panel === "ide" ? "ide-share-invite-email" : "v2-share-invite-email");
  if (!target || !emailInput) return;
  const email = emailInput.value.trim();
  if (!email) return;
  await runShareAction(panel, async () => {
    await mutateSharesPwa((s) => SHARE_MODEL.addGuest(s, { kind: target.kind, id: target.id, title: target.title, email, mode: "read", now: Date.now() }));
    emailInput.value = "";
  });
}

async function handleShareModeClick(panel, email, mode) {
  const target = currentShareTarget(panel);
  if (!target) return;
  if (mode === "control" && !window.confirm(`Give ${SHARE_UI.nameFromEmail(email)} control of this chat? Your own view becomes read-only until you take it back.`)) return;
  await runShareAction(panel, () =>
    mutateSharesPwa((s) => SHARE_MODEL.setGuestMode(s, { kind: target.kind, id: target.id, email, mode, now: Date.now() })),
  );
}

async function handleShareRemoveClick(panel, email) {
  const target = currentShareTarget(panel);
  if (!target) return;
  if (!window.confirm(`Stop sharing with ${email} and remove them from this list?`)) return;
  await runShareAction(panel, () =>
    mutateSharesPwa((s) => SHARE_MODEL.removeGuest(s, { kind: target.kind, id: target.id, email, now: Date.now() })),
  );
}

async function handleShareCopyLinkClick(panel) {
  const target = currentShareTarget(panel);
  if (!target) return;
  const link = SHARE_MODEL.buildShareLink({
    baseUrl: CONFIG.sharing.shareLinkBase,
    kind: target.kind,
    id: target.id,
    hostUpn: CONFIG.sharing.hostUpn,
  });
  try {
    await navigator.clipboard.writeText(link);
    setStatusBadge("Link copied", "ok");
  } catch {
    window.prompt("Copy this link:", link);
  }
}

async function handleShareStopAllClick() {
  const stopped = cachedShares && cachedShares.sharingEnabled === false;
  if (!stopped && !window.confirm("Stop all sharing? Every guest loses access to every shared item right away. Their modes are remembered.")) return;
  await runShareAction("chat", () =>
    mutateSharesPwa((s) => SHARE_MODEL.setSharingEnabled(s, { enabled: stopped, now: Date.now() })),
  );
}

/** "Take back control" / "Take over" from the access banner or the settings row. */
async function handleV2AccessAction(actionId) {
  const id = activeV2DetailSessionId;
  if (!id) return;
  clearV2DetailError();
  setV2Busy(true);
  try {
    if (actionId === "take-back") {
      await mutateSharesPwa((s) => SHARE_MODEL.takeBackControl(s, { kind: "session", id, now: Date.now() }));
    } else if (actionId === "take-over") {
      if (!window.confirm("Take over from the laptop? Close the mc session on the laptop first, or both will send to the same chat.")) return;
      await v2TakeOverFromDevice(id);
    }
  } catch (err) {
    showV2DetailError(err.message);
  } finally {
    setV2Busy(false);
    await renderV2Detail(id).catch((err) => showV2DetailError(err.message));
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

/** Last-loaded health.json, cached for the diagnostics panel
 *  (SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal) -- same
 *  idea as cachedRelayStatus/cachedShares below. */
let cachedHealthStatus = null;

/** Fetches the daemon-published health status and renders it. Never throws. */
async function loadHealth() {
  if (!CONFIG.health) return;
  try {
    const { json } = await loadJson(CONFIG.health.endpoint);
    cachedHealthStatus = json;
    renderHealthBadge(json);
  } catch {
    // Read-only, best-effort -- a failed fetch just leaves the badge as it
    // was (or hidden, if it never loaded), same "don't crash the rest of
    // the app over a secondary signal" posture as the IDE-tabs mirror.
    cachedHealthStatus = null;
  }
}

// =============================================================================
// 7c-2. Daemon start/stop control strip (SPEC-DELTA-2026-09-27-
//        daemon-control-watchdog)
// =============================================================================
//
// Background this exists to fix: the daemon's only prior health signal
// (health.json, above) is written BY the daemon itself, so a dead daemon
// can never report its own death -- the badge just kept showing the last
// known "ok" forever. daemon-status.json is written by a SEPARATE watchdog
// process (a Windows Scheduled Task, \Cursor\MobileCockpitControlWatchdog,
// every ~1 min) on EVERY tick regardless of whether a start/stop request
// came in, so its own absence/staleness is itself informative -- see
// DAEMON_CONTROL_MODEL#deriveControlDisplayState.
//
// Viktor's stated design (verbatim, translated): "it should warn me if
// something is still running and ask if I'm sure before stopping, then
// force-stop." The confirmation dialog below is the ONLY gate -- once a
// stop request reaches the watchdog, it always hard-kills, no separate
// graceful-vs-force request type.

/** Renders one daemon-control status into the always-visible footer strip
 *  (never gated behind the settings sheet -- Viktor wants this visible
 *  without extra taps). Never throws. */
function renderDaemonControlBadge(status) {
  const badgeEl = document.getElementById("daemon-control-badge");
  const agoEl = document.getElementById("daemon-control-checked-ago");
  const startBtn = document.getElementById("btn-daemon-start");
  const stopBtn = document.getElementById("btn-daemon-stop");
  const repairBtn = document.getElementById("btn-daemon-repair");
  if (!badgeEl || !DAEMON_CONTROL_MODEL) return;

  const nowMs = Date.now();
  const staleAfterMs =
    ((CONFIG && CONFIG.daemonControl && CONFIG.daemonControl.staleAfterSeconds) || 180) * 1000;
  const state = DAEMON_CONTROL_MODEL.deriveControlDisplayState({ status, nowMs, staleAfterMs });

  const labels = {
    running: "daemon: running",
    stopped: "daemon: stopped",
    error: "daemon: error",
    unknown: "daemon: unknown",
  };
  badgeEl.dataset.state = state;
  badgeEl.textContent = labels[state] || `daemon: ${state}`;
  badgeEl.title =
    status && status.lastAction && status.lastAction.result === "error" && status.lastAction.error
      ? status.lastAction.error
      : "";

  if (agoEl) {
    const secs = DAEMON_CONTROL_MODEL.secondsSinceChecked(status, nowMs);
    agoEl.textContent = secs === null ? "" : `(checked ${secs}s ago)`;
  }

  // Start when we're confidently NOT running; Stop when we're confidently
  // running. "unknown"/"error" show neither -- guessing which action makes
  // sense when the watchdog itself hasn't reported anything trustworthy
  // yet would be the same silent-guess failure mode this feature exists to
  // avoid.
  if (startBtn) startBtn.hidden = state !== "stopped";
  if (stopBtn) stopBtn.hidden = state !== "running";
  // SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal: the one
  // case Start/Stop deliberately leave with no button at all -- give the
  // host something to actually try. Both "error" (a prior action failed)
  // and "unknown" (the watchdog itself hasn't reported in -- the exact
  // symptom of a wedged WSL, SPEC.md's motivating incident) are covered.
  if (repairBtn) repairBtn.hidden = state !== "error" && state !== "unknown";
}

/** Last-loaded daemon-status.json, cached for the diagnostics panel
 *  (SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal). */
let cachedDaemonStatus = null;

/** Fetches the watchdog-published daemon-status.json and renders it.
 *  Never throws (same best-effort posture as loadHealth). */
async function loadDaemonControlStatus() {
  if (!CONFIG || !CONFIG.daemonControl) return;
  try {
    const { json } = await loadJson(CONFIG.daemonControl.statusEndpoint);
    cachedDaemonStatus = json;
    renderDaemonControlBadge(json);
  } catch {
    cachedDaemonStatus = null;
    renderDaemonControlBadge(null);
  }
}

/** Internal request id, never shown to the user -- same shape as
 *  generateInternalSessionId's `mcv2-` prefix convention above. */
function generateControlRequestId(now) {
  const t = Math.floor(now).toString(36).padStart(8, "0");
  const rand = Math.random().toString(36).slice(2, 8);
  return `mcctl-${t}-${rand}`;
}

/** Plain overwrite of daemon-control.json -- single writer (the PWA),
 *  nothing to merge, same reasoning writeRefreshNudge's sibling health.json
 *  write already documents. */
async function writeDaemonControlRequest(action) {
  const cfg = CONFIG && CONFIG.daemonControl;
  if (!cfg || !cfg.controlEndpoint) return;
  const request = DAEMON_CONTROL_MODEL.buildControlRequest({
    action,
    requestId: generateControlRequestId(Date.now()),
    nowIso: new Date().toISOString(),
  });
  const res = await graphFetch(`${cfg.controlEndpoint}:/content`, {
    method: "PUT",
    body: JSON.stringify(request, null, 2) + "\n",
  });
  if (!res.ok) {
    throw new Error(`daemon-control PUT failed: ${res.status}`);
  }
}

/** True if any v2 session in the current index is currently `running` --
 *  the "N session(s) currently running -- stop anyway?" confirm gate. */
async function anyV2SessionRunning() {
  try {
    const { index } = await loadV2Index();
    return (index.sessions || []).filter((s) => s.status === "running");
  } catch {
    // Best-effort: if the index can't be read, don't block the stop
    // request on it -- fail open on the CONFIRMATION copy (worst case the
    // user isn't warned about a running session), not on the action itself.
    return [];
  }
}

async function handleDaemonStartClick() {
  if (daemonControlRequestInFlight) return;
  daemonControlRequestInFlight = true;
  try {
    await writeDaemonControlRequest("start");
    setTimeout(loadDaemonControlStatus, 1500);
  } catch (err) {
    window.alert(`Failed to send start request: ${err.message}`);
  } finally {
    daemonControlRequestInFlight = false;
  }
}

async function handleDaemonStopClick() {
  if (daemonControlRequestInFlight) return;
  daemonControlRequestInFlight = true;
  try {
    const running = await anyV2SessionRunning();
    if (running.length > 0) {
      const proceed = window.confirm(
        `${running.length} session(s) currently running — stop anyway?`,
      );
      if (!proceed) return;
    }
    await writeDaemonControlRequest("stop");
    setTimeout(loadDaemonControlStatus, 1500);
  } catch (err) {
    window.alert(`Failed to send stop request: ${err.message}`);
  } finally {
    daemonControlRequestInFlight = false;
  }
}

/** SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal: restarts
 *  the whole WSL VM (`wsl --shutdown`), not just the daemon -- strictly
 *  more disruptive than Stop, so this ALWAYS confirms first, regardless of
 *  whether any v2 session looks "running" (a wedged watchdog can't even
 *  report that reliably, which is exactly why this button exists). */
async function handleDaemonRepairClick() {
  if (daemonControlRequestInFlight) return;
  daemonControlRequestInFlight = true;
  try {
    const proceed = window.confirm(
      "This restarts your WSL environment entirely, interrupting anything else running there too. Continue?",
    );
    if (!proceed) return;
    await writeDaemonControlRequest("repair:wsl-restart");
    setTimeout(loadDaemonControlStatus, 1500);
  } catch (err) {
    window.alert(`Failed to send repair request: ${err.message}`);
  } finally {
    daemonControlRequestInFlight = false;
  }
}

function wireDaemonControlButtons() {
  const startBtn = document.getElementById("btn-daemon-start");
  const stopBtn = document.getElementById("btn-daemon-stop");
  const repairBtn = document.getElementById("btn-daemon-repair");
  if (startBtn) startBtn.addEventListener("click", handleDaemonStartClick);
  if (stopBtn) stopBtn.addEventListener("click", handleDaemonStopClick);
  if (repairBtn) repairBtn.addEventListener("click", handleDaemonRepairClick);
}

// =============================================================================
// 7c-3. Global diagnostics panel (SPEC-DELTA-2026-10-01-health-visibility-
//        and-manual-self-heal)
// =============================================================================
//
// One place to see every monitored check, not just the daemon-control strip
// above -- auth (health.json), daemon (daemon-status.json), the sharing
// relay and each active guest's connection (share-relay-status.json +
// shares.json). All severity/likely-cause/root-cause logic lives in
// cockpit-health-model.mjs#buildHealthItems; this only renders whatever
// that returns.

const DIAGNOSTICS_SEVERITY_LABEL = { ok: "OK", degraded: "Degraded", error: "Error", unknown: "Unknown" };

/** Renders the full health-item list into #diagnostics-list. Never throws
 *  (same best-effort posture as the other render* functions here). */
function renderDiagnosticsList() {
  const list = document.getElementById("diagnostics-list");
  if (!list || !COCKPIT_HEALTH_MODEL) return;
  const items = COCKPIT_HEALTH_MODEL.buildHealthItems({
    health: cachedHealthStatus,
    daemonStatus: cachedDaemonStatus,
    relayStatus: cachedRelayStatus,
    shares: cachedShares,
    nowMs: Date.now(),
  });
  const labelById = Object.fromEntries(items.map((it) => [it.id, it.label]));

  list.innerHTML = "";
  for (const item of items) {
    const row = document.createElement("div");
    row.className = "v2-settings-row";
    const label = document.createElement("span");
    label.className = "v2-settings-row-label";
    label.textContent = item.label;
    const badge = document.createElement("span");
    badge.className = "v2-status-chip";
    badge.dataset.tone = item.severity;
    badge.textContent = DIAGNOSTICS_SEVERITY_LABEL[item.severity] || item.severity;
    row.append(label, badge);
    list.appendChild(row);

    const hintText = item.rootCause
      ? `Likely cause: see "${labelById[item.rootCause.id] || item.rootCause.id}" above`
      : item.reason || (item.likelyCauses && item.likelyCauses.join(" ")) || "";
    if (hintText) {
      const hint = document.createElement("div");
      hint.className = "diagnostics-item-hint";
      hint.textContent = hintText;
      list.appendChild(hint);
    }
  }
}

/** Opens the sheet and refreshes everything it shows -- health/daemon-status
 *  are already kept fresh by their own always-on polling loops, but
 *  shares/relay status are only ever refreshed lazily from views that need
 *  them (refreshSharesIfStale), which this panel otherwise never visits. */
async function openDiagnosticsSheet() {
  const sheet = document.getElementById("diagnostics-sheet");
  if (!sheet) return;
  sheet.hidden = false;
  try {
    await refreshSharesIfStale(0);
  } catch {
    // Best-effort -- render with whatever is already cached rather than
    // leaving the panel blank over a secondary signal's fetch failure.
  }
  renderDiagnosticsList();
}

function closeDiagnosticsSheet() {
  const sheet = document.getElementById("diagnostics-sheet");
  if (sheet) sheet.hidden = true;
}

function wireDiagnosticsButton() {
  const btn = document.getElementById("btn-open-diagnostics");
  if (!btn) return;
  btn.addEventListener("click", () => {
    const sheet = document.getElementById("diagnostics-sheet");
    if (sheet && !sheet.hidden) {
      closeDiagnosticsSheet();
    } else {
      openDiagnosticsSheet();
    }
  });
}

// =============================================================================
// 7d. Sharing overview + guest mode entry (SPEC-DELTA-2026-09-29-session-sharing-stage2)
// =============================================================================
//
// Host: the "Shared" tab is "Shared by me" -- every item in shares.json with
// its guests and modes; tapping one opens it. Guest: the same two views are
// driven by guest-app.mjs from the guest's OWN drive (see bootstrap). The
// Stage 1 /me/drive/sharedWithMe + /drives/{id}/items/{id} path was removed:
// in this tenant it can only ever return 403 for the guest.

let GUEST_APP = null; // guest-app.mjs controller when running in guest mode

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

/**
 * Spin a refresh button while `fn` runs (Live L2 findings 2026-09-30: the
 * animation only ever existed on the removed v1 button and the IDE ones).
 * Keeps spinning at least `minMs` so a fast refresh is still visible.
 */
async function withRefreshSpin(buttonId, fn, minMs = 500) {
  const el = document.getElementById(buttonId);
  const started = Date.now();
  if (el) {
    el.classList.add("cockpit-btn-refreshing");
    el.setAttribute("aria-busy", "true");
  }
  try {
    return await fn();
  } finally {
    const left = minMs - (Date.now() - started);
    if (left > 0) await sleepMs(left);
    if (el) {
      el.classList.remove("cockpit-btn-refreshing");
      el.setAttribute("aria-busy", "false");
    }
  }
}

// "Shared by me" (host): By person (default) or By item, per-person pause
// switch, per-item mode buttons and remove (Live L2 findings 2026-09-30).
let sharedViewMode = "person";

function openSharedItem(kind, id) {
  if (kind === "session") {
    setView("v2-detail", { sessionId: id });
    return;
  }
  const source = kind === "claude-tab" ? "claude-code" : "cursor";
  if (ideTrackerSource !== source) setIdeTrackerSource(source);
  loadIdeTabs()
    .then(() => setView("ide-tab-detail", { composerId: id }))
    .catch((err) => showSharedListError(err.message));
}

async function runSharedListAction(fn) {
  clearSharedListError();
  try {
    await fn();
  } catch (err) {
    showSharedListError(err.message);
  }
  await renderSharedList().catch((err) => showSharedListError(err.message));
}

function buildModeButtons(modes, onPick, ariaLabel) {
  const seg = document.createElement("div");
  seg.className = "v2-share-modes";
  seg.setAttribute("role", "radiogroup");
  seg.setAttribute("aria-label", ariaLabel);
  for (const m of modes) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "v2-share-mode-btn";
    btn.textContent = m.label;
    btn.dataset.mode = m.id;
    btn.setAttribute("role", "radio");
    btn.setAttribute("aria-checked", String(m.selected));
    if (!m.selected) {
      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        onPick(m.id);
      });
    }
    seg.appendChild(btn);
  }
  return seg;
}

function renderSharedByPerson(ul) {
  const people = SHARE_UI.deriveSharesByPerson({ shares: cachedShares, relayStatus: cachedRelayStatus });
  for (const p of people) {
    const li = document.createElement("li");
    li.className = "shared-person";
    const head = document.createElement("div");
    head.className = "shared-person-head";
    head.tabIndex = 0;
    head.setAttribute("aria-expanded", "false");
    const who = document.createElement("div");
    who.className = "v2-share-who";
    const name = document.createElement("span");
    name.className = "v2-share-name";
    name.textContent = p.name;
    name.title = p.email;
    who.appendChild(name);
    const meta = document.createElement("span");
    meta.className = "v2-share-conn";
    meta.dataset.state = p.paused ? "paused" : p.connection;
    const countText = `${p.itemCount} item${p.itemCount === 1 ? "" : "s"}`;
    meta.textContent = [p.paused ? "Paused" : p.connectionLabel, countText].filter(Boolean).join(" · ");
    who.appendChild(meta);
    head.appendChild(who);
    const sw = document.createElement("label");
    sw.className = "shared-person-switch";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = !p.paused;
    box.setAttribute("aria-label", `Sharing with ${p.name}`);
    box.addEventListener("change", () => {
      runSharedListAction(() =>
        mutateSharesPwa((s) => SHARE_MODEL.setGuestPaused(s, { email: p.email, paused: !box.checked, now: Date.now() })),
      );
    });
    sw.appendChild(box);
    sw.appendChild(document.createTextNode(p.paused ? "Off" : "On"));
    sw.addEventListener("click", (ev) => ev.stopPropagation());
    head.appendChild(sw);
    li.appendChild(head);
    const list = document.createElement("ul");
    list.className = "shared-person-items";
    list.hidden = true;
    for (const it of p.items) {
      const row = document.createElement("li");
      row.className = "shared-person-item";
      const title = document.createElement("span");
      title.className = "shared-person-item-title";
      title.textContent = `${it.kindLabel}: ${it.title}`;
      title.addEventListener("click", () => openSharedItem(it.kind, it.id));
      row.appendChild(title);
      row.appendChild(
        buildModeButtons(
          it.modes,
          (mode) => {
            if (mode === "control" && !window.confirm(`Give ${p.name} control of this chat? Your own view becomes read-only until you take it back.`)) return;
            runSharedListAction(() =>
              mutateSharesPwa((s) => SHARE_MODEL.setGuestMode(s, { kind: it.kind, id: it.id, email: p.email, mode, now: Date.now() })),
            );
          },
          `Access for ${p.name}`,
        ),
      );
      const del = document.createElement("button");
      del.type = "button";
      del.className = "cockpit-btn cockpit-btn-icon v2-share-remove";
      del.textContent = "🗑";
      del.setAttribute("aria-label", `Stop sharing ${it.title} with ${p.email}`);
      del.addEventListener("click", () => {
        if (!window.confirm(`Stop sharing "${it.title}" with ${p.email}?`)) return;
        runSharedListAction(() =>
          mutateSharesPwa((s) => SHARE_MODEL.removeGuest(s, { kind: it.kind, id: it.id, email: p.email, now: Date.now() })),
        );
      });
      row.appendChild(del);
      list.appendChild(row);
    }
    li.appendChild(list);
    const toggle = () => {
      list.hidden = !list.hidden;
      head.setAttribute("aria-expanded", String(!list.hidden));
    };
    head.addEventListener("click", toggle);
    head.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        toggle();
      }
    });
    ul.appendChild(li);
  }
  return people.length;
}

function renderSharedByItem(ul) {
  const items = (cachedShares && cachedShares.items) || [];
  const kindLabel = { session: "Chat", "cursor-tab": "Cursor tab", "claude-tab": "Claude Code" };
  for (const item of items) {
    const li = document.createElement("li");
    li.className = "cockpit-session-row";
    li.tabIndex = 0;
    const title = document.createElement("span");
    title.className = "cockpit-row-title";
    title.textContent = `${kindLabel[item.kind] || item.kind}: ${item.title || "(untitled)"}`;
    li.appendChild(title);
    const meta = document.createElement("span");
    meta.className = "cockpit-row-meta";
    meta.textContent = SHARE_UI.deriveShareRows({ shares: cachedShares, kind: item.kind, id: item.id, relayStatus: cachedRelayStatus })
      .map((r) => `${r.name} · ${r.paused ? "Paused" : SHARE_UI.MODE_LABELS[r.mode]}${r.connectionLabel ? ` (${r.connectionLabel})` : ""}`)
      .join(", ");
    li.appendChild(meta);
    li.addEventListener("click", () => openSharedItem(item.kind, item.id));
    li.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        openSharedItem(item.kind, item.id);
      }
    });
    ul.appendChild(li);
  }
  return items.length;
}

async function renderSharedList() {
  if (GUEST_APP) return GUEST_APP.renderGuestList();
  clearSharedListError();
  const ul = document.getElementById("shared-session-list");
  const empty = document.getElementById("shared-list-empty-state");
  if (!ul || !empty || !SHARE_MODEL || !SHARE_UI) return;
  await Promise.all([loadShares(), loadRelayStatus().catch(() => null)]);
  sharesLoadedAt = Date.now();
  const controls = document.getElementById("shared-host-controls");
  if (controls) controls.hidden = false;
  for (const btn of document.querySelectorAll("#shared-view-toggle [data-shared-view]")) {
    btn.setAttribute("aria-checked", String(btn.dataset.sharedView === sharedViewMode));
  }
  const stopAll = document.getElementById("btn-shared-stop-all");
  if (stopAll) {
    const st = SHARE_UI.deriveStopAllState(cachedShares);
    stopAll.textContent = st.label;
    stopAll.dataset.stopped = String(st.stopped);
  }
  ul.innerHTML = "";
  const count = sharedViewMode === "person" ? renderSharedByPerson(ul) : renderSharedByItem(ul);
  empty.textContent = "Nothing shared yet. Open a chat or an IDE tab and use its Share panel.";
  empty.hidden = count > 0;
  if (cachedShares && cachedShares.sharingEnabled === false && count) {
    empty.textContent = "Sharing is stopped for everyone. Tap “Sharing is stopped — resume” to turn it back on.";
    empty.hidden = false;
  }
}

async function renderSharedDetail(payload) {
  if (GUEST_APP) return GUEST_APP.renderGuestItem(payload);
  // Host never lands here (its "Shared by me" rows open the real views).
  setView("shared-list");
}

/**
 * Header account menu (SPEC-DELTA-2026-09-29-app-menu): the button toggles
 * the status panel; Escape / an outside tap closes it; the red dot follows
 * the status elements via a MutationObserver, so none of their existing
 * render functions needed to change. Wired first thing in bootstrap so the
 * panel also works when sign-in or config loading fails.
 */
/**
 * One immediate refresh of whatever is on screen, used when the page becomes
 * visible again (SPEC-DELTA-2026-09-29-pwa-polling-hygiene, AC-132).
 */
function refreshVisibleViewOnce() {
  loadHealth();
  loadDaemonControlStatus();
  const v = document.body.dataset.view;
  if (v === "v2-list") {
    renderV2List().catch((err) => showV2ListError(err.message));
  } else if (v === "ide-tabs") {
    renderIdeTabsList().catch((err) => showIdeTabsError(err.message));
  } else if (v === "v2-detail" && activeV2DetailSessionId) {
    const id = activeV2DetailSessionId;
    loadV2Record(id)
      .then(({ record }) => {
        if (!record || activeV2DetailSessionId !== id) return;
        renderV2Messages(record);
        syncComposerFromRecord(record);
        syncV2DetailPoll();
      })
      .catch((err) => showV2DetailError(err.message));
  }
}

function wireAppMenu() {
  const btn = document.getElementById("btn-app-menu");
  const menu = document.getElementById("app-menu");
  if (!btn || !menu) return;
  const setOpen = (open) => {
    menu.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
  };
  btn.addEventListener("click", (ev) => {
    ev.stopPropagation();
    setOpen(menu.hidden);
  });
  document.addEventListener("click", (ev) => {
    if (!menu.hidden && !menu.contains(ev.target) && !btn.contains(ev.target)) setOpen(false);
  });
  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape" && !menu.hidden) setOpen(false);
  });
  const watched = ["status-badge", "health-badge", "daemon-control-badge"]
    .map((id) => document.getElementById(id))
    .filter(Boolean);
  const observer = new MutationObserver(() => updateAppMenuAlert());
  for (const el of watched) {
    observer.observe(el, { attributes: true, childList: true, characterData: true, subtree: true });
  }
  updateAppMenuAlert();
}

function updateAppMenuAlert() {
  const dot = document.getElementById("app-menu-alert-dot");
  const btn = document.getElementById("btn-app-menu");
  if (!dot) return;
  const statusEl = document.getElementById("status-badge");
  const healthEl = document.getElementById("health-badge");
  const daemonEl = document.getElementById("daemon-control-badge");
  const input = {
    statusBadge: statusEl ? statusEl.dataset.status : undefined,
    health: healthEl ? healthEl.dataset.state : undefined,
    daemon: daemonEl ? daemonEl.dataset.state : undefined,
    daemonChecked: daemonEl ? !/checking/.test(daemonEl.textContent || "") : false,
  };
  // Before the helpers load (e.g. a config error), a plain error check.
  const r = APP_MENU_STATE
    ? APP_MENU_STATE.deriveMenuAlert(input)
    : { alert: input.statusBadge === "error", reasons: input.statusBadge === "error" ? ["error"] : [] };
  dot.hidden = !r.alert;
  if (btn) {
    btn.title = r.alert ? `Needs attention: ${r.reasons.join(", ")}` : "Account and status";
  }
}

// =============================================================================
// 8. Bootstrap
// =============================================================================

async function bootstrap() {
  wireAppMenu();
  wireIdeRowSummaryPopover();
  // SPEC-DELTA-2026-10-07-cockpit-pin-density-unified-view.md: auto-pin on
  // load + click-to-toggle + real-state refresh on focus, all pin buttons.
  wirePinButtons();
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
    let GRAPH_BACKOFF_HELPERS;
    [WRITE_HELPERS, IDE_HELPERS, REFRESH_HELPERS, V2_MODEL, SCROLLBACK_HELPERS, DAEMON_CONTROL_MODEL, COMPOSER_STATE, APP_MENU_STATE, GRAPH_BACKOFF_HELPERS, SHARE_MODEL, SHARE_UI, COCKPIT_HEALTH_MODEL, MODEL_CHOICE] = await Promise.all([
      import("./write-helpers.mjs?v=7d8a394"),
      import("./ide-helpers.mjs?v=7d8a394"),
      import("./refresh-helpers.mjs?v=7d8a394"),
      import("./transcript-model.mjs?v=7d8a394"),
      import("./scrollback-helpers.mjs?v=7d8a394"),
      import("./daemon-control-model.mjs?v=7d8a394"),
      import("./composer-state.mjs?v=7d8a394"),
      import("./app-menu-state.mjs?v=7d8a394"),
      import("./graph-backoff.mjs?v=7d8a394"),
      import("./share-model.mjs?v=7d8a394"),
      import("./share-ui-state.mjs?v=7d8a394"),
      import("./cockpit-health-model.mjs?v=7d8a394"),
      import("./model-choice.mjs?v=7d8a394"),
    ]);
    graphBackoff = GRAPH_BACKOFF_HELPERS.createGraphBackoff();
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

  // SPEC-DELTA-2026-09-29-session-sharing-stage2 (AC-150): anyone but the
  // cockpit's owner is a guest -- guest-app.mjs takes over the page and the
  // host wiring below never runs (no host views, no reads or writes under the
  // guest's own cursor-cockpit/).
  if (SHARE_UI.isGuestAccount(activeAccount.username, CONFIG.sharing && CONFIG.sharing.hostUpn)) {
    document.body.dataset.guest = "true";
    setStatusBadge(`signed in: ${activeAccount.username} (guest)`, "ok");
    if (connEl) connEl.textContent = "online";
    try {
      const guestModule = await import("./guest-app.mjs?v=7d8a394");
      GUEST_APP = guestModule.startGuestMode({
        config: CONFIG,
        account: activeAccount,
        graphFetch,
        loadJson,
        putJson,
        shareModel: SHARE_MODEL,
        shareUi: SHARE_UI,
        scrollback: SCROLLBACK_HELPERS,
        ideHelpers: IDE_HELPERS,
        composerState: COMPOSER_STATE,
        setView,
        setStatusBadge,
        translateErrorMessage,
        populateModelSelect,
        withRefreshSpin,
      });
    } catch (err) {
      setStatusBadge(`guest mode error: ${err.message}`, "error");
      return;
    }
    for (const back of document.querySelectorAll(".cockpit-back-btn")) {
      back.addEventListener("click", () => setView(back.dataset.targetView || "shared-list"));
    }
    const btnSharedRefreshGuest = document.getElementById("btn-shared-refresh");
    if (btnSharedRefreshGuest) {
      btnSharedRefreshGuest.addEventListener("click", () => withRefreshSpin("btn-shared-refresh", () => GUEST_APP.renderGuestList()));
    }
    const btnSharedDetailRefreshGuest = document.getElementById("btn-shared-detail-refresh");
    if (btnSharedDetailRefreshGuest) {
      btnSharedDetailRefreshGuest.addEventListener("click", () => withRefreshSpin("btn-shared-detail-refresh", () => GUEST_APP.refreshOnce()));
    }
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) GUEST_APP.refreshOnce();
    });
    GUEST_APP.start(window.location.hash);
    return;
  }

  setStatusBadge(`signed in: ${activeAccount.username} (read-write)`, "ok");
  const initialsEl = document.getElementById("app-menu-initials");
  if (initialsEl && APP_MENU_STATE) {
    initialsEl.textContent = APP_MENU_STATE.deriveInitials(activeAccount.name, activeAccount.username);
  }
  if (connEl) connEl.textContent = "online";

  // SPEC task 12 (AC-008): fire-and-forget, does not block the rest of
  // bootstrap -- a slow/failed health fetch must never delay sign-in.
  loadHealth();
  if (CONFIG.health && Number.isFinite(CONFIG.health.pollIntervalSeconds)) {
    setInterval(() => { if (!document.hidden) loadHealth(); }, CONFIG.health.pollIntervalSeconds * 1000);
  }

  // SPEC-DELTA-2026-09-27-daemon-control-watchdog: same fire-and-forget,
  // never-blocks-boot posture as loadHealth above.
  loadDaemonControlStatus();
  if (CONFIG.daemonControl && Number.isFinite(CONFIG.daemonControl.pollIntervalSeconds)) {
    setInterval(() => { if (!document.hidden) loadDaemonControlStatus(); }, CONFIG.daemonControl.pollIntervalSeconds * 1000);
  }
  wireDaemonControlButtons();
  wireDiagnosticsButton();
  // AC-132: timers skip ticks while hidden, so coming back refreshes once
  // right away instead of showing up-to-a-poll-interval-old data.
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refreshVisibleViewOnce();
  });

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
  for (const btn of document.querySelectorAll(".cockpit-ide-tracker-btn")) {
    btn.addEventListener("click", () => {
      const source = btn.dataset.ideTrackerSource;
      if (source) setIdeTrackerSource(source);
    });
  }
  syncIdeTrackerSourceToggle();

  // v2 (chat-model) button wiring (mobile follow-along, 2026-09-24). Back
  // buttons + the mode-toggle pill are already generic (data-target-view),
  // handled by the loops above.
  const btnV2Refresh = document.getElementById("btn-v2-refresh");
  if (btnV2Refresh) {
    btnV2Refresh.addEventListener("click", () => {
      withRefreshSpin("btn-v2-refresh", () => renderV2List()).catch((err) => showV2ListError(err.message));
    });
  }
  const btnSharedRefresh = document.getElementById("btn-shared-refresh");
  if (btnSharedRefresh) {
    btnSharedRefresh.addEventListener("click", () => {
      withRefreshSpin("btn-shared-refresh", () => renderSharedList()).catch((err) => showSharedListError(err.message));
    });
  }
  for (const btn of document.querySelectorAll("#shared-view-toggle [data-shared-view]")) {
    btn.addEventListener("click", () => {
      sharedViewMode = btn.dataset.sharedView === "item" ? "item" : "person";
      renderSharedList().catch((err) => showSharedListError(err.message));
    });
  }
  const btnSharedStopAll = document.getElementById("btn-shared-stop-all");
  if (btnSharedStopAll) {
    btnSharedStopAll.addEventListener("click", () => {
      handleShareStopAllClick()
        .then(() => renderSharedList())
        .catch((err) => showSharedListError(err.message));
    });
  }
  const btnV2DetailRefresh = document.getElementById("btn-v2-detail-refresh");
  if (btnV2DetailRefresh) {
    btnV2DetailRefresh.addEventListener("click", () => {
      if (activeV2DetailSessionId) {
        const id = activeV2DetailSessionId;
        withRefreshSpin("btn-v2-detail-refresh", () => renderV2Detail(id)).catch((err) => showV2DetailError(err.message));
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
  // "↓ New message" (AC-105): tap to jump; hides itself once the reader
  // reaches the bottom by any means.
  const v2MessagesEl = document.getElementById("v2-messages");
  const btnV2NewMessage = document.getElementById("btn-v2-new-message");
  if (v2MessagesEl && btnV2NewMessage) {
    btnV2NewMessage.addEventListener("click", () => {
      // Instant, not smooth: a smooth scroll was measured stepping ~40 px/s
      // when the tab is throttled, leaving the reader stranded mid-history.
      v2MessagesEl.scrollTop = v2MessagesEl.scrollHeight;
      btnV2NewMessage.hidden = true;
    });
    v2MessagesEl.addEventListener("scroll", () => {
      if (btnV2NewMessage.hidden || !SCROLLBACK_HELPERS) return;
      const atBottom = SCROLLBACK_HELPERS.shouldAutoScrollToBottom({
        scrollTop: v2MessagesEl.scrollTop,
        scrollHeight: v2MessagesEl.scrollHeight,
        clientHeight: v2MessagesEl.clientHeight,
      });
      if (atBottom) btnV2NewMessage.hidden = true;
    }, { passive: true });
  }
  // Contextual buttons + auto-grow follow the box content; Ctrl/Cmd+Enter
  // submits (plain Enter stays a newline).
  const v2ComposerText = document.getElementById("v2-composer-text");
  if (v2ComposerText) {
    v2ComposerText.addEventListener("input", () => {
      applyComposerButtons();
      autoGrowComposer();
    });
    v2ComposerText.addEventListener("keydown", (ev) => {
      if (!COMPOSER_STATE || !COMPOSER_STATE.isSubmitShortcut(ev)) return;
      ev.preventDefault();
      const send = document.getElementById("btn-v2-composer-send");
      if (send && !send.disabled) send.click();
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
  const btnV2Delete = document.getElementById("btn-v2-delete");
  if (btnV2Delete) {
    btnV2Delete.addEventListener("click", () => {
      handleV2DetailDeleteClick().catch((err) => showV2DetailError(err.message));
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
  // SPEC-DELTA-2026-09-29-session-sharing-stage2: Share panels (chat + IDE
  // tab), Stop all sharing, the access banner and the take-over row.
  const v2ShareInviteForm = document.getElementById("v2-share-invite-form");
  if (v2ShareInviteForm) {
    v2ShareInviteForm.addEventListener("submit", (evt) => {
      handleShareInviteSubmit(evt, "chat").catch((err) => showV2DetailError(err.message));
    });
  }
  const ideShareInviteForm = document.getElementById("ide-share-invite-form");
  if (ideShareInviteForm) {
    ideShareInviteForm.addEventListener("submit", (evt) => {
      handleShareInviteSubmit(evt, "ide").catch((err) => showIdeDetailError(err.message));
    });
  }
  const btnShareStopAll = document.getElementById("btn-v2-share-stop-all");
  if (btnShareStopAll) {
    btnShareStopAll.addEventListener("click", () => {
      handleShareStopAllClick().catch((err) => showV2DetailError(err.message));
    });
  }
  const btnAccessAction = document.getElementById("btn-v2-access-action");
  if (btnAccessAction) {
    btnAccessAction.addEventListener("click", () => {
      handleV2AccessAction(btnAccessAction.dataset.action).catch((err) => showV2DetailError(err.message));
    });
  }
  const btnTakeOver = document.getElementById("btn-v2-take-over");
  if (btnTakeOver) {
    btnTakeOver.addEventListener("click", () => {
      handleV2AccessAction("take-over").catch((err) => showV2DetailError(err.message));
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
