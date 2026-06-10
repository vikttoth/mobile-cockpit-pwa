// mobile-cockpit / pwa / write-helpers.mjs -- Stage A2 pure helpers.
//
// Why a separate module:
//   - pwa/app.js is a browser-side script that binds DOMContentLoaded
//     (behind a typeof-check guard) and reaches for an auth library
//     exposed on the global namespace at boot. Importing it from a
//     Node-side unit test would still trigger the auth flow at module
//     load time, which is brittle for hermetic tests.
//   - The four helpers below are PURE -- no DOM, no fetch, no auth-lib
//     coupling. Living in their own ES module lets Node tests `import`
//     them cleanly (see tests/flows/mobile-cockpit/pwa-write-helpers-unit.sh).
//   - app.js consumes them via a one-shot dynamic import() during boot
//     (it stays a classic script for back-compat with the deferred
//     auth-lib bundle that loads first).
//
// All four helpers mirror semantics of daemon/append-test-session.mjs
// so the read-modify-write+retry-on-412 pattern is identical on both
// sides of the OneDrive state.json. The only practical difference is
// the random-bytes source (browser crypto.getRandomValues vs Node
// crypto.randomBytes); production code passes the wrapper, tests
// inject a deterministic Uint8Array stub.

"use strict";

// =============================================================================
// generateSessionId
// =============================================================================
//
// Format: `pwa-<now-in-base36>-<6hex-from-rng>`.
//
// Daemon uses the `test-` prefix; PWA uses `pwa-` so the two surfaces stay
// trivially distinguishable when scanning state.json — useful while we still
// run the append-test CLI alongside the PWA for write-path round-trips.
//
// rngFn is REQUIRED to be a function returning a Uint8Array (or Buffer
// in Node) of at least the requested length. Production callers should
// pass:
//
//     (n) => { const a = new Uint8Array(n); crypto.getRandomValues(a); return a; }
//
// Tests can pass a deterministic stub. NO default — leaving this
// implicit risks accidentally calling Math.random() in some future
// refactor, which would weaken id uniqueness across tabs / users.

export function generateSessionId(now, rngFn) {
  if (typeof now !== "number" || !Number.isFinite(now)) {
    throw new Error("generateSessionId(now, rngFn): `now` must be a finite number (ms epoch)");
  }
  if (typeof rngFn !== "function") {
    throw new Error("generateSessionId(now, rngFn): `rngFn` must be a function returning a Uint8Array");
  }
  const t = Math.floor(now).toString(36).padStart(8, "0");
  const bytes = rngFn(4);
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    hex += bytes[i].toString(16).padStart(2, "0");
  }
  return `pwa-${t}-${hex.slice(0, 6)}`;
}

// =============================================================================
// validateCreateInputs
// =============================================================================
//
// Returns { valid: boolean, errors: string[] }.
// Surface used by the new-session form before we commit a PUT.
//
// Rules (mirror daemon side):
//   - prompt: required, non-empty after trim, ≤ 8000 chars (design.md §3
//     legacy SharePoint cap, carried over as a conservative practical
//     bound for the OneDrive JSON to keep state.json under ~1 MB even
//     with hundreds of sessions).
//   - cwd: optional. If provided, must appear in allowedCwds (case-
//     sensitive, exact match). Empty string is acceptable — daemon
//     reads it as "use daemon default".
//   - model: optional, if provided must be a string. Daemon accepts
//     null = its own default.

const PROMPT_MAX = 8000;

export function validateCreateInputs(input, allowedCwds) {
  const errors = [];
  const i = input && typeof input === "object" ? input : {};
  const allow = Array.isArray(allowedCwds) ? allowedCwds : [];

  // prompt
  const prompt = typeof i.prompt === "string" ? i.prompt : "";
  if (prompt.trim().length === 0) {
    errors.push("prompt is required and must not be empty");
  } else if (prompt.length > PROMPT_MAX) {
    errors.push(`prompt is too long (${prompt.length} chars; max ${PROMPT_MAX})`);
  }

  // cwd
  if (i.cwd !== undefined && i.cwd !== null && i.cwd !== "") {
    if (typeof i.cwd !== "string") {
      errors.push("cwd must be a string or empty");
    } else if (!allow.includes(i.cwd)) {
      errors.push(
        `cwd "${i.cwd}" is not in the allowlist; expected one of: ${allow.join(", ") || "(empty)"}`,
      );
    }
  }

  // model
  if (i.model !== undefined && i.model !== null && i.model !== "") {
    if (typeof i.model !== "string") {
      errors.push("model must be a string when provided");
    }
  }

  return { valid: errors.length === 0, errors };
}

// =============================================================================
// mergeAppendSession
// =============================================================================
//
// Pure state-merger. Returns a new state object with `session` appended.
// Does NOT mutate the input.
//
// - prevState `null` / non-object → seeded with `{ schemaVersion: 1, sessions: [] }`.
// - Collision (existing session with same sessionId) → throws Error with
//   `code: "ID_COLLISION"`. The PWA caller does NOT auto-replace —
//   collisions should be impossible in practice (the time-based id is
//   ~2^48 of entropy) and if one happens we want loud failure rather
//   than silent data overwrite.

export function mergeAppendSession(prevState, session) {
  if (!session || typeof session !== "object" || typeof session.sessionId !== "string") {
    throw new Error("mergeAppendSession: session.sessionId is required");
  }
  const base =
    prevState && typeof prevState === "object"
      ? prevState
      : { schemaVersion: 1, sessions: [] };
  const sessions = Array.isArray(base.sessions) ? [...base.sessions] : [];
  const exists = sessions.some((s) => s && s.sessionId === session.sessionId);
  if (exists) {
    const err = new Error(`session id collision: ${session.sessionId}`);
    err.code = "ID_COLLISION";
    throw err;
  }
  sessions.push(session);
  return { ...base, schemaVersion: base.schemaVersion || 1, sessions };
}

// =============================================================================
// mergeUpdateStatus
// =============================================================================
//
// Pure state-updater. Finds the session by `sessionId`, sets its `status`
// to `newStatus` and `lastUpdated` to `now.toISOString()`. Returns a NEW
// state object; does not mutate the input.
//
// - Missing session → throws Error with `code: "SESSION_NOT_FOUND"`.
// - Empty sessions array → throws same code (logically identical:
//   "the id you asked for is not present").
// - newStatus is accepted as any string; the schema enum (pending /
//   running / approved / cancelled / done / failed) is enforced by the
//   daemon, not the PWA, so the helper does NOT gate. This keeps the
//   helper pure (no enum import) and lets future statuses land without
//   needing a coordinated PWA+daemon deploy.

export function mergeUpdateStatus(prevState, sessionId, newStatus, now) {
  if (typeof sessionId !== "string" || sessionId.length === 0) {
    throw new Error("mergeUpdateStatus: sessionId is required");
  }
  if (typeof newStatus !== "string" || newStatus.length === 0) {
    throw new Error("mergeUpdateStatus: newStatus is required");
  }
  const isoNow = now instanceof Date ? now.toISOString() : new Date(now ?? Date.now()).toISOString();

  const base = prevState && typeof prevState === "object" ? prevState : { schemaVersion: 1, sessions: [] };
  const sessions = Array.isArray(base.sessions) ? base.sessions : [];
  const idx = sessions.findIndex((s) => s && s.sessionId === sessionId);
  if (idx < 0) {
    const err = new Error(`session not found: ${sessionId}`);
    err.code = "SESSION_NOT_FOUND";
    throw err;
  }
  const nextSessions = sessions.map((s, i) =>
    i === idx ? { ...s, status: newStatus, lastUpdated: isoNow } : s,
  );
  return { ...base, schemaVersion: base.schemaVersion || 1, sessions: nextSessions };
}

// =============================================================================
// mergeQueueFollowUp
// =============================================================================
//
// Re-queue a finished session with a new prompt, preserving the cursor-agent
// chat id for `--resume`. Clears prior output/errors and sets status back to
// pending (or approved when autoApprove is true).

export function mergeQueueFollowUp(prevState, sessionId, { prompt, now, autoApprove }) {
  if (typeof sessionId !== "string" || sessionId.length === 0) {
    throw new Error("mergeQueueFollowUp: sessionId is required");
  }
  if (typeof prompt !== "string" || prompt.trim().length === 0) {
    throw new Error("mergeQueueFollowUp: prompt is required");
  }
  const isoNow = now instanceof Date ? now.toISOString() : new Date(now ?? Date.now()).toISOString();
  const base = prevState && typeof prevState === "object" ? prevState : { schemaVersion: 1, sessions: [] };
  const sessions = Array.isArray(base.sessions) ? base.sessions : [];
  const idx = sessions.findIndex((s) => s && s.sessionId === sessionId);
  if (idx < 0) {
    const err = new Error(`session not found: ${sessionId}`);
    err.code = "SESSION_NOT_FOUND";
    throw err;
  }
  const current = sessions[idx];
  if (current.status !== "done") {
    const err = new Error(`follow-up requires status=done (got ${current.status})`);
    err.code = "INVALID_STATUS";
    throw err;
  }
  const resumeId = current.cursorAgentId || current.agentId;
  if (!resumeId) {
    const err = new Error("follow-up requires cursorAgentId on the session");
    err.code = "NO_AGENT_ID";
    throw err;
  }
  const nextSessions = sessions.map((s, i) =>
    i === idx
      ? {
          ...s,
          prompt: prompt.trim(),
          output: "",
          errorMessage: undefined,
          errorRetryable: undefined,
          completedAt: undefined,
          startedAt: undefined,
          resumeAgentId: resumeId,
          cursorAgentId: resumeId,
          status: autoApprove ? "approved" : "pending",
          lastUpdated: isoNow,
        }
      : s,
  );
  return { ...base, schemaVersion: base.schemaVersion || 1, sessions: nextSessions };
}

// =============================================================================
// Hash routing — deep links (#detail/<sessionId>) for Teams notify + bookmarks
// =============================================================================
//
// MSAL OAuth redirects use fragments like #code=... or #error=... — never
// treat those as app routes. Only #list, #new, #detail/<id> are ours.

const MSAL_OAUTH_FRAGMENT_RE =
  /^(code|error|error_description|state|client_info|session_state)=/i;

/**
 * Parse a location hash string (e.g. "#detail/<sessionId>") into a route, or null if not ours.
 * @returns {{ view: 'list'|'new'|'detail', sessionId?: string } | null}
 */
export function parseLocationHash(hash) {
  if (hash == null || hash === "" || hash === "#") {
    return { view: "list" };
  }
  const raw = String(hash).replace(/^#/, "").trim();
  if (!raw) return { view: "list" };
  if (MSAL_OAUTH_FRAGMENT_RE.test(raw)) return null;
  if (raw === "list") return { view: "list" };
  if (raw === "new") return { view: "new" };
  const detailMatch = /^detail\/([^/?#]+)$/.exec(raw);
  if (detailMatch) {
    const sessionId = decodeURIComponent(detailMatch[1]);
    if (sessionId) return { view: "detail", sessionId };
  }
  return null;
}

/**
 * Build a hash string for the given view (without leading #).
 * Returns null for views that should not update the URL (ide-* modals).
 */
export function formatViewHash(viewId, payload) {
  if (viewId === "list") return "list";
  if (viewId === "new") return "new";
  if (viewId === "detail" && payload && payload.sessionId) {
    return `detail/${encodeURIComponent(payload.sessionId)}`;
  }
  return null;
}

// =============================================================================
// cryptoRandomBytes -- production rng wrapper for the browser
// =============================================================================
//
// Exposed so app.js can pass it to generateSessionId without re-implementing
// the boilerplate. Relies on globalThis.crypto.getRandomValues which is
// present in every modern browser and in Node >= 19 (the unit test runs
// on Node 24, well above that floor). If a future runtime drops it, the
// helper throws a clear error rather than silently falling back to
// Math.random() -- session-id entropy is a correctness invariant.
//
// We intentionally do NOT load node:crypto here. .mjs ES modules have no
// `require()` and a static `import` of "node:crypto" would break this
// file's ability to load in a browser. If you ever need a Node-only rng
// path, do it in the caller (Node test or daemon), not in this shared
// pure-module.

export function cryptoRandomBytes(n) {
  if (
    typeof globalThis.crypto !== "undefined" &&
    typeof globalThis.crypto.getRandomValues === "function"
  ) {
    const a = new Uint8Array(n);
    globalThis.crypto.getRandomValues(a);
    return a;
  }
  throw new Error(
    "cryptoRandomBytes: globalThis.crypto.getRandomValues is not available; " +
      "callers in non-browser environments should pass a Uint8Array-returning rngFn directly.",
  );
}
