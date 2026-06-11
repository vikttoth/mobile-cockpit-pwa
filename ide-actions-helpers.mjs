// mobile-cockpit / pwa / ide-actions-helpers.mjs -- M2.2.3 pure helpers.
//
// Why a separate module (mirrors the rationale in write-helpers.mjs):
//   - pwa/app.js is a browser-side classic script that binds to
//     DOMContentLoaded and reaches for MSAL on the global namespace at
//     boot. Importing it into a Node test would still trigger that flow
//     at module-load, which is brittle for hermetic tests.
//   - The helpers below are PURE -- no DOM, no fetch, no MSAL, no
//     Date.now() reads. They take `now` and `rngFn` as explicit
//     arguments so tests can be deterministic.
//   - app.js consumes them via a one-shot dynamic import() during boot
//     (parallel with write-helpers.mjs + ide-helpers.mjs).
//
// Mirror of the extension's `extension/src/lib/validate.ts`. We rebuild
// the validation in JS because:
//   - The PWA cannot share TypeScript types with the extension (no build
//     step on the PWA side; the extension bundles esbuild -> single .js).
//   - The validation contract is small and rarely changes; static
//     guards (`pwa-ide-actions-helpers-unit.sh` cross-asserts the bound
//     constants against the extension's types.ts file) keep drift loud.

"use strict";

// =============================================================================
// Contract constants -- MUST stay in lockstep with extension's
// `extension/src/lib/types.ts` and `lib/config.mjs`. The static guard
// `tests/flows/mobile-cockpit/pwa-config-coherence.sh` enforces this.
// =============================================================================

export const SUPPORTED_KINDS = Object.freeze([
  "send_message",
  "new_agent",
  "close_tab",
  "stop_agent",
]);

/** Hard cap on `text` field (matches extension validate.ts -- 20 000 chars). */
export const MAX_TEXT_LEN = 20_000;

/** Minimum/maximum length for `actionId` -- matches extension validate.ts. */
export const ACTION_ID_MIN_LEN = 8;
export const ACTION_ID_MAX_LEN = 128;

/**
 * Allowed characters in `actionId` -- alphanumeric + hyphen + underscore.
 * Defensive against path-traversal if the extension ever materializes
 * actionId into a filename (e.g., for crash-recovery audit logs).
 */
export const ACTION_ID_CHAR_RE = /^[A-Za-z0-9_-]+$/;

// =============================================================================
// generateActionId
// =============================================================================
//
// Format: `ide-<now-in-base36>-<6hex-from-rng>`.
//
// Prefix `ide-` distinguishes mobile-issued IDE actions from session-create
// (`pwa-` prefix in write-helpers.mjs) and daemon-side append-test (`test-`).
//
// rngFn signature: `(n: number) => Uint8Array` (or Buffer in Node). The
// production caller passes `cryptoRandomBytes` from write-helpers.mjs;
// tests inject a deterministic stub.

/** Coerce epoch-ms from number or Date (defensive — callers should pass Date.now()). */
export function coerceNowMs(now) {
  if (now instanceof Date) return now.getTime();
  if (typeof now === "number" && Number.isFinite(now)) return now;
  throw new Error("coerceNowMs(now): `now` must be a finite number (ms epoch) or Date");
}

export function generateActionId(now, rngFn) {
  const ms = coerceNowMs(now);
  if (typeof rngFn !== "function") {
    throw new Error("generateActionId(now, rngFn): `rngFn` must be a function returning a Uint8Array");
  }
  const t = Math.floor(ms).toString(36).padStart(8, "0");
  const bytes = rngFn(4);
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    hex += bytes[i].toString(16).padStart(2, "0");
  }
  return `ide-${t}-${hex.slice(0, 6)}`;
}

// =============================================================================
// buildSendMessageAction / buildNewAgentAction / buildCloseTabAction
// =============================================================================
//
// Build the action JSON object that gets appended to ide-actions.json.
// Each helper takes the per-kind required fields + the same {now, rngFn}
// pair as `generateActionId` (so the caller doesn't need a separate id
// generation step). Returns the fully-formed action object with
// `status: "pending"`.
//
// These do NOT validate -- call `validateActionInputs(action)` separately
// before submitting. We split build + validate because build always
// succeeds (or throws on type errors), whereas validate returns a
// {valid, errors} pair that the UI can use to show inline form errors.

export function buildSendMessageAction(opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("buildSendMessageAction(opts): opts is required");
  }
  const { tabId, text, now, rngFn, sendMode } = opts;
  if (typeof tabId !== "string") throw new Error("buildSendMessageAction: tabId must be a string");
  if (typeof text !== "string") throw new Error("buildSendMessageAction: text must be a string");
  const ms = coerceNowMs(now);
  const action = {
    actionId: generateActionId(ms, rngFn),
    createdAt: new Date(ms).toISOString(),
    kind: "send_message",
    status: "pending",
    tabId: tabId,
    text: text,
  };
  if (sendMode === "interrupt") {
    action.sendMode = "interrupt";
  }
  return action;
}

export function buildNewAgentAction(opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("buildNewAgentAction(opts): opts is required");
  }
  const { workspace, text, model, now, rngFn } = opts;
  if (typeof workspace !== "string") throw new Error("buildNewAgentAction: workspace must be a string");
  const action = {
    actionId: generateActionId(now, rngFn),
    createdAt: new Date(now).toISOString(),
    kind: "new_agent",
    status: "pending",
    workspace: workspace,
  };
  if (typeof text === "string" && text.length > 0) action.text = text;
  if (typeof model === "string" && model.length > 0) action.model = model;
  return action;
}

export function buildCloseTabAction(opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("buildCloseTabAction(opts): opts is required");
  }
  const { tabId, now, rngFn } = opts;
  if (typeof tabId !== "string") throw new Error("buildCloseTabAction: tabId must be a string");
  return {
    actionId: generateActionId(now, rngFn),
    createdAt: new Date(now).toISOString(),
    kind: "close_tab",
    status: "pending",
    tabId: tabId,
  };
}

export function buildStopAgentAction(opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("buildStopAgentAction(opts): opts is required");
  }
  const { tabId, now, rngFn } = opts;
  if (typeof tabId !== "string") throw new Error("buildStopAgentAction: tabId must be a string");
  const ms = coerceNowMs(now);
  return {
    actionId: generateActionId(ms, rngFn),
    createdAt: new Date(ms).toISOString(),
    kind: "stop_agent",
    status: "pending",
    tabId: tabId,
  };
}

// =============================================================================
// Tab id matching (mirror extension/src/lib/tab-id.ts)
// =============================================================================

const UUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export function bareComposerId(tabId) {
  const t = typeof tabId === "string" ? tabId.trim() : tabId == null ? "" : String(tabId).trim();
  if (!t) return "";
  if (t.startsWith("composer:")) return t.slice("composer:".length);
  return t;
}

export function tabIdsMatch(a, b) {
  const ba = bareComposerId(a).toLowerCase();
  const bb = bareComposerId(b).toLowerCase();
  if (!ba || !bb) return false;
  if (UUID_RE.test(ba) && UUID_RE.test(bb)) return ba === bb;
  return ba === bb;
}

const TERMINAL_RESULT_STATUS = new Set(["done", "error", "skipped"]);

export function isTerminalResult(result) {
  return result && TERMINAL_RESULT_STATUS.has(result.status);
}

/**
 * Actions for `tabId` not yet finished per `resultsFile` (extension outcome).
 * @returns {Array<{action: object, result: object|null, phase: string}>}
 */
export function listOutstandingForTab(actionFile, resultsFile, tabId) {
  if (!actionFile || typeof actionFile !== "object") return [];
  const actions = Array.isArray(actionFile.actions) ? actionFile.actions : [];
  const out = [];
  for (const action of actions) {
    if (!action || typeof action !== "object") continue;
    if (action.kind !== "send_message" && action.kind !== "stop_agent") continue;
    if (!tabIdsMatch(action.tabId, tabId)) continue;
    const result = findResultForAction(resultsFile, action.actionId);
    if (isTerminalResult(result)) continue;
    const phase = result ? "delivering" : "queued";
    out.push({ action, result, phase });
  }
  out.sort((a, b) => {
    const ta = a.action.createdAt || "";
    const tb = b.action.createdAt || "";
    return ta < tb ? -1 : ta > tb ? 1 : 0;
  });
  return out;
}

export function removeActionById(prevFile, actionId, now) {
  if (!actionId || typeof actionId !== "string") {
    throw new Error("removeActionById: actionId is required");
  }
  const isoNow = now instanceof Date ? now.toISOString() : new Date(now ?? Date.now()).toISOString();
  const base =
    prevFile && typeof prevFile === "object"
      ? prevFile
      : { schemaVersion: 1, snapshotAt: isoNow, actions: [] };
  const actions = Array.isArray(base.actions) ? [...base.actions] : [];
  const next = actions.filter((a) => !(a && a.actionId === actionId));
  if (next.length === actions.length) {
    const err = new Error(`action not found: ${actionId}`);
    err.code = "NOT_FOUND";
    throw err;
  }
  return {
    schemaVersion: base.schemaVersion || 1,
    snapshotAt: isoNow,
    actions: next,
  };
}

export function patchActionById(prevFile, actionId, patch, now) {
  if (!actionId || typeof actionId !== "string") {
    throw new Error("patchActionById: actionId is required");
  }
  if (!patch || typeof patch !== "object") {
    throw new Error("patchActionById: patch is required");
  }
  const isoNow = now instanceof Date ? now.toISOString() : new Date(now ?? Date.now()).toISOString();
  const base =
    prevFile && typeof prevFile === "object"
      ? prevFile
      : { schemaVersion: 1, snapshotAt: isoNow, actions: [] };
  const actions = Array.isArray(base.actions) ? [...base.actions] : [];
  let found = false;
  const nextActions = actions.map((a) => {
    if (!a || a.actionId !== actionId) return a;
    found = true;
    return { ...a, ...patch };
  });
  if (!found) {
    const err = new Error(`action not found: ${actionId}`);
    err.code = "NOT_FOUND";
    throw err;
  }
  return {
    schemaVersion: base.schemaVersion || 1,
    snapshotAt: isoNow,
    actions: nextActions,
  };
}

// =============================================================================
// validateActionInputs
// =============================================================================
//
// Returns {valid: boolean, errors: string[]}. Mirrors the extension's
// `validateAction` in `extension/src/lib/validate.ts` -- if the PWA-side
// check passes, the extension-side check will pass too. This is a
// pre-flight: it catches bad input BEFORE the OneDrive PUT, so the user
// gets immediate feedback in the form rather than waiting for the
// extension to skip the action.
//
// `allowedCwds` is required for `new_agent` validation. Pass `[]` if
// the caller does not have it (validation will then reject every
// new_agent workspace, which is the safer default).

export function validateActionInputs(action, allowedCwds) {
  const errors = [];
  if (!action || typeof action !== "object") {
    return { valid: false, errors: ["action must be an object"] };
  }
  const allow = Array.isArray(allowedCwds) ? allowedCwds : [];

  // actionId
  if (typeof action.actionId !== "string") {
    errors.push("actionId is required and must be a string");
  } else if (action.actionId.length < ACTION_ID_MIN_LEN) {
    errors.push(`actionId is too short (${action.actionId.length} chars; min ${ACTION_ID_MIN_LEN})`);
  } else if (action.actionId.length > ACTION_ID_MAX_LEN) {
    errors.push(`actionId is too long (${action.actionId.length} chars; max ${ACTION_ID_MAX_LEN})`);
  } else if (!ACTION_ID_CHAR_RE.test(action.actionId)) {
    errors.push("actionId contains disallowed characters (only A-Z, a-z, 0-9, -, _)");
  }

  // createdAt
  if (typeof action.createdAt !== "string") {
    errors.push("createdAt is required and must be an ISO-8601 string");
  } else if (!Number.isFinite(Date.parse(action.createdAt))) {
    errors.push("createdAt is not a parseable ISO-8601 string");
  }

  // kind
  if (typeof action.kind !== "string") {
    errors.push("kind is required and must be a string");
  } else if (!SUPPORTED_KINDS.includes(action.kind)) {
    errors.push(`kind "${action.kind}" is not supported; expected one of: ${SUPPORTED_KINDS.join(", ")}`);
  } else {
    // Per-kind required fields
    if (action.kind === "send_message") {
      if (typeof action.tabId !== "string" || action.tabId.length === 0) {
        errors.push("send_message: tabId is required");
      }
      if (typeof action.text !== "string" || action.text.length === 0) {
        errors.push("send_message: text is required and must not be empty");
      } else if (action.text.length > MAX_TEXT_LEN) {
        errors.push(`send_message: text is too long (${action.text.length} chars; max ${MAX_TEXT_LEN})`);
      }
      if (action.sendMode !== undefined && action.sendMode !== null) {
        if (action.sendMode !== "queue" && action.sendMode !== "interrupt") {
          errors.push('send_message: sendMode must be "queue" or "interrupt"');
        }
      }
    } else if (action.kind === "new_agent") {
      if (typeof action.workspace !== "string" || action.workspace.length === 0) {
        errors.push("new_agent: workspace is required");
      } else if (!allow.includes(action.workspace)) {
        errors.push(
          `new_agent: workspace "${action.workspace}" is not in the allowlist; expected one of: ${allow.join(", ") || "(empty)"}`,
        );
      }
      if (action.text !== undefined && action.text !== null) {
        if (typeof action.text !== "string") {
          errors.push("new_agent: text must be a string when provided");
        } else if (action.text.length > MAX_TEXT_LEN) {
          errors.push(`new_agent: text is too long (${action.text.length} chars; max ${MAX_TEXT_LEN})`);
        }
      }
      if (action.model !== undefined && action.model !== null && typeof action.model !== "string") {
        errors.push("new_agent: model must be a string when provided");
      }
    } else if (action.kind === "close_tab" || action.kind === "stop_agent") {
      if (typeof action.tabId !== "string" || action.tabId.length === 0) {
        errors.push(`${action.kind}: tabId is required`);
      }
    }
  }

  return { valid: errors.length === 0, errors };
}

// =============================================================================
// mergeAppendAction
// =============================================================================
//
// Pure state-merger. Returns a new ActionFile with `action` appended.
// Does NOT mutate the input.
//
// - prevFile null / non-object / missing actions[] -> seeded
//     `{ schemaVersion: 1, snapshotAt, actions: [] }`.
// - Collision (existing action with same actionId) -> throws Error with
//   `code: "ID_COLLISION"`. Same loud-failure pattern as
//   `mergeAppendSession` in write-helpers.mjs.
// - `snapshotAt` is always set to `now.toISOString()` so the extension
//   can detect new writes via the top-level timestamp.

export function mergeAppendAction(prevFile, action, now) {
  if (!action || typeof action !== "object" || typeof action.actionId !== "string") {
    throw new Error("mergeAppendAction: action.actionId is required");
  }
  const isoNow = now instanceof Date ? now.toISOString() : new Date(now ?? Date.now()).toISOString();
  const base =
    prevFile && typeof prevFile === "object"
      ? prevFile
      : { schemaVersion: 1, snapshotAt: isoNow, actions: [] };
  const actions = Array.isArray(base.actions) ? [...base.actions] : [];
  const exists = actions.some((a) => a && a.actionId === action.actionId);
  if (exists) {
    const err = new Error(`action id collision: ${action.actionId}`);
    err.code = "ID_COLLISION";
    throw err;
  }
  actions.push(action);
  return {
    schemaVersion: base.schemaVersion || 1,
    snapshotAt: isoNow,
    actions: actions,
  };
}

// =============================================================================
// findResultForAction
// =============================================================================
//
// Look up the extension's outcome for a given actionId. Returns the
// result object (`{actionId, completedAt, status, error?, outcome?}`) or
// `null` if the extension has not yet written one.
//
// Defensive on garbage input (returns null) -- the polling caller in
// app.js loops on this and we don't want a malformed results file to
// crash the whole UI loop.

export function findResultForAction(resultFile, actionId) {
  if (!resultFile || typeof resultFile !== "object") return null;
  if (typeof actionId !== "string" || actionId.length === 0) return null;
  const results = Array.isArray(resultFile.results) ? resultFile.results : [];
  for (const r of results) {
    if (r && r.actionId === actionId) return r;
  }
  return null;
}

// =============================================================================
// summarizePendingActions
// =============================================================================
//
// Counts by status for the UI badges (e.g., "3 pending, 1 in-progress").
// Defensive on garbage input.

export function summarizePendingActions(actionFile) {
  const out = { pending: 0, in_progress: 0, done: 0, error: 0, skipped: 0, total: 0 };
  if (!actionFile || typeof actionFile !== "object") return out;
  const actions = Array.isArray(actionFile.actions) ? actionFile.actions : [];
  for (const a of actions) {
    out.total += 1;
    const s = a && typeof a.status === "string" ? a.status : "pending";
    if (s in out) out[s] += 1;
    else out.pending += 1;
  }
  return out;
}
