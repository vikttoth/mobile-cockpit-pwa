// mobile-cockpit / lib / daemon-control-model.mjs
//
// SPEC-DELTA-2026-09-27-daemon-control-watchdog. Pure logic (no I/O, no
// `vscode`/`node:`/DOM imports) shared byte-for-byte with
// pwa/daemon-control-model.mjs (enforced by
// tests/flows/mobile-cockpit/pwa-daemon-control-coherence.sh, same shape as
// pwa-transcript-model-coherence.sh) so the browser writes the exact same
// request shape the watchdog reads, and renders the exact same display
// state the watchdog's own status write implies.
//
// Background this exists to fix: the daemon's only prior health signal
// (health.json) is written BY the daemon itself, so a dead daemon can never
// report its own death -- the cockpit just showed the last known "ok"
// forever. daemon-status.json is written by a SEPARATE watchdog process on
// every tick regardless of whether a start/stop request came in, so its
// absence/staleness is itself informative (see deriveControlDisplayState).

"use strict";

export const DAEMON_CONTROL_SCHEMA_VERSION = 1;
export const DAEMON_STATUS_SCHEMA_VERSION = 1;

/** @typedef {'start'|'stop'|'repair:wsl-restart'} DaemonControlAction */

const VALID_ACTIONS = ["start", "stop", "repair:wsl-restart"];

/**
 * Build the JSON the PWA writes to `daemon-control.json`.
 *
 * @param {object} opts
 * @param {DaemonControlAction} opts.action
 * @param {string} opts.requestId
 * @param {string} opts.nowIso
 */
export function buildControlRequest(opts) {
  const { action, requestId, nowIso } = opts || {};
  if (!VALID_ACTIONS.includes(action)) {
    throw new Error(
      `buildControlRequest: action must be one of ${VALID_ACTIONS.join(", ")}, got ${JSON.stringify(action)}`,
    );
  }
  if (typeof requestId !== "string" || !requestId.trim()) {
    throw new Error("buildControlRequest: requestId is required");
  }
  if (typeof nowIso !== "string" || !nowIso) {
    throw new Error("buildControlRequest: nowIso is required");
  }
  return {
    schemaVersion: DAEMON_CONTROL_SCHEMA_VERSION,
    action,
    requestId,
    requestedAt: nowIso,
  };
}

/**
 * Defensive parse of whatever is currently in `daemon-control.json` (or
 * `null`/garbage if the file has never been written, mirrors
 * refresh-helpers.mjs#parseRefreshSignals's "treat corrupt as empty"
 * posture, except there is no meaningful "empty" control request, so this
 * returns `null` instead of a default object).
 *
 * @param {unknown} raw
 * @returns {{schemaVersion: number, action: DaemonControlAction, requestId: string, requestedAt: string|null} | null}
 */
export function parseDaemonControl(raw) {
  if (!raw || typeof raw !== "object") return null;
  const o = /** @type {Record<string, unknown>} */ (raw);
  if (!VALID_ACTIONS.includes(/** @type {string} */ (o.action))) return null;
  if (typeof o.requestId !== "string" || !o.requestId) return null;
  return {
    schemaVersion: typeof o.schemaVersion === "number" ? o.schemaVersion : DAEMON_CONTROL_SCHEMA_VERSION,
    action: o.action,
    requestId: o.requestId,
    requestedAt: typeof o.requestedAt === "string" ? o.requestedAt : null,
  };
}

/**
 * True when `control` names a requestId the watchdog has not already acted
 * on -- the guard against re-acting on a stale request every tick (a
 * one-shot-per-invocation script has no in-memory "already did this"
 * state, so this comparison against a caller-supplied last-handled id is
 * the whole mechanism).
 *
 * @param {ReturnType<typeof parseDaemonControl>} control
 * @param {string|null|undefined} lastHandledRequestId
 */
export function isRequestUnhandled(control, lastHandledRequestId) {
  if (!control || !control.requestId) return false;
  if (!lastHandledRequestId) return true;
  return control.requestId !== lastHandledRequestId;
}

/**
 * Build the JSON the watchdog writes to `daemon-status.json` on EVERY
 * tick, regardless of whether a request came in.
 *
 * @param {object} opts
 * @param {string} opts.nowIso
 * @param {boolean} opts.daemonAlive
 * @param {boolean} opts.ideMirrorAlive
 * @param {{action: DaemonControlAction, requestId: string, at?: string, result?: 'ok'|'error', error?: string|null}|null} [opts.lastAction]
 */
export function buildStatusPayload(opts) {
  const { nowIso, daemonAlive, ideMirrorAlive, lastAction } = opts || {};
  if (typeof nowIso !== "string" || !nowIso) {
    throw new Error("buildStatusPayload: nowIso is required");
  }
  return {
    schemaVersion: DAEMON_STATUS_SCHEMA_VERSION,
    watchdogCheckedAt: nowIso,
    daemonAlive: !!daemonAlive,
    ideMirrorAlive: !!ideMirrorAlive,
    lastAction: lastAction
      ? {
          action: lastAction.action ?? null,
          requestId: lastAction.requestId ?? null,
          at: typeof lastAction.at === "string" ? lastAction.at : nowIso,
          result: lastAction.result === "error" ? "error" : "ok",
          error: lastAction.error ?? null,
        }
      : null,
  };
}

/**
 * Pure derivation of the PWA's always-visible status-strip state from a
 * raw `daemon-status.json` payload (or `null`/garbage if it has never been
 * published). Distinguishes "confidently know it's stopped" from "don't
 * actually know" (missing/stale watchdog) rather than defaulting either
 * one to a guessed "running" -- the whole point of this feature is that a
 * dead watchdog must never look the same as a healthy "stopped".
 *
 * @param {object} opts
 * @param {any} opts.status
 * @param {number} opts.nowMs
 * @param {number} opts.staleAfterMs
 * @returns {'running'|'stopped'|'error'|'unknown'}
 */
export function deriveControlDisplayState(opts) {
  const { status, nowMs, staleAfterMs } = opts || {};
  if (!status || typeof status.watchdogCheckedAt !== "string") return "unknown";
  const checkedAtMs = Date.parse(status.watchdogCheckedAt);
  if (!Number.isFinite(checkedAtMs)) return "unknown";
  if (typeof staleAfterMs === "number" && typeof nowMs === "number" && nowMs - checkedAtMs > staleAfterMs) {
    return "unknown";
  }
  if (status.lastAction && status.lastAction.result === "error") return "error";
  if (status.daemonAlive && status.ideMirrorAlive) return "running";
  return "stopped";
}

/**
 * How long ago (seconds) `watchdogCheckedAt` was, for the "checked Ns ago"
 * readout. Returns `null` when it cannot be computed (mirrors
 * deriveControlDisplayState's own defensiveness).
 *
 * @param {any} status
 * @param {number} nowMs
 * @returns {number|null}
 */
export function secondsSinceChecked(status, nowMs) {
  if (!status || typeof status.watchdogCheckedAt !== "string") return null;
  const checkedAtMs = Date.parse(status.watchdogCheckedAt);
  if (!Number.isFinite(checkedAtMs) || typeof nowMs !== "number") return null;
  return Math.max(0, Math.round((nowMs - checkedAtMs) / 1000));
}
