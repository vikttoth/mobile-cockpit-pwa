// mobile-cockpit / lib / cockpit-health-model.mjs
//
// SPEC-DELTA-2026-10-01-health-visibility-and-manual-self-heal. Pure logic
// (no I/O, no `vscode`/`node:`/DOM imports) shared byte-for-byte with
// pwa/cockpit-health-model.mjs (enforced by
// tests/flows/mobile-cockpit/pwa-cockpit-health-coherence.sh, same shape as
// pwa-daemon-control-coherence.sh).
//
// Aggregates FOUR already-existing, independently-written signals
// (health.json, daemon-status.json, share-relay-status.json, shares.json)
// into one list of {scope, id, severity, reason, likelyCauses, rootCause}
// items, at READ time only -- none of those files' own writers change.
// Two rules carried over from deriveControlDisplayState (daemon-control-
// model.mjs) and generalized here:
//   1. A stale signal is "unknown", never silently trusted as its last
//      known value -- this is what protects "the laptop is off/asleep" and
//      "the guest isn't online" from ever being misread as a cockpit fault.
//   2. A narrower-scope item whose trouble is already explained by a
//      broader one carries a `rootCause` pointer instead of its own guess
//      list, so the UI can link to the one place with an actual fix
//      instead of presenting N independent mysteries for one real problem.

"use strict";

import { deriveControlDisplayState } from "./daemon-control-model.mjs";
import { activeEntriesByGuest } from "./share-model.mjs";

/**
 * The one canonical "how to sign in again" text (SPEC-DELTA-2026-11-03-auto-
 * signin-and-guidance). Interactive sign-in only works on the laptop -- the
 * phone cannot complete it -- so this is steps, not a link (the PWA's "Sign in"
 * button asks the laptop to start it; the Desktop shortcut is the fallback). The laptop's own
 * health publisher (lib/msal-cache-health.mjs) reuses these constants so the
 * PWA's fallback text and the daemon's published remediation never drift.
 */
export const AUTH_SIGNIN_STEPS =
  "Sign-in is lost on the laptop, so it cannot send anything to this phone -- Repair will not help. " +
  "Tap \"Sign in\" next to the daemon status here: the laptop opens the Microsoft page in Edge by itself and usually " +
  "completes it without a click (otherwise approve it at the laptop, or double-click \"Cockpit sign-in\" on its Desktop). " +
  "It recovers within a minute or two.";
export const AUTH_SIGNIN_RUNNING =
  "Automatic sign-in is running on the laptop -- if a Microsoft page opened in Edge, approve it. No other action needed.";

const LIKELY_CAUSES = {
  auth: [AUTH_SIGNIN_STEPS],
  daemon: [
    "Sign-in may be lost on the laptop (it then cannot send anything) -- see Sign-in above; Repair will not help in that case.",
    "WSL itself may be unresponsive (service wedge) -- try Repair.",
    "The laptop may be off or asleep -- nothing can be done remotely until it's back.",
  ],
  "share-relay": [
    "WSL itself may be unresponsive (service wedge) -- try Repair on the daemon.",
    "The relay's own browser session may have died -- it recovers automatically within a few minutes.",
  ],
  guest: ["The guest's own device may be off or offline.", "The guest may not have tapped Connect yet."],
};

/**
 * @param {object} opts
 * @param {{state: string, remediation?: string|null, checkedAt: string}|null} opts.health
 * @param {number} opts.nowMs
 * @param {number} opts.staleAfterMs
 * @returns {{severity: 'ok'|'degraded'|'error'|'unknown', reason: string|null}}
 */
export function deriveAuthSeverity({ health, nowMs, staleAfterMs }) {
  if (!health || typeof health.checkedAt !== "string") return { severity: "unknown", reason: null };
  const checkedAtMs = Date.parse(health.checkedAt);
  if (!Number.isFinite(checkedAtMs)) return { severity: "unknown", reason: null };
  if (typeof staleAfterMs === "number" && typeof nowMs === "number" && nowMs - checkedAtMs > staleAfterMs) {
    return { severity: "unknown", reason: null };
  }
  if (health.state === "ok") return { severity: "ok", reason: null };
  if (health.state === "unknown") return { severity: "unknown", reason: null };
  if (health.state === "expiring_soon") return { severity: "degraded", reason: health.remediation ?? null };
  return { severity: "error", reason: health.remediation ?? null };
}

/**
 * @param {object} opts
 * @param {any} opts.daemonStatus
 * @param {number} opts.nowMs
 * @param {number} opts.staleAfterMs
 * @returns {{severity: 'ok'|'error'|'unknown', reason: string|null}}
 */
export function deriveDaemonSeverity({ daemonStatus, nowMs, staleAfterMs }) {
  const state = deriveControlDisplayState({ status: daemonStatus, nowMs, staleAfterMs });
  if (state === "running" || state === "stopped") return { severity: "ok", reason: null };
  if (state === "error") {
    const err = daemonStatus && daemonStatus.lastAction ? daemonStatus.lastAction.error : null;
    return { severity: "error", reason: err ?? null };
  }
  return { severity: "unknown", reason: null }; // state === "unknown"
}

/**
 * share-relay-status.json is only republished on a real change (no
 * heartbeat of its own, unlike health.json/daemon-status.json) -- so a
 * stale-but-unchanged snapshot cannot be told apart from a genuinely dead
 * relay by its own timestamp alone. Once the heartbeat-guaranteed daemon
 * signal says the whole WSL environment is unreachable, treat this one as
 * equally untrustworthy, regardless of what `poolAlive` last recorded.
 *
 * @param {object} opts
 * @param {{poolAlive?: boolean}|null} opts.relayStatus
 * @param {'ok'|'error'|'unknown'} opts.daemonSeverity
 * @returns {{severity: 'ok'|'error'|'unknown', reason: string|null}}
 */
export function deriveShareRelaySeverity({ relayStatus, daemonSeverity }) {
  if (daemonSeverity === "unknown") return { severity: "unknown", reason: null };
  if (!relayStatus) return { severity: "unknown", reason: null };
  if (relayStatus.poolAlive === false) return { severity: "error", reason: null };
  return { severity: "ok", reason: null };
}

/**
 * @param {object} opts
 * @param {string} opts.email
 * @param {{guests?: Record<string, {connected: boolean, error: string|null}>}|null} opts.relayStatus
 * @param {'ok'|'error'|'unknown'} opts.shareRelaySeverity
 * @returns {{severity: 'ok'|'degraded'|'error'|'unknown', reason: string|null, rootCauseId: 'share-relay'|null}}
 */
export function deriveGuestSeverity({ email, relayStatus, shareRelaySeverity }) {
  if (shareRelaySeverity !== "ok") {
    return { severity: shareRelaySeverity, reason: null, rootCauseId: "share-relay" };
  }
  const g = relayStatus && relayStatus.guests ? relayStatus.guests[email] : undefined;
  if (!g) return { severity: "unknown", reason: null, rootCauseId: null };
  if (g.error === "waiting for the guest to tap Connect") {
    return { severity: "degraded", reason: g.error, rootCauseId: null };
  }
  if (g.error) return { severity: "error", reason: g.error, rootCauseId: null };
  return { severity: "ok", reason: null, rootCauseId: null };
}

/**
 * Top-level aggregation: builds the full, UI-ready list of health items.
 *
 * @param {object} opts
 * @param {any} opts.health
 * @param {any} opts.daemonStatus
 * @param {any} opts.relayStatus
 * @param {any} opts.shares
 * @param {number} opts.nowMs
 * @param {number} [opts.authStaleAfterMs] default 30 min (3x health.json's own 10-min heartbeat)
 * @param {number} [opts.daemonStaleAfterMs] default 3 min (3x the watchdog's 1-min cadence)
 * @returns {Array<{scope: 'global'|'guest', id: string, label: string, severity: string, reason: string|null, likelyCauses: string[]|null, rootCause: {scope: 'global', id: string}|null}>}
 */
export function buildHealthItems(opts) {
  const {
    health,
    daemonStatus,
    relayStatus,
    shares,
    nowMs,
    authStaleAfterMs = 1_800_000,
    daemonStaleAfterMs = 180_000,
  } = opts || {};

  const items = [];

  const auth = deriveAuthSeverity({ health, nowMs, staleAfterMs: authStaleAfterMs });
  items.push({
    scope: "global",
    id: "auth",
    label: "Sign-in",
    severity: auth.severity,
    reason: auth.reason,
    likelyCauses: auth.severity !== "ok" && !auth.reason ? LIKELY_CAUSES.auth : null,
    rootCause: null,
  });

  const daemon = deriveDaemonSeverity({ daemonStatus, nowMs, staleAfterMs: daemonStaleAfterMs });
  items.push({
    scope: "global",
    id: "daemon",
    label: "Daemon",
    severity: daemon.severity,
    reason: daemon.reason,
    likelyCauses: daemon.severity !== "ok" && !daemon.reason ? LIKELY_CAUSES.daemon : null,
    rootCause: null,
  });

  const shareRelay = deriveShareRelaySeverity({ relayStatus, daemonSeverity: daemon.severity });
  items.push({
    scope: "global",
    id: "share-relay",
    label: "Sharing relay",
    severity: shareRelay.severity,
    reason: shareRelay.reason,
    likelyCauses: shareRelay.severity !== "ok" && daemon.severity !== "unknown" ? LIKELY_CAUSES["share-relay"] : null,
    rootCause: shareRelay.severity !== "ok" && daemon.severity === "unknown" ? { scope: "global", id: "daemon" } : null,
  });

  const activeByGuest = activeEntriesByGuest(shares);
  for (const email of Object.keys(activeByGuest).sort()) {
    const g = deriveGuestSeverity({ email, relayStatus, shareRelaySeverity: shareRelay.severity });
    items.push({
      scope: "guest",
      id: `relay:${email}`,
      label: email,
      severity: g.severity,
      reason: g.reason,
      likelyCauses: g.severity !== "ok" && !g.reason && !g.rootCauseId ? LIKELY_CAUSES.guest : null,
      rootCause: g.rootCauseId ? { scope: "global", id: g.rootCauseId } : null,
    });
  }

  return items;
}
