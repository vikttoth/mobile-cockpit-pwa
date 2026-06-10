// mobile-cockpit / pwa / refresh-helpers.mjs
//
// Pure helpers for the OneDrive `cursor-cockpit/refresh-signals.json` nudge
// file. The PWA writes a timestamp when the user taps ↻; desktop daemons
// (`daemon/poll.mjs`, `ide-mirror/poll.mjs`) poll this file and run an
// immediate mirror tick when they see a newer request.
//
// Node daemons re-import this module from `lib/refresh-signals.mjs` — keep
// this file the single source of truth for the JSON shape.

"use strict";

export const REFRESH_SIGNALS_SCHEMA_VERSION = 1;

/** @typedef {'sessions' | 'ideTabs' | 'both'} RefreshScope */

/**
 * @param {unknown} raw
 * @returns {{ schemaVersion: number, sessionsRequestedAt: string|null, ideTabsRequestedAt: string|null }}
 */
export function parseRefreshSignals(raw) {
  const empty = emptyRefreshSignals();
  if (!raw || typeof raw !== "object") return empty;
  const o = /** @type {Record<string, unknown>} */ (raw);
  return {
    schemaVersion:
      typeof o.schemaVersion === "number" ? o.schemaVersion : REFRESH_SIGNALS_SCHEMA_VERSION,
    sessionsRequestedAt: isoOrNull(o.sessionsRequestedAt),
    ideTabsRequestedAt: isoOrNull(o.ideTabsRequestedAt),
  };
}

export function emptyRefreshSignals() {
  return {
    schemaVersion: REFRESH_SIGNALS_SCHEMA_VERSION,
    sessionsRequestedAt: null,
    ideTabsRequestedAt: null,
  };
}

/**
 * Merge a refresh nudge into an existing document (or empty).
 *
 * @param {unknown} existing
 * @param {RefreshScope} scope
 * @param {string} nowIso
 */
export function applyNudge(existing, scope, nowIso) {
  const base = parseRefreshSignals(existing);
  const out = { ...base };
  if (scope === "sessions" || scope === "both") {
    out.sessionsRequestedAt = nowIso;
  }
  if (scope === "ideTabs" || scope === "both") {
    out.ideTabsRequestedAt = nowIso;
  }
  return out;
}

/**
 * True when `doc` has a scope timestamp newer than `lastProcessedIso`.
 *
 * @param {ReturnType<typeof parseRefreshSignals>} doc
 * @param {'sessions' | 'ideTabs'} scope
 * @param {string|null|undefined} lastProcessedIso
 */
export function isNudgePending(doc, scope, lastProcessedIso) {
  const field = scope === "sessions" ? "sessionsRequestedAt" : "ideTabsRequestedAt";
  const at = doc[field];
  if (!at || typeof at !== "string") return false;
  if (!lastProcessedIso) return true;
  return at > lastProcessedIso;
}

/**
 * Pick the newest ISO timestamp from a session list (for refresh wait).
 *
 * @param {Array<{ lastUpdated?: string, createdAt?: string, created?: string }>} sessions
 */
export function maxSessionStamp(sessions) {
  let max = "";
  for (const s of sessions || []) {
    for (const k of [s.lastUpdated, s.createdAt, s.created]) {
      if (typeof k === "string" && k > max) max = k;
    }
  }
  return max || null;
}

function isoOrNull(v) {
  return typeof v === "string" && v.length > 0 ? v : null;
}
