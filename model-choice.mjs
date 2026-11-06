// mobile-cockpit / pwa / model-choice.mjs
//
// SPEC-DELTA-2026-10-08-model-catalog-drift-guard. Pure helpers (browser + Node)
// for "is this model id one the CLI will accept?". The account's CLI catalog
// (`cursor-agent models`) shrank from 57 entries to 5 on 2026-10-08; anything
// still holding an old id (a session record, the phone's last-used model) must
// degrade to `auto` instead of reaching `--model <dead id>` or `--model ""`.
//
// Single source of truth: the PWA imports this file; the daemon re-imports it
// from `daemon/lib/spawn-agent.mjs` (same pattern as
// `lib/refresh-signals.mjs` <- `pwa/refresh-helpers.mjs`).

"use strict";

/** Cursor's own default entry; first in MODEL_OPTIONS, so always a safe pick (decision D1). */
export const FALLBACK_MODEL_ID = "auto";

function asModelId(id) {
  return typeof id === "string" ? id.trim() : "";
}

function hasCatalog(optionIds) {
  return Array.isArray(optionIds) && optionIds.length > 0;
}

/**
 * Pick the model to actually use. An empty id always becomes the fallback
 * (never `--model ""`). A non-empty id is checked against the catalog when one
 * is known; with no catalog (config failed to load, or an empty list) nothing
 * can be judged, so a non-empty id is kept as-is.
 *
 * @param {unknown} id
 * @param {string[]|undefined} optionIds  ids from CONFIG.session.modelOptions / MODEL_OPTIONS
 * @param {string} [fallback]
 * @returns {{ model: string, substituted: boolean }}
 */
export function resolveModelChoice(id, optionIds, fallback = FALLBACK_MODEL_ID) {
  const wanted = asModelId(id);
  if (wanted === "") return { model: fallback, substituted: true };
  if (!hasCatalog(optionIds)) return { model: wanted, substituted: false };
  if (optionIds.includes(wanted)) return { model: wanted, substituted: false };
  return { model: fallback, substituted: true };
}

/**
 * Describe the extra <option> the session detail view should show for a
 * recorded model the catalog no longer offers, so the select does not render
 * blank. Null when there is nothing stale to show.
 *
 * @param {unknown} id
 * @param {string[]|undefined} optionIds
 * @returns {{ value: string, label: string, disabled: true } | null}
 */
export function describeStaleModelOption(id, optionIds) {
  const wanted = asModelId(id);
  if (wanted === "" || !hasCatalog(optionIds) || optionIds.includes(wanted)) return null;
  return { value: wanted, label: `${wanted} (unavailable — runs as Auto)`, disabled: true };
}
