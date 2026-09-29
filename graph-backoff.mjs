// flows/mobile-cockpit/pwa/graph-backoff.mjs
//
// PWA-side Graph back-off after a 429/503 (SPEC-DELTA-2026-09-29-pwa-polling-
// hygiene). Same Retry-After parsing as lib/graph-state.mjs (the PWA is
// published standalone and cannot import lib/; a test pins both to the same
// results). No DOM -- app.js#graphFetch consults it.

/**
 * @param {string|null|undefined} headerValue  Retry-After header (seconds)
 * @param {string} [body]  OneDrive puts error.retryAfterSeconds in the JSON body
 * @param {{ defaultMs?: number, maxMs?: number }} [opts]
 * @returns {number} ms, clamped to [1 s, maxMs]
 */
export function parseRetryAfterMs(headerValue, body, opts = {}) {
  const defaultMs = Number.isFinite(opts.defaultMs) ? opts.defaultMs : 30_000;
  const maxMs = Number.isFinite(opts.maxMs) ? opts.maxMs : 300_000;
  let sec = Number.parseFloat(headerValue ?? "");
  if (!Number.isFinite(sec) && typeof body === "string" && body) {
    try {
      sec = Number(JSON.parse(body)?.error?.retryAfterSeconds);
    } catch {
      /* not JSON */
    }
  }
  const ms = Number.isFinite(sec) && sec > 0 ? sec * 1000 : defaultMs;
  return Math.min(maxMs, Math.max(1000, ms));
}

/**
 * @param {{ now?: () => number }} [opts]
 * @returns {{ note: (status: number, retryAfterHeader: string|null, body: string) => void,
 *   remainingMs: () => number }}
 */
export function createGraphBackoff(opts = {}) {
  const now = opts.now ?? Date.now;
  let untilMs = 0;
  return {
    note(status, retryAfterHeader, body) {
      if (status !== 429 && status !== 503) return;
      untilMs = Math.max(untilMs, now() + parseRetryAfterMs(retryAfterHeader, body));
    },
    remainingMs() {
      return Math.max(0, untilMs - now());
    },
  };
}
