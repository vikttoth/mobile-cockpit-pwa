// mobile-cockpit / pwa / write-helpers.mjs -- pure hash-routing helpers.
//
// v1's five write-path helpers (generateSessionId, validateCreateInputs,
// mergeAppendSession, mergeUpdateStatus, mergeQueueFollowUp) were removed
// 2026-09-26 (SPEC-DELTA-2026-09-26-ui-cleanup, item A) along with the rest
// of the v1 "Sessions" surface -- no in-flight v1 session needed migrating.
// This module SURVIVES (it is not v1-only): `parseLocationHash`/
// `formatViewHash` also drive the v2 and "Shared with me" routes, and
// `cryptoRandomBytes` is a generic browser-crypto wrapper with no v1
// coupling at all.
//
// Why a separate module:
//   - pwa/app.js is a browser-side script that binds DOMContentLoaded
//     (behind a typeof-check guard) and reaches for an auth library
//     exposed on the global namespace at boot. Importing it from a
//     Node-side unit test would still trigger the auth flow at module
//     load time, which is brittle for hermetic tests.
//   - The helpers below are PURE -- no DOM, no fetch, no auth-lib
//     coupling. Living in their own ES module lets Node tests `import`
//     them cleanly (see tests/flows/mobile-cockpit/pwa-write-helpers-unit.sh).
//   - app.js consumes them via a one-shot dynamic import() during boot
//     (it stays a classic script for back-compat with the deferred
//     auth-lib bundle that loads first).

"use strict";

// =============================================================================
// Hash routing — deep links (#v2-detail/<id>) for Teams notify + bookmarks
// =============================================================================
//
// MSAL OAuth redirects use fragments like #code=... or #error=... — never
// treat those as app routes. Ours: #v2-list, #v2-new, #v2-detail/<id>.

const MSAL_OAUTH_FRAGMENT_RE =
  /^(code|error|error_description|state|client_info|session_state)=/i;

/**
 * Parse a location hash string (e.g. "#v2-detail/<sessionId>") into a
 * route, or null if not ours.
 *
 * @returns {{ view: 'v2-list'|'v2-new'|'v2-detail', sessionId?: string } | null}
 */
export function parseLocationHash(hash) {
  if (hash == null || hash === "" || hash === "#") {
    return { view: "v2-list" };
  }
  const raw = String(hash).replace(/^#/, "").trim();
  if (!raw) return { view: "v2-list" };
  if (MSAL_OAUTH_FRAGMENT_RE.test(raw)) return null;
  if (raw === "v2-list") return { view: "v2-list" };
  if (raw === "v2-new") return { view: "v2-new" };
  const v2DetailMatch = /^v2-detail\/([^/?#]+)$/.exec(raw);
  if (v2DetailMatch) {
    const sessionId = decodeURIComponent(v2DetailMatch[1]);
    if (sessionId) return { view: "v2-detail", sessionId };
  }
  return null;
}

/**
 * Build a hash string for the given view (without leading #).
 * Returns null for views that should not update the URL (ide-* modals).
 */
export function formatViewHash(viewId, payload) {
  if (viewId === "v2-list") return "v2-list";
  if (viewId === "v2-new") return "v2-new";
  if (viewId === "v2-detail" && payload && payload.sessionId) {
    return `v2-detail/${encodeURIComponent(payload.sessionId)}`;
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
