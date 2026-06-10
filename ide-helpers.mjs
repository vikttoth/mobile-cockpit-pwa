// flows/mobile-cockpit/pwa/ide-helpers.mjs
//
// Pure helpers for the M2.1 PWA "IDE tabs" view. These shape the OneDrive
// `cursor-cockpit/ide-tabs.json` snapshot (written by the read-only
// `flows/mobile-cockpit/ide-mirror/poll.mjs` daemon) into the structures
// that the browser-side renderer consumes.
//
// Design goals:
//   - **Pure**: no DOM, no fetch, no MSAL, no Date.now() reads. The
//     `relativeIdeTime` helper takes the reference time as an explicit
//     argument so tests can run deterministically.
//   - **Defensive**: never throw on missing/garbage input. The mirror
//     daemon already validates the envelope, but the PWA may receive an
//     in-flight write (rare; mirrored via fingerprint skip in the daemon)
//     or a malformed entry from a future schema version. Default to
//     "(untitled)" / "Idle" / "?" rather than crash the view.
//   - **Standalone**: Node-importable with `import * as H from "..."` so
//     the same module that ships to the PWA is exercised in unit tests.
//
// Coverage: tests/flows/mobile-cockpit/pwa-ide-helpers-unit.sh (this
// pairing is wired into tests/coverage-map.sh + the runner asserts).

/**
 * Sort a list of IDE tabs by `lastActivityAt` (ISO string), newest first,
 * then cap at `limit`. Tabs whose timestamp does not parse (or is missing
 * entirely) sink to the bottom but are not dropped.
 *
 * @param {Array<object>|null|undefined} tabs
 * @param {number} limit  Maximum tabs to return (defaults to a safe 100).
 * @returns {Array<object>}
 */
/**
 * Order tabs for the PWA list. When the snapshot was built from the
 * extension open-tab cache, preserve IDE tab-bar order; otherwise fall
 * back to `sortIdeTabs` (newest activity first).
 *
 * @param {Array<object>|null|undefined} tabs
 * @param {number} limit
 * @param {{openTabsSource?: string|null}} [opts]
 * @returns {Array<object>}
 */
export function orderIdeTabsForDisplay(tabs, limit, opts = {}) {
  if (!Array.isArray(tabs) || tabs.length === 0) return [];
  const cap = Number.isFinite(limit) && limit > 0 ? Math.floor(limit) : 100;
  if (opts.openTabsSource === "extension") {
    return tabs.slice(0, cap);
  }
  return sortIdeTabs(tabs, cap);
}

export function sortIdeTabs(tabs, limit) {
  if (!Array.isArray(tabs)) return [];
  const cap = Number.isFinite(limit) && limit > 0 ? Math.floor(limit) : 100;
  const decorated = tabs.map((tab) => {
    const ts = tab && tab.lastActivityAt;
    const ms = typeof ts === "string" ? Date.parse(ts) : NaN;
    return { tab, sortMs: Number.isFinite(ms) ? ms : -Infinity };
  });
  decorated.sort((a, b) => b.sortMs - a.sortMs);
  return decorated.slice(0, cap).map((d) => d.tab);
}

/**
 * Bucket tabs by `waitingOn` into `{agent, user, none, total}` for the
 * header badges. Unknown / missing values fall into the `none` bucket so
 * the totals always match.
 *
 * @param {Array<object>|null|undefined} tabs
 * @returns {{agent:number,user:number,none:number,total:number}}
 */
export function summarizeWaitingOn(tabs) {
  const out = { agent: 0, user: 0, none: 0, total: 0 };
  if (!Array.isArray(tabs)) return out;
  for (const tab of tabs) {
    out.total += 1;
    const w = tab && typeof tab.waitingOn === "string" ? tab.waitingOn : "none";
    if (w === "agent") out.agent += 1;
    else if (w === "user") out.user += 1;
    else out.none += 1;
  }
  return out;
}

/**
 * Trim a tab title for display. Falls back to "(untitled)" when missing,
 * null, or whitespace-only. Truncates with an ellipsis so the result is
 * always <= maxLen characters.
 *
 * @param {string|null|undefined} title
 * @param {number} maxLen  Cap; defaults to 80.
 * @returns {string}
 */
export function formatTabTitle(title, maxLen) {
  const cap = Number.isFinite(maxLen) && maxLen > 3 ? Math.floor(maxLen) : 80;
  const s = typeof title === "string" ? title.trim() : "";
  if (!s) return "(untitled)";
  if (s.length <= cap) return s;
  return s.slice(0, cap - 3) + "...";
}

/**
 * Human-friendly label for the `waitingOn` enum.
 *
 *   agent -> "Agent thinking"
 *   user  -> "Your turn"
 *   *     -> "Idle"
 *
 * @param {string|null|undefined} waitingOn
 * @returns {string}
 */
export function waitingOnLabel(waitingOn) {
  if (waitingOn === "agent") return "Agent thinking";
  if (waitingOn === "user") return "Your turn";
  return "Idle";
}

/**
 * Normalize a snapshot thread-entry into a renderable shape for the PWA.
 *
 * The mirror daemon's `tailThread()` has ALREADY flattened each turn into
 * `{role, text, toolCalls, hasContent}` (see
 * `flows/mobile-cockpit/ide-mirror/lib/transcripts.mjs#tailThread`), so the
 * PWA never re-parses Anthropic-style `content[]` parts. This helper just
 * defaults missing fields, attaches a human-friendly `label`, and exposes
 * `tools` (aliased from the snapshot's `toolCalls`) under a name that reads
 * naturally in the render code (`entry.tools.length`).
 *
 * @param {object|null|undefined} entry
 * @returns {{role:string, label:string, text:string, tools:Array<string>, hasContent:boolean}}
 */
export function formatThreadEntry(entry) {
  if (!entry || typeof entry !== "object") {
    return { role: "unknown", label: "System", text: "", tools: [], hasContent: false };
  }
  const role = typeof entry.role === "string" ? entry.role : "unknown";
  const label = role === "user" ? "You" : role === "assistant" ? "Agent" : "System";
  const text = typeof entry.text === "string" ? entry.text : "";
  const tools = Array.isArray(entry.toolCalls)
    ? entry.toolCalls.filter((t) => typeof t === "string")
    : [];
  const hasContent = typeof entry.hasContent === "boolean"
    ? entry.hasContent
    : (text.length > 0 || tools.length > 0);
  return { role, label, text, tools, hasContent };
}

/**
 * Render an IDE-tab activity timestamp as a relative string. Accepts an
 * ISO string OR an epoch-ms number (both come through in the snapshot:
 * `lastActivityAt` and `lastActivityAtMs` respectively).
 *
 * NOTE: `nowMs` must be supplied explicitly. Callers in the PWA pass
 * `Date.now()`; tests pass a fixed reference so assertions are
 * deterministic.
 *
 * @param {string|number|null|undefined} ts
 * @param {number} nowMs
 * @returns {string}
 */
export function relativeIdeTime(ts, nowMs) {
  let ms;
  if (typeof ts === "number" && Number.isFinite(ts)) {
    ms = ts;
  } else if (typeof ts === "string") {
    const parsed = Date.parse(ts);
    ms = Number.isFinite(parsed) ? parsed : NaN;
  } else {
    ms = NaN;
  }
  if (!Number.isFinite(ms) || !Number.isFinite(nowMs)) return "?";
  const diffSec = Math.max(0, Math.round((nowMs - ms) / 1000));
  if (diffSec < 60) return diffSec + " s ago";
  const diffMin = Math.round(diffSec / 60);
  if (diffMin < 60) return diffMin + " min ago";
  const diffHr = Math.round(diffMin / 60);
  if (diffHr < 24) return diffHr + " h ago";
  const diffDay = Math.round(diffHr / 24);
  return diffDay + " d ago";
}

/**
 * Pick the tab array for the IDE list sub-view from a v1 or v2 snapshot.
 *
 * @param {object|null|undefined} snapshot
 * @param {"open"|"history"} [mode="open"]
 * @returns {Array}
 */
export function pickIdeTabList(snapshot, mode = "open") {
  if (!snapshot || typeof snapshot !== "object") return [];
  if (mode === "history") {
    return Array.isArray(snapshot.historyTabs) ? snapshot.historyTabs : [];
  }
  if (Array.isArray(snapshot.openTabs)) return snapshot.openTabs;
  if (Array.isArray(snapshot.tabs)) return snapshot.tabs;
  return [];
}

/**
 * Find one tab by composerId across open, history, and legacy `tabs`.
 *
 * @param {object|null|undefined} snapshot
 * @param {string} composerId
 * @returns {object|null}
 */
export function findIdeTab(snapshot, composerId) {
  if (!snapshot || typeof composerId !== "string" || composerId.length === 0) {
    return null;
  }
  const want = composerId.toLowerCase();
  const merged = [
    ...(Array.isArray(snapshot.openTabs) ? snapshot.openTabs : []),
    ...(Array.isArray(snapshot.historyTabs) ? snapshot.historyTabs : []),
    ...(Array.isArray(snapshot.tabs) ? snapshot.tabs : []),
  ];
  return (
    merged.find((t) => t && String(t.composerId).toLowerCase() === want) ?? null
  );
}

/**
 * User-facing hint when open tabs are not sourced from the extension cache.
 *
 * @param {string|null|undefined} openTabsSource
 * @returns {string|null} null when no warning needed
 */
export function openTabsSourceHint(openTabsSource) {
  if (openTabsSource === "extension") return null;
  return (
    "Open tabs are estimated from recent transcript activity, not the live IDE tab bar. " +
    "Install or reload the mobile-cockpit extension in Cursor for exact sync."
  );
}
