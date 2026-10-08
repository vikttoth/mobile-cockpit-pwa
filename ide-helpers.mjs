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
 * True for brand-new IDE tabs mirrored without a transcript file yet.
 *
 * @param {object|null|undefined} tab
 * @returns {boolean}
 */
export function isEmptyIdeTab(tab) {
  if (!tab || typeof tab !== "object") return false;
  if (tab.openStub === true) return true;
  const title = typeof tab.title === "string" ? tab.title.trim() : "";
  const count = typeof tab.messageCount === "number" ? tab.messageCount : 0;
  return count === 0 && title === "New Agent";
}

/**
 * List-row / confirm-modal status for an IDE tab.
 *
 * @param {object|null|undefined} tab
 * @returns {string}
 */
export function ideTabStatusLabel(tab) {
  if (isEmptyIdeTab(tab)) return "Empty tab";
  return waitingOnLabel(tab && tab.waitingOn);
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
export function openTabsSourceHint(_openTabsSource) {
  // 2026-09-28: used to special-case openTabsSource === "extension" as "the
  // precise source, no hint needed" -- the extension this assumed would one
  // day exist turned out to be permanently infeasible (Cursor's real Agent
  // tabs expose no tab.input via the public VS Code Tabs API at all, see
  // SPEC-DELTA-2026-09-26-ui-cleanup-and-ide-open-tabs.md's dated
  // correction). mtime-heuristic is the only real source there ever is now,
  // so this always explains the estimate -- the parameter is kept (rather
  // than changing every call site) in case a real precise source ever
  // exists again.
  return (
    "Open tabs are estimated from recent transcript activity (most-recently-touched " +
    "conversations), not a live IDE tab bar."
  );
}

// Tracker view (SPEC-DELTA-2026-10-07-cockpit-pin-density-unified-view.md,
// same-day polish): Viktor wants Copilot vs Cowork distinguished in the
// Browser group, not lumped together as one undifferentiated "Copilot" row.
// A browser tab's `url` is only known once it has been the foreground tab at
// least once since the mirror started (see matching.mjs#mergeBrowserTabs) --
// live data confirmed a real Cowork tab's captured url contains a `/cowork`
// path segment (e.g. "m365.cloud.microsoft/cowork?auth=..."), while a tab
// whose url was never captured (still null) cannot be told apart at all.
// Best-guess heuristic, like the mirror's own COPILOT_PATTERN -- defaults
// the unknown case to the umbrella term "Copilot" (Viktor's own vocabulary:
// Cowork is Copilot's own agentic mode, not a separate product) rather than
// overclaiming a specific answer we don't have.
const COWORK_URL_PATTERN = /\/cowork(?:[/?#]|$)/i;

/**
 * @param {string|null|undefined} url
 * @returns {"cowork"|"copilot"}
 */
export function classifyCopilotKind(url) {
  if (typeof url === "string" && COWORK_URL_PATTERN.test(url)) return "cowork";
  return "copilot";
}

/**
 * Content hash of a row's current activity (FNV-1a, sync so the PWA and the
 * digest daemon compute the exact same value). A stored LLM digest is only
 * shown while its hash still matches (SPEC-DELTA-2026-10-08-tracker-hover-
 * digest.md, AC-256); the language is deliberately NOT part of it.
 */
export function activityHash(activity) {
  const s = `${activity?.ask ?? ""}\u0001${activity?.reply ?? ""}`;
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/** Stable key a row's digest is stored under (digests.json `entries`). */
export function digestRowKey(source, row) {
  if (source === "claude") {
    const id = row?.sessionId || row?.composerId;
    return id ? `claude:${id}` : null;
  }
  if (source === "cursor") {
    const id = row?.composerId || row?.sessionId;
    return id ? `cursor:${id}` : null;
  }
  if (source === "browser") return row?.title ? `browser:${row.title}` : null;
  return null;
}

/** The stored digest for `key`, or null when absent or stale vs the row's current activity. */
export function pickDigest(digestsDoc, key, activity) {
  if (!key) return null;
  const entry = digestsDoc && digestsDoc.entries ? digestsDoc.entries[key] : null;
  if (!entry || typeof entry.text !== "string" || !entry.text) return null;
  return entry.hash === activityHash(activity) ? entry.text : null;
}

/**
 * Hover-card text for a Tracker row: full (untruncated) title, the mirror's
 * longer summary when it adds something, then source / status / last
 * activity. Every row gets a card -- including Copilot rows, which have no
 * summary field at all -- so hovering any row always shows its detail.
 *
 * @param {{title?:string, summary?:string|null, sourceBadge?:string|null, statusLabel?:string|null, lastActivityAt?:string|number|null}} row
 * @param {number} nowMs
 * @returns {string} newline-separated lines
 */
export function trackerHoverText(row, nowMs) {
  const title = (row && typeof row.title === "string" && row.title.trim()) || "(untitled)";
  const lines = [title];
  // SPEC-DELTA-2026-10-08-tracker-hover-digest.md (AC-251): what it is doing
  // NOW (last reply / last ask) replaces the first-message summary, which goes
  // stale; the summary stays as the fallback for rows without activity.
  // Layer B (optional LLM digest, already matched to this row's current
  // activity by pickDigest): one crisp line right under the title.
  if (typeof row?.digest === "string" && row.digest.trim()) lines.push(row.digest.trim());
  const reply = typeof row?.activity?.reply === "string" ? row.activity.reply.trim() : "";
  const ask = typeof row?.activity?.ask === "string" ? row.activity.ask.trim() : "";
  if (reply || ask) {
    if (reply) lines.push(`Now: ${reply}`);
    if (ask) lines.push(`Asked: ${ask}`);
  } else if (typeof row?.summary === "string" && row.summary.trim() && row.summary.trim() !== title) {
    lines.push(row.summary.trim());
  }
  if (typeof row?.needsAction === "string" && row.needsAction.trim()) {
    lines.push(`Needs you: ${row.needsAction.trim()}`);
  }
  const meta = [];
  if (row?.sourceBadge) meta.push(row.sourceBadge);
  if (row?.statusLabel) meta.push(row.statusLabel);
  if (row?.lastActivityAt) meta.push(`last active ${relativeIdeTime(row.lastActivityAt, nowMs)}`);
  if (meta.length > 0) lines.push(meta.join(" · "));
  return lines.join("\n");
}
