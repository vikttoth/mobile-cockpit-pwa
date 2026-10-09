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

// How long after a successful launch the Start row keeps saying "Starting…": the
// next browser-tabs snapshot (poll + OneDrive + PWA refresh) lags the real Edge.
export const COPILOT_EDGE_LAUNCH_GRACE_MS = 90_000;

/**
 * SPEC-DELTA-2026-10-08-copilot-edge-start-button.md (AC-283/AC-284): the one
 * "Playwright Edge is not running" row of the Tracker's "Tracked (Playwright
 * Edge)" sub-group. null unless the snapshot says the watch Edge is DOWN --
 * a snapshot without `copilotWatch` (older daemon) is "unknown", never "down".
 * @param {{up?: boolean}|null|undefined} copilotWatch  the snapshot's field
 * @param {{state: "idle"|"starting"|"launched"|"failed", at?: number}} start  the PWA's own start state
 * @param {number} nowMs
 * @returns {null|{title: string, actionLabel: string|null, busy: boolean}}
 */
export function copilotWatchStartRow(copilotWatch, start, nowMs) {
  if (!copilotWatch || copilotWatch.up !== false) return null;
  const state = start?.state;
  if (state === "starting") return { title: "Starting…", actionLabel: null, busy: true };
  if (state === "launched" && Number.isFinite(start.at) && nowMs - start.at < COPILOT_EDGE_LAUNCH_GRACE_MS) {
    return { title: "Starting…", actionLabel: null, busy: true };
  }
  if (state === "failed") return { title: "Start failed — tap to retry", actionLabel: "Start", busy: false };
  return { title: "Playwright Edge is not running", actionLabel: "Start", busy: false };
}

/**
 * The PWA's start state after the daemon answered (or did not) a Start tap.
 * @param {{status?: string}|null|undefined} answer parsed response body, null when there was none
 * @param {number} nowMs
 * @returns {{state: "launched", at: number}|{state: "failed"}}
 */
export function copilotEdgeStartOutcome(answer, nowMs) {
  const status = answer && answer.status;
  return status === "started" || status === "already_running" ? { state: "launched", at: nowMs } : { state: "failed" };
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
  if (row?.where) meta.push(row.where);
  if (row?.statusLabel) meta.push(row.statusLabel);
  if (row?.lastActivityAt) meta.push(`last active ${relativeIdeTime(row.lastActivityAt, nowMs)}`);
  if (meta.length > 0) lines.push(meta.join(" · "));
  return lines.join("\n");
}

const HOVER_MAX_BULLETS = 4;
const HOVER_MAX_BULLET_CHARS = 160;

/** Splits free text into short bullets: lines, then sentences; markers stripped; capped. */
export function hoverBullets(text) {
  if (typeof text !== "string") return [];
  const parts = [];
  for (const line of text.split(/\r?\n/)) {
    const clean = line.replace(/^\s*(?:[-*\u2022]|\d+[.)])\s+/, "").trim();
    if (!clean) continue;
    for (const sentence of clean.split(/(?<=[.!?])\s+/)) {
      const t = sentence.trim();
      if (t) parts.push(t);
    }
  }
  const clamp = (t) => (t.length > HOVER_MAX_BULLET_CHARS ? `${t.slice(0, HOVER_MAX_BULLET_CHARS - 1).trimEnd()}\u2026` : t);
  const out = parts.slice(0, HOVER_MAX_BULLETS).map(clamp);
  if (parts.length > HOVER_MAX_BULLETS && !out[out.length - 1].endsWith("\u2026")) out[out.length - 1] += "\u2026";
  return out;
}

/**
 * Structured model for the Tracker hover card (SPEC-DELTA-2026-10-09-tracker-hover-structured.md):
 * what to see first (needs / TL;DR), then labelled bullet sections, footer last.
 * Content is passed through untranslated.
 * @returns {{title:string, status:{label:string,kind:string|null}|null, needs:string|null, tldr:string|null, sections:{label:string,bullets:string[]}[], meta:string}}
 */
export function trackerHoverModel(row, nowMs) {
  const str = (v) => (typeof v === "string" ? v.trim() : "");
  const title = str(row?.title) || "(untitled)";
  const reply = str(row?.activity?.reply);
  const ask = str(row?.activity?.ask);
  const summary = str(row?.summary);
  let tldr = str(row?.digest) || null;
  if (!tldr && !reply && !ask && summary && summary !== title) tldr = summary;
  const sections = [];
  const nowBullets = hoverBullets(reply);
  const askBullets = hoverBullets(ask);
  if (nowBullets.length) sections.push({ label: "Now", bullets: nowBullets });
  if (askBullets.length) sections.push({ label: "Asked", bullets: askBullets });
  const meta = [];
  if (row?.sourceBadge) meta.push(row.sourceBadge);
  if (row?.where) meta.push(row.where);
  if (row?.lastActivityAt) meta.push(`last active ${relativeIdeTime(row.lastActivityAt, nowMs)}`);
  return {
    title,
    status: row?.statusLabel ? { label: row.statusLabel, kind: row.statusKind || null } : null,
    needs: str(row?.needsAction) || null,
    tldr,
    sections,
    meta: meta.join(" \u00b7 "),
  };
}

// ---------------------------------------------------------------------------
// Tracker collapsible groups (SPEC-DELTA-2026-10-09-tracker-collapsible-groups.md,
// AC-288..AC-292). The state is a plain array of "closed" keys.
// ---------------------------------------------------------------------------

export const TRACKER_COLLAPSED_STORAGE_KEY = "cockpit.tracker.collapsed";

/** `claude` for a source header, `claude/Routines` for one of its sub-groups. */
export function trackerCollapseKey(source, subLabel) {
  return subLabel ? `${source}/${subLabel}` : source;
}

/** The stored closed-set; anything unparsable or of the wrong shape means "all open". */
export function parseTrackerCollapsed(raw) {
  if (typeof raw !== "string" || !raw) return [];
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return [];
  }
  if (!Array.isArray(parsed)) return [];
  return [...new Set(parsed.filter((k) => typeof k === "string" && k))];
}

export function toggleTrackerCollapsed(collapsed, key) {
  return collapsed.includes(key) ? collapsed.filter((k) => k !== key) : [...collapsed, key];
}

const TRACKER_SIGNAL_ORDER = ["user", "problem", "agent"];

/** Most urgent status among hidden rows: waiting-on-you > problem > running; null if none. */
export function trackerCollapsedSignal(rows) {
  if (!Array.isArray(rows)) return null;
  for (const kind of TRACKER_SIGNAL_ORDER) if (rows.some((r) => r && r.statusKind === kind)) return kind;
  return null;
}

/** "Collapse all" while any rendered group is open, "Expand all" once every one is closed. */
export function trackerCollapseAllLabel(allKeys, collapsed) {
  return allKeys.length > 0 && allKeys.every((k) => collapsed.includes(k)) ? "Expand all" : "Collapse all";
}

/** The closed-set after pressing the button; keys that are not on screen are left alone. */
export function trackerCollapseAll(allKeys, collapsed) {
  if (trackerCollapseAllLabel(allKeys, collapsed) === "Expand all") return collapsed.filter((k) => !allKeys.includes(k));
  return [...new Set([...collapsed, ...allKeys])];
}

// ---------------------------------------------------------------------------
// Tracker attention (SPEC-DELTA-2026-10-09-tracker-attention-auto-open.md,
// AC-305..AC-308): a closed group that hides something needing Viktor says so
// in words, and opens by itself when a row NEWLY needs him.
// ---------------------------------------------------------------------------

export const TRACKER_SEEN_STORAGE_KEY = "cockpit.tracker.attentionSeen";

/** Rows that need Viktor: waiting on him ("user") or a problem. */
export function trackerAttentionCounts(rows) {
  const counts = { user: 0, problem: 0 };
  if (!Array.isArray(rows)) return counts;
  for (const r of rows) {
    if (r && r.statusKind === "user") counts.user += 1;
    else if (r && r.statusKind === "problem") counts.problem += 1;
  }
  return counts;
}

/** Badge text for a closed header: "2 need you", "1 problem", "2 need you · 1 problem"; "" when none. */
export function trackerAttentionBadge(rows) {
  const c = trackerAttentionCounts(rows);
  const parts = [];
  if (c.user) parts.push(`${c.user} need you`);
  if (c.problem) parts.push(`${c.problem} problem${c.problem === 1 ? "" : "s"}`);
  return parts.join(" \u00b7 ");
}

/** Stable identity of a row across polls; null when it has nothing to hold on to. */
export function trackerAttentionRowKey(source, row) {
  const id = row && (row.sessionId || row.composerId || row.link || row.title);
  return id ? `${source}:${id}` : null;
}

/** Keys of the rows that currently need Viktor. */
export function trackerAttentionKeys(source, rows) {
  if (!Array.isArray(rows)) return [];
  const out = [];
  for (const r of rows) {
    if (!r || (r.statusKind !== "user" && r.statusKind !== "problem")) continue;
    const key = trackerAttentionRowKey(source, r);
    if (key) out.push(key);
  }
  return out;
}

/** The stored seen-set; anything unparsable means "never seen" (null = no baseline yet). */
export function parseTrackerSeen(raw) {
  if (typeof raw !== "string" || !raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((k) => typeof k === "string") : null;
  } catch {
    return null;
  }
}

/**
 * Opens every closed group that holds a row which needs Viktor and was not in the previous
 * render's attention set. No baseline (`seen === null`) opens nothing -- the very first render
 * must not undo his collapsing. Returns the new closed-set, the keys that opened, and the new baseline.
 * @param {string[]} collapsed
 * @param {Record<string, string[]>} attentionByGroup collapse key -> attention row keys inside it
 * @param {string[]|null} seen
 */
export function trackerOpenForAttention(collapsed, attentionByGroup, seen) {
  const current = new Set(Object.values(attentionByGroup || {}).flat());
  if (seen === null || seen === undefined) return { collapsed, opened: [], seen: [...current] };
  const before = new Set(seen);
  const opened = [];
  for (const [groupKey, keys] of Object.entries(attentionByGroup || {})) {
    if (collapsed.includes(groupKey) && keys.some((k) => !before.has(k))) opened.push(groupKey);
  }
  return { collapsed: collapsed.filter((k) => !opened.includes(k)), opened, seen: [...current] };
}

/**
 * True for a deep link that the OS (not the browser) handles: `claude://...`, `cursor://...`.
 * Web (http/https), blob/data/javascript and relative links are not -- those still open in a new tab.
 */
export function isCustomProtocolLink(url) {
  if (typeof url !== "string") return false;
  const m = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(url.trim());
  if (!m) return false;
  return !["http", "https", "about", "blob", "data", "file", "javascript", "mailto", "tel"].includes(m[1].toLowerCase());
}

// ---------------------------------------------------------------------------
// Tracker status of an IDE (Cursor GUI) tab (SPEC-DELTA-2026-10-09-tracker-ide-tab-status.md,
// AC-310..AC-313). The mirror's `waitingOn` only reads the transcript's shape, with no notion of
// time: "user" = the assistant wrote the last message (so EVERY finished chat is "your turn"),
// "agent" = no turn_ended after the last user message (so a stopped or abandoned run is
// "running" forever). The Tracker uses the same rules as the Claude rows instead.
// ---------------------------------------------------------------------------

export const TRACKER_IDE_RUNNING_STALE_MS = 30 * 60 * 1000;

/**
 * @param {{waitingOn?: string, pendingQuestion?: object|null, lastActivityAt?: string|number|null, title?: string, messageCount?: number, openStub?: boolean}} tab
 * @param {number} nowMs
 * @returns {{statusKind: "agent"|"user"|"none", statusLabel: string}}
 */
export function trackerIdeTabStatus(tab, nowMs) {
  if (isEmptyIdeTab(tab)) return { statusKind: "none", statusLabel: "Empty tab" };
  if (tab && tab.waitingOn === "user" && tab.pendingQuestion) return { statusKind: "user", statusLabel: "waiting on you" };
  const last = Date.parse(tab && tab.lastActivityAt);
  if (tab && tab.waitingOn === "agent" && Number.isFinite(last) && nowMs - last < TRACKER_IDE_RUNNING_STALE_MS) {
    return { statusKind: "agent", statusLabel: "running" };
  }
  return { statusKind: "none", statusLabel: "done" };
}
