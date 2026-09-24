// mobile-cockpit / pwa / scrollback-helpers.mjs
//
// SPEC task 3 (flows/mobile-cockpit/SPEC.md, S-003, AC-019). Pure helpers
// behind the chat-scrollback view. Task 5 wires these into the actual DOM
// (session detail on both hosts); this module holds only the decision
// logic, so it is Node-importable and hermetically unit-tested the same
// way pwa/write-helpers.mjs and pwa/ide-helpers.mjs already are -- no DOM,
// no fetch, no MSAL, no Date.now() reads.
//
// AC-019: "While a session detail is open, the UI shall render messages[]
// oldest-to-newest with the composer fixed at the bottom, and shall
// preserve scroll position when new messages arrive above the fold."

"use strict";

const ROLE_LABELS = Object.freeze({
  user: "You",
  assistant: "Agent",
  system: "System",
});

/**
 * Order a session's messages[] oldest-to-newest for display, dropping any
 * garbage (non-object) entries rather than throwing -- a malformed entry in
 * one message must not blank the whole scrollback. Defensive on
 * null/undefined input (mirrors ide-helpers.mjs#sortIdeTabs).
 *
 * Messages already arrive oldest-first from transcript-model.mjs#appendMessage
 * (each append pushes to the end), so this is intentionally NOT a sort by
 * `ts` -- preserving insertion order is correct even if two messages share a
 * timestamp, and a stable identity pass here keeps the contract explicit
 * and future-proof against a caller that hands in an already-ordered array
 * from a different source.
 *
 * @param {Array|null|undefined} messages
 * @returns {Array}
 */
export function orderMessagesForDisplay(messages) {
  if (!Array.isArray(messages)) return [];
  return messages.filter((m) => m && typeof m === "object" && !Array.isArray(m));
}

/**
 * Decide whether an about-to-be-appended message should pull the view's
 * scroll position down to the new bottom. Takes the scroll geometry
 * BEFORE the append (standard chat-UI pattern: check first, append DOM,
 * then either scroll to bottom or leave scrollTop untouched).
 *
 * Fails OPEN (returns true) on missing/garbage geometry -- a transient
 * layout read failure should not silently strand the reader mid-scroll on
 * their very first render, and the common case (a fresh view before any
 * layout pass has happened) legitimately reports zeroes for everything.
 *
 * @param {{scrollTop?: number, scrollHeight?: number, clientHeight?: number}|null|undefined} geometry
 * @param {{thresholdPx?: number}} [opts]  how many px of slack counts as
 *   "already at the bottom" (default 48 -- enough to absorb sub-pixel
 *   rounding and a half-visible last line, not enough to falsely follow
 *   when the reader is genuinely scrolled up).
 * @returns {boolean}
 */
export function shouldAutoScrollToBottom(geometry, opts = {}) {
  if (!geometry || typeof geometry !== "object") return true;
  const { scrollTop, scrollHeight, clientHeight } = geometry;
  if (
    !Number.isFinite(scrollTop) ||
    !Number.isFinite(scrollHeight) ||
    !Number.isFinite(clientHeight)
  ) {
    return true;
  }
  const thresholdPx = Number.isFinite(opts.thresholdPx) ? opts.thresholdPx : 48;
  const distanceFromBottom = scrollHeight - (scrollTop + clientHeight);
  return distanceFromBottom <= thresholdPx;
}

/**
 * Normalize one message into a renderable shape. Mirrors
 * ide-helpers.mjs#formatThreadEntry's contract (role/label/text) so the
 * eventual chat-UI rendering code can share a mental model across the
 * session view and the read-only IDE-tab thread view.
 *
 * @param {{role?: string, text?: string, ts?: number}|null|undefined} message
 * @returns {{role: string, label: string, text: string, ts: number|null}}
 */
export function formatSessionMessage(message) {
  if (!message || typeof message !== "object") {
    return { role: "unknown", label: "System", text: "", ts: null };
  }
  const role = typeof message.role === "string" ? message.role : "unknown";
  const label = ROLE_LABELS[role] || "System";
  const text = typeof message.text === "string" ? message.text : "";
  const ts = typeof message.ts === "number" && Number.isFinite(message.ts) ? message.ts : null;
  return { role, label, text, ts };
}
