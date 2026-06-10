// flows/mobile-cockpit/pwa/compose-draft.mjs
//
// Pure helpers for preserving per-tab compose textarea content across
// background IDE-tab snapshot refreshes (~20 s poll).

/**
 * @param {string|null|undefined} composerId
 * @returns {boolean}
 */
export function isValidComposeDraftKey(composerId) {
  return typeof composerId === "string" && composerId.trim().length > 0;
}

/**
 * Whether a detail re-render should keep the compose textarea untouched.
 *
 * @param {object} opts
 * @param {boolean} [opts.preserveCompose]  explicit flag from auto-refresh
 * @param {boolean} [opts.textareaFocused]
 * @param {string} [opts.textareaValue]
 * @param {string} [opts.draftText]  in-memory draft for this composer
 */
export function shouldPreserveComposeOnRefresh(opts = {}) {
  if (opts.preserveCompose === true) return true;
  if (opts.textareaFocused === true) return true;
  const v = typeof opts.textareaValue === "string" ? opts.textareaValue : "";
  if (v.length > 0) return true;
  const d = typeof opts.draftText === "string" ? opts.draftText : "";
  return d.length > 0;
}

/**
 * @param {string} text
 * @param {number} maxLen
 * @returns {{ counterText: string, over: boolean, sendDisabled: boolean }}
 */
export function composeUiStateFromText(text, maxLen) {
  const len = typeof text === "string" ? text.length : 0;
  const max = Number.isFinite(maxLen) && maxLen > 0 ? maxLen : 20000;
  const over = len > max;
  return {
    counterText: `${len} / ${max}`,
    over,
    sendDisabled: len === 0 || over,
  };
}

/**
 * Whether the AskQuestion block should survive a background snapshot refresh
 * without tearing down its DOM (prevents focus loss + broken taps on mobile).
 *
 * @param {object} opts
 * @param {boolean} [opts.actionInFlight]
 * @param {boolean} [opts.activeElementInPending]
 * @param {boolean} [opts.hasDraftAnswers]
 * @param {boolean} [opts.pendingVisible]
 */
export function shouldPreservePendingQuestionOnRefresh(opts = {}) {
  if (opts.actionInFlight === true) return true;
  if (opts.activeElementInPending === true) return true;
  if (opts.hasDraftAnswers === true) return true;
  return false;
}
