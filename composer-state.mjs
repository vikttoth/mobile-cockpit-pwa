// flows/mobile-cockpit/pwa/composer-state.mjs
//
// Pure helpers for the v2 chat composer: which buttons show, which are
// enabled, what the header status chip says. No DOM, no MSAL -- app.js loads
// this via dynamic import() and applies the result; Node tests import it
// directly. See SPEC-DELTA-2026-09-29-composer-contextual-buttons.md.

/**
 * `readOnly` (SPEC-DELTA-2026-09-29-session-sharing-stage2): a guest holds
 * Control, another device holds the lease, or this is a guest in Read mode.
 * Buttons keep their normal visibility but are all disabled (AC-155/AC-157:
 * "visible but inactive"), never hidden.
 *
 * @param {{ status?: string, text?: string, busy?: boolean, readOnly?: boolean }} [input]
 * @returns {{
 *   send: { visible: true, enabled: boolean, mode: "send"|"queue", label: string, title: string },
 *   stop: { visible: boolean, enabled: boolean },
 *   force: { visible: boolean, enabled: boolean },
 * }}
 */
export function deriveComposerButtons({ status, text, busy = false, readOnly = false } = {}) {
  const running = status === "running";
  const hasText = typeof text === "string" && text.trim().length > 0;
  const blocked = busy || readOnly;
  return {
    // Always visible (stable anchor); Send already enqueues while a turn
    // runs, so it becomes "Queue" in the same slot instead of a second button.
    send: {
      visible: true,
      enabled: hasText && !blocked,
      mode: running ? "queue" : "send",
      label: running ? "Queue" : "Send",
      title: readOnly ? "Read-only" : running ? "Queue — runs after the current turn" : "Send",
    },
    stop: { visible: running, enabled: !blocked },
    force: { visible: running && hasText, enabled: !blocked },
  };
}

const STATUS_TONES = {
  running: "running",
  done: "ok",
  failed: "error",
  stopped: "muted",
  pending: "muted",
  provisioning: "muted",
};

/**
 * @param {{ status?: string, chatId?: string|null, queue?: unknown[] }|null} record
 * @returns {{ label: string, tone: string }|null}
 */
export function deriveStatusChip(record) {
  if (!record || typeof record !== "object" || typeof record.status !== "string" || !record.status) {
    return null;
  }
  const base = record.status === "pending" && !record.chatId ? "provisioning" : record.status;
  const queued = Array.isArray(record.queue) ? record.queue.length : 0;
  return {
    label: queued > 0 ? `${base} · ${queued} queued` : base,
    tone: STATUS_TONES[base] || "muted",
  };
}

/**
 * Ctrl+Enter / Cmd+Enter submits; plain Enter stays a newline (phone keyboards).
 *
 * @param {{ key?: string, ctrlKey?: boolean, metaKey?: boolean, isComposing?: boolean }|null} ev
 * @returns {boolean}
 */
export function isSubmitShortcut(ev) {
  if (!ev || ev.key !== "Enter" || ev.isComposing) return false;
  return Boolean(ev.ctrlKey || ev.metaKey);
}

/**
 * Auto-grow: the box's height follows its content between min and max.
 *
 * @param {number} scrollHeightPx
 * @param {{ minPx: number, maxPx: number }} bounds
 * @returns {number}
 */
export function clampComposerHeight(scrollHeightPx, { minPx, maxPx }) {
  if (!Number.isFinite(scrollHeightPx)) return minPx;
  return Math.min(maxPx, Math.max(minPx, scrollHeightPx));
}
