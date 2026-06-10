// mobile-cockpit / pwa / pending-question.mjs
//
// Pure helpers for rendering Cursor AskQuestion blocks on the phone PWA:
// free-text escape detection, answer formatting, and bundle completeness.

"use strict";

/**
 * Stable key for a question object from ide-mirror pendingQuestion payload.
 * @param {{id?: string, prompt: string}} q
 */
export function questionKey(q) {
  if (!q || typeof q.prompt !== "string") return "";
  const id = typeof q.id === "string" ? q.id.trim() : "";
  return id || q.prompt.trim();
}

/**
 * Heuristic: option is the "other / free text" escape hatch (per personal-kb).
 * @param {{id?: string, label?: string}} opt
 */
export function isFreeTextEscapeOption(opt) {
  if (!opt || typeof opt !== "object") return false;
  const s = `${opt.label || ""} ${opt.id || ""}`.toLowerCase();
  return /(egyéb|egyeb|saját|sajat|sajat megfogalmaz|other|free\s*text|más megfogalmaz|mást szeretn|custom answer|type your own|write your own)/i.test(
    s,
  );
}

/**
 * @param {{options?: Array, allowMultiple?: boolean}} q
 */
export function showInlineFreeTextInput(q) {
  if (!q || !Array.isArray(q.options) || q.options.length === 0) return true;
  if (q.allowMultiple) return false;
  if (q.options.length === 1 && isFreeTextEscapeOption(q.options[0])) return true;
  return false;
}

/**
 * @param {Array} questions
 * @param {Record<string, string|string[]>} answersByKey
 */
export function formatAnswersForSend(questions, answersByKey) {
  if (!Array.isArray(questions) || questions.length === 0) return "";
  if (questions.length === 1) {
    const key = questionKey(questions[0]);
    const v = answersByKey[key];
    if (v == null) return "";
    return Array.isArray(v) ? v.filter(Boolean).join(", ") : String(v).trim();
  }
  const lines = [];
  for (const q of questions) {
    const key = questionKey(q);
    const v = answersByKey[key];
    if (v == null) continue;
    const text = Array.isArray(v)
      ? v.filter(Boolean).join(", ")
      : String(v).trim();
    if (!text) continue;
    const prefix = q.id && q.id.trim() ? q.id.trim() : q.prompt;
    lines.push(`${prefix}: ${text}`);
  }
  return lines.join("\n");
}

/**
 * @param {Array} questions
 * @param {Record<string, string|string[]>} answersByKey
 */
export function allQuestionsAnswered(questions, answersByKey) {
  if (!Array.isArray(questions) || questions.length === 0) return false;
  for (const q of questions) {
    const key = questionKey(q);
    const v = answersByKey[key];
    if (q.allowMultiple) {
      if (!Array.isArray(v) || v.length === 0) return false;
    } else if (v == null || String(v).trim() === "") {
      return false;
    }
  }
  return true;
}

/**
 * Fingerprint for invalidating in-progress answer drafts when the agent asks anew.
 * @param {Array<{id?: string, prompt: string, options?: Array}>} questions
 */
export function pendingQuestionsFingerprint(questions) {
  if (!Array.isArray(questions)) return "";
  return questions
    .map((q) => `${q.id || ""}|${q.prompt}|${(q.options || []).map((o) => o.id).join(",")}`)
    .join(";;");
}
