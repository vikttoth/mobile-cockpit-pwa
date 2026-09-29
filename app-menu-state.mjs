// flows/mobile-cockpit/pwa/app-menu-state.mjs
//
// Pure helpers for the header account menu: the initials on the button and
// whether its red "something's wrong" dot shows. No DOM -- app.js loads this
// via dynamic import() and applies it. SPEC-DELTA-2026-09-29-app-menu.md.

/**
 * Initials for the account button: from the UPN's local part first
 * ("viktor.toth@nokia.com" -> "VT"; Nokia display names are "Last, First
 * (…)" and would read backwards), then the display name, then "?".
 *
 * @param {string|null|undefined} displayName
 * @param {string|null|undefined} username
 * @returns {string}
 */
export function deriveInitials(displayName, username) {
  // First + last part ("anna-maria.kiss" -> "AK"), or the single part's letter.
  const fromParts = (parts) => {
    const p = parts.filter(Boolean);
    const picked = p.length > 1 ? [p[0], p[p.length - 1]] : p;
    return picked.map((x) => x[0].toUpperCase()).join("");
  };
  if (typeof username === "string" && username.includes("@")) {
    const local = username.split("@")[0];
    const got = fromParts(local.split(/[._-]+/));
    if (got) return got;
  }
  if (typeof displayName === "string" && displayName.trim()) {
    const cleaned = displayName.replace(/\(.*?\)/g, " ").replace(/,/g, " ");
    const got = fromParts(cleaned.trim().split(/\s+/));
    if (got) return got;
  }
  return "?";
}

/**
 * Whether the account button shows its red dot, and why.
 *
 * @param {{ statusBadge?: string, health?: string, daemon?: string, daemonChecked?: boolean }} s
 *   statusBadge: #status-badge data-status; health: #health-badge data-state;
 *   daemon: #daemon-control-badge data-state; daemonChecked: false while it
 *   still reads "checking…" at boot.
 * @returns {{ alert: boolean, reasons: string[] }}
 */
export function deriveMenuAlert(s = {}) {
  const reasons = [];
  if (s.statusBadge === "error") reasons.push("sign-in / config error");
  if (s.health && s.health !== "ok" && s.health !== "unknown") reasons.push(`auth: ${s.health}`);
  if (s.daemon === "stopped" || s.daemon === "error") reasons.push(`daemon ${s.daemon}`);
  if (s.daemon === "unknown" && s.daemonChecked) reasons.push("daemon status unknown");
  return { alert: reasons.length > 0, reasons };
}
