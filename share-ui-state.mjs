// flows/mobile-cockpit/pwa/share-ui-state.mjs
//
// SPEC-DELTA-2026-09-29-session-sharing-stage2: pure view-state helpers for the
// sharing UI -- the host Share panel rows, the host's read-only state on a
// chat (a guest holds Control, or another device holds the lease), the guest
// view's permissions, and guest/host detection. No DOM, no MSAL; app.js loads
// it via import() and Node tests import it directly. All user-facing strings
// are English (standing rule for the cockpit UI).

export const MODE_LABELS = Object.freeze({ off: "Off", read: "Read", control: "Control", concurrent: "Concurrent" });

/** "karoly.brix@nokia.com" -> "Karoly Brix" (best effort; the email when it can't). */
export function nameFromEmail(email) {
  if (typeof email !== "string" || !email.includes("@")) return email || "";
  const local = email.split("@")[0];
  const parts = local.split(/[._-]+/).filter(Boolean);
  if (!parts.length) return email;
  return parts.map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join(" ");
}

/** True when the signed-in account is NOT the cockpit's owner (AC-150). */
export function isGuestAccount(accountUpn, hostUpn) {
  if (typeof accountUpn !== "string" || !accountUpn || typeof hostUpn !== "string" || !hostUpn) return false;
  return accountUpn.trim().toLowerCase() !== hostUpn.trim().toLowerCase();
}

/**
 * Rows for one item's Share panel.
 * @param {{shares: object|null, kind: string, id: string, relayStatus?: object|null}} input
 */
export function deriveShareRows({ shares, kind, id, relayStatus }) {
  const items = Array.isArray(shares?.items) ? shares.items : [];
  const item = items.find((it) => it.kind === kind && it.id === id);
  const modes = kind === "session" ? ["off", "read", "control", "concurrent"] : ["off", "read"];
  if (!item) return [];
  return item.guests.map((g) => {
    const st = relayStatus?.guests?.[g.email] || null;
    let connection = "unknown";
    if (st) connection = st.connected ? "connected" : "waiting";
    return {
      email: g.email,
      name: nameFromEmail(g.email),
      mode: g.mode,
      modes: modes.map((m) => ({ id: m, label: MODE_LABELS[m], selected: m === g.mode })),
      isHolder: item.control?.holder === g.email,
      connection,
      connectionLabel:
        connection === "connected" ? "Connected" : connection === "waiting" ? "Waiting for Connect" : "",
    };
  });
}

/** Global "Stop all sharing" switch state for the panel header. */
export function deriveStopAllState(shares) {
  const enabled = shares ? shares.sharingEnabled !== false : true;
  return { stopped: !enabled, label: enabled ? "Stop all sharing" : "Sharing is stopped — resume" };
}

/**
 * Is the HOST's own chat view read-only right now, and why?
 *   - a guest holds Control of it (AC-157) -> Take back control
 *   - another device holds the lease, e.g. the laptop via mc (task 15, AC-163) -> Take over
 */
export function deriveHostChatAccess({ shares, record }) {
  const items = Array.isArray(shares?.items) ? shares.items : [];
  const enabled = shares ? shares.sharingEnabled !== false : true;
  const item = record ? items.find((it) => it.kind === "session" && it.id === record.id) : null;
  const holder = enabled ? item?.control?.holder || null : null;
  if (holder) {
    return {
      readOnly: true,
      reason: "guest-control",
      banner: `${nameFromEmail(holder)} has control of this chat.`,
      action: { id: "take-back", label: "Take back control" },
    };
  }
  const owner = record?.owner;
  if (owner && owner !== "daemon") {
    return {
      readOnly: true,
      reason: "leased",
      banner: owner === "laptop" ? "Controlled by the laptop (mc)." : `Controlled by ${owner}.`,
      action: { id: "take-over", label: "Take over" },
    };
  }
  return { readOnly: false, reason: null, banner: null, action: null };
}

/** The take-over row in the settings sheet (task 15). */
export function deriveTakeOverRow(record) {
  const owner = record?.owner || "daemon";
  if (owner === "daemon") return { text: "Runs on the cockpit daemon.", action: null };
  return {
    text: owner === "laptop" ? "Controlled by the laptop (mc)." : `Controlled by ${owner}.`,
    action: { id: "take-over", label: "Take over" },
  };
}

/**
 * What a GUEST may do in a mirrored chat (AC-155): write controls only in
 * Control (as holder) or Concurrent; everything stays visible but disabled
 * otherwise.
 */
export function deriveGuestChatAccess(projection, guestEmail) {
  const mode = projection?.yourMode || "off";
  const me = typeof guestEmail === "string" ? guestEmail.toLowerCase() : "";
  const canWrite = mode === "concurrent" || (mode === "control" && (projection?.controlHolder || "") === me);
  let banner = null;
  if (mode === "read") banner = "Read-only — you can follow this chat live.";
  if (mode === "control" && canWrite) banner = "You have control of this chat.";
  if (mode === "concurrent") banner = "Shared chat — you and the owner can both write.";
  return { mode, canWrite, banner };
}
