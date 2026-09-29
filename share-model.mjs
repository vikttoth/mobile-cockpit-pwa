// flows/mobile-cockpit/lib/share-model.mjs
//
// Session sharing Stage 2 (SPEC-DELTA-2026-09-29-session-sharing-stage2.md):
// the pure half. No I/O, no Date.now() reads, no node: imports -- this file
// is copied byte-for-byte to pwa/share-model.mjs (coherence-tested), the same
// pattern as transcript-model.mjs.
//
// Three documents meet here:
//   - cursor-cockpit/shares.json in the HOST drive: the single source of truth
//     for who sees what, in which mode (written by the host PWA / REST API,
//     read by the share relay).
//   - items/<kind>-<id>.json + manifest.json in the GUEST drive: what the
//     relay mirrors to a guest (projectSessionForGuest / projectTabForGuest /
//     buildManifest).
//   - outbox/<sessionId>.json in the GUEST drive: ops a guest wants applied
//     (validateOp / isOpAllowed / applyGuestOp / selectNewOps).

import {
  enqueueMessage,
  requestForce,
  requestStop,
  setModel,
  setMode,
  removeQueuedMessage,
  appendMessage,
  deriveTitle,
} from "./transcript-model.mjs";

export const SHARE_KINDS = Object.freeze(["session", "cursor-tab", "claude-tab"]);
export const SESSION_MODES = Object.freeze(["off", "read", "control", "concurrent"]);
export const TAB_MODES = Object.freeze(["off", "read"]);
export const GUEST_OP_TYPES = Object.freeze(["enqueue", "force", "stop", "setModel", "setMode", "removeQueued"]);
export const WRITE_MODES = Object.freeze(["control", "concurrent"]);
export const GUEST_ROOT_FOLDER = "cursor-cockpit-guest";
export const PERSONAL_SITE_ORIGIN = "https://nokia-my.sharepoint.com";

function isoNow(nowMs) {
  return new Date(nowMs).toISOString();
}

function requireNow(now, fn) {
  if (!Number.isFinite(now)) throw new Error(`${fn}: now must be a finite epoch-ms number`);
}

function codedError(message, code) {
  const err = new Error(message);
  err.code = code;
  return err;
}

/** Lower-cased, trimmed; throws BAD_EMAIL on anything that is not a plausible address. */
export function normalizeEmail(email) {
  const e = typeof email === "string" ? email.trim().toLowerCase() : "";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw codedError(`not an email address: ${email}`, "BAD_EMAIL");
  return e;
}

/** "karoly.brix@nokia.com" -> "karoly_brix_nokia_com" (the OneDrive personal-site slug). */
export function upnSlug(upn) {
  return normalizeEmail(upn).replace(/[.@]/g, "_");
}

/** Personal site URL of a Nokia user, e.g. https://nokia-my.sharepoint.com/personal/karoly_brix_nokia_com */
export function personalSiteUrl(upn) {
  return `${PERSONAL_SITE_ORIGIN}/personal/${upnSlug(upn)}`;
}

/**
 * Where everything for one host lives inside a guest's drive. Graph-style
 * relative path (guest PWA) and SharePoint server-relative path (relay).
 */
export function guestHostFolder({ guestUpn, hostUpn }) {
  const rel = `${GUEST_ROOT_FOLDER}/hosts/${upnSlug(hostUpn)}`;
  return {
    driveRelativePath: rel,
    siteUrl: personalSiteUrl(guestUpn),
    serverRelativePath: `/personal/${upnSlug(guestUpn)}/Documents/${rel}`,
  };
}

export function itemFileName(kind, id) {
  if (!SHARE_KINDS.includes(kind)) throw codedError(`unknown share kind: ${kind}`, "BAD_KIND");
  return `${kind}-${String(id).replace(/[^A-Za-z0-9._-]/g, "_")}.json`;
}

export function outboxFileName(sessionId) {
  return `${String(sessionId).replace(/[^A-Za-z0-9._-]/g, "_")}.json`;
}

export function modesForKind(kind) {
  if (!SHARE_KINDS.includes(kind)) throw codedError(`unknown share kind: ${kind}`, "BAD_KIND");
  return kind === "session" ? SESSION_MODES : TAB_MODES;
}

export function emptyShares(now) {
  requireNow(now, "emptyShares");
  return { schemaVersion: 1, sharingEnabled: true, updatedAt: isoNow(now), items: [] };
}

function normalizeShares(shares) {
  if (!shares || typeof shares !== "object" || !Array.isArray(shares.items)) {
    return { schemaVersion: 1, sharingEnabled: true, updatedAt: null, items: [] };
  }
  return shares;
}

export function findItem(shares, kind, id) {
  return normalizeShares(shares).items.find((it) => it.kind === kind && it.id === id) || null;
}

function mapItem(shares, kind, id, fn, now) {
  const s = normalizeShares(shares);
  const idx = s.items.findIndex((it) => it.kind === kind && it.id === id);
  if (idx < 0) throw codedError(`no share for ${kind} ${id}`, "SHARE_NOT_FOUND");
  const nextItem = fn(s.items[idx]);
  const items = nextItem ? s.items.map((it, i) => (i === idx ? nextItem : it)) : s.items.filter((_, i) => i !== idx);
  return { ...s, items, updatedAt: isoNow(now) };
}

/**
 * Add a guest to an item (creating the item if needed). Idempotent on the
 * email (AC-036 carried over): inviting the same address twice changes
 * nothing but `title`.
 */
export function addGuest(shares, { kind, id, title, email, mode = "read", now }) {
  requireNow(now, "addGuest");
  const e = normalizeEmail(email);
  if (!modesForKind(kind).includes(mode)) throw codedError(`mode ${mode} not allowed for ${kind}`, "BAD_MODE");
  if (typeof id !== "string" || !id) throw codedError("addGuest: id is required", "BAD_ID");
  const s = normalizeShares(shares);
  const existing = findItem(s, kind, id);
  let next;
  if (!existing) {
    next = {
      ...s,
      items: [
        ...s.items,
        { kind, id, title: title || null, guests: [], control: { holder: null, since: null } },
      ],
    };
  } else {
    next = s;
  }
  next = mapItem(next, kind, id, (it) => {
    const has = it.guests.some((g) => g.email === e);
    return {
      ...it,
      title: title || it.title || null,
      guests: has ? it.guests : [...it.guests, { email: e, mode: "off", addedAt: isoNow(now) }],
    };
  }, now);
  const current = findItem(next, kind, id).guests.find((g) => g.email === e);
  if (!existing || current.mode === "off") {
    return setGuestMode(next, { kind, id, email: e, mode, now });
  }
  return next;
}

/**
 * Change one guest's mode. Control is exclusive per item (AC-159): granting
 * it demotes any other holder to Read. Leaving Control clears the holder.
 */
export function setGuestMode(shares, { kind, id, email, mode, now }) {
  requireNow(now, "setGuestMode");
  const e = normalizeEmail(email);
  if (!modesForKind(kind).includes(mode)) throw codedError(`mode ${mode} not allowed for ${kind}`, "BAD_MODE");
  return mapItem(shares, kind, id, (it) => {
    if (!it.guests.some((g) => g.email === e)) throw codedError(`${e} is not a guest of ${kind} ${id}`, "GUEST_NOT_FOUND");
    let control = it.control || { holder: null, since: null };
    const guests = it.guests.map((g) => {
      if (g.email === e) return { ...g, mode };
      if (mode === "control" && g.mode === "control") return { ...g, mode: "read" };
      return g;
    });
    if (mode === "control") control = { holder: e, since: isoNow(now) };
    else if (control.holder === e) control = { holder: null, since: null };
    return { ...it, guests, control };
  }, now);
}

/** Remove a guest; drops the item entirely once nobody is left on it. */
export function removeGuest(shares, { kind, id, email, now }) {
  requireNow(now, "removeGuest");
  const e = normalizeEmail(email);
  return mapItem(shares, kind, id, (it) => {
    const guests = it.guests.filter((g) => g.email !== e);
    if (guests.length === it.guests.length) throw codedError(`${e} is not a guest of ${kind} ${id}`, "GUEST_NOT_FOUND");
    if (guests.length === 0) return null;
    const control = it.control?.holder === e ? { holder: null, since: null } : it.control;
    return { ...it, guests, control };
  }, now);
}

/** The "Stop all sharing" switch (AC-153/AC-154): modes stay remembered. */
export function setSharingEnabled(shares, { enabled, now }) {
  requireNow(now, "setSharingEnabled");
  return { ...normalizeShares(shares), sharingEnabled: !!enabled, updatedAt: isoNow(now) };
}

/** Host takes the wheel back (AC-158): the holder drops to Read. */
export function takeBackControl(shares, { kind = "session", id, now }) {
  requireNow(now, "takeBackControl");
  return mapItem(shares, kind, id, (it) => {
    const holder = it.control?.holder;
    if (!holder) return it;
    return {
      ...it,
      guests: it.guests.map((g) => (g.email === holder ? { ...g, mode: "read" } : g)),
      control: { holder: null, since: null },
    };
  }, now);
}

/** "off" whenever sharing is disabled, the item or the guest is unknown. */
export function effectiveGuestMode(shares, kind, id, email) {
  const s = normalizeShares(shares);
  if (!s.sharingEnabled) return "off";
  let e;
  try {
    e = normalizeEmail(email);
  } catch {
    return "off";
  }
  const it = findItem(s, kind, id);
  const g = it?.guests.find((x) => x.email === e);
  return g ? g.mode : "off";
}

/** Host UI read-only on a session while a guest holds Control (AC-157). */
export function hostReadOnlyFor(shares, sessionId) {
  const s = normalizeShares(shares);
  if (!s.sharingEnabled) return false;
  return !!findItem(s, "session", sessionId)?.control?.holder;
}

export function isAnyShareActive(shares) {
  const s = normalizeShares(shares);
  return s.sharingEnabled && s.items.some((it) => it.guests.some((g) => g.mode !== "off"));
}

/**
 * Per-guest view of the registry: email -> [{kind, id, title, mode}] with
 * only non-off entries. Empty object when sharing is disabled -- that is what
 * makes the relay delete every mirror (AC-153).
 */
export function activeEntriesByGuest(shares) {
  const s = normalizeShares(shares);
  const out = {};
  if (!s.sharingEnabled) return out;
  for (const it of s.items) {
    for (const g of it.guests) {
      if (g.mode === "off") continue;
      (out[g.email] ||= []).push({ kind: it.kind, id: it.id, title: it.title || null, mode: g.mode });
    }
  }
  return out;
}

/** Every guest email ever listed (the relay keeps their folders tidy even when all are off). */
export function allGuestEmails(shares) {
  const set = new Set();
  for (const it of normalizeShares(shares).items) for (const g of it.guests) set.add(g.email);
  return [...set].sort();
}

// ---------------------------------------------------------------------------
// Guest ops (outbox)
// ---------------------------------------------------------------------------

/**
 * Normalizes one outbox op or throws BAD_OP. `allowedModels` (the host's
 * MODEL_OPTIONS ids) is enforced here because transcript-model#setModel only
 * checks non-emptiness -- a guest must not be able to pick a model id the
 * host's own picker would never offer.
 */
export function validateOp(op, { allowedModels } = {}) {
  if (!op || typeof op !== "object") throw codedError("op must be an object", "BAD_OP");
  if (typeof op.opId !== "string" || !/^[A-Za-z0-9_-]{6,80}$/.test(op.opId)) throw codedError("op.opId missing or malformed", "BAD_OP");
  if (!GUEST_OP_TYPES.includes(op.type)) throw codedError(`op.type ${op.type} not supported`, "BAD_OP");
  const out = { opId: op.opId, type: op.type, ts: Number.isFinite(op.ts) ? op.ts : null };
  if (op.type === "enqueue" || op.type === "force") {
    if (typeof op.text !== "string" || !op.text.trim()) throw codedError(`${op.type} needs text`, "BAD_OP");
    out.text = op.text.slice(0, 20000);
  }
  if (op.type === "setModel") {
    if (typeof op.model !== "string" || !op.model.trim()) throw codedError("setModel needs model", "BAD_OP");
    out.model = op.model.trim();
    if (Array.isArray(allowedModels) && allowedModels.length && !allowedModels.includes(out.model)) {
      throw codedError(`model ${out.model} is not offered by the host`, "BAD_OP");
    }
  }
  if (op.type === "setMode") {
    if (typeof op.mode !== "string" || !op.mode.trim()) throw codedError("setMode needs mode", "BAD_OP");
    out.mode = op.mode.trim();
  }
  if (op.type === "removeQueued") {
    if (typeof op.queueItemId !== "string" || !op.queueItemId) throw codedError("removeQueued needs queueItemId", "BAD_OP");
    out.queueItemId = op.queueItemId;
  }
  return out;
}

/** Read and Off never write (AC-155); Control needs the guest to be the holder. */
export function isOpAllowed(shares, sessionId, email, _op) {
  const mode = effectiveGuestMode(shares, "session", sessionId, email);
  if (mode === "concurrent") return true;
  if (mode === "control") return findItem(shares, "session", sessionId)?.control?.holder === normalizeEmail(email);
  return false;
}

/** Ops from an outbox this relay has not applied yet, in outbox order. */
export function selectNewOps(outbox, seenOpIds) {
  const seen = new Set(seenOpIds || []);
  const ops = Array.isArray(outbox?.ops) ? outbox.ops : [];
  return ops.filter((op) => op && typeof op.opId === "string" && !seen.has(op.opId));
}

export function displayNameOf(guest) {
  const name = typeof guest?.name === "string" ? guest.name.trim() : "";
  return name || guest?.email || "A guest";
}

/** English system-note text for a non-message op (AC-156); null for plain messages. */
export function systemNoteForOp(op, guest) {
  const who = displayNameOf(guest);
  switch (op.type) {
    case "force":
      return `${who} forced a message ahead of the queue`;
    case "stop":
      return `${who} stopped the running turn`;
    case "setModel":
      return `${who} switched model to ${op.model}`;
    case "setMode":
      return `${who} switched mode to ${op.mode}`;
    case "removeQueued":
      return `${who} removed a queued message`;
    default:
      return null;
  }
}

/**
 * Pure transform: apply one validated guest op to a session record, tagging
 * authorship and appending the system note. Throws the underlying
 * transcript-model error (e.g. unknown model) unchanged.
 */
export function applyGuestOp(record, op, guest, now) {
  requireNow(now, "applyGuestOp");
  const author = normalizeEmail(guest?.email);
  let next;
  switch (op.type) {
    case "enqueue":
      next = enqueueMessage(record, { text: op.text, now, author });
      break;
    case "force":
      next = requestForce(record, { text: op.text, now, author });
      break;
    case "stop":
      next = requestStop(record, { now });
      break;
    case "setModel":
      next = setModel(record, { model: op.model, now });
      break;
    case "setMode":
      next = setMode(record, { mode: op.mode, now });
      break;
    case "removeQueued":
      next = removeQueuedMessage(record, { id: op.queueItemId });
      break;
    default:
      throw codedError(`op.type ${op.type} not supported`, "BAD_OP");
  }
  const note = systemNoteForOp(op, guest);
  if (note) next = appendMessage(next, { role: "system", text: note, now });
  return next;
}

// ---------------------------------------------------------------------------
// What a guest gets to see (AC-161): projections, never raw records
// ---------------------------------------------------------------------------

export function projectSessionForGuest(record, { mode, controlHolder, lastIngestedOpId, now }) {
  requireNow(now, "projectSessionForGuest");
  const messages = (Array.isArray(record?.messages) ? record.messages : []).map((m) => {
    const out = { role: m.role, text: m.text, ts: m.ts };
    if (m.author) out.author = m.author;
    return out;
  });
  const queue = (Array.isArray(record?.queue) ? record.queue : []).map((q) => {
    const out = { id: q.id, text: q.text, ts: q.ts };
    if (q.author) out.author = q.author;
    return out;
  });
  const streaming = record?.streaming && typeof record.streaming.text === "string"
    ? { text: record.streaming.text, updatedAt: record.streaming.updatedAt || null }
    : null;
  return {
    schemaVersion: 1,
    kind: "session",
    id: record.id,
    title: deriveTitle(record),
    status: record.status || null,
    model: record.model || null,
    mode: record.mode || null,
    pendingAction: record.pendingAction || null,
    archived: !!record.archived,
    messages,
    queue,
    streaming,
    yourMode: mode,
    controlHolder: controlHolder || null,
    lastIngestedOpId: lastIngestedOpId || null,
    sourceUpdatedAt: record.updatedAt || null,
    mirroredAt: isoNow(now),
  };
}

/** One IDE tab (Cursor or Claude Code) out of a mirror snapshot's openTabs/historyTabs. */
export function findTabInSnapshot(snapshot, tabId) {
  const pools = [snapshot?.openTabs, snapshot?.historyTabs, snapshot?.tabs];
  for (const pool of pools) {
    if (!Array.isArray(pool)) continue;
    const hit = pool.find((t) => t && t.composerId === tabId);
    if (hit) return hit;
  }
  return null;
}

export function projectTabForGuest(tab, { kind, id, now }) {
  requireNow(now, "projectTabForGuest");
  if (kind !== "cursor-tab" && kind !== "claude-tab") throw codedError(`not a tab kind: ${kind}`, "BAD_KIND");
  const thread = (Array.isArray(tab?.thread) ? tab.thread : []).map((t) => ({
    role: t.role,
    text: t.text,
    toolCalls: Array.isArray(t.toolCalls) ? t.toolCalls : [],
  }));
  return {
    schemaVersion: 1,
    kind,
    id,
    title: tab?.title || null,
    tracked: !!tab,
    lastActivityAt: tab?.lastActivityAt || null,
    messageCount: Number.isFinite(tab?.messageCount) ? tab.messageCount : thread.length,
    waitingOn: tab?.waitingOn || null,
    thread,
    yourMode: "read",
    mirroredAt: isoNow(now),
  };
}

export function buildManifest({ hostUpn, hostName, entries, now }) {
  requireNow(now, "buildManifest");
  return {
    schemaVersion: 1,
    host: { email: normalizeEmail(hostUpn), name: hostName || null },
    items: (entries || []).map((e) => ({
      kind: e.kind,
      id: e.id,
      title: e.title || null,
      mode: e.mode,
      file: `items/${itemFileName(e.kind, e.id)}`,
    })),
    updatedAt: isoNow(now),
  };
}

// ---------------------------------------------------------------------------
// Links + Stage 1 migration
// ---------------------------------------------------------------------------

export function buildShareLink({ baseUrl, kind, id, hostUpn }) {
  itemFileName(kind, id); // validates kind
  const base = String(baseUrl || "").replace(/#.*$/, "");
  return `${base}#shared/${encodeURIComponent(kind)}/${encodeURIComponent(id)}?host=${encodeURIComponent(normalizeEmail(hostUpn))}`;
}

/** Parses the hash part of a share link; null when it isn't one. */
export function parseShareLink(hash) {
  const m = /^#?shared\/([^/?]+)\/([^?]+)(?:\?(.*))?$/.exec(String(hash || ""));
  if (!m) return null;
  const kind = decodeURIComponent(m[1]);
  if (!SHARE_KINDS.includes(kind)) return null;
  const params = new URLSearchParams(m[3] || "");
  return { kind, id: decodeURIComponent(m[2]), host: params.get("host") || null };
}

/**
 * One-off Stage 1 -> Stage 2 migration: every session record with a
 * `sharedWith` list becomes a registry item with those guests in Read mode
 * (or Off when that record's Stage 1 toggle was off). Existing registry
 * entries win.
 */
export function migrateStage1Records(shares, records, now) {
  requireNow(now, "migrateStage1Records");
  let next = normalizeShares(shares);
  if (!next.updatedAt) next = { ...next, updatedAt: isoNow(now) };
  for (const r of records || []) {
    const list = Array.isArray(r?.sharedWith) ? r.sharedWith : [];
    for (const entry of list) {
      let email;
      try {
        email = normalizeEmail(entry?.email);
      } catch {
        continue;
      }
      const it = findItem(next, "session", r.id);
      if (it && it.guests.some((g) => g.email === email)) continue;
      next = addGuest(next, {
        kind: "session",
        id: r.id,
        title: deriveTitle(r),
        email,
        mode: r.sharingEnabled ? "read" : "off",
        now,
      });
    }
  }
  return next;
}
