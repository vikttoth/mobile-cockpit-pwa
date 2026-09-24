// mobile-cockpit / lib / transcript-model.mjs
//
// SPEC task 3 (flows/mobile-cockpit/SPEC.md, S-003). Pure data model for the
// v2 state store, which replaces v1's single `cursor-cockpit/state.json`
// blob (one `prompt` string + one `output` string per session) with two
// file kinds:
//
//   cursor-cockpit/sessions.json        light index: id, chatId, title,
//                                        status, owner, updatedAt. Polled
//                                        by the list view -- stays small no
//                                        matter how long any one
//                                        conversation gets.
//   cursor-cockpit/sessions/<id>.json   full record: messages[] (role/text/
//                                        ts), queue[], lease, model, mode,
//                                        cwd, worktree. Loaded on open.
//
// A v1 session could not render a scrollback (AC-019) because it only ever
// held the LATEST prompt/output pair, and putting the whole transcript in
// the index file the list view polls would make that poll slower with every
// message any session ever received.
//
// Pure and hermetic like graph-state.mjs's other pure helpers: no I/O, no
// fetch, no Date.now() reads -- every function takes `now` (epoch ms)
// explicitly so tests are deterministic. `queue[]` and `lease` are part of
// the record shape from the start (tasks 6/7 own their *behaviour*; this
// task only needs the shape to exist so there is no second schema bump).

"use strict";

const VALID_ROLES = Object.freeze(["user", "assistant", "system"]);
const TITLE_MAX_LEN = 60;
// Mirrors daemon/lib/spawn-agent.mjs's VALID_MODES -- kept as a separate
// constant rather than a shared import because this module has zero
// dependencies on daemon/ by design (pure model, no daemon-loop coupling).
const VALID_MODES = Object.freeze(["agent", "plan", "ask"]);
// SPEC task 7 (S-010/S-011, AC-025/026/027): the two hosts that can hold an
// exclusive write lease. "daemon" is the unleased default and is never a
// take-over target -- handBackLease is how a record returns to it.
const LEASABLE_OWNERS = Object.freeze(["laptop", "phone"]);
const DEFAULT_LEASE_EXPIRY_MS = 90_000;

function isoNow(nowMs) {
  return new Date(nowMs).toISOString();
}

function requireNonEmptyString(value, label) {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${label} is required and must be a non-empty string`);
  }
}

/**
 * Construct a brand-new session record. `owner` starts as "daemon" (the
 * default holder before any lease hand-off -- see SPEC.md AC-025/AC-026).
 *
 * `chatId` is optional (mobile follow-along, 2026-09-24): AC-001's "mint a
 * chat id via cursor-agent create-chat" needs the real CLI, which only the
 * daemon's own machine has -- a browser client (the phone PWA) can create a
 * session record directly against Graph but cannot mint one itself. Such a
 * record is built with `chatId: null` and mirrors that into the light
 * index too, exactly like every other v2 field; `assignChatId` (below) is
 * the daemon-side completion step once its own tick notices one missing.
 * `daemon/lib/local-server.mjs`'s `POST /api/sessions` (the laptop/Cursor-tab
 * path) still always mints one up front and passes it here unchanged.
 *
 * @param {object} opts
 * @param {string} opts.id                  cockpit-assigned session id
 * @param {string|null} [opts.chatId]        cursor-agent chat id (create-chat); null if not yet minted
 * @param {string|null} [opts.model]
 * @param {string|null} [opts.mode]   "agent" | "plan" | "ask"
 * @param {string|null} [opts.cwd]
 * @param {string|null} [opts.worktree]
 * @param {string|null} [opts.parentId]  sub-agent wiring (2026-09-24): the id of
 *   the session that started this one as a delegated parallel task, or null
 *   for a top-level session. Deliberately NOT a special session "type" --
 *   a sub-agent is an ordinary v2 session the daemon tick treats exactly
 *   like any other; `parentId` only exists for the UI to group/link them.
 *   No `childIds[]` either: the UI derives "children of X" by filtering
 *   the light index for `parentId === X` rather than maintaining a second,
 *   invertible list that could drift from the records it describes.
 * @param {number} opts.now           epoch ms
 * @returns {object} session record
 */
export function buildSessionRecord(opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("buildSessionRecord(opts): opts is required");
  }
  requireNonEmptyString(opts.id, "id");
  if (opts.chatId !== undefined && opts.chatId !== null) {
    requireNonEmptyString(opts.chatId, "chatId");
  }
  if (opts.parentId !== undefined && opts.parentId !== null) {
    requireNonEmptyString(opts.parentId, "parentId");
  }
  if (!Number.isFinite(opts.now)) {
    throw new Error("buildSessionRecord: now must be a finite epoch-ms number");
  }
  const nowIso = isoNow(opts.now);
  return {
    schemaVersion: 1,
    id: opts.id,
    chatId: opts.chatId ?? null,
    parentId: opts.parentId ?? null,
    status: "pending",
    owner: "daemon",
    archived: false,
    pendingAction: null,
    messages: [],
    queue: [],
    lease: null,
    streaming: null,
    model: opts.model ?? null,
    mode: opts.mode ?? null,
    cwd: opts.cwd ?? null,
    worktree: opts.worktree ?? null,
    createdAt: nowIso,
    updatedAt: nowIso,
  };
}

/**
 * Mobile follow-along (2026-09-24): the daemon-side completion of a
 * browser-created record's missing `chatId` (see `buildSessionRecord`'s
 * doc comment). Throws if the record already has one -- a chat id is
 * minted exactly once per session; re-minting would silently orphan
 * whatever `--resume` context the first one already built up.
 *
 * @param {object} record
 * @param {{chatId: string, now: number}} opts
 * @returns {object} new session record
 */
export function assignChatId(record, opts) {
  requireNonEmptyString(opts?.chatId, "chatId");
  if (!Number.isFinite(opts?.now)) {
    throw new Error("assignChatId: now must be a finite epoch-ms number");
  }
  if (record.chatId) {
    const err = new Error(`assignChatId: record ${record.id} already has a chatId`);
    err.code = "CHAT_ID_ALREADY_ASSIGNED";
    throw err;
  }
  return { ...record, chatId: opts.chatId, updatedAt: isoNow(opts.now) };
}

/**
 * Pure append: returns a NEW record with one more entry in `messages[]`.
 * Does not mutate `record`. Throws on an unsupported role or empty text --
 * loud failure, same pattern as write-helpers.mjs's mergeAppendSession.
 *
 * @param {object} record
 * @param {{role: string, text: string, now: number}} msg
 * @returns {object} new session record
 */
export function appendMessage(record, msg) {
  if (!msg || typeof msg !== "object") {
    throw new Error("appendMessage: msg is required");
  }
  if (!VALID_ROLES.includes(msg.role)) {
    throw new Error(
      `appendMessage: role "${msg.role}" is not supported; expected one of: ${VALID_ROLES.join(", ")}`,
    );
  }
  if (typeof msg.text !== "string" || msg.text.trim().length === 0) {
    throw new Error("appendMessage: text is required and must not be empty");
  }
  if (!Number.isFinite(msg.now)) {
    throw new Error("appendMessage: now must be a finite epoch-ms number");
  }
  const entry = { role: msg.role, text: msg.text, ts: msg.now };
  return {
    ...record,
    messages: [...record.messages, entry],
    updatedAt: isoNow(msg.now),
  };
}

/**
 * Human-friendly title, derived from the first user message. Falls back to
 * "(untitled)" before any user message exists (e.g. right after
 * buildSessionRecord, before the first prompt lands).
 *
 * @param {object} record
 * @returns {string}
 */
export function deriveTitle(record) {
  const messages = Array.isArray(record?.messages) ? record.messages : [];
  const firstUser = messages.find((m) => m && m.role === "user" && typeof m.text === "string");
  if (!firstUser) return "(untitled)";
  const text = firstUser.text.trim();
  if (text.length <= TITLE_MAX_LEN) return text;
  return text.slice(0, TITLE_MAX_LEN - 3) + "...";
}

/**
 * Derive the light index-entry shape from a full session record. Only
 * carries what the list view needs -- MUST NOT leak `messages[]` (that is
 * exactly the field split this task exists to enforce; see the "no
 * messages[] leak" assertion in transcript-model-unit.sh).
 *
 * @param {object} record
 * @returns {{id: string, chatId: string, parentId: string|null, title: string, status: string, owner: string, updatedAt: string}}
 */
export function buildIndexEntry(record) {
  return {
    id: record.id,
    chatId: record.chatId,
    parentId: record.parentId ?? null,
    title: deriveTitle(record),
    status: record.status,
    owner: record.owner,
    archived: record.archived,
    updatedAt: record.updatedAt,
  };
}

/**
 * Pure append to `queue[]` (SPEC task 6, S-004, AC-012): "When the user
 * sends with Queue while a turn is running, the message shall be appended
 * to queue[], shown as queued, and delivered as the next turn when the
 * current one ends." This function only owns the append -- the "delivered
 * as the next turn" half is the daemon tick noticing a run finished and
 * popping queue[0], which is daemon-loop wiring left for a later step (see
 * the SPEC task 6 note); it is not needed to make queueing itself correct
 * and testable.
 *
 * Queue item ids are derived from the record's own queue length rather
 * than `now`, so two items enqueued in the same millisecond (plausible
 * under a fast synthetic clock in tests) never collide.
 *
 * @param {object} record
 * @param {{text: string, now: number}} opts
 * @returns {object} new session record
 */
export function enqueueMessage(record, opts) {
  if (!opts || typeof opts !== "object") {
    throw new Error("enqueueMessage: opts is required");
  }
  if (typeof opts.text !== "string" || opts.text.trim().length === 0) {
    throw new Error("enqueueMessage: text is required and must not be empty");
  }
  if (!Number.isFinite(opts.now)) {
    throw new Error("enqueueMessage: now must be a finite epoch-ms number");
  }
  const queue = Array.isArray(record.queue) ? record.queue : [];
  const item = { id: `q${queue.length + 1}`, role: "user", text: opts.text, ts: opts.now };
  return {
    ...record,
    queue: [...queue, item],
    updatedAt: isoNow(opts.now),
  };
}

/**
 * Pure removal from `queue[]` (SPEC task 6, S-006, AC-015 "remove"). Throws
 * `code: "QUEUE_ITEM_NOT_FOUND"` on an unknown id -- loud failure, same
 * pattern as `removeIndexEntry`, so a stale UI can't silently no-op.
 *
 * @param {object} record
 * @param {{id: string}} opts
 * @returns {object} new session record
 */
export function removeQueuedMessage(record, opts) {
  const queue = Array.isArray(record.queue) ? record.queue : [];
  const next = queue.filter((q) => !(q && q.id === opts?.id));
  if (next.length === queue.length) {
    const err = new Error(`removeQueuedMessage: no queued item with id ${opts?.id}`);
    err.code = "QUEUE_ITEM_NOT_FOUND";
    throw err;
  }
  return { ...record, queue: next };
}

/**
 * Pure reorder of `queue[]` (SPEC task 6, S-006 area, AC-015 "reorder").
 * `orderedIds` must be exactly the current queue's ids, in the desired
 * order -- a partial list or an id that does not belong to this record's
 * queue throws, rather than silently dropping or inventing an entry.
 *
 * @param {object} record
 * @param {{orderedIds: string[]}} opts
 * @returns {object} new session record
 */
export function reorderQueuedMessages(record, opts) {
  const queue = Array.isArray(record.queue) ? record.queue : [];
  const orderedIds = Array.isArray(opts?.orderedIds) ? opts.orderedIds : [];
  const byId = new Map(queue.map((q) => [q.id, q]));
  if (orderedIds.length !== queue.length || !orderedIds.every((id) => byId.has(id))) {
    throw new Error(
      "reorderQueuedMessages: orderedIds must contain exactly the current queue's ids",
    );
  }
  return { ...record, queue: orderedIds.map((id) => byId.get(id)) };
}

/**
 * Pure model switch (SPEC task 8, S-008, AC-016): "When the user changes
 * the model, the next turn shall spawn with --model <new>; earlier turns
 * are unaffected." Only touches `model` + `updatedAt` -- `messages[]` is
 * untouched, and since `appendMessage` never rewrites history, an earlier
 * turn's outcome is unaffected by construction, not by extra bookkeeping.
 * `daemon/lib/spawn-agent.mjs#buildCursorAgentArgv` reads `session.model`
 * fresh on every call, so this setter is the whole mechanism -- the next
 * spawn just sees whatever `model` currently is.
 *
 * @param {object} record
 * @param {{model: string, now: number}} opts
 * @returns {object} new session record
 */
export function setModel(record, opts) {
  requireNonEmptyString(opts?.model, "model");
  if (!Number.isFinite(opts.now)) {
    throw new Error("setModel: now must be a finite epoch-ms number");
  }
  return { ...record, model: opts.model, updatedAt: isoNow(opts.now) };
}

/**
 * Pure mode switch (SPEC task 8, S-008 area, AC-017): "When the user
 * selects mode plan or ask, the next turn shall spawn with --mode
 * accordingly." Same shape as `setModel`. `mode` must be one of
 * VALID_MODES ("agent" | "plan" | "ask") -- "agent" is a valid value here
 * even though `buildCursorAgentArgv` omits the flag for it (the CLI has no
 * `--mode agent`; that IS its default).
 *
 * @param {object} record
 * @param {{mode: string, now: number}} opts
 * @returns {object} new session record
 */
export function setMode(record, opts) {
  if (!VALID_MODES.includes(opts?.mode)) {
    throw new Error(
      `setMode: mode '${opts?.mode}' is not supported; expected one of: ${VALID_MODES.join(", ")}`,
    );
  }
  if (!Number.isFinite(opts.now)) {
    throw new Error("setMode: now must be a finite epoch-ms number");
  }
  return { ...record, mode: opts.mode, updatedAt: isoNow(opts.now) };
}

/**
 * Pure archive (SPEC task 11, S-014 area, AC-028): "the session shall leave
 * the active list, keep its transcript, and remain restorable." Sets only
 * `archived` + `updatedAt` -- `messages[]`/`queue[]`/`status` are all left
 * exactly as they are, which IS "keep its transcript": nothing is deleted
 * or summarised, a session is just no longer in the default list view
 * (that filter is a UI concern, not this function's).
 *
 * Does NOT check `status` (e.g. "is this session running?") -- AC-029's
 * "must stop a running session before archiving it" is deliberately not
 * enforced here; see the SPEC task 11 note for why (it depends on task 6's
 * still-open force/stop daemon-loop design).
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {object} new session record
 */
export function archiveSession(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("archiveSession: now must be a finite epoch-ms number");
  }
  return { ...record, archived: true, updatedAt: isoNow(opts.now) };
}

/**
 * Pure unarchive -- the "remain restorable" half of AC-028. Symmetric with
 * `archiveSession`.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {object} new session record
 */
export function unarchiveSession(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("unarchiveSession: now must be a finite epoch-ms number");
  }
  return { ...record, archived: false, updatedAt: isoNow(opts.now) };
}

/**
 * Pure stop request (SPEC task 6 completion, S-007, AC-014): "When the user
 * presses Stop, the daemon shall terminate the process; the session stays
 * listed as stopped and remains resumable." This function only owns the
 * INTENT -- the daemon's concurrent poller (daemon/lib/child-registry.mjs)
 * is what actually notices `pendingAction: "stop"` and calls `.kill()` on
 * the running child; once the child has exited, `deriveNextAction` (below)
 * is what turns that intent into the final `status: "stopped"`.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {object} new session record
 */
export function requestStop(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("requestStop: now must be a finite epoch-ms number");
  }
  return { ...record, pendingAction: "stop", updatedAt: isoNow(opts.now) };
}

/**
 * Pure force request (SPEC task 6 completion, S-005, AC-013): "When the
 * user sends with Force while a turn is running, the daemon shall
 * terminate the current process and start the new turn immediately."
 * Composes two already-tested primitives rather than duplicating their
 * logic: `enqueueMessage` appends the new message, then
 * `reorderQueuedMessages` moves it to the front -- "immediately" means
 * "next", i.e. ahead of anything already queued, not appended behind it.
 * Same intent/execution split as `requestStop`: this only marks
 * `pendingAction: "force"`; the registry does the kill, `deriveNextAction`
 * does the "then spawn the front of the queue" half once the child exits.
 *
 * @param {object} record
 * @param {{text: string, now: number}} opts
 * @returns {object} new session record
 */
export function requestForce(record, opts) {
  const withNewMessage = enqueueMessage(record, { text: opts?.text, now: opts?.now });
  const newMessageId = withNewMessage.queue[withNewMessage.queue.length - 1].id;
  const otherIds = withNewMessage.queue.slice(0, -1).map((q) => q.id);
  const reordered = reorderQueuedMessages(withNewMessage, { orderedIds: [newMessageId, ...otherIds] });
  return { ...reordered, pendingAction: "force" };
}

/**
 * SPEC task 6 completion (S-005, S-007, AC-012's "delivered as the next
 * turn" half): the tick-transition decision the daemon did not have before
 * -- "once a turn ends (however it ended), what should happen next?" Pure
 * decision, not I/O: the daemon calls this once a child has actually
 * exited (whether by natural completion, or because `child-registry.mjs`
 * killed it for a stop/force request) and acts on the returned `type`.
 *
 * Priority: an explicit Stop always wins and ignores the queue (AC-014
 * says nothing about draining it, and "stays listed as stopped" reads as
 * "don't just silently start something else") -- the queue is left intact
 * for a future resume/follow-up, not cleared. Otherwise, a non-empty
 * queue -- whether put there by Force's priority-enqueue or by a plain
 * `enqueueMessage` while a turn was running -- always drains next; this is
 * also what AC-012's "delivered as the next turn when the current one
 * ends" means for a NATURAL completion, not just a Force. An empty queue
 * with no pending action is the ordinary case: the caller applies whatever
 * run outcome it already computed (done/failed via
 * `spawn-agent.mjs#formatRunResult`/`formatRunError`) -- this function has
 * no opinion on that mapping, only on whether the queue should intervene
 * first.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {{type: "finalize-stopped", record: object}
 *          |{type: "spawn-next", message: object, record: object}
 *          |{type: "finalize-normal", record: object}}
 */
export function deriveNextAction(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("deriveNextAction: now must be a finite epoch-ms number");
  }
  const updatedAt = isoNow(opts.now);

  if (record.pendingAction === "stop") {
    return {
      type: "finalize-stopped",
      record: { ...record, pendingAction: null, status: "stopped", updatedAt },
    };
  }

  const queue = Array.isArray(record.queue) ? record.queue : [];
  if (queue.length > 0) {
    const [message, ...rest] = queue;
    return {
      type: "spawn-next",
      message,
      record: { ...record, queue: rest, pendingAction: null, updatedAt },
    };
  }

  return {
    type: "finalize-normal",
    record: { ...record, pendingAction: null, updatedAt },
  };
}

/**
 * Pure lease take-over (SPEC task 7, S-010, AC-025/AC-027): "the transition
 * shall be atomic under eTag, and never leave two writers." This function
 * has no opinion about the CURRENT holder -- take-over always wins, by
 * design (a user explicitly clicking "take over" is preempting whoever has
 * it, not negotiating with them). Atomicity itself comes from the caller
 * writing the result through the record's own etag (session-store.mjs),
 * the same retry-on-412 contract every other mutation here already uses.
 *
 * @param {object} record
 * @param {{owner: "laptop"|"phone", now: number, expiryWindowMs?: number}} opts
 * @returns {object} new session record
 */
export function takeOverLease(record, opts) {
  if (!LEASABLE_OWNERS.includes(opts?.owner)) {
    throw new Error(
      `takeOverLease: owner '${opts?.owner}' is not leasable; expected one of: ${LEASABLE_OWNERS.join(", ")}`,
    );
  }
  if (!Number.isFinite(opts?.now)) {
    throw new Error("takeOverLease: now must be a finite epoch-ms number");
  }
  const expiryWindowMs = Number.isFinite(opts.expiryWindowMs) ? opts.expiryWindowMs : DEFAULT_LEASE_EXPIRY_MS;
  const nowIso = isoNow(opts.now);
  return {
    ...record,
    owner: opts.owner,
    lease: { owner: opts.owner, heartbeatAt: nowIso, expiresAt: isoNow(opts.now + expiryWindowMs) },
    updatedAt: nowIso,
  };
}

/**
 * Pure hand-back (SPEC task 7, AC-025/AC-027): returns ownership to the
 * daemon, clearing the lease. Symmetric with `takeOverLease`.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {object} new session record
 */
export function handBackLease(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("handBackLease: now must be a finite epoch-ms number");
  }
  return { ...record, owner: "daemon", lease: null, updatedAt: isoNow(opts.now) };
}

/**
 * Pure heartbeat renewal (SPEC task 7, AC-026's counterpart -- what keeps a
 * lease from expiring while its holder is still alive). Requires the
 * caller to name the owner it believes it is, and throws
 * `code: "LEASE_OWNER_MISMATCH"` if that does not match the lease's actual
 * holder -- a stale or wrong host renewing someone else's lease is exactly
 * the "two writers" AC-027 warns against, so this fails loudly rather than
 * silently extending the wrong lease.
 *
 * @param {object} record
 * @param {{owner: "laptop"|"phone", now: number, expiryWindowMs?: number}} opts
 * @returns {object} new session record
 */
export function renewLeaseHeartbeat(record, opts) {
  if (!record.lease) {
    const err = new Error("renewLeaseHeartbeat: no active lease to renew");
    err.code = "NO_ACTIVE_LEASE";
    throw err;
  }
  if (record.lease.owner !== opts?.owner) {
    const err = new Error(
      `renewLeaseHeartbeat: caller claims owner '${opts?.owner}' but the lease is held by '${record.lease.owner}'`,
    );
    err.code = "LEASE_OWNER_MISMATCH";
    throw err;
  }
  if (!Number.isFinite(opts?.now)) {
    throw new Error("renewLeaseHeartbeat: now must be a finite epoch-ms number");
  }
  const expiryWindowMs = Number.isFinite(opts.expiryWindowMs) ? opts.expiryWindowMs : DEFAULT_LEASE_EXPIRY_MS;
  const nowIso = isoNow(opts.now);
  return {
    ...record,
    lease: { owner: record.lease.owner, heartbeatAt: nowIso, expiresAt: isoNow(opts.now + expiryWindowMs) },
    updatedAt: nowIso,
  };
}

/**
 * Pure expiry check (SPEC task 7, S-011, AC-026): "If a lease heartbeat is
 * missed beyond the expiry window, then the lease shall lapse and
 * ownership shall return to the daemon, recorded as `lease expired`."
 * A safe no-op when there is no lease, or the lease is still within its
 * window -- the daemon calls this on every tick for every leased session
 * without needing to pre-check whether one exists.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {{expired: boolean, record: object}}
 */
export function expireLeaseIfStale(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("expireLeaseIfStale: now must be a finite epoch-ms number");
  }
  if (!record.lease || opts.now <= Date.parse(record.lease.expiresAt)) {
    return { expired: false, record: { ...record } };
  }
  const previousOwner = record.lease.owner;
  const withNote = appendMessage(record, {
    role: "system",
    text: `Lease expired for ${previousOwner}; ownership returned to daemon.`,
    now: opts.now,
  });
  return {
    expired: true,
    record: { ...withNote, owner: "daemon", lease: null },
  };
}

/**
 * SPEC task 6/7 v2-action-tick wiring (2026-09-24): pop `queue[0]` and start
 * it as the running turn -- append it to `messages[]` (a queued item becomes
 * part of the actual transcript the moment it is sent, not before), flip
 * `status` to `"running"`, clear `pendingAction` (a stale stop/force from
 * before this pick would otherwise immediately look "already handled" to
 * the kill-check on the very next tick). Throws if the queue is empty --
 * the caller (`daemon/lib/v2-action-tick.mjs`) already checked
 * `queue.length > 0` before calling this, so an empty queue here is a
 * caller bug, not a normal condition to swallow.
 *
 * @param {object} record
 * @param {{now: number}} opts
 * @returns {object} new session record
 */
export function startNextTurn(record, opts) {
  if (!Number.isFinite(opts?.now)) {
    throw new Error("startNextTurn: now must be a finite epoch-ms number");
  }
  const queue = Array.isArray(record.queue) ? record.queue : [];
  if (queue.length === 0) {
    throw new Error("startNextTurn: queue is empty, nothing to start");
  }
  const [message, ...rest] = queue;
  const appended = appendMessage(
    { ...record, queue: rest, pendingAction: null },
    { role: "user", text: message.text, now: opts.now },
  );
  return { ...appended, status: "running", startedAt: isoNow(opts.now) };
}

/**
 * Live-partial-output wiring (mobile follow-along, found missing 2026-09-24
 * building out the v2 action tick): pure setter for "here is the latest
 * partial text a still-running turn has produced." Not part of
 * `buildIndexEntry`'s shape -- the light index stays index-only, callers
 * poll the full record for this. `resolveFinishedTurn` (below) is what
 * clears it once a turn actually ends, and -- if the run never produced a
 * clean final `output` (e.g. killed mid-flight) -- promotes whatever text
 * landed here last into the transcript instead of losing it.
 *
 * @param {object} record
 * @param {{text: string, agentId: string|null, now: number}} opts
 * @returns {object} new session record
 */
export function updateStreamingOutput(record, opts) {
  if (typeof opts?.text !== "string") {
    throw new Error("updateStreamingOutput: opts.text must be a string");
  }
  if (!Number.isFinite(opts?.now)) {
    throw new Error("updateStreamingOutput: now must be a finite epoch-ms number");
  }
  return {
    ...record,
    streaming: { text: opts.text, agentId: opts.agentId ?? null, updatedAt: isoNow(opts.now) },
  };
}

/**
 * SPEC task 6/7 v2-action-tick wiring (2026-09-24): "a turn just ended --
 * what does the record look like now?" Merges the run's result delta
 * (`daemon/lib/spawn-agent.mjs#formatRunResult`/`formatRunError`'s shape:
 * `status`/`output`/`cursorAgentId`/`completedAt` or
 * `status`/`errorMessage`/`errorRetryable`/`completedAt`) into the record,
 * THEN applies `deriveNextAction`'s stop/queue-drain priority on top --
 * `deriveNextAction` deliberately has no opinion on the done/failed mapping
 * (see its own doc comment), so the merge has to happen first, here, not
 * inside it.
 *
 * Live-partial-output wiring (2026-09-24, found missing while building the
 * mobile follow-along view): the agent's reply was never landing in
 * `messages[]` at all -- only in the top-level `output` field, which no UI
 * actually rendered. Fixed here: a non-empty `resultDelta.output` (the
 * normal, clean-finish case) is appended as `role: "assistant"`. A run that
 * never produced one (killed mid-flight -- `formatRunError`'s shape has no
 * `output` field) falls back to whatever `record.streaming.text` last
 * captured, so an interrupted turn still leaves SOMETHING in the transcript
 * rather than silently vanishing. `streaming` itself is always cleared here
 * -- a turn ending, however it ended, means nothing is live-updating anymore.
 *
 * When a queued follow-up should start immediately (`deriveNextAction`'s
 * `"spawn-next"`), that follow-up is applied here too rather than left for
 * the caller to notice separately: `deriveNextAction` already popped the
 * message and shortened `queue[]` in the record it returned, so re-running
 * `startNextTurn` on that record would pop a SECOND item by mistake -- this
 * function appends the already-popped `message` directly instead. The
 * caller's signal for "a follow-up turn started, go spawn it" is simply
 * `result.status === "running"`; `formatRunResult`/`formatRunError` never
 * produce that status themselves, so it can only mean this branch fired.
 *
 * @param {object} record
 * @param {{resultDelta: object, now: number}} opts
 * @returns {object} new session record
 */
export function resolveFinishedTurn(record, opts) {
  if (!opts?.resultDelta || typeof opts.resultDelta !== "object") {
    throw new Error("resolveFinishedTurn: opts.resultDelta is required");
  }
  if (!Number.isFinite(opts?.now)) {
    throw new Error("resolveFinishedTurn: now must be a finite epoch-ms number");
  }
  let merged = { ...record, ...opts.resultDelta };
  const replyText = typeof opts.resultDelta.output === "string" && opts.resultDelta.output.trim().length > 0
    ? opts.resultDelta.output
    : (typeof record.streaming?.text === "string" && record.streaming.text.trim().length > 0
        ? record.streaming.text
        : null);
  if (replyText !== null) {
    merged = appendMessage(merged, { role: "assistant", text: replyText, now: opts.now });
  }
  merged = { ...merged, streaming: null };

  const decision = deriveNextAction(merged, { now: opts.now });
  if (decision.type === "spawn-next") {
    const appended = appendMessage(decision.record, {
      role: "user",
      text: decision.message.text,
      now: opts.now,
    });
    return { ...appended, status: "running", startedAt: isoNow(opts.now) };
  }
  return decision.record;
}

/**
 * Pure state-merger for sessions.json: insert `entry`, or replace the
 * existing row with the same id in place (no duplicate, no reordering of
 * the other rows). Does not mutate `indexFile`.
 *
 * @param {{schemaVersion?: number, sessions?: Array}|null|undefined} indexFile
 * @param {{id: string}} entry
 * @returns {{schemaVersion: number, sessions: Array}}
 */
export function upsertIndexEntry(indexFile, entry) {
  if (!entry || typeof entry.id !== "string" || entry.id.length === 0) {
    throw new Error("upsertIndexEntry: entry.id is required");
  }
  const base = indexFile && typeof indexFile === "object" ? indexFile : { schemaVersion: 1, sessions: [] };
  const sessions = Array.isArray(base.sessions) ? base.sessions : [];
  const idx = sessions.findIndex((s) => s && s.id === entry.id);
  const nextSessions =
    idx >= 0
      ? sessions.map((s, i) => (i === idx ? entry : s))
      : [...sessions, entry];
  return {
    schemaVersion: base.schemaVersion || 1,
    sessions: nextSessions,
  };
}

/**
 * Pure state-merger for sessions.json: drop the row with `id`. Does not
 * mutate `indexFile`. Throws `code: "SESSION_NOT_FOUND"` if no row matches
 * -- loud failure so a caller can't silently no-op a bad id.
 *
 * @param {{sessions?: Array}} indexFile
 * @param {string} id
 * @returns {{schemaVersion: number, sessions: Array}}
 */
export function removeIndexEntry(indexFile, id) {
  const base = indexFile && typeof indexFile === "object" ? indexFile : { schemaVersion: 1, sessions: [] };
  const sessions = Array.isArray(base.sessions) ? base.sessions : [];
  const next = sessions.filter((s) => !(s && s.id === id));
  if (next.length === sessions.length) {
    const err = new Error(`removeIndexEntry: no session with id ${id}`);
    err.code = "SESSION_NOT_FOUND";
    throw err;
  }
  return {
    schemaVersion: base.schemaVersion || 1,
    sessions: next,
  };
}
