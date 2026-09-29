// flows/mobile-cockpit/pwa/guest-app.mjs
//
// SPEC-DELTA-2026-09-29-session-sharing-stage2: the PWA when a GUEST signs in
// (anyone but config.sharing.hostUpn -- AC-150). The guest only ever touches
// their OWN drive, under cursor-cockpit-guest/hosts/<host>/:
//   - Connect (AC-151): create that folder and grant the host write access
//     to it (Graph invite on the guest's own item; no email is sent).
//   - Read: manifest.json + items/<kind>-<id>.json, written by the host's
//     share relay; the open item is polled every few seconds with
//     If-None-Match so an unchanged file costs one small request.
//   - Write (Control / Concurrent only): append ops to outbox/<sessionId>.json;
//     the relay applies them after re-checking the mode (the guest UI's own
//     disabled state is convenience, not the security boundary).
//
// The pure helpers at the top are exported for Node tests; startGuestMode()
// is the DOM half and only runs in the browser.

export const OUTBOX_KEEP = 50;
export const GUEST_POLL_MS_DEFAULT = 5000;
export const WAITING_POLL_MS = 10000;

export function guestPaths(shareModel, { guestUpn, hostUpn }) {
  const folder = shareModel.guestHostFolder({ guestUpn, hostUpn }).driveRelativePath;
  return {
    folder,
    folderEndpoint: `/me/drive/root:/${folder}`,
    manifestEndpoint: `/me/drive/root:/${folder}/manifest.json`,
    readmeEndpoint: `/me/drive/root:/${folder}/README.json`,
    itemEndpoint: (kind, id) => `/me/drive/root:/${folder}/items/${shareModel.itemFileName(kind, id)}`,
    outboxEndpoint: (sessionId) => `/me/drive/root:/${folder}/outbox/${shareModel.outboxFileName(sessionId)}`,
  };
}

export function newOpId(now, rand) {
  const r = typeof rand === "string" && rand ? rand : Math.random().toString(36).slice(2, 10);
  return `op-${Number(now).toString(36)}-${r}`.replace(/[^A-Za-z0-9_-]/g, "");
}

/** Pure outbox append: keeps the newest OUTBOX_KEEP ops, stamps the guest header. */
export function appendOutboxOp(outbox, { op, guest, sessionId }) {
  const ops = Array.isArray(outbox?.ops) ? outbox.ops : [];
  const next = [...ops, op].slice(-OUTBOX_KEEP);
  return { schemaVersion: 1, guest: { email: guest.email, name: guest.name || null }, sessionId, ops: next };
}

/**
 * What the Shared list should show for the guest.
 *   connect  -> folder missing: the Connect card
 *   waiting  -> connected, relay has not written a manifest yet
 *   list     -> manifest present (possibly with zero items)
 */
export function deriveGuestListState({ folderExists, manifest, hostName }) {
  if (!folderExists) {
    return { kind: "connect", text: `${hostName || "The cockpit owner"} shared something with you. Tap Connect to see it here.` };
  }
  if (!manifest) {
    return { kind: "waiting", text: `Connected — waiting for ${hostName || "the owner"}'s cockpit…` };
  }
  const items = Array.isArray(manifest.items) ? manifest.items : [];
  return {
    kind: "list",
    items,
    text: items.length ? null : `${hostName || "The owner"} is not sharing anything with you right now.`,
  };
}

/** Composer + selects state for a guest looking at a mirrored chat. */
export function deriveGuestComposer({ projection, guestEmail, text, busy, composerState, shareUi }) {
  const access = shareUi.deriveGuestChatAccess(projection, guestEmail);
  const buttons = composerState.deriveComposerButtons({ status: projection?.status, text, busy, readOnly: !access.canWrite });
  return { access, buttons, selectsEnabled: access.canWrite && !busy };
}

// ---------------------------------------------------------------------------
// DOM half
// ---------------------------------------------------------------------------

/**
 * @param {object} ctx  everything app.js already has: config, account, graphFetch,
 *   loadJson, putJson, shareModel, shareUi, scrollback, composerState, setView,
 *   setStatusBadge, translateErrorMessage, populateModelSelect
 */
export function startGuestMode(ctx) {
  const { config, account, graphFetch, loadJson, putJson, shareModel, shareUi, scrollback, composerState } = ctx;
  const hostUpn = config.sharing.hostUpn;
  const hostName = config.sharing.hostDisplayName || shareUi.nameFromEmail(hostUpn);
  const guestEmail = String(account.username || "").toLowerCase();
  const paths = guestPaths(shareModel, { guestUpn: guestEmail, hostUpn });
  const pollMs = Math.max(3, Number(config.sharing.guestPollIntervalSeconds) || 5) * 1000;
  const $ = (id) => document.getElementById(id);

  let listTimer = null;
  let itemTimer = null;
  let openItem = null; // {kind, id}
  let openEtag = null;
  let openProjection = null;
  let busy = false;
  let pendingTarget = null;

  function showListError(msg) {
    const el = $("shared-list-error-state");
    if (el) {
      el.textContent = ctx.translateErrorMessage(msg);
      el.hidden = false;
    }
  }
  function showItemError(msg) {
    const el = $("shared-detail-error-state");
    if (el) {
      el.textContent = ctx.translateErrorMessage(msg);
      el.hidden = false;
    }
  }

  async function folderExists() {
    const res = await graphFetch(`${paths.folderEndpoint}?$select=id`);
    if (res.status === 404) return false;
    if (!res.ok) throw new Error(`folder check failed: ${res.status}`);
    return true;
  }

  async function loadManifest() {
    const res = await graphFetch(`${paths.manifestEndpoint}:/content`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`manifest read failed: ${res.status}`);
    return res.json();
  }

  /** AC-151: create the folder in the guest's own drive, then grant the host write access to it. */
  async function connect() {
    const status = $("guest-connect-status");
    const btn = $("btn-guest-connect");
    if (btn) btn.disabled = true;
    if (status) {
      status.hidden = false;
      status.textContent = "Connecting…";
    }
    try {
      await putJson(paths.readmeEndpoint, {
        purpose: `Mobile Cockpit: items ${hostName} shares with you appear here. Deleting this folder disconnects.`,
        host: hostUpn,
        createdAt: new Date().toISOString(),
      }, null);
      const res = await graphFetch(`${paths.folderEndpoint}:/invite`, {
        method: "POST",
        body: JSON.stringify({ recipients: [{ email: hostUpn }], requireSignIn: true, sendInvitation: false, roles: ["write"] }),
      });
      if (!res.ok) {
        const t = await res.text().catch(() => "");
        throw new Error(`could not share the folder with ${hostName}: ${res.status} ${t.slice(0, 200)}`);
      }
      if (status) status.textContent = `Connected. Waiting for ${hostName}'s cockpit…`;
      await renderGuestList();
    } catch (err) {
      if (status) status.textContent = ctx.translateErrorMessage(err.message);
      if (btn) btn.disabled = false;
    }
  }

  async function renderGuestList() {
    const errEl = $("shared-list-error-state");
    if (errEl) errEl.hidden = true;
    const heading = $("shared-list-heading");
    if (heading) heading.textContent = `Shared by ${hostName}`;
    const card = $("guest-connect-card");
    const ul = $("shared-session-list");
    const empty = $("shared-list-empty-state");
    let state;
    try {
      const exists = await folderExists();
      state = shareUi && deriveGuestListState({ folderExists: exists, manifest: exists ? await loadManifest() : null, hostName });
    } catch (err) {
      showListError(err.message);
      return;
    }
    if (card) card.hidden = state.kind !== "connect" && state.kind !== "waiting";
    const txt = $("guest-connect-text");
    if (txt) txt.textContent = state.text || "";
    const btn = $("btn-guest-connect");
    if (btn) btn.hidden = state.kind !== "connect";
    if (ul) ul.innerHTML = "";
    if (empty) {
      empty.hidden = !(state.kind === "list" && state.text);
      empty.textContent = state.text || "";
    }
    scheduleListPoll(state.kind === "waiting");
    if (state.kind !== "list") return;
    for (const item of state.items) {
      const li = document.createElement("li");
      li.className = "cockpit-session-row";
      li.tabIndex = 0;
      const title = document.createElement("span");
      title.className = "cockpit-row-title";
      title.textContent = item.title || "(untitled)";
      li.appendChild(title);
      const meta = document.createElement("span");
      meta.className = "cockpit-row-meta";
      const kindLabel = { session: "Chat", "cursor-tab": "Cursor tab", "claude-tab": "Claude Code" }[item.kind] || item.kind;
      meta.textContent = `${kindLabel} · ${shareUi.MODE_LABELS[item.mode] || item.mode}`;
      li.appendChild(meta);
      const go = () => ctx.setView("shared-detail", { kind: item.kind, id: item.id });
      li.addEventListener("click", go);
      li.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter" || ev.key === " ") {
          ev.preventDefault();
          go();
        }
      });
      ul.appendChild(li);
    }
    if (pendingTarget) {
      const hit = state.items.find((i) => i.kind === pendingTarget.kind && i.id === pendingTarget.id);
      if (hit) {
        pendingTarget = null;
        ctx.setView("shared-detail", { kind: hit.kind, id: hit.id });
      }
    }
  }

  function scheduleListPoll(on) {
    if (listTimer) clearInterval(listTimer);
    listTimer = null;
    if (!on) return;
    listTimer = setInterval(() => {
      if (document.hidden) return;
      if (document.body.dataset.view !== "shared-list") return;
      renderGuestList();
    }, WAITING_POLL_MS);
  }

  async function fetchItem(kind, id, etag) {
    const endpoint = paths.itemEndpoint(kind, id);
    const headers = etag ? { "If-None-Match": etag } : {};
    const meta = await graphFetch(`${endpoint}?$select=eTag`, { headers });
    if (meta.status === 304) return { changed: false, etag };
    if (meta.status === 404) return { changed: true, etag: null, projection: null };
    if (!meta.ok) throw new Error(`shared item read failed: ${meta.status}`);
    const m = await meta.json();
    if (etag && m.eTag === etag) return { changed: false, etag };
    const content = await graphFetch(`${endpoint}:/content`);
    if (!content.ok) throw new Error(`shared item content failed: ${content.status}`);
    return { changed: true, etag: m.eTag, projection: await content.json() };
  }

  function renderMessages(projection) {
    const container = $("shared-messages");
    if (!container) return;
    const wasAtBottom = scrollback.shouldAutoScrollToBottom({
      scrollTop: container.scrollTop,
      scrollHeight: container.scrollHeight,
      clientHeight: container.clientHeight,
    });
    container.innerHTML = "";
    const add = (cls, labelText, bodyText) => {
      const bubble = document.createElement("div");
      bubble.className = cls;
      const label = document.createElement("div");
      label.className = "v2-message-role-label";
      label.textContent = labelText;
      bubble.appendChild(label);
      const text = document.createElement("div");
      text.className = "v2-message-text";
      text.textContent = bodyText;
      bubble.appendChild(text);
      container.appendChild(bubble);
    };
    if (projection.kind === "session") {
      for (const raw of scrollback.orderMessagesForDisplay(projection.messages || [])) {
        const m = scrollback.formatSessionMessage(raw);
        const label = m.role === "user" && !m.author ? hostName : m.author && m.author.toLowerCase() === guestEmail ? "You" : m.label;
        add(`v2-message v2-role-${m.role}`, label, m.text);
      }
      if (projection.streaming && projection.streaming.text) add("v2-message v2-role-assistant v2-message-streaming", "Agent (typing…)", projection.streaming.text);
      if (Array.isArray(projection.queue) && projection.queue.length) {
        add("v2-queue-panel", `Queued (${projection.queue.length})`, projection.queue.map((q) => `• ${q.text}`).join("\n"));
      }
      if (!(projection.messages || []).length && !projection.streaming) add("v2-message v2-role-system", "System", "No messages yet.");
    } else {
      if (!projection.tracked) add("v2-message v2-role-system", "System", "This tab is no longer tracked by the owner's cockpit.");
      for (const t of projection.thread || []) {
        add(`v2-message v2-role-${t.role === "user" ? "user" : "assistant"}`, t.role === "user" ? hostName : "Agent", t.text || "");
      }
    }
    if (wasAtBottom) container.scrollTop = container.scrollHeight;
  }

  function applyComposer() {
    const wrap = $("guest-composer");
    const isSession = openProjection && openProjection.kind === "session";
    if (wrap) wrap.hidden = !isSession;
    const banner = $("shared-detail-banner");
    if (!openProjection) return;
    if (!isSession) {
      if (banner) banner.textContent = "Read-only — you can follow this tab live.";
      return;
    }
    const textEl = $("guest-composer-text");
    const st = deriveGuestComposer({ projection: openProjection, guestEmail, text: textEl ? textEl.value : "", busy, composerState, shareUi });
    if (banner) banner.textContent = st.access.banner || "";
    if (textEl) textEl.disabled = !st.access.canWrite;
    const set = (id, b) => {
      const el = $(id);
      if (!el) return;
      el.hidden = !b.visible;
      el.disabled = !b.enabled;
    };
    set("btn-guest-send", st.buttons.send);
    set("btn-guest-stop", st.buttons.stop);
    set("btn-guest-force", st.buttons.force);
    const sendBtn = $("btn-guest-send");
    if (sendBtn) sendBtn.title = st.buttons.send.title;
    for (const id of ["guest-model-select", "guest-mode-select"]) {
      const el = $(id);
      if (el) el.disabled = !st.selectsEnabled;
    }
    const modelSel = $("guest-model-select");
    if (modelSel && document.activeElement !== modelSel) modelSel.value = openProjection.model || "auto";
    const modeSel = $("guest-mode-select");
    if (modeSel && document.activeElement !== modeSel) modeSel.value = openProjection.mode || "agent";
  }

  async function renderGuestItem(payload) {
    const errEl = $("shared-detail-error-state");
    if (errEl) errEl.hidden = true;
    const same = openItem && openItem.kind === payload.kind && openItem.id === payload.id;
    if (!same) {
      openItem = { kind: payload.kind, id: payload.id };
      openEtag = null;
      openProjection = null;
      const c = $("shared-messages");
      if (c) c.innerHTML = "";
    }
    try {
      const r = await fetchItem(openItem.kind, openItem.id, openEtag);
      if (r.changed) {
        openEtag = r.etag;
        openProjection = r.projection;
        if (!openProjection) {
          showItemError("This item is no longer shared with you.");
          const t = $("shared-detail-title");
          if (t) t.textContent = "Not shared";
          applyComposer();
          return;
        }
        const t = $("shared-detail-title");
        if (t) t.textContent = openProjection.title || "(untitled)";
        renderMessages(openProjection);
        const note = $("guest-composer-note");
        if (note && openProjection.lastIngestedOpId && note.dataset.waitingFor === openProjection.lastIngestedOpId) {
          note.hidden = true;
          note.dataset.waitingFor = "";
        }
      }
      applyComposer();
    } catch (err) {
      showItemError(err.message);
    }
    scheduleItemPoll();
  }

  function scheduleItemPoll() {
    if (itemTimer) return;
    itemTimer = setInterval(() => {
      if (document.hidden) return;
      if (!openItem || document.body.dataset.view !== "shared-detail") return;
      renderGuestItem(openItem);
    }, pollMs);
  }

  function leaveItem() {
    if (itemTimer) clearInterval(itemTimer);
    itemTimer = null;
    openItem = null;
    openEtag = null;
    openProjection = null;
  }

  async function sendOp(partial) {
    if (!openItem || openItem.kind !== "session" || !openProjection) return;
    const st = shareUi.deriveGuestChatAccess(openProjection, guestEmail);
    if (!st.canWrite) return;
    busy = true;
    applyComposer();
    const op = { opId: newOpId(Date.now()), ts: Date.now(), ...partial };
    const endpoint = paths.outboxEndpoint(openItem.id);
    try {
      for (let attempt = 0; attempt < 3; attempt++) {
        const { json, etag } = await loadJson(endpoint);
        const next = appendOutboxOp(json, { op, guest: { email: guestEmail, name: account.name || null }, sessionId: openItem.id });
        try {
          await putJson(endpoint, next, etag);
          break;
        } catch (err) {
          if (err.code !== "PRECONDITION_FAILED" || attempt >= 2) throw err;
        }
      }
      const note = $("guest-composer-note");
      if (note) {
        note.hidden = false;
        note.textContent = `Sent — waiting for ${hostName}'s cockpit to pick it up…`;
        note.dataset.waitingFor = op.opId;
      }
      if (partial.type === "enqueue" || partial.type === "force") {
        const textEl = $("guest-composer-text");
        if (textEl) textEl.value = "";
      }
    } catch (err) {
      showItemError(err.message);
    } finally {
      busy = false;
      applyComposer();
    }
  }

  // wiring
  const btnConnect = $("btn-guest-connect");
  if (btnConnect) btnConnect.addEventListener("click", () => connect());
  const textEl = $("guest-composer-text");
  if (textEl) textEl.addEventListener("input", () => applyComposer());
  const sendBtn = $("btn-guest-send");
  if (sendBtn) sendBtn.addEventListener("click", () => {
    const text = textEl ? textEl.value.trim() : "";
    if (text) sendOp({ type: "enqueue", text });
  });
  const forceBtn = $("btn-guest-force");
  if (forceBtn) forceBtn.addEventListener("click", () => {
    const text = textEl ? textEl.value.trim() : "";
    if (text && window.confirm("Stop the running turn and send this now?")) sendOp({ type: "force", text });
  });
  const stopBtn = $("btn-guest-stop");
  if (stopBtn) stopBtn.addEventListener("click", () => {
    if (window.confirm("Stop the running turn?")) sendOp({ type: "stop" });
  });
  if (typeof ctx.populateModelSelect === "function") ctx.populateModelSelect("guest-model-select");
  const modelSel = $("guest-model-select");
  if (modelSel) modelSel.addEventListener("change", () => sendOp({ type: "setModel", model: modelSel.value }));
  const modeSel = $("guest-mode-select");
  if (modeSel) modeSel.addEventListener("change", () => sendOp({ type: "setMode", mode: modeSel.value }));

  return {
    start(hash) {
      const target = shareModel.parseShareLink(hash);
      if (target) pendingTarget = { kind: target.kind, id: target.id };
      ctx.setView("shared-list");
    },
    renderGuestList,
    renderGuestItem,
    leaveItem,
    refreshOnce() {
      if (document.body.dataset.view === "shared-detail" && openItem) renderGuestItem(openItem);
      else if (document.body.dataset.view === "shared-list") renderGuestList();
    },
  };
}
