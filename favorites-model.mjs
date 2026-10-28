// mobile-cockpit / lib / favorites-model.mjs
//
// SPEC-DELTA-2026-10-28-favorites-tab. Pure logic (no I/O, no DOM, no node:
// imports) shared byte-for-byte with pwa/favorites-model.mjs (enforced by
// tests/flows/mobile-cockpit/pwa-favorites-coherence-selftest.mjs).
//
// favorites.json = { schemaVersion, items: [...] }. Each item sits in one of
// three lists (latest | pinned | archive); its position inside a list is its
// position in `items` among same-list entries (manual order). New items are
// appended, so "latest" reads oldest-first / newest at the bottom.

"use strict";

export const FAVORITES_SCHEMA_VERSION = 1;
export const FAVORITE_LISTS = Object.freeze(["latest", "pinned", "archive"]);
export const OLD_AFTER_DAYS = 90;

/** @returns {{schemaVersion: number, items: any[]}} */
export function emptyFavorites() {
  return { schemaVersion: FAVORITES_SCHEMA_VERSION, items: [] };
}

/** Defensive parse: anything unusable collapses to an empty, valid document. */
export function parseFavorites(raw) {
  if (!raw || typeof raw !== "object" || !Array.isArray(raw.items)) return emptyFavorites();
  const items = [];
  const seen = new Set();
  for (const it of raw.items) {
    if (!it || typeof it !== "object") continue;
    if (typeof it.id !== "string" || !it.id || seen.has(it.id)) continue;
    if (typeof it.url !== "string" || !/^https?:\/\//i.test(it.url)) continue;
    seen.add(it.id);
    items.push({
      id: it.id,
      title: typeof it.title === "string" && it.title.trim() ? it.title.trim() : it.url,
      url: it.url,
      kind: typeof it.kind === "string" ? it.kind : kindFromUrl(it.url),
      note: typeof it.note === "string" ? it.note : "",
      addedAt: typeof it.addedAt === "string" ? it.addedAt : null,
      openedAt: typeof it.openedAt === "string" ? it.openedAt : null,
      list: FAVORITE_LISTS.includes(it.list) ? it.list : "latest",
    });
  }
  return { schemaVersion: FAVORITES_SCHEMA_VERSION, items };
}

/** Guess a display kind from a title/path/url extension. */
export function kindFromUrl(s) {
  const m = /\.(docx?|pdf|xlsx?|pptx?|md|txt|html?)(?:[?#]|$)/i.exec(String(s || ""));
  if (!m) return "link";
  const e = m[1].toLowerCase();
  if (e.startsWith("doc")) return "word";
  if (e.startsWith("xls")) return "excel";
  if (e.startsWith("ppt")) return "powerpoint";
  if (e === "pdf") return "pdf";
  return "file";
}

/** @param {{now?: string, rand?: string}} [o] */
export function newItemId(o = {}) {
  const t = Date.parse(o.now || "") || Date.now();
  return `fav-${t.toString(36)}-${(o.rand || Math.random().toString(36).slice(2, 8))}`;
}

/**
 * Append a new item (end of its list). Re-adding an existing URL does not
 * duplicate it -- the existing item is returned unchanged.
 */
export function addItem(doc, item, nowIso) {
  const base = parseFavorites(doc);
  const url = String(item && item.url || "").trim();
  if (!/^https?:\/\//i.test(url)) throw new Error("addItem: url must be http(s)");
  if (base.items.some((x) => x.url === url)) return base;
  const entry = {
    id: item.id || newItemId({ now: nowIso }),
    title: (item.title || "").trim() || url,
    url,
    kind: item.kind || kindFromUrl(item.title || url),
    note: item.note || "",
    addedAt: nowIso,
    openedAt: null,
    list: FAVORITE_LISTS.includes(item.list) ? item.list : "latest",
  };
  return { ...base, items: [...base.items, entry] };
}

/**
 * Move an item to `toList` at position `toIndex` (index among that list's
 * items AFTER removing the item from its old spot; clamped; default = end).
 * Same list = manual reorder.
 */
export function moveItem(doc, id, toList, toIndex) {
  const base = parseFavorites(doc);
  if (!FAVORITE_LISTS.includes(toList)) throw new Error(`moveItem: bad list ${toList}`);
  const item = base.items.find((x) => x.id === id);
  if (!item) return base;
  const rest = base.items.filter((x) => x.id !== id);
  const moved = { ...item, list: toList };
  const inList = rest.filter((x) => x.list === toList);
  const idx = Number.isInteger(toIndex) ? Math.max(0, Math.min(toIndex, inList.length)) : inList.length;
  // Absolute insertion point: before the idx-th same-list item, else after the last one.
  let at;
  if (idx < inList.length) at = rest.indexOf(inList[idx]);
  else if (inList.length) at = rest.indexOf(inList[inList.length - 1]) + 1;
  else at = rest.length;
  const items = [...rest.slice(0, at), moved, ...rest.slice(at)];
  return { ...base, items };
}

export function markOpened(doc, id, nowIso) {
  const base = parseFavorites(doc);
  return { ...base, items: base.items.map((x) => (x.id === id ? { ...x, openedAt: nowIso } : x)) };
}

export function removeItem(doc, id) {
  const base = parseFavorites(doc);
  return { ...base, items: base.items.filter((x) => x.id !== id) };
}

export function renameItem(doc, id, title) {
  const base = parseFavorites(doc);
  const t = String(title || "").trim();
  if (!t) return base;
  return { ...base, items: base.items.map((x) => (x.id === id ? { ...x, title: t } : x)) };
}

/** Items of one list in manual order. */
export function listItems(doc, list) {
  return parseFavorites(doc).items.filter((x) => x.list === list);
}

/** Case-insensitive AND-match of whitespace-separated terms over title/note/kind/url. */
export function filterItems(items, query) {
  const terms = String(query || "").toLowerCase().split(/\s+/).filter(Boolean);
  if (!terms.length) return items;
  return items.filter((x) => {
    const hay = `${x.title} ${x.note} ${x.kind} ${x.url}`.toLowerCase();
    return terms.every((t) => hay.includes(t));
  });
}

/** "Old" = added more than OLD_AFTER_DAYS ago and not opened within that window. Pinned never old. */
export function isOld(item, nowMs) {
  if (!item || item.list === "pinned") return false;
  const cutoff = nowMs - OLD_AFTER_DAYS * 86400000;
  const added = Date.parse(item.addedAt || "");
  const opened = Date.parse(item.openedAt || "");
  const last = Math.max(Number.isFinite(added) ? added : 0, Number.isFinite(opened) ? opened : 0);
  return last > 0 && last < cutoff;
}

/** Short "opened 3d ago" / "never opened" label. */
export function openedLabel(item, nowMs) {
  const t = Date.parse(item && item.openedAt || "");
  if (!Number.isFinite(t)) return "never opened";
  const s = Math.max(0, Math.round((nowMs - t) / 1000));
  if (s < 90) return "opened just now";
  if (s < 5400) return `opened ${Math.round(s / 60)}m ago`;
  if (s < 129600) return `opened ${Math.round(s / 3600)}h ago`;
  return `opened ${Math.round(s / 86400)}d ago`;
}
