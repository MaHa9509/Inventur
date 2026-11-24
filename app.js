'use strict';
/* Büro-Inventur Technik – Offline-PWA
 * Daten: IndexedDB (Store "devices" + Store "meta" mit den Einstellungen)
 * Excel: ExcelJS (lokal in vendor/, wird erst bei Bedarf geladen)
 */

const APP_VERSION = '1.1.0';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const NO_TEAM = 'Gemeinschaftlich/ohne Team';

const STATUSES = ['In Benutzung', 'Lager/Reserve', 'Defekt', 'Auszusondern', 'Verliehen'];

const DEFAULT_TYPES = ['Laptop', 'PC', 'Monitor', 'Dockingstation', 'Drucker', 'Telefon', 'Smartphone',
  'Tablet', 'Tastatur', 'Maus', 'Headset', 'Webcam', 'Beamer', 'Netzwerkgerät', 'Sonstiges'];

const DEFAULT_TEAMS = [
  { name: 'Team 1', color: '#FBD49A' },
  { name: 'Team 2', color: '#DFF5C4' },
  { name: 'Team 3', color: '#D6E2EE' },
  { name: 'Team 4', color: '#DCD6F7' },
  { name: NO_TEAM, color: '#E6E6E6' },
];

// Draufsicht: Koordinaten aus dem Raumplan (x, y, Breite, Höhe in Plan-Einheiten)
const PLAN = { w: 949, h: 316 };
// gestrichelte Fläche ohne Raumnummer (nicht anklickbar)
const PLAN_DECOR = [{ x: 511, y: 0, w: 72, h: 121 }];

const DEFAULT_ROOMS = [
  { name: '4180', team: 'Team 1', x: 0, y: 0, w: 72, h: 121 },
  { name: '4190', team: 'Team 2', x: 72, y: 0, w: 73, h: 121 },
  { name: '4200', team: NO_TEAM, x: 145, y: 0, w: 73, h: 121 },
  { name: '4220', team: 'Team 1', x: 218, y: 0, w: 146, h: 121 },
  { name: '4240', team: 'Team 1', x: 364, y: 0, w: 147, h: 121 },
  { name: '4260', team: 'Team 3', x: 583, y: 0, w: 73, h: 121 },
  { name: '4270', team: 'Team 3', x: 656, y: 0, w: 73, h: 121 },
  { name: '4280', team: 'Team 3', x: 729, y: 0, w: 73, h: 121 },
  { name: '4290', team: 'Team 3', x: 802, y: 0, w: 73, h: 121 },
  { name: '4300', team: 'Team 3', x: 875, y: 0, w: 74, h: 121 },
  { name: '4170', team: 'Team 2', x: 0, y: 195, w: 121, h: 121 },
  { name: '4160', team: 'Team 2', x: 121, y: 195, w: 122, h: 121 },
  { name: '4150/4140', team: NO_TEAM, x: 243, y: 195, w: 97, h: 121, dashed: true },
  { name: '4120', team: NO_TEAM, x: 413, y: 195, w: 73, h: 121, dashed: true },
  { name: '4110', team: 'Team 1', x: 486, y: 195, w: 73, h: 121 },
  { name: '4100/4090', team: NO_TEAM, x: 559, y: 195, w: 195, h: 121 },
  { name: '4080', team: 'Team 4', x: 754, y: 195, w: 73, h: 121 },
  { name: '4070', team: 'Team 4', x: 827, y: 195, w: 122, h: 121 },
  { name: 'Lager / Sonstiges', team: NO_TEAM, dashed: true },
];
const inPlan = r => r && Number.isFinite(r.x) && Number.isFinite(r.y) && r.w > 0 && r.h > 0;

// Spalten der Excel-Tabelle "Inventar" (Reihenfolge = Spaltenreihenfolge)
const FIELDS = [
  { key: 'id', label: 'ID', width: 10 },
  { key: 'typ', label: 'Gerätetyp', width: 16 },
  { key: 'hersteller', label: 'Hersteller', width: 16 },
  { key: 'modell', label: 'Modell', width: 24 },
  { key: 'seriennr', label: 'Seriennummer', width: 22 },
  { key: 'altInvNr', label: 'Alte Inventarnummer', width: 19 },
  { key: 'raum', label: 'Raum', width: 16 },
  { key: 'arbeitsplatz', label: 'Arbeitsplatz', width: 16 },
  { key: 'person', label: 'Person', width: 20 },
  { key: 'team', label: 'Team', width: 24 },
  { key: 'status', label: 'Status', width: 15 },
  { key: 'kaufdatum', label: 'Kaufdatum', width: 12, kind: 'date' },
  { key: 'garantie', label: 'Garantie bis', width: 12, kind: 'date' },
  { key: 'anmerkungen', label: 'Anmerkungen', width: 40 },
  { key: 'erfasstVon', label: 'Erfasst von', width: 16 },
  { key: 'erfasstAm', label: 'Erfasst am', width: 17, kind: 'datetime' },
  { key: 'geaendertAm', label: 'Zuletzt geändert', width: 17, kind: 'datetime' },
];
const REC_KEYS = FIELDS.map(f => f.key);
const COMPARE_KEYS = REC_KEYS.filter(k => k !== 'erfasstAm' && k !== 'geaendertAm');
const FORM_KEYS = ['typ', 'hersteller', 'modell', 'seriennr', 'altInvNr', 'raum', 'arbeitsplatz', 'person',
  'team', 'status', 'kaufdatum', 'garantie', 'anmerkungen'];

/* ---------------- Hilfsfunktionen ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n, l = 2) => String(n).padStart(l, '0');
const clone = o => JSON.parse(JSON.stringify(o));

function todayYmd(d = new Date()) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function nowIso() { const d = new Date(); d.setMilliseconds(0); return d.toISOString(); }
function ts(iso) { const t = Date.parse(iso || ''); return isNaN(t) ? 0 : t; }
function fmtDate(ymd) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd || ''); return m ? `${m[3]}.${m[2]}.${m[1]}` : ''; }
function fmtDateTime(iso) {
  const t = ts(iso); if (!t) return '–';
  const d = new Date(t);
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function idParts(id) { const m = /^([A-Z]+)-(\d+)$/.exec(id || ''); return m ? { p: m[1], n: +m[2] } : null; }
function cmpId(a, b) {
  const x = idParts(a), y = idParts(b);
  if (x && y) return x.p === y.p ? x.n - y.n : x.p.localeCompare(y.p);
  return String(a).localeCompare(String(b), 'de', { numeric: true });
}
const IGNORED_SERIALS = new Set(['', 'NA', 'NV', 'KEINE', 'KEIN', 'UNBEKANNT', 'OHNE', 'X', '0']);
function normSerial(s) { const k = String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); return IGNORED_SERIALS.has(k) ? '' : k; }
function normRec(r) {
  const o = {};
  for (const k of REC_KEYS) { const v = r[k]; o[k] = v == null ? '' : String(v).trim(); }
  o.id = o.id.toUpperCase();
  return o;
}
function sameContent(a, b) { return COMPARE_KEYS.every(k => (a[k] || '') === (b[k] || '')); }
function teamColor(name) { const t = S.teams.find(t => t.name === name); return t ? t.color : '#E6E6E6'; }
function roomByName(name) { return S.rooms.find(r => r.name === name); }
function plural(n, one, many) { return `${n} ${n === 1 ? one : many}`; }
function statusClass(s) {
  return { 'Defekt': 's-defekt', 'Auszusondern': 's-auszusondern', 'Lager/Reserve': 's-lager', 'Verliehen': 's-verliehen' }[s] || '';
}

/* ---------------- IndexedDB ---------------- */
const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('buero-inventur', 1);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains('devices')) db.createObjectStore('devices', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta', { keyPath: 'key' });
      };
      r.onsuccess = () => { this.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  run(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = this.db.transaction(store, mode);
      let out;
      const req = fn(t.objectStore(store));
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => res(out);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error || new Error('Speichern abgebrochen'));
    });
  },
  getAll(store) { return this.run(store, 'readonly', s => s.getAll()); },
  get(store, key) { return this.run(store, 'readonly', s => s.get(key)); },
  put(store, v) { return this.run(store, 'readwrite', s => s.put(v)); },
  putMany(store, arr) { return this.run(store, 'readwrite', s => { arr.forEach(v => s.put(v)); }); },
  del(store, key) { return this.run(store, 'readwrite', s => s.delete(key)); },
  clear(store) { return this.run(store, 'readwrite', s => s.clear()); },
};

/* ---------------- Zustand ---------------- */
let S = null;                 // Einstellungen
const devices = new Map();    // id -> Datensatz

function defaultSettings() {
  return {
    key: 'settings', prefix: '', name: '', counters: {},
    types: [...DEFAULT_TYPES], teams: clone(DEFAULT_TEAMS), rooms: clone(DEFAULT_ROOMS),
    roomsDone: {}, lastExport: null, lastBackup: null,
  };
}
async function saveSettings() { await DB.put('meta', S); }

function nextId() {
  const p = S.prefix;
  let max = S.counters[p] || 0;
  for (const id of devices.keys()) { const x = idParts(id); if (x && x.p === p) max = Math.max(max, x.n); }
  return { num: max + 1, id: `${p}-${pad(max + 1, 4)}` };
}

/* ---------------- UI-Bausteine ---------------- */
let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
}

/** Zeigt einen Dialog. buttons: [{label, value, cls, action}] – action() darf false liefern, um offen zu bleiben. */
function modal({ title = '', body = '', buttons = [{ label: 'OK', value: true, cls: 'primary' }], dismissable = true, onOpen } = {}) {
  return new Promise(resolve => {
    const root = $('#modal-root');
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = `<div class="modal" role="dialog" aria-modal="true">
      ${title ? `<h3>${esc(title)}</h3>` : ''}
      <div class="modal-body">${body}</div>
      <div class="modal-actions">${buttons.map((b, i) => `<button type="button" class="btn block ${b.cls || ''}" data-i="${i}">${esc(b.label)}</button>`).join('')}</div>
    </div>`;
    const close = v => { back.remove(); document.removeEventListener('keydown', onKey); resolve(v); };
    const onKey = e => { if (e.key === 'Escape' && dismissable) close(null); };
    back.addEventListener('click', async e => {
      if (e.target === back && dismissable) return close(null);
      const btn = e.target.closest('.modal-actions button');
      if (!btn) return;
      const b = buttons[+btn.dataset.i];
      if (b.action) {
        const r = await b.action(back);
        if (r === false) return;
      }
      close(b.value !== undefined ? b.value : b.label);
    });
    document.addEventListener('keydown', onKey);
    root.appendChild(back);
    if (onOpen) onOpen(back, close);
  });
}
const confirmModal = (title, body, okLabel = 'OK', danger = false) => modal({
  title, body,
  buttons: [{ label: okLabel, value: true, cls: danger ? 'danger-fill' : 'primary' }, { label: 'Abbrechen', value: false }],
});
const infoModal = (title, body) => modal({ title, body });

let busyEl = null;
function busy(msg) {
  if (busyEl) { busyEl.remove(); busyEl = null; }
  if (!msg) return;
  busyEl = document.createElement('div');
  busyEl.className = 'modal-backdrop';
  busyEl.innerHTML = `<div class="modal" style="text-align:center"><div class="spinner"></div><div>${esc(msg)}</div></div>`;
  document.body.appendChild(busyEl);
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.rel = 'noopener';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

/** Bietet eine fertige Datei zum Teilen (Mail, AirDrop, Teams …) oder Herunterladen an. */
async function offerFile(blob, filename, onDone) {
  const file = new File([blob], filename, { type: blob.type });
  let canShare = false;
  try { canShare = !!(navigator.canShare && navigator.canShare({ files: [file] })); } catch (_) { canShare = false; }
  const buttons = [];
  if (canShare) {
    buttons.push({
      label: 'Teilen (Mail, AirDrop, Teams …)', cls: 'primary', value: 'share',
      action: async () => {
        try { await navigator.share({ files: [file], title: filename }); if (onDone) onDone(); }
        catch (e) { if (e.name !== 'AbortError') { download(blob, filename); if (onDone) onDone(); } else return false; }
      },
    });
  }
  buttons.push({ label: 'Herunterladen / In Dateien sichern', cls: canShare ? '' : 'primary', value: 'download', action: () => { download(blob, filename); if (onDone) onDone(); } });
  buttons.push({ label: 'Schließen', value: null });
  return modal({
    title: 'Datei ist fertig',
    body: `<p><strong>${esc(filename)}</strong></p><p class="muted small">Am besten direkt per Mail/Teams verschicken oder in „Dateien“ bzw. OneDrive sichern.</p>`,
    buttons,
  });
}

let excelPromise = null;
function loadExcel() {
  if (window.ExcelJS) return Promise.resolve();
  if (!excelPromise) {
    excelPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'vendor/exceljs.min.js';
      s.onload = () => res();
      s.onerror = () => { excelPromise = null; rej(new Error('Die Excel-Bibliothek konnte nicht geladen werden.')); };
      document.head.appendChild(s);
    });
  }
  return excelPromise;
}

/* ---------------- Navigation ---------------- */
const VIEWS = ['setup', 'list', 'rooms', 'export', 'settings', 'form'];
const TITLES = { setup: 'Einrichtung', list: 'Geräte', rooms: 'Räume', export: 'Excel-Export & Import', settings: 'Einstellungen', form: 'Gerät' };
let currentView = null;
let listScroll = 0;
let formFromApp = false;
let pendingPrefill = null;

function showView(name) {
  if (currentView === 'list' && name !== 'list') listScroll = window.scrollY;
  VIEWS.forEach(v => { $('#view-' + v).hidden = v !== name; });
  $$('.tabbar a').forEach(a => a.classList.toggle('active', a.dataset.view === name));
  const chrome = name !== 'setup' && name !== 'form';
  $('#tabbar').hidden = !chrome;
  document.body.classList.toggle('no-tabbar', !chrome);
  $('#fab').hidden = !(name === 'list' || name === 'rooms');
  $('#title').textContent = TITLES[name] || 'Büro-Inventur';
  const badge = $('#device-badge');
  badge.hidden = !S.prefix;
  badge.textContent = S.prefix ? `Handy ${S.prefix}${S.name ? ' · ' + S.name : ''}` : '';
  const prev = currentView;
  currentView = name;
  window.scrollTo(0, name === 'list' && prev !== 'list' ? listScroll : 0);
}

function route() {
  if (!S.prefix) { showView('setup'); return; }
  const h = decodeURIComponent(location.hash.replace(/^#/, ''));
  if (h === 'neu') return openForm(null);
  if (h.startsWith('edit/')) return openForm(h.slice(5));
  if (h === 'raeume') { renderRooms(); return showView('rooms'); }
  if (h === 'export') { renderExport(); return showView('export'); }
  if (h === 'einstellungen') { renderSettings(); return showView('settings'); }
  renderList(); showView('list');
}

function goNew(prefill) {
  pendingPrefill = prefill || null;
  if (currentView !== 'form') formFromApp = true;
  if (location.hash === '#neu') route();
  else if (currentView === 'form') location.replace('#neu');
  else location.hash = '#neu';
}
function goEdit(id) {
  formFromApp = true;
  location.hash = '#edit/' + encodeURIComponent(id);
}
function leaveForm() {
  formDirty = false;
  if (formFromApp) { formFromApp = false; history.back(); }
  else location.replace('#liste');
}

/* ---------------- Liste ---------------- */
function serialGroups(list) {
  const m = new Map();
  for (const r of list) {
    const k = normSerial(r.seriennr);
    if (!k) continue;
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return [...m.values()].filter(g => g.length > 1).map(g => g.sort((a, b) => cmpId(a.id, b.id)));
}
function dupMap(list) {
  const map = new Map();
  for (const g of serialGroups(list)) for (const r of g) map.set(r.id, g.filter(x => x.id !== r.id).map(x => x.id));
  return map;
}

function fillSelect(sel, options, value, placeholder) {
  const opts = options.map(o => typeof o === 'string' ? { value: o, label: o } : o);
  if (value && !opts.some(o => o.value === value)) opts.push({ value, label: value });
  sel.innerHTML = (placeholder != null ? `<option value="">${esc(placeholder)}</option>` : '') +
    opts.map(o => `<option value="${esc(o.value)}">${esc(o.label)}</option>`).join('');
  sel.value = value || '';
}
function uniqueValues(key) {
  return [...new Set([...devices.values()].map(d => d[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
}

function lastSafeCopy() { return Math.max(ts(S.lastExport), ts(S.lastBackup)); }

function renderBanner() {
  const b = $('#backup-banner');
  const last = lastSafeCopy();
  const old = !last || Date.now() - last > 24 * 3600 * 1000;
  if (devices.size && old) {
    b.innerHTML = `<strong>Daten sichern!</strong> Die Daten liegen nur auf diesem Handy. ${last ? `Letzter Export/Backup: ${fmtDateTime(new Date(last).toISOString())}.` : 'Es wurde noch nie exportiert.'}
      <button type="button" class="btn small block" data-go="export">Jetzt als Excel exportieren</button>`;
    b.hidden = false;
  } else b.hidden = true;
}

function renderList() {
  renderBanner();
  const all = [...devices.values()];
  const keep = id => $(id).value;
  const rooms = [...new Set([...S.rooms.map(r => r.name), ...uniqueValues('raum')])];
  const teams = [...new Set([...S.teams.map(t => t.name), ...uniqueValues('team')])];
  const types = [...new Set([...S.types, ...uniqueValues('typ')])];
  fillSelect($('#flt-raum'), rooms, keep('#flt-raum'), 'Alle Räume');
  fillSelect($('#flt-team'), teams, keep('#flt-team'), 'Alle Teams');
  fillSelect($('#flt-typ'), types, keep('#flt-typ'), 'Alle Gerätetypen');
  fillSelect($('#flt-status'), STATUSES, keep('#flt-status'), 'Alle Status');

  const f = { raum: keep('#flt-raum'), team: keep('#flt-team'), typ: keep('#flt-typ'), status: keep('#flt-status') };
  const activeFilters = Object.values(f).filter(Boolean).length;
  $('#filter-toggle').textContent = activeFilters ? `Filter (${activeFilters})` : 'Filter';
  $('#filter-toggle').classList.toggle('on', activeFilters > 0);

  const q = $('#search').value.trim().toLowerCase();
  const words = q ? q.split(/\s+/) : [];
  const list = all.filter(d => {
    for (const k in f) if (f[k] && d[k] !== f[k]) return false;
    if (!words.length) return true;
    const hay = REC_KEYS.map(k => d[k]).join(' ').toLowerCase();
    return words.every(w => hay.includes(w));
  }).sort((a, b) => (ts(b.erfasstAm) - ts(a.erfasstAm)) || cmpId(b.id, a.id));

  const dups = dupMap(all);
  $('#list-count').textContent = list.length === all.length
    ? plural(all.length, 'Gerät', 'Geräte')
    : `${list.length} von ${plural(all.length, 'Gerät', 'Geräten')}`;

  $('#device-list').innerHTML = list.map(d => {
    const title = [d.hersteller, d.modell].filter(Boolean).join(' ');
    const sub1 = [d.raum ? 'Raum ' + d.raum : '', d.arbeitsplatz, d.person].filter(Boolean).join(' · ');
    const sub2 = [d.seriennr ? 'SN ' + d.seriennr : '', d.team && d.team !== NO_TEAM ? d.team : ''].filter(Boolean).join(' · ');
    return `<li><button type="button" class="item" data-edit="${esc(d.id)}" style="--team:${esc(teamColor(d.team))}">
      <div class="item-main">
        <div class="item-top"><span class="item-id">${esc(d.id)}</span><span class="item-title">${esc(d.typ)}${title ? ' – ' + esc(title) : ''}</span></div>
        ${sub1 ? `<div class="item-sub">${esc(sub1)}</div>` : ''}
        ${sub2 ? `<div class="item-sub">${esc(sub2)}</div>` : ''}
      </div>
      <div class="stack" style="align-items:flex-end">
        <span class="chip ${statusClass(d.status)}">${esc(d.status)}</span>
        ${dups.has(d.id) ? `<span class="chip dup" title="Seriennummer doppelt">SN doppelt</span>` : ''}
      </div>
    </button></li>`;
  }).join('');

  const empty = $('#list-empty');
  if (!all.length) {
    empty.innerHTML = `<p><strong>Noch keine Geräte erfasst.</strong></p><p>Tippe auf <strong>+</strong>, um das erste Gerät zu erfassen, oder wähle unter „Räume“ einen Raum aus.</p>`;
    empty.hidden = false;
  } else if (!list.length) {
    empty.innerHTML = '<p>Keine Treffer für diese Suche/Filter.</p>';
    empty.hidden = false;
  } else empty.hidden = true;
}

/* ---------------- Räume / Grundriss ---------------- */
function roomCounts() {
  const c = {};
  for (const d of devices.values()) c[d.raum] = (c[d.raum] || 0) + 1;
  return c;
}
function planHtml({ selected = '', showDone = true } = {}) {
  const counts = roomCounts();
  const pos = r => `left:${r.x / PLAN.w * 100}%;top:${r.y / PLAN.h * 100}%;width:${r.w / PLAN.w * 100}%;height:${r.h / PLAN.h * 100}%`;
  const tile = (r, style, compact) => {
    const n = counts[r.name] || 0;
    const done = showDone && S.roomsDone[r.name];
    return `<button type="button" class="tile ${r.dashed ? 'dashed' : ''} ${selected === r.name ? 'selected' : ''}"
      style="--c:${esc(teamColor(r.team))};${style}" data-room="${esc(r.name)}"
      aria-label="Raum ${esc(r.name)}, ${n} Geräte${done ? ', fertig erfasst' : ''}">
      <span class="tile-name">${esc(r.name)}</span>
      <span class="tile-count">${n ? (compact ? String(n) : plural(n, 'Gerät', 'Geräte')) : '–'}</span>
      ${done ? '<span class="done-mark">✓</span>' : ''}
    </button>`;
  };
  const placed = S.rooms.filter(inPlan);
  const others = S.rooms.filter(r => !inPlan(r));
  let html = '';
  if (placed.length) {
    html += `<div class="plan-scroll"><div class="plan-map">
      ${PLAN_DECOR.map(d => `<div class="plan-decor" style="${pos(d)}"></div>`).join('')}
      ${placed.map(r => tile(r, pos(r), true)).join('')}
    </div></div>`;
  }
  if (others.length) {
    html += `<div><div class="plan-row-label">Weitere Orte</div><div class="plan-tiles">${others.map(r => tile(r, '', false)).join('')}</div></div>`;
  }
  return html;
}
function legendHtml() {
  return S.teams.map(t => `<span><i style="background:${esc(t.color)}"></i>${esc(t.name)}</span>`).join('') +
    `<span><i style="background:#fff;border-style:dashed"></i>gestrichelt im Plan</span>`;
}
function renderRooms() {
  const total = S.rooms.length;
  const done = S.rooms.filter(r => S.roomsDone[r.name]).length;
  const pct = total ? Math.round(done / total * 100) : 0;
  $('#rooms-progress').innerHTML = `<div><strong>${done} von ${total} Räumen fertig erfasst</strong> · ${plural(devices.size, 'Gerät', 'Geräte')}</div>
    <div class="progress-bar"><div style="width:${pct}%"></div></div>`;
  $('#plan').innerHTML = planHtml();
  $('#legend').innerHTML = legendHtml();
}

async function openRoomSheet(name) {
  const room = roomByName(name);
  const list = [...devices.values()].filter(d => d.raum === name).sort((a, b) => cmpId(a.id, b.id));
  const done = !!S.roomsDone[name];
  const team = room ? room.team : '';
  const body = `<p class="muted">${esc(team || '')}${done ? ' · <strong style="color:var(--ok)">fertig erfasst ✓</strong>' : ''}</p>
    ${list.length ? `<ul class="room-devices">${list.map(d => `<li><button type="button" class="item" data-edit-room="${esc(d.id)}" style="--team:${esc(teamColor(d.team))}">
      <div class="item-main"><div class="item-top"><span class="item-id">${esc(d.id)}</span><span class="item-title">${esc(d.typ)} ${esc([d.hersteller, d.modell].filter(Boolean).join(' '))}</span></div>
      <div class="item-sub">${esc([d.arbeitsplatz, d.person].filter(Boolean).join(' · '))}</div></div>
      <span class="chip ${statusClass(d.status)}">${esc(d.status)}</span></button></li>`).join('')}</ul>`
      : '<p>Noch keine Geräte in diesem Raum.</p>'}`;
  const res = await modal({
    title: `Raum ${name}`, body,
    buttons: [
      { label: '+ Gerät in diesem Raum erfassen', cls: 'primary', value: 'add' },
      { label: done ? 'Markierung „fertig“ entfernen' : 'Raum als fertig erfasst markieren', value: 'toggle' },
      { label: 'Schließen', value: null },
    ],
    onOpen: (el, close) => {
      el.addEventListener('click', e => {
        const b = e.target.closest('[data-edit-room]');
        if (b) { close(null); goEdit(b.dataset.editRoom); }
      });
    },
  });
  if (res === 'add') goNew({ raum: name, team: team || '', status: 'In Benutzung' });
  if (res === 'toggle') {
    if (done) delete S.roomsDone[name]; else S.roomsDone[name] = nowIso();
    await saveSettings();
    renderRooms();
    toast(done ? `Raum ${name}: Markierung entfernt` : `Raum ${name} als fertig markiert`);
  }
}

function pickRoomModal(current) {
  return modal({
    title: 'Raum auswählen',
    body: `<div class="plan">${planHtml({ selected: current, showDone: false })}</div>`,
    buttons: [{ label: 'Abbrechen', value: null }],
    onOpen: (el, close) => {
      el.addEventListener('click', e => {
        const t = e.target.closest('[data-room]');
        if (t) close(t.dataset.room);
      });
    },
  });
}

/* ---------------- Formular ---------------- */
let formId = null;
let formDirty = false;

function setDatalist(id, values) { $(id).innerHTML = values.map(v => `<option value="${esc(v)}"></option>`).join(''); }

function openForm(id) {
  const rec = id ? devices.get(id) : null;
  if (id && !rec) { toast('Eintrag nicht gefunden'); location.replace('#liste'); return; }
  formId = id;
  const pre = rec || Object.assign({ status: 'In Benutzung' }, pendingPrefill || {});
  pendingPrefill = null;

  fillSelect($('#f-typ'), S.types, pre.typ, 'Bitte wählen …');
  fillSelect($('#f-raum'), S.rooms.map(r => ({ value: r.name, label: r.team && r.team !== NO_TEAM ? `${r.name} (${r.team})` : r.name })), pre.raum, 'Bitte wählen …');
  fillSelect($('#f-team'), S.teams.map(t => t.name), pre.team, '–');
  fillSelect($('#f-status'), STATUSES, pre.status, 'Bitte wählen …');
  for (const k of ['hersteller', 'modell', 'seriennr', 'altInvNr', 'arbeitsplatz', 'person', 'kaufdatum', 'garantie', 'anmerkungen']) {
    $('#f-' + k).value = pre[k] || '';
  }
  setDatalist('#dl-hersteller', uniqueValues('hersteller'));
  setDatalist('#dl-modell', uniqueValues('modell'));
  setDatalist('#dl-arbeitsplatz', uniqueValues('arbeitsplatz'));
  setDatalist('#dl-person', uniqueValues('person'));
  $$('#device-form .invalid').forEach(e => e.classList.remove('invalid'));

  $('#form-id').textContent = id ? id : `Neues Gerät · ID ${nextId().id}`;
  $('#form-meta').textContent = rec ? `Erfasst von ${rec.erfasstVon || '–'} am ${fmtDateTime(rec.erfasstAm)} · zuletzt geändert ${fmtDateTime(rec.geaendertAm)}` : 'Die ID wird beim Speichern vergeben.';
  $('#form-delete').hidden = !id;
  checkSerial();
  formDirty = false;
  showView('form');
  TITLES.form = id ? 'Gerät bearbeiten' : 'Gerät erfassen';
  $('#title').textContent = TITLES.form;
}

function readForm() {
  const o = {};
  for (const k of FORM_KEYS) o[k] = $('#f-' + k).value.trim();
  return o;
}

function checkSerial() {
  const k = normSerial($('#f-seriennr').value);
  const w = $('#serial-warning');
  if (!k) { w.hidden = true; return; }
  const hits = [...devices.values()].filter(d => d.id !== formId && normSerial(d.seriennr) === k);
  if (!hits.length) { w.hidden = true; return; }
  w.innerHTML = `⚠️ Diese Seriennummer existiert bereits: ${hits.map(d => `<strong>${esc(d.id)}</strong> (${esc(d.typ)}, Raum ${esc(d.raum)})`).join(', ')}. Wurde das Gerät schon erfasst?`;
  w.hidden = false;
}

async function saveForm(next) {
  const f = readForm();
  let bad = null;
  for (const k of ['typ', 'raum', 'status']) {
    const el = $('#f-' + k);
    el.classList.toggle('invalid', !f[k]);
    if (!f[k] && !bad) bad = el;
  }
  if (bad) { bad.scrollIntoView({ block: 'center', behavior: 'smooth' }); toast('Bitte Gerätetyp, Raum und Status angeben.'); return; }

  let rec;
  const isNew = !formId;
  try {
    if (!isNew) {
      const old = devices.get(formId);
      rec = normRec({ ...old, ...f });
      if (sameContent(rec, old)) { formDirty = false; if (!next) { leaveForm(); return; } }
      else rec.geaendertAm = nowIso();
    } else {
      const { id, num } = nextId();
      const now = nowIso();
      rec = normRec({ ...f, id, erfasstVon: S.name, erfasstAm: now, geaendertAm: now });
      S.counters[S.prefix] = num;
      await saveSettings();
    }
    await DB.put('devices', rec);
    devices.set(rec.id, rec);
  } catch (e) {
    infoModal('Speichern fehlgeschlagen', `<p>${esc(e.message || e)}</p>`);
    return;
  }
  formDirty = false;

  if (isNew) {
    await modal({
      title: 'Gespeichert',
      body: `<p>Bitte diese ID auf das Gerät schreiben:</p><div class="big-id">${esc(rec.id)}</div>
        <p class="muted small">${esc(rec.typ)}${rec.hersteller || rec.modell ? ' – ' + esc([rec.hersteller, rec.modell].filter(Boolean).join(' ')) : ''} · Raum ${esc(rec.raum)}</p>`,
      buttons: [{ label: next ? 'Weiter zum nächsten Gerät' : 'OK', cls: 'primary', value: true }],
      dismissable: false,
    });
  } else toast(`${rec.id} gespeichert`);

  if (next) goNew({ raum: rec.raum, arbeitsplatz: rec.arbeitsplatz, person: rec.person, team: rec.team, status: 'In Benutzung' });
  else leaveForm();
}

async function deleteCurrent() {
  const rec = devices.get(formId);
  if (!rec) return;
  const ok = await confirmModal('Gerät löschen?', `<p><strong>${esc(rec.id)}</strong> – ${esc(rec.typ)} ${esc(rec.hersteller)} ${esc(rec.modell)} wird gelöscht.</p><p class="muted small">Die ID wird nicht erneut vergeben.</p>`, 'Endgültig löschen', true);
  if (!ok) return;
  await DB.del('devices', rec.id);
  devices.delete(rec.id);
  toast(`${rec.id} gelöscht`);
  leaveForm();
}

/* ---------------- Excel-Export ---------------- */
function xlDate(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd || '');
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
}
function xlDateTime(iso) {
  const t = ts(iso); if (!t) return null;
  const d = new Date(t);
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()));
}
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
const argb = hex => 'FF' + String(hex || '#E6E6E6').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);
const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F5FA8' } };
const SUB_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F2' } };
const thin = { style: 'thin', color: { argb: 'FFD6DCE4' } };

function styleHeader(row) {
  row.height = 22;
  row.eachCell(c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = HEAD_FILL;
    c.alignment = { vertical: 'middle' };
  });
}

async function buildWorkbook(list) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Büro-Inventur';
  wb.created = new Date();
  list = [...list].sort((a, b) => cmpId(a.id, b.id));
  const dups = dupMap(list);
  const colOf = key => colLetter(FIELDS.findIndex(f => f.key === key) + 1);

  // --- Inventar ---
  const ws = wb.addWorksheet('Inventar', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [...FIELDS.map(f => ({ header: f.label, key: f.key, width: f.width })), { header: 'Prüfhinweis', key: 'pruef', width: 38 }];
  styleHeader(ws.getRow(1));
  for (const d of list) {
    const vals = {};
    for (const f of FIELDS) {
      if (f.kind === 'date') vals[f.key] = xlDate(d[f.key]);
      else if (f.kind === 'datetime') vals[f.key] = xlDateTime(d[f.key]);
      else vals[f.key] = d[f.key] || null;
    }
    if (dups.has(d.id)) vals.pruef = `Seriennummer auch bei ${dups.get(d.id).join(', ')}`;
    const row = ws.addRow(vals);
    FIELDS.forEach((f, i) => {
      const c = row.getCell(i + 1);
      if (f.kind === 'date') c.numFmt = 'dd.mm.yyyy';
      if (f.kind === 'datetime') c.numFmt = 'dd.mm.yyyy hh:mm';
      if (typeof c.value === 'string') c.numFmt = '@';
    });
    row.getCell('anmerkungen').alignment = { wrapText: true, vertical: 'top' };
    if (d.team) row.getCell('team').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(teamColor(d.team)) } };
    if (d.status === 'Defekt') row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
    if (d.status === 'Auszusondern') row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE5CC' } };
    if (vals.pruef) row.getCell('pruef').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
  }
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: FIELDS.length + 1 } };

  // --- Übersicht ---
  const ov = wb.addWorksheet('Übersicht');
  ov.columns = [{ width: 34 }, { width: 14 }, { width: 16 }];
  ov.addRow(['Büro-Inventur Technik']).font = { bold: true, size: 16 };
  const stand = ov.addRow(['Stand', xlDateTime(nowIso())]);
  stand.getCell(2).numFmt = 'dd.mm.yyyy hh:mm';
  stand.getCell(2).alignment = { horizontal: 'left' };
  ov.addRow(['Geräte gesamt', { formula: 'COUNTA(Inventar!A:A)-1', result: list.length }]).font = { bold: true };
  const roomsDone = S.rooms.filter(r => S.roomsDone[r.name]).length;
  ov.addRow(['Räume fertig erfasst', `${roomsDone} von ${S.rooms.length}`]).font = { bold: true };
  const prefixes = [...new Set(list.map(d => (idParts(d.id) || {}).p).filter(Boolean))].sort();
  ov.addRow(['Erfasst mit Handys', prefixes.join(', ') || '–']);

  const crit = v => String(v).replace(/"/g, '""');
  const count = (key, v) => list.filter(d => d[key] === v).length;
  const section = (title, key, values, extra) => {
    ov.addRow([]);
    const h = ov.addRow(extra ? [title, 'Anzahl', extra.label] : [title, 'Anzahl']);
    h.eachCell(c => { c.font = { bold: true }; c.fill = SUB_FILL; c.border = { bottom: thin }; });
    for (const v of values) {
      const col = colOf(key);
      const r = ov.addRow([v, { formula: `COUNTIF(Inventar!${col}:${col},"${crit(v)}")`, result: count(key, v) }]);
      r.getCell(1).numFmt = '@';
      if (extra) r.getCell(3).value = extra.value(v);
    }
  };
  const present = key => [...new Set(list.map(d => d[key]).filter(Boolean))];
  section('Nach Gerätetyp', 'typ', [...new Set([...S.types.filter(t => count('typ', t)), ...present('typ')])]);
  section('Nach Raum', 'raum', [...new Set([...S.rooms.map(r => r.name), ...present('raum')])],
    { label: 'Fertig erfasst', value: v => (S.roomsDone[v] ? 'Ja' : 'Nein') });
  section('Nach Team', 'team', [...new Set([...S.teams.map(t => t.name), ...present('team')])]);
  section('Nach Status', 'status', [...new Set([...STATUSES, ...present('status')])]);

  // --- Räume (auch für den Re-Import) ---
  const rs = wb.addWorksheet('Räume', { views: [{ state: 'frozen', ySplit: 1 }] });
  rs.columns = [
    { header: 'Raum', key: 'raum', width: 18 }, { header: 'Team', key: 'team', width: 26 },
    { header: 'Bereich', key: 'reihe', width: 14 }, { header: 'Geräte', key: 'anzahl', width: 10 },
    { header: 'Fertig erfasst', key: 'fertig', width: 14 },
  ];
  styleHeader(rs.getRow(1));
  const raumCol = colOf('raum');
  for (const r of S.rooms) {
    const row = rs.addRow({
      raum: r.name, team: r.team || '', reihe: inPlan(r) ? 'Grundriss' : 'Weitere Orte',
      anzahl: { formula: `COUNTIF(Inventar!${raumCol}:${raumCol},"${crit(r.name)}")`, result: count('raum', r.name) },
      fertig: S.roomsDone[r.name] ? 'Ja' : 'Nein',
    });
    row.getCell('raum').numFmt = '@';
    row.getCell('team').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(teamColor(r.team)) } };
  }

  // --- Prüfen ---
  const groups = serialGroups(list);
  if (groups.length) {
    const ps = wb.addWorksheet('Prüfen', { views: [{ state: 'frozen', ySplit: 1 }] });
    ps.columns = [{ header: 'Seriennummer', key: 'sn', width: 24 }, { header: 'IDs', key: 'ids', width: 22 }, { header: 'Geräte', key: 'det', width: 70 }];
    styleHeader(ps.getRow(1));
    for (const g of groups) {
      const row = ps.addRow({
        sn: g[0].seriennr, ids: g.map(d => d.id).join(', '),
        det: g.map(d => `${d.id}: ${d.typ} ${[d.hersteller, d.modell].filter(Boolean).join(' ')}, Raum ${d.raum}, erfasst von ${d.erfasstVon || '–'}`).join(' | '),
      });
      row.getCell('sn').numFmt = '@';
      row.getCell('det').alignment = { wrapText: true, vertical: 'top' };
    }
  }

  return wb.xlsx.writeBuffer();
}

async function exportExcel() {
  if (!devices.size) { toast('Noch keine Geräte erfasst.'); return; }
  busy('Excel-Datei wird erstellt …');
  try {
    await loadExcel();
    const list = [...devices.values()];
    const buf = await buildWorkbook(list);
    const prefixes = [...new Set(list.map(d => (idParts(d.id) || {}).p).filter(Boolean))];
    const label = prefixes.length === 1 ? prefixes[0] : 'Gesamt';
    const name = `Inventur_${label}_${todayYmd()}.xlsx`;
    busy(false);
    await offerFile(new Blob([buf], { type: XLSX_MIME }), name, async () => {
      S.lastExport = nowIso(); await saveSettings(); renderBanner();
      if (currentView === 'export') renderExport();
    });
  } catch (e) {
    busy(false);
    infoModal('Export fehlgeschlagen', `<p>${esc(e.message || e)}</p>`);
  }
}

/* ---------------- Import & Zusammenführen ---------------- */
function cellText(v) {
  if (v == null) return '';
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
    if ('result' in v) return cellText(v.result);
    if ('text' in v) return cellText(v.text);
    if (v.error) return '';
    return '';
  }
  return String(v);
}
function excelSerialToUtc(n) { return new Date(Math.round((n - 25569) * 86400) * 1000); }
function parseYmd(v) {
  if (v instanceof Date) return isNaN(v) ? '' : `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  const s = String(v).trim();
  if (!s) return '';
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s); if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
  if (/^\d+(\.\d+)?$/.test(s)) return parseYmd(excelSerialToUtc(+s));
  return '';
}
function parseDateTime(v) {
  if (v instanceof Date) {
    if (isNaN(v)) return '';
    const r = new Date(Math.round(v.getTime() / 1000) * 1000);
    return new Date(r.getUTCFullYear(), r.getUTCMonth(), r.getUTCDate(), r.getUTCHours(), r.getUTCMinutes(), r.getUTCSeconds()).toISOString();
  }
  const s = String(v).trim();
  if (!s) return '';
  let m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)).toISOString();
  if (/^\d+(\.\d+)?$/.test(s)) return parseDateTime(excelSerialToUtc(+s));
  const t = Date.parse(s);
  return isNaN(t) ? '' : new Date(t).toISOString();
}

const LABEL_TO_KEY = Object.fromEntries(FIELDS.map(f => [f.label.toLowerCase(), f.key]));
Object.assign(LABEL_TO_KEY, { 'abteilung': 'team', 'inventarnummer': 'altInvNr', 'alte inventarnr.': 'altInvNr', 'typ': 'typ', 'seriennr.': 'seriennr', 'geändert am': 'geaendertAm' });

/** Liest eine Datei (xlsx oder JSON-Backup) und liefert {records, rooms, roomsDone, teams, skipped} */
async function parseImportFile(file) {
  const out = { name: file.name, records: [], rooms: [], roomsDone: {}, teams: [], skipped: 0 };
  if (/\.json$/i.test(file.name) || file.type === 'application/json') {
    const data = JSON.parse(await file.text());
    if (!data || !Array.isArray(data.devices)) throw new Error(`${file.name}: kein gültiges Backup.`);
    for (const d of data.devices) { const r = normRec(d); if (r.id) out.records.push(r); else out.skipped++; }
    if (data.settings) {
      out.rooms = data.settings.rooms || [];
      out.roomsDone = data.settings.roomsDone || {};
      out.teams = data.settings.teams || [];
    }
    return out;
  }
  await loadExcel();
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(await file.arrayBuffer()); }
  catch (e) { throw new Error(`${file.name}: Datei konnte nicht als Excel gelesen werden.`); }
  const ws = wb.getWorksheet('Inventar') || wb.worksheets[0];
  if (!ws) throw new Error(`${file.name}: keine Tabelle gefunden.`);
  const map = {};
  ws.getRow(1).eachCell((c, col) => {
    const key = LABEL_TO_KEY[String(cellText(c.value)).trim().toLowerCase()];
    if (key) map[col] = key;
  });
  if (!Object.values(map).includes('id')) throw new Error(`${file.name}: Spalte „ID“ fehlt – ist das ein Export dieser App?`);
  const kinds = Object.fromEntries(FIELDS.map(f => [f.key, f.kind]));
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const r = {};
    for (const col in map) {
      const key = map[col];
      const v = cellText(row.getCell(+col).value);
      if (kinds[key] === 'date') r[key] = v === '' ? '' : parseYmd(v);
      else if (kinds[key] === 'datetime') r[key] = v === '' ? '' : parseDateTime(v);
      else r[key] = v instanceof Date ? parseYmd(v) : v;
    }
    const rec = normRec(r);
    if (!rec.id) { if (REC_KEYS.some(k => rec[k])) out.skipped++; return; }
    out.records.push(rec);
  });
  const rs = wb.getWorksheet('Räume');
  if (rs) {
    const hm = {};
    rs.getRow(1).eachCell((c, col) => { hm[String(cellText(c.value)).trim().toLowerCase()] = col; });
    rs.eachRow((row, n) => {
      if (n === 1 || !hm['raum']) return;
      const name = String(cellText(row.getCell(hm['raum']).value)).trim();
      if (!name) return;
      const team = hm['team'] ? String(cellText(row.getCell(hm['team']).value)).trim() : '';
      const area = hm['bereich'] ? String(cellText(row.getCell(hm['bereich']).value)).trim() : '';
      void area;
      out.rooms.push({ name, team });
      if (hm['fertig erfasst'] && /^ja$/i.test(String(cellText(row.getCell(hm['fertig erfasst']).value)).trim())) out.roomsDone[name] = nowIso();
    });
  }
  return out;
}

async function importFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  busy('Dateien werden gelesen …');
  const parsed = [];
  const errors = [];
  for (const f of files) {
    try { parsed.push(await parseImportFile(f)); } catch (e) { errors.push(e.message || String(e)); }
  }
  busy(false);
  if (!parsed.length) { infoModal('Import fehlgeschlagen', errors.map(e => `<p>${esc(e)}</p>`).join('')); return; }

  // Vorschau berechnen
  const work = new Map(devices);
  const stat = { neu: 0, aktualisiert: 0, unveraendert: 0, lokalNeuer: 0 };
  const changes = new Map(); // id -> rec, das geschrieben werden soll
  let ownPrefixNew = 0;
  for (const p of parsed) {
    for (const r of p.records) {
      const cur = work.get(r.id);
      if (!cur) {
        work.set(r.id, r); changes.set(r.id, r);
        if (!devices.has(r.id)) { stat.neu++; const x = idParts(r.id); if (x && x.p === S.prefix) ownPrefixNew++; }
        else stat.aktualisiert++;
      } else if (sameContent(cur, r)) {
        if (!changes.has(r.id)) stat.unveraendert++;
      } else if (ts(r.geaendertAm) > ts(cur.geaendertAm)) {
        work.set(r.id, r);
        if (!changes.has(r.id) && devices.has(r.id)) stat.aktualisiert++;
        changes.set(r.id, r);
      } else {
        stat.lokalNeuer++;
      }
    }
  }
  const groups = serialGroups([...work.values()]);
  const skipped = parsed.reduce((s, p) => s + p.skipped, 0);
  const total = parsed.reduce((s, p) => s + p.records.length, 0);

  let body = `<p>${plural(parsed.length, 'Datei', 'Dateien')} mit ${plural(total, 'Eintrag', 'Einträgen')} gelesen:</p>
    <ul>${parsed.map(p => `<li>${esc(p.name)} – ${p.records.length}</li>`).join('')}</ul>
    <div class="stats">
      <div><b>${stat.neu}</b>neu</div>
      <div><b>${stat.aktualisiert}</b>aktualisiert</div>
      <div><b>${stat.unveraendert}</b>unverändert</div>
      <div><b>${stat.lokalNeuer}</b>hier neuer (bleibt)</div>
    </div>`;
  if (errors.length) body += `<div class="hint warn">${errors.map(esc).join('<br>')}</div>`;
  if (skipped) body += `<p class="muted small">${plural(skipped, 'Zeile', 'Zeilen')} ohne ID wurden übersprungen.</p>`;
  if (ownPrefixNew) body += `<div class="hint warn">⚠️ ${plural(ownPrefixNew, 'neuer Eintrag trägt', 'neue Einträge tragen')} das Kürzel dieses Handys (<strong>${esc(S.prefix)}</strong>). Prüft, ob zwei Handys dasselbe Kürzel benutzen!</div>`;
  if (groups.length) body += `<div class="hint warn">⚠️ ${plural(groups.length, 'Seriennummer kommt', 'Seriennummern kommen')} danach unter mehreren IDs vor (mögliche Doppelerfassung):<br>${groups.slice(0, 8).map(g => `<strong>${esc(g[0].seriennr)}</strong>: ${g.map(d => esc(d.id)).join(', ')}`).join('<br>')}${groups.length > 8 ? '<br>…' : ''}<br>Sie werden im Excel-Export im Blatt „Prüfen“ aufgeführt.</div>`;

  const go = await modal({
    title: 'Zusammenführen – Vorschau', body,
    buttons: [
      { label: changes.size ? `Zusammenführen (${changes.size} Änderungen)` : 'Übernehmen', cls: 'primary', value: true },
      { label: 'Abbrechen', value: false },
    ],
  });
  if (!go) return;

  try {
    const recs = [...changes.values()];
    if (recs.length) await DB.putMany('devices', recs);
    for (const r of recs) devices.set(r.id, r);
    // Räume, Teams, Fertig-Markierungen übernehmen
    let settingsChanged = false;
    for (const p of parsed) {
      for (const t of p.teams || []) if (t.name && !S.teams.some(x => x.name === t.name)) { S.teams.push({ name: t.name, color: t.color || '#E6E6E6' }); settingsChanged = true; }
      for (const r of p.rooms || []) {
        if (r.name && !roomByName(r.name)) { S.rooms.push(Object.assign({ name: r.name, team: r.team || NO_TEAM }, inPlan(r) ? { x: r.x, y: r.y, w: r.w, h: r.h } : {}, r.dashed ? { dashed: true } : {})); settingsChanged = true; }
      }
      for (const name in p.roomsDone || {}) if (!S.roomsDone[name]) { S.roomsDone[name] = p.roomsDone[name]; settingsChanged = true; }
    }
    for (const r of recs) {
      if (r.team && !S.teams.some(x => x.name === r.team)) { S.teams.push({ name: r.team, color: '#E6E6E6' }); settingsChanged = true; }
      if (r.raum && !roomByName(r.raum)) { S.rooms.push({ name: r.raum, team: r.team || NO_TEAM }); settingsChanged = true; }
      if (r.typ && !S.types.includes(r.typ)) { S.types.push(r.typ); settingsChanged = true; }
    }
    if (settingsChanged) await saveSettings();
  } catch (e) {
    infoModal('Zusammenführen fehlgeschlagen', `<p>${esc(e.message || e)}</p>`);
    return;
  }
  toast(`Zusammengeführt: ${stat.neu} neu, ${stat.aktualisiert} aktualisiert`);
  renderExport();
}

/* ---------------- Export-Ansicht ---------------- */
function renderExport() {
  const list = [...devices.values()];
  const prefixes = [...new Set(list.map(d => (idParts(d.id) || {}).p).filter(Boolean))].sort();
  const groups = serialGroups(list);
  $('#view-export').innerHTML = `
    <div class="card">
      <h2>Als Excel exportieren</h2>
      <p>${plural(list.length, 'Gerät', 'Geräte')}${prefixes.length ? ` · Kürzel: ${prefixes.map(esc).join(', ')}` : ''}</p>
      <button type="button" class="btn primary block" id="btn-export" ${list.length ? '' : 'disabled'}>Als Excel exportieren</button>
      <p class="muted small" style="margin-top:8px">Letzter Export: ${fmtDateTime(S.lastExport)}. Die Datei enthält die Blätter „Inventar“, „Übersicht“, „Räume“ und bei Bedarf „Prüfen“.</p>
    </div>
    <div class="card">
      <h2>Dateien importieren &amp; zusammenführen</h2>
      <p class="small">Excel-Exporte der anderen Handys hier einlesen (mehrere gleichzeitig möglich). Danach exportierst du oben die <strong>Gesamtliste</strong>.</p>
      <label class="btn block">Dateien auswählen …
        <input type="file" id="import-file" multiple accept=".xlsx,.json,${XLSX_MIME},application/json" hidden>
      </label>
      <p class="muted small" style="margin-top:8px">Gleiche ID → die neuere Änderung gewinnt. Gleiche Seriennummer unter verschiedenen IDs → wird zum Prüfen markiert. Funktioniert auch im Browser am PC.</p>
    </div>
    <div class="card">
      <h2>Zu prüfen</h2>
      ${groups.length ? groups.map(g => `<div class="dup-group">Seriennummer <strong>${esc(g[0].seriennr)}</strong> kommt mehrfach vor:<br>
        ${g.map(d => `<button type="button" class="linklike" data-edit="${esc(d.id)}">${esc(d.id)}</button> (${esc(d.typ)}, Raum ${esc(d.raum)})`).join(' · ')}</div>`).join('')
      : '<p class="muted">Keine doppelten Seriennummern gefunden. 👍</p>'}
    </div>`;
}

/* ---------------- Einstellungen ---------------- */
async function renderSettings() {
  let persisted = null;
  try { persisted = navigator.storage && navigator.storage.persisted ? await navigator.storage.persisted() : null; } catch (_) { /* egal */ }
  const teamOpts = sel => S.teams.map(t => `<option value="${esc(t.name)}" ${t.name === sel ? 'selected' : ''}>${esc(t.name)}</option>`).join('');
  $('#view-settings').innerHTML = `
    <div class="card">
      <h2>Dieses Handy</h2>
      <dl class="kv"><dt>Kürzel</dt><dd>${esc(S.prefix)}</dd><dt>Name</dt><dd>${esc(S.name || '–')}</dd>
        <dt>Nächste ID</dt><dd>${esc(nextId().id)}</dd></dl>
      <button type="button" class="btn block" data-act="edit-device">Kürzel / Name ändern</button>
    </div>

    <div class="card">
      <h2>Räume</h2>
      <p class="muted small">Team-Zuordnung je Raum. Neue Räume erscheinen unter „Weitere Orte“. Das Team wird beim Erfassen automatisch vorgeschlagen.</p>
      <ul class="set-list">${S.rooms.map((r, i) => `<li class="room-item">
        <span class="room-name">${esc(r.name)} <span class="muted small" style="font-weight:400">${inPlan(r) ? '· im Grundriss' : '· weitere Orte'}</span></span>
        <select data-room-team="${i}" aria-label="Team für Raum ${esc(r.name)}">${teamOpts(r.team)}</select>
        <button type="button" class="icon-btn" data-del-room="${i}" aria-label="Raum ${esc(r.name)} entfernen">×</button>
      </li>`).join('')}</ul>
      <div class="add-row">
        <input id="new-room" placeholder="Neuer Raum, z. B. 4310">
        <select id="new-room-team">${teamOpts(NO_TEAM)}</select>
        <button type="button" class="btn small" data-act="add-room">Hinzufügen</button>
      </div>
    </div>

    <div class="card">
      <h2>Teams</h2>
      <p class="muted small">Tippe auf einen Namen, um ihn umzubenennen. Räume und erfasste Geräte werden automatisch angepasst.</p>
      <ul class="set-list">${S.teams.map((t, i) => `<li>
        <input type="color" value="${esc(t.color)}" data-team-color="${i}" aria-label="Farbe ${esc(t.name)}">
        <span class="grow">${t.name === NO_TEAM ? esc(t.name) : `<button type="button" class="linklike" data-rename-team="${i}">${esc(t.name)}</button>`}</span>
        ${t.name === NO_TEAM ? '' : `<button type="button" class="icon-btn" data-del-team="${i}" aria-label="Team ${esc(t.name)} entfernen">×</button>`}
      </li>`).join('')}</ul>
      <div class="add-row">
        <input id="new-team" placeholder="Neues Team">
        <button type="button" class="btn small" data-act="add-team">Hinzufügen</button>
      </div>
    </div>

    <div class="card">
      <h2>Gerätetypen</h2>
      <ul class="set-list">${S.types.map((t, i) => `<li><span class="grow">${esc(t)}</span>
        <button type="button" class="icon-btn" data-del-type="${i}" aria-label="${esc(t)} entfernen">×</button></li>`).join('')}</ul>
      <div class="add-row">
        <input id="new-type" placeholder="Neuer Gerätetyp">
        <button type="button" class="btn small" data-act="add-type">Hinzufügen</button>
      </div>
    </div>

    <div class="card">
      <h2>Datensicherung</h2>
      <p class="small">Die Daten liegen <strong>nur auf diesem Handy</strong> im Browser. Regelmäßig als Excel exportieren oder ein Backup erstellen.</p>
      <dl class="kv"><dt>Letztes Backup</dt><dd>${fmtDateTime(S.lastBackup)}</dd><dt>Letzter Export</dt><dd>${fmtDateTime(S.lastExport)}</dd>
        <dt>Dauerspeicher</dt><dd>${persisted === true ? 'aktiv ✓' : persisted === false ? 'nicht bestätigt – App auf den Home-Bildschirm legen' : 'unbekannt'}</dd></dl>
      <div class="stack">
        <button type="button" class="btn block" data-act="backup">Backup erstellen (JSON)</button>
        <label class="btn block">Backup wiederherstellen …
          <input type="file" id="restore-file" accept=".json,application/json" hidden>
        </label>
      </div>
    </div>

    <div class="card">
      <h2>Gefahrenzone</h2>
      <button type="button" class="btn danger block" data-act="wipe">Alle Daten löschen</button>
    </div>
    <p class="muted small" style="text-align:center">Büro-Inventur · Version ${APP_VERSION}</p>`;
}

async function changeDeviceSettings() {
  const input = {};
  const ok0 = await modal({
    title: 'Kürzel / Name ändern',
    body: `<label class="field"><span>Kürzel dieses Handys (nur Buchstaben)</span>
      <input id="m-prefix" maxlength="3" autocapitalize="characters" autocomplete="off" value="${esc(S.prefix)}"></label>
      <label class="field"><span>Dein Name</span><input id="m-name" value="${esc(S.name)}"></label>`,
    buttons: [
      {
        label: 'Speichern', cls: 'primary', value: true,
        action: el => {
          const p = $('#m-prefix', el).value.trim().toUpperCase();
          if (!/^[A-Z]{1,3}$/.test(p)) { $('#m-prefix', el).classList.add('invalid'); toast('Kürzel: 1–3 Buchstaben (A–Z)'); return false; }
          input.p = p; input.n = $('#m-name', el).value.trim();
        },
      },
      { label: 'Abbrechen', value: false },
    ],
  });
  if (!ok0) return;
  const { p, n } = input;
  if (p !== S.prefix) {
    const existing = [...devices.keys()].filter(id => (idParts(id) || {}).p === S.prefix).length;
    const ok = await confirmModal('Kürzel ändern?',
      `<p>Neue Geräte bekommen ab jetzt IDs wie <strong>${esc(p)}-${pad((S.counters[p] || 0) + 1, 4)}</strong>.</p>
      ${existing ? `<div class="hint warn">${plural(existing, 'Gerät behält', 'Geräte behalten')} ihre bisherige ID mit „${esc(S.prefix)}“.</div>` : ''}
      <p class="small"><strong>Achtung:</strong> Das Kürzel „${esc(p)}“ darf auf keinem anderen Handy verwendet werden.</p>`, 'Kürzel ändern');
    if (!ok) return;
    S.prefix = p;
  }
  S.name = n;
  await saveSettings();
  renderSettings();
  showView('settings');
  toast('Gespeichert');
}

function buildBackup() {
  return {
    app: 'buero-inventur', version: 1, createdAt: nowIso(),
    settings: S, devices: [...devices.values()].sort((a, b) => cmpId(a.id, b.id)),
  };
}
async function createBackup() {
  const blob = new Blob([JSON.stringify(buildBackup(), null, 1)], { type: 'application/json' });
  await offerFile(blob, `Inventur-Backup_${S.prefix}_${todayYmd()}.json`, async () => {
    S.lastBackup = nowIso(); await saveSettings();
    if (currentView === 'settings') renderSettings();
  });
}
async function restoreBackup(file) {
  let data;
  try {
    data = JSON.parse(await file.text());
    if (!data || data.app !== 'buero-inventur' || !Array.isArray(data.devices) || !data.settings) throw new Error();
  } catch (_) { infoModal('Ungültige Datei', '<p>Das ist kein Backup dieser App.</p>'); return false; }
  const ok = await confirmModal('Backup wiederherstellen?',
    `<p>Backup vom ${fmtDateTime(data.createdAt)} mit ${plural(data.devices.length, 'Gerät', 'Geräten')} (Kürzel ${esc(data.settings.prefix || '–')}).</p>
     ${devices.size ? `<div class="hint warn">Die aktuellen ${plural(devices.size, 'Gerät', 'Geräte')} auf diesem Handy werden ersetzt. Zum Zusammenführen stattdessen „Excel → Dateien importieren“ verwenden.</div>` : ''}`,
    'Wiederherstellen', devices.size > 0);
  if (!ok) return false;
  const counters = Object.assign({}, S ? S.counters : {});
  const ns = Object.assign(defaultSettings(), data.settings, { key: 'settings' });
  for (const p in counters) ns.counters[p] = Math.max(ns.counters[p] || 0, counters[p]);
  await DB.clear('devices');
  const recs = data.devices.map(normRec).filter(r => r.id);
  await DB.putMany('devices', recs);
  S = ns;
  await saveSettings();
  devices.clear();
  for (const r of recs) devices.set(r.id, r);
  toast('Backup wiederhergestellt');
  return true;
}

async function wipeAll() {
  const a = await confirmModal('Alle Daten löschen?', `<p>Alle ${plural(devices.size, 'Gerät', 'Geräte')} und Einstellungen werden von diesem Handy gelöscht.</p><p><strong>Vorher exportieren!</strong></p>`, 'Weiter', true);
  if (!a) return;
  const b = await modal({
    title: 'Wirklich alles löschen?',
    body: `<p>Zur Bestätigung <strong>LÖSCHEN</strong> eintippen:</p><input id="m-wipe" autocomplete="off" autocapitalize="characters">
      <p class="muted small" style="margin-top:8px">Der ID-Zähler bleibt erhalten, damit keine ID doppelt vergeben wird.</p>`,
    buttons: [
      { label: 'Endgültig löschen', cls: 'danger-fill', value: true, action: el => { if ($('#m-wipe', el).value.trim().toUpperCase() !== 'LÖSCHEN') { toast('Bitte LÖSCHEN eintippen'); return false; } } },
      { label: 'Abbrechen', value: false },
    ],
  });
  if (!b) return;
  const counters = S.counters;
  await DB.clear('devices');
  devices.clear();
  S = defaultSettings();
  S.counters = counters;
  await saveSettings();
  toast('Alle Daten gelöscht');
  location.replace('#liste');
  route();
}

/* ---------------- Ereignisse ---------------- */
function bindEvents() {
  window.addEventListener('hashchange', route);

  // Einrichtung
  $('#setup-prefix').addEventListener('input', e => { e.target.value = e.target.value.toUpperCase().replace(/[^A-Z]/g, ''); });
  $('#setup-save').addEventListener('click', async () => {
    const p = $('#setup-prefix').value.trim().toUpperCase();
    const n = $('#setup-name').value.trim();
    $('#setup-prefix').classList.toggle('invalid', !/^[A-Z]{1,3}$/.test(p));
    $('#setup-name').classList.toggle('invalid', !n);
    if (!/^[A-Z]{1,3}$/.test(p)) { toast('Bitte ein Kürzel aus 1–3 Buchstaben eingeben.'); return; }
    if (!n) { toast('Bitte deinen Namen eingeben.'); return; }
    S.prefix = p; S.name = n;
    await saveSettings();
    requestPersist();
    location.replace('#liste');
    route();
  });
  $('#setup-restore').addEventListener('change', async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (f && await restoreBackup(f)) { location.replace('#liste'); route(); }
  });

  // Globale Klicks
  document.addEventListener('click', e => {
    const ed = e.target.closest('[data-edit]');
    if (ed) { goEdit(ed.dataset.edit); return; }
    const go = e.target.closest('[data-go]');
    if (go) { location.hash = '#' + go.dataset.go; return; }
    const room = e.target.closest('#plan [data-room]');
    if (room) { openRoomSheet(room.dataset.room); return; }
    if (e.target.closest('#btn-export')) { exportExcel(); }
  });

  $('#fab').addEventListener('click', () => goNew(null));

  // Liste
  $('#search').addEventListener('input', renderList);
  $('#filter-toggle').addEventListener('click', () => { $('#filters').hidden = !$('#filters').hidden; });
  $$('#filters select').forEach(s => s.addEventListener('change', renderList));
  $('#flt-reset').addEventListener('click', () => { $$('#filters select').forEach(s => { s.value = ''; }); renderList(); });

  // Formular
  const form = $('#device-form');
  form.addEventListener('input', () => { formDirty = true; });
  form.addEventListener('change', e => {
    formDirty = true;
    if (e.target.classList.contains('invalid') && e.target.value) e.target.classList.remove('invalid');
  });
  form.addEventListener('submit', e => { e.preventDefault(); saveForm(false); });
  $('#form-save-next').addEventListener('click', () => saveForm(true));
  $('#f-seriennr').addEventListener('input', checkSerial);
  $('#f-raum').addEventListener('change', () => {
    const r = roomByName($('#f-raum').value);
    if (r && r.team) $('#f-team').value = r.team;
  });
  $('#f-raum-plan').addEventListener('click', async () => {
    const name = await pickRoomModal($('#f-raum').value);
    if (!name) return;
    $('#f-raum').value = name;
    $('#f-raum').classList.remove('invalid');
    const r = roomByName(name);
    if (r && r.team) $('#f-team').value = r.team;
    formDirty = true;
  });
  $('#form-cancel').addEventListener('click', async () => {
    if (formDirty && !(await confirmModal('Änderungen verwerfen?', '<p>Die Eingaben wurden noch nicht gespeichert.</p>', 'Verwerfen', true))) return;
    leaveForm();
  });
  $('#form-delete').addEventListener('click', deleteCurrent);

  // Export / Import
  $('#view-export').addEventListener('change', e => {
    if (e.target.id === 'import-file') { const files = [...e.target.files]; e.target.value = ''; importFiles(files); }
  });

  // Einstellungen
  const vs = $('#view-settings');
  vs.addEventListener('click', async e => {
    const t = e.target.closest('button');
    if (!t) return;
    const act = t.dataset.act;
    if (act === 'edit-device') return changeDeviceSettings();
    if (act === 'backup') return createBackup();
    if (act === 'wipe') return wipeAll();
    if (act === 'add-room') {
      const name = $('#new-room').value.trim();
      if (!name) return;
      if (roomByName(name)) { toast('Raum existiert bereits'); return; }
      S.rooms.push({ name, team: $('#new-room-team').value });
    } else if (act === 'add-team') {
      const name = $('#new-team').value.trim();
      if (!name) return;
      if (S.teams.some(x => x.name === name)) { toast('Team existiert bereits'); return; }
      const palette = ['#FFD6D6', '#D6F5F0', '#FFF2B3', '#E0D6CC', '#F7D6F0'];
      S.teams.splice(S.teams.length - 1, 0, { name, color: palette[S.teams.length % palette.length] });
    } else if (act === 'add-type') {
      const name = $('#new-type').value.trim();
      if (!name) return;
      if (S.types.includes(name)) { toast('Gerätetyp existiert bereits'); return; }
      S.types.push(name);
    } else if (t.dataset.delRoom != null) {
      const r = S.rooms[+t.dataset.delRoom];
      const n = [...devices.values()].filter(d => d.raum === r.name).length;
      if (!(await confirmModal(`Raum ${r.name} entfernen?`, n ? `<p>${plural(n, 'Gerät behält', 'Geräte behalten')} den Raum trotzdem.</p>` : '<p>Der Raum wird aus der Auswahl entfernt.</p>', 'Entfernen', true))) return;
      S.rooms.splice(+t.dataset.delRoom, 1);
    } else if (t.dataset.renameTeam != null) {
      const tm = S.teams[+t.dataset.renameTeam];
      const input = {};
      const ok = await modal({
        title: 'Team umbenennen',
        body: `<label class="field"><span>Neuer Name</span><input id="m-team" value="${esc(tm.name)}" autocapitalize="words"></label>`,
        buttons: [
          {
            label: 'Umbenennen', cls: 'primary', value: true,
            action: el => {
              const v = $('#m-team', el).value.trim();
              if (!v) { toast('Bitte einen Namen eingeben'); return false; }
              if (v !== tm.name && S.teams.some(x => x.name === v)) { toast('Team existiert bereits'); return false; }
              input.v = v;
            },
          },
          { label: 'Abbrechen', value: false },
        ],
        onOpen: el => { const i = $('#m-team', el); i.focus(); i.select(); },
      });
      if (!ok || input.v === tm.name) return;
      const old = tm.name;
      tm.name = input.v;
      S.rooms.forEach(r => { if (r.team === old) r.team = input.v; });
      const changed = [...devices.values()].filter(d => d.team === old).map(d => ({ ...d, team: input.v, geaendertAm: nowIso() }));
      if (changed.length) { await DB.putMany('devices', changed); changed.forEach(d => devices.set(d.id, d)); }
      toast(`Umbenannt${changed.length ? ` (${plural(changed.length, 'Gerät', 'Geräte')} angepasst)` : ''}`);
    } else if (t.dataset.delTeam != null) {
      const tm = S.teams[+t.dataset.delTeam];
      if (!(await confirmModal(`Team „${tm.name}“ entfernen?`, '<p>Räume dieses Teams werden auf „ohne Team“ gesetzt. Erfasste Geräte behalten ihren Eintrag.</p>', 'Entfernen', true))) return;
      S.teams.splice(+t.dataset.delTeam, 1);
      S.rooms.forEach(r => { if (r.team === tm.name) r.team = NO_TEAM; });
    } else if (t.dataset.delType != null) {
      S.types.splice(+t.dataset.delType, 1);
    } else return;
    await saveSettings();
    const y = window.scrollY;
    await renderSettings();
    window.scrollTo(0, y);
  });
  vs.addEventListener('change', async e => {
    const t = e.target;
    if (t.id === 'restore-file') {
      const f = t.files[0]; t.value = '';
      if (f && await restoreBackup(f)) { renderSettings(); }
      return;
    }
    if (t.dataset.roomTeam != null) S.rooms[+t.dataset.roomTeam].team = t.value;
    else if (t.dataset.teamColor != null) S.teams[+t.dataset.teamColor].color = t.value;
    else return;
    await saveSettings();
    toast('Gespeichert');
  });
}

async function requestPersist() {
  try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch (_) { /* egal */ }
}

/* ---------------- Start ---------------- */
// Ältere Versionen hatten "Flure" statt Draufsicht: Plan-Koordinaten nachrüsten, Team-Zuordnung behalten
function migrateRooms() {
  let changed = false;
  for (const r of S.rooms) {
    if ('row' in r) { delete r.row; changed = true; }
    const d = DEFAULT_ROOMS.find(x => x.name === r.name);
    if (d && inPlan(d) && !inPlan(r)) { Object.assign(r, { x: d.x, y: d.y, w: d.w, h: d.h }); changed = true; }
    else if (!inPlan(r) && 'w' in r) { delete r.w; changed = true; }
  }
  return changed;
}

async function init() {
  try {
    await DB.open();
  } catch (e) {
    document.querySelector('main').innerHTML = `<div class="card"><h2>Speicher nicht verfügbar</h2><p>Der Browser erlaubt keine lokale Speicherung (privater Modus?). Bitte im normalen Modus öffnen.</p></div>`;
    return;
  }
  S = (await DB.get('meta', 'settings')) || defaultSettings();
  S = Object.assign(defaultSettings(), S);
  if (migrateRooms()) await saveSettings();
  for (const d of await DB.getAll('devices')) devices.set(d.id, d);
  bindEvents();
  if (S.prefix) requestPersist();
  route();

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    const hadController = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('service-worker.js').catch(() => { /* offline-Cache nicht verfügbar */ });
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (hadController) toast('Neue Version installiert – App bitte neu öffnen.');
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
