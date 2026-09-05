// Shared helpers. No dependencies.

/* ---------------- DOM ---------------- */

export const $  = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

/**
 * Element builder.
 *
 * There is deliberately no `html` option. Children are appended as text nodes
 * unless they are already Nodes, and attributes are set through setAttribute,
 * so a task name typed into a shared spreadsheet by a subcontractor can never
 * be interpreted as markup. If markup is ever genuinely needed, build it with
 * el() rather than reintroducing a string path.
 */
export function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v === true ? '' : String(v));
  }
  const kids = Array.isArray(children) ? children : [children];
  for (const c of kids) {
    if (c === null || c === undefined || c === false) continue;
    node.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

export function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

export function mount(node, ...children) {
  clear(node);
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    node.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

/* Inline SVG icons — 16px, stroke-based, drawn to match the hairline rules. */
const ICONS = {
  dashboard: 'M3 3h7v7H3zM14 3h7v4h-7zM14 11h7v10h-7zM3 14h7v7H3z',
  projects:  'M3 6h7l2 3h9v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z',
  tasks:     'M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2',
  gantt:     'M4 6h9M4 11h13M4 16h6M4 3v18',
  register:  'M12 3l9 16H3zM12 9v5M12 17h.01',
  insights:  'M12 3a6 6 0 0 0-3.5 10.9V17h7v-3.1A6 6 0 0 0 12 3zM9.5 21h5',
  data:      'M4 6c0-1.7 3.6-3 8-3s8 1.3 8 3-3.6 3-8 3-8-1.3-8-3zM4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3',
  settings:  'M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM19.4 15a1.6 1.6 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.6 1.6 0 0 0-2.7 1.1V21a2 2 0 1 1-4 0v-.1A1.6 1.6 0 0 0 7.9 19.4a1.6 1.6 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.6 1.6 0 0 0-1.1-2.7H2a2 2 0 1 1 0-4h.1A1.6 1.6 0 0 0 3.7 7.9a1.6 1.6 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.6 1.6 0 0 0 1.8.3H8a1.6 1.6 0 0 0 1-1.5V2a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 2.7 1.1l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.6 1.6 0 0 0-.3 1.8V8a1.6 1.6 0 0 0 1.5 1H22a2 2 0 1 1 0 4h-.1a1.6 1.6 0 0 0-1.5 1z',
  upload:    'M12 16V4M7 9l5-5 5 5M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2',
  download:  'M12 4v12M7 11l5 5 5-5M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2',
  plus:      'M12 5v14M5 12h14',
  close:     'M6 6l12 12M18 6L6 18',
  menu:      'M4 7h16M4 12h16M4 17h16',
  refresh:   'M21 12a9 9 0 1 1-3-6.7M21 4v5h-5',
  trash:     'M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13',
  edit:      'M4 20h4L20 8l-4-4L4 16z',
  save:      'M5 3h11l3 3v15H5zM8 3v6h7V3M8 14h8v7H8z',
  check:     'M5 13l4 4L19 7',
  warn:      'M12 3l9 16H3zM12 9v5M12 17h.01',
};

export function icon(name, size = 16) {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('width', size);
  svg.setAttribute('height', size);
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.7');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  const p = document.createElementNS(ns, 'path');
  p.setAttribute('d', ICONS[name] || ICONS.tasks);
  svg.appendChild(p);
  return svg;
}

export function svgEl(tag, attrs = {}) {
  const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined) continue;
    node.setAttribute(k, String(v));
  }
  return node;
}

/* ---------------- ids & misc ---------------- */

export function uid(prefix = 'id') {
  const r = (crypto?.randomUUID?.() || Math.random().toString(36).slice(2) + Date.now().toString(36));
  return `${prefix}_${r.replace(/-/g, '').slice(0, 12)}`;
}

export function debounce(fn, ms = 400) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

export function clamp(n, lo, hi) { return Math.min(hi, Math.max(lo, n)); }

export function deepClone(v) {
  if (typeof structuredClone === 'function') return structuredClone(v);
  return JSON.parse(JSON.stringify(v));
}

export function groupBy(arr, keyFn) {
  const m = new Map();
  for (const item of arr) {
    const k = keyFn(item);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(item);
  }
  return m;
}

export function sum(arr, f = (x) => x) { return arr.reduce((a, b) => a + (Number(f(b)) || 0), 0); }

/* ---------------- dates ---------------- */

const MS_DAY = 86400000;

/** Parse most things a spreadsheet throws at us into an ISO yyyy-mm-dd string. */
export function toISO(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value)) return fmtISO(value);

  if (typeof value === 'number') {
    // Excel serial date (1900 system). 25569 = days between 1970-01-01 and Excel epoch.
    if (value > 20000 && value < 80000) {
      const d = new Date(Math.round((value - 25569) * MS_DAY));
      return fmtISO(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    }
    return null;
  }

  const s = String(value).trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return mk(+m[1], +m[2], +m[3]);

  // dd/mm/yyyy and mm/dd/yyyy are ambiguous. Prefer day-first (the convention
  // on most construction programmes outside the US); flip when impossible.
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (m) {
    let [, a, b, y] = m.map(Number);
    if (y < 100) y += y < 70 ? 2000 : 1900;
    if (a > 12 && b <= 12) return mk(y, b, a);
    if (b > 12 && a <= 12) return mk(y, a, b);
    return mk(y, b, a);
  }

  m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{2,4})$/);
  if (m) {
    const mo = monthIndex(m[2]);
    let y = +m[3]; if (y < 100) y += y < 70 ? 2000 : 1900;
    if (mo >= 0) return mk(y, mo + 1, +m[1]);
  }

  const d = new Date(s);
  return isNaN(d) ? null : fmtISO(d);

  function mk(y, mo, dd) {
    if (!y || !mo || !dd || mo > 12 || dd > 31) return null;
    return `${y}-${String(mo).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;
  }
}

function monthIndex(name) {
  const n = name.slice(0, 3).toLowerCase();
  return ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(n);
}

export function fmtISO(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function parseISO(iso) {
  if (!iso) return null;
  const [y, m, d] = String(iso).split('-').map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

export function addDays(iso, n) {
  const d = parseISO(iso);
  if (!d) return null;
  d.setDate(d.getDate() + n);
  return fmtISO(d);
}

export function diffDays(a, b) {
  const da = parseISO(a), db = parseISO(b);
  if (!da || !db) return null;
  return Math.round((db - da) / MS_DAY);
}

export function todayISO() { return fmtISO(new Date()); }

export function fmtDate(iso) {
  const d = parseISO(iso);
  if (!d) return '—';
  return `${String(d.getDate()).padStart(2, '0')} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()]} ${String(d.getFullYear()).slice(2)}`;
}

export function fmtDateLong(iso) {
  const d = parseISO(iso);
  if (!d) return '—';
  return d.toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' });
}

/* ---------- working-day calendar ----------
   workingDays is an array of weekday indices (0=Sun..6=Sat) that count as work.
   holidays is an array of ISO strings.

   Every date operation here is O(1) rather than a day-by-day walk. That matters:
   a 2000-task chain spans tens of thousands of calendar days, and the naive
   version made scheduling quadratic — about 25 seconds where this takes well
   under one. The index is built lazily and attached to the calendar object, so
   callers pass the same plain { workingDays, holidays } shape as before. */

// 1900-01-01 was a Monday, and sits before any date a programme will contain,
// so every real day index is positive.
const CAL_EPOCH_MS = Date.UTC(1900, 0, 1);

function dayIndexOf(iso) {
  const d = parseISO(iso);
  if (!d) return null;
  return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - CAL_EPOCH_MS) / MS_DAY);
}

function isoFromDayIndex(i) {
  const d = new Date(CAL_EPOCH_MS + i * MS_DAY);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

// Epoch is a Monday (weekday 1), so weekday cycles 1,2,3,4,5,6,0.
function weekdayOfIndex(i) { return (((1 + i) % 7) + 7) % 7; }

/** Build (once) the lookup tables that make the arithmetic constant-time. */
function ensureIndex(cal) {
  if (cal && cal.__idx) return cal.__idx;

  const days = (cal?.workingDays?.length ? cal.workingDays : [1, 2, 3, 4, 5])
    .map(Number).filter((d) => d >= 0 && d <= 6);
  const wset = new Set(days.length ? days : [1, 2, 3, 4, 5]);
  const perWeek = wset.size;

  // prefix[r] = working weekdays among the first r days of a week window
  const prefix = [0];
  for (let j = 0; j < 7; j++) prefix.push(prefix[j] + (wset.has(weekdayOfIndex(j)) ? 1 : 0));

  // offsets[k] = day offset within a week of the k-th working day
  const offsets = [];
  for (let j = 0; j < 7; j++) if (wset.has(weekdayOfIndex(j))) offsets.push(j);

  // Only holidays that land on a working weekday actually remove a day.
  const holIdx = [...new Set(cal?.holidays || [])]
    .map(dayIndexOf)
    .filter((i) => i !== null && i >= 0 && wset.has(weekdayOfIndex(i)))
    .sort((a, b) => a - b);

  const idx = { wset, perWeek, prefix, offsets, holIdx };
  if (cal && typeof cal === 'object') {
    Object.defineProperty(cal, '__idx', { value: idx, enumerable: false, writable: true, configurable: true });
  }
  return idx;
}

/** Number of holiday day-indices strictly less than i. */
function holidaysBefore(idx, i) {
  const a = idx.holIdx;
  if (!a.length) return 0;
  let lo = 0, hi = a.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (a[mid] < i) lo = mid + 1; else hi = mid;
  }
  return lo;
}

/** Working days strictly before day index i, counting from the epoch. */
function workingCountBefore(idx, i) {
  if (i <= 0) return 0;
  const weeks = Math.floor(i / 7);
  const rem = i - weeks * 7;
  return weeks * idx.perWeek + idx.prefix[rem] - holidaysBefore(idx, i);
}

function isWorkingIndex(idx, i) {
  if (!idx.wset.has(weekdayOfIndex(i))) return false;
  const a = idx.holIdx;
  if (!a.length) return true;
  let lo = 0, hi = a.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (a[mid] === i) return false;
    if (a[mid] < i) lo = mid + 1; else hi = mid - 1;
  }
  return true;
}

function nextWorkingIndex(idx, i, dir = 1) {
  let cur = i, guard = 0;
  while (!isWorkingIndex(idx, cur) && guard++ < 400) cur += dir;
  return cur;
}

/** Day index of the k-th working day since the epoch (0-based). */
function indexFromOrdinal(idx, k) {
  const target = Math.max(0, k);
  const weeks = Math.floor(target / idx.perWeek);
  const rem = target - weeks * idx.perWeek;
  let i = weeks * 7 + idx.offsets[rem];

  // Holidays only ever remove working days, so the no-holiday guess is a lower
  // bound and we walk forward from it. Converges in one or two passes.
  for (let it = 0; it < 64; it++) {
    i = nextWorkingIndex(idx, i);
    const c = workingCountBefore(idx, i);
    if (c === target) return i;
    const deficit = target - c;
    if (deficit <= 0) return i;
    i += Math.max(1, Math.ceil((deficit * 7) / idx.perWeek));
  }
  return i;
}

export function isWorkingDay(iso, cal) {
  const i = dayIndexOf(iso);
  if (i === null) return false;
  return isWorkingIndex(ensureIndex(cal), i);
}

export function nextWorkingDay(iso, cal, dir = 1) {
  const i = dayIndexOf(iso);
  if (i === null) return iso;
  return isoFromDayIndex(nextWorkingIndex(ensureIndex(cal), i, dir));
}

/** Add n working days to a date, where day 1 is the start date itself. */
export function addWorkingDays(iso, n, cal) {
  const i = dayIndexOf(iso);
  if (i === null) return null;
  const idx = ensureIndex(cal);
  const start = nextWorkingIndex(idx, i);
  const ord = workingCountBefore(idx, start) + Math.max(1, Math.round(n)) - 1;
  return isoFromDayIndex(indexFromOrdinal(idx, ord));
}

/** Inclusive count of working days between two dates. */
export function workingDaysBetween(a, b, cal) {
  const ia = dayIndexOf(a), ib = dayIndexOf(b);
  if (ia === null || ib === null) return null;
  if (ib < ia) return -workingDaysBetween(b, a, cal);
  const idx = ensureIndex(cal);
  return workingCountBefore(idx, ib + 1) - workingCountBefore(idx, ia);
}

/** Shift by n working days (0 = snap to the next working day). */
export function shiftWorkingDays(iso, n, cal) {
  const i = dayIndexOf(iso);
  if (i === null) return iso;
  const idx = ensureIndex(cal);
  const base = nextWorkingIndex(idx, i);
  if (n === 0) return isoFromDayIndex(base);
  const ord = workingCountBefore(idx, base) + Math.round(n);
  return isoFromDayIndex(indexFromOrdinal(idx, Math.max(0, ord)));
}

/* ---------------- numbers ---------------- */

export function fmtNum(n, dp = 0) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

export function fmtPct(n, dp = 0) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return `${Number(n).toFixed(dp)}%`;
}

export function fmtMoney(n, currency = 'USD') {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const compact = abs >= 1_000_000 ? { notation: 'compact', maximumFractionDigits: 2 } : { maximumFractionDigits: 0 };
  try {
    return Number(n).toLocaleString(undefined, { style: 'currency', currency, ...compact });
  } catch {
    return `${currency} ${fmtNum(n)}`;
  }
}

export function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(/[^0-9.\-]/g, '');
  if (!s || s === '-' || s === '.') return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}

/** Percent values arrive as 0.45, 45, "45%" — normalise to 0..100. */
export function toPct(v) {
  if (v === null || v === undefined || v === '') return null;
  const raw = String(v).trim();
  let n = toNum(raw);
  if (n === null) return null;
  if (typeof v === 'number' && v > 0 && v <= 1 && !Number.isInteger(v)) n = v * 100;
  else if (raw.endsWith('%')) n = toNum(raw);
  return clamp(n, 0, 100);
}

/* ---------------- text ---------------- */

export function norm(s) {
  return String(s ?? '').toLowerCase().replace(/[\s_\-./\\]+/g, '');
}

export function titleCase(s) {
  return String(s ?? '').replace(/\b\w/g, (c) => c.toUpperCase());
}

/* ---------------- feedback ---------------- */

export function toast(message, tone = 'blueprint', ms = 4200) {
  let host = $('.toasts');
  if (!host) {
    host = el('div', { class: 'toasts', role: 'status', 'aria-live': 'polite' });
    document.body.appendChild(host);
  }
  const t = el('div', { class: 'toast', dataset: { tone }, text: message });
  host.appendChild(t);
  setTimeout(() => t.remove(), ms);
}

/** Promise-based modal. resolve(true) on confirm, resolve(false) on cancel. */
export function confirmDialog({ title, message, confirmLabel = 'Confirm', tone = 'primary' }) {
  return new Promise((resolve) => {
    const dlg = el('dialog', { class: 'modal' });
    const cancel = el('button', { class: 'btn', text: 'Cancel', onclick: () => { dlg.close(); resolve(false); } });
    const ok = el('button', {
      class: `btn btn--${tone === 'danger' ? 'danger' : 'primary'}`,
      text: confirmLabel,
      onclick: () => { dlg.close(); resolve(true); },
    });
    mount(dlg,
      el('div', { class: 'modal__head' }, [el('h2', { text: title })]),
      el('div', { class: 'modal__body' }, [el('p', { text: message })]),
      el('div', { class: 'modal__foot' }, [cancel, ok]),
    );
    dlg.addEventListener('close', () => dlg.remove());
    document.body.appendChild(dlg);
    dlg.showModal();
    ok.focus();
  });
}

/** Generic modal shell. Returns { dialog, body, foot, close }. */
export function openModal({ title, wide = false, onClose } = {}) {
  const dlg = el('dialog', { class: `modal${wide ? ' modal--wide' : ''}` });
  const body = el('div', { class: 'modal__body' });
  const foot = el('div', { class: 'modal__foot' });
  const closeBtn = el('button', {
    class: 'iconbtn', 'aria-label': 'Close', onclick: () => dlg.close(),
  }, [icon('close')]);
  mount(dlg,
    el('div', { class: 'modal__head' }, [el('h2', { text: title }), closeBtn]),
    body, foot,
  );
  dlg.addEventListener('close', () => { dlg.remove(); onClose?.(); });
  document.body.appendChild(dlg);
  dlg.showModal();
  return { dialog: dlg, body, foot, close: () => dlg.close() };
}

/* ---------------- file helpers ---------------- */

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = el('a', { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export function downloadJSON(obj, filename) {
  downloadBlob(new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' }), filename);
}

export function pickFile(accept) {
  return new Promise((resolve) => {
    const input = el('input', { type: 'file', accept, class: 'hidden' });
    input.addEventListener('change', () => { resolve(input.files?.[0] || null); input.remove(); });
    document.body.appendChild(input);
    input.click();
  });
}
