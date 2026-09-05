/**
 * Session state and persistence.
 *
 * Three layers, deliberately independent:
 *   1. In-memory state — the live session.
 *   2. IndexedDB autosave — survives a reload or a crash. This is what makes a
 *      part-finished set of reports safe when the tab dies mid-run.
 *   3. A session FILE the user saves and keeps — survives a cleared browser, a
 *      different machine, and can be sent to a colleague.
 *
 * API KEYS ARE EXCLUDED FROM 2 AND 3. They are the one thing here worth
 * stealing, and a session file is likely to be emailed or synced. Exclusion is
 * the point, not an oversight — the loader says so plainly.
 */

import { downloadJSON, pickFile, uid, todayISO } from './util.js';

export const SESSION_FORMAT = 'bimtrack.session';
export const SESSION_VERSION = 1;

const DB_NAME = 'bim-tracker';
const STORE = 'session';
const RECORD = 'current';
const LS_KEY = 'bimtrack:session';

/** Any key matching these never reaches disk or a portable file. */
const FORBIDDEN = [/api[_-]?key/i, /\bsecret\b/i, /password/i, /bearer/i, /\bcredential/i];

export function stripSecrets(obj) {
  if (obj == null) return obj;
  if (Array.isArray(obj)) return obj.map(stripSecrets);
  if (typeof obj !== 'object') return obj;
  const out = {};
  for (const k of Object.keys(obj)) {
    if (FORBIDDEN.some((re) => re.test(k))) continue;
    out[k] = stripSecrets(obj[k]);
  }
  return out;
}

/* ------------------------------ live state ------------------------------ */

const listeners = new Set();
let state = blank();
let saveState = 'idle';

export function blank() {
  return {
    id: uid('ses'),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    file: null,          // { name, size }
    model: null,         // parsed workbook
    confirmed: false,    // user has accepted the interpretation
    selection: [],       // site codes ticked for the next run
    includeMaster: false,
    reports: {},         // key -> report object (site code, or '__master__')
    settings: {
      provider: 'gemini',
      model: '',
      rememberKeyForSession: false,
      reviewPayload: true,
      maxTokens: 4096,
    },
  };
}

export const MASTER_KEY = '__master__';

export function get() { return state; }
export function subscribe(fn) { listeners.add(fn); return () => listeners.delete(fn); }
function emit() { for (const fn of listeners) fn(state); }
export function getSaveState() { return saveState; }

let timer = null;
export function update(mutator, { silent = false } = {}) {
  mutator(state);
  state.updatedAt = new Date().toISOString();
  saveState = 'dirty';
  if (!silent) emit();
  clearTimeout(timer);
  timer = setTimeout(async () => {
    saveState = 'saving';
    emit();
    const ok = await autosave();
    saveState = ok ? 'saved' : 'error';
    emit();
  }, 500);
}

export function replace(next) {
  state = { ...blank(), ...next };
  saveState = 'dirty';
  emit();
  autosave();
}

export function reset() {
  state = blank();
  saveState = 'idle';
  emit();
  autosave();
}

/** Clears generated reports but keeps the loaded workbook and selection. */
export function clearReports() {
  update((s) => { s.reports = {}; });
}

/* ------------------------------ IndexedDB ------------------------------ */

let dbPromise = null;
function openDB() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) { reject(new Error('no indexeddb')); return; }
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function autosave() {
  const payload = stripSecrets(state);
  try {
    const db = await openDB();
    await new Promise((res, rej) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(payload, RECORD);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
    return true;
  } catch {
    try { localStorage.setItem(LS_KEY, JSON.stringify(payload)); return true; }
    catch { return false; }
  }
}

export async function restore() {
  try {
    const db = await openDB();
    const v = await new Promise((res, rej) => {
      const tx = db.transaction(STORE, 'readonly');
      const r = tx.objectStore(STORE).get(RECORD);
      r.onsuccess = () => res(r.result || null);
      r.onerror = () => rej(r.error);
    });
    if (v) { state = { ...blank(), ...v }; saveState = 'saved'; emit(); return true; }
  } catch { /* fall through */ }
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) { state = { ...blank(), ...JSON.parse(raw) }; saveState = 'saved'; emit(); return true; }
  } catch { /* nothing usable */ }
  emit();
  return false;
}

export async function wipe() {
  try {
    const db = await openDB();
    await new Promise((res) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(RECORD);
      tx.oncomplete = () => res();
      tx.onerror = () => res();
    });
  } catch { /* ignore */ }
  try { localStorage.removeItem(LS_KEY); } catch { /* ignore */ }
}

/* --------------------------- session file I/O --------------------------- */

export function buildSessionFile() {
  return {
    format: SESSION_FORMAT,
    version: SESSION_VERSION,
    app: 'BIM Multi-Site Delivery Tracker',
    savedAt: new Date().toISOString(),
    note: 'API keys are deliberately not included. You will be asked to re-enter yours after loading.',
    state: stripSecrets(state),
  };
}

export function saveSessionToFile() {
  const payload = buildSessionFile();
  const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
  const name = `bimtrack-session-${stamp}.json`;
  downloadJSON(payload, name);
  return name;
}

export async function loadSessionFromFile() {
  const file = await pickFile('.json,application/json');
  if (!file) return { cancelled: true };

  let parsed;
  try {
    parsed = JSON.parse(await file.text());
  } catch {
    throw new Error('That is not a session file. Session files are the JSON produced by "Save session". A report export cannot be loaded back.');
  }
  if (!parsed || parsed.format !== SESSION_FORMAT) {
    throw new Error(`That file is not a ${SESSION_FORMAT} session. Check you picked the right one.`);
  }
  if (parsed.version > SESSION_VERSION) {
    throw new Error(`That session was saved by a newer version of the app (v${parsed.version}). Update before opening it.`);
  }
  if (!parsed.state || typeof parsed.state !== 'object') {
    throw new Error('That session file has no state in it.');
  }
  replace(stripSecrets(parsed.state));
  return { ok: true, filename: file.name, savedAt: parsed.savedAt };
}

/* ------------------------------- API key ------------------------------- */

const KEY_SESSION = 'bimtrack:key';
let memoryKey = '';

export function setApiKey(key, remember) {
  memoryKey = key || '';
  try {
    if (remember && memoryKey) sessionStorage.setItem(KEY_SESSION, memoryKey);
    else sessionStorage.removeItem(KEY_SESSION);
  } catch { /* private mode: memory only */ }
}

export function getApiKey() {
  if (memoryKey) return memoryKey;
  try {
    const k = sessionStorage.getItem(KEY_SESSION);
    if (k) { memoryKey = k; return k; }
  } catch { /* ignore */ }
  return '';
}

export function clearApiKey() {
  memoryKey = '';
  try { sessionStorage.removeItem(KEY_SESSION); } catch { /* ignore */ }
}

export function hasApiKey() { return !!getApiKey(); }

export function maskKey(key) {
  if (!key) return '';
  if (key.length <= 10) return '•'.repeat(key.length);
  return `${key.slice(0, 4)}${'•'.repeat(Math.min(16, key.length - 8))}${key.slice(-4)}`;
}
