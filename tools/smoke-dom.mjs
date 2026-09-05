/**
 * Headless UI smoke test.
 *
 * A passing syntax check says nothing about whether the app renders. The
 * commonest failure for an app like this is a thrown error inside a render
 * function producing a blank page with a perfectly green build, so this boots
 * the real thing under jsdom and walks the whole workflow: load a workbook,
 * confirm, select, generate (against a stubbed provider), export, save and
 * reload a session.
 *
 * Run: node --experimental-vm-modules tools/smoke-dom.mjs
 */

import { JSDOM } from 'jsdom';
import 'fake-indexeddb/auto';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const require = createRequire(import.meta.url);
const XLSX = require(path.join(root, 'assets/vendor/xlsx.full.min.js'));

const dom = new JSDOM(
  fs.readFileSync(path.join(root, 'index.html'), 'utf8').replace(/<script[\s\S]*?<\/script>/g, ''),
  { url: 'http://localhost/', pretendToBeVisual: true },
);
const { window } = dom;

global.window = window;
global.document = window.document;
for (const k of ['navigator', 'location']) {
  Object.defineProperty(globalThis, k, { value: window[k], writable: true, configurable: true });
}
global.HTMLElement = window.HTMLElement;
global.Node = window.Node;
global.Blob = window.Blob;
global.URL = window.URL;
global.AbortController = window.AbortController || AbortController;
// The token meter and limits store use localStorage/sessionStorage. In a real
// browser these are globals; under jsdom they hang off window only, so without
// this the meter silently records nothing and the test would pass a broken app.
for (const k of ['localStorage', 'sessionStorage']) {
  Object.defineProperty(globalThis, k, { value: window[k], writable: true, configurable: true });
}
global.requestAnimationFrame = (fn) => setTimeout(fn, 0);
global.queueMicrotask = queueMicrotask;
window.scrollTo = () => {};
window.XLSX = XLSX;
window.open = () => null;                    // popup blocked, deliberately
window.URL.createObjectURL = () => 'blob:x';
window.URL.revokeObjectURL = () => {};

// jsdom implements neither <dialog> behaviour nor sessionStorage quotas.
Object.defineProperty(window.HTMLElement.prototype, 'showModal', {
  value() { this.setAttribute('open', ''); }, writable: true, configurable: true,
});
Object.defineProperty(window.HTMLElement.prototype, 'close', {
  value() { this.removeAttribute('open'); this.dispatchEvent(new window.Event('close')); },
  writable: true, configurable: true,
});
if (!window.crypto?.randomUUID) {
  Object.defineProperty(window, 'crypto', {
    value: { randomUUID: () => `${Math.random().toString(36).slice(2)}${Date.now().toString(36)}` },
    writable: true, configurable: true,
  });
}

let errors = [];
const realError = console.error;
console.error = (...a) => { errors.push(a.map(String).join(' ')); };
window.addEventListener('error', (e) => errors.push(`window: ${e.message}`));

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++; else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};
const wait = (ms = 60) => new Promise((r) => setTimeout(r, ms));
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => [...document.querySelectorAll(sel)];
const textOf = (sel) => ($(sel)?.textContent || '');

/* ------------------------------ fixture ------------------------------ */

function fixtureBuffer() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Site code', 'Description (generic)', 'Wave / batch', 'Start date', 'Target submission', 'Status', 'Coordinator (R code)', 'Detail sheet'],
    ['A-01', 'Type B', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 2, 1), 'In progress', 'R1', 'A-01'],
    ['A-02', 'Type B', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 3, 1), 'In progress', 'R1', 'A-02'],
    ['A-03', 'Type C', 'Wave 2', new Date(2026, 2, 2), new Date(2026, 5, 1), 'Planned', 'R6', ''],
  ]), 'Sites');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Order', 'ID', 'Category', 'Standard task / scope item', 'Definition of done (category)', 'Typical duration (days)', 'Applies to', 'Default prerequisites (PT IDs)'],
    [10, 'C010', 'Initial setup', '', 'Set up', 3, 'All', 'PT01'],
    ['', 'T010-01', 'Initial setup', 'Create folders', '', '', 'All', ''],
    [20, 'C030', 'Draft Modelling', '', 'Modelled', 10, 'All', ''],
    ['', 'T030-01', 'Draft Modelling', 'Foundations', '', '', 'Struct', ''],
    [],
    ['Prereq ID', 'Prerequisite', 'Why it is needed', 'Usually provided by', 'Needed before (category ID)'],
    ['PT01', 'Template file', 'Sets standards', 'BIM lead', 'C010'],
  ]), 'Task Template');

  const F = 'Finished', P = 'WIP', B = 'Blocked', Wt = 'Waiting on', N = 'Not started';
  const head = ['Task ID', 'Category', 'Discipline', 'Task / component', 'Weight', 'Type', 'Date added', 'Target week'];
  const tail = ['Completed - week', 'Completed - date', 'Responsible (R code)', 'Prerequisite (P IDs)', 'Depends on (Task IDs)'];
  const wd = (n) => new Date(2026, 0, 4 + 7 * n);

  const site = (code, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    [`SITE ${code}`],
    ['Site code', 'Start date', 'Target submission'],
    [code, new Date(2026, 0, 5), new Date(2026, 2, 1)],
    [],
    ['TABLE 1 - TASK LIST AND WEEKLY STATUS'],
    ['', '', '', '', '', '', '', '', 'W01', 'W02', 'W03', 'W04'],
    [...head, wd(1), wd(2), wd(3), wd(4), ...tail],
    ...rows,
    [],
    ['TABLE 2 - PREREQUISITES FOR THIS SITE'],
    ['Prereq ID', 'Prerequisite', 'From template (PT ID)', 'Needed for (Task IDs)', 'Provided by (R code / party)', 'Required by week', 'Status', 'Date received', 'Notes'],
    ['P01', 'Template file', 'PT01', 'T010-01', 'R1', 1, 'Received', new Date(2026, 0, 6), ''],
    ['P02', 'Ceiling void zones', 'PT06', 'T030-02', 'R3', 2, 'Outstanding', '', ''],
    [],
    ['TABLE 3 - WAITING ON / BLOCKED LOG'],
    ['Log ID', 'Week (W no)', 'Task ID', 'Status raised', 'Reason', 'Waiting on (R code / party)', 'Raised date', 'Expected clear date', 'Cleared date', 'Notes'],
    ['L01', 3, 'T030-02', 'Waiting on', 'Ceiling void zones not confirmed', 'R3', new Date(2026, 0, 19), new Date(2026, 0, 26), '', ''],
  ]), code);

  site('A-01', [
    ['C010', 'Initial setup', 'All', '[Category] Initial setup', '', '', '', '', F, F, F, F, '', '', '', '', ''],
    ['T010-01', 'Initial setup', 'All', 'Create folders', 1, 'Regular', '', 1, F, F, F, F, 1, new Date(2026, 0, 11), 'R1', 'P01', ''],
    ['C030', 'Draft Modelling', 'All', '[Category] Draft Modelling', '', '', '', '', N, P, P, P, '', '', '', '', ''],
    ['T030-01', 'Draft Modelling', 'Struct', 'Foundations', 3, 'Regular', '', 2, N, P, F, F, 3, new Date(2026, 0, 25), 'R2', '', 'T010-01'],
    ['T030-02', 'Draft Modelling', 'MEP', 'Ductwork', 2, 'Regular', '', 3, N, N, Wt, Wt, '', '', 'R3', 'P02', ''],
    ['T030-03', 'Draft Modelling', 'Arch', 'Walls', 1, 'Additional', new Date(2026, 0, 20), 4, '', '', N, N, '', '', 'R2', '', ''],
  ]);
  site('A-02', [
    ['C010', 'Initial setup', 'All', '[Category] Initial setup', '', '', '', '', F, F, F, F, '', '', '', '', ''],
    ['T010-01', 'Initial setup', 'All', 'Create folders', 1, 'Regular', '', 1, F, F, F, F, 1, new Date(2026, 0, 11), 'R1', 'P01', ''],
    ['T030-01', 'Draft Modelling', 'Struct', 'Foundations', 3, 'Regular', '', 2, N, N, B, B, '', '', 'R3', '', ''],
  ]);
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

const BUF = fixtureBuffer();
const fakeFile = { name: 'tracker.xlsx', size: BUF.byteLength, arrayBuffer: async () => BUF };

/* ---------------------------- boot the app ---------------------------- */

const S = await import('../js/core/session.js');
const app = await import('../js/views/app.js');
const parser = await import('../js/core/parser.js');
const { analysePortfolio } = await import('../js/core/engine.js');

errors = [];
app.boot();
await wait(120);

check('app boots without errors', errors.length === 0, errors.join(' | '));
check('shell rendered', !!$('.shell'));
check('step 1 rendered', textOf('#app').includes('Load the workbook'));
check('usage rail rendered', textOf('.siderail').includes('Requests this minute'));
check('later steps hidden until a workbook is loaded', !textOf('#app').includes('Choose what to analyse'));

/* ------------------------- step 1: load ------------------------- */

errors = [];
const model = parser.parseWorkbook(await parser.readWorkbook(fakeFile));
S.update((s) => { s.file = { name: fakeFile.name, size: fakeFile.size }; s.model = model; s.confirmed = false; });
await wait(120);

check('loading a workbook logs no errors', errors.length === 0, errors.join(' | '));
check('step 2 appears after loading', textOf('#app').includes('Check what the app read'));
check('three sites listed', model.sites.length === 3, String(model.sites.length));
check('confirmation panel names each site', ['A-01', 'A-02', 'A-03'].every((c) => textOf('#app').includes(c)));
check('confirmation shows how each sheet was matched', textOf('#app').includes('named in register'));
check('site without a detail sheet is shown as unmatched', textOf('#app').includes('— none —'));
check('privacy scan result shown', /privacy scan|may identify/i.test(textOf('#app')));
check('run step still hidden before confirming', !textOf('#app').includes('Choose what to analyse'));

/* ------------------------- step 2: confirm ------------------------- */

errors = [];
const confirmBtn = $$('button').find((b) => /This is correct/.test(b.textContent));
check('confirm button present', !!confirmBtn);
confirmBtn.click();
await wait(140);

check('confirming logs no errors', errors.length === 0, errors.join(' | '));
check('run step appears', textOf('#app').includes('Choose what to analyse'));
check('reports step appears', textOf('#app').includes('Reports'));
check('a checkbox per site plus master', $$('.pick').length === 4, String($$('.pick').length));
check('master analysis option present', !!$('.pick--master'));
check('reset button present', $$('button').some((b) => /Reset/.test(b.textContent)));
check('timeline chart rendered in the rail', textOf('.siderail').includes('Timeline'));
check('computed risks shown without any AI call', textOf('.siderail').includes('Computed risks'));

/* ------------------------- selection ------------------------- */

errors = [];
$$('button').find((b) => b.textContent.trim() === 'Select all').click();
await wait(120);
check('select all ticks every box', S.get().selection.length === 3 && S.get().includeMaster === true,
  `${S.get().selection.length} sites, master=${S.get().includeMaster}`);
check('estimate reacts to the selection', /API calls/.test(textOf('#app')));
check('4 calls estimated for 3 sites plus master', /4/.test(textOf('#app')));

$$('button').find((b) => b.textContent.trim() === 'Clear selection').click();
await wait(120);
check('clear selection empties it', S.get().selection.length === 0 && !S.get().includeMaster);
check('generate button disabled with nothing selected',
  $$('button').find((b) => /Generate insight/.test(b.textContent))?.disabled === true);

const firstBox = $$('.pick input')[0];
firstBox.click();
await wait(120);
check('ticking one box selects one site', S.get().selection.length === 1, JSON.stringify(S.get().selection));
check('selection logs no errors', errors.length === 0, errors.join(' | '));

/* --------------------- generation, with the network stubbed --------------------- */

const portfolio = analysePortfolio(model);
let captured = null;
let callCount = 0;

window.fetch = async (url, opts) => {
  callCount++;
  captured = { url: String(url), body: JSON.parse(opts.body) };
  return {
    ok: true, status: 200,
    headers: { get: () => null },
    json: async () => ({
      candidates: [{ content: { parts: [{ text: JSON.stringify({
        headline: 'Ductwork is the constraint',
        verdict: 'at_risk',
        confidence: 'medium',
        confidenceReason: 'only four weeks of data',
        summary: 'T030-02 has been waiting on R3 for two weeks.',
        bottlenecks: [{ taskId: 'T030-02', issue: 'Waiting on ceiling voids', whoToChase: 'R3', suggestedAction: 'Escalate', urgency: 'high' }],
        actions: [{ action: 'Chase R3', owner: 'R1', byWhen: 'W05', priority: 'high', expectedEffect: 'Unblocks MEP' }],
        dataGaps: ['No target week on some tasks'],
      }) }] } }],
      usageMetadata: { promptTokenCount: 1200, candidatesTokenCount: 400 },
    }),
    text: async () => '',
  };
};
global.fetch = window.fetch;

S.setApiKey('AIzaTESTKEY-not-real-000', false);
S.update((s) => { s.settings.reviewPayload = false; s.selection = ['A-01']; s.includeMaster = false; });
await wait(120);

errors = [];
$$('button').find((b) => /Generate insight/.test(b.textContent)).click();
await wait(400);

check('one API call made', callCount === 1, String(callCount));
check('call went to the selected provider', /generativelanguage\.googleapis\.com/.test(captured.url), captured?.url);
check('report stored under the site code', !!S.get().reports['A-01']);
check('generation logs no errors', errors.length === 0, errors.join(' | '));
check('report rendered', textOf('#app').includes('Ductwork is the constraint'));
check('verdict chip rendered', textOf('#app').includes('at risk'));
check('actions rendered', textOf('#app').includes('Chase R3'));
check('completed site marked done in the picker', $$('.pick--done').length === 1, String($$('.pick--done').length));
check('completed site removed from the selection', !S.get().selection.includes('A-01'));
check('export buttons appear once a report exists', textOf('#app').includes('Excel (.xlsx)'));

/* ---- the boundary that matters: no raw grid, no key, real numbers ---- */

const sent = JSON.stringify(captured.body);
check('API key is not in the request body', !sent.includes('AIzaTESTKEY'));
check('request carries the computed payload', sent.includes('percentByWeight'));
check('request carries no raw weekly grid', !sent.includes('"weekly"'));
check('request tells the model not to recalculate', /do not recalculate/i.test(sent));
check('request carries site codes only, no descriptions of real places',
  !/Marina|Tower|Street/i.test(sent));

const usageAfter = (await import('../js/core/tokens.js')).getUsage();
check('provider-reported usage recorded', usageAfter.input === 1200 && usageAfter.output === 400,
  `${usageAfter.input}/${usageAfter.output}`);
check('usage is not marked estimated when the provider reported it', usageAfter.estimated === false);
check('usage rail shows the tokens', textOf('.siderail').includes('1.6k') || textOf('.siderail').includes('Tokens used'));

/* -------------------------- error handling -------------------------- */

window.fetch = async () => ({
  ok: false, status: 429, headers: { get: () => null },
  text: async () => '{"error":{"message":"Please retry in 30s"}}',
  json: async () => ({}),
});
global.fetch = window.fetch;

S.update((s) => { s.selection = ['A-02']; });
await wait(100);
errors = [];
$$('button').find((b) => /Generate insight/.test(b.textContent)).click();
await wait(400);

check('a 429 does not store a report', !S.get().reports['A-02']);
check('a 429 does not crash the app', !!$('.shell') && errors.length === 0, errors.join(' | '));
check('the earlier report survives the failure', !!S.get().reports['A-01']);
check('rate limit is explained to the user', /rate limit/i.test(textOf('.toasts') || ''), textOf('.toasts'));

/* ---------------------------- exports ---------------------------- */

const { buildWorkbook, buildRTF, buildPrintHTML, reportBlocks } = await import('../js/core/exports.js');
errors = [];
const wbOut = buildWorkbook(portfolio, S.get().reports);
check('excel export builds', wbOut.SheetNames.length > 3, String(wbOut.SheetNames.length));
check('excel export includes the AI report', wbOut.SheetNames.includes('AI reports'));
const rtf = buildRTF(reportBlocks(portfolio, S.get().reports));
check('rtf export builds', rtf.startsWith('{\\rtf1'));
check('rtf includes the narrative', rtf.includes('Ductwork is the constraint'));
const printed = buildPrintHTML(portfolio, S.get().reports);
check('print html builds', printed.includes('<!DOCTYPE html>'));
check('exports log no errors', errors.length === 0, errors.join(' | '));

const pdfBtn = $$('button').find((b) => /PDF \(print\)/.test(b.textContent));
check('pdf button present', !!pdfBtn);
errors = [];
pdfBtn.click();
await wait(80);
check('a blocked popup is reported rather than failing silently',
  /pop-?up/i.test(textOf('.toasts') || ''), textOf('.toasts'));

/* --------------------------- session file --------------------------- */

const file = S.buildSessionFile();
const fileStr = JSON.stringify(file);
check('session file has the expected format tag', file.format === 'bimtrack.session');
check('session file carries the reports', !!file.state.reports['A-01']);
check('session file carries the parsed model', !!file.state.model);
check('API KEY IS NOT IN THE SESSION FILE', !fileStr.includes('AIzaTESTKEY'));
check('no part of the key leaks into the file', !fileStr.includes('AIza'));
check('no key-shaped field survives in the session file', !/"apiKey"|"api_key"/.test(fileStr));
check('session file says the key was excluded', /API keys are deliberately not included/i.test(file.note));

/* --------------------------- persistence --------------------------- */

await wait(600);   // let the debounced autosave land
const before = Object.keys(S.get().reports).length;
S.reset();
await wait(200);
check('reset clears the session', !S.get().model && Object.keys(S.get().reports).length === 0);
const restored = await S.restore();
check('a reset session does not resurrect old reports',
  Object.keys(S.get().reports).length === 0, String(Object.keys(S.get().reports).length));

// Clear the key first, so this genuinely tests that the FILE carries no key
// rather than observing one that was already in memory.
S.clearApiKey();
check('key cleared before the load', !S.hasApiKey());
S.replace(file.state);
await wait(200);
check('loading a session file restores the reports', Object.keys(S.get().reports).length === before, `${Object.keys(S.get().reports).length} vs ${before}`);
check('loading a session file restores the model', !!S.get().model);
check('loading a session file does NOT bring back an API key', !S.hasApiKey());
await wait(120);
check('app re-renders from a loaded session', textOf('#app').includes('Ductwork is the constraint'));

/* --------------------------- edge cases --------------------------- */

errors = [];
S.update((s) => { s.model = { ...s.model, sites: [] }; });
await wait(150);
check('an empty register does not crash the app', !!$('.shell') && errors.length === 0, errors.join(' | '));

S.reset();
await wait(150);
errors = [];
check('reset back to step 1 renders', textOf('#app').includes('Load the workbook') && errors.length === 0, errors.join(' | '));

/* ------------------------------ report ------------------------------ */

console.error = realError;
console.log(`\n${pass} passed, ${fail} failed\n`);
if (failures.length) {
  console.log('Failures:');
  for (const f of failures) console.log('  ✗', f);
  process.exit(1);
}
console.log('UI smoke test passed.');
