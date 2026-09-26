import {
  $, el, mount, icon, toast, openModal, confirmDialog, pickFile, fmtDate,
  fmtNum, clamp, debounce, todayISO, titleCase,
} from '../core/util.js';
import { readWorkbook, parseWorkbook, screenText, xlsxAvailable, STATUS } from '../core/parser.js';
import { analysePortfolio, comparePortfolios, compareSites } from '../core/engine.js';
import {
  computeSiteReport, computeMasterReport, scoreRisks, SITE_SECTIONS, MASTER_SECTIONS,
  PROVENANCE_NOTE, METRIC_DEFINITIONS, RAG_THRESHOLDS, ENGINE, AI, AI_ANALYST,
} from '../core/report.js';
import * as S from '../core/session.js';
import * as T from '../core/tokens.js';
import { PROVIDERS, run as runAI, buildSitePayload, buildMasterPayload, previewPayload, CancelledError, overloadedModels, clearOverloadMemory } from '../core/ai.js';
import { exportExcel, exportWord, exportPDF } from '../core/exports.js';
import {
  renderCurve, renderThroughput, renderSiteBars, renderStatusGrid, renderStatusMix,
  renderCategoryProgress, renderBlockerAges, renderResourceLoad, meter, statusLegend,
} from '../core/charts.js';
import {
  renderRagMatrix, renderKpiDashboard, renderScheduleAcrossSites, renderCompletionVsTime,
  renderMilestoneTimeline, renderCategoryBySite, renderCategoryProgressReport,
  renderPartyDependency, renderConstraintAgeing, renderRiskHeatMap,
  renderResourceAcrossSites, renderFindingsBySite, renderIpi,
} from '../core/report-charts.js';

const MASTER = S.MASTER_KEY;

/* Live run state, deliberately outside the session — a cancel token has no
   meaning after a reload, and persisting it would leave a dead "running" flag. */
let running = false;
let controller = null;
let queue = [];
let runSelection = [];
let currentKey = null;

let cachedPortfolio = null;
let cachedStamp = '';
let cachedPrev = null;
let cachedPrevStamp = '';

/* Which sections and reports the user has collapsed. Kept in memory rather than
   the session: it is a view preference, and re-expanding everything on reload
   is the safer default. Exports ignore it entirely — they are built from data,
   never from what happens to be visible. */
const collapsed = new Set();
let retryStatus = null;    // live countdown while a provider is overloaded

/** The previous week's analysis, recomputed by the same engine. */
function previousPortfolio() {
  const prev = S.get().previous;
  if (!prev?.model) return null;
  const stamp = `${prev.loadedAt}|${prev.model.sites?.length}`;
  if (cachedPrev && cachedPrevStamp === stamp) return cachedPrev;
  try { cachedPrev = analysePortfolio(prev.model); } catch { cachedPrev = null; }
  cachedPrevStamp = stamp;
  return cachedPrev;
}

/**
 * The computed report for one site. Recomputed from the current workbook so the
 * figures and the narrative can never drift apart within a run.
 */
function siteReportFor(code) {
  const p = portfolio();
  const st = S.get();
  const site = st.model?.sites.find((x) => x.code === code);
  const a = p?.sites.find((x) => x.code === code);
  if (!site || !a) return null;
  return computeSiteReport(site, a, st.model.template, { reportDate: todayISO() });
}

/**
 * The master report over the sites the user actually ticked — never the whole
 * workbook. The template is explicit that a portfolio report must not include,
 * count or compare a site that was not selected.
 */
function masterReportFor(codes) {
  const p = portfolio();
  if (!p) return null;
  const reports = (codes || []).map(siteReportFor).filter(Boolean);
  if (!reports.length) return null;
  const scoped = { ...p, sites: p.sites.filter((s) => codes.includes(s.code)) };
  return computeMasterReport(reports, scoped, { reportDate: todayISO() });
}

function portfolio() {
  const st = S.get();
  if (!st.model) return null;
  const stamp = `${st.model.parsedAt}|${st.model.sites.length}`;
  if (cachedPortfolio && cachedStamp === stamp) return cachedPortfolio;
  cachedPortfolio = analysePortfolio(st.model);
  cachedStamp = stamp;
  return cachedPortfolio;
}

/* ========================================================================
   Shell
   ======================================================================== */

let host, rail;

export function boot() {
  const app = $('#app');
  host = el('main', { class: 'work' });
  rail = el('div', { class: 'rail__scroll' });

  const railEl = el('aside', { class: 'rail' }, [
    el('div', { class: 'rail__brand' }, [
      el('span', { class: 'rail__mark', text: 'BIM' }),
      el('span', {}, [
        el('div', { class: 'rail__name', text: 'Multi-Site Delivery Tracker' }),
        el('div', { class: 'rail__sub', text: 'Codes only — runs in your browser' }),
      ]),
    ]),
    rail,
    el('div', { class: 'rail__foot' }, [
      el('button', { class: 'btn btn--sm', onclick: onSaveSession }, [icon('save', 13), 'Save']),
      el('button', { class: 'btn btn--sm', onclick: onLoadSession }, [icon('upload', 13), 'Load']),
      el('button', { class: 'btn btn--sm btn--wide', onclick: openSettings }, [icon('settings', 13), 'Settings & API key']),
    ]),
  ]);

  mount(app, el('div', { class: 'shell grid-field' }, [railEl, host]));

  S.subscribe(debounce(render, 30));
  S.restore().then((found) => {
    render();
    if (found && S.get().model) toast('Previous session restored from this browser.', 'sign');
    if (!xlsxAvailable()) toast('The spreadsheet reader did not load — imports will not work.', 'survey', 9000);
  });
}

/*
 * The work column is rebuilt on each state change. Rebuilding empties it first,
 * which collapses the page height and makes the browser clamp the scroll
 * position to the top — the jump you saw after ticking a checkbox. The scroll
 * position is captured before the rebuild and restored after it, and the
 * column's height is held while it is empty so nothing collapses in between.
 */
function render() {
  const st = S.get();
  const y = window.scrollY;
  const h = host.offsetHeight;
  if (h) host.style.minHeight = `${h}px`;

  mount(host,
    workbar(st),
    stepImport(st),
    st.model ? stepConfirm(st) : null,
    st.model && st.confirmed ? stepPeriod(st) : null,
    st.model && st.confirmed ? stepRun(st) : null,
    st.model && st.confirmed ? stepReports(st) : null,
  );
  renderRail();

  host.style.minHeight = '';
  if (y) window.scrollTo(0, y);
}

function workbar(st) {
  return el('div', { class: 'workbar' }, [
    el('div', { class: 'grow' }, [
      el('div', { class: 'workbar__title', text: st.file ? st.file.name.replace(/\.[^.]+$/, '') : 'New analysis' }),
      el('div', {
        class: 'workbar__sub',
        text: st.model
          ? `${st.model.sites.length} sites · ${Object.keys(st.reports).length} report${Object.keys(st.reports).length === 1 ? '' : 's'}${st.previous ? ` · compared with ${st.previous.fileName}` : ''}`
          : 'Load a workbook to begin',
      }),
    ]),
    el('button', { class: 'btn btn--sm', onclick: () => toggleAll(true) }, ['Collapse all']),
    el('button', { class: 'btn btn--sm', onclick: () => toggleAll(false) }, ['Expand all']),
    el('button', { class: 'btn btn--sm btn--clear', onclick: onClearSession }, [icon('trash', 13), 'Clear session']),
  ]);
}

function toggleAll(collapse) {
  const st = S.get();
  const keys = ['step:1', 'step:2', 'step:3', 'step:4', 'step:5',
    ...Object.keys(st.reports).map((k) => `report:${k}`)];
  for (const k of keys) { if (collapse) collapsed.add(k); else collapsed.delete(k); }
  render();
}

async function onClearSession() {
  const st = S.get();
  const n = Object.keys(st.reports || {}).length;
  const ok = await confirmDialog({
    title: 'Clear this session?',
    message: `This removes the loaded workbook${n ? `, ${n} generated report${n === 1 ? '' : 's'}` : ''}${st.previous ? ', the previous week you loaded' : ''} and your period notes from this browser. Your provider settings and API key are kept. If you want any of this later, press Save in the left panel first.`,
    confirmLabel: 'Clear session', tone: 'danger',
  });
  if (!ok) return;
  cachedPortfolio = null; cachedPrev = null;
  collapsed.clear();
  await S.clearSession();
  window.scrollTo(0, 0);
  toast('Session cleared. Load a workbook to start a new analysis.', 'sign');
}

/* ========================================================================
   Step 1 — import
   ======================================================================== */

function stepImport(st) {
  /* ---- this week's workbook ---- */
  const drop = el('div', { class: `dropzone${st.file ? ' is-loaded' : ''}` }, [
    el('div', { style: { marginBottom: '6px' } }, [icon('upload', 22)]),
    el('h3', { text: st.file ? `This week: ${st.file.name}` : 'This week\u2019s workbook' }),
    el('p', { class: 'small muted', style: { marginTop: '4px' } }, [
      st.file ? 'Click or drop to replace it.' : 'Drop the tracker here, or click to choose. Register, template, then one sheet per site.',
    ]),
  ]);

  const load = async (file) => {
    if (!file) return;
    if (!xlsxAvailable()) { toast('The spreadsheet reader did not load. Reload the page.', 'survey'); return; }
    try {
      const wb = await readWorkbook(file);
      const model = parseWorkbook(wb);
      cachedPortfolio = null;
      S.update((s) => {
        s.file = { name: file.name, size: file.size };
        s.model = model;
        s.confirmed = false;
        s.selection = [];
        s.includeMaster = false;
        // Reports from a previous workbook would be misleading against new data.
        if (Object.keys(s.reports).length) s.reports = {};
      });
      toast(`Read ${model.sites.length} site${model.sites.length === 1 ? '' : 's'} from ${file.name}.`, 'sign');
    } catch (e) {
      toast(e.message || 'That file could not be read.', 'survey', 8000);
    }
  };
  wireDrop(drop, load, '.xlsx,.xlsm,.xls');

  /* ---- previous week, for comparison ---- */
  const prev = st.previous;
  const prevDrop = el('div', { class: `dropzone dropzone--small${prev ? ' is-loaded-prev' : ''}` }, [
    el('div', { style: { marginBottom: '4px' } }, [icon('upload', 18)]),
    el('h3', { text: prev ? `Previous: ${prev.fileName}` : 'Previous week (optional)' }),
    el('p', { class: 'small muted', style: { marginTop: '4px' } }, [
      prev
        ? `${prev.kind === 'session' ? `Session with ${Object.keys(prev.reports || {}).length} report${Object.keys(prev.reports || {}).length === 1 ? '' : 's'}` : 'Workbook'} \u00b7 loaded ${fmtDate((prev.loadedAt || '').slice(0, 10))}`
        : 'Last week\u2019s workbook (.xlsx) or saved session (.json). Used to show what changed.',
    ]),
  ]);

  const loadPrev = async (file) => {
    if (!file) return;
    try {
      const entry = await readPrevious(file);
      cachedPrev = null;
      S.update((s) => { s.previous = entry; });
      const cur = portfolio();
      const pp = previousPortfolio();
      const shared = cur && pp ? cur.sites.filter((x) => pp.sites.some((y) => y.code === x.code)).length : 0;
      toast(`Previous week loaded from ${file.name}${cur ? ` \u2014 ${shared} site${shared === 1 ? '' : 's'} can be compared` : ''}.`, 'sign', 6000);
    } catch (e) {
      toast(e.message || 'That file could not be read.', 'survey', 8000);
    }
  };
  wireDrop(prevDrop, loadPrev, '.xlsx,.xlsm,.xls,.json,application/json');

  const prevHelp = el('div', { class: 'notice', style: { alignSelf: 'stretch' } }, [
    el('h4', { text: 'Why load the previous week' }),
    el('p', { class: 'small', text: 'Both weeks go through the same engine, so the change \u2014 tasks finished, blockers cleared, forecast movement \u2014 is calculated rather than guessed. A saved session also brings last week\u2019s reports, so the review can say whether its actions were acted on.' }),
    prev ? el('button', {
      class: 'btn btn--sm', style: { marginTop: '6px' },
      onclick: () => { cachedPrev = null; S.update((s) => { s.previous = null; }); toast('Previous week removed.'); },
    }, [icon('trash', 13), 'Remove previous week']) : null,
  ]);

  return section('1', 'Load the workbook', el('div', { class: 'sheet__body stack' }, [
    drop,
    el('div', { class: 'prevload' }, [prevDrop, prevHelp]),
  ]), st.file ? `${st.file.name}${prev ? ' + previous week' : ''}` : '');
}

function wireDrop(zone, onFile, accept) {
  zone.addEventListener('click', async () => onFile(await pickFile(accept)));
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('is-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
  zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('is-over'); onFile(e.dataTransfer?.files?.[0]); });
}

/**
 * Read a previous week from either format.
 *  - A saved session (.json): brings the parsed workbook AND its reports.
 *  - A workbook (.xlsx): brings the figures only, parsed like this week's.
 * Either way the comparison is computed by the same engine, never inferred.
 */
async function readPrevious(file) {
  const name = file.name || 'previous';
  if (/\.json$/i.test(name)) {
    let parsed;
    try { parsed = JSON.parse(await file.text()); }
    catch { throw new Error('That .json is not a session file. Use a file saved with the Save button, or last week\u2019s .xlsx.'); }
    if (parsed?.format !== S.SESSION_FORMAT || !parsed.state) {
      throw new Error('That .json is not a session saved by this app.');
    }
    const st = S.stripSecrets(parsed.state);
    if (!st.model) throw new Error('That session has no workbook in it, so there is nothing to compare against.');
    return { kind: 'session', fileName: name, loadedAt: new Date().toISOString(), model: st.model, reports: st.reports || {} };
  }
  if (!xlsxAvailable()) throw new Error('The spreadsheet reader did not load.');
  const model = parseWorkbook(await readWorkbook(file));
  if (!model.sites.length) throw new Error('No sites found in that workbook.');
  return { kind: 'workbook', fileName: name, loadedAt: new Date().toISOString(), model, reports: {} };
}

/* ========================================================================
   Step 2 — confirm what was read
   ======================================================================== */

function stepConfirm(st) {
  const m = st.model;
  const p = portfolio();
  const body = el('div', { class: 'sheet__body stack' });

  body.appendChild(el('div', { class: 'band' }, [
    band('Sites in register', String(m.sites.length), m.register ? `from "${m.register.sheetName}"` : 'no register found', m.register ? 'blueprint' : 'survey'),
    band('With a detail sheet', String(m.sites.filter((s) => s.detail).length), `${m.sites.filter((s) => !s.detail).length} without`, 'blueprint'),
    band('Template categories', String(m.template.categories.length), m.template.sheetName ? `from "${m.template.sheetName}"` : 'none found', 'blueprint'),
    band('Tasks read', String(p ? p.sites.reduce((n, s) => n + (s.taskCount || 0), 0) : 0), 'across all sites', 'blueprint'),
    band('Weeks of status', String(Math.max(0, ...m.sites.map((s) => s.weekCount || 0))), 'longest site', 'blueprint'),
  ]));

  body.appendChild(el('div', { class: 'tablewrap' }, [
    el('table', { class: 'data' }, [
      el('thead', {}, [el('tr', {}, [
        el('th', { text: 'Site' }), el('th', { text: 'Detail sheet' }), el('th', { text: 'Matched by' }),
        el('th', { class: 'n', text: 'Tasks' }), el('th', { class: 'n', text: 'Weeks' }),
        el('th', { class: 'n', text: 'Reported to' }), el('th', { text: 'Target' }),
      ])]),
      el('tbody', {}, m.sites.map((s) => {
        const a = p?.sites.find((x) => x.code === s.code);
        const conf = s.matchedBy === 'named in register' ? 'sign'
          : s.matchedBy ? 'hivis' : 'survey';
        return el('tr', {}, [
          el('td', { class: 'strong', text: s.code }),
          el('td', { class: 'small', text: s.detail?.sheetName || '— none —' }),
          el('td', {}, [el('span', { class: 'chip', dataset: { tone: conf }, text: s.matchedBy || 'not matched' })]),
          el('td', { class: 'n', text: String(s.taskCount) }),
          el('td', { class: 'n', text: String(s.weekCount) }),
          el('td', { class: 'n small', text: a?.reportingWeek ? `${a.reportingWeek.label} (${fmtDate(a.reportingWeek.end)})` : '—' }),
          el('td', { class: 'small', text: s.target ? fmtDate(s.target) : '—' }),
        ]);
      })),
    ]),
  ]));

  if (m.problems.length) {
    body.appendChild(notice('survey', `${m.problems.length} thing${m.problems.length === 1 ? '' : 's'} to check in the workbook`,
      m.problems.slice(0, 8), m.problems.length > 8 ? `…and ${m.problems.length - 8} more.` : ''));
  }
  if (m.notes.length) {
    body.appendChild(notice('hivis', 'Notes', m.notes.slice(0, 6)));
  }

  // Privacy screen over everything that would be transmitted.
  const strings = [];
  for (const s of m.sites) {
    strings.push({ where: `${s.code} description`, text: s.description });
    strings.push({ where: `${s.code} notes`, text: s.notes });
    for (const t of s.detail?.tasks || []) strings.push({ where: `${s.code} ${t.id}`, text: t.name });
    for (const q of s.detail?.prereqs || []) strings.push({ where: `${s.code} ${q.id}`, text: q.name });
    for (const l of s.detail?.log || []) strings.push({ where: `${s.code} ${l.id}`, text: `${l.reason} ${l.waitingOn}` });
  }
  const hits = screenText(strings);
  if (hits.length) {
    body.appendChild(notice('survey', `${hits.length} name${hits.length === 1 ? '' : 's'} in your text may identify a real place or person`,
      hits.slice(0, 6).map((h) => `${h.where}: "${h.phrase}"`),
      'These would be sent to the AI provider. If any is a real client, site or person, change it in the workbook and re-load. This scan is a safety net, not a guarantee — you remain the check.'));
  } else {
    body.appendChild(notice('sign', 'Privacy scan found nothing that looks like a real name',
      [], 'Site and resource codes only. This scan cannot catch everything, so the responsibility remains yours.'));
  }

  body.appendChild(el('div', { class: 'row', style: { marginTop: '4px' } }, [
    st.confirmed
      ? el('span', { class: 'chip', dataset: { tone: 'sign' }, text: '✓ Confirmed' })
      : el('button', {
          class: 'btn btn--primary',
          onclick: () => { S.update((s) => { s.confirmed = true; }); toast('Confirmed. Choose what to analyse.', 'sign'); },
        }, [icon('check', 14), 'This is correct — continue']),
    st.confirmed
      ? el('button', { class: 'btn btn--sm', onclick: () => S.update((s) => { s.confirmed = false; }) }, ['Review again'])
      : null,
  ]));

  return section('2', 'Check what the app read', body);
}


/* ========================================================================
   Step 3 — reporting period and what came before
   ======================================================================== */

function stepPeriod(st) {
  const p = portfolio();
  const per = st.period || {};
  const body = el('div', { class: 'sheet__body stack' });

  const field = (key, label, hint, type = 'text', placeholder = '') => {
    const node = el('input', { class: 'input', type, value: per[key] || '', placeholder });
    node.addEventListener('change', () => S.update((x) => { x.period = { ...x.period, [key]: node.value }; }, { silent: true }));
    return el('div', { class: 'field' }, [
      el('label', { text: label }), node,
      hint ? el('span', { class: 'hint', text: hint }) : null,
    ]);
  };

  const notes = el('textarea', {
    class: 'textarea',
    placeholder: 'What happened since the last review that the spreadsheet does not show? Decisions taken, people moved, client instructions, anything the status columns cannot say.',
  });
  notes.value = per.notes || '';
  notes.addEventListener('change', () => S.update((x) => { x.period = { ...x.period, notes: notes.value }; }, { silent: true }));

  const carry = el('input', { type: 'checkbox', checked: per.carryPrevious !== false ? true : null });
  carry.addEventListener('change', () => S.update((x) => { x.period = { ...x.period, carryPrevious: carry.checked }; }));

  const stored = Object.values(st.reports || {});
  const latest = stored.sort((a, b) => String(b.at).localeCompare(String(a.at)))[0];

  body.appendChild(el('p', { class: 'small muted', text: 'Optional, but it is what turns a snapshot into a trend. Anything here is sent with every analysis so the model can say what has changed rather than describing this week in isolation.' }));

  body.appendChild(el('div', { class: 'formgrid' }, [
    field('label', 'Label for this review', 'e.g. "Week 6 review" or "March monthly".', 'text', 'Week 6 review'),
    field('previousDate', 'Date of the previous analysis', 'Leave blank if this is the first.', 'date'),
    field('previousWeek', 'Week the previous analysis covered', 'e.g. W04.', 'text', 'W04'),
  ]));

  body.appendChild(el('div', { class: 'field' }, [
    el('label', { text: 'What has happened since the last review' }), notes,
    el('span', { class: 'hint', text: 'Keep it code-only: R3, A-02. No client or personal names — this is sent to the AI.' }),
  ]));

  body.appendChild(el('label', { class: 'check' }, [carry, el('span', {}, [
    el('strong', { text: 'Carry the previous report forward as context. ' }),
    el('span', {
      class: 'muted',
      text: latest
        ? `The stored review for each site is included so the model can compare. Most recent: ${latest.title}, ${fmtDate((latest.at || '').slice(0, 10))}.`
        : 'Nothing stored yet — this takes effect once you have generated a report and come back next week.',
    }),
  ])]));

  if (latest) {
    body.appendChild(el('div', { class: 'notice prevbox' }, [
      el('h4', { text: 'Previous review on file' }),
      el('p', { class: 'small', text: latest.result?.conclusions?.statement || latest.result?.introduction || latest.result?.headline || 'Stored report available.' }),
      el('p', { class: 'xs muted', style: { margin: '4px 0 0' }, text: 'This is carried into the next run so progress can be measured against it, not just against the plan.' }),
    ]));
  }

  const filled = [per.label, per.previousDate, per.previousWeek, per.notes].filter(Boolean).length;
  return section('3', 'Reporting period', body, filled ? `${filled} of 4 filled` : 'optional');
}

/* ========================================================================
   Step 4 — choose and run
   ======================================================================== */

function stepRun(st) {
  const p = portfolio();
  const body = el('div', { class: 'sheet__body stack' });
  const done = new Set(Object.keys(st.reports));

  const toggle = (key) => {
    S.update((s) => {
      if (key === MASTER) { s.includeMaster = !s.includeMaster; return; }
      s.selection = s.selection.includes(key)
        ? s.selection.filter((k) => k !== key)
        : [...s.selection, key];
    });
  };

  const boxes = p.sites.map((s) => {
    const isDone = done.has(s.code);
    const checked = st.selection.includes(s.code);
    const cb = el('input', { type: 'checkbox', checked: checked ? true : null, disabled: running ? true : null });
    cb.addEventListener('change', () => toggle(s.code));
    return el('label', {
      class: `pick${isDone ? ' pick--done' : ''}${currentKey === s.code ? ' pick--active' : ''}`,
    }, [
      cb,
      el('span', { class: 'grow' }, [
        el('span', { class: 'strong', text: s.code }),
        el('span', { class: 'xs dim', text: s.noData ? ' · no weekly data' : ` · ${s.finishedCount}/${s.liveCount} done` }),
      ]),
      isDone ? el('span', { class: 'xs', style: { color: 'var(--sign)' }, text: '✓ done' }) : null,
      currentKey === s.code ? el('span', { class: 'spin' }) : null,
    ]);
  });

  const masterDone = done.has(MASTER);
  const masterCb = el('input', { type: 'checkbox', checked: st.includeMaster ? true : null, disabled: running ? true : null });
  masterCb.addEventListener('change', () => toggle(MASTER));
  const masterBox = el('label', {
    class: `pick pick--master${masterDone ? ' pick--done' : ''}${currentKey === MASTER ? ' pick--active' : ''}`,
  }, [
    masterCb,
    el('span', { class: 'grow' }, [
      el('span', { class: 'strong', text: 'Master analysis' }),
      el('span', { class: 'xs dim', text: ' · compares every site, finds systemic issues' }),
    ]),
    masterDone ? el('span', { class: 'xs', style: { color: 'var(--sign)' }, text: '✓ done' }) : null,
    currentKey === MASTER ? el('span', { class: 'spin' }) : null,
  ]);

  const selectedKeys = [...st.selection, ...(st.includeMaster ? [MASTER] : [])];
  const est = estimateSelection(st, p, selectedKeys);
  const cap = T.capacityCheck(st.settings.provider, est);

  body.appendChild(el('div', { class: 'row row--wrap', style: { gap: '8px' } }, [
    el('button', {
      class: 'btn btn--sm', disabled: running ? true : null,
      onclick: () => S.update((s) => { s.selection = p.sites.map((x) => x.code); s.includeMaster = true; }),
    }, ['Select all']),
    el('button', {
      class: 'btn btn--sm', disabled: running ? true : null,
      onclick: () => S.update((s) => { s.selection = p.sites.filter((x) => !done.has(x.code)).map((x) => x.code); s.includeMaster = !done.has(MASTER); }),
    }, ['Select what is left']),
    el('button', {
      class: 'btn btn--sm', disabled: running ? true : null,
      onclick: () => S.update((s) => { s.selection = []; s.includeMaster = false; }),
    }, ['Clear selection']),
    el('span', { class: 'grow' }),
    el('button', {
      class: 'btn btn--sm btn--ghost', style: { color: 'var(--survey)' }, disabled: running ? true : null,
      onclick: async () => {
        const ok = await confirmDialog({
          title: 'Reset the run?',
          message: `This clears the ${Object.keys(st.reports).length} generated report${Object.keys(st.reports).length === 1 ? '' : 's'} and the current selection. The workbook stays loaded. Anything you have not saved to a file or exported is lost.`,
          confirmLabel: 'Reset', tone: 'danger',
        });
        if (!ok) return;
        S.update((s) => { s.reports = {}; s.selection = []; s.includeMaster = false; });
        toast('Reports cleared.');
      },
    }, [icon('refresh', 13), 'Reset']),
  ]));

  body.appendChild(el('div', { class: 'picks' }, [...boxes, masterBox]));

  body.appendChild(el('div', { class: 'notice', dataset: { tone: cap.level === 'high' ? 'survey' : (cap.level === 'medium' ? 'hivis' : 'blueprint') } }, [
    el('div', { class: 'row row--wrap', style: { gap: '14px' } }, [
      stat('Selected', `${selectedKeys.length}`),
      stat('API calls', `${est.calls}`),
      stat('Est. input', T.formatTokens(est.input)),
      stat('Est. total', T.formatTokens(est.total)),
    ]),
    el('p', { class: 'small', style: { margin: '8px 0 0' }, text: cap.message }),
    el('p', { class: 'xs muted', style: { margin: '4px 0 0' }, text: 'Token figures are estimates — real tokenisation is model-specific and cannot be computed in the browser. Actual usage replaces the estimate once each call returns.' }),
  ]));

  const runBtn = el('button', { class: 'btn btn--primary', disabled: (!selectedKeys.length && !running) ? true : null });
  mount(runBtn, running ? el('span', { class: 'spin' }) : icon('insights', 14),
    running ? `Cancel (${queue.length} left)` : `Generate insight${selectedKeys.length > 1 ? `s (${selectedKeys.length})` : ''}`);
  runBtn.addEventListener('click', () => (running ? cancelRun() : startRun()));

  // Stated before the run, not only inside the report: the user should know
  // which parts of what they are about to generate are calculated and which
  // are a model's judgement.
  body.appendChild(el('div', { class: 'notice', dataset: { tone: 'blueprint' } }, [
    el('h4', { text: 'What the app calculates, and what the AI judges' }),
    el('p', { class: 'small', text: PROVENANCE_NOTE }),
    el('div', { class: 'row row--wrap', style: { gap: '6px', marginTop: '6px' } }, [
      el('span', { class: 'chip', dataset: { tone: 'blueprint' }, text: 'ENGINE' }),
      el('span', { class: 'xs muted', text: `${SITE_SECTIONS.filter((x) => x.by === ENGINE).length} of 14 sections` }),
      el('span', { class: 'chip', dataset: { tone: 'plum' }, text: 'AI' }),
      el('span', { class: 'xs muted', text: 'the conclusion' }),
      el('span', { class: 'chip', dataset: { tone: 'hivis' }, text: 'AI + ANALYST' }),
      el('span', { class: 'xs muted', text: 'executive summary, risks and actions — check before issuing' }),
    ]),
  ]));

  body.appendChild(el('div', {
    id: 'retrybar', class: `retrybar${retryStatus ? '' : ' hidden'}`,
    text: retryStatus ? `Provider busy — retrying in ${retryStatus.secondsLeft}s` : '',
  }));

  body.appendChild(el('div', { class: 'row row--wrap' }, [
    runBtn,
    el('button', {
      class: 'btn', disabled: (running || !selectedKeys.length) ? true : null,
      onclick: () => showPayload(selectedKeys[0]),
    }, ['Inspect what would be sent']),
    !S.hasApiKey() ? el('span', { class: 'chip', dataset: { tone: 'hivis' }, text: 'No API key set' }) : null,
  ]));

  return section('4', 'Choose what to analyse', body, `${Object.keys(st.reports).length} of ${p.sites.length + 1} done`);
}

/**
 * Attach the reporting-period context and, where one exists, a compact summary
 * of the last review of the same scope. Only the parts a comparison needs are
 * carried — a whole previous report would double the payload for little gain.
 */
function withPeriod(payload, key) {
  const st = S.get();
  const per = st.period || {};

  // Last week's review of this same scope: from this session if one was run,
  // otherwise from a previous-week session file if one was loaded.
  const ownPrev = per.carryPrevious !== false ? st.reports?.[key] : null;
  const prevReport = ownPrev || st.previous?.reports?.[key] || null;

  // Computed change against the previous week, where one is loaded.
  let computedChange = null;
  const pp = previousPortfolio();
  const cur = portfolio();
  if (pp && cur) {
    if (key === MASTER) {
      computedChange = comparePortfolios(pp, cur);
    } else {
      const a = cur.sites.find((x) => x.code === key);
      const b = pp.sites.find((x) => x.code === key);
      computedChange = a && b
        ? compareSites(b, a)
        : { comparable: false, reason: 'site was not in the previous week\u2019s workbook' };
    }
  }

  const anything = per.label || per.previousDate || per.previousWeek || per.notes || prevReport || computedChange;
  if (!anything) return payload;

  return {
    ...payload,
    previousPeriod: {
      reviewLabel: per.label || undefined,
      previousAnalysisDate: per.previousDate || (prevReport ? String(prevReport.at).slice(0, 10) : undefined),
      previousWeekCovered: per.previousWeek || undefined,
      previousSource: st.previous?.fileName || undefined,
      whatHappenedSince: per.notes || undefined,
      computedChange: computedChange || undefined,
      lastReview: prevReport ? {
        verdict: prevReport.result?.conclusions?.verdict || undefined,
        statement: prevReport.result?.conclusions?.statement || prevReport.result?.introduction || undefined,
        actionsRaised: (prevReport.result?.additional?.actions || []).slice(0, 6),
        nextStepsSet: (prevReport.result?.conclusions?.nextSteps || []).slice(0, 5),
      } : undefined,
    },
  };
}

function estimateSelection(st, p, keys) {
  let input = 0;
  for (const k of keys) {
    if (k === MASTER) {
        const mm = masterReportFor(st.selection.length ? st.selection : p.sites.map((x) => x.code));
      if (mm) input += T.estimateTokens(JSON.stringify(withPeriod(buildMasterPayload(mm), MASTER))) + 1800;
      continue;
    }
    // A selection can outlive the workbook it was made against — for instance
    // after loading a session file recorded from a different one. Skip codes
    // that no longer exist rather than estimating a payload for a ghost.
    const sr = siteReportFor(k);
    if (!sr) continue;
    input += T.estimateTokens(JSON.stringify(withPeriod(buildSitePayload(sr), k))) + 1800;   // + schema and instructions
  }
  const maxTokens = Number(st.settings.maxTokens) || 4096;
  return {
    calls: keys.length,
    input,
    output: Math.ceil(maxTokens * 0.7) * keys.length,
    total: input + Math.ceil(maxTokens * 0.7) * keys.length,
  };
}

/* ------------------------------ running ------------------------------ */

async function startRun() {
  const st = S.get();
  const p = portfolio();
  const keys = [...st.selection, ...(st.includeMaster ? [MASTER] : [])];
  if (!keys.length) return;
  if (!S.hasApiKey()) {
    toast('No API key in this page. Keys are not kept after a refresh unless you tick "Keep the key if I refresh the page" in Settings.', 'hivis', 9000);
    openSettings();
    return;
  }

  if (st.settings.reviewPayload) {
    const ok = await showPayload(keys[0], true, keys.length);
    if (!ok) return;
  }

  running = true;
  controller = new AbortController();
  // Frozen at run start: sites drop out of the selection as they complete, and
  // the master must still cover everything that was ticked.
  runSelection = [...st.selection];
  queue = [...keys];
  render();

  let okCount = 0;
  let failed = null;

  while (queue.length) {
    const key = queue[0];
    currentKey = key;
    render();
    try {
      const site = key === MASTER ? null : p.sites.find((s) => s.code === key);
      if (key !== MASTER && !site) {
        // Silently dropping it would leave the box ticked forever.
        S.update((s) => { s.selection = s.selection.filter((k) => k !== key); });
        queue.shift();
        continue;
      }
      const computed = key === MASTER
        ? masterReportFor(runSelection.length ? runSelection : p.sites.map((x) => x.code))
        : siteReportFor(key);
      const payload = withPeriod(
        key === MASTER ? buildMasterPayload(computed) : buildSitePayload(computed), key);
      const report = await runAI({
        kind: key === MASTER ? 'master' : 'site',
        payload, settings: S.get().settings, signal: controller.signal,
        onStatus: showRetry,
      });
      showRetry(null);
      report.periodContext = payload.previousPeriod || null;
      // Keep the computed half with the narrative half: the report is the two
      // together, and exports must not depend on the workbook still being loaded.
      report.computed = computed;
      if (report.result?.risks) report.result.risks = scoreRisks(report.result.risks);
      if (report.fellBack && !report.skippedPrimary) {
        toast(`${report.requestedModel} was overloaded, so ${key === MASTER ? 'the master' : key} was generated with ${report.model} instead. The rest of this run will go straight to ${report.model}.`, 'hivis', 10000);
      }
      refreshRail();
      // Persist after EVERY call, so a crash or a cancel keeps what is done.
      S.update((s) => {
        s.reports = { ...s.reports, [key]: report };
        s.selection = s.selection.filter((k) => k !== key);
        if (key === MASTER) s.includeMaster = false;
      });
      okCount++;
      queue.shift();
    } catch (e) {
      if (e instanceof CancelledError || controller.signal.aborted) {
        toast(`Cancelled. ${okCount} report${okCount === 1 ? '' : 's'} completed and saved.`, 'hivis');
        break;
      }
      failed = e;
      break;
    }
  }

  running = false;
  currentKey = null;
  controller = null;
  showRetry(null);
  const left = queue.length;
  queue = [];
  render();

  if (failed) {
    toast(failed.message, 'survey', 10000);
    if (okCount) toast(`${okCount} report${okCount === 1 ? '' : 's'} finished before the error and ${left ? 'are' : 'is'} saved. Use Continue to pick up where it stopped.`, 'hivis', 9000);
  } else if (okCount && !left) {
    toast(`${okCount} report${okCount === 1 ? '' : 's'} generated.`, 'sign');
  }
}

function cancelRun() {
  controller?.abort();
  queue = [];
}

/**
 * Updates the retry banner in place. Deliberately NOT a full render: a countdown
 * ticks every second, and rebuilding the page that often is exactly what made
 * it jump around before.
 */
function showRetry(status) {
  retryStatus = status;
  const node = document.getElementById('retrybar');
  if (node) {
    if (!status) {
      node.classList.add('hidden');
      node.textContent = '';
    } else {
      node.classList.remove('hidden');
      node.textContent = (status.secondsLeft
        ? `Provider busy \u2014 ${status.why}. Retrying in ${status.secondsLeft}s (attempt ${status.attempt} of ${status.maxAttempts}). You can cancel.`
        : `${status.why}\u2026`)
        + (status.providerMessage ? ` Google said: \u201c${status.providerMessage}\u201d` : '');
    }
  }
  refreshRail();
}

/*
 * Re-render only the rail, preserving its own scroll position. The rail used to
 * be redrawn only on state changes, so during a retry it kept showing "0 / 15
 * this minute" while requests were actually firing.
 */
let railTick = 0;
function refreshRail() {
  if (!rail) return;
  const now = Date.now();
  if (now - railTick < 900) return;   // at most about once a second
  railTick = now;
  const y = rail.scrollTop;
  try { renderRail(); } catch (e) { console.error('rail refresh failed', e); }
  rail.scrollTop = y;
}

/* ========================================================================
   Step 5 — reports
   ======================================================================== */

function stepReports(st) {
  const p = portfolio();
  const keys = Object.keys(st.reports);
  const body = el('div', { class: 'sheet__body stack' });

  if (!keys.length) {
    body.appendChild(el('div', { class: 'empty' }, [
      el('h3', { text: 'No reports yet' }),
      el('p', { text: 'Tick the sites you want above and press Generate. Each one is saved the moment it finishes, so a crash or a cancel never loses completed work.' }),
    ]));
  } else {
    body.appendChild(el('div', { class: 'row row--wrap' }, [
      el('button', { class: 'btn', onclick: () => exportExcel(p, st.reports) }, [icon('download', 13), 'Excel (.xlsx)']),
      el('button', { class: 'btn', onclick: () => exportWord(p, st.reports) }, [icon('download', 13), 'Word (.rtf)']),
      el('button', {
        class: 'btn',
        onclick: () => { if (!exportPDF(p, st.reports)) toast('Your browser blocked the print window. Allow pop-ups for this page.', 'hivis', 8000); },
      }, [icon('download', 13), 'PDF (print)']),
      el('span', { class: 'grow' }),
      el('span', { class: 'xs dim', text: 'Exports include the computed figures whether or not an AI report exists.' }),
    ]));

    const order = [MASTER, ...p.sites.map((s) => s.code)].filter((k) => st.reports[k]);
    for (const k of order) body.appendChild(reportCard(st.reports[k], p));
  }

  return section('5', 'Reports', body, keys.length ? `${keys.length} generated` : '');
}

/**
 * The report, in the eight fixed sections.
 *
 * Each section pairs COMPUTED FIGURES, rendered by the app, with the model's
 * INTERPRETATION of them, visually separated so the reader always knows which
 * is which. The model is told not to restate figures, so the two do not
 * duplicate each other.
 */
/**
 * The report, rendered in the fourteen sections of the standard template.
 *
 * Each section carries a provenance badge — ENGINE, AI, or AI + ANALYST — so a
 * reader always knows whether a statement was calculated from their workbook or
 * judged by a language model. That distinction is the whole basis on which this
 * report can be trusted, so it is shown on every section rather than explained
 * once and forgotten.
 */
function reportCard(rep, p) {
  const r = rep.result || {};
  const c = rep.computed || null;
  const isMaster = rep.kind === 'master';
  const sections = isMaster ? MASTER_SECTIONS : SITE_SECTIONS;
  const ckey = `report:${rep.key}`;
  const wrap = el('div', { class: `sheet report${collapsed.has(ckey) ? ' is-collapsed' : ''}` });

  const toggle = () => {
    if (collapsed.has(ckey)) collapsed.delete(ckey); else collapsed.add(ckey);
    wrap.classList.toggle('is-collapsed', collapsed.has(ckey));
  };

  const head = el('button', { class: 'report__head' }, [
    el('span', { class: 'collapse-btn', 'aria-hidden': 'true' }, ['▼']),
    el('span', { class: 'grow row', style: { gap: '8px' } }, [
      icon('insights', 15),
      el('strong', { text: isMaster ? 'Master Analysis Report' : `Site Status Report — ${rep.key}` }),
      c?.overall ? ragChip(c.overall) : null,
    ]),
    el('span', { class: 'xs num', text: `${rep.providerLabel} · ${T.formatTokens(rep.tokens.input + rep.tokens.output)} tok${rep.tokens.estimated ? ' est' : ''}${rep.fellBack ? ` · fallback ${rep.model}` : ''}` }),
  ]);
  head.addEventListener('click', toggle);

  const body = el('div', { class: 'report__body' });

  if (!c) {
    mount(body, el('div', { class: 'rsec__body' }, [
      el('p', { text: 'This report was generated before the report structure was updated, so its computed half is missing. Re-run it to get the full fourteen-section report.' }),
    ]));
    mount(wrap, head, body);
    return wrap;
  }

  /* ---- document control + provenance key ---- */
  body.appendChild(el('div', { class: 'rsec' }, [
    el('div', { class: 'rsec__body' }, [
      el('div', { class: 'figs' }, Object.entries({
        'Report ID': c.docControl.reportId,
        'Report type': c.docControl.reportType,
        ...(isMaster
          ? { Sites: c.docControl.sites.join(', ') }
          : { Site: c.docControl.siteCode, Priority: c.docControl.priority, 'Status date': fmtDate(c.docControl.statusDate) }),
        'Report date': fmtDate(c.docControl.reportDate),
      }).map(([k, v]) => el('div', { class: 'fig' }, [
        el('div', { class: 'fig__k', text: k }),
        el('div', { class: 'fig__v', style: { fontSize: 'var(--t-sm)' }, text: String(v ?? '—') }),
      ]))),
      el('p', { class: 'xs muted', style: { marginTop: '8px' }, text: PROVENANCE_NOTE }),
    ]),
  ]));

  const sec = (id, ...kids) => {
    const meta = sections.find((x) => x.id === id);
    if (!meta) return;
    body.appendChild(el('div', { class: 'rsec' }, [
      el('div', { class: 'rsec__head' }, [
        el('span', { class: 'rsec__n', text: String(meta.n) }),
        el('h4', { text: meta.title }),
        el('span', { class: 'grow' }),
        provChip(meta.by),
      ]),
      el('div', { class: 'rsec__body' }, kids.flat().filter(Boolean)),
    ]));
  };

  const ai = (text) => (text ? el('div', { class: 'aiprose' }, [el('p', { text })]) : null);
  const list = (title, arr, fmt) => (Array.isArray(arr) && arr.length
    ? el('div', {}, [
      title ? el('h4', { style: { margin: '12px 0 4px' }, text: title }) : null,
      el('ul', {}, arr.map((x) => el('li', { text: typeof x === 'string' ? x : fmt(x) }))),
    ])
    : null);

  /* ================= 1. Introduction ================= */
  sec('introduction',
    el('p', { text: isMaster
      ? `This report consolidates ${c.docControl.sites.length} selected sites into one portfolio view: how the group is performing, where problems repeat, where sites compete for the same people, and which interventions matter most.`
      : `This report gives the Project Manager a decision-ready view of BIM delivery for site ${c.code}: where it stands against its baseline, what is constraining it, what could go wrong, and what is needed next.` }),
    el('h4', { text: 'Basis' }),
    el('p', { class: 'small', text: c.docControl.basis }),
    el('h4', { text: 'Limitations' }),
    el('p', { class: 'small muted', text: isMaster
      ? 'Portfolio figures describe the selected sites only. Resource loading counts assignments on these sites; a resource may also be committed to sites that were not selected. Totals combine sites of different size and state, so active-site figures are shown alongside.'
      : 'All figures are computed from the workbook as recorded; no field verification was carried out. Progress is measured by task status, not by physical percentage complete. Where the workbook conflicts with itself, the Sites register is taken as the baseline and the conflict is reported in the data-integrity section. Forecasts are indicative.' }));

  /* ================= 2. Executive Summary ================= */
  const ex = r.executive || {};
  sec('executive',
    ex.bottomLine ? el('p', { class: 'strong', style: { fontSize: 'var(--t-md)' }, text: `Bottom line: ${ex.bottomLine}` }) : null,
    list('Key messages', ex.keyMessages, (x) => x),
    Array.isArray(ex.decisions) && ex.decisions.length
      ? table(['ID', 'Decision', 'Owner', 'Needed by'], ex.decisions.map((d) => [d.id, d.decision, d.owner, d.byWhen || d.neededBy]))
      : el('p', { class: 'small muted', text: 'No decisions were raised for this period.' }));

  /* ================= 3. Dashboard ================= */
  if (isMaster) {
    sec('dashboard',
      figStrip([
        ['Sites', String(c.counts.sites)], ['Active', String(c.counts.active)],
        ['Submitted', String(c.counts.submitted)],
        ['Completion (count)', pctText(c.kpis.completionCount)],
        ['Completion (weighted)', pctText(c.kpis.completionWeighted)],
        ['Active sites only', pctText(c.kpis.activeCompletionWeighted)],
        ['On-time delivery', pctText(c.kpis.onTimeDelivery)],
        ['Back-end concentration', pctText(c.kpis.backEndConcentration)],
      ]),
      el('h4', { text: 'Status by site and dimension' }),
      chartHost((h) => renderRagMatrix(h, c.ragMatrix)),
      table(['Site', 'Schedule', 'Scope', 'Constraints', 'Resources', 'Quality', 'Data', 'Overall'],
        c.ragMatrix.map((x) => [x.code, ragChip(x.schedule), ragChip(x.scope), ragChip(x.constraints),
          ragChip(x.resources), ragChip(x.quality), ragChip(x.dataIntegrity), ragChip(x.overall)])));
  } else {
    sec('dashboard',
      chartHost((h) => renderKpiDashboard(h, c)),
      el('h4', { text: 'Key performance indicators' }),
      table(['Indicator', 'Actual', 'Target / threshold', 'RAG'],
        c.kpis.map((k) => [k.k, k.v ?? '—', k.target, k.rag ? ragChip(k.rag) : '—'])),
      el('h4', { text: 'Status by performance dimension' }),
      table(['Dimension', 'RAG', 'Basis'],
        Object.entries(c.dimensions).filter(([k]) => k !== 'overall')
          .map(([k, v]) => [titleCase(k.replace(/([A-Z])/g, ' $1')), ragChip(v.rag), v.basis])));
  }

  /* ================= 4. Schedule ================= */
  if (isMaster) {
    sec('schedule',
      chartHost((h) => renderScheduleAcrossSites(h, c.schedule)),
      chartHost((h) => renderCompletionVsTime(h, c.schedule)),
      table(['Site', 'Target', 'Forecast / actual', 'Variance (d)', 'Completion', 'Time used', 'Status'],
        c.schedule.map((s) => [s.code, fmtDate(s.target), fmtDate(s.forecast),
          s.variance == null ? '—' : String(s.variance), pctText(s.completionWeighted),
          pctText(s.timeElapsed), ragChip(s.overall)])),
      ai(r.notes?.schedule));
  } else {
    sec('schedule',
      el('h4', { text: 'Milestones' }),
      table(['ID', 'Milestone', 'Baseline', 'Actual / forecast', 'Variance', 'Status', 'RAG'],
        c.milestones.map((m) => [m.id, m.name, m.baseline ? fmtDate(m.baseline) : '—',
          m.actual && /^\d{4}-/.test(String(m.actual)) ? fmtDate(m.actual) : (m.actual || '—'),
          m.variance == null ? '—' : `${m.variance > 0 ? '+' : ''}${m.variance} d`,
          m.status, ragChip(m.rag)])),
      chartHost((h) => renderMilestoneTimeline(h, c)),
      el('h4', { text: 'Forecast' }),
      el('p', { class: 'small', text: forecastSentence(c) }),
      el('h4', { text: 'Week-on-week movement' }),
      c.movement.comparable
        ? el('div', {}, [
          el('p', { class: 'small', text: `${c.movement.fromWeek} → ${c.movement.toWeek}: ${c.movement.finishedThisWeek} task${c.movement.finishedThisWeek === 1 ? '' : 's'} reached Finished.` }),
          c.movement.changes.length
            ? el('ul', {}, c.movement.changes.slice(0, 20).map((x) => el('li', { class: 'small', text: `${x.id} ${x.name}: ${x.from || 'blank'} → ${x.to || 'blank'}${x.isCategory ? ' (category row)' : ''}` })))
            : el('p', { class: 'small muted', text: 'No status changed between the last two weeks.' }),
        ])
        : el('p', { class: 'small muted', text: `Not comparable: ${c.movement.reason}.` }),
      ai(r.notes?.schedule));
  }

  /* ================= 5. Work breakdown / bottlenecks ================= */
  if (isMaster) {
    sec('bottlenecks',
      chartHost((h) => renderCategoryBySite(h, c.categories, c.docControl.sites)),
      table(['Category', 'Applicable', 'Finished', 'Open', 'Sites with stuck work'],
        c.categories.filter((x) => x.applicable > 0).map((x) => [
          `${x.id} ${x.name}`, String(x.applicable), String(x.finished), String(x.open),
          x.stuckSites.join(', ') || '—'])),
      ai(r.notes?.bottlenecks));
  } else {
    sec('wbs',
      el('p', { class: 'small muted', text: `"Reported" is the coordinator's category row; "Roll-up check" compares it with the arithmetic of its tasks. A mismatch is not necessarily an error, but it should be explained. Discrepancies found: ${c.wbs.discrepancies}.` }),
      chartHost((h) => renderCategoryProgressReport(h, c.wbs.rows)),
      table(['ID', 'Category', 'Appl.', 'Fin.', 'WIP', 'Blk / WO', '% done', 'Reported', 'Roll-up check'],
        c.wbs.rows.map((x) => [x.id, x.name, String(x.liveCount), String(x.doneCount),
          String(Math.max(0, x.liveCount - x.doneCount - x.stuck)), String(x.stuck),
          x.liveCount ? `${x.computedPct}%` : 'N/A', x.selfStatus || '—',
          x.consistent ? 'Consistent' : x.check])),
      ai(r.notes?.wbs));
  }

  /* ================= 6. Scope ================= */
  sec('scope',
    isMaster
      ? table(['Site', 'Source', 'Flagged additional', 'Set N/A afterwards', 'Unconfirmed items'],
        c.scope.map((s) => [s.code, s.source, String(s.additional), String(s.setNA), String(s.openItems)]))
      : el('div', {}, [
        el('p', { class: 'small muted', text: `Source: ${c.scope.source}. ${c.scope.additionalCount} task${c.scope.additionalCount === 1 ? '' : 's'} added after kickoff (${c.scope.growthPct}% growth)${c.scope.additionalSetNA ? `, of which ${c.scope.additionalSetNA} were later set N/A` : ''}.` }),
        c.scope.items.length
          ? table(['Ref', 'Scope item', 'Type', 'Raised', 'Status', 'Comment'],
            c.scope.items.map((s) => [s.ref, s.item, s.type, s.raised ? fmtDate(s.raised) : '—', s.status, s.comment || '—']))
          : null,
        c.scope.additional.length
          ? table(['Task', 'Name', 'Date added', 'Status'],
            c.scope.additional.map((x) => [x.taskId, x.name, x.added ? fmtDate(x.added) : '—', x.status || '—']))
          : null,
      ]),
    ai(r.notes?.scope));

  /* ================= 7. Constraints ================= */
  sec('constraints',
    isMaster
      ? el('div', {}, [
        chartHost((h) => renderPartyDependency(h, c.parties)),
        table(['Party', 'Open items', 'Sites', 'Items'],
          c.parties.map((x) => [x.party, String(x.count), x.sites.join(', '), x.items.slice(0, 6).join(', ')])),
      ])
      : el('div', {}, [
        el('p', { class: 'small muted', text: 'Make-ready view: each prerequisite must be in hand before the work that depends on it can start.' }),
        table(['ID', 'Prerequisite', 'Template', 'Needed for', 'Provider', 'Req. by', 'Status', 'RAG'],
          c.constraints.rows.map((x) => [x.id, x.name, x.fromTemplate, (x.neededFor || []).join(', ') || '—',
            x.provider, x.byWeek == null ? '—' : `W${x.byWeek}`, x.status, ragChip(x.rag)])),
        c.constraints.noReceivedDates
          ? el('p', { class: 'small', style: { color: 'var(--survey)' }, text: 'No received dates are recorded for any prerequisite, so constraint ageing cannot be measured.' })
          : null,
      ]),
    ai(r.notes?.constraints));

  /* ================= 8. Log (site only) ================= */
  if (!isMaster) {
    sec('log',
      c.log.rows.length
        ? el('div', {}, [
          chartHost((h) => renderConstraintAgeing(h, c.log.open)),
          table(['Log', 'Week', 'Task', 'Type', 'Reason', 'Waiting on', 'Raised', 'Age (d)', 'Control', 'RAG'],
            c.log.rows.map((l) => [l.id, l.week == null ? '—' : `W${l.week}`, l.taskId || '—', l.kind,
              l.reason, l.waitingOn || '—', l.raised ? fmtDate(l.raised) : '—',
              l.age == null ? '—' : String(l.age), l.control, ragChip(l.rag)])),
          el('p', { class: 'xs muted', text: 'Age measured in calendar days at the report date. Control: Internal = Blocked (the team can fix it); External = Waiting on (must be chased).' }),
        ])
        : el('p', { class: 'small muted', text: 'No entries recorded.' }),
      ai(r.notes?.log));
  }

  /* ================= 9. Risk register ================= */
  const risks = r.risks || [];
  sec('risks',
    risks.length
      ? el('div', {}, [
        chartHost((h) => renderRiskHeatMap(h, risks)),
        table(isMaster
          ? ['ID', 'Risk (cause → effect)', 'Sites', 'P', 'I', 'Score', 'Rating', 'Strategy', 'Response', 'Owner']
          : ['ID', 'Risk (cause → effect)', 'P', 'I', 'Score', 'Rating', 'Strategy', 'Response', 'Owner'],
        risks.map((x) => (isMaster
          ? [x.id, x.risk, (x.sitesAffected || []).join(', '), String(x.probability), String(x.impact), String(x.score), ragChip(x.rating), x.strategy, x.response, x.owner]
          : [x.id, x.risk, String(x.probability), String(x.impact), String(x.score), ragChip(x.rating), x.strategy, x.response, x.owner]))),
        el('p', { class: 'xs muted', text: 'P and I are the model\'s judgement, each 1–5. Score = P × I and the rating band are computed by this app: RED ≥ 15, AMBER 8–14, GREEN ≤ 7. Review and adjust before issuing.' }),
      ])
      : el('p', { class: 'small muted', text: 'No risks were proposed.' }));

  /* ================= 10. Resources ================= */
  sec('resources',
    isMaster
      ? el('div', {}, [
        chartHost((h) => renderResourceAcrossSites(h, c.resourceLoad)),
        table(['R code', 'Selected sites', 'Active sites', 'Open tasks', 'Load'],
          c.resourceLoad.map((x) => [x.resource, x.sites.join(', '), String(x.activeSites), String(x.open), x.load])),
        el('p', { class: 'xs muted', text: 'Loading reflects the selected sites only; a resource may also be committed to sites that were not selected.' }),
      ])
      : el('div', {}, [
        table(['R code', 'Open', 'In progress', 'Stuck'],
          c.resources.rows.map((x) => [x.resource, String(x.open), String(x.wip), String(x.stuck)])),
        c.resources.unassigned?.open
          ? el('p', { class: 'small', style: { color: 'var(--hivis)' }, text: `${c.resources.unassigned.open} open task${c.resources.unassigned.open === 1 ? ' has' : 's have'} no responsible resource.` })
          : null,
        c.resources.externalParties.length
          ? el('div', {}, [
            el('h4', { text: 'External dependencies (this site only)' }),
            table(['Party', 'Open items', 'Items'],
              c.resources.externalParties.map((x) => [x.party, String(x.items.length), x.items.join(', ')])),
          ])
          : null,
      ]),
    ai(r.notes?.resources));

  /* ================= 11. Quality and data integrity ================= */
  sec('quality',
    isMaster
      ? el('div', {}, [
        el('p', { class: 'small muted', text: 'A finding that appears on several sites is treated as a process issue to fix once, centrally; a finding on one site stays with that site\'s report.' }),
        chartHost((h) => renderFindingsBySite(h, c.patterns, c.docControl.sites)),
        table(['Pattern', 'Sites', 'Count', 'Systemic', 'Correction'],
          c.patterns.map((x) => [x.pattern, x.sites.join(', '), String(x.count), x.systemic ? 'Yes' : 'No', x.correction])),
      ])
      : (c.quality.findings.length
        ? table(['ID', 'Finding', 'Evidence', 'Impact', 'Correction'],
          c.quality.findings.map((f) => [f.id, f.finding, f.evidence, f.impact, f.correction]))
        : el('p', { class: 'small', text: 'No findings.' })),
    ai(r.notes?.quality));

  /* ================= 11/12. Prioritisation (master) ================= */
  if (isMaster) {
    sec('priority',
      c.ipi.length
        ? el('div', {}, [
          chartHost((h) => renderIpi(h, c.ipi)),
          table(['Rank', 'Site', 'IPI', 'RAG', 'Priority', 'Days to target', 'SPI proxy'],
            c.ipi.map((x, i) => [String(i + 1), x.code, String(x.ipi), ragChip(x.overall), x.priority,
              x.daysToTarget == null ? '—' : String(x.daysToTarget), x.spiProxy == null ? '—' : String(x.spiProxy)])),
          el('p', { class: 'xs muted', text: 'IPI = overall RAG (3/2/1) + site priority (2/1/0) + time pressure (3/2/1/0) + 3 × (1 − SPI proxy). Active sites only; ties broken by the earlier target date.' }),
          list('Recommended interventions', r.interventions, (x) => `${x.site}: ${x.recommendedIntervention}`),
        ])
        : el('p', { class: 'small muted', text: 'No active sites: prioritisation does not apply. See the closeout summary.' }),
      ai(r.notes?.priority));
  }

  /* ================= 12. Actions and lookahead ================= */
  sec('actions',
    Array.isArray(r.actions) && r.actions.length
      ? table(['ID', 'Action', 'Owner', 'Due', 'Priority', 'Links'],
        r.actions.map((x) => [x.id, x.action, x.owner, x.due, x.priority, x.links || '—']))
      : el('p', { class: 'small muted', text: 'No actions were raised.' }),
    Array.isArray(r.lookahead) && r.lookahead.length
      ? el('div', {}, [
        el('h4', { text: 'Two-week lookahead' }),
        table(['Week', 'Planned focus'], r.lookahead.map((x) => [x.week, x.focus])),
      ])
      : null);

  /* ================= 13. Conclusion ================= */
  sec('conclusion', ai(r.conclusion) || el('p', { class: 'small muted', text: 'No conclusion was returned.' }));

  /* ================= 14. References ================= */
  sec('references',
    el('ol', { class: 'small' }, c.references.map(([ref, used]) => el('li', {}, [
      el('span', { text: ref }),
      el('span', { class: 'xs muted', text: ` — Used for: ${used}` }),
    ]))));

  /* ================= Appendices ================= */
  body.appendChild(el('div', { class: 'rsec' }, [
    el('div', { class: 'rsec__head' }, [
      el('span', { class: 'rsec__n', text: 'A' }),
      el('h4', { text: 'Appendix A. Metric definitions and RAG thresholds' }),
      el('span', { class: 'grow' }), provChip(ENGINE),
    ]),
    el('div', { class: 'rsec__body' }, [
      table(['Metric', 'Definition / formula'], METRIC_DEFINITIONS),
      el('h4', { text: 'RAG thresholds' }),
      table(['Dimension', 'GREEN', 'AMBER', 'RED'], RAG_THRESHOLDS),
    ]),
  ]));

  body.appendChild(el('div', { class: 'rsec' }, [
    el('div', { class: 'rsec__head' }, [
      el('span', { class: 'rsec__n', text: 'B' }),
      el('h4', { text: isMaster ? 'Appendix B. Site summary register' : 'Appendix B. Task-level status register' }),
      el('span', { class: 'grow' }), provChip(ENGINE),
    ]),
    el('div', { class: 'rsec__body' }, [
      isMaster
        ? table(['Site', 'Priority', 'Resources', 'Overall', 'Site report ID'],
          c.appendixB.map((x) => [x.code, x.priority, x.resources.join(', ') || '—', ragChip(x.overall), x.reportId]))
        : table(['Task ID', 'Task', 'Disc.', 'Wt', 'Type', ...(c.weeks || []).map((w) => w.label)],
          c.appendixB.map((t) => [t.id, t.name, t.discipline, String(t.weight), t.type,
            ...(t.weekly || []).map((s) => s || '—')])),
    ]),
  ]));

  /* ---- footer ---- */
  body.appendChild(el('div', { class: 'rsec' }, [
    el('div', { class: 'rsec__body row row--wrap' }, [
      el('button', { class: 'btn btn--sm', disabled: running ? true : null, onclick: () => continueReport(rep, p) }, ['Continue — ask a follow-up']),
      el('button', {
        class: 'btn btn--sm', disabled: running ? true : null,
        onclick: () => { S.update((x) => { const n = { ...x.reports }; delete n[rep.key]; x.reports = n; }); toast('Report removed.'); },
      }, [icon('trash', 13), 'Remove']),
      el('span', { class: 'grow' }),
      el('span', { class: 'xs dim', text: `${rep.model} · ${fmtDate((rep.at || '').slice(0, 10))}${rep.followUp ? ` · asked: ${rep.followUp}` : ''}` }),
    ]),
  ]));

  mount(wrap, head, body);
  return wrap;
}

/* ---------------- small builders ---------------- */

function ragChip(rag) {
  const tone = rag === 'RED' ? 'survey' : (rag === 'AMBER' ? 'hivis' : 'sign');
  return el('span', { class: 'chip', dataset: { tone }, text: rag });
}

function provChip(by) {
  const map = {
    [ENGINE]: ['blueprint', 'ENGINE', 'Calculated by this app from your workbook'],
    [AI]: ['plum', 'AI', 'Written by the language model from the calculated figures'],
    [AI_ANALYST]: ['hivis', 'AI + ANALYST', 'Drafted by the model — check and approve before issuing'],
  };
  const [tone, label, title] = map[by] || map[ENGINE];
  return el('span', { class: 'chip', dataset: { tone }, title, text: label });
}

function pctText(v) { return v == null ? '—' : `${v}%`; }

function figStrip(pairs) {
  return el('div', { class: 'figs' }, pairs.map(([k, v]) => el('div', { class: 'fig' }, [
    el('div', { class: 'fig__k', text: k }),
    el('div', { class: 'fig__v', text: String(v) }),
  ])));
}

/** Table builder. Cells may be strings or already-built nodes (e.g. RAG chips). */
function table(headers, rows) {
  return el('div', { class: 'tablewrap' }, [
    el('table', { class: 'data' }, [
      el('thead', {}, [el('tr', {}, headers.map((h) => el('th', { text: h })))]),
      el('tbody', {}, (rows || []).map((row) => el('tr', {},
        row.map((cell) => el('td', {}, [cell instanceof Node ? cell : el('span', { text: cell == null ? '—' : String(cell) })]))))),
    ]),
  ]);
}

/** Defers chart drawing until the node is in the document. */
function chartHost(draw) {
  const h = el('div');
  queueMicrotask(() => { try { draw(h); } catch (e) { console.error('chart failed', e); } });
  return h;
}

function forecastSentence(c) {
  const m = c.metrics;
  if (c.submitted) return `Submitted ${m.forecastDate ? fmtDate(m.forecastDate) : ''}${m.scheduleVariance != null ? `, ${m.scheduleVariance > 0 ? `${m.scheduleVariance} days after` : `${Math.abs(m.scheduleVariance)} days before`} the target.` : '.'}`;
  if (m.forecastMethod === 'throughput') {
    return `Throughput-based forecast: at the rate measured over the last weeks the earliest submission is ${fmtDate(m.forecastDate)}${m.scheduleVariance != null ? `, ${m.scheduleVariance > 0 ? `${m.scheduleVariance} days past` : `${Math.abs(m.scheduleVariance)} days inside`} the target.` : '.'}`;
  }
  if (m.indicativeDays != null) {
    return `Throughput-based forecasting is not possible (nothing finished in the last week). Indicative forecast using the template's typical durations: ${m.indicativeDays} working days of remaining categories, giving an earliest submission of ${fmtDate(m.forecastDate)}${m.scheduleVariance != null ? `, ${m.scheduleVariance > 0 ? `${m.scheduleVariance} days past` : `${Math.abs(m.scheduleVariance)} days inside`} the target.` : '.'}`;
  }
  return 'No forecast is possible: nothing has completed and the template records no typical durations for the remaining categories.';
}

function continueReport(rep, p) {
  const { body, foot, close } = openModal({ title: `Continue — ${rep.title}` });
  const input = el('textarea', { class: 'textarea', placeholder: 'e.g. If R5 clears the ceiling void question next week, does this site still make its date?' });
  mount(body,
    el('p', { class: 'small muted', text: 'This runs a fresh analysis of the same computed data with your question added. It replaces the report above; the previous one is not kept.' }),
    el('div', { class: 'field' }, [el('label', { text: 'Your question' }), input]),
  );
  mount(foot,
    el('button', { class: 'btn', onclick: close }, ['Cancel']),
    el('button', {
      class: 'btn btn--primary',
      onclick: async () => {
        const q = input.value.trim();
        if (!q) { toast('Type a question first.', 'hivis'); return; }
        close();
        running = true;
        controller = new AbortController();
        currentKey = rep.key;
        render();
        try {
          const payload = withPeriod(
            rep.key === MASTER
              ? buildMasterPayload(masterReportFor(Object.keys(S.get().reports).filter((k) => k !== MASTER)))
              : buildSitePayload(siteReportFor(rep.key)),
            rep.key,
          );
          const next = await runAI({
            kind: rep.kind, payload, settings: S.get().settings,
            followUp: q, signal: controller.signal, onStatus: showRetry,
          });
          next.periodContext = payload.previousPeriod || null;
          next.computed = rep.kind === 'master'
            ? masterReportFor(Object.keys(S.get().reports).filter((k) => k !== MASTER))
            : siteReportFor(rep.key);
          if (next.result?.risks) next.result.risks = scoreRisks(next.result.risks);
          S.update((s) => { s.reports = { ...s.reports, [rep.key]: next }; });
          toast('Updated.', 'sign');
        } catch (e) {
          toast(e instanceof CancelledError ? 'Cancelled.' : e.message, e instanceof CancelledError ? 'hivis' : 'survey');
        } finally {
          running = false; currentKey = null; controller = null; showRetry(null); render();
        }
      },
    }, ['Ask']),
  );
}

/* ========================================================================
   Payload review
   ======================================================================== */

function showPayload(key, asConfirm = false, total = 1) {
  return new Promise((resolve) => {
    const p = portfolio();
    const st = S.get();
    const payload = withPeriod(
      key === MASTER
        ? buildMasterPayload(masterReportFor(st.selection.length ? st.selection : p.sites.map((x) => x.code)))
        : buildSitePayload(siteReportFor(key)),
      key,
    );
    const pv = previewPayload({
      kind: key === MASTER ? 'master' : 'site',
      payload, settings: st.settings,
    });

    const { body, foot, close } = openModal({
      title: asConfirm ? 'Confirm before sending' : 'What would be sent',
      wide: true,
      onClose: () => resolve(false),
    });

    mount(body,
      el('div', { class: 'notice', dataset: { tone: 'hivis' } }, [
        el('h4', { text: `This leaves your browser and goes to ${pv.destination}` }),
        el('ul', { class: 'small', style: { margin: '6px 0 0', paddingLeft: '18px' } }, [
          el('li', { text: `Model: ${pv.model}` }),
          el('li', { text: `${total} call${total === 1 ? '' : 's'} in this run. Shown below: ${key === MASTER ? 'the master analysis' : `site ${key}`}${total > 1 ? ' (the others follow the same shape)' : ''}.` }),
          el('li', { text: `About ${fmtNum(Math.round(pv.bytes / 1024))} KB, roughly ${T.formatTokens(pv.estimatedInputTokens)} input tokens (estimated).` }),
          el('li', { text: 'Site and resource codes only — no client, site or personal names, unless you put them in a task name.' }),
          el('li', { text: 'Every number below was computed by this app. The model is instructed not to recalculate them.' }),
        ]),
      ]),
      el('h4', { text: 'Computed data', style: { marginTop: '14px' } }),
      el('pre', { class: 'payload-preview', text: JSON.stringify(pv.payload, null, 2) }),
    );
    mount(foot,
      el('button', { class: 'btn', onclick: () => { close(); resolve(false); } }, [asConfirm ? 'Cancel' : 'Close']),
      asConfirm ? el('button', { class: 'btn btn--primary', onclick: () => { close(); resolve(true); } }, ['Send it']) : null,
    );
  });
}

/* ========================================================================
   Right rail — deterministic dashboard + token meter
   ======================================================================== */

/**
 * The fixed rail. Usage against your limits and where every site stands are the
 * two things worth glancing at constantly, so they never scroll away with the
 * work column.
 */
function renderRail() {
  const st = S.get();
  const p = portfolio();
  const nodes = [usageCard(st)];

  if (p && st.confirmed) {
    nodes.push(el('div', { class: 'railcard' }, [
      el('div', { class: 'railcard__head' }, [
        el('h3', { text: 'Site progress' }),
        el('span', { class: 'tag', text: 'computed' }),
      ]),
      ...p.sites.map((s) => el('div', { class: 'railsite' }, [
        el('span', { class: 'code', text: s.code }),
        el('div', { class: 'track' }, [
          el('div', { style: { width: `${s.noData ? 0 : clamp(s.pctByWeight, 0, 100)}%`, background: s.slipWeeks > 0 ? 'var(--survey)' : 'var(--sign)' } }),
        ]),
        el('span', {
          class: `val${s.slipWeeks > 0 ? ' late' : ''}`,
          text: s.noData ? '—' : `${Math.round(s.pctByWeight)}%`,
        }),
      ])),
      el('div', { class: 'railnote', style: { marginTop: '8px' } }, [
        `${p.totalDone} of ${p.totalTasks} tasks complete across ${p.sitesWithData} measured site${p.sitesWithData === 1 ? '' : 's'}.`,
      ]),
    ]));

    if (p.risks.length) {
      nodes.push(el('div', { class: 'railcard' }, [
        el('div', { class: 'railcard__head' }, [
          el('h3', { text: 'Computed risks' }),
          el('span', { class: 'tag', text: String(p.risks.length) }),
        ]),
        ...p.risks.slice(0, 5).map((r) => el('div', { style: { marginBottom: '9px' } }, [
          el('div', {
            style: {
              fontSize: 'var(--t-xs)', fontWeight: '600',
              color: r.level === 'high' ? '#FF9C9C' : '#FFD27A', lineHeight: '1.35',
            },
            text: r.title,
          }),
          el('div', { class: 'railnote', text: r.detail.length > 130 ? `${r.detail.slice(0, 128)}…` : r.detail }),
        ])),
      ]));
    }
  }

  mount(rail, ...nodes);
}

function usageCard(st) {
  const usage = T.getUsage();
  const win = T.requestWindows(st.settings.provider);
  const lim = T.getLimits(st.settings.provider);

  const bar = (used, max) => {
    const pc = max ? clamp((used / max) * 100, 0, 100) : 0;
    return el('div', {
      class: 'railbar',
      dataset: { tone: pc > 85 ? 'survey' : (pc > 60 ? 'hivis' : 'blueprint') },
    }, [el('div', { style: { width: `${pc}%` } })]);
  };

  return el('div', { class: 'railcard' }, [
    el('div', { class: 'railcard__head' }, [
      el('h3', { text: 'Usage' }),
      el('span', { class: 'tag', text: PROVIDERS[st.settings.provider]?.label || st.settings.provider }),
    ]),
    el('div', { class: 'railrow' }, [el('span', { text: 'Requests this minute' }), el('strong', { text: `${win.lastMinute} / ${lim.rpm ?? '—'}` })]),
    lim.rpm ? bar(win.lastMinute, lim.rpm) : null,
    win.lastMinute && lim.rpm && win.lastMinute >= lim.rpm
      ? el('div', { class: 'railnote', style: { color: '#FF9C9C', marginBottom: '8px' }, text: `Capacity returns in about ${win.nextMinuteSlotIn}s` })
      : null,
    el('div', { class: 'railrow' }, [el('span', { text: 'Requests today' }), el('strong', { text: `${win.lastDay} / ${lim.rpd ?? '—'}` })]),
    lim.rpd ? bar(win.lastDay, lim.rpd) : null,
    el('div', { class: 'railrow' }, [el('span', { text: 'Tokens used' }), el('strong', { text: `${T.formatTokens(usage.total)}${usage.estimated ? ' est' : ''}` })]),
    el('div', { class: 'railrow' }, [el('span', { text: 'In / out' }), el('strong', { text: `${T.formatTokens(usage.input)} / ${T.formatTokens(usage.output)}` })]),
    el('div', { class: 'railrow' }, [el('span', { text: 'Calls made' }), el('strong', { text: String(usage.calls) })]),
    ...overloadedModels().map((o) => el('div', { class: 'railnote', style: { marginTop: '6px', color: '#FFB86B' } }, [
      `${o.model} overloaded \u2014 skipped for ${o.minutesLeft} more min. `,
      el('a', {
        href: '#', style: { color: '#fff' },
        onclick: (e) => { e.preventDefault(); clearOverloadMemory(o.model); renderRail(); toast(`${o.model} will be tried again on the next run.`); },
      }, ['Try it again']),
    ])),
    (() => {
      const ks = S.keyStatus();
      return el('div', { class: 'railrow', style: { marginTop: '6px' } }, [
        el('span', { text: 'API key' }),
        el('strong', {
          style: { color: ks.set ? '#7EE2BD' : '#FFB86B' },
          text: ks.set ? (ks.where === 'tab' ? 'saved \u00b7 this tab' : 'saved \u00b7 until refresh') : 'not set',
        }),
      ]);
    })(),
    el('div', { class: 'railnote', style: { marginTop: '8px' }, text: `${lim.label}${lim.isDefault ? ' — default figures, edit in Settings to match your account' : ''}` }),
    el('div', { class: 'row', style: { gap: '6px', marginTop: '10px' } }, [
      el('button', { class: 'btn btn--sm', onclick: () => { T.resetUsage(); render(); } }, ['Reset tokens']),
      el('button', { class: 'btn btn--sm', onclick: () => { T.resetRequests(st.settings.provider); render(); } }, ['Reset requests']),
    ]),
  ]);
}

/* ========================================================================
   Settings, session, helpers
   ======================================================================== */

function openSettings() {
  const st = S.get();
  const { body, foot, close } = openModal({ title: 'Settings', wide: true });

  const prov = el('select', { class: 'select' },
    Object.entries(PROVIDERS).map(([k, v]) => el('option', { value: k, selected: k === st.settings.provider ? true : null, text: v.label })));
  const model = el('input', { class: 'input', value: st.settings.model || '', placeholder: PROVIDERS[st.settings.provider].defaultModel });
  const key = el('input', { class: 'input', type: 'password', autocomplete: 'off', spellcheck: 'false',
    placeholder: S.hasApiKey() ? 'Paste a new key only to replace the saved one' : PROVIDERS[st.settings.provider].keyHint });

  // The box is deliberately never filled with the real key. This block is what
  // says a key IS saved — the grey placeholder alone looked like an empty field.
  const ks = S.keyStatus();
  const keyState = el('div', {
    class: 'notice', dataset: { tone: ks.set ? 'sign' : 'hivis' },
    style: { padding: '8px 12px', marginBottom: '6px' },
  }, [
    el('div', { class: 'row', style: { gap: '8px', flexWrap: 'wrap' } }, [
      el('strong', { class: 'small', text: ks.set ? `\u2713 Key saved: ${ks.masked}` : 'No key saved' }),
      el('span', { class: 'xs muted', text: ks.lasts }),
    ]),
  ]);
  const remember = el('input', { type: 'checkbox', checked: st.settings.rememberKeyForSession ? true : null });
  const review = el('input', { type: 'checkbox', checked: st.settings.reviewPayload ? true : null });
  const maxTok = el('input', { class: 'input', type: 'number', value: st.settings.maxTokens, min: 1024, max: 16000 });
  const fallback = el('input', {
    class: 'input', value: st.settings.fallbackModel ?? '',
    placeholder: `e.g. ${PROVIDERS[st.settings.provider].defaultModel}`,
  });

  const lim = T.getLimits(st.settings.provider);
  const tier = el('select', { class: 'select' }, [
    el('option', { value: 'free', selected: lim.tier === 'free' ? true : null, text: 'Free / evaluation' }),
    el('option', { value: 'paid', selected: lim.tier === 'paid' ? true : null, text: 'Paid' }),
  ]);
  const rpm = el('input', { class: 'input', type: 'number', value: lim.rpm ?? '' });
  const rpd = el('input', { class: 'input', type: 'number', value: lim.rpd ?? '' });
  const tpm = el('input', { class: 'input', type: 'number', value: lim.tpm ?? '' });

  prov.addEventListener('change', () => {
    model.placeholder = PROVIDERS[prov.value].defaultModel;
    key.placeholder = S.hasApiKey() ? 'Paste a new key only to replace the saved one' : PROVIDERS[prov.value].keyHint;
  });

  mount(body,
    el('div', { class: 'notice', dataset: { tone: 'hivis' } }, [
      el('h4', { text: 'Before you paste a key' }),
      el('p', { class: 'small', text: 'This app has no server. Your key stays in this browser tab and goes only to the provider you pick. It is never written into a session file or an export. But anything with access to this tab can read it, so use a key created for this tool, restricted to one model, with a spend cap.' }),
    ]),
    el('div', { class: 'formgrid', style: { marginTop: '12px' } }, [
      el('div', { class: 'field' }, [el('label', { text: 'Provider' }), prov]),
      el('div', { class: 'field' }, [
        el('label', { text: 'Model' }), model,
        el('span', { class: 'hint', text: `Blank uses ${PROVIDERS[st.settings.provider].defaultModel}. Model names change — a 404 means this string is stale.` }),
      ]),
      el('div', { class: 'field span-2' }, [
        el('label', { text: S.hasApiKey() ? 'API key (saved — leave blank to keep it)' : 'API key' }),
        keyState, key,
        el('span', { class: 'hint' }, ['Get one from ', el('a', { href: PROVIDERS[st.settings.provider].keyUrl, target: '_blank', rel: 'noopener noreferrer', text: 'the provider console' }), '. The field clears after saving.']),
      ]),
      el('div', { class: 'field' }, [
        el('label', { text: 'Fallback model if overloaded' }), fallback,
        el('span', { class: 'hint', text: 'Used only when your model is still returning 503 after all retries. Pick an established model — the newest ones are the most crowded. Leave blank to never switch.' }),
      ]),
      el('div', { class: 'field' }, [
        el('label', { text: 'Max output tokens per call' }), maxTok,
        el('span', { class: 'hint', text: 'Lower is cheaper and less likely to be cut short.' }),
      ]),
    ]),
    el('div', { class: 'stack', style: { marginTop: '10px', gap: '8px' } }, [
      el('label', { class: 'check' }, [remember, el('span', {}, [
        el('strong', { text: 'Keep the key if I refresh the page. ' }),
        el('span', { class: 'muted', text: 'Without this, a refresh wipes the key and you must paste it again. With it, the key survives refreshes but still disappears when you close the tab. Leave off on a shared computer.' }),
      ])]),
      el('label', { class: 'check' }, [review, el('span', {}, [
        el('strong', { text: 'Show me the payload before every run. ' }),
        el('span', { class: 'muted', text: 'Recommended.' }),
      ])]),
    ]),
    el('h4', { text: 'Your account limits', style: { marginTop: '16px' } }),
    el('p', { class: 'small muted', text: lim.caution || 'These figures drive the capacity warning before a run. They are defaults, not read from your account — edit them to match what your provider actually gives you.' }),
    el('div', { class: 'formgrid' }, [
      el('div', { class: 'field' }, [el('label', { text: 'Tier' }), tier]),
      el('div', { class: 'field' }, [el('label', { text: 'Requests / minute' }), rpm]),
      el('div', { class: 'field' }, [el('label', { text: 'Requests / day' }), rpd]),
      el('div', { class: 'field' }, [el('label', { text: 'Tokens / minute' }), tpm]),
    ]),
  );

  mount(foot,
    S.hasApiKey() ? el('button', {
      class: 'btn', style: { marginRight: 'auto' },
      onclick: () => { S.clearApiKey(); toast('Key cleared from this browser.'); close(); },
    }, ['Clear key']) : null,
    el('button', { class: 'btn', onclick: close }, ['Cancel']),
    el('button', {
      class: 'btn btn--primary',
      onclick: () => {
        S.update((s) => {
          s.settings = {
            ...s.settings,
            provider: prov.value,
            model: model.value.trim(),
            rememberKeyForSession: remember.checked,
            reviewPayload: review.checked,
            maxTokens: clamp(Number(maxTok.value) || 4096, 512, 16000),
            fallbackModel: fallback.value.trim(),
          };
        });
        if (key.value.trim()) S.setApiKey(key.value.trim(), remember.checked);
        else if (S.hasApiKey()) S.setApiKey(S.getApiKey(), remember.checked);
        T.saveLimits(prov.value, { tier: tier.value, rpm: rpm.value, rpd: rpd.value, tpm: tpm.value });
        close();
        toast('Settings saved.', 'sign');
        render();
      },
    }, ['Save']),
  );
}

function onSaveSession() {
  const name = S.saveSessionToFile();
  toast(`Saved ${name}. Your API key is deliberately not in it.`, 'sign', 6000);
}

async function onLoadSession() {
  const st = S.get();
  if (st.model || Object.keys(st.reports).length) {
    const ok = await confirmDialog({
      title: 'Replace the current session?',
      message: 'Loading a session file replaces everything currently open, including any reports you have generated. Save the current one first if you need it.',
      confirmLabel: 'Load file',
    });
    if (!ok) return;
  }
  try {
    const res = await S.loadSessionFromFile();
    if (res.cancelled) return;
    cachedPortfolio = null;
    render();
    toast(`Loaded session from ${res.filename}. Re-enter your API key in Settings to generate anything new.`, 'sign', 8000);
  } catch (e) {
    toast(e.message, 'survey', 9000);
  }
}

/* ------------------------------- helpers ------------------------------- */

function section(n, title, bodyEl, meta) {
  const key = `step:${n}`;
  const wrap = el('div', { class: `sheet step${collapsed.has(key) ? ' is-collapsed' : ''}` });
  const btn = el('button', {
    class: 'collapse-btn', title: 'Collapse or expand', 'aria-label': `Collapse or expand ${title}`,
    onclick: (e) => {
      e.stopPropagation();
      if (collapsed.has(key)) collapsed.delete(key); else collapsed.add(key);
      wrap.classList.toggle('is-collapsed', collapsed.has(key));
    },
  }, ['▼']);
  mount(wrap,
    el('div', { class: 'sheet__head' }, [
      btn,
      el('span', { class: 'step__n', text: n }),
      el('h3', { text: title }),
      el('span', { class: 'grow' }),
      meta ? el('span', { class: 'xs dim', text: meta }) : null,
    ]),
    bodyEl,
  );
  return wrap;
}

function band(k, v, note, tone) {
  return el('div', { class: 'band__cell', dataset: { tone } }, [
    el('div', { class: 'band__k', text: k }),
    el('div', { class: 'band__v', text: v }),
    note ? el('div', { class: 'band__note', text: note }) : null,
  ]);
}

function stat(k, v) {
  return el('span', { class: 'row', style: { gap: '5px' } }, [
    el('span', { class: 'xs dim', text: k }),
    el('span', { class: 'small strong num', text: v }),
  ]);
}

function notice(tone, title, items = [], footer = '') {
  return el('div', { class: 'notice', dataset: { tone } }, [
    el('h4', { text: title }),
    items.length ? el('ul', { class: 'small', style: { margin: '4px 0 0', paddingLeft: '18px' } },
      items.map((t) => el('li', { text: t }))) : null,
    footer ? el('p', { class: 'small', style: { margin: '6px 0 0' }, text: footer }) : null,
  ]);
}
