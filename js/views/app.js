import {
  $, el, mount, icon, toast, openModal, confirmDialog, pickFile, fmtDate,
  fmtNum, clamp, debounce,
} from '../core/util.js';
import { readWorkbook, parseWorkbook, screenText, xlsxAvailable, STATUS } from '../core/parser.js';
import { analysePortfolio } from '../core/engine.js';
import * as S from '../core/session.js';
import * as T from '../core/tokens.js';
import { PROVIDERS, run as runAI, buildSitePayload, buildMasterPayload, previewPayload, CancelledError } from '../core/ai.js';
import { exportExcel, exportWord, exportPDF } from '../core/exports.js';
import {
  renderCurve, renderThroughput, renderSiteBars, renderStatusGrid, renderStatusMix,
  renderCategoryProgress, renderBlockerAges, renderResourceLoad, meter, statusLegend,
} from '../core/charts.js';

const MASTER = S.MASTER_KEY;

/* Live run state, deliberately outside the session — a cancel token has no
   meaning after a reload, and persisting it would leave a dead "running" flag. */
let running = false;
let controller = null;
let queue = [];
let currentKey = null;

let cachedPortfolio = null;
let cachedStamp = '';

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

function render() {
  const st = S.get();
  mount(host,
    stepImport(st),
    st.model ? stepConfirm(st) : null,
    st.model && st.confirmed ? stepPeriod(st) : null,
    st.model && st.confirmed ? stepRun(st) : null,
    st.model && st.confirmed ? stepReports(st) : null,
  );
  renderRail();
}

/* ========================================================================
   Step 1 — import
   ======================================================================== */

function stepImport(st) {
  const drop = el('div', { class: 'dropzone' }, [
    el('div', { style: { marginBottom: '6px' } }, [icon('upload', 22)]),
    el('h3', { text: st.file ? `Loaded: ${st.file.name}` : 'Drop your tracker workbook here, or click to choose' }),
    el('p', { class: 'small muted', style: { marginTop: '4px' } }, [
      'Sheet 1 the site register, sheet 2 the task template, then one sheet per site.',
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

  drop.addEventListener('click', async () => load(await pickFile('.xlsx,.xlsm,.xls')));
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('is-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('is-over'));
  drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('is-over'); load(e.dataTransfer?.files?.[0]); });

  return section('1', 'Load the workbook', el('div', { class: 'sheet__body' }, [drop]));
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
  const prevReport = per.carryPrevious !== false ? st.reports?.[key] : null;
  const anything = per.label || per.previousDate || per.previousWeek || per.notes || prevReport;
  if (!anything) return payload;

  return {
    ...payload,
    previousPeriod: {
      reviewLabel: per.label || undefined,
      previousAnalysisDate: per.previousDate || (prevReport ? String(prevReport.at).slice(0, 10) : undefined),
      previousWeekCovered: per.previousWeek || undefined,
      whatHappenedSince: per.notes || undefined,
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
        input += T.estimateTokens(JSON.stringify(withPeriod(buildMasterPayload(p), MASTER))) + 1400;
      continue;
    }
    // A selection can outlive the workbook it was made against — for instance
    // after loading a session file recorded from a different one. Skip codes
    // that no longer exist rather than estimating a payload for a ghost.
    const site = p.sites.find((s) => s.code === k);
    if (!site) continue;
    input += T.estimateTokens(JSON.stringify(withPeriod(buildSitePayload(site), k))) + 1400;   // + schema and instructions
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
  if (!S.hasApiKey()) { toast('Add an API key in Settings first.', 'hivis'); openSettings(); return; }

  if (st.settings.reviewPayload) {
    const ok = await showPayload(keys[0], true, keys.length);
    if (!ok) return;
  }

  running = true;
  controller = new AbortController();
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
      const payload = withPeriod(key === MASTER ? buildMasterPayload(p) : buildSitePayload(site), key);
      const report = await runAI({
        kind: key === MASTER ? 'master' : 'site',
        payload, settings: S.get().settings, signal: controller.signal,
      });
      report.periodContext = payload.previousPeriod || null;
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

/* ========================================================================
   Step 4 — reports
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
function reportCard(rep, p) {
  const r = rep.result || {};
  const a = rep.kind === 'master' ? null : p.sites.find((s) => s.code === rep.key);
  const wrap = el('div', { class: 'sheet report' });
  const open = { v: true };

  const head = el('button', { class: 'report__head' }, [
    el('span', { class: 'grow row', style: { gap: '8px' } }, [
      icon('insights', 15),
      el('strong', { text: rep.title }),
      r.conclusions?.verdict ? el('span', {
        class: 'chip',
        dataset: { tone: r.conclusions.verdict === 'on_track' ? 'sign' : (r.conclusions.verdict === 'at_risk' ? 'hivis' : 'survey') },
        text: String(r.conclusions.verdict).replace(/_/g, ' '),
      }) : null,
    ]),
    el('span', { class: 'xs num', text: `${rep.providerLabel} · ${T.formatTokens(rep.tokens.input + rep.tokens.output)} tok${rep.tokens.estimated ? ' est' : ''}` }),
  ]);

  const bodyEl = el('div', { class: 'report__body' });
  head.addEventListener('click', () => { open.v = !open.v; bodyEl.classList.toggle('hidden', !open.v); });

  const prose = (text) => (text ? el('div', { class: 'aiprose' }, [el('p', { text })]) : null);
  const list = (title, arr, fmt) => {
    if (!Array.isArray(arr) || !arr.length) return null;
    return el('div', {}, [
      el('h4', { style: { margin: '12px 0 4px' }, text: title }),
      el('ul', {}, arr.map((x) => el('li', { text: typeof x === 'string' ? x : fmt(x) }))),
    ]);
  };
  const sec = (n, title, ...kids) => el('div', { class: 'rsec' }, [
    el('div', { class: 'rsec__head' }, [
      el('span', { class: 'rsec__n', text: String(n) }),
      el('h4', { text: title }),
    ]),
    el('div', { class: 'rsec__body' }, kids.filter(Boolean)),
  ]);
  const figs = (items) => el('div', { class: 'figs' }, items.filter(Boolean).map((f) => el('div', {
    class: 'fig', dataset: f.tone ? { tone: f.tone } : {},
  }, [
    el('div', { class: 'fig__k', text: f.k }),
    el('div', { class: 'fig__v', text: f.v }),
    f.n ? el('div', { class: 'fig__n', text: f.n }) : null,
  ])));

  /* ---- 1. Site details and introduction ---- */
  bodyEl.appendChild(sec(1, rep.kind === 'master' ? 'Programme details' : 'Site details and introduction',
    a ? figs([
      { k: 'Site', v: a.code, n: a.description || undefined },
      { k: 'Wave', v: a.wave || '—' },
      { k: 'Coordinator', v: a.coordinator || '—' },
      { k: 'Register status', v: a.registerStatus || '—' },
      { k: 'Target submission', v: a.target ? fmtDate(a.target) : '—' },
    ]) : figs([
      { k: 'Sites', v: String(p.siteCount) },
      { k: 'With weekly data', v: String(p.sitesWithData), tone: p.sitesWithData < p.siteCount ? 'hivis' : 'sign' },
      { k: 'Live tasks', v: String(p.totalTasks) },
      { k: 'Complete', v: `${p.pctByWeight}%` },
    ]),
    prose(r.introduction)));

  /* ---- 2. Timeline ---- */
  const tl = a?.timeline;
  bodyEl.appendChild(sec(2, 'Timeline',
    figs([
      { k: 'Today', v: fmtDate(tl?.today || new Date().toISOString().slice(0, 10)) },
      { k: 'Weeks reported', v: tl ? String(tl.weeksElapsed) : '—', n: tl?.reportingWeekLabel ? `to ${tl.reportingWeekLabel}` : undefined },
      {
        k: 'Weeks to target', v: tl?.weeksRemainingToTarget == null ? '—' : String(tl.weeksRemainingToTarget),
        tone: tl?.weeksRemainingToTarget != null && tl.weeksRemainingToTarget < 0 ? 'survey' : undefined,
      },
      { k: 'Planned span', v: tl?.totalPlannedWeeks ? `${tl.totalPlannedWeeks} wks` : '—', n: tl?.percentOfPlannedTimeUsed != null ? `${tl.percentOfPlannedTimeUsed}% of time used` : undefined },
      {
        k: 'Forecast', v: a && a.forecastRecent ? `${a.forecastRecent.weeksNeeded} wks more` : '—',
        n: a?.forecastRecent?.finishDate ? fmtDate(a.forecastRecent.finishDate) : undefined,
        tone: a?.slipWeeks > 0 ? 'survey' : undefined,
      },
      {
        k: 'Against target', v: a?.slipWeeks == null ? '—' : (a.slipWeeks > 0 ? `+${a.slipWeeks}w late` : `${Math.abs(a.slipWeeks)}w spare`),
        tone: a?.slipWeeks > 0 ? 'survey' : 'sign',
      },
    ]),
    rep.periodContext?.previousAnalysisDate || rep.periodContext?.reviewLabel
      ? el('p', { class: 'small muted', text: `Previous analysis: ${rep.periodContext.previousAnalysisDate ? fmtDate(rep.periodContext.previousAnalysisDate) : 'not recorded'}${rep.periodContext.previousWeekCovered ? `, covering ${rep.periodContext.previousWeekCovered}` : ''}.` })
      : el('p', { class: 'small muted', text: 'No previous analysis recorded for comparison.' }),
    prose(r.timelineNote)));

  /* ---- 3. Task status ---- */
  bodyEl.appendChild(sec(3, 'Task status',
    a ? figs([
      { k: 'Total tasks', v: String(a.taskCount), n: a.naCount ? `${a.naCount} marked N/A` : undefined },
      { k: 'Live', v: String(a.liveCount) },
      { k: 'Completed', v: String(a.finishedCount), tone: 'sign' },
      { k: 'Pending', v: String(a.remaining), tone: a.remaining ? 'hivis' : 'sign' },
      { k: 'In progress', v: String(a.wipCount) },
      { k: 'Not started', v: String(a.notStartedCount) },
      { k: 'Complete', v: `${a.pctByWeight}%`, n: `${a.pctByCount}% by count` },
      { k: 'Rate', v: `${a.recentVelocity}/wk`, n: `all-time ${a.avgVelocity}/wk` },
    ]) : figs([
      { k: 'Live tasks', v: String(p.totalTasks) },
      { k: 'Completed', v: String(p.totalDone), tone: 'sign' },
      { k: 'Pending', v: String(p.totalTasks - p.totalDone), tone: 'hivis' },
      { k: 'Past target week', v: String(p.totalOverdue), tone: p.totalOverdue ? 'survey' : 'sign' },
      { k: 'Added after kickoff', v: String(p.totalAdditional) },
    ]),
    prose(r.taskStatusInterpretation),
    a && a.overdue?.length
      ? list(`Past their target week (${a.overdue.length})`, a.overdue.slice(0, 10),
          (x) => `${x.taskId} ${x.name} — target ${x.targetWeek}, ${x.weeksLate}w late, ${x.status}`)
      : null,
    a && a.scopeGrowth?.additionalCount
      ? list(`Added after kickoff (${a.scopeGrowth.additionalCount}, ${a.scopeGrowth.growthPct}% growth)`, a.scopeGrowth.items.slice(0, 8),
          (x) => `${x.taskId} ${x.name}${x.added ? ` — added ${fmtDate(x.added)}` : ''}`)
      : null));

  /* ---- 4. Prerequisites ---- */
  const pre = a?.prereqs || [];
  const preOut = a?.outstandingPrereqs || [];
  bodyEl.appendChild(sec(4, 'Prerequisites',
    figs([
      { k: 'Total', v: String(pre.length) },
      { k: 'Received', v: String(Math.max(0, pre.length - preOut.length)), tone: 'sign' },
      { k: 'Outstanding', v: String(preOut.length), tone: preOut.length ? 'hivis' : 'sign' },
      { k: 'Overdue', v: String(preOut.filter((x) => x.overdue).length), tone: preOut.some((x) => x.overdue) ? 'survey' : 'sign' },
      { k: 'Tasks held up', v: String(preOut.reduce((n, x) => n + x.blocksCount, 0)), tone: 'hivis' },
    ]),
    prose(r.prerequisiteInterpretation),
    list('Outstanding', preOut, (x) => `${x.id} ${x.name} — from ${x.provider || 'not recorded'}${x.byWeek ? `, required by week ${x.byWeek}` : ''}${x.overdue ? ' (OVERDUE)' : ''}, holding up ${x.blocksCount} task${x.blocksCount === 1 ? '' : 's'}`)));

  /* ---- 5. Waiting on / blocked ---- */
  const stuck = a?.stuckDetail || [];
  bodyEl.appendChild(sec(5, 'Waiting on and blocked',
    figs([
      { k: 'Stuck now', v: String(a ? a.stuckCount : p.totalStuck), tone: (a ? a.stuckCount : p.totalStuck) ? 'survey' : 'sign' },
      { k: 'Blocked', v: String(a?.blockedCount ?? '—'), n: 'inside your control' },
      { k: 'Waiting on', v: String(a?.waitingCount ?? '—'), n: 'outside it' },
      { k: 'Longest', v: stuck.length ? `${stuck[0].weeksStuck}w` : '—', n: stuck.length ? stuck[0].taskId : undefined, tone: stuck[0]?.weeksStuck >= 3 ? 'survey' : undefined },
      { k: 'Open log entries', v: String((a?.openLog || []).length) },
    ]),
    prose(r.blockedInterpretation),
    list('Currently stuck', stuck.slice(0, 12),
      (x) => `${x.taskId} ${x.name} — ${x.status} ${x.weeksStuck}w${x.waitingOn ? `, on ${x.waitingOn}` : ''}${x.reason ? `: ${x.reason}` : ' (no reason logged)'}`),
    a && a.stalled?.length
      ? list(`Stalled in WIP (${a.stalled.length})`, a.stalled.slice(0, 8), (x) => `${x.taskId} ${x.name} — ${x.weeksInWip} weeks in WIP with no change`)
      : null));

  /* ---- 6. Additional interpretations ---- */
  const add = r.additional || {};
  bodyEl.appendChild(sec(6, 'Additional interpretations',
    a && a.risks?.length
      ? el('div', {}, [
          el('h4', { style: { margin: '0 0 4px' }, text: 'Computed by the app' }),
          el('ul', {}, a.risks.slice(0, 8).map((x) => el('li', { text: `[${x.level}] ${x.title} — ${x.detail}` }))),
        ])
      : null,
    !a && p.risks?.length
      ? el('div', {}, [
          el('h4', { style: { margin: '0 0 4px' }, text: 'Computed by the app' }),
          el('ul', {}, p.risks.slice(0, 8).map((x) => el('li', { text: `[${x.level}] ${x.title} — ${x.detail}` }))),
        ])
      : null,
    list('Risks identified in review', add.risks, (x) => `[${x.impact || '—'}] ${x.risk} — ${x.why}`),
    list('Patterns', add.patterns, (x) => x),
    list('Systemic issues', add.systemicIssues, (x) => `${x.issue} (${(x.sitesAffected || []).join(', ')}) — ${x.rootCauseHypothesis}${x.fixOnceCentrally ? `. Fix centrally: ${x.fixOnceCentrally}` : ''}`),
    list('Resource concerns', add.resourceConcerns, (x) => `${x.resource}: ${x.concern} — ${x.suggestedAction}`),
    list('Actions', add.actions, (x) => `[${x.priority || 'action'}] ${x.action}${x.owner ? ` — ${x.owner}` : ''}${x.byWhen ? ` by ${x.byWhen}` : ''}${x.expectedEffect ? `. ${x.expectedEffect}` : ''}`),
    list('Watch next week', add.watchNextWeek, (x) => x),
    list('Going well', add.whatIsGoingWell, (x) => x),
    a && a.dataIssues?.length
      ? list(`Data quality (${a.dataIssues.length})`, a.dataIssues.slice(0, 8), (x) => x.detail)
      : null));

  /* ---- 7. Visualisations ---- */
  const charts = el('div');
  bodyEl.appendChild(sec(7, 'Visualisation', prose(r.visualisationNote), charts));
  queueMicrotask(() => drawReportCharts(charts, a, p));

  /* ---- 8. Conclusions ---- */
  const c = r.conclusions || {};
  bodyEl.appendChild(sec(8, 'Conclusions',
    el('div', { class: 'row row--wrap', style: { gap: '6px', marginBottom: '10px' } }, [
      c.verdict ? el('span', {
        class: 'chip',
        dataset: { tone: c.verdict === 'on_track' ? 'sign' : (c.verdict === 'at_risk' ? 'hivis' : 'survey') },
        text: String(c.verdict).replace(/_/g, ' '),
      }) : null,
      c.confidence ? el('span', { class: 'chip', dataset: { tone: 'conc' }, text: `${c.confidence} confidence` }) : null,
    ]),
    c.confidenceReason ? el('p', { class: 'small muted', text: c.confidenceReason }) : null,
    prose(c.statement),
    list('Before the next review', c.nextSteps, (x) => x)));

  /* ---- footer ---- */
  bodyEl.appendChild(el('div', { class: 'rsec' }, [
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

  mount(wrap, head, bodyEl);
  return wrap;
}

/** Charts for a report. Site reports get the full set; the master gets the comparison. */
function drawReportCharts(host, a, p) {
  const add = (fn, ...args) => {
    const h = el('div');
    host.appendChild(h);
    try { fn(h, ...args); } catch (e) { console.error('chart failed', e); h.remove(); }
  };
  if (!a) {
    add(renderSiteBars, p.sites);
    add(renderResourceLoad, p.resourceLoad, { showSites: true });
    return;
  }
  if (a.noData) {
    mount(host, el('div', { class: 'empty' }, [
      el('h3', { text: 'Nothing to plot' }),
      el('p', { text: 'This site has no weekly status recorded, so there is no trend, no throughput and no grid to draw.' }),
    ]));
    return;
  }
  add(renderStatusMix, a);
  add(renderCurve, a.curve);
  add(renderThroughput, a.velocity);
  add(renderCategoryProgress, a.categories);
  if (a.stuckDetail.length) add(renderBlockerAges, a.stuckDetail);
  add(renderResourceLoad, a.resources);
  add(renderStatusGrid, a, { maxRows: 60 });
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
            rep.key === MASTER ? buildMasterPayload(p) : buildSitePayload(p.sites.find((s) => s.code === rep.key)),
            rep.key,
          );
          const next = await runAI({
            kind: rep.kind, payload, settings: S.get().settings,
            followUp: q, signal: controller.signal,
          });
          S.update((s) => { s.reports = { ...s.reports, [rep.key]: next }; });
          toast('Updated.', 'sign');
        } catch (e) {
          toast(e instanceof CancelledError ? 'Cancelled.' : e.message, e instanceof CancelledError ? 'hivis' : 'survey');
        } finally {
          running = false; currentKey = null; controller = null; render();
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
      key === MASTER ? buildMasterPayload(p) : buildSitePayload(p.sites.find((s) => s.code === key)),
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
    placeholder: S.hasApiKey() ? S.maskKey(S.getApiKey()) : PROVIDERS[st.settings.provider].keyHint });
  const remember = el('input', { type: 'checkbox', checked: st.settings.rememberKeyForSession ? true : null });
  const review = el('input', { type: 'checkbox', checked: st.settings.reviewPayload ? true : null });
  const maxTok = el('input', { class: 'input', type: 'number', value: st.settings.maxTokens, min: 1024, max: 16000 });

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
    key.placeholder = S.hasApiKey() ? S.maskKey(S.getApiKey()) : PROVIDERS[prov.value].keyHint;
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
        el('label', { text: 'API key' }), key,
        el('span', { class: 'hint' }, ['Get one from ', el('a', { href: PROVIDERS[st.settings.provider].keyUrl, target: '_blank', rel: 'noopener noreferrer', text: 'the provider console' }), '. The field clears after saving.']),
      ]),
      el('div', { class: 'field' }, [
        el('label', { text: 'Max output tokens per call' }), maxTok,
        el('span', { class: 'hint', text: 'Lower is cheaper and less likely to be cut short.' }),
      ]),
    ]),
    el('div', { class: 'stack', style: { marginTop: '10px', gap: '8px' } }, [
      el('label', { class: 'check' }, [remember, el('span', {}, [
        el('strong', { text: 'Keep the key for this browser session. ' }),
        el('span', { class: 'muted', text: 'Held in sessionStorage, gone when the tab closes. Leave off on a shared machine.' }),
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
  return el('div', { class: 'sheet step' }, [
    el('div', { class: 'sheet__head' }, [
      el('span', { class: 'step__n', text: n }),
      el('h3', { text: title }),
      el('span', { class: 'grow' }),
      meta ? el('span', { class: 'xs dim', text: meta }) : null,
    ]),
    bodyEl,
  ]);
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
