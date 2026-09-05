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
import { renderCurve, renderThroughput, renderSiteBars, meter } from '../core/charts.js';

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
  host = el('main', { class: 'view view--wide' });
  rail = el('aside', { class: 'siderail' });
  mount(app,
    el('div', { class: 'shell' }, [
      el('header', { class: 'topbar' }, [
        el('div', { class: 'brand' }, [
          el('span', { class: 'brand__mark', text: 'BIM' }),
          el('span', {}, [
            el('div', { class: 'brand__name', text: 'Multi-Site Delivery Tracker' }),
            el('div', { class: 'brand__sub', text: 'Codes only · runs entirely in your browser' }),
          ]),
        ]),
        el('span', { class: 'topbar__spacer' }),
        el('button', { class: 'btn btn--sm', onclick: onSaveSession }, [icon('save', 13), 'Save session']),
        el('button', { class: 'btn btn--sm', onclick: onLoadSession }, [icon('upload', 13), 'Load session']),
        el('button', { class: 'btn btn--sm', onclick: openSettings }, [icon('settings', 13), 'Settings']),
      ]),
      el('div', { class: 'body' }, [host, rail]),
    ]),
  );
  S.subscribe(debounce(render, 30));
  S.restore().then((found) => {
    render();
    if (found && S.get().model) {
      toast('Previous session restored from this browser.', 'sign');
    }
    if (!xlsxAvailable()) toast('The spreadsheet reader did not load — imports will not work.', 'survey', 9000);
  });
}

function render() {
  const st = S.get();
  mount(host,
    stepImport(st),
    st.model ? stepConfirm(st) : null,
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
   Step 3 — choose and run
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

  return section('3', 'Choose what to analyse', body, `${Object.keys(st.reports).length} of ${p.sites.length + 1} done`);
}

function estimateSelection(st, p, keys) {
  let input = 0;
  for (const k of keys) {
    if (k === MASTER) {
      input += T.estimateTokens(JSON.stringify(buildMasterPayload(p))) + 900;
      continue;
    }
    // A selection can outlive the workbook it was made against — for instance
    // after loading a session file recorded from a different one. Skip codes
    // that no longer exist rather than estimating a payload for a ghost.
    const site = p.sites.find((s) => s.code === k);
    if (!site) continue;
    input += T.estimateTokens(JSON.stringify(buildSitePayload(site))) + 900;   // + schema and instructions
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
      const payload = key === MASTER ? buildMasterPayload(p) : buildSitePayload(site);
      const report = await runAI({
        kind: key === MASTER ? 'master' : 'site',
        payload, settings: S.get().settings, signal: controller.signal,
      });
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

  return section('4', 'Reports', body, keys.length ? `${keys.length} generated` : '');
}

function reportCard(rep, p) {
  const r = rep.result || {};
  const wrap = el('div', { class: 'sheet report' });
  const open = { v: true };

  const head = el('button', { class: 'report__head' }, [
    el('span', { class: 'grow row', style: { gap: '8px' } }, [
      icon('insights', 14),
      el('strong', { text: rep.title }),
      r.verdict ? el('span', {
        class: 'chip',
        dataset: { tone: r.verdict === 'on_track' ? 'sign' : (r.verdict === 'at_risk' ? 'hivis' : 'survey') },
        text: String(r.verdict).replace(/_/g, ' '),
      }) : null,
    ]),
    el('span', { class: 'xs dim num', text: `${rep.providerLabel} · ${rep.model} · ${T.formatTokens(rep.tokens.input + rep.tokens.output)} tok${rep.tokens.estimated ? ' (est)' : ''}` }),
  ]);

  const bodyEl = el('div', { class: 'report__body' });
  head.addEventListener('click', () => { open.v = !open.v; bodyEl.classList.toggle('hidden', !open.v); });

  if (r.headline) bodyEl.appendChild(el('h3', { text: r.headline }));
  if (r.confidenceReason) bodyEl.appendChild(el('p', { class: 'small muted', text: `Confidence ${r.confidence || '—'}: ${r.confidenceReason}` }));
  if (r.summary) bodyEl.appendChild(el('p', { text: r.summary }));

  const list = (title, arr, fmt) => {
    if (!Array.isArray(arr) || !arr.length) return;
    bodyEl.appendChild(el('h4', { text: title }));
    bodyEl.appendChild(el('ul', {}, arr.map((x) => el('li', { text: typeof x === 'string' ? x : fmt(x) }))));
  };
  list('What is driving it', r.whatIsDrivingIt, (x) => `${x.point} — ${x.evidence}${x.effect ? ` (${x.effect})` : ''}`);
  list('Bottlenecks', r.bottlenecks, (x) => `${x.taskId || ''} ${x.issue}${x.whoToChase ? ` — chase ${x.whoToChase}` : ''}${x.suggestedAction ? `. ${x.suggestedAction}` : ''}`);
  list('Site ranking', r.siteRanking, (x) => `${x.site}: ${x.standing} — ${x.why}`);
  list('Systemic issues', r.systemicIssues, (x) => `${x.issue} (${(x.sitesAffected || []).join(', ')}) — ${x.rootCauseHypothesis}${x.fixOnceCentrally ? `. Fix centrally: ${x.fixOnceCentrally}` : ''}`);
  list('Resource concerns', r.resourceConcerns, (x) => `${x.resource}: ${x.concern} — ${x.suggestedAction}`);
  list('Sequencing and overlaps', r.sequencingAndOverlaps, (x) => x);
  list('Actions', r.actions || r.priorityActions, (x) => `[${x.priority || 'action'}] ${x.action}${x.owner ? ` — ${x.owner}` : ''}${x.byWhen ? ` by ${x.byWhen}` : ''}`);
  list('Watch next week', r.watchNextWeek, (x) => x);
  list('Going well', r.whatIsGoingWell, (x) => x);
  list('Data gaps', r.dataGaps, (x) => x);

  bodyEl.appendChild(el('div', { class: 'row row--wrap', style: { marginTop: '12px' } }, [
    el('button', {
      class: 'btn btn--sm', disabled: running ? true : null,
      onclick: () => continueReport(rep, p),
    }, ['Continue — ask a follow-up']),
    el('button', {
      class: 'btn btn--sm', disabled: running ? true : null,
      onclick: () => {
        S.update((s) => { const n = { ...s.reports }; delete n[rep.key]; s.reports = n; });
        toast('Report removed.');
      },
    }, [icon('trash', 13), 'Remove']),
    rep.followUp ? el('span', { class: 'xs dim', text: `Asked: ${rep.followUp}` }) : null,
  ]));

  mount(wrap, head, bodyEl);
  return wrap;
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
          const payload = rep.key === MASTER
            ? buildMasterPayload(p)
            : buildSitePayload(p.sites.find((s) => s.code === rep.key));
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
    const payload = key === MASTER
      ? buildMasterPayload(p)
      : buildSitePayload(p.sites.find((s) => s.code === key));
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

function renderRail() {
  const st = S.get();
  const p = portfolio();
  const nodes = [];

  nodes.push(tokenRail(st));

  if (p && st.confirmed) {
    const bars = el('div');
    nodes.push(el('div', { class: 'sheet' }, [
      el('div', { class: 'sheet__head' }, [el('h3', { text: 'Site progress' }), el('span', { class: 'xs dim', text: 'computed' })]),
      el('div', { class: 'sheet__body' }, [bars]),
    ]));
    renderSiteBars(bars, p.sites);

    if (p.risks.length) {
      nodes.push(el('div', { class: 'sheet' }, [
        el('div', { class: 'sheet__head' }, [el('h3', { text: 'Computed risks' }), el('span', { class: 'badge', dataset: { tone: p.risks.some((r) => r.level === 'high') ? 'survey' : 'hivis' }, text: String(p.risks.length) })]),
        el('div', { class: 'sheet__body sheet__body--flush' }, [
          el('ul', { class: 'register' }, p.risks.slice(0, 8).map((r) => el('li', { class: 'register__item', dataset: { sev: r.level } }, [
            el('div', { class: 'register__spine' }),
            el('div', { class: 'register__body' }, [
              el('div', { class: 'register__title', text: r.title }),
              el('div', { class: 'register__detail', text: r.detail }),
            ]),
          ]))),
        ]),
      ]));
    }

    const firstWithData = p.sites.find((s) => !s.noData);
    if (firstWithData) {
      const curveHost = el('div');
      const thruHost = el('div');
      const sel = el('select', { class: 'select', style: { maxWidth: '130px' } },
        p.sites.filter((s) => !s.noData).map((s) => el('option', { value: s.code, text: s.code })));
      const draw = () => {
        const s = p.sites.find((x) => x.code === sel.value) || firstWithData;
        renderCurve(curveHost, s.curve);
        renderThroughput(thruHost, s.velocity);
      };
      sel.addEventListener('change', draw);
      nodes.push(el('div', { class: 'sheet' }, [
        el('div', { class: 'sheet__head' }, [el('h3', { text: 'Timeline' }), el('span', { class: 'grow' }), sel]),
        el('div', { class: 'sheet__body' }, [curveHost, el('h4', { text: 'Finished per week', style: { marginTop: '12px' } }), thruHost]),
      ]));
      queueMicrotask(draw);
    }
  }

  mount(rail, ...nodes);
}

function tokenRail(st) {
  const usage = T.getUsage();
  const win = T.requestWindows(st.settings.provider);
  const lim = T.getLimits(st.settings.provider);
  const reasons = T.recentRequestReasons();

  const bar = (used, max, tone) => {
    const p = max ? clamp((used / max) * 100, 0, 100) : 0;
    return el('div', { class: 'meter', style: { marginTop: '3px' } }, [
      el('div', { class: 'meter__fill', dataset: { tone: p > 85 ? 'survey' : (p > 60 ? 'hivis' : tone) }, style: { width: `${p}%` } }),
    ]);
  };

  return el('div', { class: 'sheet' }, [
    el('div', { class: 'sheet__head' }, [
      el('h3', { text: 'Usage' }),
      el('span', { class: 'grow' }),
      el('span', { class: 'xs dim', text: PROVIDERS[st.settings.provider]?.label || st.settings.provider }),
    ]),
    el('div', { class: 'sheet__body stack', style: { gap: '10px' } }, [
      el('div', {}, [
        el('div', { class: 'row row--between xs' }, [
          el('span', { class: 'dim', text: 'Requests this minute' }),
          el('span', { class: 'num', text: `${win.lastMinute}${lim.rpm ? ` / ${lim.rpm}` : ''}` }),
        ]),
        lim.rpm ? bar(win.lastMinute, lim.rpm, 'blueprint') : null,
        win.lastMinute && lim.rpm && win.lastMinute >= lim.rpm
          ? el('div', { class: 'xs', style: { color: 'var(--survey)' }, text: `Capacity returns in about ${win.nextMinuteSlotIn}s` })
          : null,
      ]),
      el('div', {}, [
        el('div', { class: 'row row--between xs' }, [
          el('span', { class: 'dim', text: 'Requests today' }),
          el('span', { class: 'num', text: `${win.lastDay}${lim.rpd ? ` / ${lim.rpd}` : ''}` }),
        ]),
        lim.rpd ? bar(win.lastDay, lim.rpd, 'blueprint') : null,
      ]),
      el('div', { class: 'row row--between xs' }, [
        el('span', { class: 'dim', text: 'Tokens used (all time)' }),
        el('span', { class: 'num', text: `${T.formatTokens(usage.total)}${usage.estimated ? ' est' : ''}` }),
      ]),
      el('div', { class: 'row row--between xs' }, [
        el('span', { class: 'dim', text: 'In / out' }),
        el('span', { class: 'num', text: `${T.formatTokens(usage.input)} / ${T.formatTokens(usage.output)}` }),
      ]),
      el('div', { class: 'row row--between xs' }, [
        el('span', { class: 'dim', text: 'Calls made' }),
        el('span', { class: 'num', text: String(usage.calls) }),
      ]),
      reasons.length > 1
        ? el('div', { class: 'xs dim', text: `Recent requests: ${reasons.map((r) => `${r.n}× ${r.reason}`).join(', ')} — retries count against your limit too.` })
        : null,
      el('div', { class: 'xs dim', text: `${lim.label}${lim.isDefault ? ' (default figures — edit in Settings to match your account)' : ' (your figures)'}` }),
      el('div', { class: 'row', style: { gap: '6px' } }, [
        el('button', { class: 'btn btn--sm', onclick: () => { T.resetUsage(); render(); toast('Token counter reset.'); } }, ['Reset tokens']),
        el('button', { class: 'btn btn--sm', onclick: () => { T.resetRequests(st.settings.provider); render(); toast('Request counter reset.'); } }, ['Reset requests']),
      ]),
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
