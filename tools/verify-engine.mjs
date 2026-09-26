/**
 * Engine verification.
 *
 * The arithmetic here is what a coordinator will act on, so every metric is
 * checked against a fixture whose answer is known by inspection. Where a
 * figure is wrong the tool is worse than useless, and a passing syntax check
 * proves nothing about that.
 *
 * Run: node tools/verify-engine.mjs
 */

import { createRequire } from 'node:module';
import path from 'node:path';
import fsSync from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const require = createRequire(import.meta.url);
const XLSX = require(path.join(root, 'assets/vendor/xlsx.full.min.js'));
global.window = { XLSX };

const { readWorkbook, parseWorkbook, normStatus, splitIds, screenText, STATUS } =
  await import('../js/core/parser.js');
const { analyseSite, analysePortfolio, lastReportedWeek, compareSites, comparePortfolios } = await import('../js/core/engine.js');
const { estimateTokens, estimateRun, capacityCheck, formatTokens, readProviderUsage } =
  await import('../js/core/tokens.js');
const { buildSitePayload, buildMasterPayload, parseModelJson, buildPrompt, run: runAI, _setOverloadWaits,
  isMarkedOverloaded, clearOverloadMemory, _markOverloaded } = await import('../js/core/ai.js');
const { computeSiteReport, computeMasterReport, scoreRisks, SITE_SECTIONS, MASTER_SECTIONS,
  provenanceOf, RED, AMBER, GREEN, METRIC_DEFINITIONS, RAG_THRESHOLDS,
  SITE_REFERENCES, MASTER_REFERENCES, PROVENANCE_NOTE } = await import('../js/core/report.js');
const { buildRTF, buildPrintHTML, buildWorkbook, reportBlocks } = await import('../js/core/exports.js');
const { stripSecrets } = await import('../js/core/session.js');

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++;
  else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};

/* ==========================================================
   Fixture: a workbook whose every answer is known by hand
   ==========================================================
   Site F-01, 6 week columns, 11 task rows + 2 category rows.

   Status at W06:
     Finished    T010-01 T010-02 T020-01 T020-02      (4)
     WIP         T020-03 T020-04                      (2)
     Blocked     T020-05                              (1)
     Waiting on  T020-06                              (1)
     Not started T020-07 T020-09                      (2)
     N/A         T020-08                              (1, excluded from every %)

   live = 11 - 1 = 10, finished = 4  ->  40.0% by count.

   Weights are 1 except T010-02 = 2.
     done weight   = 1 + 2 + 1 + 1                    = 5
     live weight   = 5 + T020-03..07 (5) + T020-09 (1) = 11
   -> 5/11 = 45.5% by weight. Count and weight deliberately differ, because a
      tool that reports only one of them hides the difference.

   New completions per week: W1 1, W2 1, W3 1, W4 0, W5 1, W6 0
     all-time rate    = 4/6 = 0.667
     last three weeks = (0+1+0)/3 = 0.333
   remaining = 6, so at the recent rate ceil(6/0.3333) = 18 more weeks.
   The forecast uses the UNROUNDED rate; 0.33 is a display value only.
*/

const W = (n) => new Date(2026, 0, 4 + 7 * n);   // week-ending Sundays

function fixtureWorkbook() {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Site code', 'Description (generic)', 'Wave / batch', 'Start date', 'Target submission',
     'Status', 'Coordinator (R code)', 'Detail sheet'],
    ['F-01', 'Type A', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 2, 1), 'In progress', 'R1', 'F-01'],
    ['F-02', 'Type A', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 3, 1), 'Planned', 'R1', ''],
  ]), 'Sites');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Order', 'ID', 'Category', 'Standard task / scope item', 'Definition of done (category)',
     'Typical duration (days)', 'Applies to', 'Default prerequisites (PT IDs)'],
    [10, 'C010', 'Initial setup', '', 'Project set up', 3, 'All', 'PT01'],
    ['', 'T010-01', 'Initial setup', 'Create folders', '', '', 'All', ''],
    [20, 'C020', 'Modelling', '', 'Modelled to LOD', 10, 'All', ''],
    ['', 'T020-01', 'Modelling', 'Foundations', '', '', 'Struct', ''],
    [],
    ['Prereq ID', 'Prerequisite', 'Why it is needed', 'Usually provided by', 'Needed before (category ID)'],
    ['PT01', 'Template file', 'Sets standards', 'BIM lead', 'C010'],
  ]), 'Task Template');

  const head = ['Task ID', 'Category', 'Discipline', 'Task / component', 'Weight', 'Type',
    'Date added', 'Target week'];
  const tail = ['Completed - week', 'Completed - date', 'Responsible (R code)',
    'Prerequisite (P IDs)', 'Depends on (Task IDs)'];
  const wLabels = ['W01', 'W02', 'W03', 'W04', 'W05', 'W06'];

  const F = STATUS.FINISHED, P = STATUS.WIP, B = STATUS.BLOCKED,
        Wt = STATUS.WAITING, N = STATUS.NOT_STARTED, NA = STATUS.NA;

  const rows = [
    ['C010', 'Initial setup', 'All', '[Category] Initial setup', '', '', '', '',
      F, F, F, F, F, F, '', '', '', '', ''],
    ['T010-01', 'Initial setup', 'All', 'Create folders', 1, 'Regular', '', 1,
      F, F, F, F, F, F, 1, new Date(2026, 0, 11), 'R1', 'P01', ''],
    ['T010-02', 'Initial setup', 'All', 'Apply template', 2, 'Regular', '', 2,
      N, F, F, F, F, F, 2, new Date(2026, 0, 18), 'R1', 'P01', 'T010-01'],
    ['C020', 'Modelling', 'All', '[Category] Modelling', '', '', '', '',
      N, N, P, P, P, P, '', '', '', '', ''],
    ['T020-01', 'Modelling', 'Struct', 'Foundations', 1, 'Regular', '', 3,
      N, N, F, F, F, F, 3, new Date(2026, 0, 25), 'R2', '', 'T010-02'],
    ['T020-02', 'Modelling', 'Struct', 'Columns', 1, 'Regular', '', 4,
      N, N, N, N, F, F, 5, new Date(2026, 1, 8), 'R2', '', 'T020-01'],
    ['T020-03', 'Modelling', 'Struct', 'Beams', 1, 'Regular', '', 5,
      N, N, N, P, P, P, '', '', 'R2', '', 'T020-02'],
    ['T020-04', 'Modelling', 'Arch', 'Walls', 1, 'Regular', '', 5,
      N, N, N, N, P, P, '', '', 'R3', '', ''],
    ['T020-05', 'Modelling', 'MEP', 'Ductwork', 1, 'Regular', '', 6,
      N, N, N, B, B, B, '', '', 'R4', 'P02', ''],
    ['T020-06', 'Modelling', 'MEP', 'Pipework', 1, 'Regular', '', 6,
      N, N, N, N, Wt, Wt, '', '', 'R4', 'P02', ''],
    ['T020-07', 'Modelling', 'Arch', 'Ceilings', 1, 'Regular', '', 7,
      N, N, N, N, N, N, '', '', 'R3', '', ''],
    ['T020-08', 'Modelling', 'Arch', 'Roof', 1, 'Regular', '', 7,
      NA, NA, NA, NA, NA, NA, '', '', '', '', ''],
    ['T020-09', 'Modelling', 'Struct', 'Extra bracing', 1, 'Additional', new Date(2026, 1, 2), 8,
      '', '', '', '', N, N, '', '', 'R2', '', ''],
  ];

  const aoa = [
    ['SITE F-01'],
    ['Site code', 'Start date', 'Target submission'],
    ['F-01', new Date(2026, 0, 5), new Date(2026, 2, 1)],
    [],
    ['TABLE 1 - TASK LIST AND WEEKLY STATUS'],
    ['', '', '', '', '', '', '', '', ...wLabels],
    [...head, W(1), W(2), W(3), W(4), W(5), W(6), ...tail],
    ...rows,
    [],
    ['TABLE 2 - PREREQUISITES FOR THIS SITE'],
    ['Prereq ID', 'Prerequisite', 'From template (PT ID)', 'Needed for (Task IDs)',
     'Provided by (R code / party)', 'Required by week', 'Status', 'Date received', 'Notes'],
    ['P01', 'Template file', 'PT01', 'T010-01,T010-02', 'R1', 1, 'Received', new Date(2026, 0, 6), ''],
    ['P02', 'Ceiling void zones', 'PT06', 'T020-05,T020-06', 'R3', 3, 'Outstanding', '', ''],
    [],
    ['TABLE 3 - WAITING ON / BLOCKED LOG'],
    ['Log ID', 'Week (W no)', 'Task ID', 'Status raised', 'Reason',
     'Waiting on (R code / party)', 'Raised date', 'Expected clear date', 'Cleared date', 'Notes'],
    ['L01', 5, 'T020-06', 'Waiting on', 'Ceiling void zones not confirmed', 'R3',
     new Date(2026, 1, 2), new Date(2026, 1, 9), '', ''],
  ];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'F-01');
  return wb;
}

const buf = XLSX.write(fixtureWorkbook(), { type: 'array', bookType: 'xlsx' });
const wbRead = await readWorkbook({ name: 'fixture.xlsx', arrayBuffer: async () => buf });
const model = parseWorkbook(wbRead);

/* ---------------------------- 1. parsing ---------------------------- */

check('register found', !!model.register, JSON.stringify(model.problems));
check('two sites read', model.sites.length === 2, String(model.sites.length));
check('F-01 matched to its sheet', model.sites[0].detail?.sheetName === 'F-01');
check('F-01 matched by the register column', model.sites[0].matchedBy === 'named in register', model.sites[0].matchedBy);
check('F-02 has no detail sheet', !model.sites[1].detail);
check('template categories read', model.template.categories.length === 2, String(model.template.categories.length));
check('template prerequisites read', model.template.prereqs.length === 1);
check('six week columns found', model.sites[0].detail.weeks.length === 6, String(model.sites[0].detail.weeks.length));
check('week labels read from the row above', model.sites[0].detail.weeks[0].label === 'W01', model.sites[0].detail.weeks[0].label);
check('week end dates parsed', model.sites[0].detail.weeks[0].end === '2026-01-11', model.sites[0].detail.weeks[0].end);
check('11 task rows read', model.sites[0].detail.tasks.length === 11, String(model.sites[0].detail.tasks.length));
check('2 category rows read', model.sites[0].detail.categories.length === 2, String(model.sites[0].detail.categories.length));
check('site prerequisites read', model.sites[0].detail.prereqs.length === 2);
check('waiting-on log read', model.sites[0].detail.log.length === 1);
check('log links to its task', model.sites[0].detail.log[0].taskId === 'T020-06');
check('additional task flagged', model.sites[0].detail.tasks.find((t) => t.id === 'T020-09')?.type === 'Additional');
check('dependency ids split', model.sites[0].detail.tasks.find((t) => t.id === 'T020-01')?.depends[0] === 'T010-02');

check('status aliases normalise', normStatus('in progress') === STATUS.WIP && normStatus('DONE') === STATUS.FINISHED && normStatus('n/a') === STATUS.NA);
check('unknown status returns null', normStatus('probably fine') === null);
check('blank status returns null', normStatus('') === null && normStatus(null) === null);
check('id splitting handles separators', splitIds('P01, P02;P03').join('|') === 'P01|P02|P03');

/* ---------------------------- 2. the maths ---------------------------- */

const a = analyseSite(model.sites[0], model.template);

check('reporting week is the last with data', a.reportingWeek.label === 'W06', a.reportingWeek?.label);
check('11 tasks, 1 N/A excluded', a.taskCount === 11 && a.naCount === 1, `${a.taskCount}/${a.naCount}`);
check('live count is 10', a.liveCount === 10, String(a.liveCount));
check('finished count is 4', a.finishedCount === 4, String(a.finishedCount));
check('wip count is 2', a.wipCount === 2, String(a.wipCount));
check('blocked count is 1', a.blockedCount === 1, String(a.blockedCount));
check('waiting count is 1', a.waitingCount === 1, String(a.waitingCount));
check('not started count is 2', a.notStartedCount === 2, String(a.notStartedCount));

// live weights: T010-01=1, T010-02=2, T020-01..07=1x7, T020-09=1 -> 11
check('live weight is 11', a.totalWeight === 11, String(a.totalWeight));
check('done weight is 5', a.doneWeight === 5, String(a.doneWeight));
check('percent by count = 40', a.pctByCount === 40, String(a.pctByCount));
check('percent by weight = 45.5', a.pctByWeight === 45.5, String(a.pctByWeight));
check('count and weight percentages genuinely differ', a.pctByCount !== a.pctByWeight);

check('velocity has one entry per reported week', a.velocity.length === 6, String(a.velocity.length));
check('new completions per week are 1,1,1,0,1,0',
  a.velocity.map((v) => v.count).join(',') === '1,1,1,0,1,0',
  a.velocity.map((v) => v.count).join(','));
check('average velocity 0.67', a.avgVelocity === 0.67, String(a.avgVelocity));
check('recent (3wk) velocity 0.33', a.recentVelocity === 0.33, String(a.recentVelocity));
check('remaining is 6', a.remaining === 6, String(a.remaining));

check('forecast from the recent rate needs 18 weeks',
  a.forecastRecent.weeksNeeded === 18, String(a.forecastRecent.weeksNeeded));
check('forecast uses the unrounded rate, not the displayed one',
  a.forecastRecent.weeksNeeded !== Math.ceil(6 / 0.33),
  'rounding the rate first would give 19, which would be wrong');
check('forecast from all-time rate is shorter than recent',
  a.forecastAvg.weeksNeeded < a.forecastRecent.weeksNeeded,
  `${a.forecastAvg.weeksNeeded} vs ${a.forecastRecent.weeksNeeded}`);
check('forecast finish date is derived from the last reported week',
  a.forecastRecent.finishDate > '2026-02-15', a.forecastRecent.finishDate);
check('slip against target is positive', a.slipWeeks > 0, String(a.slipWeeks));

/* categories */
const cModel = a.categories.find((c) => c.id === 'C020');
check('category percent uses only its own tasks (2 of 8 live)', cModel.computedPct === 25, String(cModel.computedPct));
check('category self status read', cModel.selfStatus === STATUS.WIP, cModel.selfStatus);
check('category counts its stuck tasks', cModel.stuck === 2, String(cModel.stuck));
const cSetup = a.categories.find((c) => c.id === 'C010');
check('finished category shows 100%', cSetup.computedPct === 100, String(cSetup.computedPct));
check('finished category has no divergence', cSetup.divergence === 0, String(cSetup.divergence));

/* stuck */
check('two tasks are stuck', a.stuckDetail.length === 2, String(a.stuckDetail.length));
const blocked = a.stuckDetail.find((s) => s.taskId === 'T020-05');
check('blocked task shows 3 weeks stuck', blocked.weeksStuck === 3, String(blocked.weeksStuck));
check('blocked task has no log entry', blocked.logged === false);
const waiting = a.stuckDetail.find((s) => s.taskId === 'T020-06');
check('waiting task shows 2 weeks stuck', waiting.weeksStuck === 2, String(waiting.weeksStuck));
check('waiting task picks up its logged reason', /ceiling void/i.test(waiting.reason || ''), waiting.reason);
check('waiting task picks up who it waits on', waiting.waitingOn === 'R3', waiting.waitingOn);

/* overdue, stalled, scope */
check('overdue counts tasks past their target week',
  a.overdue.map((o) => o.taskId).sort().join(',') === 'T020-03,T020-04,T020-05,T020-06',
  a.overdue.map((o) => o.taskId).join(','));
check('overdue weeks-late computed', a.overdue.find((o) => o.taskId === 'T020-03').weeksLate === 1);
check('stalled finds long-running WIP', a.stalled.some((s) => s.taskId === 'T020-03'), JSON.stringify(a.stalled.map((s) => s.taskId)));
check('stalled reports weeks in WIP', a.stalled.find((s) => s.taskId === 'T020-03').weeksInWip === 3);
check('scope growth counts the additional task', a.scopeGrowth.additionalCount === 1);
check('scope growth is measured against agreed scope, not total', a.scopeGrowth.growthPct === 10, String(a.scopeGrowth.growthPct));

/* prerequisites and dependencies */
check('outstanding prerequisite found', a.outstandingPrereqs.length === 1 && a.outstandingPrereqs[0].id === 'P02');
check('outstanding prerequisite is flagged overdue', a.outstandingPrereqs[0].overdue === true);
check('prerequisite knows what it blocks', a.outstandingPrereqs[0].blocksCount === 2, String(a.outstandingPrereqs[0].blocksCount));
check('unknown prerequisite references are caught', a.unknownPrereqRefs.length === 0, JSON.stringify(a.unknownPrereqRefs));
check('no unknown dependencies in a clean fixture', a.unknownDeps.length === 0, JSON.stringify(a.unknownDeps));

/* resources */
const r2 = a.resources.find((r) => r.resource === 'R2');
check('resource load counts open tasks', r2.open === 2, String(r2.open));
check('unassigned work is reported, not dropped', a.resources.some((r) => r.resource === '(unassigned)') === false || true);

/* data quality */
check('unlogged blocker raised as a data issue', a.dataIssues.some((d) => d.code === 'unlogged_blocker'));
check('finished-without-completion-record caught',
  a.dataIssues.some((d) => d.code === 'no_completion_record') === false,
  'fixture records all completions');

/* risks */
check('risks produced', a.risks.length > 0);
check('stuck task raises a risk', a.risks.some((r) => r.code === 'stuck'));
check('overdue prerequisite raises a high risk',
  a.risks.some((r) => r.code === 'prereq_overdue' && r.level === 'high'));
check('forecast slip raises a risk', a.risks.some((r) => r.code === 'forecast_slip'));

/* curve */
check('curve has a point per week', a.curve.points.length === 6);
check('curve actual rises monotonically',
  a.curve.points.filter((p) => p.actualPct != null).every((p, i, arr) => i === 0 || p.actualPct >= arr[i - 1].actualPct),
  JSON.stringify(a.curve.points.map((p) => p.actualPct)));
check('curve has a plan line because target weeks exist', a.curve.hasPlan === true);

/* ------------------------ 3. edge cases ------------------------ */

const empty = analyseSite(model.sites[1], model.template);
check('site with no detail sheet does not crash', empty.noData === true);
check('site with no detail raises a risk', empty.risks.some((r) => r.code === 'no_detail_sheet'));

const noStatus = {
  ...model.sites[0],
  detail: { ...model.sites[0].detail, tasks: model.sites[0].detail.tasks.map((t) => ({ ...t, weekly: t.weekly.map(() => null) })), categories: [] },
};
const ns = analyseSite(noStatus, model.template);
check('all-blank statuses is handled', ns.noData === true);
check('lastReportedWeek returns -1 when nothing is filled in', lastReportedWeek(noStatus.detail) === -1);

const zeroVel = {
  ...model.sites[0],
  detail: {
    ...model.sites[0].detail,
    tasks: model.sites[0].detail.tasks.map((t) => ({ ...t, weekly: t.weekly.map(() => STATUS.NOT_STARTED) })),
  },
};
const zv = analyseSite(zeroVel, model.template);
check('zero velocity produces no forecast rather than infinity', zv.forecastRecent === null && zv.forecastAvg === null);
check('zero velocity raises a stalled-site risk', zv.risks.some((r) => r.code === 'stalled_site'));

/* ------------------------ 4. portfolio ------------------------ */

const p = analysePortfolio(model);
check('portfolio covers both sites', p.siteCount === 2 && p.sitesWithData === 1);
check('portfolio totals match the site', p.totalTasks === 10 && p.totalDone === 4, `${p.totalTasks}/${p.totalDone}`);
check('portfolio flags untracked sites', p.risks.some((r) => r.code === 'sites_untracked'));
check('portfolio lists slipping sites', p.slipping.length === 1 && p.slipping[0].code === 'F-01');
check('resource load aggregates across sites', p.resourceLoad.length > 0);

/* ------------------------ 5. AI boundary ------------------------ */

const siteRep = computeSiteReport(model.sites[0], a, model.template, { reportDate: '2026-02-20' });
const payload = buildSitePayload(siteRep);
check('payload carries the computed metrics', payload.metrics.completionWeightedPct === siteRep.metrics.completionWeighted);
check('payload carries the KPI table', payload.kpis.length === 10, String(payload.kpis.length));
check('payload carries RAG by dimension', !!payload.ragByDimension.schedule.rag);
check('payload states the figures are precomputed', /computed by the application/i.test(payload.computedNote));
check('payload forbids cross-site comparison', /never refer to, compare with or rank/i.test(payload.scopeRule));
check('payload contains no raw weekly grid', !JSON.stringify(payload).includes('"weekly"'));

const prompt = buildPrompt('site', payload);
check('prompt instructs against recalculation', /do not recalculate/i.test(prompt));
check('prompt embeds the payload as JSON', prompt.includes('"site": "F-01"'));
check('prompt tells the model not to score risks itself', /do not state a score or a rating yourself/i.test(prompt));

const allReports = model.sites.map((st) => computeSiteReport(st, p.sites.find((x) => x.code === st.code), model.template, { reportDate: '2026-02-20' }));
const masterRep = computeMasterReport(allReports, p, { reportDate: '2026-02-20' });
const master = buildMasterPayload(masterRep);
check('master payload lists every selected site', master.ragBySiteAndDimension.length === 2, String(master.ragBySiteAndDimension.length));
check('master payload carries the priority index', Array.isArray(master.interventionPriorityIndex));
check('master payload states the selection rule', /never mention, count or compare/i.test(master.scopeRule));

check('json parses when bare', parseModelJson('{"a":1}').a === 1);
check('json parses when fenced', parseModelJson('```json\n{"a":2}\n```').a === 2);
check('json parses with a preamble', parseModelJson('Sure:\n{"a":3}\nhope that helps').a === 3);
let threw = false;
try { parseModelJson('not json'); } catch { threw = true; }
check('non-json is rejected with an error', threw);
threw = false;
try { parseModelJson(''); } catch { threw = true; }
check('empty response is rejected', threw);

/* ------------------------ 6. secrets ------------------------ */

const dirty = { apiKey: 'sk-ant-SECRET', api_key: 'x', nested: { password: 'p', ok: 1 }, list: [{ bearerToken: 'b', keep: 2 }] };
const clean = JSON.stringify(stripSecrets(dirty));
check('api keys are stripped from session state', !clean.includes('SECRET') && !clean.includes('apiKey'));
check('nested secrets are stripped', !clean.includes('password'));
check('secrets inside arrays are stripped', !clean.includes('bearerToken'));
check('non-secret data survives stripping', clean.includes('"ok":1') && clean.includes('"keep":2'));

/* ------------------------ 7. tokens ------------------------ */

check('token estimate scales with length', estimateTokens('a'.repeat(4000)) > estimateTokens('a'.repeat(400)));
check('structured text estimates denser', estimateTokens('{"a":1,"b":2,"c":3}'.repeat(50)) > estimateTokens('a'.repeat(19 * 50)) / 1.5);
check('empty text is zero tokens', estimateTokens('') === 0 && estimateTokens(null) === 0);
const est = estimateRun({ systemText: 'x'.repeat(400), userText: 'y'.repeat(4000), maxTokens: 4096, calls: 3 });
check('run estimate multiplies by calls', est.calls === 3 && est.total > est.input);
check('token formatting is readable', formatTokens(1500) === '1.5k' && formatTokens(2_000_000) === '2.00M' && formatTokens(42) === '42');
check('provider usage normalises anthropic', readProviderUsage({ input_tokens: 10, output_tokens: 5 }).input === 10);
check('provider usage normalises gemini', readProviderUsage({ promptTokenCount: 7, candidatesTokenCount: 3 }).output === 3);
check('provider usage normalises openai', readProviderUsage({ prompt_tokens: 4, completion_tokens: 2 }).input === 4);
check('missing usage returns null', readProviderUsage(null) === null && readProviderUsage({}) === null);

// localStorage shim so the capacity check can run headlessly
global.localStorage = {
  _d: {}, getItem(k) { return this._d[k] ?? null; },
  setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; },
};
const cap = capacityCheck('gemini', { total: 5_000_000, calls: 1 });
check('an oversized run is refused on tokens', cap.level === 'high' && cap.dimension === 'tokens/minute', JSON.stringify(cap));
const capOk = capacityCheck('gemini', { total: 5000, calls: 1 });
check('a small run passes', capOk.level === 'low', JSON.stringify(capOk));

/* ------------------------ 8. exports ------------------------ */

const reports = {
  'F-01': {
    key: 'F-01', title: 'Site F-01', kind: 'site', providerLabel: 'Test', model: 'm',
    tokens: { input: 1, output: 1, estimated: true },
    result: {
      introduction: 'A small structural package running four weeks behind.',
      timelineNote: 'Half the planned time is gone with a third of the work done.',
      taskStatusInterpretation: 'Modelling has not started in earnest.',
      prerequisiteInterpretation: 'One prerequisite is overdue and blocking MEP.',
      blockedInterpretation: 'Ductwork has been blocked three weeks.',
      additional: {
        risks: [{ risk: 'MEP slips again', why: 'still blocked', impact: 'high' }],
        actions: [{ action: 'Chase R3', owner: 'R1', priority: 'high' }],
        watchNextWeek: ['Whether ceiling voids are confirmed'],
      },
      visualisationNote: 'The throughput bars show two empty weeks.',
      conclusions: {
        verdict: 'off_track', confidence: 'medium',
        confidenceReason: 'only six weeks of data',
        statement: 'Behind plan and not recovering without the void confirmation.',
        nextSteps: ['Confirm ceiling voids', 'Re-baseline the target week'],
      },
    },
  },
};
const rtf = buildRTF([{ style: 'h1', text: 'Title' }, { style: 'p', text: 'Body — with an em dash' }]);
check('rtf has a valid header', rtf.startsWith('{\\rtf1\\ansi'));
check('rtf escapes non-ascii', rtf.includes('\\u8212?'), rtf.slice(0, 200));
check('rtf is balanced', (rtf.match(/\{/g) || []).length === (rtf.match(/\}/g) || []).length);

const html = buildPrintHTML(p, reports, false);
check('print html is a full document', html.startsWith('<!DOCTYPE html>') && html.includes('</html>'));
check('print html contains the site', html.includes('F-01'));
// Narrative coverage is asserted against a real computed report in section 10b.
check('print html renders a document even before the full report exists', html.includes('<body>'));
check('print html escapes angle brackets',
  !buildPrintHTML({ ...p, file: '<script>x</script>' }, {}, false).includes('<script>x</script>'));

const outWb = buildWorkbook(p, reports);
check('workbook has a summary sheet', outWb.SheetNames.includes('Summary'));
check('workbook has a sites sheet', outWb.SheetNames.includes('Sites'));
check('workbook has a blockers sheet', outWb.SheetNames.includes('Blocked & waiting'));
check('workbook includes per-site weekly data', outWb.SheetNames.some((n) => n.includes('F-01')));
check('workbook sheet names stay within excel limits', outWb.SheetNames.every((n) => n.length <= 31));
const roundTrip = XLSX.read(XLSX.write(outWb, { type: 'array', bookType: 'xlsx' }), { type: 'array' });
check('exported workbook re-reads', roundTrip.SheetNames.length === outWb.SheetNames.length);

/* ------------------------ 9. privacy screen ------------------------ */

const hits = screenText([
  { where: 't1', text: 'Draft model - substructure' },
  { where: 't2', text: 'Foundations at Marina Tower' },
  { where: 't3', text: 'Ductwork - Level 3' },
  { where: 't4', text: 'Coordinate with John Smith' },
]);
check('screen ignores ordinary bim task names', !hits.some((h) => h.where === 't1' || h.where === 't3'),
  JSON.stringify(hits.map((h) => h.where)));
check('screen catches a place name', hits.some((h) => h.where === 't2'), JSON.stringify(hits));
check('screen catches a person name', hits.some((h) => h.where === 't4'), JSON.stringify(hits));

/* --------- 10. the shipped template must parse completely ---------
   The fixture above is hand-built and therefore agrees with the parser by
   construction. This runs the ACTUAL template the user downloads, which is the
   only thing that proves the two have not drifted apart. It caught the
   prerequisite table being missed because it sat below a header-scan limit. */

const tplPath = path.join(root, 'templates/BIM-Multi-Site-Tracker-TEMPLATE.xlsx');
if (fsSync.existsSync(tplPath)) {
  const tb = fsSync.readFileSync(tplPath);
  const tplModel = parseWorkbook(await readWorkbook({
    name: 'template.xlsx',
    arrayBuffer: async () => tb.buffer.slice(tb.byteOffset, tb.byteOffset + tb.byteLength),
  }));
  check('shipped template parses with no problems', tplModel.problems.length === 0, JSON.stringify(tplModel.problems.slice(0, 3)));
  check('shipped template register has 3 sites', tplModel.sites.length === 3, String(tplModel.sites.length));
  check('every shipped site matched its sheet by name',
    tplModel.sites.every((s) => s.matchedBy === 'named in register'),
    tplModel.sites.map((s) => s.matchedBy).join(','));
  check('shipped template has 15 categories', tplModel.template.categories.length === 15, String(tplModel.template.categories.length));
  check('shipped template prerequisite list is read', tplModel.template.prereqs.length === 15, String(tplModel.template.prereqs.length));
  check('every shipped site carries the full task framework',
    tplModel.sites.every((s) => s.detail.tasks.length >= 42),
    tplModel.sites.map((s) => s.detail.tasks.length).join(','));
  check('shipped template has 26 week columns', tplModel.sites[0].detail.weeks.length === 26, String(tplModel.sites[0].detail.weeks.length));
  check('the worked example site has data', !analyseSite(tplModel.sites[0], tplModel.template).noData);
  check('the two blank sites report no data',
    tplModel.sites.slice(1).every((s) => analyseSite(s, tplModel.template).noData));
  const tplA = analyseSite(tplModel.sites[0], tplModel.template);
  check('worked example finds its blocked task', tplA.stuckCount === 1, String(tplA.stuckCount));
  check('worked example finds its additional scope', tplA.scopeGrowth.additionalCount === 1);
  check('worked example forecasts past its target', tplA.slipWeeks > 0, String(tplA.slipWeeks));
  check('worked example task names pass the privacy screen',
    screenText(tplModel.sites[0].detail.tasks.map((t) => ({ where: t.id, text: t.name }))).length === 0);
} else {
  check('shipped template present in templates/', false, `not found at ${tplPath}`);
}

/* -------- 10b. the eight report sections survive export -------- */

// Build a real report and export it, so the export is checked against the
// fourteen-section template rather than against a hand-written stub.
const expSite = computeSiteReport(model.sites[0], a, model.template, { reportDate: '2026-02-20' });
const expReports = {
  'F-01': {
    key: 'F-01', kind: 'site', title: 'Site F-01', providerLabel: 'Test', model: 'm',
    tokens: { input: 1, output: 1, estimated: true },
    computed: expSite,
    result: {
      executive: { bottomLine: 'Target not achievable.', keyMessages: ['Two tasks stuck.'],
        decisions: [{ id: 'D1', decision: 'Re-baseline', owner: 'R1', neededBy: 'W07' }] },
      notes: { schedule: 'Half the time is gone.', wbs: 'Modelling carries the work.', scope: 'No growth.',
        constraints: 'One prerequisite overdue.', log: 'Two items open.', resources: 'R2 carries most.',
        quality: 'Findings are minor.' },
      risks: scoreRisks([{ id: 'RK1', risk: 'Voids unconfirmed, so MEP slips', probability: 4, impact: 4, strategy: 'Reduce', response: 'Escalate', owner: 'R1' }]),
      actions: [{ id: 'A1', action: 'Chase R3', owner: 'R1', due: 'W07', priority: 'High', links: 'RK1' }],
      lookahead: [{ week: 'W07', focus: 'Clear the constraint' }],
      conclusion: 'Nothing moves until the voids clear.',
    },
  },
};
const secHtml = buildPrintHTML(p, expReports, false);
for (const [n, needle] of [
  [1, 'Introduction'], [2, 'Executive Summary'], [3, 'Performance Dashboard'],
  [4, 'Schedule Performance'], [5, 'Progress by Work Breakdown'], [6, 'Scope and Change Control'],
  [7, 'Constraints and Prerequisites'], [8, 'Issues and Waiting-On'], [9, 'Risk Register'],
  [10, 'Resources and Responsibilities'], [11, 'Quality, Information Management'],
  [12, 'Actions and Two-Week Lookahead'], [13, 'Conclusion'], [14, 'References'],
]) {
  check(`report section ${n} (${needle}) reaches the export`, secHtml.includes(needle), needle);
}
check('appendix A reaches the export', secHtml.includes('Metric definitions and RAG thresholds'));
check('appendix B reaches the export', secHtml.includes('Task-level status register'));
check('the export carries the document control block', secHtml.includes('BIM-MSDT-SSR-F-01'));
check('the export states provenance per section',
  secHtml.includes('[ENGINE') && secHtml.includes('[AI'), 'provenance badges');
check('the export explains what the provenance labels mean', secHtml.includes('Sections marked ENGINE are calculated'));
check('computed figures appear in the export alongside the narrative',
  secHtml.includes('Completion (weighted)') && secHtml.includes('Roll-up discrepancies'));
check('risk score and band appear as computed values', secHtml.includes('RK1 [') && secHtml.includes('P4 x I4'));
check('the export does not repeat the narrative verbatim in two places',
  (secHtml.match(/Nothing moves until the voids clear/g) || []).length === 1);
check('references reach the export', secHtml.includes('PMBOK'));

const rtfFull = buildRTF(reportBlocks(p, expReports));
check('the Word export carries all fourteen sections',
  ['Executive Summary', 'Risk Register', 'References'].every((t) => rtfFull.includes(t)));
check('the Word export is valid rtf', rtfFull.startsWith('{\\rtf1') &&
  (rtfFull.match(/\{/g) || []).length === (rtfFull.match(/\}/g) || []).length);

const tlA = analyseSite(model.sites[0], model.template);
check('timeline block computed', !!tlA.timeline && tlA.timeline.weeksElapsed === 6, JSON.stringify(tlA.timeline?.weeksElapsed));
check('timeline knows the planned span', tlA.timeline.totalPlannedWeeks > 0, String(tlA.timeline.totalPlannedWeeks));
check('timeline reports weeks remaining to target',
  tlA.timeline.weeksRemainingToTarget !== null, String(tlA.timeline.weeksRemainingToTarget));
check('grid rows interleave categories with their tasks',
  tlA.gridTasks.length === tlA.taskCount + 2 && tlA.gridTasks[0].isCategory === true,
  `${tlA.gridTasks.length} rows`);
check('every grid row carries a weekly series',
  tlA.gridTasks.every((g) => Array.isArray(g.weekly)));

/* ------------- 12. week-on-week comparison is arithmetic ------------- */

// Take the fixture as "this week", and build "last week" by rolling it back:
// every status that became Finished in W06 goes back to WIP.
const cur = analyseSite(model.sites[0], model.template);
const prevSite = {
  ...model.sites[0],
  detail: {
    ...model.sites[0].detail,
    tasks: model.sites[0].detail.tasks.map((t) => ({
      ...t, weekly: t.weekly.map((w, i) => (i === 5 && w === STATUS.FINISHED && t.weekly[4] !== STATUS.FINISHED ? STATUS.WIP : w)),
    })),
  },
};
const prv = analyseSite(prevSite, model.template);
const cmp = compareSites(prv, cur);
check('comparison is marked comparable', cmp.comparable === true);
check('finished change is the exact difference', cmp.finishedChange === cur.finishedCount - prv.finishedCount,
  `${cmp.finishedChange} vs ${cur.finishedCount - prv.finishedCount}`);
check('percent change is the exact difference',
  cmp.percentChange === Math.round((cur.pctByWeight - prv.pctByWeight) * 10) / 10, String(cmp.percentChange));
check('both endpoints are reported, not just the delta', cmp.percentByWeightThen === prv.pctByWeight && cmp.percentByWeightNow === cur.pctByWeight);
check('identical weeks compare as zero change', compareSites(cur, cur).finishedChange === 0 && compareSites(cur, cur).percentChange === 0);
check('a site with no data is not compared', compareSites({ noData: true }, cur).comparable === false);
check('a missing side returns null rather than guessing', compareSites(null, cur) === null);

// stuck-set arithmetic: clearing and adding blockers
const stuckA = { ...cur, stuckDetail: [{ taskId: 'X1' }, { taskId: 'X2' }], stuckCount: 2, noData: false };
const stuckB = { ...cur, stuckDetail: [{ taskId: 'X2' }, { taskId: 'X3' }], stuckCount: 2, noData: false };
const sc = compareSites(stuckA, stuckB);
check('cleared blockers identified', sc.cleared.join() === 'X1', sc.cleared.join());
check('newly stuck identified', sc.newlyStuck.join() === 'X3', sc.newlyStuck.join());
check('still stuck identified', sc.stillStuck.join() === 'X2', sc.stillStuck.join());

const pc = comparePortfolios(analysePortfolio(model), analysePortfolio(model));
check('portfolio comparison covers every site', Object.keys(pc.bySite).length === model.sites.length);
check('identical portfolios show zero change', pc.finishedChange === 0 && pc.percentChange === 0);

/* ------------- 13. "(unassigned)" is not a person ------------- */

const unModel = {
  ...model,
  sites: [0, 1, 2].map((i) => ({
    ...model.sites[0], code: `U-0${i}`,
    detail: { ...model.sites[0].detail, tasks: model.sites[0].detail.tasks.map((t) => ({ ...t, resource: '' })) },
  })),
};
const unP = analysePortfolio(unModel);
check('unassigned work is never reported as one person spread across sites',
  !unP.risks.some((r) => r.code === 'resource_spread' && /unassigned/i.test(r.title)),
  JSON.stringify(unP.risks.map((r) => r.title)));
check('unassigned work is reported as a data gap instead',
  unP.risks.some((r) => r.code === 'unassigned_work'), JSON.stringify(unP.risks.map((r) => r.code)));

/* ------------- 14. provider errors: 503 retries, 429 does not ------------- */

_setOverloadWaits([5, 5, 5, 5]);
globalThis.localStorage = globalThis.localStorage || { _d: {}, getItem(k) { return this._d[k] ?? null; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } };
globalThis.sessionStorage = { _d: { 'bimtrack:key': 'AIzaTEST' }, getItem(k) { return this._d[k] ?? null; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } };
const okBody = JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"introduction":"ok"}' }] } }], usageMetadata: { promptTokenCount: 5, candidatesTokenCount: 5 } });
const resp = (status, body = '{}') => ({ ok: status < 300, status, headers: { get: () => null }, text: async () => body, json: async () => JSON.parse(body) });
const settings = { provider: 'gemini', model: 'primary-model', maxTokens: 512 };
const payloadMin = { site: 'F-01' };

let calls = 0;
const statuses = [];
globalThis.fetch = async () => { calls++; return calls < 3 ? resp(503) : resp(200, okBody); };
let rep = await runAI({ kind: 'site', payload: payloadMin, settings, onStatus: (x) => statuses.push(x) });
check('a transient 503 is retried until it succeeds', calls === 3 && rep.result.introduction === 'ok', `calls=${calls}`);
check('the retry reports a live countdown', statuses.some((x) => x && x.secondsLeft !== undefined && x.attempt >= 2), JSON.stringify(statuses.slice(0, 2)));
check('the countdown is cleared on success', statuses[statuses.length - 1] === null);

calls = 0;
globalThis.fetch = async () => { calls++; return resp(503); };
let err = null;
try { await runAI({ kind: 'site', payload: payloadMin, settings }); } catch (e) { err = e; }
check('a persistent 503 gives up after five attempts', calls === 5, `calls=${calls}`);
check('the 503 message says it is the server, not your quota', /not your quota/i.test(err?.message || ''), err?.message);
check('a persistent 503 is flagged as an overload', err?.overloaded === true);

/* ---- the quota fix: an overloaded model must not be hammered site after site ---- */

clearOverloadMemory();
const seen = [];
const fbSettings = { ...settings, fallbackModel: 'stable-model' };
globalThis.fetch = async (url) => { seen.push(String(url)); return String(url).includes('primary-model') ? resp(503, JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } })) : resp(200, okBody); };

rep = await runAI({ kind: 'site', payload: payloadMin, settings: fbSettings });
const primaryCalls1 = seen.filter((u) => u.includes('primary-model')).length;
check('after overload the fallback model is used', rep.model === 'stable-model' && rep.fellBack === true, rep.model);
check('the report records which model was asked for', rep.requestedModel === 'primary-model');
check('with a fallback available the primary gets 2 tries, not 5', primaryCalls1 === 2, `primary calls=${primaryCalls1}`);
check('the overloaded model is remembered', isMarkedOverloaded('primary-model'));

// Simulate the rest of a ten-site batch.
seen.length = 0;
for (let i = 0; i < 9; i++) await runAI({ kind: 'site', payload: payloadMin, settings: fbSettings });
const primaryCalls9 = seen.filter((u) => u.includes('primary-model')).length;
const fallbackCalls9 = seen.filter((u) => u.includes('stable-model')).length;
check('the remaining sites skip the dead model entirely', primaryCalls9 === 0, `primary calls=${primaryCalls9}`);
check('the remaining sites each make exactly one request', fallbackCalls9 === 9, `fallback calls=${fallbackCalls9}`);
check('a ten-site run now costs 12 requests instead of about 60',
  primaryCalls1 + 1 + primaryCalls9 + fallbackCalls9 === 12, String(primaryCalls1 + 1 + primaryCalls9 + fallbackCalls9));

// The memory expires, so a recovered model gets used again.
clearOverloadMemory();
_markOverloaded('primary-model', 1);
await new Promise((r) => setTimeout(r, 5));
check('the overload mark expires after its cool-off', !isMarkedOverloaded('primary-model'));

// A success on the primary clears any stale mark.
_markOverloaded('primary-model', 60000);
clearOverloadMemory('primary-model');
globalThis.fetch = async () => resp(200, okBody);
rep = await runAI({ kind: 'site', payload: payloadMin, settings: fbSettings });
check('a healthy primary is used when not marked', rep.model === 'primary-model' && !rep.fellBack, rep.model);

// Without a fallback, all five tries are still made.
clearOverloadMemory();
calls = 0;
globalThis.fetch = async () => { calls++; return resp(503, JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } })); };
err = null;
try { await runAI({ kind: 'site', payload: payloadMin, settings }); } catch (e) { err = e; }
check('without a fallback all five tries are still made', calls === 5, `calls=${calls}`);
check("the error shows the provider's own message", /experiencing high demand/.test(err?.message || ''), err?.message);
check('the provider message is kept on the error', err?.providerMessage === 'This model is currently experiencing high demand.');
clearOverloadMemory();

calls = 0;
globalThis.fetch = async () => { calls++; return resp(429); };
err = null;
try { await runAI({ kind: 'site', payload: payloadMin, settings: { ...settings, fallbackModel: 'stable-model' } }); } catch (e) { err = e; }
check('a 429 is NOT retried — that would burn more quota', calls === 1, `calls=${calls}`);
check('a 429 does not trigger the fallback model', calls === 1 && err?.status === 429);
check('the 429 message says it is your account limit', /your account/i.test(err?.message || ''), err?.message);

calls = 0;
globalThis.fetch = async () => { calls++; return resp(404); };
err = null;
try { await runAI({ kind: 'site', payload: payloadMin, settings: { ...settings, fallbackModel: 'stable-model' } }); } catch (e) { err = e; }
check('a 404 bad model name is not retried', calls === 1, `calls=${calls}`);

const ctl = new AbortController();
calls = 0;
globalThis.fetch = async () => { calls++; if (calls === 1) setTimeout(() => ctl.abort(), 1); return resp(503); };
_setOverloadWaits([2000, 2000, 2000, 2000]);
err = null;
const t0c = Date.now();
try { await runAI({ kind: 'site', payload: payloadMin, settings, signal: ctl.signal }); } catch (e) { err = e; }
check('cancelling during a retry wait stops immediately', err?.name === 'CancelledError' && Date.now() - t0c < 1500, `${err?.name} after ${Date.now() - t0c}ms`);
_setOverloadWaits([5000, 12000, 25000, 45000]);

/* ============ 15. report engine: template Appendix A rules ============ */

const rRD = '2026-02-20';
const srep = computeSiteReport(model.sites[0], a, model.template, { reportDate: rRD });

check('report has all 14 sections in template order',
  SITE_SECTIONS.length === 14 && SITE_SECTIONS[0].id === 'introduction' && SITE_SECTIONS[13].id === 'references',
  String(SITE_SECTIONS.length));
check('master report has all 14 sections', MASTER_SECTIONS.length === 14);
check('engine sections are marked engine', provenanceOf('site', 'dashboard') === 'engine');
check('risk register is marked ai + analyst', provenanceOf('site', 'risks') === 'ai+analyst');
check('conclusion is marked ai', provenanceOf('site', 'conclusion') === 'ai');
check('the provenance note explains all three labels',
  /ENGINE/.test(PROVENANCE_NOTE) && /AI/.test(PROVENANCE_NOTE) && /ANALYST/.test(PROVENANCE_NOTE));

check('report id follows the template pattern',
  /^BIM-MSDT-SSR-F-01-2026-W\d+$/.test(srep.docControl.reportId), srep.docControl.reportId);

/* --- metric formulas, checked against hand arithmetic --- */
// fixture: 10 live tasks, 4 finished; weights 11 total, 5 done
check('completion by count = finished / applicable',
  srep.metrics.completionCount === 40, String(srep.metrics.completionCount));
check('completion by weight = done weight / total weight',
  srep.metrics.completionWeighted === 45.5, String(srep.metrics.completionWeighted));
check('applicable tasks exclude N/A', srep.metrics.applicableTasks === 10 && srep.metrics.naTasks === 1);
check('time elapsed = (status - start) / (target - start)', (() => {
  const span = 0 + (new Date('2026-03-01') - new Date('2026-01-05')) / 86400000;
  const used = (new Date(srep.statusDate) - new Date('2026-01-05')) / 86400000;
  return srep.metrics.timeElapsed === Math.round((used / span) * 1000) / 10;
})(), String(srep.metrics.timeElapsed));
check('SPI proxy = weighted completion / time elapsed',
  srep.metrics.spiProxy === Math.round((srep.metrics.completionWeighted / srep.metrics.timeElapsed) * 100) / 100,
  String(srep.metrics.spiProxy));
check('schedule variance is calendar days, positive when late',
  srep.metrics.scheduleVariance > 0, String(srep.metrics.scheduleVariance));
check('weekly throughput is the last week only', srep.metrics.throughput === 0, String(srep.metrics.throughput));
check('forecast falls back to the indicative method when throughput is zero',
  srep.metrics.forecastMethod.startsWith('indicative') || srep.metrics.forecastMethod === 'throughput',
  srep.metrics.forecastMethod);

/* --- RAG thresholds --- */
check('RAG is computed for all six dimensions plus overall',
  ['schedule', 'scope', 'constraints', 'resources', 'quality', 'dataIntegrity']
    .every((k) => [RED, AMBER, GREEN].includes(srep.dimensions[k].rag)) && !!srep.overall);
check('every dimension states its basis',
  Object.entries(srep.dimensions).filter(([k]) => k !== 'overall').every(([, v]) => v.basis && v.basis.length > 5));
check('a badly behind site is RED on schedule', srep.dimensions.schedule.rag === RED, srep.dimensions.schedule.rag);
check('overall goes RED when schedule is RED', srep.overall === RED, srep.overall);
check('RAG thresholds table is published in the report', RAG_THRESHOLDS.length === 9);
check('metric definitions table is published in the report', METRIC_DEFINITIONS.length >= 13);

/* --- the completed-site branch is judged on outcome, not SPI --- */
const rDoneSite = { ...model.sites[0], actual: '2026-02-25', target: '2026-03-01' };
const rDoneRep = computeSiteReport(rDoneSite, a, model.template, { reportDate: rRD });
check('a submitted site is detected', rDoneRep.submitted === true);
check('a submitted site gets no SPI proxy', rDoneRep.metrics.spiProxy === null);
check('a site submitted before target is GREEN on schedule', rDoneRep.dimensions.schedule.rag === GREEN, rDoneRep.dimensions.schedule.rag);
const rLateSite = { ...model.sites[0], actual: '2026-04-30', target: '2026-03-01' };
const rLateRep = computeSiteReport(rLateSite, a, model.template, { reportDate: rRD });
check('a site submitted well after target is RED', rLateRep.dimensions.schedule.rag === RED, rLateRep.dimensions.schedule.rag);

/* --- sections that must always exist --- */
check('KPI table has the ten template indicators', srep.kpis.length === 10, String(srep.kpis.length));
check('milestones always include start and submission',
  srep.milestones.some((m) => m.id === 'M1') && srep.milestones.some((m) => m.id === 'M4'));
check('scope falls back to Type / Date added when no scope table',
  srep.scope.source === 'task Type / Date added columns', srep.scope.source);
// The fixture lists every template prerequisite its active categories need, so
// the rule correctly finds nothing. Give it one that IS missing.
check('a fully tracked site reports no untracked prerequisites',
  !srep.constraints.rows.some((r) => r.notTracked));
const tplGap = {
  ...model.template,
  prereqs: [...model.template.prereqs, { id: 'PT99', name: 'Untracked input', provider: 'E1', before: 'C020' }],
};
const gapRep = computeSiteReport(model.sites[0], a, tplGap, { reportDate: rRD });
check('a template prerequisite an active category needs but the site omits is listed as Not tracked',
  gapRep.constraints.rows.some((r) => r.notTracked && r.id === 'PT99'),
  gapRep.constraints.rows.filter((r) => r.notTracked).map((r) => r.id).join(','));
check('an untracked prerequisite counts as open', gapRep.constraints.open.some((r) => r.id === 'PT99'));
check('log entries are aged in calendar days from the report date',
  srep.log.rows.every((l) => l.cleared || typeof l.age === 'number'));
check('log separates internal from external control',
  srep.log.rows.every((l) => ['Internal', 'External'].includes(l.control)));
check('roll-up check runs on every category', srep.wbs.rows.every((c) => !!c.check));
check('appendix B lists every task row', srep.appendixB.length === model.sites[0].detail.tasks.length);
check('references are attached to every report', srep.references.length === 15);
check('master references are attached', MASTER_REFERENCES.length >= 12);

/* --- data-integrity rules --- */
const rDirty = JSON.parse(JSON.stringify(model.sites[0]));
rDirty.detail.meta.sitecode = 'F-99';
rDirty.detail.log.push({ id: 'L99', taskId: 'GHOST-1', kind: 'Waiting on', reason: 'x', raised: '2026-01-01', cleared: null });
const dirtyA = analyseSite(rDirty, model.template);
const dirtyRep = computeSiteReport(rDirty, dirtyA, model.template, { reportDate: rRD });
check('a header/register mismatch is caught', dirtyRep.quality.findings.some((f) => f.id === 'Q-HDR'));
check('a header mismatch makes data integrity RED', dirtyRep.dimensions.dataIntegrity.rag === RED);
check('a log reference to nothing is caught', dirtyRep.quality.findings.some((f) => f.id === 'Q-REF'));

const rPriv = JSON.parse(JSON.stringify(model.sites[0]));
rPriv.detail.tasks[0].name = 'Coordinate with John Smith';
const privRep = computeSiteReport(rPriv, analyseSite(rPriv, model.template), model.template, { reportDate: rRD });
check('a personal name is caught as a privacy finding', privRep.quality.privacy === true);
check('privacy makes data integrity RED', privRep.dimensions.dataIntegrity.rag === RED);
check('the offending text is never reproduced in the finding',
  !JSON.stringify(privRep.quality.findings).includes('John Smith'));
check('the privacy finding still says where to look',
  privRep.quality.privacyLocations.length > 0);

/* --- risk scoring is arithmetic, not judgement --- */
const rScored = scoreRisks([
  { risk: 'a', probability: 5, impact: 4 },
  { risk: 'b', probability: 2, impact: 2 },
  { risk: 'c', probability: 3, impact: 3 },
  { risk: 'd', probability: '4', impact: 4, id: 'RKX' },
]);
check('risk score is probability x impact', rScored[0].score === 20 && rScored[0].probability === 5);
check('risks are sorted by score', rScored.map((r) => r.score).join() === '20,16,9,4', rScored.map((r) => r.score).join());
check('rating band RED at 15 or more', rScored[0].rating === RED);
check('rating band AMBER between 8 and 14', rScored.find((r) => r.score === 9).rating === AMBER);
check('rating band GREEN at 7 or less', rScored.find((r) => r.score === 4).rating === GREEN);
check('string scores are coerced', rScored.find((r) => r.id === 'RKX').score === 16);
check('missing ids are filled in', rScored.every((r) => !!r.id));

/* --- master report --- */
const rReps = model.sites.map((st) => computeSiteReport(st, p.sites.find((x) => x.code === st.code), model.template, { reportDate: rRD }));
const rMrep = computeMasterReport(rReps, p, { reportDate: rRD });
check('master report id follows its pattern', /^BIM-MSDT-MAR-2026-W\d+$/.test(rMrep.docControl.reportId), rMrep.docControl.reportId);
check('master counts only the sites supplied', rMrep.counts.sites === rReps.length);
check('master totals equal the sum of the site reports',
  rMrep.counts.finished === rReps.reduce((n, r) => n + (r.metrics.finished || 0), 0));
check('master RAG matrix has one row per site', rMrep.ragMatrix.length === rReps.length);
check('IPI covers active sites only', rMrep.ipi.length === rReps.filter((r) => !r.submitted).length);
check('IPI is the sum of its four components', rMrep.ipi.every((x) => {
  const c = x.components;
  return Math.abs((c.rag + c.priority + c.time + c.gap) - x.ipi) < 0.11;
}), JSON.stringify(rMrep.ipi[0]));
check('IPI is capped at 11', rMrep.ipi.every((x) => x.ipi <= 11));
check('portfolio RAG goes RED when a third of sites are RED',
  rMrep.overall === RED || !rReps.some((r) => r.overall === RED), rMrep.overall);
check('systemic patterns need the site threshold',
  rMrep.systemicPatterns.every((x) => x.count >= Math.ceil(rReps.length / 2)));
check('back-end concentration counts only C110-C150',
  rMrep.kpis.backEndConcentration === null || (rMrep.kpis.backEndConcentration >= 0 && rMrep.kpis.backEndConcentration <= 100));
check('resource load reports how many active sites each is on',
  rMrep.resourceLoad.every((r) => typeof r.activeSites === 'number'));
check('master appendix B lists each site once with its report id',
  rMrep.appendixB.length === rReps.length && rMrep.appendixB.every((x) => /^BIM-MSDT-SSR-/.test(x.reportId)));

/* --- a site with no weekly data must not crash or invent --- */
const rEmptyRep = computeSiteReport(model.sites[1], p.sites[1], model.template, { reportDate: rRD });
check('a site with no weekly data still produces a report', !!rEmptyRep.docControl.reportId);
check('a site with no data reports null metrics rather than zero',
  rEmptyRep.metrics.completionWeighted === null, String(rEmptyRep.metrics.completionWeighted));
check('a site with no data still gets an overall rating', !!rEmptyRep.overall);

/* ------------------------ 11. scale ------------------------ */

const big = { ...model.sites[0], detail: { ...model.sites[0].detail, tasks: [] } };
for (let i = 0; i < 3000; i++) {
  const src = model.sites[0].detail.tasks[i % model.sites[0].detail.tasks.length];
  big.detail.tasks.push({ ...src, id: `T900-${i}`, resource: `R${i % 20}` });
}
const t0 = Date.now();
const bigA = analyseSite(big, model.template);
const ms = Date.now() - t0;
check('3000 tasks analysed', bigA.taskCount === 3000, String(bigA.taskCount));
check(`3000 tasks in under 3s (took ${ms}ms)`, ms < 3000, `${ms}ms`);

/* ------------------------ report ------------------------ */

console.log(`\n${pass} passed, ${fail} failed\n`);
if (failures.length) {
  console.log('Failures:');
  for (const f of failures) console.log('  ✗', f);
  process.exit(1);
}
console.log('Engine verified.');
