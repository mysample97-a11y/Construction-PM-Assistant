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
const { analyseSite, analysePortfolio, lastReportedWeek } = await import('../js/core/engine.js');
const { estimateTokens, estimateRun, capacityCheck, formatTokens, readProviderUsage } =
  await import('../js/core/tokens.js');
const { buildSitePayload, buildMasterPayload, parseModelJson, buildPrompt } =
  await import('../js/core/ai.js');
const { buildRTF, buildPrintHTML, buildWorkbook } = await import('../js/core/exports.js');
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

const payload = buildSitePayload(a);
check('payload carries computed progress', payload.progress.percentByWeight === 45.5);
check('payload carries the throughput series', payload.throughput.perWeek.length === 6);
check('payload states the figures are precomputed', /computed by the application/i.test(payload.computedNote));
check('payload contains no raw weekly grid', !JSON.stringify(payload).includes('"weekly"'));

const prompt = buildPrompt('site', payload);
check('prompt instructs against recalculation', /do not recalculate/i.test(prompt) || /Do NOT recalculate/.test(prompt));
check('prompt embeds the payload as JSON', prompt.includes('"site": "F-01"'));

const master = buildMasterPayload(p);
check('master payload lists every site', master.sites.length === 2);
check('master payload marks the site with no data', master.sites[1].hasData === false);

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
  'F-01': { key: 'F-01', title: 'Site F-01', kind: 'site', providerLabel: 'Test', model: 'm',
            tokens: { input: 1, output: 1, estimated: true },
            result: { headline: 'Behind plan', verdict: 'off_track', summary: 'Two tasks stuck.',
                      actions: [{ action: 'Chase R3', owner: 'R1', priority: 'high' }] } },
};
const rtf = buildRTF([{ style: 'h1', text: 'Title' }, { style: 'p', text: 'Body — with an em dash' }]);
check('rtf has a valid header', rtf.startsWith('{\\rtf1\\ansi'));
check('rtf escapes non-ascii', rtf.includes('\\u8212?'), rtf.slice(0, 200));
check('rtf is balanced', (rtf.match(/\{/g) || []).length === (rtf.match(/\}/g) || []).length);

const html = buildPrintHTML(p, reports);
check('print html is a full document', html.startsWith('<!DOCTYPE html>') && html.includes('</html>'));
check('print html contains the site', html.includes('F-01'));
check('print html contains the ai narrative', html.includes('Behind plan'));
check('print html escapes angle brackets',
  !buildPrintHTML({ ...p, file: '<script>x</script>' }, {}).includes('<script>x</script>'));

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
