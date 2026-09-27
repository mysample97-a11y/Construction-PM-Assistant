/**
 * Exports: Excel, Word and PDF.
 *
 * Deliberate format choices, so the app carries no extra dependencies:
 *  - Excel  .xlsx via the vendored SheetJS build.
 *  - Word   .rtf — a real Word-openable format that can be written as plain
 *           text. A true .docx needs a zip writer and an OOXML template for no
 *           gain here; Word opens .rtf natively and keeps the formatting.
 *  - PDF    the browser's own print-to-PDF, driven by a purpose-built print
 *           window. It produces better typography than a bundled PDF library
 *           and always matches what the user sees.
 */

import { downloadBlob, fmtDate } from './util.js';
import { chartsToHTML } from './charts.js';
import { PROVENANCE_NOTE, METRIC_DEFINITIONS, RAG_THRESHOLDS } from './report.js';
import { figureHTML, FIGURE_NAMES } from './report-charts.js';

/* ================================ Excel ================================ */

/**
 * Restrict an export to what the user actually analysed.
 *
 * The export used to walk every site in the workbook's register, so analysing
 * one site produced a document covering all twelve — including sites nobody
 * reviewed. Reports are the unit of work: only sites with a generated report
 * (or covered by a generated master analysis) are exported.
 */
export function scopeToReports(portfolio, reports) {
  const all = Object.values(reports || {});
  const codes = new Set();
  for (const r of all) {
    if (r.kind === 'master') (r.computed?.docControl?.sites || []).forEach((c) => codes.add(c));
    else if (r.key) codes.add(r.key);
  }
  const sites = (portfolio?.sites || []).filter((s) => codes.has(s.code));
  // Programme-level risks are computed across the WHOLE register, so one about
  // "5 of 5 sites" would drag unanalysed sites into a single-site export. Keep a
  // programme risk only if every site code it names is in scope.
  const risks = (portfolio?.risks || []).filter((x) => {
    const named = new Set((`${x.title} ${x.detail}`.match(/\b[A-Z]-\d+\b/g) || []));
    if (!named.size) return !!all.some((r) => r.kind === 'master');   // unnamed = portfolio-wide
    return [...named].every((c) => codes.has(c));
  });
  const withData = sites.filter((s) => !s.noData);
  const sum = (f) => withData.reduce((n, x) => n + (Number(f(x)) || 0), 0);
  const live = sum((x) => x.liveCount);
  const done = sum((x) => x.finishedCount);
  const tw = sum((x) => x.totalWeight);
  const dw = sum((x) => x.doneWeight);
  return {
    ...portfolio,
    sites,
    risks,
    siteCount: sites.length,
    sitesWithData: withData.length,
    totalTasks: live,
    totalDone: done,
    pctByCount: live ? Math.round((done / live) * 1000) / 10 : 0,
    pctByWeight: tw ? Math.round((dw / tw) * 1000) / 10 : 0,
    totalStuck: sum((x) => x.stuckCount),
    totalOverdue: withData.reduce((n, x) => n + (x.overdue?.length || 0), 0),
    totalAdditional: withData.reduce((n, x) => n + (x.scopeGrowth?.additionalCount || 0), 0),
    resourceLoad: (portfolio?.resourceLoad || [])
      .map((r) => ({ ...r, sites: (r.sites || []).filter((x) => codes.has(x.code)) }))
      .filter((r) => r.sites.length),
    scopedTo: [...codes],
  };
}

export function buildWorkbook(portfolioAll, reports) {
  if (!window.XLSX) throw new Error('The spreadsheet writer did not load.');
  const portfolio = scopeToReports(portfolioAll, reports);
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();
  const add = (name, aoa, widths) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (widths) ws['!cols'] = widths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  };

  add('Summary', [
    ['BIM MULTI-SITE DELIVERY — ANALYSIS EXPORT'],
    ['Generated', new Date().toISOString().slice(0, 16).replace('T', ' ')],
    ['Source file', portfolio.file || ''],
    [],
    ['Sites analysed', portfolio.siteCount],
    ['Sites with weekly data', portfolio.sitesWithData],
    ['Sites in this export', portfolio.scopedTo.join(', ') || 'none'],
    ['Live tasks', portfolio.totalTasks],
    ['Finished', portfolio.totalDone],
    ['Progress by count (%)', portfolio.pctByCount],
    ['Progress by weight (%)', portfolio.pctByWeight],
    ['Tasks blocked or waiting', portfolio.totalStuck],
    ['Tasks past target week', portfolio.totalOverdue],
    ['Tasks added after kickoff', portfolio.totalAdditional],
    [],
    ['All figures computed by the application from the weekly status grid.'],
  ], [34, 22]);

  add('Sites', [
    ['Site', 'Wave', 'Coordinator', 'Target', 'Tasks live', 'Finished', '% weight',
     'Remaining', 'Rate/wk (recent)', 'Weeks needed', 'Forecast finish', 'Weeks past target',
     'Blocked/waiting', 'Past target week', 'Added scope'],
    ...portfolio.sites.map((s) => [
      s.code, s.wave || '', s.coordinator || '', s.target || '',
      s.noData ? '' : s.liveCount, s.noData ? '' : s.finishedCount,
      s.noData ? '' : s.pctByWeight, s.noData ? '' : s.remaining,
      s.noData ? '' : s.recentVelocity,
      s.noData ? '' : (s.forecastRecent?.weeksNeeded ?? ''),
      s.noData ? '' : (s.forecastRecent?.finishDate || ''),
      s.noData ? '' : (s.slipWeeks ?? ''),
      s.noData ? '' : s.stuckCount, s.noData ? '' : s.overdue.length,
      s.noData ? '' : s.scopeGrowth.additionalCount,
    ]),
  ], [10, 10, 12, 12, 10, 10, 10, 10, 15, 13, 15, 16, 15, 15, 12]);

  const stuck = [['Site', 'Task ID', 'Task', 'Category', 'Status', 'Weeks stuck', 'Since', 'Waiting on', 'Reason', 'Expected clear']];
  const overdue = [['Site', 'Task ID', 'Task', 'Category', 'Target week', 'Weeks late', 'Status', 'Resource']];
  const risks = [['Scope', 'Level', 'Title', 'Detail']];
  const growth = [['Site', 'Task ID', 'Task', 'Category', 'Date added', 'Weight', 'Status']];
  const data = [['Site', 'Issue']];

  for (const r of portfolio.risks) risks.push(['Programme', r.level, r.title, r.detail]);
  for (const s of portfolio.sites) {
    if (s.noData) continue;
    for (const x of s.stuckDetail) stuck.push([s.code, x.taskId, x.name, x.category, x.status, x.weeksStuck, x.sinceWeek || '', x.waitingOn || '', x.reason || '(not logged)', x.expected || '']);
    for (const x of s.overdue) overdue.push([s.code, x.taskId, x.name, x.category, x.targetWeek, x.weeksLate, x.status, x.resource || '']);
    for (const x of s.risks) risks.push([s.code, x.level, x.title, x.detail]);
    for (const x of s.scopeGrowth.items) growth.push([s.code, x.taskId, x.name, x.category, x.added || '', x.weight, x.status || '']);
    for (const x of s.dataIssues) data.push([s.code, x.detail]);
  }
  add('Blocked & waiting', stuck, [8, 11, 34, 22, 12, 12, 10, 20, 46, 14]);
  add('Past target week', overdue, [8, 11, 34, 22, 12, 11, 12, 11]);
  add('Risks', risks, [12, 8, 52, 78]);
  add('Scope growth', growth, [8, 11, 34, 22, 12, 8, 12]);
  add('Data quality', data, [8, 100]);

  add('Resource load', [
    ['Resource', 'Open tasks', 'WIP', 'Stuck', 'Sites', 'Breakdown'],
    ...portfolio.resourceLoad.map((r) => [
      r.resource, r.open, r.wip, r.stuck, r.siteCount,
      r.sites.map((s) => `${s.code}(${s.open})`).join(' '),
    ]),
  ], [14, 12, 8, 8, 8, 40]);

  for (const s of portfolio.sites) {
    if (s.noData) continue;
    add(`${s.code} weekly`, [
      [`${s.code} — weekly throughput`],
      ['Week', 'Week ending', 'Finished that week', 'Cumulative %'],
      ...s.velocity.map((v, i) => [
        v.label, v.end || '', v.count,
        s.curve.points[i]?.actualPct ?? '',
      ]),
    ], [10, 14, 20, 14]);
  }

  const rep = [['Report', 'Section', 'Content']];
  for (const r of Object.values(reports || {})) {
    for (const [k, v] of Object.entries(r.result || {})) {
      rep.push([r.title, k, typeof v === 'string' ? v : JSON.stringify(v)]);
    }
  }
  if (rep.length > 1) add('AI reports', rep, [22, 26, 120]);

  return wb;
}

export function exportExcel(portfolio, reports) {
  const wb = buildWorkbook(portfolio, reports);
  const stamp = new Date().toISOString().slice(0, 10);
  window.XLSX.writeFile(wb, `bim-tracker-analysis-${stamp}.xlsx`, { bookType: 'xlsx' });
}

/* ================================= RTF ================================= */

function rtfEscape(s) {
  return String(s ?? '')
    .replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}')
    // RTF is not Unicode-native; non-ASCII must be emitted as \uN escapes.
    .replace(/[\u0080-\uFFFF]/g, (c) => `\\u${c.charCodeAt(0)}?`)
    .replace(/\n/g, '\\par ');
}

/** @param {Array<{style:string, text:string}>} blocks */
export function buildRTF(blocks) {
  const head = '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\fs20 ';
  const body = blocks.map((b) => {
    const t = rtfEscape(b.text);
    switch (b.style) {
      case 'h1': return `\\pard\\sa200\\b\\fs36 ${t}\\b0\\fs20\\par `;
      case 'h2': return `\\pard\\sa160\\sb160\\b\\fs28 ${t}\\b0\\fs20\\par `;
      case 'h3': return `\\pard\\sa120\\sb120\\b\\fs24 ${t}\\b0\\fs20\\par `;
      case 'meta': return `\\pard\\sa120\\i\\fs18 ${t}\\i0\\fs20\\par `;
      case 'bullet': return `\\pard\\fi-200\\li400\\sa60 \\bullet\\tab ${t}\\par `;
      case 'rule': return '\\pard\\brdrb\\brdrs\\brdrw10\\brsp20\\par ';
      case 'pagebreak': return '\\page ';
      case 'figure': return `\\pard\\sa120\\i [Figure: ${rtfEscape(b.text)} — shown in the app and in the PDF export]\\i0\\par `;
      default: return `\\pard\\sa120 ${t}\\par `;
    }
  }).join('');
  return `${head}${body}}`;
}

/**
 * The report as a flat block list, used by both the Word and PDF exports.
 *
 * Follows the standardised template exactly: fourteen numbered sections, then
 * Appendix A and Appendix B, then the references. Each section carries its
 * provenance so the printed copy says which parts were calculated and which
 * were judged by a model — the same distinction the screen makes.
 *
 * Collapsing a section on screen never affects this: exports are built from the
 * data, never from what happens to be visible.
 */
export function reportBlocks(portfolio, reports) {
  const b = [];
  const all = Object.values(reports || {});
  const master = reports?.__master__ || null;
  const sites = all.filter((r) => r.kind !== 'master');

  b.push({ style: 'h1', text: 'BIM Multi-Site Delivery Tracker' });
  b.push({ style: 'meta', text: `Generated ${new Date().toLocaleString()} · source: ${portfolio?.file || 'workbook'}` });
  b.push({ style: 'meta', text: PROVENANCE_NOTE });
  b.push({ style: 'rule', text: '' });

  if (master) b.push(...oneReport(master));
  for (const r of sites) {
    b.push({ style: 'pagebreak', text: '' });
    b.push(...oneReport(r));
  }

  if (!all.length) {
    // No analysis run means nothing to report. Listing the register here would
    // put sites nobody reviewed into a document someone else will read.
    b.push({ style: 'h2', text: 'No reports generated' });
    b.push({ style: 'p', text: 'No analysis has been run, so there is nothing to export. Select sites in section 4 and generate a report first.' });
  }
  return b;
}

/** One complete report — computed sections and model-written sections together. */
function oneReport(rep) {
  const b = [];
  const c = rep.computed;
  const r = rep.result || {};
  const isMaster = rep.kind === 'master';

  if (!c) {
    b.push({ style: 'h2', text: rep.title || rep.key });
    b.push({ style: 'p', text: 'This report predates the current report structure; re-run it for the full version.' });
    return b;
  }

  const sections = c.sections || [];
  const titleOf = (id) => {
    const s = sections.find((x) => x.id === id);
    return s ? `${s.n}. ${s.title}` : id;
  };
  const badge = (id) => {
    const s = sections.find((x) => x.id === id);
    if (!s) return '';
    return s.by === 'engine' ? '[ENGINE — calculated]'
      : (s.by === 'ai' ? '[AI — written by the model]' : '[AI + ANALYST — check before issuing]');
  };
  const sec = (id) => {
    b.push({ style: 'h2', text: titleOf(id) });
    b.push({ style: 'meta', text: badge(id) });
  };
  // Figures go where they sit on screen, not in a pile at the end.
  const fig = (name) => b.push({ style: 'figure', fig: name, c, risks: r.risks || [], text: FIGURE_NAMES[name] || name });
  const para = (t) => { if (t) b.push({ style: 'p', text: t }); };
  const bullets = (arr, fmt) => {
    for (const x of (Array.isArray(arr) ? arr : [])) b.push({ style: 'bullet', text: typeof x === 'string' ? x : fmt(x) });
  };
  const row = (label, value) => b.push({ style: 'bullet', text: `${label}: ${value}` });

  b.push({ style: 'h1', text: isMaster ? 'Master Analysis Report' : `Site Status Report — ${c.code}` });
  b.push({ style: 'meta', text: `${c.docControl.reportId} · ${c.docControl.reportType} · report date ${c.docControl.reportDate}${c.statusDate ? ` · status date ${c.statusDate}` : ''} · OVERALL ${c.overall}` });

  /* 1 */ sec('introduction');
  para(c.docControl.basis);

  /* 2 */ sec('executive');
  const ex = r.executive || {};
  if (ex.bottomLine) para(`Bottom line: ${ex.bottomLine}`);
  if (ex.keyMessages?.length) { b.push({ style: 'h3', text: 'Key messages' }); bullets(ex.keyMessages, (x) => x); }
  if (ex.decisions?.length) {
    b.push({ style: 'h3', text: 'Decisions required' });
    bullets(ex.decisions, (d) => `${d.id}: ${d.decision} — ${d.owner || 'owner not set'}, by ${d.byWhen || d.neededBy || 'not set'}`);
  }

  /* 3 */ sec('dashboard');
  if (isMaster) {
    row('Sites', `${c.counts.sites} (${c.counts.active} active, ${c.counts.submitted} submitted)`);
    row('Completion (count / weighted)', `${c.kpis.completionCount}% / ${c.kpis.completionWeighted}%`);
    row('Active sites only (weighted)', `${c.kpis.activeCompletionWeighted}%`);
    row('On-time delivery', c.kpis.onTimeDelivery == null ? 'no submitted sites' : `${c.kpis.onTimeDelivery}%`);
    row('Back-end concentration (C110–C150)', `${c.kpis.backEndConcentration}% of open tasks`);
    b.push({ style: 'h3', text: 'RAG by site and dimension' });
    for (const m of c.ragMatrix) {
      b.push({ style: 'bullet', text: `${m.code}: schedule ${m.schedule}, scope ${m.scope}, constraints ${m.constraints}, resources ${m.resources}, quality ${m.quality}, data ${m.dataIntegrity} — overall ${m.overall}` });
    }
    fig('ragMatrix');
  } else {
    fig('kpiDashboard');
    for (const k of c.kpis) b.push({ style: 'bullet', text: `${k.k}: ${k.v ?? '—'} (target ${k.target})${k.rag ? ` — ${k.rag}` : ''}` });
    b.push({ style: 'h3', text: 'Status by dimension' });
    for (const [k, v] of Object.entries(c.dimensions)) {
      if (k === 'overall') continue;
      b.push({ style: 'bullet', text: `${k}: ${v.rag} — ${v.basis}` });
    }
  }

  /* 4 */ sec('schedule');
  if (isMaster) {
    fig('scheduleAcrossSites');
    fig('completionVsTime');
    for (const s of c.schedule) {
      b.push({ style: 'bullet', text: `${s.code}: target ${s.target || '—'}, ${s.submitted ? 'submitted' : 'forecast'} ${s.forecast || '—'}${s.variance == null ? '' : `, variance ${s.variance > 0 ? '+' : ''}${s.variance} d`}, ${s.completionWeighted ?? '—'}% complete against ${s.timeElapsed ?? '—'}% of time — ${s.overall}` });
    }
  } else {
    fig('milestoneTimeline');
    b.push({ style: 'h3', text: 'Milestones' });
    for (const m of c.milestones) {
      b.push({ style: 'bullet', text: `${m.id} ${m.name}: baseline ${m.baseline || '—'}, actual/forecast ${m.actual || '—'}${m.variance == null ? '' : `, variance ${m.variance > 0 ? '+' : ''}${m.variance} d`} — ${m.status} (${m.rag})` });
    }
    b.push({ style: 'h3', text: 'Forecast' });
    row('Method', c.metrics.forecastMethod);
    row('Earliest submission', c.metrics.forecastDate || 'not forecastable');
    if (c.metrics.indicativeDays != null) row('Remaining template durations', `${c.metrics.indicativeDays} working days`);
    b.push({ style: 'h3', text: 'Week-on-week movement' });
    if (c.movement.comparable) {
      row('Weeks compared', `${c.movement.fromWeek} to ${c.movement.toWeek}`);
      row('Finished this week', String(c.movement.finishedThisWeek));
      bullets(c.movement.changes.slice(0, 20), (x) => `${x.id} ${x.name}: ${x.from || 'blank'} to ${x.to || 'blank'}`);
    } else para(`Not comparable: ${c.movement.reason}.`);
  }
  para(r.notes?.schedule);

  /* 5 */ if (isMaster) {
    sec('bottlenecks');
    fig('categoryBySite');
    for (const x of c.categories.filter((y) => y.applicable > 0)) {
      b.push({ style: 'bullet', text: `${x.id} ${x.name}: ${x.finished}/${x.applicable} finished, ${x.open} open${x.stuckSites.length ? `, stuck on ${x.stuckSites.join(', ')}` : ''}` });
    }
    para(r.notes?.bottlenecks);
  } else {
    sec('wbs');
    fig('categoryProgress');
    row('Roll-up discrepancies', String(c.wbs.discrepancies));
    for (const x of c.wbs.rows) {
      b.push({ style: 'bullet', text: `${x.id} ${x.name}: ${x.doneCount}/${x.liveCount} (${x.liveCount ? `${x.computedPct}%` : 'N/A'}), reported ${x.selfStatus || '—'} — ${x.consistent ? 'consistent' : x.check}` });
    }
    para(r.notes?.wbs);
  }

  /* 6 */ sec('scope');
  if (isMaster) {
    for (const s of c.scope) b.push({ style: 'bullet', text: `${s.code}: ${s.additional} added (${s.setNA} later N/A), ${s.openItems} unconfirmed, source ${s.source}` });
  } else {
    row('Source', c.scope.source);
    row('Added after kickoff', `${c.scope.additionalCount} (${c.scope.growthPct}% growth)`);
    for (const x of c.scope.items) b.push({ style: 'bullet', text: `${x.ref} ${x.item} — ${x.type}, ${x.status}${x.raised ? `, raised ${x.raised}` : ''}` });
    for (const x of c.scope.additional) b.push({ style: 'bullet', text: `${x.taskId} ${x.name}${x.added ? ` — added ${x.added}` : ''}` });
  }
  para(r.notes?.scope);

  /* 7 */ sec('constraints');
  if (isMaster) {
    fig('partyDependency');
    for (const x of c.parties) b.push({ style: 'bullet', text: `${x.party}: ${x.count} open across ${x.sites.join(', ')}` });
  } else {
    for (const x of c.constraints.rows) {
      b.push({ style: 'bullet', text: `${x.id} ${x.name} — from ${x.provider}${x.byWeek == null ? '' : `, required W${x.byWeek}`}, ${x.status} (${x.rag})${x.notTracked ? ' [not tracked on this site]' : ''}` });
    }
    if (c.constraints.noReceivedDates) para('No received dates are recorded for any prerequisite.');
  }
  para(r.notes?.constraints);

  /* 8 — site only */
  if (!isMaster) {
    sec('log');
    if (c.log.open.length) fig('constraintAgeing');
    if (c.log.rows.length) {
      for (const l of c.log.rows) {
        b.push({ style: 'bullet', text: `${l.id} (${l.kind}, ${l.control}): ${l.reason} — waiting on ${l.waitingOn || '—'}, raised ${l.raised || '—'}, age ${l.age == null ? 'undated' : `${l.age} d`} (${l.rag})` });
      }
    } else para('No entries recorded.');
    para(r.notes?.log);
  }

  /* 9 */ sec('risks');
  if (r.risks?.length) {
    fig('riskHeatMap');
    for (const x of r.risks) {
      b.push({ style: 'bullet', text: `${x.id} [${x.rating} ${x.score}] ${x.risk} — P${x.probability} x I${x.impact}; ${x.strategy}: ${x.response} (${x.owner})${x.sitesAffected ? ` [${x.sitesAffected.join(', ')}]` : ''}` });
    }
    para('Probability and impact are the model\'s judgement. Score and rating band are computed by the app: RED 15+, AMBER 8-14, GREEN 7 or less.');
  } else para('No risks were proposed.');

  /* 10 */ sec('resources');
  if (isMaster) {
    fig('resourceAcrossSites');
    for (const x of c.resourceLoad) b.push({ style: 'bullet', text: `${x.resource}: ${x.open} open across ${x.sites.join(', ')} — ${x.activeSites} active, load ${x.load}` });
    para('Loading reflects the selected sites only.');
  } else {
    for (const x of c.resources.rows) b.push({ style: 'bullet', text: `${x.resource}: ${x.open} open, ${x.wip} in progress, ${x.stuck} stuck` });
    if (c.resources.unassigned?.open) para(`${c.resources.unassigned.open} open tasks have no responsible resource.`);
    for (const x of c.resources.externalParties) b.push({ style: 'bullet', text: `Waiting on ${x.party}: ${x.items.join(', ')}` });
  }
  para(r.notes?.resources);

  /* 11 */ sec('quality');
  if (isMaster) {
    fig('findingsBySite');
    for (const x of c.patterns) b.push({ style: 'bullet', text: `${x.pattern} — ${x.count} site(s): ${x.sites.join(', ')}${x.systemic ? ' [SYSTEMIC]' : ''}. ${x.correction}` });
  } else if (c.quality.findings.length) {
    for (const f of c.quality.findings) b.push({ style: 'bullet', text: `${f.id} ${f.finding} — ${f.evidence} Impact: ${f.impact} Correction: ${f.correction}` });
  } else para('No findings.');
  para(r.notes?.quality);

  /* 11b master prioritisation */
  if (isMaster) {
    sec('priority');
    if (c.ipi.length) fig('ipi');
    if (c.ipi.length) {
      c.ipi.forEach((x, i) => b.push({ style: 'bullet', text: `${i + 1}. ${x.code} — IPI ${x.ipi} (RAG ${x.components.rag}, priority ${x.components.priority}, time ${x.components.time}, gap ${x.components.gap}); ${x.overall}, ${x.daysToTarget == null ? 'no target' : `${x.daysToTarget} days to target`}` }));
      bullets(r.interventions, (x) => `${x.site}: ${x.recommendedIntervention}`);
    } else para('No active sites; prioritisation does not apply.');
    para(r.notes?.priority);
  }

  /* 12 */ sec('actions');
  if (r.actions?.length) {
    for (const x of r.actions) b.push({ style: 'bullet', text: `${x.id} [${x.priority}] ${x.action} — ${x.owner}, due ${x.due}${x.links ? ` (${x.links})` : ''}` });
  } else para('No actions were raised.');
  if (r.lookahead?.length) {
    b.push({ style: 'h3', text: 'Two-week lookahead' });
    bullets(r.lookahead, (x) => `${x.week}: ${x.focus}`);
  }

  /* 13 */ sec('conclusion');
  para(r.conclusion || 'No conclusion was returned.');

  /* 14 */ sec('references');
  (c.references || []).forEach(([ref, used], i) => b.push({ style: 'bullet', text: `[${i + 1}] ${ref} — Used for: ${used}` }));

  /* Appendices */
  b.push({ style: 'h2', text: 'Appendix A. Metric definitions and RAG thresholds' });
  b.push({ style: 'meta', text: '[ENGINE — calculated]' });
  for (const [k, v] of METRIC_DEFINITIONS) b.push({ style: 'bullet', text: `${k}: ${v}` });
  b.push({ style: 'h3', text: 'RAG thresholds' });
  for (const [d, g, am, rd] of RAG_THRESHOLDS) b.push({ style: 'bullet', text: `${d} — GREEN: ${g}; AMBER: ${am}; RED: ${rd}` });

  b.push({ style: 'h2', text: isMaster ? 'Appendix B. Site summary register' : 'Appendix B. Task-level status register' });
  b.push({ style: 'meta', text: '[ENGINE — calculated]' });
  if (isMaster) {
    for (const x of c.appendixB) b.push({ style: 'bullet', text: `${x.code} (${x.priority}) — ${x.overall}, resources ${x.resources.join(', ') || '—'}, report ${x.reportId}` });
  } else {
    const weeks = (c.weeks || []).map((w) => w.label);
    for (const t of c.appendixB) {
      b.push({ style: 'bullet', text: `${t.id} ${t.name} [${t.discipline}, wt ${t.weight}, ${t.type}] — ${(t.weekly || []).map((s, i) => `${weeks[i] || `W${i + 1}`}:${s || '—'}`).join(' ')}` });
    }
  }
  return b;
}

export function exportWord(portfolio, reports) {
  const rtf = buildRTF(reportBlocks(portfolio, reports));
  const stamp = new Date().toISOString().slice(0, 10);
  downloadBlob(new Blob([rtf], { type: 'application/rtf' }), `bim-tracker-report-${stamp}.rtf`);
}

/* ================================= PDF ================================= */

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/**
 * The printable report. Charts are serialised straight out of the same
 * renderers the screen uses, so what gets printed is what was reviewed rather
 * than a table of numbers standing in for the visuals.
 *
 * `withCharts` is optional because the function is also called from the test
 * harness, where there is no layout engine to render SVG into.
 */
/*
 * Colours for printing.
 *
 * The figures are drawn with the app's CSS variables. On screen those resolve
 * against the page stylesheet; in the separate print window they only resolve
 * if every variable is declared again there — and an undeclared colour in SVG
 * silently falls back to BLACK. That is how bar tracks printed as black blocks
 * and legend swatches printed blank.
 *
 * So every var(--x) is replaced with its literal value at the moment the figure
 * is serialised. The printed file then depends on no variables at all, and a
 * future figure using a new colour cannot break printing.
 */
const PRINT_PALETTE = {
  '--paper': '#EFEDFB', '--paper-deep': '#E3DFF8', '--sheet': '#FFFFFF', '--sheet-alt': '#F7F5FE',
  '--block': '#241E58', '--block-2': '#332B76', '--block-3': '#1A1544',
  '--ink': '#1E1B36', '--ink-2': '#4A4570', '--ink-3': '#7C769D', '--ink-inv': '#EAE7FB', '--ink-inv-2': '#ABA3DB',
  '--rule': '#D9D4F3', '--rule-soft': '#EAE7FA', '--rule-hard': '#B5ACE9', '--rule-inv': '#3C3486',
  '--blueprint': '#5B4FE9', '--blueprint-lo': '#EDEBFD', '--blueprint-hi': '#4A3FD6',
  '--sign': '#0E9F6E', '--sign-lo': '#E2F7EF', '--hivis': '#C2740A', '--hivis-lo': '#FCF2DF',
  '--survey': '#DC2626', '--survey-lo': '#FDEAEA', '--conc': '#7C769D', '--conc-lo': '#ECEAF7',
  '--plum': '#8B5CF6', '--plum-lo': '#F2EDFE',
  '--st-not': '#C3BEDF', '--st-wip': '#5B4FE9', '--st-blocked': '#DC2626',
  '--st-waiting': '#C2740A', '--st-done': '#0E9F6E', '--st-na': '#E6E3F2',
  '--font-ui': 'Arial, Helvetica, sans-serif', '--font-data': 'Consolas, monospace',
};

export function resolveCssVars(html) {
  let live = null;
  try {
    if (typeof getComputedStyle === 'function' && typeof document !== 'undefined') {
      live = getComputedStyle(document.documentElement);
    }
  } catch { live = null; }
  const valueOf = (name) => {
    const v = live ? live.getPropertyValue(name).trim() : '';
    return v || PRINT_PALETTE[name] || null;
  };
  // Also handles the fallback form var(--x, #fff).
  return String(html).replace(/var\((--[a-zA-Z0-9-]+)\s*(?:,\s*([^)]+))?\)/g, (m, name, fb) => {
    const v = valueOf(name);
    if (v) return v;
    if (fb) return fb.trim();
    return '#7C769D';          // never black: a neutral grey is the safe failure
  });
}

export function buildPrintHTML(portfolio, reports, withCharts = true) {
  const blocks = reportBlocks(portfolio, reports);
  const body = blocks.map((b) => {
    const t = esc(b.text);
    switch (b.style) {
      case 'h1': return `<h1>${t}</h1>`;
      case 'h2': return `<h2>${t}</h2>`;
      case 'h3': return `<h3>${t}</h3>`;
      case 'meta': return `<p class="meta">${t}</p>`;
      case 'bullet': return `<li>${t}</li>`;
      case 'rule': return '<hr>';
      case 'pagebreak': return '<div class="pagebreak"></div>';
      case 'figure': {
        const html = withCharts ? resolveCssVars(figureHTML(b.fig, b.c, b.risks)) : '';
        return html
          ? `<div class="figure">${html}</div>`
          : `<p class="meta">[Figure: ${t} — not rendered in this export]</p>`;
      }
      default: return `<p>${t}</p>`;
    }
  }).join('\n')
    // Wrap runs of <li> so bullets render as real lists.
    .replace(/(<li>[\s\S]*?<\/li>)(?!\s*<li>)/g, (m) => `<ul>${m}</ul>`)
    .replace(/<\/ul>\s*<ul>/g, '');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>BIM Multi-Site Delivery — Analysis Report</title>
<style>
  @page { margin: 16mm 14mm; }
  /* The charts are serialised SVG that refers to the app's CSS variables, so
     the print document has to declare them or every shape renders black. */
  :root {
    --ink: #1E1B36; --ink-2: #4A4570; --ink-3: #7C769D;
    --sheet: #FFFFFF; --sheet-alt: #F7F5FE;
    --rule: #D9D4F3; --rule-soft: #EAE7FA; --rule-hard: #B5ACE9;
    --blueprint: #5B4FE9; --blueprint-lo: #EDEBFD;
    --sign: #0E9F6E; --hivis: #C2740A; --survey: #DC2626;
    --conc: #7C769D; --plum: #8B5CF6;
    --st-not: #C3BEDF; --st-wip: #5B4FE9; --st-blocked: #DC2626;
    --st-waiting: #C2740A; --st-done: #0E9F6E; --st-na: #E6E3F2;
    --font-ui: Arial, Helvetica, sans-serif;
    --font-data: Consolas, monospace;
  }
  body { font: 11pt/1.5 Arial, Helvetica, sans-serif; color: #12212F; max-width: 175mm; }
  h1 { font-size: 19pt; margin: 0 0 4pt; }
  h2 { font-size: 14pt; margin: 18pt 0 6pt; border-bottom: 1px solid #C6D0DA; padding-bottom: 3pt; page-break-after: avoid; }
  h3 { font-size: 11.5pt; margin: 12pt 0 4pt; page-break-after: avoid; }
  p { margin: 0 0 7pt; }
  .meta { font-size: 9pt; color: #5A6B7B; font-style: italic; }
  ul { margin: 0 0 8pt; padding-left: 16pt; }
  li { margin-bottom: 3pt; page-break-inside: avoid; }
  hr { border: 0; border-top: 1px solid #C6D0DA; margin: 14pt 0; }
  .pagebreak { page-break-before: always; }
  .chartwrap { margin: 0 0 14pt; page-break-inside: avoid; }
  .charttitle { font-size: 10.5pt; font-weight: 600; margin-bottom: 1pt; }
  .chartnote { font-size: 8.5pt; color: #5A6B7B; margin-bottom: 4pt; }
  .chartbox svg { max-width: 100%; height: auto; }
  .legend { display: flex; flex-wrap: wrap; gap: 10pt; font-size: 8pt; color: #4A4570; margin-top: 3pt; }
  .legend span { display: inline-flex; align-items: center; gap: 3pt; }
  .swatch { width: 8pt; height: 8pt; border-radius: 1pt; display: inline-block; }
  .meter { position: relative; height: 6pt; background: #E3DFF8; border-radius: 3pt; overflow: hidden; }
  .meter__fill { position: absolute; inset: 0 auto 0 0; background: #5B4FE9; border-radius: 3pt; }
  .row { display: flex; align-items: center; }
  .grow { flex: 1; }
</style></head><body>${body}
<p class="meta">Site and resource codes are pseudonyms held only in the author's private reference file.</p>
</body></html>`;
}

/**
 * Opens a print window. Returns false when the browser blocked the popup, so
 * the caller can say so rather than appearing to do nothing.
 */
export function exportPDF(portfolio, reports) {
  const html = buildPrintHTML(portfolio, reports);
  const w = window.open('', '_blank');
  if (!w) return false;
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(() => { try { w.print(); } catch { /* user can print manually */ } }, 350);
  return true;
}
