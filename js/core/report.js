/**
 * Report engine — the deterministic half of the standardised report.
 *
 * This layer sits on top of engine.js and computes everything the report
 * template defines in its Appendix A: metric formulas, RAG thresholds,
 * milestones, validation findings, and the portfolio aggregation and
 * Intervention Priority Index.
 *
 * DIVISION OF LABOUR — this is the rule the whole tool rests on:
 *   ENGINE   anything that follows from the workbook by arithmetic or a stated
 *            rule. Computed here, never asked of a model.
 *   AI       judgement that cannot be derived: what a pattern means, how to
 *            word a decision, which risk matters most. Always labelled in the
 *            report so the reader knows a model wrote it.
 *   ANALYST  items the template marks as needing human approval before the
 *            report is issued (risks, actions, decisions).
 *
 * Every section below declares its provenance, and the UI and exports render
 * that badge rather than assuming.
 */

import {
  toISO, todayISO, addDays, diffDays, parseISO, clamp, sum, groupBy, norm,
  addWorkingDays,
} from './util.js';
import { STATUS, STUCK, screenText } from './parser.js';
import { statusAt } from './engine.js';

/* ============================================================
   Provenance
   ============================================================ */

export const ENGINE = 'engine';
export const AI = 'ai';
export const AI_ANALYST = 'ai+analyst';

/** Section list for the single-site report, in template order. */
export const SITE_SECTIONS = [
  { n: 1,  id: 'introduction',   title: 'Introduction',                                      by: ENGINE },
  { n: 2,  id: 'executive',      title: 'Executive Summary',                                 by: AI_ANALYST },
  { n: 3,  id: 'dashboard',      title: 'Performance Dashboard',                             by: ENGINE },
  { n: 4,  id: 'schedule',       title: 'Schedule Performance',                              by: ENGINE },
  { n: 5,  id: 'wbs',            title: 'Progress by Work Breakdown',                        by: ENGINE },
  { n: 6,  id: 'scope',          title: 'Scope and Change Control',                          by: ENGINE },
  { n: 7,  id: 'constraints',    title: 'Constraints and Prerequisites',                     by: ENGINE },
  { n: 8,  id: 'log',            title: 'Issues and Waiting-On / Blocked Log',               by: ENGINE },
  { n: 9,  id: 'risks',          title: 'Risk Register',                                     by: AI_ANALYST },
  { n: 10, id: 'resources',      title: 'Resources and Responsibilities',                    by: ENGINE },
  { n: 11, id: 'quality',        title: 'Quality, Information Management and Data Integrity', by: ENGINE },
  { n: 12, id: 'actions',        title: 'Actions and Two-Week Lookahead',                    by: AI_ANALYST },
  { n: 13, id: 'conclusion',     title: 'Conclusion',                                        by: AI },
  { n: 14, id: 'references',     title: 'References',                                        by: ENGINE },
];

export const MASTER_SECTIONS = [
  { n: 1,  id: 'introduction', title: 'Introduction',                                 by: ENGINE },
  { n: 2,  id: 'executive',    title: 'Executive Summary',                            by: AI_ANALYST },
  { n: 3,  id: 'dashboard',    title: 'Portfolio Dashboard',                          by: ENGINE },
  { n: 4,  id: 'schedule',     title: 'Schedule Performance Across Sites',            by: ENGINE },
  { n: 5,  id: 'bottlenecks',  title: 'Progress and Bottlenecks Across Sites',        by: ENGINE },
  { n: 6,  id: 'scope',        title: 'Scope and Change Across Sites',                by: ENGINE },
  { n: 7,  id: 'constraints',  title: 'Constraints and External Dependencies',        by: ENGINE },
  { n: 8,  id: 'resources',    title: 'Shared Resources and Capacity',                by: ENGINE },
  { n: 9,  id: 'risks',        title: 'Portfolio Risk Register',                      by: AI_ANALYST },
  { n: 10, id: 'quality',      title: 'Systemic Quality and Data-Integrity Patterns', by: ENGINE },
  { n: 11, id: 'priority',     title: 'Prioritisation and Decisions',                 by: ENGINE },
  { n: 12, id: 'actions',      title: 'Portfolio Actions and Two-Week Lookahead',     by: AI_ANALYST },
  { n: 13, id: 'conclusion',   title: 'Conclusion',                                   by: AI },
  { n: 14, id: 'references',   title: 'References',                                   by: ENGINE },
];

export function provenanceOf(kind, id) {
  const list = kind === 'master' ? MASTER_SECTIONS : SITE_SECTIONS;
  return list.find((s) => s.id === id)?.by || ENGINE;
}

export const PROVENANCE_NOTE =
  'Sections marked ENGINE are calculated by this app from your workbook — the same input always gives the same figures. '
  + 'Sections marked AI are written by the language model you selected, reading only those calculated figures; it never recalculates them. '
  + 'Sections marked AI + ANALYST are drafted by the model and are meant to be checked and approved by you before the report is issued.';

/* ============================================================
   RAG
   ============================================================ */

export const RED = 'RED';
export const AMBER = 'AMBER';
export const GREEN = 'GREEN';
const rank = { [GREEN]: 0, [AMBER]: 1, [RED]: 2 };
const worst = (...v) => v.filter(Boolean).sort((a, b) => rank[b] - rank[a])[0] || GREEN;

/* ============================================================
   Small helpers
   ============================================================ */

const pct = (n, d) => (d > 0 ? (n / d) * 100 : null);
const r1 = (n) => (n == null ? null : Math.round(n * 10) / 10);
const r2 = (n) => (n == null ? null : Math.round(n * 100) / 100);

/** ISO week number, for the report ID. */
function isoWeek(iso) {
  const d = parseISO(iso);
  if (!d) return null;
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const start = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - start) / 86400000 + 1) / 7);
}

/** A site counts as submitted when its register records an actual submission
 *  or its submission category is finished. Several rules switch on this. */
function isSubmitted(site, a) {
  if (site.actual) return true;
  const sub = (a.categories || []).find((c) => /^C150/i.test(c.id));
  return !!sub && sub.computedPct === 100 && sub.liveCount > 0;
}

/* ============================================================
   Site report
   ============================================================ */

/**
 * @param {object} site      register row + parsed detail
 * @param {object} a         analyseSite() output for the same site
 * @param {object} template  parsed Task Template
 * @param {object} opts      { reportDate, previousAnalysis }
 */
export function computeSiteReport(site, a, template, opts = {}) {
  const reportDate = opts.reportDate || todayISO();
  const submitted = isSubmitted(site, a);
  const statusDate = a.reportingWeek?.end || null;

  /* ---------- 1. metrics (Appendix A.1) ---------- */

  // Time elapsed = (status date - start) / (target - start), capped at 100%.
  let timeElapsed = null;
  if (statusDate && site.start && site.target) {
    const span = diffDays(site.start, site.target);
    const used = diffDays(site.start, statusDate);
    if (span && span > 0 && used != null) timeElapsed = clamp((used / span) * 100, 0, 100);
  }

  const completionCount = a.noData ? null : pct(a.finishedCount, a.liveCount);
  const completionWeighted = a.noData ? null : pct(a.doneWeight, a.totalWeight);

  // Time-based SPI proxy. Not an EVM SPI: it assumes planned progress is
  // linear in time, which is an approximation and is labelled as one.
  // Not computed for a submitted site — there is no remaining time to be
  // behind against.
  const spiProxy = (!submitted && timeElapsed > 0 && completionWeighted != null)
    ? completionWeighted / timeElapsed : null;

  const throughput = a.velocity?.length ? a.velocity[a.velocity.length - 1].count : null;

  /* ---------- indicative forecast (Appendix A.1) ----------
     Used when throughput is zero or unmeasurable: sum the template's typical
     durations for the categories that still have open work, and run them from
     the status date on the working calendar. It is a floor, not a plan. */
  const catById = new Map((template?.categories || []).map((c) => [c.id, c]));
  const openCategories = (a.categories || [])
    .filter((c) => c.liveCount > 0 && c.doneCount < c.liveCount)
    .sort((x, y) => (catById.get(x.id)?.order ?? 9999) - (catById.get(y.id)?.order ?? 9999));
  const remainingDurations = openCategories.map((c) => ({
    id: c.id, name: c.name, days: catById.get(c.id)?.duration ?? null,
  }));
  const knownDays = remainingDurations.filter((x) => x.days != null);
  const indicativeDays = knownDays.length ? sum(knownDays, (x) => x.days) : null;
  const cal = opts.calendar || { workingDays: [1, 2, 3, 4, 5], holidays: [] };
  const indicativeDate = (indicativeDays != null && statusDate)
    ? addWorkingDays(statusDate, indicativeDays, cal) : null;

  // Throughput forecast when work is actually completing, otherwise indicative.
  const throughputForecast = (!submitted && a.forecastRecent?.finishDate) ? a.forecastRecent.finishDate : null;
  const forecastDate = submitted ? (site.actual || null) : (throughputForecast || indicativeDate);
  const forecastMethod = submitted ? 'actual'
    : (throughputForecast ? 'throughput' : (indicativeDate ? 'indicative (template durations)' : 'none'));

  // Schedule variance in calendar days. Positive = late.
  const scheduleVariance = (site.target && forecastDate) ? diffDays(site.target, forecastDate) : null;

  /* ---------- roll-up check (Appendix A.1) ----------
     The coordinator's category row against the arithmetic of its tasks. A
     mismatch is not automatically an error — it is a question to answer. */
  const rollup = (a.categories || []).map((c) => {
    const implied = c.liveCount === 0 ? STATUS.NA
      : (c.doneCount === c.liveCount ? STATUS.FINISHED
        : (c.doneCount > 0 || c.stuck > 0 ? STATUS.WIP : STATUS.NOT_STARTED));
    const hidesStuck = c.stuck > 0 && c.selfStatus && !STUCK.includes(c.selfStatus);
    const mismatch = !!c.selfStatus && c.selfStatus !== implied;
    return {
      ...c, implied, hidesStuck,
      check: mismatch || hidesStuck
        ? (hidesStuck ? 'Category hides a blocked or waiting task' : `Reported ${c.selfStatus}, tasks imply ${implied}`)
        : 'Consistent',
      consistent: !(mismatch || hidesStuck),
    };
  });
  const rollupDiscrepancies = rollup.filter((c) => !c.consistent).length;

  /* ---------- 4. milestones ----------
     Derived by fixed rules so two runs of the same workbook always agree:
       M1 site start (register)
       M2 scope freeze, only when scope items are still open
       M3 the latest in-progress category that still has later work waiting
       M4 submission (the C150 task) */
  const milestones = [];
  if (site.start) {
    milestones.push({
      id: 'M1', name: 'Site start', baseline: site.start, actual: site.start,
      variance: 0, status: 'Achieved', rag: GREEN,
    });
  }
  const scope = computeScope(site, a, reportDate);
  if (scope.openItems.length) {
    milestones.push({
      id: 'M2', name: `Scope freeze (${scope.openItems.length} item${scope.openItems.length === 1 ? '' : 's'} to confirm)`,
      baseline: null, actual: 'Pending', variance: null, status: 'Open', rag: RED,
    });
  }
  const gating = [...openCategories].reverse().find((c) => {
    const order = catById.get(c.id)?.order ?? 0;
    return (a.categories || []).some((o) => (catById.get(o.id)?.order ?? 0) > order && o.liveCount > 0 && o.doneCount < o.liveCount);
  });
  if (gating) {
    milestones.push({
      id: 'M3', name: `${gating.name} closed (${gating.id})`, baseline: null,
      actual: 'Not dated', variance: null, status: gating.doneCount > 0 ? 'WIP' : 'Not started', rag: AMBER,
    });
  }
  const subTask = (site.detail?.tasks || []).find((t) => /^T150/i.test(t.id));
  milestones.push({
    id: 'M4', name: `Submission${subTask ? ` (${subTask.id})` : ''}`,
    baseline: site.target || null,
    actual: submitted ? (site.actual || 'Submitted') : (forecastDate ? `${forecastDate} (earliest)` : 'Not forecast'),
    variance: scheduleVariance,
    status: submitted ? 'Achieved' : (subTask ? statusAt(subTask, 9999) || 'Not started' : 'Not started'),
    rag: submitted ? (scheduleVariance > 0 ? AMBER : GREEN) : (scheduleVariance > 0 ? RED : AMBER),
  });

  /* ---------- 4.3 week-on-week movement ---------- */
  const movement = computeMovement(site, a);

  /* ---------- 7/8 constraints and log ---------- */
  const constraints = computeConstraints(site, a, template, reportDate);
  const logEntries = computeLog(site, a, reportDate);

  /* ---------- 10 resources ---------- */
  const resources = computeResources(site, a);

  /* ---------- 11 data integrity ---------- */
  const quality = computeQuality(site, a, rollup, logEntries, constraints);

  /* ---------- 3.2 RAG by dimension (Appendix A.2) ---------- */
  const dimensions = computeRag({
    submitted, spiProxy, scheduleVariance, scope, constraints, logEntries,
    resources, quality, a, milestones,
  });
  const overall = dimensions.overall;

  /* ---------- 3.1 KPI table ---------- */
  const kpis = [
    { k: 'Completion (count)', v: completionCount == null ? null : `${r1(completionCount)}% (${a.finishedCount}/${a.liveCount})`,
      target: timeElapsed == null ? '—' : `≈ ${Math.round(timeElapsed)}% at this date`,
      rag: ragCompletion(completionCount, timeElapsed, submitted) },
    { k: 'Completion (weighted)', v: completionWeighted == null ? null : `${r1(completionWeighted)}%`,
      target: timeElapsed == null ? '—' : `≈ ${Math.round(timeElapsed)}% at this date`,
      rag: ragCompletion(completionWeighted, timeElapsed, submitted) },
    { k: 'Time elapsed', v: timeElapsed == null ? null : `${r1(timeElapsed)}%${site.start && statusDate ? ` (${diffDays(site.start, statusDate)} of ${diffDays(site.start, site.target)} days)` : ''}`,
      target: '—', rag: null },
    { k: 'Time-based SPI (proxy)', v: spiProxy == null ? (submitted ? 'not applicable (submitted)' : null) : String(r2(spiProxy)),
      target: '≥ 0.95', rag: dimensions.schedule.rag },
    { k: 'Weekly throughput', v: throughput == null ? null : `${throughput} finished`,
      target: '> 0', rag: throughput === 0 ? RED : (throughput == null ? null : GREEN) },
    { k: 'Work in progress', v: a.noData ? null : `${a.wipCount} task${a.wipCount === 1 ? '' : 's'}${a.liveCount ? ` (${Math.round(pct(a.wipCount, a.liveCount))}% of applicable)` : ''}`,
      target: 'Limit to critical path', rag: dimensions.resources.rag },
    { k: 'Unconfirmed scope items', v: `${scope.openItems.length} of ${scope.items.length || scope.openItems.length}`,
      target: '0 before production', rag: scope.openItems.length ? RED : GREEN },
    { k: 'Open waiting-on / blocked items', v: `${logEntries.open.length} logged${logEntries.oldestAge != null ? `, oldest ${logEntries.oldestAge} days` : ''}`,
      target: '0 or ≤ 5 days', rag: dimensions.constraints.rag },
    { k: 'Category roll-up discrepancies', v: String(rollupDiscrepancies), target: '0', rag: rollupDiscrepancies ? AMBER : GREEN },
    { k: 'Data-integrity findings', v: `${quality.findings.length}${quality.privacy ? ' (incl. 1 privacy)' : ''}`,
      target: '0', rag: dimensions.dataIntegrity.rag },
  ];

  /* ---------- document control ---------- */
  const week = isoWeek(reportDate);
  const docControl = {
    reportId: `BIM-MSDT-SSR-${site.code}-${String(reportDate).slice(0, 4)}-W${week}`,
    reportType: 'Site status (highlight) report',
    siteCode: site.code,
    priority: site.priority || '—',
    statusDate, reportDate,
    basis: site.detail?.sheetName ? `Sites register row and site sheet "${site.detail.sheetName}"` : 'Sites register row',
    preparedBy: 'BIM Multi-Site Delivery Tracker (engine) with AI interpretation',
  };

  return {
    kind: 'site',
    code: site.code,
    docControl,
    overall,
    submitted,
    statusDate,
    reportDate,
    metrics: {
      statusDate,
      applicableTasks: a.liveCount ?? null,
      totalTasks: a.taskCount ?? null,
      naTasks: a.naCount ?? null,
      finished: a.finishedCount ?? null,
      wip: a.wipCount ?? null,
      notStarted: a.notStartedCount ?? null,
      stuck: a.stuckCount ?? null,
      completionCount: r1(completionCount),
      completionWeighted: r1(completionWeighted),
      timeElapsed: r1(timeElapsed),
      spiProxy: r2(spiProxy),
      throughput,
      scheduleVariance,
      forecastDate,
      forecastMethod,
      indicativeDays,
      remainingDurations,
      weightsDefaulted: (site.detail?.tasks || []).filter((t) => t.weight == null || t.weight === '').length,
    },
    kpis,
    dimensions,
    milestones,
    movement,
    wbs: { rows: rollup, discrepancies: rollupDiscrepancies },
    scope,
    constraints,
    log: logEntries,
    resources,
    quality,
    appendixB: (site.detail?.tasks || []).map((t) => ({
      id: t.id, name: t.name, discipline: t.discipline, weight: t.weight ?? 1,
      type: t.type, weekly: t.weekly, resource: t.resource,
    })),
    weeks: site.detail?.weeks || [],
    references: SITE_REFERENCES,
    sections: SITE_SECTIONS,
  };
}

/* ---------------- scope (section 6) ---------------- */

function computeScope(site, a, reportDate) {
  // Preferred source: an explicit Scope table on the site sheet. Falls back to
  // the Type / Date added columns, exactly as the template's inclusion rules say.
  const table = site.detail?.scope || [];
  const items = table.length ? table.map((s, i) => ({
    ref: s.ref || `S${i + 1}`,
    item: s.item || s.name || '',
    type: s.type || 'Baseline',
    raised: s.raised || s.added || null,
    status: s.status || 'Confirmed',
    comment: s.comment || '',
    open: /open|to confirm|unconfirmed|pending/i.test(`${s.status || ''} ${s.type || ''}`),
  })) : [];

  const additional = (a.scopeGrowth?.items || []);
  const regular = a.scopeGrowth?.regularCount ?? null;

  return {
    source: table.length ? 'scope table' : 'task Type / Date added columns',
    hasTable: table.length > 0,
    items,
    openItems: items.filter((x) => x.open),
    additionalCount: a.scopeGrowth?.additionalCount ?? 0,
    regularCount: regular,
    growthPct: a.scopeGrowth?.growthPct ?? 0,
    additional,
    // Additional tasks that were later set N/A never became real growth, but
    // they also mean growth is not being measured honestly.
    additionalSetNA: additional.filter((x) => x.status === STATUS.NA).length,
    taskCountVsTemplate: {
      site: a.taskCount ?? 0,
      applicable: a.liveCount ?? 0,
    },
  };
}

/* ---------------- constraints (section 7) ---------------- */

function computeConstraints(site, a, template, reportDate) {
  const rows = (site.detail?.prereqs || []).map((p) => {
    const received = /receiv|complete|done|closed/i.test(p.status || '');
    const partial = /partial/i.test(p.status || '');
    const overdueByWeek = !received && p.byWeek != null && a.reportingWeek
      ? p.byWeek <= (a.weeks?.length ? a.reportingWeek.number ?? 0 : 0) : false;
    return {
      id: p.id, name: p.name, fromTemplate: p.fromTemplate || '—',
      neededFor: p.neededFor || [], provider: p.provider || '—',
      byWeek: p.byWeek ?? null, status: p.status || 'Outstanding',
      received: p.received || null,
      rag: received && !partial ? GREEN : (partial ? AMBER : (overdueByWeek ? RED : AMBER)),
      open: !received || partial,
    };
  });

  // Template prerequisites that an active category needs but the site never
  // listed. The template calls these "Not tracked" rather than silently absent.
  const listedTpl = new Set(rows.map((r) => r.fromTemplate).filter((x) => x && x !== '—'));
  const activeCats = new Set((a.categories || []).filter((c) => c.liveCount > 0).map((c) => c.id));
  const notTracked = (template?.prereqs || [])
    .filter((pt) => activeCats.has(pt.before) && !listedTpl.has(pt.id))
    .map((pt) => ({
      id: pt.id, name: pt.name, fromTemplate: pt.id, neededFor: [pt.before],
      provider: pt.provider || '—', byWeek: null, status: 'Not tracked',
      received: null, rag: AMBER, open: true, notTracked: true,
    }));

  const all = [...rows, ...notTracked];
  return {
    rows: all,
    open: all.filter((x) => x.open),
    noReceivedDates: rows.length > 0 && rows.every((r) => !r.received),
  };
}

/* ---------------- log (section 8) ---------------- */

function computeLog(site, a, reportDate) {
  const rows = (site.detail?.log || []).map((l) => {
    const age = l.raised && !l.cleared ? diffDays(l.raised, reportDate) : null;
    const control = l.kind === STATUS.BLOCKED ? 'Internal' : 'External';
    const rag = l.cleared ? GREEN
      : (age == null ? RED                       // undated open item
        : (age > 10 ? RED : (age > 5 ? AMBER : GREEN)));
    return { ...l, age, control, rag, open: !l.cleared };
  });
  const open = rows.filter((x) => x.open);
  return {
    rows, open,
    oldestAge: open.length ? Math.max(...open.map((x) => x.age ?? 0)) : null,
    undated: open.filter((x) => x.age == null).length,
  };
}

/* ---------------- resources (section 10) ---------------- */

function computeResources(site, a) {
  const rows = (a.resources || []).filter((r) => r.resource !== '(unassigned)');
  const unassigned = (a.resources || []).find((r) => r.resource === '(unassigned)') || null;
  const production = rows.filter((r) => r.open > 0);
  const maxWip = production.length ? Math.max(...production.map((r) => r.wip)) : 0;
  const soleProviders = production.filter((r) => r.open >= 2 && production.length <= 2);
  const causesBlocked = (a.stuckDetail || [])
    .filter((s) => s.status === STATUS.BLOCKED && s.waitingOn)
    .map((s) => s.waitingOn);

  // Who the site is waiting on, from the log and the prerequisites.
  const parties = new Map();
  for (const l of (site.detail?.log || [])) {
    if (l.cleared || !l.waitingOn) continue;
    if (!parties.has(l.waitingOn)) parties.set(l.waitingOn, { party: l.waitingOn, items: [] });
    parties.get(l.waitingOn).items.push(l.id);
  }
  for (const p of (site.detail?.prereqs || [])) {
    if (/receiv|complete|done|closed/i.test(p.status || '') && !/partial/i.test(p.status || '')) continue;
    const who = p.provider;
    if (!who) continue;
    if (!parties.has(who)) parties.set(who, { party: who, items: [] });
    parties.get(who).items.push(p.id);
  }

  return {
    rows, production, unassigned,
    maxWip,
    wipPerResource: production.length ? r1(sum(production, (r) => r.wip) / production.length) : null,
    soleProviders: soleProviders.map((r) => r.resource),
    causesBlocked: [...new Set(causesBlocked)],
    externalParties: [...parties.values()].sort((x, y) => y.items.length - x.items.length),
  };
}

/* ---------------- quality and data integrity (section 11) ---------------- */

function computeQuality(site, a, rollup, logEntries, constraints) {
  const findings = [];
  const add = (id, finding, evidence, impact, correction, severity = AMBER) =>
    findings.push({ id, finding, evidence, impact, correction, severity });

  // Privacy: report that a name was found, never reproduce it.
  const strings = [];
  for (const t of (site.detail?.tasks || [])) strings.push({ where: `task ${t.id}`, text: t.name });
  for (const p of (site.detail?.prereqs || [])) strings.push({ where: `prerequisite ${p.id}`, text: `${p.name} ${p.notes || ''}` });
  for (const l of (site.detail?.log || [])) strings.push({ where: `log ${l.id}`, text: `${l.reason} ${l.notes || ''}` });
  const hits = screenText(strings);
  const privacy = hits.length > 0;
  if (privacy) {
    add('Q-PRIV', 'Possible personal or place name in the data (privacy)',
      `${hits.length} text${hits.length === 1 ? '' : 's'} look like real names: ${[...new Set(hits.map((h) => h.where))].slice(0, 5).join(', ')}. The text itself is deliberately not reproduced here.`,
      'Breaches the workbook privacy rule; this data would be sent to the AI provider.',
      'Replace with codes (R1, E1) in the workbook and re-import.', RED);
  }

  // Header site code vs register.
  const headerCode = site.detail?.meta?.sitecode || site.detail?.meta?.code || null;
  if (headerCode && norm(headerCode) !== norm(site.code)) {
    add('Q-HDR', 'Sheet header disagrees with the register',
      `Sheet header code = ${headerCode}; register row = ${site.code}.`,
      'Figures could be attached to the wrong site.',
      `Set the sheet header code to ${site.code}.`, RED);
  }

  // Log references that resolve to nothing.
  const taskIds = new Set((site.detail?.tasks || []).map((t) => t.id));
  const prereqIds = new Set((site.detail?.prereqs || []).map((p) => p.id));
  const badRefs = (site.detail?.log || []).filter((l) => l.taskId && !taskIds.has(l.taskId) && !prereqIds.has(l.taskId));
  if (badRefs.length) {
    add('Q-REF', 'Log references IDs that do not exist',
      `${badRefs.map((l) => `${l.id} → ${l.taskId}`).join('; ')}.`,
      'The log cannot be linked to tasks automatically.',
      'Use Task IDs that appear in Table 1, or prerequisite IDs from Table 2.');
  }

  // Log vs prerequisite conflict.
  const conflicts = (site.detail?.log || []).filter((l) => {
    if (l.cleared) return false;
    const p = (site.detail?.prereqs || []).find((x) => x.id === l.taskId);
    return p && /receiv|complete|done|closed/i.test(p.status || '');
  });
  if (conflicts.length) {
    add('Q-CONF', 'Log conflicts with prerequisite status',
      `${conflicts.map((l) => `${l.id} open but ${l.taskId} marked ${(site.detail.prereqs.find((x) => x.id === l.taskId) || {}).status}`).join('; ')}.`,
      'Constraint status is ambiguous.',
      'Add cleared dates to the log, or reopen the prerequisite.');
  }

  // Finished tasks with no completion record.
  const noDone = (site.detail?.tasks || []).filter((t) => statusAt(t, 9999) === STATUS.FINISHED && !t.doneWeek && !t.doneDate);
  if (noDone.length) {
    add('Q-DONE', 'Completion dates not recorded',
      `${noDone.length} finished task${noDone.length === 1 ? '' : 's'} have no completion week or date.`,
      'Cycle times and lessons learned cannot be measured.',
      'Fill in the completion column when a task is set to Finished.');
  }

  // Sequence gate bypassed.
  if ((a.outOfOrder || []).length) {
    add('Q-SEQ', 'Sequence gate bypassed',
      `${a.outOfOrder.slice(0, 3).map((o) => `${o.taskId} is ${o.status} while ${o.dependsOn} is ${o.dependsOnStatus}`).join('; ')}.`,
      'Later-stage work may need redoing once the upstream task closes.',
      'Hold the later task until its dependency is Finished, or record why not.');
  }

  // Prerequisites with no received dates.
  if (constraints.noReceivedDates) {
    add('Q-PREQ', 'No received dates on prerequisites',
      'Every prerequisite row has an empty "Date received".',
      'Constraint ageing and make-ready reliability cannot be measured.',
      'Record the date each prerequisite actually arrived.');
  }

  // Missing target weeks.
  const noTarget = (site.detail?.tasks || []).filter((t) => !t.targetWeek && statusAt(t, 9999) !== STATUS.NA);
  if (noTarget.length) {
    add('Q-TGT', 'Tasks without a target week',
      `${noTarget.length} applicable task${noTarget.length === 1 ? '' : 's'} have no target week.`,
      'Those tasks can never be reported as late.',
      'Fill in the Target week column.');
  }

  // Roll-up discrepancies.
  const bad = rollup.filter((c) => !c.consistent);
  if (bad.length) {
    add('Q-ROLL', 'Category roll-up discrepancies',
      bad.map((c) => `${c.id}: ${c.check}`).join('; '),
      'The reported category status does not match its tasks.',
      'Correct whichever is wrong, or record why they differ.');
  }

  // Register status against the sheet.
  const regSaysNotStarted = /not started|planned/i.test(site.status || '');
  if (regSaysNotStarted && (a.finishedCount || 0) + (a.wipCount || 0) > 0) {
    add('Q-REG', 'Register status is stale',
      `Register says "${site.status}" but the sheet shows ${a.finishedCount} finished and ${a.wipCount} in progress.`,
      'The portfolio view misstates this site.',
      'Update the Status column on the Sites register.');
  }

  return {
    findings,
    privacy,
    privacyLocations: [...new Set(hits.map((h) => h.where))],
    count: findings.length,
  };
}

/* ---------------- RAG by dimension (Appendix A.2) ---------------- */

function ragCompletion(completion, timeElapsed, submitted) {
  if (completion == null) return null;
  if (submitted) return GREEN;
  if (timeElapsed == null) return null;
  const ratio = timeElapsed > 0 ? completion / timeElapsed : null;
  if (ratio == null) return null;
  if (ratio >= 0.95) return GREEN;
  if (ratio >= 0.85) return AMBER;
  return RED;
}

function computeRag({ submitted, spiProxy, scheduleVariance, scope, constraints, logEntries, resources, quality, a, milestones }) {
  /* Schedule */
  let schedule;
  if (submitted) {
    // Completed: judged on outcome only.
    const late = scheduleVariance ?? 0;
    schedule = late <= 0 ? GREEN : (late <= 15 ? AMBER : RED);
  } else {
    const targetPassed = scheduleVariance != null && scheduleVariance > 0;
    if (spiProxy == null) schedule = AMBER;
    else if (spiProxy >= 0.95 && !targetPassed) schedule = GREEN;
    else if (spiProxy >= 0.85 && (scheduleVariance ?? 0) <= 5) schedule = AMBER;
    else schedule = RED;
  }
  const scheduleBasis = submitted
    ? `Submitted ${scheduleVariance > 0 ? `${scheduleVariance} days late` : 'on or before target'}.`
    : `SPI proxy ${spiProxy == null ? 'not computable' : r2(spiProxy)}; variance ${scheduleVariance == null ? 'unknown' : `${scheduleVariance} d`}.`;

  /* Scope */
  const scopeRag = scope.openItems.length
    ? RED
    : (scope.additionalCount > 0 && scope.additionalSetNA === scope.additionalCount ? AMBER : GREEN);

  /* Constraints */
  const worstLog = logEntries.open.length
    ? worst(...logEntries.open.map((x) => x.rag)) : GREEN;
  const openPre = constraints.open.length ? AMBER : GREEN;
  const constraintsRag = worst(worstLog, openPre, constraints.noReceivedDates ? AMBER : GREEN);

  /* Resources */
  let resourcesRag = GREEN;
  if (resources.maxWip > 4 || resources.causesBlocked.length) resourcesRag = RED;
  else if (resources.maxWip >= 3 || resources.soleProviders.length) resourcesRag = AMBER;

  /* Quality */
  const qaCat = (a.categories || []).find((c) => /^C130/i.test(c.id));
  let qualityRag = GREEN;
  if (qaCat && qaCat.liveCount === 0) qualityRag = AMBER;               // QA set N/A
  if ((a.outOfOrder || []).length) qualityRag = worst(qualityRag, AMBER);
  if (submitted && qaCat && qaCat.computedPct < 100 && qaCat.liveCount > 0) qualityRag = RED;

  /* Data integrity */
  const dataRag = quality.privacy ? RED
    : (quality.findings.some((f) => f.severity === RED) ? RED
      : (quality.findings.length ? AMBER : GREEN));

  const delivery = [schedule, scopeRag, constraintsRag, resourcesRag, qualityRag];
  const overall = submitted
    ? schedule
    : (schedule === RED || scopeRag === RED ? RED : (delivery.includes(AMBER) || delivery.includes(RED) ? AMBER : GREEN));

  const dim = (rag, basis) => ({ rag, basis });
  return {
    schedule: dim(schedule, scheduleBasis),
    scope: dim(scopeRag, scope.openItems.length
      ? `${scope.openItems.length} scope item${scope.openItems.length === 1 ? '' : 's'} unconfirmed.`
      : `${scope.additionalCount} task${scope.additionalCount === 1 ? '' : 's'} added after kickoff.`),
    constraints: dim(constraintsRag, `${logEntries.open.length} open log item${logEntries.open.length === 1 ? '' : 's'}${logEntries.oldestAge != null ? `, oldest ${logEntries.oldestAge} days` : ''}; ${constraints.open.length} prerequisite${constraints.open.length === 1 ? '' : 's'} outstanding.`),
    resources: dim(resourcesRag, `${resources.production.length} production resource${resources.production.length === 1 ? '' : 's'}; highest WIP ${resources.maxWip}.`),
    quality: dim(qualityRag, qaCat && qaCat.liveCount === 0 ? 'QA category set N/A.' : `${(a.outOfOrder || []).length} sequence anomal${(a.outOfOrder || []).length === 1 ? 'y' : 'ies'}.`),
    dataIntegrity: dim(dataRag, `${quality.findings.length} finding${quality.findings.length === 1 ? '' : 's'}${quality.privacy ? ', including a privacy breach' : ''}.`),
    overall,
  };
}

/* ---------------- movement (section 4.3) ---------------- */

function computeMovement(site, a) {
  const weeks = site.detail?.weeks || [];
  const tasks = site.detail?.tasks || [];
  if (weeks.length < 2 || a.noData) {
    return { comparable: false, reason: 'only one status week is recorded', changes: [] };
  }
  const last = a.velocity.length - 1;
  const prev = last - 1;
  if (prev < 0) return { comparable: false, reason: 'only one status week is recorded', changes: [] };

  const changes = [];
  for (const t of tasks) {
    const before = statusAt(t, prev);
    const now = statusAt(t, last);
    if (before !== now) changes.push({ id: t.id, name: t.name, from: before, to: now });
  }
  for (const c of (site.detail?.categories || [])) {
    const before = statusAt(c, prev);
    const now = statusAt(c, last);
    if (before !== now) changes.push({ id: c.id, name: c.name, from: before, to: now, isCategory: true });
  }
  return {
    comparable: true,
    fromWeek: weeks[prev]?.label, toWeek: weeks[last]?.label,
    finishedThisWeek: a.velocity[last].count,
    changes,
  };
}

/* ============================================================
   Master report
   ============================================================ */

/**
 * @param {object[]} siteReports  computeSiteReport() output for each selected site
 * @param {object}   portfolio    analysePortfolio() output, restricted to the selection
 */
export function computeMasterReport(siteReports, portfolio, opts = {}) {
  const reportDate = opts.reportDate || todayISO();
  const reports = siteReports.filter(Boolean);
  const active = reports.filter((r) => !r.submitted);
  const done = reports.filter((r) => r.submitted);

  const sumOf = (list, f) => sum(list, f);
  const applicable = sumOf(reports, (r) => r.metrics.applicableTasks || 0);
  const finished = sumOf(reports, (r) => r.metrics.finished || 0);
  const activeApplicable = sumOf(active, (r) => r.metrics.applicableTasks || 0);
  const activeFinished = sumOf(active, (r) => r.metrics.finished || 0);

  // Weighted figures come from each site's own totals, never recomputed differently.
  const bySite = new Map(portfolio.sites.map((s) => [s.code, s]));
  const weightAll = sumOf(reports, (r) => bySite.get(r.code)?.totalWeight || 0);
  const weightDone = sumOf(reports, (r) => bySite.get(r.code)?.doneWeight || 0);
  const weightActiveAll = sumOf(active, (r) => bySite.get(r.code)?.totalWeight || 0);
  const weightActiveDone = sumOf(active, (r) => bySite.get(r.code)?.doneWeight || 0);

  const onTime = done.length ? done.filter((r) => (r.metrics.scheduleVariance ?? 0) <= 0).length / done.length * 100 : null;

  const latestForecast = active
    .map((r) => r.metrics.forecastDate).filter(Boolean).sort().pop() || null;
  const latestTarget = reports.map((r) => r.docControl && bySite.get(r.code)?.target).filter(Boolean).sort().pop() || null;

  /* ---------- back-end concentration (C110–C150) ---------- */
  const BACK_END = /^C1[1-5]0$/i;
  let openBackEnd = 0;
  let openAll = 0;
  for (const r of active) {
    for (const c of r.wbs.rows) {
      const open = c.liveCount - c.doneCount;
      if (open <= 0) continue;
      openAll += open;
      if (BACK_END.test(c.id)) openBackEnd += open;
    }
  }

  /* ---------- resource load across the selection ---------- */
  const load = new Map();
  for (const r of reports) {
    for (const res of r.resources.rows) {
      if (!load.has(res.resource)) load.set(res.resource, { resource: res.resource, sites: [], activeSites: 0, open: 0 });
      const e = load.get(res.resource);
      e.sites.push(r.code);
      e.open += res.open;
      if (!r.submitted) e.activeSites++;
    }
  }
  const resourceLoad = [...load.values()].map((e) => ({
    ...e,
    load: e.activeSites >= 3 ? 'CRITICAL' : (e.activeSites === 2 ? 'HIGH' : 'OK'),
  })).sort((a, b) => b.activeSites - a.activeSites || b.open - a.open);

  /* ---------- party dependency ---------- */
  const party = new Map();
  for (const r of reports) {
    for (const p of r.resources.externalParties) {
      if (!party.has(p.party)) party.set(p.party, { party: p.party, items: [], sites: new Set() });
      const e = party.get(p.party);
      p.items.forEach((i) => e.items.push(`${i} (${r.code})`));
      e.sites.add(r.code);
    }
  }
  const parties = [...party.values()]
    .map((p) => ({ party: p.party, count: p.items.length, sites: [...p.sites], items: p.items }))
    .sort((a, b) => b.count - a.count);

  /* ---------- systemic patterns ----------
     A finding on three or more sites, or on half of them when fewer than six
     are selected, is a process issue rather than a site issue. */
  const threshold = reports.length < 6 ? Math.ceil(reports.length / 2) : 3;
  const patternMap = new Map();
  for (const r of reports) {
    for (const f of r.quality.findings) {
      if (!patternMap.has(f.finding)) patternMap.set(f.finding, { pattern: f.finding, sites: [], correction: f.correction, impact: f.impact });
      patternMap.get(f.finding).sites.push(r.code);
    }
  }
  const patterns = [...patternMap.values()]
    .map((p) => ({ ...p, count: p.sites.length, systemic: p.sites.length >= threshold }))
    .sort((a, b) => b.count - a.count);

  /* ---------- Intervention Priority Index (Appendix A) ---------- */
  const ipi = active.map((r) => {
    const s = bySite.get(r.code) || {};
    const ragPts = r.overall === RED ? 3 : (r.overall === AMBER ? 2 : 1);
    const prioPts = /high/i.test(s.priority || '') ? 2 : (/medium/i.test(s.priority || '') ? 1 : 0);
    const toTarget = s.target ? diffDays(reportDate, s.target) : null;
    const timePts = toTarget == null ? 0 : (toTarget < 0 ? 3 : (toTarget <= 7 ? 2 : (toTarget <= 21 ? 1 : 0)));
    const gapPts = clamp(3 * (1 - (r.metrics.spiProxy ?? 1)), 0, 3);
    return {
      code: r.code, overall: r.overall, priority: s.priority || '—',
      daysToTarget: toTarget, spiProxy: r.metrics.spiProxy,
      components: { rag: ragPts, priority: prioPts, time: timePts, gap: r1(gapPts) },
      ipi: r1(ragPts + prioPts + timePts + gapPts),
      target: s.target || null,
    };
  }).sort((a, b) => b.ipi - a.ipi || String(a.target).localeCompare(String(b.target)));

  /* ---------- portfolio RAG ---------- */
  const redCount = reports.filter((r) => r.overall === RED).length;
  const anyHighRed = active.some((r) => r.overall === RED && /high/i.test((bySite.get(r.code) || {}).priority || ''));
  const portfolioRag = (anyHighRed || redCount >= reports.length / 3) ? RED
    : (reports.some((r) => r.overall !== GREEN) ? AMBER : GREEN);

  const week = isoWeek(reportDate);
  return {
    kind: 'master',
    docControl: {
      reportId: `BIM-MSDT-MAR-${String(reportDate).slice(0, 4)}-W${week}`,
      reportType: 'Master analysis (portfolio) report',
      sites: reports.map((r) => r.code),
      reportDate,
      basis: `Computed site reports for ${reports.length} selected site${reports.length === 1 ? '' : 's'}`,
    },
    overall: portfolioRag,
    counts: {
      sites: reports.length, active: active.length, submitted: done.length,
      applicable, finished,
      activeApplicable, activeFinished,
    },
    kpis: {
      completionCount: r1(pct(finished, applicable)),
      completionWeighted: r1(pct(weightDone, weightAll)),
      activeCompletionCount: r1(pct(activeFinished, activeApplicable)),
      activeCompletionWeighted: r1(pct(weightActiveDone, weightActiveAll)),
      onTimeDelivery: r1(onTime),
      latestForecast, latestTarget,
      backEndConcentration: r1(pct(openBackEnd, openAll)),
      openTasksActive: openAll,
      openBackEnd,
    },
    ragMatrix: reports.map((r) => ({
      code: r.code,
      schedule: r.dimensions.schedule.rag,
      scope: r.dimensions.scope.rag,
      constraints: r.dimensions.constraints.rag,
      resources: r.dimensions.resources.rag,
      quality: r.dimensions.quality.rag,
      dataIntegrity: r.dimensions.dataIntegrity.rag,
      overall: r.overall,
    })),
    schedule: reports.map((r) => {
      const s = bySite.get(r.code) || {};
      return {
        code: r.code, start: s.start || null, target: s.target || null,
        forecast: r.metrics.forecastDate, variance: r.metrics.scheduleVariance,
        submitted: r.submitted, overall: r.overall,
        completionWeighted: r.metrics.completionWeighted,
        timeElapsed: r.metrics.timeElapsed,
      };
    }),
    categories: aggregateCategories(reports),
    scope: reports.map((r) => ({
      code: r.code, source: r.scope.source, additional: r.scope.additionalCount,
      setNA: r.scope.additionalSetNA, openItems: r.scope.openItems.length,
      hasTable: r.scope.hasTable,
    })),
    constraints: reports.flatMap((r) => r.log.open.map((l) => ({ code: r.code, ...l }))),
    parties,
    resourceLoad,
    patterns,
    systemicPatterns: patterns.filter((p) => p.systemic),
    ipi,
    appendixB: reports.map((r) => {
      const s = bySite.get(r.code) || {};
      return {
        code: r.code, priority: s.priority || '—',
        resources: r.resources.rows.map((x) => x.resource),
        overall: r.overall, reportId: r.docControl.reportId,
        submitted: r.submitted,
      };
    }),
    references: MASTER_REFERENCES,
    sections: MASTER_SECTIONS,
    noActiveSites: active.length === 0,
  };
}

function aggregateCategories(reports) {
  const map = new Map();
  for (const r of reports) {
    for (const c of r.wbs.rows) {
      if (!map.has(c.id)) map.set(c.id, { id: c.id, name: c.name, applicable: 0, finished: 0, open: 0, stuckSites: [], bySite: {} });
      const e = map.get(c.id);
      e.applicable += c.liveCount;
      e.finished += c.doneCount;
      e.open += Math.max(0, c.liveCount - c.doneCount);
      if (c.stuck > 0) e.stuckSites.push(r.code);
      e.bySite[r.code] = c.liveCount ? Math.round(pct(c.doneCount, c.liveCount)) : null;
    }
  }
  return [...map.values()];
}

/* ============================================================
   Risk scoring — the split between judgement and arithmetic

   The model supplies probability and impact (judgement). The score and the
   rating band are multiplication and a lookup, so they are computed here. That
   way two reports with the same P and I can never disagree on the rating.
   ============================================================ */

export function scoreRisks(risks) {
  return (Array.isArray(risks) ? risks : []).map((r, i) => {
    const p = clamp(Math.round(Number(r.probability) || 0), 0, 5);
    const impact = clamp(Math.round(Number(r.impact) || 0), 0, 5);
    const score = p * impact;
    return {
      ...r,
      id: r.id || `RK${i + 1}`,
      probability: p,
      impact,
      score,
      rating: score >= 15 ? RED : (score >= 8 ? AMBER : GREEN),
    };
  }).sort((a, b) => b.score - a.score);
}

/* ============================================================
   References — fixed bibliography from the report standard
   ============================================================ */

export const SITE_REFERENCES = [
  ['Project Management Institute (2021). A Guide to the Project Management Body of Knowledge (PMBOK Guide), 7th ed., and The Standard for Project Management. Newtown Square, PA: PMI.', 'performance domains (measurement, uncertainty, delivery); status and forecast reporting; issue and risk logs'],
  ['Project Management Institute (2019). The Standard for Earned Value Management. Newtown Square, PA: PMI.', 'SPI concept behind the time-based SPI proxy, schedule variance and forecasting logic'],
  ['ISO 21502:2020. Project, programme and portfolio management — Guidance on project management. Geneva: ISO.', 'report content and control practices: progress, change control, issues, risks, decision requests'],
  ['AXELOS (2023). Managing Successful Projects with PRINCE2, 7th ed. London: TSO.', 'Highlight Report and End Project Report product descriptions; document control block; closeout content'],
  ['ISO 19650-1:2018 and ISO 19650-2:2018. Organization and digitization of information about buildings and civil engineering works, including BIM. Geneva: ISO.', 'information-delivery framing, CDE and information-quality findings'],
  ['ISO 31000:2018. Risk management — Guidelines. Geneva: ISO.', 'risk statement format (cause to effect), probability x impact scoring and response strategies'],
  ['Ballard, G. (2000). The Last Planner System of Production Control. PhD thesis, University of Birmingham.', 'make-ready constraint analysis, lookahead planning and week-on-week reliability'],
  ['Kunz, J. and Fischer, M. (2012). Virtual Design and Construction: Themes, Case Studies and Implementation Suggestions. CIFE Working Paper #097, Stanford University.', 'VDC practice of setting measurable production metrics and reporting them against objectives'],
  ['U.S. Government Accountability Office (2015). GAO Schedule Assessment Guide: Best Practices for Project Schedules (GAO-16-89G). Washington, DC.', 'schedule-quality checks: logic/sequence anomalies, realistic forecasts, baseline integrity'],
  ['Infrastructure and Projects Authority (UK). Delivery Confidence Assessment (DCA) definitions, IPA Annual Reports on Major Projects. London: HM Treasury / Cabinet Office.', 'structure of the single overall RAG rating with a one-line rationale'],
  ['Kerzner, H. (2017). Project Management Metrics, KPIs, and Dashboards, 3rd ed. Hoboken, NJ: Wiley.', 'KPI selection, thresholds and dashboard layout'],
  ['Minto, B. (2009). The Pyramid Principle: Logic in Writing and Thinking, 3rd ed. Harlow: Pearson.', 'answer-first executive summary: bottom line, supporting key messages, decisions'],
  ['Harvard Business Review (2012). HBR Guide to Project Management. Boston, MA: Harvard Business Review Press.', 'stakeholder-facing status communication and decision requests'],
  ['Moavenzadeh, F. (2009). 1.040 Project Management, Spring 2009. MIT OpenCourseWare.', 'construction project control structure: scope, schedule, resources and risk as linked control areas'],
  ['User input: the BIM Multi-Site Delivery Tracker workbook supplied for this report.', 'all project data, status definitions, category Definitions of Done and typical durations'],
];

export const MASTER_REFERENCES = [
  ['Project Management Institute (2017). The Standard for Portfolio Management, 4th ed. Newtown Square, PA: PMI.', 'portfolio performance reporting, prioritisation and capacity balancing'],
  ['Project Management Institute (2017). The Standard for Program Management, 4th ed. Newtown Square, PA: PMI.', 'cross-project dependencies, shared resources and benefits-level reporting'],
  ['ISO 21504:2022. Project, programme and portfolio management — Guidance on portfolio management. Geneva: ISO.', 'portfolio monitoring, prioritisation and balancing'],
  ['AXELOS (2020). Managing Successful Programmes (MSP), 5th ed. London: TSO.', 'programme-level dashboard, issue and risk aggregation, dependency management'],
  ['Project Management Institute (2021). PMBOK Guide, 7th ed. Newtown Square, PA: PMI.', 'measurement and uncertainty performance domains; forecasting'],
  ['Project Management Institute (2019). The Standard for Earned Value Management. Newtown Square, PA: PMI.', 'basis for the time-based SPI proxy and schedule variance'],
  ['ISO 21502:2020. Project, programme and portfolio management — Guidance on project management. Geneva: ISO.', 'progress, change, issue and decision reporting content'],
  ['ISO 31000:2018. Risk management — Guidelines. Geneva: ISO.', 'portfolio risk statement format and probability x impact scoring'],
  ['Ballard, G. (2000). The Last Planner System of Production Control. PhD thesis, University of Birmingham.', 'constraint analysis and make-ready reliability across sites'],
  ['U.S. Government Accountability Office (2015). GAO Schedule Assessment Guide (GAO-16-89G). Washington, DC.', 'schedule quality and forecast credibility'],
  ['Kerzner, H. (2017). Project Management Metrics, KPIs, and Dashboards, 3rd ed. Hoboken, NJ: Wiley.', 'portfolio KPI selection and dashboard layout'],
  ['Minto, B. (2009). The Pyramid Principle, 3rd ed. Harlow: Pearson.', 'answer-first executive summary and decision requests'],
  ['User input: the BIM Multi-Site Delivery Tracker workbook supplied for this report.', 'all project data for the selected sites'],
];

/* ============================================================
   Appendix A content — definitions shown in the report itself
   ============================================================ */

export const METRIC_DEFINITIONS = [
  ['Status date', 'End date of the latest status week column on the site sheet. All progress figures are at this date.'],
  ['Applicable tasks', 'Task rows whose latest status is not N/A. Category rows are excluded from counts.'],
  ['Completion (count)', 'Finished applicable tasks / applicable tasks.'],
  ['Completion (weighted)', 'Sum of weight of Finished applicable tasks / sum of weight of applicable tasks. Missing weight defaults to 1 and is reported as a data finding.'],
  ['Time elapsed', '(Status date − register start) / (register target − register start), capped at 100%.'],
  ['Time-based SPI (proxy)', 'Weighted completion / time elapsed. Assumes linear planned progress; an indicator, not an EVM SPI. Not computed for a submitted site.'],
  ['Schedule variance', 'Actual or forecast submission − register target, in calendar days. Positive means late.'],
  ['Weekly throughput', 'Tasks that moved to Finished between the last two status weeks.'],
  ['Stalled task', 'Task in WIP for the last two status weeks with no status change.'],
  ['Constraint age', 'Report date − raised date, in calendar days, for log entries with no cleared date.'],
  ['Scope growth', 'Additional tasks / Regular tasks, by count, with the date added.'],
  ['Roll-up discrepancy', 'Category row status differs from the status implied by its tasks, or the category hides a Blocked or Waiting-on task.'],
  ['Indicative forecast', 'When throughput is zero or unmeasurable: sum of the template typical durations of the categories with open work, run forward from the status date on the working calendar.'],
  ['Intervention Priority Index', 'Overall RAG (3/2/1) + site priority (2/1/0) + time pressure (3/2/1/0) + 3 x (1 − SPI proxy). Active sites only; ties broken by the earlier target date.'],
];

export const RAG_THRESHOLDS = [
  ['Schedule (active)', 'SPI ≥ 0.95 and forecast ≤ target', 'SPI 0.85–0.94 or forecast ≤ 5 days late', 'SPI < 0.85, forecast > 5 days late, or target passed'],
  ['Schedule (completed)', 'Submitted on or before target', '≤ 15 days late', '> 15 days late'],
  ['Scope', 'Baseline confirmed; growth recorded', 'Growth recorded but unapproved', 'Scope unconfirmed near target, or growth unrecorded'],
  ['Constraints', 'No open items', 'Open items ≤ 10 days old with dates', 'Any item > 10 days, past expected date, or undated'],
  ['Resources', '≤ 2 WIP per production resource; no single-provider dependency', '3–4 WIP per resource, or a single provider on the critical path', '> 4 WIP per resource, or a resource is the cause of a Blocked item'],
  ['Quality', 'QA definition of done met or planned', 'QA set N/A without a record, or a sequence risk', 'QA skipped at submission with defects found'],
  ['Data integrity', 'No findings', 'Findings with no reporting impact', 'Privacy breach or wrong site baseline'],
  ['Overall (active)', 'All delivery dimensions GREEN', 'Any delivery dimension AMBER', 'Schedule or Scope RED'],
  ['Overall (completed)', 'Schedule outcome GREEN', 'Schedule outcome AMBER', 'Schedule outcome RED'],
];
