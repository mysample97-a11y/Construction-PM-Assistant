/**
 * Deterministic analysis engine.
 *
 * Every number the app shows or sends to the AI is produced here, in the
 * browser, from the spreadsheet. The model never calculates — it reads these
 * results and writes commentary. If a figure in this file is wrong, the whole
 * tool is worse than useless, so each metric is defined explicitly below and
 * covered by tools/verify-engine.mjs.
 *
 * Conventions used throughout:
 *  - "live" tasks exclude anything marked N/A in the reporting week.
 *  - "as at week W" means the status recorded in the LAST week column that has
 *    any data, unless a specific week is requested.
 *  - Weight defaults to 1. Percentages are reported both by count and by
 *    weight, because they diverge and neither is the whole truth.
 */

import { STATUS, STUCK } from './parser.js';
import { toISO, addDays, diffDays } from './util.js';

const sum = (a, f = (x) => x) => a.reduce((n, x) => n + (Number(f(x)) || 0), 0);
const pct = (n, d) => (d > 0 ? (n / d) * 100 : 0);

/**
 * Status of a task in a given week index, carrying the last known status
 * forward. A blank week means "unchanged", not "unknown" — people fill in what
 * moved, not every cell every week.
 *
 * `weekly` is defended because a malformed sheet can produce a row without it,
 * and a crash here would take the whole analysis down rather than reporting the
 * one bad row.
 */
function statusAt(task, wi) {
  const wk = task?.weekly;
  if (!Array.isArray(wk) || !wk.length) return null;
  for (let i = Math.min(wi, wk.length - 1); i >= 0; i--) {
    if (wk[i]) return wk[i];
  }
  return null;
}

/** Index of the last week column that any task has data in. */
export function lastReportedWeek(detail) {
  if (!detail) return -1;
  let last = -1;
  const scan = (rows) => {
    for (const r of rows || []) {
      const wk = Array.isArray(r?.weekly) ? r.weekly : [];
      for (let i = 0; i < wk.length; i++) if (wk[i]) last = Math.max(last, i);
    }
  };
  scan(detail.tasks);
  scan(detail.categories);
  return last;
}

/* ========================================================================
   Per-site analysis
   ======================================================================== */

export function analyseSite(site, template, opts = {}) {
  const d = site.detail;
  const weeks = d?.weeks || [];
  const asAt = opts.weekIndex !== undefined && opts.weekIndex !== null
    ? opts.weekIndex
    : lastReportedWeek(d);

  const base = {
    code: site.code,
    description: site.description,
    wave: site.wave,
    coordinator: site.coordinator,
    start: site.start,
    target: site.target,
    actualSubmission: site.actual,
    registerStatus: site.status,
    priority: site.priority,
    sheet: d?.sheetName || null,
    hasDetail: !!d,
    weekCount: weeks.length,
    reportingWeek: asAt >= 0 ? weeks[asAt] : null,
    risks: [],
  };

  if (!d || asAt < 0) {
    base.taskCount = d?.tasks.length || 0;
    base.noData = true;
    base.gridTasks = [];
    base.timeline = {
      today: toISO(new Date()),
      startDate: site.start || null,
      targetSubmission: site.target || null,
      weeksElapsed: 0,
      weeksRemainingToTarget: null,
      weekColumnsAvailable: weeks.length,
      previousAnalysis: opts.previous || null,
    };
    if (!d) {
      base.risks.push({
        level: 'medium', code: 'no_detail_sheet',
        title: `${site.code} has no detail sheet`,
        detail: 'It appears in the register but has no task list, so nothing about it can ever be reported as late. It is being tracked in name only.',
      });
    } else {
      base.risks.push({
        level: 'medium', code: 'no_status_yet',
        title: `${site.code} has ${d.tasks.length} tasks but no weekly status filled in`,
        detail: 'Progress cannot be measured until at least one week column has statuses.',
      });
    }
    return base;
  }

  const tasks = d.tasks;
  const live = tasks.filter((t) => statusAt(t, asAt) !== STATUS.NA);
  const naCount = tasks.length - live.length;
  const finished = live.filter((t) => statusAt(t, asAt) === STATUS.FINISHED);
  const wip = live.filter((t) => statusAt(t, asAt) === STATUS.WIP);
  const blocked = live.filter((t) => statusAt(t, asAt) === STATUS.BLOCKED);
  const waiting = live.filter((t) => statusAt(t, asAt) === STATUS.WAITING);
  const notStarted = live.filter((t) => {
    const s = statusAt(t, asAt);
    return s === STATUS.NOT_STARTED || s === null;
  });

  const wLive = sum(live, (t) => t.weight || 1);
  const wDone = sum(finished, (t) => t.weight || 1);

  /* ---- velocity: newly finished tasks in each week ---- */
  const velocity = [];
  for (let i = 0; i <= asAt; i++) {
    const done = live.filter((t) => {
      const now = statusAt(t, i);
      const before = i === 0 ? null : statusAt(t, i - 1);
      return now === STATUS.FINISHED && before !== STATUS.FINISHED;
    });
    velocity.push({
      weekIndex: i,
      label: weeks[i]?.label || `W${i + 1}`,
      end: weeks[i]?.end || null,
      count: done.length,
      weight: sum(done, (t) => t.weight || 1),
    });
  }

  const activeWeeks = velocity.length;
  const avgVelocity = activeWeeks ? sum(velocity, (v) => v.count) / activeWeeks : 0;
  const recent = velocity.slice(-3);
  const recentVelocity = recent.length ? sum(recent, (v) => v.count) / recent.length : 0;
  const avgWeightVelocity = activeWeeks ? sum(velocity, (v) => v.weight) / activeWeeks : 0;

  /* ---- forecast ----
     Two rates are reported because they answer different questions: the
     all-time average is stable, the 3-week average reflects what is happening
     now. Where they disagree the disagreement is itself the finding. */
  const remaining = live.length - finished.length;
  const remainingWeight = wLive - wDone;

  const forecastFrom = (rate) => {
    if (!rate || rate <= 0) return null;
    const weeksNeeded = Math.ceil(remaining / rate);
    return {
      rate: Math.round(rate * 100) / 100,
      weeksNeeded,
      finishWeek: asAt + 1 + weeksNeeded,
      finishDate: weeks[asAt]?.end ? addDays(weeks[asAt].end, weeksNeeded * 7) : null,
    };
  };
  const forecastAvg = forecastFrom(avgVelocity);
  const forecastRecent = forecastFrom(recentVelocity);

  /* ---- target comparison ---- */
  const targetWeekIndex = site.target && weeks.length
    ? weeks.findIndex((w) => w.end >= site.target)
    : -1;
  const targetWeekNo = targetWeekIndex >= 0 ? targetWeekIndex + 1 : null;

  let slipWeeks = null;
  let forecastVsTarget = null;
  const chosen = forecastRecent || forecastAvg;
  if (chosen && site.target) {
    if (chosen.finishDate) {
      const days = diffDays(site.target, chosen.finishDate);
      slipWeeks = days === null ? null : Math.round(days / 7);
    } else if (targetWeekNo) {
      slipWeeks = chosen.finishWeek - targetWeekNo;
    }
    forecastVsTarget = slipWeeks;
  }

  /* ---- categories: self-reported vs computed ---- */
  const catMap = new Map((template?.categories || []).map((c) => [c.id, c]));
  const categories = d.categories.map((c) => {
    const kids = tasks.filter((t) => t.categoryId === c.id);
    const kidsLive = kids.filter((t) => statusAt(t, asAt) !== STATUS.NA);
    const kidsDone = kidsLive.filter((t) => statusAt(t, asAt) === STATUS.FINISHED);
    const computedPct = pct(sum(kidsDone, (t) => t.weight || 1), sum(kidsLive, (t) => t.weight || 1));
    const self = statusAt(c, asAt);
    const impliedSelf = self === STATUS.FINISHED ? 100 : (self === STATUS.NOT_STARTED || !self ? 0 : null);
    return {
      id: c.id,
      name: c.name || catMap.get(c.id)?.name || c.id,
      order: catMap.get(c.id)?.order ?? null,
      definitionOfDone: catMap.get(c.id)?.definitionOfDone || '',
      selfStatus: self,
      taskCount: kids.length,
      liveCount: kidsLive.length,
      doneCount: kidsDone.length,
      computedPct: Math.round(computedPct),
      // A category the coordinator calls finished while tasks under it are open,
      // or vice versa. Neither is automatically wrong; the gap is the point.
      divergence: impliedSelf === null ? null : Math.round(impliedSelf - computedPct),
      stuck: kidsLive.filter((t) => STUCK.includes(statusAt(t, asAt))).length,
    };
  });

  /* ---- stuck work, with how long ---- */
  const stuckDetail = [...blocked, ...waiting].map((t) => {
    let since = asAt;
    while (since > 0 && STUCK.includes(statusAt(t, since - 1))) since--;
    const entry = (d.log || []).find((l) => l.taskId === t.id && !l.cleared);
    return {
      taskId: t.id,
      name: t.name,
      category: t.category,
      discipline: t.discipline,
      status: statusAt(t, asAt),
      resource: t.resource,
      weeksStuck: asAt - since + 1,
      sinceWeek: weeks[since]?.label || null,
      reason: entry?.reason || null,
      waitingOn: entry?.waitingOn || null,
      expected: entry?.expected || null,
      logged: !!entry,
    };
  }).sort((a, b) => b.weeksStuck - a.weeksStuck);

  /* ---- stalled: WIP for several weeks with no change ---- */
  const stalled = wip.filter((t) => {
    let n = 0;
    for (let i = asAt; i >= 0 && statusAt(t, i) === STATUS.WIP; i--) n++;
    t._wipWeeks = n;
    return n >= (opts.stalledWeeks || 3);
  }).map((t) => ({
    taskId: t.id, name: t.name, category: t.category,
    resource: t.resource, weeksInWip: t._wipWeeks,
  })).sort((a, b) => b.weeksInWip - a.weeksInWip);

  /* ---- overdue against the task's own target week ---- */
  const overdue = live.filter((t) =>
    t.targetWeek && t.targetWeek <= asAt + 1 && statusAt(t, asAt) !== STATUS.FINISHED)
    .map((t) => ({
      taskId: t.id, name: t.name, category: t.category, resource: t.resource,
      targetWeek: t.targetWeek, weeksLate: (asAt + 1) - t.targetWeek,
      status: statusAt(t, asAt),
    })).sort((a, b) => b.weeksLate - a.weeksLate);

  /* ---- scope growth ---- */
  const additional = tasks.filter((t) => t.type === 'Additional');
  const scopeGrowth = {
    additionalCount: additional.length,
    additionalWeight: sum(additional, (t) => t.weight || 1),
    regularCount: tasks.length - additional.length,
    growthPct: tasks.length - additional.length > 0
      ? Math.round(pct(additional.length, tasks.length - additional.length)) : 0,
    items: additional.map((t) => ({
      taskId: t.id, name: t.name, category: t.category,
      added: t.added, weight: t.weight || 1,
      status: statusAt(t, asAt),
    })),
  };

  /* ---- prerequisites ---- */
  const prereqIndex = new Map((d.prereqs || []).map((p) => [p.id, p]));
  const outstandingPrereqs = (d.prereqs || [])
    .filter((p) => !/receiv|done|complete|closed|not required/i.test(p.status || ''))
    .map((p) => {
      const blocks = tasks.filter((t) => t.prereqs.includes(p.id));
      return {
        id: p.id, name: p.name, provider: p.provider,
        byWeek: p.byWeek, status: p.status || 'Outstanding',
        overdue: p.byWeek ? p.byWeek <= asAt + 1 : false,
        blocksCount: blocks.length,
        blocks: blocks.map((t) => t.id),
      };
    });

  const unknownPrereqRefs = [];
  for (const t of tasks) {
    for (const p of t.prereqs) {
      if (!prereqIndex.has(p)) unknownPrereqRefs.push({ taskId: t.id, ref: p });
    }
  }

  /* ---- dependencies ---- */
  const taskIndex = new Map(tasks.map((t) => [t.id, t]));
  const unknownDeps = [];
  const outOfOrder = [];
  for (const t of tasks) {
    for (const depId of t.depends) {
      const dep = taskIndex.get(depId);
      if (!dep) { unknownDeps.push({ taskId: t.id, ref: depId }); continue; }
      const tStatus = statusAt(t, asAt);
      const depStatus = statusAt(dep, asAt);
      const tStarted = tStatus === STATUS.WIP || tStatus === STATUS.FINISHED;
      if (tStarted && depStatus !== STATUS.FINISHED && depStatus !== STATUS.NA) {
        outOfOrder.push({
          taskId: t.id, name: t.name, status: tStatus,
          dependsOn: dep.id, dependsOnName: dep.name, dependsOnStatus: depStatus,
        });
      }
    }
  }

  /* ---- resource load in the reporting week ---- */
  const openByResource = new Map();
  for (const t of live) {
    const s = statusAt(t, asAt);
    if (s === STATUS.FINISHED) continue;
    const r = t.resource || '(unassigned)';
    if (!openByResource.has(r)) openByResource.set(r, { resource: r, open: 0, wip: 0, stuck: 0, weight: 0 });
    const e = openByResource.get(r);
    e.open++;
    e.weight += t.weight || 1;
    if (s === STATUS.WIP) e.wip++;
    if (STUCK.includes(s)) e.stuck++;
  }
  const resources = [...openByResource.values()].sort((a, b) => b.open - a.open);

  /* ---- data quality ---- */
  const dataIssues = [];
  const noTarget = live.filter((t) => !t.targetWeek).length;
  if (noTarget) {
    dataIssues.push({
      code: 'no_target_week',
      detail: `${noTarget} of ${live.length} live tasks have no target week, so they can never be reported as late.`,
    });
  }
  for (const t of tasks) {
    // Progress cannot go backwards; if it does, it is rework or a typo.
    for (let i = 1; i <= asAt; i++) {
      if (statusAt(t, i - 1) === STATUS.FINISHED && statusAt(t, i) && statusAt(t, i) !== STATUS.FINISHED && statusAt(t, i) !== STATUS.NA) {
        dataIssues.push({
          code: 'regression',
          detail: `${t.id} "${t.name}" was Finished in ${weeks[i - 1]?.label} then ${statusAt(t, i)} in ${weeks[i]?.label}. Rework, or a reporting error.`,
        });
        break;
      }
    }
    if (t.doneDate && statusAt(t, asAt) !== STATUS.FINISHED) {
      dataIssues.push({
        code: 'completion_mismatch',
        detail: `${t.id} has a completion date of ${t.doneDate} but its status is "${statusAt(t, asAt) || 'blank'}".`,
      });
    }
    if (statusAt(t, asAt) === STATUS.FINISHED && !t.doneWeek && !t.doneDate) {
      dataIssues.push({ code: 'no_completion_record', detail: `${t.id} is Finished but has no completion week or date recorded.` });
    }
  }
  for (const u of unknownDeps) {
    dataIssues.push({ code: 'unknown_dependency', detail: `${u.taskId} depends on "${u.ref}", which is not a task on this sheet.` });
  }
  for (const u of unknownPrereqRefs) {
    dataIssues.push({ code: 'unknown_prerequisite', detail: `${u.taskId} references prerequisite "${u.ref}", which is not in Table 2.` });
  }
  const unloggedStuck = stuckDetail.filter((s) => !s.logged);
  if (unloggedStuck.length) {
    dataIssues.push({
      code: 'unlogged_blocker',
      detail: `${unloggedStuck.length} task${unloggedStuck.length === 1 ? ' is' : 's are'} Blocked or Waiting on with no entry in the log, so the reason is unknown: ${unloggedStuck.slice(0, 5).map((s) => s.taskId).join(', ')}.`,
    });
  }

  /* ---- risks ---- */
  const risks = [];
  if (slipWeeks !== null && slipWeeks > 0) {
    risks.push({
      level: slipWeeks > 4 ? 'high' : 'medium', code: 'forecast_slip',
      title: `${site.code} forecasts ${slipWeeks} week${slipWeeks === 1 ? '' : 's'} past its target submission`,
      detail: `${remaining} of ${live.length} tasks remain. At the recent rate of ${Math.round(recentVelocity * 100) / 100} tasks/week that needs about ${chosen.weeksNeeded} more weeks, landing around ${chosen.finishDate || `week ${chosen.finishWeek}`} against a target of ${site.target}.`,
    });
  }
  for (const s of stuckDetail) {
    if (s.weeksStuck >= 2) {
      risks.push({
        level: s.weeksStuck >= 4 ? 'high' : 'medium', code: 'stuck',
        title: `${s.taskId} has been ${s.status} for ${s.weeksStuck} weeks`,
        detail: `${s.name}${s.waitingOn ? ` — waiting on ${s.waitingOn}` : ''}${s.reason ? `: ${s.reason}` : '. No reason logged.'}`,
      });
    }
  }
  for (const p of outstandingPrereqs) {
    if (p.overdue && p.blocksCount) {
      risks.push({
        level: 'high', code: 'prereq_overdue',
        title: `Prerequisite ${p.id} is overdue and blocks ${p.blocksCount} task${p.blocksCount === 1 ? '' : 's'}`,
        detail: `"${p.name}" was required by week ${p.byWeek}${p.provider ? `, from ${p.provider}` : ''}. Status: ${p.status}.`,
      });
    }
  }
  if (recentVelocity === 0 && remaining > 0 && activeWeeks >= 2) {
    risks.push({
      level: 'high', code: 'stalled_site',
      title: `${site.code} completed nothing in the last ${Math.min(3, activeWeeks)} weeks`,
      detail: `${remaining} tasks remain open. A site with zero throughput has no forecast finish date at all.`,
    });
  }
  if (scopeGrowth.growthPct >= 15) {
    risks.push({
      level: 'medium', code: 'scope_growth',
      title: `${site.code} scope has grown ${scopeGrowth.growthPct}%`,
      detail: `${scopeGrowth.additionalCount} tasks added after kickoff against ${scopeGrowth.regularCount} agreed originally.`,
    });
  }
  if (outOfOrder.length) {
    risks.push({
      level: 'medium', code: 'out_of_sequence',
      title: `${outOfOrder.length} task${outOfOrder.length === 1 ? ' is' : 's are'} running ahead of what they depend on`,
      detail: outOfOrder.slice(0, 3).map((o) => `${o.taskId} is ${o.status} but ${o.dependsOn} is only ${o.dependsOnStatus}`).join('; ') + '. Usually a sign that rework is coming.',
    });
  }
  for (const c of categories) {
    if (c.divergence !== null && Math.abs(c.divergence) >= 40 && c.liveCount > 0) {
      risks.push({
        level: 'medium', code: 'category_divergence',
        title: `Category ${c.id} ${c.name}: your status says "${c.selfStatus}" but its tasks are ${c.computedPct}% done`,
        detail: 'Either the category row is ahead of the task detail, or the tasks have not been updated. Worth resolving before this goes into a report.',
      });
    }
  }

  /* ---- timeline: where this site is in its own calendar ----
     Weeks are counted against the site's own start and target, not the
     workbook's, because sites in different waves start months apart. */
  const firstWeek = weeks[0]?.end || null;
  const nowWeek = weeks[asAt]?.end || null;
  const weeksElapsed = asAt + 1;
  let weeksToTarget = null;
  let totalPlannedWeeks = null;
  if (site.target && nowWeek) {
    const d = diffDays(nowWeek, site.target);
    weeksToTarget = d === null ? null : Math.round(d / 7);
  }
  if (site.target && firstWeek) {
    const d = diffDays(firstWeek, site.target);
    totalPlannedWeeks = d === null ? null : Math.round(d / 7) + 1;
  }
  const timeline = {
    today: toISO(new Date()),
    startDate: site.start || firstWeek,
    targetSubmission: site.target,
    firstReportedWeek: weeks[0]?.label || null,
    reportingWeekLabel: weeks[asAt]?.label || null,
    reportingWeekEnding: nowWeek,
    weeksElapsed,
    weeksRemainingToTarget: weeksToTarget,
    totalPlannedWeeks,
    weekColumnsAvailable: weeks.length,
    percentOfPlannedTimeUsed: totalPlannedWeeks && totalPlannedWeeks > 0
      ? Math.round((weeksElapsed / totalPlannedWeeks) * 100) : null,
    previousAnalysis: opts.previous || null,
  };

  /* Rows for the weekly status grid: categories interleaved with their tasks,
     in template order, so the chart reads like the spreadsheet. */
  const gridTasks = [];
  const orderOf = (cid) => catMap.get(cid)?.order ?? 9999;
  const catIds = [...new Set([...d.categories.map((c) => c.id), ...tasks.map((t) => t.categoryId)])]
    .sort((x, y) => orderOf(x) - orderOf(y) || String(x).localeCompare(String(y)));
  for (const cid of catIds) {
    const cat = d.categories.find((c) => c.id === cid);
    if (cat) gridTasks.push({ id: cat.id, name: cat.name || cid, weekly: cat.weekly, isCategory: true });
    for (const t of tasks.filter((x) => x.categoryId === cid)) {
      gridTasks.push({ id: t.id, name: t.name, weekly: t.weekly, isCategory: false });
    }
  }

  return {
    ...base,
    noData: false,
    timeline,
    gridTasks,
    taskCount: tasks.length,
    liveCount: live.length,
    naCount,
    finishedCount: finished.length,
    wipCount: wip.length,
    blockedCount: blocked.length,
    waitingCount: waiting.length,
    notStartedCount: notStarted.length,
    stuckCount: blocked.length + waiting.length,
    pctByCount: Math.round(pct(finished.length, live.length) * 10) / 10,
    pctByWeight: Math.round(pct(wDone, wLive) * 10) / 10,
    totalWeight: wLive,
    doneWeight: wDone,
    remaining,
    remainingWeight,
    velocity,
    avgVelocity: Math.round(avgVelocity * 100) / 100,
    recentVelocity: Math.round(recentVelocity * 100) / 100,
    avgWeightVelocity: Math.round(avgWeightVelocity * 100) / 100,
    forecastAvg,
    forecastRecent,
    targetWeekNo,
    slipWeeks,
    forecastVsTarget,
    categories,
    stuckDetail,
    stalled,
    overdue,
    scopeGrowth,
    prereqs: d.prereqs || [],
    outstandingPrereqs,
    log: d.log || [],
    openLog: (d.log || []).filter((l) => !l.cleared),
    outOfOrder,
    unknownDeps,
    unknownPrereqRefs,
    resources,
    dataIssues,
    risks: risks.sort((a, b) => (a.level === 'high' ? -1 : 1) - (b.level === 'high' ? -1 : 1)),
    curve: buildCurve(live, weeks, asAt),
  };
}

/**
 * Cumulative completion curve. "Planned" is derived from each task's target
 * week, so it only exists where the user filled that column in — the app says
 * so rather than inventing a plan.
 */
function buildCurve(live, weeks, asAt) {
  const withTarget = live.filter((t) => t.targetWeek);
  const totalWeight = sum(live, (t) => t.weight || 1) || 1;
  const points = [];
  for (let i = 0; i < weeks.length; i++) {
    const doneW = sum(live.filter((t) => statusAt(t, i) === STATUS.FINISHED), (t) => t.weight || 1);
    const plannedW = sum(withTarget.filter((t) => t.targetWeek <= i + 1), (t) => t.weight || 1);
    points.push({
      weekIndex: i,
      label: weeks[i].label,
      end: weeks[i].end,
      actualPct: i <= asAt ? Math.round(pct(doneW, totalWeight) * 10) / 10 : null,
      plannedPct: withTarget.length ? Math.round(pct(plannedW, totalWeight) * 10) / 10 : null,
    });
  }
  return { points, hasPlan: withTarget.length > 0, planCoverage: Math.round(pct(withTarget.length, live.length)) };
}

/* ========================================================================
   Portfolio (master) analysis
   ======================================================================== */

export function analysePortfolio(model, opts = {}) {
  const sites = model.sites.map((s) => analyseSite(s, model.template, opts));
  const withData = sites.filter((s) => !s.noData);

  const totalTasks = sum(withData, (s) => s.liveCount);
  const totalDone = sum(withData, (s) => s.finishedCount);
  const totalWeight = sum(withData, (s) => s.totalWeight);
  const doneWeight = sum(withData, (s) => s.doneWeight);

  /* Resource load across every site — the thing a single-site view cannot see. */
  const load = new Map();
  for (const s of withData) {
    for (const r of s.resources) {
      if (!load.has(r.resource)) load.set(r.resource, { resource: r.resource, open: 0, wip: 0, stuck: 0, sites: [] });
      const e = load.get(r.resource);
      e.open += r.open; e.wip += r.wip; e.stuck += r.stuck;
      e.sites.push({ code: s.code, open: r.open });
    }
  }
  const resourceLoad = [...load.values()]
    .map((e) => ({ ...e, siteCount: e.sites.length }))
    .sort((a, b) => b.open - a.open);

  /* Blockers that recur across sites are systemic, not local. */
  const blockerText = new Map();
  for (const s of withData) {
    for (const b of s.stuckDetail) {
      const key = (b.waitingOn || b.reason || b.name || '').toLowerCase().slice(0, 60);
      if (!key) continue;
      if (!blockerText.has(key)) blockerText.set(key, { key, label: b.waitingOn || b.reason || b.name, sites: new Set(), count: 0 });
      const e = blockerText.get(key);
      e.sites.add(s.code); e.count++;
    }
  }
  const commonBlockers = [...blockerText.values()]
    .map((e) => ({ label: e.label, count: e.count, siteCount: e.sites.size, sites: [...e.sites] }))
    .filter((e) => e.siteCount > 1)
    .sort((a, b) => b.siteCount - a.siteCount || b.count - a.count);

  const waves = new Map();
  for (const s of sites) {
    const w = s.wave || '(no wave)';
    if (!waves.has(w)) waves.set(w, { wave: w, sites: [], pct: 0 });
    waves.get(w).sites.push(s.code);
  }

  const risks = [];
  const slipping = withData.filter((s) => s.slipWeeks !== null && s.slipWeeks > 0)
    .sort((a, b) => b.slipWeeks - a.slipWeeks);
  if (slipping.length) {
    risks.push({
      level: slipping.some((s) => s.slipWeeks > 4) ? 'high' : 'medium',
      code: 'portfolio_slip',
      title: `${slipping.length} of ${withData.length} sites forecast past their target submission`,
      detail: slipping.slice(0, 5).map((s) => `${s.code} +${s.slipWeeks}w`).join(', ') + '.',
    });
  }
  for (const r of resourceLoad) {
    if (r.siteCount > 1 && r.open >= (opts.overloadThreshold || 10)) {
      risks.push({
        level: 'medium', code: 'resource_spread',
        title: `${r.resource} holds ${r.open} open tasks across ${r.siteCount} sites`,
        detail: r.sites.map((s) => `${s.code} (${s.open})`).join(', ') + '. One person cannot be the constraint on several sites at once without one of them slipping.',
      });
    }
  }
  for (const b of commonBlockers) {
    risks.push({
      level: 'high', code: 'systemic_blocker',
      title: `The same blocker is holding up ${b.siteCount} sites`,
      detail: `"${b.label}" appears on ${b.sites.join(', ')}. A blocker on one site is a site problem; the same one on several is a process problem and needs fixing once, centrally.`,
    });
  }
  const noDetail = sites.filter((s) => s.noData);
  if (noDetail.length) {
    risks.push({
      level: 'medium', code: 'sites_untracked',
      title: `${noDetail.length} site${noDetail.length === 1 ? '' : 's'} in the register cannot be measured`,
      detail: noDetail.map((s) => s.code).join(', ') + ' — no detail sheet or no weekly status yet.',
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    file: model.file,
    siteCount: sites.length,
    sitesWithData: withData.length,
    totalTasks,
    totalDone,
    pctByCount: Math.round(pct(totalDone, totalTasks) * 10) / 10,
    pctByWeight: Math.round(pct(doneWeight, totalWeight) * 10) / 10,
    totalStuck: sum(withData, (s) => s.stuckCount),
    totalOverdue: sum(withData, (s) => s.overdue.length),
    totalAdditional: sum(withData, (s) => s.scopeGrowth.additionalCount),
    sites,
    slipping: slipping.map((s) => ({ code: s.code, slipWeeks: s.slipWeeks, target: s.target })),
    resourceLoad,
    commonBlockers,
    waves: [...waves.values()],
    risks,
  };
}

export { statusAt };
