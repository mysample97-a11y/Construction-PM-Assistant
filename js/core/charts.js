/**
 * Charts, hand-rolled in SVG.
 *
 * No charting library: the dataset is small, and the two charts that matter
 * most here (the weekly status grid and the blocker-age bars) need behaviour a
 * generic library does not give — an honest "no plan line" state, a marker
 * separating recorded history from empty future columns, and status colouring
 * that matches the spreadsheet the reader is looking at.
 *
 * Every renderer takes a host element and paints into it, and every one is
 * safe to call with empty data.
 */

import { el, mount, svgEl, clamp, fmtDate } from './util.js';
import { STATUS } from './parser.js';

const STATUS_COLOUR = {
  [STATUS.NOT_STARTED]: 'var(--st-not)',
  [STATUS.WIP]:         'var(--st-wip)',
  [STATUS.BLOCKED]:     'var(--st-blocked)',
  [STATUS.WAITING]:     'var(--st-waiting)',
  [STATUS.FINISHED]:    'var(--st-done)',
  [STATUS.NA]:          'var(--st-na)',
};

export function statusLegend(extra = []) {
  const items = [
    [STATUS.FINISHED, 'Finished'], [STATUS.WIP, 'WIP'],
    [STATUS.BLOCKED, 'Blocked'], [STATUS.WAITING, 'Waiting on'],
    [STATUS.NOT_STARTED, 'Not started'], [STATUS.NA, 'N/A'],
  ];
  return el('div', { class: 'legend' }, [
    ...items.map(([k, label]) => el('span', {}, [
      el('span', { class: 'swatch', style: { background: STATUS_COLOUR[k] } }), label,
    ])),
    ...extra,
  ]);
}

function chart(title, note, body, legend) {
  return el('div', { class: 'chartwrap' }, [
    title ? el('div', { class: 'charttitle', text: title }) : null,
    note ? el('div', { class: 'chartnote', text: note }) : null,
    el('div', { class: 'chartbox' }, [body]),
    legend || null,
  ]);
}

function emptyNote(text) {
  return el('div', { class: 'small dim', style: { padding: '10px 0' }, text });
}

/* ============================================================
   1. Weekly status grid — the Gantt for this kind of data.

   A classic bar Gantt needs durations and dates per task. This programme is
   recorded as a status per task per week, so the honest equivalent is a grid:
   one row per task, one column per week, coloured by the status recorded that
   week. It shows exactly what the spreadsheet says, including the gaps.
   ============================================================ */

export function renderStatusGrid(host, analysis, opts = {}) {
  const weeks = analysis.reportingWeek ? (analysis.curve?.points || []) : [];
  const tasks = (opts.tasks || analysis.gridTasks || []);
  if (!tasks.length || !weeks.length) {
    mount(host, chart('Weekly status grid', '', emptyNote('No weekly status recorded yet.')));
    return;
  }

  const maxRows = opts.maxRows || 60;
  const shown = tasks.slice(0, maxRows);
  const LABEL = 250;
  const CELL = Math.max(14, Math.min(30, Math.floor(760 / weeks.length)));
  const ROW = 19;
  const HEAD = 34;
  const W = LABEL + weeks.length * CELL + 8;
  const H = HEAD + shown.length * ROW + 4;

  const svg = svgEl('svg', {
    viewBox: `0 0 ${W} ${H}`, width: W, height: H,
    role: 'img', 'aria-label': 'Task status by week',
  });

  // week headers
  weeks.forEach((w, i) => {
    const x = LABEL + i * CELL;
    const t = svgEl('text', {
      x: x + CELL / 2, y: 13, 'font-size': 9.5, fill: 'var(--ink-3)',
      'text-anchor': 'middle', 'font-family': 'var(--font-ui)',
    });
    t.textContent = w.label.replace(/^W0?/, 'W');
    svg.appendChild(t);
    if (i % 4 === 0) {
      const d = svgEl('text', {
        x: x + CELL / 2, y: 25, 'font-size': 8, fill: 'var(--ink-3)', 'text-anchor': 'middle',
      });
      d.textContent = fmtDate(w.end).slice(0, 6);
      svg.appendChild(d);
    }
  });
  svg.appendChild(svgEl('line', {
    x1: 0, y1: HEAD - 4, x2: W, y2: HEAD - 4, stroke: 'var(--rule)', 'stroke-width': 1,
  }));

  const lastIdx = analysis.reportingWeek ? weeks.findIndex((w) => w.label === analysis.reportingWeek.label) : -1;

  shown.forEach((t, r) => {
    const y = HEAD + r * ROW;
    const isCat = t.isCategory;

    if (isCat) {
      svg.appendChild(svgEl('rect', { x: 0, y: y - 1, width: W, height: ROW, fill: 'var(--blueprint-lo)' }));
    } else if (r % 2 === 1) {
      svg.appendChild(svgEl('rect', { x: 0, y: y - 1, width: W, height: ROW, fill: 'var(--sheet-alt)' }));
    }

    const label = svgEl('text', {
      x: isCat ? 6 : 16, y: y + 12,
      'font-size': isCat ? 10.5 : 10, fill: isCat ? 'var(--blueprint)' : 'var(--ink-2)',
      'font-weight': isCat ? 700 : 400, 'font-family': 'var(--font-ui)',
    });
    const name = `${t.id}  ${t.name}`;
    label.textContent = name.length > 38 ? `${name.slice(0, 37)}…` : name;
    const ttl = svgEl('title');
    ttl.textContent = name;
    label.appendChild(ttl);
    svg.appendChild(label);

    (t.weekly || []).slice(0, weeks.length).forEach((s, i) => {
      if (!s) return;
      const cell = svgEl('rect', {
        x: LABEL + i * CELL + 1, y: y + 2, width: CELL - 2, height: ROW - 6,
        rx: 2, fill: STATUS_COLOUR[s] || 'var(--st-not)',
        opacity: isCat ? 0.55 : 1,
      });
      const ct = svgEl('title');
      ct.textContent = `${t.id} — ${weeks[i].label}: ${s}`;
      cell.appendChild(ct);
      svg.appendChild(cell);
    });
  });

  // where recorded history stops
  if (lastIdx >= 0 && lastIdx < weeks.length - 1) {
    const x = LABEL + (lastIdx + 1) * CELL;
    svg.appendChild(svgEl('line', {
      x1: x, y1: HEAD - 8, x2: x, y2: H, stroke: 'var(--ink-2)',
      'stroke-width': 1.4, 'stroke-dasharray': '4 3',
    }));
  }

  const extra = tasks.length > maxRows
    ? [el('span', { class: 'dim', text: `showing ${maxRows} of ${tasks.length} rows` })]
    : [];
  mount(host, chart(
    'Weekly status grid',
    'One row per task, one column per week, coloured by the status you recorded. The dashed line is where reporting stops.',
    svg, statusLegend(extra),
  ));
}

/* ============================================================
   2. Cumulative progress against plan
   ============================================================ */

export function renderCurve(host, curve) {
  const pts = curve?.points || [];
  if (pts.length < 2) {
    mount(host, chart('Progress over time', '', emptyNote('Two or more weeks of status are needed before a trend can be drawn.')));
    return;
  }

  const W = 780, H = 250, M = { t: 14, r: 16, b: 34, l: 42 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const x = (i) => M.l + (pts.length === 1 ? 0 : (i / (pts.length - 1)) * iw);
  const y = (p) => M.t + ih - (clamp(p, 0, 100) / 100) * ih;

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Progress against plan' });

  for (let p = 0; p <= 100; p += 25) {
    svg.appendChild(svgEl('line', { x1: M.l, y1: y(p), x2: W - M.r, y2: y(p), stroke: 'var(--rule-soft)' }));
    const t = svgEl('text', { x: M.l - 7, y: y(p) + 3.5, 'font-size': 10, fill: 'var(--ink-3)', 'text-anchor': 'end' });
    t.textContent = `${p}%`;
    svg.appendChild(t);
  }

  const series = (key) => pts.map((p, i) => (p[key] == null ? null : `${x(i).toFixed(1)},${y(p[key]).toFixed(1)}`)).filter(Boolean);

  if (curve.hasPlan) {
    const pl = series('plannedPct');
    if (pl.length > 1) {
      svg.appendChild(svgEl('polyline', {
        points: pl.join(' '), fill: 'none', stroke: 'var(--blueprint)',
        'stroke-width': 2, 'stroke-dasharray': '6 4',
      }));
    }
  }
  const ac = series('actualPct');
  if (ac.length > 1) {
    svg.appendChild(svgEl('polygon', {
      points: `${M.l},${y(0)} ${ac.join(' ')} ${x(ac.length - 1)},${y(0)}`,
      fill: 'var(--sign)', opacity: 0.09,
    }));
    svg.appendChild(svgEl('polyline', { points: ac.join(' '), fill: 'none', stroke: 'var(--sign)', 'stroke-width': 2.6 }));
  }
  pts.forEach((p, i) => {
    if (p.actualPct == null) return;
    const c = svgEl('circle', { cx: x(i), cy: y(p.actualPct), r: 3.2, fill: 'var(--sign)', stroke: '#fff', 'stroke-width': 1.5 });
    const ttl = svgEl('title');
    ttl.textContent = `${p.label} (${fmtDate(p.end)}): ${p.actualPct}%`;
    c.appendChild(ttl);
    svg.appendChild(c);
  });

  const lastIdx = pts.reduce((n, p, i) => (p.actualPct != null ? i : n), -1);
  if (lastIdx >= 0 && lastIdx < pts.length - 1) {
    svg.appendChild(svgEl('line', {
      x1: x(lastIdx), y1: M.t, x2: x(lastIdx), y2: M.t + ih,
      stroke: 'var(--ink-2)', 'stroke-width': 1, 'stroke-dasharray': '3 3',
    }));
  }

  [0, pts.length - 1].forEach((i) => {
    const t = svgEl('text', {
      x: x(i), y: H - 12, 'font-size': 9.5, fill: 'var(--ink-3)',
      'text-anchor': i === 0 ? 'start' : 'end',
    });
    t.textContent = `${pts[i].label} ${fmtDate(pts[i].end)}`;
    svg.appendChild(t);
  });

  mount(host, chart('Progress over time', 'Share of weighted scope reported complete, week by week.', svg,
    el('div', { class: 'legend' }, [
      el('span', {}, [el('span', { class: 'swatch', style: { background: 'var(--sign)' } }), 'Reported complete']),
      curve.hasPlan
        ? el('span', {}, [el('span', { class: 'swatch', style: { background: 'var(--blueprint)' } }), `Plan from target weeks (${curve.planCoverage}% of tasks have one)`])
        : el('span', { class: 'dim', text: 'No plan line — no target weeks filled in' }),
    ])));
}

/* ============================================================
   3. Weekly throughput
   ============================================================ */

export function renderThroughput(host, velocity) {
  if (!velocity?.length) {
    mount(host, chart('Tasks finished each week', '', emptyNote('No weeks reported yet.')));
    return;
  }
  const W = 780, H = 165, M = { t: 14, r: 12, b: 30, l: 30 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const max = Math.max(1, ...velocity.map((v) => v.count));
  const bw = iw / velocity.length;

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Tasks finished per week' });
  svg.appendChild(svgEl('line', { x1: M.l, y1: M.t + ih, x2: W - M.r, y2: M.t + ih, stroke: 'var(--rule-hard)' }));

  const avg = velocity.reduce((n, v) => n + v.count, 0) / velocity.length;
  const ay = M.t + ih - (avg / max) * ih;
  svg.appendChild(svgEl('line', {
    x1: M.l, y1: ay, x2: W - M.r, y2: ay,
    stroke: 'var(--plum)', 'stroke-width': 1.3, 'stroke-dasharray': '5 4',
  }));
  const at = svgEl('text', { x: W - M.r, y: ay - 4, 'font-size': 9, fill: 'var(--plum)', 'text-anchor': 'end' });
  at.textContent = `avg ${Math.round(avg * 100) / 100}/wk`;
  svg.appendChild(at);

  velocity.forEach((v, i) => {
    const h = (v.count / max) * ih;
    const x = M.l + i * bw + bw * 0.18;
    const w = bw * 0.64;
    const zero = v.count === 0;
    const g = svgEl('g');
    g.appendChild(svgEl('rect', {
      x, y: zero ? M.t + ih - 3 : M.t + ih - h,
      width: w, height: zero ? 3 : Math.max(h, 2), rx: 2,
      fill: zero ? 'var(--survey)' : 'var(--blueprint)',
      opacity: zero ? 0.55 : 0.9,
    }));
    const t = svgEl('text', { x: x + w / 2, y: M.t + ih - (zero ? 8 : h + 4), 'font-size': 9, fill: 'var(--ink-2)', 'text-anchor': 'middle' });
    t.textContent = String(v.count);
    g.appendChild(t);
    if (bw > 24) {
      const lab = svgEl('text', { x: x + w / 2, y: H - 10, 'font-size': 8.5, fill: 'var(--ink-3)', 'text-anchor': 'middle' });
      lab.textContent = v.label.replace(/^W0?/, 'W');
      g.appendChild(lab);
    }
    const ttl = svgEl('title');
    ttl.textContent = `${v.label}: ${v.count} finished`;
    g.appendChild(ttl);
    svg.appendChild(g);
  });

  mount(host, chart('Tasks finished each week', 'Red marks a week where nothing completed — the shape of a stall.', svg));
}

/* ============================================================
   4. Status mix — one stacked bar
   ============================================================ */

export function renderStatusMix(host, a) {
  const parts = [
    { k: STATUS.FINISHED, n: a.finishedCount, label: 'Finished' },
    { k: STATUS.WIP, n: a.wipCount, label: 'WIP' },
    { k: STATUS.BLOCKED, n: a.blockedCount, label: 'Blocked' },
    { k: STATUS.WAITING, n: a.waitingCount, label: 'Waiting on' },
    { k: STATUS.NOT_STARTED, n: a.notStartedCount, label: 'Not started' },
  ].filter((p) => p.n > 0);
  const total = parts.reduce((n, p) => n + p.n, 0);
  if (!total) { mount(host, chart('Where the work stands', '', emptyNote('No live tasks.'))); return; }

  const W = 780, H = 52;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Status mix' });
  let x = 0;
  for (const p of parts) {
    const w = (p.n / total) * W;
    const g = svgEl('g');
    g.appendChild(svgEl('rect', { x, y: 0, width: w, height: 26, fill: STATUS_COLOUR[p.k] }));
    if (w > 34) {
      const t = svgEl('text', { x: x + w / 2, y: 18, 'font-size': 11, fill: '#fff', 'text-anchor': 'middle', 'font-weight': 600 });
      t.textContent = String(p.n);
      g.appendChild(t);
    }
    if (w > 62) {
      const l = svgEl('text', { x: x + w / 2, y: 42, 'font-size': 9.5, fill: 'var(--ink-2)', 'text-anchor': 'middle' });
      l.textContent = p.label;
      g.appendChild(l);
    }
    const ttl = svgEl('title');
    ttl.textContent = `${p.label}: ${p.n} of ${total}`;
    g.appendChild(ttl);
    svg.appendChild(g);
    x += w;
  }
  mount(host, chart('Where the work stands', `${total} live tasks, excluding anything marked N/A.`, svg, statusLegend()));
}

/* ============================================================
   5. Category progress
   ============================================================ */

export function renderCategoryProgress(host, categories) {
  const cats = (categories || []).filter((c) => c.liveCount > 0);
  if (!cats.length) { mount(host, chart('Progress by category', '', emptyNote('No categories with live tasks.'))); return; }

  const rows = cats.map((c) => el('div', { class: 'row', style: { gap: '10px', padding: '3px 0' } }, [
    el('span', { class: 'xs', style: { flex: '0 0 46px', fontFamily: 'var(--font-data)', color: 'var(--ink-3)' }, text: c.id }),
    el('span', { class: 'small truncate', style: { flex: '0 0 175px' }, text: c.name }),
    el('div', { class: 'grow' }, [
      el('div', { class: 'meter' }, [
        el('div', {
          class: 'meter__fill',
          dataset: { tone: c.computedPct >= 100 ? 'sign' : (c.stuck ? 'survey' : 'blueprint') },
          style: { width: `${clamp(c.computedPct, 0, 100)}%` },
        }),
      ]),
    ]),
    el('span', { class: 'xs num', style: { flex: '0 0 74px', textAlign: 'right' }, text: `${c.doneCount}/${c.liveCount}` }),
    el('span', {
      class: 'xs num', style: { flex: '0 0 42px', textAlign: 'right', color: c.stuck ? 'var(--survey)' : 'var(--ink-3)' },
      text: c.stuck ? `${c.stuck} stuck` : '',
    }),
  ]));

  mount(host, chart('Progress by category',
    'Computed from the tasks under each category, which may differ from the status you recorded on the category row.',
    el('div', {}, rows)));
}

/* ============================================================
   6. Blocker ages
   ============================================================ */

export function renderBlockerAges(host, stuck) {
  if (!stuck?.length) {
    mount(host, chart('How long work has been stuck', '', emptyNote('Nothing is blocked or waiting.')));
    return;
  }
  const max = Math.max(...stuck.map((s) => s.weeksStuck), 1);
  const rows = stuck.slice(0, 14).map((s) => el('div', { class: 'row', style: { gap: '10px', padding: '3px 0' } }, [
    el('span', { class: 'xs', style: { flex: '0 0 74px', fontFamily: 'var(--font-data)' }, text: s.taskId }),
    el('span', { class: 'small truncate', style: { flex: '0 0 165px' }, text: s.name }),
    el('div', { class: 'grow' }, [
      el('div', { class: 'meter' }, [
        el('div', {
          class: 'meter__fill',
          style: {
            width: `${(s.weeksStuck / max) * 100}%`,
            background: s.status === 'Blocked' ? 'var(--st-blocked)' : 'var(--st-waiting)',
          },
        }),
      ]),
    ]),
    el('span', { class: 'xs num', style: { flex: '0 0 54px', textAlign: 'right' }, text: `${s.weeksStuck}w` }),
    el('span', { class: 'xs truncate', style: { flex: '0 0 100px', color: 'var(--ink-3)' }, text: s.waitingOn || '—' }),
  ]));
  mount(host, chart('How long work has been stuck',
    'Longest first. The rightmost column is who it is waiting on.',
    el('div', {}, rows),
    el('div', { class: 'legend' }, [
      el('span', {}, [el('span', { class: 'swatch', style: { background: 'var(--st-blocked)' } }), 'Blocked — inside your control']),
      el('span', {}, [el('span', { class: 'swatch', style: { background: 'var(--st-waiting)' } }), 'Waiting on — outside it']),
    ])));
}

/* ============================================================
   7. Resource load
   ============================================================ */

export function renderResourceLoad(host, resources, opts = {}) {
  const rs = (resources || []).slice(0, 12);
  if (!rs.length) { mount(host, chart('Open work by resource', '', emptyNote('No resources assigned.'))); return; }
  const max = Math.max(...rs.map((r) => r.open), 1);
  const rows = rs.map((r) => el('div', { class: 'row', style: { gap: '10px', padding: '3px 0' } }, [
    el('span', { class: 'small strong', style: { flex: '0 0 74px' }, text: r.resource }),
    el('div', { class: 'grow' }, [
      el('div', { class: 'meter' }, [
        el('div', { class: 'meter__fill', dataset: { tone: 'blueprint' }, style: { width: `${(r.open / max) * 100}%` } }),
        r.stuck ? el('div', { class: 'meter__fill', style: { width: `${(r.stuck / max) * 100}%`, background: 'var(--st-blocked)' } }) : null,
      ]),
    ]),
    el('span', { class: 'xs num', style: { flex: '0 0 60px', textAlign: 'right' }, text: `${r.open} open` }),
    opts.showSites && r.siteCount
      ? el('span', { class: 'xs', style: { flex: '0 0 110px', color: 'var(--ink-3)' }, text: `${r.siteCount} site${r.siteCount === 1 ? '' : 's'}` })
      : null,
  ]));
  mount(host, chart('Open work by resource', 'Red portion is work that is blocked or waiting.', el('div', {}, rows)));
}

/* ============================================================
   8. Site comparison (portfolio)
   ============================================================ */

export function renderSiteBars(host, sites) {
  const rows = (sites || []).map((s) => el('div', { class: 'row', style: { gap: '10px', padding: '4px 0' } }, [
    el('span', { class: 'small strong', style: { flex: '0 0 74px' }, text: s.code }),
    el('div', { class: 'grow' }, [
      el('div', { class: 'meter' }, [
        el('div', {
          class: 'meter__fill',
          dataset: { tone: s.noData ? 'conc' : (s.slipWeeks > 0 ? 'survey' : 'sign') },
          style: { width: `${s.noData ? 0 : clamp(s.pctByWeight, 0, 100)}%` },
        }),
      ]),
    ]),
    el('span', { class: 'small num', style: { flex: '0 0 46px', textAlign: 'right' }, text: s.noData ? '—' : `${Math.round(s.pctByWeight)}%` }),
    el('span', {
      class: 'xs num', style: { flex: '0 0 86px', textAlign: 'right', color: s.slipWeeks > 0 ? 'var(--survey)' : 'var(--ink-3)' },
      text: s.noData ? 'no data' : (s.slipWeeks == null ? '—' : (s.slipWeeks > 0 ? `+${s.slipWeeks}w late` : `${Math.abs(s.slipWeeks)}w spare`)),
    }),
  ]));
  mount(host, chart('Sites compared', 'Weighted completion, with forecast against each target submission.', el('div', {}, rows)));
}

export function meter(actualPct, plannedPct) {
  const p = clamp(Number(actualPct) || 0, 0, 100);
  const tone = plannedPct == null ? 'blueprint'
    : (p - plannedPct >= -2 ? 'sign' : (p - plannedPct >= -12 ? 'hivis' : 'survey'));
  const bar = el('div', { class: 'meter' }, [
    el('div', { class: 'meter__fill', dataset: { tone }, style: { width: `${p}%` } }),
  ]);
  if (plannedPct != null) {
    bar.appendChild(el('div', { class: 'meter__plan', style: { left: `${clamp(plannedPct, 0, 100)}%` } }));
  }
  return bar;
}

/**
 * Renders every chart for a site into a detached element and returns the HTML.
 * Used by the PDF export so the printed report carries the same visuals as the
 * screen rather than a table of numbers.
 */
export function chartsToHTML(analysis) {
  const box = document.createElement('div');
  const add = (fn, ...args) => {
    const h = document.createElement('div');
    box.appendChild(h);
    try { fn(h, ...args); } catch { h.remove(); }
  };
  add(renderStatusMix, analysis);
  add(renderCurve, analysis.curve);
  add(renderThroughput, analysis.velocity);
  add(renderCategoryProgress, analysis.categories);
  if (analysis.stuckDetail?.length) add(renderBlockerAges, analysis.stuckDetail);
  add(renderResourceLoad, analysis.resources);
  add(renderStatusGrid, analysis, { maxRows: 40 });
  return box.innerHTML;
}
