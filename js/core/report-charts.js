/**
 * Figures for the standardised report.
 *
 * One renderer per figure the template calls for. All hand-rolled SVG, all
 * driven by the computed report — nothing here decides anything, it only draws
 * what report.js already worked out.
 */

import { el, mount, svgEl, clamp, fmtDate } from './util.js';

const RAG_FILL = { RED: 'var(--survey)', AMBER: 'var(--hivis)', GREEN: 'var(--sign)' };
const ragColour = (r) => RAG_FILL[r] || 'var(--conc)';

function chart(title, note, body, legend) {
  return el('div', { class: 'chartwrap' }, [
    title ? el('div', { class: 'charttitle', text: title }) : null,
    note ? el('div', { class: 'chartnote', text: note }) : null,
    el('div', { class: 'chartbox' }, [body]),
    legend || null,
  ]);
}
const empty = (t) => el('div', { class: 'small dim', style: { padding: '8px 0' }, text: t });

function text(x, y, s, opts = {}) {
  const t = svgEl('text', {
    x, y, 'font-size': opts.size || 10, fill: opts.fill || 'var(--ink-2)',
    'text-anchor': opts.anchor || 'start', 'font-weight': opts.weight || 400,
    'font-family': 'var(--font-ui)',
  });
  t.textContent = s;
  return t;
}

function ragLegend() {
  return el('div', { class: 'legend' }, ['RED', 'AMBER', 'GREEN'].map((r) => el('span', {}, [
    el('span', { class: 'swatch', style: { background: ragColour(r) } }), r,
  ])));
}

/* ============ 1. Site dashboard: status mix, progress vs time, RAG ============ */

export function renderKpiDashboard(host, c) {
  const m = c.metrics;
  const W = 760, H = 190;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Performance dashboard' });

  // panel 1 — task mix
  const parts = [
    ['Finished', m.finished, 'var(--st-done)'],
    ['WIP', m.wip, 'var(--st-wip)'],
    ['Stuck', m.stuck, 'var(--st-blocked)'],
    ['Not started', m.notStarted, 'var(--st-not)'],
  ].filter(([, n]) => (n || 0) > 0);
  const total = parts.reduce((n, [, v]) => n + v, 0) || 1;
  svg.appendChild(text(0, 12, 'Task status', { weight: 700, fill: 'var(--ink)' }));
  let x = 0;
  const barW = 240;
  for (const [label, n, fill] of parts) {
    const w = (n / total) * barW;
    svg.appendChild(svgEl('rect', { x, y: 22, width: w, height: 22, fill }));
    if (w > 22) svg.appendChild(text(x + w / 2, 37, String(n), { fill: '#fff', anchor: 'middle', weight: 600 }));
    x += w;
  }
  parts.forEach(([label, n, fill], i) => {
    svg.appendChild(svgEl('rect', { x: 0, y: 56 + i * 15, width: 9, height: 9, fill, rx: 2 }));
    svg.appendChild(text(14, 64 + i * 15, `${label} — ${n}`, { size: 9.5 }));
  });

  // panel 2 — completion against time
  const px = 290;
  svg.appendChild(text(px, 12, 'Progress against time', { weight: 700, fill: 'var(--ink)' }));
  const gauge = (y, label, value, fill) => {
    svg.appendChild(text(px, y + 9, label, { size: 9.5 }));
    svg.appendChild(svgEl('rect', { x: px + 92, y, width: 150, height: 12, rx: 6, fill: 'var(--paper-deep)' }));
    svg.appendChild(svgEl('rect', { x: px + 92, y, width: Math.max(2, (clamp(value ?? 0, 0, 100) / 100) * 150), height: 12, rx: 6, fill }));
    svg.appendChild(text(px + 250, y + 10, value == null ? '—' : `${value}%`, { size: 9.5, weight: 600 }));
  };
  gauge(24, 'Completion', m.completionWeighted, 'var(--sign)');
  gauge(44, 'Time elapsed', m.timeElapsed, 'var(--blueprint)');
  if (m.spiProxy != null) {
    svg.appendChild(text(px, 78, `Time-based SPI proxy ${m.spiProxy}`, {
      size: 10, weight: 700, fill: m.spiProxy >= 0.95 ? 'var(--sign)' : (m.spiProxy >= 0.85 ? 'var(--hivis)' : 'var(--survey)'),
    }));
  }

  // panel 3 — RAG by dimension
  svg.appendChild(text(0, 110, 'RAG by dimension', { weight: 700, fill: 'var(--ink)' }));
  const dims = Object.entries(c.dimensions).filter(([k]) => k !== 'overall');
  dims.forEach(([k, v], i) => {
    const bx = (i % 6) * 126;
    const by = 120;
    svg.appendChild(svgEl('rect', { x: bx, y: by, width: 118, height: 40, rx: 4, fill: ragColour(v.rag), opacity: 0.18, stroke: ragColour(v.rag) }));
    svg.appendChild(text(bx + 59, by + 17, k.replace(/([A-Z])/g, ' $1').replace(/^./, (s) => s.toUpperCase()), { size: 9, anchor: 'middle' }));
    svg.appendChild(text(bx + 59, by + 32, v.rag, { size: 12, anchor: 'middle', weight: 700, fill: ragColour(v.rag) }));
  });

  mount(host, chart('Performance dashboard',
    `Task status, progress against time, and RAG by dimension at ${fmtDate(m.statusDate)}.`, svg));
}

/* ============ 2. Milestone timeline (site) ============ */

export function renderMilestoneTimeline(host, c) {
  const rows = (c.milestones || []).filter((m) => m.baseline || /^\d{4}-/.test(String(m.actual || '')));
  if (!rows.length) { mount(host, chart('Milestones', '', empty('No dated milestones.'))); return; }

  const dates = rows.flatMap((m) => [m.baseline, /^\d{4}-/.test(String(m.actual)) ? m.actual : null]).filter(Boolean).sort();
  const min = dates[0], max = dates[dates.length - 1];
  const span = Math.max(1, (new Date(max) - new Date(min)) / 86400000);
  const W = 760, L = 150, R = 30, H = 40 + rows.length * 30;
  const px = (d) => L + (((new Date(d) - new Date(min)) / 86400000) / span) * (W - L - R);

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Milestones' });
  svg.appendChild(svgEl('line', { x1: L, y1: 22, x2: W - R, y2: 22, stroke: 'var(--rule-hard)' }));
  svg.appendChild(text(L, 14, fmtDate(min), { size: 9, fill: 'var(--ink-3)' }));
  svg.appendChild(text(W - R, 14, fmtDate(max), { size: 9, anchor: 'end', fill: 'var(--ink-3)' }));

  rows.forEach((m, i) => {
    const y = 40 + i * 30;
    svg.appendChild(text(0, y + 4, `${m.id} ${m.name}`.slice(0, 26), { size: 9.5 }));
    if (m.baseline) {
      svg.appendChild(svgEl('circle', { cx: px(m.baseline), cy: y, r: 5, fill: 'none', stroke: 'var(--blueprint)', 'stroke-width': 2 }));
    }
    if (/^\d{4}-/.test(String(m.actual))) {
      const ax = px(m.actual);
      svg.appendChild(svgEl('circle', { cx: ax, cy: y, r: 5, fill: ragColour(m.rag) }));
      if (m.baseline && m.actual !== m.baseline) {
        svg.appendChild(svgEl('line', { x1: px(m.baseline), y1: y, x2: ax, y2: y, stroke: ragColour(m.rag), 'stroke-width': 1.6, 'stroke-dasharray': '3 3' }));
      }
    }
  });

  mount(host, chart('Baseline against actual or forecast', 'Hollow = baseline, solid = actual or forecast.', svg, ragLegend()));
}

/* ============ 3. Category progress (site) ============ */

export function renderCategoryProgressReport(host, rows) {
  const live = (rows || []).filter((c) => c.liveCount > 0);
  if (!live.length) { mount(host, chart('Progress by category', '', empty('No categories with applicable tasks.'))); return; }
  const W = 760, rowH = 20, H = 20 + live.length * rowH;
  const L = 230, barW = W - L - 60;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Progress by category' });

  live.forEach((c, i) => {
    const y = 12 + i * rowH;
    svg.appendChild(text(0, y + 10, `${c.id} ${c.name}`.slice(0, 34), { size: 9.5 }));
    svg.appendChild(svgEl('rect', { x: L, y: y + 2, width: barW, height: 12, rx: 3, fill: 'var(--paper-deep)' }));
    const w = Math.max(1, (c.computedPct / 100) * barW);
    svg.appendChild(svgEl('rect', { x: L, y: y + 2, width: w, height: 12, rx: 3, fill: c.stuck ? 'var(--st-blocked)' : 'var(--sign)' }));
    if (c.stuck) {
      svg.appendChild(svgEl('rect', { x: L, y: y, width: barW, height: 16, rx: 3, fill: 'none', stroke: 'var(--survey)', 'stroke-width': 1.2 }));
    }
    svg.appendChild(text(W - 52, y + 12, `${c.doneCount}/${c.liveCount}`, { size: 9 }));
  });

  mount(host, chart('Task status by category', 'A red outline marks a category holding blocked or waiting work.', svg));
}

/* ============ 4. Constraint ageing (site) ============ */

export function renderConstraintAgeing(host, open) {
  if (!open?.length) { mount(host, chart('Age of open items', '', empty('No open waiting-on or blocked items.'))); return; }
  const max = Math.max(10, ...open.map((x) => x.age ?? 0));
  const W = 760, H = 24 + open.length * 24, L = 180, barW = W - L - 60;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Constraint ageing' });

  // thresholds at 5 and 10 days
  [[5, 'var(--hivis)'], [10, 'var(--survey)']].forEach(([d, col]) => {
    const x = L + (d / max) * barW;
    svg.appendChild(svgEl('line', { x1: x, y1: 8, x2: x, y2: H - 8, stroke: col, 'stroke-dasharray': '3 3', 'stroke-width': 1 }));
    svg.appendChild(text(x + 3, 14, `${d}d`, { size: 8, fill: col }));
  });

  open.forEach((l, i) => {
    const y = 20 + i * 24;
    svg.appendChild(text(0, y + 11, `${l.id} ${l.waitingOn || ''}`.slice(0, 26), { size: 9.5 }));
    const w = Math.max(2, ((l.age ?? 0) / max) * barW);
    svg.appendChild(svgEl('rect', { x: L, y: y + 2, width: w, height: 13, rx: 3, fill: ragColour(l.rag) }));
    svg.appendChild(text(L + w + 6, y + 13, l.age == null ? 'undated' : `${l.age} d`, { size: 9 }));
  });

  mount(host, chart('Age of open waiting-on / blocked items',
    'Measured in calendar days at the report date, against the 5 and 10 day thresholds.', svg, ragLegend()));
}

/* ============ 5. Risk heat map (both) ============ */

export function renderRiskHeatMap(host, risks) {
  if (!risks?.length) { mount(host, chart('Risk heat map', '', empty('No risks proposed.'))); return; }
  const cell = 46, pad = 34, W = pad + 5 * cell + 10, H = pad + 5 * cell + 24;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Risk heat map' });

  for (let p = 1; p <= 5; p++) {
    for (let i = 1; i <= 5; i++) {
      const score = p * i;
      const fill = score >= 15 ? 'var(--survey)' : (score >= 8 ? 'var(--hivis)' : 'var(--sign)');
      svg.appendChild(svgEl('rect', {
        x: pad + (i - 1) * cell, y: pad + (5 - p) * cell, width: cell - 2, height: cell - 2,
        fill, opacity: 0.16, stroke: 'var(--rule)',
      }));
    }
  }
  for (let i = 1; i <= 5; i++) {
    svg.appendChild(text(pad + (i - 1) * cell + cell / 2 - 1, H - 10, String(i), { size: 9, anchor: 'middle' }));
    svg.appendChild(text(pad - 8, pad + (5 - i) * cell + cell / 2 + 3, String(i), { size: 9, anchor: 'end' }));
  }
  svg.appendChild(text(pad + 2.5 * cell, H - 1, 'Impact →', { size: 9, anchor: 'middle', fill: 'var(--ink-3)' }));

  // group risks sharing a cell so labels never sit on top of each other
  const groups = new Map();
  for (const r of risks) {
    const key = `${r.probability}|${r.impact}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }
  for (const [key, list] of groups) {
    const [p, i] = key.split('|').map(Number);
    if (!p || !i) continue;
    const cx = pad + (i - 1) * cell + cell / 2 - 1;
    const cy = pad + (5 - p) * cell + cell / 2 - 1;
    svg.appendChild(svgEl('circle', { cx, cy, r: 13, fill: ragColour(list[0].rating), opacity: 0.92 }));
    const label = list.map((r) => r.id).join(',');
    svg.appendChild(text(cx, cy + 3.5, label.length > 8 ? `${list.length} risks` : label,
      { size: 9, anchor: 'middle', fill: '#fff', weight: 700 }));
    const ttl = svgEl('title');
    ttl.textContent = list.map((r) => `${r.id}: ${r.risk} (P${r.probability} x I${r.impact} = ${r.score})`).join('\n');
    svg.appendChild(ttl);
  }

  mount(host, chart('Risk heat map', 'Probability (vertical) x impact (horizontal). RED ≥ 15, AMBER 8–14, GREEN ≤ 7.', svg));
}

/* ============ 6. RAG matrix (master) ============ */

export function renderRagMatrix(host, matrix) {
  if (!matrix?.length) { mount(host, chart('RAG by site', '', empty('No sites.'))); return; }
  const dims = ['schedule', 'scope', 'constraints', 'resources', 'quality', 'dataIntegrity', 'overall'];
  const labels = ['Schedule', 'Scope', 'Constraints', 'Resources', 'Quality', 'Data', 'Overall'];
  const cw = 92, rh = 26, L = 70, T = 30;
  const W = L + dims.length * cw, H = T + matrix.length * rh + 26;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'RAG by site and dimension' });

  labels.forEach((l, i) => svg.appendChild(text(L + i * cw + cw / 2, T - 10, l, { size: 9, anchor: 'middle', weight: 600 })));
  matrix.forEach((row, r) => {
    const y = T + r * rh;
    svg.appendChild(text(0, y + 17, row.code, { size: 10, weight: 600, fill: 'var(--ink)' }));
    dims.forEach((d, i) => {
      const v = row[d];
      svg.appendChild(svgEl('rect', { x: L + i * cw, y: y + 2, width: cw - 3, height: rh - 5, rx: 3, fill: ragColour(v), opacity: d === 'overall' ? 0.95 : 0.22, stroke: ragColour(v) }));
      svg.appendChild(text(L + i * cw + (cw - 3) / 2, y + 17, v || '—',
        { size: 9, anchor: 'middle', weight: 700, fill: d === 'overall' ? '#fff' : ragColour(v) }));
    });
  });

  const reds = dims.map((d) => matrix.filter((m) => m[d] === 'RED').length);
  svg.appendChild(text(0, H - 8, 'RED count', { size: 9, fill: 'var(--ink-3)' }));
  reds.forEach((n, i) => svg.appendChild(text(L + i * cw + cw / 2, H - 8, String(n),
    { size: 9, anchor: 'middle', weight: 700, fill: n ? 'var(--survey)' : 'var(--ink-3)' })));

  mount(host, chart('RAG status of every selected site', 'Column totals count RED sites.', svg, ragLegend()));
}

/* ============ 7. Schedule across sites (master) ============ */

export function renderScheduleAcrossSites(host, rows) {
  const dated = (rows || []).filter((s) => s.target || s.forecast);
  if (!dated.length) { mount(host, chart('Schedule by site', '', empty('No dated sites.'))); return; }
  const dates = dated.flatMap((s) => [s.start, s.target, s.forecast]).filter(Boolean).sort();
  const min = dates[0], max = dates[dates.length - 1];
  const span = Math.max(1, (new Date(max) - new Date(min)) / 86400000);
  const L = 70, R = 30, W = 760, H = 34 + dated.length * 28;
  const px = (d) => L + (((new Date(d) - new Date(min)) / 86400000) / span) * (W - L - R);

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Schedule by site' });
  svg.appendChild(text(L, 12, fmtDate(min), { size: 9, fill: 'var(--ink-3)' }));
  svg.appendChild(text(W - R, 12, fmtDate(max), { size: 9, anchor: 'end', fill: 'var(--ink-3)' }));

  dated.forEach((s, i) => {
    const y = 30 + i * 28;
    svg.appendChild(text(0, y + 12, s.code, { size: 10, weight: 600, fill: 'var(--ink)' }));
    if (s.start && s.target) {
      svg.appendChild(svgEl('rect', { x: px(s.start), y: y + 4, width: Math.max(2, px(s.target) - px(s.start)), height: 11, rx: 3, fill: 'var(--blueprint)', opacity: 0.22 }));
    }
    if (s.target) svg.appendChild(svgEl('line', { x1: px(s.target), y1: y, x2: px(s.target), y2: y + 19, stroke: 'var(--blueprint)', 'stroke-width': 2 }));
    if (s.forecast) {
      svg.appendChild(svgEl('circle', { cx: px(s.forecast), cy: y + 9.5, r: 5, fill: ragColour(s.overall) }));
      if (s.target && s.variance > 0) {
        svg.appendChild(svgEl('line', { x1: px(s.target), y1: y + 9.5, x2: px(s.forecast), y2: y + 9.5, stroke: 'var(--survey)', 'stroke-width': 1.6, 'stroke-dasharray': '3 3' }));
      }
    }
  });

  mount(host, chart('Baseline, target and forecast by site',
    'Bar = start to target. Vertical line = target. Dot = actual or forecast submission.', svg, ragLegend()));
}

/* ============ 8. Completion vs time scatter (master) ============ */

export function renderCompletionVsTime(host, rows) {
  const pts = (rows || []).filter((s) => s.completionWeighted != null && s.timeElapsed != null);
  if (!pts.length) { mount(host, chart('Completion against time', '', empty('No site has both figures.'))); return; }
  const W = 480, H = 320, M = { t: 16, r: 16, b: 40, l: 46 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const x = (v) => M.l + (clamp(v, 0, 100) / 100) * iw;
  const y = (v) => M.t + ih - (clamp(v, 0, 100) / 100) * ih;

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Completion against time elapsed' });
  for (let v = 0; v <= 100; v += 25) {
    svg.appendChild(svgEl('line', { x1: M.l, y1: y(v), x2: W - M.r, y2: y(v), stroke: 'var(--rule-soft)' }));
    svg.appendChild(text(M.l - 6, y(v) + 3.5, `${v}%`, { size: 9, anchor: 'end', fill: 'var(--ink-3)' }));
    svg.appendChild(text(x(v), H - 22, `${v}%`, { size: 9, anchor: 'middle', fill: 'var(--ink-3)' }));
  }
  // on-plan diagonal
  svg.appendChild(svgEl('line', { x1: x(0), y1: y(0), x2: x(100), y2: y(100), stroke: 'var(--ink-2)', 'stroke-dasharray': '5 4' }));
  svg.appendChild(text(x(72), y(78), 'on plan', { size: 9, fill: 'var(--ink-3)' }));
  svg.appendChild(text(W / 2, H - 6, 'Time elapsed', { size: 9.5, anchor: 'middle', fill: 'var(--ink-2)' }));

  for (const s of pts) {
    svg.appendChild(svgEl('circle', { cx: x(s.timeElapsed), cy: y(s.completionWeighted), r: 6, fill: ragColour(s.overall), opacity: 0.9 }));
    svg.appendChild(text(x(s.timeElapsed) + 9, y(s.completionWeighted) + 3.5, s.code, { size: 9, weight: 600 }));
  }

  mount(host, chart('Completion against time elapsed',
    'Points below the dashed line are behind plan. Colour is the site\'s overall status.', svg, ragLegend()));
}

/* ============ 9. Category by site (master) ============ */

export function renderCategoryBySite(host, cats, sites) {
  const rows = (cats || []).filter((c) => c.applicable > 0);
  if (!rows.length || !sites?.length) { mount(host, chart('Completion by category and site', '', empty('Nothing to plot.'))); return; }
  const cw = Math.max(52, Math.min(96, Math.floor(620 / sites.length)));
  const L = 200, T = 28, rh = 22;
  const W = L + sites.length * cw, H = T + rows.length * rh + 10;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Completion by category and site' });

  sites.forEach((s, i) => svg.appendChild(text(L + i * cw + cw / 2, T - 10, s, { size: 9, anchor: 'middle', weight: 600 })));
  rows.forEach((c, r) => {
    const yy = T + r * rh;
    svg.appendChild(text(0, yy + 15, `${c.id} ${c.name}`.slice(0, 30), { size: 9 }));
    sites.forEach((s, i) => {
      const v = c.bySite?.[s];
      const stuck = (c.stuckSites || []).includes(s);
      svg.appendChild(svgEl('rect', {
        x: L + i * cw, y: yy + 3, width: cw - 3, height: rh - 7, rx: 3,
        fill: v == null ? 'var(--st-na)' : 'var(--sign)',
        opacity: v == null ? 0.5 : clamp(0.18 + (v / 100) * 0.8, 0.18, 0.98),
        stroke: stuck ? 'var(--survey)' : 'none', 'stroke-width': stuck ? 1.6 : 0,
      }));
      svg.appendChild(text(L + i * cw + (cw - 3) / 2, yy + 16, v == null ? 'N/A' : `${v}%`,
        { size: 8.5, anchor: 'middle', fill: v != null && v > 55 ? '#fff' : 'var(--ink-2)' }));
    });
  });

  mount(host, chart('Completion by category and site', 'A red outline marks a category holding blocked or waiting work.', svg));
}

/* ============ 10. Party dependency (master) ============ */

export function renderPartyDependency(host, parties) {
  if (!parties?.length) { mount(host, chart('Open items by party', '', empty('No external dependencies recorded.'))); return; }
  const rows = parties.slice(0, 12);
  const max = Math.max(...rows.map((p) => p.count), 1);
  const L = 140, W = 700, H = 16 + rows.length * 24, barW = W - L - 60;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Open items by party' });
  rows.forEach((p, i) => {
    const y = 10 + i * 24;
    svg.appendChild(text(0, y + 12, p.party.slice(0, 20), { size: 9.5, weight: 600 }));
    const w = Math.max(2, (p.count / max) * barW);
    svg.appendChild(svgEl('rect', { x: L, y: y + 3, width: w, height: 13, rx: 3, fill: 'var(--hivis)' }));
    svg.appendChild(text(L + w + 6, y + 14, `${p.count} (${p.sites.join(', ')})`, { size: 9 }));
  });
  mount(host, chart('Open items each party is holding', 'Site codes in brackets.', svg));
}

/* ============ 11. Resource across sites (master) ============ */

export function renderResourceAcrossSites(host, load) {
  if (!load?.length) { mount(host, chart('Resource assignment', '', empty('No resources assigned.'))); return; }
  const rows = load.slice(0, 14);
  const max = Math.max(...rows.map((r) => r.open), 1);
  const L = 90, W = 700, H = 16 + rows.length * 24, barW = W - L - 150;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Resource assignment' });
  rows.forEach((r, i) => {
    const y = 10 + i * 24;
    const col = r.load === 'CRITICAL' ? 'var(--survey)' : (r.load === 'HIGH' ? 'var(--hivis)' : 'var(--blueprint)');
    svg.appendChild(text(0, y + 12, r.resource, { size: 10, weight: 600, fill: 'var(--ink)' }));
    const w = Math.max(2, (r.open / max) * barW);
    svg.appendChild(svgEl('rect', { x: L, y: y + 3, width: w, height: 13, rx: 3, fill: col }));
    svg.appendChild(text(L + w + 6, y + 14, `${r.open} open · ${r.activeSites} active site${r.activeSites === 1 ? '' : 's'} · ${r.load}`, { size: 9 }));
  });
  mount(host, chart('Resource assignment across the selected sites',
    'Load: CRITICAL on three or more active sites, HIGH on two.', svg));
}

/* ============ 12. Findings by site (master) ============ */

export function renderFindingsBySite(host, patterns, sites) {
  if (!patterns?.length) { mount(host, chart('Findings by site', '', empty('No findings.'))); return; }
  const counts = (sites || []).map((s) => ({ site: s, n: patterns.filter((p) => p.sites.includes(s)).length }));
  const max = Math.max(...counts.map((c) => c.n), 1);
  const bw = Math.max(36, Math.min(80, Math.floor(600 / Math.max(counts.length, 1))));
  const W = 40 + counts.length * bw, H = 150;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Findings by site' });
  svg.appendChild(svgEl('line', { x1: 30, y1: H - 32, x2: W, y2: H - 32, stroke: 'var(--rule-hard)' }));
  counts.forEach((c, i) => {
    const h = (c.n / max) * (H - 60);
    const x = 34 + i * bw;
    svg.appendChild(svgEl('rect', { x, y: H - 32 - h, width: bw - 12, height: Math.max(h, 1), rx: 3, fill: c.n ? 'var(--hivis)' : 'var(--conc)' }));
    svg.appendChild(text(x + (bw - 12) / 2, H - 36 - h, String(c.n), { size: 9.5, anchor: 'middle', weight: 600 }));
    svg.appendChild(text(x + (bw - 12) / 2, H - 16, c.site, { size: 9, anchor: 'middle' }));
  });
  mount(host, chart('Data-integrity and quality findings by site', 'Count of distinct findings recorded against each site.', svg));
}

/* ============ 13. Intervention Priority Index (master) ============ */

export function renderIpi(host, ipi) {
  if (!ipi?.length) { mount(host, chart('Intervention Priority Index', '', empty('No active sites to rank.'))); return; }
  const rows = ipi.slice(0, 14);
  const L = 80, W = 700, H = 26 + rows.length * 26, barW = W - L - 60;
  const scale = 11;
  const parts = [['rag', 'var(--survey)', 'Overall RAG'], ['priority', 'var(--hivis)', 'Site priority'],
    ['time', 'var(--blueprint)', 'Time pressure'], ['gap', 'var(--plum)', 'Completion gap']];

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Intervention priority index' });
  rows.forEach((r, i) => {
    const y = 16 + i * 26;
    svg.appendChild(text(0, y + 12, r.code, { size: 10, weight: 600, fill: 'var(--ink)' }));
    let x = L;
    for (const [key, col] of parts) {
      const v = r.components[key] || 0;
      const w = (v / scale) * barW;
      if (w > 0.5) {
        svg.appendChild(svgEl('rect', { x, y: y + 2, width: w, height: 14, fill: col, opacity: 0.88 }));
        if (w > 18) svg.appendChild(text(x + w / 2, y + 13, String(v), { size: 8.5, anchor: 'middle', fill: '#fff', weight: 600 }));
      }
      x += w;
    }
    svg.appendChild(text(x + 6, y + 13, String(r.ipi), { size: 10, weight: 700 }));
  });

  mount(host, chart('Intervention Priority Index and its components',
    'Higher means intervene sooner. Maximum 11.', svg,
    el('div', { class: 'legend' }, parts.map(([, col, label]) => el('span', {}, [
      el('span', { class: 'swatch', style: { background: col } }), label,
    ])))));
}
