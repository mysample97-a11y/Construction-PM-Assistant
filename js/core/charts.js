/**
 * Charts, hand-rolled in SVG. The dataset is small enough that a charting
 * library would cost more than it saves, and these need behaviour a generic
 * one does not give: an honest "no plan line" state, and a reporting-week
 * marker that separates recorded history from empty future columns.
 */

import { el, mount, svgEl, clamp, fmtDate } from './util.js';

export function meter(actualPct, plannedPct) {
  const p = clamp(Number(actualPct) || 0, 0, 100);
  const tone = plannedPct == null ? 'blueprint'
    : (p - plannedPct >= -2 ? 'sign' : (p - plannedPct >= -12 ? 'hivis' : 'survey'));
  const bar = el('div', { class: 'meter' }, [
    el('div', { class: 'meter__fill', dataset: { tone }, style: { width: `${p}%` } }),
  ]);
  if (plannedPct != null) {
    bar.appendChild(el('div', {
      class: 'meter__plan', style: { left: `${clamp(plannedPct, 0, 100)}%` },
      title: `Planned ${Math.round(plannedPct)}%`,
    }));
  }
  return bar;
}

/** Cumulative completion against the plan implied by target weeks. */
export function renderCurve(host, curve, opts = {}) {
  const pts = curve?.points || [];
  if (pts.length < 2) {
    mount(host, el('div', { class: 'empty' }, [
      el('h3', { text: 'Not enough weeks to draw a curve' }),
      el('p', { text: 'Two or more weeks of status are needed before a trend can be shown.' }),
    ]));
    return;
  }

  const W = 760, H = 240, M = { t: 14, r: 14, b: 34, l: 38 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const x = (i) => M.l + (pts.length === 1 ? 0 : (i / (pts.length - 1)) * iw);
  const y = (p) => M.t + ih - (clamp(p, 0, 100) / 100) * ih;

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Progress against plan' });

  for (let p = 0; p <= 100; p += 25) {
    svg.appendChild(svgEl('line', { x1: M.l, y1: y(p), x2: W - M.r, y2: y(p), stroke: 'var(--rule-soft)', 'stroke-width': 1 }));
    const t = svgEl('text', { x: M.l - 6, y: y(p) + 3.5, 'font-size': 10, fill: 'var(--ink-3)', 'text-anchor': 'end' });
    t.textContent = `${p}%`;
    svg.appendChild(t);
  }

  const line = (key) => pts
    .map((p, i) => (p[key] == null ? null : `${x(i).toFixed(1)},${y(p[key]).toFixed(1)}`))
    .filter(Boolean);

  if (curve.hasPlan) {
    const pl = line('plannedPct');
    if (pl.length > 1) {
      svg.appendChild(svgEl('polyline', {
        points: pl.join(' '), fill: 'none', stroke: 'var(--blueprint)',
        'stroke-width': 2, 'stroke-dasharray': '5 4',
      }));
    }
  }

  const ac = line('actualPct');
  if (ac.length > 1) {
    svg.appendChild(svgEl('polyline', { points: ac.join(' '), fill: 'none', stroke: 'var(--sign)', 'stroke-width': 2.4 }));
  }
  pts.forEach((p, i) => {
    if (p.actualPct == null) return;
    const c = svgEl('circle', { cx: x(i), cy: y(p.actualPct), r: 3, fill: 'var(--sign)', stroke: 'var(--sheet)', 'stroke-width': 1.4 });
    const ttl = svgEl('title');
    ttl.textContent = `${p.label} (${fmtDate(p.end)}): ${p.actualPct}%`;
    c.appendChild(ttl);
    svg.appendChild(c);
  });

  // Where recorded history stops.
  const lastIdx = pts.reduce((n, p, i) => (p.actualPct != null ? i : n), -1);
  if (lastIdx >= 0 && lastIdx < pts.length - 1) {
    svg.appendChild(svgEl('line', {
      x1: x(lastIdx), y1: M.t, x2: x(lastIdx), y2: M.t + ih,
      stroke: 'var(--ink-2)', 'stroke-width': 1, 'stroke-dasharray': '3 3',
    }));
    const t = svgEl('text', { x: x(lastIdx) + 4, y: M.t + 10, 'font-size': 9, fill: 'var(--ink-2)' });
    t.textContent = 'reported to here';
    svg.appendChild(t);
  }

  const label = (i) => {
    const t = svgEl('text', { x: x(i), y: H - 12, 'font-size': 9, fill: 'var(--ink-3)', 'text-anchor': i === 0 ? 'start' : 'end' });
    t.textContent = `${pts[i].label} ${fmtDate(pts[i].end)}`;
    return t;
  };
  svg.appendChild(label(0));
  svg.appendChild(label(pts.length - 1));

  mount(host,
    el('div', { class: 'chartbox' }, [svg]),
    el('div', { class: 'gantt-legend', style: { marginTop: '6px' } }, [
      el('span', {}, [el('span', { class: 'dkey', style: { background: 'var(--sign)' } }), 'Reported complete']),
      curve.hasPlan
        ? el('span', {}, [el('span', { class: 'dkey', style: { background: 'var(--blueprint)' } }), `Plan from target weeks (${curve.planCoverage}% of tasks have one)`])
        : el('span', { class: 'dim', text: 'No plan line — no target weeks filled in' }),
    ]),
  );
}

/** Tasks finished each week. Zero weeks are the interesting ones. */
export function renderThroughput(host, velocity) {
  if (!velocity?.length) { mount(host, el('div', { class: 'dim small', text: 'No weeks reported yet.' })); return; }
  const W = 760, H = 150, M = { t: 12, r: 10, b: 28, l: 28 };
  const iw = W - M.l - M.r, ih = H - M.t - M.b;
  const max = Math.max(1, ...velocity.map((v) => v.count));
  const bw = iw / velocity.length;

  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Tasks finished per week' });
  svg.appendChild(svgEl('line', { x1: M.l, y1: M.t + ih, x2: W - M.r, y2: M.t + ih, stroke: 'var(--rule-hard)' }));

  velocity.forEach((v, i) => {
    const h = (v.count / max) * ih;
    const x = M.l + i * bw + bw * 0.16;
    const w = bw * 0.68;
    const g = svgEl('g');
    g.appendChild(svgEl('rect', {
      x, y: M.t + ih - h, width: w, height: Math.max(h, v.count === 0 ? 0 : 2),
      fill: v.count === 0 ? 'var(--survey)' : 'var(--blueprint)',
      opacity: v.count === 0 ? 0.25 : 0.85, rx: 2,
    }));
    if (v.count === 0) {
      g.appendChild(svgEl('rect', { x, y: M.t + ih - 3, width: w, height: 3, fill: 'var(--survey)', rx: 1 }));
    }
    const t = svgEl('text', { x: x + w / 2, y: M.t + ih - h - 4, 'font-size': 9, fill: 'var(--ink-2)', 'text-anchor': 'middle' });
    t.textContent = String(v.count);
    g.appendChild(t);
    if (bw > 26) {
      const lab = svgEl('text', { x: x + w / 2, y: H - 10, 'font-size': 8.5, fill: 'var(--ink-3)', 'text-anchor': 'middle' });
      lab.textContent = v.label;
      g.appendChild(lab);
    }
    const ttl = svgEl('title');
    ttl.textContent = `${v.label}: ${v.count} finished`;
    g.appendChild(ttl);
    svg.appendChild(g);
  });

  mount(host, el('div', { class: 'chartbox' }, [svg]));
}

/** Horizontal comparison of sites. */
export function renderSiteBars(host, sites) {
  const rows = sites.map((s) => el('div', { class: 'row', style: { gap: '10px', padding: '4px 0' } }, [
    el('span', { class: 'small strong', style: { flex: '0 0 74px' }, text: s.code }),
    el('div', { class: 'grow' }, [meter(s.noData ? 0 : s.pctByWeight, null)]),
    el('span', { class: 'small num', style: { flex: '0 0 44px', textAlign: 'right' }, text: s.noData ? '—' : `${Math.round(s.pctByWeight)}%` }),
    el('span', {
      class: 'xs num', style: { flex: '0 0 74px', textAlign: 'right', color: s.slipWeeks > 0 ? 'var(--survey)' : 'var(--ink-3)' },
      text: s.noData ? 'no data' : (s.slipWeeks == null ? '—' : (s.slipWeeks > 0 ? `+${s.slipWeeks}w late` : `${Math.abs(s.slipWeeks)}w spare`)),
    }),
  ]));
  mount(host, ...rows);
}
