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

/* ================================ Excel ================================ */

export function buildWorkbook(portfolio, reports) {
  if (!window.XLSX) throw new Error('The spreadsheet writer did not load.');
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
    ['Sites in register', portfolio.siteCount],
    ['Sites with weekly data', portfolio.sitesWithData],
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
      default: return `\\pard\\sa120 ${t}\\par `;
    }
  }).join('');
  return `${head}${body}}`;
}

export function reportBlocks(portfolio, reports) {
  const b = [];
  b.push({ style: 'h1', text: 'BIM Multi-Site Delivery — Analysis Report' });
  b.push({ style: 'meta', text: `Generated ${new Date().toLocaleString()} · source: ${portfolio.file || 'workbook'} · sites: ${portfolio.siteCount}` });
  b.push({ style: 'meta', text: 'All figures were computed by the application from the weekly status grid. Narrative sections were written by an AI model reading those figures; it did not calculate them.' });
  b.push({ style: 'rule', text: '' });

  b.push({ style: 'h2', text: 'Programme position' });
  b.push({ style: 'p', text: `${portfolio.sitesWithData} of ${portfolio.siteCount} sites have weekly data. ${portfolio.totalDone} of ${portfolio.totalTasks} live tasks are finished — ${portfolio.pctByWeight}% by weight, ${portfolio.pctByCount}% by count. ${portfolio.totalStuck} tasks are blocked or waiting on someone. ${portfolio.totalOverdue} are past their target week. ${portfolio.totalAdditional} were added after kickoff.` });

  if (portfolio.risks.length) {
    b.push({ style: 'h3', text: 'Computed risks' });
    for (const r of portfolio.risks) b.push({ style: 'bullet', text: `[${r.level.toUpperCase()}] ${r.title} — ${r.detail}` });
  }

  const master = reports?.__master__;
  if (master) {
    b.push({ style: 'h2', text: 'Programme review' });
    b.push(...aiBlocks(master.result));
  }

  for (const s of portfolio.sites) {
    b.push({ style: 'rule', text: '' });
    b.push({ style: 'h2', text: `Site ${s.code}${s.description ? ` — ${s.description}` : ''}` });
    if (s.noData) {
      b.push({ style: 'p', text: 'No weekly data for this site. It appears in the register only.' });
    } else {
      b.push({ style: 'p', text: `${s.finishedCount} of ${s.liveCount} tasks finished (${s.pctByWeight}% by weight). Recent rate ${s.recentVelocity} tasks/week. ${s.remaining} remaining. ${s.forecastRecent ? `At that rate about ${s.forecastRecent.weeksNeeded} more weeks, landing near ${s.forecastRecent.finishDate || `week ${s.forecastRecent.finishWeek}`}.` : 'No completions recorded, so no forecast is possible.'}${s.slipWeeks !== null ? ` That is ${s.slipWeeks > 0 ? `${s.slipWeeks} weeks past` : `${Math.abs(s.slipWeeks)} weeks inside`} the target of ${s.target}.` : ''}` });
      if (s.stuckDetail.length) {
        b.push({ style: 'h3', text: 'Blocked and waiting' });
        for (const x of s.stuckDetail) {
          b.push({ style: 'bullet', text: `${x.taskId} ${x.name} — ${x.status} for ${x.weeksStuck} week(s)${x.waitingOn ? `, waiting on ${x.waitingOn}` : ''}${x.reason ? `: ${x.reason}` : ' (no reason logged)'}` });
        }
      }
      if (s.risks.length) {
        b.push({ style: 'h3', text: 'Computed risks' });
        for (const r of s.risks) b.push({ style: 'bullet', text: `[${r.level.toUpperCase()}] ${r.title} — ${r.detail}` });
      }
    }
    const rep = reports?.[s.code];
    if (rep) {
      b.push({ style: 'h3', text: 'Review' });
      b.push(...aiBlocks(rep.result));
    }
  }
  return b;
}

function aiBlocks(r) {
  const b = [];
  if (!r) return b;
  if (r.headline) b.push({ style: 'h3', text: r.headline });
  if (r.verdict) b.push({ style: 'meta', text: `Verdict: ${String(r.verdict).replace(/_/g, ' ')}${r.confidence ? ` · confidence ${r.confidence}` : ''}${r.confidenceReason ? ` — ${r.confidenceReason}` : ''}` });
  if (r.summary) b.push({ style: 'p', text: r.summary });
  const list = (title, arr, fmt) => {
    if (!Array.isArray(arr) || !arr.length) return;
    b.push({ style: 'h3', text: title });
    for (const x of arr) b.push({ style: 'bullet', text: typeof x === 'string' ? x : fmt(x) });
  };
  list('What is driving it', r.whatIsDrivingIt, (x) => `${x.point} — ${x.evidence}${x.effect ? ` (${x.effect})` : ''}`);
  list('Bottlenecks', r.bottlenecks, (x) => `${x.taskId || ''} ${x.issue}${x.whoToChase ? ` — chase ${x.whoToChase}` : ''}${x.suggestedAction ? `. ${x.suggestedAction}` : ''}`);
  list('Site ranking', r.siteRanking, (x) => `${x.site}: ${x.standing} — ${x.why}`);
  list('Systemic issues', r.systemicIssues, (x) => `${x.issue} (${(x.sitesAffected || []).join(', ')}) — ${x.rootCauseHypothesis}${x.fixOnceCentrally ? `. Fix: ${x.fixOnceCentrally}` : ''}`);
  list('Resource concerns', r.resourceConcerns, (x) => `${x.resource}: ${x.concern} — ${x.suggestedAction}`);
  list('Sequencing and overlaps', r.sequencingAndOverlaps, (x) => x);
  list('Actions', r.actions || r.priorityActions, (x) => `[${x.priority || 'action'}] ${x.action}${x.owner ? ` — ${x.owner}` : ''}${x.byWhen ? ` by ${x.byWhen}` : ''}${x.expectedEffect ? `. ${x.expectedEffect}` : ''}`);
  list('Watch next week', r.watchNextWeek, (x) => x);
  list('Going well', r.whatIsGoingWell, (x) => x);
  list('Data gaps', r.dataGaps, (x) => x);
  return b;
}

export function exportWord(portfolio, reports) {
  const rtf = buildRTF(reportBlocks(portfolio, reports));
  const stamp = new Date().toISOString().slice(0, 10);
  downloadBlob(new Blob([rtf], { type: 'application/rtf' }), `bim-tracker-report-${stamp}.rtf`);
}

/* ================================= PDF ================================= */

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export function buildPrintHTML(portfolio, reports) {
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
      default: return `<p>${t}</p>`;
    }
  }).join('\n')
    // Wrap runs of <li> so bullets render as real lists.
    .replace(/(<li>[\s\S]*?<\/li>)(?!\s*<li>)/g, (m) => `<ul>${m}</ul>`)
    .replace(/<\/ul>\s*<ul>/g, '');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>BIM Multi-Site Delivery — Analysis Report</title>
<style>
  @page { margin: 18mm 16mm; }
  body { font: 11pt/1.5 Arial, Helvetica, sans-serif; color: #12212F; max-width: 175mm; }
  h1 { font-size: 19pt; margin: 0 0 4pt; }
  h2 { font-size: 14pt; margin: 18pt 0 6pt; border-bottom: 1px solid #C6D0DA; padding-bottom: 3pt; page-break-after: avoid; }
  h3 { font-size: 11.5pt; margin: 12pt 0 4pt; page-break-after: avoid; }
  p { margin: 0 0 7pt; }
  .meta { font-size: 9pt; color: #5A6B7B; font-style: italic; }
  ul { margin: 0 0 8pt; padding-left: 16pt; }
  li { margin-bottom: 3pt; page-break-inside: avoid; }
  hr { border: 0; border-top: 1px solid #C6D0DA; margin: 14pt 0; }
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
