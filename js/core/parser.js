/**
 * Workbook parser.
 *
 * Reads the three-part tracker: a Sites register, a Task Template, and one
 * detail sheet per site holding three tables (task list with weekly status,
 * prerequisites, waiting-on log).
 *
 * Nothing here guesses silently. Every assumption the parser makes is recorded
 * in `notes` or `problems` and shown to the user for confirmation before a
 * single figure is computed.
 */

import { toISO, toNum, norm, uid } from './util.js';

export const STATUS = {
  NOT_STARTED: 'Not started',
  WIP: 'WIP',
  BLOCKED: 'Blocked',
  WAITING: 'Waiting on',
  FINISHED: 'Finished',
  NA: 'N/A',
};

export const STATUS_ORDER = [
  STATUS.NOT_STARTED, STATUS.WIP, STATUS.BLOCKED, STATUS.WAITING, STATUS.FINISHED, STATUS.NA,
];

const STATUS_ALIASES = {
  notstarted: STATUS.NOT_STARTED, notyetstarted: STATUS.NOT_STARTED, pending: STATUS.NOT_STARTED,
  ns: STATUS.NOT_STARTED, todo: STATUS.NOT_STARTED, planned: STATUS.NOT_STARTED,
  wip: STATUS.WIP, inprogress: STATUS.WIP, ongoing: STATUS.WIP, started: STATUS.WIP, active: STATUS.WIP,
  blocked: STATUS.BLOCKED, stopped: STATUS.BLOCKED, onhold: STATUS.BLOCKED, hold: STATUS.BLOCKED,
  waitingon: STATUS.WAITING, waiting: STATUS.WAITING, awaiting: STATUS.WAITING, held: STATUS.WAITING,
  finished: STATUS.FINISHED, complete: STATUS.FINISHED, completed: STATUS.FINISHED,
  done: STATUS.FINISHED, closed: STATUS.FINISHED,
  na: STATUS.NA, notapplicable: STATUS.NA, notrequired: STATUS.NA, excluded: STATUS.NA,
};

export function normStatus(v) {
  if (v === null || v === undefined || String(v).trim() === '') return null;
  return STATUS_ALIASES[norm(v)] || null;
}

/** Statuses that mean "cannot proceed" — the two are reported separately. */
export const STUCK = [STATUS.BLOCKED, STATUS.WAITING];

/* ------------------------------------------------------------------ */

export function xlsxAvailable() {
  return typeof window !== 'undefined' && !!window.XLSX;
}

export async function readWorkbook(file) {
  if (!xlsxAvailable()) {
    throw new Error('The spreadsheet reader did not load. Reload the page; if it persists, check that assets/vendor/xlsx.full.min.js was deployed.');
  }
  const buf = await file.arrayBuffer();
  const wb = window.XLSX.read(buf, { type: 'array', cellDates: true });
  return {
    name: file.name,
    size: file.size,
    sheets: wb.SheetNames.map((name) => ({
      name,
      rows: window.XLSX.utils.sheet_to_json(wb.Sheets[name], {
        header: 1, raw: true, defval: null, blankrows: true,
      }).map((r) => (Array.isArray(r) ? r : [])),
    })),
  };
}

/* ---------------------------- cell helpers ---------------------------- */

const txt = (v) => (v === null || v === undefined ? '' : String(v).trim());
const isDateCell = (v) => v instanceof Date && !isNaN(v);

/**
 * Locate a header row by the labels it must contain.
 *
 * The default limit is generous because secondary tables sit a long way down a
 * sheet: on the real template the prerequisite table starts around row 62,
 * below fifteen categories and forty task rows. A tighter limit silently
 * returned zero prerequisites and nothing complained, which is exactly the kind
 * of quiet failure that makes a tool untrustworthy.
 */
function findHeaderRow(rows, mustContain, limit = 500) {
  const want = mustContain.map(norm);
  for (let r = 0; r < Math.min(rows.length, limit); r++) {
    const cells = (rows[r] || []).map((c) => norm(c));
    if (want.every((w) => cells.includes(w))) return r;
  }
  return -1;
}

/** Map header labels to column indexes using an alias dictionary. */
function mapColumns(headerRow, dict) {
  const out = {};
  const cells = (headerRow || []).map((c) => norm(c));
  for (const [key, aliases] of Object.entries(dict)) {
    for (let i = 0; i < cells.length; i++) {
      if (!cells[i] || out[key] !== undefined) continue;
      if (aliases.includes(cells[i])) { out[key] = i; break; }
    }
    if (out[key] === undefined) {
      for (let i = 0; i < cells.length; i++) {
        if (!cells[i] || out[key] !== undefined) continue;
        if (aliases.some((a) => a.length >= 4 && cells[i].includes(a))) { out[key] = i; break; }
      }
    }
  }
  return out;
}

const REGISTER_COLS = {
  code:        ['sitecode', 'code', 'siteid', 'id', 'ref'],
  description: ['descriptiongeneric', 'description', 'generic', 'label'],
  package:     ['packagetype', 'package', 'scopetype', 'type'],
  wave:        ['wavebatch', 'wave', 'batch', 'group', 'phase'],
  start:       ['startdate', 'start', 'plannedstart', 'commencement'],
  target:      ['targetsubmission', 'target', 'targetdate', 'submissiondate', 'duedate', 'deadline'],
  actual:      ['actualsubmission', 'actualsubmissiondate', 'submitted', 'actual'],
  category:    ['currentcategory', 'currentstage', 'stage', 'category'],
  status:      ['status', 'state'],
  priority:    ['priority'],
  coordinator: ['coordinatorrcode', 'coordinator', 'lead', 'owner', 'responsible'],
  sheet:       ['detailsheet', 'sheet', 'tab', 'sheetname', 'worksheet'],
  notes:       ['notes', 'remarks', 'comment', 'comments'],
};

const TASK_COLS = {
  id:        ['taskid', 'id', 'ref'],
  category:  ['category', 'stage', 'group'],
  discipline:['discipline', 'trade'],
  name:      ['taskcomponent', 'task', 'taskname', 'component', 'description', 'scopeitem'],
  weight:    ['weight', 'size', 'effort'],
  type:      ['type', 'scopetype', 'regularadditional'],
  added:     ['dateadded', 'added', 'dateaddedtoscope'],
  target:    ['targetweek', 'target', 'plannedweek', 'dueweek'],
  doneWeek:  ['completedweek', 'completionweek', 'finishedweek'],
  doneDate:  ['completeddate', 'completiondate', 'finisheddate', 'actualfinish'],
  resource:  ['responsiblercode', 'responsible', 'resource', 'resourcecode', 'owner', 'assignedto'],
  prereq:    ['prerequisitepids', 'prerequisite', 'prerequisites', 'prereq', 'prereqids'],
  depends:   ['dependsontaskids', 'dependson', 'dependency', 'dependencies', 'predecessor', 'predecessors'],
};

const PREREQ_COLS = {
  id:        ['prereqid', 'prerequisiteid', 'id'],
  name:      ['prerequisite', 'description', 'item', 'name'],
  fromTpl:   ['fromtemplateptid', 'fromtemplate', 'templateid', 'ptid'],
  neededFor: ['neededfortaskids', 'neededfor', 'tasks', 'taskids'],
  provider:  ['providedbyrcodeparty', 'providedby', 'provider', 'source', 'responsible'],
  byWeek:    ['requiredbyweek', 'requiredby', 'neededbyweek', 'week'],
  status:    ['status', 'state'],
  received:  ['datereceived', 'received', 'receiveddate'],
  notes:     ['notes', 'remarks', 'comment'],
};

const LOG_COLS = {
  id:       ['logid', 'id', 'ref'],
  week:     ['weekwno', 'week', 'wno', 'weekno', 'weeknumber'],
  taskId:   ['taskid', 'task', 'id'],
  kind:     ['statusraised', 'status', 'type', 'kind'],
  reason:   ['reason', 'cause', 'why', 'description'],
  waitingOn:['waitingonrcodeparty', 'waitingon', 'party', 'with', 'owner'],
  raised:   ['raiseddate', 'raised', 'dateraised', 'from'],
  expected: ['expectedcleardate', 'expected', 'expectedclear', 'eta'],
  cleared:  ['cleareddate', 'cleared', 'resolved', 'resolveddate'],
  notes:    ['notes', 'remarks', 'comment'],
};

/* ------------------------------ register ------------------------------ */

function parseRegister(sheet) {
  const hr = findHeaderRow(sheet.rows, ['site code']) >= 0
    ? findHeaderRow(sheet.rows, ['site code'])
    : findHeaderRow(sheet.rows, ['sitecode']);
  if (hr < 0) return null;

  const cols = mapColumns(sheet.rows[hr], REGISTER_COLS);
  if (cols.code === undefined) return null;

  const sites = [];
  for (let r = hr + 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r] || [];
    const code = txt(row[cols.code]);
    if (!code) continue;
    if (/^table\b/i.test(code)) break;
    sites.push({
      id: uid('site'),
      code,
      description: txt(row[cols.description]),
      package: txt(row[cols.package]),
      wave: txt(row[cols.wave]),
      start: toISO(row[cols.start]),
      target: toISO(row[cols.target]),
      actual: toISO(row[cols.actual]),
      currentCategory: txt(row[cols.category]),
      status: txt(row[cols.status]),
      priority: txt(row[cols.priority]),
      coordinator: txt(row[cols.coordinator]),
      sheet: txt(row[cols.sheet]),
      notes: txt(row[cols.notes]),
      sourceRow: r + 1,
    });
  }
  return { sheetName: sheet.name, headerRow: hr, columns: cols, sites };
}

/* ------------------------------ template ------------------------------ */

function parseTemplate(sheet) {
  const rows = sheet.rows;
  const tplHr = findHeaderRow(rows, ['order', 'category']);
  const categories = [];
  const tasks = [];

  if (tplHr >= 0) {
    const cols = mapColumns(rows[tplHr], {
      order: ['order', 'seq', 'sequence', 'no'],
      id: ['id', 'code', 'ref'],
      category: ['category', 'stage'],
      task: ['standardtaskscopeitem', 'standardtask', 'task', 'scopeitem'],
      dod: ['definitionofdonecategory', 'definitionofdone', 'dod', 'doneWhen'],
      duration: ['typicaldurationdays', 'typicalduration', 'duration', 'days'],
      applies: ['appliesto', 'applies', 'discipline'],
      prereq: ['defaultprerequisitesptids', 'defaultprerequisites', 'prerequisites', 'prereq'],
    });
    for (let r = tplHr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const id = txt(row[cols.id]);
      if (/^prereq/i.test(txt(row[0])) || /^table\b/i.test(txt(row[0]))) break;
      if (!id) continue;
      if (/^C\d+/i.test(id)) {
        categories.push({
          id: id.toUpperCase(),
          order: toNum(row[cols.order]) ?? categories.length * 10 + 10,
          name: txt(row[cols.category]),
          definitionOfDone: txt(row[cols.dod]),
          duration: toNum(row[cols.duration]),
          appliesTo: txt(row[cols.applies]),
          defaultPrereqs: splitIds(txt(row[cols.prereq])),
        });
      } else if (/^T[\d]/i.test(id)) {
        tasks.push({
          id: id.toUpperCase(),
          categoryId: `C${id.replace(/^T/i, '').split('-')[0]}`,
          category: txt(row[cols.category]),
          name: txt(row[cols.task]),
          discipline: txt(row[cols.applies]),
        });
      }
    }
  }

  // Prerequisite template table
  const ptHr = findHeaderRow(rows, ['prereq id']) >= 0
    ? findHeaderRow(rows, ['prereq id'])
    : findHeaderRow(rows, ['prereqid']);
  const prereqs = [];
  if (ptHr >= 0) {
    const cols = mapColumns(rows[ptHr], {
      id: ['prereqid', 'id'],
      name: ['prerequisite', 'name', 'description'],
      why: ['whyitisneeded', 'why', 'reason'],
      provider: ['usuallyprovidedby', 'providedby', 'provider'],
      before: ['neededbeforecategoryid', 'neededbefore', 'before', 'category'],
    });
    for (let r = ptHr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const id = txt(row[cols.id]);
      if (!id) continue;
      if (!/^PT\d+/i.test(id)) continue;
      prereqs.push({
        id: id.toUpperCase(),
        name: txt(row[cols.name]),
        why: txt(row[cols.why]),
        provider: txt(row[cols.provider]),
        before: txt(row[cols.before]).toUpperCase(),
      });
    }
  }

  return { sheetName: sheet.name, categories, tasks, prereqs };
}

export function splitIds(s) {
  return String(s || '')
    .split(/[,;/|\n]+/)
    .map((x) => x.trim().toUpperCase())
    .filter(Boolean);
}

/* ----------------------------- site sheet ----------------------------- */

function parseSiteSheet(sheet) {
  const rows = sheet.rows;
  const problems = [];

  const hr = findHeaderRow(rows, ['task id']) >= 0
    ? findHeaderRow(rows, ['task id'])
    : findHeaderRow(rows, ['taskid']);
  if (hr < 0) return null;

  const header = rows[hr] || [];
  const cols = mapColumns(header, TASK_COLS);

  // Week columns are identified by a date in the header — unambiguous, and it
  // means the user can add weeks to the right without telling the app.
  const weeks = [];
  for (let c = 0; c < header.length; c++) {
    if (!isDateCell(header[c])) continue;
    const labelRow = rows[hr - 1] || [];
    const label = txt(labelRow[c]);
    weeks.push({
      col: c,
      end: toISO(header[c]),
      label: label || `W${String(weeks.length + 1).padStart(2, '0')}`,
      number: weeks.length + 1,
    });
  }
  if (!weeks.length) {
    problems.push('No week columns found. Their headers must be dates (the date the week ends).');
  }

  // Meta block above the table: label row / value row pairs.
  const meta = {};
  for (let r = 0; r < hr - 1; r++) {
    const labels = rows[r] || [];
    const values = rows[r + 1] || [];
    for (let c = 0; c < labels.length; c++) {
      const k = txt(labels[c]);
      const v = values[c];
      if (!k || v === null || v === undefined || txt(v) === '') continue;
      if (k.length > 40) continue;
      if (!(norm(k) in meta)) meta[norm(k)] = isDateCell(v) ? toISO(v) : txt(v);
    }
  }

  const tasks = [];
  const categories = [];
  let lastRow = hr;

  for (let r = hr + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const id = txt(row[cols.id]);
    if (/^table\b/i.test(id)) break;
    if (!id) continue;

    const weekly = weeks.map((w) => normStatus(row[w.col]));
    const rawWeekly = weeks.map((w) => txt(row[w.col]));
    for (let i = 0; i < weeks.length; i++) {
      if (rawWeekly[i] && !weekly[i]) {
        problems.push(`Row ${r + 1}, ${weeks[i].label}: "${rawWeekly[i]}" is not a recognised status.`);
      }
    }

    if (/^C\d+/i.test(id)) {
      categories.push({
        id: id.toUpperCase(),
        name: txt(row[cols.category]),
        weekly,
        row: r + 1,
      });
      lastRow = r;
      continue;
    }
    if (!/^T/i.test(id)) continue;

    const name = txt(row[cols.name]);
    if (!name) { problems.push(`Row ${r + 1}: task ${id} has no name and was skipped.`); continue; }

    tasks.push({
      id: id.toUpperCase(),
      categoryId: `C${id.replace(/^T/i, '').split('-')[0]}`,
      category: txt(row[cols.category]),
      discipline: txt(row[cols.discipline]) || 'All',
      name,
      weight: toNum(row[cols.weight]) ?? 1,
      type: /add/i.test(txt(row[cols.type])) ? 'Additional' : 'Regular',
      added: toISO(row[cols.added]),
      targetWeek: toNum(row[cols.target]),
      weekly,
      doneWeek: toNum(row[cols.doneWeek]),
      doneDate: toISO(row[cols.doneDate]),
      resource: txt(row[cols.resource]),
      prereqs: splitIds(txt(row[cols.prereq])),
      depends: splitIds(txt(row[cols.depends])),
      row: r + 1,
    });
    lastRow = r;
  }

  // Prerequisites table
  const prereqs = [];
  const pHr = findHeaderRow(rows.slice(lastRow), ['prereq id']) >= 0
    ? lastRow + findHeaderRow(rows.slice(lastRow), ['prereq id'])
    : (findHeaderRow(rows.slice(lastRow), ['prereqid']) >= 0
        ? lastRow + findHeaderRow(rows.slice(lastRow), ['prereqid']) : -1);
  if (pHr >= 0) {
    const pc = mapColumns(rows[pHr], PREREQ_COLS);
    for (let r = pHr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const id = txt(row[pc.id]);
      if (/^table\b/i.test(id) || /^log/i.test(id)) break;
      if (!/^P\d+/i.test(id)) continue;
      prereqs.push({
        id: id.toUpperCase(),
        name: txt(row[pc.name]),
        fromTemplate: txt(row[pc.fromTpl]).toUpperCase(),
        neededFor: splitIds(txt(row[pc.neededFor])),
        provider: txt(row[pc.provider]),
        byWeek: toNum(row[pc.byWeek]),
        status: txt(row[pc.status]),
        received: toISO(row[pc.received]),
        notes: txt(row[pc.notes]),
        row: r + 1,
      });
    }
  }

  // Waiting-on / blocked log
  const log = [];
  const lHr = findHeaderRow(rows, ['log id']) >= 0
    ? findHeaderRow(rows, ['log id'])
    : findHeaderRow(rows, ['logid']);
  if (lHr >= 0) {
    const lc = mapColumns(rows[lHr], LOG_COLS);
    for (let r = lHr + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const id = txt(row[lc.id]);
      if (!/^L\d+/i.test(id)) continue;
      log.push({
        id: id.toUpperCase(),
        week: toNum(row[lc.week]),
        taskId: txt(row[lc.taskId]).toUpperCase(),
        kind: normStatus(row[lc.kind]) || txt(row[lc.kind]),
        reason: txt(row[lc.reason]),
        waitingOn: txt(row[lc.waitingOn]),
        raised: toISO(row[lc.raised]),
        expected: toISO(row[lc.expected]),
        cleared: toISO(row[lc.cleared]),
        notes: txt(row[lc.notes]),
        row: r + 1,
      });
    }
  }

  return {
    sheetName: sheet.name,
    headerRow: hr,
    meta,
    weeks,
    categories,
    tasks,
    prereqs,
    log,
    problems,
  };
}

/* ------------------------------ top level ----------------------------- */

const NON_DATA = [/^read ?me/i, /^how to use/i, /^instructions?$/i, /^guide$/i, /^legend$/i, /^notes?$/i];

/**
 * Parse the whole workbook into the model the engine consumes.
 * Returns { register, template, sites[], notes[], problems[] }.
 */
export function parseWorkbook(wb) {
  const notes = [];
  const problems = [];

  let register = null;
  let template = null;
  const consumed = new Set();

  for (const sheet of wb.sheets) {
    if (NON_DATA.some((re) => re.test(sheet.name))) { consumed.add(sheet.name); continue; }
    if (!register) {
      const r = parseRegister(sheet);
      if (r && r.sites.length) { register = r; consumed.add(sheet.name); continue; }
    }
    if (!template) {
      const t = parseTemplate(sheet);
      if (t && (t.categories.length || t.prereqs.length) && !t.tasks.some((x) => x.weekly)) {
        // A template sheet has categories but no week columns.
        const hasWeeks = (sheet.rows[findHeaderRow(sheet.rows, ['task id'])] || [])
          .some((c) => isDateCell(c));
        if (!hasWeeks && (t.categories.length >= 2 || t.prereqs.length >= 2)) {
          template = t; consumed.add(sheet.name); continue;
        }
      }
    }
  }

  if (!register) {
    problems.push('No site register found. The first sheet should have a "Site code" column listing your sites.');
  }
  if (!template) {
    notes.push('No task template sheet was recognised. Categories will be taken from the site sheets themselves.');
    template = { sheetName: null, categories: [], tasks: [], prereqs: [] };
  }

  // Attach each register row to its detail sheet.
  const sites = [];
  const byName = new Map(wb.sheets.map((s) => [norm(s.name), s]));
  const claimed = new Set(consumed);

  for (const entry of (register?.sites || [])) {
    let sheet = null;
    let matchedBy = null;
    if (entry.sheet && byName.has(norm(entry.sheet)) && !claimed.has(byName.get(norm(entry.sheet)).name)) {
      sheet = byName.get(norm(entry.sheet)); matchedBy = 'named in register';
    } else if (byName.has(norm(entry.code)) && !claimed.has(byName.get(norm(entry.code)).name)) {
      sheet = byName.get(norm(entry.code)); matchedBy = 'tab name matches site code';
    } else {
      const hit = wb.sheets.find((s) => !claimed.has(s.name) && norm(s.name).includes(norm(entry.code)) && norm(entry.code).length >= 2);
      if (hit) { sheet = hit; matchedBy = 'tab name contains site code'; }
    }

    let parsed = null;
    if (sheet) {
      claimed.add(sheet.name);
      parsed = parseSiteSheet(sheet);
      if (!parsed) {
        problems.push(`Sheet "${sheet.name}" was matched to ${entry.code} but has no "Task ID" header row, so no tasks were read from it.`);
      } else if (parsed.problems.length) {
        for (const p of parsed.problems) problems.push(`${entry.code}: ${p}`);
      }
    } else {
      notes.push(`${entry.code} has no detail sheet. It will be tracked in the register only, with no tasks.`);
    }

    sites.push({
      ...entry,
      matchedBy,
      detail: parsed,
      taskCount: parsed?.tasks.length || 0,
      weekCount: parsed?.weeks.length || 0,
    });
  }

  const orphans = wb.sheets.filter((s) => !claimed.has(s.name) && !NON_DATA.some((re) => re.test(s.name)));
  for (const o of orphans) {
    const p = parseSiteSheet(o);
    if (p && p.tasks.length) {
      problems.push(`Sheet "${o.name}" holds ${p.tasks.length} tasks but no register row points at it, so it will NOT be analysed. Add a row to the Sites sheet.`);
    }
  }

  return {
    file: wb.name,
    parsedAt: new Date().toISOString(),
    register,
    template,
    sites,
    orphanSheets: orphans.map((o) => o.name),
    notes,
    problems,
  };
}

/* --------------------------- privacy screening --------------------------- */

/**
 * Scans free text heading for the AI for anything that looks like a real name
 * or place. This is a safety net, not a guarantee — the user remains the check,
 * and the panel says so.
 */
const SAFE_WORDS = new Set([
  'level', 'ground', 'first', 'second', 'third', 'north', 'south', 'east', 'west',
  'draft', 'model', 'sheet', 'design', 'revision', 'check', 'qa', 'submission',
  'arch', 'struct', 'mep', 'civil', 'all', 'zone', 'block', 'type', 'phase',
  'foundations', 'columns', 'beams', 'slabs', 'walls', 'ductwork', 'pipework',
  'containment', 'envelope', 'substructure', 'superstructure', 'clash', 'audit',
  'schedule', 'schedules', 'quantity', 'extract', 'report', 'transmittal', 'cde',
  'revit', 'families', 'template', 'coordinates', 'survey', 'point', 'cloud',
  'internal', 'external', 'primary', 'routes', 'secondary', 'elements', 'annotation',
  'dimensions', 'titleblock', 'federate', 'detection', 'standards', 'update',
  'updates', 'produce', 'issue', 'receive', 'confirm', 'assess', 'implement',
  'create', 'apply', 'set', 'add', 'log', 'return', 'export', 'package', 'develop',
  'review', 'run', 'resolve', 'purge', 'cleanup', 'setup', 'views', 'general',
  'arrangement', 'option', 'impact', 'information', 'team', 'work', 'scope',
  'prerequisite', 'preparation', 'deliverables', 'discipline', 'comments',
  'tuning', 'individual', 'initial', 'modelling', 'modeling', 'kickoff', 'folder',
  'structure', 'file', 'files', 'project', 'site', 'sites', 'week', 'day', 'days',
]);

export function screenText(strings) {
  const hits = [];
  const seen = new Set();
  for (const { where, text } of strings) {
    if (!text) continue;
    // Two or more consecutive capitalised words that are not known BIM terms.
    const re = /\b([A-Z][a-z]{2,})(\s+(?:of|the|at|in|de|al|el)\s+|\s+)([A-Z][a-z]{2,})\b/g;
    let m;
    while ((m = re.exec(text))) {
      const phrase = m[0];
      const words = phrase.toLowerCase().split(/\s+/).filter((w) => !['of', 'the', 'at', 'in', 'de', 'al', 'el'].includes(w));
      if (words.every((w) => SAFE_WORDS.has(w))) continue;
      const key = `${where}|${phrase}`;
      if (seen.has(key)) continue;
      seen.add(key);
      hits.push({ where, phrase, text });
    }
  }
  return hits;
}
