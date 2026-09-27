/**
 * AI layer.
 *
 * The model receives COMPUTED RESULTS, never the spreadsheet. Percentages,
 * velocity, forecasts, blocker ages and risks are all calculated in
 * js/core/engine.js and passed in as fixed facts. The model's job is
 * interpretation: what it means, what to do, what to watch.
 *
 * Asking a model to compute a forecast from a status grid produces numbers that
 * look right and are not, and a coordinator would act on them. That is the one
 * mistake this tool cannot afford, so the boundary is enforced here and stated
 * in the system prompt.
 */

import { getApiKey } from './session.js';
import {
  recordRequest, recordUsage, readProviderUsage, estimateTokens,
  readRateLimitHeaders,
} from './tokens.js';

export const PROVIDERS = {
  gemini: {
    label: 'Google Gemini',
    keyHint: 'Paste your Google AI Studio key',
    keyUrl: 'https://aistudio.google.com/apikey',
    defaultModel: 'gemini-2.5-flash',
    suggestions: ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash'],
    note: 'Calls go directly from your browser to Google.',
  },
  anthropic: {
    label: 'Anthropic Claude',
    keyHint: 'Paste your Anthropic console key',
    keyUrl: 'https://console.anthropic.com/settings/keys',
    defaultModel: 'claude-sonnet-4-5',
    suggestions: ['claude-sonnet-4-5', 'claude-haiku-4-5', 'claude-opus-4-1'],
    note: 'Browser calls need direct-access enabled on the request, which this app sets.',
  },
  openai: {
    label: 'OpenAI',
    keyHint: 'Paste your OpenAI platform key',
    keyUrl: 'https://platform.openai.com/api-keys',
    defaultModel: 'gpt-4.1-mini',
    suggestions: ['gpt-4.1-mini', 'gpt-4.1', 'gpt-4o'],
    note: 'Calls go directly from your browser to OpenAI.',
  },
};

/* ============================ system prompt ============================ */

const SYSTEM = `You are a delivery manager reviewing a BIM production programme running across several similar sites. Your reader is the coordinator who runs it: technically fluent, short of time, accountable for the submission dates.

GROUND RULES — these are not stylistic preferences.

1. Every number in the input was computed by the application from the source spreadsheet. Treat them as given facts. Do NOT recalculate percentages, velocity, forecasts or dates, and do NOT invent any figure that is not in the input. If you want to express a rate or a total that is not provided, say what you would need instead.
2. Sites are identified by code only (A-01, people by R1, R2). You will never be told real names, and you must not speculate about them or ask for them.
3. Be specific and name things. "Monitor the MEP package" is worthless. "T030-04 has been Waiting on for 2 weeks against R5 for ceiling void confirmation, and it blocks the clash run" is useful.
4. Separate what has happened from what is forecast, and say what a forecast rests on. A forecast built on three weeks of data is not the same as one built on twelve.
5. Where the data cannot support a conclusion, say what is missing rather than filling the gap.
6. Plain professional English. No filler, no restating the input back, no motivational language, no headings that just repeat the schema keys.
7. The report is divided into fixed sections. SAY EACH THING ONCE, in the section where it belongs. If a blocker belongs in the blocked-log section, do not repeat it in the task-status section or the conclusion. The conclusion draws the threads together in new words; it is not a summary that repeats earlier sentences.
8. The application renders the figures and the charts itself. Do not list numbers back that are already in the input — interpret them. "43 tasks with 9 done" is wasted space; "the four weeks since kickoff have produced setup work only, with no modelling closed out" is not.

Return strictly valid JSON matching the requested shape. No markdown fences, no commentary outside the JSON.`;

/* ============================== payloads ==============================
 *
 * The model is handed the COMPUTED REPORT, not the workbook. Every figure has
 * already been calculated by js/core/report.js, so the model's only job is the
 * judgement the template marks as AI: the executive summary, one finding per
 * engine section, the risk register, the actions and the conclusion.
 *
 * Nothing here contains the raw weekly grid, and nothing asks for a number to
 * be derived.
 * ===================================================================== */

const MAX_ROWS = 40;
const trim = (a, n = MAX_ROWS) => (Array.isArray(a) ? a.slice(0, n) : []);

export function buildSitePayload(rep) {
  if (!rep || rep.kind !== 'site') {
    return {
      reportType: 'site status report', site: rep?.code || '(unknown)', hasData: false,
      note: 'No computed report is available for this site. Say plainly that it cannot be assessed and why.',
    };
  }
  const m = rep.metrics;
  return {
    reportType: 'site status report',
    site: rep.code,
    hasData: true,
    computedNote: 'Every figure below was computed by the application from the workbook. Treat them as fact. Do not recalculate them, and do not state any figure that is not here.',
    scopeRule: 'This report covers this site only. Never refer to, compare with or rank against any other site, even if others were analysed in the same run.',

    documentControl: rep.docControl,
    overallRag: rep.overall,
    submitted: rep.submitted,

    metrics: {
      statusDate: m.statusDate,
      applicableTasks: m.applicableTasks, totalTasks: m.totalTasks, naTasks: m.naTasks,
      finished: m.finished, wip: m.wip, notStarted: m.notStarted, stuck: m.stuck,
      completionCountPct: m.completionCount, completionWeightedPct: m.completionWeighted,
      timeElapsedPct: m.timeElapsed, timeBasedSpiProxy: m.spiProxy,
      weeklyThroughput: m.throughput,
      scheduleVarianceDays: m.scheduleVariance,
      forecastSubmission: m.forecastDate, forecastMethod: m.forecastMethod,
      indicativeWorkingDaysRemaining: m.indicativeDays,
    },
    kpis: rep.kpis.map((k) => ({ indicator: k.k, actual: k.v, target: k.target, rag: k.rag })),
    ragByDimension: Object.fromEntries(
      Object.entries(rep.dimensions).filter(([k]) => k !== 'overall')
        .map(([k, v]) => [k, { rag: v.rag, basis: v.basis }])),

    milestones: rep.milestones,
    weekOnWeekMovement: rep.movement.comparable
      ? {
        from: rep.movement.fromWeek, to: rep.movement.toWeek,
        finishedThisWeek: rep.movement.finishedThisWeek, changes: trim(rep.movement.changes, 25),
      }
      : { comparable: false, reason: rep.movement.reason },

    progressByCategory: rep.wbs.rows.map((c) => ({
      id: c.id, name: c.name, applicable: c.liveCount, finished: c.doneCount,
      percent: c.computedPct, reported: c.selfStatus, rollUpCheck: c.check, stuck: c.stuck,
    })),
    rollUpDiscrepancies: rep.wbs.discrepancies,

    scope: {
      source: rep.scope.source,
      items: trim(rep.scope.items, 20),
      openItems: rep.scope.openItems.length,
      addedAfterKickoff: rep.scope.additionalCount,
      addedThenSetNA: rep.scope.additionalSetNA,
      growthPct: rep.scope.growthPct,
    },
    prerequisites: trim(rep.constraints.rows, 25).map((p) => ({
      id: p.id, name: p.name, neededFor: p.neededFor, provider: p.provider,
      requiredByWeek: p.byWeek, status: p.status, rag: p.rag,
      notTracked: p.notTracked || undefined,
    })),
    waitingOnBlockedLog: trim(rep.log.rows, 25).map((l) => ({
      id: l.id, week: l.week, task: l.taskId, type: l.kind, reason: l.reason,
      waitingOn: l.waitingOn, raised: l.raised, ageDays: l.age, control: l.control,
      rag: l.rag, cleared: l.cleared || undefined,
    })),
    resources: {
      assigned: rep.resources.rows.map((r) => ({ resource: r.resource, open: r.open, wip: r.wip, stuck: r.stuck })),
      highestWip: rep.resources.maxWip,
      soleProviders: rep.resources.soleProviders,
      causesBlocked: rep.resources.causesBlocked,
      externalParties: rep.resources.externalParties,
      unassignedOpenTasks: rep.resources.unassigned?.open || 0,
    },
    dataIntegrityFindings: rep.quality.findings.map((f) => ({
      id: f.id, finding: f.finding, evidence: f.evidence, impact: f.impact,
      correction: f.correction, severity: f.severity,
    })),
  };
}

export function buildMasterPayload(mr) {
  return {
    reportType: 'master analysis (portfolio) report',
    computedNote: 'Every figure below was computed by the application from the individual site reports. Treat them as fact and do not recalculate them.',
    scopeRule: 'Only the sites listed here were selected. Never mention, count or compare any site that is not in this list.',
    documentControl: mr.docControl,
    overallRag: mr.overall,
    counts: mr.counts,
    portfolioKpis: mr.kpis,
    ragBySiteAndDimension: mr.ragMatrix,
    scheduleBySite: mr.schedule,
    progressByCategory: mr.categories,
    scopeBySite: mr.scope,
    openConstraints: trim(mr.constraints, 30),
    externalParties: mr.parties,
    resourceLoad: mr.resourceLoad,
    resourceLoadCaveat: 'Loading reflects the selected sites only; a resource may also be committed to sites that were not selected.',
    qualityPatterns: mr.patterns,
    systemicPatterns: mr.systemicPatterns,
    interventionPriorityIndex: mr.ipi,
    siteSummaries: mr.appendixB,
    noActiveSites: mr.noActiveSites,
  };
}

/* ============================== schemas ============================== */

const SITE_SCHEMA = `{
  "executive": {
    "bottomLine": "one sentence: the single most important thing about this site, stated as an answer not a summary. Lead with the conclusion.",
    "keyMessages": ["3-5 bullets, each a claim followed by the evidence from the figures given. Never a figure that is not in the input."],
    "decisions": [{"id": "D1", "decision": "a decision only a manager can take, phrased as an ask", "owner": "role or R code from the input", "neededBy": "a date or week from the input"}]
  },

  "notes": {
    "schedule": "2-4 sentences on the schedule position and what the forecast rests on. Section 4.",
    "wbs": "2-3 sentences on where the work sits across categories and what the roll-up checks mean. Section 5.",
    "scope": "2-3 sentences on scope and change control, including whether growth is being measured honestly. Section 6.",
    "constraints": "2-3 sentences on prerequisites: what is outstanding and what it gates. Section 7.",
    "log": "2-3 sentences on blocked and waiting items, distinguishing Internal (the team can fix) from External (must be chased). Section 8.",
    "resources": "2-3 sentences on loading, single-provider exposure and who is holding work up. Section 10.",
    "quality": "2-3 sentences on what the data-integrity findings mean for trusting this report. Section 11."
  },

  "risks": [{
    "id": "RK1",
    "risk": "stated as cause then effect, e.g. 'X is not confirmed, so Y will slip'",
    "probability": 1, "impact": 1,
    "strategy": "Avoid" | "Reduce" | "Transfer" | "Accept",
    "response": "the specific action that changes the probability or the impact",
    "owner": "role or R code from the input"
  }],

  "actions": [{"id": "A1", "action": "...", "owner": "role or R code", "due": "date or week", "priority": "High"|"Medium"|"Low", "links": "risk, milestone or finding IDs this addresses"}],

  "lookahead": [{"week": "the next status week label", "focus": "what must happen that week"}],

  "conclusion": "3-5 sentences. State the position, the single most consequential choice, and what changes if it is taken. New words: do not repeat sentences from the sections above."
}`;

const MASTER_SCHEMA = `{
  "executive": {
    "bottomLine": "one sentence covering the selected sites as a group.",
    "keyMessages": ["4-6 bullets: portfolio position, schedule, the dominant bottleneck, shared-resource exposure, governance patterns. Each with its evidence."],
    "decisions": [{"id": "MD1", "decision": "a portfolio-level decision", "owner": "role", "neededBy": "date"}]
  },

  "notes": {
    "schedule": "2-4 sentences comparing the sites against their own dates. Section 4.",
    "bottlenecks": "2-4 sentences on where the remaining work is concentrated and why that matters. Section 5.",
    "scope": "2-3 sentences on whether scope and change are captured consistently across sites. Section 6.",
    "constraints": "2-3 sentences on external parties holding work across sites. Section 7.",
    "resources": "2-4 sentences on shared resources and capacity. Section 8.",
    "quality": "2-3 sentences on which findings are systemic rather than local. Section 10.",
    "priority": "2-3 sentences justifying the ranking and naming the tie-break. Section 11."
  },

  "risks": [{
    "id": "PR1",
    "risk": "a risk that spans sites or arises from the combination; site-specific risks belong in the site reports",
    "sitesAffected": ["A-01"],
    "probability": 1, "impact": 1,
    "strategy": "Avoid" | "Reduce" | "Transfer" | "Accept",
    "response": "...", "owner": "role"
  }],

  "interventions": [{"site": "A-01", "recommendedIntervention": "what to do on this site, one line"}],

  "actions": [{"id": "PA1", "action": "...", "owner": "role", "due": "date", "priority": "High"|"Medium"|"Low", "links": "risk or decision IDs"}],

  "lookahead": [{"week": "week label", "focus": "portfolio focus for that week"}],

  "conclusion": "4-6 sentences: the portfolio position, the pattern behind it, and the few decisions that change the outcome most."
}`;

/* ============================== prompts ============================== */

export function buildPrompt(kind, payload, followUp = '') {
  const parts = [];
  if (kind === 'master') {
    parts.push('You are writing the judgement sections of a Master Analysis (portfolio) report over the selected sites. The application has already computed every figure; your job is interpretation, not calculation.');
    parts.push('Find what is systemic rather than local, where sites compete for the same people, and which decisions change the outcome most.');
    parts.push('\nReturn JSON matching exactly this shape:\n' + MASTER_SCHEMA);
  } else {
    parts.push(`You are writing the judgement sections of a Site Status Report for site ${payload.site}. The application has already computed every figure; your job is interpretation, not calculation.`);
    parts.push('Write for a Project Manager who needs to decide something this week. Lead with the answer, then the evidence.');
    parts.push('\nReturn JSON matching exactly this shape:\n' + SITE_SCHEMA);
  }
  parts.push('\nRules for the judgement you are asked for:');
  parts.push('- Score each risk with a probability and an impact from 1 to 5. The application multiplies them and applies the rating band; do not state a score or a rating yourself.');
  parts.push('- Every claim must trace to a figure in the input. If the data cannot support a point, say what is missing instead of filling the gap.');
  parts.push('- Name things: task IDs, prerequisite IDs, R codes, dates. "Monitor progress" is not an action.');
  parts.push('- Say each thing once, in the section where it belongs. The conclusion draws threads together in new words rather than repeating earlier sentences.');
  if (payload.previousPeriod) {
    parts.push('\nA previous review of this same scope is included in the data as "previousPeriod". Use it: say what has moved, what has not, and whether actions raised last time were acted on. Do not simply repeat it.');
  }
  if (followUp) {
    parts.push('\nThe reader has asked specifically:\n' + followUp + '\nAnswer that within the same JSON shape, in the summary and actions.');
  }
  parts.push('\nComputed programme data:\n```json\n' + JSON.stringify(payload, null, 1) + '\n```');
  return parts.join('\n');
}

export function previewPayload({ kind, payload, followUp, settings }) {
  const prompt = buildPrompt(kind, payload, followUp);
  const provider = PROVIDERS[settings.provider];
  return {
    destination: provider?.label || settings.provider,
    model: (settings.model || '').trim() || provider?.defaultModel,
    system: SYSTEM,
    prompt,
    payload,
    estimatedInputTokens: estimateTokens(SYSTEM) + estimateTokens(prompt),
    bytes: new Blob([prompt]).size,
  };
}

/* ============================== transport ============================== */

/** Merges a caller cancel signal with a hard timeout. */
function makeSignal(external, ms) {
  const ctrl = new AbortController();
  // If the caller already cancelled, abort before fetch attaches its listener
  // and fetch never rejects — so bail out explicitly instead.
  if (external?.aborted) { ctrl.abort(); return { signal: ctrl.signal, done: () => {}, preAborted: true }; }
  const timer = setTimeout(() => ctrl.abort(new Error('timeout')), ms);
  if (external) external.addEventListener('abort', () => ctrl.abort(new Error('cancelled')), { once: true });
  return { signal: ctrl.signal, done: () => clearTimeout(timer), preAborted: false };
}

export class CancelledError extends Error {
  constructor() { super('Cancelled.'); this.name = 'CancelledError'; }
}

function retryAfterSeconds(response, bodyText) {
  const h = response?.headers?.get?.('retry-after');
  if (h && Number.isFinite(Number(h))) return Number(h);
  const m = String(bodyText || '').match(/retry in ([\d.]+)\s*s/i);
  return m ? Math.ceil(Number(m[1])) : null;
}

const sleep = (ms, signal) => new Promise((res, rej) => {
  const t = setTimeout(res, ms);
  signal?.addEventListener('abort', () => { clearTimeout(t); rej(new CancelledError()); }, { once: true });
});

/**
 * One request, with retry.
 *
 * WHY 503s HAPPEN: a 503 (or Anthropic's 529) means the provider's servers for
 * that model are overloaded. It is not your quota and not your token count —
 * the request never got far enough to be counted against either. The newest
 * models are the most crowded, and free-tier traffic is the first to be turned
 * away when they are busy.
 *
 * The old version retried three times over about 3.6 seconds, which is far too
 * short: overload clears in tens of seconds, not three. This waits roughly 5s,
 * 12s, 25s, 45s, with random jitter so many clients don't all retry in lockstep,
 * and honours Retry-After when the server sends one. Every retry is a real
 * request against your per-minute limit, which is why this caps at five.
 *
 * 429 is different: it is YOUR rate limit, and retrying immediately just burns
 * more of it. That fails fast with an explanation.
 */
let OVERLOAD_WAITS = [5000, 12000, 25000, 45000];

/** Test hook only: the real waits total over a minute, far too slow for a suite. */
export function _setOverloadWaits(arr) { OVERLOAD_WAITS = arr; }

/** Pull the provider's own error text out of a response body, if it sent one. */
function providerMessage(body) {
  try {
    const j = JSON.parse(body);
    return j?.error?.message || j?.message || '';
  } catch { return String(body || '').slice(0, 200); }
}

async function callWithRetry(provider, doFetch, { signal, onStatus, maxAttempts: cap } = {}) {
  const maxAttempts = Math.min(cap || Infinity, OVERLOAD_WAITS.length + 1);
  let lastStatus = null;
  let lastMessage = '';

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    if (signal?.aborted) throw new CancelledError();
    recordRequest(provider, attempt === 0 ? 'first attempt' : `retry after ${lastStatus || 'error'}`);

    let response;
    try {
      response = await doFetch();
    } catch (e) {
      if (signal?.aborted) throw new CancelledError();
      if (e?.name === 'AbortError') throw new Error('The request timed out. Try fewer sites at once.');
      lastStatus = 'network error';
      if (attempt === maxAttempts - 1) throw new Error(`Could not reach ${provider}. Check your connection. (${e.message})`);
      await waitWithStatus(OVERLOAD_WAITS[attempt], attempt, maxAttempts, 'connection failed', onStatus, signal);
      continue;
    }

    if (response.ok) { onStatus?.(null); return response; }

    const body = await response.text().catch(() => '');
    if (response.status === 429) {
      const wait = retryAfterSeconds(response, body);
      const e = new Error(
        `Rate limit reached (429) — this is your account's limit, not the server being busy.${wait ? ` The provider asked to wait about ${wait}s.` : ''} ` +
        'Retrying straight away would only burn more of your request quota. Wait, then run fewer sites at a time.');
      e.status = 429;
      throw e;
    }
    if (response.status === 401 || response.status === 403) {
      throw new Error(`${provider} rejected the API key (${response.status}). Check it is correct, active, and allowed for this model.`);
    }
    if (response.status === 404) {
      const e = new Error(`${provider} does not recognise the model name (404). Model names change — check the current list in Settings.`);
      e.status = 404;
      throw e;
    }
    if ([500, 502, 503, 504, 529].includes(response.status)) {
      lastStatus = response.status;
      lastMessage = providerMessage(body);
      if (attempt === maxAttempts - 1) {
        const waited = Math.round(OVERLOAD_WAITS.slice(0, maxAttempts - 1).reduce((a, b) => a + b, 0) / 1000);
        const e = new Error(
          `${provider} is still overloaded (${response.status}) after ${maxAttempts} attempt${maxAttempts === 1 ? '' : 's'}${waited ? ` over about ${waited}s` : ''}.` +
          `${lastMessage ? ` The provider said: "${lastMessage}"` : ''} ` +
          'This is the provider\'s servers being busy for this model, not your quota or tokens. Try again later, or switch to a more established model in Settings.');
        e.status = response.status;
        e.overloaded = true;
        e.providerMessage = lastMessage;
        throw e;
      }
      const hinted = retryAfterSeconds(response, body);
      const base = hinted ? hinted * 1000 : OVERLOAD_WAITS[attempt];
      await waitWithStatus(base, attempt, maxAttempts, `server busy (${response.status})`, onStatus, signal, lastMessage);
      continue;
    }
    let msg = body.slice(0, 300);
    try { msg = JSON.parse(body)?.error?.message || msg; } catch { /* keep raw */ }
    throw new Error(`${provider} returned ${response.status}: ${msg}`);
  }
  throw new Error('Request failed.');
}

/** Waits with jitter, reporting a live countdown so the user can see it is working. */
async function waitWithStatus(baseMs, attempt, maxAttempts, why, onStatus, signal, message = '') {
  const ms = Math.round(baseMs * (0.8 + Math.random() * 0.4));
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (signal?.aborted) { onStatus?.(null); throw new CancelledError(); }
    onStatus?.({
      why,
      attempt: attempt + 2,
      maxAttempts,
      secondsLeft: Math.ceil((until - Date.now()) / 1000),
      providerMessage: message,
    });
    await sleep(Math.min(1000, until - Date.now()), signal);
  }
}

async function callGemini({ model, prompt, signal, maxTokens, onStatus, maxAttempts }) {
  const key = getApiKey();
  const { signal: s, done } = makeSignal(signal, 120000);
  try {
    const res = await callWithRetry('gemini', () => fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
      {
        method: 'POST', signal: s,
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: SYSTEM }] },
          contents: [{ role: 'user', parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: maxTokens },
        }),
      }), { signal, onStatus, maxAttempts });
    const data = await res.json();
    const text = data?.candidates?.[0]?.content?.parts?.map((p) => p.text).filter(Boolean).join('') || '';
    if (!text && data?.promptFeedback?.blockReason) {
      throw new Error(`Gemini declined the request (${data.promptFeedback.blockReason}).`);
    }
    return { text, usage: data?.usageMetadata || null, headers: null };
  } finally { done(); }
}

async function callAnthropic({ model, prompt, signal, maxTokens, onStatus, maxAttempts }) {
  const key = getApiKey();
  const { signal: s, done } = makeSignal(signal, 120000);
  try {
    const res = await callWithRetry('anthropic', () => fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST', signal: s,
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model, max_tokens: maxTokens, temperature: 0.3, system: SYSTEM,
        messages: [{ role: 'user', content: prompt }],
      }),
    }), { signal, onStatus, maxAttempts });
    const headers = readRateLimitHeaders(res);
    const data = await res.json();
    const text = (data?.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    return { text, usage: data?.usage || null, headers };
  } finally { done(); }
}

async function callOpenAI({ model, prompt, signal, maxTokens, onStatus, maxAttempts }) {
  const key = getApiKey();
  const { signal: s, done } = makeSignal(signal, 120000);
  try {
    const res = await callWithRetry('openai', () => fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST', signal: s,
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
      body: JSON.stringify({
        model, temperature: 0.3, max_tokens: maxTokens,
        response_format: { type: 'json_object' },
        messages: [{ role: 'system', content: SYSTEM }, { role: 'user', content: prompt }],
      }),
    }), { signal, onStatus, maxAttempts });
    const data = await res.json();
    return { text: data?.choices?.[0]?.message?.content || '', usage: data?.usage || null, headers: null };
  } finally { done(); }
}

const ADAPTERS = { gemini: callGemini, anthropic: callAnthropic, openai: callOpenAI };

/* ============================== parsing ============================== */

export function parseModelJson(text) {
  if (!text) throw new Error('The model returned an empty response. This usually means the reply was cut short — try again, or reduce how much you run at once.');
  let t = String(text).trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  try { return JSON.parse(t); } catch { /* try harder */ }
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a >= 0 && b > a) {
    try { return JSON.parse(t.slice(a, b + 1)); } catch { /* fall through */ }
  }
  throw new Error('The model did not return valid JSON. Try again, or switch to a stronger model in Settings.');
}

/* ============================== orchestration ============================== */

/*
 * Overload memory ("circuit breaker").
 *
 * Without this, every site in a batch retried an overloaded model five times
 * before trying the fallback. On a ten-site run that spent about fifty of a
 * 200-a-day free allowance, and a minute and a half per site, on a model that
 * was never going to answer. Once a model has proved overloaded, the rest of
 * the run goes straight to the fallback, and it is only tried again once the
 * cool-off has passed.
 */
const OVERLOAD_COOLOFF_MS = 10 * 60 * 1000;
const overloadedUntil = new Map();   // model -> timestamp

export function isMarkedOverloaded(model, now = Date.now()) {
  const t = overloadedUntil.get(model);
  if (!t) return false;
  if (now >= t) { overloadedUntil.delete(model); return false; }
  return true;
}
export function overloadedModels(now = Date.now()) {
  return [...overloadedUntil.entries()]
    .filter(([m]) => isMarkedOverloaded(m, now))
    .map(([m, t]) => ({ model: m, minutesLeft: Math.ceil((t - now) / 60000) }));
}
function markOverloaded(model) { overloadedUntil.set(model, Date.now() + OVERLOAD_COOLOFF_MS); }

export function clearOverloadMemory(model) {
  if (model) overloadedUntil.delete(model); else overloadedUntil.clear();
}
/** Test hook. */
export function _markOverloaded(model, ms) { overloadedUntil.set(model, Date.now() + ms); }

/**
 * Run one analysis. Returns a report object ready to store.
 * @param {'site'|'master'} kind
 */
export async function run({ kind, payload, settings, followUp, signal, onStatus }) {
  if (!getApiKey()) throw new Error('No API key set. Add one in Settings — it stays in this browser and is never written to a session file.');
  const provider = PROVIDERS[settings.provider];
  if (!provider) throw new Error(`Unknown provider "${settings.provider}".`);
  const model = (settings.model || '').trim() || provider.defaultModel;
  const maxTokens = Number(settings.maxTokens) || 4096;

  const prompt = buildPrompt(kind, payload, followUp);
  const started = Date.now();

  let usedModel = model;
  let fellBack = false;
  let skippedPrimary = false;
  let out;
  const fallback = (settings.fallbackModel || '').trim();
  const hasFallback = !!fallback && fallback !== model;

  // A model already seen overloaded gets two tries, not five. Without this a
  // run where BOTH models are down spends about 95 seconds per site proving
  // something it already knew.
  const capFor = (mdl, withAlternative) =>
    (isMarkedOverloaded(mdl) ? 2 : (withAlternative ? 2 : undefined));

  if (hasFallback && isMarkedOverloaded(model) && !isMarkedOverloaded(fallback)) {
    skippedPrimary = true;
    onStatus?.({ why: `${model} is overloaded (seen in the last few minutes) — using ${fallback}`, attempt: 1, maxAttempts: 1, secondsLeft: 0 });
    try {
      out = await ADAPTERS[settings.provider]({ model: fallback, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(fallback, false) });
    } catch (e) {
      if (e?.overloaded) markOverloaded(fallback);
      throw e;
    }
    usedModel = fallback;
    fellBack = true;
  } else {
    try {
      out = await ADAPTERS[settings.provider]({
        model, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(model, hasFallback),
      });
    } catch (e) {
      if (e?.overloaded) markOverloaded(model);
      // Only an overload is worth a second model. A bad key, a 429 or a 404
      // would fail identically on any model, so those propagate unchanged.
      if (!e?.overloaded || !hasFallback || signal?.aborted) throw e;
      onStatus?.({ why: `${model} overloaded — switching to ${fallback}`, attempt: 1, maxAttempts: 1, secondsLeft: 0, providerMessage: e.providerMessage });
      try {
        out = await ADAPTERS[settings.provider]({ model: fallback, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(fallback, false) });
      } catch (e2) {
        if (e2?.overloaded) markOverloaded(fallback);
        if (e2?.overloaded) {
          // Say plainly that BOTH were tried, so the reader does not think the
          // fallback was never reached.
          const both = new Error(
            `Both models are overloaded right now: ${model} and the fallback ${fallback}.`
            + `${e2.providerMessage ? ` The provider said: "${e2.providerMessage}".` : ''}`
            + ' This is the provider\'s capacity for these models, not your key, your quota or your token count.'
            + ' Wait a few minutes, or set a more established model in Settings.');
          both.overloaded = true;
          both.status = e2.status;
          both.bothModels = [model, fallback];
          both.providerMessage = e2.providerMessage;
          throw both;
        }
        throw e2;
      }
      usedModel = fallback;
      fellBack = true;
    }
  }
  // A success clears any stale mark on whichever model answered.
  overloadedUntil.delete(usedModel);
  const { text, usage, headers } = out;
  const result = parseModelJson(text);

  const reported = readProviderUsage(usage);
  const tokens = reported || {
    input: estimateTokens(SYSTEM) + estimateTokens(prompt),
    output: estimateTokens(text),
  };
  recordUsage({ ...tokens, estimated: !reported });

  return {
    id: `${kind}-${Date.now()}`,
    kind,
    key: kind === 'master' ? '__master__' : payload.site,
    title: kind === 'master' ? 'Programme overview' : `Site ${payload.site}`,
    result,
    followUp: followUp || null,
    provider: settings.provider,
    providerLabel: provider.label,
    model: usedModel,
    requestedModel: model,
    fellBack,
    skippedPrimary,
    at: new Date().toISOString(),
    ms: Date.now() - started,
    tokens: { ...tokens, estimated: !reported },
    rateLimit: headers || null,
  };
}
