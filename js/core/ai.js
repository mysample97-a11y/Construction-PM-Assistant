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
    keyHint: 'Starts with "AIza"',
    keyUrl: 'https://aistudio.google.com/apikey',
    defaultModel: 'gemini-2.5-flash',
    suggestions: ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash'],
    note: 'Calls go directly from your browser to Google.',
  },
  anthropic: {
    label: 'Anthropic Claude',
    keyHint: 'Starts with "sk-ant-"',
    keyUrl: 'https://console.anthropic.com/settings/keys',
    defaultModel: 'claude-sonnet-4-5',
    suggestions: ['claude-sonnet-4-5', 'claude-haiku-4-5', 'claude-opus-4-1'],
    note: 'Browser calls need direct-access enabled on the request, which this app sets.',
  },
  openai: {
    label: 'OpenAI',
    keyHint: 'Starts with "sk-"',
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

/* ============================== payloads ============================== */

/**
 * Compact, human-checkable payload for one site.
 *
 * `noData` is derived rather than trusted: a caller that hands in a missing or
 * half-built analysis must produce an empty payload, not crash the render. A
 * thrown error here takes the whole page down, which is a far worse outcome
 * than a site reporting that it has nothing to say.
 */
export function buildSitePayload(analysis) {
  const a = analysis && typeof analysis === 'object' ? analysis : {};
  const trim = (arr, n) => (Array.isArray(arr) ? arr : []).slice(0, n);
  const noData = a.noData !== false || !Array.isArray(a.velocity);
  if (noData) {
    return {
      site: a.code || '(unknown)',
      description: a.description || undefined,
      targetSubmission: a.target || undefined,
      hasWeeklyData: false,
      note: 'This site is in the register but has no weekly status data, so nothing about its progress can be measured. Say so plainly rather than inferring anything.',
      computedRisks: trim(a.risks, 10),
    };
  }
  return {
    site: a.code,
    description: a.description || undefined,
    wave: a.wave || undefined,
    coordinator: a.coordinator || undefined,
    startDate: a.start || undefined,
    targetSubmission: a.target || undefined,
    reportingWeek: a.reportingWeek
      ? { label: a.reportingWeek.label, weekEnding: a.reportingWeek.end }
      : null,
    timeline: a.timeline || null,
    // What the last review said, so this one can speak to what changed rather
    // than starting from nothing every week.
    previousPeriod: a.previousPeriod || null,
    computedNote: 'All figures below were computed by the application from the weekly status grid. Do not recalculate them.',
    hasWeeklyData: true,

    progress: {
      tasksLive: a.liveCount,
      tasksFinished: a.finishedCount,
      tasksWip: a.wipCount,
      tasksBlocked: a.blockedCount,
      tasksWaitingOn: a.waitingCount,
      tasksNotStarted: a.notStartedCount,
      tasksNotApplicable: a.naCount,
      percentByCount: a.pctByCount,
      percentByWeight: a.pctByWeight,
      remaining: a.remaining,
    },
    throughput: {
      perWeek: a.velocity.map((v) => ({ week: v.label, finished: v.count })),
      averagePerWeek: a.avgVelocity,
      lastThreeWeeksAverage: a.recentVelocity,
    },
    forecast: {
      basis: 'remaining tasks divided by completion rate',
      usingAllTimeRate: a.forecastAvg,
      usingRecentRate: a.forecastRecent,
      targetWeekNumber: a.targetWeekNo,
      weeksPastTarget: a.slipWeeks,
    },
    categories: a.categories.map((c) => ({
      id: c.id, name: c.name,
      yourStatus: c.selfStatus,
      tasksDone: `${c.doneCount}/${c.liveCount}`,
      computedPercent: c.computedPct,
      divergenceFromYourStatus: c.divergence,
      stuckTasks: c.stuck,
    })),
    stuck: trim(a.stuckDetail, 20).map((s) => ({
      taskId: s.taskId, task: s.name, category: s.category, discipline: s.discipline,
      status: s.status, weeksStuck: s.weeksStuck, since: s.sinceWeek,
      reason: s.reason || '(no reason logged)',
      waitingOn: s.waitingOn || undefined,
      expectedClear: s.expected || undefined,
    })),
    stalledInWip: trim(a.stalled, 12),
    pastTargetWeek: trim(a.overdue, 20),
    outstandingPrerequisites: trim(a.outstandingPrereqs, 20),
    scopeGrowth: {
      addedAfterKickoff: a.scopeGrowth.additionalCount,
      originallyAgreed: a.scopeGrowth.regularCount,
      growthPercent: a.scopeGrowth.growthPct,
      items: trim(a.scopeGrowth.items, 15),
    },
    runningAheadOfDependencies: trim(a.outOfOrder, 12),
    resourceLoad: trim(a.resources, 15),
    dataQualityIssues: trim(a.dataIssues, 20).map((d) => d.detail),
    computedRisks: trim(a.risks, 20),
  };
}

export function buildMasterPayload(p) {
  return {
    scope: 'portfolio',
    generatedAt: p.generatedAt,
    computedNote: 'All figures were computed by the application. Do not recalculate them.',
    totals: {
      sites: p.siteCount,
      sitesWithData: p.sitesWithData,
      tasksLive: p.totalTasks,
      tasksFinished: p.totalDone,
      percentByCount: p.pctByCount,
      percentByWeight: p.pctByWeight,
      stuckTasks: p.totalStuck,
      tasksPastTargetWeek: p.totalOverdue,
      tasksAddedAfterKickoff: p.totalAdditional,
    },
    sites: p.sites.map((s) => ({
      site: s.code,
      wave: s.wave || undefined,
      coordinator: s.coordinator || undefined,
      targetSubmission: s.target || undefined,
      hasData: !s.noData,
      percentByWeight: s.noData ? null : s.pctByWeight,
      tasksLive: s.noData ? null : s.liveCount,
      remaining: s.noData ? null : s.remaining,
      recentRatePerWeek: s.noData ? null : s.recentVelocity,
      forecastWeeksNeeded: s.noData ? null : s.forecastRecent?.weeksNeeded ?? null,
      weeksPastTarget: s.noData ? null : s.slipWeeks,
      stuckTasks: s.noData ? null : s.stuckCount,
      topRisk: s.risks?.[0]?.title || null,
    })),
    sitesForecastLate: p.slipping,
    resourceLoadAcrossSites: p.resourceLoad.slice(0, 20),
    blockersAffectingMoreThanOneSite: p.commonBlockers,
    waves: p.waves,
    computedRisks: p.risks,
  };
}

/* ============================== schemas ============================== */

const SITE_SCHEMA = `{
  "introduction": "2-3 sentences introducing this site: what it is, where it sits in its own programme, and the single thing a reader needs to know before the detail. Do not list figures.",

  "timelineNote": "2-4 sentences on time: how far through the planned period this site is against how much work is done, and whether those two are in step. If a previous analysis is given, say what has changed since it. Nothing about individual tasks here.",

  "taskStatusInterpretation": "3-5 sentences interpreting the task position: what the completion rate and the mix of WIP/not-started actually mean for delivery, which categories are carrying the work and which have not started. Do not mention blocked or waiting work here — that has its own section.",

  "prerequisiteInterpretation": "2-4 sentences on prerequisites: what is outstanding, what it is holding up, and whether the pattern suggests an upstream problem. If none are outstanding, say so in one line and move on.",

  "blockedInterpretation": "3-5 sentences on blocked and waiting work: what is stuck, for how long, on whom, and what it will cost if it stays stuck. Distinguish Blocked (inside the team's control) from Waiting on (outside it) because the response differs.",

  "additional": {
    "risks": [{"risk": "...", "why": "the evidence from the input", "impact": "high"|"medium"|"low"}],
    "patterns": ["something the figures reveal that is not obvious from any single one of them"],
    "actions": [{"action": "...", "owner": "R code from the input where known, else 'not recorded'", "byWhen": "a week number or date", "priority": "high"|"medium"|"low", "expectedEffect": "..."}],
    "watchNextWeek": ["the specific thing that would tell you early if this is getting worse"]
  },

  "visualisationNote": "1-3 sentences telling the reader what to look for in the charts below — the shape that matters, not a description of the axes.",

  "conclusions": {
    "verdict": "on_track" | "at_risk" | "off_track",
    "confidence": "high" | "medium" | "low",
    "confidenceReason": "what limits confidence",
    "statement": "3-4 sentences drawing the threads together in NEW words. Do not repeat sentences from earlier sections.",
    "nextSteps": ["the two or three things that must happen before the next review"]
  }
}`;

const MASTER_SCHEMA = `{
  "introduction": "2-3 sentences introducing the programme: how many sites, what they have in common, and the headline position.",

  "timelineNote": "2-4 sentences on where the sites sit against their own dates, including which are in the same wave and therefore competing for the same people. If a previous analysis is given, say what has moved since.",

  "taskStatusInterpretation": "3-5 sentences comparing the sites: who is ahead, who is behind, and whether the spread is explained by start dates or by something else.",

  "prerequisiteInterpretation": "2-4 sentences on prerequisites across the sites, especially any provider appearing on more than one.",

  "blockedInterpretation": "3-5 sentences on blocked and waiting work across the programme. A blocker on one site is a site problem; the same blocker on several is a process problem — say which you are looking at.",

  "additional": {
    "systemicIssues": [{"issue": "...", "sitesAffected": ["A-01"], "rootCauseHypothesis": "...", "howToTest": "what to check to confirm it", "fixOnceCentrally": "..."}],
    "resourceConcerns": [{"resource": "R code", "concern": "...", "suggestedAction": "..."}],
    "actions": [{"action": "...", "owner": "...", "byWhen": "...", "priority": "high"|"medium"|"low", "affectsSites": ["A-01"]}],
    "whatIsGoingWell": ["worth saying — a report that is only bad news gets discounted"]
  },

  "visualisationNote": "1-3 sentences on what to look for in the charts below.",

  "conclusions": {
    "verdict": "on_track" | "at_risk" | "off_track",
    "confidence": "high" | "medium" | "low",
    "confidenceReason": "...",
    "statement": "3-4 sentences in NEW words, not a repeat of the sections above.",
    "nextSteps": ["..."]
  }
}`;

/* ============================== prompts ============================== */

export function buildPrompt(kind, payload, followUp = '') {
  const parts = [];
  if (kind === 'master') {
    parts.push('Review this multi-site BIM delivery programme as a whole. Compare the sites against each other, find what is systemic rather than local, and say what the coordinator should do next week.');
    parts.push('\nReturn JSON matching exactly this shape:\n' + MASTER_SCHEMA);
  } else {
    parts.push(`Review site ${payload.site} for this reporting week.`);
    parts.push('\nReturn JSON matching exactly this shape:\n' + SITE_SCHEMA);
  }
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
 * One HTTP attempt with retry.
 *
 * 529/503 are server-side overload and worth retrying with backoff.
 * 429 is a real rate limit: retrying makes it worse and burns request quota,
 * so it fails fast with an explanation instead.
 */
async function callWithRetry(provider, doFetch, { signal, maxAttempts = 3 } = {}) {
  let lastErr = null;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    if (signal?.aborted) throw new CancelledError();
    recordRequest(provider, attempt === 0 ? 'first attempt' : `retry after ${lastErr?.status || 'error'}`);
    let response;
    try {
      response = await doFetch();
    } catch (e) {
      if (signal?.aborted) throw new CancelledError();
      if (e?.name === 'AbortError') throw new Error('The request timed out. Try fewer sites at once.');
      lastErr = e;
      if (attempt === maxAttempts - 1) throw new Error(`Could not reach ${provider}. Check your connection. (${e.message})`);
      await sleep(800 * (attempt + 1), signal);
      continue;
    }

    if (response.ok) return response;

    const body = await response.text().catch(() => '');
    if (response.status === 429) {
      const wait = retryAfterSeconds(response, body);
      throw new Error(
        `Rate limit reached (429).${wait ? ` The provider asked to wait about ${wait}s.` : ''} ` +
        'Retrying immediately would only burn more of your request quota. Wait, then run fewer sites at a time.');
    }
    if (response.status === 401 || response.status === 403) {
      throw new Error(`${provider} rejected the API key (${response.status}). Check it is correct, active, and allowed for this model.`);
    }
    if (response.status === 404) {
      throw new Error(`${provider} does not recognise that model name (404). Model names change — check the current list in Settings.`);
    }
    if (response.status === 529 || response.status === 503 || response.status === 500) {
      lastErr = { status: response.status };
      if (attempt === maxAttempts - 1) {
        throw new Error(`${provider} is overloaded (${response.status}) and did not recover after ${maxAttempts} attempts. Try again shortly.`);
      }
      await sleep(1200 * Math.pow(2, attempt), signal);
      continue;
    }
    let msg = body.slice(0, 300);
    try { msg = JSON.parse(body)?.error?.message || msg; } catch { /* keep raw */ }
    throw new Error(`${provider} returned ${response.status}: ${msg}`);
  }
  throw new Error('Request failed.');
}

async function callGemini({ model, prompt, signal, maxTokens }) {
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
      }), { signal });
    const data = await res.json();
    const text = data?.candidates?.[0]?.content?.parts?.map((p) => p.text).filter(Boolean).join('') || '';
    if (!text && data?.promptFeedback?.blockReason) {
      throw new Error(`Gemini declined the request (${data.promptFeedback.blockReason}).`);
    }
    return { text, usage: data?.usageMetadata || null, headers: null };
  } finally { done(); }
}

async function callAnthropic({ model, prompt, signal, maxTokens }) {
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
    }), { signal });
    const headers = readRateLimitHeaders(res);
    const data = await res.json();
    const text = (data?.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    return { text, usage: data?.usage || null, headers };
  } finally { done(); }
}

async function callOpenAI({ model, prompt, signal, maxTokens }) {
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
    }), { signal });
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

/**
 * Run one analysis. Returns a report object ready to store.
 * @param {'site'|'master'} kind
 */
export async function run({ kind, payload, settings, followUp, signal }) {
  if (!getApiKey()) throw new Error('No API key set. Add one in Settings — it stays in this browser and is never written to a session file.');
  const provider = PROVIDERS[settings.provider];
  if (!provider) throw new Error(`Unknown provider "${settings.provider}".`);
  const model = (settings.model || '').trim() || provider.defaultModel;
  const maxTokens = Number(settings.maxTokens) || 4096;

  const prompt = buildPrompt(kind, payload, followUp);
  const started = Date.now();
  const { text, usage, headers } = await ADAPTERS[settings.provider]({ model, prompt, signal, maxTokens });
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
    model,
    at: new Date().toISOString(),
    ms: Date.now() - started,
    tokens: { ...tokens, estimated: !reported },
    rateLimit: headers || null,
  };
}
