/**
 * Token and request accounting.
 *
 * Two dimensions are tracked, because they fail differently:
 *   TOKENS   — what a run costs.
 *   REQUESTS — what actually throttles you. Free tiers publish generous token
 *              allowances alongside a low requests-per-minute cap, so a
 *              token-only meter shows comfortable headroom right up to the
 *              moment a 429 lands.
 *
 * Pre-flight figures are ESTIMATES and are labelled as such everywhere they
 * appear. Real tokenisation is model-specific and cannot be computed in the
 * browser without shipping the tokeniser. Provider-reported usage always
 * overrides the estimate once a call returns.
 */

const TOK_KEY = 'bimtrack:tokens:v1';
const REQ_KEY = 'bimtrack:requests:v1';
const LIM_KEY = 'bimtrack:limits:v1';
const DAY_MS = 86400000;
const MIN_MS = 60000;

/* ------------------------------ estimation ------------------------------ */

/**
 * ~4 characters per token for English prose. Structured text (JSON) is denser
 * because punctuation creates more token boundaries, so a surcharge applies.
 */
export function estimateTokens(text) {
  if (!text) return 0;
  const s = String(text);
  const structural = (s.match(/[{}[\]":,]/g) || []).length;
  const density = structural / Math.max(s.length, 1) > 0.04 ? 3.2 : 4.0;
  return Math.ceil(s.length / density);
}

export function estimateRun({ systemText = '', userText = '', maxTokens = 4096, calls = 1 } = {}) {
  const input = estimateTokens(systemText) + estimateTokens(userText);
  // Output rarely reaches the cap; ~70% is a fair planning figure and
  // over-estimating is the safer error because it warns earlier.
  const output = Math.ceil(maxTokens * 0.7);
  return { input: input * calls, output: output * calls, total: (input + output) * calls, calls };
}

/* ------------------------------ accounting ------------------------------ */

const EMPTY = { input: 0, output: 0, total: 0, calls: 0, lastRun: null, estimated: false };

function read(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    const p = JSON.parse(raw);
    return p && typeof p === 'object' ? p : fallback;
  } catch { return fallback; }
}
function write(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch { /* never fail a run over accounting */ }
}

export function getUsage() { return read(TOK_KEY, null) || { ...EMPTY }; }

export function recordUsage({ input = 0, output = 0, estimated = false } = {}) {
  const cur = getUsage();
  const next = {
    input: cur.input + (Number(input) || 0),
    output: cur.output + (Number(output) || 0),
    total: cur.total + (Number(input) || 0) + (Number(output) || 0),
    calls: cur.calls + 1,
    lastRun: new Date().toISOString(),
    // Sticky: once any component is an estimate, the total is.
    estimated: cur.estimated || !!estimated,
  };
  write(TOK_KEY, next);
  return next;
}

export function resetUsage() {
  try { localStorage.removeItem(TOK_KEY); } catch { /* ignore */ }
  return { ...EMPTY };
}

/** Normalise a provider usage object. Returns null when nothing was reported. */
export function readProviderUsage(raw) {
  if (!raw || typeof raw !== 'object') return null;
  if (typeof raw.input_tokens === 'number' || typeof raw.output_tokens === 'number') {
    return { input: Number(raw.input_tokens) || 0, output: Number(raw.output_tokens) || 0 };
  }
  if (typeof raw.promptTokenCount === 'number' || typeof raw.candidatesTokenCount === 'number') {
    return { input: Number(raw.promptTokenCount) || 0, output: Number(raw.candidatesTokenCount) || 0 };
  }
  if (typeof raw.prompt_tokens === 'number' || typeof raw.completion_tokens === 'number') {
    return { input: Number(raw.prompt_tokens) || 0, output: Number(raw.completion_tokens) || 0 };
  }
  return null;
}

/* --------------------------- request windows --------------------------- */

function reqKey(provider) { return `${REQ_KEY}:${provider}`; }

function readReq(provider) {
  const list = read(reqKey(provider), []);
  return Array.isArray(list) ? list.filter((t) => typeof t === 'number') : [];
}

/**
 * Records one HTTP request. Retries count too — each is a real request the
 * provider counts, which is why the number can rise by more than one per click.
 */
export function recordRequest(provider, reason = 'call', now = Date.now()) {
  const list = readReq(provider).filter((t) => now - t < DAY_MS);
  list.push(now);
  write(reqKey(provider), list);
  const log = read(`${REQ_KEY}:log`, []);
  write(`${REQ_KEY}:log`, [...(Array.isArray(log) ? log : []), { at: now, provider, reason }].slice(-20));
  return requestWindows(provider, now, list);
}

export function requestWindows(provider, now = Date.now(), preloaded = null) {
  const list = (preloaded || readReq(provider)).filter((t) => now - t < DAY_MS);
  const inMin = list.filter((t) => now - t < MIN_MS).sort((a, b) => a - b);
  return {
    lastMinute: inMin.length,
    lastDay: list.length,
    nextMinuteSlotIn: inMin.length ? Math.max(0, Math.ceil((MIN_MS - (now - inMin[0])) / 1000)) : 0,
  };
}

export function recentRequestReasons(now = Date.now()) {
  const log = read(`${REQ_KEY}:log`, []);
  const recent = (Array.isArray(log) ? log : []).filter((r) => r && now - r.at < 120000);
  const counts = {};
  for (const r of recent) counts[r.reason] = (counts[r.reason] || 0) + 1;
  return Object.entries(counts).map(([reason, n]) => ({ reason, n }));
}

export function resetRequests(provider) {
  write(reqKey(provider), []);
  return { lastMinute: 0, lastDay: 0, nextMinuteSlotIn: 0 };
}

/* ------------------------------- limits ------------------------------- */

/**
 * Seed values only. Providers change these without notice, so they are
 * defaults the user can edit — never assertions of fact. The UI says so.
 */
export const PUBLISHED_LIMITS = {
  gemini: {
    free: { rpm: 15, rpd: 200, tpm: 250000, label: 'Gemini free tier (per model)',
            caution: 'Limits are per model and Google revises them. Check your own quota and edit these to match. Free-tier prompts may be used by Google to improve their products.' },
    paid: { rpm: 150, rpd: null, tpm: 1000000, label: 'Gemini paid (Tier 1)' },
  },
  anthropic: {
    free: { rpm: 5, rpd: null, tpm: 20000, label: 'Anthropic evaluation tier',
            caution: 'New organisations start on reduced limits that rise with usage history.' },
    paid: { rpm: 50, rpd: null, tpm: 40000, label: 'Anthropic Tier 1' },
  },
  openai: {
    free: { rpm: 3, rpd: 200, tpm: 40000, label: 'OpenAI free/trial',
            caution: 'Trial limits are low and vary by model.' },
    paid: { rpm: 500, rpd: null, tpm: 200000, label: 'OpenAI Tier 1' },
  },
};

export function getLimits(provider = 'gemini') {
  const p = PUBLISHED_LIMITS[provider] ? provider : 'gemini';
  const declared = read(LIM_KEY, {})[p] || null;
  const tier = declared?.tier || 'free';
  const base = PUBLISHED_LIMITS[p][tier] || PUBLISHED_LIMITS[p].free;
  return {
    provider: p,
    tier,
    rpm: declared?.rpm != null ? Number(declared.rpm) : base.rpm,
    rpd: declared?.rpd != null ? Number(declared.rpd) : base.rpd,
    tpm: declared?.tpm != null ? Number(declared.tpm) : base.tpm,
    label: base.label,
    caution: base.caution || '',
    isDefault: !declared,
  };
}

export function saveLimits(provider, { tier, rpm, rpd, tpm } = {}) {
  const all = read(LIM_KEY, {});
  all[provider] = {
    tier: tier || 'free',
    rpm: rpm === '' || rpm == null ? null : Number(rpm),
    rpd: rpd === '' || rpd == null ? null : Number(rpd),
    tpm: tpm === '' || tpm == null ? null : Number(tpm),
  };
  write(LIM_KEY, all);
  return getLimits(provider);
}

/**
 * Verdict on whether a planned run will fit. Requests are checked first,
 * because that is the dimension that actually throttles this usage pattern.
 */
export function capacityCheck(provider, estimate, now = Date.now()) {
  const lim = getLimits(provider);
  const win = requestWindows(provider, now);
  const est = estimate?.total || 0;
  const calls = estimate?.calls || 1;

  if (lim.rpd && win.lastDay + calls > lim.rpd) {
    return { level: 'high', dimension: 'requests/day',
      message: `Daily request limit reached (${win.lastDay} of ${lim.rpd}). This resets on a rolling 24-hour window.` };
  }
  if (lim.rpm && win.lastMinute + calls > lim.rpm) {
    return { level: 'high', dimension: 'requests/minute',
      message: `Per-minute request limit reached (${win.lastMinute} of ${lim.rpm}). Capacity returns in about ${win.nextMinuteSlotIn}s.` };
  }
  if (lim.tpm && est > lim.tpm) {
    return { level: 'high', dimension: 'tokens/minute',
      message: `This run is estimated at about ${formatTokens(est)} tokens against a ${formatTokens(lim.tpm)}/min limit. Run fewer sites at a time.` };
  }
  if (lim.rpm && win.lastMinute + calls > lim.rpm * 0.7) {
    return { level: 'medium', dimension: 'requests/minute',
      message: `Approaching the per-minute request limit (${win.lastMinute} of ${lim.rpm}).` };
  }
  if (lim.tpm && est > lim.tpm * 0.6) {
    return { level: 'medium', dimension: 'tokens/minute',
      message: 'Large run against the declared per-minute token limit. Consider running sites in two batches.' };
  }
  return { level: 'low', dimension: null, message: 'Within declared limits.' };
}

export function formatTokens(n) {
  const v = Number(n) || 0;
  if (v >= 1e6) return `${(v / 1e6).toFixed(2)}M`;
  if (v >= 1000) return `${(v / 1000).toFixed(1)}k`;
  return String(v);
}

/* --------------------------- exact counting --------------------------- */

/**
 * Exact input token count from the provider. Free in tokens, but it costs one
 * REQUEST against RPM/RPD — which is the binding limit — so it is only called
 * when the user explicitly asks for a precise figure, never on every keystroke.
 * Returns null on any failure so the caller falls back to the estimate.
 */
export async function countTokensExact({ provider, apiKey, model, systemText, userText }) {
  try {
    if (provider === 'anthropic') {
      const r = await fetch('https://api.anthropic.com/v1/messages/count_tokens', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': (apiKey || '').trim(),
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model, system: systemText || undefined,
          messages: [{ role: 'user', content: String(userText || '') }],
        }),
      });
      recordRequest(provider, 'token count');
      if (!r.ok) return null;
      const d = await r.json().catch(() => null);
      const n = d?.input_tokens;
      return Number.isFinite(n) ? { input: n, exact: true } : null;
    }
    if (provider === 'gemini') {
      const r = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:countTokens`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-goog-api-key': (apiKey || '').trim() },
          body: JSON.stringify({ contents: [{ parts: [{ text: `${systemText || ''}\n${userText || ''}` }] }] }),
        });
      recordRequest(provider, 'token count');
      if (!r.ok) return null;
      const d = await r.json().catch(() => null);
      const n = d?.totalTokens;
      return Number.isFinite(n) ? { input: n, exact: true } : null;
    }
    return null;   // OpenAI has no free counting endpoint
  } catch { return null; }
}

/**
 * Anthropic reports remaining quota as response headers. Browsers can only read
 * them if the server lists them in Access-Control-Expose-Headers, which cannot
 * be determined from documentation — only at runtime. Returns null when they
 * are unreadable so the caller degrades to the estimate rather than showing
 * blanks.
 */
export function readRateLimitHeaders(response) {
  if (!response || typeof response.headers?.get !== 'function') return null;
  const num = (n) => {
    const v = response.headers.get(n);
    if (v == null || v === '') return null;
    const p = Number(v);
    return Number.isFinite(p) ? p : null;
  };
  const out = {
    requestsRemaining: num('anthropic-ratelimit-requests-remaining'),
    requestsLimit: num('anthropic-ratelimit-requests-limit'),
    inputRemaining: num('anthropic-ratelimit-input-tokens-remaining'),
    outputRemaining: num('anthropic-ratelimit-output-tokens-remaining'),
    retryAfter: num('retry-after'),
  };
  return Object.values(out).some((v) => v !== null) ? out : null;
}
