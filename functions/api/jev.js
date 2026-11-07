// POST /api/jev: asks the Jev decision model (DIT AI-ONE gateway, POST /decisions/) typed questions about one text.
// Jev does not write text: it answers yes/no, one-of-N and ordered-score questions with probabilities and a confidence.
// The browser (or an Apps Script) never sees the API key; only the whitelisted fields below reach the gateway.
// Request:  { code, state:<string|object>, questions:{ name:{ type:'noul'|'choice'|'score', instructions, criteria? } } } or { code, check:true }
//   noul   criteria optional {"true":"...","false":"..."}
//   choice criteria { option: description } (2 to MAX_CHOICES)
//   score  criteria ["lowest", ..., "highest"] (2 to 10)
// Success:  { ok:true, model, answers:{ name: {type:'noul', p} | {type:'choice', choice, confidence, probs}
//                                              | {type:'score', score, level, label, confidence, probs} }, usage }
// Failure:  JSON { error:<Korean>, code:bad_code|rate|budget|disabled|too_long|bad_request|upstream[, retryAfter][, detail] }
import {
  readLimits,
  isKillSwitchOff,
  readJsonBody,
  parseCodes,
  matchCode,
  clientKey,
  createStore,
  badCodeGate,
  admit,
} from '../_lib/limits.js';
import { isAllowedOrigin, jsonResponse, errorResponse, preflightResponse } from '../_lib/cors.js';

const DEFAULT_BASE = 'https://factchat-cloud.mindlogic.ai/v1/gateway';
const DEFAULT_MODEL = 'jev-latest';
const TIMEOUT_MS = 25_000;

export const JEV_LIMITS = {
  maxBodyBytes: 160 * 1024,
  maxCode: 64,
  maxState: 20000, // chars; Korean is ~1 token per char and the model budget is ~32k tokens for state + the longest question
  maxQuestions: 10,
  maxInstructions: 1500,
  maxChoices: 12,
  maxDescription: 300,
  maxLevels: 10,
  maxLevelLabel: 60,
};

const MSG = {
  disabled: 'AI 기능이 지금 꺼져 있습니다. 강사에게 문의해 주세요.',
  method: '허용되지 않은 요청 방식입니다.',
  origin: '허용되지 않은 사이트에서 보낸 요청입니다.',
  noCode: '수업 코드를 입력해 주세요.',
  longCode: '수업 코드가 너무 깁니다. 강사가 알려 준 코드를 확인해 주세요.',
  badCode: '수업 코드가 맞지 않습니다. 강사가 알려 준 코드를 확인해 주세요.',
  auth: 'AI ONE 연결 설정에 문제가 있습니다. 강사에게 알려 주세요.',
  credit: 'AI ONE 크레딧이 부족합니다. 강사에게 알려 주세요.',
  rate: 'AI ONE 서버에 요청이 몰렸습니다. 1분쯤 뒤에 다시 시도해 주세요.',
  noModel: 'Jev 판정 모델을 이 계정에서 쓸 수 없습니다. 강사에게 알려 주세요.',
  slow: 'Jev 판정 서버 응답이 너무 늦습니다. 잠시 후 다시 시도해 주세요.',
  busy: 'Jev 판정 서버가 바쁩니다. 잠시 후 다시 시도해 주세요.',
  rejected: 'Jev가 질문 형식을 받아들이지 못했습니다. 질문을 확인해 주세요.',
  internal: '판정 요청을 처리하지 못했습니다. 잠시 후 다시 시도해 주세요.',
};

const NAME_RE = /^[A-Za-z0-9_가-힣-]{1,40}$/;
const TYPES = new Set(['noul', 'choice', 'score']);

function bad(message, code = 'bad_request', status = 400) {
  return { ok: false, status, code, message };
}

const isPlain = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
const isText = (v, max) => typeof v === 'string' && v.trim().length > 0 && v.length <= max;

// Returns { ok:true, state:<string>, questions } holding only whitelisted fields, or { ok:false, status, code, message }.
export function validateJev(body) {
  const L = JEV_LIMITS;
  let state = body.state;
  if (state !== null && typeof state === 'object') state = JSON.stringify(state);
  if (typeof state !== 'string' || state.trim().length === 0) return bad('판정할 글(state)이 비어 있습니다.');
  if (state.length > L.maxState) return bad(`판정할 글이 너무 깁니다. ${L.maxState.toLocaleString('en-US')}자 안으로 줄여 주세요.`, 'too_long', 413);
  if (!isPlain(body.questions)) return bad('질문(questions)은 { 이름: 질문 } 형태여야 합니다.');
  const names = Object.keys(body.questions);
  if (names.length === 0) return bad('질문이 하나도 없습니다.');
  if (names.length > L.maxQuestions) return bad(`질문은 ${L.maxQuestions}개까지 보낼 수 있습니다.`, 'too_long', 413);
  const questions = {};
  for (const name of names) {
    const q = body.questions[name];
    const at = `질문 «${name.slice(0, 40)}»`;
    if (!NAME_RE.test(name)) return bad(`${at}의 이름은 글자, 숫자, _, - 40자 안으로 써 주세요.`);
    if (!isPlain(q) || !TYPES.has(q.type)) return bad(`${at}의 type은 noul, choice, score 중 하나여야 합니다.`);
    if (!isText(q.instructions, L.maxInstructions)) return bad(`${at}의 instructions를 ${L.maxInstructions}자 안의 글로 적어 주세요.`);
    const out = { type: q.type, instructions: q.instructions };
    if (q.type === 'noul') {
      if (q.criteria != null) {
        const c = q.criteria;
        if (!isPlain(c) || !isText(c.true, L.maxDescription) || !isText(c.false, L.maxDescription)) {
          return bad(`${at}의 criteria는 {"true":"…","false":"…"} 형태여야 합니다.`);
        }
        out.criteria = { true: c.true, false: c.false };
      }
    } else if (q.type === 'choice') {
      const keys = isPlain(q.criteria) ? Object.keys(q.criteria) : [];
      if (keys.length < 2 || keys.length > L.maxChoices) return bad(`${at}의 선택지(criteria)는 2개 이상 ${L.maxChoices}개 이하의 { 선택지: 설명 }이어야 합니다.`);
      const crit = {};
      for (const k of keys) {
        if (!NAME_RE.test(k)) return bad(`${at}의 선택지 이름 «${k.slice(0, 40)}»은 글자, 숫자, _, - 40자 안으로 써 주세요.`);
        if (!isText(q.criteria[k], L.maxDescription)) return bad(`${at}의 선택지 «${k}» 설명을 ${L.maxDescription}자 안의 글로 적어 주세요.`);
        crit[k] = q.criteria[k];
      }
      out.criteria = crit;
    } else {
      const c = q.criteria;
      if (!Array.isArray(c) || c.length < 2 || c.length > L.maxLevels || !c.every((x) => isText(x, L.maxLevelLabel))) {
        return bad(`${at}의 단계(criteria)는 낮은 단계부터 적은 글 2개 이상 ${L.maxLevels}개 이하의 배열이어야 합니다.`);
      }
      out.criteria = c.slice();
    }
    questions[name] = out;
  }
  return { ok: true, state, questions };
}

const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : null);

// Gateway answers -> one small shape per type, so callers do not parse the vendor format.
export function flattenAnswers(answers) {
  const out = {};
  if (!isPlain(answers)) return out;
  for (const [name, a] of Object.entries(answers)) {
    if (!isPlain(a)) continue;
    if (a.type === 'noul') out[name] = { type: 'noul', p: num(a.noul) };
    else if (a.type === 'choice') {
      out[name] = { type: 'choice', choice: typeof a.choice === 'string' ? a.choice : null, confidence: num(a.confidence), probs: isPlain(a.probabilities) ? a.probabilities : {} };
    } else if (a.type === 'score') {
      const probs = isPlain(a.probabilities) ? a.probabilities : {};
      let level = null;
      let best = -1;
      for (const [k, p] of Object.entries(probs)) if (typeof p === 'number' && p > best) [level, best] = [Number(k), p];
      if (!Number.isInteger(level)) level = num(a.score) == null ? null : Math.round(a.score);
      const legend = isPlain(a.legend) ? a.legend : {};
      out[name] = { type: 'score', score: num(a.score), level, label: level == null ? null : (legend[String(level)] ?? null), confidence: num(a.confidence), probs };
    }
  }
  return out;
}

function upstreamMessage(text) {
  try {
    const j = JSON.parse(text);
    const m = (j && j.error && (j.error.message || j.error)) || (j && (j.detail || j.message));
    return typeof m === 'string' ? m.slice(0, 300) : '';
  } catch {
    return '';
  }
}

export async function onRequest(context) {
  const origin = context.request.headers.get('Origin');
  try {
    return await handle(context, origin);
  } catch (e) {
    console.error('[jev] internal error', e && e.name);
    return errorResponse(503, 'upstream', MSG.internal, origin);
  }
}

async function handle(context, origin) {
  const { request, env } = context;
  const waitUntil = typeof context.waitUntil === 'function' ? context.waitUntil.bind(context) : undefined;

  if (request.method === 'OPTIONS') return preflightResponse(origin);
  if (request.method !== 'POST') {
    const r = errorResponse(405, 'bad_request', MSG.method, origin);
    r.headers.set('Allow', 'POST, OPTIONS');
    return r;
  }
  if (origin != null && !isAllowedOrigin(origin)) return errorResponse(403, 'bad_request', MSG.origin, origin);

  const codesRaw = String(env.AI_CLASS_CODE ?? '');
  if (isKillSwitchOff(env) || parseCodes(codesRaw).length === 0) return errorResponse(503, 'disabled', MSG.disabled, origin);
  const apiKey = String(env.AI_GATEWAY_API_KEY ?? '').trim();
  if (!apiKey) {
    console.error('[jev] AI_GATEWAY_API_KEY is not set');
    return errorResponse(503, 'disabled', MSG.disabled, origin);
  }

  const parsed = await readJsonBody(request, JEV_LIMITS.maxBodyBytes);
  if (!parsed.ok) return errorResponse(parsed.status, parsed.code, parsed.message, origin);
  const body = parsed.body;
  if (typeof body.code !== 'string' || !body.code.trim()) return errorResponse(401, 'bad_code', MSG.noCode, origin);
  if (body.code.length > JEV_LIMITS.maxCode) return errorResponse(400, 'bad_request', MSG.longCode, origin);

  // Access code before any budget: a wrong code never reaches the gateway.
  const limits = readLimits(env);
  const store = createStore(env, waitUntil);
  const ip = clientKey(request);
  const gate = await badCodeGate(store, ip, Date.now(), limits.badCode10m);
  if (gate.locked) return errorResponse(gate.locked.status, gate.locked.code, gate.locked.message, origin, gate.locked.retryAfter);
  const codeId = await matchCode(body.code, codesRaw);
  if (!codeId) {
    gate.recordWrong();
    return errorResponse(401, 'bad_code', MSG.badCode, origin);
  }
  if (body.check === true) return jsonResponse(200, { ok: true }, origin);

  const v = validateJev(body);
  if (!v.ok) return errorResponse(v.status, v.code, v.message, origin);

  const adm = await admit(store, { ip, codeId, now: Date.now(), limits });
  if (!adm.ok) return errorResponse(adm.status, adm.code, adm.message, origin, adm.retryAfter);
  try {
    const model = /^jev-[a-z0-9.-]{1,30}$/.test(String(env.JEV_MODEL ?? '').trim()) ? String(env.JEV_MODEL).trim() : DEFAULT_MODEL;
    const base = String(env.AI_GATEWAY_BASE_URL || DEFAULT_BASE).trim().replace(/\/+$/, '');
    let res;
    try {
      res = await fetch(`${base}/decisions/`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ model, state: v.state, questions: v.questions }),
        signal: AbortSignal.timeout(TIMEOUT_MS),
      });
    } catch (e) {
      const timeout = e && (e.name === 'TimeoutError' || e.name === 'AbortError');
      console.warn('[jev] upstream fetch failed', e && e.name);
      return errorResponse(503, 'upstream', timeout ? MSG.slow : MSG.busy, origin, 10);
    }
    const text = await res.text();
    if (!res.ok) return failure(res, text, origin);
    let j;
    try {
      j = JSON.parse(text);
    } catch {
      console.error('[jev] upstream returned non-JSON');
      return errorResponse(503, 'upstream', MSG.busy, origin, 10);
    }
    const answers = flattenAnswers(j.answers);
    if (Object.keys(answers).length === 0) {
      console.error('[jev] upstream returned no answers');
      return errorResponse(503, 'upstream', MSG.busy, origin, 10);
    }
    const usage = isPlain(j.usage) ? { input_tokens: num(j.usage.input_tokens), output_tokens: num(j.usage.output_tokens) } : null;
    return jsonResponse(200, { ok: true, model: typeof j.model === 'string' ? j.model : model, answers, usage }, origin);
  } finally {
    adm.release();
  }
}

function failure(res, text, origin) {
  const s = res.status;
  if (s === 401 || s === 403) {
    console.error('[jev] upstream auth failed');
    return errorResponse(503, 'upstream', MSG.auth, origin, 60);
  }
  if (s === 402) {
    console.error('[jev] upstream credit exhausted');
    return errorResponse(429, 'budget', MSG.credit, origin, 600);
  }
  if (s === 429) {
    const ra = Number(res.headers.get('Retry-After'));
    return errorResponse(429, 'rate', MSG.rate, origin, Number.isFinite(ra) && ra > 0 ? Math.min(60, Math.max(5, ra)) : 20);
  }
  if (s === 404) {
    console.error('[jev] upstream model not available');
    return errorResponse(503, 'upstream', MSG.noModel, origin);
  }
  if (s === 400) {
    return jsonResponse(400, { error: MSG.rejected, code: 'bad_request', detail: upstreamMessage(text) }, origin);
  }
  console.error('[jev] upstream status', s);
  return errorResponse(503, 'upstream', MSG.busy, origin, 10);
}
