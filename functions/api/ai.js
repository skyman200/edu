// POST /api/ai: runs a workshop exercise prompt through the DIT AI-ONE (Mindlogic) gateway.
// The browser never sees the API key; the server picks the model (../_lib/route.js).
// Request:  { code, ex, messages:[{role:'user'|'assistant', content}][, files:[{name,text}]][, images:[{name,url}]] }
//           or { code, check:true }. files carry text extracted in the browser; images are data:image/...;base64 URLs.
//           Both are resent on every turn and go into the first user message (../_lib/route.js withAttachments).
// Success:  text/event-stream of  data:{"meta":{model,label,tier,why}}, data:{"delta":"..."}...,
//           then data:{"done":true[,"truncated":true][,"usage":{...}]}  or  data:{"error":"..."}
// Failure:  JSON { error:<Korean>, code:bad_code|rate|budget|disabled|too_long|bad_request|upstream[, retryAfter] }
import { pickModel, systemPromptFor, withAttachments, labelOf, fallbackWhy, TIERS, MODEL_LABELS } from '../_lib/route.js';
import { normalizeView, applyView } from '../_lib/views.js';
import {
  SIZE,
  readLimits,
  isKillSwitchOff,
  readJsonBody,
  validateMessages,
  validateAttachments,
  parseCodes,
  matchCode,
  clientKey,
  createStore,
  badCodeGate,
  admit,
} from '../_lib/limits.js';
import { isAllowedOrigin, jsonResponse, errorResponse, preflightResponse, baseHeaders } from '../_lib/cors.js';

const DEFAULT_BASE = 'https://factchat-cloud.mindlogic.ai/v1/gateway';
const FIRST_CONTENT_MS = 100_000; // until the first answer text, across fallbacks and 429 waits
const IDLE_MS = 60_000; // between upstream chunks once the answer has started
const HARD_MS = 300_000; // whole request
const RATE_RETRIES = 2; // extra tries of the same model after an upstream 429

const MSG = {
  disabled: 'AI 기능이 지금 꺼져 있습니다. 강사에게 문의해 주세요.',
  method: '허용되지 않은 요청 방식입니다.',
  origin: '허용되지 않은 사이트에서 보낸 요청입니다.',
  noCode: '수업 코드를 입력해 주세요.',
  longCode: '수업 코드가 너무 깁니다. 강사가 알려 준 코드를 확인해 주세요.',
  badCode: '수업 코드가 맞지 않습니다. 강사가 알려 준 코드를 확인해 주세요.',
  busy: 'AI ONE 서버가 바쁩니다. 잠시 후 다시 시도해 주세요.',
  slow: 'AI ONE 서버 응답이 너무 늦습니다. 잠시 후 다시 시도해 주세요.',
  auth: 'AI ONE 연결 설정에 문제가 있습니다. 강사에게 알려 주세요.',
  credit: 'AI ONE 크레딧이 부족합니다. 강사에게 알려 주세요.',
  upstreamRate: 'AI ONE 서버에 요청이 몰렸습니다. 1분쯤 뒤에 다시 시도해 주세요.',
  rejected: 'AI ONE 서버가 요청을 처리하지 못했습니다. 내용을 조금 바꿔 다시 시도해 주세요.',
  internal: 'AI 요청을 처리하지 못했습니다. 잠시 후 다시 시도해 주세요.',
  midCut: 'AI 응답이 중간에 끊겼습니다. 받은 부분까지만 표시됩니다. 다시 시도해 주세요.',
  midIdle: 'AI 응답이 1분 넘게 멈춰서 중단했습니다. 받은 부분까지만 표시됩니다. 다시 시도해 주세요.',
  midHard: 'AI 응답이 5분을 넘겨 중단했습니다. 받은 부분까지만 표시됩니다. 요청을 나눠서 다시 시도해 주세요.',
};

export async function onRequest(context) {
  const origin = context.request.headers.get('Origin');
  try {
    return await handle(context, origin);
  } catch (e) {
    console.error('[ai] internal error', e && e.name);
    return errorResponse(503, 'upstream', MSG.internal, origin);
  }
}

function fromFailure(f, origin) {
  return errorResponse(f.status, f.code, f.message, origin, f.retryAfter);
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
  if (isKillSwitchOff(env) || parseCodes(codesRaw).length === 0) {
    return errorResponse(503, 'disabled', MSG.disabled, origin);
  }
  const apiKey = String(env.AI_GATEWAY_API_KEY ?? '').trim();
  if (!apiKey) {
    console.error('[ai] AI_GATEWAY_API_KEY is not set');
    return errorResponse(503, 'disabled', MSG.disabled, origin);
  }

  const parsed = await readJsonBody(request);
  if (!parsed.ok) return fromFailure(parsed, origin);
  const body = parsed.body;
  if (typeof body.code !== 'string' || !body.code.trim()) return errorResponse(401, 'bad_code', MSG.noCode, origin);
  if (body.code.length > SIZE.maxCode) return errorResponse(400, 'bad_request', MSG.longCode, origin);

  // Access code before any budget: a wrong code never reaches the gateway.
  const limits = readLimits(env);
  const store = createStore(env, waitUntil);
  const ip = clientKey(request);
  const gate = await badCodeGate(store, ip, Date.now(), limits.badCode10m);
  if (gate.locked) return fromFailure(gate.locked, origin);
  const codeId = await matchCode(body.code, codesRaw);
  if (!codeId) {
    gate.recordWrong();
    return errorResponse(401, 'bad_code', MSG.badCode, origin);
  }
  if (body.check === true) return jsonResponse(200, { ok: true }, origin);

  // Only roles and text pass through; model, system, max_tokens, tools... from the client are ignored.
  const v = validateMessages(body.messages);
  if (!v.ok) return fromFailure(v, origin);
  const att = validateAttachments(body.files, body.images, body.messages);
  if (!att.ok) return fromFailure(att, origin);
  const hasAttachments = att.files.length + att.images.length > 0;
  const view = normalizeView(body.view);
  const route = applyView(pickModel({ ex: body.ex, messages: v.messages, hasAttachments }), view, TIERS, MODEL_LABELS);

  const adm = await admit(store, { ip, codeId, now: Date.now(), limits });
  if (!adm.ok) return fromFailure(adm, origin);
  // Without the enable_request_signal flag request.signal never fires; with it, a closed tab frees the in-flight slot at once (release is idempotent)
  if (request.signal) request.signal.addEventListener('abort', adm.release, { once: true });

  const ctl = createControl(request);
  const upstream = {
    url: `${String(env.AI_GATEWAY_BASE_URL || DEFAULT_BASE).trim().replace(/\/+$/, '')}/chat/completions`,
    apiKey,
    // Serialized once: with attachments it can be megabytes, and connect() may send it up to 9 times.
    messagesJson: JSON.stringify([{ role: 'system', content: systemPromptFor(route.ex, hasAttachments, view) }, ...withAttachments(v.messages, att)]),
  };
  let conn;
  try {
    conn = await connect(route, upstream, ctl);
  } catch (e) {
    console.error('[ai] upstream connect error', e && e.name);
    conn = { ok: false, kind: 'busy' };
  }
  if (!conn.ok) {
    ctl.clear();
    adm.release();
    return failureResponse(conn, origin);
  }
  return streamResponse(conn, route, ctl, adm.release, origin);
}

function failureResponse(conn, origin) {
  switch (conn.kind) {
    case 'auth':
      console.error('[ai] upstream auth failed');
      return errorResponse(503, 'upstream', MSG.auth, origin, 60);
    case 'credit':
      console.error('[ai] upstream credit exhausted');
      return errorResponse(429, 'budget', MSG.credit, origin, 600);
    case 'rate':
      console.warn('[ai] upstream rate limited');
      return errorResponse(429, 'rate', MSG.upstreamRate, origin, Math.min(60, Math.max(10, conn.retryAfter || 30)));
    case 'timeout':
      return errorResponse(503, 'upstream', MSG.slow, origin, 10);
    case 'rejected':
      return errorResponse(503, 'upstream', MSG.rejected, origin);
    default:
      return errorResponse(503, 'upstream', MSG.busy, origin, 10);
  }
}

// ---------- abort / timeouts ----------

function createControl(request) {
  const ac = new AbortController();
  const timers = {};
  const ctl = {
    signal: ac.signal,
    reason: null,
    abort(reason) {
      if (!ctl.reason) ctl.reason = reason;
      ac.abort();
    },
    contentStarted() {
      clearTimeout(timers.first);
    },
    armIdle() {
      clearTimeout(timers.idle);
      timers.idle = setTimeout(() => ctl.abort('idle'), IDLE_MS);
    },
    disarmIdle() {
      clearTimeout(timers.idle);
    },
    clear() {
      clearTimeout(timers.first);
      clearTimeout(timers.idle);
      clearTimeout(timers.hard);
    },
  };
  timers.first = setTimeout(() => ctl.abort('first'), FIRST_CONTENT_MS);
  timers.hard = setTimeout(() => ctl.abort('hard'), HARD_MS);
  const rs = request.signal;
  if (rs && typeof rs.addEventListener === 'function') {
    if (rs.aborted) ctl.abort('client');
    else rs.addEventListener('abort', () => ctl.abort('client'), { once: true });
  }
  return ctl;
}

function sleep(ms, signal) {
  return new Promise((resolve) => {
    if (signal.aborted) return resolve(false);
    const onAbort = () => {
      clearTimeout(t);
      resolve(false);
    };
    const t = setTimeout(() => {
      signal.removeEventListener('abort', onAbort);
      resolve(true);
    }, ms);
    signal.addEventListener('abort', onAbort, { once: true });
  });
}

export function parseRetryAfter(value, now = Date.now()) {
  if (value == null || String(value).trim() === '') return null;
  const n = Number(value);
  if (Number.isFinite(n)) return Math.max(0, n);
  const t = Date.parse(value);
  return Number.isFinite(t) ? Math.max(0, (t - now) / 1000) : null;
}

// ---------- upstream ----------

function discard(res) {
  if (res.body) res.body.cancel().catch(() => {});
}

async function peekText(res, maxBytes) {
  if (!res.body) return '';
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let s = '';
  try {
    while (s.length < maxBytes) {
      const { value, done } = await reader.read();
      if (done) break;
      s += dec.decode(value, { stream: true });
    }
  } catch {
    // ignore: only used to classify the error
  }
  reader.cancel().catch(() => {});
  return s;
}

// One HTTP attempt. Reads the stream until the first answer text so a failure that happens
// before any content can still fall back to another model.
async function tryModel(model, route, upstream, ctl) {
  let res;
  try {
    res = await fetch(upstream.url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${upstream.apiKey}`,
        'Content-Type': 'application/json',
        Accept: 'text/event-stream',
      },
      body: `{"model":${JSON.stringify(model)},"messages":${upstream.messagesJson},"max_tokens":${route.maxTokens},"stream":true}`,
      signal: ctl.signal,
    });
  } catch {
    return ctl.signal.aborted ? { kind: 'aborted' } : { kind: 'network', retry: true };
  }
  const st = res.status;
  if (st === 401 || st === 403) {
    discard(res);
    return { kind: 'auth' };
  }
  if (st === 402) {
    discard(res);
    return { kind: 'credit' };
  }
  if (st === 429) {
    discard(res);
    return { kind: 'rate', retryAfter: parseRetryAfter(res.headers.get('Retry-After')) };
  }
  if (!res.ok || !res.body) {
    let retry = st === 404 || st >= 500;
    if (!retry && (st === 400 || st === 422)) retry = /model/i.test(await peekText(res, 4096));
    else discard(res);
    return { kind: 'status', status: st, retry };
  }

  const reader = res.body.getReader();
  const parser = createSseParser();
  try {
    for (;;) {
      const { value, done } = await reader.read();
      const text = done ? parser.end() : parser.push(value);
      const s = parser.state;
      if (s.error && !text) {
        reader.cancel().catch(() => {});
        if (s.errorKind === 'rate') return { kind: 'rate', retryAfter: null };
        if (s.errorKind === 'credit') return { kind: 'credit' };
        return { kind: 'stream-error', retry: true };
      }
      if (text || s.done || s.finishReason) return { ok: true, model, reader, parser, first: text, ended: done };
      if (done) return { kind: 'empty', retry: true };
    }
  } catch {
    reader.cancel().catch(() => {});
    return ctl.signal.aborted ? { kind: 'aborted' } : { kind: 'network', retry: true };
  }
}

// Primary model, then fallbacks on 5xx/404/model errors/early stream failures.
// 429 retries the same model (the gateway limit is per member, so another model would not help);
// 401/403 and 402 stop at once.
async function connect(route, upstream, ctl) {
  const candidates = [route.model, ...route.fallbacks];
  for (let i = 0; i < candidates.length; i++) {
    const model = candidates[i];
    for (let rateTry = 0; ; rateTry++) {
      const r = await tryModel(model, route, upstream, ctl);
      if (r.ok) return r;
      if (r.kind === 'aborted') return { ok: false, kind: ctl.reason === 'first' ? 'timeout' : 'aborted' };
      if (r.kind === 'auth' || r.kind === 'credit') return { ok: false, kind: r.kind };
      if (r.kind === 'rate') {
        if (rateTry >= RATE_RETRIES) return { ok: false, kind: 'rate', retryAfter: r.retryAfter };
        const waitSec = Math.min(10, Math.max(1, r.retryAfter ?? 3));
        if (!(await sleep(waitSec * 1000, ctl.signal))) {
          return { ok: false, kind: ctl.reason === 'first' ? 'timeout' : 'aborted' };
        }
        continue;
      }
      if (!r.retry) return { ok: false, kind: 'rejected' };
      if (i + 1 < candidates.length) console.warn(`[ai] upstream ${r.status || r.kind} on ${model}; trying ${candidates[i + 1]}`);
      break;
    }
  }
  return { ok: false, kind: 'busy' };
}

// ---------- SSE parsing ----------

function classifyError(err) {
  let s = '';
  try {
    s = JSON.stringify(err).slice(0, 500);
  } catch {
    // ignore
  }
  if (/\b(402|6004|6010)\b|credit|insufficient/i.test(s)) return 'credit';
  if (/\b429\b|rate.?limit/i.test(s)) return 'rate';
  return 'other';
}

// Plain text chunks skip the full JSON.parse: on the Workers free plan (10 ms CPU per request)
// that halves the cost of a 12k-token answer. Anything else (finish, usage, error) is fully parsed.
const FAST_CONTENT = /"delta":\s*\{"content":\s*("(?:[^"\\]|\\.)*"|null)/;
const FINISH_NULL = /"finish_reason":\s*null/;
const USAGE_OBJECT = /"usage":\s*\{/;

function fastText(payload) {
  if (payload.charCodeAt(payload.length - 1) !== 125 || payload.includes('"error"')) return undefined;
  if (!FINISH_NULL.test(payload) || USAGE_OBJECT.test(payload)) return undefined;
  const m = FAST_CONTENT.exec(payload);
  if (!m) return undefined;
  if (m[1] === 'null') return '';
  try {
    return JSON.parse(m[1]);
  } catch {
    return undefined;
  }
}

// Incremental parser for OpenAI-style chat.completion.chunk SSE. push() returns the answer text
// found in that chunk; state holds finishReason, usage, done ([DONE] seen) and error.
export function createSseParser() {
  const decoder = new TextDecoder();
  const state = { finishReason: null, usage: null, done: false, error: false, errorKind: null };
  let buf = '';
  let pending = '';
  let eventName = '';

  function handle(obj, out) {
    if (eventName === 'error' || (obj && obj.error)) {
      state.error = true;
      state.errorKind = classifyError(obj && obj.error ? obj.error : obj);
      return;
    }
    const c = obj && Array.isArray(obj.choices) ? obj.choices[0] : null;
    if (c) {
      const t = c.delta && c.delta.content;
      if (typeof t === 'string') out.text += t;
      if (c.finish_reason) state.finishReason = String(c.finish_reason);
    }
    const u = obj && obj.usage;
    if (u && typeof u === 'object') {
      const usage = {};
      for (const k of ['prompt_tokens', 'completion_tokens', 'total_tokens']) {
        if (Number.isFinite(u[k])) usage[k] = u[k];
      }
      if (Object.keys(usage).length) state.usage = usage;
    }
  }

  // Returns true when the payload was consumed.
  function tryPayload(payload, force, out) {
    if (payload.trim() === '[DONE]') {
      state.done = true;
      return true;
    }
    if (!eventName) {
      const t = fastText(payload);
      if (t !== undefined) {
        out.text += t;
        return true;
      }
    }
    let obj;
    try {
      obj = JSON.parse(payload);
    } catch {
      return force;
    }
    handle(obj, out);
    return true;
  }

  function line(l, out) {
    if (state.done) return;
    if (l === '') {
      if (pending) tryPayload(pending, true, out);
      pending = '';
      eventName = '';
      return;
    }
    if (l[0] === ':') return;
    const i = l.indexOf(':');
    const field = i < 0 ? l : l.slice(0, i);
    let value = i < 0 ? '' : l.slice(i + 1);
    if (value[0] === ' ') value = value.slice(1);
    if (field === 'event') {
      eventName = value;
      return;
    }
    if (field !== 'data') return;
    if (pending) {
      pending += `\n${value}`;
      if (tryPayload(pending, false, out)) pending = '';
    } else if (!tryPayload(value, false, out)) {
      pending = value;
    }
  }

  function drain(final) {
    const out = { text: '' };
    let start = 0;
    if (buf.indexOf('\r') === -1) {
      for (let nl = buf.indexOf('\n'); nl !== -1; nl = buf.indexOf('\n', start)) {
        line(buf.slice(start, nl), out);
        start = nl + 1;
      }
    } else {
      const re = /\r\n|\r|\n/g;
      let m;
      while ((m = re.exec(buf))) {
        if (!final && m[0] === '\r' && m.index === buf.length - 1) break; // maybe half of \r\n
        line(buf.slice(start, m.index), out);
        start = re.lastIndex;
      }
    }
    buf = buf.slice(start);
    if (final) {
      if (buf) line(buf, out);
      buf = '';
      line('', out);
    }
    return out.text;
  }

  return {
    state,
    push(chunk) {
      buf += decoder.decode(chunk, { stream: true });
      return drain(false);
    },
    end() {
      buf += decoder.decode();
      return drain(true);
    },
  };
}

// ---------- response stream ----------

const encoder = new TextEncoder();
function sse(obj) {
  return encoder.encode(`data: ${JSON.stringify(obj)}\n\n`);
}

function streamResponse(conn, route, ctl, release, origin) {
  const { reader, parser } = conn;
  ctl.contentStarted();
  const meta = { model: conn.model, label: labelOf(conn.model), tier: route.tier, why: fallbackWhy(route, conn.model) };
  let readerDone = conn.ended;
  let finished = false;

  const cleanup = () => {
    if (finished) return;
    finished = true;
    ctl.clear();
    if (!readerDone) reader.cancel().catch(() => {});
    release();
  };
  const endWith = (controller, obj) => {
    controller.enqueue(sse(obj));
    controller.close();
    cleanup();
  };
  const doneEvent = () => {
    const ev = { done: true };
    const fr = parser.state.finishReason;
    if (fr === 'length' || fr === 'max_tokens') ev.truncated = true;
    if (parser.state.usage) ev.usage = parser.state.usage;
    return ev;
  };
  // Ends the stream when the upstream answer is complete or broken. Returns true when ended.
  const settle = (controller) => {
    const s = parser.state;
    if (s.error) endWith(controller, { error: MSG.midCut });
    else if (s.done) endWith(controller, doneEvent());
    else if (readerDone) endWith(controller, s.finishReason ? doneEvent() : { error: MSG.midCut });
    else return false;
    return true;
  };

  const stream = new ReadableStream({
    start(controller) {
      controller.enqueue(sse({ meta }));
      if (conn.first) controller.enqueue(sse({ delta: conn.first }));
      settle(controller);
    },
    async pull(controller) {
      if (finished) return;
      try {
        for (;;) {
          ctl.armIdle();
          const { value, done } = await reader.read();
          ctl.disarmIdle();
          if (done) readerDone = true;
          const text = done ? parser.end() : parser.push(value);
          if (text) controller.enqueue(sse({ delta: text }));
          if (settle(controller) || text) return;
        }
      } catch {
        if (finished) return;
        readerDone = true;
        if (ctl.reason === 'client') {
          cleanup();
          return;
        }
        if (parser.state.finishReason) {  // the answer was complete; only the trailing [DONE] was lost
          try {
            endWith(controller, doneEvent());
          } catch {
            cleanup();
          }
          return;
        }
        try {
          endWith(controller, { error: ctl.reason === 'idle' ? MSG.midIdle : ctl.reason === 'hard' ? MSG.midHard : MSG.midCut });
        } catch {
          cleanup();
        }
      }
    },
    cancel() {
      ctl.abort('client');
      cleanup();
    },
  });

  const headers = baseHeaders(origin, { 'Content-Type': 'text/event-stream; charset=utf-8', 'X-Accel-Buffering': 'no' });
  return new Response(stream, { status: 200, headers });
}
