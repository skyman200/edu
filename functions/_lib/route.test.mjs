// Plain-node checks for the /api/ai proxy: `node functions/_lib/route.test.mjs`. No test framework.
// The upstream is a fake fetch; nothing here calls the network.
import assert from 'node:assert/strict';
import { pickModel, normalizeEx, systemPromptFor, fallbackWhy, TIERS, isHeavyChat } from './route.js';
import { validateMessages, readLimits, matchCode, clientKey, createStore, badCodeGate, admit, readJsonBody, _resetMemory } from './limits.js';
import { isAllowedOrigin } from './cors.js';
import { onRequest, createSseParser, parseRetryAfter } from '../api/ai.js';

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const user = (content) => ({ role: 'user', content });

test('route table: exercise -> tier, model, maxTokens', () => {
  const table = {
    's1-1': ['fast', 1500], 's2-5': ['fast', 1500],
    's1-2': ['balanced', 4000], 's1-4': ['balanced', 4000], 's2-2': ['balanced', 4000], 's2-3': ['balanced', 4000],
    's2-4': ['balanced', 4000], 's3-1': ['balanced', 4000], 's3-2': ['balanced', 4000], 's3-4': ['balanced', 4000],
    's3-5': ['balanced', 4000], 'p-1': ['balanced', 4000], 'p-2': ['balanced', 4000], 'p-3': ['balanced', 4000],
    'p-4': ['balanced', 4000], 'p-9': ['balanced', 4000],
    's1-3': ['balanced', 6000], 's2-1': ['balanced', 6000], 's3-3': ['balanced', 6000],
    's4-1': ['balanced', 8000], 's4-2': ['balanced', 8000], 's4-3': ['balanced', 8000],
    's4-4': ['strong', 12000], vibe: ['strong', 12000],
  };
  const models = { fast: 'claude-haiku-5-5', balanced: 'claude-sonnet-5-5', strong: 'claude-opus-5-5' };
  for (const [ex, [tier, max]] of Object.entries(table)) {
    const r = pickModel({ ex, messages: [user('hi')] });
    assert.equal(r.tier, tier, ex);
    assert.equal(r.maxTokens, max, ex);
    assert.equal(r.model, models[tier], ex);
    assert.ok(/[가-힣]/.test(r.why) && r.why.includes(r.label), ex);
  }
  assert.equal(pickModel({ ex: 's1-3' }).why, '표와 긴 계획을 정확히 짜야 해서 균형형 Claude Sonnet 5.5를 골랐습니다');
  assert.equal(pickModel({ ex: 's1-1' }).why, '짧은 요약과 질문이라 빠르고 가벼운 Claude Haiku 5.5를 골랐습니다');
  assert.equal(pickModel({ ex: 's4-4' }).why, '코드 한 파일을 끝까지 틀리지 않게 써야 해서 가장 강한 Claude Opus 5.5를 골랐습니다');
  assert.equal(pickModel({ ex: 's4-4' }).label, 'Claude Opus 5.5');
});

test('fallbacks per tier, Claude only', () => {
  assert.deepEqual(pickModel({ ex: 's1-2' }).fallbacks, ['claude-sonnet-5', 'claude-haiku-4-5-20251001']);
  assert.deepEqual(pickModel({ ex: 's1-1' }).fallbacks, ['claude-haiku-4-5-20251001', 'claude-sonnet-5-5']);
  assert.deepEqual(pickModel({ ex: 'vibe' }).fallbacks, ['claude-sonnet-5-5', 'claude-sonnet-5']);
  for (const t of Object.values(TIERS)) for (const m of [t.model, ...t.fallbacks]) assert.ok(m.startsWith('claude-'));
  const r = pickModel({ ex: 's1-2' });
  r.fallbacks.push('x');
  assert.equal(pickModel({ ex: 's1-2' }).fallbacks.length, 2, 'returned arrays are copies');
});

test('unknown or malformed ex -> chat', () => {
  for (const ex of ['s5-1', 's9-9', 'zzz', '', null, 42, 'S1-1 ', 'a'.repeat(17), 's1_1', '../x']) {
    const want = ex === 'S1-1 ' ? 's1-1' : 'chat';
    assert.equal(normalizeEx(ex), want, String(ex));
  }
  assert.equal(pickModel({ ex: 'nope', messages: [user('안녕하세요')] }).ex, 'chat');
});

test('chat classification', () => {
  const tierOf = (text, more = []) => pickModel({ ex: 'chat', messages: [user(text), ...more] }).tier;
  assert.equal(tierOf('플립러닝이 뭔가요?'), 'fast');
  assert.equal(tierOf('학습목표를 한 줄로 다듬어 줘'), 'fast', '목표 is not 표');
  assert.equal(tierOf('발표 수업 팁 세 가지'), 'fast');
  assert.equal(tierOf('표현을 부드럽게 바꿔 줘'), 'fast');
  for (const t of ['이 내용을 표로 정리해 줘', '표 형식으로', '비교표 만들어 줘', '15주 계획서 초안', '슬라이드 문구', 'HTML 페이지 만들어 줘', 'Apps Script 코드', 'make a table', '```js\nx\n```', '엑셀 스크립트']) {
    assert.equal(tierOf(t), 'balanced', t);
  }
  assert.equal(tierOf('가'.repeat(299)), 'fast');
  assert.equal(tierOf('가'.repeat(300)), 'balanced');
  assert.equal(pickModel({ ex: 'chat', messages: [user('가'.repeat(300))] }).maxTokens, 4000);
  assert.equal(pickModel({ ex: 'chat', messages: [user('짧은 질문')] }).maxTokens, 1500);
  // follow-ups classify on the first user message only
  assert.equal(tierOf('짧은 질문', [{ role: 'assistant', content: 'a' }, user('이제 표로 만들어 줘')]), 'fast');
  assert.equal(tierOf('표로 만들어 줘', [{ role: 'assistant', content: 'a' }, user('고마워')]), 'balanced');
  assert.equal(isHeavyChat(undefined), false);
});

test('known exercises keep their model on follow-ups', () => {
  const long = [user('a'), { role: 'assistant', content: 'b' }, user('표로 바꿔 줘 '.repeat(50))];
  assert.equal(pickModel({ ex: 's1-1', messages: long }).model, 'claude-haiku-5-5');
});

test('system prompt', () => {
  const base = systemPromptFor('chat');
  assert.ok(base.length <= 600, `base ${base.length}`);
  assert.ok(base.includes('확인 필요') && base.includes('Markdown 표'));
  for (const ex of ['vibe', 's4-4']) {
    const s = systemPromptFor(ex);
    assert.ok(s.includes('```html') && s.includes('CDN') && s.length <= 700, ex);
  }
  for (const ex of ['s4-1', 's4-2', 's4-3']) assert.ok(systemPromptFor(ex).includes('```javascript'), ex);
  assert.equal(systemPromptFor('s1-1'), base);
});

test('fallback why mentions both models', () => {
  const r = pickModel({ ex: 's4-4' });
  assert.equal(fallbackWhy(r, r.model), r.why);
  assert.ok(fallbackWhy(r, 'claude-sonnet-5-5').startsWith('Claude Opus 5.5가 지금 바빠서 Claude Sonnet 5.5로 대신 답합니다.'));
});

test('validateMessages', () => {
  assert.equal(validateMessages([user('hi')]).ok, true);
  const sys = validateMessages([{ role: 'system', content: 'x' }, user('hi')]);
  assert.deepEqual([sys.status, sys.code], [400, 'bad_request']);
  assert.equal(validateMessages([{ role: 'tool', content: 'x' }]).status, 400);
  assert.equal(validateMessages([{ role: 'user', content: [{ type: 'text', text: 'x' }] }]).status, 400);
  assert.equal(validateMessages([]).status, 400);
  assert.equal(validateMessages('x').status, 400);
  assert.equal(validateMessages([user('a'), { role: 'assistant', content: 'b' }]).status, 400, 'last must be user');
  assert.equal(validateMessages([user('  ')]).status, 400);
  const tooMany = validateMessages(Array.from({ length: 25 }, () => user('x')));
  assert.deepEqual([tooMany.status, tooMany.code], [413, 'too_long']);
  assert.equal(validateMessages(Array.from({ length: 24 }, () => user('x'))).ok, true);
  assert.equal(validateMessages([user('x'.repeat(12001))]).status, 413);
  assert.equal(validateMessages([user('x'.repeat(12000))]).ok, true);
  const total = validateMessages([user('x'.repeat(12000)), user('x'.repeat(12000)), user('x'.repeat(12000)), user('x'.repeat(4001))]);
  assert.equal(total.status, 413);
  const cleaned = validateMessages([{ role: 'assistant', content: '안녕하세요' }, user('a'), { role: 'assistant', content: '' }, user('b')]);
  assert.deepEqual(cleaned.messages, [user('a'), user('b')]);
  const extra = validateMessages([{ role: 'user', content: 'a', name: 'x', tool_calls: [] }]);
  assert.deepEqual(extra.messages, [user('a')], 'only role/content pass through');
});

test('readJsonBody caps size and rejects non-objects', async () => {
  const req = (body, headers = {}) => new Request('http://x/api/ai', { method: 'POST', body, headers });
  assert.equal((await readJsonBody(req('{"a":1}'))).ok, true);
  assert.equal((await readJsonBody(req('[1]'))).status, 400);
  assert.equal((await readJsonBody(req('{bad'))).status, 400);
  assert.equal((await readJsonBody(req('x'.repeat(600 * 1024)))).status, 413);
  assert.equal((await readJsonBody(req('{}', { 'Content-Length': '999999999' }))).status, 413);
});

test('limits: defaults and env overrides', () => {
  assert.deepEqual(readLimits({}), { ip10m: 1500, inflight: 150, badCode10m: 1000, codeDay: 2500, globalDay: 3000, globalMin: 100 });
  const l = readLimits({ AI_LIMIT_IP_10M: '40', AI_LIMIT_CODE_DAY: ' 400 ', AI_LIMIT_GLOBAL_DAY: 'abc', AI_LIMIT_INFLIGHT: '0', AI_LIMIT_GLOBAL_MIN: '5', AI_LIMIT_BADCODE_10M: '15' });
  assert.deepEqual(l, { ip10m: 40, inflight: 150, badCode10m: 15, codeDay: 400, globalDay: 3000, globalMin: 5 });
});

test('access code: trimmed, case-insensitive, list', async () => {
  _resetMemory();
  const raw = ' Alpha , beta2 ,,';
  assert.ok(await matchCode('alpha', raw));
  assert.ok(await matchCode('  ALPHA ', raw));
  assert.ok(await matchCode('BETA2', raw));
  assert.equal(await matchCode('alph', raw), null);
  assert.equal(await matchCode('', raw), null);
  assert.equal(await matchCode('alpha', ''), null);
  assert.notEqual(await matchCode('alpha', raw), await matchCode('beta2', raw));
  assert.ok(!(await matchCode('alpha', raw)).includes('alpha'), 'id is a hash, not the code');
});

test('client key: IPv4 as is, IPv6 /64', () => {
  const k = (ip) => clientKey(new Request('http://x', { headers: ip ? { 'CF-Connecting-IP': ip } : {} }));
  assert.equal(k('203.0.113.7'), '203.0.113.7');
  assert.equal(k('2001:db8:1:2:aaaa::1'), '2001:db8:1:2::/64');
  assert.equal(k('2001:0db8:0001:0002:ffff:ffff:ffff:ffff'), '2001:db8:1:2::/64');
  assert.equal(k('::1'), '0:0:0:0::/64');
  assert.equal(k('::ffff:198.51.100.4'), '198.51.100.4');
  assert.equal(k(''), 'unknown');
});

test('CORS allow-list', () => {
  for (const o of ['https://kimcook-edu.pages.dev', 'https://skyman200.github.io', 'http://localhost:8788', 'http://127.0.0.1:8788', 'https://abc123.kimcook-edu.pages.dev']) {
    assert.ok(isAllowedOrigin(o), o);
  }
  for (const o of ['https://evil.com', 'https://kimcook-edu.pages.dev.evil.com', 'https://a.b.kimcook-edu.pages.dev', 'http://kimcook-edu.pages.dev', 'null', 'https://skyman200.github.io.evil.com', null]) {
    assert.ok(!isAllowedOrigin(o), String(o));
  }
});

test('parseRetryAfter', () => {
  assert.equal(parseRetryAfter('7'), 7);
  assert.equal(parseRetryAfter(null), null);
  assert.equal(parseRetryAfter('soon'), null);
  assert.equal(Math.round(parseRetryAfter(new Date(10_000).toUTCString(), 4_000)), 6);
});

const enc = new TextEncoder();
function feed(parser, parts) {
  let text = '';
  for (const p of parts) text += parser.push(typeof p === 'string' ? enc.encode(p) : p);
  return text + parser.end();
}
const chunk = (content, extra = {}) => `data: ${JSON.stringify({ choices: [{ delta: { content }, finish_reason: null, ...extra }] })}\n\n`;

test('SSE parser: splits, CRLF, null content, DONE, usage, UTF-8 boundaries', () => {
  const finish = `data: ${JSON.stringify({ choices: [{ delta: { content: null }, finish_reason: 'stop' }], usage: { prompt_tokens: 3, completion_tokens: 4, total_tokens: 7, junk: 'x' } })}\r\n\r\n`;
  const full = chunk(null).replace(/\n/g, '\r\n') + chunk('안녕') + ': keepalive\n\n' + chunk('하세요') + finish + 'data: [DONE]\n\n' + chunk('ignored');
  const bytes = enc.encode(full);
  for (const size of [1, 2, 3, 5, 7, 64, bytes.length]) {
    const p = createSseParser();
    const parts = [];
    for (let i = 0; i < bytes.length; i += size) parts.push(bytes.subarray(i, i + size));
    assert.equal(feed(p, parts), '안녕하세요', `size ${size}`);
    assert.equal(p.state.done, true);
    assert.equal(p.state.finishReason, 'stop');
    assert.deepEqual(p.state.usage, { prompt_tokens: 3, completion_tokens: 4, total_tokens: 7 });
  }
  const multi = createSseParser();
  assert.equal(feed(multi, ['data: {"choices":[{"delta":\n', 'data: {"content":"x"}}]}\n\n']), 'x', 'multi-line data');
  const err = createSseParser();
  feed(err, ['data: {"error":{"message":"overloaded"}}\n\n']);
  assert.equal(err.state.error, true);
  const credit = createSseParser();
  feed(credit, ['data: {"error":{"code":6004}}\n\n']);
  assert.equal(credit.state.errorKind, 'credit');
  const len = createSseParser();
  feed(len, [chunk('a', { finish_reason: 'length' })]);
  assert.equal(len.state.finishReason, 'length');
});

test('SSE parser: fast path equals full JSON parse on tricky text', () => {
  const tricky = ['say "error" now', 'a\\b \\"c', 'line1\nline2\r\n', '"finish_reason": null', '"usage": {"x":1}', '\u2028\u0000 끝', '}', '{"delta": {"content": "x"}}', '```html\n<p class="a">x</p>\n```'];
  const gateway = (t) => `data: {"id": "m", "choices": [{"delta": {"content": ${JSON.stringify(t)}, "role": null}, "finish_reason": null, "index": 0}], "usage": null}\n\n`;
  for (const t of tricky) {
    assert.equal(feed(createSseParser(), [gateway(t)]), t, JSON.stringify(t));
    assert.equal(feed(createSseParser(), [chunk(t)]), t, JSON.stringify(t));
  }
  const p = createSseParser();
  assert.equal(feed(p, [gateway('a'), 'data: {"choices": [{"delta": {"content": "b"}, "finish_reason": "stop"}], "usage": {"total_tokens": 5}}\n\n', 'data: [DONE]\n\n']), 'ab');
  assert.deepEqual([p.state.finishReason, p.state.usage], ['stop', { total_tokens: 5 }]);
  const e = createSseParser();
  feed(e, ['data: {"choices": [{"delta": {"content": "x"}, "finish_reason": null}], "error": {"code": 429}}\n\n']);
  assert.deepEqual([e.state.error, e.state.errorKind], [true, 'rate']);
});

test('rate limits in memory (no KV bound)', async () => {
  _resetMemory();
  const store = createStore({});
  const limits = { ...readLimits({}), ip10m: 3, inflight: 2, codeDay: 100, globalDay: 100, globalMin: 100 };
  const now = Date.now();
  const a = await admit(store, { ip: '1.1.1.1', codeId: 'c', now, limits });
  const b = await admit(store, { ip: '1.1.1.1', codeId: 'c', now, limits });
  const c = await admit(store, { ip: '1.1.1.1', codeId: 'c', now, limits });
  assert.ok(a.ok && b.ok);
  assert.deepEqual([c.status, c.code], [429, 'rate'], 'in-flight cap');
  a.release();
  a.release(); // idempotent
  const d = await admit(store, { ip: '1.1.1.1', codeId: 'c', now, limits });
  assert.ok(d.ok);
  d.release();
  const e = await admit(store, { ip: '1.1.1.1', codeId: 'c', now, limits });
  assert.equal(e.status, 429, 'per-IP window cap');
  assert.ok(e.retryAfter > 0 && e.retryAfter <= 600);
  assert.ok((await admit(store, { ip: '2.2.2.2', codeId: 'c', now, limits })).ok, 'other IP unaffected');
  const later = await admit(store, { ip: '1.1.1.1', codeId: 'c', now: now + 11 * 60_000, limits });
  assert.ok(later.ok, 'window slides');
});

test('daily and per-minute budgets', async () => {
  _resetMemory();
  const store = createStore({});
  const base = { ...readLimits({}), ip10m: 1000, inflight: 1000 };
  const now = Date.now();
  const lim = { ...base, codeDay: 2 };
  assert.ok((await admit(store, { ip: 'a', codeId: 'k1', now, limits: lim })).ok);
  assert.ok((await admit(store, { ip: 'b', codeId: 'k1', now, limits: lim })).ok);
  const over = await admit(store, { ip: 'c', codeId: 'k1', now, limits: lim });
  assert.deepEqual([over.status, over.code], [429, 'budget']);
  assert.ok(over.retryAfter > 0 && over.retryAfter <= 86400);
  assert.ok((await admit(store, { ip: 'c', codeId: 'k2', now, limits: lim })).ok, 'another code has its own day cap');
  _resetMemory();
  const g = { ...base, globalDay: 1 };
  assert.ok((await admit(store, { ip: 'a', codeId: 'k1', now, limits: g })).ok);
  assert.equal((await admit(store, { ip: 'b', codeId: 'k2', now, limits: g })).code, 'budget');
  _resetMemory();
  const m = { ...base, globalMin: 2 };
  const t0 = Math.floor(now / 60000) * 60000 + 30_000;
  await admit(store, { ip: 'a', codeId: 'k', now: t0, limits: m });
  await admit(store, { ip: 'b', codeId: 'k', now: t0, limits: m });
  const third = await admit(store, { ip: 'c', codeId: 'k', now: t0, limits: m });
  assert.deepEqual([third.status, third.code, Math.round(third.retryAfter)], [429, 'rate', 30]);
  assert.ok((await admit(store, { ip: 'c', codeId: 'k', now: t0 + 60_000, limits: m })).ok, 'next minute bucket');
});

test('wrong-code lockout blocks the IP, also for a right code', async () => {
  _resetMemory();
  const store = createStore({});
  const now = Date.now();
  for (let i = 0; i < 15; i++) {
    const g = await badCodeGate(store, '9.9.9.9', now, 15);
    assert.equal(g.locked, null, `attempt ${i + 1}`);
    g.recordWrong();
  }
  const g16 = await badCodeGate(store, '9.9.9.9', now, 15);
  assert.deepEqual([g16.locked.status, g16.locked.code], [429, 'rate']);
  assert.equal((await badCodeGate(store, '8.8.8.8', now, 15)).locked, null);
});

test('KV path: counts shared through KV; KV errors fall back to memory', async () => {
  _resetMemory();
  const data = new Map();
  const kv = { get: async (k) => data.get(k) ?? null, put: async (k, v, o) => { assert.ok(o.expirationTtl >= 60); data.set(k, v); } };
  const pending = [];
  const store = createStore({ AI_KV: kv }, (p) => pending.push(p));
  const limits = { ...readLimits({}), ip10m: 2 };
  const now = Date.now();
  const a = await admit(store, { ip: 'k', codeId: 'c', now, limits });
  await Promise.all(pending);
  assert.ok([...data.keys()].some((k) => k.startsWith('r:k:')) && [...data.keys()].some((k) => k.startsWith('f:')));
  a.release();
  await Promise.all(pending);
  _resetMemory(); // a different isolate: memory empty, KV remembers
  const b = await admit(store, { ip: 'k', codeId: 'c', now, limits });
  assert.ok(b.ok);
  await Promise.all(pending);
  _resetMemory();
  const c = await admit(store, { ip: 'k', codeId: 'c', now, limits });
  assert.equal(c.status, 429, 'cap seen through KV after isolate reset');
  _resetMemory();
  for (let i = 0; i < 3; i++) (await badCodeGate(store, 'w', now, 3)).recordWrong();
  await Promise.all(pending);
  _resetMemory();
  assert.equal((await badCodeGate(store, 'w', now, 3)).locked.status, 429, 'lockout seen through KV after isolate reset');
  _resetMemory();
  const broken = { get: async () => { throw new Error('KV down'); }, put: async () => { throw new Error('quota'); } };
  const warn = console.warn;
  console.warn = () => {};
  try {
    const s2 = createStore({ AI_KV: broken }, () => {});
    assert.ok((await admit(s2, { ip: 'z', codeId: 'c', now, limits })).ok);
    assert.ok((await admit(s2, { ip: 'z', codeId: 'c', now, limits })).ok);
    assert.equal((await admit(s2, { ip: 'z', codeId: 'c', now, limits })).status, 429, 'memory still limits');
  } finally {
    console.warn = warn;
  }
});

// ---------- handler with a fake upstream ----------

function sseResponse(parts, status = 200) {
  const stream = new ReadableStream({
    start(c) {
      for (const p of parts) c.enqueue(enc.encode(p));
      c.close();
    },
  });
  return new Response(stream, { status, headers: { 'Content-Type': 'text/event-stream' } });
}
const okStream = (text = 'OK') => sseResponse([chunk(null), chunk(text), chunk(null, { finish_reason: 'stop' }), 'data: [DONE]\n\n']);

async function call(body, { env = {}, origin = 'https://kimcook-edu.pages.dev', ip = '10.0.0.1', method = 'POST', upstream } = {}) {
  const calls = [];
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const req = JSON.parse(init.body);
    calls.push({ url, model: req.model, maxTokens: req.max_tokens, messages: req.messages, auth: init.headers.Authorization });
    return upstream ? upstream(req, calls.length) : okStream();
  };
  try {
    const headers = { 'Content-Type': 'application/json', 'CF-Connecting-IP': ip };
    if (origin) headers.Origin = origin;
    const request = new Request('http://local/api/ai', { method, headers, body: method === 'POST' ? JSON.stringify(body) : undefined });
    const res = await onRequest({ request, env: { AI_GATEWAY_API_KEY: 'sk-test-secret', AI_CLASS_CODE: 'TESTCODE', ...env }, waitUntil() {} });
    const text = await res.text();
    return { res, text, calls, events: text.split('\n\n').filter(Boolean).map((l) => JSON.parse(l.replace(/^data: /, ''))) };
  } finally {
    globalThis.fetch = realFetch;
  }
}
const parseJson = (r) => JSON.parse(r.text);

test('handler: wrong code -> 401, no upstream call, no secret in body', async () => {
  _resetMemory();
  const r = await call({ code: 'nope', ex: 's1-1', messages: [user('hi')] });
  assert.equal(r.res.status, 401);
  assert.equal(parseJson(r).code, 'bad_code');
  assert.equal(r.calls.length, 0);
  assert.equal(r.res.headers.get('Cache-Control'), 'no-store');
  assert.equal(r.res.headers.get('X-Robots-Tag'), 'noindex');
  assert.equal(r.res.headers.get('Access-Control-Allow-Origin'), 'https://kimcook-edu.pages.dev');
});

test('handler: check:true validates the code only', async () => {
  _resetMemory();
  const ok = await call({ code: ' testcode ', check: true });
  assert.deepEqual([ok.res.status, parseJson(ok)], [200, { ok: true }]);
  assert.equal(ok.calls.length, 0);
  const bad = await call({ code: 'x', check: true });
  assert.equal(bad.res.status, 401);
});

test('handler: stream contract, server-chosen model, client params ignored', async () => {
  _resetMemory();
  const r = await call({ code: 'TESTCODE', ex: 's1-1', model: 'gpt-x', max_tokens: 99999, system: 'evil', messages: [user('hi')] });
  assert.equal(r.res.status, 200);
  assert.equal(r.res.headers.get('Content-Type'), 'text/event-stream; charset=utf-8');
  assert.equal(r.calls.length, 1);
  assert.deepEqual([r.calls[0].model, r.calls[0].maxTokens], ['claude-haiku-5-5', 1500]);
  assert.equal(r.calls[0].messages[0].role, 'system');
  assert.equal(r.calls[0].messages.filter((m) => m.role === 'system').length, 1);
  assert.equal(r.calls[0].auth, 'Bearer sk-test-secret');
  assert.ok(r.calls[0].url.endsWith('/v1/gateway/chat/completions'));
  assert.deepEqual(r.events[0].meta, { model: 'claude-haiku-5-5', label: 'Claude Haiku 5.5', tier: 'fast', why: pickModel({ ex: 's1-1' }).why });
  assert.deepEqual(r.events[1], { delta: 'OK' });
  assert.deepEqual(r.events.at(-1), { done: true });
  assert.ok(!r.text.includes('sk-test-secret'));
});

test('handler: 5xx falls back, meta shows the model used', async () => {
  _resetMemory();
  const warn = console.warn;
  console.warn = () => {};
  try {
    const r = await call({ code: 'TESTCODE', ex: 's4-4', messages: [user('Say OK')] }, { upstream: (req, n) => (n === 1 ? new Response('{"detail":"boom https://internal"}', { status: 500 }) : okStream()) });
    assert.deepEqual(r.calls.map((c) => c.model), ['claude-opus-5-5', 'claude-sonnet-5-5']);
    assert.equal(r.calls[1].maxTokens, 12000);
    assert.equal(r.events[0].meta.model, 'claude-sonnet-5-5');
    assert.ok(r.events[0].meta.why.startsWith('Claude Opus 5.5가 지금 바빠서'));
    const all = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: () => new Response('{"detail":"secret-req-id"}', { status: 503 }) });
    assert.equal(all.res.status, 503);
    assert.equal(parseJson(all).code, 'upstream');
    assert.equal(all.calls.length, 3);
    assert.ok(!all.text.includes('secret-req-id'));
    const model404 = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: (req, n) => (n === 1 ? new Response('{"detail":{"message":"Model not found"}}', { status: 400 }) : okStream()) });
    assert.equal(model404.events[0].meta.model, 'claude-sonnet-5');
    const plain400 = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: () => new Response('{"detail":"bad"}', { status: 400 }) });
    assert.deepEqual([plain400.res.status, plain400.calls.length], [503, 1]);
  } finally {
    console.warn = warn;
  }
});

test('handler: upstream 429 retries the same model, never falls back', async () => {
  _resetMemory();
  const warn = console.warn;
  console.warn = () => {};
  try {
    const t0 = Date.now();
    const r = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: (req, n) => (n < 3 ? new Response('', { status: 429, headers: { 'Retry-After': '0' } }) : okStream()) });
    assert.deepEqual(r.calls.map((c) => c.model), ['claude-sonnet-5-5', 'claude-sonnet-5-5', 'claude-sonnet-5-5']);
    assert.equal(r.res.status, 200);
    assert.ok(Date.now() - t0 >= 1900, 'waited at least 1 s per retry');
    const still = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: () => new Response('', { status: 429, headers: { 'Retry-After': '1' } }) });
    assert.deepEqual([still.res.status, parseJson(still).code, still.calls.length], [429, 'rate', 3]);
    assert.ok(Number(still.res.headers.get('Retry-After')) >= 10);
  } finally {
    console.warn = warn;
  }
});

test('handler: 402 credit -> budget, 401 -> upstream, no fallback, generic logs', async () => {
  _resetMemory();
  const logs = [];
  const err = console.error;
  console.error = (...a) => logs.push(a.join(' '));
  try {
    const r = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: () => new Response('{"detail":{"code":6004}}', { status: 402 }) });
    assert.deepEqual([r.res.status, parseJson(r).code, parseJson(r).error, r.calls.length], [429, 'budget', 'AI ONE 크레딧이 부족합니다. 강사에게 알려 주세요.', 1]);
    const a = await call({ code: 'TESTCODE', ex: 's1-2', messages: [user('x')] }, { upstream: () => new Response('{"detail":"invalid key sk-..."}', { status: 401 }) });
    assert.deepEqual([a.res.status, parseJson(a).code, a.calls.length], [503, 'upstream', 1]);
    assert.deepEqual(logs, ['[ai] upstream credit exhausted', '[ai] upstream auth failed']);
  } finally {
    console.error = err;
  }
});

test('handler: truncated, mid-stream failure, early stream error falls back', async () => {
  _resetMemory();
  const t = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, { upstream: () => sseResponse([chunk('a'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n']) });
  assert.deepEqual(t.events.at(-1), { done: true, truncated: true });
  const cut = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, { upstream: () => sseResponse([chunk('part')]) });
  assert.deepEqual(cut.events.slice(1).map((e) => Object.keys(e)[0]), ['delta', 'error']);
  const warn = console.warn;
  console.warn = () => {};
  try {
    const early = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, { upstream: (req, n) => (n === 1 ? sseResponse(['data: {"error":{"message":"overloaded"}}\n\n']) : okStream('fine')) });
    assert.equal(early.events[0].meta.model, 'claude-haiku-4-5-20251001');
    assert.deepEqual(early.events[1], { delta: 'fine' });
  } finally {
    console.warn = warn;
  }
});

test('handler: validation, origin, method, disabled', async () => {
  _resetMemory();
  const sys = await call({ code: 'TESTCODE', ex: 's1-1', messages: [{ role: 'system', content: 'x' }, user('hi')] });
  assert.deepEqual([sys.res.status, sys.calls.length], [400, 0]);
  const big = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x'.repeat(12001))] });
  assert.deepEqual([big.res.status, parseJson(big).code], [413, 'too_long']);
  const evil = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('hi')] }, { origin: 'https://evil.example' });
  assert.equal(evil.res.status, 403);
  assert.equal(evil.res.headers.get('Access-Control-Allow-Origin'), null);
  assert.equal(evil.res.headers.get('Vary'), 'Origin');
  const get = await call(null, { method: 'GET' });
  assert.equal(get.res.status, 405);
  const pre = await call(null, { method: 'OPTIONS', origin: 'https://skyman200.github.io' });
  assert.equal(pre.res.status, 204);
  assert.equal(pre.res.headers.get('Access-Control-Allow-Origin'), 'https://skyman200.github.io');
  assert.equal((await call(null, { method: 'OPTIONS', origin: 'https://evil.example' })).res.status, 403);
  const off = await call({ code: 'TESTCODE', check: true }, { env: { AI_ENABLED: 'false' } });
  assert.deepEqual([off.res.status, parseJson(off).code], [503, 'disabled']);
  const noCodes = await call({ code: 'TESTCODE', check: true }, { env: { AI_CLASS_CODE: ' , ' } });
  assert.equal(noCodes.res.status, 503);
  const err = console.error;
  console.error = () => {};
  try {
    const noKey = await call({ code: 'TESTCODE', check: true }, { env: { AI_GATEWAY_API_KEY: '' } });
    assert.equal(noKey.res.status, 503);
  } finally {
    console.error = err;
  }
});

let failed = 0;
for (const [name, fn] of tests) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (e) {
    failed++;
    console.log(`FAIL ${name}\n     ${e && e.stack ? e.stack.split('\n').slice(0, 3).join('\n     ') : e}`);
  }
}
console.log(failed ? `${failed} of ${tests.length} failed` : `all ${tests.length} passed`);
process.exitCode = failed ? 1 : 0;
