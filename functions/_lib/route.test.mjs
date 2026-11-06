// Plain-node checks for the /api/ai proxy: `node functions/_lib/route.test.mjs`. No test framework.
// The upstream is a fake fetch; nothing here calls the network.
import assert from 'node:assert/strict';
import { pickModel, normalizeEx, systemPromptFor, withAttachments, fallbackWhy, TIERS, isHeavyChat } from './route.js';
import { SIZE, validateMessages, validateAttachments, readLimits, matchCode, clientKey, createStore, badCodeGate, admit, readJsonBody, _resetMemory } from './limits.js';
import { isAllowedOrigin } from './cors.js';
import { normalizeView, applyView, viewPrompt, VIEWS } from './views.js';
import { onRequest, createSseParser, parseRetryAfter } from '../api/ai.js';

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const user = (content) => ({ role: 'user', content });

test('route table: exercise -> tier, model, maxTokens', () => {
  const table = {
    's1-1': ['fast', 1500], 's2-5': ['fast', 1500],
    's1-2': ['balanced', 4000], 's1-4': ['balanced', 4000], 's2-2': ['balanced', 4000], 's2-3': ['balanced', 4000],
    's2-4': ['balanced', 4000], 's3-1': ['balanced', 4000], 's3-2': ['balanced', 4000], 's3-4': ['balanced', 4000],
    's3-5': ['balanced', 4000], 'p-1': ['balanced', 6000], 'p-2': ['balanced', 6000], 'p-3': ['balanced', 6000],
    'p-4': ['balanced', 6000], 'p-9': ['balanced', 6000],
    's1-3': ['balanced', 6000], 's2-1': ['balanced', 6000], 's3-3': ['balanced', 6000],
    's4-1': ['balanced', 8000], 's4-2': ['balanced', 8000], 's4-3': ['balanced', 8000],
    's4-4': ['strong', 12000], vibe: ['strong', 12000],
    'm1-1': ['balanced', 6000], 'm1-2': ['balanced', 4000], 'm1-3': ['balanced', 8000], 'lv-group': ['balanced', 3000],
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
  assert.equal(SIZE.maxBodyBytes, 6 * 1024 * 1024);
  assert.equal((await readJsonBody(req('x'.repeat(SIZE.maxBodyBytes + 1)))).status, 413);
  assert.equal((await readJsonBody(req(JSON.stringify({ t: 'x'.repeat(SIZE.maxBodyBytes - 20) })))).ok, true);
  assert.equal((await readJsonBody(req('{}', { 'Content-Length': '999999999' }))).status, 413);
  // no Content-Length and an endless body: reading stops at the byte cap
  let pulled = 0;
  const endless = new ReadableStream({ pull(c) { pulled += 65536; c.enqueue(new Uint8Array(65536)); } });
  const cut = await readJsonBody(new Request('http://x/api/ai', { method: 'POST', body: endless, duplex: 'half' }));
  assert.equal(cut.status, 413);
  assert.ok(pulled <= SIZE.maxBodyBytes + 4 * 65536, `pulled ${pulled}`);
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
    calls.push({ url, model: req.model, maxTokens: req.max_tokens, messages: req.messages, auth: init.headers.Authorization, body: req });
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
  assert.deepEqual(t.events.at(-1), { done: true, truncated: true, continued: 3 }, 'a model that never finishes is continued 3 times, then reported as truncated');
  assert.equal(t.calls.length, 4);
  assert.equal(deltas(t), 'aaaa', 'text held at a seam is not lost when a continuation is itself cut off');
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

const usageChunk = (p, c) => `data: ${JSON.stringify({ choices: [], usage: { prompt_tokens: p, completion_tokens: c, total_tokens: p + c } })}\n\n`;
const deltas = (r) => r.events.filter((e) => 'delta' in e).map((e) => e.delta).join('');

test('validateMessages: long earlier answers are accepted, user turns stay capped', () => {
  const a = (n) => ({ role: 'assistant', content: 'a'.repeat(n) });
  assert.equal(validateMessages([user('q'), a(60000), user('more')]).ok, true);
  assert.equal(validateMessages([user('q'), a(200000), user('more')]).ok, true);
  assert.equal(validateMessages([user('q'), a(200001), user('more')]).status, 413);
  assert.equal(validateMessages([user('q'), a(120000), user('x'), a(80000), user('y'), a(1), user('z')]).status, 413);
  assert.equal(validateMessages([user('q'), a(5000), user('x'.repeat(12001))]).status, 413);
});

test('handler: a cut-off answer is continued inside the same stream', async () => {
  _resetMemory();
  const r = await call(
    { code: 'TESTCODE', ex: 's1-1', messages: [user('write a lot')], files: [{ name: 'a.txt', text: 'FILEBODY' }] },
    {
      upstream: (req, n) =>
        n === 1
          ? sseResponse([chunk('Hello, '), chunk(null, { finish_reason: 'length' }), usageChunk(100, 10), 'data: [DONE]\n\n'])
          : sseResponse([chunk('world.'), chunk(null, { finish_reason: 'stop' }), usageChunk(120, 5), 'data: [DONE]\n\n']),
    },
  );
  assert.equal(r.calls.length, 2);
  assert.equal(deltas(r), 'Hello, world.');
  assert.deepEqual(r.events.filter((e) => 'cont' in e), [{ cont: 1 }]);
  assert.deepEqual(r.events.at(-1), { done: true, continued: 1, usage: { prompt_tokens: 220, completion_tokens: 15, total_tokens: 235 } });
  const m = r.calls[1].messages;
  assert.equal(m[0].role, 'system');
  assert.deepEqual(m.slice(-2).map((x) => x.role), ['assistant', 'user']);
  assert.equal(m.at(-2).content, 'Hello, ');
  assert.ok(m.at(-1).content.includes('이어서'));
  assert.ok(m.some((x) => x.role === 'user' && x.content.includes('FILEBODY')), 'attachments stay in the first user turn');
  assert.deepEqual([r.calls[1].model, r.calls[1].maxTokens], [r.calls[0].model, r.calls[0].maxTokens]);
  assert.ok(!r.text.includes('sk-test-secret'));
});

test('handler: a repeated tail at the seam is dropped, short continuations are not lost', async () => {
  _resetMemory();
  const tail = 'THE-LAST-LINE-WAS-RIGHT-HERE-0123456789';
  const first = 'start of the answer. ' + tail;
  const r = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
    upstream: (req, n) => n === 1
      ? sseResponse([chunk(first), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n'])
      : sseResponse([chunk(tail.slice(0, 20)), chunk(tail.slice(20) + ' and more.'), chunk(null, { finish_reason: 'stop' }), 'data: [DONE]\n\n']),
  });
  assert.equal(deltas(r), first + ' and more.');
  const long = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
    upstream: (req, n) => n === 1
      ? sseResponse([chunk(first), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n'])
      : sseResponse([chunk(tail + ' ' + 'z'.repeat(300)), chunk(null, { finish_reason: 'stop' }), 'data: [DONE]\n\n']),
  });
  assert.equal(deltas(long), first + ' ' + 'z'.repeat(300));
  const plain = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
    upstream: (req, n) => n === 1
      ? sseResponse([chunk('abc'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n'])
      : sseResponse([chunk('abc'), chunk(null, { finish_reason: 'stop' }), 'data: [DONE]\n\n']),
  });
  assert.equal(deltas(plain), 'abcabc', 'short overlaps are real text, not a seam');
});

test('handler: continuations count against the per-minute cap', async () => {
  _resetMemory();
  const warn = console.warn;
  console.warn = () => {};
  try {
    const r = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
      env: { AI_LIMIT_GLOBAL_MIN: '1' },
      upstream: () => sseResponse([chunk('cut'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n']),
    });
    assert.equal(r.calls.length, 1, 'no second gateway call once the minute cap is used');
    assert.equal(deltas(r), 'cut');
    assert.deepEqual(r.events.at(-1), { done: true, truncated: true });
    _resetMemory();
    const ok = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
      env: { AI_LIMIT_GLOBAL_MIN: '2' },
      upstream: (req, n) => n === 1 ? sseResponse([chunk('cut'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n']) : okStream('ted'),
    });
    assert.equal(deltas(ok), 'cutted');
  } finally {
    console.warn = warn;
  }
});

test('handler: continuation that cannot connect ends as truncated with the text so far', async () => {
  _resetMemory();
  const r = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
    upstream: (req, n) => n === 1
      ? sseResponse([chunk('part one'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n'])
      : new Response('{"detail":"boom"}', { status: 500 }),
  });
  assert.equal(deltas(r), 'part one');
  assert.deepEqual(r.events.at(-1), { done: true, truncated: true });
  const again = await call({ code: 'TESTCODE', ex: 's1-1', messages: [user('x')] }, {
    upstream: (req, n) => n === 1
      ? sseResponse([chunk('part one'), chunk(null, { finish_reason: 'length' }), 'data: [DONE]\n\n'])
      : sseResponse([chunk('part two'), 'data: {"error":{"message":"overloaded"}}\n\n']),
  });
  assert.equal(again.events.at(-1).error !== undefined, true, 'a failure after the continuation started is an error event');
  assert.ok(deltas(again).startsWith('part one'));
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

// ---------- attachments ----------

const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const JPG = 'data:image/jpeg;base64,/9j/4AAQSkZJRg==';
const jpgOfLength = (n) => `data:image/jpeg;base64,${'A'.repeat(n - 23)}`;
const doc = (text = '본문', name = 'a.pdf') => ({ name, text });
const pic = (url = PNG, name = 'p.png') => ({ name, url });
const CODE = 'TESTCODE';

test('validateAttachments: counts and sizes', () => {
  const va = (files, images, messages = [user('q')]) => validateAttachments(files, images, messages);
  const none = va(undefined, null);
  assert.deepEqual([none.ok, none.files, none.images], [true, [], []]);
  assert.deepEqual([va([], []).files, va([], []).images], [[], []]);
  assert.equal(va(Array.from({ length: 3 }, () => doc())).ok, true);
  assert.deepEqual([va(Array.from({ length: 4 }, () => doc())).status, va(Array.from({ length: 4 }, () => doc())).code], [413, 'too_long']);
  assert.equal(va([doc('x'.repeat(100000))]).ok, true);
  assert.equal(va([doc('x'.repeat(100001))]).status, 413);
  assert.equal(va([doc('x'.repeat(75000)), doc('x'.repeat(75000))]).ok, true);
  assert.equal(va([doc('x'.repeat(75000)), doc('x'.repeat(75001))]).status, 413);
  assert.equal(va(null, Array.from({ length: 4 }, () => pic())).ok, true);
  assert.equal(va(null, Array.from({ length: 5 }, () => pic())).status, 413);
  const max = jpgOfLength(1_399_999); // the longest valid base64 length under the 1,400,000 cap
  assert.equal(va(null, [pic(max)]).ok, true);
  assert.equal(va(null, [pic(jpgOfLength(1_400_003))]).status, 413);
  assert.equal(va(null, [pic(max), pic(max)]).ok, true);
  assert.equal(va(null, [pic(max), pic(max), pic(max)]).status, 413, 'all images together over 4,000,000');
  const msgOnly = validateMessages([user('x'.repeat(12000))]);
  assert.ok(msgOnly.ok && va([doc('x'.repeat(100000))], null, msgOnly.messages).ok, 'message caps do not count attachments');
});

test('validateAttachments: formats, empty text, first turn', () => {
  const va = (files, images, messages = [user('q')]) => validateAttachments(files, images, messages);
  const empty = va([doc(' \n\t ')]);
  assert.deepEqual([empty.status, empty.code], [400, 'bad_request']);
  assert.ok(empty.message.includes('«a.pdf»') && empty.message.includes('읽을 수 있는 글자'), empty.message);
  for (const bad of ['x', {}, 3, true]) {
    assert.equal(va(bad).status, 400, `files ${JSON.stringify(bad)}`);
    assert.equal(va(null, bad).status, 400, `images ${JSON.stringify(bad)}`);
  }
  for (const bad of [[null], ['text'], [{ name: 'a' }], [{ text: 1 }], [{ name: 5, text: 'x' }]]) assert.equal(va(bad).status, 400, JSON.stringify(bad));
  for (const bad of [[null], [PNG], [{ name: 'a' }], [{ url: 1 }]]) assert.equal(va(null, bad).status, 400, JSON.stringify(bad));
  const badUrls = [
    'data:image/gif;base64,R0lGODlh',
    'data:image/svg+xml;base64,PHN2Zz4=',
    'data:image/PNG;base64,iVBORw0KGgo=',
    'data:image/png;base64,',
    'data:image/png;base64,abc!',
    'data:image/png;base64,abc',
    'data:image/png;base64,ab==cd==',
    'data:image/png;base64,iVBO RW0K',
    'data:image/png;base64,iVBORw0KGgo=\n',
    ' data:image/png;base64,iVBORw0KGgo=',
    'data:image/png,iVBORw0KGgo=',
    'data:text/html;base64,PGI+',
    `data:image/jpeg;base64,${'é'.repeat(8)}`,
    'https://example.com/a.png',
  ];
  for (const url of badUrls) {
    const r = va(null, [pic(url)]);
    assert.deepEqual([r.status, r.code], [400, 'bad_request'], JSON.stringify(url));
  }
  for (const url of [PNG, JPG, 'data:image/webp;base64,UklGRg==', 'data:image/png;base64,ab+/']) assert.equal(va(null, [pic(url)]).ok, true, url);
  const t0 = Date.now();
  assert.equal(va(null, [pic(`${jpgOfLength(1_399_998)}!`)]).status, 400, 'bad last char of a long url');
  assert.ok(Date.now() - t0 < 1000, 'linear-time check');
  const lead = va([doc()], null, [{ role: 'assistant', content: '안녕하세요' }, user('q')]);
  assert.deepEqual([lead.status, lead.code], [400, 'bad_request']);
  assert.equal(va(null, null, [{ role: 'assistant', content: '안녕하세요' }, user('q')]).ok, true, 'no attachments: unchanged');
});

test('validateAttachments: names are cleaned and cut', () => {
  const va = (files, images) => validateAttachments(files, images, [user('q')]);
  const r = va([doc('t', `a\nb\r\nc\u0000d\u202e«e»\u2028f\t${'g'.repeat(200)}`)], [pic(PNG, '\u0007\n'), pic(JPG, '  사진  1 ')]);
  assert.ok(r.ok);
  const n = r.files[0].name;
  assert.ok(n.length <= 80, `length ${n.length}`);
  assert.ok(!/[\u0000-\u001f\u007f-\u009f\u2028\u2029\u202e«»]/.test(n), JSON.stringify(n));
  assert.ok(n.startsWith('a b c de f ggg'), JSON.stringify(n));
  assert.deepEqual(r.images.map((im) => im.name), ['이미지 1', '사진 1']);
  assert.equal(va([doc('t', `${'x'.repeat(79)}😀`)]).files[0].name, 'x'.repeat(79), 'no half surrogate');
  assert.deepEqual(va([{ text: 'a' }, { name: '', text: 'b' }]).files.map((f) => f.name), ['파일 1', '파일 2']);
  assert.equal(va([doc('본문', 'x'.repeat(5000))]).files[0].name.length, 80);
});

test('withAttachments: message shapes and order', () => {
  const msgs = [user('원래 질문'), { role: 'assistant', content: '답' }, user('후속 질문')];
  assert.equal(withAttachments(msgs, { files: [], images: [] }), msgs);
  const files = [doc('A본문', 'a.pdf'), doc('B본문', 'b.txt')];
  const block = '[첨부 자료 «a.pdf» 시작]\nA본문\n[첨부 자료 «a.pdf» 끝]\n\n[첨부 자료 «b.txt» 시작]\nB본문\n[첨부 자료 «b.txt» 끝]';
  const f = withAttachments(msgs, { files, images: [] });
  assert.deepEqual(f[0], user(`${block}\n\n원래 질문`));
  assert.deepEqual(f.slice(1), msgs.slice(1));
  const images = [pic(JPG, 'x.jpg'), pic(PNG, 'y.png')];
  const i = withAttachments(msgs, { files: [], images });
  assert.deepEqual(i[0], {
    role: 'user',
    content: [
      { type: 'text', text: '첨부 이미지 2장: x.jpg, y.png\n\n원래 질문' },
      { type: 'image_url', image_url: { url: JPG } },
      { type: 'image_url', image_url: { url: PNG } },
    ],
  });
  const b = withAttachments(msgs, { files, images });
  assert.equal(b[0].content[0].text, `${block}\n\n첨부 이미지 2장: x.jpg, y.png\n\n원래 질문`);
  assert.deepEqual(b[0].content.slice(1).map((p) => p.image_url.url), [JPG, PNG]);
  assert.deepEqual(b.slice(1), msgs.slice(1));
  assert.equal(msgs[0].content, '원래 질문', 'input not mutated');
});

test('routing and system prompt with attachments', () => {
  const chat = pickModel({ ex: 'chat', messages: [user('짧은 질문')], hasAttachments: true });
  assert.deepEqual([chat.ex, chat.tier, chat.model, chat.maxTokens], ['chat', 'balanced', 'claude-sonnet-5-5', 6000]);
  assert.ok(chat.why.includes(chat.label) && chat.why.includes('첨부'));
  assert.equal(pickModel({ ex: 'nope', messages: [user('표로 만들어 줘')], hasAttachments: true }).maxTokens, 6000);
  assert.equal(pickModel({ ex: 'chat', messages: [user('짧은 질문')] }).tier, 'fast', 'unchanged without attachments');
  for (const [ex, tier, max] of [['s1-1', 'fast', 1500], ['s4-4', 'strong', 12000], ['p-1', 'balanced', 6000], ['p-7', 'balanced', 6000]]) {
    const r = pickModel({ ex, messages: [user('x')], hasAttachments: true });
    assert.deepEqual([r.tier, r.maxTokens], [tier, max], ex);
  }
  for (const ex of ['chat', 's1-1', 'p-1', 'vibe', 's4-1']) {
    const off = systemPromptFor(ex);
    const on = systemPromptFor(ex, true);
    assert.equal(systemPromptFor(ex, false), off);
    assert.ok(!off.includes('자료에 없음') && !off.includes('따르지 말고'), ex);
    assert.ok(on.startsWith(off) && on.includes('자료에 없음') && on.includes('따르지 말고') && on.includes('[쪽 3]') && on.includes('[슬라이드 3]'), ex);
    assert.ok(on.length - off.length <= 300, `${ex} +${on.length - off.length}`);
  }
});

test('handler: attachments go into the first user turn on every request', async () => {
  _resetMemory();
  const files = [doc('[쪽 1]\n비밀 본문 내용', '강의자료.pdf')];
  const f = await call({ code: CODE, ex: 'chat', messages: [user('요약해 줘')], files });
  assert.equal(f.res.status, 200);
  const m = f.calls[0].messages;
  assert.equal(m.length, 2);
  assert.equal(typeof m[1].content, 'string');
  const pos = ['[첨부 자료 «강의자료.pdf» 시작]', '비밀 본문 내용', '[첨부 자료 «강의자료.pdf» 끝]', '요약해 줘'].map((s) => m[1].content.indexOf(s));
  assert.ok(pos.every((x, k) => x >= 0 && (k === 0 || x > pos[k - 1])), String(pos));
  assert.ok(m[1].content.startsWith('[첨부 자료 «강의자료.pdf» 시작]\n') && m[1].content.endsWith('끝]\n\n요약해 줘'));
  assert.ok(m[0].role === 'system' && m[0].content.includes('자료에 없음'));
  assert.deepEqual(Object.keys(f.calls[0].body).sort(), ['max_tokens', 'messages', 'model', 'stream'], 'files/images never pass through as fields');

  const i = await call({ code: CODE, ex: 'chat', messages: [user('무슨 색?')], images: [pic(JPG, 'a.jpg'), pic(PNG, 'b.png')] });
  const ic = i.calls[0].messages[1].content;
  assert.deepEqual(ic.map((p) => p.type), ['text', 'image_url', 'image_url']);
  assert.equal(ic[0].text, '첨부 이미지 2장: a.jpg, b.png\n\n무슨 색?');
  assert.deepEqual(ic.slice(1).map((p) => p.image_url.url), [JPG, PNG]);
  assert.ok(i.calls[0].messages[0].content.includes('자료에 없음'));

  const both = await call({ code: CODE, ex: 'chat', messages: [user('설명해 줘')], files, images: [pic(JPG, 'a.jpg')] });
  const bc = both.calls[0].messages[1].content;
  assert.deepEqual(bc.map((p) => p.type), ['text', 'image_url']);
  assert.ok(bc[0].text.startsWith('[첨부 자료 «강의자료.pdf» 시작]\n') && bc[0].text.endsWith('끝]\n\n첨부 이미지 1장: a.jpg\n\n설명해 줘'), bc[0].text);

  const turns = [user('첫 질문'), { role: 'assistant', content: '첫 답' }, user('둘째 질문')];
  const follow = await call({ code: CODE, ex: 'chat', messages: turns, files, images: [pic(JPG, 'a.jpg')] });
  const fm = follow.calls[0].messages;
  assert.equal(fm.length, 4);
  assert.ok(fm[1].content[0].text.includes('[첨부 자료 «강의자료.pdf» 시작]') && fm[1].content[0].text.endsWith('첫 질문'));
  assert.deepEqual(fm.slice(2), turns.slice(1), 'later turns untouched');
  assert.equal(follow.events[0].meta.tier, 'balanced', 'follow-ups keep the attachment route');

  const big = await call({ code: CODE, ex: 'chat', messages: [user('x'.repeat(12000))], files: [doc('가'.repeat(100000), 'a'), doc('나'.repeat(50000), 'b')] });
  assert.equal(big.res.status, 200, 'attachments do not count against the message caps');

  const plain = await call({ code: CODE, ex: 'chat', messages: [user('질문')] });
  assert.ok(!plain.calls[0].messages[0].content.includes('자료에 없음'));
  assert.equal(plain.calls[0].messages[1].content, '질문');
});

test('handler: routing with attachments', async () => {
  _resetMemory();
  const files = [doc('본문', 'a.txt')];
  const chat = await call({ code: CODE, ex: 'chat', messages: [user('짧은 질문')], files });
  assert.deepEqual([chat.calls[0].model, chat.calls[0].maxTokens, chat.events[0].meta.tier], ['claude-sonnet-5-5', 6000, 'balanced']);
  const imgs = await call({ code: CODE, messages: [user('hi')], images: [pic()] });
  assert.deepEqual([imgs.calls[0].maxTokens, imgs.events[0].meta.tier], [6000, 'balanced']);
  const p1 = await call({ code: CODE, ex: 'p-1', messages: [user('정리해 줘')], files });
  assert.deepEqual([p1.calls[0].model, p1.calls[0].maxTokens], ['claude-sonnet-5-5', 6000]);
  const s11 = await call({ code: CODE, ex: 's1-1', messages: [user('요약')], files });
  assert.deepEqual([s11.calls[0].model, s11.calls[0].maxTokens], ['claude-haiku-5-5', 1500], 'exercise models stay fixed');
  const vibe = await call({ code: CODE, ex: 'vibe', messages: [user('앱')], files });
  assert.ok(vibe.calls[0].messages[0].content.includes('```html') && vibe.calls[0].messages[0].content.includes('자료에 없음'));
  const plain = await call({ code: CODE, ex: 'chat', messages: [user('짧은 질문')] });
  assert.deepEqual([plain.calls[0].maxTokens, plain.events[0].meta.tier], [1500, 'fast']);
});

test('handler: bad attachments -> 4xx before any budget or upstream call', async () => {
  _resetMemory();
  const env = { AI_LIMIT_IP_10M: '1' };
  const q = (extra) => call({ code: CODE, ex: 'chat', messages: [user('q')], ...extra }, { env });
  const cases = [
    [{ files: Array.from({ length: 4 }, () => doc()) }, 413],
    [{ files: [doc('x'.repeat(100001))] }, 413],
    [{ files: [doc('x'.repeat(80000)), doc('x'.repeat(80000))] }, 413],
    [{ files: [doc('  \n ')] }, 400],
    [{ files: 'x' }, 400],
    [{ images: [pic('data:image/gif;base64,R0lGODlh')] }, 400],
    [{ images: [pic('data:image/png;base64,abc!')] }, 400],
    [{ images: [pic('https://example.com/a.png')] }, 400],
    [{ images: Array.from({ length: 5 }, () => pic()) }, 413],
    [{ images: [pic(jpgOfLength(1_400_003))] }, 413],
    [{ images: [pic(jpgOfLength(1_399_999)), pic(jpgOfLength(1_399_999)), pic(jpgOfLength(1_399_999))] }, 413],
    [{ images: [pic(jpgOfLength(SIZE.maxBodyBytes + 100))] }, 413],
    [{ messages: [{ role: 'assistant', content: '안녕하세요' }, user('q')], files: [doc()] }, 400],
  ];
  for (const [extra, status] of cases) {
    const r = await q(extra);
    const label = JSON.stringify(extra).slice(0, 90);
    assert.equal(r.res.status, status, label);
    assert.ok(/[가-힣]/.test(parseJson(r).error), label);
    assert.equal(parseJson(r).code, status === 413 ? 'too_long' : 'bad_request', label);
    assert.equal(r.calls.length, 0, label);
  }
  const wrong = await q({ code: 'nope', files: 'junk', images: 'junk' });
  assert.deepEqual([wrong.res.status, wrong.calls.length], [401, 0], 'the code check comes first');
  const ok = await q({ files: [doc()] });
  assert.equal(ok.res.status, 200, 'rejected requests used no budget (ip10m = 1)');
  const over = await q({ files: [doc()] });
  assert.equal(over.res.status, 429, 'accepted requests do count');
});

test('handler: check:true ignores attachments', async () => {
  _resetMemory();
  const r = await call({ code: CODE, check: true, files: Array.from({ length: 9 }, () => ({ text: '' })), images: 'junk' });
  assert.deepEqual([r.res.status, parseJson(r), r.calls.length], [200, { ok: true }, 0]);
  const good = await call({ code: CODE, check: true, files: [doc()], images: [pic()] });
  assert.deepEqual([good.res.status, parseJson(good), good.calls.length], [200, { ok: true }, 0]);
});

test('handler: attachment names, text and urls never reach the logs', async () => {
  _resetMemory();
  const logs = [];
  const saved = { log: console.log, info: console.info, warn: console.warn, error: console.error, debug: console.debug };
  for (const k of Object.keys(saved)) console[k] = (...a) => logs.push(a.map(String).join(' '));
  try {
    const files = [doc('SECRETTEXT', 'SECRETNAME.pdf')];
    const images = [pic('data:image/jpeg;base64,U0VDUkVUSU1H', 'SECRETIMG.jpg')];
    const r = await call({ code: CODE, ex: 's4-4', messages: [user('x')], files, images }, { upstream: (req, n) => (n === 1 ? new Response('{"detail":"boom"}', { status: 500 }) : okStream()) });
    assert.equal(r.res.status, 200);
    await call({ code: CODE, ex: 'chat', messages: [user('x')], files: [doc(' ', 'SECRETEMPTY')] });
    await call({ code: CODE, ex: 'chat', messages: [user('x')], images: [pic('data:image/png;base64,SECRET!', 'SECRETBAD')] });
    await call({ code: CODE, ex: 'chat', messages: [user('x')], files, images }, { upstream: () => { throw new TypeError('boom'); } });
  } finally {
    Object.assign(console, saved);
  }
  assert.ok(logs.length >= 1, 'the fallback was logged');
  assert.ok(!logs.some((l) => /SECRET|U0VDUkVU|base64/.test(l)), logs.join('\n'));
});

test('views: unknown view ignored, known view forces balanced model and length, adds its prompt', () => {
  assert.equal(normalizeView('x'), null);
  assert.equal(normalizeView({}), null);
  assert.equal(normalizeView('constructor'), null);
  const labels = { 'claude-sonnet-5-5': 'Claude Sonnet 5.5' };
  const base = pickModel({ ex: 's1-1', messages: [user('hi')] });
  assert.equal(base.tier, 'fast');
  assert.equal(applyView(base, null, TIERS, labels), base);
  const r = applyView(base, 'doc', TIERS, labels);
  assert.equal(r.model, 'claude-sonnet-5-5');
  assert.equal(r.maxTokens, VIEWS.doc.maxTokens);
  assert.ok(r.why.includes('Claude Sonnet 5.5'));
  assert.ok(systemPromptFor('s2-3', false, 'doc').includes(viewPrompt('doc', 's2-3')));
  assert.ok(viewPrompt('doc', 's2-3').length > 1000);
  assert.ok(!systemPromptFor('s2-3', false, null).includes(viewPrompt('doc', 's2-3')));
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
