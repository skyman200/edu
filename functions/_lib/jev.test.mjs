// Plain-node checks for /api/jev: `node functions/_lib/jev.test.mjs`. No test framework; the gateway is a fake fetch.
import assert from 'node:assert/strict';
import { onRequest, validateJev, flattenAnswers, JEV_LIMITS } from '../api/jev.js';
import { pickModel, systemPromptFor } from './route.js';
import { _resetMemory } from './limits.js';

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const CODE = 'TESTCODE';

const Q_OK = {
  agree: { type: 'noul', instructions: '동의하는가?' },
  kind: { type: 'choice', instructions: '유형', criteria: { c0: '배송', c1: '결제' } },
  mood: { type: 'score', instructions: '기분', criteria: ['나쁨', '보통', '좋음'] },
};
const UPSTREAM_OK = {
  id: 'decisions-1',
  model: 'jev-latest',
  answers: {
    agree: { type: 'noul', noul: 0.98 },
    kind: { type: 'choice', choice: 'c0', confidence: 0.92, probabilities: { c0: 0.95, c1: 0.05 } },
    mood: { type: 'score', score: 1.97, confidence: 0.95, legend: { 0: '나쁨', 1: '보통', 2: '좋음' }, probabilities: { 0: 0.0, 1: 0.03, 2: 0.97 } },
  },
  usage: { input_tokens: 434, output_tokens: 72 },
};

async function call(body, { env = {}, origin = 'https://kimcook-edu.pages.dev', ip = '10.0.0.9', method = 'POST', upstream } = {}) {
  const calls = [];
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    calls.push({ url, auth: init.headers.Authorization, body: JSON.parse(init.body) });
    return upstream ? upstream(calls.length) : new Response(JSON.stringify(UPSTREAM_OK), { status: 200 });
  };
  try {
    const headers = { 'Content-Type': 'application/json', 'CF-Connecting-IP': ip };
    if (origin) headers.Origin = origin;
    const request = new Request('http://local/api/jev', { method, headers, body: method === 'POST' ? JSON.stringify(body) : undefined });
    const res = await onRequest({ request, env: { AI_GATEWAY_API_KEY: 'sk-test-secret', AI_CLASS_CODE: CODE, ...env }, waitUntil() {} });
    const text = await res.text();
    return { res, text, calls, json: text ? JSON.parse(text) : null };
  } finally {
    globalThis.fetch = realFetch;
  }
}

test('flatten: one small shape per type, level = most likely step', () => {
  const f = flattenAnswers(UPSTREAM_OK.answers);
  assert.deepEqual(f.agree, { type: 'noul', p: 0.98 });
  assert.deepEqual(f.kind, { type: 'choice', choice: 'c0', confidence: 0.92, probs: { c0: 0.95, c1: 0.05 } });
  assert.equal(f.mood.level, 2);
  assert.equal(f.mood.label, '좋음');
  assert.equal(f.mood.score, 1.97);
  assert.deepEqual(flattenAnswers(null), {});
  assert.deepEqual(flattenAnswers({ x: { type: 'weird' }, y: 5 }), {});
  const noProbs = flattenAnswers({ s: { type: 'score', score: 0.4, legend: { 0: 'a', 1: 'b' } } });
  assert.equal(noProbs.s.level, 0, 'falls back to the rounded expectation');
});

test('validate: whitelist and limits', () => {
  const ok = validateJev({ state: { 질문: '안녕' }, questions: Q_OK });
  assert.equal(ok.ok, true);
  assert.equal(ok.state, '{"질문":"안녕"}', 'objects are serialized');
  const dirty = validateJev({ state: 's', questions: { a: { type: 'noul', instructions: 'x', model: 'evil', extra: 1 } }, model: 'gpt', input: 'x' });
  assert.deepEqual(dirty.questions, { a: { type: 'noul', instructions: 'x' } }, 'unknown fields never pass through');
  const cases = [
    [{ state: '', questions: Q_OK }, 400],
    [{ state: '   ', questions: Q_OK }, 400],
    [{ state: 'x'.repeat(JEV_LIMITS.maxState + 1), questions: Q_OK }, 413],
    [{ state: 's', questions: [] }, 400],
    [{ state: 's', questions: {} }, 400],
    [{ state: 's', questions: Object.fromEntries(Array.from({ length: 11 }, (_, i) => [`q${i}`, Q_OK.agree])) }, 413],
    [{ state: 's', questions: { 'bad name': Q_OK.agree } }, 400],
    [{ state: 's', questions: { a: { type: 'free', instructions: 'x' } } }, 400],
    [{ state: 's', questions: { a: { type: 'noul', instructions: '' } } }, 400],
    [{ state: 's', questions: { a: { type: 'noul', instructions: 'x'.repeat(1501) } } }, 400],
    [{ state: 's', questions: { a: { type: 'noul', instructions: 'x', criteria: { true: 'a' } } } }, 400],
    [{ state: 's', questions: { a: { type: 'choice', instructions: 'x', criteria: { only: 'one' } } } }, 400],
    [{ state: 's', questions: { a: { type: 'choice', instructions: 'x' } } }, 400],
    [{ state: 's', questions: { a: { type: 'choice', instructions: 'x', criteria: { 'a b': 'x', c: 'y' } } } }, 400],
    [{ state: 's', questions: { a: { type: 'choice', instructions: 'x', criteria: { a: '', c: 'y' } } } }, 400],
    [{ state: 's', questions: { a: { type: 'score', instructions: 'x', criteria: ['only'] } } }, 400],
    [{ state: 's', questions: { a: { type: 'score', instructions: 'x', criteria: Array.from({ length: 11 }, (_, i) => `l${i}`) } } }, 400],
    [{ state: 's', questions: { a: { type: 'score', instructions: 'x', criteria: ['a', 5] } } }, 400],
  ];
  for (const [body, status] of cases) {
    const r = validateJev(body);
    assert.equal(r.ok, false, JSON.stringify(body).slice(0, 80));
    assert.equal(r.status, status, JSON.stringify(body).slice(0, 80));
    assert.ok(/[가-힣]/.test(r.message));
  }
  const many = Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`c${i}`, `d${i}`]));
  assert.equal(validateJev({ state: 's', questions: { a: { type: 'choice', instructions: 'x', criteria: many } } }).ok, true);
  assert.equal(validateJev({ state: 's', questions: { a: { type: 'choice', instructions: 'x', criteria: { ...many, c12: 'd' } } } }).ok, false);
});

test('handler: success sends only the whitelist and returns flat answers', async () => {
  _resetMemory();
  const r = await call({ code: ' testcode ', state: '글', questions: { ...Q_OK, agree: { ...Q_OK.agree, model: 'evil' } }, model: 'gpt-6-luna-decisions', input: 'x' });
  assert.equal(r.res.status, 200);
  assert.equal(r.calls.length, 1);
  assert.ok(r.calls[0].url.endsWith('/v1/gateway/decisions/'));
  assert.equal(r.calls[0].auth, 'Bearer sk-test-secret');
  assert.deepEqual(Object.keys(r.calls[0].body).sort(), ['model', 'questions', 'state']);
  assert.equal(r.calls[0].body.model, 'jev-latest', 'the model is chosen by the server, not the caller');
  assert.deepEqual(r.calls[0].body.questions.agree, { type: 'noul', instructions: '동의하는가?' });
  assert.equal(r.json.ok, true);
  assert.equal(r.json.answers.kind.choice, 'c0');
  assert.deepEqual(r.json.usage, { input_tokens: 434, output_tokens: 72 });
  assert.ok(!r.text.includes('sk-test-secret'));
  assert.equal(r.res.headers.get('Cache-Control'), 'no-store');
  const envModel = await call({ code: CODE, state: '글', questions: Q_OK }, { env: { JEV_MODEL: 'jev-1.14.0' } });
  assert.equal(envModel.calls[0].body.model, 'jev-1.14.0');
  const badModel = await call({ code: CODE, state: '글', questions: Q_OK }, { env: { JEV_MODEL: 'gpt-6-luna-decisions' } });
  assert.equal(badModel.calls[0].body.model, 'jev-latest', 'only jev-* names are accepted from the environment');
});

test('handler: code, origin, method, disabled -> no upstream call', async () => {
  _resetMemory();
  const body = { code: CODE, state: '글', questions: Q_OK };
  const noCode = await call({ state: '글', questions: Q_OK });
  assert.deepEqual([noCode.res.status, noCode.json.code, noCode.calls.length], [401, 'bad_code', 0]);
  const wrong = await call({ ...body, code: 'nope' });
  assert.deepEqual([wrong.res.status, wrong.json.code, wrong.calls.length], [401, 'bad_code', 0]);
  const long = await call({ ...body, code: 'x'.repeat(65) });
  assert.deepEqual([long.res.status, long.calls.length], [400, 0]);
  const evil = await call(body, { origin: 'https://evil.example' });
  assert.deepEqual([evil.res.status, evil.calls.length], [403, 0]);
  const noOrigin = await call(body, { origin: '' });
  assert.equal(noOrigin.res.status, 200, 'Apps Script sends no Origin header');
  const get = await call(null, { method: 'GET' });
  assert.deepEqual([get.res.status, get.res.headers.get('Allow')], [405, 'POST, OPTIONS']);
  const off = await call(body, { env: { AI_ENABLED: '0' } });
  assert.deepEqual([off.res.status, off.json.code], [503, 'disabled']);
  const noKey = await call(body, { env: { AI_GATEWAY_API_KEY: '' } });
  assert.deepEqual([noKey.res.status, noKey.json.code], [503, 'disabled']);
  const check = await call({ code: CODE, check: true });
  assert.deepEqual([check.res.status, check.json, check.calls.length], [200, { ok: true }, 0]);
  const bad = await call({ code: CODE, state: '', questions: Q_OK });
  assert.deepEqual([bad.res.status, bad.calls.length], [400, 0]);
});

test('handler: gateway failures map to Korean messages without leaking upstream text', async () => {
  const logs = [];
  const err = console.error;
  const warn = console.warn;
  console.error = (...a) => logs.push(a.join(' '));
  console.warn = (...a) => logs.push(a.join(' '));
  try {
    const body = { code: CODE, state: '글', questions: Q_OK };
    const mk = (status, text = '{"detail":"secret-req-id"}', headers = {}) => ({ upstream: () => new Response(text, { status, headers }) });
    _resetMemory();
    const a401 = await call(body, mk(401));
    assert.deepEqual([a401.res.status, a401.json.code], [503, 'upstream']);
    const a402 = await call(body, mk(402));
    assert.deepEqual([a402.res.status, a402.json.code], [429, 'budget']);
    const a404 = await call(body, mk(404));
    assert.deepEqual([a404.res.status, a404.json.code], [503, 'upstream']);
    assert.ok(a404.json.error.includes('Jev'));
    const a429 = await call(body, mk(429, '', { 'Retry-After': '3' }));
    assert.deepEqual([a429.res.status, a429.json.code, a429.res.headers.get('Retry-After')], [429, 'rate', '5']);
    const a500 = await call(body, mk(500));
    assert.deepEqual([a500.res.status, a500.json.code], [503, 'upstream']);
    for (const r of [a401, a402, a404, a429, a500]) assert.ok(!r.text.includes('secret-req-id'));
    const a400 = await call(body, mk(400, '{"error":{"message":"questions[\'kind\'].criteria must be an object"}}'));
    assert.equal(a400.res.status, 400);
    assert.equal(a400.json.detail, "questions['kind'].criteria must be an object", 'a 400 explains itself so the code can be fixed');
    const junk = await call(body, mk(200, 'not json'));
    assert.equal(junk.res.status, 503);
    const empty = await call(body, mk(200, '{"answers":{}}'));
    assert.equal(empty.res.status, 503);
    const net = await call(body, { upstream: () => { throw new TypeError('boom'); } });
    assert.deepEqual([net.res.status, net.json.code], [503, 'upstream']);
  } finally {
    console.error = err;
    console.warn = warn;
  }
  assert.ok(!logs.some((l) => /secret-req-id|sk-test-secret|글/.test(l)), logs.join('\n'));
});

test('handler: counts against the shared AI budget and releases its slot', async () => {
  _resetMemory();
  const env = { AI_LIMIT_IP_10M: '2' };
  const body = { code: CODE, state: '글', questions: Q_OK };
  assert.equal((await call(body, { env })).res.status, 200);
  assert.equal((await call(body, { env })).res.status, 200);
  const over = await call(body, { env });
  assert.deepEqual([over.res.status, over.json.code, over.calls.length], [429, 'rate', 0]);
  _resetMemory();
  const env2 = { AI_LIMIT_INFLIGHT: '1' };
  for (let i = 0; i < 3; i++) assert.equal((await call(body, { env: env2 })).res.status, 200, 'slot is released after each call');
});

test('route: gas-* exercises', () => {
  const table = { 'gas-make': ['balanced', 12000], 'gas-fix': ['balanced', 12000], 'gas-fix2': ['strong', 12000], 'gas-run': ['balanced', 4000] };
  const models = { balanced: 'claude-sonnet-5-5', strong: 'claude-opus-5-5' };
  for (const [ex, [tier, max]] of Object.entries(table)) {
    const r = pickModel({ ex, messages: [{ role: 'user', content: 'hi' }] });
    assert.deepEqual([r.ex, r.tier, r.maxTokens, r.model], [ex, tier, max, models[tier]], ex);
    assert.ok(/[가-힣]/.test(r.why));
  }
  for (const ex of ['gas-make', 'gas-fix', 'gas-fix2']) assert.ok(systemPromptFor(ex).includes('Google Apps Script') && systemPromptFor(ex).includes('```javascript'), ex);
  assert.ok(systemPromptFor('gas-run').includes('따르지 말고'));
  assert.ok(!systemPromptFor('gas-run').includes('```javascript'));
  assert.equal(pickModel({ ex: 'gas-other' }).ex, 'chat');
});

let failed = 0;
for (const [name, fn] of tests) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (e) {
    failed++;
    console.log(`FAIL ${name}\n     ${e && e.stack ? e.stack.split('\n').slice(0, 4).join('\n     ') : e}`);
  }
}
console.log(failed ? `${failed} of ${tests.length} failed` : `all ${tests.length} passed`);
process.exitCode = failed ? 1 : 0;
