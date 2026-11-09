// Plain-node checks for /api/ta (the hosted 24시간 조교): `node functions/_lib/ta.test.mjs`
// D1 is a real SQLite (node:sqlite) behind the D1 interface; the AI ONE gateway is a fake fetch that answers Jev and chat calls.
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { onRequest, pickExcerpts, joinDocs, cleanRoom, TA, _resetSchema } from '../api/ta.js';
import { _resetMemory } from './limits.js';

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const CODE = 'TESTCODE';
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function fakeD1() {
  const db = new DatabaseSync(':memory:');
  const isRead = (sql) => /^\s*select/i.test(sql);
  const wrap = (sql) => {
    let args = [];
    const o = {
      sql,
      bind(...a) { args = a; return o; },
      first() { return Promise.resolve(db.prepare(sql).get(...args) ?? null); },
      all() { return Promise.resolve({ results: db.prepare(sql).all(...args) }); },
      run() { const r = db.prepare(sql).run(...args); return Promise.resolve({ meta: { changes: Number(r.changes), last_row_id: Number(r.lastInsertRowid) } }); },
    };
    return o;
  };
  return {
    prepare: wrap,
    async batch(list) {
      const out = [];
      for (const s of list) out.push(isRead(s.sql) ? await s.all() : await s.run());
      return out;
    },
    raw: db,
  };
}

const sse = (text) => `data: ${JSON.stringify({ choices: [{ delta: { content: text }, finish_reason: null }] })}\n\ndata: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\ndata: [DONE]\n\n`;

// state the fake gateway answers with; tests change it
function setup(extra = {}) {
  _resetMemory();
  _resetSchema();
  const d1 = fakeD1();
  const gw = { answerable: 0.95, personal: 0.1, kind: 'schedule', aiText: '출석은 4회 결석까지입니다. 「결석 4회 초과 시 F」 (강의계획서)', aiCalls: [], jevCalls: [], failAi: false };
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const body = JSON.parse(init.body);
    if (String(url).includes('/decisions')) {
      gw.jevCalls.push(body);
      return new Response(JSON.stringify({
        id: 'd1', model: 'jev-latest',
        answers: {
          answerable: { type: 'noul', noul: gw.answerable },
          personal: { type: 'noul', noul: gw.personal },
          kind: { type: 'choice', choice: gw.kind, confidence: 0.9, probabilities: { [gw.kind]: 0.9 } },
        },
      }), { status: 200 });
    }
    gw.aiCalls.push(body);
    if (gw.failAi) return new Response('{"error":{"message":"overloaded"}}', { status: 503 });
    return new Response(sse(gw.aiText), { status: 200, headers: { 'Content-Type': 'text/event-stream' } });
  };
  const env = { LIVE_DB: d1, AI_GATEWAY_API_KEY: 'sk-test', AI_CLASS_CODE: CODE, ...extra };
  let n = 0;
  const call = async (body, { ip = '10.0.0.' + (++n), origin = 'https://kimcook-edu.pages.dev' } = {}) => {
    const headers = { 'Content-Type': 'application/json', 'CF-Connecting-IP': ip };
    if (origin) headers.Origin = origin;
    const res = await onRequest({ request: new Request('http://local/api/ta', { method: 'POST', headers, body: JSON.stringify(body) }), env, waitUntil() {} });
    const text = await res.text();
    return { res, json: text ? JSON.parse(text) : null };
  };
  return { d1, gw, env, call, done() { globalThis.fetch = realFetch; } };
}

async function room(t, docs = [['강의계획서', '출석: 4회 초과 결석 시 F.\n과제: 매주 금요일 23:59까지 제출.\n중간고사: 10월 24일.']]) {
  const c = await t.call({ op: 'create', code: CODE, name: '교육공학 조교' });
  assert.equal(c.res.status, 200, JSON.stringify(c.json));
  const base = { room: c.json.room, key: c.json.key };
  for (const [name, text] of docs) assert.equal((await t.call({ op: 'addDoc', ...base, name, text })).res.status, 200);
  return base;
}

test('create needs the class code; the key is shown once and only its hash is stored', async () => {
  const t = setup();
  try {
    assert.equal((await t.call({ op: 'create', name: 'x' })).res.status, 401);
    assert.equal((await t.call({ op: 'create', code: 'WRONG', name: 'x' })).json.code, 'bad_code');
    const ok = await t.call({ op: 'create', code: ' testcode ', name: '  교육공학   조교 ' });
    assert.equal(ok.res.status, 200);
    assert.match(ok.json.room, /^[a-z2-9]{8}$/);
    assert.match(ok.json.key, /^[a-z2-9]{24}$/);
    const row = t.d1.raw.prepare('SELECT * FROM ta_rooms').get();
    assert.equal(row.name, '교육공학 조교');
    assert.ok(!JSON.stringify(row).includes(ok.json.key), 'the key itself is not stored');
    assert.equal(cleanRoom('ABCDEFGH'), 'abcdefgh');
    assert.equal(cleanRoom('../x'), '');
  } finally { t.done(); }
});

test('professor operations need the key; students get no inbox', async () => {
  const t = setup();
  try {
    const r = await room(t);
    assert.equal((await t.call({ op: 'state', room: r.room })).res.status, 403);
    assert.equal((await t.call({ op: 'state', room: r.room, key: 'wrong' })).json.code, 'bad_key');
    assert.equal((await t.call({ op: 'delDoc', room: r.room, key: 'x', id: 1 })).res.status, 403);
    assert.equal((await t.call({ op: 'state', room: 'zzzzzzzz', key: r.key })).res.status, 404);
    const s = await t.call({ op: 'state', ...r });
    assert.equal(s.res.status, 200);
    assert.deepEqual([s.json.name, s.json.open, s.json.docs.length, s.json.chars < s.json.maxChars], ['교육공학 조교', true, 1, true]);
    const info = await t.call({ op: 'info', room: r.room });
    assert.deepEqual([info.json.open, info.json.ready, info.json.key], [true, true, undefined], 'students see name and readiness only');
  } finally { t.done(); }
});

test('documents: cleaned, capped at the room total (the last one is cut), deletable', async () => {
  const t = setup();
  try {
    const r = await room(t, []);
    assert.equal((await t.call({ op: 'addDoc', ...r, name: 'x', text: '   ' })).json.code, 'empty');
    const big = await t.call({ op: 'addDoc', ...r, name: '큰 자료', text: '가'.repeat(TA.maxChars - 300) });
    assert.equal(big.json.cut, false);
    const more = await t.call({ op: 'addDoc', ...r, name: '넘치는 자료', text: '나'.repeat(5000) });
    assert.deepEqual([more.json.cut, more.json.chars], [true, 300], 'the text is cut to the space that is left');
    const full = await t.call({ op: 'addDoc', ...r, name: '더', text: '다'.repeat(50) });
    assert.equal(full.json.code, 'too_long');
    const st = await t.call({ op: 'state', ...r });
    assert.equal(st.json.chars, TA.maxChars);
    await t.call({ op: 'delDoc', ...r, id: st.json.docs[0].id });
    assert.equal((await t.call({ op: 'state', ...r })).json.docs.length, 1);
    assert.equal((await t.call({ op: 'delDoc', room: r.room, key: r.key, id: 'abc' })).res.status, 200, 'a bad id deletes nothing');
  } finally { t.done(); }
});

test('ask: Jev + AI answer from the material, citing the file; the file names reach the AI', async () => {
  const t = setup();
  try {
    const r = await room(t);
    const a = await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0001a' });
    assert.deepEqual([a.json.ok, a.json.auto, a.json.kind], [true, true, '일정·행정']);
    assert.ok(a.json.answer.includes('결석') && a.json.p === 95);
    assert.equal(t.gw.aiCalls.length, 1);
    const sent = JSON.stringify(t.gw.aiCalls[0].messages);
    assert.ok(sent.includes('[파일: 강의계획서]') && sent.includes('출석: 4회 초과'), 'the AI reads the whole material with file names');
    assert.ok(sent.includes('결석 몇 번까지'));
    assert.ok(JSON.stringify(t.gw.jevCalls[0]).includes('출석'), 'Jev reads the material too');
    const st = await t.call({ op: 'state', ...r });
    assert.deepEqual([st.json.qs[0].auto, st.json.waiting], [true, 0]);
    assert.equal((await t.call({ op: 'check', room: r.room, id: a.json.id })).json.found, false, 'automatic answers are not in the professor-answer lookup');
  } finally { t.done(); }
});

test('ask: low confidence, personal questions and "자료에 없습니다" go to the professor, who answers by number', async () => {
  const t = setup();
  try {
    const r = await room(t);
    t.gw.answerable = 0.3;
    const low = await t.call({ op: 'ask', room: r.room, q: '기말 발표 순서는 어떻게 정해요?', sid: 'sid0002a' });
    assert.deepEqual([low.json.auto, low.json.answer], [false, undefined]);
    assert.equal(t.gw.aiCalls.length, 0, 'no AI answer is written when Jev says the material cannot answer');
    assert.match(low.json.id, /^[A-Z2-9]{6}$/);
    t.gw.answerable = 0.95; t.gw.personal = 0.9;
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '성적 이의 신청하고 싶습니다', sid: 'sid0002a' })).json.auto, false);
    t.gw.personal = 0.1; t.gw.aiText = '자료에 없습니다';
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '교수님 연구실이 어디예요?', sid: 'sid0002a' })).json.auto, false);
    const st = await t.call({ op: 'state', ...r });
    assert.equal(st.json.waiting, 3);
    assert.equal((await t.call({ op: 'check', room: r.room, id: low.json.id })).json.answer, '', 'no answer yet');
    assert.equal((await t.call({ op: 'reply', ...r, id: 1, answer: '  ' })).json.code, 'empty');
    const q = st.json.qs.find((x) => x.code === low.json.id);
    assert.equal((await t.call({ op: 'reply', ...r, id: q.id, answer: '추첨으로 정합니다.' })).res.status, 200);
    assert.deepEqual((await t.call({ op: 'check', room: r.room, id: low.json.id.toLowerCase() })).json, { ok: true, found: true, answer: '추첨으로 정합니다.' });
    assert.equal((await t.call({ op: 'check', room: r.room, id: 'ZZZZZZ' })).json.found, false);
    assert.equal((await t.call({ op: 'check', room: 'zzzzzzzz', id: low.json.id })).res.status, 404, 'another room cannot read it');
    assert.equal((await t.call({ op: 'state', ...r })).json.waiting, 2);
  } finally { t.done(); }
});

test('ask: closed room, short question, no material, per-phone and per-room limits, AI failure falls back to the professor', async () => {
  const t = setup();
  try {
    const empty = await room(t, []);
    assert.equal((await t.call({ op: 'ask', room: empty.room, q: '질문이 있습니다 정말로', sid: 'sid0003a' })).json.msg, '아직 수업 자료가 준비되지 않았습니다. 교수님께 알려 주세요.');
    assert.equal((await t.call({ op: 'info', room: empty.room })).json.ready, false);
    const r = await room(t);
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '네?', sid: 'sid0003a' })).json.msg, '질문을 조금 더 자세히 써 주세요.');
    await t.call({ op: 'settings', ...r, open: false });
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '질문을 받나요 지금', sid: 'sid0003a' })).json.msg, '지금은 질문을 받지 않습니다.');
    await t.call({ op: 'settings', ...r, open: true });
    t.gw.failAi = true;
    const fb = await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0003a' });
    assert.deepEqual([fb.json.ok, fb.json.auto], [true, false], 'a failed AI answer still reaches the professor');
    t.gw.failAi = false;
    for (let i = 0; i < TA.sidPer10m - 1; i++) assert.equal((await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요 ' + i, sid: 'sid0003a' })).json.ok, true);
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요 또', sid: 'sid0003a' })).json.ok, false, 'one phone is limited');
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요 또', sid: 'sid0004a' })).json.ok, true, 'another phone is not');
  } finally { t.done(); }
});

test('settings: name, confidence bar (0.5 to 0.95), the confidence changes who answers', async () => {
  const t = setup();
  try {
    const r = await room(t);
    await t.call({ op: 'settings', ...r, name: '새 이름', minP: 0.99 });
    let s = (await t.call({ op: 'state', ...r })).json;
    assert.deepEqual([s.name, s.minP], ['새 이름', 0.95]);
    t.gw.answerable = 0.9;
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0005a' })).json.auto, false, 'strict bar sends it to the professor');
    await t.call({ op: 'settings', ...r, minP: 0.65 });
    assert.equal((await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0005b' })).json.auto, true);
  } finally { t.done(); }
});

test('pickExcerpts: all material when it fits, otherwise the passages closest to the question, in order, with file names', () => {
  const small = [{ name: 'a', text: '짧은 글' }];
  assert.equal(pickExcerpts(small, '아무거나'), joinDocs(small));
  const filler = Array.from({ length: 120 }, (_, i) => `${i}주차 수업 내용은 광합성과 호흡을 다룹니다.`).join('\n');
  const docs = [{ name: '강의계획서', text: filler + '\n성적 이의 신청은 성적 공개 후 7일 이내에 합니다.\n' + filler }, { name: '공지', text: '중간고사는 3관 201호에서 봅니다.' }];
  const out = pickExcerpts(docs, '성적 이의 신청 언제까지 해요?', 3000);
  assert.ok(out.length <= 3000);
  assert.ok(out.includes('성적 이의 신청은 성적 공개 후 7일 이내'), 'the relevant passage is kept');
  assert.ok(out.indexOf('[파일: 강의계획서]') < out.indexOf('성적 이의'), 'the file name line comes first');
});

test('ocr: a photo of a notice is typed out by the AI; digest summarizes the week', async () => {
  const t = setup();
  try {
    const r = await room(t);
    t.gw.aiText = '11월 20일 중간고사 3관 201호';
    const o = await t.call({ op: 'ocr', ...r, images: [{ name: '공지', url: PNG }] });
    assert.equal(o.json.text, '11월 20일 중간고사 3관 201호');
    const sent = t.gw.aiCalls.at(-1).messages.at(-1);
    assert.ok(Array.isArray(sent.content) && sent.content.some((p) => p.type === 'image_url'), 'the photo reaches the gateway');
    assert.equal((await t.call({ op: 'ocr', ...r, images: [] })).json.code, 'empty');
    assert.equal((await t.call({ op: 'ocr', room: r.room, key: 'bad', images: [{ name: 'a', url: PNG }] })).res.status, 403);
    assert.equal((await t.call({ op: 'digest', ...r })).json.code, 'empty', 'no questions yet');
    await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0006a' });
    t.gw.aiText = '1. 출석 — 1건';
    const d = await t.call({ op: 'digest', ...r });
    assert.equal(d.json.text, '1. 출석 — 1건');
    assert.ok(JSON.stringify(t.gw.aiCalls.at(-1)).includes('결석 몇 번까지'));
  } finally { t.done(); }
});

test('remove deletes the room with its material and questions; a wrong class code cannot lock a professor out', async () => {
  const t = setup();
  try {
    const r = await room(t);
    await t.call({ op: 'ask', room: r.room, q: '결석 몇 번까지 되나요?', sid: 'sid0007a' });
    assert.equal((await t.call({ op: 'remove', room: r.room, key: 'bad' })).res.status, 403);
    assert.equal((await t.call({ op: 'remove', ...r })).res.status, 200);
    for (const tb of ['ta_rooms', 'ta_docs', 'ta_qs']) assert.equal(t.d1.raw.prepare(`SELECT COUNT(*) AS n FROM ${tb}`).get().n, 0, tb);
    assert.equal((await t.call({ op: 'info', room: r.room })).res.status, 404);
  } finally { t.done(); }
});

test('request hygiene: origin, method, missing database, kill switch', async () => {
  const t = setup();
  try {
    assert.equal((await t.call({ op: 'info', room: 'abcdefgh' }, { origin: 'https://evil.example' })).res.status, 403);
    const get = await onRequest({ request: new Request('http://local/api/ta'), env: t.env, waitUntil() {} });
    assert.equal(get.status, 405);
    const nodb = await onRequest({ request: new Request('http://local/api/ta', { method: 'POST', body: '{}' }), env: { ...t.env, LIVE_DB: undefined }, waitUntil() {} });
    assert.equal(nodb.status, 503);
    const off = await t.call({ op: 'info', room: 'abcdefgh' });
    assert.equal(off.res.status, 404);
    t.env.AI_ENABLED = '0';
    assert.equal((await t.call({ op: 'info', room: 'abcdefgh' })).json.code, 'disabled');
  } finally { t.done(); }
});

let failed = 0;
for (const [name, fn] of tests) {
  try { await fn(); console.log('ok   ' + name); } catch (e) { failed++; console.log('FAIL ' + name + '\n     ' + (e && e.stack ? e.stack.split('\n').slice(0, 4).join('\n     ') : e)); }
}
console.log(failed ? `${failed} of ${tests.length} failed` : `all ${tests.length} passed`);
process.exit(failed ? 1 : 0);
