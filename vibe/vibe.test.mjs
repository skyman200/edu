// node vibe/vibe.test.mjs : 고칠 곳 끼워 넣기(VibePatch)와 예시 앱 9개가 약속을 지키는지 확인합니다.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { SAMPLES } from './build.mjs';

const P = createRequire(import.meta.url)('./patch.js');
const root = new URL('..', import.meta.url);
const read = (p) => readFileSync(new URL(p, root), 'utf8');
const blk = (a, b) => `<<<<<<< 찾기\n${a}\n=======\n${b}\n>>>>>>> 바꾸기`;

test('patch: unique match is replaced, text around the block is ignored', () => {
  const base = '<h1 style="font-size:20px">제목</h1>\n<p>끝</p>';
  const r = P.apply(base, '고쳤어요\n```\n' + blk('font-size:20px', 'font-size:40px') + '\n```');
  assert.equal(r.applied, 1); assert.deepEqual(r.failed, []);
  assert.equal(r.code, '<h1 style="font-size:40px">제목</h1>\n<p>끝</p>');
});
test('patch: indentation differences are tolerated and re-applied', () => {
  const base = 'function f() {\n    a();\n    b();\n}';
  const r = P.apply(base, blk('a();\nb();', 'a();\nc();\nb();'));
  assert.deepEqual(r.failed, []);
  assert.equal(r.code, 'function f() {\n    a();\n    c();\n    b();\n}');
});
test('patch: several blocks apply in order; any failure leaves the code untouched', () => {
  const base = 'one\ntwo\nthree';
  assert.equal(P.apply(base, blk('one', '1') + '\n' + blk('three', '3')).code, '1\ntwo\n3');
  const bad = P.apply(base, blk('one', '1') + '\n' + blk('nope', 'x'));
  assert.equal(bad.code, base); assert.equal(bad.failed.length, 1); assert.equal(bad.applied, 1);
});
test('patch: ambiguous or missing targets and answers without blocks are refused', () => {
  assert.match(P.apply('x\nx', blk('x', 'y')).failed[0].why, /여러 곳/);
  assert.match(P.apply('a', blk('b', 'c')).failed[0].why, /찾지 못/);
  assert.equal(P.apply('a', '그냥 설명').failed.length, 1);
  assert.equal(P.apply('a', '<<<<<<< 찾기\na\n=======\nb').failed.length, 1);
  assert.equal(P.has(blk('a', 'b')), true); assert.equal(P.has('<html></html>'), false);
});
test('prompts ask for changed spots only and carry the code', () => {
  const p = P.fixPrompt('크게', '<p>x</p>');
  assert.match(p, /전체 코드를 다시 쓰지 마/); assert.match(p, /<<<<<<< 찾기/); assert.ok(p.includes('<p>x</p>'));
  assert.ok(P.errorPrompt(['ReferenceError: a is not defined (줄 3)'], '<p/>').includes('ReferenceError'));
  assert.equal(P.fixPrompt('크게', '').includes('[지금 코드]'), false);
});

test('samples: each is one self-contained, flat, emoji-free page with an editable content block', () => {
  assert.equal(new Set(SAMPLES.map((s) => s.id)).size, SAMPLES.length);
  for (const s of SAMPLES) {
    const code = read(`vibe/${s.id}.html`);
    assert.ok(/^<!doctype html>/i.test(code), s.id);
    assert.ok(code.length < 12000, `${s.id} is too long to edit comfortably`);
    assert.ok(!/gradient|box-shadow|backdrop|blur\(|filter\s*:|color-mix/i.test(code), `${s.id} breaks the flat design`);
    assert.ok(!/https?:\/\/|<link|\bsrc=["']http/i.test(code), `${s.id} uses the network`);
    assert.ok(!/\b(alert|confirm|prompt)\(/.test(code), `${s.id} uses a browser dialog`);
    assert.ok(!/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(code), `${s.id} has emoji`);
    if (s.id !== 'order') assert.ok(code.includes('여기를 고치면 내용이 바뀝니다'), `${s.id} lacks the editable-content marker`);
    assert.ok(s.ask && s.desc && s.name);
  }
});
test('built files: studio.html is standalone and vibe-kit.js lists every sample', () => {
  const studio = read('studio.html'), kit = read('vibe-kit.js');
  assert.ok(!studio.includes('/*VIBEKIT*/'));
  assert.ok(!/<script[^>]*\ssrc=/i.test(studio), 'studio.html must not load other files');
  for (const s of SAMPLES) { assert.ok(studio.includes(s.name) && kit.includes(s.name), s.id); }
  assert.equal((studio.match(/<\/script>/g) || []).length, 2, 'sample code must not close the page script early');
});
