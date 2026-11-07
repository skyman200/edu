// Checks the pure helpers behind gas.html and the generated recipe bundle: `node gas/diag.test.mjs`
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { RECIPES, buildRecipes } from './build.mjs';
import { validateJev } from '../functions/api/jev.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const D = createRequire(import.meta.url)('../gas-diag.js');
const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('parseError: Korean popup, stack frames, English, plain', () => {
  const a = D.parseError("TypeError: Cannot read properties of null (reading 'getRange') (줄 12, 파일 \"코드\")");
  assert.deepEqual([a.message, a.lines], ["TypeError: Cannot read properties of null (reading 'getRange')", [12]]);
  const b = D.parseError('Exception: 초과\n    at setup(코드:42:7)\n    at onOpen(코드:5:3)');
  assert.deepEqual([b.lines, b.fn], [[5, 42], 'setup']);
  assert.deepEqual(D.parseError('ReferenceError: x is not defined (line 7, file "Code")').lines, [7]);
  assert.deepEqual(D.parseError('그냥 글').lines, []);
  assert.deepEqual(D.parseError('Error (줄 99999999)').lines, []);
});

test('signature: same kind matches, different does not', () => {
  const s1 = D.signature("TypeError: Cannot read properties of null (reading 'getRange') (줄 12)");
  assert.equal(s1, D.signature("TypeError: Cannot read properties of null (reading 'appendRow') (줄 88)"));
  assert.notEqual(s1, D.signature('ReferenceError: foo is not defined (줄 3)'));
});

test('diagnose: common failures, specific beats generic', () => {
  const ids = (t) => D.diagnose(t).map((d) => d.id);
  assert.deepEqual(ids("TypeError: Cannot read properties of null (reading 'getRange')"), ['noSheet']);
  assert.deepEqual(ids("TypeError: Cannot read properties of undefined (reading 'name')"), ['readNull']);
  assert.ok(ids('Exception: Authorization is required to perform that action.').includes('perm'));
  assert.deepEqual(ids('Exception: Exceeded maximum execution time'), ['time']);
  assert.ok(ids('Exception: Service invoked too many times for one day: email.').includes('mailQuota'));
  assert.ok(ids('Exception: Service invoked too many times for one day: urlfetch.').includes('fetchQuota'));
  assert.deepEqual(ids('ReferenceError: jev_ is not defined'), ['notDefinedHelper']);
  assert.deepEqual(ids('ReferenceError: totl is not defined'), ['notDefined']);
  assert.deepEqual(ids('TypeError: sh.getRnage is not a function'), ['notFn']);
  assert.ok(ids('Error: 수업 코드가 아직 없습니다. 시트 메뉴에서 "수업 코드 입력"을 먼저 실행하세요.').includes('classCodeMissing'));
  assert.ok(ids('Error: AI ONE 429: AI 호출 한도에 도달했습니다.').includes('aiRate'));
  assert.ok(ids("Error: AI ONE 400: Jev가 질문 형식을 받아들이지 못했습니다. (questions['kind'].criteria must be an object)").includes('jevFormat'));
  assert.deepEqual(ids('SyntaxError: Unexpected end of input'), ['syntax']);
  assert.deepEqual(ids('SyntaxError: Unexpected token < in JSON at position 0'), ['jsonParse']);
  assert.ok(ids('Exception: Cannot call SpreadsheetApp.getUi() from this context.').includes('getUi'));
  assert.ok(ids('Script function not found: doGet').includes('noFn'));
  assert.deepEqual(ids('Error at doGet(코드:5:3) something odd'), []);
  assert.deepEqual(ids('totally unrelated text'), []);
  for (const r of D.RULES) assert.ok(r.title && r.cause && /[가-힣]/.test(r.fix), r.id);
});

test('excerpt and extractCode', () => {
  const code = Array.from({ length: 30 }, (_, i) => 'line' + (i + 1)).join('\n');
  assert.deepEqual(D.excerpt(code, [5, 20], 1).map((x) => x.n), [4, 5, 6, 0, 19, 20, 21]);
  assert.deepEqual(D.excerpt(code, [99], 2), []);
  const md = '설명\n```javascript\nfunction a() {}\n```\n```js\nfunction bigger() {\n  return 1;\n}\n```\n```html\n<div></div>\n```';
  assert.deepEqual(D.extractCode(md), { code: 'function bigger() {\n  return 1;\n}', truncated: false });
  assert.deepEqual(D.extractCode('```javascript\nfunction a() {\n  var x'), { code: 'function a() {\n  var x', truncated: true });
  assert.equal(D.extractCode('코드를 만들지 못했어요').code, '');
  assert.equal(D.extractCode('').code, '');
});

test('staticCheck: syntax, brackets, helpers, duplicates, keys', () => {
  assert.deepEqual(D.staticCheck('function a() { return 1; }'), []);
  assert.ok(D.staticCheck('').some((x) => x.level === 'error'));
  const r = D.staticCheck('function a() {\n  if (x) {\n    foo(');
  assert.ok(r.some((x) => x.level === 'error' && x.msg.startsWith('문법 오류')));
  assert.ok(r.some((x) => /닫히지 않았습니다/.test(x.msg) && x.line === 3));
  assert.ok(D.staticCheck('function a(){ ai_("x"); }').some((x) => /ai_ 함수를 쓰는데/.test(x.msg)));
  assert.ok(D.staticCheck('function a(){ ai_("x"); }\nfunction ai_(p){}').every((x) => !/ai_ 함수/.test(x.msg)));
  assert.ok(D.staticCheck('function a(){ return SERVER; }').some((x) => /SERVER/.test(x.msg)));
  assert.ok(D.staticCheck('function f() {}\nfunction g() {}\nfunction f() {}').some((x) => x.level === 'warn' && x.line === 3));
  assert.ok(D.staticCheck('var k = "AIzaSyA1234567890abcdefghijklmnopqrstuv";').some((x) => /API 키/.test(x.msg)));
  assert.deepEqual(D.staticCheck("var s = 'it\\'s'; var t = \"a'b\"; /* ( */ // [\nfunction f(){}"), []);
});

test('riskScan', () => {
  const code = 'function a(){ MailApp.sendEmail(t,"s","b"); MailApp.sendEmail(t,"s","b"); sh.clearContent(); sh.deleteRow(3);\n file.setSharing(x,y); ScriptApp.newTrigger("a"); UrlFetchApp.fetch("https://api.example.com/x"); DocumentApp.create("x"); }';
  const by = Object.fromEntries(D.riskScan(code).risks.map((x) => [x.key, x]));
  assert.equal(by.mail.count, 2);
  assert.equal(by.delete.count, 1);
  assert.equal(by.clear.level, 'danger');
  assert.equal(by.share.count, 1);
  assert.equal(by.fetch.count, 1);
  assert.deepEqual(D.riskScan(code).hosts, ['api.example.com']);
  assert.deepEqual(D.riskScan('function a(){ return 1 }'), { risks: [], hosts: [], scopes: [] });
});

test('lineDiff', () => {
  const a = 'a\nb\nc\nd\ne';
  assert.deepEqual(D.lineDiff(a, a), { added: 0, removed: 0, hunks: [] });
  const d = D.lineDiff(a, 'a\nb\nX\nY\nd\ne');
  assert.deepEqual([d.added, d.removed], [2, 1]);
  assert.deepEqual(d.hunks, [{ aStart: 3, bStart: 3, removed: ['c'], added: ['X', 'Y'] }]);
  assert.equal(D.lineDiff('1\n2\n3\n4\n5\n6', '1\nB\n3\n4\nE\n6').hunks.length, 2);
});

test('prompts', () => {
  const mk = D.buildMakePrompt({ request: '  퀴즈를 5문항으로  ', title: '칠판', hasBase: true });
  assert.ok(mk.includes('고쳐') && mk.includes('퀴즈를 5문항으로') && mk.includes('```javascript'));
  const ex = D.buildExternalPrompt({ request: '요청', title: 'T', code: 'function a(){}\n' });
  assert.ok(ex.includes('[기준 코드]\n```javascript\nfunction a(){}\n```') && !ex.includes('첨부한'));
  const fx = D.buildFixPrompt({ error: "TypeError: Cannot read properties of null (reading 'getRange') (줄 12)", notes: D.diagnose("reading 'getRange' of null"), attempt: 0 });
  assert.ok(fx.includes('[오류가 난 줄] 12') && !fx.includes('이미 한 번 고쳤는데'));
  assert.ok(D.buildFixPrompt({ error: 'x', attempt: 1 }).includes('이미 한 번 고쳤는데'));
  assert.ok(D.buildFixPrompt({ error: 'x'.repeat(5000), attempt: 0 }).length < 3500);
});

test('recipes: bundle up to date, Jev presets pass the real validator, sources really ask those questions, helpers + recipe pass staticCheck', () => {
  assert.equal(fs.readFileSync(path.join(here, '..', 'gas-recipes.js'), 'utf8'), buildRecipes(), 'run: node gas/build.mjs');
  assert.equal(RECIPES.length, 7);
  const common = fs.readFileSync(path.join(here, '00-common.gs'), 'utf8');
  for (const r of RECIPES) {
    const v = validateJev({ state: r.jevPreset.state, questions: r.jevPreset.questions });
    assert.equal(v.ok, true, `${r.id}: ${v.message}`);
    assert.ok(r.examples.length >= 3 && r.steps.length >= 4, r.id);
    if (!r.file) continue;
    const src = fs.readFileSync(path.join(here, r.file), 'utf8');
    for (const q of Object.keys(r.jevPreset.questions)) assert.ok(new RegExp('\\b' + q + ':\\s*\\{\\s*type:').test(src), `${r.id}: ${q}`);
    assert.deepEqual(D.staticCheck(common + '\n' + src).filter((x) => x.level === 'error'), [], r.id);
  }
});

test('gas.html: inline script parses, every $("#id") exists, scripts load in order', () => {
  const html = fs.readFileSync(path.join(here, '..', 'gas.html'), 'utf8');
  const scripts = [...html.matchAll(/<script(?: src="([^"]+)")?>([\s\S]*?)<\/script>/g)];
  assert.deepEqual(scripts.filter((s) => s[1]).map((s) => s[1]), ['gas-diag.js', 'gas-recipes.js']);
  const inline = scripts.find((s) => !s[1])[2];
  new vm.Script(inline, { filename: 'gas.html' });
  const ids = new Set([...html.matchAll(/\bid="([\w-]+)"/g)].map((m) => m[1]));
  for (const m of inline.matchAll(/\$\('#([\w-]+)'\)/g)) assert.ok(ids.has(m[1]), '#' + m[1]);
  for (const f of ['gas-diag.js', 'gas-recipes.js']) assert.ok(fs.existsSync(path.join(here, '..', f)), f);
});

let failed = 0;
for (const [name, fn] of tests) {
  try { await fn(); console.log(`ok   ${name}`); } catch (e) {
    failed++;
    console.log(`FAIL ${name}\n     ${e && e.stack ? e.stack.split('\n').slice(0, 5).join('\n     ') : e}`);
  }
}
console.log(failed ? `${failed} of ${tests.length} failed` : `all ${tests.length} passed`);
process.exitCode = failed ? 1 : 0;
