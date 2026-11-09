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
  assert.deepEqual(ids('Script function not found: doGet'), ['noDoGet'], 'a missing doGet is explained as "no web page in this code", not as a paste error');
  assert.deepEqual(ids('칠판 사진이 수업 노트로 스크립트 함수(doGet)를 찾을 수 없습니다.'), ['noDoGet'], 'the Korean web-app error page is recognized');
  assert.ok(ids('Script function not found: setup').includes('noFn'));
  assert.ok(!ids('Script function not found: setup').includes('noDoGet'));
  assert.deepEqual(ids('Error at doGet(코드:5:3) something odd'), []);
  assert.deepEqual(ids('totally unrelated text'), []);
  for (const r of D.RULES) assert.ok(r.title && r.cause && /[가-힣]/.test(r.fix), r.id);
});

test('excerpt', () => {
  const code = Array.from({ length: 30 }, (_, i) => 'line' + (i + 1)).join('\n');
  assert.deepEqual(D.excerpt(code, [5, 20], 1).map((x) => x.n), [4, 5, 6, 0, 19, 20, 21]);
  assert.deepEqual(D.excerpt(code, [99], 2), []);
});

test('staticCheck: syntax, brackets, helpers, duplicates, keys', () => {
  assert.deepEqual(D.staticCheck('function a() { return 1; }'), []);
  assert.ok(D.staticCheck('').some((x) => x.level === 'error'));
  const r = D.staticCheck('function a() {\n  if (x) {\n    foo(');
  assert.ok(r.some((x) => x.level === 'error' && x.msg.startsWith('문법 오류')));
  assert.ok(r.some((x) => /닫히지 않았습니다/.test(x.msg) && x.line === 3));
  assert.ok(D.staticCheck('function a(){ ai_("x"); }').some((x) => /ai_ 함수를 쓰는데/.test(x.msg)));
  const web = 'function page_(t){return t}\nfunction onOpen(){}\nfunction show(){ return page_("a","b"); }';
  assert.ok(D.staticCheck(web).some((x) => /doGet 함수가 없습니다/.test(x.msg)), 'a page without doGet is flagged');
  assert.ok(!D.staticCheck(web + '\nfunction doGet(){ return show(); }').some((x) => /doGet/.test(x.msg)));
  assert.ok(!D.staticCheck('function page_(t){return t}\n// return page_("a")\nfunction x(){}').some((x) => /doGet/.test(x.msg)), 'comments do not count');
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
  assert.ok(mk.includes('<<<<<<< 찾기') && mk.includes('>>>>>>> 바꾸기'));
  assert.ok(mk.includes('바꾸거나 새로 만든 함수만') && mk.includes('// 삭제: 함수이름') && !mk.includes('수정한 전체 코드'), 'AI ONE gets function-level patches');
  const ex = D.buildExternalPrompt({ request: '요청', title: 'T', code: 'function a(){}\n' });
  assert.ok(ex.includes('[기준 코드]\n```javascript\nfunction a(){}\n```') && !ex.includes('첨부한'));
  assert.ok(ex.includes('수정한 전체 코드') && !ex.includes('바꾸거나 새로 만든 함수만'), 'external AIs still return the whole file');
  const fx = D.buildFixPrompt({ error: "TypeError: Cannot read properties of null (reading 'getRange') (줄 12)", notes: D.diagnose("reading 'getRange' of null"), attempt: 0 });
  assert.ok(fx.includes('[오류가 난 줄] 12') && !fx.includes('이미 한 번 고쳤는데'));
  assert.ok(fx.includes('고친 함수만') && !fx.includes('고친 전체 코드'));
  assert.ok(D.buildFixPrompt({ error: 'x', attempt: 1 }).includes('이미 한 번 고쳤는데'));
  assert.ok(D.buildFixPrompt({ error: 'x'.repeat(5000), attempt: 0 }).length < 3500);
});

const blk = (o, n, ind = '') => ind + '<<<<<<< 찾기\n' + o + '\n' + ind + '=======\n' + n + '\n' + ind + '>>>>>>> 바꾸기\n';
const unchanged = (base, patch) => { const r = D.applyPatch(base, patch); assert.ok(r.failed.length || r.incomplete, 'must be refused: ' + patch.slice(0, 60)); assert.equal(r.code, base, 'a refused patch leaves the code as it was'); return r; };

test('applyPatch: replaces, adds and deletes functions by name', () => {
  const base = [
    "// header", "var CFG = { a: 1 };", "const NAME = 'x';", "",
    "// 시작", "function start() {", "  var s = '}{';", "  return /[})]/.test(s) ? 1 : 2; // }", "}", "",
    "async function later(a = {}) {", "  return `${a}}`;", "}", "", "function gone() { return 0; }", "",
  ].join('\n');
  const r = D.applyPatch(base, ["function start() {", "  return 42;", "}", "const NAME = 'y';", "// 새 함수", "function extra_() { return 3; }", "// 삭제: gone", ""].join('\n'));
  assert.deepEqual([r.replaced, r.added, r.removed, r.failed, r.incomplete], [['start', 'NAME'], ['extra_'], ['gone'], [], false]);
  assert.ok(r.code.includes('// 시작\nfunction start() {\n  return 42;\n}\n'), 'the comment above a replaced function stays');
  assert.ok(r.code.includes("const NAME = 'y';") && !r.code.includes("'x'") && r.code.includes('var CFG = { a: 1 };'));
  assert.ok(r.code.includes('async function later(a = {})') && !r.code.includes('gone()'));
  assert.ok(r.code.trimEnd().endsWith('// 새 함수\nfunction extra_() { return 3; }'));
  new Function(r.code);
  assert.equal(D.applyPatch(base, '').count, 0);
  assert.equal(D.applyPatch(base, 'function broken() {\n  var x = {').incomplete, true);
  assert.equal(D.applyPatch('', 'function a() {}').code.trim(), 'function a() {}');
  assert.equal(D.applyPatch(base, 'function start() { return 1; }\nfunction start() { return 2; }').code.includes('return 2'), true, 'a later duplicate wins');
});

test('applyPatch: search/replace blocks edit a few lines inside a big function', () => {
  const base = 'function a() {\n  var x = 1;\n  if (x) {\n    y();\n  }\n}\nfunction b() { y(); }\n';
  const ok = D.applyPatch(base, blk('  var x = 1;', '  var x = 2;\n  z();'));
  assert.deepEqual([ok.edited, ok.failed, ok.incomplete], [1, [], false]);
  assert.equal(ok.code, 'function a() {\n  var x = 2;\n  z();\n  if (x) {\n    y();\n  }\n}\nfunction b() { y(); }\n');
  // the anchor has a different indentation than the base: found by trimmed lines, the new lines take the base indentation
  const loose = D.applyPatch('function f() {\n    if (a) {\n        go();\n    }\n}\n', blk('if (a) {\n    go();\n}', 'if (a) {\n    go2();\n}'));
  assert.equal(loose.failed.length, 0);
  assert.equal(loose.code, 'function f() {\n    if (a) {\n        go2();\n    }\n}\n');
  const both = D.applyPatch(base, blk('  var x = 1;', '  var x = 5;') + 'function b() { return 9; }\n');
  assert.deepEqual([both.edited, both.replaced], [1, ['b']]);
  assert.equal(D.applyPatch(base, blk('y();', 'q();')).failed[0].why.slice(0, 8), '같은 글이 여러');
  assert.equal(D.applyPatch(base, blk('nothing here', 'x')).failed.length, 1);
  // indented marker lines are still blocks
  assert.equal(D.applyPatch(base, blk('  var x = 1;', '  var x = 3;', '    ')).edited, 1);
});

test('applyPatch refuses anything it cannot place safely and then changes nothing', () => {
  const base = "var HTML = '<a>';\n\nfunction f() { return HTML; }\nfunction g() {\n  var url = webUrl_() || set();\n  return url;\n}\n";
  // an unfinished or tangled marker block
  unchanged(base, '<<<<<<< 찾기\nvar x\n=======\nvar x = 3');
  unchanged(base, blk('var url = webUrl_() || set();', 'NEW1').replace('>>>>>>> 바꾸기\n', '') + blk('return url;', 'NEW2'));
  // an anchor that is missing or ambiguous anywhere in the answer fails the whole answer, even if other parts are fine
  unchanged(base, blk('return url;', 'return 1;') + blk('not there', 'x') + 'function h_() {}\n');
  // loose text next to the code, a bare statement, a class
  unchanged(base, 'function h_() {}\nconsole.log(1);\n');
  unchanged(base, 'class A { m() {} }\n');
  // delete of a name that does not exist
  assert.match(unchanged(base, '// 삭제: nothing_\nfunction h_() {}\n').failed[0].why, /기준 코드에 없습니다/);
  // an unfinished function
  unchanged(base, 'function h_() {\n  var a = [');
  // a base whose braces do not match cannot take function patches, but marker blocks still work on it
  const broken = 'function a() {\n  if (x) {\n    y();\n}\nfunction b() { return 1; }\n';
  unchanged(broken, 'function c_() {}\n');
  assert.equal(D.applyPatch(broken, blk('    y();', '    z();')).edited, 1);
  // a multi-line var: a comment between the lines must not cut the declaration
  const r = D.applyPatch(base, "var HTML = '<c>' +\n  // note\n  '<d>';\n");
  assert.deepEqual(r.replaced, ['HTML']);
  assert.equal(r.code, "var HTML = '<c>' +\n  // note\n  '<d>';\n\nfunction f() { return HTML; }\nfunction g() {\n  var url = webUrl_() || set();\n  return url;\n}\n");
});

test('applyPatch: scanning survives nested templates, regexes and strings; new constants go before their first user; comments are kept clean', () => {
  const base = "function list_(xs) { return `<ul>${xs.map((x) => `<li>${x}</li>`).join('')}</ul>`; }\nfunction later_() { return 1; }\nvar ASK = '<b>' + 'x';\nfunction re_() { return /\\/\\*|[\"'`]/.test(s) ? 1 : a / b; }\nfunction tail_() { return 2; }\n";
  const u = D.scanUnits(base);
  assert.deepEqual(u.map((x) => x.name), ['list_', 'later_', 'ASK', 're_', 'tail_']);
  assert.ok(!u.some((x) => x.incomplete));
  const r = D.applyPatch(base, "function later_() { return 11; }\n// 지금은 안 쓰는 값\nconst NOTICE = 'n';\nvar ASK = '<main>' + NOTICE;\n/*\n  도우미\n*/\nfunction helper_() { return 5; }\n");
  assert.deepEqual([r.replaced, r.added.sort(), r.failed], [['later_', 'ASK'], ['NOTICE', 'helper_'], []]);
  assert.ok(r.code.indexOf("const NOTICE = 'n';") < r.code.indexOf("var ASK = '<main>' + NOTICE;"), 'the constant is defined before the declaration that reads it');
  assert.ok(r.code.includes('/*\n  도우미\n*/\nfunction helper_()'), 'a block comment above a new function is kept whole');
  new Function(r.code);
  // the delete directive inside a replaced function body is just a comment, not an instruction
  const sneaky = D.applyPatch(base, "function later_() {\n  // 제거: tail_ 호출은 뺐습니다\n  return 1;\n}\n");
  assert.deepEqual([sneaky.removed, sneaky.failed], [[], []]);
  assert.ok(sneaky.code.includes('function tail_()'));
});

test('applyPatch on the real recipes: every unit round-trips and a patched function splices cleanly', () => {
  const common = fs.readFileSync(path.join(here, '00-common.gs'), 'utf8');
  for (const f of fs.readdirSync(here).filter((x) => /^r\d.*\.gs$/.test(x))) {
    const base = common + '\n' + fs.readFileSync(path.join(here, f), 'utf8');
    const units = D.scanUnits(base);
    assert.equal(units.filter((u) => u.kind === 'function').length, (base.match(/^(?:async\s+)?function\s+\w+/gm) || []).length, f);
    assert.equal(units.filter((u) => u.kind === 'var').length, (base.match(/^(?:const|let|var)\s+\w+/gm) || []).length, f);
    assert.ok(!units.some((u) => u.incomplete), f);
    for (const u of units) {
      const text = base.slice(u.start, u.end);
      if (u.kind === 'function') new Function(text);
      const same = D.applyPatch(base, text);
      assert.deepEqual([same.replaced, same.failed, same.code], [[u.name], [], base], f + ' ' + u.name + ': replacing a unit by itself changes nothing');
    }
    const t = units.filter((u) => u.kind === 'function')[3];
    const mine = base.slice(t.start, t.end).replace('{', '{ var patched_ = 1;');
    const r = D.applyPatch(base, D.extractPatch('```javascript\n' + mine + '\n```').code);
    assert.deepEqual([r.replaced, r.added, r.removed], [[t.name], [], []], f);
    const d = D.lineDiff(base, r.code);
    assert.ok(d.added <= 2 && d.removed <= 2, f + ' only the patched function changes');
    new Function(r.code);
  }
});

test('extractPatch: joins every code block, takes marker blocks in any fence, flags an unclosed one', () => {
  assert.deepEqual(D.extractPatch('설명\n```javascript\nfunction a() {}\n```\n그리고\n```js\nfunction b() {}\n```'), { code: 'function a() {}\n\nfunction b() {}', truncated: false });
  assert.equal(D.extractPatch('```javascript\nfunction a() {}\n```\n```javascript\nfunction b() {').truncated, true);
  assert.equal(D.extractPatch('```json\n{"a":1}\n```').code, '');
  assert.equal(D.extractPatch('').code, '');
  assert.equal(D.extractPatch('```javascript\nfunction n_() {}\n```\n```diff\n' + blk('a', 'b') + '```').code.includes('<<<<<<< 찾기'), true, 'a marker block in a diff fence is not dropped');
  assert.equal(D.extractPatch('```html\n<div></div>\n```').code, '');
});

test('recipes: bundle up to date, Jev presets pass the real validator, sources really ask those questions, helpers + recipe pass staticCheck', () => {
  assert.equal(fs.readFileSync(path.join(here, '..', 'gas-recipes.js'), 'utf8'), buildRecipes(), 'run: node gas/build.mjs');
  assert.equal(RECIPES.length, 7);
  const common = fs.readFileSync(path.join(here, '00-common.gs'), 'utf8');
  for (const r of RECIPES) {
    const v = validateJev({ state: r.jevPreset.state, questions: r.jevPreset.questions });
    assert.equal(v.ok, true, `${r.id}: ${v.message}`);
    assert.ok(r.examples.length >= 3 && r.steps.length >= 4 && r.web && typeof r.web.deploy === 'boolean' && Array.isArray(r.web.links), r.id);
    if (r.web.deploy) assert.ok(r.web.links.length >= 1, r.id);
    if (!r.file) continue;
    const src = fs.readFileSync(path.join(here, r.file), 'utf8');
    for (const q of Object.keys(r.jevPreset.questions)) assert.ok(new RegExp('\\b' + q + ':\\s*\\{\\s*type:').test(src), `${r.id}: ${q}`);
    assert.deepEqual(D.staticCheck(common + '\n' + src).filter((x) => x.level === 'error'), [], r.id);
  }
});

test('gas.html: inline script parses, every $("#id") exists, scripts load in order', () => {
  const html = fs.readFileSync(path.join(here, '..', 'gas.html'), 'utf8');
  const scripts = [...html.matchAll(/<script(?: src="([^"]+)")?>([\s\S]*?)<\/script>/g)];
  assert.deepEqual(scripts.filter((s) => s[1]).map((s) => s[1]), ['vendor/qrcode/qrcode.js', 'gas-diag.js', 'gas-recipes.js']);
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
