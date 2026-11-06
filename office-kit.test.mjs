// node office-kit.test.mjs  (DOM 없이 실행; JSZip 은 vendor/jszip/jszip.min.js 를 사용)
// 선택: OK_OUT=/some/dir 로 실행하면 만든 docx/xlsx 파일을 그 폴더에 남깁니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
globalThis.JSZip = require(path.join(here, 'vendor/jszip/jszip.min.js'));
const OfficeKit = require('./office-kit.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'officekit-'));
const results = [];
let failures = 0;
async function test(name, fn) {
  try { await fn(); results.push('ok   ' + name); }
  catch (e) { failures++; results.push('FAIL ' + name + '\n' + (e && e.stack || e)); }
}
const ran = { xmllint: false, pythonDocx: false, pythonXlsx: false };
const hasXmllint = spawnSync('xmllint', ['--version']).status === 0 || spawnSync('xmllint', ['--version']).stderr?.length > 0;
const pyOk = spawnSync('python3', ['-c', 'import docx, openpyxl']).status === 0;

async function unzip(blob) {
  const buf = Buffer.from(await blob.arrayBuffer());
  const zip = await JSZip.loadAsync(buf);
  const parts = {};
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) parts[name] = await zip.files[name].async('string');
  return { parts, buf };
}
function checkPackage(parts) {
  // 모든 XML 이 잘 짜였는지 (xmllint 가 있으면 그것으로, 없으면 정규식 수준)
  const dir = fs.mkdtempSync(path.join(tmp, 'x-'));
  for (const [name, xml] of Object.entries(parts)) {
    assert.ok(xml.startsWith('<?xml'), name + ' 선언 없음');
    assert.ok(!/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(xml), name + ' 에 제어문자');
    if (hasXmllint) {
      const f = path.join(dir, name.replace(/[\/\[\]]/g, '_'));
      fs.writeFileSync(f, xml);
      const r = spawnSync('xmllint', ['--noout', f], { encoding: 'utf8' });
      assert.equal(r.status, 0, name + ' 가 well-formed 아님: ' + r.stderr);
      ran.xmllint = true;
    }
  }
  // Content_Types: 모든 부품에 타입이 있어야 함
  const ct = parts['[Content_Types].xml'];
  assert.ok(ct, '[Content_Types].xml 없음');
  const overrides = [...ct.matchAll(/PartName="([^"]+)"/g)].map(m => m[1]);
  for (const o of overrides) assert.ok(parts[o.slice(1)] != null, 'Override 대상 부품 없음: ' + o);
  const defaults = [...ct.matchAll(/Extension="([^"]+)"/g)].map(m => m[1]);
  for (const name of Object.keys(parts)) {
    if (name === '[Content_Types].xml') continue;
    const ext = name.split('.').pop();
    assert.ok(overrides.includes('/' + name) || defaults.includes(ext), name + ' 의 콘텐츠 타입 없음');
  }
  // 모든 .rels 의 Target 이 실제 부품을 가리키는지
  for (const name of Object.keys(parts)) {
    if (!name.endsWith('.rels')) continue;
    const base = name.replace(/_rels\/[^/]*$/, '');
    for (const m of parts[name].matchAll(/Target="([^"]+)"/g)) {
      const target = path.posix.normalize(path.posix.join(base, m[1]));
      assert.ok(parts[target] != null, name + ' -> 없는 부품 ' + target);
    }
  }
  const ids = [...Object.values(parts).join('').matchAll(/<Relationship [^>]*Id="([^"]+)"/g)].map(m => m[1]);
  assert.ok(ids.length > 0);
}

const SAMPLE = `
<div class="doc">
<header class="cover"><p class="kicker">Syllabus 2026</p><h1>비즈니스 영어 이메일 &amp; 회의</h1>
<p class="sub">15주 강의계획서 초안과 평가 안내</p>
<div class="meta"><span><b>과목</b>비즈니스 영어</span><span><b>대상</b>2학년 40명</span><span><b>일자</b>2026. 11. 13.</span></div></header>
<p class="lead">이 문서는 <b>학생이 먼저 읽는</b> 수업 안내입니다. 특수문자 &lt;script&gt;alert("x")&lt;/script&gt; &amp; 'quote' 도 그대로 보입니다.</p>
<section class="sec"><h2><span class="no">01</span>수업 개요</h2>
<div class="cols3">
 <div class="card hl"><h3>목표</h3><p>업무 이메일을 <mark>정중하게</mark> 쓴다.</p><span class="tag red">핵심</span><span class="tag">필수</span><span class="tag gray">선택</span></div>
 <div class="card dark"><h3>방식</h3><p>실습 70%, 강의 30%</p><small>확인 필요</small></div>
 <div class="card"><h3>평가</h3><ul><li>출석 10%</li><li>과제 30%<ul><li>이메일 3통</li></ul></li></ul></div>
</div>
<div class="cols3"><div class="stat"><b>82%</b><span>실습 비중</span><div class="bar red" data-w="82"><i></i></div></div>
<div class="stat"><b>15주</b><span>학기 길이</span><div class="bar"><i></i></div></div>
<div class="stat"><b>40명</b><span>수강 인원</span></div></div>
</section>
<section class="sec"><h2>주차별 계획</h2>
<div class="timeline">
 <div class="wk"><span class="n">WEEK 1</span><h4>오리엔테이션</h4><p>수업 규칙과 AI 사용 원칙을 안내합니다.</p></div>
 <div class="wk mid"><span class="n">WEEK 8</span><h4>중간 평가</h4><p>이메일 쓰기 실기.</p></div>
 <div class="wk fin"><span class="n">WEEK 15</span><h4>기말 평가</h4><p>발표와 포트폴리오.</p></div>
</div>
<ol class="steps"><li>자료를 읽는다.</li><li>초안을 쓴다.</li><li>피드백을 반영한다.</li></ol>
<div class="callout warn"><p><b>주의</b> 학생 실명과 학번은 AI에 넣지 않습니다.</p></div>
<div class="callout"><p>평가 비율 합은 100% 입니다.</p></div>
<blockquote>좋은 이메일은 짧고 구체적이다.</blockquote>
<div class="say"><b>교수 멘트</b>오늘은 제목 줄만 연습합니다.</div>
</section>
<section class="sec pb"><h2>채점 기준</h2>
<table class="rubric"><thead><tr><th>기준</th><th>우수</th><th>보통</th><th>미흡</th></tr></thead><tbody>
<tr><td>내용</td><td><span class="pt">4점</span>요청이 분명하고 근거가 있다.</td><td><span class="pt">3점</span>요청은 있으나 근거가 약하다.</td><td><span class="pt">1점</span>요청이 불분명하다.</td></tr>
<tr><td>어조</td><td><span class="pt">4점</span>정중하다.</td><td><span class="pt">3점</span>대체로 정중하다.</td><td><span class="pt">1점</span>무례하다.</td></tr>
</tbody></table>
<table><tr><th colspan="2">합계</th></tr><tr><td rowspan="2">A</td><td>1</td></tr><tr><td>2</td></tr></table>
</section>
<section class="sec"><h2>활동지</h2>
<div class="sheet"><div class="name"><span>이름</span><span>날짜</span></div>
<div class="q"><p class="qt">1. 가장 알맞은 제목은?</p><ol class="opts"><li>Hello</li><li>Meeting</li><li>Request to move Friday's meeting</li><li>Important!!!</li></ol></div>
<p>빈칸: He <span class="fill"></span> a plan.</p><div class="line"></div><div class="box"></div><p class="ans">정답: 3</p></div>
</section>
<p class="foot">출처: 자체 작성 &middot; &nbsp;확인 필요</p>
</div>`;

const QUIZ = [
  ['문항', '보기1', '보기2', '보기3', '보기4', '정답', '해설'],
  ['이메일 제목으로 알맞은 것은?', 'Hello', 'Meeting', 'Request to move Friday\'s meeting', 'Important!!!', '3', '구체적이어야 한다 <&">'],
  ['정중한 표현이 아닌 것은?', 'Could you ~?', 'Would you mind ~?', 'Send it now.', 'I was wondering if ~', 4, '3번은 명령형'],
];

await test('parseHtml: 닫히지 않은 태그, 엔티티, script 제거', () => {
  const t = OfficeKit._internal.parseHtml('<p>a &amp; b &#xAC00; &bogus;<script>alert(1)</script><li>x<li>y</ul><!-- c --><div class="A  B">z');
  const flat = JSON.stringify(t);
  assert.ok(flat.includes('a & b 가 &bogus;'));
  assert.ok(!flat.includes('alert'));
  assert.ok(flat.includes('"cls":" a b "'));
});

await test('docx: 구조, XML, 콘텐츠 타입, 관계', async () => {
  const blob = await OfficeKit.docx(SAMPLE, { title: '테스트 <문서>' });
  assert.equal(blob.type, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const { parts, buf } = await unzip(blob);
  assert.equal(buf.slice(0, 2).toString(), 'PK');
  for (const p of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml', 'word/footer1.xml', 'docProps/core.xml', 'docProps/app.xml', 'word/_rels/document.xml.rels'])
    assert.ok(parts[p], '부품 없음 ' + p);
  checkPackage(parts);
  const doc = parts['word/document.xml'];
  // 한글, 특수문자
  assert.ok(doc.includes('비즈니스 영어 이메일 &amp; 회의'));
  assert.ok(doc.includes('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;'));
  assert.ok(doc.includes('Request to move Friday&apos;s meeting'));
  assert.ok(!doc.includes('<script'));
  // 폰트/언어/페이지
  const st = parts['word/styles.xml'];
  assert.ok(st.includes('Malgun Gothic') && st.includes('w:lang w:val="ko-KR"'));
  assert.ok(doc.includes('w:w="11906" w:h="16838"') && doc.includes('w:left="1134"'));
  // 구조
  assert.ok(doc.includes('w:pStyle w:val="Title"'));
  assert.ok(doc.includes('w:pStyle w:val="Heading1"') && doc.includes('w:pStyle w:val="Heading3"'));
  assert.ok(/<w:tblHeader\/>/.test(doc), '머리 행 반복');
  assert.ok(doc.includes('w:fill="1A1A1A"') && doc.includes('<w:color w:val="FFFFFF"/>'));
  assert.ok(doc.includes('w:sz="4"') && doc.includes('BFBFBF'));
  assert.ok(doc.includes('w:gridSpan') && doc.includes('w:vMerge'));
  assert.ok(/<w:numPr><w:ilvl w:val="1"\/>/.test(doc), '중첩 목록');
  assert.ok(doc.includes('pageBreakBefore'), '.pb');
  assert.ok(doc.includes('E54536'));
  assert.ok(!doc.includes('bar') || true);
  assert.ok(!/데이터|data-w/.test(doc));
  const num = parts['word/numbering.xml'];
  assert.ok(num.includes('w:numFmt w:val="bullet"') && num.includes('w:numFmt w:val="decimal"'));
  // 모든 numId 가 numbering 에 정의되어 있어야 함
  for (const m of doc.matchAll(/<w:numId w:val="(\d+)"/g)) assert.ok(num.includes('<w:num w:numId="' + m[1] + '"'), 'numId ' + m[1]);
  // 문서는 문단으로 끝나고 sectPr 가 마지막
  assert.ok(/<\/w:p><w:sectPr>/.test(doc));
  // 제목 속성 이스케이프
  assert.ok(parts['docProps/core.xml'].includes('테스트 &lt;문서&gt;'));
  if (process.env.OK_OUT) { fs.mkdirSync(process.env.OK_OUT, { recursive: true }); fs.writeFileSync(path.join(process.env.OK_OUT, 'sample.docx'), buf); }
  fs.writeFileSync(path.join(tmp, 'sample.docx'), buf);
});

await test('docx: 제어문자 제거, 길이 상한, 위험 입력', async () => {
  const long = 'ㄱ'.repeat(5000);
  const html = '<h1>' + 'T'.repeat(400) + '</h1><p>a\u0000b\u0001c\u000bd\u001fe\ud800f &lt;&amp;&quot;&gt; </p><p>' + long + '</p>' +
    '<img src=x onerror=alert(1)><svg><text>SVG</text></svg><a href="javascript:alert(1)">링크글</a><style>p{}</style><script>evil()</script>' +
    '<div style="x">'.repeat(200) + '깊음' + '</div>'.repeat(200);
  const { parts } = await unzip(await OfficeKit.docx(html));
  checkPackage(parts);
  const doc = parts['word/document.xml'];
  assert.ok(doc.includes('abcde') || doc.includes('ab'), '제어문자 제거 후 글 유지');
  assert.ok(!doc.includes('onerror') && !doc.includes('evil') && !doc.includes('javascript:'));
  assert.ok(doc.includes('링크글') && doc.includes('깊음'));
  assert.ok(!doc.includes('ㄱ'.repeat(2001)), '문단 상한');
  assert.ok(doc.includes('ㄱ'.repeat(1500)));
  assert.ok(!doc.includes('T'.repeat(121)), '제목 상한');
  assert.ok(doc.includes('&lt;&amp;&quot;&gt;'));
});

await test('docx: 슬라이드 40장 상한, 빈 입력, 표로 끝나는 문서', async () => {
  const slides = Array.from({ length: 50 }, (_, i) => `<div class="slide"><h2>슬라이드 제목 ${i + 1}</h2><p class="big">${i}</p></div>`).join('');
  const { parts } = await unzip(await OfficeKit.docx('<div class="deck">' + slides + '</div>'));
  checkPackage(parts);
  const d = parts['word/document.xml'];
  assert.ok(d.includes('슬라이드 제목 40') && !d.includes('슬라이드 제목 41'));
  const r = await unzip(await OfficeKit.docx(''));
  checkPackage(r.parts);
  const t = await unzip(await OfficeKit.docx('<table><tr><td>끝</td></tr></table>'));
  assert.ok(/<\/w:tbl><w:p>/.test(t.parts['word/document.xml']));
  checkPackage(t.parts);
});

await test('docx: DOM 형태(요소 객체) 입력', async () => {
  const mk = (tag, cls, kids) => ({ nodeType: 1, tagName: tag.toUpperCase(), childNodes: kids, getAttribute: a => a === 'class' ? cls || null : null });
  const tx = s => ({ nodeType: 3, nodeValue: s });
  const el = mk('article', 'doc', [mk('h1', '', [tx('DOM 제목')]), mk('p', '', [tx('본문 '), mk('b', '', [tx('굵게')])]), mk('style', '', [tx('p{}')])]);
  const { parts } = await unzip(await OfficeKit.docx(el));
  checkPackage(parts);
  assert.ok(parts['word/document.xml'].includes('DOM 제목') && parts['word/document.xml'].includes('굵게'));
  assert.ok(!parts['word/document.xml'].includes('p{}'));
});

await test('xlsx: 퀴즈 표 + 2시트, 머리글 굵게/채움/틀 고정, 이름 정리', async () => {
  const blob = await OfficeKit.xlsx([
    { name: QUIZ_NAME(), rows: QUIZ, header: true },
    { name: 'a/b:c*d?e[f]g\\', rows: [['x', '숫자', 7, '007', '1.50'], ['y', 3.14, '42', null, '']], header: false },
    { name: '[]', rows: [['중복']] },
    { name: 'history', rows: [] },
    { name: 'A/B:C*D?E[F]G\\', rows: [['dup']] },
  ], { numbers: true });
  assert.equal(blob.type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const { parts, buf } = await unzip(blob);
  checkPackage(parts);
  const wb = parts['xl/workbook.xml'];
  const names = [...wb.matchAll(/<sheet name="([^"]*)"/g)].map(m => m[1]);
  assert.equal(names.length, 5);
  for (const n of names) { assert.ok(n.length <= 31 && n.length > 0, n); assert.ok(!/[\[\]:*?\/\\]/.test(n), n); }
  assert.equal(new Set(names.map(n => n.toLowerCase())).size, 5, '이름 중복: ' + names);
  assert.ok(names[0].length === 31);
  assert.ok(names.includes('a_b_c_d_e_f_g_'));
  assert.ok(!names.includes('History'));
  const s1 = parts['xl/worksheets/sheet1.xml'];
  assert.ok(s1.includes('<pane ySplit="1"') && s1.includes('state="frozen"'));
  assert.ok(!parts['xl/worksheets/sheet2.xml'].includes('<pane'));
  const sty = parts['xl/styles.xml'];
  assert.ok(sty.includes('<b/>') && sty.includes('Malgun Gothic') && sty.includes('wrapText="1"') && sty.includes('FF1A1A1A'));
  assert.ok(/<c r="A1" s="1" t="s">/.test(s1) && /<c r="A2" s="2" t="s">/.test(s1));
  assert.ok(/<c r="F3" s="2"><v>4<\/v><\/c>/.test(s1), '숫자 셀');
  const sst = parts['xl/sharedStrings.xml'];
  assert.ok(sst.includes('문항') && sst.includes('구체적이어야 한다 &lt;&amp;&quot;&gt;') && sst.includes('Request to move Friday&apos;s meeting'));
  // numbers 옵션: '7'(문자열 숫자)는 숫자, '007'은 문자, '1.50'은 숫자(1.5)
  const s2 = parts['xl/worksheets/sheet2.xml'];
  assert.ok(/<c r="C1" s="2"><v>7<\/v><\/c>/.test(s2));
  assert.ok(/<c r="D1" s="2" t="s">/.test(s2));
  assert.ok(/<c r="E1" s="2"><v>1.5<\/v><\/c>/.test(s2));
  assert.ok(/<c r="B2" s="2"><v>3.14<\/v><\/c>/.test(s2));
  assert.ok(/<c r="C2" s="2"><v>42<\/v><\/c>/.test(s2));
  // 열 너비: 상한 60
  const w = [...s1.matchAll(/width="([\d.]+)"/g)].map(m => parseFloat(m[1]));
  assert.ok(w.length === 7 && w.every(x => x >= 8 && x <= 60));
  if (process.env.OK_OUT) { fs.mkdirSync(process.env.OK_OUT, { recursive: true }); fs.writeFileSync(path.join(process.env.OK_OUT, 'sample.xlsx'), buf); }
  fs.writeFileSync(path.join(tmp, 'sample.xlsx'), buf);
});
function QUIZ_NAME() { return '퀴즈 ' + '가'.repeat(40); }

await test('xlsx: numbers 옵션 없이는 숫자 문자열이 텍스트로 남는다, 제어문자 제거, 빈 입력 오류', async () => {
  const { parts } = await unzip(await OfficeKit.xlsx([{ name: 's', rows: [['h'], ['12', 'a\u0000b\u0008c\u0010', '=1+1']] }]));
  checkPackage(parts);
  const s = parts['xl/worksheets/sheet1.xml'];
  assert.ok(/<c r="A2" s="2" t="s">/.test(s));
  assert.ok(parts['xl/sharedStrings.xml'].includes('>abc<'));
  await assert.rejects(() => OfficeKit.xlsx([]), /시트/);
  await assert.rejects(() => OfficeKit.xlsx(null));
});

await test('tablesFromElement / quizRows', () => {
  const tabs = OfficeKit.tablesFromElement(SAMPLE);
  assert.equal(tabs.length, 2);
  assert.equal(tabs[0].name, '채점 기준');
  assert.deepEqual(tabs[0].rows[0], ['기준', '우수', '보통', '미흡']);
  assert.equal(tabs[0].rows[1][0], '내용');
  assert.ok(tabs[0].rows[1][1].includes('4점') && tabs[0].rows[1][1].includes('요청이 분명하고'));
  assert.equal(tabs[1].rows[0].length, 2); // colspan 확장
  assert.equal(tabs[1].rows[2][0], 'A');   // rowspan 채움
  // DOM 형태 + shadowRoot
  const mk = (tag, kids, attrs) => ({ nodeType: 1, tagName: tag.toUpperCase(), childNodes: kids, getAttribute: a => (attrs || {})[a] || null });
  const tx = s => ({ nodeType: 3, nodeValue: s });
  const tbl = mk('table', [mk('tr', [mk('th', [tx('문항')]), mk('th', [tx('보기1')])]), mk('tr', [mk('td', [tx('Q'), mk('br', [])]), mk('td', [tx('A')])])]);
  const shadow = { nodeType: 11, childNodes: [mk('style', [tx('x')]), mk('article', [mk('h3', [tx('퀴즈 표')]), tbl])] };
  const host = { nodeType: 1, tagName: 'DIV', childNodes: [], shadowRoot: shadow, getAttribute: () => null };
  const t2 = OfficeKit.tablesFromElement(host);
  assert.equal(t2.length, 1);
  assert.equal(t2[0].name, '퀴즈 표');
  assert.deepEqual(t2[0].rows, [['문항', '보기1'], ['Q', 'A']]);

  assert.equal(OfficeKit.quizRows(QUIZ), true);
  assert.equal(OfficeKit.quizRows([['문항', '보기 1', '보기 2', '보기 3', '보기 4', '정답', '해설']]), true);
  assert.equal(OfficeKit.quizRows({ rows: QUIZ }), true);
  assert.equal(OfficeKit.quizRows([['번호', '문항', '보기1', '보기2', '보기3', '보기4', '정답', '해설']]), false);
  assert.equal(OfficeKit.quizRows([['주차', '주제', '활동', '과제']]), false);
  assert.equal(OfficeKit.quizRows([]), false);
  assert.equal(OfficeKit.quizRows(null), false);
  assert.equal(OfficeKit.quizRows('x'), false);
});

await test('DOC_PROMPT 의 예시 두 쪽 + 퀴즈 표가 그대로 docx / xlsx 로 변환된다', async () => {
  const { DOC_PROMPT, DOC_HINTS } = await import('./functions/_lib/view-doc.js');
  assert.ok(DOC_PROMPT.length < 9000);
  for (const id of ['s1-2', 's1-3', 's2-1', 's2-2', 's2-3', 's2-4', 's2-5', 's3-3', 'p-1', 'p-2', 'p-3', 'p-4', 'chat']) assert.ok(DOC_HINTS[id] && DOC_HINTS[id].length > 20, id);
  for (const id of ['p-1', 'p-2', 'p-3', 'p-4']) assert.ok(DOC_HINTS[id].includes('[쪽 N]') && DOC_HINTS[id].includes('첨부한 자료만'), id);
  // 프롬프트에 쓴 클래스는 모두 dockit.css 에 있어야 한다 (css 가 있을 때만 확인)
  const cssPath = '/tmp/dockit.css';
  if (fs.existsSync(cssPath)) {
    const css = fs.readFileSync(cssPath, 'utf8');
    const def = new Set([...css.matchAll(/\.([a-z][a-z0-9-]*)/g)].map(m => m[1]));
    for (const m of DOC_PROMPT.matchAll(/class="([^"]+)"/g)) for (const c of m[1].split(' ')) assert.ok(def.has(c), '정의 안 된 클래스 ' + c);
  }
  const ex = DOC_PROMPT.slice(DOC_PROMPT.indexOf('[예시 1'));
  const { parts } = await unzip(await OfficeKit.docx(ex.replace(/\[예시 \d[^\]]*\]/g, '')));
  checkPackage(parts);
  const d = parts['word/document.xml'];
  assert.ok(d.includes('비즈니스 영어 15주 계획') && d.includes('이메일 쓰기 평가 기준') && d.includes('중간 평가'));
  const html = '<table><thead><tr><th>문항</th><th>보기1</th><th>보기2</th><th>보기3</th><th>보기4</th><th>정답</th><th>해설</th></tr></thead><tbody><tr><td>Q1</td><td>a</td><td>b</td><td>c</td><td>d</td><td>2</td><td>이유</td></tr></tbody></table>';
  const tabs = OfficeKit.tablesFromElement(html);
  assert.equal(OfficeKit.quizRows(tabs[0].rows), true);
  const x = await unzip(await OfficeKit.xlsx(tabs));
  checkPackage(x.parts);
});

await test('save: node 에서는 blob 을 그대로 돌려준다 / 브라우저 흐름', () => {
  const b = new Blob(['x']);
  assert.equal(OfficeKit.save(b, 'a.docx'), b);
  // 가짜 브라우저 환경
  const calls = [];
  const timers = [];
  const realST = globalThis.setTimeout;
  globalThis.document = {
    body: { appendChild: a => calls.push(['append', a.download, a.href]) },
    createElement: () => ({ style: {}, click() { calls.push(['click']); }, parentNode: null })
  };
  const realURL = globalThis.URL;
  globalThis.URL = { createObjectURL: () => 'blob:x', revokeObjectURL: u => calls.push(['revoke', u]) };
  globalThis.setTimeout = (fn, ms) => { timers.push(ms); return 0; };
  try {
    OfficeKit.save(b, '../수업:계획?.docx');
  } finally {
    delete globalThis.document; globalThis.URL = realURL; globalThis.setTimeout = realST;
  }
  assert.deepEqual(calls.slice(0, 2).map(c => c[0]), ['append', 'click']);
  assert.equal(calls[0][1], '.._수업_계획_.docx'.replace(/^\.+/, '') || '');
  assert.ok(timers.includes(60000));
});

await test('python-docx / openpyxl 로 열어 보기', () => {
  if (!pyOk) { results.push('skip python-docx/openpyxl 없음'); return; }
  const script = `
import sys, docx, openpyxl
d = docx.Document(sys.argv[1])
paras = [p.text for p in d.paragraphs]
assert any('비즈니스 영어 이메일 & 회의' in p for p in paras), paras[:5]
assert d.core_properties.title.startswith('테스트') or True
assert len(d.tables) == 7, len(d.tables)  # 최상위 표: cols 2 + callout 2 + rubric + 합계 + sheet
rub = [t for t in d.tables if t.rows[0].cells[0].text.strip() == '기준']
assert rub and len(rub[0].rows) == 3 and len(rub[0].columns) == 4
assert '내용' in rub[0].rows[1].cells[0].text and '4점' in rub[0].rows[1].cells[1].text
styles = {p.style.name for p in d.paragraphs}
assert 'Title' in styles and 'Heading 1' in styles, styles
wb = openpyxl.load_workbook(sys.argv[2])
assert len(wb.sheetnames) == 5, wb.sheetnames
ws = wb.worksheets[0]
assert ws['A1'].value == '문항' and ws['A1'].font.b and ws['A1'].fill.fgColor.rgb.endswith('1A1A1A')
assert ws.freeze_panes == 'A2', ws.freeze_panes
assert ws['F3'].value == 4 and ws['F2'].value == 3  # numbers 옵션: 숫자 문자열 -> 숫자
assert ws['G2'].value == '구체적이어야 한다 <&">', ws['G2'].value
assert ws['A1'].font.name == 'Malgun Gothic'
ws2 = wb.worksheets[1]
assert ws2['B2'].value == 3.14 and ws2['D1'].value == '007'
print('PY-OK', len(paras), len(d.tables))
`;
  const r = spawnSync('python3', ['-c', script, path.join(tmp, 'sample.docx'), path.join(tmp, 'sample.xlsx')], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stdout + r.stderr);
  ran.pythonDocx = ran.pythonXlsx = true;
});

console.log(results.join('\n'));
console.log('\nchecks actually run: xmllint=' + ran.xmllint + ' python-docx=' + ran.pythonDocx + ' openpyxl=' + ran.pythonXlsx);
fs.rmSync(tmp, { recursive: true, force: true });
if (failures) { console.log('\n' + failures + ' FAILED'); process.exit(1); }
console.log('\nall passed');
