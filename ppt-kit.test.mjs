// node ppt-kit.test.mjs  (DOM 없이 실행; JSZip 은 vendor/jszip, PptxGenJS 는 vendor/pptxgenjs 를 씁니다)
// 선택: PK_OUT=/some/dir 로 실행하면 만든 .pptx 를 그 폴더에 남깁니다. PK_REAL=/path/a.pptx 로 실제 덱을 지정합니다.
// 실제 덱(읽기만 함)은 /Volumes/kim/2026, ~/Downloads, ~/Desktop 에서 찾습니다. 원본은 절대 고치지 않습니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
globalThis.JSZip = require(path.join(here, 'vendor/jszip/jszip.min.js'));
process.removeAllListeners('warning');
const bundle = fs.readFileSync(path.join(here, 'vendor/pptxgenjs/pptxgen.bundle.js'), 'utf8');
const loader = vm.constants && vm.constants.USE_MAIN_CONTEXT_DEFAULT_LOADER;
new vm.Script(bundle, loader ? { filename: 'pptxgen.bundle.js', importModuleDynamically: loader } : { filename: 'pptxgen.bundle.js' }).runInThisContext();
globalThis.JSZip = require(path.join(here, 'vendor/jszip/jszip.min.js'));
const PptKit = require('./ppt-kit.js');
const { layoutSlide, measure, MORPH_XML, PALETTES, countAccent, parseXml } = PptKit._internal;

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pptkit-'));
const OUT = process.env.PK_OUT || '';
const results = [];
let failures = 0;
async function test(name, fn) {
  try { await fn(); results.push('ok   ' + name); }
  catch (e) { failures++; results.push('FAIL ' + name + '\n' + (e && e.stack || e)); }
}
const ran = { xmllint: false, pythonPptx: false, realDecks: [] };
const hasXmllint = spawnSync('xmllint', ['--version']).status === 0;
const hasPyPptx = spawnSync('python3', ['-c', 'import pptx']).status === 0;
const sha = b => crypto.createHash('sha256').update(b).digest('hex');
const save = (name, buf) => { if (OUT) fs.writeFileSync(path.join(OUT, name), buf); };

async function unzip(buf) {
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
  const raw = {}, text = {};
  for (const n of names) { raw[n] = await zip.files[n].async('uint8array'); if (/\.(xml|rels)$/.test(n)) text[n] = Buffer.from(raw[n]).toString('utf8'); }
  return { names, raw, text };
}
function xmlOk(parts, label) {
  const dir = fs.mkdtempSync(path.join(tmp, 'x-')), files = [];
  Object.entries(parts).forEach(([n, x], i) => {
    assert.ok(!/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(x), label + ' ' + n + ' 에 제어문자');
    const f = path.join(dir, i + '.xml'); fs.writeFileSync(f, x); files.push(f);
  });
  if (hasXmllint) {
    const r = spawnSync('xmllint', ['--noout', ...files], { encoding: 'utf8' });
    assert.equal(r.status, 0, label + ' xmllint: ' + r.stderr.slice(0, 800));
    ran.xmllint = true;
  } else Object.entries(parts).forEach(([n, x]) => assert.ok(!parseXml(x).broken, label + ' ' + n + ' 닫히지 않은 태그'));
}
function pyRead(file) {
  if (!hasPyPptx) return null;
  const r = spawnSync('python3', ['-c', 'import sys,json\nfrom pptx import Presentation\np=Presentation(sys.argv[1])\nout=[]\ndef walk(shs):\n  for sh in shs:\n    if sh.shape_type==6:\n      yield from walk(sh.shapes)\n    else:\n      yield sh\nfor s in p.slides:\n  t=[sh.text_frame.text for sh in walk(s.shapes) if sh.has_text_frame and sh.text_frame.text.strip()]\n  n=s.notes_slide.notes_text_frame.text if s.has_notes_slide and s.notes_slide.notes_text_frame is not None else ""\n  out.append({"t":t,"n":n})\nprint(json.dumps(out,ensure_ascii=False))', file], { encoding: 'utf8' });
  assert.equal(r.status, 0, 'python-pptx 가 열지 못함: ' + r.stderr.slice(-600));
  ran.pythonPptx = true;
  return JSON.parse(r.stdout);
}
const ab = blob => blob.arrayBuffer();

const NOTE = '이 장에서는 핵심 하나만 말씀드립니다. 학생이 무엇을 할 수 있게 되는지가 먼저입니다. 예시를 하나 들고 넘어갑니다.';
const DECK = {
  title: 'AI로 수업 설계하기', subtitle: '교수자를 위한 실습 워크숍', author: '김강훈', date: '2026.11.13', theme: 'brutal', morph: true,
  slides: [
    { layout: 'cover', kicker: 'AI CLASS LAB', title: 'AI로 수업을\n다시 설계하기', body: '부산외대 교수자 워크숍. 실습 중심 3시간 과정입니다.', notes: NOTE },
    { layout: 'section', kicker: '01', title: '왜 지금 수업 설계인가', body: '학생의 도구가 바뀌면 과제와 평가도 함께 바뀌어야 합니다.', notes: NOTE },
    { layout: 'statement', title: '도구보다 먼저\n학습목표를 정합니다', accent: '학습목표', body: '목표가 분명하면 AI를 어디에 쓸지 저절로 정해집니다.', notes: NOTE },
    { layout: 'text', kicker: 'BACKGROUND', title: '학생은 이미 AI를 쓰고 있습니다', body: '과제 초안, 요약, 번역에 AI를 쓰는 학생이 늘었습니다.\n금지보다 쓰는 방법을 가르치는 편이 효과적입니다.', notes: NOTE },
    { layout: 'cards', title: '오늘 만드는 세 가지', items: [{ head: '강의계획서', text: '15주 흐름과 평가 비율을 한 장에 정리합니다.' }, { head: '활동지', text: '수업 중에 바로 쓰는 A4 한 장 활동지입니다.' }, { head: '루브릭', text: '기준 4개와 수준 4단계로 채점을 맞춥니다.' }], notes: NOTE },
    { layout: 'stat', title: '과제에 AI를 써 본 학생', stat: { value: '82%', label: '설문 응답자 중 한 번 이상 사용' }, body: '수치는 예시입니다. 우리 학교 자료로 바꿔 쓰세요.', notes: NOTE },
    { layout: 'steps', title: '백워드 설계 4단계', items: [{ head: '목표', text: '학생이 할 수 있게 될 일을 동사로 씁니다.' }, { head: '증거', text: '목표를 이뤘다는 증거를 정합니다.' }, { head: '활동', text: '증거를 만드는 수업 활동을 짭니다.' }, { head: '점검', text: 'AI 사용 원칙을 함께 정합니다.' }], notes: NOTE },
    { layout: 'timeline', title: '오늘의 흐름', items: [{ num: '10:00', head: '열기', text: '오늘 목표 확인' }, { num: '10:30', head: '설계', text: '강의계획서 만들기' }, { num: '11:30', head: '실습', text: '활동지와 루브릭' }, { num: '12:30', head: '정리', text: '내 수업에 적용' }], notes: NOTE },
    { layout: 'table', title: '평가 비율과 근거', table: { head: ['항목', '비율', '근거'], rows: [['과제', '40%', '매주 짧은 글쓰기로 꾸준함을 봅니다'], ['중간', '20%', '8주 실기 평가'], ['기말', '30%', '포트폴리오 발표'], ['출석', '10%', '학칙 기준']] }, notes: NOTE },
    { layout: 'compare', title: 'AI 사용, 어디까지 허용할까', left: { head: '허용', items: ['아이디어 내기와 개요 잡기', '맞춤법과 문장 다듬기'] }, right: { head: '금지', items: ['답안 전체를 그대로 제출', '출처 없는 인용'] }, notes: NOTE },
    { layout: 'quote', title: '기억할 한 문장', quote: { text: '좋은 질문이 좋은 답을 만듭니다', by: '오늘 워크숍 정리' }, accent: '좋은 질문', notes: NOTE },
    { layout: 'closing', title: '내 수업에\n오늘 바로 써 봅니다', accent: '오늘', body: '만든 자료는 "내 결과물"에 모두 저장됩니다.', notes: NOTE }
  ]
};

await test('THEMES: 네 가지 테마와 label/desc', () => {
  assert.deepEqual(Object.keys(PptKit.THEMES), ['brutal', 'clean', 'navy', 'warm']);
  for (const t of Object.values(PptKit.THEMES)) { assert.ok(t.label && t.desc); }
  assert.equal(PptKit.THEMES.brutal.colors.accent, '#E54536');
  assert.equal(PptKit.THEMES.brutal.colors.bg, '#F2F2F2');
});

await test('parseModelJson: 펜스 / 앞뒤 설명 / 끝 쉼표 / 잘린 답', () => {
  const P = PptKit.parseModelJson;
  assert.deepEqual(P('```json\n{"a":1,"b":[1,2,],}\n```'), { a: 1, b: [1, 2] });
  assert.deepEqual(P('네, 만들었습니다 {이건 아님} 아래입니다.\n{"title":"x","slides":[]} 끝.'), { title: 'x', slides: [] });
  assert.deepEqual(P('{"t":"줄\n바꿈"}'), { t: '줄\n바꿈' });
  const tr = P('{"title":"덱","slides":[{"layout":"cover","title":"하나"},{"layout":"text","title":"둘","body":"끝까지 안 쓴');
  assert.equal(tr.__truncated, true); assert.equal(tr.slides.length, 1); assert.equal(tr.slides[0].title, '하나');
  const tr2 = P('{"title":"덱","slides":[{"layout":"cover","title":"하나"},{"layout":"text","ti');
  assert.equal(tr2.slides.length, 1);
  const tr3 = P('{"summary":"요약","edits":[{"slide":1,"shape":"2","text":["a","b"]},{"slide":2,"shape":"3","text":["c"');
  assert.equal(tr3.edits.length, 1); assert.equal(tr3.__truncated, true);
  const tr4 = P('{"title":"덱 제목","slides":[');
  assert.deepEqual(tr4.slides, []); assert.equal(tr4.title, '덱 제목');
  assert.equal(P('no json here'), null); assert.equal(P(''), null); assert.equal(P(null), null); assert.equal(P('{"a":'), null);
  assert.equal(P('[1,2]'), null);
});

await test('validateDeck: 모르는 구성, 빠진 값, 상한, 제어문자, 경고', () => {
  const many = { title: 'x'.repeat(300), theme: 'nope', slides: Array.from({ length: 45 }, (_, i) => ({ layout: i % 2 ? 'mystery' : 'text', title: '제목\u0001' + i, body: 'b'.repeat(5000), notes: 'n'.repeat(9000) })) };
  const { spec, warnings } = PptKit.validateDeck(many);
  assert.equal(spec.slides.length, 40); assert.equal(spec.theme, 'brutal'); assert.ok(spec.title.length <= 120);
  assert.equal(spec.slides[1].layout, 'text'); assert.ok(!spec.slides[0].title.includes('\u0001'));
  assert.ok(spec.slides[0].notes.length <= 6000); assert.ok(spec.slides[0].body.length <= 4000);
  assert.ok(warnings.some(w => /mystery/.test(w))); assert.ok(warnings.some(w => /같은 구성/.test(w)));
  const empty = PptKit.validateDeck(null);
  assert.equal(empty.spec.slides.length, 1); assert.equal(empty.spec.slides[0].layout, 'cover');
  const fb = PptKit.validateDeck({ slides: [{ layout: 'stat', title: '숫자 없음' }, { layout: 'cards', title: '항목 없음' }] }).spec;
  assert.deepEqual(fb.slides.map(s => s.layout), ['text', 'text']);
  const again = PptKit.validateDeck(PptKit.validateDeck(DECK).spec).spec;
  assert.deepEqual(again, PptKit.validateDeck(DECK).spec, '두 번 다듬어도 같아야 함');
  const reds = PptKit.validateDeck(DECK);
  assert.ok(!reds.warnings.some(w => /빨간|같은 구성|줄였/.test(w)), reds.warnings.join(' / '));
});

await test('previewHtml: 스크립트 없음, 주입된 < 는 이스케이프, cqw', () => {
  const evil = { title: '<script>alert(1)</script>', slides: [{ layout: 'text', title: '<img src=x onerror=alert(1)>', body: '"><svg onload=x> & </div>' }, ...DECK.slides] };
  const h = PptKit.previewHtml(evil);
  assert.ok(h.startsWith('<div class="ppt-prev"'));
  assert.ok(!/<script/i.test(h) && !/<img/i.test(h) && !/<svg/i.test(h));
  const tags = h.match(/<[^>]*>/g); for (const t of tags) assert.match(t, /^<\/?(div|span|br)\b/, t);
  assert.ok(h.includes('container-type:inline-size') && h.includes('cqw'));
  assert.ok(h.includes('&lt;img src=x onerror=alert(1)&gt;'));
});

let deckBuf = null;
await test('buildDeck: 12장 brutal 덱, 모든 구성, XML 정상, 노트, 이름, 모핑 위치', async () => {
  deckBuf = Buffer.from(await ab(await PptKit.buildDeck(DECK)));
  save('deck-brutal.pptx', deckBuf); fs.writeFileSync(path.join(tmp, 'deck.pptx'), deckBuf);
  const { text } = await unzip(deckBuf);
  xmlOk(text, 'deck');
  const info = await PptKit.inspect(deckBuf);
  assert.equal(info.slideCount, 12); assert.equal(info.size.w, 13.333);
  DECK.slides.forEach((s, i) => {
    const sl = info.slides[i], all = sl.shapes.map(x => x.text).join('\n');
    if (s.title && s.layout !== 'quote') assert.equal(sl.title, s.title.replace(/\n/g, ' '), 'slide ' + (i + 1));
    assert.equal(sl.notes, NOTE);
    (s.items || []).forEach(it => assert.ok(all.includes(it.head), it.head));
    if (s.stat) assert.ok(all.includes('82%'));
    if (s.table) assert.ok(all.includes('포트폴리오 발표'));
  });
  for (let k = 1; k <= 12; k++) {
    const x = text['ppt/slides/slide' + k + '.xml'];
    assert.ok(x.includes('name="!!title"') || k === 11, 'slide ' + k + ' !!title');
    if (k === 1) { assert.ok(!x.includes('p:transition')); continue; }
    const at = x.indexOf(MORPH_XML);
    assert.ok(at > x.indexOf('</p:clrMapOvr>') && x.indexOf('</p:clrMapOvr>') > 0, 'morph 는 clrMapOvr 뒤');
    const tm = x.indexOf('<p:timing'); assert.ok(tm < 0 || at < tm, 'morph 는 timing 앞');
    assert.equal(x.split('<p:transition').length - 1, 2);
    assert.match(x.slice(0, x.indexOf('>', x.indexOf('<p:sld'))), /xmlns:p14="http:\/\/schemas.microsoft.com\/office\/powerpoint\/2010\/main"/);
  }
  assert.ok(text['ppt/presentation.xml'].includes('cx="12192000" cy="6858000"'));
  const py = pyRead(path.join(tmp, 'deck.pptx'));
  if (py) { assert.equal(py.length, 12); assert.ok(py[5].t.includes('82%')); }
});

await test('글자 맞춤: 모든 글 상자가 어림 높이 안, 11pt 이상, 장 안에 있음 (긴 한글 덱 포함)', () => {
  const long = '가나다라마바사아자차카타파하'.repeat(10);
  const stress = { theme: 'brutal', slides: PptKit.LAYOUTS.map(l => ({ layout: l, kicker: 'KICKER LABEL', title: long.slice(0, 120), body: long.repeat(3), accent: '마바', items: Array.from({ length: 8 }, (_, i) => ({ num: '20' + (10 + i), head: long.slice(0, 80), text: long.repeat(2).slice(0, 600) })), stat: { value: '12,345명', label: long.slice(0, 100) }, table: { head: ['가', '나', '다', '라', '마', '바'], rows: Array.from({ length: 8 }, () => Array(6).fill(long.slice(0, 200))) }, left: { head: long.slice(0, 60), items: Array(6).fill(long.slice(0, 300)) }, right: { head: '오른쪽', items: Array(6).fill(long.slice(0, 300)) }, quote: { text: long.slice(0, 300), by: '출처' } })) };
  for (const deck of [DECK, stress, { ...DECK, theme: 'clean' }, { ...DECK, theme: 'navy' }, { ...DECK, theme: 'warm' }]) {
    const spec = PptKit.validateDeck(deck).spec;
    spec.slides.forEach((s, i) => {
      const L = layoutSlide(spec, i);
      for (const p of L.prims) {
        if (p.t !== 'text') continue;
        assert.ok(p.size >= 11, s.layout + ' ' + p.role + ' size ' + p.size);
        const h = measure(p.paras, p.size, p.w, p).h;
        assert.ok(h <= p.h + 1e-6, s.layout + ' ' + p.role + ' 넘침 ' + h.toFixed(2) + ' > ' + p.h.toFixed(2));
        if (p.role !== 'ghost' && !p.rotate) assert.ok(p.x >= 0 && p.y >= 0 && p.x + p.w <= 13.34 && p.y + p.h <= 7.5, s.layout + ' ' + p.role + ' 장 밖');
      }
      if (deck === DECK) {
        assert.ok(countAccent(L, PALETTES.brutal.accent) <= 2, s.layout + ' 빨강 ' + countAccent(L, PALETTES.brutal.accent));
        const t = L.prims.find(p => p.role === 'title' || p.role === 'quote');
        assert.ok(t && !t.trunc, s.layout + ' 제목이 잘림');
      }
    });
  }
});

await test('buildDeck: 적대적 글 (<&"> 와 제어문자) 이스케이프', async () => {
  const buf = Buffer.from(await ab(await PptKit.buildDeck({ title: 'a<b', slides: [{ layout: 'text', title: '<p:sld>&"\u0001\u0008', body: '</a:t><a:t>x\u000B', notes: '노트 <&>\u001F' }] }, { theme: 'navy', morph: false })));
  const { text } = await unzip(buf);
  xmlOk(text, 'evil');
  const s1 = text['ppt/slides/slide1.xml'];
  assert.ok(s1.includes('&lt;p:sld&gt;&amp;&quot;') && s1.includes('&lt;/a:t&gt;&lt;a:t&gt;x'));
  assert.ok(!s1.includes('p:transition'));
  const info = await PptKit.inspect(buf);
  assert.equal(info.slides[0].title, '<p:sld>&"'); assert.equal(info.slides[0].notes, '노트 <&>');
});

// 손으로 만든 덱: 여러 글 조각·여러 문단, 표, 묶음, 슬라이드 순서 바꿈, 노트 없는 장
async function handDeck({ dropNotesMaster = false } = {}) {
  const p = new PptxGenJS(); p.layout = 'LAYOUT_WIDE';
  const a = p.addSlide();
  a.addText([{ text: '굵은 빨강 ', options: { bold: true, color: 'FF0000' } }, { text: '보통 글', options: { breakLine: true } }, { text: '둘째 문단 기울임', options: { italic: true } }], { x: 1, y: 1, w: 6, h: 2, fontSize: 20, objectName: 'Multi' });
  a.addTable([[{ text: '머리1' }, { text: '머리2' }], [{ text: '칸1' }, { text: '칸2' }]], { x: 1, y: 4, w: 6 });
  a.addNotes('첫 장 원래 노트');
  const b = p.addSlide(); b.addText('두 번째 장', { x: 1, y: 1, w: 5, h: 1 });
  const zip = await JSZip.loadAsync(await p.write({ outputType: 'arraybuffer' }));
  let s1 = await zip.file('ppt/slides/slide1.xml').async('string');
  const i0 = s1.indexOf('<p:sp>'), i1 = s1.indexOf('</p:sp>', i0) + 7;
  s1 = s1.slice(0, i0) + '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="100" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + s1.slice(i0, i1) + '</p:grpSp>' + s1.slice(i1);
  zip.file('ppt/slides/slide1.xml', s1);
  let pres = await zip.file('ppt/presentation.xml').async('string');
  const ids = pres.match(/<p:sldId [^>]*\/>/g); pres = pres.replace(ids[0] + ids[1], ids[1] + ids[0]);
  zip.file('ppt/presentation.xml', pres);
  // 두 번째로 만든 장(이제 1번)의 노트 쪽을 통째로 지웁니다
  const rels = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string');
  const m = /<Relationship[^>]*notesSlide[^>]*Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"[^>]*\/>/.exec(rels);
  zip.file('ppt/slides/_rels/slide2.xml.rels', rels.replace(m[0], ''));
  zip.remove('ppt/notesSlides/' + m[1]); zip.remove('ppt/notesSlides/_rels/' + m[1] + '.rels');
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(new RegExp('<Override PartName="/ppt/notesSlides/' + m[1] + '"[^>]*/>'), ''));
  if (dropNotesMaster) {
    const pr = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    zip.file('ppt/_rels/presentation.xml.rels', pr.replace(/<Relationship[^>]*notesMaster"[^>]*\/>|<Relationship[^>]*Target="notesMasters\/[^"]*"[^>]*\/>/g, ''));
  }
  return Buffer.from(await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

await test('inspect: 실제 순서, 묶음 안 도형, 표 칸 id, 노트', async () => {
  const info = await PptKit.inspect(await handDeck());
  assert.equal(info.slideCount, 2);
  assert.deepEqual(info.slides[0].shapes.map(s => s.text), ['두 번째 장']);
  assert.equal(info.slides[0].notes, '');
  const s2 = info.slides[1], multi = s2.shapes.find(s => s.name === 'Multi');
  assert.deepEqual(multi.paragraphs, ['굵은 빨강 보통 글', '둘째 문단 기울임']);
  const cells = s2.shapes.filter(s => s.kind === 'table');
  assert.deepEqual(cells.map(c => c.text), ['머리1', '머리2', '칸1', '칸2']);
  assert.match(cells[3].id, /^\d+\.r2c2$/);
  assert.ok(s2.objects.some(o => o.kind === 'group' && o.id === '100'));
  assert.equal(s2.notes, '첫 장 원래 노트');
  const o = PptKit.outline(info);
  assert.ok(o.includes('[슬라이드 2]') && o.includes('#' + multi.id + ' text: "굵은 빨강 보통 글 | 둘째 문단 기울임"') && o.includes('#notes: "첫 장 원래 노트"'));
  assert.ok(PptKit.outline(info, { maxChars: 40 }).includes('생략'));
});

await test('applyEdits: 손으로 만든 덱 (서식 유지, 표 칸, 묶음, 노트 새로 만들기)', async () => {
  const buf = await handDeck(), info = await PptKit.inspect(buf);
  const multi = info.slides[1].shapes.find(s => s.name === 'Multi'), cell = info.slides[1].shapes.find(s => /r2c2$/.test(s.id));
  const res = await PptKit.applyEdits(buf, [
    { slide: 2, shape: multi.id, text: ['바뀐 첫 문단', '둘째 문단 기울임', '셋째 <&"> 문단\u0001'] },
    { slide: 2, shape: cell.id, text: '바뀐 칸' },
    { slide: 2, shape: '100', text: '묶음에는 글 없음' },
    { slide: 2, shape: '100', name: '!!grp' },
    { slide: 1, notes: '[발표 대본 · 약 20초] 새 노트\n[강조 포인트] 하나' },
    { slide: 3, shape: 2, text: 'x' }, { slide: 1, shape: '#999', text: 'x' }, 'junk'
  ], { morph: true });
  assert.equal(res.applied, 4, JSON.stringify(res.skipped));
  assert.equal(res.skipped.length, 4);
  assert.ok(res.skipped.some(s => /없는 슬라이드/.test(s.reason)) && res.skipped.some(s => /없는 도형/.test(s.reason)) && res.skipped.some(s => /글을 넣을 수 없는/.test(s.reason)));
  const out = Buffer.from(await ab(res.blob)); save('hand-edited.pptx', out); fs.writeFileSync(path.join(tmp, 'hand-edited.pptx'), out);
  const before = await unzip(buf), after = await unzip(out);
  xmlOk(after.text, 'hand');
  const info2 = await PptKit.inspect(out);
  const m2 = info2.slides[1].shapes.find(s => s.id === multi.id);
  assert.deepEqual(m2.paragraphs, ['바뀐 첫 문단', '둘째 문단 기울임', '셋째 <&"> 문단']);
  assert.ok(info2.slides[1].shapes.some(s => s.id === cell.id && s.text === '바뀐 칸'));
  assert.equal(info2.slides[0].notes, '[발표 대본 · 약 20초] 새 노트\n[강조 포인트] 하나');
  assert.ok(info2.slides[1].objects.some(o => o.name === '!!grp'));
  const x = after.text['ppt/slides/slide1.xml'], xo = before.text['ppt/slides/slide1.xml'];
  const sp = s => s.slice(s.indexOf('name="Multi"'), s.indexOf('</p:sp>', s.indexOf('name="Multi"')));
  const paras = s => sp(s).match(/<a:p>[\s\S]*?<\/a:p>/g);
  const firstRPr = p => (p.match(/<a:rPr[\s\S]*?(?:\/>|<\/a:rPr>)/) || [''])[0];
  assert.equal(firstRPr(paras(x)[0]), firstRPr(paras(xo)[0]), '첫 문단은 첫 글 조각 서식 유지');
  assert.ok(/b="1"/.test(firstRPr(paras(x)[0])) && /FF0000/.test(firstRPr(paras(x)[0])));
  assert.equal(paras(x)[1], paras(xo)[1], '같은 글 문단은 바이트 그대로 (기울임 유지)');
  assert.equal(firstRPr(paras(x)[2]), firstRPr(paras(xo)[1]), '늘어난 문단은 마지막 문단 서식 복제');
  assert.ok(x.includes('셋째 &lt;&amp;&quot;&gt; 문단'));
  const np = Object.keys(after.text).filter(n => /notesSlides\/notesSlide\d+\.xml$/.test(n) && !before.text[n]);
  assert.equal(np.length, 1, '새 노트 쪽 하나');
  assert.ok(after.text['[Content_Types].xml'].includes('PartName="/' + np[0] + '"'));
  assert.ok(after.text['ppt/slides/_rels/slide2.xml.rels'].includes('notesSlide'));
  assert.ok(after.text['ppt/slides/slide1.xml'].includes(MORPH_XML) && !after.text['ppt/slides/slide2.xml'].includes('p:transition'), '모핑은 2..N 장(순서는 presentation.xml 기준)');
  const py = pyRead(path.join(tmp, 'hand-edited.pptx'));
  if (py) { assert.ok(py[0].n.includes('새 노트')); assert.ok(py[1].t.some(t => t.includes('바뀐 첫 문단'))); }
  const nm = await PptKit.applyEdits(await handDeck({ dropNotesMaster: true }), [{ slide: 1, notes: '노트' }]);
  assert.equal(nm.applied, 0); assert.match(nm.skipped[0].reason, /no notes page/);
});

await test('applyEdits: 만든 12장 덱 왕복 (손대지 않은 항목 SHA-256 동일, 모핑 교체)', async () => {
  const info = await PptKit.inspect(deckBuf);
  const byName = (n, nm) => info.slides[n - 1].shapes.find(s => s.name === nm) || info.slides[n - 1].objects.find(s => s.name === nm);
  const body4 = byName(4, '!!body');
  assert.equal(body4.paragraphs.length, 2);
  const edits = [
    { slide: 3, shape: byName(3, '!!title').id, text: '도구보다\n목표가 먼저' },
    { slide: 4, shape: body4.id, text: ['첫 문단 고침', '둘째', '셋째 <&"> 추가\u0001\u0008'] },
    { slide: 5, shape: byName(5, '!!head').id, text: '새 카드 제목' },
    { slide: 5, shape: byName(5, '!!head').id, text: '새 카드 제목 (마지막)' },
    { slide: 2, notes: '[발표 대본 · 약 30초] 새 대본\n[강조 포인트] 목표\n── 원본 메모 ──\n' + NOTE },
    { slide: 6, shape: byName(6, '!!stat').id, name: '!!number' },
    { slide: 99, shape: 2, text: 'x' }, { slide: 3, shape: 9999, text: 'x' }, { slide: 3, shape: byName(3, '!!title').id, text: '도구보다\n목표가 먼저' }
  ];
  const res = await PptKit.applyEdits(deckBuf, edits, { morph: true });
  assert.equal(res.applied, 5, JSON.stringify(res.skipped));
  assert.ok(res.skipped.some(s => /마지막 것만/.test(s.reason)));
  const out = Buffer.from(await ab(res.blob)); save('deck-edited.pptx', out); fs.writeFileSync(path.join(tmp, 'deck-edited.pptx'), out);
  const before = await unzip(deckBuf), after = await unzip(out);
  xmlOk(after.text, 'deck-edited');
  assert.deepEqual(after.names, before.names, '항목 목록과 순서 그대로');
  const notes2 = 'ppt/' + /Target="\.\.\/(notesSlides\/notesSlide\d+\.xml)"/.exec(before.text['ppt/slides/_rels/slide2.xml.rels'])[1];
  const touched = new Set(['ppt/slides/slide3.xml', 'ppt/slides/slide4.xml', 'ppt/slides/slide5.xml', 'ppt/slides/slide6.xml', notes2]);
  for (const n of before.names) {
    if (touched.has(n)) assert.notEqual(sha(after.raw[n]), sha(before.raw[n]), n + ' 바뀌어야 함');
    else assert.equal(sha(after.raw[n]), sha(before.raw[n]), n + ' 는 바이트 그대로여야 함');
  }
  const x3 = after.text['ppt/slides/slide3.xml'], x4 = after.text['ppt/slides/slide4.xml'];
  assert.ok(/<a:t>도구보다<\/a:t><\/a:r><a:br>/.test(x3), '\\n 은 a:br');
  assert.ok(x4.includes('셋째 &lt;&amp;&quot;&gt; 추가</a:t>') && !/[\u0001\u0008]/.test(x4));
  for (let k = 2; k <= 12; k++) assert.equal(after.text['ppt/slides/slide' + k + '.xml'].split(MORPH_XML).length - 1, 1, 'slide ' + k + ' 모핑 하나만');
  const info2 = await PptKit.inspect(out);
  assert.equal(info2.slides[2].title, '도구보다 목표가 먼저');
  assert.deepEqual(info2.slides[3].shapes.find(s => s.id === body4.id).paragraphs, ['첫 문단 고침', '둘째', '셋째 <&"> 추가']);
  assert.ok(info2.slides[1].notes.startsWith('[발표 대본 · 약 30초]') && info2.slides[1].notes.endsWith(NOTE));
  assert.ok(info2.slides[5].shapes.some(s => s.name === '!!number'));
  const py = pyRead(path.join(tmp, 'deck-edited.pptx'));
  if (py) { assert.equal(py.length, 12); assert.ok(py[3].t.some(t => t.includes('첫 문단 고침'))); }
});

await test('applyEdits: keepOriginal 은 원래 노트를 "── 원본 메모 ──" 아래에 한 번만 남김', async () => {
  const r1 = await PptKit.applyEdits(deckBuf, [{ slide: 3, notes: '[발표 대본 · 약 20초]\n첫 대본', keepOriginal: true }]);
  const b1 = Buffer.from(await ab(r1.blob));
  const r2 = await PptKit.applyEdits(b1, [{ slide: 3, notes: '[발표 대본 · 약 25초]\n둘째 대본', keepOriginal: true }]);
  const info = await PptKit.inspect(Buffer.from(await ab(r2.blob)));
  assert.equal(info.slides[2].notes, '[발표 대본 · 약 25초]\n둘째 대본\n── 원본 메모 ──\n' + NOTE);
});

await test('view-ppt.js: 상수, 길이, 예시 JSON 이 그대로 읽힘', async () => {
  const V = await import('./functions/_lib/view-ppt.js');
  assert.equal(V.DECK_MAX_TOKENS, 10000); assert.equal(V.EDIT_MAX_TOKENS, 8000);
  assert.ok(V.DECK_PROMPT.length < 9000 && V.EDIT_PROMPT.length < 9000);
  for (const l of PptKit.LAYOUTS) assert.ok(V.DECK_PROMPT.includes(l + ' '), l);
  const deck = PptKit.parseModelJson(V.DECK_PROMPT.slice(V.DECK_PROMPT.indexOf('[예시')));
  const v = PptKit.validateDeck(deck);
  assert.equal(v.spec.slides.length, 3); assert.deepEqual(v.warnings, []);
  const ed = PptKit.parseModelJson(V.EDIT_PROMPT.slice(V.EDIT_PROMPT.indexOf('[예시]')));
  assert.ok(Array.isArray(ed.edits) && ed.morph === false && ed.edits.some(e => e.keepOriginal === true));
  assert.ok(V.EDIT_PROMPT.includes('[발표 대본 · 약 N초]') && V.EDIT_PROMPT.includes('── 원본 메모 ──') && V.EDIT_PROMPT.includes('[강조 포인트]'));
});

function findReal() {
  if (process.env.PK_REAL) return process.env.PK_REAL.split(path.delimiter).filter(f => fs.existsSync(f));
  const roots = ['/Volumes/kim/2026', path.join(os.homedir(), 'Downloads'), path.join(os.homedir(), 'Desktop')], found = [];
  const walk = (d, depth) => {
    let ents = []; try { ents = fs.readdirSync(d, { withFileTypes: true }); } catch (e) { return; }
    for (const e of ents) {
      if (found.length >= 40 || e.name.startsWith('.') || e.name.startsWith('~$')) continue;
      const f = path.join(d, e.name);
      if (e.isDirectory() && depth < 3 && !/node_modules|ai-class-lab/.test(e.name)) walk(f, depth + 1);
      else if (e.isFile() && /\.pptx$/i.test(e.name)) { const sz = fs.statSync(f).size; if (sz > 40000 && sz < 25e6) found.push([f, sz]); }
    }
  };
  roots.forEach(r => walk(r, 0));
  return found.sort((a, b) => a[1] - b[1]).filter((x, i, a) => i === 0 || i === a.length - 1 || i === a.length >> 1).map(x => x[0]).slice(0, 3);
}

await test('실제 덱: inspect + applyEdits (읽기만, 결과는 임시 폴더)', async () => {
  const files = findReal();
  if (!files.length) { results.push('     (실제 .pptx 를 찾지 못해 건너뜀 — 손으로 만든 덱 시험으로 대신함)'); return; }
  for (const f of files) {
    const buf = fs.readFileSync(f), info = await PptKit.inspect(buf);
    assert.ok(info.slideCount > 0, f);
    const outline = PptKit.outline(info); assert.ok(outline.includes('[슬라이드 1]'));
    const edits = [];
    info.slides.slice(0, 3).forEach(s => { const sh = s.shapes.find(x => x.kind !== 'table') || s.shapes[0]; if (sh) edits.push({ slide: s.n, shape: sh.id, text: sh.paragraphs.map((p, i) => i === 0 ? '[고침] ' + p : p) }); });
    edits.push({ slide: 1, notes: '[발표 대본 · 약 20초] 시험 노트\n── 원본 메모 ──\n' + info.slides[0].notes });
    const res = await PptKit.applyEdits(buf, edits, { morph: true });
    const out = Buffer.from(await ab(res.blob));
    const before = await unzip(buf), after = await unzip(out);
    xmlOk(Object.fromEntries(Object.entries(after.text).filter(([n]) => n.startsWith('ppt/slides/') || n.startsWith('ppt/notesSlides/') || n === '[Content_Types].xml')), path.basename(f));
    const changedParts = new Set();
    for (const n of before.names) if (sha(after.raw[n]) !== sha(before.raw[n])) changedParts.add(n);
    for (const n of changedParts) assert.ok(/^ppt\/(slides\/slide\d+\.xml|slides\/_rels\/slide\d+\.xml\.rels|notesSlides\/notesSlide\d+\.xml)$|^\[Content_Types\]\.xml$/.test(n), path.basename(f) + ': 바뀌면 안 되는 항목 ' + n);
    const info2 = await PptKit.inspect(out);
    edits.filter(e => e.shape).forEach(e => assert.ok(info2.slides[e.slide - 1].shapes.some(s => s.id === e.shape && s.text.startsWith('[고침] ')), path.basename(f) + ' slide ' + e.slide));
    const noteSkip = res.skipped.find(x => x.edit && x.edit.notes != null);
    assert.ok(info2.slides[0].notes.startsWith('[발표 대본') || (noteSkip && /no notes page/.test(noteSkip.reason)), path.basename(f) + ' 노트: ' + JSON.stringify(noteSkip || info2.slides[0].notes.slice(0, 80)));
    const name = 'real-' + path.basename(f).replace(/[^\w.-]+/g, '_');
    fs.writeFileSync(path.join(tmp, name), out); save(name, out);
    const py = pyRead(path.join(tmp, name));
    if (py) assert.equal(py.length, info.slideCount);
    ran.realDecks.push(path.basename(f) + ' (' + info.slideCount + '장, 고침 ' + res.applied + ', 건너뜀 ' + res.skipped.length + ', 모핑 ' + res.morphed.length + '장)');
  }
});

console.log(results.join('\n'));
console.log('\nxmllint: ' + (ran.xmllint ? '사용함' : '없음(정규식 점검)') + ' · python-pptx: ' + (ran.pythonPptx ? '열어 봄' : hasPyPptx ? '안 씀' : '없음'));
console.log('실제 덱: ' + (ran.realDecks.length ? ran.realDecks.join(' / ') : '없음'));
fs.rmSync(tmp, { recursive: true, force: true });
if (failures) { console.log('\n' + failures + '개 실패'); process.exit(1); }
console.log('\n모두 통과');
