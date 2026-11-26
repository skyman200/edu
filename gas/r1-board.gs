/* ===== 레시피 1. 실시간 교실 반응판 =====
   사용법: 웹 도구(https://kimcook-edu.pages.dev/#gas)에서 이 레시피를 고르고 안내를 따르세요. 코드를 붙여 넣고 start 를 실행하면 준비가 끝납니다.
   질문과 묶음은 강의장 큰 화면에서 바로 바꿀 수 있습니다(교수 PIN 필요). [설정] 시트에서 바꿔도 됩니다.
   답은 [응답] 시트에 쌓입니다. 학생이 폰에서 묶음 칩을 고르면 그 묶음으로, 안 고르면 먼저 낱말 맞추기(API 없음)로, 그래도 못 정하면 Jev가 묶음을 고릅니다.
   큰 화면의 [분류](교수 PIN)는 기타·미분류·오류인 답을 다시 나누고, 묶음이 없거나 기타가 많으면 비슷한 답끼리 묶어 바로 저장합니다. 응답을 한꺼번에 AI에 보내 새 묶음을 만들 때만 교수의 동의가 필요합니다. */
var BOARD_HEAD = ['시간', '응답', '묶음', '확신도', '상태', '번호', '자동 묶음'];
var BOARD_AI_PROP = 'BOARD_AI_GROUPING', BOARD_AI_GEN = 'BOARD_AI_GENERATION';
var BOARD_SOLO = typeof ensureKey_ === 'function', BOARD_AI_NAME = BOARD_SOLO ? 'Gemini' : CLASS_AI_NAME;
var BOARD_AI_BUSY = '다른 작업이 진행 중이라 동의 상태를 바꾸지 못했습니다. 잠시 후 다시 시도하십시오.';
var AUTO_EMPTY = '미분류 (자동, 묶음 없음)'; // 쉼표를 쓸 수 없는 실제 묶음 이름과 구분합니다.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('반응판')
    .addItem('처음 준비', 'setup').addItem('새 질문 시작(지금 응답은 보관)', 'newRound').addItem('수업 코드 입력', 'setClassCode').addItem('주소와 교수 PIN 보기', 'showUrls').addItem('웹 앱 주소 넣기', 'setWebUrl')
    .addItem('AI 묶기 끄기', 'disableBoardAiGrouping').addToUi();
}

function disableBoardAiGrouping() {
  menuOnly_();
  var props = PropertiesService.getScriptProperties(), lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { say_(BOARD_AI_BUSY); return false; }
  try {
    props.setProperty(BOARD_AI_PROP, '0');
    props.setProperty(BOARD_AI_GEN, String(Number(props.getProperty(BOARD_AI_GEN) || 0) + 1));
  } finally { lock.releaseLock(); }
  say_('응답을 한꺼번에 보내 새 묶음을 만드는 기능을 껐습니다. 이미 전송이 승인된 요청은 전송될 수 있으며, 전송된 요청은 취소할 수 없습니다. 다시 켜려면 큰 화면의 [분류]에서 동의해 주십시오.');
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  var s = ss.getSheetByName('설정');
  if (!s) {
    s = ss.insertSheet('설정');
    s.getRange('A1:B4').setValues([
      ['질문', 'AI를 수업에 쓰는 것에서 가장 걱정되는 점은?'],
      ['묶음(쉼표로 구분, 10개까지)', '부정행위, 평가 방법, 학습 효과 저하, 업무 부담'],
      ['받는 중(O/X)', 'O'],
      ['강의장 화면 새로고침(초)', 4]
    ]);
    s.setColumnWidth(1, 220); s.setColumnWidth(2, 520);
  }
  var gh = sheet_('응답', BOARD_HEAD).getRange(1, 7);
  if (String(gh.getValue()) === '') gh.setValue('자동 묶음').setFontWeight('bold'); // 옛 6열 사본의 빈 머리글
  // 수업 서버판이면 판정 서버 연결을 확인합니다. 실패해도 준비는 계속하고, 이유를 알립니다(수업 코드 없이 시작하면 답이 모두 기타로 보이기 때문).
  var warn = '';
  var why = typeof checkServer_ === 'function' ? checkServer_() : '';
  if (why) warn = '\n\n※ 판정 서버에 연결하지 못했습니다: ' + why + '\n시트 메뉴 반응판 → 수업 코드 입력을 하세요. 그대로 두면 답은 받지만 묶음을 고르지 못해 기타로 보입니다. 큰 화면 [분류]로 나중에 다시 나눌 수 있습니다.';
  say_('준비 끝! 이제 웹 도구 화면의 다음 단계(배포)를 따라 하세요.\n큰 화면에서 질문을 바꿀 때 쓰는 교수 PIN: ' + pin_('board') + ' — 학생에게 알리지 마세요. 시트 메뉴 반응판 → 주소와 교수 PIN 보기에서 다시 볼 수 있습니다.' + warn);
}

// 화면이 다시 그려야 하는 변화(질문 바꾸기, 보관)마다 1씩 올립니다. 큰 화면은 이보다 낮은 번호의 늦게 온 응답을 버립니다.
function bump_() {
  try {
    var p = PropertiesService.getScriptProperties(), r = Number(p.getProperty('BOARD_REV') || 0) + 1;
    p.setProperty('BOARD_REV', String(r));
    return r;
  } finally {
    var cache = CacheService.getScriptCache();
    try { cache.remove('board'); } finally { cache.remove('q'); }
  }
}

// 지금까지의 응답을 [응답_월일_시분초] 시트로 옮겨 보관하고, 빈 [응답] 시트를 새로 만듭니다. 지우지 않습니다.
// 부르는 쪽이 스크립트 잠금을 잡고 있어야 합니다(그 사이 들어온 답이 머리글 없는 시트로 가지 않게).
function archive_() {
  var ss = ss_(), sh = ss.getSheetByName('응답');
  if (!sh || sh.getLastRow() <= 1) return false;
  var base = '응답_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd_HHmmss'), name = base;
  for (var k = 2; ss.getSheetByName(name); k++) name = base + '_' + k;
  sh.setName(name);
  sheet_('응답', BOARD_HEAD);
  bump_();
  SpreadsheetApp.flush();
  return true;
}

function newRound() {
  menuOnly_();
  var lock = LockService.getScriptLock(), done;
  lock.waitLock(20000);
  try { done = archive_(); } finally { lock.releaseLock(); }
  say_(done ? '보관했습니다. 큰 화면이 곧 비워집니다. 질문은 큰 화면의 "질문 바꾸기"에서 바꿀 수 있습니다.' : '비울 응답이 아직 없습니다.');
}

function showUrls() {
  menuOnly_();
  var url = webUrl_() || setWebUrl();
  say_((url ? '학생용: ' + url + '\n강의장 화면: ' + url + '?v=board\n' : '웹 앱 주소가 아직 없습니다. 배포 → 배포 관리에서 웹 앱 URL(/exec)을 복사해 메뉴 반응판 → 웹 앱 주소 넣기에 붙여 넣으세요.\n') + '질문 바꾸기 PIN: ' + pin_('board'));
}

// 묶음 정리: 줄바꿈이나 쉼표로 나누고, 빈 것, 같은 이름, '기타'(따로 붙음)를 빼고 10개까지. 시트에서 고쳐도 같은 규칙입니다.
function cats_(text) {
  var seen = {};
  return String(text || '').split(/[\n,]/).map(function (x) { return x.replace(/\s+/g, ' ').trim().slice(0, 24); })
    .filter(function (x) { if (!x || x === '기타' || seen[x]) return false; seen[x] = 1; return true; }).slice(0, 10);
}

function cfg_() {
  var v = ss_().getSheetByName('설정').getRange('B1:B4').getValues();
  return { q: String(v[0][0]), cats: cats_(v[1][0]), open: String(v[2][0]).trim().toUpperCase() === 'O', refresh: Math.max(2, Number(v[3][0]) || 4) };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.v === 'board') {
    // 참여 QR 주소: 시트 메뉴로 넣어 둔 주소가 먼저이고, 없으면 웹 도구가 큰 화면 주소에 붙여 준 join 을 이 화면에서만 씁니다(저장하지 않음).
    var saved = webUrl_(), asked = execUrl_(e.parameter.join), id = function (u) { var m = /\/s\/([\w-]+)\/exec$/.exec(u); return m ? m[1] : ''; };
    var boot = { join: saved || asked, warn: !!(saved && asked && id(saved) !== id(asked)) };
    var html = BOARD_HTML.replace('/*BOOT*/null', function () { return JSON.stringify(boot).replace(/</g, '\\u003c'); });
    return page_('반응판', html, { raw: true, css: BOARD_CSS + STAGE_CSS });
  }
  return page_('내 생각 보내기', ASK_HTML, { raw: true, css: ASK_CSS });
}

function rev_() { return Number(PropertiesService.getScriptProperties().getProperty('BOARD_REV') || 0); }
function round_() { var sh = ss_().getSheetByName('응답'); return sh ? sh.getSheetId() : 0; }

// 학생 폰이 12초마다 묻습니다. cats는 폰에 칩으로 보이는 묶음입니다(두 개 미만이면 묶지 않는 질문이라 빈 목록). 사람이 많아도 시트를 매번 읽지 않게 5초 동안 기억합니다.
// round: 지금 [응답] 시트. 보관 뒤에는 바뀌므로, 폰은 이전 질문에 쓴 답을 새 판에 보내지 않습니다.
// 읽는 사이 질문이 바뀌었으면(rev) 기억하지 않습니다.
function getQuestion() {
  var cache = CacheService.getScriptCache(), hit = cache.get('q');
  if (hit) return JSON.parse(hit);
  var rev = rev_(), c = cfg_(), out = { q: c.q, open: c.open, round: round_(), cats: c.cats.length >= 2 ? c.cats : [] };
  if (rev_() === rev) cache.put('q', JSON.stringify(out), 5);
  return out;
}

// ===== API 없이 묶기: 낱말 맞추기와 비슷한 답 묶기 (AI를 부르지 않아 두 판 모두 비용이 없습니다) =====
// 뜻이 같아도 낱말이 다르면('시험 때 몰래' ↔ '부정행위') 맞추지 못합니다. 그런 답은 Jev가 고릅니다.
// 묶음 이름의 낱말이라도 너무 흔한 말(방법, 효과, 부담, 학생…)은 혼자서는 묶음을 정하지 못합니다('사용 방법을 몰라서'가 '평가 방법'으로 가지 않게).
var WORD_STOP = ('방법 문제 효과 부분 관련 사용 활용 증가 감소 부족 저하 향상 부담 정도 내용 생각 걱정 우려 기타 능력 시간 결과 과정 ' +
  '학생 교수 수업 강의 대한 때문 같은 같아 같다 정말 너무 그냥 진짜 조금 많이 많아 아직 모르 몰라 어떻 어려 하는 하기 하고 해서 ' +
  '해요 했다 되는 되고 될까 있는 있을 없는 없을 싶은 싶어 만들 바꾸 바꿔 늘리 줄이 주기 받기 좋겠 좋을 추가 변경 더욱 지금 이번 다음 ai ' +
  '정리 작성 처리 관리 진행 제공 확인').split(' ');
var JOSA_RE = /(으로써|으로서|에서는|에게서|이라도|까지|부터|처럼|보다|으로|에서|에게|한테|이랑|이나|하고|은|는|이|가|을|를|의|에|도|만|로|와|과|랑)$/;
// 글자·숫자만 남기고 낱말로 나눕니다. 낱말 뒤 조사는 두 글자 이상 남을 때만 뗍니다(업무가 → 업무, 평가 → 평가).
function words_(s) {
  return String(s || '').toLowerCase().replace(/[^0-9a-z가-힣]+/g, ' ').trim().split(' ').filter(String).map(function (w) {
    var t = w.replace(JOSA_RE, '');
    return t.length >= 2 ? t : w;
  });
}
function compact_(s) { return String(s || '').toLowerCase().replace(/[^0-9a-z가-힣]+/g, ''); }

// 낱말 맞추기: 답이 묶음 하나에만 분명히 맞으면 그 묶음 번호, 아니면 -1(Jev로 넘김).
// 맞음(묶음 이름이 한 낱말): 이름 낱말(흔한 말·질문에 나온 말·다른 묶음과 겹치는 말은 제외)로 시작하는 낱말이 답에 있음. 이름이 답 낱말 안에 들어 있기만 한 것('정보' ↔ '개인정보')은 맞지 않습니다.
// 맞음(이름이 두 낱말 이상): 띄어쓰기를 뺀 답에 이름이 통째로 있거나, 이름의 낱말 머리가 둘 이상 답에 있고 그중 하나는 이 묶음만 쓰는 낱말임. 흔한 말·질문 낱말로만 된 이름은 낱말 규칙을 따르지 않습니다.
//   낱말 하나만 맞는 것('AI 학습에 쓰이는 것' ↔ '학습 효과 저하')은 뜻이 다를 수 있어 Jev가 정합니다.
// 부정·반대 낱말(안/않/없/못/괜찮/아니/아닌/아냐/줄/덜/오히려/별로/딱히/쉬웠·쉽·편했, 붙여 쓴 '안함' '안해' '없음' 포함)이 있으면 뜻이 뒤집힐 수 있어 정하지 않습니다.
var NEG_RE = /^(안$|안(함|해|하|했|돼|되|된|써|쓰|봐|본|갈|가|느|넣)|않|없|못|괜찮|아니|아닌|아냐|줄|덜|오히려|별로|딱히|쉬웠|쉬워|쉬운|쉬움|쉽|편했|편해)/;
// 낱말 안쪽이나 붙여 쓴 꼴('걱정없음', '안될', '아닐', '~보다는', '좋아질')도 띄어쓰기를 뺀 답 전체에서 찾아 Jev로 넘깁니다.
var NEG_IN_RE = /없|않|못|안[되됨될돼된생]|아닐|아니|괜찮|별로|딱히|오히려|덜|줄[어었을]|보다|말고|빼고|대신|좋아|높아|늘[어었]|향상|개선|감소|가벼/;
function matchCat_(text, cats, q) {
  var ws = words_(text), flat = compact_(text);
  if (!flat || NEG_IN_RE.test(flat) || ws.some(function (w) { return NEG_RE.test(w); })) return -1;
  var qw = {}, own = {};
  words_(q).forEach(function (w) { qw[w] = 1; });
  var uniq = function (name) { return words_(name).filter(function (w, i, a) { return w.length >= 2 && a.indexOf(w) === i; }); };
  var head = function (cw) { return ws.some(function (w) { return w.indexOf(cw) === 0; }); };
  var all = cats.map(uniq), lists = all.map(function (l) { return l.filter(function (w) { return WORD_STOP.indexOf(w) < 0 && !qw[w]; }); });
  lists.forEach(function (l) { l.forEach(function (w) { own[w] = (own[w] || 0) + 1; }); });
  var hits = [];
  cats.forEach(function (name, i) {
    var nw = words_(name).length, multi = nw >= 2, mine = lists[i].filter(function (cw) { return own[cw] === 1; }), hit = false;
    if (!lists[i].length) return;
    if (!multi) hit = mine.some(head);
    else {
      var full = compact_(name);
      hit = full.length >= 2 && flat.indexOf(full) >= 0;
      // 세 낱말 이상인 이름은 통째로 있을 때만 맞춥니다(낱말 둘만 겹치면 뜻이 반대일 수 있음).
      if (!hit && nw < 3 && mine.some(head)) hit = all[i].filter(function (cw) { return !qw[cw] && head(cw); }).length >= 2;
    }
    if (hit) hits.push(i);
  });
  return hits.length === 1 ? hits[0] : -1;
}

// 비슷한 답 묶기: 여러 답에 함께 나오는 낱말 머리(2글자 이상)를 찾아, 가장 많은 답이 나누는 것부터 한 묶음씩 떼어 냅니다.
// 질문에 나온 말, 흔한 말(로 시작하는 말: 늘리기, 만들어서…), 답의 40%를 넘게 덮는 말은 쓰지 않습니다. 묶음 이름은 그 묶음 답의 60% 이상에 나오는 가장 긴 낱말 머리(12글자 안)입니다.
// 묶이지 않은 답은 기타로 둡니다. 돌려주는 값: [{name, idx: [답 번호…]}] (큰 묶음부터)
function clusterAnswers_(texts, q, max, avoid) {
  var n = texts.length, minSize = n >= 20 ? 3 : 2, stems = texts.map(words_), banned = {}, out = [], left = {};
  words_(q).forEach(function (w) { for (var k = 2; k <= w.length; k++) banned[w.slice(0, k)] = 1; });
  // 이미 있는 묶음 이름은 통째로도, 이름의 낱말 머리로도('평가 방법' 옆에 '평가') 새 묶음 이름이 되지 못합니다.
  (avoid || []).forEach(function (a) {
    banned[compact_(a)] = 1;
    words_(a).forEach(function (w) { if (w.length >= 2) for (var k = 2; k <= w.length; k++) banned[w.slice(0, k)] = 1; });
  });
  var stop = function (p) { return banned[p] || WORD_STOP.some(function (w) { return p.indexOf(w) === 0; }); };
  var feats = stems.map(function (ws) {
    var f = {};
    ws.forEach(function (w) { for (var k = 2; k <= Math.min(w.length, 12); k++) if (!stop(w.slice(0, k))) f[w.slice(0, k)] = 1; });
    return f;
  });
  var df = {};
  feats.forEach(function (f) { Object.keys(f).forEach(function (k) { df[k] = (df[k] || 0) + 1; }); });
  for (var i = 0; i < n; i++) left[i] = 1;
  while (out.length < (max || 6)) {
    var cnt = {};
    Object.keys(left).forEach(function (i) { Object.keys(feats[i]).forEach(function (k) { if (df[k] <= Math.max(minSize, n * 0.4)) cnt[k] = (cnt[k] || 0) + 1; }); });
    var best = Object.keys(cnt).filter(function (k) { return cnt[k] >= minSize; })
      .sort(function (a, b) { return cnt[b] - cnt[a] || b.length - a.length || (a < b ? -1 : 1); })[0];
    if (!best) break;
    var idx = Object.keys(left).map(Number).filter(function (i) { return feats[i][best]; });
    idx.forEach(function (i) { delete left[i]; });
    var name = best, need = Math.ceil(idx.length * 0.6), longest = {};
    idx.forEach(function (i) { stems[i].forEach(function (w) { if (w.indexOf(best) === 0) for (var k = best.length + 1; k <= Math.min(w.length, 12); k++) longest[w.slice(0, k)] = 1; }); });
    Object.keys(longest).forEach(function (p) {
      var c = idx.filter(function (i) { return stems[i].some(function (w) { return w.indexOf(p) === 0; }); }).length;
      if (c >= need && (p.length > name.length || (p.length === name.length && p < name))) name = p;
    });
    out.push({ name: name, idx: idx });
  }
  return out;
}

// 학생 답 하나를 Jev로 묶음에 넣습니다. 돌려주는 값: {cat, conf, bad}. 확신도가 JEV_MIN보다 낮으면 '기타'.
var JEV_MIN = 0.45;
function jevCat_(c, text, withBad) {
  var crit = {};
  c.cats.forEach(function (name, i) { crit['c' + i] = name; });
  crit.other = '위 항목 어디에도 맞지 않는 내용';
  var qs = { cat: { type: 'choice', instructions: '교수의 질문에 대한 응답이 어느 묶음에 속하는지 고르세요.', criteria: crit } };
  if (withBad) qs.bad = { type: 'noul', instructions: '응답이 질문과 무관한 광고, 욕설, 장난, 개인정보 노출인가?' };
  var a = jev_({ 질문: c.q, 응답: text }, qs);
  // 답 모양을 엄격히 봅니다: 'other'만 기타이고, 모르는 모양이면 오류로 던져 '오류:' 흔적을 남깁니다(조용히 기타로 가지 않게).
  var k = a && a.cat && a.cat.choice, conf = a && a.cat && a.cat.confidence, n = -1;
  if (typeof conf !== 'number' || !isFinite(conf)) throw new Error('Jev 답 형식이 다릅니다: confidence ' + conf);
  if (k !== 'other') {
    if (typeof k !== 'string' || !/^c\d+$/.test(k) || !c.cats[Number(k.slice(1))]) throw new Error('Jev 답 형식이 다릅니다: ' + String(k).slice(0, 20));
    n = Number(k.slice(1));
  }
  conf = Math.round(conf * 100) / 100;
  return { cat: conf < JEV_MIN || n < 0 ? '기타' : c.cats[n], conf: conf, bad: withBad ? a.bad.p > 0.8 : false };
}
function errMark_(err) { return '오류: ' + String(err && err.message || err).replace(/\s+/g, ' ').slice(0, 80); }

// pick: 학생이 칩으로 고른 묶음(없으면 빈 값). 지금 묶음 목록에 있는 이름이나 '기타'일 때만 받고, 그러면 묶음은 Jev로 고르지 않습니다(스팸 판정만 합니다).
// 칩이 없으면 먼저 낱말 맞추기(API 없음, 확신도 칸 '낱말'), 못 정하면 Jev가 고릅니다. 스팸 판정은 어느 길이든 Jev 한 번으로 합니다.
// Jev가 실패하면 답은 그대로 받고 묶음을 비운 채 확신도 칸에 '오류: 이유'를 남깁니다. 큰 화면의 [분류]가 이 답을 다시 나눕니다.
// 자동 묶음 칸(G): 이 코드가 쓴 묶음 또는 미분류 표시. 교수의 수동 분류는 [분류]가 건드리지 않습니다.
function submitAnswer(text, sid, round, pick) {
  var c = cfg_();
  if (!c.open) return { ok: false, msg: '지금은 답을 받지 않습니다.' };
  text = String(text || '').replace(/\s+/g, ' ').trim().slice(0, 140);
  if (text.length < 2) return { ok: false, msg: '두 글자 이상 써 주세요.' };
  sid = String(sid || 'x').slice(0, 40);
  if (!throttle_('s:' + sid, 6, 600) || !throttle_('all', 90, 60)) return { ok: false, msg: '잠시 뒤에 다시 보내 주세요.' };
  var cat = '', conf = '', status = '표시', jevOther = false;
  pick = String(pick || '').trim();
  var picked = c.cats.length >= 2 && (pick === '기타' || c.cats.indexOf(pick) >= 0);
  var word = !picked && c.cats.length >= 2 ? matchCat_(text, c.cats, c.q) : -1;
  if (picked || word >= 0) {
    cat = picked ? pick : c.cats[word]; conf = picked ? '직접' : '낱말';
    try {
      var chk = jev_({ 질문: c.q, 응답: text }, { bad: { type: 'noul', instructions: '응답이 질문과 무관한 광고, 욕설, 장난, 개인정보 노출인가?' } });
      if (chk.bad.p > 0.8) status = '숨김';
    } catch (err) {}
  } else if (c.cats.length >= 2) {
    try {
      var j = jevCat_(c, text, true);
      cat = j.cat; conf = j.conf; jevOther = j.cat === '기타';
      if (j.bad) status = '숨김';
    } catch (err) { cat = ''; conf = errMark_(err); }
  }
  // 번호: 답마다 한 번만 붙는 고유 번호입니다. 교수가 시트에서 줄을 지우거나 정렬해도 큰 화면이 답을 헷갈리지 않습니다.
  var lock = LockService.getScriptLock(), no;
  lock.waitLock(10000);
  try {
    var p = PropertiesService.getScriptProperties(), id = Math.max(1000000, Number(p.getProperty('BOARD_SEQ') || 0)) + 1;
    p.setProperty('BOARD_SEQ', String(id));
    var old = ss_().getSheetByName('응답');
    if (round && old && old.getSheetId() !== Number(round)) return { ok: false, stale: true, msg: '그사이 새 질문으로 넘어갔습니다. 질문을 확인하고 다시 보내 주세요.' };
    var sh = sheet_('응답', BOARD_HEAD);
    row_(sh, [new Date(), text, cat, conf, status, id, cat || AUTO_EMPTY]);
    // Jev가 이 묶음 목록에서 기타로 본 답은 [분류]가 같은 목록으로 다시 묻지 않게 기억합니다(이미 잡은 잠금 안).
    if (jevOther) seenAdd_(p, c.cats.join('|'), [id]);
    no = sh.getLastRow() - 1;
  } finally { lock.releaseLock(); }
  return { ok: true, cat: cat, no: no };
}

// 큰 화면의 [분류] (교수만): 한 번 누르면 알아서 묶입니다.
// (가) 묶음이 있으면 '기타'·미분류·오류인 답(지금 목록에 없는 묶음 이름 포함)을 낱말 맞추기로, 그래도 못 정하면 Jev로 다시 나눠 시트에 씁니다.
//     학생이 칩으로 고른 답('직접')과 교수가 시트에서 손으로 고친 묶음(자동 묶음 칸과 다름)은 건드리지 않습니다.
//     Jev는 한 번에 CLASSIFY_MAX개, 시간은 CLASSIFY_SEC초까지만 부릅니다(Apps Script 6분 한도, 수업 API 한도). 남은 답은 다시 누르면 이어서 합니다.
//     같은 묶음 목록에서 Jev가 이미 '기타'라고 본 답(제출 때 포함)은 다음에 또 묻지 않습니다.
// (나) 묶음이 없거나 (가) 뒤에도 기타가 많으면(4개 이상, 25% 이상) 기타인 답끼리 스스로 묶어 바로 저장하고 답을 옮깁니다.
//     먼저 API 없이 비슷한 낱말로 묶어 보고, 답의 60% 이상이 묶이면 그것을 씁니다(API 0번).
//     그보다 적게 묶이고 교수가 동의한 경우에만 AI 글쓰기(ai_)로 묶고, AI가 빠뜨린 답만 Jev로 나눕니다.
//     실패하면 이유를 알립니다. 되돌리려면 질문 바꾸기에서 이번에 새로 생긴 묶음 이름만 지우고 바꾸기를 누릅니다(그 답은 기타로 돌아갑니다).
var CLASSIFY_MAX = typeof ensureKey_ === 'function' ? 10 : 20, CLASSIFY_SEC = typeof ensureKey_ === 'function' ? 120 : 150;
var CLASSIFY_RATE = typeof ensureKey_ === 'function' ? 30 : 90;
var AUTO_SHARE = 0.6, AUTO_LIST = 80;
function classifyBoard(cred, confirmKey) {
  var chk = auth_(cred, 'board');
  if (!chk.ok) return chk;
  var cache = CacheService.getScriptCache(), props = PropertiesService.getScriptProperties(), generation = props.getProperty(BOARD_AI_GEN) || '0';
  var prefix = 'board-ai:' + chk.token + ':';
  if (cache.get('classifying')) return { ok: false, busy: true, msg: '이미 분류하고 있습니다. 끝나면 결과가 뜹니다.', token: chk.token };
  if (confirmKey) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { ok: false, msg: BOARD_AI_BUSY, token: chk.token, consentBusy: true };
    try {
      generation = props.getProperty(BOARD_AI_GEN) || '0';
      var consent = typeof confirmKey === 'string' && /^[0-9a-f-]{36}$/.test(confirmKey) ? prefix + confirmKey : '';
      if (!consent || cache.get(consent) !== generation) return { ok: false, msg: '동의 확인이 만료되었습니다. [분류]를 다시 누르십시오.', token: chk.token };
      cache.remove(consent);
      props.setProperty(BOARD_AI_PROP, '1');
    } finally { lock.releaseLock(); }
  } else if (!throttle_('classify', 4, 60)) return { ok: false, msg: '잠시 뒤에 다시 눌러 주세요.', token: chk.token };
  cache.put('classifying', '1', 400);
  try {
    var out = classify_(chk.token, generation);
    if (out.askAi && generation === (props.getProperty(BOARD_AI_GEN) || '0')) {
      out.confirmKey = Utilities.getUuid();
      cache.put(prefix + out.confirmKey, generation, 300);
    }
    return out;
  } finally { cache.remove('classifying'); }
}

var LIMIT_MSG = typeof ensureKey_ === 'function' ? 'AI 사용 한도(1분 30번)에 닿았습니다. 1분쯤 뒤에 다시 눌러 주세요.' : '수업 API 한도(1분 90번)에 닿았습니다.';
function isOther_(c, r) { return c.cats.length >= 2 && c.cats.indexOf(String(r[2])) < 0; }
// 직접 고른 답과 수동 분류는 보호합니다. G가 빈 옛 행도 C가 채워져 있으면 수동으로 봅니다.
function isFree_(r) {
  var auto = String(r[6] == null ? '' : r[6]);
  return String(r[3]) !== '직접' && String(r[2]) === (auto === AUTO_EMPTY ? '' : auto);
}
function isErr_(r) { return /^오류/.test(String(r[3])) || (String(r[2]) === '' && String(r[3]) === ''); }

// Jev가 '기타'라고 본 답 번호를 이 묶음 목록(key)에 붙여 기억합니다(최근 800개). 부르는 쪽이 스크립트 잠금을 잡고 있어야 합니다.
function seenAdd_(props, key, ids) {
  var seen = {};
  try { seen = JSON.parse(props.getProperty('CLASSIFY_SEEN') || '{}'); } catch (e) {}
  props.setProperty('CLASSIFY_SEEN', JSON.stringify({ key: key, ids: (seen.key === key ? seen.ids || [] : []).concat(ids).slice(-800) }));
}

// 시트에 씁니다(잠금 안). Jev를 기다리는 사이 줄이 지워지거나 정렬되었을 수 있어 번호(F)로 다시 찾고, 그사이 누가 고친 줄은 건너뜁니다.
// o: {key, q} 읽을 때의 묶음 목록·질문(그사이 바뀌었으면 쓰지 않음), cats 새 묶음 목록(주면 함께 저장), judged Jev가 기타로 본 번호, seenKey 기억에 쓸 목록.
// 바뀐 줄은 연속한 범위마다 C:G를 씁니다. partial은 쓰기 오류, needsReload는 화면 동기화 오류입니다.
function commit_(sh, changes, o) {
  var props = PropertiesService.getScriptProperties(), res = { moved: 0, aborted: '' };
  if (!changes.length && !o.cats && !(o.judged || []).length) return res;
  var lock = LockService.getScriptLock(), writing = false;
  lock.waitLock(20000);
  try {
    // 그사이 새 질문으로 넘어가 [응답] 시트가 보관되었으면 아무것도 쓰지 않습니다.
    var cur = ss_().getSheetByName('응답');
    if (!cur || cur.getSheetId() !== sh.getSheetId()) { res.aborted = '그사이 새 질문으로 넘어가 이 분류 결과는 저장하지 않았습니다.'; return res; }
    var cc = cfg_();
    if (cc.cats.join('|') !== o.key || cc.q !== o.q) { res.aborted = '그사이 질문이나 묶음이 바뀌어 이 분류 결과는 저장하지 않았습니다. [분류]를 다시 누르십시오.'; return res; }
    var now = changes.length ? sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 7).getValues() : [], at = {}, hit = [];
    now.forEach(function (r, k) { if (String(r[5]) !== '') at[r[5]] = k; });
    changes.forEach(function (ch) {
      var k = at[ch.r[5]], r = k == null ? null : now[k];
      if (!r || !isFree_(r) || String(r[2]) !== String(ch.r[2]) || String(r[3]) !== String(ch.r[3]) || String(r[6]) !== String(ch.r[6]) || String(r[4]) !== String(ch.r[4])) return;
      r[2] = ch.cat; r[3] = ch.conf; r[6] = ch.cat || AUTO_EMPTY;
      if (ch.hide) r[4] = '숨김';
      hit.push(k);
      if (ch.cat !== '기타' && !ch.hide) res.moved++;
    });
    // 바뀐 줄만 씁니다. 사이에 끼어 있는 수동 행의 수식이나 값도 그대로 둡니다.
    hit.sort(function (a, b) { return a - b; });
    for (var i = 0; i < hit.length;) {
      var lo = hit[i], hi = hit[i++];
      while (hit[i] === hi + 1) hi = hit[i++];
      writing = true;
      sh.getRange(lo + 2, 3, hi - lo + 1, 5).setValues(now.slice(lo, hi + 1).map(function (r) { return r.slice(2, 7).map(cell_); }));
    }
    if (o.cats) {
      writing = true;
      ss_().getSheetByName('설정').getRange('B2').setValue(cell_(o.cats.join(', ')));
    }
    if ((o.judged || []).length) seenAdd_(props, o.seenKey != null ? o.seenKey : o.key, o.judged);
    SpreadsheetApp.flush();
  } catch (e) {
    if (!writing) throw e;
    res.partial = true; res.moved = 0;
    try { SpreadsheetApp.flush(); } catch (ignored) {} // 앞서 성공한 쓰기는 남을 수 있습니다.
  } finally {
    try {
      if (writing) {
        try { props.setProperty('BOARD_MIX', String(Number(props.getProperty('BOARD_MIX') || 0) + 1)); }
        catch (e) { res.needsReload = true; }
        try { bump_(); } catch (e) { res.needsReload = true; }
      }
    } finally { lock.releaseLock(); }
  }
  return res;
}

function classify_(token, generation) {
  var c = cfg_(), sh = ss_().getSheetByName('응답'), t0 = Date.now(), grouped = c.cats.length >= 2, props = PropertiesService.getScriptProperties(), key = c.cats.join('|');
  var out = { ok: true, token: token, grouped: grouped, moved: 0, words: 0, jev: 0, left: 0, more: 0, stopped: '', auto: null, autoErr: '', autoNone: false, askAi: false, aiOff: props.getProperty(BOARD_AI_PROP) !== '1' };
  if (!sh) { out.board = getBoard(); out.left = grouped ? out.board.counts['기타'] || 0 : out.board.total; return out; }
  var head = sh.getRange(1, 7);
  if (String(head.getValue()) === '') head.setValue('자동 묶음').setFontWeight('bold');
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : [];
  var seen = {}, done = {};
  try { seen = JSON.parse(props.getProperty('CLASSIFY_SEEN') || '{}'); } catch (e) {}
  if (seen.key === key) (seen.ids || []).forEach(function (i) { done[i] = 1; });
  var visible = rows.filter(function (r) { return r[4] !== '숨김' && String(r[1]) !== '' && String(r[5]) !== ''; });
  var changes = [], judged = [];
  if (grouped) {
    visible.filter(function (r) { return isOther_(c, r) && isFree_(r); }).forEach(function (r) {
      var w = matchCat_(r[1], c.cats, c.q);
      if (w >= 0) { changes.push({ r: r, cat: c.cats[w], conf: '낱말' }); out.words++; return; }
      var err = isErr_(r);
      if (!err && String(r[2]) === '기타' && done[r[5]]) return;
      if (out.stopped || out.jev >= CLASSIFY_MAX || Date.now() - t0 > CLASSIFY_SEC * 1000) { out.more++; return; }
      if (!throttle_('all', CLASSIFY_RATE, 60)) { out.stopped = LIMIT_MSG; out.more++; return; }
      try {
        var j = jevCat_(c, String(r[1]), err);
        out.jev++;
        changes.push({ r: r, cat: j.cat, conf: j.conf, hide: j.bad });
        if (j.cat === '기타') judged.push(r[5]);
      } catch (e) { out.stopped = String(e && e.message || e).slice(0, 160); out.more++; }
    });
  }
  var res = commit_(sh, changes, { key: key, q: c.q, judged: judged });
  out.moved = res.moved;
  if (res.partial) out.partial = true;
  if (res.needsReload) out.needsReload = true;
  if (res.aborted) { out.stopped = out.stopped || res.aborted; changes = []; }
  function readBoard() {
    try { return getBoard(); } catch (e) {
      if (!out.partial && !out.needsReload) throw e;
      out.needsReload = true;
      return null;
    }
  }
  var board = readBoard(), left = board ? (grouped ? board.counts['기타'] || 0 : board.total) : 0;
  // (나) 묶음이 없거나 기타가 많으면 스스로 묶습니다. 위에서 막혔거나(stopped) 아직 남은 답이 있으면(more) 먼저 그것부터 끝냅니다.
  if (!out.partial && !out.needsReload && !res.aborted && !out.stopped && !out.more && (!grouped || (left >= 4 && left >= board.total * 0.25))) {
    try {
      var a = autoGroup_(c, sh, key, t0, generation);
      if (a.needsReload) out.needsReload = true;
      if (a.askAi) { out.askAi = true; out.aiOff = true; }
      else if (a.partial) out.partial = true;
      else if (a.none) { out.autoNone = true; out.autoFull = !!a.full; }
      else if (a.err) out.autoErr = a.err;
      else { out.auto = { kind: a.kind, groups: a.groups, moved: a.moved }; out.jev += a.jev; out.more += a.more; if (a.stopped) out.stopped = a.stopped; }
    } catch (e) { out.autoErr = String(e && e.message || e).slice(0, 160); }
    board = readBoard();
    left = board ? (board.cats.length ? board.counts['기타'] || 0 : board.total) : 0;
  }
  out.board = board; out.left = left;
  return out;
}

// 묶음 이름은 12글자 안: 넘으면 12글자 안에서 낱말 단위로 자르고 끝의 조사('및')를 뗍니다.
function shortName_(nm) {
  if (nm.length <= 12) return nm;
  var cut = nm.slice(0, 12), sp = cut.lastIndexOf(' ');
  return (sp >= 3 ? cut.slice(0, sp) : cut).replace(/\s*(및|와|과|의|,)$/, '').trim();
}

// 연락처 전체를 같은 길이로 가린 뒤 자릅니다. 80자 경계에 걸친 연락처 일부도 보내지 않습니다.
function maskContact_(text) {
  var mask = function (s) { return '*'.repeat(s.length); };
  return text.replace(/[\p{L}\p{N}._%+'-]+@[\p{L}\p{N}.-]+\.[\p{L}]{2,}/gu, mask)
    .replace(/(^|[^\d+])(?:(?:\+?82[\s.-]*)?(?:\(?0\d{1,2}\)?|10)[\s.-]*\d{3,4}[\s.-]*\d{4}|\+\d{1,3}(?:[\s.-]+\(?\d{1,4}\)?){1,4}[\s.-]+\d{3,4}|\d{4}[\s.-]+\d{4})(?!\d)/g,
      function (s, prefix) { return prefix + mask(s.slice(prefix.length)); });
}

// 스스로 묶기. {none:true}(묶을 답이 적음) | {askAi:true}(동의 없음) | {partial:true} | {err} | {kind, groups, moved, jev, more, stopped}. 화면 동기화 오류는 needsReload로 따로 알립니다.
// 묶을 답: 숨김이 아니고 건드려도 되며, 묶음이 있으면 그 밖의 답. 쓰기는 commit_ 한 번(묶음 목록 저장 + 답 옮기기)입니다.
function autoGroup_(c, sh, key, t0, generation) {
  var room = Math.min(6, 10 - c.cats.length), grouped = c.cats.length >= 2;
  var pool = (sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : []).filter(function (r) {
    return r[4] !== '숨김' && String(r[1]) !== '' && String(r[5]) !== '' && isFree_(r) && (!grouped || isOther_(c, r));
  });
  if (room < 2) return { none: true, full: true };
  if (pool.length < 4) return { none: true };
  var texts = pool.map(function (r) { return String(r[1]); }), assign = {}, names = [], kind = 'words';
  var cl = clusterAnswers_(texts, c.q, room, c.cats), covered = 0;
  cl.forEach(function (g) { covered += g.idx.length; });
  if (cl.length >= 2 && covered >= pool.length * AUTO_SHARE) {
    cl.forEach(function (g) { names.push(g.name); g.idx.forEach(function (i) { assign[i] = g.name; }); });
  } else {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty(BOARD_AI_PROP) !== '1') return { askAi: true };
    kind = 'ai';
    // AI는 기다림이 길어(재시도 포함 수 분) 6분 한도를 넘지 않게 시작 1분 안에만 부릅니다.
    if (Date.now() - t0 > Math.min(60, CLASSIFY_SEC) * 1000) return { err: '시간이 모자라 AI로 묶지 못했어요. [분류]를 한 번 더 누르세요.' };
    if (!throttle_('all', CLASSIFY_RATE, 60) || !throttle_('autogroup', 3, 60)) return { err: '잠시 뒤에 다시 눌러 주세요. (' + LIMIT_MSG + ')' };
    var recent = pool.map(function (r, i) { return i; }).sort(function (a, b) { return Number(pool[a][5]) - Number(pool[b][5]); }).slice(-AUTO_LIST);
    var list = recent.map(function (i) { return texts[i]; });
    var prompt = '아래는 수업에서 학생 ' + list.length + '명이 쓴 "' + maskContact_(c.q) + '"에 대한 답이야. 하나도 빠뜨리지 말고 비슷한 것끼리 3~' + Math.max(3, room) + '개 묶음으로 정리해 줘.\n' +
      '형식을 정확히 지켜. 한 줄에 묶음 하나, 다른 말은 쓰지 마.\n묶음 이름 | 답 번호들(쉼표로 구분)\n' +
      '규칙:\n- 묶음 이름은 공백 포함 12글자 안쪽입니다.\n- 번호는 아래 목록의 번호입니다. 모든 번호를 어느 한 묶음에 넣으세요.\n- "기타"라는 묶음은 만들지 않습니다.' +
      (c.cats.length ? '\n- 이미 있는 묶음 이름(' + maskContact_(c.cats.join(', ')) + ')은 다시 쓰지 않습니다.' : '') + '\n\n[답 목록]\n' +
      list.map(function (t, i) { return (i + 1) + '. ' + maskContact_(t.replace(/\s+/g, ' ')).slice(0, 80); }).join('\n');
    var raw, withdrawn = {};
    try {
      raw = ai_(prompt, { beforeSend: function () {
        // 완성된 본문을 보내기 직전(재시도 포함)에만 끄기와 승인 읽기를 직렬화합니다.
        var lock = LockService.getScriptLock();
        if (!lock.tryLock(1000)) throw new Error('전송 동의를 확인하지 못했습니다. 잠시 후 다시 시도하십시오.');
        try {
          if (props.getProperty(BOARD_AI_PROP) !== '1' || (props.getProperty(BOARD_AI_GEN) || '0') !== generation) throw withdrawn;
        } finally { lock.releaseLock(); }
      } });
    } catch (e) {
      if (e === withdrawn) return { askAi: true };
      return { err: 'AI가 묶음을 만들지 못했습니다. ' + String(e && e.message || e).slice(0, 160) };
    }
    var used = {};
    String(raw).split('\n').forEach(function (line) {
      // Markdown 표('| 이름 | 1, 2 |')로 답해도 읽도록 줄 앞뒤의 '|'를 먼저 뗍니다.
      var m = /^\s*(?:#+\s*|[-*•·]\s*|\d+[.)]\s*)?(.+?)\s*[|｜]\s*(.+)$/.exec(line.replace(/^\s*[|｜]\s*|\s*[|｜]\s*$/g, ''));
      if (!m || names.length >= room) return;
      var nm = shortName_(cats_(m[1].replace(/\*\*/g, ''))[0] || ''), nums = (m[2].match(/\d+/g) || []).map(Number).filter(function (n, k, a) { return n >= 1 && n <= list.length && !used[n] && a.indexOf(n) === k; });
      if (!nm || c.cats.indexOf(nm) >= 0 || names.indexOf(nm) >= 0 || nums.length < 2) return;
      names.push(nm); nums.forEach(function (n) { used[n] = 1; assign[recent[n - 1]] = nm; });
    });
    if (names.length < 2) return { err: 'AI 답에서 묶음을 찾지 못했습니다. 다시 눌러 주세요.' };
  }
  var cats = c.cats.concat(names).slice(0, 10), c2 = { q: c.q, cats: cats }, changes = [], judged = [], jev = 0, more = 0, stopped = '', moved = 0;
  pool.forEach(function (r, i) {
    if (assign[i]) { changes.push({ r: r, cat: assign[i], conf: kind === 'ai' ? 'AI' : '낱말' }); return; }
    if (kind !== 'ai') return;
    // AI가 빠뜨린 답만 Jev로 나눕니다(새 묶음 목록으로).
    if (stopped || jev >= CLASSIFY_MAX || Date.now() - t0 > CLASSIFY_SEC * 1000) { more++; return; }
    if (!throttle_('all', CLASSIFY_RATE, 60)) { stopped = LIMIT_MSG; more++; return; }
    try {
      var j = jevCat_(c2, String(r[1]), isErr_(r));
      jev++;
      changes.push({ r: r, cat: j.cat, conf: j.conf, hide: j.bad });
      if (j.cat === '기타') judged.push(r[5]);
    } catch (e) { stopped = String(e && e.message || e).slice(0, 160); more++; }
  });
  var res = commit_(sh, changes, { key: key, q: c.q, cats: cats, judged: judged, seenKey: cats.join('|') });
  if (res.partial) return { partial: true, needsReload: res.needsReload };
  if (res.aborted) return { err: res.aborted };
  return { kind: kind, groups: names.length, moved: res.moved, jev: jev, more: more, stopped: stopped, needsReload: res.needsReload };
}

// 큰 화면의 "AI로 묶음 만들기": 질문을 보고 묶음 후보를 줄 단위로 돌려줍니다. 저장은 하지 않고, 교수가 고쳐서 바꾸기를 눌러야 반영됩니다.
function suggestCats(cred, q) {
  var chk = auth_(cred, 'board');
  if (!chk.ok) return chk;
  q = String(q || '').replace(/\s+/g, ' ').trim().slice(0, 160);
  if (q.length < 2) return { ok: false, msg: '질문을 먼저 적어 주세요.', token: chk.token };
  if (!throttle_('suggest', 12, 60)) return { ok: false, msg: '잠시 뒤에 다시 눌러 주세요.', token: chk.token };
  var text;
  try {
    text = ai_('수업 중 학생들에게 던질 질문입니다.\n질문: ' + q + '\n\n학생들의 짧은 답을 모아 보기 좋게 나눌 묶음 이름을 4~6개 만드세요. 규칙:\n- 한 줄에 하나, 이름만 씁니다. 번호, 기호, 설명을 붙이지 않습니다.\n- 이름은 공백 포함 12자 이내입니다.\n- 서로 겹치지 않게, 이 질문에 나올 만한 답을 고르게 덮도록 합니다.\n- "기타"는 쓰지 않습니다(따로 붙습니다).');
  } catch (err) { return { ok: false, msg: 'AI가 묶음을 만들지 못했습니다. ' + String(err && err.message || err).slice(0, 160), token: chk.token }; }
  var list = cats_(text.split('\n').map(function (x) { return x.replace(/^\s*(?:[-*•·]|\d+[.)])\s*/, ''); }).join('\n'));
  if (list.length < 2) return { ok: false, msg: 'AI 답에서 묶음을 찾지 못했습니다. 다시 눌러 주세요.', token: chk.token };
  return { ok: true, cats: list.join('\n'), token: chk.token };
}

// 큰 화면에서 질문과 묶음을 바로 바꿉니다. cred는 교수 PIN, 또는 한 번 맞힌 뒤 받은 화면용 표입니다.
// fresh가 참이면 지금까지의 응답을 먼저 보관하고 화면을 비웁니다(보관이 실패하면 질문도 바꾸지 않습니다).
function setBoard(cred, q, cats, open, fresh) {
  var chk = auth_(cred, 'board');
  if (!chk.ok) return chk;
  q = String(q || '').replace(/\s+/g, ' ').trim().slice(0, 160);
  if (q.length < 2) return { ok: false, msg: '질문을 두 글자 이상 적어 주세요.', token: chk.token };
  var list = cats_(cats);
  if (list.length === 1) return { ok: false, msg: '묶음은 두 개 이상 적거나, 묶지 않으려면 비워 두세요.', token: chk.token };
  var lock = LockService.getScriptLock(), archived = false;
  lock.waitLock(20000);
  try {
    if (fresh) archived = archive_();
    ss_().getSheetByName('설정').getRange('B1:B3').setValues([[cell_(q)], [cell_(list.join(', '))], [open ? 'O' : 'X']]);
    bump_();
  } finally { lock.releaseLock(); }
  return { ok: true, archived: archived, token: chk.token, board: getBoard() };
}

// 큰 화면용 자료: 묶음별 개수는 숨김을 뺀 모든 응답으로 세고, 글은 최근 240개만 보냅니다. 읽기만 하고 시트를 만들지 않습니다.
// 묶음 이름이 바뀌어 지금 목록에 없는 예전 답은 '기타'로 셉니다. round는 [응답] 시트가 보관으로 바뀌면 달라집니다.
function getBoard() {
  var cache = CacheService.getScriptCache(), hit = cache.get('board');
  if (hit) return JSON.parse(hit);
  // rev는 읽기 전에 잡습니다. 읽는 사이 질문이 바뀌면 이 자료는 예전 번호를 달고 나가 화면이 버리게 되고, 기억하지도 않습니다.
  var rev = rev_(), c = cfg_(), sh = ss_().getSheetByName('응답'), n = sh ? sh.getLastRow() - 1 : 0, tz = Session.getScriptTimeZone();
  var rows = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [], items = [];
  var cats = c.cats.length >= 2 ? c.cats.concat(['기타']) : [], counts = {};
  cats.forEach(function (k) { counts[k] = 0; });
  var bad = 0, why = '';
  rows.forEach(function (r, k) {
    if (r[4] === '숨김' || String(r[1]) === '') return;
    if (cats.length && /^오류/.test(String(r[3]))) { bad++; why = String(r[3]).replace(/^오류:?\s*/, '').slice(0, 60); }
    var cat = '';
    if (cats.length) { cat = String(r[2] || ''); if (cats.indexOf(cat) < 0) cat = '기타'; counts[cat]++; }
    items.push({ i: Number(r[5]) || k + 2, no: k + 1, h: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'HH:mm') : '', t: String(r[1]), c: cat });
  });
  // mix: [분류]가 답을 옮길 때마다 1씩 올라갑니다. 화면은 이 값이 바뀌면 막대를 처음부터 다시 그립니다(옮겨진 칸이 제 묶음으로 가게).
  var out = { q: c.q, cats: cats, counts: counts, open: c.open, refresh: c.refresh, total: items.length,
    round: sh ? sh.getSheetId() : 0, rev: rev, mix: Number(PropertiesService.getScriptProperties().getProperty('BOARD_MIX') || 0), items: items.slice(-240) };
  // err: 판정에 실패해 '오류:' 흔적이 남은 답의 수와 가장 최근 이유(큰 화면이 한 번 알립니다). 학생 폰 화면은 getBoard를 부르지 않습니다.
  if (bad) out.err = { n: bad, why: why };
  if (rev_() === rev) cache.put('board', JSON.stringify(out), 3);
  return out;
}

// 학생 폰 화면. 큰 화면과 같은 말씨: 종이 바탕 위의 흰 종이 한 장(슬립)에 생각을 적어 보냅니다.
// 보내면 그 종이가 위로 날아가고, 아래 '내가 보낸 생각'에 번호와 묶음이 찍힌 종이로 남습니다. 큰 화면의 칸에도 같은 번호가 붙습니다.
var ASK_CSS = [
  ':root{--bg:#ffffff;--ink:#111111;--mut:#6e6e6e;--hair:rgba(17,17,17,.14);--slip:#fff;--tile:#f2f2f0;--red:#e54536;--e:cubic-bezier(.2,.8,.2,1)}',
  '*{box-sizing:border-box}html{-webkit-text-size-adjust:100%}',
  'body{margin:0;min-height:100vh;background:var(--bg);color:var(--ink);font:16px/1.5 Pretendard,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;-webkit-font-smoothing:antialiased;word-break:keep-all}',
  '.ask{max-width:560px;margin:0 auto;padding:clamp(18px,5vw,40px) clamp(18px,5vw,28px) 80px}',
  '.bar{display:flex;justify-content:space-between;align-items:baseline;padding-bottom:12px;border-bottom:1px solid var(--ink);font:700 11px/1 Montserrat,Pretendard,sans-serif;letter-spacing:.2em}',
  '.bar span:last-child{letter-spacing:.06em;font-family:Pretendard,sans-serif;font-size:13px}.bar .shut{color:var(--red)}',
  '.qk{margin:26px 0 8px;font:800 12px/1 Montserrat,sans-serif;letter-spacing:.2em;color:var(--red)}',
  'h1{margin:0 0 26px;font-size:clamp(25px,7.4vw,36px);font-weight:800;letter-spacing:-.035em;line-height:1.2;text-wrap:balance}',
  '.w{display:inline-block;overflow:hidden;vertical-align:top;padding-bottom:.1em;margin-bottom:-.1em}',
  '.w>span{display:inline-block;transform:translateY(108%);transition:transform .8s cubic-bezier(.2,.7,.1,1)}.on .w>span{transform:none}',
  '.card{position:relative;background:var(--slip);border:1px solid var(--ink);border-radius:2px;transition:transform .5s var(--e),opacity .4s}',
  '.card:focus-within{outline:2px solid var(--ink);outline-offset:-1px}',
  '.card.go{transform:translateY(-46px) rotate(-1.2deg);opacity:0;transition:transform .55s cubic-bezier(.55,0,.3,1),opacity .45s ease .1s}',
  '.card.come{transform:translateY(18px);opacity:0;transition:none}',
  '.card.shut{opacity:.45}',
  '.ch{display:flex;justify-content:space-between;padding:12px 16px 0;font:700 10px/1 Montserrat,Pretendard,sans-serif;letter-spacing:.16em;color:var(--mut)}',
  '.ch span:last-child{font-family:Pretendard,sans-serif;letter-spacing:.02em;font-weight:500;font-size:12px}',
  'textarea{display:block;width:100%;min-height:8.5em;margin:0;padding:10px 16px 14px;border:0;background:transparent;color:var(--ink);font:500 19px/1.5 Pretendard,sans-serif;resize:none;outline:none}',
  'textarea::placeholder{color:#a8a49c}',
  '.chips{display:flex;flex-wrap:wrap;gap:6px;padding:0 16px 12px}.chips[hidden]{display:none}.chips p{flex:none;width:100%;margin:0 0 2px;font:600 12px/1.3 Pretendard,sans-serif;color:var(--mut)}',
  '.chips button{padding:8px 12px;border:1px solid var(--ink);border-radius:2px;background:transparent;color:var(--ink);font:600 14px/1.2 Pretendard,sans-serif;cursor:pointer}',
  '.chips button.on{background:var(--ink);color:#fff}.chips button:disabled{opacity:.4;cursor:default}',
  '.cf{display:flex;align-items:center;gap:12px;padding:0 16px 14px}',
  '.meter{position:relative;flex:1;height:1px;background:var(--hair)}.meter i{position:absolute;left:0;top:-1px;width:100%;height:3px;background:var(--ink);transform-origin:0 50%;transform:scaleX(0);transition:transform .25s var(--e)}.meter.full i{background:var(--red)}',
  '.cnt{font:600 12px/1 Montserrat,sans-serif;letter-spacing:.04em;color:var(--mut);font-variant-numeric:tabular-nums}',
  '#b{display:flex;justify-content:space-between;align-items:center;width:100%;margin-top:12px;padding:18px 20px;border:1px solid var(--ink);border-radius:2px;background:var(--ink);color:#fff;font:700 17px/1 Pretendard,sans-serif;cursor:pointer;transition:transform .15s,opacity .3s}',
  '#b:active{transform:translateY(1px)}#b:disabled{opacity:.35;cursor:default}#b i{font-style:normal;font-family:Montserrat,sans-serif;transition:transform .3s var(--e)}#b:not(:disabled):hover i{transform:translateX(4px)}',
  '.ai-notice{margin:10px 16px 0;color:var(--mut);font-size:14px;line-height:1.5}',
  '#m{min-height:1.5em;margin:14px 0 0}.msg{margin:0;font-size:15px;font-weight:600;line-height:1.45}.msg b{color:var(--red)}.msg.err{color:var(--red)}',
  '.msg.rise,.sent.rise{animation:rise .55s var(--e) both}@keyframes rise{from{opacity:0;transform:translateY(12px)}}',
  '.mine{margin-top:34px}.mine[hidden]{display:none}',
  '.mine h2{display:flex;justify-content:space-between;margin:0 0 10px;padding-top:10px;border-top:1px solid var(--ink);font:700 11px/1 Montserrat,Pretendard,sans-serif;letter-spacing:.2em}',
  '.mine h2 span:first-child{font-family:Pretendard,sans-serif;font-size:13px;letter-spacing:-.01em}.mine h2 span:last-child{letter-spacing:.04em}',
  '.sent{display:flex;margin:0 0 3px;background:var(--tile);font-size:16px;line-height:1.45}',
  '.sent .no{flex:none;display:flex;align-items:center;padding:12px 12px;border-right:1px solid var(--hair);font:700 11px/1 Montserrat,sans-serif;letter-spacing:.06em;color:var(--mut);font-variant-numeric:tabular-nums}',
  '.sent p{flex:1;margin:0;padding:11px 14px;overflow-wrap:anywhere}',
  '.sent small{display:block;margin-top:3px;font-size:12px;font-weight:700;color:var(--mut)}.sent:first-child small{color:var(--red)}',
  '@media (prefers-reduced-motion:reduce){*{transition-duration:.01ms!important;animation-duration:.01ms!important}}'
].join('');

function askClient_() {
  var $ = function (s) { return document.querySelector(s); };
  var live = typeof requestAnimationFrame === 'function', sid, cq = null, round = 0, sending = false, open = true, lt = 0, pick = '', chipKey = '', chipBtns = [], mine = { q: '', round: 0, list: [] };
  try { sid = localStorage.getItem('sid'); if (!sid) { sid = Math.random().toString(36).slice(2, 12); localStorage.setItem('sid', sid); } } catch (e) { sid = Math.random().toString(36).slice(2, 12); }
  try { mine = JSON.parse(localStorage.getItem('board-mine') || '') || mine; } catch (e) {}
  function mk(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function pad(n) { return n < 1000 ? ('00' + n).slice(-3) : String(n); }
  function cls(e, c, on) { if (e.classList) e.classList.toggle(c, on); else e.className = on ? c : ''; }
  // 질문이 바뀌었거나 마감 같은 상태 안내만 여기에 적습니다. 버튼 결과는 공통 토스트(toast, fail)로 알립니다.
  function note(t, bad) { var m = $('#m'), p = mk('p', 'msg rise' + (bad ? ' err' : ''), t); m.textContent = ''; m.appendChild(p); }
  function words(e, text) {
    if (!live) { e.textContent = text; return; }
    e.textContent = ''; cls(e, 'on', false);
    String(text).split(/\s+/).filter(String).forEach(function (w, k) {
      if (k) e.appendChild(document.createTextNode(' '));
      var o = mk('span', 'w'), i = mk('span', '', w); i.style.transitionDelay = (k * 50) + 'ms'; o.appendChild(i); e.appendChild(o);
    });
    requestAnimationFrame(function () { requestAnimationFrame(function () { cls(e, 'on', true); }); });
  }
  function count() {
    var n = $('#t').value.length;
    $('#n').textContent = n + ' / 140';
    if ($('#mt').style) $('#mt').style.transform = 'scaleX(' + Math.min(1, n / 140).toFixed(3) + ')';
    cls($('#meter'), 'full', n >= 130);
  }
  function drawMine(fresh) {
    var box = $('#mine'), list = $('#ml'); list.textContent = '';
    box.hidden = !mine.list.length;
    $('#mc').textContent = mine.list.length + '개';
    mine.list.slice().reverse().forEach(function (x, k) {
      var e = mk('div', 'sent' + (fresh && k === 0 ? ' rise' : '')), p = mk('p', '', x.t);
      e.appendChild(mk('span', 'no', x.no ? pad(x.no) : '—'));
      p.appendChild(mk('small', '', (x.c ? '→ ' + x.c : '보냄') + (x.h ? ' · ' + x.h : '')));
      e.appendChild(p); list.appendChild(e);
    });
  }
  // 묶음 칩: 누르면 그 묶음으로 보내고, 한 번 더 누르면 풉니다. 아무것도 안 고르면 AI가 묶음을 고릅니다.
  function drawChips(cats) {
    var key = cats.join('|');
    if (key === chipKey) return;
    var keep = pick && (pick === '기타' || cats.indexOf(pick) >= 0) ? pick : '';
    if (pick && !keep) note('묶음이 바뀌어 고른 묶음을 풀었습니다.');
    chipKey = key; pick = keep;
    var box = $('#chips'); box.textContent = ''; chipBtns = [];
    box.hidden = !cats.length;
    if (!cats.length) return;
    box.appendChild(mk('p', '', '묶음을 골라 보낼 수 있어요 (안 골라도 됩니다)'));
    cats.concat(['기타']).forEach(function (name) {
      var b = mk('button', '', name); b.type = 'button';
      b.onclick = function () {
        pick = pick === name ? '' : name;
        chipBtns.forEach(function (x) { cls(x, 'on', x.textContent === pick); });
      };
      cls(b, 'on', name === pick); b.disabled = !open; chipBtns.push(b); box.appendChild(b);
    });
  }
  function clearPick() {
    pick = '';
    chipBtns.forEach(function (x) { cls(x, 'on', false); });
  }
  function save() { try { localStorage.setItem('board-mine', JSON.stringify(mine)); } catch (e) {} }
  function setOpen(o) {
    open = o;
    var st = $('#st'); st.textContent = o ? '받는 중' : '마감'; st.className = o ? '' : 'shut';
    cls($('#card'), 'shut', !o);
    chipBtns.forEach(function (x) { x.disabled = !o; });
    if (!sending) $('#b').disabled = !o;
  }
  function load() {
    google.script.run.withSuccessHandler(function (r) {
      if (cq !== null && r.q !== cq) { note('질문이 바뀌었습니다. 새 질문에 답해 주세요.'); }
      if (r.q !== cq) { words($('#q'), r.q); if (cq !== null) clearPick(); }
      drawChips(r.cats || []);
      if (mine.q !== r.q || mine.round !== r.round) { mine = { q: r.q, round: r.round, list: [] }; save(); }
      if (r.q !== cq || r.round !== round) drawMine(false);
      round = r.round;
      if (cq !== null && open && !r.open) note('지금은 답을 받지 않습니다.', 1);
      if (cq !== null && !open && r.open) note('다시 답을 받습니다.');
      if (cq === null && !r.open) note('지금은 답을 받지 않습니다.', 1);
      cq = r.q; setOpen(r.open);
      clearTimeout(lt); lt = setTimeout(load, 12000);
    }).withFailureHandler(function () { clearTimeout(lt); lt = setTimeout(load, 20000); }).getQuestion();
  }
  $('#t').oninput = count;
  $('#b').onclick = function () {
    var t = $('#t').value;
    if (sending) return;
    if (t.trim().length < 2) { fail('생각 보내기를 못했어요. 두 글자 이상 적어 주세요.'); return; }
    sending = true; $('#b').disabled = true;
    google.script.run.withSuccessHandler(function (r) {
      sending = false; $('#b').disabled = !open;
      if (!r.ok) { fail('생각 보내기를 못했어요. ' + r.msg); if (r.stale) load(); return; }
      var now = new Date(), h = ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
      mine.list.push({ t: t.replace(/\s+/g, ' ').trim().slice(0, 140), c: r.cat || '', no: r.no || 0, h: h }); mine.list = mine.list.slice(-20); save();
      $('#m').textContent = '';
      toast('생각을 보냈어요' + (r.no ? ' · No.' + pad(r.no) : '') + (r.cat ? ' · 「' + r.cat + '」 쪽으로 모였어요.' : '. 큰 화면을 보세요.'));
      var card = $('#card');
      function reset() { $('#t').value = ''; clearPick(); count(); drawMine(true); }
      if (!live) { reset(); return; }
      cls(card, 'go', true);
      setTimeout(function () {
        reset(); cls(card, 'go', false); cls(card, 'come', true);
        requestAnimationFrame(function () { requestAnimationFrame(function () { cls(card, 'come', false); }); });
      }, 520);
    }).withFailureHandler(function (e) { sending = false; $('#b').disabled = !open; fail('생각 보내기를 못했어요. ' + e.message); }).submitAnswer(t, sid, round, pick);
  };
  count(); drawMine(false); load();
}

var BOARD_BATCH_NOTICE = '새 묶음을 만들기 위해 질문·기존 묶음 이름과 응답 최대 80건의 앞 80자씩을 ' + BOARD_AI_NAME + '에 한꺼번에 보냅니다. 이때 답의 이메일·전화번호 패턴을 가립니다.';
var BOARD_SINGLE_NOTICE = '기존 묶음에 분류하거나 광고·욕설을 확인할 때, 또는 일괄 분류에서 빠진 답을 나눌 때는 ' +
  (BOARD_SOLO ? 'Gemini' : '수업 AI(' + BOARD_AI_NAME + '의 Jev)') + '에 답 원문 전체를 가리지 않고 한 건씩 보낼 수 있습니다. 답에 적은 이름도 전송될 수 있습니다.';
var ASK_HTML = '<main class="ask"><div class="bar"><span>LIVE BOARD</span><span id="st">받는 중</span></div>' +
  '<p class="qk">Q.</p><h1 id="q">…</h1>' +
  '<div class="card" id="card"><div class="ch"><span>MY THOUGHT</span><span>이름 입력란 없음</span></div>' +
  '<p class="ai-notice" id="ai-notice">교수가 동의하면 ' + BOARD_BATCH_NOTICE + ' ' + BOARD_SINGLE_NOTICE + ' 이름·연락처는 적지 마십시오.</p>' +
  '<textarea id="t" maxlength="140" aria-describedby="ai-notice" placeholder="한두 문장으로 적어 주세요"></textarea><div class="chips" id="chips" hidden></div>' +
  '<div class="cf"><div class="meter" id="meter"><i id="mt"></i></div><span class="cnt" id="n">0 / 140</span></div></div>' +
  '<button id="b">보내기<i>→</i></button><div id="m"></div>' +
  '<section class="mine" id="mine" hidden><h2><span>내가 보낸 생각</span><span id="mc">0개</span></h2><div id="ml"></div></section></main>' +
  '<script>(' + askClient_.toString() + ')();</script>';

// 강의장 큰 화면. 답 하나가 한 칸(종이 띠)이고, 묶음마다 띠가 쌓여 막대가 됩니다. 막대 높이가 곧 답의 수입니다.
// 새 답은 먼저 화면 위쪽에 크게 한 단어씩 나타나고, 묶음이 정해지면 줄어들며 그 막대의 맨 위 칸으로 들어갑니다.
// 답이 많아 칸이 얇아지면 글 대신 선으로 쌓고, 막대 위에 가장 최근 답 한 줄을 남깁니다. 강조색(빨강)은 가장 많은 묶음의 숫자에만 씁니다.
var BOARD_CSS = [
  ':root{--bg:#ffffff;--fg:#111111;--ink:#111111;--mut:#6e6e6e;--hair:rgba(17,17,17,.14);--slip:#f1f1ef;--bar:#cfcfcb;--red:#e54536;--e:cubic-bezier(.2,.8,.2,1)}',
  'body.dark{--bg:#121212;--fg:#efede8;--ink:#efede8;--mut:#8f8b83;--hair:rgba(239,237,232,.18);--slip:#232322;--bar:#5d5b56}',
  '*{box-sizing:border-box}html,body{height:100%}',
  'body{margin:0;overflow:hidden;background:var(--bg);color:var(--ink);font:16px/1.5 Pretendard,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;-webkit-font-smoothing:antialiased;transition:background-color .6s,color .6s}',
  '.stage{position:fixed;inset:0;display:grid;grid-template-rows:auto minmax(0,1fr);gap:clamp(12px,2.6vh,34px);padding:clamp(18px,3vw,58px) clamp(18px,3.4vw,66px) clamp(16px,2.6vw,44px)}',
  '.top{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:clamp(16px,3vw,56px);align-items:start}',
  '.meta{display:flex;flex-wrap:wrap;gap:.4em 1.4em;font-size:clamp(12px,1.05vw,18px);font-weight:500;color:var(--mut);font-variant-numeric:tabular-nums}',
  '.meta b{color:var(--ink);font-weight:700}.meta .shut{color:var(--red)}',
  'h1{margin:.4em 0 0;max-width:16em;font-size:clamp(28px,4.3vw,84px);font-weight:800;letter-spacing:-.04em;line-height:1.1;word-break:keep-all;text-wrap:balance}',
  '.w{display:inline-block;overflow:hidden;vertical-align:top;padding-bottom:.1em;margin-bottom:-.1em}',
  '.w>span{display:inline-block;transform:translateY(108%);transition:transform .85s cubic-bezier(.2,.7,.1,1)}',
  '.on .w>span{transform:none}.out .w>span{transform:translateY(-108%);transition-duration:.42s;transition-timing-function:cubic-bezier(.6,0,.8,.2);transition-delay:0s!important}',
  '.join{display:flex;align-items:flex-end;gap:clamp(10px,1vw,18px)}.join[hidden]{display:none}',
  '.join p{margin:0;text-align:right;font-size:clamp(12px,1.05vw,18px);line-height:1.35;color:var(--mut)}.join p b{display:block;color:var(--ink);font-size:1.25em}',
  '.qr{width:clamp(86px,8.6vw,156px);aspect-ratio:1;padding:6px;background:#fff;border:1px solid var(--ink)}.qr svg{display:block;width:100%;height:100%}.qr path{fill:#141414}',
  '.field{position:relative;display:grid;grid-template-rows:minmax(0,1fr) auto auto;min-height:0}',
  '.cols,.labs{display:grid;grid-template-columns:repeat(var(--n,5),minmax(0,1fr));gap:0 clamp(10px,1.5vw,28px)}',
  '.cols{min-height:0}',
  '.labs{padding:0 0 clamp(40px,6.5vh,64px)}', // 오른쪽 아래 단추 줄이 마지막 막대 이름을 가리지 않게 이름을 위로 올립니다
  '.wall .toast{bottom:auto;top:clamp(150px,34vh,360px);z-index:60}', // 큰 화면 알림은 막대 이름과 겹치지 않는 가운데 빈 자리에 띄웁니다
  '.col{display:flex;flex-direction:column;justify-content:flex-end;min-height:0;overflow:hidden}',
  '.stack{display:flex;flex-direction:column;transform-origin:50% 100%;transition:transform 1s var(--e)}',
  '.intro .stack{transform:scaleY(0);transition-duration:.5s}',
  '.slip{flex:none;display:flex;align-items:center;height:var(--u,64px);padding:0 .8em;overflow:hidden;background:var(--slip);color:var(--ink);border-top:3px solid var(--bg);font-size:var(--fs,20px);font-weight:500;line-height:1.3;transition:height .6s var(--e),background-color 1.1s,color 1.1s}',
  '.col.lead .slip{background:var(--ink);color:var(--bg)}',
  '.slip span{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:var(--lines,2);overflow:hidden;word-break:keep-all;overflow-wrap:anywhere;transition:opacity .35s}',
  '.slip.g0{height:0;border-top-width:0}.slip.hid span,.slip.hid .no{opacity:0}',
  '.slip .no{order:-1;flex:none;align-self:stretch;display:flex;align-items:center;margin-right:.7em;padding-right:.7em;border-right:1px solid var(--hair);font:700 calc(var(--fs,20px)*.52)/1 Montserrat,Pretendard,sans-serif;font-style:normal;letter-spacing:.06em;color:var(--mut);font-variant-numeric:tabular-nums;transition:opacity .35s,color 1.1s}',
  '.col.lead .slip .no{color:var(--bg);opacity:.55;border-right-color:var(--bg)}.slip.fresh .no{color:#fff;opacity:.85;border-right-color:#fff}.thin .slip .no{display:none}',
  '.slip.fresh,.col.lead .slip.fresh{background:var(--red);color:#fff;transition-duration:.25s}',
  '.thin .slip{padding:0;background:var(--bar);border-top-width:1px}.thin .slip span{display:none}.thin .col.lead .slip{background:var(--ink)}',
  '.older{flex:none;display:flex;align-items:center;justify-content:center;height:0;overflow:hidden;background:var(--bar);color:var(--ink);font-size:clamp(11px,.9vw,15px);font-weight:600;transition:height .6s var(--e)}.col.lead .older{background:var(--ink);color:var(--bg)}',
  '.call{margin:0 0 .55em;padding-left:.6em;border-left:2px solid var(--ink);font-size:var(--fs,18px);line-height:1.3;word-break:keep-all;overflow-wrap:anywhere;display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:3;overflow:hidden}',
  '.cols:not(.thin) .call{display:none}.call:empty{display:none}',
  '.base{height:2px;background:var(--ink);transform-origin:0 50%;transition:transform 1.1s cubic-bezier(.7,0,.2,1)}.intro .base{transform:scaleX(0);transition-duration:.4s}',
  '.cnt{flex:none;margin:0 0 .18em;font-size:calc(var(--fs,20px)*2.3);font-weight:800;line-height:1;letter-spacing:-.04em;font-variant-numeric:tabular-nums;display:flex}',
  '.intro .cnt{opacity:0}.cnt{transition:opacity .6s,color .6s}.col.lead .cnt{color:var(--red)}',
  '.lab{padding-top:clamp(8px,1.2vh,16px);transition:opacity .7s,transform .7s var(--e)}',
  '.intro .lab{opacity:0;transform:translateY(10px)}',
  '.lab b{font-size:var(--fl,22px);font-weight:700;letter-spacing:-.02em;line-height:1.2;word-break:keep-all;min-width:0}',
  '#c{display:inline-flex;vertical-align:top;margin-top:.25em;line-height:1}',
  '.dg{display:inline-block;height:1em;overflow:hidden}.st{display:flex;flex-direction:column;transition:transform .8s var(--e)}.st span{display:block;height:1em;line-height:1}',
  '.head{position:relative}.band{position:absolute;left:0;right:0;top:0;z-index:4;min-height:100%;background:var(--bg);transition:background-color .5s}',
  '.band[hidden]{display:none}.band.fly{background-color:transparent}',
  '.band small{display:flex;gap:.6em;font-size:clamp(12px,1.05vw,18px);font-weight:600;color:var(--mut);transition:opacity .3s}.band.fly small{opacity:0}.band small b{color:var(--red)}',
  '.band p{margin:.4em 0 0;max-width:17em;font-size:clamp(28px,4vw,78px);font-weight:800;letter-spacing:-.04em;line-height:1.1;word-break:keep-all;transform-origin:0 0}',
  '.top h1{transition:opacity .3s}.top.arriving h1{opacity:0}',
  '.wait{position:absolute;left:0;top:8%;max-width:15em;margin:0;font-size:clamp(20px,2.6vw,48px);font-weight:700;letter-spacing:-.03em;line-height:1.25;color:var(--mut)}.wait[hidden]{display:none}',
  '.edit{position:fixed;top:0;right:0;bottom:0;z-index:40;width:min(520px,94vw);overflow:auto;padding:clamp(22px,2.6vw,40px);background:var(--bg);border-left:1px solid var(--ink);transform:translateX(102%);visibility:hidden;transition:transform .55s var(--e),visibility 0s .55s}',
  '.edit.on{transform:none;visibility:visible;transition:transform .55s var(--e),visibility 0s}',
  '.edit h2{margin:0 0 1.1em;font-size:24px;font-weight:800;letter-spacing:-.03em}',
  '.edit label[hidden]{display:none}.edit label{display:block;margin:0 0 16px;font-size:14px;font-weight:700}.edit label small{display:block;margin:2px 0 6px;font-weight:500;color:var(--mut)}',
  '.edit textarea,.edit input[type=text],.edit input[type=password]{display:block;width:100%;margin-top:6px;padding:10px 12px;border:1px solid var(--ink);border-radius:2px;background:transparent;color:var(--ink);font:500 16px/1.45 Pretendard,sans-serif;resize:vertical}',
  '.edit .ck{display:flex;align-items:center;gap:10px;font-weight:600}.edit .ck input{width:18px;height:18px;margin:0;accent-color:#141414}',
  '.edit .row,.ai-confirm .row{display:flex;gap:8px;margin-top:22px}.edit button,.ai-confirm button{padding:11px 18px;border:1px solid var(--ink);border-radius:2px;background:var(--ink);color:var(--bg);font:700 15px Pretendard,sans-serif;cursor:pointer}',
  '.edit button.g,.ai-confirm button.g{background:transparent;color:var(--ink)}.edit .msg{min-height:1.4em;margin:14px 0 0;font-size:14px;color:var(--mut)}.edit .msg.err{color:var(--red)}',
  '.ai-confirm{width:min(520px,94vw);padding:26px;border:1px solid var(--ink);background:var(--bg);color:var(--ink)}.ai-confirm h2{margin:0;font-size:24px}.ai-confirm #ai-error{color:var(--red)}'
].join('');

function boardClient_() {
  var $ = function (s) { return document.querySelector(s); };
  var live = !!(document.body && document.documentElement && typeof requestAnimationFrame === 'function');
  var S = { key: null, round: null, rev: -1, gen: 0, redraw: null, cols: {}, names: [], seen: 0, data: null, queue: [], busy: false, shownTotal: 0, };
  var MAXSLIPS = 60, sortBtn = null;
  if (live) document.body.classList.add('wall');
  function mk(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  // 예약한 일은 화면을 다시 그리면(세대 S.gen이 바뀌면) 버립니다. 이전 질문의 도착 연출이 새 화면을 건드리지 않게 합니다.
  function later(f, ms) { var g = S.gen; return setTimeout(function () { if (g === S.gen) f(); }, ms); }
  function frame(f) { var g = S.gen; if (live) requestAnimationFrame(function () { requestAnimationFrame(function () { if (g === S.gen) f(); }); }); else f(); }
  function pad(n) { return n < 1000 ? ('00' + n).slice(-3) : String(n); }
  function cls(e, c, on) { if (e.classList) e.classList.toggle(c, on); else e.className = on ? c : ''; }
  function qr(url) {
    if (typeof qrcode !== 'function') return '';
    var q = qrcode(0, 'M'); q.addData(url); q.make();
    var n = q.getModuleCount(), d = '';
    for (var y = 0; y < n; y++) for (var x = 0; x < n; x++) if (q.isDark(y, x)) d += 'M' + x + ',' + y + 'h1v1h-1z';
    return '<svg viewBox="-2 -2 ' + (n + 4) + ' ' + (n + 4) + '" shape-rendering="crispEdges"><path d="' + d + '"/></svg>';
  }
  // 숫자: 자리마다 0~9 띠가 위아래로 굴러 새 값에 멈춥니다.
  function num(e, n) {
    if (!live || !e.insertBefore) { e.textContent = String(n); return; }
    var s = String(n), made = false;
    if (!e.children.length) e.textContent = '';
    while (e.children.length < s.length) {
      var d = mk('span', 'dg'), st = mk('span', 'st');
      for (var k = 0; k < 10; k++) st.appendChild(mk('span', '', String(k)));
      d.appendChild(st); e.insertBefore(d, e.firstChild); made = true;
    }
    while (e.children.length > s.length) e.removeChild(e.firstChild);
    var set = function () { for (var i = 0; i < s.length; i++) e.children[i].firstChild.style.transform = 'translateY(' + (-Number(s.charAt(i)) * 10) + '%)'; };
    if (made) frame(set); else set();
  }
  // 글을 단어 단위로 아래에서 위로 드러냅니다.
  function words(e, text, delay) {
    if (!live) { e.textContent = text; return; }
    e.textContent = ''; cls(e, 'on', false); cls(e, 'out', false);
    String(text).split(/\s+/).filter(String).forEach(function (w, k) {
      if (k) e.appendChild(document.createTextNode(' '));
      var o = mk('span', 'w'), i = mk('span', '', w);
      i.style.transitionDelay = (delay + k * 55) + 'ms'; o.appendChild(i); e.appendChild(o);
    });
    frame(function () { cls(e, 'on', true); });
  }
  function build(d) {
    var g = $('#g'), labs = $('#labs');
    S.names = d.cats.length ? d.cats.slice() : ['모두']; S.cols = {};
    g.textContent = ''; labs.textContent = '';
    if (live) { g.style.setProperty('--n', S.names.length); labs.style.setProperty('--n', S.names.length); }
    S.names.forEach(function (name) {
      var col = mk('section', 'col'), call = mk('p', 'call'), n = mk('div', 'cnt'), stack = mk('div', 'stack'), older = mk('div', 'older');
      stack.appendChild(older); col.appendChild(call); col.appendChild(n); col.appendChild(stack); g.appendChild(col);
      var lab = mk('div', 'lab');
      lab.appendChild(mk('b', '', name)); labs.appendChild(lab);
      S.cols[name] = { col: col, call: call, stack: stack, older: older, lab: lab, n: n, slips: {}, ids: [], shown: 0 };
      num(n, 0);
    });
  }
  function nameOf(x) { return S.names.length === 1 && S.names[0] === '모두' ? '모두' : (S.cols[x.c] ? x.c : '기타'); }
  function colOf(x) { return S.cols[nameOf(x)]; }
  function countOf(d, name) { return name === '모두' ? d.total : (d.counts[name] || 0); }
  // 칸 높이 u는 가장 많은 묶음이 화면 높이를 채우도록 정합니다. 칸이 얇아지면 글 대신 선으로 쌓습니다.
  function size() {
    if (!live) return;
    var g = $('#g'), h = g.clientHeight, w = g.clientWidth, n = S.names.length, max = 1;
    S.names.forEach(function (k) { max = Math.max(max, S.cols[k].shown); });
    var colW = (w - (n - 1) * Math.min(28, Math.max(10, w * 0.015))) / n;
    var fs = Math.max(13, Math.min(26, colW * 0.07));
    var thin = (h / max) < fs * 2.3 || colW < 110;
    var u = Math.max(2, Math.min(h / 3.4, (h - fs * 3.4 - (thin ? fs * 4.4 : 0)) / max));
    var lines = u > fs * 3.9 + 16 ? 3 : (u > fs * 2.6 + 12 ? 2 : 1);
    g.style.setProperty('--u', u.toFixed(1) + 'px'); g.style.setProperty('--fs', fs.toFixed(1) + 'px'); g.style.setProperty('--lines', lines);
    $('#labs').style.setProperty('--fl', Math.max(14, Math.min(26, colW * 0.075)).toFixed(1) + 'px');
    cls(g, 'thin', thin); S.u = u;
  }
  function slipFor(c, x, hidden) {
    var e = mk('div', 'slip' + (hidden ? ' g0 hid' : '')); e.appendChild(mk('span', '', x.t)); e.appendChild(mk('i', 'no', pad(x.no || 0)));
    c.slips[x.i] = e; c.stack.insertBefore(e, c.stack.firstChild);
    c.ids.unshift(x.i);
    while (c.ids.length > MAXSLIPS) { var old = c.ids.pop(); if (c.slips[old]) { c.stack.removeChild(c.slips[old]); delete c.slips[old]; } }
    return e;
  }
  function olderFor(c, total) {
    var rest = Math.max(0, total - c.ids.length);
    c.older.textContent = rest > 0 ? '이전 ' + rest + '개' : '';
    if (live) c.older.style.height = rest > 0 ? (rest * (S.u || 4)).toFixed(1) + 'px' : '0px';
  }
  function lead(d) {
    var best = '', top = -1, tie = false;
    S.names.forEach(function (k) { var v = S.cols[k].shown; if (v > top) { top = v; best = k; tie = false; } else if (v === top) tie = true; });
    S.names.forEach(function (k) { var on = S.names.length > 1 && !tie && top > 0 && k === best; cls(S.cols[k].lab, 'lead', on); cls(S.cols[k].col, 'lead', on); });
  }
  function meta(d) {
    num($('#c'), S.shownTotal);
    var st = $('#st'); st.textContent = d.open ? '받는 중' : '마감'; st.className = d.open ? '' : 'shut';
    $('#wait').hidden = S.shownTotal > 0;
  }
  // 처음 그리기, 또는 질문이 바뀐 뒤 다시 그리기: 막대가 바닥에서 자라고 숫자가 굴러 올라갑니다.
  function full(d) {
    var field = $('#field');
    cls(field, 'intro', true);
    build(d);
    S.seen = 0;
    d.items.forEach(function (x) { var c = colOf(x); if (c && !c.slips[x.i]) slipFor(c, x, false); S.seen = Math.max(S.seen, x.i); });
    S.names.forEach(function (k) { S.cols[k].shown = countOf(d, k); });
    S.shownTotal = d.total;
    size();
    S.names.forEach(function (k, i) { var c = S.cols[k]; olderFor(c, c.shown); c.call.textContent = c.ids.length ? c.slips[c.ids[0]].firstChild.textContent : ''; if (live) c.stack.style.transitionDelay = (300 + i * 110) + 'ms'; });
    words($('#q'), d.q, 0);
    frame(function () {
      cls(field, 'intro', false);
      S.names.forEach(function (k) { var c = S.cols[k]; num(c.n, c.shown); });
      lead(d); meta(d);
    });
    if (live) later(function () { S.names.forEach(function (k) { S.cols[k].stack.style.transitionDelay = ''; }); }, 2200);
  }
  function land(x) {
    var c = colOf(x); if (!c) return;
    c.shown++; S.shownTotal++;
    num(c.n, c.shown); size(); olderFor(c, c.shown); lead(S.data); meta(S.data);
    c.call.textContent = x.t;
  }
  // 새 답 하나: 위쪽 띠에 크게 → 묶음 이름 → 그 막대 맨 위 칸으로 줄어들며 들어감 → 칸이 잠깐 먹색으로 찍혔다가 돌아옴.
  function arrive(x, done) {
    var c = colOf(x), band = $('#band'), lab = band.querySelector('small'), p = band.querySelector('p');
    var e = slipFor(c, x, true);
    band.className = 'band'; band.hidden = false; cls($('#top'), 'arriving', true); p.style.transition = 'none'; p.style.transform = ''; p.style.opacity = '';
    var head = '방금 들어온 생각 · No.' + pad(x.no || 0) + (x.h ? ' · ' + x.h : '');
    lab.textContent = head;
    words(p, x.t, 120);
    later(function () { lab.textContent = ''; lab.appendChild(document.createTextNode(head)); lab.appendChild(mk('b', '', '→ ' + (x.c || '모두'))); }, 1150);
    later(function () { cls(e, 'g0', false); land(x); }, 1750);
    later(function () {
      var a = p.getBoundingClientRect(), b = e.getBoundingClientRect(), fs = parseFloat(getComputedStyle(p).fontSize) || 60;
      var k = (parseFloat(getComputedStyle(e).fontSize) || 20) / fs;
      p.style.transition = 'transform .75s cubic-bezier(.65,0,.25,1),opacity .45s ease .3s';
      p.style.transform = 'translate(' + Math.round(b.left + 12 - a.left) + 'px,' + Math.round(b.top + b.height / 2 - a.top - fs * k * 0.65) + 'px) scale(' + k.toFixed(3) + ')';
      p.style.opacity = '0'; cls(band, 'fly', true);
    }, 2350);
    later(function () {
      band.hidden = true; cls($('#top'), 'arriving', false); cls(e, 'hid', false); cls(e, 'fresh', true);
      later(function () { cls(e, 'fresh', false); }, 2600);
      done();
    }, 3100);
  }
  function pump() {
    if (S.busy || !S.queue.length) return;
    S.busy = true;
    // 한꺼번에 여러 개가 오면 앞의 것들은 막대에 바로 쌓고, 마지막 하나만 크게 보여 줍니다.
    while (S.queue.length > 1) { var x = S.queue.shift(), e = slipFor(colOf(x), x, true); (function (e, x, k) { later(function () { cls(e, 'g0', false); cls(e, 'hid', false); land(x); }, Math.min(k * 90, 2400)); })(e, x, S.queue.length); }
    var last = S.queue.shift();
    arrive(last, function () { S.busy = false; pump(); });
  }
  // 판정 연결 오류: '오류:' 흔적이 있는 답이 새로 생기거나 이유가 바뀔 때 한 번만 알립니다(실패 알림은 8초 뒤 접힙니다).
  function alertErr(d) {
    var why = d.err ? d.err.why || '이유를 알 수 없음' : '';
    if (why === S.errWhy) return;
    S.errWhy = why;
    if (why && typeof fail === 'function') fail('판정 연결 오류 ' + d.err.n + '개: ' + why + ' · [분류]로 다시 나눌 수 있어요');
  }
  function keyOf(d) { return d.q + '\u0001' + d.cats.join('\u0001') + '\u0001' + (d.mix || 0); }
  function update(d) {
    // 늦게 도착한 예전 응답(질문을 바꾸기 전에 보낸 요청)은 버립니다.
    if (d.rev < S.rev) return;
    S.rev = d.rev;
    alertErr(d);
    // 다시 그리는 중이면 가장 새 자료만 기억했다가 그것으로 그립니다.
    if (S.redraw) { S.redraw.d = d; return; }
    S.data = d;
    var key = keyOf(d);
    if (S.key === null) { S.key = key; S.round = d.round; full(d); return; }
    // 질문이나 묶음이 바뀌었거나, 보관으로 [응답] 시트가 새로 생겼으면(round) 처음부터 다시 그립니다.
    if (key !== S.key || d.round !== S.round) {
      S.gen++; S.key = key; S.round = d.round; S.queue = []; S.busy = false; $('#band').hidden = true; cls($('#top'), 'arriving', false);
      cls($('#q'), 'out', true); cls($('#field'), 'intro', true);
      var r = S.redraw = { d: d };
      later(function () { S.redraw = null; S.data = r.d; S.key = keyOf(r.d); S.round = r.d.round; full(r.d); }, live ? 560 : 0);
      return;
    }
    // 화면에서 사라진 답(교수가 시트에서 '숨김'으로 바꾼 것)은 칸에서 뺍니다.
    // 자료에 모든 답이 왔으면 없는 칸은 모두 빼고, 최근 것만 왔으면 그 범위 안에서만 뺍니다(시트를 정렬해도 번호로 찾습니다).
    var keep = {}, low = Infinity; d.items.forEach(function (x) { keep[x.i] = 1; low = Math.min(low, x.i); });
    var all = d.items.length >= d.total;
    S.names.forEach(function (k) {
      var c = S.cols[k];
      c.ids = c.ids.filter(function (i) {
        if ((all || i >= low) && !keep[i] && i <= S.seen) { if (c.slips[i]) { c.stack.removeChild(c.slips[i]); delete c.slips[i]; } return false; }
        return true;
      });
    });
    var fresh = d.items.filter(function (x) { return x.i > S.seen; });
    fresh.forEach(function (x) { S.seen = Math.max(S.seen, x.i); S.queue.push(x); });
    // 숨김 처리 등으로 줄어든 개수는 바로 맞춥니다(도착 중인 것은 빼고).
    var pend = {}; S.queue.forEach(function (x) { var k = nameOf(x); pend[k] = (pend[k] || 0) + 1; });
    S.names.forEach(function (k) { var c = S.cols[k], want = countOf(d, k) - (pend[k] || 0); if (!S.busy && c.shown !== want) { c.shown = want; num(c.n, want); olderFor(c, want); } });
    if (!S.busy) { S.shownTotal = d.total - S.queue.length; size(); lead(d); meta(d); }
    if (!live) { S.queue.forEach(function (x) { var c = colOf(x); slipFor(c, x, false); c.shown++; S.shownTotal++; num(c.n, c.shown); }); S.queue = []; lead(d); meta(d); return; }
    pump();
  }
  function tick() {
    // 화면 그리기에서 오류가 나도 다음 묻기는 꼭 예약합니다(프로젝터 화면이 조용히 멈추지 않게).
    google.script.run.withSuccessHandler(function (d) { try { update(d); } finally { setTimeout(tick, ((d && d.refresh) || 4) * 1000); } })
      .withFailureHandler(function () { setTimeout(tick, 8000); }).getBoard();
  }
  // 질문 바꾸기(E): 오른쪽에서 칸이 열립니다. 교수 PIN은 이 창을 닫을 때까지 기억합니다.
  function editor() {
    var ed = $('#ed');
    function open() {
      var d = S.data || { q: '', cats: [], open: true };
      $('#eq').value = d.q; $('#ec').value = d.cats.filter(function (k) { return k !== '기타'; }).join('\n');
      $('#eo').checked = !!d.open; $('#ef').checked = false; touched = false; $('#em').textContent = ''; $('#em').className = 'msg';
      $('#ep').value = ''; $('#epl').hidden = !!tok();
      cls(ed, 'on', true); later(function () { $('#eq').focus(); }, 80);
    }
    function close() { cls(ed, 'on', false); pending = ''; }
    // PIN이 한 번 맞으면 서버가 이 화면용 표를 줍니다. 창을 닫을 때까지 PIN을 다시 묻지 않습니다.
    var touched = false, sorting = false, pending = '';
    function tok() { try { return sessionStorage.getItem('board-tok') || ''; } catch (e) { return ''; } }
    function keep(t) { try { if (t) sessionStorage.setItem('board-tok', t); else sessionStorage.removeItem('board-tok'); } catch (e) {} }
    $('#eq').oninput = function () { if (!touched && S.data) $('#ef').checked = this.value.trim() !== S.data.q; };
    $('#ef').onchange = function () { touched = true; };
    $('#ex').onclick = close;
    $('#es').onclick = function () {
      var b = $('#es'), m = $('#em'), pin = $('#ep').value.trim(), cred = pin || tok();
      if (!cred) { m.textContent = '교수 PIN은 시트 메뉴 반응판 → 주소와 교수 PIN 보기에 있습니다.'; m.className = 'msg err'; fail('질문 바꾸기를 못했어요. 교수 PIN 4자리를 넣어 주세요.'); return; }
      if (pin && !/^\d{4}$/.test(pin)) { m.textContent = ''; fail('질문 바꾸기를 못했어요. PIN은 숫자 4자리입니다.'); return; }
      b.disabled = true; m.textContent = '바꾸는 중…'; m.className = 'msg';
      google.script.run.withSuccessHandler(function (r) {
        b.disabled = false;
        if (r.expired) { keep(''); $('#epl').hidden = false; }
        if (r.token) keep(r.token);
        if (!r.ok) { m.textContent = r.msg; m.className = 'msg err'; fail('질문 바꾸기를 못했어요. ' + r.msg); return; }
        close(); update(r.board);
        toast(r.archived ? '질문을 바꿨어요. 지난 응답은 시트에 보관했습니다.' : '질문을 바꿨어요.');
      }).withFailureHandler(function (e) { b.disabled = false; m.textContent = e.message; m.className = 'msg err'; fail('질문 바꾸기를 못했어요. ' + e.message); })
        .setBoard(cred, $('#eq').value, $('#ec').value, $('#eo').checked, $('#ef').checked);
    };
    // AI로 묶음 만들기: 질문을 보고 후보를 묶음 칸에 채웁니다. 저장은 "바꾸기"를 눌러야 됩니다.
    $('#eg').onclick = function () {
      var b = $('#eg'), m = $('#em'), pin = $('#ep').value.trim(), cred = pin || tok();
      if (!cred) { m.textContent = '교수 PIN은 시트 메뉴 반응판 → 주소와 교수 PIN 보기에 있습니다.'; m.className = 'msg err'; fail('묶음 만들기를 못했어요. 교수 PIN 4자리를 넣어 주세요.'); return; }
      if (pin && !/^\d{4}$/.test(pin)) { m.textContent = ''; fail('묶음 만들기를 못했어요. PIN은 숫자 4자리입니다.'); return; }
      b.disabled = true; $('#es').disabled = true; m.textContent = 'AI가 묶음을 만드는 중…'; m.className = 'msg';
      google.script.run.withSuccessHandler(function (r) {
        b.disabled = false; $('#es').disabled = false;
        if (r.expired) { keep(''); $('#epl').hidden = false; }
        if (r.token) keep(r.token);
        if (!r.ok) { m.textContent = r.msg; m.className = 'msg err'; fail('묶음 만들기를 못했어요. ' + r.msg); return; }
        if (ed.classList && !ed.classList.contains('on')) return; // 그 사이 창을 닫았으면 채우지 않습니다
        $('#ec').value = r.cats; m.textContent = '묶음을 채웠어요. 고친 뒤 "바꾸기"를 누르면 학생 폰에 보입니다.'; m.className = 'msg';
        toast('AI가 묶음을 만들었어요. 확인하고 바꾸기를 눌러 주세요.');
      }).withFailureHandler(function (e) { b.disabled = false; $('#es').disabled = false; m.textContent = e.message; m.className = 'msg err'; fail('묶음 만들기를 못했어요. ' + e.message); })
        .suggestCats(cred, $('#eq').value);
    };
    // [분류](C): 한 번 누르면 알아서 묶입니다. 기타·미분류·오류인 답은 낱말 맞추기 → Jev로 다시 나누고, 묶음이 없거나 기타가 많으면
    // 서버가 비슷한 답끼리 스스로 묶어 바로 저장합니다. 생성형 AI가 필요하면 동의한 뒤에만 이어서 합니다.
    // 교수 PIN이 필요해 편집 칸이 열린 경우(pending)에는 PIN이 다 들어오면(4자리, 또는 Enter) 바로 이어서 합니다.
    var aiConfirm = $('#ai-confirm'), aiToken = '', aiKey = '';
    function showAiConfirm(token, key, msg) {
      aiToken = token; aiKey = key;
      var err = $('#ai-error'); err.textContent = msg || ''; err.hidden = !msg;
      aiConfirm.showModal();
    }
    $('#ai-no').onclick = function () { aiConfirm.close(); };
    $('#ai-yes').onclick = function () { aiConfirm.close(); if (aiKey) runClassify(aiToken, aiKey); aiKey = ''; };
    function classify() {
      if (aiConfirm.open) return;
      if (sorting) { toast('분류하는 중이에요. 잠시만 기다려 주세요.'); return; }
      var pin = ed.classList.contains('on') ? $('#ep').value.trim() : '', cred = pin || tok();
      if (!cred) {
        open(); pending = 'classify'; $('#em').textContent = '분류는 교수만 할 수 있어요. 교수 PIN 4자리를 넣으면 바로 분류합니다.'; $('#em').className = 'msg err';
        later(function () { $('#ep').focus(); }, 120);
        fail('분류를 못했어요. 교수 PIN 4자리를 넣어 주세요.'); return;
      }
      if (pin && !/^\d{4}$/.test(pin)) { fail('분류를 못했어요. PIN은 숫자 4자리입니다.'); return; }
      pending = '';
      runClassify(cred, '');
    }
    function runClassify(cred, confirmKey) {
      sorting = true; if (sortBtn) sortBtn.label('분류 중…');
      toast('분류하는 중이에요. 기타에 있는 답을 다시 나누고, 묶음이 없으면 알아서 묶습니다.');
      var end = function () { sorting = false; if (sortBtn) sortBtn.label('분류'); };
      google.script.run.withSuccessHandler(function (r) {
        end();
        if (r.expired) { keep(''); $('#epl').hidden = false; }
        if (r.token) keep(r.token);
        if (r.consentBusy) { showAiConfirm(r.token, confirmKey, r.msg); return; }
        if (!r.ok) { if (r.expired || !r.token) { open(); pending = 'classify'; $('#em').textContent = r.msg; $('#em').className = 'msg err'; } fail('분류를 못했어요. ' + r.msg); return; }
        if (ed.classList.contains('on')) close();
        if (r.board) update(r.board);
        if (r.needsReload) { fail((r.partial ? '일부만 저장되었을 수 있습니다. 화면 동기화에 실패했습니다.' : '시트에는 저장했지만 화면 동기화에 실패했습니다.') + ' 화면을 새로고침하십시오.'); return; }
        if (r.partial) { fail('일부만 저장되었을 수 있습니다. 화면을 다시 불러옵니다.'); return; }
        var tot = (r.moved || 0) + (r.auto ? r.auto.moved : 0), msg = '';
        if (r.grouped && r.moved) msg = r.moved + '개를 묶음으로 옮겼어요. ';
        if (r.auto) msg += r.auto.kind === 'ai' ? 'AI로 ' + r.auto.groups + '묶음으로 묶었어요. ' : 'AI 없이 ' + r.auto.groups + '묶음으로 묶었어요. ';
        else if (r.grouped && !msg) msg = '묶음으로 옮긴 답은 없어요. ';
        if (r.auto || r.grouped) msg += '기타에 ' + r.left + '개가 남았어요.';
        if (r.autoNone) msg += (msg ? ' ' : '') + (r.autoFull ? '묶음이 거의 꽉 차서 새 묶음은 만들지 않았어요.' : '새로 묶을 답이 아직 적어요(4개 이상 필요). 답이 더 모이면 다시 눌러 주세요.');
        if (r.more) msg += ' ' + r.more + '개는 이번에 다 못 나눴어요. [분류]를 한 번 더 누르세요.';
        var off = r.aiOff ? ' 응답을 한꺼번에 보내 새 묶음을 만드는 기능이 꺼져 있습니다.' : '';
        // 멈춘 이유(r.stopped)와 묶기 실패(r.autoErr)는 몇 개를 처리했든 알립니다.
        if (r.autoErr) fail('묶음을 만들지 못했어요. ' + r.autoErr + (tot ? ' (' + tot + '개는 옮겼어요.)' : '') + off);
        else if (r.stopped) fail('분류를 다 못 했어요. ' + r.stopped + (r.more ? ' ' + r.more + '개를 그대로 두었어요.' : '') + (tot || r.words ? ' (' + (tot || r.words) + '개는 옮겼어요.)' : '') + off);
        else toast((msg + off).trim());
        if (r.confirmKey) showAiConfirm(r.token, r.confirmKey);
      }).withFailureHandler(function (e) { end(); fail('분류를 못했어요. ' + e.message); })
        .classifyBoard(cred, confirmKey);
    }
    // PIN이 모자라 열린 편집 칸: 4자리가 다 들어오거나 Enter를 치면 바로 분류합니다.
    $('#ep').oninput = function () { if (pending === 'classify' && /^\d{4}$/.test(this.value.trim())) classify(); };
    $('#ep').onkeydown = function (e) { if (e.key === 'Enter' && pending === 'classify') { e.preventDefault(); classify(); } };
    $('#ek').onclick = classify;
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') close(); });
    return { toggle: function () { if (ed.classList.contains('on')) close(); else open(); }, classify: classify };
  }
  // 참여 QR: 페이지를 만들 때 서버가 넣어 준 주소(BOOT.join)로 한 번 그립니다. 믿을 주소가 없으면 깨진 QR 대신 안내를 띄웁니다.
  var boot = (typeof BOOT === 'object' && BOOT) || {};
  if (boot.join) {
    var svg = qr(boot.join);
    if (svg) { $('#qr').innerHTML = svg; $('#join').hidden = false; }
    else if (typeof fail === 'function') fail('참여 QR을 그리지 못했어요. QR 도구를 불러오지 못했습니다. 인터넷을 확인하고 새로고침하세요.');
  } else if (typeof fail === 'function') fail('참여 QR을 만들 주소가 없어요. 웹 도구에서 받은 "강의장 큰 화면 주소"로 이 화면을 다시 열거나, 시트 메뉴 반응판 → 웹 앱 주소 넣기에 /exec 주소를 넣으세요.');
  if (boot.warn && typeof fail === 'function') fail('이 화면을 연 주소와 시트에 저장된 웹 앱 주소가 다릅니다. QR은 시트에 저장된 주소로 만들었습니다. 바꾸려면 시트 메뉴 반응판 → 웹 앱 주소 넣기를 하세요.');
  if (live) {
    var ed = editor();
    try { if (localStorage.getItem('board-theme') === 'dark') document.body.classList.add('dark'); } catch (e) {}
    if (window.stage) {
      window.stage.add('질문 바꾸기', 'E', ed.toggle);
      sortBtn = window.stage.add('분류', 'C', ed.classify);
      var th = window.stage.add(document.body.classList.contains('dark') ? '밝게' : '어둡게', 'T', function () {
        var dk = document.body.classList.toggle('dark'); th.label(dk ? '밝게' : '어둡게'); toast(dk ? '어두운 화면으로 바꿨어요.' : '밝은 화면으로 바꿨어요.');
        try { localStorage.setItem('board-theme', dk ? 'dark' : 'light'); } catch (e) {}
      });
    }
    addEventListener('resize', function () { size(); S.names.forEach(function (k) { olderFor(S.cols[k], S.cols[k].shown); }); });
  }
  tick();
}

var BOARD_HTML = '<div class="stage"><header class="top" id="top"><div class="head"><div class="meta"><span>교실 반응판</span><span>응답 <b id="c">0</b></span><span id="st">받는 중</span></div><h1 id="q">…</h1>' +
  '<div class="band" id="band" hidden><small></small><p></p></div></div>' +
  '<div class="join" id="join" hidden><p><b>폰으로 참여</b>QR을 찍고 한 줄</p><div class="qr" id="qr"></div></div></header>' +
  '<main class="field intro" id="field"><div class="cols" id="g"></div><div class="base"></div><div class="labs" id="labs"></div>' +
  '<p class="wait" id="wait" hidden>첫 번째 생각을 기다리고 있습니다. QR을 찍고 한 줄을 보내 주세요.</p></main></div>' +
  '<aside class="edit" id="ed"><h2>질문 바꾸기</h2>' +
  '<label>질문<textarea id="eq" rows="3" maxlength="160"></textarea></label>' +
  '<label>묶음<small>한 줄에 하나, 10개까지. 비우면 묶지 않고 한 막대에 모읍니다. 맞지 않는 답은 \'기타\'로 갑니다.</small><textarea id="ec" rows="6"></textarea></label>' +
  '<div style="margin:-6px 0 16px"><button type="button" class="g" id="eg">AI로 묶음 만들기</button></div>' +
  '<label class="ck"><input type="checkbox" id="eo">답 받는 중</label>' +
  '<label class="ck"><input type="checkbox" id="ef">지금까지 응답은 시트에 보관하고 화면을 비웁니다</label>' +
  '<label id="epl">교수 PIN<small>시트 메뉴 반응판 → 주소와 교수 PIN 보기. 한 번 맞으면 이 창을 닫을 때까지 다시 묻지 않습니다.</small><input type="password" id="ep" inputmode="numeric" maxlength="4" autocomplete="off"></label>' +
  '<div class="row"><button id="es">바꾸기</button><button id="ek" class="g">분류</button><button id="ex" class="g">닫기</button></div><p class="msg" id="em"></p></aside>' +
  '<dialog class="ai-confirm" id="ai-confirm" aria-labelledby="ai-title" aria-describedby="ai-desc"><h2 id="ai-title">응답을 한꺼번에 보내 새 묶음을 만드는 기능을 켜시겠습니까?</h2>' +
  '<p id="ai-desc">' + BOARD_BATCH_NOTICE + ' ' + BOARD_SINGLE_NOTICE + ' 이름·연락처를 적지 않도록 안내해 주십시오.</p>' +
  '<p>동의는 저장되며 시트 메뉴 반응판 → AI 묶기 끄기에서 취소할 수 있습니다.</p>' +
  '<p id="ai-error" role="alert" hidden></p>' +
  '<div class="row"><button type="button" class="g" id="ai-no" autofocus>지금은 사용하지 않기</button><button type="button" id="ai-yes">동의하고 기능 켜기</button></div></dialog>' +
  '<script src="https://kimcook-edu.pages.dev/vendor/qrcode/qrcode.js"></script>' + STAGE_JS + '<script>var BOOT = /*BOOT*/null;</script>' +
  '<script>(' + boardClient_.toString() + ')();</script>';
