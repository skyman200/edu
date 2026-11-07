/* ===== 레시피 1. 실시간 교실 반응판 =====
   사용법: 웹 도구(https://kimcook-edu.pages.dev/#gas)에서 이 레시피를 고르고 안내를 따르세요. 코드를 붙여 넣고 start 를 실행하면 준비가 끝납니다.
   질문과 묶음은 강의장 큰 화면에서 바로 바꿀 수 있습니다(교수 PIN 필요). [설정] 시트에서 바꿔도 됩니다.
   답은 [응답] 시트에 쌓이고, Jev가 묶음을 고릅니다. */
var BOARD_HEAD = ['시간', '응답', '묶음', '확신도', '상태', '번호'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('반응판')
    .addItem('처음 준비', 'setup').addItem('새 질문 시작(지금 응답은 보관)', 'newRound').addItem('수업 코드 입력', 'setClassCode').addItem('주소와 교수 PIN 보기', 'showUrls').addToUi();
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
  sheet_('응답', BOARD_HEAD);
  say_('준비 끝! 이제 웹 도구 화면의 다음 단계(배포)를 따라 하세요.\n큰 화면에서 질문을 바꿀 때 쓰는 교수 PIN: ' + pin_('board') + ' — 학생에게 알리지 마세요. 시트 메뉴 반응판 → 주소와 교수 PIN 보기에서 다시 볼 수 있습니다.');
}

// 화면이 다시 그려야 하는 변화(질문 바꾸기, 보관)마다 1씩 올립니다. 큰 화면은 이보다 낮은 번호의 늦게 온 응답을 버립니다.
function bump_() {
  var p = PropertiesService.getScriptProperties(), r = Number(p.getProperty('BOARD_REV') || 0) + 1;
  p.setProperty('BOARD_REV', String(r));
  CacheService.getScriptCache().remove('board'); CacheService.getScriptCache().remove('q');
  return r;
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
  var url = ScriptApp.getService().getUrl();
  say_(url ? '학생용: ' + url + '\n강의장 화면: ' + url + '?v=board\n질문 바꾸기 PIN: ' + pin_('board') : '아직 웹 앱으로 배포하지 않았습니다. 배포 → 새 배포 → 웹 앱을 먼저 하세요.');
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
  if (e && e.parameter && e.parameter.v === 'board') return page_('반응판', BOARD_HTML, { raw: true, css: BOARD_CSS + STAGE_CSS });
  return page_('내 생각 보내기', ASK_HTML, { raw: true, css: ASK_CSS });
}

function rev_() { return Number(PropertiesService.getScriptProperties().getProperty('BOARD_REV') || 0); }
function round_() { var sh = ss_().getSheetByName('응답'); return sh ? sh.getSheetId() : 0; }

// 학생 폰이 12초마다 묻습니다. 사람이 많아도 시트를 매번 읽지 않게 5초 동안 기억합니다.
// round: 지금 [응답] 시트. 보관 뒤에는 바뀌므로, 폰은 이전 질문에 쓴 답을 새 판에 보내지 않습니다.
// 읽는 사이 질문이 바뀌었으면(rev) 기억하지 않습니다.
function getQuestion() {
  var cache = CacheService.getScriptCache(), hit = cache.get('q');
  if (hit) return JSON.parse(hit);
  var rev = rev_(), c = cfg_(), out = { q: c.q, open: c.open, round: round_() };
  if (rev_() === rev) cache.put('q', JSON.stringify(out), 5);
  return out;
}

function submitAnswer(text, sid, round) {
  var c = cfg_();
  if (!c.open) return { ok: false, msg: '지금은 답을 받지 않습니다.' };
  text = String(text || '').replace(/\s+/g, ' ').trim().slice(0, 140);
  if (text.length < 2) return { ok: false, msg: '두 글자 이상 써 주세요.' };
  sid = String(sid || 'x').slice(0, 40);
  if (!throttle_('s:' + sid, 6, 600) || !throttle_('all', 90, 60)) return { ok: false, msg: '잠시 뒤에 다시 보내 주세요.' };
  var cat = '', conf = '', status = '표시';
  if (c.cats.length >= 2) {
    try {
      var crit = {};
      c.cats.forEach(function (name, i) { crit['c' + i] = name; });
      crit.other = '위 항목 어디에도 맞지 않는 내용';
      var a = jev_({ 질문: c.q, 응답: text }, {
        cat: { type: 'choice', instructions: '교수의 질문에 대한 응답이 어느 묶음에 속하는지 고르세요.', criteria: crit },
        bad: { type: 'noul', instructions: '응답이 질문과 무관한 광고, 욕설, 장난, 개인정보 노출인가?' }
      });
      var k = a.cat.choice, n = k && k.charAt(0) === 'c' ? Number(k.slice(1)) : -1;
      cat = c.cats[n] || '기타';
      conf = Math.round((a.cat.confidence || 0) * 100) / 100;
      if (conf < 0.45) cat = '기타';
      if (a.bad.p > 0.8) status = '숨김';
    } catch (err) { cat = ''; }
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
    row_(sh, [new Date(), text, cat, conf, status, id]);
    no = sh.getLastRow() - 1;
  } finally { lock.releaseLock(); }
  return { ok: true, cat: cat, no: no };
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
  rows.forEach(function (r, k) {
    if (r[4] === '숨김' || String(r[1]) === '') return;
    var cat = '';
    if (cats.length) { cat = String(r[2] || ''); if (cats.indexOf(cat) < 0) cat = '기타'; counts[cat]++; }
    items.push({ i: Number(r[5]) || k + 2, no: k + 1, h: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'HH:mm') : '', t: String(r[1]), c: cat });
  });
  var out = { q: c.q, cats: cats, counts: counts, open: c.open, refresh: c.refresh, total: items.length, join: ScriptApp.getService().getUrl() || '',
    round: sh ? sh.getSheetId() : 0, rev: rev, items: items.slice(-240) };
  if (rev_() === rev) cache.put('board', JSON.stringify(out), 3);
  return out;
}

// 학생 폰 화면. 큰 화면과 같은 말씨: 종이 바탕 위의 흰 종이 한 장(슬립)에 생각을 적어 보냅니다.
// 보내면 그 종이가 위로 날아가고, 아래 '내가 보낸 생각'에 번호와 묶음이 찍힌 종이로 남습니다. 큰 화면의 칸에도 같은 번호가 붙습니다.
var ASK_CSS = [
  ':root{--bg:#efede8;--ink:#141414;--mut:#77736b;--hair:rgba(20,20,20,.16);--slip:#fff;--red:#e54536;--e:cubic-bezier(.2,.8,.2,1)}',
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
  '.cf{display:flex;align-items:center;gap:12px;padding:0 16px 14px}',
  '.meter{position:relative;flex:1;height:1px;background:var(--hair)}.meter i{position:absolute;left:0;top:-1px;width:100%;height:3px;background:var(--ink);transform-origin:0 50%;transform:scaleX(0);transition:transform .25s var(--e)}.meter.full i{background:var(--red)}',
  '.cnt{font:600 12px/1 Montserrat,sans-serif;letter-spacing:.04em;color:var(--mut);font-variant-numeric:tabular-nums}',
  '#b{display:flex;justify-content:space-between;align-items:center;width:100%;margin-top:12px;padding:18px 20px;border:1px solid var(--ink);border-radius:2px;background:var(--ink);color:#fff;font:700 17px/1 Pretendard,sans-serif;cursor:pointer;transition:transform .15s,opacity .3s}',
  '#b:active{transform:translateY(1px)}#b:disabled{opacity:.35;cursor:default}#b i{font-style:normal;font-family:Montserrat,sans-serif;transition:transform .3s var(--e)}#b:not(:disabled):hover i{transform:translateX(4px)}',
  '#m{min-height:1.5em;margin:14px 0 0}.msg{margin:0;font-size:15px;font-weight:600;line-height:1.45}.msg b{color:var(--red)}.msg.err{color:var(--red)}',
  '.msg.rise,.sent.rise{animation:rise .55s var(--e) both}@keyframes rise{from{opacity:0;transform:translateY(12px)}}',
  '.mine{margin-top:34px}.mine[hidden]{display:none}',
  '.mine h2{display:flex;justify-content:space-between;margin:0 0 10px;padding-top:10px;border-top:1px solid var(--ink);font:700 11px/1 Montserrat,Pretendard,sans-serif;letter-spacing:.2em}',
  '.mine h2 span:first-child{font-family:Pretendard,sans-serif;font-size:13px;letter-spacing:-.01em}.mine h2 span:last-child{letter-spacing:.04em}',
  '.sent{display:flex;margin:0 0 3px;background:var(--slip);font-size:16px;line-height:1.45}',
  '.sent .no{flex:none;display:flex;align-items:center;padding:12px 12px;border-right:1px solid var(--hair);font:700 11px/1 Montserrat,sans-serif;letter-spacing:.06em;color:var(--mut);font-variant-numeric:tabular-nums}',
  '.sent p{flex:1;margin:0;padding:11px 14px;overflow-wrap:anywhere}',
  '.sent small{display:block;margin-top:3px;font-size:12px;font-weight:700;color:var(--mut)}.sent:first-child small{color:var(--red)}',
  '@media (prefers-reduced-motion:reduce){*{transition-duration:.01ms!important;animation-duration:.01ms!important}}'
].join('');

function askClient_() {
  var $ = function (s) { return document.querySelector(s); };
  var live = typeof requestAnimationFrame === 'function', sid, cq = null, round = 0, sending = false, open = true, lt = 0, mine = { q: '', round: 0, list: [] };
  try { sid = localStorage.getItem('sid'); if (!sid) { sid = Math.random().toString(36).slice(2, 12); localStorage.setItem('sid', sid); } } catch (e) { sid = Math.random().toString(36).slice(2, 12); }
  try { mine = JSON.parse(localStorage.getItem('board-mine') || '') || mine; } catch (e) {}
  function mk(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function pad(n) { return n < 1000 ? ('00' + n).slice(-3) : String(n); }
  function cls(e, c, on) { if (e.classList) e.classList.toggle(c, on); else e.className = on ? c : ''; }
  function note(t, bad, tag) {
    var m = $('#m'), p = mk('p', 'msg rise' + (bad ? ' err' : ''));
    m.textContent = ''; p.appendChild(document.createTextNode(t));
    if (tag) { p.appendChild(document.createTextNode(' 「')); p.appendChild(mk('b', '', tag)); p.appendChild(document.createTextNode('」 쪽으로 모였어요. 큰 화면을 보세요.')); }
    m.appendChild(p);
  }
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
  function save() { try { localStorage.setItem('board-mine', JSON.stringify(mine)); } catch (e) {} }
  function setOpen(o) {
    open = o;
    var st = $('#st'); st.textContent = o ? '받는 중' : '마감'; st.className = o ? '' : 'shut';
    cls($('#card'), 'shut', !o);
    if (!sending) $('#b').disabled = !o;
  }
  function load() {
    google.script.run.withSuccessHandler(function (r) {
      if (cq !== null && r.q !== cq) { note('질문이 바뀌었습니다. 새 질문에 답해 주세요.'); }
      if (r.q !== cq) words($('#q'), r.q);
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
    if (t.trim().length < 2 || sending) return;
    sending = true; $('#b').disabled = true;
    google.script.run.withSuccessHandler(function (r) {
      sending = false; $('#b').disabled = !open;
      if (!r.ok) { note(r.msg, 1); if (r.stale) load(); return; }
      var now = new Date(), h = ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
      mine.list.push({ t: t.replace(/\s+/g, ' ').trim().slice(0, 140), c: r.cat || '', no: r.no || 0, h: h }); mine.list = mine.list.slice(-20); save();
      note('보냈습니다' + (r.no ? ' · No.' + pad(r.no) : '') + '.' + (r.cat ? '' : ' 큰 화면을 보세요.'), 0, r.cat || '');
      var card = $('#card');
      function reset() { $('#t').value = ''; count(); drawMine(true); }
      if (!live) { reset(); return; }
      cls(card, 'go', true);
      setTimeout(function () {
        reset(); cls(card, 'go', false); cls(card, 'come', true);
        requestAnimationFrame(function () { requestAnimationFrame(function () { cls(card, 'come', false); }); });
      }, 520);
    }).withFailureHandler(function (e) { sending = false; $('#b').disabled = !open; note('보내지 못했습니다: ' + e.message, 1); }).submitAnswer(t, sid, round);
  };
  count(); drawMine(false); load();
}

var ASK_HTML = '<main class="ask"><div class="bar"><span>LIVE BOARD</span><span id="st">받는 중</span></div>' +
  '<p class="qk">Q.</p><h1 id="q">…</h1>' +
  '<div class="card" id="card"><div class="ch"><span>MY THOUGHT</span><span>이름 없이 보냅니다</span></div>' +
  '<textarea id="t" maxlength="140" placeholder="한두 문장으로 적어 주세요"></textarea>' +
  '<div class="cf"><div class="meter" id="meter"><i id="mt"></i></div><span class="cnt" id="n">0 / 140</span></div></div>' +
  '<button id="b">보내기<i>→</i></button><div id="m"></div>' +
  '<section class="mine" id="mine" hidden><h2><span>내가 보낸 생각</span><span id="mc">0개</span></h2><div id="ml"></div></section></main>' +
  '<script>(' + askClient_.toString() + ')();</script>';

// 강의장 큰 화면. 답 하나가 한 칸(종이 띠)이고, 묶음마다 띠가 쌓여 막대가 됩니다. 막대 높이가 곧 답의 수입니다.
// 새 답은 먼저 화면 위쪽에 크게 한 단어씩 나타나고, 묶음이 정해지면 줄어들며 그 막대의 맨 위 칸으로 들어갑니다.
// 답이 많아 칸이 얇아지면 글 대신 선으로 쌓고, 막대 위에 가장 최근 답 한 줄을 남깁니다. 강조색(빨강)은 가장 많은 묶음의 숫자에만 씁니다.
var BOARD_CSS = [
  ':root{--bg:#efede8;--fg:#141414;--ink:#141414;--mut:#77736b;--hair:rgba(20,20,20,.16);--slip:#ffffff;--bar:#bdb9b0;--red:#e54536;--e:cubic-bezier(.2,.8,.2,1)}',
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
  '#c{display:inline-flex;vertical-align:bottom}',
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
  '.edit .row{display:flex;gap:8px;margin-top:22px}.edit button{padding:11px 18px;border:1px solid var(--ink);border-radius:2px;background:var(--ink);color:var(--bg);font:700 15px Pretendard,sans-serif;cursor:pointer}',
  '.edit button.g{background:transparent;color:var(--ink)}.edit .msg{min-height:1.4em;margin:14px 0 0;font-size:14px;color:var(--mut)}.edit .msg.err{color:var(--red)}'
].join('');

function boardClient_() {
  var $ = function (s) { return document.querySelector(s); };
  var live = !!(document.body && document.documentElement && typeof requestAnimationFrame === 'function');
  var S = { key: null, round: null, rev: -1, gen: 0, redraw: null, cols: {}, names: [], seen: 0, data: null, queue: [], busy: false, shownTotal: 0, qrDone: false };
  var MAXSLIPS = 60;
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
  function keyOf(d) { return d.q + '\u0001' + d.cats.join('\u0001'); }
  function update(d) {
    // 늦게 도착한 예전 응답(질문을 바꾸기 전에 보낸 요청)은 버립니다.
    if (d.rev < S.rev) return;
    S.rev = d.rev;
    // 다시 그리는 중이면 가장 새 자료만 기억했다가 그것으로 그립니다.
    if (S.redraw) { S.redraw.d = d; return; }
    S.data = d;
    var key = keyOf(d);
    if (d.join && !S.qrDone) { var svg = qr(d.join); if (svg) { $('#qr').innerHTML = svg; $('#join').hidden = false; } S.qrDone = true; }
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
    function close() { cls(ed, 'on', false); }
    // PIN이 한 번 맞으면 서버가 이 화면용 표를 줍니다. 창을 닫을 때까지 PIN을 다시 묻지 않습니다.
    var touched = false;
    function tok() { try { return sessionStorage.getItem('board-tok') || ''; } catch (e) { return ''; } }
    function keep(t) { try { if (t) sessionStorage.setItem('board-tok', t); else sessionStorage.removeItem('board-tok'); } catch (e) {} }
    $('#eq').oninput = function () { if (!touched && S.data) $('#ef').checked = this.value.trim() !== S.data.q; };
    $('#ef').onchange = function () { touched = true; };
    $('#ex').onclick = close;
    $('#es').onclick = function () {
      var b = $('#es'), m = $('#em'), pin = $('#ep').value.trim(), cred = pin || tok();
      if (!cred) { m.textContent = '교수 PIN 4자리를 넣어 주세요. 시트 메뉴 반응판 → 주소와 교수 PIN 보기에 있습니다.'; m.className = 'msg err'; return; }
      if (pin && !/^\d{4}$/.test(pin)) { m.textContent = 'PIN은 숫자 4자리입니다.'; m.className = 'msg err'; return; }
      b.disabled = true; m.textContent = '바꾸는 중…'; m.className = 'msg';
      google.script.run.withSuccessHandler(function (r) {
        b.disabled = false;
        if (r.expired) { keep(''); $('#epl').hidden = false; }
        if (r.token) keep(r.token);
        if (!r.ok) { m.textContent = r.msg; m.className = 'msg err'; return; }
        close(); update(r.board);
        if (window.stage) window.stage.toast(r.archived ? '바꿨습니다. 지난 응답은 시트에 보관했습니다.' : '바꿨습니다.');
      }).withFailureHandler(function (e) { b.disabled = false; m.textContent = '바꾸지 못했습니다: ' + e.message; m.className = 'msg err'; })
        .setBoard(cred, $('#eq').value, $('#ec').value, $('#eo').checked, $('#ef').checked);
    };
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') close(); });
    return function () { if (ed.classList.contains('on')) close(); else open(); };
  }
  if (live) {
    var toggle = editor();
    try { if (localStorage.getItem('board-theme') === 'dark') document.body.classList.add('dark'); } catch (e) {}
    if (window.stage) {
      window.stage.add('질문 바꾸기', 'E', toggle);
      var th = window.stage.add(document.body.classList.contains('dark') ? '밝게' : '어둡게', 'T', function () {
        var dk = document.body.classList.toggle('dark'); th.label(dk ? '밝게' : '어둡게');
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
  '<label class="ck"><input type="checkbox" id="eo">답 받는 중</label>' +
  '<label class="ck"><input type="checkbox" id="ef">지금까지 응답은 시트에 보관하고 화면을 비웁니다</label>' +
  '<label id="epl">교수 PIN<small>시트 메뉴 반응판 → 주소와 교수 PIN 보기. 한 번 맞으면 이 창을 닫을 때까지 다시 묻지 않습니다.</small><input type="password" id="ep" inputmode="numeric" maxlength="4" autocomplete="off"></label>' +
  '<div class="row"><button id="es">바꾸기</button><button id="ex" class="g">닫기</button></div><p class="msg" id="em"></p></aside>' +
  '<script src="https://kimcook-edu.pages.dev/vendor/qrcode/qrcode.js"></script>' + STAGE_JS +
  '<script>(' + boardClient_.toString() + ')();</script>';
