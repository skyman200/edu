/* ===== 레시피 1. 실시간 교실 반응판 =====
   1) 새 구글 시트 → 확장 프로그램 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.
   2) 시트를 새로고침하면 [반응판] 메뉴가 생깁니다. 처음 준비 → 수업 코드 입력을 차례로 실행합니다.
   3) 배포 → 새 배포 → 웹 앱: 실행 사용자 "나", 액세스 "모든 사용자". 주소 보기 메뉴에서 학생용과 강의장용 주소를 봅니다.
   질문과 묶음 기준은 [설정] 시트에서 바꿉니다. 답은 [응답] 시트에 쌓이고, Jev가 묶음을 고릅니다. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('반응판')
    .addItem('처음 준비', 'setup').addItem('수업 코드 입력', 'setClassCode').addItem('주소 보기', 'showUrls').addToUi();
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
    s.setColumnWidth(1, 190); s.setColumnWidth(2, 520);
  }
  sheet_('응답', ['시간', '응답', '묶음', '확신도', '상태']);
  SpreadsheetApp.getUi().alert('준비됐습니다. 이어서 메뉴의 수업 코드 입력을 실행하고, 배포 → 웹 앱으로 배포하세요.\n질문과 묶음은 [설정] 시트에서 바꿉니다.');
}

function showUrls() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? '학생용: ' + url + '\n강의장 화면: ' + url + '?v=board' : '아직 웹 앱으로 배포하지 않았습니다. 배포 → 새 배포 → 웹 앱을 먼저 하세요.');
}

function cfg_() {
  var v = ss_().getSheetByName('설정').getRange('B1:B4').getValues();
  var cats = String(v[1][0]).split(',').map(function (x) { return x.trim(); }).filter(String).slice(0, 10);
  return { q: String(v[0][0]), cats: cats, open: String(v[2][0]).trim().toUpperCase() === 'O', refresh: Math.max(2, Number(v[3][0]) || 4) };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.v === 'board') return page_('반응판', BOARD_HTML);
  return page_('내 생각 보내기', ASK_HTML);
}

function getQuestion() {
  var c = cfg_();
  return { q: c.q, open: c.open };
}

function submitAnswer(text, sid) {
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
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sheet_('응답').appendRow([new Date(), text, cat, conf, status]); } finally { lock.releaseLock(); }
  return { ok: true, cat: cat };
}

function getBoard() {
  var cache = CacheService.getScriptCache(), hit = cache.get('board');
  if (hit) return JSON.parse(hit);
  var c = cfg_(), sh = sheet_('응답'), n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [], items = [];
  rows.forEach(function (r) { if (r[4] !== '숨김') items.push({ t: String(r[1]), c: String(r[2] || '') }); });
  var out = { q: c.q, cats: c.cats.concat(c.cats.length ? ['기타'] : []), open: c.open, refresh: c.refresh, items: items.slice(-200) };
  cache.put('board', JSON.stringify(out), 3);
  return out;
}

var ASK_HTML = '<h1 id="q">…</h1><textarea id="t" maxlength="140" placeholder="한두 문장으로 적어 주세요 (이름 없이)"></textarea>' +
  '<div class="row"><button id="b">보내기</button><span class="muted" id="n">0 / 140</span></div><div id="m"></div>' +
  '<script>var sid;try{sid=localStorage.getItem("sid");if(!sid){sid=Math.random().toString(36).slice(2,12);localStorage.setItem("sid",sid)}}catch(e){sid=Math.random().toString(36).slice(2,12)}' +
  'var $=function(s){return document.querySelector(s)};' +
  'google.script.run.withSuccessHandler(function(r){$("#q").textContent=r.q;if(!r.open){$("#b").disabled=true;$("#m").innerHTML="<p class=msg>지금은 답을 받지 않습니다.</p>"}}).getQuestion();' +
  '$("#t").oninput=function(){$("#n").textContent=this.value.length+" / 140"};' +
  '$("#b").onclick=function(){var t=$("#t").value;if(t.trim().length<2){return}$("#b").disabled=true;' +
  'google.script.run.withSuccessHandler(function(r){$("#b").disabled=false;var m=$("#m");m.textContent="";var p=document.createElement("p");p.className=r.ok?"msg":"msg err";' +
  'p.textContent=r.ok?"보냈습니다."+(r.cat?" 「"+r.cat+"」 쪽으로 모였어요.":""):r.msg;m.appendChild(p);if(r.ok){$("#t").value="";$("#n").textContent="0 / 140"}})' +
  '.withFailureHandler(function(e){$("#b").disabled=false;$("#m").textContent="보내지 못했습니다: "+e.message}).submitAnswer(t,sid)};</script>';

var BOARD_HTML = '<h1 id="q" style="font-size:2rem">…</h1><p class="muted" id="c"></p><div id="g" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;align-items:start"></div>' +
  '<script>var $=function(s){return document.querySelector(s)};' +
  'function draw(d){$("#q").textContent=d.q;$("#c").textContent="모인 생각 "+d.items.length+"개"+(d.open?"":" · 받기를 멈췄습니다");var cols=d.cats.length?d.cats:["모두"],g=$("#g");g.textContent="";' +
  'cols.forEach(function(name){var col=document.createElement("section"),h=document.createElement("h2"),k=d.items.filter(function(x){return d.cats.length?x.c===name||(name==="기타"&&!x.c):true});' +
  'h.textContent=name+" · "+k.length;col.appendChild(h);k.slice(-12).reverse().forEach(function(x){var b=document.createElement("div");b.className="bubble";b.textContent=x.t;col.appendChild(b)});g.appendChild(col)})}' +
  'function tick(){google.script.run.withSuccessHandler(function(d){draw(d);setTimeout(tick,d.refresh*1000)}).withFailureHandler(function(){setTimeout(tick,8000)}).getBoard()}tick();</script>';
