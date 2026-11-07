/* ===== 레시피 1. 실시간 교실 반응판 =====
   사용법: 웹 도구(https://kimcook-edu.pages.dev/#gas)에서 이 레시피를 고르고 안내를 따르세요. 코드를 붙여 넣고 start 를 실행하면 준비가 끝납니다.
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
  say_('준비 끝! 질문과 묶음은 [설정] 시트에서 바꿀 수 있습니다. 이제 웹 도구 화면의 다음 단계(배포)를 따라 하세요.');
}

function showUrls() {
  var url = ScriptApp.getService().getUrl();
  say_(url ? '학생용: ' + url + '\n강의장 화면: ' + url + '?v=board' : '아직 웹 앱으로 배포하지 않았습니다. 배포 → 새 배포 → 웹 앱을 먼저 하세요.');
}

function cfg_() {
  var v = ss_().getSheetByName('설정').getRange('B1:B4').getValues();
  var cats = String(v[1][0]).split(',').map(function (x) { return x.trim(); }).filter(String).slice(0, 10);
  return { q: String(v[0][0]), cats: cats, open: String(v[2][0]).trim().toUpperCase() === 'O', refresh: Math.max(2, Number(v[3][0]) || 4) };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.v === 'board') return page_('반응판', BOARD_HTML, { raw: true, css: BOARD_CSS });
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
  var out = { q: c.q, cats: c.cats.concat(c.cats.length ? ['기타'] : []), open: c.open, refresh: c.refresh, total: items.length, join: ScriptApp.getService().getUrl() || '', items: items.slice(-200) };
  cache.put('board', JSON.stringify(out), 3);
  return out;
}

var ASK_HTML = '<div class="kick">LIVE · <span class="ko">교실 반응판</span></div><h1 id="q">…</h1><textarea id="t" maxlength="140" placeholder="한두 문장으로 적어 주세요 (이름 없이)"></textarea>' +
  '<div class="row"><button id="b">보내기</button><span class="muted" id="n">0 / 140</span></div><div id="m"></div>' +
  '<script>var sid;try{sid=localStorage.getItem("sid");if(!sid){sid=Math.random().toString(36).slice(2,12);localStorage.setItem("sid",sid)}}catch(e){sid=Math.random().toString(36).slice(2,12)}' +
  'var $=function(s){return document.querySelector(s)};' +
  'google.script.run.withSuccessHandler(function(r){$("#q").textContent=r.q;if(!r.open){$("#b").disabled=true;$("#m").innerHTML="<p class=msg>지금은 답을 받지 않습니다.</p>"}}).getQuestion();' +
  '$("#t").oninput=function(){$("#n").textContent=this.value.length+" / 140"};' +
  '$("#b").onclick=function(){var t=$("#t").value;if(t.trim().length<2){return}$("#b").disabled=true;' +
  'google.script.run.withSuccessHandler(function(r){$("#b").disabled=false;var m=$("#m");m.textContent="";var p=document.createElement("p");p.className=r.ok?"msg":"msg err";' +
  'p.textContent=r.ok?"보냈습니다."+(r.cat?" 「"+r.cat+"」 쪽으로 모였어요.":""):r.msg;m.appendChild(p);if(r.ok){$("#t").value="";$("#n").textContent="0 / 140"}})' +
  '.withFailureHandler(function(e){$("#b").disabled=false;$("#m").textContent="보내지 못했습니다: "+e.message}).submitAnswer(t,sid)};</script>';

// 강의장 큰 화면: 어두운 종이색, 큰 글씨, 1px 선과 각진 칸. 색은 빨강 하나만 쓰고, 가장 많이 모인 묶음에만 붙입니다.
// 작은 칸(파워포인트에 넣을 때)에서도 글자와 간격이 칸 크기를 따라 줄어듭니다.
var BOARD_CSS = [
  ':root{--bg:#161616;--fg:#f2f2f2;--mut:#8e8e8e;--line:rgba(242,242,242,.16);--red:#e54536}',
  '*{box-sizing:border-box}html,body{height:100%}',
  'body{margin:0;overflow:hidden;background:var(--bg);color:var(--fg);font:16px/1.5 Pretendard,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;-webkit-font-smoothing:antialiased}',
  '.stage{height:100vh;display:grid;grid-template-rows:auto auto minmax(0,1fr);gap:clamp(12px,2.4vh,26px);padding:clamp(14px,2.8vw,48px)}',
  '.top{display:flex;justify-content:space-between;align-items:flex-end;gap:clamp(14px,3vw,48px)}',
  '.live{display:flex;align-items:center;gap:.9em;font:700 clamp(10px,1vw,14px)/1 Montserrat,Pretendard,sans-serif;letter-spacing:.3em;color:var(--mut)}',
  '.live:before{content:"";width:.8em;height:.8em;background:var(--red)}',
  'h1{word-break:keep-all;max-width:19em;margin:.6em 0 0;font-size:clamp(22px,4vw,64px);font-weight:800;letter-spacing:-.03em;line-height:1.15;text-wrap:balance}',
  '.side{display:flex;align-items:flex-end;gap:clamp(14px,2.4vw,40px);flex:none}',
  '.total{text-align:right;font:800 clamp(40px,7.4vw,120px)/.82 Montserrat,Pretendard,sans-serif;letter-spacing:-.05em}',
  '.total small{display:block;margin-top:1em;font:600 clamp(10px,.9vw,13px)/1 Pretendard,sans-serif;letter-spacing:.24em;color:var(--mut)}',
  '.join{display:flex;align-items:flex-end;gap:clamp(8px,1vw,16px)}.join[hidden]{display:none}',
  '.qr{width:clamp(76px,9vw,140px);aspect-ratio:1;padding:5px;background:#fff}',
  '.qr svg{display:block;width:100%;height:100%}.qr path{fill:#000}',
  '.join b{display:block;font-size:clamp(12px,1.2vw,19px);line-height:1.3}.join span{color:var(--mut);font-size:clamp(10px,.9vw,13px)}',
  '.share{display:flex;gap:2px;height:clamp(6px,1vh,10px)}',
  '.share i{display:block;min-width:0;flex:0 1 0;background:#f2f2f2;transition:flex-grow .9s cubic-bezier(.2,.8,.2,1)}',
  '.share i.k0{background:var(--red)}.share i.k1{opacity:.7}.share i.k2{opacity:.45}.share i.k3{opacity:.28}.share i.k4{opacity:.16}',
  '.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(clamp(150px,17vw,300px),1fr));gap:clamp(12px,1.6vw,26px);min-height:0}',
  '.col{display:flex;flex-direction:column;min-height:0;overflow:hidden;border-top:1px solid var(--fg);padding-top:clamp(8px,1vw,16px)}',
  '.col.lead{border-top:4px solid var(--red);padding-top:calc(clamp(8px,1vw,16px) - 3px)}',
  '.col h2{display:grid;grid-template-columns:1fr auto;align-items:end;gap:2px 10px;margin:0 0 clamp(10px,1.4vh,18px);font-size:clamp(14px,1.5vw,24px);font-weight:700;letter-spacing:-.01em;line-height:1.2}',
  '.col h2 span{grid-column:1;grid-row:2}.col h2 u{grid-column:1;grid-row:1;text-decoration:none;font:600 clamp(9px,.8vw,12px)/1 Montserrat,Pretendard,sans-serif;letter-spacing:.2em;color:var(--mut)}',
  '.col h2 em{grid-column:2;grid-row:1/3;font:800 clamp(26px,3.6vw,58px)/.9 Montserrat,Pretendard,sans-serif;font-style:normal;letter-spacing:-.04em}',
  '.col.lead h2 em{color:var(--red)}',
  '.msgs{display:flex;flex-direction:column;gap:clamp(6px,1vh,12px);min-height:0;overflow:hidden}',
  '.ko{letter-spacing:.04em}',
  '.b{padding:.1em 0 .1em .85em;word-break:keep-all;overflow-wrap:anywhere;border-left:2px solid var(--line);font-size:clamp(13px,1.4vw,22px);line-height:1.4}',
  '.b.anim{animation:in .5s cubic-bezier(.2,.8,.2,1) both}.b.last{animation-delay:1.9s}',
  '.b.new{border-left-color:var(--red)}',
  '@keyframes in{from{opacity:0;transform:translateX(-14px)}}',
  '.arr{position:fixed;left:0;right:0;bottom:7vh;z-index:5;width:min(72vw,900px);margin:0 auto;padding:clamp(14px,2vw,28px) clamp(18px,2.6vw,40px);background:#f2f2f2;color:#161616;border-left:8px solid var(--red);transform-origin:0 0}',
  '.arr[hidden]{display:none}',
  '.arr small{display:block;margin-bottom:.8em;font:700 clamp(10px,1vw,14px)/1 Montserrat,Pretendard,sans-serif;letter-spacing:.26em;color:#666}',
  '.arr p{margin:0;font-size:clamp(20px,3vw,46px);font-weight:800;letter-spacing:-.02em;line-height:1.25;word-break:keep-all}',
  '.arr.in{animation:rise .55s cubic-bezier(.2,.8,.2,1) both}',
  '.arr.go{transition:transform .7s cubic-bezier(.6,0,.2,1),opacity .7s ease}',
  '@keyframes rise{from{opacity:0;transform:translateY(48px)}}',
  '.total span.bump{animation:bump .45s ease}@keyframes bump{40%{transform:scale(1.12)}}',
  '.wait{grid-column:1/-1;align-self:center;justify-self:start;max-width:16em;color:var(--mut);font-size:clamp(16px,2.4vw,38px);font-weight:700;letter-spacing:-.02em;line-height:1.3}'
].join('');

function boardClient_() {
  var $ = function (s) { return document.querySelector(s); };
  var prev = -1, qrDone = false, token = 0, counts = {}, shownTotal = 0;
  function mk(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function qr(url) {
    if (typeof qrcode !== 'function') return '';
    var q = qrcode(0, 'M'); q.addData(url); q.make();
    var n = q.getModuleCount(), d = '';
    for (var y = 0; y < n; y++) for (var x = 0; x < n; x++) if (q.isDark(y, x)) d += 'M' + x + ',' + y + 'h1v1h-1z';
    return '<svg viewBox="-2 -2 ' + (n + 4) + ' ' + (n + 4) + '" shape-rendering="crispEdges"><path d="' + d + '"/></svg>';
  }
  // 숫자가 이전 값에서 새 값까지 올라갑니다. 애니메이션 기능이 없는 환경에서는 바로 새 값을 씁니다.
  function tween(e, from, to) {
    if (from === to || typeof requestAnimationFrame !== 'function') { e.textContent = String(to); return; }
    var t0 = Date.now();
    (function step() {
      var k = Math.min(1, (Date.now() - t0) / 700), v = 1 - Math.pow(1 - k, 3);
      e.textContent = String(Math.round(from + (to - from) * v));
      if (k < 1) requestAnimationFrame(step);
    })();
  }
  // 새 답이 화면 가운데에 크게 뜨고, 분류가 끝나면 해당 묶음으로 날아갑니다.
  function arrive(text, label, colEl) {
    var a = $('#arr');
    if (!a || !a.querySelector || typeof setTimeout !== 'function') return;
    var my = ++token, lab = a.querySelector('small'), p = a.querySelector('p');
    a.hidden = false; a.className = 'arr'; a.style.transform = ''; a.style.opacity = '';
    lab.textContent = '방금 도착 · 분류하는 중'; p.textContent = text;
    void a.offsetWidth; a.className = 'arr in';
    setTimeout(function () { if (my === token) lab.textContent = '→ ' + label; }, 1000);
    setTimeout(function () {
      if (my !== token) return;
      var r = a.getBoundingClientRect(), t = colEl.getBoundingClientRect();
      a.className = 'arr go';
      a.style.transform = 'translate(' + Math.round(t.left + 12 - r.left) + 'px,' + Math.round(t.top + 40 - r.top) + 'px) scale(.2)';
      a.style.opacity = '0';
    }, 1900);
    setTimeout(function () { if (my === token) a.hidden = true; }, 2700);
  }
  function draw(d) {
    $("#q").textContent = d.q;
    var tot = $("#c");
    tween(tot, shownTotal, d.total);
    if (d.total !== shownTotal && shownTotal > 0 && tot.className !== undefined) { tot.className = ''; void tot.offsetWidth; tot.className = 'bump'; }
    shownTotal = d.total;
    if (d.join && !qrDone) { var svg = qr(d.join); if (svg) { $("#qr").innerHTML = svg; $("#join").hidden = false; } qrDone = true; }
    var g = $("#g"), sh = $("#sh");
    g.textContent = ''; sh.textContent = '';
    if (!d.total) { g.appendChild(mk('div', 'wait', '첫 번째 생각을 기다리고 있습니다. QR을 스캔해 보내 주세요.')); prev = 0; return; }
    var cols = d.cats.length ? d.cats : ['모두'];
    var vh = typeof innerHeight === 'number' ? innerHeight : 800;
    var per = Math.max(3, Math.min(10, Math.floor((vh - 300) / 74)));
    var first = prev < 0, fresh = first ? 0 : Math.max(0, d.total - prev);
    var groups = cols.map(function (name) {
      var mine = [];
      d.items.forEach(function (x, idx) {
        if (d.cats.length ? (x.c === name || (name === '기타' && !x.c)) : true) mine.push({ t: x.t, isNew: idx >= d.items.length - fresh, last: fresh > 0 && idx === d.items.length - 1 });
      });
      return { name: name, mine: mine };
    });
    var order = groups.map(function (gr, i) { return i; }).sort(function (a, b) { return groups[b].mine.length - groups[a].mine.length || a - b; });
    var landing = null, n = 0;
    groups.forEach(function (gr, i) {
      var rank = order.indexOf(i), lead = rank === 0 && d.cats.length > 1 && gr.mine.length > (groups[order[1]] ? groups[order[1]].mine.length : 0);
      var col = mk('section', 'col' + (lead ? ' lead' : '')), h = mk('h2'), box = mk('div', 'msgs'), em = mk('em', '', String(gr.mine.length));
      var label = (i < 9 ? '0' : '') + String(i + 1);
      h.appendChild(mk('u', '', label));
      h.appendChild(mk('span', '', gr.name)); h.appendChild(em);
      tween(em, counts[gr.name] || 0, gr.mine.length); counts[gr.name] = gr.mine.length;
      col.appendChild(h);
      gr.mine.slice(-per).reverse().forEach(function (x) {
        var b = mk('div', 'b' + (x.isNew ? ' new' : '') + (first || x.isNew ? ' anim' : '') + (x.last ? ' last' : ''), x.t);
        if (first && b.style) b.style.animationDelay = Math.min(1.2, n * 0.07) + 's';
        n++;
        if (x.last) landing = { text: x.t, label: label + ' ' + gr.name, col: col };
        box.appendChild(b);
      });
      col.appendChild(box); g.appendChild(col);
      var seg = mk('i', 'k' + (lead ? 0 : Math.max(1, Math.min(rank, 4))));
      if (first && typeof requestAnimationFrame === 'function') {
        seg.style.flexGrow = '0';
        requestAnimationFrame(function () { requestAnimationFrame(function () { seg.style.flexGrow = String(gr.mine.length); }); });
      } else seg.style.flexGrow = String(gr.mine.length);
      sh.appendChild(seg);
    });
    if (landing) arrive(landing.text, landing.label, landing.col);
    prev = d.total;
  }
  function tick() {
    google.script.run.withSuccessHandler(function (d) { draw(d); setTimeout(tick, d.refresh * 1000); })
      .withFailureHandler(function () { setTimeout(tick, 8000); }).getBoard();
  }
  tick();
}

var BOARD_HTML = '<div class="stage"><div class="top"><div><div class="live"><span class="dot"></span>LIVE · <span class="ko">교실 반응판</span></div><h1 id="q">…</h1></div>' +
  '<div class="side"><div class="join" id="join" hidden><div class="qr" id="qr"></div><div><b>휴대폰으로 참여</b><span>QR을 스캔하세요</span></div></div>' +
  '<div class="total"><span id="c">0</span><small>모인 생각</small></div></div></div>' +
  '<div class="share" id="sh"></div><div class="cols" id="g"></div><div class="arr" id="arr" hidden><small></small><p></p></div></div>' +
  '<script src="https://kimcook-edu.pages.dev/vendor/qrcode/qrcode.js"></script>' +
  '<script>(' + boardClient_.toString() + ')();</script>';
