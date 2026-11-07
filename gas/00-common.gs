// 코드 버전: 웹 도구의 '코드 복사' 옆 번호와 화면 오른쪽 아래 번호가 같아야 새 코드가 배포된 것입니다.
var CODE_VERSION = 'v2026.11.08-a';

/* ===== 공통 도우미: 모든 레시피 맨 위에 한 번만 둡니다 =====
   post_  : 수업 서버(kimcook-edu.pages.dev)로 보내는 공통 통로. 수업 코드는 스크립트 속성에만 저장됩니다.
   ai_    : 글을 쓰는 AI(Claude)에게 부탁합니다.
   jev_   : 판정 모델 Jev에게 예/아니오, 여럿 중 하나, 단계 점수를 확률로 묻습니다. 글은 쓰지 않고 고르기만 합니다.
   ss_ sheet_ throttle_ page_ : 시트 찾기, 시트 만들기, 연타 막기, 웹 화면 만들기 */
// 편집기의 ▶ 실행이 맨 처음 함수를 고르므로 start 를 가장 위에 둡니다: 시트 탭, 메뉴, 자동 실행을 한 번에 준비합니다.
function start() {
  if (typeof setup !== 'function') { say_('이 코드에는 start 단계가 없습니다. 만들어진 함수를 직접 실행하세요.'); return; }
  setup();
  if (typeof onOpen === 'function') onOpen();
}

// 편집기 실행 로그와 시트 알림창에 같이 알립니다(자동 실행 중에는 알림창이 없어서 로그만 남습니다).
function say_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

var SERVER = 'https://kimcook-edu.pages.dev';

function setClassCode() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('수업 코드', '강사가 알려 준 수업 코드를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var code = r.getResponseText().trim();
  if (!code) { ui.alert('수업 코드가 비어 있습니다.'); return; }
  PropertiesService.getScriptProperties().setProperty('CLASS_CODE', code);
  try {
    post_('/api/jev', { check: true });
    ui.alert('수업 코드를 저장했고 연결도 확인했습니다.');
  } catch (e) {
    ui.alert('저장했지만 연결 확인에 실패했습니다.\n' + e.message);
  }
}

function post_(path, body) {
  var code = PropertiesService.getScriptProperties().getProperty('CLASS_CODE') || (typeof EMBEDDED_CODE === 'string' ? EMBEDDED_CODE : '');
  if (!code) throw new Error('수업 코드가 아직 없습니다. 웹 도구에서 수업 코드를 넣고 코드를 다시 복사하거나, 시트 메뉴의 "수업 코드 입력"을 실행하세요.');
  body.code = code;
  var last;
  for (var i = 0; i < 3; i++) {
    var res = UrlFetchApp.fetch(SERVER + path, { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    var status = res.getResponseCode(), text = res.getContentText();
    if (status === 200) return text;
    var msg = text.slice(0, 200), wait = 0;
    try {
      var j = JSON.parse(text);
      msg = j.error || msg;
      wait = j.retryAfter || 0;
      if (j.detail) msg += ' (' + j.detail + ')';
    } catch (e) {}
    last = new Error('AI ONE ' + status + ': ' + msg);
    if ((status === 429 || status === 503) && i < 2 && wait <= 20) { Utilities.sleep(Math.max(wait, 2) * 1000); continue; }
    throw last;
  }
  throw last;
}

// prompt: 글자 하나, 또는 [{role: 'user' | 'assistant', content: '…'}] 배열(마지막은 user).
// opts.images: [{name: '칠판.jpg', url: 'data:image/jpeg;base64,…'}]  opts.files: [{name: '강의계획서', text: '…'}] (긴 자료는 여기에)
function ai_(prompt, opts) {
  opts = opts || {};
  var body = { ex: 'gas-run', messages: typeof prompt === 'string' ? [{ role: 'user', content: prompt }] : prompt };
  if (opts.images) body.images = opts.images;
  if (opts.files) body.files = opts.files;
  var lines = post_('/api/ai', body).split('\n'), out = '';
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf('data:') !== 0) continue;
    var ev;
    try { ev = JSON.parse(lines[i].slice(5)); } catch (e) { continue; }
    if (ev.error) throw new Error(ev.error);
    if (ev.delta) out += ev.delta;
  }
  return out.trim();
}

// state: 판정할 글(또는 객체). questions: {이름: {type: 'noul' | 'choice' | 'score', instructions: '…', criteria: …}}
// 돌려주는 값: {이름: {type:'noul', p} | {type:'choice', choice, confidence, probs} | {type:'score', score, level, label, confidence, probs}}
function jev_(state, questions) {
  return JSON.parse(post_('/api/jev', { state: state, questions: questions })).answers;
}

function ss_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  var a = SpreadsheetApp.getActive();
  if (!a) throw new Error('시트 메뉴에서 "처음 준비"를 먼저 실행하세요.');
  p.setProperty('SS_ID', a.getId());
  return a;
}

function sheet_(name, header) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) { sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1); }
  }
  return sh;
}

// 학생에게 줄 웹 앱 주소. ScriptApp.getService().getUrl() 은 쓰지 않습니다: 편집기에서만 열리는 /dev 주소나
// 있지도 않은 배포 번호의 /exec 주소를 돌려줄 때가 있어, 학생 폰에 "현재 파일을 열 수 없습니다"가 뜹니다.
// 대신 교수가 배포 창에서 복사한 /exec 주소를 시트 메뉴 "웹 앱 주소 넣기"로 한 번 넣어 두고(WEB_URL) 그것만 씁니다.
var EXEC_RE = /^https:\/\/script\.google\.com\/(?:macros\/s\/[\w-]+|a\/(?:macros\/)?[\w.-]+\/(?:macros\/)?s\/[\w-]+)\/exec$/;
// 주소 정리: 뒤의 ?…, #…, / 와 여러 계정 로그인 때 붙는 /u/숫자 를 떼고, /exec 웹 앱 주소가 아니면 '' 를 돌려줍니다.
// 학교 계정 주소(/a/macros/학교/s/…, /a/학교/macros/s/…)는 /macros/s/… 로 바꿉니다. 학교 주소는 학교 로그인을 요구해
// 로그인하지 않은 학생 폰에 "현재 파일을 열 수 없습니다"가 뜹니다. 액세스 "모든 사용자"면 /macros/s/… 주소로 누구나 열립니다.
function execUrl_(raw) {
  var u = String(raw || '').trim().replace(/[?#][\s\S]*$/, '').replace(/\/+$/, '').replace(/\/u\/\d+(?=\/s\/)/, '');
  if (!EXEC_RE.test(u)) return '';
  return u.replace(/^https:\/\/script\.google\.com\/a\/(?:macros\/)?[\w.-]+\/(?:macros\/)?s\//, 'https://script.google.com/macros/s/');
}
function webUrl_() { return execUrl_(PropertiesService.getScriptProperties().getProperty('WEB_URL')); }
// 시트 메뉴 "웹 앱 주소 넣기". 받는 값이 없는 함수라 웹 화면(google.script.run)에서는 주소를 바꿀 수 없습니다.
function setWebUrl() {
  menuOnly_();
  var ui = SpreadsheetApp.getUi(), now = webUrl_();
  var r = ui.prompt('웹 앱 주소 넣기', '배포 → 배포 관리에서 복사한 웹 앱 URL(끝이 /exec)을 붙여 넣으세요.' + (now ? '\n지금 주소: ' + now : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return now;
  var raw = r.getResponseText().trim(), u = execUrl_(raw);
  if (!u) {
    ui.alert(/\/dev\b/.test(raw) ? '끝이 /dev 인 주소는 편집하는 사람만 열 수 있는 시험용입니다. 배포 → 새 배포(또는 배포 관리)의 웹 앱 URL(/exec)을 넣으세요.' : '웹 앱 주소가 아닙니다. https://script.google.com/…/exec 모양이어야 합니다.');
    return now;
  }
  PropertiesService.getScriptProperties().setProperty('WEB_URL', u);
  ui.alert('저장했습니다: ' + u);
  return u;
}

// 학생이 쓴 글을 시트에 넣을 때: = + - @ 로 시작하면 시트가 수식으로 계산하므로, 앞에 ' 를 붙여 글자 그대로 둡니다.
function cell_(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }
function row_(sh, arr) { sh.appendRow(arr.map(cell_)); }

// key 하나당 sec초 안에 max번까지만 true. 학생이 버튼을 마구 눌러도 수업의 AI 한도를 지키려는 장치입니다.
function throttle_(key, max, sec) {
  var cache = CacheService.getScriptCache(), n = Number(cache.get('t:' + key) || 0);
  if (n >= max) return false;
  cache.put('t:' + key, String(n + 1), sec);
  return true;
}

var FONT_LINK = '<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css"><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap">';

// 실습실 본 페이지와 같은 말씨: 종이색 바탕, 잉크색 글씨, 1px 선, 각진 모서리, 빨강 한 가지 포인트.
var PAGE_CSS = [
  ':root{--ink:#1a1a1a;--ink2:#4a4a4a;--mut:#8e8e8e;--paper:#f2f2f2;--sheet:#f8f8f8;--line:#d2d2d2;--line2:#bcbcbc;--red:#e54536}',
  '*{box-sizing:border-box}',
  '.ko{letter-spacing:.04em}',
  'body{margin:0;word-break:keep-all;background:var(--paper);color:var(--ink);font:17px/1.65 Pretendard,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;-webkit-font-smoothing:antialiased}',
  'main{max-width:620px;margin:0 auto;padding:40px 20px 88px}',
  '.kick{margin:0 0 20px;padding-left:10px;border-left:3px solid var(--red);font:700 11px/1.1 Montserrat,Pretendard,sans-serif;letter-spacing:.28em;color:var(--ink2)}',
  'h1{margin:0 0 22px;font-size:clamp(24px,6vw,33px);font-weight:800;letter-spacing:-.025em;line-height:1.25;text-wrap:balance}',
  'h2{margin:34px 0 12px;padding-top:12px;border-top:1px solid var(--ink);font:700 12px/1 Montserrat,Pretendard,sans-serif;letter-spacing:.22em;color:var(--ink2)}',
  'textarea,input,select{display:block;width:100%;margin:6px 0 12px;padding:12px 14px;border:1px solid var(--line2);border-radius:2px;background:#fff;color:inherit;font:inherit;font-size:16px}',
  'textarea:focus,input:focus,select:focus{outline:2px solid var(--ink);outline-offset:-1px;border-color:var(--ink)}',
  'textarea{min-height:120px;resize:vertical}',
  'button{padding:12px 24px;border:1px solid var(--ink);border-radius:2px;background:var(--ink);color:#fff;font:inherit;font-weight:700;cursor:pointer}',
  'button:hover{background:#000}button:disabled{opacity:.35;cursor:default}',
  'button.alt{background:transparent;color:var(--ink)}button.alt:hover{background:#fff}',
  '.muted{color:var(--mut);font-size:.9rem}',
  '.msg{margin:16px 0;padding:12px 16px;border-left:3px solid var(--ink);background:var(--sheet)}',
  '.msg.err{border-left-color:var(--red);color:#8a1f12}',
  '.chip{display:inline-block;padding:2px 9px;border:1px solid var(--line2);font-size:.82rem}',
  '.bubble{max-width:94%;margin:0 0 10px;padding:11px 15px;border:1px solid var(--line);background:#fff;font-size:1.02rem;word-break:break-word}',
  '.me{margin-left:auto;border-color:var(--ink);background:var(--ink);color:#fff}',
  '.row{display:flex;flex-wrap:wrap;align-items:center;gap:12px}',
  '.gl{margin-top:14px;font:700 11px/1.4 Montserrat,Pretendard,sans-serif;letter-spacing:.14em;color:var(--ink2)}',
  '.gauge{height:6px;margin:6px 0 18px;background:var(--line)}',
  '.gauge i{display:block;width:0;height:100%;background:var(--ink);transition:width 1.2s cubic-bezier(.2,.8,.2,1)}',
  '.gauge.hot i{background:var(--red)}',
  '.rise{animation:rise .55s cubic-bezier(.2,.8,.2,1) both}@keyframes rise{from{opacity:0;transform:translateY(12px)}}'
].join('');

// 글자가 한 글자씩 써지는 효과. 브라우저 안에서만 움직이고, 움직일 수 없는 환경에서는 전체 글이 바로 보입니다.
// 페이지 스크립트 맨 앞에 TYPE_JS 를 붙여 typeOut(요소, 글) 로 부릅니다.
function typeOut_(el, text) {
  text = String(text);
  el.textContent = text;
  if (typeof requestAnimationFrame !== 'function') return;
  var n = text.length, t0 = Date.now(), dur = Math.min(2800, 500 + n * 26);
  el.textContent = '';
  (function step() {
    var i = Math.ceil(n * Math.min(1, (Date.now() - t0) / dur));
    el.textContent = text.slice(0, i);
    if (i < n) requestAnimationFrame(step);
  })();
}
var TYPE_JS = typeOut_.toString().replace('function typeOut_', 'function typeOut');

// 버튼 결과 알림(모든 웹 화면 공통): 화면 아래 가운데에 뜹니다.
// toast('…했어요') 는 먹색으로 3초 동안 보였다가 저절로 사라집니다(아래 빨간 선이 줄어듭니다).
// fail('…을 못했어요. 이유') 는 빨간 바탕으로, 닫기 버튼을 누를 때까지 남습니다. 글은 textContent로만 넣습니다.
var TOAST_CSS = [
  '.toast{position:fixed;left:50%;bottom:clamp(20px,4vh,40px);z-index:90;display:flex;align-items:flex-start;gap:12px;width:max-content;max-width:min(600px,calc(100vw - 32px));padding:12px 16px 14px;overflow:hidden;background:#141414;color:#fff;border-radius:2px;font:500 clamp(14px,1.1vw,18px)/1.5 Pretendard,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;word-break:keep-all;transform:translate(-50%,14px);opacity:0;pointer-events:none;transition:transform .35s cubic-bezier(.2,.8,.2,1),opacity .25s}',
  '.toast.on{transform:translate(-50%,0);opacity:1;pointer-events:auto}',
  '.toast b{flex:none;padding-right:12px;border-right:1px solid rgba(255,255,255,.35);font-size:.8em;font-weight:800;letter-spacing:.12em;line-height:1.9}',
  '.toast span{white-space:pre-wrap;overflow-wrap:anywhere}',
  '.toast i{position:absolute;left:0;right:0;bottom:0;height:3px;background:#e54536;transform-origin:0 50%;transform:scaleX(0)}',
  '.toast.run i{animation:tbar 3s linear forwards}@keyframes tbar{from{transform:scaleX(1)}to{transform:scaleX(0)}}',
  '.toast.bad{background:#e54536}',
  '.toast button{flex:none;margin:-3px -6px -3px 4px;padding:3px 11px;border:1px solid #fff;border-radius:2px;background:transparent;color:#fff;font:600 .9em Pretendard,sans-serif;cursor:pointer}',
  '.toast button:hover{background:#fff;color:#e54536}',
  '@media (prefers-reduced-motion:reduce){.toast{transition:none}.toast.run i{animation:none}}'
].join('');

function toastKit_() {
  var box = null, tt = 0;
  function mk(tag, text) { var e = document.createElement(tag); if (text != null) e.textContent = text; return e; }
  return function (msg, bad) {
    box = box || document.querySelector('#toast');
    if (!box) return;
    clearTimeout(tt); box.textContent = '';
    box.className = 'toast'; void box.offsetWidth;
    if (box.setAttribute) box.setAttribute('role', bad ? 'alert' : 'status');
    box.appendChild(mk('b', bad ? '실패' : '완료')); box.appendChild(mk('span', String(msg)));
    if (bad) { var x = mk('button', '닫기'); x.type = 'button'; x.onclick = function () { box.className = 'toast bad'; }; box.appendChild(x); }
    else { box.appendChild(mk('i')); tt = setTimeout(function () { box.className = 'toast'; }, 3000); }
    box.className = 'toast on ' + (bad ? 'bad' : 'run');
  };
}
// page_ 가 모든 화면의 맨 앞에 붙입니다. 화면 스크립트는 toast(글) 와 fail(글) 을 부르면 됩니다.
var TOAST_JS = '<div class="toast" id="toast" aria-live="polite"></div><script>var toast_ = (' + toastKit_.toString() + ')();function toast(m) { toast_(m, false); }function fail(m) { toast_(m, true); }</script>';

// 웹 화면 하나를 만듭니다. body에는 <script>까지 넣습니다. opts.raw 이면 기본 틀 없이 opts.css 만 씁니다(강의장 큰 화면용).
// ALLOWALL: 파워포인트의 웹 뷰어 같은 다른 페이지 안에 이 화면을 넣을 수 있게 합니다.
function page_(title, body, opts) {
  opts = opts || {};
  var css = (opts.raw ? (opts.css || '') : PAGE_CSS + (opts.css || '')) + TOAST_CSS;
  var inner = TOAST_JS + (opts.raw ? body : '<main>' + body + '</main>');
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>' + title + '</title>' + FONT_LINK + '<style>' + css + '</style></head><body>' + inner + '<div style="position:fixed;right:10px;bottom:6px;font:600 10px/1 Montserrat,sans-serif;letter-spacing:.12em;color:#8e8e8e;opacity:.7;pointer-events:none">' + CODE_VERSION + '</div></body></html>')
    .setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 교수용 PIN: 4자리 숫자를 스크립트 속성에 둡니다(시트에는 적지 않아 학생 답이나 화면에 섞여 나올 수 없습니다). 없으면 만듭니다.
// 교수는 시트 메뉴의 "주소와 PIN 보기"에서 확인합니다. name: 'board'(반응판) 'class'(튜터 교수 화면)
function pin_(name) {
  var p = PropertiesService.getScriptProperties(), k = 'PIN_' + name, v = String(p.getProperty(k) || '');
  if (!/^\d{4}$/.test(v)) { v = String(1000 + Math.floor(Math.random() * 9000)); p.setProperty(k, v); }
  return v;
}

// 교수 확인. cred는 PIN 4자리, 또는 PIN이 맞았을 때 내준 화면용 표(6시간 유효)입니다.
// 표를 가진 화면은 계속 쓸 수 있고, PIN은 10분 동안 10번 틀리면 새로 넣는 PIN을 받지 않습니다(무작위 대입 방지). 세는 일은 잠금 안에서 합니다.
function auth_(cred, name) {
  cred = String(cred || '').trim();
  var cache = CacheService.getScriptCache();
  if (/^t[0-9a-f]{24,}$/.test(cred)) return cache.get('tok:' + name + ':' + cred) ? { ok: true, token: cred } : { ok: false, expired: true, msg: '화면을 연 지 오래되었습니다. PIN을 다시 넣어 주세요.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var key = 'pinfail:' + name, fails = Number(cache.get(key) || 0);
    if (fails >= 10) return { ok: false, locked: true, msg: 'PIN을 여러 번 틀렸습니다. 10분 뒤에 다시 하세요.' };
    if (/^\d{4}$/.test(cred) && cred === pin_(name)) {
      var t = 't' + Utilities.getUuid().replace(/-/g, '');
      cache.put('tok:' + name + ':' + t, '1', 21600);
      return { ok: true, token: t };
    }
    cache.put(key, String(fails + 1), 600);
    return { ok: false, msg: 'PIN이 맞지 않습니다.' };
  } finally { lock.releaseLock(); }
}

// 시트 메뉴(편집기)에서만 돌아야 하는 함수의 맨 앞에 둡니다. 웹 앱 화면에서 google.script.run 으로 부르면 멈춥니다.
function menuOnly_() {
  try { SpreadsheetApp.getUi(); } catch (e) { throw new Error('시트 메뉴에서만 실행할 수 있습니다.'); }
}

// 강의장 큰 화면 공통 도구: 오른쪽 아래 버튼 줄(전체화면 F), 마우스를 멈추면 버튼과 커서를 숨김, 전체화면 동안 화면 꺼짐 방지.
// 다른 화면은 window.stage.add(이름, 단축키, 함수)로 버튼을 더합니다.
var STAGE_CSS = [
  '.kit{position:fixed;right:clamp(10px,1.4vw,26px);bottom:clamp(10px,1.4vw,26px);z-index:30;display:flex;gap:6px;transition:opacity .45s ease}',
  'body.idle .kit{opacity:0;pointer-events:none}body.idle{cursor:none}',
  '.kit button{display:flex;align-items:center;gap:8px;padding:9px 12px;border:1px solid var(--fg);border-radius:2px;background:var(--bg);color:var(--fg);font:600 13px/1 Pretendard,sans-serif;cursor:pointer}',
  '.kit button:hover{background:var(--fg);color:var(--bg)}',
  '.kit kbd{font:700 10px/1 Montserrat,sans-serif;letter-spacing:.08em;opacity:.55}',
  '@media (prefers-reduced-motion:reduce){*{transition-duration:.01ms!important;animation-duration:.01ms!important}}'
].join('');

function stageKit_() {
  if (typeof document === 'undefined' || !document.body || !document.addEventListener) return;
  var bar = document.createElement('div'), keys = {}, idle = 0, lock = null;
  bar.className = 'kit'; document.body.appendChild(bar);
  // 알림은 공통 토스트(TOAST_JS)를 씁니다.
  function toast(t) { if (typeof window.toast === 'function') window.toast(t); }
  function oops(t) { if (typeof window.fail === 'function') window.fail(t); }
  function wake() { document.body.classList.remove('idle'); clearTimeout(idle); idle = setTimeout(function () { document.body.classList.add('idle'); }, 2600); }
  ['mousemove', 'mousedown', 'keydown', 'touchstart'].forEach(function (e) { document.addEventListener(e, wake, { passive: true }); });
  wake();
  function isFs() { return !!(document.fullscreenElement || document.webkitFullscreenElement); }
  function hint() { oops('전체화면을 못했어요. 이 창에서는 버튼이 막혀 있습니다. 키보드로 하세요: 윈도우 F11, 맥 Control + Command + F'); }
  function fs() {
    if (isFs()) { (document.exitFullscreen || document.webkitExitFullscreen).call(document); return; }
    var d = document.documentElement, req = d.requestFullscreen || d.webkitRequestFullscreen;
    if (!req || document.fullscreenEnabled === false) { hint(); return; }
    try { var p = req.call(d); if (p && p.catch) p.catch(hint); } catch (e) { hint(); }
  }
  function add(label, key, fn) {
    var b = document.createElement('button'), k = document.createElement('kbd');
    b.type = 'button'; b.appendChild(document.createTextNode(label)); k.textContent = key; b.appendChild(k);
    b.onclick = function () { fn(); b.blur(); };
    bar.insertBefore(b, bar.firstChild); keys[key.toLowerCase()] = fn;
    return { el: b, label: function (t) { b.firstChild.textContent = t; } };
  }
  var fsb = add('전체화면', 'F', fs);
  document.addEventListener('fullscreenchange', function () {
    fsb.label(isFs() ? '전체화면 끄기' : '전체화면');
    try { if (isFs() && navigator.wakeLock) navigator.wakeLock.request('screen').then(function (l) { lock = l; }, function () {}); else if (lock) { lock.release(); lock = null; } } catch (e) {}
  });
  document.addEventListener('keydown', function (e) {
    var t = e.target && e.target.tagName;
    if (t === 'INPUT' || t === 'TEXTAREA' || t === 'SELECT' || e.metaKey || e.ctrlKey || e.altKey) return;
    var f = keys[String(e.key || '').toLowerCase()];
    if (f) { e.preventDefault(); f(); }
  });
  window.stage = { add: add, toast: toast, fail: oops };
}
var STAGE_JS = '<script>(' + stageKit_.toString() + ')();</script>';

// 메일을 보냅니다. 글자판(text)을 그대로 두고, 같은 내용을 실습실과 같은 말씨의 HTML로도 함께 보냅니다.
// "[제목]" 한 줄은 작은 소제목이 됩니다.
function mail_(to, subject, text) {
  var esc = function (x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var body = String(text).split('\n').map(function (l) {
    if (/^\[.+\]$/.test(l)) return '<div style="margin:24px 0 6px;font:700 11px/1.4 Arial,sans-serif;letter-spacing:.18em;color:#8e8e8e">' + esc(l.slice(1, -1)) + '</div>';
    if (!l.trim()) return '<div style="height:10px"></div>';
    return '<div>' + esc(l).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" style="color:#1a1a1a">$1</a>') + '</div>';
  }).join('');
  var html = '<div style="max-width:560px;padding:8px 2px 28px;font:15px/1.75 \'Malgun Gothic\',\'Apple SD Gothic Neo\',sans-serif;color:#1a1a1a">' +
    '<div style="border-top:3px solid #1a1a1a;padding-top:12px;font:700 11px/1 Arial,sans-serif;letter-spacing:.26em;color:#8e8e8e">PROF.KIMCOOK · AI CLASS LAB</div>' +
    '<div style="margin:18px 0 16px;font-size:20px;line-height:1.35;font-weight:700">' + esc(subject) + '</div>' + body + '</div>';
  MailApp.sendEmail(to, subject, text, { htmlBody: html });
}
