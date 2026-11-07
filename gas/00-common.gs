// 코드 버전: 웹 도구의 '코드 복사' 옆 번호와 화면 오른쪽 아래 번호가 같아야 새 코드가 배포된 것입니다.
var CODE_VERSION = 'v2026.11.07-d';

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

// 웹 화면 하나를 만듭니다. body에는 <script>까지 넣습니다. opts.raw 이면 기본 틀 없이 opts.css 만 씁니다(강의장 큰 화면용).
// ALLOWALL: 파워포인트의 웹 뷰어 같은 다른 페이지 안에 이 화면을 넣을 수 있게 합니다.
function page_(title, body, opts) {
  opts = opts || {};
  var css = opts.raw ? (opts.css || '') : PAGE_CSS + (opts.css || '');
  var inner = opts.raw ? body : '<main>' + body + '</main>';
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>' + title + '</title>' + FONT_LINK + '<style>' + css + '</style></head><body>' + inner + '<div style="position:fixed;right:10px;bottom:6px;font:600 10px/1 Montserrat,sans-serif;letter-spacing:.12em;color:#8e8e8e;opacity:.7;pointer-events:none">' + CODE_VERSION + '</div></body></html>')
    .setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
