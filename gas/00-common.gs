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

var PAGE_CSS = 'body{margin:0;background:#f6f3ec;color:#1b1a17;font:17px/1.6 -apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif}' +
  'main{max-width:720px;margin:0 auto;padding:24px 18px 60px}h1{font-size:1.35rem;line-height:1.35;margin:0 0 14px}h2{font-size:1rem;margin:22px 0 8px;color:#6b675e}' +
  'textarea,input,select{width:100%;box-sizing:border-box;font:inherit;padding:10px 12px;border:1px solid #cfc8b6;border-radius:8px;background:#fffdf8;color:inherit;margin:4px 0 10px}' +
  'textarea{min-height:96px;resize:vertical}button{font:inherit;font-weight:600;padding:10px 18px;border:0;border-radius:8px;background:#1b1a17;color:#fff;cursor:pointer}' +
  'button:disabled{opacity:.45;cursor:default}button.alt{background:#fffdf8;color:#1b1a17;border:1px solid #cfc8b6}.muted{color:#6b675e;font-size:.9rem}' +
  '.msg{margin:12px 0;padding:10px 12px;border-radius:8px;background:#ece6d6}.err{background:#f6dcd5;color:#7c2412}.chip{display:inline-block;padding:2px 9px;border-radius:99px;background:#ece6d6;font-size:.85rem}' +
  '.bubble{background:#fffdf8;border:1px solid #d9d3c4;border-radius:10px;padding:8px 12px;margin:0 0 8px;font-size:1.02rem;word-break:break-word}' +
  '.me{background:#1b1a17;color:#fff;border-color:#1b1a17}.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}';

// 웹 화면 하나를 만듭니다. body에는 <script>까지 넣습니다.
function page_(title, body) {
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>' + title + '</title><style>' + PAGE_CSS + '</style></head><body><main>' + body + '</main></body></html>')
    .setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
