// 고칠 곳만 받아 코드에 끼워 넣습니다. 형식(줄 맨 앞에 쓰고, 한 답에 여러 개 가능):
// <<<<<<< 찾기 / (지금 코드의 그 부분) / ======= / (바꿀 내용) / >>>>>>> 바꾸기
// 찾을 글이 코드에 한 번만 있어야 바꿉니다. 들여쓰기가 달라도 줄 내용이 같으면 찾습니다.
var VibePatch = (function () {
  var BLOCK = /^[ \t]*<<<<<<<[^\n]*\n([\s\S]*?)\n[ \t]*=======[ \t]*\n([\s\S]*?)\n?[ \t]*>>>>>>>[^\n]*$/gm;
  var MARK = /^[ \t]*(?:<<<<<<<|=======[ \t]*$|>>>>>>>)/m;
  function all(hay, s) { var at = []; for (var i = hay.indexOf(s); s && i >= 0; i = hay.indexOf(s, i + 1)) at.push(i); return at; }
  function locate(code, old) {
    var ex = all(code, old);
    if (ex.length) return ex.length === 1 ? { a: ex[0], b: ex[0] + old.length } : { many: true };
    var want = old.split('\n').map(function (l) { return l.trim(); });
    while (want.length && !want[0]) want.shift();
    while (want.length && !want[want.length - 1]) want.pop();
    if (!want.length) return null;
    var lines = code.split('\n'), offs = [], o = 0, hits = [];
    lines.forEach(function (l) { offs.push(o); o += l.length + 1; });
    for (var i = 0; i + want.length <= lines.length; i++) {
      var ok = true;
      for (var k = 0; k < want.length && ok; k++) if (lines[i + k].trim() !== want[k]) ok = false;
      if (ok) hits.push(i);
    }
    if (hits.length !== 1) return hits.length ? { many: true } : null;
    var a = hits[0], b = a + want.length - 1;
    return { a: offs[a], b: offs[b] + lines[b].length, indent: lines[a].match(/^[ \t]*/)[0], loose: true };
  }
  // text 에 찾기/바꾸기 블록이 있는지
  function has(text) { return /^[ \t]*<<<<<<<[^\n]*\n[\s\S]*?\n[ \t]*=======[ \t]*\n/m.test(String(text)); }
  // { code, applied, failed:[{label, why}] }. failed 가 하나라도 있으면 code 는 처음 그대로입니다.
  function apply(base, text) {
    var out = { code: String(base || ''), applied: 0, failed: [] }, rest = String(text || ''), cur = out.code;
    var blocks = [];
    rest.replace(BLOCK, function (m, o, n) { blocks.push([o, n]); return ''; });
    rest = rest.replace(BLOCK, '');
    if (MARK.test(rest)) out.failed.push({ label: '<<<<<<<', why: '바꿀 곳 표시가 끝나지 않았거나 형식이 깨졌습니다' });
    blocks.forEach(function (b) {
      var old = b[0], neu = b[1], label = (old.split('\n').filter(function (l) { return l.trim(); })[0] || '').trim().slice(0, 60);
      if (MARK.test(old) || MARK.test(neu)) { out.failed.push({ label: label, why: '바꿀 곳 표시가 서로 엉켜 있습니다' }); return; }
      var at = locate(cur, old);
      if (!at || at.many) { out.failed.push({ label: label, why: at ? '같은 글이 여러 곳에 있어 어디인지 알 수 없습니다' : '지금 코드에서 그 글을 찾지 못했습니다' }); return; }
      if (at.loose && at.indent) {
        var first = neu.split('\n').filter(function (l) { return l.trim(); })[0] || '';
        if (!/^[ \t]/.test(first)) neu = neu.split('\n').map(function (l) { return l.trim() ? at.indent + l : l; }).join('\n');
      }
      cur = cur.slice(0, at.a) + neu + cur.slice(at.b);
      out.applied++;
    });
    if (!blocks.length && !out.failed.length) out.failed.push({ label: '', why: '바꿀 곳(<<<<<<< 찾기) 형식이 없습니다' });
    if (!out.failed.length) out.code = cur;
    return out;
  }
  // AI 에게 보내는 고치기 요청. 전체 코드를 다시 쓰게 하지 않아 답이 짧습니다(비용은 거의 답 길이에 비례).
  var FORMAT = '고친 곳만 아래 형식으로 알려 줘. 전체 코드를 다시 쓰지 마. 표시줄은 들여쓰기 없이 줄 맨 앞에, 한 코드 블록에 여러 개 써도 돼.\n'
    + '<<<<<<< 찾기\n(지금 코드에서 바꿀 부분을 한 글자도 틀리지 않게, 3줄 안팎으로)\n=======\n(바꾼 내용)\n>>>>>>> 바꾸기\n'
    + '찾을 부분은 코드 안에서 딱 한 곳이어야 해. 설명은 한 줄만.';
  function fixPrompt(ask, code) {
    return ask.trim() + '\n나머지는 그대로 두고, ' + FORMAT + (code ? '\n\n[지금 코드]\n```html\n' + code.trim() + '\n```' : '');
  }
  function errorPrompt(errors, code) {
    return '이 웹앱을 실행하니 아래 오류가 났어.\n' + errors.slice(0, 5).map(function (x) { return '- ' + x; }).join('\n')
      + '\n원인을 한 줄로 설명하고, ' + FORMAT + (code ? '\n\n[지금 코드]\n```html\n' + code.trim() + '\n```' : '');
  }
  // 처음부터 만들 때 AI 채팅에 붙여 넣는 글. 괄호 안만 내 수업에 맞게 고칩니다.
  var SCRATCH = ['수업에서 쓸 웹앱을 만들어 줘.', '- 앱 이름: (예: 영어 면접 질문 뽑기)', '- 쓰는 사람: (예: 내가 수업 중 화면에 띄워 씀 / 학생이 휴대폰으로 씀)',
    '- 하는 일: (한 문장)', '- 꼭 있어야 할 기능: (한 줄에 하나)', '- 넣을 내용: (내 과목의 실제 내용 8~12개)',
    '', '규칙: HTML 파일 하나에 CSS와 JavaScript를 모두 넣고, 외부 라이브러리와 서버와 로그인은 쓰지 마. 글자는 프로젝터에서도 보이게 크게.',
    '디자인은 밝은 배경(#F2F2F2), 거의 검은 글씨(#1A1A1A), 강조색 하나(#E54536), 얇은 선만. 그라데이션과 이모지는 쓰지 마.',
    '내용은 코드 맨 위의 const 배열 하나에 모아 두고 "여기를 고치면 내용이 바뀝니다" 주석을 달아 줘.', '코드는 ```html 블록 하나로, 설명은 세 줄 이하로.'].join('\n');
  return { apply: apply, has: has, fixPrompt: fixPrompt, errorPrompt: errorPrompt, FORMAT: FORMAT, SCRATCH: SCRATCH };
})();
if (typeof module === 'object' && module.exports) module.exports = VibePatch;
