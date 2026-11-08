/* GAS 작업실의 순수 함수 모음: 오류 문구 읽기, 원인 사전, 코드 점검, 위험 스캔, 줄 단위 비교, 프롬프트 만들기.
   브라우저(window.GasDiag)와 node 시험(gas/diag.test.mjs)에서 같은 파일을 씁니다. DOM이나 네트워크에 손대지 않습니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GasDiag = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- 오류 문구 읽기 ----------
  // Apps Script는 오류를 이렇게 알려 줍니다.
  //   TypeError: Cannot read properties of null (reading 'getRange') (줄 12, 파일 "코드")
  //   Exception: ... at 함수이름(코드:42:7)
  function parseError(text) {
    const raw = String(text || '').replace(/\r/g, '').trim();
    const lines = new Set();
    const add = (n) => { n = Number(n); if (n >= 1 && n <= 20000) lines.add(n); };
    for (const m of raw.matchAll(/(?:줄|line)\s*(\d+)/gi)) add(m[1]);
    for (const m of raw.matchAll(/\(([^()\s:]+):(\d+)(?::\d+)?\)/g)) add(m[2]);
    for (const m of raw.matchAll(/^\s*at\s+.*?:(\d+):\d+\)?\s*$/gm)) add(m[1]);
    const first = raw.split('\n').map((s) => s.trim()).find(Boolean) || '';
    const message = first.replace(/\s*\((?:줄|line)\s*\d+[^)]*\)\s*$/i, '').slice(0, 300);
    const fn = (raw.match(/\bat\s+([A-Za-z_$][\w$]*)\s*\(/) || [])[1] || '';
    return { message, lines: [...lines].sort((a, b) => a - b), fn, raw };
  }

  // 같은 오류인지 알아보는 표지: 숫자와 따옴표 안 내용을 지웁니다. 같은 표지가 두 번째면 더 강한 모델로 고칩니다.
  function signature(text) {
    return parseError(text).message.replace(/\(reading '[^']*'\)/g, '').replace(/'[^']*'|"[^"]*"/g, "''").replace(/\d+/g, '#').toLowerCase().slice(0, 120);
  }

  // ---------- 원인 사전 ----------
  // ai:true 이면 사전만으로는 부족하니 AI ONE에게 고치게 하라는 뜻입니다.
  const RULES = [
    { id: 'perm', re: /not authorized|authorization is required|permission|권한|insufficient authentication scopes|Required permissions/i,
      title: '권한 승인이 필요합니다', cause: '이 스크립트가 시트, 메일, 드라이브를 쓰려면 소유자가 한 번 승인해야 합니다.',
      fix: '편집기 위쪽 함수 선택 칸에서 아무 함수나 골라 ▶ 실행을 누르고, 뜨는 창에서 "권한 검토 → 계정 선택 → 고급 → 안전하지 않음으로 이동 → 허용"을 차례로 누르세요.', ai: false },
    { id: 'noSheet', re: /Cannot read propert(?:y|ies) of (?:null|undefined) \(reading '(getRange|getLastRow|appendRow|getLastColumn|getDataRange|setFrozenRows|getSheetId)'\)/i,
      title: '시트를 찾지 못했습니다', cause: '코드가 찾는 이름의 시트 탭이 없습니다. 탭 이름이 한 글자라도 다르면 못 찾습니다.',
      fix: '메뉴의 "처음 준비"를 먼저 실행하세요. 이미 했다면 시트 아래 탭 이름이 코드에 적힌 이름과 같은지(띄어쓰기까지) 확인하세요.', ai: false },
    { id: 'getUi', re: /Cannot call (?:SpreadsheetApp|FormApp|DocumentApp)\.getUi\(\) from this context/i,
      title: '화면 알림창을 못 띄우는 곳에서 실행됐습니다', cause: '웹 앱, 트리거, 폼 제출 같은 자동 실행에서는 alert 창을 열 수 없습니다.',
      fix: '그 함수 안의 SpreadsheetApp.getUi().alert(...) 줄을 지우거나 Logger.log(...)로 바꾸세요. AI ONE에게 "자동 실행에서도 돌게 getUi 호출을 빼 줘"라고 시켜도 됩니다.', ai: true },
    { id: 'time', re: /Exceeded maximum execution time/i,
      title: '실행 시간 6분을 넘었습니다', cause: 'Apps Script는 한 번 실행에 6분까지만 허용합니다. 처리할 줄이 너무 많거나 AI를 줄마다 부르면 넘습니다.',
      fix: '한 번에 처리하는 줄 수를 줄이세요(예: 한 번에 20줄만 처리하고 처리한 줄은 표시). 시간 트리거로 여러 번 나눠 돌리는 방법도 있습니다.', ai: true },
    { id: 'mailQuota', re: /(?:email|mail|메일).{0,40}(?:quota|limit|한도|too many)|too many times.{0,40}mail|Service invoked too many times for one day: email/i,
      title: '오늘 보낼 수 있는 메일을 다 썼습니다', cause: '무료 구글 계정은 하루 메일 약 100통, 학교(워크스페이스) 계정은 약 1,500통이 한도입니다.',
      fix: '내일 다시 실행하거나, 연습할 때는 받는 사람을 본인 한 명으로 줄이세요. 메일 대신 시트에 쌓아 두고 한 통으로 요약해 보내는 방법도 좋습니다.', ai: false },
    { id: 'fetchQuota', re: /Service invoked too many times for one day: urlfetch/i,
      title: '오늘 외부 호출 한도를 다 썼습니다', cause: 'UrlFetch(AI ONE과 Jev 호출 포함)는 하루 약 20,000회가 한도입니다.',
      fix: '반복 실행을 멈추고 내일 다시 하세요. 줄마다 호출하지 말고 여러 줄을 묶어 한 번에 보내도록 고치세요.', ai: true },
    { id: 'tooMany', re: /Service invoked too many times in a short time|too many times in a short time/i,
      title: '짧은 시간에 너무 많이 호출했습니다', cause: '시트나 외부 서비스를 연달아 부르면 잠깐 막힙니다.',
      fix: '1~2분 뒤 다시 실행하세요. 코드에서는 호출 사이에 Utilities.sleep(500)을 넣거나, 셀을 하나씩 읽지 말고 getValues()로 한 번에 읽으세요.', ai: true },
    { id: 'lock', re: /Could not obtain lock|Lock timeout/i,
      title: '다른 실행이 끝나기를 기다리다 포기했습니다', cause: '여러 명이 동시에 보내면 시트 쓰기를 한 명씩 차례로 합니다. 대기 시간이 모자랐습니다.',
      fix: '잠시 뒤 다시 보내면 됩니다. 사람이 아주 많다면 lock.waitLock(10000)의 숫자를 30000으로 늘리세요.', ai: false },
    { id: 'noFn', re: /Script function not found: (\w+)/i,
      title: '실행하려는 함수가 코드에 없습니다', cause: '메뉴나 트리거가 부르는 함수 이름이 코드에 없습니다. 붙여 넣다가 일부가 빠졌거나 이름이 다릅니다.',
      fix: '코드 전체를 지우고 공통 도우미와 레시피를 처음부터 다시 붙여 넣으세요. 웹 앱이라면 배포 → 배포 관리 → 새 버전으로 다시 배포하세요.', ai: false },
    { id: 'notDefinedHelper', re: /\b((?:post|ai|jev|ss|sheet|throttle|page)_|SERVER) is not defined/,
      title: '공통 도우미가 빠져 있습니다', cause: '레시피 코드는 공통 도우미(post_, ai_, jev_ ...)를 불러 씁니다. 도우미를 같이 붙여 넣지 않았습니다.',
      fix: '"공통 도우미" 코드를 레시피 코드 위에 한 번 붙여 넣으세요. 도우미는 한 프로젝트에 한 번만 둡니다.', ai: false },
    { id: 'notDefined', re: /ReferenceError: (\w+) is not defined/,
      title: '정의되지 않은 이름을 썼습니다', cause: '변수나 함수 이름의 오타이거나, 선언 없이 썼습니다. 대소문자도 구분합니다.',
      fix: '오류에 나온 줄에서 그 이름의 철자를 확인하세요. 다른 파일에 있던 함수라면 같은 프로젝트에 붙여 넣어야 합니다.', ai: true },
    { id: 'notFn', re: /TypeError: ([\w.$]+) is not a function/,
      title: '함수가 아닌 것을 함수처럼 불렀습니다', cause: '메서드 이름의 오타이거나, 그 객체에 없는 기능을 불렀습니다. (예: getRange를 getRnage로 씀)',
      fix: '오류에 나온 줄에서 점(.) 뒤 이름의 철자를 확인하세요. 구글 문서/시트/폼 서비스마다 쓸 수 있는 메서드가 다릅니다.', ai: true },
    { id: 'readNull', re: /Cannot read propert(?:y|ies) of (?:null|undefined)/i,
      title: '비어 있는 것의 속성을 읽으려 했습니다', cause: '앞 단계에서 찾은 시트, 파일, 폼, 값이 없는데(null) 그 안을 읽으려 했습니다.',
      fix: '오류에 나온 줄 바로 앞에서 무엇을 찾는지 보세요. 이름이나 주소가 틀렸거나 아직 만들어지지 않은 것입니다. 모르겠으면 아래 AI ONE 고치기를 누르세요.', ai: true },
    { id: 'range', re: /number of (?:rows|columns) in the range must be at least 1|coordinates or dimensions of the range are invalid|out of bounds/i,
      title: '시트 범위가 잘못됐습니다', cause: '읽을 줄이 하나도 없는데(데이터가 0줄) 범위를 만들었거나, 열 번호가 시트 크기를 넘었습니다.',
      fix: '시트에 자료가 있는지 확인하세요. 코드에서는 getLastRow()가 1 이하이면 먼저 return 하도록 막아야 합니다.', ai: true },
    { id: 'argType', re: /The parameters \((.*?)\) don't match the method signature for ([\w.]+)|Cannot find method (\w+)\(/i,
      title: '함수에 넣은 값의 종류가 맞지 않습니다', cause: '숫자를 넣을 자리에 글자를, 배열을 넣을 자리에 값 하나를 넣는 식입니다.',
      fix: '오류 문구의 괄호 안(넣은 값)과 메서드 이름을 보고 그 줄의 인자를 확인하세요. setValues는 항상 2차원 배열 [[...]]이어야 합니다.', ai: true },
    { id: 'noItem', re: /No item with the given ID|Document is missing|File not found|Unexpected error while getting the method or property open on object DriveApp|Invalid argument: id/i,
      title: '문서나 폴더를 못 찾았습니다', cause: '주소나 ID가 틀렸거나, 이 계정에 공유되지 않은 문서입니다.',
      fix: '문서 주소에서 /d/ 와 /edit 사이의 긴 글자가 ID입니다. 복사가 잘렸는지, 내 계정이 열 수 있는 문서인지 확인하세요.', ai: false },
    { id: 'jsonParse', re: /Unexpected token .* in JSON|JSON\.parse|Unexpected end of JSON/i,
      title: '서버 답이 JSON이 아닙니다', cause: '요청한 주소가 틀렸거나, 서버가 점검 중이라 웹페이지(오류 화면)를 돌려줬습니다.',
      fix: '맨 위 var SERVER = ... 주소가 https://kimcook-edu.pages.dev 인지 확인하고, 몇 분 뒤 다시 실행하세요.', ai: false },
    { id: 'syntax', re: /SyntaxError|Unexpected token|Unexpected identifier|missing \) after|Invalid or unexpected token|Unexpected end of input|Unexpected string|Unterminated/i,
      title: '문법 오류입니다 (괄호나 따옴표가 짝이 안 맞음)', cause: '붙여 넣다가 코드가 잘렸거나, 따옴표 · 괄호 · 중괄호 짝이 맞지 않습니다. AI 답이 길이 제한으로 끊겼을 수도 있습니다.',
      fix: '아래 "코드 점검"을 눌러 괄호 짝을 확인하세요. 코드 끝이 잘려 있다면 AI ONE에게 "끊긴 곳부터 이어서 완성해 줘"라고 시키세요.', ai: true },
    { id: 'classCodeMissing', re: /수업 코드가 아직 없습니다/,
      title: '수업 코드를 아직 저장하지 않았습니다', cause: '이 스크립트는 수업 코드가 있어야 AI ONE과 Jev를 쓸 수 있습니다.',
      fix: '시트로 돌아가 위쪽 메뉴의 "수업 코드 입력"을 눌러 강사가 알려 준 코드를 넣으세요.', ai: false },
    { id: 'classCodeBad', re: /AI ONE 401|수업 코드가 맞지 않습니다|잘못된 수업 코드/,
      title: '수업 코드가 맞지 않습니다', cause: '저장된 수업 코드가 틀렸거나 바뀌었습니다.',
      fix: '메뉴의 "수업 코드 입력"으로 다시 넣으세요. 대문자와 숫자를 잘 보세요.', ai: false },
    { id: 'aiRate', re: /AI ONE 429|호출 한도|사용하는 사람이 많|1분쯤 뒤|몰렸습니다/,
      title: 'AI 호출이 몰려 잠시 막혔습니다', cause: '수업 전체가 같이 쓰는 호출 한도에 닿았습니다.',
      fix: '1분쯤 뒤 다시 하세요. 줄마다 부르는 코드라면 여러 줄을 묶어 한 번만 부르게 고치세요.', ai: false },
    { id: 'aiDown', re: /AI ONE 503|서버가 바쁩니다|응답이 너무 늦습니다|AI ONE 연결 설정|크레딧이 부족/,
      title: 'AI ONE 서버 쪽 문제입니다', cause: '서버가 바쁘거나 점검 중이거나 크레딧이 모자랍니다. 코드 잘못이 아닙니다.',
      fix: '잠시 뒤 다시 실행하세요. 계속되면 강사에게 알려 주세요.', ai: false },
    { id: 'jevFormat', re: /Jev가 질문 형식|AI ONE 400|질문 «/,
      title: 'Jev에게 보낸 질문 모양이 틀렸습니다', cause: '질문은 type(noul, choice, score), instructions, criteria(선택지 또는 단계)를 정해진 모양으로 적어야 합니다.',
      fix: '오류 끝 괄호 안 설명을 읽어 보세요. choice의 criteria는 { 선택지: 설명 }, score의 criteria는 낮은 단계부터 적은 글 배열입니다. 아래 Jev 시험대에서 먼저 시험해 보세요.', ai: true },
    { id: 'trigger', re: /too many triggers|trigger.*limit|트리거.*(?:많|한도)/i,
      title: '트리거가 너무 많습니다', cause: '한 프로젝트에 만들 수 있는 트리거는 20개까지입니다. 같은 트리거를 계속 만들었을 수 있습니다.',
      fix: '편집기 왼쪽 시계 모양(트리거)에서 필요 없는 것을 지우세요. 코드에서는 만들기 전에 같은 이름 트리거를 지우도록 하세요.', ai: false },
    { id: 'deploy', re: /Cannot read propert(?:y|ies) of undefined \(reading 'parameter'\)|웹 앱.*(?:배포|deploy)/i,
      title: '웹 앱 배포나 doGet 실행 문제입니다', cause: 'doGet은 웹 주소로 열릴 때만 e(주소 정보)가 생깁니다. 편집기에서 ▶로 직접 실행하면 비어 있습니다.',
      fix: '편집기에서 doGet을 실행하지 말고, 배포 → 새 배포 → 웹 앱으로 만든 주소를 브라우저에서 여세요. 코드를 고친 뒤에는 배포 관리에서 새 버전으로 올려야 반영됩니다.', ai: false },
  ];

  function diagnose(text) {
    const t = String(text || '');
    const out = [];
    for (const r of RULES) {
      const m = r.re.exec(t);
      if (m) out.push({ id: r.id, title: r.title, cause: r.cause, fix: r.fix, ai: r.ai, hit: m[0].slice(0, 120) });
    }
    // 더 구체적인 풀이가 있으면 일반 풀이는 뺍니다.
    const has = (id) => out.some((d) => d.id === id);
    return out.filter((d) => !(d.id === 'readNull' && has('noSheet')) && !(d.id === 'notDefined' && has('notDefinedHelper')) && !(d.id === 'syntax' && has('jsonParse')));
  }

  // 오류에 나온 줄 주변 코드를 보여 주기 위한 조각.
  function excerpt(code, lines, context = 2) {
    const all = String(code || '').split('\n');
    const want = new Set();
    for (const n of lines) for (let i = n - context; i <= n + context; i++) if (i >= 1 && i <= all.length) want.add(i);
    const hit = new Set(lines);
    const nums = [...want].sort((a, b) => a - b);
    const out = [];
    let prev = 0;
    for (const n of nums) {
      if (prev && n > prev + 1) out.push({ n: 0, text: '…', hit: false, gap: true });
      out.push({ n, text: all[n - 1], hit: hit.has(n) });
      prev = n;
    }
    return out;
  }

  // ---------- 코드 점검 ----------
  // 문자열과 주석을 건너뛰며 괄호 짝을 셉니다. 정규식 리터럴은 따로 다루지 않으므로 힌트로만 씁니다.
  function balance(code) {
    const stack = [];
    const pairs = { ')': '(', ']': '[', '}': '{' };
    let i = 0, line = 1;
    const n = code.length;
    while (i < n) {
      const c = code[i], d = code[i + 1];
      if (c === '\n') { line++; i++; continue; }
      if (c === '/' && d === '/') { while (i < n && code[i] !== '\n') i++; continue; }
      if (c === '/' && d === '*') { i += 2; while (i < n && !(code[i] === '*' && code[i + 1] === '/')) { if (code[i] === '\n') line++; i++; } i += 2; continue; }
      if (c === '"' || c === "'" || c === '`') {
        const q = c, startLine = line;
        i++;
        while (i < n && code[i] !== q) {
          if (code[i] === '\\') i++;
          else if (code[i] === '\n') { line++; if (q !== '`') return { ok: false, line: startLine, msg: `따옴표(${q})가 닫히지 않았습니다.` }; }
          i++;
        }
        if (i >= n) return { ok: false, line: startLine, msg: `따옴표(${q})가 닫히지 않았습니다.` };
        i++;
        continue;
      }
      if (c === '(' || c === '[' || c === '{') stack.push({ c, line });
      else if (pairs[c]) {
        const top = stack.pop();
        if (!top || top.c !== pairs[c]) return { ok: false, line, msg: `닫는 괄호 ${c} 짝이 맞지 않습니다.` };
      }
      i++;
    }
    if (stack.length) {
      const top = stack[stack.length - 1];
      return { ok: false, line: top.line, msg: `${top.c} 가 닫히지 않았습니다 (닫는 괄호가 ${stack.length}개 부족합니다). 코드가 중간에 잘렸을 수 있습니다.` };
    }
    return { ok: true };
  }

  const HELPERS = ['post_', 'ai_', 'jev_', 'ss_', 'sheet_', 'throttle_', 'page_', 'webUrl_', 'execUrl_'];

  function definedFunctions(code) {
    const names = [];
    for (const m of code.matchAll(/^[ \t]*function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) names.push({ name: m[1], line: code.slice(0, m.index).split('\n').length });
    return names;
  }

  function staticCheck(code) {
    const src = String(code || '');
    const out = [];
    if (!src.trim()) return [{ level: 'error', line: 0, msg: '코드가 비어 있습니다.' }];
    let syntaxOk = true;
    try { new Function(src); } catch (e) { syntaxOk = false; out.push({ level: 'error', line: 0, msg: '문법 오류: ' + String(e.message).slice(0, 140) }); }
    const bal = balance(src);
    if (!bal.ok) out.push({ level: syntaxOk ? 'warn' : 'error', line: bal.line, msg: bal.msg });
    const defined = new Set(definedFunctions(src).map((d) => d.name));
    for (const h of HELPERS) {
      if (new RegExp('\\b' + h + '\\(').test(src) && !defined.has(h)) out.push({ level: 'error', line: 0, msg: `${h} 함수를 쓰는데 정의가 없습니다. 공통 도우미 코드를 같이 붙여 넣으세요.` });
    }
    if (/\bSERVER\b/.test(src) && !/\b(?:var|let|const)\s+SERVER\b/.test(src)) out.push({ level: 'error', line: 0, msg: 'SERVER 값이 정의되어 있지 않습니다. 공통 도우미를 같이 붙여 넣으세요.' });
    const seen = new Map();
    for (const d of definedFunctions(src)) {
      if (seen.has(d.name)) out.push({ level: 'warn', line: d.line, msg: `함수 ${d.name}이 ${seen.get(d.name)}줄에도 있습니다. Apps Script는 뒤의 것만 실행합니다. 레시피를 두 개 붙였다면 프로젝트(파일)를 나누세요.` });
      else seen.set(d.name, d.line);
    }
    const lines = src.split('\n');
    lines.forEach((ln, i) => {
      if (/AIza[0-9A-Za-z_-]{30,}|\bsk-[A-Za-z0-9_-]{20,}|Bearer\s+[A-Za-z0-9._-]{30,}/.test(ln)) {
        out.push({ level: 'error', line: i + 1, msg: 'API 키처럼 보이는 글자가 코드에 직접 들어 있습니다. 키는 스크립트 속성에만 두세요(이 도구의 코드는 수업 코드만 속성에 저장합니다).' });
      }
      if (/getService\(\)\.getUrl\(/.test(ln.replace(/(^|[^:])\/\/.*$/, '$1'))) {
        out.push({ level: 'warn', line: i + 1, msg: '학생에게 줄 주소에 ScriptApp.getService().getUrl()을 쓰면 /dev나 없는 주소가 나와 학생 폰에 "현재 파일을 열 수 없습니다"가 뜹니다. webUrl_()을 쓰세요(시트 메뉴 웹 앱 주소 넣기로 넣은 /exec 주소).' });
      }
    });
    return out;
  }

  // 실행 전 안전 점검: 이 코드가 밖으로 무엇을 보내고 무엇을 지우거나 덮어쓰는지.
  const RISKS = [
    { key: 'mail', label: '메일을 보냅니다', re: /\b(?:MailApp|GmailApp)\.(?:sendEmail|createDraft)\b/g, level: 'warn' },
    { key: 'delete', label: '지우거나 휴지통으로 보냅니다', re: /\.(?:deleteRows?|deleteColumns?|deleteSheet|setTrashed|removeFile|deleteAllTriggers)\b|\.moveToTrash\b/g, level: 'danger' },
    { key: 'clear', label: '내용을 비웁니다/덮어씁니다', re: /\.(?:clear|clearContent|clearContents|clearFormat)\(|\.getBody\(\)\.clear\(/g, level: 'danger' },
    { key: 'share', label: '공유 설정을 바꿉니다', re: /\.(?:setSharing|addEditors?|addViewers?|addCommenters?|removeEditor)\b/g, level: 'danger' },
    { key: 'trigger', label: '자동 실행(트리거)을 만듭니다', re: /ScriptApp\.newTrigger\b/g, level: 'info' },
    { key: 'create', label: '문서, 폴더, 폼을 새로 만듭니다', re: /\b(?:DocumentApp|DriveApp|FormApp|SlidesApp)\.create\w*\(|\.insertSheet\(/g, level: 'info' },
    { key: 'fetch', label: '인터넷 주소로 데이터를 보냅니다', re: /UrlFetchApp\.fetch\b|\bpost_\(|\bai_\(|\bjev_\(/g, level: 'warn' },
    { key: 'calendar', label: '캘린더 일정을 만들거나 바꿉니다', re: /CalendarApp\.[\w.]*(?:create\w+|deleteEvent)\b/g, level: 'warn' },
  ];

  function riskScan(code) {
    const src = String(code || '');
    const out = [];
    for (const r of RISKS) {
      const m = src.match(r.re);
      if (m) out.push({ key: r.key, label: r.label, count: m.length, level: r.level });
    }
    const hosts = new Set();
    for (const m of src.matchAll(/https?:\/\/([A-Za-z0-9.-]+)/g)) hosts.add(m[1].toLowerCase());
    const scopes = [];
    if (/SpreadsheetApp/.test(src)) scopes.push('스프레드시트');
    if (/DriveApp/.test(src)) scopes.push('드라이브 파일');
    if (/DocumentApp/.test(src)) scopes.push('구글 문서');
    if (/FormApp/.test(src)) scopes.push('구글 폼');
    if (/MailApp|GmailApp/.test(src)) scopes.push('메일 보내기');
    if (/CalendarApp/.test(src)) scopes.push('캘린더');
    if (/UrlFetchApp|\b(?:post_|ai_|jev_)\(/.test(src)) scopes.push('외부 서버 연결');
    return { risks: out, hosts: [...hosts], scopes };
  }

  // ---------- 줄 단위 비교 ----------
  function lineDiff(a, b) {
    const A = String(a || '').split('\n'), B = String(b || '').split('\n');
    let s = 0;
    while (s < A.length && s < B.length && A[s] === B[s]) s++;
    let ea = A.length, eb = B.length;
    while (ea > s && eb > s && A[ea - 1] === B[eb - 1]) { ea--; eb--; }
    const a2 = A.slice(s, ea), b2 = B.slice(s, eb);
    if (a2.length * b2.length > 4_000_000) return { added: b2.length, removed: a2.length, hunks: [{ aStart: s + 1, bStart: s + 1, removed: a2.slice(0, 40), added: b2.slice(0, 40) }] };
    const n = a2.length, m = b2.length, w = m + 1;
    const L = new Uint32Array((n + 1) * w);
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i * w + j] = a2[i] === b2[j] ? L[(i + 1) * w + j + 1] + 1 : Math.max(L[(i + 1) * w + j], L[i * w + j + 1]);
    const hunks = [];
    let i = 0, j = 0, cur = null, added = 0, removed = 0;
    const flush = () => { if (cur) { hunks.push(cur); cur = null; } };
    const open = () => { if (!cur) cur = { aStart: s + i + 1, bStart: s + j + 1, removed: [], added: [] }; return cur; };
    while (i < n || j < m) {
      if (i < n && j < m && a2[i] === b2[j]) { flush(); i++; j++; }
      else if (j < m && (i >= n || L[i * w + j + 1] >= L[(i + 1) * w + j])) { open().added.push(b2[j]); added++; j++; }
      else { open().removed.push(a2[i]); removed++; i++; }
    }
    flush();
    return { added, removed, hunks };
  }

  // ---------- 바뀐 곳만 받아 기준 코드에 끼워 넣기 ----------
  // AI는 고친 함수(와 새 함수)나 바꿀 줄만 쓰고, 여기서 기준 코드에 끼워 넣습니다. 코드 전체를 다시 쓰게 하면 긴 코드에서 답이 끊기기 때문입니다.
  // 원칙: 답에서 하나라도 안전하게 못 끼우면(failed) 코드를 하나도 바꾸지 않습니다.
  const FN_HEAD = /(?:async\s+)?function(?:\s*\*\s*|\s+)([A-Za-z_$][\w$]*)\s*\(/y;
  const VAR_HEAD = /(?:const|let|var)\s+([A-Za-z_$][\w$]*)\b/y;
  const NEXT_DECL = /^[ \t]*(?:\/\/|\/\*|(?:async\s+)?function\b|const\b|let\b|var\b)/;
  const REGEX_AFTER = /[(,=:[!&|?{};+\-*%~^]$/;
  const REGEX_WORD = /(?:^|[^\w$])(?:return|typeof|case|in|of|void|delete)$/;
  const EXPR_END = /[\w$)\]}'"`/]/;

  // ' 나 " 로 시작하는 글자 덩어리의 끝(한 줄 안). 닫히지 않으면 줄 끝에서 멈춥니다.
  function skipQuoted(s, i) {
    const q = s[i], n = s.length;
    i++;
    while (i < n) {
      if (s[i] === '\\') i += 2;
      else if (s[i] === q) return i + 1;
      else if (s[i] === '\n') return i;
      else i++;
    }
    return n;
  }

  // ` 로 시작하는 템플릿 글자의 끝. ${ … } 안의 따옴표, 템플릿, 주석도 따라갑니다.
  function skipTemplate(s, i) {
    const n = s.length;
    let j = i + 1;
    while (j < n) {
      const c = s[j];
      if (c === '\\') j += 2;
      else if (c === '`') return j + 1;
      else if (c === '$' && s[j + 1] === '{') j = skipBraces(s, j + 1);
      else j++;
    }
    return n;
  }

  function skipBraces(s, i) {
    const n = s.length;
    let depth = 0, j = i;
    while (j < n) {
      const c = s[j], d = s[j + 1];
      if (c === '"' || c === "'") j = skipQuoted(s, j);
      else if (c === '`') j = skipTemplate(s, j);
      else if (c === '/' && d === '/') { while (j < n && s[j] !== '\n') j++; }
      else if (c === '/' && d === '*') { const e = s.indexOf('*/', j + 2); j = e < 0 ? n : e + 2; }
      else if (c === '{') { depth++; j++; }
      else if (c === '}') { depth--; j++; if (depth === 0) return j; }
      else j++;
    }
    return n;
  }

  function skipRegex(s, i) {
    const n = s.length;
    let j = i + 1, inClass = false;
    while (j < n && s[j] !== '\n') {
      const c = s[j];
      if (c === '\\') { j += 2; continue; }
      if (c === '[') inClass = true;
      else if (c === ']') inClass = false;
      else if (c === '/' && !inClass) { j++; while (j < n && /[a-z]/i.test(s[j])) j++; return j; }
      j++;
    }
    return -1;
  }

  // 맨 바깥의 함수 선언과 const/let/var 선언을 찾아 [start, end) 범위를 돌려줍니다. 끝이 없는 선언은 incomplete 입니다.
  function scanUnits(code) {
    const s = String(code || ''), n = s.length, units = [];
    let i = 0, depth = 0, cur = null, last = '';
    const atLineStart = (k) => { for (let j = k - 1; j >= 0 && s[j] !== '\n'; j--) if (s[j] !== ' ' && s[j] !== '\t') return false; return true; };
    const finish = (end, incomplete) => { units.push({ kind: cur.kind, name: cur.name, start: cur.start, end, incomplete: !!incomplete }); cur = null; };
    while (i < n) {
      const c = s[i], d = s[i + 1];
      if (c === '/' && d === '/') { while (i < n && s[i] !== '\n') i++; continue; }
      if (c === '/' && d === '*') { const e = s.indexOf('*/', i + 2); i = e < 0 ? n : e + 2; continue; }
      if (c === '"' || c === "'") { i = skipQuoted(s, i); last = c; continue; }
      if (c === '`') { i = skipTemplate(s, i); last = c; continue; }
      if (c === '/' && (last === '' || REGEX_AFTER.test(last) || REGEX_WORD.test(s.slice(Math.max(0, i - 8), i).trimEnd()))) {
        const e = skipRegex(s, i);
        if (e > 0) { i = e; last = '/'; continue; }
      }
      if (!cur && depth === 0 && atLineStart(i)) {
        FN_HEAD.lastIndex = i; VAR_HEAD.lastIndex = i;
        let m;
        if ((m = FN_HEAD.exec(s))) { cur = { kind: 'function', name: m[1], start: i, body: false }; i += m[0].length - 1; last = 'n'; continue; }
        if ((m = VAR_HEAD.exec(s))) { cur = { kind: 'var', name: m[1], start: i }; i += m[0].length; last = 'r'; continue; }
      }
      if (c === '(' || c === '[' || c === '{') { depth++; if (cur && cur.kind === 'function' && c === '{' && depth === 1) cur.body = true; }
      else if (c === ')' || c === ']' || c === '}') {
        if (depth > 0) depth--;
        if (cur && cur.kind === 'function' && c === '}' && depth === 0 && cur.body) { i++; last = c; finish(i, false); continue; }
      } else if (c === ';' && depth === 0 && cur && cur.kind === 'var') { i++; last = c; finish(i, false); continue; }
      else if (c === '\n' && depth === 0 && cur && cur.kind === 'var' && EXPR_END.test(last) && NEXT_DECL.test(s.slice(i + 1, i + 120))) { finish(i, false); i++; continue; }
      if (!/\s/.test(c)) last = c;
      i++;
    }
    if (cur) {
      if (cur.kind === 'var' && depth === 0) finish(n, false);
      else finish(n, true);
    }
    return units;
  }

  const DELETE_LINE = /^[ \t]*\/\/\s*(?:삭제|제거|delete|remove)\s*[:：]\s*([A-Za-z_$][\w$]*)/gim;

  // 선언 바로 위에 붙은 주석(// 줄과 /* … */ 덩어리). 삭제 표시 줄은 버리고, 주석 시작을 못 찾으면 모두 버립니다.
  function leadingComments(src, start) {
    const lines = src.slice(0, start).split('\n');
    lines.pop();
    let keep = [];
    for (;;) {
      const t = (lines[lines.length - 1] || '').trim();
      if (t.startsWith('//')) { keep.unshift(lines.pop()); continue; }
      if (t.endsWith('*/')) {
        let k = lines.length - 1;
        while (k >= 0 && !lines[k].includes('/*')) k--;
        if (k < 0) return '';
        keep = lines.splice(k).concat(keep);
        continue;
      }
      break;
    }
    const text = keep.filter((l) => !new RegExp(DELETE_LINE.source, 'i').test(l)).join('\n');
    return text ? text + '\n' : '';
  }

  // 큰 함수 안의 몇 줄만 바꿀 때: <<<<<<< 찾기 / ======= / >>>>>>> 바꾸기 블록. 찾을 글이 기준 코드에 한 번만 있어야 바꿉니다(들여쓰기가 달라도 줄 내용이 같으면 찾습니다).
  const EDIT_BLOCK = /^[ \t]*<<<<<<<[^\n]*\n([\s\S]*?)\n[ \t]*=======[ \t]*\n([\s\S]*?)\n?[ \t]*>>>>>>>[^\n]*$/gm;
  const MARKER = /^[ \t]*(?:<<<<<<<|=======[ \t]*$|>>>>>>>)/m;

  function findAll(hay, needle) {
    const at = [];
    if (!needle) return at;
    for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + 1)) at.push(i);
    return at;
  }

  function locate(code, old) {
    const exact = findAll(code, old);
    if (exact.length) return exact.length === 1 ? { start: exact[0], end: exact[0] + old.length } : { many: true };
    const want = old.split('\n').map((l) => l.trim());
    while (want.length && want[0] === '') want.shift();
    while (want.length && want[want.length - 1] === '') want.pop();
    if (!want.length) return null;
    const lines = code.split('\n'), offs = [];
    let o = 0;
    for (const l of lines) { offs.push(o); o += l.length + 1; }
    const hits = [];
    for (let i = 0; i + want.length <= lines.length; i++) {
      let ok = true;
      for (let k = 0; k < want.length && ok; k++) if (lines[i + k].trim() !== want[k]) ok = false;
      if (ok) hits.push(i);
    }
    if (hits.length !== 1) return hits.length ? { many: true } : null;
    const a = hits[0], b = a + want.length - 1;
    return { start: offs[a], end: offs[b] + lines[b].length, loose: true, indent: lines[a].match(/^[ \t]*/)[0] };
  }

  // 찾기/바꾸기 블록을 차례로 적용합니다. rest 는 블록을 뺀 나머지 답입니다.
  function applyEdits(code, patch) {
    const out = { code: String(code || ''), applied: 0, failed: [], rest: '', open: false };
    let rest = String(patch || '');
    const blocks = [...rest.matchAll(EDIT_BLOCK)];
    rest = rest.replace(EDIT_BLOCK, () => '');
    if (MARKER.test(rest)) out.open = true;
    for (const b of blocks) {
      const old = b[1], neu = b[2];
      const label = (old.split('\n').find((l) => l.trim()) || '').trim().slice(0, 60);
      if (MARKER.test(old) || MARKER.test(neu)) { out.failed.push({ label, why: '바꿀 곳 표시가 서로 엉켜 있습니다' }); continue; }
      const at = locate(out.code, old);
      if (!at || at.many) { out.failed.push({ label, why: at ? '같은 글이 여러 곳에 있어 어디인지 알 수 없습니다' : '기준 코드에서 그 글을 찾지 못했습니다' }); continue; }
      let text = neu;
      if (at.loose && at.indent) {
        const first = text.split('\n').find((l) => l.trim()) || '';
        if (!/^[ \t]/.test(first)) text = text.split('\n').map((l) => (l.trim() ? at.indent + l : l)).join('\n');
      }
      out.code = out.code.slice(0, at.start) + text + out.code.slice(at.end);
      out.applied++;
    }
    out.rest = rest;
    return out;
  }

  const escapeRe = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  // base 에 patch(함수, 선언, 찾기/바꾸기 블록, 삭제 줄)를 끼워 넣습니다. failed 가 하나라도 있으면 code 는 base 그대로입니다.
  function applyPatch(base, patch) {
    const ed = applyEdits(base, patch);
    const b = ed.code, p = ed.rest;
    const out = { code: String(base || ''), replaced: [], added: [], removed: [], edited: ed.applied, failed: ed.failed.slice(), incomplete: false, count: ed.applied + ed.failed.length };
    if (ed.open) out.failed.push({ label: '<<<<<<<', why: '바꿀 곳 표시가 끝나지 않았거나 형식이 깨졌습니다' });
    const pu = scanUnits(p);
    out.incomplete = pu.some((u) => u.incomplete);
    out.count += pu.length;
    // 단위 밖에 남는 글: 주석과 삭제 줄만 허용합니다. 그 밖의 글은 반영할 수 없으니 거절합니다.
    let outside = '', pos = 0;
    for (const u of pu) { outside += p.slice(pos, u.start) + '\n'; pos = u.end; }
    outside += p.slice(pos);
    const deletes = [...outside.matchAll(DELETE_LINE)].map((m) => m[1]);
    const stray = outside.replace(DELETE_LINE, '').replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '').trim();
    if (stray) out.failed.push({ label: stray.split('\n')[0].trim().slice(0, 60), why: '함수나 바꿀 곳 형식이 아닌 글이 섞여 있습니다' });
    if (out.incomplete || out.failed.length) return out;
    const bu = scanUnits(b);
    if ((pu.length || deletes.length) && bu.some((u) => u.incomplete)) {
      out.failed.push({ label: '기준 코드', why: '기준 코드의 괄호가 맞지 않아 함수를 안전하게 끼워 넣을 수 없습니다. 먼저 코드 점검으로 괄호를 고쳐 주세요' });
      return out;
    }
    const byName = new Map();
    for (const u of pu) byName.set(u.name, u);
    const edits = [], appended = [], fresh = [], finalText = new Map();
    for (const u of byName.values()) {
      const text = p.slice(u.start, u.end);
      let hit = null;
      for (const x of bu) if (x.name === u.name) hit = x;
      if (hit) { edits.push({ start: hit.start, end: hit.end, text }); out.replaced.push(u.name); finalText.set(hit.start, text); }
      else fresh.push({ u, text, lead: leadingComments(p, u.start) });
    }
    for (const { u, text, lead } of fresh) {
      out.added.push(u.name);
      if (u.kind === 'var') {
        // 새 상수는 그것을 쓰는 첫 선언(바뀐 선언 포함) 앞에 둡니다. 맨 끝에 두면 먼저 실행되는 선언이 아직 없는 값을 읽습니다.
        const re = new RegExp('(^|[^\\w$.])' + escapeRe(u.name) + '(?![\\w$])');
        const user = bu.find((x) => re.test(finalText.get(x.start) ?? b.slice(x.start, x.end)));
        if (user) { edits.push({ start: user.start, end: user.start, text: lead + text + '\n\n' }); continue; }
      }
      appended.push(lead + text);
    }
    for (const name of deletes) {
      let hit = null;
      for (const x of bu) if (x.name === name) hit = x;
      if (!hit) { out.failed.push({ label: name, why: '지우라고 한 이름이 기준 코드에 없습니다' }); continue; }
      if (edits.some((e) => e.start === hit.start && e.end === hit.end)) continue;
      let end = hit.end;
      while (b[end] === '\n' && b[end + 1] === '\n') end++;
      if (b[end] === '\n') end++;
      edits.push({ start: hit.start, end, text: '' });
      out.removed.push(name);
    }
    if (out.failed.length) return out;
    let code = b;
    for (const e of edits.map((x, i) => ({ ...x, i })).sort((x, y) => y.start - x.start || y.i - x.i)) code = code.slice(0, e.start) + e.text + code.slice(e.end);
    if (appended.length) code = code.replace(/\s+$/, '') + '\n\n' + appended.join('\n\n') + '\n';
    out.code = code;
    return out;
  }

  // 답 안의 코드 블록을 모두 이어 붙입니다(함수를 여러 블록에 나눠 줘도 받도록). 바꿀 곳 표시가 든 블록은 언어 표시와 상관없이 받습니다. 닫히지 않은 블록이 있으면 truncated 입니다.
  function extractPatch(markdown) {
    const s = String(markdown || '');
    const blocks = [];
    const re = /```[ \t]*([\w+-]*)[ \t]*\n([\s\S]*?)(```|$)/g;
    let m;
    while ((m = re.exec(s))) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      blocks.push({ lang: m[1].toLowerCase(), code: m[2].replace(/\s+$/, ''), closed: m[3] === '```' });
    }
    const wanted = blocks.filter((b) => b.code.trim() && (/^(javascript|js|gs|apps-?script|)$/.test(b.lang) || MARKER.test(b.code)));
    if (wanted.length) return { code: wanted.map((b) => b.code).join('\n\n'), truncated: wanted.some((b) => !b.closed) };
    if (/\bfunction\s+\w+\s*\(/.test(s) && !/```/.test(s)) return { code: s.trim(), truncated: false };
    return { code: '', truncated: false };
  }

  // ---------- 프롬프트 ----------
  // 밖의 AI(ChatGPT 등)에 붙여 넣을 때는 사람이 직접 붙이므로 전체 코드를 받습니다.
  const MAKE_RULES = [
    '규칙:',
    '- 첨부한 "기준 코드"의 구조, 함수 이름, 도우미 함수(post_, ai_, jev_, throttle_, sheet_, page_)를 그대로 쓰고, 아래 요청에 필요한 부분만 바꾸거나 더합니다.',
    '- 수정한 전체 코드를 ```javascript 코드 블록 하나로 줍니다. 일부만 보여 주지 않습니다.',
    '- 수업 코드나 API 키를 코드에 쓰지 않습니다.',
    '- 학생에게 줄 웹 앱 주소는 webUrl_()(시트 메뉴 "웹 앱 주소 넣기"로 넣은 /exec 주소)만 씁니다. ScriptApp.getService().getUrl()은 /dev나 없는 주소를 돌려줄 때가 있어 쓰지 않습니다.',
    '- 학생 이름, 학번 같은 개인정보를 새로 모으지 않습니다.',
    '- 코드 블록 밖 설명은 다섯 줄 이내로, 무엇을 바꿨는지만 씁니다.',
  ].join('\n');

  // AI ONE 에는 바뀐 함수만 받습니다. 화면이 같은 이름의 함수를 찾아 끼워 넣습니다(applyPatch).
  const PATCH_RULES = [
    '- 바꾸거나 새로 만든 함수만 ```javascript 코드 블록 하나에 줍니다. 바꾸지 않은 함수는 쓰지 않습니다. 화면이 기준 코드에서 같은 이름의 함수를 찾아 자동으로 바꿔 끼웁니다.',
    '- 바꾼 함수는 이름을 기준 코드와 똑같이 쓰고, 함수 하나를 처음부터 끝까지 빠짐없이 씁니다. 함수 안을 "..."나 "기존과 같음"으로 줄이지 않습니다.',
    '- 맨 위의 const, var 선언을 바꿀 때는 그 선언 전체를 씁니다. 새 함수는 같은 코드 블록에 이어서 씁니다.',
    '- 지울 함수는 코드 블록 안에 `// 삭제: 함수이름` 한 줄로 적습니다.',
    '- 큰 함수(boardClient_처럼 수십 줄 이상) 안의 몇 줄만 바꿀 때는 함수 전체를 다시 쓰지 말고 아래 형식으로 바꿀 곳만 씁니다. 표시줄(<<<<<<<, =======, >>>>>>>)은 들여쓰기 없이 줄 맨 앞에 쓰고, 같은 코드 블록에 여러 개 쓸 수 있습니다.',
    '<<<<<<< 찾기',
    '(기준 코드에 있는 그대로의 줄 2~6줄. 기준 코드에 한 번만 나오는 곳이어야 합니다)',
    '=======',
    '(그 자리에 들어갈 새 줄)',
    '>>>>>>> 바꾸기',
    '- 코드 블록 안에는 함수, 맨 위 const/var 선언, 위 형식의 바꿀 곳, 삭제 줄, 주석만 씁니다. 그 밖의 문장이나 문장 하나만 따로 쓰지 않습니다.',
  ];
  const MAKE_RULES_AI = MAKE_RULES.split('\n').filter((l) => !l.includes('수정한 전체 코드')).concat(PATCH_RULES).join('\n');

  function buildMakePrompt({ request, title, hasBase }) {
    return [
      `Google Apps Script 코드를 ${hasBase ? '고쳐' : '새로'} 만들어 주세요. 레시피: ${title || '새 자동화'}`,
      '',
      '[요청]',
      String(request || '').trim(),
      '',
      MAKE_RULES_AI,
    ].join('\n');
  }

  // ChatGPT, Gemini 등 밖의 AI에 붙여 넣을 때: 코드를 프롬프트 안에 함께 넣습니다.
  function buildExternalPrompt({ request, title, code }) {
    return [
      'Google Apps Script(V8) 전문가로서 아래 코드를 요청대로 고쳐 주세요.',
      `레시피: ${title || '새 자동화'}`,
      '',
      '[요청]',
      String(request || '').trim(),
      '',
      MAKE_RULES.replace('첨부한 "기준 코드"', '아래 "기준 코드"'),
      '',
      '[기준 코드]',
      '```javascript',
      String(code || '').trim(),
      '```',
    ].join('\n');
  }

  function buildFixPrompt({ error, notes, lines, attempt, code }) {
    const p = parseError(error);
    const hint = (notes || []).map((d) => `- ${d.title}: ${d.cause}`).join('\n');
    const where = (lines && lines.length ? lines : p.lines).join(', ');
    return [
      '첨부한 "현재 코드"를 실행했더니 아래 오류가 났습니다. 원인을 찾아 고친 함수를 주세요.',
      attempt > 0 ? '※ 이미 한 번 고쳤는데 같은 종류의 오류가 또 납니다. 앞의 수정이 틀렸을 수 있으니 가정부터 다시 확인하세요.' : '',
      '',
      '[오류 문구]',
      String(error || '').trim().slice(0, 2500),
      where ? `[오류가 난 줄] ${where}` : '',
      hint ? `[도구가 추정한 원인]\n${hint}` : '',
      '',
      '규칙:',
      '- 오류가 난 줄 주변을 먼저 살피고, 필요한 곳만 최소한으로 고칩니다. 다른 기능은 바꾸지 않습니다.',
      ...PATCH_RULES.map((l) => l.replace('바꾸거나 새로 만든 함수만', '고친 함수만')),
      '- 코드 블록 밖에는 "원인 한 줄, 고친 곳 한 줄"만 씁니다.',
      '- 사용자가 시트를 직접 만들거나 이름을 맞춰야 하는 문제라면 코드를 고치지 말고 그 방법만 알려 줍니다.',
    ].filter((x) => x !== '').join('\n');
  }

  return { parseError, signature, diagnose, excerpt, balance, staticCheck, riskScan, lineDiff, scanUnits, applyEdits, applyPatch, extractPatch, buildMakePrompt, buildExternalPrompt, buildFixPrompt, RULES, HELPERS };
});
