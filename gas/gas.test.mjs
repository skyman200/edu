// Runs every recipe under node against an in-memory Apps Script (mock-gas.mjs): `node gas/gas.test.mjs`
// The fake server uses the real /api/jev and /api/ai validators, so a request the classroom server would reject fails here.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createEnv, readGs as readGs0, choice, noul, score, defaultAnswers, runPage } from './mock-gas.mjs';
import { soloCommon, soloRecipe } from './build.mjs';

// SOLO=1 node gas/gas.test.mjs runs the same recipes as the solo build (Gemini key, no class server); tests about the class server are skipped there.
const SOLO = !!process.env.SOLO;
const readGs = (n) => (SOLO && n === '00-common.gs' ? soloCommon(readGs0(n), readGs0('solo.gs')) : SOLO ? soloRecipe(readGs0(n)) : readGs0(n));

const tests = [];
const test = (name, fn) => { if (!(SOLO && /^server:/.test(name))) tests.push([name, fn]); };
const COMMON = readGs('00-common.gs');

function boot(recipe, opts) {
  const env = createEnv(opts);
  env.run(COMMON + '\n' + readGs(recipe), recipe);
  // results made inside the vm context have another Object prototype: clone plain data so strict deepEqual works
  env.call = (fn, ...args) => {
    const r = env.ctx[fn](...args);
    return r && typeof r === 'object' && !('html' in r) ? JSON.parse(JSON.stringify(r)) : r;
  };
  env.sheet = (n) => env.active.getSheetByName(n);
  env.rows = (n) => env.sheet(n).getRange(1, 1, Math.max(1, env.sheet(n).getLastRow()), 12).getValues().slice(1).map((r) => r.slice(0));
  env.setup = () => { env.call('onOpen'); env.call('setup'); };
  return env;
}

test('server: common: setClassCode checks the connection; post_ retries 503 then fails with the server message', () => {
  const env = boot('r1-board.gs');
  env.ctx.PropertiesService.getScriptProperties().setProperty('CLASS_CODE', 'old');
  env.call('setClassCode');
  assert.equal(env.ctx.PropertiesService.getScriptProperties().getProperty('CLASS_CODE'), 'TESTCODE', 'the code is trimmed');
  assert.ok(env.log.alerts.at(-1).includes('연결도 확인'));
  const bad = boot('r1-board.gs');
  bad.ctx.PropertiesService.getScriptProperties().setProperty('CLASS_CODE', 'WRONG');
  assert.throws(() => bad.call('jev_', 'x', { a: { type: 'noul', instructions: 'q' } }), /401.*수업 코드가 맞지 않습니다/);
  const none = boot('r1-board.gs');
  none.ctx.PropertiesService.getScriptProperties().setProperty('CLASS_CODE', '');
  assert.throws(() => none.call('ai_', 'x'), /수업 코드가 아직 없습니다/);
});

test('server: common: ai_ joins the stream, sends files/images, surfaces stream errors', () => {
  const env = boot('r1-board.gs', { ai: ({ files, images }) => `답:${files.length}:${images.length}` });
  assert.equal(env.call('ai_', 'hi', { files: [{ name: 'f', text: 't' }] }), '답:1:0');
  assert.equal(env.call('ai_', [{ role: 'user', content: 'a' }, { role: 'assistant', content: 'b' }, { role: 'user', content: 'c' }]), '답:0:0');
  assert.equal(env.call('ai_', 'pic', { images: [{ name: 'p', url: 'data:image/jpeg;base64,/9j/AAAA' }] }), '답:0:1');
  const err = boot('r1-board.gs');
  err.ctx.UrlFetchApp.fetch = () => ({ getResponseCode: () => 200, getContentText: () => 'data: {"delta":"앞"}\n\ndata: {"error":"끊겼습니다"}\n\n' });
  assert.throws(() => err.call('ai_', 'x'), /끊겼습니다/);
});

test('common: throttle_ allows max then blocks', () => {
  const env = boot('r1-board.gs');
  assert.deepEqual([1, 2, 3, 4].map(() => env.call('throttle_', 'k', 3, 60)), [true, true, true, false]);
});

// ---------- recipe 1 ----------
test('r1: setup, pages, submit grouped by Jev, moderation, throttle, Jev outage', () => {
  const env = boot('r1-board.gs');
  env.setup();
  assert.equal(env.log.menus[0].name, '반응판');
  for (const [, fn] of env.log.menus[0].items) assert.equal(typeof env.ctx[fn], 'function', fn);
  assert.ok(env.sheet('설정') && env.sheet('응답'));
  const studentHtml = env.call('doGet', { parameter: {} }).html;
  assert.ok(studentHtml.includes('submitAnswer'));
  assert.match(studentHtml, /id="ai-notice"[^>]*>[^<]*<\/p><textarea[^>]*aria-describedby="ai-notice"/);
  const boardHtml = env.call('doGet', { parameter: { v: 'board' } }).html;
  assert.ok(boardHtml.includes('getBoard'));
  assert.ok(boardHtml.includes('응답을 한꺼번에 보내 새 묶음을 만드는 기능을 켜시겠습니까?'));
  const batch = '새 묶음을 만들기 위해 질문·기존 묶음 이름과 응답 최대 80건의 앞 80자씩을 ' + (SOLO ? 'Gemini' : 'DIT AI ONE') + '에 한꺼번에 보냅니다. 이때 답의 이메일·전화번호 패턴을 가립니다.';
  const single = '일괄 분류에서 빠진 답을 나눌 때는 ' + (SOLO ? 'Gemini' : '수업 AI(DIT AI ONE의 Jev)') + '에 답 원문 전체를 가리지 않고 한 건씩 보낼 수 있습니다. 답에 적은 이름도 전송될 수 있습니다.';
  const studentNotice = /id="ai-notice">([^<]*)/.exec(studentHtml)[1], instructorNotice = /id="ai-desc">([^<]*)/.exec(boardHtml)[1];
  assert.ok(studentNotice.startsWith('교수가 동의하면 ' + batch));
  assert.ok(studentNotice.endsWith(single + ' 이름·연락처는 적지 마십시오.'));
  assert.ok(instructorNotice.startsWith(batch));
  assert.ok(instructorNotice.endsWith(single + ' 이름·연락처를 적지 않도록 안내해 주십시오.'));
  const gq = env.call('getQuestion');
  assert.deepEqual([gq.q, gq.open, gq.round], ['AI를 수업에 쓰는 것에서 가장 걱정되는 점은?', true, env.sheet('응답').getSheetId()]);
  const r = env.call('submitAnswer', '시험 때 몰래 쓸까 봐 걱정입니다', 'sid1');
  assert.deepEqual([r.ok, r.cat], [true, '부정행위']);
  const jevReq = env.log.jevCalls.at(-1);
  assert.deepEqual(Object.keys(jevReq.questions.cat.criteria), ['c0', 'c1', 'c2', 'c3', 'other']);
  assert.equal(env.rows('응답')[0][4], '표시');
  // low confidence falls back to 기타, a flagged answer is stored but hidden from the board
  const low = boot('r1-board.gs', { jev: (s, q) => ({ ...defaultAnswers(q), cat: choice('c1', 0.3) }) });
  low.setup();
  assert.equal(low.call('submitAnswer', '잘 모르겠어요 그냥', 's').cat, '기타');
  const spam = boot('r1-board.gs', { jev: (s, q) => ({ ...defaultAnswers(q), bad: noul(0.95) }) });
  spam.setup();
  spam.call('submitAnswer', '광고 광고 광고', 's');
  assert.equal(spam.rows('응답')[0][4], '숨김');
  assert.equal(spam.call('getBoard').items.length, 0);
  assert.equal(env.call('getBoard').items.length, 1);
  // a new round archives the answers and empties the board
  env.call('newRound');
  assert.equal(env.call('getBoard').items.length, 0, 'the board starts empty again');
  assert.equal(env.rows('응답').length, 0);
  env.call('submitAnswer', '새 질문에 대한 첫 생각', 's9');
  assert.equal(env.rows('응답').length, 1, 'new answers land in the fresh sheet');
  assert.ok([...env.ctx.SpreadsheetApp.getActive().sheets.keys()].some((n) => /^응답_\d{4}_\d{6}$/.test(n)), 'the old answers are archived, not deleted');
  // too short, closed, throttled
  assert.equal(env.call('submitAnswer', '가', 's2').ok, false);
  env.sheet('설정').getRange('B3').setValue('X');
  assert.equal(env.call('submitAnswer', '닫힌 뒤에 보냄', 's3').ok, false);
  env.sheet('설정').getRange('B3').setValue('O');
  const results = Array.from({ length: 8 }, (_, i) => env.call('submitAnswer', '연타 ' + i + ' 번째', 'spammer').ok);
  assert.deepEqual(results, [true, true, true, true, true, true, false, false], 'six per ten minutes per student');
  // Jev down: the answer is still kept
  const down = boot('r1-board.gs');
  down.setup();
  down.log.jevFail = 503;
  const d = down.call('submitAnswer', '판정이 안 될 때도 저장', 's9');
  assert.deepEqual([d.ok, d.cat], [true, '']);
  assert.equal(down.rows('응답').length, 1);
  assert.match(String(down.rows('응답')[0][3]), /^오류: .*503/, 'the failure leaves a trace with the reason, so [분류] can retry it');
  assert.ok(down.log.sleeps >= 2, 'retried before giving up');
});

// ---------- recipe 2 ----------
test('r2: Jev routes to auto answer or to the professor; answers come back by number; weekly report', () => {
  let answerable = 0.95, personal = 0.1, aiText = '출석은 4회 결석까지입니다. 「결석 4회 초과 시 F」';
  const seen = [];
  const env = boot('r2-assistant.gs', {
    jev: (s, q) => ({ ...defaultAnswers(q), answerable: noul(answerable), personal: noul(personal), kind: choice('schedule') }),
    ai: (req) => { seen.push(req); return aiText; },
  });
  env.setup();
  env.sheet('자료').appendRow(['출석: 4회 초과 결석 시 F']);
  const a = env.call('ask', '결석 몇 번까지 되나요?', 's1');
  assert.deepEqual([a.ok, a.auto], [true, true]);
  assert.ok(a.answer.includes('결석'));
  assert.equal(seen[0].files[0].text.includes('4회 초과'), true, 'the material travels as an attachment, not in the message');
  assert.equal(env.rows('질문기록')[0][5], '자동답변');
  answerable = 0.3;
  const b = env.call('ask', '제 성적이 이상한 것 같아요', 's1');
  assert.deepEqual([b.ok, b.auto, b.id.length], [true, false, 6]);
  assert.equal(env.call('checkAnswer', b.id).answer, '');
  env.sheet('질문기록').getRange(3, 7).setValue('확인해 보니 입력 오류였습니다.');
  assert.deepEqual(env.call('checkAnswer', b.id.toLowerCase()), { found: true, answer: '확인해 보니 입력 오류였습니다.' });
  assert.equal(env.call('checkAnswer', 'ZZZZZZ').found, false);
  answerable = 0.95;
  aiText = '자료에 없습니다';
  assert.equal(env.call('ask', '자료에 없는 질문입니다', 's1').auto, false, 'an AI answer of "no data" goes to the professor');
  personal = 0.9;
  aiText = '답';
  assert.equal(env.call('ask', '개인 사정 상담 질문입니다', 's2').auto, false);
  assert.equal(env.call('ask', '짧', 's3').ok, false);
  env.call('weeklyOn');
  assert.equal(env.log.triggers[0].fn, 'weeklyReport');
  env.call('weeklyReport');
  assert.equal(env.log.mails.at(-1).to, 'prof@example.com');
  assert.ok(env.log.mails.at(-1).body.includes('교수 확인 필요'));
  const empty = boot('r2-assistant.gs');
  empty.setup();
  assert.equal(empty.call('ask', '자료가 없을 때 질문', 's').ok, false);
});

// ---------- recipe 3 ----------
test('r3: Jev picks the move, the tutor follows it, history survives, wrap invites submission', () => {
  const asked = [];
  let move = 'hint', level = 1;
  const env = boot('r3-tutor.gs', {
    jev: (s, q) => ({ ...defaultAnswers(q), understanding: score(level, ['0', '1', '2', '3']), move: choice(move), off: noul(0.1), misconception: noul(0.2) }),
    ai: (req) => { asked.push(req.messages); return '그렇다면 빛은 재료일까요, 에너지일까요?'; },
  });
  env.setup();
  const sid = 'abcdef123456';
  const t1 = env.call('tutorTurn', sid, '파랑고래', '식물은 햇빛을 먹고 산다고 생각해요.');
  assert.deepEqual([t1.ok, t1.done, t1.turn], [true, false, 1]);
  assert.equal(asked[0].length, 1);
  assert.ok(asked[0][0].content.includes('소크라테스식 튜터') && asked[0][0].content.includes('[이번 턴에 할 일] 작은 힌트'));
  move = 'counter'; level = 2;
  env.call('tutorTurn', sid, '파랑고래', '그럼 빛이 재료인가요?');
  assert.deepEqual(asked[1].map((m) => m.role), ['user', 'assistant', 'user']);
  assert.ok(asked[1][2].content.includes('반례나 모순'));
  assert.ok(!asked[1][1].content.includes('이번 턴에 할 일'), 'only the last turn carries the move');
  const log = env.rows('대화록');
  assert.deepEqual(log.map((r) => [r[4], r[3]]), [['학생', 1], ['튜터', 1], ['학생', 2], ['튜터', 2]]);
  assert.deepEqual([log[0][6], log[2][6], log[2][7]], [1, 2, 'counter']);
  // cache lost: history is rebuilt from the sheet
  env.cache.clear();
  move = 'wrap'; level = 3;
  const t3 = env.call('tutorTurn', sid, '파랑고래', '빛은 에너지이고 물과 이산화탄소가 재료예요.');
  assert.equal(asked[2].length, 5);
  assert.equal(t3.done, true);
  // submit
  assert.equal(env.call('submitTranscript', sid, '파랑고래', '짧음').ok, false);
  const s = env.call('submitTranscript', sid, '파랑고래', '식물은 물과 이산화탄소를 빛에너지로 포도당과 산소로 바꾼다.');
  assert.equal(s.ok, true);
  const sub = env.rows('제출')[0];
  assert.deepEqual([sub[2], sub[3], sub[4], sub[5]], ['파랑고래', 3, 3, '1 → 2 → 3']);
  assert.equal(typeof sub[8], 'string');
  // max turns force a wrap; invalid ids and closed rooms are refused
  const cap = boot('r3-tutor.gs', { jev: (s2, q) => ({ ...defaultAnswers(q), move: choice('probe') }), ai: () => '질문' });
  cap.setup();
  cap.sheet('설정').getRange('B4').setValue(4);
  const outs = [1, 2, 3, 4, 5].map((i) => cap.call('tutorTurn', 'sid' + 'x'.repeat(9), '별', '생각 ' + i));
  assert.deepEqual(outs.map((o) => o.done), [false, false, false, true, true]);
  assert.equal(cap.call('tutorTurn', 'x', '별', '짧은 세션').ok, false);
  cap.sheet('설정').getRange('B5').setValue('X');
  assert.equal(cap.call('tutorTurn', 'sidxxxxxxxxxx', '별', '닫힘').ok, false);
  assert.ok(!env.call('doGet').html.includes('이해도'), 'students never see the score');
});

// ---------- recipe 4 ----------
test('r4: letter saved, delivered on its date, reflection compared by Jev + AI', () => {
  const seen = [];
  const env = boot('r4-future.gs', {
    jev: (s, q) => ({ ...defaultAnswers(q), growth: score(2, ['없음', '조금', '뚜렷', '크게']), tone: choice('mixed') }),
    ai: (req) => { seen.push(req.messages[0].content); return '처음엔 걱정이 컸는데 지금은 달라졌네요.'; },
  });
  env.setup();
  const info = env.call('getInfo');
  assert.match(info.date, /^\d{4}-\d{2}-\d{2}$/, 'the default date stays text even though sheets turn dates into dates');
  const tomorrow = env.ctx.Utilities.formatDate(new Date(Date.now() + 86400000), 'x', 'yyyy-MM-dd');
  const today = env.ctx.Utilities.formatDate(new Date(), 'x', 'yyyy-MM-dd');
  const draft = { alias: '파랑고래', email: 'Student@Example.com', goals: '발표를 두려워하지 않기, 매주 복습하기', worry: '시간 관리', date: tomorrow };
  assert.equal(env.call('saveLetter', { ...draft, email: 'nope' }).ok, false);
  assert.equal(env.call('saveLetter', { ...draft, goals: '짧음' }).ok, false);
  assert.equal(env.call('saveLetter', { ...draft, date: '2020-01-01' }).ok, false);
  assert.equal(env.call('saveLetter', { ...draft, date: '내일' }).ok, false);
  assert.deepEqual(env.call('saveLetter', draft), { ok: true, date: tomorrow });
  assert.equal(env.rows('편지')[0][3], 'student@example.com');
  assert.equal(env.rows('편지')[0][7], '예약');
  assert.equal(env.log.mails[0].to, 'student@example.com');
  env.call('saveLetter', { ...draft, alias: '둘째', email: 'two@example.com', date: today });
  env.props.set('WEB_URL', 'https://script.google.com/macros/s/AKfycbTEST/exec');
  env.call('deliverLetters');
  const delivered = env.log.mails.filter((m) => m.subject.includes('편지를 보낸') || m.subject.includes('나에게] 학기 초'));
  assert.equal(delivered.length, 1, 'only the letter due today is delivered');
  assert.equal(delivered[0].to, 'two@example.com');
  assert.ok(delivered[0].body.includes('?r='), 'the mail carries the reflection link');
  const dueId = env.rows('편지')[1][1], futureId = env.rows('편지')[0][1];
  assert.equal(env.call('getLetter', futureId).found, false, 'not delivered yet');
  assert.equal(env.call('getLetter', dueId).found, true);
  assert.ok(env.call('doGet', { parameter: { r: dueId } }).html.includes('id="' + dueId + '"'));
  assert.ok(!env.call('doGet', { parameter: { r: dueId + '"><script>' } }).html.includes('<script>alert'), 'the id is sanitized');
  assert.equal(env.call('submitReflection', futureId, '아직 도착 전에 쓰는 회고입니다.').ok, false);
  const r = env.call('submitReflection', dueId, '발표는 아직 떨리지만 매주 복습은 지켰어요.');
  assert.equal(r.ok, true);
  assert.ok(seen[0].includes('변화 뚜렷') && seen[0].includes('발표를 두려워하지'));
  const row = env.rows('편지')[1];
  assert.deepEqual([row[7], row[9]], ['회고완료', 2]);
  const n = env.log.requests.length;
  assert.equal(env.call('submitReflection', dueId, '다시 쓰는 회고라도 같은 코멘트').comment, r.comment);
  assert.equal(env.log.requests.length, n, 'a finished letter does not call AI again');
  // without a saved web app address nothing is mailed (a broken reflection link cannot be taken back)
  env.props.delete('WEB_URL');
  env.call('saveLetter', { ...draft, alias: '셋째', email: 'three@example.com', date: today });
  env.log.prompts.push(null);
  const mailsBefore = env.log.mails.length;
  env.call('demoDeliverAll');
  assert.ok(env.log.alerts.at(-1).startsWith('웹 앱 주소가 없어 배달하지 않았습니다'));
  assert.equal(env.log.mails.length, mailsBefore);
  env.call('deliverLetters');
  assert.ok(env.log.mails.at(-1).subject.includes('배달을 기다립니다'), 'the daily trigger tells the owner instead');
  const afterHeld = env.log.mails.length;
  env.call('deliverLetters');
  assert.equal(env.log.mails.length, afterHeld, 'the owner is told once a day, however often it is called');
  // the menu asks for the address once; /dev and junk are refused
  env.log.prompts.push('https://script.google.com/macros/s/AKfycbTEST/dev');
  env.call('setWebUrl');
  assert.ok(env.log.alerts.at(-1).includes('/dev'));
  env.log.prompts.push('https://script.google.com/macros/u/1/s/AKfycbTEST/exec?v=x');
  env.call('demoDeliverAll');
  assert.equal(env.props.get('WEB_URL'), 'https://script.google.com/macros/s/AKfycbTEST/exec', 'the /u/1 account part and the query are cleaned off');
  assert.ok(env.log.alerts.at(-1).startsWith('2통'), 'the held letter and the future one (rehearsal sends all)');
  assert.ok(env.log.mails.at(-1).body.includes('https://script.google.com/macros/s/AKfycbTEST/exec?r='), 'the letter links to the saved /exec address');
  env.call('dailyOn');
  assert.deepEqual(env.log.triggers.map((t) => t.fn), ['deliverLetters']);
});

// ---------- recipe 5 ----------
const file = (id, name, mime, size) => ({ getId: () => id, getName: () => name, getMimeType: () => mime, getBlob: () => ({ getContentType: () => mime, getBytes: () => Buffer.alloc(size, 7) }) });

test('r5: photo -> notes doc; unreadable asks for a retake; big or HEIC photos go through a resized thumbnail', () => {
  const NOTES = '# 광합성 정리\n## 재료\n- 물\n- 이산화탄소 [?]\n## 복습 퀴즈\n1. 산소는 어디서? 정답: 물';
  let unreadable = 0.1, kind = 'lecture';
  const fetched = [];
  const env = boot('r5-boardphoto.gs', {
    ai: () => NOTES,
    jev: (s, q) => ({ ...defaultAnswers(q), kind: choice(kind), unreadable: noul(unreadable) }),
    fetchExtra: (url, opts) => { fetched.push([url, opts.headers.Authorization]); return { getResponseCode: () => 200, getContentText: () => '', getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => Buffer.alloc(300000, 1) }) }; },
  });
  env.setup();
  const inbox = env.folders.get(env.props.get('INBOX'));
  inbox.files.push(file('p1', '칠판1.jpg', 'image/jpeg', 200000), file('x1', '메모.txt', 'text/plain', 10));
  env.call('processPhotos');
  const row = env.rows('기록')[0];
  assert.deepEqual([row[1], row[2], row[4]], ['칠판1.jpg', '수업 내용', '완료']);
  const doc = [...env.docs.values()][0];
  assert.deepEqual(doc.lines.map((l) => l.kind), ['P', 'P', 'H1', 'H2', 'LI', 'LI', 'H2', 'P']);
  assert.equal(doc.images.length, 1, 'the original blackboard photo sits at the top of the note');
  assert.equal(doc.lines[1].t, '원본 칠판 사진');
  assert.equal(doc.lines[2].t, '광합성 정리');
  assert.equal(env.folders.get(env.props.get('NOTES')).files.length, 1, 'the doc was moved into the result folder');
  assert.ok(env.log.mails.at(-1).body.includes(doc.getUrl()));
  env.call('processPhotos');
  assert.equal(env.rows('기록').length, 1, 'a photo is processed once');
  // unreadable -> retake mail, no doc
  unreadable = 0.9;
  inbox.files.push(file('p2', '흐림.jpg', 'image/jpeg', 100000));
  env.call('processPhotos');
  assert.equal(env.rows('기록')[1][4], '다시 찍기 필요');
  assert.equal(env.docs.size, 1);
  assert.ok(env.log.mails.at(-1).subject.includes('다시 찍어'));
  // assignment notices are flagged
  unreadable = 0.1; kind = 'assignment';
  inbox.files.push(file('p3', '공지.jpg', 'image/jpeg', 100000));
  env.call('processPhotos');
  assert.ok(env.log.mails.at(-1).body.includes('과제·시험 공지로 보입니다'));
  // big jpeg and HEIC use the thumbnail with the OAuth token, and the result passes the server's image caps
  kind = 'lecture';
  inbox.files.push(file('p4', '큰사진.jpg', 'image/jpeg', 4000000), file('p5', '아이폰.heic', 'image/heic', 900000));
  env.call('processPhotos');
  assert.equal(fetched.length, 2);
  assert.ok(fetched[0][0].includes('sz=w1600') && fetched[0][1] === 'Bearer tok');
  assert.deepEqual(env.rows('기록').slice(3).map((r) => r[4]), ['완료', '완료']);
  // a file the server cannot take is recorded as a failure, not retried forever
  const fail = boot('r5-boardphoto.gs', { ai: () => NOTES, fetchExtra: () => ({ getResponseCode: () => 404, getContentText: () => '' }) });
  fail.setup();
  fail.folders.get(fail.props.get('INBOX')).files.push(file('p9', '거대.jpg', 'image/jpeg', 5000000));
  fail.call('processPhotos');
  assert.ok(fail.rows('기록')[0][4].startsWith('실패'));
  fail.call('processPhotos');
  assert.equal(fail.rows('기록').length, 1);
  env.call('autoOn');
  assert.deepEqual(env.log.triggers.map((t) => [t.fn, t.spec[0][0]]), [['processPhotos', 'timeBased']]);
});

// ---------- recipe 6 ----------
test('r6: form + trigger, every response analysed, urgent mails at once, digest every N', () => {
  const env = boot('r6-survey.gs', {
    jev: (s, q) => ({ ...defaultAnswers(q), mood: score(3), topic: choice('pace'), suggest: noul(s.includes('더 천천히') ? 0.9 : 0.1), urgent: noul(s.includes('괴롭') ? 0.95 : 0.05) }),
    ai: () => '키워드: 속도, 실습\n칭찬 한 줄: 좋았어요',
  });
  env.setup();
  assert.deepEqual(env.log.triggers.map((t) => [t.fn, t.spec.map((s) => s[0]).join('>')]), [['onSurveySubmit', 'forForm>onFormSubmit']]);
  env.call('setup');
  assert.equal(env.log.triggers.length, 1, 'running setup twice does not duplicate the form or trigger');
  env.call('showUrls');
  assert.ok(env.log.alerts.at(-1).includes('forms.gle') && env.log.alerts.at(-1).includes('웹 도구에서 받은 주소'));
  const sub = (sat, good, want) => env.call('onSurveySubmit', { namedValues: { '오늘 수업 만족도': [String(sat)], '좋았던 점': [good], '아쉬웠던 점이나 바라는 점': [want] } });
  env.sheet('설정').getRange('B1').setValue(3);
  sub(5, '예시가 좋았어요', '더 천천히 해 주세요');
  sub(4, '실습이 좋았어요', '');
  let rows = env.rows('분석');
  assert.deepEqual([rows[0][1], rows[0][2], rows[0][3], rows[0][4], rows[0][5]], [5, 3, '속도', 'O', '']);
  assert.equal(env.log.mails.length, 0, 'no digest before three responses');
  sub(2, '', '이 수업에서 괴롭힘을 당했어요');
  assert.ok(env.log.mails[0].subject.includes('긴급'));
  const digest = env.log.mails.at(-1);
  assert.ok(digest.subject.includes('새 응답 3개 요약'));
  assert.ok(digest.body.includes('평균 만족도: 3.7') && digest.body.includes('키워드: 속도'), digest.body);
  assert.equal(env.props.get('SENT'), '3');
  assert.ok(digest.html.includes('PROF.KIMCOOK') && digest.html.includes('평균 만족도') && !digest.html.includes('<script'), 'every mail also has an HTML body in the lab style');
  assert.ok(!/gradient|border-radius|box-shadow/.test(digest.html), 'no decorative styling in mail');
  sub(5, '좋음', '좋음');
  assert.equal(env.log.mails.length, 2, 'the next digest waits for three more');
  env.call('digestNow');
  assert.ok(env.log.mails.at(-1).subject.includes('새 응답 1개'));
  env.call('digestNow');
  assert.ok(env.log.alerts.at(-1).includes('새로 쌓인 응답이 없습니다'));
  // rating only, and a Jev outage, still produce a row
  sub(3, '', '');
  env.log.jevFail = 503;
  sub(1, '판정 실패', '판정 실패');
  rows = env.rows('분석');
  assert.deepEqual([rows.at(-2)[1], rows.at(-2)[2], rows.at(-1)[3]], [3, '', '판정 실패']);
});

test('r3 class screen: PIN-gated, shows aliases and levels, never conversation text', () => {
  let level = 1;
  const env = boot('r3-tutor.gs', { jev: (s, q) => ({ ...defaultAnswers(q), understanding: score(level), move: choice('probe') }), ai: () => '왜 그렇게 생각하나요?' });
  env.setup();
  const pin = env.call('pin_', 'class');
  assert.match(pin, /^\d{4}$/);
  assert.equal(env.call('pin_', 'class'), pin, 'the PIN stays the same');
  assert.ok(env.log.alerts.at(-1).includes(pin), 'setup tells the professor the PIN');
  env.call('tutorTurn', 'sidaaaa1', '파랑고래', '빛을 먹어요');
  level = 3;
  env.call('tutorTurn', 'sidaaaa1', '파랑고래', '빛은 에너지예요');
  level = 0;
  env.call('tutorTurn', 'sidbbbb2', '노랑곰', '모르겠어요');
  const wrong = env.call('getClass', pin === '0000' ? '1111' : '0000');
  assert.equal(wrong.ok, false);
  const r = env.call('getClass', pin);
  assert.equal(r.total, 2);
  assert.deepEqual(r.students.find((x) => x.a === '파랑고래').lv, [1, 3]);
  assert.deepEqual(r.dist, [1, 0, 0, 1]);
  assert.ok(!JSON.stringify(r).includes('에너지예요'), 'no conversation text leaves the server');
  const doc = env.call('doGet', { parameter: { v: 'class' } });
  assert.ok(doc.html.includes('교수 화면 PIN') && doc.xframe === 'ALLOWALL');
  assert.ok(!env.rows('설정').flat().includes(pin), 'the PIN is not written into the sheet');
  // the right PIN hands out a screen token; with it the screen keeps working even when someone else locks the PIN out
  const tok = r.token;
  assert.match(tok, /^t[0-9a-f]{24,}$/);
  for (let k = 0; k < 10; k++) env.call('getClass', pin === '0000' ? '1111' : '0000');
  assert.equal(env.call('getClass', pin).locked, true, 'ten wrong PINs lock new PIN entries');
  assert.equal(env.call('getClass', tok).ok, true, 'an open screen keeps its token');
  assert.equal(env.call('getClass', 't' + 'a'.repeat(32)).expired, true, 'an unknown token asks for the PIN again without counting');
  env.cache.delete('pinfail:class');
  const pg = runPage(env, env.ctx.CLASS_HTML);
  pg.set('#pin', '12');
  pg.click('#go');
  assert.equal(pg.el('#err').textContent, 'PIN은 숫자 4자리입니다.');
  assert.deepEqual(pg.calls, [], 'a malformed PIN is not sent (and does not count toward the lockout)');
  pg.set('#pin', pin);
  pg.click('#go');
  assert.equal(pg.el('#n').textContent, '2');
  assert.equal(pg.el('#grid').children.length, 2);
  assert.equal(pg.el('#gate').hidden, true);
  pg.click('#go');
  assert.equal(pg.timers.filter((t) => t[1] === 5000).length, 1, 'one polling loop, however often the button is pressed');
});

test('r6 report screen: totals and rankings for the big screen, never urgent rows or complaints', () => {
  const env = boot('r6-survey.gs', {
    jev: (s, q) => ({ ...defaultAnswers(q), mood: score(3), topic: choice('pace'), suggest: noul(0.1), urgent: noul(s.includes('괴롭') ? 0.95 : 0.05) }),
    ai: () => '키워드: 속도',
  });
  env.setup();
  const sub = (sat, good, want) => env.call('onSurveySubmit', { namedValues: { '오늘 수업 만족도': [String(sat)], '좋았던 점': [good], '아쉬웠던 점이나 바라는 점': [want] } });
  sub(5, '예시가 좋았어요', '조금 빨랐어요');
  sub(4, '실습이 좋았어요', '');
  sub(2, '숨겨질 칭찬', '이 수업에서 괴롭힘을 당했어요');
  const r = env.call('getReport');
  assert.equal(r.n, 3);
  assert.equal(r.avg, 3.7);
  assert.deepEqual(r.sat, [0, 1, 0, 1, 1]);
  assert.equal(r.mood[3], 3);
  assert.deepEqual(r.topics, [{ k: '속도', n: 3 }]);
  const wire = JSON.stringify(r);
  assert.ok(!wire.includes('괴롭') && !wire.includes('숨겨질') && !wire.includes('조금 빨랐어요'), 'urgent rows and the "want" column stay off the screen');
  assert.deepEqual(r.quotes, ['예시가 좋았어요', '실습이 좋았어요']);
  const doc = env.call('doGet', {});
  assert.equal(doc.xframe, 'ALLOWALL');
  const pg = runPage(env, env.ctx.REPORT_HTML);
  assert.equal(pg.el('#avg').textContent, '3.7');
  assert.equal(pg.el('#hist').children.length, 5);
  assert.equal(pg.el('#mood').children.length, 5);
  assert.equal(pg.el('#tp').children.length, 1);
  assert.equal(pg.el('#wait').hidden, true);
  const empty = boot('r6-survey.gs', {});
  empty.setup();
  const pe = runPage(empty, empty.ctx.REPORT_HTML);
  assert.equal(pe.el('#main').hidden, true, 'an empty survey shows the waiting message');
});

test('one-step start: pasting the code and running start() prepares everything, with the class code embedded', () => {
  const expectTriggers = { 'r1-board.gs': [], 'r2-assistant.gs': ['weeklyReport'], 'r3-tutor.gs': [], 'r4-future.gs': ['deliverLetters'], 'r5-boardphoto.gs': ['processPhotos'], 'r6-survey.gs': ['onSurveySubmit'] };
  for (const [f, trig] of Object.entries(expectTriggers)) {
    const env = createEnv();
    env.props.delete('CLASS_CODE');
    env.run(COMMON + '\n' + readGs(f) + "\n\nvar EMBEDDED_CODE = 'TESTCODE';\n", f);
    assert.match(env.run('start.toString()').slice(0, 20), /function start/);
    env.ctx.start();
    assert.ok(env.log.alerts.at(-1).startsWith('준비 끝'), f + ': ' + env.log.alerts.at(-1));
    assert.ok(env.log.logged.at(-1).startsWith('준비 끝'), 'the message also reaches the execution log');
    assert.equal(env.log.menus.length, 1, f + ': the menu is installed without reopening the sheet');
    assert.deepEqual(env.log.triggers.map((t) => t.fn), trig, f);
    assert.ok(env.active.sheets.size >= 1, f);
    // no CLASS_CODE property: the embedded code is used
    assert.deepEqual(JSON.parse(JSON.stringify(env.ctx.jev_('x', { a: { type: 'noul', instructions: 'q' } }))).a.type, 'noul');
    assert.ok(!/수업 코드 입력/.test(env.log.alerts.at(-1)), 'no leftover step about typing the code');
  }
  const blank = createEnv();
  blank.run(COMMON, 'blank');
  blank.ctx.start();
  assert.ok(blank.log.alerts.at(-1).includes('start 단계가 없습니다'));
  assert.ok(/^function start\(/m.exec(COMMON) && COMMON.indexOf('function start') < COMMON.search(/function (?!start)\w+\(/), 'start is the first function, so the editor selects it by default');
});

test('r1 big screen edits: PIN-gated question change, optional archive, brute force locked out', () => {
  const env = boot('r1-board.gs');
  env.setup();
  const pin = env.call('pin_', 'board');
  assert.ok(env.log.alerts.at(-1).includes(pin), 'setup tells the professor the PIN');
  env.call('submitAnswer', '시험 때 몰래 쓸까 봐 걱정입니다', 's1');
  const wrong = pin === '0000' ? '1111' : '0000';
  assert.equal(env.call('setBoard', wrong, '새 질문', '가, 나', true, true).ok, false);
  assert.equal(env.call('setBoard', pin, '가', '', true, false).msg, '질문을 두 글자 이상 적어 주세요.');
  assert.ok(env.call('setBoard', pin, '오늘 수업에서 가장 헷갈린 것은?', '용어', true, false).msg.includes('두 개 이상'));
  const r = env.call('setBoard', pin, '오늘 수업에서 가장 헷갈린 것은?', '용어\n개념, 개념\n기타\n예시', false, true);
  assert.equal(r.ok, true);
  assert.match(r.token, /^t[0-9a-f]{24,}$/, 'the screen gets a token so the PIN is typed once');
  assert.equal(env.call('setBoard', r.token, '오늘 수업에서 가장 헷갈린 것은?', '용어\n개념\n예시', false, false).ok, true, 'the token works like the PIN');
  assert.equal(r.archived, true, 'old answers are kept in an archive sheet');
  assert.deepEqual(r.board.cats, ['용어', '개념', '예시', '기타'], 'duplicates and a typed 기타 are dropped');
  assert.equal(r.board.q, '오늘 수업에서 가장 헷갈린 것은?');
  assert.equal(r.board.open, false);
  assert.equal(r.board.total, 0, 'the screen starts empty');
  assert.ok([...env.ctx.SpreadsheetApp.getActive().sheets.keys()].some((n) => n.startsWith('응답_')));
  assert.equal(env.call('submitAnswer', '닫힌 뒤', 's2').ok, false, 'closing from the screen stops answers');
  const kept = env.call('setBoard', pin, '같은 질문 유지', '', true, false);
  assert.equal(kept.archived, false);
  assert.deepEqual(kept.board.cats, [], 'empty groups mean one column');
  env.call('submitAnswer', '하나로 모입니다', 's3');
  env.cache.clear();
  assert.equal(env.call('getBoard').items[0].c, '');
  const t2 = env.call('setBoard', pin, '같은 질문 유지', '', true, false).token;
  for (let k = 0; k < 10; k++) env.call('setBoard', wrong, '틀린 PIN', '', true, false);
  assert.equal(env.call('setBoard', pin, '맞는 PIN이지만 잠김', '', true, false).ok, false, 'after ten wrong PINs even the right one waits');
  assert.equal(env.call('setBoard', t2, '표가 있는 화면은 계속', '', true, false).ok, true, 'a screen that already has a token is not locked out');
});

test('r1 groups on the student phone: chips from the question, a picked group skips the Jev classifier, AI suggests groups for the professor', () => {
  let jevCalls = [];
  const env = boot('r1-board.gs', { jev: (st, q) => { jevCalls.push(Object.keys(q)); return { ...defaultAnswers(q), bad: noul(String(st).includes('광고') ? 0.95 : 0.05) }; }, ai: ({ messages }) => '1. 부정행위\n- 평가 방법\n  학습 효과 저하 \n기타\n업무 부담\n부정행위' });
  env.setup();
  assert.deepEqual(env.call('getQuestion').cats, ['부정행위', '평가 방법', '학습 효과 저하', '업무 부담']);
  // a picked group is stored as chosen; only the spam filter runs
  const p = env.call('submitAnswer', '직접 골랐어요', 's1', undefined, '평가 방법');
  assert.deepEqual([p.ok, p.cat], [true, '평가 방법']);
  assert.deepEqual(jevCalls, [['bad']], 'no classifier call for a picked group');
  assert.deepEqual(env.rows('응답')[0].slice(2, 5), ['평가 방법', '직접', '표시']);
  assert.equal(env.call('submitAnswer', '기타로 보냅니다', 's1', undefined, '기타').cat, '기타');
  // a name that is not in the list is ignored and Jev classifies as before
  jevCalls = [];
  const ign = env.call('submitAnswer', '없는 묶음을 눌렀다고 해요', 's1', undefined, '지어낸 묶음');
  assert.equal(ign.cat, '부정행위');
  assert.deepEqual(jevCalls, [['cat', 'bad']]);
  // a picked answer that is spam is still hidden
  env.call('submitAnswer', '광고 광고 광고', 's2', undefined, '업무 부담');
  assert.equal(env.rows('응답').at(-1)[4], '숨김');
  // no groups: no chips, a pick means nothing
  const pin = env.call('pin_', 'board');
  env.call('setBoard', pin, '그룹 없는 질문', '', true, false);
  env.cache.clear();
  assert.deepEqual(env.call('getQuestion').cats, []);
  assert.equal(env.call('submitAnswer', '묶음 없이 보냄', 's3', undefined, '평가 방법').cat, '');
  // AI suggestion: PIN-gated, cleaned to one name per line, never saved by itself
  assert.equal(env.call('suggestCats', pin === '0000' ? '1111' : '0000', '오늘 가장 어려웠던 점은?').ok, false);
  assert.equal(env.call('suggestCats', pin, '가').msg, '질문을 먼저 적어 주세요.');
  const sug = env.call('suggestCats', pin, '오늘 가장 어려웠던 점은?');
  assert.deepEqual([sug.ok, sug.cats], [true, '부정행위\n평가 방법\n학습 효과 저하\n업무 부담'], 'numbers and bullets stripped, 기타 and duplicates dropped');
  assert.match(sug.token, /^t[0-9a-f]{24,}$/);
  assert.equal(env.call('getBoard').q, '그룹 없는 질문', 'suggesting does not change the board');
  assert.equal(env.call('suggestCats', sug.token, '오늘 가장 어려웠던 점은?').ok, true, 'the screen token works');
  const dead = boot('r1-board.gs', { ai: () => { throw new Error('AI ONE 401: 수업 코드가 맞지 않습니다'); } });
  dead.setup();
  assert.match(dead.call('suggestCats', dead.call('pin_', 'board'), '질문입니다').msg, /수업 코드가 맞지 않습니다/, 'the reason reaches the professor');
  const bad = boot('r1-board.gs', { ai: () => '하나만' });
  bad.setup();
  assert.match(bad.call('suggestCats', bad.call('pin_', 'board'), '질문입니다').msg, /묶음을 찾지 못했습니다/);
});

test('r1 safety: student text never runs as a formula, archives never collide, answers keep their id, web callers cannot archive', () => {
  const env = boot('r1-board.gs');
  env.setup();
  const pin = env.call('pin_', 'board');
  assert.ok(!env.rows('설정').flat().includes(pin), 'the PIN is not in the sheet, so no formula can read it');
  for (const t of ['=설정!B5', "='설정'!B1&\" 걱정\"", '+1+1 같아요', '@멘션', '-걱정입니다']) assert.equal(env.call('submitAnswer', t, 's' + t.length).ok, true);
  const shown = env.call('getBoard').items.map((x) => x.t);
  assert.deepEqual(shown, ['=설정!B5', "='설정'!B1&\" 걱정\"", '+1+1 같아요', '@멘션', '-걱정입니다'], 'stored and shown as typed text');
  assert.ok(!JSON.stringify(env.rows('응답')).includes('#FORMULA'), 'no cell was evaluated');
  const ok = env.call('setBoard', pin, '=1+1 질문', '=A, B', true, false);
  assert.equal(ok.board.q, '=1+1 질문');
  assert.deepEqual(ok.board.cats, ['=A', 'B', '기타']);
  // ids survive a deleted row; numbers follow the order in the round
  env.cache.clear();
  const before = env.call('getBoard').items;
  assert.deepEqual(before.map((x) => x.no), [1, 2, 3, 4, 5]);
  assert.equal(new Set(before.map((x) => x.i)).size, 5);
  env.sheet('응답').rows.splice(2, 1);
  env.cache.clear();
  const after = env.call('getBoard').items;
  assert.deepEqual(after.map((x) => x.i), before.filter((_, k) => k !== 1).map((x) => x.i), 'deleting a row in the sheet does not shift the other answers');
  assert.equal(env.call('submitAnswer', '하나 더', 'zz').no, 5);
  // two archives in the same second get different names; the round id changes
  const round = env.call('getBoard').round;
  env.call('setBoard', pin, '새 질문 하나', '', true, true);
  env.call('submitAnswer', '두 번째 판', 'yy');
  const again = env.call('setBoard', pin, '새 질문 둘', '', true, true);
  assert.equal(again.archived, true);
  assert.equal([...env.active.sheets.keys()].filter((n) => n.startsWith('응답_')).length, 2);
  assert.notEqual(again.board.round, round);
  assert.ok(again.board.rev > 0);
  // a phone that still shows the old round cannot drop its answer into the new one
  const oldRound = again.board.round;
  env.call('setBoard', pin, '새 질문 셋', '', true, true);
  env.call('submitAnswer', '셋째 판 첫 답', 'ww');
  env.call('setBoard', pin, '새 질문 넷', '', true, true);
  const stale = env.call('submitAnswer', '예전 화면에서 보냄', 'vv', env.call('getBoard').round - 1);
  assert.deepEqual([stale.ok, stale.stale], [false, true]);
  assert.equal(env.call('submitAnswer', '지금 화면에서 보냄', 'uu', env.call('getQuestion').round).ok, true);
  assert.notEqual(oldRound, env.call('getQuestion').round);
  // the menu-only archive refuses web callers (google.script.run has no sheet UI)
  env.call('submitAnswer', '세 번째 판', 'xx');
  env.ctx.SpreadsheetApp.getUi = () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); };
  assert.throws(() => env.call('newRound'), /시트 메뉴에서만/);
  assert.throws(() => env.call('showUrls'), /시트 메뉴에서만/);
  env.cache.clear();
  assert.equal(env.call('getBoard').total, 2, 'nothing was archived');
});

test('r1 keyword stage: one clear group name decides without Jev; generic, shared, negated or two-group words go to Jev', () => {
  let calls = [];
  const env = boot('r1-board.gs', { jev: (st, q) => { calls.push(Object.keys(q)); return { ...defaultAnswers(q), bad: noul(String(st).includes('광고') ? 0.95 : 0.05) }; } });
  env.setup();
  const send = (t) => { calls = []; return env.call('submitAnswer', t, 'k' + Math.random()); };
  assert.equal(send('업무 부담이 커질까 걱정').cat, '업무 부담');
  assert.deepEqual(calls, [['bad']], 'decided by words: only the spam check calls Jev');
  assert.deepEqual(env.rows('응답')[0].slice(2, 7), ['업무 부담', '낱말', '표시', env.rows('응답')[0][5], '업무 부담'], 'the method and the auto group are recorded');
  assert.equal(send('보고서 평가 방법이 궁금합니다').cat, '평가 방법', 'one whole two-word name decides');
  assert.equal(send('학습 효과 저하가 걱정').cat, '학습 효과 저하', 'a three-word name decides only when it is there whole');
  for (const t of ['사용 방법을 잘 몰라서', '부정행위는 별로 걱정 안 돼요', '과제 평가 때문에 업무가 늘어요', '시험 때 몰래 쓸까 봐 걱정입니다']) {
    send(t);
    assert.deepEqual(calls, [['cat', 'bad']], t + ' goes to Jev');
    assert.notEqual(env.rows('응답').at(-1)[3], '낱말');
  }
  // review 4: negations (also written together), a lone shared word, a word inside a longer word, and an easy-looking answer go to Jev
  for (const t of ['부정행위는 안생길 듯합니다', '새 도구의 사용 순서가 헷갈립니다', '실습 때 학습 효과가 달라질 듯합니다', '사전리허설로 연습했어요', '발표 평가 방법은 쉬웠어요', '조교 일이 늘어나는 듯합니다', '부정행위는우려없음', '모의부정행위로 연습했습니다',
    // re-review 1: negation written inside a word, opposite claims without a negation word, comparisons, and a three-word name met only in part
    '부정행위는 걱정없음', '부정행위 걱정은 안됨', '부정행위는 큰 문제 아닐 듯', '학습 효과가 좋아질 것 같아요', '업무 부담이 감소할 듯',
    '부정행위보다는 평가가 더 걱정돼요', '학습 효과는 그대로일 것', '혹시 업무 부담이 커지는 건 아닐까요']) {
    send(t);
    assert.deepEqual(calls, [['cat', 'bad']], t + ' goes to Jev');
  }
  send('업무 부담이 크다는 광고 클릭');
  assert.deepEqual(env.rows('응답').at(-1).slice(3, 5), ['낱말', '숨김'], 'spam is still hidden on the word path');
  // pure helpers
  const cats = ['용어', '개념 이해', '과제', '수업 속도'], q = '오늘 수업에서 가장 어려웠던 점은?';
  assert.equal(env.ctx.matchCat_('전문 용어가 생소했다', cats, q), 0);
  assert.equal(env.ctx.matchCat_('수업 자료를 미리 못 받아서', cats, q), -1, 'a word of the question (수업) never decides');
  assert.equal(env.ctx.matchCat_('과제 설명이 이해가 안 됐어요', cats, q), -1, 'two groups: Jev decides');
  assert.equal(env.ctx.matchCat_('과제는 쉬웠어요', cats, q), -1, 'q2: an easy-looking answer (쉬웠어요) is not 과제 trouble');
  assert.equal(env.ctx.matchCat_('용어보다 개념이 어려웠어요', cats, q), -1, 'a comparison (~보다) names the other thing');
  assert.equal(env.ctx.matchCat_('개인정보 유출이 걱정', ['정보', '보안 사고'], 'q'), -1, 'a one-word name is a word head, never a piece of a longer word (정보 in 개인정보)');
  assert.equal(env.ctx.matchCat_('정보 유출이 걱정', ['정보', '보안 사고'], 'q'), 0);
  assert.equal(env.ctx.matchCat_('시간이 부족해요', ['시간', '비용'], 'q'), -1, 'a common word (시간) never decides');
  assert.equal(env.ctx.matchCat_('시간 관리가 어려워요', ['시간 관리', '비용'], 'q'), -1, 'a name made only of common words never decides, even whole');
  assert.equal(env.ctx.matchCat_('개인정보 보호가 걱정', ['개인정보 보호', '비용'], 'q'), 0, 'a two-word name inside the answer');
  assert.equal(env.ctx.matchCat_('학습 방법', ['학습 효과', '학습 방법'], 'q'), 1, 'the shared word 학습 does not count, the whole name does');
});

test('r1 [분류]: PIN only, re-sorts 기타/unsorted/error rows by words then Jev, keeps chosen and hand-edited rows, caps Jev, blocks double presses', () => {
  const jevFor = (st) => (String(st).includes('몰래') ? choice('c0', 1) : String(st).includes('늘어나는') ? choice('c3', 0.9) : choice('other', 1));
  const env = boot('r1-board.gs', { jev: (st, q) => ({ ...defaultAnswers(q), ...(q.cat ? { cat: jevFor(st) } : {}) }) });
  env.setup();
  const pin = env.call('pin_', 'board'), sh = env.sheet('응답'), now = new Date();
  sh.appendRow([now, '아직 생각이 잘 안납니다', '', '', '표시', 1000001]);
  sh.appendRow([now, '팀 실습 뒤 업무 부담이 늘어나는 건 아닐까요', '', '', '표시', 1000002]);
  sh.appendRow([now, '발표 평가 방법이 궁금합니다', '', '오류: AI ONE 401: 수업 코드가 맞지 않습니다.', '표시', 1000003]);
  sh.appendRow([now, '시험 때 몰래 쓸까 봐', '기타', 0.4, '표시', 1000004, '기타']);
  sh.appendRow([now, '업무가 많아서 기타로 골랐어요', '기타', '직접', '표시', 1000005, '기타']);
  sh.appendRow([now, '업무 부담은 아니고 다른 걱정', '기타', 0.9, '표시', 1000006, '업무 부담']);
  sh.appendRow([now, '업무 광고입니다', '', '', '숨김', 1000007]);
  env.cache.clear();
  assert.equal(env.call('getBoard').counts['기타'], 6);
  const wrong = pin === '0000' ? '1111' : '0000';
  const no = env.call('classifyBoard', wrong);
  assert.equal(no.ok, false);
  assert.equal(env.rows('응답')[0][2], '', 'a wrong PIN changes nothing');
  const before = env.log.jevCalls.length;
  const r = env.call('classifyBoard', pin);
  assert.equal(r.ok, true);
  assert.match(r.token, /^t[0-9a-f]{24,}$/);
  const rows = env.rows('응답');
  assert.deepEqual(rows.map((x) => x[2]), ['기타', '업무 부담', '평가 방법', '부정행위', '기타', '기타', ''], 'sorted; chosen, hand-edited and hidden rows untouched');
  assert.deepEqual(rows.map((x) => x[3]), [1, 0.9, '낱말', 1, '직접', 0.9, ''], 'a hedged answer (~아닐까) goes to Jev, the whole name is placed by words');
  assert.deepEqual(rows.slice(0, 4).map((x) => x[6]), ['기타', '업무 부담', '평가 방법', '부정행위'], 'the auto column follows what was written');
  assert.equal(env.log.jevCalls.length - before, 3, 'Jev only for the three the words could not place');
  assert.deepEqual(Object.keys(env.log.jevCalls.at(-1).questions), ['cat'], 'a row that already had a spam check is not checked again');
  assert.deepEqual([r.moved, r.words, r.jev, r.left, r.more], [3, 1, 3, 3, 0]);
  assert.deepEqual(r.board.counts, { 부정행위: 1, '평가 방법': 1, '학습 효과 저하': 0, '업무 부담': 1, 기타: 3 });
  assert.ok(r.board.mix > 0 && r.board.rev > 0, 'every open board redraws its columns');
  // again: what Jev already called 기타 under this list is not asked again
  const again = env.call('classifyBoard', r.token);
  assert.deepEqual([again.ok, again.jev, again.moved], [true, 0, 0]);
  // double press while running, and the per-minute limit
  env.cache.set('classifying', '1');
  assert.equal(env.call('classifyBoard', pin).busy, true);
  env.cache.delete('classifying');
  env.cache.delete('t:classify');
  const tries = Array.from({ length: 5 }, () => env.call('classifyBoard', pin).ok);
  assert.deepEqual(tries, [true, true, true, true, false], 'four presses a minute');
  // the Jev cap: the rest waits for the next press
  env.cache.delete('t:classify');
  for (let k = 0; k < 5; k++) env.call('submitAnswer', '아직 잘 모르겠어요 ' + k, 'cap' + k);
  env.ctx.CLASSIFY_MAX = 3;
  env.log.jevFail = 0;
  const rowsOf = () => env.rows('응답').slice(-5).map((x) => x[3]);
  // make them unsorted error rows as if Jev had been down
  env.rows('응답').forEach((x, k) => { if (String(x[1]).startsWith('아직 잘')) { env.sheet('응답').getRange(k + 2, 3, 1, 2).setValues([['', '오류: AI ONE 503: Jev down']]); env.sheet('응답').getRange(k + 2, 7).setValue(''); } });
  const capped = env.call('classifyBoard', pin);
  assert.deepEqual([capped.jev, capped.more], [3, 2]);
  assert.equal(rowsOf().filter((x) => /^오류/.test(String(x))).length, 2, 'two left for the next press');
  assert.deepEqual([env.call('classifyBoard', pin).jev, rowsOf().filter((x) => /^오류/.test(String(x))).length], [2, 0]);
  // Jev down while sorting: nothing is overwritten, the reason comes back
  env.call('submitAnswer', '그냥 궁금해서 써 봐요', 'z9');
  env.log.jevFail = 503;
  env.call('submitAnswer', '또 다른 답입니다', 'z8');
  env.cache.delete('t:classify');
  const down = env.call('classifyBoard', pin);
  assert.ok(down.ok && down.jev === 0 && down.more >= 1 && /503/.test(down.stopped));
  assert.match(String(env.rows('응답').at(-1)[3]), /^오류/);
  // a new question (archive) while Jev is thinking: nothing is written into the archived answers
  let arch = null;
  const race = boot('r1-board.gs', { jev: (st, q) => { if (q.cat && arch) { arch.ctx.archive_(); arch = null; } return { ...defaultAnswers(q), ...(q.cat ? { cat: choice('c0', 1) } : {}) }; } });
  race.setup();
  race.sheet('응답').appendRow([new Date(), '시험 때 몰래 쓸까 봐', '', '오류: x', '표시', 1000001]);
  arch = race;
  const rr = race.call('classifyBoard', race.call('pin_', 'board'));
  assert.deepEqual([rr.ok, rr.moved], [true, 0]);
  const old = [...race.active.sheets.keys()].find((n) => n.startsWith('응답_'));
  assert.deepEqual(race.active.getSheetByName(old).rows[1].slice(2, 4), ['', '오류: x'], 'the archived row is untouched');
});

test('r1 [분류] without groups: one press groups similar answers by itself with no API, saves the names, and can be undone', () => {
  const env = boot('r1-board.gs', { jev: (st, q) => ({ ...defaultAnswers(q), ...(q.cat ? { cat: choice('other', 1) } : {}) }) });
  env.setup();
  const pin = env.call('pin_', 'board'), q = '다음 학기에 수업에서 바꿔 보고 싶은 한 가지는?';
  env.call('setBoard', pin, q, '', true, false);
  const answers = ['과제를 AI 활용형으로 바꾸기', '과제 수를 줄이고 깊게', '과제 대신 프로젝트로', '과제 제출을 온라인으로', '피드백을 빨리 돌려주기', '매주 피드백 시간 만들기',
    '개별 피드백 늘리기', '시험 방식을 오픈북으로', '서술형 시험 늘리기', '시험을 구술로', '토론 시간 늘리기', '실습 시간 늘리기', '아직 모르겠어요', '강의실 바꾸기'];
  answers.forEach((t, k) => env.call('submitAnswer', t, 'g' + k));
  assert.equal(env.log.jevCalls.length, 0, 'an ungrouped question never calls Jev');
  assert.deepEqual([env.rows('응답')[0][2], env.rows('응답')[0][6]], ['', '미분류 (자동, 묶음 없음)'], 'a new unsorted row records a blank automatic result');
  const reqs = env.log.requests.length;
  const r = env.call('classifyBoard', pin);
  assert.equal(r.ok, true);
  assert.equal(r.grouped, false);
  assert.deepEqual([r.askAi, r.aiOff], [false, true], 'no consent prompt when words are sufficient');
  assert.deepEqual([r.auto.kind, r.auto.groups, r.auto.moved, r.autoErr, r.stopped], ['words', 3, 10, '', ''], 'about 70% grouped by words: used as is');
  assert.equal(env.log.requests.length, reqs, 'no API request at all (no Jev, no AI)');
  assert.deepEqual(r.board.cats, ['과제', '피드백', '시험', '기타'], 'the names are saved at once, no suggestion box');
  assert.deepEqual([r.board.counts['과제'], r.board.counts['피드백'], r.board.counts['시험'], r.board.counts['기타']], [4, 3, 3, 4]);
  assert.equal(r.left, 4);
  assert.equal('groups' in r, false, 'no suggestion payload any more');
  const rows = env.rows('응답');
  assert.deepEqual(rows[0].slice(2, 4), ['과제', '낱말']);
  assert.equal(rows[0][6], '과제', 'the auto column follows the group');
  assert.equal(env.sheet('설정').getRange('B2').getValue(), '과제, 피드백, 시험');
  // the saved list also reaches the student phone chips
  assert.deepEqual(env.call('getQuestion').cats, ['과제', '피드백', '시험']);
  // undo: empty the group box and press 바꾸기 -> the board is ungrouped again, nothing archived
  const back = env.call('setBoard', r.token, q, '', true, false);
  assert.deepEqual([back.ok, back.archived, back.board.cats, back.board.total], [true, false, [], 14]);
  assert.deepEqual(env.call('getQuestion').cats, []);
  // groups exist but many answers stay in 기타: the next press groups those too, next to the old names
  env.call('setBoard', r.token, q, '과제\n발표', true, false);
  const many = env.call('classifyBoard', r.token);
  assert.ok(many.auto && many.auto.groups >= 1 && many.board.cats.includes('과제') && many.board.cats.length <= 11, 'new names are added after the old ones');
  assert.equal(many.board.cats.filter((k) => k === '과제').length, 1, 'an existing name is not made twice');
  assert.equal(env.ctx.clusterAnswers_(['하나', '둘'], 'q', 6, []).length, 0, 'nothing in common: no group');
  // hand-edited, chosen and hidden rows stay out of the grouping
  const keep = boot('r1-board.gs');
  keep.setup();
  keep.call('setBoard', keep.call('pin_', 'board'), q, '', true, false);
  const kpin = keep.call('pin_', 'board'), ksh = keep.sheet('응답'), now = new Date();
  answers.forEach((t, k) => ksh.appendRow([now, t, '', '', k === 0 ? '숨김' : '표시', 2000001 + k, '']));
  ksh.appendRow([now, '과제는 손으로 고친 답', '내가 고침', 0.9, '표시', 2000100, '과제']);
  ksh.getRange(4, 3).setValue('교수의 묶음'); // legacy G is blank; this row is between auto writes
  const writes = [], put = ksh.put.bind(ksh);
  ksh.put = (row, col, val) => { writes.push(row); put(row, col, val); };
  const kr = keep.call('classifyBoard', kpin);
  assert.ok(kr.auto && kr.auto.kind === 'words');
  const krows = keep.rows('응답');
  assert.equal(krows[0][2], '', 'hidden row untouched');
  assert.equal(krows[14][2], '내가 고침', 'hand-edited row untouched');
  assert.deepEqual([krows[2][2], krows[2][6], writes.includes(4)], ['교수의 묶음', '', false], 'manual row without G is not even written back');
  assert.deepEqual(krows[1].slice(2, 4), ['과제', '낱말'], 'the rest is placed');
  keep.cache.clear();
  const grouped = keep.call('classifyBoard', kpin);
  assert.equal(grouped.ok, true);
  assert.deepEqual([krows[2][2], krows[2][6], writes.includes(4)], ['교수의 묶음', '', false], 'manual row is protected from the existing-group pass too');
});

test('r1 [분류] without a word grouping: opt in for AI; only what AI missed goes to Jev; a failed AI saves nothing', () => {
  const email = 'student.fixture' + '@' + 'example.invalid', phone = ['010', '0'.repeat(4), '0'.repeat(4)].join('-');
  const answers = ['수업 시간에 졸려요', '친구들과 이야기할 시간이 없어요', '교실이 너무 추워요', '집에 가는 길이 멀어요', '점심이 늦어요 ' + email, '화장실이 멀어요 ' + phone,
    '프로젝터 소리가 작아요', '강의실 의자가 불편해요', '와이파이가 느려요', '콘센트가 부족해요', '주차장이 좁아요', '자판기가 고장났어요'];
  const aiCalls = [];
  const jevSeen = [];
  const env = boot('r1-board.gs', {
    ai: ({ messages }) => { aiCalls.push(messages.at(-1).content); return '## 시설 | 3, 4, 5, 6\n환경 | 7,8,9\n기타 | 1, 2\n'; },
    jev: (st, q) => { if (q.cat) jevSeen.push(JSON.parse(st).응답); return { ...defaultAnswers(q), ...(q.cat ? { cat: choice('c0', 0.9) } : {}) }; },
  });
  env.setup();
  const pin = env.call('pin_', 'board'), q = '수업이 불편한 점은?';
  env.call('setBoard', pin, q, '', true, false);
  answers.forEach((t, k) => env.call('submitAnswer', t, 'a' + k));
  const wrong = String((Number(pin) + 1) % 10000).padStart(4, '0');
  assert.equal(env.call('classifyBoard', pin, true).ok, false, 'a bare flag cannot opt in');
  const pending = env.call('classifyBoard', pin);
  assert.deepEqual([pending.askAi, pending.aiOff, pending.auto, aiCalls.length], [true, true, null, 0], 'without a prior prompt, no text is sent or grouped');
  assert.equal(env.sheet('설정').getRange('B2').getValue(), '');
  assert.equal(env.props.has('BOARD_AI_GROUPING'), false);
  assert.equal(env.call('classifyBoard', wrong, pending.confirmKey).ok, false, 'an unauthenticated caller cannot opt in');
  const r = env.call('classifyBoard', pending.token, pending.confirmKey);
  assert.equal(r.ok, true);
  assert.equal(env.call('classifyBoard', pending.token, pending.confirmKey).ok, false, 'a used key cannot be replayed');
  assert.equal(aiCalls.length, 1, 'one AI call');
  assert.equal(r.aiOff, false);
  assert.equal(env.props.get('BOARD_AI_GROUPING'), '1');
  assert.ok(aiCalls[0].includes('하나도 빠뜨리지 말고') && aiCalls[0].includes('12글자') && aiCalls[0].includes('1. 수업 시간에 졸려요'), 'the answer list is numbered');
  assert.ok(!aiCalls[0].includes(email) && !aiCalls[0].includes(phone));
  assert.ok(aiCalls[0].includes('점심이 늦어요 ' + '*'.repeat(email.length)) && aiCalls[0].includes('화장실이 멀어요 ' + '*'.repeat(phone.length)));
  assert.equal(r.auto.kind, 'ai');
  assert.equal(r.auto.groups, 2);
  assert.deepEqual(r.board.cats, ['시설', '환경', '기타'], '기타 stays separate and is never created as a group');
  const rows = env.rows('응답');
  assert.deepEqual(rows.slice(2, 6).map((x) => [x[2], x[3], x[6]]), Array(4).fill(['시설', 'AI', '시설']));
  assert.deepEqual(jevSeen.sort(), [...answers.slice(0, 2), ...answers.slice(9)].sort(), 'only the answers AI left out go to Jev (new list), about CLASSIFY_MAX at most');
  assert.deepEqual(Object.keys(env.log.jevCalls.at(-1).questions.cat.criteria), ['c0', 'c1', 'other'], 'Jev sees the saved list');
  env.cache.delete('t:all'); // advance the shared request window between rounds
  env.call('setBoard', r.token, q + ' (다음 질문)', '', true, true);
  answers.forEach((t, k) => env.call('submitAnswer', t, 'repeat' + k));
  assert.equal(env.call('classifyBoard', r.token).askAi, false, 'consent persists for the next round');
  assert.equal(aiCalls.length, 2);
  const getUi = env.ctx.SpreadsheetApp.getUi;
  env.ctx.SpreadsheetApp.getUi = () => { throw new Error('spreadsheet unavailable from a web app'); };
  assert.throws(() => env.call('disableBoardAiGrouping'), /시트 메뉴에서만/);
  env.ctx.SpreadsheetApp.getUi = getUi;
  env.call('disableBoardAiGrouping');
  assert.equal(env.props.get('BOARD_AI_GROUPING'), '0');
  env.cache.delete('t:all');
  env.call('setBoard', r.token, q + ' (마지막 질문)', '', true, true);
  answers.forEach((t, k) => env.call('submitAnswer', t, 'revoke' + k));
  assert.equal(env.call('classifyBoard', r.token).askAi, true, 'revocation requires new consent');
  assert.equal(aiCalls.length, 2);
  // AI fails: nothing is saved, the reason comes back
  const bad = boot('r1-board.gs');
  bad.setup();
  bad.call('setBoard', bad.call('pin_', 'board'), q, '', true, false);
  answers.forEach((t, k) => bad.call('submitAnswer', t, 'b' + k));
  bad.ctx.UrlFetchApp.fetch = () => ({ getResponseCode: () => 500, getContentText: () => '{"error":"AI 서버 점검 중"}' });
  const bp = bad.call('classifyBoard', bad.call('pin_', 'board'));
  const br = bad.call('classifyBoard', bp.token, bp.confirmKey);
  assert.equal(br.ok, true);
  assert.match(br.autoErr, /AI가 묶음을 만들지 못했습니다.*점검/);
  assert.equal(br.auto, null);
  assert.deepEqual([bad.call('getBoard').cats, bad.sheet('설정').getRange('B2').getValue()], [[], ''], 'nothing saved');
  assert.ok(bad.rows('응답').every((x) => x[2] === ''), 'no row moved');
  // AI answers in a shape that has no groups: also nothing saved
  const odd = boot('r1-board.gs', { ai: () => '죄송합니다. 묶을 수 없습니다.' });
  odd.setup();
  odd.call('setBoard', odd.call('pin_', 'board'), q, '', true, false);
  answers.forEach((t, k) => odd.call('submitAnswer', t, 'c' + k));
  const op = odd.call('classifyBoard', odd.call('pin_', 'board'));
  const orr = odd.call('classifyBoard', op.token, op.confirmKey);
  assert.match(orr.autoErr, /묶음을 찾지 못/);
  assert.equal(odd.sheet('설정').getRange('B2').getValue(), '');
  // AI answers with a Markdown table and repeats a number on one line: names carry no '|', a repeated number counts once
  const tbl = boot('r1-board.gs', { ai: () => '| 묶음 이름 | 답 번호 |\n|---|---|\n| 시설 | 3, 3, 4, 5, 6 |\n| 환경 | 7, 8, 9 |\n| 잡담 | 1, 1 |' });
  tbl.setup();
  tbl.call('setBoard', tbl.call('pin_', 'board'), q, '', true, false);
  answers.forEach((t, k) => tbl.call('submitAnswer', t, 'd' + k));
  const tp = tbl.call('classifyBoard', tbl.call('pin_', 'board'));
  const tr = tbl.call('classifyBoard', tp.token, tp.confirmKey);
  assert.equal(tr.auto.kind, 'ai');
  assert.deepEqual(tr.board.cats, ['시설', '환경', '기타'], 'table cells read as names; a group of one repeated number is dropped');
  assert.ok(tbl.rows('응답').every((x) => !String(x[2]).includes('|')), 'no pipe in a saved name');
});

function pendingBatch() {
  const env = boot('r1-board.gs');
  env.setup();
  env.ctx.clusterAnswers_ = () => [];
  const pin = env.call('pin_', 'board'), sh = env.sheet('응답');
  env.call('setBoard', pin, '합성 질문', '', true, false);
  for (let i = 1; i <= 4; i++) sh.appendRow([new Date(), '합성 답 ' + i, '', '', '표시', i, '']);
  const pending = env.call('classifyBoard', pin);
  assert.ok(pending.confirmKey);
  return { env, pending };
}

test('r1 [분류]: turning off invalidates pending consent and prevents a batch even after JSON serialization', () => {
  const { env, pending } = pendingBatch();
  env.call('disableBoardAiGrouping');
  assert.match(env.log.alerts.at(-1), /이미 전송이 승인된 요청은 전송될 수 있으며, 전송된 요청은 취소할 수 없습니다/);
  const before = env.log.requests.length;
  const stale = env.call('classifyBoard', pending.token, pending.confirmKey);
  assert.equal(stale.ok, false);
  assert.match(stale.msg, /동의 확인이 만료/);
  assert.deepEqual([env.props.get('BOARD_AI_GROUPING'), env.log.requests.length], ['0', before]);
  const fresh = env.call('classifyBoard', pending.token);
  assert.ok(fresh.confirmKey && fresh.confirmKey !== pending.confirmKey);
  assert.equal(env.call('classifyBoard', pending.token, pending.confirmKey).ok, false);

  let serialized = false, fetched = 0;
  env.ctx.JSON = Object.assign(Object.create(JSON), { stringify(value) {
    const payload = JSON.stringify(value);
    if (payload.includes('[답 목록]')) {
      serialized = true;
      assert.equal(env.props.get('BOARD_AI_GROUPING'), '1');
      env.call('disableBoardAiGrouping');
    }
    return payload;
  } });
  env.ctx.UrlFetchApp.fetch = () => { fetched++; throw new Error('unexpected fetch after revocation'); };
  const revoked = env.call('classifyBoard', fresh.token, fresh.confirmKey);
  assert.equal(serialized, true);
  assert.equal(fetched, 0);
  assert.deepEqual([revoked.auto, revoked.aiOff, revoked.confirmKey, env.props.get('BOARD_AI_GROUPING'), env.log.requests.length], [null, true, undefined, '0', before]);
});

test('r1 [분류]: revocation during 429/503 backoff blocks retries even if a new generation is active', () => {
  for (const status of [429, 503]) {
    const { env, pending } = pendingBatch();
    let locked = false, fetched = 0, waited = 0;
    env.ctx.LockService.getScriptLock = () => ({
      waitLock() { assert.equal(locked, false); locked = true; },
      tryLock() { this.waitLock(); return true; },
      releaseLock() { assert.equal(locked, true); locked = false; },
    });
    env.ctx.UrlFetchApp.fetch = () => {
      assert.equal(locked, false, 'fetch does not hold the lock');
      fetched++;
      return { getResponseCode: () => status, getContentText: () => '{}' };
    };
    env.ctx.Utilities.sleep = () => {
      assert.equal(locked, false, 'backoff does not hold the lock');
      waited++;
      env.call('disableBoardAiGrouping');
      if (status === 503) env.props.set('BOARD_AI_GROUPING', '1'); // a different approval generation
    };
    const r = env.call('classifyBoard', pending.token, pending.confirmKey);
    assert.deepEqual([fetched, waited, r.auto, r.askAi, r.confirmKey, r.autoErr], [1, 1, null, true, undefined, '']);
  }
});

test('common: non-batch AI and Jev still retry 429/503 without batch consent', () => {
  for (const fn of ['ai_', 'jev_']) {
    const env = boot('r1-board.gs');
    env.props.set('BOARD_AI_GROUPING', '0');
    env.ctx.LockService.getScriptLock = () => assert.fail('no batch consent lock');
    const fetch = env.ctx.UrlFetchApp.fetch;
    let attempts = 0;
    env.ctx.UrlFetchApp.fetch = (...args) => {
      const status = [429, 503, 200][attempts++];
      return status === 200 ? fetch(...args) : { getResponseCode: () => status, getContentText: () => '{}' };
    };
    const result = fn === 'ai_' ? env.call(fn, '합성 질문') : env.call(fn, '합성 답', { a: { type: 'noul', instructions: '합성 질문' } });
    assert.ok(result);
    assert.deepEqual([attempts, env.log.sleeps], [3, 2]);
  }
});

test('r1 [분류]: consent lock timeout does not change state or consume a pending key', () => {
  const { env, pending } = pendingBatch();
  const msg = '다른 작업이 진행 중이라 동의 상태를 바꾸지 못했습니다. 잠시 후 다시 시도하십시오.';
  const lock = env.ctx.LockService.getScriptLock, before = env.log.requests.length;
  env.ctx.LockService.getScriptLock = () => ({ tryLock: () => false, releaseLock: () => assert.fail('no lock acquired') });
  const timedOut = env.call('classifyBoard', pending.token, pending.confirmKey);
  assert.deepEqual([timedOut.ok, timedOut.msg, timedOut.consentBusy, env.log.requests.length], [false, msg, true, before]);
  assert.equal(env.props.has('BOARD_AI_GROUPING'), false);
  assert.equal(env.cache.get('board-ai:' + pending.token + ':' + pending.confirmKey), '0');
  env.ctx.LockService.getScriptLock = lock;
  assert.equal(env.call('classifyBoard', pending.token, pending.confirmKey).ok, true);
  assert.equal(env.props.get('BOARD_AI_GROUPING'), '1');
  env.ctx.LockService.getScriptLock = () => ({ tryLock: () => false, releaseLock: () => assert.fail('no lock acquired') });
  assert.equal(env.call('disableBoardAiGrouping'), false);
  assert.deepEqual([env.props.get('BOARD_AI_GROUPING'), env.props.has('BOARD_AI_GENERATION'), env.log.alerts.at(-1)], ['1', false, msg]);
});

test('r1 [분류]: after the second range write fails, reload the partly saved sheet and never write a manual row', () => {
  for (const batch of [false, true]) {
    const env = boot('r1-board.gs');
    env.setup();
    const pin = env.call('pin_', 'board'), sh = env.sheet('응답');
    if (batch) {
      env.call('setBoard', pin, '합성 질문', '', true, false);
      env.ctx.clusterAnswers_ = () => [{ name: '부정행위', idx: [0, 1] }, { name: '업무 부담', idx: [2, 3] }];
    }
    ['부정행위', '합성 수동 답', '업무 부담', '부정행위', '업무 부담'].forEach((text, i) => {
      sh.appendRow([new Date(), text, i === 1 ? '평가 방법' : '', '', '표시', i + 1, '']);
    });
    env.cache.delete('board'); // direct fixture writes do not invalidate the board cached by setBoard
    const before = env.call('getBoard');
    assert.equal(before.items[0].c, batch ? '' : '기타');
    assert.ok(env.cache.has('board'));
    const getRange = sh.getRange.bind(sh), writes = [];
    sh.getRange = (...args) => {
      const range = getRange(...args), setValues = range.setValues.bind(range);
      range.setValues = (v) => {
        writes.push(args);
        if (writes.length === 2) throw new Error('두 번째 범위 저장 실패');
        return setValues(v);
      };
      return range;
    };
    const r = env.call('classifyBoard', pin);
    assert.deepEqual([r.ok, r.partial, r.auto, r.moved], [true, true, null, 0]);
    assert.deepEqual(writes, [[2, 3, 1, 5], [4, 3, 3, 5]]);
    assert.deepEqual(env.rows('응답').map((row) => row[2]), ['부정행위', '평가 방법', '', '', '']);
    assert.deepEqual(r.board, env.call('getBoard'));
    assert.deepEqual([r.board.rev, r.board.mix], [before.rev + 1, before.mix + 1]);
    if (!batch) assert.deepEqual(r.board.items.map((item) => item.c), ['부정행위', '평가 방법', '기타', '기타', '기타']);
  }
});

test('r1 [분류]: property failures after a sheet write still invalidate both caches and preserve the write result', () => {
  for (const [batch, property] of [[false, 'BOARD_MIX'], [true, 'BOARD_REV']]) {
    const env = boot('r1-board.gs');
    env.setup();
    const pin = env.call('pin_', 'board'), sh = env.sheet('응답');
    if (batch) {
      env.call('setBoard', pin, '합성 질문', '', true, false);
      env.ctx.clusterAnswers_ = () => [{ name: '부정행위', idx: [0, 1] }, { name: '업무 부담', idx: [2, 3] }];
    }
    ['부정행위', '부정행위', '업무 부담', '업무 부담'].forEach((text, i) => sh.appendRow([new Date(), text, '', '', '표시', i + 1, '']));
    env.cache.delete('board');
    const before = env.call('getBoard');
    env.call('getQuestion');
    assert.ok(env.cache.has('board') && env.cache.has('q'));
    const getProps = env.ctx.PropertiesService.getScriptProperties;
    env.ctx.PropertiesService.getScriptProperties = () => {
      const props = getProps(), set = props.setProperty;
      props.setProperty = (k, v) => {
        if (k === property) {
          assert.equal(env.rows('응답')[0][2], '부정행위', 'write preceded the failing property update');
          throw new Error('합성 속성 저장 실패');
        }
        set(k, v);
      };
      return props;
    };
    const r = env.call('classifyBoard', pin);
    assert.deepEqual([r.ok, r.partial, r.needsReload, batch ? r.auto.moved : r.moved], [true, undefined, true, 4]);
    assert.deepEqual(r.board.items.map((x) => x.c), ['부정행위', '부정행위', '업무 부담', '업무 부담']);
    assert.deepEqual([r.board.mix, r.board.rev], [before.mix + Number(property !== 'BOARD_MIX'), before.rev + Number(property !== 'BOARD_REV')]);
    assert.equal(env.cache.has('q'), false);
    assert.deepEqual(env.call('getBoard'), r.board);
    if (batch) assert.deepEqual(env.call('getQuestion').cats, ['부정행위', '업무 부담']);
  }
});

test('r1 [분류]: generation sends only the latest 80 eligible answers and masks contacts that straddle the 80-character cutoff', () => {
  let prompt = '';
  const env = boot('r1-board.gs', { ai: ({ messages }) => { prompt = messages.at(-1).content; return '첫째 | 1, 2\n둘째 | 79, 80'; } });
  env.setup();
  env.ctx.clusterAnswers_ = () => [];
  env.props.set('BOARD_AI_GROUPING', '1');
  const pin = env.call('pin_', 'board'), sh = env.sheet('응답');
  env.call('setBoard', pin, '어떤 점을 바꾸고 싶습니까?', '', true, false);
  const phone = ['+82', '10', '0'.repeat(4), '0'.repeat(4)].join(' ');
  for (let i = 82; i > 0; i--) sh.appendRow([new Date(), i === 82 ? '가'.repeat(74) + ' ' + phone + ' 제외된 뒤쪽' : '합성 답 ' + i, '', '', '표시', 1000000 + i, '']);
  sh.appendRow([new Date(), '보내지 않는 수동 분류', '교수의 묶음', '', '표시', 1000083, '']);
  const r = env.call('classifyBoard', pin);
  assert.equal(r.auto.kind, 'ai');
  const snippets = prompt.split('[답 목록]\n')[1].split('\n');
  assert.deepEqual([snippets.length, snippets[0], snippets[78], snippets[79]], [80, '1. 합성 답 3', '79. 합성 답 81', '80. ' + '가'.repeat(74) + ' *****']);
  assert.equal(sh.getRange(2, 3).getValue(), '둘째', 'AI numbers map back to the right rows after reordering');
  assert.equal(sh.getRange(84, 3).getValue(), '교수의 묶음');
});

test('r1 [분류]: a list changed while sorting, or an archive, writes nothing; Jev seen as 기타 at submit time is not asked again', () => {
  const env = boot('r1-board.gs', { jev: (st, q) => ({ ...defaultAnswers(q), ...(q.cat ? { cat: choice('other', 1) } : {}) }) });
  env.setup();
  const pin = env.call('pin_', 'board'), sh = env.sheet('응답');
  const r = env.call('submitAnswer', '아직 잘 모르겠어요 정말', 's1');
  assert.equal(r.cat, '기타');
  const seen = JSON.parse(env.props.get('CLASSIFY_SEEN'));
  assert.deepEqual([seen.key, seen.ids.length], ['부정행위|평가 방법|학습 효과 저하|업무 부담', 1], 'recorded inside the submit lock with the list Jev judged against');
  const calls = env.log.jevCalls.length;
  const c1 = env.call('classifyBoard', pin);
  assert.deepEqual([c1.jev, env.log.jevCalls.length - calls], [0, 0], 'not asked again for the same list');
  // a row that Jev failed on (empty group and empty trace) is retried, with the spam check
  sh.appendRow([new Date(), '옛 실패 행입니다', '', '', '표시', 1000009, '']);
  env.cache.delete('t:classify');
  const c2 = env.call('classifyBoard', pin);
  assert.equal(c2.jev, 1);
  assert.deepEqual(Object.keys(env.log.jevCalls.at(-1).questions), ['cat', 'bad'], 'an old failed row also gets the spam check');
  // the list changes while Jev is thinking: nothing is written
  let edit = null;
  const race = boot('r1-board.gs', { jev: (st, q) => { if (q.cat && edit) { edit.sheet('설정').getRange('B2').setValue('가, 나'); edit = null; } return { ...defaultAnswers(q), ...(q.cat ? { cat: choice('c0', 1) } : {}) }; } });
  race.setup();
  race.sheet('응답').appendRow([new Date(), '시험 때 몰래 쓸까 봐', '', '오류: x', '표시', 1000001, '']);
  edit = race;
  const rr = race.call('classifyBoard', race.call('pin_', 'board'));
  assert.deepEqual([rr.moved, rr.board.counts['기타']], [0, 1]);
  assert.match(rr.stopped, /바뀌어/);
  assert.deepEqual(race.rows('응답')[0].slice(2, 4), ['', '오류: x']);
  // Clearing G alone makes a filled C a manual entry, even during Jev's request.
  let manual;
  const pending = boot('r1-board.gs', { jev: (st, q) => {
    if (q.cat) manual.getRange(2, 7).setValue('');
    return { ...defaultAnswers(q), ...(q.cat ? { cat: choice('c0', 1) } : {}) };
  } });
  pending.setup();
  manual = pending.sheet('응답');
  manual.appendRow([new Date(), '어디에 넣을지 궁금합니다', '기타', 0.8, '표시', 1000001, '기타']);
  const edited = pending.call('classifyBoard', pending.call('pin_', 'board'));
  assert.deepEqual([edited.jev, edited.moved], [1, 0]);
  assert.deepEqual(pending.rows('응답')[0].slice(2, 7), ['기타', 0.8, '표시', 1000001, '']);
});

test('r1 strict Jev answers: only "other" is 기타; an unknown choice, a missing confidence or a vanished group leaves an error trace', () => {
  const run = (answer) => {
    const env = boot('r1-board.gs', { jev: (st, q) => ({ ...defaultAnswers(q), cat: answer }) });
    env.setup();
    env.call('submitAnswer', '시험 때 몰래 쓸까 봐 걱정입니다 정말로', 'z');
    return env.rows('응답')[0];
  };
  assert.deepEqual(run(choice('other', 0.9)).slice(2, 4), ['기타', 0.9]);
  assert.deepEqual(run(choice('c1', 0.9)).slice(2, 4), ['평가 방법', 0.9]);
  if (SOLO) return; // the Gemini stand-in rebuilds every verdict from the listed choices, so an unknown choice cannot reach jevCat_ there
  const odd = run({ type: 'choice', choice: 'weird', confidence: 0.9, probs: {} });
  assert.deepEqual([odd[2], /^오류: Jev 답 형식이 다릅니다/.test(String(odd[3]))], ['', true], 'an unknown choice is an error, not a silent 기타');
  const none = run({ type: 'choice', choice: 'c0', probs: {} });
  assert.deepEqual([none[2], /^오류: Jev 답 형식이 다릅니다.*confidence/.test(String(none[3]))], ['', true], 'a missing confidence is an error');
  const gone = run(choice('c9', 0.9));
  assert.equal(gone[2], '', 'a group number that does not exist is an error');
});

test('r1 errors are shown: setup checks the connection, getBoard reports the failed rows (never to students), [분류] says why it stopped', () => {
  if (!SOLO) {
    const bad = boot('r1-board.gs');
    bad.ctx.PropertiesService.getScriptProperties().setProperty('CLASS_CODE', 'WRONG');
    bad.call('onOpen');
    bad.call('setup');
    assert.ok(bad.sheet('응답'), 'setup goes on');
    assert.match(bad.log.alerts.at(-1), /준비 끝!.*판정 서버에 연결하지 못했습니다.*수업 코드/s);
  }
  const ok = boot('r1-board.gs');
  ok.setup();
  assert.ok(!/연결하지 못했습니다/.test(ok.log.alerts.at(-1)));
  ok.log.jevFail = 503;
  ok.call('submitAnswer', '판정이 안 될 때의 첫 답입니다', 'e1');
  ok.cache.clear();
  const b = ok.call('getBoard');
  assert.deepEqual([b.err.n, /503/.test(b.err.why), b.err.why.length <= 60], [1, true, true]);
  const fine = boot('r1-board.gs');
  fine.setup();
  assert.equal('err' in fine.call('getBoard'), false, 'no failed row, no err');
  assert.ok(!ok.call('doGet', { parameter: {} }).html.includes('getBoard'), 'the student page never reads getBoard');
  // [분류] stops after a partial run: the reason is in the result whatever r.jev is
  ok.log.jevFail = 0;
  const pin = ok.call('pin_', 'board');
  for (let k = 0; k < 3; k++) ok.sheet('응답').appendRow([new Date(), '시험 때 몰래 ' + k, '', '오류: x', '표시', 1000100 + k, '']);
  let n = 0;
  const orig = ok.ctx.jev_;
  ok.ctx.jev_ = (s, qs) => { if (++n === 2) throw new Error('Jev 점검 중'); return orig(s, qs); };
  const r = ok.call('classifyBoard', pin);
  assert.ok(r.jev >= 1 && /점검/.test(r.stopped), 'a partial failure still reports the reason');
});

// The big-screen browser script (editor, [분류] button, PIN path) needs a real DOM: it is checked in headless Chrome when Chrome is installed, skipped otherwise.
const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
test('r1 big screen in a real browser: [분류] authorizes with PIN and asks for AI consent without sending it until approved', () => {
  if (SOLO || !fs.existsSync(CHROME)) return;
  const env = boot('r1-board.gs');
  env.setup();
  // Use a fake join address and the local QR library so startup doesn't show an unrelated QR warning.
  const html = env.call('doGet', { parameter: { v: 'board', join: 'https://script.google.com/macros/s/TEST/exec' } }).html;
  const qr = fs.readFileSync(new URL('../vendor/qrcode/qrcode.js', import.meta.url), 'utf8');
  const pin = '2468';
  const board = { q: 'Q', cats: ['가', '나', '기타'], counts: { 가: 0, 나: 0, 기타: 1 }, open: true, refresh: 4, total: 1, round: 1, rev: 1, mix: 0, items: [{ i: 1, no: 1, h: '', t: '오류 시험 답', c: '기타' }], err: { n: 1, why: 'AI ONE 401: 수업 코드가 맞지 않습니다.' } };
  const stub = `<script>(function(){var BOARD=${JSON.stringify(board)},log=[],T='t'+'0'.repeat(32),confirmed=0;
    function run(ok,bad){return new Proxy({},{get:function(_,n){if(n==='withSuccessHandler')return function(f){return run(f,bad)};if(n==='withFailureHandler')return function(f){return run(ok,f)};
      return function(){var a=[].slice.call(arguments);log.push([n,a[0],a[1]]);var askAi=!a[1]&&log.filter(function(x){return x[0]==='classifyBoard'}).length>=3;
        var r=n==='getBoard'?BOARD:n==='classifyBoard'?{ok:true,token:T,grouped:true,moved:0,words:0,jev:0,left:1,more:0,stopped:'',auto:null,autoErr:'',autoNone:!askAi,confirmKey:askAi?'fixture-confirm':'',aiOff:!a[1],partial:!!a[1],board:BOARD}:{ok:false,msg:'x'};
        if(n==='classifyBoard'&&a[1]&&++confirmed===1)r={ok:false,token:T,consentBusy:true,msg:'다른 작업이 진행 중이라 동의 상태를 바꾸지 못했습니다. 잠시 후 다시 시도하십시오.'};
        setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(r)))},20)}}})}
    window.google={script:{run:run()}};
    function key(k,t){document.dispatchEvent(new KeyboardEvent('keydown',{key:k,code:t||''}))}
    function out(tag){var d=document.createElement('pre'),err=document.getElementById('ai-error');d.id='res-'+tag;d.textContent=JSON.stringify({log:log,edOn:document.getElementById('ed').className,toast:document.getElementById('toast').textContent,dialog:document.getElementById('ai-confirm').open,consentError:err.hidden?'':err.textContent});document.body.appendChild(d)}
    setTimeout(function(){out('e')},150);
    setTimeout(function(){key('ㅊ','KeyC')},300);
    setTimeout(function(){out('a');var ep=document.getElementById('ep');ep.value=${JSON.stringify(pin)};ep.dispatchEvent(new Event('input'))},700);
    setTimeout(function(){out('b');sessionStorage.removeItem('board-tok');document.getElementById('ed').className='edit';key('c','KeyC')},1100);
    setTimeout(function(){var ep=document.getElementById('ep');ep.value='5678';ep.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'}));out('c')},1500);
    setTimeout(function(){document.getElementById('ek').onclick();out('d')},1900);
    setTimeout(function(){out('prompt');document.getElementById('ai-no').click()},2100);
    setTimeout(function(){out('declined');key('c','KeyC')},2400);
    setTimeout(function(){document.getElementById('ai-yes').click()},2700);
    setTimeout(function(){out('busy');document.getElementById('ai-yes').click()},3000);
    setTimeout(function(){out('accepted')},3300);
  })();</script>`;
  const file = path.join(os.tmpdir(), 'board-pin-' + process.pid + '.html');
  fs.writeFileSync(file, html.replace('<body>', '<body>' + stub)
    .replace('<script src="https://kimcook-edu.pages.dev/vendor/qrcode/qrcode.js"></script>', () => '<script>' + qr + '</script>'));
  const dom = execFileSync(CHROME, ['--headless=new', '--disable-gpu', '--window-size=1600,900', '--virtual-time-budget=4000', '--dump-dom', 'file://' + file], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], maxBuffer: 1 << 26 });
  fs.unlinkSync(file);
  const res = (t) => JSON.parse(new RegExp('<pre id="res-' + t + '">([^<]*)</pre>').exec(dom)[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
  const e = res('e'), a = res('a'), b = res('b'), d = res('d');
  const sent = (r) => r.log.filter((x) => x[0] === 'classifyBoard').map((x) => x[1]);
  assert.match(e.toast, /판정 연결 오류 1개: AI ONE 401: 수업 코드가 맞지 않습니다\. · \[분류\]로 다시 나눌 수 있어요/, 'the big screen says why rows failed');
  assert.match(a.edOn, /\bon\b/, 'C (even typed as the Korean ㅊ) without a PIN opens the edit box');
  assert.deepEqual(sent(a), [], 'nothing is sent before a PIN');
  assert.deepEqual(sent(b), [pin], 'the fourth digit runs [분류] at once');
  assert.deepEqual(sent(d), [pin, '5678', 't' + '0'.repeat(32)], 'Enter in the PIN box runs it too; the box [분류] button then reuses the saved ticket, no PIN');
  const prompt = res('prompt'), declined = res('declined'), busy = res('busy'), accepted = res('accepted');
  assert.equal(prompt.dialog, true, 'the projector shows an in-page confirmation');
  assert.equal(declined.dialog, false);
  assert.match(declined.toast, /응답을 한꺼번에 보내 새 묶음을 만드는 기능이 꺼져 있습니다/);
  assert.deepEqual(sent(declined), sent(d), 'declining sends no follow-up request');
  assert.equal(busy.dialog, true);
  assert.equal(busy.consentError, '다른 작업이 진행 중이라 동의 상태를 바꾸지 못했습니다. 잠시 후 다시 시도하십시오.');
  assert.equal(accepted.dialog, false);
  assert.match(accepted.toast, /일부만 저장되었을 수 있습니다\. 화면을 다시 불러옵니다\./);
  assert.deepEqual(accepted.log.filter((x) => x[0] === 'classifyBoard').map((x) => x[2]), ['', '', '', '', 'fixture-confirm', 'fixture-confirm']);
});

test('web app address: never getUrl(); the board QR uses the saved /exec or the wizard join, nothing anonymous is saved', () => {
  const env = boot('r1-board.gs');
  env.setup();
  const A = 'https://script.google.com/macros/s/AKfycbAAA/exec', B = 'https://script.google.com/macros/s/AKfycbBBB/exec';
  const bootOf = (q) => JSON.parse(/var BOOT = (.*?);<\/script>/.exec(env.call('doGet', { parameter: { v: 'board', ...q } }).html)[1]);
  assert.deepEqual(bootOf({}), { join: '', warn: false }, 'no address known: no QR (never a guessed one)');
  assert.deepEqual(bootOf({ join: A }), { join: A, warn: false });
  assert.deepEqual(bootOf({ join: 'https://script.google.com/macros/s/AKfycbAAA/dev' }).join, '', '/dev only opens for editors');
  assert.deepEqual(bootOf({ join: 'https://evil.example/macros/s/X/exec' }).join, '');
  assert.ok(!env.call('doGet', { parameter: { v: 'board', join: A + '"</script><script>alert(1)</script>' } }).html.includes('alert(1)'), 'a crafted join never reaches the page');
  assert.equal(env.props.get('WEB_URL'), undefined, 'a visitor-supplied join is never saved');
  assert.equal(bootOf({ join: 'https://script.google.com/a/macros/bufs.ac.kr/s/AKfycbAAA/exec' }).join, A, 'school (Workspace) form becomes the form signed-out phones can open');
  assert.equal(bootOf({ join: 'https://script.google.com/a/bufs.ac.kr/macros/s/AKfycbAAA/exec' }).join, A, 'old school form too');
  env.log.prompts.push(A);
  env.call('showUrls');
  assert.equal(env.props.get('WEB_URL'), A, 'the menu asks once and keeps it');
  assert.ok(env.log.alerts.at(-1).includes(A + '?v=board') && /PIN: \d{4}/.test(env.log.alerts.at(-1)));
  assert.deepEqual(bootOf({ join: B }), { join: A, warn: true }, 'the saved address wins; a different deployment is flagged on the projector');
  assert.deepEqual(bootOf({ join: 'https://script.google.com/a/macros/bufs.ac.kr/s/AKfycbAAA/exec' }).warn, false, 'same deployment in another form is not a mismatch');
  env.ctx.SpreadsheetApp.getUi = () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); };
  assert.throws(() => env.call('setWebUrl'), /시트 메뉴에서만/, 'web callers cannot change it');
  for (const f of ['r1-board.gs', 'r2-assistant.gs', 'r3-tutor.gs', 'r4-future.gs', 'r6-survey.gs']) assert.ok(!/getService\(\)\.getUrl\(/.test(readGs(f).replace(/\/\/.*$/gm, '')), f + ' does not use getUrl()');
});

test('every recipe answers a visit to its web address: a tool without a web page says so instead of "doGet not found"', () => {
  for (const f of ['r1-board', 'r2-assistant', 'r3-tutor', 'r4-future', 'r5-boardphoto', 'r6-survey']) {
    const env = boot(f + '.gs');
    assert.equal(typeof env.ctx.doGet, 'function', f + ' has doGet');
    const out = env.ctx.doGet({ parameter: {}, parameters: {} });
    assert.ok(out && out.html, f + ' doGet returns a page');
  }
  for (const f of ['r2-assistant', 'r4-future']) {
    const fresh = boot(f + '.gs'); // a student opens the address before the professor ran the first preparation
    assert.ok(fresh.ctx.doGet({ parameter: {} }).html.includes('아직 준비 중입니다'), f + ' shows a "not ready" page, not a script error');
  }
  const r5 = boot('r5-boardphoto.gs');
  r5.setup();
  const html = r5.ctx.doGet({ parameter: {} }).html;
  assert.ok(html.includes('웹 주소로 열지 않습니다'));
  assert.ok(!/drive\.google\.com/.test(html), 'the public page does not reveal the Drive folder');
});

test('design guard: pages and mails stay flat and in the lab style (no gradients, glow, blur, big radii, emoji); mail text is escaped', () => {
  const env = boot('r1-board.gs');
  env.setup();
  const pages = [env.ctx.PAGE_CSS, env.ctx.BOARD_CSS, env.ctx.ASK_CSS, env.ctx.STAGE_CSS, env.ctx.TOAST_CSS];
  for (const f of ['r2-assistant', 'r3-tutor', 'r4-future']) pages.push(readGs(f + '.gs'));
  for (const css of pages) {
    assert.ok(!/gradient|box-shadow|backdrop|blur\(|filter:|color-mix/i.test(css), 'no decorative effects');
    for (const m of css.matchAll(/border-radius:\s*([\d.]+)px/g)) assert.ok(Number(m[1]) <= 4, 'radius ' + m[1]);
    assert.ok(!/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(css), 'no emoji');
  }
  env.call('mail_', 'a@b.co', '제목 <b>', '[소제목]\n학생 글: <script>alert(1)</script> https://x.test/a?b=1&c=2');
  const m = env.log.mails.at(-1);
  assert.ok(!m.html.includes('<script>') && m.html.includes('&lt;script&gt;'), 'user text is escaped');
  assert.ok(m.html.includes('href="https://x.test/a?b=1&amp;c=2"'));
  assert.equal(m.body, '[소제목]\n학생 글: <script>alert(1)</script> https://x.test/a?b=1&c=2', 'the plain-text body is unchanged');
});

test('every recipe: menu items exist, the shared helpers are not redefined, pages carry no template literals', () => {
  for (const f of ['r1-board', 'r2-assistant', 'r3-tutor', 'r4-future', 'r5-boardphoto', 'r6-survey']) {
    const src = readGs(f + '.gs');
    const env = boot(f + '.gs');
    env.call('onOpen');
    assert.equal(env.log.menus.length, 1, f);
    for (const [label, fn] of env.log.menus[0].items) assert.equal(typeof env.ctx[fn], 'function', `${f}: ${label} -> ${fn}`);
    for (const shared of ['post_', 'ai_', 'jev_', 'ss_', 'sheet_', 'throttle_', 'page_', 'setClassCode', 'setApiKey', 'gemini_', 'ensureKey_']) {
      assert.ok(!new RegExp('function ' + shared + '\\b').test(src), `${f} redefines ${shared}`);
    }
    assert.ok(!/api[_-]?key|sk-[A-Za-z0-9]{10}|AIza[\w-]{20}/i.test(src.replace(/setApiKey/g, '')), `${f} must not contain keys`);
    assert.ok((SOLO ? /AI 키 넣기|setApiKey/ : /수업 코드 입력|setClassCode/).test(src), f);
  }
});

test('web pages: each browser script runs against the real server functions', () => {
  // recipe 1: student page sends an answer, board page draws columns
  const r1 = boot('r1-board.gs');
  r1.setup();
  const ask = runPage(r1, r1.ctx.ASK_HTML);
  assert.deepEqual(ask.calls, ['getQuestion']);
  assert.equal(ask.el('#q').textContent, 'AI를 수업에 쓰는 것에서 가장 걱정되는 점은?');
  const toastOf = (pg) => pg.texts('#toast').join(' ');
  ask.click('#b');
  assert.deepEqual(ask.calls, ['getQuestion'], 'an empty answer is not sent');
  assert.ok(toastOf(ask).startsWith('실패 생각 보내기를 못했어요.') && ask.el('#toast').className.includes('bad'), 'a failed button shows the red toast');
  assert.equal(ask.el('#toast').children.at(-1).textContent, '닫기', 'the failure toast has a close button');
  ask.el('#toast').children.at(-1).onclick();
  assert.ok(!ask.el('#toast').className.includes('on'), 'closing hides it');
  const chipBox = ask.el('#chips');
  assert.equal(chipBox.hidden, false);
  assert.deepEqual(chipBox.children.slice(1).map((c) => c.textContent), ['부정행위', '평가 방법', '학습 효과 저하', '업무 부담', '기타']);
  const chip = (name) => chipBox.children.find((c) => c.textContent === name);
  chip('평가 방법').onclick();
  assert.ok(chip('평가 방법').className.includes('on'));
  chip('평가 방법').onclick();
  assert.ok(!chip('평가 방법').className.includes('on'), 'a second press lets go');
  chip('부정행위').onclick();
  ask.set('#t', '시험 때 몰래 쓸까 봐 걱정입니다');
  ask.click('#b');
  assert.deepEqual(ask.calls, ['getQuestion', 'submitAnswer']);
  assert.equal(r1.rows('응답')[0][3], '직접', 'the chosen chip travelled with the answer');
  assert.ok(!chip('부정행위').className.includes('on'), 'the chip is let go after sending');
  assert.equal(toastOf(ask), '완료 생각을 보냈어요 · No.001 · 「부정행위」 쪽으로 모였어요.', 'the student sees the number the big screen shows');
  assert.equal(ask.el('#toast').className, 'toast on run');
  const hide = ask.timers.find((t) => t[1] === 3000);
  assert.ok(hide, 'a success toast hides itself after three seconds');
  hide[0]();
  assert.equal(ask.el('#toast').className, 'toast');
  assert.equal(ask.el('#t').value, '');
  assert.equal(ask.el('#mine').hidden, false);
  assert.equal(ask.el('#ml').children.length, 1);
  assert.match(ask.el('#ml').children[0].textContent, /^001시험 때 몰래 쓸까 봐 걱정입니다→ 부정행위 · \d\d:\d\d$/, 'my sent thought stays as a numbered slip with its group');
  r1.cache.clear();
  const boardHtml = (q) => r1.call('doGet', { parameter: { v: 'board', ...q } }).html.match(/<body[^>]*>([\s\S]*)<\/body>/)[1].replace(r1.ctx.TOAST_JS, '');
  const board = runPage(r1, boardHtml({ join: 'https://script.google.com/macros/s/AKfycbTEST/exec' }));
  assert.deepEqual(board.calls, ['getBoard']);
  assert.equal(board.el('#c').textContent, '1');
  assert.equal(board.el('#g').children.length, 5, 'four groups plus the catch-all');
  assert.equal(board.el('#labs').children.length, 5, 'one label per group');
  const counts = () => board.el('#g').children.map((c) => c.children[1].textContent);
  const slips = (k) => board.el('#g').children[k].children[2].children.filter((x) => x.className.includes('slip')).map((x) => x.children[0].textContent);
  assert.deepEqual(counts(), ['1', '0', '0', '0', '0']);
  assert.deepEqual(slips(0), ['시험 때 몰래 쓸까 봐 걱정입니다'], 'each answer is one slip in its group');
  assert.equal(board.el('#g').children[0].children[2].children.find((x) => x.className.includes('slip')).children[1].textContent, '001', 'each slip carries its number');
  assert.equal(board.el('#labs').children[0].className.includes('lead'), true, 'the single largest group is marked');
  assert.equal(board.el('#wait').hidden, true);
  // the next poll adds only the new answer, on top of its group
  ask.set('#t', '과제를 AI가 대신 써 오면 구분이 안 돼요');
  ask.click('#b');
  r1.cache.clear();
  board.timers.shift()[0]();
  assert.deepEqual(slips(0), ['과제를 AI가 대신 써 오면 구분이 안 돼요', '시험 때 몰래 쓸까 봐 걱정입니다'], 'newest on top');
  assert.deepEqual(counts(), ['2', '0', '0', '0', '0']);
  assert.equal(board.el('#c').textContent, '2');
  const waiting = boot('r1-board.gs');
  waiting.setup();
  const w = runPage(waiting, waiting.ctx.BOARD_HTML);
  assert.ok(!w.el('#qr').innerHTML, 'no trusted address: no QR at all');
  assert.ok(w.texts('#toast').join(' ').includes('참여 QR을 만들 주소가 없어요'), '…and the projector says what to do');
  assert.equal(w.el('#wait').hidden, false, 'an empty board shows the waiting line');
  assert.equal(w.el('#labs').children.length, 5, 'the groups are already visible before the first answer');
  assert.equal(w.el('#c').textContent, '0');
  const doc = r1.call('doGet', { parameter: { v: 'board' } });
  assert.equal(doc.xframe, 'ALLOWALL', 'the page can be embedded in PowerPoint web viewers');
  assert.ok(doc.html.includes('qrcode.js') && !doc.html.includes('<main>'), 'raw board has no narrow main wrapper');
  assert.ok(r1.call('doGet', { parameter: {} }).html.includes('<main class="ask">'));
  assert.equal(board.timers.length, 1, 'the board keeps polling');
  // recipe 2
  const r2 = boot('r2-assistant.gs', { jev: (s, q) => ({ ...defaultAnswers(q), answerable: noul(0.95) }), ai: () => '바로 답하면 이렇게 됩니다' });
  r2.setup();
  r2.sheet('자료').appendRow(['과제는 매주 금요일까지']);
  const p2 = runPage(r2, r2.ctx.ASK_HTML);
  p2.set('#t', '과제 마감이 언제예요?');
  p2.click('#b');
  assert.equal(p2.el('#m').children[0].textContent, '바로 답하면 이렇게 됩니다');
  p2.set('#i', 'ZZZZZZ');
  p2.click('#k');
  assert.ok(p2.texts('#toast').join(' ').includes('답변 확인을 못했어요. 해당하는 질문 번호가 없습니다') && p2.el('#toast').className.includes('bad'));
  // recipe 3
  const r3 = boot('r3-tutor.gs', { jev: (s, q) => ({ ...defaultAnswers(q), move: choice('wrap') }), ai: () => '잘했어요. 이제 한 문단으로 정리해 보세요.' });
  r3.setup();
  const p3 = runPage(r3, r3.ctx.TUTOR_HTML);
  assert.equal(p3.el('#tp').textContent, '광합성: 식물은 빛으로 무엇을 만들까');
  p3.click('#go');
  p3.set('#al', '파랑고래');
  p3.set('#t', '빛을 먹는다고 생각해요');
  p3.click('#b');
  assert.equal(p3.el('#fin').hidden, false, 'a wrap move opens the summary box');
  p3.set('#ft', '식물은 물과 이산화탄소를 빛으로 포도당과 산소로 바꾼다.');
  p3.click('#sub');
  assert.equal(r3.rows('제출').length, 1);
  assert.ok(p3.texts('#m').some((t) => t.includes('제출했습니다')));
  // recipe 4
  const r4 = boot('r4-future.gs', { jev: (s, q) => ({ ...defaultAnswers(q), growth: score(1, ['a', 'b', 'c', 'd']) }), ai: () => '변화가 보입니다' });
  r4.setup();
  const p4 = runPage(r4, r4.ctx.WRITE_HTML);
  assert.match(p4.el('#dt').value, /^\d{4}-\d{2}-\d{2}$/);
  const tomorrow = r4.ctx.Utilities.formatDate(new Date(Date.now() + 86400000), 'x', 'yyyy-MM-dd');
  for (const [k, v] of [['#al', '파랑고래'], ['#em', 'a@b.co'], ['#go', '발표를 두려워하지 않기'], ['#wo', ''], ['#dt', tomorrow]]) p4.set(k, v);
  p4.click('#b');
  assert.ok(p4.el('#m').children[0].textContent.includes('도착합니다'));
  r4.props.set('WEB_URL', 'https://script.google.com/macros/s/AKfycbTEST/exec');
  r4.call('demoDeliverAll');
  const id = r4.rows('편지')[0][1];
  const p4r = runPage(r4, r4.call('doGet', { parameter: { r: id } }).html);
  assert.deepEqual(p4r.calls, ['getLetter']);
  assert.equal(p4r.el('#ref').hidden, false);
  p4r.set('#t', '학기를 돌아보니 발표가 조금 편해졌습니다.');
  p4r.click('#b');
  assert.ok(p4r.texts('#m').some((t) => t.includes('변화가 보입니다')));
  // a page that calls a missing server function must fail loudly
  assert.throws(() => runPage(r1, '<script>google.script.run.nope()</script>'), /does not exist/);
});

let failed = 0;
for (const [name, fn] of tests) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (e) {
    failed++;
    console.log(`FAIL ${name}\n     ${e && e.stack ? e.stack.split('\n').slice(0, 6).join('\n     ') : e}`);
  }
}
console.log(failed ? `${failed} of ${tests.length} failed` : `all ${tests.length} passed`);
process.exitCode = failed ? 1 : 0;
