// Model router for /api/ai: picks a Claude model per exercise. Pure functions, no I/O.
// Every label ends in "5", so the Korean particles 가/로 used in fallbackWhy() stay correct.

export const MODEL_LABELS = {
  'claude-opus-5-5': 'Claude Opus 5.5',
  'claude-sonnet-5-5': 'Claude Sonnet 5.5',
  'claude-haiku-5-5': 'Claude Haiku 5.5',
  'claude-sonnet-5': 'Claude Sonnet 5',
  'claude-haiku-4-5-20251001': 'Claude Haiku 4.5',
};

export const TIERS = {
  fast: { model: 'claude-haiku-5-5', fallbacks: ['claude-haiku-4-5-20251001', 'claude-sonnet-5-5'] },
  balanced: { model: 'claude-sonnet-5-5', fallbacks: ['claude-sonnet-5', 'claude-haiku-4-5-20251001'] },
  strong: { model: 'claude-opus-5-5', fallbacks: ['claude-sonnet-5-5', 'claude-sonnet-5'] },
};

const FAST = '빠르고 가벼운 Claude Haiku 5.5를 골랐습니다';
const BAL = '균형형 Claude Sonnet 5.5를 골랐습니다';
const STRONG = '가장 강한 Claude Opus 5.5를 골랐습니다';

// [tier, maxTokens, why]
export const EXERCISES = {
  's1-1': ['fast', 1500, `짧은 요약과 질문이라 ${FAST}`],
  's1-2': ['balanced', 4000, `블룸 분류에 맞춰 목표 문장을 정확히 다듬어야 해서 ${BAL}`],
  's1-3': ['balanced', 6000, `표와 긴 계획을 정확히 짜야 해서 ${BAL}`],
  's1-4': ['balanced', 4000, `정답과 보기가 틀리면 안 되는 퀴즈 표라 ${BAL}`],
  's2-1': ['balanced', 6000, `목표, 평가, 활동을 한 흐름으로 길게 짜야 해서 ${BAL}`],
  's2-2': ['balanced', 4000, `학생이 바로 쓸 활동지를 짜임새 있게 만들어야 해서 ${BAL}`],
  's2-3': ['balanced', 4000, `기준, 수준, 배점이 맞물리는 루브릭 표라 ${BAL}`],
  's2-4': ['balanced', 4000, `학생 답안을 읽고 존중하는 말투로 피드백해야 해서 ${BAL}`],
  's2-5': ['fast', 1500, `짧은 원칙 문서라 ${FAST}`],
  's3-1': ['balanced', 4000, `레퍼런스를 고르는 기준을 따져 봐야 해서 ${BAL}`],
  's3-2': ['balanced', 4000, `느낌을 숫자와 규칙으로 정확히 옮겨야 해서 ${BAL}`],
  's3-3': ['balanced', 6000, `명세를 지키며 여러 장의 슬라이드 문구를 길게 써야 해서 ${BAL}`],
  's3-4': ['balanced', 4000, `어색한 문장과 디자인 신호를 꼼꼼히 찾아야 해서 ${BAL}`],
  's3-5': ['balanced', 4000, `슬라이드 사이 개체 이동을 단계별로 설명해야 해서 ${BAL}`],
  's4-1': ['balanced', 8000, `퀴즈 표를 구글 폼으로 바꾸는 Apps Script를 한 번에 돌아가게 써야 해서 ${BAL}`],
  's4-2': ['balanced', 8000, `학생별 메일을 보내는 Apps Script를 실수 없이 써야 해서 ${BAL}`],
  's4-3': ['balanced', 8000, `미제출자를 골라 내는 Apps Script를 한 번에 돌아가게 써야 해서 ${BAL}`],
  's4-4': ['strong', 12000, `코드 한 파일을 끝까지 틀리지 않게 써야 해서 ${STRONG}`],
  vibe: ['strong', 12000, `코드 한 파일을 끝까지 틀리지 않게 써야 해서 ${STRONG}`],
  'p-1': ['balanced', 4000, `내 강의자료를 읽고 학생용 가이드로 정리해야 해서 ${BAL}`],
  'p-2': ['balanced', 4000, `같은 목표로 난이도별 과제를 나눠 짜야 해서 ${BAL}`],
  'p-3': ['balanced', 4000, `정답 오류와 애매한 보기를 꼼꼼히 찾아야 해서 ${BAL}`],
  'p-4': ['balanced', 4000, `주관식 의견을 묶어 개선계획으로 정리해야 해서 ${BAL}`],
};

const P_DEFAULT = ['balanced', 4000, `자료를 읽고 정리하는 실습이라 ${BAL}`];
const CHAT_FAST = ['fast', 1500, `짧은 질문이라 ${FAST}`];
const CHAT_HEAVY = ['balanced', 4000, `표, 코드, 문서를 만드는 질문이라 ${BAL}`];
const CHAT_LONG = ['balanced', 4000, `긴 질문이라 꼼꼼히 읽도록 ${BAL}`];

export const CHAT_SHORT_LIMIT = 300;

// Table, code, script, slide or plan requests go to the balanced model.
// A bare 표 counts only as its own word (표로, 표를, 표 형식...), so 목표/발표/대표/표현 stay short questions.
const HEAVY_RE = new RegExp(
  [
    '```',
    '\\|\\s*:?-{3,}',
    '\\b(code|coding|table|html|css|javascript|python|script|slides?|ppt|apps?)\\b',
    '코드|코딩|프로그램|스크립트|슬라이드|계획서|웹앱',
    '시간표|일정표|비교표|평가표|점검표|진도표|성적표|계획표|도표',
    '(?<![가-힣])표(?=$|[^가-힣]|로|를|는|가|에|와|만|형식|형태)',
  ].join('|'),
  'i',
);

export function normalizeEx(ex) {
  if (typeof ex !== 'string') return 'chat';
  const id = ex.trim().toLowerCase();
  if (id.length === 0 || id.length > 16 || !/^[a-z0-9-]+$/.test(id)) return 'chat';
  if (Object.hasOwn(EXERCISES, id) || /^p-[a-z0-9-]+$/.test(id)) return id;
  return 'chat';
}

export function isHeavyChat(text) {
  return typeof text === 'string' && (text.length >= CHAT_SHORT_LIMIT || HEAVY_RE.test(text));
}

function build(id, [tier, maxTokens, why]) {
  const t = TIERS[tier];
  return { ex: id, model: t.model, fallbacks: [...t.fallbacks], maxTokens, tier, label: MODEL_LABELS[t.model], why };
}

// Deterministic per exercise: follow-up turns get the same model as turn 1.
// "chat" classifies on the first user message only, so a conversation never switches models.
export function pickModel({ ex, messages } = {}) {
  const id = normalizeEx(ex);
  if (Object.hasOwn(EXERCISES, id)) return build(id, EXERCISES[id]);
  if (id.startsWith('p-')) return build(id, P_DEFAULT);
  const first = Array.isArray(messages) ? messages.find((m) => m && m.role === 'user' && typeof m.content === 'string') : null;
  const text = first ? first.content : '';
  if (!isHeavyChat(text)) return build('chat', CHAT_FAST);
  return build('chat', text.length >= CHAT_SHORT_LIMIT ? CHAT_LONG : CHAT_HEAVY);
}

export function labelOf(model) {
  return MODEL_LABELS[model] || model;
}

export function fallbackWhy(route, usedModel) {
  if (usedModel === route.model) return route.why;
  return `${route.label}가 지금 바빠서 ${labelOf(usedModel)}로 대신 답합니다. ${route.why}`;
}

const BASE_SYSTEM = [
  '당신은 대학 교수자의 수업 설계를 돕는 조교입니다.',
  '사용자가 다른 언어를 요청하지 않으면 한국어로 답합니다.',
  '군더더기 없이 바로 수업에 쓸 수 있게 간결하게 답합니다.',
  '표가 필요하면 Markdown 표로 씁니다.',
  '웹을 검색하거나 링크를 열 수 없으니 그런 척하지 않습니다.',
  '출처, 논문, URL을 지어내지 말고 확실하지 않으면 "확인 필요"라고 적습니다.',
  'API 키나 비밀번호 같은 비밀 정보는 출력하지 않습니다.',
  '코드를 요청받으면 완성된 파일 하나를 코드 블록 하나에 담아 줍니다.',
].join(' ');

const HTML_APP_SYSTEM = [
  '이번 요청은 수업용 웹앱입니다.',
  '완성된 HTML 문서 하나를 ```html 코드 블록 하나에만 담고, CSS와 JavaScript도 그 파일 안에 넣습니다.',
  '외부 네트워크 호출, CDN, 외부 글꼴은 쓰지 않고 로그인 기능도 넣지 않습니다.',
  '글자는 크고 읽기 쉽게 만듭니다. 코드 블록 밖 설명은 세 줄 이내로 씁니다.',
].join(' ');

const APPS_SCRIPT_SYSTEM = [
  '이번 요청은 Google Apps Script 코드입니다.',
  '완성된 코드를 ```javascript 코드 블록 하나에 담고, 맨 위에 설정 방법을 한국어 주석 세 줄로 적습니다.',
].join(' ');

export function systemPromptFor(ex) {
  const id = normalizeEx(ex);
  if (id === 'vibe' || id === 's4-4') return `${BASE_SYSTEM}\n\n${HTML_APP_SYSTEM}`;
  if (id === 's4-1' || id === 's4-2' || id === 's4-3') return `${BASE_SYSTEM}\n\n${APPS_SCRIPT_SYSTEM}`;
  return BASE_SYSTEM;
}
