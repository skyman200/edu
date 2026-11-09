// vibe/*.html 예시 앱 + vibe/patch.js 를 한 파일(vibe-kit.js)로 묶고, studio.src.html 에 끼워 넣어 혼자 쓰는 studio.html 을 만듭니다.
// 사용: node vibe/build.mjs  (deploy.sh 가 먼저 실행합니다)
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(root, p), 'utf8');

// 위에서부터 보여 줍니다. title/desc 는 학생 갤러리 카드에도 쓰입니다.
export const SAMPLES = [
  { id: 'order', name: '발표 순서 뽑기', tag: '진행', desc: '번호를 중복 없이 한 명씩 뽑아 크게 보여 줍니다', ask: '글자를 더 크게 하고, 뽑힌 번호가 나올 때 짧은 소리가 나게 해 줘' },
  { id: 'voice-card', name: '발음 듣기 단어 카드', tag: '외국어', desc: '카드를 넘기며 영어·일본어·중국어·스페인어 발음을 듣습니다', ask: '카드를 넘기면 자동으로 발음이 나오게 해 줘' },
  { id: 'dictation', name: '듣고 받아쓰기', tag: '외국어', desc: '문장을 듣고 받아 적으면 틀린 단어에 밑줄을 그어 줍니다', ask: '문장을 일본어로 바꿔 줘' },
  { id: 'roleplay', name: '상황 역할극 카드', tag: '외국어', desc: '상황 카드를 뽑아 역할과 꼭 써야 할 표현을 보여 주고 3분을 잽니다', ask: '타이머를 5분으로 바꿔 줘' },
  { id: 'match', name: '개념 짝 맞추기', tag: '복습', desc: '용어 카드와 설명 카드를 짝지으며 복습합니다', ask: '용어를 내 과목 용어로 바꿔 줘' },
  { id: 'quiz', name: '복습 퀴즈', tag: '복습', desc: '객관식을 한 문제씩 풀고 바로 해설을 봅니다', ask: '문제를 내 과목에서 8개로 바꿔 줘' },
  { id: 'teams', name: '조 편성기', tag: '진행', desc: '번호를 정한 조 수로 고르게 섞어 나눕니다', ask: '조장 표시를 빼 줘' },
  { id: 'debate', name: '토론 타이머', tag: '진행', desc: '찬성·반대 발언 시간을 차례로 재고 크게 보여 줍니다', ask: '발언 시간을 3분으로 바꿔 줘' },
  { id: 'rubric', name: '루브릭 점수 계산기', tag: '평가', desc: '기준마다 수준을 고르면 총점과 피드백 문장을 만들어 줍니다', ask: '기준 이름을 내 과제에 맞게 바꿔 줘' },
];

const samples = SAMPLES.map((s) => ({ ...s, code: read(`vibe/${s.id}.html`).trim() }));
const patch = read('vibe/patch.js').replace(/\nif \(typeof module[^\n]*\n?$/, '\n');
const kit = `/* 자동 생성: node vibe/build.mjs . vibe/*.html 과 vibe/patch.js 를 고친 뒤 다시 만드세요. */\n${patch}\nvar VIBE_SAMPLES = ${JSON.stringify(samples)};\n`;
writeFileSync(join(root, 'vibe-kit.js'), kit);
const page = read('studio.src.html');
if (!page.includes('/*VIBEKIT*/')) throw new Error('studio.src.html 에 /*VIBEKIT*/ 표시가 없습니다');
writeFileSync(join(root, 'studio.html'), page.replace('/*VIBEKIT*/', () => kit.replace(/<\/script/gi, '<\\/script')));
console.log(`vibe-kit.js ${kit.length}자, 예시 ${samples.length}개`);
