// Builds ../gas-recipes.js (window.GAS_RECIPES) from the .gs sources and the metadata below: `node gas/build.mjs`.
// deploy.sh runs it, and gas.test.mjs fails when the checked-in file is out of date.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { EXTRA, PROMPT_RULE, TROUBLE, SOLO_TROUBLE } from './story.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const read = (f) => fs.readFileSync(path.join(here, f), 'utf8').replace(/\r\n/g, '\n').trimEnd() + '\n';

const Q = {
  noul: (instructions, criteria) => ({ type: 'noul', instructions, ...(criteria ? { criteria } : {}) }),
  choice: (instructions, criteria) => ({ type: 'choice', instructions, criteria }),
  score: (instructions, criteria) => ({ type: 'score', instructions, criteria }),
};

export const RECIPES = [
  {
    id: 'r1', web: { deploy: true, ppt: { suffix: '?v=board', join: true, title: '교실 반응판' }, links: [{ label: '참석자용 주소 (QR로 띄우세요)', suffix: '' }, { label: '강의장 큰 화면 주소 (프로젝터에 띄우세요. 이 주소로 열어야 화면의 참여 QR이 맞게 나옵니다)', suffix: '?v=board', join: true }], after: ['큰 화면 주소를 열고 F(전체화면)를 누르세요. 질문과 묶음은 그 화면에서 E를 눌러 바로 바꿉니다. 교수 PIN 4자리는 start 실행 때 알림창과 시트 메뉴 반응판 → 주소와 교수 PIN 보기에 있습니다. 한 번 맞히면 그 창을 닫을 때까지 다시 묻지 않습니다.', '수업 중간에 새 질문으로 넘어갈 때는 E 칸에서 "지금까지 응답은 보관"을 켜고 바꾸면 화면이 비워집니다. 지난 응답은 시트에 남습니다.'] }, file: 'r1-board.gs', name: '실시간 교실 반응판', tag: '참석자 폰 → 큰 화면',
    short: '질문을 던지면 참석자가 폰으로 한 줄씩 답하고, 큰 화면에서 답 하나가 한 칸씩 묶음 막대로 쌓입니다. 질문은 띄운 화면에서 바로 바꿉니다.',
    jev: ['답이 어느 묶음에 속하는지 고르기(학생이 폰에서 묶음을 직접 고르면 건너뜀)', '광고·장난·욕설이면 화면에서 숨기기'],
    steps: [
      '새 구글 시트를 만들고 확장 프로그램 → Apps Script를 엽니다.',
      '아래 "공통 도우미"와 이 레시피 코드를 한 파일에 위아래로 붙여 넣고 저장합니다.',
      '시트로 돌아가 새로고침 → [반응판] 메뉴 → 처음 준비 → 수업 코드 입력.',
      '배포 → 새 배포 → 웹 앱(실행 사용자: 나, 액세스: 모든 사용자). 웹 도구에 /exec 주소를 붙여 넣어 참석자용 QR과 강의장 큰 화면 주소를 받습니다. 시트 메뉴 반응판 → 웹 앱 주소 넣기에도 같은 주소를 넣어 두면 메뉴의 주소 보기와 큰 화면 QR이 그 주소를 씁니다.',
      '[설정] 시트에서 질문과 묶음 이름을 내 수업에 맞게 바꿉니다.',
    ],
    examples: [
      "질문을 '이번 학기 가장 기대되는 AI 활용'으로 바꾸고, 묶음을 수업 준비, 채점, 학생 상담, 연구로 바꿔 줘.",
      "학생이 답을 보내면 '○○ 쪽으로 모였어요' 대신 응원 한 줄을 AI가 써서 보여 줘.",
      '강의장 화면에서 묶음 제목 옆에 개수를 막대로 보여 주고 가장 큰 묶음을 맨 왼쪽에 놓아 줘.',
      '같은 기기에서 같은 답을 두 번 보내면 막아 줘.',
    ],
    jevPreset: {
      label: '반응판: 답을 묶음으로 나누기',
      state: { 질문: 'AI를 수업에 쓰는 것에서 가장 걱정되는 점은?', 응답: '시험 때 몰래 쓸까 봐 걱정입니다' },
      questions: {
        cat: Q.choice('교수의 질문에 대한 응답이 어느 묶음에 속하는지 고르세요.', { 부정행위: '부정행위', 평가방법: '평가 방법', 학습효과: '학습 효과 저하', 업무부담: '업무 부담', other: '위 항목 어디에도 맞지 않는 내용' }),
        bad: Q.noul('응답이 질문과 무관한 광고, 욕설, 장난, 개인정보 노출인가?'),
      },
    },
  },
  {
    id: 'r2', web: { deploy: true, links: [{ label: '학생용 주소 (LMS나 QR로 나눠 주세요)', suffix: '' }], after: ['처음 준비(start)가 내 드라이브에 만든 "수업 조교 자료" 폴더에 강의계획서, 평가 기준, 공지 파일을 넣으세요(구글 문서, PDF, 워드, 한글 hwpx, 파워포인트, 텍스트, 사진). 붙여 넣을 필요가 없습니다. 메뉴 [자료 폴더 보기]에서 조교가 읽은 파일과 읽지 못한 파일을 확인합니다. 이 내용에만 근거해 답합니다.'] }, file: 'r2-assistant.gs', name: '24시간 수업 조교', tag: '학생 질문 → 자동 답변 또는 교수님께',
    short: '강의계획서를 읽고 학생 질문에 답하고, 자료에 없거나 개인 사정이면 교수님께 넘깁니다. 매주 가장 헷갈린 것 TOP5를 메일로 받습니다.',
    jev: ['자료만으로 답할 수 있는지 (확률로)', '성적 이의·개인 사정처럼 교수가 직접 답할 질문인지', '질문의 종류'],
    steps: [
      '새 구글 시트 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.',
      'start 를 실행하면 내 드라이브에 "수업 조교 자료" 폴더가 만들어집니다.',
      '강의계획서, 평가 기준, 공지 파일(구글 문서, PDF, 워드, 한글, 파워포인트, 텍스트, 사진)을 그 폴더에 넣습니다. 이 내용에만 근거해 답합니다.',
      '배포 → 새 배포 → 웹 앱(실행 사용자: 나, 액세스: 모든 사용자)으로 학생용 /exec 주소를 받습니다.',
      '메뉴의 매주 월요일 리포트 켜기로 주간 요약 메일을 받습니다.',
    ],
    examples: [
      '영어로 묻는 질문에는 영어로 답하게 해 줘.',
      '교수님께로 넘어간 질문 중 아직 답하지 않은 것을 매일 저녁 7시에 메일 한 통으로 보내 줘.',
      '시험 기간(2026-12-10 ~ 2026-12-20)에는 자동 답변을 끄고 모든 질문을 교수님께로 보내 줘.',
      '자동 답변이 틀렸다고 학생이 누를 수 있는 버튼을 넣고, 누른 질문은 시트에 빨간색으로 표시해 줘.',
    ],
    jevPreset: {
      label: '조교: 자료로 답할 수 있나',
      state: { 자료: '출석: 4회 초과 결석 시 F. 과제: 매주 금요일 23:59까지 제출. 중간고사: 10월 24일.', 질문: '결석 몇 번까지 괜찮아요?' },
      questions: {
        answerable: Q.noul('자료에 적힌 내용만으로 질문에 정확히 답할 수 있는가? 자료에 근거가 없으면 아니오.'),
        personal: Q.noul('성적 이의, 건강, 개인 사정, 상담처럼 교수가 직접 답해야 하는 질문인가?'),
        kind: Q.choice('질문의 종류', { content: '수업 내용과 개념', task: '과제, 평가, 시험', schedule: '일정, 장소, 출석 같은 행정', personal: '개인 사정, 성적 이의, 상담', other: '기타' }),
      },
    },
  },
  {
    id: 'r3', web: { deploy: true, links: [{ label: '학생용 주소 (LMS나 QR로 나눠 주세요)', suffix: '' }, { label: '교수 화면 주소 (프로젝터에 띄우세요. 처음 준비 때 알림창에 나온 PIN 4자리가 필요합니다)', suffix: '?v=class' }], after: ['시트의 [설정] 탭에서 주제, 학습 목표, 교수용 핵심 개념을 바꾸세요. 핵심 개념은 학생 화면에 나오지 않습니다.', 'PIN은 시트 메뉴 튜터 → 주소와 교수 화면 PIN 보기에서 다시 볼 수 있습니다(시트 칸에는 적지 않습니다). 학생에게 알리지 마세요. 한 번 맞히면 그 창을 닫을 때까지 다시 묻지 않습니다.'] }, file: 'r3-tutor.gs', name: '소크라테스 튜터와 대화록 제출', tag: '정답 대신 되묻는 AI',
    short: '정답을 알려 주지 않고 되묻는 튜터입니다. 대화 전체와 이해도 흐름이 시트에 저장되어 "AI를 어떻게 썼는지"를 평가할 수 있습니다.',
    jev: ['지금 이해도 (0~3 단계)', '오개념이 있는지', '튜터의 다음 행동 고르기 (힌트·반례·이유 묻기·심화·정리)'],
    steps: [
      '새 구글 시트 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.',
      '[튜터] 메뉴 → 처음 준비 → 수업 코드 입력.',
      '[설정] 시트에 주제, 학습 목표, 교수용 핵심 개념을 적습니다. 핵심 개념은 학생 화면에 나오지 않습니다.',
      '배포 → 새 배포 → 웹 앱(실행 사용자: 나, 액세스: 모든 사용자)으로 학생용 /exec 주소를 받습니다.',
      '[대화록]과 [제출] 시트를 평가 근거로 씁니다. 이해도는 교수만 보는 참고값이며 성적에 쓰지 않습니다.',
    ],
    examples: [
      "주제를 '근대 물리치료의 역사'로 바꾸고, 튜터가 편한 말투로 말하게 해 줘.",
      '이해도가 3이 되면 학생에게 비슷한 문제를 하나 내고 풀어 보게 한 뒤 정리로 넘어가 줘.',
      '[제출] 시트에서 마지막 이해도가 0~1인 학생 행만 노란색으로 칠하는 메뉴를 추가해 줘.',
      '학생이 같은 말을 세 번 반복하면 튜터가 힌트를 한 단계 더 크게 주게 해 줘.',
    ],
    jevPreset: {
      label: '튜터: 이해도와 다음 행동',
      state: { 주제: '광합성', 교수용핵심개념: '이산화탄소와 물이 빛에너지로 포도당과 산소가 된다. 빛은 재료가 아니라 에너지다.', 학생의마지막말: '식물은 햇빛을 먹고 산다고 생각해요.' },
      questions: {
        understanding: Q.score('학생의 마지막 말로 볼 때 핵심 개념을 얼마나 이해했는가?', ['거의 이해하지 못함', '일부만 이해함', '대체로 이해함', '정확하고 깊게 이해함']),
        misconception: Q.noul('학생의 마지막 말에 핵심 개념과 어긋나는 명백한 오개념이 있는가?'),
        move: Q.choice('튜터가 다음에 할 일 하나를 고르세요. 정답을 직접 알려 주는 선택지는 없습니다.', { hint: '작은 힌트나 쉬운 예로 한 걸음 안내하는 질문', counter: '오개념을 반례나 모순으로 스스로 알아차리게 하는 질문', probe: '맞는 것 같지만 이유를 더 설명하게 하는 질문', deepen: '정확히 이해했으니 한 단계 더 어려운 질문', wrap: '충분히 이해했으니 스스로 한 문단으로 정리하게 하기' }),
      },
    },
  },
  {
    id: 'r4', web: { deploy: true, links: [{ label: '학생용 주소 (학기 초 편지 쓰기)', suffix: '' }], after: ['시트 메뉴 미래 편지 → 웹 앱 주소 넣기에 위의 /exec 주소를 꼭 넣으세요. 편지 메일의 회고 링크에 씁니다(지금 모두 배달을 누르면 먼저 물어봅니다). 연습할 때는 본인 메일로 편지를 쓰고, 시트 메뉴 [미래 편지] → 지금 모두 배달(시연용)을 눌러 회고 메일이 오는 것까지 확인하세요.'] }, file: 'r4-future.gs', name: '미래의 나에게', tag: '학기 초 편지 → 학기말 도착',
    short: '학기 초에 목표와 걱정을 쓰면 도착일에 메일로 돌아오고, 회고를 쓰면 처음 글과 비교한 코멘트를 받습니다.',
    jev: ['처음 글과 비교한 변화·성장 (0~3 단계)', '회고글의 정서 (뿌듯함·섞임·아쉬움·무덤덤)'],
    steps: [
      '새 구글 시트 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.',
      '[미래 편지] 메뉴 → 처음 준비 → 수업 코드 입력.',
      '배포 → 새 배포 → 웹 앱(실행 사용자: 나, 액세스: 모든 사용자)으로 학생용 /exec 주소를 받고, 시트 메뉴 미래 편지 → 웹 앱 주소 넣기에 넣습니다(편지 메일의 회고 링크). 메뉴의 매일 아침 배달 켜기를 누릅니다.',
      '연습할 때는 본인 메일로 편지를 쓰고, 메뉴의 지금 모두 배달(시연용)을 눌러 회고 링크가 오는 것까지 확인합니다.',
      '메일 주소는 편지 배달에만 쓰이고 내 시트에만 남습니다. 학생에게 이 사실을 먼저 알려 주세요.',
    ],
    examples: [
      '도착일을 학기말이 아니라 편지를 쓴 날부터 30일 뒤로 고정해 줘.',
      '회고 코멘트 끝에 다음 학기에 해 볼 학습 활동 두 가지를 덧붙여 줘.',
      '메일 제목에 학생 별칭을 넣고, 하루 메일 한도(100통)를 넘으면 남은 편지는 다음 날 배달하게 해 줘.',
      '회고를 쓰지 않은 학생에게 도착 3일 뒤 부드러운 재촉 메일을 한 번만 보내 줘.',
    ],
    jevPreset: {
      label: '미래 편지: 변화와 정서',
      state: { 처음목표: '발표를 두려워하지 않기, 매주 복습하기', 처음걱정: '시간 관리', 지금의회고: '발표는 아직 떨리지만 매주 복습은 지켰어요. 시간 관리는 여전히 어렵습니다.' },
      questions: {
        growth: Q.score('처음 글과 지금의 회고를 비교할 때 학생의 변화나 성장이 얼마나 뚜렷한가?', ['변화가 보이지 않음', '조금 나아짐', '뚜렷이 나아짐', '크게 성장함']),
        tone: Q.choice('회고글의 정서', { proud: '뿌듯하고 자신감이 있다', mixed: '성취와 아쉬움이 섞여 있다', regret: '아쉬움과 후회가 크다', flat: '무덤덤하거나 구체성이 없다' }),
      },
    },
  },
  {
    id: 'r5', web: { deploy: false, links: [], after: ['실행 후 뜨는 알림창의 첫 번째 폴더에 칠판 사진을 올리세요. 10분 안에 두 번째 폴더에 노트가 생기고 메일이 옵니다.', '이 도구는 웹 앱으로 배포하지 않습니다. 배포 단계가 없고, 배포해서 주소를 열면 안내 화면만 보입니다.'] }, file: 'r5-boardphoto.gs', name: '칠판 사진이 수업 노트로', tag: '폰 사진 → 구글 문서 + 복습 퀴즈',
    short: '폰으로 찍은 칠판 사진을 드라이브 폴더에 올리면 10분 안에 수업 노트와 복습 퀴즈 3문항이 구글 문서로 생깁니다.',
    jev: ['수업 내용인지, 과제·시험 공지인지, 일정 공지인지', '너무 흐려서 다시 찍어야 하는지'],
    steps: [
      '새 구글 시트 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.',
      '[칠판 노트] 메뉴 → 처음 준비(드라이브에 폴더 두 개가 생김) → 수업 코드 입력 → 자동 처리 켜기.',
      '메뉴의 폴더 주소 보기에서 "사진 올리기" 폴더를 폰에 즐겨찾기 합니다.',
      '칠판 사진을 그 폴더에 올립니다. 얼굴이나 학생 이름이 찍히지 않게 하세요.',
      '10분 안에 "결과" 폴더에 노트가 생기고 메일이 옵니다.',
    ],
    examples: [
      '퀴즈를 5문항으로 늘리고, 정답은 문서 맨 아래에 따로 모아 줘.',
      '노트를 한국어와 영어 두 가지로 만들어 문서 하나에 위아래로 넣어 줘.',
      '공지로 분류된 사진은 노트 대신 사진 속 날짜를 읽어 구글 캘린더에 일정으로 등록해 줘.',
      '노트 문서 제목에 수업 날짜와 과목명(설정 시트에서 읽기)을 넣어 줘.',
    ],
    jevPreset: {
      label: '칠판: 종류와 판독 가능 여부',
      state: '# 광합성 정리\n## 재료\n- 물\n- 이산화탄소 [?]\n## 산물\n- 포도당, 산소\n다음 주 월요일까지 보고서 제출',
      questions: {
        kind: Q.choice('이 노트는 어떤 종류의 칠판 내용인가?', { lecture: '수업 내용을 설명하는 판서', assignment: '과제나 시험에 대한 공지', schedule: '일정이나 장소에 대한 공지', other: '수업과 무관하거나 내용이 거의 없음' }),
        unreadable: Q.noul('노트에 [?] 표시가 많거나 대부분을 읽지 못해 쓸 만한 노트가 아닌가?'),
      },
    },
  },
  {
    id: 'r6', web: { deploy: true, ppt: { suffix: '', title: '수업 후 설문 결과' }, links: [{ label: '강의장 결과 화면 (프로젝터나 슬라이드에 띄우세요. 응답이 올 때마다 움직입니다)', suffix: '' }], after: ['실행 후 뜨는 알림창의 설문 주소를 학생에게 보여 주세요(메뉴의 주소 보기에서 다시 볼 수 있습니다). 응답이 오면 [분석] 탭에 쌓이고, 위 결과 화면이 움직이고, 내 메일로 요약이 옵니다.'] }, file: 'r6-survey.gs', name: '수업 후 설문 즉석 분석', tag: '폼 제출 → 감정·키워드·건의 → 내 메일',
    short: '설문 폼이 제출될 때마다 감정, 주제, 건의 여부, 긴급 여부를 판정해 시트에 쌓고, 일정 개수가 모이면 요약을 메일로 보냅니다.',
    jev: ['수업에 대한 감정 (0~4 단계)', '가장 비중 있게 말한 주제', '구체적인 건의가 있는지', '교수가 바로 알아야 할 긴급 내용인지'],
    steps: [
      '새 구글 시트 → Apps Script에 공통 도우미 + 이 코드를 붙여 넣고 저장합니다.',
      '[설문 분석] 메뉴 → 처음 준비(설문 폼이 만들어지고 이 시트에 연결됨) → 수업 코드 입력.',
      '메뉴의 주소 보기에서 학생용 설문 주소를 받아 QR로 보여 줍니다.',
      '[설정] 시트에서 요약 메일을 받을 개수와 주소를 정합니다.',
      '긴급으로 판정된 응답은 즉시 메일로 오고, [분석] 시트에 한 줄씩 쌓입니다.',
    ],
    examples: [
      '만족도가 2 이하인 응답만 [주의] 시트에 따로 모으고, 매주 금요일 오후 5시에 그 목록을 메일로 보내 줘.',
      "설문에 '이번 시간에 가장 헷갈린 개념'을 주관식으로 추가하고, 요약 메일에 그 개념 TOP3를 넣어 줘.",
      '요약 메일을 두 번째 주소(조교)에게도 같이 보내 줘.',
      '주제별 평균 만족도를 [주제별] 시트에 표로 만들어 줘.',
    ],
    jevPreset: {
      label: '설문: 감정·주제·건의·긴급',
      state: '좋았던 점: 예시가 많아서 이해가 잘 됐어요\n아쉬웠던 점이나 바라는 점: 실습 때 진행이 너무 빨라서 따라가기 힘들었어요. 더 천천히 해 주세요.',
      questions: {
        mood: Q.score('수업 전체에 대한 학생의 감정은 어떤가?', ['매우 부정', '부정', '중립', '긍정', '매우 긍정']),
        topic: Q.choice('학생이 가장 비중 있게 말한 주제', { level: '수업 내용의 난이도', pace: '진행 속도', material: '자료와 교재', practice: '실습과 활동', manner: '교수의 태도와 소통', env: '강의실, 시설, 온라인 환경', other: '기타' }),
        suggest: Q.noul('다음 수업에 반영할 수 있는 구체적인 건의나 요청이 들어 있는가?'),
        urgent: Q.noul('교수가 지금 바로 알아야 할 심각한 불만, 안전 문제, 차별, 괴롭힘 이야기인가?'),
      },
    },
  },
  {
    id: 'blank', web: { deploy: false, links: [], after: ['만든 코드를 붙여 넣은 편집기에서 위쪽 함수 선택 칸으로 새 함수를 고르고 ▶ 실행하세요. 오류가 나면 아래 4번 칸에 붙여 넣어 고칩니다.'] }, file: null, name: '빈 종이에서 시작', tag: '하고 싶은 일을 말로 설명',
    short: '공통 도우미만 두고, 하고 싶은 자동화를 말로 설명하면 AI가 코드를 만듭니다. 막히면 오류를 붙여 넣어 바로 고칩니다.',
    jev: ['글을 읽고 예/아니오, 여럿 중 하나, 단계 점수를 확률로 판정 (jev_ 도우미)'],
    steps: [
      '새 구글 시트 → Apps Script에 아래 공통 도우미를 붙여 넣고 저장합니다.',
      '하고 싶은 일을 아래 칸에 적고 "AI ONE으로 코드 만들기"를 누릅니다.',
      '만들어진 코드를 Apps Script 편집기의 코드 아래에 붙여 넣고 ▶ 실행합니다.',
      '오류가 나면 오류 문구를 그대로 복사해 "실행하다 오류가 났다면" 칸에 붙여 넣고 고칩니다.',
    ],
    examples: [
      'A열에 학생 이름, B열에 점수가 있는 시트에서 60점 미만 이름 칸을 빨갛게 칠하고 그 목록을 내 메일로 보내는 함수를 만들어 줘.',
      '구글 폼 응답이 들어올 때마다 응답자에게 감사 메일을 자동으로 보내 줘.',
      '매주 월요일 아침 8시에 이번 주 내 캘린더 일정을 메일 한 통으로 보내 줘.',
      'A열의 학생 글을 Jev로 판정해서 "주제와 관련 있는가"를 B열에 확률로 적어 줘. jev_ 도우미를 써.',
      'A열의 긴 글을 읽고 B열에 한 줄 요약을 ai_ 도우미로 채우는 메뉴를 만들어 줘. 한 번에 20줄씩만 처리해 줘.',
    ],
    jevPreset: {
      label: '빈 종이: 글이 주제와 관련 있나',
      state: '오늘 수업에서 배운 광합성은 식물이 빛에너지로 포도당을 만드는 과정이다.',
      questions: {
        related: Q.noul('이 글이 광합성에 대한 내용인가?'),
        level: Q.score('글의 설명이 얼마나 정확하고 구체적인가?', ['틀렸거나 막연함', '부분적으로 맞음', '대체로 정확함', '정확하고 구체적임']),
      },
    },
  },
];

// 혼자 쓰는 판: 공통 도우미의 "수업 서버" 부분(// <<SERVER … // SERVER>>)만 solo.gs 로 바꾸고, 레시피 글의 수업 코드·Jev 말을 내 API 키·AI로 바꿉니다.
const SERVER_BLOCK = /\/\/ <<SERVER[\s\S]*?\/\/ SERVER>>\n?/;
const SOLO_WORDS = [["'수업 코드 입력', 'setClassCode'", "'AI 키 넣기', 'setApiKey'"],
  ['사용법: 웹 도구(https://kimcook-edu.pages.dev/#gas) 또는 함께 받은 PDF 키트의 안내를 따르세요.', '사용법: 함께 받은 PDF 키트의 안내를 따르세요.'],
  ['사용법: 웹 도구(https://kimcook-edu.pages.dev/#gas)에서 이 레시피를 고르고 안내를 따르세요.', '사용법: 함께 받은 PDF 키트의 안내를 따르세요.'],
  // 강사 서버에 올려 둔 QR 그림 도구 대신 공개 CDN(같은 라이브러리)에서 받습니다
  ['https://kimcook-edu.pages.dev/vendor/qrcode/qrcode.js', 'https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js'], ['JEV 판정', 'AI 판정'], ['Jev가', 'AI가'], ['Jev', 'AI']];
export function soloCommon(common, solo) {
  if (!SERVER_BLOCK.test(common)) throw new Error('00-common.gs 에 // <<SERVER … // SERVER>> 표시가 없습니다.');
  return common.replace(SERVER_BLOCK, () => solo.replace(/\s+$/, '') + '\n');
}
export function soloRecipe(code) {
  return SOLO_WORDS.reduce((c, [from, to]) => c.split(from).join(to), code);
}

// 브라우저(키트 PDF)에서 연수용 코드를 혼자 쓰는 판으로 바꾸는 함수. 이 함수 본문이 gas-recipes.js 에 그대로 들어가므로 D 외의 바깥 변수를 쓰지 않습니다.
function makeToSolo(D) {
  return function toSolo(code) {
    var c = String(code == null ? '' : code), had = c.indexOf(D.common.code) === 0;
    if (had) c = D.solo.code + c.slice(D.common.code.length);
    else c = c.replace(new RegExp(D.serverBlock), function () { return D.soloBlock; });
    D.soloWords.forEach(function (w) { c = c.split(w[0]).join(w[1]); });
    return c;
  };
}

export function buildRecipes() {
  const common = read('00-common.gs');
  const out = {
    soloWords: SOLO_WORDS, serverBlock: SERVER_BLOCK.source, soloBlock: read('solo.gs').replace(/\s+$/, '') + '\n',
    common: { code: common, version: (common.match(/CODE_VERSION = '([^']+)'/) || [])[1] }, solo: { code: soloCommon(common, read('solo.gs')) }, promptRule: PROMPT_RULE, trouble: TROUBLE, soloTrouble: SOLO_TROUBLE,
    recipes: RECIPES.map((r) => ({
      id: r.id, web: r.web, wow: EXTRA[r.id].wow, story: EXTRA[r.id].story, name: r.name, tag: r.tag, short: r.short, jev: r.jev, steps: r.steps, examples: r.examples,
      code: r.file ? read(r.file) : '', soloCode: r.file ? soloRecipe(read(r.file)) : '', jevPreset: r.jevPreset,
    })),
  };
  return '/* 자동 생성 파일: node gas/build.mjs 로 gas/*.gs 에서 만듭니다. 직접 고치지 마세요. */\nwindow.GAS_RECIPES = ' + JSON.stringify(out) + ';\nwindow.GAS_RECIPES.toSolo = (' + makeToSolo.toString() + ')(window.GAS_RECIPES);\n';
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const target = path.join(here, '..', 'gas-recipes.js');
  fs.writeFileSync(target, buildRecipes());
  console.log('wrote', target);
}
