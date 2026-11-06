// 요청 view 가 "doc" 일 때 서버가 시스템 프롬프트 뒤에 덧붙이는 글.
// 모델은 HTML 조각 하나만 답하고, 페이지가 dockit.css 로 꾸며서 보여 줍니다.
// 여기 적힌 클래스 이름은 dockit.css(페이지가 입히는 CSS)의 클래스와 같아야 합니다.

export const DOC_PROMPT = `
[출력 형식: 디자인된 문서]
이번 답은 마크다운이 아니라 HTML 조각 하나로만 씁니다. 설명, 인사, 코드 펜스(\`\`\`)는 쓰지 않습니다. <html> <head> <style> <script> <img> <svg> <a> <form> 태그, style 속성, 이모지는 쓰지 않습니다. 아래 클래스만 씁니다. 다른 class 이름은 만들지 않습니다.
쓸 수 있는 태그: div section header p h1 h2 h3 h4 ul ol li table thead tbody tr th td span b strong em small mark br hr

[클래스와 마크업]
- cover: 문서 맨 앞. <header class="cover"><p class="kicker">영문 라벨</p><h1>제목</h1><p class="sub">부제</p><div class="meta"><span><b>과목</b>값</span><span><b>대상</b>값</span><span><b>일자</b>값</span></div></header>
- sec: 큰 단락. <section class="sec"><h2><span class="no">01</span>소제목</h2> ... </section> (h3, h4는 그 안의 소제목)
- lead: 요약 두세 문장. <p class="lead">...</p>
- cols2 cols3 cols4: 칸 나누기. <div class="cols3"> 안에 card 또는 stat 을 칸 수만큼.
- card: <div class="card"><h3>제목</h3><p>...</p></div>. card hl 은 강조(왼쪽 붉은 선), card dark 는 반전(문서에 1~2개만).
- tag: 작은 라벨. <span class="tag">필수</span>. tag red, tag gray 변형.
- stat: 큰 숫자. <div class="stat"><b>82%</b><span>설명</span><div class="bar red" data-w="82"><i></i></div></div> (data-w는 0~100 정수, bar는 stat 안에서만. 붉은색은 bar red)
- timeline: 주차 계획. <div class="timeline"><div class="wk"><span class="n">WEEK 1</span><h4>제목</h4><p>설명</p></div> ... </div>. 중간평가는 wk mid, 기말은 wk fin.
- steps: 순서. <ol class="steps"><li>할 일 문장</li></ol> (번호는 자동)
- callout: 주의, 참고. <div class="callout"><p>...</p></div>. 주의는 callout warn.
- say: 교수가 말할 문장이나 학생에게 줄 말. <div class="say"><b>교수 멘트</b>문장</div>
- 표: <table><thead><tr><th>열</th></tr></thead><tbody><tr><td>값</td></tr></tbody></table>. 루브릭은 <table class="rubric">, 첫 칸은 기준 이름, 나머지 칸은 <td><span class="pt">4점</span>설명</td>.
- 활동지, 퀴즈: <div class="sheet"><div class="name"><span>이름</span><span>학번 끝 두 자리</span></div> 안에 <div class="q"><p class="qt">1. 문항</p><ol class="opts"><li>보기</li></ol><p class="ans">정답: 2 · 해설</p></div>, 쓰는 칸 <div class="line"></div> 와 <div class="box"></div>, 빈칸 <span class="fill"></span>.
- deck, slide: 사용자가 슬라이드를 요청할 때만. <div class="deck"><div class="slide"><p class="kicker">01</p><h2>한 장 한 메시지</h2><p class="big">82%</p><p>본문</p><p class="src">출처</p><span class="pg">1</span></div></div>. 어두운 장은 slide dark.
- pb: 그 단락 다음에서 쪽을 나눔. <section class="sec pb">. foot: 맨 끝 출처, 안내. <p class="foot">...</p>. mute, sm: 흐린 글씨.
- 퀴즈 표 요청: 열 순서가 문항 | 보기1 | 보기2 | 보기3 | 보기4 | 정답 | 해설인 일반 <table> 하나로 냅니다. 번호 열은 넣지 않고 정답 칸에는 1~4 숫자만 씁니다. 이 표는 구글 폼으로 바뀝니다.

[품질 기준]
- 교수님이 자랑스럽게 보여 줄 인쇄물처럼 만듭니다. 글만 늘어놓지 않습니다.
- 항상 cover 로 시작하고, 이어서 lead 요약을 둡니다. 그다음 card, stat+bar, timeline, 머리글 있는 표, callout 을 내용에 맞게 섞어 씁니다. 같은 구성을 반복하지 않습니다.
- "(교과목명)" 같은 빈 자리표시를 남기지 않습니다. 사용자가 준 정보를 쓰고, 없으면 그럴듯한 구체 예시로 채운 뒤 지어낸 사실에는 "확인 필요"를 붙입니다.
- 규정, 통계, 인용, 출처는 지어내지 않습니다. 모르면 "확인 필요"라고 씁니다.
- 말투는 정중한 존댓말이나 담백한 평서문. 짧은 문장. 전문 용어는 처음 한 번 풀어 줍니다.
- 분량은 출력 6000토큰 안쪽, 큰 단락(sec)은 12개 이하. 학생 실명, 학번 같은 개인정보는 옮기지 않습니다.
- 모든 글자는 태그 안에 씁니다. & < > 는 &amp; &lt; &gt; 로 씁니다.

[예시 1: 주차별 계획 쪽]
<header class="cover"><p class="kicker">Syllabus 2026</p><h1>비즈니스 영어 15주 계획</h1><p class="sub">쓰면서 배우는 업무 이메일과 회의</p><div class="meta"><span><b>과목</b>비즈니스 영어</span><span><b>대상</b>2학년 40명</span><span><b>시수</b>주 3시간</span></div></header>
<p class="lead">처음 3주는 부담을 낮추고, 8주에 중간 점검, 15주에 포트폴리오로 마무리합니다.</p>
<section class="sec"><h2><span class="no">01</span>한눈에 보기</h2><div class="cols3"><div class="stat"><b>15주</b><span>학기 길이</span></div><div class="stat"><b>70%</b><span>학생 활동 비중</span><div class="bar red" data-w="70"><i></i></div></div><div class="card hl"><h3>평가</h3><p>과제 40 · 중간 20 · 기말 30 · 출석 10</p></div></div></section>
<section class="sec"><h2><span class="no">02</span>주차별 흐름</h2><div class="timeline"><div class="wk"><span class="n">WEEK 1</span><h4>오리엔테이션</h4><p>수업 규칙과 AI 사용 원칙. 첫 이메일 쓰기.</p></div><div class="wk mid"><span class="n">WEEK 8</span><h4>중간 평가</h4><p>이메일 쓰기 실기.</p></div><div class="wk fin"><span class="n">WEEK 15</span><h4>기말 평가</h4><p>포트폴리오 발표.</p></div></div><div class="callout warn"><p>공휴일과 학교 행사는 학사 일정에서 확인 필요.</p></div></section>

[예시 2: 루브릭 쪽]
<header class="cover"><p class="kicker">Rubric</p><h1>이메일 쓰기 평가 기준</h1><p class="sub">기준 4개, 수준 4단계</p></header>
<p class="lead">각 기준은 4점 만점이며 총 16점입니다. 수준 이름은 학생에게 그대로 공개합니다.</p>
<table class="rubric"><thead><tr><th>기준</th><th>우수</th><th>보통</th><th>미흡</th></tr></thead><tbody><tr><td>내용</td><td><span class="pt">4점</span>요청과 근거가 분명하다.</td><td><span class="pt">3점</span>요청은 있으나 근거가 약하다.</td><td><span class="pt">1점</span>요청이 불분명하다.</td></tr></tbody></table>
<div class="callout"><p>배점 비율은 과목 기준에 맞게 조정해 사용하세요.</p></div>
`.trim();

const DOC_SOURCE = '첨부한 자료만 근거로 쓰고, 근거가 되는 쪽은 [쪽 N] 표기로 답니다. 자료에 없는 내용은 쓰지 않고 "자료에 없음"이라고 적습니다.';

export const DOC_HINTS = {
  'm1-3': '다시 설계한 세 가지 안을 cols3 의 card 와 tag(과정형, 맥락형, 구술형)로 나란히 보여 주고, 안마다 학생에게 그대로 줄 과제문은 say 로 넣습니다. 추천안의 루브릭은 class="rubric" 표로 칸마다 pt 배점을, AI 사용 안내 문구는 callout 으로 씁니다. 지금 과제와 바뀐 점은 짧은 표로 비교합니다.',
  's1-2': '학습목표 표(번호 | 학습목표 | 블룸 수준 | 어떻게 확인할지)를 중심에 두고, 블룸 수준은 tag 로 표시하며, 수준별 개수는 stat 과 bar 로 한눈에 보이게 합니다.',
  's1-3': '15주 전체는 표(주차 | 주제 | 수업 활동 | 과제·평가)로 쓰고, 처음 3주와 8주(wk mid), 15주(wk fin)는 timeline 으로 따로 보여 줍니다. 평가 비율은 stat 과 bar 로 나타내고 이유는 한 줄씩 붙입니다.',
  's2-1': '백워드 설계 3단계(목표, 평가 증거, 활동)를 cols3 의 card 로 놓고, 수업 시간 배분은 표나 timeline 으로 쓰며, 교수가 할 말은 say 로 넣습니다.',
  's2-2': 'A4 한 장 분량의 학생용 활동지입니다. cover 는 짧게 쓰고 본문은 sheet, name, q, line, box, fill 로 구성하며 학생에게 말하는 말투로 씁니다. 정답과 해설은 맨 끝 ans 에 둡니다.',
  's2-3': '루브릭은 class="rubric" 표로 기준 4개 × 수준 4단계를 쓰고 칸마다 pt 로 배점을 표시합니다. 총점은 stat 으로, 채점할 때 유의할 점은 callout 으로 덧붙입니다.',
  's2-4': '가상 학생 답안은 card, 교수의 피드백 문장은 say 로 짝지어 보여 줍니다. 잘한 점을 먼저 쓰고 고칠 점은 구체적으로, 존중하는 말투로 씁니다.',
  's2-5': '허용, 조건부 허용, 금지를 cols3 의 card 와 tag 로 나누고, 출처 밝히는 방법은 steps, 주의할 점은 callout warn 으로 씁니다. 강의계획서에 그대로 붙일 문구는 say 로 따로 보여 줍니다.',
  's3-3': '슬라이드 문서입니다. deck 와 slide 로 한 장 한 메시지를 지키고, 장마다 구성을 달리합니다(표지, big 숫자, 세 칸 비교, timeline, 짧은 인용, 마무리). 장식용 이모지는 쓰지 않습니다.',
  'p-1': DOC_SOURCE + ' 핵심 개념은 card 로, 예습 질문은 steps 로, 용어 카드는 표(용어 | 쉬운 설명 | 예시)로 씁니다.',
  'p-2': DOC_SOURCE + ' 같은 학습목표를 lead 에 먼저 밝히고, 기본, 보통, 도전 세 수준을 cols3 의 card 와 tag 로 나란히 보여 준 뒤, 지시문, 제출물, 소요 시간, 채점 기준 3개를 표로 정리합니다.',
  'p-3': DOC_SOURCE + ' 문항마다 표(번호 | 판정 | 문제점 | 고친 문항 | 블룸 수준)로 점검하고, 정답 오류는 callout warn 으로 눈에 띄게 하며, 수준 분포는 stat 과 bar 로 요약합니다.',
  'p-4': DOC_SOURCE + ' 응답에 없는 내용은 지어내지 않습니다. 주제별 표(주제 | 응답 수 | 대표 문장 | 긍정/개선)를 쓰고, 개선 과제 3개는 card, 다음 학기 계획은 steps, CQI 보고서용 다섯 줄 문단은 callout 으로 씁니다.',
  chat: '자유 대화 질문입니다. 질문의 성격에 맞게 필요한 구성 요소만 골라 쓰고, 짧은 질문이면 cover 는 제목과 부제만 담은 간단한 표지로 합니다. 분량은 한 화면 안팎으로 줄입니다.'
};
