# AI Class Lab

부산외국어대학교 PSC교수학습기술센터 교수자 AI 활용 맞춤형 연수(2026. 11. 13.) 실습 페이지입니다.

- 주소: https://kimcook-edu.pages.dev/ (덱의 버튼과 QR이 쓰는 주소. GitHub를 거치지 않아 github.io가 막힌 곳에서도 열립니다)
- 예비 주소: https://skyman200.github.io/edu/ (같은 파일)
- 바이브코딩 스튜디오: `#vibe` — 기획서로 만든 프롬프트를 AI에 보내고, AI가 준 답을 붙여 넣으면 페이지 안의 샌드박스 iframe에서 바로 실행합니다. 오류는 줄 번호와 함께 잡아 AI에게 물어볼 문장으로 복사하고, 버전 저장과 HTML 내려받기를 지원합니다.
- 강의 자료 PDF: `ai-class-slides.pdf` (51쪽, 2026-11-07 미션판) — 맨 위 '강의 자료 PDF 내려받기'. 원본 덱은 `../부산외대_AI특강_prof.kimcook_미션판.pptx`(본편 29장 + 숨긴 부록 21장 + 마무리). 부록(자료실 STEP 1~5)은 슬라이드 쇼에서 숨겨져 있고 PDF에는 들어 있습니다. PDF 안의 OPEN 버튼과 웹앱 줄은 PPT와 같은 주소로 열립니다.
- 접속 점검: `/check` — 연수 장소 네트워크에서 실습 사이트·AI ONE·참여판·AI 서비스·구글이 열리는지 확인하고, 막히면 전산 요청문을 만들어 줍니다. 미션 2 발음 앱용으로 이 컴퓨터의 언어별 목소리(소리 읽기)도 점검합니다.
- 과제 바로가기: `#s1-1` ~ `#s5-4`, `#p-1` ~ `#p-4` (예: `/#s2-1`)
- AI ONE 바로 실행: 구독이 없는 교수님용. 실습마다 'AI ONE으로 바로 실행' 버튼이 DIT AI-ONE 게이트웨이(Claude)로 프롬프트를 보내 답을 이 페이지에서 받습니다. 모델은 서버가 실습별로 자동 선택하고(짧은 요약은 Haiku 5.5, 표·계획·슬라이드·Apps Script는 Sonnet 5.5, 바이브코딩 웹앱은 Opus 5.5), 고른 모델과 이유를 답 위에 표시합니다. API 키는 브라우저에 오지 않고 Cloudflare Pages Function(`functions/api/ai.js`)에만 있습니다. 수업 코드(`AI_CLASS_CODE`)를 아는 사람만 쓸 수 있고, IP·코드·하루 전체 호출 한도가 걸려 있습니다. `/?c=코드` 링크(QR)로 열면 코드가 자동 입력됩니다.
- 연수 흐름(2026-11-07 개편): 첫 화면은 "오늘의 미션 2개"입니다. 미션 1(`#mission1`, `#m1-1`~`#m1-3`)은 내 과제를 AI가 학생처럼 풀게 하고, 교수가 채점해 점수를 강의장 화면으로 보낸 뒤, AI 시대에 맞게 과제를 다시 설계합니다. 미션 2(`#vibe`)는 바이브코딩 스튜디오로 수업 앱을 만들고, 옆 교수와 바꿔 써 본 뒤, 전시판에 올려 학생용 휴대폰 링크를 받습니다. 기존 STEP 1~5·PLUS·AI 작업실은 `#library`(자료실) 아래에 그대로 있고 주소(`#s1-1` 등)도 같습니다.
- 참여판(LIVE): `/live` — 강의장 화면용. 1 걱정(QR로 휴대폰 참여, "AI로 묶어 보기"), 2 AI 답안 점수(평균·분포·비교), 3 전시(올라온 앱 열기·휴대폰 QR·하나 뽑기), 4 다음 학기 한 가지. 숫자 키 1~4로 넘기고 Esc로 앱을 닫습니다. "지금부터 새로"는 리허설 데이터를 화면에서만 치웁니다(저장된 내용은 그대로). 휴대폰 참여 주소는 `/live?j=worry`·`/live?j=take`, 학생용 앱 주소는 `/live?app=<토큰>`(수업 코드 없이 열림). 방을 나누려면 `?r=이름`(기본 main)을 붙입니다.
- GAS 작업실: `/gas` — Apps Script 레시피 6개(반응판, 24시간 조교, 소크라테스 튜터, 미래의 나에게, 칠판 사진 노트, 설문 즉석 분석)와 빈 종이. 말로 요청하면 AI ONE(`gas-make`)이 코드를 고치고, 오류 문구를 붙여 넣으면 줄 번호를 읽고 원인 사전으로 처방하거나 AI ONE(`gas-fix`, 같은 오류가 반복되면 `gas-fix2`=Opus)이 고칩니다. 코드 점검, 실행 전 안전 점검, 처음 코드와의 비교, Jev 시험대가 있습니다. 소스는 `gas/*.gs`(공통 도우미 `00-common.gs` + 레시피)이고 `node gas/build.mjs`가 `gas-recipes.js`를 만듭니다(deploy.sh가 실행). 시험: `node gas/gas.test.mjs`(가짜 구글 환경에서 레시피 실행), `node gas/diag.test.mjs`.
- Jev API: `functions/api/jev.js`(`/api/jev`)가 DIT AI-ONE의 판정 모델(`jev-latest`, 환경변수 `JEV_MODEL`로 `jev-1.14.0` 등 지정 가능)에 예/아니오·여럿 중 하나·단계 점수 질문을 보냅니다. 수업 코드와 AI 호출 한도를 `/api/ai`와 같이 씁니다. Apps Script는 `jev_(글, 질문)` 도우미로 부릅니다. 시험: `node functions/_lib/jev.test.mjs`.
- 입력한 수업 정보는 브라우저(localStorage)에만 저장되고 어디에도 전송되지 않습니다.
- 내 결과물: 실습마다 'MY RESULT' 칸에 AI 답을 붙여 두면 `#mine`에 모이고, Word(.docx)·PDF(인쇄)·전체 복사로 한 번에 가져갈 수 있습니다. 이것도 브라우저에만 저장됩니다. .docx는 외부 라이브러리 없이 페이지 안에서 만듭니다.

## 고친 뒤 반영

`./deploy.sh` 한 번이면 됩니다. 공개할 파일(index.html, live.html, check.html, ping.js, ai-class-slides.pdf, vendor/)과 `functions/`만 임시 폴더에 모아 Cloudflare Pages(`kimcook-edu`)에 올립니다. wrangler는 `functions/`를 올리는 폴더가 아니라 현재 폴더에서 찾기 때문에, `wrangler pages deploy .`처럼 직접 올리지 마세요(README·`._*`·`.dev.vars`가 함께 공개되거나 Function이 빠집니다).
GitHub Pages 예비 주소는 `git push`(main)로 반영합니다. 예비 주소에는 Function이 없어서 AI ONE 버튼은 `kimcook-edu.pages.dev/api/ai`를 호출합니다(CORS 허용 목록에 있음).

## AI ONE 운영 (Cloudflare Pages 설정)

- 비밀값(Settings → Variables and Secrets, 또는 `npx wrangler pages secret put <이름> --project-name kimcook-edu`): `AI_GATEWAY_API_KEY`(DIT AI-ONE 키), `AI_CLASS_CODE`(수업 코드. 쉼표로 여러 개 가능). 바꾼 값은 다음 배포부터 적용됩니다.
- 호환성 플래그 `enable_request_signal`(Production·Preview)이 켜져 있어야 브라우저를 닫을 때 동시 처리 슬롯이 바로 풀립니다.
- 한도 기본값(코드 수정 없이 환경변수로 덮어쓸 수 있음): `AI_LIMIT_IP_10M=1500`, `AI_LIMIT_INFLIGHT=150`, `AI_LIMIT_BADCODE_10M=1000`, `AI_LIMIT_CODE_DAY=2500`, `AI_LIMIT_GLOBAL_DAY=3000`, `AI_LIMIT_GLOBAL_MIN=100`(게이트웨이는 회원당 60초 120건). 연수장은 IP 하나를 다 같이 쓰므로 IP별 한도는 사실상 연수장 전체 한도입니다.
- KV(`AI_KV`)는 연결하지 않았습니다. 연결하지 않으면 한도는 서버 인스턴스별 메모리로 셉니다. Workers Free의 KV 쓰기는 하루 1,000회라 모델 요청 약 160건이면 바닥납니다.
- 비상 정지: AI-ONE에서 키를 폐기·교체합니다(`AI_ENABLED=0`도 있지만 재배포가 필요하고, 이전 배포 주소 `<해시>.kimcook-edu.pages.dev`는 옛 설정으로 남습니다). 행사 뒤에는 키를 교체하고 이전 배포를 지웁니다.
- 참여판 저장소: D1 데이터베이스 `kimcook-edu-live`를 Pages 프로젝트에 `LIVE_DB`로 연결했습니다(Production·Preview). 표는 `functions/api/live.js`가 처음 요청 때 스스로 만듭니다. 쓰기는 모두 수업 코드(`AI_CLASS_CODE`)가 필요하고, 학생용 앱 주소만 코드 없이 읽힙니다. 끄려면 `LIVE_ENABLED=0`. 연습 데이터 지우기: `npx wrangler d1 execute kimcook-edu-live --remote --command "DELETE FROM posts WHERE room='main'"`. Free 한도(하루 쓰기 10만·읽기 500만 행) 안에서 넉넉합니다.
- 로컬 시험: `.dev.vars.example`을 보고 `.dev.vars`를 만든 뒤 `npx wrangler pages dev public --d1 LIVE_DB=local` (deploy.sh처럼 공개 파일을 public/에 모은 뒤). 라우터 시험: `node functions/_lib/route.test.mjs`.

prof.kimcook · 김강훈 (동의과학대학교 물리치료과)
