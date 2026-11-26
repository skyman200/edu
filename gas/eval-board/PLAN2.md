# 2차 계획 (리뷰 반영 + [분류] 한 번에 묶기)

원칙: 리뷰 1~5(꼭), 6~14(낮음), 15(사용자 요구)를 모두 고친다. live.html 은 고치지 않는다. 배포·커밋·push 없음.

1. r1-board.gs 서버
   - 1 setup 끝 연결 확인(post_ check, 서버판만) · getBoard.err(개수+이유, 학생 화면 없음) · 큰 화면 한 번 알림 · classify 토스트 r.stopped 우선.
   - 2 jevCat_ 엄격(other만 기타, 모양 이상이면 throw) · 시험 2개.
   - 4 matchCat_ 부정 넓힘, 통째 규칙은 두 낱말 이상 이름만, 한 낱말 이름은 낱말 머리 규칙만. 회귀 예문 시험. eval 재실행.
   - 5 submitAnswer 잠금 안에서 SEEN 기록 · CLASSIFY_MAX 서버 20/혼자 10, SEC 혼자 120 · 문구.
   - 9/10 잠금 안 key 재확인, F 빈 행 건너뜀, 쓰기 블록·상한 / err 판별 /^오류/||(C===''&&D==='').
   - 8 G1 빈 머리글이면 '자동 묶음'.
   - B 한 번에 묶기: classify_ 에서 묶음 없음/기타 많음이면 clusterAnswers_ → 60% 이상이면 saveGroups_(setBoard 와 같은 저장) → 낱말·Jev 로 옮김. 아니면 ai_ 로 3~6묶음+번호 → 저장, C·G 기록, 빠진 답만 Jev. 실패하면 저장 안 함, 이유 토스트. 제안 칸 제거(12). 되돌리기: 묶음 비우고 바꾸기.
2. 큰 화면 클라이언트: PIN pending(3), 편집 칸 [분류] 단추, 단축키 e.code(14), .labs 여백(6), 토스트 위치·실패 8초(7), 결과 토스트 문구.
3. 시험: gas.test.mjs 에 항목별 추가. 6개 명령 통과.
4. eval 재실행(Jev 호출 없이 캐시, 회귀 예문은 Jev 1.14 몇 번만), capture.mjs 갱신(단추 안 가림, 한 번에 묶기, 오류 알림).
5. RESULT.md '2차' 절.
