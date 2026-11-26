# 반응판 분류 고치기 계획

1. 서버 모델 고정: functions/api/jev.js DEFAULT_MODEL='jev-1.14.0' (env JEV_MODEL 유지), 시험·주석 맞춤.
2. r1-board.gs submitAnswer: Jev 전에 낱말 맞추기(정규화 후 묶음 이름/낱말 포함 여부, 한 묶음에만 맞으면 확정, conf='낱말').
   Jev 실패 시 cat='' 대신 오류 흔적(conf='오류') 남김 → [분류]가 다시 처리.
3. 큰 화면 [분류] 단추(교수만, auth_): 서버 함수 classifyBoard
   (a) 묶음 있으면 기타·미분류·오류 답을 낱말→Jev 로 다시 분류(직접·손 고친 칸 보존, 상한, 연타 막기)
   (b) 묶음 없거나 기타 많으면 2-gram 자카드 군집으로 묶음 제안 → 큰 화면에 보이고 [바꾸기]로 setBoard 저장.
4. 혼자 판(solo.gs)에서도 낱말/군집은 API 없이 동작.
5. 설명 글(story.mjs, README, 레시피 설명) 맞춤.
6. 평가: eval-board/ 에 질문 3개 이상 × 25~40 답 + 정답, 낱말 단계/낱말+Jev/Jev만/군집 표. 문턱 0.45 판단.
7. 시험: gas.test.mjs 에 낱말, 오류 흔적, classifyBoard(인증·보존·상한), 군집 추가. 지정 시험 5개 + build.mjs.
8. RESULT.md 단계마다 갱신.
