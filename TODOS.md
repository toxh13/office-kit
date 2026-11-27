# TODOS

## v0.2.1 이후

### 탭 잠재우기 (뷰 내리기)
- **What:** 오래 안 쓴 저장된 Office 탭의 뷰를 메모리에서 내리고, 누르면 다시 불러오기.
- **Why:** Office 탭이 많을 때 메모리 절약. v0.2.0 Eng Review D2에서 연기.
- **Pros:** 탭을 많이 열어도 앱이 가벼움. **Cons:** 되돌리기 기록·커서가 사라짐, 상태 판정 시험 필요.
- **Context:** 먼저 Office 뷰 실제 메모리를 잰다(editor-window.mjs:52의 0.5GB는 측정 없음). v0.2.0의 지연 로드 장치를 재사용.
- **Depends on:** v0.2.0 탭 창(M3).

### 탭 끌어서 새 창으로 떼어 내기 / 합치기
- **What:** 탭을 끌어 새 탭 창으로 분리하거나 다른 창으로 옮기기.
- **Why:** 크롬식 사용감 완성. 설계 전제 1에서 연기.
- **Pros:** 두 모니터 작업이 편함. **Cons:** 뷰 이동 중 상태·IPC 주인 변경 처리.
- **Context:** D3 구조(editor-window가 뷰 소유)에서 host만 바꾸면 됨.
- **Depends on:** v0.2.0 탭 창.

### 나란히 보기
- **What:** 한 탭 창을 좌우로 나눠 두 문서를 동시에.
- **Why:** 사용자가 원한 네 기능 중 하나(v0.2.x로 연기).
- **Pros:** 원본 보며 작성·비교. **Cons:** 배치·포커스·단축키 규칙 추가.
- **Context:** tab-window.mjs 배치 계산에 분할 영역 추가.
- **Depends on:** v0.2.0 탭 창.

### 작업 묶음(탭 세트) 저장 + AI 레시피
- **What:** 열린 탭 묶음과 AI 요청을 이름 붙여 저장하고, 파일만 바꿔 다시 실행.
- **Why:** 사용자가 원한 기능 + 2차 의견의 '레시피' 아이디어.
- **Pros:** 반복 업무(월간 공문 등) 자동화. **Cons:** 저장 형식·파일 바꿔 끼우기 UI.
- **Context:** v0.2.0의 세션 저장(R8)과 읽기 재료 구조(M1)를 재사용.
- **Depends on:** v0.2.0 M1·복원.

### 고칠 문서 여러 개
- **What:** 한 대화에서 문서 여러 개를 함께 고치기.
- **Why:** v0.2.0은 고칠 문서 1개로 제한(전제 4).
- **Pros:** 여러 문서 일괄 수정. **Cons:** 쓰기 범위·문서별 버전·되돌리기·세션 키 설계 필요.
- **Context:** aiMcpEnv({edit:[...]}) 확장, keepVersion을 문서별로.
- **Depends on:** v0.2.0 M1.

### 탭 세션 가장자리 경우 3가지 (v0.2.0 /review Minor, QA에서 재현 못 함)
- **What:** ① 끝내는 중에 탐색기로 연 파일이 버려지고, 끝내기를 취소해도 다시 열리지 않음 ② 크래시로 올린 세션에서 '지난 탭 다시 열기'가 활성 탭을 바로 불러옴 ③ 올린 세션(promoteCrashed)이 활성 탭을 홈으로 기록.
- **Why:** 문서를 다시 열어야 하거나, 앱을 죽게 만든 문서를 다시 불러올 수 있음.
- **Pros:** 복원이 더 믿을 만해짐. **Cons:** 끝내기 취소 시 대기열·크래시 표시를 세션 형식에 더해야 함.
- **Context:** src/app/main.mjs openFromExplorer·quitInOrder, src/app/tab-session.mjs restorePlan·promoteCrashed. 2026-11-27 /qa 보고서 참고.
- **Depends on:** 없음.

### 탭 닫기 단추를 tablist 밖으로 (접근성 권고, LOW)
- **What:** 탭 줄 tablist 안에 tab 말고 button(닫기)이 있음. ARIA 소유 규칙상 tablist의 자식은 tab만.
- **Why:** 화면 읽기 프로그램이 탭 목록을 읽을 때 닫기 단추가 섞여 들림. 키보드는 Delete로 닫을 수 있음.
- **Pros:** 자동 검사(axe) 경고 없앰. **Cons:** 탭 모양·배치를 바꿔야 함.
- **Context:** src/app/renderer/tabstrip.js renderTab.
- **Depends on:** 없음.
