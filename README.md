<p align="center"><img src="build/icon.png" width="96" alt="Office Kit"></p>

# Office Kit

한글(HWP·HWPX)·Word·PowerPoint·Excel 문서를 **열고, 확인하고, AI 대화로 고치는** Windows 데스크톱 도구.

## 한눈에

| 무엇을 | 어떻게 |
|---|---|
| 문서 보기 | 창에 끌어다 놓기 → 왼쪽에 미리보기 |
| 새 문서 | `새 문서` (Ctrl+N) → 한글·Word·PowerPoint·Excel 빈 문서, 또는 AI에게 내용 부탁 |
| 문서 고치기 | 오른쪽 AI 창에 "담당자 칸에 홍길동 넣어줘"처럼 말하기 |
| 한글 직접 편집 | `직접 편집` (Ctrl+E) → 앱 안 한글 편집 창 (인터넷 없이, 한컴오피스 없이) |
| 문서·표·슬라이드 직접 편집 | Word·Excel·PowerPoint 형식(docx·xlsx·pptx)도 `직접 편집` (Ctrl+E) → 앱 안 편집 창 (인터넷 없이, MS Office 없이) |
| AI | 내 PC에 로그인해 둔 **Claude Code** 또는 **Codex**를 그대로 씀 (앱은 로그인 정보를 안 만짐) |
| 다른 AI 앱에서 | Claude Desktop·Claude Code·Codex에 문서 도구(MCP) 연결 |
| 명령줄 | `office-kit text 보고서.hwp` |
| 설치 | [Releases](https://github.com/toxh13/office-kit/releases)에서 `OfficeKit-Setup-<버전>.exe` 받아 실행 |

- 한컴오피스·MS Office가 **없어도** 한글/Word/PowerPoint/Excel 파일을 읽고 고칠 수 있다 (내장 엔진).
- 프로그램이 **있으면** 그 프로그램으로 원본 그대로 처리하고 PDF도 만든다.
- 한글 문서는 한컴오피스가 **없어도 PDF로 저장**된다 (내장 엔진, 미리보기와 같은 모양).

---

## 설치

1. [GitHub Releases](https://github.com/toxh13/office-kit/releases/latest)에서 **`OfficeKit-Setup-<버전>.exe`** 를 받는다.
2. 실행한다. 관리자 권한 없이 내 계정에만 설치된다 (기본 위치 `%LOCALAPPDATA%\Programs\Office Kit`).
3. "Windows의 PC 보호" 창이 뜨면 **추가 정보 → 실행** (코드 서명 전이라 뜨는 경고).

설치하면 생기는 것:

| 항목 | 내용 |
|---|---|
| 바로 가기 | 바탕화면 · 시작 메뉴 "Office Kit" |
| 오른쪽 클릭 메뉴 | 문서 파일에 **"Office Kit로 열기"** (기본 연결 프로그램은 그대로) |
| 대상 형식 | hwp hwpx hwt · docx doc rtf odt · pptx ppt odp · xlsx xls xlsm csv ods |

지우기: 설정 → 앱 → "Office Kit" 제거. 오른쪽 클릭 메뉴와 명령줄 파일도 함께 지워진다.

- 남는 것: 설정·최근 파일·**문서 버전**(연 문서의 사본 — 문서마다 최근 30개, 모두 합쳐 200개·2GB까지)·**미리보기 PDF**(최근 40개·512MB까지) `%APPDATA%\office-kit`, AI 작업 폴더 `%LOCALAPPDATA%\office-kit`. 필요 없으면 직접 지운다.
- Office Kit로 연 문서는 Windows "최근 항목"에 남고, 탐색기의 그 형식 "연결 프로그램" 목록에 Office Kit가 보일 수 있다 (기본 연결 프로그램은 그대로).

---

## 사용

### 문서 열기
- 창에 **끌어다 놓기**, 또는 `파일 열기`, 또는 탐색기에서 오른쪽 클릭 → **Office Kit로 열기**.
- 문서 버전을 저절로 남긴다: 문서를 열 때(**처음 연 상태**), 편집 창에서 저장하기 전(**직접 편집 저장 전**), AI 대화마다 고치기 전(**AI 대화 전**), 되돌리기 전(**원래대로 전**). 내용이 같으면 하나만 둔다 (문서마다 최근 30개).
- 상단 **원래대로**: 마지막 변경 전으로 되돌린다 (한 번 더 누르면 되돌리기 전으로). 옆 **▾** 를 누르면 버전 목록(시각·까닭)에서 **이 버전으로 되돌리기**.
- **프로그램에서 열기**: 설치된 Word·Excel·PowerPoint 창에서 고칠 수 있게(읽기·쓰기) 연다. 쓰기 암호가 걸렸거나 다른 프로그램이 쓰고 있으면 읽기 전용으로 열고 알려 준다.

### 새 문서
- 상단 **새 문서** (Ctrl+N) → 형식(한글 .hwpx · 한글 호환 .hwp · Word · PowerPoint · Excel), 파일 이름, 저장 위치를 고른다.
  저장 위치 기본값은 열린 문서의 폴더(없으면 문서 폴더). 같은 이름이 있으면 바꿀 이름을 알려 준다.
- 시작 방법: **빈 문서** 또는 **AI에게 내용 부탁하기** ("주간 회의록 양식 만들어줘"). 만든 문서는 바로 열리고 원본 백업도 만든다.
- **만든 뒤 바로 편집 창 열기**를 고를 수 있다 (기본 켜짐). 한글은 한글 편집 창, Word·Excel·PowerPoint 형식은 문서·표·슬라이드 편집 창으로 열린다.

### 직접 편집 (한글)
- 한글 문서(hwp·hwpx·hwt)를 연 뒤 상단 **직접 편집** (Ctrl+E) → 문서마다 편집 창이 하나 열린다 (다시 누르면 그 창으로).
- 편집기는 앱에 들어 있는 rhwp-studio다. **인터넷 없이** 동작하고 한컴오피스가 필요 없다. 앱 밖으로 나가는 요청은 모두 막는다
  (웹 요청뿐 아니라 WebRTC·WebTransport 같은 다른 연결과 맞춤법 사전 내려받기도). 편집기는 주 창과 다른 프로세스에서 돈다.
- 문서 안 링크(http·https·mailto)는 "기본 브라우저로 열까요?"를 물은 뒤 연다. 이 PC·사설망 주소는 열지 않는다.
- 저장 `Ctrl+S`: 원래 형식 그대로 (hwp·hwt → HWP, hwpx → HWPX). 덮어쓰기 전 내용은 백업 폴더에 남고, 왼쪽 미리보기도 바로 바뀐다.
  다른 이름으로 저장 `Ctrl+Shift+S` (.hwpx / .hwp). 닫기 `Ctrl+W` — 저장 안 한 내용이 있으면 저장 / 저장 안 함 / 취소를 묻는다.
  저장하는 동안 친 글은 **저장 안 됨**으로 남는다 (다시 저장하거나 닫을 때 묻는다). 파일의 보안 설정(나만 읽기 등)은 저장해도 바뀌지 않는다.
- 다른 곳(AI·다른 프로그램)에서 파일이 바뀌면: 편집 창에서 고친 것이 없으면 저절로 다시 불러오고, 있으면 **다시 불러오기 / 내 편집 유지**를 묻는다.
  **원래대로**도 열려 있는 편집 창에 똑같이 알린다 (편집 창에 저장 안 한 내용이 있으면 파일이 이미 원본과 같아도 묻는다).
- 편집 창이 열려 있으면 상태 막대에 **편집 창 열림** (저장 안 한 내용이 있으면 **· 저장 안 됨**)이 보이고, 누르면 그 창으로 간다.
- 메인 창을 닫으면 편집 창도 닫는다 (저장 안 한 창은 먼저 묻고, 취소하면 둘 다 그대로).
- Word·Excel·PowerPoint 형식은 아래 **직접 편집 (Word·Excel·PowerPoint 형식)**.

### 직접 편집 (Word·Excel·PowerPoint 형식)
- docx·xlsx·pptx 문서를 연 뒤 상단 **직접 편집** (Ctrl+E) → 형식에 따라 **문서 편집 · 표 편집 · 슬라이드 편집** 창이 열린다 (문서마다 하나).
- 편집기는 앱에 들어 있다 (ONLYOFFICE Docs 기반, 아래 '라이선스'). **인터넷 없이**, Microsoft Office 없이 동작하고, 한글 편집 창처럼 앱 밖으로 나가는 요청을 모두 막는다.
- 글꼴은 이 PC에 설치된 것을 쓴다. 문서의 '맑은 고딕'·'바탕' 같은 한국어 이름도 찾는다. 첫 편집 창을 열 때 글꼴 목록을 만든다 (글꼴 미리보기 그림은 앱이 한가할 때 뒤에서).
- 저장 `Ctrl+S`: 원래 형식 그대로 (docx → docx). 덮어쓰기 전 내용은 백업 폴더에 남고, 왼쪽 미리보기도 바로 바뀐다.
  Word·Excel·PowerPoint가 경고 없이 열도록 저장할 때 몇 가지를 고친다 (pptx 노트 테마, 표에 더한 행의 세로 병합 표시 등).
  다른 형식(PDF 등)으로 내려받기는 저장 대화상자로, 인쇄는 PDF로 만들어 기본 PDF 앱으로 연다.
- 자동 저장: 편집기가 스스로 저장하려 해도 **원본 파일은 덮어쓰지 않고** 복구용 사본만 만든다.
  다음에 같은 파일을 열 때 사본이 더 새로우면 **자동 저장된 편집 내용이 있습니다 — 복구 / 버리기**를 묻는다.
- 바깥 변경(AI·다른 프로그램)·원래대로·닫기·AI 보내기 전 확인은 한글 편집 창과 같다.
- 위 막대: 파일 이름, 저장 상태, **저장**, **정보 · 라이선스** (편집기의 원 개발사, 라이선스, 이 판의 소스 주소).
- 직접 편집하지 않는 것: 암호 걸린 문서, 매크로 문서(docm·xlsm·pptm), 옛 형식(doc·xls·ppt·odt·ods·odp) → **프로그램에서 열기**.

### AI 대화
- 오른쪽 **AI 도우미** 창에 할 일을 쓴다. `Enter` 보내기, `Shift+Enter` 줄바꿈.
- 예: "표 2행 금액 합계 넣어줘", "{{담당자}}를 김선호로 채워줘", "이 문서 PDF로 저장해줘".
- AI는 Office Kit 문서 도구만 쓴다. 고친 결과는 미리보기에 바로 반영된다.
- AI는 **저장된 파일**을 고친다. 편집 창에 저장 안 한 내용이 있으면 보내기 전에 묻는다 (**편집 창 보기** / **그대로 보내기** / 취소).

### 형식별 전용 기능
찾아 바꾸기·칸 채우기 말고도 형식마다 고유한 기능을 AI와 명령줄에서 쓴다.
AI는 먼저 구조를 살펴보고(`…_inspect`) 번호를 확인한 뒤 고치고(`…_edit`), 쪽 그림(`office_page_image`)으로 결과를 확인한다.

| 형식 | 할 수 있는 것 | 필요한 것 |
|---|---|---|
| Word | 스타일·서식, 목록, 표·행, 그림·캡션, 머리말·꼬리말·쪽 번호, 쪽 설정, 메모, 변경 내용 추적, 목차, 각주, 링크 | 설치된 Word |
| PowerPoint | 슬라이드 추가·삭제·복제·순서, 제목·본문, 발표자 노트, 글상자·그림·표·도형·차트, 배경, 화면 전환, 테마, 슬라이드 크기 | 설치된 PowerPoint |
| Excel | 시트, 행·열, 서식·표시 형식, 너비·높이, 병합, 틀 고정, 정렬, 필터, 표, 차트, 조건부 서식, 입력 제한, 이름 정의 | Excel 우선. 없으면 내장 엔진(xlsx·csv, 정렬·차트·시트 복사·다시 계산 제외) |
| 한글 | 원하는 자리에 문단 넣기·지우기, 글자·문단 서식, 스타일, 표 칸·행, 머리말·꼬리말, 각주, 그림, 쪽 설정 | 없음 — 내장 엔진(rhwp) |

앱 AI 창에 이렇게 말해 보기:

| 형식 | 예 |
|---|---|
| Word | "제목들을 제목 1 스타일로 정리하고 맨 앞에 목차 넣어줘" · "'중요'를 모두 빨간 굵은 글씨로" · "꼬리말에 쪽 번호 넣어줘" |
| PowerPoint | "슬라이드마다 발표자 노트 달아줘" · "2번 슬라이드에 분기별 매출 막대 차트 넣어줘" · "모든 슬라이드에 페이드 전환 넣어줘" |
| Excel | "1행 굵게 하고 틀 고정해줘" · "A1:C13으로 꺾은선 차트 만들어줘" · "C열에서 100 넘는 칸 빨간색으로 표시해줘" |
| 한글 | "3번 문단 뒤에 '추진 일정' 문단 넣어줘" · "표 1 맨 아래에 합계 행 추가해줘" · "용지를 가로로 바꾸고 꼬리말에 쪽 번호 넣어줘" |

- 작업 여러 개를 순서대로 적용하고 **한 번만 저장**한다. 하나라도 실패하면 아무것도 저장하지 않는다.
- 앱의 첫 화면 추천 문장도 문서 종류에 맞춰 바뀐다.
- 작업 목록·인수·출력 한도·알려진 한계: [`docs/FEATURES.md`](docs/FEATURES.md).

### 미리보기
| 방식 | 한글 | Word·PowerPoint·Excel |
|---|---|---|
| 빠르게 (기본) | 내장 엔진(rhwp)으로 바로 그림 | 설치된 프로그램으로 PDF |
| 정확하게 | 한컴오피스로 PDF (보안 모듈 필요) | 설치된 프로그램으로 PDF |

프로그램이 없으면 글자만 보여 준다.
- 프로그램으로 만든 PDF는 기억해 두고 파일(크기·수정 시각)이 그대로면 다시 쓴다 → 다시 열 때 바로 보인다. 처음 만들 때는 걸린 시간을 상태 막대에 보여 준다.
- Excel은 미리보기·AI 쪽 그림(`office_page_image`)에서 **시트마다 한 쪽 너비**에 맞춘다 (넓은 표·차트가 여러 쪽으로 잘리지 않게). 임시 사본에서만 바꾸고 파일의 인쇄 설정은 그대로다 — **PDF로 저장**은 문서의 인쇄 설정대로.
- 미리보기(내장 엔진)는 표 다음 문단을 표의 바깥 아래 여백만큼 띄운다 (한글과 같게, 파일은 그대로).

### PDF로 저장
상단 **PDF로 저장**, AI에게 "PDF로 저장해줘", MCP `office_convert`, 명령줄 `office-kit convert 보고서.hwp 보고서.pdf` 모두 같은 규칙이다.

| 형식 | 프로그램이 있으면 | 없으면 |
|---|---|---|
| 한글 (hwp·hwpx·hwt) | 한컴오피스 + 자동화 보안 모듈로 (원본 그대로) | **내장 엔진(rhwp)으로** — 결과에 "엔진: 내장(rhwp)" 표시 |
| Word·PowerPoint·Excel | 그 프로그램으로 | 만들 수 없음 |

내장 엔진 PDF (`vendor/rhwp-cli/rhwp.exe`, rhwp v0.8.7 명령줄 프로그램, MIT):
- **인터넷 없이**, 한컴오피스·Chromium 없이 만든다. 앱 미리보기와 같은 rhwp 조판 엔진(같은 판)으로 그려 **미리보기와 같은 모양**이다.
- 글자는 PDF 글꼴로 들어간다 (이 PC 글꼴에서 쓴 글자만, 예: 함초롬바탕·맑은 고딕) → **글자 선택·검색이 된다**. 쪽 크기·가로 방향·머리말·꼬리말·쪽 번호·표·그림 그대로.
- 한컴오피스 PDF와 다른 점: 글꼴이 PC에 없으면 비슷한 글꼴(바탕·맑은 고딕)로 바뀌고, 줄바꿈·쪽 나눔이 조금 다를 수 있다.
  글자를 한 자씩 놓아서 PDF 뷰어에 따라 복사한 글에 띄어쓰기가 빠질 수 있다. 큰 문서는 느리다 (약 0.5초/쪽, 77쪽 문서 약 40초).
- 빈 누름틀의 안내문(미리보기의 빨간 글)은 PDF에 나오지 않는다 (한컴오피스 인쇄와 같음). 채운 값은 나온다.
- 원본 파일은 건드리지 않는다 (미리보기용 사본을 임시 폴더에서 PDF로 만든 뒤 고른 자리에 옮긴다).

---

## AI 연결

Office Kit는 **로그인 정보를 다루지 않는다.** 내 PC에 설치·로그인한 CLI를 실행할 뿐이다.

| AI | 준비 | 로그인 |
|---|---|---|
| Claude | [Claude Code](https://docs.anthropic.com/claude-code) 설치 | 터미널에서 `claude` 실행 → 내 Claude 계정으로 로그인 |
| Codex | [Codex CLI](https://github.com/openai/codex) 설치 (`npm i -g @openai/codex`) | `codex login` → 내 ChatGPT 계정으로 로그인 |

- 앱 **설정 → AI 프로그램** 에서 어느 쪽을 쓸지 고른다. `다시 확인` 으로 설치 여부를 다시 찾는다.
- 모델 칸을 비우면 각 CLI의 기본 모델을 쓴다.

---

## Claude Desktop · Claude Code · Codex에서 쓰기 (MCP)

Office Kit 문서 도구를 다른 AI 앱에서도 쓰게 연결한다.

**방법 1 — 앱에서:** 설정 → **외부 연결** → 원하는 항목 켜기 → 그 프로그램 다시 시작.

**방법 2 — 명령줄에서:**

```powershell
office-kit setup claude-desktop     # %APPDATA%\Claude\claude_desktop_config.json 에 추가
office-kit setup claude-code        # claude mcp add --scope user office-kit ...
office-kit setup codex              # ~/.codex/config.toml 에 [mcp_servers.office-kit] 추가
office-kit setup codex --remove     # 연결 해제 (다른 대상도 같음)
```

| 대상 | 바뀌는 곳 | 안전장치 |
|---|---|---|
| Claude Desktop | `claude_desktop_config.json` 의 `mcpServers.office-kit` | 다른 설정 보존, 고칠 때마다 백업, 깨진 JSON은 안 건드림 |
| Claude Code | `claude mcp add --scope user` (사용자 범위) | Claude Code CLI가 직접 저장. 먼저 쓰이는 local·project 항목은 지우고, 끝나면 실제로 쓰이는 항목을 다시 확인 |
| Codex | `config.toml` 의 `[mcp_servers.office-kit]` 표 (바로 위 안내 줄 `# office-kit: …`) | 그 표 밖은 안 바꿈 (Codex가 나중에 붙인 표도 그대로), 고칠 때마다 백업, 해석 못 하는 파일은 안 건드림, `CODEX_HOME` 존중 |

- 백업: 같은 폴더에 `<설정 파일>.office-kit-날짜-시각.bak` (Office Kit가 만든 것만 최근 10개 유지).
- 연결 상태는 값(실행 파일·인수·환경 변수)으로 비교한다. Codex가 `codex mcp add` 등으로 표 모양을 바꿔도 그대로 연결됨으로 본다.

수동으로 넣을 때 (Claude Desktop 예):

```json
{
  "mcpServers": {
    "office-kit": {
      "command": "C:\\Users\\<나>\\AppData\\Local\\Programs\\Office Kit\\Office Kit.exe",
      "args": ["C:\\Users\\<나>\\AppData\\Local\\Programs\\Office Kit\\resources\\app\\src\\mcp\\server.mjs"],
      "env": { "ELECTRON_RUN_AS_NODE": "1" }
    }
  }
}
```

도구 목록: `office_status` `office_info` `office_read` `office_fields` `office_fill` `office_replace` `office_set_cells` `office_convert` `office_create` `office_page_image`
· 형식별: `hwp_inspect` `hwp_edit` `word_inspect` `word_edit` `ppt_inspect` `ppt_edit` `excel_inspect` `excel_edit`

> PC에 Node.js가 없어도 된다. 설치된 `Office Kit.exe` 가 `ELECTRON_RUN_AS_NODE=1` 로 Node처럼 서버를 실행한다.

---

## 명령줄

설정 → 외부 연결 → **명령줄** 을 켜면 (또는 `office-kit setup cli`) 어디서나 `office-kit` 을 쓸 수 있다.
`%LOCALAPPDATA%\Microsoft\WindowsApps\office-kit.cmd` 가 생긴다 (이미 PATH에 있는 폴더). 새 터미널부터 적용.
`office-kit setup cli` 는 **일반 터미널**에서 실행한다. 스토어판 Claude Desktop 같은 MSIX 앱 안의 터미널·에이전트에서 실행하면
파일이 그 앱 전용 폴더에 저장돼 다른 터미널에서는 보이지 않는다 (이때는 안내가 뜬다).

| 명령 | 하는 일 |
|---|---|
| `office-kit status` | 설치된 프로그램 · 한글 보안 모듈 · 연결 상태 |
| `office-kit text 보고서.hwp` | 글자 출력 |
| `office-kit fields 신청서.hwpx` | 채울 칸 목록 |
| `office-kit fill 신청서.hwpx 값.json -o 채움.hwpx` | 칸 채우기 |
| `office-kit replace 계약서.docx 갑 을` | 찾아 바꾸기 |
| `office-kit cells 표.xlsx 셀.json` | 엑셀 셀 쓰기 |
| `office-kit convert 보고서.hwp 보고서.pdf` | 형식 변환 |
| `office-kit new 원고.md 새문서.hwpx` | 마크다운 원고로 새 문서 |
| `office-kit image 보고서.hwp 2 -o 2쪽.png` | 한 쪽을 그림으로 |
| `office-kit inspect 발표.pptx --slide 2` | 형식별 구조 보기 (문단·표·슬라이드·도형·시트 번호) |
| `office-kit edit 보고서.docx 작업.json -o 수정.docx` | 형식별 편집 (작업 목록을 한 번에 적용) |
| `office-kit help edit` | 형식별 편집 작업 목록 |
| `office-kit help` | 전체 도움말 |

공통: `-o` 출력 경로(생략하면 원본 덮어씀), `--engine auto|app|file`. `edit`는 `-o`에 다른 파일이 있으면 `--overwrite`가 있어야 덮어쓴다.

`작업.json` 예 (Word):

```json
[
  { "op": "insert_toc", "at": "start", "title": "목차" },
  { "op": "format_text", "find": "중요", "all": true, "bold": true, "color": "#c00000" },
  { "op": "set_footer", "text": "- {page} / {pages} -", "align": "center" }
]
```
예제 파일: [`examples/`](examples/) (신청서.hwp/hwpx, 값.json, 신청서.md).

---

## 한글 자동화 보안 모듈

한컴오피스를 프로그램으로 조작하면 파일마다 **"접근 허용" 창**이 떠서 자동 작업이 멈춘다.
한컴이 배포하는 **보안 승인 모듈**(FilePathCheckerModuleExample.dll)을 등록하면 창 없이 동작한다.

| 상태 | 한글 문서 처리 |
|---|---|
| 모듈 없음 (기본) | 내장 엔진(rhwp)으로 읽기·쓰기·**PDF 저장**. 정확한 미리보기(한컴 PDF)는 불가 |
| 모듈 등록됨 | 한컴오피스로 처리. PDF도 한컴으로 (원본 그대로), 정확한 미리보기 가능 |

등록하기:
- 앱: 설정 → **한글 자동화** → **한글 자동화 연결** (한컴 공식 배포본을 받아 등록) 또는 **직접 DLL 선택…**
- 명령줄: `office-kit setup hancom --download` 또는 `office-kit setup hancom --dll <경로>`
- 해제: `office-kit setup hancom --remove`

알아 둘 것:
- 받는 곳: [한컴 개발자센터](https://developer.hancom.com/hwpautomation)가 안내하는 공식 저장소 `github.com/hancom-io/devcenter-archive` 의 `보안모듈(Automation).zip`. 이 파일은 Office Kit에 포함돼 있지 않고, **버튼을 눌렀을 때만** 받는다.
- 설치된 한컴오피스의 비트 수(대개 32비트)에 맞는 DLL을 고른다.
- 레지스트리: `HKCU\Software\HNC\HwpAutomation\Modules` 의 `FilePathCheckerModule` = DLL 전체 경로 (따옴표 없이).

---

## 개발

필요: Windows 10/11, Node.js 24, Git.

```powershell
git clone https://github.com/toxh13/office-kit.git
cd office-kit
npm install
npm run editor:build   # 앱 안 한글 편집기 묶음 만들기 (처음 한 번, 아래 참고)
npm run office-editor:build   # 앱 안 문서·표·슬라이드 편집기 묶음 만들기 (처음 한 번, 약 100MB 받음)
npm run rhwp-cli:build   # 한글→PDF 변환 프로그램 받기 (처음 한 번, 약 19MB 받음)
npm start          # 앱 실행 (Electron)
npm test           # 시험 (node --test)
npm run dist       # 설치 파일 → dist/OfficeKit-Setup-<버전>.exe (편집기 묶음도 먼저 만든다. 묶음 없이 electron-builder를 바로 돌리면 멈춘다)
npm run icon       # build/icon.ico · icon.png 다시 그리기
```

**오프라인 한글 편집기 묶음** (`npm run editor:build` = `node scripts/build-editor.mjs`)
- rhwp 저장소의 고정 태그 `v0.8.7`(커밋까지 확인)에서 rhwp-studio를 받아 Rust 없이 빌드한다. `@rhwp/core` 0.8.7의 WASM을 그대로 쓴다.
- 결과: `vendor/rhwp-studio/`(편집기, 약 22MB) · `vendor/rhwp-editor-sdk/`(연결 SDK). `vendor/`는 만들어지는 폴더라 git에 넣지 않는다.
- 서비스 워커·샘플 문서는 빼고, 바깥 글꼴 주소는 끈 채로 빌드한다. 받은 소스·의존성 캐시는 `tmp/editor-build` (`OFFICEKIT_EDITOR_CACHE`로 바꿈).
- 이미 같은 것이 있으면 건너뛴다. `--force` 다시 빌드, `--with-fonts` 오픈 라이선스 웹 글꼴(약 19MB)을 함께 넣음 (기본은 시스템 글꼴).
- 묶음 없이 `npm start`하면 앱은 그대로 쓰고, **직접 편집**만 "한글 편집기를 열 수 없어요" 안내를 띄운다.
- rhwp 버전은 함께 올린다: `package.json`의 `@rhwp/core`(정확한 버전)·`@rhwp/editor`, `scripts/build-editor.mjs`의 `TAG`·`COMMIT`·`CORE_VERSION`·`EDITOR_VERSION`.

**한글→PDF 변환 프로그램** (`npm run rhwp-cli:build` = `node scripts/build-rhwp-cli.mjs`)
- rhwp v0.8.7 GitHub Release의 `rhwp-v0.8.7-windows-x86_64.zip`을 받아 크기·SHA-256을 확인하고, `rhwp.exe`·`LICENSE`(역시 고정 SHA-256)와
  같은 태그 커밋의 `THIRD_PARTY_LICENSES.md`를 `vendor/rhwp-cli/`에 넣는다 (약 47MB). `rhwp.exe --version`도 확인한다.
- `node_modules/@rhwp/core`가 정확히 0.8.7이 아니면 멈춘다 (PDF와 미리보기가 같은 엔진 판이어야 같은 모양).
- `--verify` 고정 값 그대로인지 확인만, `--force` 다시 만듦, `--offline` 캐시만. 받은 파일 캐시: `tmp/rhwp-cli-build/download` (`OFFICEKIT_RHWP_CLI_CACHE`로 바꿈).
- 없이 `npm start`하면 한컴오피스(+보안 모듈)가 없는 PC에서 한글 PDF 저장만 "내장 PDF 변환 프로그램이 없습니다"를 알린다.
  다른 위치의 rhwp를 쓰려면 `OFFICEKIT_RHWP_CLI`, 시간 제한(기본 60초 + 쪽당 5초, 최대 20분)은 `OFFICEKIT_RHWP_TIMEOUT_MS`.

**오프라인 문서·표·슬라이드 편집기 묶음** (`npm run office-editor:build` = `node scripts/build-office-editor.mjs`)
- filex-office-editor v0.3.1 릴리스의 `ui.zip`(ONLYOFFICE Docs CE 9.4.0.129 편집기 + x2t 변환기)을 받아 크기·SHA-256과 파일마다의 SHA-256을 확인한다.
  `patches/office-editor/filex/`에 패치가 있으면 filex 소스(고정 커밋)에 얹어 앱 페이지만 다시 빌드한다 (Docker 필요 없음).
- 쓰지 않는 파일(다른 나라 말, 휴대폰용 화면, 묶음 글꼴)을 **통째로만** 빼고 `vendor/office-editor/`에 쓴다 (약 213MB, 파일 약 1,480개). 남은 파일은 고치지 않는다.
- 고정 값: `scripts/office-editor/pins.mjs`. 받은 파일 캐시: `tmp/office-editor-build` (`OFFICEKIT_OFFICE_EDITOR_CACHE`로 바꿈).
- 묶음 없이 `npm start`하면 앱은 그대로 쓰고, Word·Excel·PowerPoint 형식의 **직접 편집**만 이유를 알린다.
- 소스 꾸러미(AGPL): `node scripts/make-office-source.mjs` → `dist/office-editor-source-<버전>.zip` ([`docs/LICENSES-OFFICE-EDITOR.md`](docs/LICENSES-OFFICE-EDITOR.md)).

| 폴더 | 내용 |
|---|---|
| `src/core/` | 문서 엔진 (UI 없음) |
| `src/mcp/` | MCP 서버 (stdio) |
| `src/cli.mjs` | 명령줄 도구 |
| `src/setup/` | 외부 연결 · 명령줄 등록 · 한글 보안 모듈 |
| `src/app/` | Electron 앱 (main, preload, renderer, AI 중계, 한글·Office 편집 창) |
| `scripts/office-editor/`, `patches/office-editor/` | Office 편집기 묶음의 고정 값·덜어 내기 규칙, filex 패치 (AGPL) |
| `src/mcp/tools/` | 형식별 MCP 도구 (`hwp`·`word`·`ppt`·`excel`) |
| `scripts/com/` | 설치된 프로그램을 COM으로 조작 (`office.ps1` + 형식별 `word-ops.ps1`·`ppt-ops.ps1`·`excel-ops.ps1`) |
| `build/` | 아이콘, NSIS 추가 설정(`installer.nsh`) |

구조와 모듈 약속: [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)

시험은 실제 설정 파일·레지스트리·네트워크를 건드리지 않는다 (임시 폴더와 가짜 실행기 사용).

---

## 배포

`v*` 태그를 올리면 GitHub Actions가 설치 파일을 만들어 Release에 올린다.

```powershell
npm version 0.1.0 --no-git-tag-version   # package.json 버전 = 태그 버전이어야 함
git commit -am "v0.1.0"
git tag v0.1.0
git push origin main v0.1.0
```

- 워크플로: [`.github/workflows/release.yml`](.github/workflows/release.yml) — 시험(먼저 `npm run rhwp-cli:build`) → `npm run editor:build` → `npm run office-editor:build` → `npm run rhwp-cli:build` → `electron-builder --publish always` (`GH_TOKEN` = 기본 `GITHUB_TOKEN`)
  → Office 편집기 소스 꾸러미 `office-editor-source-<버전>.zip`을 같은 Release에 붙인다.
  설치 파일에는 `vendor/rhwp-studio`·`vendor/rhwp-editor-sdk`(설치 파일 약 7MB, 설치 폴더 약 23MB 늘어남)와
  `vendor/office-editor`(설치 파일 약 50~65MB, 설치 폴더 약 213MB 늘어남), `vendor/rhwp-cli`(설치 파일 약 14MB, 설치 폴더 약 47MB 늘어남)가 들어간다.
- **남에게 배포하기 전에** [`docs/LICENSES-OFFICE-EDITOR.md`](docs/LICENSES-OFFICE-EDITOR.md)의 체크리스트 (법률 검토, 거울 저장소, 소스 꾸러미 확인).
- 수동 실행(Actions → CI · Release → Run workflow)은 초안(draft) Release로 올린다.
- main push / PR 마다 `npm test` 와 편집기 빌드 스크립트 확인이 돌고, main push에서는 편집기 묶음도 실제로 빌드해 본다 (태그 전에 깨진 것을 알게).

---

## 알려진 한계

- **Windows 전용.** COM 자동화와 설치 파일이 Windows 기준이다.
- **코드 서명 없음.** 처음 실행할 때 SmartScreen 경고가 뜬다.
- **한컴 보안 모듈이 없으면** 한글 문서의 정확한 미리보기(한컴 PDF)가 안 된다. 읽기·쓰기·PDF 저장은 내장 엔진으로 된다 —
  내장 엔진 PDF는 한컴오피스 PDF와 글꼴·줄바꿈·쪽 나눔이 조금 다를 수 있고, 큰 문서는 느리다 (위 'PDF로 저장').
- 한글 이외의 PDF 변환(docx·pptx·xlsx → pdf)과 한글 → docx·odt·txt·html 변환은 해당 프로그램이 있어야 한다.
- 내장 엔진(rhwp) 미리보기는 한컴오피스와 글꼴·줄바꿈이 조금 다를 수 있다.
- 앱 안 한글 편집기(rhwp-studio)는 한컴오피스와 같지 않다: 기본 설치에는 웹 글꼴이 없어 PC에 있는 글꼴로 보이고, 복잡한 개체·서식은 다르게 보이거나 저장될 수 있다.
  중요한 문서는 저장한 뒤 미리보기(또는 한컴오피스)로 확인한다. `.hwt`(서식 파일)는 저장할 때 HWP 내용으로 쓴다.
  병합된 칸이 많은 표에 줄·칸을 넣으면 표 모양이 흐트러질 수 있어요 (Ctrl+Z로 되돌리기). 복잡한 서식 문서는 **프로그램에서 열기**로 고치세요.
  일부 문서는 편집 창에서 표 바로 다음 문단이 표 마지막 줄에 겹쳐 보인다 (rhwp-studio 0.8.7이 표의 바깥 아래 여백을 무시함 — 저장한 파일은 그대로, 미리보기는 Office Kit가 그 여백만큼 띄워 보여 줌).
  편집 창 아래 상태 줄의 '… 2페이지 (239.0ms)' 같은 그리기 시간은 rhwp-studio가 쓰는 글이다.
- 앱 안 문서·표·슬라이드 편집기는 Word·Excel·PowerPoint와 같지 않다. 저장하면 파일 꾸러미를 다시 쓴다 (`styles.xml`이 커짐, 한글·영문 사이 간격이 달라 보임 등).
  변경 추적·메모·차트·SmartArt·필드·포함 개체가 든 중요한 문서는 저장한 뒤 미리보기(또는 그 프로그램)로 확인하거나 **프로그램에서 열기**로 고친다.
  편집기 화면은 원래 모양·로고 그대로이고 일부 안내(새 기능 풍선 등)는 영어다. 함께 편집·맞춤법 검사는 없다.
  편집기 상태 막대의 '모든 변경 사항이 저장되었습니다'는 편집기 안 기록 이야기다 — 파일 저장 여부는 위 막대를 본다 (편집기 파일을 고치지 않고는 숨길 수 없음).
  새 기능 풍선('신규')은 끈다. 기존 문서의 글자 언어는 문서에 적힌 대로 보인다 (Office Kit가 만든 새 문서는 한국어).
  편집 창마다 메모리를 약 0.5GB 쓴다.
- doc · ppt · xls · rtf · odt · odp · ods 는 해당 프로그램이 있어야 열고 고칠 수 있다 (csv는 읽기와 `excel_edit`의 값 작업을 내장 엔진으로도 한다).
- Word·PowerPoint 전용 기능(`word_*`·`ppt_*`)은 해당 프로그램이 있어야 한다. PowerPoint 차트 넣기는 작업 동안 최소화된 PowerPoint 창이 잠깐 열린다.
- 형식별 기능의 작업마다 한계는 [`docs/FEATURES.md`](docs/FEATURES.md) 각 절 끝에 적었다.
- AI 대화는 Claude Code 또는 Codex CLI가 설치·로그인돼 있어야 한다.
- 자동 업데이트는 아직 없다. 새 버전은 Releases에서 받아 덮어 설치한다.
- Office Kit를 지워도 Claude Desktop·Codex 설정의 `office-kit` 항목은 남는다 → 지우기 전에 설정 → 외부 연결에서 끄기.

---

## 다음 할 일

- [x] 첫 배포: `v0.1.0` 태그 push, Release에 설치 파일 올라오는지 확인
- [x] 실제 PC에서 설치·제거, 화면 조작, Claude Desktop·Claude Code·Codex·명령줄 연결 확인 (v0.1.1에 반영)
- [ ] Codex로 앱 안 AI 대화 끝까지 확인 (계정 사용 한도가 풀린 뒤)
- [ ] 코드 서명 인증서 적용 (SmartScreen 경고 없애기)
- [ ] 자동 업데이트 (electron-updater + GitHub Release)
- [ ] 한컴 보안 모듈 받기를 실제 한컴오피스 2022/2024에서 끝까지 확인
- [ ] 설치/제거 시험 자동화 (설치 파일을 CI에서 조용히 설치 → 오른쪽 클릭 메뉴 확인 → 제거)
- [ ] 제거할 때 Claude Desktop·Codex 연결도 함께 정리
- [ ] 미리보기 글꼴 차이 줄이기 (함초롬 계열 글꼴 안내)
- [ ] 편집기 웹 글꼴을 설치 파일에 넣을지 결정 (`npm run editor:build -- --with-fonts`, 설치 파일 약 19MB 더 큼)
- [ ] 영어 README
- [ ] Office 편집기: 첫 배포 전 법률 검토, 위쪽 소스 거울 저장소 만들기 ([`docs/LICENSES-OFFICE-EDITOR.md`](docs/LICENSES-OFFICE-EDITOR.md) 4·9절)
- [ ] Office 편집기: 실제 MS IME(두벌식·세벌식·한자 변환) 사람 시험

## 라이선스

[MIT](LICENSE) © toxh13. 함께 배포하는 오픈소스(rhwp·rhwp-studio·@rhwp/editor — MIT, Edward Kim 외 / CanvasKit — BSD-3-Clause / @noble/hashes — MIT)는
[LICENSE](LICENSE)의 THIRD-PARTY 항목과 `vendor/rhwp-studio/THIRD_PARTY_LICENSES.md` 참고.
한글→PDF 변환 프로그램 `vendor/rhwp-cli/rhwp.exe`는 rhwp v0.8.7 Release의 바이너리 그대로다 (MIT, `vendor/rhwp-cli/`에 LICENSE·THIRD_PARTY_LICENSES.md 동봉).

**예외 — AGPL인 것**: 문서·표·슬라이드 편집 창의 편집기 묶음 `vendor/office-editor/`와, 이 저장소에서 머리에
`SPDX-License-Identifier: AGPL-3.0-or-later`·`AGPL-3.0-only`가 붙은 파일(`patches/office-editor/`의 filex 패치 — AGPL-3.0-or-later, `src/app/office-editor/fonts/`의 글꼴 목록 생성기 — AGPL-3.0-only + ONLYOFFICE 추가 조항)은 MIT가 아니라 **GNU AGPL 3판**이다. 전문: [`LICENSES/AGPL-3.0.txt`](LICENSES/AGPL-3.0.txt), 추가 조항: [`LICENSES/ONLYOFFICE-additional-terms.txt`](LICENSES/ONLYOFFICE-additional-terms.txt).

- 편집기는 **ONLYOFFICE Docs** Community Edition 9.4.0(빌드 9.4.0.129, Ascensio System SIA)을 바탕으로 한 **수정판**이다 —
  AGPL-3.0 + 추가 조항, 그림·아이콘은 CC BY-SA 4.0. 바꾼 것과 날짜: `vendor/office-editor/NOTICE-OfficeKit.md`.
- 그 편집기를 Document Server 없이 돌리는 앱 페이지·다리는 [filex-office-editor](https://github.com/BRF-Tech/filex-office-editor) v0.3.1 (BRF Tech, AGPL-3.0-or-later), 변환기 x2t는 ONLYOFFICE core + CryptPad·BRF Tech 패치 (AGPL).
- 변환기 x2t·글꼴 엔진(WebAssembly) 안 제3자 코드(Boost, ICU, FreeType, libpng, libjpeg, Xpdf 등)와 x2t 안 URW++ PDF 기본 14 글꼴(AGPL-3.0 + PDF 포함 예외)의 고지: `vendor/office-editor/THIRD-PARTY-NATIVE.txt` ([`LICENSES/office-editor-third-party-native.txt`](LICENSES/office-editor-third-party-native.txt)). 편집용 글꼴 파일은 넣지 않는다 (이 PC의 글꼴을 읽는다).
- 소스: Release마다 붙는 `office-editor-source-<버전>.zip` + ONLYOFFICE 소스 태그 `v9.4.0.129`. 앱 안에서는 편집 창 위 막대 **정보 · 라이선스**.
- Office Kit 나머지는 MIT 그대로다 (편집기와는 메시지로만 주고받는 별개 프로그램). 자세한 경계·배포 전 할 일: [`docs/LICENSES-OFFICE-EDITOR.md`](docs/LICENSES-OFFICE-EDITOR.md).
- ONLYOFFICE®·ONLYOFFICE Docs™는 Ascensio System SIA의 상표다. Office Kit은 Ascensio System SIA와 관계가 없고 보증·후원을 받지 않는다.
  Word·Excel·PowerPoint는 Microsoft의 상표이며 파일 형식을 가리킬 때만 쓴다.
