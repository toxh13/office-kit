<p align="center"><img src="build/icon.png" width="96" alt="Office Kit"></p>

# Office Kit

한글(HWP·HWPX)·Word·PowerPoint·Excel 문서를 **열고, 확인하고, AI 대화로 고치는** Windows 데스크톱 도구.

## 한눈에

| 무엇을 | 어떻게 |
|---|---|
| 문서 보기 | 창에 끌어다 놓기 → 왼쪽에 미리보기 |
| 문서 고치기 | 오른쪽 AI 창에 "담당자 칸에 홍길동 넣어줘"처럼 말하기 |
| AI | 내 PC에 로그인해 둔 **Claude Code** 또는 **Codex**를 그대로 씀 (앱은 로그인 정보를 안 만짐) |
| 다른 AI 앱에서 | Claude Desktop·Claude Code·Codex에 문서 도구(MCP) 연결 |
| 명령줄 | `office-kit text 보고서.hwp` |
| 설치 | [Releases](https://github.com/toxh13/office-kit/releases)에서 `OfficeKit-Setup-<버전>.exe` 받아 실행 |

- 한컴오피스·MS Office가 **없어도** 한글/Word/PowerPoint/Excel 파일을 읽고 고칠 수 있다 (내장 엔진).
- 프로그램이 **있으면** 그 프로그램으로 원본 그대로 처리하고 PDF도 만든다.

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

- 남는 것: 설정·최근 파일·**문서 백업**(연 문서의 사본, 최근 50개) `%APPDATA%\office-kit`, AI 작업 폴더 `%LOCALAPPDATA%\office-kit`. 필요 없으면 직접 지운다.
- Office Kit로 연 문서는 Windows "최근 항목"에 남고, 탐색기의 그 형식 "연결 프로그램" 목록에 Office Kit가 보일 수 있다 (기본 연결 프로그램은 그대로).

---

## 사용

### 문서 열기
- 창에 **끌어다 놓기**, 또는 `파일 열기`, 또는 탐색기에서 오른쪽 클릭 → **Office Kit로 열기**.
- 열 때마다 원본을 자동 백업한다 (최근 50개). 상단 **원래대로** 로 되돌린다.

### AI 대화
- 오른쪽 **AI 도우미** 창에 할 일을 쓴다. `Enter` 보내기, `Shift+Enter` 줄바꿈.
- 예: "표 2행 금액 합계 넣어줘", "{{담당자}}를 김선호로 채워줘", "이 문서 PDF로 저장해줘".
- AI는 Office Kit 문서 도구만 쓴다. 고친 결과는 미리보기에 바로 반영된다.

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
| 모듈 없음 (기본) | 내장 엔진(rhwp)으로 읽기·쓰기. PDF 변환·정확한 미리보기는 불가 |
| 모듈 등록됨 | 한컴오피스로 처리. PDF 변환·정확한 미리보기 가능 |

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
npm start          # 앱 실행 (Electron)
npm test           # 시험 (node --test)
npm run dist       # 설치 파일 → dist/OfficeKit-Setup-<버전>.exe
npm run icon       # build/icon.ico · icon.png 다시 그리기
```

| 폴더 | 내용 |
|---|---|
| `src/core/` | 문서 엔진 (UI 없음) |
| `src/mcp/` | MCP 서버 (stdio) |
| `src/cli.mjs` | 명령줄 도구 |
| `src/setup/` | 외부 연결 · 명령줄 등록 · 한글 보안 모듈 |
| `src/app/` | Electron 앱 (main, preload, renderer, AI 중계) |
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

- 워크플로: [`.github/workflows/release.yml`](.github/workflows/release.yml) — 시험 → `electron-builder --publish always` (`GH_TOKEN` = 기본 `GITHUB_TOKEN`).
- 수동 실행(Actions → CI · Release → Run workflow)은 초안(draft) Release로 올린다.
- main push / PR 마다 `npm test` 가 돈다.

---

## 알려진 한계

- **Windows 전용.** COM 자동화와 설치 파일이 Windows 기준이다.
- **코드 서명 없음.** 처음 실행할 때 SmartScreen 경고가 뜬다.
- **한컴 보안 모듈이 없으면** 한글 문서의 PDF 변환·정확한 미리보기가 안 된다 (읽기·쓰기는 내장 엔진으로 됨).
- 내장 엔진(rhwp) 미리보기는 한컴오피스와 글꼴·줄바꿈이 조금 다를 수 있다.
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
- [ ] 영어 README

## 라이선스

[MIT](LICENSE) © toxh13. 함께 배포하는 오픈소스(rhwp — MIT, Edward Kim 외)는 [LICENSE](LICENSE)의 THIRD-PARTY 항목 참고.
