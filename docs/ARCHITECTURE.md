# Office Kit 구조

한글(HWP/HWPX)·Word·PowerPoint·Excel 문서를 열고, 확인하고, AI 대화로 고치는 Windows 데스크톱 도구.

## 한눈에

```
src/core/      문서 엔진 (UI 없음, Node만 있으면 동작)
src/mcp/       MCP 서버 — Claude/Codex가 문서 도구를 쓰게 함 (stdio)
src/cli.mjs    명령줄 도구 office-kit
src/setup/     외부 연결: Claude Desktop/Claude Code/Codex 등록, 명령줄 등록, 한글 보안 모듈
src/app/       Electron 앱 (main.mjs, preload.cjs, renderer/, ai.mjs)
scripts/com/   office.ps1 — 설치된 프로그램을 COM으로 조작
build/         아이콘, NSIS 추가 설정
```

## 엔진 두 가지

| 엔진 | 쓰는 것 | 장점 | 언제 |
|---|---|---|---|
| app | 설치된 한컴/Word/PowerPoint/Excel (COM, `scripts/com/office.ps1`) | 원본 그대로, PDF 변환 | 쓰기 작업 기본값 |
| file | rhwp(WASM)·ZIP/XML·exceljs | 프로그램 불필요, 빠름 | 읽기 기본값, 프로그램 없을 때 |

- 한컴은 **자동화 보안 모듈**(`HKCU\Software\HNC\HwpAutomation\Modules` 의 `FilePathCheckerModule`)이 없으면
  파일마다 "접근 허용" 창이 떠서 멈춘다 → 모듈 없으면 자동 작업은 rhwp로 처리 (`comAvailableFor(ext, {interactive})`).
- COM 호출은 `src/core/com.mjs` 큐로 프로세스 안에서 한 번에 하나씩, `office.ps1`의 이름 있는 뮤텍스
  (`Local\OfficeKitCom`)로 **프로세스끼리도** 한 번에 하나씩 실행한다.
- 자동화로 여는 Office 문서는 매크로를 막는다 (`AutomationSecurity = 3`, 사람이 보는 창은 2).
- 시간 초과 시 `office.ps1`이 적어 둔 PID로 남은 Office 프로세스를 정리한다.

## 환경 변수

| 이름 | 쓰는 곳 | 뜻 |
|---|---|---|
| `OFFICEKIT_WRITE_ROOTS` | MCP 서버 | 저장을 허락할 폴더들 (`;` 구분). 앱 안 AI 대화는 열린 문서의 폴더만 넘긴다 |
| `OFFICEKIT_POWERSHELL` | core/com | 시험용 PowerShell 경로 바꿔 끼우기 |
| `OFFICEKIT_SMOKE`, `_CHAT`, `_SETTINGS` | 앱 | 화면 캡처 확인 모드 |
| `OFFICEKIT_USER_DATA` | 앱 | 설정·백업 폴더 바꾸기 (시험용) |
| `OFFICEKIT_AI_MODULE` | 앱(개발용) | 가짜 AI 모듈로 바꿔 끼우기 |
| `OFFICEKIT_COM_TESTS=1` | 테스트 | Word/PowerPoint/Excel COM 시험 켜기 (한컴 제외) |
| `OFFICEKIT_DEBUG` | CLI/MCP | 오류 스택 출력 |

## 모듈 계약

### src/core/index.mjs (있음)
`readDocument(path, {engine})`, `documentInfo(path)`, `replaceText(path, pairs, {output, engine})`,
`listFields(path)`, `fillFields(path, values, {output, engine})`, `setCells(path, cells, {sheet, output, engine})`,
`convertDocument(path, output)`, `renderPreview(path, outDir, {mode})`, `openInApp(path)`,
`createDocument(output, markdown)`, `probe()`, `Hwp`, `OPENABLE_EXTS`, `extOf`, `kindOf`.
- `output` 생략 = 원본 덮어쓰기. engine: `'auto' | 'app' | 'file'`.
- `renderPreview` → `{kind:'pdf', file}` | `{kind:'svg', pages}` | `{kind:'text', docKind, ...readDocument}`.
  프로그램 PDF는 임시 폴더에 받은 뒤 `outDir`로 옮긴다 (MSIX 앱 안에서 실행되면 가상화된 `%APPDATA%`를 Office가 못 본다).
- `replaceText` 결과에 쌍마다의 개수 `counts`가 있다. csv는 파일 엔진으로도 읽는다. `.hwt`(한글 서식) 지원.
- 모든 오류 메시지는 한국어.

### src/core/render.mjs
`pageImage(path, page = 1, { dpi }) → { png: Buffer, page, pages, engine }` — AI가 결과를 눈으로 확인하는 용도.
- hwp/hwpx: rhwp SVG → `@resvg/resvg-js` (시스템 글꼴).
- 그 외: 프로그램으로 PDF 변환 → `office.ps1`의 `app:'pdf', op:'pageImage'` (Windows.Data.Pdf)로 PNG.
- `.pdf` 입력도 바로 처리.

### src/mcp/server.mjs
stdio MCP 서버 `office-kit`. stdout은 프로토콜 전용, 로그는 stderr.
도구: `office_status`, `office_info`, `office_read`, `office_fields`, `office_fill`, `office_replace`,
`office_set_cells`, `office_convert`, `office_create`, `office_page_image`, `hwp_edit`.
- 쓰기 도구는 기본이 제자리 저장. 다른 이름(`output`)이 이미 있으면 `overwrite=true`일 때만 덮어쓴다.
- 저장 가능한 확장자만 허용하고, `OFFICEKIT_WRITE_ROOTS`가 있으면 그 폴더 밖에는 쓰지 않는다.
- `office_read`는 `offset`/`max_chars`로 나눠 읽는다 (잘라 버리지 않고 다음 위치를 알려 준다).

### src/app/ai.mjs
```js
detectProviders() → Promise<[{ id: 'claude'|'codex', name, available, version, command, hint }]>
startChat({ provider, prompt, sessionId, cwd, filePath, mcp: { command, args, env }, model, onEvent })
  → { cancel(), done: Promise<{ sessionId, text }> }
// onEvent: {type:'text', text, delta:true} | {type:'text_reset', text, drop}
//          {type:'tool', id, name(한국어 이름), tool(원래 도구 이름), server, input} | {type:'tool_result', id, ok, summary}
//          {type:'status', message, sessionId?} | {type:'done', sessionId, text, cancelled?} | {type:'error', message, detail?}
defaultWorkspace() → 고정 작업 폴더 (%LOCALAPPDATA%\office-kit\ai). Claude는 같은 폴더에서만 대화를 이어 간다.
```
- 앱은 로그인 정보를 다루지 않는다. 사용자가 직접 로그인한 CLI(`claude`, `codex`)를 실행할 뿐이다.
- AI가 쓸 수 있는 도구는 office-kit MCP 도구로 제한한다 (Claude: 기본 도구 끔 + `--strict-mcp-config`,
  Codex: `--ignore-user-config` + 읽기 전용 샌드박스).
- `done`은 거절하지 않는다. 실행마다 `done`(취소 시 `cancelled:true`) 또는 `error`가 정확히 한 번 온다.
- `text_reset`: 스트림이 끊겨 CLI가 다시 받을 때, 이미 보낸 글 끝에서 `drop`글자를 지운다.

### src/setup/integrations.mjs
```js
mcpLaunchSpec({ execPath, appRoot }) → { command, args, env }   // Electron이면 ELECTRON_RUN_AS_NODE=1
integrationStatus(spec) → { claudeDesktop, claudeCode, codex, cli }
connectClaudeDesktop(spec) / disconnectClaudeDesktop()
connectClaudeCode(spec) / disconnectClaudeCode()
connectCodex(spec) / disconnectCodex()
installCliShim(spec) / uninstallCliShim()     // %LOCALAPPDATA%\Microsoft\WindowsApps\office-kit.cmd
```
- 상태 객체: `{ connected, matches, available, error, detail }` (항목마다 조금씩 다름).
- 설정 파일을 고치기 전에 시간 붙은 백업을 만들고, 원자적으로 쓴다. 다른 항목은 그대로 둔다.
- `office-kit.cmd`는 실행 줄이 마지막 줄이고 그 줄 안에서 뒷정리한다 (`setup cli [--remove]`가 실행 중인 자기 파일을
  바꾸거나 지워도 깨지지 않게). 종료 코드는 0/1. MSIX 앱 안에서 쓰면 결과에 `warning`이 붙는다.
- 모든 함수가 경로·실행기를 바꿔 끼울 수 있어 시험은 실제 설정을 건드리지 않는다.

### src/setup/hancom.mjs
```js
hancomSecurityStatus() → { hancomInstalled, registered, dllPath, dllExists }
setupHancomSecurity({ dllPath }) → 레지스트리 등록
downloadHancomSecurityModule(destDir) → 한컴 공식 배포본만 받음 (사용자가 버튼을 눌렀을 때만)
```
- 공식 배포처: 한컴 개발자센터 https://developer.hancom.com/hwpautomation →
  GitHub `hancom-io/devcenter-archive` 의 `hwp-automation/보안모듈(Automation).zip`.
- 설치된 Hwp.exe의 비트(PE 헤더)에 맞는 DLL을 고른다. HTTPS·GitHub 주소만, 30MB 제한.

### src/app (Electron)
- `main.mjs`: 창, 단일 인스턴스(두 번째 실행·"연결 프로그램"의 파일 경로를 기존 창에서 연다), IPC, 원본 백업,
  파일 변경 감시 → 미리보기 자동 새로고침, AI 대화 중계.
- `preload.cjs`: `window.officeKit` API만 노출 (contextIsolation, `webUtils.getPathForFile`로 끌어놓기 경로).
- `renderer/`: 왼쪽 문서 미리보기 + 오른쪽 AI 대화창. 바닐라 JS, 번들러 없음.
- `store.mjs`: 설정(JSON, 원자적 쓰기, AI마다 모델 따로 `models: {claude, codex}`), 최근 파일, 백업
  (`userData/backups`, 최근 50개).
- '원래대로'는 그 문서의 AI 대화를 멈춘 뒤 마지막 백업으로 되돌리고, 다음 메시지부터 새 대화로 시작한다.
- AI 대화 세션은 문서마다 메모리에만 둔다 (앱을 다시 켜면 새 대화).
- 확인용: `OFFICEKIT_SMOKE=<png경로>` 환경 변수로 실행하면 인수로 받은 파일을 열고 화면을 캡처한 뒤 종료 (못 열면 알림을 찍고 실패로 끝남). 최근 파일 목록·Windows 최근 항목은 건드리지 않는다.
- ESM 진입 파일에서 `await app.whenReady()`를 최상위에서 기다리면 멈춘다 → `start()`를 기다리지 않고 부른다.

## 배포
- `electron-builder` NSIS 설치 파일 (사용자별 설치, asar 끔 — PowerShell 스크립트와 MCP 서버를 그대로 실행하기 위해).
- 탐색기 오른쪽 클릭 "Office Kit로 열기" (기본 연결 프로그램은 바꾸지 않음).
- GitHub Actions: `v*` 태그를 올리면 Release에 설치 파일 업로드.
