# Office Kit 구조

한글(HWP/HWPX)·Word·PowerPoint·Excel 문서를 열고, 확인하고, AI 대화로 고치는 Windows 데스크톱 도구.

## 한눈에

```
src/core/      문서 엔진 (UI 없음, Node만 있으면 동작). 형식별: word.mjs, ppt.mjs, excel.mjs, hwp.mjs
src/mcp/       MCP 서버 — Claude/Codex가 문서 도구를 쓰게 함 (stdio). 형식별 도구는 tools/*.mjs
src/cli.mjs    명령줄 도구 office-kit
src/setup/     외부 연결: Claude Desktop/Claude Code/Codex 등록, 명령줄 등록, 한글 보안 모듈
src/app/       Electron 앱 (main.mjs, preload.cjs, renderer/, ai.mjs)
scripts/com/   office.ps1 — 설치된 프로그램을 COM으로 조작. 형식별 작업은 word-ops.ps1, ppt-ops.ps1, excel-ops.ps1
build/         아이콘, NSIS 추가 설정
```

형식별 전용 기능(살펴보기·편집)의 작업 목록·인수·한계는 [`FEATURES.md`](FEATURES.md).

형식별 기능이 흐르는 길 (앱 안 AI도 MCP 서버를 거친다):

```
MCP 도구  src/mcp/tools/<형식>.mjs ─┐
명령줄    src/cli.mjs inspect/edit ─┴→ core  src/core/<형식>.mjs ─┬→ Word·PowerPoint·Excel(app): com.mjs → office.ps1 → <형식>-ops.ps1
                                                                   ├→ Excel(file): exceljs (Node 안)
                                                                   └→ 한글: rhwp (Node 안, 한컴 COM 안 씀)
```

## 엔진 두 가지

| 엔진 | 쓰는 것 | 장점 | 언제 |
|---|---|---|---|
| app | 설치된 한컴/Word/PowerPoint/Excel (COM, `scripts/com/office.ps1`) | 원본 그대로, PDF 변환 | 쓰기 작업 기본값 |
| file | rhwp(WASM)·ZIP/XML·exceljs | 프로그램 불필요, 빠름 | 읽기 기본값, 프로그램 없을 때 |

- 한컴은 **자동화 보안 모듈**(`HKCU\Software\HNC\HwpAutomation\Modules` 의 `FilePathCheckerModule`)이 없으면
  파일마다 "접근 허용" 창이 떠서 멈춘다 → 모듈 없으면 자동 작업은 rhwp로 처리 (`comAvailableFor(ext, {interactive})`).
  형식별 한글 기능(`hwp_inspect`·`hwp_edit`)은 모듈과 상관없이 늘 rhwp다.
- 형식별 Word·PowerPoint 기능은 app 엔진만, Excel은 둘 다 (`excel_inspect` auto = 파일 엔진 우선, `excel_edit` auto = Excel 우선).
- 자동화로 여는 Office 문서는 매크로를 막는다 (`AutomationSecurity = 3`, 사람이 보는 창은 2).
- COM 호출의 줄 세우기·시간 초과는 아래 '시간 초과와 잠금'.

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
`pageImage(path, page = 1, { dpi }) → { png: Buffer, page, pages, engine, width, height }` — AI가 결과를 눈으로 확인하는 용도.
- hwp/hwpx/hwt: rhwp SVG → `@resvg/resvg-js` (시스템 글꼴).
- PowerPoint(pptx·ppt·odp): PDF를 거치지 않고 `office.ps1`의 `app:'ppt', op:'slideImage'`(`Slide.Export`, 읽기 전용으로 엶) →
  **쪽 = 슬라이드 번호(숨긴 슬라이드 포함)**. 암호 문서는 열기 전에 거르고(`pptProtection`), pptx·odp는 꾸러미에서 슬라이드 수를 먼저 세어 범위를 확인한다.
  그림은 PowerPoint 기본 크기까지 (작은 dpi는 resvg로 줄임). 같은 파일·슬라이드 그림은 24개까지 기억.
- Word·Excel: 프로그램으로 PDF 변환(8개까지 기억) → `office.ps1`의 `app:'pdf', op:'pageImage'` (Windows.Data.Pdf)로 PNG.
- `.pdf` 입력도 바로 처리.

### src/mcp/server.mjs
stdio MCP 서버 `office-kit`. stdout은 프로토콜 전용, 로그는 stderr.
공통 도구 10개: `office_status`, `office_info`, `office_read`, `office_fields`, `office_fill`, `office_replace`,
`office_set_cells`, `office_convert`, `office_create`, `office_page_image`.
형식별 도구 8개(`src/mcp/tools/*.mjs`): `hwp_inspect`·`hwp_edit`, `word_inspect`·`word_edit`, `ppt_inspect`·`ppt_edit`, `excel_inspect`·`excel_edit`.
- 쓰기 도구는 기본이 제자리 저장. 다른 이름(`output`)이 이미 있으면 `overwrite=true`일 때만 덮어쓴다.
- 저장 가능한 확장자만 허용하고, `OFFICEKIT_WRITE_ROOTS`가 있으면 그 폴더 밖에는 쓰지 않는다.
- `office_read`는 `offset`/`max_chars`로 나눠 읽는다 (잘라 버리지 않고 다음 위치를 알려 준다).
- SDK가 만드는 입력 검증 오류도 `koreanError`로 한국어로 바꿔 돌려준다 (흔한 zod 영어 문구만).

### src/mcp/tools/*.mjs (형식별 도구 모듈)
```js
export function register(server, ctx)   // server.mjs의 TOOL_MODULES = [hwp, word, ppt, excel] 순서로 부른다
// ctx: { z, scalar, tool, text, json, log, abs, mustExist, outPath, koreanError,
//        pathSchema, outputSchema, overwriteSchema, engineSchema, EDIT_EXTS, CONVERT_EXTS, HWP_EXTS }
```

| ctx | 하는 일 |
|---|---|
| `tool(fn)` | 도구 함수를 감싸 예외를 한국어 `isError` 결과(`오류: …`)로 |
| `text(t)` / `json(o)` | MCP 글 결과 / 보기 좋은 JSON 문자열 |
| `abs(p)` | 서버 작업 폴더 기준 절대 경로 (따옴표 벗김). 그림·테마 경로도 이걸로 바꿔 core에 넘긴다 |
| `mustExist(p)` | 없으면 "파일이 없습니다: …" |
| `outPath(output, { source, allow, overwrite })` | 저장 경로 확인: 확장자(`allow`), `OFFICEKIT_WRITE_ROOTS`, 다른 파일 덮어쓰기 — 모두 통과한 뒤에만 폴더를 만든다 |
| `koreanError(e)` | Node(EBUSY·ENOENT…)·zod 오류를 한국어로 |
| `pathSchema` · `outputSchema` · `overwriteSchema` · `engineSchema` | 공통 입력 스키마 (한국어 설명 포함) |
| `EDIT_EXTS` · `CONVERT_EXTS` · `HWP_EXTS` | 저장할 수 있는 확장자 |

- 모듈은 입력 스키마(zod, `z.locales.ko()` 한국어 로캘)·한국어 설명·결과 글 모양만 갖고, 실제 일은 core에 맡긴다.
- 편집 도구 계약: `{ path, operations, output?, overwrite? }` (Excel은 `engine?`) → 작업을 모두 적용한 뒤 한 번 저장,
  하나라도 실패하면 저장하지 않음. 결과는 작업마다 한 줄 + `참고:` 줄 + `저장: 경로`.
- 작업 스키마가 사는 곳 (명령줄도 같은 확인을 거치게):

| 형식 | MCP 스키마 | core 확인 | 명령줄 |
|---|---|---|---|
| Word | tools/word.mjs (모르는 인수는 버림) | `normalizeWordOps` | core만 |
| PowerPoint | core `pptOpSchema` (엄격) | `validatePptOps` | core (같은 스키마) |
| Excel | tools/excel.mjs (엄격) | `normalizeOperations` + `checkKeys` (엄격) | core만 |
| 한글 | tools/hwp.mjs `hwpOp` (엄격) | `Hwp` 메서드가 번호·크기 범위 | `hwpOp`을 import해서 확인 |

- 한글 편집은 tools/hwp.mjs가 직접 돌린다: `Hwp.open` → `applyHwpOp`(작업 하나 → 결과 한 줄) 반복 → `h.save`.
  저장하지 못하면 `outPath`가 만든 빈 폴더를 지운다. `inspectText`(살펴보기 글)도 이 모듈에 있다.
- 새 형식 도구를 더하려면: `tools/<형식>.mjs`에 `register`를 만들고 `TOOL_MODULES`에 넣는다 (`test/mcp.test.mjs`의 도구 목록도).

### 형식별 core (src/core/word.mjs · ppt.mjs · excel.mjs · hwp.mjs)

| 모듈 | 살펴보기 | 편집 → 결과 |
|---|---|---|
| word.mjs | `wordInspect(path, {offset, limit})` → 원자료, `formatWordInspect(info)` → 글 | `wordEdit(path, ops, {output})` → `{ output, engine:'word', results, lines }` |
| ppt.mjs | `pptInspect(path, {slide, offset, limit})`, `formatPptInspect(info)` | `pptEdit(path, ops, {output})` → `{ engine:'powerpoint', results(줄), notes, output, raw }` |
| excel.mjs | `excelInspect(path, {engine})`, `inspectText(info)` | `excelEdit(path, ops, {output, engine})` → `{ engine:'excel'\|'file', results, notes, output }` |
| hwp.mjs | `Hwp#inspect({offset, limit})` (글은 tools/hwp.mjs `inspectText`) | `Hwp` 메서드 (`insertParagraph`, `formatText`, `addTableRow` …) |

- core가 하는 일: 경로·형식 확인 → 프로그램 확인(`probe`/`comAvailableFor`, 없으면 대안을 붙인 한국어 오류) → 작업 정리·검증
  ("N번째 작업(op) 입력 오류: …") → COM 요청 → 영어 오류 코드를 한국어로.
- COM 오류 코드 → 한국어:
  - Word: `op <n> (<name>): <영어 코드>` (예: `paragraph out of range: 9 of 8`) → `word.mjs` `ERRORS`.
  - PowerPoint: `opfail|<n>|<op>|E|<code>|인수…` (예: `E|slide_range|99|3` → "슬라이드 99번이 없습니다 (지금 3장)") → `ppt.mjs` `ERRORS`.
  - Excel: `opfail|<n>|<op>|E_<CODE>|인수…` (예: `E_TABLE_HEADER`) → `excel.mjs` `COM_CODES`.
  - 작업 밖(열기·저장): `open failed: …`, `save failed: …`, `read only`, `locked for editing` → `com.mjs`·각 core가 한국어로.
- 결과 원자료는 작업마다 해시 하나. Word는 `paragraphs`(작업 뒤 문단 수)와 `ms`(Word 안에서 그 작업에 걸린 시간 — 잠금 대기 제외, 시험·진단용)를 붙인다.

### scripts/com/office.ps1 + *-ops.ps1 (COM)
- `office.ps1 <요청.json> <결과.json>`: Node(`com.mjs`)가 임시 폴더에 요청 JSON을 쓰고 PowerShell을 띄운다.
- 시작하자마자 `word-ops.ps1`·`ppt-ops.ps1`·`excel-ops.ps1`을 (있으면) **점 소싱**해 같은 범위에 불러온다.
  - 각 파일은 `Invoke-<App>Inspect` / `Invoke-<App>Edit`을 정의한다 (`Invoke-WordEdit`, `Invoke-PptEdit`, `Invoke-ExcelEdit` …).
  - 같은 범위라 이름이 부딪치지 않게 도우미에 형식 접두를 붙인다: Word `Wd…`, PowerPoint `…-Ppt…`, Excel `Xl…`.
- `Invoke-Word` / `Invoke-Ppt` / `Invoke-Excel`이 문서를 열고 `switch ($req.op)`로 나눈다:
  `'inspect'` → `$res.inspect`, `'edit'` → `$res.results`(작업별 원자료). 저장할 것(`save`)이 없으면 읽기 전용으로 연다.
- 요청: `{ app, op:'inspect'|'edit', path, operations, save, saveFormat, pidFile, …형식별 }`
  (Word `offset`·`limit`, PowerPoint `slide`·`offset`·`limit`·`maxShapes`, Excel `csvSheets`). 결과: `{ ok:true, inspect | results, saved }` 또는 `{ ok:false, error }`.
- 파일은 **ASCII만** (PowerShell 5.1은 BOM 없는 스크립트를 ANSI로 읽음). 한국어는 JSON 요청으로 받고, 오류는 짧은 영어 코드로 던져
  JS(core)가 한국어로 바꾼다.
- COM 객체가 살아 있는 채로 Quit하면 숨은 Office 프로세스가 남을 수 있다 → 일은 안쪽 함수에서 하고 돌아오기 전에 GC (`XlCollect`, excel-ops.ps1).
  PowerShell COM 어댑터는 처음 넣은 값의 변환 형식을 기억하므로 Excel 값은 리플렉션(`InvokeMember`)으로 넣는다.
- PowerPoint `add_chart`는 창이 있어야 해서 최소화된 창을 열고, 그 호출이 띄운 차트 데이터용 Excel만 정리한다.

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
- 지침(`systemPrompt`)은 열린 문서 형식의 `…_inspect` → `…_edit`를 먼저 쓰게 하고(`FORMAT_TOOLS`), 글자만 바꿀 때는 `office_replace`/`office_fill`,
  고친 뒤에는 `office_page_image`로 확인하게 한다. 화면(renderer)은 편집 도구(`EDIT_TOOLS`: office_fill·replace·set_cells·create·convert,
  hwp/word/ppt/excel_edit) 결과가 오면 미리보기를 새로 고치고, 첫 화면 추천 문장을 문서 종류(한글·Word·PowerPoint·Excel)에 맞춰 보여 준다.
- 앱은 로그인 정보를 다루지 않는다. 사용자가 직접 로그인한 CLI(`claude`, `codex`)를 실행할 뿐이다.
- AI가 쓸 수 있는 도구는 office-kit MCP 도구로 제한한다 (Claude: 기본 도구 끔 + `--strict-mcp-config`,
  Codex: `--ignore-user-config` + 읽기 전용 샌드박스).
- `done`은 거절하지 않는다. 실행마다 `done`(취소 시 `cancelled:true`) 또는 `error`가 정확히 한 번 온다.
- `text_reset`: 스트림이 끊겨 CLI가 다시 받을 때, 이미 보낸 글 끝에서 `drop`글자를 지운다.

### src/cli.mjs (형식별 부분)
- `inspect`·`edit`는 확장자로 형식을 고르고(`FORMAT_OF`) core 함수를 그대로 부른다. 한글은 `tools/hwp.mjs`의 `hwpOp`·`applyHwpOp`·`inspectText`를 쓴다.
- `--engine` 확인: 한글 `app`, Word·PowerPoint `file`은 한국어 오류. core 글의 MCP 표기는 명령줄 표기로 바꿔 보여 준다:
  오류·참고(`engine:'app'` → `--engine app`, `output` → `-o`, `cliText`), 살펴보기 이어 보기 안내(`offset=N` → `--offset N`, `slide=N` → `--slide N`, `INSPECT_HINTS`).
- PowerPoint `--offset/--limit`은 슬라이드 단위로 `pptInspect`에 넘긴다 (`--limit`이 없으면 끝까지). 편집의 `notes`는 `참고:` 줄로 찍는다.
- `help edit`은 core가 내보내는 작업 이름 목록(`WORD_OPS`, `OP_NAMES`, `EXCEL_OPS`, `hwpOp`)으로 만든다 → 작업이 늘면 저절로 따라간다.

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

## 편집 원자성

규칙 하나: **모든 작업을 적용한 뒤 한 번만 저장, 하나라도 실패하면 아무것도 쓰지 않는다.**

| 엔진 | 어떻게 지키나 |
|---|---|
| Word·PowerPoint·Excel (COM) | `Invoke-<App>Edit`가 작업을 차례로 돌린다. 예외가 나면 `office.ps1`은 저장 단계에 가지 않고, `finally`에서 문서를 **저장 없이** 닫는다. 저장은 끝에 한 번: 원본이면 `Save()`, 아니면 `SaveAs(output, 형식)` |
| Excel 파일 엔진 (exceljs) | 메모리에서 모든 작업 → 만든 파일을 다시 읽어 확인 → 같은 폴더 임시 파일에 쓰고 이름 바꾸기(`atomicWrite`). csv도 같음 |
| 한글 (rhwp) | 메모리 문서에 모든 작업 → `Hwp.save`: 임시 파일에 다 쓴 뒤 이름 바꾸기 (백신·미리보기가 잠깐 잡으면 몇 번 다시) |

저장 전에 거르는 것 (그대로 두면 보이지 않는 창에서 멈추거나 조용히 저장이 안 됨):
- Word: 원본 자리 저장인데 파일이 읽기 전용이면 Word를 띄우기 전에 거절, 다른 프로그램이 열고 있어 읽기 전용으로 열리면 작업 전에 거절(`read only`).
- Excel: 원본 자리 저장인데 읽기 전용으로 열리면 거절(`locked for editing`). 저장 뒤 파일 수정 시각이 그대로면 실패로 본다.
  Excel로 csv를 만들 때는 임시 폴더에 저장한 뒤 인코딩·줄바꿈을 맞춰 `atomicWrite`.
- PowerPoint: 열기 암호·.ppt 암호 문서는 열기 전에 거절. 쓰기 암호 pptx는 쓰기 암호를 뺀 임시 사본으로 편집하고, 결과 pptx에 같은 암호를 다시 건다.
- 시험: 네 형식 모두 "실패하면 원본이 바이트 그대로"를 확인한다 (`test/<형식>.test.mjs`, COM은 `OFFICEKIT_COM_TESTS=1`).

## 시간 초과와 잠금

줄 세우기는 두 겹이다.
1. **프로세스 안**: `com.mjs`의 `comCall`은 큐 — 앞 호출이 끝나야 다음 PowerShell을 띄운다.
2. **프로세스끼리**: `office.ps1`이 이름 있는 뮤텍스 `Local\OfficeKitCom`을 잡은 뒤에만 Office를 띄운다 (앱·MCP 서버·명령줄이 동시에 돌아도 하나씩).
   `probe`(설치 확인)와 `pdf`(PDF → PNG)는 Office를 띄우지 않아 잠그지 않는다.

호출마다 한도(`timeoutMs`)가 있다 — 형식별 값은 [`FEATURES.md`](FEATURES.md) '엔진과 시간 제한'.
한도를 넘으면:
- PowerShell을 끝내고, `office.ps1`이 Office를 띄울 때 적어 둔 **PID 파일**로 그 Office를 끝낸다 (DCOM으로 뜬 Office는 PowerShell의 자식이 아님).
- Word 편집은 그 Word가 남긴 `~$` 소유자 파일을 지운다 (이번 호출이 시작된 뒤 생긴 작은 파일만 — 사용자가 연 Word의 것은 안 지움).
- PowerPoint `add_chart`의 차트 데이터 Excel은 PID를 몰라 남을 수 있다.

잠금 대기와 한도:
- 한도는 **프로그램이 문서를 다루는 시간**만 재야 한다. 다른 프로세스가 잠금을 오래 쥐고 있을 때 기다린 시간까지 세면,
  Office가 멀쩡히 일하고 있어도(또는 아직 시작도 못 했는데) "응답하지 않아 작업을 멈췄습니다"로 끝난다.
- 그래서 PID 파일과 같은 방식으로 **시작 표시(started marker)** 를 둔다:
  1. Node가 요청 JSON에 시작 표시 파일 경로를 함께 보낸다.
  2. `office.ps1`은 잠금을 잡은 **직후**, Office를 띄우기 전에 그 파일을 만든다.
  3. Node는 시작 표시가 생긴 때부터 작업 한도를 잰다. 그 전(잠금을 기다리는 시간)은 따로, 더 긴 한도와 한국어 안내로 다룬다.
  4. 시작 표시 없이 한도를 넘었다면 아직 Office를 띄우지 않았으므로 PID 정리도 하지 않는다.
- 구현: 요청의 `startedFile`(office.ps1이 잠금을 잡은 직후 씀)을 `com.mjs`가 200ms마다 확인한다.
  잠금 대기 한도는 `comCall`의 `lockTimeoutMs` (기본 `LOCK_WAIT_MS` = 15분) — 넘으면 "다른 Office 작업이 끝나기를 기다리다 멈췄습니다".
  `probe`·`pdf`는 잠그지 않으므로 바로 한도를 잰다.
- 쓰기 암호가 걸린 문서: 그 자리 저장이 아니면 읽기 전용으로 열고, 그 자리 저장은 OOXML 안의 쓰기 보호 표시를 미리 보고 바로 거절한다.
- 앱 이름 뒤 조사는 `com.mjs`의 `josa()`로 붙인다 (Word가, Excel이, 한글이).

## 배포
- `electron-builder` NSIS 설치 파일 (사용자별 설치, asar 끔 — PowerShell 스크립트와 MCP 서버를 그대로 실행하기 위해).
  `docs/`도 함께 들어간다 (`office-kit help edit`이 `docs/FEATURES.md` 경로를 알려 줌).
- 탐색기 오른쪽 클릭 "Office Kit로 열기" (기본 연결 프로그램은 바꾸지 않음).
- GitHub Actions: `v*` 태그를 올리면 Release에 설치 파일 업로드.
