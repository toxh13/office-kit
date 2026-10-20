# Office Kit 구조

한글(HWP/HWPX)·Word·PowerPoint·Excel 문서를 열고, 확인하고, AI 대화로 고치는 Windows 데스크톱 도구.

## 한눈에

```
src/core/      문서 엔진 (UI 없음, Node만 있으면 동작). 형식별: word.mjs, ppt.mjs, excel.mjs, hwp.mjs
src/mcp/       MCP 서버 — Claude/Codex가 문서 도구를 쓰게 함 (stdio). 형식별 도구는 tools/*.mjs
src/cli.mjs    명령줄 도구 office-kit
src/setup/     외부 연결: Claude Desktop/Claude Code/Codex 등록, 명령줄 등록, 한글 보안 모듈
src/app/       Electron 앱 (main.mjs, preload.cjs, renderer/, ai.mjs)
               편집 창 (종류 kind = 'hwp' | 'office'): editor-window.mjs(창·저장·감시), editor-server.mjs(127.0.0.1·.2·.3 정적 서버)
               한글: editor/(host 페이지·preload) / 문서·표·슬라이드: office-editor/(host 페이지·preload, fonts/ 글꼴 목록 생성기 — AGPL)
src/core/ooxml-fixups.mjs  Office 편집 창이 저장한 docx·xlsx·pptx를 Office가 경고 없이 열게 고치기
src/core/hwppdf.mjs  한컴오피스 없이 한글 → PDF (vendor/rhwp-cli/rhwp.exe 실행)
src/core/diff.mjs    두 버전의 바뀐 글 (문단·글자 단위 Myers, 엑셀은 칸 단위) — '바뀐 내용 보기'
src/core/pii.mjs     개인정보 찾기(글자 규칙)·가린 사본 만들기(replaceText) + 다시 읽어 남은 것 확인
src/core/batch.mjs   한글 → HWPX 일괄 변환(convertDocument) + 다시 열어 쪽·표·글 비교, CSV 보고서
src/setup/fileassoc.mjs  기본 프로그램 후보 등록 상태 읽기·고쳐 쓰기 (HKCU, build/installer.nsh와 같은 키)
scripts/com/   office.ps1 — 설치된 프로그램을 COM으로 조작. 형식별 작업은 word-ops.ps1, ppt-ops.ps1, excel-ops.ps1
scripts/build-editor.mjs  한글 편집기 묶음 만들기 (rhwp-studio v0.8.7 → vendor/)
scripts/build-office-editor.mjs  Office 편집기 묶음 만들기 (filex-office-editor v0.3.1 ui.zip → vendor/office-editor), 고정 값·규칙은 scripts/office-editor/
scripts/build-rhwp-cli.mjs  한글→PDF 변환 프로그램 받기 (rhwp v0.8.7 Release zip, 고정 SHA-256 → vendor/rhwp-cli)
scripts/make-office-source.mjs   Office 편집기 AGPL 소스 꾸러미 (office-editor-source-<버전>.zip)
patches/office-editor/filex/     filex 앱 패치 (AGPL-3.0-or-later)
LICENSES/                        AGPL 3판 전문·ONLYOFFICE 추가 조항·편집기 기계어 제3자 고지 (third-party.mjs가 만듦)
vendor/        만들어지는 폴더 (git 밖): rhwp-studio/ 편집기, rhwp-editor-sdk/ 연결 SDK, office-editor/ 문서·표·슬라이드 편집기 (AGPL),
               rhwp-cli/ 한글→PDF 변환 프로그램 (rhwp.exe, MIT)
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
- 한글 → PDF는 한컴(+모듈)이 있으면 app, 없으면 rhwp 명령줄 프로그램(아래 'src/core/hwppdf.mjs'). 한컴 COM은 모듈이 없으면 부르지 않는다.
- 형식별 Word·PowerPoint 기능은 app 엔진만, Excel은 둘 다 (`excel_inspect` auto = 파일 엔진 우선, `excel_edit` auto = Excel 우선).
- 자동화로 여는 Office 문서는 매크로를 막는다 (`AutomationSecurity = 3`, 사람이 보는 창은 2).
- COM 호출의 줄 세우기·시간 초과는 아래 '시간 초과와 잠금'.

## 환경 변수

| 이름 | 쓰는 곳 | 뜻 |
|---|---|---|
| `OFFICEKIT_WRITE_ROOTS` | MCP 서버 | 저장을 허락할 폴더들 (`;` 구분). 앱 안 AI 대화는 열린 문서의 폴더만 넘긴다 |
| `OFFICEKIT_POWERSHELL` | core/com | 시험용 PowerShell 경로 바꿔 끼우기 |
| `OFFICEKIT_RHWP_CLI` | core/hwppdf | 한글→PDF에 쓸 rhwp 실행 파일 (기본 `vendor/rhwp-cli/rhwp.exe`). 시험·다른 OS용 |
| `OFFICEKIT_RHWP_TIMEOUT_MS` | core/hwppdf | 한글→PDF 시간 제한 (기본 60초 + 쪽당 5초, 최대 20분) |
| `OFFICEKIT_RHWP_CLI_CACHE` | build-rhwp-cli | 받은 rhwp zip·THIRD_PARTY_LICENSES.md 캐시 폴더 (기본 `tmp/rhwp-cli-build/download`) |
| `OFFICEKIT_SMOKE`, `_CHAT`, `_SETTINGS`, `_NEWDOC`, `_EDITOR` | 앱 | 화면 캡처 확인 모드 (`_EDITOR=1`: 편집 창도 열어 `<png>-editor.png`로 찍음, `_PDF=<경로>`: 'PDF로 저장' 단추를 눌러 저장 대화상자 대신 그 경로에 만들고 알림까지 확인 → `[smoke]` 줄의 `pdf`) |
| `OFFICEKIT_EDITOR_VENDOR` | 앱(개발용) | 편집기 묶음 폴더 바꾸기 (`<폴더>/rhwp-studio`, `rhwp-editor-sdk`). 없는 폴더 = 묶음 없음 시험 |
| `OFFICEKIT_EDITOR_TEST=1` | 앱(개발용) | 편집 창 host 페이지에 시험 고리 `window.__editorTest` (`?test=1`) |
| `OFFICEKIT_EDITOR_CACHE` | build-editor | rhwp 소스·의존성 캐시 폴더 (기본 `tmp/editor-build`) |
| `OFFICEKIT_OFFICE_EDITOR_CACHE` | build-office-editor, make-office-source | Office 편집기 받은 파일·filex 소스 캐시 폴더 (기본 `tmp/office-editor-build`) |
| `OFFICEKIT_FILEX_SRC` | make-office-source | filex 고정 커밋이 있는 git 폴더 (`--filex-src`와 같음) |
| `OFFICEKIT_FILEX_REPO` | make-office-source | filex 소스를 받을 곳 — 거울 저장소를 쓸 때 (`--filex-repo`와 같음) |
| `SOURCE_DATE_EPOCH` | make-office-source | 소스 꾸러미 안 파일 시각 (없으면 HEAD 커밋 시각 → 같은 입력이면 같은 바이트) |
| `OFFICEKIT_EDITOR_SMOKE=1` | 테스트 | `test/editor-integration.test.mjs`의 실제 Electron 캡처 시험 켜기 (창이 잠깐 뜸) |
| `OFFICEKIT_USER_DATA` | 앱 | 설정·백업 폴더 바꾸기 (시험용) |
| `OFFICEKIT_AI_MODULE` | 앱(개발용) | 가짜 AI 모듈로 바꿔 끼우기 |
| `OFFICEKIT_COM_TESTS=1` | 테스트 | Word/PowerPoint/Excel COM 시험 켜기 (한컴 제외) |
| `OFFICEKIT_DEBUG` | CLI/MCP | 오류 스택 출력 |

## 모듈 계약

### src/core/index.mjs (있음)
`readDocument(path, {engine})`, `documentInfo(path)`, `replaceText(path, pairs, {output, engine})`,
`listFields(path)`, `fillFields(path, values, {output, engine})`, `setCells(path, cells, {sheet, output, engine})`,
`convertDocument(path, output)`, `renderPreview(path, outDir, {mode})`, `openInApp(path)`,
`createDocument(output, markdown)`, `probe()`, `Hwp`, `OPENABLE_EXTS`, `extOf`, `kindOf`, `hwpPdfStatus()`, `HWP_PDF_NOTE`.
- `convertDocument` → `{engine, output, pages?, note?}`. 한글 → pdf: 한컴(+보안 모듈)이 있으면 `engine:'hwp'`, 없으면 `hwpToPdf`로 `engine:'rhwp'` +
  `pages`·`note`(`HWP_PDF_NOTE`: 한컴 없이 내장 엔진으로 만듦 — 화면·MCP·명령줄이 이 안내를 그대로 보여 준다).
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
  변환은 `index.mjs`의 `previewPdf`: Excel은 임시 사본을 시트마다 한 쪽 너비로(`fitWidth` → `Set-XlFitWidth`: Zoom 끔, FitToPagesWide 1, FitToPagesTall 끔) 내보낸다 — 넓은 표·차트가 쪽마다 잘리지 않게. 원본은 건드리지 않는다. 앱 미리보기도 같은 함수.
- 한글 미리보기(`Hwp.openForView`): rhwp가 표의 바깥 아래 여백을 무시하므로 화면용 사본에서만 표 다음 문단의 문단 위 간격에 그 여백을 더한다 (`_gapAfterTables`).
- `.pdf` 입력도 바로 처리.

### src/core/hwppdf.mjs
`hwpToPdf(path, output, {timeoutMs}) → {engine:'rhwp', output, pages, bytes, ms}`, `hwpPdfStatus() → {ok, path, reason?}`.
- 한컴오피스 없이 한글(hwp·hwpx·hwt) → PDF. Node만으로 돈다 (Electron·Chromium 불필요) → 앱·MCP(`ELECTRON_RUN_AS_NODE`)·명령줄이 같은 길.
- 순서: `Hwp.openForView`(미리보기와 같은 화면용 사본: 표 다음 문단 간격, 줄 정보 없는 hwpx는 hwp로 거침) → `toBytes('hwp')` → 임시 폴더 `doc.hwp`
  → `rhwp.exe export-pdf doc.hwp -o doc.pdf --backend svg --json --fallback-serif Batang --fallback-sans "Malgun Gothic" --fallback-mono GulimChe`
  (`execFile`, 셸 없이, 작업 폴더 = 임시 폴더, 시간 제한이 지나면 강제 종료) → `%PDF-` 확인 → 같은 폴더 임시 이름에 복사 후 이름 바꿔 `output`에 → 임시 폴더 지움.
- 원본은 읽기만 한다. `--json` 결과(stdout 마지막 줄)의 `pageCount`를 쪽 수로 돌려준다.
- 엔진 고르기(2026-10 시험, `tmp/hwp-pdf`): ① rhwp 명령줄 svg 방식 — 미리보기와 같은 모양, 글자 100% 선택·검색(ToUnicode), 부분 글꼴, Node만으로 됨, 느림(약 0.5초/쪽).
  ② rhwp `--backend direct` — 빠르지만(77쪽 4초) 굵기가 틀리고 그러데이션이 든 문서는 실패 → 쓰지 않음.
  ③ rhwp SVG → 숨긴 BrowserWindow → `printToPDF` — 빠르지만(77쪽 10초) 장평(가로 비율)이 든 글자를 Type3 그림 글꼴로 넣어 실제 공문서에서 한글 글자의 절반 가까이가
  검색되지 않았고, MCP·명령줄에서는 Electron을 따로 띄워야 함 → 쓰지 않음.

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
  예외: 사람에게 넘기는 `op:'open', visible:true`('프로그램에서 열기')는 읽기·쓰기로 연다. 쓰기 암호가 있으면(OOXML 미리 확인, Word는 열기 실패로도 앎) 읽기 전용으로 열고
  `readOnly:true, readOnlyReason:'write password'`, 다른 프로그램이 쥐고 있어 읽기 전용이 되면 `'in use or read-only file'` → `openInApp`이 `readOnlyReason: 'writePassword'|'inUse'`로, 화면은 알림.
- Excel `freeze`: 숨은 Excel은 SplitRow·FreezePanes를 못 건다 → .xlsx·.xlsm으로 저장할 때는 `freezeInFile`: Excel은 창만 풀고 사본에 저장, Node(`writeFrozenPanes`)가 시트 보기에 `<pane state="frozen">`을 써서 `atomicWrite`
  (그 자리 저장이면 쓰기 암호를 Node가 먼저 확인). 다른 형식은 COM으로 해 보고 안 되면 `E_FREEZE_HIDDEN`.
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

### src/setup/fileassoc.mjs
```js
registrationEntries(exe) → [{ key, name, value }]   // HKCU 아래: ProgID OfficeKit.<확장자>, .<확장자>\OpenWithProgids,
                                                     // Software\OfficeKit\Capabilities, RegisteredApplications, Applications\Office Kit.exe
assocStatus({ exe }) → { registered, matches, types: [{ ext, desc, kind, progId, isDefault }] }
registerFileTypes({ exe })                           // reg.exe add (설치된 앱에서만 — main이 확인)
DEFAULT_APPS_URI = 'ms-settings:defaultapps?registeredAppUser=Office%20Kit'
```
- 기본 프로그램(UserChoice)은 쓰지 않는다 — Windows가 해시로 막고, 고르는 것은 사용자 몫이다. `.<확장자>` 키의 기본값도 건드리지 않는다.
- 읽기는 PowerShell + .NET 레지스트리 API(한글 경로가 깨지지 않게, `Get-ItemProperty`는 없는 키에서 느려서 안 씀). `UserChoiceLatest`(24H2)가 `UserChoice`보다 앞선다.
- 설치 프로그램(`build/installer.nsh`)이 같은 키를 쓰고, 진짜 제거일 때만 지운다 (업데이트 중에는 그대로 — 사용자가 고른 기본 프로그램이 끊기지 않게).

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
  파일 변경 감시 → 미리보기 자동 새로고침, AI 대화 중계, 새 문서(`doc:newInfo`·`doc:newCheck`·`doc:new`), 한글 편집 창 잇기(아래).
- `preload.cjs`: `window.officeKit` API만 노출 (contextIsolation, `webUtils.getPathForFile`로 끌어놓기 경로).
- `renderer/`: 왼쪽 문서 미리보기 + 오른쪽 AI 대화창. 바닐라 JS, 번들러 없음.
- `store.mjs`: 설정(JSON, 원자적 쓰기, AI마다 모델 따로 `models: {claude, codex}`), 최근 파일, 백업
  = 문서 버전 (`userData/backups`, 항목마다 까닭 `reason`: `open`·`editor`·`ai`·`restore`, 문서마다 최근 30개(main의 `maxPerSource`), 모두 합쳐 200개·2GB),
  미리보기 PDF 기억 `createPreviewCache` (`userData/preview-cache`, 열쇠 = 경로·크기·수정 시각·방식의 해시, 40개·512MB).
- 버전을 남기는 때: 문서를 열 때, 편집 창이 덮어쓰기 전, AI 대화(`ai:send`)마다 시작 전, 되돌리기 전. 같은 내용이면 있던 것을 다시 쓴다 (처음 까닭 그대로).
- '원래대로'(`doc:restore`)는 그 문서의 AI 대화를 멈춘 뒤 버전으로 되돌리고, 다음 메시지부터 새 대화로 시작한다.
  인수 없으면 마지막 변경 전(지금 내용과 다른 가장 최근 버전), `{ file }`이면 그 버전 (`doc:versions`의 목록에 있는 것만).
  화면은 원래대로 옆 ▾에서 버전 목록(시각·까닭·'이 버전으로 되돌리기')을 띄운다.
- 미리보기(`doc:preview`): 프로그램으로 만든 PDF는 `preview-cache`에서 찾고, 없으면 만든 뒤 (만드는 동안 파일이 그대로였으면) 기억한다.
- PDF로 저장(`doc:savePdf`)은 두 번: `{step:'pick'}` 저장 대화상자 → `{step:'make'}` 고른 곳에 만들기 (화면은 고른 뒤에만 'PDF 만드는 중').
- AI 대화 세션은 문서마다 메모리에만 둔다 (앱을 다시 켜면 새 대화).
- 바뀐 내용(`doc:diff { file }`): 그 버전 → 지금 파일(임시 사본)을 `core/diff.mjs`로 비교. AI 대화가 끝날 때(`ai:event` `end`) 파일이 대화 전 버전과 다르면
  `changed: true, before: <버전 파일>`을 붙여 보내고, 화면은 대화창에 '바뀐 내용 보기 · 이 대화 전으로 되돌리기'를 띄운다.
- 문서를 열면 편집 창도(`settings.openAction`: `'edit'` 기본 | `'preview'`): 사용자가 연 것(탐색기·끌어놓기·최근 목록·열기 대화상자)이면
  `doc:opened`에 `autoEdit: true`를 붙이고, 화면이 '직접 편집'과 같은 길(`doc:edit`)로 편집 창을 띄운다. 새 문서는 자기 '바로 편집 창 열기'를 따른다.
- 기본 프로그램(`assoc:status`·`assoc:repair`·`assoc:openSettings`): 설치된 앱에서만 등록을 고쳐 쓰고 Windows 기본 앱 설정을 연다.
- 도구: 개인정보(`pii:scan`·`pii:mask` — 원본 옆 `<이름>_가림.<확장자>`), 한글 → HWPX 일괄 변환(`batch:pick`·`batch:scan`·`batch:start`(진행은 `batch:progress` 이벤트)·`batch:cancel`·`batch:saveReport`).
- 확인용: `OFFICEKIT_SMOKE=<png경로>` 환경 변수로 실행하면 인수로 받은 파일을 열고 화면을 캡처한 뒤 종료 (못 열면 알림을 찍고 실패로 끝남). 최근 파일 목록·Windows 최근 항목은 건드리지 않는다.
  `OFFICEKIT_SMOKE_AUTOEDIT=1`: '열면 바로 편집 창'이 실제로 창을 띄웠는지 `[smoke]` 줄의 `autoEditor`로. `OFFICEKIT_SMOKE_DIFF=<파일>`: 연 뒤 그 파일로 덮어쓰고 '바뀐 내용' 창을 찍는다.
  `OFFICEKIT_SMOKE_TOOL=pii|pii-mask|batch` (+ `OFFICEKIT_SMOKE_BATCH=<폴더>`): 도구 창을 찍는다. `OFFICEKIT_SMOKE_SETTINGS=open`: 설정의 파일 열기 쪽.
  `OFFICEKIT_SMOKE_EDITOR=1`이면 그 문서를 편집 창으로도 열어, 다 뜨면(또는 실패 안내가 뜨면) `<png>-editor.png`로 찍고 주 창을 찍는다 (편집기가 못 뜨면 종료 코드 1).
- ESM 진입 파일에서 `await app.whenReady()`를 최상위에서 기다리면 멈춘다 → `start()`를 기다리지 않고 부른다.
- 정리(감시 끄기·AI 대화 멈춤·미리보기 PDF 지우기)는 `will-quit`에서 한다. '끝내기'가 편집 창의 저장 확인에서 **취소**될 수 있어서 `before-quit`에서는 하지 않는다.

### 한글 편집 창 (src/app/editor-window.mjs · editor-server.mjs · editor/)

앱 안에서 한글 문서(hwp·hwpx·hwt)를 직접 고치는 창. 편집기는 rhwp-studio(웹 앱)를 묶어 넣은 것(`vendor/`)이고, 인터넷 없이 동작한다.

```
주 창(renderer) ─ doc:edit ─→ main.mjs editDocument() ─→ createEditorManager().openEditor(path)
                                                          │  문서마다 BrowserWindow 하나 (partition 'officekit-editor', 디스크에 안 남음)
                                                          ▼
  http://127.0.0.1:<A>/host/host.html  (src/app/editor/host.*, preload editor-preload.cjs → window.officeKitEditor)
        └ iframe  http://127.0.0.2:<B>/studio/index.html?chrome=embed  (vendor/rhwp-studio, 따로 렌더러 프로세스)
          ↑ SDK(vendor/rhwp-editor-sdk) createEditor → MessageChannel로 loadFile · exportHwp/exportHwpx · notifySaved · commands.context · getDocumentState
```

- **두 사이트**: host(`127.0.0.1:A`, `/host/*`·`/sdk/*`)와 studio(`127.0.0.2:B`, `/studio/*`)를 다른 사이트로 띄운다 → studio 스크립트가 host의 preload API에 손대지 못하고,
  Chromium 사이트 격리로 렌더러 프로세스도 나뉜다 (믿을 수 없는 문서를 푸는 WASM이 host의 저장 IPC와 한 프로세스에 있지 않게). 127.0.0.2를 못 쓰는 PC면 `127.0.0.1`의 다른 포트로 (출처만 분리).
  SDK는 http(s) 출처만 받으므로 `file:`이 아니라 로컬 서버를 쓴다 (`editor-server.mjs`: GET/HEAD만, Host 머리글 확인, 경로 탈출·점 이름·`sw.js`류·`*.cjs` 404 — 8.3 짧은 이름으로 와도 실제 긴 이름으로 다시 확인,
  host HTML에 CSP `frame-src <studio 출처>`, studio HTML에 `studioCsp`: 자기 출처만(`connect-src 'self' data: blob:`, `script-src 'self' 'wasm-unsafe-eval'`), `frame-ancestors <host 출처>`).
- **네트워크 차단** (인터넷 없이, 문서 내용이 밖으로 나가지 않게):
  - 편집 창 세션의 `webRequest`가 두 로컬 출처(와 blob:/data:) 밖의 요청을 모두 취소한다 (빌드에 남은 jsdelivr 글꼴 주소 포함).
  - `webRequest`에 안 잡히는 연결(WebRTC STUN·TURN, WebTransport): 세션 프록시를 죽은 주소(`socks5://127.0.0.1:9`, 이 PC 주소만 바로)로, 창마다 `setWebRTCIPHandlingPolicy('disable_non_proxied_udp')`, studio CSP `connect-src 'self'`.
  - 맞춤법 사전: Chromium이 사전(ko-3-0.bdic, 11MB)을 인터넷에서 받지 않게 편집 창 세션과 주 창(기본 세션) 모두 `setSpellCheckerDictionaryDownloadURL('http://127.0.0.1:9/')`.
  - 권한은 `clipboard-sanitized-write`·`local-fonts`만, 내려받기·다른 곳으로 이동은 막는다.
- **새 창·링크**: `window.open`은 늘 거절하고 주소만 본다. 이 PC·사설망 주소(편집기 서버 포함)는 열지 않고, 바깥 http(s)·mailto는 "기본 브라우저로 열까요?"를 물은 뒤 `shell.openExternal`.
  studio의 문서 링크는 `window.open('about:blank')` 뒤 주소를 바꾸므로, `about:blank`만 숨긴 창(스크립트 끔, 같은 세션)으로 열어 주소를 바꾸려는 순간(`will-navigate`) 가로채 같은 확인을 거친 뒤 창을 없앤다.
- **IPC** (편집 창 host의 최상위 틀에서만, 경로를 받지 않음 — 그 창이 연 파일에만 작용):
  `editor:get-document`, `editor:save`, `editor:save-as`(main의 저장 대화상자) / `editor:dirty`, `editor:log`, `editor:close`, `editor:reply` / main → host `editor:external-change`, `editor:request-close`, `editor:command`.
- **저장**: host가 원래 형식으로 내보냄(hwp·hwt → `exportHwp`, hwpx → `exportHwpx`) → main이 형식 서명(OLE/ZIP)·512MB·읽기 전용을 확인 → 덮어쓰기 전 백업(`getBackups` = 버전 '직접 편집 저장 전', 문서마다 최근 30개) →
  같은 폴더 임시 파일에 쓰고 fsync → 있는 파일이면 내용만 덮어 복사(파일의 ACL·하드 링크가 그대로), 새 파일이면 이름 바꾸기 (`atomicWrite`) → `onSaved(path)`.
  단축키는 `before-input-event`로 받는다 (Ctrl+S, Ctrl+Shift+S, Ctrl+W). 다른 이름으로 저장에서 확장자를 붙인 이름이 이미 있으면 다시 묻는다.
- **저장하는 동안 고친 것**: `notifySaved`는 studio를 무조건 '저장됨'으로 만든다. 그래서 그 뒤 `getDocumentState().documentSha256`(형식이 다르면 다시 내보내 비교)을 저장한 바이트와 견주고,
  다르면(내보낸 뒤 ~ notifySaved 사이에 친 글) host가 '저장 안 됨'(`lateEdits`)으로 둔다 → 닫을 때 묻고, 주 창의 AI 보내기 경고도 그대로. notifySaved 뒤에 고친 것은 studio가 스스로 '바뀜'으로 표시한다.
- **바깥 변경**: 창마다 그 파일의 폴더를 감시하고 내용(sha1)이 같으면 무시한다. 고친 것이 없으면 host가 다시 불러오고, 있으면 알림 띠(다시 불러오기 / 내 편집 유지).
- **닫기**: host에 '저장 안 한 내용이 있는지' 묻고(3초) 있으면 저장 / 저장 안 함 / 취소. host의 `beforeunload`는 무시한다.

main.mjs가 잇는 것:

| 언제 | 무엇을 |
|---|---|
| 시작 | `createEditorManager({ getBackups, onSaved, onDirtyChange, icon, dirs?, testHooks? })` 한 번. 모듈을 못 불러와도 앱은 뜨고 `직접 편집`만 이유를 알린다 |
| `app:init` | `editor: { available, reason, exts }` (묶음이 있는지, 한국어 이유) — 한글·Office 편집기를 종류(kind)마다 따로 알린다 |
| `doc:edit` | 열린 문서가 편집 창 형식(한글 hwp·hwpx·hwt, Office docx·xlsx·pptx)이 아니면 `UNSUPPORTED`(프로그램에서 열기 안내), 그 종류의 묶음이 없으면 `EDITOR_MISSING`(한국어 이유), 아니면 `openEditor` → `{ path, reused }` |
| 문서 정보·`doc:editor` 이벤트 | `editor: { open, dirty }` — 창이 열리고 닫히고 고쳐질 때(`onDirtyChange`, 창 `closed`, 저장 뒤) 주 창에 보낸다 (상태 막대 '편집 창 열림', 대화창 경고 띠) |
| `onSaved` | 최근 목록 맨 위로, 열린 문서면 `doc:changed { by:'editor' }` → 미리보기 새로 고침 (감시보다 먼저, 알림 없이). 백업은 편집 창이 저장 전에 만든다 |
| 파일 감시(주 창) | 바뀌거나 없어지면 `notifyExternalChange(path)` (내용이 같으면 편집 창은 무시) |
| 원래대로 | 되돌린 뒤 `notifyExternalChange(path, { force: 편집 창에 저장 안 한 내용 })` → 결과의 `editorNotified`. 파일이 이미 원본과 같아도 저장 안 한 편집이 있으면 알려 편집 창이 다시 불러오기 / 내 편집 유지를 묻는다 |
| `ai:send` | 그 문서 편집 창에 저장 안 한 내용이 있으면 `EDITOR_DIRTY`로 거절. 화면이 묻고(편집 창 보기 / 그대로 보내기 / 취소) 그대로 보내면 `editorOk:true` |
| `doc:new` | `edit:true`이고 편집 창 형식(한글·docx·xlsx·pptx)이면 만든 뒤 `doc:edit`와 같은 길로 연다 → `editor` 또는 `editorError` (문서는 만들어졌으므로 실패로 돌려주지 않음). 같은 경로의 문서를 지웠다가 다시 만든 경우 예전 '원래대로' 기준·AI 대화는 버린다 |
| 주 창 닫기 | 편집 창이 있으면 `closeAll()` (저장 안 한 창마다 묻기) → 모두 닫혔을 때만 주 창도 닫는다. 취소하면 둘 다 그대로 |

- 시험: `test/editor-server.test.mjs`(서버·저장 도우미), `test/editor-integration.test.mjs`(가짜 electron으로 main.mjs의 위 흐름, `OFFICEKIT_EDITOR_SMOKE=1`이면 실제 앱 캡처),
  끝까지 시험 `tmp/editor-int/e2e.mjs`(CDP로 두 창 조작 — 개발용, 저장소에 넣지 않음).

### 문서·표·슬라이드 편집 창 (kind 'office' — src/app/office-editor/ · vendor/office-editor/)

docx·xlsx·pptx를 앱 안에서 고치는 창. 한글 편집 창과 같은 틀(창·IPC·저장·감시·닫기)에 편집기 종류만 하나 더 있다.
편집기는 **ONLYOFFICE Docs CE 9.4.0(빌드 9.4.0.129)** 을 바탕으로 한 filex-office-editor v0.3.1 묶음이고 **AGPL**이다 (아래 '라이선스 경계').
자세한 설계·근거: [`OFFICE-EDITOR-DESIGN.md`](OFFICE-EDITOR-DESIGN.md).

```
주 창 ─ doc:edit ─→ openEditor(path)  (확장자 docx·xlsx·pptx → kind 'office', 창 제목 '<파일> - 문서 편집 / 표 편집 / 슬라이드 편집 (Office Kit)')
                       │ 문서마다 BrowserWindow 하나 (partition 'officekit-office', 디스크에 안 남음)
                       ▼
  http://127.0.0.1:<A>/office-host/host.html   (src/app/office-editor/host.*, preload office-preload.cjs)
        └ iframe sandbox="allow-scripts"  http://127.0.0.3:<C>/office/<ver>/index.html   (filex 앱 페이지, CSP sandbox → opaque origin)
             └ iframe  …/editor/web-apps/apps/<documenteditor|spreadsheeteditor|presentationeditor>/main/index.html
                  ├ socket.io.min.js 자리 = filex 다리 (Document Server 흉내, 틀 안에서만)
                  └ Worker: x2t.wasm (docx/xlsx/pptx ↔ 편집기 형식)
        ↑ MessagePort (filex 앱 SDK v1): session.get · file.read · file.save{reason} · ui.dirty · ui.download · ui.print · state.* · {hid, request:'save'}
```

- **세 출처**: host(`127.0.0.1:A`, preload API가 붙는 곳), 한글 studio(`127.0.0.2:B`), Office 편집기(`127.0.0.3:C`). 편집기 스크립트는 host에 닿지 못한다.
  127.0.0.3을 못 쓰는 PC면 127.0.0.1의 다른 포트 (출처만 분리). 서버가 받은 포트가 Chromium 금지 포트면 다시 받는다.
- **글꼴**: 묶음에 글꼴이 없다. `office-editor/fonts/`가 이 PC의 글꼴(`C:\Windows\Fonts`, 사용자 글꼴 폴더)로 `AllFonts.js`·미리보기 그림을 만들어
  `userData/office-editor/fonts/<편집기 버전>-<폴더 목록 해시>/`에 두고, 서버가 그 주소에 대신 내준다. 글꼴 파일은 복사하지 않고 읽어서 앞 32바이트만 XOR해 보낸다.
- **IPC**: 한글 편집 창의 `editor:*`를 그대로 쓰고 Office에만 `editor:download`(다른 형식 내려받기), `editor:print`(PDF → 기본 PDF 앱), `editor:state`(편집기 설정 보관),
  `editor:draft`(자동 저장 사본)가 더 있다. `editor:get-document`는 office면 `about: { appVersion, bundle(VERSION.json 고정 값), sources }`도 준다 ('정보 · 라이선스').
- **저장**: Ctrl+S → host가 앱에 `{hid, request:'save'}` → 앱이 x2t로 원래 형식 바이트 → `file.save{reason:'user'}` → `editor:save` →
  main이 형식 확인(ZIP·`[Content_Types].xml`·주 부분) → `src/core/ooxml-fixups.mjs`(pptx 노트 테마 분리 — **필수**, xlsx 남는 테마, docx 혼자 있는 `vMerge`) →
  백업 → `atomicWrite` → `onSaved`. 고치기가 실패하면 원래 바이트를 저장하고 로그만 남긴다.
- **자동 저장**: filex 앱이 스스로 부르는 `file.save{reason:'auto'}`는 원본에 쓰지 않고 `userData/office-drafts/<경로 sha1>.<ext>`에 복구 사본만 둔다.
- **보안**: 한글 편집 창과 같은 네트워크 차단(webRequest·죽은 프록시·`disable_non_proxied_udp`·맞춤법 사전 주소) + 편집기 HTML에 CSP `sandbox allow-scripts`
  (`'unsafe-eval'`은 ONLYOFFICE가 꼭 쓰므로 네트워크 없는 opaque origin 틀 안에 가둔다). host는 편집기 틀 안에 스크립트를 넣지 않는다. 매크로·플러그인 없음.
- **이름**: 창 제목·메뉴·기능 이름에 'ONLYOFFICE'를 쓰지 않는다. 편집기 안 로고·'정보'·filex의 아래쪽 법적 고지 줄은 그대로 둔다.
- 시험: `test/editor-server.test.mjs`, `test/editor-window-office.test.mjs`, `test/office-host.test.mjs`, `test/ooxml-fixups.test.mjs`, `scripts/office-editor/office-editor.test.mjs`(묶음 만들기),
  e2e는 `OFFICEKIT_EDITOR_SMOKE=1` (설계 10장).

**묶음 만들기** (`scripts/build-office-editor.mjs`, 고정 값 `scripts/office-editor/pins.mjs`)
- filex 릴리스 `ui.zip`(크기·SHA-256 고정) → 편집기 파일마다 filex `upstream/editor.lock.json`과 비교, x2t 두 파일은 고정 SHA-256 →
  (패치가 있으면) filex 소스 고정 커밋 + `patches/office-editor/filex/*.patch`로 앱 페이지만 다시 빌드 → 덜어 내기(`scripts/office-editor/rules.mjs`, 파일 통째로만) →
  `vendor/office-editor/` + `VERSION.json` + `NOTICE-OfficeKit.md`(날짜 붙은 수정 고지).
- `beforePack`(`scripts/check-editor-bundle.cjs`)이 묶음·라이선스 파일·`VERSION.json`이 없으면 설치 파일을 만들지 않는다.

**라이선스 경계** ([`LICENSES-OFFICE-EDITOR.md`](LICENSES-OFFICE-EDITOR.md))
- AGPL: `vendor/office-editor/` 전체, `patches/office-editor/`, `src/app/office-editor/fonts/` (파일 머리 `SPDX-License-Identifier`: 패치는 `AGPL-3.0-or-later`, 글꼴 생성기는 `AGPL-3.0-only` + ONLYOFFICE 추가 조항). 전문은 `LICENSES/`, 편집기 묶음 안 기계어(WebAssembly) 제3자 고지는 `THIRD-PARTY-NATIVE.txt` (`scripts/office-editor/third-party.mjs`).
  나머지(호스트 페이지·서버·창·`ooxml-fixups`·묶음 스크립트)는 MIT — 편집기와는 MessagePort·로컬 서버로만 주고받는다. ONLYOFFICE·filex·CryptPad 코드를 MIT 파일에 붙여 넣지 않는다.
- 소스 제공(AGPL 6조): `scripts/make-office-source.mjs` → `office-editor-source-<버전>.zip`
  (filex 소스 고정 커밋 `git archive` 그대로 + 묶음 스크립트·규칙·패치·글꼴 생성기 + `NOTICE-OfficeKit.md`·`VERSION.json`·ONLYOFFICE 라이선스 파일
  + `UPSTREAM.md`: ONLYOFFICE sdkjs·web-apps·core·build_tools `v9.4.0.129` 커밋·x2t 빌드 방법). 빠진 부분이 있거나 고정 값이 서로 다르면 멈춘다.
  같은 입력이면 같은 바이트. 거울 저장소는 만들지 않는다 (사람이 할 일).
- 화면 고지(추가 조항 3): host 위 막대 **정보 · 라이선스** 대화상자. 글의 기준: [`office-editor-legal-ko.md`](office-editor-legal-ko.md).

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
- 한글 편집기 묶음 `vendor/rhwp-studio`·`vendor/rhwp-editor-sdk`도 들어간다. `npm run dist`/`release`와 CI 배포 작업이 먼저 `npm run editor:build`를 돈다.
  electron-builder를 바로 실행해도 `beforePack`(`scripts/check-editor-bundle.cjs`)이 묶음이 없으면 멈춘다. CI 시험 작업은 빌드 스크립트 문법·`--help`를 늘 확인하고, main push에서는 실제 빌드도 한다.
  - 버전 고정: `@rhwp/core`(정확히 0.8.7, `^` 없이)·`@rhwp/editor`·`build-editor.mjs`의 `TAG`·`COMMIT`·`CORE_VERSION`·`EDITOR_VERSION`은 함께 올린다.
  - `scripts/build-editor.mjs`: rhwp 태그 `v0.8.7`을 성긴 복제(커밋 `1a76570e…` 확인) → `@rhwp/core` 0.8.7 파일을 `pkg/`에 넣고 Rust 없이 vite 빌드
    (`RHWP_WITHOUT_HWPCTRL=1`, `RHWP_DISABLE_EXTERNAL_WEBFONTS=1`, `--base=./`) → 샘플·서비스 워커(PWA)·manifest를 빼고 index.html에서 등록 코드를 지운다
    (남아 있으면 빌드 실패) → 라이선스 파일과 `VERSION.json`(파일마다 크기·sha256)을 넣어 `vendor/`로 바꿔 놓는다. 같은 것이 있으면 건너뛴다.
  - `@rhwp/editor`(SDK)는 devDependency로 받아 필요한 세 파일만 `vendor/rhwp-editor-sdk`에 복사한다 (`node_modules`로는 배포하지 않음).
- 한글→PDF 변환 프로그램 `vendor/rhwp-cli`(rhwp.exe 약 47MB, 설치 파일 약 +14MB)도 들어간다. `npm run dist`/`release`와 CI가 `npm run rhwp-cli:build`를 돈다.
  - `scripts/build-rhwp-cli.mjs`: rhwp v0.8.7 Release의 `rhwp-v0.8.7-windows-x86_64.zip`(고정 크기·SHA-256) → `rhwp.exe`·`LICENSE`(고정 SHA-256)를 꺼내고
    같은 태그 커밋의 `THIRD_PARTY_LICENSES.md`(고정 SHA-256)와 `VERSION.json`을 쓴다. `rhwp.exe --version`이 `rhwp v0.8.7`인지, `@rhwp/core`가 정확히 0.8.7인지 확인.
    `--verify`는 받지 않고 고정 값과 맞춰 보기만. CI는 받은 zip을 캐시하고(키 = 스크립트 해시), 시험 작업도 먼저 받는다(`test/hwp-pdf.test.mjs`).
  - `beforePack`이 `vendor/rhwp-cli`의 네 파일과 `VERSION.json`에 적힌 크기를 확인한다 (없으면 설치 파일을 만들지 않음).
  - rhwp를 올릴 때 `@rhwp/core`·`build-editor.mjs`·`build-rhwp-cli.mjs`를 함께 올린다 (미리보기·편집기·PDF가 같은 엔진 판).
- Office 편집기 묶음 `vendor/office-editor`도 들어간다 (설치 파일 약 +50~65MB, 설치 폴더 약 +213MB). `npm run dist`/`release`와 CI가 `editor:build` 다음에 `office-editor:build`를 돈다.
  - 고정 값(filex 태그·커밋·`ui.zip` SHA-256, x2t SHA-256, ONLYOFFICE 판)은 `scripts/office-editor/pins.mjs` 한 곳. 올릴 때 `make-office-source.mjs`의 소스 주소(같은 태그의 커밋)와
    `docs/office-editor-legal-ko.md`도 함께 바꾼다 (`make-office-source.mjs`가 서로 맞는지 확인하고, 다르면 멈춘다).
  - 릴리스마다 `node scripts/make-office-source.mjs`로 `office-editor-source-<버전>.zip`(+ `.sha256`)을 만들어 같은 Release에 붙이고, 공개 전에 붙었는지 확인한다 (AGPL 6조).
- 탐색기 오른쪽 클릭 "Office Kit로 열기" (기본 연결 프로그램은 바꾸지 않음).
- GitHub Actions: `v*` 태그를 올리면 Release에 설치 파일 업로드.
