# Office 편집기 설계 — Word·Excel·PowerPoint를 앱 안에서 직접 편집

> 상태: 설계안 (2026-10-18). 아직 코드 없음. 근거 실험은 `tmp/spike-*`(git 밖)에 있다.

## 한눈에

- **목표**: docx·xlsx·pptx를 Office Kit 창 안에서 고치고 저장한다. Microsoft Office 없이, 인터넷 없이.
- **결정**: **filex-office-editor v0.3.1**을 바탕으로 한다.
  - 안에 든 것: ONLYOFFICE Docs Community Edition 9.4.0(빌드 9.4.0.129) 편집기 파일(**손대지 않은 공식 파일**), 같은 태그로 빌드한 x2t(WASM 변환기), Document Server를 흉내 내는 다리 코드.
  - Office Kit이 더하는 것: 호스트 페이지(약 150줄), 시스템 글꼴 목록 만들기, 저장 뒤 OOXML 고치기, 네트워크 차단, 법적 고지 화면.
- **판정 실험** (filex 시제품을 직접 다시 돌림):
  - 9개 문서(형식 3개 × 만든 방법 3개)를 열고 고치고 저장했다. 9개 모두 Word·Excel·PowerPoint에서 경고 없이 열렸다.
  - 열기 1.1~2.2초, 저장 0.04~0.2초. 밖으로 나간 요청 0건.
- **꼭 넣을 것 두 가지** (이번 판정에서 확인):
  1. pptx 저장 뒤 고치기: 빼면 Office Kit '새 문서'로 만든 pptx를 저장했을 때 PowerPoint가 복구 창에서 멈춘다.
  2. `setWebRTCIPHandlingPolicy('disable_non_proxied_udp')`: 빼면 편집기 안에서 STUN(UDP)이 인터넷에 닿는다.
- **크기**: 설치 파일 +50~65MB (지금 131MB → 약 185~195MB). 설치 폴더 +214MB, 파일 약 1,530개.
- **일**: 한 사람 3~4주. 일감 9개로 나눠 3명이 하면 약 2주.
- **다음 행동 하나**: W1(묶음 만들기 스크립트)과 W6(저장 뒤 고치기)부터 시작한다. 둘 다 다른 일감을 기다리지 않는다.

---

## 1. 결정과 버린 안

### 고른 것: filex-office-editor 바탕 + 다른 실험의 좋은 부분

| 부분 | 가져오는 곳 | 까닭 |
|---|---|---|
| 편집기 파일 (web-apps, sdkjs) | filex v0.3.1 `ui.zip` 안의 `editor/` = ONLYOFFICE 공식 CE 이미지 9.4.0.1 (digest 고정) | 공식 파일 그대로라 출처·라이선스가 깨끗하다. 파일마다 SHA-256이 `upstream/editor.lock.json`에 있다 |
| 변환기 x2t (WASM) | 같은 `ui.zip`의 `x2t/` (ONLYOFFICE core v9.4.0.129 + CryptPad WASM 패치) | 편집기와 **같은 버전**이다. 빌드 방법이 고정되어 있어 다시 빌드하면 같은 바이트가 나온다 |
| Document Server 흉내 | filex `socket.io.min.js` 자리의 다리(bridge·shim) | 편집기 파일을 고치지 않는다. 자물쇠(lock) 규칙은 ONLYOFFICE 서버 코드를 옮긴 것이다 |
| 호스트 ↔ 앱 약속 | filex 앱 SDK v1 (MessagePort) | 메서드 10개 남짓이라 호스트가 작다 |
| 글꼴 목록 만들기 | CryptPad 실험 `gen-allfonts.mjs`(모든 설치 글꼴, 한국어 이름) + filex 실험 `make-allfonts.mjs`(9.4 형식, XOR) | 묶음에 한글 글꼴이 없어 그대로 두면 한글이 □로 보인다 |
| 글꼴 미리보기 그림 | CryptPad 실험 `gen-thumbs.mjs` (resvg — 이미 의존성) | 글꼴 목록에서 이름 미리보기가 나오게 |
| 저장 뒤 고치기 | 세 실험 공통 (pptx 테마) + oow 실험 (xlsx 남는 테마) | PowerPoint 복구 창을 막는다 |
| 실제 IME 시험 도구, Office 엄격 열기 | CryptPad 실험 `ime-send.ps1`·`strict-open.ps1` | `strict-open.ps1`은 시험 계획에 그대로 쓴다. `ime-send.ps1`(OS 키 입력)은 쓰지 않고 손으로 확인 |

### 버린 안

| 안 | 버린 까닭 | 그래도 가져오는 것 |
|---|---|---|
| CryptPad 빌드 (onlyoffice-editor v9.3.2+3) | v9부터 ONLYOFFICE가 아니라 **Euro-Office 포크(9.3.0.140)** 기반이다. 머리글에서 로고 유지 조항을 지웠고, ONLYOFFICE와 다툼이 남아 있다. zip에 web-apps·sdkjs LICENSE 파일이 없다. 재배포할 수 없는 Arial·Calibri가 들어 있다. CryptPad만의 패치(`connectMockServer`, `parent.APP`)에 묶인다 | 글꼴 목록·미리보기 생성기, 실제 IME 시험, 엄격 열기, 안전하지 않은 포트 피하기 |
| onlyoffice-web (yuhaibohotmail) | 공식 CE 파일이고 법적 고지도 좋다. 하지만 TS 1.5만 줄에 React가 붙고, 실험에서 열기가 8초였다. 마지막 e2e(run4)는 다시 불러오기에서 실패했다 | 레지스트리 API로 Docker 없이 이미지 층을 받는 방법, xlsx 남는 테마 지우기, 법적 고지 화면 아이디어 |
| Document Server를 Docker·WSL로 띄우기 | 사용자 PC에 서버·Docker·WSL이 필요하다. 무겁고 설치가 깨지기 쉽다 | — |
| 9.4용 Document Server 흉내를 처음부터 새로 짜기 | filex가 이미 한 일(약 1.5천 줄 + 시험 280개)을 되풀이하게 된다. filex는 AGPL이라 버려지면 포크해서 이어 갈 수 있다 | — |
| Microsoft Office COM으로 편집 | Office가 있어야 한다 — 목표와 반대다 | (검증에만 씀) |
| LibreOffice·Collabora | 사용자가 ONLYOFFICE 계열을 골랐다. 더 무겁다 | — |

### 세 실험 비교 (판정)

| | CryptPad | **filex (선택)** | onlyoffice-web |
|---|---|---|---|
| 편집기 출처 | Euro-Office 포크 9.3 | ONLYOFFICE CE 9.4.0.129 공식, 그대로 | ONLYOFFICE CE 9.4.0.1 공식 |
| x2t | CryptPad v9.3.2+3 | core 9.4.0.129 (버전 같음) | CryptPad 빌드 |
| 호스트 코드 | oo.js 283줄 (CryptPad 패치 필요) | host.js 102줄 | vite 빌드 + React 컴포넌트 |
| 열기 | 1.7~2.2초 | 1.1~2.2초 (판정 다시 돌림) | 8~8.5초 |
| 격리 | 다른 사이트(127.0.0.2) | CSP `sandbox` → opaque origin | 같은 출처 |
| 실제 한국어 IME | 확인함 (MS IME, 9.3) | CDP 조합만 (아래 '한계') | CDP만 |
| PDF 내보내기 | 없음 | 됨 (x2t, 글꼴 포함) | 일부 |
| 판정 | 2위: 출처·라이선스 위험 | **1위** | 3위: 크고 느림 |

---

## 2. 판정 실험 — 다시 돌려 본 것

`tmp/spike-judge/`에서 filex 시제품을 복사해 경로만 바꿔 돌렸다. 문서는 새로 만들었다.

| 문서 | 만든 방법 |
|---|---|
| `judge.*` | `createDocument` (docx 라이브러리·exceljs·pptxgenjs) — 한국어 마크다운 |
| `blank.*` | `createBlankDocument` — **Office Kit '새 문서' 경로 그대로** |
| `office.*` | Word·Excel·PowerPoint COM으로 꾸민 문서 (표 서식, 머리글, 수식, 병합, 도형) |

결과:

| 확인 | 결과 |
|---|---|
| 열기·고치기·저장 (9개) | 모두 됨. 문단 추가·굵게, 표 행 추가, 셀 값·`=SUM`, 행 삽입, 슬라이드 추가·글상자 |
| Office 엄격 열기 (경고 모두 켬, 복구 끔, 감시 타이머) | **9/9 통과** (pptx 고치기 적용) |
| pptx 고치기 없이 | `judge-raw.pptx`·`blank-raw.pptx` **둘 다 PowerPoint 복구 창에서 멈춤** |
| Word에서 다시 읽기 | 추가 문단·굵게·표 새 행이 그대로. 단 표가 `Uniform=false`가 됨 (아래 vMerge) |
| Excel에서 다시 읽기 | `60 {=SUM(B10:D10)}` — 행 삽입 뒤 수식 참조도 맞게 밀림 |
| 네트워크 | 문서마다 로컬 요청 261~278개, 막힌 것 0, 404 0 |
| 일부러 밖으로 보내기 | `session.fetch` 4곳, 편집기 틀 `fetch` 3곳, WebSocket — 모두 실패 |
| WebRTC | **정책 없이**: STUN 응답 후보(srflx)가 생김 = UDP가 인터넷에 닿음 ✗ / **정책 켬**: host·app·editor 세 틀 모두 후보 0 ✓ |
| 한국어 조합 (CDP) | '판정', '세종', '한글'이 겹침·빠짐 없이 들어감 |
| 실제 MS IME (SendInput) | 10키 뒤 사용자 Chrome이 앞으로 와 **스스로 멈춤** → 판정에서는 확인 못 함 (출시 전 수동 시험 필수) |

증거: `tmp/spike-judge/out/j1/*.strict.json`, `tmp/spike-judge/shots/j1-*.png`, `tmp/spike-judge/out/j1/office.docx-p1.png`(Word 렌더).

---

## 3. 구조

한글 편집 창과 같은 모양으로 만든다. 편집기 종류(kind)만 하나 더 생긴다.

```
주 창(renderer) ─ doc:edit ─→ main.mjs editDocument() ─→ editorManager.openEditor(path)   (kind = 확장자로 고름)
                                                            │ 문서마다 BrowserWindow 하나 (partition 'officekit-office', 디스크에 안 남음)
                                                            ▼
  http://127.0.0.1:<A>/office-host/host.html   (src/app/office-editor/host.*, preload office-preload.cjs)
        └ iframe sandbox="allow-scripts"  http://127.0.0.3:<C>/office/<ver>/index.html   (filex 앱 페이지, opaque origin)
             └ iframe  …/editor/web-apps/apps/<documenteditor|spreadsheeteditor|presentationeditor>/main/index.html
                  ├ socket.io.min.js 자리 = filex 다리 (Document Server 흉내, 틀 안에서만)
                  └ Worker: x2t.wasm (docx/xlsx/pptx ↔ Editor.bin)
        ↑ MessagePort (filex 앱 SDK v1): session.get · file.read · file.save · ui.dirty · ui.download · ui.print · state.* · {hid, request:'save'}
```

| 출처 | 내주는 것 | 왜 따로 |
|---|---|---|
| `127.0.0.1:A` | `/host/*`(한글), `/office-host/*`(Office 호스트) | preload API가 붙는 곳. 편집기 스크립트가 닿지 못해야 한다 |
| `127.0.0.2:B` | `/studio/*` (rhwp-studio, 지금 그대로) | — |
| `127.0.0.3:C` | `/office/<ver>/*` (vendor/office-editor) + 글꼴 경로 | 다른 사이트 → 렌더러 프로세스가 나뉜다. 거기에 CSP `sandbox`로 opaque origin까지 |

- 127.0.0.3을 못 쓰는 PC면 127.0.0.1의 다른 포트로 (출처만 분리 — 한글 편집기와 같은 대비책).
- 서버가 `listen(0)`으로 받은 포트가 Chromium 금지 포트(1719, 1720, 1723, 2049, 3659, 4045, 4190, 5060, 5061, 6000, 6566, 6665~6669, 6679, 6697, 10080)면 다시 받는다.
  이 PC는 임시 포트가 낮은 번호(2161, 3982, 4017, 9357…)에서 나와 실제로 걸릴 수 있다. **한글 편집기에도 같은 버그가 있다** (`editor-server.mjs` `listen`).

---

## 4. 묶음 만들기 (asset pipeline)

### `scripts/build-office-editor.mjs` (`npm run office-editor:build`)

| 단계 | 하는 일 | 실패하면 |
|---|---|---|
| 1. 받기 | filex 릴리스 `ui.zip`을 캐시 폴더로 (`OFFICEKIT_OFFICE_EDITOR_CACHE`, 기본 `tmp/office-editor-build`). 이어받기, 3번 다시 | 멈춤 |
| 2. 확인 | 크기·SHA-256이 고정 값과 같은지 | 멈춤 (파일 지움) |
| 3. 풀기 | jszip(이미 의존성)으로 캐시에 풂 | 멈춤 |
| 4. 파일별 확인 | `editor/` 파일마다 filex `upstream/editor.lock.json`(고정 커밋에서 받고 그 파일의 SHA-256도 고정)과 비교. 다른 것은 `socket.io.min.js`(filex 다리) 하나뿐이어야 함. x2t 두 파일은 고정 SHA-256 | 멈춤 |
| 5. (패치가 있을 때) 앱 다시 빌드 | filex 소스를 고정 커밋으로 받고(`git` sparse) → `patches/office-editor/filex/*.patch` 적용 → `npm ci --ignore-scripts` → `node scripts/build-app.mjs --editor <4의 editor> --x2t <4의 x2t> --no-zip` → 앱 페이지·`filex/*`·`socket.io.min.js`만 가져옴. Docker 필요 없음 (filex 실험에서 Windows·Node 24로 릴리스와 같은 바이트 확인) | 멈춤 |
| 6. 덜어 내기 | 아래 표 | — |
| 7. 쓰기 | `vendor/.office-editor.tmp`에 다 쓴 뒤 `vendor/office-editor`로 바꿔 넣음. `VERSION.json`(고정 값, 규칙 해시, 파일마다 크기·SHA-256), `NOTICE-OfficeKit.md`(날짜 붙은 수정 고지) | 이전 묶음 그대로 |
| 8. 건너뛰기 | `VERSION.json`의 고정 값·규칙 해시가 같으면 아무것도 안 함. `--force`로 다시 | — |

고정 값 (올릴 때는 여기만 바꾸고 함께 올린다):

| 이름 | 값 |
|---|---|
| filex 태그·커밋 | `v0.3.1` · `57dda3dae99af78eb26b3c64b69f1e3f711f0757` (태그 커밋 — 다음 main 커밋 `acc87f88da47`은 README.md만 다름) |
| `ui.zip` | `https://github.com/BRF-Tech/filex-office-editor/releases/download/v0.3.1/ui.zip` · 103,098,008바이트 · SHA-256 `cd3bc0741f97e471ec26d639da65835fccff2605fc255974ecb5f21d2b179baf` |
| x2t | `x2t.js` `9f2f65ac557470b2895beaa0f9678ee1a998545d49c8b93d67306ff3e7d7fe4b` · `x2t.wasm` `cabaa6e33574f61703c53463d714e33529949d17d388b4e198994199e29c368b` (filex 릴리스 `v9.4.0.129+2`) |
| ONLYOFFICE | Docs CE 9.4.0 빌드 9.4.0.129 · `docker.io/onlyoffice/documentserver@sha256:3ab6ebc7c605e5a32b7ae3ff19daed4925090245acc8100ce2230bd766c88212` · 소스 태그 `v9.4.0.129` |
| x2t 소스 | core `a016fc2868b9f99f59dd30aa2088e68bbe5104af`, build_tools `cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5` (filex `upstream/onlyoffice.json`) |

### 덜어 내기 (판정에서 잰 값)

| 빼는 것 | 줄어듦 | 까닭 |
|---|---|---|
| 편집기 언어 파일 (ko·en 말고) | −46.0MB, 264개 | 한국어 UI + 영어 대비 |
| 수식 언어 (ko·en 말고) | −6.3MB | |
| 휴대폰·embed·forms 앱 | −18.9MB, 611개 | 데스크톱 창만 씀 |
| 묶음 글꼴 `editor/fonts` (+ 그 라이선스) | −38.5MB | 시스템 글꼴을 쓴다 (5장). 묶음에는 한글 글꼴도 없다 |
| **남는 것** | **213.7MB, 1,530개, gzip-9 64.1MB** | sdkjs slide 50.6 · x2t 37.0 · cell 33.9 · word 30.9 · common 30.3 · web-apps 29.6 |

- 빼지 않는 것: `LICENSE.txt`, `3rd-Party.txt`, `license/`, filex `LICENSE`·`NOTICE`·`editor/filex/CHANGES.txt`, 모든 머리글. **파일을 통째로 빼기만 하고, 남는 파일은 한 바이트도 고치지 않는다** (줄이기·압축 다시 하기 없음).
- 슬라이드 테마(약 20MB)는 '디자인' 탭에 필요해 남긴다.

### CI · 배포

- `package.json`: `"office-editor:build"`, `dist`·`dist:dir`·`release`가 `editor:build` 다음에 부른다. `build.files`에 `vendor/office-editor/**/*`.
- `scripts/check-editor-bundle.cjs`(beforePack): `vendor/office-editor/index.html`, `x2t/x2t.wasm`, `editor/LICENSE.txt`, `VERSION.json`이 없으면 멈춘다.
- `.github/workflows/release.yml`:
  - test 작업: `node --check` + `--help`는 늘, main push에서는 실제 빌드.
  - release 작업: "Office 편집기 묶음 (filex v0.3.1 → vendor/office-editor)" 단계를 한글 편집기 빌드 다음에.
  - `actions/cache`: 키 = 고정 값 해시, 경로 = 캐시 폴더 (100MB 받기를 줄임).
  - 소스 꾸러미 올리기 (8장 4번).

---

## 5. 글꼴

| 문제 | 해결 |
|---|---|
| 묶음 글꼴에는 쓸 만한 한글이 없다 → 한글이 □ (filex 실험 `shots/explore-docx-1.png`) | PC마다 설치된 글꼴로 `AllFonts.js`를 만든다 |
| 문서의 '맑은 고딕'·'바탕'이 이름으로 안 찾아진다 | 한국어 이름(맑은 고딕, 바탕, 굴림, 돋움, 궁서, 함초롬바탕…)을 별칭으로 넣는다 |
| Calibri·Arial 문단 안 한글 | 유니코드 대체 범위의 한글 칸을 Malgun Gothic으로 |
| 글꼴 목록 미리보기가 비어 있다 | `fonts_thumbnail*.png`를 다시 그린다 (resvg) |

- 만드는 곳: `src/app/office-editor/fonts/allfonts.mjs` (worker_thread에서 돌림, 주 창은 멈추지 않음).
  - 읽는 폴더: `C:\Windows\Fonts`, `%LOCALAPPDATA%\Microsoft\Windows\Fonts`.
  - 형식: 9.4 sdkjs `common/libfont/map.js`의 `CFontSelect.fromStream`이 읽는 v2 이진 + 글꼴 목록 + 대체 범위.
  - CryptPad 실험 것(모든 글꼴, fontkit — 598파일·285가족·1.5초)과 filex 실험 것(9.4 형식 확인됨)을 합친다.
- 캐시: `userData/office-editor/fonts/<편집기 버전>-<폴더 목록 해시>/`. 목록(이름·크기·시각)이 바뀌면 다시 만든다.
  - 목록 1.5초는 첫 편집 창을 열 때 기다린다. 미리보기 그림(약 24초)은 앱이 한가할 때 뒤에서 만든다. 다 되기 전에는 미리보기만 비어 있다.
- 내주기: `/office/<ver>/editor/sdkjs/common/AllFonts.js`·`Images/fonts_thumbnail*` → 캐시. `/office/<ver>/editor/fonts/<번호>` → 실제 글꼴 파일을 읽기 전용으로 열고, **앞 32바이트만 XOR**해서 보낸다 (ONLYOFFICE가 기대하는 모양). 복사하지 않는다. **설치 파일에 편집용 글꼴을 넣지 않는다** (x2t.wasm 안에 ONLYOFFICE core가 넣어 둔 URW++ PDF 기본 14 글꼴만 있다 — 8장).
- PDF 내보내기·인쇄: filex 틀이 쓰인 글꼴을 x2t에 넘긴다 → 한글이 PDF에 들어간다 (filex 실험 `out/download-com-test.pdf`).
- 시험의 기준: 생성기를 묶음의 core-fonts 폴더에 돌려, 묶음에 든 `AllFonts.js`(ONLYOFFICE allfontsgen이 만든 것)와 같은 결과인지 비교한다.

---

## 6. 실행 (runtime)

### 호스트 ↔ 앱 약속 (filex 앱 SDK v1, MessagePort)

| 앱 → 호스트 | 호스트가 하는 일 |
|---|---|
| `session.get` | `{ v:1, locale:'ko-KR', theme:{mode}, user:{name: Office에 적힌 사용자 이름(HKCU\Software\Microsoft\Office\Common\UserInfo\UserName) → 없으면 Windows 계정 이름 → 그것도 없으면 '나'}, files:[{index:0, name, ext, size, mime, readOnly}], grants:[files:read, files:write, ui, ui:eval, ui:wasm-eval, ui:package-fetch, ui:frame-package, ui:connect-blob, ui:download, ui:print] }` |
| `file.read` | 원본 바이트 (ArrayBuffer, transfer) |
| `file.save {data, reason}` | 바이트는 이미 원래 형식. `reason:'user'`면 저장, `'auto'`면 복구용 사본 (아래 '자동 저장'). `reason`은 패치로 더한다 (W2) |
| `ui.dirty {dirty}` | 창 제목 `*`, 주 창 '편집 창 열림·저장 안 됨' |
| `ui.download {name, data, mime}` | '다른 형식으로 내려받기'·PDF → main의 저장 대화상자 |
| `ui.print {data: PDF}` | 임시 PDF → 기본 PDF 앱으로 열기 (또는 저장 대화상자). 창마다 처음 한 번 '인쇄용 PDF를 기본 PDF 앱으로 열까요?'를 묻고, 한 번에 하나·3초에 한 번까지 (틀이 PDF 앱을 마구 띄우지 못하게) |
| `state.get` / `state.set` | 편집기 설정 (값마다 8KB까지) → `store.mjs`의 `officeEditor.settings` |
| `ui.toast` / `ui.title` / `ui.close` | 로그 / 무시 / 창 닫기 흐름 |
| `license.get` | `{status:'free'}` |
| `coedit.*` | `unknown_method` → 혼자 편집 |

| 호스트 → 앱 | 언제 |
|---|---|
| `{hid, request:'save'}` | Ctrl+S, 호스트 막대 '저장', 닫기 대화상자의 '저장' → 앱이 `file.save` 뒤 `{hid, result:{saved:true}}` |
| `{event:'close.request'}` | 창 닫기 직전 (설정 보관) |
| `{event:'theme', data:{mode}}` | 앱 테마 바뀜 |

- 손잡기: 앱이 `{type:'filex:hello'}`를 보내면 호스트가 `ev.source === iframe.contentWindow`를 확인하고 `postMessage({type:'filex:port', v:1}, '*', [port2])`.
- 호스트는 편집기 틀 안에 **스크립트를 넣지 않는다** (편집기 안에서 잡히지 않은 예외가 나면 ONLYOFFICE가 오류 창을 띄우고 읽기 전용이 된다 — 두 실험 모두 봄).

### IPC (한글 편집 창의 것을 일반화)

지금의 `editor:*` 채널을 그대로 쓰고, 창 정보(entry)에 `kind`('hwp' | 'office')를 둔다. Office에만 셋이 더 생긴다.

| 채널 | 방향 | 비고 |
|---|---|---|
| `editor:get-document` | host → main | `+ appUrl, locale` |
| `editor:save(bytes, {reason})` | host → main | 바이트 확인 → 고치기 → 백업 → `atomicWrite` |
| `editor:save-as(bytes, ext)` | host → main | 같은 형식만 (docx→docx). 다른 형식은 `editor:download` |
| `editor:download(name, bytes, mime)` | host → main | **새로 생김**. 저장 대화상자 |
| `editor:print(pdfBytes)` | host → main | **새로 생김** |
| `editor:state(op, key, value)` | host → main | **새로 생김**. 편집기 설정 보관 |
| `editor:dirty`, `editor:log`, `editor:close`, `editor:reply` | host → main | 그대로 |
| `editor:external-change`, `editor:request-close`, `editor:command` | main → host | 그대로 |

보낸 쪽 확인도 그대로: 그 창의 최상위 틀 + host 출처만.

### 열기

1. 주 창 `직접 편집`(Ctrl+E) → `doc:edit` → `openEditor(path)`. 확장자 `docx·xlsx·pptx` → kind `office`.
2. 열기 전 거르기 (한국어 안내):
   - 암호 걸린 OOXML(CFB 안에 `EncryptionInfo`) → "암호가 걸린 문서는 직접 편집할 수 없습니다. 프로그램에서 여세요."
   - 512MB 넘음 → 거절 (지금 규칙).
3. host가 바이트를 받아 앱 틀을 만든다 → 앱이 x2t로 `Editor.bin`을 만든다 → 편집기가 뜬다 (판정 1.1~2.2초).
4. 문서가 읽기 전용이면 `readOnly:true` → 보기 모드.

### 저장

```
Ctrl+S (main before-input-event) ─→ host ─{hid, request:'save'}─→ 앱: asc_nativeGetFile + x2t → file.save{reason:'user'}
   → host → editor:save → main: 형식 확인 → ooxml-fixups → 백업 → atomicWrite → onSaved → 미리보기 새로 고침
```

- 형식 확인: ZIP 서명 + `[Content_Types].xml` + 주 부분(`word/document.xml`·`xl/workbook.xml`·`ppt/presentation.xml`).
- 저장 뒤 고치기 `src/core/ooxml-fixups.mjs` (JSZip, 저장할 바이트에만):

| 형식 | 고치는 것 | 근거 |
|---|---|---|
| pptx | 노트 마스터가 슬라이드 마스터와 같은 테마를 쓰면 테마를 복사해 따로 줌 | **필수**. 판정: 없으면 pptxgenjs 문서(새 문서 포함)가 PowerPoint 복구 창에서 멈춤 |
| xlsx | 어디서도 안 쓰는 `xl/theme/theme.xml` 지움 | Excel은 견디지만 깔끔하게 (oow 실험) |
| xlsx | 차트 계열의 항목(`c:cat`) 범위: 값 범위와 같은 방향인데 길이가 다르면(시작이나 끝이 같을 때) 값 범위에 맞춤. 원본과 범위 글이 같은 계열은 원본 항목 글(캐시)을 시트에서 따라가 범위를 옮기고 늘림. 고친 범위의 캐시는 시트 값으로 다시 씀. 확실하지 않으면 그대로 두고 경고(로그)만. 원본에도 같은 범위 글로 있던 길이 차이는 그대로 | 편집기는 데이터 범위 안에 행을 넣으면 값 범위만 늘리거나(행 머리글 → '위의 행 삽입') 둘 다 그대로 둔다(셀 하나 → 행 삽입) → Excel에서 새 막대에 항목 이름이 없거나 마지막 행이 차트에서 빠짐 (v0.1.4 최종 평가, `test/fixtures/office-editor/carry/chart*.xlsx`) |
| docx | 혼자 있는 `<w:vMerge w:val="restart"/>`(아래 행이 이어 받지 않음) 지움 | 편집기에서 추가한 표 행에 붙음 → Word가 표를 '세로 병합'으로 보고 `Rows.Item`이 실패 → **Office Kit 자신의 `word_edit` 표 작업이 깨질 수 있음** (판정: `Uniform=false`) |

- 고치기가 실패하면 원래 바이트를 저장하고 로그만 남긴다 (저장을 막지 않음). 단 pptx 고치기 실패는 알림 띠로 알린다.
- 저장하는 동안 친 글: 앱이 저장 뒤 `ui.dirty`를 다시 보낸다 → 그대로 따른다.

### 자동 저장

- filex 앱은 고친 채 10분이 지나면 스스로 `file.save`를 부른다. 지금은 사용자 저장과 구별할 수 없다.
- 정책: **사용자 파일을 몰래 덮어쓰지 않는다.**
  - 패치 W2로 `file.save`에 `reason`을 더한다 (filex에 PR로도 보냄).
  - `'auto'` → `userData/office-drafts/<경로 sha1>.<ext>`에 복구용 사본. 정상 저장·정상 닫기에서 지운다.
  - 다음에 같은 파일을 열 때 사본이 파일보다 새로우면 띠: "자동 저장된 편집 내용이 있습니다 — 복구 / 버리기".

### 바깥 변경 · 원래대로 · 닫기 · 새 문서 · AI

| 상황 | 동작 (한글 편집 창과 같음) |
|---|---|
| 다른 프로그램·AI가 파일을 바꿈 | 폴더 감시 → 내용이 다르면 띠. 고친 게 없으면 앱 틀을 새로 만들어 다시 열기, 있으면 '다시 불러오기 / 내 편집 유지' |
| 주 창 '원래대로' | `notifyExternalChange(path, {force, reason:'restore'})` |
| 창 닫기 | host에 dirty를 묻고(3초) 저장 / 저장 안 함 / 취소. 저장은 `{hid, request:'save'}` |
| 새 문서 (`doc:new`, edit:true) | docx·xlsx·pptx도 만든 뒤 같은 길로 연다 (pptx 고치기가 꼭 필요한 경로) |
| AI 보내기 | 그 문서의 편집 창에 저장 안 한 내용이 있으면 `EDITOR_DIRTY` (지금 규칙) |
| 렌더러가 죽음 | "편집 창이 예기치 않게 멈췄습니다" → 마지막 저장본 다시 열기 + 복구 사본 안내 |

### 창 · 이름

- 제목: `<파일> - 문서 편집 (Office Kit)` / `표 편집` / `슬라이드 편집`. **'ONLYOFFICE'를 제목·메뉴·기능 이름에 쓰지 않는다** (8장).
- host 위 막대(28px): 파일 이름, 저장 상태, `저장`, `정보 · 라이선스`.
- 한 번에 열린 Office 편집 창이 4개를 넘으면 "메모리를 많이 씁니다" 안내 (창마다 약 0.5GB).

---

## 7. 보안

| 막을 것 | 어떻게 | 판정 확인 |
|---|---|---|
| HTTP(S)·WebSocket | 세션 `webRequest`: host·office 두 출처와 `blob:`·`data:` 말고 모두 취소 | ✓ (fetch·WS 실패) |
| TCP로 새는 것 (TURN-TCP, WebTransport) | 세션 프록시 `socks5://127.0.0.1:9`. 바로 가는 것은 우리 서버 두 곳의 **주소:포트**뿐 — `proxyBypassRules: '<-loopback>;127.0.0.1:<A>;127.0.0.3:<C>'` (Chromium이 늘 바로 보내는 127.0.0.0/8 규칙을 `<-loopback>`으로 끔 → 이 PC의 다른 서비스로 TURN-TCP 못 감) | ✓ (보안 검증: TURN-TCP 카나리 1 → 0) |
| **DNS** (`<link rel=dns-prefetch>`로 문서 내용을 DNS 이름에 실어 보내기) | 앱 시작 전 `--host-resolver-rules="MAP * ~NOTFOUND, EXCLUDE localhost, EXCLUDE 127.0.0.1~3"` (main.mjs) — 이 PC 밖 이름은 풀지 않는다. 세션 요청 막기·프록시·CSP·`X-DNS-Prefetch-Control`은 DNS를 막지 못한다 | ✓ (보안 검증: netlog에 DNS_TRANSACTION 0) |
| **UDP STUN** | **창마다 `webContents.setWebRTCIPHandlingPolicy('disable_non_proxied_udp')`** | ✓ 없으면 srflx 생김, 켜면 후보 0 |
| 맞춤법 사전 받기 | `setSpellCheckerDictionaryDownloadURL('http://127.0.0.1:9/')`, 맞춤법 끔 | ✓ |
| 편집기가 preload에 닿기 | preload는 host 최상위 틀에만. 앱·편집기 틀은 CSP `sandbox allow-scripts`(opaque origin)·다른 사이트(127.0.0.3) | ✓ |
| 편집기 스크립트 범위 | 묶음 HTML CSP: `default-src 'none'; script-src <pkg> 'unsafe-eval' 'wasm-unsafe-eval'; style-src <pkg> 'unsafe-inline'; img-src <pkg> data: blob:; font-src <pkg> data:; connect-src <pkg> blob:; worker-src blob:; frame-src <pkg>; object-src 'none'; base-uri 'none'; form-action 'none'; sandbox allow-scripts`. 나머지 파일 `Access-Control-Allow-Origin: *`·`Cross-Origin-Resource-Policy: cross-origin` | ✓ |
| host 페이지 | CSP `default-src 'none'; script-src 'self'; frame-src <office 출처>; frame-ancestors 'none'` | — |
| 권한 | `clipboard-sanitized-write`만. 장치·화면 공유·알림 거절 | ✓ |
| 내려받기·이동·새 창 | `will-download` 막음, `will-navigate` 막음, `window.open` 거절 → 문서 링크는 지금의 `openLink`(묻고 기본 브라우저) | ✓ |
| 서버 | GET/HEAD만, Host 머리글 확인, 경로 탈출·점 이름 거절 (지금 `editor-server.mjs` 규칙). office 경로는 `/office/<판>-<실행마다 새 열쇠>/` — 다른 프로그램·웹 페이지가 미리 알고 글꼴 목록·글꼴 파일을 읽지 못하게 | — |
| 믿을 수 없는 문서 | x2t는 sandboxed 틀의 Worker에서만 돈다. 매크로·플러그인 없음 (`plugins.json` 비움) | — |

`'unsafe-eval'`은 ONLYOFFICE가 꼭 쓴다 (템플릿·`new Function`). 그래서 네트워크 없는 opaque origin 틀 안에 가둔다.

---

## 8. 라이선스 지키기

> 법률 자문이 아니라 파일을 읽고 정리한 것이다. **첫 배포 전에 법률 검토를 받는다.**

### 들어가는 것

| 부분 | 라이선스 | 근거 파일 |
|---|---|---|
| ONLYOFFICE web-apps·sdkjs 9.4.0.129 | AGPL-3.0 + 추가 조항 1~5. 코드 아닌 것(그림·아이콘·문서 글)은 CC BY-SA 4.0 | `editor/LICENSE.txt`(662줄부터), 모든 파일 머리글 |
| x2t (WASM) | ONLYOFFICE core AGPL-3.0 + CryptPad 패치 AGPL-3.0-or-later + BRF Tech 패치 | filex `NOTICE`, `upstream/onlyoffice.json` |
| filex-office-editor | AGPL-3.0-or-later. `src/locks.ts`만 AGPL-3.0-only (ONLYOFFICE 서버 코드를 옮김) | filex `LICENSE`, `NOTICE` |
| 편집기 안 제3자 JS | MIT·BSD·Apache-2.0 (Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO…) | `editor/3rd-Party.txt`, `editor/license/` |
| x2t·sdkjs WebAssembly 안 제3자 코드, `es6-promise`·`fetch` | BSL·ICU·FTL·MIT·BSD·IJG·Apache·MPL·GPL/LGPL 등 | `THIRD-PARTY-NATIVE.txt` (Office Kit이 더함, `scripts/office-editor/third-party.mjs`) |
| 편집용 글꼴 | **넣지 않음** (사용자 PC의 글꼴을 실행 중에 읽기만) | — |
| PDF 기본 14 글꼴 (URW++, x2t.wasm 안) | AGPL-3.0 + PDF·PostScript 포함 예외 — ONLYOFFICE core가 넣어 둔 것, 뺄 수 없다 | `THIRD-PARTY-NATIVE.txt` |
| Office Kit 호스트·앱 | MIT (MessagePort로 말하는 별개 프로그램) — 검토 대상 | — |

### ONLYOFFICE 추가 조항 (LICENSE.txt 그대로 요약)

1. 저작권·라이선스·보증 부인·출처 고지를 **모두 남긴다**.
2. 수정판은 **수정했다는 것과 날짜**, ONLYOFFICE(Ascensio System SIA) 원본을 바탕으로 했다는 것을 눈에 띄게 적는다.
3. 대화형 UI는 **쉽게 찾을 수 있고 눈에 띄는 UI**로 (i) 원 개발사가 ONLYOFFICE임, (ii) 수정판일 수 있음, (iii) 라이선스 정보를 보여 준다.
4. 상표·로고·브랜드 사용권은 **없다** (onlyoffice.com/trademark-policy).
5. 코드 아닌 것은 CC BY-SA 4.0.

### Office Kit이 할 일

| # | 할 일 | 어디서 |
|---|---|---|
| 1 | `LICENSE.txt`, `3rd-Party.txt`, `license/`, filex `LICENSE`·`NOTICE`·`CHANGES.txt`를 묶음에 그대로 둔다. 머리글을 지우지 않는다 (덜어 내기는 파일 단위만) | W1 |
| 2 | `vendor/office-editor/NOTICE-OfficeKit.md`: 날짜 붙은 수정 목록 (덜어 낸 것, filex 패치, 실행 중 AllFonts.js 덮어쓰기), "ONLYOFFICE Docs(Ascensio System SIA)를 바탕으로 함" | W1·W8 |
| 3 | 고지 UI: filex 앱의 아래쪽 법적 고지 줄을 **늘 보이게 둔다** + 편집기 자체의 '정보'를 둔다 + host 막대의 **`정보 · 라이선스`** → 한국어 대화상자 (원 개발사, 수정판일 수 있음, AGPL 전문·추가 조항, 설치된 판의 소스 주소) | W4·W8 |
| 4 | 소스 제공 (AGPL §6): 릴리스마다 `office-editor-source-<버전>.zip`(filex 소스 커밋 그대로 + 우리 패치 + `build-office-editor.mjs` + 덜어 내기 규칙 + 글꼴 생성기)을 Release에 붙인다. ONLYOFFICE sdkjs·web-apps·core·build_tools `v9.4.0.129`와 x2t 빌드 방법은 정확한 태그 주소 + **우리 GitHub에 거울(포크)** 을 둬서 배포하는 동안 계속 받을 수 있게 한다. 앱 안 링크는 **설치된 버전의 태그**를 가리킨다 | W8 (release.yml) |
| 5 | 상표: 'ONLYOFFICE' 이름·로고를 Office Kit 이름, 기능 이름, 창 제목, 설치 파일, 아이콘, 소개 글에 쓰지 않는다. "ONLYOFFICE Docs 기반"처럼 사실만 적는다. 편집기 안 기본 로고는 **그대로 둔다** (지우면 1번과 부딪히고, 바꾸면 상용판 기능인 화이트 라벨이 된다) | W4·W7·W8 |
| 6 | CC BY-SA 4.0: 아이콘·그림 출처를 '정보 · 라이선스'에 적는다. 고친 아이콘은 같은 조건으로 | W8 |
| 7 | 우리 파일 경계: 앱은 MIT 그대로. filex 패치는 `SPDX-License-Identifier: AGPL-3.0-or-later`, 글꼴 생성기(ONLYOFFICE AllFontsGen 형식을 옮긴 것)는 `AGPL-3.0-only`(+ ONLYOFFICE 추가 조항) 머리글. 전문은 저장소 `LICENSES/`. README 라이선스 절에 "vendor/office-editor와 AGPL 표시 파일은 AGPL" 추가 | W5·W8 |
| 8 | §13(네트워크 사용자)은 해당 없음 (이 PC 사용자만 127.0.0.x에 붙음). §6은 해당 (설치 파일을 배포하므로) | — |
| 9 | 배포 관문: beforePack이 라이선스 파일·`VERSION.json`·`NOTICE-OfficeKit.md`를 확인, release 작업이 소스 꾸러미가 붙었는지 확인 | W1·W8 |

'Word·Excel·PowerPoint'는 Microsoft 상표다. 형식 설명("Word 문서(.docx)")으로만 쓰고, 편집 창 이름은 '문서 편집·표 편집·슬라이드 편집'으로 한다.

---

## 9. 크기

| | 값 |
|---|---|
| filex `ui.zip` 전체 | 323.4MB, 2,633개 (zip 98.3MB) |
| 덜어 낸 묶음 | **213.7MB, 1,530개** (gzip-9 64.1MB) |
| 설치 파일 늘어남 | **+50~65MB** (NSIS LZMA는 gzip-9보다 작거나 같음) → 약 185~195MB |
| 메모리 | 편집 창마다 약 0.46~0.53GB (CryptPad 실험 측정, 같은 계열 엔진) |

줄이는 방법과 판단:

| 방법 | 효과 | 판단 |
|---|---|---|
| 위 덜어 내기 | −110MB | **함** |
| `.br` 사본만 두고 `Content-Encoding: br`로 내주기 | 설치 폴더 약 절반 | 나중에 (서버가 복잡해짐) |
| 슬라이드 테마 빼기 | −20MB | 안 함 ('디자인' 탭이 빔) |
| 편집기를 따로 내려받는 추가 구성 요소로 | 기본 설치 파일 그대로 | 안 함 (인터넷 없는 PC에서도 바로 되어야 함) |
| x2t 워커를 열기 뒤 끝내고 저장 때 다시 | 창마다 수십 MB | filex 앱 동작 확인 후 |

---

## 10. 시험 계획

| 층 | 무엇 | 언제 |
|---|---|---|
| 단위 (`npm test`) | `ooxml-fixups` (판정의 raw 출력 3종을 `test/fixtures/office-editor/`에), 형식 확인, 서버 경로·CSP·금지 포트, `allfonts` (기준: 묶음 core-fonts → 묶음 `AllFonts.js`와 같음), 호스트 약속 (가짜 MessagePort) | 늘 |
| 통합 (가짜 electron) | `test/editor-integration.test.mjs` 확장: doc:edit → office kind, 저장 IPC, 닫기 흐름, 자동 저장 → 사본, AI 보내기 막기 | 늘 |
| e2e (Electron + CDP) | 형식 3 × 만든 방법 3 (`createDocument`, `createBlankDocument`, Office로 만든 고정 파일): 열기 → 입력(키·CDP 조합 한국어) → Ctrl+S → `readDocument(engine:'file')`로 확인. 네트워크 기록: 밖 0건. WebRTC 시험: 후보 0 | `OFFICEKIT_EDITOR_SMOKE=1` |
| Office 검증 (COM) | 저장한 파일을 경고 모두 켜고 엄격히 열기 (`strict-open.ps1`, 감시 타이머가 그 Office만 끝냄). docx 표 `Rows.Item(n)`, PowerPoint 열기. 원본에서 옮겨 오기(테마 글꼴·한국어 줄 바꿈·호환 설정·셀 글꼴·메모·조건부 서식·구역)를 거친 파일도 (`test/fixtures/office-editor/carry/`) | `OFFICEKIT_COM_TESTS=1` |
| 충실도 (COM, 사람이 돌림) | Office로 만든 한국어 문서(형식마다 하나: 표·각주·변경 추적·메모·차트·SmartArt·내용 컨트롤 / 서식·조건부 서식·메모·차트·그림 / 구역·숨김·전환·애니메이션)를 편집 창에서 고쳐 저장 → 저장 전후를 COM으로 비교(글꼴 이름·행 높이·개체 크기·구역 수·쪽 수)하고 `src/core/render.mjs pageImage`로 쪽 그림을 나란히 본다 | 릴리스 전, x2t·편집기 판을 올릴 때 |
| 사람 (출시 관문) | 실제 MS IME 두벌식·세벌식·한자 변환, Word·Excel·한글에서 복사 붙이기, 인쇄, 50쪽 문서, 다크 모드, 고해상도, 창 4개 | 릴리스 전 체크리스트 |

- 실제 MS IME는 **사람이 손으로** 확인한다 (`docs/QA-OFFICE-EDITOR.md` 2.1). OS 키 입력 자동화(`ime-send.ps1`)는 저장소에 두지 않는다 — 다른 창으로 키가 갈 위험이 있어서.
- CDP의 IME 흉내는 엔진에 따라 글자가 겹친 적이 있다 (CryptPad 빌드). 그래서 IME 판정은 사람 시험으로 한다.

---

## 11. 알려진 한계

- **완전히 같지 않음**: x2t가 꾸러미를 다시 쓴다. Office로 만든 한국어 문서 5개(보고서·회의록 docx, 예산·매출 xlsx, 발표 pptx)를 COM으로 저장 전후 비교한 결과:
  - 지켜짐: 변경 추적·메모·각주·차트·SmartArt 글·필드·그룹·하이퍼링크·머리글/바닥글, 셀 값·수식·병합·조건부 서식·유효성·이름·표(ListObject)·스파크라인·자동 필터·인쇄 설정, 마스터·레이아웃·숨김·전환·애니메이션·노트.
  - x2t가 잃어 저장 뒤 고치기가 **덮어쓸 원래 파일에서 옮겨 오는 것**(6장 고치기, `src/core/ooxml-fixups.mjs`): 테마 글꼴 묶음의 스크립트별 글꼴(Hang=맑은 고딕 — 없으면 한글이 바탕으로 보임), `w:themeFontLang`의 eastAsia(ko-KR → zh-CN으로 바뀜), 기본 문단 스타일의 `wordWrap`·`autoSpaceDE`·`autoSpaceDN`(없으면 "3 분기"처럼 띄고 단어 단위로 줄 바꿈), `w:compat` 항목(표 줄 높이 등), 내용 컨트롤 종류(일반 텍스트 → 서식 있는 텍스트가 됨), xlsx 셀 글꼴 이름(맑은 고딕 → Calibri — 행 높이 17.4 → 13.8pt, 차트·그림 찌그러짐), 원래 메모(스레드 메모로 바뀜)·기본 조건부 서식(x14 확장으로만 옮겨짐 — Excel 없는 파일 엔진이 거절), pptx 구역. 편집기에 이것을 바꾸는 화면이 없어서 사용자 편집을 되돌리지 않는다. 새 문서(원래 파일 없음)는 `themeFontLang`만 문서 기본 언어로.
  - 남는 차이 (Office·Office Kit 동작에는 지장 없음): docx `styles.xml`이 49KB → 640KB(용어집 30KB → 637KB)로 커지고 스타일 ID가 숫자로 바뀜(Normal → 773 — styleId로 찾는 바깥 도구는 깨짐, `word_edit`는 이름으로 찾아 괜찮음), docDefaults의 `w:kern`이 빠짐, `docProps/app.xml` Application이 `ONLYOFFICE/9.4.0.129`. xlsx `calcChain.xml`·`printerSettings1.bin`이 빠짐(Excel이 다시 만들고 인쇄 설정은 남음), 시트 안 링크 `'코드표'!A1`이 `코드표!A1`로(동작 같음). pptx `docProps/thumbnail.jpeg`이 빠져 탐색기 미리보기 없음, 6~8쪽에 빈 노트 슬라이드가 생김.
  - 저장한 5개 모두 엄격 열기(경고·복구 창 없음) 통과, `word_edit`·`ppt_edit` 동작 같음.
- 표에 새 행 → `vMerge` (6장 고치기로 막음).
- 암호 문서, 매크로 문서(docm·xlsm·pptm), 옛 형식(doc·xls·ppt·odt·ods·odp)은 1차에서 뺀다. 옛 형식은 2차: 열어서 OOXML로 '다른 이름으로 저장'만.
- 함께 편집 없음, 맞춤법 검사 없음 (한국어 사전 없음, 꺼 둠).
- 편집기 UI는 ONLYOFFICE 모양·로고 그대로 (라이선스 조건). 새 기능 풍선·셀 스타일 이름 일부는 영어. 아래쪽 법적 고지 줄은 filex 문구(영어) — 한국어 문구는 W2 패치로.
- 상태 막대의 '모든 변경 사항이 저장되었습니다'는 편집기 안 기록 이야기이지 파일 저장이 아니다 (호스트 막대가 진짜 상태를 보여 줌).
- 메모리: 창마다 약 0.5GB.
- 실제 MS IME는 이번 판정에서 확인하지 못했다 (사용자 Chrome이 앞으로 와 시험이 스스로 멈춤). Euro-Office 9.3 엔진에서는 CryptPad 실험이 확인했다.

---

## 12. 위험과 대비

| 위험 | 크기 | 대비 |
|---|---|---|
| filex가 나온 지 열흘 된 한 회사 프로젝트 (0.1.0 2026-10-09, 0.3.1 2026-10-10) | 중 | 태그·커밋·SHA-256 고정. 우리 쪽 약속은 메서드 10개뿐. AGPL이라 포크해 이어 갈 수 있음 (다리 약 1.5천 줄 + 시험 280개). 우리 패치는 PR로 보냄 |
| ONLYOFFICE 올릴 때 깨짐 | 중 | filex가 매주 새 판을 확인. 우리는 올릴 때마다 e2e + COM 엄격 열기 전체를 통과해야 고정 값을 바꿈 |
| x2t ↔ Office 호환 문제가 더 있음 | 중 | 고치기 단계를 한 모듈(`ooxml-fixups`)에 모으고, 실제 한국어 문서 모음으로 COM 회귀 시험 |
| 라이선스 해석 (MIT 앱 + AGPL 편집기) | 중 | 법률 검토, 경계 문서화, 소스 꾸러미 자동 첨부 |
| 글꼴 형식(`AllFonts.js`·선택 이진)은 sdkjs 내부 형식 | 중 | 기준 비교 시험(5장). 실패하면 묶음 글꼴 + 한글 대체만으로 물러남 |
| 메모리·설치 크기 | 낮음 | 9장 |
| **대안 (plan B)** | — | CryptPad 빌드 + 우리 oo.js (출처 위험을 받아들일 때), 또는 onlyoffice-web 컴포넌트 |

---

## 13. 만들기 계획

### 일감 (파일 주인이 겹치지 않게)

| ID | 일 | 주인 파일 | 기다리는 것 | 크기 |
|---|---|---|---|---|
| W1 | 묶음 만들기·확인·덜어 내기·CI | `scripts/build-office-editor.mjs`, `scripts/office-editor/*`(규칙·lock 확인), `scripts/check-editor-bundle.cjs`, `package.json`(scripts·build.files), `.github/workflows/release.yml`(묶음·캐시 단계) | 없음 | 2~3일 |
| W2 | filex 패치 2개 + PR: `file.save`에 `reason`, 앱 문구 한국어(법적 고지 줄 포함) | `patches/office-editor/filex/*.patch` | 없음 (W1 5단계로 들어감) | 1~2일 |
| W3 | 서버·창 일반화: office 출처(127.0.0.3), CSP, 글꼴 경로, 금지 포트 피하기(한글 편집기도), kind별 확장자·제목·형식 확인·고치기 훅·새 IPC 3개 | `src/app/editor-server.mjs`, `src/app/editor-window.mjs`, `test/editor-server.test.mjs` | W4와 약속(위 IPC 표)만 먼저 정함 | 4~5일 |
| W4 | Office 호스트 페이지: filex SDK 응답, hid 저장, 막대(저장·정보 · 라이선스), 바깥 변경 띠, 복구 사본 띠 | `src/app/office-editor/host.html·host.js·host.css`, `src/app/office-editor/office-preload.cjs` | 약속 | 3~4일 |
| W5 | 글꼴: 목록·미리보기 생성기, 캐시, XOR 내주기용 색인 | `src/app/office-editor/fonts/*`, `test/office-fonts.test.mjs` | 없음 (기준 시험은 W1 묶음) | 3일 |
| W6 | 저장 뒤 고치기 3종 + 시험 + COM 엄격 열기 시험 | `src/core/ooxml-fixups.mjs`, `test/ooxml-fixups.test.mjs`, `test/fixtures/office-editor/*`, `scripts/dev/strict-open.ps1` | 없음 | 2~3일 |
| W7 | 앱 잇기: doc:edit·app:init(kind별 사용 가능)·doc:new(edit)·AI 막기·자동 저장 사본 정리, 화면 문구 | `src/app/main.mjs`, `src/app/renderer/*`, `src/app/store.mjs`(officeEditor 설정) | W3·W4 약속 | 2~3일 |
| W8 | 라이선스: NOTICE·정보 대화상자 글, 소스 꾸러미 만들기·첨부, 거울 저장소, README·ARCHITECTURE·FEATURES 갱신, 법률 검토 요청 | `docs/*`, `README.md`, `scripts/make-office-source.mjs`, `release.yml`(소스 단계 — W1과 순서 맞춤) | 없음 | 2일 |
| W9 | e2e·검증: CDP e2e, 네트워크·WebRTC 탐지, COM 검증 묶음, 사람 체크리스트(실제 IME는 손으로) | `test/office-editor-e2e.test.mjs`, `scripts/dev/office-e2e.mjs`, `docs/QA-OFFICE-EDITOR.md` | W1·W3·W4 | 3~4일 |

### 순서

```
1주   W1 ─┬─ W2            W5        W6        W8(초안)
          └─ W3 뼈대 ── W4 뼈대   (약속 먼저 고정: 6장 IPC·SDK 표)
2주   W3 마무리 · W4 마무리 · W7 · W5 마무리
3주   W9 (e2e·COM·사람 시험) · W8 마무리 · 법률 검토 · 릴리스 후보
```

### 끝났다고 보는 기준

- [ ] 형식 3 × 만든 방법 3 문서가 열기·고치기·저장되고 Word·Excel·PowerPoint가 경고 없이 연다
- [ ] Office Kit '새 문서' pptx를 고쳐 저장해도 PowerPoint 복구 창이 안 뜬다
- [ ] 편집 중 밖으로 나간 요청 0건, WebRTC 후보 0개
- [ ] 한글이 맑은 고딕 등 실제 글꼴로 보이고 글꼴 목록에 한국어 이름·미리보기가 있다
- [ ] 실제 MS IME 사람 시험 통과
- [ ] 자동 저장이 원본을 덮어쓰지 않는다
- [ ] 라이선스 파일·수정 고지·`정보 · 라이선스` 화면·소스 꾸러미가 있다
- [ ] 설치 파일 200MB 안

---

## 14. 근거 파일

| 무엇 | 경로 (git 밖, `tmp/`) |
|---|---|
| 판정 시제품·스크립트 | `spike-judge/proto/`, `spike-judge/scripts/judge.mjs`, `rtc-probe.mjs`, `size.mjs`, `realime.mjs` |
| 판정 결과 | `spike-judge/out/j1-report.json`, `spike-judge/out/j1/*.strict.json`, `spike-judge/shots/j1-*.png` |
| filex 실험 | `spike-filex/proto/`, `spike-filex/scripts/make-allfonts.mjs`, `spike-filex/repo/` (소스·`dist/ui`), `spike-filex/shots/explore-docx-1.png`(□ 글꼴) |
| CryptPad 실험 | `spike-cryptpad/tools/gen-allfonts.mjs`·`gen-thumbs.mjs`, `spike-cryptpad/scripts/strict-open.ps1`·`ime-send.ps1`, `spike-cryptpad/out/final/`, `spike-cryptpad/dl/licenses/` |
| onlyoffice-web 실험 | `spike-oow/electron/ooxml-fixups.cjs`, `spike-oow/scripts/pull-layers.sh`, `spike-oow/legal/NOTICE.md`, `spike-oow/logs/e2e-run3.json` |
