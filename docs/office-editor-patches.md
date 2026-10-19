# Office 편집기 — filex 패치

> 상태: 패치 3개 완성·검증 (2026-10-18). 업스트림 PR은 **초안만** 있다 — 보내지 않았다.
> 바탕: filex-office-editor `v0.3.1` · `57dda3dae99af78eb26b3c64b69f1e3f711f0757` (태그 커밋). 설계: [`OFFICE-EDITOR-DESIGN.md`](OFFICE-EDITOR-DESIGN.md) 6장(자동 저장)·8장(라이선스)·13장 W2.

## 한눈에

| 패치 | 하는 일 | 꼭 필요? |
|---|---|---|
| `0001-app-file-save-says-why.patch` | `file.save`에 `reason`을 더한다: 사용자 저장 `"user"`, 10분 자동 저장 `"auto"`. 자동 저장 중 누른 저장은 그 뒤에 따로 한다 | **예** — 없으면 자동 저장이 사용자 파일을 몰래 덮어쓴다 |
| `0002-app-korean-strings.patch` | 앱 페이지 문구 한국어 (아래쪽 **법적 고지 줄** 포함). `ko` 지역 → `ko-KR` | 예 — 한국어 UI |
| `0003-build-source-for-modified-builds.patch` | `build-app.mjs --source <주소>`: 수정판의 소스 위치를 법적 고지 줄에 넣는다 | 권장 — 안 주면 고지 줄이 filex 원본 태그를 가리킨다 |
| `0004-app-no-feature-tips.patch` | 편집기 설정 `customization.features.featuresTips: false` — '신규' 기능 풍선(새 버튼 안내)을 끈다 | 권장 — 샌드박스 틀 안 편집기는 풍선을 닫은 것을 기억하지 못해 열 때마다 도구 모음을 가린다 |

- 모두 **차례대로** 얹는다 (0001 → 0002 → 0003 → 0004). 0002·0003·0004는 앞 패치 위에서 만들었다.
- 패치를 얹어 다시 빌드한 묶음은 릴리스 `ui.zip`과 **filex가 빌드하는 세 파일만** 다르다 (2,633개 중): `filex/app.js`(패치 내용 + 머리글), 다리 `editor/web-apps/vendor/socketio/socket.io.min.js`·`filex/x2t-worker.js`(코드는 같고 맨 위 머리글만 `MODIFIED build … Source of this build: <주소>` — 0003). 편집기 파일·x2t는 같은 바이트.
- 검증: filex 단위 시험 280 → 293개 통과, `tsc` 통과. Electron 하네스에서 docx·xlsx·pptx 각 10개 확인 모두 통과 (아래 '검증').

---

## 1. 얹는 방법 (W1 5단계)

```bash
# filex 소스: 고정 커밋 (전체 1.6MB, 92파일 — 성긴 받기를 쓰면 패치가 고치는 tests/도 받을 것)
git clone --depth 1 --branch v0.3.1 https://github.com/BRF-Tech/filex-office-editor filex-src
git -C filex-src rev-parse HEAD        # 57dda3dae99af78eb26b3c64b69f1e3f711f0757 이어야 함
for p in patches/office-editor/filex/*.patch; do
  git -C filex-src apply --check "$p" && git -C filex-src apply "$p"
done
cd filex-src && npm ci --ignore-scripts
node scripts/build-app.mjs --editor <lock 그대로의 editor> --x2t <x2t> --out <폴더> --no-zip \
  --source github.com/toxh13/office-kit/releases/tag/v<Office Kit 버전>
```

주의:

| 할 것 | 까닭 |
|---|---|
| `git -C <filex 폴더> apply` (그 폴더가 git 저장소여야 함) | 다른 저장소의 하위 폴더에서 `git apply`를 돌리면 그 폴더 밖 경로를 **조용히 건너뛴다** (`0 files changed`) |
| 패치를 **LF**로 넘긴다 | CRLF로 바뀐 패치는 `patch does not apply`로 실패한다 (확인함). Office Kit 저장소엔 `.gitattributes`가 없고 이 PC는 `core.autocrlf=true` → Windows CI 체크아웃에서 CRLF가 될 수 있다. W1 `applyPatches`는 실패하면 CR을 지우고 다시 해 본다. `patches/** -text` 속성을 두면 더 좋다 |
| 성긴 받기에 `tests/` 포함 | 0001·0002·0003이 `tests/*.test.ts`를 고친다. 맨 위 파일(`README.md`, `CHANGELOG.md`)도 고친다 (cone 모드는 늘 받음) |
| `--source`는 0003을 얹었을 때만 | 0003 없는 `build-app.mjs`는 `unknown argument --source`로 멈춘다. 주소는 ASCII 한 줄, 공백·따옴표·`<>`·`\`·`` ` `` 없이 4~200자 |
| 패치 머리말은 지우지 않는다 | 첫 줄 `SPDX-License-Identifier: AGPL-3.0-or-later`, `Subject:`, `Date:` (W1이 `NOTICE-OfficeKit.md`에 날짜 붙은 수정 목록으로 옮김). `git apply`는 첫 `diff --git` 앞 글을 건너뛴다 |

시험까지 돌리려면 (선택): `npx tsc -p tsconfig.json && npx vitest run` (x2t-wasm 시험은 x2t가 없으면 건너뜀).

---

## 2. 약속 — `file.save`의 `reason` (호스트가 따를 것)

### 앱 → 호스트

```js
// MessagePort, filex 앱 SDK v1
{ id, method: 'file.save', params: { index: 0, data: ArrayBuffer /* transfer */, mime, reason: 'user' | 'auto' } }
// 함께 편집(coedit, Office Kit은 쓰지 않음)일 때만 through: number 가 더 붙는다
```

| `reason` | 언제 |
|---|---|
| `'user'` | 호스트의 `{hid, request:'save'}` (Office Kit: main의 Ctrl+S, 막대 '저장', 닫기 대화상자 '저장') · 편집기 안 저장 단추·Ctrl+S (main이 가로채지 않았을 때) · 앱 페이지 위 Ctrl+S (SDK) · 휴대폰 보기 전환 · 함께 편집의 마지막 사람이 나갈 때 |
| `'auto'` | 고친 채 10분 (`AUTOSAVE_MS`, 30초마다 확인). 함께 편집이면 저장 담당자의 10분 |

- 패치 없는 빌드는 `reason`이 없다 → 호스트는 '부탁한 저장 중이면 `user`, 아니면 `auto`'로 본다 (W4 `host.js`가 이미 그렇게 함). 그 밖의 값은 오지 않는다 (`saveParams`가 거절).
- 답: 두 경우 모두 성공이면 `{ saved: true, size }`. 실패(`error`)면 앱이 "문서를 저장하지 못했습니다: …"를 띄우고 편집기는 '저장 안 됨' 그대로.

### 호스트가 지킬 것

1. **`'auto'`는 사용자 파일에 쓰지 않는다** → `userData/office-drafts/<경로 sha1>.<ext>` (설계 6장).
2. **`'auto'` 뒤에도 '저장 안 됨'을 유지한다.** 앱은 어떤 저장이든 성공하면 편집기에 `saved`를 알리고 `ui.dirty(false)`를 보낸다. 호스트는 `autoPending`(자동 저장 사본만 있음)을 따로 들고, 보여 줄 dirty = `appDirty || autoPending`. `'user'` 저장이 성공하면 `autoPending = false`.
3. `{hid, request:'save'}`의 답 `{hid, result:{saved:true}}`는 그 `'user'` `file.save`에 호스트가 답한 **뒤에** 온다.
4. 자동 저장이 도는 중에 `{hid, request:'save'}`를 보내면: 앱이 자동 저장을 끝낸 뒤 `'user'` 저장을 **따로 한 번** 하고, 그다음에 hid 답을 보낸다 (전에는 자동 저장의 답으로 끝나 버렸다). 그사이 여러 번 눌러도 사용자 저장은 한 번.

---

## 3. 패치별 설명

### 0001 — 모든 `file.save`가 까닭을 말한다

| 파일 | 바뀐 것 |
|---|---|
| `src/app/config.ts` | `SaveReason = 'user' \| 'auto'`, `SAVE_REASONS`, `saveParams(data, mime, reason, through?)` (다른 값은 throw) |
| `src/app/save-queue.ts` (새 파일) | `SaveQueue`: 저장은 한 번에 하나. 같은 종류나 '사용자 중 자동'은 진행 중인 것에 합침, '자동 중 사용자'는 끝난 뒤 한 번 더 |
| `src/app/main.ts` | `save(why)`; 10분 `tick()` → `'auto'`, 편집기 저장·`fx.onSave`·보기 전환·마지막 사람 → `'user'`. 혼자일 때 SDK 0.53의 `fx.save()`에는 `reason` 칸이 없어 `fx.request('file.save', saveParams(...), [bytes])` |
| `src/app/coedit-client.ts`, `src/app/together.ts` | 함께 편집 저장에도 `reason` (`through`와 함께) |
| `tests/save-reason.test.ts` (새), `tests/together.test.ts` | 시험 9개 + 1개 고침 |
| `README.md`, `CHANGELOG.md` | 설명, `[Unreleased]` |

filex에서 달라지는 것: `file.save`에 칸 하나가 더 붙는다 (filex는 둘 다 새 버전으로 씀 — `through`가 붙던 것과 같은 방식). 자동 저장 중 누른 저장이 버전 하나를 더 만든다 (자동 저장의 스냅숏이 마지막 입력보다 오래됐을 수 있으므로 맞는 동작).

### 0002 — 앱 페이지 한국어

| 파일 | 바뀐 것 |
|---|---|
| `src/app/config.ts` | `UiLang`에 `'ko'`, `uiLang('ko'·'ko-KR'·'KO_kr')` → `'ko'` (`'kok'`은 아님). `editorRegion('ko')` → `'ko-KR'` (전에는 `'ko-KO'`) |
| `src/app/strings.ts` | `ko` 문구 전부 + 새 키 `editorReloaded` (en·tr·ko). 전에는 `main.ts`에 영어 그대로 박혀 있었다 |
| `src/app/main.ts` | `openFailed(t.editorReloaded)` |
| `tests/app-config.test.ts` | 언어마다 키가 같은지, 한국어 법적 고지 줄 필수 내용, 저장 거절 문구 |
| `README.md`, `CHANGELOG.md` | 설명 |

`filex-app.json`(filex 앱 목록용 설명)은 영어·터키어 그대로 둔다 — Office Kit은 쓰지 않는다.

**법적 고지 줄** (편집기 아래, 늘 보임, 마우스를 올리면 줄바꿈해 전부 보임). 뜻과 필수 내용은 영어 원문과 같다:

| 필수 내용 (ONLYOFFICE 추가 조항 2·3, AGPL 6조) | 영어 원문 | 한국어 |
|---|---|---|
| 원 개발사 | Based on ONLYOFFICE Docs by Ascensio System SIA | Ascensio System SIA의 ONLYOFFICE Docs를 바탕으로 하며 |
| 수정판일 수 있음 | this version may have been modified | 이 버전은 수정되었을 수 있습니다 |
| 판·빌드·소스 태그 | (Docs 9.4.0, build 129, ONLYOFFICE source tag v9.4.0.129) | (Docs 9.4.0, 빌드 129, ONLYOFFICE 소스 태그 v9.4.0.129) |
| 라이선스 | Free software under the GNU AGPL version 3 | GNU AGPL 버전 3에 따른 자유 소프트웨어입니다 |
| 소스 위치 | source code: github.com/BRF-Tech/filex-office-editor/tree/v0.3.1 | 소스 코드: (같은 주소, 0003 `--source`면 그 주소) |
| 상표 | ONLYOFFICE® is a registered trademark of Ascensio System SIA; this app is not affiliated with or endorsed by it. | ONLYOFFICE®는 Ascensio System SIA의 등록 상표이며, 이 앱은 Ascensio System SIA와 제휴 관계가 없고 그 보증을 받지 않았습니다. |
| 접근성 이름 (`aria-label`) | About this editor | 이 편집기 정보 |

- 'ONLYOFFICE'는 사실을 적는 곳(바탕·상표 고지)에만 있다. Office Kit 이름·창 제목·메뉴·기능 이름에는 없다 (설계 8장 5번).
- 낱말: vault → '금고', version(문서) → '편집본', 보기(viewer) → '보기 화면'. 문체는 Office Kit 대화상자와 같은 '-습니다'.

### 0003 — 수정판의 소스 위치 (`--source`)

| 파일 | 바뀐 것 |
|---|---|
| `scripts/build-app.mjs` | `--source <주소>` → `checkSource()`(한 줄, 공백·따옴표·`<>`·`\`·`` ` `` 없음, 4~200자) → esbuild define `__APP_SOURCE__`. 없으면 `''`. 주면 만든 스크립트(`filex/app.js` 등) 맨 위 머리글도 `MODIFIED build of filex-office-editor … Source of this build: <주소>`로 (수정판 표시) |
| `src/app/strings.ts` | `LegalFacts.source?`, `sourceOf(app, source?)` — 주면 그 주소, 아니면 예전대로 filex 태그 |
| `src/app/main.ts` | `setLegal`이 `__APP_SOURCE__`를 넘김 |
| `tests/app-config.test.ts`, `tests/app-bundle.test.ts` | 세 언어에서 주소가 바뀌고 나머지 고지는 그대로인지, define이 들어가는지, 나쁜 값 거절 |
| `README.md`, `CHANGELOG.md` | 설명 |

까닭: 패치를 얹은 앱의 고지 줄이 `…/filex-office-editor/tree/v0.3.1`(패치 없는 원본)을 가리키면 '지금 돌리는 것의 소스'가 아니다. Office Kit 빌드는 `github.com/toxh13/office-kit/releases/tag/v<버전>`(소스 꾸러미 `office-editor-source-<버전>.zip`이 붙는 곳, 설계 8장 4번)을 넘긴다 — W1 `build-office-editor.mjs`가 이미 그렇게 함. `--source` 없이 빌드하면 바이트가 예전과 같다.

### 0004 — 새 기능 풍선 끄기 (2026-10-19, v0.1.4)

| 파일 | 바뀐 것 |
|---|---|
| `src/app/config.ts` | `editorConfig().editorConfig.customization.features`에 `featuresTips: false` (ONLYOFFICE 9.4 설정 — `api.js`의 `features.featuresTips`) |
| `tests/app-config.test.ts` | 그 값이 들어가는지 |

까닭: 편집기는 새 기능마다 '신규 새 버튼: 서식' 같은 풍선을 처음 한 번 띄우고 닫았다는 것을 브라우저 저장소에 적는데, filex의 편집기 틀은 `sandbox`(opaque origin)라 저장소가 없어 열 때마다 다시 뜬다 (v0.1.3 최종 평가). 편집기 파일은 고치지 않는다 — 편집기가 받는 설정만 바꾼다.

---

## 4. 검증 (2026-10-18, Windows 11, Node 24.18.0)

| 확인 | 결과 |
|---|---|
| 깨끗한 클론(acc87f8)에 `git apply --check` → `git apply` 3개 | 모두 됨. 결과 트리 = 작업 트리 (`diff -r` 같음) |
| CRLF로 바꾼 패치 | `patch does not apply` → CR 지우면 됨 |
| filex `tsc -p tsconfig.json` | 오류 0 |
| filex `vitest run` | 280 → **293 통과**, 17 건너뜀 (x2t-wasm 등, 전과 같음) |
| 패치 없는 소스로 `bundleScripts` | `filex/app.js`·`x2t-worker.js`·`socket.io.min.js` 모두 릴리스와 **같은 바이트** (빌드 환경이 릴리스와 같다는 확인) |
| 패치 얹고 `build-app.mjs --no-zip --source …` | 2,633파일 323.4MiB, 27초. 릴리스 `ui/`와 다른 파일은 `filex/app.js` 하나 |
| Office Kit 한글 편집기 시험 | `editor-integration` 17 통과·0 실패, `editor-server` 25/25 |

### Electron 하네스 (`tmp/W2/harness`, `tmp/spike-judge/proto` 바탕)

- host가 filex SDK를 흉내 내고 `params.reason`을 그대로 main에 넘긴다. main은 `'user'` → 문서 파일, `'auto'` → `drafts/` 사본만 (Office Kit 정책 흉내). Ctrl+S는 main `before-input-event`가 가로채 `{hid, request:'save'}` (Office Kit 경로).
- 자동 저장 부르기: 앱 틀의 `Date.now`를 CDP로 11분 앞당김 → 30초 안의 `tick()`이 `'auto'` 저장. 입력은 모두 CDP (`Input.insertText`·`dispatchKeyEvent`). OS 키 입력 없음.
- 문서: `tmp/spike-judge/docs/judge.{docx,xlsx,pptx}`, locale `ko-KR`.

| # | 확인 | docx | xlsx | pptx |
|---|---|---|---|---|
| 1 | 아래쪽 고지 줄이 한국어 (필수 내용·`aria-label` '이 편집기 정보') | ✓ | ✓ | ✓ |
| 2 | 자동 저장 → `reason:'auto'`, params 키 `data,index,mime,reason` | ✓ | ✓ | ✓ |
| 3 | 자동 저장은 원본 파일을 안 만들고 사본만 (사본에 입력한 글) | ✓ | ✓ | ✓ |
| 4 | 자동 저장 뒤 앱은 `ui.dirty(false)`, 호스트는 '저장 안 됨' 유지 | ✓ | ✓ | ✓ |
| 5 | Ctrl+S (main → host hid) → `reason:'user'`, hid 답 `{saved:true}` | ✓ | ✓ | ✓ |
| 6 | 사용자 저장 파일에 두 번의 입력이 다 있음 | ✓ | ✓ | ✓ |
| 7 | 사용자 저장 뒤 '저장됨' | ✓ | ✓ | ✓ |
| 8 | 가로채지 않은 Ctrl+S (편집기 자체 저장 → 다리) → `'user'` | ✓ | ✓ | ✓ |
| 9 | 자동 저장 도중(호스트가 5초 붙잡음) Ctrl+S → `'auto'` 다음 `'user'`, hid 답은 사용자 저장 뒤 | ✓ | ✓ | ✓ |
| 10 | 밖으로 나간 요청 0건 (265~278개 모두 로컬) | ✓ | ✓ | ✓ |

증거: `tmp/W2/out/{docx,xlsx,pptx}-report.json`, `tmp/W2/shots/*-2-autosaved.png`·`*-4-final.png`, `tmp/W2/shots/legal-ko-KR-bottom.png`(고지 줄 전체, `--source` 빌드), `legal-en-US.png`(영어 그대로). 다시 돌리기: `electron tmp/W2/harness/main.cjs --control 47367` 뒤 `node tmp/W2/work/verify.mjs <문서> <tag>`.

---

## 5. 라이선스

- 패치는 AGPL 프로그램(filex-office-editor, AGPL-3.0-or-later)을 고치는 것 → 각 `.patch` 첫 줄 `SPDX-License-Identifier: AGPL-3.0-or-later`. 새 파일도 SPDX 머리글: `save-queue.ts`는 "Contributed by Office Kit"(BRF Tech 저작권 줄은 붙이지 않음), `save-reason.test.ts`는 filex 시험 파일 머리글 그대로.
- AGPL 5조 a(수정·날짜 고지): 패치 머리말의 `Date:`·`Subject:` + `vendor/office-editor/NOTICE-OfficeKit.md`(W1이 패치 목록을 날짜와 함께 씀) + 소스 꾸러미에 패치 그대로 (W8).
- 업스트림의 영어·터키어 문구는 고치지 않았다. 영어 한 줄(`the editor page reloaded`)은 같은 글 그대로 문자열 표로 옮겼고, 터키어는 그 새 키 하나만 더했다.

---

## 6. filex를 올릴 때

1. 새 태그를 받아 `git apply --3way`로 차례대로 얹는다. 걸리면 고쳐서 `git diff --full-index`로 다시 만든다 (머리말은 이 문서 형식 그대로).
2. `tsc` + `vitest run` + 하네스 `verify.mjs` 세 형식.
3. 업스트림이 받아 준 패치는 지운다 (0001·0003은 그대로 PR 초안).

---

## 7. 업스트림 PR 초안 (보내지 않음)

> 사용자 허락 없이 PR·이슈·포크를 만들지 않는다. 보낼 때는 각 패치를 `main`에 다시 얹어(0002·0003은 `CHANGELOG.md`가 겹침) 따로 보낸다.

### PR 1 — `app: every file.save says why (reason "user" | "auto")`

```markdown
## What

Every `file.save` the app makes now carries `reason`:

- `"user"` — a save the person asked for: the editor's Save or Ctrl+S, filex's
  Save (the host's `{hid, request: "save"}`), Ctrl+S on the app page, a switch
  between the phone app and the editor, the last writer leaving a session;
- `"auto"` — the ten-minute save while there are changes (alone, or the
  saver's in a session).

filex writes both as a new version, exactly as today, and can ignore the field
(it is passed like `through` already is). A host that keeps an autosave apart
from the file — a recovery copy, rather than writing over a file the person did
not ask to save — can now tell the two apart. We embed the app in a desktop
host (Electron, no filex server) that does exactly that.

Alone, the save goes through the SDK's raw `request("file.save", …)` because
`@brftech/filex-app-ui` 0.53's `save()` has no field for it; together,
`coedit-client` passes it beside `through`. If filex wants it in the SDK, a
`reason?: "user" | "auto"` on `SaveParams` and on `save()`'s options would do.

## Also fixed

A Save pressed while the ten-minute save was running was answered by that
autosave — whose snapshot could be older than the person's last change. Saves
are still one at a time (`src/app/save-queue.ts`), but the person's save asked
for during an autosave is now made after it (once, however often it was asked
for); an autosave due during the person's save is that save.

## Tests

`tests/save-reason.test.ts` (params, the coedit save, the queue: same kind,
auto during user, user during auto, a failed autosave, a failed user save);
`tests/together.test.ts` updated. `npm test`: 289 passed. Measured in an
Electron host with docx, xlsx and pptx: the autosave arrives with `"auto"`,
Ctrl+S and the editor's Save with `"user"`, and a Ctrl+S during a held autosave
gives `"auto"` then `"user"`.
```

### PR 2 — `app: the app page in Korean`

```markdown
## What

The app page's own words in Korean for a `ko` locale, beside English and
Turkish: the notice under the editor (based on ONLYOFFICE Docs by Ascensio
System SIA, may have been modified, Docs version / build / source tag, GNU AGPL
version 3, where the source is, the trademark line), the opening, failure,
save-refusal and editing-together messages.

- `uiLang`: `ko`, `ko-KR`, `ko_KR` → `ko`.
- `editorRegion("ko")` → `ko-KR` (it was `ko-KO`, not a region); English keeps
  `en-US`, everything else as before.
- "the editor page reloaded", an English literal in `main.ts`, is now a string
  in all three languages (`editorReloaded`; the English words are unchanged).
- The manifest (`filex-app.json`) is left in English and Turkish.

## Tests

`tests/app-config.test.ts`: every language has every key the English has (and
its own words), the Korean legal line holds every required part with the
version, build and tag as given, the save refusals say what happened and what
to do. `npm test` passes. Seen in an Electron host with `locale: "ko-KR"`.
```

### PR 3 — `build: --source <address> for builds from changed sources`

```markdown
## What

`node scripts/build-app.mjs --source <address>` for a build made from changed
sources — a fork, or a redistribution that applies patches of its own. The
notice under the editor then names that address as where the source is,
instead of this project's tag, so the person is pointed at the source of what
they actually run (AGPL section 6). Without `--source` the bundle is byte for
byte what it was (`__APP_SOURCE__` is defined as `""`).

The address is checked: one line, no spaces, no quotes, `<`, `>`, backslash or
backtick, 4 to 200 characters.

## Tests

`tests/app-config.test.ts` (all three languages name the given address and keep
the rest of the notice; none given → this project's tag) and
`tests/app-bundle.test.ts` (`checkSource`, and the address reaching
`filex/app.js` through `bundleScripts`). `npm test` passes.
```
