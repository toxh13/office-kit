# Office 편집기 라이선스 — 무엇이 AGPL이고, 배포 전에 무엇을 하나

> **법률 자문이 아니다.** 라이선스 파일을 읽고 정리한 것이다. **처음 남에게 배포하기 전에 법률 검토를 받는다** (맨 아래 '검토 때 물을 것').
> 설계: [`OFFICE-EDITOR-DESIGN.md`](OFFICE-EDITOR-DESIGN.md) 8장. 대화상자 글: [`office-editor-legal-ko.md`](office-editor-legal-ko.md).

## 한눈에

- **Office Kit 자체는 MIT**다. 그대로 둔다.
- **Office 편집기 묶음 `vendor/office-editor/`는 AGPL-3.0**이다 (ONLYOFFICE 부분은 추가 조항 1~5가 붙음).
  - Office Kit 저장소 안에서 AGPL인 파일은 두 종류뿐: filex 패치(`patches/office-editor/`, 머리 `SPDX-License-Identifier: AGPL-3.0-or-later`), 글꼴 목록 생성기(`src/app/office-editor/fonts/`, 머리 `SPDX-License-Identifier: AGPL-3.0-only` — ONLYOFFICE core를 옮긴 것이라 그 라이선스 그대로 + 추가 조항).
  - 그 전문은 저장소의 `LICENSES/AGPL-3.0.txt`(filex `LICENSE`와 같은 바이트), 추가 조항은 `LICENSES/ONLYOFFICE-additional-terms.txt`(core `v9.4.0.129` `LICENSE` 끝부분 그대로).
- **설치 파일을 남에게 주는 것 = AGPL 6조의 '목적 코드 배포'** → 소스를 줄 의무가 생긴다.
  - 릴리스마다 `office-editor-source-<버전>.zip`을 같은 Release에 붙인다 (`scripts/make-office-source.mjs`).
  - ONLYOFFICE 원본 소스는 정확한 태그 주소 + **우리 거울 저장소**(아직 없음 — 할 일).
- 편집 창 위 막대의 **`정보 · 라이선스`** 가 ONLYOFFICE 추가 조항 3의 화면 고지다. 편집기 안 로고·'정보'는 지우지 않는다.
- **'ONLYOFFICE'는 사실을 말할 때만** 쓴다 ("ONLYOFFICE Docs 기반"). 제품·창·메뉴·기능 이름, 설치 파일, 아이콘에는 쓰지 않는다.
- **다음 행동 하나**: 아래 '배포 전 체크리스트'의 법률 검토와 거울 저장소를 사람이 처리한다.

---

## 1. 무엇이 어떤 라이선스인가

| 들어가는 것 | 어디 | 라이선스 | 저작권자 | 근거 |
|---|---|---|---|---|
| 편집기 web-apps·sdkjs (ONLYOFFICE Docs CE 9.4.0, 빌드 9.4.0.129) | `vendor/office-editor/editor/` | AGPL-3.0 + 추가 조항 1~5. 코드 아닌 것(그림·아이콘·설명 글)은 CC BY-SA 4.0 | Ascensio System SIA | `editor/LICENSE.txt` 662줄부터, 모든 파일 머리글 |
| 변환기 x2t (WASM) `v9.4.0.129+2` | `vendor/office-editor/x2t/` | ONLYOFFICE core AGPL-3.0 + CryptPad 패치 AGPL-3.0-or-later + BRF Tech 패치 | Ascensio System SIA, CryptPad 개발자, BRF Tech | filex `NOTICE` 4번, `upstream/onlyoffice.json` |
| filex-office-editor v0.3.1 (앱 페이지·다리 코드) | `vendor/office-editor/index.html`, `filex/`, `editor/web-apps/vendor/socketio/socket.io.min.js` | AGPL-3.0-or-later. `src/locks.ts`만 AGPL-3.0-only (ONLYOFFICE 서버 코드를 옮김) | BRF Tech | filex `LICENSE`, `NOTICE` |
| 편집기 안 제3자 JS | `editor/web-apps/vendor/` 등 | MIT · BSD · Apache-2.0 (Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO …) | 각 저작자 | `editor/3rd-Party.txt`, `editor/license/` |
| x2t·sdkjs WebAssembly 안 제3자 코드, 라이선스 파일 없는 `es6-promise`·`fetch` | `x2t/x2t.wasm`, `editor/sdkjs/common/*/engine/*.wasm`, `editor/web-apps/vendor/{es6-promise,fetch}` | BSL-1.0(Boost·Crypto++), ICU, FTL(FreeType), MIT(HarfBuzz·Brotli·libxml2·katana·md4c·musl …), BSD(OpenJPEG·Adobe CMap), IJG(libjpeg), libtiff, libpng, JasPer, Apache-2.0(gumbo, libc++ LLVM 예외), MPL(hyphen·librevenge …), GPL-2.0-or-later(DjVuLibre·JBIG-KIT), GPL-2/3(Xpdf 4.03), LGPL-2.1+(libpsd) | 각 저작자 | **`THIRD-PARTY-NATIVE.txt`** (Office Kit이 더함 — `scripts/office-editor/third-party.mjs`가 고정 주소·SHA-256으로 받아 `LICENSES/office-editor-third-party-native.txt`에 엮음) |
| filex 패치 | `patches/office-editor/filex/*.patch` | **AGPL-3.0-or-later** (filex를 고치는 것) | toxh13 | 파일 머리 |
| 글꼴 목록·미리보기 생성기 | `src/app/office-editor/fonts/*` | **AGPL-3.0-only + ONLYOFFICE 추가 조항** (ONLYOFFICE core AllFontsGen의 형식·규칙을 옮김 — core가 AGPL 3판만이라 -or-later로 바꿀 수 없다) | toxh13 (실험의 CryptPad·filex 생성기 코드를 가져오면 그 저작자도 머리에 적는다) | 파일 머리 |
| 묶음 만들기·덜어 내기 규칙·소스 꾸러미 스크립트 | `scripts/build-office-editor.mjs`, `scripts/office-editor/*`, `scripts/make-office-source.mjs` | MIT (우리 코드). 단 대응 소스의 일부로 소스 꾸러미에 넣는다 | toxh13 | — |
| 호스트 페이지·서버·창·저장 뒤 고치기 | `src/app/office-editor/host.*`, `editor-server.mjs`, `editor-window.mjs`, `src/core/ooxml-fixups.mjs` | MIT | toxh13 | `LICENSE` |
| 편집용 글꼴 | — | **넣지 않음**. 실행 중에 이 PC의 글꼴을 읽기만 한다 (묶음의 `editor/fonts`는 덜어 냄) | — | — |
| PDF 기본 14 글꼴 (URW++ Nimbus Sans·Roman·Mono, Standard Symbols L, Dingbats — Type 1) | `x2t/x2t.wasm` 안 (core `PdfFile/Resources/Fonts`) | **AGPL-3.0 + PDF·PostScript 포함 예외** (자유 글꼴) | URW++ | core `PdfFile/Resources/Fonts/LICENSE`, `THIRD-PARTY-NATIVE.txt` |

- 묶음은 filex `ui.zip`에서 오는데, filex가 ONLYOFFICE 파일을 묶으며 바꾼 것이 있다 (HTML 안 스크립트를 옆 파일로 옮김, 라이선스 파일 이름에 `.txt`, 빈 문서 `.bin`, `socket.io.min.js` 자리에 다리). 모두 filex `NOTICE`와 `editor/filex/CHANGES.txt`에 날짜와 함께 적혀 있다.
- Office Kit이 하는 수정(파일 통째로 빼기, filex 패치, 실행 중 `AllFonts.js` 바꿔 내주기)은 `vendor/office-editor/NOTICE-OfficeKit.md`에 날짜와 함께 적는다.

## 2. 경계 — MIT 앱과 AGPL 편집기가 한 설치 파일에 있는 이유

- 편집기는 **따로 도는 프로그램**으로 다룬다: 로컬 서버(127.0.0.3)가 파일을 내주고, 편집기는 sandbox 틀(opaque origin)·다른 렌더러 프로세스에서 돈다. Office Kit 호스트와는 **MessagePort 메시지**(filex 앱 SDK v1)로만 말한다. 서로의 코드를 링크하거나 한 파일로 묶지 않는다.
- 그래서 설치 파일은 MIT 프로그램과 AGPL 프로그램을 **함께 담은 것**(AGPL 5조 '집합물')으로 본다 — **검토 대상**이다.
- 경계를 지키는 규칙:
  - ONLYOFFICE·filex·CryptPad 코드를 옮기거나 고친 파일은 AGPL로 둔다 (위 표의 두 종류). 실험(`tmp/spike-*`)에서 코드를 가져올 때도 같다.
  - MIT 파일에 AGPL 코드를 붙여 넣지 않는다. 필요하면 그 부분을 AGPL 파일로 따로 둔다.
  - 호스트는 편집기 틀 안에 스크립트를 넣지 않는다 (설계 6장).

## 3. AGPL과 ONLYOFFICE 추가 조항이 요구하는 것

| 조항 | 요구 | Office Kit이 하는 것 |
|---|---|---|
| AGPL 4·5조, 추가 조항 1 | 저작권·라이선스·보증 부인·출처 고지를 **모두 남긴다** | 묶음은 파일을 통째로만 빼고 남은 파일은 한 바이트도 안 고친다. `LICENSE.txt`, `3rd-Party.txt`, `license/`, filex `LICENSE`·`NOTICE`·`CHANGES.txt`는 늘 남긴다 (`check-editor-bundle.cjs`가 확인) |
| AGPL 5조 a, 추가 조항 2 | 수정판은 **수정했다는 것과 날짜**, ONLYOFFICE(Ascensio System SIA) 원본 바탕임을 눈에 띄게 | `NOTICE-OfficeKit.md` (묶음 안, 소스 꾸러미 안) |
| AGPL 5조 d, 추가 조항 3 | 대화형 화면에서 쉽게 찾을 수 있게: (i) 원 개발사 ONLYOFFICE, (ii) 수정판일 수 있음, (iii) 라이선스 정보 | filex 앱 아래쪽 법적 고지 줄을 **늘 보이게** + 편집기 자체 '정보' + 위 막대 **`정보 · 라이선스`** 대화상자 ([`office-editor-legal-ko.md`](office-editor-legal-ko.md)) |
| 추가 조항 4 | 상표·로고 사용권 **없음** | 아래 5절 |
| 추가 조항 5 | 코드 아닌 것은 CC BY-SA 4.0 | 아래 6절 |
| AGPL 6조 | 목적 코드(설치 파일)를 주면 **대응 소스**를 함께 또는 같은 곳에서 | 릴리스마다 소스 꾸러미를 같은 Release에 + 위쪽 태그 주소 + 거울 (아래 7·8절) |
| AGPL 13조 | 네트워크로 쓰게 하면 그 사용자에게 소스 제공 | **해당 없음**: 편집기는 이 PC 사용자만 127.0.0.x로 쓴다. 누군가 고쳐서 남에게 서비스하면 그때는 해당 |
| AGPL 15·16조 | 보증 없음, 책임 제한 | 대화상자에 "아무런 보증이 없습니다" |

## 4. 남에게 배포하기 전 체크리스트

GitHub Release로 설치 파일을 올리는 것도, 메일·USB·사내 서버로 주는 것도 배포다.

- [ ] **법률 검토** (11절 질문 목록을 들고)
- [ ] **거울 저장소** 만들기 (9절 — 스크립트는 만들지 않는다, 사람이 한다)
- [ ] Release에 `office-editor-source-<버전>.zip`과 `.sha256`이 붙었는지 (CI가 확인하지만 눈으로도)
- [ ] 꾸러미의 `MANIFEST.json`이 `"complete": true`인지 (`--allow-missing`으로 만든 개발용이 아닌지), `officeKit.dirty`가 `false`인지
- [ ] 설치한 앱에서 Office 편집 창 → **`정보 · 라이선스`** 를 열어 본다: 원 개발사, 수정판 안내, 라이선스 전문(LICENSE.txt), 소스 주소가 **설치된 버전의 태그**를 가리키는지
- [ ] 편집기 아래쪽 법적 고지 줄이 보이고, 편집기 자체 '정보'(로고 포함)가 그대로인지
- [ ] 상표 점검 (5절): 설치 파일 이름·설명, 창 제목, 메뉴, 기능 이름, 아이콘, README 첫머리에 'ONLYOFFICE'가 없는지
- [ ] `NOTICE-OfficeKit.md`의 날짜·수정 목록이 이번 묶음과 맞는지
- [ ] **GitHub Release 말고 다른 길로 줄 때**: 설치 파일과 함께 소스 꾸러미도 준다. 못 주면 3년 이상 유효한 서면 제안(소스를 달라고 하면 주겠다는 글)을 함께 준다 (AGPL 6조 b)
- [ ] 묶음을 더 고쳤다면: `NOTICE-OfficeKit.md`에 날짜·내용을 더하고 소스 꾸러미를 다시 만든다

## 5. 상표 규칙

| 이름 | 써도 되는 곳 (사실만) | 쓰지 않는 곳 |
|---|---|---|
| **ONLYOFFICE®, ONLYOFFICE Docs™** (Ascensio System SIA) | "ONLYOFFICE Docs 기반", "원 개발사 Ascensio System SIA(ONLYOFFICE)" — 라이선스 대화상자, NOTICE, README 라이선스·구조 절, 이 문서 | 제품 이름, 창 제목, 메뉴, 기능 이름, 설치 파일 이름·설명, 아이콘, README·소개 글의 첫머리, 광고 |
| ONLYOFFICE **로고** | 편집기 안에 원래 있는 그대로 (지우면 추가 조항 1과 부딪히고, 바꾸면 상용판 기능인 화이트 라벨이 된다) | Office Kit 화면·아이콘·설치 파일에 새로 넣기, 편집기 안 로고 지우기·바꾸기 |
| **Microsoft Word·Excel·PowerPoint** | 형식 설명: "Word 문서(.docx)", "PowerPoint 형식" | 편집 창 이름 → '문서 편집 · 표 편집 · 슬라이드 편집' |
| **filex** (BRF Tech), **CryptPad** | 출처·감사: "filex-office-editor(BRF Tech)를 씀" | 제품·기능 이름 |

- 늘 함께 적는 문장: "ONLYOFFICE®·ONLYOFFICE Docs™는 Ascensio System SIA의 상표다. Office Kit은 Ascensio System SIA와 관계가 없고 보증·후원을 받지 않는다."
- 상표 정책: <https://www.onlyoffice.com/trademark-policy>

## 6. CC BY-SA 4.0 (코드 아닌 것)

- 편집기 안 그림·아이콘·설명 글은 CC BY-SA 4.0이다. `정보 · 라이선스`에 저작자(Ascensio System SIA)와 라이선스 주소를 적는다.
- 그런 아이콘·그림을 고쳐 쓰면 고친 것도 CC BY-SA 4.0으로 내고, 고쳤다는 것을 적는다. Office Kit 자신의 아이콘(`build/icon.*`)과는 섞지 않는다.

## 7. 소스 꾸러미 — `scripts/make-office-source.mjs`

```powershell
npm run office-editor:build                   # 먼저 묶음 (vendor/office-editor)
node scripts/make-office-source.mjs           # → dist/office-editor-source-<버전>.zip + .sha256
node scripts/make-office-source.mjs --dry-run --allow-missing   # 확인만 (개발용)
```

| 꾸러미 안 | 무엇 |
|---|---|
| `filex-office-editor/` | filex 소스, 고정 커밋 `57dda3da…`(태그 `v0.3.1`) 그대로 (`git archive`, 파일마다 git blob 해시 확인). x2t 빌드 방법 `scripts/x2t/` 포함 |
| `office-kit/` | `scripts/build-office-editor.mjs`, `scripts/office-editor/`(덜어 내기 규칙), `patches/office-editor/`, `src/app/office-editor/fonts/`, 이 스크립트, 라이선스 문서, `package.json`, `LICENSE` |
| `NOTICE-OfficeKit.md`, `bundle/VERSION.json`, `licenses/` | 묶음의 수정 고지·고정 값, ONLYOFFICE `LICENSE.txt`·`3rd-Party.txt`·`license/` |
| `UPSTREAM.md`, `upstream.json` | ONLYOFFICE sdkjs·web-apps·core·build_tools `v9.4.0.129`의 주소·커밋, x2t 빌드 방법, 거울 할 일 |
| `README.md`, `MANIFEST.json` | 다시 만드는 법, 파일마다 SHA-256 |

- 빠진 부분(묶음을 안 만들었거나 다른 일감 파일이 없음)이 있으면 멈춘다. `--allow-missing`은 개발용이고 꾸러미 `README.md`에 "개발용" 경고가 붙는다.
- filex 소스의 `upstream/onlyoffice.json`이 스크립트의 고정 값과 다르면 멈춘다. `docs/office-editor-legal-ko.md`의 판 번호·주소가 다르면 멈춘다.
- 같은 입력이면 같은 바이트다 (파일 시각 = `SOURCE_DATE_EPOCH` 또는 Office Kit HEAD 커밋 시각).
- filex 커밋 메모: 고정 커밋은 태그 `v0.3.1`이 가리키는 `57dda3da`(릴리스 `ui.zip`을 만든 소스, `scripts/office-editor/pins.mjs`)다. 설계 4장 표의 `acc87f88`은 그 다음 **README.md만** 고친 main 커밋이다 (GitHub 비교로 확인).
- 고정 값은 `scripts/office-editor/pins.mjs`(묶음 만들기와 같은 것)에서 가져오고, 같은 태그의 저장소 커밋만 스크립트에 더 적는다. 둘이 어긋나면 멈춘다.
- CI(release 작업)가 설치 파일을 올린 뒤 이 꾸러미를 만들어 같은 Release에 붙이고, 공개 전에 붙었는지 확인한다.

## 8. 위쪽 소스 주소 (2026-10-18 `git ls-remote`로 확인)

| 저장소 | 태그 | 커밋 | 무엇 |
|---|---|---|---|
| [ONLYOFFICE/sdkjs](https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129) | `v9.4.0.129` | `75b53d504bfd81bfa35d59bb57dd9e56f7c9235d` | 편집기 엔진 |
| [ONLYOFFICE/web-apps](https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129) | `v9.4.0.129` | `1993a6d89731af0fcb5129ebf0ee380076fea5a9` | 편집기 화면 |
| [ONLYOFFICE/core](https://github.com/ONLYOFFICE/core/tree/v9.4.0.129) | `v9.4.0.129` | `a016fc2868b9f99f59dd30aa2088e68bbe5104af` | x2t 소스 |
| [ONLYOFFICE/build_tools](https://github.com/ONLYOFFICE/build_tools/tree/v9.4.0.129) | `v9.4.0.129` | `cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5` | 공식 빌드 도구 |
| [ONLYOFFICE/server](https://github.com/ONLYOFFICE/server/tree/v9.4.0.129) | `v9.4.0.129` | `13142e41dfc9e1c337bbe76f0222d4b2ce7efbc8` | filex `src/locks.ts`가 옮긴 잠금 규칙의 원본 |
| [BRF-Tech/filex-office-editor](https://github.com/BRF-Tech/filex-office-editor/tree/v0.3.1) | `v0.3.1` | `57dda3dae99af78eb26b3c64b69f1e3f711f0757` | 앱 페이지·다리·x2t 빌드 방법 |
| [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm) | `v9.3.2+3` · `v9.3.0+0` | `7debf5e6d6d4f19874d764cd34446f7e63c6bc32` · `96886ff143e05471144c4426fb304b4d794370d2` | x2t WASM 빌드 방법·core 변경의 출처 |

- 같은 태그의 관련 저장소(sdkjs-forms, document-templates, dictionaries, core-fonts, document-server-package)와 DocumentServer `v9.4.0`의 커밋은 소스 꾸러미 `UPSTREAM.md`에 있다.
- 편집기 파일은 공식 Docker 이미지 `docker.io/onlyoffice/documentserver:9.4.0.1` (`sha256:3ab6ebc7…88212`)에서 꺼낸 것이다.
- 앱 안 링크는 **설치된 버전의 태그**를 가리킨다 (`{{appVersion}}`).

## 9. 거울 저장소 — 할 일 (아직 없음)

위쪽 태그가 지워지거나 옮겨져도, 이 판을 배포하는 동안 소스를 줄 수 있게 둔다 (AGPL 6조: 소스가 다른 서버에 있어도 **계속 받을 수 있게 할 책임은 배포하는 쪽**).
Office Kit 스크립트·에이전트는 거울·포크를 만들지 않는다. 사람이 만들고 아래 표의 주소 칸을 채운다.

| 원본 | 지킬 것 | 거울 주소 |
|---|---|---|
| ONLYOFFICE/sdkjs | 태그 `v9.4.0.129` | (비어 있음) |
| ONLYOFFICE/web-apps | 태그 `v9.4.0.129` | (비어 있음) |
| ONLYOFFICE/core | 태그 `v9.4.0.129` | (비어 있음) |
| ONLYOFFICE/build_tools | 태그 `v9.4.0.129` | (비어 있음) |
| ONLYOFFICE/server | 태그 `v9.4.0.129` | (비어 있음) |
| BRF-Tech/filex-office-editor | 태그 `v0.3.1`(커밋 `57dda3da`), 릴리스 파일 `ui.zip`·`x2t.zip` | (비어 있음) |
| cryptpad/onlyoffice-x2t-wasm | 태그 `v9.3.2+3`, `v9.3.0+0` | (비어 있음) |
| x2t를 만드는 데 쓰는 제3자 소스 (filex `upstream/onlyoffice.json` `x2t.build.sources`) | boost 1.84.0 (`archives.boost.io` `boost_1_84_0.tar.bz2`, SHA-256 `cc4b893a…`), OpenSSL `OpenSSL_1_1_1f` 헤더, hunspell/hyphen `dcd90f00` | (비어 있음) |
| core 받기 스크립트가 받는 것 (core `Common/3dParty/*/fetch.py`·`make.py`) | gumbo-parser `aa91b27b`, katana-parser `be6df458`, md4c `481fbfbd`, glm `33b4a621`, mdds `07831589`, librevenge `becd044b`, libodfgen `8ef8c171`, libetonyek `cb396b4a`, harfbuzz `894a1f72`, brotli `a47d7475` | (비어 있음) |
| emscripten 4.0.11 ICU 포트 | unicode-org/icu `release-68-2` (`icu4c-68_2-src.zip`) | (비어 있음) |
| filex 앱 UI (npm) | `@brftech/filex-app-ui` 0.53.0 (filex `package-lock.json`) | (비어 있음) |

- 위쪽이 사라지면 소스를 줄 수 없게 되는 것은 ONLYOFFICE·filex 저장소만이 아니다: x2t.wasm의 대응 소스에는 위 제3자 소스도 들어간다. 거울 대신 그 소스 묶음(tarball)을 Release마다 붙여도 된다 (검토 11절 9번).

```bash
# 포크보다 --mirror가 확실하다 (태그까지 그대로 옮긴다)
git clone --mirror https://github.com/ONLYOFFICE/sdkjs.git
cd sdkjs.git
git push --mirror https://github.com/<거울 계정>/sdkjs.git
```

- 거울을 만든 뒤: 이 표, `office-editor-legal-ko.md`의 '소스 코드' 절(거울 주소 한 줄), 소스 꾸러미 `UPSTREAM.md`(스크립트의 `PINS`)에 주소를 더한다.
- 큰 저장소(core)는 그 태그만 밀어도 된다: `git push <거울> refs/tags/v9.4.0.129`.

## 10. 버전 올릴 때 함께 바꾸는 것

- `scripts/build-office-editor.mjs`의 고정 값 (filex 태그·커밋·`ui.zip` 크기·SHA-256, x2t SHA-256)
- `scripts/make-office-source.mjs`의 `PINS` (+ 새 태그의 커밋을 `git ls-remote`로 확인)
- `docs/office-editor-legal-ko.md`의 판 번호·소스 주소 (스크립트가 맞는지 확인)
- `LICENSE`의 THIRD-PARTY 항목, 이 문서의 1·8·9절, `README.md` 라이선스 절
- 올린 뒤: e2e + Office 엄격 열기 전체 통과 (설계 12장)

## 11. 법률 검토 때 물을 것

1. MIT 앱과 AGPL 편집기 묶음을 한 설치 파일로 내는 것이 '집합물'(AGPL 5조)로 충분한가. 경계(2절)를 어떻게 더 분명히 하면 좋은가.
2. 소스 제공 방법: 같은 GitHub Release의 소스 꾸러미 + 위쪽 태그 주소 + 거울 — AGPL 6조 d로 충분한가. 사내·USB 배포 때는 어떻게 하나 (회사 안에서 직원에게 주는 것도 배포인가).
3. ONLYOFFICE 추가 조항(특히 로고·고지 유지)을 모두 따르는 것을 전제로, 그 해석과 범위 (7조 '추가 제한' 문단과의 관계 포함).
4. 상표: 라이선스 대화상자·README·NOTICE에서 "ONLYOFFICE Docs 기반"이라고 쓰는 것, 편집기 안 로고를 그대로 두는 것.
5. filex의 `src/locks.ts`(AGPL-3.0-only)와 나머지(AGPL-3.0-or-later)를 함께 내는 것, filex 패치는 -or-later, 글꼴 생성기는 ONLYOFFICE core를 따라 AGPL-3.0-only(+ 추가 조항)로 두는 것.
6. 글꼴: 설치 파일에 넣지 않고 실행 중에 사용자 PC의 글꼴을 읽어(앞 32바이트 XOR) 편집기에 주는 것, PDF 내보내기에서 그 글꼴이 PDF에 들어가는 것. x2t.wasm 안 URW++ 기본 14 글꼴(AGPL-3.0 + PDF 포함 예외)의 고지 방법(`THIRD-PARTY-NATIVE.txt`).
7. CC BY-SA 4.0 표시(6절)가 충분한가.
8. 대화상자 문구([`office-editor-legal-ko.md`](office-editor-legal-ko.md))가 추가 조항 3의 (i)~(iii)을 채우는가.
9. x2t.wasm 대응 소스의 제3자 부분(boost·ICU·hyphen·core 받기 스크립트가 받는 것, 9절 표): 거울로 둘지, Release마다 소스 묶음으로 붙일지.
10. `THIRD-PARTY-NATIVE.txt`의 범위: core의 qmake 프로젝트·3DPARTY.md를 따라 엮었다 (링크된 것보다 넓을 수 있음). 특히 Xpdf 4.03(GPL-2.0 또는 GPL-3.0 — GPL-3.0으로 AGPL-3.0과 함께), DjVuLibre·JBIG-KIT(GPL-2.0-or-later), libpsd(LGPL-2.1-or-later)가 ONLYOFFICE core(AGPL-3.0) 안에 든 것을 그대로 내는 것.
