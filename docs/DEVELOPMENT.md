# 개발 · 배포

Office Kit를 소스에서 빌드하고 릴리스하는 방법. 사용 방법은 [README](../README.md).

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
- 소스 꾸러미(AGPL): `node scripts/make-office-source.mjs` → `dist/office-editor-source-<버전>.zip` ([`LICENSES-OFFICE-EDITOR.md`](LICENSES-OFFICE-EDITOR.md)).

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
| `site/` | 안내 사이트 (GitHub Pages) |

구조와 모듈 약속: [`ARCHITECTURE.md`](ARCHITECTURE.md)

시험은 실제 설정 파일·레지스트리·네트워크를 건드리지 않는다 (임시 폴더와 가짜 실행기 사용).

## 배포

`v*` 태그를 올리면 GitHub Actions가 설치 파일을 만들어 Release에 올린다.

```powershell
npm version <버전> --no-git-tag-version   # package.json 버전 = 태그 버전이어야 함
git commit -am "v<버전>"
git tag v<버전>
git push origin main v<버전>
```

- 워크플로: [`.github/workflows/release.yml`](../.github/workflows/release.yml) — 시험(먼저 `npm run rhwp-cli:build`) → `npm run editor:build` → `npm run office-editor:build` → `npm run rhwp-cli:build` → `electron-builder --publish always` (`GH_TOKEN` = 기본 `GITHUB_TOKEN`)
  → Office 편집기 소스 꾸러미 `office-editor-source-<버전>.zip`을 같은 Release에 붙인다.
  설치 파일에는 `vendor/rhwp-studio`·`vendor/rhwp-editor-sdk`(설치 파일 약 7MB, 설치 폴더 약 23MB 늘어남)와
  `vendor/office-editor`(설치 파일 약 50~65MB, 설치 폴더 약 213MB 늘어남), `vendor/rhwp-cli`(설치 파일 약 14MB, 설치 폴더 약 47MB 늘어남)가 들어간다.
- **남에게 배포하기 전에** [`LICENSES-OFFICE-EDITOR.md`](LICENSES-OFFICE-EDITOR.md)의 체크리스트 (법률 검토, 거울 저장소, 소스 꾸러미 확인).
- 수동 실행(Actions → CI · Release → Run workflow)은 초안(draft) Release로 올린다.
- main push / PR 마다 `npm test` 와 편집기 빌드 스크립트 확인이 돌고, main push에서는 편집기 묶음도 실제로 빌드해 본다 (태그 전에 깨진 것을 알게).
