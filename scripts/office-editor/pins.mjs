// pins.mjs — Office 편집기 묶음의 고정 값 (scripts/build-office-editor.mjs, make-office-source 등이 함께 씀).
//
// 올릴 때는 이 파일만 바꾸고, e2e·COM 엄격 열기 시험을 다시 통과한 뒤 함께 올린다
// (docs/OFFICE-EDITOR-DESIGN.md 4장 '고정 값'). CI 캐시 키도 이 파일의 해시다 (.github/workflows/release.yml).
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';

// filex-office-editor (BRF Tech, AGPL-3.0-or-later)
//   commit은 태그 v0.3.1이 가리키는 커밋이다 (ui.zip을 만든 소스). 실험(tmp/spike-filex/repo)이 받은
//   acc87f88da47…는 그 다음 main 커밋으로 README.md만 다르다 — 빌드에 쓰는 파일은 같다.
export const FILEX = Object.freeze({
  name: 'filex-office-editor',
  repo: 'https://github.com/BRF-Tech/filex-office-editor',
  git: 'https://github.com/BRF-Tech/filex-office-editor.git',
  tag: 'v0.3.1',
  version: '0.3.1',
  commit: '57dda3dae99af78eb26b3c64b69f1e3f711f0757',
  license: 'AGPL-3.0-or-later',
});

// 릴리스에 붙은 앱 묶음 (편집기 + x2t + 앱 페이지)
export const UI_ZIP = Object.freeze({
  url: `${FILEX.repo}/releases/download/${FILEX.tag}/ui.zip`,
  file: `ui-${FILEX.tag}.zip`,
  bytes: 103098008,
  sha256: 'cd3bc0741f97e471ec26d639da65835fccff2605fc255974ecb5f21d2b179baf',
  files: 2633,
});

// filex upstream/editor.lock.json (고정 커밋의 것) — 편집기 파일마다 "sha256 크기"
export const EDITOR_LOCK = Object.freeze({
  url: `https://raw.githubusercontent.com/BRF-Tech/filex-office-editor/${FILEX.commit}/upstream/editor.lock.json`,
  path: 'upstream/editor.lock.json',
  file: `editor.lock-${FILEX.commit.slice(0, 12)}.json`,
  bytes: 423699,
  sha256: '8b5ac33823287cd05110bca0d79a3965a3b5a8d972e6a375bc925de9435f3375',
  files: 2623,
});

// ONLYOFFICE Docs Community Edition (Ascensio System SIA, AGPL-3.0 + 추가 조항)
export const ONLYOFFICE = Object.freeze({
  product: 'ONLYOFFICE Docs Community Edition',
  vendor: 'Ascensio System SIA',
  version: '9.4.0',
  build: '9.4.0.129',
  image: 'docker.io/onlyoffice/documentserver@sha256:3ab6ebc7c605e5a32b7ae3ff19daed4925090245acc8100ce2230bd766c88212',
  sourceTag: 'v9.4.0.129',
  sources: Object.freeze({
    'web-apps': 'https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129',
    sdkjs: 'https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129',
    core: 'https://github.com/ONLYOFFICE/core/tree/v9.4.0.129',
    build_tools: 'https://github.com/ONLYOFFICE/build_tools/tree/v9.4.0.129',
  }),
});

// x2t (WASM 변환기) — filex 릴리스 v9.4.0.129+2, 소스는 filex upstream/onlyoffice.json
export const X2T = Object.freeze({
  release: 'v9.4.0.129+2',
  files: Object.freeze({
    'x2t.js': '9f2f65ac557470b2895beaa0f9678ee1a998545d49c8b93d67306ff3e7d7fe4b',
    'x2t.wasm': 'cabaa6e33574f61703c53463d714e33529949d17d388b4e198994199e29c368b',
  }),
  core: Object.freeze({ repo: 'https://github.com/ONLYOFFICE/core.git', tag: 'v9.4.0.129', commit: 'a016fc2868b9f99f59dd30aa2088e68bbe5104af' }),
  buildTools: Object.freeze({ repo: 'https://github.com/ONLYOFFICE/build_tools.git', tag: 'v9.4.0.129', commit: 'cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5' }),
  recipe: 'https://github.com/BRF-Tech/filex-office-editor/tree/57dda3dae99af78eb26b3c64b69f1e3f711f0757/scripts/x2t',
});

// 5단계(앱 다시 빌드)에만: filex build-app.mjs는 편집기 파일이 lock과 똑같기를 바란다.
// 릴리스 묶음에는 다리(bridge)가 이 파일 자리에 들어가 있어, 원래 파일(Socket.IO 4.5.3 클라이언트)을
// ONLYOFFICE web-apps 같은 태그에서 받아 lock의 SHA-256으로 확인한다.
export const SOCKET_IO_ORIGINAL = Object.freeze({
  path: 'web-apps/vendor/socketio/socket.io.min.js',
  url: `https://raw.githubusercontent.com/ONLYOFFICE/web-apps/${ONLYOFFICE.sourceTag}/vendor/socketio/socket.io.min.js`,
  file: `socket.io.min-${ONLYOFFICE.sourceTag}.js`,
  bytes: 44004,
  sha256: '93708097d4f4097bc5884b4bf8f55722e98ecdb4ce8bde297bcdc223d6941464',
});

// 5단계에서 받는 filex 소스 폴더 (성긴 받기). 패치가 다른 폴더를 고치면 그 폴더도 더 받는다.
export const FILEX_SPARSE_DIRS = Object.freeze(['app', 'scripts', 'src', 'upstream']);

/** 고정 값 전체 (VERSION.json, 건너뛰기 비교, 캐시 키) */
export const PINS = Object.freeze({ filex: FILEX, uiZip: UI_ZIP, editorLock: EDITOR_LOCK, onlyoffice: ONLYOFFICE, x2t: X2T, socketIoOriginal: SOCKET_IO_ORIGINAL });

export function pinsSha256() {
  return createHash('sha256').update(JSON.stringify(PINS)).digest('hex');
}

/** 캐시 폴더: OFFICEKIT_OFFICE_EDITOR_CACHE, 없으면 <저장소>/tmp/office-editor-build */
export function defaultCache(root) {
  return resolve(process.env.OFFICEKIT_OFFICE_EDITOR_CACHE || join(root, 'tmp', 'office-editor-build'));
}

/**
 * 캐시 폴더 안 자리 (OFFICEKIT_OFFICE_EDITOR_CACHE, 기본 <저장소>/tmp/office-editor-build).
 *   download/   ui.zip·lock·원래 socket.io (CI가 이 폴더만 캐시한다)
 *   ui-v0.3.1/  ui.zip을 그대로 푼 전체 (덜어 내기 전 — editor/fonts·AllFonts.js 기준 시험용)
 *   filex-src-<커밋 12자>/  5단계 filex 소스 (고정 커밋, 패치 적용)
 */
export function cachePaths(cache) {
  return {
    root: cache,
    download: join(cache, 'download'),
    uiZip: join(cache, 'download', UI_ZIP.file),
    lock: join(cache, 'download', EDITOR_LOCK.file),
    socketIo: join(cache, 'download', SOCKET_IO_ORIGINAL.file),
    ui: join(cache, `ui-${FILEX.tag}`),
    filexSrc: join(cache, `filex-src-${FILEX.commit.slice(0, 12)}`),
    lockEditor: join(cache, 'filex-lock-editor'),
    appOut: join(cache, 'filex-app-out'),
  };
}
