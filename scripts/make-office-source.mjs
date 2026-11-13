#!/usr/bin/env node
// SPDX-License-Identifier: MIT
// make-office-source.mjs — Office 편집기(AGPL) 소스 꾸러미 office-editor-source-<버전>.zip 만들기 (AGPL 6조).
//
//   node scripts/make-office-source.mjs [--out 폴더] [--version x.y.z] [--filex-src git폴더] [--filex-repo 주소]
//                                       [--offline] [--allow-missing] [--dry-run] [--root 폴더]
//
// 꾸러미 맨 위 폴더 office-editor-source-<버전>/ 안:
//   filex-office-editor/  filex 소스 — 고정 커밋을 git archive 그대로 (x2t 빌드 방법 scripts/x2t/ 포함)
//   office-kit/           Office Kit 쪽 — 묶음 만들기(build-office-editor.mjs), 덜어 내기 규칙(scripts/office-editor/),
//                         filex 패치(patches/office-editor/), 글꼴 목록 생성기(src/app/office-editor/fonts/), 이 스크립트, 문서
//   bundle/VERSION.json   만든 묶음의 고정 값·파일 목록 (vendor/office-editor/VERSION.json)
//   licenses/             ONLYOFFICE LICENSE.txt(AGPL 3판 + 추가 조항)·3rd-Party.txt·license/
//   NOTICE-OfficeKit.md   날짜 붙은 수정 고지 (vendor/office-editor/의 것, 없으면 --allow-missing일 때만 대신 씀)
//   UPSTREAM.md · upstream.json   ONLYOFFICE sdkjs·web-apps·core·build_tools v9.4.0.129 주소·커밋, x2t 빌드 방법, 거울 할 일
//   README.md · MANIFEST.json     다시 만드는 법 / 파일마다 크기·sha256, 빠진 것
//
// - 고정 값은 scripts/office-editor/pins.mjs(묶음 만들기와 같은 것) + 아래 소스 주소. docs/office-editor-legal-ko.md도 맞춰 본다.
// - filex 소스의 upstream/onlyoffice.json이 고정 값과 다르면 멈춘다 (다른 커밋을 담지 않게).
// - 같은 입력 = 같은 바이트: 파일 시각은 SOURCE_DATE_EPOCH 또는 Office Kit HEAD 커밋 시각, 이름순, 폴더 항목 없음.
// - filex 소스: --filex-src → OFFICEKIT_FILEX_SRC → 캐시(OFFICEKIT_OFFICE_EDITOR_CACHE, 기본 tmp/office-editor-build)의
//   git 폴더 중 그 커밋이 있는 것 → 없으면 그 커밋만 받는다 (--filex-repo·OFFICEKIT_FILEX_REPO = 거울 주소, 기본 GitHub.
//   --offline이면 받지 않고 멈춤).
// - 위쪽 소스 사본은 UPSTREAM_MIRROR에 따로 올려 두었다 (이 스크립트는 받거나 올리지 않는다).
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import JSZip from 'jszip';

import { FILEX, ONLYOFFICE, UI_ZIP, X2T } from './office-editor/pins.mjs';

const SCRIPT = fileURLToPath(import.meta.url);
const DEFAULT_ROOT = resolve(dirname(SCRIPT), '..');

// ── 고정 값 ──
// filex·ui.zip·ONLYOFFICE 판·x2t는 묶음 만들기와 같은 곳(scripts/office-editor/pins.mjs)에서 가져온다.
// 여기에는 소스 주소만 더한다: 같은 태그의 저장소 커밋(2026-10-18 git ls-remote), 이미지 태그, x2t 빌드 방법의 출처.
// 올릴 때: pins.mjs → 아래 REPOS·X2T_RECIPE → docs/office-editor-legal-ko.md (checkPins·checkLegalDoc이 맞는지 확인).
const SRC_TAG = 'v9.4.0.129';
// 위쪽 소스 사본 (커밋 그대로의 압축본 + SHA256SUMS) — docs/LICENSES-OFFICE-EDITOR.md 9절
const UPSTREAM_MIRROR = 'https://github.com/toxh13/office-kit-upstream/releases/tag/office-editor-9.4.0.129';
const REPOS = [
  { name: 'sdkjs', url: 'https://github.com/ONLYOFFICE/sdkjs', tag: SRC_TAG, commit: '75b53d504bfd81bfa35d59bb57dd9e56f7c9235d', main: true, role: '편집기 엔진 (묶음의 editor/sdkjs)' },
  { name: 'web-apps', url: 'https://github.com/ONLYOFFICE/web-apps', tag: SRC_TAG, commit: '1993a6d89731af0fcb5129ebf0ee380076fea5a9', main: true, role: '편집기 화면 (묶음의 editor/web-apps)' },
  { name: 'core', url: 'https://github.com/ONLYOFFICE/core', tag: SRC_TAG, commit: 'a016fc2868b9f99f59dd30aa2088e68bbe5104af', main: true, role: '변환기 x2t의 소스 (묶음의 x2t/)' },
  { name: 'build_tools', url: 'https://github.com/ONLYOFFICE/build_tools', tag: SRC_TAG, commit: 'cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5', main: true, role: 'sdkjs·web-apps를 묶는 공식 빌드 도구, x2t 빌드에도 씀' },
  { name: 'server', url: 'https://github.com/ONLYOFFICE/server', tag: SRC_TAG, commit: '13142e41dfc9e1c337bbe76f0222d4b2ce7efbc8', main: false, role: 'Document Server — filex src/locks.ts가 옮긴 잠금 규칙의 원본' },
  { name: 'sdkjs-forms', url: 'https://github.com/ONLYOFFICE/sdkjs-forms', tag: SRC_TAG, commit: '1a883281c41fa9aa5e98c64cce13b14782f92f56', main: false, role: 'sdkjs의 양식(forms) 모듈 — 묶음에 들어가는지 확인 필요' },
  { name: 'document-templates', url: 'https://github.com/ONLYOFFICE/document-templates', tag: SRC_TAG, commit: '71430c9f183489e8912f54f9dc859e369cf0dfb4', main: false, role: '빈 문서 서식 (묶음의 editor/document-templates)' },
  { name: 'dictionaries', url: 'https://github.com/ONLYOFFICE/dictionaries', tag: SRC_TAG, commit: 'd3223bbb777883db66ac3cd249f71c6ebdc992c7', main: false, role: '맞춤법 사전 (묶음에 없음 — 참고)' },
  { name: 'core-fonts', url: 'https://github.com/ONLYOFFICE/core-fonts', tag: SRC_TAG, commit: '7030c6681fb5bbed560675cb42422f91df15d5c9', main: false, role: '묶음 글꼴 (Office Kit은 덜어 냄 — 참고)' },
  { name: 'document-server-package', url: 'https://github.com/ONLYOFFICE/document-server-package', tag: SRC_TAG, commit: '0e73a4ff3c565da6b74b675959aad3ebbd48c793', main: false, role: 'Docker 이미지·설치 꾸러미 만드는 방법' },
  { name: 'DocumentServer', url: 'https://github.com/ONLYOFFICE/DocumentServer', tag: 'v9.4.0', commit: 'fb73d33c85f59d1a5d2e5a5ed05388ebc2418337', main: false, role: '릴리스 안내·라이선스' },
];
const X2T_RECIPE = { repo: 'https://github.com/cryptpad/onlyoffice-x2t-wasm', release: 'v9.3.2+3', commit: '7debf5e6d6d4f19874d764cd34446f7e63c6bc32', changes: 'v9.3.0+0', changesCommit: '96886ff143e05471144c4426fb304b4d794370d2' };
const [IMAGE, DIGEST] = ONLYOFFICE.image.split('@');

export const PINS = Object.freeze({
  filex: {
    name: FILEX.name,
    repo: FILEX.repo,
    git: FILEX.git || `${FILEX.repo}.git`,
    tag: FILEX.tag,
    commit: FILEX.commit,
    // 설계 4장 표의 acc87f88(main)은 이 태그 커밋 다음 README.md만 고친 커밋이다 (GitHub 비교로 확인)
    mainCommit: 'acc87f88da472bcf984f8c27baafe6ad788e3f61',
    uiZip: { url: UI_ZIP.url, bytes: UI_ZIP.bytes, sha256: UI_ZIP.sha256 },
    license: 'AGPL-3.0-or-later (src/locks.ts만 AGPL-3.0-only)',
    copyright: 'Copyright (C) 2026 BRF Tech',
  },
  onlyoffice: {
    product: ONLYOFFICE.product,
    vendor: ONLYOFFICE.vendor,
    version: ONLYOFFICE.version,
    build: ONLYOFFICE.build,
    sourceTag: ONLYOFFICE.sourceTag,
    image: IMAGE,
    imageTag: '9.4.0.1',
    digest: DIGEST,
    release: 'https://github.com/ONLYOFFICE/DocumentServer/releases/tag/v9.4.0',
    license: 'AGPL-3.0 + 추가 조항(7조) 1~5, 코드 아닌 것은 CC BY-SA 4.0',
    repos: REPOS,
  },
  x2t: {
    release: X2T.release,
    files: { ...X2T.files },
    recipe: X2T_RECIPE,
    license: 'ONLYOFFICE core AGPL-3.0 + CryptPad 패치 AGPL-3.0-or-later + BRF Tech 패치',
  },
  officeKit: {
    repo: 'https://github.com/toxh13/office-kit',
  },
});

/** pins.mjs와 여기 소스 주소가 서로 맞는지 → 틀린 것 목록 (pins.mjs만 올리고 여기를 잊은 경우) */
export function checkPins() {
  const bad = [];
  const repo = (n) => REPOS.find((r) => r.name === n);
  if (ONLYOFFICE.sourceTag !== SRC_TAG) bad.push(`ONLYOFFICE 소스 태그 ${ONLYOFFICE.sourceTag} ≠ REPOS의 ${SRC_TAG}`);
  for (const [name, url] of Object.entries(ONLYOFFICE.sources || {})) {
    const r = repo(name);
    if (!r || url !== `${r.url}/tree/${r.tag}`) bad.push(`${name} 주소 ${url}`);
  }
  if (X2T.core?.commit && X2T.core.commit !== repo('core').commit) bad.push(`x2t core 커밋 ${X2T.core.commit}`);
  if (X2T.buildTools?.commit && X2T.buildTools.commit !== repo('build_tools').commit) bad.push(`x2t build_tools 커밋 ${X2T.buildTools.commit}`);
  if (!/^[0-9a-f]{40}$/.test(FILEX.commit || '')) bad.push(`filex 커밋 ${FILEX.commit}`);
  if (!/^sha256:[0-9a-f]{64}$/.test(DIGEST || '')) bad.push(`이미지 digest ${ONLYOFFICE.image}`);
  return bad;
}

// 꾸러미에 넣는 Office Kit 쪽 부분: [id, 저장소 경로, 꾸러미 경로, 필수, 설명]
const PARTS = [
  { id: 'build', src: 'scripts/build-office-editor.mjs', dest: 'office-kit/scripts/build-office-editor.mjs', required: true, what: '묶음 만들기 (받기·확인·덜어 내기·패치 적용)' },
  { id: 'rules', src: 'scripts/office-editor', dest: 'office-kit/scripts/office-editor', dir: true, required: true, what: '덜어 내기 규칙·lock 확인' },
  { id: 'patches', src: 'patches/office-editor', dest: 'office-kit/patches/office-editor', dir: true, required: true, what: 'filex 패치 (AGPL-3.0-or-later)' },
  { id: 'fonts', src: 'src/app/office-editor/fonts', dest: 'office-kit/src/app/office-editor/fonts', dir: true, required: true, what: '글꼴 목록·미리보기 생성기 (AGPL-3.0-only + ONLYOFFICE 추가 조항)' },
  { id: 'agpl-texts', src: 'LICENSES', dest: 'office-kit/LICENSES', dir: true, required: true, what: 'AGPL 3판 전문·ONLYOFFICE 추가 조항·기계어 제3자 고지' },
  { id: 'self', src: 'scripts/make-office-source.mjs', dest: 'office-kit/scripts/make-office-source.mjs', required: true, what: '이 꾸러미를 만든 스크립트' },
  { id: 'licenses-doc', src: 'docs/LICENSES-OFFICE-EDITOR.md', dest: 'office-kit/docs/LICENSES-OFFICE-EDITOR.md', required: true, what: '라이선스 설명 (한국어)' },
  { id: 'legal-doc', src: 'docs/office-editor-legal-ko.md', dest: 'office-kit/docs/office-editor-legal-ko.md', required: true, what: "'정보 · 라이선스' 대화상자 글" },
  { id: 'design', src: 'docs/OFFICE-EDITOR-DESIGN.md', dest: 'office-kit/docs/OFFICE-EDITOR-DESIGN.md', required: false, what: '설계' },
  { id: 'patches-doc', src: 'docs/office-editor-patches.md', dest: 'office-kit/docs/office-editor-patches.md', required: false, what: 'filex 패치 설명' },
  { id: 'package', src: 'package.json', dest: 'office-kit/package.json', required: true, what: 'npm 스크립트 이름·의존성 버전' },
  { id: 'mit', src: 'LICENSE', dest: 'office-kit/LICENSE', required: true, what: 'Office Kit 라이선스 (MIT)' },
  { id: 'notices', src: 'THIRD-PARTY-NOTICES.txt', dest: 'office-kit/THIRD-PARTY-NOTICES.txt', required: true, what: '제3자 목록과 AGPL 범위' },
  { id: 'third-party', src: 'vendor/office-editor/THIRD-PARTY-NATIVE.txt', dest: 'licenses/THIRD-PARTY-NATIVE.txt', required: true, what: 'x2t·sdkjs WebAssembly 안 제3자 고지 (묶음의 것)' },
  { id: 'notice', src: 'vendor/office-editor/NOTICE-OfficeKit.md', dest: 'NOTICE-OfficeKit.md', required: true, what: '날짜 붙은 수정 고지 (묶음의 것)' },
  { id: 'version', src: 'vendor/office-editor/VERSION.json', dest: 'bundle/VERSION.json', required: true, what: '묶음의 고정 값·파일 목록' },
  { id: 'oo-license', src: 'vendor/office-editor/editor/LICENSE.txt', dest: 'licenses/ONLYOFFICE-LICENSE.txt', required: true, what: 'AGPL 3판 + ONLYOFFICE 추가 조항' },
  { id: 'oo-3rd', src: 'vendor/office-editor/editor/3rd-Party.txt', dest: 'licenses/ONLYOFFICE-3rd-Party.txt', required: false, what: '편집기 안 제3자 구성 요소' },
  { id: 'oo-license-dir', src: 'vendor/office-editor/editor/license', dest: 'licenses/onlyoffice-license', dir: true, required: false, what: '제3자 라이선스 전문' },
];

// 폴더에서 빼는 것 (생성물·개인 파일)
const SKIP = /(^|\/)(node_modules|\.git|\.cache|coverage)(\/|$)|(^|\/)\.DS_Store$|(^|\/)Thumbs\.db$|\.(log|tmp)$/i;
const MAX_FILE = 16 * 1024 * 1024;

const log = (s) => process.stdout.write(`[office-source] ${s}\n`);
export class SourceError extends Error {}
const fail = (msg) => { throw new SourceError(msg); };

const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');
const posix = (p) => p.split(sep).join('/');
const kb = (n) => (n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)} MB` : `${Math.ceil(n / 1024)} KB`);
const isDir = (p) => { try { return statSync(p).isDirectory(); } catch { return false; } };
const isFile = (p) => { try { return statSync(p).isFile(); } catch { return false; } };

const HELP = `Office 편집기 소스 꾸러미 만들기 → <out>/office-editor-source-<버전>.zip (+ .sha256)

  node scripts/make-office-source.mjs [옵션]

  --out 폴더        꾸러미를 둘 곳 (기본 dist)
  --version x.y.z   꾸러미 버전 (기본 package.json의 version)
  --filex-src 폴더  filex 고정 커밋(${PINS.filex.commit.slice(0, 8)})이 들어 있는 git 폴더 (OFFICEKIT_FILEX_SRC)
  --filex-repo 주소 filex 소스를 받을 곳 (기본 ${PINS.filex.git}, 거울을 쓸 때 — OFFICEKIT_FILEX_REPO)
  --offline         filex 소스를 받지 않음 (이 PC에 없으면 멈춤)
  --allow-missing   빠진 부분이 있어도 만듦 (개발용 — 배포용 꾸러미에는 쓰지 않는다)
  --dry-run         확인만 하고 파일을 쓰지 않음
  --root 폴더       Office Kit 폴더 (기본 이 스크립트의 상위, 시험용)

  담는 것: filex ${PINS.filex.tag} 소스, 묶음 만들기·덜어 내기 규칙·filex 패치·글꼴 생성기, NOTICE-OfficeKit.md,
           ONLYOFFICE ${PINS.onlyoffice.sourceTag} 소스 주소(sdkjs·web-apps·core·build_tools …)와 x2t 빌드 방법.`;

export function parseArgs(argv) {
  const o = { out: null, version: null, filexSrc: null, filexRepo: null, offline: false, allowMissing: false, dryRun: false, root: null };
  const need = (i, name) => {
    const v = argv[i + 1];
    if (!v || v.startsWith('--')) fail(`${name} 뒤에 값이 필요합니다`);
    return v;
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') o.out = need(i++, a);
    else if (a === '--version') o.version = need(i++, a);
    else if (a === '--filex-src') o.filexSrc = need(i++, a);
    else if (a === '--filex-repo') o.filexRepo = need(i++, a);
    else if (a === '--root') o.root = need(i++, a);
    else if (a === '--offline') o.offline = true;
    else if (a === '--allow-missing') o.allowMissing = true;
    else if (a === '--dry-run') o.dryRun = true;
    else if (a === '--help' || a === '-h') o.help = true;
    else fail(`모르는 옵션: ${a} (--help로 목록)`);
  }
  if (o.version && !/^\d+\.\d+\.\d+([-+][0-9A-Za-z.-]+)?$/.test(o.version)) fail(`버전 모양이 아닙니다: ${o.version}`);
  return o;
}

// ── git ──
// offline이면 부분 받기(partial clone)의 빠진 파일을 git이 몰래 받지 않게 (GIT_NO_LAZY_FETCH)
function git(args, { cwd, buffer = false, what, offline = false } = {}) {
  const env = offline ? { ...process.env, GIT_NO_LAZY_FETCH: '1' } : process.env;
  const r = spawnSync('git', ['-c', 'core.autocrlf=false', ...args], {
    cwd, env, encoding: buffer ? 'buffer' : 'utf8', windowsHide: true, maxBuffer: 256 * 1024 * 1024,
  });
  if (r.error?.code === 'ENOENT') fail('git을 찾을 수 없습니다 — Git을 설치하세요.');
  if (r.error) fail(`${what || 'git'} 실행 실패: ${r.error.message}`);
  return r;
}
const gitText = (args, cwd) => { const r = git(args, { cwd }); return r.status === 0 ? r.stdout.trim() : null; };
/** `git status --porcelain -z` → 경로들 ('XY 경로'; 이름 바꿈·복사는 다음 칸이 옛 이름 — 뺀다) */
export function porcelainPaths(z) {
  const out = [];
  const items = String(z).split('\0');
  for (let i = 0; i < items.length; i++) {
    const e = items[i];
    if (e.length < 4) continue;
    out.push(e.slice(3));
    if (e[0] === 'R' || e[0] === 'C') i++;
  }
  return out;
}

// 부분 받기 저장소는 없는 개체를 몰래 받으러 가므로 늘 GIT_NO_LAZY_FETCH로 묻는다
const hasCommit = (dir, commit) => isDir(dir) && git(['cat-file', '-e', `${commit}^{commit}`], { cwd: dir, offline: true }).status === 0;
// 그 커밋의 파일이 모두 이 PC에 있는지 (부분 받기면 일부가 빠져 있을 수 있다)
// (얕은 받기라 부모 커밋이 없어도 되게 그 커밋의 파일 목록만 본다)
function complete(dir, commit) {
  const ls = git(['ls-tree', '-r', '-z', commit], { cwd: dir, offline: true });
  if (ls.status !== 0) return false;
  const blobs = ls.stdout.split('\0').map((rec) => /^\d+ blob ([0-9a-f]+)\t/.exec(rec)?.[1]).filter(Boolean);
  const r = spawnSync('git', ['cat-file', '--batch-check'], {
    cwd: dir, env: { ...process.env, GIT_NO_LAZY_FETCH: '1' }, input: `${blobs.join('\n')}\n`, encoding: 'utf8', windowsHide: true, maxBuffer: 64 * 1024 * 1024,
  });
  return r.status === 0 && blobs.length > 0 && !/ missing$/m.test(r.stdout);
}

// 캐시 폴더 바로 아래 git 폴더들 (W1의 받은 소스 포함)
function cacheRepos(cache) {
  if (!isDir(cache)) return [];
  const out = [];
  for (const ent of readdirSync(cache, { withFileTypes: true })) {
    if (!ent.isDirectory()) continue;
    const d = join(cache, ent.name);
    if (existsSync(join(d, '.git')) || (isFile(join(d, 'HEAD')) && isDir(join(d, 'objects')))) out.push(d);
  }
  return out.sort();
}

function cacheDir() {
  return resolve(process.env.OFFICEKIT_OFFICE_EDITOR_CACHE || join(DEFAULT_ROOT, 'tmp', 'office-editor-build'));
}

/** filex 고정 커밋이 있는 git 폴더를 찾거나 받는다 → { dir, how } */
export function locateFilex({ filexSrc, filexRepo, offline } = {}) {
  const commit = PINS.filex.commit;
  const given = filexSrc || process.env.OFFICEKIT_FILEX_SRC;
  if (given) {
    const dir = resolve(given);
    if (!hasCommit(dir, commit)) fail(`${dir}에 filex 커밋 ${commit}이 없습니다 (git 폴더인지, 그 커밋을 받았는지 확인하세요).`);
    if (offline && !complete(dir, commit)) fail(`${dir}은(는) 부분 받기라 커밋 ${commit.slice(0, 12)}의 파일 일부가 없습니다 — --offline 없이 실행하세요.`);
    return { dir, how: 'given' };
  }
  const cache = cacheDir();
  // 파일이 다 있는 것 먼저, 부분 받기는 인터넷이 될 때만 (빠진 파일을 git이 받는다)
  const found = cacheRepos(cache).filter((d) => hasCommit(d, commit));
  const full = found.find((d) => complete(d, commit));
  if (full) return { dir: full, how: 'cache' };
  if (found.length && !offline) return { dir: found[0], how: 'cache' };
  if (offline) fail(`filex 소스(커밋 ${commit.slice(0, 12)})를 이 PC에서 찾지 못했습니다 — --filex-src로 주거나 --offline 없이 실행하세요.`);

  // 그 커밋만 받는다 (GitHub은 닿을 수 있는 커밋을 SHA로 받게 해 준다)
  const dir = join(cache, 'filex-source');
  const part = `${dir}.partial`;
  rmSync(part, { recursive: true, force: true });
  mkdirSync(part, { recursive: true });
  const url = filexRepo || process.env.OFFICEKIT_FILEX_REPO || PINS.filex.git;
  log(`filex 소스를 받습니다 (커밋 ${commit.slice(0, 12)}만): ${url}`);
  const steps = [
    [['init', '-q'], 'git init'],
    [['remote', 'add', 'origin', url], 'git remote'],
    [['fetch', '-q', '--depth', '1', 'origin', commit], 'filex 소스 받기 (git fetch — 인터넷 연결 확인)'],
  ];
  for (const [args, what] of steps) {
    const r = git(args, { cwd: part, what });
    if (r.status !== 0) fail(`${what}에 실패했습니다: ${String(r.stderr).trim().split('\n').pop() || r.status}`);
  }
  if (!hasCommit(part, commit)) fail(`받은 filex 소스에 커밋 ${commit}이 없습니다.`);
  rmSync(dir, { recursive: true, force: true });
  renameSync(part, dir);
  return { dir, how: 'fetched' };
}

/** filex 고정 커밋의 추적 파일 전부 → [{ path, data, mode }] */
export async function filexFiles(dir, { offline = false } = {}) {
  const commit = PINS.filex.commit;
  const r = git(['archive', '--format=zip', commit], { cwd: dir, buffer: true, what: 'git archive', offline });
  if (r.status !== 0) fail(`filex 소스를 꺼내지 못했습니다 (git archive): ${String(r.stderr).trim()}`);
  const zip = await JSZip.loadAsync(r.stdout);
  // 커밋의 파일 목록: 모드·blob 해시 (zip의 권한 칸은 git·OS마다 달라 믿지 않는다)
  const tree = new Map();
  const ls = git(['ls-tree', '-r', '-z', commit], { cwd: dir, what: 'git ls-tree', offline });
  if (ls.status !== 0) fail(`filex 파일 목록을 읽지 못했습니다 (git ls-tree): ${String(ls.stderr).trim()}`);
  for (const rec of ls.stdout.split('\0').filter(Boolean)) {
    const m = /^(\d+) (\w+) ([0-9a-f]+)\t(.+)$/s.exec(rec);
    if (m && m[2] === 'blob') tree.set(m[4], { mode: parseInt(m[1], 8), blob: m[3] });
  }
  const out = [];
  for (const f of Object.values(zip.files)) {
    if (f.dir) continue;
    const t = tree.get(f.name);
    if (!t) fail(`filex 커밋에 없는 파일이 나왔습니다: ${f.name}`);
    const data = await f.async('nodebuffer');
    // 바이트가 커밋의 blob과 같은지 (줄바꿈 변환 등이 끼지 않았는지)
    const blob = createHash('sha1').update(`blob ${data.length}\0`).update(data).digest('hex');
    if (blob !== t.blob) fail(`filex 파일 바이트가 커밋과 다릅니다: ${f.name}`);
    out.push({ path: f.name, data, mode: t.mode === 0o100755 ? 0o100755 : 0o100644 });
  }
  out.sort((a, b) => (a.path < b.path ? -1 : 1));
  if (tree.size !== out.length) fail(`filex 소스 파일 수가 맞지 않습니다 (커밋 ${tree.size}개, 꺼낸 것 ${out.length}개).`);
  return out;
}

/** filex upstream/onlyoffice.json이 PINS와 같은지 → 그 내용 */
export function checkFilexPins(files) {
  const f = files.find((x) => x.path === 'upstream/onlyoffice.json');
  if (!f) fail('filex 소스에 upstream/onlyoffice.json이 없습니다 — 고정 커밋이 맞는지 확인하세요.');
  const j = JSON.parse(f.data.toString('utf8'));
  const P = PINS.onlyoffice;
  const want = [
    ['build', j.build, P.build], ['source_tag', j.source_tag, P.sourceTag], ['digest', j.digest, P.digest], ['tag', j.tag, P.imageTag],
    ['x2t.release', j.x2t?.release, PINS.x2t.release],
    ['x2t.files.x2t.js', j.x2t?.files?.['x2t.js'], PINS.x2t.files['x2t.js']],
    ['x2t.files.x2t.wasm', j.x2t?.files?.['x2t.wasm'], PINS.x2t.files['x2t.wasm']],
    ['x2t core commit', j.x2t?.build?.sources?.core?.commit, P.repos.find((r) => r.name === 'core').commit],
    ['x2t build_tools commit', j.x2t?.build?.sources?.build_tools?.commit, P.repos.find((r) => r.name === 'build_tools').commit],
    ['x2t recipe commit', j.x2t?.build?.recipe?.commit, PINS.x2t.recipe.commit],
    ['x2t changes commit', j.x2t?.build?.recipe?.changes_commit, PINS.x2t.recipe.changesCommit],
  ];
  const bad = want.filter(([, got, exp]) => got !== exp);
  if (bad.length) {
    fail(`filex upstream/onlyoffice.json이 고정 값과 다릅니다:\n${bad.map(([k, got, exp]) => `  ${k}: ${got} (기대 ${exp})`).join('\n')}\n` +
      'scripts/office-editor/pins.mjs와 make-office-source.mjs의 소스 주소를 함께 올리세요.');
  }
  const pkg = files.find((x) => x.path === 'package.json');
  const ver = pkg && JSON.parse(pkg.data.toString('utf8')).version;
  if (`v${ver}` !== PINS.filex.tag) fail(`filex package.json 버전(${ver})이 태그 ${PINS.filex.tag}와 다릅니다.`);
  return j;
}

/** docs/office-editor-legal-ko.md가 고정 값과 맞는지 → 틀린 것 목록 */
export function checkLegalDoc(text) {
  const P = PINS.onlyoffice;
  const need = [
    ['{{appVersion}}', '설치된 판의 소스 주소 자리표시'],
    [P.build, 'ONLYOFFICE 빌드'],
    [`/tree/${P.sourceTag}`, 'ONLYOFFICE 소스 태그 주소'],
    [`${PINS.filex.repo}/tree/${PINS.filex.tag}`, 'filex 태그 주소'],
    ['office-editor-source-{{appVersion}}.zip', '소스 꾸러미 이름'],
    ['local:editor/LICENSE.txt', 'AGPL 전문·추가 조항 연결'],
    ['local:NOTICE-OfficeKit.md', '수정 고지 연결'],
    ['local:THIRD-PARTY-NATIVE.txt', '기계어(WebAssembly) 제3자 고지 연결'],
    ['Ascensio System SIA', '원 개발사'],
    ['수정된 판일 수 있습니다', '수정판 안내'],
    ['https://www.onlyoffice.com/trademark-policy', '상표 정책'],
  ];
  for (const r of P.repos.filter((x) => x.main)) need.push([`${r.url}/tree/${r.tag}`, `${r.name} 소스 주소`]);
  return need.filter(([s]) => !text.includes(s)).map(([s, what]) => `${what} (${s})`);
}

// ── Office Kit 쪽 파일 ──
function walk(dir, base = dir) {
  const out = [];
  for (const ent of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, ent.name);
    const rel = posix(relative(base, full));
    if (SKIP.test(rel)) continue;
    if (ent.isDirectory()) out.push(...walk(full, base));
    else if (ent.isFile()) out.push(rel);
  }
  return out.sort();
}

export function collectParts(root) {
  const files = [];
  const missing = [];
  const parts = [];
  for (const p of PARTS) {
    const src = join(root, p.src);
    let list = [];
    if (p.dir) {
      if (isDir(src)) list = walk(src).map((rel) => ({ src: join(src, rel), dest: `${p.dest}/${rel}` }));
    } else if (isFile(src)) list = [{ src, dest: p.dest }];
    if (!list.length) {
      parts.push({ id: p.id, what: p.what, src: p.src, dest: p.dest, files: 0, missing: true, required: p.required });
      if (p.required) missing.push({ id: p.id, src: p.src, what: p.what });
      continue;
    }
    let bytes = 0;
    for (const f of list) {
      const st = statSync(f.src);
      if (st.size > MAX_FILE) fail(`${posix(relative(root, f.src))}가 너무 큽니다 (${kb(st.size)}) — 소스 꾸러미에 넣을 파일이 맞는지 확인하세요.`);
      const data = readFileSync(f.src);
      bytes += data.length;
      files.push({ path: f.dest, data, from: posix(relative(root, f.src)) });
    }
    parts.push({ id: p.id, what: p.what, src: p.src, dest: p.dest, files: list.length, bytes, missing: false, required: p.required });
  }
  return { files, missing, parts };
}

function officeKitState(root, version) {
  // root가 저장소 맨 위일 때만 커밋을 적는다 (다른 저장소 안 폴더면 그 저장소의 커밋이 아님)
  const top = gitText(['rev-parse', '--show-toplevel'], root);
  const same = !!top && resolve(top).toLowerCase() === resolve(root).toLowerCase();
  const commit = same ? gitText(['rev-parse', 'HEAD'], root) : null;
  const paths = PARTS.filter((p) => !p.src.startsWith('vendor/')).map((p) => p.src);
  // porcelain 줄은 'XY 경로' — 앞 칸이 공백일 수 있으므로 다듬지(trim) 않고 -z로 나눈다
  const st = commit ? git(['status', '--porcelain', '-z', '--', ...paths], { cwd: root }) : null;
  const status = st && st.status === 0 ? st.stdout : null;
  return {
    repo: PINS.officeKit.repo,
    version,
    tag: `v${version}`,
    commit,
    dirty: status == null ? null : status.length > 0,
    dirtyFiles: status ? porcelainPaths(status) : [],
  };
}

function sourceEpoch(root) {
  const env = process.env.SOURCE_DATE_EPOCH;
  if (env && /^\d+$/.test(env)) return Number(env);
  const ct = gitText(['log', '-1', '--format=%ct'], root);
  if (ct && /^\d+$/.test(ct)) return Number(ct);
  return Math.floor(Date.now() / 1000);
}

// ── 만들어 넣는 글 ──
const ymd = (epoch) => new Date(epoch * 1000).toISOString().slice(0, 10);

/** 묶음이 없을 때(--allow-missing) 대신 넣는 수정 고지. 묶음의 진짜 고지는 scripts/office-editor/notice.mjs가 쓴다 */
export function renderNotice({ date, version = null, removed = null, patches = null } = {}) {
  const P = PINS.onlyoffice;
  const d = date || new Date().toISOString().slice(0, 10);
  const rm = Array.isArray(removed) && removed.length
    ? removed.map((r) => `  - ${r.what}${r.files != null ? ` (${r.files}개)` : ''}`).join('\n')
    : [
      '  - 편집기 언어 파일 (한국어·영어 말고)',
      '  - 수식 함수 이름 언어 파일 (한국어·영어 말고)',
      '  - 휴대폰·embed·forms 앱',
      '  - 묶음 글꼴 `editor/fonts`와 그 라이선스 (편집용 글꼴 파일은 넣지 않는다 — 다만 x2t.wasm 안의 URW++ PDF 기본 14 글꼴은 남는다, `THIRD-PARTY-NATIVE.txt`)',
    ].join('\n');
  const pt = Array.isArray(patches) && patches.length
    ? patches.map((p) => `  - \`${p}\``).join('\n')
    : '  - `patches/office-editor/filex/*.patch` (소스 꾸러미 `office-kit/patches/`)';
  return `# 수정 고지 — Office Kit의 Office 편집기 묶음

이 폴더(\`vendor/office-editor\`)는 **수정된 판**이다. 원본은 ${P.vendor}가 개발한 **ONLYOFFICE Docs**
(${P.product} ${P.version}, 빌드 ${P.build}, 소스 태그 \`${P.sourceTag}\`)이고,
BRF Tech의 filex-office-editor ${PINS.filex.tag}가 그 편집기 파일을 묶은 것(\`ui.zip\`)을 Office Kit이 다시 고쳐 쓴다.
ONLYOFFICE 부분은 GNU AGPL 3판과 그 추가 조항(\`editor/LICENSE.txt\`)을 따른다.

## ${d} — Office Kit${version ? ` ${version}` : ''} (toxh13)

1. **덜어 냄** — 파일을 통째로 뺐다. 남은 파일은 한 바이트도 고치지 않았다 (머리글·저작권 고지 그대로).
${rm}
2. **filex 앱 패치** — filex 소스(커밋 \`${PINS.filex.commit.slice(0, 12)}\`)에 아래 패치를 적용해 앱 페이지를 다시 빌드했다 (패치가 있을 때).
${pt}
3. **실행 중 바꿔 내주는 것** — 디스크의 파일은 그대로 두고, Office Kit의 로컬 서버가 아래 주소에 다른 내용을 보낸다.
  - \`editor/sdkjs/common/AllFonts.js\`, \`editor/sdkjs/common/Images/fonts_thumbnail*.png\`: 이 PC에 설치된 글꼴로 만든 목록·미리보기
  - \`editor/fonts/<번호>\`: 이 PC의 글꼴 파일 (ONLYOFFICE가 기대하는 대로 앞 32바이트만 XOR, 복사하지 않음)
  - HTML 응답의 보안 머리글(CSP)

filex-office-editor가 ONLYOFFICE 파일에 한 수정(BRF Tech, 2026-10-08~10)은 \`NOTICE\`와 \`editor/filex/CHANGES.txt\`에 있다.

## Notice (English)

This folder is a **modified version** of ONLYOFFICE Docs by ${P.vendor} (Docs ${P.version}, build ${P.build},
source tag ${P.sourceTag}), as bundled by BRF Tech's filex-office-editor ${PINS.filex.tag}. Modified by Office Kit (toxh13)
on ${d}: whole files left out (interface and formula languages other than Korean and English; the phone, embed and
forms apps; the bundled fonts), the filex app rebuilt with Office Kit's patches when there are any, and, at run time,
AllFonts.js, the font thumbnails and the font files served from this PC's installed fonts. No remaining file is edited.
ONLYOFFICE® is a registered trademark of ${P.vendor}; Office Kit is not affiliated with or endorsed by it.
`;
}

function upstreamJson(filexJson, office, version) {
  const P = PINS.onlyoffice;
  return {
    about: `Office Kit ${version}의 Office 편집기 묶음(vendor/office-editor)을 만든 소스의 주소. 커밋은 git ls-remote로 확인한 값.`,
    officeKit: office,
    filex: { ...PINS.filex, note: 'commit = 태그가 가리키는 커밋 (릴리스 ui.zip을 만든 소스). mainCommit은 그 다음 README.md만 고친 커밋.' },
    onlyoffice: { ...P },
    x2t: { ...PINS.x2t, build: filexJson?.x2t?.build ?? null, howTo: 'filex-office-editor/scripts/x2t/build.sh (docker), filex-office-editor/README.md "x2t, the converter"' },
    mirrors: { status: 'done', url: UPSTREAM_MIRROR, note: '위 커밋 그대로의 소스 압축본 + SHA256SUMS — docs/LICENSES-OFFICE-EDITOR.md 9절' },
  };
}

function renderUpstream(filexJson, version) {
  const P = PINS.onlyoffice;
  const b = filexJson?.x2t?.build || {};
  const src = b.sources || {};
  const row = (r) => `| ${r.name} | ${r.url} | \`${r.tag}\` | \`${r.commit}\` | ${r.role} |`;
  const lines = [
    `# 위쪽(upstream) 소스 — Office Kit ${version}의 Office 편집기`,
    '',
    '이 꾸러미에 든 것 말고, 묶음을 이루는 ONLYOFFICE·CryptPad 소스는 아래 주소의 **정확한 태그·커밋**이다.',
    '커밋은 2026-10-18에 `git ls-remote`로 확인했다 (태그가 옮겨지면 커밋을 믿는다).',
    '',
    `## ONLYOFFICE Docs ${P.version} (빌드 ${P.build}) — ${P.vendor}`,
    '',
    `- 편집기 파일은 공식 Docker 이미지 \`${P.image}:${P.imageTag}\` (\`${P.digest}\`)에서 꺼낸 것이다 (filex \`scripts/extract-editor.sh\`).`,
    `- 릴리스: ${P.release}`,
    `- 라이선스: ${P.license}. 전문은 \`licenses/ONLYOFFICE-LICENSE.txt\`.`,
    '',
    '묶음에 직접 든 것:',
    '',
    '| 저장소 | 주소 | 태그 | 커밋 | 무엇 |',
    '|---|---|---|---|---|',
    ...P.repos.filter((r) => r.main).map(row),
    '',
    '같은 태그의 관련 저장소 (빌드 도구가 함께 받거나, 원본 규칙을 옮겨 온 것):',
    '',
    '| 저장소 | 주소 | 태그 | 커밋 | 무엇 |',
    '|---|---|---|---|---|',
    ...P.repos.filter((r) => !r.main).map(row),
    '',
    `sdkjs·web-apps를 소스에서 묶는 방법은 build_tools \`${P.sourceTag}\`의 안내를 따른다 (Closure Compiler·Grunt로 \`sdk-all*.js\`와 앱 묶음을 만든다).`,
    '',
    `## filex-office-editor ${PINS.filex.tag} — BRF Tech`,
    '',
    `- 라이선스: ${PINS.filex.license}`,
    `- 저장소: ${PINS.filex.repo}`,
    `- 이 꾸러미의 \`filex-office-editor/\` = 커밋 \`${PINS.filex.commit}\` 그대로 (\`git archive\`).`,
    `  태그 \`${PINS.filex.tag}\`가 가리키는 커밋이다 (릴리스 \`ui.zip\`을 만든 소스). main의 \`${PINS.filex.mainCommit.slice(0, 12)}\`은 그 다음 README.md만 고친 커밋이다.`,
    `- 릴리스 \`ui.zip\`: ${PINS.filex.uiZip.url} — ${PINS.filex.uiZip.bytes.toLocaleString('en-US')}바이트, SHA-256 \`${PINS.filex.uiZip.sha256}\``,
    '- 같은 `ui.zip`을 다시 만드는 법: `filex-office-editor/README.md` "Building a release" (Docker만 있으면 됨).',
    '- 묶음에 든 npm 꾸러미(예: `@brftech/filex-app-ui`, MIT)는 `filex-office-editor/package-lock.json`이 버전·무결성 값을 고정한다.',
    '',
    `## x2t (변환기) ${PINS.x2t.release}`,
    '',
    `- 라이선스: ${PINS.x2t.license}`,
    `- 파일: \`x2t.js\` SHA-256 \`${PINS.x2t.files['x2t.js']}\`, \`x2t.wasm\` SHA-256 \`${PINS.x2t.files['x2t.wasm']}\``,
    '- 빌드 방법: `filex-office-editor/scripts/x2t/` (`build.sh`, `Dockerfile`, `steps.sh`, `main1.cpp`, `pre-js.js`, `patches/01~05`).',
    '  `bash filex-office-editor/scripts/x2t/build.sh` → 두 파일의 SHA-256이 위 값과 같아야 한다 (약 20분, 작업 폴더 약 4.5GB).',
    `- 빌드 방법의 출처: CryptPad ${PINS.x2t.recipe.repo} \`${PINS.x2t.recipe.release}\` (커밋 \`${PINS.x2t.recipe.commit}\`),`,
    `  core 변경은 \`${PINS.x2t.recipe.changes}\` (커밋 \`${PINS.x2t.recipe.changesCommit}\`).`,
  ];
  if (b.toolchain) lines.push(`- 도구: \`${b.toolchain.image}\` (\`${b.toolchain.digest}\`), Ubuntu 스냅샷 \`${b.toolchain.apt_snapshot}\``);
  const extra = [];
  for (const [name, s] of Object.entries(src)) {
    if (name === 'core' || name === 'build_tools') continue;
    extra.push(`  - ${name}: ${s.repository || s.url}${s.tag ? ` \`${s.tag}\`` : ''}${s.commit ? ` 커밋 \`${s.commit}\`` : ''}${s.sha256 ? ` SHA-256 \`${s.sha256}\`` : ''}`);
  }
  if (extra.length) lines.push('- 함께 쓰는 소스:', ...extra);
  lines.push(
    '- core의 받기 스크립트가 그 밖의 제3자 코드(harfbuzz 등)를 core가 적은 커밋으로 받는다 (core `3DPARTY.md`).',
    '',
    '## 거울 (위쪽 소스 사본)',
    '',
    `위쪽 저장소의 태그가 지워지거나 옮겨져도 소스를 받을 수 있게, 아래 모두를 커밋 그대로 압축해 ${UPSTREAM_MIRROR} 에 올려 두었다 (SHA256SUMS 포함).`,
    '',
    ...P.repos.map((r) => `- [x] ${r.name} \`${r.tag}\` (\`${r.commit.slice(0, 12)}\`)`),
    `- [x] filex-office-editor \`${PINS.filex.tag}\`·\`${PINS.filex.commit.slice(0, 12)}\` (+ 릴리스 \`ui.zip\`, \`x2t.zip\`)`,
    `- [x] onlyoffice-x2t-wasm (CryptPad) \`${PINS.x2t.recipe.release}\`·\`${PINS.x2t.recipe.changes}\``,
    ...Object.entries(src).filter(([name]) => name !== 'core' && name !== 'build_tools')
      .map(([name, s]) => `- [x] x2t 소스 ${name}: ${s.repository || s.url}${s.tag ? ` \`${s.tag}\`` : ''}${s.commit ? ` (\`${s.commit.slice(0, 12)}\`)` : ''}${s.url ? ' (내려받은 파일 그대로)' : ''}`),
    '- [x] core가 받는 제3자 코드 (core `Common/3dParty/*/fetch.py`·`make.py`의 커밋): gumbo-parser, katana-parser, md4c, glm, mdds, librevenge, libodfgen, libetonyek, harfbuzz, brotli — 목록과 커밋은 `licenses/THIRD-PARTY-NATIVE.txt`',
    '- [x] emscripten 4.0.11의 ICU 포트 (`icu4c-68_2-src.zip`, unicode-org/icu `release-68-2`)',
    '- [x] filex 앱 UI 패키지 `@brftech/filex-app-ui` (npm, filex `package-lock.json`의 판)',
    '',
    '```bash',
    '# 받은 뒤 확인: 사본 Release의 SHA256SUMS와 맞춰 본다',
    'sha256sum -c SHA256SUMS',
    '```',
    '',
  );
  return lines.join('\n');
}

function renderReadme({ version, parts, missing }) {
  const P = PINS.onlyoffice;
  const partRows = parts.map((p) => `| \`${p.src}\` | \`${p.dest}\` | ${p.what} | ${p.missing ? (p.required ? '**빠짐**' : '없음') : `${p.files}개`} |`);
  return `# Office Kit ${version} — Office 편집기 소스 꾸러미

Office Kit ${version} 설치 파일의 Office 편집 창(문서 편집 · 표 편집 · 슬라이드 편집)이 쓰는 편집기 묶음
\`resources/app/vendor/office-editor/\`는 **GNU Affero General Public License 3판(AGPL-3.0)** 을 따른다.
이 꾸러미는 그 묶음의 **대응 소스**(Corresponding Source, AGPL 1조·6조)를 모은 것이다.

- 원 개발사: **${P.vendor}** (ONLYOFFICE Docs ${P.version}, 빌드 ${P.build}). 이 판은 **수정된 판**이다 → \`NOTICE-OfficeKit.md\`.
- 다리 코드·앱 페이지: BRF Tech filex-office-editor ${PINS.filex.tag} — ${PINS.filex.license}.
- Office Kit 나머지(창·호스트 페이지·문서 엔진)는 MIT다: ${PINS.officeKit.repo}/tree/v${version}

## 들어 있는 것

| 꾸러미 안 | 무엇 |
|---|---|
| \`filex-office-editor/\` | filex 소스, 커밋 \`${PINS.filex.commit}\` 그대로 (\`git archive\`). 앱 페이지·다리 코드·x2t 빌드 방법(\`scripts/x2t/\`)·편집기 파일 목록(\`upstream/editor.lock.json\`) |
| \`office-kit/\` | Office Kit 저장소의 같은 경로 파일 (아래 표) |
| \`bundle/VERSION.json\` | 만든 묶음의 고정 값과 파일마다 크기·SHA-256 |
| \`licenses/\` | ONLYOFFICE \`LICENSE.txt\`(AGPL 3판 + 추가 조항), \`3rd-Party.txt\`, \`license/\` |
| \`NOTICE-OfficeKit.md\` | 날짜 붙은 수정 고지 |
| \`UPSTREAM.md\`, \`upstream.json\` | ONLYOFFICE sdkjs·web-apps·core·build_tools \`${P.sourceTag}\`의 주소·커밋, x2t 빌드 방법 |
| \`MANIFEST.json\` | 이 꾸러미의 파일마다 크기·SHA-256 |

| Office Kit 경로 | 꾸러미 안 | 무엇 | 파일 |
|---|---|---|---|
${partRows.join('\n')}
${missing.length ? `
> **주의 — 빠진 부분이 있는 개발용 꾸러미다.** 배포에 쓰지 않는다: ${missing.map((m) => `\`${m.src}\``).join(', ')}
` : ''}
## 다시 만드는 법

1. Office Kit 저장소에서 태그 \`v${version}\`의 소스를 받는다 (\`office-kit/\`의 파일과 같다). Windows, Node.js 24, Git.
2. \`npm ci\` → \`npm run office-editor:build\`
   - filex 릴리스 \`ui.zip\`을 받아 크기·SHA-256(\`${PINS.filex.uiZip.sha256.slice(0, 16)}…\`)을 확인하고, 편집기 파일마다 \`upstream/editor.lock.json\`과 견준다.
   - 패치가 있으면 filex 소스(이 꾸러미의 \`filex-office-editor/\`와 같은 커밋)에 \`patches/office-editor/filex/*.patch\`를 적용해 앱 페이지를 다시 빌드한다.
   - 덜어 내기 규칙(\`scripts/office-editor/\`)대로 파일을 통째로 빼고 \`vendor/office-editor/\`에 쓴다.
3. 편집기 파일(web-apps·sdkjs)을 소스에서: \`UPSTREAM.md\`의 ONLYOFFICE 저장소 \`${P.sourceTag}\`와 build_tools.
   같은 파일은 공식 Docker 이미지에서도 꺼낼 수 있다: \`bash filex-office-editor/scripts/extract-editor.sh\` (digest 고정, 파일마다 SHA-256 확인).
4. x2t: \`bash filex-office-editor/scripts/x2t/build.sh\` (Docker) → \`x2t.js\`·\`x2t.wasm\`의 SHA-256이 \`UPSTREAM.md\`의 값과 같다.
5. 글꼴 목록(\`AllFonts.js\`)은 설치 파일에 없다. 편집 창을 처음 열 때 \`office-kit/src/app/office-editor/fonts/\`가 그 PC의 글꼴로 만든다.

## 라이선스

- ONLYOFFICE 코드: Copyright (C) ${P.vendor}. AGPL 3판 + 7조 추가 조항 1~5 (고지 유지, 수정 고지, 화면 안 법적 고지, 상표권 없음, 코드 아닌 것은 CC BY-SA 4.0).
- filex-office-editor: ${PINS.filex.copyright}. ${PINS.filex.license}.
- x2t: ${PINS.x2t.license}.
- Office Kit의 filex 패치: AGPL-3.0-or-later, 글꼴 생성기: AGPL-3.0-only + ONLYOFFICE 추가 조항 (파일 머리 \`SPDX-License-Identifier\`, 전문은 \`office-kit/LICENSES/\`). 나머지 Office Kit 파일은 MIT.
- x2t·sdkjs WebAssembly 안 제3자 코드(Boost, ICU, FreeType, libpng, libjpeg, Xpdf 등)와 x2t.wasm 안 URW++ 글꼴: \`licenses/THIRD-PARTY-NATIVE.txt\`.
- 상표: ONLYOFFICE®·ONLYOFFICE Docs™는 ${P.vendor}의 상표다. Office Kit은 ${P.vendor}와 관계가 없고 보증·후원을 받지 않는다.
- 자세한 설명: \`office-kit/docs/LICENSES-OFFICE-EDITOR.md\`

## English

This archive is the Corresponding Source (GNU AGPL v3, sections 1 and 6) of the office editor bundle
(\`resources/app/vendor/office-editor/\`) shipped in Office Kit ${version}. The bundle is a modified version of
ONLYOFFICE Docs ${P.version} (build ${P.build}) by ${P.vendor}, as packaged by BRF Tech's filex-office-editor ${PINS.filex.tag}
(source commit \`${PINS.filex.commit}\`, in \`filex-office-editor/\`). Office Kit's build script, trimming rules,
patches and font list generator are in \`office-kit/\`; the exact ONLYOFFICE source coordinates (sdkjs, web-apps, core,
build_tools at \`${P.sourceTag}\`) and the x2t build recipe are in \`UPSTREAM.md\`. ONLYOFFICE® is a registered trademark
of ${P.vendor}; Office Kit is not affiliated with or endorsed by it.
`;
}

// ── 꾸러미 ──
async function buildZip(entries, epoch) {
  const zip = new JSZip();
  const date = new Date(epoch * 1000);
  const seen = new Set();
  for (const e of [...entries].sort((a, b) => (a.path < b.path ? -1 : 1))) {
    if (seen.has(e.path)) fail(`꾸러미 안 경로가 겹칩니다: ${e.path}`);
    seen.add(e.path);
    const opt = { date, createFolders: false, binary: true };
    opt.unixPermissions = e.mode || 0o100644;
    zip.file(e.path, e.data, opt);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 }, platform: 'UNIX' });
}

/** 소스 꾸러미 만들기 → { zipPath, sha256, bytes, manifest } (dryRun이면 zipPath 없음) */
export async function makeOfficeSource(opts = {}) {
  const root = resolve(opts.root || DEFAULT_ROOT);
  const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
  const version = opts.version || pkg.version;
  const name = `office-editor-source-${version}`;
  const outDir = opts.out ? resolve(opts.out) : join(root, 'dist');

  // 1. Office Kit 쪽
  const pinErr = checkPins();
  if (pinErr.length) fail(`scripts/office-editor/pins.mjs와 소스 주소(make-office-source.mjs)가 맞지 않습니다:\n  ${pinErr.join('\n  ')}\n함께 올리세요.`);

  const kit = collectParts(root);
  const legal = kit.files.find((f) => f.path === 'office-kit/docs/office-editor-legal-ko.md');
  if (legal) {
    const bad = checkLegalDoc(legal.data.toString('utf8'));
    if (bad.length) fail(`docs/office-editor-legal-ko.md가 고정 값과 맞지 않습니다:\n  ${bad.join('\n  ')}`);
  }
  if (kit.missing.length && !opts.allowMissing) {
    fail(`소스 꾸러미에 넣을 부분이 없습니다:\n${kit.missing.map((m) => `  ${m.src} — ${m.what}`).join('\n')}\n` +
      '먼저 npm run office-editor:build로 묶음을 만드세요 (개발용으로만 --allow-missing).');
  }

  // 2. filex 소스
  const { dir, how } = locateFilex(opts);
  const filex = await filexFiles(dir, opts);
  const filexJson = checkFilexPins(filex);
  const filexHow = { given: '--filex-src로 준 git 폴더', cache: '빌드 캐시의 git 폴더', fetched: '고정 커밋만 받음' }[how];
  log(`filex ${PINS.filex.tag} 소스: 파일 ${filex.length}개 (${filexHow}: ${dir})`);

  // 3. 만들어 넣는 글
  const epoch = sourceEpoch(root);
  const office = officeKitState(root, version);
  const generated = [];
  if (!kit.files.some((f) => f.path === 'NOTICE-OfficeKit.md')) {
    generated.push({ path: 'NOTICE-OfficeKit.md', data: Buffer.from(renderNotice({ date: ymd(epoch), version })), from: '(만듦 — 묶음 없음)' });
  }
  generated.push({ path: 'UPSTREAM.md', data: Buffer.from(renderUpstream(filexJson, version)), from: '(만듦)' });
  generated.push({ path: 'upstream.json', data: Buffer.from(`${JSON.stringify(upstreamJson(filexJson, office, version), null, 2)}\n`), from: '(만듦)' });
  generated.push({ path: 'README.md', data: Buffer.from(renderReadme({ version, parts: kit.parts, missing: kit.missing })), from: '(만듦)' });

  const entries = [
    ...filex.map((f) => ({ path: `filex-office-editor/${f.path}`, data: f.data, mode: f.mode, from: `filex@${PINS.filex.commit.slice(0, 12)}` })),
    ...kit.files,
    ...generated,
  ];
  const manifest = {
    name,
    version,
    epoch,
    date: new Date(epoch * 1000).toISOString(),
    officeKit: office,
    filex: { repo: PINS.filex.repo, tag: PINS.filex.tag, commit: PINS.filex.commit, files: filex.length, how },
    onlyoffice: { version: PINS.onlyoffice.version, build: PINS.onlyoffice.build, sourceTag: PINS.onlyoffice.sourceTag },
    parts: kit.parts,
    missing: kit.missing,
    complete: kit.missing.length === 0,
    files: entries.map((e) => ({ path: e.path, bytes: e.data.length, sha256: sha256(e.data) })).sort((a, b) => (a.path < b.path ? -1 : 1)),
  };
  entries.push({ path: 'MANIFEST.json', data: Buffer.from(`${JSON.stringify(manifest, null, 2)}\n`) });

  const zipBuf = await buildZip(entries.map((e) => ({ ...e, path: `${name}/${e.path}` })), epoch);
  const digest = sha256(zipBuf);
  if (opts.dryRun) return { zipPath: null, sha256: digest, bytes: zipBuf.length, manifest };

  mkdirSync(outDir, { recursive: true });
  const zipPath = join(outDir, `${name}.zip`);
  const tmp = `${zipPath}.tmp-${process.pid}`;
  writeFileSync(tmp, zipBuf);
  renameSync(tmp, zipPath);
  writeFileSync(`${zipPath}.sha256`, `${digest}  ${name}.zip\n`);
  return { zipPath, sha256: digest, bytes: zipBuf.length, manifest };
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) { console.log(HELP); return; }
  const r = await makeOfficeSource(opts);
  const m = r.manifest;
  for (const p of m.parts) log(`${p.missing ? (p.required ? '빠짐' : '없음') : `${String(p.files).padStart(3)}개`}  ${p.src} — ${p.what}`);
  if (m.officeKit.dirty) log(`주의: 커밋하지 않은 파일이 들어갔습니다 (${m.officeKit.dirtyFiles.length}개) — 배포용은 깨끗한 태그에서 만드세요.`);
  if (!m.complete) log(`주의: 빠진 부분이 있는 개발용 꾸러미입니다 (${m.missing.map((x) => x.src).join(', ')}).`);
  const what = `${m.name}.zip — 파일 ${m.files.length + 1}개, ${kb(r.bytes)}, SHA-256 ${r.sha256}`;
  log(r.zipPath ? `만들었습니다: ${relative(process.cwd(), r.zipPath) || r.zipPath} (${what})` : `확인만 했습니다 (--dry-run): ${what}`);
}

const isMain = process.argv[1] && pathToFileURL(resolve(process.argv[1])).href.toLowerCase() === import.meta.url.toLowerCase();
if (isMain) {
  main().catch((e) => {
    console.error(`[office-source] 오류: ${e instanceof SourceError ? e.message : e.stack || e.message}`);
    process.exitCode = 1;
  });
}
