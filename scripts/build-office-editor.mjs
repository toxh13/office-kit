#!/usr/bin/env node
// build-office-editor.mjs — 앱 안 Office 편집기(문서·표·슬라이드) 묶음을 만들어 vendor/office-editor에 넣는다.
//
//   node scripts/build-office-editor.mjs [--force] [--rebuild-app] [--patches 폴더] [--out 폴더] [--offline]
//   node scripts/build-office-editor.mjs --verify      (npm run office-editor:build)
//
// 하는 일 (docs/OFFICE-EDITOR-DESIGN.md 4장)
//   1. 받기       filex-office-editor v0.3.1 릴리스의 ui.zip을 캐시 폴더로 (이어받기, 3번 다시)
//   2. 확인       크기·SHA-256이 고정 값과 같은지 (다르면 파일을 지우고 멈춤)
//   3. 풀기       jszip으로 캐시에 풂 (<캐시>/ui-v0.3.1 — 덜어 내기 전 전체)
//   4. 파일별 확인 editor/ 파일마다 filex upstream/editor.lock.json(고정 커밋, 그 파일의 SHA-256도 고정)과 비교.
//                 filex 다리(socket.io.min.js)·CHANGES.txt 덧붙임·빈 목록 두 개만 다를 수 있다. x2t 두 파일은 고정 SHA-256
//   5. (패치가 있을 때) filex 소스를 고정 커밋으로 받아 patches/office-editor/filex/*.patch를 얹고
//                 npm ci --ignore-scripts → node scripts/build-app.mjs --no-zip → 앱 페이지·filex/*·다리만 가져옴
//   6. 덜어 내기  scripts/office-editor/rules.mjs (파일을 통째로 빼기만, 남는 파일은 그대로)
//   7. 쓰기       vendor/.office-editor.tmp에 다 쓴 뒤 vendor/office-editor로 바꿔 넣음 (+ VERSION.json, NOTICE-OfficeKit.md)
//   8. 건너뛰기   VERSION.json의 고정 값·규칙 해시·패치가 같고 파일이 그대로면 아무것도 안 함 (--force: 다시)
//
// 캐시: tmp/office-editor-build (OFFICEKIT_OFFICE_EDITOR_CACHE로 바꿈). 고정 값: scripts/office-editor/pins.mjs
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, lstatSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { basename, dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

import { FetchError, fetchPinned } from './office-editor/fetch.mjs';
import {
  StepError, applyPatches, buildApp, ensureFilexSource, installDeps, listPatches, patchesSha256, sparseDirsFor,
} from './office-editor/filex-source.mjs';
import { ADDED, BRIDGE, CHANGES, EDITOR_PREFIX, X2T_FILES, makeLockEditor, parseLock, sha256, verifyUi } from './office-editor/lock.mjs';
import { renderNotice } from './office-editor/notice.mjs';
import { EDITOR_LOCK, FILEX, ONLYOFFICE, PINS, SOCKET_IO_ORIGINAL, UI_ZIP, X2T, cachePaths, defaultCache, pinsSha256 } from './office-editor/pins.mjs';
import { DROP_RULES, NOTICE_DATED, partition, rulesSha256 } from './office-editor/rules.mjs';
import { THIRD_PARTY_BUNDLE_PATH, THIRD_PARTY_FILE, checkThirdPartyFile } from './office-editor/third-party.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(join(ROOT, 'package.json'));
const CACHE = defaultCache(ROOT);
const PATHS = cachePaths(CACHE);
const DEFAULT_OUT = join(ROOT, 'vendor', 'office-editor');
const DEFAULT_PATCHES = join(ROOT, 'patches', 'office-editor', 'filex');
// 설계에서 잰 값 (9장) — 결과와 나란히 보여 준다
const DESIGN = { bytes: 213.7 * 1048576, files: 1530, gzipBytes: 64.1 * 1048576 };

const log = (s) => process.stdout.write(`[office-editor] ${s}\n`);
class BuildError extends Error {}
const fail = (msg) => { throw new BuildError(msg); };
const mb = (n) => `${(n / 1048576).toFixed(1)}MB`;
const num = (n) => n.toLocaleString('en-US');
const posix = (p) => p.split(sep).join('/');
const readJson = (p) => JSON.parse(readFileSync(p, 'utf8'));
const shown = (p) => (relative(ROOT, p).startsWith('..') ? p : relative(ROOT, p) || p);

const HELP = `Office 편집기 묶음 만들기 (${FILEX.name} ${FILEX.tag} · ONLYOFFICE Docs ${ONLYOFFICE.build} 바탕) → vendor/office-editor

  node scripts/build-office-editor.mjs [옵션]          (npm run office-editor:build)

  --force          결과가 이미 있어도 다시 만듦
  --rebuild-app    패치가 없어도 5단계(filex 앱 페이지 다시 빌드)를 함 — 릴리스와 같은 바이트인지 확인용
  --patches 폴더   filex 패치 폴더 (기본 patches/office-editor/filex, *.patch 이름 순)
  --out 폴더       결과 폴더 (기본 vendor/office-editor)
  --offline        인터넷을 쓰지 않음 (캐시에 받아 둔 것만)
  --measure        끝에 gzip-9 크기도 잼 (설치 파일 크기 어림, 30초 안팎)
  --verify         만들지 않고, 결과 폴더의 파일을 VERSION.json과 하나하나 맞춰 봄

  캐시 폴더: ${shown(CACHE)} (OFFICEKIT_OFFICE_EDITOR_CACHE로 바꿈)
    download/        ui.zip(${mb(UI_ZIP.bytes)})·lock 파일 — CI가 이 폴더를 캐시한다
    ui-${FILEX.tag}/       ui.zip을 그대로 푼 전체 (editor/fonts·AllFonts.js 기준 시험용)
    filex-src-*/     5단계 filex 소스 (패치가 있을 때만, git·npm 필요)`;

function parseArgs(argv) {
  const o = { force: false, rebuildApp: false, patches: DEFAULT_PATCHES, out: DEFAULT_OUT, offline: false, measure: false, verify: false, help: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const val = () => {
      const v = argv[++i];
      if (!v || v.startsWith('--')) fail(`${a} 다음에 폴더를 주세요.`);
      return resolve(v);
    };
    if (a === '--force' || a === '-f') o.force = true;
    else if (a === '--rebuild-app') o.rebuildApp = true;
    else if (a === '--patches') o.patches = val();
    else if (a === '--out') o.out = val();
    else if (a === '--offline') o.offline = true;
    else if (a === '--measure') o.measure = true;
    else if (a === '--verify') o.verify = true;
    else if (a === '--help' || a === '-h') o.help = true;
    else fail(`모르는 옵션: ${a} (--help로 옵션 보기)`);
  }
  return o;
}

/** 이 빌드 스크립트들의 해시 (줄바꿈 LF로 맞춤) */
function builderSha256() {
  const dir = join(ROOT, 'scripts', 'office-editor');
  const files = [join(ROOT, 'scripts', 'build-office-editor.mjs'),
    ...readdirSync(dir).filter((f) => f.endsWith('.mjs') && !f.endsWith('.test.mjs')).sort().map((f) => join(dir, f))];
  const h = createHash('sha256');
  for (const f of files) h.update(`${basename(f)}\n`).update(readFileSync(f, 'utf8').replace(/\r\n/g, '\n'));
  // 묶음에 넣는 제3자 고지 (저장소 LICENSES/)
  h.update(`${basename(THIRD_PARTY_FILE)}\n`).update(readFileSync(THIRD_PARTY_FILE, 'utf8').replace(/\r\n/g, '\n'));
  return h.digest('hex');
}

function listFiles(dir, base = dir) {
  const out = [];
  for (const ent of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, ent.name);
    if (ent.isSymbolicLink()) fail(`${full}: 링크는 받지 않습니다.`);
    if (ent.isDirectory()) out.push(...listFiles(full, base));
    else out.push(posix(relative(base, full)));
  }
  return out.sort();
}

// ── 8. 건너뛰기 ──
function upToDate(out, want) {
  try {
    const v = readJson(join(out, 'VERSION.json'));
    for (const k of ['name', 'pinsSha256', 'rulesSha256', 'builderSha256', 'patchesSha256', 'rebuiltApp', 'officeKitVersion']) {
      if (v[k] !== want[k]) return false;
    }
    for (const f of v.files) {
      const p = join(out, ...f.path.split('/'));
      if (statSync(p).size !== f.bytes) return false;
    }
    return v;
  } catch {
    return false;
  }
}

// ── --verify ──
function verifyOutput(out) {
  const vp = join(out, 'VERSION.json');
  if (!existsSync(vp)) fail(`${shown(vp)}가 없습니다 — 먼저 npm run office-editor:build`);
  const v = readJson(vp);
  const problems = [];
  const listed = new Set(v.files.map((f) => f.path));
  for (const f of v.files) {
    const p = join(out, ...f.path.split('/'));
    if (!existsSync(p)) { problems.push(`- ${f.path} (없음)`); continue; }
    const buf = readFileSync(p);
    if (buf.length !== f.bytes || sha256(buf) !== f.sha256) problems.push(`~ ${f.path} (바뀜)`);
  }
  for (const p of listFiles(out)) if (p !== 'VERSION.json' && !listed.has(p)) problems.push(`+ ${p} (VERSION.json에 없음)`);
  if (problems.length) fail(`${shown(out)}가 VERSION.json과 다릅니다:\n  ${problems.slice(0, 40).join('\n  ')}${problems.length > 40 ? `\n  … ${problems.length - 40}개 더` : ''}`);
  log(`확인 완료: ${shown(out)} — VERSION.json에 적힌 파일 ${num(v.files.length)}개가 모두 크기·SHA-256 그대로이고 다른 파일은 없습니다 (filex ${v.filex?.version}, ONLYOFFICE ${v.onlyoffice?.build}).`);
}

// ── 3. 풀기 ──
function safeZipName(name) {
  if (!name || name.includes('\\') || name.startsWith('/') || /^[a-z]:/i.test(name)) return false;
  const segs = name.split('/');
  return segs.every((s) => s !== '' && s !== '.' && s !== '..' && !/[<>:"|?*\x00-\x1f]/.test(s));
}

async function extractUi(zipPath, dest) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(readFileSync(zipPath), { createFolders: false });
  const part = `${dest}.partial`;
  rmSync(part, { recursive: true, force: true });
  mkdirSync(part, { recursive: true });
  const files = new Map();
  const lower = new Set();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    if (!safeZipName(name)) fail(`ui.zip에 이상한 경로가 있습니다: ${JSON.stringify(name)}`);
    const mode = entry.unixPermissions;
    if (typeof mode === 'number' && (mode & 0o170000) === 0o120000) fail(`ui.zip에 링크가 있습니다: ${name}`);
    const key = name.toLowerCase();
    if (lower.has(key)) fail(`ui.zip에 대소문자만 다른 경로가 있습니다: ${name}`);
    lower.add(key);
    const data = await entry.async('nodebuffer');
    const abs = join(part, ...name.split('/'));
    mkdirSync(dirname(abs), { recursive: true });
    writeFileSync(abs, data);
    files.set(name, { bytes: data.length, sha256: sha256(data) });
  }
  if (files.size !== UI_ZIP.files) fail(`ui.zip의 파일 수가 ${files.size}개입니다 (고정: ${UI_ZIP.files}개).`);
  rmSync(dest, { recursive: true, force: true });
  renameRetry(part, dest);
  return files;
}

// ── 5. 앱 다시 빌드 ──
async function rebuildApp({ patches, opts, lock, release, source }) {
  const { socketIo: socketPath } = PATHS;
  await fetchPinned({ url: SOCKET_IO_ORIGINAL.url, dest: socketPath, bytes: SOCKET_IO_ORIGINAL.bytes, sha256: SOCKET_IO_ORIGINAL.sha256, label: `원래 socket.io (ONLYOFFICE web-apps ${ONLYOFFICE.sourceTag})`, log, offline: opts.offline });
  if (opts.offline && !existsSync(join(PATHS.filexSrc, '.git'))) fail('--offline인데 filex 소스가 캐시에 없습니다.');
  const sparse = sparseDirsFor(patches);
  ensureFilexSource({ dir: PATHS.filexSrc, sparseDirs: sparse, log });
  applyPatches({ dir: PATHS.filexSrc, patches, scratch: join(CACHE, 'patches-lf'), log });
  installDeps({ dir: PATHS.filexSrc, log });
  const made = makeLockEditor({ uiDir: PATHS.ui, lock, outDir: PATHS.lockEditor, socketIo: readFileSync(socketPath) });
  log(`  lock 그대로의 편집기 폴더: 파일 ${num(made.files)}개 (하드 링크 ${num(made.linked)}개)`);
  const { ui, source: usedSource } = buildApp({ dir: PATHS.filexSrc, editorDir: PATHS.lockEditor, x2tDir: join(PATHS.ui, 'x2t'), outDir: PATHS.appOut, source, log });

  // 다시 빌드한 것에서는 filex가 만드는 파일만 가져온다. 편집기 파일·x2t는 릴리스와 바이트가 같아야 한다.
  const lockCopied = new Set(Object.keys(lock.files).filter((p) => p !== BRIDGE && p !== CHANGES).map((p) => EDITOR_PREFIX + p));
  const own = new Set([BRIDGE, CHANGES, ...Object.keys(ADDED)].map((p) => EDITOR_PREFIX + p));
  const rebuilt = new Map();
  const problems = [];
  for (const p of listFiles(ui)) {
    const abs = join(ui, ...p.split('/'));
    if (lockCopied.has(p)) {
      if (statSync(abs).size !== release.get(p)?.bytes) problems.push(`~ ${p} (편집기 파일이 릴리스와 다름)`);
      continue;
    }
    const buf = readFileSync(abs);
    const info = { bytes: buf.length, sha256: sha256(buf), abs };
    if (X2T_FILES.includes(p)) {
      if (info.sha256 !== X2T.files[p.slice(4)]) problems.push(`~ ${p} (x2t가 고정 값과 다름)`);
      continue;
    }
    if (p.startsWith(EDITOR_PREFIX) && !own.has(p)) { problems.push(`+ ${p} (패치가 편집기 파일을 더함 — 편집기 파일은 고치지 않는다)`); continue; }
    rebuilt.set(p, info);
  }
  for (const p of lockCopied) if (!existsSync(join(ui, ...p.split('/')))) problems.push(`- ${p} (다시 빌드한 것에 없음)`);
  if (problems.length) fail(`다시 빌드한 filex 묶음이 예상과 다릅니다:\n  ${problems.slice(0, 30).join('\n  ')}`);
  const changed = [...rebuilt].filter(([p, i]) => release.get(p)?.sha256 !== i.sha256).map(([p, i]) => ({ path: p, bytes: i.bytes, sha256: i.sha256, before: release.get(p)?.sha256 || null }));
  const gone = [...release.keys()].filter((p) => !p.startsWith(EDITOR_PREFIX) && !X2T_FILES.includes(p) && !rebuilt.has(p));
  for (const p of own) if (!rebuilt.has(p)) gone.push(p);
  return { rebuilt, changed, gone, source: usedSource };
}

// 폴더 이름 바꾸기: Windows에서 백신·색인기가 막 쓴 파일을 잠깐 잡고 있으면 실패하므로 몇 번 다시 한다
function renameRetry(from, to) {
  for (let i = 0; ; i++) {
    try { renameSync(from, to); return; } catch (e) {
      if (i >= 5 || !['EPERM', 'EBUSY', 'EACCES'].includes(e.code)) throw e;
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 300 * (i + 1));
    }
  }
}

function replaceDir(stage, dest) {
  const old = join(dirname(dest), `.${basename(dest)}.old-${process.pid}`);
  try {
    if (existsSync(dest)) renameRetry(dest, old);
    renameRetry(stage, dest);
  } catch (e) {
    if (!existsSync(dest) && existsSync(old)) { try { renameSync(old, dest); } catch { /* 그대로 둠 */ } }
    fail(`${shown(dest)}를 바꾸지 못했습니다 (${e.code || e.message}). Office Kit 앱이나 탐색기가 그 폴더를 쓰고 있으면 닫고 다시 실행하세요. 이전 묶음은 그대로입니다.`);
  }
  rmSync(old, { recursive: true, force: true });
}

function elapsed(ms) {
  const s = Math.round(ms / 1000);
  return s >= 60 ? `${Math.floor(s / 60)}분 ${s % 60}초` : `${s}초`;
}

async function main() {
  const t0 = Date.now();
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) { console.log(HELP); return; }
  if (opts.verify) { verifyOutput(opts.out); return; }

  const pkg = readJson(join(ROOT, 'package.json'));
  const officeKitVersion = pkg.version;
  // 패치한 수정판의 소스가 있는 곳 (편집기 아래 법적 고지 줄 — 이 판의 Release에 소스 꾸러미가 붙는다)
  const appSource = `${String(pkg.homepage || 'https://github.com/toxh13/office-kit').replace(/^https?:\/\//, '').replace(/\/+$/, '')}/releases/tag/v${officeKitVersion}`;
  const patches = listPatches(opts.patches);
  const rebuildWanted = patches.length > 0 || opts.rebuildApp;
  const want = {
    name: 'office-editor',
    pinsSha256: pinsSha256(),
    rulesSha256: rulesSha256(),
    builderSha256: builderSha256(),
    patchesSha256: patchesSha256(patches),
    rebuiltApp: rebuildWanted,
    officeKitVersion,
  };
  const current = !opts.force && upToDate(opts.out, want);
  if (current) {
    log(`이미 준비되어 있습니다: ${shown(opts.out)} (filex ${FILEX.version}, ONLYOFFICE ${ONLYOFFICE.build}, 파일 ${num(current.fileCount + 1)}개(VERSION.json 포함), ${mb(current.totalBytes)}${patches.length ? `, 패치 ${patches.length}개` : ''}) — 다시 만들려면 --force`);
    return;
  }

  const stage = join(dirname(opts.out), `.${basename(opts.out)}.tmp`);
  rmSync(stage, { recursive: true, force: true });

  // 1·2. 받기·확인
  log(`1·2단계: filex ${FILEX.tag} ui.zip 받기·확인 (${shown(PATHS.uiZip)})`);
  const zipFrom = await fetchPinned({ url: UI_ZIP.url, dest: PATHS.uiZip, bytes: UI_ZIP.bytes, sha256: UI_ZIP.sha256, label: 'ui.zip', log, offline: opts.offline });
  await fetchPinned({ url: EDITOR_LOCK.url, dest: PATHS.lock, bytes: EDITOR_LOCK.bytes, sha256: EDITOR_LOCK.sha256, label: `filex ${EDITOR_LOCK.path}`, log, offline: opts.offline });
  const lock = parseLock(readFileSync(PATHS.lock));

  // 3. 풀기
  log(`3단계: 풀기 → ${shown(PATHS.ui)}`);
  const release = await extractUi(PATHS.uiZip, PATHS.ui);

  // 4. 파일별 확인
  const checked = verifyUi({ files: release, read: (p) => readFileSync(join(PATHS.ui, ...p.split('/'))), lock });
  log(`4단계: 편집기 파일 ${num(checked.editorFiles)}개가 filex lock(커밋 ${FILEX.commit.slice(0, 12)})과 같음 — 다른 것은 filex 다리·CHANGES.txt 덧붙임(${checked.changes.appendedBytes}바이트)·빈 목록 ${checked.added.length}개뿐, x2t 두 파일 고정 SHA-256 일치`);

  // 최종 묶음의 파일: 묶음 안 경로 → { bytes, sha256, abs }
  const final = new Map();
  for (const [p, i] of release) final.set(p, { ...i, abs: join(PATHS.ui, ...p.split('/')) });

  // 5. (패치가 있을 때) 앱 다시 빌드
  let rebuilt = null;
  if (rebuildWanted) {
    log(`5단계: filex 앱 페이지 다시 빌드 (${patches.length ? `패치 ${patches.length}개: ${patches.map((p) => p.name).join(', ')}` : '패치 없음 — --rebuild-app'})`);
    const r = await rebuildApp({ patches, opts, lock, release, source: patches.length ? appSource : null });
    for (const p of r.gone) final.delete(p);
    for (const [p, i] of r.rebuilt) final.set(p, i);
    rebuilt = { source: r.source, changed: r.changed.map(({ path, bytes, sha256: s, before }) => ({ path, bytes, sha256: s, releaseSha256: before })), removed: r.gone };
    log(r.changed.length
      ? `  릴리스와 달라진 파일 ${r.changed.length}개: ${r.changed.map((c) => c.path).join(', ')}`
      : '  다시 빌드한 앱 페이지·다리가 릴리스와 바이트까지 같습니다.');
  } else {
    log('5단계: 건너뜀 (filex 패치 없음)');
  }

  // 6. 덜어 내기
  const { kept, removed } = partition([...final.keys()].sort());
  const removedInfo = DROP_RULES.map((r) => {
    const paths = removed.get(r.id);
    return { id: r.id, what: r.what, why: r.why, files: paths.length, bytes: paths.reduce((s, p) => s + final.get(p).bytes, 0) };
  });
  for (const r of removedInfo) log(`6단계: 뺌 — ${r.what}: ${num(r.files)}개, ${mb(r.bytes)}`);

  // 7. 쓰기
  mkdirSync(stage, { recursive: true });
  const files = [];
  for (const p of kept) {
    const i = final.get(p);
    const dest = join(stage, ...p.split('/'));
    mkdirSync(dirname(dest), { recursive: true });
    copyFileSync(i.abs, dest);
    if (statSync(dest).size !== i.bytes) fail(`${p}를 복사하지 못했습니다 (크기가 다름).`);
    files.push({ path: p, bytes: i.bytes, sha256: i.sha256 });
  }
  const builtAt = new Date().toISOString();
  const keptBytes = files.reduce((s, f) => s + f.bytes, 0);
  const noticeBuf = Buffer.from(renderNotice({ builtAt, officeKitVersion, removed: removedInfo, patches, rebuilt, fileCount: files.length, totalBytes: keptBytes }), 'utf8');
  writeFileSync(join(stage, 'NOTICE-OfficeKit.md'), noticeBuf);
  files.push({ path: 'NOTICE-OfficeKit.md', bytes: noticeBuf.length, sha256: sha256(noticeBuf) });
  // 기계어(wasm)·라이선스 파일 없는 JS의 제3자 고지 (scripts/office-editor/third-party.mjs가 만든 LICENSES/ 파일)
  const tpText = readFileSync(THIRD_PARTY_FILE, 'utf8').replace(/\r\n/g, '\n');
  const tpBad = checkThirdPartyFile(tpText);
  if (tpBad.length) fail(`${shown(THIRD_PARTY_FILE)}가 목록과 다릅니다 (${tpBad.join(', ')}) — node scripts/office-editor/third-party.mjs --write`);
  const tpBuf = Buffer.from(tpText, 'utf8');
  writeFileSync(join(stage, THIRD_PARTY_BUNDLE_PATH), tpBuf);
  files.push({ path: THIRD_PARTY_BUNDLE_PATH, bytes: tpBuf.length, sha256: sha256(tpBuf) });
  files.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));

  const fileCount = files.length;
  const totalBytes = files.reduce((s, f) => s + f.bytes, 0);
  const contentSha256 = sha256(files.map((f) => `${f.path}\t${f.bytes}\t${f.sha256}\n`).join(''));
  const version = {
    ...want,
    about: 'Office Kit 문서·표·슬라이드 편집 창의 편집기 묶음 — scripts/build-office-editor.mjs가 만든다. 고지: NOTICE-OfficeKit.md',
    editorVersion: ONLYOFFICE.build,
    urlVersion: `${ONLYOFFICE.build}-${contentSha256.slice(0, 10)}`,
    contentSha256,
    onlyoffice: { product: ONLYOFFICE.product, vendor: ONLYOFFICE.vendor, version: ONLYOFFICE.version, build: ONLYOFFICE.build, image: ONLYOFFICE.image, sourceTag: ONLYOFFICE.sourceTag },
    filex: { name: FILEX.name, repo: FILEX.repo, tag: FILEX.tag, version: FILEX.version, commit: FILEX.commit, license: FILEX.license },
    uiZip: { url: UI_ZIP.url, bytes: UI_ZIP.bytes, sha256: UI_ZIP.sha256, files: UI_ZIP.files },
    editorLock: { url: EDITOR_LOCK.url, bytes: EDITOR_LOCK.bytes, sha256: EDITOR_LOCK.sha256, files: EDITOR_LOCK.files },
    x2t: { release: X2T.release, files: X2T.files, core: X2T.core.commit, coreTag: X2T.core.tag, buildTools: X2T.buildTools.commit, recipe: X2T.recipe },
    bridge: final.get(EDITOR_PREFIX + BRIDGE) ? { path: EDITOR_PREFIX + BRIDGE, bytes: final.get(EDITOR_PREFIX + BRIDGE).bytes, sha256: final.get(EDITOR_PREFIX + BRIDGE).sha256 } : null,
    patches: patches.map(({ name, bytes, sha256: s, subject, date, touched }) => ({ name, bytes, sha256: s, subject, date, touched })),
    rebuilt,
    removed: removedInfo,
    noticeDated: NOTICE_DATED,
    pins: PINS,
    builtAt,
    builder: 'scripts/build-office-editor.mjs',
    node: process.version,
    fileCount,
    totalBytes,
    files,
  };
  writeFileSync(join(stage, 'VERSION.json'), `${JSON.stringify(version, null, 2)}\n`);
  mkdirSync(dirname(opts.out), { recursive: true });
  replaceDir(stage, opts.out);
  // 5단계의 중간 결과(약 320MB)는 다음 빌드에서 다시 만든다 — 디스크를 비운다 (filex 소스·node_modules는 둠)
  for (const d of [PATHS.appOut, PATHS.lockEditor]) rmSync(d, { recursive: true, force: true });

  // 결과 (설계 값과 나란히)
  const removedFiles = removedInfo.reduce((s, r) => s + r.files, 0);
  const removedBytes = removedInfo.reduce((s, r) => s + r.bytes, 0);
  const lines = [
    '',
    'Office 편집기 묶음 만들기 완료',
    `  원본   ${FILEX.name} ${FILEX.tag} ui.zip (${zipFrom === 'cache' ? '캐시 사용' : '새로 받음'}, ${num(release.size)}개 ${mb([...release.values()].reduce((s, i) => s + i.bytes, 0))}) · ONLYOFFICE Docs ${ONLYOFFICE.build} · x2t ${X2T.release}`,
    `  확인   편집기 파일 ${num(checked.editorFiles)}개 lock 일치, x2t 고정 SHA-256 일치${patches.length ? `, 패치 ${patches.length}개로 앱 다시 빌드` : ''}`,
    `  뺀 것  ${num(removedFiles)}개, ${mb(removedBytes)}`,
    `  결과   ${shown(opts.out)} — 파일 ${num(fileCount + 1)}개 (VERSION.json 포함), ${mb(totalBytes)}`,
    `         설계 값: ${mb(DESIGN.bytes)}, ${num(DESIGN.files)}개 (gzip-9 ${mb(DESIGN.gzipBytes)})`,
  ];
  if (opts.measure) {
    let gz = 0;
    for (const f of files) gz += gzipSync(readFileSync(join(opts.out, ...f.path.split('/'))), { level: 9 }).length;
    lines.push(`  gzip-9 ${mb(gz)} (설치 파일이 늘어나는 크기 어림 — NSIS LZMA는 이보다 작거나 같음)`);
  }
  lines.push(`  걸린 시간 ${elapsed(Date.now() - t0)}`);
  console.log(lines.join('\n'));
}

main().catch((e) => {
  if (e instanceof BuildError || e instanceof FetchError || e instanceof StepError) console.error(`\n[office-editor] 오류: ${e.message}`);
  // lock.mjs·rules.mjs가 던지는 보통 Error(한국어 안내)도 알려진 오류로
  else if (e?.constructor === Error) console.error(`\n[office-editor] 오류: ${e.message}`);
  else console.error(`\n[office-editor] 예상하지 못한 오류: ${e?.stack || e}`);
  process.exit(1);
});
