#!/usr/bin/env node
// build-rhwp-cli.mjs — 한컴오피스 없이 한글 문서를 PDF로 만드는 rhwp 명령줄 프로그램(rhwp.exe)을 vendor/rhwp-cli에 넣는다.
//
//   node scripts/build-rhwp-cli.mjs [--force] [--offline] [--verify]     (npm run rhwp-cli:build)
//
// 하는 일
//   1. github.com/edwardkim/rhwp Release v0.8.7의 rhwp-v0.8.7-windows-x86_64.zip을 받는다 (고정 크기·SHA-256, 다르면 지우고 멈춤).
//      THIRD_PARTY_LICENSES.md는 같은 태그의 고정 커밋에서 받는다 (고정 크기·SHA-256).
//   2. zip에서 rhwp/rhwp.exe와 rhwp/LICENSE만 꺼낸다 (README는 뺀다). 꺼낸 파일도 고정 SHA-256과 맞춘다.
//   3. Windows면 rhwp.exe --version이 'rhwp v0.8.7'인지 확인한다.
//   4. vendor/rhwp-cli에 rhwp.exe·LICENSE·THIRD_PARTY_LICENSES.md·VERSION.json을 쓴다.
//
// 앱 미리보기(@rhwp/core WASM)와 같은 조판 엔진 판이어야 PDF가 미리보기와 같게 나온다 → node_modules/@rhwp/core가 정확히 0.8.7이 아니면 멈춘다.
// 이미 같은 결과가 있으면 아무것도 하지 않는다 (--force: 다시 만듦). --verify: 받지 않고 vendor/rhwp-cli만 고정 값과 맞춰 본다.
// 받은 파일 캐시: tmp/rhwp-cli-build/download (OFFICEKIT_RHWP_CLI_CACHE로 바꿈, CI가 캐시한다).
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { FetchError, fetchPinned } from './office-editor/fetch.mjs';

// ── 고정 값 (올릴 때는 여기만 바꾸고, 받은 파일을 직접 확인한 뒤 크기·SHA-256을 적는다) ──
// 올릴 때 함께: package.json의 @rhwp/core, scripts/build-editor.mjs의 TAG·COMMIT (미리보기·편집기·PDF가 같은 엔진 판이 되게)
const TAG = 'v0.8.7';
const COMMIT = '1a76570e833917d15817415a53c09ad61ab3203f';
const CORE_VERSION = '0.8.7';
const ZIP = {
  name: 'rhwp-v0.8.7-windows-x86_64.zip',
  url: 'https://github.com/edwardkim/rhwp/releases/download/v0.8.7/rhwp-v0.8.7-windows-x86_64.zip',
  bytes: 19_796_112,
  sha256: 'f7ac326e98f595ae8fecd8da7a614ee9a37db77bdf938782a8e07c9da473ec6f',
};
// zip 안 파일 → vendor/rhwp-cli 안 이름
const FROM_ZIP = [
  { entry: 'rhwp/rhwp.exe', path: 'rhwp.exe', bytes: 49_458_688, sha256: '1e1b7ae6c9ffd895e1a330a191666a2629cc7eba2fc9e37442e9f1755741bb62' },
  { entry: 'rhwp/LICENSE', path: 'LICENSE', bytes: 1093, sha256: 'd6ea0e91bf1e6136b606e288e7e22dc44e39b4f087d4bdcde2bd0d8aaf5314dd' },
];
const THIRD_PARTY = {
  path: 'THIRD_PARTY_LICENSES.md',
  url: `https://raw.githubusercontent.com/edwardkim/rhwp/${COMMIT}/THIRD_PARTY_LICENSES.md`,
  bytes: 13_162,
  sha256: 'bb40fa5d0fa44bf2ff8b2b5dd7085562b58557a69eaa08b6bb5a4fc94252af88',
};
const VERSION_LINE = `rhwp ${TAG}`;

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const CACHE = resolve(process.env.OFFICEKIT_RHWP_CLI_CACHE || join(ROOT, 'tmp', 'rhwp-cli-build', 'download'));
const OUT = join(ROOT, 'vendor', 'rhwp-cli');
const CORE_PKG = join(ROOT, 'node_modules', '@rhwp', 'core', 'package.json');

const log = (s) => process.stdout.write(`[rhwp-cli] ${s}\n`);
class BuildError extends Error {}
const fail = (msg) => { throw new BuildError(msg); };
const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');
const mb = (n) => `${(n / 1048576).toFixed(1)} MB`;

const HELP = `한글→PDF 변환 프로그램(rhwp ${TAG} 명령줄) 받기 → vendor/rhwp-cli

  node scripts/build-rhwp-cli.mjs [--force] [--offline] [--verify]

  --force    결과가 이미 있어도 다시 만듦
  --offline  받지 않고 캐시만 씀 (없으면 실패)
  --verify   vendor/rhwp-cli가 고정 값(파일·크기·SHA-256)과 같은지 확인만

  받는 것   ${ZIP.url}
            (${mb(ZIP.bytes)}, SHA-256 ${ZIP.sha256.slice(0, 16)}…)
  캐시 폴더 ${relative(ROOT, CACHE) || CACHE} (OFFICEKIT_RHWP_CLI_CACHE로 바꿈)`;

function parseArgs(argv) {
  const o = { force: false, offline: false, verify: false, help: false };
  for (const a of argv) {
    if (a === '--force' || a === '-f') o.force = true;
    else if (a === '--offline') o.offline = true;
    else if (a === '--verify') o.verify = true;
    else if (a === '--help' || a === '-h') o.help = true;
    else fail(`모르는 옵션: ${a} (--force, --offline, --verify, --help)`);
  }
  return o;
}

function checkCore() {
  if (!existsSync(CORE_PKG)) fail('node_modules/@rhwp/core가 없습니다 — 저장소 폴더에서 먼저 npm ci를 실행하세요.');
  const v = JSON.parse(readFileSync(CORE_PKG, 'utf8')).version;
  if (v !== CORE_VERSION) {
    fail(`node_modules/@rhwp/core 버전이 ${v}입니다 (필요: 정확히 ${CORE_VERSION}). PDF가 미리보기와 같은 엔진으로 만들어지도록 판을 맞추세요.`);
  }
}

// 고정한 파일 목록 (VERSION.json 제외) → [{ path, bytes, sha256 }]
const PINNED = [...FROM_ZIP.map(({ path, bytes, sha256: s }) => ({ path, bytes, sha256: s })),
  { path: THIRD_PARTY.path, bytes: THIRD_PARTY.bytes, sha256: THIRD_PARTY.sha256 }];

/** vendor/rhwp-cli가 고정 값 그대로인지 → 문제 목록 (없으면 빈 배열) */
function problemsIn(dir) {
  const out = [];
  for (const f of PINNED) {
    const p = join(dir, f.path);
    if (!existsSync(p)) { out.push(`- ${f.path} (없음)`); continue; }
    if (statSync(p).size !== f.bytes) { out.push(`~ ${f.path} (크기가 다름)`); continue; }
    if (sha256(readFileSync(p)) !== f.sha256) out.push(`~ ${f.path} (SHA-256이 다름)`);
  }
  let v = null;
  try { v = JSON.parse(readFileSync(join(dir, 'VERSION.json'), 'utf8')); } catch { out.push('- VERSION.json (없거나 읽을 수 없음)'); }
  if (v && (v.name !== 'rhwp-cli' || v.tag !== TAG || v.zip?.sha256 !== ZIP.sha256)) out.push('~ VERSION.json (다른 판)');
  return out;
}

function checkVersion(exe) {
  if (process.platform !== 'win32') { log('Windows가 아니라 rhwp.exe 실행 확인은 건너뜁니다.'); return null; }
  const r = spawnSync(exe, ['--version'], { encoding: 'utf8', windowsHide: true, timeout: 30_000 });
  if (r.error) fail(`rhwp.exe를 실행하지 못했습니다: ${r.error.message}`);
  const line = String(r.stdout || '').trim().split(/\r?\n/)[0];
  if (r.status !== 0 || line !== VERSION_LINE) fail(`rhwp.exe --version이 '${VERSION_LINE}'이 아닙니다: ${line || r.stderr}`);
  return line;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) { console.log(HELP); return; }
  if (opts.verify) {
    const p = problemsIn(OUT);
    if (p.length) fail(`vendor/rhwp-cli가 고정 값과 다릅니다 — npm run rhwp-cli:build -- --force\n  ${p.join('\n  ')}`);
    checkVersion(join(OUT, 'rhwp.exe'));
    log(`확인 완료: vendor/rhwp-cli — ${PINNED.map((f) => f.path).join('·')}가 고정 크기·SHA-256 그대로 (${VERSION_LINE})`);
    return;
  }
  checkCore();
  if (!opts.force && !problemsIn(OUT).length) {
    log(`이미 있음: vendor/rhwp-cli (${VERSION_LINE}) — 다시 만들려면 --force`);
    return;
  }
  const t0 = Date.now();
  const zipPath = join(CACHE, ZIP.name);
  const tpPath = join(CACHE, `${THIRD_PARTY.path}-${COMMIT.slice(0, 12)}`);
  log(`1단계: ${ZIP.name} (${mb(ZIP.bytes)}) 받기`);
  const zipFrom = await fetchPinned({ ...ZIP, dest: zipPath, label: ZIP.name, log, offline: opts.offline });
  const tpFrom = await fetchPinned({ ...THIRD_PARTY, dest: tpPath, label: THIRD_PARTY.path, log, offline: opts.offline });

  log('2단계: zip에서 rhwp.exe·LICENSE 꺼내기');
  const zip = await JSZip.loadAsync(readFileSync(zipPath));
  const stage = `${OUT}.new`;
  rmSync(stage, { recursive: true, force: true });
  mkdirSync(stage, { recursive: true });
  for (const f of FROM_ZIP) {
    const entry = zip.file(f.entry);
    if (!entry) fail(`zip 안에 ${f.entry}가 없습니다.`);
    const buf = await entry.async('nodebuffer');
    if (buf.length !== f.bytes || sha256(buf) !== f.sha256) fail(`zip 안 ${f.entry}가 고정 값과 다릅니다 (${buf.length}바이트, ${sha256(buf)}).`);
    writeFileSync(join(stage, f.path), buf);
  }
  writeFileSync(join(stage, THIRD_PARTY.path), readFileSync(tpPath));

  log('3단계: rhwp.exe 실행 확인');
  const version = checkVersion(join(stage, 'rhwp.exe'));

  const builderSha = sha256(readFileSync(fileURLToPath(import.meta.url)));
  writeFileSync(join(stage, 'VERSION.json'), `${JSON.stringify({
    name: 'rhwp-cli',
    purpose: '한컴오피스 없이 한글(hwp·hwpx·hwt) 문서를 PDF로 (export-pdf, svg 방식)',
    repo: 'https://github.com/edwardkim/rhwp',
    license: 'MIT',
    tag: TAG,
    commit: COMMIT,
    coreVersion: CORE_VERSION,
    version: version ?? VERSION_LINE,
    zip: ZIP,
    files: PINNED,
    builtAt: new Date().toISOString(),
    builder: 'scripts/build-rhwp-cli.mjs',
    builderSha256: builderSha,
  }, null, 2)}\n`);

  // 바꿔 끼우기 (반쯤 쓴 폴더가 남지 않게)
  const old = `${OUT}.old`;
  rmSync(old, { recursive: true, force: true });
  if (existsSync(OUT)) renameSync(OUT, old);
  try { renameSync(stage, OUT); } catch (e) {
    if (!existsSync(OUT) && existsSync(old)) { try { renameSync(old, OUT); } catch { /* 그대로 둠 */ } }
    throw e;
  }
  rmSync(old, { recursive: true, force: true });
  const p = problemsIn(OUT);
  if (p.length) fail(`만든 vendor/rhwp-cli가 고정 값과 다릅니다:\n  ${p.join('\n  ')}`);
  console.log([
    '',
    '한글→PDF 변환 프로그램(rhwp 명령줄) 준비 완료',
    `  원본   edwardkim/rhwp ${TAG} Release ${ZIP.name} (${zipFrom === 'cache' ? '캐시 사용' : '새로 받음'}, SHA-256 일치) · THIRD_PARTY_LICENSES.md (${tpFrom === 'cache' ? '캐시' : '새로 받음'})`,
    `  결과   vendor/rhwp-cli — rhwp.exe ${mb(FROM_ZIP[0].bytes)}, LICENSE(MIT), THIRD_PARTY_LICENSES.md, VERSION.json`,
    `  걸린 시간 ${((Date.now() - t0) / 1000).toFixed(1)}초`,
  ].join('\n'));
}

main().catch((e) => {
  if (e instanceof BuildError || e instanceof FetchError) console.error(`\n[rhwp-cli] 오류: ${e.message}`);
  else console.error(`\n[rhwp-cli] 예상하지 못한 오류: ${e.stack || e}`);
  process.exit(1);
});
