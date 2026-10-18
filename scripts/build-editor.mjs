#!/usr/bin/env node
// build-editor.mjs — 앱 안 한글 편집기(rhwp-studio)를 고정된 태그에서 빌드해 vendor/에 넣는다.
//
//   node scripts/build-editor.mjs [--force] [--with-fonts]     (npm run editor:build)
//
// 하는 일
//   1. github.com/edwardkim/rhwp 태그 v0.8.7을 필요한 폴더만 받는다 (sparse·blob:none, 캐시 폴더에 두고 다시 씀).
//      받은 커밋이 COMMIT과 다르면 멈춘다. node_modules/@rhwp/core·@rhwp/editor가 정확히 0.8.7이 아니어도 멈춘다.
//   2. Rust 없이: npm @rhwp/core 0.8.7의 WASM 파일을 <소스>/pkg/와 rhwp-studio/public/에 넣는다.
//   3. rhwp-studio에서 npm ci(스크립트 실행 안 함) → vite build --base=./
//      (RHWP_WITHOUT_HWPCTRL=1: hwpctrl 플러그인 뺌, RHWP_DISABLE_EXTERNAL_WEBFONTS=1: CDN 웹폰트 끔).
//   4. 결과를 vendor/rhwp-studio로 옮기며 오프라인 앱에 필요 없는 것을 뺀다:
//      - samples/ (예제 문서), 최상위 rhwp.js·rhwp_bg.wasm·*.d.ts (번들은 assets/의 WASM을 씀)
//      - PWA·서비스 워커: sw.js, workbox-*.js, registerSW.js, manifest.webmanifest, PWA 설치용 아이콘
//      - index.html에서 <link rel="manifest">와 <script id="vite-plugin-pwa:register-sw">를 지운다
//        → 페이지가 서비스 워커를 등록하지 않는다. 남은 파일에 serviceWorker·registerSW·workbox 글자가 있으면 실패.
//      - fonts (core.symlinks=false로 받아 심볼릭 링크 대신 생긴 글자 파일). --with-fonts면 오픈 라이선스 글꼴을 대신 넣는다.
//      그리고 rhwp LICENSE·THIRD_PARTY_LICENSES.md, 번들에 들어간 npm 패키지 라이선스(licenses/), VERSION.json을 쓴다.
//   5. @rhwp/editor SDK의 호스트 페이지용 파일을 vendor/rhwp-editor-sdk로 복사한다.
//
// 이미 같은 결과가 있으면 아무것도 하지 않는다 (--force: 다시 빌드). 캐시: tmp/editor-build (OFFICEKIT_EDITOR_CACHE로 바꿈).
// 남은 한계: 번들 JS에 jsdelivr 글꼴 주소 문자열이 남아 있다(실행 시 꺼짐) → 앱이 편집기의 외부 요청을 막아야 한다.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync, existsSync, lstatSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync,
} from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// ── 고정 값 (올릴 때는 여기만 바꾸고 커밋을 직접 확인) ──
// 올릴 때 함께: package.json의 @rhwp/core(정확한 버전, ^ 없이)·@rhwp/editor(devDependencies)도 같은 버전으로
const REPO_URL = 'https://github.com/edwardkim/rhwp.git';
const TAG = 'v0.8.7';
const COMMIT = '1a76570e833917d15817415a53c09ad61ab3203f';
const CORE_VERSION = '0.8.7';
const EDITOR_VERSION = '0.8.7';
const SPARSE_DIRS = ['rhwp-studio', 'rhwp-shared', 'npm/editor', 'assets/logo', 'typescript'];
const FONT_DIR = 'assets/fonts';
const CORE_FILES = ['rhwp.js', 'rhwp_bg.wasm', 'rhwp.d.ts', 'rhwp_bg.wasm.d.ts'];
const SDK_FILES = ['index.js', 'transport.js', 'document-agent-contract.js'];
// 번들(assets/)에 실제로 들어가는 rhwp-studio 의존성 → 라이선스 전문을 함께 싣는다
const BUNDLED_DEPS = ['canvaskit-wasm', '@noble/hashes'];
// --with-fonts: 재배포 조건이 분명한 글꼴만 (SIL OFL 1.1·GUST). 카페24·행복고딕("무료 배포")은 뺀다
const FONT_EXCLUDE = /^(Cafe24|Happiness)/i;
const BUILD_ENV = { RHWP_WITHOUT_HWPCTRL: '1', RHWP_DISABLE_EXTERNAL_WEBFONTS: '1', RHWP_SUBSECOND: '0' };

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const CACHE = resolve(process.env.OFFICEKIT_EDITOR_CACHE || join(ROOT, 'tmp', 'editor-build'));
const SRC = join(CACHE, `rhwp-${TAG}`);
const STUDIO = join(SRC, 'rhwp-studio');
const VENDOR = join(ROOT, 'vendor');
const OUT = join(VENDOR, 'rhwp-studio');
const SDK_OUT = join(VENDOR, 'rhwp-editor-sdk');
const CORE_DIR = join(ROOT, 'node_modules', '@rhwp', 'core');
const EDITOR_DIR = join(ROOT, 'node_modules', '@rhwp', 'editor');

// dist에서 빼는 것: [규칙, 까닭]
const DROP = [
  [(p) => p === 'samples' || p.startsWith('samples/'), '예제 문서 (samples/)'],
  [(p) => /^(sw\.js|registerSW\.js|manifest\.webmanifest)$/.test(p) || /^workbox-[^/]*\.js$/.test(p), 'PWA·서비스 워커 (sw.js, workbox-*.js, registerSW.js, manifest.webmanifest)'],
  [(p) => /^icons\/icon-(128|192|512)\.png$/.test(p), 'PWA 설치용 아이콘 (icon-256만 남김)'],
  [(p) => /^(rhwp\.js|rhwp_bg\.wasm)$/.test(p), '최상위 rhwp.js·rhwp_bg.wasm 사본 (번들은 assets/의 WASM을 씀)'],
  [(p) => p.endsWith('.d.ts') || p.endsWith('.map'), '타입 정의·소스맵'],
  [(p) => p === 'fonts' || p.startsWith('fonts/'), 'fonts 링크 자리 파일 (core.symlinks=false라 심볼릭 링크 대신 생긴 글자 파일 — --with-fonts면 글꼴을 따로 넣음)'],
];
// 남으면 안 되는 글자 (서비스 워커 등록 흔적)
const SW_PATTERNS = [/serviceWorker/, /registerSW/, /workbox/i, /manifest\.webmanifest/];

const log = (s) => process.stdout.write(`[editor] ${s}\n`);

class BuildError extends Error {}
const fail = (msg) => { throw new BuildError(msg); };

function parseArgs(argv) {
  const opts = { force: false, fonts: false };
  for (const a of argv) {
    if (a === '--force' || a === '-f') opts.force = true;
    else if (a === '--with-fonts') opts.fonts = true;
    else if (a === '--help' || a === '-h') opts.help = true;
    else fail(`모르는 옵션: ${a} (--force, --with-fonts, --help)`);
  }
  return opts;
}

const HELP = `한글 편집기(rhwp-studio ${TAG}) 빌드 → vendor/rhwp-studio, vendor/rhwp-editor-sdk

  node scripts/build-editor.mjs [--force] [--with-fonts]

  --force       결과가 이미 있어도 다시 빌드
  --with-fonts  오픈 라이선스 웹 글꼴(약 20MB)을 함께 넣음 (기본: 넣지 않음 — 시스템 글꼴 사용)

  캐시 폴더: ${relative(ROOT, CACHE) || CACHE} (지우면 처음부터 다시 받음, OFFICEKIT_EDITOR_CACHE로 바꿈)`;

// ── 실행 도우미 ──
function run(cmd, args, { cwd, env, what }) {
  const r = spawnSync(cmd, args, { cwd, env: env || process.env, stdio: 'inherit', windowsHide: true });
  if (r.error?.code === 'ENOENT') fail(`${cmd}을(를) 찾을 수 없습니다 — 설치되어 있고 PATH에 있는지 확인하세요.`);
  if (r.error) fail(`${what} 실행 실패: ${r.error.message}`);
  if (r.status !== 0) fail(`${what}에 실패했습니다 (종료 코드 ${r.status}).`);
}

function gitOut(args, cwd) {
  const r = spawnSync('git', args, { cwd, encoding: 'utf8', windowsHide: true });
  if (r.error?.code === 'ENOENT') fail('git을 찾을 수 없습니다 — Git을 설치하세요.');
  if (r.status !== 0) return null;
  return r.stdout.trim();
}

// npm은 Windows에서 .cmd라 셸 없이 못 띄운다 → npm-cli.js를 node로 직접 실행
function npmCommand() {
  const cands = [
    process.env.npm_execpath,
    join(dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js'),
    join(dirname(process.execPath), '..', 'lib', 'node_modules', 'npm', 'bin', 'npm-cli.js'),
  ];
  const cli = cands.find((p) => p && /npm-cli\.js$/.test(p) && existsSync(p));
  if (!cli) fail('npm을 찾을 수 없습니다 (Node.js와 함께 설치되는 npm이 필요합니다).');
  return cli;
}

const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');
const readJson = (p) => JSON.parse(readFileSync(p, 'utf8'));
const mb = (n) => `${(n / 1024 / 1024).toFixed(1)} MB`;
const posix = (p) => p.split(sep).join('/');

function listFiles(dir, base = dir) {
  const out = [];
  for (const ent of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, ent.name);
    if (ent.isDirectory()) out.push(...listFiles(full, base));
    else out.push(posix(relative(base, full)));
  }
  return out.sort();
}

// ── 1. 버전 확인 ──
function checkPackages() {
  const core = join(CORE_DIR, 'package.json');
  if (!existsSync(core)) fail('node_modules/@rhwp/core가 없습니다 — 저장소 폴더에서 먼저 npm ci를 실행하세요.');
  const cv = readJson(core).version;
  if (cv !== CORE_VERSION) {
    fail(`node_modules/@rhwp/core 버전이 ${cv}입니다 (필요: 정확히 ${CORE_VERSION}). WASM 버전이 rhwp-studio ${TAG}에 맞지 않으면 편집기가 깨집니다 — npm ci로 맞추거나, 스크립트의 버전들을 함께 올리세요.`);
  }
  for (const f of CORE_FILES) if (!existsSync(join(CORE_DIR, f))) fail(`node_modules/@rhwp/core/${f}가 없습니다 — npm ci를 다시 실행하세요.`);
  const ed = join(EDITOR_DIR, 'package.json');
  if (!existsSync(ed)) fail('node_modules/@rhwp/editor가 없습니다 — 저장소 폴더에서 npm ci를 실행하세요.');
  const ev = readJson(ed).version;
  if (ev !== EDITOR_VERSION) fail(`node_modules/@rhwp/editor 버전이 ${ev}입니다 (필요: 정확히 ${EDITOR_VERSION}).`);
  for (const f of SDK_FILES) if (!existsSync(join(EDITOR_DIR, f))) fail(`node_modules/@rhwp/editor/${f}가 없습니다.`);
  return { wasmSha256: sha256(readFileSync(join(CORE_DIR, 'rhwp_bg.wasm'))) };
}

// ── 이미 만든 결과가 그대로인지 ──
function upToDate(opts, builderSha) {
  try {
    const v = readJson(join(OUT, 'VERSION.json'));
    if (v.tag !== TAG || v.commit !== COMMIT || v.coreVersion !== CORE_VERSION || v.fonts !== opts.fonts || v.builderSha256 !== builderSha) return false;
    for (const f of v.files) {
      const p = join(OUT, f.path);
      if (!existsSync(p) || statSync(p).size !== f.bytes) return false;
    }
    const s = readJson(join(SDK_OUT, 'VERSION.json'));
    if (s.version !== EDITOR_VERSION) return false;
    for (const f of s.files) if (!existsSync(join(SDK_OUT, f.path))) return false;
    return v;
  } catch {
    return false;
  }
}

// ── 2. 소스 받기 (캐시 재사용) ──
function verifyCommit(dir) {
  const head = gitOut(['rev-parse', 'HEAD^{commit}'], dir);
  const tag = gitOut(['rev-parse', `refs/tags/${TAG}^{commit}`], dir);
  if (head !== COMMIT || tag !== COMMIT) {
    fail(`받은 rhwp ${TAG}의 커밋이 예상과 다릅니다 (받음 ${head || tag || '없음'}, 예상 ${COMMIT}).\n` +
      '원본 저장소의 태그가 바뀌었을 수 있어 빌드를 멈춥니다. 새 커밋을 직접 확인한 뒤 scripts/build-editor.mjs의 COMMIT을 고치세요.');
  }
}

function ensureSource(opts) {
  const dirs = opts.fonts ? [...SPARSE_DIRS, FONT_DIR] : SPARSE_DIRS;
  if (existsSync(join(SRC, '.git'))) {
    const head = gitOut(['rev-parse', 'HEAD^{commit}'], SRC);
    if (head === COMMIT) {
      log(`캐시된 소스를 다시 씁니다: ${SRC}`);
      // 지난 빌드가 덮어쓴 추적 파일(public/rhwp.js)을 원래대로
      run('git', ['-C', SRC, 'reset', '-q', '--hard', 'HEAD'], { what: 'git reset' });
      run('git', ['-C', SRC, 'sparse-checkout', 'set', '--cone', ...dirs], { what: 'git sparse-checkout' });
      verifyCommit(SRC);
      return 'cache';
    }
    log(`캐시된 소스의 커밋이 다릅니다 (${head || '알 수 없음'}) — 지우고 새로 받습니다.`);
    rmSync(SRC, { recursive: true, force: true });
  }
  mkdirSync(CACHE, { recursive: true });
  const part = `${SRC}.partial`;
  rmSync(part, { recursive: true, force: true });
  log(`rhwp ${TAG} 소스를 받습니다 (필요한 폴더만): ${REPO_URL}`);
  // 심볼릭 링크·줄바꿈 변환을 꺼서 OS와 상관없이 같은 파일로 빌드
  run('git', ['clone', '-q', '--depth', '1', '--branch', TAG, '--filter=blob:none', '--sparse',
    '-c', 'core.symlinks=false', '-c', 'core.autocrlf=false', '-c', 'advice.detachedHead=false', REPO_URL, part],
  { what: 'rhwp 소스 받기 (git clone — 인터넷 연결 확인)' });
  run('git', ['-C', part, 'sparse-checkout', 'set', '--cone', ...dirs], { what: 'git sparse-checkout' });
  verifyCommit(part);
  renameSync(part, SRC);
  return 'clone';
}

// ── 3. WASM 넣기 → npm ci → vite build ──
function prepareWasm() {
  const pkg = join(SRC, 'pkg');
  mkdirSync(pkg, { recursive: true });
  for (const f of CORE_FILES) copyFileSync(join(CORE_DIR, f), join(pkg, f));
  for (const f of ['rhwp.js', 'rhwp_bg.wasm']) copyFileSync(join(CORE_DIR, f), join(STUDIO, 'public', f));
}

function installStudioDeps() {
  const lock = join(STUDIO, 'package-lock.json');
  const hasLock = existsSync(lock);
  const stamp = join(STUDIO, 'node_modules', '.officekit-stamp');
  // 설치 스크립트를 돌리지 않으므로 Node 버전은 결과에 영향이 없다 → 잠금 파일 내용만 본다
  const want = hasLock ? sha256(readFileSync(lock)) : sha256(readFileSync(join(STUDIO, 'package.json')));
  if (existsSync(stamp) && readFileSync(stamp, 'utf8') === want && existsSync(join(STUDIO, 'node_modules', 'vite', 'package.json'))) {
    log('rhwp-studio 의존성이 이미 설치되어 있습니다.');
    return;
  }
  log(`rhwp-studio 의존성을 설치합니다 (npm ${hasLock ? 'ci' : 'install'}, 설치 스크립트 실행 안 함)...`);
  const env = { ...process.env };
  // 상위 npm run이 넘기는 설정이 하위 설치에 섞이지 않게
  for (const k of Object.keys(env)) if (/^npm_config_/i.test(k) && !/^npm_config_(cache|registry|userconfig|proxy|https_proxy|noproxy|strict_ssl|cafile)$/i.test(k)) delete env[k];
  run(process.execPath, [npmCommand(), hasLock ? 'ci' : 'install', '--ignore-scripts', '--no-audit', '--no-fund', '--loglevel=error'],
    { cwd: STUDIO, env, what: 'rhwp-studio 의존성 설치' });
  writeFileSync(stamp, want);
}

function viteBuild() {
  const vitePkg = readJson(join(STUDIO, 'node_modules', 'vite', 'package.json'));
  const bin = typeof vitePkg.bin === 'string' ? vitePkg.bin : vitePkg.bin.vite;
  log(`vite ${vitePkg.version}로 rhwp-studio를 빌드합니다 (hwpctrl 없음, 외부 웹폰트 끔)...`);
  rmSync(join(STUDIO, 'dist'), { recursive: true, force: true });
  run(process.execPath, [join(STUDIO, 'node_modules', 'vite', bin), 'build', '--base=./', '--logLevel', 'warn'],
    { cwd: STUDIO, env: { ...process.env, ...BUILD_ENV, NODE_ENV: 'production' }, what: 'vite build' });
  if (!existsSync(join(STUDIO, 'dist', 'index.html'))) fail('vite build가 dist/index.html을 만들지 않았습니다.');
  return vitePkg.version;
}

// ── 4. 정리해서 vendor로 ──
function stageStudio(opts) {
  const dist = join(STUDIO, 'dist');
  const stage = join(VENDOR, '.rhwp-studio.tmp');
  rmSync(stage, { recursive: true, force: true });
  mkdirSync(stage, { recursive: true });

  const dropped = new Map();
  for (const rel of listFiles(dist)) {
    const rule = DROP.find(([test]) => test(rel));
    if (rule) {
      const d = dropped.get(rule[1]) || { files: 0, bytes: 0 };
      d.files++; d.bytes += lstatSync(join(dist, rel)).size;
      dropped.set(rule[1], d);
      continue;
    }
    mkdirSync(dirname(join(stage, rel)), { recursive: true });
    copyFileSync(join(dist, rel), join(stage, rel));
  }

  // index.html: PWA manifest 링크·서비스 워커 등록 스크립트 지우기
  const indexPath = join(stage, 'index.html');
  let html = readFileSync(indexPath, 'utf8');
  const patches = [
    [/<link\b[^>]*\brel=["']manifest["'][^>]*>\s*/gi, 'index.html: <link rel="manifest"> 제거'],
    [/<script\b[^>]*\bid=["']vite-plugin-pwa:register-sw["'][^>]*>\s*<\/script>\s*/gi, 'index.html: 서비스 워커 등록 스크립트(registerSW.js) 제거'],
  ];
  const applied = [];
  for (const [re, what] of patches) {
    const next = html.replace(re, '');
    if (next === html) fail(`index.html에서 지울 PWA 태그를 찾지 못했습니다 (${what}) — rhwp-studio 빌드 모양이 바뀌었는지 확인하세요.`);
    html = next;
    applied.push(what);
  }
  writeFileSync(indexPath, html);

  // hwpctrl 플러그인 청크가 빠졌는지, 서비스 워커 흔적이 없는지 (글자 파일 전부)
  const hwpctrl = listFiles(stage).find((rel) => /studio-plugin|hwpctrl/i.test(rel));
  if (hwpctrl) fail(`RHWP_WITHOUT_HWPCTRL=1인데 hwpctrl 플러그인 파일(${hwpctrl})이 남아 있습니다.`);
  for (const rel of listFiles(stage)) {
    if (!/\.(html|js|mjs|css|json|webmanifest)$/i.test(rel)) continue;
    const text = readFileSync(join(stage, rel), 'utf8');
    const hit = SW_PATTERNS.find((re) => re.test(text));
    if (hit) fail(`정리한 결과의 ${rel}에 서비스 워커 흔적(${hit})이 남아 있습니다 — 빌드를 멈춥니다.`);
  }

  // 글꼴 (선택)
  let fontInfo = null;
  if (opts.fonts) {
    const src = join(SRC, FONT_DIR);
    if (!existsSync(src)) fail(`${FONT_DIR} 폴더가 없습니다 — 캐시 폴더를 지우고 다시 실행하세요.`);
    const fonts = join(stage, 'fonts');
    mkdirSync(fonts, { recursive: true });
    const skipped = [];
    let n = 0;
    for (const f of readdirSync(src)) {
      if (!/\.(woff2|txt|md)$/i.test(f)) continue;
      if (FONT_EXCLUDE.test(f)) { skipped.push(f); continue; }
      copyFileSync(join(src, f), join(fonts, f));
      if (f.endsWith('.woff2')) n++;
    }
    fontInfo = { included: n, skipped, note: 'SIL OFL 1.1·GUST 글꼴만 (FONTS.md, SourceHanSerifK-OFL.txt 동봉)' };
  }

  // 라이선스
  copyFileSync(join(SRC, 'LICENSE'), join(stage, 'LICENSE'));
  copyFileSync(join(SRC, 'THIRD_PARTY_LICENSES.md'), join(stage, 'THIRD_PARTY_LICENSES.md'));
  mkdirSync(join(stage, 'licenses'), { recursive: true });
  const bundled = [];
  for (const name of BUNDLED_DEPS) {
    const dir = join(STUDIO, 'node_modules', ...name.split('/'));
    const pkg = readJson(join(dir, 'package.json'));
    const lic = readdirSync(dir).find((f) => /^(licen[sc]e|copying)(\.(md|txt))?$/i.test(f));
    if (!lic) fail(`${name}의 라이선스 파일을 찾지 못했습니다.`);
    const dest = `${name.replace('@', '').replace('/', '-')}-LICENSE.txt`;
    copyFileSync(join(dir, lic), join(stage, 'licenses', dest));
    bundled.push({ name, version: pkg.version, license: pkg.license, file: `licenses/${dest}` });
  }

  return {
    stage,
    removed: [...dropped].map(([what, d]) => ({ what, files: d.files, bytes: d.bytes })),
    patches: applied,
    fontInfo,
    bundled,
  };
}

function writeVersion(dir, info) {
  const files = listFiles(dir).filter((p) => p !== 'VERSION.json').map((p) => {
    const buf = readFileSync(join(dir, p));
    return { path: p, bytes: buf.length, sha256: sha256(buf) };
  });
  const totalBytes = files.reduce((s, f) => s + f.bytes, 0);
  const v = { ...info, files, fileCount: files.length, totalBytes };
  writeFileSync(join(dir, 'VERSION.json'), `${JSON.stringify(v, null, 2)}\n`);
  return v;
}

function stageSdk() {
  const stage = join(VENDOR, '.rhwp-editor-sdk.tmp');
  rmSync(stage, { recursive: true, force: true });
  mkdirSync(stage, { recursive: true });
  for (const f of SDK_FILES) copyFileSync(join(EDITOR_DIR, f), join(stage, f));
  // npm 꾸러미에 LICENSE가 없어 같은 저장소(rhwp, MIT)의 것을 싣는다
  copyFileSync(join(SRC, 'LICENSE'), join(stage, 'LICENSE'));
  return stage;
}

function replaceDir(stage, dest) {
  const old = `${dest}.old-${process.pid}`;
  try {
    if (existsSync(dest)) renameSync(dest, old);
    renameSync(stage, dest);
  } catch (e) {
    if (!existsSync(dest) && existsSync(old)) { try { renameSync(old, dest); } catch { /* 그대로 둠 */ } }
    fail(`${relative(ROOT, dest)}를 바꾸지 못했습니다 (${e.code || e.message}). Office Kit 앱이나 탐색기가 그 폴더를 쓰고 있으면 닫고 다시 실행하세요.`);
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

  const { wasmSha256 } = checkPackages();
  const builderSha = sha256(readFileSync(fileURLToPath(import.meta.url)));
  const current = !opts.force && upToDate(opts, builderSha);
  if (current) {
    log(`이미 준비되어 있습니다: vendor/rhwp-studio (rhwp ${TAG}, 파일 ${current.fileCount}개, ${mb(current.totalBytes)}) — 다시 만들려면 --force`);
    return;
  }

  mkdirSync(VENDOR, { recursive: true });
  for (const leftover of ['.rhwp-studio.tmp', '.rhwp-editor-sdk.tmp']) rmSync(join(VENDOR, leftover), { recursive: true, force: true });

  const source = ensureSource(opts);
  prepareWasm();
  installStudioDeps();
  const viteVersion = viteBuild();
  const staged = stageStudio(opts);
  const sdkStage = stageSdk();

  const studioVersion = writeVersion(staged.stage, {
    name: 'rhwp-studio',
    repo: REPO_URL,
    tag: TAG,
    commit: COMMIT,
    coreVersion: CORE_VERSION,
    coreWasmSha256: wasmSha256,
    editorSdkVersion: EDITOR_VERSION,
    vite: viteVersion,
    buildEnv: { ...BUILD_ENV, base: './' },
    fonts: opts.fonts,
    fontInfo: staged.fontInfo,
    bundledDependencies: staged.bundled,
    removed: staged.removed,
    patches: staged.patches,
    builtAt: new Date().toISOString(),
    builder: 'scripts/build-editor.mjs',
    builderSha256: builderSha,
  });
  const sdkVersion = writeVersion(sdkStage, {
    name: '@rhwp/editor',
    version: EDITOR_VERSION,
    note: 'createEditor(container, { studioUrl })의 studioUrl에는 앱이 띄운 vendor/rhwp-studio 주소(http/https)를 넘긴다',
    builtAt: studioVersion.builtAt,
  });

  replaceDir(staged.stage, OUT);
  replaceDir(sdkStage, SDK_OUT);

  const removedBytes = staged.removed.reduce((s, r) => s + r.bytes, 0);
  const lines = [
    '',
    '한글 편집기(rhwp-studio) 빌드 완료',
    `  원본   edwardkim/rhwp ${TAG} (${COMMIT.slice(0, 12)}, ${source === 'cache' ? '캐시 사용' : '새로 받음'}) · @rhwp/core ${CORE_VERSION} · vite ${viteVersion}`,
    `  결과   vendor/rhwp-studio — 파일 ${studioVersion.fileCount}개, ${mb(studioVersion.totalBytes)}`,
    `         vendor/rhwp-editor-sdk — 파일 ${sdkVersion.fileCount}개 (@rhwp/editor ${EDITOR_VERSION})`,
    `  뺀 것  ${staged.removed.map((r) => r.what.replace(/ \(.*\)$/, '')).join(', ')} — ${mb(removedBytes)}`,
    '  고침   index.html의 PWA manifest 링크·서비스 워커 등록 스크립트 제거 (서비스 워커 흔적 검사 통과)',
    `  글꼴   ${opts.fonts ? `오픈 라이선스 글꼴 ${staged.fontInfo.included}개 포함` : '넣지 않음 (시스템 글꼴 사용, --with-fonts로 포함)'}`,
    `  걸린 시간 ${elapsed(Date.now() - t0)}`,
  ];
  console.log(lines.join('\n'));
}

main().catch((e) => {
  if (e instanceof BuildError) console.error(`\n[editor] 오류: ${e.message}`);
  else console.error(`\n[editor] 예상하지 못한 오류: ${e.stack || e}`);
  process.exit(1);
});
