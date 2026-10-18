// filex-source.mjs — 5단계: filex 소스를 고정 커밋으로 받아 우리 패치를 얹고 앱 페이지만 다시 빌드한다.
//
//   patches/office-editor/filex/*.patch (이름 순) → git apply → npm ci --ignore-scripts
//   → node scripts/build-app.mjs --editor <lock 그대로의 편집기> --x2t <x2t> --out <폴더> --no-zip
// Docker가 필요 없다 (편집기 파일은 4단계에서 확인한 릴리스의 것, x2t도 릴리스의 것).
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';

import { FILEX, FILEX_SPARSE_DIRS } from './pins.mjs';

export class StepError extends Error {}
const fail = (msg) => { throw new StepError(msg); };

// ── 실행 도우미 ──
export function run(cmd, args, { cwd, env, what, quiet = false }) {
  const r = spawnSync(cmd, args, { cwd, env: env || process.env, stdio: quiet ? ['ignore', 'pipe', 'pipe'] : 'inherit', windowsHide: true, encoding: 'utf8' });
  if (r.error?.code === 'ENOENT') fail(`${cmd}을(를) 찾을 수 없습니다 — 설치되어 있고 PATH에 있는지 확인하세요.`);
  if (r.error) fail(`${what} 실행 실패: ${r.error.message}`);
  if (r.status !== 0) fail(`${what}에 실패했습니다 (종료 코드 ${r.status}).${quiet && (r.stderr || r.stdout) ? `\n${(r.stderr || r.stdout).trim()}` : ''}`);
  return r;
}

function gitOut(args, cwd) {
  const r = spawnSync('git', args, { cwd, encoding: 'utf8', windowsHide: true });
  if (r.error?.code === 'ENOENT') fail('git을 찾을 수 없습니다 — Git을 설치하세요 (패치가 있을 때만 필요).');
  if (r.status !== 0) return null;
  return r.stdout.trim();
}

// npm은 Windows에서 .cmd라 셸 없이 못 띄운다 → npm-cli.js를 node로 직접 실행
function npmCli() {
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
const lf = (buf) => Buffer.from(buf.toString('utf8').replace(/\r\n/g, '\n'), 'utf8');

// ── 패치 ──
/**
 * 패치 폴더의 *.patch (이름 순). 해시는 줄바꿈을 LF로 맞춘 내용 (Windows 체크아웃이 CRLF로 바꿔도 같게).
 * @returns {{name:string, path:string, bytes:number, sha256:string, subject:string|null, date:string|null, touched:string[]}[]}
 */
export function listPatches(dir) {
  if (!dir || !existsSync(dir) || !statSync(dir).isDirectory()) return [];
  return readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.patch'))
    .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))
    .map((name) => {
      const path = join(dir, name);
      const raw = readFileSync(path);
      const text = lf(raw).toString('utf8');
      const touched = new Set();
      for (const m of text.matchAll(/^diff --git a\/(\S+) b\/(\S+)$/gm)) { touched.add(m[1]); touched.add(m[2]); }
      for (const m of text.matchAll(/^(?:---|\+\+\+) [ab]\/(\S+)/gm)) touched.add(m[1]);
      const subject = /^Subject: (?:\[PATCH[^\]]*\] )?(.+)$/m.exec(text)?.[1]?.trim() || null;
      const dateRaw = /^Date: (.+)$/m.exec(text)?.[1]?.trim();
      const d = dateRaw ? new Date(dateRaw) : null;
      const date = d && !Number.isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : null;
      if (!touched.size) fail(`${name}: 고칠 파일이 보이지 않습니다 (git diff 모양의 패치가 아님).`);
      return { name, path, bytes: raw.length, sha256: sha256(lf(raw)), subject, date, touched: [...touched].sort() };
    });
}

export function patchesSha256(patches) {
  if (!patches.length) return null;
  return sha256(patches.map((p) => `${p.name}\n${p.sha256}\n`).join(''));
}

// ── 소스 받기 ──
function verifyCommit(dir) {
  const head = gitOut(['rev-parse', 'HEAD^{commit}'], dir);
  const tag = gitOut(['rev-parse', `refs/tags/${FILEX.tag}^{commit}`], dir);
  if (head !== FILEX.commit || (tag && tag !== FILEX.commit)) {
    fail(`받은 filex ${FILEX.tag}의 커밋이 예상과 다릅니다 (받음 ${head || '없음'}, 태그 ${tag || '없음'}, 예상 ${FILEX.commit}).\n` +
      '원본 저장소의 태그가 바뀌었을 수 있어 멈춥니다. 새 커밋을 직접 확인한 뒤 scripts/office-editor/pins.mjs를 고치세요.');
  }
}

/** 패치가 고치는 파일의 맨 위 폴더 (성긴 받기에 더함). 맨 위 파일은 cone 모드가 늘 받는다 */
export function sparseDirsFor(patches) {
  const dirs = new Set(FILEX_SPARSE_DIRS);
  for (const p of patches) for (const f of p.touched) if (f.includes('/')) dirs.add(f.split('/')[0]);
  return [...dirs].sort();
}

/** filex 소스를 dir에 준비한다 (있으면 고정 커밋으로 되돌려 다시 씀). @returns 'cache' | 'clone' */
export function ensureFilexSource({ dir, sparseDirs, log }) {
  if (existsSync(join(dir, '.git'))) {
    const head = gitOut(['rev-parse', 'HEAD^{commit}'], dir);
    if (head === FILEX.commit) {
      log(`캐시된 filex 소스를 다시 씁니다: ${dir}`);
      run('git', ['-C', dir, 'sparse-checkout', 'set', '--cone', ...sparseDirs], { what: 'git sparse-checkout' });
      // 지난 패치·빌드가 남긴 것을 지운다 (node_modules는 둠)
      run('git', ['-C', dir, 'reset', '-q', '--hard', 'HEAD'], { what: 'git reset' });
      run('git', ['-C', dir, 'clean', '-q', '-fdx', '-e', 'node_modules'], { what: 'git clean' });
      verifyCommit(dir);
      return 'cache';
    }
    log(`캐시된 filex 소스의 커밋이 다릅니다 (${head || '알 수 없음'}) — 지우고 새로 받습니다.`);
    rmSync(dir, { recursive: true, force: true });
  }
  mkdirSync(dirname(dir), { recursive: true });
  const part = `${dir}.partial`;
  rmSync(part, { recursive: true, force: true });
  log(`filex ${FILEX.tag} 소스를 받습니다 (필요한 폴더만: ${sparseDirs.join(', ')}): ${FILEX.git}`);
  run('git', ['clone', '-q', '--depth', '1', '--branch', FILEX.tag, '--filter=blob:none', '--sparse',
    '-c', 'core.symlinks=false', '-c', 'core.autocrlf=false', '-c', 'advice.detachedHead=false', FILEX.git, part],
  { what: 'filex 소스 받기 (git clone — 인터넷 연결 확인)' });
  run('git', ['-C', part, 'sparse-checkout', 'set', '--cone', ...sparseDirs], { what: 'git sparse-checkout' });
  verifyCommit(part);
  renameSync(part, dir);
  return 'clone';
}

/** 패치를 이름 순으로 적용한다. CRLF로 바뀐 패치(Windows 체크아웃)는 LF로 맞춰 다시 해 본다 */
export function applyPatches({ dir, patches, scratch, log }) {
  for (const p of patches) {
    const tryApply = (file) => spawnSync('git', ['-C', dir, 'apply', '--check', '--whitespace=nowarn', file], { encoding: 'utf8', windowsHide: true });
    let file = p.path;
    let r = tryApply(file);
    if (r.status !== 0) {
      const raw = readFileSync(p.path);
      if (raw.includes('\r\n')) {
        mkdirSync(scratch, { recursive: true });
        file = join(scratch, basename(p.path));
        writeFileSync(file, lf(raw));
        r = tryApply(file);
      }
    }
    if (r.status !== 0) fail(`패치 ${p.name}을(를) filex ${FILEX.tag}(${FILEX.commit.slice(0, 12)})에 적용할 수 없습니다:\n${(r.stderr || r.stdout || '').trim()}`);
    run('git', ['-C', dir, 'apply', '--whitespace=nowarn', file], { what: `패치 ${p.name} 적용`, quiet: true });
    log(`  패치 적용: ${p.name}${p.subject ? ` — ${p.subject}` : ''}`);
  }
}

/** npm ci --ignore-scripts (잠금 파일이 같으면 건너뜀) */
export function installDeps({ dir, log }) {
  const lockFile = join(dir, 'package-lock.json');
  if (!existsSync(lockFile)) fail('filex 소스에 package-lock.json이 없습니다.');
  const want = sha256(Buffer.concat([readFileSync(lockFile), readFileSync(join(dir, 'package.json'))]));
  const stamp = join(dir, 'node_modules', '.officekit-stamp');
  if (existsSync(stamp) && readFileSync(stamp, 'utf8') === want && existsSync(join(dir, 'node_modules', 'esbuild', 'package.json'))) {
    log('filex 빌드 도구가 이미 설치되어 있습니다.');
    return;
  }
  log('filex 빌드 도구를 설치합니다 (npm ci, 설치 스크립트 실행 안 함)...');
  const env = { ...process.env };
  // 상위 npm run이 넘기는 설정이 하위 설치에 섞이지 않게
  for (const k of Object.keys(env)) if (/^npm_config_/i.test(k) && !/^npm_config_(cache|registry|userconfig|proxy|https_proxy|noproxy|strict_ssl|cafile)$/i.test(k)) delete env[k];
  run(process.execPath, [npmCli(), 'ci', '--ignore-scripts', '--no-audit', '--no-fund', '--loglevel=error'], { cwd: dir, env, what: 'filex 빌드 도구 설치 (npm ci)' });
  writeFileSync(stamp, want);
}

/** 패치를 얹은 build-app.mjs가 --source(수정판의 소스 주소, 패치 0003)를 받는지 */
export function buildAppTakesSource(dir) {
  try { return readFileSync(join(dir, 'scripts', 'build-app.mjs'), 'utf8').includes("'--source'"); } catch { return false; }
}

/**
 * filex scripts/build-app.mjs → outDir/ui (zip 없이).
 * source: 수정판의 소스 주소 — build-app.mjs가 받을 때만 넘긴다 (편집기 아래 법적 고지 줄에 들어감, AGPL 6조)
 * @returns {{ ui: string, source: string|null }}
 */
export function buildApp({ dir, editorDir, x2tDir, outDir, source = null, log }) {
  rmSync(outDir, { recursive: true, force: true });
  mkdirSync(outDir, { recursive: true });
  const args = [join(dir, 'scripts', 'build-app.mjs'), '--editor', editorDir, '--x2t', x2tDir, '--out', outDir, '--no-zip'];
  const withSource = !!source && buildAppTakesSource(dir);
  if (withSource) args.push('--source', source);
  log(`filex 앱 페이지를 다시 빌드합니다 (scripts/build-app.mjs --no-zip${withSource ? ` --source ${source}` : ''})...`);
  run(process.execPath, args, { cwd: dir, what: 'filex build-app.mjs' });
  const ui = join(outDir, 'ui');
  if (!existsSync(join(ui, 'index.html'))) fail('filex build-app.mjs가 ui/index.html을 만들지 않았습니다.');
  return { ui, source: withSource ? source : null };
}
