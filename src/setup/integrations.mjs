// integrations.mjs — Office Kit MCP 서버를 Claude Desktop / Claude Code / Codex에 등록하고,
// 명령줄(office-kit.cmd)을 설치한다.
// 모든 함수는 경로·실행기를 옵션으로 바꿀 수 있어서, 테스트는 실제 설정 파일을 건드리지 않는다.
//   opts = { env, home, runner, claudeDesktopConfig, codexConfig, shimPath }
import { execFile } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync, renameSync, copyFileSync, mkdirSync, rmSync, readdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, join, resolve, delimiter, isAbsolute } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));

export const SERVER_NAME = 'office-kit';
/** 저장소(또는 설치된 앱의 resources/app) 루트 */
export const APP_ROOT = resolve(here, '../..');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Windows 환경 변수는 대소문자를 가리지 않는다 (주입한 일반 객체도 같은 규칙으로 읽는다)
function getEnv(env, name) {
  if (env[name] !== undefined) return env[name];
  const key = Object.keys(env).find((k) => k.toLowerCase() === name.toLowerCase());
  return key ? env[key] : undefined;
}

// ── 실행 방법 ────────────────────────────────────────────

/** Electron 실행 파일인가 (개발용 electron.exe 또는 설치된 'Office Kit.exe') */
export function isElectronBinary(execPath = process.execPath) {
  const name = basename(execPath).toLowerCase();
  if (/^electron(\.exe)?$/.test(name) || /^office ?kit(\.exe)?$/.test(name)) return true;
  return execPath === process.execPath && !!process.versions.electron;
}

/**
 * MCP 서버를 띄우는 방법 → { command, args, env }
 * Electron이면 ELECTRON_RUN_AS_NODE=1로 Node처럼 실행한다 (사용자 PC에 Node가 없어도 됨).
 * appRoot는 열거되지 않는 속성으로 붙인다 (JSON으로 내보낼 때 섞이지 않게).
 */
export function mcpLaunchSpec({ execPath = process.execPath, appRoot = APP_ROOT } = {}) {
  // asar로 묶인 경우를 대비해 풀린 경로로 바꾼다 (기본 설정은 asar 끔)
  appRoot = appRoot.replace(/([\\/])app\.asar(?=[\\/]|$)/, '$1app.asar.unpacked');
  const electron = isElectronBinary(execPath);
  const spec = {
    command: execPath,
    args: [join(appRoot, 'src', 'mcp', 'server.mjs')],
    env: electron ? { ELECTRON_RUN_AS_NODE: '1' } : {},
  };
  Object.defineProperty(spec, 'appRoot', { value: appRoot, enumerable: false });
  return spec;
}

const specRoot = (spec) => spec.appRoot ?? dirname(dirname(dirname(spec.args[0])));
const sameSpec = (a, b) => !!a && !!b && a.command === b.command
  && JSON.stringify(a.args ?? []) === JSON.stringify(b.args ?? [])
  && JSON.stringify(a.env ?? {}) === JSON.stringify(b.env ?? {});

// 설정 파일에 넣을 모양 (env가 비면 뺀다)
function entryOf(spec) {
  const e = { command: spec.command, args: [...spec.args] };
  if (spec.env && Object.keys(spec.env).length) e.env = { ...spec.env };
  return e;
}

// ── 파일 쓰기 공통 ───────────────────────────────────────

/** 임시 파일에 쓴 뒤 바꿔치기. 다른 프로그램이 잡고 있으면 잠깐 기다렸다 다시 시도 */
async function atomicWrite(path, text) {
  mkdirSync(dirname(path), { recursive: true });
  const tmp = `${path}.${process.pid}.${Date.now()}.tmp`;
  writeFileSync(tmp, text, 'utf8');
  for (let i = 0; ; i++) {
    try { renameSync(tmp, path); return; } catch (e) {
      if (i >= 6 || !['EPERM', 'EBUSY', 'EACCES'].includes(e.code)) {
        rmSync(tmp, { force: true });
        throw new Error(`설정 파일을 저장하지 못했습니다: ${path} (${e.code ?? e.message})`);
      }
      await sleep(100 * (i + 1));
    }
  }
}

const BACKUP_KEEP = 10;

/**
 * 고칠 때마다 시각을 붙인 백업 (<파일>.office-kit-YYYYMMDD-HHMMSS.bak) → 백업 경로 | null(원본 없음)
 * Office Kit가 만든 백업만 최근 BACKUP_KEEP개 남기고 정리한다.
 */
export function backupFile(path, now = new Date()) {
  if (!existsSync(path)) return null;
  const p2 = (v) => String(v).padStart(2, '0');
  const stamp = `${now.getFullYear()}${p2(now.getMonth() + 1)}${p2(now.getDate())}-${p2(now.getHours())}${p2(now.getMinutes())}${p2(now.getSeconds())}`;
  let bak = `${path}.office-kit-${stamp}.bak`;
  for (let k = 2; existsSync(bak); k++) bak = `${path}.office-kit-${stamp}-${k}.bak`;
  copyFileSync(path, bak);
  const re = /^\.office-kit-(\d{8})-(\d{6})(?:-(\d+))?\.bak$/;
  const base = basename(path);
  let mine = [];
  try { mine = readdirSync(dirname(path)).filter((f) => f.startsWith(base) && re.test(f.slice(base.length))); } catch {}
  const order = (f) => { const m = re.exec(f.slice(base.length)); return [m[1], m[2], Number(m[3] ?? 1)]; };
  mine.sort((a, b) => { const x = order(a), y = order(b); return x[0].localeCompare(y[0]) || x[1].localeCompare(y[1]) || x[2] - y[2]; });
  for (const f of mine.slice(0, -BACKUP_KEEP)) rmSync(join(dirname(path), f), { force: true });
  return bak;
}

// ── Claude Desktop ───────────────────────────────────────

/**
 * claude_desktop_config.json 위치.
 * 보통 %APPDATA%\Claude. 스토어(MSIX)판은 %LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude 를 쓸 수 있다.
 */
export function claudeDesktopConfigPath(opts = {}) {
  if (opts.claudeDesktopConfig) return opts.claudeDesktopConfig;
  const env = opts.env ?? process.env;
  const appData = getEnv(env, 'APPDATA') ?? join(opts.home ?? homedir(), 'AppData', 'Roaming');
  const local = getEnv(env, 'LOCALAPPDATA') ?? join(opts.home ?? homedir(), 'AppData', 'Local');
  const main = join(appData, 'Claude', 'claude_desktop_config.json');
  if (existsSync(main)) return main;
  let pkgs = [];
  try { pkgs = readdirSync(join(local, 'Packages')).filter((d) => /^Claude_/i.test(d)); } catch {}
  const msix = pkgs.map((d) => join(local, 'Packages', d, 'LocalCache', 'Roaming', 'Claude'));
  return join(msix.find((d) => existsSync(join(d, 'claude_desktop_config.json'))) ?? msix.find((d) => existsSync(d)) ?? dirname(main),
    'claude_desktop_config.json');
}

/** JSON 설정 읽기. 없으면 빈 객체, 깨져 있으면 덮어쓰지 않고 오류 */
function readJsonConfig(path, label) {
  if (!existsSync(path)) return { data: {}, raw: null };
  const raw = readFileSync(path, 'utf8');
  const body = raw.replace(/^\uFEFF/, '');
  if (!body.trim()) return { data: {}, raw };
  let data;
  try { data = JSON.parse(body); } catch (e) {
    throw new Error(`${label} 설정 파일이 올바른 JSON이 아니라서 고치지 않았습니다: ${path}\n파일을 직접 고치거나 지운 뒤 다시 시도하세요. (${e.message})`);
  }
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error(`${label} 설정 파일 형식이 예상과 다릅니다(최상위가 객체가 아님): ${path}`);
  if (data.mcpServers !== undefined && (typeof data.mcpServers !== 'object' || data.mcpServers === null || Array.isArray(data.mcpServers))) {
    throw new Error(`${label} 설정 파일의 mcpServers 형식이 예상과 다릅니다: ${path}`);
  }
  return { data, raw };
}

// 원래 파일의 들여쓰기·줄바꿈·끝 줄바꿈을 따라 쓴다
function formatJsonLike(data, raw) {
  const indent = raw?.match(/\n([ \t]+)"/)?.[1] ?? 2;
  const eol = raw?.includes('\r\n') ? '\r\n' : '\n';
  let text = JSON.stringify(data, null, indent).replace(/\n/g, eol);
  if (raw == null || /\r?\n$/.test(raw)) text += eol;
  return text;
}

export async function connectClaudeDesktop(spec = mcpLaunchSpec(), opts = {}) {
  const configPath = claudeDesktopConfigPath(opts);
  const { data, raw } = readJsonConfig(configPath, 'Claude Desktop');
  const entry = entryOf(spec);
  if (sameSpec(data.mcpServers?.[SERVER_NAME], entry)) return { changed: false, configPath };
  data.mcpServers = { ...(data.mcpServers ?? {}), [SERVER_NAME]: entry };
  const backup = backupFile(configPath);
  await atomicWrite(configPath, formatJsonLike(data, raw));
  return { changed: true, configPath, backup, restartNeeded: true };
}

export async function disconnectClaudeDesktop(opts = {}) {
  const configPath = claudeDesktopConfigPath(opts);
  if (!existsSync(configPath)) return { changed: false, configPath };
  const { data, raw } = readJsonConfig(configPath, 'Claude Desktop');
  if (!data.mcpServers || !(SERVER_NAME in data.mcpServers)) return { changed: false, configPath };
  delete data.mcpServers[SERVER_NAME];
  const backup = backupFile(configPath);
  await atomicWrite(configPath, formatJsonLike(data, raw));
  return { changed: true, configPath, backup, restartNeeded: true };
}

// 화면·명령줄용 공통 필드: connected(연결됨), detail(한 줄 설명)
function withSummary(st, where) {
  const detail = st.error ? st.error
    : !st.configured ? ''
    : st.matches ? where
    : '다른 위치의 Office Kit로 등록돼 있습니다. 다시 연결하세요.';
  return { ...st, connected: !!st.configured, detail };
}

export function claudeDesktopStatus(spec = mcpLaunchSpec(), opts = {}) {
  const configPath = claudeDesktopConfigPath(opts);
  const out = { configPath, exists: existsSync(configPath), configured: false, matches: false };
  try {
    const entry = readJsonConfig(configPath, 'Claude Desktop').data.mcpServers?.[SERVER_NAME];
    out.configured = !!entry;
    out.matches = sameSpec(entry, entryOf(spec));
  } catch (e) { out.error = e.message; }
  return withSummary(out, configPath);
}

// ── 외부 명령 실행 ───────────────────────────────────────

/** PATH(+PATHEXT)에서 실행 파일 찾기 → 전체 경로 | null */
export function findExecutable(name, { env = process.env, extra = [] } = {}) {
  if (isAbsolute(name)) return existsSync(name) ? name : null;
  const exts = (getEnv(env, 'PATHEXT') ?? '.COM;.EXE;.BAT;.CMD').split(';').filter(Boolean);
  const dirs = (getEnv(env, 'PATH') ?? '').split(delimiter).filter(Boolean).map((d) => d.replace(/^"|"$/g, ''));
  for (const d of dirs) for (const e of exts) {
    const p = join(d, name + e.toLowerCase());
    if (existsSync(p)) return p;
  }
  return extra.find((p) => existsSync(p)) ?? null;
}

// cmd.exe /c 한 줄 인수: 큰따옴표로 감싸고 안의 " 는 "" 로
const cmdQuote = (a) => `"${String(a).replace(/"/g, '""')}"`;

/**
 * 기본 실행기: runner(cmd, args, { timeoutMs }) → { code, stdout, stderr, notFound? }
 * .cmd/.bat(npm 설치본)은 cmd.exe로 감싸 실행한다 (Node는 .cmd를 직접 실행하지 않음).
 */
export function runCommand(cmd, args = [], { timeoutMs = 60_000, env = process.env } = {}) {
  const home = homedir();
  const extra = cmd === 'claude' ? [join(home, '.local', 'bin', 'claude.exe'), join(getEnv(env, 'APPDATA') ?? '', 'npm', 'claude.cmd')]
    : cmd === 'codex' ? [join(getEnv(env, 'APPDATA') ?? '', 'npm', 'codex.cmd')] : [];
  const exe = findExecutable(cmd, { env, extra });
  if (!exe) return Promise.resolve({ code: null, notFound: true, stdout: '', stderr: `${cmd} 명령을 찾을 수 없습니다` });
  const batch = /\.(cmd|bat)$/i.test(exe);
  const file = batch ? (getEnv(env, 'ComSpec') ?? 'cmd.exe') : exe;
  const argv = batch ? ['/d', '/s', '/c', `"${[exe, ...args].map(cmdQuote).join(' ')}"`] : args;
  return new Promise((done) => {
    execFile(file, argv, { env, timeout: timeoutMs, windowsHide: true, windowsVerbatimArguments: batch, maxBuffer: 8 * 1024 * 1024, encoding: 'utf8' },
      (err, stdout, stderr) => done({
        code: err ? (typeof err.code === 'number' ? err.code : null) : 0,
        stdout: stdout ?? '', stderr: stderr ?? (err?.message ?? ''),
        ...(err?.killed ? { timedOut: true } : {}),
        ...(err?.code === 'ENOENT' ? { notFound: true } : {}),
      }));
  });
}

async function run(opts, cmd, args, timeoutMs) {
  const runner = opts.runner ?? runCommand;
  try { return await runner(cmd, args, { timeoutMs, env: opts.env ?? process.env }); } catch (e) {
    if (e?.code === 'ENOENT') return { code: null, notFound: true, stdout: '', stderr: e.message };
    throw e;
  }
}

// ── Claude Code (claude CLI) ─────────────────────────────

const normPath = (p) => (p ?? '').replace(/\\/g, '/').toLowerCase();
const NO_CLAUDE = 'Claude Code(claude 명령)를 찾을 수 없습니다. Claude Code를 설치하고 한 번 로그인한 뒤 다시 시도하세요.';
const SCOPE_NAME = { local: '이 폴더 전용(local)', project: '프로젝트 공유(project, .mcp.json)', user: '사용자 전체(user)' };

/**
 * `claude mcp get`의 Scope 줄 → 'local' | 'project' | 'user' | undefined
 * 예: "Local config (private to you in this project)" — 설명에 project가 들어 있으니 앞 단어로 가른다.
 */
export function claudeScopeOf(label = '') {
  return /^\s*local\b/i.test(label) ? 'local' : /^\s*project\b/i.test(label) ? 'project' : /^\s*user\b/i.test(label) ? 'user' : undefined;
}

/**
 * `claude mcp get office-kit` 결과 → { available, configured, matches, scope, command }
 * Claude Code가 실제로 쓰는 항목(우선순위 local > project > user 중 맨 앞)을 본다.
 */
export async function claudeCodeStatus(spec = mcpLaunchSpec(), opts = {}) {
  const r = await run(opts, 'claude', ['mcp', 'get', SERVER_NAME], 45_000);
  if (r.notFound) return { available: false, configured: false, matches: false, connected: false, detail: 'Claude Code가 설치돼 있지 않습니다' };
  const out = `${r.stdout}\n${r.stderr}`;
  if (r.code !== 0) {
    const missing = /No MCP server|not found/i.test(out);
    return withSummary({ available: true, configured: false, matches: false, ...(missing ? {} : { error: out.trim() || 'claude 명령 실행 실패' }) });
  }
  const field = (name) => out.match(new RegExp(`^\\s*${name}:\\s*(.*)$`, 'mi'))?.[1]?.trim() ?? '';
  const command = field('Command');
  const argsLine = field('Args');
  const scope = claudeScopeOf(field('Scope'));
  const matches = normPath(command) === normPath(spec.command)
    && spec.args.every((a) => normPath(argsLine).includes(normPath(a)))
    && Object.entries(spec.env ?? {}).every(([k, v]) => out.includes(`${k}=${v}`));
  const st = withSummary({ available: true, configured: true, matches, scope, command }, `범위: ${scope ?? '알 수 없음'}`);
  if (scope && scope !== 'user') st.detail = `${SCOPE_NAME[scope]} 설정이 먼저 쓰입니다. 다시 연결하면 사용자 전체(user)로 옮깁니다.`;
  return st;
}

// 한 범위의 office-kit 지우기. 실패하면 멈춘다 (지우지 못한 항목이 새 등록을 가린다)
async function removeClaudeCode(opts, scope) {
  const r = await run(opts, 'claude', ['mcp', 'remove', ...(scope ? ['--scope', scope] : []), SERVER_NAME], 30_000);
  if (r.notFound) throw new Error(NO_CLAUDE);
  if (r.code !== 0) throw new Error(`Claude Code에서 ${scope ? `${SCOPE_NAME[scope]} ` : ''}설정을 지우지 못했습니다: ${(r.stderr || r.stdout).trim()}`);
}

/**
 * user 범위로 등록한다. 앞서는 local/project 항목이나 다른 경로의 항목은 먼저 지우고,
 * 끝나면 Claude Code가 실제로 쓰는 항목이 이 앱을 가리키는지 다시 확인한다.
 */
export async function connectClaudeCode(spec = mcpLaunchSpec(), opts = {}) {
  let changed = false;
  let added = false;
  for (let round = 0; round < 5; round++) {
    const st = await claudeCodeStatus(spec, opts);
    if (!st.available) throw new Error(NO_CLAUDE);
    if (st.error) throw new Error(`Claude Code 연결 상태를 확인하지 못했습니다: ${st.error}`);
    if (st.configured && st.matches && st.scope === 'user') return { changed };
    if (st.configured) {
      if (added && st.scope === 'user') break; // 방금 넣은 것이 다르게 보이면 되풀이하지 않는다
      await removeClaudeCode(opts, st.scope);
      changed = true;
      continue;
    }
    if (added) break;
    // -e 는 값을 여러 개 받는 옵션이라 이름을 먼저 쓰고 '--' 로 끊는다
    const envArgs = Object.entries(spec.env ?? {}).flatMap(([k, v]) => ['-e', `${k}=${v}`]);
    const r = await run(opts, 'claude', ['mcp', 'add', '--scope', 'user', SERVER_NAME, ...envArgs, '--', spec.command, ...spec.args], 30_000);
    if (r.notFound) throw new Error(NO_CLAUDE);
    if (r.code !== 0) throw new Error(`Claude Code에 등록하지 못했습니다: ${(r.stderr || r.stdout).trim()}`);
    added = changed = true;
  }
  const last = await claudeCodeStatus(spec, opts);
  throw new Error(`Claude Code의 office-kit 설정이 이 앱을 가리키게 하지 못했습니다 (지금 쓰이는 설정: ${last.configured ? `${SCOPE_NAME[last.scope] ?? last.scope ?? '알 수 없는 범위'}, ${last.command || '명령 없음'}` : '없음'}). `
    + '터미널에서 claude mcp list 로 office-kit 항목을 확인하세요.');
}

/** 보이는 office-kit 항목을 범위마다 지운다 (local → project → user 순으로 드러난다) */
export async function disconnectClaudeCode(opts = {}) {
  let changed = false;
  for (let round = 0; round < 4; round++) {
    const st = await claudeCodeStatus(mcpLaunchSpec(), opts);
    if (!st.available) throw new Error(NO_CLAUDE);
    if (st.error) throw new Error(`Claude Code 연결 상태를 확인하지 못했습니다: ${st.error}`);
    if (!st.configured) return { changed };
    await removeClaudeCode(opts, st.scope);
    changed = true;
  }
  throw new Error('Claude Code에서 office-kit 설정을 모두 지우지 못했습니다. 터미널에서 claude mcp list 로 확인하세요.');
}

// ── Codex (~/.codex/config.toml) ─────────────────────────

export function codexConfigPath(opts = {}) {
  if (opts.codexConfig) return opts.codexConfig;
  const env = opts.env ?? process.env;
  return join(getEnv(env, 'CODEX_HOME') || join(opts.home ?? homedir(), '.codex'), 'config.toml');
}

// 표 바로 위에 두는 안내 줄. 지울 범위는 이 줄이 아니라 TOML 표 이름으로 정한다
// (Codex가 설정을 저장하면 이 줄은 사라지고, 새 표는 우리 표 뒤에 붙는다)
const MARK = '# office-kit: Office Kit 앱이 넣은 설정입니다. 앱의 설정 → 외부 연결에서 끄면 지워집니다.';
const ADDED_EOL = '[+eol]'; // 원래 파일 끝에 줄바꿈이 없어서 하나 붙였다는 표시 (지울 때 되돌림)
// 안내 줄 + 예전 판의 표시 줄(# >>> / # <<<). 따로 떨어져 남아 있어도 지운다
const MARK_RE = /^[ \t]*# (?:office-kit: Office Kit|>>> office-kit >>>|<<< office-kit <<<)/;
const OK_PATH = ['mcp_servers', SERVER_NAME];

/** TOML 기본 문자열 ("...") — 역슬래시·따옴표·제어 문자 이스케이프 */
export function tomlString(s) {
  return `"${String(s).replace(/[\\"\u0000-\u001f\u007f]/g, (c) => ({ '\\': '\\\\', '"': '\\"', '\b': '\\b', '\t': '\\t', '\n': '\\n', '\f': '\\f', '\r': '\\r' }[c]
    ?? `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`))}"`;
}
const tomlKey = (k) => (/^[A-Za-z0-9_-]+$/.test(k) ? k : tomlString(k));

// ── 작은 TOML 읽개: 줄 단위 문장과 그 위치만 안다 ───────
// 값은 문자열·배열·인라인 표만 풀고, 숫자·불·날짜는 원문(TomlRaw) 그대로 둔다.

class TomlRaw { constructor(raw) { this.raw = raw; } }
const isTable = (v) => !!v && typeof v === 'object' && !Array.isArray(v) && !(v instanceof TomlRaw);
const startsWithPath = (path, base) => base.every((p, k) => path[k] === p);

function putPath(obj, path, value) {
  if (!path.length) { if (isTable(value)) Object.assign(obj, value); return; }
  let o = obj;
  for (const p of path.slice(0, -1)) { if (!isTable(o[p])) o[p] = Object.create(null); o = o[p]; }
  o[path.at(-1)] = value;
}

/** → [{ kind: 'blank'|'comment'|'header'|'kv', start, end(다음 줄 시작), text?, path?, value? }] */
function scanToml(text) {
  const n = text.length;
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0; // BOM은 어느 문장에도 넣지 않는다
  const fail = (msg) => {
    const line = text.slice(0, i).split('\n').length;
    throw new Error(`config.toml ${line}번째 줄을 해석하지 못해 고치지 않았습니다 (${msg}). 파일을 직접 확인하세요.`);
  };
  const ws = () => { while (text[i] === ' ' || text[i] === '\t') i++; };
  const newline = () => {
    if (text[i] === '\n') { i++; return true; }
    if (text[i] === '\r' && text[i + 1] === '\n') { i += 2; return true; }
    return false;
  };
  const comment = () => { if (text[i] === '#') while (i < n && text[i] !== '\n') i++; };
  const lineEnd = () => { ws(); comment(); if (!newline() && i < n) fail('줄 끝에 알 수 없는 글자가 있습니다'); };
  const gap = () => { for (;;) { ws(); comment(); if (!newline()) return; } }; // 배열·인라인 표 안의 빈칸·줄바꿈·주석
  const ESC = { b: '\b', t: '\t', n: '\n', f: '\f', r: '\r', e: '\x1b', '"': '"', '\\': '\\' };

  function basic(multi) {
    i += multi ? 3 : 1;
    if (multi) newline(); // 여는 따옴표 바로 뒤 줄바꿈은 내용이 아니다
    let s = '';
    for (;;) {
      if (i >= n) fail('따옴표가 닫히지 않았습니다');
      const c = text[i];
      if (multi && text.startsWith('"""', i)) {
        let q = 3; while (q < 5 && text[i + q] === '"') q++; // """" → 끝 따옴표 앞의 " 는 내용
        s += '"'.repeat(q - 3); i += q; return s;
      }
      if (!multi && c === '"') { i++; return s; }
      if (c === '\\') {
        const d = text[i + 1];
        if (multi && /[ \t\r\n]/.test(d ?? '')) { // 줄 끝 역슬래시: 다음 글자까지 빈칸·줄바꿈을 건너뛴다
          let j = i + 1; while (text[j] === ' ' || text[j] === '\t') j++;
          if (text[j] !== '\n' && !(text[j] === '\r' && text[j + 1] === '\n')) fail('잘못된 역슬래시');
          i = j; while (/[ \t\r\n]/.test(text[i] ?? '')) i++;
          continue;
        }
        if (d === 'u' || d === 'U' || d === 'x') {
          const len = { u: 4, U: 8, x: 2 }[d];
          const hex = text.slice(i + 2, i + 2 + len);
          if (!new RegExp(`^[0-9A-Fa-f]{${len}}$`).test(hex)) fail('잘못된 이스케이프');
          try { s += String.fromCodePoint(parseInt(hex, 16)); } catch { fail('잘못된 이스케이프'); }
          i += 2 + len; continue;
        }
        if (!Object.hasOwn(ESC, d ?? '')) fail('잘못된 이스케이프');
        s += ESC[d]; i += 2; continue;
      }
      if (!multi && (c === '\n' || c === '\r')) fail('문자열 안에 줄바꿈이 있습니다');
      s += c; i++;
    }
  }

  function literal(multi) {
    const q = multi ? "'''" : "'";
    i += q.length;
    if (multi) newline();
    const j = text.indexOf(q, i);
    if (j < 0) fail('따옴표가 닫히지 않았습니다');
    let end = j;
    if (multi) while (end - j < 2 && text[end + 3] === "'") end++; // '''' → 끝 따옴표 앞의 ' 는 내용
    const s = text.slice(i, end);
    if (!multi && /[\r\n]/.test(s)) fail('문자열 안에 줄바꿈이 있습니다');
    i = end + q.length;
    return s;
  }

  const BARE = /[A-Za-z0-9_-]+/y;
  function key() {
    const parts = [];
    for (;;) {
      ws();
      if (text[i] === '"') parts.push(basic(false));
      else if (text[i] === "'") parts.push(literal(false));
      else {
        BARE.lastIndex = i;
        const m = BARE.exec(text);
        if (!m) fail('키 이름이 없습니다');
        parts.push(m[0]); i += m[0].length;
      }
      ws();
      if (text[i] !== '.') return parts;
      i++;
    }
  }

  const SCALAR = /[^,\]}#\r\n]*/y;
  function value(depth = 0) {
    if (depth > 64) fail('너무 깊게 겹쳐 있습니다');
    const c = text[i];
    if (c === '"') return basic(text.startsWith('"""', i));
    if (c === "'") return literal(text.startsWith("'''", i));
    if (c === '[') {
      i++;
      const arr = [];
      for (;;) {
        gap();
        if (text[i] === ']') { i++; return arr; }
        arr.push(value(depth + 1));
        gap();
        if (text[i] === ',') { i++; continue; }
        if (text[i] === ']') { i++; return arr; }
        fail('배열이 올바르지 않습니다');
      }
    }
    if (c === '{') {
      i++;
      const obj = Object.create(null);
      for (;;) {
        gap();
        if (text[i] === '}') { i++; return obj; }
        const k = key();
        if (text[i] !== '=') fail('= 가 없습니다');
        i++; gap();
        putPath(obj, k, value(depth + 1));
        gap();
        if (text[i] === ',') { i++; continue; }
        if (text[i] === '}') { i++; return obj; }
        fail('인라인 표가 올바르지 않습니다');
      }
    }
    SCALAR.lastIndex = i;
    const raw = SCALAR.exec(text)[0].trimEnd();
    if (!raw) fail('값이 없습니다');
    i += raw.length;
    return new TomlRaw(raw);
  }

  const out = [];
  let table = [];
  while (i < n) {
    const start = i;
    ws();
    if (i >= n || text[i] === '#' || text[i] === '\n' || text[i] === '\r') {
      const kind = text[i] === '#' ? 'comment' : 'blank';
      lineEnd();
      out.push({ kind, start, end: i, text: text.slice(start, i) });
    } else if (text[i] === '[') {
      const array = text[i + 1] === '[';
      i += array ? 2 : 1;
      const path = key();
      if (text[i] !== ']' || (array && text[i + 1] !== ']')) fail('표 이름이 올바르지 않습니다');
      i += array ? 2 : 1;
      lineEnd();
      table = path;
      out.push({ kind: 'header', start, end: i, path, array });
    } else {
      const k = key();
      if (text[i] !== '=') fail('= 가 없습니다');
      i++; ws();
      const v = value();
      lineEnd();
      out.push({ kind: 'kv', start, end: i, path: [...table, ...k], value: v });
    }
  }
  return out;
}

/**
 * config.toml 안의 office-kit 설정 찾기.
 * → { entry(command·args·env 값), configured, cuts(지울 범위), anchor(제자리에서 바꿀 [mcp_servers.office-kit] 표),
 *     locked(한 줄짜리 mcp_servers = {...} 안에 있음), inlineParent, addedEol }
 * 표의 범위 = 표 이름 줄부터, 다음 표 이름 전의 마지막 키까지. 뒤쪽 빈 줄·주석과 다른 표는 건드리지 않는다.
 */
function inspectCodexToml(text) {
  const stmts = scanToml(text);
  const entry = Object.create(null);
  const cuts = [];
  let anchor = null, ours = false, configured = false, locked = false, inlineParent = false, addedEol = false;
  // 표 앞의 빈 줄은 그 표 몫으로 함께 지운다
  const lead = (k) => { let j = k; while (j > 0 && stmts[j - 1].kind === 'blank') j--; return stmts[j].start; };
  stmts.forEach((s, k) => {
    if (s.kind === 'comment') {
      if (!MARK_RE.test(s.text)) return;
      const flag = s.text.trimEnd().endsWith(ADDED_EOL);
      addedEol ||= flag;
      cuts.push({ start: s.start, end: s.end, lead: s.start, mark: true, addedEol: flag, legacy: !s.text.trimStart().startsWith('# office-kit:') });
    } else if (s.kind === 'header') {
      ours = startsWithPath(s.path, OK_PATH);
      if (!ours) return;
      configured = true;
      let end = s.end;
      for (let j = k + 1; j < stmts.length && stmts[j].kind !== 'header'; j++) if (stmts[j].kind === 'kv') end = stmts[j].end;
      const cut = { start: s.start, end, lead: lead(k) };
      cuts.push(cut);
      if (!anchor && !s.array && s.path.length === OK_PATH.length) anchor = cut;
    } else if (s.kind === 'kv') {
      if (startsWithPath(s.path, OK_PATH)) {
        configured = true;
        putPath(entry, s.path.slice(OK_PATH.length), s.value);
        if (!ours) cuts.push({ start: s.start, end: s.end, lead: lead(k) }); // [mcp_servers] 아래 office-kit = {...} 같은 줄
      } else if (startsWithPath(OK_PATH, s.path)) {
        // mcp_servers = { ... } 처럼 위쪽 표를 한 줄로 쓴 경우: 그 안은 고칠 수 없다
        inlineParent = true;
        let v = s.value;
        for (const p of OK_PATH.slice(s.path.length)) v = isTable(v) ? v[p] : undefined;
        if (v !== undefined) { configured = locked = true; putPath(entry, [], v); }
      }
    }
  });
  return { entry, configured, cuts, anchor, locked, inlineParent, addedEol };
}

// 값으로 비교 (Codex가 표를 다시 쓰면 따옴표 종류·env 위치가 바뀐다)
function codexEntryMatches(e, spec) {
  if (typeof e.command !== 'string' || normPath(e.command) !== normPath(spec.command)) return false;
  if (!Array.isArray(e.args) || e.args.length !== spec.args.length
    || e.args.some((a, k) => typeof a !== 'string' || normPath(a) !== normPath(spec.args[k]))) return false;
  const want = Object.entries(spec.env ?? {});
  const have = e.env ?? {};
  if (!isTable(have) || Object.keys(have).length !== want.length || want.some(([k, v]) => have[k] !== v)) return false;
  return !(e.enabled instanceof TomlRaw && e.enabled.raw === 'false');
}

function renderCodexTable(spec, eol) {
  const env = Object.entries(spec.env ?? {});
  return [
    `[mcp_servers.${SERVER_NAME}]`,
    `command = ${tomlString(spec.command)}`,
    `args = [${spec.args.map(tomlString).join(', ')}]`,
    ...(env.length ? [`env = { ${env.map(([k, v]) => `${tomlKey(k)} = ${tomlString(v)}`).join(', ')} }`] : []),
  ].join(eol) + eol;
}

/**
 * config.toml 내용에 office-kit 표를 넣거나(spec) 빼고(null) 새 내용을 돌려준다.
 * office-kit 표·키와 안내 줄 밖의 내용은 바꾸지 않는다 (지운 표 앞의 빈 줄만 함께 지운다).
 * 이미 [mcp_servers.office-kit] 표가 있으면 그 자리에서 바꾸고, 없으면 파일 끝에 붙인다.
 */
export function editCodexToml(text, spec) {
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const t = inspectCodexToml(text);
  if (t.locked) {
    throw new Error(`config.toml의 office-kit 설정이 한 줄짜리 mcp_servers = { ... } 안에 있어 고치지 않았습니다. [mcp_servers.${SERVER_NAME}] 표로 바꾸거나 직접 지운 뒤 다시 시도하세요.`);
  }
  if (spec && t.inlineParent) {
    throw new Error('config.toml의 mcp_servers가 한 줄짜리 표(mcp_servers = { ... })라서 office-kit을 넣지 못했습니다. [mcp_servers.이름] 표 형식으로 바꾼 뒤 다시 시도하세요.');
  }
  const anchor = spec ? t.anchor : null;
  const markLine = (addedEol) => MARK + (addedEol ? ` ${ADDED_EOL}` : '') + eol;
  const keep = anchor ? t.cuts.find((c) => c.mark && c.end === anchor.start) : null; // 표 바로 위 안내 줄은 둔다
  const pieces = t.cuts.map((c) => (c === anchor ? { start: c.start, end: c.end, text: renderCodexTable(spec, eol) }
    : c === keep ? { start: c.start, end: c.end, text: c.legacy ? markLine(c.addedEol) : text.slice(c.start, c.end) } // 예전 표시 줄은 새 안내 줄로
    : { start: c.lead, end: c.end, text: '' }))
    .sort((a, b) => a.start - b.start);
  let out = '';
  let pos = 0;
  for (const p of pieces) { out += text.slice(pos, Math.max(pos, p.start)) + p.text; pos = Math.max(pos, p.end); }
  out += text.slice(pos);
  // 처음 붙일 때 더한 끝 줄바꿈은, 그 안내 줄부터 파일 끝까지를 모두 지웠을 때만 되돌린다
  const first = !anchor && t.cuts.find((c) => c.mark && c.addedEol);
  if (first) {
    let cur = first.start;
    for (const p of pieces) if (p.start <= cur && p.end > cur) cur = p.end;
    if (cur >= text.length) out = out.replace(/\r?\n$/, '');
  }
  if (spec && !anchor) {
    const addEol = out.replace(/^﻿/, '').length > 0 && !out.endsWith('\n');
    out += (addEol ? eol : '') + markLine(addEol) + renderCodexTable(spec, eol);
  }
  return out;
}

function readCodex(path) {
  if (!existsSync(path)) return null;
  const buf = readFileSync(path);
  const text = buf.toString('utf8');
  if (!Buffer.from(text, 'utf8').equals(buf)) throw new Error(`Codex 설정 파일이 UTF-8이 아니라서 고치지 않았습니다: ${path}`);
  return text;
}

export async function connectCodex(spec = mcpLaunchSpec(), opts = {}) {
  const configPath = codexConfigPath(opts);
  const text = readCodex(configPath) ?? '';
  const t = inspectCodexToml(text);
  if (t.configured && codexEntryMatches(t.entry, spec)) return { changed: false, configPath };
  const next = editCodexToml(text, spec);
  if (next === text) return { changed: false, configPath };
  const backup = backupFile(configPath);
  await atomicWrite(configPath, next);
  return { changed: true, configPath, backup };
}

export async function disconnectCodex(opts = {}) {
  const configPath = codexConfigPath(opts);
  const text = readCodex(configPath);
  if (text == null) return { changed: false, configPath };
  const next = editCodexToml(text, null);
  if (next === text) return { changed: false, configPath };
  const backup = backupFile(configPath);
  await atomicWrite(configPath, next);
  return { changed: true, configPath, backup };
}

/** → { configPath, exists, configured, matches, connected, detail, error? } */
export function codexStatus(spec = mcpLaunchSpec(), opts = {}) {
  const configPath = codexConfigPath(opts);
  const out = { configPath, exists: existsSync(configPath), configured: false, matches: false };
  try {
    const t = inspectCodexToml(readCodex(configPath) ?? '');
    out.configured = t.configured;
    out.matches = t.configured && codexEntryMatches(t.entry, spec);
    if (t.locked && !out.matches) out.error = '한 줄짜리 mcp_servers = { ... } 안의 office-kit 설정은 Office Kit가 고칠 수 없습니다. 직접 고치세요.';
  } catch (e) { out.error = e.message; }
  return withSummary(out, configPath);
}

// ── 명령줄 (office-kit.cmd) ──────────────────────────────

export function cliShimPath(opts = {}) {
  if (opts.shimPath) return opts.shimPath;
  const env = opts.env ?? process.env;
  const local = getEnv(env, 'LOCALAPPDATA') ?? join(opts.home ?? homedir(), 'AppData', 'Local');
  return join(local, 'Microsoft', 'WindowsApps', 'office-kit.cmd');
}

const SHIM_MARK = 'rem office-kit-shim';
// 배치 파일 안의 % 는 %% 로 (경로에 큰따옴표는 올 수 없다)
const batchEsc = (s) => String(s).replace(/%/g, '%%');

/**
 * office-kit.cmd 내용.
 * cmd는 배치 파일을 한 줄씩 "현재 코드 페이지"로 읽으므로, chcp 65001 앞은 ASCII만 쓰고
 * 한글이 들어갈 수 있는 경로 줄은 그 뒤에 둔다. 끝나면 원래 코드 페이지로 되돌린다.
 */
export function renderCliShim(spec = mcpLaunchSpec()) {
  const cli = join(specRoot(spec), 'src', 'cli.mjs');
  return [
    '@echo off',
    `${SHIM_MARK} (generated by Office Kit)`,
    'setlocal',
    'for /f "tokens=2 delims=:." %%a in (\'chcp\') do set "OFFICEKIT_CP=%%a"',
    'chcp 65001 >nul',
    'rem Office Kit 명령줄 — 앱 설정에서 다시 만들거나 지울 수 있습니다',
    ...Object.entries(spec.env ?? {}).map(([k, v]) => `set "${k}=${batchEsc(v)}"`),
    `"${batchEsc(spec.command)}" "${batchEsc(cli)}" %*`,
    'set "OFFICEKIT_EXIT=%ERRORLEVEL%"',
    'if defined OFFICEKIT_CP chcp %OFFICEKIT_CP% >nul',
    'exit /b %OFFICEKIT_EXIT%',
  ].join('\r\n') + '\r\n';
}

export async function installCliShim(spec = mcpLaunchSpec(), opts = {}) {
  const path = cliShimPath(opts);
  const content = renderCliShim(spec);
  if (existsSync(path)) {
    const old = readFileSync(path, 'utf8');
    if (old === content) return { changed: false, path };
    if (!old.includes(SHIM_MARK)) throw new Error(`같은 이름의 다른 파일이 이미 있어 덮어쓰지 않았습니다: ${path}`);
  }
  await atomicWrite(path, content); // BOM 없이 UTF-8 (BOM이 있으면 첫 줄 @echo off가 깨진다)
  return { changed: true, path };
}

export async function uninstallCliShim(opts = {}) {
  const path = cliShimPath(opts);
  if (!existsSync(path)) return { changed: false, path };
  if (!readFileSync(path, 'utf8').includes(SHIM_MARK)) throw new Error(`Office Kit가 만든 파일이 아니라서 지우지 않았습니다: ${path}`);
  rmSync(path, { force: true });
  return { changed: true, path };
}

export function cliShimStatus(spec = mcpLaunchSpec(), opts = {}) {
  const path = cliShimPath(opts);
  const env = opts.env ?? process.env;
  const dirs = (getEnv(env, 'PATH') ?? '').split(delimiter).map((d) => normPath(d.replace(/^"|"$/g, '')).replace(/\/$/, ''));
  const onPath = dirs.includes(normPath(dirname(path)));
  const out = { path, exists: false, ours: false, matches: false, onPath, connected: false, detail: '' };
  if (!existsSync(path)) return out;
  const text = readFileSync(path, 'utf8');
  const ours = text.includes(SHIM_MARK);
  const matches = text === renderCliShim(spec);
  const detail = !ours ? `같은 이름의 다른 파일이 있습니다: ${path}`
    : !matches ? '다른 위치의 Office Kit를 가리킵니다. 다시 설치하세요.'
    : onPath ? path : `${path} (이 폴더가 PATH에 없어 전체 경로로 실행해야 합니다)`;
  return { ...out, exists: true, ours, matches, connected: ours, detail };
}

// ── 한꺼번에 ─────────────────────────────────────────────

/** 연결 상태 → { claudeDesktop, claudeCode, codex, cli } */
export async function integrationStatus(spec = mcpLaunchSpec(), opts = {}) {
  const claudeCode = claudeCodeStatus(spec, opts).catch((e) => ({ available: false, configured: false, matches: false, error: e.message }));
  return {
    claudeDesktop: claudeDesktopStatus(spec, opts),
    claudeCode: await claudeCode,
    codex: codexStatus(spec, opts),
    cli: cliShimStatus(spec, opts),
  };
}
