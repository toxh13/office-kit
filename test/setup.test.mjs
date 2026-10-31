// setup.test.mjs — 외부 연결(설정 파일 편집, 명령줄 파일)과 한글 보안 모듈 등록을 시험한다.
// 실제 설정 파일·레지스트리·네트워크는 건드리지 않는다 (모두 임시 폴더와 가짜 실행기로).
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, existsSync, rmSync, readdirSync, utimesSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, dirname } from 'node:path';
import { execFile } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import {
  mcpLaunchSpec, isElectronBinary, APP_ROOT, SERVER_NAME, backupFile, runCommand,
  claudeDesktopConfigPath, connectClaudeDesktop, disconnectClaudeDesktop, claudeDesktopStatus,
  connectClaudeCode, disconnectClaudeCode, claudeCodeStatus, claudeScopeOf,
  codexConfigPath, editCodexToml, tomlString, connectCodex, disconnectCodex, codexStatus,
  renderCliShim, installCliShim, uninstallCliShim, cliShimStatus, cliShimPath, packagedCopyOf, integrationStatus,
} from '../src/setup/integrations.mjs';
import {
  hancomSecurityStatus, setupHancomSecurity, removeHancomSecurity, downloadHancomSecurityModule,
  httpsDownload, pickModuleDll, peMachine, findHwpExe, MODULES_KEY, HANCOM_SECURITY_MODULE_URL,
} from '../src/setup/hancom.mjs';
import { HANCOM_SECURITY_MODULE } from '../src/core/com.mjs';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
let root;
before(() => { root = mkdtempSync(join(tmpdir(), 'officekit-setup-')); });
after(() => { rmSync(root, { recursive: true, force: true }); });
let n = 0;
const dir = (name = 'd') => { const d = join(root, `${name}-${n++}`); mkdirSync(d, { recursive: true }); return d; };

const EXE = 'C:\\Users\\홍 길동\\AppData\\Local\\Programs\\Office Kit\\Office Kit.exe';
const APP = 'C:\\Users\\홍 길동\\AppData\\Local\\Programs\\Office Kit\\resources\\app';
const spec = () => mcpLaunchSpec({ execPath: EXE, appRoot: APP });

// ── 실행 방법 ────────────────────────────────────────────

describe('mcpLaunchSpec', () => {
  test('설치된 Office Kit.exe → Electron을 Node처럼 실행', () => {
    const s = spec();
    assert.equal(s.command, EXE);
    assert.deepEqual(s.args, [join(APP, 'src', 'mcp', 'server.mjs')]);
    assert.deepEqual(s.env, { ELECTRON_RUN_AS_NODE: '1' });
    assert.equal(s.appRoot, APP);
    // appRoot는 JSON에 섞이지 않는다
    assert.deepEqual(Object.keys(JSON.parse(JSON.stringify(s))), ['command', 'args', 'env']);
  });

  test('개발용 electron.exe도 Electron으로 본다', () => {
    const s = mcpLaunchSpec({ execPath: 'D:\\x\\node_modules\\electron\\dist\\electron.exe', appRoot: 'D:\\x' });
    assert.deepEqual(s.env, { ELECTRON_RUN_AS_NODE: '1' });
    assert.ok(isElectronBinary('D:\\x\\electron.exe'));
    assert.ok(!isElectronBinary('C:\\Program Files\\nodejs\\node.exe'));
  });

  test('node.exe → env 없이 node로 실행, 기본 appRoot는 저장소 루트', () => {
    const s = mcpLaunchSpec({ execPath: 'C:\\Program Files\\nodejs\\node.exe' });
    assert.equal(s.command, 'C:\\Program Files\\nodejs\\node.exe');
    assert.deepEqual(s.env, {});
    assert.equal(APP_ROOT, repoRoot);
    assert.equal(s.args[0], join(repoRoot, 'src', 'mcp', 'server.mjs'));
  });

  test('asar 안의 경로는 풀린 경로로 바꾼다', () => {
    const s = mcpLaunchSpec({ execPath: EXE, appRoot: 'C:\\A\\resources\\app.asar' });
    assert.equal(s.args[0], join('C:\\A\\resources\\app.asar.unpacked', 'src', 'mcp', 'server.mjs'));
  });
});

// ── Claude Desktop (JSON) ────────────────────────────────

describe('Claude Desktop 설정', () => {
  test('파일이 없으면 만든다 (백업 없음)', async () => {
    const p = join(dir(), 'Claude', 'claude_desktop_config.json');
    const r = await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
    assert.equal(r.changed, true);
    assert.equal(r.backup, null);
    const j = JSON.parse(readFileSync(p, 'utf8'));
    assert.deepEqual(j.mcpServers[SERVER_NAME], { command: EXE, args: spec().args, env: { ELECTRON_RUN_AS_NODE: '1' } });
    assert.deepEqual(backupsIn(dirname(p)), []);
  });

  test('다른 키·다른 서버를 보존하고, 고칠 때마다 시각을 붙인 백업을 남긴다', async () => {
    const p = join(dir(), 'claude_desktop_config.json');
    const original = JSON.stringify({ mcpServers: { other: { command: 'x', args: ['y'] } }, globalShortcut: 'Ctrl+Space', preferences: { a: [1, 2] } }, null, 2);
    writeFileSync(p, original);
    const r1 = await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
    assert.equal(r1.changed, true);
    assert.match(r1.backup, /claude_desktop_config\.json\.office-kit-\d{8}-\d{6}\.bak$/);
    assert.equal(readFileSync(r1.backup, 'utf8'), original);
    const j = JSON.parse(readFileSync(p, 'utf8'));
    assert.deepEqual(j.mcpServers.other, { command: 'x', args: ['y'] });
    assert.equal(j.globalShortcut, 'Ctrl+Space');
    assert.deepEqual(j.preferences, { a: [1, 2] });
    assert.ok(j.mcpServers[SERVER_NAME]);

    // 같은 내용이면 다시 쓰지 않는다
    const before2 = readFileSync(p, 'utf8');
    const r2 = await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
    assert.equal(r2.changed, false);
    assert.equal(readFileSync(p, 'utf8'), before2);

    // 경로가 바뀌면 갱신하고, 바로 앞 내용을 새 백업으로 (처음 백업은 그대로)
    const s2 = mcpLaunchSpec({ execPath: 'D:\\New\\Office Kit.exe', appRoot: 'D:\\New\\resources\\app' });
    const r3 = await connectClaudeDesktop(s2, { claudeDesktopConfig: p });
    assert.equal(r3.changed, true);
    assert.notEqual(r3.backup, r1.backup);
    assert.equal(readFileSync(r3.backup, 'utf8'), before2);
    assert.equal(readFileSync(r1.backup, 'utf8'), original);
    assert.equal(JSON.parse(readFileSync(p, 'utf8')).mcpServers[SERVER_NAME].command, 'D:\\New\\Office Kit.exe');
    assert.equal(claudeDesktopStatus(s2, { claudeDesktopConfig: p }).matches, true);
    assert.equal(claudeDesktopStatus(spec(), { claudeDesktopConfig: p }).matches, false);
    const r4 = await disconnectClaudeDesktop({ claudeDesktopConfig: p });
    assert.ok(r4.backup && readFileSync(r4.backup, 'utf8').includes('D:\\\\New'), '해제할 때도 백업');
  });

  test('백업은 Office Kit가 만든 것만 최근 10개 남긴다', () => {
    const d = dir('bak');
    const p = join(d, 'config.toml');
    writeFileSync(p, 'x');
    writeFileSync(join(d, 'config.toml.bak'), 'user');      // 사용자가 만든 백업
    writeFileSync(join(d, 'other.toml.office-kit-20200101-000000.bak'), 'o'); // 다른 파일의 백업
    const made = [];
    for (let k = 0; k < 12; k++) made.push(backupFile(p, new Date(2026, 0, 1, 10, 0, k)));
    made.push(backupFile(p, new Date(2026, 0, 1, 10, 0, 11))); // 같은 초 → -2
    assert.match(made.at(-1), /20260101-100011-2\.bak$/);
    const left = readdirSync(d).filter((f) => f.startsWith('config.toml.office-kit-')).sort();
    assert.equal(left.length, 10);
    assert.deepEqual(left.slice(-2), ['config.toml.office-kit-20260101-100011-2.bak', 'config.toml.office-kit-20260101-100011.bak'].sort());
    assert.ok(!left.includes('config.toml.office-kit-20260101-100000.bak'), '가장 오래된 것부터 지운다');
    assert.ok(existsSync(join(d, 'config.toml.bak')));
    assert.ok(existsSync(join(d, 'other.toml.office-kit-20200101-000000.bak')));
    assert.equal(backupFile(join(d, 'none.toml')), null);
  });

  test('들여쓰기·CRLF·BOM 있는 파일도 읽고 형식을 따른다', async () => {
    const p = join(dir(), 'c.json');
    writeFileSync(p, '\uFEFF{\r\n    "mcpServers": {}\r\n}\r\n');
    await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
    const text = readFileSync(p, 'utf8');
    assert.ok(text.startsWith('{\r\n    "mcpServers"'));
    assert.ok(text.endsWith('}\r\n'));
  });

  test('깨진 JSON은 덮어쓰지 않고 한국어 오류', async () => {
    const p = join(dir(), 'c.json');
    const bad = '{ "mcpServers": { "a": 1, }';
    writeFileSync(p, bad);
    await assert.rejects(connectClaudeDesktop(spec(), { claudeDesktopConfig: p }), /올바른 JSON이 아니라서/);
    assert.equal(readFileSync(p, 'utf8'), bad);
    assert.deepEqual(backupsIn(dirname(p)), []);
    await assert.rejects(disconnectClaudeDesktop({ claudeDesktopConfig: p }), /JSON/);
    assert.match(claudeDesktopStatus(spec(), { claudeDesktopConfig: p }).error, /JSON/);
    writeFileSync(p, '[1,2]');
    await assert.rejects(connectClaudeDesktop(spec(), { claudeDesktopConfig: p }), /최상위가 객체가 아님/);
    writeFileSync(p, '{"mcpServers": []}');
    await assert.rejects(connectClaudeDesktop(spec(), { claudeDesktopConfig: p }), /mcpServers 형식/);
  });

  test('연결 해제는 office-kit만 지운다', async () => {
    const p = join(dir(), 'c.json');
    writeFileSync(p, JSON.stringify({ mcpServers: { other: { command: 'x' } }, theme: 'dark' }));
    await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
    assert.equal(claudeDesktopStatus(spec(), { claudeDesktopConfig: p }).configured, true);
    assert.equal((await disconnectClaudeDesktop({ claudeDesktopConfig: p })).changed, true);
    assert.deepEqual(JSON.parse(readFileSync(p, 'utf8')), { mcpServers: { other: { command: 'x' } }, theme: 'dark' });
    assert.equal((await disconnectClaudeDesktop({ claudeDesktopConfig: p })).changed, false);
    assert.equal((await disconnectClaudeDesktop({ claudeDesktopConfig: join(dir(), 'none.json') })).changed, false);
  });

  test('mcpServers가 없던 설정은 연결 → 해제하면 원본과 바이트까지 같다 (빈 mcpServers를 남기지 않음)', async () => {
    const p = join(dir(), 'c.json');
    for (const original of [
      '{\n  "coworkUserFilesPath": "C:\\\\Users\\\\홍 길동\\\\Documents",\n  "preferences": {\n    "menuBarEnabled": false\n  }\n}\n',
      '{\r\n    "preferences": {\r\n        "a": [\r\n            1\r\n        ]\r\n    }\r\n}',
    ]) {
      writeFileSync(p, original);
      await connectClaudeDesktop(spec(), { claudeDesktopConfig: p });
      assert.ok(JSON.parse(readFileSync(p, 'utf8')).mcpServers[SERVER_NAME]);
      assert.equal((await disconnectClaudeDesktop({ claudeDesktopConfig: p })).changed, true);
      assert.equal(readFileSync(p, 'utf8'), original);
    }
    // 다른 서버가 있으면 mcpServers는 그대로 둔다
    writeFileSync(p, JSON.stringify({ mcpServers: { [SERVER_NAME]: { command: 'x' }, other: { command: 'y' } } }));
    await disconnectClaudeDesktop({ claudeDesktopConfig: p });
    assert.deepEqual(JSON.parse(readFileSync(p, 'utf8')), { mcpServers: { other: { command: 'y' } } });
  });

  test('설정 파일 위치: %APPDATA% → 스토어판 LocalCache → 기본값', () => {
    const appData = dir('Roaming');
    const local = dir('Local');
    const env = { APPDATA: appData, LOCALAPPDATA: local };
    assert.equal(claudeDesktopConfigPath({ env }), join(appData, 'Claude', 'claude_desktop_config.json'));
    const msix = join(local, 'Packages', 'Claude_pzs8sxrjxfjjc', 'LocalCache', 'Roaming', 'Claude');
    mkdirSync(msix, { recursive: true });
    assert.equal(claudeDesktopConfigPath({ env }), join(msix, 'claude_desktop_config.json'));
    mkdirSync(join(appData, 'Claude'));
    writeFileSync(join(appData, 'Claude', 'claude_desktop_config.json'), '{}');
    assert.equal(claudeDesktopConfigPath({ env }), join(appData, 'Claude', 'claude_desktop_config.json'));
  });
});

// ── Claude Code (가짜 claude CLI) ────────────────────────

const SCOPE_LABEL = {
  local: 'Local config (private to you in this project)',
  project: 'Project config (shared via .mcp.json)',
  user: 'User config (available in all your projects)',
};

// claude 2.1.x 흉내: 범위마다 항목 하나, get은 local > project > user 중 앞선 것만 보여 준다
//   stuck: 지우기가 실패하는 범위, ghost: 지웠다고 해 놓고 남는 범위
function fakeClaude(state = {}) {
  const calls = [];
  state.entries ??= {};
  const runner = async (cmd, args) => {
    calls.push([cmd, ...args]);
    if (state.missing) { const e = new Error('spawn claude ENOENT'); e.code = 'ENOENT'; throw e; }
    const [, sub] = args;
    const scopeArg = args.includes('--scope') ? args[args.indexOf('--scope') + 1] : undefined;
    if (sub === 'get') {
      const scope = ['local', 'project', 'user'].find((s) => state.entries[s]);
      if (!scope) return { code: 1, stdout: '', stderr: `No MCP server named "${SERVER_NAME}".` };
      const { command, args: a, env = {} } = state.entries[scope];
      return { code: 0, stderr: '', stdout: `${SERVER_NAME}:\n  Scope: ${SCOPE_LABEL[scope]}\n  Status: ✔ Connected\n  Type: stdio\n  Command: ${command.replace(/\\/g, '/')}\n  Args: ${a.join(' ').replace(/\\/g, '/')}\n  Environment:\n${Object.entries(env).map(([k, v]) => `    ${k}=${v}\n`).join('')}` };
    }
    if (sub === 'add') {
      if (state.entries[scopeArg]) return { code: 1, stdout: '', stderr: `MCP server ${SERVER_NAME} already exists in ${scopeArg} config` };
      const dash = args.indexOf('--');
      const env = {};
      args.forEach((x, i) => { if (x === '-e') { const [k, v] = args[i + 1].split('='); env[k] = v; } });
      state.entries[scopeArg] = { command: args[dash + 1], args: args.slice(dash + 2), env };
      return { code: 0, stdout: 'Added', stderr: '' };
    }
    if (sub === 'remove') {
      const scope = scopeArg ?? ['local', 'project', 'user'].find((s) => state.entries[s]);
      if (!state.entries[scope] || state.stuck === scope) return { code: 1, stdout: '', stderr: `No MCP server named "${SERVER_NAME}" in ${scope} config` };
      if (state.ghost !== scope) delete state.entries[scope];
      return { code: 0, stdout: 'Removed', stderr: '' };
    }
    return { code: 1, stdout: '', stderr: 'unknown' };
  };
  return { runner, calls, state };
}
const mine = () => ({ command: EXE, args: spec().args, env: { ELECTRON_RUN_AS_NODE: '1' } });
const OLD = { command: 'node', args: ['old.mjs'] };
const subs = (f) => f.calls.filter((c) => c[2] !== 'get').map((c) => [c[2], c[c.indexOf('--scope') + 1]]);

describe('Claude Code 등록 (claude CLI)', () => {
  test('Scope 줄은 앞 단어로 가른다 (Local 설명 속 "project"에 속지 않음)', () => {
    assert.equal(claudeScopeOf(SCOPE_LABEL.local), 'local');
    assert.equal(claudeScopeOf(SCOPE_LABEL.project), 'project');
    assert.equal(claudeScopeOf(SCOPE_LABEL.user), 'user');
    assert.equal(claudeScopeOf('user'), 'user');
    assert.equal(claudeScopeOf(''), undefined);
  });

  test('없으면 user 범위로 add, -e 는 이름 뒤 / -- 앞', async () => {
    const f = fakeClaude();
    assert.deepEqual(await connectClaudeCode(spec(), { runner: f.runner }), { changed: true });
    const add = f.calls.find((c) => c[2] === 'add');
    assert.deepEqual(add, ['claude', 'mcp', 'add', '--scope', 'user', SERVER_NAME, '-e', 'ELECTRON_RUN_AS_NODE=1', '--', EXE, ...spec().args]);
    const st = await claudeCodeStatus(spec(), { runner: f.runner });
    assert.deepEqual(st, { available: true, configured: true, matches: true, scope: 'user', command: EXE.replace(/\\/g, '/'), connected: true, detail: '범위: user' });
  });

  test('이미 같으면 아무것도 안 하고, 다르면 지우고 다시 추가', async () => {
    const f = fakeClaude({ entries: { user: mine() } });
    assert.deepEqual(await connectClaudeCode(spec(), { runner: f.runner }), { changed: false });
    assert.deepEqual(subs(f), []);
    f.state.entries.user = { command: 'C:\\old\\Office Kit.exe', args: ['C:\\old\\server.mjs'], env: { ELECTRON_RUN_AS_NODE: '1' } };
    assert.deepEqual(await connectClaudeCode(spec(), { runner: f.runner }), { changed: true });
    assert.deepEqual(subs(f), [['remove', 'user'], ['add', 'user']]);
    assert.equal(f.state.entries.user.command, EXE);
  });

  test('local 범위의 옛 항목: project로 읽지 않고 local을 지운 뒤 user로 등록', async () => {
    const f = fakeClaude({ entries: { local: OLD } });
    const st = await claudeCodeStatus(spec(), { runner: f.runner });
    assert.equal(st.scope, 'local');
    assert.equal(st.matches, false);
    assert.match(st.detail, /이 폴더 전용\(local\) 설정이 먼저 쓰입니다/);
    assert.deepEqual(await connectClaudeCode(spec(), { runner: f.runner }), { changed: true });
    assert.deepEqual(subs(f), [['remove', 'local'], ['add', 'user']]);
    assert.deepEqual(Object.keys(f.state.entries), ['user']);
    assert.equal((await claudeCodeStatus(spec(), { runner: f.runner })).matches, true);
  });

  test('user가 이미 맞아도 앞서는 local 항목은 지운다 (add는 안 함)', async () => {
    const f = fakeClaude({ entries: { local: OLD, user: mine() } });
    assert.deepEqual(await connectClaudeCode(spec(), { runner: f.runner }), { changed: true });
    assert.deepEqual(subs(f), [['remove', 'local']]);
  });

  test('지우기 실패는 멈추고 add 하지 않는다', async () => {
    const f = fakeClaude({ entries: { local: OLD }, stuck: 'local' });
    await assert.rejects(connectClaudeCode(spec(), { runner: f.runner }), /Claude Code에서 이 폴더 전용\(local\) 설정을 지우지 못했습니다: No MCP server/);
    assert.ok(!f.calls.some((c) => c[2] === 'add'));
    await assert.rejects(disconnectClaudeCode({ runner: f.runner }), /지우지 못했습니다/);
  });

  test('지웠는데도 다른 항목이 계속 먼저 쓰이면 성공이라 하지 않는다', async () => {
    const f = fakeClaude({ entries: { local: OLD }, ghost: 'local' });
    await assert.rejects(connectClaudeCode(spec(), { runner: f.runner }), /가리키게 하지 못했습니다 \(지금 쓰이는 설정: 이 폴더 전용\(local\), node\)/);
    await assert.rejects(disconnectClaudeCode({ runner: f.runner }), /모두 지우지 못했습니다/);
  });

  test('node로 실행할 때는 -e 없이 등록', async () => {
    const f = fakeClaude();
    const s = mcpLaunchSpec({ execPath: 'C:\\node\\node.exe', appRoot: 'D:\\repo' });
    await connectClaudeCode(s, { runner: f.runner });
    assert.ok(!f.calls.find((c) => c[2] === 'add').includes('-e'));
  });

  test('연결 해제는 보이는 범위를 모두 지운다, claude 없음', async () => {
    const f = fakeClaude({ entries: { local: OLD, user: mine() } });
    assert.deepEqual(await disconnectClaudeCode({ runner: f.runner }), { changed: true });
    assert.deepEqual(subs(f), [['remove', 'local'], ['remove', 'user']]);
    assert.deepEqual(await disconnectClaudeCode({ runner: f.runner }), { changed: false });
    const none = fakeClaude({ missing: true });
    const st = await claudeCodeStatus(spec(), { runner: none.runner });
    assert.equal(st.available, false);
    assert.equal(st.connected, false);
    await assert.rejects(connectClaudeCode(spec(), { runner: none.runner }), /Claude Code\(claude 명령\)를 찾을 수 없습니다/);
    const notFound = { runner: async () => ({ code: null, notFound: true, stdout: '', stderr: '' }) };
    await assert.rejects(connectClaudeCode(spec(), notFound), /찾을 수 없습니다/);
  });

  test('add 실패·상태 확인 실패는 한국어 오류로', async () => {
    const runner = async (_c, args) => (args[1] === 'get' ? { code: 1, stdout: '', stderr: 'No MCP server named' } : { code: 1, stdout: '', stderr: 'boom' });
    await assert.rejects(connectClaudeCode(spec(), { runner }), /Claude Code에 등록하지 못했습니다: boom/);
    const broken = async () => ({ code: 1, stdout: '', stderr: 'config parse error' });
    await assert.rejects(connectClaudeCode(spec(), { runner: broken }), /연결 상태를 확인하지 못했습니다: config parse error/);
  });
});

// ── Codex (TOML) ─────────────────────────────────────────

// office-kit 표의 줄을 JSON으로 다시 읽어 값 확인 (TOML 기본 문자열의 \\ \" \uXXXX 는 JSON과 같다)
function blockValues(text) {
  const block = text.slice(text.indexOf('[mcp_servers.office-kit]'));
  const v = (key) => JSON.parse(block.match(new RegExp(`^${key} = (.*)$`, 'm'))[1]);
  return { command: v('command'), args: v('args') };
}
// 문자열을 임시 config.toml로 써서 상태 보기
function statusOf(text, s = spec()) {
  const p = join(dir('st'), 'config.toml');
  writeFileSync(p, text);
  return codexStatus(s, { codexConfig: p });
}
const backupsIn = (d) => readdirSync(d).filter((f) => /\.office-kit-\d{8}-\d{6}(-\d+)?\.bak$/.test(f));

describe('Codex 설정 (config.toml)', () => {
  const ORIGINAL = '# 내 설정\nmodel = "gpt-5"\napproval_policy = "on-request"\n\n[mcp_servers.other]\ncommand = "npx"\nargs = ["-y", "x"]\n\n[profiles.fast]\nmodel = "o4-mini"\n';

  test('윈도 경로 역슬래시를 올바르게 이스케이프', () => {
    const out = editCodexToml('', spec());
    assert.match(out, /^# office-kit: Office Kit/);
    assert.match(out, /^\[mcp_servers\.office-kit\]$/m);
    assert.ok(out.includes('command = "C:\\\\Users\\\\홍 길동\\\\AppData\\\\Local\\\\Programs\\\\Office Kit\\\\Office Kit.exe"'));
    assert.deepEqual(blockValues(out), { command: EXE, args: spec().args });
    assert.match(out, /^env = \{ ELECTRON_RUN_AS_NODE = "1" \}$/m);
    assert.ok(!out.includes('<<<'), '끝 표시 줄은 더 쓰지 않는다');
    assert.equal(tomlString('a"b\\c\n\t\u0001한글'), '"a\\"b\\\\c\\n\\t\\u0001한글"');
    assert.equal(statusOf(out).matches, true);
  });

  test('추가 → 제거하면 원본과 바이트까지 같다 (LF / CRLF / 끝 줄바꿈 없음 / BOM)', () => {
    for (const original of [ORIGINAL, ORIGINAL.replace(/\n/g, '\r\n'), ORIGINAL.trimEnd(), '', 'model = "x"', '\uFEFFmodel = "x"\n', 'a = 1\n\n']) {
      const added = editCodexToml(original, spec());
      assert.ok(added.startsWith(original), '앞부분 그대로');
      if (original.includes('\r\n')) assert.ok(!/[^\r]\n/.test(added), 'CRLF 유지');
      assert.equal(editCodexToml(added, null), original);
      assert.equal(editCodexToml(editCodexToml(added, spec()), null), original);
    }
  });

  test('경로가 바뀌면 같은 자리에서 고치고, 뒤에 붙은 사용자 내용은 보존', () => {
    const added = editCodexToml(ORIGINAL, spec()) + '\n[profiles.slow]\nmodel = "o3"\n';
    const s2 = mcpLaunchSpec({ execPath: 'D:\\새 폴더\\Office Kit.exe', appRoot: 'D:\\새 폴더\\resources\\app' });
    const updated = editCodexToml(added, s2);
    assert.equal(blockValues(updated).command, 'D:\\새 폴더\\Office Kit.exe');
    assert.ok(updated.startsWith(ORIGINAL));
    assert.ok(updated.endsWith('\n[profiles.slow]\nmodel = "o3"\n'));
    assert.equal(editCodexToml(updated, null), `${ORIGINAL}\n[profiles.slow]\nmodel = "o3"\n`);
  });

  test('우리 표 뒤에 Codex가 붙인 표([features])는 연결·해제 모두 남긴다 — 예전 끝 표시 줄 앞에 끼어든 경우 포함', async () => {
    const added = editCodexToml('model = "gpt-5"\n', spec());
    const legacy = 'model = "gpt-5"\n# >>> office-kit >>> (Office Kit가 자동으로 관리하는 부분입니다. 직접 고치지 마세요)\n'
      + added.slice(added.indexOf('[mcp_servers.office-kit]')) + '\n[features]\nx = true\n# <<< office-kit <<<\n';
    for (const text of [`${added}\n[features]\nx = true\n`, legacy]) {
      const p = join(dir('codex-features'), 'config.toml');
      writeFileSync(p, text);
      assert.equal(codexStatus(spec(), { codexConfig: p }).matches, true, '값이 같으면 연결됨');
      assert.deepEqual(await connectCodex(spec(), { codexConfig: p }), { changed: false, configPath: p });
      const r = await disconnectCodex({ codexConfig: p });
      assert.equal(r.changed, true);
      assert.equal(readFileSync(r.backup, 'utf8'), text, '해제 전에도 백업');
      assert.equal(readFileSync(p, 'utf8'), 'model = "gpt-5"\n\n[features]\nx = true\n');
      // 다른 경로로 다시 연결해도 [features]는 그대로
      writeFileSync(p, text);
      const s2 = mcpLaunchSpec({ execPath: 'D:\\x\\electron.exe', appRoot: 'D:\\x' });
      assert.equal((await connectCodex(s2, { codexConfig: p })).changed, true);
      const after = readFileSync(p, 'utf8');
      assert.ok(after.includes('\n[features]\nx = true\n'), after);
      assert.ok(!after.includes('<<<'));
      assert.equal(blockValues(after).command, 'D:\\x\\electron.exe');
      assert.equal(codexStatus(s2, { codexConfig: p }).matches, true);
    }
  });

  test('Codex가 다시 쓴 모양(작은따옴표 문자열, [mcp_servers.office-kit.env] 표)도 값으로 비교·정리', async () => {
    const s = spec();
    const rewritten = `model = "gpt-5"\n\n[mcp_servers.office-kit]\ncommand = '${EXE}'\nargs = ['${s.args[0]}']\n\n[mcp_servers.office-kit.env]\nELECTRON_RUN_AS_NODE = "1"\n\n[mcp_servers.mytool]\ncommand = "node"\nargs = ["m.js"]\n`;
    assert.equal(statusOf(rewritten).matches, true);
    assert.equal(statusOf(rewritten.replace('"1"', '"0"')).matches, false);
    assert.equal(statusOf(rewritten.replace('[mcp_servers.office-kit.env]\nELECTRON_RUN_AS_NODE = "1"\n', '')).matches, false);
    const p = join(dir('codex-rw'), 'config.toml');
    writeFileSync(p, rewritten);
    assert.equal((await connectCodex(s, { codexConfig: p })).changed, false);
    // 다른 경로로 연결: 제자리에서 한 표로 바꾸고 env 표는 지운다
    const s2 = mcpLaunchSpec({ execPath: 'D:\\x\\electron.exe', appRoot: 'D:\\x' });
    await connectCodex(s2, { codexConfig: p });
    assert.equal(readFileSync(p, 'utf8'), `model = "gpt-5"\n\n[mcp_servers.office-kit]\ncommand = "D:\\\\x\\\\electron.exe"\nargs = ["D:\\\\x\\\\src\\\\mcp\\\\server.mjs"]\nenv = { ELECTRON_RUN_AS_NODE = "1" }\n\n[mcp_servers.mytool]\ncommand = "node"\nargs = ["m.js"]\n`);
    await disconnectCodex({ codexConfig: p });
    assert.equal(readFileSync(p, 'utf8'), 'model = "gpt-5"\n\n[mcp_servers.mytool]\ncommand = "node"\nargs = ["m.js"]\n');
  });

  test('예전 표시 줄이 한쪽만 남아 있어도 멈추지 않고 정리한다', () => {
    const endOnly = 'model = "gpt-5"\n# <<< office-kit <<<\n'; // codex mcp remove office-kit 뒤 모양
    assert.equal(statusOf(endOnly).configured, false);
    assert.equal(statusOf(endOnly).error, undefined);
    assert.equal(editCodexToml(endOnly, null), 'model = "gpt-5"\n');
    const connected = editCodexToml(endOnly, spec());
    assert.ok(!connected.includes('<<<'));
    assert.equal(editCodexToml(connected, null), 'model = "gpt-5"\n');
    const startOnly = '# >>> office-kit >>> (Office Kit가 자동으로 관리하는 부분입니다. 직접 고치지 마세요)\n[mcp_servers.office-kit]\ncommand = "x"\n\n[mcp_servers.mytool]\ncommand = "node"\n';
    assert.equal(editCodexToml(startOnly, null), '\n[mcp_servers.mytool]\ncommand = "node"\n');
  });

  test('표시 줄 없이 들어 있는 office-kit도 찾아서 바꾸고 지운다 (여러 TOML 모양)', () => {
    for (const t of ['[mcp_servers.office-kit]\ncommand = "x"\n', '[ mcp_servers . "office-kit" ]\ncommand = "x"\n',
      '[mcp_servers]\nother = { command = "y" }\noffice-kit = { command = "x" }\n', 'mcp_servers.office-kit.command = "x"\n',
      '[mcp_servers.office-kit.env]\nA = "1"\n']) {
      assert.equal(statusOf(t).configured, true, t);
      assert.equal(statusOf(t).matches, false, t);
      const on = editCodexToml(t, spec());
      assert.equal(statusOf(on).matches, true, `${t}\n→\n${on}`);
      assert.equal((on.match(/office-kit(?!: Office Kit)/g) ?? []).length, 1, `office-kit 표는 하나만: ${on}`); // 안내 줄은 빼고 센다
      const off = editCodexToml(on, null);
      assert.ok(!off.includes('office-kit'), off);
      assert.ok(!editCodexToml(t, null).includes('office-kit'));
    }
    assert.equal(editCodexToml('[mcp_servers]\nother = { command = "y" }\noffice-kit = { command = "x" }\n', null), '[mcp_servers]\nother = { command = "y" }\n');
    // 이름만 비슷한 것은 우리 것이 아니다
    const similar = '[mcp_servers.office-kit2]\ncommand = "x"\n';
    assert.equal(statusOf(similar).configured, false);
    assert.ok(editCodexToml(similar, spec()).startsWith(similar));
    assert.equal(editCodexToml(similar, null), similar);
  });

  test('한 줄짜리 mcp_servers = { ... } 는 고치지 않고 알려 준다', () => {
    const inside = 'mcp_servers = { office-kit = { command = "x" } }\n';
    assert.throws(() => editCodexToml(inside, spec()), /한 줄짜리 mcp_servers/);
    assert.throws(() => editCodexToml(inside, null), /한 줄짜리 mcp_servers/);
    assert.match(statusOf(inside).error, /고칠 수 없습니다/);
    const other = 'mcp_servers = { other = { command = "y" } }\n';
    assert.throws(() => editCodexToml(other, spec()), /넣지 못했습니다/);
    assert.equal(editCodexToml(other, null), other);
  });

  test('TOML 해석: 여러 줄 문자열·배열·주석 속의 [표 이름]은 표로 보지 않는다', () => {
    const tricky = '\uFEFF# [mcp_servers.office-kit] 주석\nnotes = """\n[mcp_servers.office-kit]\n가짜 \\" 따옴표"""\nlit =\'\'\'\n[mcp_servers.office-kit]\'\'\'\n'
      + 'arr = [\n  "a", # ] [주석]\n  ["[x]"],\n]\nesc = "\\u00e9 \\t \\\\"\nwhen = 1979-05-27 07:32:00Z  # 날짜\n'
      + 'inline = { a = 1, b.c = "d", e = [1, 2] }\n\n[mcp_servers.office-kit]\ncommand = "x"   # 끝 주석\n';
    assert.equal(statusOf(tricky).configured, true);
    const on = editCodexToml(tricky, spec());
    assert.ok(on.startsWith(tricky.slice(0, tricky.indexOf('\n\n[mcp_servers.office-kit]\ncommand'))), '앞부분 그대로');
    assert.equal(statusOf(on).matches, true);
    const off = editCodexToml(tricky, null);
    assert.equal(off, tricky.slice(0, tricky.indexOf('\n\n[mcp_servers.office-kit]\ncommand') + 1));
  });

  test('해석할 수 없는 파일은 건드리지 않는다', async () => {
    const p = join(dir('bad-toml'), 'config.toml');
    for (const bad of ['a = "닫히지 않음\n', 'a = [1, 2\n[b]\n', '[a\nb = 1\n', 'a = \n', 'a\n']) {
      writeFileSync(p, bad);
      await assert.rejects(connectCodex(spec(), { codexConfig: p }), /해석하지 못해 고치지 않았습니다/, bad);
      await assert.rejects(disconnectCodex({ codexConfig: p }), /해석하지 못해/);
      assert.match(codexStatus(spec(), { codexConfig: p }).error, /번째 줄을 해석하지 못해/);
      assert.equal(readFileSync(p, 'utf8'), bad);
      assert.deepEqual(backupsIn(dirname(p)), []);
    }
  });

  test('파일 연결·해제·상태 (CODEX_HOME 존중, 고칠 때마다 백업)', async () => {
    const home = dir('codex-home');
    const opts = { env: { CODEX_HOME: home } };
    assert.equal(codexConfigPath(opts), join(home, 'config.toml'));
    assert.equal(codexConfigPath({ env: {}, home: 'C:\\Users\\u' }), join('C:\\Users\\u', '.codex', 'config.toml'));
    const p = join(home, 'config.toml');
    writeFileSync(p, ORIGINAL);
    const r1 = await connectCodex(spec(), opts);
    assert.equal(r1.changed, true);
    assert.equal(readFileSync(r1.backup, 'utf8'), ORIGINAL);
    assert.equal((await connectCodex(spec(), opts)).changed, false);
    assert.deepEqual(codexStatus(spec(), opts), { configPath: p, exists: true, configured: true, matches: true, connected: true, detail: p });
    assert.equal(codexStatus(mcpLaunchSpec({ execPath: 'D:\\x\\electron.exe', appRoot: 'D:\\x' }), opts).matches, false);
    const connected = readFileSync(p, 'utf8');
    const r2 = await disconnectCodex(opts);
    assert.equal(r2.changed, true);
    assert.notEqual(r2.backup, r1.backup);
    assert.equal(readFileSync(r2.backup, 'utf8'), connected);
    assert.equal(readFileSync(p, 'utf8'), ORIGINAL);
    assert.equal((await disconnectCodex(opts)).changed, false);
    assert.equal(backupsIn(home).length, 2);
    // 파일이 없으면 새로 만든다 (백업 없음)
    const fresh = { codexConfig: join(dir('codex2'), 'sub', 'config.toml') };
    const r3 = await connectCodex(spec(), fresh);
    assert.equal(r3.changed, true);
    assert.equal(r3.backup, null);
    assert.deepEqual(blockValues(readFileSync(fresh.codexConfig, 'utf8')).args, spec().args);
  });

  test('UTF-8이 아닌 파일은 건드리지 않는다', async () => {
    const p = join(dir(), 'config.toml');
    const bytes = Buffer.from([0x6d, 0x3d, 0x22, 0xc7, 0xd1, 0x22, 0x0a]); // CP949 "한"
    writeFileSync(p, bytes);
    await assert.rejects(connectCodex(spec(), { codexConfig: p }), /UTF-8이 아니라서/);
    assert.ok(readFileSync(p).equals(bytes));
  });

  // 설치된 codex가 있으면 실제 Codex 저장 방식과 함께 확인 (임시 CODEX_HOME — 사용자 설정은 안 건드림)
  test('실제 codex CLI와 함께: features enable · mcp add · mcp remove 뒤에도 정상', async (t) => {
    const home = dir('codex-real');
    const env = { ...process.env, CODEX_HOME: home };
    const codex = (...args) => runCommand('codex', args, { env, timeoutMs: 60_000 });
    const v = await codex('--version');
    if (v.notFound || v.code !== 0) { t.skip('codex 명령 없음'); return; }
    const p = join(home, 'config.toml');
    writeFileSync(p, 'model = "gpt-5"\n');
    const s = mcpLaunchSpec({ execPath: 'C:\\Users\\홍 길동\\Office Kit\\Office Kit.exe', appRoot: 'C:\\Users\\홍 길동\\Office Kit\\resources\\app' });
    const opts = { env };
    await connectCodex(s, opts);
    const got = JSON.parse((await codex('mcp', 'get', SERVER_NAME, '--json')).stdout);
    assert.deepEqual([got.transport.command, got.transport.args, got.transport.env], [s.command, s.args, s.env], 'Codex가 같은 값으로 읽는다');
    assert.equal((await codex('features', 'enable', 'apply_patch_freeform')).code, 0);
    assert.equal((await codex('mcp', 'add', 'mytool', '--', 'node', 'mytool.js')).code, 0);
    assert.equal(codexStatus(s, opts).matches, true, readFileSync(p, 'utf8'));
    assert.equal((await connectCodex(s, opts)).changed, false);
    assert.equal((await disconnectCodex(opts)).changed, true);
    const after = readFileSync(p, 'utf8');
    assert.ok(after.includes('[features]') && after.includes('[mcp_servers.mytool]'), after);
    assert.ok(!after.includes('office-kit'), after);
    assert.notEqual((await codex('mcp', 'get', SERVER_NAME)).code, 0);
    // Codex가 직접 지운 뒤에도 다시 연결된다
    await connectCodex(s, opts);
    assert.equal((await codex('mcp', 'remove', SERVER_NAME)).code, 0);
    assert.equal(codexStatus(s, opts).configured, false);
    assert.equal((await connectCodex(s, opts)).changed, true);
    assert.equal(JSON.parse((await codex('mcp', 'get', SERVER_NAME, '--json')).stdout).transport.command, s.command);
    assert.equal(JSON.parse((await codex('mcp', 'list', '--json')).stdout).length, 2);
  });
});

// ── 명령줄 (office-kit.cmd) ──────────────────────────────

const runCmd = (shim, args) => new Promise((done) => {
  const line = `""${shim}" ${args.map((a) => `"${a}"`).join(' ')}"`;
  execFile(process.env.ComSpec ?? 'cmd.exe', ['/d', '/s', '/c', line], { windowsVerbatimArguments: true, encoding: 'utf8', windowsHide: true },
    (err, stdout, stderr) => done({ code: err?.code ?? 0, stdout, stderr }));
});

describe('명령줄 office-kit.cmd', () => {
  test('내용: chcp 65001 앞은 ASCII, 경로는 따옴표, %* 전달, CRLF', () => {
    const text = renderCliShim(spec());
    const lines = text.split('\r\n');
    assert.equal(lines[0], '@echo off');
    assert.ok(!/[^\r]\n/.test(text));
    const chcp = lines.indexOf('chcp 65001 >nul');
    assert.ok(chcp > 0);
    assert.ok(lines.slice(0, chcp).every((l) => /^[\x00-\x7f]*$/.test(l)), 'chcp 앞은 ASCII만');
    assert.ok(lines.includes('set "ELECTRON_RUN_AS_NODE=1"'));
    // 실행 줄이 마지막 줄 (그 뒤로 cmd가 파일을 다시 읽지 않게, 뒷정리도 같은 줄에서)
    assert.equal(lines.at(-1), '');
    const run = lines.at(-2);
    assert.ok(run.startsWith(`"${EXE}" "${join(APP, 'src', 'cli.mjs')}" %* && ((goto) 2>nul & chcp %OFFICEKIT_CP% >nul & cmd /d /c exit 0) || `), run);
    assert.ok(run.endsWith('((goto) 2>nul & chcp %OFFICEKIT_CP% >nul & cmd /d /c exit 1)'), run);
    assert.ok(lines.indexOf(run) > chcp);
    // % 는 %% 로
    assert.ok(renderCliShim(mcpLaunchSpec({ execPath: 'C:\\100%\\node.exe', appRoot: 'C:\\100%' })).includes('"C:\\100%%\\node.exe"'));
  });

  test('설치·상태·제거, 남의 파일은 건드리지 않음', async () => {
    const d = dir('WindowsApps');
    const opts = { shimPath: join(d, 'office-kit.cmd'), env: { PATH: `C:\\Windows;${d}\\` } };
    assert.deepEqual(cliShimStatus(spec(), opts), { path: opts.shimPath, exists: false, ours: false, matches: false, onPath: true, connected: false, detail: '' });
    assert.equal((await installCliShim(spec(), opts)).changed, true);
    assert.equal((await installCliShim(spec(), opts)).changed, false);
    assert.ok(!readFileSync(opts.shimPath).subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf])), 'BOM 없음');
    assert.equal(cliShimStatus(spec(), opts).matches, true);
    assert.equal(cliShimStatus(spec(), opts).connected, true);
    assert.match(cliShimStatus(spec(), { ...opts, env: { PATH: 'C:\\Windows' } }).detail, /PATH에 없어/);
    assert.equal((await uninstallCliShim(opts)).changed, true);
    assert.ok(!existsSync(opts.shimPath));
    assert.equal((await uninstallCliShim(opts)).changed, false);
    writeFileSync(opts.shimPath, '@echo off\r\necho mine\r\n');
    await assert.rejects(installCliShim(spec(), opts), /다른 파일이 이미 있어/);
    await assert.rejects(uninstallCliShim(opts), /Office Kit가 만든 파일이 아니라서/);
    assert.equal(readFileSync(opts.shimPath, 'utf8'), '@echo off\r\necho mine\r\n');
    const status = cliShimStatus(spec(), { ...opts, env: { PATH: 'C:\\Windows' } });
    assert.equal(status.ours, false);
    assert.equal(status.onPath, false);
  });

  test('실제로 실행: 한글·공백 경로, 인수 전달, 종료 코드', { skip: process.platform !== 'win32' }, async () => {
    const app = join(dir('앱 폴더 한글'), 'app');
    mkdirSync(join(app, 'src'), { recursive: true });
    writeFileSync(join(app, 'src', 'cli.mjs'),
      'console.log(JSON.stringify({ args: process.argv.slice(2), env: process.env.ELECTRON_RUN_AS_NODE ?? null, msg: "한글 출력" }));\nprocess.exit(process.argv.includes("fail") ? 7 : 0);\n');
    const opts = { shimPath: join(dir('bin 폴더'), 'office-kit.cmd') };
    await installCliShim(mcpLaunchSpec({ execPath: process.execPath, appRoot: app }), opts);
    const ok = await runCmd(opts.shimPath, ['read', 'C:\\문서 폴더\\보고서.hwp']);
    assert.equal(ok.code, 0, ok.stderr);
    assert.deepEqual(JSON.parse(ok.stdout.trim().split(/\r?\n/).pop()), { args: ['read', 'C:\\문서 폴더\\보고서.hwp'], env: null, msg: '한글 출력' });
    const bad = await runCmd(opts.shimPath, ['fail']);
    assert.equal(bad.code, 1, '실패는 1로 (cli.mjs도 0/1만 돌려준다)');
  });

  // cmd는 실행 중인 배치 파일을 명령마다 다시 읽는다 → setup cli [--remove]가 자기 파일을 지우거나 바꿔도 깨지면 안 된다
  test('실행 중에 자기 파일이 지워지거나 바뀌어도: 오류 없음, 종료 코드, 코드 페이지·환경 변수 원래대로', { skip: process.platform !== 'win32' }, async () => {
    const app = join(dir('앱 폴더'), 'app');
    mkdirSync(join(app, 'src'), { recursive: true });
    const shim = join(dir('bin 폴더'), 'office-kit.cmd');
    writeFileSync(join(app, 'src', 'cli.mjs'), [
      "import { rmSync, writeFileSync } from 'node:fs';",
      `const shim = ${JSON.stringify(shim)};`,
      "const [op, code] = process.argv.slice(2);",
      "if (op === 'remove') rmSync(shim);",
      // 다른 위치·새 형식으로 다시 만든 경우처럼 길이가 다른 내용
      "if (op === 'replace') writeFileSync(shim, '@echo off\\r\\nrem office-kit-shim (new)\\r\\nrem ' + 'x'.repeat(300) + '\\r\\necho REPLACED-RAN\\r\\n');",
      "console.log(`ran ${op} env=${process.env.OFFICEKIT_SHIM_TEST}`);",
      'process.exit(Number(code));',
    ].join('\n'));
    // env가 있는 실행 방법 (Electron의 ELECTRON_RUN_AS_NODE처럼): 부른 쪽으로 새지 않아야 한다
    const s = { ...mcpLaunchSpec({ execPath: process.execPath, appRoot: app }), env: { OFFICEKIT_SHIM_TEST: '1' } };
    const runLine = (line) => new Promise((done) => execFile(process.env.ComSpec ?? 'cmd.exe', ['/d', '/v:on', '/s', '/c', `"${line}"`],
      { windowsVerbatimArguments: true, encoding: 'utf8', windowsHide: true }, (err, stdout, stderr) => done({ code: err?.code ?? 0, stdout, stderr })));
    for (const [op, code, want] of [['remove', 0, 0], ['remove', 3, 1], ['replace', 0, 0], ['replace', 1, 1], ['none', 0, 0]]) {
      await installCliShim(s, { shimPath: shim });
      const r = await runLine(`chcp 437 >nul & chcp & call "${shim}" ${op} ${code} & echo rc=!errorlevel! & chcp`
        + ' & (if defined OFFICEKIT_SHIM_TEST echo LEAK-ENV) & (if defined OFFICEKIT_CP echo LEAK-CP)');
      const label = `${op} ${code}: ${r.stdout}${r.stderr}`;
      assert.equal(r.stderr.trim(), '', label);
      assert.match(r.stdout, new RegExp(`ran ${op} env=1`), label);
      assert.match(r.stdout, new RegExp(`rc=${want}\\s`), label);
      assert.doesNotMatch(r.stdout, /REPLACED-RAN|LEAK-/, label);
      const pages = r.stdout.split(/\r?\n/).filter((l) => /:\s*\d+\s*$/.test(l)).map((l) => l.match(/(\d+)\s*$/)[1]);
      assert.equal(pages.length, 2, label);
      assert.equal(pages[1], pages[0], `코드 페이지를 되돌린다 — ${label}`);
      assert.equal(existsSync(shim), op !== 'remove');
    }
  });

  test('예전 형식(또는 다른 위치)의 명령줄 파일은 다시 연결 필요로 보이고, 다시 연결하면 바뀐다', async () => {
    const d = dir('WindowsApps');
    const opts = { shimPath: join(d, 'office-kit.cmd'), env: { PATH: d } };
    const old = renderCliShim(spec()).replace(/ && \(\(goto\).*$/m, '\r\nset "OFFICEKIT_EXIT=%ERRORLEVEL%"\r\nexit /b %OFFICEKIT_EXIT%');
    writeFileSync(opts.shimPath, old);
    const st = cliShimStatus(spec(), opts);
    assert.equal(st.connected, true);
    assert.equal(st.matches, false);
    assert.match(st.detail, /예전 판이거나 다른 위치/);
    assert.equal((await installCliShim(spec(), opts)).changed, true);
    assert.equal(cliShimStatus(spec(), opts).matches, true);
  });

  test('MSIX 앱 안에서 써서 그 앱 전용 폴더로 옮겨졌으면 알린다', async () => {
    const local = dir('Local');
    const shimPath = join(local, 'Microsoft', 'WindowsApps', 'office-kit.cmd');
    const copy = join(local, 'Packages', 'Claude_pzs8sxrjxfjjc', 'LocalCache', 'Local', 'Microsoft', 'WindowsApps', 'office-kit.cmd');
    const opts = { env: { LOCALAPPDATA: local, PATH: '' } };
    assert.equal(cliShimPath(opts), shimPath);
    // 다른 앱 폴더에 옛 사본만 있으면(오래전에 쓴 것) 알리지 않는다
    mkdirSync(dirname(copy), { recursive: true });
    writeFileSync(copy, 'old');
    utimesSync(copy, new Date(2020, 0, 1), new Date(2020, 0, 1));
    mkdirSync(join(local, 'Packages', 'Other.App_abc'), { recursive: true });
    let r = await installCliShim(spec(), opts);
    assert.equal(r.changed, true);
    assert.equal(r.warning, undefined);
    // 방금 쓴 파일이 앱 전용 폴더에 생겼다 (가상화된 쓰기를 흉내)
    rmSync(shimPath);
    writeFileSync(copy, renderCliShim(spec()));
    r = await installCliShim(spec(), opts);
    assert.equal(r.packaged, 'Claude');
    assert.match(r.warning, /Claude 앱 안에서 실행돼 .*일반 터미널/);
    assert.equal(packagedCopyOf(join(dir('elsewhere'), 'x.cmd'), 0, opts), null, 'LOCALAPPDATA 밖은 보지 않는다');
  });
});

// ── 한꺼번에 상태 ────────────────────────────────────────

test('integrationStatus는 네 가지를 모두 돌려준다', async () => {
  const d = dir('all');
  const opts = {
    claudeDesktopConfig: join(d, 'claude.json'), codexConfig: join(d, 'config.toml'), shimPath: join(d, 'office-kit.cmd'),
    runner: fakeClaude().runner, env: { PATH: d },
  };
  await connectClaudeDesktop(spec(), opts);
  await connectCodex(spec(), opts);
  const st = await integrationStatus(spec(), opts);
  assert.equal(st.claudeDesktop.matches, true);
  assert.equal(st.codex.matches, true);
  assert.equal(st.claudeCode.configured, false);
  assert.equal(st.cli.exists, false);
  assert.equal(st.claudeDesktop.connected, true);
  assert.equal(st.codex.connected, true);
  assert.equal(st.cli.onPath, true);
});

// ── 한글 보안 모듈 ───────────────────────────────────────

// 머신 종류만 맞춘 가짜 PE 파일
function fakePe(machine, tag = '') {
  const b = Buffer.alloc(512);
  b.write('MZ', 0, 'latin1');
  b.writeUInt32LE(0x80, 0x3c);
  b.write('PE\0\0', 0x80, 'latin1');
  b.writeUInt16LE(machine, 0x84);
  b.write(tag, 0x100, 'utf8');
  return b;
}
const X86 = 0x14c;
const X64 = 0x8664;

function fakeReg(initial = {}) {
  const values = new Map(Object.entries(initial));
  const calls = [];
  return {
    values, calls,
    async get(key, name = '') { return values.get(`${key}|${name}`) ?? null; },
    async set(key, name, data) { calls.push(['set', key, name, data]); values.set(`${key}|${name}`, data); },
    async remove(key, name) { calls.push(['remove', key, name]); values.delete(`${key}|${name}`); },
  };
}
const CLSID = '{2291CF00-64A1-4877-A9B4-68CFE89612D6}';
function fakeHancom(machine = X86) {
  const bin = join(dir('Hnc'), 'Office 2024', 'HOffice130', 'Bin');
  mkdirSync(bin, { recursive: true });
  const hwp = join(bin, 'Hwp.exe');
  writeFileSync(hwp, fakePe(machine));
  const reg = fakeReg({
    'HKCR\\HWPFrame.HwpObject\\CLSID|': CLSID,
    [`HKCR\\WOW6432Node\\CLSID\\${CLSID}\\LocalServer32|`]: `${hwp} -Automation`,
  });
  return { hwp, reg, env: { ProgramFiles: dir('empty') } };
}

describe('한글 자동화 보안 모듈', () => {
  test('레지스트리 값 이름은 com.mjs의 HANCOM_SECURITY_MODULE과 같다', () => {
    assert.equal(HANCOM_SECURITY_MODULE, 'FilePathCheckerModule');
    assert.equal(MODULES_KEY, 'HKCU\\Software\\HNC\\HwpAutomation\\Modules');
    assert.match(HANCOM_SECURITY_MODULE_URL, /^https:\/\/github\.com\/hancom-io\/devcenter-archive\/raw\/[0-9a-f]{40}\/hwp-automation\//, '바뀔 수 있는 브랜치가 아니라 커밋으로 고정');
  });

  test('PE 머신 종류 읽기', () => {
    assert.equal(peMachine(fakePe(X86)), 'x86');
    assert.equal(peMachine(fakePe(X64)), 'x64');
    assert.equal(peMachine(Buffer.from('not a pe file at all, just some text here....................................')), null);
    const p = join(dir(), 'a.dll');
    writeFileSync(p, fakePe(X64));
    assert.equal(peMachine(p), 'x64');
    assert.equal(peMachine(join(root, 'nope.dll')), null);
  });

  test('한컴 없음', async () => {
    const st = await hancomSecurityStatus({ reg: fakeReg(), env: { ProgramFiles: dir('pf'), 'ProgramFiles(x86)': dir('pf86') } });
    assert.equal(st.hancomInstalled, false);
    assert.equal(st.registered, false);
    assert.equal(st.moduleName, HANCOM_SECURITY_MODULE);
  });

  test('COM 등록 정보로 Hwp.exe를 찾고, 없으면 설치 폴더를 뒤진다', async () => {
    const h = fakeHancom();
    assert.equal(await findHwpExe({ reg: h.reg, env: h.env }), h.hwp);
    const pf = dir('pf86');
    const bin = join(pf, 'Hnc', 'Office 2022', 'HOffice120', 'Bin');
    mkdirSync(bin, { recursive: true });
    writeFileSync(join(bin, 'Hwp.exe'), fakePe(X86));
    assert.equal(await findHwpExe({ reg: fakeReg(), env: { 'ProgramFiles(x86)': pf } }), join(bin, 'Hwp.exe'));
  });

  test('등록: 따옴표 없는 전체 경로, 비트 수 확인, 해제', async () => {
    const h = fakeHancom(X86);
    const dll = join(dir('mod 폴더'), 'FilePathCheckerModuleExample.dll');
    writeFileSync(dll, fakePe(X86));
    const st = await setupHancomSecurity({ dllPath: `"${dll}"`, reg: h.reg, env: h.env });
    assert.deepEqual(h.reg.calls, [['set', MODULES_KEY, HANCOM_SECURITY_MODULE, dll]]);
    assert.equal(st.registered, true);
    assert.equal(st.dllExists, true);
    assert.equal(st.hwpArch, 'x86');
    assert.equal(st.archMatch, true);
    assert.equal(st.warning, undefined);

    assert.deepEqual(await removeHancomSecurity({ reg: h.reg }), { changed: true });
    assert.deepEqual(h.reg.calls.at(-1), ['remove', MODULES_KEY, HANCOM_SECURITY_MODULE]);
    assert.deepEqual(await removeHancomSecurity({ reg: h.reg }), { changed: false });
    assert.equal((await hancomSecurityStatus({ reg: h.reg, env: h.env })).registered, false);
  });

  test('등록 거절: 비트 수 다름 / 파일 없음 / DLL 아님', async () => {
    const h = fakeHancom(X86);
    const d = dir();
    writeFileSync(join(d, 'x64.dll'), fakePe(X64));
    await assert.rejects(setupHancomSecurity({ dllPath: join(d, 'x64.dll'), reg: h.reg, env: h.env }), /32비트인데 이 보안 모듈은 64비트/);
    await assert.rejects(setupHancomSecurity({ dllPath: join(d, 'none.dll'), reg: h.reg, env: h.env }), /보안 모듈 파일이 없습니다/);
    writeFileSync(join(d, 'text.dll'), 'hello');
    await assert.rejects(setupHancomSecurity({ dllPath: join(d, 'text.dll'), reg: h.reg, env: h.env }), /올바른 DLL 파일이 아닙니다/);
    await assert.rejects(setupHancomSecurity({ reg: h.reg }), /경로가 필요합니다/);
    assert.equal(h.reg.calls.length, 0);
  });

  test('상태 경고: 따옴표로 감싼 경로, 사라진 파일', async () => {
    const h = fakeHancom(X86);
    h.reg.values.set(`${MODULES_KEY}|${HANCOM_SECURITY_MODULE}`, '"C:\\x\\a.dll"');
    assert.match((await hancomSecurityStatus({ reg: h.reg, env: h.env })).warning, /따옴표/);
    h.reg.values.set(`${MODULES_KEY}|${HANCOM_SECURITY_MODULE}`, 'C:\\없는\\a.dll');
    const st = await hancomSecurityStatus({ reg: h.reg, env: h.env });
    assert.equal(st.dllExists, false);
    assert.match(st.warning, /파일이 없습니다/);
  });
});

// ── 공식 배포본 받기 (가짜 네트워크) ─────────────────────

async function moduleZip() {
  const zip = new JSZip();
  zip.file('보안모듈(Automation)/x64/FilePathCheckerModuleExample.dll', fakePe(X64, 'x64'));
  zip.file('보안모듈(Automation)/x86/FilePathCheckerModuleExample.dll', fakePe(X86, 'x86'));
  zip.file('보안모듈(Automation)/레지스트리.JPG', Buffer.from('jpg'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

function fakeFetch(routes) {
  const seen = [];
  const fetch = async (url, init) => {
    seen.push(url);
    assert.equal(init.redirect, 'manual');
    const r = routes[url];
    if (!r) return new Response('not found', { status: 404 });
    if (r.location) return new Response(null, { status: 302, headers: { location: r.location } });
    return new Response(r.body, { status: 200, headers: r.headers ?? {} });
  };
  return { fetch, seen };
}

describe('보안 모듈 받기', () => {
  const RAW = 'https://raw.githubusercontent.com/hancom-io/devcenter-archive/main/hwp-automation/x.zip';

  test('리디렉션을 따라가 zip에서 설치된 한글과 같은 비트의 DLL을 고른다', async () => {
    const zip = await moduleZip();
    const f = fakeFetch({ [HANCOM_SECURITY_MODULE_URL]: { location: RAW }, [RAW]: { body: zip } });
    const dest = dir('dest');
    const h = fakeHancom(X86);
    const dll = await downloadHancomSecurityModule(dest, { fetch: f.fetch, reg: h.reg, env: h.env });
    assert.equal(dll, join(dest, 'FilePathCheckerModuleExample.dll'));
    assert.ok(readFileSync(dll).equals(fakePe(X86, 'x86')));
    assert.deepEqual(f.seen, [HANCOM_SECURITY_MODULE_URL, RAW]);
    // 64비트를 원하면 64비트
    const dll64 = await downloadHancomSecurityModule(dir('dest64'), { fetch: f.fetch, arch: 'x64' });
    assert.equal(peMachine(dll64), 'x64');
  });

  test('DLL 하나만 오는 경우와 상대 경로 리디렉션', async () => {
    const f = fakeFetch({
      'https://github.com/a': { location: '/b' },
      'https://github.com/b': { body: fakePe(X86) },
    });
    const dll = await downloadHancomSecurityModule(dir(), { url: 'https://github.com/a', fetch: f.fetch, arch: 'x86' });
    assert.equal(peMachine(dll), 'x86');
  });

  test('HTTPS 아님 / 공식 아닌 주소 / 너무 큼 / 형식 이상 / 비트 없음 → 거절', async () => {
    const zip = await moduleZip();
    await assert.rejects(httpsDownload('http://github.com/x', { fetch: fakeFetch({}).fetch }), /HTTPS로만/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: fakeFetch({ 'https://github.com/x': { location: 'http://raw.githubusercontent.com/y' } }).fetch }), /HTTPS로만/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: fakeFetch({ 'https://github.com/x': { location: 'https://evil.example.com/y' } }).fetch }), /공식 배포처가 아닌 주소\(evil\.example\.com\)/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: fakeFetch({ 'https://github.com/x': { body: Buffer.alloc(2048) } }).fetch, maxBytes: 1024 }), /너무 큽니다/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: fakeFetch({ 'https://github.com/x': { body: 'x', headers: { 'content-length': '999999' } } }).fetch, maxBytes: 1024 }), /너무 큽니다/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: fakeFetch({}).fetch }), /HTTP 404/);
    const loop = fakeFetch({ 'https://github.com/x': { location: 'https://github.com/x' } });
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: loop.fetch }), /리디렉션이 너무 많습니다/);
    await assert.rejects(httpsDownload('https://github.com/x', { fetch: async () => { throw new TypeError('fetch failed'); } }), /인터넷 연결/);
    await assert.rejects(pickModuleDll(Buffer.from('<html>moved</html>'), 'x86'), /zip이나 DLL이 아닙니다/);
    const only64 = new JSZip();
    only64.file('a/x64.dll', fakePe(X64));
    await assert.rejects(pickModuleDll(await only64.generateAsync({ type: 'nodebuffer' }), 'x86'), /32비트용 보안 모듈이 없습니다 \(들어 있는 것: 64비트\)/);
    const empty = new JSZip();
    empty.file('readme.txt', 'x');
    await assert.rejects(pickModuleDll(await empty.generateAsync({ type: 'nodebuffer' }), 'x86'), /DLL이 없습니다/);
    assert.equal((await pickModuleDll(zip, 'x86')).arch, 'x86');
  });
});
