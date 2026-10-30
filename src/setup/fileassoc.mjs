// fileassoc.mjs — Office Kit를 Windows '기본 프로그램' 후보로 등록하고, 지금 어느 형식이 Office Kit로 열리는지 본다.
//   등록은 사용자별(HKCU)이고 기본 연결 프로그램은 바꾸지 않는다:
//     ProgID(OfficeKit.<확장자>) + .<확장자>\OpenWithProgids + Capabilities + RegisteredApplications
//   → 탐색기 '연결 프로그램'과 Windows 설정 > 앱 > 기본 앱에 Office Kit가 나오고, 고르는 것은 사용자가 한다.
//   (Windows 10부터 앱이 기본 프로그램을 직접 바꿀 수 없다 — UserChoice는 Windows만 쓴다)
//   설치 프로그램(build/installer.nsh)이 같은 키를 쓰고 지운다. 앱은 키가 없거나 다른 경로를 가리킬 때 고쳐 쓴다.
import { execFile } from 'node:child_process';

export const APP_NAME = 'Office Kit';
export const PROGID_PREFIX = 'OfficeKit.';
export const CAPABILITIES_KEY = 'Software\\OfficeKit\\Capabilities';

/** 등록하는 형식과 탐색기 '유형'에 보일 이름 (build/installer.nsh, src/core/index.mjs OPENABLE_EXTS와 같게) */
export const ASSOC_TYPES = {
  hwp: '한글 문서', hwpx: '한글 문서 (HWPX)', hwt: '한글 서식 파일',
  docx: 'Word 문서', doc: 'Word 97-2003 문서', rtf: '서식 있는 텍스트 문서', odt: 'ODF 텍스트 문서',
  pptx: 'PowerPoint 프레젠테이션', ppt: 'PowerPoint 97-2003 프레젠테이션', odp: 'ODF 프레젠테이션',
  xlsx: 'Excel 통합 문서', xls: 'Excel 97-2003 통합 문서', xlsm: 'Excel 매크로 사용 통합 문서', csv: 'CSV 파일', ods: 'ODF 스프레드시트',
  pdf: 'PDF 문서',
};

/** 설정 화면에 묶어 보일 순서 */
export const ASSOC_GROUPS = [
  { kind: '한글', exts: ['hwp', 'hwpx', 'hwt'] },
  { kind: 'Word', exts: ['docx', 'doc', 'rtf', 'odt'] },
  { kind: 'PowerPoint', exts: ['pptx', 'ppt', 'odp'] },
  { kind: 'Excel', exts: ['xlsx', 'xls', 'xlsm', 'csv', 'ods'] },
  { kind: 'PDF', exts: ['pdf'] },
];

const FILE_EXTS_KEY = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts';

/** Windows 설정 > 기본 앱 > Office Kit 쪽 주소 */
export const DEFAULT_APPS_URI = `ms-settings:defaultapps?registeredAppUser=${encodeURIComponent(APP_NAME)}`;

/** 탐색기가 실행할 명령 */
export const openCommand = (exe) => `"${exe}" "%1"`;

/**
 * 써야 할 레지스트리 값 목록 → [{ key, name, value }] (key는 HKCU 아래, name ''은 기본값).
 * root: 시험용으로 'Software' 대신 다른 키 아래에 쓴다
 */
export function registrationEntries(exe, { root = 'Software' } = {}) {
  const icon = `"${exe}",0`;
  const cls = `${root}\\Classes`;
  const caps = root === 'Software' ? CAPABILITIES_KEY : `${root}\\OfficeKit\\Capabilities`;
  const out = [];
  const put = (key, name, value) => out.push({ key, name, value });
  for (const [ext, desc] of Object.entries(ASSOC_TYPES)) {
    const id = `${PROGID_PREFIX}${ext}`;
    put(`${cls}\\${id}`, '', desc);
    put(`${cls}\\${id}`, 'FriendlyTypeName', desc);
    put(`${cls}\\${id}\\DefaultIcon`, '', icon);
    put(`${cls}\\${id}\\Application`, 'ApplicationName', APP_NAME);
    put(`${cls}\\${id}\\shell\\open\\command`, '', openCommand(exe));
    put(`${cls}\\.${ext}\\OpenWithProgids`, id, '');
    put(`${caps}\\FileAssociations`, `.${ext}`, id);
    put(`${cls}\\Applications\\${APP_NAME}.exe\\SupportedTypes`, `.${ext}`, '');
  }
  put(`${cls}\\Applications\\${APP_NAME}.exe`, 'FriendlyAppName', APP_NAME);
  put(`${cls}\\Applications\\${APP_NAME}.exe\\DefaultIcon`, '', icon);
  put(`${cls}\\Applications\\${APP_NAME}.exe\\shell\\open\\command`, '', openCommand(exe));
  put(caps, 'ApplicationName', APP_NAME);
  put(caps, 'ApplicationDescription', '한글·Word·PowerPoint·Excel·PDF 문서를 보고, 직접 고치고, AI 대화로 수정하는 프로그램');
  put(caps, 'ApplicationIcon', icon);
  put(`${root}\\RegisteredApplications`, APP_NAME, caps);
  return out;
}

// ── 레지스트리 ──────────────────────────────────────────

function regExe(args, { timeoutMs = 15_000 } = {}) {
  return new Promise((done) => {
    execFile('reg.exe', args, { windowsHide: true, timeout: timeoutMs, encoding: 'utf8' },
      (err, stdout, stderr) => done({ code: err ? (typeof err.code === 'number' ? err.code : 1) : 0, stdout: String(stdout ?? ''), stderr: String(stderr ?? '') }));
  });
}

// 읽기는 PowerShell로 (reg.exe 출력은 시스템 코드 페이지라 한글 사용자 폴더 경로가 깨진다). 스크립트는 ASCII만
const STATUS_PS = String.raw`
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$cu = [Microsoft.Win32.Registry]::CurrentUser
function V($key, $name) { $k = $cu.OpenSubKey($key); if ($k) { $v = $k.GetValue($name); $k.Close(); if ($v -ne $null) { return [string]$v } }; return '' }
$root = $env:OK_ROOT
$fe = 'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.'
$o = [ordered]@{ registered = (V "$root\RegisteredApplications" 'Office Kit'); types = [ordered]@{} }
foreach ($e in ($env:OK_EXTS -split ',')) {
  $o.types[$e] = [ordered]@{
    latest = (V "$fe$e\UserChoiceLatest" 'ProgId')
    user = (V "$fe$e\UserChoice" 'ProgId')
    command = (V "$root\Classes\OfficeKit.$e\shell\open\command" '')
    listed = (V "$($env:OK_CAPS)\FileAssociations" ".$e")
  }
}
$o | ConvertTo-Json -Depth 4 -Compress
`;

function readStatusRaw({ root, caps, timeoutMs = 20_000 }) {
  const enc = Buffer.from(STATUS_PS, 'utf16le').toString('base64');
  const env = { ...process.env, OK_ROOT: root, OK_CAPS: caps, OK_EXTS: Object.keys(ASSOC_TYPES).join(',') };
  return new Promise((done, fail) => {
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', enc],
      { env, windowsHide: true, timeout: timeoutMs, encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 },
      (err, stdout) => {
        if (err) { fail(new Error(`기본 프로그램 상태를 읽지 못했습니다: ${err.message}`)); return; }
        try { done(JSON.parse(String(stdout).trim().replace(/^\uFEFF/, ''))); } catch (e) { fail(new Error(`기본 프로그램 상태를 읽지 못했습니다: ${e.message}`)); }
      });
  });
}

const capsKey = (root) => (root === 'Software' ? CAPABILITIES_KEY : `${root}\\OfficeKit\\Capabilities`);

/**
 * 읽은 값 → 상태. raw = { registered, types: { ext: { latest, user, command, listed } } }
 * registered: Office Kit가 기본 앱 목록에 있음 / matches: 모든 형식이 등록돼 있고 열기 명령이 이 exe를 가리킴
 * isDefault: 그 형식을 지금 Office Kit로 연다 (사용자가 Windows에서 고른 것. Windows 11 24H2부터는 UserChoiceLatest가 앞선다)
 */
export function summarizeStatus(raw, exe) {
  const want = exe ? openCommand(exe).toLowerCase() : null;
  let all = true;
  const types = Object.entries(ASSOC_TYPES).map(([ext, desc]) => {
    const t = raw?.types?.[ext] ?? {};
    const progId = t.latest || t.user || null;
    if (!t.listed || !want || String(t.command ?? '').toLowerCase() !== want) all = false;
    return {
      ext, desc, kind: ASSOC_GROUPS.find((g) => g.exts.includes(ext))?.kind ?? '',
      progId, isDefault: !!progId && progId.toLowerCase() === `${PROGID_PREFIX}${ext}`.toLowerCase(),
    };
  });
  return { registered: !!raw?.registered, matches: all, exe: exe ?? null, types };
}

/** 지금 등록 상태와 형식마다 기본 프로그램 (summarizeStatus 참고) */
export async function assocStatus({ exe, root = 'Software', read = readStatusRaw } = {}) {
  return summarizeStatus(await read({ root, caps: capsKey(root) }), exe);
}

// ── Office Kit가 아닌 프로그램으로 열기 ────────────────────
// Office Kit가 그 형식의 기본 프로그램이면 '프로그램에서 열기'가 Office Kit 자신을 다시 연다 → 다른 프로그램을 찾는다.
// 1) Windows가 지금 쓰는 열기 명령(AssocQueryString)이 Office Kit가 아니면 그것, 2) 아니면 그 형식을 여는 다른 ProgID
//    (HKCR\.<확장자> 기본값, OpenWithProgids — OfficeKit.*·AppX* 제외)의 열기 명령. 스크립트는 ASCII만
const OTHER_PS = String.raw`
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public static class OkAssoc {
  [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
  static extern uint AssocQueryString(uint f, uint s, string a, string x, StringBuilder o, ref uint n);
  public static string Cmd(string a) { uint n = 2048; var sb = new StringBuilder(2048); return AssocQueryString(0, 1, a, null, sb, ref n) == 0 ? sb.ToString() : ""; }
}
"@
$ext = '.' + $env:OK_EXT
$cr = [Microsoft.Win32.Registry]::ClassesRoot
function CmdOf($progId) {
  $s = $cr.OpenSubKey("$progId\shell"); $verb = 'open'
  if ($s) { $d = [string]$s.GetValue(''); if ($d) { $verb = $d } }
  $k = $cr.OpenSubKey("$progId\shell\$verb\command"); if (-not $k) { $k = $cr.OpenSubKey("$progId\shell\open\command") }
  if ($k) { return [string]$k.GetValue('') } ; return ''
}
$eff = [OkAssoc]::Cmd($ext)
$list = New-Object System.Collections.ArrayList
$ids = New-Object System.Collections.ArrayList
$e = $cr.OpenSubKey($ext)
if ($e) { $d = [string]$e.GetValue(''); if ($d) { [void]$ids.Add($d) } }
$ow = $cr.OpenSubKey("$ext\OpenWithProgids")
if ($ow) { foreach ($n in $ow.GetValueNames()) { if (-not $ids.Contains($n)) { [void]$ids.Add($n) } } }
foreach ($id in $ids) {
  if ($id -like 'OfficeKit.*' -or $id -like 'AppX*') { continue }
  $c = CmdOf $id
  if ($c -and $c -notmatch 'Office Kit\.exe') { [void]$list.Add([ordered]@{ id = $id; command = $c }) }
}
$o = [ordered]@{ effective = $eff; candidates = @($list) }
$o | ConvertTo-Json -Compress
`;

// 다른 프로그램을 고를 때 먼저 볼 ProgID (설치된 오피스·한컴·PDF 뷰어). 목록에 없는 것은 레지스트리 순서대로
const PREFERRED_PROGIDS = [/^Word\./i, /^Excel\./i, /^PowerPoint\./i, /^Hwp\.Document/i, /^MSEdgePDF/i, /^AcroExch/i, /^Acrobat/i,
  /^ChromePDF|^ChromeHTML/i, /^HWord\./i, /^HCell\./i, /^HShow\./i];

/** otherOpenCommand가 읽은 것 { effective, candidates: [{ id, command }] } → 쓸 명령 | null */
export function pickOpenCommand(raw) {
  const eff = String(raw?.effective ?? '');
  if (eff && !/Office Kit\.exe/i.test(eff) && !/OpenWith\.exe/i.test(eff)) return eff;
  const list = [].concat(raw?.candidates ?? []).filter((c) => c && c.command && !/Office Kit\.exe/i.test(c.command));
  const rank = (id) => { const i = PREFERRED_PROGIDS.findIndex((re) => re.test(String(id))); return i < 0 ? PREFERRED_PROGIDS.length : i; };
  list.sort((a, b) => rank(a.id) - rank(b.id));
  return list[0]?.command ?? null;
}

/**
 * 그 형식을 Office Kit 말고 다른 프로그램으로 여는 명령 (예: "C:\...\Hwp.exe" "%1") → 문자열 | null
 */
export async function otherOpenCommand(ext, { timeoutMs = 20_000 } = {}) {
  const enc = Buffer.from(OTHER_PS, 'utf16le').toString('base64');
  return new Promise((done) => {
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', enc],
      { env: { ...process.env, OK_EXT: String(ext).replace(/^\./, '') }, windowsHide: true, timeout: timeoutMs, encoding: 'utf8' },
      (err, stdout) => {
        if (err) { done(null); return; }
        try { done(pickOpenCommand(JSON.parse(String(stdout).trim().replace(/^\uFEFF/, '')))); } catch { done(null); }
      });
  });
}

/**
 * 열기 명령 글 → { file, args } (args는 한 줄 그대로 — Windows 명령줄 규칙으로 넘긴다).
 * %1·%L·%l은 파일 경로로, %*·%2~%9 같은 나머지 자리는 지운다. %ProgramFiles% 같은 환경 변수는 풀어 준다
 */
export function parseOpenCommand(command, file, env = process.env) {
  let s = String(command ?? '').trim();
  if (!s) return null;
  s = s.replace(/%([A-Za-z_][A-Za-z0-9_()]*)%/g, (m, name) => {
    const k = Object.keys(env).find((x) => x.toLowerCase() === name.toLowerCase());
    return k ? env[k] : m;
  });
  let exe; let rest;
  if (s.startsWith('"')) {
    const end = s.indexOf('"', 1);
    if (end < 0) return null;
    exe = s.slice(1, end); rest = s.slice(end + 1);
  } else {
    // 따옴표 없는 경로에 빈칸이 있으면 (예: %ProgramFiles%\x\y.exe) .exe까지를 실행 파일로 본다 (Windows도 그렇게 찾는다)
    const m = /^(.*?\.(?:exe|com|bat|cmd))(?=\s|$)(.*)$/is.exec(s) ?? /^(\S+)(.*)$/s.exec(s);
    exe = m[1]; rest = m[2];
  }
  let used = false;
  rest = rest.replace(/%[1Ll]/g, () => { used = true; return file; }).replace(/%[*2-9]/g, '').trim();
  if (!used) rest = `${rest} "${file}"`.trim();
  return { file: exe, args: rest };
}

/** 등록하기(고쳐 쓰기). 기본 프로그램은 바꾸지 않는다 → 쓴 값 개수 */
export async function registerFileTypes({ exe, root = 'Software', run = regExe } = {}) {
  if (!exe) throw new Error('실행 파일 경로가 필요합니다');
  const entries = registrationEntries(exe, { root });
  for (const e of entries) {
    const args = ['add', `HKCU\\${e.key}`, ...(e.name ? ['/v', e.name] : ['/ve']), '/t', 'REG_SZ', '/d', e.value, '/f'];
    const r = await run(args);
    if (r.code !== 0) throw new Error(`기본 프로그램 후보로 등록하지 못했습니다: ${(r.stderr || r.stdout).trim() || `reg add 실패 (${r.code})`}`);
  }
  return entries.length;
}

/** 시험용: root 아래에 쓴 것 지우기 (실제 'Software'는 지우지 않는다 — 제거는 설치 프로그램이 한다) */
export async function unregisterForTest({ root, run = regExe } = {}) {
  if (!root || root === 'Software') throw new Error('시험용 키만 지울 수 있습니다');
  await run(['delete', `HKCU\\${root}`, '/f']);
}
