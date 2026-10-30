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
