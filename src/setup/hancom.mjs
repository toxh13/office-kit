// hancom.mjs — 한글 자동화 보안 모듈(파일 접근 승인 모듈) 등록·해제·받기.
// 이 모듈이 없으면 한글을 COM으로 열 때마다 "접근 허용" 창이 떠서 자동 작업이 멈춘다.
// 레지스트리: HKCU\Software\HNC\HwpAutomation\Modules 의 값 이름 = HANCOM_SECURITY_MODULE, 값 = DLL 전체 경로(따옴표 없이)
// 한컴 COM은 여기서 절대 부르지 않는다 (레지스트리와 파일만 본다).
import { execFile } from 'node:child_process';
import { existsSync, writeFileSync, mkdirSync, openSync, readSync, closeSync, readdirSync, renameSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import JSZip from 'jszip';
import { HANCOM_SECURITY_MODULE } from '../core/com.mjs';

export { HANCOM_SECURITY_MODULE };
export const MODULES_KEY = 'HKCU\\Software\\HNC\\HwpAutomation\\Modules';

// 공식 배포처 (2026-10-17 확인)
//   한컴 개발자센터 "한글 오토메이션" 페이지 https://developer.hancom.com/hwpautomation 의 보안모듈 다운로드 링크가
//   한컴 공식 GitHub 저장소 hancom-io/devcenter-archive 의 hwp-automation/보안모듈(Automation).zip 을 가리킨다.
//   (https://github.com/hancom-io/devcenter-archive/tree/main/hwp-automation)
//   개발자센터 설명상 압축 안에 FilePathCheckerModuleExample.dll 이 있다 (폴더 구성은 몰라도 되게 PE 헤더로 고른다).
export const HANCOM_SECURITY_MODULE_PAGE = 'https://developer.hancom.com/hwpautomation';
// 브랜치(main)가 아니라 확인한 커밋으로 고정한다 (저장소가 바뀌어도 그 커밋의 파일만 받는다).
// 커밋 ec0c158c… = hwp-automation 폴더의 마지막 변경 (2025-04-21). 공식 배포처가 새 판을 내면 이 값을 확인해 바꾼다
export const HANCOM_SECURITY_MODULE_COMMIT = 'ec0c158c69195070dabcb086e11dd197c336bf8e';
export const HANCOM_SECURITY_MODULE_URL =
  `https://github.com/hancom-io/devcenter-archive/raw/${HANCOM_SECURITY_MODULE_COMMIT}/hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88(Automation).zip`;
// 압축을 풀 때 한도 (압축 폭탄 막기): 항목 하나 16MB, 모두 합쳐 64MB
const MAX_ENTRY_BYTES = 16 * 1024 * 1024;
const MAX_TOTAL_BYTES = 64 * 1024 * 1024;

// 받는 중 리디렉션은 GitHub 공식 주소로만 허용한다
const ALLOWED_HOSTS = new Set(['github.com', 'raw.githubusercontent.com', 'media.githubusercontent.com',
  'objects.githubusercontent.com', 'release-assets.githubusercontent.com', 'codeload.github.com']);
const MAX_BYTES = 30 * 1024 * 1024;
const DLL_NAME = 'FilePathCheckerModuleExample.dll';

// ── 레지스트리 (테스트에서는 가짜로 바꿔 끼운다) ─────────────

// 콘솔 출력은 보통 UTF-8이 아니므로(한국어 Windows는 CP949) 둘 다 시도
function decodeOem(buf) {
  if (!buf?.length) return '';
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch {}
  try { return new TextDecoder('euc-kr').decode(buf); } catch { return buf.toString('latin1'); }
}

// 읽기는 PowerShell(UTF-8 출력)로 해서 한글 경로도 깨지지 않게 한다. 키·이름은 환경 변수로 넘긴다.
const PS_GET = "$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.Encoding]::UTF8;"
  + "try{$k=Get-Item -LiteralPath ('Registry::'+$env:OFFICEKIT_REG_KEY);"
  + "$v=$k.GetValue($env:OFFICEKIT_REG_NAME,$null,'DoNotExpandEnvironmentNames');"
  + "if($null -eq $v){exit 2};[Console]::Out.Write([string]$v)}catch{exit 2}";

function exec(file, args, opts = {}) {
  return new Promise((done) => execFile(file, args, { windowsHide: true, timeout: 20_000, encoding: 'buffer', ...opts },
    (err, stdout, stderr) => done({ code: err ? (typeof err.code === 'number' ? err.code : -1) : 0, stdout, stderr })));
}

/** 실제 레지스트리: get(key, name) → 문자열|null, set(key, name, data), remove(key, name). name '' = 기본값 */
export const systemRegistry = {
  async get(key, name = '') {
    const r = await exec('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', PS_GET],
      { env: { ...process.env, OFFICEKIT_REG_KEY: key, OFFICEKIT_REG_NAME: name } });
    return r.code === 0 ? r.stdout.toString('utf8') : null;
  },
  async set(key, name, data) {
    const r = await exec('reg.exe', ['add', key, '/v', name, '/t', 'REG_SZ', '/d', data, '/f']);
    if (r.code !== 0) throw new Error(`레지스트리에 쓰지 못했습니다: ${decodeOem(r.stderr).trim() || `reg.exe 종료 코드 ${r.code}`}`);
  },
  async remove(key, name) {
    const r = await exec('reg.exe', ['delete', key, '/v', name, '/f']);
    if (r.code !== 0) throw new Error(`레지스트리에서 지우지 못했습니다: ${decodeOem(r.stderr).trim() || `reg.exe 종료 코드 ${r.code}`}`);
  },
};

// ── 한컴오피스 찾기 ──────────────────────────────────────

const MACHINE = { 0x14c: 'x86', 0x8664: 'x64', 0xaa64: 'arm64' };

/** PE(exe/dll) 머신 종류 → 'x86' | 'x64' | 'arm64' | null. 파일 경로나 Buffer */
export function peMachine(src) {
  let buf = src;
  if (typeof src === 'string') {
    buf = Buffer.alloc(4096);
    let fd;
    try { fd = openSync(src, 'r'); buf = buf.subarray(0, readSync(fd, buf, 0, 4096, 0)); } catch { return null; } finally { if (fd !== undefined) closeSync(fd); }
  }
  if (!buf || buf.length < 0x40 || buf.toString('latin1', 0, 2) !== 'MZ') return null;
  const pe = buf.readUInt32LE(0x3c);
  if (pe + 6 > buf.length || buf.toString('latin1', pe, pe + 4) !== 'PE\0\0') return null;
  return MACHINE[buf.readUInt16LE(pe + 4)] ?? null;
}

// "C:\...\Hwp.exe" -Automation → C:\...\Hwp.exe
function exeFromCommand(cmd) {
  const s = (cmd ?? '').trim();
  if (s.startsWith('"')) return s.slice(1, s.indexOf('"', 1));
  return s.match(/^(.*?\.exe)\b/i)?.[1] ?? (s.split(/\s+-/)[0] || null);
}

/** 설치된 한글(Hwp.exe) 위치. COM 등록 정보 → 흔한 설치 폴더 순서로 찾는다 */
export async function findHwpExe({ reg = systemRegistry, env = process.env } = {}) {
  const clsid = (await reg.get('HKCR\\HWPFrame.HwpObject\\CLSID', ''))?.trim();
  if (clsid) {
    for (const root of ['HKCR\\WOW6432Node\\CLSID', 'HKCR\\CLSID']) {
      const exe = exeFromCommand(await reg.get(`${root}\\${clsid}\\LocalServer32`, ''));
      if (exe && existsSync(exe)) return exe;
    }
  }
  for (const pf of [env['ProgramFiles(x86)'], env.ProgramFiles].filter(Boolean)) {
    const hnc = join(pf, 'Hnc');
    let offices = [];
    try { offices = readdirSync(hnc).sort().reverse(); } catch { continue; }
    for (const o of offices) {
      let subs = [];
      try { subs = readdirSync(join(hnc, o)).filter((d) => /^HOffice/i.test(d)); } catch { continue; }
      for (const s of subs) { const exe = join(hnc, o, s, 'Bin', 'Hwp.exe'); if (existsSync(exe)) return exe; }
    }
  }
  return null;
}

const BITS = { x86: '32비트', x64: '64비트', arm64: 'ARM64' };

/** → { hancomInstalled, hwpPath, hwpArch, registered, dllPath, dllExists, dllArch, archMatch, moduleName, warning? } */
export async function hancomSecurityStatus({ reg = systemRegistry, env = process.env } = {}) {
  const [hwpPath, value] = await Promise.all([findHwpExe({ reg, env }), reg.get(MODULES_KEY, HANCOM_SECURITY_MODULE)]);
  const raw = value?.trim() || null;
  const dllPath = raw ? raw.replace(/^"(.*)"$/, '$1') : null;
  const dllExists = !!dllPath && existsSync(dllPath);
  const hwpArch = hwpPath ? peMachine(hwpPath) : null;
  const dllArch = dllExists ? peMachine(dllPath) : null;
  const out = {
    hancomInstalled: !!hwpPath, hwpPath, hwpArch,
    registered: !!raw, dllPath, dllExists, dllArch,
    archMatch: !!(hwpArch && dllArch && hwpArch === dllArch),
    moduleName: HANCOM_SECURITY_MODULE,
  };
  if (raw && raw !== dllPath) out.warning = '레지스트리 경로가 따옴표로 감싸져 있어 한글이 모듈을 찾지 못할 수 있습니다. 다시 등록하세요.';
  else if (raw && !dllExists) out.warning = `등록된 보안 모듈 파일이 없습니다: ${dllPath}`;
  else if (hwpArch && dllArch && hwpArch !== dllArch) out.warning = `한컴오피스는 ${BITS[hwpArch]}인데 보안 모듈은 ${BITS[dllArch]}입니다.`;
  return out;
}

/** 보안 모듈 DLL을 레지스트리에 등록 → 상태 */
export async function setupHancomSecurity({ dllPath, reg = systemRegistry, env = process.env } = {}) {
  if (!dllPath) throw new Error('보안 모듈 DLL 경로가 필요합니다');
  dllPath = resolve(String(dllPath).trim().replace(/^"(.*)"$/, '$1'));
  if (!existsSync(dllPath)) throw new Error(`보안 모듈 파일이 없습니다: ${dllPath}`);
  const dllArch = peMachine(dllPath);
  if (!dllArch) throw new Error(`올바른 DLL 파일이 아닙니다: ${dllPath}`);
  const hwpPath = await findHwpExe({ reg, env });
  const hwpArch = hwpPath ? peMachine(hwpPath) : null;
  if (hwpArch && hwpArch !== dllArch) {
    throw new Error(`설치된 한컴오피스는 ${BITS[hwpArch]}인데 이 보안 모듈은 ${BITS[dllArch]}입니다. ${BITS[hwpArch]}용 DLL을 골라 주세요.`);
  }
  await reg.set(MODULES_KEY, HANCOM_SECURITY_MODULE, dllPath); // 따옴표 없이 전체 경로
  return hancomSecurityStatus({ reg, env });
}

/** 등록 해제 (DLL 파일은 지우지 않는다) */
export async function removeHancomSecurity({ reg = systemRegistry } = {}) {
  if (!(await reg.get(MODULES_KEY, HANCOM_SECURITY_MODULE))) return { changed: false };
  await reg.remove(MODULES_KEY, HANCOM_SECURITY_MODULE);
  return { changed: true };
}

// ── 공식 배포본 받기 ─────────────────────────────────────

/** HTTPS로만, 허용된 주소로만 리디렉션을 따라가며, 크기 제한 안에서 받는다 → Buffer */
export async function httpsDownload(url, { fetch: fetchImpl = globalThis.fetch, maxBytes = MAX_BYTES, timeoutMs = 60_000, maxRedirects = 5, allowedHosts = ALLOWED_HOSTS } = {}) {
  let current = new URL(url);
  for (let hop = 0; hop <= maxRedirects; hop++) {
    if (current.protocol !== 'https:') throw new Error(`보안 모듈은 HTTPS로만 받습니다: ${current.href}`);
    if (!allowedHosts.has(current.hostname)) throw new Error(`공식 배포처가 아닌 주소(${current.hostname})로 이동하려 해서 중단했습니다`);
    const ac = new AbortController();
    const timer = setTimeout(() => ac.abort(), timeoutMs);
    try {
      let res;
      try { res = await fetchImpl(current.href, { redirect: 'manual', signal: ac.signal, headers: { 'user-agent': 'OfficeKit' } }); } catch (e) {
        throw new Error(ac.signal.aborted ? `보안 모듈 받기 시간이 초과됐습니다 (${Math.round(timeoutMs / 1000)}초)` : `보안 모듈을 받지 못했습니다. 인터넷 연결을 확인하세요. (${e.cause?.code ?? e.message})`);
      }
      if ([301, 302, 303, 307, 308].includes(res.status)) {
        const loc = res.headers.get('location');
        if (!loc) throw new Error('보안 모듈 주소가 잘못된 곳으로 이동했습니다');
        current = new URL(loc, current);
        continue;
      }
      if (res.status !== 200) throw new Error(`보안 모듈을 받지 못했습니다 (HTTP ${res.status})`);
      if (Number(res.headers.get('content-length')) > maxBytes) throw new Error('보안 모듈 파일이 너무 큽니다');
      const chunks = [];
      let size = 0;
      try {
        for await (const chunk of res.body ?? []) {
          size += chunk.length;
          if (size > maxBytes) { ac.abort(); throw new Error('보안 모듈 파일이 너무 큽니다'); }
          chunks.push(Buffer.from(chunk));
        }
      } catch (e) {
        if (ac.signal.aborted && size <= maxBytes) throw new Error(`보안 모듈 받기 시간이 초과됐습니다 (${Math.round(timeoutMs / 1000)}초)`);
        throw e;
      }
      return Buffer.concat(chunks);
    } finally { clearTimeout(timer); }
  }
  throw new Error('보안 모듈 주소의 리디렉션이 너무 많습니다');
}

// 한국어 압축 파일은 이름이 CP949인 경우가 많다
function zipName(bytes) {
  const b = Buffer.from(bytes);
  try { return new TextDecoder('utf-8', { fatal: true }).decode(b); } catch { return new TextDecoder('euc-kr').decode(b); }
}

/** 받은 파일(zip 또는 dll)에서 원하는 비트 수의 DLL 고르기 → { name, data, arch } */
export async function pickModuleDll(bytes, arch, depth = 0) {
  if (bytes.toString('latin1', 0, 2) === 'MZ') {
    const a = peMachine(bytes);
    if (!a) throw new Error('받은 파일이 올바른 DLL이 아닙니다');
    if (arch && a !== arch) throw new Error(`받은 보안 모듈은 ${BITS[a]}용이라 ${BITS[arch]} 한컴오피스에 맞지 않습니다`);
    return { name: DLL_NAME, data: bytes, arch: a };
  }
  if (bytes.toString('latin1', 0, 4) !== 'PK\x03\x04') throw new Error('받은 파일이 zip이나 DLL이 아닙니다. 공식 배포처가 바뀌었을 수 있습니다.');
  const zip = await JSZip.loadAsync(bytes, { decodeFileName: zipName });
  const found = [];
  let total = 0;
  // 풀기 전에 적힌 크기로 거른다 (적힌 크기를 속여도 아래에서 실제 크기를 다시 본다)
  const sizeOf = (entry) => Number(entry?._data?.uncompressedSize ?? 0);
  const take = async (entry) => {
    if (sizeOf(entry) > MAX_ENTRY_BYTES || total + sizeOf(entry) > MAX_TOTAL_BYTES) throw new Error('받은 압축 파일의 내용이 너무 큽니다. 공식 배포처가 바뀌었을 수 있습니다.');
    const data = await entry.async('nodebuffer');
    total += data.length;
    if (data.length > MAX_ENTRY_BYTES || total > MAX_TOTAL_BYTES) throw new Error('받은 압축 파일의 내용이 너무 큽니다. 공식 배포처가 바뀌었을 수 있습니다.');
    return data;
  };
  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;
    if (/\.dll$/i.test(entry.name)) {
      const data = await take(entry);
      found.push({ name: entry.name, data, arch: peMachine(data) });
    } else if (/\.zip$/i.test(entry.name) && depth < 1) {
      try { found.push(await pickModuleDll(await take(entry), arch, depth + 1)); } catch (err) { if (/너무 큽니다/.test(err?.message)) throw err; }
    }
  }
  const dlls = found.filter((f) => f.arch);
  if (!dlls.length) throw new Error('받은 압축 파일 안에 보안 모듈 DLL이 없습니다. 공식 배포처가 바뀌었을 수 있습니다.');
  const fit = arch ? dlls.filter((f) => f.arch === arch) : dlls;
  if (!fit.length) {
    throw new Error(`받은 압축 파일에 ${BITS[arch]}용 보안 모듈이 없습니다 (들어 있는 것: ${[...new Set(dlls.map((d) => BITS[d.arch]))].join(', ')})`);
  }
  // 이름에 FilePathChecker가 들어간 것 → 경로가 짧은 것 순
  fit.sort((a, b) => (/FilePathChecker/i.test(b.name) - /FilePathChecker/i.test(a.name)) || a.name.length - b.name.length);
  return fit[0];
}

/**
 * 한컴 공식 배포본을 받아 destDir에 DLL을 저장 → DLL 경로.
 * 사용자가 버튼을 눌렀을 때만 부른다. 비트 수는 설치된 Hwp.exe에 맞춘다 (모르면 32비트).
 */
export async function downloadHancomSecurityModule(destDir, { url = HANCOM_SECURITY_MODULE_URL, fetch, arch, reg = systemRegistry, env = process.env, maxBytes, timeoutMs } = {}) {
  if (!destDir) throw new Error('보안 모듈을 저장할 폴더가 필요합니다');
  if (!arch) {
    const hwp = await findHwpExe({ reg, env });
    arch = (hwp && peMachine(hwp)) || 'x86';
  }
  const bytes = await httpsDownload(url, { fetch, maxBytes, timeoutMs });
  const dll = await pickModuleDll(bytes, arch);
  const base = basename(dll.name.replace(/\\/g, '/'));
  const name = /^[\w.-]+\.dll$/i.test(base) ? base : DLL_NAME;
  mkdirSync(destDir, { recursive: true });
  const dllPath = resolve(destDir, name);
  const tmp = `${dllPath}.${process.pid}.tmp`;
  writeFileSync(tmp, dll.data);
  renameSync(tmp, dllPath);
  return dllPath;
}

/** 받기 + 등록을 한 번에 (설정 화면의 "한글 자동화 연결" 버튼용) → 상태 */
export async function installHancomSecurity({ destDir, ...opts } = {}) {
  const dllPath = await downloadHancomSecurityModule(destDir, opts);
  return setupHancomSecurity({ dllPath, reg: opts.reg, env: opts.env });
}

