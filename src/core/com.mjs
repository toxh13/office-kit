// com.mjs — 설치된 한컴/Word/PowerPoint/Excel을 PowerShell COM으로 조작한다.
// 네이티브 모듈 없이 동작하도록 요청·응답을 임시 JSON 파일로 주고받는다.
import { execFile } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync, rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SCRIPT = resolve(here, '../../scripts/com/office.ps1');

// PATH에 기대지 않고 시스템 PowerShell을 직접 부른다 (OFFICEKIT_POWERSHELL: 시험용으로 바꿔 끼우기)
const SYSTEM_PS = join(process.env.SystemRoot ?? process.env.windir ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const POWERSHELL = process.env.OFFICEKIT_POWERSHELL || (existsSync(SYSTEM_PS) ? SYSTEM_PS : 'powershell.exe');

// 한글 자동화 보안 승인 모듈의 레지스트리 이름 (setup에서 등록)
export const HANCOM_SECURITY_MODULE = 'FilePathCheckerModule';

const APP_NAME = { hwp: '한글', word: 'Word', ppt: 'PowerPoint', excel: 'Excel', pdf: 'PDF 변환', probe: '프로그램 확인' };

// office.ps1은 영어 코드로 실패를 알린다 → 한국어로 바꾼다
const PS_ERRORS = [
  [/password is incorrect|password you supplied|incorrect password|암호가 (?:틀|올바르지)|잘못된 암호/i, () => '암호가 걸린 문서는 열 수 없습니다'],
  [/being used by another|used by another process|locked for editing|잠겨 있|다른 사용자가 사용/i, () => '문서가 다른 프로그램에서 열려 있어 작업할 수 없습니다'],
  [/^open failed: (.*)$/s, (m) => `문서를 열지 못했습니다: ${m[1]}`],
  [/^save failed: (.*)$/s, (m) => `문서를 저장하지 못했습니다: ${m[1]}`],
  [/^unsupported format: (.*)$/s, (m) => `이 프로그램으로는 .${m[1]} 형식으로 저장할 수 없습니다`],
  [/^unknown (?:app|pdf op|op): (.*)$/s, (m) => `알 수 없는 요청입니다: ${m[1]}`],
  [/^sheet not found: (.*)$/s, (m) => `시트 없음: ${m[1]}`],
  [/^find too long$/, () => 'Word에서는 찾을 말과 바꿀 말이 255자까지만 됩니다'],
];

export function koreanComError(msg, app) {
  const s = String(msg ?? '').trim();
  for (const [re, fmt] of PS_ERRORS) { const m = s.match(re); if (m) return fmt(m); }
  if (/[가-힣]/.test(s)) return s;
  return `${APP_NAME[app] ?? '프로그램'} 작업 실패: ${s || '알 수 없는 오류'}`;
}

// 같은 프로그램을 동시에 두 번 띄우면 COM이 꼬이므로 한 번에 하나씩 실행한다.
// (다른 프로세스와는 office.ps1의 이름 있는 뮤텍스로 줄을 세운다)
let queue = Promise.resolve();

// 시간 초과로 PowerShell을 죽여도 DCOM이 띄운 Office는 남으므로 office.ps1이 적어 둔 PID로 정리한다
function killOrphan(pidPath) {
  let pid = 0;
  try { pid = Number(readFileSync(pidPath, 'utf8').trim()); } catch { return; }
  if (!Number.isInteger(pid) || pid <= 0) return;
  try { process.kill(pid); } catch { /* 이미 끝남 */ }
}

function runPs(request, timeoutMs) {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-'));
  const reqPath = join(dir, 'req.json');
  const resPath = join(dir, 'res.json');
  const pidPath = join(dir, 'pid.txt');
  writeFileSync(reqPath, JSON.stringify({ securityModule: HANCOM_SECURITY_MODULE, pidFile: pidPath, ...request }), 'utf8');
  return new Promise((done, fail) => {
    execFile(POWERSHELL,
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, reqPath, resPath],
      { timeout: timeoutMs, windowsHide: true },
      (err, _stdout, stderr) => {
        try {
          if (existsSync(resPath)) {
            const res = JSON.parse(readFileSync(resPath, 'utf8').replace(/^\uFEFF/, ''));
            if (res.ok) return done(res);
            return fail(new Error(koreanComError(res.error, request.app)));
          }
          if (err?.killed) {
            killOrphan(pidPath);
            const sec = Math.round(timeoutMs / 1000);
            return fail(new Error(request.app === 'hwp'
              ? `한글이 응답하지 않습니다 (${sec}초 초과). 보안 확인 창이 떠 있는지 확인하세요.`
              : `${APP_NAME[request.app] ?? '프로그램'}이(가) 응답하지 않아 작업을 멈췄습니다 (${sec}초 초과). 문서가 크거나 프로그램에 확인 창이 떠 있을 수 있습니다.`));
          }
          if (err?.code === 'ENOENT') return fail(new Error('PowerShell을 찾을 수 없어 설치된 프로그램을 쓸 수 없습니다'));
          const msg = stderr?.trim() || err?.message || '';
          if (/PSSecurityException|UnauthorizedAccess|cannot be loaded|실행할 수 없|디지털 서명/i.test(msg)) {
            return fail(new Error('PowerShell 실행 정책(그룹 정책) 때문에 Office Kit 스크립트를 실행할 수 없습니다. 프로그램 없이 되는 작업만 할 수 있습니다.'));
          }
          fail(new Error(koreanComError(msg || 'COM 호출 실패', request.app)));
        } finally {
          rmSync(dir, { recursive: true, force: true });
        }
      });
  });
}

export function comCall(request, { timeoutMs = 120_000 } = {}) {
  const job = queue.then(() => runPs(request, timeoutMs));
  queue = job.catch(() => {});
  return job;
}

const NONE = { hancom: false, word: false, powerpoint: false, excel: false, hancomSecurityModule: false };
let probeCache;
let probeFailedAt = 0;
/**
 * 설치된 프로그램 확인: { hancom, word, powerpoint, excel, hancomSecurityModule }
 * PowerShell을 못 쓰면 모두 false (+ error)로 돌려 파일 엔진 작업은 계속되게 한다.
 */
export function probe({ refresh = false } = {}) {
  if (process.platform !== 'win32') return Promise.resolve({ ...NONE });
  // 실패 결과는 30초만 기억하고 다시 확인한다
  if (probeFailedAt && Date.now() - probeFailedAt > 30_000) refresh = true;
  if (!probeCache || refresh) {
    probeFailedAt = 0;
    probeCache = comCall({ app: 'probe' }, { timeoutMs: 20_000 }).catch((e) => {
      console.error('[office-kit] 설치된 프로그램을 확인하지 못했습니다:', e.message);
      probeFailedAt = Date.now();
      return { ...NONE, error: e.message };
    });
  }
  return probeCache;
}

// 확장자 → COM 대상 프로그램
const APP_BY_EXT = {
  hwp: 'hwp', hwpx: 'hwp', hwt: 'hwp',
  docx: 'word', doc: 'word', rtf: 'word', odt: 'word',
  pptx: 'ppt', ppt: 'ppt', odp: 'ppt',
  xlsx: 'excel', xls: 'excel', xlsm: 'excel', csv: 'excel', ods: 'excel',
};
const PROBE_KEY = { hwp: 'hancom', word: 'word', ppt: 'powerpoint', excel: 'excel' };

export function comAppFor(ext) { return APP_BY_EXT[ext]; }

/**
 * 이 형식을 프로그램 엔진으로 다룰 수 있는지.
 * 한글은 보안 승인 모듈이 없으면 파일마다 "접근 허용" 창이 떠서 자동 작업이 멈추므로,
 * 사람이 보는 작업(interactive)이 아니면 모듈이 등록돼 있을 때만 쓴다.
 */
export async function comAvailableFor(ext, { interactive = false } = {}) {
  const app = APP_BY_EXT[ext];
  if (!app) return false;
  const p = await probe();
  if (!p[PROBE_KEY[app]]) return false;
  if (app === 'hwp' && !interactive && !p.hancomSecurityModule) return false;
  return true;
}
