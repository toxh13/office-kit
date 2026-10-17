// com.mjs — 설치된 한컴/Word/PowerPoint/Excel을 PowerShell COM으로 조작한다.
// 네이티브 모듈 없이 동작하도록 요청·응답을 임시 JSON 파일로 주고받는다.
import { execFile } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync, rmSync, existsSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, resolve, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SCRIPT = resolve(here, '../../scripts/com/office.ps1');

// PATH에 기대지 않고 시스템 PowerShell을 직접 부른다 (OFFICEKIT_POWERSHELL: 시험용으로 바꿔 끼우기)
const SYSTEM_PS = join(process.env.SystemRoot ?? process.env.windir ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const POWERSHELL = process.env.OFFICEKIT_POWERSHELL || (existsSync(SYSTEM_PS) ? SYSTEM_PS : 'powershell.exe');

// 한글 자동화 보안 승인 모듈의 레지스트리 이름 (setup에서 등록)
export const HANCOM_SECURITY_MODULE = 'FilePathCheckerModule';

const APP_NAME = { hwp: '한글', word: 'Word', ppt: 'PowerPoint', excel: 'Excel', pdf: 'PDF 변환', probe: '프로그램 확인' };

// office.ps1이 공유 잠금(Local\OfficeKitCom)을 잡고 일하는 프로그램. 잠금을 잡으면 startedFile을 쓴다.
const LOCKED_APPS = new Set(['hwp', 'word', 'ppt', 'excel']);
// 다른 프로세스(앱 미리보기·AI의 MCP 서버)의 Office 작업이 끝나기를 기다리는 최대 시간
export const LOCK_WAIT_MS = 15 * 60_000;
const POLL_MS = 200;

// ── 조사 (이/가, 은/는, 을/를, 과/와, 으로/로) ─────────────────

const JOSA = { '이/가': ['이', '가'], '은/는': ['은', '는'], '을/를': ['을', '를'], '과/와': ['과', '와'], '으로/로': ['으로', '로'] };
// 영어 이름은 읽는 소리의 끝 글자로: Word(워드)·PowerPoint(파워포인트)·PDF(피디에프)는 받침 없음, Excel(엑셀)은 ㄹ
const SPOKEN_END = { word: '드', powerpoint: '트', excel: '셀', pdf: '프', powershell: '셸', office: '스', hwp: '피', hwpx: '스', kit: '트' };
// 받침 번호 (0 없음, 1 ㄱ, 4 ㄴ, 8 ㄹ, 16 ㅁ, 21 ㅇ) — 숫자·알파벳 한 글자를 읽을 때
const DIGIT_END = [21, 8, 0, 16, 0, 0, 1, 8, 8, 0]; // 영 일 이 삼 사 오 육 칠 팔 구
const LETTER_END = { L: 8, R: 8, M: 16, N: 4 };     // 엘 알 엠 엔 (나머지 알파벳은 받침 없음)

/** 끝소리 받침 번호 (0 = 받침 없음, 8 = ㄹ), 모르면 null */
function finalSound(word) {
  const s = String(word ?? '').trim().replace(/[\s.,!?'"’”)\]]+$/u, '');
  if (!s) return null;
  const last = s.at(-1);
  const code = last.charCodeAt(0);
  if (code >= 0xac00 && code <= 0xd7a3) return (code - 0xac00) % 28;
  const lastWord = s.split(/[\s\-_/]+/).at(-1);
  const spoken = SPOKEN_END[lastWord.toLowerCase()];
  if (spoken) return finalSound(spoken);
  if (/[0-9]/.test(last)) return DIGIT_END[Number(last)];
  if (/^[A-Z]+$/.test(lastWord)) return LETTER_END[last] ?? 0; // 약어는 글자 이름으로 (HTML → 엘)
  if (/[a-z]/i.test(last)) {
    if (/l$/i.test(lastWord)) return 8;
    if (/(?:m|n|ng)$/i.test(lastWord)) return 4;
    return 0;
  }
  return null;
}

/**
 * 낱말 + 알맞은 조사: josa('Word') → 'Word가', josa('Excel') → 'Excel이', josa('한글', '을/를') → '한글을'.
 * 끝소리를 모르면 '이(가)'처럼 둘 다 쓴다.
 */
export function josa(word, pair = '이/가') {
  const [withFinal, withoutFinal] = JOSA[pair] ?? String(pair).split('/');
  const f = finalSound(word);
  if (pair === '으로/로') return `${word}${f == null ? '(으)로' : f && f !== 8 ? '으로' : '로'}`;
  if (f == null) return `${word}${withFinal}(${withoutFinal})`;
  return `${word}${f ? withFinal : withoutFinal}`;
}

// ── 오류 메시지 ──────────────────────────────────────────

// 다른 프로그램이 문서를 쥐고 있을 때 Office·Windows가 내는 말 (영어 UI·한국어 UI)
const IN_USE = /in use by another|being used by another|used by another (?:process|user|application|program)|locked for editing|is locked by|file is (?:locked|in use|currently in use)|currently in use by|sharing violation|already open elsewhere|잠겨 있|잠갔|다른 사용자가 (?:사용|편집)|다른 프로세스(?:에서|가) (?:파일을 )?사용/i;
const PASSWORD = /password is incorrect|password you supplied|incorrect password|암호가 (?:틀|올바르지)|잘못된 암호/i;
// 작업별 오류(형식 모듈이 번호·작업 이름을 붙여 해석): 'opfail|n|op|…' (PowerPoint·Excel), 'op n (name): …' (Word)
const OP_ERROR = /\bopfail\||^op \d+ \([a-z_]*\): /;

// office.ps1은 영어 코드로 실패를 알린다 → 한국어로 바꾼다
const PS_ERRORS = [
  [PASSWORD, () => '암호가 걸린 문서는 열 수 없습니다'],
  // Word가 편집 제한(보호)된 문서에 내는 말: 'locked for editing'이 들어 있지만 다른 프로그램 때문이 아니다
  [/not available because the (?:document|object) is (?:locked|protected)/i, () => '편집이 제한된(보호된) 문서라 고칠 수 없습니다'],
  [IN_USE, () => '문서가 다른 프로그램에서 열려 있어 작업할 수 없습니다'],
  [/^write protected: (.*)$/s, (m) => `쓰기 암호가 걸린 문서라 그 자리에 저장할 수 없습니다: ${basename(m[1].trim())} — output(명령줄은 -o)으로 다른 파일에 저장하세요`],
  [/^open failed: (.*)$/s, (m) => `문서를 열지 못했습니다: ${m[1]}`],
  [/^save failed: (.*)$/s, (m) => `문서를 저장하지 못했습니다: ${m[1]}`],
  [/^unsupported format: (.*)$/s, (m) => `이 프로그램으로는 .${m[1]} 형식으로 저장할 수 없습니다`],
  [/^unknown (?:app|pdf op|op): (.*)$/s, (m) => `알 수 없는 요청입니다: ${m[1]}`],
  [/^sheet not found: (.*)$/s, (m) => `시트 없음: ${m[1]}`],
  [/^find too long$/, () => 'Word에서는 찾을 말과 바꿀 말이 255자까지만 됩니다'],
];

export function koreanComError(msg, app) {
  const s = String(msg ?? '').trim();
  const name = APP_NAME[app] ?? '프로그램';
  // 작업별 오류는 그대로 넘겨 형식 모듈(word·ppt·excel.mjs)이 몇 번째 작업인지와 함께 한국어로 바꾸게 한다
  if (OP_ERROR.test(s)) return `${name} 작업 실패: ${s}`;
  for (const [re, fmt] of PS_ERRORS) { const m = s.match(re); if (m) return fmt(m); }
  if (/[가-힣]/.test(s)) return s;
  return `${name} 작업 실패: ${s || '알 수 없는 오류'}`;
}

/** 120초 → '120초', 300초 → '5분', 150초 → '2분 30초' */
function duration(ms) {
  const sec = Math.round(ms / 1000);
  if (sec < 120) return `${sec}초`;
  const min = Math.floor(sec / 60);
  return sec % 60 ? `${min}분 ${sec % 60}초` : `${min}분`;
}

function timeoutMessage(app, ms) {
  if (app === 'hwp') return `한글이 응답하지 않습니다 (${duration(ms)} 초과). 보안 확인 창이 떠 있는지 확인하세요.`;
  const name = APP_NAME[app] ?? '프로그램';
  if (!LOCKED_APPS.has(app)) return `${josa(name)} 응답하지 않아 작업을 멈췄습니다 (${duration(ms)} 초과).`;
  return `${josa(name)} 응답하지 않아 작업을 멈췄습니다 (${duration(ms)} 초과). 문서가 크거나 프로그램에 확인 창이 떠 있을 수 있습니다.`;
}

function lockWaitMessage(ms) {
  return `다른 Office 작업이 끝나기를 기다리다 멈췄습니다 (${duration(ms)} 넘게 기다림). `
    + '다른 창이나 AI 대화에서 Office Kit가 Word·Excel·PowerPoint·한글로 일하는 중일 수 있습니다 — 그 작업이 끝난 뒤 다시 시도하세요.';
}

// ── 실행 ─────────────────────────────────────────────────

// 같은 프로그램을 동시에 두 번 띄우면 COM이 꼬이므로 한 번에 하나씩 실행한다.
// (다른 프로세스와는 office.ps1의 이름 있는 뮤텍스로 줄을 세운다)
let queue = Promise.resolve();

// 시간 초과로 PowerShell을 죽여도 DCOM이 띄운 Office는 남으므로 office.ps1이 적어 둔 PID로 정리한다 → 끝냈으면 true
function killOrphan(pidPath) {
  let pid = 0;
  try { pid = Number(readFileSync(pidPath, 'utf8').trim()); } catch { return false; }
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try { process.kill(pid); return true; } catch { return false; /* 이미 끝남 */ }
}

// 끝낸 Office가 남긴 '~$' 소유자 파일 이름 (Word는 긴 이름의 앞 1~2글자를 '~$'로 바꾼다)
function ownerFiles(app, file) {
  const name = basename(file);
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? dot : name.length;
  const own = app !== 'word' || stem <= 6 ? `~$${name}` : stem === 7 ? `~$${name.slice(1)}` : `~$${name.slice(2)}`;
  return join(dirname(file), own);
}
// 이번 호출이 시작된 뒤 생긴 작은 소유자 파일만 지운다 (사용자의 프로그램이 쓰는 중이면 지워지지 않음)
function removeOwnerFiles(request, since) {
  if (!['word', 'ppt', 'excel'].includes(request.app)) return;
  for (const f of new Set([request.path, request.save].filter(Boolean).map(String))) {
    const o = ownerFiles(request.app, f);
    try {
      const st = statSync(o);
      if (st.size < 4096 && st.mtimeMs >= since - 2000) rmSync(o);
    } catch { /* 없음·사용 중 */ }
  }
}

function runPs(request, timeoutMs, lockTimeoutMs) {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-'));
  const reqPath = join(dir, 'req.json');
  const resPath = join(dir, 'res.json');
  const pidPath = join(dir, 'pid.txt');
  const startedPath = join(dir, 'started.txt');
  writeFileSync(reqPath, JSON.stringify({ securityModule: HANCOM_SECURITY_MODULE, pidFile: pidPath, startedFile: startedPath, ...request }), 'utf8');
  const since = Date.now();
  return new Promise((done, fail) => {
    let stopped = null; // 'timeout' | 'lock'
    let timer = null;
    let poll = null;
    const child = execFile(POWERSHELL,
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, reqPath, resPath],
      { windowsHide: true },
      (err, _stdout, stderr) => {
        clearTimeout(timer); clearInterval(poll);
        try {
          let res = null;
          if (existsSync(resPath)) {
            // 끊기는 순간에 쓰던 응답은 깨져 있을 수 있다
            try { res = JSON.parse(readFileSync(resPath, 'utf8').replace(/^\uFEFF/, '')); } catch {
              if (!stopped) throw new Error(`${APP_NAME[request.app] ?? '프로그램'} 작업 결과를 읽지 못했습니다`);
            }
          }
          if (res) {
            if (res.ok) return done(res);
            return fail(new Error(koreanComError(res.error, request.app)));
          }
          if (stopped) {
            // 일하다 끊긴 경우만 Office가 떠 있다 (잠금을 기다리는 동안에는 아무것도 띄우지 않음)
            if (killOrphan(pidPath)) removeOwnerFiles(request, since);
            return fail(new Error(stopped === 'lock' ? lockWaitMessage(lockTimeoutMs) : timeoutMessage(request.app, timeoutMs)));
          }
          if (err?.code === 'ENOENT') return fail(new Error('PowerShell을 찾을 수 없어 설치된 프로그램을 쓸 수 없습니다'));
          const msg = stderr?.trim() || err?.message || '';
          if (/PSSecurityException|UnauthorizedAccess|cannot be loaded|실행할 수 없|디지털 서명/i.test(msg)) {
            return fail(new Error('PowerShell 실행 정책(그룹 정책) 때문에 Office Kit 스크립트를 실행할 수 없습니다. 프로그램 없이 되는 작업만 할 수 있습니다.'));
          }
          fail(new Error(koreanComError(msg || 'COM 호출 실패', request.app)));
        } catch (e) {
          fail(e);
        } finally {
          // 끝낸 PowerPoint가 사본(ppt-src)을 잠깐 더 쥐고 있을 수 있다: 지우기 실패로 프로세스가 죽지 않게
          try { rmSync(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }); } catch { /* 임시 폴더 */ }
        }
      });
    const stop = (why) => {
      if (stopped) return;
      stopped = why;
      clearInterval(poll);
      child.kill();
    };
    // 일하는 시간 제한은 office.ps1이 공유 잠금을 잡은 뒤(startedFile)부터 잰다. 그 전에는 잠금 대기 제한만.
    const startWork = () => {
      clearInterval(poll); clearTimeout(timer);
      timer = setTimeout(() => stop('timeout'), timeoutMs);
    };
    if (LOCKED_APPS.has(request.app)) {
      timer = setTimeout(() => stop('lock'), lockTimeoutMs);
      poll = setInterval(() => { if (!stopped && existsSync(startedPath)) startWork(); }, POLL_MS);
    } else {
      startWork();
    }
  });
}

/**
 * office.ps1 요청 하나를 실행한다 (프로세스 안에서는 차례로).
 * timeoutMs: 일하는 시간 제한 — 다른 프로세스의 Office 작업이 끝나기를 기다린 시간은 넣지 않는다.
 * lockTimeoutMs: 그 기다림의 상한 (기본 15분).
 */
export function comCall(request, { timeoutMs = 120_000, lockTimeoutMs = LOCK_WAIT_MS } = {}) {
  const job = queue.then(() => runPs(request, timeoutMs, lockTimeoutMs));
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
