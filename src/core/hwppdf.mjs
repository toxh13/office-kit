// hwppdf.mjs — 한컴오피스 없이 한글(hwp·hwpx·hwt) 문서를 PDF로.
//   rhwp 명령줄 프로그램(vendor/rhwp-cli/rhwp.exe, MIT, scripts/build-rhwp-cli.mjs가 고정 SHA-256으로 넣음)의 export-pdf(svg 방식)를 쓴다.
//   글자는 PDF 글꼴로 들어가(이 PC 글꼴의 쓴 글자만) 선택·검색이 되고, 쪽 크기·방향은 문서 그대로다.
//   입력은 앱 미리보기와 같은 화면용 사본(Hwp.openForView: 표 다음 문단 간격 등)을 hwp로 내보낸 임시 파일 → PDF가 미리보기와 같은 모양.
//   Node만으로 돈다 (Electron·Chromium 불필요) → 앱·MCP·명령줄이 같은 결과를 낸다. 원본 파일은 건드리지 않는다.
import { execFile } from 'node:child_process';
import { copyFileSync, existsSync, mkdtempSync, openSync, readSync, closeSync, renameSync, rmSync, statSync, unlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Hwp } from './hwp.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
export const RHWP_CLI_DEFAULT = join(ROOT, 'vendor', 'rhwp-cli', 'rhwp.exe');

/** 결과 안내 (앱·MCP·명령줄 공통) */
export const HWP_PDF_NOTE = '한컴오피스 없이 내장 엔진(rhwp)으로 만든 PDF입니다. 글꼴·줄바꿈·쪽 나눔이 한컴오피스와 조금 다를 수 있습니다 (앱 미리보기와 같은 모양).';

// PC에 없는 글꼴 대신 쓸 Windows 기본 글꼴 (rhwp 기본값 D2Coding 등은 보통 없다 → 고정폭 글자가 빈칸이 될 수 있음)
const FALLBACK_FONTS = ['--fallback-serif', 'Batang', '--fallback-sans', 'Malgun Gothic', '--fallback-mono', 'GulimChe'];

/** rhwp.exe 경로 (OFFICEKIT_RHWP_CLI로 바꿀 수 있음 — 시험·다른 OS용) */
export function rhwpCliPath() {
  return process.env.OFFICEKIT_RHWP_CLI ? resolve(process.env.OFFICEKIT_RHWP_CLI) : RHWP_CLI_DEFAULT;
}

/** 한컴오피스 없이 PDF를 만들 수 있는지 → { ok, path, reason? } */
export function hwpPdfStatus() {
  const path = rhwpCliPath();
  if (!process.env.OFFICEKIT_RHWP_CLI && process.platform !== 'win32') return { ok: false, path, reason: '내장 PDF 변환은 Windows에서만 됩니다' };
  if (!existsSync(path)) return { ok: false, path, reason: `내장 PDF 변환 프로그램이 없습니다: ${path} (개발 중이면 npm run rhwp-cli:build)` };
  return { ok: true, path };
}

// 쪽 수에 맞춘 시간 제한: 큰 문서(77쪽)도 1분 안팎이므로 넉넉히, 그래도 끝없이 기다리지는 않는다
const timeoutFor = (pages) => {
  const env = Number(process.env.OFFICEKIT_RHWP_TIMEOUT_MS);
  if (env > 0) return env;
  return Math.min(60_000 + pages * 5_000, 20 * 60_000);
};

const secs = (ms) => (ms < 10_000 ? Math.round(ms / 100) / 10 : Math.round(ms / 1000));

function run(exe, args, { cwd, timeoutMs }) {
  return new Promise((done) => {
    execFile(exe, args, { cwd, timeout: timeoutMs, killSignal: 'SIGKILL', windowsHide: true, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024, shell: false },
      (err, stdout, stderr) => done({ err, stdout: String(stdout ?? ''), stderr: String(stderr ?? '') }));
  });
}

function isPdf(file) {
  let fd;
  try {
    fd = openSync(file, 'r');
    const b = Buffer.alloc(5);
    return readSync(fd, b, 0, 5, 0) === 5 && b.toString('latin1') === '%PDF-';
  } catch { return false; } finally { if (fd !== undefined) closeSync(fd); }
}

// 쓰는 도중 끊겨도 이미 있던 output이 반쯤 망가지지 않게: 같은 폴더 임시 파일에 다 쓴 뒤 이름을 바꿔 끼운다
function place(from, output) {
  const part = `${output}.${process.pid}.${Date.now()}.tmp`;
  try {
    copyFileSync(from, part);
    renameSync(part, output);
  } catch (e) {
    try { unlinkSync(part); } catch {}
    const busy = (err) => ['EPERM', 'EACCES', 'EBUSY'].includes(err?.code);
    if (!busy(e) && e.code !== 'EXDEV') throw e;
    try { copyFileSync(from, output); } catch (e2) {
      if (busy(e2)) throw new Error(`PDF를 저장하지 못했습니다 — 같은 이름의 파일이 다른 프로그램(PDF 뷰어 등)에서 열려 있거나 쓰기 권한이 없습니다: ${output}`);
      throw e2;
    }
  }
}

/**
 * 한글 문서 → PDF (한컴오피스 없이). → { engine: 'rhwp', output, pages, ms }
 * opts.timeoutMs: 시간 제한 (기본: 쪽 수에 맞춤). 원본은 바꾸지 않는다.
 */
export async function hwpToPdf(path, output, { timeoutMs } = {}) {
  path = resolve(path);
  output = resolve(output);
  const st = hwpPdfStatus();
  if (!st.ok) throw new Error(`한글 문서를 PDF로 바꿀 수 없습니다: ${st.reason}`);
  const started = Date.now();
  // 미리보기와 같은 화면용 사본 → hwp 바이트 (파일을 못 읽으면 여기서 한국어 오류)
  let bytes; let pages;
  let h;
  try { h = Hwp.openForView(path); } catch (e) {
    throw new Error(`한글 문서를 열지 못해 PDF로 바꿀 수 없습니다: ${e?.message ?? e}`);
  }
  try { pages = h.pageCount; bytes = h.toBytes('hwp'); } finally { h.close(); }
  const work = mkdtempSync(join(tmpdir(), 'officekit-hwppdf-'));
  try {
    const src = join(work, 'doc.hwp');
    const pdf = join(work, 'doc.pdf');
    writeFileSync(src, bytes);
    const limit = timeoutMs > 0 ? timeoutMs : timeoutFor(pages);
    // 셸 없이 인수 배열로 (경로에 공백·특수 문자가 있어도 안전), 작업 폴더는 임시 폴더
    const r = await run(st.path, ['export-pdf', src, '-o', pdf, '--backend', 'svg', '--json', ...FALLBACK_FONTS], { cwd: work, timeoutMs: limit });
    if (r.err) {
      if (r.err.killed || r.err.signal) throw new Error(`한글 문서를 PDF로 바꾸는 데 시간이 너무 오래 걸려 멈췄습니다 (${secs(limit)}초, ${pages}쪽)`);
      if (r.err.code === 'ENOENT' || r.err.code === 'EACCES') throw new Error(`내장 PDF 변환 프로그램을 실행하지 못했습니다: ${st.path} (${r.err.code})`);
      const why = r.stderr.split(/\r?\n/).map((l) => l.trim()).filter(Boolean).filter((l) => !/^WARN|^\s*(한글 또는|--fallback)/.test(l)).pop() ?? r.err.message;
      throw new Error(`한글 문서를 PDF로 바꾸지 못했습니다 (내장 엔진 rhwp): ${why.replace(/^오류:\s*/, '')}`);
    }
    if (!existsSync(pdf) || !isPdf(pdf)) throw new Error('한글 문서를 PDF로 바꾸지 못했습니다 (내장 엔진 rhwp): 결과 PDF가 만들어지지 않았습니다');
    let manifest = {};
    try { manifest = JSON.parse(r.stdout.trim().split(/\r?\n/).pop()); } catch {}
    place(pdf, output);
    return { engine: 'rhwp', output, pages: Number(manifest.pageCount) || pages, bytes: statSync(output).size, ms: Date.now() - started };
  } finally {
    // 끊은 프로세스가 파일을 잠깐 쥐고 있을 수 있다 → 몇 번 다시, 그래도 안 되면 남겨 둔다 (원래 오류를 가리지 않게)
    try { rmSync(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }); } catch {}
  }
}
