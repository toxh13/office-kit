// SPDX-License-Identifier: AGPL-3.0-only
// index.mjs — Office 편집기 글꼴: 설치된 글꼴로 AllFonts.js·미리보기 그림을 만들어 캐시에 두고, 글꼴 파일을 내주는 도우미.
//   Office Kit 글꼴 생성기의 일부 (src/app/office-editor/fonts/*는 AGPL-3.0-only — ONLYOFFICE AllFontsGen 규칙을 옮김).
//
//   const fonts = await ensureFonts({ userData: app.getPath('userData'), editorVersion });
//   server.setOfficeFonts(fonts);   // { allFontsJsPath, thumbnailsDir, fontFileByIndex(번호) → 글꼴 파일 절대 경로 }
//   // 또는 /office/<판>/editor/<rel> 요청마다: fonts.resolve(rel) → null | { type:'file', file, contentType } | { type:'font', file, size, contentType, open() }
//
// 캐시: <userData>/office-editor/fonts/<편집기 버전>-<폴더 목록 해시>/ (목록 = 경로·크기·수정 시각, 바뀌면 다시 만든다)
import { createHash } from 'node:crypto';
import { createReadStream, existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { Transform } from 'node:stream';
import { Worker } from 'node:worker_threads';
import { FONT_XOR_KEY, GENERATOR_VERSION, listFontFiles } from './allfonts.mjs';
import { CACHE_FILES, runAllFontsJob, runThumbsJob } from './jobs.mjs';
import { isThumbName } from './thumbs.mjs';

export { FONT_XOR_KEY, GENERATOR_VERSION, generateAllFonts, listFontFiles, PROFILES, KOREAN_ALIASES } from './allfonts.mjs';
export { THUMB_FILES, decodeSpriteBin, encodeSpriteBin, spriteBinToPng, thumbFileName } from './thumbs.mjs';

/** 읽는 글꼴 폴더: C:\Windows\Fonts, %LOCALAPPDATA%\Microsoft\Windows\Fonts (사용자별 설치) */
export function defaultFontDirs(env = process.env) {
  const dirs = [join(env.WINDIR || env.SystemRoot || 'C:\\Windows', 'Fonts')];
  if (env.LOCALAPPDATA) dirs.push(join(env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'));
  return dirs;
}

/** 캐시 뿌리 폴더 */
export const fontsCacheRoot = (userDataDir) => join(userDataDir, 'office-editor', 'fonts');

const safeVersion = (v) => (String(v || 'unknown').replace(/[^0-9A-Za-z._+-]/g, '_').slice(0, 48) || 'unknown');

/** 캐시 키 = <편집기 버전>-<해시 16자> (해시: 생성기 판·설정·폴더·파일 목록(경로·크기·수정 시각)) */
export function fontsCacheKey({ editorVersion, listing, fontDirs = [], profile = 'officekit' }) {
  const h = createHash('sha256');
  h.update(JSON.stringify({ g: GENERATOR_VERSION, profile, dirs: fontDirs.map((d) => resolve(d).toLowerCase()) }));
  for (const f of listing) h.update(`\n${f.path.toLowerCase()}|${f.size}|${f.mtimeMs}`);
  return `${safeVersion(editorVersion)}-${h.digest('hex').slice(0, 16)}`;
}

/** 버퍼의 앞 32바이트를 XOR (offset = 이 버퍼가 파일에서 시작하는 위치) → 새 버퍼 */
export function obfuscateFontBytes(buf, offset = 0) {
  const out = Buffer.from(buf);
  for (let i = 0; i < out.length && offset + i < 32; i++) out[i] ^= FONT_XOR_KEY[(offset + i) % 16];
  return out;
}

/**
 * 글꼴 파일을 읽기 전용으로 열어 앞 32바이트만 XOR해서 흘려보낸다 (ONLYOFFICE 웹 글꼴 모양, 복사하지 않음)
 * → Readable. 파일 오류는 스트림 'error'로
 */
export function createObfuscatedStream(path) {
  let pos = 0;
  const t = new Transform({
    transform(chunk, _enc, cb) {
      const out = pos < 32 ? obfuscateFontBytes(chunk, pos) : chunk;
      pos += chunk.length;
      cb(null, out);
    },
  });
  const rs = createReadStream(path, { flags: 'r' });
  rs.on('error', (err) => t.destroy(err));
  t.on('close', () => rs.destroy());
  return rs.pipe(t);
}

// ── worker ──
const running = new Set();
function runWorker(workerData, { onMessage = () => {}, unref = false } = {}) {
  const w = new Worker(new URL('./worker.mjs', import.meta.url), { workerData });
  running.add(w);
  if (unref) w.unref();
  const p = new Promise((ok, fail) => {
    let settled = false;
    w.on('message', (m) => {
      if (m?.type === 'done') { settled = true; ok(m.result); }
      else if (m?.type === 'error') { settled = true; fail(Object.assign(new Error(m.message), { workerStack: m.stack })); }
      else onMessage(m);
    });
    w.on('error', (err) => { if (!settled) { settled = true; fail(err); } });
    w.on('exit', (code) => { running.delete(w); if (!settled) { settled = true; fail(new Error(`글꼴 작업이 멈췄습니다 (종료 코드 ${code})`)); } });
  });
  return { worker: w, promise: p };
}

/** 돌고 있는 글꼴 작업(미리보기 그리기 등)을 멈춘다 — 앱을 끝낼 때 */
export async function stopFontsWork() {
  await Promise.all([...running].map(async (w) => { try { w.postMessage({ type: 'stop' }); } catch {} await w.terminate().catch(() => {}); }));
  running.clear();
}

const MIME_JS = 'text/javascript; charset=utf-8';
const MIME_BIN = 'application/octet-stream';

/** 완성된 캐시 폴더 → 손잡이 (W3 editor-server의 setOfficeFonts가 allFontsJsPath·thumbnailsDir·fontFileByIndex를 읽는다) */
function makeHandle(dir, manifest, { thumbnails }) {
  const byNum = new Map(manifest.files.map((f) => [Number(f.id), f]));
  const stageDone = (s) => existsSync(join(dir, CACHE_FILES.stageDone(s)));
  const allFontsPath = join(dir, CACHE_FILES.allFonts);
  const imagesDir = join(dir, CACHE_FILES.images);
  /** 번호 → { id, path, size } | null (그 뒤에 파일이 바뀌었으면(크기 다름·없음) null) */
  const info = (n) => {
    if (typeof n === 'string' && !/^\d{1,6}$/.test(n)) return null;
    if (typeof n !== 'string' && !Number.isInteger(n)) return null;
    const f = byNum.get(Number(n));
    if (!f) return null;
    try { const st = statSync(f.path); if (!st.isFile() || st.size !== f.size) return null; } catch { return null; }
    return { id: f.id, path: f.path, size: f.size };
  };
  const handle = {
    key: manifest.key, dir, editorVersion: manifest.editorVersion,
    /** 캐시의 AllFonts.js (절대 경로) — /office/<판>/editor/sdkjs/common/AllFonts.js 대신 */
    allFontsJsPath: allFontsPath,
    allFontsPath,
    /** 미리보기 그림 폴더 (절대 경로) — /office/<판>/editor/sdkjs/common/Images/fonts_thumbnail*.png(.bin) 대신.
     *  그림 10개가 늘 있고, 빈 그림 → 기본 글꼴 → 글꼴마다 그린 것으로 제자리에서 바뀐다 */
    thumbnailsDir: imagesDir,
    stats: manifest.stats,
    fileCount: manifest.files.length,
    /** 미리보기 단계: 'full' | 'plain' | 'blank' */
    thumbnailsStage() { return stageDone('full') ? 'full' : stageDone('plain') ? 'plain' : 'blank'; },
    /** 미리보기 파일 이름 (fonts_thumbnail…png.bin) → 경로 | null */
    thumbnailPath(name) {
      if (!isThumbName(name)) return null;
      const p = join(imagesDir, name);
      return existsSync(p) ? p : null;
    },
    /**
     * 웹 글꼴 번호(주소 /office/<판>/editor/fonts/<번호>, 12 또는 '012') → 실제 글꼴 파일 절대 경로 | null.
     * 내주는 쪽은 그 파일을 읽기 전용으로 열고 앞 32바이트만 XOR한다 (createObfuscatedStream)
     */
    fontFileByIndex(n) { return info(n)?.path ?? null; },
    /** 번호 → { id, path, size } | null */
    fontFileInfo: info,
    /** 번호 → 앞 32바이트를 XOR한 스트림 | null */
    createFontStream(n) { const f = info(n); return f ? createObfuscatedStream(f.path) : null; },
    /**
     * 편집기 폴더 기준 경로(rel, 예: 'sdkjs/common/AllFonts.js', 'sdkjs/common/Images/fonts_thumbnail@2x.png.bin', 'fonts/012')
     * → null(글꼴 경로가 아님·없음) | { type:'file', file, contentType } | { type:'font', id, file, size, contentType, open() }
     */
    resolve(rel) {
      if (typeof rel !== 'string') return null;
      const r = rel.replace(/^\/+/, '');
      if (r === 'sdkjs/common/AllFonts.js') return { type: 'file', file: allFontsPath, contentType: MIME_JS };
      let m = /^sdkjs\/common\/Images\/([^/]+)$/.exec(r);
      if (m && isThumbName(m[1])) {
        const file = handle.thumbnailPath(m[1]);
        return file ? { type: 'file', file, contentType: m[1].endsWith('.bin') ? MIME_BIN : 'image/png' } : null;
      }
      m = /^fonts\/(\d{1,6})$/.exec(r);
      if (m) {
        const f = info(m[1]);
        return f ? { type: 'font', id: f.id, file: f.path, size: f.size, contentType: MIME_BIN, open: () => createObfuscatedStream(f.path) } : null;
      }
      return null;
    },
    /** 미리보기 그리기가 끝나면 → 마지막 단계 ('full' | 'plain' | 'blank') */
    thumbnails,
  };
  return handle;
}

function readManifest(dir) {
  try {
    const m = JSON.parse(readFileSync(join(dir, CACHE_FILES.manifest), 'utf8'));
    if (m?.v !== 1 || m.generator !== GENERATOR_VERSION || !Array.isArray(m.files)) return null;
    if (!existsSync(join(dir, CACHE_FILES.allFonts)) || !existsSync(join(dir, CACHE_FILES.thumbs))) return null;
    return m;
  } catch { return null; }
}

/** 다른 키의 옛 캐시 지우기 (이 프로세스가 쓰고 있는 캐시는 두고, 만드는 중인 임시 폴더는 1시간 지난 것만) */
function cleanOld(root, keep) {
  let ents = [];
  try { ents = readdirSync(root, { withFileTypes: true }); } catch { return; }
  for (const e of ents) {
    const p = join(root, e.name);
    if (!e.isDirectory() || e.name === keep || liveDirs.has(p)) continue;
    if (/\.tmp-/.test(e.name)) { try { if (Date.now() - statSync(p).mtimeMs < 3600e3) continue; } catch { continue; } }
    try { rmSync(p, { recursive: true, force: true }); } catch {}
  }
}

const inflight = new Map(); // 캐시 폴더 → Promise<handle>
const liveDirs = new Set(); // 이 프로세스가 내주고 있는 캐시 폴더 (지우지 않는다)
const thumbJobs = new Map(); // 캐시 폴더 → Promise

/**
 * 글꼴 캐시를 준비한다 (없거나 글꼴 폴더가 바뀌었으면 worker_thread에서 새로 만든다) → Promise<손잡이>
 * opts: {
 *   userData,                     // 필수 (app.getPath('userData'), userDataDir도 됨) — 또는 cacheRoot로 캐시 뿌리를 바로
 *   editorVersion,                // 편집기 묶음 버전 (vendor/office-editor/VERSION.json) — 캐시 키에 들어감
 *   fontDirs = defaultFontDirs(), profile = 'officekit',
 *   thumbnails = true,            // 미리보기 그림을 뒤에서 그린다 (plain 약 1초 → full 약 6초)
 *   waitPlainMs = 2500,           // 새로 만들 때 기본 글꼴 미리보기를 이만큼까지 기다린다 (첫 창의 글꼴 목록이 비지 않게)
 *   thumbDelayMs = 3000,          // full 미리보기를 시작하기 전에 쉬는 시간 (첫 편집 창이 뜨는 동안 비켜 준다)
 *   worker = true,                // false면 같은 스레드에서 (시험용)
 *   log = () => {},
 * }
 * 손잡이: { allFontsJsPath, thumbnailsDir, fontFileByIndex(n) → 절대 경로 | null,   ← editor-server setOfficeFonts가 쓰는 것
 *          key, dir, stats, fileCount, fontFileInfo(n), createFontStream(n), resolve(rel),
 *          thumbnailPath(name), thumbnailsStage(), thumbnails: Promise<'full'|'plain'|'blank'> }
 */
export async function ensureFonts(opts = {}) {
  const { editorVersion = 'unknown', fontDirs = defaultFontDirs(), profile = 'officekit', log = () => {} } = opts;
  const userData = opts.userData ?? opts.userDataDir;
  const root = opts.cacheRoot ?? (userData ? fontsCacheRoot(userData) : null);
  if (!root) throw Object.assign(new Error('글꼴 캐시 폴더(userData)가 필요합니다'), { code: 'FONTS_ARGS' });
  const listing = listFontFiles(fontDirs);
  const key = fontsCacheKey({ editorVersion, listing, fontDirs, profile });
  const dir = join(root, key);
  if (!inflight.has(dir)) {
    const p = prepare({ ...opts, root, dir, key, listing, editorVersion, profile, log }).catch((err) => { inflight.delete(dir); liveDirs.delete(dir); throw err; });
    inflight.set(dir, p);
  }
  return inflight.get(dir);
}

async function prepare({ root, dir, key, listing, editorVersion, profile, log, worker = true, thumbnails = true, waitPlainMs = 2500, thumbDelayMs = 3000 }) {
  liveDirs.add(dir);
  let manifest = readManifest(dir);
  let fresh = false;
  if (!manifest) {
    fresh = true;
    const t0 = Date.now();
    mkdirSync(root, { recursive: true });
    const tmp = join(root, `${key}.tmp-${process.pid}-${Math.random().toString(36).slice(2, 8)}`);
    const data = { job: 'allfonts', outDir: tmp, listing, profile, key, editorVersion };
    try {
      if (worker) await runWorker(data).promise; else runAllFontsJob(data);
      try { renameSync(tmp, dir); } catch (err) {
        // 다른 창(프로세스)이 먼저 만들었으면 그것을 쓴다
        if (!readManifest(dir)) throw err;
        rmSync(tmp, { recursive: true, force: true });
      }
    } catch (err) {
      try { rmSync(tmp, { recursive: true, force: true }); } catch {}
      throw Object.assign(new Error(`글꼴 목록을 만들지 못했습니다: ${err.message}`), { code: 'FONTS_GENERATE', cause: err });
    }
    manifest = readManifest(dir);
    if (!manifest) throw Object.assign(new Error('글꼴 목록을 만들지 못했습니다 (캐시를 읽을 수 없음)'), { code: 'FONTS_GENERATE' });
    log(`글꼴 목록: 파일 ${manifest.stats.files}개, 가족 ${manifest.stats.families}개, 한글 대체 ${manifest.stats.hangulFallback ?? '없음'} (${Date.now() - t0}ms)`);
    cleanOld(root, key);
  }

  // 미리보기 그림 (뒤에서). 이미 진행 중이면 그것을 쓴다
  let thumbsPromise;
  if (!thumbnails) thumbsPromise = Promise.resolve(stageOf(dir));
  else if (thumbJobs.has(dir)) thumbsPromise = thumbJobs.get(dir);
  else {
    let plainDone; const plainSeen = new Promise((ok) => { plainDone = ok; });
    thumbsPromise = (async () => {
      try {
        if (!existsSync(join(dir, CACHE_FILES.stageDone('plain')))) {
          const data = { job: 'thumbs', dir, stages: ['plain'] };
          if (worker) await runWorker(data).promise; else await runThumbsJob(data);
        }
        plainDone();
        if (thumbDelayMs) await new Promise((ok) => setTimeout(ok, thumbDelayMs));
        if (!existsSync(join(dir, CACHE_FILES.stageDone('full')))) {
          const t0 = Date.now();
          const data = { job: 'thumbs', dir, stages: ['full'] };
          if (worker) await runWorker(data).promise; else await runThumbsJob(data);
          log(`글꼴 미리보기 그림을 만들었습니다 (${Date.now() - t0}ms)`);
        }
      } catch (err) {
        log(`글꼴 미리보기 그림을 만들지 못했습니다: ${err.message}`);
      } finally {
        plainDone();
      }
      return stageOf(dir);
    })();
    thumbJobs.set(dir, thumbsPromise);
    thumbsPromise.finally(() => thumbJobs.delete(dir));
    // 처음 만들 때만: 기본 글꼴 미리보기를 잠깐 기다린다 (첫 편집 창 글꼴 목록이 비지 않게)
    if (fresh && waitPlainMs > 0) {
      let timer;
      await Promise.race([plainSeen, new Promise((ok) => { timer = setTimeout(ok, waitPlainMs); })]);
      clearTimeout(timer);
    }
  }
  return makeHandle(dir, manifest, { thumbnails: thumbsPromise });
}

function stageOf(dir) {
  return ['full', 'plain', 'blank'].find((s) => existsSync(join(dir, CACHE_FILES.stageDone(s)))) ?? 'blank';
}
