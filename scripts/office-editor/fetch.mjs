// fetch.mjs — 고정한 파일 받기: 이어받기, 3번 다시, 크기·SHA-256 확인 (1·2단계).
import { createHash } from 'node:crypto';
import { createReadStream, existsSync, mkdirSync, renameSync, rmSync, statSync } from 'node:fs';
import { open } from 'node:fs/promises';
import { dirname } from 'node:path';
import { pipeline } from 'node:stream/promises';

const RETRIES = 3;
const IDLE_MS = Number(process.env.OFFICEKIT_FETCH_IDLE_MS) || 60_000;
const RETRY_MS = process.env.OFFICEKIT_FETCH_RETRY_MS !== undefined ? Number(process.env.OFFICEKIT_FETCH_RETRY_MS) : 2000;

export class FetchError extends Error {}

export async function fileSha256(path) {
  const h = createHash('sha256');
  await pipeline(createReadStream(path), h);
  return h.digest('hex');
}

/** path가 크기·SHA-256 그대로인지 */
export async function matches(path, { bytes, sha256 }) {
  try {
    if (statSync(path).size !== bytes) return false;
  } catch {
    return false;
  }
  return (await fileSha256(path)) === sha256;
}

const mb = (n) => `${(n / 1048576).toFixed(1)}MB`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * url을 dest로 받는다. 이미 있고 맞으면 그대로 쓴다. 맞지 않으면 지우고 멈춘다.
 * @param {{url:string, dest:string, bytes:number, sha256:string, label:string, log?:(s:string)=>void, offline?:boolean}} o
 * @returns {Promise<'cache'|'download'>}
 */
export async function fetchPinned({ url, dest, bytes, sha256, label, log = () => {}, offline = false, retryMs = RETRY_MS, idleMs = IDLE_MS }) {
  if (existsSync(dest)) {
    if (await matches(dest, { bytes, sha256 })) return 'cache';
    log(`캐시의 ${label}이(가) 고정 값과 달라 지우고 다시 받습니다.`);
    rmSync(dest, { force: true });
  }
  if (offline) throw new FetchError(`${label}이(가) 캐시에 없습니다 (--offline). 인터넷이 되는 곳에서 한 번 받으세요: ${url}`);
  mkdirSync(dirname(dest), { recursive: true });
  const part = `${dest}.partial`;
  let lastErr = null;
  for (let attempt = 0; attempt <= RETRIES; attempt++) {
    let have = existsSync(part) ? statSync(part).size : 0;
    if (have > bytes) { rmSync(part, { force: true }); have = 0; }
    if (have === bytes) break;
    if (attempt > 0) {
      log(`  ${label} 다시 받기 ${attempt}/${RETRIES}${have ? ` (${mb(have)}부터 이어받기)` : ''} — 앞 오류: ${lastErr?.message || lastErr}`);
      await sleep(retryMs * attempt);
    } else if (have) {
      log(`  ${label}: 지난번에 받다 만 것(${mb(have)})부터 이어받습니다.`);
    }
    try {
      await getOnce({ url, part, have, bytes, label, log, idleMs });
      lastErr = null;
    } catch (e) {
      lastErr = e;
    }
    if (!lastErr && existsSync(part) && statSync(part).size === bytes) break;
  }
  const got = existsSync(part) ? statSync(part).size : 0;
  if (got !== bytes) {
    if (got > bytes) rmSync(part, { force: true });
    throw new FetchError(`${label}을(를) 받지 못했습니다 (${mb(got)} / ${mb(bytes)}): ${lastErr?.message || lastErr || '크기가 다름'}\n  주소: ${url}\n  인터넷 연결을 확인하고 다시 실행하면 이어서 받습니다.`);
  }
  const hash = await fileSha256(part);
  if (hash !== sha256) {
    rmSync(part, { force: true });
    throw new FetchError(`받은 ${label}의 SHA-256이 고정 값과 다릅니다 — 파일을 지우고 멈춥니다.\n  받음 ${hash}\n  고정 ${sha256}\n  원본이 바뀌었거나 전송 중 손상됐습니다. 원본이 바뀌었다면 내용을 직접 확인한 뒤 scripts/office-editor/pins.mjs를 고치세요.`);
  }
  renameSync(part, dest);
  return 'download';
}

async function getOnce({ url, part, have, bytes, label, log, idleMs }) {
  const ac = new AbortController();
  let timer = setTimeout(() => ac.abort(new Error(`${idleMs / 1000}초 동안 받은 것이 없음`)), idleMs);
  const kick = () => {
    clearTimeout(timer);
    timer = setTimeout(() => ac.abort(new Error(`${idleMs / 1000}초 동안 받은 것이 없음`)), idleMs);
  };
  try {
    const headers = { 'user-agent': 'office-kit-build (+https://github.com/toxh13/office-kit)' };
    if (have) headers.range = `bytes=${have}-`;
    const res = await fetch(url, { headers, redirect: 'follow', signal: ac.signal });
    let append = false;
    if (have && res.status === 206) {
      const range = res.headers.get('content-range') || '';
      if (!range.startsWith(`bytes ${have}-`)) throw new Error(`이어받기 응답 범위가 이상합니다 (${range})`);
      append = true;
    } else if (res.status === 200) {
      if (have) log(`  서버가 이어받기를 받지 않아 처음부터 받습니다.`);
    } else {
      throw new Error(`HTTP ${res.status} ${res.statusText}`);
    }
    if (!res.body) throw new Error('응답 본문이 없음');
    const total = bytes;
    let done = append ? have : 0;
    let nextReport = done + total / 10;
    // 조각마다 파일에 쓴 뒤 다음 조각을 읽는다 → 끊겨도 받은 만큼은 남아 다음에 이어받는다
    const fh = await open(part, append ? 'a' : 'w');
    try {
      const reader = res.body.getReader();
      for (;;) {
        const { done: end, value } = await reader.read();
        if (end) break;
        kick();
        if (done + value.length > total) throw new Error('고정한 크기보다 큼');
        await fh.write(value);
        done += value.length;
        if (total > 4 * 1048576 && done >= nextReport) {
          log(`  ${label} ${mb(done)} / ${mb(total)}`);
          nextReport += total / 10;
        }
      }
    } finally {
      await fh.close();
    }
  } finally {
    clearTimeout(timer);
  }
}
