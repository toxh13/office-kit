// atomic.mjs — 파일을 원자적으로 쓴다 (동기 판). 도중에 실패해도 원본은 그대로 남는다.
// editor-window.mjs의 atomicWrite와 같은 규칙: 같은 폴더 임시 파일에 쓰고 fsync → 기존 파일이면 덮어 복사(ACL·숨김 유지),
// 안 되면 이름 바꾸기 → 끝에 임시 파일 삭제.
import { randomBytes } from 'node:crypto';
import {
  closeSync, copyFileSync, existsSync, fsyncSync, openSync, readFileSync, renameSync, rmSync, writeFileSync, writeSync,
} from 'node:fs';
import { basename, dirname, join } from 'node:path';

const sleepSync = (ms) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
const busy = (err) => ['EBUSY', 'EPERM', 'EACCES'].includes(err?.code);

// Windows '웹에서 받은 파일' 표시(Zone.Identifier 스트림): 바꿔 끼우면 사라질 수 있어 읽어 두었다가 다시 쓴다 (최선만)
function readZone(path) {
  if (process.platform !== 'win32') return null;
  try { return readFileSync(`${path}:Zone.Identifier`); } catch { return null; }
}
function restoreZone(path, zone) {
  if (!zone) return;
  try {
    const now = readZone(path);
    if (!now || !now.equals(zone)) writeFileSync(`${path}:Zone.Identifier`, zone);
  } catch { /* 표시를 다시 쓰지 못해도 저장은 끝났다 */ }
}

/** data(Buffer)를 path에 원자적으로 쓴다. 실패하면 원본은 그대로이고 첫 오류를 던진다 */
export function writeFileAtomic(path, data, opts = {}) {
  const zone = readZone(path);
  replaceAtomic(path, data, opts);
  restoreZone(path, zone);
}

function replaceAtomic(path, data, { retries = 8 } = {}) {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomBytes(4).toString('hex')}.tmp`);
  const fd = openSync(tmp, 'wx');
  try {
    for (let off = 0; off < data.length;) off += writeSync(fd, data, off, data.length - off);
    fsyncSync(fd);
  } catch (err) {
    try { closeSync(fd); } catch {}
    rmSync(tmp, { force: true });
    throw err;
  }
  closeSync(fd);
  try {
    if (existsSync(path)) {
      for (let i = 0; ; i++) {
        try { copyFileSync(tmp, path); return; } catch (err) {
          if (busy(err) && i < retries) { sleepSync(60 * (i + 1)); continue; }
          // 숨김 속성 등으로 덮어 복사가 안 되거나 도중에 끊겼으면 이름 바꾸기로 (그래도 안 되면 처음 오류)
          try { renameSync(tmp, path); return; } catch { throw err; }
        }
      }
    }
    for (let i = 0; ; i++) {
      try { renameSync(tmp, path); return; } catch (err) {
        if (!busy(err) || i >= retries) throw err;
        sleepSync(60 * (i + 1));
      }
    }
  } finally {
    rmSync(tmp, { force: true });
  }
}

/** from의 내용을 to에 원자적으로 복사한다 */
export function copyFileAtomic(from, to) {
  writeFileAtomic(to, readFileSync(from));
}
