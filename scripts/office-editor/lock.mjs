// lock.mjs — 풀어 놓은 filex ui.zip이 고정한 그대로인지 파일마다 확인한다 (4단계).
//
// editor/ 파일은 filex upstream/editor.lock.json("sha256 크기")과 똑같아야 한다. 예외는 filex
// build-app.mjs가 릴리스 묶음을 만들 때 하는 일뿐이다:
//   - web-apps/vendor/socketio/socket.io.min.js : filex 다리(bridge)로 바뀜 — lock과 달라야 하고 filex 머리글로 시작
//   - filex/CHANGES.txt                         : lock의 내용 뒤에 앱 빌드 설명이 덧붙음 — 앞부분이 lock과 같아야 함
//   - themes.json, plugins.json                 : 더해진 빈 목록 — 내용이 정확히 같아야 함
// editor/ 밖은 앱 페이지(index.html, filex/*), filex LICENSE·NOTICE, x2t 두 파일(고정 SHA-256)만 있어야 한다.
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, linkSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

import { EDITOR_LOCK, FILEX, ONLYOFFICE, SOCKET_IO_ORIGINAL, X2T } from './pins.mjs';

export const EDITOR_PREFIX = 'editor/';
export const BRIDGE = SOCKET_IO_ORIGINAL.path;
export const CHANGES = 'filex/CHANGES.txt';
export const ADDED = Object.freeze({
  'themes.json': '{"themes":[]}\n',
  'plugins.json': '{"pluginsData":[]}\n',
});
export const APP_FILES = Object.freeze(['index.html', 'filex/app.css', 'filex/app.js', 'filex/x2t-worker.js', 'LICENSE', 'NOTICE']);
export const X2T_FILES = Object.freeze(Object.keys(X2T.files).map((f) => `x2t/${f}`));

export const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');

/** filex가 만든 스크립트의 첫 줄 (scripts/build-app.mjs의 머리글) */
export const bridgeBanner = (version = FILEX.version) =>
  `/* filex-office-editor ${version} - ${EDITOR_PREFIX}${BRIDGE}, built from src/frame/main.ts`;

/** lock 파일 바이트 → 객체. 고정 SHA-256·모양·편집기 버전을 확인한다 */
export function parseLock(buf, { pinned = EDITOR_LOCK } = {}) {
  if (pinned) {
    if (buf.length !== pinned.bytes || sha256(buf) !== pinned.sha256) {
      throw new Error(`filex ${EDITOR_LOCK.path}가 고정한 것과 다릅니다 (크기 ${buf.length}, 예상 ${pinned.bytes}바이트·SHA-256 ${pinned.sha256.slice(0, 16)}…).`);
    }
  }
  let lock;
  try { lock = JSON.parse(buf.toString('utf8')); } catch (e) { throw new Error(`lock 파일을 읽지 못했습니다: ${e.message}`); }
  if (!lock || typeof lock.files !== 'object') throw new Error('lock 파일에 files가 없습니다.');
  if (pinned && lock.onlyoffice?.build !== ONLYOFFICE.build) throw new Error(`lock의 ONLYOFFICE 빌드가 ${lock.onlyoffice?.build}입니다 (고정: ${ONLYOFFICE.build}).`);
  if (pinned && lock.onlyoffice?.image !== ONLYOFFICE.image) throw new Error('lock의 Document Server 이미지(digest)가 고정 값과 다릅니다.');
  for (const [p, v] of Object.entries(lock.files)) {
    if (typeof v !== 'string' || !/^[0-9a-f]{64} \d+$/.test(v)) throw new Error(`lock 항목 모양이 이상합니다: ${p}`);
    if (p.startsWith('/') || p.split('/').some((s) => s === '' || s === '.' || s === '..')) throw new Error(`lock 경로가 이상합니다: ${p}`);
  }
  if (pinned && Object.keys(lock.files).length !== pinned.files) throw new Error(`lock의 파일 수가 ${Object.keys(lock.files).length}개입니다 (고정: ${pinned.files}개).`);
  return lock;
}

const lockEntry = (v) => {
  const [hash, size] = v.split(' ');
  return { sha256: hash, bytes: Number(size) };
};

/**
 * 묶음 전체를 확인한다.
 * @param {object} o
 * @param {Map<string,{bytes:number, sha256:string}>} o.files  묶음 안 경로 → 크기·SHA-256
 * @param {(p:string)=>Buffer} o.read                          묶음 안 파일 읽기 (다리·CHANGES.txt·더해진 것만 읽음)
 * @param {object} o.lock                                      parseLock 결과
 * @returns {{ editorFiles:number, bridge:{path,bytes,sha256}, changes:{path,bytes,appendedBytes}, added:string[], app:string[] }}
 */
export function verifyUi({ files, read, lock, filexVersion = FILEX.version, x2t = X2T.files }) {
  const problems = [];
  const seen = new Set();
  let bridge = null;
  let changes = null;

  for (const [p, v] of Object.entries(lock.files)) {
    const zp = EDITOR_PREFIX + p;
    const got = files.get(zp);
    seen.add(zp);
    if (!got) { problems.push(`- ${zp} (lock에 있는데 묶음에 없음)`); continue; }
    const want = lockEntry(v);
    if (p === BRIDGE) {
      if (got.sha256 === want.sha256) { problems.push(`~ ${zp} (filex 다리가 아니라 원래 socket.io 파일)`); continue; }
      const head = read(zp).subarray(0, 300).toString('utf8');
      if (!head.startsWith(bridgeBanner(filexVersion))) { problems.push(`~ ${zp} (filex ${filexVersion} 다리의 머리글이 아님)`); continue; }
      bridge = { path: zp, bytes: got.bytes, sha256: got.sha256 };
    } else if (p === CHANGES) {
      const data = read(zp);
      if (data.length <= want.bytes || sha256(data.subarray(0, want.bytes)) !== want.sha256) {
        problems.push(`~ ${zp} (앞부분이 lock의 것과 다름)`);
        continue;
      }
      changes = { path: zp, bytes: data.length, appendedBytes: data.length - want.bytes };
    } else if (got.sha256 !== want.sha256 || got.bytes !== want.bytes) {
      problems.push(`~ ${zp} (lock과 다름)`);
    }
  }

  const added = [];
  const app = [];
  for (const [zp, got] of files) {
    if (seen.has(zp)) continue;
    if (zp.startsWith(EDITOR_PREFIX)) {
      const rel = zp.slice(EDITOR_PREFIX.length);
      if (Object.hasOwn(ADDED, rel)) {
        if (read(zp).toString('utf8') !== ADDED[rel]) problems.push(`~ ${zp} (filex가 더한 빈 목록이 아님)`);
        else added.push(zp);
      } else {
        problems.push(`+ ${zp} (lock에 없음)`);
      }
    } else if (X2T_FILES.includes(zp)) {
      const want = x2t[zp.slice(4)];
      if (got.sha256 !== want) problems.push(`~ ${zp} (고정 SHA-256과 다름 — x2t ${X2T.release})`);
    } else if (APP_FILES.includes(zp)) {
      app.push(zp);
    } else {
      problems.push(`+ ${zp} (모르는 파일)`);
    }
  }
  for (const f of [...APP_FILES, ...X2T_FILES]) if (!files.has(f)) problems.push(`- ${f} (없음)`);
  for (const rel of Object.keys(ADDED)) if (!files.has(EDITOR_PREFIX + rel)) problems.push(`- ${EDITOR_PREFIX}${rel} (없음)`);

  if (problems.length) {
    const shown = problems.slice(0, 40).join('\n  ');
    throw new Error(`묶음 파일이 고정한 것과 다릅니다 (filex ${EDITOR_LOCK.path}·x2t 고정 값):\n  ${shown}${problems.length > 40 ? `\n  … ${problems.length - 40}개 더` : ''}`);
  }
  return { editorFiles: Object.keys(lock.files).length, bridge, changes, added: added.sort(), app: app.sort() };
}

/**
 * filex build-app.mjs에 넘길 'lock 그대로'의 편집기 폴더를 만든다 (5단계).
 * 같은 파일은 하드 링크(안 되면 복사), 다리 자리는 원래 socket.io, CHANGES.txt는 lock 길이만큼.
 */
export function makeLockEditor({ uiDir, lock, outDir, socketIo }) {
  rmSync(outDir, { recursive: true, force: true });
  let linked = 0;
  for (const [p, v] of Object.entries(lock.files)) {
    const want = lockEntry(v);
    const dest = join(outDir, ...p.split('/'));
    mkdirSync(dirname(dest), { recursive: true });
    const src = join(uiDir, 'editor', ...p.split('/'));
    if (p === BRIDGE) {
      if (socketIo.length !== want.bytes || sha256(socketIo) !== want.sha256) throw new Error('받은 원래 socket.io 파일이 lock과 다릅니다.');
      writeFileSync(dest, socketIo);
    } else if (p === CHANGES) {
      writeFileSync(dest, readFileSync(src).subarray(0, want.bytes));
    } else {
      if (!existsSync(src)) throw new Error(`편집기 파일이 없습니다: ${src}`);
      try { linkSync(src, dest); linked++; } catch { copyFileSync(src, dest); }
    }
  }
  return { files: Object.keys(lock.files).length, linked };
}
