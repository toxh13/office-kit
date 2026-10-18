// editor-window.mjs — 한글(HWP/HWPX) 편집 창. 문서마다 창 하나, 내장 rhwp-studio를 로컬 서버(127.0.0.1·127.0.0.2)에서 띄운다 (인터넷 없이).
//   Electron 모듈은 주입받는다 (main.mjs가 부르고, 시험 하니스도 같은 길로 부른다).
import { createHash, randomBytes } from 'node:crypto';
import { closeSync, copyFileSync, existsSync, fsyncSync, openSync, readFileSync, renameSync, rmSync, statSync, watch, writeSync } from 'node:fs';
import { isIP } from 'node:net';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { editorAvailable, startEditorServer } from './editor-server.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** 편집할 수 있는 형식 (hwt는 저장할 때 HWP로 내보낸다) */
export const EDITABLE_EXTS = ['hwp', 'hwpx', 'hwt'];
const SAVE_AS_EXTS = ['hwp', 'hwpx'];
const PARTITION = 'officekit-editor'; // 'persist:'가 아님 → 저장소·서비스 워커가 디스크에 남지 않는다
const TITLE_SUFFIX = ' - 한글 편집 (Office Kit)';
const MAX_BYTES = 512 * 1024 * 1024;
// 편집기가 쓰는 권한만: 복사(클립보드 쓰기), 설치된 글꼴 읽기(옵션 창에서 사용자가 누를 때).
// 나머지(클립보드 읽기 API·카메라·위치·알림 …)는 거절. 메뉴의 '붙이기'는 enableDeprecatedPaste로 된다
const ALLOWED_PERMISSIONS = new Set(['clipboard-sanitized-write', 'local-fonts']);
const REPLY_TIMEOUT_MS = 3000;
const SAVE_TIMEOUT_MS = 180_000;
// 아무 곳에도 이어지지 않는 주소: 맞춤법 사전 내려받기·프록시를 여기로 보내 인터넷에 나가지 못하게 한다
const DEAD_URL = 'http://127.0.0.1:9/';
const DEAD_PROXY = 'socks5://127.0.0.1:9';

const extOf = (p) => extname(p).slice(1).toLowerCase();
const keyOf = (p) => (process.platform === 'win32' ? resolve(p).toLowerCase() : resolve(p));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fail = (message, code) => Object.assign(new Error(message), { expected: true }, code ? { code } : {});

// 내보낸 바이트가 형식과 맞는지 (HWP·HWT = OLE 복합 문서, HWPX = ZIP)
const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP = [0x50, 0x4b, 0x03, 0x04];
const startsWith = (bytes, sig) => bytes.length >= sig.length && sig.every((b, i) => bytes[i] === b);
export function bytesMatchExt(bytes, ext) {
  if (!(bytes instanceof Uint8Array)) return false;
  return ext === 'hwpx' ? startsWith(bytes, ZIP) : startsWith(bytes, OLE);
}

function hashOf(buf) { return createHash('sha1').update(buf).digest('hex'); }
function statOf(p) {
  try { const st = statSync(p); return { mtimeMs: st.mtimeMs, size: st.size, readOnly: (st.mode & 0o200) === 0 }; } catch { return null; }
}

/** 파일 오류 → 한국어 */
function saveError(err, name) {
  const code = err?.code ?? '';
  if (['EBUSY', 'EPERM', 'EACCES'].includes(code)) {
    return fail(`‘${name}’ 파일을 다른 프로그램이 쓰고 있거나 쓸 권한이 없어 저장하지 못했습니다. 그 프로그램을 닫고 다시 저장하거나, 다른 이름으로 저장하세요.`, 'LOCKED');
  }
  if (code === 'ENOSPC') return fail('디스크 공간이 부족해 저장하지 못했습니다.', 'NO_SPACE');
  if (code === 'ENOENT') return fail('문서가 있던 폴더를 찾을 수 없어 저장하지 못했습니다. 다른 이름으로 저장하세요.', 'NOT_FOUND');
  if (code === 'EROFS') return fail('읽기 전용 드라이브라 저장할 수 없습니다. 다른 이름으로 저장하세요.', 'READ_ONLY');
  return fail(`저장하지 못했습니다: ${err?.message ?? err}`, code || 'SAVE_FAILED');
}

/**
 * 같은 폴더 임시 파일에 다 쓰고(fsync) 문서 자리에 옮긴다 (백신·미리보기가 잠깐 잡으면 몇 번 다시).
 * - 새 파일: 이름 바꾸기
 * - 있는 파일: 내용만 덮어 복사 → 파일의 보안 설정(ACL)·하드 링크·추가 스트림이 그대로 남는다
 *   (이름 바꾸기는 폴더의 기본 권한을 물려받아 '나만 읽기' 같은 설정이 풀린다). 복사가 끝내 안 되면 이름 바꾸기로
 */
export async function atomicWrite(path, bytes, { retries = 8 } = {}) {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomBytes(4).toString('hex')}.tmp`);
  const fd = openSync(tmp, 'wx');
  try {
    for (let off = 0; off < bytes.length;) off += writeSync(fd, bytes, off, bytes.length - off);
    fsyncSync(fd);
  } catch (err) {
    closeSync(fd);
    rmSync(tmp, { force: true });
    throw err;
  }
  closeSync(fd);
  const busy = (err) => ['EBUSY', 'EPERM', 'EACCES'].includes(err?.code);
  try {
    if (existsSync(path)) {
      for (let i = 0; ; i++) {
        try { copyFileSync(tmp, path); return; } catch (err) {
          if (busy(err) && i < retries) { await sleep(60 * (i + 1)); continue; }
          // 숨김 속성 등으로 덮어 복사가 안 되거나 도중에 끊겼으면 이름 바꾸기로 (그래도 안 되면 처음 오류)
          try { renameSync(tmp, path); return; } catch { throw err; }
        }
      }
    }
    for (let i = 0; ; i++) {
      try { renameSync(tmp, path); return; } catch (err) {
        if (!busy(err) || i >= retries) throw err;
        await sleep(60 * (i + 1));
      }
    }
  } finally {
    rmSync(tmp, { force: true });
  }
}

/** 브라우저로 넘기면 안 되는 주소: 이 PC·사설망 (편집기 서버 주소 포함) */
export function isPrivateHost(hostname) {
  let h = String(hostname ?? '').toLowerCase().replace(/^\[|\]$/g, '');
  if (!h || h === 'localhost' || h.endsWith('.localhost') || h === '0.0.0.0') return true;
  const dotted = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(h);
  if (dotted) h = dotted[1];
  const hex = /^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(h);
  if (hex) { const a = parseInt(hex[1], 16); const b = parseInt(hex[2], 16); h = [a >> 8, a & 255, b >> 8, b & 255].join('.'); }
  if (isIP(h) === 4) {
    const [a, b] = h.split('.').map(Number);
    return a === 127 || a === 10 || a === 0 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31)
      || (a === 192 && b === 168) || (a === 100 && b >= 64 && b <= 127);
  }
  if (isIP(h) === 6) return h === '::1' || h === '::' || /^f[cd]/.test(h) || /^fe[89ab]/.test(h);
  return false;
}

/**
 * 한글 편집 창 관리자.
 * @param {object} o
 * @param {import('electron').App} o.app
 * @param {typeof import('electron').BrowserWindow} o.BrowserWindow
 * @param {import('electron').IpcMain} o.ipcMain
 * @param {typeof import('electron').session} o.session
 * @param {import('electron').Dialog} o.dialog
 * @param {import('electron').Shell} o.shell
 * @param {string} o.appRoot              앱 폴더 (vendor/, src/app/editor 기준)
 * @param {string} [o.userData]           (지금은 쓰지 않음, 예약)
 * @param {() => ({ backup(path): any }) | { backup(path): any }} [o.getBackups]  저장 전에 원본 백업
 * @param {(path: string) => void} [o.onSaved]      저장(다른 이름 포함) 뒤
 * @param {(path: string, dirty: boolean) => void} [o.onDirtyChange]
 * @param {{ studio?: string, sdk?: string, host?: string }} [o.dirs]  편집기 폴더 바꾸기 (시험·임시 묶음)
 * @param {object} [o.nativeTheme]  창 배경색 고르기용
 * @param {string} [o.icon]
 * @param {boolean} [o.testHooks]   host 페이지에 시험용 고리를 연다 (?test=1)
 * @param {(...args) => void} [o.log]
 */
export function createEditorManager({
  app, BrowserWindow, ipcMain, session, dialog, shell, appRoot, userData, getBackups, onSaved, onDirtyChange,
  dirs, nativeTheme, icon, testHooks = false, log = (...a) => console.log('[editor]', ...a),
} = {}) {
  void userData;
  const root = resolve(appRoot ?? join(here, '..', '..'));
  const dirOverrides = { ...(dirs ?? {}) };
  const entries = new Map();    // 문서 키 → 창 정보
  const byWc = new Map();       // webContents id → 창 정보
  const pending = new Map();    // 요청 id → { resolve, timer, entry }
  let server = null; let serverStarting = null;
  let sessionReady = false;
  let quitting = false;
  let seq = 0;

  const available = () => editorAvailable(root, dirOverrides, { dev: !app?.isPackaged });

  // ── 서버 · 세션 ───────────────────────────────────────

  async function ensureServer() {
    if (server) return server;
    // 세션(요청 막기·프록시)을 다 갖춘 뒤에 창을 연다
    serverStarting ??= startEditorServer({ root, dirs: dirOverrides })
      .then(async (s) => {
        try { await setupSession(s); } catch (err) { await s.close().catch(() => {}); throw err; }
        server = s;
        return s;
      })
      .finally(() => { serverStarting = null; });
    return serverStarting;
  }

  function allowedUrl(url, srv = server) {
    if (/^(blob|data):/i.test(url) || /^(devtools|chrome-devtools):/i.test(url) || url === 'about:blank' || url === 'about:srcdoc') return true;
    if (!srv) return false;
    let u;
    try { u = new URL(url); } catch { return false; }
    return u.protocol === 'http:' && (u.origin === srv.origin || u.origin === srv.studioOrigin);
  }

  async function setupSession(srv) {
    if (sessionReady) return;
    const ses = session.fromPartition(PARTITION);
    // 맞춤법 사전: 끄기만 해도 Chromium이 인터넷에서 사전(ko-3-0.bdic, 11MB)을 받으므로 받을 곳을 죽은 주소로
    try { ses.setSpellCheckerDictionaryDownloadURL(DEAD_URL); } catch {}
    try { ses.setSpellCheckerEnabled(false); } catch {}
    // HTTP 밖의 연결(WebRTC TURN-TCP·WebTransport 등)은 요청 막기에 안 잡히므로 죽은 프록시로 보낸다 (이 PC 주소만 바로)
    await ses.setProxy({ mode: 'fixed_servers', proxyRules: DEAD_PROXY, proxyBypassRules: '127.0.0.1;127.0.0.2' });
    // 내 서버 밖으로 나가는 요청은 모두 막는다 (CDN 글꼴·외부 그림 등)
    ses.webRequest.onBeforeRequest({ urls: ['<all_urls>'] }, (details, cb) => {
      const ok = allowedUrl(details.url, server ?? srv);
      if (!ok) log('막은 요청:', details.url.slice(0, 200));
      cb({ cancel: !ok });
    });
    ses.setPermissionRequestHandler((wc, permission, cb, details) => {
      const ok = ALLOWED_PERMISSIONS.has(permission) && byWc.has(wc?.id);
      if (!ok) log('거절한 권한:', permission, details?.requestingUrl ?? '');
      cb(ok);
    });
    ses.setPermissionCheckHandler((wc, permission) => ALLOWED_PERMISSIONS.has(permission) && (!wc || byWc.has(wc.id)));
    try { ses.setDevicePermissionHandler(() => false); } catch {}
    try { ses.setDisplayMediaRequestHandler((_req, cb) => cb({})); } catch {}
    ses.on('will-download', (e) => e.preventDefault());
    sessionReady = true;
  }

  // ── IPC ──────────────────────────────────────────────

  /** 보낸 쪽이 편집 창의 host 페이지(최상위 틀)인지 → 그 창 정보 */
  function senderEntry(e) {
    const entry = byWc.get(e.sender?.id);
    if (!entry || !entry.win || entry.win.isDestroyed()) return null;
    const frame = e.senderFrame;
    if (!frame || frame !== entry.win.webContents.mainFrame) return null;
    try { if (new URL(frame.url).origin !== server?.origin) return null; } catch { return null; }
    return entry;
  }

  function handle(channel, fn) {
    ipcMain.handle(channel, async (e, ...args) => {
      try {
        const entry = senderEntry(e);
        if (!entry) throw fail('허용되지 않은 요청입니다');
        return { ok: true, value: await fn(entry, ...args) };
      } catch (err) {
        if (!err?.expected) log(`[${channel}]`, err);
        return { ok: false, error: err?.message ?? String(err), code: err?.code ?? null };
      }
    });
  }
  function on(channel, fn) {
    ipcMain.on(channel, (e, ...args) => {
      const entry = senderEntry(e);
      if (!entry) return;
      try { fn(entry, ...args); } catch (err) { log(`[${channel}]`, err); }
    });
  }

  handle('editor:get-document', (entry) => {
    const ok = available();
    const st = statOf(entry.path);
    if (!st) throw fail(`문서를 찾을 수 없습니다: ${basename(entry.path)}`, 'NOT_FOUND');
    if (st.size > MAX_BYTES) throw fail('문서가 너무 커서 편집기로 열 수 없습니다 (512MB 넘음)', 'TOO_BIG');
    const buf = readFileSync(entry.path);
    entry.known = { mtimeMs: st.mtimeMs, size: buf.length, hash: hashOf(buf), missing: false };
    return {
      name: basename(entry.path), ext: extOf(entry.path), size: buf.length, mtimeMs: st.mtimeMs, readOnly: st.readOnly,
      bytes: new Uint8Array(buf), // 복사본 (Buffer 풀의 다른 내용이 함께 가지 않게)
      available: ok.available, reason: ok.reason, studioUrl: server.studioUrl,
    };
  });

  handle('editor:save', async (entry, bytes) => saveTo(entry, entry.path, bytes, extOf(entry.path) === 'hwpx' ? 'hwpx' : 'hwp'));

  handle('editor:save-as', async (entry, bytes, ext) => {
    if (!SAVE_AS_EXTS.includes(ext)) throw fail('한글 문서(.hwp, .hwpx)로만 저장할 수 있습니다');
    checkBytes(bytes, ext);
    const stem = basename(entry.path).replace(/\.[^.]+$/, '');
    const r = await dialog.showSaveDialog(entry.win, {
      title: '다른 이름으로 저장',
      defaultPath: join(dirname(entry.path), `${stem}.${ext}`),
      filters: [{ name: ext === 'hwpx' ? '한글 문서 (HWPX)' : '한글 문서 (HWP)', extensions: [ext] }],
      properties: ['showOverwriteConfirmation'],
    });
    if (r.canceled || !r.filePath) return { canceled: true };
    let target = resolve(r.filePath);
    if (extOf(target) !== ext) {
      target = `${target}.${ext}`;
      // 덮어쓰기 확인은 사용자가 적은 이름으로 받았으므로, 확장자를 붙인 이름이 이미 있으면 다시 묻는다
      if (statOf(target)) {
        const { response } = await dialog.showMessageBox(entry.win, {
          type: 'warning', title: '다른 이름으로 저장',
          message: `‘${basename(target)}’ 파일이 이미 있습니다. 바꿀까요?`,
          detail: '바꾸면 원래 파일의 내용은 백업 폴더에만 남습니다.',
          buttons: ['바꾸기', '취소'], defaultId: 1, cancelId: 1, noLink: true,
        });
        if (response !== 0) return { canceled: true };
      }
    }
    const other = entries.get(keyOf(target));
    if (other && other !== entry) throw fail('그 문서는 다른 편집 창에서 열려 있습니다. 그 창을 닫고 다시 저장하세요.', 'IN_USE');
    const res = await saveTo(entry, target, bytes, ext);
    if (keyOf(target) !== entry.key) rekey(entry, target);
    return res;
  });

  on('editor:dirty', (entry, dirty) => {
    dirty = dirty === true;
    if (entry.dirty === dirty) return;
    entry.dirty = dirty;
    updateTitle(entry);
    try { onDirtyChange?.(entry.path, dirty); } catch {}
  });
  on('editor:log', (entry, message) => log(`(${basename(entry.path)})`, String(message ?? '').slice(0, 2000)));
  on('editor:close', (entry) => { entry.win.close(); });
  on('editor:reply', (entry, msg) => {
    const p = pending.get(msg?.id);
    if (!p || p.entry !== entry) return;
    pending.delete(msg.id);
    clearTimeout(p.timer);
    p.resolve(typeof msg.value === 'boolean' ? msg.value : null);
  });

  /** host 페이지에 묻고 답(boolean | null)을 기다린다. 시간이 지나거나 창이 없어지면 null */
  function ask(entry, channel, payload = {}, timeoutMs = REPLY_TIMEOUT_MS) {
    if (!entry.win || entry.win.isDestroyed()) return Promise.resolve(null);
    const id = `${++seq}`;
    return new Promise((res) => {
      const timer = setTimeout(() => { pending.delete(id); res(null); }, timeoutMs);
      pending.set(id, { resolve: res, timer, entry });
      try { entry.win.webContents.send(channel, { ...payload, id }); } catch { clearTimeout(timer); pending.delete(id); res(null); }
    });
  }

  // ── 저장 ─────────────────────────────────────────────

  function checkBytes(bytes, ext) {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) throw fail('저장할 내용이 비어 있습니다');
    if (bytes.length > MAX_BYTES) throw fail('문서가 너무 커서 저장할 수 없습니다 (512MB 넘음)');
    if (!bytesMatchExt(bytes, ext)) throw fail(`편집기가 내보낸 내용이 .${ext} 형식이 아니라 저장하지 않았습니다`, 'BAD_FORMAT');
  }

  function backupIfExists(path) {
    if (!getBackups || !statOf(path)) return null;
    try {
      const store = typeof getBackups === 'function' ? getBackups() : getBackups;
      const b = store?.backup?.(path);
      return b?.file ?? b?.path ?? null;
    } catch (err) {
      log('백업 실패:', err?.message ?? err);
      return null;
    }
  }

  async function saveTo(entry, target, bytes, ext) {
    checkBytes(bytes, ext);
    const name = basename(target);
    const before = statOf(target);
    if (before?.readOnly) throw fail(`‘${name}’ 파일은 읽기 전용이라 저장할 수 없습니다. 다른 이름으로 저장하세요.`, 'READ_ONLY');
    if (entry.saving) throw fail('저장하는 중입니다. 잠시 뒤 다시 시도하세요.', 'BUSY');
    entry.saving = true;
    try {
      // 덮어쓰기 전의 파일을 백업해 둔다 (실수로 저장해도 백업 폴더에서 되찾게)
      const backup = backupIfExists(target);
      const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      try { await atomicWrite(target, buf); } catch (err) { throw saveError(err, name); }
      const st = statOf(target);
      if (keyOf(target) === entry.key) entry.known = { mtimeMs: st?.mtimeMs ?? 0, size: buf.length, hash: hashOf(buf), missing: false };
      log(`저장: ${target} (${buf.length}바이트)`);
      try { onSaved?.(target); } catch (err) { log('onSaved 오류:', err); }
      return { name, ext: extOf(target), size: buf.length, mtimeMs: st?.mtimeMs ?? null, backup };
    } finally {
      entry.saving = false;
    }
  }

  // 다른 이름으로 저장 뒤: 이 창은 새 파일을 편집한다
  function rekey(entry, target) {
    stopWatch(entry);
    entries.delete(entry.key);
    entry.path = target;
    entry.key = keyOf(target);
    entries.set(entry.key, entry);
    const buf = readFileSync(target);
    const st = statOf(target);
    entry.known = { mtimeMs: st?.mtimeMs ?? 0, size: buf.length, hash: hashOf(buf), missing: false };
    updateTitle(entry);
    startWatch(entry);
  }

  // ── 바깥 변경 감시 ────────────────────────────────────
  // 편집기가 새 파일로 바꿔치기하는 경우도 잡도록 폴더를 지켜본다. 내용(해시)이 같으면 알리지 않는다 (자기 저장 포함)

  /**
   * force: 내용이 같아도 알린다 ('원래대로'를 눌렀는데 편집 창에 저장 안 한 편집이 있을 때 → 편집 창이 다시 불러올지 묻는다)
   * reason: 편집 창 안내 문구용 ('restore' = 주 창의 원래대로)
   */
  function checkExternal(entry, { force = false, reason = null } = {}) {
    if (!entry.win || entry.win.isDestroyed() || entry.saving || !entry.known) return false;
    const st = statOf(entry.path);
    if (!st) {
      if (entry.known.missing) return false;
      entry.known = { ...entry.known, missing: true };
      entry.win.webContents.send('editor:external-change', { missing: true, size: 0, mtimeMs: null });
      return true;
    }
    if (!force && !entry.known.missing && st.mtimeMs === entry.known.mtimeMs && st.size === entry.known.size) return false;
    let hash;
    try { hash = hashOf(readFileSync(entry.path)); } catch { return false; } // 쓰는 중이면 다음 알림에서
    const { hash: before, missing: wasMissing } = entry.known;
    entry.known = { mtimeMs: st.mtimeMs, size: st.size, hash, missing: false };
    if (!force && hash === before && !wasMissing) return false;
    entry.win.webContents.send('editor:external-change', { missing: false, size: st.size, mtimeMs: st.mtimeMs, reason });
    return true;
  }

  function startWatch(entry) {
    const name = basename(entry.path).toLowerCase();
    try {
      entry.watcher = watch(dirname(entry.path), { persistent: false }, (_ev, file) => {
        if (file && String(file).toLowerCase() !== name) return;
        clearTimeout(entry.watchTimer);
        entry.watchTimer = setTimeout(() => checkExternal(entry), 600);
      });
      entry.watcher.on('error', () => {});
    } catch { entry.watcher = null; }
  }
  function stopWatch(entry) {
    clearTimeout(entry.watchTimer);
    try { entry.watcher?.close(); } catch {}
    entry.watcher = null;
  }

  // ── 창 ───────────────────────────────────────────────

  function titleOf(entry) { return `${entry.dirty ? '* ' : ''}${basename(entry.path)}${TITLE_SUFFIX}`; }
  function updateTitle(entry) { if (entry.win && !entry.win.isDestroyed()) entry.win.setTitle(titleOf(entry)); }

  function focus(entry) {
    const w = entry.win;
    if (!w || w.isDestroyed()) return;
    if (w.isMinimized()) w.restore();
    w.show();
    w.focus();
  }

  function isEditorUrl(url) {
    try {
      const u = new URL(url);
      return u.origin === server?.studioOrigin && u.pathname.startsWith('/studio/');
    } catch { return false; }
  }

  /** 닫기 흐름: 저장 안 한 변경이 있으면 저장 / 저장 안 함 / 취소 → 닫혔으면 true */
  async function requestClose(entry) {
    if (!entry.win || entry.win.isDestroyed()) return true;
    if (entry.closing) return entry.closing;
    entry.closing = (async () => {
      const answer = await ask(entry, 'editor:request-close');
      const dirty = answer ?? entry.dirty;
      if (dirty) {
        focus(entry);
        const { response } = await dialog.showMessageBox(entry.win, {
          type: 'warning', title: '한글 편집',
          message: `‘${basename(entry.path)}’의 바뀐 내용을 저장할까요?`,
          detail: '저장하지 않으면 마지막으로 저장한 뒤 고친 내용이 사라집니다.',
          buttons: ['저장', '저장 안 함', '취소'], defaultId: 0, cancelId: 2, noLink: true,
        });
        if (response === 2) return false;
        if (response === 0) {
          const saved = await ask(entry, 'editor:command', { command: 'save' }, SAVE_TIMEOUT_MS);
          if (saved !== true) return false; // 저장 실패는 창에 알림 띠로 보인다
        }
      }
      if (!entry.win || entry.win.isDestroyed()) return true;
      const closed = new Promise((r) => entry.win.once('closed', () => r(true)));
      entry.allowClose = true;
      entry.win.close();
      return closed;
    })();
    try { return await entry.closing; } finally { entry.closing = null; }
  }

  // ── 문서 안 링크 ─────────────────────────────────────
  // 사람이 누른 링크든 스크립트가 연 창이든, 이 PC·사설망 주소는 열지 않고 나머지는 물어본 뒤 기본 브라우저로

  let linkAsking = false;
  async function openLink(entry, url) {
    let u;
    try { u = new URL(String(url)); } catch { return false; }
    if (!['http:', 'https:', 'mailto:'].includes(u.protocol) || u.username || u.password) { log('열지 않은 주소:', String(url).slice(0, 200)); return false; }
    const local = u.protocol !== 'mailto:' && (isPrivateHost(u.hostname)
      || (server && [server.origin, server.studioOrigin].includes(u.origin)));
    if (local) { log('열지 않은 주소(이 PC·사설망):', u.href.slice(0, 200)); return false; }
    if (linkAsking || !entry.win || entry.win.isDestroyed()) return false; // 한 번에 하나만 묻는다
    linkAsking = true;
    try {
      const { response } = await dialog.showMessageBox(entry.win, {
        type: 'question', title: '한글 편집',
        message: u.protocol === 'mailto:' ? '이 주소로 메일을 쓸까요?' : '이 주소를 기본 브라우저로 열까요?',
        detail: u.href.length > 500 ? `${u.href.slice(0, 500)}…` : u.href,
        buttons: ['열기', '취소'], defaultId: 0, cancelId: 1, noLink: true,
      });
      if (response !== 0) return false;
      await shell?.openExternal(u.href);
      return true;
    } catch (err) {
      log('링크를 열지 못했습니다:', err?.message ?? err);
      return false;
    } finally {
      linkAsking = false;
    }
  }

  /** studio가 연 빈 창: 주소를 바꾸려는 순간 막고 그 주소를 openLink로 넘긴 뒤 창을 없앤다 */
  function catchLinkWindow(entry, child) {
    const cwc = child.webContents;
    let taken = false;
    const take = (e, url) => {
      if (!url || url === 'about:blank') return;
      try { e?.preventDefault?.(); } catch {}
      if (taken) return;
      taken = true;
      void openLink(entry, url);
      setImmediate(() => { if (!child.isDestroyed()) child.destroy(); });
    };
    cwc.on('will-navigate', (e, url) => take(e, e?.url ?? url));
    cwc.on('will-frame-navigate', (e) => take(e, e?.url));
    cwc.on('will-redirect', (e, url) => take(e, e?.url ?? url));
    cwc.setWindowOpenHandler(() => ({ action: 'deny' }));
    // 아무 데도 가지 않는 빈 창은 곧 없앤다
    const timer = setTimeout(() => { if (!child.isDestroyed()) child.destroy(); }, 5000);
    child.once('closed', () => clearTimeout(timer));
  }

  function createWindow(entry) {
    const dark = !!nativeTheme?.shouldUseDarkColors;
    const win = new BrowserWindow({
      width: 1280, height: 900, minWidth: 720, minHeight: 520,
      title: titleOf(entry), show: false,
      backgroundColor: dark ? '#15171c' : '#f3f4f6',
      autoHideMenuBar: true,
      ...(icon ? { icon } : {}),
      webPreferences: {
        preload: join(here, 'editor', 'editor-preload.cjs'),
        contextIsolation: true, sandbox: true, nodeIntegration: false, nodeIntegrationInSubFrames: false,
        partition: PARTITION,
        webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
        // studio 메뉴·오른쪽 클릭의 '붙이기'는 execCommand('paste')를 쓴다 (Ctrl+V는 이것 없이도 된다)
        enableDeprecatedPaste: true,
        devTools: !app?.isPackaged,
      },
    });
    entry.win = win;
    const wc = win.webContents;
    byWc.set(wc.id, entry);
    // 앱 메뉴 단축키(Ctrl+O, F5 …)가 이 창에서 주 창 일을 하지 않게
    win.removeMenu();
    win.once('ready-to-show', () => win.show());
    win.on('page-title-updated', (e) => e.preventDefault());

    // WebRTC가 프록시를 거치지 않는 UDP(STUN)로 밖에 나가지 못하게 (TCP는 세션의 죽은 프록시로 간다)
    try { wc.setWebRTCIPHandlingPolicy('disable_non_proxied_udp'); } catch {}

    // 새 창 막기: 웹 주소는 물어본 뒤 기본 브라우저로.
    // studio의 문서 링크는 빈 창(about:blank)을 연 뒤 그 창의 주소를 바꾸므로, 빈 창만 숨겨서 열고 이동을 가로챈다
    wc.setWindowOpenHandler(({ url }) => {
      if (url === 'about:blank') {
        return {
          action: 'allow',
          overrideBrowserWindowOptions: {
            show: false, skipTaskbar: true, width: 200, height: 100,
            webPreferences: {
              partition: PARTITION, sandbox: true, contextIsolation: true, nodeIntegration: false,
              javascript: false, webviewTag: false, spellcheck: false, devTools: false,
            },
          },
        };
      }
      void openLink(entry, url);
      return { action: 'deny' };
    });
    wc.on('did-create-window', (child) => catchLinkWindow(entry, child));
    wc.on('will-navigate', (e) => e.preventDefault());
    wc.on('will-frame-navigate', (e) => {
      if (e.isMainFrame) { e.preventDefault(); return; }
      const u = e.url;
      if (u === 'about:blank' || u === 'about:srcdoc' || isEditorUrl(u)) return;
      e.preventDefault();
    });
    wc.on('will-attach-webview', (e) => e.preventDefault());
    // studio의 beforeunload(저장 안 함 경고)는 무시하고 닫기 확인은 여기서 한다
    wc.on('will-prevent-unload', (e) => e.preventDefault());
    // 단축키: studio(iframe) 안에 커서가 있어도 먼저 받는다
    wc.on('before-input-event', (e, input) => {
      if (input.type !== 'keyDown') return;
      const code = input.code || '';
      // 개발 중에만: F12 / Ctrl+Shift+I 개발자 도구 (메뉴를 뗐으므로)
      if (!app?.isPackaged && (code === 'F12' || ((input.control || input.meta) && input.shift && code === 'KeyI'))) {
        e.preventDefault();
        wc.toggleDevTools();
        return;
      }
      if (!(input.control || input.meta) || input.alt) return;
      if (code === 'KeyS') {
        e.preventDefault();
        void ask(entry, 'editor:command', { command: input.shift ? 'save-as' : 'save' }, SAVE_TIMEOUT_MS);
      } else if (code === 'KeyW' && !input.shift) {
        e.preventDefault();
        win.close();
      }
    });
    wc.on('render-process-gone', (_e, d) => {
      log('편집 창 화면 종료:', d.reason);
      if (d.reason === 'clean-exit' || win.isDestroyed()) return;
      entry.dirty = false;
      updateTitle(entry);
      dialog.showMessageBox(win, {
        type: 'error', title: '한글 편집', noLink: true, buttons: ['확인'],
        message: '편집 창이 예기치 않게 멈췄습니다.',
        detail: '마지막으로 저장한 문서를 다시 엽니다. 저장하지 않은 내용은 사라졌을 수 있습니다.',
      }).then(() => { if (!win.isDestroyed()) wc.loadURL(hostUrl()); }).catch(() => {});
    });

    win.on('close', (e) => {
      if (entry.allowClose) return;
      e.preventDefault();
      void requestClose(entry).then((closed) => {
        if (!closed) quitting = false;
        else if (quitting) app?.quit();
      });
    });
    win.on('closed', () => {
      stopWatch(entry);
      byWc.delete(wc.id);
      if (entries.get(entry.key) === entry) entries.delete(entry.key);
      for (const [id, p] of pending) if (p.entry === entry) { clearTimeout(p.timer); pending.delete(id); p.resolve(null); }
      entry.win = null;
    });

    wc.loadURL(hostUrl()).catch((err) => log('편집 창을 열지 못했습니다:', err?.message ?? err));
    return win;
  }

  const hostUrl = () => `${server.hostUrl}${testHooks ? '?test=1' : ''}`;

  /** 문서 경로 확인 → 절대 경로 */
  function editablePath(p) {
    if (typeof p !== 'string' || !p.trim() || p.length > 4096) throw fail('잘못된 파일 경로입니다', 'BAD_PATH');
    if (!isAbsolute(p)) throw fail('파일의 전체 경로가 필요합니다', 'BAD_PATH');
    p = resolve(p);
    let st;
    try { st = statSync(p); } catch { throw fail(`파일을 찾을 수 없습니다: ${basename(p)}`, 'NOT_FOUND'); }
    if (!st.isFile()) throw fail('파일이 아닙니다', 'BAD_PATH');
    if (!EDITABLE_EXTS.includes(extOf(p))) throw fail(`한글 문서(.hwp, .hwpx, .hwt)만 편집할 수 있습니다 (.${extOf(p) || '확장자 없음'})`, 'UNSUPPORTED');
    return p;
  }

  if (app) {
    app.on('before-quit', () => { quitting = true; });
    app.on('will-quit', () => { void server?.close(); });
  }

  return {
    /**
     * 문서를 편집 창으로 연다 (이미 열려 있으면 그 창을 앞으로) → { path, reused, available, reason }
     * 편집기 묶음이 없으면 창 안에 한국어 이유를 보여 준다 (available:false)
     */
    async openEditor(path) {
      const p = editablePath(path);
      const ok = available();
      const existing = entries.get(keyOf(p));
      if (existing?.win && !existing.win.isDestroyed()) {
        focus(existing);
        return { path: existing.path, reused: true, ...ok };
      }
      await ensureServer();
      const again = entries.get(keyOf(p)); // 서버를 기다리는 사이 같은 문서를 또 열었으면
      if (again?.win && !again.win.isDestroyed()) { focus(again); return { path: again.path, reused: true, ...ok }; }
      const entry = { path: p, key: keyOf(p), win: null, dirty: false, known: null, saving: false, allowClose: false, closing: null, watcher: null, watchTimer: null };
      entries.set(entry.key, entry);
      try { createWindow(entry); } catch (err) { entries.delete(entry.key); throw err; }
      startWatch(entry);
      return { path: p, reused: false, ...ok };
    },
    /** 그 문서가 편집 창에 열려 있는지 */
    isOpen(path) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return !!(e?.win && !e.win.isDestroyed());
    },
    /** 그 문서에 저장 안 한 변경이 있는지 (host가 알려 준 마지막 상태) */
    isDirty(path) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return !!(e?.win && !e.win.isDestroyed() && e.dirty);
    },
    /**
     * 다른 곳(AI·다른 프로그램)에서 파일을 바꿨을 때 알려 주기 → 편집 창에 알렸으면 true (내용이 같으면 false)
     * force: 내용이 같아도 알린다, reason: 'restore'(원래대로) 등 편집 창 안내 문구용
     */
    notifyExternalChange(path, { force = false, reason = null } = {}) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return e ? checkExternal(e, { force: force === true, reason: reason === 'restore' ? 'restore' : null }) : false;
    },
    /** 열린 편집 창 문서 경로들 */
    list() { return [...entries.values()].filter((e) => e.win && !e.win.isDestroyed()).map((e) => e.path); },
    /** 편집기 묶음 확인 → { available, reason } */
    available,
    /**
     * 모든 편집 창 닫기 → 다 닫혔으면 true (사용자가 '취소'하면 false)
     * force: 묻지 않고 닫는다 (저장 안 한 변경은 사라짐)
     */
    async closeAll({ force = false } = {}) {
      const list = [...entries.values()].filter((e) => e.win && !e.win.isDestroyed());
      if (force) {
        for (const e of list) { e.allowClose = true; e.win.destroy(); }
        return true;
      }
      for (const e of list) {
        if (!(await requestClose(e))) return false;
      }
      return true;
    },
    /** 서버 주소 (진단용) */
    serverInfo: () => (server ? { origin: server.origin, studioOrigin: server.studioOrigin } : null),
    /** 서버 닫기 (창을 다 닫은 뒤) */
    async dispose() { await server?.close(); server = null; },
  };
}
