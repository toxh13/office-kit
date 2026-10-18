// editor-window.mjs의 Office 종류(docx·xlsx·pptx) 시험 — 가짜 electron으로 창 관리자만 (서버·파일·IPC는 진짜)
//   글꼴 모듈(W5)·저장 뒤 고치기 모듈(W6)은 흉내 낸 것을 넣는다
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, utimesSync, writeFileSync } from 'node:fs';
import { request } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { createBlankDocument } from '../src/core/create.mjs';
import {
  ALL_EDITABLE_EXTS, checkOoxmlBytes, createEditorManager, draftPathFor, EDITABLE_EXTS, EDITOR_KINDS, editorKindOf,
  officeOpenProblem, OFFICE_DRAFTS_DIR, OFFICE_EXTS,
} from '../src/app/editor-window.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (b) => createHash('sha1').update(b).digest('hex');
async function until(fn, ms = 5000, step = 20) {
  const end = Date.now() + ms;
  for (;;) { const v = await fn(); if (v || Date.now() > end) return v; await sleep(step); }
}
const put = (p, data = 'x') => { mkdirSync(dirname(p), { recursive: true }); writeFileSync(p, data); };

function get(url) {
  return new Promise((ok, fail) => {
    const req = request(url, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => ok({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
    });
    req.on('error', fail);
    req.end();
  });
}

// ── 가짜 electron ────────────────────────────────────────

function makeFake() {
  const fake = {
    windows: [], dialogs: [], dialogAnswers: [], saveDialogs: [], saveDialogAnswers: [], opened: [], openPathResult: '',
    external: [], sessions: new Map(), replies: new Map(), commands: [], commandAnswer: null,
  };
  let wcSeq = 0;
  const app = Object.assign(new EventEmitter(), { isPackaged: false, getVersion: () => '0.2.0-test', quit() {} });
  class WebContents extends EventEmitter {
    constructor() { super(); this.id = ++wcSeq; this.sent = []; this.mainFrame = { url: '' }; }
    send(channel, data) {
      this.sent.push({ channel, data });
      // host 흉내: 닫기 전 질문·명령에 답한다
      if (channel === 'editor:request-close') setImmediate(() => emit('editor:reply', this, { id: data.id, value: fake.replies.get(this.id) ?? false }));
      if (channel === 'editor:command') {
        fake.commands.push(data.command);
        if (fake.commandAnswer) setImmediate(async () => emit('editor:reply', this, { id: data.id, value: await fake.commandAnswer(this, data.command) }));
      }
    }
    setWindowOpenHandler(fn) { this.openHandler = fn; }
    setWebRTCIPHandlingPolicy(p) { this.webrtcPolicy = p; }
    loadURL(u) { this.mainFrame.url = u; return Promise.resolve(); }
    toggleDevTools() {}
  }
  class BrowserWindow extends EventEmitter {
    constructor(opts = {}) {
      super();
      this.opts = opts; this.title = opts.title ?? ''; this.destroyed = false;
      this.webContents = new WebContents();
      fake.windows.push(this);
    }
    isDestroyed() { return this.destroyed; }
    close() {
      if (this.destroyed) return;
      let prevented = false;
      this.emit('close', { preventDefault: () => { prevented = true; } });
      if (!prevented) this.destroy();
    }
    destroy() { if (this.destroyed) return; this.destroyed = true; this.emit('closed'); }
    setTitle(t) { this.title = t; }
    getTitle() { return this.title; }
    show() {} focus() {} restore() {} removeMenu() {}
    isMinimized() { return false; }
  }
  const ipcMain = {
    handlers: new Map(), listeners: new Map(),
    handle(c, fn) { if (this.handlers.has(c)) throw new Error(`이미 등록된 채널: ${c}`); this.handlers.set(c, fn); },
    on(c, fn) { if (!this.listeners.has(c)) this.listeners.set(c, []); this.listeners.get(c).push(fn); },
  };
  const evOf = (wc, frame = wc.mainFrame) => ({ sender: wc, senderFrame: frame });
  function emit(channel, wc, ...args) { for (const fn of ipcMain.listeners.get(channel) ?? []) fn(evOf(wc), ...args); }
  async function invoke(wc, channel, ...args) {
    const fn = ipcMain.handlers.get(channel);
    if (!fn) throw new Error(`없는 채널: ${channel}`);
    return fn(evOf(wc), ...args);
  }
  const session = {
    fromPartition(name) {
      if (!fake.sessions.has(name)) {
        const s = Object.assign(new EventEmitter(), {
          name, handlers: {},
          webRequest: { onBeforeRequest(_f, fn) { s.handlers.beforeRequest = fn; } },
          setPermissionRequestHandler(fn) { s.handlers.permission = fn; },
          setPermissionCheckHandler(fn) { s.handlers.permissionCheck = fn; },
          setDevicePermissionHandler() {}, setDisplayMediaRequestHandler() {},
          setSpellCheckerEnabled(v) { s.spellcheck = v; },
          setSpellCheckerDictionaryDownloadURL(u) { s.dictionaryUrl = u; },
          async setProxy(cfg) { s.proxy = cfg; },
        });
        fake.sessions.set(name, s);
      }
      return fake.sessions.get(name);
    },
  };
  Object.assign(fake, {
    app, BrowserWindow, ipcMain, session, emit, invoke, evOf,
    dialog: {
      async showMessageBox(_w, opts) { fake.dialogs.push(opts); return { response: fake.dialogAnswers.shift() ?? 2 }; },
      async showSaveDialog(_w, opts) { fake.saveDialogs.push(opts); return fake.saveDialogAnswers.shift() ?? { canceled: true }; },
    },
    shell: {
      async openPath(p) { fake.opened.push(p); return fake.openPathResult; },
      async openExternal(u) { fake.external.push(u); },
    },
    nativeTheme: { shouldUseDarkColors: false },
  });
  return fake;
}

/** 시험용 묶음: 한글(studio·sdk) + Office(vendor/office-editor) + Office host 폴더 */
function makeBundles(dir) {
  put(join(dir, 'rhwp-studio', 'index.html'), '<!doctype html><title>studio</title>');
  put(join(dir, 'rhwp-studio', 'rhwp_bg.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(dir, 'rhwp-editor-sdk', f), 'export {};');
  const office = join(dir, 'office-editor');
  put(join(office, 'index.html'), '<!doctype html><title>app</title>');
  put(join(office, 'x2t', 'x2t.js'), 'var x;');
  put(join(office, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(office, 'editor', 'LICENSE.txt'), 'AGPL');
  put(join(office, 'editor', 'sdkjs', 'common', 'AllFonts.js'), '// 묶음 글꼴');
  put(join(office, 'VERSION.json'), JSON.stringify({ onlyoffice: { version: '9.4.0', build: '9.4.0.129' }, filex: { tag: 'v0.3.1' }, files: { a: 1 } }));
  const officeHost = join(dir, 'office-host');
  for (const f of ['host.html', 'host.js', 'host.css', 'office-preload.cjs']) put(join(officeHost, f), '//');
  return { studio: join(dir, 'rhwp-studio'), sdk: join(dir, 'rhwp-editor-sdk'), host: join(ROOT, 'src', 'app', 'editor'), office, officeHost };
}

/** OOXML 바이트에 부분 하나를 더한 사본 (고치기 흉내) */
async function withPart(bytes, name, text) {
  const z = await JSZip.loadAsync(bytes);
  z.file(name, text);
  return new Uint8Array(await z.generateAsync({ type: 'uint8array' }));
}

test('종류 나누기: 확장자로 hwp / office', () => {
  assert.deepEqual(EDITABLE_EXTS, ['hwp', 'hwpx', 'hwt'], 'main.mjs가 쓰는 이름은 한글만 (예전 그대로)');
  assert.deepEqual(OFFICE_EXTS, ['docx', 'xlsx', 'pptx']);
  assert.deepEqual(ALL_EDITABLE_EXTS, ['hwp', 'hwpx', 'hwt', 'docx', 'xlsx', 'pptx']);
  assert.deepEqual(EDITOR_KINDS, { hwp: EDITABLE_EXTS, office: OFFICE_EXTS });
  for (const p of ['a.hwp', 'C:\\x\\b.HWPX', 'c.hwt', 'hwpx']) assert.equal(editorKindOf(p), 'hwp', p);
  for (const p of ['a.docx', 'C:\\x\\b.XLSX', 'c.pptx', 'pptx']) assert.equal(editorKindOf(p), 'office', p);
  for (const p of ['a.doc', 'b.xlsm', 'c.pdf', 'd', '', null]) assert.equal(editorKindOf(p), null, String(p));
});

test('Office 편집 창 (가짜 electron)', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-office-win-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  mkdirSync(docs, { recursive: true });
  const dirs = makeBundles(join(base, 'vendor'));
  const docx = join(docs, '메모.docx');
  const xlsx = join(docs, '표.xlsx');
  const pptx = join(docs, '발표.pptx');
  const hwpx = join(docs, '보고서.hwpx');
  for (const p of [docx, xlsx, pptx, hwpx]) await createBlankDocument(p);

  const fake = makeFake();
  const logs = [];
  const fontsCalls = [];
  const fontsDir = join(base, 'fonts');
  put(join(fontsDir, 'AllFonts.js'), '// 이 PC 글꼴');
  const fixCalls = []; const fixOriginals = [];
  let fixMode = 'add';
  const saved = []; const downloaded = []; const dirtyChanges = []; const backups = [];
  const stateStore = new Map();
  const mgr = createEditorManager({
    app: fake.app, BrowserWindow: fake.BrowserWindow, ipcMain: fake.ipcMain, session: fake.session, dialog: fake.dialog,
    shell: fake.shell, nativeTheme: fake.nativeTheme, appRoot: base, userData: ud, dirs,
    getBackups: () => ({ backup: (p) => { backups.push(p); return { file: `${p}.bak` }; } }),
    onSaved: (p) => saved.push(p), onDownloaded: (p) => downloaded.push(p), onDirtyChange: (p, d) => dirtyChanges.push([p, d]),
    officeState: { get: (k) => stateStore.get(k), set: (k, v) => { stateStore.set(k, v); } },
    fonts: {
      async ensureFonts(o) {
        fontsCalls.push(o);
        return { allFontsJsPath: join(fontsDir, 'AllFonts.js'), thumbnailsDir: join(fontsDir, 'thumbs'), fontFileByIndex: () => null };
      },
    },
    fixups: {
      async applyFixups(ext, bytes, o) {
        fixCalls.push(ext);
        fixOriginals.push(o?.original ? sha(o.original) : null);
        if (fixMode === 'throw') throw new Error('고치기 실패 흉내');
        if (fixMode === 'broken') return { bytes: new Uint8Array([1, 2, 3]), applied: ['깨뜨림'] };
        if (fixMode === 'none') return { bytes, applied: [] };
        return { bytes: await withPart(bytes, 'okfix.txt', ext), applied: [`${ext} 고침`] };
      },
    },
    log: (...a) => logs.push(a.join(' ')),
    userName: '시험 사용자',
    limits: { printIntervalMs: 300 },
  });
  t.after(async () => {
    await mgr.closeAll({ force: true });
    fake.app.emit('will-quit');
    await mgr.dispose();
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  });

  const win = (i) => fake.windows[i];
  const call = (w, channel, ...args) => fake.invoke(w.webContents, channel, ...args);

  let dw;
  await t.test('열기: 제목·partition·preload·host 주소·WebRTC 정책', async () => {
    assert.deepEqual(mgr.available('office'), { available: true, reason: null });
    assert.equal(mgr.available().available, true, '기본은 한글 묶음 확인 (main.mjs 그대로)');
    const r = await mgr.openEditor(docx);
    assert.deepEqual({ ...r }, { path: docx, kind: 'office', reused: false, available: true, reason: null, notice: null });
    dw = win(0);
    assert.equal(dw.getTitle(), '메모.docx - 문서 편집 (Office Kit)');
    assert.equal(dw.opts.webPreferences.partition, 'officekit-office');
    assert.equal(dw.opts.webPreferences.preload, join(dirs.officeHost, 'office-preload.cjs'));
    assert.equal(dw.opts.webPreferences.sandbox, true);
    assert.equal(dw.opts.webPreferences.contextIsolation, true);
    assert.equal(dw.opts.webPreferences.nodeIntegrationInSubFrames, false);
    assert.match(dw.webContents.mainFrame.url, /^http:\/\/127\.0\.0\.1:\d+\/office-host\/host\.html$/);
    assert.equal(dw.webContents.webrtcPolicy, 'disable_non_proxied_udp');
    assert.doesNotMatch(dw.getTitle(), /ONLYOFFICE/i);
    await mgr.openEditor(xlsx);
    await mgr.openEditor(pptx);
    assert.equal(win(1).getTitle(), '표.xlsx - 표 편집 (Office Kit)');
    assert.equal(win(2).getTitle(), '발표.pptx - 슬라이드 편집 (Office Kit)');
    const again = await mgr.openEditor(docx);
    assert.equal(again.reused, true);
    assert.equal(fake.windows.length, 3);
    assert.equal(mgr.isOpen(docx), true);
    assert.equal(mgr.kindOf(docx), 'office');
  });

  await t.test('세션: officekit-office — 죽은 프록시(127.0.0.3만 바로), 맞춤법 사전 주소 죽임, 요청 막기, 권한', async () => {
    const s = fake.sessions.get('officekit-office');
    const info = mgr.serverInfo();
    const officeHost = new URL(info.officeOrigin).hostname;
    // 바로 가는 것은 우리 서버 두 곳의 주소:포트뿐 (<-loopback>: 이 PC의 다른 포트는 죽은 프록시로 → TURN-TCP로 이 PC 서비스에 못 닿음)
    assert.deepEqual(s.proxy, { mode: 'fixed_servers', proxyRules: 'socks5://127.0.0.1:9', proxyBypassRules: `<-loopback>;${new URL(info.origin).host};${new URL(info.officeOrigin).host}` });
    if (process.platform === 'win32') assert.equal(officeHost, '127.0.0.3');
    assert.equal(s.dictionaryUrl, 'http://127.0.0.1:9/');
    assert.equal(s.spellcheck, false);
    const allowed = (url) => new Promise((r) => s.handlers.beforeRequest({ url }, (x) => r(!x.cancel)));
    const pkg = `${info.officeOrigin}/office/${info.officeKey}/`;
    for (const url of [`${info.origin}/office-host/host.html`, `${pkg}index.html`, `${pkg}x2t/x2t.wasm`, 'blob:null/abc', 'data:text/plain,a', 'about:blank']) {
      assert.equal(await allowed(url), true, url);
    }
    for (const url of ['https://example.com/', 'http://example.com/', `${info.origin}/host/host.html`, `${info.origin}/sdk/index.js`,
      `${info.studioOrigin}/studio/index.html`, `${info.officeOrigin}/office/other/index.html`, `ws://${new URL(info.officeOrigin).host}/`,
      'https://www.onlyoffice.com/', 'http://127.0.0.1:9/', 'file:///C:/Windows/win.ini']) {
      assert.equal(await allowed(url), false, url);
    }
    // 권한: 클립보드 쓰기만, 이 관리자의 Office 창에서만
    const perm = (wc, p) => new Promise((r) => s.handlers.permission(wc, p, r, {}));
    assert.equal(await perm(dw.webContents, 'clipboard-sanitized-write'), true);
    for (const p of ['local-fonts', 'media', 'geolocation', 'notifications', 'display-capture', 'clipboard-read', 'openExternal']) {
      assert.equal(await perm(dw.webContents, p), false, p);
    }
    assert.equal(await perm({ id: 9999 }, 'clipboard-sanitized-write'), false);
    assert.equal(s.handlers.permissionCheck(dw.webContents, 'clipboard-sanitized-write'), true);
    assert.equal(s.handlers.permissionCheck(dw.webContents, 'local-fonts'), false);
  });

  await t.test('한글 창과 같이: 세션·주소가 따로 (한글 세션은 예전 그대로)', async () => {
    const r = await mgr.openEditor(hwpx);
    assert.equal(r.kind, 'hwp');
    const hw = fake.windows.at(-1);
    assert.equal(hw.getTitle(), '보고서.hwpx - 한글 편집 (Office Kit)');
    assert.equal(hw.opts.webPreferences.partition, 'officekit-editor');
    assert.match(hw.opts.webPreferences.preload, /editor[\\/]editor-preload\.cjs$/);
    assert.match(hw.webContents.mainFrame.url, /\/host\/host\.html$/);
    const s = fake.sessions.get('officekit-editor');
    const info = mgr.serverInfo();
    assert.deepEqual(s.proxy, { mode: 'fixed_servers', proxyRules: 'socks5://127.0.0.1:9', proxyBypassRules: `<-loopback>;${new URL(info.origin).host};${new URL(info.studioOrigin).host}` });
    const allowed = (url) => new Promise((res) => s.handlers.beforeRequest({ url }, (x) => res(!x.cancel)));
    assert.equal(await allowed(`${info.studioOrigin}/studio/index.html`), true);
    assert.equal(await allowed(`${info.officeOrigin}/office/${info.officeKey}/index.html`), false, '한글 창은 Office 출처에 닿지 않는다');
    const perm = (wc, p) => new Promise((res) => s.handlers.permission(wc, p, res, {}));
    assert.equal(await perm(hw.webContents, 'local-fonts'), true, '한글 창 권한은 예전 그대로');
    assert.equal(await perm(dw.webContents, 'local-fonts'), false, 'Office 창은 한글 세션 권한을 받지 못한다');
    // 한글 창은 Office 전용 IPC를 못 쓴다
    const d = await call(hw, 'editor:download', 'a.pdf', new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), 'application/pdf');
    assert.equal(d.ok, false);
    assert.match(d.error, /허용되지 않은 요청/);
    const doc = await call(hw, 'editor:get-document');
    assert.equal(doc.ok, true, doc.error);
    assert.equal(doc.value.kind, 'hwp');
    assert.ok(doc.value.studioUrl);
    assert.equal(doc.value.appUrl, undefined);
    hw.destroy();
  });

  await t.test('보낸 쪽 확인: host 최상위 틀 + /office-host/ 경로만', async () => {
    const wc = dw.webContents;
    const orig = wc.mainFrame.url;
    const fn = fake.ipcMain.handlers.get('editor:get-document');
    // 하위 틀에서 온 요청
    let r = await fn(fake.evOf(wc, { url: orig }));
    assert.equal(r.ok, false);
    // 같은 출처라도 한글 host 경로면 거절
    wc.mainFrame.url = orig.replace('/office-host/', '/host/');
    r = await fn(fake.evOf(wc));
    assert.equal(r.ok, false);
    wc.mainFrame.url = orig;
    r = await fn(fake.evOf(wc));
    assert.equal(r.ok, true, r.error);
  });

  await t.test('editor:get-document: 앱 주소·locale·about, 글꼴 목록을 먼저 준비 (W5 ensureFonts)', async () => {
    const r = await call(dw, 'editor:get-document');
    assert.equal(r.ok, true, r.error);
    const v = r.value;
    const info = mgr.serverInfo();
    assert.equal(v.kind, 'office');
    assert.equal(v.name, '메모.docx');
    assert.equal(v.ext, 'docx');
    assert.equal(v.appUrl, `${info.officeOrigin}/office/${info.officeKey}/index.html`);
    assert.equal(v.locale, 'ko-KR');
    assert.equal(v.mime, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    assert.equal(v.title, '문서 편집');
    assert.equal(v.theme, 'light');
    assert.equal(v.readOnly, false);
    assert.equal(v.available, true);
    assert.equal(v.draft, null);
    assert.equal(v.userName, '시험 사용자', '변경 내용 추적·메모 작성자 (\'나\'가 아님)');
    assert.equal(v.fontsWarning, null);
    assert.ok(v.bytes instanceof Uint8Array);
    assert.equal(sha(v.bytes), sha(readFileSync(docx)));
    assert.equal(v.about.appVersion, '0.2.0-test');
    assert.deepEqual(v.about.bundle, { onlyoffice: { version: '9.4.0', build: '9.4.0.129' }, filex: { tag: 'v0.3.1' } }, '파일 목록은 빼고');
    assert.equal(fontsCalls.length, 1, '글꼴 준비는 한 번');
    assert.deepEqual(fontsCalls[0], { userData: ud, editorVersion: info.officeVersion });
    // 서버가 이 PC 글꼴 목록을 내준다
    const all = await get(`${info.officeOrigin}/office/${info.officeKey}/editor/sdkjs/common/AllFonts.js`);
    assert.match(all.body.toString(), /이 PC 글꼴/);
    const x = await call(win(1), 'editor:get-document');
    assert.equal(x.value.title, '표 편집');
    assert.equal(fontsCalls.length, 1);
  });

  await t.test('저장: 형식 확인 → 고치기(W6) → 백업 → 쓰기 → onSaved', async () => {
    const before = readFileSync(docx);
    // 형식이 아닌 바이트는 저장하지 않는다
    for (const bad of [new Uint8Array([1, 2, 3]), new Uint8Array(readFileSync(hwpx)), new Uint8Array(readFileSync(xlsx))]) {
      const r = await call(dw, 'editor:save', bad, { reason: 'user' });
      assert.equal(r.ok, false);
      assert.equal(r.code, 'BAD_FORMAT');
      assert.match(r.error, /Word 문서\(\.docx\)/);
    }
    assert.equal(sha(readFileSync(docx)), sha(before));
    assert.equal(fixCalls.length, 0);
    const edited = await withPart(readFileSync(docx), 'edited.txt', '편집');
    const r = await call(dw, 'editor:save', edited, { reason: 'user' });
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(fixCalls, ['docx']);
    assert.deepEqual(fixOriginals, [sha(before)], '고치기에 덮어쓰기 전 원래 파일을 넘긴다 (원본에서 옮기기)');
    assert.deepEqual(r.value.fixups, ['docx 고침']);
    assert.equal(r.value.warning, null);
    assert.equal(r.value.backup, `${docx}.bak`);
    assert.deepEqual(backups, [docx]);
    assert.deepEqual(saved, [docx]);
    const z = await JSZip.loadAsync(readFileSync(docx));
    assert.equal(await z.file('okfix.txt')?.async('string'), 'docx', '고친 바이트를 저장');
    assert.equal(await z.file('edited.txt')?.async('string'), '편집');
    // reason이 없어도 사용자 저장
    fixMode = 'none';
    const r2 = await call(dw, 'editor:save', edited);
    assert.equal(r2.ok, true, r2.error);
    assert.equal(sha(readFileSync(docx)), sha(edited));
    // 바깥 변경으로 보지 않는다 (자기 저장)
    assert.equal(mgr.notifyExternalChange(docx), false);
  });

  await t.test('고치기가 실패하면 원래 바이트로 저장, pptx만 알림 띠(warning)', async () => {
    fixMode = 'throw';
    const d = await withPart(readFileSync(docx), 'v3.txt', '3');
    let r = await call(dw, 'editor:save', d);
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.warning, null);
    assert.equal(sha(readFileSync(docx)), sha(d));
    const pw = win(2);
    assert.equal((await call(pw, 'editor:get-document')).ok, true);
    const p = await withPart(readFileSync(pptx), 'v.txt', 'p');
    r = await call(pw, 'editor:save', p);
    assert.equal(r.ok, true, r.error);
    assert.match(r.value.warning, /PowerPoint/);
    assert.equal(sha(readFileSync(pptx)), sha(p));
    // 고친 결과가 형식에 맞지 않으면 원래 바이트
    fixMode = 'broken';
    const p2 = await withPart(readFileSync(pptx), 'v2.txt', 'p');
    r = await call(pw, 'editor:save', p2);
    assert.equal(r.ok, true, r.error);
    assert.equal(sha(readFileSync(pptx)), sha(p2));
    assert.match(r.value.warning, /PowerPoint/);
    assert.ok(logs.some((l) => /고친 결과가 형식에 맞지 않습니다/.test(l)));
    fixMode = 'add';
  });

  await t.test('fixOoxml 모양(W6 실제 모듈 이름)도 받는다: { bytes, fixes, error, notify }', async () => {
    const base2 = mkdtempSync(join(tmpdir(), 'ok-office-fix-'));
    const fake2 = makeFake();
    const doc = join(base2, '발표.pptx');
    await createBlankDocument(doc);
    const m2 = createEditorManager({
      app: fake2.app, BrowserWindow: fake2.BrowserWindow, ipcMain: fake2.ipcMain, session: fake2.session, dialog: fake2.dialog,
      shell: fake2.shell, appRoot: base2, userData: join(base2, 'ud'), dirs, log: () => {},
      fonts: { ensureFonts: async () => { throw new Error('글꼴 실패 흉내'); } },
      fixups: { async fixOoxml(bytes, ext) { return { bytes, changed: false, fixes: [], error: `${ext} 고치기 실패 (흉내)`, notify: true }; } },
    });
    try {
      await m2.openEditor(doc);
      const w = fake2.windows[0];
      const g = await fake2.invoke(w.webContents, 'editor:get-document');
      assert.equal(g.ok, true, g.error);
      assert.match(g.value.fontsWarning, /글꼴 목록을 만들지 못해/, '글꼴을 못 만들어도 연다');
      const r = await fake2.invoke(w.webContents, 'editor:save', new Uint8Array(readFileSync(doc)));
      assert.equal(r.ok, true, r.error);
      assert.equal(r.value.warning, 'pptx 고치기 실패 (흉내)');
    } finally {
      await m2.closeAll({ force: true }); // 창 감시(폴더)를 먼저 끈다
      await m2.dispose();
      rmSync(base2, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });

  await t.test('글꼴 준비가 너무 늦으면 기다리지 않고 묶음 글꼴로 연다 (limits.fontsTimeoutMs)', async () => {
    const base3 = mkdtempSync(join(tmpdir(), 'ok-office-slow-'));
    const fake3 = makeFake();
    const doc = join(base3, '느림.docx');
    await createBlankDocument(doc);
    const lines = [];
    const m3 = createEditorManager({
      app: fake3.app, BrowserWindow: fake3.BrowserWindow, ipcMain: fake3.ipcMain, session: fake3.session, dialog: fake3.dialog,
      shell: fake3.shell, appRoot: base3, userData: join(base3, 'ud'), dirs, log: (...a) => lines.push(a.join(' ')),
      limits: { fontsTimeoutMs: 150 },
      fonts: { ensureFonts: () => new Promise(() => {}) }, // 끝나지 않는 글꼴 준비
    });
    try {
      await m3.openEditor(doc);
      const t0 = Date.now();
      const g = await fake3.invoke(fake3.windows[0].webContents, 'editor:get-document');
      assert.equal(g.ok, true, g.error);
      assert.ok(Date.now() - t0 < 3000);
      assert.ok(lines.some((l) => /글꼴 목록 준비가 늦어/.test(l)));
      // 설정 값: 함수처럼 JSON으로 못 바꾸는 값은 거절 (메모리 보관)
      const w = fake3.windows[0].webContents;
      assert.equal((await fake3.invoke(w, 'editor:state', 'set', 'k', 1)).ok, true);
      assert.equal((await fake3.invoke(w, 'editor:state', 'get', 'k')).value, 1);
      assert.equal((await fake3.invoke(w, 'editor:state', 'set', 'k', null)).ok, true);
      assert.equal((await fake3.invoke(w, 'editor:state', 'get', 'k')).value, null);
      assert.equal((await fake3.invoke(w, 'editor:state', 'set', 'f', () => 1)).code, 'BAD_VALUE');
    } finally {
      await m3.closeAll({ force: true });
      await m3.dispose();
      rmSync(base3, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });

  await t.test('자동 저장(reason:auto): 원본은 그대로, userData/office-drafts/<sha1>.<ext> 사본 → 다음 열기에 복구 제안', async () => {
    const before = sha(readFileSync(docx));
    const n0 = saved.length; const b0 = backups.length; const f0 = fixCalls.length;
    const draftBytes = await withPart(readFileSync(docx), 'draft.txt', '자동 저장');
    const r = await call(dw, 'editor:save', draftBytes, { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.draft, true);
    const file = draftPathFor(ud, docx);
    assert.equal(file, join(ud, OFFICE_DRAFTS_DIR, `${sha(process.platform === 'win32' ? docx.toLowerCase() : docx)}.docx`));
    assert.equal(mgr.draftPath(docx), file);
    assert.equal(sha(readFileSync(file)), sha(draftBytes));
    assert.equal(sha(readFileSync(docx)), before, '사용자 파일을 몰래 덮어쓰지 않는다');
    assert.equal(saved.length, n0, 'onSaved 없음');
    assert.equal(backups.length, b0, '백업 없음');
    assert.equal(fixCalls.length, f0, '사본에는 고치기 없음');
    // 형식이 아니면 사본도 쓰지 않는다
    assert.equal((await call(dw, 'editor:save', new Uint8Array([9]), { reason: 'auto' })).ok, false);
    // 사본이 파일보다 새로우면 제안
    const old = new Date(Date.now() - 60_000);
    utimesSync(docx, old, old);
    let g = await call(dw, 'editor:get-document');
    assert.equal(g.value.draft.size, draftBytes.length);
    assert.ok(g.value.draft.mtimeMs > g.value.mtimeMs);
    assert.equal(g.value.draft.bytes, undefined, '바이트는 editor:draft read로');
    const rd = await call(dw, 'editor:draft', 'read');
    assert.equal(rd.ok, true, rd.error);
    assert.equal(sha(rd.value.bytes), sha(draftBytes));
    // 파일이 더 새로우면 제안하지 않는다
    const future = new Date(Date.now() + 60_000);
    utimesSync(docx, future, future);
    g = await call(dw, 'editor:get-document');
    assert.equal(g.value.draft, null);
    // 버리기
    const dd = await call(dw, 'editor:draft', 'discard');
    assert.deepEqual(dd.value, { discarded: true });
    assert.equal(existsSync(file), false);
    const none = await call(dw, 'editor:draft', 'read');
    assert.equal(none.ok, false);
    assert.equal(none.code, 'NO_DRAFT');
    assert.equal((await call(dw, 'editor:draft', 'write')).ok, false);
    // 정상 저장하면 사본을 지운다
    assert.equal((await call(dw, 'editor:save', draftBytes, { reason: 'auto' })).ok, true);
    assert.equal(existsSync(file), true);
    assert.equal((await call(dw, 'editor:save', draftBytes, { reason: 'user' })).ok, true);
    assert.equal(existsSync(file), false);
  });

  await t.test('복구 / 버리기에 답하지 않고 닫으면 복구 사본을 남기고 다음에 다시 묻는다', async () => {
    const doc = join(docs, '복구.docx');
    await createBlankDocument(doc);
    await mgr.openEditor(doc);
    let w = fake.windows.at(-1);
    const draftBytes = await withPart(readFileSync(doc), 'draft.txt', '자동 저장');
    assert.equal((await call(w, 'editor:save', draftBytes, { reason: 'auto' })).ok, true);
    const file = draftPathFor(ud, doc);
    // 앱이 죽은 뒤 다시 연 것처럼: 파일이 사본보다 오래됨 → 띠
    w.destroy();
    const old = new Date(Date.now() - 60_000);
    utimesSync(doc, old, old);
    await mgr.openEditor(doc);
    w = fake.windows.at(-1);
    let g = await call(w, 'editor:get-document');
    assert.equal(g.value.draft?.size, draftBytes.length);
    // 답하지 않고 Ctrl+W (편집한 것 없음 → 묻지 않고 닫힘)
    w.close();
    assert.ok(await until(() => w.isDestroyed()));
    assert.equal(existsSync(file), true, '유일한 사본을 지우지 않는다');
    await mgr.openEditor(doc);
    w = fake.windows.at(-1);
    g = await call(w, 'editor:get-document');
    assert.equal(g.value.draft?.size, draftBytes.length, '다시 연 창에서 다시 묻는다');
    // 사용자 저장도 답하기 전에는 사본을 지우지 않는다
    fixMode = 'none';
    assert.equal((await call(w, 'editor:save', new Uint8Array(readFileSync(doc)), { reason: 'user' })).ok, true);
    assert.equal(existsSync(file), true);
    // 복구하면 닫을 때 지운다 (편집 창이 그 내용을 들고 있으므로 저장 여부는 닫기 대화상자가 묻는다)
    utimesSync(doc, old, old);
    g = await call(w, 'editor:get-document');
    assert.ok(g.value.draft);
    assert.equal((await call(w, 'editor:draft', 'read')).ok, true);
    w.close();
    assert.ok(await until(() => w.isDestroyed()));
    assert.equal(existsSync(file), false);
  });

  await t.test('다른 이름으로 저장: 같은 형식만, 다른 형식은 내려받기로', async () => {
    const bytes = new Uint8Array(readFileSync(docx));
    let r = await call(dw, 'editor:save-as', bytes, 'pdf');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'UNSUPPORTED');
    assert.match(r.error, /다른 형식으로 내려받기/);
    r = await call(dw, 'editor:save-as', new Uint8Array(readFileSync(xlsx)), 'xlsx');
    assert.equal(r.ok, false);
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '메모 사본') });
    r = await call(dw, 'editor:save-as', bytes, 'docx');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '메모 사본.docx');
    assert.deepEqual(fake.saveDialogs.at(-1).filters, [{ name: 'Word 문서 (.docx)', extensions: ['docx'] }]);
    assert.equal(dw.getTitle(), '메모 사본.docx - 문서 편집 (Office Kit)');
    assert.equal(mgr.isOpen(join(docs, '메모 사본.docx')), true);
    assert.equal(mgr.isOpen(docx), false);
    // 취소
    fake.saveDialogAnswers.push({ canceled: true });
    r = await call(dw, 'editor:save-as', bytes, 'docx');
    assert.deepEqual(r.value, { canceled: true });
  });

  await t.test('editor:download: 저장 대화상자 → 파일, 형식 허용 목록·바이트 확인·열린 문서 거절', async () => {
    const pdf = new Uint8Array(Buffer.from('%PDF-1.7\n가나다'));
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '내보낸.pdf') });
    let r = await call(dw, 'editor:download', '메모.pdf', pdf, 'application/pdf');
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(r.value, { saved: true, name: '내보낸.pdf', ext: 'pdf', size: pdf.length, path: join(docs, '내보낸.pdf') });
    assert.equal(readFileSync(join(docs, '내보낸.pdf'), 'utf8'), '%PDF-1.7\n가나다');
    const sd = fake.saveDialogs.at(-1);
    assert.equal(sd.defaultPath, join(docs, '메모.pdf'));
    assert.deepEqual(sd.filters, [{ name: 'PDF 문서 (.pdf)', extensions: ['pdf'] }]);
    assert.deepEqual(downloaded, [join(docs, '내보낸.pdf')]);
    // 이름의 경로 구분자는 지운다
    fake.saveDialogAnswers.push({ canceled: true });
    r = await call(dw, 'editor:download', '..\\..\\evil/x.pdf', pdf, 'application/pdf');
    assert.deepEqual(r.value, { canceled: true });
    assert.equal(fake.saveDialogs.at(-1).defaultPath, join(docs, '.._.._evil_x.pdf'));
    // 받지 않는 형식, 형식과 다른 바이트
    for (const [name, bytes] of [['a.exe', pdf], ['a.bat', pdf], ['a.docm', pdf], ['a.lnk', pdf]]) {
      r = await call(dw, 'editor:download', name, bytes, '');
      assert.equal(r.ok, false, name);
      assert.equal(r.code, 'UNSUPPORTED', name);
    }
    for (const [name, bytes] of [['a.pdf', new Uint8Array([1, 2])], ['a.odt', pdf], ['a.docx', pdf], ['a.png', pdf]]) {
      r = await call(dw, 'editor:download', name, bytes, '');
      assert.equal(r.ok, false, name);
      assert.equal(r.code, 'BAD_FORMAT', name);
    }
    assert.equal((await call(dw, 'editor:download', 'a.pdf', new Uint8Array(0), '')).ok, false);
    // 편집 창에 열린 문서로는 내려받지 않는다
    fake.saveDialogAnswers.push({ canceled: false, filePath: xlsx });
    r = await call(dw, 'editor:download', '표.xlsx', new Uint8Array(readFileSync(xlsx)), '');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'IN_USE');
    // 텍스트 형식은 그대로, 확장자가 없으면 붙인다
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '글') });
    r = await call(dw, 'editor:download', '메모.txt', new Uint8Array(Buffer.from('글')), 'text/plain');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '글.txt');
  });

  await t.test('editor:print: 임시 PDF → 기본 PDF 앱, 없으면 저장 대화상자', async () => {
    const pdf = new Uint8Array(Buffer.from('%PDF-1.4\n인쇄'));
    // 창마다 처음 한 번은 묻는다 (취소 → 열지 않음)
    let d0 = fake.dialogs.length; const o0 = fake.opened.length;
    fake.dialogAnswers.push(1);
    let r = await call(dw, 'editor:print', pdf);
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(r.value, { canceled: true });
    assert.equal(fake.dialogs[d0].message, '인쇄용 PDF를 기본 PDF 앱으로 열까요?');
    assert.equal(fake.opened.length, o0);
    await sleep(350);
    fake.dialogAnswers.push(0);
    r = await call(dw, 'editor:print', pdf);
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.opened, true);
    // 바로 다시 부르면 거절 (편집기 틀이 PDF 앱을 연달아 띄우지 못하게), 간격이 지나면 묻지 않고 연다
    const busy = await call(dw, 'editor:print', pdf);
    assert.equal(busy.ok, false);
    assert.equal(busy.code, 'BUSY');
    await sleep(350);
    d0 = fake.dialogs.length;
    r = await call(dw, 'editor:print', pdf);
    assert.equal(r.value.opened, true);
    assert.equal(fake.dialogs.length, d0, '두 번째부터는 묻지 않는다');
    rmSync(dirname(fake.opened.at(-2)), { recursive: true, force: true });
    const file = fake.opened.at(-1);
    assert.equal(r.value.path, file);
    assert.match(file, /officekit-print[\\/][0-9a-f]{12}[\\/]메모 사본\.pdf$/);
    assert.equal(readFileSync(file, 'utf8'), '%PDF-1.4\n인쇄');
    // PDF가 아니면 거절
    r = await call(dw, 'editor:print', new Uint8Array([1, 2, 3, 4, 5]));
    assert.equal(r.ok, false);
    assert.equal(r.code, 'BAD_FORMAT');
    // PDF 앱이 없으면 저장
    fake.openPathResult = '연결된 앱 없음';
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '인쇄본.pdf') });
    await sleep(350);
    r = await call(dw, 'editor:print', pdf);
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.opened, false);
    assert.equal(r.value.saved, true);
    assert.equal(readFileSync(join(docs, '인쇄본.pdf'), 'utf8'), '%PDF-1.4\n인쇄');
    fake.openPathResult = '';
    rmSync(dirname(file), { recursive: true, force: true });
  });

  await t.test('editor:state: 설정 get/set, 값마다 8KB, 이름 확인', async () => {
    let r = await call(dw, 'editor:state', 'set', 'de-settings.zoom', { zoom: 120, list: [1, 2] });
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(stateStore.get('de-settings.zoom'), { zoom: 120, list: [1, 2] });
    r = await call(dw, 'editor:state', 'get', 'de-settings.zoom');
    assert.deepEqual(r.value, { zoom: 120, list: [1, 2] });
    assert.equal((await call(dw, 'editor:state', 'get', 'none')).value, null);
    r = await call(dw, 'editor:state', 'set', 'big', 'x'.repeat(9000));
    assert.equal(r.ok, false);
    assert.equal(r.code, 'TOO_BIG');
    for (const key of ['', 'a b', '../x', 'x'.repeat(200), 5]) {
      assert.equal((await call(dw, 'editor:state', 'get', key)).ok, false, String(key));
    }
    assert.equal((await call(dw, 'editor:state', 'delete', 'x')).ok, false);
  });

  await t.test('dirty → 제목 *·onDirtyChange, Ctrl+S → host에 저장 명령', async () => {
    const p = mgr.list().find((x) => x.endsWith('.docx'));
    fake.emit('editor:dirty', dw.webContents, true);
    assert.equal(dw.getTitle(), '* 메모 사본.docx - 문서 편집 (Office Kit)');
    assert.equal(mgr.isDirty(p), true);
    assert.deepEqual(dirtyChanges.at(-1), [p, true]);
    let prevented = false;
    dw.webContents.emit('before-input-event', { preventDefault: () => { prevented = true; } }, { type: 'keyDown', code: 'KeyS', control: true });
    assert.equal(prevented, true);
    assert.equal(fake.commands.at(-1), 'save');
    fake.emit('editor:dirty', dw.webContents, false);
    assert.equal(dw.getTitle(), '메모 사본.docx - 문서 편집 (Office Kit)');
  });

  await t.test('틀 이동: Office 묶음 주소만 (studio·바깥은 막음)', async () => {
    const info = mgr.serverInfo();
    const nav = (url, isMainFrame = false) => { let p = false; dw.webContents.emit('will-frame-navigate', { url, isMainFrame, preventDefault: () => { p = true; } }); return !p; };
    assert.equal(nav(`${info.officeOrigin}/office/${info.officeKey}/editor/web-apps/apps/documenteditor/main/index.html`), true);
    assert.equal(nav('about:blank'), true);
    assert.equal(nav(`${info.studioOrigin}/studio/index.html`), false);
    assert.equal(nav(`${info.officeOrigin}/office/x/index.html`), false);
    assert.equal(nav('https://example.com/'), false);
    assert.equal(nav(`${info.origin}/office-host/host.html`, true), false, '최상위 이동은 늘 막음');
    // 새 창: 내부 주소는 열지 않는다
    assert.deepEqual(dw.webContents.openHandler({ url: `${info.officeOrigin}/x` }), { action: 'deny' });
    await sleep(20);
    assert.deepEqual(fake.external, []);
  });

  await t.test('열기 전 거르기: 암호 문서·옛 형식·손상·너무 큼 → 한국어 오류, 창을 열지 않음', async () => {
    const n = fake.windows.length;
    const enc = join(docs, '암호.docx');
    const cfb = Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), Buffer.alloc(500), Buffer.from('EncryptionInfo', 'utf16le'), Buffer.alloc(100)]);
    writeFileSync(enc, cfb);
    await assert.rejects(mgr.openEditor(enc), (e) => e.code === 'ENCRYPTED' && e.message === '암호가 걸린 문서는 직접 편집할 수 없습니다. 프로그램에서 여세요.');
    const enc2 = join(docs, '암호2.pptx');
    writeFileSync(enc2, Buffer.concat([cfb.subarray(0, 8), Buffer.alloc(64), Buffer.from('EncryptedPackage', 'utf16le')]));
    await assert.rejects(mgr.openEditor(enc2), { code: 'ENCRYPTED' });
    const legacy = join(docs, '옛날.xlsx');
    writeFileSync(legacy, Buffer.concat([cfb.subarray(0, 8), Buffer.alloc(64), Buffer.from('Workbook', 'utf16le')]));
    await assert.rejects(mgr.openEditor(legacy), (e) => e.code === 'OLD_FORMAT' && /프로그램에서 여세요/.test(e.message));
    const notzip = join(docs, '글.docx');
    writeFileSync(notzip, '그냥 글');
    await assert.rejects(mgr.openEditor(notzip), (e) => e.code === 'BAD_FORMAT' && /Word 문서\(\.docx\) 형식이 아닙니다/.test(e.message));
    const hwpxAsDocx = join(docs, '한글.docx');
    writeFileSync(hwpxAsDocx, readFileSync(hwpx));
    await assert.rejects(mgr.openEditor(hwpxAsDocx), (e) => e.code === 'BAD_FORMAT' && /Content_Types/.test(e.message));
    const xlsxAsDocx = join(docs, '표인데.docx');
    writeFileSync(xlsxAsDocx, readFileSync(xlsx));
    await assert.rejects(mgr.openEditor(xlsxAsDocx), (e) => e.code === 'BAD_FORMAT' && /word\/document\.xml/.test(e.message));
    const empty = join(docs, '빈.pptx');
    writeFileSync(empty, '');
    await assert.rejects(mgr.openEditor(empty), { code: 'BAD_FORMAT' });
    assert.equal(officeOpenProblem(docx.replace('메모', '없음')).code, 'NOT_FOUND');
    assert.equal(officeOpenProblem(pptx, { maxBytes: 10 }).code, 'TOO_BIG');
    assert.equal(officeOpenProblem(pptx), null);
    assert.equal(fake.windows.length, n, '창을 열지 않는다');
    await assert.rejects(mgr.openEditor(join(docs, '없음.docx')), { code: 'NOT_FOUND' });
    writeFileSync(join(docs, 'a.doc'), 'x');
    await assert.rejects(mgr.openEditor(join(docs, 'a.doc')), (e) => e.code === 'UNSUPPORTED' && /\.docx/.test(e.message));
  });

  await t.test('checkOoxmlBytes: 주 부분 이름이 달라도 _rels/.rels가 가리키면 맞음', async () => {
    const z = await JSZip.loadAsync(readFileSync(docx));
    const doc = await z.file('word/document.xml').async('string');
    z.remove('word/document.xml');
    z.file('word/document2.xml', doc);
    const rels = (await z.file('_rels/.rels').async('string')).replace(/word\/document\.xml/, 'word/document2.xml');
    z.file('_rels/.rels', rels);
    const b = new Uint8Array(await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    assert.equal(checkOoxmlBytes(b, 'docx'), null);
    assert.match(checkOoxmlBytes(b, 'xlsx'), /xl\/workbook\.xml/);
    assert.equal(checkOoxmlBytes(new Uint8Array(readFileSync(xlsx)), 'xlsx'), null);
    assert.equal(checkOoxmlBytes(new Uint8Array(readFileSync(pptx)), 'pptx'), null);
    assert.match(checkOoxmlBytes(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0]), 'docx'), /손상/);
  });

  await t.test('메모리 안내: Office 편집 창이 4개를 넘으면 notice', async () => {
    const extra = [];
    for (let i = 0; i < 3; i++) { const p = join(docs, `더${i}.docx`); await createBlankDocument(p); extra.push(p); }
    const counts = [];
    for (const p of extra) counts.push((await mgr.openEditor(p)).notice);
    // 지금 열린 Office 창: 메모 사본·표·발표 + 3
    assert.equal(counts[0], null);
    assert.match(counts[1], /문서 편집 창이 5개 열려 있습니다/);
    assert.match(counts[2], /6개/);
  });

  await t.test('렌더러가 죽으면 안내 후 host를 다시 연다 (복구 사본 안내)', async () => {
    const d0 = fake.dialogs.length;
    dw.webContents.mainFrame.url = 'about:blank';
    dw.webContents.emit('render-process-gone', {}, { reason: 'crashed' });
    await until(() => fake.dialogs.length > d0);
    const d = fake.dialogs.at(-1);
    assert.equal(d.title, '문서 편집');
    assert.equal(d.message, '편집 창이 예기치 않게 멈췄습니다.');
    assert.match(d.detail, /자동 저장된 편집 내용이 있으면/);
    assert.ok(await until(() => /\/office-host\/host\.html$/.test(dw.webContents.mainFrame.url)));
  });

  await t.test('닫기: 고친 창은 묻고(저장 → host 저장 명령), 정상 닫기는 복구 사본을 지운다, 강제 닫기는 남긴다', async () => {
    // 저장 → host가 editor:save → 닫힘
    const pw = win(2);
    fake.replies.set(pw.webContents.id, true);
    fake.dialogAnswers.push(0); // 저장
    const s0 = saved.length;
    fake.commandAnswer = async (wc, command) => {
      if (command !== 'save') return false;
      const r = await fake.invoke(wc, 'editor:save', new Uint8Array(readFileSync(pptx)), { reason: 'user' });
      return r.ok;
    };
    let d0 = fake.dialogs.length;
    pw.close();
    assert.ok(await until(() => pw.isDestroyed()), '저장 뒤 닫힌다');
    assert.equal(fake.dialogs[d0].title, '슬라이드 편집');
    assert.deepEqual(fake.dialogs[d0].buttons, ['저장', '저장 안 함', '취소']);
    assert.equal(saved.length, s0 + 1);
    // 저장 안 함 → 닫히고 복구 사본도 지운다 (정상 닫기)
    const xw = win(1);
    assert.equal((await call(xw, 'editor:save', new Uint8Array(readFileSync(xlsx)), { reason: 'auto' })).ok, true);
    const draft = draftPathFor(ud, xlsx);
    assert.equal(existsSync(draft), true);
    fake.replies.set(xw.webContents.id, true);
    fake.dialogAnswers.push(1); // 저장 안 함
    d0 = fake.dialogs.length;
    xw.close();
    assert.ok(await until(() => xw.isDestroyed()));
    assert.equal(fake.dialogs[d0].title, '표 편집');
    assert.equal(existsSync(draft), false, '정상 닫기 → 복구 사본 지움');
    // 취소 → 그대로
    const ow = fake.windows.find((w) => !w.isDestroyed() && /더0\.docx/.test(w.getTitle()));
    const od = draftPathFor(ud, join(docs, '더0.docx'));
    assert.equal((await call(ow, 'editor:save', new Uint8Array(readFileSync(join(docs, '더0.docx'))), { reason: 'auto' })).ok, true);
    fake.replies.set(ow.webContents.id, true);
    fake.dialogAnswers.push(2); // 취소
    assert.equal(await mgr.closeAll(), false);
    assert.equal(ow.isDestroyed(), false);
    // 강제 닫기(closeAll force)는 사본을 남긴다
    fake.commandAnswer = null;
    await mgr.closeAll({ force: true });
    assert.equal(ow.isDestroyed(), true);
    assert.equal(existsSync(od), true, '강제로 닫으면 복구 사본을 남긴다');
    assert.deepEqual(mgr.list(), []);
  });
});
