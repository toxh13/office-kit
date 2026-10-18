// 앱(main.mjs)과 한글 편집 창 잇기 시험.
//  1) 가짜 electron 모듈로 main.mjs를 그대로 띄워 IPC 흐름을 확인한다 (창은 가짜, 편집기 서버·파일·백업은 진짜)
//  2) OFFICEKIT_EDITOR_SMOKE=1 이고 vendor/ 묶음이 있으면 실제 Electron으로 두 창을 찍는다 (창이 잠깐 뜬다)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { register } from 'node:module';
import { EventEmitter } from 'node:events';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBlankDocument } from '../src/core/create.mjs';
import { readDocument } from '../src/core/index.mjs';

const readDocumentText = async (p) => (await readDocument(p, { engine: 'file' })).text ?? '';

const here = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(here, '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (p) => createHash('sha1').update(readFileSync(p)).digest('hex');

async function until(fn, ms = 8000, step = 40) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) return v;
    await sleep(step);
  }
}

// ── 가짜 electron ────────────────────────────────────────

function makeFakeElectron({ userData, documents }) {
  const fake = { windows: [], dialogAnswers: [], dialogs: [], replies: new Map(), sessionHandlers: {}, external: [], saveDialogAnswers: [] };
  let wcSeq = 0;

  const app = Object.assign(new EventEmitter(), {
    isPackaged: false,
    paths: { userData, documents },
    setPath(k, v) { this.paths[k] = v; },
    getPath(k) { return this.paths[k] ?? tmpdir(); },
    requestSingleInstanceLock: () => true,
    quit() { this.quitCalls += 1; },
    exit(code) { this.exitCode = code; },
    whenReady: () => Promise.resolve(),
    getVersion: () => '0.0.0-test',
    addRecentDocument() {},
    quitCalls: 0,
    exitCode: undefined,
  });

  class WebContents extends EventEmitter {
    constructor() {
      super();
      this.id = ++wcSeq;
      this.sent = [];
      this.session = new EventEmitter();
      this.mainFrame = { url: '' };
    }
    send(channel, data) {
      this.sent.push({ channel, data });
      // 편집 창 host 흉내: 닫기 전 질문에 답한다 (fake.replies: webContents id → 저장 안 한 내용 있음?)
      if (channel === 'editor:request-close') {
        setImmediate(() => emit('editor:reply', this, { id: data.id, value: fake.replies.get(this.id) ?? false }));
      }
    }
    setWindowOpenHandler(fn) { this.openHandler = fn; }
    setWebRTCIPHandlingPolicy(p) { this.webrtcPolicy = p; }
    loadFile(f) { this.mainFrame.url = pathToFileURL(f).href; return Promise.resolve(); }
    loadURL(u) { this.mainFrame.url = u; return Promise.resolve(); }
    toggleDevTools() {}
    executeJavaScript() { return Promise.resolve(null); }
  }

  class BrowserWindow extends EventEmitter {
    static all = new Set();
    static getAllWindows() { return [...BrowserWindow.all]; }
    constructor(opts = {}) {
      super();
      this.opts = opts;
      this.title = opts.title ?? '';
      this.destroyed = false;
      this.webContents = new WebContents();
      BrowserWindow.all.add(this);
      fake.windows.push(this);
      app.emit('browser-window-created', {}, this);
    }
    isDestroyed() { return this.destroyed; }
    close() {
      if (this.destroyed) return;
      let prevented = false;
      this.emit('close', { preventDefault: () => { prevented = true; } });
      if (!prevented) this.destroy();
    }
    destroy() {
      if (this.destroyed) return;
      this.destroyed = true;
      BrowserWindow.all.delete(this);
      this.emit('closed');
    }
    loadFile(f) { return this.webContents.loadFile(f); }
    loadURL(u) { return this.webContents.loadURL(u); }
    setTitle(t) { this.title = t; }
    getTitle() { return this.title; }
    show() {} focus() {} restore() {} maximize() {} moveTop() {} removeMenu() {}
    isMinimized() { return false; }
    isMaximized() { return false; }
    getNormalBounds() { return { x: 0, y: 0, width: 1200, height: 800 }; }
  }

  const ipcMain = {
    handlers: new Map(), listeners: new Map(),
    handle(channel, fn) {
      if (this.handlers.has(channel)) throw new Error(`이미 등록된 채널: ${channel}`);
      this.handlers.set(channel, fn);
    },
    on(channel, fn) {
      if (!this.listeners.has(channel)) this.listeners.set(channel, []);
      this.listeners.get(channel).push(fn);
    },
  };
  const evOf = (wc) => ({ sender: wc, senderFrame: wc.mainFrame });
  function emit(channel, wc, ...args) { for (const fn of ipcMain.listeners.get(channel) ?? []) fn(evOf(wc), ...args); }
  async function invokeFrom(wc, channel, ...args) {
    const fn = ipcMain.handlers.get(channel);
    if (!fn) throw new Error(`없는 채널: ${channel}`);
    return fn(evOf(wc), ...args);
  }

  const editorSession = {
    webRequest: { onBeforeRequest(_filter, fn) { fake.sessionHandlers.beforeRequest = fn; } },
    setPermissionRequestHandler(fn) { fake.sessionHandlers.permission = fn; },
    setPermissionCheckHandler() {}, setDevicePermissionHandler() {}, setDisplayMediaRequestHandler() {},
    setSpellCheckerEnabled() {}, on() {},
    setSpellCheckerDictionaryDownloadURL(u) { fake.sessionHandlers.dictionaryUrl = u; },
    async setProxy(cfg) { fake.sessionHandlers.proxy = cfg; },
  };

  Object.assign(fake, {
    app, BrowserWindow, ipcMain, emit, invokeFrom,
    Menu: { buildFromTemplate: (t) => t, setApplicationMenu() {} },
    dialog: {
      async showMessageBox(_win, opts) { fake.dialogs.push(opts); return { response: fake.dialogAnswers.shift() ?? 2 }; },
      async showSaveDialog() { return fake.saveDialogAnswers.shift() ?? { canceled: true }; },
      async showOpenDialog() { return { canceled: true, filePaths: [] }; },
      showErrorBox(title, text) { fake.errorBox = `${title}: ${text}`; },
    },
    nativeTheme: { shouldUseDarkColors: false, themeSource: 'system' },
    screen: { getAllDisplays: () => [{ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }] },
    session: {
      defaultSession: { setPermissionRequestHandler() {}, setPermissionCheckHandler() {}, setSpellCheckerDictionaryDownloadURL(u) { fake.defaultDictionaryUrl = u; } },
      fromPartition: () => editorSession,
    },
    shell: { openPath: async () => '', openExternal: async (u) => { fake.external.push(u); }, showItemInFolder() {} },
  });
  return fake;
}

/** 시험용 편집기 묶음 (vendor/rhwp-studio, vendor/rhwp-editor-sdk 흉내) */
function makeVendor(dir) {
  const put = (p, text = 'x') => { mkdirSync(dirname(p), { recursive: true }); writeFileSync(p, text); };
  put(join(dir, 'rhwp-studio', 'index.html'), '<!doctype html><title>studio</title>');
  put(join(dir, 'rhwp-studio', 'assets', 'rhwp_bg-test.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(dir, 'rhwp-editor-sdk', f), 'export {};');
}

const FAKE_AI = `
export async function detectProviders() { return [{ id: 'claude', name: 'Claude (시험)', available: true, version: '1', hint: '' }]; }
export function defaultWorkspace() { return process.env.OFFICEKIT_TEST_WS; }
export function startChat({ prompt, onEvent }) {
  globalThis.__okPrompts = [...(globalThis.__okPrompts ?? []), prompt];
  const done = Promise.resolve().then(() => { onEvent({ type: 'done', sessionId: 's1', text: '네' }); return { sessionId: 's1', text: '네' }; });
  return { cancel() {}, done };
}
`;

test('main.mjs ↔ 한글 편집 창 (가짜 electron)', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-int-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  const docPath = join(docs, '보고서.hwpx');
  await createBlankDocument(docPath, { title: '처음 내용' });
  const docxPath = join(docs, '메모.docx');
  await createBlankDocument(docxPath);
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);

  const saved = { argv: process.argv, env: { ...process.env } };
  Object.assign(process.env, {
    OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_VENDOR: vendor, OFFICEKIT_AI_MODULE: aiModule, OFFICEKIT_TEST_WS: join(base, 'ws'),
  });
  for (const k of ['OFFICEKIT_SMOKE', 'OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_TEST']) delete process.env[k];
  process.argv = [process.argv[0], docPath];

  const fake = makeFakeElectron({ userData: ud, documents: docs });
  globalThis.__okFakeElectron = fake;
  const names = ['app', 'BrowserWindow', 'Menu', 'dialog', 'ipcMain', 'nativeTheme', 'screen', 'session', 'shell'];
  const fakeSource = `const f = globalThis.__okFakeElectron; ${names.map((n) => `export const ${n} = f.${n};`).join(' ')} export default f;`;
  register(`data:text/javascript,${encodeURIComponent(`
    export async function resolve(spec, ctx, next) {
      if (spec === 'electron') return { url: ${JSON.stringify(`data:text/javascript,${encodeURIComponent(fakeSource)}`)}, shortCircuit: true };
      return next(spec, ctx);
    }`)}`);

  t.after(async () => {
    fake.app.emit('will-quit');
    for (const w of fake.BrowserWindow.getAllWindows()) w.destroy();
    await sleep(100);
    process.argv = saved.argv;
    for (const k of Object.keys(process.env)) if (!(k in saved.env)) delete process.env[k];
    Object.assign(process.env, saved.env);
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  });

  await import(pathToFileURL(join(ROOT, 'src', 'app', 'main.mjs')).href);
  assert.ok(await until(() => fake.windows.length >= 1, 20_000), '주 창이 만들어져야 한다');
  const mainWin = fake.windows[0];
  const mainWc = mainWin.webContents;
  const call = (channel, ...args) => fake.invokeFrom(mainWc, channel, ...args);
  const lastSent = (channel, pred = () => true) => mainWc.sent.filter((s) => s.channel === channel && pred(s.data)).at(-1)?.data;
  assert.equal(fake.app.exitCode, undefined, fake.errorBox);

  let init;
  await t.test('app:init: 편집기 묶음 상태를 알려 준다 (없으면 한국어 이유)', async () => {
    init = await call('app:init');
    assert.equal(init.ok, true, init.error);
    assert.deepEqual(init.value.editor.exts, ['hwp', 'hwpx', 'hwt']);
    assert.equal(init.value.editor.available, false);
    assert.match(init.value.editor.reason, /한글 편집기/);
    assert.equal(init.value.opening, true);
  });

  await t.test('문서 정보에 편집 창 상태가 들어 있다', async () => {
    const opened = await until(() => lastSent('doc:opened', (d) => d?.path === docPath));
    assert.ok(opened, '인수로 받은 문서를 연다');
    assert.deepEqual(opened.editor, { open: false, dirty: false });
  });

  await t.test('doc:edit: Word 문서는 거절하고 프로그램에서 열기를 안내', async () => {
    assert.equal((await call('doc:open', docxPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'UNSUPPORTED');
    assert.match(r.error, /프로그램에서 열기/);
    assert.equal(fake.windows.length, 1, '편집 창을 띄우지 않는다');
  });

  await t.test('doc:edit: 편집기 묶음이 없으면 EDITOR_MISSING (창을 띄우지 않음)', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'EDITOR_MISSING');
    assert.match(r.error, /한글 편집기 파일이 없습니다/);
    assert.equal(fake.windows.length, 1);
  });

  makeVendor(vendor);
  let editorWin; let hostCall;
  await t.test('doc:edit: 묶음이 있으면 편집 창을 열고, 다시 누르면 같은 창', async () => {
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.path, docPath);
    assert.equal(r.value.reused, false);
    editorWin = fake.windows[1];
    assert.ok(editorWin, '편집 창');
    assert.match(editorWin.getTitle(), /보고서\.hwpx - 한글 편집 \(Office Kit\)/);
    assert.match(editorWin.webContents.mainFrame.url, /^http:\/\/127\.0\.0\.1:\d+\/host\/host\.html$/);
    assert.equal(editorWin.opts.webPreferences.partition, 'officekit-editor');
    hostCall = (channel, ...args) => fake.invokeFrom(editorWin.webContents, channel, ...args);
    const again = await call('doc:edit');
    assert.equal(again.value.reused, true);
    assert.equal(fake.windows.length, 2);
    const st = await until(() => lastSent('doc:editor', (d) => d.open));
    assert.deepEqual(st, { path: docPath, open: true, dirty: false });
    // 편집 창 밖으로 나가는 요청은 막는다
    const blocked = await new Promise((r) => fake.sessionHandlers.beforeRequest({ url: 'https://cdn.jsdelivr.net/x.woff2' }, r));
    assert.equal(blocked.cancel, true);
    // HTTP 밖의 연결·맞춤법 사전 내려받기도 인터넷에 나가지 못하게 (죽은 프록시·사전 주소, WebRTC는 프록시 안 거친 UDP 금지)
    assert.deepEqual(fake.sessionHandlers.proxy, { mode: 'fixed_servers', proxyRules: 'socks5://127.0.0.1:9', proxyBypassRules: '127.0.0.1;127.0.0.2' });
    assert.equal(fake.sessionHandlers.dictionaryUrl, 'http://127.0.0.1:9/');
    assert.equal(fake.defaultDictionaryUrl, 'http://127.0.0.1:9/');
    assert.equal(editorWin.webContents.webrtcPolicy, 'disable_non_proxied_udp');
  });

  await t.test('편집 창 새 창·링크: 이 PC·사설망 주소는 열지 않고, 바깥 주소는 물어본 뒤 기본 브라우저로', async () => {
    const open = editorWin.webContents.openHandler;
    const d0 = fake.dialogs.length;
    for (const url of ['http://127.0.0.2:5555/studio/print.html', 'http://127.0.0.1:1/x', 'http://localhost/', 'http://192.168.0.10/',
      'http://10.1.2.3/', 'http://[::1]/', 'http://[::ffff:127.0.0.1]/', 'http://2130706433/', 'file:///C:/Windows/win.ini', 'javascript:alert(1)',
      'https://user:pw@example.com/']) {
      assert.deepEqual(open({ url }), { action: 'deny' }, url);
    }
    await sleep(50);
    assert.equal(fake.dialogs.length, d0, '내부 주소는 묻지도 않는다');
    assert.deepEqual(fake.external, []);
    // 바깥 주소: 취소하면 열지 않고, 열기를 고르면 연다
    fake.dialogAnswers.push(1);
    assert.deepEqual(open({ url: 'https://example.com/a?b=1' }), { action: 'deny' });
    await until(() => fake.dialogs.length > d0);
    await sleep(20);
    assert.match(fake.dialogs.at(-1).message, /기본 브라우저로 열까요/);
    assert.equal(fake.dialogs.at(-1).detail, 'https://example.com/a?b=1');
    assert.deepEqual(fake.external, []);
    fake.dialogAnswers.push(0);
    open({ url: 'https://example.com/a?b=1' });
    assert.ok(await until(() => fake.external.length === 1));
    assert.deepEqual(fake.external, ['https://example.com/a?b=1']);
    // studio의 문서 링크(빈 창 → 주소 바꾸기): 빈 창만 숨겨서, 스크립트 없이
    const blank = open({ url: 'about:blank' });
    assert.equal(blank.action, 'allow');
    assert.equal(blank.overrideBrowserWindowOptions.show, false);
    assert.equal(blank.overrideBrowserWindowOptions.webPreferences.javascript, false);
    assert.equal(blank.overrideBrowserWindowOptions.webPreferences.partition, 'officekit-editor');
    fake.external.length = 0;
  });

  await t.test('편집 창이 고친 상태면 AI 요청을 막고, editorOk면 보낸다', async () => {
    const doc = await hostCall('editor:get-document');
    assert.equal(doc.ok, true, doc.error);
    assert.equal(doc.value.name, '보고서.hwpx');
    fake.emit('editor:dirty', editorWin.webContents, true);
    const st = await until(() => lastSent('doc:editor', (d) => d.dirty));
    assert.deepEqual(st, { path: docPath, open: true, dirty: true });
    assert.ok(editorWin.getTitle().startsWith('* '));
    const refused = await call('ai:send', { prompt: '요약해줘' });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, 'EDITOR_DIRTY');
    assert.match(refused.error, /저장하지 않은 내용/);
    const sent = await call('ai:send', { prompt: '요약해줘', editorOk: true });
    assert.equal(sent.ok, true, sent.error);
    assert.ok(sent.value.chatId);
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end')), '대화가 끝난다');
  });

  const before = sha(docPath);
  let savedBytes;
  await t.test('편집 창 저장 → 백업·최근 목록·미리보기 알림(doc:changed by editor)', async () => {
    const tmpDoc = join(base, 'saved.hwpx');
    await createBlankDocument(tmpDoc, { title: '편집 창에서 고친 내용' });
    savedBytes = readFileSync(tmpDoc);
    const res = await hostCall('editor:save', new Uint8Array(savedBytes));
    assert.equal(res.ok, true, res.error);
    assert.equal(sha(docPath), createHash('sha1').update(savedBytes).digest('hex'));
    const changed = await until(() => lastSent('doc:changed', (d) => d.by === 'editor'));
    assert.equal(changed.path, docPath);
    assert.equal(changed.size, savedBytes.length);
    const settings = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));
    assert.equal(settings.recent[0].path, docPath);
    const index = JSON.parse(readFileSync(join(ud, 'backups', 'index.json'), 'utf8'));
    assert.ok(index.some((e) => e.source === docPath && e.hash === before), '덮어쓰기 전 내용이 백업돼 있다');
    fake.emit('editor:dirty', editorWin.webContents, false);
    assert.ok(await until(() => lastSent('doc:editor', (d) => !d.dirty)));
  });

  await t.test('바깥에서 파일이 바뀌면 편집 창에 알린다 (감시 → notifyExternalChange)', async () => {
    const n0 = editorWin.webContents.sent.filter((s) => s.channel === 'editor:external-change').length;
    const other = join(base, 'other.hwpx');
    await createBlankDocument(other, { title: '바깥에서 바꾼 내용' });
    writeFileSync(docPath, readFileSync(other));
    const got = await until(() => editorWin.webContents.sent.filter((s) => s.channel === 'editor:external-change').length > n0, 8000);
    assert.ok(got, '편집 창이 바깥 변경을 받아야 한다');
    assert.ok(await until(() => lastSent('doc:changed', (d) => !d.by && d.size === statSync(docPath).size), 8000), '주 창 미리보기도');
  });

  await t.test('원래대로 → 편집 창에도 알림 (editorNotified)', async () => {
    const n0 = editorWin.webContents.sent.filter((s) => s.channel === 'editor:external-change').length;
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(sha(docPath), before, '열 때의 내용으로');
    assert.equal(editorWin.webContents.sent.filter((s) => s.channel === 'editor:external-change').length, n0 + 1);
  });

  await t.test('원래대로: 파일이 이미 원본과 같아도 편집 창에 저장 안 한 편집이 있으면 알린다 (다시 불러올지 묻게)', async () => {
    assert.equal(sha(docPath), before);
    const count = () => editorWin.webContents.sent.filter((s) => s.channel === 'editor:external-change').length;
    // 깨끗하면 알릴 것이 없다
    let n0 = count();
    let r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, false);
    assert.equal(count(), n0);
    // 저장 안 한 편집이 있으면 내용이 같아도 알린다
    fake.emit('editor:dirty', editorWin.webContents, true);
    await until(() => lastSent('doc:editor', (d) => d.dirty));
    n0 = count();
    r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(count(), n0 + 1);
    assert.equal(sha(docPath), before);
    fake.emit('editor:dirty', editorWin.webContents, false);
    assert.ok(await until(() => lastSent('doc:editor', (d) => !d.dirty)));
  });

  await t.test('새 문서: 지웠던 문서와 같은 경로로 다시 만들면 예전 원래대로 기준·AI 대화를 잇지 않는다', async () => {
    const old = join(docs, '재생성.docx');
    await createBlankDocument(old, { title: '옛 보고서' });
    assert.equal((await call('doc:open', old)).ok, true);
    const sent = await call('ai:send', { prompt: '요약해줘' });
    assert.equal(sent.ok, true, sent.error);
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end' && d.path === old)), '대화가 끝난다');
    assert.ok(await until(async () => (await call('doc:open', old)).value.hasSession), '대화가 이어질 상태');
    assert.equal((await call('doc:open', docxPath)).ok, true);
    rmSync(old);
    const made = await call('doc:new', { format: 'docx', folder: docs, name: '재생성', start: 'blank' });
    assert.equal(made.ok, true, made.error);
    assert.equal(made.value.doc.path, old);
    assert.equal(made.value.doc.hasSession, false, '예전 AI 대화를 잇지 않는다');
    const blank = sha(old);
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(sha(old), blank, '원래대로는 새로 만든 빈 문서로 (지운 문서 내용이 아님)');
    assert.equal((await readDocumentText(old)).includes('옛 보고서'), false);
  });

  let newEditor;
  await t.test('새 문서 edit:true → 한글이면 편집 창, Word면 무시', async () => {
    const hw = await call('doc:new', { format: 'hwpx', folder: docs, name: '새 편집', start: 'blank', edit: true });
    assert.equal(hw.ok, true, hw.error);
    assert.equal(hw.value.editorError, null);
    assert.equal(hw.value.editor.path, join(docs, '새 편집.hwpx'));
    newEditor = fake.windows.at(-1);
    assert.match(newEditor.getTitle(), /새 편집\.hwpx/);
    const wd = await call('doc:new', { format: 'docx', folder: docs, name: '새 워드', start: 'blank', edit: true });
    assert.equal(wd.ok, true, wd.error);
    assert.equal(wd.value.editor, null);
    assert.equal(wd.value.editorError, null);
    assert.equal(fake.windows.length, 3, 'Word는 편집 창을 열지 않는다');
  });

  await t.test('다른 이름으로 저장: 확장자를 붙인 이름이 이미 있으면 다시 묻고, 취소하면 그대로', async () => {
    const h2 = (channel, ...args) => fake.invokeFrom(newEditor.webContents, channel, ...args);
    assert.equal((await h2('editor:get-document')).ok, true);
    const victim = join(docs, 'victim.docx.hwpx');
    writeFileSync(victim, 'VICTIM');
    const bytes = new Uint8Array(readFileSync(join(docs, '새 편집.hwpx')));
    const d0 = fake.dialogs.length;
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, 'victim.docx') });
    fake.dialogAnswers.push(1); // 취소
    let r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(r.value, { canceled: true });
    assert.equal(fake.dialogs.length, d0 + 1);
    assert.match(fake.dialogs.at(-1).message, /victim\.docx\.hwpx/);
    assert.equal(readFileSync(victim, 'utf8'), 'VICTIM');
    // 바꾸기를 고르면 저장
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, 'victim.docx') });
    fake.dialogAnswers.push(0);
    r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, 'victim.docx.hwpx');
    assert.equal(sha(victim), createHash('sha1').update(bytes).digest('hex'));
    // 확장자가 맞고 없는 이름이면 묻지 않는다
    const d1 = fake.dialogs.length;
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '다른 이름.hwpx') });
    r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.equal(fake.dialogs.length, d1);
  });

  let closeAsked = 0;
  await t.test('주 창 닫기: 고친 편집 창이 있으면 묻고, 취소하면 모두 그대로', async () => {
    fake.replies.set(editorWin.webContents.id, true);
    fake.dialogAnswers.push(2); // 취소
    closeAsked = fake.dialogs.length;
    mainWin.close();
    await until(() => fake.dialogs.length >= closeAsked + 1);
    await sleep(100);
    assert.equal(fake.dialogs.length, closeAsked + 1);
    assert.deepEqual(fake.dialogs[closeAsked].buttons, ['저장', '저장 안 함', '취소']);
    assert.match(fake.dialogs[closeAsked].message, /보고서\.hwpx/);
    assert.equal(mainWin.isDestroyed(), false);
    assert.equal(editorWin.isDestroyed(), false);
    assert.equal(newEditor.isDestroyed(), false);
  });

  await t.test('주 창 닫기: 저장 안 함 → 편집 창 둘 다 닫고 주 창도 닫힌다 (파일은 그대로)', async () => {
    const hash = sha(docPath);
    fake.dialogAnswers.push(1); // 저장 안 함
    mainWin.close();
    assert.ok(await until(() => mainWin.isDestroyed(), 8000), '주 창이 닫혀야 한다');
    assert.equal(editorWin.isDestroyed(), true);
    assert.equal(newEditor.isDestroyed(), true);
    assert.equal(fake.dialogs.length, closeAsked + 2, '깨끗한 편집 창은 묻지 않는다');
    assert.equal(sha(docPath), hash);
  });
});

// ── 실제 Electron 확인용 캡처 (선택) ───────────────────────

const wantSmoke = process.env.OFFICEKIT_EDITOR_SMOKE === '1';
const electronExe = !wantSmoke ? '' : (() => { try { return spawnSync(process.execPath, ['-e', "process.stdout.write(require('electron'))"], { cwd: ROOT, encoding: 'utf8' }).stdout; } catch { return ''; } })();
const smokeReady = wantSmoke && existsSync(join(ROOT, 'vendor', 'rhwp-studio', 'index.html')) && !!electronExe && existsSync(electronExe);

test('실제 앱: OFFICEKIT_SMOKE_EDITOR로 주 창 + 편집 창 캡처', { skip: !smokeReady && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/ 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async () => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-smoke-'));
  try {
    const doc = join(base, '캡처 시험.hwpx');
    await createBlankDocument(doc, { title: '캡처 시험 문서' });
    const png = join(base, 'shot.png');
    const env = { ...process.env, OFFICEKIT_SMOKE: png, OFFICEKIT_SMOKE_EDITOR: '1', OFFICEKIT_USER_DATA: join(base, 'ud') };
    delete env.ELECTRON_RUN_AS_NODE;
    const r = spawnSync(electronExe, ['.', doc], { cwd: ROOT, env, encoding: 'utf8', timeout: 180_000 });
    const line = (r.stdout ?? '').split(/\r?\n/).find((l) => l.startsWith('[smoke] '));
    assert.equal(r.status, 0, `${r.stdout}\n${r.stderr}`);
    const info = JSON.parse(line.slice('[smoke] '.length));
    assert.equal(info.ok, true);
    assert.equal(info.editor.ok, true);
    assert.equal(info.editor.state, 'saved');
    assert.equal(info.editor.title, '캡처 시험.hwpx - 한글 편집 (Office Kit)');
    assert.ok(statSync(png).size > 10_000);
    assert.ok(statSync(info.editor.png).size > 10_000);
  } finally {
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
});
