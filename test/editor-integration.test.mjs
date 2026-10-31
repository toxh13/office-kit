// 앱(main.mjs)과 편집 창(한글 · 문서·표·슬라이드) 잇기 시험.
//  1) 가짜 electron 모듈로 main.mjs를 그대로 띄워 IPC 흐름을 확인한다 (창은 가짜, 편집기 서버·파일·백업은 진짜)
//  2) Office 편집기 설정 보관(store.mjs), '정보 · 라이선스' 글(legal-ko.json ↔ docs/office-editor-legal-ko.md)
//  3) OFFICEKIT_EDITOR_SMOKE=1 이고 vendor/ 묶음이 있으면 실제 Electron으로 두 창을 찍는다 (창이 잠깐 뜬다)
//     한글(hwpx) + 문서·표·슬라이드(docx·xlsx·pptx). OFFICEKIT_SMOKE_OUT=<폴더>면 찍은 그림을 거기에 남긴다
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { register } from 'node:module';
import { EventEmitter } from 'node:events';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, utimesSync, writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBlankDocument, createDocument } from '../src/core/create.mjs';
import { readDocument, Hwp, hwpPdfStatus } from '../src/core/index.mjs';
import { createSettings, OFFICE_STATE_MAX_KEYS } from '../src/app/store.mjs';
import { draftPathFor } from '../src/app/editor-window.mjs';
import { startEditorServer } from '../src/app/editor-server.mjs';
import { NOTICE_FILES, aboutModel, legalSections } from '../src/app/office-editor/host.js';

const readDocumentText = async (p) => (await readDocument(p, { engine: 'file' })).text ?? '';

const here = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(here, '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (p) => createHash('sha1').update(readFileSync(p)).digest('hex');
const shaBytes = (b) => createHash('sha1').update(b).digest('hex');
const DAY = 24 * 60 * 60 * 1000;

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
  const fake = { windows: [], dialogAnswers: [], dialogs: [], replies: new Map(), sessions: {}, external: [], saveDialogAnswers: [], opened: [] };
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
    commandLine: { switches: new Map(), appendSwitch(k, v = '') { this.switches.set(k, v); } },
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
      // 저장 명령: 이 시험에서는 저장하지 못한 것으로 답한다 (180초 기다리지 않게)
      if (channel === 'editor:command') setImmediate(() => emit('editor:reply', this, { id: data.id, value: false }));
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
    show() { this.visible = true; } hide() { this.visible = false; } focus() {} restore() {} maximize() {} unmaximize() {} moveTop() {} removeMenu() {}
    isVisible() { return this.visible !== false; }
    isMinimized() { return false; }
    isMaximized() { return false; }
    getNormalBounds() { return { ...this.getBounds() }; }
    getBounds() { return { ...(this.bounds ?? { x: 0, y: 0, width: 1200, height: 800 }) }; }
    setBounds(b) { this.bounds = { ...this.getBounds(), ...b }; }
    setMinimumSize(w, h) { this.minSize = [w, h]; }
    getMinimumSize() { return this.minSize ?? [1000, 680]; }
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

  // 편집 창 세션 (partition마다 따로: 'officekit-editor' = 한글, 'officekit-office' = 문서·표·슬라이드)
  function makeSession(partition) {
    const h = {};
    fake.sessions[partition] = h;
    return {
      webRequest: { onBeforeRequest(_filter, fn) { h.beforeRequest = fn; } },
      setPermissionRequestHandler(fn) { h.permission = fn; },
      setPermissionCheckHandler(fn) { h.permissionCheck = fn; }, setDevicePermissionHandler() {}, setDisplayMediaRequestHandler() {},
      setSpellCheckerEnabled() {}, on() {},
      setSpellCheckerDictionaryDownloadURL(u) { h.dictionaryUrl = u; },
      async setProxy(cfg) { h.proxy = cfg; },
    };
  }
  const partitions = new Map();

  Object.assign(fake, {
    app, BrowserWindow, ipcMain, emit, invokeFrom,
    Menu: { buildFromTemplate: (t) => { fake.menu = t; return t; }, setApplicationMenu() {} },
    dialog: {
      async showMessageBox(_win, opts) { fake.dialogs.push(opts); return { response: fake.dialogAnswers.shift() ?? 2 }; },
      async showSaveDialog() { return fake.saveDialogAnswers.shift() ?? { canceled: true }; },
      async showOpenDialog() { return { canceled: true, filePaths: [] }; },
      showErrorBox(title, text) { fake.errorBox = `${title}: ${text}`; },
    },
    nativeTheme: { shouldUseDarkColors: false, themeSource: 'system' },
    screen: {
      getAllDisplays: () => [{ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }],
      getDisplayMatching: () => ({ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }),
    },
    session: {
      defaultSession: { setPermissionRequestHandler() {}, setPermissionCheckHandler() {}, setSpellCheckerDictionaryDownloadURL(u) { fake.defaultDictionaryUrl = u; } },
      fromPartition: (p) => { if (!partitions.has(p)) partitions.set(p, makeSession(p)); return partitions.get(p); },
    },
    shell: { openPath: async (p) => { fake.opened.push(p); return ''; }, openExternal: async (u) => { fake.external.push(u); }, showItemInFolder() {} },
  });
  return fake;
}

const put = (p, data = 'x') => { mkdirSync(dirname(p), { recursive: true }); writeFileSync(p, data); };

/** 시험용 한글 편집기 묶음 (vendor/rhwp-studio, vendor/rhwp-editor-sdk 흉내) */
function makeVendor(dir) {
  put(join(dir, 'rhwp-studio', 'index.html'), '<!doctype html><title>studio</title>');
  put(join(dir, 'rhwp-studio', 'assets', 'rhwp_bg-test.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(dir, 'rhwp-editor-sdk', f), 'export {};');
}

const FAKE_VERSION = {
  name: 'office-editor', officeKitVersion: '0.0.0-test', editorVersion: '9.4.0.129',
  onlyoffice: { version: '9.4.0', build: '9.4.0.129', sourceTag: 'v9.4.0.129' },
  filex: { tag: 'v0.3.1', commit: '57dda3dae99af78eb26b3c64b69f1e3f711f0757' },
  files: [],
};

/** 시험용 문서·표·슬라이드 편집기 묶음 (vendor/office-editor 흉내: 있는지 확인하는 파일만) */
function makeOfficeVendor(dir) {
  const o = join(dir, 'office-editor');
  put(join(o, 'index.html'), '<!doctype html><title>app</title>');
  put(join(o, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(o, 'x2t', 'x2t.js'), '// x2t');
  put(join(o, 'editor', 'LICENSE.txt'), 'GNU AFFERO GENERAL PUBLIC LICENSE');
  put(join(o, 'VERSION.json'), JSON.stringify(FAKE_VERSION));
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
// 글꼴 모듈 대신 (이 PC 글꼴을 읽지 않는다): 묶음 글꼴로
const FAKE_FONTS = `
export async function ensureFonts() { globalThis.__okFontCalls = (globalThis.__okFontCalls ?? 0) + 1; return { allFontsJsPath: null, thumbnailsDir: null, fontFileByIndex: () => null }; }
`;

test('main.mjs ↔ 편집 창 (가짜 electron): 한글 · 문서·표·슬라이드', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-int-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  const docPath = join(docs, '보고서.hwpx');
  await createBlankDocument(docPath, { title: '처음 내용' });
  const docxPath = join(docs, '메모.docx');
  await createDocument(docxPath, '# 회의 메모\n\n처음 쓴 내용입니다.\n\n- 안건 하나\n- 안건 둘');
  const csvPath = join(docs, '목록.csv');
  writeFileSync(csvPath, '이름,수량\n사과,3\n');
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);
  const fontsModule = join(base, 'fake-fonts.mjs');
  writeFileSync(fontsModule, FAKE_FONTS);

  // 지난 실행에서 남은 자동 저장 사본 (시작할 때 정리: 30일 넘은 것·쓰다 남은 임시 파일은 지우고 나머지는 둔다)
  const drafts = join(ud, 'office-drafts');
  const oldDraft = join(drafts, `${'a'.repeat(40)}.docx`);
  const freshDraft = join(drafts, `${'b'.repeat(40)}.xlsx`);
  const oldTmp = join(drafts, `.${'c'.repeat(40)}.docx.1234.abcd.tmp`);
  const freshTmp = join(drafts, `.${'d'.repeat(40)}.pptx.1234.abcd.tmp`);
  for (const f of [oldDraft, freshDraft, oldTmp, freshTmp]) put(f, 'PK');
  utimesSync(oldDraft, new Date(Date.now() - 40 * DAY), new Date(Date.now() - 40 * DAY));
  utimesSync(oldTmp, new Date(Date.now() - 2 * 60 * 60 * 1000), new Date(Date.now() - 2 * 60 * 60 * 1000));

  const saved = { argv: process.argv, env: { ...process.env } };
  Object.assign(process.env, {
    OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_VENDOR: vendor, OFFICEKIT_AI_MODULE: aiModule, OFFICEKIT_TEST_WS: join(base, 'ws'),
    OFFICEKIT_OFFICE_FONTS_MODULE: fontsModule,
    // 이 시험은 주 창에서 문서를 여는 흐름을 본다 (탐색기에서 연 문서의 '문서 창 방식'은 끈다)
    OFFICEKIT_DOC_WINDOW: '0',
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
  const sentTo = (w, channel) => w.webContents.sent.filter((s) => s.channel === channel);
  assert.equal(fake.app.exitCode, undefined, fake.errorBox);

  await t.test('시작: 이 PC 밖 이름은 풀지 않는다 (편집기 틀의 dns-prefetch로 내용이 새지 않게), 편집기 서버 주소는 뺀다', () => {
    const rules = fake.app.commandLine.switches.get('host-resolver-rules');
    assert.equal(rules, 'MAP * ~NOTFOUND, EXCLUDE localhost, EXCLUDE 127.0.0.1, EXCLUDE 127.0.0.2, EXCLUDE 127.0.0.3');
  });

  await t.test('시작: 오래된 자동 저장 사본과 쓰다 남은 임시 파일을 지운다 (최근 것은 둔다)', () => {
    assert.equal(existsSync(oldDraft), false, '30일 넘은 사본');
    assert.equal(existsSync(oldTmp), false, '한 시간 넘은 임시 파일');
    assert.equal(existsSync(freshDraft), true);
    assert.equal(existsSync(freshTmp), true, '쓰는 중일 수 있는 임시 파일은 남긴다');
  });

  await t.test('메뉴: 직접 편집(Ctrl+E)은 종류를 가리지 않는 이름', () => {
    const items = fake.menu?.[0]?.submenu ?? [];
    const edit = items.find((i) => i.accelerator === 'CmdOrCtrl+E');
    assert.equal(edit?.label, '직접 편집');
    assert.doesNotMatch(JSON.stringify(fake.menu.map((m) => m.label)), /ONLYOFFICE/i);
  });

  let init;
  await t.test('app:init: 종류별 편집기 묶음 상태를 알려 준다 (없으면 한국어 이유)', async () => {
    init = await call('app:init');
    assert.equal(init.ok, true, init.error);
    const ed = init.value.editor;
    assert.deepEqual(ed.exts, ['hwp', 'hwpx', 'hwt', 'docx', 'xlsx', 'pptx']);
    assert.deepEqual(ed.kinds.hwp.exts, ['hwp', 'hwpx', 'hwt']);
    assert.deepEqual(ed.kinds.office.exts, ['docx', 'xlsx', 'pptx']);
    assert.equal(ed.kinds.hwp.available, false);
    assert.match(ed.kinds.hwp.reason, /한글 편집기/);
    assert.equal(ed.kinds.office.available, false);
    assert.match(ed.kinds.office.reason, /문서 편집기 파일이 없습니다 \(vendor\/office-editor\)/);
    assert.match(ed.kinds.office.reason, /npm run office-editor:build/, '개발 실행이면 만드는 법');
    assert.equal(init.value.opening, true);
  });

  await t.test('문서 정보에 편집 창 상태가 들어 있다', async () => {
    const opened = await until(() => lastSent('doc:opened', (d) => d?.path === docPath));
    assert.ok(opened, '인수로 받은 문서를 연다');
    assert.deepEqual(opened.editor, { open: false, dirty: false });
  });

  await t.test('doc:edit: 편집 창 형식이 아니면(csv) 거절하고 프로그램에서 열기를 안내', async () => {
    assert.equal((await call('doc:open', csvPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'UNSUPPORTED');
    assert.match(r.error, /\.csv/);
    assert.match(r.error, /프로그램에서 열기/);
    assert.equal(fake.windows.length, 1, '편집 창을 띄우지 않는다');
  });

  await t.test('doc:edit: Word 문서도 직접 편집 대상 — 묶음이 없으면 EDITOR_MISSING (창을 띄우지 않음)', async () => {
    assert.equal((await call('doc:open', docxPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'EDITOR_MISSING');
    assert.match(r.error, /문서 편집기 파일이 없습니다/);
    assert.equal(fake.windows.length, 1);
  });

  await t.test('doc:edit: 한글 편집기 묶음이 없으면 EDITOR_MISSING (창을 띄우지 않음)', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'EDITOR_MISSING');
    assert.match(r.error, /한글 편집기 파일이 없습니다/);
    assert.equal(fake.windows.length, 1);
  });

  makeVendor(vendor);
  makeOfficeVendor(vendor);

  await t.test('app:init: 묶음을 만들면 두 종류 모두 쓸 수 있다', async () => {
    const r = await call('app:init');
    assert.equal(r.value.editor.kinds.hwp.available, true, r.value.editor.kinds.hwp.reason);
    assert.equal(r.value.editor.kinds.office.available, true, r.value.editor.kinds.office.reason);
    assert.equal(r.value.editor.kinds.office.reason, null);
  });

  let editorWin; let hostCall;
  await t.test('doc:edit: 묶음이 있으면 편집 창을 열고, 다시 누르면 같은 창', async () => {
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.path, docPath);
    assert.equal(r.value.kind, 'hwp');
    assert.equal(r.value.reused, false);
    assert.equal(r.value.notice, null);
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
    const ses = fake.sessions['officekit-editor'];
    const blocked = await new Promise((r) => ses.beforeRequest({ url: 'https://cdn.jsdelivr.net/x.woff2' }, r));
    assert.equal(blocked.cancel, true);
    // HTTP 밖의 연결·맞춤법 사전 내려받기도 인터넷에 나가지 못하게 (죽은 프록시·사전 주소, WebRTC는 프록시 안 거친 UDP 금지)
    // 바로 가는 것은 host·studio 서버의 주소:포트뿐 (<-loopback>: 이 PC의 다른 포트는 죽은 프록시로)
    assert.match(ses.proxy.proxyBypassRules, /^<-loopback>;127\.0\.0\.1:\d+;127\.0\.0\.[12]:\d+$/);
    assert.deepEqual({ ...ses.proxy, proxyBypassRules: null }, { mode: 'fixed_servers', proxyRules: 'socks5://127.0.0.1:9', proxyBypassRules: null });
    assert.equal(ses.dictionaryUrl, 'http://127.0.0.1:9/');
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
    assert.equal(doc.value.kind, 'hwp');
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
    assert.equal(sha(docPath), shaBytes(savedBytes));
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
    const n0 = sentTo(editorWin, 'editor:external-change').length;
    const other = join(base, 'other.hwpx');
    await createBlankDocument(other, { title: '바깥에서 바꾼 내용' });
    writeFileSync(docPath, readFileSync(other));
    const got = await until(() => sentTo(editorWin, 'editor:external-change').length > n0, 8000);
    assert.ok(got, '편집 창이 바깥 변경을 받아야 한다');
    assert.ok(await until(() => lastSent('doc:changed', (d) => !d.by && d.size === statSync(docPath).size), 8000), '주 창 미리보기도');
  });

  await t.test('원래대로 → 편집 창에도 알림 (editorNotified)', async () => {
    const n0 = sentTo(editorWin, 'editor:external-change').length;
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(sha(docPath), before, '열 때의 내용으로');
    assert.equal(sentTo(editorWin, 'editor:external-change').length, n0 + 1);
  });

  await t.test('원래대로: 파일이 이미 그 버전과 같아도 편집 창에 저장 안 한 편집이 있으면 알린다 (다시 불러올지 묻게)', async () => {
    assert.equal(sha(docPath), before);
    const count = () => sentTo(editorWin, 'editor:external-change').length;
    const same = (await call('doc:versions')).value.versions.find((v) => v.current);
    assert.ok(same, '지금 내용과 같은 버전');
    // 깨끗하면 알릴 것이 없다
    let n0 = count();
    let r = await call('doc:restore', { file: same.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, false);
    assert.equal(count(), n0);
    // 저장 안 한 편집이 있으면 내용이 같아도 알린다
    fake.emit('editor:dirty', editorWin.webContents, true);
    await until(() => lastSent('doc:editor', (d) => d.dirty));
    n0 = count();
    r = await call('doc:restore', { file: same.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(count(), n0 + 1);
    assert.equal(sha(docPath), before);
    fake.emit('editor:dirty', editorWin.webContents, false);
    assert.ok(await until(() => lastSent('doc:editor', (d) => !d.dirty)));
  });

  await t.test('버전 목록: 열 때·편집 창 저장 전·AI 대화 전·원래대로 전, 원래대로(인수 없음) = 마지막 변경 전', async () => {
    const w = join(base, 'ver-w.hwpx'); const x = join(base, 'ver-x.hwpx'); const y = join(base, 'ver-y.hwpx');
    await createBlankDocument(w, { title: '편집 창 저장 전 내용' });
    await createBlankDocument(x, { title: 'AI 대화 전 내용' });
    await createBlankDocument(y, { title: 'AI가 고친 내용' });
    writeFileSync(docPath, readFileSync(w));
    // 편집 창 저장 = 덮어쓰기 전 상태를 버전으로
    const saved = await hostCall('editor:save', new Uint8Array(readFileSync(x)));
    assert.equal(saved.ok, true, saved.error);
    assert.equal(sha(docPath), sha(x));
    // AI 대화 한 번 = 고치기 전 상태를 버전으로
    assert.equal((await call('ai:send', { prompt: '고쳐줘', editorOk: true })).ok, true);
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end')));
    writeFileSync(docPath, readFileSync(y)); // AI가 고친 것처럼
    let v = (await call('doc:versions')).value;
    const reasons = v.versions.map((e) => e.reason);
    for (const want of ['open', 'editor', 'ai', 'restore']) assert.ok(reasons.includes(want), `${want}: ${reasons.join(', ')}`);
    assert.ok(v.versions.length <= 30);
    const undo = v.versions.find((e) => e.file === v.undo);
    assert.equal(undo.reason, 'ai', '마지막 변경(AI) 전');
    assert.equal(v.versions.some((e) => e.current), false, 'AI가 고친 지금 내용은 아직 버전이 아님');
    let r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.reason, 'ai');
    assert.equal(sha(docPath), sha(x));
    // 되돌리기 전 상태(AI가 고친 것)가 '원래대로 전'으로 남아 다시 돌아올 수 있다
    v = (await call('doc:versions')).value;
    assert.equal(v.versions[0].reason, 'restore');
    assert.equal(v.undo, v.versions[0].file);
    assert.equal(v.versions.find((e) => e.current)?.reason, 'ai');
    // 처음 연 상태로
    const opened = v.versions.find((e) => e.reason === 'open');
    r = await call('doc:restore', { file: opened.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(sha(docPath), before);
    // 목록에 없는 이름은 거절
    r = await call('doc:restore', { file: '..\settings.json' });
    assert.equal(r.ok, false);
    assert.equal(r.code, 'NO_VERSION');
    assert.equal(sha(docPath), before);
  });

  // ── 문서·표·슬라이드 편집 창 (Office) ──

  let officeWin; let officeCall;
  const docxBefore = sha(docxPath);
  await t.test('한글 편집 창: 자동 저장은 복구 사본에만, 다시 열면 복구/버리기, 원본은 그대로', async () => {
    const before = readFileSync(docPath);
    const bytes = new Uint8Array(before);
    const r = await hostCall('editor:save', bytes, { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.draft, true);
    const draft = draftPathFor(ud, docPath);
    assert.ok(existsSync(draft), '복구 사본');
    assert.deepEqual(readFileSync(docPath), before, '원본은 그대로');
    // 파일보다 새로운 사본 → 다시 열면 알린다
    utimesSync(draft, new Date(Date.now() + 5000), new Date(Date.now() + 5000));
    const d = await hostCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.ok(d.value.draft && d.value.draft.size === bytes.length, JSON.stringify(d.value.draft));
    const read = await hostCall('editor:draft', 'read');
    assert.equal(read.ok, true, read.error);
    assert.equal(read.value.bytes.length, bytes.length);
    // 형식이 아닌 바이트는 사본으로도 쓰지 않는다
    const bad = await hostCall('editor:save', new Uint8Array(Buffer.from('not hwp')), { reason: 'auto' });
    assert.equal(bad.ok, false);
    assert.equal((await hostCall('editor:draft', 'discard')).ok, true);
    assert.equal(existsSync(draft), false);
  });

  await t.test('한글 편집 창: 지금 내용을 PDF로 저장 (편집기 인쇄 대신)', { skip: !hwpPdfStatus().ok && '내장 PDF 엔진(rhwp.exe) 없음' }, async () => {
    const h = Hwp.open(docPath);
    let hwpBytes;
    try { hwpBytes = new Uint8Array(h.toBytes('hwp')); } finally { h.close(); }
    const target = join(docs, '보고서 출력.pdf');
    fake.saveDialogAnswers.push({ canceled: false, filePath: target });
    const r = await hostCall('editor:hwp-pdf', hwpBytes, 'save');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '보고서 출력.pdf');
    assert.equal(readFileSync(target).subarray(0, 5).toString('latin1'), '%PDF-');
    // hwp가 아닌 바이트는 거절
    const no = await hostCall('editor:hwp-pdf', new Uint8Array(Buffer.from('PK\x03\x04')), 'save');
    assert.equal(no.ok, false);
    assert.equal(no.code, 'BAD_FORMAT');
  });

  await t.test('Office: docx를 문서 편집 창으로 연다 (따로 된 세션·네트워크 막기)', async () => {
    assert.equal((await call('doc:open', docxPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.kind, 'office');
    assert.equal(r.value.path, docxPath);
    assert.equal(r.value.reused, false);
    assert.equal(r.value.notice, null);
    officeWin = fake.windows.at(-1);
    assert.equal(officeWin.getTitle(), '메모.docx - 문서 편집 (Office Kit)');
    assert.doesNotMatch(officeWin.getTitle(), /ONLYOFFICE/i);
    assert.match(officeWin.webContents.mainFrame.url, /^http:\/\/127\.0\.0\.1:\d+\/office-host\/host\.html$/);
    assert.equal(officeWin.opts.webPreferences.partition, 'officekit-office');
    assert.match(officeWin.opts.webPreferences.preload, /office-preload\.cjs$/);
    assert.equal(officeWin.webContents.webrtcPolicy, 'disable_non_proxied_udp');
    const ses = fake.sessions['officekit-office'];
    assert.equal(ses.proxy.proxyRules, 'socks5://127.0.0.1:9');
    const blocked = await new Promise((res) => ses.beforeRequest({ url: 'https://example.com/x.js' }, res));
    assert.equal(blocked.cancel, true);
    officeCall = (channel, ...args) => fake.invokeFrom(officeWin.webContents, channel, ...args);
    const st = await until(() => lastSent('doc:editor', (d) => d.path === docxPath && d.open));
    assert.deepEqual(st, { path: docxPath, open: true, dirty: false });
    // 한글 편집 창의 host에서 Office 전용 요청은 거절
    const no = await hostCall('editor:state', 'get', 'x');
    assert.equal(no.ok, false);
  });

  await t.test('Office: get-document → 편집기 주소·한국어·판 정보 (복구 사본 없음)', async () => {
    const d = await officeCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    const v = d.value;
    assert.equal(v.kind, 'office');
    assert.equal(v.name, '메모.docx');
    assert.equal(v.ext, 'docx');
    assert.equal(v.locale, 'ko-KR');
    assert.equal(v.available, true);
    assert.equal(v.title, '문서 편집');
    assert.match(v.appUrl, /^http:\/\/127\.0\.0\.\d+:\d+\/office\/[0-9a-f]{16}-[0-9a-f]{32}\/index\.html$/, '판-실행마다 새 열쇠');
    assert.equal(typeof v.userName, 'string');
    assert.ok(v.userName.length > 0, '사용자 이름 (Office 이름 → Windows 계정 이름)');
    assert.equal(v.draft, null);
    assert.equal(v.about.appVersion, '0.0.0-test');
    assert.equal(v.about.bundle.onlyoffice.build, '9.4.0.129');
    assert.equal(v.about.bundle.files, undefined, '파일 목록은 넘기지 않는다');
    assert.equal(shaBytes(v.bytes), docxBefore);
    assert.ok(globalThis.__okFontCalls >= 1, '글꼴 목록을 준비한다');
  });

  await t.test('Office: 편집기 설정(state.set/get)은 settings.json의 officeEditor.settings에', async () => {
    const value = { zoom: 120, rulers: true, name: '보기 설정' };
    let r = await officeCall('editor:state', 'set', 'de-settings', value);
    assert.equal(r.ok, true, r.error);
    r = await officeCall('editor:state', 'get', 'de-settings');
    assert.deepEqual(r.value, value);
    const file = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));
    assert.deepEqual(file.officeEditor.settings['de-settings'], value);
    // 화면(주 창) 설정에는 내놓지 않는다
    const pub = await call('settings:get');
    assert.equal(pub.value.officeEditor, undefined);
    // 지우기
    r = await officeCall('editor:state', 'set', 'de-settings', null);
    assert.equal(r.ok, true, r.error);
    assert.equal((await officeCall('editor:state', 'get', 'de-settings')).value, null);
    // 너무 큰 값·이상한 이름은 거절
    assert.equal((await officeCall('editor:state', 'set', 'big', 'x'.repeat(9000))).code, 'TOO_BIG');
    assert.equal((await officeCall('editor:state', 'set', '../x', 1)).ok, false);
    r = await officeCall('editor:state', 'set', 'de-settings', value);
    assert.equal(r.ok, true);
  });

  const draftFile = draftPathFor(ud, docxPath);
  let editedDocx;
  await t.test('Office 자동 저장(reason auto) → 복구 사본에만, 문서 파일은 그대로', async () => {
    const tmp = join(base, 'edited.docx');
    await createDocument(tmp, '# 회의 메모\n\n편집 창에서 고친 내용입니다.\n\n- 안건 하나\n- 안건 둘\n- 새 안건');
    editedDocx = readFileSync(tmp);
    const changes0 = mainWc.sent.filter((s) => s.channel === 'doc:changed' && s.data.by === 'editor').length;
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.draft, true);
    assert.equal(sha(docxPath), docxBefore, '사용자 파일을 몰래 덮어쓰지 않는다');
    assert.ok(existsSync(draftFile), '복구 사본');
    assert.equal(sha(draftFile), shaBytes(editedDocx));
    assert.equal(dirname(draftFile), join(ud, 'office-drafts'));
    await sleep(100);
    assert.equal(mainWc.sent.filter((s) => s.channel === 'doc:changed' && s.data.by === 'editor').length, changes0, '저장 알림 없음');
  });

  await t.test('Office: 저장 안 한 편집 → 창 제목 *, 주 창 상태, AI 보내기 막기', async () => {
    fake.emit('editor:dirty', officeWin.webContents, true);
    const st = await until(() => lastSent('doc:editor', (d) => d.path === docxPath && d.dirty));
    assert.deepEqual(st, { path: docxPath, open: true, dirty: true });
    assert.equal(officeWin.getTitle(), '* 메모.docx - 문서 편집 (Office Kit)');
    const info = await call('doc:open', docxPath);
    assert.deepEqual(info.value.editor, { open: true, dirty: true });
    const refused = await call('ai:send', { prompt: '표 하나 넣어줘' });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, 'EDITOR_DIRTY');
  });

  await t.test('Office 원래대로: 저장 안 한 편집이 있으면 편집 창에 알린다 (reason restore)', async () => {
    const n0 = sentTo(officeWin, 'editor:external-change').length;
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    const got = sentTo(officeWin, 'editor:external-change');
    assert.equal(got.length, n0 + 1);
    assert.equal(got.at(-1).data.reason, 'restore');
    assert.equal(sha(docxPath), docxBefore);
  });

  await t.test('Office 저장(reason user) → 형식 확인·고치기·백업·쓰기, 복구 사본 지움, 미리보기 알림', async () => {
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'user' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '메모.docx');
    assert.ok(Array.isArray(r.value.fixups));
    assert.equal(r.value.warning, null);
    assert.match(await readDocumentText(docxPath), /편집 창에서 고친 내용/);
    assert.equal(existsSync(draftFile), false, '정상 저장 → 복구 사본은 필요 없다');
    const changed = await until(() => lastSent('doc:changed', (d) => d.by === 'editor' && d.path === docxPath));
    assert.ok(changed);
    const index = JSON.parse(readFileSync(join(ud, 'backups', 'index.json'), 'utf8'));
    assert.ok(index.some((e) => e.source === docxPath && e.hash === docxBefore), '덮어쓰기 전 내용이 백업돼 있다');
    // 형식이 아닌 바이트는 저장하지 않는다
    const bad = await officeCall('editor:save', new Uint8Array(Buffer.from('not a zip')), { reason: 'user' });
    assert.equal(bad.ok, false);
    assert.equal(bad.code, 'BAD_FORMAT');
    fake.emit('editor:dirty', officeWin.webContents, false);
    assert.ok(await until(() => lastSent('doc:editor', (d) => d.path === docxPath && !d.dirty)));
    assert.equal((await call('ai:send', { prompt: '요약해줘' })).ok, true, '저장한 뒤에는 보낸다');
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end' && d.path === docxPath)));
  });

  await t.test('Office 닫기: 저장 안 함 → 창 닫힘, 파일 그대로, 복구 사본 지움 (정상 닫기)', async () => {
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.ok(existsSync(draftFile));
    fake.emit('editor:dirty', officeWin.webContents, true);
    fake.replies.set(officeWin.webContents.id, true);
    const hash = sha(docxPath);
    const d0 = fake.dialogs.length;
    fake.dialogAnswers.push(1); // 저장 안 함
    officeWin.close();
    assert.ok(await until(() => officeWin.isDestroyed(), 8000), '창이 닫혀야 한다');
    assert.equal(fake.dialogs.length, d0 + 1);
    assert.deepEqual(fake.dialogs[d0].buttons, ['저장', '저장 안 함', '취소']);
    assert.equal(fake.dialogs[d0].title, '문서 편집');
    assert.match(fake.dialogs[d0].message, /메모\.docx/);
    assert.equal(sha(docxPath), hash);
    assert.equal(existsSync(draftFile), false, '정상 닫기 → 복구 사본 지움');
    assert.ok(await until(() => lastSent('doc:editor', (d) => d.path === docxPath && !d.open)));
  });

  await t.test('Office 복구 사본: 파일보다 새로우면 다시 열 때 알리고, 읽기·버리기', async () => {
    put(draftFile, editedDocx);
    const future = new Date(Date.now() + 60_000);
    utimesSync(draftFile, future, future);
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    officeWin = fake.windows.at(-1);
    officeCall = (channel, ...args) => fake.invokeFrom(officeWin.webContents, channel, ...args);
    const d = await officeCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.equal(d.value.draft?.size, editedDocx.length);
    const read = await officeCall('editor:draft', 'read');
    assert.equal(read.ok, true, read.error);
    assert.equal(shaBytes(read.value.bytes), shaBytes(editedDocx));
    const gone = await officeCall('editor:draft', 'discard');
    assert.deepEqual(gone.value, { discarded: true });
    assert.equal(existsSync(draftFile), false);
    assert.equal((await officeCall('editor:draft', 'read')).code, 'NO_DRAFT');
  });

  let newEditor;
  await t.test('새 문서 edit:true → 한글은 한글 편집, pptx는 슬라이드 편집, xlsx는 표 편집 창', async () => {
    const hw = await call('doc:new', { format: 'hwpx', folder: docs, name: '새 편집', start: 'blank', edit: true });
    assert.equal(hw.ok, true, hw.error);
    assert.equal(hw.value.editorError, null);
    assert.equal(hw.value.editor.path, join(docs, '새 편집.hwpx'));
    assert.equal(hw.value.editor.kind, 'hwp');
    newEditor = fake.windows.at(-1);
    assert.match(newEditor.getTitle(), /새 편집\.hwpx - 한글 편집/);
    const pp = await call('doc:new', { format: 'pptx', folder: docs, name: '새 발표', start: 'blank', edit: true });
    assert.equal(pp.ok, true, pp.error);
    assert.equal(pp.value.editorError, null);
    assert.equal(pp.value.editor.kind, 'office');
    assert.equal(fake.windows.at(-1).getTitle(), '새 발표.pptx - 슬라이드 편집 (Office Kit)');
    const xl = await call('doc:new', { format: 'xlsx', folder: docs, name: '새 표', start: 'blank', edit: true });
    assert.equal(xl.ok, true, xl.error);
    assert.equal(fake.windows.at(-1).getTitle(), '새 표.xlsx - 표 편집 (Office Kit)');
    // edit 없이 만들면 편집 창을 열지 않는다
    const n = fake.windows.length;
    const wd = await call('doc:new', { format: 'docx', folder: docs, name: '새 워드', start: 'blank' });
    assert.equal(wd.ok, true, wd.error);
    assert.equal(wd.value.editor, null);
    assert.equal(fake.windows.length, n);
  });

  await t.test('Office 편집 창이 4개를 넘으면 메모리 안내 (notice)', async () => {
    // 지금 Office 창: 메모.docx, 새 발표.pptx, 새 표.xlsx
    const notices = [];
    for (const name of ['넷째', '다섯째']) {
      const p = join(docs, `${name}.docx`);
      await createBlankDocument(p);
      assert.equal((await call('doc:open', p)).ok, true);
      const r = await call('doc:edit');
      assert.equal(r.ok, true, r.error);
      notices.push(r.value.notice);
    }
    assert.equal(notices[0], null, '4개까지는 안내 없음');
    assert.match(notices[1], /5개 열려 있습니다/);
    assert.match(notices[1], /메모리/);
  });

  await t.test('Office: 암호 문서·옛 형식은 창을 열기 전에 한국어로 거절', async () => {
    const old = join(docs, '옛 문서.docx');
    const ole = Buffer.alloc(1024);
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(ole);
    writeFileSync(old, ole);
    assert.equal((await call('doc:open', old)).ok, true);
    const n = fake.windows.length;
    let r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'OLD_FORMAT');
    assert.match(r.error, /프로그램에서 여세요/);
    Buffer.from('EncryptionInfo', 'utf16le').copy(ole, 512);
    writeFileSync(old, ole);
    r = await call('doc:edit');
    assert.equal(r.code, 'ENCRYPTED');
    assert.match(r.error, /암호가 걸린 문서는 직접 편집할 수 없습니다/);
    assert.equal(fake.windows.length, n);
  });

  await t.test('한글: 내장 엔진이 못 여는 문서(손상)는 편집 창을 띄우기 전에 거절 — 엔진 원래 글은 detail로', async () => {
    const bad = join(docs, '깨진 문서.hwp');
    const bytes = Buffer.alloc(4096, 0x5a);
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(bytes);
    writeFileSync(bad, bytes);
    assert.equal((await call('doc:open', bad)).ok, true);
    const n = fake.windows.length;
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'BAD_FORMAT');
    assert.match(r.error, /^‘깨진 문서\.hwp’ 파일은 손상됐거나 지원하지 않는 한글 문서라 직접 편집할 수 없어요/);
    assert.ok(r.detail, '엔진이 낸 원래 글');
    assert.equal(fake.windows.length, n, '창을 띄우지 않음');
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
    assert.equal(sha(victim), shaBytes(bytes));
    // 확장자가 맞고 없는 이름이면 묻지 않는다
    const d1 = fake.dialogs.length;
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '다른 이름.hwpx') });
    r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.equal(fake.dialogs.length, d1);
  });

  await t.test('Office 내려받기(PDF 등) → 열 수 있는 형식이면 최근 목록에 (PDF도 이제 열 수 있음)', async () => {
    const w = fake.windows.at(-1); // 다섯째.docx
    const dl = (...a) => fake.invokeFrom(w.webContents, 'editor:download', ...a);
    const odt = join(docs, '내보낸 문서.odt');
    fake.saveDialogAnswers.push({ canceled: false, filePath: odt });
    let r = await dl('내보낸 문서.odt', new Uint8Array(readFileSync(docxPath)), 'application/vnd.oasis.opendocument.text');
    assert.equal(r.ok, true, r.error);
    let recent = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).recent;
    assert.equal(recent[0].path, odt);
    const pdf = join(docs, '내보낸 문서.pdf');
    fake.saveDialogAnswers.push({ canceled: false, filePath: pdf });
    r = await dl('내보낸 문서.pdf', new Uint8Array(Buffer.from('%PDF-1.7\n%%EOF\n')), 'application/pdf');
    assert.equal(r.ok, true, r.error);
    recent = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).recent;
    assert.equal(recent[0].path, pdf);
  });

  await t.test("편집 창의 'AI · 도구' → 주 창이 그 문서의 옆 창이 되어 편집 창 오른쪽에 붙고, 닫으면 보통 주 창으로", async () => {
    officeWin.setBounds({ x: 100, y: 50, width: 1700, height: 900 });
    fake.emit('editor:panel', officeWin.webContents);
    const on = await until(() => lastSent('panel', (d) => d.on === true));
    assert.equal(on.path, docxPath);
    assert.ok(await until(() => lastSent('doc:opened', (d) => d?.path === docxPath)), '옆 창은 그 문서를 연다');
    const pb = mainWin.getBounds();
    const eb = officeWin.getBounds();
    assert.equal(pb.width, 420, '좁은 옆 창');
    assert.equal(pb.x, eb.x + eb.width, '편집 창 바로 오른쪽');
    assert.ok(eb.x + eb.width + pb.width <= 1920, '자리가 없으면 편집 창을 줄인다');
    assert.deepEqual(mainWin.getMinimumSize(), [320, 480]);
    // 처음부터 주 창을 쓰던 경우(Office Kit를 직접 실행): 옆 창을 닫으면 보통 주 창으로 돌아온다 (숨기지 않음)
    assert.equal((await call('panel:close')).ok, true);
    assert.ok(await until(() => lastSent('panel', (d) => d.on === false)));
    assert.equal(mainWin.isVisible(), true);
    assert.deepEqual(mainWin.getMinimumSize(), [1000, 680]);
  });

  await t.test('탐색기에서 연 옛 형식(csv) → 새 형식(xlsx)으로 바꿔 편집 창, 첫 저장은 원본 옆 새 파일 (원본 그대로)', async () => {
    const csv = join(docs, '명단.csv');
    writeFileSync(csv, '이름,점수\n홍길동,90\n김영희,85\n', 'utf8');
    const csvBefore = readFileSync(csv);
    process.env.OFFICEKIT_DOC_WINDOW = '1';
    const before = fake.windows.length;
    try { fake.app.emit('second-instance', {}, [process.execPath, csv], docs); } finally { delete process.env.OFFICEKIT_DOC_WINDOW; process.env.OFFICEKIT_DOC_WINDOW = '0'; }
    assert.ok(await until(() => fake.windows.slice(before).some((w) => /명단\.csv → \.xlsx - 표 편집/.test(w.getTitle())), 20_000), fake.windows.map((w) => w.getTitle()).join(' | '));
    const w = fake.windows.slice(before).find((x) => /명단\.csv/.test(x.getTitle()));
    const c = (ch, ...a) => fake.invokeFrom(w.webContents, ch, ...a);
    const d = await c('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.equal(d.value.ext, 'xlsx');
    assert.equal(d.value.convertedFrom, '명단.csv');
    const temp = (await import('node:path')).dirname(join(ud, 'converted'));
    void temp;
    // 첫 저장: 원본 옆 '명단.xlsx'를 고르게 하고, 그 파일을 편집하는 창이 된다
    const target = join(docs, '명단.xlsx');
    fake.saveDialogAnswers.push({ canceled: false, filePath: target });
    const r = await c('editor:save', new Uint8Array(d.value.bytes), { reason: 'user' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '명단.xlsx');
    assert.ok(existsSync(target));
    assert.deepEqual(readFileSync(csv), csvBefore, '원본 csv는 그대로');
    assert.equal(w.getTitle(), '명단.xlsx - 표 편집 (Office Kit)');
    assert.deepEqual(readdirSync(join(ud, 'converted')), [], '임시 사본은 지운다');
    // 두 번째 저장은 그 파일에 바로
    const r2 = await c('editor:save', new Uint8Array(readFileSync(target)), { reason: 'user' });
    assert.equal(r2.ok, true, r2.error);
    assert.equal(r2.value.name, '명단.xlsx');
    w.close();
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

  await t.test('주 창 닫기: 저장 안 함 → 편집 창 모두 닫고 주 창도 닫힌다 (파일은 그대로)', async () => {
    const hash = sha(docPath);
    const editorsOpen = fake.windows.slice(1).filter((w) => !w.isDestroyed());
    assert.ok(editorsOpen.length >= 6, '한글 2 + Office 5');
    fake.dialogAnswers.push(1); // 저장 안 함
    mainWin.close();
    assert.ok(await until(() => mainWin.isDestroyed(), 8000), '주 창이 닫혀야 한다');
    for (const w of editorsOpen) assert.equal(w.isDestroyed(), true, w.getTitle());
    assert.equal(fake.dialogs.length, closeAsked + 2, '깨끗한 편집 창은 묻지 않는다');
    assert.equal(sha(docPath), hash);
    assert.deepEqual(readdirSync(join(ud, 'office-drafts')).filter((n) => /^[0-9a-f]{40}\./.test(n) && n !== basename(freshDraft)), [], '정상 닫기 → 복구 사본 없음');
  });
});

// ── Office 편집기 설정 보관 (store.mjs) ──────────────────

test('store: officeGet/officeSet — 같은 값이면 다시 쓰지 않고, 한도·이상한 이름을 거른다', () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-office-state-'));
  try {
    const s = createSettings(dir);
    assert.equal(s.officeGet('a'), null);
    assert.equal(s.officeSet('a', { x: 1 }), true);
    assert.deepEqual(s.officeGet('a'), { x: 1 });
    const got = s.officeGet('a');
    got.x = 2;
    assert.deepEqual(s.officeGet('a'), { x: 1 }, '사본을 돌려준다');
    const m0 = statSync(s.file).mtimeMs;
    assert.equal(s.officeSet('a', { x: 1 }), false, '같은 값');
    assert.equal(statSync(s.file).mtimeMs, m0);
    // '__proto__'도 그냥 이름 (프로토타입을 바꾸지 않는다)
    assert.equal(s.officeSet('__proto__', { polluted: true }), true);
    assert.equal(({}).polluted, undefined);
    assert.deepEqual(s.officeGet('__proto__'), { polluted: true });
    assert.throws(() => s.officeSet('a/b', 1), /잘못된 설정 이름/);
    assert.throws(() => s.officeSet('big', 'x'.repeat(9000)), /8KB/);
    // 다시 읽어도 그대로, 한도를 넘으면 오래된 것부터
    for (let i = 0; i < OFFICE_STATE_MAX_KEYS + 5; i++) s.officeSet(`k${i}`, i);
    const again = createSettings(dir);
    assert.equal(again.officeGet('a'), null, '오래된 것은 밀려난다');
    assert.equal(again.officeGet(`k${OFFICE_STATE_MAX_KEYS + 4}`), OFFICE_STATE_MAX_KEYS + 4);
    assert.equal(Object.keys(again.get('officeEditor').settings).length, OFFICE_STATE_MAX_KEYS);
    assert.equal(again.officeSet(`k${OFFICE_STATE_MAX_KEYS + 4}`, null), true, '지우기');
    assert.equal(again.officeGet(`k${OFFICE_STATE_MAX_KEYS + 4}`), null);
    // 망가진 설정 파일도 받아 준다
    writeFileSync(s.file, JSON.stringify({ officeEditor: { settings: { ok: 1, 'bad key': 2, huge: 'y'.repeat(9000) } } }));
    assert.deepEqual(createSettings(dir).get('officeEditor').settings, { ok: 1 });
    writeFileSync(s.file, JSON.stringify({ officeEditor: 'x' }));
    assert.deepEqual(createSettings(dir).get('officeEditor'), { settings: {} });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

// ── '정보 · 라이선스' 글: docs/office-editor-legal-ko.md → src/app/office-editor/legal-ko.json ──

const LEGAL_MD = join(ROOT, 'docs', 'office-editor-legal-ko.md');
const LEGAL_JSON = join(ROOT, 'src', 'app', 'office-editor', 'legal-ko.json');

/** 글 한 줄 → 조각 [{ text, bold?, code?, href?, local? }] (**굵게**, `코드`, [글](https://… | local:경로)만) */
function legalInline(s) {
  const spans = [];
  let last = 0;
  for (const m of s.matchAll(/\*\*(.+?)\*\*|`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\)/g)) {
    if (m.index > last) spans.push({ text: s.slice(last, m.index) });
    if (m[1] !== undefined) spans.push({ text: m[1], bold: true });
    else if (m[2] !== undefined) spans.push({ text: m[2], code: true });
    else if (m[4].startsWith('local:')) spans.push({ text: m[3], local: m[4].slice('local:'.length) });
    else if (m[4].startsWith('https://')) spans.push({ text: m[3], href: m[4] });
    else throw new Error(`허용되지 않는 링크 대상: ${m[4]}`);
    last = m.index + m[0].length;
  }
  if (last < s.length) spans.push({ text: s.slice(last) });
  return spans;
}

/** 법적 고지 글(마크다운 일부 문법) → { title, intro, sections: [{ title, blocks: [{ p } | { ul }] }] } */
function legalFromMarkdown(md) {
  let text = md.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  if (text.trimStart().startsWith('<!--')) text = text.slice(text.indexOf('-->') + 3); // 맨 위 주석(읽는 쪽 안내)은 뺀다
  const out = { title: null, intro: [], sections: [] };
  let blocks = out.intro; let para = null; let list = null;
  const endPara = () => { if (para) { blocks.push({ p: legalInline(para.join(' ')) }); para = null; } };
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    if (!line) { endPara(); list = null; continue; }
    if (/^##\s/.test(line)) {
      endPara(); list = null;
      const s = { title: line.replace(/^##\s+/, ''), blocks: [] };
      out.sections.push(s);
      blocks = s.blocks;
    } else if (/^#\s/.test(line)) {
      endPara(); list = null;
      if (out.title) throw new Error('# 제목은 하나만 쓸 수 있습니다');
      out.title = line.replace(/^#\s+/, '');
    } else if (/^-\s/.test(line)) {
      endPara();
      if (!list) { list = { ul: [] }; blocks.push(list); }
      list.ul.push(legalInline(line.replace(/^-\s+/, '')));
    } else if (/^#{3,}\s|^>|^\d+\.\s|^\|/.test(line)) {
      throw new Error(`이 글에서 쓰지 않는 문법입니다: ${line}`);
    } else {
      list = null;
      (para ??= []).push(line);
    }
  }
  endPara();
  return out;
}

test('정보 · 라이선스 글: legal-ko.json이 docs/office-editor-legal-ko.md와 같다', () => {
  const want = legalFromMarkdown(readFileSync(LEGAL_MD, 'utf8'));
  if (process.env.OFFICEKIT_UPDATE_LEGAL === '1') {
    const json = {
      $comment: 'docs/office-editor-legal-ko.md에서 만든 Office 편집 창 \'정보 · 라이선스\' 글. 고치려면 그 문서를 고친 뒤 OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs. {{appVersion}}은 편집 창이 앱 버전으로 바꾼다.',
      source: 'docs/office-editor-legal-ko.md',
      ...want,
    };
    writeFileSync(LEGAL_JSON, `${JSON.stringify(json, null, 2)}\n`);
  }
  const have = JSON.parse(readFileSync(LEGAL_JSON, 'utf8'));
  assert.deepEqual({ title: have.title, intro: have.intro, sections: have.sections }, want,
    'legal-ko.json이 문서와 다릅니다. OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs 로 다시 만드세요');
  assert.equal(have.source, 'docs/office-editor-legal-ko.md');
});

test('정보 · 라이선스 글: 설계 8장 고지 (원 개발사·수정판·라이선스·소스·상표)와 안전한 링크', () => {
  const legal = JSON.parse(readFileSync(LEGAL_JSON, 'utf8'));
  const secs = legalSections(legal, { appVersion: '0.1.2' });
  assert.ok(secs, '절로 바뀐다');
  const titles = secs.map((s) => s.title);
  for (const t of ['원 개발사', '수정된 판일 수 있습니다', '라이선스', '소스 코드', '상표']) assert.ok(titles.includes(t), t);
  const spans = secs.flatMap((s) => s.blocks.flatMap((b) => b.p ?? b.ul.flat()));
  const all = spans.map((s) => s.text).join('');
  assert.match(all, /Ascensio System SIA/);
  assert.match(all, /수정된 판일 수 있습니다/);
  assert.match(all, /AGPL-3\.0/);
  assert.match(all, /아무런 보증이 없습니다/);
  assert.doesNotMatch(JSON.stringify(secs), /\{\{/, '자리표시는 모두 바뀐다');
  // 소스 꾸러미 주소는 설치된 판의 태그를 가리킨다
  const hrefs = spans.filter((s) => s.href).map((s) => s.href);
  assert.ok(hrefs.includes('https://github.com/toxh13/office-kit/releases/download/v0.1.2/office-editor-source-0.1.2.zip'));
  assert.ok(hrefs.every((h) => h.startsWith('https://')));
  // 묶음 안 고지 파일 링크는 '라이선스 원문' 탭에 있는 파일만
  const locals = spans.filter((s) => s.local).map((s) => s.local);
  assert.ok(locals.length >= 4);
  for (const l of locals) assert.ok(NOTICE_FILES.some((n) => n.path === l), l);
  const office = join(ROOT, 'vendor', 'office-editor');
  if (existsSync(join(office, 'VERSION.json'))) for (const l of locals) assert.ok(existsSync(join(office, ...l.split('/'))), `묶음에 ${l}`);
  // 판을 모르면 릴리스 목록으로
  const unknown = legalSections(legal, {}).flatMap((s) => s.blocks.flatMap((b) => b.p ?? b.ul.flat()));
  assert.ok(unknown.some((s) => s.href === 'https://github.com/toxh13/office-kit/releases'));
  assert.doesNotMatch(JSON.stringify(unknown), /\{\{/);
  // 대화상자 내용: 글이 있으면 그 절을 쓰고, 설치된 판 정보는 그대로
  const m = aboutModel({ appVersion: '0.1.2', bundle: FAKE_VERSION }, 'http://127.0.0.3:5000/office/abc/index.html', legal);
  assert.equal(m.legal, true);
  assert.equal(m.sections[0].title, null, '첫 문단 (기반 소개)');
  assert.ok(m.facts.some((f) => f.label === '편집기'));
  assert.equal(aboutModel({ appVersion: '0.1.2' }, 'x', { sections: [] }).legal, false, '읽을 수 없는 글이면 들어 있는 글로');
  // 나쁜 링크는 글자로만
  const evil = legalSections({ sections: [{ title: 't', blocks: [{ p: [{ text: 'a', href: 'javascript:alert(1)' }, { text: 'b', local: '../../x' }, { text: 'c', html: '<b>' }] }] }] });
  assert.deepEqual(evil[0].blocks[0].p, [{ text: 'a' }, { text: 'b' }, { text: 'c' }]);
});

test('office 서버: 고지 파일(.md·LICENSE·NOTICE)은 글로 내준다 (정보 · 라이선스 보기 틀)', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-office-notice-'));
  const office = join(dir, 'office');
  put(join(office, 'NOTICE-OfficeKit.md'), '# 수정 고지\n');
  put(join(office, 'NOTICE'), 'filex NOTICE');
  put(join(office, 'LICENSE'), 'AGPL');
  put(join(office, 'VERSION.json'), '{}');
  put(join(office, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d]));
  const srv = await startEditorServer({ root: dir, dirs: { office, officeHost: join(ROOT, 'src', 'app', 'office-editor') } });
  try {
    assert.ok(srv.officeBase, srv.officeError);
    for (const f of ['NOTICE-OfficeKit.md', 'NOTICE', 'LICENSE']) {
      const r = await fetch(`${srv.officeBase}${f}`);
      assert.equal(r.status, 200, f);
      assert.equal(r.headers.get('content-type'), 'text/plain; charset=utf-8', f);
      await r.arrayBuffer();
    }
    const w = await fetch(`${srv.officeBase}x2t/x2t.wasm`);
    assert.equal(w.headers.get('content-type'), 'application/wasm');
    await w.arrayBuffer();
    // host 페이지 옆의 legal-ko.json (편집 창이 읽는다)
    const j = await fetch(`${srv.origin}/office-host/legal-ko.json`);
    assert.equal(j.status, 200);
    assert.match(j.headers.get('content-type'), /application\/json/);
    assert.equal((await j.json()).source, 'docs/office-editor-legal-ko.md');
  } finally {
    await srv.close();
    rmSync(dir, { recursive: true, force: true });
  }
});

// ── 실제 Electron 확인용 캡처 (선택) ───────────────────────

const wantSmoke = process.env.OFFICEKIT_EDITOR_SMOKE === '1';
const electronExe = !wantSmoke ? '' : (() => { try { return spawnSync(process.execPath, ['-e', "process.stdout.write(require('electron'))"], { cwd: ROOT, encoding: 'utf8' }).stdout; } catch { return ''; } })();
const hasElectron = !!electronExe && existsSync(electronExe);
const smokeReady = wantSmoke && hasElectron && existsSync(join(ROOT, 'vendor', 'rhwp-studio', 'index.html'));
const officeSmokeReady = wantSmoke && hasElectron && existsSync(join(ROOT, 'vendor', 'office-editor', 'VERSION.json'));

/** 실제 앱을 OFFICEKIT_SMOKE(+_EDITOR)로 띄워 doc을 열고 두 창을 찍는다 → [smoke] 줄의 정보 */
function runSmoke(base, doc, tag) {
  const png = join(base, `${tag}.png`);
  const ud = join(base, `ud-${tag}`);
  // 주 창 미리보기는 글자만 (Word·Excel·PowerPoint를 띄우지 않게)
  put(join(ud, 'settings.json'), JSON.stringify({ previewMode: 'text' }));
  const env = { ...process.env, OFFICEKIT_SMOKE: png, OFFICEKIT_SMOKE_EDITOR: '1', OFFICEKIT_USER_DATA: ud };
  delete env.ELECTRON_RUN_AS_NODE;
  for (const k of ['OFFICEKIT_EDITOR_VENDOR', 'OFFICEKIT_OFFICE_FONTS_MODULE', 'OFFICEKIT_AI_MODULE']) delete env[k];
  const r = spawnSync(electronExe, ['.', doc], { cwd: ROOT, env, encoding: 'utf8', timeout: 200_000 });
  const line = (r.stdout ?? '').split(/\r?\n/).find((l) => l.startsWith('[smoke] '));
  assert.equal(r.status, 0, `${r.stdout}\n${r.stderr}`);
  const info = JSON.parse(line.slice('[smoke] '.length));
  const out = process.env.OFFICEKIT_SMOKE_OUT;
  if (out) {
    mkdirSync(out, { recursive: true });
    copyFileSync(png, join(out, `${tag}.png`));
    const e = info.editor ?? {};
    for (const [f, name] of [[e.png, 'editor'], [e.aboutInfo?.png, 'editor-about'], [e.aboutInfo?.textsPng, 'editor-about-texts']]) {
      if (f && existsSync(f)) copyFileSync(f, join(out, `${tag}-${name}.png`));
    }
    writeFileSync(join(out, `${tag}.json`), JSON.stringify(info, null, 2));
  }
  return { info, png };
}

test('실제 앱: OFFICEKIT_SMOKE_EDITOR로 주 창 + 한글 편집 창 캡처', { skip: !smokeReady && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/ 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async () => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-smoke-'));
  try {
    const doc = join(base, '캡처 시험.hwpx');
    await createBlankDocument(doc, { title: '캡처 시험 문서' });
    const { info, png } = runSmoke(base, doc, 'hwpx');
    assert.equal(info.ok, true);
    assert.equal(info.editable, true, '직접 편집 단추');
    assert.equal(info.editor.ok, true);
    assert.equal(info.editor.state, 'saved');
    assert.equal(info.editor.title, '캡처 시험.hwpx - 한글 편집 (Office Kit)');
    assert.ok(statSync(png).size > 10_000);
    assert.ok(statSync(info.editor.png).size > 10_000);
  } finally {
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
});

// 실제 한국어 문서 (원고 → createDocument)
const OFFICE_SMOKE_DOCS = {
  docx: {
    name: '2026년 4분기 사업 보고서.docx', title: '문서 편집',
    md: [
      '# 2026년 4분기 사업 보고서', '>> 기획팀 김민지 · 2026년 10월 18일', '',
      '## 1. 요약', '4분기 매출은 전년 같은 분기보다 12% 늘었습니다. 새로 연 부산 지점이 늘어난 몫의 절반을 차지했습니다.', '',
      '- 매출: 13억 4,400만 원 (전년 대비 +12%)', '- 새 고객: 182곳 (목표 150곳)', '- 고객 만족도: 4.6점 / 5점', '',
      '## 2. 지점별 실적', '| 지점 | 3분기 | 4분기 | 늘어난 비율 |', '|---|---|---|---|', '| 서울 본점 | 6억 2,000만 | 6억 5,100만 | 5% |',
      '| 부산 지점 | 2억 1,000만 | 3억 4,300만 | 63% |', '| 대전 지점 | 3억 | 3억 5,000만 | 17% |', '',
      '## 3. 다음 분기 계획', '광주 지점 문을 열고, 고객 상담 시간을 저녁 8시까지 늘립니다.',
    ].join('\n'),
  },
  xlsx: {
    name: '부서별 예산표.xlsx', title: '표 편집',
    md: [
      '# 2027년 부서별 예산 (단위: 만 원)', '',
      '| 부서 | 인건비 | 운영비 | 장비 | 합계 |', '|---|---|---|---|---|',
      '| 기획팀 | 42000 | 8000 | 3000 | 53000 |', '| 개발팀 | 96000 | 12000 | 21000 | 129000 |',
      '| 영업팀 | 51000 | 15000 | 2000 | 68000 |', '| 고객지원팀 | 38000 | 6000 | 1500 | 45500 |',
    ].join('\n'),
  },
  pptx: {
    name: '신제품 소개 발표.pptx', title: '슬라이드 편집',
    md: [
      '# 한글 문서 도우미 소개', '>> 2026년 가을 신제품 설명회', '- 인터넷 없이 이 PC 안에서 문서를 고칩니다', '- 한글 · 워드 · 엑셀 · 파워포인트를 한 곳에서', '',
      '# 이런 분께 좋습니다', '- 보고서를 자주 고치는 직장인', '- 표 계산이 많은 회계 담당자', '- 발표 자료를 빨리 만들어야 하는 팀장님',
    ].join('\n'),
  },
};

test('실제 앱: OFFICEKIT_SMOKE_EDITOR로 주 창 + 문서·표·슬라이드 편집 창 캡처 (한국어 문서, 이 PC 글꼴)', { skip: !officeSmokeReady && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/office-editor 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-office-smoke-'));
  t.after(() => rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }));
  for (const [ext, d] of Object.entries(OFFICE_SMOKE_DOCS)) {
    await t.test(ext, async () => {
      const doc = join(base, d.name);
      await createDocument(doc, d.md);
      const before = sha(doc);
      const { info, png } = runSmoke(base, doc, ext);
      assert.equal(info.ok, true, JSON.stringify(info));
      assert.equal(info.editable, true, '직접 편집 단추를 누를 수 있다');
      const e = info.editor;
      assert.equal(e.ok, true, JSON.stringify(e));
      assert.equal(e.kind, 'office');
      assert.equal(e.phase, 'ready', '편집기가 문서를 다 열었다');
      assert.equal(e.title, `${basename(doc)} - ${d.title} (Office Kit)`);
      assert.doesNotMatch(e.title, /ONLYOFFICE/i);
      assert.equal(e.about, '정보 · 라이선스', '정보 · 라이선스 단추');
      // 정보 · 라이선스 대화상자: legal-ko.json 글 + 수정 고지 원문을 글로 보여 준다
      assert.equal(e.aboutInfo?.legal, true, JSON.stringify(e.aboutInfo));
      for (const t of ['원 개발사', '라이선스', '소스 코드', '상표', '설치된 판']) assert.ok(e.aboutInfo.titles.includes(t), t);
      assert.match(e.aboutInfo.text.url, /\/NOTICE-OfficeKit\.md$/);
      assert.equal(e.aboutInfo.text.note, null, '보기 틀에 글로 보인다');
      assert.ok(e.fonts?.count > 50, `이 PC 글꼴 목록 (${JSON.stringify(e.fonts)})`);
      assert.equal(e.fonts.malgun, true, '맑은 고딕이 글꼴 목록에');
      assert.ok(statSync(png).size > 10_000);
      assert.ok(statSync(e.png).size > 30_000);
      assert.equal(sha(doc), before, '열기만 해서는 파일이 바뀌지 않는다');
    });
  }
});
