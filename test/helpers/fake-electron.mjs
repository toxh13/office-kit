// 시험용 가짜 electron (공용). 시험 두 파일에 따로 있던 가짜의 합집합이다.
// 상태는 makeFakeElectron 호출마다 따로 만든다 (클래스·차례 번호가 시험끼리 섞이지 않게).
import { EventEmitter } from 'node:events';
import { register } from 'node:module';
import { tmpdir } from 'node:os';
import { pathToFileURL } from 'node:url';

export const ELECTRON_EXPORTS = ['app', 'BaseWindow', 'BrowserWindow', 'WebContentsView', 'View', 'Menu', 'dialog', 'ipcMain', 'nativeTheme', 'screen', 'session', 'shell'];

export function makeFakeElectron({ userData, documents, version = '0.0.0-test', commandAnswer = null } = {}) {
  const fake = {
    windows: [], views: [], dialogs: [], dialogAnswers: [], saveDialogs: [], saveDialogAnswers: [], openDialogAnswers: [],
    opened: [], openPathResult: '', external: [], shown: [], sessions: new Map(), replies: new Map(),
    commands: [], commandAnswer, noUserAct: false, menu: undefined, errorBox: undefined, defaultDictionaryUrl: undefined,
  };
  let wcSeq = 0;
  let winSeq = 0;

  const app = Object.assign(new EventEmitter(), {
    isPackaged: false,
    paths: { userData, documents },
    setPath(k, v) { this.paths[k] = v; },
    getPath(k) { return this.paths[k] ?? tmpdir(); },
    requestSingleInstanceLock: () => true,
    quit() { this.quitCalls += 1; },
    exit(code) { this.exitCode = code; },
    whenReady: () => Promise.resolve(),
    getVersion: () => version,
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
      this.destroyed = false;
      this.focused = false;
    }
    send(channel, data) {
      this.sent.push({ channel, data });
      // 편집 창 host 흉내: 닫기 전 질문에 답한다 (fake.replies: webContents id → 저장 안 한 내용 있음?)
      if (channel === 'editor:request-close') {
        setImmediate(() => emit('editor:reply', this, { id: data.id, value: fake.replies.get(this.id) ?? false }));
      }
      // 명령: 늘 기록하고, commandAnswer가 있을 때만 그 값으로 답한다
      if (channel === 'editor:command') {
        fake.commands.push(data.command);
        if (fake.commandAnswer) {
          setImmediate(async () => emit('editor:reply', this, { id: data.id, value: await fake.commandAnswer(this, data.command) }));
        }
      }
    }
    setWindowOpenHandler(fn) { this.openHandler = fn; }
    setWebRTCIPHandlingPolicy(p) { this.webrtcPolicy = p; }
    setBackgroundThrottling(v) { this.backgroundThrottling = v; }
    loadFile(f) { this.mainFrame.url = pathToFileURL(f).href; return Promise.resolve(); }
    loadURL(u) { this.mainFrame.url = u; return Promise.resolve(); }
    toggleDevTools() {}
    executeJavaScript() { return Promise.resolve(null); }
    focus() { this.focused = true; }
    isFocused() { return this.focused; }
    close() {
      if (this.destroyed) return;
      this.destroyed = true;
      this.emit('destroyed');
    }
    isDestroyed() { return this.destroyed; }
  }

  class View extends EventEmitter {
    constructor() {
      super();
      this.children = [];
      this.bounds = { x: 0, y: 0, width: 0, height: 0 };
      this.visible = true;
    }
    addChildView(view, index) {
      const at = this.children.indexOf(view);
      if (at >= 0) this.children.splice(at, 1); // 이미 있으면 빼고 다시 넣는다 = 맨 위
      if (index === undefined) this.children.push(view); else this.children.splice(index, 0, view);
    }
    removeChildView(view) {
      const at = this.children.indexOf(view);
      if (at >= 0) this.children.splice(at, 1);
    }
    setBounds(b) { this.bounds = { ...this.bounds, ...b }; }
    getBounds() { return { ...this.bounds }; }
    setVisible(v) { this.visible = !!v; }
    getVisible() { return this.visible; }
    setBackgroundColor(c) { this.backgroundColor = c; }
  }

  class WebContentsView extends View {
    constructor(opts = {}) {
      super();
      this.opts = opts;
      this.webContents = new WebContents();
      fake.views.push(this);
    }
  }

  const allWindows = new Set();
  let focusedWindow = null;

  class BaseWindow extends EventEmitter {
    static getAllWindows() { return [...allWindows]; }
    static fromId(id) { return [...allWindows].find((w) => w.id === id) ?? null; }
    static getFocusedWindow() { return focusedWindow && !focusedWindow.destroyed ? focusedWindow : null; }
    constructor(opts = {}) {
      super();
      this.opts = opts;
      this.id = ++winSeq;
      this.title = opts.title ?? '';
      this.destroyed = false;
      this.contentView = new View();
      allWindows.add(this);
      fake.windows.push(this);
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
      allWindows.delete(this);
      if (focusedWindow === this) focusedWindow = null;
      this.emit('closed');
    }
    focus() { focusedWindow = this; }
    isFocused() { return focusedWindow === this; }
    show() { this.visible = true; fake.shown.push(this); }
    hide() { this.visible = false; }
    isVisible() { return this.visible !== false; }
    minimize() { this.minimized = true; }
    restore() { this.minimized = false; this.maximized = false; }
    isMinimized() { return !!this.minimized; }
    maximize() { this.maximized = true; }
    unmaximize() { this.maximized = false; }
    isMaximized() { return !!this.maximized; }
    getBounds() { return { ...(this.bounds ?? { x: 0, y: 0, width: 1200, height: 800 }) }; }
    setBounds(b) { this.bounds = { ...this.getBounds(), ...b }; }
    getNormalBounds() { return { ...this.getBounds() }; }
    getContentBounds() { return { ...this.getBounds() }; }
    setMinimumSize(w, h) { this.minSize = [w, h]; }
    getMinimumSize() { return this.minSize ?? [1000, 680]; }
    setTitle(t) { this.title = t; }
    getTitle() { return this.title; }
    moveTop() {}
    removeMenu() {}
  }

  // 실제 Electron과 달리 destroy가 webContents에 'destroyed'를 보내지 않는다 (기존 시험 그대로)
  class BrowserWindow extends BaseWindow {
    static getAllWindows() { return [...allWindows].filter((w) => w instanceof BrowserWindow); }
    constructor(opts = {}) {
      super(opts);
      this.webContents = new WebContents();
      app.emit('browser-window-created', {}, this);
    }
    loadFile(f) { return this.webContents.loadFile(f); }
    loadURL(u) { return this.webContents.loadURL(u); }
  }

  const ipcMain = {
    handlers: new Map(), listeners: new Map(),
    handle(channel, fn) {
      if (this.handlers.has(channel)) throw new Error('이미 등록된 채널: ' + channel);
      this.handlers.set(channel, fn);
    },
    on(channel, fn) {
      if (!this.listeners.has(channel)) this.listeners.set(channel, []);
      this.listeners.get(channel).push(fn);
    },
  };
  const evOf = (wc, frame = wc.mainFrame) => ({ sender: wc, senderFrame: frame });
  function emitAs(ev, channel, ...args) { for (const fn of ipcMain.listeners.get(channel) ?? []) fn(ev, ...args); }
  function emit(channel, wc, ...args) { emitAs(evOf(wc), channel, ...args); }
  function emitSync(wc, channel, ...args) {
    const ev = evOf(wc);
    emitAs(ev, channel, ...args);
    return ev.returnValue;
  }
  async function invokeAs(ev, channel, ...args) {
    const fn = ipcMain.handlers.get(channel);
    if (!fn) throw new Error('없는 채널: ' + channel);
    return fn(ev, ...args);
  }
  // 저장·내려받기·인쇄는 사용자가 누른 것으로 친다 (사용자 조작 없는 요청 거절은 noUserAct로 따로 시험)
  async function invoke(wc, channel, ...args) {
    if (!fake.noUserAct && ['editor:save', 'editor:download', 'editor:print'].includes(channel)) wc.emit?.('input-event', {}, { type: 'mouseDown' });
    return invokeAs(evOf(wc), channel, ...args);
  }
  const invokeFrom = (wc, channel, ...args) => invokeAs(evOf(wc), channel, ...args);

  function keyDown(wc, input) {
    let prevented = false;
    wc.emit('before-input-event', { preventDefault: () => { prevented = true; } }, input);
    return { prevented };
  }
  function crash(wc, reason = 'crashed') { wc.emit('render-process-gone', {}, { reason }); }

  function makeSession(name) {
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
    return s;
  }

  Object.assign(fake, {
    app, BaseWindow, BrowserWindow, WebContentsView, View, ipcMain,
    evOf, emit, emitAs, emitSync, invoke, invokeFrom, invokeAs, keyDown, crash,
    Menu: { buildFromTemplate: (t) => { fake.menu = t; return t; }, setApplicationMenu() {} },
    dialog: {
      async showMessageBox(_w, opts) { fake.dialogs.push(opts); return { response: fake.dialogAnswers.shift() ?? 2 }; },
      async showSaveDialog(_w, opts) { fake.saveDialogs.push(opts); return fake.saveDialogAnswers.shift() ?? { canceled: true }; },
      async showOpenDialog() { return fake.openDialogAnswers.shift() ?? { canceled: true, filePaths: [] }; },
      showErrorBox(title, text) { fake.errorBox = title + ': ' + text; },
    },
    nativeTheme: { shouldUseDarkColors: false, themeSource: 'system' },
    screen: {
      getAllDisplays: () => [{ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }],
      getDisplayMatching: () => ({ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }),
    },
    session: {
      defaultSession: { setPermissionRequestHandler() {}, setPermissionCheckHandler() {}, setSpellCheckerDictionaryDownloadURL(u) { fake.defaultDictionaryUrl = u; } },
      fromPartition(name) {
        if (!fake.sessions.has(name)) fake.sessions.set(name, makeSession(name));
        return fake.sessions.get(name);
      },
    },
    shell: {
      async openPath(p) { fake.opened.push(p); return fake.openPathResult; },
      async openExternal(u) { fake.external.push(u); },
      showItemInFolder() {},
    },
  });
  return fake;
}

/** 'electron' 불러오기를 가짜로 바꿔 끼운다. 한 프로세스에서 두 번째부터는 __okFakeElectron만 바꾼다. */
export function installFakeElectron(fake) {
  globalThis.__okFakeElectron = fake;
  if (globalThis.__okFakeElectronHooked) return;
  globalThis.__okFakeElectronHooked = true;
  const fakeSource = 'const f = globalThis.__okFakeElectron; ' + ELECTRON_EXPORTS.map((n) => 'export const ' + n + ' = f.' + n + ';').join(' ') + ' export default f;';
  register('data:text/javascript,' + encodeURIComponent(
    'export async function resolve(spec, ctx, next) {\n' +
    '  if (spec === \'electron\') return { url: ' + JSON.stringify('data:text/javascript,' + encodeURIComponent(fakeSource)) + ', shortCircuit: true };\n' +
    '  return next(spec, ctx);\n}'));
}
