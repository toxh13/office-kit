// preload.cjs — 화면(renderer)에 window.officeKit 만 노출한다. (sandbox라 CommonJS)
const { contextBridge, ipcRenderer, webUtils } = require('electron');

// 응답은 { ok, value } / { ok:false, error, code } 그대로 넘기고, 화면 쪽에서 풀어 쓴다
const invoke = (channel) => (...args) => ipcRenderer.invoke(channel, ...args);

// 구독 → 해제 함수 반환
const subscribe = (channel) => (callback) => {
  if (typeof callback !== 'function') return () => {};
  const listener = (_event, data) => callback(data);
  ipcRenderer.on(channel, listener);
  return () => ipcRenderer.removeListener(channel, listener);
};

contextBridge.exposeInMainWorld('officeKit', {
  init: invoke('app:init'),

  doc: {
    pick: invoke('doc:pick'),
    open: invoke('doc:open'),
    close: invoke('doc:close'),
    preview: invoke('doc:preview'),
    restore: invoke('doc:restore'),
    openInApp: invoke('doc:openInApp'),
    savePdf: invoke('doc:savePdf'),
    showInFolder: invoke('doc:showInFolder'),
    // 편집 창 열기 (한글 → 한글 편집, docx·xlsx·pptx → 문서·표·슬라이드 편집. 이미 열려 있으면 앞으로)
    // → { path, kind: 'hwp' | 'office', reused, notice }
    edit: invoke('doc:edit'),
    // 새 문서
    newInfo: invoke('doc:newInfo'),
    newCheck: invoke('doc:newCheck'),
    pickFolder: invoke('doc:pickFolder'),
    create: invoke('doc:new'),
  },

  recent: {
    list: invoke('recent:list'),
    remove: invoke('recent:remove'),
    clear: invoke('recent:clear'),
  },

  settings: {
    get: invoke('settings:get'),
    set: invoke('settings:set'),
  },

  ai: {
    providers: invoke('ai:providers'),
    send: invoke('ai:send'),
    cancel: invoke('ai:cancel'),
    reset: invoke('ai:reset'),
  },

  integrations: {
    status: invoke('integrations:status'),
    set: invoke('integrations:set'),
  },

  hancom: {
    status: invoke('hancom:status'),
    setup: invoke('hancom:setup'),
    remove: invoke('hancom:remove'),
  },

  shell: {
    showItem: invoke('shell:showItem'),
    openBackups: invoke('shell:openBackups'),
    openExternal: invoke('shell:openExternal'),
  },

  /** 끌어 놓은 File → 실제 경로 (File.path는 더 이상 없다) */
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file) || ''; } catch { return ''; }
  },

  on: {
    docOpened: subscribe('doc:opened'),
    docChanged: subscribe('doc:changed'),
    /** 편집 창 상태 { path, open, dirty } */
    docEditor: subscribe('doc:editor'),
    aiEvent: subscribe('ai:event'),
    menu: subscribe('menu'),
    toast: subscribe('toast'),
  },

  /** 확인용 캡처 신호 (OFFICEKIT_SMOKE일 때만 main이 듣는다) */
  smokeReady: (info) => ipcRenderer.send('smoke:ready', info && typeof info === 'object' ? { ok: !!info.ok, kind: String(info.kind ?? ''), error: String(info.error ?? ''), editable: info.editable === true } : {}),
});
