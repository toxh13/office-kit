// tabstrip-preload.cjs — 탭 줄 화면(renderer/tabstrip.html)에 window.okTabs 만 노출한다. (sandbox라 CommonJS)
const { contextBridge, ipcRenderer } = require('electron');

// 허용 채널 표 (src/app/ipc-channels.mjs). 표에 없는 채널은 불러올 때 바로 던진다
// <ipc-channels> 자동 생성: node scripts/gen-ipc-preload.mjs (손으로 고치지 않는다)
const CHANNELS = { invoke: new Set(['tabs:activate', 'tabs:close', 'tabs:home', 'tabs:togglePanel']), on: new Set(['i18n:get', 'tabs:ready']), send: new Set(['tabs:list']) };
// </ipc-channels>
const need = (call, channel) => {
  if (!CHANNELS[call].has(channel)) throw new Error('IPC 표에 없는 채널: ' + channel);
  return channel;
};
const sendTo = (channel, ...args) => { need('on', channel); ipcRenderer.send(channel, ...args); };
const sendSync = (channel) => { need('on', channel); return ipcRenderer.sendSync(channel); };

// 응답은 { ok, value } / { ok:false, error, code } 그대로 넘긴다
const invoke = (channel) => {
  need('invoke', channel);
  return (...args) => ipcRenderer.invoke(channel, ...args);
};

// 구독 → 해제 함수 반환
const subscribe = (channel) => {
  need('send', channel);
  return (callback) => {
    if (typeof callback !== 'function') return () => {};
    const listener = (_event, data) => callback(data);
    ipcRenderer.on(channel, listener);
    return () => ipcRenderer.removeListener(channel, listener);
  };
};

// 화면 언어와 번역 사전 (처음 그리기 전에 — 한국어면 빈 사전). 표 검사는 try 밖에서 먼저
const I18N_CHANNEL = need('on', 'i18n:get');
const i18n = (() => { try { return sendSync(I18N_CHANNEL) ?? null; } catch { return null; } })();

contextBridge.exposeInMainWorld('okTabs', {
  /** 탭 고르기 (id) */
  activate: invoke('tabs:activate'),
  /** 탭 닫기 (id) — 저장 확인은 main이 한다 */
  close: invoke('tabs:close'),
  /** 홈 탭으로 */
  home: invoke('tabs:home'),
  /** AI 패널 펼치기·접기 */
  togglePanel: invoke('tabs:togglePanel'),
  /** 화면 준비 끝 → main이 탭 목록을 보낸다 */
  ready: () => sendTo('tabs:ready'),
  /** 탭 목록 { tabs: [{ id, type, title, dirty, loaded, zone3, busy }], active, panelCollapsed } */
  onList: subscribe('tabs:list'),
  /** { lang: 'ko'|'en'|'ja', setting, dict } */
  i18n,
});
