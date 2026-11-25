// preload.cjs — 화면(renderer)에 window.officeKit 만 노출한다. (sandbox라 CommonJS)
const { contextBridge, ipcRenderer, webUtils } = require('electron');


// 허용 채널 표 (src/app/ipc-channels.mjs). 표에 없는 채널은 불러올 때 바로 던진다
// <ipc-channels> 자동 생성: node scripts/gen-ipc-preload.mjs (손으로 고치지 않는다)
const CHANNELS = { invoke: new Set(['ai:cancel', 'ai:materials', 'ai:providers', 'ai:reset', 'ai:send', 'app:init', 'app:relaunch', 'assoc:openSettings', 'assoc:status', 'batch:cancel', 'batch:pick', 'batch:pickOut', 'batch:saveReport', 'batch:scan', 'batch:start', 'doc:close', 'doc:diff', 'doc:edit', 'doc:new', 'doc:newCheck', 'doc:newInfo', 'doc:open', 'doc:openInApp', 'doc:openInWord', 'doc:pick', 'doc:pickFolder', 'doc:preview', 'doc:print', 'doc:restore', 'doc:savePdf', 'doc:showInFolder', 'doc:versions', 'hancom:remove', 'hancom:setup', 'hancom:status', 'integrations:set', 'integrations:status', 'panel:close', 'panel:full', 'pdftools:pick', 'pdftools:run', 'pii:mask', 'pii:scan', 'recent:clear', 'recent:list', 'recent:pin', 'recent:remove', 'settings:get', 'settings:set', 'shell:openBackups', 'shell:openExternal', 'shell:showItem']), on: new Set(['i18n:get', 'smoke:ready']), send: new Set(['ai:event', 'batch:progress', 'doc:changed', 'doc:editor', 'doc:opened', 'menu', 'panel', 'toast']) };
// </ipc-channels>
const need = (call, channel) => {
  if (!CHANNELS[call].has(channel)) throw new Error('IPC 표에 없는 채널: ' + channel);
  return channel;
};
const sendTo = (channel, ...args) => { need('on', channel); ipcRenderer.send(channel, ...args); };
const sendSync = (channel) => { need('on', channel); return ipcRenderer.sendSync(channel); };

// 응답은 { ok, value } / { ok:false, error, code } 그대로 넘기고, 화면 쪽에서 풀어 쓴다
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

contextBridge.exposeInMainWorld('officeKit', {
  init: invoke('app:init'),
  /** { lang: 'ko'|'en'|'ja', setting: 'system'|…, dict } */
  i18n,
  /** 앱 다시 시작 (언어를 바꾼 뒤) */
  relaunch: invoke('app:relaunch'),

  doc: {
    pick: invoke('doc:pick'),
    open: invoke('doc:open'),
    close: invoke('doc:close'),
    preview: invoke('doc:preview'),
    // 버전으로 되돌리기: 인수 없으면 마지막 변경 전으로, { file }이면 그 버전으로
    restore: invoke('doc:restore'),
    // 버전 목록 → { versions: [{ file, time, reason, size, current }], undo }
    versions: invoke('doc:versions'),
    // 버전 이후 바뀐 글 { file } → { kind: 'text'|'cells', hunks|cells, stats, same, time, reason }
    diff: invoke('doc:diff'),
    openInApp: invoke('doc:openInApp'),
    // PDF를 Word 창에서 열기 (Word가 편집할 수 있는 문서로 바꾼다 — 확인 창은 사용자가 누른다)
    openInWord: invoke('doc:openInWord'),
    savePdf: invoke('doc:savePdf'),
    // 인쇄: 한글은 인쇄 대화상자 { printed }, 그 밖은 PDF로 만들어 PDF 프로그램에서 { opened }
    print: invoke('doc:print'),
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
    /** 고정 · 풀기 (path, on) → 목록 */
    pin: invoke('recent:pin'),
  },

  settings: {
    get: invoke('settings:get'),
    set: invoke('settings:set'),
  },

  ai: {
    providers: invoke('ai:providers'),
    send: invoke('ai:send'),
    cancel: invoke('ai:cancel'),
    materials: invoke('ai:materials'),
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

  // 개인정보: 열린 문서에서 찾기 → { items, total, output(가린 사본 이름) } · 가린 사본 만들기 { values } → { output, items, remaining }
  pii: {
    scan: invoke('pii:scan'),
    mask: invoke('pii:mask'),
  },

  // 한글 → HWPX 일괄 변환: 고르기·세기·시작(진행은 on.batchProgress)·멈추기·보고서 저장
  batch: {
    pick: invoke('batch:pick'),
    pickOut: invoke('batch:pickOut'),
    scan: invoke('batch:scan'),
    start: invoke('batch:start'),
    cancel: invoke('batch:cancel'),
    saveReport: invoke('batch:saveReport'),
  },

  // 옆 창 (편집 창 옆의 AI·도구 창): 전체 창으로 · 닫기
  // M3에서 지움: 탭 창 AI 패널 뷰로 바뀜
  panel: {
    full: invoke('panel:full'),
    close: invoke('panel:close'),
  },

  // PDF 도구: 고르기 → [{ path, name, pages, error }] · 실행 { action: merge|extract|delete|rotate, inputs, pages, degrees } → { output, pages }
  pdftools: {
    pick: invoke('pdftools:pick'),
    run: invoke('pdftools:run'),
  },

  // 기본 프로그램: 상태(형식마다 Office Kit로 여는지) · Windows 기본 앱 설정 열기
  assoc: {
    status: invoke('assoc:status'),
    openSettings: invoke('assoc:openSettings'),
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
    batchProgress: subscribe('batch:progress'),
    // M3에서 지움: 탭 창 AI 패널 뷰로 바뀜
    /** 옆 창 모드 { on } */
    panel: subscribe('panel'),
  },

  /** 확인용 캡처 신호 (OFFICEKIT_SMOKE일 때만 main이 듣는다) */
  smokeReady: (info) => sendTo('smoke:ready', info && typeof info === 'object' ? { ok: !!info.ok, kind: String(info.kind ?? ''), error: String(info.error ?? ''), editable: info.editable === true } : {}),
});
