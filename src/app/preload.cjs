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
  panel: {
    full: invoke('panel:full'),
    close: invoke('panel:close'),
  },

  // PDF 도구: 고르기 → [{ path, name, pages, error }] · 실행 { action: merge|extract|delete|rotate, inputs, pages, degrees } → { output, pages }
  pdftools: {
    pick: invoke('pdftools:pick'),
    run: invoke('pdftools:run'),
  },

  // 기본 프로그램: 상태(형식마다 Office Kit로 여는지) · 등록 고치기 · Windows 기본 앱 설정 열기
  assoc: {
    status: invoke('assoc:status'),
    repair: invoke('assoc:repair'),
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
    /** 옆 창 모드 { on } */
    panel: subscribe('panel'),
  },

  /** 확인용 캡처 신호 (OFFICEKIT_SMOKE일 때만 main이 듣는다) */
  smokeReady: (info) => ipcRenderer.send('smoke:ready', info && typeof info === 'object' ? { ok: !!info.ok, kind: String(info.kind ?? ''), error: String(info.error ?? ''), editable: info.editable === true } : {}),
});
