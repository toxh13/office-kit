// office-preload.cjs — Office 편집 창(host.html)에 window.officeKitOffice 만 노출한다. (sandbox라 CommonJS, electron 모듈만)
//   host 최상위 틀에만 붙는다. 앱·편집기 틀(127.0.0.3, sandbox)에는 닿지 않는다. 약속: docs/OFFICE-EDITOR-DESIGN.md 6장 IPC.
const { contextBridge, ipcRenderer } = require('electron');

// 바이트는 Uint8Array로만 넘긴다 (그 밖이면 main이 거절)
const toBytes = (v) => {
  if (v instanceof Uint8Array) return v;
  if (v instanceof ArrayBuffer) return new Uint8Array(v);
  return null;
};

// 구독 → 해제 함수 반환
const subscribe = (channel) => (callback) => {
  if (typeof callback !== 'function') return () => {};
  const listener = (_event, data) => { try { callback(data); } catch {} };
  ipcRenderer.on(channel, listener);
  return () => ipcRenderer.removeListener(channel, listener);
};

// main이 묻는 요청({ id, ... })에 콜백 결과로 답한다. 콜백이 없거나 실패하면 null
const answer = (channel, pick) => (callback) => {
  if (typeof callback !== 'function') return () => {};
  const listener = async (_event, req) => {
    let value = null;
    try { value = await callback(pick(req)); } catch { value = null; }
    ipcRenderer.send('editor:reply', { id: req?.id, value: typeof value === 'boolean' ? value : null });
  };
  ipcRenderer.on(channel, listener);
  return () => ipcRenderer.removeListener(channel, listener);
};

// 응답은 { ok, value } / { ok:false, error, code } 그대로 넘기고, 화면 쪽에서 풀어 쓴다
contextBridge.exposeInMainWorld('officeKitOffice', {
  /**
   * 열린 문서 → { kind:'office', name, ext, bytes, size, mtimeMs, readOnly, available, reason,
   *               appUrl, locale, draft: null | { size, mtimeMs }, about: { appVersion, bundle, sources? } }
   */
  getDocument: () => ipcRenderer.invoke('editor:get-document'),
  /** 복구용 사본(자동 저장) 읽기 → { bytes, size, mtimeMs } */
  readDraft: () => ipcRenderer.invoke('editor:draft', 'read'),
  /** 복구용 사본 지우기 → { discarded: true } */
  discardDraft: () => ipcRenderer.invoke('editor:draft', 'discard'),
  /**
   * 저장 (bytes는 원래 형식). reason 'user' = 문서 파일에, 'auto' = 복구용 사본에만
   * → user: { name, ext, size, mtimeMs, backup, warning? } / auto: { draft: true, size, mtimeMs }
   */
  save: (bytes, opts) => ipcRenderer.invoke('editor:save', toBytes(bytes), { reason: opts?.reason === 'auto' ? 'auto' : 'user' }),
  /** 다른 이름으로 저장 (같은 형식만, main의 저장 대화상자) → { canceled } | { name, ext, size, mtimeMs } */
  saveAs: (bytes, ext) => ipcRenderer.invoke('editor:save-as', toBytes(bytes), String(ext ?? '')),
  /** 다른 형식으로 내려받기·PDF (main의 저장 대화상자) → { canceled } | { saved: true, size, path } */
  download: (name, bytes, mime) => ipcRenderer.invoke('editor:download', String(name ?? ''), toBytes(bytes), String(mime ?? '')),
  /** 인쇄 (PDF 바이트 → 기본 PDF 앱) → { opened: true } | { canceled } */
  print: (bytes) => ipcRenderer.invoke('editor:print', toBytes(bytes)),
  /** 편집기 설정 보관: ('get', key) → 값 | null, ('set', key, value) → null. 값마다 8KB까지 */
  state: (op, key, value) => (op === 'set'
    ? ipcRenderer.invoke('editor:state', 'set', String(key ?? ''), value === undefined ? null : value)
    : ipcRenderer.invoke('editor:state', 'get', String(key ?? ''))),
  /** 다른 곳에서 파일이 바뀜 → cb({ missing, size, mtimeMs, reason }) */
  onExternalChange: subscribe('editor:external-change'),
  /** 창을 닫기 전에 main이 묻는다 → cb()가 '저장 안 한 변경이 있는지'(boolean)를 돌려준다 */
  onRequestClose: answer('editor:request-close', () => undefined),
  /** main이 시키는 일 ('save' | 'save-as', 단축키·닫기 확인창) → cb(command)가 성공 여부(boolean)를 돌려준다 */
  onCommand: answer('editor:command', (req) => String(req?.command ?? '')),
  /** 저장 안 한 변경 여부 알리기 (창 제목 *, 주 창 '편집 창 열림·저장 안 됨', 닫기 확인) */
  reportDirty: (dirty) => ipcRenderer.send('editor:dirty', dirty === true),
  /** 진단 로그 (main 콘솔) */
  log: (message) => ipcRenderer.send('editor:log', String(message ?? '').slice(0, 2000)),
  /** 창 닫기 (저장 안 한 변경이 있으면 main이 확인창을 띄운다) */
  close: () => ipcRenderer.send('editor:close'),
  /** 'AI · 도구' 단추: 이 문서의 옆 창(AI 대화·버전·도구)을 연다 */
  panel: () => ipcRenderer.send('editor:panel'),
});
