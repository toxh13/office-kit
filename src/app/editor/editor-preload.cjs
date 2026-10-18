// editor-preload.cjs — 한글 편집 창(host.html)에 window.officeKitEditor 만 노출한다. (sandbox라 CommonJS, electron 모듈만)
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
contextBridge.exposeInMainWorld('officeKitEditor', {
  /** 열린 문서 → { name, ext, bytes, size, mtimeMs, readOnly, studioUrl, available, reason } */
  getDocument: () => ipcRenderer.invoke('editor:get-document'),
  /** 원래 자리에 저장 (bytes는 원래 형식으로 내보낸 것) → { name, ext, size, mtimeMs, backup } */
  save: (bytes) => ipcRenderer.invoke('editor:save', toBytes(bytes)),
  /** 다른 이름으로 저장 (ext: 'hwp' | 'hwpx', bytes는 그 형식) → { canceled } | { name, ext, size, mtimeMs } */
  saveAs: (bytes, ext) => ipcRenderer.invoke('editor:save-as', toBytes(bytes), String(ext ?? '')),
  /** 다른 곳에서 파일이 바뀜 → cb({ missing, size, mtimeMs }) */
  onExternalChange: subscribe('editor:external-change'),
  /** 창을 닫기 전에 main이 묻는다 → cb()가 '저장 안 한 변경이 있는지'(boolean)를 돌려준다 */
  onRequestClose: answer('editor:request-close', () => undefined),
  /** main이 시키는 일 ('save' | 'save-as', 단축키·닫기 확인창) → cb(command)가 성공 여부(boolean)를 돌려준다 */
  onCommand: answer('editor:command', (req) => String(req?.command ?? '')),
  /** 저장 안 한 변경 여부 알리기 (창 제목·닫기 확인에 쓴다) */
  reportDirty: (dirty) => ipcRenderer.send('editor:dirty', dirty === true),
  /** 진단 로그 (main 콘솔) */
  log: (message) => ipcRenderer.send('editor:log', String(message ?? '').slice(0, 2000)),
  /** 창 닫기 (저장 안 한 변경이 있으면 main이 확인창을 띄운다) */
  close: () => ipcRenderer.send('editor:close'),
});
