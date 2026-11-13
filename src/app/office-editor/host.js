// host.js — Office 편집 창(문서·표·슬라이드 편집)의 host 페이지.
//   filex 앱 페이지를 sandbox iframe(다른 사이트, opaque origin)으로 띄우고 filex 앱 SDK v1(MessagePort)에 답한다.
//   약속: docs/OFFICE-EDITOR-DESIGN.md 6장. 편집기 틀 안에는 스크립트를 넣지 않는다 (MessagePort로만 말한다).
//   메시지 모양은 filex 앱 SDK(@brftech/filex-app-ui 0.53, MIT)의 protocol.ts를 따른다. 이 파일은 새로 쓴 것이라 MIT (AGPL 코드 없음).
//   Node에서도 불러올 수 있다 (시험: createBridge·sessionInfo·aboutModel …). 화면 부분은 브라우저에서만 돈다.

// ── 약속 (filex 앱 SDK v1) ──────────────────────────────

export const BRIDGE_VERSION = 1;
/** 앱의 첫 말 (window.parent로) */
export const HELLO = 'filex:hello';
/** host의 답 (포트를 함께 넘김) */
export const PORT = 'filex:port';

export const LIMITS = Object.freeze({
  maxStateBytes: 8 << 10,            // state.set 값 하나 (JSON)
  maxKeyChars: 128,                  // state 키 (영문·숫자·_ . : - 만, main 규칙과 같음)
  maxLineChars: 300,                 // ui.toast · ui.title 글자 수
  maxTextBytes: 32 << 20,            // file.read {as:'text'}
  maxDownloadBytes: 256 << 20,       // ui.download · ui.print
  maxFileBytes: 512 * 1024 * 1024,   // file.save (main 규칙과 같음)
  maxNameChars: 200,                 // 내려받기 이름
});

/** host가 부탁한 저장({hid, request:'save'})을 기다리는 시간 (main은 180초 기다린다) */
export const SAVE_WAIT_MS = 170_000;

export const MIME = Object.freeze({
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
});

/** 형식 → 화면 이름 ('Word·Excel·PowerPoint'는 상표라 쓰지 않는다) */
export const KINDS = Object.freeze({
  docx: { badge: '문서', title: '문서 편집' },
  xlsx: { badge: '표', title: '표 편집' },
  pptx: { badge: '슬라이드', title: '슬라이드 편집' },
});

/** session.get의 grants (설계 6장). 읽기 전용이면 files:write를 뺀다 */
export const GRANTS = Object.freeze([
  'files:read', 'files:write', 'ui', 'ui:eval', 'ui:wasm-eval', 'ui:package-fetch',
  'ui:frame-package', 'ui:connect-blob', 'ui:download', 'ui:print',
]);

/** main이 이름을 주지 못했을 때 (변경 내용 추적·메모 작성자) */
export const USER_NAME = '나';
const TONES = new Set(['info', 'success', 'warning', 'error']);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 앱에 돌려줄 오류 (code = filex ErrorCode) */
function bridgeError(code, message) {
  return Object.assign(new Error(message || code), { bridgeCode: code });
}
/** host 안 오류 (code는 우리 것: not_connected, timeout, closed …) */
function hostError(code, message) {
  return Object.assign(new Error(message || code), { code });
}

/** 바이트로 (ArrayBuffer·TypedArray·Blob·ReadableStream). 모르는 모양이면 null. max를 넘으면 too_large */
export async function toBytes(data, max = Infinity) {
  let u8 = null;
  if (data instanceof ArrayBuffer) u8 = new Uint8Array(data);
  else if (ArrayBuffer.isView(data)) u8 = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  else if (typeof Blob !== 'undefined' && data instanceof Blob) {
    if (data.size > max) throw bridgeError('too_large', '파일이 너무 큽니다');
    u8 = new Uint8Array(await data.arrayBuffer());
  } else if (data && typeof data.getReader === 'function') u8 = await readStream(data, max);
  else return null;
  if (u8.byteLength > max) throw bridgeError('too_large', '파일이 너무 큽니다');
  return u8;
}

async function readStream(stream, max) {
  const reader = stream.getReader();
  const parts = []; let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = value instanceof Uint8Array ? value : ArrayBuffer.isView(value)
      ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength) : value instanceof ArrayBuffer ? new Uint8Array(value) : null;
    if (!chunk) { try { await reader.cancel(); } catch {} throw bridgeError('invalid', '스트림에 바이트가 아닌 것이 들어 있습니다'); }
    total += chunk.byteLength;
    if (total > max) { try { await reader.cancel(); } catch {} throw bridgeError('too_large', '파일이 너무 큽니다'); }
    parts.push(chunk);
  }
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { out.set(p, off); off += p.byteLength; }
  return out;
}

/** 내려받기 이름: 경로·제어 문자를 빼고 길이 제한 */
export function safeFileName(name, fallback = '내려받기') {
  let s = String(name ?? '').replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_').replace(/^[\s.]+|[\s.]+$/g, '');
  if (s.length > LIMITS.maxNameChars) {
    const dot = s.lastIndexOf('.');
    const ext = dot > 0 && s.length - dot <= 10 ? s.slice(dot) : '';
    s = s.slice(0, LIMITS.maxNameChars - ext.length) + ext;
  }
  return s || fallback;
}

/** 목적격 조사: 마지막 한글 글자에 받침이 있으면 '을', 없으면 '를' */
export function objectJosa(word) {
  const c = String(word ?? '').trim().slice(-1).charCodeAt(0);
  return c >= 0xac00 && c <= 0xd7a3 && (c - 0xac00) % 28 !== 0 ? '을' : '를';
}

/** UTF-8 길이 */
const utf8Length = (s) => (typeof TextEncoder !== 'undefined' ? new TextEncoder().encode(s).byteLength : s.length * 3);

/**
 * session.get 답 (설계 6장)
 * doc: { name, ext, size?, bytes?, readOnly?, locale?, userName? }
 *   userName: main이 정한 사용자 이름 (Office에 적힌 이름 → Windows 계정 이름). 변경 내용 추적·메모·마지막 수정한 사람에 들어간다
 */
export function sessionInfo(doc, { mode = 'light', version = '' } = {}) {
  const ext = String(doc?.ext ?? '').toLowerCase();
  const readOnly = doc?.readOnly === true;
  const size = doc?.bytes instanceof Uint8Array ? doc.bytes.byteLength : Number(doc?.size) || 0;
  return {
    v: BRIDGE_VERSION,
    app: { name: 'office-editor', version: String(version ?? '') },
    view: { id: 'editor', placement: 'viewer' },
    locale: typeof doc?.locale === 'string' && doc.locale ? doc.locale : 'ko-KR',
    dir: 'ltr',
    theme: { mode: mode === 'dark' ? 'dark' : 'light', tokens: {} },
    user: { name: userNameOf(doc) },
    files: [{ index: 0, name: String(doc?.name ?? ''), ext, size, mime: MIME[ext] ?? 'application/octet-stream', readOnly }],
    grants: GRANTS.filter((g) => !(readOnly && g === 'files:write')),
  };
}

/** 편집기에 넘길 사용자 이름: 제어 문자 빼고 64자까지, 없으면 USER_NAME */
export function userNameOf(doc) {
  const v = typeof doc?.userName === 'string' ? doc.userName.replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 64) : '';
  return v || USER_NAME;
}

/** 저장 안 한 내용이 있는지 (host가 보는 진짜 상태) */
export function effectiveDirty({ appDirty = false, autoPending = false, restoredDraft = false } = {}) {
  return !!(appDirty || autoPending || restoredDraft);
}

/**
 * host ↔ 앱 다리. 포트 하나(앱 틀 한 번 띄울 때마다 새로)에 답한다.
 * @param {object} o
 * @param {() => object} o.getDoc     지금 문서 { name, ext, bytes: Uint8Array, readOnly, locale }
 * @param {() => 'light'|'dark'} [o.getTheme]
 * @param {string | (() => string)} [o.version]  session.app.version (filex 앱 판)
 * @param {object} o.api              main에 맡기는 일 (값을 돌려주거나 { message, code } 오류를 던진다)
 *   save(bytes, { reason: 'user'|'auto' }) · saveAs(bytes, ext) · download(name, bytes, mime) · print(bytes) · state(op, key, value)
 * @param {object} [o.ui]             화면에 알리기 (모두 고를 수 있음)
 *   onConnect() · onExtraHello() · onDirty(bool) · onSaveStart({as}) · onSaveEnd({ok, result, error, canceled, as})
 *   onAutoSaved(result) · onActionFailed('download'|'print', error) · onToast(text, tone) · onClose()
 * @param {(msg: string) => void} [o.log]
 * @param {number} [o.saveWaitMs]
 */
export function createBridge({ getDoc, getTheme = () => 'light', version = '', api, ui = {}, log = () => {}, saveWaitMs = SAVE_WAIT_MS } = {}) {
  let port = null;
  let hidSeq = 0;
  const hids = new Map();      // hid → { resolve, reject, timer, port }
  let intent = null;           // host가 부탁한 저장: { as, consumed, wrote, result, error, canceled }
  let saving = null;           // requestSave 진행 중 (같은 것을 돌려준다)
  const inflight = new Set();  // 처리 중인 앱 요청 (닫기 전에 잠깐 기다린다)

  const tell = (name, ...a) => {
    try { return ui[name]?.(...a); } catch (err) { log(`화면 처리 오류(${name}): ${err?.message ?? err}`); return undefined; }
  };
  const post = (p, msg, transfer = []) => {
    if (!p) return false;
    try { p.postMessage(msg, transfer); return true; } catch (err) { log(`앱에 보내지 못했습니다: ${err?.message ?? err}`); return false; }
  };
  const reply = (p, id, result, transfer) => post(p, { id, result: result === undefined ? null : result }, transfer);
  const refuse = (p, id, code, message) => post(p, { id, error: message ? { code, message: String(message) } : { code } });

  /** main에 맡긴 일(내려받기·인쇄)이 실패하면 화면에 알리고(한국어) 앱에도 실패로 답한다 */
  const mainCall = async (what, fn) => {
    try { return await fn(); } catch (err) {
      if (err && typeof err === 'object' && !err.bridgeCode) err.bridgeCode = 'failed';
      tell('onActionFailed', what, err);
      throw err;
    }
  };

  const keyOf = (params) => {
    const key = params.key;
    if (typeof key !== 'string' || key.length > LIMITS.maxKeyChars || !/^[\w.:-]+$/.test(key)) throw bridgeError('invalid', '잘못된 키입니다');
    return key;
  };

  // ── 앱 → host 요청 ──
  const METHODS = new Map([
    ['session.get', (p, id) => reply(p, id, sessionInfo(getDoc(), { mode: getTheme(), version: typeof version === 'function' ? version() : version }))],

    ['file.read', (p, id, params) => {
      const doc = getDoc();
      if ((params.index ?? 0) !== 0) return refuse(p, id, 'not_found', '그런 파일이 없습니다');
      if (!(doc?.bytes instanceof Uint8Array)) return refuse(p, id, 'unavailable', '문서를 아직 받지 못했습니다');
      const ext = String(doc.ext ?? '').toLowerCase();
      const head = { name: String(doc.name ?? ''), size: doc.bytes.byteLength, mime: MIME[ext] ?? 'application/octet-stream' };
      const as = params.as ?? 'bytes';
      if (as === 'bytes') {
        const buf = doc.bytes.slice().buffer; // 사본을 넘긴다 (다시 불러오기에 원본을 남김)
        return reply(p, id, { ...head, bytes: buf }, [buf]);
      }
      if (as === 'stream') {
        const copy = doc.bytes.slice();
        const stream = new ReadableStream({ start(c) { c.enqueue(copy); c.close(); } });
        return reply(p, id, { ...head, stream }, [stream]);
      }
      if (as === 'text') {
        if (doc.bytes.byteLength > LIMITS.maxTextBytes) return refuse(p, id, 'too_large', '글로 읽기에는 파일이 너무 큽니다');
        return reply(p, id, { ...head, text: new TextDecoder().decode(doc.bytes) });
      }
      return refuse(p, id, 'invalid', '읽는 방법(as)이 잘못됐습니다');
    }],

    ['file.save', (p, id, params) => fileSave(p, id, params)],

    ['ui.dirty', (p, id, params) => {
      if (p === port) tell('onDirty', params.dirty === true);
      return reply(p, id, null);
    }],

    ['ui.download', async (p, id, params) => {
      const doc = getDoc();
      const stem = String(doc?.name ?? '문서').replace(/\.[^.]+$/, '') || '문서';
      const name = safeFileName(params.name, stem);
      const bytes = typeof params.data === 'string' ? new TextEncoder().encode(params.data) : await toBytes(params.data, LIMITS.maxDownloadBytes);
      if (!bytes) return refuse(p, id, 'invalid', '내려받을 내용이 없습니다');
      if (bytes.byteLength > LIMITS.maxDownloadBytes) return refuse(p, id, 'too_large', '내려받기에는 파일이 너무 큽니다');
      const mime = typeof params.mime === 'string' ? params.mime.slice(0, 200) : '';
      const r = await mainCall('download', () => api.download(name, bytes, mime));
      if (r?.canceled) return refuse(p, id, 'cancelled');
      return reply(p, id, { saved: true, size: bytes.byteLength });
    }],

    ['ui.print', async (p, id, params) => {
      const bytes = await toBytes(params.data, LIMITS.maxDownloadBytes);
      if (!bytes || bytes.byteLength < 5) return refuse(p, id, 'invalid', '인쇄할 내용이 없습니다');
      // PDF만 받는다 (%PDF-)
      if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46 && bytes[4] === 0x2d)) {
        return refuse(p, id, 'invalid', '인쇄할 내용이 PDF가 아닙니다');
      }
      const r = await mainCall('print', () => api.print(bytes));
      if (r?.canceled) return refuse(p, id, 'cancelled');
      return reply(p, id, { printed: true, size: bytes.byteLength });
    }],

    ['state.get', async (p, id, params) => {
      const key = keyOf(params);
      const value = await api.state('get', key);
      return reply(p, id, value === undefined ? null : value);
    }],

    ['state.set', async (p, id, params) => {
      const key = keyOf(params);
      const value = params.value === undefined ? null : params.value;
      let json;
      try { json = JSON.stringify(value); } catch { return refuse(p, id, 'invalid', 'JSON으로 바꿀 수 없는 값입니다'); }
      if (json === undefined) return refuse(p, id, 'invalid', 'JSON으로 바꿀 수 없는 값입니다');
      if (utf8Length(json) > LIMITS.maxStateBytes) return refuse(p, id, 'too_large', '설정 값이 너무 큽니다 (8KB 넘음)');
      await api.state('set', key, JSON.parse(json));
      return reply(p, id, null);
    }],

    ['ui.toast', (p, id, params) => {
      const text = String(params.text ?? '').slice(0, LIMITS.maxLineChars);
      const tone = TONES.has(params.tone) ? params.tone : 'info';
      log(`앱 알림(${tone}): ${text}`);
      if (text && p === port) tell('onToast', text, tone);
      return reply(p, id, null);
    }],

    ['ui.title', (p, id) => reply(p, id, null)], // 창 제목은 main이 정한다

    ['ui.close', (p, id) => {
      reply(p, id, null);
      if (p === port) tell('onClose');
    }],

    ['license.get', (p, id) => reply(p, id, { status: 'free' })],
    // coedit.* · file.saveAs · ui.confirm · clipboard.write · engine.call · job.submit → unknown_method (혼자 편집)
  ]);

  async function fileSave(p, id, params) {
    const doc = getDoc();
    if ((params.index ?? 0) !== 0) return refuse(p, id, 'not_found', '그런 파일이 없습니다');
    // reason은 filex 패치(W2)로 온다. 없으면: host가 부탁한 저장 중이면 'user', 아니면 'auto' (사용자 파일을 몰래 덮어쓰지 않는다)
    const reason = params.reason === 'auto' || params.reason === 'user' ? params.reason : intent && !intent.consumed ? 'user' : 'auto';
    const it = reason === 'user' && intent && !intent.consumed && intent.port === p ? intent : null;
    if (it) it.consumed = true;
    const failWith = (code, message) => {
      if (it) it.error = hostError(code, message);
      return refuse(p, id, code, message);
    };
    if (doc?.readOnly === true) return failWith('read_only', '읽기 전용으로 연 문서라 저장할 수 없습니다');
    let bytes;
    try { bytes = await toBytes(params.data, LIMITS.maxFileBytes); } catch (err) {
      return failWith(err?.bridgeCode ?? 'invalid', err?.bridgeCode === 'too_large' ? '문서가 너무 커서 저장할 수 없습니다 (512MB 넘음)' : err?.message);
    }
    if (!bytes || bytes.byteLength === 0) return failWith('invalid', '저장할 내용이 비어 있습니다');

    if (reason === 'auto') {
      // 복구용 사본만 (userData/office-drafts). 문서 파일은 그대로
      try {
        const r = await api.save(bytes, { reason: 'auto' });
        if (p === port) tell('onAutoSaved', r);
        return reply(p, id, { saved: true, size: bytes.byteLength });
      } catch (err) {
        log(`자동 저장(복구용 사본) 실패: ${err?.message ?? err}`);
        return refuse(p, id, 'failed', err?.message ?? String(err));
      }
    }

    const as = !!it?.as;
    tell('onSaveStart', { as });
    let r;
    try {
      r = as ? await api.saveAs(bytes, String(doc?.ext ?? '').toLowerCase()) : await api.save(bytes, { reason: 'user' });
    } catch (err) {
      if (it) it.error = err;
      tell('onSaveEnd', { ok: false, error: err, as });
      return refuse(p, id, err?.code === 'READ_ONLY' ? 'read_only' : 'failed', err?.message ?? String(err));
    }
    if (r?.canceled) {
      if (it) it.canceled = true;
      tell('onSaveEnd', { ok: false, canceled: true, as });
      return refuse(p, id, 'cancelled', '저장을 취소했습니다');
    }
    if (it) { it.wrote = true; it.result = r; }
    tell('onSaveEnd', { ok: true, result: r, as });
    return reply(p, id, { saved: true, size: Number(r?.size) || bytes.byteLength });
  }

  async function handle(p, m) {
    if (!m || typeof m !== 'object') return;
    if (typeof m.hid === 'number') { settleHid(p, m); return; }
    if (typeof m.id !== 'number' || typeof m.method !== 'string') return;
    const { id, method } = m;
    const params = m.params && typeof m.params === 'object' ? m.params : {};
    const fn = METHODS.get(method);
    if (!fn) { refuse(p, id, 'unknown_method'); return; }
    const job = (async () => {
      try { await fn(p, id, params); } catch (err) {
        if (!err?.bridgeCode) log(`[${method}] ${err?.message ?? err}`);
        refuse(p, id, err?.bridgeCode ?? 'failed', err?.message ?? String(err));
      }
    })();
    inflight.add(job);
    try { await job; } finally { inflight.delete(job); }
  }

  // ── host → 앱 요청 ({hid, request}) ──
  function settleHid(p, m) {
    const w = hids.get(m.hid);
    if (!w || w.port !== p) return;
    hids.delete(m.hid);
    clearTimeout(w.timer);
    if (m.error) {
      const code = String(m.error?.code ?? 'failed');
      w.reject(hostError(code, String(m.error?.message ?? code)));
    } else w.resolve(m.result ?? null);
  }

  function request(name, params, timeoutMs = saveWaitMs) {
    const p = port;
    if (!p) return Promise.reject(hostError('not_connected', '편집기가 아직 준비되지 않았습니다'));
    const hid = ++hidSeq;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { hids.delete(hid); reject(hostError('timeout', '편집기가 응답하지 않습니다')); }, timeoutMs);
      hids.set(hid, { resolve, reject, timer, port: p });
      const msg = params === undefined ? { hid, request: name } : { hid, request: name, params };
      if (!post(p, msg)) { clearTimeout(timer); hids.delete(hid); reject(hostError('closed', '편집기와 연결이 끊겼습니다')); }
    });
  }

  /**
   * 저장 부탁 (Ctrl+S, 막대 '저장', 닫기 확인창의 '저장'): 앱이 file.save로 바이트를 넘기고 {hid, result}로 답한다.
   * as: 다른 이름으로 저장 (같은 형식만, main의 저장 대화상자)
   * → { ok:true, result, as } | { ok:false, canceled:true } | { ok:false, nothing:true } | { ok:false, error, shown }
   *   shown: 저장 실패를 이미 onSaveEnd로 알렸음 (화면이 두 번 알리지 않게)
   */
  function requestSave({ as = false } = {}) {
    if (saving) return saving;
    const it = { as: !!as, consumed: false, wrote: false, result: null, error: null, canceled: false, port };
    intent = it;
    saving = (async () => {
      try {
        await request('save');
      } catch (err) {
        if (!it.wrote && !it.canceled) return { ok: false, error: it.error ?? err, shown: !!it.error };
      } finally {
        if (intent === it) intent = null;
      }
      if (it.canceled) return { ok: false, canceled: true };
      if (it.error) return { ok: false, error: it.error, shown: true };
      if (!it.wrote) return { ok: false, nothing: true };
      return { ok: true, result: it.result, as: it.as };
    })().finally(() => { saving = null; });
    return saving;
  }

  // ── 포트 ──
  function attach(p) {
    reset();
    port = p;
    p.onmessage = (ev) => { void handle(p, ev?.data); };
    try { p.start?.(); } catch {}
  }

  /** 앱 틀이 없어짐 (다시 불러오기·창 닫기): 포트를 닫고 기다리던 것을 끝낸다 */
  function reset() {
    const p = port;
    port = null;
    for (const [hid, w] of hids) {
      clearTimeout(w.timer);
      hids.delete(hid);
      w.reject(hostError('closed', '편집기 화면이 다시 열려 요청이 끊겼습니다'));
    }
    if (p) {
      try { p.onmessage = null; } catch {}
      try { p.close(); } catch {}
    }
  }

  /** 손잡기: 우리가 그린 틀에서 온 hello만 → 포트를 넘긴다. 받았으면 true */
  function accept(ev, frameWin) {
    if (!frameWin || !ev || ev.source !== frameWin) return false;
    const d = ev.data;
    if (!d || typeof d !== 'object' || d.type !== HELLO) return false;
    if (d.v !== BRIDGE_VERSION) { log(`모르는 손잡기 판: ${String(d.v).slice(0, 20)}`); return false; }
    if (port) { tell('onExtraHello'); return false; } // 틀 한 번에 한 번만
    const ch = new MessageChannel();
    attach(ch.port1);
    try {
      frameWin.postMessage({ type: PORT, v: BRIDGE_VERSION }, '*', [ch.port2]); // sandbox 틀은 출처가 "null" → '*'
    } catch (err) {
      log(`포트를 넘기지 못했습니다: ${err?.message ?? err}`);
      reset();
      return false;
    }
    tell('onConnect');
    return true;
  }

  /** host → 앱 이벤트 ({event, data}) */
  function emit(event, data) {
    return post(port, data === undefined ? { event } : { event, data });
  }

  /** 처리 중인 앱 요청이 끝나기를 잠깐 기다린다 (최대 ms) */
  function idle(ms = 300) {
    const jobs = [...inflight];
    if (!jobs.length) return Promise.resolve();
    return Promise.race([Promise.allSettled(jobs), sleep(ms)]).then(() => {});
  }

  /** 창 닫기 직전: 앱이 설정을 보관할 틈을 준다 */
  async function closeRequest({ waitMs = 150, idleMs = 400 } = {}) {
    if (!emit('close.request')) return;
    await sleep(waitMs);
    await idle(idleMs);
  }

  return {
    accept, attach, reset, emit, request, requestSave, idle, closeRequest,
    /** 시험용: 포트 없이 메시지 하나 처리 */
    handle: (m, p = port) => handle(p, m),
    get connected() { return !!port; },
    get saving() { return !!saving; },
  };
}

// ── 정보 · 라이선스 (설계 8장) ─────────────────────────

const GH = 'https://github.com';
// 위쪽 소스 사본 (docs/LICENSES-OFFICE-EDITOR.md 9절)
const UPSTREAM_MIRROR = 'https://github.com/toxh13/office-kit-upstream/releases/tag/office-editor-9.4.0.129';
const OFFICE_KIT_REPO = `${GH}/toxh13/office-kit`;
const short = (s, n = 12) => (typeof s === 'string' && s.length > n ? s.slice(0, n) : s);
const str = (v) => (typeof v === 'string' && v.trim() ? v.trim() : typeof v === 'number' ? String(v) : null);
const httpsUrl = (u) => {
  try { const x = new URL(String(u)); return x.protocol === 'https:' && !x.username && !x.password ? x.href : null; } catch { return null; }
};
// 깃 이름(태그·커밋)으로 쓸 수 있는 글자만
const ref = (s) => (typeof s === 'string' && /^[\w.+-]{1,80}$/.test(s) ? s : null);

/** 묶음과 함께 내주는 고지 파일 (appUrl 폴더 기준). 묶음에서 빼면 안 되는 것들 (설계 8장 1번) */
export const NOTICE_FILES = Object.freeze([
  { path: 'editor/LICENSE.txt', label: 'AGPL 전문과 추가 조항', note: 'ONLYOFFICE 편집기' },
  { path: 'NOTICE-OfficeKit.md', label: 'Office Kit 수정 고지', note: '바꾼 것과 날짜' },
  { path: 'editor/3rd-Party.txt', label: '제3자 구성 요소', note: '편집기 안 오픈 소스' },
  { path: 'THIRD-PARTY-NATIVE.txt', label: '변환기·글꼴 엔진 제3자 구성 요소', note: 'WebAssembly 안 오픈 소스·URW++ 글꼴' },
  { path: 'NOTICE', label: 'filex-office-editor 고지', note: 'BRF Tech' },
  { path: 'editor/filex/CHANGES.txt', label: 'filex 변경 목록', note: '편집기 파일에서 바꾼 것' },
  { path: 'LICENSE', label: 'filex-office-editor 라이선스', note: 'AGPL-3.0' },
]);

/**
 * VERSION.json(main이 about.bundle로 넘김)에서 판 정보를 고른다. 이름이 조금 달라도 읽는다
 * → { onlyoffice: {version, build, sourceTag, image}, filex: {tag, commit}, x2t: {release, core, buildTools}, builtAt } (모르는 것은 null)
 */
export function bundleFacts(bundle) {
  const b = bundle && typeof bundle === 'object' ? bundle : {};
  const oo = b.onlyoffice ?? b.onlyOffice ?? b.editor ?? {};
  const fx = b.filex ?? {};
  const x2 = b.x2t ?? {};
  const build = str(oo.build ?? oo.buildNumber);
  return {
    onlyoffice: {
      version: str(oo.version),
      build,
      sourceTag: ref(str(oo.sourceTag ?? oo.tag)) ?? (build && /^\d+(\.\d+){3}$/.test(build) ? `v${build}` : null),
      image: str(oo.image ?? oo.digest),
    },
    filex: { tag: ref(str(fx.tag ?? fx.version)), commit: ref(str(fx.commit)) },
    x2t: { release: str(x2.release ?? x2.tag), core: ref(str(x2.core ?? x2.coreCommit)), buildTools: ref(str(x2.buildTools ?? x2.build_tools)) },
    builtAt: ((d) => (d && /^\d{4}-\d\d-\d\dT/.test(d) ? d.slice(0, 10) : d))(str(b.builtAt ?? b.generatedAt ?? b.date)),
  };
}

/** legal-ko.json의 자리표시 (앱 버전) */
export const LEGAL_VERSION_SLOT = '{{appVersion}}';

/**
 * '정보 · 라이선스' 글(legal-ko.json — docs/office-editor-legal-ko.md에서 만듦) → 대화상자 절.
 * legal: { title, intro: [블록], sections: [{ title, blocks }] }, 블록 = { p: [조각] } | { ul: [[조각]] },
 *   조각 = { text, bold?, code?, href?(https만), local?(묶음 안 고지 파일) }
 * {{appVersion}}은 앱 버전으로 바꾼다 (모르면 글은 '(알 수 없음)', 주소는 릴리스 목록).
 * → [{ title, blocks }] (읽을 수 없으면 null). 글은 모두 글자로만 그린다 (HTML 없음)
 */
export function legalSections(legal, { appVersion } = {}) {
  if (!legal || typeof legal !== 'object' || !Array.isArray(legal.sections) || !legal.sections.length) return null;
  const ver = ref(str(appVersion));
  const fillText = (s) => String(s).replaceAll(LEGAL_VERSION_SLOT, ver ?? '(알 수 없음)');
  const span = (x) => {
    if (!x || typeof x !== 'object' || typeof x.text !== 'string' || !x.text) return null;
    const text = fillText(x.text).slice(0, 2000);
    if (typeof x.href === 'string') {
      const href = !ver && x.href.includes(LEGAL_VERSION_SLOT) ? `${OFFICE_KIT_REPO}/releases` : httpsUrl(fillText(x.href));
      return href ? { text, href } : { text };
    }
    if (typeof x.local === 'string') return NOTICE_FILES.some((n) => n.path === x.local) ? { text, local: x.local } : { text };
    return { text, ...(x.bold === true ? { bold: true } : {}), ...(x.code === true ? { code: true } : {}) };
  };
  const spans = (list) => (Array.isArray(list) ? list.map(span).filter(Boolean) : []);
  const block = (b) => {
    if (b && Array.isArray(b.p)) { const p = spans(b.p); return p.length ? { p } : null; }
    if (b && Array.isArray(b.ul)) { const ul = b.ul.map(spans).filter((s) => s.length); return ul.length ? { ul } : null; }
    return null;
  };
  const blocks = (list) => (Array.isArray(list) ? list.map(block).filter(Boolean) : []);
  const out = [];
  const intro = blocks(legal.intro);
  if (intro.length) out.push({ title: null, blocks: intro });
  for (const s of legal.sections) {
    if (!s || typeof s !== 'object' || typeof s.title !== 'string') continue;
    const b = blocks(s.blocks);
    if (b.length) out.push({ title: fillText(s.title).slice(0, 200), blocks: b });
  }
  return out.length ? out : null;
}

/**
 * '정보 · 라이선스' 대화상자 내용 (한국어). 화면 없이 만든다 (시험).
 * about: main이 get-document로 넘긴 { appVersion, bundle(VERSION.json 고정 값), sources?: [{label, url}] }
 * appUrl: 앱 페이지 주소 (고지 파일 주소의 기준)
 * legal: legal-ko.json 내용 (있으면 그 글을 절로 쓴다 → sections: [{ title, blocks }], legal: true)
 * → { known, legal, sections: [{ title, paras?, items? } | { title, blocks }], facts: [{label, value}], links: [{label, url}], files: [{label, note, url}] }
 */
export function aboutModel(about, appUrl, legal = null) {
  const a = about && typeof about === 'object' ? about : {};
  const f = bundleFacts(a.bundle);
  const known = !!(f.onlyoffice.version || f.onlyoffice.build);
  const appVersion = ref(str(a.appVersion)) ?? ref(str(a.bundle?.officeKitVersion));
  const oo = f.onlyoffice;
  const ooName = `ONLYOFFICE Docs Community Edition${oo.version ? ` ${oo.version}` : ''}${oo.build ? ` (빌드 ${oo.build})` : ''}`;
  const legalSecs = legalSections(legal, { appVersion });

  const sections = legalSecs ?? [
    {
      title: '원 개발사',
      paras: [
        `이 창의 편집기는 Ascensio System SIA가 만든 ${ooName}를 바탕으로 합니다. 원 개발사는 ONLYOFFICE(Ascensio System SIA, https://www.onlyoffice.com)입니다.`,
        'Office Kit은 이 편집기를 Document Server 없이, 인터넷 없이 이 PC 안에서만 돌립니다.',
      ],
    },
    {
      title: '수정판일 수 있음',
      paras: ['지금 쓰는 편집기는 원본 그대로가 아닌 수정판일 수 있습니다.'],
      items: [
        'filex-office-editor(BRF Tech)가 편집기를 Document Server 없이 돌리려고 일부 파일을 바꿨습니다 (목록: filex 변경 목록).',
        'Office Kit은 쓰지 않는 언어·휴대폰용 화면·묶음 글꼴 파일을 덜어 내고(남은 파일은 고치지 않음), 실행 중에 이 PC에 설치된 글꼴로 글꼴 목록(AllFonts.js)을 새로 만들어 씁니다 (바꾼 것과 날짜: Office Kit 수정 고지).',
      ],
    },
    {
      title: '라이선스',
      items: [
        'ONLYOFFICE 편집기(web-apps·sdkjs)와 변환기(x2t): GNU Affero 일반 공중 사용 허가서 3판(AGPL-3.0)과 그 7조에 따른 추가 조항. 그림·아이콘·문서 글처럼 코드가 아닌 것은 CC BY-SA 4.0.',
        'filex-office-editor: AGPL-3.0-or-later (src/locks.ts만 AGPL-3.0-only).',
        '편집기 안의 다른 오픈 소스(Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO 등): MIT·BSD·Apache-2.0 (목록: 제3자 구성 요소).',
        'Office Kit 앱: MIT. 편집기와는 MessagePort로만 주고받는 별개 프로그램입니다.',
        '글꼴은 들어 있지 않습니다. 이 PC에 설치된 글꼴을 실행 중에 읽기만 합니다.',
      ],
    },
    {
      title: 'ONLYOFFICE 추가 조항 (요약 — 원문은 ‘AGPL 전문과 추가 조항’ 끝)',
      items: [
        '1. 저작권·라이선스·보증 부인·출처 고지를 모두 남겨야 합니다.',
        '2. 수정판은 수정했다는 것과 그 날짜, ONLYOFFICE(Ascensio System SIA) 원본을 바탕으로 했다는 것을 눈에 띄게 적어야 합니다.',
        '3. 대화형 화면은 원 개발사가 ONLYOFFICE임, 수정판일 수 있음, 라이선스 정보를 쉽게 찾을 수 있게 보여 줘야 합니다.',
        '4. 상표·로고·브랜드를 쓸 권리는 주지 않습니다 (https://www.onlyoffice.com/trademark-policy).',
        '5. 그림·아이콘·문서 글처럼 코드가 아닌 것은 CC BY-SA 4.0입니다.',
      ],
    },
    {
      title: '보증 없음',
      paras: ['이 편집기는 관련 법이 허용하는 한도에서 어떤 보증도 없이 제공됩니다 (AGPL-3.0 15·16조).'],
    },
    {
      title: '상표',
      paras: [
        'ONLYOFFICE®와 ONLYOFFICE Docs™는 Ascensio System SIA의 등록 상표 또는 상표입니다. Office Kit은 Ascensio System SIA와 관계가 없고, 그 후원이나 승인을 받지 않았습니다.',
        'Word·Excel·PowerPoint는 Microsoft의 상표이며, 여기서는 파일 형식을 가리킬 때만 씁니다.',
      ],
    },
  ];

  // 설치된 판
  const facts = [];
  if (appVersion) facts.push({ label: 'Office Kit', value: `v${appVersion}` });
  if (known) facts.push({ label: '편집기', value: ooName });
  if (oo.sourceTag) facts.push({ label: '편집기 소스 태그', value: oo.sourceTag });
  if (f.filex.tag || f.filex.commit) facts.push({ label: 'filex-office-editor', value: [f.filex.tag, f.filex.commit && `커밋 ${short(f.filex.commit)}`].filter(Boolean).join(' · ') });
  if (f.x2t.release || f.x2t.core) {
    facts.push({ label: '변환기 x2t', value: [f.x2t.release, f.x2t.core && `core ${short(f.x2t.core)}`, f.x2t.buildTools && `build_tools ${short(f.x2t.buildTools)}`].filter(Boolean).join(' · ') });
  }
  if (oo.image) facts.push({ label: '편집기 파일 출처', value: oo.image });
  if (f.builtAt) facts.push({ label: '묶음 만든 날', value: f.builtAt });

  // 소스 주소 (설치된 판의 태그·커밋)
  const links = [];
  const add = (label, url) => { const u = httpsUrl(url); if (u && !links.some((l) => l.url === u)) links.push({ label, url: u }); };
  for (const s of [...(Array.isArray(a.sources) ? a.sources : []), ...(Array.isArray(a.bundle?.sources) ? a.bundle.sources : [])]) {
    if (s && typeof s === 'object') add(String(s.label ?? s.name ?? '').slice(0, 120) || '소스', s.url);
  }
  add(appVersion ? `Office Kit v${appVersion} 릴리스 (소스 꾸러미 office-editor-source-*.zip)` : 'Office Kit 릴리스 (소스 꾸러미)',
    appVersion ? `${OFFICE_KIT_REPO}/releases/tag/v${appVersion}` : `${OFFICE_KIT_REPO}/releases`);
  // 위쪽 소스 사본: ONLYOFFICE·filex·x2t를 이룬 커밋 그대로의 소스 압축본 (위쪽 태그가 사라져도 받을 수 있게)
  add('위쪽 소스 사본 (ONLYOFFICE·filex·x2t를 이룬 커밋 그대로)', UPSTREAM_MIRROR);
  const tag = oo.sourceTag;
  add(`ONLYOFFICE sdkjs${tag ? ` ${tag}` : ''}`, `${GH}/ONLYOFFICE/sdkjs${tag ? `/tree/${tag}` : ''}`);
  add(`ONLYOFFICE web-apps${tag ? ` ${tag}` : ''}`, `${GH}/ONLYOFFICE/web-apps${tag ? `/tree/${tag}` : ''}`);
  add(`ONLYOFFICE core (x2t)${f.x2t.core ? ` ${short(f.x2t.core)}` : ''}`, `${GH}/ONLYOFFICE/core${f.x2t.core ? `/tree/${f.x2t.core}` : ''}`);
  add(`ONLYOFFICE build_tools${f.x2t.buildTools ? ` ${short(f.x2t.buildTools)}` : ''}`, `${GH}/ONLYOFFICE/build_tools${f.x2t.buildTools ? `/tree/${f.x2t.buildTools}` : ''}`);
  const fxRef = f.filex.commit ?? f.filex.tag;
  add(`filex-office-editor${f.filex.tag ? ` ${f.filex.tag}` : ''}`, `${GH}/BRF-Tech/filex-office-editor${fxRef ? `/tree/${fxRef}` : ''}`);
  add('GNU AGPL 3판 원문', 'https://www.gnu.org/licenses/agpl-3.0.html');

  // 묶음 안 고지 파일 (편집기 출처에서 글로 보여 준다)
  let files = [];
  try {
    const base = new URL('.', appUrl);
    if (base.protocol === 'http:' || base.protocol === 'https:') files = NOTICE_FILES.map((x) => ({ label: x.label, note: x.note, url: new URL(x.path, base).href }));
  } catch {}

  return { known, legal: !!legalSecs, sections, facts, links, files };
}

// ── 화면 (브라우저에서만) ────────────────────────────────

const IS_BROWSER = typeof window !== 'undefined' && typeof document !== 'undefined';
if (IS_BROWSER) startPage();

// 화면 언어: 틀(단추·알림)의 한국어를 사전으로 바꾼다. 편집기 안 메뉴는 편집기 자체 언어 (일본어 메뉴는 없어 영어)
function startI18n(d) {
  const lib = globalThis.OfficeKitI18n;
  if (!d || !lib || !['en', 'ja'].includes(d.lang)) return;
  document.documentElement.lang = d.lang;
  lib.translateDom(document.documentElement, lib.compile(d.dict, d.lang));
}

function startPage() {
  const api = window.officeKitOffice;
  startI18n(api?.i18n);
  const TEST = new URLSearchParams(location.search).get('test') === '1';
  const $ = (id) => document.getElementById(id);
  const HELLO_WAIT_MS = 30_000;
  const SAVED_GRACE_MS = 2500; // 저장 뒤 앱이 '바뀜 없음'을 알려 오기를 기다리는 시간
  const darkQuery = window.matchMedia?.('(prefers-color-scheme: dark)');

  const state = {
    doc: null, failed: null, connected: false,
    appDirty: false, autoPending: false, restoredDraft: false,
    saving: 0, savedAt: 0, reported: null,
    banner: null, frame: null, helloTimer: 0, graceTimer: 0,
    suppressUntil: 0, // 앱이 보낸 저장 실패 알림을 이 시각까지 숨긴다 (host가 이미 띠로 알림)
    reloading: false,
  };

  const log = (message) => { try { api?.log(String(message)); } catch {} };

  /** main 응답 { ok, value } 풀기 */
  function unwrap(res) {
    if (res && res.ok) return res.value;
    const err = new Error(res?.error || '알 수 없는 오류가 났습니다');
    err.code = res?.code ?? null;
    throw err;
  }
  const main = {
    save: async (bytes, opts) => unwrap(await api.save(bytes, opts)),
    saveAs: async (bytes, ext) => unwrap(await api.saveAs(bytes, ext)),
    download: async (name, bytes, mime) => unwrap(await api.download(name, bytes, mime)),
    print: async (bytes) => unwrap(await api.print(bytes)),
    state: async (op, key, value) => unwrap(await api.state(op, key, value)),
  };

  const theme = () => (darkQuery?.matches ? 'dark' : 'light');
  const dirtyNow = () => effectiveDirty(state);

  // ── 화면 조각 ──

  function svgIcon(name) {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', 'i');
    const use = document.createElementNS('http://www.w3.org/2000/svg', 'use');
    use.setAttribute('href', `#i-${name}`);
    svg.append(use);
    return svg;
  }
  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function setDocInfo(doc) {
    const ext = String(doc.ext ?? '').toLowerCase();
    const kind = KINDS[ext] ?? { badge: '문서', title: '문서 편집' };
    $('doc-name').textContent = doc.name ?? '';
    $('doc-name').title = doc.name ?? '';
    $('doc-kind').textContent = kind.badge;
    $('doc-kind').dataset.ext = ext;
    $('doc-readonly').hidden = doc.readOnly !== true;
    document.title = `${doc.name ?? ''} - ${kind.title}`;
  }

  function renderState() {
    const badge = $('doc-state');
    const dirty = dirtyNow();
    let s; let text; let tip = '';
    if (state.failed) { s = 'error'; text = '열 수 없음'; }
    else if (!state.connected) { s = 'loading'; text = '불러오는 중'; }
    else if (state.saving > 0) { s = 'saving'; text = '저장 중…'; }
    else if (dirty && state.appDirty && !state.autoPending && !state.restoredDraft && Date.now() - state.savedAt < SAVED_GRACE_MS) { s = 'saving'; text = '저장 중…'; }
    else if (dirty) {
      s = 'dirty'; text = '저장 안 됨';
      if (state.restoredDraft) tip = '복구한 내용은 아직 파일에 저장되지 않았습니다';
      else if (state.autoPending) tip = '바뀐 내용을 복구용 사본에만 자동 저장했습니다. 파일에 넣으려면 저장하세요';
    } else { s = 'saved'; text = '저장됨'; }
    badge.dataset.state = s;
    badge.textContent = text;
    badge.title = tip;
    const ready = state.connected && !state.failed && state.doc?.readOnly !== true;
    $('btn-save').disabled = !ready || state.saving > 0;
  }

  function reportDirty() {
    const dirty = state.failed ? false : dirtyNow();
    if (state.reported === dirty) return;
    state.reported = dirty;
    try { api?.reportDirty(dirty); } catch {}
  }
  const refresh = () => { renderState(); reportDirty(); };

  /** 알림 띠. actions: [{ label, primary, onClick }] */
  function showBanner(kind, text, actions = [], tone = 'warn') {
    state.banner = kind;
    const banner = $('banner');
    banner.dataset.kind = tone;
    $('banner-text').textContent = text;
    $('banner-actions').replaceChildren(...actions.map((a) => {
      const b = el('button', `btn sm${a.primary ? ' primary' : ''}`, a.label);
      b.type = 'button';
      b.addEventListener('click', () => a.onClick());
      return b;
    }));
    banner.hidden = false;
  }
  function hideBanner(kind) {
    if (kind && state.banner !== kind) return;
    state.banner = null;
    $('banner').hidden = true;
  }
  const closeAction = (kind) => ({ label: '닫기', onClick: () => hideBanner(kind) });

  function toast(text, type = 'success') {
    const t = el('div', `toast ${type}`);
    t.append(svgIcon(type === 'error' ? 'alert' : type === 'info' || type === 'warning' ? 'info' : 'check'), el('span', null, text));
    $('toasts').append(t);
    setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 200); }, type === 'error' ? 5000 : 2800);
  }

  function setLoading(text) {
    $('loading').hidden = !text;
    if (text) $('loading-text').textContent = text;
  }

  /** 편집기를 쓸 수 없을 때 (이유를 한국어로) */
  function fail(reason, detail = '', title = '편집기를 열 수 없습니다') {
    state.failed = reason || title;
    clearTimeout(state.helloTimer);
    setLoading(null);
    $('failure-title').textContent = title;
    $('failure-reason').textContent = reason || '';
    $('failure-detail').textContent = detail || '';
    $('failure-detail').hidden = !detail;
    $('failure').hidden = false;
    refresh();
    log(`실패: ${reason}${detail ? ` (${detail})` : ''}`);
  }

  // ── 다리 ──

  const bridge = createBridge({
    getDoc: () => state.doc,
    getTheme: theme,
    version: () => String(bundleFacts(state.doc?.about?.bundle).filex.tag ?? '').replace(/^v/, ''),
    api: main,
    log,
    ui: {
      onConnect() {
        clearTimeout(state.helloTimer);
        state.connected = true;
        setLoading(null);
        refresh();
      },
      onExtraHello() {
        log('앱 틀이 다시 손잡기를 청함 (틀이 다시 열림)');
        showBanner('rehello', '편집기 화면이 저절로 다시 열렸습니다. 문서를 다시 불러와야 편집을 이어 갈 수 있습니다.', [
          { label: '다시 불러오기', primary: true, onClick: () => { void reloadFromDisk('문서를 다시 불러왔습니다'); } },
        ], 'error');
      },
      onDirty(dirty) {
        state.appDirty = dirty;
        if (!dirty) clearTimeout(state.graceTimer);
        refresh();
      },
      onSaveStart() { state.saving++; refresh(); },
      onSaveEnd({ ok, result, error, canceled, as }) {
        state.saving = Math.max(0, state.saving - 1);
        if (ok) {
          state.autoPending = false;
          state.restoredDraft = false;
          state.savedAt = Date.now();
          clearTimeout(state.graceTimer);
          state.graceTimer = setTimeout(refresh, SAVED_GRACE_MS + 50);
          // 다른 이름으로 저장, 또는 옛 형식에서 바꿔 연 문서의 첫 저장(새 이름) → 이름을 바꿔 보인다
          if (result?.name && (as || result.name !== state.doc?.name)) { state.doc = { ...state.doc, name: result.name, ext: result.ext ?? state.doc.ext, convertedFrom: null }; setDocInfo(state.doc); }
          if (state.banner === 'converted') hideBanner('converted');
          if (state.banner === 'save-error' || state.banner === 'missing' || state.banner === 'changed' || state.banner === 'draft') hideBanner();
          if (result?.warning) showBanner('save-warning', String(result.warning), [closeAction('save-warning')]);
          toast(as ? `다른 이름으로 저장했습니다: ${result?.name ?? ''}` : '저장했습니다');
        } else if (!canceled) {
          state.suppressUntil = Date.now() + 4000;
          showBanner('save-error', `저장하지 못했습니다: ${error?.message ?? error}`, [closeAction('save-error')], 'error');
        }
        refresh();
      },
      onAutoSaved() {
        state.autoPending = true;
        log('자동 저장: 복구용 사본에만 썼습니다');
        refresh();
      },
      onActionFailed(what, err) {
        state.suppressUntil = Date.now() + 4000;
        toast(what === 'print' ? `인쇄하지 못했습니다: ${err?.message ?? err}` : `내려받지 못했습니다: ${err?.message ?? err}`, 'error');
      },
      onToast(text, tone) {
        // 저장 실패는 host가 띠로 알리므로 앱 알림은 로그로만
        if ((tone === 'error' || tone === 'warning') && (bridge.saving || Date.now() < state.suppressUntil)) return;
        toast(text, tone === 'error' ? 'error' : tone === 'success' ? 'success' : 'info');
      },
      onClose() { try { api?.close(); } catch {} },
    },
  });

  window.addEventListener('message', (ev) => {
    const frame = state.frame;
    if (frame && bridge.accept(ev, frame.contentWindow)) log('앱과 연결됨');
  });

  /** 앱 틀을 (다시) 만든다 — 문서 바이트는 state.doc */
  function startApp() {
    bridge.reset();
    state.connected = false;
    state.appDirty = false;
    clearTimeout(state.helloTimer);
    state.frame?.remove();
    const frame = document.createElement('iframe');
    frame.setAttribute('sandbox', 'allow-scripts');
    frame.setAttribute('referrerpolicy', 'no-referrer');
    frame.setAttribute('title', KINDS[String(state.doc.ext ?? '').toLowerCase()]?.title ?? '문서 편집');
    frame.src = state.doc.appUrl;
    state.frame = frame;
    $('app-box').append(frame);
    setLoading('편집기를 준비하는 중…');
    state.helloTimer = setTimeout(() => {
      if (!state.connected) fail('편집기가 시작되지 않았습니다. 편집기 파일이 빠졌거나 손상됐을 수 있으니 Office Kit를 다시 설치하세요.');
    }, HELLO_WAIT_MS);
    refresh();
  }

  async function reloadFromDisk(message) {
    if (state.reloading) return false;
    state.reloading = true;
    try {
      const doc = unwrap(await api.getDocument());
      if (!doc.available) { fail(doc.reason || '편집기 파일이 없습니다.'); return false; }
      state.doc = doc;
      state.autoPending = false;
      state.restoredDraft = false;
      setDocInfo(doc);
      hideBanner();
      try { await bridge.closeRequest({ waitMs: 50, idleMs: 200 }); } catch {}
      startApp();
      toast(message, 'info');
      return true;
    } catch (err) {
      showBanner('reload-error', `다시 불러오지 못했습니다: ${err.message}`, [closeAction('reload-error')], 'error');
      return false;
    } finally {
      state.reloading = false;
      refresh();
    }
  }

  // ── 저장 ──

  async function save({ as = false } = {}) {
    if (!state.connected || state.failed) return false;
    if (state.doc?.readOnly === true) {
      showBanner('save-error', '읽기 전용으로 연 문서는 저장할 수 없습니다.', [closeAction('save-error')], 'error');
      return false;
    }
    state.saving++;
    refresh();
    let r;
    try { r = await bridge.requestSave({ as }); } finally { state.saving = Math.max(0, state.saving - 1); refresh(); }
    if (r.ok) return true;
    if (r.canceled) return false;
    if (r.nothing) {
      showBanner('save-error', '편집기가 저장할 문서를 넘기지 않았습니다. 읽기 전용으로 열렸거나 아직 여는 중일 수 있습니다.', [closeAction('save-error')], 'error');
      return false;
    }
    if (!r.shown) {
      const code = r.error?.code;
      const why = code === 'unavailable' ? '아직 문서를 여는 중입니다. 잠시 뒤 다시 저장하세요.'
        : code === 'timeout' ? '편집기가 응답하지 않습니다.'
          : code === 'closed' || code === 'not_connected' ? '편집기와 연결이 끊겼습니다.'
            : `편집기가 문서를 내보내지 못했습니다 (${r.error?.message ?? code ?? '알 수 없는 오류'})`;
      showBanner('save-error', `저장하지 못했습니다: ${why}`, [closeAction('save-error')], 'error');
    }
    return false;
  }

  // ── 복구용 사본 ──

  function offerDraft(draft) {
    const when = Number.isFinite(draft?.mtimeMs) ? new Date(draft.mtimeMs).toLocaleString(document.documentElement.lang || 'ko', { dateStyle: 'medium', timeStyle: 'short' }) : '';
    showBanner('draft', `자동 저장된 편집 내용이 있습니다${when ? ` (${when})` : ''}. 복구하면 그 내용으로 다시 열고, 저장해야 파일에 들어갑니다.`, [
      { label: '복구', primary: true, onClick: () => { void restoreDraft(); } },
      { label: '버리기', onClick: () => { void discardDraft(); } },
    ]);
  }

  async function restoreDraft() {
    try {
      const d = unwrap(await api.readDraft());
      if (!(d?.bytes instanceof Uint8Array) || !d.bytes.byteLength) throw new Error('복구용 사본이 비어 있습니다');
      if (state.saving) { toast('저장하는 중에는 복구할 수 없습니다', 'info'); return; }
      state.doc = { ...state.doc, bytes: d.bytes, size: d.bytes.byteLength };
      hideBanner('draft');
      startApp();
      state.restoredDraft = true;
      refresh();
      toast('자동 저장된 내용으로 다시 열었습니다. 저장하면 파일에 들어갑니다', 'info');
    } catch (err) {
      showBanner('draft-error', `복구하지 못했습니다: ${err.message}`, [closeAction('draft-error')], 'error');
    }
  }

  async function discardDraft() {
    try {
      unwrap(await api.discardDraft());
      hideBanner('draft');
      toast('자동 저장된 내용을 버렸습니다', 'info');
    } catch (err) {
      showBanner('draft-error', `복구용 사본을 지우지 못했습니다: ${err.message}`, [closeAction('draft-error')], 'error');
    }
  }

  // ── 정보 · 라이선스 ──

  // '정보 · 라이선스' 글 (legal-ko.json, 이 페이지와 같은 폴더). undefined = 아직 안 읽음, null = 못 읽음 → 들어 있는 글로
  // 화면 언어가 한국어가 아니면 영어 글(legal-en.json — 고지문은 조각 번역하지 않는다)
  const LEGAL_FILE = (document.documentElement.lang || 'ko') === 'ko' ? 'legal-ko.json' : 'legal-en.json';
  let legal;
  let legalLoading = null;
  function loadLegal() {
    if (legal !== undefined) return Promise.resolve(legal);
    legalLoading ??= fetch(new URL(LEGAL_FILE, location.href), { cache: 'no-cache' })
      .then((r) => (r.ok ? r.json() : null))
      .catch(() => null)
      .then((v) => { legal = v && typeof v === 'object' ? v : null; if (!legal) log(`정보 · 라이선스 글(${LEGAL_FILE})을 읽지 못해 들어 있는 글로 보여 줍니다`); return legal; });
    return legalLoading;
  }

  /** 글 조각 → 노드 (굵게·코드·바깥 링크·묶음 안 고지 파일 링크) */
  function spanNode(s) {
    if (s.href) {
      const a = el('a', null, s.text);
      a.href = s.href;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.title = s.href;
      return a;
    }
    if (s.local) {
      const b = el('button', 'inline-link', s.text);
      b.type = 'button';
      b.title = '라이선스 원문 탭에서 보기';
      b.addEventListener('click', () => { setTab('texts'); showFile(NOTICE_FILES.findIndex((n) => n.path === s.local)); });
      return b;
    }
    if (s.bold) return el('strong', null, s.text);
    if (s.code) return el('code', null, s.text);
    return document.createTextNode(s.text);
  }
  const spansNode = (tag, spans) => { const e = el(tag); e.append(...spans.map(spanNode)); return e; };

  function renderAbout() {
    const m = aboutModel(state.doc?.about, state.doc?.appUrl, legal);
    const body = $('about-info');
    const parts = [];
    const shown = new Set(); // 글에 이미 있는 주소 (아래 소스 주소 목록에서 겹치지 않게)
    for (const s of m.sections) {
      const sec = el('section', 'about-sec');
      if (s.title) sec.append(el('h3', null, s.title));
      for (const b of s.blocks ?? []) {
        if (b.p) sec.append(spansNode('p', b.p));
        else if (b.ul) { const ul = el('ul'); for (const it of b.ul) ul.append(spansNode('li', it)); sec.append(ul); }
        for (const sp of [...(b.p ?? []), ...(b.ul ?? []).flat()]) if (sp.href) shown.add(sp.href);
      }
      for (const p of s.paras ?? []) sec.append(el('p', null, p));
      if (s.items?.length) {
        const ul = el('ul');
        for (const it of s.items) ul.append(el('li', null, it));
        sec.append(ul);
      }
      sec.dataset.title = s.title ?? '';
      // 이미 화면 언어로 쓴 고지문(legal-en.json)은 번역하지 않는다
      if (m.legal && LEGAL_FILE !== 'legal-ko.json') sec.dataset.noI18n = '';
      parts.push(sec);
    }
    const ver = el('section', 'about-sec');
    ver.append(el('h3', null, m.legal ? '설치된 판' : '설치된 판과 소스 코드'));
    if (!m.known) ver.append(el('p', 'warn', '설치된 편집기 묶음의 판 정보(VERSION.json)를 읽지 못했습니다. 설치가 손상됐을 수 있습니다. 아래 주소에서 소스를 받을 수 있습니다.'));
    if (m.facts.length) {
      const dl = el('dl', 'facts');
      for (const f of m.facts) { dl.append(el('dt', null, f.label), el('dd', null, f.value)); }
      ver.append(dl);
    }
    const links = m.links.filter((l) => !shown.has(l.url));
    if (links.length) {
      ver.append(el('p', null, m.legal
        ? '설치된 묶음(VERSION.json)에 적힌 정확한 소스 주소입니다. 누르면 기본 브라우저로 열지 물어봅니다.'
        : '이 판의 소스 코드(AGPL 6조)는 아래 주소에 있습니다. 누르면 기본 브라우저로 열지 물어봅니다.'));
      const ul = el('ul', 'links');
      for (const l of links) {
        const li = el('li');
        const a = el('a', null, l.label);
        a.href = l.url;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        li.append(a, el('span', 'url', l.url));
        ul.append(li);
      }
      ver.append(ul);
    }
    // 글이 있으면 '소스 코드' 절 다음에, 없으면 '수정판일 수 있음' 다음에
    const at = m.legal ? parts.findIndex((p) => p.dataset.title === '소스 코드') + 1 : 2;
    parts.splice(at > 0 ? at : parts.length, 0, ver);
    body.replaceChildren(...parts);

    const list = $('about-files');
    list.replaceChildren(...m.files.map((f, i) => {
      const b = el('button', 'file-item');
      b.type = 'button';
      b.dataset.url = f.url;
      b.append(el('span', null, f.label), el('small', null, f.note));
      b.addEventListener('click', () => showFile(i));
      return b;
    }));
    $('about-files-empty').hidden = m.files.length > 0;
  }

  /** 틀이 한 번 다 뜰 때까지 (최대 ms) → 떴으면 true */
  const frameLoad = (frame, ms) => new Promise((resolve) => {
    const on = () => done(true);
    const t = setTimeout(() => done(false), ms);
    function done(v) { clearTimeout(t); frame.removeEventListener('load', on); resolve(v); }
    frame.addEventListener('load', on);
  });

  async function showFile(i) {
    const items = [...$('about-files').querySelectorAll('.file-item')];
    items.forEach((b, j) => b.setAttribute('aria-current', j === i ? 'true' : 'false'));
    const item = items[i];
    const url = item?.dataset.url;
    const viewer = $('about-viewer');
    const note = $('about-viewer-note');
    if (!url || viewer.dataset.url === url) return;
    // 먼저 비운다: 글로 내주지 않는 파일이면(내려받기로 막힘) 앞 파일 내용이 남아 다른 파일처럼 보이지 않게
    viewer.dataset.url = url;
    note.hidden = true;
    const blank = frameLoad(viewer, 500);
    viewer.setAttribute('src', 'about:blank');
    await blank;
    if (viewer.dataset.url !== url) return;
    const loaded = frameLoad(viewer, 4000);
    viewer.setAttribute('src', url);
    if ((await loaded) || viewer.dataset.url !== url) return;
    // 여기서 보여 주지 못함 → 설치 폴더의 파일 위치를 알려 준다
    const rel = decodeURIComponent(new URL(url).pathname.replace(/^\/office\/[^/]+\//, '')).replaceAll('/', '\\');
    const label = item.querySelector('span')?.textContent ?? '';
    note.textContent = `‘${label}’${objectJosa(label)} 여기서 보여 주지 못했습니다. Office Kit 설치 폴더의 vendor\\office-editor\\${rel} 파일에서 볼 수 있습니다.`;
    note.hidden = false;
  }

  function setTab(name) {
    for (const t of document.querySelectorAll('#about [role="tab"]')) {
      const on = t.dataset.tab === name;
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      t.tabIndex = on ? 0 : -1;
    }
    $('about-info').hidden = name !== 'info';
    $('about-texts').hidden = name !== 'texts';
    if (name === 'texts' && !$('about-viewer').dataset.url) showFile(0);
  }

  async function openAbout(tab = 'info') {
    // 글을 아직 못 받았으면 잠깐만 기다린다 (못 받으면 들어 있는 글로)
    await Promise.race([loadLegal(), sleep(1500)]);
    renderAbout();
    setTab(tab);
    const dlg = $('about');
    if (!dlg.open) dlg.showModal();
    $('about-info').scrollTop = 0;
  }

  // ── 버튼 · 단축키 ──

  function wireUi() {
    $('btn-save').addEventListener('click', () => { void save(); });
    $('btn-about').addEventListener('click', () => openAbout());
    $('btn-panel')?.addEventListener('click', () => api?.panel?.());
    $('about-close').addEventListener('click', () => $('about').close());
    $('btn-failure-close').addEventListener('click', () => api?.close());
    $('btn-failure-about').addEventListener('click', () => openAbout());
    for (const t of document.querySelectorAll('#about [role="tab"]')) t.addEventListener('click', () => setTab(t.dataset.tab));
    // 탭: 왼쪽·오른쪽 화살표로 옮겨 가기
    document.querySelector('#about [role="tablist"]').addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      const tabs = [...document.querySelectorAll('#about [role="tab"]')];
      const i = tabs.findIndex((t) => t.getAttribute('aria-selected') === 'true');
      const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
      setTab(next.dataset.tab);
      next.focus();
      e.preventDefault();
    });
    $('about').addEventListener('click', (e) => { if (e.target === $('about')) $('about').close(); }); // 바깥(배경) 누름
    darkQuery?.addEventListener?.('change', () => { bridge.emit('theme', { mode: theme(), tokens: {} }); });
  }

  // ── 앱(main)과 잇기 ──

  function wireApp() {
    // 다른 곳(다른 프로그램·AI·주 창의 원래대로)에서 파일이 바뀜
    api.onExternalChange(async (info) => {
      if (!state.doc || state.failed) return;
      if (info?.missing) {
        showBanner('missing', '문서 파일이 없어졌습니다 (옮겨졌거나 지워짐). 저장하면 같은 자리에 다시 만듭니다.', [closeAction('missing')]);
        return;
      }
      if (state.banner === 'missing') hideBanner('missing');
      while (state.saving > 0) await sleep(100);
      const restore = info?.reason === 'restore';
      if (!dirtyNow()) { await reloadFromDisk(restore ? '원래대로 되돌린 내용을 다시 불러왔습니다' : '다른 곳에서 바뀐 내용을 다시 불러왔습니다'); return; }
      showBanner('changed', restore
        ? '주 창에서 문서를 원래대로 되돌렸습니다. 다시 불러오면 여기서 고친 내용은 사라지고, 내 편집을 유지한 채 저장하면 되돌린 내용을 덮어씁니다.'
        : '다른 곳에서 이 문서가 바뀌었습니다. 다시 불러오면 여기서 고친 내용은 사라지고, 내 편집을 유지한 채 저장하면 바깥에서 바뀐 내용을 덮어씁니다.', [
        { label: '다시 불러오기', primary: true, onClick: () => { void reloadFromDisk(restore ? '원래대로 되돌린 내용을 다시 불러왔습니다' : '바뀐 내용을 다시 불러왔습니다'); } },
        { label: '내 편집 유지', onClick: () => hideBanner('changed') },
      ]);
    });

    // 창을 닫기 전: 앱에 알리고(설정 보관), 저장 안 한 내용이 있는지 답한다
    api.onRequestClose(async () => {
      if (state.failed) return false;
      try { await bridge.closeRequest(); } catch {}
      while (state.saving > 0) await sleep(100);
      return dirtyNow();
    });

    // 단축키(Ctrl+S, Ctrl+Shift+S)와 닫기 확인창의 '저장'
    api.onCommand(async (command) => {
      if (command === 'save') return save();
      if (command === 'save-as') return save({ as: true });
      return false;
    });
  }

  // ── 시작 ──

  async function boot() {
    wireUi();
    refresh();
    void loadLegal(); // '정보 · 라이선스' 글을 미리 받아 둔다
    if (!api) { fail('편집 창을 앱과 연결하지 못했습니다.'); return; }
    wireApp();
    let doc;
    try { doc = unwrap(await api.getDocument()); } catch (err) { fail(err.message); return; }
    state.doc = doc;
    setDocInfo(doc);
    if (!doc.available) { fail(doc.reason || '편집기 파일이 없습니다.'); return; }
    if (!(doc.bytes instanceof Uint8Array) || typeof doc.appUrl !== 'string') { fail('문서를 받지 못했습니다.'); return; }
    if (doc.draft) offerDraft(doc.draft);
    // 옛 형식(doc·xls·ppt …)을 바꿔 연 문서: 저장하면 새 형식으로 따로 저장된다고 알린다
    else if (doc.convertedFrom) {
      showBanner('converted', `‘${doc.convertedFrom}’을(를) 고칠 수 있게 .${doc.ext} 형식으로 바꿔 열었어요. 저장하면 원본 옆에 새 .${doc.ext} 파일로 따로 저장되고, 원본은 그대로예요.`, [closeAction('converted')], 'info');
    }
    // 글꼴 목록을 만들지 못함 (편집은 됨): 복구 띠가 있으면 알림으로만
    if (typeof doc.fontsWarning === 'string' && doc.fontsWarning) {
      if (state.banner) toast(doc.fontsWarning, 'info');
      else showBanner('fonts', doc.fontsWarning, [closeAction('fonts')]);
    }
    startApp();
  }

  if (TEST) {
    window.__officeTest = {
      state: () => ({
        connected: state.connected, failed: state.failed, dirty: dirtyNow(), appDirty: state.appDirty,
        autoPending: state.autoPending, restoredDraft: state.restoredDraft, saving: state.saving,
        name: state.doc?.name ?? null, ext: state.doc?.ext ?? null, readOnly: state.doc?.readOnly === true,
        banner: state.banner, bannerText: state.banner ? $('banner-text').textContent : null,
        badge: $('doc-state').textContent, title: document.title, aboutOpen: $('about').open,
      }),
      save: (as) => save({ as: !!as }),
      emit: (event, data) => bridge.emit(event, data),
      openAbout: (tab) => openAbout(tab),
      closeAbout: () => $('about').close(),
      bannerAction: (label) => {
        const b = [...$('banner-actions').querySelectorAll('button')].find((x) => x.textContent === label);
        b?.click();
        return !!b;
      },
      frameRect: () => {
        const r = state.frame?.getBoundingClientRect();
        return r ? { x: r.x, y: r.y, width: r.width, height: r.height } : null;
      },
    };
  }

  boot().catch((err) => fail('편집 창을 준비하지 못했습니다.', err?.message));
}
