// office-editor host 단위 시험 — filex 앱 SDK v1(MessagePort) 약속을 가짜 포트로, preload는 가짜 electron으로
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { ROOT } from './helpers/fixtures.mjs';
import {
  BRIDGE_VERSION, GRANTS, HELLO, KINDS, LIMITS, MIME, NOTICE_FILES, PORT,
  aboutModel, bundleFacts, createBridge, effectiveDirty, objectJosa, safeFileName, sessionInfo, toBytes,
} from '../src/app/office-editor/host.js';

const HOST_DIR = join(ROOT, 'src', 'app', 'office-editor');
const DOCX = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6]);
const PDF = new TextEncoder().encode('%PDF-1.7\n%…\n');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 가짜 MessagePort: 보낸 것을 모으고, 기다리는 사람에게 알린다 */
class FakePort {
  constructor() { this.sent = []; this.closed = false; this.onmessage = null; this.waiters = []; this.started = false; }
  postMessage(msg, transfer = []) {
    if (this.closed) throw new Error('port closed');
    this.sent.push({ msg, transfer });
    for (const w of [...this.waiters]) if (w.pred(msg)) { this.waiters.splice(this.waiters.indexOf(w), 1); w.resolve({ msg, transfer }); }
  }
  start() { this.started = true; }
  close() { this.closed = true; }
  /** 앱이 보낸 것처럼 */
  deliver(msg) { this.onmessage?.({ data: msg }); }
  waitFor(pred, ms = 2000) {
    const hit = this.sent.find((s) => pred(s.msg));
    if (hit) return Promise.resolve(hit);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { const i = this.waiters.indexOf(w); if (i >= 0) { this.waiters.splice(i, 1); reject(new Error('기다린 메시지가 오지 않음')); } }, ms);
      const w = { pred, resolve: (v) => { clearTimeout(timer); resolve(v); } };
      this.waiters.push(w);
    });
  }
}

let nextId = 0;
/** 앱 → host 요청 하나 보내고 답을 기다린다 → { msg, transfer } */
function ask(port, method, params) {
  const id = ++nextId;
  port.deliver(params === undefined ? { id, method } : { id, method, params });
  return port.waitFor((m) => m.id === id);
}

/** main 흉내: 부른 것을 적어 두고 정해 둔 답을 준다 */
function fakeMain(over = {}) {
  const calls = [];
  const settings = new Map();
  const api = {
    save: async (bytes, opts) => { calls.push(['save', bytes, opts]); return over.save ? over.save(bytes, opts) : opts?.reason === 'auto' ? { draft: true, size: bytes.length, mtimeMs: 1 } : { name: 'a.docx', ext: 'docx', size: bytes.length, mtimeMs: 2, backup: null }; },
    saveAs: async (bytes, ext) => { calls.push(['saveAs', bytes, ext]); return over.saveAs ? over.saveAs(bytes, ext) : { name: 'b.docx', ext, size: bytes.length, mtimeMs: 3 }; },
    download: async (name, bytes, mime) => { calls.push(['download', name, bytes, mime]); return over.download ? over.download(name, bytes, mime) : { saved: true, size: bytes.length, path: `C:\\x\\${name}` }; },
    print: async (bytes) => { calls.push(['print', bytes]); return over.print ? over.print(bytes) : { opened: true }; },
    state: async (op, key, value) => {
      calls.push(['state', op, key, value]);
      if (op === 'set') { settings.set(key, value); return null; }
      return settings.has(key) ? settings.get(key) : null;
    },
  };
  return { api, calls, settings };
}

function setup({ doc: docOver = {}, main = {}, theme = 'light', saveWaitMs } = {}) {
  const doc = { name: '보고서.docx', ext: 'docx', bytes: DOCX.slice(), readOnly: false, locale: 'ko-KR', ...docOver };
  const m = fakeMain(main);
  const events = [];
  const logs = [];
  const ui = new Proxy({}, { get: (_t, name) => (...a) => { events.push([name, ...a]); } });
  const bridge = createBridge({ getDoc: () => doc, getTheme: () => theme, version: '0.3.1', api: m.api, ui, log: (s) => logs.push(s), saveWaitMs });
  const port = new FakePort();
  bridge.attach(port);
  return { bridge, port, doc, main: m, events, logs };
}

const ev = (events, name) => events.filter((e) => e[0] === name);

// ── 손잡기 ─────────────────────────────────────────────

test('손잡기: 우리가 그린 틀의 hello만 받고, 포트를 * 로 넘긴다 (틀 한 번에 한 번)', () => {
  const m = fakeMain();
  const events = [];
  const bridge = createBridge({ getDoc: () => ({ name: 'a.docx', ext: 'docx', bytes: DOCX }), api: m.api, ui: { onConnect: () => events.push('connect'), onExtraHello: () => events.push('extra') } });
  const posted = [];
  const frameWin = { postMessage: (msg, target, transfer) => posted.push({ msg, target, transfer }) };
  const otherWin = { postMessage: () => { throw new Error('다른 틀에 보내면 안 됨'); } };

  assert.equal(bridge.accept({ source: otherWin, data: { type: HELLO, v: 1 } }, frameWin), false, '다른 틀');
  assert.equal(bridge.accept({ source: frameWin, data: { type: 'filex-office-editor:hello', v: 1 } }, frameWin), false, '다른 종류');
  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: 2 } }, frameWin), false, '모르는 판');
  assert.equal(bridge.accept({ source: frameWin, data: 'filex:hello' }, frameWin), false, '문자열');
  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: 1 } }, null), false, '틀 없음');
  assert.equal(posted.length, 0);
  assert.equal(bridge.connected, false);

  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: BRIDGE_VERSION } }, frameWin), true);
  assert.equal(posted.length, 1);
  assert.deepEqual(posted[0].msg, { type: PORT, v: 1 });
  assert.equal(posted[0].target, '*', 'sandbox 틀은 출처가 null이라 *');
  assert.equal(posted[0].transfer.length, 1);
  assert.equal(typeof posted[0].transfer[0].postMessage, 'function');
  assert.equal(bridge.connected, true);
  assert.deepEqual(events, ['connect']);

  // 같은 틀의 두 번째 hello는 무시 (화면에만 알림)
  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: 1 } }, frameWin), false);
  assert.equal(posted.length, 1);
  assert.deepEqual(events, ['connect', 'extra']);

  // 틀을 새로 만들면(reset) 다시 받는다
  bridge.reset();
  assert.equal(bridge.connected, false);
  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: 1 } }, frameWin), true);
  bridge.reset();
  try { posted[0].transfer[0].close(); posted[1].transfer[0].close(); } catch {}
});

test('손잡기: 실제 MessageChannel로 넘긴 포트에서 session.get이 답한다', async () => {
  const m = fakeMain();
  const bridge = createBridge({ getDoc: () => ({ name: 'a.xlsx', ext: 'xlsx', bytes: DOCX }), api: m.api });
  let appPort = null;
  const frameWin = { postMessage: (_msg, _t, transfer) => { appPort = transfer[0]; } };
  assert.equal(bridge.accept({ source: frameWin, data: { type: HELLO, v: 1 } }, frameWin), true);
  const answer = new Promise((resolve) => { appPort.onmessage = (e) => resolve(e.data); });
  appPort.postMessage({ id: 1, method: 'session.get' });
  const res = await answer;
  assert.equal(res.id, 1);
  assert.equal(res.result.files[0].ext, 'xlsx');
  assert.equal(res.result.files[0].mime, MIME.xlsx);
  bridge.reset();
  appPort.close();
});

// ── 앱 → host 요청 ─────────────────────────────────────

test('session.get: 설계 6장 모양 (ko-KR, 이름이 없으면 나, 파일 하나, grants)', async () => {
  const { port } = setup({ theme: 'dark' });
  const { msg } = await ask(port, 'session.get');
  const s = msg.result;
  assert.equal(s.v, 1);
  assert.equal(s.locale, 'ko-KR');
  assert.deepEqual(s.theme, { mode: 'dark', tokens: {} });
  assert.deepEqual(s.user, { name: '나' });
  assert.equal(s.app.version, '0.3.1');
  assert.deepEqual(s.files, [{ index: 0, name: '보고서.docx', ext: 'docx', size: DOCX.length, mime: MIME.docx, readOnly: false }]);
  assert.deepEqual(s.grants, [...GRANTS]);
  for (const g of ['files:read', 'files:write', 'ui', 'ui:eval', 'ui:wasm-eval', 'ui:package-fetch', 'ui:frame-package', 'ui:connect-blob', 'ui:download', 'ui:print']) {
    assert.ok(s.grants.includes(g), g);
  }
});

test('session.get: 사용자 이름은 main이 준 것 (변경 내용 추적·메모 작성자), 제어 문자 빼고 64자까지', async () => {
  const { port } = setup({ doc: { userName: '김선호' } });
  const { msg } = await ask(port, 'session.get');
  assert.deepEqual(msg.result.user, { name: '김선호' });
  assert.deepEqual(sessionInfo({ name: 'a.docx', ext: 'docx', userName: ' 이\u0007름 ' }).user, { name: '이름' });
  assert.equal(sessionInfo({ name: 'a.docx', ext: 'docx', userName: 'ㄱ'.repeat(100) }).user.name.length, 64);
  for (const v of ['', '   ', null, 7, {}]) assert.deepEqual(sessionInfo({ name: 'a.docx', ext: 'docx', userName: v }).user, { name: '나' }, String(v));
});

test('session.get: 읽기 전용이면 readOnly:true, files:write 없음', () => {
  const s = sessionInfo({ name: 'x.pptx', ext: 'PPTX', size: 9, readOnly: true });
  assert.equal(s.files[0].readOnly, true);
  assert.equal(s.files[0].ext, 'pptx');
  assert.equal(s.files[0].size, 9);
  assert.equal(s.files[0].mime, MIME.pptx);
  assert.ok(!s.grants.includes('files:write'));
  assert.ok(s.grants.includes('files:read'));
  assert.equal(s.locale, 'ko-KR');
});

test('file.read: 사본 바이트를 transfer로, 원본은 그대로 · 없는 index는 not_found', async () => {
  const { port, doc } = setup();
  const { msg, transfer } = await ask(port, 'file.read', { index: 0, as: 'bytes' });
  assert.equal(msg.result.name, '보고서.docx');
  assert.equal(msg.result.size, DOCX.length);
  assert.equal(msg.result.mime, MIME.docx);
  assert.ok(msg.result.bytes instanceof ArrayBuffer);
  assert.deepEqual(new Uint8Array(msg.result.bytes), DOCX);
  assert.equal(transfer.length, 1);
  assert.equal(transfer[0], msg.result.bytes);
  assert.notEqual(msg.result.bytes, doc.bytes.buffer, '원본 버퍼를 넘기면 안 됨');

  const nf = await ask(port, 'file.read', { index: 1 });
  assert.equal(nf.msg.error.code, 'not_found');

  const bad = await ask(port, 'file.read', { as: 'blob' });
  assert.equal(bad.msg.error.code, 'invalid');

  const st = await ask(port, 'file.read', { as: 'stream' });
  assert.ok(st.msg.result.stream);
  assert.deepEqual(await toBytes(st.msg.result.stream), DOCX);
});

test('file.save reason:user → editor:save {reason:user} → {saved:true,size}', async () => {
  const { port, main, events } = setup();
  const data = DOCX.slice().buffer;
  const { msg } = await ask(port, 'file.save', { index: 0, data, mime: MIME.docx, reason: 'user' });
  assert.deepEqual(msg.result, { saved: true, size: DOCX.length });
  assert.equal(main.calls.length, 1);
  assert.equal(main.calls[0][0], 'save');
  assert.ok(main.calls[0][1] instanceof Uint8Array);
  assert.deepEqual(main.calls[0][1], DOCX);
  assert.deepEqual(main.calls[0][2], { reason: 'user' });
  assert.deepEqual(ev(events, 'onSaveStart'), [['onSaveStart', { as: false }]]);
  const end = ev(events, 'onSaveEnd');
  assert.equal(end.length, 1);
  assert.equal(end[0][1].ok, true);
  assert.equal(end[0][1].result.name, 'a.docx');
});

test('file.save reason:auto → 복구용 사본만 (문서 파일 저장 아님)', async () => {
  const { port, main, events } = setup();
  const { msg } = await ask(port, 'file.save', { data: DOCX.slice().buffer, reason: 'auto' });
  assert.equal(msg.result.saved, true);
  assert.deepEqual(main.calls[0][2], { reason: 'auto' });
  assert.equal(ev(events, 'onAutoSaved').length, 1);
  assert.equal(ev(events, 'onSaveStart').length, 0, '자동 저장은 저장 중 표시를 하지 않음');
});

test('file.save reason 없음: host가 부탁한 저장 중이 아니면 auto (몰래 덮어쓰지 않음)', async () => {
  const { port, main } = setup();
  await ask(port, 'file.save', { data: DOCX.slice().buffer });
  assert.deepEqual(main.calls[0][2], { reason: 'auto' });
});

test('file.save 데이터 모양: TypedArray·ReadableStream 받음, 문자열·빈 것은 invalid', async () => {
  const { port, main } = setup();
  const view = new Uint8Array(20); view.set(DOCX, 5);
  let r = await ask(port, 'file.save', { data: view.subarray(5, 5 + DOCX.length), reason: 'user' });
  assert.equal(r.msg.result.saved, true);
  assert.deepEqual(main.calls.at(-1)[1], DOCX);

  const stream = new ReadableStream({ start(c) { c.enqueue(DOCX.slice(0, 4)); c.enqueue(DOCX.slice(4)); c.close(); } });
  r = await ask(port, 'file.save', { data: stream, reason: 'user' });
  assert.equal(r.msg.result.saved, true);
  assert.deepEqual(main.calls.at(-1)[1], DOCX);

  const n = main.calls.length;
  r = await ask(port, 'file.save', { data: 'PK…', reason: 'user' });
  assert.equal(r.msg.error.code, 'invalid');
  r = await ask(port, 'file.save', { data: new ArrayBuffer(0), reason: 'user' });
  assert.equal(r.msg.error.code, 'invalid');
  r = await ask(port, 'file.save', { reason: 'user' });
  assert.equal(r.msg.error.code, 'invalid');
  r = await ask(port, 'file.save', { index: 3, data: DOCX.slice().buffer, reason: 'user' });
  assert.equal(r.msg.error.code, 'not_found');
  assert.equal(main.calls.length, n, '거절한 것은 main에 가지 않음');
});

test('file.save: 읽기 전용 문서는 read_only, main에 가지 않음', async () => {
  const { port, main } = setup({ doc: { readOnly: true } });
  const r = await ask(port, 'file.save', { data: DOCX.slice().buffer, reason: 'user' });
  assert.equal(r.msg.error.code, 'read_only');
  assert.match(r.msg.error.message, /읽기 전용/);
  assert.equal(main.calls.length, 0);
});

test('file.save: main 저장 실패 → failed + 한국어 이유, 화면에 알림', async () => {
  const { port, events } = setup({ main: { save: () => { throw Object.assign(new Error('‘보고서.docx’ 파일을 다른 프로그램이 쓰고 있거나…'), { code: 'LOCKED' }); } } });
  const r = await ask(port, 'file.save', { data: DOCX.slice().buffer, reason: 'user' });
  assert.equal(r.msg.error.code, 'failed');
  assert.match(r.msg.error.message, /다른 프로그램이 쓰고 있거나/);
  const end = ev(events, 'onSaveEnd');
  assert.equal(end[0][1].ok, false);
  assert.equal(end[0][1].error.code, 'LOCKED');

  const ro = setup({ main: { save: () => { throw Object.assign(new Error('읽기 전용이라…'), { code: 'READ_ONLY' }); } } });
  const r2 = await ask(ro.port, 'file.save', { data: DOCX.slice().buffer, reason: 'user' });
  assert.equal(r2.msg.error.code, 'read_only');
});

test('ui.dirty → 화면에 그대로', async () => {
  const { port, events } = setup();
  await ask(port, 'ui.dirty', { dirty: true });
  await ask(port, 'ui.dirty', { dirty: false });
  await ask(port, 'ui.dirty', { dirty: 'yes' });
  assert.deepEqual(ev(events, 'onDirty').map((e) => e[1]), [true, false, false]);
});

test('ui.download → editor:download(이름, 바이트, mime) · 취소는 cancelled · 이름의 경로 글자는 지움', async () => {
  const { port, main } = setup();
  let r = await ask(port, 'ui.download', { name: '..\\보고서.pdf', data: PDF.slice().buffer, mime: 'application/pdf' });
  assert.deepEqual(r.msg.result, { saved: true, size: PDF.length });
  const call = main.calls.at(-1);
  assert.equal(call[0], 'download');
  assert.equal(call[1], '_보고서.pdf');
  assert.deepEqual(call[2], PDF);
  assert.equal(call[3], 'application/pdf');

  const c = setup({ main: { download: () => ({ canceled: true }) } });
  r = await ask(c.port, 'ui.download', { name: 'a.odt', data: DOCX.slice().buffer });
  assert.equal(r.msg.error.code, 'cancelled');

  r = await ask(port, 'ui.download', { name: '', data: DOCX.slice().buffer });
  assert.equal(main.calls.at(-1)[1], '보고서', '이름이 없으면 문서 이름');
  r = await ask(port, 'ui.download', { name: 'x.txt' });
  assert.equal(r.msg.error.code, 'invalid');
});

test('ui.download·ui.print: main 실패는 failed + 한국어 이유, 화면에도 알림', async () => {
  const boom = () => { throw Object.assign(new Error('디스크 공간이 부족합니다'), { code: 'NO_SPACE' }); };
  const { port, events } = setup({ main: { download: boom, print: boom } });
  let r = await ask(port, 'ui.download', { name: 'a.pdf', data: PDF.slice().buffer, mime: 'application/pdf' });
  assert.equal(r.msg.error.code, 'failed');
  assert.equal(r.msg.error.message, '디스크 공간이 부족합니다');
  r = await ask(port, 'ui.print', { data: PDF.slice().buffer });
  assert.equal(r.msg.error.code, 'failed');
  assert.deepEqual(ev(events, 'onActionFailed').map((e) => [e[1], e[2].message]), [['download', '디스크 공간이 부족합니다'], ['print', '디스크 공간이 부족합니다']]);
});

test('ui.print → PDF만 editor:print로', async () => {
  const { port, main } = setup();
  let r = await ask(port, 'ui.print', { name: '보고서.pdf', data: PDF.slice().buffer, mime: 'application/pdf' });
  assert.equal(r.msg.result.printed, true);
  assert.equal(main.calls.at(-1)[0], 'print');
  assert.deepEqual(main.calls.at(-1)[1], PDF);
  const n = main.calls.length;
  r = await ask(port, 'ui.print', { data: DOCX.slice().buffer });
  assert.equal(r.msg.error.code, 'invalid');
  assert.equal(main.calls.length, n);
});

test('state.get/set → editor:state · 8KB 넘는 값은 too_large · 잘못된 키는 invalid', async () => {
  const { port, main } = setup();
  let r = await ask(port, 'state.get', { key: 'editor-settings' });
  assert.equal(r.msg.result, null);
  r = await ask(port, 'state.set', { key: 'editor-settings', value: { zoom: 120, unit: 'cm' } });
  assert.equal(r.msg.result, null);
  r = await ask(port, 'state.get', { key: 'editor-settings' });
  assert.deepEqual(r.msg.result, { zoom: 120, unit: 'cm' });
  assert.deepEqual(main.calls.filter((c) => c[0] === 'state').map((c) => c[1]), ['get', 'set', 'get']);

  const n = main.calls.length;
  r = await ask(port, 'state.set', { key: 'big', value: 'ㄱ'.repeat(3000) }); // UTF-8로 9KB 넘음
  assert.equal(r.msg.error.code, 'too_large');
  r = await ask(port, 'state.set', { key: 'ok', value: 'a'.repeat(LIMITS.maxStateBytes - 2) }); // 따옴표 포함 딱 8KB
  assert.equal(r.msg.result, null);
  r = await ask(port, 'state.get', { key: '' });
  assert.equal(r.msg.error.code, 'invalid');
  r = await ask(port, 'state.set', { key: 'x'.repeat(LIMITS.maxKeyChars + 1), value: 1 });
  assert.equal(r.msg.error.code, 'invalid');
  r = await ask(port, 'state.get', { key: '설정 이름' });
  assert.equal(r.msg.error.code, 'invalid', 'main 규칙: 영문·숫자·_ . : - 만');
  r = await ask(port, 'state.set', { key: 'f', value: 10n });
  assert.equal(r.msg.error.code, 'invalid');
  assert.equal(main.calls.length, n + 1, '거절한 것은 main에 가지 않음');
});

test('ui.toast(로그+화면) · ui.title(무시) · ui.close(닫기 흐름) · license.get(free)', async () => {
  const { port, events, logs } = setup();
  let r = await ask(port, 'ui.toast', { text: 'x'.repeat(400), tone: 'error' });
  assert.equal(r.msg.result, null);
  const t = ev(events, 'onToast');
  assert.equal(t[0][1].length, LIMITS.maxLineChars);
  assert.equal(t[0][2], 'error');
  assert.ok(logs.some((l) => l.startsWith('앱 알림(error)')));
  await ask(port, 'ui.toast', { text: '안내', tone: 'loud' });
  assert.equal(ev(events, 'onToast')[1][2], 'info');

  r = await ask(port, 'ui.title', { text: '제목' });
  assert.equal(r.msg.result, null);
  assert.equal(ev(events, 'onTitle').length, 0);

  r = await ask(port, 'ui.close');
  assert.equal(r.msg.result, null);
  assert.equal(ev(events, 'onClose').length, 1);

  r = await ask(port, 'license.get');
  assert.deepEqual(r.msg.result, { status: 'free' });
});

test('coedit.* 와 모르는 메서드는 unknown_method (혼자 편집), 이상한 메시지는 무시', async () => {
  const { port, main } = setup();
  for (const m of ['coedit.join', 'coedit.append', 'coedit.leave', 'file.saveAs', 'ui.confirm', 'clipboard.write', 'engine.call', 'job.submit', '__proto__', 'constructor', 'toString']) {
    const r = await ask(port, m, {});
    assert.equal(r.msg.error.code, 'unknown_method', m);
  }
  const before = port.sent.length;
  port.deliver(null);
  port.deliver('session.get');
  port.deliver({ method: 'session.get' });
  port.deliver({ id: '1', method: 'session.get' });
  port.deliver({ id: 5 });
  await sleep(20);
  assert.equal(port.sent.length, before);
  assert.equal(main.calls.length, 0);
});

// ── host → 앱 ──────────────────────────────────────────

/** 앱 흉내: {hid, request:'save'}를 받으면 file.save를 보내고 답한다 */
function fakeApp(port, { reason = 'user', answer = 'ok', saveFirst = true, data = () => DOCX.slice().buffer } = {}) {
  const seen = [];
  port.waitFor((m) => m.request === 'save', 3000).then(async ({ msg }) => {
    seen.push(msg);
    if (saveFirst) {
      const id = ++nextId;
      port.deliver({ id, method: 'file.save', params: { index: 0, data: data(), mime: MIME.docx, ...(reason ? { reason } : {}) } });
      const { msg: res } = await port.waitFor((m) => m.id === id);
      seen.push(res);
      if (res.error && answer === 'ok') { port.deliver({ hid: msg.hid, error: { code: 'failed', message: res.error.message } }); return; }
    }
    if (answer === 'ok') port.deliver({ hid: msg.hid, result: { saved: true } });
    else if (answer === 'unavailable') port.deliver({ hid: msg.hid, error: { code: 'unavailable', message: 'this app has no save handler' } });
  }).catch(() => {});
  return seen;
}

test('{hid, request:"save"} → 앱이 file.save(user) 뒤 {hid, result} → 저장됨', async () => {
  const { bridge, port, main } = setup();
  const seen = fakeApp(port);
  const r = await bridge.requestSave();
  assert.equal(r.ok, true);
  assert.equal(r.result.name, 'a.docx');
  assert.equal(seen[0].request, 'save');
  assert.equal(typeof seen[0].hid, 'number');
  assert.deepEqual(Object.keys(seen[0]).sort(), ['hid', 'request']);
  assert.deepEqual(main.calls.map((c) => [c[0], c[2]]), [['save', { reason: 'user' }]]);
});

test('저장 부탁 중 reason 없는 file.save는 user로 (패치 전 filex)', async () => {
  const { bridge, port, main } = setup();
  fakeApp(port, { reason: null });
  const r = await bridge.requestSave();
  assert.equal(r.ok, true);
  assert.deepEqual(main.calls[0][2], { reason: 'user' });
});

test('다른 이름으로 저장: 같은 저장 부탁을 editor:save-as(같은 형식)로 · 취소', async () => {
  const { bridge, port, main } = setup();
  fakeApp(port);
  let r = await bridge.requestSave({ as: true });
  assert.equal(r.ok, true);
  assert.equal(r.as, true);
  assert.equal(main.calls[0][0], 'saveAs');
  assert.equal(main.calls[0][2], 'docx');

  const c = setup({ main: { saveAs: () => ({ canceled: true }) } });
  const seen = fakeApp(c.port);
  r = await c.bridge.requestSave({ as: true });
  assert.deepEqual(r, { ok: false, canceled: true });
  assert.equal(seen[1].error.code, 'cancelled', '앱에는 cancelled');
});

test('저장 부탁: 앱이 file.save 없이 답하면 nothing, 저장 처리기 없음은 unavailable', async () => {
  let { bridge, port } = setup();
  fakeApp(port, { saveFirst: false });
  let r = await bridge.requestSave();
  assert.deepEqual(r, { ok: false, nothing: true });

  ({ bridge, port } = setup());
  fakeApp(port, { saveFirst: false, answer: 'unavailable' });
  r = await bridge.requestSave();
  assert.equal(r.ok, false);
  assert.equal(r.shown, false);
  assert.equal(r.error.code, 'unavailable');
});

test('저장 부탁: main 저장 실패는 shown:true (화면은 이미 onSaveEnd로 알림)', async () => {
  const { bridge, port, events } = setup({ main: { save: () => { throw Object.assign(new Error('디스크 공간이 부족해 저장하지 못했습니다.'), { code: 'NO_SPACE' }); } } });
  fakeApp(port);
  const r = await bridge.requestSave();
  assert.equal(r.ok, false);
  assert.equal(r.shown, true);
  assert.equal(r.error.code, 'NO_SPACE');
  assert.equal(ev(events, 'onSaveEnd')[0][1].ok, false);
});

test('저장 부탁: 시간 초과, 진행 중이면 같은 약속, 틀이 없어지면(reset) 끊김', async () => {
  let { bridge, port } = setup({ saveWaitMs: 60 });
  let r = await bridge.requestSave();
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'timeout');

  ({ bridge, port } = setup());
  const a = bridge.requestSave();
  const b = bridge.requestSave();
  assert.equal(a, b, '진행 중이면 같은 것');
  assert.equal(bridge.saving, true);
  await port.waitFor((m) => m.request === 'save');
  assert.equal(port.sent.filter((s) => s.msg.request === 'save').length, 1);
  bridge.reset();
  r = await a;
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'closed');
  assert.equal(port.closed, true);
  assert.equal(bridge.saving, false);

  const none = createBridge({ getDoc: () => ({}), api: fakeMain().api });
  r = await none.requestSave();
  assert.equal(r.error.code, 'not_connected');
});

test('다른 포트의 hid 답·다른 hid는 무시', async () => {
  const { bridge, port } = setup({ saveWaitMs: 200 });
  const p = bridge.requestSave();
  const { msg } = await port.waitFor((m) => m.request === 'save');
  port.deliver({ hid: msg.hid + 1, result: { saved: true } });
  const other = new FakePort();
  await bridge.handle({ hid: msg.hid, result: { saved: true } }, other);
  const r = await p;
  assert.equal(r.error.code, 'timeout');
});

test('이벤트: close.request · theme', async () => {
  const { bridge, port } = setup();
  assert.equal(bridge.emit('theme', { mode: 'dark', tokens: {} }), true);
  assert.deepEqual(port.sent.at(-1).msg, { event: 'theme', data: { mode: 'dark', tokens: {} } });
  const t0 = Date.now();
  await bridge.closeRequest({ waitMs: 10, idleMs: 50 });
  assert.deepEqual(port.sent.find((s) => s.msg.event === 'close.request').msg, { event: 'close.request' });
  assert.ok(Date.now() - t0 < 1000);
  bridge.reset();
  assert.equal(bridge.emit('theme', { mode: 'light' }), false, '포트 없으면 보내지 않음');
});

test('close.request 뒤 앱이 보낸 state.set을 잠깐 기다린다', async () => {
  const { bridge, port, main } = setup({ main: {} });
  // 앱 흉내: close.request를 받으면 설정을 보관한다 (main은 느리게)
  main.api.state = async (op, key, value) => { await sleep(80); main.calls.push(['state', op, key, value]); return null; };
  port.waitFor((m) => m.event === 'close.request').then(() => { port.deliver({ id: ++nextId, method: 'state.set', params: { key: 's', value: { a: 1 } } }); });
  await bridge.closeRequest({ waitMs: 20, idleMs: 500 });
  assert.deepEqual(main.calls.at(-1), ['state', 'set', 's', { a: 1 }]);
});

test('다시 불러온 뒤 옛 포트의 ui.dirty는 화면에 가지 않는다', async () => {
  const { bridge, port, events } = setup();
  const fresh = new FakePort();
  bridge.attach(fresh);
  assert.equal(port.closed, true);
  await bridge.handle({ id: 99, method: 'ui.dirty', params: { dirty: true } }, port);
  assert.equal(ev(events, 'onDirty').length, 0);
  fresh.deliver({ id: 100, method: 'ui.dirty', params: { dirty: true } });
  await fresh.waitFor((m) => m.id === 100);
  assert.deepEqual(ev(events, 'onDirty').map((e) => e[1]), [true]);
});

// ── 도우미 ─────────────────────────────────────────────

test('effectiveDirty: 앱 바뀜 · 자동 저장만 됨 · 복구한 내용 중 하나라도', () => {
  assert.equal(effectiveDirty({}), false);
  assert.equal(effectiveDirty({ appDirty: true }), true);
  assert.equal(effectiveDirty({ autoPending: true }), true, '복구용 사본에만 저장 = 파일은 아직 옛것');
  assert.equal(effectiveDirty({ restoredDraft: true }), true);
});

test('objectJosa: 받침 있으면 을, 없으면 를', () => {
  assert.equal(objectJosa('AGPL 전문과 추가 조항'), '을');
  assert.equal(objectJosa('filex 변경 목록'), '을');
  assert.equal(objectJosa('filex-office-editor 고지'), '를');
  assert.equal(objectJosa('제3자 구성 요소'), '를');
  assert.equal(objectJosa('NOTICE'), '를');
});

test('toBytes · safeFileName', async () => {
  assert.equal(await toBytes('x'), null);
  assert.equal(await toBytes(null), null);
  await assert.rejects(toBytes(new Uint8Array(10), 5), (e) => e.bridgeCode === 'too_large');
  const big = new ReadableStream({ start(c) { c.enqueue(new Uint8Array(4)); c.enqueue(new Uint8Array(4)); c.close(); } });
  await assert.rejects(toBytes(big, 6), (e) => e.bridgeCode === 'too_large');
  assert.deepEqual(await toBytes(new Blob([DOCX])), DOCX);
  assert.equal(safeFileName('a/b:c*.pdf'), 'a_b_c_.pdf');
  assert.equal(safeFileName('  ..  '), '내려받기');
  assert.equal(safeFileName('', '보고서'), '보고서');
  const long = safeFileName(`${'가'.repeat(300)}.pdf`);
  assert.equal(long.length, LIMITS.maxNameChars);
  assert.ok(long.endsWith('.pdf'));
});

// ── 정보 · 라이선스 (설계 8장) ─────────────────────────

const VERSION_JSON = {
  onlyoffice: { version: '9.4.0', build: '9.4.0.129', sourceTag: 'v9.4.0.129', image: 'docker.io/onlyoffice/documentserver@sha256:3ab6ebc7c605e5a32b7ae3ff19daed4925090245acc8100ce2230bd766c88212' },
  filex: { tag: 'v0.3.1', commit: 'acc87f88da472bcf984f8c27baafe6ad788e3f61' },
  x2t: { release: 'v9.4.0.129+2', core: 'a016fc2868b9f99f59dd30aa2088e68bbe5104af', buildTools: 'cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5' },
  builtAt: '2026-10-20',
};
const APP_URL = 'http://127.0.0.3:5123/office/9.4.0.129/index.html';

test('정보 · 라이선스: 원 개발사 · 수정판일 수 있음 · AGPL·추가 조항 · 설치된 판의 소스 주소', () => {
  const m = aboutModel({ appVersion: '0.2.0', bundle: VERSION_JSON }, APP_URL);
  const text = JSON.stringify(m);
  assert.equal(m.known, true);
  // (i) 원 개발사
  assert.match(text, /원 개발사는 ONLYOFFICE\(Ascensio System SIA/);
  // (ii) 수정판일 수 있음 (+ 수정 고지 파일)
  assert.match(text, /수정판일 수 있습니다/);
  // (iii) 라이선스 정보: AGPL 전문·추가 조항 1~5, CC BY-SA, 제3자, 상표, 보증 없음
  assert.match(text, /AGPL-3\.0/);
  for (const n of ['1. ', '2. ', '3. ', '4. ', '5. ']) assert.ok(m.sections.some((s) => s.items?.some((i) => i.startsWith(n))), `추가 조항 ${n}`);
  assert.match(text, /CC BY-SA 4\.0/);
  assert.match(text, /trademark-policy/);
  assert.match(text, /보증도 없이/);
  assert.match(text, /후원이나 승인을 받지 않았습니다/);
  // 설치된 판 (VERSION.json에서)
  const facts = Object.fromEntries(m.facts.map((f) => [f.label, f.value]));
  assert.equal(facts['Office Kit'], 'v0.2.0');
  assert.match(facts['편집기'], /9\.4\.0 \(빌드 9\.4\.0\.129\)/);
  assert.equal(facts['편집기 소스 태그'], 'v9.4.0.129');
  assert.match(facts['filex-office-editor'], /v0\.3\.1 · 커밋 acc87f88da47/);
  assert.match(facts['변환기 x2t'], /v9\.4\.0\.129\+2/);
  assert.equal(facts['묶음 만든 날'], '2026-10-20');
  // 소스 주소는 설치된 판의 태그·커밋
  const urls = m.links.map((l) => l.url);
  for (const u of [
    'https://github.com/toxh13/office-kit/releases/tag/v0.2.0',
    'https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129',
    'https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129',
    'https://github.com/ONLYOFFICE/core/tree/a016fc2868b9f99f59dd30aa2088e68bbe5104af',
    'https://github.com/ONLYOFFICE/build_tools/tree/cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5',
    'https://github.com/BRF-Tech/filex-office-editor/tree/acc87f88da472bcf984f8c27baafe6ad788e3f61',
  ]) assert.ok(urls.includes(u), u);
  assert.ok(urls.every((u) => u.startsWith('https://')));
  // 묶음 안 고지 파일은 편집기 출처에서 (LICENSE.txt = AGPL 전문 + 추가 조항)
  assert.deepEqual(m.files.map((f) => f.url), NOTICE_FILES.map((f) => `http://127.0.0.3:5123/office/9.4.0.129/${f.path}`));
  assert.equal(m.files[0].url.endsWith('/editor/LICENSE.txt'), true);
});

test('정보 · 라이선스: 거울 주소(sources)를 앞에, https만 · VERSION.json이 없으면 known:false', () => {
  const m = aboutModel({
    appVersion: '0.2.0', bundle: VERSION_JSON,
    sources: [{ label: '거울: sdkjs', url: 'https://github.com/toxh13/onlyoffice-sdkjs/tree/v9.4.0.129' }, { label: '나쁜 것', url: 'javascript:alert(1)' }, { label: 'http', url: 'http://example.com' }],
  }, APP_URL);
  assert.equal(m.links[0].label, '거울: sdkjs');
  assert.ok(!m.links.some((l) => /javascript|http:\/\//.test(l.url)));

  const u = aboutModel({ appVersion: '0.2.0' }, 'not a url');
  assert.equal(u.known, false);
  // 앱 판을 못 받으면 묶음을 만든 Office Kit 판, 만든 날은 날짜만
  const b = aboutModel({ bundle: { ...VERSION_JSON, officeKitVersion: '0.1.2', builtAt: '2026-10-18T08:36:58.727Z' } }, APP_URL);
  assert.ok(b.links.some((l) => l.url === 'https://github.com/toxh13/office-kit/releases/tag/v0.1.2'));
  assert.equal(b.facts.find((x) => x.label === '묶음 만든 날').value, '2026-10-18');
  assert.equal(u.files.length, 0);
  assert.ok(u.links.some((l) => l.url === 'https://github.com/ONLYOFFICE/sdkjs'));
  assert.ok(u.sections.length >= 5, '판 정보가 없어도 법적 고지는 그대로');

  // 태그·커밋 자리에 이상한 글자는 쓰지 않는다
  const f = bundleFacts({ onlyoffice: { build: '9.4.0.129', sourceTag: 'v1/../../x' }, filex: { commit: 'abc def' } });
  assert.equal(f.onlyoffice.sourceTag, 'v9.4.0.129', '이상한 태그 대신 빌드 번호로');
  assert.equal(bundleFacts({ onlyoffice: { sourceTag: 'v1/../../x' } }).onlyoffice.sourceTag, null);
  assert.equal(f.filex.commit, null);
  assert.equal(bundleFacts({ onlyoffice: { build: '9.4.0.129' } }).onlyoffice.sourceTag, 'v9.4.0.129');
});

// ── 정적 파일 ─────────────────────────────────────────

test('host.html: 스크립트는 host.js 모듈 하나, 인라인 스크립트·style 속성 없음, 막대에 저장·정보 · 라이선스', () => {
  const html = readFileSync(join(HOST_DIR, 'host.html'), 'utf8');
  const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)];
  assert.equal(scripts.length, 1);
  assert.match(scripts[0][0], /type="module"\s+src="host\.js"/);
  assert.equal(scripts[0][1].trim(), '');
  assert.doesNotMatch(html, /\sstyle="/i);
  assert.doesNotMatch(html, /\son[a-z]+="/i, '인라인 이벤트 처리기 없음');
  assert.match(html, /<link rel="stylesheet" href="host\.css">/);
  assert.match(html, /id="btn-save"[^>]*>.*저장/);
  assert.match(html, /id="btn-about"[^>]*>.*정보 · 라이선스/);
  assert.match(html, /<html lang="ko">/);
  // 상표: Office Kit 화면 이름·제목에 ONLYOFFICE를 쓰지 않는다 (고지 글은 host.js가 만든다)
  assert.doesNotMatch(html, /ONLYOFFICE/i);
  for (const k of Object.values(KINDS)) assert.doesNotMatch(`${k.badge} ${k.title}`, /ONLYOFFICE|Word|Excel|PowerPoint/);
  // 고지 파일 보기 틀은 스크립트 없이
  assert.match(html, /id="about-viewer"[^>]*sandbox=""/);
  const css = readFileSync(join(HOST_DIR, 'host.css'), 'utf8');
  assert.match(css, /--bar-h:\s*28px/);
});

test('host.js는 편집기 틀에 스크립트를 넣지 않는다 (MessagePort로만)', () => {
  const js = readFileSync(join(HOST_DIR, 'host.js'), 'utf8');
  assert.doesNotMatch(js, /contentWindow\.(eval|document|location)|contentDocument|executeJavaScript|srcdoc|\.eval\(/);
  assert.match(js, /setAttribute\('sandbox', 'allow-scripts'\)/);
});

// ── preload ───────────────────────────────────────────

test('office-preload.cjs: window.officeKitOffice만, editor:* 채널로', async () => {
  const require = createRequire(import.meta.url);
  const Module = require('node:module');
  const exposed = {};
  const invoked = []; const sent = []; const listeners = new Map();
  const fakeElectron = {
    contextBridge: { exposeInMainWorld: (name, api) => { exposed[name] = api; } },
    ipcRenderer: {
      invoke: (ch, ...a) => { invoked.push([ch, ...a]); return Promise.resolve({ ok: true, value: null }); },
      send: (ch, ...a) => { sent.push([ch, ...a]); },
      on: (ch, fn) => { listeners.set(ch, fn); },
      removeListener: (ch) => { listeners.delete(ch); },
    },
  };
  const orig = Module._load;
  Module._load = function (request, ...rest) { return request === 'electron' ? fakeElectron : orig.call(this, request, ...rest); };
  const file = join(HOST_DIR, 'office-preload.cjs');
  try { delete require.cache[file]; require(file); } finally { Module._load = orig; }

  assert.deepEqual(Object.keys(exposed), ['officeKitOffice']);
  const api = exposed.officeKitOffice;
  assert.deepEqual(Object.keys(api).sort(), [
    'close', 'discardDraft', 'download', 'getDocument', 'log', 'onCommand', 'onExternalChange', 'onRequestClose',
    'panel', 'print', 'readDraft', 'reportDirty', 'save', 'saveAs', 'state',
  ]);
  api.panel();
  assert.deepEqual(sent.at(-1), ['editor:panel'], "'AI · 도구' 단추 → main에 옆 창을 부탁");
  sent.length = 0;
  const bytes = new Uint8Array([1, 2]);
  await api.getDocument();
  await api.readDraft();
  await api.discardDraft();
  await api.save(bytes, { reason: 'auto' });
  await api.save(bytes.buffer, { reason: 'weird' });
  await api.save('문자열');
  await api.saveAs(bytes, 'docx');
  await api.download('a.pdf', bytes, 'application/pdf');
  await api.print(bytes);
  await api.state('get', 'k');
  await api.state('set', 'k', { a: 1 });
  await api.state('drop', 'k', 1);
  assert.deepEqual(invoked.map((c) => c[0]), [
    'editor:get-document', 'editor:draft', 'editor:draft', 'editor:save', 'editor:save', 'editor:save', 'editor:save-as',
    'editor:download', 'editor:print', 'editor:state', 'editor:state', 'editor:state',
  ]);
  assert.deepEqual(invoked[1], ['editor:draft', 'read']);
  assert.deepEqual(invoked[2], ['editor:draft', 'discard']);
  assert.deepEqual(invoked[3][2], { reason: 'auto' });
  assert.ok(invoked[4][1] instanceof Uint8Array);
  assert.deepEqual(invoked[4][2], { reason: 'user' }, '모르는 reason은 user');
  assert.equal(invoked[5][1], null, '바이트가 아니면 null (main이 거절)');
  assert.deepEqual(invoked[7], ['editor:download', 'a.pdf', bytes, 'application/pdf']);
  assert.deepEqual(invoked[9], ['editor:state', 'get', 'k']);
  assert.deepEqual(invoked[10], ['editor:state', 'set', 'k', { a: 1 }]);
  assert.deepEqual(invoked[11], ['editor:state', 'get', 'k'], '모르는 op는 get');

  api.reportDirty(1);
  api.reportDirty(true);
  api.log('x'.repeat(3000));
  api.close();
  assert.deepEqual(sent.map((s) => s[0]), ['editor:dirty', 'editor:dirty', 'editor:log', 'editor:close']);
  assert.equal(sent[0][1], false);
  assert.equal(sent[1][1], true);
  assert.equal(sent[2][1].length, 2000);

  // main이 묻는 것에 답하기: editor:reply {id, value(boolean|null)}
  api.onCommand(async (cmd) => cmd === 'save');
  await listeners.get('editor:command')({}, { id: '7', command: 'save' });
  assert.deepEqual(sent.at(-1), ['editor:reply', { id: '7', value: true }]);
  api.onRequestClose(async () => 'yes');
  await listeners.get('editor:request-close')({}, { id: '8' });
  assert.deepEqual(sent.at(-1), ['editor:reply', { id: '8', value: null }]);
  const off = api.onExternalChange(() => {});
  assert.ok(listeners.has('editor:external-change'));
  off();
  assert.ok(!listeners.has('editor:external-change'));
});
