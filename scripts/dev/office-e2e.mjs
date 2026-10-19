#!/usr/bin/env node
// office-e2e.mjs — 문서·표·슬라이드 편집 창 e2e (설계 docs/OFFICE-EDITOR-DESIGN.md 10장).
//   실제 앱(electron .)을 --remote-debugging-port로 띄우고 CDP로만 조작한다 (OS 키 입력·실제 IME 없음).
//   형식 3(docx·xlsx·pptx) × 만든 방법 3(createDocument, '새 문서'(doc:new edit:true), Word·Excel·PowerPoint COM)
//   → 편집 창 열기 → 한국어 입력(insertText·IME 조합 흉내)·굵게·표 행/수식·행 삽입/새 슬라이드 → Ctrl+S
//   → readDocument(engine:'file') + (선택) COM 엄격 열기(scripts/dev/strict-open.ps1)로 경고 없이 열리고 고친 내용이 있는지.
//   함께: 밖으로 나간 요청 0건(CDP 네트워크 기록), WebRTC 후보 0개, 자동 저장 → 복구용 사본만, 바깥 변경 띠,
//         원래대로, 화면 멈춤 뒤 복구 사본, 닫기 확인(취소·저장·저장 안 함), AI 보내기 막기.
//
//   node scripts/dev/office-e2e.mjs [--out 폴더] [--formats docx,xlsx,pptx] [--methods create,blank,office]
//                                   [--com | --no-com] [--no-flows] [--office-docs 폴더] [--keep]
//   결과: <out>/report.json, <out>/shots/*.png, 끝 코드 0 = 모두 통과, 1 = 실패 있음, 2 = 실행 못 함
//   시험 묶음: test/office-editor-e2e.test.mjs (OFFICEKIT_EDITOR_SMOKE=1, COM은 OFFICEKIT_COM_TESTS=1)
//
// 창이 몇 개 뜨고 닫힌다. 닫기 확인창·'멈췄습니다' 창의 단추는 UI Automation으로 찾아 BN_CLICKED 메시지만 보낸다
// (키보드·마우스 입력을 보내지 않으므로 다른 창이 입력을 받지 않는다).
import { spawn, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
export const ROOT = resolve(here, '..', '..');

export const FORMATS = ['docx', 'xlsx', 'pptx'];
export const METHODS = ['create', 'blank', 'office'];
const LABEL = { docx: '문서 편집', xlsx: '표 편집', pptx: '슬라이드 편집' };

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (p) => { try { return createHash('sha1').update(readFileSync(p)).digest('hex'); } catch { return null; } };
const mtime = (p) => { try { return statSync(p).mtimeMs; } catch { return null; } };
const flat = (s) => String(s ?? '').replace(/\s+/g, '');

async function waitFor(fn, ms = 20_000, step = 250) {
  const end = Date.now() + ms;
  let last;
  while (Date.now() < end) {
    try { last = await fn(); if (last) return last; } catch {}
    await sleep(step);
  }
  return null;
}

function freePort() {
  return new Promise((ok, no) => {
    const s = createServer();
    s.once('error', no);
    s.listen(0, '127.0.0.1', () => { const { port } = s.address(); s.close(() => ok(port)); });
  });
}

/** 이 PC 주소·내장 주소인지 (밖으로 나가지 않는 요청) */
export function isLocalUrl(url) {
  const u = String(url ?? '');
  if (/^(blob|data|about|chrome|chrome-extension|devtools|chrome-devtools|file):/i.test(u)) return true;
  try {
    const h = new URL(u).hostname.replace(/^\[|\]$/g, '');
    return /^127\.\d+\.\d+\.\d+$/.test(h) || h === 'localhost' || h === '::1';
  } catch { return false; }
}

// ── CDP (브라우저 한 연결, flatten 세션) ─────────────────────

class Cdp {
  constructor(ws) {
    this.ws = ws; this.seq = 0; this.pending = new Map(); this.handlers = new Map();
    ws.addEventListener('message', (m) => {
      const d = JSON.parse(m.data);
      if (d.id && this.pending.has(d.id)) {
        const p = this.pending.get(d.id); this.pending.delete(d.id);
        if (d.error) p.no(Object.assign(new Error(`${p.method}: ${d.error.message}`), { cdp: d.error })); else p.ok(d.result);
        return;
      }
      for (const fn of this.handlers.get(d.method) ?? []) { try { fn(d.params ?? {}, d.sessionId); } catch (err) { console.error('[cdp]', d.method, err); } }
    });
    ws.addEventListener('close', () => { for (const p of this.pending.values()) p.no(new Error('CDP 연결이 끊겼습니다')); this.pending.clear(); });
  }
  static async connect(url) {
    const ws = new WebSocket(url);
    await new Promise((ok, no) => { ws.addEventListener('open', ok, { once: true }); ws.addEventListener('error', () => no(new Error('CDP 연결 실패')), { once: true }); });
    return new Cdp(ws);
  }
  send(method, params = {}, sessionId, timeoutMs = 60_000) {
    if (this.ws.readyState !== WebSocket.OPEN) return Promise.reject(new Error('CDP 연결이 끊겼습니다'));
    const id = ++this.seq;
    return new Promise((ok, no) => {
      const t = setTimeout(() => { this.pending.delete(id); no(new Error(`${method}: 시간 초과`)); }, timeoutMs);
      this.pending.set(id, { ok: (v) => { clearTimeout(t); ok(v); }, no: (e) => { clearTimeout(t); no(e); }, method });
      this.ws.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });
  }
  on(method, fn) { if (!this.handlers.has(method)) this.handlers.set(method, []); this.handlers.get(method).push(fn); }
  close() { try { this.ws.close(); } catch {} }
}

/** 모든 창·틀·워커에 붙어 네트워크·실행 문맥을 기록한다 (새 대상은 시작 전에 멈춰 두고 기록을 켠 뒤 돌린다) */
class Tracker {
  constructor(cdp) {
    this.cdp = cdp;
    this.sessions = new Map(); // sid → { sid, type, url, targetId, parent, root, frames: Map(frameId→url), ctx: Map(frameId→contextId) }
    this.requests = [];        // { root, type, url, kind, failed, blocked }
    this.byReq = new Map();
    this.exceptions = [];
    this.crashed = new Set();
    cdp.on('Target.attachedToTarget', (p, parentSid) => { void this.attached(p, parentSid); });
    cdp.on('Target.detachedFromTarget', (p) => { this.sessions.delete(p.sessionId); });
    cdp.on('Target.targetInfoChanged', (p) => { for (const s of this.sessions.values()) if (s.targetId === p.targetInfo.targetId) s.url = p.targetInfo.url; });
    cdp.on('Inspector.targetCrashed', (_p, sid) => { this.crashed.add(sid); });
    cdp.on('Page.frameNavigated', (p, sid) => { const s = this.sessions.get(sid); if (s) s.frames.set(p.frame.id, p.frame.url + (p.frame.urlFragment ?? '')); if (s && !p.frame.parentId) s.url = p.frame.url; });
    cdp.on('Page.frameDetached', (p, sid) => { const s = this.sessions.get(sid); if (s) { s.frames.delete(p.frameId); s.ctx.delete(p.frameId); } });
    cdp.on('Runtime.executionContextCreated', ({ context: c }, sid) => {
      const s = this.sessions.get(sid);
      if (s && c.auxData?.isDefault && c.auxData.frameId) s.ctx.set(c.auxData.frameId, c.id);
    });
    cdp.on('Runtime.executionContextDestroyed', (p, sid) => {
      const s = this.sessions.get(sid);
      if (s) for (const [f, id] of s.ctx) if (id === p.executionContextId) s.ctx.delete(f);
    });
    cdp.on('Runtime.executionContextsCleared', (_p, sid) => { this.sessions.get(sid)?.ctx.clear(); });
    cdp.on('Runtime.exceptionThrown', (p, sid) => {
      const s = this.sessions.get(sid);
      this.exceptions.push({ root: s?.root, url: s?.url, text: String(p.exceptionDetails?.exception?.description ?? p.exceptionDetails?.text ?? '').slice(0, 400) });
    });
    cdp.on('Network.requestWillBeSent', (p, sid) => {
      const s = this.sessions.get(sid);
      const r = { root: s?.root ?? sid, rootUrl: this.sessions.get(s?.root)?.url ?? s?.url, session: s?.type, type: p.type ?? '', url: p.request.url, local: isLocalUrl(p.request.url) };
      this.requests.push(r);
      this.byReq.set(`${sid}:${p.requestId}`, r);
    });
    cdp.on('Network.loadingFailed', (p, sid) => {
      const r = this.byReq.get(`${sid}:${p.requestId}`);
      if (r) { r.failed = p.errorText; r.blocked = p.blockedReason ?? null; }
    });
    cdp.on('Network.webSocketCreated', (p, sid) => {
      const s = this.sessions.get(sid);
      this.requests.push({ root: s?.root ?? sid, rootUrl: this.sessions.get(s?.root)?.url ?? s?.url, session: s?.type, type: 'WebSocket', url: p.url, local: isLocalUrl(p.url) });
    });
  }

  async attached({ sessionId: sid, targetInfo: t, waitingForDebugger }, parentSid) {
    const parent = parentSid ? this.sessions.get(parentSid) : null;
    const s = { sid, type: t.type, url: t.url, targetId: t.targetId, parent: parentSid ?? null, root: parent ? parent.root : sid, frames: new Map(), ctx: new Map() };
    this.sessions.set(sid, s);
    const send = (m, p) => this.cdp.send(m, p ?? {}, sid).catch(() => null);
    // 기다리지 않고 차례로 보낸다 (멈춰 둔 새 창은 렌더러가 아직 없어 틀 쪽 명령의 답이 늦다). 같은 세션 안에서는 보낸 순서대로 처리된다
    const jobs = [];
    if (['page', 'iframe', 'worker', 'shared_worker'].includes(t.type)) jobs.push(send('Network.enable'), send('Runtime.enable'));
    if (t.type === 'page' || t.type === 'iframe') {
      jobs.push(send('Page.enable'), send('Target.setAutoAttach', { autoAttach: true, waitForDebuggerOnStart: true, flatten: true }));
    }
    if (waitingForDebugger) jobs.push(send('Runtime.runIfWaitingForDebugger'));
    await Promise.all(jobs);
    if (t.type === 'page' || t.type === 'iframe') {
      const tree = await send('Page.getFrameTree');
      const walk = (n) => { if (!n) return; if (!s.frames.has(n.frame.id)) s.frames.set(n.frame.id, n.frame.url); for (const c of n.childFrames ?? []) walk(c); };
      walk(tree?.frameTree);
    }
  }

  pages(re) { return [...this.sessions.values()].filter((s) => s.type === 'page' && re.test(s.url)); }
  page(re) { return this.pages(re).at(-1) ?? null; }

  /** 주소가 re에 맞는 틀 → { sid, ctx, url } (가장 나중 것) */
  frame(re, root) {
    let hit = null;
    for (const s of this.sessions.values()) {
      if (root && s.root !== root) continue;
      for (const [fid, url] of s.frames) {
        if (!re.test(url)) continue;
        const ctx = s.ctx.get(fid);
        if (ctx) hit = { sid: s.sid, ctx, url };
      }
    }
    return hit;
  }

  async evalIn(re, expression, { root, timeoutMs = 60_000 } = {}) {
    const f = this.frame(re, root);
    if (!f) throw new Error(`틀 없음: ${re}`);
    const r = await this.cdp.send('Runtime.evaluate', { expression, contextId: f.ctx, returnByValue: true, awaitPromise: true }, f.sid, timeoutMs);
    if (r.exceptionDetails) throw new Error(`${expression.slice(0, 80)}: ${r.exceptionDetails.exception?.description ?? r.exceptionDetails.text}`);
    return r.result?.value;
  }
}

// ── 앱 ──────────────────────────────────────────────────

/** 틀 주소 (주 창, Office host, filex 앱, 편집기) */
export const RE = {
  main: /\/src\/app\/renderer\/index\.html/,
  host: /\/office-host\/host\.html/,
  app: /^http:\/\/127\.0\.0\.\d+:\d+\/office\/[^/]+\/index\.html/,
  editor: /\/office\/[^/]+\/editor\/web-apps\/apps\/\w+\/main\/index\.html/,
};

/** 실제 앱 하나 (electron . + remote debugging, CDP 한 연결). 손으로 쓰는 확인 스크립트도 쓴다 */
export class App {
  constructor({ out, log }) { this.out = out; this.log = log; this.child = null; this.stdout = ''; this.exit = null; }

  async start({ ud }) {
    const electron = (await import('electron')).default;
    this.port = await freePort();
    const env = { ...process.env, OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_TEST: '1' };
    for (const k of ['ELECTRON_RUN_AS_NODE', 'OFFICEKIT_SMOKE', 'OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_VENDOR', 'OFFICEKIT_OFFICE_FONTS_MODULE', 'OFFICEKIT_AI_MODULE']) delete env[k];
    // 가려진 창도 계속 그리고 타이머를 늦추지 않게 (사람이 다른 창을 써도 시험이 흔들리지 않게)
    const args = ['.', `--remote-debugging-port=${this.port}`, '--disable-features=CalculateNativeWinOcclusion', '--disable-renderer-backgrounding', '--disable-background-timer-throttling'];
    // DNS 조회 기록 (보안 7장: 이 PC 밖 이름은 풀지 않는다 — main.mjs host-resolver-rules). 끝에 netlogDns()로 본다
    if (this.out) {
      this.netlog = join(resolve(this.out), 'netlog.json');
      rmSync(this.netlog, { force: true });
      args.push(`--log-net-log=${this.netlog}`);
    }
    this.child = spawn(electron, args, { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
    const take = (b) => { this.stdout += b; if (this.stdout.length > 4_000_000) this.stdout = this.stdout.slice(-2_000_000); };
    this.child.stdout.on('data', take);
    this.child.stderr.on('data', take);
    this.child.on('exit', (c) => { this.exit = c ?? -1; });
    const ver = await waitFor(async () => (await fetch(`http://127.0.0.1:${this.port}/json/version`)).json(), 30_000, 200);
    if (!ver?.webSocketDebuggerUrl) throw new Error(`앱이 뜨지 않았습니다:\n${this.stdout.slice(-2000)}`);
    this.browser = ver.Browser;
    this.cdp = await Cdp.connect(ver.webSocketDebuggerUrl);
    this.t = new Tracker(this.cdp);
    await this.cdp.send('Target.setDiscoverTargets', { discover: true }); // targetInfoChanged (주소 바뀜)
    await this.cdp.send('Target.setAutoAttach', { autoAttach: true, waitForDebuggerOnStart: true, flatten: true });
    // 이미 떠 있던 대상 (자동으로 붙지 않은 것)
    const { targetInfos } = await this.cdp.send('Target.getTargets');
    for (const ti of targetInfos) {
      if (ti.attached || ti.type !== 'page') continue;
      await this.cdp.send('Target.attachToTarget', { targetId: ti.targetId, flatten: true }).catch(() => null);
    }
    const main = await waitFor(() => this.t.page(RE.main) && this.t.frame(RE.main), 30_000);
    if (!main) throw new Error('주 창을 찾지 못했습니다');
    await waitFor(() => this.main('typeof window.officeKit?.doc?.open === "function" && document.readyState === "complete"'), 30_000);
    await sleep(800);
  }

  main(expr) { return this.t.evalIn(RE.main, expr, { root: this.t.page(RE.main)?.sid }); }
  /** 주 창 preload API 부르기 → value (실패면 { error, code }를 담은 Error) */
  async api(path, ...args) {
    const r = await this.main(`window.officeKit.${path}(...${JSON.stringify(args)})`);
    if (!r?.ok) throw Object.assign(new Error(r?.error ?? `${path} 실패`), { code: r?.code ?? null });
    return r.value;
  }
  async apiRaw(path, ...args) { return this.main(`window.officeKit.${path}(...${JSON.stringify(args)})`); }

  /** 편집 창 PowerShell 도우미 (UIA: 창 제목, 확인창 글·단추) */
  ps(action, arg = '', timeoutMs = 15_000) {
    if (!this.child) return null;
    const out = join(this.out, `ps-${process.pid}-${Date.now()}.json`);
    const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psHelper(this.out), '-ProcId', String(this.child.pid),
      '-Action', action, '-ArgB64', Buffer.from(`|${arg}`, 'utf8').toString('base64'), '-TimeoutMs', String(timeoutMs), '-Out', out], { encoding: 'utf8', timeout: timeoutMs + 20_000 });
    let v = null;
    try { v = JSON.parse(readFileSync(out, 'utf8').replace(/^﻿/, '')); } catch { v = { error: (r.stderr || r.stdout || 'no output').slice(0, 400) }; }
    rmSync(out, { force: true });
    return v;
  }

  /**
   * netlog에서 DNS 조회 → { read, events, lookups: [이 PC 밖 이름], transactions } (파일을 못 읽으면 read:false).
   * 강제로 끝낸 앱의 netlog는 끝이 잘려 있을 수 있어 줄마다 읽는다.
   */
  netlogDns() {
    const out = { read: false, events: 0, lookups: [], transactions: 0 };
    let text;
    try { text = readFileSync(this.netlog, 'utf8'); } catch { return out; }
    const at = text.indexOf('"events"');
    let types;
    try { types = JSON.parse(`${text.slice(0, at).replace(/,\s*$/, '')}}`).constants.logEventTypes; } catch { return out; }
    const id = (n) => types[n];
    const job = id('HOST_RESOLVER_MANAGER_JOB'); const req = id('HOST_RESOLVER_MANAGER_REQUEST'); const tx = id('DNS_TRANSACTION');
    const local = (h) => /^(localhost|127\.\d+\.\d+\.\d+|\[?::1\]?)(:\d+)?$/i.test(String(h ?? '').replace(/^[a-z]+:\/\//i, '').replace(/\/.*$/, ''));
    const names = new Set();
    for (const line of text.slice(at).split('\n')) {
      const l = line.trim().replace(/,$/, '');
      if (!l.startsWith('{')) continue;
      let e; try { e = JSON.parse(l); } catch { continue; }
      out.events++;
      if (e.type === tx) out.transactions++;
      if ((e.type === job || e.type === req) && e.params) {
        const h = e.params.host ?? e.params.hostname ?? e.params.key?.host;
        if (h && !local(h)) names.add(String(h));
      }
    }
    out.read = out.events > 0;
    out.lookups = [...names];
    return out;
  }

  async stop() {
    try { this.cdp?.close(); } catch {}
    if (this.child && this.exit === null) {
      spawnSync('taskkill', ['/pid', String(this.child.pid), '/T', '/F'], { stdio: 'ignore' });
      await waitFor(() => this.exit !== null, 10_000, 100);
    }
  }
}

// ── 편집 창 (host 페이지 세션 하나) ─────────────────────────

/** 편집 창 하나 (host 페이지 세션) */
export class Editor {
  constructor(app, page) { this.app = app; this.t = app.t; this.cdp = app.cdp; this.sid = page.sid; }
  get alive() { return this.t.sessions.has(this.sid); }
  host(expr) { return this.t.evalIn(RE.host, expr, { root: this.sid }); }
  appFrame(expr) { return this.t.evalIn(RE.app, expr, { root: this.sid }); }
  ed(code) { return this.t.evalIn(RE.editor, `(()=>{try{ return (${code}); }catch(e){ return 'ERR '+(e&&e.message||e); }})()`, { root: this.sid }); }
  state() { return this.host('window.__officeTest.state()'); }
  input(method, params) { return this.cdp.send(method, params, this.sid); }

  async waitReady(ms = 120_000) {
    const t0 = Date.now();
    const st = await waitFor(async () => { const s = await this.state(); return s?.connected || s?.failed ? s : null; }, ms);
    if (!st || st.failed) throw new Error(`편집 창이 준비되지 않았습니다: ${JSON.stringify(st)}`);
    const phase = await waitFor(async () => {
      const ph = await this.appFrame('document.documentElement.dataset.fxPhase || ""');
      if (ph === 'failed') throw new Error(`앱 단계 failed: ${await this.appFrame('document.getElementById("fx-status-text")?.textContent ?? ""')}`);
      return ph === 'ready' ? ph : null;
    }, Math.max(1000, ms - (Date.now() - t0)), 300);
    if (!phase) throw new Error('편집기가 문서를 다 열지 못했습니다 (fxPhase ≠ ready)');
    await waitFor(() => this.t.frame(RE.editor, this.sid), 20_000);
    try { await this.input('Emulation.setFocusEmulationEnabled', { enabled: true }); } catch {}
    return Date.now() - t0;
  }

  // 입력 (CDP만)
  async key(name, mods = '') {
    const MOD = { alt: 1, ctrl: 2, meta: 4, shift: 8 };
    const KEYS = {
      Enter: { code: 'Enter', vk: 13, text: '\r' }, Tab: { code: 'Tab', vk: 9 }, Backspace: { code: 'Backspace', vk: 8 },
      Home: { code: 'Home', vk: 36 }, End: { code: 'End', vk: 35 }, ArrowDown: { code: 'ArrowDown', vk: 40 }, ArrowUp: { code: 'ArrowUp', vk: 38 },
      ArrowLeft: { code: 'ArrowLeft', vk: 37 }, ArrowRight: { code: 'ArrowRight', vk: 39 }, Escape: { code: 'Escape', vk: 27 }, Delete: { code: 'Delete', vk: 46 },
    };
    const m = mods ? mods.split('+').reduce((a, x) => a | MOD[x], 0) : 0;
    let k = KEYS[name];
    if (!k) {
      const up = name.toUpperCase();
      k = { code: /^[a-z]$/i.test(name) ? `Key${up}` : name, vk: up.charCodeAt(0), text: name };
    }
    const withText = k.text && !(m & (MOD.ctrl | MOD.alt | MOD.meta));
    const keyName = m & MOD.ctrl && name.length === 1 ? name.toLowerCase() : name;
    const base = { key: keyName, code: k.code, windowsVirtualKeyCode: k.vk, nativeVirtualKeyCode: k.vk, modifiers: m };
    await this.input('Input.dispatchKeyEvent', { type: withText ? 'keyDown' : 'rawKeyDown', ...base, ...(withText ? { text: k.text, unmodifiedText: k.text } : {}) });
    await this.input('Input.dispatchKeyEvent', { type: 'keyUp', ...base });
    await sleep(70);
  }
  async type(text) { await this.input('Input.insertText', { text }); await sleep(160); }
  /** 한국어 IME 조합 흉내: 조합 단계마다 imeSetComposition, 끝에 확정 */
  async ime(steps, commit) {
    for (const s of steps) { await this.input('Input.imeSetComposition', { text: s, selectionStart: s.length, selectionEnd: s.length }); await sleep(90); }
    await this.input('Input.insertText', { text: commit });
    await sleep(160);
  }
  async click(x, y, count = 1) {
    await this.input('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, buttons: 0 });
    await sleep(50);
    for (let i = 1; i <= count; i++) {
      await this.input('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', buttons: 1, clickCount: i });
      await sleep(50);
      await this.input('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0, clickCount: i });
      await sleep(50);
    }
    await sleep(250);
  }
  async drag(x0, y0, x1, y1) {
    await this.input('Input.dispatchMouseEvent', { type: 'mouseMoved', x: x0, y: y0, buttons: 0 });
    await this.input('Input.dispatchMouseEvent', { type: 'mousePressed', x: x0, y: y0, button: 'left', buttons: 1, clickCount: 1 });
    for (let i = 1; i <= 10; i++) { await this.input('Input.dispatchMouseEvent', { type: 'mouseMoved', x: x0 + ((x1 - x0) * i) / 10, y: y0 + ((y1 - y0) * i) / 10, button: 'left', buttons: 1 }); await sleep(20); }
    await this.input('Input.dispatchMouseEvent', { type: 'mouseReleased', x: x1, y: y1, button: 'left', buttons: 0, clickCount: 1 });
    await sleep(400);
  }
  /** 편집기 틀 안 좌표 → 창 좌표 (host 틀 + 앱의 frameEditor 틀) */
  async origin() {
    const h = await this.host('window.__officeTest.frameRect()');
    const a = await this.appFrame('(()=>{const r=document.querySelector("iframe[name=frameEditor]")?.getBoundingClientRect(); return r ? [r.x, r.y] : null})()');
    return [h.x + (a?.[0] ?? 0), h.y + (a?.[1] ?? 0)];
  }
  async rectOf(sel) {
    const [ox, oy] = await this.origin();
    const r = await this.ed(`(()=>{const e=document.querySelector(${JSON.stringify(sel)}); if(!e) return null; const r=e.getBoundingClientRect(); return [r.x, r.y, r.width, r.height];})()`);
    if (!Array.isArray(r)) throw new Error(`편집기에 ${sel} 없음`);
    return [ox + r[0], oy + r[1], r[2], r[3]];
  }
  /** 편집기의 '새 기능' 풍선 닫기 (화면을 가린다) */
  closeTips() {
    return this.ed(`(()=>{let n=0;for(const b of document.querySelectorAll('.asc-synchronizetip .btn-div, .asc-synchronizetip .close')){ if(b.offsetParent){b.click();n++;} } return n})()`).catch(() => 0);
  }
  find(text) {
    return this.ed(`(()=>{const s=new AscCommon.CSearchSettings(); s.put_Text(${JSON.stringify(text)}); s.put_MatchCase(true); return Asc.editor.asc_findText(s, true);})()`);
  }
  async shot(name) {
    try {
      const r = await this.input('Page.captureScreenshot', { format: 'png' });
      const p = join(this.app.out, 'shots', `${name}.png`);
      mkdirSync(dirname(p), { recursive: true });
      writeFileSync(p, Buffer.from(r.data, 'base64'));
      return p;
    } catch { return null; }
  }
  /** Ctrl+S (main의 before-input-event가 받아 host에 저장을 부탁한다) → 저장이 끝날 때까지 */
  async ctrlS(file, beforeHash, ms = 90_000) {
    const t0 = Date.now();
    await this.key('s', 'ctrl');
    const st = await waitFor(async () => {
      const s = await this.state();
      return s && s.saving === 0 && !s.dirty && s.badge === '저장됨' && sha(file) !== beforeHash ? s : null;
    }, ms, 200);
    return { ok: !!st, ms: Date.now() - t0, state: st ?? await this.state().catch(() => null) };
  }
  async waitDirty(want = true, ms = 15_000) { return !!(await waitFor(async () => (await this.state())?.dirty === want, ms, 200)); }
  async closeByKey() { await this.key('w', 'ctrl'); }
  async waitClosed(ms = 20_000) { return !!(await waitFor(() => !this.alive, ms, 200)); }
}

// ── 고치기 (형식별, CDP 입력) ─────────────────────────────

const WORD_TEXT = { p1: '판정 추가 문단: 굵게 바꾼 한국어 문장.', p2: '판정 두 번째 줄, 보통 글씨 값=7!', row: ['총무팀', '박새롬', '3,300'] };
const CELL = { name: '제주', vals: ['10', '20', '30'], sum: '60', ins: '세종', insVals: ['1', '2'] };
const SLIDE = { title: ' — 판정 편집', newText: '판정이 추가한 슬라이드', ime: '한글 입력 확인' };

async function editWord(e, spec) {
  const [x, y, w, h] = await e.rectOf('#editor_sdk');
  await e.click(x + w / 2, y + h / 3);
  await e.key('End', 'ctrl');
  await e.key('Enter');
  await e.type(WORD_TEXT.p1);
  await e.key('Home', 'shift');
  await e.key('b', 'ctrl');
  await e.key('End');
  await e.key('Enter');
  await e.key('b', 'ctrl');
  await e.ime(['ㅍ', '파', '판'], '판');
  await e.ime(['ㅈ', '저', '정'], '정');
  await e.type(' 두 번째 줄, 보통 글씨 값=7!');
  const out = { anchor: spec.anchor ?? null };
  if (spec.anchor) {
    out.found = await e.find(spec.anchor);
    out.addRow = await e.ed('(Asc.editor.addRowBelow(), "ok")');
    await e.key('ArrowLeft');
    await e.type(WORD_TEXT.row[0]); await e.key('Tab'); await e.type(WORD_TEXT.row[1]); await e.key('Tab'); await e.type(WORD_TEXT.row[2]);
  } else {
    // 표가 없는 문서(새 문서): 끝에 3열 1행 표를 넣고 채운 뒤 마지막 칸에서 Tab → 새 행
    await e.key('End', 'ctrl');
    await e.key('Enter');
    out.table = await e.ed('(Asc.editor.put_Table(3, 1), "ok")');
    await sleep(300);
    await e.type('부서'); await e.key('Tab'); await e.type('담당'); await e.key('Tab'); await e.type('예산');
    await e.key('Tab'); // 마지막 칸에서 Tab → 새 행
    await e.type(WORD_TEXT.row[0]); await e.key('Tab'); await e.type(WORD_TEXT.row[1]); await e.key('Tab'); await e.type(WORD_TEXT.row[2]);
  }
  return out;
}

async function editCell(e, spec) {
  const goto = async (ref) => {
    const [x, y, w, h] = await e.rectOf('#ce-cell-name');
    await e.click(x + w / 2, y + h / 2);
    await e.key('a', 'ctrl'); await e.type(ref); await e.key('Enter'); await sleep(250);
    return e.ed('document.querySelector("#ce-cell-name").value');
  };
  const out = { plan: spec.plan };
  out.at = await goto(`A${spec.plan.add}`);
  await e.type(CELL.name); await e.key('Tab');
  for (const v of CELL.vals) { await e.type(v); await e.key('Tab'); }
  for (const ch of `=SUM(B${spec.plan.add}:D${spec.plan.add})`) await e.type(ch);
  await e.key('Enter');
  await goto(`A${spec.plan.add}`); await e.key('b', 'ctrl');
  await goto(`A${spec.plan.ins}`);
  out.insertRow = await e.ed('(Asc.editor.asc_insertCells(Asc.c_oAscInsertOptions.InsertRows), "ok")');
  await sleep(400);
  await goto(`A${spec.plan.ins}`);
  await e.ime(['ㅅ', '세'], '세'); await e.ime(['ㅈ', '조', '종'], '종');
  await e.key('Tab'); await e.type(CELL.insVals[0]); await e.key('Tab'); await e.type(CELL.insVals[1]); await e.key('Enter');
  return out;
}

async function editSlide(e) {
  const [vx, vy, vw, vh] = await e.rectOf('#id_main_view');
  const empty = [vx + 40, vy + vh - 30];
  const out = { before: await e.ed('Asc.editor.getCountPages()') };
  await e.click(...empty);
  await e.key('Tab'); await e.key('Enter'); await e.key('End', 'ctrl');
  await e.type(SLIDE.title);
  await e.key('Home', 'shift'); await e.key('b', 'ctrl');
  await e.key('Escape'); await e.key('Escape');
  await e.click(...empty);
  await e.key('m', 'ctrl'); await sleep(900);
  out.after = await e.ed('Asc.editor.getCountPages()');
  await e.key('Tab');
  out.placeholder = await e.ed('Asc.editor.getSelectedElements().some((x) => x.get_ObjectType() === Asc.c_oAscTypeSelectElement.Shape)');
  if (out.placeholder === true) {
    await e.key('Enter'); await e.type(SLIDE.newText); await e.key('Escape'); await e.key('Tab'); await e.key('Enter');
  } else {
    await e.key('Escape');
    await e.ed('(Asc.editor.StartAddShape("textRect", true), "ok")');
    await e.drag(vx + vw * 0.3, vy + vh * 0.3, vx + vw * 0.3 + 400, vy + vh * 0.3 + 80);
    await e.type(SLIDE.newText); await e.key('Enter');
  }
  await e.ime(['ㅎ', '하', '한'], '한'); await e.ime(['ㄱ', '그', '글'], '글');
  await e.type(' 입력 확인'); await e.key('Escape'); await e.key('Escape');
  return out;
}

// ── 저장한 파일 확인 (readDocument + XML) ──────────────────

async function zipText(path, name) {
  const { default: JSZip } = await import('jszip');
  const z = await JSZip.loadAsync(readFileSync(path));
  const f = z.file(name);
  return f ? f.async('string') : null;
}
const unxml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** docx: 그 글이 든 문단의 글자 조각이 모두 굵은지 → true | false | null(문단 없음) */
export function docxParagraphBold(xml, text) {
  for (const m of xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) {
    const p = m[0];
    const runs = [...p.matchAll(/<w:r[ >][\s\S]*?<\/w:r>/g)].map((r) => r[0]).filter((r) => /<w:t[ >]/.test(r));
    const all = runs.map((r) => [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((x) => unxml(x[1])).join('')).join('');
    if (!flat(all).includes(flat(text))) continue;
    const bold = (r) => { const b = /<w:b(?: w:val="([^"]*)")?\s*\/>/.exec(r); return !!b && !['0', 'false', 'off'].includes(b[1] ?? ''); };
    return runs.filter((r) => flat(r.replace(/<[^>]+>/g, ''))).every(bold);
  }
  return null;
}

/** xlsx: 첫 시트에서 값이 text인 칸 → { ref, bold, formula } | null */
export async function xlsxCell(path, text) {
  const shared = await zipText(path, 'xl/sharedStrings.xml') ?? '';
  const strings = [...shared.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t(?: [^>]*)?>([^<]*)<\/t>/g)].map((x) => unxml(x[1])).join(''));
  const wbRels = await zipText(path, 'xl/_rels/workbook.xml.rels') ?? '';
  const wb = await zipText(path, 'xl/workbook.xml') ?? '';
  const rid = /<sheet [^>]*r:id="([^"]+)"/.exec(wb)?.[1];
  const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(wbRels);
  const sheetPath = `xl/${(target?.[1] ?? target?.[2] ?? 'worksheets/sheet1.xml').replace(/^\/?xl\//, '').replace(/^\//, '')}`;
  const sheet = await zipText(path, sheetPath) ?? '';
  const styles = await zipText(path, 'xl/styles.xml') ?? '';
  const fonts = [...(/<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)?.[1] ?? '').matchAll(/<font>[\s\S]*?<\/font>|<font\/>|<font [^>]*\/>|<font [^>]*>[\s\S]*?<\/font>/g)].map((m) => m[0]);
  const xfs = [...(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  const cells = [...sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)];
  const byRef = new Map(cells.map((c) => [c[1], c]));
  for (const c of cells) {
    const attrs = c[2]; const body = c[3] ?? '';
    const v = /<v>([^<]*)<\/v>/.exec(body)?.[1];
    let value = null;
    if (/t="s"/.test(attrs) && v != null) value = strings[Number(v)];
    else if (/t="inlineStr"/.test(attrs)) value = [...body.matchAll(/<t(?: [^>]*)?>([^<]*)<\/t>/g)].map((x) => unxml(x[1])).join('');
    else if (/t="str"/.test(attrs) && v != null) value = unxml(v);
    if (value !== text) continue;
    const s = Number(/ s="(\d+)"/.exec(attrs)?.[1] ?? 0);
    const fontId = Number(/fontId="(\d+)"/.exec(xfs[s] ?? '')?.[1] ?? 0);
    const bold = /<b\/>|<b val="(1|true)"\/>/.test(fonts[fontId] ?? '');
    const row = Number(/\d+$/.exec(c[1])[0]);
    const sumCell = byRef.get(`E${row}`);
    const formula = sumCell ? unxml(/<f[^>]*>([^<]*)<\/f>/.exec(sumCell[3] ?? '')?.[1] ?? '') || null : null;
    const sumValue = sumCell ? /<v>([^<]*)<\/v>/.exec(sumCell[3] ?? '')?.[1] ?? null : null;
    return { ref: c[1], row, bold, formula, sumValue };
  }
  return null;
}

/**
 * pptx: 글 조각이 굵은지 → true | false | null(없음).
 * text를 주면 그 글이 든(또는 그 글의 일부인 두 글자 넘는) 조각, 안 주면 첫 슬라이드의 첫 조각 (제목)
 */
export async function pptxRunBold(path, text = null) {
  const { default: JSZip } = await import('jszip');
  const z = await JSZip.loadAsync(readFileSync(path));
  const names = Object.keys(z.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
  for (const name of text == null ? names.slice(0, 1) : names) {
    const xml = await z.file(name).async('string');
    for (const m of xml.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)) {
      const t = flat(unxml(/<a:t>([^<]*)<\/a:t>/.exec(m[1])?.[1] ?? ''));
      const want = flat(text ?? '');
      if (!t || (text != null && !t.includes(want) && !(want.includes(t) && t.length >= 2))) continue;
      return /<a:rPr[^>]* b="(1|true)"/.test(m[1]);
    }
  }
  return null;
}

// ── PowerShell 도우미 (실행 폴더에 써서 쓴다. Windows PowerShell 5.1은 BOM 없는 파일을 ANSI로 읽으므로 ASCII만) ──

export const PS_HELPER = String.raw`# office-e2e helper: UI Automation only (no keystrokes, no mouse input).
# -Action titles            : top-level window titles of the process
# -Action dialog [-ArgB64 b]: message boxes (#32770) of the process -> texts; with a button name, press it (WM_COMMAND BN_CLICKED)
param([int]$ProcId, [string]$Action, [string]$ArgB64 = '', [int]$TimeoutMs = 15000, [string]$Out)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$res = [ordered]@{ ok = $false }
function Done([int]$code) { [System.IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 6), $utf8); exit $code }
try {
  Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
  Add-Type -Namespace OkE2E -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
'@
  $arg = ''
  # ArgB64 = base64('|' + name): never empty, so PowerShell -File does not drop it
  if ($ArgB64) { $arg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ArgB64)).Substring(1) }
  $AE = [System.Windows.Automation.AutomationElement]
  $TS = [System.Windows.Automation.TreeScope]
  $procCond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $ProcId)
  $dlgClass = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, '#32770')
  if ($Action -eq 'titles') {
    $titles = @()
    foreach ($w in $AE::RootElement.FindAll($TS::Children, $procCond)) { if ($w.Current.Name) { $titles += $w.Current.Name } }
    $res.titles = $titles; $res.ok = $true; Done 0
  }
  $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
  do {
    $dialogs = @()
    foreach ($w in $AE::RootElement.FindAll($TS::Children, $procCond)) {
      if ($w.Current.ClassName -eq '#32770') { $dialogs += $w; continue }
      foreach ($d in $w.FindAll($TS::Children, $dlgClass)) { $dialogs += $d }
    }
    foreach ($d in $dialogs) {
      $texts = @(); $buttons = @(); $hit = $null
      foreach ($b in $d.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        $n = $b.Current.Name
        if (-not $n) { continue }
        if ($b.Current.ClassName -eq 'CCPushButton' -or $b.Current.ControlType.ProgrammaticName -eq 'ControlType.Button') { $buttons += $n.Trim(); if ($arg -and $n.Trim() -eq $arg) { $hit = $b } }
        elseif ($texts -notcontains $n) { $texts += $n }
      }
      $res.title = $d.Current.Name; $res.texts = $texts; $res.buttons = $buttons
      if (-not $arg) { $res.ok = $true; Done 0 }
      if ($hit) {
        $h = [IntPtr]$hit.Current.NativeWindowHandle
        if ($h -ne [IntPtr]::Zero) {
          $id = [OkE2E.Win]::GetDlgCtrlID($h)
          [void][OkE2E.Win]::PostMessage([OkE2E.Win]::GetParent($h), 0x111, [IntPtr]($id -band 0xffff), $h)
          $res.clicked = $arg; $res.ok = $true; Done 0
        }
      }
    }
    Start-Sleep -Milliseconds 250
  } while ((Get-Date) -lt $deadline)
  $res.error = 'no dialog'
  Done 1
} catch { $res.error = $_.Exception.Message; Done 1 }
`;

function psHelper(out) {
  const p = join(out, 'office-e2e-helper.ps1');
  if (!existsSync(p)) { mkdirSync(out, { recursive: true }); writeFileSync(p, PS_HELPER, 'ascii'); }
  return p;
}

// Word·Excel·PowerPoint COM으로 문서 만들기 (한국어 글은 UTF-8 JSON으로 넘긴다). 사용자 Office가 떠 있으면 하지 않는다
export const PS_MAKE_OFFICE = String.raw`# office-e2e: make test documents with Word/Excel/PowerPoint COM. Korean text comes from a UTF-8 JSON file.
# Refuses (exit 3) while any Word/Excel/PowerPoint is running (never touches the user's Office); quits/kills only the instance it started.
param([string]$Dir, [string]$JsonPath, [string]$Out, [string]$Only = 'docx,xlsx,pptx')
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$t = [System.IO.File]::ReadAllText($JsonPath, $utf8) | ConvertFrom-Json
$res = [ordered]@{ ok = $false; made = @(); busy = $false; error = $null }
function Done([int]$code) { [System.IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 5), $utf8); exit $code }
function Pids($n) { @(Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
$lock = New-Object System.Threading.Mutex($false, 'Local\OfficeKitCom')
$held = $false
try { $held = $lock.WaitOne(120000) } catch [System.Threading.AbandonedMutexException] { $held = $true }
if (-not $held) { $res.busy = $true; $res.error = 'Local\OfficeKitCom busy'; Done 3 }
try {
  foreach ($n in 'WINWORD', 'EXCEL', 'POWERPNT') { if ((Pids $n).Count) { $res.busy = $true; throw "$n is already running" } }
  $want = $Only.Split(',')
  function Finish($app, [string]$proc, $before) {
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in (Pids $proc)) { if ($before -notcontains $p) { $pr = Get-Process -Id $p -ErrorAction SilentlyContinue; if ($pr -and -not $pr.WaitForExit(15000)) { Stop-Process -Id $p -Force } } }
  }
  if ($want -contains 'docx') {
    $before = Pids 'WINWORD'
    $w = New-Object -ComObject Word.Application
    try {
      $w.Visible = $false; $w.DisplayAlerts = 0
      $doc = $w.Documents.Add()
      $doc.Content.Text = (@($t.wTitle, $t.wLead, $t.wBody, $t.wCaption) -join [char]13)
      $doc.Content.Font.Name = $t.malgun
      $doc.Paragraphs.Item(1).Range.Style = -2
      $doc.Paragraphs.Item(1).Range.Font.Name = $t.malgun
      $doc.Paragraphs.Item(2).Range.Font.Bold = 1
      $doc.Paragraphs.Item(2).Range.Font.Color = 12611584
      $doc.Content.InsertParagraphAfter()
      $rng = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
      $rows = @($t.wTable)
      $tbl = $doc.Tables.Add($rng, $rows.Count, 3)
      $tbl.Borders.Enable = 1
      for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $tbl.Cell($r + 1, $c + 1).Range.Text = [string]$rows[$r][$c] } }
      $tbl.Rows.Item(1).Range.Font.Bold = 1
      $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 15652797
      $doc.Content.InsertAfter($t.wAfter)
      $sec = $doc.Sections.Item(1)
      $sec.Headers.Item(1).Range.Text = $t.header
      [void]$sec.Footers.Item(1).PageNumbers.Add(1)
      $p = Join-Path $Dir 'office.docx'
      $doc.SaveAs2([string]$p, 16)
      $doc.Close($false)
      $res.made += $p
    } finally { Finish $w 'WINWORD' $before }
  }
  if ($want -contains 'xlsx') {
    $before = Pids 'EXCEL'
    $x = New-Object -ComObject Excel.Application
    try {
      $x.Visible = $false; $x.DisplayAlerts = $false
      $wb = $x.Workbooks.Add()
      $ws = $wb.Worksheets.Item(1)
      $ws.Name = $t.sheet1
      $ws.Range('A1').Value2 = $t.xTitle
      $ws.Range('A1:E1').Merge(); $ws.Range('A1').HorizontalAlignment = -4108; $ws.Range('A1').Font.Size = 16; $ws.Range('A1').Font.Bold = $true
      $head = @($t.xHead)
      for ($c = 0; $c -lt 5; $c++) { $ws.Cells.Item(2, $c + 1).Value2 = [string]$head[$c] }
      $data = @($t.xRows)
      for ($r = 0; $r -lt $data.Count; $r++) {
        $ws.Cells.Item($r + 3, 1).Value2 = [string]$data[$r][0]
        for ($c = 1; $c -le 3; $c++) { $ws.Cells.Item($r + 3, $c + 1).Value2 = [double]$data[$r][$c] }
      }
      $last = $data.Count + 2
      $ws.Range("E3:E$last").Formula = '=SUM(B3:D3)'
      $ws.Cells.Item($last + 1, 1).Value2 = $t.sum
      $ws.Range("B$($last + 1):E$($last + 1)").Formula = "=SUM(B3:B$last)"
      $ws.Range('A2:E2').Font.Bold = $true
      $ws.Range('A2:E2').Interior.Color = 15652797
      $ws.Range("A$($last + 1):E$($last + 1)").Font.Bold = $true
      $ws.Range("B3:E$($last + 1)").NumberFormat = '#,##0'
      $ws.Range("A2:E$($last + 1)").Borders.LineStyle = 1
      $ws.Columns('A:E').ColumnWidth = 14
      $ws2 = $wb.Worksheets.Add([Type]::Missing, $ws)
      $ws2.Name = $t.sheet2
      $ws2.Range('A1').Value2 = $t.memo
      $ws.Activate()
      $p = Join-Path $Dir 'office.xlsx'
      $wb.SaveAs([string]$p, 51)
      $wb.Close($false)
      $res.made += $p
    } finally { Finish $x 'EXCEL' $before }
  }
  if ($want -contains 'pptx') {
    $before = Pids 'POWERPNT'
    $pp = New-Object -ComObject PowerPoint.Application
    try {
      $pres = $pp.Presentations.Add(0)
      $s1 = $pres.Slides.Add(1, 1)
      $s1.Shapes.Item(1).TextFrame.TextRange.Text = $t.pTitle
      $s1.Shapes.Item(2).TextFrame.TextRange.Text = $t.pSub
      $s2 = $pres.Slides.Add(2, 2)
      $s2.Shapes.Item(1).TextFrame.TextRange.Text = $t.p2Title
      $s2.Shapes.Item(2).TextFrame.TextRange.Text = (@($t.p2Body) -join [char]13)
      $sh = $s2.Shapes.AddShape(5, 560, 330, 340, 110)
      $sh.Fill.ForeColor.RGB = 13998939
      $sh.TextFrame.TextRange.Text = $t.box
      $sh.TextFrame.TextRange.Font.Name = $t.malgun
      $sh.TextFrame.TextRange.Font.Size = 20
      $sh.TextFrame.TextRange.Font.Bold = 1
      $ar = $s1.Shapes.AddShape(33, 60, 420, 120, 60)
      $ar.Fill.ForeColor.RGB = 3381759
      $p = Join-Path $Dir 'office.pptx'
      $pres.SaveAs([string]$p, 24)
      $pres.Close()
      $res.made += $p
    } finally { Finish $pp 'POWERPNT' $before }
  }
  $res.ok = $true
} catch { $res.error = $_.Exception.Message }
finally { try { $lock.ReleaseMutex() } catch {}; $lock.Dispose() }
if ($res.ok) { Done 0 } elseif ($res.busy) { Done 3 } else { Done 1 }
`;

const OFFICE_TEXT = {
  malgun: '맑은 고딕',
  wTitle: '2026년 하반기 사업 계획', wLead: '굵은 파란 글씨: Word로 만든 시험 문서', wBody: '이 문서는 Word COM으로 만들었습니다. 표 서식·머리글·쪽 번호가 있습니다.',
  wCaption: '부서별 담당자', wTable: [['부서', '담당', '예산'], ['기획팀', '김하늘', '1,000'], ['개발팀', '이바다', '2,500'], ['영업팀', '최유진', '1,800']],
  wAfter: '표 뒤 마지막 문단입니다.', header: 'Office Kit 시험 머리글',
  sheet1: '예산', sheet2: '메모', xTitle: '2027년 부서별 예산 (단위: 만 원)', xHead: ['부서', '인건비', '운영비', '장비', '합계'],
  xRows: [['기획팀', 4200, 800, 300], ['개발팀', 9600, 1200, 2100], ['영업팀', 5100, 1500, 200], ['지원팀', 3800, 600, 150]],
  sum: '합계', memo: '작성자: 김민지',
  pTitle: '한글 문서 도우미 소개', pSub: '2026년 가을 설명회 (PowerPoint로 만듦)', p2Title: '이런 분께 좋습니다',
  p2Body: ['보고서를 자주 고치는 직장인', '표 계산이 많은 회계 담당자'], box: '인터넷 없이 고칩니다',
};

/** COM으로 Office 문서 만들기 → { ok, made: [path], busy, error } */
export function makeOfficeDocs(dir, formats = FORMATS) {
  dir = resolve(dir); // Word SaveAs2는 전체 경로만 받는다
  mkdirSync(dir, { recursive: true });
  const ps1 = join(dir, 'make-office.ps1');
  const json = join(dir, 'make-office.json');
  const out = join(dir, 'make-office.result.json');
  writeFileSync(ps1, PS_MAKE_OFFICE, 'ascii');
  writeFileSync(json, JSON.stringify(OFFICE_TEXT), 'utf8');
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-Dir', dir, '-JsonPath', json, '-Out', out, '-Only', formats.join(',')], { encoding: 'utf8', timeout: 300_000 });
  try { return JSON.parse(readFileSync(out, 'utf8').replace(/^﻿/, '')); } catch { return { ok: false, error: (r.stderr || r.stdout || `exit ${r.status}`).slice(0, 600) }; }
}

/** scripts/dev/strict-open.ps1 (경고 모두 켬, 복구 끔, 감시 타이머) → 결과 JSON */
export function strictOpen(file, out, timeoutSec = 90) {
  file = resolve(file);
  const json = `${resolve(out)}.strict.json`;
  rmSync(json, { force: true });
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(ROOT, 'scripts', 'dev', 'strict-open.ps1'), '-Path', file, '-Out', json, '-TimeoutSec', String(timeoutSec)], { encoding: 'utf8', timeout: (timeoutSec + 120) * 1000 });
  try { return { exit: r.status, ...JSON.parse(readFileSync(json, 'utf8').replace(/^﻿/, '')) }; } catch { return { exit: r.status, ok: false, error: (r.stderr || r.stdout || '').slice(0, 600) }; }
}

/** Office가 깔려 있는지 (COM 등록) */
export function officeInstalled() {
  const r = spawnSync('powershell', ['-NoProfile', '-Command', "foreach($p in 'Word.Application','Excel.Application','PowerPoint.Application'){ if([type]::GetTypeFromProgID($p)){'1'}else{'0'} }"], { encoding: 'utf8', timeout: 30_000 });
  const [word, excel, ppt] = (r.stdout ?? '').trim().split(/\s+/).map((x) => x === '1');
  return { word: !!word, excel: !!excel, ppt: !!ppt };
}

// 일부러 밖으로 보내 보기 (fetch·WebSocket·그림) → 모두 실패해야 한다. 이 주소는 '밖으로 나간 요청' 집계에서 뺀다
export const OUT_PROBE_MARK = 'officekit-e2e-probe';
const OUT_PROBE = `(async()=>{const u='https://example.com/?${OUT_PROBE_MARK}';const o={};`
  + `try{await fetch(u,{mode:'no-cors',cache:'no-store'});o.fetch='열림'}catch(e){o.fetch='막힘'}`
  + `o.ws=await new Promise(r=>{try{const w=new WebSocket('wss://example.com/${OUT_PROBE_MARK}');w.onopen=()=>{w.close();r('열림')};w.onerror=()=>r('막힘');setTimeout(()=>r('시간 초과'),8000)}catch(e){r('막힘')}});`
  + `o.img=await new Promise(r=>{const i=new Image();i.onload=()=>r('열림');i.onerror=()=>r('막힘');i.src=u+'&img';setTimeout(()=>r('시간 초과'),8000)});return o})()`;

const RTC_PROBE = `(async()=>{const o={};for(const [n,cfg] of [['stun',{iceServers:[{urls:'stun:stun.l.google.com:19302'}]}],['none',{}]]){try{const pc=new RTCPeerConnection(cfg);pc.createDataChannel('x');await pc.setLocalDescription(await pc.createOffer());o[n]=await new Promise(r=>{const c=[];pc.onicecandidate=e=>{if(e.candidate)c.push(e.candidate.type||'?');else r({done:true,c})};setTimeout(()=>r({done:false,c}),6000)});pc.close()}catch(e){o[n]={threw:String(e&&e.message||e),c:[]}}};return o})()`;

// ── 실행 ────────────────────────────────────────────────

const JUDGE_MD = `# 판정용 시험 문서
**굵은 첫 줄: 오프라인 편집기 판정**
이 문서는 Office Kit 편집기 판정을 위한 한국어 시험 문서입니다.
- 첫째: 한글 입력과 저장
- 둘째: 표와 서식 유지
| 부서 | 담당 | 예산 |
|---|---|---|
| 기획팀 | 김하늘 | 1,000 |
| 개발팀 | 이바다 | 2,500 |
담당자: {{담당자}}
---page---
## 두 번째 쪽
마지막 문단입니다.`;

/**
 * e2e 한 번 → 보고서 { ok, checks: [{ group, name, ok, info }], docs, flows, net, rtc, ... }
 * opts: { out, formats, methods, com (true|false|'auto'), flows, officeDocs, keep, log }
 */
export async function runOfficeE2E(opts = {}) {
  const out = resolve(opts.out ?? join(ROOT, 'tmp', 'office-e2e', new Date().toISOString().replace(/[:.]/g, '-')));
  const formats = (opts.formats ?? FORMATS).filter((f) => FORMATS.includes(f));
  const methods = (opts.methods ?? METHODS).filter((m) => METHODS.includes(m));
  const log = opts.log ?? ((...a) => console.log(...a));
  mkdirSync(join(out, 'shots'), { recursive: true });
  const docsDir = join(out, 'docs');
  mkdirSync(docsDir, { recursive: true });
  const report = { ok: false, out, startedAt: new Date().toISOString(), checks: [], docs: [], flows: {}, net: null, rtc: {}, notes: [] };
  const check = (group, name, ok, info = '') => {
    report.checks.push({ group, name, ok: !!ok, info });
    log(`${ok ? 'PASS' : 'FAIL'}  [${group}] ${name}${info === '' ? '' : `  — ${typeof info === 'string' ? info : JSON.stringify(info)}`.slice(0, 700)}`);
    return !!ok;
  };
  const note = (s) => { report.notes.push(s); log(`NOTE  ${s}`); };
  const core = await import(pathToFileURL(join(ROOT, 'src', 'core', 'index.mjs')).href);
  const create = await import(pathToFileURL(join(ROOT, 'src', 'core', 'create.mjs')).href);
  const { draftPathFor } = await import(pathToFileURL(join(ROOT, 'src', 'app', 'editor-window.mjs')).href);

  if (!existsSync(join(ROOT, 'vendor', 'office-editor', 'VERSION.json'))) throw new Error('vendor/office-editor 묶음이 없습니다 (npm run office-editor:build)');

  // COM: true = 꼭, false = 안 함, 'auto' = Office가 있으면
  const office = opts.com === false ? { word: false, excel: false, ppt: false } : officeInstalled();
  const comFor = (ext) => opts.com !== false && office[{ docx: 'word', xlsx: 'excel', pptx: 'ppt' }[ext]];
  report.office = office;

  // 1) 문서 준비
  const specs = [];
  for (const ext of formats) {
    for (const method of methods) {
      const id = `${method}-${ext}`;
      const spec = { id, ext, method, path: null, anchor: null, plan: null, skip: null };
      if (method === 'create') {
        spec.path = join(docsDir, `판정 문서.${ext}`);
        await create.createDocument(spec.path, JUDGE_MD);
        spec.anchor = '이바다';
      } else if (method === 'blank') {
        spec.path = join(docsDir, `새 문서.${ext}`); // 앱의 '새 문서'(doc:new edit:true)가 만든다
      } else {
        spec.path = join(docsDir, `office.${ext}`);
        spec.anchor = '최유진';
      }
      specs.push(spec);
    }
  }
  const officeSpecs = specs.filter((s) => s.method === 'office');
  if (officeSpecs.length) {
    if (opts.officeDocs) {
      for (const s of officeSpecs) {
        const src = join(resolve(opts.officeDocs), `office.${s.ext}`);
        if (existsSync(src)) copyFileSync(src, s.path); else s.skip = `Office로 만든 파일 없음: ${src}`;
      }
    } else {
      const need = officeSpecs.filter((s) => comFor(s.ext)).map((s) => s.ext);
      for (const s of officeSpecs) if (!need.includes(s.ext)) s.skip = 'COM을 쓰지 않음 (Office 없음 또는 --no-com)';
      if (need.length) {
        const made = makeOfficeDocs(docsDir, need);
        report.officeMade = made;
        for (const s of officeSpecs) if (!s.skip && !existsSync(s.path)) s.skip = `COM으로 만들지 못함: ${made.error ?? '?'}`;
      }
    }
  }

  const ud = join(out, 'userdata');
  rmSync(ud, { recursive: true, force: true });
  mkdirSync(ud, { recursive: true });
  // 주 창 미리보기는 글자만 (Word·PowerPoint를 띄우지 않게 → COM 엄격 열기와 겹치지 않음)
  writeFileSync(join(ud, 'settings.json'), JSON.stringify({ previewMode: 'text' }));

  const app = new App({ out, log });
  report.app = { started: false };
  try {
    const t0 = Date.now();
    await app.start({ ud });
    report.app = { started: true, ms: Date.now() - t0, browser: app.browser };
    const init = await app.api('init');
    check('app', 'app:init: Office 편집 창 사용 가능', init?.editor?.kinds?.office?.available === true, init?.editor?.kinds?.office);

    /** 문서를 주 창에 열고 편집 창을 띄운다 → Editor */
    const openEditor = async (spec) => {
      const before = new Set(app.t.pages(RE.host).map((p) => p.sid));
      let r;
      if (spec.method === 'blank' && !existsSync(spec.path)) {
        r = await app.api('doc.create', { start: 'blank', folder: dirname(spec.path), name: basename(spec.path).replace(/\.[^.]+$/, ''), format: spec.ext, edit: true });
        spec.created = { doc: r?.doc?.path ?? null, editor: r?.editor ?? null, editorError: r?.editorError ?? null };
        if (!r?.editor) throw new Error(`새 문서 편집 창을 열지 못했습니다: ${JSON.stringify(r?.editorError)}`);
        r = r.editor;
      } else {
        await app.api('doc.open', spec.path);
        r = await app.api('doc.edit');
      }
      const page = await waitFor(() => app.t.pages(RE.host).find((p) => !before.has(p.sid)), 30_000);
      if (!page) throw new Error(`편집 창(host 페이지)이 뜨지 않았습니다: ${JSON.stringify([...app.t.sessions.values()].map((x) => [x.type, x.url]))}`);
      return { editor: new Editor(app, page), result: r };
    };

    // 2) 형식 3 × 만든 방법 3
    const rtcDone = new Set();
    for (const spec of specs) {
      const d = { id: spec.id, ext: spec.ext, method: spec.method, path: spec.path };
      report.docs.push(d);
      const g = spec.id;
      if (spec.skip) { d.skipped = spec.skip; note(`${g}: 건너뜀 — ${spec.skip}`); continue; }
      let e = null;
      try {
        const original = spec.method === 'blank' ? null : sha(spec.path);
        if (spec.ext === 'xlsx') {
          const rows = spec.method === 'blank' ? 0 : ((await core.readDocument(spec.path)).sheets?.[0]?.rows?.length ?? 0);
          spec.plan = { add: Math.max(rows, 1) + 2, ins: rows >= 3 ? 3 : 2 };
        }
        const pptBefore = spec.ext === 'pptx' && spec.method !== 'blank' ? (await core.readDocument(spec.path)).slides?.length : null;
        // 제목 줄을 골라 Ctrl+B → 굵기가 뒤집힌다 (createDocument 제목은 원래 굵음)
        const titleBold = spec.ext === 'pptx' && spec.method !== 'blank' ? await pptxRunBold(spec.path) : null;
        const netStart = app.t.requests.length;
        const t0 = Date.now();
        const o = await openEditor(spec);
        e = o.editor;
        check(g, '편집 창 열림 (kind office)', o.result?.kind === 'office', o.result);
        d.openMs = await e.waitReady();
        d.openTotalMs = Date.now() - t0;
        const original2 = original ?? sha(spec.path);
        if (spec.method === 'blank') {
          check(g, "'새 문서'(doc:new edit:true)로 만든 파일", existsSync(spec.path) && spec.created?.doc && resolve(spec.created.doc) === resolve(spec.path), spec.created);
        }
        const st0 = await e.state();
        check(g, '열기: 연결됨·저장됨', st0.connected && !st0.failed && st0.badge === '저장됨' && !st0.dirty, { openMs: d.openMs, badge: st0.badge, title: st0.title });
        check(g, 'host 제목에 창 이름 (ONLYOFFICE 없음)', st0.title?.endsWith(` - ${LABEL[spec.ext]}`) && !/onlyoffice/i.test(st0.title), st0.title);
        await sleep(1500);
        await e.closeTips();
        d.shots = { open: await e.shot(`${g}-1-open`) };

        // 고치기
        const tEdit = Date.now();
        d.edit = spec.ext === 'docx' ? await editWord(e, spec) : spec.ext === 'xlsx' ? await editCell(e, spec) : await editSlide(e);
        d.editMs = Date.now() - tEdit;
        await sleep(700);
        await e.closeTips();
        d.shots.edited = await e.shot(`${g}-2-edited`);
        const dirty = await e.waitDirty(true);
        const st1 = await e.state();
        check(g, '고치면 저장 안 됨 표시', dirty && st1.badge === '저장 안 됨', { badge: st1.badge, edit: d.edit });
        const titles = app.ps('titles')?.titles ?? [];
        const winTitle = titles.find((x) => x.includes(`${basename(spec.path)} - ${LABEL[spec.ext]}`)); // 주 창('<파일> - Office Kit')은 빼고
        check(g, "창 제목: '* <파일> - " + LABEL[spec.ext] + " (Office Kit)'", winTitle === `* ${basename(spec.path)} - ${LABEL[spec.ext]} (Office Kit)`, winTitle ?? titles);
        const ai = await app.apiRaw('ai.send', { prompt: '시험: 보내지 않아야 함' });
        check(g, 'AI 보내기 막힘 (EDITOR_DIRTY)', ai?.ok === false && ai?.code === 'EDITOR_DIRTY', ai);
        const stEditor = await app.main('document.getElementById("st-editor")?.textContent ?? ""');
        check(g, '주 창 상태 막대: 편집 창 열림 · 저장 안 됨', stEditor.includes(`${LABEL[spec.ext]} 창 열림 · 저장 안 됨`), stEditor);

        // Ctrl+S
        const draft = draftPathFor(ud, spec.path);
        const s = await e.ctrlS(spec.path, original2);
        d.save = { ms: s.ms, ok: s.ok, state: s.state };
        check(g, 'Ctrl+S → 파일 저장·저장됨', s.ok, { ms: s.ms, state: s.state });
        check(g, '정상 저장 뒤 복구 사본 없음', !existsSync(draft), draft);
        d.shots.saved = await e.shot(`${g}-3-saved`);
        const titles2 = app.ps('titles')?.titles ?? [];
        check(g, '저장 뒤 창 제목에 * 없음', titles2.includes(`${basename(spec.path)} - ${LABEL[spec.ext]} (Office Kit)`), titles2);

        // readDocument(engine:'file')
        const read = await core.readDocument(spec.path);
        d.read = { text: String(read.text ?? '').slice(0, 3000) };
        if (spec.ext === 'docx') {
          const text = flat(read.text);
          check(g, 'readDocument: 추가 문단 두 개', text.includes(flat(WORD_TEXT.p1)) && text.includes(flat(WORD_TEXT.p2)), d.read.text.slice(-400));
          const rows = (read.tables ?? []).flatMap((tb) => tb.rows ?? tb);
          const hasRow = rows.some((r) => Array.isArray(r) && WORD_TEXT.row.every((v, i) => flat(r[i]) === flat(v)));
          check(g, 'readDocument: 표 새 행 (총무팀 | 박새롬 | 3,300)', hasRow, rows.slice(-3));
          const xml = await zipText(spec.path, 'word/document.xml');
          check(g, '굵게: 첫 추가 문단은 굵게', docxParagraphBold(xml, WORD_TEXT.p1) === true);
          check(g, '굵게: 둘째 추가 문단은 보통', docxParagraphBold(xml, WORD_TEXT.p2) === false);
          const restarts = (xml.match(/<w:vMerge w:val="restart"\/>/g) ?? []).length;
          const continues = (xml.match(/<w:vMerge\/>|<w:vMerge w:val="continue"\/>/g) ?? []).length;
          check(g, '저장 뒤 고치기: 혼자 남은 vMerge restart 없음', restarts === 0 || continues > 0, { restarts, continues });
        } else if (spec.ext === 'xlsx') {
          const rows = (read.sheets?.[0]?.rows ?? []).map((r) => r.map((v) => String(v ?? '')));
          const jeju = rows.find((r) => r[0] === CELL.name);
          check(g, 'readDocument: 제주 10 20 30 = 60', !!jeju && jeju.slice(1, 5).join(',') === [...CELL.vals, CELL.sum].join(','), jeju);
          const sejong = rows.find((r) => r[0] === CELL.ins);
          check(g, 'readDocument: IME로 넣은 세종 1 2', !!sejong && sejong.slice(1, 3).join(',') === CELL.insVals.join(','), sejong);
          const cell = await xlsxCell(spec.path, CELL.name);
          const row = spec.plan.add + 1;
          check(g, `수식이 행 삽입 뒤 밀림 (=SUM(B${row}:D${row}))`, cell?.formula === `SUM(B${row}:D${row})` && cell?.row === row, cell);
          check(g, '굵게: 제주 칸', cell?.bold === true, cell);
          check(g, '행 삽입: 세종이 삽입한 행에', rows[spec.plan.ins - 1]?.[0] === CELL.ins, rows[spec.plan.ins - 1]);
        } else {
          const slides = read.slides ?? [];
          d.slides = slides.length;
          check(g, '새 슬라이드 (Ctrl+M)', d.edit.after === d.edit.before + 1 && (pptBefore == null || slides.length === pptBefore + 1), { before: d.edit.before, after: d.edit.after, file: slides.length, pptBefore });
          const text = flat(read.text);
          check(g, 'readDocument: 제목 뒤 추가 글·새 슬라이드 글·IME 글', text.includes(flat(SLIDE.title)) && text.includes(flat(SLIDE.newText)) && text.includes(flat(SLIDE.ime)), d.read.text.slice(0, 400));
          const nowBold = await pptxRunBold(spec.path, SLIDE.title);
          check(g, `Ctrl+B: 제목 줄 굵기가 뒤집힘 (${titleBold === true ? '굵음 → 보통' : '보통 → 굵음'})`, nowBold === !(titleBold === true), { before: titleBold, after: nowBold });
        }

        // COM 엄격 열기
        if (comFor(spec.ext)) {
          const r = strictOpen(spec.path, join(out, 'docs', spec.id));
          d.com = r;
          check(g, 'COM 엄격 열기: 경고·복구 창 없이 열림', r.ok === true && r.opened === true && !r.hung, { ok: r.ok, hung: r.hung, error: r.error, prompt: r.prompt, openMs: r.openMs });
          if (r.ok && spec.ext === 'docx') {
            const p1 = (r.paras ?? []).find((p) => flat(p.text).includes(flat(WORD_TEXT.p1)));
            const p2 = (r.paras ?? []).find((p) => flat(p.text).includes(flat(WORD_TEXT.p2)));
            check(g, 'Word: 추가 문단 굵게 / 둘째 보통', p1 && Number(p1.bold) !== 0 && p2 && Number(p2.bold) === 0, { p1, p2 });
            const t = (r.tables ?? []).find((x) => (x.cells ?? []).some((c) => c.endsWith(`=${WORD_TEXT.row[0]}`)));
            check(g, 'Word: 표 새 행 + Rows.Item 됨 (vMerge 고치기)', !!t && t.rowsItem === true, t ? { rows: t.rows, uniform: t.uniform, rowsItem: t.rowsItem, err: t.rowsItemError } : r.tables);
          } else if (r.ok && spec.ext === 'xlsx') {
            const rows = r.sheets?.[0]?.rows ?? [];
            const j = rows.find((x) => x.includes(CELL.name));
            check(g, 'Excel: 제주 행 굵게 + 60 {=SUM}', !!j && j.includes(`**${CELL.name}**`) && /60 \{=SUM\(B\d+:D\d+\)\}/.test(j), j ?? rows);
          } else if (r.ok && spec.ext === 'pptx') {
            const all = (r.slides ?? []).flatMap((x) => (x.shapes ?? []).map((sh) => sh.text ?? '')).join('\n');
            check(g, 'PowerPoint: 슬라이드 수·추가 글', (r.slides ?? []).length === d.slides && flat(all).includes(flat(SLIDE.newText)), { slides: (r.slides ?? []).length, text: all.slice(0, 300) });
          }
        }

        // WebRTC (형식마다 한 번, host·앱·편집기 세 틀)
        if (!rtcDone.has(spec.ext)) {
          rtcDone.add(spec.ext);
          const rtc = {};
          for (const [n, re] of [['host', RE.host], ['app', RE.app], ['editor', RE.editor]]) {
            try { rtc[n] = await app.t.evalIn(re, RTC_PROBE, { root: e.sid, timeoutMs: 30_000 }); } catch (err) { rtc[n] = { error: err.message }; }
          }
          report.rtc[spec.ext] = rtc;
          const cands = Object.values(rtc).flatMap((x) => [...(x?.stun?.c ?? []), ...(x?.none?.c ?? [])]);
          check(g, 'WebRTC: host·앱·편집기 틀 모두 후보 0개 (STUN 서버를 줘도)', cands.length === 0 && Object.values(rtc).every((x) => !x?.error), rtc);
          const probe = {};
          for (const [n, re] of [['host', RE.host], ['app', RE.app], ['editor', RE.editor]]) {
            try { probe[n] = await app.t.evalIn(re, OUT_PROBE, { root: e.sid, timeoutMs: 30_000 }); } catch (err) { probe[n] = { error: err.message }; }
          }
          report.outProbe = { ...(report.outProbe ?? {}), [spec.ext]: probe };
          check(g, '일부러 밖으로 보내기 (fetch·WebSocket·그림): 세 틀 모두 막힘', Object.values(probe).every((x) => x && !x.error && Object.values(x).every((v) => v === '막힘')), probe);
        }

        // 이 문서 동안 밖으로 나간 요청
        const reqs = app.t.requests.slice(netStart).filter((r) => RE.host.test(r.rootUrl ?? '') && !r.url.includes(OUT_PROBE_MARK));
        const ext = reqs.filter((r) => !r.local);
        d.net = { requests: reqs.length, external: ext.map((r) => r.url).slice(0, 20), dead: reqs.filter((r) => /^https?:\/\/127\.0\.0\.1:9\//.test(r.url)).length };
        check(g, '네트워크: 밖으로 나간 요청 0건', reqs.length > 50 && ext.length === 0, d.net);

        // 닫기 (저장했으므로 묻지 않고 닫힘)
        await e.closeByKey();
        const closed = await e.waitClosed();
        if (!closed) { const dlg = app.ps('dialog', '저장 안 함', 3000); note(`${g}: 닫기 확인창이 떴음 ${JSON.stringify(dlg)}`); await e.waitClosed(); }
        check(g, '저장한 뒤 Ctrl+W → 묻지 않고 닫힘', closed);
      } catch (err) {
        d.error = String(err?.stack ?? err);
        check(g, '흐름이 끝까지 돎', false, d.error.slice(0, 600));
        if (e) { try { d.shots = { ...(d.shots ?? {}), error: await e.shot(`${g}-9-error`) }; } catch {} }
        if (e?.alive) { await e.closeByKey().catch(() => {}); if (!(await e.waitClosed(4000))) { app.ps('dialog', '저장 안 함', 4000); await e.waitClosed(8000); } }
      }
    }

    // 3) 흐름 (자동 저장·바깥 변경·멈춤 뒤 복구·원래대로·닫기 확인)
    if (opts.flows !== false) await runFlows({ app, report, check, ud, core, create, draftPathFor, openEditor });

    // 4) 모든 편집 창 요청 모아 보기
    const all = app.t.requests.filter((r) => RE.host.test(r.rootUrl ?? '') && !r.url.includes(OUT_PROBE_MARK));
    const probes = app.t.requests.filter((r) => r.url.includes(OUT_PROBE_MARK));
    report.net = {
      editorRequests: all.length,
      external: all.filter((r) => !r.local).map((r) => r.url),
      dead: [...new Set(all.filter((r) => /^https?:\/\/127\.0\.0\.1:9\//.test(r.url)).map((r) => r.url))],
      failed: all.filter((r) => r.failed && r.failed !== 'net::ERR_ABORTED').map((r) => `${r.failed} ${r.url}`).slice(0, 30),
      mainWindowExternal: app.t.requests.filter((r) => RE.main.test(r.rootUrl ?? '') && !r.local).map((r) => r.url),
      // 일부러 보낸 것이 네트워크 층까지 갔다면 (CSP를 지나) 세션이 막았는지
      probes: probes.map((r) => ({ url: r.url, type: r.type, failed: r.failed ?? null, blocked: r.blocked ?? null })),
    };
    // loadingFailed가 왔으면 실패 (CSP가 막으면 errorText는 비고 blockedReason='csp'). WebSocket은 위 틀 안 결과('막힘')로 본다
    check('net', '일부러 보낸 요청이 네트워크 층에 닿았으면 모두 실패', probes.every((r) => r.blocked || typeof r.failed === 'string' || r.type === 'WebSocket'),
      [...new Set(report.net.probes.map((r) => `${r.type} ${r.blocked ?? r.failed ?? '실패 안 함'}`))]);
    check('net', '편집 창 전체: 밖으로 나간 요청 0건', report.net.external.length === 0 && all.length > 0, { requests: all.length, external: report.net.external.slice(0, 10) });
    const blockedLog = app.stdout.split(/\r?\n/).filter((l) => l.includes('막은 요청') && !l.includes(OUT_PROBE_MARK));
    report.net.blockedBySession = blockedLog.slice(0, 20);
    check('net', '세션이 막은 요청 0건 (main 로그)', blockedLog.length === 0, blockedLog.slice(0, 5));
    report.exceptions = app.t.exceptions.filter((x) => RE.host.test(app.t.sessions.get(x.root)?.url ?? '') || /\/office\//.test(x.url ?? '')).slice(0, 40);
  } finally {
    report.log = app.stdout.split(/\r?\n/).filter((l) => /\[editor\]|error|오류|실패/i.test(l)).slice(-200);
    await app.stop();
    // DNS: 일부러 보낸 example.com 등도 시스템 DNS로 나가지 않았어야 한다 (host-resolver-rules → 풀지 않음)
    if (app.netlog) {
      const dns = app.netlogDns();
      report.net = { ...(report.net ?? {}), dns };
      check('net', 'DNS: 시스템 DNS 조회(DNS_TRANSACTION) 0건 — 이 PC 밖 이름은 풀지 않는다', dns.read && dns.transactions === 0, dns);
    }
    report.finishedAt = new Date().toISOString();
    report.ok = report.checks.length > 0 && report.checks.every((c) => c.ok);
    writeFileSync(join(out, 'report.json'), JSON.stringify(report, null, 2));
    if (!opts.keep) rmSync(ud, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
  return report;
}

/**
 * 편집 창 흐름: 바깥 변경(깨끗할 때·고쳤을 때) → 자동 저장 → 멈춤 뒤 복구 → 원래대로 → 닫기 확인.
 * host·main 쪽 흐름이라 종류마다 같다 → 문서(docx) 하나로 (편집기 안 글 찾기도 docx API)
 */
async function runFlows({ app, report, check, ud, core, create, draftPathFor, openEditor }) {
  const ext = 'docx';
  const g = `flow-${ext}`;
  const f = report.flows;
  const dir = join(report.out, 'docs', 'flow');
  mkdirSync(dir, { recursive: true });
  const path = join(dir, `흐름 시험.${ext}`);
  const md = (marker) => `# 흐름 시험 문서\n${marker}\n\n| 부서 | 담당 | 예산 |\n|---|---|---|\n| 기획팀 | 김하늘 | 1,000 |`;
  await create.createDocument(path, md('처음 문단입니다.'));
  const original = sha(path);
  const draft = draftPathFor(ud, path);
  let e = null;
  const typeAtEnd = async (text) => {
    const [x, y, w, h] = await e.rectOf('#editor_sdk');
    await e.click(x + w / 2, y + h / 3);
    await e.key('End', 'ctrl'); await e.key('Enter'); await e.type(text);
  };
  const has = async (text) => { const n = await e.find(text); return typeof n === 'number' ? n > 0 : n === true; };
  const writeOutside = async (marker) => {
    const tmp = join(dir, `바깥.${ext}`);
    await create.createDocument(tmp, md(marker));
    copyFileSync(tmp, path); // 다른 프로그램이 덮어쓴 것처럼 (내용만 바꿈)
    rmSync(tmp, { force: true });
  };
  try {
    const o = await openEditor({ method: 'create', ext, path });
    e = o.editor;
    await e.waitReady();
    await sleep(1200);
    await e.closeTips();

    // 바깥 변경 (고친 것 없음) → 저절로 다시 불러옴
    const appCtx1 = app.t.frame(RE.app, e.sid)?.ctx;
    await writeOutside('바깥에서 바꾼 첫 문단');
    const reloaded = await waitFor(async () => {
      const ctx = app.t.frame(RE.app, e.sid)?.ctx;
      if (!ctx || ctx === appCtx1) return false;
      const ph = await e.appFrame('document.documentElement.dataset.fxPhase || ""');
      return ph === 'ready' && app.t.frame(RE.editor, e.sid) && await has('바깥에서 바꾼 첫 문단');
    }, 60_000, 400);
    const st1 = await e.state();
    f.externalClean = { reloaded: !!reloaded, state: st1 };
    check(g, '바깥 변경(고친 것 없음) → 띠 없이 저절로 다시 불러옴', !!reloaded && !st1.banner && !st1.dirty, f.externalClean);
    await sleep(1000);
    await e.closeTips();

    // 고친 뒤 바깥 변경 → 띠 '다시 불러오기 / 내 편집 유지'
    await typeAtEnd('내 편집 유지 확인');
    await e.waitDirty(true);
    await writeOutside('바깥에서 바꾼 둘째 문단');
    const banner = await waitFor(async () => { const s = await e.state(); return s.banner === 'changed' ? s : null; }, 15_000);
    f.externalDirty = { banner: banner?.bannerText ?? null };
    check(g, '고친 뒤 바깥 변경 → 띠 (다시 불러오기 / 내 편집 유지)', !!banner && banner.bannerText.includes('다른 곳에서 이 문서가 바뀌었습니다'), banner?.bannerText);
    await e.shot(`${g}-1-changed-banner`);
    const kept = await e.host('window.__officeTest.bannerAction("내 편집 유지")');
    await sleep(300);
    const st2 = await e.state();
    check(g, "'내 편집 유지' → 띠 닫힘, 고친 내용 그대로", kept && !st2.banner && st2.dirty && await has('내 편집 유지 확인'), st2);

    // 자동 저장 (앱 틀의 시계를 11분 앞으로 → 30초 안에 tick이 'auto' 저장) → 복구용 사본만
    const fileBefore = sha(path);
    const mtimeBefore = mtime(path);
    await e.appFrame('(()=>{ if (!Date.__e2eShift) { const real = Date.now.bind(Date); Date.__e2eShift = 11*60*1000; Date.now = () => real() + Date.__e2eShift; } return true; })()');
    const drafted = await waitFor(() => existsSync(draft), 50_000, 500);
    await sleep(800);
    const st3 = await e.state();
    f.autosave = { draft: drafted ? draft : null, fileUnchanged: sha(path) === fileBefore && mtime(path) === mtimeBefore, state: st3 };
    check(g, '자동 저장 → userData/office-drafts 사본에만 씀', !!drafted, draft);
    check(g, '자동 저장이 문서 파일을 덮어쓰지 않음', f.autosave.fileUnchanged, { before: fileBefore, after: sha(path) });
    check(g, '자동 저장 뒤에도 저장 안 됨 (autoPending)', st3.dirty && st3.autoPending && st3.badge === '저장 안 됨', st3);
    if (drafted) {
      const dr = await core.readDocument(draft);
      check(g, '복구용 사본에 고친 내용', flat(dr.text).includes(flat('내 편집 유지 확인')), String(dr.text).slice(0, 300));
    }
    // 자동 저장 뒤 앱은 '바뀜 없음'이지만 Ctrl+S는 파일에 써야 한다 (사본만 있고 파일에는 없는 상태를 남기지 않게)
    const sAfterAuto = await e.ctrlS(path, sha(path));
    const rdAuto = await core.readDocument(path);
    const stAuto = await e.state();
    f.saveAfterAuto = { ok: sAfterAuto.ok, draft: existsSync(draft), state: stAuto };
    check(g, '자동 저장 뒤 Ctrl+S → 파일에 들어가고 사본 지워짐, 저장됨', sAfterAuto.ok && flat(rdAuto.text).includes(flat('내 편집 유지 확인')) && !existsSync(draft) && !stAuto.autoPending, f.saveAfterAuto);

    // 화면 멈춤(렌더러 죽음) → 안내 → 다시 열림 → 복구 띠 → 복구 → 저장
    try {
      await typeAtEnd('멈춤 전 편집');
      await e.waitDirty(true);
      await e.appFrame('(()=>{ Date.__e2eShift += 11*60*1000; return Date.__e2eShift; })()');
      const drafted2 = await waitFor(() => existsSync(draft), 50_000, 500);
      check(g, '다시 고친 뒤 자동 저장 → 새 복구 사본', !!drafted2, draft);
      await e.input('Page.crash').catch(() => null);
      const dlg = app.ps('dialog', '확인', 20_000);
      f.crash = { dialog: dlg };
      check(g, "멈춤 안내: '편집 창이 예기치 않게 멈췄습니다'", dlg?.ok && (dlg.texts ?? []).some((x) => x.includes('편집 창이 예기치 않게 멈췄습니다')), dlg);
      const offer = await waitFor(async () => { const s = await e.state(); return s?.banner === 'draft' ? s : null; }, 60_000, 400);
      f.crash.offer = offer?.bannerText ?? null;
      check(g, "다시 열면 띠 '자동 저장된 편집 내용이 있습니다 — 복구 / 버리기'", !!offer && offer.bannerText.includes('자동 저장된 편집 내용이 있습니다'), offer);
      await e.shot(`${g}-2-draft-offer`);
      await waitFor(async () => (await e.appFrame('document.documentElement.dataset.fxPhase || ""')) === 'ready', 60_000, 400);
      await e.host('window.__officeTest.bannerAction("복구")');
      await sleep(500);
      await e.waitReady();
      await sleep(1200);
      const st4 = await e.state();
      check(g, "'복구' → 사본 내용으로 다시 열림, 저장 안 됨", st4.restoredDraft && st4.dirty && await has('멈춤 전 편집'), st4);
      const s = await e.ctrlS(path, sha(path));
      const rd = await core.readDocument(path);
      check(g, '복구한 내용을 Ctrl+S → 파일에 들어가고 사본 지워짐', s.ok && flat(rd.text).includes(flat('멈춤 전 편집')) && !existsSync(draft), { save: s.ok, draft: existsSync(draft) });
    } catch (err) {
      check(g, '멈춤 뒤 복구 흐름', false, String(err?.message ?? err));
    }

    // 원래대로 (주 창) — 고친 것 없을 때: 편집 창이 저절로 다시 불러옴, 파일은 처음 연 상태 버전 (버전 목록에서 고름)
    const appCtx2 = app.t.frame(RE.app, e.sid)?.ctx;
    const openedVersion = async () => (await app.apiRaw('doc.versions'))?.value?.versions?.find((v) => v.reason === 'open')?.file;
    const rr = await app.apiRaw('doc.restore', { file: await openedVersion() });
    const back = await waitFor(async () => {
      const ctx = app.t.frame(RE.app, e.sid)?.ctx;
      if (!ctx || ctx === appCtx2) return false;
      return (await e.appFrame('document.documentElement.dataset.fxPhase || ""')) === 'ready' && app.t.frame(RE.editor, e.sid) && await has('처음 문단입니다');
    }, 60_000, 400);
    f.restore = { result: rr, reloaded: !!back, same: sha(path) === original };
    check(g, "주 창 '원래대로' → 파일이 처음 내용, 편집 창이 다시 불러옴", rr?.ok && rr.value?.editorNotified === true && f.restore.same && !!back, f.restore);
    await sleep(1000);
    await e.closeTips();

    // 원래대로 — 고친 것 있을 때: 띠
    await typeAtEnd('되돌리기 전 편집');
    await e.waitDirty(true);
    const rr2 = await app.apiRaw('doc.restore', { file: await openedVersion() });
    const b2 = await waitFor(async () => { const s = await e.state(); return s.banner === 'changed' ? s : null; }, 15_000);
    check(g, "고친 뒤 '원래대로' → 띠 '주 창에서 문서를 원래대로 되돌렸습니다'", rr2?.ok && !!b2 && b2.bannerText.includes('원래대로 되돌렸습니다'), b2?.bannerText ?? rr2);
    await e.host('window.__officeTest.bannerAction("다시 불러오기")');
    const clean = await waitFor(async () => { const s = await e.state(); return s.connected && !s.dirty && !s.banner ? s : null; }, 30_000);
    await e.waitReady();
    check(g, "'다시 불러오기' → 고친 내용 버리고 깨끗함", !!clean && !(await has('되돌리기 전 편집')), clean);
    await sleep(1000);
    await e.closeTips();

    // 닫기 확인: 취소 → 그대로, 저장 → 저장하고 닫힘
    await typeAtEnd('닫기 전에 저장');
    await e.waitDirty(true);
    await e.closeByKey();
    const ask = app.ps('dialog', '', 15_000);
    f.closePrompt = ask;
    check(g, "닫기 확인창: '바뀐 내용을 저장할까요?' 저장 / 저장 안 함 / 취소", ask?.ok && (ask.texts ?? []).some((x) => x.includes('바뀐 내용을 저장할까요?')) && ['저장', '저장 안 함', '취소'].every((b) => (ask.buttons ?? []).includes(b)), ask);
    const c1 = app.ps('dialog', '취소', 10_000);
    await sleep(1200);
    check(g, "'취소' → 창 그대로, 고친 내용 그대로", c1?.clicked === '취소' && e.alive && (await e.state()).dirty, c1);
    await e.closeByKey();
    const before = sha(path);
    const c2 = app.ps('dialog', '저장', 15_000);
    const closed2 = await e.waitClosed(60_000);
    const rd2 = await core.readDocument(path);
    check(g, "'저장' → 저장하고 닫힘", c2?.clicked === '저장' && closed2 && sha(path) !== before && flat(rd2.text).includes(flat('닫기 전에 저장')), { clicked: c2?.clicked, closed: closed2 });
    check(g, '정상 닫기 뒤 복구 사본 없음', !existsSync(draft));

    // 저장 안 함 → 닫히고 파일 그대로
    const o2 = await openEditor({ method: 'create', ext, path });
    e = o2.editor;
    await e.waitReady();
    await sleep(1200);
    await e.closeTips();
    await typeAtEnd('버릴 편집');
    await e.waitDirty(true);
    const before3 = sha(path);
    await e.closeByKey();
    const c3 = app.ps('dialog', '저장 안 함', 15_000);
    const closed3 = await e.waitClosed(30_000);
    check(g, "'저장 안 함' → 닫히고 파일 그대로", c3?.clicked === '저장 안 함' && closed3 && sha(path) === before3, { clicked: c3?.clicked, closed: closed3 });
    const st = await app.main('document.getElementById("st-editor")?.hidden');
    check(g, '편집 창을 닫으면 주 창 상태 막대에서 사라짐', st === true, st);
  } catch (err) {
    f.error = String(err?.stack ?? err);
    check(g, '흐름이 끝까지 돎', false, f.error.slice(0, 600));
    if (e?.alive) { try { await e.shot(`${g}-9-error`); } catch {} await e.closeByKey().catch(() => {}); if (!(await e.waitClosed(4000))) { app.ps('dialog', '저장 안 함', 4000); await e.waitClosed(8000); } }
  }
}

// ── 명령줄 ──────────────────────────────────────────────

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const opt = (n) => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
  const has = (n) => argv.includes(n);
  if (has('--help') || has('-h')) {
    console.log('사용법: node scripts/dev/office-e2e.mjs [--out 폴더] [--formats docx,xlsx,pptx] [--methods create,blank,office] [--com|--no-com] [--no-flows] [--office-docs 폴더] [--keep]');
    process.exit(0);
  }
  const list = (s) => (s ? s.split(',').map((x) => x.trim()).filter(Boolean) : undefined);
  try {
    const r = await runOfficeE2E({
      out: opt('--out'), formats: list(opt('--formats')), methods: list(opt('--methods')),
      com: has('--no-com') ? false : has('--com') ? true : 'auto', flows: !has('--no-flows'),
      officeDocs: opt('--office-docs'), keep: has('--keep'),
    });
    const fails = r.checks.filter((c) => !c.ok);
    console.log(`\n${r.checks.length - fails.length}/${r.checks.length} 통과 · 보고서 ${join(r.out, 'report.json')}`);
    process.exit(fails.length ? 1 : 0);
  } catch (err) {
    console.error('[office-e2e] 실행하지 못했습니다:', err?.stack ?? err);
    process.exit(2);
  }
}
