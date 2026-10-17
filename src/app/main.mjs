// main.mjs — Office Kit 창, 단일 인스턴스, IPC, 원본 백업, 파일 감시, AI 대화 중계.
import { app, BrowserWindow, Menu, dialog, ipcMain, nativeTheme, screen, session, shell } from 'electron';
import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, statSync, watch, writeFileSync } from 'node:fs';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBackups, createSettings, samePath } from './store.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const APP_ROOT = resolve(here, '../..');
const INDEX_HTML = join(here, 'renderer', 'index.html');
const SMOKE = process.env.OFFICEKIT_SMOKE ? resolve(process.env.OFFICEKIT_SMOKE) : '';
const SMOKE_CHAT = process.env.OFFICEKIT_SMOKE_CHAT === '1';
const SMOKE_SETTINGS = ['ai', 'view', 'connect', 'hancom'].includes(process.env.OFFICEKIT_SMOKE_SETTINGS) ? process.env.OFFICEKIT_SMOKE_SETTINGS : null;

// 테스트용으로 설정 폴더를 바꿀 수 있게 (단일 인스턴스 잠금도 폴더별로 걸린다)
if (process.env.OFFICEKIT_USER_DATA) app.setPath('userData', resolve(process.env.OFFICEKIT_USER_DATA));

if (!SMOKE && !app.requestSingleInstanceLock()) {
  app.quit();
} else {
  // 최상위 await로 기다리면 ESM 진입점에서는 ready가 오지 않으므로 그냥 실행한다
  start().catch((err) => {
    console.error('[office-kit] 시작 실패:', err);
    dialog.showErrorBox('Office Kit', `시작하지 못했습니다.\n${err?.message ?? err}`);
    app.exit(1);
  });
}

async function start() {
  // 문서 엔진은 잠금을 얻은 뒤에 불러온다 (WASM 초기화가 있어서)
  const core = await import('../core/index.mjs');
  const com = await import('../core/com.mjs');
  const { OPENABLE_EXTS, extOf, kindOf } = core;

  const userData = app.getPath('userData');
  const settings = createSettings(userData);
  const backups = createBackups(join(userData, 'backups'));
  const previewDir = join(userData, 'preview');

  /** @type {BrowserWindow | null} */
  let win = null;
  let current = null;       // 열린 문서 { path, backup, stat, renderedStat, quietUntil }
  let pendingFile = null;   // 창이 준비되기 전에 받은 파일
  let rendererReady = false;

  // ── 선택 모듈 (다른 사람이 함께 만드는 중일 수 있어 없으면 기능만 끈다) ──
  const modules = {};
  async function optional(name, spec) {
    if (modules[name]) return modules[name];
    try { modules[name] = await import(spec); return modules[name]; } catch (err) {
      const e = new Error(`${{ ai: 'AI 대화', integrations: '외부 연결', hancom: '한글 자동화' }[name]} 기능을 불러오지 못했습니다: ${err.message}`);
      e.code = 'MODULE_MISSING';
      throw e;
    }
  }
  // 개발 중 시험용으로 AI 모듈을 바꿔 끼울 수 있다 (설치본에서는 무시)
  const AI_MODULE = !app.isPackaged && process.env.OFFICEKIT_AI_MODULE ? pathToFileURL(resolve(process.env.OFFICEKIT_AI_MODULE)).href : './ai.mjs';
  const loadAi = () => optional('ai', AI_MODULE);
  const loadIntegrations = () => optional('integrations', '../setup/integrations.mjs');
  const loadHancom = () => optional('hancom', '../setup/hancom.mjs');

  async function mcpSpec() {
    try {
      const m = await loadIntegrations();
      return await m.mcpLaunchSpec({ execPath: process.execPath, appRoot: APP_ROOT });
    } catch {
      // integrations.mjs가 없을 때도 같은 규칙으로 MCP 서버를 띄운다
      return { command: process.execPath, args: [join(APP_ROOT, 'src', 'mcp', 'server.mjs')], env: { ELECTRON_RUN_AS_NODE: '1' } };
    }
  }

  // ── 검증 도우미 ──────────────────────────────────────

  // 예상한 오류(잘못된 요청·없는 파일 등)는 화면에만 알리고 로그를 남기지 않는다
  const fail = (message, code) => Object.assign(new Error(message), { expected: true }, code ? { code } : {});

  function docPath(p) {
    if (typeof p !== 'string' || !p.trim() || p.length > 4096) throw fail('잘못된 파일 경로입니다');
    if (!isAbsolute(p)) throw fail('파일의 전체 경로가 필요합니다');
    p = resolve(p);
    let st;
    try { st = statSync(p); } catch { throw fail(`파일을 찾을 수 없습니다: ${basename(p)}`, 'NOT_FOUND'); }
    if (!st.isFile()) throw fail('파일이 아닙니다');
    if (!OPENABLE_EXTS.includes(extOf(p))) throw fail(`지원하지 않는 형식입니다 (.${extOf(p) || '확장자 없음'})`, 'UNSUPPORTED');
    return p;
  }

  function existingFile(p, exts) {
    if (typeof p !== 'string' || !isAbsolute(p) || p.length > 4096) throw fail('잘못된 파일 경로입니다');
    p = resolve(p);
    if (!existsSync(p) || !statSync(p).isFile()) throw fail(`파일을 찾을 수 없습니다: ${basename(p)}`);
    if (exts && !exts.includes(extOf(p))) throw fail('허용되지 않는 파일 형식입니다');
    return p;
  }

  function needDoc() {
    if (!current) throw fail('열린 문서가 없습니다');
    if (!existsSync(current.path)) throw fail(`문서를 찾을 수 없습니다: ${basename(current.path)}`, 'NOT_FOUND');
    return current;
  }

  /**
   * 명령줄 인수에서 열 파일 찾기 (뒤에서부터, 스위치·앱 경로 제외).
   * 열 수 있는 파일이 없으면 마지막 인수를 돌려준다 → 열 때 '파일을 찾을 수 없습니다' 등을 알린다 (조용히 무시하지 않게)
   */
  function fileFromArgv(argv, cwd = process.cwd()) {
    const args = argv.slice(1).filter((a) => a && !a.startsWith('-'));
    if (process.defaultApp) args.shift(); // 개발 실행(electron .)의 앱 경로
    const paths = args.map((a) => resolve(cwd, a));
    for (const p of [...paths].reverse()) {
      if (!OPENABLE_EXTS.includes(extOf(p))) continue;
      try { if (statSync(p).isFile()) return p; } catch {}
    }
    return paths.at(-1) ?? null;
  }

  function send(channel, data) {
    if (win && !win.isDestroyed()) win.webContents.send(channel, data);
  }

  // ── 문서 ─────────────────────────────────────────────

  function statOf(p) {
    try { const st = statSync(p); return { mtimeMs: st.mtimeMs, size: st.size }; } catch { return null; }
  }

  function docInfo() {
    if (!current) return null;
    const st = statOf(current.path);
    return {
      path: current.path, name: basename(current.path), dir: dirname(current.path),
      ext: extOf(current.path), kind: kindOf(current.path),
      size: st?.size ?? 0, mtime: st?.mtimeMs ?? null, missing: !st,
      backup: current.backup ? { time: current.backup.time, file: current.backup.file } : null,
      backupError: current.backupError ?? null,
      hasSession: sessions.has(key(current.path)),
    };
  }

  // 이번 실행에서 문서를 처음 열 때 만든 백업 (문서 경로 → 백업). 다시 열어도 '원래대로' 기준은 이것
  const restorePoints = new Map();
  const backupExists = (e) => !!e?.file && existsSync(join(backups.dir, e.file));
  // 백업 정리 때 지우면 안 되는 것: 열었던 문서들의 '원래대로' 기준
  const protectedBackups = () => [...restorePoints.values()].map((e) => e.file);

  function focusWindow() {
    if (win) { if (win.isMinimized()) win.restore(); win.focus(); }
  }

  async function openDocument(p) {
    p = docPath(p);
    // 이미 열린 문서를 다시 열면(탐색기 더블클릭·끌어놓기·최근 목록) 백업 기준과 감시는 그대로 두고 앞으로 가져오기만 한다
    if (current && samePath(current.path, p)) {
      if (!watcher) startWatch(current.path);
      if (!SMOKE) settings.addRecent(current.path);
      const doc = docInfo();
      send('doc:opened', doc);
      focusWindow();
      return doc;
    }
    stopWatch();
    const k = key(p);
    const first = restorePoints.get(k);
    let backup = null; let backupError = null;
    try {
      // 지금 상태는 늘 남기되, 이번 실행에서 연 적이 있으면 처음 열 때의 백업으로 되돌린다 (그 사이 AI가 고친 것이 기준이 되지 않게)
      const made = backups.backup(p, new Date(), { protect: protectedBackups() });
      backup = backupExists(first) ? first : made;
    } catch (err) {
      if (backupExists(first)) backup = first;
      else backupError = `원본을 백업하지 못했습니다: ${err.message}`;
    }
    if (backup) restorePoints.set(k, backup);
    current = { path: p, backup, backupError, stat: statOf(p), missing: false, renderedStat: null, quietUntil: 0 };
    // 확인용 캡처(SMOKE)는 최근 파일 목록과 Windows 최근 항목을 건드리지 않는다
    // (Windows 최근 항목에 넣으면 탐색기가 그 형식의 '연결 프로그램' 목록에도 Office Kit를 적어 둔다)
    if (!SMOKE) {
      settings.addRecent(p);
      try { app.addRecentDocument(p); } catch {}
    }
    startWatch(p);
    const doc = docInfo();
    send('doc:opened', doc);
    focusWindow();
    return doc;
  }

  function closeDocument() {
    stopWatch();
    current = null;
    send('doc:opened', null);
  }

  // 파일 감시: 편집기가 새 파일로 바꿔치기하는 경우도 잡도록 폴더를 지켜본다
  let watcher = null; let watchTimer = null;
  function startWatch(p) {
    const name = basename(p).toLowerCase();
    const check = () => {
      if (!current || current.path !== p) return;
      if (Date.now() < current.quietUntil) return;
      const st = statOf(p);
      if (!st) {
        // 없어진 순간에만 알리고, 크기·시각을 잊어 두어 같은 파일이 돌아와도 알린다
        if (!current.missing) { current.missing = true; current.stat = null; send('doc:changed', { path: p, missing: true }); }
        return;
      }
      const back = current.missing;
      current.missing = false;
      const last = current.stat;
      if (!back && last && last.mtimeMs === st.mtimeMs && last.size === st.size) return;
      current.stat = st;
      send('doc:changed', { path: p, missing: false, mtime: st.mtimeMs, size: st.size });
    };
    try {
      watcher = watch(dirname(p), { persistent: false }, (_ev, file) => {
        if (file && String(file).toLowerCase() !== name) return;
        clearTimeout(watchTimer);
        watchTimer = setTimeout(check, 600);
      });
      watcher.on('error', () => {});
    } catch { watcher = null; }
  }
  function stopWatch() {
    clearTimeout(watchTimer);
    try { watcher?.close(); } catch {}
    watcher = null;
  }

  function pdfPageCount(file) {
    try {
      const s = readFileSync(file).toString('latin1');
      const pages = s.match(/\/Type\s*\/Page(?![a-zA-Z])/g)?.length ?? 0;
      if (pages) return pages;
      const counts = [...s.matchAll(/\/Count\s+(\d+)/g)].map((m) => Number(m[1]));
      return counts.length ? Math.max(...counts) : null;
    } catch { return null; }
  }

  /** 미리보기 폴더 정리 (keep 폴더만 남김) */
  function cleanPreviews(keep) {
    try {
      for (const f of readdirSync(previewDir)) {
        const full = join(previewDir, f);
        if (keep && samePath(full, keep)) continue;
        try { rmSync(full, { recursive: true, force: true }); } catch {} // PDF 뷰어가 잡고 있으면 다음에 지운다
      }
    } catch {}
  }

  // PDF 뷰어 제목에 내부 이름(s.xlsx.123.pdf) 대신 문서 이름이 보이도록 폴더를 나눠 옮긴다
  let pdfSeq = 0;
  function tidyPdf(file, docPath) {
    try {
      const dir = join(previewDir, `${Date.now()}-${++pdfSeq}`);
      mkdirSync(dir, { recursive: true });
      const out = join(dir, `${basename(docPath).replace(/\.[^.]+$/, '')}.pdf`);
      renameSync(file, out);
      return { file: out, dir };
    } catch { return { file, dir: null }; }
  }

  const clip = (rows, maxRows = 500, maxCols = 60) => (rows ?? []).slice(0, maxRows).map((r) => (r ?? []).slice(0, maxCols).map((v) => String(v ?? '')));

  async function preview({ mode, ifChanged } = {}) {
    const doc = needDoc();
    const st = statOf(doc.path);
    if (ifChanged && doc.renderedStat && st && doc.renderedStat.mtimeMs === st.mtimeMs && doc.renderedStat.size === st.size) {
      return { unchanged: true, doc: docInfo() };
    }
    const want = mode ?? settings.get('previewMode');
    const m = ['exact', 'text'].includes(want) ? want : 'auto';
    mkdirSync(previewDir, { recursive: true });
    const started = Date.now();
    let res;
    // '글자만': 프로그램을 띄우지 않고 파일 엔진으로 바로 읽는다 (한글은 원래 빠르므로 그대로 그림)
    const ext = extOf(doc.path);
    if (m === 'text' && ['docx', 'pptx', 'xlsx'].includes(ext)) res = { ...(await core.readDocument(doc.path, { engine: 'file' })), kind: 'text' };
    else res = await core.renderPreview(doc.path, previewDir, { mode: m === 'exact' ? 'exact' : 'auto' });
    if (current === doc) { doc.renderedStat = st; doc.stat = st; }
    const base = { mode: m, ms: Date.now() - started, doc: current === doc ? docInfo() : null };
    if (res.kind === 'pdf') {
      const pdf = tidyPdf(res.file, doc.path);
      cleanPreviews(pdf.dir ?? pdf.file);
      return { ...base, kind: 'pdf', engine: res.engine, url: `${pathToFileURL(pdf.file).href}#view=FitH&navpanes=0`, pages: pdfPageCount(pdf.file) };
    }
    if (res.kind === 'svg') {
      return { ...base, kind: 'svg', engine: res.engine, total: res.total,
        pages: res.pages.map((svg) => Buffer.from(String(svg), 'utf8').toString('base64')) };
    }
    // 그 외는 글자 미리보기 (readDocument 결과)
    return {
      ...base, kind: 'text', engine: res.engine,
      text: String(res.text ?? '').slice(0, 300_000),
      tables: (res.tables ?? []).slice(0, 50).map((t) => clip(t, 200, 30)),
      slides: (res.slides ?? []).slice(0, 200).map((s) => String(s ?? '')),
      sheets: (res.sheets ?? []).slice(0, 30).map((s) => ({ name: String(s.name ?? ''), rows: clip(s.rows) })),
      pages: res.pages ?? null,
    };
  }

  // 되돌리기 중인 문서와, 문서별 되돌린 횟수 (그 사이 시작하던 AI 대화를 막는 데 쓴다)
  const restoring = new Set();
  const restoreGen = new Map();

  /** 이 문서를 고치는 AI 대화를 멈추고 끝날 때까지 기다린다 → 멈춘 개수 (시간 안에 안 끝나면 -1) */
  async function stopChatsFor(p, ms = 10_000) {
    const running = [...chats.values()].filter((c) => samePath(c.path, p));
    if (!running.length) return 0;
    for (const c of running) { try { c.handle.cancel(); } catch {} }
    let timer;
    const ended = await Promise.race([
      Promise.allSettled(running.map((c) => Promise.resolve(c.handle.done))).then(() => true),
      new Promise((r) => { timer = setTimeout(() => r(false), ms); }),
    ]);
    clearTimeout(timer);
    return ended ? running.length : -1;
  }

  async function restore() {
    const doc = needDoc();
    const k = key(doc.path);
    if (restoring.has(k)) throw fail('이미 되돌리는 중입니다');
    const entry = doc.backup ?? backups.latest(doc.path);
    if (!entry) throw fail('이 문서의 백업이 없습니다');
    restoring.add(k);
    restoreGen.set(k, (restoreGen.get(k) ?? 0) + 1);
    try {
      // AI가 이 문서를 고치는 중이면 먼저 멈춘다 (안 그러면 되돌린 파일에 AI가 다시 쓴다)
      const stoppedAi = await stopChatsFor(doc.path);
      if (stoppedAi < 0) throw fail('AI 작업이 멈추지 않아 되돌리지 않았습니다. 잠시 뒤 다시 시도하세요.');
      // 되돌리기 전 상태도 남겨 둔다 (실수로 눌렀을 때 백업 폴더에서 찾을 수 있게). 되돌릴 백업은 정리하지 않는다
      try { backups.backup(doc.path, new Date(), { protect: [entry.file, ...protectedBackups()] }); } catch {}
      doc.quietUntil = Date.now() + 1500;
      try { backups.restore(doc.path, entry); } catch (err) {
        if (/EBUSY|EPERM|EACCES/.test(err.code ?? err.message)) throw fail('문서가 다른 프로그램에서 열려 있어 되돌릴 수 없습니다. 프로그램을 닫고 다시 시도하세요.');
        throw err;
      }
      doc.stat = statOf(doc.path);
      doc.missing = false;
      // AI가 기억하는 문서 내용은 이제 틀리므로 다음 메시지부터 새 대화로 시작한다
      const newChat = sessions.delete(k);
      return { time: entry.time, stoppedAi, newChat, doc: current === doc ? docInfo() : null };
    } finally {
      restoring.delete(k);
    }
  }

  /** 이 형식을 프로그램으로 다룰 수 있는지 미리 확인 (한글은 보안 모듈이 없으면 COM을 쓰지 않는다) */
  async function appCheck(p, what) {
    const ext = extOf(p);
    const apps = await com.probe();
    if (com.comAppFor(ext) === 'hwp') {
      if (!apps.hancom) throw fail(`한컴오피스가 설치돼 있지 않아 ${what}할 수 없습니다`);
      if (!apps.hancomSecurityModule) throw fail(`한글 문서를 ${what}하려면 설정에서 "한글 자동화 연결"을 먼저 해주세요.`, 'HANCOM_SETUP');
    } else if (!(await com.comAvailableFor(ext))) {
      throw fail(`${kindOf(p)} 프로그램이 설치돼 있지 않아 ${what}할 수 없습니다`);
    }
  }

  async function openInProgram() {
    const doc = needDoc();
    const ext = extOf(doc.path);
    const apps = await com.probe().catch(() => ({}));
    // 한글은 보안 모듈이 없으면 COM 대신 연결 프로그램으로 연다 (허용 창에서 멈추지 않게)
    if (com.comAppFor(ext) === 'hwp' && !apps.hancomSecurityModule) {
      const err = await shell.openPath(doc.path);
      if (err) throw fail(`프로그램에서 열지 못했습니다: ${err}`);
      return { via: 'shell' };
    }
    try {
      await core.openInApp(doc.path);
      return { via: 'app' };
    } catch (err) {
      const e2 = await shell.openPath(doc.path);
      if (e2) throw fail(`프로그램에서 열지 못했습니다: ${err.message}`);
      return { via: 'shell' };
    }
  }

  async function savePdf() {
    const doc = needDoc();
    await appCheck(doc.path, 'PDF로 저장');
    const stem = basename(doc.path).replace(/\.[^.]+$/, '');
    const r = await dialog.showSaveDialog(win, {
      title: 'PDF로 저장', defaultPath: join(dirname(doc.path), `${stem}.pdf`),
      filters: [{ name: 'PDF 문서', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { canceled: true };
    const output = r.filePath.toLowerCase().endsWith('.pdf') ? r.filePath : `${r.filePath}.pdf`;
    if (samePath(output, doc.path)) throw fail('원본 문서와 같은 이름으로는 저장할 수 없습니다');
    const res = await core.convertDocument(doc.path, output);
    return { output: res.output, engine: res.engine };
  }

  // ── AI 대화 ──────────────────────────────────────────

  const key = (p) => (process.platform === 'win32' ? resolve(p).toLowerCase() : resolve(p));
  const sessions = new Map();  // 문서 경로 → { provider, sessionId }
  const chats = new Map();     // chatId → { handle, path }
  let providerCache = null;

  async function providers({ refresh } = {}) {
    const ai = await loadAi();
    if (!providerCache || refresh || Date.now() - providerCache.at > 60_000) {
      const list = await ai.detectProviders();
      providerCache = { at: Date.now(), list: Array.isArray(list) ? list : [] };
    }
    const list = providerCache.list.map((p) => ({
      id: String(p.id), name: String(p.name ?? p.id), short: String(p.name ?? p.id).replace(/\s*\(.*\)\s*$/, '') || String(p.id),
      available: !!p.available, loggedIn: p.loggedIn ?? null,
      version: p.version ? String(p.version) : '', hint: p.hint ? String(p.hint) : '',
    }));
    const want = settings.get('provider');
    const selected = list.find((p) => p.id === want && p.available)?.id ?? list.find((p) => p.available)?.id ?? null;
    return { providers: list, selected, preferred: want };
  }

  async function sendChat({ prompt } = {}) {
    if (typeof prompt !== 'string' || !prompt.trim()) throw fail('보낼 내용을 입력하세요');
    if (prompt.length > 20_000) throw fail('메시지가 너무 깁니다 (2만 자 이하)');
    const doc = needDoc();
    const path = doc.path;
    const k = key(path);
    if (restoring.has(k)) throw fail('문서를 원래대로 되돌리는 중입니다. 끝난 뒤 다시 보내 주세요.');
    const gen = restoreGen.get(k) ?? 0;
    const ai = await loadAi();
    const { providers: list, selected } = await providers();
    if (!selected) throw fail('사용할 수 있는 AI가 없습니다. Claude Code나 Codex를 설치하고 로그인하세요.', 'NO_PROVIDER');
    const base = await mcpSpec();
    // 앱 안 AI는 열린 문서의 폴더에만 저장할 수 있다 (MCP 서버가 OFFICEKIT_WRITE_ROOTS로 막는다)
    const mcp = { ...base, env: { ...(base.env ?? {}), OFFICEKIT_WRITE_ROOTS: dirname(path) } };
    // 준비하는 사이 '원래대로'를 눌렀으면 시작하지 않는다 (되돌린 파일을 다시 고치지 않게)
    if (restoring.has(k) || (restoreGen.get(k) ?? 0) !== gen) throw fail('문서를 원래대로 되돌려서 이 요청은 보내지 않았습니다.');
    const sess = sessions.get(k);
    const chatId = randomUUID();
    // 모델은 AI마다 따로 (다른 AI의 모델 이름을 넘기면 모든 요청이 실패한다)
    const model = settings.get('models')?.[selected] || undefined;
    let errored = false;
    const onEvent = (ev) => {
      if (!ev || typeof ev !== 'object') return;
      if (ev.type === 'done' && ev.sessionId) sessions.set(k, { provider: selected, sessionId: String(ev.sessionId) });
      if (ev.type === 'error') errored = true;
      try { send('ai:event', { ...ev, chatId, path }); } catch {
        send('ai:event', { type: ev.type, text: ev.text, message: ev.message, id: ev.id, name: ev.name, ok: ev.ok, summary: ev.summary, delta: ev.delta, chatId, path });
      }
    };
    const handle = ai.startChat({
      provider: selected, prompt, sessionId: sess?.provider === selected ? sess.sessionId : undefined,
      // 대화 기록은 같은 작업 폴더에서만 이어지므로 ai.mjs의 고정 폴더를 쓴다
      cwd: typeof ai.defaultWorkspace === 'function' ? ai.defaultWorkspace() : join(userData, 'ai'),
      filePath: path, mcp, model, onEvent,
    });
    chats.set(chatId, { handle, path });
    Promise.resolve(handle.done)
      .then((r) => { if (r?.sessionId) sessions.set(k, { provider: selected, sessionId: String(r.sessionId) }); })
      .catch((err) => { if (!errored) send('ai:event', { type: 'error', message: err?.message ?? String(err), chatId, path }); })
      .finally(() => { chats.delete(chatId); send('ai:event', { type: 'end', chatId, path }); });
    return { chatId, provider: list.find((p) => p.id === selected)?.name ?? selected };
  }

  function cancelChat(chatId) {
    if (typeof chatId !== 'string') throw fail('잘못된 요청입니다');
    const c = chats.get(chatId);
    if (c) { try { c.handle.cancel(); } catch {} }
    return { canceled: !!c };
  }

  // ── 설정 ─────────────────────────────────────────────

  const PUBLIC_KEYS = ['provider', 'models', 'previewMode', 'theme', 'chatWidth', 'chatCollapsed'];
  const publicSettings = () => Object.fromEntries(PUBLIC_KEYS.map((k) => [k, k === 'models' ? { ...settings.get(k) } : settings.get(k)]));
  const PROVIDER_ID = /^[a-z0-9_-]{1,32}$/i;

  function setSettings(patch) {
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw fail('잘못된 설정입니다');
    const clean = {};
    for (const [k, v] of Object.entries(patch)) {
      if (k === 'provider') {
        if (v !== null && (typeof v !== 'string' || !PROVIDER_ID.test(v))) throw fail('잘못된 AI 선택입니다');
        clean.provider = v;
      } else if (k === 'models') {
        // 준 AI의 모델만 바꾼다. 빈 글자는 기본 모델로
        if (!v || typeof v !== 'object' || Array.isArray(v)) throw fail('잘못된 모델 설정입니다');
        const models = { ...settings.get('models') };
        for (const [id, name] of Object.entries(v)) {
          if (!PROVIDER_ID.test(id)) throw fail('잘못된 AI 선택입니다');
          if (typeof name !== 'string' || name.length > 100 || /[\r\n]/.test(name)) throw fail('모델 이름이 올바르지 않습니다');
          if (name.trim()) models[id] = name.trim(); else delete models[id];
        }
        clean.models = models;
      } else if (k === 'previewMode') {
        if (!['auto', 'exact', 'text'].includes(v)) throw fail('잘못된 미리보기 방식입니다');
        clean.previewMode = v;
      } else if (k === 'theme') {
        if (!['system', 'light', 'dark'].includes(v)) throw fail('잘못된 테마입니다');
        clean.theme = v;
        nativeTheme.themeSource = v;
      } else if (k === 'chatWidth') {
        if (!Number.isFinite(v)) throw fail('잘못된 너비입니다');
        clean.chatWidth = Math.round(Math.min(800, Math.max(280, v)));
      } else if (k === 'chatCollapsed') {
        clean.chatCollapsed = !!v;
      } else throw fail(`알 수 없는 설정: ${k}`);
    }
    settings.set(clean);
    return publicSettings();
  }

  // ── 외부 연결 · 한글 자동화 ───────────────────────────

  const INTEGRATIONS = {
    claudeDesktop: ['connectClaudeDesktop', 'disconnectClaudeDesktop'],
    claudeCode: ['connectClaudeCode', 'disconnectClaudeCode'],
    codex: ['connectCodex', 'disconnectCodex'],
    cli: ['installCliShim', 'uninstallCliShim'],
  };

  async function integrationStatus() {
    const m = await loadIntegrations();
    const raw = await m.integrationStatus(await mcpSpec());
    // 값이 true/false든 객체든 화면에서 쓰기 좋게 맞춘다
    const out = {};
    for (const k of Object.keys(INTEGRATIONS)) {
      const v = raw?.[k];
      if (v && typeof v === 'object') {
        out[k] = {
          connected: !!(v.connected ?? v.configured ?? v.installed ?? v.registered ?? v.ok),
          available: v.available !== false,
          matches: v.matches !== false,
          error: v.error ? String(v.error) : '',
          detail: String(v.detail ?? v.message ?? v.hint ?? v.path ?? v.configPath ?? ''),
        };
      } else out[k] = { connected: !!v, available: true, matches: true, error: '', detail: '' };
    }
    return out;
  }

  async function setIntegration({ target, on } = {}) {
    if (!Object.hasOwn(INTEGRATIONS, target) || typeof on !== 'boolean') throw fail('잘못된 요청입니다');
    const m = await loadIntegrations();
    const fn = m[INTEGRATIONS[target][on ? 0 : 1]];
    if (typeof fn !== 'function') throw fail('이 연결은 아직 지원하지 않습니다');
    const spec = await mcpSpec();
    const res = on ? await fn(spec) : await fn();
    return { result: res && typeof res === 'object' ? JSON.parse(JSON.stringify(res)) : res ?? null, status: await integrationStatus() };
  }

  async function hancomStatus() {
    const h = await loadHancom();
    const s = (await h.hancomSecurityStatus()) ?? {};
    return {
      hancomInstalled: !!s.hancomInstalled, registered: !!s.registered,
      dllPath: s.dllPath ? String(s.dllPath) : '', dllExists: !!s.dllExists,
      warning: s.warning ? String(s.warning) : '',
      source: { page: String(h.HANCOM_SECURITY_MODULE_PAGE ?? ''), url: String(h.HANCOM_SECURITY_MODULE_URL ?? '') },
      canRemove: typeof h.removeHancomSecurity === 'function',
    };
  }

  async function hancomSetup({ consent, mode } = {}) {
    if (consent !== true) throw fail('동의가 필요합니다');
    const h = await loadHancom();
    let dllPath;
    if (mode === 'pick') {
      const r = await dialog.showOpenDialog(win, {
        title: '한글 보안 승인 모듈(DLL) 선택', properties: ['openFile'],
        filters: [{ name: '보안 승인 모듈', extensions: ['dll'] }],
      });
      if (r.canceled || !r.filePaths[0]) return { canceled: true };
      dllPath = existingFile(r.filePaths[0], ['dll']);
    } else if (mode === 'download') {
      const got = await h.downloadHancomSecurityModule(join(userData, 'hancom'));
      dllPath = typeof got === 'string' ? got : got?.dllPath ?? got?.path;
      if (!dllPath) throw fail('보안 모듈을 내려받지 못했습니다');
    } else throw fail('잘못된 요청입니다');
    await h.setupHancomSecurity({ dllPath });
    await com.probe({ refresh: true }).catch(() => {});
    return { status: await hancomStatus() };
  }

  // ── IPC ──────────────────────────────────────────────

  // 모든 요청은 { ok, value } / { ok:false, error, code } 로 돌려준다
  function handle(channel, fn) {
    ipcMain.handle(channel, async (e, ...args) => {
      try {
        if (!win || e.sender !== win.webContents || !e.senderFrame || e.senderFrame !== win.webContents.mainFrame) {
          throw fail('허용되지 않은 요청입니다');
        }
        return { ok: true, value: await fn(...args) };
      } catch (err) {
        if (!err?.expected && err?.code !== 'MODULE_MISSING') console.error(`[${channel}]`, err);
        return { ok: false, error: err?.message ?? String(err), code: err?.code ?? null };
      }
    });
  }

  handle('app:init', () => {
    rendererReady = true;
    const opening = !!pendingFile;
    // 응답이 먼저 가도록 다음 틱에 연다
    if (pendingFile) {
      const f = pendingFile; pendingFile = null;
      setTimeout(() => openDocument(f).catch((err) => send('toast', { type: 'error', message: err.message })), 0);
    }
    return {
      version: app.getVersion(), packaged: app.isPackaged, platform: process.platform,
      exts: OPENABLE_EXTS, settings: publicSettings(), recent: settings.recent(), doc: docInfo(),
      smoke: SMOKE ? { chat: SMOKE_CHAT, settings: SMOKE_SETTINGS } : null, opening,
    };
  });

  handle('doc:pick', async () => {
    const all = OPENABLE_EXTS;
    const r = await dialog.showOpenDialog(win, {
      title: '문서 열기', properties: ['openFile'],
      defaultPath: current ? dirname(current.path) : app.getPath('documents'),
      filters: [
        { name: '모든 문서', extensions: all },
        { name: '한글', extensions: all.filter((e) => kindOf(`a.${e}`) === '한글') },
        { name: 'Word', extensions: all.filter((e) => kindOf(`a.${e}`) === 'Word') },
        { name: 'PowerPoint', extensions: all.filter((e) => kindOf(`a.${e}`) === 'PowerPoint') },
        { name: 'Excel', extensions: all.filter((e) => kindOf(`a.${e}`) === 'Excel') },
      ],
    });
    if (r.canceled || !r.filePaths[0]) return null;
    return openDocument(r.filePaths[0]);
  });
  handle('doc:open', (p) => openDocument(p));
  handle('doc:close', () => { closeDocument(); return null; });
  handle('doc:preview', (opts = {}) => preview({
    mode: ['auto', 'exact', 'text'].includes(opts?.mode) ? opts.mode : undefined, ifChanged: opts?.ifChanged === true,
  }));
  handle('doc:restore', () => restore());
  handle('doc:openInApp', () => openInProgram());
  handle('doc:savePdf', () => savePdf());
  handle('doc:showInFolder', () => { shell.showItemInFolder(needDoc().path); return null; });
  handle('shell:showItem', (p) => { shell.showItemInFolder(existingFile(p)); return null; });
  handle('shell:openBackups', async () => { mkdirSync(backups.dir, { recursive: true }); await shell.openPath(backups.dir); return null; });
  handle('shell:openExternal', async (url) => {
    let u;
    try { u = new URL(String(url)); } catch { throw fail('잘못된 주소입니다'); }
    if (u.protocol !== 'https:' && u.protocol !== 'http:') throw fail('웹 주소만 열 수 있습니다');
    await shell.openExternal(u.href);
    return null;
  });

  handle('recent:list', () => settings.recent());
  handle('recent:remove', (p) => { if (typeof p !== 'string') throw fail('잘못된 경로입니다'); settings.removeRecent(p); return settings.recent(); });
  handle('recent:clear', () => { settings.clearRecent(); return []; });

  handle('settings:get', () => publicSettings());
  handle('settings:set', (patch) => setSettings(patch));

  handle('ai:providers', (opts) => providers({ refresh: opts?.refresh === true }));
  handle('ai:send', (args) => sendChat(args));
  handle('ai:cancel', (chatId) => cancelChat(chatId));
  handle('ai:reset', () => { if (current) sessions.delete(key(current.path)); return null; });

  handle('integrations:status', () => integrationStatus());
  handle('integrations:set', (args) => setIntegration(args));
  handle('hancom:status', async () => {
    const apps = await com.probe({ refresh: true }).catch(() => ({}));
    let status = null; let error = null;
    try { status = await hancomStatus(); } catch (err) { error = err.message; }
    return { status, error, apps: { hancom: !!apps.hancom, word: !!apps.word, powerpoint: !!apps.powerpoint, excel: !!apps.excel, hancomSecurityModule: !!apps.hancomSecurityModule } };
  });
  handle('hancom:setup', (args) => hancomSetup(args));
  handle('hancom:remove', async (args) => {
    if (args?.consent !== true) throw fail('동의가 필요합니다');
    const h = await loadHancom();
    if (typeof h.removeHancomSecurity !== 'function') throw fail('연결 해제를 지원하지 않습니다');
    await h.removeHancomSecurity();
    await com.probe({ refresh: true }).catch(() => {});
    return { status: await hancomStatus() };
  });

  // ── 확인용 캡처 (OFFICEKIT_SMOKE) ─────────────────────

  if (SMOKE) {
    const bail = setTimeout(() => { console.error('[smoke] 시간 초과'); app.exit(1); }, 240_000);
    ipcMain.on('smoke:ready', async (e, info = {}) => {
      if (!win || e.sender !== win.webContents) return;
      try {
        await new Promise((r) => setTimeout(r, info.kind === 'pdf' ? 3000 : 900));
        let img;
        for (let i = 0; i < 8 && !img; i++) {
          try { img = await win.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true }); } catch (err) {
            if (i === 7) throw err;
            await new Promise((r) => setTimeout(r, 700));
          }
        }
        mkdirSync(dirname(SMOKE), { recursive: true });
        writeFileSync(SMOKE, img.toPNG());
        console.log('[smoke]', JSON.stringify({ ...info, png: SMOKE, size: img.getSize() }));
        clearTimeout(bail);
        app.exit(info.ok === false ? 1 : 0);
      } catch (err) {
        console.error('[smoke] 캡처 실패:', err);
        app.exit(1);
      }
    });
  }

  // ── 창 ───────────────────────────────────────────────

  function initialBounds() {
    const base = { width: 1440, height: 920 };
    const w = SMOKE ? null : settings.get('window');
    if (!w || !Number.isFinite(w.width) || !Number.isFinite(w.height)) return base;
    const size = { width: Math.max(1000, Math.round(w.width)), height: Math.max(680, Math.round(w.height)) };
    // 저장된 위치가 지금 화면 안에 있을 때만 쓴다 (모니터를 뺀 경우)
    const visible = Number.isFinite(w.x) && Number.isFinite(w.y) && screen.getAllDisplays().some(({ workArea: a }) =>
      w.x < a.x + a.width - 80 && w.x + size.width > a.x + 80 && w.y >= a.y - 10 && w.y < a.y + a.height - 80);
    return visible ? { ...size, x: Math.round(w.x), y: Math.round(w.y) } : size;
  }

  function isPreviewUrl(u) {
    if (!u.startsWith('file:')) return false;
    try {
      const p = fileURLToPath(u.split('#')[0]);
      const rel = relative(previewDir, p);
      return !!rel && !rel.startsWith('..') && !isAbsolute(rel);
    } catch { return false; }
  }

  function createWindow() {
    const saved = SMOKE ? null : settings.get('window');
    win = new BrowserWindow({
      ...initialBounds(), minWidth: 1000, minHeight: 680, title: 'Office Kit', show: false,
      backgroundColor: nativeTheme.shouldUseDarkColors ? '#16181d' : '#f4f5f7',
      autoHideMenuBar: true,
      icon: [join(APP_ROOT, 'build', 'icon.ico'), join(APP_ROOT, 'build', 'icon.png')].find((f) => existsSync(f)),
      webPreferences: {
        preload: join(here, 'preload.cjs'),
        contextIsolation: true, nodeIntegration: false, sandbox: true,
        plugins: true,            // 내장 PDF 뷰어
        webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
        backgroundThrottling: !SMOKE,
      },
    });
    if (saved?.maximized) win.maximize();
    win.once('ready-to-show', () => { win.show(); if (SMOKE) win.focus(); });

    const wc = win.webContents;
    // 새 창·이동 막기: 웹 주소는 기본 브라우저로
    wc.setWindowOpenHandler(({ url }) => {
      if (/^https?:\/\//i.test(url)) shell.openExternal(url).catch(() => {});
      return { action: 'deny' };
    });
    wc.on('will-navigate', (e) => e.preventDefault());
    wc.on('will-frame-navigate', (e) => {
      if (e.isMainFrame) { e.preventDefault(); return; }
      const u = e.url;
      // 미리보기 PDF와 내장 PDF 뷰어만 허용
      if (u === 'about:blank' || u.startsWith('chrome-extension://') || isPreviewUrl(u)) return;
      e.preventDefault();
      if (/^https?:\/\//i.test(u)) shell.openExternal(u).catch(() => {});
    });
    wc.on('will-attach-webview', (e) => e.preventDefault());
    // 내려받기는 PDF 뷰어의 저장 단추(미리보기 PDF)만 허용 → 저장 위치를 묻는다
    wc.session.on('will-download', (e, item) => {
      if (!isPreviewUrl(item.getURL())) { e.preventDefault(); return; }
      const name = current ? `${basename(current.path).replace(/\.[^.]+$/, '')}.pdf` : item.getFilename();
      item.setSaveDialogOptions({ title: 'PDF로 저장', defaultPath: join(current ? dirname(current.path) : app.getPath('documents'), name), filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] });
    });
    wc.on('render-process-gone', (_e, d) => console.error('[renderer] 종료:', d.reason));
    if (SMOKE) wc.on('console-message', (e) => { if (e.level === 'error' || e.level === 'warning') console.log('[renderer]', e.message); });

    win.on('close', () => {
      if (SMOKE || !win) return;
      try { settings.set({ window: { ...win.getNormalBounds(), maximized: win.isMaximized() } }); } catch {}
    });
    win.on('closed', () => { win = null; });

    win.loadFile(INDEX_HTML);
  }

  function buildMenu() {
    const template = [
      { label: '파일', submenu: [
        { label: '파일 열기…', accelerator: 'CmdOrCtrl+O', click: () => send('menu', 'open') },
        { label: '새로고침', accelerator: 'F5', click: () => send('menu', 'refresh') },
        { type: 'separator' },
        { label: '설정', accelerator: 'CmdOrCtrl+,', click: () => send('menu', 'settings') },
        { type: 'separator' },
        { label: '끝내기', role: 'quit' },
      ] },
    ];
    if (!app.isPackaged) {
      template.push({ label: '개발', submenu: [
        { label: '개발자 도구', accelerator: 'CmdOrCtrl+Shift+I', click: () => win?.webContents.toggleDevTools() },
        { label: '화면 다시 불러오기', accelerator: 'CmdOrCtrl+Shift+R', click: () => win?.webContents.reloadIgnoringCache() },
      ] });
    }
    Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  }

  // ── 시작 ─────────────────────────────────────────────

  pendingFile = fileFromArgv(process.argv);

  app.on('second-instance', (_e, argv, cwd) => {
    const f = fileFromArgv(argv, cwd);
    if (win) { if (win.isMinimized()) win.restore(); win.show(); win.focus(); }
    if (!f) return;
    if (rendererReady) openDocument(f).catch((err) => send('toast', { type: 'error', message: err.message }));
    else pendingFile = f;
  });

  app.on('window-all-closed', () => app.quit());
  app.on('before-quit', () => {
    stopWatch();
    for (const c of chats.values()) { try { c.handle.cancel(); } catch {} }
  });
  // 창이 다 닫힌 뒤 미리보기 PDF(문서 사본)를 지운다. 못 지운 것은 다음 시작 때 지운다
  app.on('will-quit', () => cleanPreviews(null));

  await app.whenReady();
  nativeTheme.themeSource = ['light', 'dark'].includes(settings.get('theme')) ? settings.get('theme') : 'system';

  // 권한 요청은 PDF 뷰어에 필요한 것만
  const allowed = new Set(['fullscreen', 'clipboard-sanitized-write']);
  session.defaultSession.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  session.defaultSession.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));

  cleanPreviews(null);
  buildMenu();
  createWindow();
}
