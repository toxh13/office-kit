// main.mjs — Office Kit 창, 단일 인스턴스, IPC, 원본 백업, 파일 감시, AI 대화 중계.
import { app, BrowserWindow, Menu, dialog, ipcMain, nativeTheme, screen, session, shell } from 'electron';
import { randomUUID } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, renameSync, rmSync, statSync, watch, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBackups, createPreviewCache, createSettings, hashFile, samePath } from './store.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const APP_ROOT = resolve(here, '../..');
const INDEX_HTML = join(here, 'renderer', 'index.html');
const SMOKE = process.env.OFFICEKIT_SMOKE ? resolve(process.env.OFFICEKIT_SMOKE) : '';
const SMOKE_CHAT = process.env.OFFICEKIT_SMOKE_CHAT === '1';
const SMOKE_SETTINGS = ['ai', 'view', 'open', 'connect', 'hancom'].includes(process.env.OFFICEKIT_SMOKE_SETTINGS) ? process.env.OFFICEKIT_SMOKE_SETTINGS : null;
// 새 문서 창을 연 채로 찍기: '1' = 빈 문서, 'ai' = AI에게 부탁하기를 고른 모양
const SMOKE_NEWDOC = ['1', 'ai', 'template'].includes(process.env.OFFICEKIT_SMOKE_NEWDOC) ? process.env.OFFICEKIT_SMOKE_NEWDOC : null;
// 편집 창(한글·문서·표·슬라이드)도 열어 두 창을 함께 찍기 (편집 창은 <png>-editor.png)
const SMOKE_EDITOR = process.env.OFFICEKIT_SMOKE_EDITOR === '1';
// 'PDF로 저장'도 눌러 보기: 저장 대화상자 대신 이 경로에 만든다 (OFFICEKIT_SMOKE와 함께일 때만). 결과는 [smoke] 줄의 pdf
const SMOKE_PDF = SMOKE && process.env.OFFICEKIT_SMOKE_PDF ? resolve(process.env.OFFICEKIT_SMOKE_PDF) : '';
// '바뀐 내용' 창 찍기: 문서를 연 뒤(처음 연 상태 버전이 생긴 뒤) 이 파일로 덮어쓰고, 처음 연 상태 → 지금을 비교해 보인다
const SMOKE_DIFF = SMOKE && process.env.OFFICEKIT_SMOKE_DIFF ? resolve(process.env.OFFICEKIT_SMOKE_DIFF) : '';
// 도구 창 찍기: 'pii'(찾은 목록) | 'pii-mask'(가린 사본까지) | 'batch'(OFFICEKIT_SMOKE_BATCH 폴더를 바꾼 결과)
const SMOKE_TOOL = SMOKE && ['pii', 'pii-mask', 'batch', 'pdftools', 'print'].includes(process.env.OFFICEKIT_SMOKE_TOOL) ? process.env.OFFICEKIT_SMOKE_TOOL : null;
const SMOKE_BATCH = SMOKE_TOOL === 'batch' && process.env.OFFICEKIT_SMOKE_BATCH ? resolve(process.env.OFFICEKIT_SMOKE_BATCH) : '';
const APP_ICON = [join(APP_ROOT, 'build', 'icon.ico'), join(APP_ROOT, 'build', 'icon.png')].find((f) => existsSync(f));

// 테스트용으로 설정 폴더를 바꿀 수 있게 (단일 인스턴스 잠금도 폴더별로 걸린다)
if (process.env.OFFICEKIT_USER_DATA) app.setPath('userData', resolve(process.env.OFFICEKIT_USER_DATA));
// 확인용 캡처: 다른 창에 가려진 창도 계속 그리게 (가려진 창은 Chromium이 그리기를 멈춰 예전 화면이 찍힌다)
if (SMOKE) app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
// 이 PC 밖 이름은 풀지 않는다 (ready 전에): 편집기 틀의 스크립트가 <link rel=dns-prefetch>로 문서 내용을 DNS 이름에 실어
// 내보내지 못하게. 세션 요청 막기·죽은 프록시·CSP는 DNS를 막지 못한다 (보안 검증: 시스템 DNS로 실제 조회가 나감).
// 앱은 인터넷에 Chromium 네트워크를 쓰지 않는다 (AI·MCP는 자식 프로세스, 주 창 CSP connect-src 'none', 링크는 기본 브라우저로).
// 편집기 서버 주소(127.0.0.1~3)는 빼야 한다 — 빼지 않으면 편집 창이 뜨지 않는다
export const HOST_RESOLVER_RULES = 'MAP * ~NOTFOUND, EXCLUDE localhost, EXCLUDE 127.0.0.1, EXCLUDE 127.0.0.2, EXCLUDE 127.0.0.3';
app.commandLine.appendSwitch('host-resolver-rules', HOST_RESOLVER_RULES);

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
  const create = await import('../core/create.mjs');
  const templatesMod = await import('../core/templates.mjs');
  const { OPENABLE_EXTS, extOf, kindOf } = core;

  const userData = app.getPath('userData');
  const settings = createSettings(userData);
  const backups = createBackups(join(userData, 'backups'));
  const previewDir = join(userData, 'preview');
  // 프로그램으로 만든 미리보기 PDF는 파일이 그대로면 다시 쓴다 (Excel·Word 변환은 5~25초)
  const previewCache = createPreviewCache(join(userData, 'preview-cache'));

  /** @type {BrowserWindow | null} */
  let win = null;
  let current = null;       // 열린 문서 { path, backup, stat, renderedStat, quietUntil }
  let pendingFile = null;   // 창이 준비되기 전에 받은 파일
  // 문서 창 방식: 탐색기에서 문서를 열면(두 번 누르기·연결 프로그램) 그 문서의 편집 창만 띄운다 (Word처럼).
  // 주 창은 숨겨 두고, 편집 창의 'AI · 도구'를 누르면 그 옆에 좁은 '옆 창'으로 보인다. 마지막 편집 창을 닫으면 끝난다.
  // hubPrimary: 사용자가 Office Kit 자체를 실행했다 (주 창이 처음부터 보임) → 주 창을 닫으면 앱이 끝난다 (예전처럼)
  let hubPrimary = true;
  let panelOf = null;       // 옆 창이 붙어 있는 편집 창 (BrowserWindow)
  let hubBounds = null;     // 옆 창이 되기 전 주 창 크기 (전체 창으로 돌아갈 때)
  let appQuitting = false;
  let hubMaximizeLater = false; // 숨긴 채 시작한 주 창: 처음 보일 때 최대화 (저장된 창 상태)
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
      editor: editorState(current.path),
    };
  }

  // 이번 실행에서 문서를 처음 열 때 만든 백업 (문서 경로 → 백업). 버전 목록의 '처음 연 상태' — 정리할 때 지우지 않는다
  const restorePoints = new Map();
  // 문서 버전(백업): 문서를 열 때·편집 창 저장 전·AI 대화 전·원래대로 전에 남긴다. 문서마다 최근 몇 개까지
  const VERSIONS_PER_DOC = 30;
  const keepVersion = (p, reason) => backups.backup(p, new Date(), { protect: protectedBackups(), maxPerSource: VERSIONS_PER_DOC, reason });
  const backupExists = (e) => !!e?.file && existsSync(join(backups.dir, e.file));
  // 백업 정리 때 지우면 안 되는 것: 열었던 문서들의 '원래대로' 기준
  const protectedBackups = () => [...restorePoints.values()].map((e) => e.file);

  function focusWindow() {
    if (win) { if (win.isMinimized()) win.restore(); win.focus(); }
  }

  /** 열 때 편집 창도 바로 띄울지 (설정 '문서를 열면': 편집 창 / 미리보기). 편집 창으로 열 수 있는 형식만 */
  // 확인용 캡처는 OFFICEKIT_SMOKE_AUTOEDIT=1일 때만 (다른 캡처는 주 창만 찍는다)
  const wantsEditor = (p) => (!SMOKE || process.env.OFFICEKIT_SMOKE_AUTOEDIT === '1') && settings.get('openAction') !== 'preview' && !!editorKind(p);

  /**
   * 문서 열기. autoEdit: 사용자가 연 것(탐색기·끌어놓기·최근 목록·열기 대화상자)이면 설정에 따라 편집 창도 연다
   * → 화면이 doc:opened의 autoEdit를 보고 '직접 편집'과 같은 길로 편집 창을 띄운다
   */
  async function openDocument(p, { autoEdit = false } = {}) {
    p = docPath(p);
    const edit = autoEdit && wantsEditor(p);
    // 이미 열린 문서를 다시 열면(탐색기 더블클릭·끌어놓기·최근 목록) 백업 기준과 감시는 그대로 두고 앞으로 가져오기만 한다
    if (current && samePath(current.path, p)) {
      if (!watcher) startWatch(current.path);
      if (!SMOKE) settings.addRecent(current.path);
      const doc = docInfo();
      send('doc:opened', edit ? { ...doc, autoEdit: true } : doc);
      if (!edit) focusWindow();
      return doc;
    }
    stopWatch();
    const k = key(p);
    const first = restorePoints.get(k);
    let backup = null; let backupError = null;
    try {
      // 지금 상태는 늘 남기되, 이번 실행에서 연 적이 있으면 처음 열 때의 백업으로 되돌린다 (그 사이 AI가 고친 것이 기준이 되지 않게)
      const made = keepVersion(p, 'open');
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
    send('doc:opened', edit ? { ...doc, autoEdit: true } : doc);
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
        if (!current.missing) {
          current.missing = true; current.stat = null;
          send('doc:changed', { path: p, missing: true });
          editors?.notifyExternalChange(p);
        }
        return;
      }
      const back = current.missing;
      current.missing = false;
      const last = current.stat;
      if (!back && last && last.mtimeMs === st.mtimeMs && last.size === st.size) return;
      current.stat = st;
      send('doc:changed', { path: p, missing: false, mtime: st.mtimeMs, size: st.size });
      // 편집 창이 열려 있으면 알린다 (고친 것 없으면 다시 불러오고, 있으면 편집 창이 물어본다. 내용이 같으면 무시)
      editors?.notifyExternalChange(p);
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
    const cacheKey = st && m !== 'text' ? { path: doc.path, size: st.size, mtimeMs: st.mtimeMs, mode: m } : null;
    let cached = null;
    try { cached = cacheKey && previewCache.get(cacheKey); } catch {}
    if (cached) {
      // 기억해 둔 PDF의 사본을 미리보기 폴더에 (보는 동안 기억 정리가 지우지 못하게)
      const file = join(previewDir, `${basename(doc.path)}.${Date.now()}.pdf`);
      try { copyFileSync(cached.file, file); res = { kind: 'pdf', file, engine: cached.engine, pages: cached.pages, cached: true }; } catch {}
    }
    if (!res && m === 'text' && ['docx', 'pptx', 'xlsx'].includes(ext)) res = { ...(await core.readDocument(doc.path, { engine: 'file' })), kind: 'text' };
    else if (!res) res = await core.renderPreview(doc.path, previewDir, { mode: m === 'exact' ? 'exact' : 'auto' });
    if (res.kind === 'pdf' && !res.cached) {
      // 빈 시트처럼 쪽이 하나도 없는 PDF는 PDF 뷰어가 빈 화면에서 멈춘다 → 글자 미리보기로
      const pages = pdfPageCount(res.file);
      if (pages === 0) {
        try { rmSync(res.file, { force: true }); } catch {}
        res = { ...(await core.readDocument(doc.path, { engine: 'file' })), kind: 'text', note: '내용이 없어 미리 볼 쪽이 없어요' };
      } else {
        res = { ...res, pages };
        // 변환하는 동안 파일이 바뀌지 않았을 때만 기억한다
        const after = statOf(doc.path);
        if (cacheKey && after && after.size === st.size && after.mtimeMs === st.mtimeMs) {
          try { previewCache.put(cacheKey, res.file, { engine: res.engine, pages }); } catch (err) { console.error('[preview] 기억', err?.message ?? err); }
        }
      }
    }
    if (current === doc) { doc.renderedStat = st; doc.stat = st; }
    const base = { mode: m, ms: Date.now() - started, doc: current === doc ? docInfo() : null };
    if (res.kind === 'pdf') {
      const pdf = tidyPdf(res.file, doc.path);
      cleanPreviews(pdf.dir ?? pdf.file);
      return { ...base, kind: 'pdf', engine: res.engine, cached: !!res.cached, url: `${pathToFileURL(pdf.file).href}#view=FitH&navpanes=0`, pages: res.pages };
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
      note: res.note ? String(res.note) : null,
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

  // 지금 파일 내용의 해시 (크기·수정 시각이 같으면 다시 읽지 않는다)
  const hashCache = new Map();
  function currentHash(p) {
    const st = statOf(p);
    if (!st) return null;
    const c = hashCache.get(key(p));
    if (c && c.mtimeMs === st.mtimeMs && c.size === st.size) return c.hash;
    const hash = hashFile(p);
    hashCache.set(key(p), { ...st, hash });
    return hash;
  }

  /**
   * 열린 문서의 버전 목록 → { versions: [{ file, time, reason, size, current }], undo: 파일 이름 | null }
   * 최근에 쓴 것부터. current = 지금 파일과 내용이 같음. undo = 마지막 변경 전 (지금과 다른 가장 최근 버전)
   */
  function versions() {
    const doc = needDoc();
    let now = null;
    try { now = currentHash(doc.path); } catch {}
    const list = backups.list(doc.path).map((e) => ({ file: e.file, time: e.time, usedAt: e.usedAt ?? e.time, reason: e.reason ?? null, size: e.size ?? 0, current: !!now && e.hash === now }));
    return { versions: list, undo: list.find((e) => !e.current)?.file ?? null };
  }

  /** 버전으로 되돌리기. file을 주지 않으면 마지막 변경 전으로 (되돌리기 전 상태도 '원래대로 전' 버전으로 남는다) */
  async function restore(args) {
    const doc = needDoc();
    const k = key(doc.path);
    if (restoring.has(k)) throw fail('이미 되돌리는 중입니다');
    const want = args && typeof args === 'object' && typeof args.file === 'string' ? args.file : null;
    let entry;
    if (want) {
      entry = backups.find(doc.path, want);
      if (!entry) throw fail('그 버전을 찾을 수 없어요 (정리됐을 수 있어요). 목록을 다시 열어 주세요.', 'NO_VERSION');
    } else {
      const v = versions();
      if (!v.versions.length) throw fail('이 문서의 백업이 없습니다');
      entry = v.undo ? backups.find(doc.path, v.undo) : null;
      // 바꾼 것이 없으면 (모든 버전이 지금과 같음) 파일은 그대로 두고 편집 창에만 알린다
      entry ??= backups.find(doc.path, v.versions[0].file);
    }
    if (!entry) throw fail('이 문서의 백업이 없습니다');
    restoring.add(k);
    restoreGen.set(k, (restoreGen.get(k) ?? 0) + 1);
    try {
      // AI가 이 문서를 고치는 중이면 먼저 멈춘다 (안 그러면 되돌린 파일에 AI가 다시 쓴다)
      const stoppedAi = await stopChatsFor(doc.path);
      if (stoppedAi < 0) throw fail('AI 작업이 멈추지 않아 되돌리지 않았습니다. 잠시 뒤 다시 시도하세요.');
      // 되돌리기 전 상태도 버전으로 남겨 둔다 (다시 '원래대로'로 돌아올 수 있게). 되돌릴 백업은 정리하지 않는다
      try { backups.backup(doc.path, new Date(), { protect: [entry.file, ...protectedBackups()], maxPerSource: VERSIONS_PER_DOC, reason: 'restore' }); } catch {}
      doc.quietUntil = Date.now() + 1500;
      try { backups.restore(doc.path, entry); } catch (err) {
        if (/EBUSY|EPERM|EACCES/.test(err.code ?? err.message)) throw fail('문서가 다른 프로그램에서 열려 있어 되돌릴 수 없습니다. 프로그램을 닫고 다시 시도하세요.');
        throw err;
      }
      doc.stat = statOf(doc.path);
      doc.missing = false;
      // 편집 창이 열려 있으면 되돌린 내용으로 다시 불러오게 한다 (저장 안 한 편집이 있으면 편집 창이 물어본다).
      // 파일이 이미 원본과 같아도(열고 고치기만 하고 저장 안 함) 저장 안 한 편집이 있으면 알린다 → 버릴지 고르게
      const editorNotified = !!editors?.notifyExternalChange(doc.path, { force: !!editors?.isDirty(doc.path), reason: 'restore' });
      // AI가 기억하는 문서 내용은 이제 틀리므로 다음 메시지부터 새 대화로 시작한다
      const newChat = sessions.delete(k);
      return { time: entry.time, reason: entry.reason ?? null, file: entry.file, stoppedAi, newChat, editorNotified, doc: current === doc ? docInfo() : null };
    } finally {
      restoring.delete(k);
    }
  }

  /**
   * 버전(백업) 이후 바뀐 글 → { file, time, reason, kind: 'text'|'cells', hunks|cells, stats, same, … } (core/diff.mjs)
   * 지금 파일은 사본으로 읽는다 (다른 프로그램이 열고 있어도 읽히게, 읽는 사이 바뀌어도 한 시점으로)
   */
  let loadedDiff = null;
  async function diffVersion(args) {
    const doc = needDoc();
    const file = args && typeof args === 'object' && typeof args.file === 'string' ? args.file : null;
    if (!file) throw fail('비교할 버전을 골라 주세요');
    const entry = backups.find(doc.path, file);
    if (!entry) throw fail('그 버전을 찾을 수 없어요 (정리됐을 수 있어요). 목록을 다시 열어 주세요.', 'NO_VERSION');
    const m = await (loadedDiff ??= import('../core/diff.mjs'));
    const work = mkdtempSync(join(tmpdir(), 'officekit-diff-'));
    try {
      const now = join(work, `now.${extOf(doc.path)}`);
      try { copyFileSync(doc.path, now); } catch (err) { throw fail(`문서를 읽지 못했습니다: ${err.message}`); }
      let d;
      try { d = await m.diffDocuments(entry.path, now); } catch (err) {
        throw Object.assign(fail(`바뀐 내용을 비교하지 못했어요: ${err?.message ?? err}`, 'DIFF_FAILED'), { detail: err?.stack ?? '' });
      }
      return { file: entry.file, time: entry.time, reason: entry.reason ?? null, ...d };
    } finally {
      try { rmSync(work, { recursive: true, force: true, maxRetries: 3 }); } catch {}
    }
  }

  // ── 개인정보 찾기·가리기 (core/pii.mjs) ──────────────
  let loadedPii = null;
  const loadPii = async () => (loadedPii ??= await import('../core/pii.mjs'));

  async function piiScan() {
    const doc = needDoc();
    const m = await loadPii();
    try {
      const r = await m.scanDocument(doc.path);
      return { path: doc.path, items: r.items, total: r.total, chars: r.chars, types: m.PII_TYPES, output: m.maskedPathFor(doc.path) };
    } catch (err) { throw Object.assign(fail(`개인정보를 찾지 못했어요: ${err?.message ?? err}`, 'PII_FAILED'), { detail: err?.stack ?? '' }); }
  }

  /** 가린 사본 만들기 (원본은 그대로, 원본 옆 '<이름>_가림.<확장자>') → { output, items, remaining } */
  async function piiMask(args) {
    const doc = needDoc();
    const values = Array.isArray(args?.values) ? args.values.filter((v) => typeof v === 'string' && v.length <= 200).slice(0, 2000) : null;
    if (!values?.length) throw fail('가릴 항목을 하나 이상 골라 주세요');
    const m = await loadPii();
    const output = m.maskedPathFor(doc.path);
    let r;
    try { r = await m.maskDocument(doc.path, { output, values }); } catch (err) {
      if (/EBUSY|EPERM|EACCES/.test(err?.code ?? err?.message)) throw fail('가린 사본을 저장하지 못했어요. 문서 폴더에 쓸 수 없거나 같은 이름의 파일이 열려 있어요.');
      throw Object.assign(fail(`가린 사본을 만들지 못했어요: ${err?.message ?? err}`, 'PII_FAILED'), { detail: err?.stack ?? '' });
    }
    if (!SMOKE) { try { settings.addRecent(r.output); } catch {} }
    return { output: r.output, items: r.items, remaining: r.remaining, engine: r.engine };
  }

  // ── 한글 → HWPX 일괄 변환 (core/batch.mjs) ────────────
  let loadedBatch = null;
  const loadBatch = async () => (loadedBatch ??= await import('../core/batch.mjs'));
  let batchRun = null; // { cancelled, results }
  let batchLast = null;

  const pathList = (v) => (Array.isArray(v) ? v : []).filter((p) => typeof p === 'string' && isAbsolute(p) && p.length <= 4096).slice(0, 500);

  async function batchPick({ kind } = {}) {
    const folder = kind === 'folder';
    const r = await dialog.showOpenDialog(win, folder
      ? { title: '한글 문서가 든 폴더 고르기', properties: ['openDirectory', 'multiSelections'], defaultPath: current ? dirname(current.path) : app.getPath('documents') }
      : { title: '한글 문서 고르기', properties: ['openFile', 'multiSelections'], defaultPath: current ? dirname(current.path) : app.getPath('documents'), filters: [{ name: '한글 문서 (.hwp)', extensions: ['hwp'] }] });
    return r.canceled ? [] : r.filePaths;
  }

  async function batchScan({ inputs, recursive } = {}) {
    const m = await loadBatch();
    const { items, truncated } = m.collectHwpFiles(pathList(inputs), { recursive: recursive !== false });
    return { count: items.length, truncated, files: items.slice(0, 200).map((it) => it.src) };
  }

  async function batchStart({ inputs, recursive, overwrite, outDir } = {}) {
    if (batchRun) throw fail('이미 변환하는 중이에요');
    const m = await loadBatch();
    const { items } = m.collectHwpFiles(pathList(inputs), { recursive: recursive !== false });
    if (!items.length) throw fail('바꿀 한글(.hwp) 파일이 없어요');
    const out = typeof outDir === 'string' && outDir && isAbsolute(outDir) ? resolve(outDir) : undefined;
    const run = { cancelled: false };
    batchRun = run;
    try {
      const r = await m.convertHwpBatch(items, {
        outDir: out, overwrite: overwrite === true,
        isCancelled: () => run.cancelled,
        onProgress: ({ index, total, result }) => send('batch:progress', { index, total, result }),
      });
      batchLast = r.results;
      return r;
    } finally { batchRun = null; }
  }

  async function batchSaveReport() {
    if (!batchLast?.length) throw fail('저장할 결과가 없어요');
    const m = await loadBatch();
    const dir = dirname(batchLast[0].out);
    const r = await dialog.showSaveDialog(win, { title: '변환 결과 보고서 저장', defaultPath: join(dir, 'HWPX 변환 결과.csv'), filters: [{ name: 'CSV (엑셀에서 열기)', extensions: ['csv'] }] });
    if (r.canceled || !r.filePath) return { canceled: true };
    const file = r.filePath.toLowerCase().endsWith('.csv') ? r.filePath : `${r.filePath}.csv`;
    m.writeBatchReport(file, batchLast);
    return { output: file };
  }

  // ── PDF 도구 (core/pdftools.mjs) ──────────────────────
  let loadedPdfTools = null;
  const loadPdfTools = async () => (loadedPdfTools ??= await import('../core/pdftools.mjs'));
  const PDF_ACTIONS = { merge: '합침', extract: '쪽뽑기', delete: '쪽지움', rotate: '돌림' };

  async function pdfPick() {
    const r = await dialog.showOpenDialog(win, { title: 'PDF 고르기', properties: ['openFile', 'multiSelections'],
      defaultPath: current ? dirname(current.path) : app.getPath('documents'), filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] });
    if (r.canceled) return [];
    const m = await loadPdfTools();
    const out = [];
    for (const p of r.filePaths.slice(0, 200)) {
      let pages = null; let error = null;
      try { pages = await m.pdfPageCount(p); } catch (err) { error = err.message; }
      out.push({ path: p, name: basename(p), pages, error });
    }
    return out;
  }

  async function pdfRun({ action, inputs, pages, degrees } = {}) {
    if (!Object.hasOwn(PDF_ACTIONS, action)) throw fail('잘못된 작업입니다');
    const list = pathList(inputs).filter((p) => extOf(p) === 'pdf');
    if (!list.length) throw fail('PDF를 골라 주세요');
    if (action === 'merge' && list.length < 2) throw fail('합칠 PDF를 둘 이상 골라 주세요');
    const spec = typeof pages === 'string' ? pages.slice(0, 500) : '';
    const first = list[0];
    const stem = basename(first).replace(/\.[^.]+$/, '');
    let suggest = join(dirname(first), `${stem}_${PDF_ACTIONS[action]}.pdf`);
    for (let i = 2; existsSync(suggest); i++) suggest = join(dirname(first), `${stem}_${PDF_ACTIONS[action]} (${i}).pdf`);
    // 쪽 번호가 맞는지는 저장할 곳을 묻기 전에 확인한다
    const m = await loadPdfTools();
    if (action !== 'merge') {
      let total;
      try { total = await m.pdfPageCount(first); } catch (err) { throw fail(err.message); }
      if (action !== 'rotate' || spec.trim()) { try { m.parsePages(spec, total); } catch (err) { throw fail(err.message, 'BAD_PAGES'); } }
    }
    const r = SMOKE && process.env.OFFICEKIT_SMOKE_PDFTOOLS_OUT ? { canceled: false, filePath: resolve(process.env.OFFICEKIT_SMOKE_PDFTOOLS_OUT) }
      : await dialog.showSaveDialog(win, { title: '결과 PDF 저장', defaultPath: suggest, filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] });
    if (r.canceled || !r.filePath) return { canceled: true };
    const output = r.filePath.toLowerCase().endsWith('.pdf') ? r.filePath : `${r.filePath}.pdf`;
    try {
      const res = action === 'merge' ? await m.mergePdfs(list, output)
        : action === 'extract' ? await m.extractPages(first, spec, output)
          : action === 'delete' ? await m.deletePages(first, spec, output)
            : await m.rotatePages(first, Number(degrees) || 90, spec, output);
      if (!SMOKE) { try { settings.addRecent(res.output); } catch {} }
      return res;
    } catch (err) {
      if (/EBUSY|EPERM|EACCES/.test(err?.code ?? err?.message)) throw fail('저장하지 못했어요 — 같은 이름의 파일이 열려 있거나 쓰기 권한이 없어요');
      throw fail(err.message);
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

  /**
   * PDF를 Word 창에서 연다. Word가 'PDF를 편집 가능한 문서로 변환' 확인 창을 띄우므로 자동(COM)으로는 하지 않고,
   * Word를 그냥 실행해 사용자가 확인하게 한다 (Word 설정은 바꾸지 않는다). Start-Process가 App Paths로 winword를 찾는다
   */
  async function openPdfInWord() {
    const doc = needDoc();
    if (extOf(doc.path) !== 'pdf') throw fail('PDF 문서만 Word로 바꿔 열 수 있어요');
    const apps = await com.probe().catch(() => ({}));
    if (!apps.word) throw fail('Word가 설치돼 있지 않아 PDF를 고칠 수 없어요. 원본 문서를 고친 뒤 PDF로 다시 저장해 주세요.');
    const ps = "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue'; Start-Process -FilePath 'winword.exe' -ArgumentList ('\"' + $env:OK_FILE + '\"')";
    const enc = Buffer.from(ps, 'utf16le').toString('base64');
    const { execFile } = await import('node:child_process');
    await new Promise((done, bad) => execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', enc],
      { env: { ...process.env, OK_FILE: doc.path }, windowsHide: true, timeout: 30_000 },
      (err) => (err ? bad(fail(`Word를 열지 못했어요: ${String(err.message).trim().split(/\r?\n/)[0]}`)) : done())));
    return { opened: true };
  }

  /**
   * 연결 프로그램으로 열기 — 단 Office Kit 자신은 빼고 (Office Kit가 그 형식의 기본 프로그램이면 자기를 다시 열게 되므로).
   * 다른 프로그램이 없으면 Windows '연결 프로그램' 고르기 창을 띄운다
   */
  async function openOutside(p) {
    if (process.platform !== 'win32') { const e = await shell.openPath(p); if (e) throw fail(`프로그램에서 열지 못했습니다: ${e}`); return; }
    const m = await loadAssoc();
    const cmd = await m.otherOpenCommand(extOf(p));
    const run = cmd ? m.parseOpenCommand(cmd, p) : null;
    const { spawn } = await import('node:child_process');
    const launch = (file, args, verbatim) => new Promise((done, bad) => {
      try {
        const child = spawn(file, args, { detached: true, stdio: 'ignore', windowsHide: false, windowsVerbatimArguments: verbatim });
        child.once('error', bad);
        child.once('spawn', () => { child.unref(); done(); });
      } catch (err) { bad(err); }
    });
    if (run) {
      try { await launch(run.file, [run.args], true); return; } catch (err) { console.error('[open-outside]', err?.message ?? err); }
    }
    // 다른 프로그램을 못 찾음 → '연결 프로그램' 고르기
    await launch('rundll32.exe', ['shell32.dll,OpenAs_RunDLL', p], false);
  }

  async function openInProgram() {
    const doc = needDoc();
    const ext = extOf(doc.path);
    // 인터넷에서 받은 문서: 자동화(COM)로 열면 보호된 보기를 거치지 않는다 → Windows 연결 프로그램으로 (Office가 보호된 보기로 연다)
    if (!core.isHwp(ext) && core.isFromInternet(doc.path)) { await openOutside(doc.path); return { via: 'shell' }; }
    const apps = await com.probe().catch(() => ({}));
    // 한글은 보안 모듈이 없으면 COM 대신 연결 프로그램으로 연다 (허용 창에서 멈추지 않게)
    if (com.comAppFor(ext) === 'hwp' && !apps.hancomSecurityModule) {
      await openOutside(doc.path);
      return { via: 'shell' };
    }
    try {
      const r = await core.openInApp(doc.path);
      return { via: 'app', readOnly: !!r?.readOnly, readOnlyReason: r?.readOnlyReason ?? null };
    } catch (err) {
      try { await openOutside(doc.path); } catch { throw fail(`프로그램에서 열지 못했습니다: ${err.message}`); }
      return { via: 'shell' };
    }
  }

  // PDF로 저장은 두 번에 나눈다: 저장 대화상자(step 'pick') → 고른 곳에 만들기(step 'make').
  // 화면은 자리를 고른 뒤에만 'PDF 만드는 중'을 보인다. 만들 곳은 main이 기억한 것만 쓴다 (화면이 경로를 주지 않음)
  let pdfTarget = null;
  async function savePdf({ step } = {}) {
    const doc = needDoc();
    if (extOf(doc.path) === 'pdf') throw fail('이미 PDF 문서예요');
    if (step === 'make') {
      const t = pdfTarget;
      pdfTarget = null;
      if (!t || !samePath(t.doc, doc.path)) throw fail('저장할 곳을 다시 골라 주세요');
      const res = await core.convertDocument(doc.path, t.output);
      return { output: res.output, engine: res.engine, pages: res.pages ?? null, note: res.note ?? null };
    }
    await pdfCheck(doc.path);
    const stem = basename(doc.path).replace(/\.[^.]+$/, '');
    const r = SMOKE_PDF ? { canceled: false, filePath: SMOKE_PDF } : await dialog.showSaveDialog(win, {
      title: 'PDF로 저장', defaultPath: join(dirname(doc.path), `${stem}.pdf`),
      filters: [{ name: 'PDF 문서', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { canceled: true };
    const output = r.filePath.toLowerCase().endsWith('.pdf') ? r.filePath : `${r.filePath}.pdf`;
    if (samePath(output, doc.path)) throw fail('원본 문서와 같은 이름으로는 저장할 수 없습니다');
    if (step === 'pick') { pdfTarget = { doc: doc.path, output }; return { output, picked: true }; }
    const res = await core.convertDocument(doc.path, output);
    return { output: res.output, engine: res.engine, pages: res.pages ?? null, note: res.note ?? null };
  }

  /**
   * PDF로 저장할 수 있는지 저장 대화상자 전에 확인.
   * 한글: 한컴오피스 + 자동화 보안 모듈이 있으면 한컴으로, 없으면 내장 엔진(rhwp)으로 만든다 (core.convertDocument가 고름)
   */
  async function pdfCheck(p) {
    if (!core.isHwp(extOf(p))) return appCheck(p, 'PDF로 저장');
    if (await com.comAvailableFor(extOf(p))) return;
    const st = core.hwpPdfStatus();
    if (!st.ok) throw fail(`PDF로 저장할 수 없습니다: ${st.reason}`);
  }

  // ── 새 문서 ──────────────────────────────────────────

  const NEW_DOC_TOOLS = { 한글: 'hwp', Word: 'word', PowerPoint: 'ppt', Excel: 'excel' };
  const NEW_DOC_REQUEST_MAX = 10_000;
  const isDir = (p) => { try { return statSync(p).isDirectory(); } catch { return false; } };

  /** 새 문서 기본 폴더: 열린 문서의 폴더, 없으면 문서 폴더 */
  function newDocFolder() {
    const dir = current ? dirname(current.path) : null;
    return dir && isDir(dir) ? dir : app.getPath('documents');
  }

  function blankFormat(format) {
    const f = typeof format === 'string' ? format.toLowerCase() : '';
    if (!create.BLANK_FORMATS.includes(f)) throw fail('지원하지 않는 형식입니다', 'BAD_FORMAT');
    return f;
  }

  /** 새 문서 창 기본값 → { folder, name, documents, formats }. folder를 주면 그 폴더에서 겹치지 않는 이름 */
  function newDocInfo({ format = create.BLANK_FORMATS[0], folder } = {}) {
    const ext = blankFormat(format);
    const dir = typeof folder === 'string' && folder.length <= 4096 && isAbsolute(folder) ? resolve(folder) : newDocFolder();
    const templates = Object.fromEntries(create.BLANK_FORMATS.map((f) => [f, templatesMod.templatesFor(f)]));
    return { folder: dir, name: create.freeDocName(dir, ext), documents: app.getPath('documents'), formats: [...create.BLANK_FORMATS], templates };
  }

  /** 입력하는 동안 이름 확인 → { ok, path, name, code, error, suggestion } (오류를 던지지 않고 돌려준다) */
  function newDocCheck({ folder, name, format } = {}) {
    try {
      const plan = create.planBlankDocument({ folder, name, format });
      return { ok: true, path: plan.path, name: plan.name, code: null, error: null, suggestion: null };
    } catch (err) {
      if (!err?.code) throw err;
      const suggestion = err.code === 'EXISTS' ? create.freeDocName(resolve(folder), format, create.cleanDocName(name, format)) : null;
      return { ok: false, path: null, name: null, code: err.code, error: err.message, suggestion };
    }
  }

  async function pickFolder(start) {
    const r = await dialog.showOpenDialog(win, {
      title: '새 문서를 저장할 폴더', properties: ['openDirectory'],
      defaultPath: typeof start === 'string' && isAbsolute(start) && isDir(start) ? start : newDocFolder(),
    });
    if (r.canceled || !r.filePaths[0]) return null;
    return resolve(r.filePaths[0]);
  }

  /** AI에게 줄 첫 메시지: 방금 만든 빈 문서를 요청대로 처음부터 채우게 */
  function newDocPrompt(path, request) {
    const kind = kindOf(path);
    const p = NEW_DOC_TOOLS[kind];
    return [
      `[새 문서] 이 ${kind} 문서(${basename(path)})는 방금 Office Kit에서 빈 문서로 만들었고 아직 내용이 없다. 아래 요청대로 내용을 처음부터 채운다.`,
      `- ${p}_inspect로 빈 문서 구조를 확인하고 ${p}_edit로 내용을 넣는다. 원고 하나로 한 번에 쓰는 편이 나으면 office_create로 이 경로(${path})에 overwrite=true로 다시 만들어도 된다.`,
      '- 다른 파일을 만들거나 이름을 바꾸지 말고 이 문서에 저장한다.',
      '- 다 채운 뒤 office_page_image로 쪽(슬라이드)을 보고 결과를 확인한다.',
      ...(p === 'hwp' ? ['- 한글 쪽 이미지는 내장 엔진(rhwp) 그림이라 한컴오피스와 쪽 나눔이 다를 수 있다. 쪽 나눔을 맞추려고 쪽 나누기·빈 문단을 넣지 말고, 필요하면 문단 서식의 keep_with_next(다음 문단과 함께)를 쓴다.'] : []),
      '',
      `요청: ${request}`,
    ].join('\n');
  }

  /**
   * 새 문서 만들기: 확인 → 빈 문서 → 열기(백업·감시·최근 목록) → (AI) 대화 시작.
   * → { doc, chat: { chatId, provider } | null, chatError: { message, code } | null }
   */
  async function newDocument(args) {
    if (!args || typeof args !== 'object' || Array.isArray(args)) throw fail('잘못된 요청입니다');
    const start = args.start ?? 'blank';
    if (!['blank', 'ai', 'template'].includes(start)) throw fail('잘못된 시작 방법입니다');
    let request = '';
    if (start === 'ai') {
      request = typeof args.request === 'string' ? args.request.trim() : '';
      if (!request) throw fail('AI에게 부탁할 내용을 적어 주세요', 'NO_REQUEST');
      if (request.length > NEW_DOC_REQUEST_MAX) throw fail(`부탁할 내용이 너무 길어요 (${NEW_DOC_REQUEST_MAX.toLocaleString('ko-KR')}자 이하)`, 'NO_REQUEST');
    }
    let plan;
    try { plan = create.planBlankDocument({ folder: args.folder, name: args.name, format: blankFormat(args.format) }); } catch (err) {
      throw err?.code ? fail(err.message, err.code) : err;
    }
    // 서식 파일: 원고로 만든다 (같은 이름이 있으면 거절 — 빈 문서와 같게)
    let tplMarkdown = null;
    if (start === 'template') {
      tplMarkdown = templatesMod.templateMarkdown(String(args.template ?? ''), extOf(plan.path));
      if (!tplMarkdown) throw fail('고른 서식을 이 형식으로 만들 수 없어요', 'NO_TEMPLATE');
    }
    try {
      if (tplMarkdown) {
        if (existsSync(plan.path)) throw Object.assign(new Error('exists'), { code: 'EXISTS' });
        await create.createDocument(plan.path, tplMarkdown);
      } else await create.createBlankDocument(plan.path);
    } catch (err) {
      if (err.code === 'EXISTS') throw fail(`이 폴더에 같은 이름의 파일이 이미 있어요: ${basename(plan.path)}`, 'EXISTS');
      if (['EACCES', 'EPERM', 'EROFS'].includes(err.code)) throw fail('이 폴더에는 파일을 만들 수 없어요 (권한이 없거나 읽기 전용). 다른 폴더를 골라 주세요.', 'WRITE_FAILED');
      if (err.code === 'ENOSPC') throw fail('디스크 공간이 부족해 문서를 만들지 못했어요', 'WRITE_FAILED');
      if (err.code === 'ENOENT') throw fail('저장할 폴더를 찾을 수 없어요', 'NO_FOLDER');
      throw fail(`문서를 만들지 못했어요: ${err.message}`, 'WRITE_FAILED');
    }
    // 같은 경로의 문서를 지웠다가 다시 만든 경우 예전 '원래대로' 기준·AI 대화를 잇지 않는다
    const k = key(plan.path);
    restorePoints.delete(k);
    sessions.delete(k);
    if (current && samePath(current.path, plan.path)) { stopWatch(); current = null; }
    const doc = await openDocument(plan.path);
    let chat = null; let chatError = null;
    if (start === 'ai') {
      // 문서는 이미 만들었으므로 AI를 못 띄워도 실패로 돌려주지 않는다 (화면이 대화창에 알린다)
      try { chat = await sendChat({ prompt: newDocPrompt(plan.path, request) }); } catch (err) {
        if (!err?.expected && err?.code !== 'MODULE_MISSING') console.error('[doc:new] AI', err);
        chatError = { message: err?.message ?? String(err), code: err?.code ?? null };
      }
    }
    // 편집 창으로 열 수 있는 형식(한글·docx·xlsx·pptx)은 만든 뒤 바로 편집 창으로 (못 열어도 문서는 만들었으므로 실패로 돌려주지 않는다)
    let editor = null; let editorError = null;
    if (args.edit === true && editorKind(plan.path)) {
      try { editor = await editDocument(); } catch (err) {
        editorError = { message: err?.message ?? String(err), code: err?.code ?? null };
      }
    }
    return { doc, chat, chatError, editor, editorError };
  }

  // ── 편집 창 (한글 · 문서·표·슬라이드) ─────────────────
  // editor-window.mjs가 문서마다 창 하나를 띄운다 (인터넷 없이). 종류(kind)는 확장자로 고른다:
  //   'hwp' = 한글 hwp·hwpx·hwt → 내장 rhwp-studio, 'office' = docx·xlsx·pptx → vendor/office-editor
  // 모듈을 못 불러와도 앱은 뜨고 '직접 편집'만 이유를 알린다

  let EDITOR_KINDS = { hwp: ['hwp', 'hwpx', 'hwt'], office: ['docx', 'xlsx', 'pptx'] };
  // 창 이름 ('ONLYOFFICE'는 쓰지 않는다 — 설계 8장)
  const EDITOR_LABELS = { hwp: '한글 편집', docx: '문서 편집', xlsx: '표 편집', pptx: '슬라이드 편집' };
  let OFFICE_DRAFTS = 'office-drafts';
  /** 파일 경로 → 'hwp' | 'office' | null */
  const editorKind = (p) => Object.keys(EDITOR_KINDS).find((k) => EDITOR_KINDS[k].includes(extOf(p))) ?? null;
  /** 파일 경로 → '한글 편집' | '문서 편집' | '표 편집' | '슬라이드 편집' */
  const editorLabel = (p) => (editorKind(p) === 'hwp' ? EDITOR_LABELS.hwp : EDITOR_LABELS[extOf(p)] ?? '문서 편집');
  let editors = null; let editorLoadError = null;
  try {
    const m = await import('./editor-window.mjs');
    if (m.EDITOR_KINDS && typeof m.EDITOR_KINDS === 'object') {
      EDITOR_KINDS = Object.fromEntries(Object.entries(m.EDITOR_KINDS).filter(([, v]) => Array.isArray(v)).map(([k, v]) => [k, [...v]]));
    } else if (Array.isArray(m.EDITABLE_EXTS)) EDITOR_KINDS = { ...EDITOR_KINDS, hwp: [...m.EDITABLE_EXTS] };
    if (typeof m.OFFICE_DRAFTS_DIR === 'string') OFFICE_DRAFTS = m.OFFICE_DRAFTS_DIR;
    // 개발 중 시험용: 편집기 묶음 폴더 바꾸기(없는 폴더 = 묶음 없음), 편집 창 시험 고리(window.__editorTest·__officeTest),
    // 글꼴 모듈 바꾸기 (설치본에서는 모두 무시)
    const dev = !app.isPackaged;
    const vendor = dev && process.env.OFFICEKIT_EDITOR_VENDOR ? resolve(process.env.OFFICEKIT_EDITOR_VENDOR) : null;
    let fonts;
    if (dev && process.env.OFFICEKIT_OFFICE_FONTS_MODULE) {
      try { fonts = await import(pathToFileURL(resolve(process.env.OFFICEKIT_OFFICE_FONTS_MODULE)).href); } catch (err) {
        console.error('[editor] 시험용 글꼴 모듈을 불러오지 못했습니다:', err?.message ?? err);
      }
    }
    editors = m.createEditorManager({
      app, BrowserWindow, ipcMain, session, dialog, shell, nativeTheme,
      appRoot: APP_ROOT, userData, icon: APP_ICON,
      dirs: vendor ? { studio: join(vendor, 'rhwp-studio'), sdk: join(vendor, 'rhwp-editor-sdk'), office: join(vendor, 'office-editor') } : undefined,
      ...(fonts ? { fonts } : {}),
      testHooks: dev && process.env.OFFICEKIT_EDITOR_TEST === '1',
      // 편집 창이 덮어쓰기 전에 백업한다 (열었던 문서들의 '원래대로' 기준은 정리하지 않게).
      // 자주 저장해도 다른 문서의 백업을 밀어내지 않도록 이 문서 몫은 최근 몇 개만 둔다
      getBackups: () => ({ backup: (p) => keepVersion(p, 'editor') }),
      // 문서·표·슬라이드 편집기의 설정(보기·확대 등)은 settings.json의 officeEditor.settings에
      officeState: { get: (k) => settings.officeGet(k), set: (k, v) => { settings.officeSet(k, v); } },
      onSaved: (p) => editorSaved(p),
      onDownloaded: (p) => editorDownloaded(p),
      onDirtyChange: () => pushEditorState(),
      onPanel: (p, ew) => { showPanel(p, ew).catch((err) => console.error('[panel]', err)); },
      // 한글 편집 창의 인쇄·PDF로 저장: 편집 중인 hwp 바이트 → PDF (한컴+자동화 연결이 있으면 한컴, 없으면 내장 엔진)
      toPdf: (bytes) => hwpBytesToPdf(bytes),
      toPrint: (bytes, ew) => printHwpBytes(bytes, ew),
      // 편집 창의 Ctrl+O · Ctrl+N
      onOpenRequest: (ew) => { pickAndOpen(ew).catch((err) => console.error('[open]', err)); },
      onNewRequest: () => { showHub(); send('menu', 'new'); },
      onClosed: () => setTimeout(editorClosed, 0),
    });
  } catch (err) {
    editorLoadError = err?.message ?? String(err);
    console.error('[editor] 편집 창 기능을 불러오지 못했습니다:', err);
  }

  /** 그 문서의 편집 창 상태 → { open, dirty } */
  function editorState(p) {
    if (!editors || !p) return { open: false, dirty: false };
    return { open: editors.isOpen(p), dirty: editors.isDirty(p) };
  }

  /**
   * 편집기 묶음이 있는지 → { exts(모든 형식), kinds: { hwp|office: { exts, available, reason } } } (reason은 한국어).
   * 화면은 문서 확장자로 종류를 골라 '직접 편집'·새 문서 '바로 편집 창 열기'를 켜고 끈다
   */
  function editorInfo() {
    const kinds = {};
    for (const [kind, exts] of Object.entries(EDITOR_KINDS)) {
      let a = { available: false, reason: `편집 창 기능을 불러오지 못했습니다: ${editorLoadError}` };
      if (editors) { try { a = editors.available(kind); } catch (err) { a = { available: false, reason: err?.message ?? String(err) }; } }
      kinds[kind] = { exts: [...exts], available: !!a.available, reason: a.reason ?? null };
    }
    return { exts: Object.values(kinds).flatMap((k) => k.exts), kinds };
  }

  // ── 자동 저장 사본 정리 (Office) ─────────────────────
  // 편집기의 10분 자동 저장은 userData/office-drafts/<경로 sha1>.<ext>에만 쓴다. 정상 저장·정상 닫기에서는 편집 창이 지운다.
  // 강제 종료·충돌로 남은 사본은 다음에 그 문서를 열 때 '복구 / 버리기'로 묻고, 오래 안 연 것은 여기서 지운다
  const DRAFT_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const DRAFT_KEEP = 50;
  const DRAFT_TMP_AGE_MS = 60 * 60 * 1000;
  const DRAFT_NAME = /^[0-9a-f]{40}\.(docx|xlsx|pptx|hwp|hwpx|hwt)$/i;

  /** 30일 넘은 사본, 쓰다 남은 임시 파일, 50개 넘는 것(오래된 것부터) 지우기. 열린 편집 창 문서의 사본은 남긴다 → 지운 개수 */
  function cleanOfficeDrafts({ now = Date.now() } = {}) {
    const dir = join(userData, OFFICE_DRAFTS);
    let names;
    try { names = readdirSync(dir); } catch { return 0; }
    const inUse = new Set();
    for (const p of editors?.list() ?? []) {
      const d = editors.draftPath?.(p);
      if (d) inUse.add(basename(d).toLowerCase());
    }
    let removed = 0;
    const drop = (file) => { try { rmSync(file, { force: true }); removed++; } catch {} };
    const drafts = [];
    for (const name of names) {
      const file = join(dir, name);
      let st;
      try { st = statSync(file); } catch { continue; }
      if (!st.isFile() || inUse.has(name.toLowerCase())) continue;
      const age = now - st.mtimeMs;
      if (!DRAFT_NAME.test(name)) { if (age > DRAFT_TMP_AGE_MS) drop(file); continue; }
      if (age > DRAFT_MAX_AGE_MS) drop(file);
      else drafts.push({ file, mtimeMs: st.mtimeMs });
    }
    drafts.sort((a, b) => b.mtimeMs - a.mtimeMs);
    for (const d of drafts.slice(DRAFT_KEEP)) drop(d.file);
    if (removed) console.log(`[editor] 오래된 자동 저장 사본 ${removed}개를 지웠습니다`);
    return removed;
  }

  // 편집 창이 열리고 닫히고 고쳐질 때 화면(상태 막대·대화창 경고)에 알린다.
  // 창이 닫힌 뒤 정리·'다른 이름으로 저장' 뒤 경로 바뀜이 끝난 다음에 보도록 조금 늦춘다
  let editorPushTimer = 0;
  function pushEditorState() {
    clearTimeout(editorPushTimer);
    editorPushTimer = setTimeout(() => {
      if (current) send('doc:editor', { path: current.path, ...editorState(current.path) });
    }, 40);
  }

  /** 편집 창이 저장함 → 바깥 변경과 같은 길로 미리보기를 새로 고치고 최근 목록에 올린다 (감시보다 먼저) */
  function editorSaved(p) {
    if (!SMOKE) {
      try { settings.addRecent(p); } catch {}
      try { app.addRecentDocument(p); } catch {}
    }
    if (current && samePath(current.path, p)) {
      const st = statOf(p);
      current.stat = st;
      current.missing = !st;
      if (st) send('doc:changed', { path: current.path, missing: false, mtime: st.mtimeMs, size: st.size, by: 'editor' });
    }
    pushEditorState();
  }

  /** 편집 창이 다른 형식으로 내려받음 (PDF 등) → 이 앱으로 열 수 있는 형식이면 최근 목록에 */
  function editorDownloaded(p) {
    if (SMOKE || typeof p !== 'string' || !OPENABLE_EXTS.includes(extOf(p))) return;
    try { settings.addRecent(p); } catch {}
    try { app.addRecentDocument(p); } catch {}
  }

  /**
   * 열린 문서를 편집 창으로 연다 (이미 열려 있으면 그 창을 앞으로) → { path, kind, reused, notice }
   * notice: Office 편집 창이 많이 열렸을 때 메모리 안내 (없으면 null)
   */
  /** 한글 문서를 내장 엔진(rhwp)으로 열어 본다 → 문제 없으면 null, 있으면 { code, message, detail } */
  function hwpOpenProblem(p) {
    try { const h = core.Hwp.openForView(p); h.close(); return null; } catch (err) {
      const why = String(err?.message ?? err);
      if (/암호|password|encrypt|distribut|배포용/i.test(why)) {
        return { code: 'ENCRYPTED', message: '암호가 걸렸거나 배포용인 한글 문서라 직접 편집할 수 없어요. ‘프로그램에서 열기’로 열어 주세요.', detail: why };
      }
      return { code: 'BAD_FORMAT', message: `‘${basename(p)}’ 파일은 손상됐거나 지원하지 않는 한글 문서라 직접 편집할 수 없어요. ‘프로그램에서 열기’로 열어 보세요.`, detail: why };
    }
  }

  async function editDocument() {
    const doc = needDoc();
    const kind = editorKind(doc.path);
    if (!kind && extOf(doc.path) === 'pdf') {
      throw fail('PDF는 앱 안에서 바로 고칠 수 없어요. 글자를 고치려면 원본 문서(한글·Word 등)를 열어 고친 뒤 PDF로 다시 저장해 주세요. AI 대화로 내용을 묻거나 요약하는 것은 PDF에서도 돼요.', 'PDF_EDIT');
    }
    if (!kind) {
      const ext = extOf(doc.path);
      throw fail(`${kindOf(doc.path)} 문서(.${ext || '확장자 없음'})는 앱 안에서 직접 편집할 수 없어요. 직접 편집은 한글(HWP·HWPX)·Word(DOCX)·Excel(XLSX)·PowerPoint(PPTX) 문서만 돼요. ‘프로그램에서 열기’로 고쳐 주세요.`, 'UNSUPPORTED');
    }
    const label = editorLabel(doc.path);
    // 한글: 내장 엔진이 못 여는 문서(손상·암호·지원 안 하는 판)는 창을 띄우기 전에 거른다 (Office는 editor-window가 거름)
    if (kind === 'hwp' && !editors?.isOpen(doc.path)) {
      const bad = hwpOpenProblem(doc.path);
      if (bad) throw Object.assign(fail(bad.message, bad.code), { detail: bad.detail });
    }
    if (!editors) throw fail(`${label} 기능을 불러오지 못했습니다: ${editorLoadError}`, 'EDITOR_MISSING');
    const a = editors.available(kind);
    if (!a.available) throw fail(a.reason || `${label}기 파일이 없습니다.`, 'EDITOR_MISSING');
    let r;
    try { r = await editors.openEditor(doc.path); } catch (err) {
      // 암호 문서·옛 형식·손상·너무 큼 등은 한국어 이유와 code(ENCRYPTED·OLD_FORMAT·BAD_FORMAT·TOO_BIG …)를 그대로 넘긴다
      if (err?.expected) throw err;
      console.error('[doc:edit]', err);
      throw fail(err?.message ?? String(err), err?.code ?? 'EDITOR_FAILED');
    }
    pushEditorState();
    return { path: r.path, kind: r.kind ?? kind, reused: !!r.reused, notice: typeof r.notice === 'string' && r.notice ? r.notice : null };
  }

  /** hwp 바이트 → PDF 바이트 (임시 폴더에서) */
  async function hwpBytesToPdf(bytes) {
    const work = mkdtempSync(join(tmpdir(), 'officekit-hwppdf-'));
    try {
      const src = join(work, 'doc.hwp');
      const out = join(work, 'doc.pdf');
      writeFileSync(src, bytes);
      await core.convertDocument(src, out);
      return new Uint8Array(readFileSync(out));
    } finally { try { rmSync(work, { recursive: true, force: true, maxRetries: 3 }); } catch {} }
  }

  // ── 인쇄 ─────────────────────────────────────────────
  // 한글: 내장 엔진이 그린 쪽 그림(SVG)을 숨은 창에 펴 놓고 Windows 인쇄 대화상자로 (쪽 크기·방향은 쪽마다).
  // 숨은 창은 스크립트를 끄고, 파일·data: 말고는 아무 데도 읽으러 가지 않는다 (문서 내용이 밖으로 나가지 않게)
  let printSession = null;
  let printAllowedUrl = null;
  function printSes() {
    if (printSession) return printSession;
    printSession = session.fromPartition('officekit-print');
    // 지금 인쇄하는 print.html과 data: 그림만 (다른 파일·네트워크 공유(file://서버)·인터넷은 막는다)
    printSession.webRequest.onBeforeRequest((d, cb) => cb({ cancel: !(d.url.startsWith('data:') || (printAllowedUrl && d.url === printAllowedUrl)) }));
    printSession.setPermissionRequestHandler((_w, _p, cb) => cb(false));
    return printSession;
  }

  async function printHwpPath(path, parent) {
    const h = core.Hwp.openForView(path);
    let pages;
    try { pages = Array.from({ length: h.pageCount }, (_, i) => h.svg(i)); } finally { h.close(); }
    if (!pages.length) throw fail('인쇄할 쪽이 없습니다');
    const mm = (px) => (Number(px) * 25.4 / 96).toFixed(2);
    const css = []; const body = [];
    pages.forEach((svg, i) => {
      const m = /<svg\b[^>]*\bwidth="([\d.]+)"[^>]*\bheight="([\d.]+)"/.exec(svg);
      const w = m ? mm(m[1]) : '210'; const hgt = m ? mm(m[2]) : '297';
      css.push(`@page p${i} { size: ${w}mm ${hgt}mm; margin: 0 } .p${i} { page: p${i}; width: ${w}mm; height: ${hgt}mm }`);
      body.push(`<div class="pg p${i}">${svg}</div>`);
    });
    const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>인쇄</title><style>html,body{margin:0;padding:0;background:#fff}
      .pg{overflow:hidden;break-after:page}.pg:last-child{break-after:auto}.pg>svg{display:block;width:100%;height:100%}${css.join('\n')}</style></head><body>${body.join('')}</body></html>`;
    const dir = join(userData, 'print', randomUUID().slice(0, 12));
    mkdirSync(dir, { recursive: true });
    const file = join(dir, 'print.html');
    writeFileSync(file, html, 'utf8');
    const pw = new BrowserWindow({ show: false, width: 800, height: 1100, ...(parent && !parent.isDestroyed() ? { parent } : {}),
      webPreferences: { session: printSes(), sandbox: true, contextIsolation: true, nodeIntegration: false, javascript: false, spellcheck: false, devTools: false } });
    pw.webContents.on('will-navigate', (e) => e.preventDefault());
    pw.webContents.on('will-frame-navigate', (e) => { if (e.url !== printAllowedUrl) e.preventDefault(); });
    pw.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    try {
      printAllowedUrl = pathToFileURL(file).href;
      await pw.loadFile(file);
      // 확인용: 인쇄 대화상자 대신 PDF로 (쪽 크기·쪽 수 확인)
      if (SMOKE && process.env.OFFICEKIT_SMOKE_PRINT) {
        writeFileSync(resolve(process.env.OFFICEKIT_SMOKE_PRINT), await pw.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true }));
        return { printed: true, reason: null };
      }
      return await new Promise((done) => {
        pw.webContents.print({ silent: false, printBackground: true }, (ok, reason) => done({ printed: !!ok, reason: ok ? null : String(reason ?? 'cancelled') }));
      });
    } finally {
      printAllowedUrl = null;
      try { pw.destroy(); } catch {}
      try { rmSync(dir, { recursive: true, force: true }); } catch {}
    }
  }

  /** 인쇄용 사본(userData/print) 정리: 한 시간 넘은 것 (끝낼 때는 모두) — 문서 사본이 남지 않게 */
  function cleanPrintCopies({ all = false } = {}) {
    const root = join(userData, 'print');
    let names;
    try { names = readdirSync(root); } catch { return; }
    for (const n of names) {
      const p = join(root, n);
      try { if (all || Date.now() - statSync(p).mtimeMs > 60 * 60 * 1000) rmSync(p, { recursive: true, force: true }); } catch {}
    }
  }

  /** 한글 편집 창의 인쇄: 편집 중인 hwp 바이트 → 임시 파일 → 인쇄 대화상자 */
  async function printHwpBytes(bytes, parent) {
    const work = mkdtempSync(join(tmpdir(), 'officekit-print-'));
    try {
      const src = join(work, 'doc.hwp');
      writeFileSync(src, bytes);
      return await printHwpPath(src, parent);
    } finally { try { rmSync(work, { recursive: true, force: true, maxRetries: 3 }); } catch {} }
  }

  /** 주 창의 인쇄 (Ctrl+P): 한글은 인쇄 대화상자, PDF·Word·Excel·PowerPoint는 PDF로 만들어 PDF 프로그램에서 */
  async function printCurrent() {
    const doc = needDoc();
    const ext = extOf(doc.path);
    if (core.isHwp(ext)) return printHwpPath(doc.path, win);
    let pdf = doc.path;
    if (ext !== 'pdf') {
      await appCheck(doc.path, '인쇄');
      const dir = join(userData, 'print', randomUUID().slice(0, 12));
      mkdirSync(dir, { recursive: true });
      const work = mkdtempSync(join(tmpdir(), 'officekit-print-'));
      try {
        const tmp = join(work, 'print.pdf');
        await core.convertDocument(doc.path, tmp);
        pdf = join(dir, `${basename(doc.path).replace(/\.[^.]+$/, '')}.pdf`);
        copyFileSync(tmp, pdf);
      } finally { try { rmSync(work, { recursive: true, force: true }); } catch {} }
    }
    await openOutside(pdf);
    return { opened: true };
  }

  /** 편집 창의 Ctrl+O: 문서를 골라 열기 (편집할 수 있으면 새 편집 창, 아니면 주 창) */
  async function pickAndOpen(parent) {
    const all = OPENABLE_EXTS;
    const r = await dialog.showOpenDialog(parent && !parent.isDestroyed() ? parent : win, {
      title: '문서 열기', properties: ['openFile', 'multiSelections'],
      defaultPath: current ? dirname(current.path) : app.getPath('documents'),
      filters: [{ name: '모든 문서', extensions: all }],
    });
    if (r.canceled) return;
    for (const f of r.filePaths) {
      if (wantsDocWindow(f)) await openAsDocument(f);
      else await openInHub(f);
    }
  }

  // ── 문서 창 (Word처럼) ───────────────────────────────

  /** 옛 형식 → 편집기로 열 새 형식 (설치된 Word·Excel·PowerPoint가 바꾼다. csv는 Office Kit가 직접) */
  const CONVERT_FOR_EDIT = { doc: 'docx', rtf: 'docx', odt: 'docx', xls: 'xlsx', xlsm: 'xlsx', ods: 'xlsx', csv: 'xlsx', ppt: 'pptx', odp: 'pptx' };
  const wantsDocWindow = (p) => settings.get('openAction') !== 'preview' && OPENABLE_EXTS.includes(extOf(p))
    && (!!editorKind(p) || !!CONVERT_FOR_EDIT[extOf(p)]);

  /** 주 창을 (옆 창이 아닌) 보통 창으로 보인다 */
  function showHub() {
    if (!win) return;
    leavePanel();
    if (win.isMinimized()) win.restore();
    win.show(); win.focus();
    if (hubMaximizeLater) { hubMaximizeLater = false; win.maximize(); }
  }

  /** 주 창에서 문서 열기 (창이 준비 전이면 준비되면 연다) */
  function openInHub(p) {
    showHub();
    if (rendererReady) return openDocument(p).catch((err) => send('toast', { type: 'error', message: err.message }));
    pendingFile = p;
    return null;
  }

  /** 문서를 열 때 하는 일 (주 창의 openDocument와 같게): 처음 연 상태 버전, 최근 목록 */
  function noteOpened(p) {
    const k = key(p);
    try {
      const made = keepVersion(p, 'open');
      if (!backupExists(restorePoints.get(k))) restorePoints.set(k, made);
    } catch (err) { console.error('[doc] 백업', err?.message ?? err); }
    if (!SMOKE) {
      try { settings.addRecent(p); } catch {}
      try { app.addRecentDocument(p); } catch {}
    }
  }

  // 옛 형식을 바꾸는 동안 띄우는 작은 창 (주 창이 숨어 있어 아무것도 안 보이지 않게)
  function busyWindow(text) {
    const w = new BrowserWindow({ width: 380, height: 96, frame: false, resizable: false, skipTaskbar: false, alwaysOnTop: true, show: false,
      title: 'Office Kit', icon: APP_ICON, backgroundColor: nativeTheme.shouldUseDarkColors ? '#1d2027' : '#ffffff',
      webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, javascript: false } });
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const dark = nativeTheme.shouldUseDarkColors;
    w.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(`<!doctype html><meta charset="utf-8"><body style="margin:0;font:13.5px 'Malgun Gothic',sans-serif;display:flex;align-items:center;gap:12px;height:100vh;padding:0 22px;box-sizing:border-box;color:${dark ? '#e6e8ec' : '#1f2328'};background:${dark ? '#1d2027' : '#fff'};border:1px solid ${dark ? '#333842' : '#d0d5dd'}"><b style="color:#3b5bdb">Office Kit</b><span>${esc(text)}</span>`)}`).catch(() => {});
    w.once('ready-to-show', () => { if (!w.isDestroyed()) w.show(); });
    return { close: () => { try { if (!w.isDestroyed()) w.destroy(); } catch {} } };
  }

  /** 옛 형식 → userData/converted/<…>/<이름>.<새 형식> (원본은 그대로) */
  async function convertForEditing(p, to) {
    const ext = extOf(p);
    const dir = join(userData, 'converted', `${Date.now()}-${randomUUID().slice(0, 8)}`);
    mkdirSync(dir, { recursive: true });
    const out = join(dir, `${basename(p).replace(/\.[^.]+$/, '')}.${to}`);
    try {
      if (ext === 'csv') {
        const doc = await core.readDocument(p, { engine: 'file' });
        const { default: ExcelJS } = await import('exceljs');
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet(basename(p).replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1');
        for (const row of doc.sheets?.[0]?.rows ?? []) ws.addRow(row.map((v) => (v !== '' && /^-?\d+(\.\d+)?$/.test(v) && v.length < 16 ? Number(v) : v)));
        await wb.xlsx.writeFile(out);
        return out;
      }
      if (!(await com.comAvailableFor(ext))) {
        throw fail(`.${ext} 문서를 고치려면 ${kindOf(p)} 프로그램이 필요해요 (새 형식으로 바꿔 열어야 해서). 미리보기로 열게요.`, 'NEED_APP');
      }
      if (core.isFromInternet(p)) throw fail(`${core.UNTRUSTED_APP_MESSAGE} 미리보기로 열게요.`, 'UNTRUSTED');
      // 프로그램(COM)은 MSIX 앱 안에서 가상화된 폴더를 못 보므로 임시 폴더에 받은 뒤 옮긴다
      const work = mkdtempSync(join(tmpdir(), 'officekit-convert-'));
      try {
        const tmpOut = join(work, `doc.${to}`);
        await core.convertDocument(p, tmpOut);
        copyFileSync(tmpOut, out);
      } finally { try { rmSync(work, { recursive: true, force: true }); } catch {} }
      return out;
    } catch (err) {
      try { rmSync(dir, { recursive: true, force: true }); } catch {}
      throw err;
    }
  }

  /**
   * 문서 창으로 열기 (탐색기에서 연 문서): 편집 창만 띄운다. 편집할 수 없는 문서(PDF·암호 문서 등)는 주 창에서 연다
   */
  async function openAsDocument(p) {
    p = docPath(p);
    const ext = extOf(p);
    const kind = editorKind(p);
    try {
      if (!editors) throw fail(`편집 창 기능을 불러오지 못했습니다: ${editorLoadError}`, 'EDITOR_MISSING');
      if (kind) {
        if (!editors.available(kind).available) throw fail('편집기가 없어 미리보기로 엽니다', 'EDITOR_MISSING');
        if (kind === 'hwp' && !editors.isOpen(p)) {
          const bad = hwpOpenProblem(p);
          if (bad) throw fail(bad.message, bad.code);
        }
        noteOpened(p);
        await editors.openEditor(p);
        return { path: p, mode: 'editor' };
      }
      const to = CONVERT_FOR_EDIT[ext];
      if (!to) return openInHub(p);
      if (!editors.available('office').available) throw fail('편집기가 없어 미리보기로 엽니다', 'EDITOR_MISSING');
      noteOpened(p);
      const busy = busyWindow(`‘${basename(p)}’를 고칠 수 있게 바꾸는 중…`);
      let tmp;
      try { tmp = await convertForEditing(p, to); } finally { busy.close(); }
      await editors.openEditor(tmp, { convertedFrom: p });
      return { path: p, mode: 'converted' };
    } catch (err) {
      // 편집 창으로 못 열면 주 창에서 미리보기로 (이유를 알린다)
      console.error('[doc-window]', err?.message ?? err);
      await openInHub(p);
      setTimeout(() => send('toast', { type: 'warn', message: err?.message ?? String(err) }), 400);
      return { path: p, mode: 'hub' };
    }
  }

  /** 편집 창이 닫힌 뒤: 주 창이 숨어 있거나 옆 창뿐이고 편집 창이 하나도 없으면 끝낸다 (Word처럼) */
  function editorClosed() {
    pushEditorState();
    if (!win || win.isDestroyed() || appQuitting) return;
    const left = editors?.list().length ?? 0;
    if (panelOf && (panelOf.isDestroyed() || left === 0)) {
      panelOf = null;
      if (!hubPrimary) { win.hide(); leavePanel(); }
    }
    if (left === 0 && !hubPrimary && !win.isVisible()) app.quit();
  }

  const PANEL_W = 420;

  /** 편집 창 옆에 옆 창(AI 대화·버전·도구)을 붙인다 */
  async function showPanel(p, ew) {
    if (!win || win.isDestroyed() || !ew || ew.isDestroyed()) return;
    if (!current || !samePath(current.path, p)) {
      if (!rendererReady) return;
      await openDocument(p);
    }
    if (!win.isVisible() || !panelOf) {
      if (win.isVisible() && !panelOf && !win.isMinimized()) hubBounds = win.getNormalBounds();
    }
    send('panel', { on: true, path: p });
    win.setMinimumSize(320, 480);
    if (win.isMaximized()) win.unmaximize();
    if (ew.isMaximized()) ew.unmaximize();
    const wa = screen.getDisplayMatching(ew.getBounds()).workArea;
    let eb = ew.getBounds();
    // 오른쪽에 자리가 없으면 편집 창을 줄인다
    if (eb.x + eb.width + PANEL_W > wa.x + wa.width) {
      const width = Math.max(720, wa.x + wa.width - PANEL_W - eb.x);
      const x = Math.max(wa.x, Math.min(eb.x, wa.x + wa.width - PANEL_W - width));
      eb = { ...eb, x, width };
      ew.setBounds(eb);
    }
    win.setBounds({ x: eb.x + eb.width, y: eb.y, width: PANEL_W, height: eb.height });
    win.show(); win.focus();
    panelOf = ew;
    if (!ew.__okPanelFollow) {
      ew.__okPanelFollow = true;
      const follow = () => {
        if (panelOf !== ew || ew.isDestroyed() || !win || win.isDestroyed() || !win.isVisible() || ew.isMinimized()) return;
        const b = ew.getBounds();
        win.setBounds({ x: b.x + b.width, y: b.y, width: win.getBounds().width, height: b.height });
      };
      ew.on('move', follow);
      ew.on('resize', follow);
    }
  }

  /** 옆 창 모드 끝 (주 창 크기를 되돌린다) */
  function leavePanel() {
    if (!win || win.isDestroyed()) return;
    const was = !!panelOf || win.getMinimumSize()[0] < 1000;
    panelOf = null;
    send('panel', { on: false });
    win.setMinimumSize(1000, 680);
    if (was) win.setBounds(hubBounds ?? initialBounds());
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

  async function sendChat({ prompt, editorOk } = {}) {
    if (typeof prompt !== 'string' || !prompt.trim()) throw fail('보낼 내용을 입력하세요');
    if (prompt.length > 20_000) throw fail('메시지가 너무 깁니다 (2만 자 이하)');
    const doc = needDoc();
    const path = doc.path;
    const k = key(path);
    if (restoring.has(k)) throw fail('문서를 원래대로 되돌리는 중입니다. 끝난 뒤 다시 보내 주세요.');
    // 편집 창에 저장 안 한 내용이 있으면 먼저 묻는다: AI는 저장된 파일을 고치므로 편집 창 내용과 엇갈린다
    // (화면이 '그대로 보내기'를 고르면 editorOk:true로 다시 보낸다)
    if (editorOk !== true && editors?.isDirty(path)) {
      throw fail('편집 창에 저장하지 않은 내용이 있어요. 편집 창에서 먼저 저장하거나, 그대로 보내기를 골라 주세요.', 'EDITOR_DIRTY');
    }
    const gen = restoreGen.get(k) ?? 0;
    const ai = await loadAi();
    const { providers: list, selected } = await providers();
    if (!selected) throw fail('사용할 수 있는 AI가 없습니다. Claude Code나 Codex를 설치하고 로그인하세요.', 'NO_PROVIDER');
    const base = await mcpSpec();
    // 앱 안 AI는 열린 문서의 폴더에만 저장할 수 있다 (MCP 서버가 OFFICEKIT_WRITE_ROOTS로 막는다)
    // 읽기는 그 폴더 안만, 고치기는 열린 문서와 AI가 새로 만든 파일만 (다른 파일은 덮어쓰지 않음)
    const mcp = { ...base, env: { ...(base.env ?? {}), OFFICEKIT_WRITE_ROOTS: dirname(path), OFFICEKIT_READ_ROOTS: dirname(path), OFFICEKIT_EDIT_FILES: path } };
    // 준비하는 사이 '원래대로'를 눌렀으면 시작하지 않는다 (되돌린 파일을 다시 고치지 않게)
    if (restoring.has(k) || (restoreGen.get(k) ?? 0) !== gen) throw fail('문서를 원래대로 되돌려서 이 요청은 보내지 않았습니다.');
    // AI가 고치기 전 상태를 버전으로 남긴다 (대화 한 번마다. 내용이 같으면 있던 것을 다시 씀)
    // → 끝날 때 이 버전과 다르면 화면에 '바뀐 내용 보기 · 이 대화 전으로 되돌리기'를 띄운다
    let before = null;
    try { before = keepVersion(path, 'ai'); } catch (err) { console.error('[ai:send] 백업', err?.message ?? err); }
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
      .finally(() => {
        chats.delete(chatId);
        let changed = false;
        try { changed = !!before?.hash && existsSync(path) && currentHash(path) !== before.hash; } catch {}
        send('ai:event', { type: 'end', chatId, path, ...(changed ? { changed: true, before: before.file } : {}) });
      });
    return { chatId, provider: list.find((p) => p.id === selected)?.name ?? selected };
  }

  function cancelChat(chatId) {
    if (typeof chatId !== 'string') throw fail('잘못된 요청입니다');
    const c = chats.get(chatId);
    if (c) { try { c.handle.cancel(); } catch {} }
    return { canceled: !!c };
  }

  // ── 설정 ─────────────────────────────────────────────

  const PUBLIC_KEYS = ['provider', 'models', 'previewMode', 'openAction', 'theme', 'chatWidth', 'chatCollapsed'];
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
      } else if (k === 'openAction') {
        if (!['edit', 'preview'].includes(v)) throw fail('잘못된 열기 방식입니다');
        clean.openAction = v;
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

  // ── 기본 프로그램 (Windows 파일 연결) ─────────────────
  // 설치 프로그램이 Office Kit를 기본 프로그램 후보로 등록한다. 고르는 것은 Windows 설정에서 사용자가 한다.
  // 설치된 앱에서만: 개발 실행(electron.exe)을 등록하면 탐색기가 엉뚱한 프로그램을 띄운다
  let loadedAssoc = null;
  const loadAssoc = async () => (loadedAssoc ??= await import('../setup/fileassoc.mjs'));
  const assocExe = () => (app.isPackaged && process.platform === 'win32' ? process.execPath : null);

  async function assocStatus() {
    const m = await loadAssoc();
    if (process.platform !== 'win32') return { supported: false, reason: 'Windows에서만 쓸 수 있어요', groups: [] };
    const exe = assocExe();
    const st = await m.assocStatus({ exe });
    const groups = m.ASSOC_GROUPS.map((g) => ({ kind: g.kind, types: st.types.filter((t) => g.exts.includes(t.ext)) }));
    return {
      supported: true, packaged: !!exe, registered: st.registered, matches: st.matches, groups,
      reason: exe ? null : '설치 프로그램으로 설치한 Office Kit에서만 기본 프로그램으로 정할 수 있어요 (지금은 개발 실행)',
    };
  }

  /** 등록이 없거나 다른 경로를 가리키면 고쳐 쓴다 (기본 프로그램은 바꾸지 않는다) */
  async function assocRepair() {
    const exe = assocExe();
    if (!exe) throw fail('설치 프로그램으로 설치한 Office Kit에서만 등록할 수 있어요');
    const m = await loadAssoc();
    await m.registerFileTypes({ exe });
    return assocStatus();
  }

  /** Windows 설정 > 기본 앱 > Office Kit 쪽을 연다 (등록이 어긋나 있으면 먼저 고친다) */
  async function assocOpenSettings() {
    const m = await loadAssoc();
    const st = await assocStatus();
    if (st.packaged && (!st.registered || !st.matches)) await assocRepair();
    await shell.openExternal(m.DEFAULT_APPS_URI);
    return null;
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
        // detail: 엔진이 낸 원래 글 (화면이 '자세히'로 접어 보여 준다)
        return { ok: false, error: err?.message ?? String(err), code: err?.code ?? null, ...(err?.detail ? { detail: String(err.detail).slice(0, 4000) } : {}) };
      }
    });
  }

  handle('app:init', () => {
    rendererReady = true;
    const opening = !!pendingFile;
    // 응답이 먼저 가도록 다음 틱에 연다
    if (pendingFile) {
      const f = pendingFile; pendingFile = null;
      setTimeout(() => openDocument(f, { autoEdit: true })
        .then(() => { if (SMOKE_DIFF && current) copyFileSync(SMOKE_DIFF, current.path); })
        .catch((err) => send('toast', { type: 'error', message: err.message })), 0);
    }
    return {
      version: app.getVersion(), packaged: app.isPackaged, platform: process.platform,
      exts: OPENABLE_EXTS, settings: publicSettings(), recent: settings.recent(), doc: docInfo(),
      editor: editorInfo(),
      smoke: SMOKE ? { chat: SMOKE_CHAT, settings: SMOKE_SETTINGS, newdoc: SMOKE_NEWDOC, editor: SMOKE_EDITOR, diff: !!SMOKE_DIFF, tool: SMOKE_TOOL, batch: SMOKE_BATCH } : null, opening,
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
        { name: 'PDF', extensions: ['pdf'] },
      ],
    });
    if (r.canceled || !r.filePaths[0]) return null;
    return openDocument(r.filePaths[0], { autoEdit: true });
  });
  handle('doc:open', (p) => openDocument(p, { autoEdit: true }));
  handle('doc:newInfo', (opts) => newDocInfo(opts && typeof opts === 'object' ? opts : {}));
  handle('doc:newCheck', (args) => newDocCheck(args && typeof args === 'object' ? args : {}));
  handle('doc:pickFolder', (start) => pickFolder(start));
  handle('doc:new', (args) => newDocument(args));
  handle('doc:close', () => { closeDocument(); return null; });
  handle('doc:preview', (opts = {}) => preview({
    mode: ['auto', 'exact', 'text'].includes(opts?.mode) ? opts.mode : undefined, ifChanged: opts?.ifChanged === true,
  }));
  handle('doc:restore', (args) => restore(args));
  handle('doc:versions', () => versions());
  handle('doc:diff', (args) => diffVersion(args));
  handle('pii:scan', () => piiScan());
  handle('pii:mask', (args) => piiMask(args));
  handle('batch:pick', (args) => batchPick(args && typeof args === 'object' ? args : {}));
  handle('batch:pickOut', async () => {
    const r = await dialog.showOpenDialog(win, { title: 'HWPX를 모을 폴더', properties: ['openDirectory', 'createDirectory'], defaultPath: current ? dirname(current.path) : app.getPath('documents') });
    return r.canceled ? null : r.filePaths[0] ?? null;
  });
  handle('batch:scan', (args) => batchScan(args && typeof args === 'object' ? args : {}));
  handle('batch:start', (args) => batchStart(args && typeof args === 'object' ? args : {}));
  handle('batch:cancel', () => { if (batchRun) batchRun.cancelled = true; return null; });
  handle('batch:saveReport', () => batchSaveReport());
  handle('pdftools:pick', () => pdfPick());
  handle('doc:print', () => printCurrent());
  handle('pdftools:run', (args) => pdfRun(args && typeof args === 'object' ? args : {}));
  handle('doc:openInApp', () => openInProgram());
  handle('doc:openInWord', () => openPdfInWord());
  handle('panel:full', () => { showHub(); return null; });
  handle('panel:close', () => {
    // 옆 창 닫기: Office Kit를 직접 실행해 쓰던 주 창이면 보통 주 창으로 돌아온다.
    // 탐색기에서 문서를 열어 시작했으면 숨긴다 (편집 창은 그대로). 편집 창이 하나도 없으면 끝낸다
    if (!win) return null;
    if (hubPrimary) { showHub(); return null; }
    win.hide();
    leavePanel();
    if (!hubPrimary && !(editors?.list().length)) app.quit();
    return null;
  });
  handle('doc:edit', () => editDocument());
  handle('doc:savePdf', (args) => savePdf(args && typeof args === 'object' ? { step: ['pick', 'make'].includes(args.step) ? args.step : undefined } : {}));
  handle('doc:showInFolder', () => { shell.showItemInFolder(needDoc().path); return null; });
  handle('shell:showItem', (p) => { shell.showItemInFolder(existingFile(p)); return null; });
  handle('shell:openBackups', async () => { mkdirSync(backups.dir, { recursive: true }); await shell.openPath(backups.dir); return null; });
  handle('shell:openExternal', async (url) => {
    const r = await openWebLink(url);
    if (r === 'bad') throw fail('웹 주소만 열 수 있습니다');
    if (r === 'private') throw fail('이 PC나 내부망 주소는 열지 않아요');
    return null;
  });

  handle('recent:list', () => settings.recent());
  handle('recent:remove', (p) => { if (typeof p !== 'string') throw fail('잘못된 경로입니다'); settings.removeRecent(p); return settings.recent(); });
  handle('recent:clear', () => { settings.clearRecent(); return settings.recent(); });
  handle('recent:pin', (p, on) => {
    if (typeof p !== 'string' || !isAbsolute(p) || p.length > 4096) throw fail('잘못된 경로입니다');
    try { settings.pinRecent(resolve(p), on === true); } catch (err) { throw fail(err.message); }
    return settings.recent();
  });

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
  handle('assoc:status', () => assocStatus());
  handle('assoc:repair', () => assocRepair());
  handle('assoc:openSettings', () => assocOpenSettings());
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
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    async function capture(w) {
      let img;
      for (let i = 0; i < 8 && !img; i++) {
        try { img = await w.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true }); } catch (err) {
          if (i === 7) throw err;
          await sleep(700);
        }
      }
      return img;
    }

    /** 화면의 'PDF로 저장' 단추를 눌러(저장 대화상자 대신 SMOKE_PDF) PDF가 생기고 알림이 뜨는지 → { ok, toast, file, bytes, ms } */
    async function smokePdf() {
      const ask = (code) => win.webContents.executeJavaScript(code).catch(() => null);
      const started = Date.now();
      try { rmSync(SMOKE_PDF, { force: true }); } catch {}
      await ask("document.getElementById('btn-pdf')?.click()");
      let toast = null;
      for (const end = Date.now() + 180_000; Date.now() < end && !toast;) {
        toast = await ask("[...document.querySelectorAll('.toast .toast-msg')].map((t) => t.textContent).find((t) => /PDF/.test(t)) ?? null");
        if (!toast) await sleep(300);
      }
      const bytes = existsSync(SMOKE_PDF) ? statSync(SMOKE_PDF).size : 0;
      return { ok: bytes > 0 && /PDF로 저장했어요/.test(toast ?? ''), toast, file: SMOKE_PDF, bytes, ms: Date.now() - started };
    }

    /** 열린 문서를 편집 창으로 열고, 다 뜨면(또는 실패 카드가 뜨면) 그 창을 <png>-editor.png로 찍는다 */
    async function smokeEditor() {
      const png = `${SMOKE.replace(/\.png$/i, '')}-editor.png`;
      if (!editors) return { ok: false, error: `편집 창 기능 없음: ${editorLoadError}` };
      if (!current) return { ok: false, error: '열린 문서가 없습니다' };
      const before = new Set(BrowserWindow.getAllWindows());
      // 묶음이 없어도 창을 연다 (편집 창의 실패 안내를 찍으려고)
      try { await editors.openEditor(current.path); } catch (err) { return { ok: false, error: err.message }; }
      pushEditorState();
      let ew = null;
      for (let i = 0; i < 100 && !ew; i++) {
        ew = BrowserWindow.getAllWindows().find((w) => !before.has(w) && !w.isDestroyed()) ?? null;
        if (!ew) await sleep(100);
      }
      if (!ew) return { ok: false, error: '편집 창이 뜨지 않았습니다' };
      const kind = editorKind(current.path);
      const ask = (code) => ew.webContents.executeJavaScript(code).catch(() => null);
      let state = null;
      for (const end = Date.now() + 90_000; Date.now() < end;) {
        state = await ask("document.getElementById('doc-state')?.dataset.state ?? null");
        if (state === 'saved' || state === 'error') break;
        await sleep(250);
      }
      const reason = state === 'error' ? await ask("document.getElementById('failure-reason')?.textContent ?? ''") : '';
      // Office: host가 '저장됨'이 된 뒤에도 틀 안 편집기가 문서를 그리는 중일 수 있다 → filex 앱 단계(data-fx-phase)가 ready가 될 때까지
      let phase = null; let fonts = null; let about = null;
      if (kind === 'office') {
        about = await ask("document.getElementById('btn-about')?.textContent?.trim() ?? null");
        const frameOf = (re) => { try { return ew.webContents.mainFrame.framesInSubtree.find((f) => re.test(f.url)) ?? null; } catch { return null; } };
        const inFrame = async (re, code) => { const f = frameOf(re); try { return f ? await f.executeJavaScript(code, true) : null; } catch { return null; } };
        for (const end = Date.now() + 90_000; state === 'saved' && Date.now() < end;) {
          phase = await inFrame(/\/office\/[^/]+\/index\.html/, 'document.documentElement.dataset.fxPhase || ""');
          if (phase === 'ready' || phase === 'failed') break;
          await sleep(300);
        }
        if (phase === 'ready') {
          await sleep(3000); // 글꼴을 받아 쪽을 다 그리게
          // 편집기 틀이 새 화면을 실제로 그릴 때까지 (가려지거나 바쁜 창은 예전 화면이 찍힌다)
          try { ew.moveTop(); } catch {}
          await Promise.race([
            inFrame(/\/web-apps\/apps\/\w+\/main\/index\.html/, 'new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(true))))'),
            sleep(8000),
          ]);
          try { ew.webContents.invalidate(); } catch {}
          await sleep(400);
          // 글꼴 목록 (이 PC 글꼴로 만든 AllFonts.js인지 — 편집기가 읽은 뒤 AscFonts.g_font_infos에 둔다)
          fonts = await inFrame(/\/web-apps\/apps\/\w+\/main\/index\.html/,
            '(() => { const i = window.AscFonts?.g_font_infos; if (!Array.isArray(i)) return null; const n = i.map((x) => String(x?.Name ?? "")); return { count: n.length, malgun: n.some((x) => /^(맑은 고딕|Malgun Gothic)$/.test(x)) }; })()');
        }
      } else {
        await sleep(1500); // 쪽이 다 그려지게
      }
      const img = await capture(ew);
      mkdirSync(dirname(png), { recursive: true });
      writeFileSync(png, img.toPNG());
      // Office: '정보 · 라이선스' 대화상자도 찍는다 (<png>-editor-about.png = 정보, -about-texts.png = 수정 고지 원문)
      let aboutInfo = null;
      if (kind === 'office' && phase === 'ready') {
        const shot = async (suffix) => { const i = await capture(ew); const f = png.replace(/\.png$/i, `-${suffix}.png`); writeFileSync(f, i.toPNG()); return f; };
        await ask("document.getElementById('btn-about')?.click()");
        let titles = null;
        for (const end = Date.now() + 8000; Date.now() < end && !titles?.length;) {
          titles = await ask("document.getElementById('about')?.open ? [...document.querySelectorAll('#about-info .about-sec h3')].map((h) => h.textContent) : null");
          if (!titles?.length) await sleep(200);
        }
        await sleep(500);
        const infoPng = await shot('about');
        // 글 안 '수정 고지' 링크 → 라이선스 원문 탭에서 그 파일을 글로 보여 주는지
        await ask("document.querySelector('#about-info .inline-link')?.click()");
        await sleep(2500);
        const textShown = await ask("(() => { const v = document.getElementById('about-viewer'); return { url: v?.dataset.url ?? null, note: document.getElementById('about-viewer-note')?.hidden === false ? document.getElementById('about-viewer-note').textContent : null }; })()");
        const textsPng = await shot('about-texts');
        await ask("document.getElementById('about')?.close()");
        aboutInfo = { titles, legal: Array.isArray(titles) && titles.includes('수정된 판일 수 있습니다'), text: textShown, png: infoPng, textsPng };
      }
      // 주 창을 앞으로 (상태 막대의 '편집 창 열림'이 보이게)
      if (win && !win.isDestroyed()) { win.moveTop(); win.focus(); }
      const ok = state === 'saved' && (kind !== 'office' || phase === 'ready');
      return { ok, kind, state, phase, fonts, about, aboutInfo, reason: reason || null, title: ew.getTitle(), png, size: img.getSize() };
    }

    ipcMain.on('smoke:ready', async (e, info = {}) => {
      if (!win || e.sender !== win.webContents) return;
      try {
        // 열면 바로 편집 창 (OFFICEKIT_SMOKE_AUTOEDIT=1): 화면이 편집 창을 띄웠는지 (30초까지 기다림)
        let autoEditor = null;
        if (process.env.OFFICEKIT_SMOKE_AUTOEDIT === '1' && current) {
          for (const end = Date.now() + 30_000; Date.now() < end && !editors?.isOpen(current.path);) await sleep(200);
          autoEditor = { open: !!editors?.isOpen(current.path), windows: BrowserWindow.getAllWindows().length };
        }
        const editor = SMOKE_EDITOR && info.ok !== false ? await smokeEditor() : null;
        const pdf = SMOKE_PDF && info.ok !== false ? await smokePdf() : null;
        await sleep(info.kind === 'pdf' ? 3000 : 900);
        const img = await capture(win);
        mkdirSync(dirname(SMOKE), { recursive: true });
        writeFileSync(SMOKE, img.toPNG());
        console.log('[smoke]', JSON.stringify({ ...info, ...(autoEditor ? { autoEditor } : {}), ...(editor ? { editor } : {}), ...(pdf ? { pdf } : {}), png: SMOKE, size: img.getSize() }));
        clearTimeout(bail);
        app.exit(info.ok === false || editor?.ok === false || pdf?.ok === false ? 1 : 0);
      } catch (err) {
        console.error('[smoke] 캡처 실패:', err);
        app.exit(1);
      }
    });
  }

  // ── 웹 링크 ──────────────────────────────────────────
  // 문서(PDF 미리보기)·AI 답의 링크는 문서가 정한 주소일 수 있다 → 주소를 보여 주고 묻는다. 이 PC·내부망 주소는 열지 않는다
  // (주소에 문서 내용을 실어 보내거나 내부망 장비를 건드리지 않게). Office Kit 안내 링크(아래 목록)는 묻지 않는다
  const TRUSTED_LINK_HOSTS = new Set(['github.com', 'docs.anthropic.com', 'claude.ai', 'claude.com', 'code.claude.com', 'openai.com', 'developer.hancom.com', 'www.onlyoffice.com', 'onlyoffice.com']);
  let askingLink = false;
  /** → 'opened' | 'canceled' | 'bad' | 'private' */
  async function openWebLink(raw) {
    let u;
    try { u = new URL(String(raw)); } catch { return 'bad'; }
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return 'bad';
    const { isPrivateHost } = await import('./editor-window.mjs');
    if (isPrivateHost(u.hostname)) {
      send('toast', { type: 'warn', message: `이 PC나 내부망 주소는 열지 않아요: ${u.host}` });
      return 'private';
    }
    if (!TRUSTED_LINK_HOSTS.has(u.hostname.toLowerCase())) {
      if (askingLink) return 'canceled';
      askingLink = true;
      try {
        const shown = u.href.length > 300 ? `${u.href.slice(0, 300)}…` : u.href;
        const { response } = await dialog.showMessageBox(win, {
          type: 'question', title: 'Office Kit', message: '이 주소를 기본 브라우저로 열까요?', detail: shown,
          buttons: ['열기', '취소'], defaultId: 1, cancelId: 1, noLink: true,
        });
        if (response !== 0) return 'canceled';
      } finally { askingLink = false; }
    }
    await shell.openExternal(u.href);
    return 'opened';
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

  function createWindow({ hidden = false } = {}) {
    const saved = SMOKE ? null : settings.get('window');
    win = new BrowserWindow({
      ...initialBounds(), minWidth: 1000, minHeight: 680, title: 'Office Kit', show: false,
      backgroundColor: nativeTheme.shouldUseDarkColors ? '#16181d' : '#f4f5f7',
      autoHideMenuBar: true,
      icon: APP_ICON,
      webPreferences: {
        preload: join(here, 'preload.cjs'),
        contextIsolation: true, nodeIntegration: false, sandbox: true,
        plugins: true,            // 내장 PDF 뷰어
        webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
        backgroundThrottling: !SMOKE,
        devTools: !app.isPackaged,
      },
    });
    // 숨긴 채 시작하면(문서 창 방식) 최대화는 처음 보일 때 한다 — maximize()는 숨긴 창도 보이게 한다
    if (saved?.maximized && !hidden) win.maximize();
    else if (saved?.maximized) hubMaximizeLater = true;
    win.once('ready-to-show', () => { if (!hidden) win.show(); if (SMOKE) win.focus(); });

    const wc = win.webContents;
    // 새 창·이동 막기: 웹 주소는 기본 브라우저로
    wc.setWindowOpenHandler(({ url }) => {
      if (/^https?:\/\//i.test(url)) void openWebLink(url);
      return { action: 'deny' };
    });
    wc.on('will-navigate', (e) => e.preventDefault());
    wc.on('will-frame-navigate', (e) => {
      if (e.isMainFrame) { e.preventDefault(); return; }
      const u = e.url;
      // 미리보기 PDF와 내장 PDF 뷰어만 허용
      if (u === 'about:blank' || u.startsWith('chrome-extension://') || isPreviewUrl(u)) return;
      e.preventDefault();
      if (/^https?:\/\//i.test(u)) void openWebLink(u);
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

    // 주 창을 닫으면 편집 창도 닫는다. 저장 안 한 편집 창은 저장 / 저장 안 함 / 취소를 묻고, 취소하면 주 창도 그대로 둔다
    let editorsClosed = false; let closingEditors = null;
    win.on('close', (e) => {
      // 문서 창 방식(탐색기에서 열어 시작)에서는 주 창·옆 창을 닫아도 편집 창은 그대로 둔다 → 숨기기만
      if (!hubPrimary && !appQuitting && editors?.list().length) {
        e.preventDefault();
        win.hide();
        leavePanel();
        return;
      }
      if (panelOf) leavePanel();
      if (!editorsClosed && editors?.list().length) {
        e.preventDefault();
        closingEditors ??= editors.closeAll()
          .then((ok) => { if (ok && win && !win.isDestroyed()) { editorsClosed = true; win.close(); } })
          .catch((err) => console.error('[editor] 편집 창을 닫지 못했습니다:', err))
          .finally(() => { closingEditors = null; });
        return;
      }
      if (SMOKE || !win) return;
      try { settings.set({ window: { ...win.getNormalBounds(), maximized: win.isMaximized() } }); } catch {}
    });
    win.on('closed', () => { win = null; });

    win.loadFile(INDEX_HTML);
  }

  function buildMenu() {
    const template = [
      { label: '파일', submenu: [
        { label: '새 문서…', accelerator: 'CmdOrCtrl+N', click: () => send('menu', 'new') },
        { label: '파일 열기…', accelerator: 'CmdOrCtrl+O', click: () => send('menu', 'open') },
        { label: '직접 편집', accelerator: 'CmdOrCtrl+E', click: () => send('menu', 'edit') },
        { label: '새로고침', accelerator: 'F5', click: () => send('menu', 'refresh') },
        { label: '인쇄…', accelerator: 'CmdOrCtrl+P', click: () => send('menu', 'print') },
        { type: 'separator' },
        { label: '개인정보 찾아 가리기…', click: () => send('menu', 'pii') },
        { label: '한글 → HWPX 일괄 변환…', click: () => send('menu', 'batch') },
        { label: 'PDF 합치기 · 나누기 · 돌리기…', click: () => send('menu', 'pdftools') },
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
  // 탐색기에서 편집할 수 있는 문서로 시작 → 문서 창 방식 (주 창은 숨김). 확인용 캡처는 OFFICEKIT_SMOKE_DOCWINDOW=1일 때만
  // OFFICEKIT_DOC_WINDOW=0: 끈다 (시험용)
  const docWindowStart = !!pendingFile && process.env.OFFICEKIT_DOC_WINDOW !== '0' && (!SMOKE || process.env.OFFICEKIT_SMOKE_DOCWINDOW === '1')
    && (() => { try { return wantsDocWindow(pendingFile) && statSync(pendingFile).isFile(); } catch { return false; } })();
  const launchDoc = docWindowStart ? pendingFile : null;
  if (docWindowStart) { hubPrimary = false; pendingFile = null; }

  app.on('before-quit', () => { appQuitting = true; });

  app.on('second-instance', (_e, argv, cwd) => {
    const f = fileFromArgv(argv, cwd);
    // 탐색기에서 연 문서 → 그 문서의 편집 창만 (주 창은 건드리지 않는다)
    if (f && wantsDocWindow(f) && !SMOKE && process.env.OFFICEKIT_DOC_WINDOW !== '0') {
      openAsDocument(f).catch((err) => { showHub(); send('toast', { type: 'error', message: err.message }); });
      return;
    }
    if (win) { if (win.isMinimized()) win.restore(); leavePanel(); win.show(); win.focus(); }
    if (!f) return;
    if (rendererReady) openDocument(f, { autoEdit: true }).catch((err) => send('toast', { type: 'error', message: err.message }));
    else pendingFile = f;
  });

  app.on('window-all-closed', () => app.quit());
  // 편집 창이 열려 있으면 '끝내기'가 취소될 수 있으므로(저장 확인 → 취소) 정리는 정말 끝날 때(will-quit) 한다
  app.on('will-quit', () => {
    stopWatch();
    for (const c of chats.values()) { try { c.handle.cancel(); } catch {} }
    // 창이 다 닫힌 뒤 미리보기 PDF(문서 사본)를 지운다. 못 지운 것은 다음 시작 때 지운다
    cleanPreviews(null);
    cleanPrintCopies({ all: true });
  });
  // 편집 창이 닫히면 상태 막대를 맞춘다
  app.on('browser-window-created', (_e, w) => { w.once('closed', () => pushEditorState()); });

  await app.whenReady();
  nativeTheme.themeSource = ['light', 'dark'].includes(settings.get('theme')) ? settings.get('theme') : 'system';
  // 맞춤법 검사는 쓰지 않는다. 창을 만들기만 해도 Chromium이 인터넷에서 사전(ko-3-0.bdic, 11MB)을 받으므로 받을 곳을 막는다
  try { session.defaultSession.setSpellCheckerDictionaryDownloadURL('http://127.0.0.1:9/'); } catch {}

  // 권한 요청은 PDF 뷰어에 필요한 것만
  const allowed = new Set(['fullscreen', 'clipboard-sanitized-write']);
  session.defaultSession.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  session.defaultSession.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));

  cleanPreviews(null);
  cleanPrintCopies();
  // 오래 안 연 문서의 자동 저장 사본(Office)·쓰다 남은 임시 파일 정리
  try { cleanOfficeDrafts(); } catch (err) { console.error('[editor] 자동 저장 사본 정리 실패:', err?.message ?? err); }
  buildMenu();
  createWindow({ hidden: docWindowStart });
  if (launchDoc) openAsDocument(launchDoc).catch((err) => { showHub(); send('toast', { type: 'error', message: err.message }); });
}
