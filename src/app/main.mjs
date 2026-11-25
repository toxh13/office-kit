// main.mjs — Office Kit 창, 단일 인스턴스, IPC, 원본 백업, 파일 감시, AI 대화 중계.
import { app, BaseWindow, BrowserWindow, Menu, WebContentsView, dialog as electronDialog, ipcMain, nativeTheme, screen, session, shell } from 'electron';
import { randomUUID } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, watch, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { allows, channelSpec, VIEW_ROUTES } from './ipc-channels.mjs';
import { createI18n, LANG_SETTINGS, localizeDialogOptions, localizedDialog, localizeMenu, resolveLang } from './i18n/index.mjs';
import { createBackups, createPreviewCache, createPreviewFiles, createSettings, hashFile, samePath } from './store.mjs';
import { createTabWindow } from './tab-window.mjs';
import { createTabList, docKey, tabCommandOf } from './tabs.mjs';

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
// 화면 언어 (설정 'lang' → 시작 때 정한다. 바꾸면 다시 시작). 알림 창은 늘 이 언어로 번역해 보인다
let i18n = createI18n('ko');
const dialog = localizedDialog(electronDialog, () => i18n.t);

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
  // AI 대화마다 MCP 서버가 띄운 Office의 PID를 적는 폴더 (userData/com-pids/<chatId>)
  const COM_PIDS = 'com-pids';
  // 프로그램으로 만든 미리보기 PDF는 파일이 그대로면 다시 쓴다 (Excel·Word 변환은 5~25초)
  const previewCache = createPreviewCache(join(userData, 'preview-cache'));

  // 허브(Hub): 홈 종류 뷰(홈 탭·문서 탭)마다 따로 두는 문맥 (판정 2·3: 홈 탭은 문서를 품지 않는다 → 홈 허브의 current는 늘 null)
  //   { wc, tw: 그 탭 창, tabId: 문서 탭 id (홈이면 'home'), current: 열린 문서 { path, backup, stat, renderedStat, quietUntil } | null,
  //     watcher·watchTimer: 열린 문서의 파일 감시, pendingFile: 화면이 준비되면 열 파일(첫 탭 창의 홈만), rendererReady }
  const hubs = new Map();   // webContents → Hub
  function makeHub(wc, tw, tabId) {
    const hub = { wc, tw, tabId, current: null, watcher: null, watchTimer: null, pendingFile: null, rendererReady: false };
    hubs.set(wc, hub);
    // 뷰가 닫히면 허브도 정리한다 (감시 끄기)
    wc.once('destroyed', () => { stopWatch(hub); hubs.delete(wc); });
    return hub;
  }
  /** 보낸 쪽 webContents → 그 허브 (허브가 아닌 화면이면 null) */
  const hubOf = (wc) => (wc && hubs.get(wc)) || null;
  const allHubs = () => [...hubs.values()];
  // 탭 창: shell.id → tw = { shell, tabs: TabList, home: View, panel: View, docViews: Map<tabId, View>, edViews: Map<tabId, View>, titles: Map<tabId, string> }
  const tabWindows = new Map();
  const windowOrder = [];   // 최근에 초점 받은 탭 창 차례 (shell.id, 앞이 최근)
  // 화면 뷰(webContents) → { kinds: IPC 뷰 종류, tw, role: 'home' | 'doc' | 'panel' | 'tabstrip' }
  const viewRecs = new Map();
  let appQuitting = false;

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

  function needDoc(hub) {
    if (!hub?.current) throw fail('열린 문서가 없습니다');
    if (!existsSync(hub.current.path)) throw fail(`문서를 찾을 수 없습니다: ${basename(hub.current.path)}`, 'NOT_FOUND');
    return hub.current;
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

  /** 살아 있는 화면(wc)에만 보낸다. 채널은 IPC 표로 먼저 검사한다 */
  function sendTo(wc, channel, data) {
    channelSpec(channel, 'send');
    if (wc && !wc.isDestroyed()) wc.send(channel, data);
  }

  // ── 문서 ─────────────────────────────────────────────

  function statOf(p) {
    try { const st = statSync(p); return { mtimeMs: st.mtimeMs, size: st.size }; } catch { return null; }
  }

  function docInfo(hub) {
    if (!hub?.current) return null;
    const st = statOf(hub.current.path);
    return {
      path: hub.current.path, name: basename(hub.current.path), dir: dirname(hub.current.path),
      ext: extOf(hub.current.path), kind: kindOf(hub.current.path),
      size: st?.size ?? 0, mtime: st?.mtimeMs ?? null, missing: !st,
      backup: hub.current.backup ? { time: hub.current.backup.time, file: hub.current.backup.file } : null,
      backupError: hub.current.backupError ?? null,
      hasSession: sessions.has(key(hub.current.path)),
      editor: editorState(hub.current.path),
      aiBusy: aiBusyFor(hub.current.path),
    };
  }

  /** 그 경로로 진행 중인 AI 대화가 있는가 (대화 칸 키 = path) */
  const aiBusyFor = (p) => [...chats.values()].some((c) => samePath(c.path, p));

  /** AI 패널에 보일 대상 문서 정보 (ai:target) */
  function targetInfo(p) {
    if (!p) return null;
    return {
      path: p, name: basename(p), dir: dirname(p), ext: extOf(p), kind: kindOf(p),
      editor: editorState(p), hasSession: sessions.has(key(p)), aiBusy: aiBusyFor(p),
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

  /** 열 때 편집 탭으로 열지 (설정 '문서를 열면': 편집 / 미리보기). 편집 탭으로 열 수 있는 형식만 */
  // 확인용 캡처는 OFFICEKIT_SMOKE_AUTOEDIT=1일 때만 (다른 캡처는 주 창만 찍는다)
  const wantsEditor = (p) => (!SMOKE || process.env.OFFICEKIT_SMOKE_AUTOEDIT === '1') && settings.get('openAction') !== 'preview' && !!editorKind(p);

  /**
   * 문서 탭 뷰(허브)에 문서 열기: 처음 연 상태 버전·최근 목록·감시. 편집 탭으로 열지는 openPath가 정한다 (autoEdit 왕복 없음)
   */
  function openDocument(hub, p) {
    p = docPath(p);
    // 이미 열린 문서를 다시 열면 백업 기준과 감시는 그대로 둔다
    if (hub.current && samePath(hub.current.path, p)) {
      if (!hub.watcher) startWatch(hub, hub.current.path);
      if (!SMOKE) settings.addRecent(hub.current.path);
      const doc = docInfo(hub);
      sendTo(hub.wc, 'doc:opened', doc);
      return doc;
    }
    stopWatch(hub);
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
    hub.current = { path: p, backup, backupError, stat: statOf(p), missing: false, renderedStat: null, quietUntil: 0 };
    // 확인용 캡처(SMOKE)는 최근 파일 목록과 Windows 최근 항목을 건드리지 않는다
    // (Windows 최근 항목에 넣으면 탐색기가 그 형식의 '연결 프로그램' 목록에도 Office Kit를 적어 둔다)
    if (!SMOKE) {
      settings.addRecent(p);
      try { app.addRecentDocument(p); } catch {}
    }
    startWatch(hub, p);
    const doc = docInfo(hub);
    sendTo(hub.wc, 'doc:opened', doc);
    return doc;
  }

  function closeDocument(hub) {
    stopWatch(hub);
    hub.current = null;
    sendTo(hub.wc, 'doc:opened', null);
  }

  // 파일 감시: 편집기가 새 파일로 바꿔치기하는 경우도 잡도록 폴더를 지켜본다 (허브마다 hub.watcher·hub.watchTimer)
  function startWatch(hub, p) {
    const name = basename(p).toLowerCase();
    const check = () => {
      if (!hub.current || hub.current.path !== p) return;
      if (Date.now() < hub.current.quietUntil) return;
      const st = statOf(p);
      if (!st) {
        // 없어진 순간에만 알리고, 크기·시각을 잊어 두어 같은 파일이 돌아와도 알린다
        if (!hub.current.missing) {
          hub.current.missing = true; hub.current.stat = null;
          sendTo(hub.wc, 'doc:changed', { path: p, missing: true });
          editors?.notifyExternalChange(p);
        }
        return;
      }
      const back = hub.current.missing;
      hub.current.missing = false;
      const last = hub.current.stat;
      if (!back && last && last.mtimeMs === st.mtimeMs && last.size === st.size) return;
      hub.current.stat = st;
      sendTo(hub.wc, 'doc:changed', { path: p, missing: false, mtime: st.mtimeMs, size: st.size });
      // 편집 창이 열려 있으면 알린다 (고친 것 없으면 다시 불러오고, 있으면 편집 창이 물어본다. 내용이 같으면 무시)
      editors?.notifyExternalChange(p);
    };
    try {
      hub.watcher = watch(dirname(p), { persistent: false }, (_ev, file) => {
        if (file && String(file).toLowerCase() !== name) return;
        clearTimeout(hub.watchTimer);
        hub.watchTimer = setTimeout(check, 600);
      });
      hub.watcher.on('error', () => {});
    } catch { hub.watcher = null; }
  }
  function stopWatch(hub) {
    clearTimeout(hub.watchTimer);
    try { hub.watcher?.close(); } catch {}
    hub.watcher = null;
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

  // 미리보기 폴더 정리 (문서 하나의 PDF만 남긴다)
  const previewFiles = createPreviewFiles(previewDir);
  const cleanPreviews = (keep) => previewFiles.clean(keep);

  const clip = (rows, maxRows = 500, maxCols = 60) => (rows ?? []).slice(0, maxRows).map((r) => (r ?? []).slice(0, maxCols).map((v) => String(v ?? '')));

  async function preview(hub, { mode, ifChanged } = {}) {
    const doc = needDoc(hub);
    const st = statOf(doc.path);
    if (ifChanged && doc.renderedStat && st && doc.renderedStat.mtimeMs === st.mtimeMs && doc.renderedStat.size === st.size) {
      return { unchanged: true, doc: docInfo(hub) };
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
    if (hub.current === doc) { doc.renderedStat = st; doc.stat = st; }
    const base = { mode: m, ms: Date.now() - started, doc: hub.current === doc ? docInfo(hub) : null };
    if (res.kind === 'pdf') {
      const pdf = previewFiles.settle(res.file, doc.path, { current: hub.current === doc });
      // 늦게 끝난 미리보기: 지금 문서의 PDF를 지우지 않고, 화면도 이 결과를 버린다
      if (!pdf) return { ...base, kind: 'stale' };
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
    // 대화(MCP 서버)를 끝내도 그 서버가 띄운 Office는 남아 문서를 쥐고 있다 → 되돌리기 전에 끝낸다
    // 시간 안에 안 끝났으면 MCP 서버가 아직 그 Office로 일하는 중일 수 있다: 대화의 마무리(finally)에 맡긴다
    if (ended) for (const c of running) killChatOffice(c.pidDir);
    return ended ? running.length : -1;
  }

  /** 대화의 MCP 서버가 띄우고 못 끝낸 Office를 끝낸다 (remove: 대화가 끝났으면 PID 폴더도 지운다) */
  function killChatOffice(dir, remove = false) {
    if (!dir) return;
    try { com.killRegisteredOffice?.(dir); } catch (err) { console.error('[ai] 남은 Office 정리 실패:', err?.message ?? err); }
    if (remove) { try { rmSync(dir, { recursive: true, force: true }); } catch { /* 지우기 실패 */ } }
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
  function versions(doc) {
    let now = null;
    try { now = currentHash(doc.path); } catch {}
    const list = backups.list(doc.path).map((e) => ({ file: e.file, time: e.time, usedAt: e.usedAt ?? e.time, reason: e.reason ?? null, size: e.size ?? 0, current: !!now && e.hash === now }));
    return { versions: list, undo: list.find((e) => !e.current)?.file ?? null };
  }

  /**
   * 버전 요청의 문서 (판정 5): args.path가 보낸 뷰의 문서이거나 그 탭 창의 AI 대상일 때만 → { doc, hub }
   * doc은 그 뷰 허브의 current(있으면) 또는 { path }. 아니면 '허용되지 않은 요청' (code NOT_OPEN_TAB)
   */
  function docFor(ctx, args) {
    const p = args && typeof args === 'object' && typeof args.path === 'string' ? args.path : null;
    const bad = () => fail('허용되지 않은 요청입니다', 'NOT_OPEN_TAB');
    if (!p) throw bad();
    const hub = ctx.hub;
    if (hub?.current && samePath(hub.current.path, p)) return { doc: needDoc(hub), hub };
    const target = targetOf(ctx.tw);
    if (!target || !samePath(target, p)) throw bad();
    if (!existsSync(target)) throw fail(`문서를 찾을 수 없습니다: ${basename(target)}`, 'NOT_FOUND');
    // AI 대상이 문서 탭이면 그 허브의 문서 기록을 쓴다 (감시 쉬기·상태가 같이 맞게)
    const owner = allHubs().find((h) => h.tw === ctx.tw && h.current && samePath(h.current.path, target)) ?? null;
    return owner ? { doc: owner.current, hub: owner } : { doc: { path: target }, hub: null };
  }

  /** 버전으로 되돌리기. file을 주지 않으면 마지막 변경 전으로 (되돌리기 전 상태도 '원래대로 전' 버전으로 남는다) */
  async function restore(doc, args, hub = null) {
    const k = key(doc.path);
    if (restoring.has(k)) throw fail('이미 되돌리는 중입니다');
    const want = args && typeof args === 'object' && typeof args.file === 'string' ? args.file : null;
    let entry;
    if (want) {
      entry = backups.find(doc.path, want);
      if (!entry) throw fail('그 버전을 찾을 수 없어요 (정리됐을 수 있어요). 목록을 다시 열어 주세요.', 'NO_VERSION');
    } else {
      const v = versions(doc);
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
      // 남기지 못하면 되돌리지 않는다 (되돌리면 지금 내용을 영영 잃을 수 있다). 파일이 이미 없으면 남길 내용이 없으니 그대로 진행
      if (existsSync(doc.path)) {
        try { backups.backup(doc.path, new Date(), { protect: [entry.file, ...protectedBackups()], maxPerSource: VERSIONS_PER_DOC, reason: 'restore' }); } catch (err) {
          // 문서를 다른 프로그램이 쥐고 있어 읽지 못한 것이면 디스크 공간 안내는 틀리다
          if (/EBUSY|EPERM|EACCES/.test(err?.code ?? '')) throw fail('문서가 다른 프로그램에서 열려 있어 되돌릴 수 없습니다. 프로그램을 닫고 다시 시도하세요.');
          throw fail('지금 내용을 버전으로 남기지 못해 되돌리지 않았습니다. 디스크 공간을 확인하세요.', 'BACKUP_FAILED');
        }
      }
      // 그 문서를 보는 문서 탭들의 감시는 잠깐 쉰다 (되돌린 것을 바깥 변경으로 알리지 않게)
      const quiet = Date.now() + 1500;
      doc.quietUntil = quiet;
      for (const h of allHubs()) if (h.current && samePath(h.current.path, doc.path)) h.current.quietUntil = quiet;
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
      for (const tw of tabWindows.values()) pushTarget(tw);
      return { time: entry.time, reason: entry.reason ?? null, file: entry.file, stoppedAi, newChat, editorNotified, doc: hub && hub.current === doc ? docInfo(hub) : null };
    } finally {
      restoring.delete(k);
    }
  }

  /**
   * 버전(백업) 이후 바뀐 글 → { file, time, reason, kind: 'text'|'cells', hunks|cells, stats, same, … } (core/diff.mjs)
   * 지금 파일은 사본으로 읽는다 (다른 프로그램이 열고 있어도 읽히게, 읽는 사이 바뀌어도 한 시점으로)
   */
  let loadedDiff = null;
  async function diffVersion(doc, args) {
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

  async function piiScan(hub) {
    const doc = needDoc(hub);
    const m = await loadPii();
    try {
      const r = await m.scanDocument(doc.path);
      return { path: doc.path, items: r.items, total: r.total, chars: r.chars, types: m.PII_TYPES, output: m.maskedPathFor(doc.path) };
    } catch (err) { throw Object.assign(fail(`개인정보를 찾지 못했어요: ${err?.message ?? err}`, 'PII_FAILED'), { detail: err?.stack ?? '' }); }
  }

  /** 가린 사본 만들기 (원본은 그대로, 원본 옆 '<이름>_가림.<확장자>') → { output, items, remaining } */
  async function piiMask(hub, args) {
    const doc = needDoc(hub);
    const m = await loadPii();
    m.piiTargetCheck(args?.path, doc.path); // 찾을 때의 문서가 아니면 거부
    const values = Array.isArray(args?.values) ? args.values.filter((v) => typeof v === 'string' && v.length <= 200).slice(0, 2000) : null;
    if (!values?.length) throw fail('가릴 항목을 하나 이상 골라 주세요');
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

  async function batchPick(hub, { kind } = {}) {
    const folder = kind === 'folder';
    const r = await dialog.showOpenDialog(winOf(hub), folder
      ? { title: '한글 문서가 든 폴더 고르기', properties: ['openDirectory', 'multiSelections'], defaultPath: hub.current ? dirname(hub.current.path) : app.getPath('documents') }
      : { title: '한글 문서 고르기', properties: ['openFile', 'multiSelections'], defaultPath: hub.current ? dirname(hub.current.path) : app.getPath('documents'), filters: [{ name: '한글 문서 (.hwp)', extensions: ['hwp'] }] });
    return r.canceled ? [] : r.filePaths;
  }

  async function batchScan({ inputs, recursive } = {}) {
    const m = await loadBatch();
    const { items, truncated } = m.collectHwpFiles(pathList(inputs), { recursive: recursive !== false });
    return { count: items.length, truncated, files: items.slice(0, 200).map((it) => it.src) };
  }

  async function batchStart(wc, { inputs, recursive, overwrite, outDir } = {}) {
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
        onProgress: ({ index, total, result }) => sendTo(wc, 'batch:progress', { index, total, result }),
      });
      batchLast = r.results;
      return r;
    } finally { batchRun = null; }
  }

  async function batchSaveReport() {
    if (!batchLast?.length) throw fail('저장할 결과가 없어요');
    const m = await loadBatch();
    const dir = dirname(batchLast[0].out);
    const r = await dialog.showSaveDialog(winOf(null), { title: '변환 결과 보고서 저장', defaultPath: join(dir, 'HWPX 변환 결과.csv'), filters: [{ name: 'CSV (엑셀에서 열기)', extensions: ['csv'] }] });
    if (r.canceled || !r.filePath) return { canceled: true };
    const file = r.filePath.toLowerCase().endsWith('.csv') ? r.filePath : `${r.filePath}.csv`;
    m.writeBatchReport(file, batchLast);
    return { output: file };
  }

  // ── PDF 도구 (core/pdftools.mjs) ──────────────────────
  let loadedPdfTools = null;
  const loadPdfTools = async () => (loadedPdfTools ??= await import('../core/pdftools.mjs'));
  const PDF_ACTIONS = { merge: '합침', extract: '쪽뽑기', delete: '쪽지움', rotate: '돌림' };

  async function pdfPick(hub) {
    const r = await dialog.showOpenDialog(winOf(hub), { title: 'PDF 고르기', properties: ['openFile', 'multiSelections'],
      defaultPath: hub.current ? dirname(hub.current.path) : app.getPath('documents'), filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] });
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

  async function pdfRun(hub, { action, inputs, pages, degrees } = {}) {
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
      : await dialog.showSaveDialog(winOf(hub), { title: '결과 PDF 저장', defaultPath: suggest, filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] });
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
  async function openPdfInWord(hub) {
    const doc = needDoc(hub);
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

  async function openInProgram(hub) {
    const doc = needDoc(hub);
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
  async function savePdf(hub, { step } = {}) {
    const doc = needDoc(hub);
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
    const r = SMOKE_PDF ? { canceled: false, filePath: SMOKE_PDF } : await dialog.showSaveDialog(winOf(hub), {
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
  function newDocFolder(hub) {
    const dir = hub?.current ? dirname(hub.current.path) : null;
    return dir && isDir(dir) ? dir : app.getPath('documents');
  }

  function blankFormat(format) {
    const f = typeof format === 'string' ? format.toLowerCase() : '';
    if (!create.BLANK_FORMATS.includes(f)) throw fail('지원하지 않는 형식입니다', 'BAD_FORMAT');
    return f;
  }

  /** 새 문서 창 기본값 → { folder, name, documents, formats }. folder를 주면 그 폴더에서 겹치지 않는 이름 */
  function newDocInfo(hub, { format = create.BLANK_FORMATS[0], folder } = {}) {
    const ext = blankFormat(format);
    const dir = typeof folder === 'string' && folder.length <= 4096 && isAbsolute(folder) ? resolve(folder) : newDocFolder(hub);
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

  async function pickFolder(hub, start) {
    const r = await dialog.showOpenDialog(winOf(hub), {
      title: '새 문서를 저장할 폴더', properties: ['openDirectory'],
      defaultPath: typeof start === 'string' && isAbsolute(start) && isDir(start) ? start : newDocFolder(hub),
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
  async function newDocument(ctx, args) {
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
    const tw = ctx.tw && tabWindows.has(ctx.tw.shell.id) ? ctx.tw : lastTw() ?? createTabWindowFor({ bounds: initialBounds() });
    // 지워진 그 문서의 문서 탭이 남아 있으면 닫고 새로 연다
    const old = findTab(plan.path);
    if (old?.tab.type === 'doc') closeDocTab(old.tw, old.tab.id);
    // 편집 탭으로 열 수 있는 형식(한글·docx·xlsx·pptx)은 만든 뒤 바로 편집 탭으로 (못 열어도 문서는 만들었으므로 실패로 돌려주지 않고 문서 탭으로)
    let editor = null; let editorError = null;
    if (args.edit === true && editorKind(plan.path)) {
      try { editor = await editPath(tw, plan.path); } catch (err) {
        editorError = { message: err?.message ?? String(err), code: err?.code ?? null };
      }
    }
    if (!editor) {
      if (findTab(plan.path)) await openPath(tw, plan.path);
      else openDocTab(tw, plan.path);
    }
    let chat = null; let chatError = null;
    if (start === 'ai') {
      // 문서는 이미 만들었으므로 AI를 못 띄워도 실패로 돌려주지 않는다 (화면이 알린다).
      // 그 탭 창의 AI 패널에 'start'를 먼저 보내 패널이 대화를 만든다 (announce)
      try { chat = await sendChat({ tw, path: plan.path }, { prompt: newDocPrompt(plan.path, request), announce: request }); } catch (err) {
        if (!err?.expected && err?.code !== 'MODULE_MISSING') console.error('[doc:new] AI', err);
        chatError = { message: err?.message ?? String(err), code: err?.code ?? null };
      }
    }
    return { doc: targetInfo(plan.path), chat, chatError, editor, editorError };
  }

  // ── 편집 창 (한글 · 문서·표·슬라이드) ─────────────────
  // editor-window.mjs가 문서마다 창 하나를 띄운다 (인터넷 없이). 종류(kind)는 확장자로 고른다:
  //   'hwp' = 한글 hwp·hwpx·hwt → 내장 rhwp-studio, 'office' = docx·xlsx·pptx → vendor/office-editor
  // 모듈을 못 불러와도 앱은 뜨고 '직접 편집'만 이유를 알린다

  let EDITOR_KINDS = { hwp: ['hwp', 'hwpx', 'hwt'], office: ['docx', 'xlsx', 'pptx'] };
  // 창 이름 ('ONLYOFFICE'는 쓰지 않는다 — 설계 8장)
  const EDITOR_LABELS = { hwp: '한글 편집', docx: '문서 편집', xlsx: '표 편집', pptx: '슬라이드 편집' };
  let OFFICE_DRAFTS = 'office-drafts';
  /** 시작할 때 바꿔 연 임시 사본 정리 (editor-window.mjs) */
  let cleanConvertedCopies = null;
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
    if (typeof m.cleanConvertedCopies === 'function') cleanConvertedCopies = m.cleanConvertedCopies;
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
      app, WebContentsView, ipcMain, session, dialog, shell, nativeTheme,
      appRoot: APP_ROOT, userData, icon: APP_ICON,
      // 편집 뷰를 붙일 탭 창 (openEditorTab이 정한 탭 창, 아니면 최근 탭 창)
      hostFor: (info) => editorHostFor(info),
      onTabChange: (p, change) => editorTabChanged(p, change),
      t: (x) => i18n.t(x), lang: () => i18n.lang, i18nData: () => ({ lang: i18n.lang, dict: i18n.dict }),
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
      // 편집기의 'AI · 도구' 단추 → 그 탭을 보이고 같은 탭 창의 AI 패널을 펼친다
      onPanelRequest: (p) => showPanelFor(p),
      // 한글 편집 창의 인쇄·PDF로 저장: 편집 중인 hwp 바이트 → PDF (한컴+자동화 연결이 있으면 한컴, 없으면 내장 엔진)
      toPdf: (bytes) => hwpBytesToPdf(bytes),
      toPrint: (bytes, ew) => printHwpBytes(bytes, ew),
      // 편집 탭의 Ctrl+O (부모 = 탭 창)
      onOpenRequest: (parent) => { pickAndOpen(parent).catch((err) => console.error('[open]', err)); },
      // 탭 단축키(Ctrl+Tab·Ctrl+1…9·Ctrl+T·Ctrl+N)는 아직 쓰지 않는다 (탭 단축키 작업에서 연결)
      onTabCommand: (p, cmd) => { const hit = findTab(p); if (hit) runTabCommand(hit.tw, cmd); },
      // 끝내기 중 저장 확인에서 취소 → 끝내기 상태를 되돌린다
      onQuitCanceled: () => { appQuitting = false; },
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

  // 편집 탭이 열리고 닫히고 고쳐질 때 화면에 알린다: 문서가 열린 허브(문서 탭 상태 막대)마다 doc:editor, 탭 창마다 AI 패널에 ai:target.
  // 탭이 닫힌 뒤 정리·'다른 이름으로 저장' 뒤 경로 바뀜이 끝난 다음에 보도록 조금 늦춘다
  let editorPushTimer = 0;
  /** AI 대화가 시작·끝날 때: AI 패널(ai:target)·탭 줄(busy)·그 문서를 보는 문서 탭(doc:editor의 aiBusy)에 알린다 */
  function pushAiState() {
    for (const hub of allHubs()) {
      if (hub.current) sendTo(hub.wc, 'doc:editor', { path: hub.current.path, ...editorState(hub.current.path), aiBusy: aiBusyFor(hub.current.path) });
    }
    for (const tw of tabWindows.values()) { pushTarget(tw); sendTabs(tw); }
  }

  function pushEditorState() {
    clearTimeout(editorPushTimer);
    editorPushTimer = setTimeout(() => {
      for (const hub of allHubs()) {
        if (hub.current) sendTo(hub.wc, 'doc:editor', { path: hub.current.path, ...editorState(hub.current.path), aiBusy: aiBusyFor(hub.current.path) });
      }
      for (const tw of tabWindows.values()) pushTarget(tw);
    }, 40);
  }

  /** 편집 창이 저장함 → 바깥 변경과 같은 길로 미리보기를 새로 고치고 최근 목록에 올린다 (감시보다 먼저) */
  function editorSaved(p) {
    if (!SMOKE) {
      try { settings.addRecent(p); } catch {}
      try { app.addRecentDocument(p); } catch {}
    }
    // 그 문서가 열린 허브마다
    for (const hub of allHubs()) {
      if (!hub.current || !samePath(hub.current.path, p)) continue;
      const st = statOf(p);
      hub.current.stat = st;
      hub.current.missing = !st;
      if (st) sendTo(hub.wc, 'doc:changed', { path: hub.current.path, missing: false, mtime: st.mtimeMs, size: st.size, by: 'editor' });
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

  /** 편집 탭으로 열 수 있는 문서인지 미리 확인 → 종류 ('hwp' | 'office'). 아니면 한국어 이유(code)로 던진다 */
  function editorCheck(p) {
    const kind = editorKind(p);
    if (!kind && extOf(p) === 'pdf') {
      throw fail('PDF는 앱 안에서 바로 고칠 수 없어요. 글자를 고치려면 원본 문서(한글·Word 등)를 열어 고친 뒤 PDF로 다시 저장해 주세요. AI 대화로 내용을 묻거나 요약하는 것은 PDF에서도 돼요.', 'PDF_EDIT');
    }
    if (!kind) {
      const ext = extOf(p);
      throw fail(`${kindOf(p)} 문서(.${ext || '확장자 없음'})는 앱 안에서 직접 편집할 수 없어요. 직접 편집은 한글(HWP·HWPX)·Word(DOCX)·Excel(XLSX)·PowerPoint(PPTX) 문서만 돼요. ‘프로그램에서 열기’로 고쳐 주세요.`, 'UNSUPPORTED');
    }
    const label = editorLabel(p);
    // 한글: 내장 엔진이 못 여는 문서(손상·암호·지원 안 하는 판)는 탭을 띄우기 전에 거른다 (Office는 editor-window가 거름)
    if (kind === 'hwp' && !editors?.isOpen(p)) {
      const bad = hwpOpenProblem(p);
      if (bad) throw Object.assign(fail(bad.message, bad.code), { detail: bad.detail });
    }
    if (!editors) throw fail(`${label} 기능을 불러오지 못했습니다: ${editorLoadError}`, 'EDITOR_MISSING');
    const a = editors.available(kind);
    if (!a.available) throw fail(a.reason || `${label}기 파일이 없습니다.`, 'EDITOR_MISSING');
    return kind;
  }

  /** 편집 탭으로 열기 (확인 → openEditorTab) → { path, kind, reused, notice, tabId } */
  async function editPath(tw, p, opts = {}) {
    const kind = editorCheck(p);
    let r;
    try { r = await openEditorTab(tw, p, opts); } catch (err) {
      // 암호 문서·옛 형식·손상·너무 큼 등은 한국어 이유와 code(ENCRYPTED·OLD_FORMAT·BAD_FORMAT·TOO_BIG …)를 그대로 넘긴다
      if (err?.expected) throw err;
      console.error('[doc:edit]', err);
      throw fail(err?.message ?? String(err), err?.code ?? 'EDITOR_FAILED');
    }
    return { path: r.path, kind: r.kind ?? kind, reused: !!r.reused, notice: r.notice, tabId: r.tabId };
  }

  /** 문서 탭의 '직접 편집': 그 탭을 같은 자리 편집 탭으로 바꾼다 (판정 7) */
  async function editDocument(ctx) {
    const doc = needDoc(ctx.hub);
    return editPath(ctx.hub.tw ?? ctx.tw, doc.path, { replaceId: ctx.hub.tabId, noted: true });
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
  async function printCurrent(hub) {
    const doc = needDoc(hub);
    const ext = extOf(doc.path);
    if (core.isHwp(ext)) return printHwpPath(doc.path, winOf(hub));
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

  /** 그 탭 창에서 보이는 홈 종류 뷰(문서 탭이면 그 뷰, 아니면 홈 뷰)에 알림 */
  function toastIn(tw, message, type = 'error') {
    const t = tw?.tabs.active;
    const view = t?.type === 'doc' ? tw.docViews.get(t.id) : tw?.home;
    sendTo(view?.webContents, 'toast', { type, message });
  }

  /** 편집 탭의 Ctrl+O: 문서를 골라 그 탭 창에 열기 (이미 열린 문서면 그 탭으로). parent: 그 탭 창 (BaseWindow) */
  async function pickAndOpen(parent) {
    const tw = twOfWin(parent) ?? lastTw();
    const target = targetOf(tw);
    const r = await dialog.showOpenDialog(parent && !parent.isDestroyed() ? parent : winOf(null), {
      title: '문서 열기', properties: ['openFile', 'multiSelections'],
      defaultPath: target ? dirname(target) : app.getPath('documents'),
      filters: [{ name: '모든 문서', extensions: OPENABLE_EXTS }],
    });
    if (r.canceled) return;
    for (const f of r.filePaths) {
      try {
        if (wantsDocWindow(f)) await openAsDocument(f, tw);
        else await openPath(tw, f);
      } catch (err) { toastIn(tw, err.message); }
    }
  }
  // ── 탐색기에서 연 문서 (Word처럼 바로 편집) ─────────────

  /** 옛 형식 → 편집기로 열 새 형식 (설치된 Word·Excel·PowerPoint가 바꾼다. csv는 Office Kit가 직접) */
  const CONVERT_FOR_EDIT = { doc: 'docx', rtf: 'docx', odt: 'docx', xls: 'xlsx', xlsm: 'xlsx', ods: 'xlsx', csv: 'xlsx', ppt: 'pptx', odp: 'pptx' };
  const wantsDocWindow = (p) => settings.get('openAction') !== 'preview' && OPENABLE_EXTS.includes(extOf(p))
    && (!!editorKind(p) || !!CONVERT_FOR_EDIT[extOf(p)]);


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
    w.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(`<!doctype html><meta charset="utf-8"><body style="margin:0;font:13.5px 'Malgun Gothic',sans-serif;display:flex;align-items:center;gap:12px;height:100vh;padding:0 22px;box-sizing:border-box;color:${dark ? '#e6e8ec' : '#1f2328'};background:${dark ? '#1d2027' : '#fff'};border:1px solid ${dark ? '#333842' : '#d0d5dd'}"><b style="color:#3b5bdb">Office Kit</b><span>${esc(i18n.t(text))}</span>`)}`).catch(() => {});
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
        const { csvCellValue } = await import('../core/officefile.mjs');
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet(basename(p).replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1');
        for (const row of doc.sheets?.[0]?.rows ?? []) ws.addRow(row.map(csvCellValue));
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
   * 탐색기에서 연 문서: 편집할 수 있으면 편집 탭으로 (옛 형식은 새 형식으로 바꿔 연다). 편집할 수 없는 문서(PDF·암호 문서 등)는 문서 탭으로.
   * 이미 열린 문서(바꿔 연 원본 포함)면 그 탭으로
   */
  async function openAsDocument(p, tw = null) {
    p = docPath(p);
    tw = tw && tabWindows.has(tw.shell.id) ? tw : lastTw() ?? createTabWindowFor({ bounds: initialBounds() });
    const hit = findTab(p);
    if (hit) { activateTab(hit.tw, hit.tab.id); focusShell(hit.tw); return { path: p, mode: 'reused' }; }
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
        await openEditorTab(tw, p, { noted: true });
        return { path: p, mode: 'editor' };
      }
      const to = CONVERT_FOR_EDIT[ext];
      if (!to) { openDocTab(tw, p); return { path: p, mode: 'doc' }; }
      if (!editors.available('office').available) throw fail('편집기가 없어 미리보기로 엽니다', 'EDITOR_MISSING');
      noteOpened(p);
      const busy = busyWindow(`‘${basename(p)}’를 고칠 수 있게 바꾸는 중…`);
      let tmp;
      try { tmp = await convertForEditing(p, to); } finally { busy.close(); }
      try {
        await openEditorTab(tw, tmp, { convertedFrom: p, noted: true });
      } catch (err) {
        // 편집 탭을 못 열었으면 방금 만든 임시 사본 폴더(userData/converted/<…>)를 지운다 (탭이 열려 있으면 탭이 닫힐 때 지운다)
        if (!editors.isOpen(tmp)) { try { rmSync(dirname(tmp), { recursive: true, force: true }); } catch {} }
        throw err;
      }
      return { path: p, mode: 'converted' };
    } catch (err) {
      // 편집 탭으로 못 열면 문서 탭(미리보기)으로 (이유를 알린다)
      console.error('[doc-window]', err?.message ?? err);
      const r = openDocTab(tw, p);
      const hub = hubOf(tw.docViews.get(r.tabId)?.webContents);
      setTimeout(() => sendTo(hub?.wc, 'toast', { type: 'warn', message: err?.message ?? String(err) }), 400);
      return { path: p, mode: 'doc' };
    }
  }

  // ── 탭 창 (홈 · 문서 탭 · 편집 탭 · AI 패널) ──────────
  // 탭 창 = 껍데기(tab-window.mjs) + 탭 목록(tabs.mjs) + 홈 뷰(?view=home, 탭 'home') + AI 패널 뷰(?view=panel).
  // 편집 탭 뷰는 editor-window가 만들고 정리한다 (hostFor로 이 탭 창 껍데기를 받는다). 문서 탭 뷰(?view=doc)는 여기서

  const lastTw = () => {
    for (const id of windowOrder) { const tw = tabWindows.get(id); if (tw) return tw; }
    return tabWindows.values().next().value ?? null;
  };
  /** 그 허브(없으면 최근 탭 창)의 대화상자 부모 창 (판정 12) */
  const winOf = (hub) => ((hub?.tw && tabWindows.has(hub.tw.shell.id) ? hub.tw : lastTw())?.shell.dialogParent() ?? undefined);
  const twOfWin = (w) => [...tabWindows.values()].find((tw) => tw.shell.win === w) ?? null;
  const homeHubOf = (tw) => (tw?.home ? hubOf(tw.home.webContents) : null);

  /** 탭 창을 앞으로 */
  function focusShell(tw) {
    const w = tw?.shell.win;
    if (!w || w.isDestroyed()) return;
    if (w.isMinimized()) w.restore();
    w.show();
    w.focus();
  }

  /** 열려 있는 탭 찾기 (모든 탭 창, 대소문자·바꿔 연 원본까지) → { tw, tab } | null */
  function findTab(p) {
    for (const tw of tabWindows.values()) {
      const tab = tw.tabs.find(p);
      if (tab) return { tw, tab };
    }
    return null;
  }

  /** 홈 종류(home·doc)·패널 뷰: 지금 주 창과 같은 webPreferences, 탐색·새 창·내려받기 막기 */
  const guardedSessions = new WeakSet();
  function makeHubView(tw, role) {
    const view = new WebContentsView({
      webPreferences: {
        preload: join(here, 'preload.cjs'),
        contextIsolation: true, nodeIntegration: false, sandbox: true,
        plugins: true,            // 내장 PDF 뷰어
        webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
        backgroundThrottling: !SMOKE,
        devTools: !app.isPackaged,
      },
    });
    view.setBackgroundColor?.(nativeTheme.shouldUseDarkColors ? '#16181d' : '#f4f5f7');
    const wc = view.webContents;
    viewRecs.set(wc, { kinds: role === 'panel' ? ['panel'] : ['home'], tw, role });
    wc.once('destroyed', () => viewRecs.delete(wc));
    // 새 창·이동 막기: 웹 주소는 기본 브라우저로
    wc.setWindowOpenHandler(({ url }) => {
      if (/^https?:\/\//i.test(url)) void openWebLink(url, wc);
      return { action: 'deny' };
    });
    bindTabKeys(wc, tw);
    wc.on('will-navigate', (e) => e.preventDefault());
    wc.on('will-frame-navigate', (e) => {
      if (e.isMainFrame) { e.preventDefault(); return; }
      const u = e.url;
      // 미리보기 PDF와 내장 PDF 뷰어만 허용
      if (u === 'about:blank' || u.startsWith('chrome-extension://') || isPreviewUrl(u)) return;
      e.preventDefault();
      if (/^https?:\/\//i.test(u)) void openWebLink(u, wc);
    });
    wc.on('will-attach-webview', (e) => e.preventDefault());
    // 내려받기는 PDF 뷰어의 저장 단추(미리보기 PDF)만 허용 → 저장 위치를 묻는다 (이름·폴더는 그 뷰 허브의 문서). 세션마다 한 번만 건다
    const ses = wc.session;
    if (ses && !guardedSessions.has(ses)) {
      guardedSessions.add(ses);
      ses.on('will-download', (e, item, from) => {
        if (!isPreviewUrl(item.getURL())) { e.preventDefault(); return; }
        const cur = hubOf(from)?.current ?? null;
        const name = cur ? `${basename(cur.path).replace(/\.[^.]+$/, '')}.pdf` : item.getFilename();
        item.setSaveDialogOptions(localizeDialogOptions({ title: 'PDF로 저장', defaultPath: join(cur ? dirname(cur.path) : app.getPath('documents'), name), filters: [{ name: 'PDF 문서', extensions: ['pdf'] }] }, i18n.t));
      });
    }
    wc.on('render-process-gone', (_e, d) => console.error('[renderer] 종료:', d.reason));
    if (SMOKE) wc.on('console-message', (e) => { if (e.level === 'error' || e.level === 'warning') console.log('[renderer]', e.message); });
    void Promise.resolve(wc.loadURL(`${pathToFileURL(INDEX_HTML).href}?view=${role}`)).catch(() => {});
    return view;
  }

  /** 탭 창 하나 만들기: 껍데기 + 탭 목록 + 홈 뷰 + AI 패널 뷰 (너비·접힘은 settings.chatWidth/chatCollapsed) */
  function createTabWindowFor({ bounds, maximized = false } = {}) {
    const shell = createTabWindow({
      BaseWindow, WebContentsView, bounds, maximized, icon: APP_ICON, dark: !!nativeTheme.shouldUseDarkColors, devTools: !app.isPackaged,
      panelWidth: settings.get('chatWidth'), panelCollapsed: !!settings.get('chatCollapsed'),
    });
    const tw = { shell, tabs: createTabList(), home: null, panel: null, docViews: new Map(), edViews: new Map(), titles: new Map(), closing: null, widthTimer: 0 };
    tabWindows.set(shell.id, tw);
    windowOrder.unshift(shell.id);
    viewRecs.set(shell.tabstrip.webContents, { kinds: ['tabstrip'], tw, role: 'tabstrip' });
    shell.tabstrip.webContents.once?.('destroyed', () => viewRecs.delete(shell.tabstrip.webContents));
    bindTabKeys(shell.tabstrip.webContents, tw);
    tw.tabs.on('change', () => sendTabs(tw));
    tw.home = makeHubView(tw, 'home');
    makeHub(tw.home.webContents, tw, 'home');
    shell.attach(tw.home, 'content');
    tw.panel = makeHubView(tw, 'panel');
    shell.attach(tw.panel, 'panel');
    shell.focusTab(tw.home);
    shell.setTitle('Office Kit');
    shell.on('focus', () => {
      const i = windowOrder.indexOf(shell.id);
      if (i > 0) { windowOrder.splice(i, 1); windowOrder.unshift(shell.id); }
    });
    // 창 닫기(X·Alt+F4) → 편집 탭을 탭 차례로 닫아 본다 (판정 13). AI 확인·끝내기 차례는 Task 10
    shell.on('close-requested', () => { void closeWindow(tw).catch((err) => console.error('[tabs] 탭 창을 닫지 못했습니다:', err)); });
    shell.on('closed', () => {
      tabWindows.delete(shell.id);
      const i = windowOrder.indexOf(shell.id);
      if (i >= 0) windowOrder.splice(i, 1);
    });
    sendTabs(tw);
    return tw;
  }

  /** 탭 목록을 탭 줄에 보낸다 */
  function sendTabs(tw) {
    const panel = tw.shell.setPanel();
    tw.shell.sendTabs({
      tabs: tw.tabs.tabs.map((t) => ({ id: t.id, type: t.type, title: t.title, path: t.path, dirty: !!t.dirty, loaded: t.loaded !== false, zone3: !!t.zone3, busy: !!t.path && aiBusyFor(t.path) })),
      active: tw.tabs.active?.id ?? null, panelCollapsed: panel.collapsed,
    });
  }

  /** 탭의 뷰 (홈 → 홈 뷰, 문서 탭 → 문서 탭 뷰, 편집 탭 → 편집 뷰) */
  function viewOfTab(tw, t) {
    if (!t) return null;
    if (t.type === 'home') return tw.home;
    if (t.type === 'doc') return tw.docViews.get(t.id) ?? null;
    return tw.edViews.get(t.id) ?? editors?.viewOf(t.path) ?? null;
  }

  /** 창 제목 = 활성 탭 제목 (홈은 'Office Kit') */
  function titleOfTab(tw, t) {
    if (!t || t.type === 'home') return 'Office Kit';
    return tw.titles.get(t.id) ?? `${t.title || basename(t.path ?? '')} - Office Kit`;
  }

  /** 탭 하나를 보이고 AI 대상을 패널에 알린다 */
  function activateTab(tw, id) {
    const t = tw.tabs.get(id);
    if (!t) return null;
    if (tw.tabs.active?.id !== id) tw.tabs.activate(id);
    const view = viewOfTab(tw, t);
    if (view) { try { tw.shell.focusTab(view); } catch (err) { console.error('[tabs] 탭 보이기:', err?.message ?? err); } }
    tw.shell.setTitle(titleOfTab(tw, t));
    sendTabs(tw);
    pushTarget(tw);
    return t;
  }

  /** 탭 하나를 목록에서 뺀다 (활성 탭이었으면 이웃 탭으로) */
  function removeTab(tw, id) {
    if (!tw.tabs.get(id)) return;
    const wasActive = tw.tabs.active?.id === id;
    const { active } = tw.tabs.remove(id);
    tw.titles.delete(id);
    tw.edViews.delete(id);
    if (wasActive) activateTab(tw, active.id);
    else pushTarget(tw);
    pushEditorState();
  }

  /** 뷰 wc 닫기 (정리는 'destroyed'에서) */
  const closeWc = (wc) => { try { if (wc && !wc.isDestroyed()) wc.close(); } catch {} };

  /** 문서 탭 닫기: 뷰를 떼고 wc를 닫는다 (허브 정리는 wc 'destroyed'에서: 감시 끄기·hubs.delete) */
  function closeDocTab(tw, id) {
    const view = tw.docViews.get(id);
    tw.docViews.delete(id);
    if (view) { try { tw.shell.detach(view); } catch {} }
    removeTab(tw, id);
    if (view) closeWc(view.webContents);
  }

  /** 탭 닫기 (탭 줄 ×·Delete): 편집 탭은 저장 확인, 문서 탭은 바로. → 닫혔으면 true */
  async function closeTab(tw, id) {
    const t = tw.tabs.get(id);
    if (!t || t.type === 'home') return false;
    if (t.type === 'doc') { closeDocTab(tw, id); return true; }
    const r = editors ? await editors.close(t.path) : null;
    // 편집 뷰가 이미 없으면 탭만 뺀다
    if (r === null) removeTab(tw, id);
    return r !== false;
  }

  /** 탭 단축키 실행 (어느 뷰에서 눌렀든 그 탭 창에서): 홈·닫기·이웃 탭·번째 탭. 새 창(Ctrl+N)은 아직 없다 */
  function runTabCommand(tw, cmd) {
    if (!tw || !cmd || !tabWindows.has(tw.shell.id)) return;
    if (cmd.cmd === 'home') activateTab(tw, 'home');
    else if (cmd.cmd === 'next' || cmd.cmd === 'prev') {
      const t = tw.tabs.step(cmd.cmd === 'next' ? 1 : -1);
      if (t) activateTab(tw, t.id);
    } else if (cmd.cmd === 'index') {
      const t = tw.tabs.at(cmd.n);
      if (t) activateTab(tw, t.id);
    } else if (cmd.cmd === 'close') {
      const t = tw.tabs.active;
      // 홈 탭은 닫지 않는다 (판정 6)
      if (t && t.type !== 'home') void closeTab(tw, t.id).catch((err) => console.error('[tabs] 탭 닫기:', err?.message ?? err));
    }
  }

  /** 홈·문서 탭·패널·탭 줄 뷰의 키 입력: 탭 단축키면 가로채 그 탭 창에서 실행 */
  function bindTabKeys(wc, tw) {
    wc.on('before-input-event', (e, input) => {
      const tc = tabCommandOf(input);
      if (!tc) return;
      e.preventDefault();
      runTabCommand(tw, tc);
    });
  }

  /** 그 탭 창의 AI 대상 (판정 3): 편집 탭 → 그 문서, 문서 탭 → 그 뷰의 문서, 홈 → null */
  function targetOf(tw) {
    const t = tw?.tabs.active;
    if (!t || t.type === 'home') return null;
    if (t.type === 'editor') return t.path;
    const view = tw.docViews.get(t.id);
    return (view && hubOf(view.webContents)?.current?.path) ?? null;
  }

  /** 패널 뷰에 AI 대상을 민다 */
  function pushTarget(tw) {
    if (!tw?.panel) return;
    sendTo(tw.panel.webContents, 'ai:target', { doc: targetInfo(targetOf(tw)) });
  }

  /** 문서 탭(?view=doc) 하나 열기: 그 뷰의 허브에 문서를 연다 (화면 준비를 기다리지 않는다 — app:init이 doc을 준다) */
  function openDocTab(tw, p, { activate = true } = {}) {
    p = docPath(p);
    const view = makeHubView(tw, 'doc');
    const t = tw.tabs.add({ type: 'doc', path: p, title: basename(p) }, { activate: false });
    const hub = makeHub(view.webContents, tw, t.id);
    tw.docViews.set(t.id, view);
    tw.shell.attach(view, 'content');
    try { openDocument(hub, p); } catch (err) { closeDocTab(tw, t.id); throw err; }
    if (activate) activateTab(tw, t.id);
    pushEditorState();
    return { tabId: t.id, mode: 'doc', path: p, reused: false, editor: editorState(p), notice: null };
  }

  // 편집 탭을 어디에 붙일지 (openEditorTab이 정하고 hostFor가 쓴다): docKey(path) → { tw, path, replaceId }
  // openEditor는 서버·세션을 기다리므로 여러 열기가 겹칠 수 있다 → 문서마다 따로 둔다
  const placings = new Map();
  // hostFor가 막 만든 편집 탭 (첫 onTabChange가 뷰를 잇는다): docKey(path) → { tw, tabId, path }
  const pendingTabs = new Map();

  /** editor-window의 hostFor: 편집 뷰를 붙일 탭 창 껍데기. 탭을 더하거나 문서 탭을 같은 자리에서 바꾼다 (판정 7) */
  function editorHostFor(info) {
    const pl = placings.get(docKey(info.path)) ?? null;
    const tw = (pl && tabWindows.has(pl.tw.shell.id) ? pl.tw : null) ?? lastTw() ?? createTabWindowFor({ bounds: initialBounds() });
    const convertedFrom = info.convertedFrom ?? null;
    const tab = { type: 'editor', path: info.path, convertedFrom, title: basename(convertedFrom ?? info.path) };
    let t;
    if (pl?.replaceId && tw.tabs.get(pl.replaceId)?.type === 'doc') {
      const old = pl.replaceId;
      const view = tw.docViews.get(old);
      t = tw.tabs.replace(old, tab);
      tw.titles.delete(old);
      tw.docViews.delete(old);
      if (view) { try { tw.shell.detach(view); } catch {} closeWc(view.webContents); }
    } else {
      t = tw.tabs.add(tab);
    }
    pendingTabs.set(docKey(info.path), { tw, tabId: t.id, path: info.path });
    return tw.shell;
  }

  /** 그 경로의 편집 탭 → { tw, tab } (편집 뷰로 먼저, 없으면 경로로) */
  function findEditorTab(p, view = editors?.viewOf(p) ?? null) {
    for (const tw of tabWindows.values()) {
      if (view) for (const [id, v] of tw.edViews) if (v === view && tw.tabs.get(id)) return { tw, tab: tw.tabs.get(id) };
    }
    for (const tw of tabWindows.values()) {
      const tab = tw.tabs.tabs.find((x) => x.type === 'editor' && samePath(x.path, p));
      if (tab) return { tw, tab };
    }
    const pend = pendingTabs.get(docKey(p));
    if (pend && pend.tw.tabs.get(pend.tabId)) return { tw: pend.tw, tab: pend.tw.tabs.get(pend.tabId) };
    return null;
  }

  /** editor-window의 onTabChange: 탭 제목·저장 안 함·경로(다른 이름으로 저장)를 고치고, 닫혔으면 탭을 뺀다 */
  function editorTabChanged(p, change = {}) {
    const view = editors?.viewOf(p) ?? null;
    const hit = findEditorTab(p, view);
    if (!hit) return;
    const { tw, tab } = hit;
    if (change.closed) {
      removeTab(tw, tab.id);
      setTimeout(editorClosed, 0);
      return;
    }
    if (view) tw.edViews.set(tab.id, view);
    const patch = {};
    if (!samePath(tab.path, p)) Object.assign(patch, { path: p, convertedFrom: null, title: basename(p) });
    if (change.dirty !== undefined) patch.dirty = !!change.dirty;
    tw.tabs.update(tab.id, patch);
    if (typeof change.title === 'string') tw.titles.set(tab.id, change.title);
    if (tw.tabs.active?.id === tab.id) {
      tw.shell.setTitle(titleOfTab(tw, tw.tabs.get(tab.id)));
      pushTarget(tw);
    }
    pushEditorState();
  }

  /** 편집 탭이 닫힌 뒤: 상태 맞추기 */
  function editorClosed() {
    pushEditorState();
  }

  /**
   * 편집 탭으로 열기 (이미 열려 있으면 그 탭으로) → { tabId, mode: 'editor'|'reused', path, kind, reused, notice, editor }
   * replaceId: 바꿀 문서 탭 (판정 7), noted: 처음 연 상태 버전·최근 목록을 이미 남겼다
   */
  async function openEditorTab(tw, p, { replaceId = null, noted = false, convertedFrom = null } = {}) {
    if (!editors) throw fail(`편집 창 기능을 불러오지 못했습니다: ${editorLoadError}`, 'EDITOR_MISSING');
    if (!noted) noteOpened(p);
    const k = docKey(p);
    const mine = { tw, path: p, replaceId };
    placings.set(k, mine);
    pendingTabs.delete(k);
    let r;
    try {
      r = await editors.openEditor(p, convertedFrom ? { convertedFrom } : undefined);
    } catch (err) {
      // 뷰를 붙이기 전에 실패했으면 이 열기가 막 만든 탭을 뺀다
      const pend = pendingTabs.get(k);
      if (pend && !editors.isOpen(pend.path)) removeTab(pend.tw, pend.tabId);
      throw err;
    } finally {
      if (placings.get(k) === mine) placings.delete(k);
      pendingTabs.delete(k);
    }
    const hit = findEditorTab(r.path);
    if (hit) { activateTab(hit.tw, hit.tab.id); if (hit.tw !== tw) focusShell(hit.tw); }
    // 다른 길로 이미 열려 있던 편집 탭으로 갔으면 바꾸려던 문서 탭은 닫는다 (같은 문서 탭 둘을 두지 않는다)
    if (r.reused && replaceId && tw.tabs.get(replaceId)?.type === 'doc') closeDocTab(tw, replaceId);
    pushEditorState();
    // 메모리 안내(D12): 그 탭 창을 부모로 한 정보 대화상자 (기다리지 않는다 — 열기를 막지 않게)
    const notice = typeof r.notice === 'string' && r.notice ? r.notice : null;
    if (notice) {
      const owner = hit?.tw ?? tw;
      void Promise.resolve(dialog.showMessageBox(owner.shell.dialogParent(), {
        type: 'info', title: 'Office Kit', message: notice, buttons: ['확인'], defaultId: 0, noLink: true,
      })).catch(() => {});
    }
    return {
      tabId: hit?.tab.id ?? null, mode: r.reused ? 'reused' : 'editor', path: r.path, kind: r.kind, reused: !!r.reused,
      notice: typeof r.notice === 'string' && r.notice ? r.notice : null, editor: editorState(r.path),
    };
  }

  /**
   * 문서 열기 (홈의 열기·끌어놓기·최근 목록, 탐색기, 편집 탭의 Ctrl+O): 이미 열린 탭이면 그 탭으로(다른 탭 창이면 그 창을 앞으로),
   * 설정이 '편집'이고 편집할 수 있으면 편집 탭, 아니면 문서 탭 → { tabId, mode: 'editor'|'doc'|'reused', path, reused, editor, notice }
   */
  async function openPath(tw, p, { activate = true } = {}) {
    p = docPath(p);
    const hit = findTab(p);
    if (hit) {
      if (activate) { activateTab(hit.tw, hit.tab.id); focusShell(hit.tw); }
      if (!SMOKE) { try { settings.addRecent(hit.tab.path); } catch {} }
      return { tabId: hit.tab.id, mode: 'reused', path: hit.tab.path, reused: true, editor: editorState(hit.tab.path), notice: null };
    }
    tw = tw && tabWindows.has(tw.shell.id) ? tw : lastTw() ?? createTabWindowFor({ bounds: initialBounds() });
    if (wantsEditor(p)) {
      try {
        editorCheck(p);
        return await openEditorTab(tw, p);
      } catch (err) {
        // 편집 탭으로 못 열면 문서 탭으로. 편집기 묶음이 없는 것은 조용히 (예전 화면의 autoEdit와 같게), 나머지는 이유를 알린다
        if (!err?.expected) console.error('[open]', err);
        const r = openDocTab(tw, p, { activate });
        if (err?.code !== 'EDITOR_MISSING') {
          const hub = hubOf(tw.docViews.get(r.tabId)?.webContents);
          setTimeout(() => sendTo(hub?.wc, 'toast', { type: 'warn', message: err?.message ?? String(err) }), 400);
        }
        return r;
      }
    }
    return openDocTab(tw, p, { activate });
  }

  /** 탭 창 닫기: 편집 탭을 탭 차례로 닫아 보고(취소하면 그 탭부터 모두 그대로, 판정 13), 다 닫히면 문서 탭·홈·패널 뷰를 닫고 창을 없앤다 */
  function closeWindow(tw) {
    if (tw.closing) return tw.closing;
    tw.closing = (async () => {
      const paths = tw.tabs.tabs.filter((t) => t.type === 'editor' && t.path).map((t) => t.path);
      if (editors && !(await editors.closeMany(paths))) return false;
      const w = tw.shell.win;
      if (!SMOKE && w && !w.isDestroyed()) {
        try { settings.set({ window: { ...w.getNormalBounds(), maximized: w.isMaximized() } }); } catch {}
      }
      for (const view of [...tw.docViews.values(), tw.home, tw.panel]) {
        if (!view) continue;
        try { tw.shell.detach(view); } catch {}
        closeWc(view.webContents);
      }
      tw.docViews.clear();
      tw.shell.destroy();
      if (appQuitting) app.quit();
      return true;
    })().finally(() => { tw.closing = null; });
    return tw.closing;
  }

  /** 편집 탭의 'AI · 도구': 그 탭을 보이고 같은 탭 창의 AI 패널을 펼쳐 초점을 준다 (옆 창 없음) */
  function showPanelFor(p) {
    const hit = findEditorTab(p) ?? findTab(p);
    if (!hit) return;
    activateTab(hit.tw, hit.tab.id);
    const r = hit.tw.shell.setPanel({ collapsed: false });
    try { settings.set({ chatCollapsed: r.collapsed }); } catch {}
    sendTabs(hit.tw);
    hit.tw.panel?.webContents.focus();
  }

  // ── AI 대화 ──────────────────────────────────────────

  const key = (p) => (process.platform === 'win32' ? resolve(p).toLowerCase() : resolve(p));
  const sessions = new Map();  // 문서 경로 → { provider, sessionId }
  // 세션 키 → 그 대화에 마지막으로 넘긴 인터넷 재료(Zone 3)의 key 목록 (빠지면 이어 하지 않고 새 대화)
  const sessionZone3 = new Map();
  // chatId → { handle, path: 화면 대화 칸 키(지금 문서), edit: 고칠 문서 | null, read: 재료 목록, pidDir, ownerWc: 보낸 쪽 화면(wc) }
  const chats = new Map();
  const MAX_MATERIALS = 10;
  let providerCache = null;

  /** AI가 고칠 수 있는 문서인가: PDF는 안 되고, 인터넷에서 받은 파일(Zone 3)은 한글 문서만 */
  const canEdit = (p) => extOf(p) !== 'pdf' && (core.isHwp(extOf(p)) || !core.isFromInternet(p));
  /** 진행 중인 대화들이 고치는 문서의 key */
  const runningEdits = () => new Set([...chats.values()].filter((c) => c.edit).map((c) => key(c.edit)));

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

  /** 편집 창에 저장 안 한 내용이 있는가: 편집 창이 알린 상태이거나, 열린 창의 자동 저장 사본이 원본 파일보다 새것 (OV6) */
  const unsaved = (p) => {
    if (editors?.isDirty(p)) return true;
    if (!editors?.isOpen(p)) return false;
    try {
      const d = editors.draftPath?.(p);
      return !!d && existsSync(d) && statSync(d).mtimeMs > statSync(p).mtimeMs;
    } catch { return false; }
  };

  /** 모든 탭 창의 편집 탭 + 문서 탭(PDF 포함) 문서 경로 (AI 재료 후보) */
  const openDocPaths = () => [
    ...(editors?.list() ?? []),
    ...allHubs().filter((h) => h.tabId !== 'home' && h.current).map((h) => h.current.path),
  ];

  /** AI 재료 후보: 고칠 문서(그 탭 창의 AI 대상)와 재료로 넘길 수 있는 열린 문서들 */
  function listMaterials(tw) {
    const target = targetOf(tw);
    if (!target) return { edit: null, open: [] };
    const busy = runningEdits();
    const cur = key(target);
    const item = (p) => ({
      path: p, name: basename(p), kind: kindOf(p), zone3: core.isFromInternet(p),
      busy: busy.has(key(p)), dirty: unsaved(p), current: key(p) === cur,
    });
    const edit = canEdit(target) ? item(target) : null;
    const seen = new Set(edit ? [cur] : []);
    const open = [];
    for (const p of [target, ...openDocPaths()]) {
      if (seen.has(key(p))) continue;
      seen.add(key(p));
      open.push(item(p));
    }
    return { edit, open };
  }

  /**
   * AI 대화 시작. to = { tw: 그 탭 창, path: 대화 대상(AI 대상) }. 이벤트는 그 탭 창의 AI 패널로만 (chatId → ownerWc).
   * announce: 시작 전에 패널에 'start' 이벤트를 보내 패널이 대화를 만들게 한다 (문자열이면 말풍선 글, 새 문서 AI 시작)
   */
  async function sendChat(to, { prompt, editorOk, target, outDir, materialsOk, saveMaterials, announce } = {}) {
    if (typeof prompt !== 'string' || !prompt.trim()) throw fail('보낼 내용을 입력하세요');
    if (prompt.length > 20_000) throw fail('메시지가 너무 깁니다 (2만 자 이하)');
    if (!to?.path) throw fail('열린 문서가 없습니다');
    if (!existsSync(to.path)) throw fail(`문서를 찾을 수 없습니다: ${basename(to.path)}`, 'NOT_FOUND');
    const doc = { path: to.path };
    const ownerWc = to.tw?.panel?.webContents ?? null;
    // uiPath = 화면 대화 칸 키(지금 문서). 이어 하기 세션·되돌리기도 이 키를 쓴다 (재료 전용이어도)
    const uiPath = doc.path;
    const k = key(uiPath);
    if (restoring.has(k)) throw fail('문서를 원래대로 되돌리는 중입니다. 끝난 뒤 다시 보내 주세요.');
    // 대화 대상 확정: 고칠 문서는 지금 문서(고칠 수 없는 문서면 없음), 재료는 화면이 넘긴 목록
    const edit = canEdit(uiPath) ? uiPath : null;
    const want = target && typeof target === 'object' ? target : {};
    // 화면이 보던 고칠 문서와 지금 문서가 다르면 (보내는 사이 문서가 바뀜) 엉뚱한 문서를 고치지 않게 거절
    if (typeof want.edit === 'string' && key(want.edit) !== key(edit ?? '')) throw fail('대화 대상이 바뀌었어요. 다시 보내 주세요.', 'TARGET_CHANGED');
    const seen = new Set(edit ? [key(edit)] : []);
    let read = [];
    for (const p of Array.isArray(want.read) ? want.read : []) {
      const f = existingFile(p, OPENABLE_EXTS);
      if (seen.has(key(f))) continue;
      seen.add(key(f));
      read.push(f);
    }
    // 고칠 문서가 없으면 지금 문서가 맨 앞 재료
    if (!edit) read = [uiPath, ...read.filter((p) => key(p) !== k)];
    if (read.length > MAX_MATERIALS) throw fail('재료는 10개까지 넘길 수 있어요');
    // AI가 고치는 중인 문서는 또 고치지도, 다른 대화의 재료로 넘기지도 않는다 (D14 · 열림 검사보다 먼저)
    const checkBusy = () => {
      const busy = runningEdits();
      if (edit && busy.has(key(edit))) throw fail('AI가 이 문서를 이미 고치는 중이에요. 끝난 뒤 다시 보내 주세요.', 'AI_BUSY');
      const p = read.find((m) => busy.has(key(m)));
      if (p) throw fail(`‘${basename(p)}’은(는) AI가 고치는 중이라 재료로 넘길 수 없어요. 그 대화가 끝난 뒤 보내 주세요.`, 'MATERIAL_BUSY');
    };
    checkBusy();
    // 재료는 지금 문서이거나 탭으로 열려 있는 문서만 (모든 탭 창의 편집 탭·문서 탭)
    const open = new Set(openDocPaths().map(key));
    for (const p of read) {
      if (key(p) !== k && !open.has(key(p))) throw fail(`열린 문서만 재료로 넘길 수 있어요: ${basename(p)}`, 'MATERIAL_NOT_OPEN');
    }
    // 저장 안 한 재료: AI는 저장된 파일을 읽으므로 편집 창 내용과 엇갈린다 (화면이 '그대로 보내기'면 materialsOk, '저장하고 보내기'면 saveMaterials)
    const dirty = read.filter(unsaved);
    if (dirty.length && materialsOk !== true) {
      if (saveMaterials === true) {
        for (const p of dirty) {
          if ((await editors.save(p)) !== true) throw fail(`‘${basename(p)}’을(를) 저장하지 못했어요. 편집 창에서 저장한 뒤 다시 보내 주세요.`, 'MATERIAL_SAVE_FAILED');
        }
      } else {
        const names = dirty.map((p) => `‘${basename(p)}’`).join(', ');
        throw fail(`${names}에 저장 안 한 편집이 있어 저장된 내용을 읽습니다.`, 'MATERIAL_DIRTY');
      }
    }
    // 편집 창에 저장 안 한 내용이 있으면 먼저 묻는다: AI는 저장된 파일을 고치므로 편집 창 내용과 엇갈린다
    // (화면이 '그대로 보내기'를 고르면 editorOk:true로 다시 보낸다)
    if (edit && editorOk !== true && editors?.isDirty(edit)) {
      throw fail('편집 창에 저장하지 않은 내용이 있어요. 편집 창에서 먼저 저장하거나, 그대로 보내기를 골라 주세요.', 'EDITOR_DIRTY');
    }
    // 재료 전용이면 새 파일을 만들 출력 폴더 (주지 않으면 첫 재료의 폴더)
    let out = null;
    if (!edit) {
      if (outDir != null) {
        if (typeof outDir !== 'string' || outDir.length > 4096 || !isAbsolute(outDir) || !isDir(outDir)) throw fail('출력 폴더를 찾을 수 없어요', 'BAD_OUT_DIR');
        out = resolve(outDir);
      } else {
        out = dirname(read[0]);
      }
    }
    const zone3Keys = read.filter((p) => core.isFromInternet(p)).map(key);
    const zone3 = zone3Keys.length > 0;
    const gen = restoreGen.get(k) ?? 0;
    const ai = await loadAi();
    const { providers: list, selected } = await providers();
    if (!selected) throw fail('사용할 수 있는 AI가 없습니다. Claude Code나 Codex를 설치하고 로그인하세요.', 'NO_PROVIDER');
    const base = await mcpSpec();
    // 준비하는 사이 '원래대로'를 눌렀으면 시작하지 않는다 (되돌린 파일을 다시 고치지 않게)
    if (restoring.has(k) || (restoreGen.get(k) ?? 0) !== gen) throw fail('문서를 원래대로 되돌려서 이 요청은 보내지 않았습니다.');
    // 준비하는 사이 다른 대화가 같은 문서를 고치기 시작했을 수 있다
    checkBusy();
    // 지난 대화에 넘긴 인터넷 재료가 이번에 빠졌으면 이어 하지 않는다 (그 내용이 대화 기록에 남아 있으므로 새 대화)
    let newChat = false;
    const keep = new Set(read.map(key));
    if ((sessionZone3.get(k) ?? []).some((z) => !keep.has(z)) && sessions.has(k)) {
      sessions.delete(k);
      newChat = true;
    }
    const chatId = randomUUID();
    // MCP 서버가 띄운 Office의 PID를 대화마다 따로 적는다 → 멈추거나 끝나면 남은 Office를 끝낸다
    const pidDir = join(userData, COM_PIDS, chatId);
    let comPidDir;
    try { mkdirSync(pidDir, { recursive: true }); comPidDir = pidDir; } catch (err) { console.error('[ai:send] PID 폴더', err?.message ?? err); }
    // 앱 안 AI의 권한 (MCP 서버가 OFFICEKIT_* 환경 변수로 막는다)
    //   고칠 문서가 있으면: 읽기는 ① 문서 폴더 ② 재료 파일 ③ AI가 새로 만든 파일, 저장은 문서 폴더에,
    //     고치기는 그 문서와 새로 만든 파일만 (다른 파일은 덮어쓰지 않음)
    //   인터넷 재료가 있으면 ① 대신 문서 파일 하나만 읽는다 (폴더 통째 읽기 없음)
    //   재료 전용(고칠 문서 없음)이면 재료 파일만 읽고, 출력 폴더에 새 파일만 만든다
    const target2 = edit ? { edit, read, zone3 } : { edit: null, read, outDir: out, zone3 };
    const mcp = { ...base, env: { ...(base.env ?? {}), ...ai.aiMcpEnv(target2, { comPidDir }) } };
    // AI가 고치기 전 상태를 버전으로 남긴다 (대화 한 번마다. 내용이 같으면 있던 것을 다시 씀)
    // → 끝날 때 이 버전과 다르면 화면에 '바뀐 내용 보기 · 이 대화 전으로 되돌리기'를 띄운다. 고칠 문서가 없으면 남기지 않는다
    let before = null;
    if (edit) { try { before = keepVersion(edit, 'ai'); } catch (err) { console.error('[ai:send] 백업', err?.message ?? err); } }
    const sess = sessions.get(k);
    // 모델은 AI마다 따로 (다른 AI의 모델 이름을 넘기면 모든 요청이 실패한다)
    const model = settings.get('models')?.[selected] || undefined;
    let errored = false;
    const path = uiPath;
    const onEvent = (ev) => {
      if (!ev || typeof ev !== 'object') return;
      if (ev.type === 'done' && ev.sessionId) sessions.set(k, { provider: selected, sessionId: String(ev.sessionId) });
      if (ev.type === 'error') errored = true;
      try { sendTo(ownerWc, 'ai:event', { ...ev, chatId, path }); } catch {
        sendTo(ownerWc, 'ai:event', { type: ev.type, text: ev.text, message: ev.message, id: ev.id, name: ev.name, ok: ev.ok, summary: ev.summary, delta: ev.delta, chatId, path });
      }
    };
    // 패널이 대화를 만들도록 다른 이벤트보다 먼저 알린다
    if (announce) sendTo(ownerWc, 'ai:event', { type: 'start', chatId, path, prompt: typeof announce === 'string' ? announce : prompt });
    let handle;
    try {
      handle = ai.startChat({
        provider: selected, prompt, sessionId: sess?.provider === selected ? sess.sessionId : undefined,
        // 대화 기록은 같은 작업 폴더에서만 이어지므로 ai.mjs의 고정 폴더를 쓴다
        cwd: typeof ai.defaultWorkspace === 'function' ? ai.defaultWorkspace() : join(userData, 'ai'),
        filePath: edit, materials: read, outDir: edit ? null : out, mcp, model, onEvent, lang: i18n.lang,
      });
    } catch (err) {
      // 대화를 시작하지 못했으면 만든 PID 폴더도 남기지 않는다
      if (comPidDir) { try { rmSync(comPidDir, { recursive: true, force: true }); } catch { /* 지우기 실패 */ } }
      throw err;
    }
    sessionZone3.set(k, zone3Keys);
    chats.set(chatId, { handle, path, edit, read, pidDir, ownerWc });
    pushAiState();
    Promise.resolve(handle.done)
      .then((r) => { if (r?.sessionId) sessions.set(k, { provider: selected, sessionId: String(r.sessionId) }); })
      .catch((err) => { if (!errored) sendTo(ownerWc, 'ai:event', { type: 'error', message: err?.message ?? String(err), chatId, path }); })
      .finally(() => {
        chats.delete(chatId);
        killChatOffice(pidDir, true);
        // 바뀜 확인은 고친 문서로 (재료 전용이면 before가 없어 바뀜 알림도 없다)
        let changed = false;
        try { changed = !!before?.hash && existsSync(edit) && currentHash(edit) !== before.hash; } catch {}
        sendTo(ownerWc, 'ai:event', { type: 'end', chatId, path, ...(changed ? { changed: true, before: before.file } : {}) });
        pushAiState();
      });
    return { chatId, provider: list.find((p) => p.id === selected)?.name ?? selected, newChat };
  }

  function cancelChat(chatId) {
    if (typeof chatId !== 'string') throw fail('잘못된 요청입니다');
    const c = chats.get(chatId);
    if (c) { try { c.handle.cancel(); } catch {} }
    return { canceled: !!c };
  }

  // ── 설정 ─────────────────────────────────────────────

  const PUBLIC_KEYS = ['provider', 'models', 'previewMode', 'openAction', 'theme', 'lang', 'chatWidth', 'chatCollapsed'];
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
      } else if (k === 'lang') {
        if (!LANG_SETTINGS.includes(v)) throw fail('잘못된 언어입니다');
        clean.lang = v;
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
    await m.registerFileTypes({ exe, lang: i18n.lang });
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
      const r = await dialog.showOpenDialog(winOf(null), {
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

  /** 보낸 쪽이 탭 창의 홈·문서 탭·패널·탭 줄 뷰의 최상위 틀이면 그 뷰 종류, 아니면 null */
  function senderKinds(e) {
    const rec = viewRecs.get(e.sender);
    if (!rec || e.sender.isDestroyed?.() || !e.senderFrame || e.senderFrame !== e.sender.mainFrame) return null;
    return rec.kinds;
  }

  /** 보낸 쪽 문맥: { wc: 보낸 화면, kinds: 뷰 종류, hub: 그 화면의 허브 | null, tw: 그 탭 창 | null } */
  const ctxOf = (e, kinds) => ({ wc: e.sender, kinds, hub: hubOf(e.sender), tw: viewRecs.get(e.sender)?.tw ?? null });

  // 모든 요청은 { ok, value } / { ok:false, error, code } 로 돌려준다. 채널과 보낸 쪽은 IPC 표(ipc-channels.mjs)로 검사한다
  // fn(ctx, ...args): ctx = 보낸 쪽 문맥 (ctxOf)
  function handle(channel, fn) {
    channelSpec(channel, 'invoke');
    ipcMain.handle(channel, async (e, ...args) => {
      try {
        const kinds = senderKinds(e);
        if (!allows(channel, kinds)) throw fail('허용되지 않은 요청입니다');
        return { ok: true, value: await fn(ctxOf(e, kinds), ...args) };
      } catch (err) {
        if (!err?.expected && err?.code !== 'MODULE_MISSING') console.error(`[${channel}]`, err);
        // detail: 엔진이 낸 원래 글 (화면이 '자세히'로 접어 보여 준다)
        return { ok: false, error: err?.message ?? String(err), code: err?.code ?? null, ...(err?.detail ? { detail: String(err.detail).slice(0, 4000) } : {}) };
      }
    });
  }

  /** ipcMain.on 등록 (send·sendSync). 열린 채널이 아니면 탭 창 뷰의 최상위 틀이 보낸 것만 받는다. fn(e, ctx, ...args) */
  function onIpc(channel, fn) {
    channelSpec(channel, 'on');
    ipcMain.on(channel, (e, ...args) => {
      const kinds = senderKinds(e);
      if (!allows(channel, kinds)) return;
      return fn(e, ctxOf(e, kinds), ...args);
    });
  }

  handle('app:init', ({ hub, tw, wc }) => {
    if (hub) hub.rendererReady = true;
    const pending = hub?.pendingFile ?? null;
    // 명령줄로 받은 파일은 첫 탭 창 홈 뷰가 준비되면 그 파일 탭을 연다 (응답이 먼저 가도록 다음 틱에)
    if (pending) {
      hub.pendingFile = null;
      setTimeout(() => {
        openPath(tw, pending)
          .then((r) => { const cur = targetOf(tw); if (SMOKE_DIFF && cur && r?.mode === 'doc') copyFileSync(SMOKE_DIFF, cur); })
          .catch((err) => sendTo(hub.wc, 'toast', { type: 'error', message: err.message }));
      }, 0);
    }
    return {
      version: app.getVersion(), packaged: app.isPackaged, platform: process.platform,
      exts: OPENABLE_EXTS, settings: publicSettings(), recent: settings.recent(),
      // 문서 탭 뷰 = 그 문서, 홈 = 없음, AI 패널 = 그 탭 창의 AI 대상
      doc: hub ? docInfo(hub) : targetInfo(targetOf(tw)),
      editor: editorInfo(),
      smoke: SMOKE ? { chat: SMOKE_CHAT, settings: SMOKE_SETTINGS, newdoc: SMOKE_NEWDOC, editor: SMOKE_EDITOR, diff: !!SMOKE_DIFF, tool: SMOKE_TOOL, batch: SMOKE_BATCH } : null,
      opening: !!pending, view: viewRecs.get(wc)?.role ?? null,
    };
  });

  handle('doc:pick', async ({ hub, tw }) => {
    const all = OPENABLE_EXTS;
    const r = await dialog.showOpenDialog(winOf(hub), {
      title: '문서 열기', properties: ['openFile'],
      defaultPath: hub?.current ? dirname(hub.current.path) : app.getPath('documents'),
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
    return openPath(tw, r.filePaths[0]);
  });
  // 홈·문서 탭의 열기·끌어놓기·최근 목록 → 새 탭 (이미 열린 문서면 그 탭으로)
  handle('doc:open', ({ tw }, p) => openPath(tw, p));
  handle('doc:newInfo', ({ hub }, opts) => newDocInfo(hub, opts && typeof opts === 'object' ? opts : {}));
  handle('doc:newCheck', (_ctx, args) => newDocCheck(args && typeof args === 'object' ? args : {}));
  handle('doc:pickFolder', ({ hub }, start) => pickFolder(hub, start));
  handle('doc:new', (ctx, args) => newDocument(ctx, args));
  // 문서 탭의 '닫기' = 그 탭 닫기
  handle('doc:close', ({ hub, tw }) => {
    if (hub?.tabId && hub.tabId !== 'home' && tw?.tabs.get(hub.tabId)) closeDocTab(tw, hub.tabId);
    else if (hub) closeDocument(hub);
    return null;
  });
  handle('doc:preview', ({ hub }, opts = {}) => preview(hub, {
    mode: ['auto', 'exact', 'text'].includes(opts?.mode) ? opts.mode : undefined, ifChanged: opts?.ifChanged === true,
  }));
  handle('doc:restore', (ctx, args) => { const { doc, hub } = docFor(ctx, args); return restore(doc, args, hub); });
  handle('doc:versions', (ctx, args) => versions(docFor(ctx, args).doc));
  handle('doc:diff', (ctx, args) => diffVersion(docFor(ctx, args).doc, args));
  handle('pii:scan', ({ hub }) => piiScan(hub));
  handle('pii:mask', ({ hub }, args) => piiMask(hub, args));
  handle('batch:pick', ({ hub }, args) => batchPick(hub, args && typeof args === 'object' ? args : {}));
  handle('batch:pickOut', async ({ hub }) => {
    const r = await dialog.showOpenDialog(winOf(hub), { title: 'HWPX를 모을 폴더', properties: ['openDirectory', 'createDirectory'], defaultPath: hub?.current ? dirname(hub.current.path) : app.getPath('documents') });
    return r.canceled ? null : r.filePaths[0] ?? null;
  });
  handle('batch:scan', (_ctx, args) => batchScan(args && typeof args === 'object' ? args : {}));
  handle('batch:start', ({ wc }, args) => batchStart(wc, args && typeof args === 'object' ? args : {}));
  handle('batch:cancel', () => { if (batchRun) batchRun.cancelled = true; return null; });
  handle('batch:saveReport', () => batchSaveReport());
  handle('pdftools:pick', ({ hub }) => pdfPick(hub));
  handle('doc:print', ({ hub }) => printCurrent(hub));
  handle('pdftools:run', ({ hub }, args) => pdfRun(hub, args && typeof args === 'object' ? args : {}));
  handle('doc:openInApp', ({ hub }) => openInProgram(hub));
  handle('doc:openInWord', ({ hub }) => openPdfInWord(hub));
  // 문서 탭의 '직접 편집' → 그 탭을 같은 자리 편집 탭으로 (판정 7)
  handle('doc:edit', (ctx) => editDocument(ctx));
  handle('doc:savePdf', ({ hub }, args) => savePdf(hub, args && typeof args === 'object' ? { step: ['pick', 'make'].includes(args.step) ? args.step : undefined } : {}));
  handle('doc:showInFolder', ({ hub }) => { shell.showItemInFolder(needDoc(hub).path); return null; });
  handle('shell:showItem', (_ctx, p) => { shell.showItemInFolder(existingFile(p)); return null; });
  handle('shell:openBackups', async () => { mkdirSync(backups.dir, { recursive: true }); await shell.openPath(backups.dir); return null; });
  handle('shell:openExternal', async ({ wc }, url) => {
    const r = await openWebLink(url, wc);
    if (r === 'bad') throw fail('웹 주소만 열 수 있습니다');
    if (r === 'private') throw fail('이 PC나 내부망 주소는 열지 않아요');
    return null;
  });

  handle('recent:list', () => settings.recent());
  handle('recent:remove', (_ctx, p) => { if (typeof p !== 'string') throw fail('잘못된 경로입니다'); settings.removeRecent(p); return settings.recent(); });
  handle('recent:clear', () => { settings.clearRecent(); return settings.recent(); });
  handle('recent:pin', (_ctx, p, on) => {
    if (typeof p !== 'string' || !isAbsolute(p) || p.length > 4096) throw fail('잘못된 경로입니다');
    try { settings.pinRecent(resolve(p), on === true); } catch (err) { throw fail(err.message); }
    return settings.recent();
  });

  handle('settings:get', () => publicSettings());
  // 언어를 바꾼 뒤 다시 시작: 열린 편집 창은 저장 여부를 묻고 닫힌다 (취소하면 그대로). 열고 시작한 파일은 다시 열지 않는다
  handle('app:relaunch', () => {
    if (SMOKE) return null;
    const args = process.argv.slice(1).filter((a) => { try { return !(isAbsolute(a) && statSync(a).isFile()); } catch { return true; } });
    app.relaunch({ args });
    app.quit();
    return null;
  });
  handle('settings:set', (_ctx, patch) => setSettings(patch));

  handle('ai:providers', (_ctx, opts) => providers({ refresh: opts?.refresh === true }));
  // AI 패널의 요청은 그 탭 창의 AI 대상 기준 (판정 3)
  handle('ai:send', ({ tw }, args) => {
    const { announce: _announce, ...rest } = args && typeof args === 'object' ? args : {};
    return sendChat({ tw, path: targetOf(tw) }, rest);
  });
  handle('ai:cancel', (_ctx, chatId) => cancelChat(chatId));
  handle('ai:materials', ({ tw }) => listMaterials(tw));
  handle('ai:reset', ({ tw }) => { const p = targetOf(tw); if (p) sessions.delete(key(p)); pushTarget(tw); return null; });
  // 패널의 '편집 창 보기': AI 대상 문서의 편집 탭으로 (문서 탭이면 같은 자리 편집 탭으로 바꾼다)
  handle('ai:showEditor', (ctx) => {
    const tw = ctx.tw;
    const t = tw?.tabs.active;
    if (!t || !targetOf(tw)) throw fail('열린 문서가 없습니다');
    if (t.type === 'editor') { activateTab(tw, t.id); focusShell(tw); return { path: t.path, reused: true }; }
    return editDocument({ ...ctx, hub: hubOf(tw.docViews.get(t.id)?.webContents) });
  });

  // ── 탭 줄 (고르기·닫기·홈·AI 패널 접기) · AI 패널 너비 ──
  const tabIdOf = (tw, id) => {
    if (typeof id !== 'string' || !tw?.tabs.get(id)) throw fail('잘못된 요청입니다');
    return id;
  };
  handle('tabs:activate', ({ tw }, id) => { activateTab(tw, tabIdOf(tw, id)); return null; });
  handle('tabs:close', ({ tw }, id) => closeTab(tw, tabIdOf(tw, id)));
  handle('tabs:home', ({ tw }) => { activateTab(tw, 'home'); return null; });
  handle('tabs:togglePanel', ({ tw }) => {
    const r = tw.shell.setPanel({ collapsed: !tw.shell.setPanel().collapsed });
    try { settings.set({ chatCollapsed: r.collapsed }); } catch {}
    sendTabs(tw);
    if (!r.collapsed) tw.panel?.webContents.focus();
    return { collapsed: r.collapsed };
  });
  onIpc('tabs:ready', (_e, { tw }) => { if (tw) sendTabs(tw); });
  // AI 패널 왼쪽 가장자리로 너비 조절 (바로 배치하고, 설정 저장은 200ms 모아서)
  handle('tabs:panelWidth', ({ tw }, w) => {
    if (!Number.isFinite(w)) throw fail('잘못된 너비입니다');
    const r = tw.shell.setPanel({ width: w });
    clearTimeout(tw.widthTimer);
    tw.widthTimer = setTimeout(() => { try { settings.set({ chatWidth: r.width }); } catch {} }, 200);
    return { width: r.width };
  });

  // ── 홈·패널 사이 이름 붙은 호출 중계 (VIEW_ROUTES, 판정 4) ──
  // 표에 있는 이름이고 보낸 쪽 종류가 받을 종류와 다를 때만 같은 탭 창(scope window)·모든 탭 창(scope all)의 그 뷰로 넘긴다
  onIpc('view:call', (_e, { kinds, tw }, msg) => {
    const name = typeof msg?.name === 'string' ? msg.name : '';
    const route = Object.hasOwn(VIEW_ROUTES, name) ? VIEW_ROUTES[name] : null;
    if (!route || !tw || kinds?.includes(route.to)) return;
    const args = Array.isArray(msg.args) ? msg.args.slice(0, 8) : [];
    const payload = { name, args };
    for (const t of route.scope === 'all' ? [...tabWindows.values()] : [tw]) {
      if (route.to === 'panel') { sendTo(t.panel?.webContents, 'view:called', payload); continue; }
      if (route.byPath) {
        // 그 문서를 보고 있는 홈 종류 뷰(문서 탭)만
        const p = typeof args[0]?.path === 'string' ? args[0].path : null;
        if (!p) continue;
        for (const h of allHubs()) if (h.tw === t && h.current && samePath(h.current.path, p)) sendTo(h.wc, 'view:called', payload);
        continue;
      }
      if (route.focusHome) activateTab(t, 'home');
      sendTo(t.home?.webContents, 'view:called', payload);
    }
  });

  handle('integrations:status', () => integrationStatus());
  handle('integrations:set', (_ctx, args) => setIntegration(args));
  handle('hancom:status', async () => {
    const apps = await com.probe({ refresh: true }).catch(() => ({}));
    let status = null; let error = null;
    try { status = await hancomStatus(); } catch (err) { error = err.message; }
    return { status, error, apps: { hancom: !!apps.hancom, word: !!apps.word, powerpoint: !!apps.powerpoint, excel: !!apps.excel, hancomSecurityModule: !!apps.hancomSecurityModule } };
  });
  handle('hancom:setup', (_ctx, args) => hancomSetup(args));
  handle('assoc:status', () => assocStatus());
  handle('assoc:openSettings', () => assocOpenSettings());
  handle('hancom:remove', async (_ctx, args) => {
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
    async function smokePdf(wc) {
      const ask = (code) => wc.executeJavaScript(code).catch(() => null);
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
    async function smokeEditor(hub) {
      const png = `${SMOKE.replace(/\.png$/i, '')}-editor.png`;
      if (!editors) return { ok: false, error: `편집 창 기능 없음: ${editorLoadError}` };
      if (!hub?.current) return { ok: false, error: '열린 문서가 없습니다' };
      const before = new Set(BrowserWindow.getAllWindows());
      // 묶음이 없어도 창을 연다 (편집 창의 실패 안내를 찍으려고)
      try { await editors.openEditor(hub.current.path); } catch (err) { return { ok: false, error: err.message }; }
      pushEditorState();
      let ew = null;
      for (let i = 0; i < 100 && !ew; i++) {
        ew = BrowserWindow.getAllWindows().find((w) => !before.has(w) && !w.isDestroyed()) ?? null;
        if (!ew) await sleep(100);
      }
      if (!ew) return { ok: false, error: '편집 창이 뜨지 않았습니다' };
      const kind = editorKind(hub.current.path);
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
      // 탭 창을 앞으로
      const front = lastTw()?.shell.win;
      if (front && !front.isDestroyed()) { front.moveTop(); front.focus(); }
      const ok = state === 'saved' && (kind !== 'office' || phase === 'ready');
      return { ok, kind, state, phase, fonts, about, aboutInfo, reason: reason || null, title: ew.getTitle(), png, size: img.getSize() };
    }

    onIpc('smoke:ready', async (e, { hub }, info = {}) => {
      try {
        // 열면 바로 편집 창 (OFFICEKIT_SMOKE_AUTOEDIT=1): 화면이 편집 창을 띄웠는지 (30초까지 기다림)
        let autoEditor = null;
        if (process.env.OFFICEKIT_SMOKE_AUTOEDIT === '1' && hub?.current) {
          for (const end = Date.now() + 30_000; Date.now() < end && !editors?.isOpen(hub.current.path);) await sleep(200);
          autoEditor = { open: !!editors?.isOpen(hub.current.path), windows: BrowserWindow.getAllWindows().length };
        }
        const editor = SMOKE_EDITOR && info.ok !== false ? await smokeEditor(hub) : null;
        const pdf = SMOKE_PDF && info.ok !== false ? await smokePdf(e.sender) : null;
        await sleep(info.kind === 'pdf' ? 3000 : 900);
        // 신호를 보낸 홈 종류 뷰(홈·문서 탭)를 찍는다 (탭 창 전체 찍기는 스모크 작업에서)
        const img = await capture({ webContents: e.sender });
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
  async function openWebLink(raw, from = null) {
    let u;
    try { u = new URL(String(raw)); } catch { return 'bad'; }
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return 'bad';
    const { isPrivateHost } = await import('./editor-window.mjs');
    if (isPrivateHost(u.hostname)) {
      sendTo(from ?? lastTw()?.home?.webContents, 'toast', { type: 'warn', message: `이 PC나 내부망 주소는 열지 않아요: ${u.host}` });
      return 'private';
    }
    if (!TRUSTED_LINK_HOSTS.has(u.hostname.toLowerCase())) {
      if (askingLink) return 'canceled';
      askingLink = true;
      try {
        const shown = u.href.length > 300 ? `${u.href.slice(0, 300)}…` : u.href;
        const { response } = await dialog.showMessageBox(winOf({ tw: viewRecs.get(from)?.tw }), {
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


  /** 최근 탭 창에서 보이는 뷰의 webContents (개발 메뉴용) */
  function activeWc() {
    const tw = lastTw();
    const view = tw ? viewOfTab(tw, tw.tabs.active) : null;
    return view?.webContents ?? null;
  }

  /** 메뉴 명령 → 최근 탭 창: 문서 탭이면 그 뷰, 아니면 홈 탭으로 가서 홈 뷰 (문서 명령은 문서 탭에서만) */
  function sendMenu(cmd) {
    const tw = lastTw();
    if (!tw) return;
    const t = tw.tabs.active;
    if (t?.type === 'doc') { sendTo(tw.docViews.get(t.id)?.webContents, 'menu', cmd); return; }
    if (['edit', 'refresh', 'print', 'pii'].includes(cmd)) return;
    if (t?.type !== 'home') activateTab(tw, 'home');
    sendTo(tw.home?.webContents, 'menu', cmd);
  }

  function buildMenu() {
    const template = [
      { label: '파일', submenu: [
        // Ctrl+N은 새 탭 창 (판정 6) — 새 문서는 홈 탭 단추·이 메뉴로
        { label: '새 문서…', click: () => sendMenu('new') },
        { label: '파일 열기…', accelerator: 'CmdOrCtrl+O', click: () => sendMenu('open') },
        { label: '직접 편집', accelerator: 'CmdOrCtrl+E', click: () => sendMenu('edit') },
        { label: '새로고침', accelerator: 'F5', click: () => sendMenu('refresh') },
        { label: '인쇄…', accelerator: 'CmdOrCtrl+P', click: () => sendMenu('print') },
        { type: 'separator' },
        { label: '개인정보 찾아 가리기…', click: () => sendMenu('pii') },
        { label: '한글 → HWPX 일괄 변환…', click: () => sendMenu('batch') },
        { label: 'PDF 합치기 · 나누기 · 돌리기…', click: () => sendMenu('pdftools') },
        { type: 'separator' },
        { label: '설정', accelerator: 'CmdOrCtrl+,', click: () => sendMenu('settings') },
        { type: 'separator' },
        { label: '끝내기', role: 'quit' },
      ] },
    ];
    if (!app.isPackaged) {
      template.push({ label: '개발', submenu: [
        { label: '개발자 도구', accelerator: 'CmdOrCtrl+Shift+I', click: () => activeWc()?.toggleDevTools() },
        { label: '화면 다시 불러오기', accelerator: 'CmdOrCtrl+Shift+R', click: () => activeWc()?.reloadIgnoringCache() },
      ] });
    }
    Menu.setApplicationMenu(Menu.buildFromTemplate(localizeMenu(template, i18n.t)));
  }

  // ── 시작 ─────────────────────────────────────────────

  const argvFile = fileFromArgv(process.argv);
  // 시작은 탭 창 하나. 명령줄 파일은 홈 뷰가 준비되면 그 파일 탭을 연다 (탐색기 시작·평소 시작 복원은 탭 열기·복원 작업에서)
  let earlyFile = null; // 탭 창을 만들기 전에 받은 탐색기 파일

  app.on('before-quit', () => { appQuitting = true; });

  app.on('second-instance', (_e, argv, cwd) => {
    const f = fileFromArgv(argv, cwd);
    const tw = lastTw();
    if (!tw) { if (f) earlyFile = f; return; }
    focusShell(tw);
    if (!f) return;
    // 탐색기에서 연 문서 → 그 탭 창의 새 탭 (편집할 수 있으면 편집 탭, 이미 열린 문서면 그 탭)
    if (wantsDocWindow(f) && !SMOKE && process.env.OFFICEKIT_DOC_WINDOW !== '0') {
      openAsDocument(f, tw).catch((err) => toastIn(tw, err.message));
      return;
    }
    openPath(tw, f).catch((err) => toastIn(tw, err.message));
  });

  app.on('window-all-closed', () => app.quit());
  // 편집 탭이 열려 있으면 '끝내기'가 취소될 수 있으므로(저장 확인 → 취소) 정리는 정말 끝날 때(will-quit) 한다
  app.on('will-quit', () => {
    for (const hub of allHubs()) stopWatch(hub);
    for (const c of chats.values()) { try { c.handle.cancel(); } catch {} }
    // 창이 다 닫힌 뒤 미리보기 PDF(문서 사본)를 지운다. 못 지운 것은 다음 시작 때 지운다
    cleanPreviews(null);
    cleanPrintCopies({ all: true });
  });

  await app.whenReady();
  i18n = createI18n(resolveLang(settings.get('lang'), app.getPreferredSystemLanguages?.() ?? [app.getLocale?.() ?? 'ko']));
  // 화면(주 창·편집 창의 preload)이 처음 그리기 전에 받는다 (sendSync) — 번역 사전은 앱에 들어 있는 공개 글이다
  onIpc('i18n:get', (e) => { e.returnValue = { lang: i18n.lang, setting: settings.get('lang'), dict: i18n.dict }; });
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
  // 강제 종료·충돌로 남은 바꿔 연 임시 사본 (아직 편집 창이 없으므로 모두 지운다)
  try { cleanConvertedCopies?.(userData); } catch (err) { console.error('[editor] 바꿔 연 임시 사본 정리 실패:', err?.message ?? err); }
  // 강제 종료로 남은 AI 대화의 PID 기록은 지우기만 한다 (그 PID는 이제 다른 프로그램일 수 있어 끝내지 않는다)
  try { rmSync(join(userData, COM_PIDS), { recursive: true, force: true }); } catch (err) { console.error('[ai] 남은 PID 기록 정리 실패:', err?.message ?? err); }
  buildMenu();
  const savedWin = SMOKE ? null : settings.get('window');
  const firstTw = createTabWindowFor({ bounds: initialBounds(), maximized: !!savedWin?.maximized });
  const firstHome = homeHubOf(firstTw);
  if (firstHome) firstHome.pendingFile = argvFile ?? earlyFile;
}
