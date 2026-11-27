// 앱(main.mjs)과 편집 창(한글 · 문서·표·슬라이드) 잇기 시험.
//  1) 가짜 electron 모듈로 main.mjs를 그대로 띄워 IPC 흐름을 확인한다 (창은 가짜, 편집기 서버·파일·백업은 진짜)
//  2) Office 편집기 설정 보관(store.mjs), '정보 · 라이선스' 글(legal-ko.json ↔ docs/office-editor-legal-ko.md)
//  3) OFFICEKIT_EDITOR_SMOKE=1 이고 vendor/ 묶음이 있으면 실제 Electron으로 두 창을 찍는다 (창이 잠깐 뜬다)
//     한글(hwpx) + 문서·표·슬라이드(docx·xlsx·pptx). OFFICEKIT_SMOKE_OUT=<폴더>면 찍은 그림을 거기에 남긴다
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { createRequire } from 'node:module';
import {
  copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, utimesSync, writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBlankDocument, createDocument } from '../src/core/create.mjs';
import { readDocument, Hwp, hwpPdfStatus } from '../src/core/index.mjs';
import { createSettings, OFFICE_STATE_MAX_KEYS } from '../src/app/store.mjs';
import { draftPathFor } from '../src/app/editor-window.mjs';
import { startEditorServer } from '../src/app/editor-server.mjs';
import { NOTICE_FILES, aboutModel, legalSections } from '../src/app/office-editor/host.js';
import { makePdf } from './helpers/fixtures.mjs';
import { makeFakeElectron, installFakeElectron } from './helpers/fake-electron.mjs';

const readDocumentText = async (p) => (await readDocument(p, { engine: 'file' })).text ?? '';

const here = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(here, '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (p) => createHash('sha1').update(readFileSync(p)).digest('hex');
const shaBytes = (b) => createHash('sha1').update(b).digest('hex');
const DAY = 24 * 60 * 60 * 1000;

async function until(fn, ms = 8000, step = 40) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) return v;
    await sleep(step);
  }
}

const put = (p, data = 'x') => { mkdirSync(dirname(p), { recursive: true }); writeFileSync(p, data); };

/** 시험용 한글 편집기 묶음 (vendor/rhwp-studio, vendor/rhwp-editor-sdk 흉내) */
function makeVendor(dir) {
  put(join(dir, 'rhwp-studio', 'index.html'), '<!doctype html><title>studio</title>');
  put(join(dir, 'rhwp-studio', 'assets', 'rhwp_bg-test.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(dir, 'rhwp-editor-sdk', f), 'export {};');
}

const FAKE_VERSION = {
  name: 'office-editor', officeKitVersion: '0.0.0-test', editorVersion: '9.4.0.129',
  onlyoffice: { version: '9.4.0', build: '9.4.0.129', sourceTag: 'v9.4.0.129' },
  filex: { tag: 'v0.3.1', commit: '57dda3dae99af78eb26b3c64b69f1e3f711f0757' },
  files: [],
};

/** 시험용 문서·표·슬라이드 편집기 묶음 (vendor/office-editor 흉내: 있는지 확인하는 파일만) */
function makeOfficeVendor(dir) {
  const o = join(dir, 'office-editor');
  put(join(o, 'index.html'), '<!doctype html><title>app</title>');
  put(join(o, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(o, 'x2t', 'x2t.js'), '// x2t');
  put(join(o, 'editor', 'LICENSE.txt'), 'GNU AFFERO GENERAL PUBLIC LICENSE');
  put(join(o, 'VERSION.json'), JSON.stringify(FAKE_VERSION));
}

const FAKE_AI = `
export async function detectProviders() { return [{ id: 'claude', name: 'Claude (시험)', available: true, version: '1', hint: '' }]; }
export function defaultWorkspace() { return process.env.OFFICEKIT_TEST_WS; }
export function aiMcpEnv(target) { globalThis.__okTargets = [...(globalThis.__okTargets ?? []), target]; return {}; }
export function startChat({ prompt, onEvent, filePath, materials, outDir }) {
  globalThis.__okPrompts = [...(globalThis.__okPrompts ?? []), prompt];
  globalThis.__okChats = [...(globalThis.__okChats ?? []), { prompt, filePath, materials, outDir }];
  // '[hold]'가 든 요청은 시험이 __okRelease()를 부를 때까지 끝나지 않는다 (진행 중인 대화 흉내)
  const hold = String(prompt).includes('[hold]') ? new Promise((r) => { globalThis.__okRelease = r; }) : Promise.resolve();
  const done = hold.then(() => { onEvent({ type: 'done', sessionId: 's1', text: '네' }); return { sessionId: 's1', text: '네' }; });
  return { cancel() { globalThis.__okRelease?.(); }, done };
}
`;
// 글꼴 모듈 대신 (이 PC 글꼴을 읽지 않는다): 묶음 글꼴로
const FAKE_FONTS = `
export async function ensureFonts() { globalThis.__okFontCalls = (globalThis.__okFontCalls ?? 0) + 1; return { allFontsJsPath: null, thumbnailsDir: null, fontFileByIndex: () => null }; }
`;

test('main.mjs ↔ 편집 창 (가짜 electron): 한글 · 문서·표·슬라이드', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-int-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  const docPath = join(docs, '보고서.hwpx');
  await createBlankDocument(docPath, { title: '처음 내용' });
  const docxPath = join(docs, '메모.docx');
  await createDocument(docxPath, '# 회의 메모\n\n처음 쓴 내용입니다.\n\n- 안건 하나\n- 안건 둘');
  const csvPath = join(docs, '목록.csv');
  writeFileSync(csvPath, '이름,수량\n사과,3\n');
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);
  const fontsModule = join(base, 'fake-fonts.mjs');
  writeFileSync(fontsModule, FAKE_FONTS);

  // 지난 실행에서 남은 자동 저장 사본 (시작할 때 정리: 30일 넘은 것·쓰다 남은 임시 파일은 지우고 나머지는 둔다)
  const drafts = join(ud, 'office-drafts');
  const oldDraft = join(drafts, `${'a'.repeat(40)}.docx`);
  const freshDraft = join(drafts, `${'b'.repeat(40)}.xlsx`);
  const oldTmp = join(drafts, `.${'c'.repeat(40)}.docx.1234.abcd.tmp`);
  const freshTmp = join(drafts, `.${'d'.repeat(40)}.pptx.1234.abcd.tmp`);
  for (const f of [oldDraft, freshDraft, oldTmp, freshTmp]) put(f, 'PK');
  utimesSync(oldDraft, new Date(Date.now() - 40 * DAY), new Date(Date.now() - 40 * DAY));
  utimesSync(oldTmp, new Date(Date.now() - 2 * 60 * 60 * 1000), new Date(Date.now() - 2 * 60 * 60 * 1000));

  const saved = { argv: process.argv, env: { ...process.env } };
  Object.assign(process.env, {
    OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_VENDOR: vendor, OFFICEKIT_AI_MODULE: aiModule, OFFICEKIT_TEST_WS: join(base, 'ws'),
    OFFICEKIT_OFFICE_FONTS_MODULE: fontsModule,
    // 이 시험은 주 창에서 문서를 여는 흐름을 본다 (탐색기에서 연 문서의 '문서 창 방식'은 끈다)
    OFFICEKIT_DOC_WINDOW: '0',
  });
  for (const k of ['OFFICEKIT_SMOKE', 'OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_TEST']) delete process.env[k];
  process.argv = [process.argv[0], docPath];
  // 문서를 열면 문서 탭(미리보기)으로: 예전 시험의 '열기 → 직접 편집' 흐름을 그대로 따른다 (가짜 화면은 autoEdit로 편집 창을 열지 않았다)
  // 지난 세션(평소 시작으로 끝낸 것): 탐색기 시작 세션은 이것을 덮지 않고, 홈의 '지난 탭 다시 열기'가 연다 (D4·D11)
  const lastPdf = join(docs, '지난 세션.pdf');
  writeFileSync(lastPdf, makePdf(1));
  const lastSession = {
    v: 1, at: 1000, explorer: false, finalized: true,
    windows: [{ bounds: { x: 300, y: 200, width: 1200, height: 800 }, maximized: false, active: 2, tabs: [{ type: 'doc', path: docPath }, { type: 'doc', path: lastPdf }] }],
  };
  put(join(ud, 'settings.json'), JSON.stringify({ openAction: 'preview', lastTabSession: lastSession }));

  const fake = makeFakeElectron({ userData: ud, documents: docs, commandAnswer: async () => false });
  installFakeElectron(fake);

  t.after(async () => {
    fake.app.emit('will-quit');
    for (const w of fake.BrowserWindow.getAllWindows()) w.destroy();
    await sleep(100);
    process.argv = saved.argv;
    for (const k of Object.keys(process.env)) if (!(k in saved.env)) delete process.env[k];
    Object.assign(process.env, saved.env);
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  });

  const mainMod = await import(pathToFileURL(join(ROOT, 'src', 'app', 'main.mjs')).href);
  assert.ok(await until(() => fake.windows.length >= 1, 20_000), '탭 창이 만들어져야 한다');
  const tw0 = fake.windows[0];
  /** 주소가 ?view=<kind>인 뷰들 (홈·AI 패널·문서 탭) */
  const viewsOf = (kind) => fake.views.filter((v) => v.webContents.mainFrame.url.endsWith(`?view=${kind}`));
  const homeWc = viewsOf('home')[0]?.webContents;
  const panelWc = viewsOf('panel')[0]?.webContents;
  const tabstripWc = fake.views.find((v) => /tabstrip\.html$/.test(v.webContents.mainFrame.url))?.webContents;
  assert.ok(homeWc && panelWc && tabstripWc, '홈·AI 패널·탭 줄 뷰');
  /** 문서 탭 뷰: ?view=doc 뷰 중 마지막 doc:opened 경로가 p인 (살아 있는) wc */
  const docWc = (p) => viewsOf('doc').map((v) => v.webContents)
    .filter((wc) => !wc.isDestroyed() && wc.sent.filter((s) => s.channel === 'doc:opened').at(-1)?.data?.path === p).at(-1) ?? null;
  /** who: 'home' | 'panel' | 'tabstrip' | 문서 경로(그 문서 탭 뷰) */
  const callAs = (who, channel, ...args) => {
    const wc = who === 'home' ? homeWc : who === 'panel' ? panelWc : who === 'tabstrip' ? tabstripWc : docWc(who);
    if (!wc) throw new Error(`${who}: ${channel}을 보낼 뷰가 없다`);
    return fake.invokeFrom(wc, channel, ...args);
  };
  // 지금 문서 (마지막으로 연 문서). 예전 주 창 call → 채널마다 맞는 뷰로:
  //   ai:* → AI 패널, 버전·바뀐 내용·되돌리기 → 그 문서 탭 뷰(없으면 = 편집 탭이면 AI 패널)에 경로를 붙여, 그 밖의 doc:*·pii:* → 문서 탭 뷰, 나머지 → 홈
  let cur = null;
  const PATH_CH = new Set(['doc:versions', 'doc:diff', 'doc:restore']);
  const HOME_DOC_CH = new Set(['doc:open', 'doc:pick', 'doc:new', 'doc:newInfo', 'doc:newCheck', 'doc:pickFolder']);
  async function call(channel, ...args) {
    if (channel.startsWith('ai:')) return callAs('panel', channel, ...args);
    if (PATH_CH.has(channel)) return callAs(docWc(cur) ? cur : 'panel', channel, { ...(args[0] ?? {}), path: cur });
    if ((channel.startsWith('doc:') || channel.startsWith('pii:')) && !HOME_DOC_CH.has(channel)) return callAs(cur, channel, ...args);
    const r = await callAs('home', channel, ...args);
    if (channel === 'doc:open' && r.ok) cur = r.value.path;
    if (channel === 'doc:new' && r.ok && r.value.doc?.path) cur = r.value.doc.path;
    return r;
  }
  /** 홈·문서 탭·AI 패널 뷰가 받은 것 (뷰 차례로 이어 붙임) */
  const allSent = () => fake.views.filter((v) => /\?view=(home|doc|panel)$/.test(v.webContents.mainFrame.url)).flatMap((v) => v.webContents.sent);
  const lastSent = (channel, pred = () => true) => allSent().filter((s) => s.channel === channel && pred(s.data)).at(-1)?.data;
  /** AI 패널이 받은 마지막 AI 대상 문서 (없으면 null) */
  const lastTarget = () => panelWc.sent.filter((s) => s.channel === 'ai:target').at(-1)?.data?.doc ?? null;
  /** 편집 탭 상태 { path, open, dirty } — 편집 탭 문서의 상태는 AI 패널의 ai:target으로 온다 (예전 주 창 doc:editor) */
  const lastEditor = (pred = () => true) => panelWc.sent.filter((s) => s.channel === 'ai:target' && s.data?.doc)
    .map((s) => ({ path: s.data.doc.path, ...s.data.doc.editor })).filter(pred).at(-1);
  const sentTo = (w, channel) => w.webContents.sent.filter((s) => s.channel === channel);
  /** 탭 줄이 받은 마지막 탭 목록 { tabs, active, panelCollapsed } */
  const tabList = () => tabstripWc.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [], active: null };
  const tabOf = (p) => tabList().tabs.find((x) => x.path === p) ?? null;
  // 탭 창 제목은 활성 탭 것이다 → 제목이 바뀔 때 보이던 편집 뷰의 제목으로 기억해 둔다
  const titleOfView = new Map();
  const setTitle0 = tw0.setTitle.bind(tw0);
  tw0.setTitle = (title) => {
    const v = tw0.contentView.children.find((c) => c.opts?.webPreferences?.partition && c.visible && !c.webContents.isDestroyed());
    if (v) titleOfView.set(v, title);
    setTitle0(title);
  };
  /** 제목이 re에 맞는 (마지막) 편집 탭 → { win: 탭 창, view, wc } | null */
  const editorOf = (re) => {
    const hit = [...titleOfView].filter(([v, title]) => !v.webContents.isDestroyed() && re.test(title)).at(-1);
    return hit ? { win: tw0, view: hit[0], wc: hit[0].webContents } : null;
  };
  /** 살아 있는 편집 뷰들 */
  const editorViews = () => fake.views.filter((v) => v.opts?.webPreferences?.partition && !v.webContents.isDestroyed());
  assert.equal(fake.app.exitCode, undefined, fake.errorBox);

  await t.test('시작: 이 PC 밖 이름은 풀지 않는다 (편집기 틀의 dns-prefetch로 내용이 새지 않게), 편집기 서버 주소는 뺀다', () => {
    const rules = fake.app.commandLine.switches.get('host-resolver-rules');
    assert.equal(rules, 'MAP * ~NOTFOUND, EXCLUDE localhost, EXCLUDE 127.0.0.1, EXCLUDE 127.0.0.2, EXCLUDE 127.0.0.3');
  });

  await t.test('시작: 오래된 자동 저장 사본과 쓰다 남은 임시 파일을 지운다 (최근 것은 둔다)', () => {
    assert.equal(existsSync(oldDraft), false, '30일 넘은 사본');
    assert.equal(existsSync(oldTmp), false, '한 시간 넘은 임시 파일');
    assert.equal(existsSync(freshDraft), true);
    assert.equal(existsSync(freshTmp), true, '쓰는 중일 수 있는 임시 파일은 남긴다');
  });

  await t.test('메뉴: 직접 편집(Ctrl+E)은 종류를 가리지 않는 이름', () => {
    const items = fake.menu?.[0]?.submenu ?? [];
    const edit = items.find((i) => i.accelerator === 'CmdOrCtrl+E');
    assert.equal(edit?.label, '직접 편집');
    assert.doesNotMatch(JSON.stringify(fake.menu.map((m) => m.label)), /ONLYOFFICE/i);
  });

  const queuedPdf = join(docs, '줄 시험.pdf');
  await t.test('second-instance: 준비 전 요청은 줄에 섰다가 차례로, 같은 문서·대소문자 다른 경로는 그 탭으로', async () => {
    writeFileSync(queuedPdf, makePdf(1));
    // 첫 탭 창 홈 뷰가 준비(app:init)되기 전에 탐색기에서 두 번 더 열었다: 다른 PDF, 시작 문서(대소문자만 다른 경로)
    fake.app.emit('second-instance', {}, [process.execPath, queuedPdf], docs);
    fake.app.emit('second-instance', {}, [process.execPath, process.platform === 'win32' ? docPath.toUpperCase() : docPath], docs);
    await sleep(150);
    assert.deepEqual(tabList().tabs.map((x) => x.type), ['home'], '준비 전에는 아무 탭도 열지 않는다');
  });

  let init;
  await t.test('app:init: 종류별 편집기 묶음 상태를 알려 준다 (없으면 한국어 이유)', async () => {
    init = await call('app:init');
    assert.equal(init.ok, true, init.error);
    const ed = init.value.editor;
    assert.deepEqual(ed.exts, ['hwp', 'hwpx', 'hwt', 'docx', 'xlsx', 'pptx']);
    assert.deepEqual(ed.kinds.hwp.exts, ['hwp', 'hwpx', 'hwt']);
    assert.deepEqual(ed.kinds.office.exts, ['docx', 'xlsx', 'pptx']);
    assert.equal(ed.kinds.hwp.available, false);
    assert.match(ed.kinds.hwp.reason, /한글 편집기/);
    assert.equal(ed.kinds.office.available, false);
    assert.match(ed.kinds.office.reason, /문서 편집기 파일이 없습니다 \(vendor\/office-editor\)/);
    assert.match(ed.kinds.office.reason, /npm run office-editor:build/, '개발 실행이면 만드는 법');
    assert.equal(init.value.opening, true);
  });

  await t.test('문서 정보에 편집 창 상태가 들어 있다', async () => {
    const opened = await until(() => lastSent('doc:opened', (d) => d?.path === docPath));
    assert.ok(opened, '인수로 받은 문서를 연다');
    assert.deepEqual(opened.editor, { open: false, dirty: false });
    // D4: 탐색기로 시작 → 홈 + 그 파일 탭. 준비 전에 줄에 선 열기는 그 뒤에 차례로 (같은 문서는 새 탭 없이 그 탭으로)
    assert.ok(await until(() => tabList().tabs.length >= 3 && tabList().active === tabOf(docPath)?.id), JSON.stringify(tabList()));
    await sleep(100);
    assert.deepEqual(tabList().tabs.map((x) => (x.type === 'home' ? '홈' : basename(x.path))), ['홈', '보고서.hwpx', '줄 시험.pdf']);
    // 줄 시험 PDF 탭은 닫아 둔다 (아래 시험은 홈 + 시작 문서에서 시작)
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(queuedPdf).id)).ok, true);
    assert.deepEqual(tabList().tabs.map((x) => (x.type === 'home' ? '홈' : basename(x.path))), ['홈', '보고서.hwpx']);
    assert.equal(tabList().active, tabOf(docPath).id);
  });

  await t.test('doc:edit: 편집 창 형식이 아니면(csv) 거절하고 프로그램에서 열기를 안내', async () => {
    assert.equal((await call('doc:open', csvPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'UNSUPPORTED');
    assert.match(r.error, /\.csv/);
    assert.match(r.error, /프로그램에서 열기/);
    assert.equal(editorViews().length, 0, '편집 창을 띄우지 않는다');
    // csv 문서 탭은 닫아 둔다 (아래 재료 시험이 '열리지 않은 파일'로 쓴다)
    await callAs('tabstrip', 'tabs:close', tabOf(csvPath).id);
  });

  await t.test('doc:edit: Word 문서도 직접 편집 대상 — 묶음이 없으면 EDITOR_MISSING (창을 띄우지 않음)', async () => {
    assert.equal((await call('doc:open', docxPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'EDITOR_MISSING');
    assert.match(r.error, /문서 편집기 파일이 없습니다/);
    assert.equal(editorViews().length, 0);
  });

  await t.test('doc:edit: 한글 편집기 묶음이 없으면 EDITOR_MISSING (창을 띄우지 않음)', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'EDITOR_MISSING');
    assert.match(r.error, /한글 편집기 파일이 없습니다/);
    assert.equal(editorViews().length, 0);
  });

  makeVendor(vendor);
  makeOfficeVendor(vendor);

  await t.test('app:init: 묶음을 만들면 두 종류 모두 쓸 수 있다', async () => {
    const r = await call('app:init');
    assert.equal(r.value.editor.kinds.hwp.available, true, r.value.editor.kinds.hwp.reason);
    assert.equal(r.value.editor.kinds.office.available, true, r.value.editor.kinds.office.reason);
    assert.equal(r.value.editor.kinds.office.reason, null);
  });

  let editorWin; let hostCall;
  await t.test('doc:edit: 묶음이 있으면 편집 창을 열고, 다시 누르면 같은 창', async () => {
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.path, docPath);
    assert.equal(r.value.kind, 'hwp');
    assert.equal(r.value.reused, false);
    assert.equal(r.value.notice, null);
    editorWin = editorOf(/보고서\.hwpx - 한글 편집/);
    assert.ok(editorWin, '편집 창');
    assert.match(editorWin.win.getTitle(), /보고서\.hwpx - 한글 편집 \(Office Kit\)/);
    assert.match(editorWin.wc.mainFrame.url, /^http:\/\/127\.0\.0\.1:\d+\/host\/host\.html$/);
    assert.equal(editorWin.view.opts.webPreferences.partition, 'officekit-editor');
    hostCall = (channel, ...args) => fake.invokeFrom(editorWin.wc, channel, ...args);
    const again = await call('doc:open', docPath);
    assert.equal(again.value.reused, true);
    assert.equal(editorViews().length, 1);
    const st = await until(() => lastEditor((d) => d.open));
    assert.deepEqual(st, { path: docPath, open: true, dirty: false });
    // 편집 창 밖으로 나가는 요청은 막는다
    const ses = fake.sessions.get('officekit-editor');
    const blocked = await new Promise((r) => ses.handlers.beforeRequest({ url: 'https://cdn.jsdelivr.net/x.woff2' }, r));
    assert.equal(blocked.cancel, true);
    // HTTP 밖의 연결·맞춤법 사전 내려받기도 인터넷에 나가지 못하게 (죽은 프록시·사전 주소, WebRTC는 프록시 안 거친 UDP 금지)
    // 바로 가는 것은 host·studio 서버의 주소:포트뿐 (<-loopback>: 이 PC의 다른 포트는 죽은 프록시로)
    assert.match(ses.proxy.proxyBypassRules, /^<-loopback>;127\.0\.0\.1:\d+;127\.0\.0\.[12]:\d+$/);
    assert.deepEqual({ ...ses.proxy, proxyBypassRules: null }, { mode: 'fixed_servers', proxyRules: 'socks5://127.0.0.1:9', proxyBypassRules: null });
    assert.equal(ses.dictionaryUrl, 'http://127.0.0.1:9/');
    assert.equal(fake.defaultDictionaryUrl, 'http://127.0.0.1:9/');
    assert.equal(editorWin.wc.webrtcPolicy, 'disable_non_proxied_udp');
  });

  await t.test('편집 창 새 창·링크: 이 PC·사설망 주소는 열지 않고, 바깥 주소는 물어본 뒤 기본 브라우저로', async () => {
    const open = editorWin.wc.openHandler;
    const d0 = fake.dialogs.length;
    for (const url of ['http://127.0.0.2:5555/studio/print.html', 'http://127.0.0.1:1/x', 'http://localhost/', 'http://192.168.0.10/',
      'http://10.1.2.3/', 'http://[::1]/', 'http://[::ffff:127.0.0.1]/', 'http://2130706433/', 'file:///C:/Windows/win.ini', 'javascript:alert(1)',
      'https://user:pw@example.com/']) {
      assert.deepEqual(open({ url }), { action: 'deny' }, url);
    }
    await sleep(50);
    assert.equal(fake.dialogs.length, d0, '내부 주소는 묻지도 않는다');
    assert.deepEqual(fake.external, []);
    // 바깥 주소: 취소하면 열지 않고, 열기를 고르면 연다
    fake.dialogAnswers.push(1);
    assert.deepEqual(open({ url: 'https://example.com/a?b=1' }), { action: 'deny' });
    await until(() => fake.dialogs.length > d0);
    await sleep(20);
    assert.match(fake.dialogs.at(-1).message, /기본 브라우저로 열까요/);
    assert.equal(fake.dialogs.at(-1).detail, 'https://example.com/a?b=1');
    assert.deepEqual(fake.external, []);
    fake.dialogAnswers.push(0);
    open({ url: 'https://example.com/a?b=1' });
    assert.ok(await until(() => fake.external.length === 1));
    assert.deepEqual(fake.external, ['https://example.com/a?b=1']);
    // studio의 문서 링크(빈 창 → 주소 바꾸기): 빈 창만 숨겨서, 스크립트 없이
    const blank = open({ url: 'about:blank' });
    assert.equal(blank.action, 'allow');
    assert.equal(blank.overrideBrowserWindowOptions.show, false);
    assert.equal(blank.overrideBrowserWindowOptions.webPreferences.javascript, false);
    assert.equal(blank.overrideBrowserWindowOptions.webPreferences.partition, 'officekit-editor');
    fake.external.length = 0;
  });

  await t.test('편집 창이 고친 상태면 AI 요청을 막고, editorOk면 보낸다', async () => {
    const doc = await hostCall('editor:get-document');
    assert.equal(doc.ok, true, doc.error);
    assert.equal(doc.value.name, '보고서.hwpx');
    assert.equal(doc.value.kind, 'hwp');
    fake.emit('editor:dirty', editorWin.wc, true);
    const st = await until(() => lastEditor((d) => d.dirty));
    assert.deepEqual(st, { path: docPath, open: true, dirty: true });
    assert.ok(editorWin.win.getTitle().startsWith('* '));
    const refused = await call('ai:send', { prompt: '요약해줘' });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, 'EDITOR_DIRTY');
    assert.match(refused.error, /저장하지 않은 내용/);
    const sent = await call('ai:send', { prompt: '요약해줘', editorOk: true });
    assert.equal(sent.ok, true, sent.error);
    assert.ok(sent.value.chatId);
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end')), '대화가 끝난다');
  });

  const before = sha(docPath);
  let savedBytes;
  await t.test('편집 창 저장 → 백업·최근 목록·미리보기 알림(doc:changed by editor)', async () => {
    const tmpDoc = join(base, 'saved.hwpx');
    await createBlankDocument(tmpDoc, { title: '편집 창에서 고친 내용' });
    savedBytes = readFileSync(tmpDoc);
    const res = await hostCall('editor:save', new Uint8Array(savedBytes));
    assert.equal(res.ok, true, res.error);
    assert.equal(sha(docPath), shaBytes(savedBytes));
    // (의도한 변경) 편집 탭 문서는 미리보기(문서 탭 뷰)가 없으므로 doc:changed by editor를 받을 화면이 없다
    // — 미리보기 알림은 새 시험 '문서 탭 미리보기: 바깥에서 파일이 바뀌면 그 문서 탭 뷰에 doc:changed'가 본다
    const settings = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));
    assert.equal(settings.recent[0].path, docPath);
    const index = JSON.parse(readFileSync(join(ud, 'backups', 'index.json'), 'utf8'));
    assert.ok(index.some((e) => e.source === docPath && e.hash === before), '덮어쓰기 전 내용이 백업돼 있다');
    fake.emit('editor:dirty', editorWin.wc, false);
    assert.ok(await until(() => lastEditor((d) => !d.dirty)));
  });

  await t.test('바깥에서 파일이 바뀌면 편집 창에 알린다 (감시 → notifyExternalChange)', async () => {
    const n0 = sentTo(editorWin.view, 'editor:external-change').length;
    const other = join(base, 'other.hwpx');
    await createBlankDocument(other, { title: '바깥에서 바꾼 내용' });
    writeFileSync(docPath, readFileSync(other));
    const got = await until(() => sentTo(editorWin.view, 'editor:external-change').length > n0, 8000);
    assert.ok(got, '편집 창이 바깥 변경을 받아야 한다');
    // (의도한 변경) 편집 탭 문서는 문서 탭 뷰(미리보기)가 없다 — 바깥 변경 알림은 편집 뷰가 직접 받는다 (위)
  });

  await t.test('원래대로 → 편집 창에도 알림 (editorNotified)', async () => {
    const n0 = sentTo(editorWin.view, 'editor:external-change').length;
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(sha(docPath), before, '열 때의 내용으로');
    assert.equal(sentTo(editorWin.view, 'editor:external-change').length, n0 + 1);
  });

  await t.test('원래대로: 파일이 이미 그 버전과 같아도 편집 창에 저장 안 한 편집이 있으면 알린다 (다시 불러올지 묻게)', async () => {
    assert.equal(sha(docPath), before);
    const count = () => sentTo(editorWin.view, 'editor:external-change').length;
    const same = (await call('doc:versions')).value.versions.find((v) => v.current);
    assert.ok(same, '지금 내용과 같은 버전');
    // 깨끗하면 알릴 것이 없다
    let n0 = count();
    let r = await call('doc:restore', { file: same.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, false);
    assert.equal(count(), n0);
    // 저장 안 한 편집이 있으면 내용이 같아도 알린다
    fake.emit('editor:dirty', editorWin.wc, true);
    await until(() => lastEditor((d) => d.dirty));
    n0 = count();
    r = await call('doc:restore', { file: same.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    assert.equal(count(), n0 + 1);
    assert.equal(sha(docPath), before);
    fake.emit('editor:dirty', editorWin.wc, false);
    assert.ok(await until(() => lastEditor((d) => !d.dirty)));
  });

  await t.test('버전 목록: 열 때·편집 창 저장 전·AI 대화 전·원래대로 전, 원래대로(인수 없음) = 마지막 변경 전', async () => {
    const w = join(base, 'ver-w.hwpx'); const x = join(base, 'ver-x.hwpx'); const y = join(base, 'ver-y.hwpx');
    await createBlankDocument(w, { title: '편집 창 저장 전 내용' });
    await createBlankDocument(x, { title: 'AI 대화 전 내용' });
    await createBlankDocument(y, { title: 'AI가 고친 내용' });
    writeFileSync(docPath, readFileSync(w));
    // 편집 창 저장 = 덮어쓰기 전 상태를 버전으로
    const saved = await hostCall('editor:save', new Uint8Array(readFileSync(x)));
    assert.equal(saved.ok, true, saved.error);
    assert.equal(sha(docPath), sha(x));
    // AI 대화 한 번 = 고치기 전 상태를 버전으로
    assert.equal((await call('ai:send', { prompt: '고쳐줘', editorOk: true })).ok, true);
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end')));
    writeFileSync(docPath, readFileSync(y)); // AI가 고친 것처럼
    let v = (await call('doc:versions')).value;
    const reasons = v.versions.map((e) => e.reason);
    for (const want of ['open', 'editor', 'ai', 'restore']) assert.ok(reasons.includes(want), `${want}: ${reasons.join(', ')}`);
    assert.ok(v.versions.length <= 30);
    const undo = v.versions.find((e) => e.file === v.undo);
    assert.equal(undo.reason, 'ai', '마지막 변경(AI) 전');
    assert.equal(v.versions.some((e) => e.current), false, 'AI가 고친 지금 내용은 아직 버전이 아님');
    let r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.reason, 'ai');
    assert.equal(sha(docPath), sha(x));
    // 되돌리기 전 상태(AI가 고친 것)가 '원래대로 전'으로 남아 다시 돌아올 수 있다
    v = (await call('doc:versions')).value;
    assert.equal(v.versions[0].reason, 'restore');
    assert.equal(v.undo, v.versions[0].file);
    assert.equal(v.versions.find((e) => e.current)?.reason, 'ai');
    // 처음 연 상태로
    const opened = v.versions.find((e) => e.reason === 'open');
    r = await call('doc:restore', { file: opened.file });
    assert.equal(r.ok, true, r.error);
    assert.equal(sha(docPath), before);
    // 목록에 없는 이름은 거절
    r = await call('doc:restore', { file: '..\settings.json' });
    assert.equal(r.ok, false);
    assert.equal(r.code, 'NO_VERSION');
    assert.equal(sha(docPath), before);
  });

  await t.test('원래대로: 지금 내용을 버전으로 남기지 못하면 되돌리지 않는다 (BACKUP_FAILED)', async () => {
    const v0 = (await call('doc:versions')).value;
    const target = v0.versions.find((e) => e.reason === 'open');
    assert.ok(target, '되돌릴 버전');
    // 지금 내용을 새 내용으로 바꿔 (같은 내용의 백업을 다시 쓰지 않게) 백업 색인 쓰기를 막는다
    const original = readFileSync(docPath);
    writeFileSync(docPath, Buffer.concat([original, Buffer.from('백업 실패 시험')]));
    const changed = sha(docPath);
    const block = join(ud, 'backups', `index.json.${process.pid}.tmp`);
    mkdirSync(block);
    try {
      const r = await call('doc:restore', { file: target.file });
      assert.equal(r.ok, false);
      assert.equal(r.code, 'BACKUP_FAILED');
      assert.match(r.error, /버전으로 남기지 못해 되돌리지 않았습니다/);
      assert.equal(sha(docPath), changed, '되돌리지 않고 지금 내용 그대로');
    } finally {
      rmSync(block, { recursive: true, force: true });
      writeFileSync(docPath, original);
    }
    assert.equal(sha(docPath), before);
  });

  await t.test('원래대로: 문서가 다른 프로그램에 잠겨 버전을 남기지 못하면 "열려 있어" 안내 (BACKUP_FAILED 아님)', { skip: process.platform !== 'win32' && 'Windows 전용 (파일 잠금)' }, async () => {
    const v0 = (await call('doc:versions')).value;
    const target = v0.versions.find((e) => e.reason === 'open');
    assert.ok(target, '되돌릴 버전');
    // 다른 프로그램이 문서를 나눠 쓰기 없이 열어 둔 흉내 (읽기도 EBUSY)
    const ready = join(ud, `lock-ready-${Date.now()}.txt`);
    const q = (s) => s.replace(/'/g, "''");
    const ps = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    const script = `$f = [System.IO.File]::Open('${q(docPath)}', 'Open', 'ReadWrite', 'None'); [System.IO.File]::WriteAllText('${q(ready)}', 'held'); Start-Sleep -Seconds 60; $f.Dispose()`;
    const holder = spawn(ps, ['-NoProfile', '-NonInteractive', '-Command', script], { windowsHide: true, stdio: 'ignore' });
    const gone = new Promise((r) => holder.once('exit', r));
    try {
      const t0 = Date.now();
      while (!existsSync(ready)) {
        if (holder.exitCode !== null || Date.now() - t0 > 60_000) throw new Error('문서를 잠그는 PowerShell이 시작하지 않음');
        await new Promise((r) => setTimeout(r, 50));
      }
      const r = await call('doc:restore', { file: target.file });
      assert.equal(r.ok, false);
      assert.notEqual(r.code, 'BACKUP_FAILED');
      assert.match(r.error, /문서가 다른 프로그램에서 열려 있어 되돌릴 수 없습니다/);
    } finally {
      holder.kill();
      await gone;
      rmSync(ready, { force: true });
    }
    assert.equal(sha(docPath), before, '되돌리지 않고 그대로');
  });

  // ── 문서·표·슬라이드 편집 창 (Office) ──

  let officeWin; let officeCall;
  const docxBefore = sha(docxPath);
  await t.test('한글 편집 창: 자동 저장은 복구 사본에만, 다시 열면 복구/버리기, 원본은 그대로', async () => {
    const before = readFileSync(docPath);
    const bytes = new Uint8Array(before);
    const r = await hostCall('editor:save', bytes, { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.draft, true);
    const draft = draftPathFor(ud, docPath);
    assert.ok(existsSync(draft), '복구 사본');
    assert.deepEqual(readFileSync(docPath), before, '원본은 그대로');
    // 파일보다 새로운 사본 → 다시 열면 알린다
    utimesSync(draft, new Date(Date.now() + 5000), new Date(Date.now() + 5000));
    const d = await hostCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.ok(d.value.draft && d.value.draft.size === bytes.length, JSON.stringify(d.value.draft));
    const read = await hostCall('editor:draft', 'read');
    assert.equal(read.ok, true, read.error);
    assert.equal(read.value.bytes.length, bytes.length);
    // 형식이 아닌 바이트는 사본으로도 쓰지 않는다
    const bad = await hostCall('editor:save', new Uint8Array(Buffer.from('not hwp')), { reason: 'auto' });
    assert.equal(bad.ok, false);
    assert.equal((await hostCall('editor:draft', 'discard')).ok, true);
    assert.equal(existsSync(draft), false);
  });

  await t.test('한글 편집 창: 지금 내용을 PDF로 저장 (편집기 인쇄 대신)', { skip: !hwpPdfStatus().ok && '내장 PDF 엔진(rhwp.exe) 없음' }, async () => {
    const h = Hwp.open(docPath);
    let hwpBytes;
    try { hwpBytes = new Uint8Array(h.toBytes('hwp')); } finally { h.close(); }
    const target = join(docs, '보고서 출력.pdf');
    fake.saveDialogAnswers.push({ canceled: false, filePath: target });
    const r = await hostCall('editor:hwp-pdf', hwpBytes, 'save');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '보고서 출력.pdf');
    assert.equal(readFileSync(target).subarray(0, 5).toString('latin1'), '%PDF-');
    // hwp가 아닌 바이트는 거절
    const no = await hostCall('editor:hwp-pdf', new Uint8Array(Buffer.from('PK\x03\x04')), 'save');
    assert.equal(no.ok, false);
    assert.equal(no.code, 'BAD_FORMAT');
  });

  await t.test('Office: docx를 문서 편집 창으로 연다 (따로 된 세션·네트워크 막기)', async () => {
    assert.equal((await call('doc:open', docxPath)).ok, true);
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.kind, 'office');
    assert.equal(r.value.path, docxPath);
    assert.equal(r.value.reused, false);
    assert.equal(r.value.notice, null);
    officeWin = editorOf(/메모\.docx - 문서 편집/);
    assert.equal(officeWin.win.getTitle(), '메모.docx - 문서 편집 (Office Kit)');
    assert.doesNotMatch(officeWin.win.getTitle(), /ONLYOFFICE/i);
    assert.match(officeWin.wc.mainFrame.url, /^http:\/\/127\.0\.0\.1:\d+\/office-host\/host\.html$/);
    assert.equal(officeWin.view.opts.webPreferences.partition, 'officekit-office');
    assert.match(officeWin.view.opts.webPreferences.preload, /office-preload\.cjs$/);
    assert.equal(officeWin.wc.webrtcPolicy, 'disable_non_proxied_udp');
    const ses = fake.sessions.get('officekit-office');
    assert.equal(ses.proxy.proxyRules, 'socks5://127.0.0.1:9');
    const blocked = await new Promise((res) => ses.handlers.beforeRequest({ url: 'https://example.com/x.js' }, res));
    assert.equal(blocked.cancel, true);
    officeCall = (channel, ...args) => fake.invokeFrom(officeWin.wc, channel, ...args);
    const st = await until(() => lastEditor((d) => d.path === docxPath && d.open));
    assert.deepEqual(st, { path: docxPath, open: true, dirty: false });
    // 한글 편집 창의 host에서 Office 전용 요청은 거절
    const no = await hostCall('editor:state', 'get', 'x');
    assert.equal(no.ok, false);
  });

  await t.test('Office: get-document → 편집기 주소·한국어·판 정보 (복구 사본 없음)', async () => {
    const d = await officeCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    const v = d.value;
    assert.equal(v.kind, 'office');
    assert.equal(v.name, '메모.docx');
    assert.equal(v.ext, 'docx');
    assert.equal(v.locale, 'ko-KR');
    assert.equal(v.available, true);
    assert.equal(v.title, '문서 편집');
    assert.match(v.appUrl, /^http:\/\/127\.0\.0\.\d+:\d+\/office\/[0-9a-f]{16}-[0-9a-f]{32}\/index\.html$/, '판-실행마다 새 열쇠');
    assert.equal(typeof v.userName, 'string');
    assert.ok(v.userName.length > 0, '사용자 이름 (Office 이름 → Windows 계정 이름)');
    assert.equal(v.draft, null);
    assert.equal(v.about.appVersion, '0.0.0-test');
    assert.equal(v.about.bundle.onlyoffice.build, '9.4.0.129');
    assert.equal(v.about.bundle.files, undefined, '파일 목록은 넘기지 않는다');
    assert.equal(shaBytes(v.bytes), docxBefore);
    assert.ok(globalThis.__okFontCalls >= 1, '글꼴 목록을 준비한다');
  });

  await t.test('Office: 편집기 설정(state.set/get)은 settings.json의 officeEditor.settings에', async () => {
    const value = { zoom: 120, rulers: true, name: '보기 설정' };
    let r = await officeCall('editor:state', 'set', 'de-settings', value);
    assert.equal(r.ok, true, r.error);
    r = await officeCall('editor:state', 'get', 'de-settings');
    assert.deepEqual(r.value, value);
    const file = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));
    assert.deepEqual(file.officeEditor.settings['de-settings'], value);
    // 화면(주 창) 설정에는 내놓지 않는다
    const pub = await call('settings:get');
    assert.equal(pub.value.officeEditor, undefined);
    // 지우기
    r = await officeCall('editor:state', 'set', 'de-settings', null);
    assert.equal(r.ok, true, r.error);
    assert.equal((await officeCall('editor:state', 'get', 'de-settings')).value, null);
    // 너무 큰 값·이상한 이름은 거절
    assert.equal((await officeCall('editor:state', 'set', 'big', 'x'.repeat(9000))).code, 'TOO_BIG');
    assert.equal((await officeCall('editor:state', 'set', '../x', 1)).ok, false);
    r = await officeCall('editor:state', 'set', 'de-settings', value);
    assert.equal(r.ok, true);
  });

  const draftFile = draftPathFor(ud, docxPath);
  let editedDocx;
  await t.test('Office 자동 저장(reason auto) → 복구 사본에만, 문서 파일은 그대로', async () => {
    const tmp = join(base, 'edited.docx');
    await createDocument(tmp, '# 회의 메모\n\n편집 창에서 고친 내용입니다.\n\n- 안건 하나\n- 안건 둘\n- 새 안건');
    editedDocx = readFileSync(tmp);
    const changes0 = allSent().filter((s) => s.channel === 'doc:changed' && s.data.by === 'editor').length;
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.draft, true);
    assert.equal(sha(docxPath), docxBefore, '사용자 파일을 몰래 덮어쓰지 않는다');
    assert.ok(existsSync(draftFile), '복구 사본');
    assert.equal(sha(draftFile), shaBytes(editedDocx));
    assert.equal(dirname(draftFile), join(ud, 'office-drafts'));
    await sleep(100);
    assert.equal(allSent().filter((s) => s.channel === 'doc:changed' && s.data.by === 'editor').length, changes0, '저장 알림 없음');
  });

  await t.test('Office: 저장 안 한 편집 → 창 제목 *, 주 창 상태, AI 보내기 막기', async () => {
    fake.emit('editor:dirty', officeWin.wc, true);
    const st = await until(() => lastEditor((d) => d.path === docxPath && d.dirty));
    assert.deepEqual(st, { path: docxPath, open: true, dirty: true });
    assert.equal(officeWin.win.getTitle(), '* 메모.docx - 문서 편집 (Office Kit)');
    const info = await call('doc:open', docxPath);
    assert.deepEqual(info.value.editor, { open: true, dirty: true });
    const refused = await call('ai:send', { prompt: '표 하나 넣어줘' });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, 'EDITOR_DIRTY');
  });

  await t.test('Office 원래대로: 저장 안 한 편집이 있으면 편집 창에 알린다 (reason restore)', async () => {
    const n0 = sentTo(officeWin.view, 'editor:external-change').length;
    const r = await call('doc:restore');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.editorNotified, true);
    const got = sentTo(officeWin.view, 'editor:external-change');
    assert.equal(got.length, n0 + 1);
    assert.equal(got.at(-1).data.reason, 'restore');
    assert.equal(sha(docxPath), docxBefore);
  });

  await t.test('Office 저장(reason user) → 형식 확인·고치기·백업·쓰기, 복구 사본 지움, 미리보기 알림', async () => {
    // 사용자가 누르거나 친 것이 없으면 편집기 틀이 원본을 저장하지 못한다
    const sneaky = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'user' });
    assert.equal(sneaky.ok, false);
    assert.equal(sneaky.code, 'NO_USER_ACTION');
    officeWin.wc.emit('input-event', {}, { type: 'mouseDown' });
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'user' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '메모.docx');
    assert.ok(Array.isArray(r.value.fixups));
    assert.equal(r.value.warning, null);
    assert.match(await readDocumentText(docxPath), /편집 창에서 고친 내용/);
    assert.equal(existsSync(draftFile), false, '정상 저장 → 복구 사본은 필요 없다');
    // (의도한 변경) 편집 탭 문서는 문서 탭 뷰가 없어 doc:changed by editor를 받을 화면이 없다
    const index = JSON.parse(readFileSync(join(ud, 'backups', 'index.json'), 'utf8'));
    assert.ok(index.some((e) => e.source === docxPath && e.hash === docxBefore), '덮어쓰기 전 내용이 백업돼 있다');
    // 형식이 아닌 바이트는 저장하지 않는다
    const bad = await officeCall('editor:save', new Uint8Array(Buffer.from('not a zip')), { reason: 'user' });
    assert.equal(bad.ok, false);
    assert.equal(bad.code, 'BAD_FORMAT');
    fake.emit('editor:dirty', officeWin.wc, false);
    assert.ok(await until(() => lastEditor((d) => d.path === docxPath && !d.dirty)));
    assert.equal((await call('ai:send', { prompt: '요약해줘' })).ok, true, '저장한 뒤에는 보낸다');
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end' && d.path === docxPath)));
  });

  // ── AI 대상 (고칠 문서 · 재료) ── 재료 = 편집 창으로 열려 있는 메모.docx, 고칠 문서 = 보고서.hwpx
  const 재료 = docxPath;
  const endOf = (chatId) => until(() => lastSent('ai:event', (d) => d.type === 'end' && d.chatId === chatId));
  await t.test('AI 대상: 대상 없이 보내면 지금 문서가 고칠 문서 (지금과 같음)', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    const r = await call('ai:send', { prompt: '요약해줘', editorOk: true });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.newChat, false);
    const tg = globalThis.__okTargets.at(-1);
    assert.equal(tg.edit, docPath);
    assert.deepEqual(tg.read, []);
    assert.equal(tg.zone3, false);
    const c = globalThis.__okChats.at(-1);
    assert.equal(c.filePath, docPath);
    assert.deepEqual(c.materials, []);
    assert.ok(await endOf(r.value.chatId));
  });
  await t.test('AI 대상: 재료는 READ 목록과 데이터 줄 재료로, 재료가 고칠 문서와 같으면 뺀다', async () => {
    const r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target: { edit: docPath, read: [재료, docPath.toUpperCase()] } });
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(globalThis.__okTargets.at(-1).read, [재료]);
    assert.deepEqual(globalThis.__okChats.at(-1).materials, [재료]);
    assert.ok(await endOf(r.value.chatId));
  });
  await t.test('AI 대상: 지금 문서가 바뀌었으면 TARGET_CHANGED, 열리지 않은 파일은 MATERIAL_NOT_OPEN, 11개는 거절', async () => {
    let r = await call('ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: 재료, read: [] } });
    assert.equal(r.code, 'TARGET_CHANGED');
    r = await call('ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: docPath, read: [csvPath] } });
    assert.equal(r.code, 'MATERIAL_NOT_OPEN');
    assert.match(r.error, /목록\.csv/);
    const many = join(base, '재료 여럿');
    mkdirSync(many, { recursive: true });
    const files = Array.from({ length: 11 }, (_, i) => { const p = join(many, `${i + 1}.csv`); writeFileSync(p, 'a,b\n'); return p; });
    r = await call('ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: docPath, read: files } });
    assert.equal(r.ok, false);
    assert.match(r.error, /10개/);
  });
  await t.test('AI 대상: AI가 고치는 중인 문서는 다른 대화의 재료가 될 수 없다 (D14)', async () => {
    const held = await call('ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    try {
      let r = await call('ai:send', { prompt: '또 고쳐줘', editorOk: true });
      assert.equal(r.code, 'AI_BUSY');
      assert.equal((await call('doc:open', 재료)).ok, true);
      r = await call('ai:send', { prompt: '고쳐줘', target: { edit: 재료, read: [docPath] } });
      assert.equal(r.code, 'MATERIAL_BUSY');
      assert.match(r.error, /보고서\.hwpx/);
      if (process.platform === 'win32') {
        r = await call('ai:send', { prompt: '고쳐줘', target: { edit: 재료, read: [docPath.toLowerCase()] } });
        assert.equal(r.code, 'MATERIAL_BUSY', '대소문자만 다른 경로도 같은 문서');
      }
    } finally {
      globalThis.__okRelease();
    }
    assert.ok(await endOf(held.value.chatId));
    const r = await call('ai:send', { prompt: '고쳐줘', target: { edit: 재료, read: [docPath] } });
    assert.equal(r.ok, true, r.error);
    assert.ok(await endOf(r.value.chatId));
    assert.equal((await call('doc:open', docPath)).ok, true);
  });
  await t.test('AI 대상: 재료 전용 — 지금 문서가 PDF면 고칠 문서 없음, 버전 저장 없음, 출력 폴더 기본=첫 재료 폴더', async () => {
    const pdf = join(docs, '자료.pdf');
    writeFileSync(pdf, makePdf(1));
    try {
      assert.equal((await call('doc:open', pdf)).ok, true);
      const n0 = (await call('doc:versions')).value.versions.length;
      let r = await call('ai:send', { prompt: '요약해줘' });
      assert.equal(r.ok, true, r.error);
      assert.deepEqual(globalThis.__okTargets.at(-1), { edit: null, read: [pdf], outDir: docs, zone3: false });
      const c = globalThis.__okChats.at(-1);
      assert.equal(c.filePath, null);
      assert.equal(c.outDir, docs);
      const end = await endOf(r.value.chatId);
      assert.ok(end);
      assert.equal(end.changed, undefined, '고칠 문서가 없으면 바뀜 알림도 없다');
      assert.equal((await call('doc:versions')).value.versions.length, n0, 'AI 전 버전을 남기지 않는다');
      r = await call('ai:send', { prompt: '요약해줘', outDir: join(base, '없는 폴더') });
      assert.equal(r.code, 'BAD_OUT_DIR');
      r = await call('ai:send', { prompt: '요약해줘', outDir: base });
      assert.equal(r.ok, true, r.error);
      assert.equal(globalThis.__okTargets.at(-1).outDir, base);
      assert.equal(globalThis.__okChats.at(-1).outDir, base);
      assert.ok(await endOf(r.value.chatId));
    } finally {
      assert.equal((await call('doc:open', docPath)).ok, true);
      rmSync(pdf, { force: true });
    }
  });
  await t.test('AI 대상: 인터넷 재료 → zone3:true, 빼고 보내면 이어 하기 세션을 버린다', { skip: process.platform !== 'win32' && 'Zone.Identifier' }, async () => {
    const mark = `${재료}:Zone.Identifier`;
    writeFileSync(mark, '[ZoneTransfer]\r\nZoneId=3\r\n');
    try {
      let r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target: { edit: docPath, read: [재료] } });
      assert.equal(r.ok, true, r.error);
      assert.equal(globalThis.__okTargets.at(-1).zone3, true);
      assert.equal(r.value.newChat, false);
      assert.ok(await endOf(r.value.chatId));
      r = await call('ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: docPath, read: [] } });
      assert.equal(r.ok, true, r.error);
      assert.equal(r.value.newChat, true, '인터넷 재료를 본 대화는 이어 하지 않는다');
      assert.equal(globalThis.__okTargets.at(-1).zone3, false);
      assert.ok(await endOf(r.value.chatId));
      r = await call('ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: docPath, read: [] } });
      assert.equal(r.ok, true, r.error);
      assert.equal(r.value.newChat, false);
      assert.ok(await endOf(r.value.chatId));
    } finally {
      rmSync(mark, { force: true });
    }
  });
  await t.test('AI 재료: 저장 안 한 재료 → MATERIAL_DIRTY, 그대로 보내기는 통과, 저장하고 보내기는 저장 실패를 알린다', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    const win = officeWin.view; // 재료 = 메모.docx 편집 창
    fake.emit('editor:dirty', win.webContents, true);
    try {
      const target = { edit: docPath, read: [재료] };
      let r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target });
      assert.equal(r.ok, false);
      assert.equal(r.code, 'MATERIAL_DIRTY');
      assert.match(r.error, /메모\.docx/);
      assert.match(r.error, /저장된 내용을 읽습니다/);
      r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, materialsOk: true, target });
      assert.equal(r.ok, true, r.error);
      assert.ok(await endOf(r.value.chatId));
      const n0 = sentTo(win, 'editor:command').length;
      r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, saveMaterials: true, target });
      assert.equal(r.ok, false);
      assert.equal(r.code, 'MATERIAL_SAVE_FAILED');
      assert.match(r.error, /메모\.docx/);
      const cmds = sentTo(win, 'editor:command').slice(n0);
      assert.equal(cmds.length, 1);
      assert.equal(cmds[0].data.command, 'save');
    } finally {
      fake.emit('editor:dirty', win.webContents, false);
    }
  });
  await t.test('AI 재료: 편집 창은 깨끗해도 파일보다 새 자동 저장 사본이 있으면 묻는다 (OV6)', async () => {
    const draft = draftPathFor(ud, 재료);
    mkdirSync(dirname(draft), { recursive: true });
    writeFileSync(draft, 'PK');
    const target = { edit: docPath, read: [재료] };
    try {
      const orig = statSync(재료).mtimeMs;
      utimesSync(draft, new Date(orig + 60_000), new Date(orig + 60_000));
      let r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target });
      assert.equal(r.code, 'MATERIAL_DIRTY');
      utimesSync(draft, new Date(orig - 60_000), new Date(orig - 60_000));
      r = await call('ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target });
      assert.equal(r.ok, true, r.error);
      assert.ok(await endOf(r.value.chatId));
    } finally {
      rmSync(draft, { force: true });
    }
  });
  await t.test('AI 재료: ai:materials — 고칠 문서·열린 편집 창·AI 작업 중·재료 전용·인터넷 표시', async () => {
    assert.equal((await call('doc:open', docPath)).ok, true);
    let r = (await call('ai:materials')).value;
    assert.equal(r.edit.path, docPath);
    assert.equal(r.edit.current, true);
    assert.ok(r.open.some((m) => m.path === 재료 && m.kind === 'Word' && m.busy === false));
    assert.ok(!r.open.some((m) => m.path === docPath));
    const held = await call('ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    try {
      assert.equal((await call('doc:open', 재료)).ok, true);
      r = (await call('ai:materials')).value;
      assert.equal(r.edit.path, 재료);
      const busy = r.open.find((m) => m.path === docPath);
      assert.ok(busy, '고칠 문서가 바뀌어도 열린 편집 창은 후보');
      assert.equal(busy.busy, true);
    } finally {
      globalThis.__okRelease();
    }
    assert.ok(await endOf(held.value.chatId));
    assert.equal((await call('doc:open', docPath)).ok, true);
    const pdf = join(docs, '자료2.pdf');
    writeFileSync(pdf, makePdf(1));
    try {
      assert.equal((await call('doc:open', pdf)).ok, true);
      r = (await call('ai:materials')).value;
      assert.equal(r.edit, null);
      assert.equal(r.open[0].path, pdf);
      assert.equal(r.open[0].current, true);
    } finally {
      assert.equal((await call('doc:open', docPath)).ok, true);
      rmSync(pdf, { force: true });
    }
    if (process.platform === 'win32') {
      const mark = `${재료}:Zone.Identifier`;
      writeFileSync(mark, '[ZoneTransfer]\r\nZoneId=3\r\n');
      try {
        r = (await call('ai:materials')).value;
        assert.equal(r.open.find((m) => m.path === 재료).zone3, true);
      } finally {
        rmSync(mark, { force: true });
      }
    }
  });
  // 아래 시험은 메모.docx가 지금 문서인 상태에서 이어진다
  assert.equal((await call('doc:open', docxPath)).ok, true);

  await t.test('Office 닫기: 저장 안 함 → 창 닫힘, 파일 그대로, 복구 사본 지움 (정상 닫기)', async () => {
    const r = await officeCall('editor:save', new Uint8Array(editedDocx), { reason: 'auto' });
    assert.equal(r.ok, true, r.error);
    assert.ok(existsSync(draftFile));
    fake.emit('editor:dirty', officeWin.wc, true);
    fake.replies.set(officeWin.wc.id, true);
    const hash = sha(docxPath);
    const d0 = fake.dialogs.length;
    fake.dialogAnswers.push(1); // 저장 안 함
    void callAs('tabstrip', 'tabs:close', tabOf(docxPath).id);
    assert.ok(await until(() => officeWin.wc.isDestroyed(), 8000), '창이 닫혀야 한다');
    assert.equal(fake.dialogs.length, d0 + 1);
    assert.deepEqual(fake.dialogs[d0].buttons, ['저장', '저장 안 함', '취소']);
    assert.equal(fake.dialogs[d0].title, '문서 편집');
    assert.match(fake.dialogs[d0].message, /메모\.docx/);
    assert.equal(sha(docxPath), hash);
    assert.equal(existsSync(draftFile), false, '정상 닫기 → 복구 사본 지움');
    assert.ok(await until(() => lastEditor((d) => d.path === docxPath && !d.open)));
  });

  await t.test('Office 복구 사본: 파일보다 새로우면 다시 열 때 알리고, 읽기·버리기', async () => {
    put(draftFile, editedDocx);
    const future = new Date(Date.now() + 60_000);
    utimesSync(draftFile, future, future);
    await call('doc:open', docxPath); // 편집 탭을 닫았으므로 문서 탭으로 다시 연다
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    officeWin = editorOf(/메모\.docx - 문서 편집/);
    officeCall = (channel, ...args) => fake.invokeFrom(officeWin.wc, channel, ...args);
    const d = await officeCall('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.equal(d.value.draft?.size, editedDocx.length);
    const read = await officeCall('editor:draft', 'read');
    assert.equal(read.ok, true, read.error);
    assert.equal(shaBytes(read.value.bytes), shaBytes(editedDocx));
    const gone = await officeCall('editor:draft', 'discard');
    assert.deepEqual(gone.value, { discarded: true });
    assert.equal(existsSync(draftFile), false);
    assert.equal((await officeCall('editor:draft', 'read')).code, 'NO_DRAFT');
  });

  let newEditor;
  await t.test('새 문서 edit:true → 한글은 한글 편집, pptx는 슬라이드 편집, xlsx는 표 편집 창', async () => {
    const hw = await call('doc:new', { format: 'hwpx', folder: docs, name: '새 편집', start: 'blank', edit: true });
    assert.equal(hw.ok, true, hw.error);
    assert.equal(hw.value.editorError, null);
    assert.equal(hw.value.editor.path, join(docs, '새 편집.hwpx'));
    assert.equal(hw.value.editor.kind, 'hwp');
    newEditor = editorOf(/새 편집\.hwpx - 한글 편집/);
    assert.match(newEditor.win.getTitle(), /새 편집\.hwpx - 한글 편집/);
    const pp = await call('doc:new', { format: 'pptx', folder: docs, name: '새 발표', start: 'blank', edit: true });
    assert.equal(pp.ok, true, pp.error);
    assert.equal(pp.value.editorError, null);
    assert.equal(pp.value.editor.kind, 'office');
    assert.equal(tw0.getTitle(), '새 발표.pptx - 슬라이드 편집 (Office Kit)');
    const xl = await call('doc:new', { format: 'xlsx', folder: docs, name: '새 표', start: 'blank', edit: true });
    assert.equal(xl.ok, true, xl.error);
    assert.equal(tw0.getTitle(), '새 표.xlsx - 표 편집 (Office Kit)');
    // edit 없이 만들면 편집 창을 열지 않는다
    const n = editorViews().length;
    const wd = await call('doc:new', { format: 'docx', folder: docs, name: '새 워드', start: 'blank' });
    assert.equal(wd.ok, true, wd.error);
    assert.equal(wd.value.editor, null);
    assert.equal(editorViews().length, n);
  });

  await t.test('Office 편집 창이 4개를 넘으면 메모리 안내 (notice)', async () => {
    // 지금 Office 창: 메모.docx, 새 발표.pptx, 새 표.xlsx
    const notices = [];
    for (const name of ['넷째', '다섯째']) {
      const p = join(docs, `${name}.docx`);
      await createBlankDocument(p);
      assert.equal((await call('doc:open', p)).ok, true);
      const r = await call('doc:edit');
      assert.equal(r.ok, true, r.error);
      notices.push(r.value.notice);
    }
    assert.equal(notices[0], null, '4개까지는 안내 없음');
    assert.match(notices[1], /5개 열려 있습니다/);
    assert.match(notices[1], /메모리/);
    // 메모리 안내는 main이 그 탭 창을 부모로 한 정보 대화상자로 보인다 (D12 — 바뀌어 닫히는 문서 탭에 맡기지 않는다)
    const shown = fake.dialogs.filter((d) => d.type === 'info' && /5개 열려 있습니다/.test(d.message ?? ''));
    assert.equal(shown.length, 1);
    assert.equal(shown[0].parent, tw0);
    assert.deepEqual(shown[0].buttons, ['확인']);
  });

  await t.test('Office: 암호 문서·옛 형식은 창을 열기 전에 한국어로 거절', async () => {
    const old = join(docs, '옛 문서.docx');
    const ole = Buffer.alloc(1024);
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(ole);
    writeFileSync(old, ole);
    assert.equal((await call('doc:open', old)).ok, true);
    const n = editorViews().length;
    let r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'OLD_FORMAT');
    assert.match(r.error, /프로그램에서 여세요/);
    Buffer.from('EncryptionInfo', 'utf16le').copy(ole, 512);
    writeFileSync(old, ole);
    r = await call('doc:edit');
    assert.equal(r.code, 'ENCRYPTED');
    assert.match(r.error, /암호가 걸린 문서는 직접 편집할 수 없습니다/);
    assert.equal(editorViews().length, n);
  });

  await t.test('한글: 내장 엔진이 못 여는 문서(손상)는 편집 창을 띄우기 전에 거절 — 엔진 원래 글은 detail로', async () => {
    const bad = join(docs, '깨진 문서.hwp');
    const bytes = Buffer.alloc(4096, 0x5a);
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(bytes);
    writeFileSync(bad, bytes);
    assert.equal((await call('doc:open', bad)).ok, true);
    const n = editorViews().length;
    const r = await call('doc:edit');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'BAD_FORMAT');
    assert.match(r.error, /^‘깨진 문서\.hwp’ 파일은 손상됐거나 지원하지 않는 한글 문서라 직접 편집할 수 없어요/);
    assert.ok(r.detail, '엔진이 낸 원래 글');
    assert.equal(editorViews().length, n, '창을 띄우지 않음');
  });

  await t.test('다른 이름으로 저장: 확장자를 붙인 이름이 이미 있으면 다시 묻고, 취소하면 그대로', async () => {
    const h2 = (channel, ...args) => fake.invokeFrom(newEditor.wc, channel, ...args);
    assert.equal((await h2('editor:get-document')).ok, true);
    const victim = join(docs, 'victim.docx.hwpx');
    writeFileSync(victim, 'VICTIM');
    const bytes = new Uint8Array(readFileSync(join(docs, '새 편집.hwpx')));
    const d0 = fake.dialogs.length;
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, 'victim.docx') });
    fake.dialogAnswers.push(1); // 취소
    let r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(r.value, { canceled: true });
    assert.equal(fake.dialogs.length, d0 + 1);
    assert.match(fake.dialogs.at(-1).message, /victim\.docx\.hwpx/);
    assert.equal(readFileSync(victim, 'utf8'), 'VICTIM');
    // 바꾸기를 고르면 저장
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, 'victim.docx') });
    fake.dialogAnswers.push(0);
    r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, 'victim.docx.hwpx');
    assert.equal(sha(victim), shaBytes(bytes));
    // 확장자가 맞고 없는 이름이면 묻지 않는다
    const d1 = fake.dialogs.length;
    fake.saveDialogAnswers.push({ canceled: false, filePath: join(docs, '다른 이름.hwpx') });
    r = await h2('editor:save-as', bytes, 'hwpx');
    assert.equal(r.ok, true, r.error);
    assert.equal(fake.dialogs.length, d1);
  });

  await t.test('Office 내려받기(PDF 등) → 열 수 있는 형식이면 최근 목록에 (PDF도 이제 열 수 있음)', async () => {
    const w = editorOf(/다섯째\.docx/);
    const dl = (...a) => fake.invokeFrom(w.wc, 'editor:download', ...a);
    assert.equal((await dl('몰래.pdf', new Uint8Array(Buffer.from('%PDF-1.7\n%%EOF\n')), 'application/pdf')).code, 'NO_USER_ACTION', '사용자 조작 없이 내려받기 창을 띄우지 못한다');
    w.wc.emit('input-event', {}, { type: 'keyDown' });
    const odt = join(docs, '내보낸 문서.odt');
    fake.saveDialogAnswers.push({ canceled: false, filePath: odt });
    let r = await dl('내보낸 문서.odt', new Uint8Array(readFileSync(docxPath)), 'application/vnd.oasis.opendocument.text');
    assert.equal(r.ok, true, r.error);
    let recent = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).recent;
    assert.equal(recent[0].path, odt);
    const pdf = join(docs, '내보낸 문서.pdf');
    fake.saveDialogAnswers.push({ canceled: false, filePath: pdf });
    r = await dl('내보낸 문서.pdf', new Uint8Array(Buffer.from('%PDF-1.7\n%%EOF\n')), 'application/pdf');
    assert.equal(r.ok, true, r.error);
    recent = JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).recent;
    assert.equal(recent[0].path, pdf);
  });

  await t.test("편집 탭의 'AI · 도구' → 같은 탭 창의 AI 패널이 펼쳐지고 ai:target이 그 문서", async () => {
    // 패널을 접어 두고 홈 탭을 보던 중
    assert.equal((await callAs('tabstrip', 'tabs:togglePanel')).ok, true);
    assert.equal(tabList().panelCollapsed, true);
    const panelView = fake.views.find((v) => v.webContents === panelWc);
    assert.equal(panelView.visible, false, '접힌 패널은 숨긴다');
    assert.equal(JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).chatCollapsed, true);
    assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true);
    assert.equal(lastTarget(), null, '홈 탭이면 AI 대상 없음');
    panelWc.focused = false;
    const windows = fake.windows.length;
    fake.emit('editor:panel', officeWin.wc);
    assert.equal(tabList().active, tabOf(docxPath).id, '그 편집 탭이 활성');
    assert.equal(tabList().panelCollapsed, false, '패널이 펼쳐진다');
    assert.equal(panelView.visible, true);
    assert.equal(panelWc.isFocused(), true, '패널에 초점');
    assert.equal(lastTarget()?.path, docxPath, 'ai:target = 그 문서');
    assert.deepEqual(lastTarget().editor, { open: true, dirty: false });
    assert.equal(JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8')).chatCollapsed, false);
    assert.equal(fake.windows.length, windows, '옆 창을 띄우지 않는다');
    cur = docxPath;
  });

  await t.test('홈에서 연 PDF는 문서 탭 뷰(?view=doc), 같은 문서를 다시 열면 그 탭', async () => {
    const pdf = join(docs, '탭 시험.pdf');
    writeFileSync(pdf, makePdf(1));
    const n = tabList().tabs.length;
    const r = await call('doc:open', pdf);
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.mode, 'doc');
    const wc = docWc(pdf);
    assert.ok(wc, '문서 탭 뷰가 그 문서를 연다');
    assert.match(wc.mainFrame.url, /\/renderer\/index\.html\?view=doc$/);
    assert.equal(tabList().tabs.length, n + 1);
    assert.equal(tabList().active, r.value.tabId);
    assert.equal(tabOf(pdf).type, 'doc');
    assert.equal(lastTarget()?.path, pdf, 'AI 대상 = 그 문서');
    assert.equal(tw0.getTitle(), '탭 시험.pdf - Office Kit');
    // 다른 탭으로 갔다가 같은 문서를 다시 열면 새 탭 없이 그 탭으로
    assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true);
    assert.equal(tabList().active, 'home');
    assert.equal(tw0.getTitle(), 'Office Kit');
    const again = await call('doc:open', process.platform === 'win32' ? pdf.toUpperCase() : pdf); // 대소문자만 다른 경로도 같은 문서
    assert.equal(again.ok, true, again.error);
    assert.equal(again.value.mode, 'reused');
    assert.equal(again.value.tabId, r.value.tabId);
    assert.equal(tabList().tabs.length, n + 1);
    assert.equal(tabList().active, r.value.tabId);
    cur = pdf;
  });

  await t.test('문서 탭 미리보기: 바깥에서 파일이 바뀌면 그 문서 탭 뷰에 doc:changed', async () => {
    const pdf = join(docs, '탭 시험.pdf');
    const wc = docWc(pdf);
    const n0 = wc.sent.filter((s) => s.channel === 'doc:changed').length;
    writeFileSync(pdf, makePdf(2));
    assert.ok(await until(() => wc.sent.filter((s) => s.channel === 'doc:changed').length > n0, 8000), '그 문서 탭 뷰가 받는다');
    assert.equal(wc.sent.filter((s) => s.channel === 'doc:changed').at(-1).data.path, pdf);
    assert.equal(homeWc.sent.some((s) => s.channel === 'doc:changed' && s.data.path === pdf), false, '홈 뷰는 문서를 품지 않는다');
  });

  await t.test('AI 작업 중인 문서 탭: doc:editor로 aiBusy가 오고, 끝나면 false', async () => {
    const pdf = join(docs, '탭 시험.pdf');
    assert.equal((await call('doc:open', pdf)).ok, true);
    const wc = docWc(pdf);
    const busyOf = () => wc.sent.filter((s) => s.channel === 'doc:editor' && s.data?.path === pdf).at(-1)?.data?.aiBusy;
    const held = await call('ai:send', { prompt: '[hold] 요약해줘' });
    assert.equal(held.ok, true, held.error);
    try {
      assert.equal(busyOf(), true, '대화 시작 → 그 문서 탭에 aiBusy (미리보기 새로 고침 이유가 ai가 되게)');
      assert.equal(lastTarget()?.aiBusy, true);
    } finally {
      globalThis.__okRelease();
    }
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end' && d.chatId === held.value.chatId)));
    assert.ok(await until(() => busyOf() === false), '끝나면 aiBusy false');
  });

  await t.test('view:call 중계: 표에 없는 이름·같은 종류로 보내기는 버림, chat.loadProviders는 모든 창 패널로', async () => {
    const pdf = join(docs, '탭 시험.pdf');
    const pdfWc = docWc(pdf);
    const got = (wc) => wc.sent.filter((s) => s.channel === 'view:called');
    const p0 = got(panelWc).length; const h0 = got(homeWc).length; const d0 = got(pdfWc).length;
    fake.emit('view:call', homeWc, { name: 'chat.nope', args: [] });
    fake.emit('view:call', panelWc, { name: 'chat.markStopping', args: [{ path: pdf }] });
    fake.emit('view:call', homeWc, { name: 'home.renderNewDoc', args: [] });
    fake.emit('view:call', tabstripWc, { name: 'chat.loadProviders', args: [] });
    assert.equal(got(panelWc).length, p0, '표에 없는 이름·같은 종류·탭 줄은 버린다');
    assert.equal(got(homeWc).length, h0);
    fake.emit('view:call', homeWc, { name: 'chat.loadProviders', args: [false] });
    assert.deepEqual(got(panelWc).at(-1)?.data, { name: 'chat.loadProviders', args: [false] });
    // 경로가 있는 이름은 그 문서를 보는 문서 탭 뷰로만
    fake.emit('view:call', panelWc, { name: 'home.refreshPreview', args: [{ path: pdf, ifChanged: true, reason: 'ai' }] });
    assert.deepEqual(got(pdfWc).at(-1)?.data, { name: 'home.refreshPreview', args: [{ path: pdf, ifChanged: true, reason: 'ai' }] });
    assert.equal(got(homeWc).length, h0, '홈 뷰는 받지 않는다');
    assert.equal(got(pdfWc).length, d0 + 1);
    // focusHome: 홈 탭으로 간 뒤 홈 뷰에
    fake.emit('view:call', panelWc, { name: 'home.pickFile', args: [] });
    assert.deepEqual(got(homeWc).at(-1)?.data, { name: 'home.pickFile', args: [] });
    assert.equal(tabList().active, 'home');
  });

  await t.test('doc:restore는 그 뷰의 문서·AI 대상 경로만', async () => {
    const pdf = join(docs, '탭 시험.pdf');
    assert.equal((await call('doc:open', docPath)).ok, true); // 편집 탭 → AI 대상
    // AI 패널: AI 대상이 아닌 경로는 거절
    for (const ch of ['doc:restore', 'doc:versions', 'doc:diff']) {
      const r = await callAs('panel', ch, { path: docxPath, file: 'x' });
      assert.deepEqual([r.ok, r.code], [false, 'NOT_OPEN_TAB'], ch);
      assert.match(r.error, /허용되지 않은 요청입니다/);
    }
    let r = await callAs('panel', 'doc:restore', {});
    assert.equal(r.code, 'NOT_OPEN_TAB', '경로가 없으면 거절');
    // AI 대상이면 된다
    r = await callAs('panel', 'doc:versions', { path: docPath });
    assert.equal(r.ok, true, r.error);
    // 문서 탭 뷰: 그 뷰의 문서는 된다, 다른 문서(AI 대상도 아님)는 거절
    r = await callAs(pdf, 'doc:versions', { path: pdf });
    assert.equal(r.ok, true, r.error);
    const hash = sha(docxPath);
    r = await callAs(pdf, 'doc:restore', { path: docxPath });
    assert.equal(r.code, 'NOT_OPEN_TAB');
    assert.equal(sha(docxPath), hash, '되돌리지 않는다');
    // 홈 뷰는 문서를 품지 않는다
    r = await callAs('home', 'doc:versions', { path: pdf });
    assert.equal(r.code, 'NOT_OPEN_TAB');
    // 탭을 닫으면 그 뷰도 닫힌다
    const pdfWc = docWc(pdf);
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(pdf).id)).ok, true);
    assert.equal(pdfWc.isDestroyed(), true);
    assert.equal(tabOf(pdf), null);
  });

  await t.test('문서 탭의 직접 편집 → 같은 자리 편집 탭 (탭 수 그대로)', async () => {
    const p = join(docs, '자리 시험.docx');
    await createBlankDocument(p);
    const opened = await call('doc:open', p);
    assert.equal(opened.value.mode, 'doc');
    const before = tabList();
    const at = before.tabs.findIndex((x) => x.id === opened.value.tabId);
    const docView = docWc(p);
    const r = await call('doc:edit');
    assert.equal(r.ok, true, r.error);
    const after = tabList();
    assert.equal(after.tabs.length, before.tabs.length, '탭 수 그대로');
    assert.equal(after.tabs[at].type, 'editor');
    assert.equal(after.tabs[at].path, p);
    assert.equal(after.active, after.tabs[at].id);
    assert.equal(docView.isDestroyed(), true, '문서 탭 뷰는 닫힌다');
    assert.ok(editorOf(/자리 시험\.docx - 문서 편집/));
    assert.equal(lastTarget()?.path, p);
  });

  await t.test('단축키: 탭 줄·홈·문서·패널·편집 어디에 초점이 있어도 Ctrl+Tab/Ctrl+1~9/Ctrl+T가 같은 탭 창에서 동작', async () => {
    const pdf = join(docs, '단축키 시험.pdf');
    writeFileSync(pdf, makePdf(1));
    assert.equal((await call('doc:open', pdf)).ok, true);
    const edWc = editorOf(/자리 시험\.docx - 문서 편집/).wc;
    const docV = docWc(pdf);
    const press = (wc, code, { shift = false, alt = false } = {}) => {
      const e = { prevented: false, preventDefault() { this.prevented = true; } };
      wc.emit('before-input-event', e, { type: 'keyDown', code, control: true, shift, alt, meta: false });
      return e.prevented;
    };
    const ids = () => tabList().tabs.map((x) => x.id);
    const goHome = async () => { assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true); assert.equal(tabList().active, 'home'); };
    const wins = fake.windows.length;
    assert.ok(ids().length >= 3, '홈 + 편집 탭 + 문서 탭');
    // Ctrl+Tab / Ctrl+Shift+Tab: 탭 줄·홈·패널·문서·편집 어디서든, 끝에서 처음으로 돈다
    await goHome();
    assert.equal(press(panelWc, 'Tab'), true);
    assert.equal(tabList().active, ids()[1], '패널에서 Ctrl+Tab → 다음 탭');
    assert.equal(press(edWc, 'Tab', { shift: true }), true);
    assert.equal(tabList().active, 'home', '편집 뷰에서 Ctrl+Shift+Tab → 이전 탭');
    assert.equal(press(tabstripWc, 'Tab', { shift: true }), true);
    assert.equal(tabList().active, ids().at(-1), '탭 줄에서 Ctrl+Shift+Tab → 처음에서 마지막으로');
    assert.equal(press(docV, 'Tab'), true);
    assert.equal(tabList().active, 'home', '문서 뷰에서 Ctrl+Tab → 마지막에서 처음으로');
    assert.equal(press(homeWc, 'Tab'), true);
    assert.equal(tabList().active, ids()[1]);
    // Ctrl+1~9: 탭 줄·홈·패널·문서 뷰에서 → 그 번째 탭 (9는 마지막)
    assert.equal(press(docV, 'Digit9'), true);
    assert.equal(tabList().active, ids().at(-1), 'Ctrl+9 → 마지막 탭');
    assert.equal(press(panelWc, 'Digit1'), true);
    assert.equal(tabList().active, 'home');
    assert.equal(press(tabstripWc, 'Digit2'), true);
    assert.equal(tabList().active, ids()[1]);
    // 편집 뷰는 Ctrl+1~9를 가로채지 않는다 (셀 서식 Ctrl+1 등 편집기 기본 동작)
    const before = tabList().active;
    assert.equal(press(edWc, 'Digit1'), false, '편집 뷰 Ctrl+1은 통과');
    assert.equal(press(edWc, 'Digit9'), false);
    assert.equal(tabList().active, before);
    // 없는 번째 탭은 무시 (막기만 한다)
    if (ids().length < 8) {
      assert.equal(press(homeWc, 'Digit8'), true);
      assert.equal(tabList().active, before);
    }
    // Ctrl+T: 홈 탭
    assert.equal(press(docV, 'KeyT'), true);
    assert.equal(tabList().active, 'home');
    await callAs('tabstrip', 'tabs:activate', ids().at(-1));
    assert.equal(press(edWc, 'KeyT'), true);
    assert.equal(tabList().active, 'home', '편집 뷰 Ctrl+T도 홈');
    // Ctrl+Alt 조합은 탭을 건드리지 않는다 (Ctrl+N 새 탭 창은 아래 여러 탭 창 시험에서)
    assert.equal(press(panelWc, 'Tab', { alt: true }), false);
    assert.equal(tabList().active, 'home');
    assert.equal(fake.windows.length, wins, '새 창을 만들지 않는다');
  });

  await t.test('Ctrl+W: 문서 탭은 닫고, 홈 탭에서는 아무 일도 없다', async () => {
    const pdf = join(docs, '단축키 시험.pdf');
    const press = (wc, code) => {
      const e = { prevented: false, preventDefault() { this.prevented = true; } };
      wc.emit('before-input-event', e, { type: 'keyDown', code, control: true, shift: false, alt: false, meta: false });
      return e.prevented;
    };
    const docV = docWc(pdf);
    const n = tabList().tabs.length;
    // 홈 탭: 어느 뷰에서 눌러도 그대로
    assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true);
    for (const wc of [homeWc, panelWc, tabstripWc]) {
      assert.equal(press(wc, 'KeyW'), true);
      assert.equal(tabList().tabs.length, n);
      assert.equal(tabList().active, 'home');
    }
    // 문서 탭: 그 탭이 닫히고 뷰도 닫힌다
    assert.equal((await callAs('tabstrip', 'tabs:activate', tabOf(pdf).id)).ok, true);
    assert.equal(press(docV, 'KeyW'), true);
    assert.ok(await until(() => tabOf(pdf) === null, 5000), '문서 탭이 닫힌다');
    assert.equal(docV.isDestroyed(), true);
    assert.equal(tabList().tabs.length, n - 1);
  });

  await t.test('겹친 편집 탭 열기: 두 문서 탭의 직접 편집을 동시에 → 각자 제자리 편집 탭, 중복 없음', async () => {
    const a = join(docs, '동시 가.docx');
    const b = join(docs, '동시 나.docx');
    await createBlankDocument(a);
    await createBlankDocument(b);
    const ta = (await call('doc:open', a)).value.tabId;
    const tb = (await call('doc:open', b)).value.tabId;
    const before = tabList();
    const ia = before.tabs.findIndex((x) => x.id === ta);
    const ib = before.tabs.findIndex((x) => x.id === tb);
    const [ra, rb] = await Promise.all([callAs(a, 'doc:edit'), callAs(b, 'doc:edit')]);
    assert.equal(ra.ok, true, ra.error);
    assert.equal(rb.ok, true, rb.error);
    const after = tabList();
    assert.equal(after.tabs.length, before.tabs.length, '탭 수 그대로');
    assert.deepEqual([after.tabs[ia].type, after.tabs[ia].path], ['editor', a]);
    assert.deepEqual([after.tabs[ib].type, after.tabs[ib].path], ['editor', b]);
    for (const p of [a, b]) {
      assert.equal(after.tabs.filter((x) => x.path === p).length, 1, p);
      assert.equal(docWc(p), null, '문서 탭 뷰는 닫힌다');
    }
    assert.ok(editorOf(/동시 가\.docx - 문서 편집/) && editorOf(/동시 나\.docx - 문서 편집/));
    cur = b;
  });

  await t.test('새 문서 AI 시작 → 패널에 ai:event start가 다른 ai:event보다 먼저', async () => {
    const n0 = panelWc.sent.length;
    const r = await call('doc:new', { format: 'docx', folder: docs, name: 'AI 새 문서', start: 'ai', request: '회의록 양식 만들어줘' });
    assert.equal(r.ok, true, r.error);
    assert.ok(r.value.chat?.chatId, JSON.stringify(r.value.chatError));
    const chatId = r.value.chat.chatId;
    assert.ok(await until(() => lastSent('ai:event', (d) => d.type === 'end' && d.chatId === chatId)));
    const evs = panelWc.sent.slice(n0).filter((s) => s.channel === 'ai:event' && s.data.chatId === chatId).map((s) => s.data);
    assert.equal(evs[0].type, 'start');
    assert.equal(evs[0].path, join(docs, 'AI 새 문서.docx'));
    assert.equal(evs[0].prompt, '회의록 양식 만들어줘');
    assert.ok(evs.length >= 2);
    assert.equal(lastTarget()?.path, join(docs, 'AI 새 문서.docx'), '그 탭 창의 AI 대상 = 새 문서');
    assert.equal(tabOf(join(docs, 'AI 새 문서.docx'))?.type, 'doc');
  });

  await t.test('IPC 표: 편집 창 ↔ 주 창 채널은 서로 거절, 하위 틀 거절, 열린 채널은 누구에게나 답', async () => {
    // 편집 창 → 주 창 채널
    let r = await fake.invokeFrom(officeWin.wc, 'ai:send', { prompt: 'x' });
    assert.deepEqual([r.ok, r.error], [false, '허용되지 않은 요청입니다']);
    // 주 창 → 편집 창 채널
    r = await call('editor:get-document');
    assert.deepEqual([r.ok, r.error], [false, '허용되지 않은 요청입니다']);
    // 홈 뷰 안의 하위 틀
    r = await fake.invokeAs(fake.evOf(homeWc, { url: 'file:///sub.html' }), 'doc:versions');
    assert.equal(r.ok, false);
    // AI 패널은 홈 채널(doc:open)을, 탭 줄은 편집·AI 채널을 못 부른다 (Section 1 #3)
    r = await callAs('panel', 'doc:open', docPath);
    assert.deepEqual([r.ok, r.error], [false, '허용되지 않은 요청입니다']);
    for (const ch of ['editor:save', 'ai:send']) {
      r = await callAs('tabstrip', ch, { prompt: 'x' });
      assert.deepEqual([r.ok, r.error], [false, '허용되지 않은 요청입니다'], ch);
    }
    // 어느 창에도 속하지 않은 화면도 번역 사전(열린 채널)은 받는다
    const stray = new fake.WebContentsView({}).webContents;
    assert.ok(fake.emitSync(stray, 'i18n:get'));
    assert.notEqual(fake.emitSync(stray, 'editor:i18n'), undefined);
  });

  await t.test('탐색기에서 연 옛 형식(csv) → 새 형식(xlsx)으로 바꿔 편집 창, 첫 저장은 원본 옆 새 파일 (원본 그대로)', async () => {
    const csv = join(docs, '명단.csv');
    writeFileSync(csv, '이름,점수\n홍길동,90\n김영희,85\n', 'utf8');
    const csvBefore = readFileSync(csv);
    // 탐색기에서 연 문서는 설정 '문서를 열면: 편집'일 때 편집 탭으로 (이 시험의 기본은 미리보기)
    assert.equal((await call('settings:set', { openAction: 'edit' })).ok, true);
    process.env.OFFICEKIT_DOC_WINDOW = '1';
    try { fake.app.emit('second-instance', {}, [process.execPath, csv], docs); } finally { delete process.env.OFFICEKIT_DOC_WINDOW; process.env.OFFICEKIT_DOC_WINDOW = '0'; }
    assert.ok(await until(() => editorOf(/명단\.csv → \.xlsx - 표 편집/), 20_000), fake.windows.map((w) => w.getTitle()).join(' | '));
    const w = editorOf(/명단\.csv/);
    const c = (ch, ...a) => fake.invokeFrom(w.wc, ch, ...a);
    const d = await c('editor:get-document');
    assert.equal(d.ok, true, d.error);
    assert.equal(d.value.ext, 'xlsx');
    assert.equal(d.value.convertedFrom, '명단.csv');
    // 원본 csv를 탐색기에서 다시 열면 바꿔 연 그 탭으로 (새 편집 뷰·탭 없이)
    const nViews = editorViews().length;
    const nTabs = tabList().tabs.length;
    assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true);
    fake.app.emit('second-instance', {}, [process.execPath, csv], docs);
    assert.ok(await until(() => tabList().active !== 'home'), '그 탭으로 가야 한다');
    assert.equal(tabList().tabs.find((x) => x.id === tabList().active)?.title, '명단.csv');
    await sleep(200);
    assert.equal(editorViews().length, nViews);
    assert.equal(tabList().tabs.length, nTabs);
    const temp = (await import('node:path')).dirname(join(ud, 'converted'));
    void temp;
    // 첫 저장: 원본 옆 '명단.xlsx'를 고르게 하고, 그 파일을 편집하는 창이 된다
    const target = join(docs, '명단.xlsx');
    fake.saveDialogAnswers.push({ canceled: false, filePath: target });
    w.wc.emit('input-event', {}, { type: 'mouseDown' });
    const r = await c('editor:save', new Uint8Array(d.value.bytes), { reason: 'user' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.value.name, '명단.xlsx');
    assert.ok(existsSync(target));
    assert.deepEqual(readFileSync(csv), csvBefore, '원본 csv는 그대로');
    assert.equal(w.win.getTitle(), '명단.xlsx - 표 편집 (Office Kit)');
    assert.deepEqual(readdirSync(join(ud, 'converted')), [], '임시 사본은 지운다');
    // 두 번째 저장은 그 파일에 바로
    const r2 = await c('editor:save', new Uint8Array(readFileSync(target)), { reason: 'user' });
    assert.equal(r2.ok, true, r2.error);
    assert.equal(r2.value.name, '명단.xlsx');
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(target).id)).ok, true);
  });

  await t.test('탐색기 인터넷 docx → 문서 탭(편집 뷰 없음), zone3Preview, doc:edit UNTRUSTED, openInApp은 연결 프로그램으로 — 의도한 변경', { skip: process.platform !== 'win32' && 'Zone.Identifier' }, async () => {
    const net = join(docs, '인터넷 자료.docx');
    const net2 = join(docs, '인터넷 자료 둘.docx');
    for (const p of [net, net2]) {
      copyFileSync(docxPath, p);
      writeFileSync(`${p}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=3\r\n');
    }
    // 설정 '문서를 열면: 편집' + 탐색기 문서 창 방식 → 예전에는 편집기로 바로 열렸다
    assert.equal((await call('settings:set', { openAction: 'edit' })).ok, true);
    const nViews = editorViews().length;
    process.env.OFFICEKIT_DOC_WINDOW = '1';
    try {
      fake.app.emit('second-instance', {}, [process.execPath, net], docs);
      assert.ok(await until(() => tabOf(net) && docWc(net)), '탭이 열려야 한다');
    } finally { process.env.OFFICEKIT_DOC_WINDOW = '0'; }
    await sleep(200);
    assert.equal(tabOf(net).type, 'doc', '편집 탭이 아니라 문서 탭(미리보기)');
    assert.equal(tabOf(net).zone3, true, '탭 줄에 인터넷 파일 표시');
    assert.equal(editorViews().length, nViews, '편집 뷰를 만들지 않는다 (editor:* 채널이 열릴 길 없음)');
    assert.equal(lastSent('doc:opened', (d) => d?.path === net)?.zone3Preview, true);
    // '직접 편집'은 거절
    const ed = await callAs(net, 'doc:edit');
    assert.equal(ed.ok, false);
    assert.equal(ed.code, 'UNTRUSTED');
    assert.equal(editorViews().length, nViews);
    assert.equal(tabOf(net).type, 'doc');
    // 홈에서 연 인터넷 문서도 문서 탭, 같은 문서를 다시 열면 그 탭으로
    const r2 = await callAs('home', 'doc:open', net2);
    assert.equal(r2.ok, true, r2.error);
    assert.equal(r2.value.mode, 'doc');
    assert.equal(editorViews().length, nViews);
    const again = await callAs('home', 'doc:open', net);
    assert.equal(again.value.mode, 'reused');
    assert.equal(again.value.tabId, tabOf(net).id);
    // 평범한 문서는 zone3Preview 아님
    assert.equal(lastSent('doc:opened', (d) => d?.path === docPath)?.zone3Preview, false);
    // 'Windows 프로그램으로 열기' → 연결 프로그램으로 (자동화 COM 아님). 실제 프로그램은 띄우지 않는다
    const req = createRequire(import.meta.url);
    const cp = req('node:child_process');
    const { syncBuiltinESMExports } = req('node:module');
    const spawned = [];
    const spawn0 = cp.spawn;
    cp.spawn = (file, args) => {
      spawned.push([file, ...(args ?? [])].join(' '));
      const child = new EventEmitter();
      child.unref = () => {};
      setImmediate(() => child.emit('spawn'));
      return child;
    };
    syncBuiltinESMExports();
    try {
      const o = await callAs(net, 'doc:openInApp');
      assert.equal(o.ok, true, o.error);
      assert.equal(o.value.via, 'shell');
    } finally {
      cp.spawn = spawn0;
      syncBuiltinESMExports();
    }
    assert.equal(spawned.length, 1, spawned.join(' | '));
    assert.ok(spawned[0].includes(net), spawned[0]);
    // 차단을 해제한 뒤 같은 문서를 다시 열면 그 탭으로 가고 띠가 사라진다 ('직접 편집'이 돌아온다, 새 편집 뷰 없음)
    rmSync(`${net}:Zone.Identifier`, { force: true });
    assert.equal((await callAs('tabstrip', 'tabs:home')).ok, true);
    const nOpened = docWc(net).sent.filter((s) => s.channel === 'doc:opened').length;
    const re = await callAs('home', 'doc:open', net);
    assert.equal(re.ok, true, re.error);
    assert.equal(re.value.mode, 'reused');
    assert.equal(re.value.tabId, tabOf(net).id);
    assert.equal(tabList().active, tabOf(net).id);
    assert.ok(docWc(net).sent.filter((s) => s.channel === 'doc:opened').length > nOpened, '문서 정보를 다시 보낸다');
    assert.equal(lastSent('doc:opened', (d) => d?.path === net)?.zone3Preview, false);
    assert.equal(tabOf(net).zone3, false, '탭 줄의 인터넷 표시도 사라진다');
    assert.equal(editorViews().length, nViews);
    // 인터넷 PDF는 편집 대상이 아니므로 띠 없음
    const netPdf = join(docs, '인터넷 문서.pdf');
    writeFileSync(netPdf, makePdf(1));
    writeFileSync(`${netPdf}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=3\r\n');
    assert.equal((await callAs('home', 'doc:open', netPdf)).ok, true);
    assert.equal(lastSent('doc:opened', (d) => d?.path === netPdf)?.zone3Preview, false);
    assert.equal(tabOf(netPdf).zone3, false);
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(netPdf).id)).ok, true);
    for (const p of [net, net2]) assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(p).id)).ok, true);
  });

  // ── 여러 탭 창 (D1 · Section 1 #1): Ctrl+N, AI 이벤트는 대화를 시작한 창의 패널로, AI 대상은 창마다 ──
  const pressCtrl = (wc, code) => {
    const e = { prevented: false, preventDefault() { this.prevented = true; } };
    wc.emit('before-input-event', e, { type: 'keyDown', code, control: true, shift: false, alt: false, meta: false });
    return e.prevented;
  };
  /** fromWc에서 Ctrl+N → 새 탭 창과 그 창의 홈·패널·탭 줄 뷰 */
  const newTabWindow = async (fromWc) => {
    const nWin = fake.windows.length;
    const nView = fake.views.length;
    assert.equal(pressCtrl(fromWc, 'KeyN'), true, 'Ctrl+N은 가로챈다');
    assert.ok(await until(() => fake.windows.length > nWin), '새 탭 창이 열려야 한다');
    const fresh = fake.views.slice(nView).map((v) => v.webContents);
    const w = {
      win: fake.windows.at(-1),
      home: fresh.find((wc) => wc.mainFrame.url.endsWith('?view=home')),
      panel: fresh.find((wc) => wc.mainFrame.url.endsWith('?view=panel')),
      tabstrip: fresh.find((wc) => /tabstrip\.html$/.test(wc.mainFrame.url)),
    };
    assert.ok(w.home && w.panel && w.tabstrip, '새 탭 창의 홈·패널·탭 줄 뷰');
    return w;
  };
  const tabsOf = (w) => w.tabstrip.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [], active: null };
  /** 그 패널 뷰가 받은 ai:event의 chatId들 */
  const chatIdsOf = (wc) => new Set(wc.sent.filter((s) => s.channel === 'ai:event' && s.data?.chatId).map((s) => s.data.chatId));
  const sameFile = (a, b) => resolve(a).toLowerCase() === resolve(b).toLowerCase();
  const pdfB = join(docs, '둘째 창.pdf');
  let win2 = null;

  await t.test('탭 창 2개: 각 창 AI 답은 자기 창 패널에만 (D1, Section 1 #1)', async () => {
    tw0.emit('focus');
    const before = fake.windows.length;
    win2 = await newTabWindow(homeWc);
    assert.equal(fake.windows.length, before + 1, '탭 창 하나만 더');
    assert.equal(tw0.isDestroyed(), false);
    assert.equal(win2.win.isDestroyed(), false);
    // 새 창은 최근 창에서 30px 어긋나게, 홈 탭만
    assert.equal(win2.win.opts.x, tw0.getBounds().x + 30);
    assert.equal(win2.win.opts.y, tw0.getBounds().y + 30);
    assert.deepEqual(tabsOf(win2).tabs.map((x) => x.type), ['home']);
    assert.equal(tabsOf(win2).active, 'home');
    // 창1: 문서 A(보고서.hwpx 편집 탭), 창2: 문서 B(PDF 문서 탭)
    assert.equal((await callAs('home', 'doc:open', docPath)).ok, true);
    writeFileSync(pdfB, makePdf(1));
    assert.equal((await fake.invokeFrom(win2.home, 'doc:open', pdfB)).ok, true);
    assert.ok(tabsOf(win2).tabs.some((x) => x.path && sameFile(x.path, pdfB)), '창2에 문서 B 탭');
    assert.equal(tabOf(pdfB), null, '창1에는 문서 B 탭이 없다');
    const r1 = await callAs('panel', 'ai:send', { prompt: '요약해줘', editorOk: true });
    assert.equal(r1.ok, true, r1.error);
    const r2 = await fake.invokeFrom(win2.panel, 'ai:send', { prompt: '요약해줘' });
    assert.equal(r2.ok, true, r2.error);
    assert.ok(await endOf(r1.value.chatId));
    assert.ok(await endOf(r2.value.chatId));
    assert.equal(chatIdsOf(panelWc).has(r1.value.chatId), true);
    assert.equal(chatIdsOf(panelWc).has(r2.value.chatId), false, '창2 대화의 이벤트는 창1 패널에 가지 않는다');
    assert.deepEqual([...chatIdsOf(win2.panel)], [r2.value.chatId], '창2 패널은 자기 대화만');
    for (const s of win2.panel.sent.filter((x) => x.channel === 'ai:event')) assert.ok(sameFile(s.data.path, pdfB));
    // D14는 창을 넘어서도: 창1이 고치는 중인 문서 A는 창2 대화의 재료가 될 수 없다
    const held = await callAs('panel', 'ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    try {
      const m = await fake.invokeFrom(win2.panel, 'ai:materials');
      assert.equal(m.ok, true, m.error);
      assert.equal(m.value.open.find((x) => sameFile(x.path, docPath))?.busy, true, '창2 재료 목록에도 고치는 중으로');
      const r = await fake.invokeFrom(win2.panel, 'ai:send', { prompt: '참고해줘', target: { read: [docPath] } });
      assert.equal(r.code, 'MATERIAL_BUSY');
      assert.equal(chatIdsOf(win2.panel).has(held.value.chatId), false);
    } finally {
      globalThis.__okRelease();
    }
    assert.ok(await endOf(held.value.chatId));
  });

  await t.test('AI 답이 오는 중 그 탭 창을 닫으면 (재료만 읽는 대화도) 한 번 묻고 멈춘 뒤 닫는다, 이벤트는 다른 창으로 새지 않고 대화 정리는 끝까지 — 의도한 변경 (Review Focus, RT4)', async () => {
    const r = await fake.invokeFrom(win2.panel, 'ai:send', { prompt: '[hold] 요약해줘' });
    assert.equal(r.ok, true, r.error);
    const chatId = r.value.chatId;
    const pidDir = join(ud, 'com-pids', chatId);
    assert.equal(existsSync(pidDir), true, '대화마다 PID 폴더');
    const warns = [];
    const warn0 = console.warn;
    console.warn = (...a) => { warns.push(a); };
    try {
      const n = fake.dialogs.length;
      fake.dialogAnswers.push(0); // 멈추고 닫기: 창 패널의 대화는 창과 함께 멈춘다
      win2.win.close();
      assert.ok(await until(() => win2.win.isDestroyed()), '창2가 닫혀야 한다');
      assert.equal(fake.dialogs.length, n + 1, '한 번만 묻는다');
      assert.equal(fake.dialogs[n].message, 'AI 작업 중인 문서가 있습니다.', '재료만 읽는 대화도 묻는다');
      assert.equal(win2.panel.isDestroyed(), true);
      assert.ok(await until(() => !existsSync(pidDir)), 'PID 폴더 정리');
      await sleep(100);
    } finally {
      console.warn = warn0;
      globalThis.__okRelease?.();
    }
    assert.ok(warns.filter((a) => a[0] === '[ai] 패널이 닫혀 이벤트를 버렸습니다' && a[1] === chatId).length <= 1, '대화마다 많아야 한 번만 알린다');
    assert.equal(chatIdsOf(panelWc).has(chatId), false, '다른 창 패널로 새지 않는다');
    // 대화가 정리됐다: 창1에서 문서 B를 열면 AI가 쓰는 중이 아니다
    assert.equal((await callAs('home', 'doc:open', pdfB)).ok, true);
    assert.equal(lastSent('doc:opened', (d) => d?.path && sameFile(d.path, pdfB))?.aiBusy, false);
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(pdfB).id)).ok, true);
  });

  let win3 = null;
  await t.test('창별 AI 대상: 창1은 편집 탭 A, 창2는 홈 → 창2 ai:materials edit null', async () => {
    tw0.emit('focus');
    win3 = await newTabWindow(panelWc);
    assert.equal((await callAs('home', 'doc:open', docPath)).ok, true);
    const m1 = await callAs('panel', 'ai:materials');
    assert.equal(m1.ok, true, m1.error);
    assert.ok(sameFile(m1.value.edit.path, docPath), '창1 대상 = 편집 탭 A');
    const m2 = await fake.invokeFrom(win3.panel, 'ai:materials');
    assert.equal(m2.ok, true, m2.error);
    assert.equal(m2.value.edit, null, '창2는 홈 → 대상 없음');
    const r = await fake.invokeFrom(win3.panel, 'ai:send', { prompt: '요약해줘' });
    assert.equal(r.ok, false, '홈 탭 창은 AI 대상이 없다');
  });

  await t.test('메뉴 명령은 마지막으로 초점 받은 탭 창의 홈으로', async () => {
    const settingsItem = fake.menu.flatMap((m) => m.submenu ?? []).find((i) => i.label === '설정');
    assert.ok(settingsItem, '설정 메뉴');
    const menus = (wc) => wc.sent.filter((s) => s.channel === 'menu').length;
    const [a0, b0] = [menus(homeWc), menus(win3.home)];
    tw0.emit('focus');
    settingsItem.click();
    assert.equal(menus(homeWc), a0 + 1, '창1이 최근 → 창1 홈');
    assert.equal(menus(win3.home), b0);
    assert.equal(tabList().active, 'home');
    win3.win.emit('focus');
    settingsItem.click();
    assert.equal(menus(win3.home), b0 + 1, '창2가 최근 → 창2 홈');
    assert.equal(menus(homeWc), a0 + 1);
    // 시험 뒤 창2를 닫는다 (아래 탭 창 닫기 시험은 창1 기준)
    win3.win.close();
    assert.ok(await until(() => win3.win.isDestroyed()));
    tw0.emit('focus');
  });

  await t.test('PDF 탭 둘(창 2개) + 셋째 미리보기: 앞 둘의 사본이 남는다 (Review Focus)', async () => {
    const [pa, pb, pc] = ['미리 가.pdf', '미리 나.pdf', '미리 다.pdf'].map((n) => join(docs, n));
    for (const p of [pa, pb, pc]) writeFileSync(p, makePdf(1));
    const fileOf = (r) => fileURLToPath(r.value.url.split('#')[0]);
    const show = async (p) => {
      const r = await callAs(p, 'doc:preview');
      assert.equal(r.ok, true, r.error);
      assert.equal(r.value.kind, 'pdf');
      return r;
    };
    assert.equal((await callAs('home', 'doc:open', pa)).ok, true);
    const ra = await show(pa);
    const w = await newTabWindow(homeWc);
    assert.equal((await fake.invokeFrom(w.home, 'doc:open', pb)).ok, true);
    const rb = await show(pb);
    tw0.emit('focus');
    assert.equal((await callAs('home', 'doc:open', pc)).ok, true);
    const rc = await show(pc);
    for (const r of [ra, rb, rc]) assert.equal(existsSync(fileOf(r)), true, fileOf(r));
    // 다시 그린 미리보기는 그 문서의 옛 사본만 지운다
    const ra2 = await show(pa);
    assert.equal(existsSync(fileOf(ra)), false, '가의 옛 사본');
    for (const r of [ra2, rb, rc]) assert.equal(existsSync(fileOf(r)), true, fileOf(r));
    // 탭을 닫은 문서의 사본은 다음 미리보기 때 지운다
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(pa).id)).ok, true);
    const rc2 = await show(pc);
    assert.equal(existsSync(fileOf(ra2)), false, '닫은 가의 사본');
    for (const r of [rb, rc2]) assert.equal(existsSync(fileOf(r)), true, fileOf(r));
    assert.equal((await callAs('tabstrip', 'tabs:close', tabOf(pc).id)).ok, true);
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    tw0.emit('focus');
  });

  await t.test('탐색기 시작 세션은 lastTabSession을 덮지 않고, 홈의 session:restoreLast가 지난 세션을 연다 (D4·D11)', async () => {
    const readS = () => JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));
    assert.ok(await until(() => readS().tabSession?.explorer === true, 3000), '탐색기 시작 세션도 지금 세션으로 적는다');
    assert.deepEqual(readS().lastTabSession, lastSession);
    // 확정해도(끝내기·마지막 창 닫기) 지난 세션은 그대로
    mainMod.sessionHooks.finalize();
    assert.equal(readS().tabSession.finalized, true);
    assert.deepEqual(readS().lastTabSession, lastSession);
    mainMod.sessionHooks.resume();
    const info = await callAs('home', 'session:info');
    assert.equal(info.ok, true, info.error);
    assert.deepEqual(info.value, { hasLast: true, count: 2 });
    const n = fake.windows.length;
    const r = await callAs('home', 'session:restoreLast');
    assert.equal(r.ok, true, r.error);
    assert.ok(await until(() => fake.windows.length === n + 1), '지난 세션 창');
    const win = fake.windows.at(-1);
    assert.deepEqual([win.opts.x, win.opts.y], [300, 200], '창 자리는 세션 값');
    const strip = win.contentView.children.map((v) => v.webContents).find((wc) => /tabstrip\.html$/.test(wc?.mainFrame.url ?? ''));
    const list = () => strip.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [] };
    // 이미 열린 문서(보고서.hwpx)는 건너뛰고, 활성 탭만 불러온다
    assert.ok(await until(() => list().tabs.find((x) => x.path === lastPdf)?.loaded), '활성 탭(지난 세션.pdf)은 불러온다');
    assert.deepEqual(list().tabs.map((x) => [x.type, x.path && basename(x.path)]), [['home', null], ['doc', '지난 세션.pdf']]);
    assert.equal((await callAs('home', 'session:info')).value.hasLast, false, '한 번 열면 단추를 감춘다');
    assert.deepEqual(readS().lastTabSession, lastSession);
    // 지난 세션을 다시 열었으면 평소 세션이 된다: 다음 확정이 지금 탭들을 지난 세션으로 쓴다 (RT3 — 안 그러면 다시 시작할 때 탭을 잃는다)
    assert.ok(await until(() => readS().tabSession?.explorer === false, 3000), '평소 세션으로 다시 적는다');
    mainMod.sessionHooks.finalize();
    assert.equal(readS().lastTabSession.finalized, true);
    assert.equal(readS().lastTabSession.explorer, false);
    assert.ok(readS().lastTabSession.windows.some((x) => x.tabs.some((y) => sameFile(y.path, lastPdf))), '지금 탭들이 지난 세션');
    mainMod.sessionHooks.resume();
    win.close();
    assert.ok(await until(() => win.isDestroyed()));
    tw0.emit('focus');
  });

  // ── 닫기 흐름: 탭·창·끝내기 차례 확인, AI 작업 중 확인 (D17) ──
  const AI_TAB_MSG = 'AI가 이 문서를 고치는 중입니다.';
  const AI_WIN_MSG = 'AI 작업 중인 문서가 있습니다.';
  const AI_DETAIL = '멈추면 지금까지 고친 내용은 남고, 버전 목록에서 대화 전으로 되돌릴 수 있어요.';
  const STOP_BUTTONS = ['멈추고 닫기', '취소'];
  // 세션 연결(Task 11)이 채울 자리: 부른 차례를 적는다 (finalize는 그때 살아 있던 탭 창 수도)
  const hookCalls = [];
  const liveWindows = () => fake.windows.filter((w) => !w.isDestroyed()).length;
  Object.assign(mainMod.sessionHooks, {
    finalize: () => { hookCalls.push({ name: 'finalize', open: liveWindows() }); },
    resume: () => { hookCalls.push({ name: 'resume' }); },
    removeWindow: (id) => { hookCalls.push({ name: 'removeWindow', id }); },
  });
  /** 새 한글 문서를 만들어 그 탭 창 홈에서 연다 → { path, wc: 편집 뷰 } */
  const openEditorIn = async (w, name) => {
    const p = join(docs, name);
    await createBlankDocument(p, { title: name });
    const before = new Set(editorViews());
    const r = await fake.invokeFrom(w.home, 'doc:open', p);
    assert.equal(r.ok, true, r.error);
    assert.ok(await until(() => editorViews().some((v) => !before.has(v))), `${name} 편집 뷰`);
    return { path: p, wc: editorViews().find((v) => !before.has(v)).webContents };
  };
  const tabIn = (w, p) => tabsOf(w).tabs.find((x) => x.path && sameFile(x.path, p)) ?? null;
  const asked = () => fake.dialogs.length;
  const quitEvent = () => ({ prevented: false, preventDefault() { this.prevented = true; } });

  await t.test('탭 창 닫기: 저장 안 한 탭을 탭 차례로 묻고, 둘째에서 취소하면 첫째는 닫히고 창·나머지는 그대로 — 의도한 변경', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, '닫기 가.hwpx');
    const b = await openEditorIn(w, '닫기 나.hwpx');
    fake.replies.set(a.wc.id, true);
    fake.replies.set(b.wc.id, true);
    const n = asked();
    fake.dialogAnswers.push(1, 2); // 가: 저장 안 함, 나: 취소
    w.win.close();
    assert.ok(await until(() => asked() >= n + 2));
    await sleep(150);
    assert.equal(asked(), n + 2, '탭마다 한 번씩만');
    assert.match(fake.dialogs[n].message, /닫기 가\.hwpx/, '탭 차례: 가 먼저');
    assert.match(fake.dialogs[n + 1].message, /닫기 나\.hwpx/);
    assert.equal(a.wc.isDestroyed(), true, '확인한 탭은 바로 닫힌다');
    assert.equal(tabIn(w, a.path), null);
    assert.equal(b.wc.isDestroyed(), false, '취소한 탭은 그대로');
    assert.ok(tabIn(w, b.path));
    assert.equal(w.win.isDestroyed(), false, '창도 그대로');
    assert.equal(hookCalls.length, 0, '닫히지 않은 창은 세션에서 빼지 않는다');
    // 정리: 깨끗하면 묻지 않고 닫힌다. 다른 탭 창이 남아 있으니 세션에서 이 창만 뺀다
    fake.replies.set(b.wc.id, false);
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    assert.deepEqual(hookCalls.splice(0), [{ name: 'removeWindow', id: w.win.id }]);
    tw0.emit('focus');
  });

  await t.test('AI가 고치는 중인 문서 탭 닫기 → 멈추고 닫기/취소 (D17)', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, 'AI 가.hwpx');
    const held = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    const chatId = held.value.chatId;
    const pidDir = join(ud, 'com-pids', chatId);
    try {
      assert.ok(sameFile(globalThis.__okChats.at(-1).filePath, a.path), '고칠 문서 = 가');
      let n = asked();
      fake.dialogAnswers.push(1); // 취소
      const r0 = await fake.invokeFrom(w.tabstrip, 'tabs:close', tabIn(w, a.path).id);
      assert.equal(r0.ok, true, r0.error);
      assert.equal(r0.value, false);
      assert.equal(asked(), n + 1);
      const d = fake.dialogs[n];
      assert.equal(d.message, AI_TAB_MSG);
      assert.equal(d.detail, AI_DETAIL);
      assert.deepEqual(d.buttons, STOP_BUTTONS);
      assert.equal(d.cancelId, 1);
      assert.equal(d.parent, w.win, '부모 = 그 탭 창');
      assert.equal(a.wc.isDestroyed(), false, '취소 → 탭 그대로');
      assert.ok(tabIn(w, a.path));
      assert.equal(existsSync(pidDir), true, '취소 → 대화 그대로');
      assert.equal(lastSent('ai:event', (x) => x.type === 'end' && x.chatId === chatId), undefined);
      // 멈추고 닫기 → 대화 정지(숨은 Office 정리 포함) → 저장 확인(깨끗하면 묻지 않음) → 닫기
      n = asked();
      fake.dialogAnswers.push(0);
      const r1 = await fake.invokeFrom(w.tabstrip, 'tabs:close', tabIn(w, a.path).id);
      assert.equal(r1.value, true);
      assert.equal(asked(), n + 1, '깨끗한 편집 탭은 저장을 묻지 않는다');
      assert.ok(await endOf(chatId), '대화가 멈췄다');
      assert.ok(await until(() => !existsSync(pidDir)), 'PID 폴더 정리');
      assert.ok(await until(() => a.wc.isDestroyed()), '탭이 닫힌다');
      assert.equal(tabIn(w, a.path), null);
    } finally {
      globalThis.__okRelease?.();
    }
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('AI 작업 중 탭 창 닫기·끝내기 → 한 번 묻고, 취소면 창 유지·끝내기 취소 (D17)', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, 'AI 나.hwpx');
    const b = await openEditorIn(w, 'AI 다.hwpx');
    fake.replies.set(a.wc.id, true);
    const held = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    const chatId = held.value.chatId;
    const pidDir = join(ud, 'com-pids', chatId);
    const quits = fake.app.quitCalls;
    try {
      assert.ok(sameFile(globalThis.__okChats.at(-1).filePath, b.path), '고칠 문서 = 다');
      // 창 닫기 → 한 번 묻고 취소: 저장 확인도 하지 않고 모두 그대로
      let n = asked();
      fake.dialogAnswers.push(1);
      w.win.close();
      assert.ok(await until(() => asked() >= n + 1));
      await sleep(150);
      assert.equal(asked(), n + 1, '한 번만 묻는다');
      assert.equal(fake.dialogs[n].message, AI_WIN_MSG);
      assert.deepEqual(fake.dialogs[n].buttons, STOP_BUTTONS);
      assert.equal(fake.dialogs[n].cancelId, 1);
      assert.equal(fake.dialogs[n].parent, w.win);
      assert.equal(w.win.isDestroyed(), false);
      assert.equal(a.wc.isDestroyed(), false);
      assert.equal(b.wc.isDestroyed(), false);
      assert.equal(existsSync(pidDir), true, '대화 그대로');
      // 끝내기 → 한 번 묻고(부모 = 최근 탭 창) 취소: 아무 창도 닫지 않고 끝내기 취소
      w.win.emit('focus');
      n = asked();
      fake.dialogAnswers.push(1);
      const ev = quitEvent();
      fake.app.emit('before-quit', ev);
      assert.equal(ev.prevented, true, '차례 확인이 끝날 때까지 끝내기를 막는다');
      assert.ok(await until(() => asked() >= n + 1));
      await sleep(150);
      assert.equal(asked(), n + 1);
      assert.equal(fake.dialogs[n].message, AI_WIN_MSG);
      assert.equal(fake.dialogs[n].parent, w.win);
      assert.equal(fake.app.quitCalls, quits);
      for (const x of [w.win, tw0]) assert.equal(x.isDestroyed(), false);
      for (const x of [a.wc, b.wc, homeWc]) assert.equal(x.isDestroyed(), false);
      assert.equal(existsSync(pidDir), true);
      assert.deepEqual(hookCalls.splice(0).map((h) => h.name), ['finalize', 'resume'], '끝내기 전에 확정, 취소하면 다시 이어 적기');
      // 창 닫기 → 멈추고 닫기: 대화를 멈춘 뒤 저장 안 한 탭을 묻고, 창이 닫힌다 (다른 창이 있으니 끝내지 않음)
      n = asked();
      fake.dialogAnswers.push(0, 1); // 멈추고 닫기, 나: 저장 안 함
      w.win.close();
      assert.ok(await until(() => w.win.isDestroyed()), '창이 닫혀야 한다');
      assert.equal(asked(), n + 2);
      assert.equal(fake.dialogs[n].message, AI_WIN_MSG);
      assert.match(fake.dialogs[n + 1].message, /AI 나\.hwpx/);
      assert.ok(await endOf(chatId), '대화가 멈췄다');
      assert.ok(await until(() => !existsSync(pidDir)), 'PID 폴더 정리');
      assert.equal(fake.app.quitCalls, quits);
      assert.deepEqual(hookCalls.splice(0), [{ name: 'removeWindow', id: w.win.id }]);
    } finally {
      globalThis.__okRelease?.();
    }
    tw0.emit('focus');
  });

  await t.test('읽기 재료 탭은 AI 작업 중이어도 묻지 않고 닫힌다', async () => {
    const w = await newTabWindow(homeWc);
    const m = join(docs, '읽기 재료.pdf');
    writeFileSync(m, makePdf(1));
    assert.equal((await fake.invokeFrom(w.home, 'doc:open', m)).ok, true);
    const a = await openEditorIn(w, 'AI 라.hwpx');
    const held = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '[hold] 참고해서 고쳐줘', editorOk: true, target: { edit: a.path, read: [m] } });
    assert.equal(held.ok, true, held.error);
    try {
      assert.deepEqual(globalThis.__okChats.at(-1).materials.map((x) => sameFile(x, m)), [true]);
      const n = asked();
      const r = await fake.invokeFrom(w.tabstrip, 'tabs:close', tabIn(w, m).id);
      assert.equal(r.value, true);
      assert.equal(asked(), n, '재료 탭은 묻지 않는다');
      assert.equal(tabIn(w, m), null);
      assert.equal(lastSent('ai:event', (x) => x.type === 'end' && x.chatId === held.value.chatId), undefined, '대화는 계속');
    } finally {
      globalThis.__okRelease?.();
    }
    assert.ok(await endOf(held.value.chatId));
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('편집 탭 안 Ctrl+W·편집 화면의 닫기도 AI가 고치는 중이면 묻는다 (D17, RT1)', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, 'AI 마.hwpx');
    const held = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '[hold] 고쳐줘', editorOk: true });
    assert.equal(held.ok, true, held.error);
    try {
      let n = asked();
      fake.dialogAnswers.push(1); // 취소
      assert.equal(pressCtrl(a.wc, 'KeyW'), true, 'Ctrl+W는 가로챈다');
      assert.ok(await until(() => asked() >= n + 1), 'AI 작업 중 확인');
      assert.equal(fake.dialogs[n].message, AI_TAB_MSG);
      assert.equal(fake.dialogs[n].parent, w.win);
      await sleep(150);
      assert.equal(a.wc.isDestroyed(), false, '취소 → 탭 그대로');
      assert.ok(tabIn(w, a.path));
      // 편집 화면의 닫기(editor:close)도 같은 확인
      n = asked();
      fake.dialogAnswers.push(1);
      fake.emit('editor:close', a.wc);
      assert.ok(await until(() => asked() >= n + 1));
      assert.equal(fake.dialogs[n].message, AI_TAB_MSG);
      await sleep(150);
      assert.equal(a.wc.isDestroyed(), false);
      assert.ok(tabIn(w, a.path));
    } finally {
      globalThis.__okRelease?.();
    }
    assert.ok(await endOf(held.value.chatId));
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('재료만 읽는 AI 대화도 창 닫기·끝내기에서 묻고, 멈추고 닫기면 대화를 멈춘다 (D17, RT4)', async () => {
    const w = await newTabWindow(homeWc);
    const m = join(docs, '재료만 읽기.pdf');
    writeFileSync(m, makePdf(1));
    assert.equal((await fake.invokeFrom(w.home, 'doc:open', m)).ok, true);
    const held = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '[hold] 요약해줘' });
    assert.equal(held.ok, true, held.error);
    const chatId = held.value.chatId;
    const pidDir = join(ud, 'com-pids', chatId);
    try {
      assert.equal(globalThis.__okTargets.at(-1).edit, null, '재료만 읽는 대화');
      // 끝내기 → 한 번 묻고 취소: 끝내지 않고 대화 그대로
      w.win.emit('focus');
      const quits = fake.app.quitCalls;
      let n = asked();
      fake.dialogAnswers.push(1);
      const ev = quitEvent();
      fake.app.emit('before-quit', ev);
      assert.equal(ev.prevented, true);
      assert.ok(await until(() => asked() >= n + 1), '끝내기 확인');
      await sleep(150);
      assert.equal(asked(), n + 1);
      assert.equal(fake.dialogs[n].message, AI_WIN_MSG);
      assert.equal(fake.app.quitCalls, quits);
      assert.equal(existsSync(pidDir), true, '취소 → 대화 그대로');
      hookCalls.splice(0);
      // 창 닫기 → 한 번 묻고 멈추고 닫기: 대화가 멈추고 창이 닫힌다
      n = asked();
      fake.dialogAnswers.push(0);
      w.win.close();
      assert.ok(await until(() => w.win.isDestroyed()), '창이 닫힌다');
      assert.equal(asked(), n + 1);
      assert.equal(fake.dialogs[n].message, AI_WIN_MSG);
      assert.ok(await until(() => !existsSync(pidDir)), '대화가 멈추고 PID 폴더 정리');
    } finally {
      globalThis.__okRelease?.();
    }
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('재료만 보내기(edit:null)인데 지금 탭이 고칠 수 있는 문서면 TARGET_CHANGED (F5)', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, '대상 확인.hwpx');
    const n = (globalThis.__okChats ?? []).length;
    const r = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '요약해줘', editorOk: true, target: { edit: null, read: [] } });
    assert.equal(r.ok, false);
    assert.equal(r.code, 'TARGET_CHANGED', r.error);
    assert.equal((globalThis.__okChats ?? []).length, n, '대화를 시작하지 않는다');
    const ok = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: a.path, read: [] } });
    assert.equal(ok.ok, true, ok.error);
    assert.ok(await endOf(ok.value.chatId));
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('고칠 문서가 인터넷에서 받은 문서면 zone3: 폴더 읽기 없이 파일 단위 (D15, F6)', { skip: process.platform !== 'win32' && 'Zone.Identifier' }, async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, '인터넷 한글.hwpx');
    writeFileSync(`${a.path}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=3\r\n');
    try {
      let r = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '고쳐줘', editorOk: true, target: { edit: a.path, read: [] } });
      assert.equal(r.ok, true, r.error);
      assert.equal(globalThis.__okTargets.at(-1).zone3, true, '고칠 문서가 인터넷 문서');
      assert.ok(await endOf(r.value.chatId));
      // 같은 문서를 다시 보내면 이어 한다 (고칠 문서는 빠지지 않았다)
      r = await fake.invokeFrom(w.panel, 'ai:send', { prompt: '더 고쳐줘', editorOk: true, target: { edit: a.path, read: [] } });
      assert.equal(r.ok, true, r.error);
      assert.equal(r.value.newChat, false);
      assert.equal(globalThis.__okTargets.at(-1).zone3, true);
      assert.ok(await endOf(r.value.chatId));
    } finally {
      rmSync(`${a.path}:Zone.Identifier`, { force: true });
    }
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('세션 확정(finalize)이 던져도 끝내기 차례는 이어진다 (F2)', async () => {
    const w = await newTabWindow(homeWc);
    const a = await openEditorIn(w, '확정 실패.hwpx');
    fake.replies.set(a.wc.id, true);
    const keep = mainMod.sessionHooks.finalize;
    mainMod.sessionHooks.finalize = () => { throw new Error('설정 파일을 쓸 수 없음 (시험)'); };
    try {
      w.win.emit('focus');
      const quits = fake.app.quitCalls;
      const n = asked();
      fake.dialogAnswers.push(2); // 저장 확인에서 취소
      const ev = quitEvent();
      fake.app.emit('before-quit', ev);
      assert.equal(ev.prevented, true);
      assert.ok(await until(() => asked() >= n + 1), '확정이 던져도 저장 확인까지 간다');
      assert.match(fake.dialogs[n].message, /확정 실패\.hwpx/);
      await sleep(150);
      assert.equal(fake.app.quitCalls, quits);
      assert.equal(w.win.isDestroyed(), false);
    } finally {
      mainMod.sessionHooks.finalize = keep;
    }
    hookCalls.splice(0);
    fake.replies.set(a.wc.id, false);
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('홈·패널 화면이 죽으면 다시 싣는다, 정상 종료는 그대로 (F10)', async () => {
    const w = await newTabWindow(homeWc);
    for (const [wc, re] of [[w.home, /\?view=home$/], [w.panel, /\?view=panel$/]]) {
      const loads = [];
      const orig = wc.loadURL.bind(wc);
      wc.loadURL = (u) => { loads.push(u); return orig(u); };
      const err0 = console.error;
      console.error = () => {};
      try {
        fake.crash(wc, 'crashed');
        fake.crash(wc, 'clean-exit');
      } finally {
        console.error = err0;
      }
      assert.equal(loads.length, 1, '죽으면 한 번 다시 싣고, 정상 종료는 그대로');
      assert.match(loads[0], re);
    }
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('패널 너비: app:init이 탭 창에 쓴 너비와 범위를 주고, 설정 저장도 같은 범위로 자른다 (M2)', async () => {
    const w = await newTabWindow(homeWc);
    const init = await fake.invokeFrom(w.panel, 'app:init');
    assert.equal(init.ok, true, init.error);
    assert.deepEqual([init.value.panel.min, init.value.panel.max], [320, 640]);
    assert.ok(init.value.panel.width >= 320 && init.value.panel.width <= 640, String(init.value.panel.width));
    const keep = (await fake.invokeFrom(w.home, 'settings:get')).value.chatWidth;
    try {
      assert.equal((await fake.invokeFrom(w.home, 'settings:set', { chatWidth: 900 })).ok, true);
      assert.equal((await fake.invokeFrom(w.home, 'settings:get')).value.chatWidth, 640);
      assert.equal((await fake.invokeFrom(w.home, 'settings:set', { chatWidth: 200 })).ok, true);
      assert.equal((await fake.invokeFrom(w.home, 'settings:get')).value.chatWidth, 320);
    } finally {
      if (Number.isFinite(keep)) await fake.invokeFrom(w.home, 'settings:set', { chatWidth: keep });
    }
    w.win.close();
    assert.ok(await until(() => w.win.isDestroyed()));
    hookCalls.splice(0);
    tw0.emit('focus');
  });

  await t.test('끝내기: 창마다 차례로 묻고, 취소하면 남은 것은 그대로, 그 뒤 탭 하나 닫기는 끝내지 않음 (Review Focus)', async () => {
    const w2 = await newTabWindow(homeWc);
    const w1 = await newTabWindow(w2.home);
    const a = await openEditorIn(w1, '끝내기 가.hwpx');
    const b = await openEditorIn(w2, '끝내기 나.hwpx');
    fake.replies.set(a.wc.id, true);
    fake.replies.set(b.wc.id, true);
    // 최근 차례: 창1 → 창2 → 처음 창
    tw0.emit('focus');
    w2.win.emit('focus');
    w1.win.emit('focus');
    const quits = fake.app.quitCalls;
    const n = asked();
    fake.dialogAnswers.push(1, 2); // 가: 저장 안 함, 나: 취소
    const ev = quitEvent();
    fake.app.emit('before-quit', ev);
    assert.equal(ev.prevented, true);
    assert.ok(await until(() => asked() >= n + 2));
    await sleep(150);
    assert.equal(asked(), n + 2);
    assert.match(fake.dialogs[n].message, /끝내기 가\.hwpx/, '최근 창부터');
    assert.match(fake.dialogs[n + 1].message, /끝내기 나\.hwpx/);
    assert.equal(a.wc.isDestroyed(), true);
    assert.equal(w1.win.isDestroyed(), true, '다 닫힌 창1은 닫힌다');
    assert.equal(b.wc.isDestroyed(), false, '취소한 탭과');
    assert.equal(w2.win.isDestroyed(), false, '그 창과');
    assert.equal(tw0.isDestroyed(), false, '남은 창은 그대로');
    assert.equal(homeWc.isDestroyed(), false);
    assert.equal(fake.app.quitCalls, quits, '끝내지 않는다');
    // 확정 → (끝내기 중 닫힌 창은 세션에서 바로 빼지 않음) → 취소: 다시 이어 적고, 닫힌 창1을 뺀다
    assert.deepEqual(hookCalls.splice(0).map((h) => h.name === 'removeWindow' ? [h.name, h.id] : h.name), ['finalize', 'resume', ['removeWindow', w1.win.id]]);
    // 끝내기가 되돌아갔다: 깨끗한 탭 하나 닫기·창 하나 닫기는 끝내지 않는다
    const c = await openEditorIn(w2, '끝내기 다.hwpx');
    const r = await fake.invokeFrom(w2.tabstrip, 'tabs:close', tabIn(w2, c.path).id);
    assert.equal(r.value, true);
    assert.ok(await until(() => c.wc.isDestroyed()));
    assert.equal(fake.app.quitCalls, quits);
    fake.replies.set(b.wc.id, false);
    w2.win.close();
    assert.ok(await until(() => w2.win.isDestroyed()));
    assert.equal(fake.app.quitCalls, quits);
    assert.deepEqual(hookCalls.splice(0), [{ name: 'removeWindow', id: w2.win.id }], '끝내기가 아니라 창 하나 닫기');
    tw0.emit('focus');
  });

  await t.test('끝내기 중 사용자가 X로 닫은 창도 끝내기를 취소하면 세션에서 뺀다', async () => {
    const w2 = await newTabWindow(homeWc);
    const w1 = await newTabWindow(w2.home);
    const a = await openEditorIn(w1, '끝내기 중 가.hwpx');
    fake.replies.set(a.wc.id, true);
    tw0.emit('focus');
    w2.win.emit('focus');
    w1.win.emit('focus');
    const quits = fake.app.quitCalls;
    const n = asked();
    fake.dialogAnswers.push(2); // 가: 취소
    const ev = quitEvent();
    fake.app.emit('before-quit', ev);
    assert.equal(ev.prevented, true);
    w2.win.close(); // 차례 확인 중에 X
    assert.ok(await until(() => w2.win.isDestroyed()), 'X로 닫은 창은 닫힌다');
    assert.ok(await until(() => asked() >= n + 1));
    await sleep(150);
    assert.equal(fake.app.quitCalls, quits, '끝내지 않는다');
    assert.equal(w1.win.isDestroyed(), false);
    assert.deepEqual(hookCalls.splice(0).map((h) => h.name === 'removeWindow' ? [h.name, h.id] : h.name), ['finalize', 'resume', ['removeWindow', w2.win.id]]);
    fake.replies.set(a.wc.id, false);
    w1.win.close();
    assert.ok(await until(() => w1.win.isDestroyed()));
    assert.deepEqual(hookCalls.splice(0), [{ name: 'removeWindow', id: w1.win.id }]);
    tw0.emit('focus');
  });

  let closeAsked = 0;
  await t.test('탭 창 닫기: 고친 편집 탭이 있으면 묻고, 취소하면 모두 그대로', async () => {
    fake.replies.set(editorWin.wc.id, true);
    fake.dialogAnswers.push(2); // 취소
    closeAsked = fake.dialogs.length;
    tw0.close();
    await until(() => fake.dialogs.length >= closeAsked + 1);
    await sleep(100);
    assert.equal(fake.dialogs.length, closeAsked + 1);
    assert.deepEqual(fake.dialogs[closeAsked].buttons, ['저장', '저장 안 함', '취소']);
    assert.match(fake.dialogs[closeAsked].message, /보고서\.hwpx/);
    assert.equal(tw0.isDestroyed(), false);
    assert.equal(editorWin.wc.isDestroyed(), false);
    assert.equal(newEditor.wc.isDestroyed(), false);
  });

  await t.test('탭 창 닫기: 저장 안 함 → 편집 탭 모두 닫고 탭 창도 닫힌다 (파일은 그대로)', async () => {
    const hash = sha(docPath);
    const editorsOpen = editorViews();
    assert.ok(editorsOpen.length >= 6, '한글 2 + Office 5');
    fake.dialogAnswers.push(1); // 저장 안 함
    tw0.close();
    assert.ok(await until(() => tw0.isDestroyed(), 8000), '탭 창이 닫혀야 한다');
    assert.deepEqual(hookCalls.splice(0), [{ name: 'finalize', open: 1 }], '마지막 탭 창: 세션에서 빼기 전에(창이 살아 있을 때) 확정');
    for (const wc of [homeWc, panelWc, tabstripWc]) assert.equal(wc.isDestroyed(), true, '홈·패널·탭 줄 뷰도 닫는다');
    for (const v of editorsOpen) assert.equal(v.webContents.isDestroyed(), true, v.opts.webPreferences.partition);
    assert.equal(fake.dialogs.length, closeAsked + 2, '깨끗한 편집 창은 묻지 않는다');
    assert.equal(sha(docPath), hash);
    assert.deepEqual(readdirSync(join(ud, 'office-drafts')).filter((n) => /^[0-9a-f]{40}\./.test(n) && n !== basename(freshDraft)), [], '정상 닫기 → 복구 사본 없음');
  });

  await t.test('마지막 탭 창을 닫은 뒤 끝내기(window-all-closed)는 막지 않고, 세션도 다시 확정하지 않는다', () => {
    const quits = fake.app.quitCalls;
    const ev = quitEvent();
    fake.app.emit('before-quit', ev);
    assert.equal(ev.prevented, false);
    assert.equal(fake.app.quitCalls, quits);
    assert.deepEqual(hookCalls, []);
  });
});

// ── Office 편집기 설정 보관 (store.mjs) ──────────────────

test('store: officeGet/officeSet — 같은 값이면 다시 쓰지 않고, 한도·이상한 이름을 거른다', () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-office-state-'));
  try {
    const s = createSettings(dir);
    assert.equal(s.officeGet('a'), null);
    assert.equal(s.officeSet('a', { x: 1 }), true);
    assert.deepEqual(s.officeGet('a'), { x: 1 });
    const got = s.officeGet('a');
    got.x = 2;
    assert.deepEqual(s.officeGet('a'), { x: 1 }, '사본을 돌려준다');
    const m0 = statSync(s.file).mtimeMs;
    assert.equal(s.officeSet('a', { x: 1 }), false, '같은 값');
    assert.equal(statSync(s.file).mtimeMs, m0);
    // '__proto__'도 그냥 이름 (프로토타입을 바꾸지 않는다)
    assert.equal(s.officeSet('__proto__', { polluted: true }), true);
    assert.equal(({}).polluted, undefined);
    assert.deepEqual(s.officeGet('__proto__'), { polluted: true });
    assert.throws(() => s.officeSet('a/b', 1), /잘못된 설정 이름/);
    assert.throws(() => s.officeSet('big', 'x'.repeat(9000)), /8KB/);
    // 다시 읽어도 그대로, 한도를 넘으면 오래된 것부터
    for (let i = 0; i < OFFICE_STATE_MAX_KEYS + 5; i++) s.officeSet(`k${i}`, i);
    const again = createSettings(dir);
    assert.equal(again.officeGet('a'), null, '오래된 것은 밀려난다');
    assert.equal(again.officeGet(`k${OFFICE_STATE_MAX_KEYS + 4}`), OFFICE_STATE_MAX_KEYS + 4);
    assert.equal(Object.keys(again.get('officeEditor').settings).length, OFFICE_STATE_MAX_KEYS);
    assert.equal(again.officeSet(`k${OFFICE_STATE_MAX_KEYS + 4}`, null), true, '지우기');
    assert.equal(again.officeGet(`k${OFFICE_STATE_MAX_KEYS + 4}`), null);
    // 망가진 설정 파일도 받아 준다
    writeFileSync(s.file, JSON.stringify({ officeEditor: { settings: { ok: 1, 'bad key': 2, huge: 'y'.repeat(9000) } } }));
    assert.deepEqual(createSettings(dir).get('officeEditor').settings, { ok: 1 });
    writeFileSync(s.file, JSON.stringify({ officeEditor: 'x' }));
    assert.deepEqual(createSettings(dir).get('officeEditor'), { settings: {} });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

// ── '정보 · 라이선스' 글: docs/office-editor-legal-ko.md → src/app/office-editor/legal-ko.json ──

const LEGAL_JSON = join(ROOT, 'src', 'app', 'office-editor', 'legal-ko.json');

/** 글 한 줄 → 조각 [{ text, bold?, code?, href?, local? }] (**굵게**, `코드`, [글](https://… | local:경로)만) */
function legalInline(s) {
  const spans = [];
  let last = 0;
  for (const m of s.matchAll(/\*\*(.+?)\*\*|`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\)/g)) {
    if (m.index > last) spans.push({ text: s.slice(last, m.index) });
    if (m[1] !== undefined) spans.push({ text: m[1], bold: true });
    else if (m[2] !== undefined) spans.push({ text: m[2], code: true });
    else if (m[4].startsWith('local:')) spans.push({ text: m[3], local: m[4].slice('local:'.length) });
    else if (m[4].startsWith('https://')) spans.push({ text: m[3], href: m[4] });
    else throw new Error(`허용되지 않는 링크 대상: ${m[4]}`);
    last = m.index + m[0].length;
  }
  if (last < s.length) spans.push({ text: s.slice(last) });
  return spans;
}

/** 법적 고지 글(마크다운 일부 문법) → { title, intro, sections: [{ title, blocks: [{ p } | { ul }] }] } */
function legalFromMarkdown(md) {
  let text = md.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  if (text.trimStart().startsWith('<!--')) text = text.slice(text.indexOf('-->') + 3); // 맨 위 주석(읽는 쪽 안내)은 뺀다
  const out = { title: null, intro: [], sections: [] };
  let blocks = out.intro; let para = null; let list = null;
  const endPara = () => { if (para) { blocks.push({ p: legalInline(para.join(' ')) }); para = null; } };
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    if (!line) { endPara(); list = null; continue; }
    if (/^##\s/.test(line)) {
      endPara(); list = null;
      const s = { title: line.replace(/^##\s+/, ''), blocks: [] };
      out.sections.push(s);
      blocks = s.blocks;
    } else if (/^#\s/.test(line)) {
      endPara(); list = null;
      if (out.title) throw new Error('# 제목은 하나만 쓸 수 있습니다');
      out.title = line.replace(/^#\s+/, '');
    } else if (/^-\s/.test(line)) {
      endPara();
      if (!list) { list = { ul: [] }; blocks.push(list); }
      list.ul.push(legalInline(line.replace(/^-\s+/, '')));
    } else if (/^#{3,}\s|^>|^\d+\.\s|^\|/.test(line)) {
      throw new Error(`이 글에서 쓰지 않는 문법입니다: ${line}`);
    } else {
      list = null;
      (para ??= []).push(line);
    }
  }
  endPara();
  return out;
}

for (const lang of ['ko', 'en']) {
  test(`정보 · 라이선스 글: legal-${lang}.json이 docs/office-editor-legal-${lang}.md와 같다`, () => {
    const md = join(ROOT, 'docs', `office-editor-legal-${lang}.md`);
    const file = join(ROOT, 'src', 'app', 'office-editor', `legal-${lang}.json`);
    const want = legalFromMarkdown(readFileSync(md, 'utf8'));
    if (process.env.OFFICEKIT_UPDATE_LEGAL === '1') {
      const json = {
        $comment: `docs/office-editor-legal-${lang}.md에서 만든 Office 편집 창 '정보 · 라이선스' 글. 고치려면 그 문서를 고친 뒤 OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs. {{appVersion}}은 편집 창이 앱 버전으로 바꾼다.`,
        source: `docs/office-editor-legal-${lang}.md`,
        ...want,
      };
      writeFileSync(file, `${JSON.stringify(json, null, 2)}\n`);
    }
    const have = JSON.parse(readFileSync(file, 'utf8'));
    assert.deepEqual({ title: have.title, intro: have.intro, sections: have.sections }, want,
      `legal-${lang}.json이 문서와 다릅니다. OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs 로 다시 만드세요`);
    assert.equal(have.source, `docs/office-editor-legal-${lang}.md`);
  });
}

test('정보 · 라이선스 글: 영어 글(화면 언어가 영어·일본어일 때)도 같은 연결·고정 값을 쓴다', () => {
  const links = (f) => [...readFileSync(join(ROOT, 'docs', f), 'utf8').replace(/<!--[\s\S]*?-->/, '').matchAll(/\]\(([^)\s]+)\)/g)].map((m) => m[1]);
  assert.deepEqual(links('office-editor-legal-en.md'), links('office-editor-legal-ko.md'));
  const en = readFileSync(join(ROOT, 'docs', 'office-editor-legal-en.md'), 'utf8').replace(/<!--[\s\S]*?-->/, '');
  assert.doesNotMatch(en, /[가-힣]/, '영어 글에 한국어가 남지 않는다');
});

test('정보 · 라이선스 글: 설계 8장 고지 (원 개발사·수정판·라이선스·소스·상표)와 안전한 링크', () => {
  const legal = JSON.parse(readFileSync(LEGAL_JSON, 'utf8'));
  const secs = legalSections(legal, { appVersion: '0.1.2' });
  assert.ok(secs, '절로 바뀐다');
  const titles = secs.map((s) => s.title);
  for (const t of ['원 개발사', '수정된 판일 수 있습니다', '라이선스', '소스 코드', '상표']) assert.ok(titles.includes(t), t);
  const spans = secs.flatMap((s) => s.blocks.flatMap((b) => b.p ?? b.ul.flat()));
  const all = spans.map((s) => s.text).join('');
  assert.match(all, /Ascensio System SIA/);
  assert.match(all, /수정된 판일 수 있습니다/);
  assert.match(all, /AGPL-3\.0/);
  assert.match(all, /아무런 보증이 없습니다/);
  assert.doesNotMatch(JSON.stringify(secs), /\{\{/, '자리표시는 모두 바뀐다');
  // 소스 꾸러미 주소는 설치된 판의 태그를 가리킨다
  const hrefs = spans.filter((s) => s.href).map((s) => s.href);
  assert.ok(hrefs.includes('https://github.com/toxh13/office-kit/releases/download/v0.1.2/office-editor-source-0.1.2.zip'));
  assert.ok(hrefs.every((h) => h.startsWith('https://')));
  // 묶음 안 고지 파일 링크는 '라이선스 원문' 탭에 있는 파일만
  const locals = spans.filter((s) => s.local).map((s) => s.local);
  assert.ok(locals.length >= 4);
  for (const l of locals) assert.ok(NOTICE_FILES.some((n) => n.path === l), l);
  const office = join(ROOT, 'vendor', 'office-editor');
  if (existsSync(join(office, 'VERSION.json'))) for (const l of locals) assert.ok(existsSync(join(office, ...l.split('/'))), `묶음에 ${l}`);
  // 판을 모르면 릴리스 목록으로
  const unknown = legalSections(legal, {}).flatMap((s) => s.blocks.flatMap((b) => b.p ?? b.ul.flat()));
  assert.ok(unknown.some((s) => s.href === 'https://github.com/toxh13/office-kit/releases'));
  assert.doesNotMatch(JSON.stringify(unknown), /\{\{/);
  // 대화상자 내용: 글이 있으면 그 절을 쓰고, 설치된 판 정보는 그대로
  const m = aboutModel({ appVersion: '0.1.2', bundle: FAKE_VERSION }, 'http://127.0.0.3:5000/office/abc/index.html', legal);
  assert.equal(m.legal, true);
  assert.equal(m.sections[0].title, null, '첫 문단 (기반 소개)');
  assert.ok(m.facts.some((f) => f.label === '편집기'));
  assert.equal(aboutModel({ appVersion: '0.1.2' }, 'x', { sections: [] }).legal, false, '읽을 수 없는 글이면 들어 있는 글로');
  // 나쁜 링크는 글자로만
  const evil = legalSections({ sections: [{ title: 't', blocks: [{ p: [{ text: 'a', href: 'javascript:alert(1)' }, { text: 'b', local: '../../x' }, { text: 'c', html: '<b>' }] }] }] });
  assert.deepEqual(evil[0].blocks[0].p, [{ text: 'a' }, { text: 'b' }, { text: 'c' }]);
});

test('office 서버: 고지 파일(.md·LICENSE·NOTICE)은 글로 내준다 (정보 · 라이선스 보기 틀)', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-office-notice-'));
  const office = join(dir, 'office');
  put(join(office, 'NOTICE-OfficeKit.md'), '# 수정 고지\n');
  put(join(office, 'NOTICE'), 'filex NOTICE');
  put(join(office, 'LICENSE'), 'AGPL');
  put(join(office, 'VERSION.json'), '{}');
  put(join(office, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d]));
  const srv = await startEditorServer({ root: dir, dirs: { office, officeHost: join(ROOT, 'src', 'app', 'office-editor') } });
  try {
    assert.ok(srv.officeBase, srv.officeError);
    for (const f of ['NOTICE-OfficeKit.md', 'NOTICE', 'LICENSE']) {
      const r = await fetch(`${srv.officeBase}${f}`);
      assert.equal(r.status, 200, f);
      assert.equal(r.headers.get('content-type'), 'text/plain; charset=utf-8', f);
      await r.arrayBuffer();
    }
    const w = await fetch(`${srv.officeBase}x2t/x2t.wasm`);
    assert.equal(w.headers.get('content-type'), 'application/wasm');
    await w.arrayBuffer();
    // host 페이지 옆의 legal-ko.json (편집 창이 읽는다)
    const j = await fetch(`${srv.origin}/office-host/legal-ko.json`);
    assert.equal(j.status, 200);
    assert.match(j.headers.get('content-type'), /application\/json/);
    assert.equal((await j.json()).source, 'docs/office-editor-legal-ko.md');
  } finally {
    await srv.close();
    rmSync(dir, { recursive: true, force: true });
  }
});

// ── 실제 Electron 확인용 캡처 (선택) ───────────────────────

const wantSmoke = process.env.OFFICEKIT_EDITOR_SMOKE === '1';
const electronExe = !wantSmoke ? '' : (() => { try { return spawnSync(process.execPath, ['-e', "process.stdout.write(require('electron'))"], { cwd: ROOT, encoding: 'utf8' }).stdout; } catch { return ''; } })();
const hasElectron = !!electronExe && existsSync(electronExe);
const smokeReady = wantSmoke && hasElectron && existsSync(join(ROOT, 'vendor', 'rhwp-studio', 'index.html'));
const officeSmokeReady = wantSmoke && hasElectron && existsSync(join(ROOT, 'vendor', 'office-editor', 'VERSION.json'));

/** 실제 앱을 OFFICEKIT_SMOKE(+_EDITOR)로 띄워 doc을 열고 두 창을 찍는다 → [smoke] 줄의 정보 */
function runSmoke(base, doc, tag) {
  const png = join(base, `${tag}.png`);
  const ud = join(base, `ud-${tag}`);
  // 주 창 미리보기는 글자만 (Word·Excel·PowerPoint를 띄우지 않게)
  // 창 제목·문구는 한국어로 단언하므로 Windows 언어와 무관하게 한국어 화면으로 (CI는 영어 Windows)
  put(join(ud, 'settings.json'), JSON.stringify({ previewMode: 'text', lang: 'ko' }));
  const env = { ...process.env, OFFICEKIT_SMOKE: png, OFFICEKIT_SMOKE_EDITOR: '1', OFFICEKIT_USER_DATA: ud };
  delete env.ELECTRON_RUN_AS_NODE;
  for (const k of ['OFFICEKIT_EDITOR_VENDOR', 'OFFICEKIT_OFFICE_FONTS_MODULE', 'OFFICEKIT_AI_MODULE']) delete env[k];
  const r = spawnSync(electronExe, ['.', doc], { cwd: ROOT, env, encoding: 'utf8', timeout: 200_000 });
  const line = (r.stdout ?? '').split(/\r?\n/).find((l) => l.startsWith('[smoke] '));
  assert.equal(r.status, 0, `${r.stdout}\n${r.stderr}`);
  const info = JSON.parse(line.slice('[smoke] '.length));
  const out = process.env.OFFICEKIT_SMOKE_OUT;
  if (out) {
    mkdirSync(out, { recursive: true });
    copyFileSync(png, join(out, `${tag}.png`));
    const e = info.editor ?? {};
    for (const [f, name] of [[e.png, 'editor'], [e.aboutInfo?.png, 'editor-about'], [e.aboutInfo?.textsPng, 'editor-about-texts']]) {
      if (f && existsSync(f)) copyFileSync(f, join(out, `${tag}-${name}.png`));
    }
    writeFileSync(join(out, `${tag}.json`), JSON.stringify(info, null, 2));
  }
  return { info, png };
}

test('실제 앱: OFFICEKIT_SMOKE_EDITOR로 주 창 + 한글 편집 창 캡처', { skip: !smokeReady && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/ 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async () => {
  const base = mkdtempSync(join(tmpdir(), 'ok-editor-smoke-'));
  try {
    const doc = join(base, '캡처 시험.hwpx');
    await createBlankDocument(doc, { title: '캡처 시험 문서' });
    const { info, png } = runSmoke(base, doc, 'hwpx');
    assert.equal(info.ok, true);
    assert.equal(info.editable, true, '직접 편집 단추');
    assert.equal(info.editor.ok, true);
    assert.equal(info.editor.state, 'saved');
    assert.equal(info.editor.title, '캡처 시험.hwpx - 한글 편집 (Office Kit)');
    assert.ok(statSync(png).size > 10_000);
    assert.ok(statSync(info.editor.png).size > 10_000);
  } finally {
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
});

// 실제 한국어 문서 (원고 → createDocument)
const OFFICE_SMOKE_DOCS = {
  docx: {
    name: '2026년 4분기 사업 보고서.docx', title: '문서 편집',
    md: [
      '# 2026년 4분기 사업 보고서', '>> 기획팀 김민지 · 2026년 10월 18일', '',
      '## 1. 요약', '4분기 매출은 전년 같은 분기보다 12% 늘었습니다. 새로 연 부산 지점이 늘어난 몫의 절반을 차지했습니다.', '',
      '- 매출: 13억 4,400만 원 (전년 대비 +12%)', '- 새 고객: 182곳 (목표 150곳)', '- 고객 만족도: 4.6점 / 5점', '',
      '## 2. 지점별 실적', '| 지점 | 3분기 | 4분기 | 늘어난 비율 |', '|---|---|---|---|', '| 서울 본점 | 6억 2,000만 | 6억 5,100만 | 5% |',
      '| 부산 지점 | 2억 1,000만 | 3억 4,300만 | 63% |', '| 대전 지점 | 3억 | 3억 5,000만 | 17% |', '',
      '## 3. 다음 분기 계획', '광주 지점 문을 열고, 고객 상담 시간을 저녁 8시까지 늘립니다.',
    ].join('\n'),
  },
  xlsx: {
    name: '부서별 예산표.xlsx', title: '표 편집',
    md: [
      '# 2027년 부서별 예산 (단위: 만 원)', '',
      '| 부서 | 인건비 | 운영비 | 장비 | 합계 |', '|---|---|---|---|---|',
      '| 기획팀 | 42000 | 8000 | 3000 | 53000 |', '| 개발팀 | 96000 | 12000 | 21000 | 129000 |',
      '| 영업팀 | 51000 | 15000 | 2000 | 68000 |', '| 고객지원팀 | 38000 | 6000 | 1500 | 45500 |',
    ].join('\n'),
  },
  pptx: {
    name: '신제품 소개 발표.pptx', title: '슬라이드 편집',
    md: [
      '# 한글 문서 도우미 소개', '>> 2026년 가을 신제품 설명회', '- 인터넷 없이 이 PC 안에서 문서를 고칩니다', '- 한글 · 워드 · 엑셀 · 파워포인트를 한 곳에서', '',
      '# 이런 분께 좋습니다', '- 보고서를 자주 고치는 직장인', '- 표 계산이 많은 회계 담당자', '- 발표 자료를 빨리 만들어야 하는 팀장님',
    ].join('\n'),
  },
};

test('실제 앱: OFFICEKIT_SMOKE_EDITOR로 주 창 + 문서·표·슬라이드 편집 창 캡처 (한국어 문서, 이 PC 글꼴)', { skip: !officeSmokeReady && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/office-editor 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-office-smoke-'));
  t.after(() => rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }));
  for (const [ext, d] of Object.entries(OFFICE_SMOKE_DOCS)) {
    await t.test(ext, async () => {
      const doc = join(base, d.name);
      await createDocument(doc, d.md);
      const before = sha(doc);
      const { info, png } = runSmoke(base, doc, ext);
      assert.equal(info.ok, true, JSON.stringify(info));
      assert.equal(info.editable, true, '직접 편집 단추를 누를 수 있다');
      const e = info.editor;
      assert.equal(e.ok, true, JSON.stringify(e));
      assert.equal(e.kind, 'office');
      assert.equal(e.phase, 'ready', '편집기가 문서를 다 열었다');
      assert.equal(e.title, `${basename(doc)} - ${d.title} (Office Kit)`);
      assert.doesNotMatch(e.title, /ONLYOFFICE/i);
      assert.equal(e.about, '정보 · 라이선스', '정보 · 라이선스 단추');
      // 정보 · 라이선스 대화상자: legal-ko.json 글 + 수정 고지 원문을 글로 보여 준다
      assert.equal(e.aboutInfo?.legal, true, JSON.stringify(e.aboutInfo));
      for (const t of ['원 개발사', '라이선스', '소스 코드', '상표', '설치된 판']) assert.ok(e.aboutInfo.titles.includes(t), t);
      assert.match(e.aboutInfo.text.url, /\/NOTICE-OfficeKit\.md$/);
      assert.equal(e.aboutInfo.text.note, null, '보기 틀에 글로 보인다');
      assert.ok(e.fonts?.count > 50, `이 PC 글꼴 목록 (${JSON.stringify(e.fonts)})`);
      assert.equal(e.fonts.malgun, true, '맑은 고딕이 글꼴 목록에');
      assert.ok(statSync(png).size > 10_000);
      assert.ok(statSync(e.png).size > 30_000);
      assert.equal(sha(doc), before, '열기만 해서는 파일이 바뀌지 않는다');
    });
  }
});

// 탭 창(Plan 2b): 스모크는 편집 창(BrowserWindow)이 아니라 탭 창 · 뷰를 찾는다. 탭 창 실제 스모크는 CI(main push·출시)에서 묶음을 만든 뒤 돈다
test('스모크 코드: 편집 창을 BrowserWindow 목록으로 찾지 않고, 탭 창 스모크 갈래와 CI 단계가 있다 (D16)', () => {
  const main = readFileSync(join(ROOT, 'src', 'app', 'main.mjs'), 'utf8');
  const start = main.indexOf('// ── 확인용 캡처 (OFFICEKIT_SMOKE)');
  const end = main.indexOf('// ── 웹 링크', start);
  assert.ok(start > 0 && end > start, '스모크 절');
  const smoke = main.slice(start, end);
  assert.doesNotMatch(smoke, /BrowserWindow\.getAllWindows/, '편집 창은 탭 창의 뷰다');
  assert.doesNotMatch(smoke, /\bew\.(moveTop|getTitle|webContents)/, '편집 창 대신 탭 창·편집 뷰');
  assert.match(main, /OFFICEKIT_SMOKE_TABS === '1'/);
  assert.match(main, /OFFICEKIT_SMOKE_TABS_FILES/);
  assert.match(smoke, /SMOKE_TABS \? 360_000 : 240_000/, '탭 스모크 시간 초과 360초');
  assert.ok(smoke.indexOf('view?.getBounds?.()') < smoke.indexOf('tw.shell.setPanel()'), '자리는 다시 배치하기 전에 읽는다');
  const yml = readFileSync(join(ROOT, '.github', 'workflows', 'release.yml'), 'utf8');
  const steps = yml.split(/\r?\n(?=      - )/).filter((s) => /name: 탭 창 실제 스모크/.test(s));
  assert.equal(steps.length, 2, 'test 잡(main push)과 release 잡에 하나씩');
  for (const s of steps) {
    assert.match(s, /shell: bash/);
    assert.match(s, /OFFICEKIT_EDITOR_SMOKE: '1'/);
    assert.match(s, /run: node --test test\/editor-integration\.test\.mjs/);
  }
  // test 잡의 단계는 묶음을 만드는 main push에서만 (같은 if), 'Office 편집기 묶음 빌드 (main)' 뒤에
  const testJob = yml.slice(yml.indexOf('\n  test:'), yml.indexOf('\n  release:'));
  assert.ok(testJob.indexOf('name: 탭 창 실제 스모크') > testJob.indexOf('name: Office 편집기 묶음 빌드 (main)'));
  assert.match(steps[0], /if: github\.event_name == 'push' && github\.ref == 'refs\/heads\/main'/);
  const relJob = yml.slice(yml.indexOf('\n  release:'));
  assert.ok(relJob.indexOf('name: 탭 창 실제 스모크') > relJob.indexOf('name: 한글→PDF 변환 프로그램 (rhwp v0.8.7 → vendor/rhwp-cli)'));
});

/** 실제 앱을 OFFICEKIT_SMOKE_TABS=1로 띄워 탭 창 하나에 files를 탭으로 연다 → [smoke] 줄의 정보 */
function runTabsSmoke(base, files) {
  const png = join(base, 'tabs.png');
  const ud = join(base, 'ud-tabs');
  // 창 제목·문구는 한국어로 단언하므로 Windows 언어와 무관하게 한국어 화면으로 (CI는 영어 Windows)
  put(join(ud, 'settings.json'), JSON.stringify({ previewMode: 'text', lang: 'ko' }));
  const env = { ...process.env, OFFICEKIT_SMOKE: png, OFFICEKIT_SMOKE_TABS: '1', OFFICEKIT_SMOKE_TABS_FILES: files.join(';'), OFFICEKIT_USER_DATA: ud };
  delete env.ELECTRON_RUN_AS_NODE;
  for (const k of ['OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_VENDOR', 'OFFICEKIT_OFFICE_FONTS_MODULE', 'OFFICEKIT_AI_MODULE']) delete env[k];
  // 앱의 탭 스모크 시간 초과(360초)보다 길게 — 넘으면 '[smoke] 시간 초과'가 보이게
  const r = spawnSync(electronExe, ['.'], { cwd: ROOT, env, encoding: 'utf8', timeout: 380_000 });
  const line = (r.stdout ?? '').split(/\r?\n/).find((l) => l.startsWith('[smoke] '));
  assert.ok(line, `${r.stdout}\n${r.stderr}`);
  const info = JSON.parse(line.slice('[smoke] '.length));
  const out = process.env.OFFICEKIT_SMOKE_OUT;
  if (out) {
    mkdirSync(out, { recursive: true });
    for (const f of [info.tabs?.png, info.tabs?.stripPng]) if (f && existsSync(f)) copyFileSync(f, join(out, basename(f)));
    writeFileSync(join(out, 'tabs.json'), JSON.stringify(info, null, 2));
  }
  assert.equal(r.status, 0, `${r.stdout}\n${r.stderr}`);
  return info;
}

test('실제 앱: 탭 창 스모크 — HWP·DOCX·PDF 탭 열기 → 전환 → 탭 닫기 → 창 닫기 (D16)', { skip: !(smokeReady && officeSmokeReady) && 'OFFICEKIT_EDITOR_SMOKE=1 과 vendor/ 묶음이 있을 때만 (창이 잠깐 뜬다)' }, async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-tabs-smoke-'));
  t.after(() => rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }));
  const hwpx = join(base, '탭 시험.hwpx');
  await createBlankDocument(hwpx, { title: '탭 시험 문서' });
  const docx = join(base, OFFICE_SMOKE_DOCS.docx.name);
  await createDocument(docx, OFFICE_SMOKE_DOCS.docx.md);
  const pdf = join(base, '탭 시험.pdf');
  writeFileSync(pdf, makePdf(2));
  const info = runTabsSmoke(base, [hwpx, docx, pdf]);
  const tabs = info.tabs;
  assert.ok(tabs, JSON.stringify(info));
  assert.equal(tabs.ok, true, JSON.stringify(tabs));
  assert.equal(tabs.opened, 4, '홈 + 세 문서 탭');
  assert.equal(tabs.switches.length, 4);
  for (const s of tabs.switches) {
    assert.deepEqual({ visible: s.visible, focused: s.focused, contentBounds: s.contentBounds }, { visible: true, focused: true, contentBounds: true }, JSON.stringify(s));
  }
  assert.equal(tabs.hwp.state, 'saved', JSON.stringify(tabs.hwp));
  assert.equal(tabs.hwp.title, '탭 시험.hwpx - 한글 편집 (Office Kit)');
  assert.equal(tabs.docx.phase, 'ready', JSON.stringify(tabs.docx));
  assert.equal(tabs.pdf.viewer, true, '내장 PDF 뷰어 틀');
  assert.equal(tabs.afterCloseTab, 3);
  assert.equal(tabs.closedDestroyed, true, "닫은 탭 뷰 wc에 'destroyed'가 온다");
  assert.equal(tabs.windowClosed, true);
  assert.ok(statSync(tabs.png).size > 10_000);
});
