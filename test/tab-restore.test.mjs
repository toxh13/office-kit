// 탭 세션 복원 시험 (Task 11): 평소 시작 = 지난 세션 자동 복원(활성 탭만 불러오기), 지연 탭, 세션 저장·끝내기.
// main.mjs는 한 프로세스에 한 번만 불러올 수 있으므로 한 시험 안에서 차례로 이어 간다
// (평소 시작 → 지연 탭 → 세션 저장 → 마지막 창 닫기·다시 열기 → 끝내기 확정 → 설정 파일 확인).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBlankDocument, createDocument } from '../src/core/create.mjs';
import { createSettings } from '../src/app/store.mjs';
import { draftPathFor } from '../src/app/editor-window.mjs';
import { promoteCrashed, restorePlan } from '../src/app/tab-session.mjs';
import { makePdf } from './helpers/fixtures.mjs';
import { makeFakeElectron, installFakeElectron } from './helpers/fake-electron.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(here, '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
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

/** 시험용 편집기 묶음 (한글 vendor/rhwp-*, 문서·표·슬라이드 vendor/office-editor 흉내: 있는지 확인하는 파일만) */
function makeVendors(dir) {
  put(join(dir, 'rhwp-studio', 'index.html'), '<!doctype html><title>studio</title>');
  put(join(dir, 'rhwp-studio', 'assets', 'rhwp_bg-test.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(dir, 'rhwp-editor-sdk', f), 'export {};');
  const o = join(dir, 'office-editor');
  put(join(o, 'index.html'), '<!doctype html><title>app</title>');
  put(join(o, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(o, 'x2t', 'x2t.js'), '// x2t');
  put(join(o, 'editor', 'LICENSE.txt'), 'GNU AFFERO GENERAL PUBLIC LICENSE');
  put(join(o, 'VERSION.json'), JSON.stringify({
    name: 'office-editor', officeKitVersion: '0.0.0-test', editorVersion: '9.4.0.129',
    onlyoffice: { version: '9.4.0', build: '9.4.0.129', sourceTag: 'v9.4.0.129' },
    filex: { tag: 'v0.3.1', commit: '57dda3dae99af78eb26b3c64b69f1e3f711f0757' }, files: [],
  }));
}

const FAKE_AI = `
export async function detectProviders() { return [{ id: 'claude', name: 'Claude (시험)', available: true, version: '1', hint: '' }]; }
export function defaultWorkspace() { return process.env.OFFICEKIT_TEST_WS; }
export function aiMcpEnv() { return {}; }
export function startChat({ onEvent }) {
  const done = Promise.resolve().then(() => { onEvent({ type: 'done', sessionId: 's1', text: '네' }); return { sessionId: 's1', text: '네' }; });
  return { cancel() {}, done };
}
`;
const FAKE_FONTS = `
export async function ensureFonts() { return { allFontsJsPath: null, thumbnailsDir: null, fontFileByIndex: () => null }; }
`;

test('평소 시작: 지난 세션 자동 복원, 활성 탭만 편집 뷰, 없는 파일은 알림 (D2·D4)', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-tab-restore-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  makeVendors(vendor);
  const a = join(docs, '가.hwpx');
  await createBlankDocument(a, { title: '가' });
  const b = join(docs, '나.docx');
  await createDocument(b, '# 나\n\n내용');
  const c = join(docs, '다.pdf');
  writeFileSync(c, makePdf(1));
  const d = join(docs, '라.pdf');
  writeFileSync(d, makePdf(1));
  const e = join(docs, '마.docx');
  await createDocument(e, '# 마\n\n재료');
  const gone = join(docs, '없음.pdf');
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);
  const fontsModule = join(base, 'fake-fonts.mjs');
  writeFileSync(fontsModule, FAKE_FONTS);

  // lastTabSession: 창 2개, 창1 [가.hwpx, 나.docx(활성), 없음.pdf, 라.pdf], 창2 [다.pdf, 마.docx] (홈 활성)
  const last = {
    v: 1, at: 1000, explorer: false, finalized: true,
    windows: [
      { bounds: { x: 100, y: 80, width: 1300, height: 850 }, maximized: false, active: 2,
        tabs: [{ type: 'editor', path: a }, { type: 'editor', path: b }, { type: 'doc', path: gone }, { type: 'doc', path: d }] },
      { bounds: { x: 200, y: 120, width: 1100, height: 760 }, maximized: true, active: 0,
        tabs: [{ type: 'doc', path: c }, { type: 'editor', path: e }] },
    ],
  };
  put(join(ud, 'settings.json'), JSON.stringify({ lastTabSession: last }));
  const readSettings = () => JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));

  const saved = { argv: process.argv, env: { ...process.env } };
  Object.assign(process.env, {
    OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_VENDOR: vendor, OFFICEKIT_AI_MODULE: aiModule, OFFICEKIT_TEST_WS: join(base, 'ws'),
    OFFICEKIT_OFFICE_FONTS_MODULE: fontsModule,
  });
  for (const k of ['OFFICEKIT_SMOKE', 'OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_TEST', 'OFFICEKIT_DOC_WINDOW']) delete process.env[k];
  process.argv = [process.argv[0]]; // 명령줄 파일 없음 = 평소 시작
  const fake = makeFakeElectron({ userData: ud, documents: docs });
  installFakeElectron(fake);
  t.after(async () => {
    fake.app.emit('will-quit');
    for (const w of fake.BaseWindow.getAllWindows()) w.destroy();
    await sleep(100);
    process.argv = saved.argv;
    for (const k of Object.keys(process.env)) if (!(k in saved.env)) delete process.env[k];
    Object.assign(process.env, saved.env);
    rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  });

  await import(pathToFileURL(join(ROOT, 'src', 'app', 'main.mjs')).href);
  assert.ok(await until(() => fake.windows.length >= 2, 20_000), '탭 창 2개');
  assert.equal(fake.app.exitCode, undefined, fake.errorBox);
  const [w1, w2] = fake.windows;
  /** 그 창에 붙은 뷰 중 주소가 re에 맞는 wc */
  const wcIn = (win, re) => win.contentView.children.map((v) => v.webContents).filter((wc) => wc && !wc.isDestroyed() && re.test(wc.mainFrame.url));
  const tabstripOf = (win) => fake.views.map((v) => v.webContents).find((wc) => /tabstrip\.html$/.test(wc.mainFrame.url) && win.contentView.children.some((c) => c.webContents === wc));
  const strip1 = tabstripOf(w1);
  const strip2 = tabstripOf(w2);
  const home1 = wcIn(w1, /\?view=home$/)[0];
  const panel1 = wcIn(w1, /\?view=panel$/)[0];
  const tabsOf = (strip) => strip.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [], active: null };
  const tabOf = (strip, p) => tabsOf(strip).tabs.find((x) => x.path && basename(x.path) === basename(p)) ?? null;
  const editorViews = () => fake.views.filter((v) => v.opts?.webPreferences?.partition && !v.webContents.isDestroyed());
  const toasts = (wc) => wc.sent.filter((s) => s.channel === 'toast').map((s) => s.data);

  // 창 자리·최대화는 세션 값 그대로
  assert.equal(fake.windows.length, 2);
  assert.deepEqual([w1.opts.x, w1.opts.y, w1.opts.width, w1.opts.height], [100, 80, 1300, 850]);
  assert.equal(w1.isMaximized(), false);
  assert.equal(w2.isMaximized(), true);
  // 활성 탭(나.docx)만 편집 뷰, 나머지는 불러오지 않은 탭
  assert.ok(await until(() => editorViews().length >= 1 && tabOf(strip1, b)?.loaded), '활성 탭 나.docx 편집 뷰');
  await sleep(200);
  assert.equal(editorViews().length, 1, '편집 뷰는 활성 탭 하나만');
  const l1 = tabsOf(strip1);
  assert.deepEqual(l1.tabs.map((x) => [x.type, x.path && basename(x.path), x.loaded]),
    [['home', null, true], ['editor', '가.hwpx', false], ['editor', '나.docx', true], ['doc', '라.pdf', false]], '없는 파일은 빼고 차례 그대로');
  assert.equal(l1.active, tabOf(strip1, b).id);
  const l2 = tabsOf(strip2);
  assert.deepEqual(l2.tabs.map((x) => [x.type, x.path && basename(x.path), x.loaded]),
    [['home', null, true], ['doc', '다.pdf', false], ['editor', '마.docx', false]]);
  assert.equal(l2.active, 'home');
  assert.equal(wcIn(w2, /\?view=doc$/).length, 0, '창2는 문서 탭 뷰도 만들지 않는다');
  // 없는 파일 알림: 첫 창 홈이 준비되면 (app:init) 받는다
  const init = await fake.invokeFrom(home1, 'app:init');
  assert.equal(init.ok, true, init.error);
  assert.ok(await until(() => toasts(home1).some((x) => x.type === 'warn' && x.message === '찾을 수 없어 열지 못한 문서: 없음.pdf')), JSON.stringify(toasts(home1)));

  await t.test('지연 탭 활성 → 그때 편집 뷰', async () => {
    const r = await fake.invokeFrom(strip1, 'tabs:activate', tabOf(strip1, a).id);
    assert.equal(r.ok, true, r.error);
    assert.ok(await until(() => editorViews().length === 2 && tabOf(strip1, a)?.loaded), '가.hwpx 편집 뷰');
    assert.equal(tabsOf(strip1).active, tabOf(strip1, a).id);
    assert.deepEqual(tabsOf(strip1).tabs.map((x) => x.type), ['home', 'editor', 'editor', 'doc'], '탭 수·자리 그대로');
    // 문서 탭: 문서 탭 뷰를 그 자리에 만든다
    const strip2r = await fake.invokeFrom(strip2, 'tabs:activate', tabOf(strip2, c).id);
    assert.equal(strip2r.ok, true, strip2r.error);
    assert.ok(await until(() => tabOf(strip2, c)?.loaded && wcIn(w2, /\?view=doc$/).length === 1), '다.pdf 문서 탭 뷰');
    const docWc = wcIn(w2, /\?view=doc$/)[0];
    assert.ok(await until(() => docWc.sent.some((s) => s.channel === 'doc:opened' && basename(s.data?.path ?? '') === '다.pdf')));
    assert.equal(tabOf(strip2, e).loaded, false, '마.docx는 아직');
  });

  await t.test('지연 탭 파일이 없어지면 탭을 빼고 알림 (Review Focus)', async () => {
    rmSync(d);
    const r = await fake.invokeFrom(strip1, 'tabs:activate', tabOf(strip1, d).id);
    assert.equal(r.ok, true, r.error);
    assert.ok(await until(() => !tabOf(strip1, d)), '탭을 뺀다');
    assert.ok(await until(() => toasts(home1).some((x) => x.message === '‘라.pdf’ 파일을 찾을 수 없어 탭을 닫았어요')), JSON.stringify(toasts(home1)));
    assert.equal(wcIn(w1, /\?view=doc$/).length, 0, '문서 탭 뷰를 만들지 않는다');
  });

  await t.test('지연 탭 재료: 자동 저장 사본이 원본보다 새로우면 MATERIAL_DIRTY (OV6)', async () => {
    // 창1에서 나.docx(편집 탭)를 AI 대상으로, 창2의 지연 탭 마.docx를 재료로
    assert.equal((await fake.invokeFrom(strip1, 'tabs:activate', tabOf(strip1, b).id)).ok, true);
    const old = new Date(Date.now() - DAY);
    utimesSync(e, old, old);
    const m0 = await fake.invokeFrom(panel1, 'ai:materials');
    assert.equal(m0.ok, true, m0.error);
    const item0 = m0.value.open.find((x) => basename(x.path) === '마.docx');
    assert.ok(item0, '지연 탭도 재료 후보');
    assert.equal(item0.dirty, false, '자동 저장 사본이 없으면 깨끗');
    put(draftPathFor(ud, e), 'PK');
    const m1 = await fake.invokeFrom(panel1, 'ai:materials');
    assert.equal(m1.value.open.find((x) => basename(x.path) === '마.docx')?.dirty, true);
    const r = await fake.invokeFrom(panel1, 'ai:send', { prompt: '참고해서 고쳐줘', editorOk: true, target: { edit: tabOf(strip1, b).path, read: [e] } });
    assert.equal(r.ok, false);
    assert.equal(r.code, 'MATERIAL_DIRTY', r.error);
    assert.equal(tabOf(strip2, e).loaded, false, '재료로 판정해도 불러오지 않는다');
    rmSync(draftPathFor(ud, e), { force: true });
  });

  /** settings.json의 tabSession.at을 지켜보며 다른 값이 몇 번 나왔는지 센다 */
  const watchWrites = async (ms) => {
    const seen = new Set();
    const end = Date.now() + ms;
    while (Date.now() < end) { seen.add(readSettings().tabSession?.at); await sleep(25); }
    return seen;
  };

  await t.test('세션 저장: 1초 안 여러 변경 → settings.json 한 번, 창 하나 닫기 → 목록에서 빠짐 (D9·D11)', async () => {
    await sleep(1200);
    const before = readSettings().tabSession;
    assert.equal(before.finalized, false);
    assert.equal(before.explorer, false);
    assert.equal(before.windows.length, 2);
    // 여러 번 바꾸기 (탭 활성·창 옮기기)
    for (const id of ['home', tabOf(strip1, a).id, tabOf(strip1, b).id]) await fake.invokeFrom(strip1, 'tabs:activate', id);
    w1.setBounds({ x: 140, y: 90, width: 1320, height: 860 });
    w1.emit('move');
    const seen = await watchWrites(1500);
    seen.delete(before.at);
    assert.equal(seen.size, 1, '한 번만 쓴다');
    const s = readSettings().tabSession;
    assert.deepEqual(s.windows[0].bounds, { x: 140, y: 90, width: 1320, height: 860 });
    assert.equal(s.windows[0].active, 2, '나.docx 활성');
    assert.deepEqual(s.windows[0].tabs.map((x) => basename(x.path)), ['가.hwpx', '나.docx']);
    assert.deepEqual(s.windows[1].tabs.map((x) => basename(x.path)), ['다.pdf', '마.docx'], '불러오지 않은 탭도 세션에 남는다');
    assert.equal(s.windows[1].maximized, true);
    // 창2 닫기 → 다른 창이 남으니 그 창만 빠진다. 지난 세션은 그대로
    w2.close();
    assert.ok(await until(() => w2.isDestroyed()));
    assert.ok(await until(() => readSettings().tabSession.windows.length === 1, 3000), '닫은 창은 목록에서 빠진다');
    assert.deepEqual(readSettings().lastTabSession, last, '지난 세션은 끝낼 때까지 그대로');
  });

  let w3 = null;
  await t.test('마지막 창 닫기 → 확정, 앱이 살아 있는 채로 새 탭 창이 열리면 다시 적는다', async () => {
    w1.close();
    assert.ok(await until(() => w1.isDestroyed()), '창1이 닫힌다 (깨끗한 편집 탭)');
    const s = readSettings();
    assert.equal(s.lastTabSession.finalized, true);
    assert.deepEqual(s.lastTabSession.windows.map((w) => w.tabs.map((x) => basename(x.path))), [['가.hwpx', '나.docx']], '마지막 창은 빼기 전에 확정');
    assert.deepEqual(s.tabSession, s.lastTabSession);
    // 탐색기에서 또 열기 → 새 탭 창 → 세션 기록을 다시 시작한다
    const n = fake.windows.length;
    fake.app.emit('second-instance', {}, [process.execPath, c], docs);
    assert.ok(await until(() => fake.windows.length > n), '새 탭 창');
    w3 = fake.windows.at(-1);
    assert.ok(await until(() => readSettings().tabSession.finalized === false, 3000), '다시 적기 (resume)');
    const now = readSettings();
    assert.deepEqual(now.tabSession.windows.map((w) => w.tabs.map((x) => basename(x.path))), [['다.pdf']]);
    assert.equal(now.lastTabSession.finalized, true, '지난 세션(확정한 것)은 그대로');
  });

  await t.test('끝내기: lastTabSession 확정, 끝내기 중 닫힘은 저장 안 함 (D11)', async () => {
    const quits = fake.app.quitCalls;
    const ev = { prevented: false, preventDefault() { this.prevented = true; } };
    fake.app.emit('before-quit', ev);
    assert.equal(ev.prevented, true, '차례 확인이 끝날 때까지 막는다');
    assert.ok(await until(() => fake.app.quitCalls === quits + 1), '모든 창 확인 통과 → 다시 app.quit');
    assert.equal(w3.isDestroyed(), true);
    await sleep(SETTLE);
    const s = readSettings();
    assert.equal(s.lastTabSession.finalized, true);
    assert.equal(s.lastTabSession.explorer, false);
    assert.deepEqual(s.lastTabSession.windows.map((w) => w.tabs.map((x) => basename(x.path))), [['다.pdf']], '끝내기 중 닫은 창이 빠지지 않는다');
    assert.deepEqual(s.tabSession, s.lastTabSession);
  });
});
const SETTLE = 1300;

test('강제 종료 뒤 시작: 확정 안 된 평소 세션을 복원한다', () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-tab-crash-'));
  try {
    const orig = join(dir, '옛 표.xls');
    writeFileSync(orig, 'x');
    const kept = join(dir, '남은.pdf');
    writeFileSync(kept, 'x');
    const settings = createSettings(dir);
    const older = { v: 1, at: 1, explorer: false, finalized: true, windows: [{ bounds: null, maximized: false, active: 0, tabs: [] }] };
    // 바꿔 연 탭: path는 시작 때 지우는 임시 사본, convertedFrom이 원본
    const crashed = { v: 1, at: 2, explorer: false, finalized: false, windows: [{ bounds: { x: 1, y: 2, width: 1200, height: 800 }, maximized: false, active: 2,
      tabs: [{ type: 'doc', path: kept }, { type: 'editor', path: join(dir, 'converted', 'x', '옛 표.xlsx'), convertedFrom: orig }] }] };
    settings.set({ lastTabSession: older, tabSession: crashed });
    assert.equal(promoteCrashed(createSettings(dir)), true);
    const again = createSettings(dir);
    assert.deepEqual(again.get('lastTabSession'), crashed);
    const plan = restorePlan(again.get('lastTabSession'));
    assert.deepEqual(plan.missing, []);
    assert.equal(plan.windows[0].active, 2, '바꿔 연 탭도 원본이 있으면 남는다');
    assert.deepEqual(plan.windows[0].tabs, [{ type: 'doc', path: kept }, { type: 'editor', path: orig, convertedFrom: orig }], '원본을 다시 바꿔 연다');
    // 탐색기 시작 세션이 강제 종료됐으면 올리지 않는다
    settings.set({ tabSession: { ...crashed, explorer: true } });
    assert.equal(promoteCrashed(createSettings(dir)), false);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
