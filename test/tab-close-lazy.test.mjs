// 닫는 창의 지연 탭 (F1): 창 닫기·끝내기·탭 닫기 중에 이웃 지연 편집 탭이 깨어나 유령 창을 만들거나 다른 창에 붙지 않는다.
// main.mjs는 한 프로세스에 한 번만 불러올 수 있으므로 한 시험 안에서 차례로 이어 간다 (창 닫기 → 불러오는 중 탭 닫기 → 끝내기).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createBlankDocument } from '../src/core/create.mjs';
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


test('닫는 창의 지연 편집 탭은 깨어나지 않는다: 창 닫기·끝내기에 유령 창 없음, 불러오는 중 닫은 탭은 되살아나지 않음 (F1)', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-tab-close-lazy-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  makeVendors(vendor);
  const file = async (name) => { const p = join(docs, name); await createBlankDocument(p, { title: name }); return p; };
  const a = await file('가.hwpx');
  const b = await file('나.hwpx');
  const c = await file('다.hwpx');
  const d = await file('라.hwpx');
  const f = await file('마.hwpx');
  const e = join(docs, '바.pdf');
  writeFileSync(e, makePdf(1));
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);
  const fontsModule = join(base, 'fake-fonts.mjs');
  writeFileSync(fontsModule, FAKE_FONTS);
  // 창1 [가(활성), 나], 창2 [다(활성), 마, 라, 바.pdf]
  const last = {
    v: 1, at: 1000, explorer: false, finalized: true,
    windows: [
      { bounds: { x: 100, y: 80, width: 1300, height: 850 }, maximized: false, active: 1,
        tabs: [{ type: 'editor', path: a }, { type: 'editor', path: b }] },
      { bounds: { x: 200, y: 120, width: 1100, height: 760 }, maximized: false, active: 1,
        tabs: [{ type: 'editor', path: c }, { type: 'editor', path: f }, { type: 'editor', path: d }, { type: 'doc', path: e }] },
    ],
  };
  put(join(ud, 'settings.json'), JSON.stringify({ lastTabSession: last }));

  const saved = { argv: process.argv, env: { ...process.env } };
  Object.assign(process.env, {
    OFFICEKIT_USER_DATA: ud, OFFICEKIT_EDITOR_VENDOR: vendor, OFFICEKIT_AI_MODULE: aiModule, OFFICEKIT_TEST_WS: join(base, 'ws'),
    OFFICEKIT_OFFICE_FONTS_MODULE: fontsModule,
  });
  for (const k of ['OFFICEKIT_SMOKE', 'OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_EDITOR_TEST', 'OFFICEKIT_DOC_WINDOW']) delete process.env[k];
  process.argv = [process.argv[0]];
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
  const [w1, w2] = fake.windows;
  const tabstripOf = (win) => fake.views.map((v) => v.webContents).find((wc) => /tabstrip\.html$/.test(wc.mainFrame.url) && win.contentView.children.some((x) => x.webContents === wc));
  const strip1 = tabstripOf(w1);
  const strip2 = tabstripOf(w2);
  const tabsOf = (strip) => strip.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [], active: null };
  const tabOf = (strip, p) => tabsOf(strip).tabs.find((x) => x.path && basename(x.path) === basename(p)) ?? null;
  const editorViews = () => fake.views.filter((v) => v.opts?.webPreferences?.partition && !v.webContents.isDestroyed());
  const liveWindows = () => fake.windows.filter((w) => !w.isDestroyed());
  // 만든 편집 뷰 수 (닫힌 것 포함): 닫는 창에서 지연 탭을 불러오기 시작했는지 본다
  const madeViews = () => fake.views.filter((v) => v.opts?.webPreferences?.partition).length;
  assert.ok(await until(() => tabOf(strip1, a)?.loaded && tabOf(strip2, c)?.loaded && editorViews().length === 2), '활성 탭 가·다만 편집 뷰');

  await t.test('창 닫기: 활성 편집 탭이 닫혀 이웃 지연 탭이 활성이 되어도 불러오지 않는다 (새 창·다른 창에 붙지 않음, 편집 entry 없음)', async () => {
    const n = fake.windows.length;
    const made = madeViews();
    w1.close();
    assert.ok(await until(() => w1.isDestroyed()), '창1이 닫힌다');
    await sleep(1500);
    assert.equal(madeViews(), made, '지연 탭 나.hwpx를 불러오지 않는다');
    assert.equal(fake.windows.length, n, '새 탭 창을 만들지 않는다');
    assert.equal(tabOf(strip2, b), null, '나.hwpx가 창2에 붙지 않는다');
    assert.equal(editorViews().length, 1, '남은 편집 뷰는 창2의 다.hwpx 하나');
  });

  await t.test('불러오는 중인 지연 탭을 닫으면 그대로 닫혀 있다 (되살아나지 않음)', async () => {
    const before = editorViews().length;
    const p1 = fake.invokeFrom(strip2, 'tabs:activate', tabOf(strip2, d).id);
    const p2 = fake.invokeFrom(strip2, 'tabs:close', tabOf(strip2, d).id);
    for (const r of await Promise.all([p1, p2])) assert.equal(r.ok, true, r.error);
    await sleep(1500);
    assert.equal(tabOf(strip2, d), null, '라.hwpx 탭은 닫힌 채');
    // 라.hwpx가 빠지면 이웃 바.pdf(문서 탭)가 활성이 된다 → 편집 뷰는 그대로 (탭 없는 편집 뷰가 남지 않는다)
    assert.equal(editorViews().length, before, '탭 없는 편집 뷰가 남지 않는다');
    assert.equal(fake.windows.length, 2);
    // 활성 탭을 다.hwpx로 (끝내기에서 다.hwpx를 닫으면 이웃 지연 탭 마.hwpx가 활성이 된다)
    assert.equal((await fake.invokeFrom(strip2, 'tabs:activate', tabOf(strip2, c).id)).ok, true);
    assert.equal(tabOf(strip2, f).loaded, false);
  });

  await t.test('끝내기: 지연 탭이 깨어나지 않아 app.quit 한 번, 유령 창 없음', async () => {
    const quits = fake.app.quitCalls;
    const n = fake.windows.length;
    const made = madeViews();
    const ev = { prevented: false, preventDefault() { this.prevented = true; } };
    fake.app.emit('before-quit', ev);
    assert.equal(ev.prevented, true);
    assert.ok(await until(() => fake.app.quitCalls === quits + 1), 'app.quit');
    await sleep(1500);
    assert.equal(fake.app.quitCalls, quits + 1, 'app.quit은 한 번');
    assert.equal(madeViews(), made, '지연 탭 마.hwpx를 불러오지 않는다');
    assert.equal(fake.windows.length, n, '새 탭 창을 만들지 않는다');
    assert.deepEqual(liveWindows(), [], '창이 다 닫혔다');
    assert.equal(editorViews().length, 0, '편집 뷰가 남지 않는다');
  });
});
