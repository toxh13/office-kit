// 강제 종료 뒤 시작 (F7·F4): 확정 못 한 평소 세션을 올려 복원하되 활성 탭도 불러오지 않고 홈을 보인다
// (앱을 죽이는 문서가 시작마다 다시 열리지 않게). 네트워크 경로(UNC) 탭은 있는지 보지 않고 지연 탭으로 남긴다.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
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


test('강제 종료 뒤 시작: 올린 세션은 활성 탭도 불러오지 않고 홈, UNC 탭은 지연 탭으로 남는다 (F7·F4)', async (t) => {
  const base = mkdtempSync(join(tmpdir(), 'ok-tab-crash-restore-'));
  const ud = join(base, 'ud');
  const docs = join(base, '문서');
  const vendor = join(base, 'vendor');
  mkdirSync(docs, { recursive: true });
  makeVendors(vendor);
  const a = join(docs, '죽이는 문서.hwpx');
  await createBlankDocument(a, { title: '가' });
  const p = join(docs, '보통.pdf');
  writeFileSync(p, makePdf(1));
  const unc = '\\\\office-kit-test-server\\share\\보고서.pdf';
  const aiModule = join(base, 'fake-ai.mjs');
  writeFileSync(aiModule, FAKE_AI);
  const fontsModule = join(base, 'fake-fonts.mjs');
  writeFileSync(fontsModule, FAKE_FONTS);
  // 확정 못 한 평소 세션 (강제 종료): 활성 탭 = 죽이는 문서.hwpx
  const crashed = {
    v: 1, at: 2000, explorer: false, finalized: false,
    windows: [{ bounds: { x: 100, y: 80, width: 1300, height: 850 }, maximized: false, active: 1,
      tabs: [{ type: 'editor', path: a }, { type: 'doc', path: unc }, { type: 'doc', path: p }] }],
  };
  put(join(ud, 'settings.json'), JSON.stringify({ lastTabSession: { v: 1, at: 1, explorer: false, finalized: true, windows: [] }, tabSession: crashed }));
  const readSettings = () => JSON.parse(readFileSync(join(ud, 'settings.json'), 'utf8'));

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
  assert.ok(await until(() => fake.windows.length >= 1, 20_000), '탭 창');
  const [w1] = fake.windows;
  const strip = fake.views.map((v) => v.webContents).find((wc) => /tabstrip\.html$/.test(wc.mainFrame.url) && w1.contentView.children.some((x) => x.webContents === wc));
  const tabsOf = () => strip.sent.filter((s) => s.channel === 'tabs:list').at(-1)?.data ?? { tabs: [], active: null };
  const tabOf = (name) => tabsOf().tabs.find((x) => x.path && basename(x.path) === name) ?? null;
  const editorViews = () => fake.views.filter((v) => v.opts?.webPreferences?.partition && !v.webContents.isDestroyed());
  await sleep(500);

  assert.deepEqual(readSettings().lastTabSession.at, 2000, '강제 종료한 세션을 지난 세션으로 올렸다');
  assert.deepEqual(tabsOf().tabs.map((x) => [x.type, x.path && basename(x.path), x.loaded]),
    [['home', null, true], ['editor', '죽이는 문서.hwpx', false], ['doc', '보고서.pdf', false], ['doc', '보통.pdf', false]], 'UNC 탭도 남는다 (있는지 보지 않음)');
  assert.equal(tabsOf().active, 'home', '홈을 보인다');
  assert.equal(editorViews().length, 0, '활성이던 탭도 불러오지 않는다');
  const homeView = w1.contentView.children.find((v) => v.webContents?.mainFrame.url.endsWith('?view=home'));
  assert.equal(homeView.visible, true);

  // 사용자가 그 탭을 누르면 그때 불러온다
  const r = await fake.invokeFrom(strip, 'tabs:activate', tabOf('죽이는 문서.hwpx').id);
  assert.equal(r.ok, true, r.error);
  assert.ok(await until(() => editorViews().length === 1 && tabOf('죽이는 문서.hwpx')?.loaded), '누르면 편집 뷰');
});
