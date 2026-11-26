// 탭 창 껍데기(src/app/tab-window.mjs)와 탭 줄 화면(src/app/renderer/tabstrip.*) 시험
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import { makeFakeElectron } from './helpers/fake-electron.mjs';
import { createTabWindow, PANEL_MAX, PANEL_MIN, TABSTRIP_H } from '../src/app/tab-window.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => readFileSync(join(ROOT, rel), 'utf8');

// 1200×800 탭 창 하나와 가짜 electron
const make = (o = {}) => {
  const f = makeFakeElectron();
  const s = createTabWindow({ BaseWindow: f.BaseWindow, WebContentsView: f.WebContentsView, bounds: { x: 0, y: 0, width: 1200, height: 800 }, ...o });
  return { f, s };
};

test('상수: 탭 줄 높이·패널 너비 범위', () => {
  assert.equal(TABSTRIP_H, 40);
  assert.equal(PANEL_MIN, 320);
  assert.equal(PANEL_MAX, 640);
});

test('createTabWindow: 탭 줄 뷰가 맨 위 40px, tabstrip preload, 창 닫기는 막고 close-requested', () => {
  const f = makeFakeElectron();
  const s = createTabWindow({ BaseWindow: f.BaseWindow, WebContentsView: f.WebContentsView, bounds: { x: 0, y: 0, width: 1200, height: 800 } });
  assert.deepEqual(s.tabstrip.getBounds(), { x: 0, y: 0, width: 1200, height: 40 });
  assert.match(s.tabstrip.opts.webPreferences.preload, /tabstrip-preload\.cjs$/);
  let asked = 0; s.on('close-requested', () => asked++);
  s.win.close();
  assert.equal(asked, 1); assert.equal(s.win.isDestroyed(), false);
});

test('createTabWindow: 탭 줄 뷰는 sandbox·contextIsolation, tabstrip.html을 싣고, 이동·새 창을 막는다', () => {
  const { f, s } = make();
  const wp = s.tabstrip.opts.webPreferences;
  assert.equal(wp.sandbox, true);
  assert.equal(wp.contextIsolation, true);
  assert.equal(wp.nodeIntegration, false);
  assert.equal('backgroundThrottling' in wp, false, 'backgroundThrottling은 기본값');
  assert.equal(s.id, s.win.id);
  assert.ok(s.win.contentView.children.includes(s.tabstrip), '탭 줄 뷰가 창에 붙어 있다');
  const wc = s.tabstrip.webContents;
  assert.match(wc.mainFrame.url, /\/renderer\/tabstrip\.html$/);
  assert.deepEqual(wc.openHandler({ url: 'https://example.com' }), { action: 'deny' });
  let prevented = false;
  wc.emit('will-navigate', { preventDefault: () => { prevented = true; } });
  assert.equal(prevented, true);
  assert.equal(f.windows.length, 1);
});

test('attach·focusTab: 활성 content만 보이고 배치, 패널은 오른쪽, 접으면 숨김', () => {
  const { f, s } = make();
  const a = new f.WebContentsView(); const b = new f.WebContentsView(); const p = new f.WebContentsView();
  s.attach(a, 'content'); s.attach(b, 'content'); s.attach(p, 'panel');
  assert.equal(a.getVisible(), false, 'focusTab 전에는 숨김');
  assert.equal(b.getVisible(), false);
  assert.ok(s.win.contentView.children.includes(a) && s.win.contentView.children.includes(p));
  s.focusTab(b);
  assert.equal(b.getVisible(), true);
  assert.equal(a.getVisible(), false);
  assert.equal(b.webContents.isFocused(), true);
  assert.deepEqual(b.getBounds(), { x: 0, y: 40, width: 800, height: 760 });
  assert.deepEqual(p.getBounds(), { x: 800, y: 40, width: 400, height: 760 });
  assert.equal(p.getVisible(), true);
  s.focusTab(a);
  assert.equal(a.getVisible(), true);
  assert.equal(b.getVisible(), false);
  assert.deepEqual(a.getBounds(), { x: 0, y: 40, width: 800, height: 760 });

  assert.deepEqual(s.setPanel({ collapsed: true }), { width: 400, collapsed: true });
  assert.equal(p.getVisible(), false);
  assert.equal(a.getBounds().width, 1200);
  assert.deepEqual(s.setPanel({ collapsed: false, width: 9999 }), { width: 640, collapsed: false });
  assert.equal(p.getVisible(), true);
  assert.deepEqual(p.getBounds(), { x: 560, y: 40, width: 640, height: 760 });
  assert.equal(a.getBounds().width, 560);
  assert.deepEqual(s.setPanel({ width: 10 }), { width: 320, collapsed: false });
  assert.deepEqual(s.setPanel({}), { width: 320, collapsed: false }, '빈 값은 그대로');
});

test('createTabWindow: panelWidth·panelCollapsed 처음 값도 범위로 자르고 반영', () => {
  const { f, s } = make({ panelWidth: 100, panelCollapsed: true });
  const c = new f.WebContentsView(); const p = new f.WebContentsView();
  s.attach(c, 'content'); s.attach(p, 'panel');
  s.focusTab(c);
  assert.equal(p.getVisible(), false);
  assert.deepEqual(c.getBounds(), { x: 0, y: 40, width: 1200, height: 760 });
  assert.deepEqual(s.setPanel({ collapsed: false }), { width: 320, collapsed: false });
});

test('detach: 떼기만 하고 wc는 닫지 않는다 (D3)', () => {
  const { f, s } = make();
  const a = new f.WebContentsView();
  s.attach(a, 'content'); s.focusTab(a);
  s.detach(a);
  assert.equal(s.win.contentView.children.includes(a), false);
  assert.equal(a.webContents.isDestroyed(), false);
  // 뗀 뷰는 더 배치하지 않는다
  const before = a.getBounds();
  s.win.setBounds({ width: 1000, height: 600 }); s.win.emit('resize');
  assert.deepEqual(a.getBounds(), before);
  s.detach(a); // 두 번 떼도 괜찮다
});

test('resize: 창 크기가 바뀌면 다시 배치', () => {
  const { f, s } = make();
  const a = new f.WebContentsView(); const p = new f.WebContentsView();
  s.attach(a, 'content'); s.attach(p, 'panel'); s.focusTab(a);
  s.win.setBounds({ width: 1000, height: 600 });
  s.win.emit('resize');
  assert.deepEqual(s.tabstrip.getBounds(), { x: 0, y: 0, width: 1000, height: 40 });
  assert.deepEqual(a.getBounds(), { x: 0, y: 40, width: 600, height: 560 });
  assert.deepEqual(p.getBounds(), { x: 600, y: 40, width: 400, height: 560 });
});

test('dialogParent는 BaseWindow, setTitle은 창 제목', () => {
  const { s } = make();
  assert.equal(s.dialogParent(), s.win);
  s.setTitle('보고서.hwpx - Office Kit');
  assert.equal(s.win.getTitle(), '보고서.hwpx - Office Kit');
});

test('focus·closed 이벤트를 낸다', () => {
  const { s } = make();
  let focus = 0; let closed = 0;
  s.on('focus', () => focus++); s.on('closed', () => closed++);
  s.win.emit('focus');
  assert.equal(focus, 1);
  s.win.destroy();
  assert.equal(closed, 1);
});

test('sendTabs: 탭 줄 wc에 tabs:list', () => {
  const { s } = make();
  const payload = { tabs: [{ id: 'home', type: 'home', title: '홈', dirty: false, loaded: true, zone3: false, busy: false }], active: 'home', panelCollapsed: false };
  s.sendTabs(payload);
  assert.deepEqual(s.tabstrip.webContents.sent, [{ channel: 'tabs:list', data: payload }]);
});

test('destroy: 탭 줄 wc destroyed, 창 closed 한 번', () => {
  const { f, s } = make();
  const a = new f.WebContentsView();
  s.attach(a, 'content');
  let closed = 0; s.on('closed', () => closed++);
  s.destroy();
  assert.equal(s.tabstrip.webContents.isDestroyed(), true);
  assert.equal(s.win.isDestroyed(), true);
  assert.equal(s.win.contentView.children.length, 0, '자식 뷰를 뗀다');
  assert.equal(a.webContents.isDestroyed(), false, '뷰 wc 정리는 주인 몫');
  s.destroy();
  assert.equal(closed, 1);
  s.sendTabs({ tabs: [], active: null, panelCollapsed: false }); // 닫힌 뒤에는 보내지 않고 조용히
  assert.equal(s.tabstrip.webContents.sent.length, 0);
});

test('탭 줄 화면: focusAfterKey — ←/→ 돌아감, Home/End, 그 밖은 그대로', () => {
  const ctx = {}; vm.runInNewContext(readFileSync(join(ROOT, 'src/app/renderer/tabstrip.js'), 'utf8'), ctx);
  const f = ctx.OKTabstrip.focusAfterKey;
  assert.equal(f(0, 'ArrowLeft', 3), 2); assert.equal(f(2, 'ArrowRight', 3), 0);
  assert.equal(f(1, 'Home', 3), 0); assert.equal(f(1, 'End', 3), 2); assert.equal(f(1, 'a', 3), 1);
  assert.equal(f(1, 'ArrowLeft', 3), 0); assert.equal(f(1, 'ArrowRight', 3), 2);
});

test('탭 줄 화면: tabStopIndex — 활성 탭이 0, 활성이 없거나 목록에 없으면 첫 탭', () => {
  const ctx = {}; vm.runInNewContext(readFileSync(join(ROOT, 'src/app/renderer/tabstrip.js'), 'utf8'), ctx);
  const s = ctx.OKTabstrip.tabStopIndex;
  assert.equal(s(['home', 't1', 't2'], 't1'), 1);
  assert.equal(s(['home', 't1', 't2'], null), 0);
  assert.equal(s(['home', 't1', 't2'], 'gone'), 0);
  assert.equal(s([], null), -1);
});

test('탭 줄 화면: role=tablist·role=tab·aria-selected·aria-label, 인라인 스크립트 없음(CSP)', () => {
  const html = read('src/app/renderer/tabstrip.html');
  assert.match(html, /<script src="tabstrip\.js" defer><\/script>/);
  assert.match(html, /<meta http-equiv="Content-Security-Policy" content="[^"]*script-src 'self'/);
  assert.doesNotMatch(html, /<script(?![^>]*\bsrc=)[^>]*>/, '본문 있는 <script> 없음');
  assert.doesNotMatch(html, /\son[a-z]+=/i, '인라인 이벤트 속성 없음');
  assert.match(html, /role="tablist"/);
  assert.match(html, /aria-label="열린 문서"/);
  const js = read('src/app/renderer/tabstrip.js');
  for (const s of ['tablist', 'aria-selected', "'tab'", 'aria-pressed', 'tabindex', 'Delete', '저장 안 한 변경 있음', '아직 불러오지 않음', '인터넷 파일', 'AI 도우미']) {
    assert.ok(html.includes(s) || js.includes(s), `탭 줄 화면에 ${s} 없음`);
  }
  assert.match(js, /‘\$\{[^}]+\}’ 닫기/, '닫기 단추 이름');
  assert.match(read('src/app/renderer/tabstrip.css'), /:focus-visible/);
});

test('탭 줄 preload: window.okTabs만 노출하고 채널은 표 블록으로 검사', () => {
  const src = read('src/app/tabstrip-preload.cjs');
  assert.match(src, /exposeInMainWorld\('okTabs'/);
  for (const k of ['activate', 'close', 'home', 'togglePanel', 'ready', 'onList', 'i18n']) assert.match(src, new RegExp('\\b' + k + ':'), k);
  assert.match(src, /const need = \(call, channel\)/);
});

test('탭 줄 화면이 죽으면 다시 싣는다, 정상 종료·닫힌 창은 그대로 (F10)', () => {
  const { f, s } = make();
  const wc = s.tabstrip.webContents;
  const loads = [];
  const orig = wc.loadFile.bind(wc);
  wc.loadFile = (file) => { loads.push(file); return orig(file); };
  const err0 = console.error;
  console.error = () => {};
  try {
    f.crash(wc, 'crashed');
    assert.equal(loads.length, 1);
    assert.match(loads[0], /tabstrip\.html$/);
    f.crash(wc, 'clean-exit');
    assert.equal(loads.length, 1, '정상 종료는 다시 싣지 않는다');
  } finally {
    console.error = err0;
  }
});
