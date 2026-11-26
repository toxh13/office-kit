// tab-window.mjs — 탭 창 껍데기: BaseWindow + 맨 위 탭 줄 뷰 + 뷰 배치.
// 문서 상태는 두지 않는다. 문서 뷰(WebContentsView)와 그 정리는 뷰 주인(editor-window.mjs 등)이 맡는다 (D3).
// 배치: 탭 줄 {0,0,W,40} · 활성 content {0,40,W-P,H-40} · 패널 {W-P,40,P,H-40} (패널이 접혔거나 없으면 P = 0)
import { EventEmitter } from 'node:events';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { channelSpec } from './ipc-channels.mjs';

const here = dirname(fileURLToPath(import.meta.url));

export const TABSTRIP_H = 40;
export const PANEL_MIN = 320;
export const PANEL_MAX = 640;

// 탭 줄에 보내는 채널은 표에 있어야 한다 (없으면 불러올 때 던진다)
channelSpec('tabs:list', 'send');

/** 패널 너비를 PANEL_MIN..PANEL_MAX로 자른다. 숫자가 아니면 fallback */
const clampPanel = (w, fallback) => {
  const n = Number(w);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(PANEL_MAX, Math.max(PANEL_MIN, Math.round(n)));
};

export function createTabWindow({
  BaseWindow, WebContentsView, bounds, maximized = false, icon, dark = false, devTools = false,
  panelWidth = 400, panelCollapsed = false, show = true,
}) {
  const shell = new EventEmitter();
  const win = new BaseWindow({
    ...(bounds ?? {}), minWidth: 1000, minHeight: 680, title: 'Office Kit', show: false,
    backgroundColor: dark ? '#16181d' : '#f4f5f7',
    autoHideMenuBar: true,
    ...(icon ? { icon } : {}),
  });

  // 탭 줄 뷰: sandbox·contextIsolation, 탭 줄 채널만 여는 preload. backgroundThrottling은 기본값(숨으면 멈춘다)
  const tabstrip = new WebContentsView({
    webPreferences: {
      preload: join(here, 'tabstrip-preload.cjs'),
      contextIsolation: true, nodeIntegration: false, sandbox: true,
      webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
      devTools: !!devTools,
    },
  });
  tabstrip.setBackgroundColor?.(dark ? '#16181d' : '#f4f5f7');
  win.contentView.addChildView(tabstrip);
  const wc = tabstrip.webContents;
  // 새 창·이동 막기 (main.mjs 주 창 규칙과 같다. 탭 줄에는 링크가 없으므로 모두 거절)
  wc.setWindowOpenHandler(() => ({ action: 'deny' }));
  wc.on('will-navigate', (e) => e.preventDefault());
  wc.on('will-frame-navigate', (e) => e.preventDefault());
  wc.on('will-attach-webview', (e) => e.preventDefault());
  const loadStrip = () => { void Promise.resolve(wc.loadFile(join(here, 'renderer', 'tabstrip.html'))).catch(() => {}); };
  loadStrip();
  // 탭 줄 화면이 죽으면(메모리 부족 등) 다시 싣는다 — 다시 뜨면 tabs:ready로 탭 목록을 다시 받는다
  wc.on('render-process-gone', (_e, d) => {
    console.error('[tabstrip] 화면 종료:', d?.reason);
    if (closed || wc.isDestroyed() || d?.reason === 'clean-exit') return;
    loadStrip();
  });

  const views = new Map(); // 붙인 뷰 → 'content' | 'panel'
  let active = null; // 보이는 content 뷰
  const panel = { width: clampPanel(panelWidth, 400), collapsed: !!panelCollapsed };
  let closed = false;

  function layout() {
    if (closed || win.isDestroyed()) return;
    const { width: W, height: H } = win.getContentBounds();
    const bodyH = Math.max(0, H - TABSTRIP_H);
    tabstrip.setBounds({ x: 0, y: 0, width: W, height: TABSTRIP_H });
    const hasPanel = [...views.values()].includes('panel');
    const P = panel.collapsed || !hasPanel ? 0 : Math.min(panel.width, W);
    for (const [view, role] of views) {
      if (role === 'panel') {
        view.setVisible(!panel.collapsed);
        if (!panel.collapsed) view.setBounds({ x: W - P, y: TABSTRIP_H, width: P, height: bodyH });
      } else if (view === active) {
        view.setBounds({ x: 0, y: TABSTRIP_H, width: W - P, height: bodyH });
        view.setVisible(true);
      } else {
        view.setVisible(false);
      }
    }
  }

  // 창 닫기는 늘 막고 알린다 — 닫을지는 main이 정한다
  win.on('close', (e) => { e.preventDefault(); shell.emit('close-requested'); });
  win.on('focus', () => shell.emit('focus'));
  win.on('resize', layout);
  win.on('maximize', layout);
  win.on('unmaximize', layout);
  win.on('closed', () => {
    closed = true;
    views.clear();
    active = null;
    if (!wc.isDestroyed()) wc.close();
    shell.emit('closed');
  });

  Object.assign(shell, {
    id: win.id,
    win,
    tabstrip,
    /** 뷰를 창에 붙인다. content는 focusTab 전까지 숨긴다 */
    attach(view, role) {
      if (role !== 'content' && role !== 'panel') throw new Error('탭 창 뷰 역할이 잘못됨: ' + role);
      if (closed) return;
      views.set(view, role);
      if (role === 'content' && view !== active) view.setVisible(false);
      win.contentView.addChildView(view);
      layout();
    },
    /** 떼기만 한다 (wc를 닫지 않는다 — 닫기는 뷰 주인 몫, D3) */
    detach(view) {
      if (!views.has(view)) return;
      views.delete(view);
      if (active === view) active = null;
      if (!win.isDestroyed()) win.contentView.removeChildView(view);
      layout();
    },
    /** 이 content 뷰만 보이고 나머지 content는 숨긴다 (Section 4 #4) */
    focusTab(view) {
      if (views.get(view) !== 'content') throw new Error('탭 창에 붙은 content 뷰가 아님');
      active = view;
      layout();
      view.webContents?.focus?.();
    },
    setTitle(title) { if (!win.isDestroyed()) win.setTitle(String(title ?? '')); },
    dialogParent() { return win; },
    /** 패널 너비·접힘 → 지금 값 { width, collapsed } */
    setPanel({ width, collapsed } = {}) {
      if (width !== undefined) panel.width = clampPanel(width, panel.width);
      if (collapsed !== undefined) panel.collapsed = !!collapsed;
      layout();
      return { ...panel };
    },
    /** 탭 목록을 탭 줄 화면에 보낸다 */
    sendTabs(payload) {
      if (closed || wc.isDestroyed()) return;
      tabstrip.webContents.send('tabs:list', payload);
    },
    /** 탭 줄 wc를 닫고 남은 뷰를 뗀 뒤 창을 없앤다 (뷰 wc 정리는 주인이 그 전에 끝낸다) */
    destroy() {
      if (closed || win.isDestroyed()) return;
      if (!wc.isDestroyed()) wc.close();
      for (const view of [...views.keys(), tabstrip]) win.contentView.removeChildView(view);
      views.clear();
      active = null;
      win.destroy();
    },
  });

  layout();
  if (show) {
    if (maximized) win.maximize();
    else win.show();
  }
  return shell;
}
