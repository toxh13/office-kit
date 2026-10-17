// app.js — Office Kit 화면. 바닐라 JS, window.officeKit(preload)만 쓴다.
'use strict';
(() => {
  const K = window.officeKit;
  const MD = window.OKMarkdown;

  // ── 작은 도우미 ──────────────────────────────────────

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  /** 요소 만들기: h('div', { class: 'a', text: '글', onclick }, ...자식) */
  function h(tag, props = {}, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props ?? {})) {
      if (v === undefined || v === null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'html') el.innerHTML = v; // md.js가 이스케이프한 결과에만 쓴다
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, String(v));
    }
    for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) el.append(c.nodeType ? c : String(c));
    return el;
  }

  const SVGNS = 'http://www.w3.org/2000/svg';
  function icon(name, cls = 'i') {
    const svg = document.createElementNS(SVGNS, 'svg');
    svg.setAttribute('class', cls);
    svg.setAttribute('aria-hidden', 'true');
    const use = document.createElementNS(SVGNS, 'use');
    use.setAttribute('href', `#i-${name}`);
    svg.append(use);
    return svg;
  }
  const spinner = () => h('span', { class: 'spinner', 'aria-hidden': 'true' });

  /** preload 응답 { ok, value } 풀기 */
  async function call(promise) {
    const r = await promise;
    if (!r || r.ok !== true) {
      const err = new Error(r?.error || '알 수 없는 오류가 났습니다');
      err.code = r?.code ?? null;
      throw err;
    }
    return r.value;
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const pathKey = (p) => String(p ?? '').toLowerCase();
  const extOf = (p) => { const m = /\.([^.\\/]+)$/.exec(String(p ?? '')); return m ? m[1].toLowerCase() : ''; };
  const baseName = (p) => String(p ?? '').split(/[\\/]/).pop();
  const dirName = (p) => { const s = String(p ?? ''); const i = Math.max(s.lastIndexOf('\\'), s.lastIndexOf('/')); return i > 0 ? s.slice(0, i) : s; };

  const KIND_BY_EXT = {
    hwp: '한글', hwpx: '한글', hwt: '한글', docx: 'Word', doc: 'Word', rtf: 'Word', odt: 'Word',
    pptx: 'PowerPoint', ppt: 'PowerPoint', odp: 'PowerPoint', xlsx: 'Excel', xls: 'Excel', xlsm: 'Excel', csv: 'Excel', ods: 'Excel',
  };
  const KIND_STYLE = {
    한글: { cls: 'k-hwp', letter: '한' }, Word: { cls: 'k-word', letter: 'W' },
    PowerPoint: { cls: 'k-ppt', letter: 'P' }, Excel: { cls: 'k-xls', letter: 'X' },
  };
  const kindStyle = (kind) => KIND_STYLE[kind] ?? { cls: 'k-etc', letter: '?' };

  const ENGINE_LABEL = {
    rhwp: '내장 엔진 (rhwp)', file: '내장 엔진', word: 'Microsoft Word', excel: 'Microsoft Excel',
    ppt: 'PowerPoint', powerpoint: 'PowerPoint', hwp: '한컴오피스', hancom: '한컴오피스',
  };
  const ENGINE_SHORT = { rhwp: '내장 엔진', file: '내장 엔진', word: 'Word 변환', excel: 'Excel 변환', ppt: 'PowerPoint 변환', powerpoint: 'PowerPoint 변환', hwp: '한컴 변환', hancom: '한컴 변환' };

  // AI 도구 이름 → 화면에 보일 말
  const TOOL_LABEL = {
    office_status: ['상태 확인 중', '상태 확인'],
    office_info: ['문서 정보 보는 중', '문서 정보 보기'],
    office_read: ['문서 읽는 중', '문서 읽기'],
    office_fields: ['채울 칸 찾는 중', '채울 칸 찾기'],
    office_fill: ['칸 채우는 중', '칸 채우기'],
    office_replace: ['바꾸는 중', '찾아 바꾸기'],
    office_set_cells: ['셀 고치는 중', '셀 고치기'],
    office_convert: ['변환 중', '형식 바꾸기'],
    office_create: ['새 문서 만드는 중', '새 문서 만들기'],
    office_page_image: ['결과 확인 중', '결과 확인'],
    hwp_edit: ['한글 편집 중', '한글 편집'],
  };
  const EDIT_TOOLS = new Set(['office_fill', 'office_replace', 'office_set_cells', 'hwp_edit', 'office_create', 'office_convert']);
  // mcp__office-kit__office_read, office-kit.office_read 같은 이름도 받아 준다
  const toolBase = (name) => String(name ?? '').split(/__|[./:]/).filter(Boolean).pop() ?? '';
  /** part = { tool, name } (tool: office_read 같은 원래 이름, name: ai.mjs가 붙인 표시 이름) */
  function toolLabel(part, done) {
    const base = toolBase(part.tool || part.name);
    const l = TOOL_LABEL[base];
    if (l) return done ? l[1] : l[0];
    const label = part.name && part.name !== base ? part.name : base || '도구';
    return done ? label : `${label} 중`;
  }

  // 시간 표시
  const two = (n) => String(n).padStart(2, '0');
  function timeText(ms) {
    const d = new Date(ms);
    const hr = d.getHours();
    return `${hr < 12 ? '오전' : '오후'} ${hr % 12 || 12}:${two(d.getMinutes())}`;
  }
  function dayDiff(ms) {
    const a = new Date(ms); a.setHours(0, 0, 0, 0);
    const b = new Date(); b.setHours(0, 0, 0, 0);
    return Math.round((b - a) / 86_400_000);
  }
  function whenText(ms) {
    if (!ms) return '';
    const dd = dayDiff(ms);
    const d = new Date(ms);
    if (dd === 0) return `오늘 ${timeText(ms)}`;
    if (dd === 1) return `어제 ${timeText(ms)}`;
    if (d.getFullYear() === new Date().getFullYear()) return `${d.getMonth() + 1}월 ${d.getDate()}일 ${timeText(ms)}`;
    return `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}. ${timeText(ms)}`;
  }
  function relText(ms) {
    if (!ms) return '';
    const s = (Date.now() - ms) / 1000;
    if (s < 60) return '방금';
    if (s < 3600) return `${Math.floor(s / 60)}분 전`;
    const dd = dayDiff(ms);
    if (dd === 0) return `${Math.floor(s / 3600)}시간 전`;
    if (dd === 1) return '어제';
    if (dd < 7) return `${dd}일 전`;
    const d = new Date(ms);
    return d.getFullYear() === new Date().getFullYear() ? `${d.getMonth() + 1}월 ${d.getDate()}일` : `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`;
  }

  // ── 상태 ─────────────────────────────────────────────

  const state = {
    exts: [],
    settings: {},
    version: '',
    doc: null,
    preview: null,
    previewSeq: 0,
    rendering: null,        // 렌더링 중인 문서 경로
    pendingRefresh: null,
    zoom: 'auto',
    providers: null,        // { providers, selected } 또는 null
    providerError: null,
    chats: new Map(),       // 문서 경로 → { items, busy, chatId }
    busy: new Map(),
    smoke: null,
    smokeSent: false,
    demo: false,
  };

  const el = {
    app: $('#app'), workspace: $('#workspace'), viewer: $('#viewer'),
    empty: $('#empty'), preview: $('#preview'), zoombar: $('#zoombar'), refreshing: $('#refreshing'),
    recent: $('#recent'), recentList: $('#recent-list'),
    docinfo: $('#docinfo'), docKind: $('#doc-kind'), docName: $('#doc-name'), docEngine: $('#doc-engine'),
    chat: $('#chat'), chatBody: $('#chat-body'), chatLog: $('#chat-log'), provider: $('#provider'),
    composer: $('#composer'), prompt: $('#prompt'), send: $('#btn-send'), composerHint: $('#composer-hint'),
    splitter: $('#splitter'), toggleChat: $('#btn-toggle-chat'),
    stEngine: $('#st-engine'), stPages: $('#st-pages'), stSaved: $('#st-saved'),
    stBusy: $('#st-busy'), stBusyText: $('#st-busy-text'), stReady: $('#st-ready'),
    overlay: $('#drop-overlay'), toasts: $('#toasts'),
    settings: $('#dlg-settings'), confirm: $('#dlg-confirm'),
  };

  // ── 바쁨 표시 · 알림 ─────────────────────────────────

  function setBusy(key, text) { state.busy.delete(key); state.busy.set(key, text); renderBusy(); }
  function clearBusy(key) { state.busy.delete(key); renderBusy(); }
  function renderBusy() {
    const last = [...state.busy.values()].pop();
    el.stBusy.hidden = !last;
    el.stReady.hidden = !!last;
    el.stBusyText.textContent = last ?? '';
  }

  const TOAST_ICON = { success: 'check', error: 'alert', info: 'info', warn: 'alert' };
  function toast(message, { type = 'info', action, timeout } = {}) {
    const t = h('div', { class: `toast ${type}`, role: type === 'error' ? 'alert' : 'status' },
      icon(TOAST_ICON[type] ?? 'info'), h('span', { class: 'toast-msg', text: message }));
    const close = () => {
      if (t.classList.contains('out')) return;
      t.classList.add('out');
      setTimeout(() => t.remove(), 180);
    };
    if (action) t.append(h('button', { class: 'toast-action', text: action.label, onclick: () => { close(); action.run(); } }));
    t.append(h('button', { class: 'toast-close', 'aria-label': '닫기', onclick: close }, icon('x')));
    el.toasts.append(t);
    while (el.toasts.children.length > 4) el.toasts.firstElementChild.remove();
    setTimeout(close, timeout ?? (type === 'error' ? 7000 : action ? 6000 : 3500));
    return close;
  }

  function showError(err, prefix) {
    const message = prefix ? `${prefix}: ${err.message}` : err.message;
    if (err.code === 'HANCOM_SETUP') toast(message, { type: 'warn', action: { label: '설정 열기', run: () => openSettings('hancom') } });
    else toast(message, { type: 'error' });
  }

  // ── 확인 창 ──────────────────────────────────────────

  /** body: 문자열 | 노드 배열. → Promise<boolean> */
  function confirmDialog({ title, body, ok = '확인', cancel = '취소', tone = 'info' }) {
    const dlg = el.confirm;
    $('#confirm-title').textContent = title;
    const bodyEl = $('#confirm-body');
    bodyEl.replaceChildren(...[].concat(body).map((b) => (typeof b === 'string' ? h('p', { text: b }) : b)));
    const ic = $('#confirm-icon');
    ic.className = `confirm-icon ${tone}`;
    ic.replaceChildren(icon(tone === 'danger' ? 'restore' : tone === 'shield' ? 'shield' : 'info'));
    const okBtn = $('#confirm-ok');
    const cancelBtn = $('#confirm-cancel');
    okBtn.textContent = ok;
    okBtn.className = `btn ${tone === 'danger' ? 'danger' : 'primary'}`;
    cancelBtn.textContent = cancel;
    return new Promise((resolve) => {
      const done = (v) => {
        okBtn.removeEventListener('click', yes);
        cancelBtn.removeEventListener('click', no);
        dlg.removeEventListener('close', no);
        if (dlg.open) dlg.close();
        resolve(v);
      };
      const yes = () => done(true);
      const no = () => done(false);
      okBtn.addEventListener('click', yes);
      cancelBtn.addEventListener('click', no);
      dlg.addEventListener('close', no);
      dlg.showModal();
      cancelBtn.focus();
    });
  }

  // ── 빈 화면 · 최근 문서 ──────────────────────────────

  function renderRecent(list) {
    const items = Array.isArray(list) ? list : [];
    el.recent.hidden = items.length === 0;
    el.recentList.replaceChildren(...items.map((r) => {
      const kind = KIND_BY_EXT[extOf(r.path)];
      const ks = kindStyle(kind);
      const open = h('button', {
        class: 'recent-open', title: r.exists ? r.path : `찾을 수 없음: ${r.path}`,
        'aria-disabled': r.exists ? null : 'true',
        onclick: () => {
          if (!r.exists) { toast('파일을 찾을 수 없어요. 옮겨졌거나 지워졌을 수 있어요.', { type: 'warn', action: { label: '목록에서 빼기', run: () => removeRecent(r.path) } }); return; }
          openPath(r.path);
        },
      },
      h('span', { class: `file-tile ${ks.cls}`, text: ks.letter, 'aria-hidden': 'true' }),
      h('span', { class: 'recent-text' },
        h('span', { class: 'recent-name', text: r.name ?? baseName(r.path) }),
        h('span', { class: 'recent-dir' }, h('bdi', { text: dirName(r.path) }))),
      r.exists ? h('span', { class: 'recent-meta', text: relText(r.openedAt) }) : h('span', { class: 'recent-missing', text: '찾을 수 없음' }));
      const remove = h('button', { class: 'icon-btn sm recent-remove', title: '목록에서 빼기', 'aria-label': `${r.name} 목록에서 빼기`, onclick: (e) => { e.stopPropagation(); removeRecent(r.path); } }, icon('x'));
      return h('li', { class: `recent-item${r.exists ? '' : ' missing'}` }, open, remove);
    }));
  }

  async function loadRecent() {
    try { renderRecent(await call(K.recent.list())); } catch { renderRecent([]); }
  }
  async function removeRecent(p) {
    try { renderRecent(await call(K.recent.remove(p))); } catch (err) { showError(err); }
  }

  // ── 문서 열기 · 닫기 ─────────────────────────────────

  async function pickFile() {
    try { await call(K.doc.pick()); } catch (err) { showError(err); }
  }

  async function openPath(p) {
    if (!p) return;
    const ext = extOf(p);
    if (!state.exts.includes(ext)) {
      toast(`지원하지 않는 형식이에요${ext ? ` (.${ext})` : ''}. 한글·Word·PowerPoint·Excel 문서를 열어 주세요.`, { type: 'warn' });
      return;
    }
    try { await call(K.doc.open(p)); } catch (err) {
      showError(err);
      if (err.code === 'NOT_FOUND') loadRecent();
    }
  }

  function onDocOpened(doc) {
    // 열려 있는 문서를 다시 연 경우: 미리보기는 그대로 두고 바뀌었을 때만 새로 그린다
    if (doc && state.doc && pathKey(doc.path) === pathKey(state.doc.path)) {
      state.doc = doc;
      renderTopbar();
      renderStatus();
      refreshPreview({ ifChanged: true, reason: 'reopen' });
      return;
    }
    state.doc = doc;
    state.preview = null;
    state.zoom = 'auto';
    state.previewSeq++;
    state.pendingRefresh = null;
    document.title = doc ? `${doc.name} - Office Kit` : 'Office Kit';
    renderTopbar();
    renderStatus();
    renderChat();
    if (!doc) {
      el.preview.hidden = true;
      el.preview.replaceChildren();
      el.zoombar.hidden = true;
      el.refreshing.hidden = true;
      el.empty.hidden = false;
      loadRecent();
      return;
    }
    el.empty.hidden = true;
    el.preview.hidden = false;
    if (doc.backupError) toast(doc.backupError, { type: 'warn' });
    refreshPreview({ reason: 'open' });
  }

  function updateDoc(doc) {
    if (!doc || !state.doc || pathKey(doc.path) !== pathKey(state.doc.path)) return;
    state.doc = { ...state.doc, ...doc };
    renderTopbar();
    renderStatus();
  }

  function renderTopbar() {
    const d = state.doc;
    el.docinfo.hidden = !d;
    for (const id of ['btn-refresh', 'btn-openapp', 'btn-pdf', 'btn-folder']) $(`#${id}`).disabled = !d;
    const restore = $('#btn-restore');
    restore.disabled = !d || !d.backup || !!state.restoring;
    restore.title = d?.backup ? `${whenText(d.backup.time)}에 만든 백업으로 되돌리기` : '이 문서의 백업이 없어요';
    if (!d) return;
    const ks = kindStyle(d.kind);
    el.docKind.className = `kind-badge ${ks.cls}`;
    el.docKind.textContent = d.kind;
    el.docName.textContent = d.name;
    el.docName.title = d.path;
    const eng = state.preview?.engine;
    el.docEngine.hidden = !eng;
    el.docEngine.textContent = ENGINE_SHORT[eng] ?? eng ?? '';
    el.docEngine.title = eng ? `미리보기: ${ENGINE_LABEL[eng] ?? eng}` : '';
  }

  function renderStatus() {
    const d = state.doc;
    const p = state.preview;
    const set = (node, text, title) => {
      node.hidden = !text;
      node.lastElementChild.textContent = text ?? '';
      node.title = title ?? '';
    };
    set(el.stEngine, d && p ? (ENGINE_LABEL[p.engine] ?? p.engine) : '', '미리보기를 만든 엔진');
    let pages = '';
    if (d && p) {
      if (p.kind === 'svg') pages = p.total > p.pages.length ? `${p.total}쪽 (앞 ${p.pages.length}쪽 표시)` : `${p.total}쪽`;
      else if (p.kind === 'pdf') pages = p.pages ? `${p.pages}쪽` : '';
      else if (p.sheets?.length) pages = `시트 ${p.sheets.length}개`;
      else if (p.slides?.length) pages = `슬라이드 ${p.slides.length}장`;
      else if (p.pages) pages = `${p.pages}쪽`;
    }
    set(el.stPages, pages);
    set(el.stSaved, d?.mtime ? `마지막 저장 ${whenText(d.mtime)}` : '', d?.path);
    if (d?.missing) set(el.stSaved, '파일을 찾을 수 없음');
  }

  // ── 미리보기 ─────────────────────────────────────────

  async function refreshPreview({ ifChanged = false, reason = 'manual' } = {}) {
    if (!state.doc) return;
    const path = state.doc.path;
    if (state.rendering && pathKey(state.rendering) === pathKey(path)) {
      // 이미 그리는 중이면 끝난 뒤 한 번 더
      const prev = state.pendingRefresh;
      state.pendingRefresh = { ifChanged: prev ? prev.ifChanged && ifChanged : ifChanged, reason: prev?.reason === 'manual' ? 'manual' : reason };
      return;
    }
    state.rendering = path;
    const seq = ++state.previewSeq;
    const had = !!state.preview;
    if (!had) showSkeleton();
    else if (!ifChanged || reason === 'manual') el.refreshing.hidden = false;
    setBusy('preview', had ? '미리보기 새로 고치는 중…' : '미리보기 만드는 중…');
    let result = null;
    try {
      const res = await call(K.doc.preview({ ifChanged }));
      if (seq !== state.previewSeq || !state.doc || pathKey(state.doc.path) !== pathKey(path)) return;
      if (res.doc) updateDoc(res.doc);
      if (res.unchanged) { result = { ok: true, kind: state.preview?.kind }; return; }
      state.preview = res;
      await renderPreview(res);
      renderTopbar();
      renderStatus();
      if (reason === 'watch') toast('문서가 바뀌어 새로 고쳤어요', { type: 'info' });
      result = { ok: true, kind: res.kind };
    } catch (err) {
      if (seq !== state.previewSeq) return;
      if (!had || reason === 'manual' || reason === 'open') showPreviewError(err);
      else showError(err, '미리보기를 새로 고치지 못했어요');
      result = { ok: false, kind: 'error', error: err.message };
    } finally {
      // 다른 문서를 그리는 중이면 그쪽 표시는 건드리지 않는다
      const mine = pathKey(state.rendering) === pathKey(path);
      if (mine) state.rendering = null;
      if (mine || seq === state.previewSeq) { el.refreshing.hidden = true; clearBusy('preview'); }
      if (result && reason === 'open') smokeSignal(result);
      const next = state.pendingRefresh;
      if (next && state.doc && pathKey(state.doc.path) === pathKey(path)) {
        state.pendingRefresh = null;
        refreshPreview(next);
      }
    }
  }

  function showSkeleton() {
    el.zoombar.hidden = true;
    el.preview.className = 'preview';
    const lines = ['w90', 'w75', 'w90', 'w60', 'gap', 'w90', 'w90', 'w40', 'gap', 'w75', 'w90', 'w60'];
    el.preview.replaceChildren(
      h('div', { class: 'skeleton', 'aria-hidden': 'true' },
        h('div', { class: 'sk-page' }, h('div', { class: 'sk-line title' }), ...lines.map((c) => h('div', { class: c === 'gap' ? 'sk-gap' : `sk-line ${c}` })))),
      h('div', { class: 'sk-label', role: 'status' }, spinner(), '미리보기를 만들고 있어요…'),
    );
  }

  function showPreviewError(err) {
    el.zoombar.hidden = true;
    el.preview.className = 'preview';
    const card = h('div', { class: 'error-card', role: 'alert' },
      h('div', { class: 'err-icon' }, icon('alert')),
      h('h3', { text: '미리보기를 만들지 못했어요' }),
      h('p', { text: err.message }),
      h('div', { class: 'btn-row' },
        h('button', { class: 'btn primary', onclick: () => refreshPreview({ reason: 'manual' }) }, icon('refresh'), '다시 시도'),
        err.code === 'HANCOM_SETUP' ? h('button', { class: 'btn', onclick: () => openSettings('hancom') }, '설정 열기') : null,
        h('button', { class: 'btn', onclick: () => actOpenInApp() }, icon('external'), '프로그램에서 열기')));
    el.preview.replaceChildren(h('div', { class: 'preview-error' }, card));
  }

  async function renderPreview(res) {
    if (res.kind === 'pdf') return renderPdf(res);
    if (res.kind === 'svg') return renderSvgPages(res);
    return renderText(res);
  }

  function renderPdf(res) {
    el.zoombar.hidden = true;
    el.preview.className = 'preview pdf';
    const frame = h('iframe', { class: 'pdf-frame', title: `${state.doc?.name ?? '문서'} 미리보기` });
    const loaded = new Promise((resolve) => {
      frame.addEventListener('load', () => resolve(), { once: true });
      setTimeout(resolve, 15_000);
    });
    frame.src = res.url;
    el.preview.replaceChildren(frame);
    return loaded;
  }

  function renderSvgPages(res) {
    el.preview.className = 'preview';
    const wrap = h('div', { class: 'pages' });
    const imgs = [];
    const scrollTop = el.preview.scrollTop;
    res.pages.forEach((b64, i) => {
      const img = h('img', { alt: `${i + 1}쪽`, draggable: 'false', decoding: 'async', loading: i < 3 ? 'eager' : 'lazy' });
      img.src = `data:image/svg+xml;base64,${b64}`;
      imgs.push(img);
      wrap.append(h('figure', { class: 'page' },
        h('div', { class: 'page-sheet' }, img),
        h('figcaption', { class: 'page-no', text: `${i + 1} / ${res.total}` })));
    });
    if (res.total > res.pages.length) {
      wrap.append(h('div', { class: 'pages-note', text: `앞 ${res.pages.length}쪽만 미리 보여요 (전체 ${res.total}쪽). 전체는 프로그램에서 열어 확인하세요.` }));
    }
    el.preview.replaceChildren(wrap);
    el.preview.scrollTop = scrollTop;
    el.zoombar.hidden = false;
    // 첫 쪽 크기를 알아야 맞춤 배율을 정할 수 있다
    const first = imgs[0];
    const ready = first ? (first.complete ? Promise.resolve() : new Promise((r) => { first.onload = r; first.onerror = r; })) : Promise.resolve();
    return ready.then(() => { applyZoom(); return Promise.all(imgs.slice(0, 3).map((im) => im.decode?.().catch(() => {}))); });
  }

  function svgImages() { return $$('.pages img', el.preview); }
  // 'auto' = 너비에 맞추되 100%를 넘지 않음(기본), 'fit' = 너비에 꽉 맞춤
  function fitZoom() {
    const imgs = svgImages();
    const w = Math.max(1, ...imgs.map((i) => i.naturalWidth || 0));
    const avail = el.viewer.clientWidth - 64 - 16; // 좌우 여백 + 세로 스크롤 막대
    return Math.max(0.25, Math.min(avail / w, 3));
  }
  function currentZoom() {
    if (state.zoom === 'fit') return fitZoom();
    if (state.zoom === 'auto') return Math.min(1, fitZoom());
    return state.zoom;
  }
  function applyZoom() {
    const z = currentZoom();
    for (const img of svgImages()) {
      if (!img.naturalWidth) { img.addEventListener('load', applyZoom, { once: true }); continue; }
      img.style.width = `${Math.round(img.naturalWidth * z)}px`;
      img.style.height = `${Math.round(img.naturalHeight * z)}px`;
    }
    $('#zoom-val').textContent = `${Math.round(z * 100)}%`;
    $('#zoom-fit').setAttribute('aria-pressed', String(state.zoom === 'fit'));
  }
  const ZOOMS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];
  function stepZoom(dir) {
    if (state.preview?.kind !== 'svg') return;
    const z = currentZoom();
    const next = dir > 0 ? ZOOMS.find((v) => v > z + 0.001) ?? ZOOMS.at(-1) : [...ZOOMS].reverse().find((v) => v < z - 0.001) ?? ZOOMS[0];
    // 화면 가운데를 기준으로 확대
    const p = el.preview;
    const cx = (p.scrollLeft + p.clientWidth / 2) / Math.max(1, p.scrollWidth);
    const cy = (p.scrollTop + p.clientHeight / 2) / Math.max(1, p.scrollHeight);
    state.zoom = next;
    applyZoom();
    p.scrollLeft = cx * p.scrollWidth - p.clientWidth / 2;
    p.scrollTop = cy * p.scrollHeight - p.clientHeight / 2;
  }

  // 글자 미리보기 (프로그램이 없을 때)
  function renderText(res) {
    el.zoombar.hidden = true;
    el.preview.className = 'preview';
    const wrap = h('div', { class: 'text-wrap' });
    if (res.sheets?.length) wrap.append(sheetView(res.sheets));
    else if (res.slides?.length) wrap.append(slideView(res.slides));
    else wrap.append(paperView(res.text ?? '', res.tables ?? []));
    el.preview.replaceChildren(wrap);
  }

  function paperView(text, tables) {
    const paper = h('article', { class: 'paper' });
    const lines = String(text).split('\n');
    let tableRows = [];
    const tableNode = (rows) => h('table', { class: 'doc-table' },
      h('tbody', {}, ...rows.map((r) => h('tr', {}, ...r.map((c) => h('td', { text: c }))))));
    const flushTable = () => {
      if (!tableRows.length) return;
      paper.append(tableNode(tableRows));
      tableRows = [];
    };
    for (const line of lines) {
      // 파일 엔진은 표를 "a | b" 줄로 풀어 준다
      if (line.includes(' | ')) { tableRows.push(line.split(' | ')); continue; }
      flushTable();
      paper.append(line.trim() ? h('p', { text: line }) : h('p', { class: 'blank' }));
    }
    flushTable();
    if (!text.trim()) for (const t of tables) paper.append(tableNode(t));
    if (!text.trim() && !tables.length) paper.append(h('p', { class: 'note', text: '문서에 글자가 없어요.' }));
    paper.append(h('p', { class: 'note', text: state.settings.previewMode === 'text'
      ? '서식 없이 글자만 보여 주는 미리보기예요. 설정에서 미리보기를 “빠르게”로 바꾸면 원본 모양으로 볼 수 있어요.'
      : '서식 없이 글자만 보여 주는 미리보기예요. 문서 프로그램이 설치돼 있으면 원본 모양으로 보여 드려요.' }));
    return paper;
  }

  function slideView(slides) {
    return h('div', { class: 'slides' }, ...slides.map((s, i) => {
      const [title, ...rest] = String(s).split('\n').filter((l) => l.trim());
      return h('div', { class: 'slide' },
        h('div', { class: 's-title', text: title ?? '' }),
        ...rest.slice(0, 12).map((l) => h('div', { class: 's-line', text: l })),
        h('span', { class: 's-no', text: String(i + 1) }));
    }));
  }

  const colName = (n) => { let s = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  function sheetView(sheets) {
    const view = h('div', { class: 'sheet-view' });
    const tabs = h('div', { class: 'sheet-tabs', role: 'tablist' });
    const scroll = h('div', { class: 'sheet-scroll' });
    const show = (idx) => {
      $$('.sheet-tab', tabs).forEach((t, i) => { t.classList.toggle('active', i === idx); t.setAttribute('aria-selected', String(i === idx)); });
      const rows = sheets[idx].rows ?? [];
      const cols = Math.max(1, ...rows.map((r) => r.length));
      const table = h('table', { class: 'grid' },
        h('thead', {}, h('tr', {}, h('th', {}), ...Array.from({ length: cols }, (_, c) => h('th', { text: colName(c) })))),
        h('tbody', {}, ...rows.map((r, ri) => h('tr', {}, h('td', { class: 'rh', text: String(ri + 1) }),
          ...Array.from({ length: cols }, (_, c) => {
            const v = r[c] ?? '';
            return h('td', { text: v, title: v.length > 30 ? v : null, class: /^-?[\d,.]+%?$/.test(v) ? 'num' : null });
          })))));
      scroll.replaceChildren(table);
    };
    sheets.forEach((s, i) => tabs.append(h('button', { class: 'sheet-tab', role: 'tab', text: s.name || `시트 ${i + 1}`, onclick: () => show(i) })));
    view.append(tabs, scroll);
    show(0);
    return view;
  }

  // ── 문서 작업 ────────────────────────────────────────

  async function actRestore() {
    const d = state.doc;
    if (!d?.backup || state.restoring) return;
    // AI가 고치는 중이면 먼저 멈춘다 (main이 멈추고 끝날 때까지 기다린 뒤 되돌린다)
    const aiBusy = !!chatFor(d.path).busy;
    const yes = await confirmDialog({
      title: '원래대로 되돌릴까요?',
      body: [
        aiBusy ? '진행 중인 AI 작업을 멈추고 되돌립니다.' : null,
        `${whenText(d.backup.time)}에 문서를 열면서 만든 백업으로 되돌립니다.`,
        '그 뒤에 바뀐 내용은 문서에서 사라지지만, 지금 상태도 백업 폴더에 따로 남겨 둡니다.',
      ].filter(Boolean),
      ok: aiBusy ? '멈추고 되돌리기' : '원래대로', tone: 'danger',
    });
    if (!yes || !state.doc || pathKey(state.doc.path) !== pathKey(d.path)) return;
    const chat = chatFor(d.path);
    if (chat.busy) markStopping(chat);
    state.restoring = true;
    renderTopbar();
    setBusy('restore', chat.busy ? 'AI 작업을 멈추고 되돌리는 중…' : '원래대로 되돌리는 중…');
    try {
      const r = await call(K.doc.restore());
      if (r.doc) updateDoc(r.doc);
      toast(r.stoppedAi > 0 ? 'AI 작업을 멈추고 원래대로 되돌렸어요' : '원래대로 되돌렸어요', { type: 'success' });
      if (state.doc && pathKey(state.doc.path) === pathKey(d.path)) await refreshPreview({ reason: 'manual' });
    } catch (err) { showError(err); } finally {
      state.restoring = false;
      clearBusy('restore');
      renderTopbar();
    }
  }

  async function actOpenInApp() {
    if (!state.doc) return;
    setBusy('openapp', '프로그램 여는 중…');
    try {
      await call(K.doc.openInApp());
      toast('프로그램에서 열었어요. 저장하면 여기 미리보기도 자동으로 바뀌어요.', { type: 'success' });
    } catch (err) { showError(err); } finally { clearBusy('openapp'); }
  }

  async function actSavePdf() {
    if (!state.doc) return;
    setBusy('pdf', 'PDF 만드는 중…');
    try {
      const r = await call(K.doc.savePdf());
      if (r?.canceled) return;
      toast(`PDF로 저장했어요: ${baseName(r.output)}`, { type: 'success', action: { label: '폴더에서 보기', run: () => call(K.shell.showItem(r.output)).catch(showError) } });
    } catch (err) { showError(err); } finally { clearBusy('pdf'); }
  }

  async function actShowFolder() {
    try { await call(K.doc.showInFolder()); } catch (err) { showError(err); }
  }

  async function actCloseDoc() {
    try { await call(K.doc.close()); } catch (err) { showError(err); }
  }

  // ── AI 대화 ──────────────────────────────────────────

  function chatFor(path) {
    const k = pathKey(path);
    let c = state.chats.get(k);
    if (!c) { c = { items: [], busy: false, chatId: null }; state.chats.set(k, c); }
    return c;
  }

  async function loadProviders(refresh = false) {
    try {
      state.providers = await call(K.ai.providers({ refresh }));
      state.providerError = null;
    } catch (err) {
      state.providers = null;
      state.providerError = err.message;
    }
    renderProviderSelect();
    renderChat();
  }

  function renderProviderSelect() {
    const sel = el.provider;
    const p = state.providers;
    sel.replaceChildren();
    if (!p || !p.providers.length) {
      sel.append(h('option', { value: '', text: state.providerError ? 'AI 없음' : '확인 중…' }));
      sel.disabled = true;
      return;
    }
    for (const it of p.providers) {
      sel.append(h('option', { value: it.id, text: it.available ? it.short : `${it.short} (설치 안 됨)`, title: it.name, disabled: !it.available }));
    }
    sel.value = p.selected ?? '';
    sel.disabled = !p.selected || chatCurrent()?.busy;
  }

  const chatCurrent = () => (state.doc ? chatFor(state.doc.path) : null);

  /** 대화창을 쓸 수 없는 이유 카드 (없으면 null) */
  function chatBlocker() {
    if (state.demo) return null;
    if (!state.doc) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('file')),
        h('h3', { text: '문서를 먼저 열어 주세요' }),
        h('p', { text: '문서를 열면 AI가 내용을 읽고 요약하거나, 글자를 바꾸고, 빈칸을 채워 드려요. 고친 결과는 왼쪽 미리보기에 바로 보여요.' }),
        h('button', { class: 'btn sm', onclick: pickFile }, icon('file-open'), '파일 열기'));
    }
    if (state.providerError) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('alert')),
        h('h3', { text: 'AI 기능을 쓸 수 없어요' }),
        h('p', { text: state.providerError }),
        h('button', { class: 'btn sm', onclick: () => loadProviders(true) }, icon('refresh'), '다시 확인'));
    }
    if (!state.providers) return h('div', { class: 'loading-line' }, spinner(), 'AI 프로그램을 확인하는 중…');
    if (!state.providers.selected) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('sparkle')),
        h('h3', { text: 'AI 프로그램이 필요해요' }),
        h('p', { text: 'Office Kit는 이 컴퓨터에 설치하고 로그인한 Claude Code 또는 Codex로 대화합니다. 아래 안내대로 하나를 설치한 뒤 다시 확인을 눌러 주세요.' }),
        h('div', { class: 'hint-list' }, ...state.providers.providers.map((p) =>
          h('div', { class: 'hint-item' }, h('b', { text: p.name }), h('span', { text: p.hint || '설치되어 있지 않아요.' })))),
        h('button', { class: 'btn sm', onclick: () => loadProviders(true) }, icon('refresh'), '다시 확인'));
    }
    return null;
  }

  const SUGGESTS = ['이 문서 요약해줘', '맞춤법 고쳐줘', '표 내용 정리해줘', 'PDF로 저장해줘'];

  /** 고른 AI가 로그인되지 않은 것 같으면 안내 (막지는 않음) */
  function loginWarning() {
    const p = state.providers?.providers.find((x) => x.id === state.providers.selected);
    if (!p || p.loggedIn !== false || !p.hint) return null;
    return h('div', { class: 'warn-note' }, icon('alert'), h('span', { text: p.hint }));
  }

  function renderChat() {
    const blocker = chatBlocker();
    const chat = chatCurrent();
    const log = el.chatLog;
    if (blocker) {
      log.replaceChildren(blocker);
    } else if (!chat.items.length) {
      log.replaceChildren(h('div', { class: 'chat-welcome' },
        h('div', { class: 'w-icon' }, icon('sparkle')),
        h('h3', { text: '무엇을 도와드릴까요?' }),
        h('p', { text: `「${state.doc.name}」 문서를 읽고 고쳐 드려요. 바꾼 내용은 왼쪽 미리보기에 바로 반영되고, 언제든 원래대로 되돌릴 수 있어요.` }),
        loginWarning(),
        h('div', { class: 'suggests' }, ...SUGGESTS.map((s) => h('button', { class: 'suggest', onclick: () => sendPrompt(s) }, icon('sparkle'), s)))));
    } else {
      log.replaceChildren(...chat.items.map(messageNode));
    }
    const off = !!blocker;
    el.composer.classList.toggle('disabled', off);
    el.prompt.disabled = off;
    el.prompt.placeholder = off ? (state.doc ? 'AI 프로그램을 먼저 준비해 주세요' : '문서를 열면 대화할 수 있어요') : '문서에 대해 물어보거나 고칠 내용을 적어 주세요';
    $('#btn-new-chat').disabled = off || !chat?.items.length || chat?.busy;
    renderSendButton();
    scrollChat(true);
  }

  function renderSendButton() {
    const chat = chatCurrent();
    const busy = !!chat?.busy;
    el.send.classList.toggle('stop', busy);
    el.send.replaceChildren(icon(busy ? 'stop' : 'send'));
    el.send.title = busy ? '멈추기' : '보내기 (Enter)';
    el.send.setAttribute('aria-label', busy ? '멈추기' : '보내기');
    el.send.disabled = !busy && (el.prompt.disabled || !el.prompt.value.trim());
    el.composerHint.textContent = busy ? 'AI가 작업 중이에요 · 멈추려면 ■ 버튼' : 'Enter 보내기 · Shift+Enter 줄바꿈';
    el.provider.disabled = busy || !state.providers?.selected;
  }

  function messageNode(item) {
    if (item.role === 'user') {
      return h('div', { class: 'msg user' }, h('div', { class: 'bubble', text: item.text }));
    }
    const body = h('div', { class: 'msg-body' });
    fillAssistant(body, item);
    item.node = body;
    return h('div', { class: 'msg assistant' }, h('div', { class: 'avatar', 'aria-hidden': 'true' }, icon('sparkle')), body);
  }

  function fillAssistant(body, item) {
    const nodes = [];
    let chips = null;
    for (const part of item.parts) {
      if (part.type === 'tool') {
        if (!chips) { chips = h('div', { class: 'tools' }); nodes.push(chips); }
        const st = part.state;
        const chip = h('span', { class: `chip ${st}`, title: part.summary || toolBase(part.tool || part.name) },
          st === 'run' ? spinner() : icon(st === 'ok' ? 'check' : st === 'fail' ? 'alert' : 'x'),
          h('span', { class: 'chip-label', text: toolLabel(part, st !== 'run') + (st === 'fail' ? ' 실패' : st === 'stop' ? ' (멈춤)' : '') }));
        chips.append(chip);
      } else if (part.text.trim()) {
        chips = null;
        nodes.push(h('div', { class: 'md', html: MD.render(part.text) }));
      }
    }
    // 마지막이 돌고 있는 도구면 칩이 이미 진행 중임을 보여 준다
    const lastPart = item.parts.at(-1);
    if (!item.done && !(lastPart?.type === 'tool' && lastPart.state === 'run' && !item.stopping)) {
      nodes.push(h('div', { class: 'thinking' }, h('span', { class: 'dots' }, h('i'), h('i'), h('i')), item.status || '생각하는 중…'));
    }
    if (item.error) nodes.push(h('div', { class: 'msg-error', role: 'alert', title: item.errorDetail || null }, icon('alert'), h('span', { text: item.error })));
    if (item.canceled && !item.error) nodes.push(h('div', { class: 'msg-note', text: '멈췄어요.' }));
    else if (item.done && item.stopping && !item.error) nodes.push(h('div', { class: 'msg-note', text: '멈추기 전에 작업이 끝났어요.' }));
    if (item.done && !item.error && !item.canceled && !item.parts.some((p) => p.type === 'text' && p.text.trim())) {
      nodes.push(h('div', { class: 'msg-note', text: '작업을 마쳤어요.' }));
    }
    body.replaceChildren(...nodes);
  }

  let chatFrame = 0;
  function updateAssistant(item) {
    if (!item.node || !item.node.isConnected) return;
    if (chatFrame) return;
    // rAF는 창이 가려지면 멈추므로 타이머로 모아서 그린다
    chatFrame = setTimeout(() => {
      chatFrame = 0;
      const nearBottom = isNearBottom();
      const chat = chatCurrent();
      if (chat) for (const it of chat.items) if (it.role === 'assistant' && it.dirty && it.node?.isConnected) { it.dirty = false; fillAssistant(it.node, it); }
      if (nearBottom) scrollChat(true);
    }, 33);
  }

  const isNearBottom = () => el.chatBody.scrollHeight - el.chatBody.scrollTop - el.chatBody.clientHeight < 80;
  function scrollChat(force) {
    if (force || isNearBottom()) el.chatBody.scrollTop = el.chatBody.scrollHeight;
  }

  async function sendPrompt(text) {
    text = String(text ?? '').trim();
    if (!text || !state.doc || chatBlocker()) return;
    const chat = chatFor(state.doc.path);
    if (chat.busy) return;
    const msg = { role: 'assistant', parts: [], done: false, status: '생각하는 중…', chatId: null };
    chat.items.push({ role: 'user', text }, msg);
    chat.busy = true;
    chat.pending = msg;
    chat.stopRequested = false;
    chat.cancelSent = null;
    el.prompt.value = '';
    autoGrow();
    renderChat();
    setBusy(`ai:${pathKey(state.doc.path)}`, 'AI가 작업 중…');
    try {
      const r = await call(K.ai.send({ prompt: text }));
      msg.chatId = r.chatId;
      if (chat.pending === msg) chat.pending = null;
      // 이미 끝났으면('end'가 먼저 옴) 대화 번호를 남기지 않는다
      if (!msg.done) chat.chatId = r.chatId;
      // 보내는 동안 멈추기를 눌렀으면 지금 멈춘다
      if (chat.stopRequested && !msg.done) cancelChatId(chat, r.chatId);
    } catch (err) {
      chat.pending = null;
      msg.done = true;
      msg.error = err.message;
      chat.busy = false;
      clearBusy(`ai:${pathKey(state.doc?.path)}`);
      if (err.code === 'NO_PROVIDER' || err.code === 'MODULE_MISSING') loadProviders(true);
      renderChat();
    }
  }

  function findMessage(chat, chatId) {
    for (let i = chat.items.length - 1; i >= 0; i--) {
      const it = chat.items[i];
      if (it.role === 'assistant' && it.chatId === chatId) return it;
    }
    // 응답보다 이벤트가 먼저 온 경우 (그 사이 멈추기를 눌렀으면 바로 멈춘다)
    if (chat.pending) {
      const m = chat.pending;
      m.chatId = chatId;
      chat.pending = null;
      if (chat.stopRequested) cancelChatId(chat, chatId);
      return m;
    }
    return null;
  }

  function onAiEvent(ev) {
    if (!ev || !ev.path) return;
    const chat = chatFor(ev.path);
    const msg = findMessage(chat, ev.chatId);
    const isCurrent = state.doc && pathKey(state.doc.path) === pathKey(ev.path);
    if (!msg) return;
    switch (ev.type) {
      case 'text': {
        const last = msg.parts.at(-1);
        const text = String(ev.text ?? '');
        if (ev.delta === false) {
          if (last?.type === 'text' && text.startsWith(last.text)) last.text = text;
          else msg.parts.push({ type: 'text', text });
        } else if (last?.type === 'text') last.text += text;
        else msg.parts.push({ type: 'text', text });
        msg.status = '답하는 중…';
        break;
      }
      case 'text_reset': {
        // 스트림이 끊겨 AI가 다시 쓰는 경우: 이미 보인 글 끝에서 drop 글자만큼 지운다
        let drop = Math.max(0, Number(ev.drop) || 0);
        for (let i = msg.parts.length - 1; i >= 0 && drop > 0; i--) {
          const p = msg.parts[i];
          if (p.type !== 'text') continue;
          const cut = Math.min(drop, p.text.length);
          p.text = p.text.slice(0, p.text.length - cut);
          drop -= cut;
          if (!p.text) msg.parts.splice(i, 1);
        }
        break;
      }
      case 'tool': {
        const part = { type: 'tool', id: ev.id, tool: String(ev.tool ?? ''), name: String(ev.name ?? ''), state: 'run', summary: '' };
        msg.parts.push(part);
        msg.status = `${toolLabel(part, false)}…`;
        break;
      }
      case 'tool_result': {
        const part = msg.parts.find((p) => p.type === 'tool' && p.id === ev.id && p.state === 'run')
          ?? msg.parts.find((p) => p.type === 'tool' && p.id === ev.id);
        if (part) {
          part.state = ev.ok === false ? 'fail' : 'ok';
          part.summary = String(ev.summary ?? '');
          if (part.state === 'ok' && EDIT_TOOLS.has(toolBase(part.tool || part.name)) && isCurrent) refreshPreview({ ifChanged: true, reason: 'ai' });
        }
        msg.status = '생각하는 중…';
        break;
      }
      case 'status':
        if (ev.message) msg.status = String(ev.message);
        break;
      case 'done':
        msg.done = true;
        if (ev.cancelled) msg.canceled = true;
        if (ev.text && !msg.parts.some((p) => p.type === 'text' && p.text.trim())) msg.parts.push({ type: 'text', text: String(ev.text) });
        break;
      case 'error':
        msg.error = String(ev.message ?? '오류가 났습니다');
        msg.errorDetail = ev.detail ? String(ev.detail).slice(-600) : '';
        msg.done = true;
        break;
      case 'end':
        msg.done = true;
        for (const p of msg.parts) if (p.type === 'tool' && p.state === 'run') p.state = msg.canceled ? 'stop' : 'ok';
        if (chat.chatId === ev.chatId) chat.chatId = null;
        chat.busy = false;
        clearBusy(`ai:${pathKey(ev.path)}`);
        if (isCurrent) {
          refreshPreview({ ifChanged: true, reason: 'ai' });
          renderChat();
          return;
        }
        break;
      default:
        return;
    }
    // 멈추기를 누른 뒤에는 다른 상태 글로 덮지 않는다
    if (msg.stopping && !msg.done) msg.status = '멈추는 중…';
    msg.dirty = true;
    if (isCurrent) updateAssistant(msg);
  }

  /** 멈추기 요청만 표시한다. '멈췄어요'는 AI가 실제로 멈췄다고 알려 올 때(done.cancelled) 붙인다 */
  function markStopping(chat) {
    chat.stopRequested = true;
    const msg = chat.items.findLast((i) => i.role === 'assistant');
    if (msg && !msg.done) { msg.stopping = true; msg.status = '멈추는 중…'; msg.dirty = true; updateAssistant(msg); }
  }

  /** 대화 하나에 멈추기를 한 번만 보낸다 */
  async function cancelChatId(chat, chatId) {
    if (!chatId || chat.cancelSent === chatId) return;
    chat.cancelSent = chatId;
    try { await call(K.ai.cancel(chatId)); } catch (err) { showError(err); }
  }

  async function stopChat() {
    const chat = chatCurrent();
    if (!chat?.busy) return;
    markStopping(chat);
    // 아직 대화 번호를 못 받았으면 받는 즉시 sendPrompt가 멈춘다
    if (chat.chatId) await cancelChatId(chat, chat.chatId);
  }

  async function newChat() {
    const chat = chatCurrent();
    if (!chat || chat.busy) return;
    try { await call(K.ai.reset()); } catch {}
    chat.items = [];
    renderChat();
    el.prompt.focus();
  }

  function autoGrow() {
    const t = el.prompt;
    t.style.height = 'auto';
    t.style.height = `${Math.min(t.scrollHeight, 180)}px`;
    renderSendButton();
  }

  // 확인용: 가짜 대화 내용
  function injectDemoChat() {
    if (!state.doc) return;
    state.demo = true;
    const chat = chatFor(state.doc.path);
    chat.items = [
      { role: 'user', text: '담당자 칸에 김민지 넣고, 표 숫자 단위도 맞춰줘' },
      { role: 'assistant', done: true, parts: [
        { type: 'tool', id: '1', tool: 'office_read', name: '문서 읽기', state: 'ok' },
        { type: 'tool', id: '2', tool: 'office_fill', name: '칸 채우기', state: 'ok' },
        { type: 'tool', id: '3', tool: 'office_replace', name: '찾아 바꾸기', state: 'ok' },
        { type: 'text', text: '**담당자** 칸에 `김민지`를 넣었어요.\n\n- 표의 금액을 모두 **만 원** 단위로 맞췄어요 (2곳)\n- 나머지 서식은 그대로 두었어요\n\n왼쪽 미리보기에서 바뀐 부분을 확인해 보세요.' },
      ] },
      { role: 'user', text: '결과 한번 확인해 줘' },
      { role: 'assistant', done: false, status: '결과 확인 중…', parts: [
        { type: 'tool', id: '4', tool: 'office_page_image', name: '쪽 이미지로 확인', state: 'run' },
      ] },
    ];
    chat.busy = true;
    renderChat();
    renderProviderSelect();
  }

  // ── 설정 창 ──────────────────────────────────────────

  function openSettings(tab = 'ai') {
    if (!el.settings.open) el.settings.showModal();
    renderSettingsView();
    // 다른 곳에서 바뀌었을 수 있으니 저장된 설정을 다시 읽어 맞춘다
    call(K.settings.get()).then((s) => { state.settings = { ...state.settings, ...s }; renderSettingsView(); }).catch(() => {});
    const loading = selectTab(tab);
    $(`.nav-item[data-tab="${tab}"]`, el.settings)?.focus();
    return loading;
  }

  function selectTab(tab) {
    $$('.nav-item', el.settings).forEach((b) => b.classList.toggle('active', b.dataset.tab === tab));
    $$('.settings-page', el.settings).forEach((p) => { p.hidden = p.dataset.page !== tab; });
    if (tab === 'ai') return renderProviderChoices(true);
    if (tab === 'connect') return loadIntegrations();
    if (tab === 'hancom') return loadHancom();
    return Promise.resolve();
  }

  function renderSettingsView() {
    const s = state.settings;
    $$('input[name="previewMode"]', el.settings).forEach((r) => { r.checked = r.value === (s.previewMode ?? 'auto'); });
    $$('#theme-seg button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.theme === (s.theme ?? 'system'))));
    renderModelInputs();
    $('#app-version').textContent = state.version ? `Office Kit ${state.version}` : '';
  }

  async function saveSettings(patch) {
    try {
      state.settings = await call(K.settings.set(patch));
      renderSettingsView();
      return true;
    } catch (err) { showError(err); return false; }
  }

  async function renderProviderChoices(reload) {
    const box = $('#provider-choices');
    if (reload || !state.providers) {
      if (!state.providers) box.replaceChildren(h('div', { class: 'loading-line' }, spinner(), '확인하는 중…'));
      await loadProviders(!!reload && !!state.providers);
    }
    if (state.providerError) { box.replaceChildren(h('div', { class: 'inline-error', text: state.providerError })); return; }
    const p = state.providers;
    if (!p) return;
    box.replaceChildren(...p.providers.map((it) => {
      const input = h('input', { type: 'radio', name: 'provider', value: it.id, disabled: !it.available });
      input.checked = it.id === p.selected;
      input.addEventListener('change', async () => {
        if (await saveSettings({ provider: it.id })) { await loadProviders(false); renderProviderChoices(false); }
      });
      return h('label', { class: `choice${it.available ? '' : ' disabled'}` }, input,
        h('span', { class: 'choice-body' },
          h('span', { class: 'choice-title' }, it.name,
            h('span', { class: `pill ${it.available ? 'on' : 'off'}`, text: it.available ? (it.version ? `설치됨 · ${it.version}` : '설치됨') : '설치 안 됨' })),
          h('span', { class: 'choice-desc', text: it.available ? (it.id === 'claude' ? 'Anthropic의 Claude Code 명령줄 도구로 대화합니다.' : it.id === 'codex' ? 'OpenAI의 Codex 명령줄 도구로 대화합니다.' : (it.hint || '')) : (it.hint || '설치되어 있지 않아요.') })));
    }));
    if (!p.providers.length) box.replaceChildren(h('div', { class: 'help', text: '찾은 AI 프로그램이 없어요.' }));
    renderModelInputs();
  }

  // AI 목록을 못 불러와도 모델은 정할 수 있게 알려진 것을 기본으로 둔다
  const KNOWN_PROVIDERS = [{ id: 'claude', short: 'Claude Code' }, { id: 'codex', short: 'Codex' }];
  const MODEL_EXAMPLE = { claude: '예: sonnet, opus', codex: '예: gpt-5' };

  /** AI마다 모델 입력 칸 하나씩 (다른 AI의 모델 이름이 넘어가지 않게) */
  function renderModelInputs() {
    const box = $('#model-list');
    const models = state.settings.models ?? {};
    const list = state.providers?.providers?.length ? state.providers.providers : KNOWN_PROVIDERS;
    const labelText = (p) => [p.short ?? p.name ?? p.id, state.providers?.selected === p.id ? h('span', { class: 'opt', text: '사용 중' }) : null];
    const existing = new Map($$('input[data-provider]', box).map((i) => [i.dataset.provider, i]));
    // 같은 목록이면 입력 중인 칸을 다시 만들지 않고 값과 이름표만 맞춘다
    if (existing.size === list.length && list.every((p) => existing.has(p.id))) {
      for (const p of list) {
        const input = existing.get(p.id);
        if (document.activeElement !== input) input.value = models[p.id] ?? '';
        $(`label[for="model-${p.id}"]`, box)?.replaceChildren(...labelText(p).filter(Boolean));
      }
      return;
    }
    box.replaceChildren(...list.flatMap((p) => {
      const id = `model-${p.id}`;
      const input = h('input', {
        class: 'input', id, type: 'text', maxlength: '100', spellcheck: 'false', autocomplete: 'off',
        placeholder: `기본 모델${MODEL_EXAMPLE[p.id] ? ` (${MODEL_EXAMPLE[p.id]})` : ''}`, dataset: { provider: p.id },
      });
      input.value = models[p.id] ?? '';
      input.addEventListener('change', () => saveSettings({ models: { [p.id]: input.value.trim() } }));
      return [h('label', { class: 'model-name', for: id }, ...labelText(p)), input];
    }));
  }

  const INTEGRATIONS = [
    { key: 'claudeDesktop', name: 'Claude Desktop', desc: 'Claude 데스크톱 앱에서 Office Kit 문서 도구를 씁니다.', icon: 'sparkle' },
    { key: 'claudeCode', name: 'Claude Code', desc: '터미널의 Claude Code에서 Office Kit 문서 도구를 씁니다.', icon: 'terminal' },
    { key: 'codex', name: 'Codex', desc: 'Codex CLI에서 Office Kit 문서 도구를 씁니다.', icon: 'terminal' },
    { key: 'cli', name: '명령줄 도구', desc: '어느 폴더에서나 office-kit 명령을 쓸 수 있게 등록합니다.', icon: 'keyboard' },
  ];

  async function loadIntegrations() {
    const list = $('#int-list');
    list.replaceChildren(...INTEGRATIONS.map((it) => integrationRow(it, null)));
    try {
      const st = await call(K.integrations.status());
      list.replaceChildren(...INTEGRATIONS.map((it) => integrationRow(it, st[it.key])));
    } catch (err) {
      list.replaceChildren(h('div', { class: 'int-row' }, h('div', { class: 'inline-error', text: err.message })));
    }
  }

  function integrationRow(it, st) {
    const pill = !st ? h('span', { class: 'pill wait', text: '확인 중' })
      : st.error ? h('span', { class: 'pill err', text: '확인 필요' })
        : st.connected && !st.matches ? h('span', { class: 'pill warn', text: '다시 연결 필요' })
          : st.connected ? h('span', { class: 'pill on', text: '연결됨' })
            : !st.available ? h('span', { class: 'pill off', text: '설치 안 됨' })
              : h('span', { class: 'pill off', text: '연결 안 됨' });
    const reconnect = st?.connected && !st.matches;
    const btn = h('button', { class: `btn sm${st && (!st.connected || reconnect) ? ' primary' : ''}`, disabled: !st || (!st.connected && !st.available), text: reconnect ? '다시 연결' : st?.connected ? '연결 해제' : '연결' });
    btn.addEventListener('click', async () => {
      const on = !st.connected || reconnect;
      btn.disabled = true;
      btn.replaceChildren(spinner(), on ? '연결 중' : '해제 중');
      try {
        const r = await call(K.integrations.set({ target: it.key, on }));
        toast(on ? `${it.name}에 연결했어요.${it.key === 'cli' ? ' 새 터미널에서 office-kit 명령을 쓸 수 있어요.' : ` ${it.name}을(를) 다시 시작하면 적용돼요.`}` : `${it.name} 연결을 해제했어요.`, { type: 'success' });
        const list = $('#int-list');
        list.replaceChildren(...INTEGRATIONS.map((x) => integrationRow(x, r.status[x.key])));
      } catch (err) {
        showError(err);
        loadIntegrations();
      }
    });
    return h('div', { class: 'int-row' },
      h('span', { class: 'int-icon' }, icon(it.icon)),
      h('div', { class: 'int-text' },
        h('div', { class: 'int-name' }, it.name, pill),
        h('div', { class: 'int-desc', text: st?.error || (st?.detail && st.connected ? st.detail : it.desc), title: st?.detail || null })),
      btn);
  }

  async function loadHancom() {
    const grid = $('#hancom-status');
    const note = $('#hancom-note');
    const removeBtn = $('#btn-hancom-remove');
    grid.replaceChildren(h('div', { class: 'loading-line' }, spinner(), '확인하는 중…'));
    note.textContent = '';
    removeBtn.hidden = true;
    $('#btn-hancom-setup').disabled = true;
    $('#btn-hancom-dll').disabled = true;
    try {
      const r = await call(K.hancom.status());
      const s = r.status;
      state.hancomSource = s?.source ?? null;
      removeBtn.hidden = !(s?.registered && s.canRemove);
      const installed = s ? s.hancomInstalled : r.apps.hancom;
      const registered = s ? s.registered : r.apps.hancomSecurityModule;
      const row = (k, v) => [h('span', { class: 'k', text: k }), h('span', { class: 'v' }, v)];
      grid.replaceChildren(
        ...row('한컴오피스', h('span', { class: `pill ${installed ? 'on' : 'off'}`, text: installed ? '설치됨' : '설치 안 됨' })),
        ...row('자동화 보안 모듈', h('span', { class: `pill ${registered ? 'on' : 'warn'}`, text: registered ? '등록됨' : '등록 안 됨' })),
        ...(s?.dllPath ? row('모듈 파일', h('span', { class: 'path', title: s.dllPath, text: s.dllExists ? s.dllPath : `${s.dllPath} (파일 없음)` })) : []),
        ...row('다른 프로그램', h('span', { text: ['Word', 'PowerPoint', 'Excel'].filter((n) => r.apps[n.toLowerCase()]).join(' · ') || '없음' })));
      $('#btn-hancom-setup').disabled = !s || !installed;
      $('#btn-hancom-dll').disabled = !s || !installed;
      $('#btn-hancom-setup').lastChild.textContent = registered ? '다시 연결' : '한글 자동화 연결';
      if (r.error) note.textContent = r.error;
      else if (s?.warning) note.textContent = s.warning;
      else if (!installed) note.textContent = '한컴오피스가 설치돼 있지 않아요. 한글 문서는 내장 엔진으로 계속 열고 고칠 수 있어요.';
      else if (registered) note.textContent = '연결돼 있어요. 미리보기를 “정확하게”로 바꾸면 한글 프로그램으로 그린 결과를 볼 수 있어요.';
    } catch (err) {
      grid.replaceChildren(h('div', { class: 'inline-error', text: err.message }));
    }
  }

  function sourceText(src) {
    try {
      const u = new URL(src?.url ?? '');
      const parts = u.pathname.split('/').filter(Boolean);
      return u.hostname === 'github.com' ? `github.com/${parts.slice(0, 2).join('/')}` : u.hostname;
    } catch { return ''; }
  }

  async function hancomRemove() {
    const yes = await confirmDialog({
      title: '한글 자동화 연결을 해제할까요?',
      body: ['레지스트리에 등록한 보안 모듈 항목만 지웁니다. 내려받은 파일은 그대로 둡니다.', '해제하면 한글 문서는 다시 내장 엔진으로 다룹니다.'],
      ok: '연결 해제', tone: 'danger',
    });
    if (!yes) return;
    try {
      await call(K.hancom.remove({ consent: true }));
      toast('한글 자동화 연결을 해제했어요', { type: 'success' });
    } catch (err) { showError(err); } finally { loadHancom(); }
  }

  async function hancomSetup(mode) {
    const download = mode === 'download';
    const where = sourceText(state.hancomSource);
    let page = '';
    try { page = state.hancomSource?.page ? new URL(state.hancomSource.page).hostname : ''; } catch {}
    const yes = await confirmDialog({
      title: download ? '한글 자동화를 연결할까요?' : '선택한 모듈을 등록할까요?',
      tone: 'shield',
      ok: download ? '동의하고 연결' : '파일 선택',
      body: download ? [
        '한컴오피스가 Office Kit의 자동 작업을 허용하도록 다음을 진행합니다.',
        h('ul', {},
          h('li', {}, '한컴 공식 배포본에서 한컴이 제공하는 “자동화 보안 승인 모듈”(DLL)을 내려받습니다.',
            where ? h('br') : null, where ? h('code', { text: where }) : null, page ? ` (한컴 개발자센터 ${page} 안내)` : null),
          h('li', { text: '받은 파일은 이 PC의 Office Kit 데이터 폴더에 저장합니다.' }),
          h('li', {}, '현재 사용자 레지스트리(', h('code', { text: 'HKCU\\Software\\HNC\\HwpAutomation\\Modules' }), ')에 등록합니다. 관리자 권한은 필요 없어요.')),
        '등록하면 한글 프로그램으로 PDF 저장과 정확한 미리보기를 할 수 있어요. 문서 내용은 어디로도 보내지 않습니다.',
      ] : [
        '가지고 있는 한컴 자동화 보안 승인 모듈(DLL)을 직접 고릅니다.',
        h('p', {}, '고른 파일을 현재 사용자 레지스트리(', h('code', { text: 'HKCU\\Software\\HNC\\HwpAutomation\\Modules' }), ')에 등록합니다. 한컴에서 받은 파일만 선택하세요.'),
      ],
    });
    if (!yes) return;
    const btn = $(download ? '#btn-hancom-setup' : '#btn-hancom-dll');
    btn.disabled = true;
    setBusy('hancom', download ? '보안 모듈 내려받는 중…' : '보안 모듈 등록 중…');
    try {
      const r = await call(K.hancom.setup({ consent: true, mode: download ? 'download' : 'pick' }));
      if (r?.canceled) return;
      toast('한글 자동화를 연결했어요', { type: 'success' });
    } catch (err) { showError(err, '한글 자동화를 연결하지 못했어요'); } finally {
      clearBusy('hancom');
      btn.disabled = false;
      loadHancom();
    }
  }

  // ── 대화창 너비 · 접기 ────────────────────────────────

  function applyChatLayout() {
    const s = state.settings;
    document.documentElement.style.setProperty('--chat-w', `${s.chatWidth ?? 400}px`);
    el.app.classList.toggle('chat-collapsed', !!s.chatCollapsed);
    el.toggleChat.setAttribute('aria-pressed', String(!s.chatCollapsed));
    if (state.preview?.kind === 'svg' && typeof state.zoom === 'string') requestAnimationFrame(applyZoom);
  }

  function clampChat(w) {
    const max = Math.min(760, el.workspace.clientWidth - 460);
    return Math.round(Math.max(300, Math.min(max, w)));
  }

  function setupSplitter() {
    const sp = el.splitter;
    let startX = 0; let startW = 0; let dragging = false;
    sp.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      dragging = true; startX = e.clientX; startW = el.chat.getBoundingClientRect().width;
      sp.setPointerCapture(e.pointerId);
      document.body.classList.add('resizing');
      e.preventDefault();
    });
    sp.addEventListener('pointermove', (e) => {
      if (!dragging) return;
      state.settings.chatWidth = clampChat(startW + (startX - e.clientX));
      document.documentElement.style.setProperty('--chat-w', `${state.settings.chatWidth}px`);
      if (state.preview?.kind === 'svg' && typeof state.zoom === 'string') applyZoom();
    });
    const end = (e) => {
      if (!dragging) return;
      dragging = false;
      try { sp.releasePointerCapture(e.pointerId); } catch {}
      document.body.classList.remove('resizing');
      saveSettings({ chatWidth: state.settings.chatWidth });
    };
    sp.addEventListener('pointerup', end);
    sp.addEventListener('pointercancel', end);
    sp.addEventListener('dblclick', () => { state.settings.chatWidth = 400; applyChatLayout(); saveSettings({ chatWidth: 400 }); });
    sp.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      e.preventDefault();
      state.settings.chatWidth = clampChat((state.settings.chatWidth ?? 400) + (e.key === 'ArrowLeft' ? 24 : -24));
      applyChatLayout();
      saveSettings({ chatWidth: state.settings.chatWidth });
    });
  }

  function toggleChat(force) {
    const collapsed = typeof force === 'boolean' ? force : !state.settings.chatCollapsed;
    state.settings.chatCollapsed = collapsed;
    applyChatLayout();
    saveSettings({ chatCollapsed: collapsed });
    if (!collapsed) setTimeout(() => el.prompt.focus(), 0);
  }

  // ── 끌어놓기 ─────────────────────────────────────────

  function setupDragDrop() {
    let depth = 0;
    let hideTimer = 0;
    const hasFiles = (e) => [...(e.dataTransfer?.types ?? [])].includes('Files');
    const show = (on) => {
      el.overlay.hidden = !on;
      document.body.classList.toggle('dragging', on);
      if (!on) depth = 0;
    };
    // dragleave가 빠지는 경우를 대비해 dragover가 끊기면 닫는다
    const keepAlive = () => { clearTimeout(hideTimer); hideTimer = setTimeout(() => show(false), 400); };
    window.addEventListener('dragenter', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth++;
      show(true);
      keepAlive();
    });
    window.addEventListener('dragover', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      keepAlive();
    });
    window.addEventListener('dragleave', (e) => {
      if (!hasFiles(e)) return;
      depth = Math.max(0, depth - 1);
      if (depth === 0) show(false);
    });
    window.addEventListener('drop', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      clearTimeout(hideTimer);
      show(false);
      const files = [...(e.dataTransfer.files ?? [])];
      if (!files.length) return;
      const paths = files.map((f) => K.pathForFile(f)).filter(Boolean);
      const ok = paths.find((p) => state.exts.includes(extOf(p)));
      if (!ok) {
        const ext = extOf(paths[0] ?? files[0].name);
        toast(`열 수 없는 형식이에요${ext ? ` (.${ext})` : ''}. 한글·Word·PowerPoint·Excel 문서를 놓아 주세요.`, { type: 'warn' });
        return;
      }
      if (paths.length > 1) toast('여러 파일 중 첫 번째 문서만 열었어요', { type: 'info' });
      openPath(ok);
    });
  }

  // ── 확인용 캡처 신호 ─────────────────────────────────

  async function smokeSignal(info) {
    if (!state.smoke || state.smokeSent) return;
    state.smokeSent = true;
    if (state.smoke.chat) injectDemoChat();
    if (state.smoke.settings) {
      try { await openSettings(state.smoke.settings); } catch {}
      await sleep(300);
    }
    setTimeout(() => K.smokeReady(info), 50);
  }

  // ── 연결 · 시작 ──────────────────────────────────────

  function bindUi() {
    $('#btn-open').addEventListener('click', pickFile);
    $('#dropzone').addEventListener('click', pickFile);
    $('#btn-refresh').addEventListener('click', () => refreshPreview({ reason: 'manual' }));
    $('#btn-restore').addEventListener('click', actRestore);
    $('#btn-openapp').addEventListener('click', actOpenInApp);
    $('#btn-pdf').addEventListener('click', actSavePdf);
    $('#btn-folder').addEventListener('click', actShowFolder);
    $('#btn-close-doc').addEventListener('click', actCloseDoc);
    $('#btn-settings').addEventListener('click', () => openSettings('ai'));
    $('#btn-toggle-chat').addEventListener('click', () => toggleChat());
    $('#btn-collapse-chat').addEventListener('click', () => toggleChat(true));
    $('#btn-new-chat').addEventListener('click', newChat);
    $('#btn-recent-clear').addEventListener('click', async () => {
      const yes = await confirmDialog({ title: '최근 문서 목록을 비울까요?', body: '목록만 지우고 파일은 그대로 둡니다.', ok: '비우기' });
      if (yes) { try { renderRecent(await call(K.recent.clear())); } catch (err) { showError(err); } }
    });

    // 확대/축소
    $('#zoom-in').addEventListener('click', () => stepZoom(1));
    $('#zoom-out').addEventListener('click', () => stepZoom(-1));
    $('#zoom-val').addEventListener('click', () => { state.zoom = 1; applyZoom(); });
    $('#zoom-fit').addEventListener('click', () => { state.zoom = state.zoom === 'fit' ? 'auto' : 'fit'; applyZoom(); });
    el.preview.addEventListener('wheel', (e) => {
      if (!e.ctrlKey || state.preview?.kind !== 'svg') return;
      e.preventDefault();
      stepZoom(e.deltaY < 0 ? 1 : -1);
    }, { passive: false });
    new ResizeObserver(() => { if (state.preview?.kind === 'svg' && typeof state.zoom === 'string') applyZoom(); }).observe(el.viewer);

    // 대화 입력
    el.prompt.addEventListener('input', autoGrow);
    el.prompt.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && e.keyCode !== 229) {
        e.preventDefault();
        if (!chatCurrent()?.busy) sendPrompt(el.prompt.value);
      }
    });
    el.send.addEventListener('click', () => (chatCurrent()?.busy ? stopChat() : sendPrompt(el.prompt.value)));
    el.provider.addEventListener('change', async () => {
      if (await saveSettings({ provider: el.provider.value || null })) loadProviders(false);
    });
    // 답변 속 링크는 기본 브라우저로
    el.chatLog.addEventListener('click', (e) => {
      const a = e.target.closest('a.ext');
      if (!a) return;
      e.preventDefault();
      call(K.shell.openExternal(a.dataset.href)).catch(showError);
    });

    // 설정 창
    $$('.nav-item', el.settings).forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.tab)));
    $('#settings-close').addEventListener('click', () => el.settings.close());
    el.settings.addEventListener('click', (e) => { if (e.target === el.settings) el.settings.close(); });
    $$('input[name="previewMode"]', el.settings).forEach((r) => r.addEventListener('change', async () => {
      if (!r.checked) return;
      if (await saveSettings({ previewMode: r.value }) && state.doc) {
        toast({ exact: '설치된 프로그램으로 미리보기를 다시 만들어요', text: '글자만 보는 미리보기로 바꿨어요' }[r.value] ?? '빠른 미리보기로 다시 만들어요', { type: 'info' });
        refreshPreview({ reason: 'manual' });
      }
    }));
    $$('#theme-seg button').forEach((b) => b.addEventListener('click', () => saveSettings({ theme: b.dataset.theme })));
    $('#btn-provider-recheck').addEventListener('click', () => renderProviderChoices(true));
    $('#btn-open-backups').addEventListener('click', () => call(K.shell.openBackups()).catch(showError));
    $('#btn-hancom-setup').addEventListener('click', () => hancomSetup('download'));
    $('#btn-hancom-dll').addEventListener('click', () => hancomSetup('pick'));
    $('#btn-hancom-remove').addEventListener('click', hancomRemove);

    // 단축키 (메뉴에 없는 것)
    window.addEventListener('keydown', (e) => {
      if (!e.ctrlKey || e.altKey) return;
      if (e.key === '=' || e.key === '+') { e.preventDefault(); stepZoom(1); }
      else if (e.key === '-') { e.preventDefault(); stepZoom(-1); }
      else if (e.key === '0') { e.preventDefault(); state.zoom = 'auto'; applyZoom(); }
    });

    setupSplitter();
    setupDragDrop();
  }

  function bindEvents() {
    K.on.docOpened(onDocOpened);
    K.on.docChanged((ev) => {
      if (!state.doc || pathKey(ev.path) !== pathKey(state.doc.path)) return;
      if (ev.missing) {
        updateDoc({ path: ev.path, missing: true });
        toast('문서를 찾을 수 없어요. 이름이 바뀌었거나 지워졌을 수 있어요.', { type: 'warn' });
        return;
      }
      updateDoc({ path: ev.path, missing: false, mtime: ev.mtime, size: ev.size });
      // AI가 고치는 중이면 대화창이 이미 알려 주므로 알림 없이 새로 고친다
      refreshPreview({ ifChanged: true, reason: chatCurrent()?.busy ? 'ai' : 'watch' });
    });
    K.on.aiEvent(onAiEvent);
    K.on.menu((cmd) => {
      if (cmd === 'open') pickFile();
      else if (cmd === 'refresh') { if (state.doc) refreshPreview({ reason: 'manual' }); }
      else if (cmd === 'settings') openSettings('ai');
    });
    K.on.toast((t) => toast(String(t?.message ?? ''), { type: t?.type ?? 'info' }));
  }

  async function start() {
    bindUi();
    bindEvents();
    let init;
    try { init = await call(K.init()); } catch (err) {
      toast(`시작하지 못했어요: ${err.message}`, { type: 'error' });
      return;
    }
    state.exts = init.exts ?? [];
    state.settings = init.settings ?? {};
    state.version = init.version ?? '';
    state.smoke = init.smoke;
    applyChatLayout();
    renderRecent(init.recent);
    if (init.doc) onDocOpened(init.doc);
    else { renderTopbar(); renderStatus(); renderChat(); }
    loadProviders(false);
    // 열 파일 없이 시작한 확인용 캡처는 빈 화면을 찍는다
    if (state.smoke && !init.doc && !init.opening) setTimeout(() => smokeSignal({ ok: true, kind: 'empty' }), 600);
  }

  start();
})();
