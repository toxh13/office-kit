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
    word_edit: ['Word 편집 중', 'Word 편집'],
    ppt_edit: ['슬라이드 편집 중', '슬라이드 편집'],
    excel_edit: ['시트 편집 중', '시트 편집'],
    hwp_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    word_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    ppt_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    excel_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
  };
  const EDIT_TOOLS = new Set(['office_fill', 'office_replace', 'office_set_cells', 'hwp_edit', 'office_create', 'office_convert',
    'word_edit', 'ppt_edit', 'excel_edit']);
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
    providersReady: null,   // 처음 AI 확인이 끝나면 풀리는 Promise
    chats: new Map(),       // 문서 경로 → { items, busy, chatId }
    busy: new Map(),
    smoke: null,
    smokeSent: false,
    demo: false,
    packaged: true,
    // 편집기 묶음 (init에서 받음): 종류마다 { exts, available, reason }. hwp = 한글 편집, office = 문서·표·슬라이드 편집
    editor: {
      exts: ['hwp', 'hwpx', 'hwt', 'docx', 'xlsx', 'pptx'],
      kinds: {
        hwp: { exts: ['hwp', 'hwpx', 'hwt'], available: false, reason: null },
        office: { exts: ['docx', 'xlsx', 'pptx'], available: false, reason: null },
      },
    },
  };

  const el = {
    app: $('#app'), workspace: $('#workspace'), viewer: $('#viewer'),
    empty: $('#empty'), preview: $('#preview'), zoombar: $('#zoombar'), refreshing: $('#refreshing'),
    recent: $('#recent'), recentList: $('#recent-list'),
    docinfo: $('#docinfo'), docKind: $('#doc-kind'), docName: $('#doc-name'), docEngine: $('#doc-engine'),
    chat: $('#chat'), chatBody: $('#chat-body'), chatLog: $('#chat-log'), provider: $('#provider'),
    composer: $('#composer'), prompt: $('#prompt'), send: $('#btn-send'), composerHint: $('#composer-hint'),
    splitter: $('#splitter'), toggleChat: $('#btn-toggle-chat'),
    stEngine: $('#st-engine'), stPages: $('#st-pages'), stSaved: $('#st-saved'), stEditor: $('#st-editor'),
    stBusy: $('#st-busy'), stBusyText: $('#st-busy-text'), stReady: $('#st-ready'),
    btnEdit: $('#btn-edit'), editorWarn: $('#editor-warn'),
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

  /**
   * body: 문자열 | 노드 배열. alt: 세 번째 단추 이름 (누르면 'alt'), cancel: null이면 취소 단추 없음
   * → Promise<true | false | 'alt'>
   */
  function confirmDialog({ title, body, ok = '확인', cancel = '취소', alt = null, tone = 'info' }) {
    const dlg = el.confirm;
    $('#confirm-title').textContent = title;
    const bodyEl = $('#confirm-body');
    bodyEl.replaceChildren(...[].concat(body).filter(Boolean).map((b) => (typeof b === 'string' ? h('p', { text: b }) : b)));
    const ic = $('#confirm-icon');
    ic.className = `confirm-icon ${tone}`;
    ic.replaceChildren(icon({ danger: 'restore', shield: 'shield', warn: 'alert' }[tone] ?? 'info'));
    const okBtn = $('#confirm-ok');
    const cancelBtn = $('#confirm-cancel');
    const altBtn = $('#confirm-alt');
    okBtn.textContent = ok;
    okBtn.className = `btn ${tone === 'danger' ? 'danger' : 'primary'}`;
    cancelBtn.textContent = cancel ?? '';
    cancelBtn.hidden = cancel === null;
    altBtn.textContent = alt ?? '';
    altBtn.hidden = !alt;
    return new Promise((resolve) => {
      const done = (v) => {
        okBtn.removeEventListener('click', yes);
        cancelBtn.removeEventListener('click', no);
        altBtn.removeEventListener('click', other);
        dlg.removeEventListener('close', no);
        if (dlg.open) dlg.close();
        resolve(v);
      };
      const yes = () => done(true);
      const no = () => done(false);
      const other = () => done('alt');
      okBtn.addEventListener('click', yes);
      cancelBtn.addEventListener('click', no);
      altBtn.addEventListener('click', other);
      dlg.addEventListener('close', no);
      dlg.showModal();
      (cancel === null ? okBtn : cancelBtn).focus();
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
    renderEditButton();
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
    // 이 문서의 편집 창이 열려 있으면 (누르면 그 창으로). 창 이름: 한글 편집 · 문서 편집 · 표 편집 · 슬라이드 편집
    const ed = d?.editor;
    const win = `${(d && editorLabel(d.path)) || '편집'} 창`;
    set(el.stEditor, ed?.open ? (ed.dirty ? `${win} 열림 · 저장 안 됨` : `${win} 열림`) : '',
      ed?.open ? (ed.dirty ? `${win}에 저장하지 않은 내용이 있어요. 눌러서 ${win} 보기` : `눌러서 ${win} 보기`) : '');
    el.stEditor.classList.toggle('dirty', !!ed?.dirty);
    renderEditorWarn();
  }

  // ── 편집 창 (한글 · 문서·표·슬라이드) ─────────────────
  // main: doc:edit(열기·앞으로), doc:editor 이벤트({ path, open, dirty }). 편집 창 자체는 editor-window.mjs
  //   종류는 확장자로: hwp·hwpx·hwt → 'hwp'(한글 편집), docx → 문서 편집, xlsx → 표 편집, pptx → 슬라이드 편집 ('office')

  // 창 이름 ('ONLYOFFICE'는 쓰지 않는다. Word·Excel·PowerPoint는 형식 설명에만)
  const EDITOR_LABEL_BY_EXT = { hwp: '한글 편집', hwpx: '한글 편집', hwt: '한글 편집', docx: '문서 편집', xlsx: '표 편집', pptx: '슬라이드 편집' };
  /** 파일 경로(또는 확장자) → 'hwp' | 'office' | null */
  function editorKindOf(p) {
    const ext = String(p ?? '').includes('.') ? extOf(p) : String(p ?? '').toLowerCase();
    for (const [kind, k] of Object.entries(state.editor.kinds ?? {})) if (k?.exts?.includes(ext)) return kind;
    return null;
  }
  const editorKindInfo = (p) => state.editor.kinds?.[editorKindOf(p)] ?? null;
  /** '한글 편집' | '문서 편집' | '표 편집' | '슬라이드 편집' | null */
  const editorLabel = (p) => (editorKindOf(p) ? EDITOR_LABEL_BY_EXT[String(p ?? '').includes('.') ? extOf(p) : String(p ?? '').toLowerCase()] ?? '문서 편집' : null);
  const isEditable = (doc) => !!doc && !!editorKindOf(doc.path);
  /** 그 종류의 편집기 묶음이 있다/없다로 고쳐 둔다 (열기 결과로 알게 된 것) */
  function setEditorAvailable(p, available, reason = null) {
    const kind = editorKindOf(p);
    if (!kind) return;
    state.editor = { ...state.editor, kinds: { ...state.editor.kinds, [kind]: { ...state.editor.kinds[kind], available, reason } } };
  }

  function renderEditButton() {
    const d = state.doc;
    const b = el.btnEdit;
    const ok = isEditable(d);
    b.disabled = !ok;
    // 편집 창으로 열 수 없는 형식(옛 형식 doc·xls·ppt, csv, odt …)에서는 숨겨 문서 이름 자리를 넓힌다 (문서가 없을 때는 안내용으로 보인다)
    b.hidden = !!d && !ok;
    b.classList.toggle('active', !!d?.editor?.open);
    if (!d) { b.title = '한글·Word·Excel·PowerPoint 문서를 열면 앱 안에서 직접 편집할 수 있어요'; return; }
    const label = editorLabel(d.path);
    if (!ok) b.title = `${d.kind} 문서(.${extOf(d.path)})는 앱 안에서 직접 편집할 수 없어요. ‘프로그램에서 열기’로 고쳐 주세요. (직접 편집: HWP·HWPX·DOCX·XLSX·PPTX)`;
    else if (!editorKindInfo(d.path)?.available) b.title = `${label}기가 설치돼 있지 않아요. 눌러서 이유를 확인하세요.`;
    else if (d.editor?.open) b.title = `열려 있는 ${label} 창으로 가기`;
    else b.title = `${label} 창에서 직접 고치기 (Ctrl+E, 인터넷 없이 동작)`;
  }

  /** 대화창 위 경고: 편집 창에 저장 안 한 내용이 있을 때 */
  function renderEditorWarn() {
    el.editorWarn.hidden = !(state.doc?.editor?.open && state.doc.editor.dirty);
  }

  /** 편집기를 쓸 수 없을 때 이유를 보여 준다 (p: 문서 경로 → 종류별 문구) */
  function showEditorMissing(reason, p) {
    const office = editorKindOf(p) === 'office';
    const label = editorLabel(p) ?? '편집';
    // 이유 끝의 '(개발 중이면 …)'은 아래에 자세히 적으므로 뺀다
    const why = String(reason || `${label}기 파일이 없습니다.`).replace(/\s*\(개발 중이면.*\)\s*$/, '');
    const cmd = office ? 'npm run office-editor:build' : 'npm run editor:build';
    const what = office ? '문서·표·슬라이드 편집기 묶음(vendor/office-editor)' : '한글 편집기 묶음(vendor/rhwp-studio)';
    return confirmDialog({
      title: `${label}기를 열 수 없어요`,
      tone: 'warn', cancel: null,
      body: [
        why,
        state.packaged ? '문서 보기와 AI 대화는 그대로 쓸 수 있어요. 직접 고치려면 ‘프로그램에서 열기’를 써 주세요.'
          : h('p', {}, '개발 중이면 저장소 폴더에서 ', h('code', { text: cmd }), `를 한 번 실행한 뒤 다시 눌러 주세요. ${what}을 인터넷 없이 쓸 수 있게 만듭니다.`),
      ],
    });
  }

  // 편집 창을 열기 전에 거른 문서 (암호·옛 형식): 프로그램에서 열기를 권한다
  const OPEN_IN_APP_CODES = new Set(['UNSUPPORTED', 'ENCRYPTED', 'OLD_FORMAT']);

  /** p: 열려던 문서 경로 */
  function editorError(err, p = state.doc?.path) {
    if (err.code === 'EDITOR_MISSING') { setEditorAvailable(p, false, err.message); renderEditButton(); return showEditorMissing(err.message, p); }
    if (OPEN_IN_APP_CODES.has(err.code)) {
      toast(err.message, { type: 'info', action: { label: '프로그램에서 열기', run: actOpenInApp } });
      return null;
    }
    showError(err, '편집 창을 열지 못했어요');
    return null;
  }

  async function actEdit() {
    if (!state.doc) return;
    const p = state.doc.path;
    const label = editorLabel(p) ?? '편집';
    setBusy('edit', state.doc.editor?.open ? `${label} 창 여는 중…` : `${label}기 여는 중…`);
    let error = null; let r = null;
    try {
      r = await call(K.doc.edit());
      // 열렸으면 편집기 묶음이 있는 것
      if (!editorKindInfo(p)?.available) { setEditorAvailable(p, true); renderEditButton(); }
    } catch (err) { error = err; } finally { clearBusy('edit'); }
    if (error) { await editorError(error, p); return; }
    // Office 편집 창이 많이 열림 → 메모리 안내
    if (r?.notice) toast(r.notice, { type: 'warn', timeout: 8000 });
  }

  function onDocEditor(ev) {
    if (!ev || !state.doc || pathKey(ev.path) !== pathKey(state.doc.path)) return;
    state.doc = { ...state.doc, editor: { open: !!ev.open, dirty: !!ev.dirty } };
    renderTopbar();
    renderStatus();
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
    // 쪽이 없는 PDF(빈 시트 등)는 PDF 뷰어가 빈 화면에서 멈추므로 안내만
    if (res.pages === 0) {
      el.preview.className = 'preview';
      const card = h('div', { class: 'error-card' },
        h('div', { class: 'err-icon' }, icon('info')),
        h('h3', { text: '미리 볼 쪽이 없어요' }),
        h('p', { text: '문서에 내용이 없어 프로그램이 그릴 쪽이 없어요. 내용을 넣으면 여기에 보여요.' }));
      el.preview.replaceChildren(h('div', { class: 'preview-error' }, card));
      return Promise.resolve();
    }
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
    if (res.note) wrap.append(h('div', { class: 'pages-note', text: res.note }));
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
        !d.editor?.open ? null : d.editor.dirty
          ? '편집 창에 저장하지 않은 내용이 있어요. 되돌린 뒤 편집 창에서 다시 불러올지, 내 편집을 유지할지 고를 수 있어요.'
          : '열려 있는 편집 창도 되돌린 내용으로 다시 불러와요.',
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
      const editorNote = !r.editorNotified ? '' : state.doc?.editor?.dirty ? '. 편집 창에서 다시 불러올지 골라 주세요' : '. 편집 창도 다시 불러와요';
      toast(`${r.stoppedAi > 0 ? 'AI 작업을 멈추고 원래대로 되돌렸어요' : '원래대로 되돌렸어요'}${editorNote}`, { type: 'success' });
      // main이 이 문서의 AI 대화를 끝냈다 → 대화창에 알린다 (AI가 되돌리기 전 내용을 기억하지 않게)
      if (r.newChat && chat.items.length) {
        chat.items.push({ role: 'note', text: '문서를 원래대로 되돌려서, 다음 메시지부터 새 대화로 시작해요.' });
        if (chatCurrent() === chat) renderChat();
      }
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
    if (ne.dlg.open) renderNewDoc();
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
        h('p', { text: '새 문서를 만들면서 AI에게 내용을 부탁할 수도 있어요.' }),
        h('div', { class: 'btn-row' },
          h('button', { class: 'btn sm', onclick: pickFile }, icon('file-open'), '파일 열기'),
          h('button', { class: 'btn sm', id: 'chat-new-doc', onclick: () => openNewDoc({ start: 'ai' }) }, icon('file-new'), '새 문서')));
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

  // 첫 화면 추천 문장: 열린 문서 종류에 맞춰 (형식별 전용 기능)
  const SUGGESTS = {
    한글: ['이 문서 요약해줘', '표 내용 정리해줘', '머리말 넣어줘', '맞춤법 고쳐줘'],
    Word: ['이 문서 요약해줘', '제목 스타일 정리해줘', '목차 넣어줘', '맞춤법 고쳐줘'],
    PowerPoint: ['발표자 노트 달아줘', '슬라이드 하나 추가해줘', '이 발표 요약해줘', 'PDF로 저장해줘'],
    Excel: ['차트 만들어줘', '머리글 굵게 하고 틀 고정해줘', '이 표 요약해줘', 'PDF로 저장해줘'],
  };
  const SUGGESTS_DEFAULT = ['이 문서 요약해줘', '맞춤법 고쳐줘', '표 내용 정리해줘', 'PDF로 저장해줘'];
  const suggestsFor = (doc) => SUGGESTS[doc?.kind] ?? SUGGESTS_DEFAULT;

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
        h('div', { class: 'suggests' }, ...suggestsFor(state.doc).map((s) => h('button', { class: 'suggest', onclick: () => sendPrompt(s) }, icon('sparkle'), s)))));
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
    if (item.role === 'note') return h('div', { class: 'chat-divider', role: 'note', text: item.text });
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

  /**
   * 대화 한 번 준비: 사용자 말풍선과 답 자리를 넣고 '작업 중'으로 둔다 → { chat, msg }.
   * 대화 번호를 받으면 chatStarted, 못 받으면 chatFailed (번호보다 먼저 온 이벤트는 chat.pending이 받는다)
   */
  function queueChat(path, text) {
    const chat = chatFor(path);
    const msg = { role: 'assistant', parts: [], done: false, status: '생각하는 중…', chatId: null };
    chat.items.push({ role: 'user', text }, msg);
    chat.busy = true;
    chat.pending = msg;
    chat.stopRequested = false;
    chat.cancelSent = null;
    setBusy(`ai:${pathKey(path)}`, 'AI가 작업 중…');
    if (chatCurrent() === chat) renderChat();
    return { chat, msg };
  }

  function chatStarted(chat, msg, chatId) {
    msg.chatId = chatId;
    if (chat.pending === msg) chat.pending = null;
    // 이미 끝났으면('end'가 먼저 옴) 대화 번호를 남기지 않는다
    if (!msg.done) chat.chatId = chatId;
    // 보내는 동안 멈추기를 눌렀으면 지금 멈춘다
    if (chat.stopRequested && !msg.done) cancelChatId(chat, chatId);
  }

  function chatFailed(chat, msg, path, err) {
    if (chat.pending === msg) chat.pending = null;
    msg.done = true;
    msg.error = err?.message || 'AI 대화를 시작하지 못했어요';
    chat.busy = false;
    clearBusy(`ai:${pathKey(path)}`);
    if (err?.code === 'NO_PROVIDER' || err?.code === 'MODULE_MISSING') loadProviders(true);
    if (chatCurrent() === chat) renderChat();
  }

  /** 편집 창에 저장 안 한 내용이 있을 때 보내기 전에 묻는다 → 'send' | 'editor' | 'cancel' */
  async function askEditorDirty() {
    const r = await confirmDialog({
      title: '편집 창에 저장하지 않은 내용이 있어요',
      tone: 'warn',
      body: [
        'AI는 저장된 파일을 고칩니다. 지금 보내면 편집 창에서 고친 내용은 AI가 보지 못하고, 나중에 편집 창에서 저장하면 AI가 고친 내용을 덮어쓸 수 있어요.',
        '편집 창에서 먼저 저장(Ctrl+S)한 뒤 보내 주세요.',
      ],
      ok: '편집 창 보기', alt: '그대로 보내기', cancel: '취소',
    });
    return r === true ? 'editor' : r === 'alt' ? 'send' : 'cancel';
  }

  async function sendPrompt(text, { editorOk = false } = {}) {
    text = String(text ?? '').trim();
    if (!text || !state.doc || chatBlocker()) return;
    const path = state.doc.path;
    if (chatFor(path).busy) return;
    // 편집 창에 저장 안 한 내용이 있으면 먼저 묻는다 (입력한 글은 그대로 둔다)
    if (!editorOk && state.doc.editor?.dirty) {
      const choice = await askEditorDirty();
      if (choice === 'editor') { actEdit(); return; }
      if (choice !== 'send' || !state.doc || pathKey(state.doc.path) !== pathKey(path) || chatFor(path).busy) return;
      editorOk = true;
    }
    el.prompt.value = '';
    autoGrow();
    const { chat, msg } = queueChat(path, text);
    try {
      chatStarted(chat, msg, (await call(K.ai.send({ prompt: text, editorOk }))).chatId);
    } catch (err) {
      if (err.code !== 'EDITOR_DIRTY') { chatFailed(chat, msg, path, err); return; }
      // 화면이 편집 창 상태를 아직 몰랐던 경우: 대화를 빼고 글을 되돌린 뒤 묻는다
      dropQueuedChat(path, { chat, msg });
      renderChat();
      if (!el.prompt.value.trim()) { el.prompt.value = text; autoGrow(); }
      if (state.doc && pathKey(state.doc.path) === pathKey(path)) onDocEditor({ path, open: true, dirty: true });
      const choice = await askEditorDirty();
      if (choice === 'editor') actEdit();
      else if (choice === 'send') sendPrompt(text, { editorOk: true });
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

  // ── 새 문서 ──────────────────────────────────────────
  // 창은 index.html #dlg-new. main 쪽: doc:newInfo(기본 폴더·이름) · doc:newCheck(이름 확인) · doc:pickFolder · doc:new(만들고 열기 + AI 시작)

  const NEW_FORMATS = {
    hwpx: { kind: '한글', hint: '개방형 표준(OWPML) 한글 문서예요.' },
    hwp: { kind: '한글', hint: '예전 한글 버전과도 주고받기 좋은 기존 한글 형식이에요.' },
    docx: { kind: 'Word', hint: 'Microsoft Word 문서예요.' },
    pptx: { kind: 'PowerPoint', hint: '빈 제목 슬라이드 하나로 시작해요.' },
    xlsx: { kind: 'Excel', hint: '빈 시트 하나로 시작해요.' },
  };
  const NEW_SUGGESTS = {
    한글: ['주간 회의록 양식 만들어줘', '공문 초안 써줘', '보고서 목차 잡아줘'],
    Word: ['주간 회의록 양식 만들어줘', '보도자료 초안 써줘', '보고서 목차 잡아줘'],
    PowerPoint: ['회사 소개 발표 5장 만들어줘', '프로젝트 시작 회의 발표 만들어줘'],
    Excel: ['월별 가계부 표 만들어줘', '업무 일정표 만들어줘', '재고 관리 표 만들어줘'],
  };
  // 이름 칸 아래에 보여 줄 오류 (나머지는 창 아래쪽)
  const NAME_CODES = new Set(['EMPTY', 'BAD_CHARS', 'SPACES', 'RESERVED', 'TOO_LONG', 'PATH_TOO_LONG', 'EXISTS']);

  // format은 창을 닫아도 기억한다 (이번 실행 동안). start는 열 때마다 '빈 문서'부터
  // edit: 만든 뒤 바로 편집 창 열기 (한글·docx·xlsx·pptx, 이번 실행 동안 기억)
  const nd = { format: 'hwpx', start: 'blank', edit: true, folder: '', autoName: '', busy: false, opening: false, seq: 0, timer: 0 };
  const ne = {
    dlg: $('#dlg-new'), name: $('#new-name'), nameBox: $('#new-name-box'), nameMsg: $('#new-name-msg'), ext: $('#new-ext'),
    folder: $('#new-folder'), folderText: $('#new-folder-text'), change: $('#new-folder-change'), formatHint: $('#new-format-hint'),
    aiChoice: $('#new-start-ai'), aiHint: $('#new-ai-hint'), aiBox: $('#new-ai-box'), request: $('#new-request'),
    requestMsg: $('#new-request-msg'), suggests: $('#new-suggests'), error: $('#new-error'),
    create: $('#new-create'), cancel: $('#new-cancel'), close: $('#new-close'),
    editField: $('#new-edit-field'), edit: $('#new-edit'), editRow: $('#new-edit-row'), editDesc: $('#new-edit-desc'),
  };
  /** 이 형식으로 만들면 편집 창을 열 수 있는지 (편집 창 형식이고 그 편집기 묶음이 있음) */
  const newDocEditable = (format) => !!NEW_FORMATS[format] && !!editorKindInfo(format)?.available;
  // 만든 뒤 바로 편집 창에서 할 일 (빈 문서로 시작할 때)
  const NEW_EDIT_DESC = {
    한글: '한글 편집 창에서 바로 글을 써요. 인터넷 없이 동작해요.',
    Word: '문서 편집 창에서 바로 글을 써요. 인터넷 없이 동작해요.',
    PowerPoint: '슬라이드 편집 창에서 바로 슬라이드를 꾸며요. 인터넷 없이 동작해요.',
    Excel: '표 편집 창에서 바로 표를 채워요. 인터넷 없이 동작해요.',
  };

  /** 새 문서 창 열기. format: 미리 고를 형식, start: 'blank' | 'ai' (AI를 못 쓰면 빈 문서) */
  async function openNewDoc({ format, start } = {}) {
    if (ne.dlg.open) { ne.name.focus(); return; }
    if (nd.opening || el.confirm.open) return;
    if (el.settings.open) el.settings.close();
    if (NEW_FORMATS[format]) nd.format = format;
    nd.start = start === 'ai' ? 'ai' : 'blank';
    nd.opening = true;
    try {
      const info = await call(K.doc.newInfo({ format: nd.format }));
      nd.folder = info.folder;
      nd.autoName = info.name;
    } catch (err) { showError(err); return; } finally { nd.opening = false; }
    ne.name.value = nd.autoName;
    ne.request.value = '';
    setNameMsg(null);
    setRequestMsg('');
    setNewError('');
    renderNewDoc();
    ne.dlg.showModal();
    (nd.start === 'ai' && newDocAi().ok ? ne.request : ne.name).focus();
    if (document.activeElement === ne.name) ne.name.select();
  }

  /** AI로 시작할 수 있는지 → { ok, wait?, reason? } */
  function newDocAi() {
    if (state.providerError) return { ok: false, reason: `AI 기능을 쓸 수 없어요: ${state.providerError}` };
    if (!state.providers) return { ok: false, wait: true, reason: 'AI 프로그램을 확인하는 중이에요…' };
    if (!state.providers.selected) return { ok: false, reason: 'AI 프로그램(Claude Code 또는 Codex)을 찾지 못해 쓸 수 없어요.' };
    return { ok: true };
  }

  function renderNewDoc() {
    const busy = nd.busy;
    const ai = newDocAi();
    if (!ai.ok && nd.start === 'ai') nd.start = 'blank';
    const f = NEW_FORMATS[nd.format];
    for (const r of $$('input[name="new-format"]', ne.dlg)) { r.checked = r.value === nd.format; r.disabled = busy; }
    for (const r of $$('input[name="new-start"]', ne.dlg)) { r.checked = r.value === nd.start; r.disabled = busy || (r.value === 'ai' && !ai.ok); }
    ne.ext.textContent = `.${nd.format}`;
    ne.formatHint.textContent = f.hint;
    ne.folderText.textContent = nd.folder;
    ne.folder.title = nd.folder;
    ne.aiChoice.classList.toggle('disabled', !ai.ok);
    ne.aiHint.hidden = ai.ok;
    ne.aiHint.replaceChildren(...[icon(ai.wait ? 'clock' : 'info'), h('span', { text: ai.reason ?? '' }),
      ai.ok || ai.wait ? null : h('button', { class: 'link-btn', type: 'button', onclick: () => { ne.dlg.close(); openSettings('ai'); } }, '설정 열기')].filter(Boolean));
    ne.aiBox.hidden = nd.start !== 'ai';
    // 편집 창 형식(한글·docx·xlsx·pptx): 만든 뒤 바로 편집 창 열기
    const editKind = editorKindOf(nd.format);
    const editOk = newDocEditable(nd.format);
    const label = editorLabel(nd.format) ?? '편집';
    ne.editField.hidden = !editKind;
    ne.edit.checked = editOk && nd.edit;
    ne.edit.disabled = busy || !editOk;
    ne.editRow.classList.toggle('disabled', !!editKind && !editOk);
    ne.editDesc.textContent = !editOk
      ? (state.packaged ? `${label}기가 설치돼 있지 않아 쓸 수 없어요.`
        : `${label}기 묶음이 없어요. ${editKind === 'office' ? 'npm run office-editor:build' : 'npm run editor:build'}로 먼저 만들어 주세요.`)
      : nd.start === 'ai' ? 'AI가 채우는 내용이 편집 창에도 바로 보여요.' : NEW_EDIT_DESC[f.kind];
    ne.request.placeholder = `예: ${NEW_SUGGESTS[f.kind][0]}`;
    ne.suggests.replaceChildren(...NEW_SUGGESTS[f.kind].map((s) => h('button', {
      class: 'suggest-chip', type: 'button', disabled: busy,
      onclick: () => { ne.request.value = s; setRequestMsg(''); ne.request.focus(); },
    }, icon('sparkle'), s)));
    ne.name.disabled = busy;
    ne.request.disabled = busy;
    ne.change.disabled = busy;
    ne.cancel.disabled = busy;
    ne.close.disabled = busy;
    ne.create.disabled = busy;
    ne.create.replaceChildren(...(busy ? [spinner(), '만드는 중…'] : nd.start === 'ai' ? [icon('sparkle'), '만들고 AI에게 부탁하기'] : ['만들기']));
    ne.dlg.classList.toggle('busy', busy);
  }

  function setNameMsg(r) {
    const bad = !!r && !r.ok;
    ne.nameBox.classList.toggle('invalid', bad);
    ne.name.setAttribute('aria-invalid', String(bad));
    ne.nameMsg.className = `nd-msg${bad ? ' error' : ''}`;
    if (!bad) { ne.nameMsg.replaceChildren(); return; }
    ne.nameMsg.replaceChildren(...[icon('alert'), h('span', { text: r.error }),
      r.suggestion ? h('button', {
        class: 'link-btn', type: 'button',
        onclick: () => { ne.name.value = r.suggestion; checkNewName(); ne.name.focus(); },
      }, `‘${r.suggestion}’(으)로 바꾸기`) : null].filter(Boolean));
  }
  function setRequestMsg(text) {
    ne.requestMsg.className = `nd-msg${text ? ' error' : ''}`;
    ne.requestMsg.replaceChildren(...(text ? [icon('alert'), h('span', { text })] : []));
  }
  function setNewError(text) {
    ne.error.hidden = !text;
    ne.error.replaceChildren(...(text ? [icon('alert'), h('span', { text })] : []));
  }

  /** 이름 확인 (main이 폴더·형식·이미 있는 파일까지 본다) → 결과, 그 사이 다시 확인했으면 null */
  async function checkNewName() {
    clearTimeout(nd.timer);
    const seq = ++nd.seq;
    let r;
    try { r = await call(K.doc.newCheck({ folder: nd.folder, name: ne.name.value, format: nd.format })); } catch (err) {
      r = { ok: false, code: err.code, error: err.message, suggestion: null };
    }
    if (seq !== nd.seq) return null;
    // 이름 칸과 상관없는 오류(폴더 등)는 창 아래쪽에
    if (r.ok || NAME_CODES.has(r.code)) { setNameMsg(r); if (r.ok) setNewError(''); } else { setNameMsg(null); setNewError(r.error); }
    return r;
  }

  /** 기본 이름('새 문서', '새 문서 (2)' …)을 지금 형식·폴더에 맞춘다 (사용자가 고친 이름은 두고) */
  async function refreshAutoName() {
    try {
      const info = await call(K.doc.newInfo({ format: nd.format, folder: nd.folder }));
      if (ne.name.value === nd.autoName) ne.name.value = info.name;
      nd.autoName = info.name;
    } catch {}
  }

  async function setNewFormat(format) {
    if (!NEW_FORMATS[format] || format === nd.format) return;
    nd.format = format;
    renderNewDoc();
    await refreshAutoName();
    checkNewName();
  }

  async function changeNewFolder() {
    let folder;
    try { folder = await call(K.doc.pickFolder(nd.folder)); } catch (err) { setNewError(err.message); return; }
    if (!folder || !ne.dlg.open) return;
    nd.folder = folder;
    renderNewDoc();
    await refreshAutoName();
    checkNewName();
  }

  async function createNewDoc() {
    if (nd.busy || !ne.dlg.open) return;
    setNewError('');
    const ai = nd.start === 'ai';
    const request = ne.request.value.trim();
    if (ai && !newDocAi().ok) { renderNewDoc(); return; }
    if (ai && !request) { setRequestMsg('AI에게 부탁할 내용을 적어 주세요'); ne.request.focus(); return; }
    const check = await checkNewName();
    if (!check) return;
    if (!check.ok) { (NAME_CODES.has(check.code) ? ne.name : ne.change).focus(); return; }
    nd.busy = true;
    renderNewDoc();
    setBusy('newdoc', '새 문서 만드는 중…');
    // 같은 경로의 문서를 지웠다가 다시 만드는 경우 예전 대화를 잇지 않는다 (main도 AI 대화·원래대로 기준을 새로 시작)
    const old = state.chats.get(pathKey(check.path));
    if (old && !old.busy) state.chats.delete(pathKey(check.path));
    // AI로 시작하면 대화를 미리 준비해 둔다 (main이 문서를 열자마자 대화를 시작하므로 먼저 오는 이벤트도 받게)
    const queued = ai && !chatFor(check.path).busy ? queueChat(check.path, request) : null;
    try {
      const edit = newDocEditable(nd.format) && nd.edit;
      const res = await call(K.doc.create({ format: nd.format, folder: nd.folder, name: ne.name.value, start: nd.start, request: ai ? request : undefined, edit }));
      nd.busy = false;
      ne.dlg.close();
      const path = res.doc?.path ?? check.path;
      if (queued) {
        if (pathKey(path) !== pathKey(check.path)) moveChat(check.path, path);
        if (res.chat?.chatId) chatStarted(queued.chat, queued.msg, res.chat.chatId);
        else chatFailed(queued.chat, queued.msg, path, res.chatError);
        if (state.settings.chatCollapsed) toggleChat(false);
      }
      if (ai && !res.chat?.chatId) toast(`${baseName(path)}을(를) 만들었지만 AI를 시작하지 못했어요`, { type: 'warn' });
      else toast(ai ? `${baseName(path)}을(를) 만들고 AI에게 부탁했어요` : `새 문서를 만들었어요: ${baseName(path)}`, { type: 'success' });
      // 편집 창은 main이 열었다 (못 열었으면 이유만 알린다. 문서는 만들어졌음)
      if (res.editorError) editorError(res.editorError, path);
      else if (res.editor?.notice) toast(res.editor.notice, { type: 'warn', timeout: 8000 });
    } catch (err) {
      if (queued) dropQueuedChat(check.path, queued);
      nd.busy = false;
      renderNewDoc();
      if (NAME_CODES.has(err.code)) {
        setNameMsg({ ok: false, error: err.message });
        if (err.code === 'EXISTS') checkNewName(); // 바꿀 이름 제안을 받는다
        ne.name.focus();
      } else if (err.code === 'NO_REQUEST') {
        setRequestMsg(err.message);
        ne.request.focus();
      } else setNewError(err.message);
    } finally { clearBusy('newdoc'); }
  }

  /** 문서를 만들지 못했을 때 미리 넣은 대화를 뺀다 */
  function dropQueuedChat(path, { chat, msg }) {
    const i = chat.items.indexOf(msg);
    if (i > 0) chat.items.splice(i - 1, 2);
    if (chat.pending === msg) chat.pending = null;
    chat.busy = false;
    clearBusy(`ai:${pathKey(path)}`);
    if (!chat.items.length) state.chats.delete(pathKey(path));
    if (chatCurrent() === chat) renderChat();
  }

  /** 미리 짐작한 경로와 실제 경로가 다를 때 대화를 옮긴다 */
  function moveChat(from, to) {
    const c = state.chats.get(pathKey(from));
    if (!c) return;
    state.chats.delete(pathKey(from));
    state.chats.set(pathKey(to), c);
    if (state.busy.has(`ai:${pathKey(from)}`)) { clearBusy(`ai:${pathKey(from)}`); setBusy(`ai:${pathKey(to)}`, 'AI가 작업 중…'); }
    if (chatCurrent() === c) renderChat();
  }

  function bindNewDoc() {
    for (const r of $$('input[name="new-format"]', ne.dlg)) r.addEventListener('change', () => { if (r.checked) setNewFormat(r.value); });
    for (const r of $$('input[name="new-start"]', ne.dlg)) {
      r.addEventListener('change', () => {
        if (!r.checked) return;
        nd.start = r.value;
        setRequestMsg('');
        renderNewDoc();
        if (nd.start === 'ai') ne.request.focus();
      });
    }
    ne.name.addEventListener('input', () => {
      setNewError('');
      clearTimeout(nd.timer);
      nd.timer = setTimeout(checkNewName, 150);
    });
    ne.request.addEventListener('input', () => setRequestMsg(''));
    ne.edit.addEventListener('change', () => { nd.edit = ne.edit.checked; });
    ne.change.addEventListener('click', changeNewFolder);
    ne.folder.addEventListener('dblclick', changeNewFolder);
    ne.create.addEventListener('click', createNewDoc);
    ne.cancel.addEventListener('click', () => ne.dlg.close());
    ne.close.addEventListener('click', () => ne.dlg.close());
    // 만드는 중에는 Esc로 닫지 않는다
    ne.dlg.addEventListener('cancel', (e) => { if (nd.busy) e.preventDefault(); });
    ne.dlg.addEventListener('close', () => { clearTimeout(nd.timer); nd.seq++; });
    // Enter = 만들기 (글 상자에서는 Shift+Enter가 줄바꿈). 단추 위의 Enter는 그 단추를 누른다
    ne.dlg.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.shiftKey || e.altKey || e.isComposing || e.keyCode === 229) return;
      if (e.target.tagName === 'BUTTON') return;
      e.preventDefault();
      createNewDoc();
    });
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
        if (r.result?.warning) toast(String(r.result.warning), { type: 'warn', timeout: 10000 });
        else toast(on ? `${it.name}에 연결했어요.${it.key === 'cli' ? ' 새 터미널에서 office-kit 명령을 쓸 수 있어요.' : ` ${it.name}을(를) 다시 시작하면 적용돼요.`}` : `${it.name} 연결을 해제했어요.`, { type: 'success' });
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
    if (state.smoke.newdoc) {
      // AI 선택지가 제 모양으로 보이게 AI 확인을 기다린다
      await Promise.race([state.providersReady, sleep(8000)]);
      await openNewDoc({ start: state.smoke.newdoc === 'ai' ? 'ai' : 'blank' });
      if (state.smoke.newdoc === 'ai') ne.request.value = '주간 회의록 양식 만들어줘';
      await sleep(300);
    }
    if (state.smoke.settings) {
      try { await openSettings(state.smoke.settings); } catch {}
      await sleep(300);
    }
    // '직접 편집' 단추를 누를 수 있는지도 함께 (편집 창 형식 확인용)
    const editable = !el.btnEdit.hidden && !el.btnEdit.disabled;
    setTimeout(() => K.smokeReady({ ...info, editable }), 50);
  }

  // ── 연결 · 시작 ──────────────────────────────────────

  function bindUi() {
    $('#btn-open').addEventListener('click', pickFile);
    $('#btn-new').addEventListener('click', () => openNewDoc());
    $('#dz-open').addEventListener('click', pickFile);
    $('#dz-new').addEventListener('click', () => openNewDoc());
    // 끌어놓기 칸의 빈 곳을 눌러도 파일 열기
    $('#dropzone').addEventListener('click', (e) => { if (!e.target.closest('button')) pickFile(); });
    $('#btn-refresh').addEventListener('click', () => refreshPreview({ reason: 'manual' }));
    $('#btn-restore').addEventListener('click', actRestore);
    el.btnEdit.addEventListener('click', actEdit);
    el.stEditor.addEventListener('click', actEdit);
    $('#editor-warn-show').addEventListener('click', actEdit);
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

    // 단축키 (메뉴에 없는 것. Ctrl+N은 메뉴에도 있지만 화면이 먼저 받으면 여기서 연다)
    window.addEventListener('keydown', (e) => {
      if (!e.ctrlKey || e.altKey) return;
      if (e.code === 'KeyN' && !e.shiftKey) { e.preventDefault(); openNewDoc(); return; }
      if (e.key === '=' || e.key === '+') { e.preventDefault(); stepZoom(1); }
      else if (e.key === '-') { e.preventDefault(); stepZoom(-1); }
      else if (e.key === '0') { e.preventDefault(); state.zoom = 'auto'; applyZoom(); }
    });

    bindNewDoc();
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
      // 편집 창이 저장했거나 AI가 고치는 중이면 (거기서 이미 알리므로) 알림 없이 새로 고친다
      refreshPreview({ ifChanged: true, reason: ev.by === 'editor' ? 'editor' : chatCurrent()?.busy ? 'ai' : 'watch' });
    });
    K.on.docEditor(onDocEditor);
    K.on.aiEvent(onAiEvent);
    K.on.menu((cmd) => {
      if (cmd === 'open') pickFile();
      else if (cmd === 'new') openNewDoc();
      else if (cmd === 'edit') { if (state.doc) actEdit(); }
      else if (cmd === 'refresh') { if (state.doc) refreshPreview({ reason: 'manual' }); }
      else if (cmd === 'settings') openSettings('ai');
    });
    K.on.toast((t) => {
      toast(String(t?.message ?? ''), { type: t?.type ?? 'info' });
      // 확인용 캡처: 인수로 받은 파일을 못 열었으면 알림이 뜬 화면을 찍고 실패로 끝낸다
      if (state.smoke && t?.type === 'error' && !state.doc) smokeSignal({ ok: false, kind: 'open-error', error: String(t?.message ?? '') });
    });
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
    state.packaged = init.packaged !== false;
    if (init.editor) state.editor = { ...state.editor, ...init.editor };
    applyChatLayout();
    renderRecent(init.recent);
    if (init.doc) onDocOpened(init.doc);
    else { renderTopbar(); renderStatus(); renderChat(); }
    state.providersReady = loadProviders(false);
    // 열 파일 없이 시작한 확인용 캡처는 빈 화면을 찍는다
    if (state.smoke && !init.doc && !init.opening) setTimeout(() => smokeSignal({ ok: true, kind: 'empty' }), 600);
  }

  start();
})();
