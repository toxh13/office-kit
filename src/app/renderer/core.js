// core.js — Office Kit 화면 공용: 화면 언어·작은 도우미·알림·확인 창·공용 상태·시작(K.init은 여기서 한 번만).
// 같은 index.html이 탭 창의 세 뷰가 된다: ?view=home(홈 탭) · ?view=doc(문서 탭) · ?view=panel(AI 패널).
// home.js(홈·문서 탭)와 panel.js(AI 대화)는 OKCore에서 받아 쓰고, 서로는 bridge(이름 붙은 호출 → 다른 뷰면 main 중계)로만 부른다.
'use strict';
(() => {
  const K = window.officeKit;
  const MD = window.OKMarkdown;

  // ── 화면 뷰 종류 ── 'home' | 'panel' | 'doc' (없으면 home). CSS는 body[data-view]로 보일 것을 고른다
  const VIEW = (() => {
    const v = new URLSearchParams(location.search).get('view');
    return ['home', 'panel', 'doc'].includes(v) ? v : 'home';
  })();
  document.body.dataset.view = VIEW;

  // ── 화면 언어 ── 한국어 원문을 사전으로 바꾼다 (문서 내용·파일 이름·대화 글은 그대로)
  const NO_I18N = [
    '.paper > p', '.doc-table', '.sheet-scroll', '.s-title', '.s-line', '.bubble', '.md',
    '.diff-text:not(:has(> .diff-empty-line))', '.diff-where', 'td.addr', '.pii-row td.val', '.pii-row td.masked', '.pii-row td.ctx',
    '.recent-name', '.recent-dir', '.pdf-name', 'td.file', '.dir', '.path', '.tool-arg',
  ].join(', ');
  const I18N = (() => {
    const d = K?.i18n; const lib = window.OfficeKitI18n;
    if (!d || !lib || !['en', 'ja'].includes(d.lang)) return { t: (x) => x, lang: 'ko', setting: d?.setting ?? 'system' };
    const tr = lib.compile(d.dict, d.lang);
    document.documentElement.lang = d.lang;
    lib.translateDom(document.documentElement, tr, { skip: NO_I18N });
    return { ...tr, setting: d.setting ?? 'system' };
  })();

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
      err.detail = r?.detail ?? null;
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
    pdf: 'PDF',
  };
  const KIND_STYLE = {
    한글: { cls: 'k-hwp', letter: '한' }, Word: { cls: 'k-word', letter: 'W' },
    PowerPoint: { cls: 'k-ppt', letter: 'P' }, Excel: { cls: 'k-xls', letter: 'X' },
    PDF: { cls: 'k-pdf', letter: 'PDF' },
  };
  const kindStyle = (kind) => KIND_STYLE[kind] ?? { cls: 'k-etc', letter: '?' };

  const ENGINE_LABEL = {
    rhwp: '내장 엔진 (rhwp)', file: '내장 엔진', pdf: 'PDF 뷰어', word: 'Microsoft Word', excel: 'Microsoft Excel',
    ppt: 'PowerPoint', powerpoint: 'PowerPoint', hwp: '한컴오피스', hancom: '한컴오피스',
  };
  const ENGINE_SHORT = { rhwp: '내장 엔진', file: '내장 엔진', word: 'Word 변환', excel: 'Excel 변환', ppt: 'PowerPoint 변환', powerpoint: 'PowerPoint 변환', hwp: '한컴 변환', hancom: '한컴 변환' };

  // 시간 표시
  const two = (n) => String(n).padStart(2, '0');
  const LOCALE = { ko: 'ko-KR', en: 'en-US', ja: 'ja-JP' };
  function timeText(ms) {
    const d = new Date(ms);
    if (I18N.lang !== 'ko') return d.toLocaleTimeString(LOCALE[I18N.lang] ?? 'en-US', { hour: 'numeric', minute: '2-digit' });
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

  // 공용 칸만 여기서 만든다. home.js·panel.js는 자기 칸을 이 객체에 더하고, 남의 칸은 읽지도 쓰지도 않는다.
  const state = {
    exts: [],
    settings: {},
    version: '',
    doc: null,
    providers: null,        // { providers, selected } 또는 null
    providerError: null,
    providersReady: null,   // 처음 AI 확인이 끝나면 풀리는 Promise
    busy: new Map(),
    smoke: null,
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

  // 공용 요소 (home.js·panel.js는 자기 요소를 따로 잡는다)
  const el = {
    app: $('#app'), toasts: $('#toasts'), confirm: $('#dlg-confirm'),
    stBusy: $('#st-busy'), stBusyText: $('#st-busy-text'), stReady: $('#st-ready'),
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

  /**
   * 오류 → { text: 쉬운 한국어, detail: 원래 글(다르면), retry: 다시 해 볼 만한지 }
   * 엔진 오류 코드(UNSUPPORTED_FILE_FORMAT 등)·전체 경로·Office의 영어 문장은 '자세히'로 미룬다
   */
  function friendlyError(err) {
    const raw = String(err?.message ?? err ?? '').trim();
    // 전체 경로 → 파일 이름 (폴더 부분만 지운다. 폴더·파일 이름에 띄어쓰기가 있어도)
    let text = raw.replace(/(?:[A-Za-z]:|\\\\[^\\/\s]+)[\\/](?:[^\\/:*?"<>|\r\n]+[\\/])+/g, '');
    let retry = true;
    // 이미 쉬운 한국어(오류 코드·경로·영어 문장 없음)면 그 글을 그대로 쓴다
    const plain = !/오류코드|[A-Z][A-Z0-9]*_[A-Z0-9_]{3,}|[A-Za-z]:[\\/]|[A-Za-z]{3,}\s+[A-Za-z]{2,}\s+[A-Za-z]{2,}/.test(raw);
    if (err?.code === 'ENCRYPTED' || /암호가 걸린|password|encrypted/i.test(raw)) {
      if (!plain || !/프로그램에서/.test(raw)) text = '암호가 걸린 문서라 열 수 없어요. ‘프로그램에서 열기’로 열어 암호를 넣어 주세요.';
      retry = false;
    } else if (/UNSUPPORTED_FILE_FORMAT|오류코드|유효하지 않은 파일|손상|corrupt|damaged|not a valid|unreadable content|cannot open the file/i.test(raw) || err?.code === 'BAD_FORMAT') {
      if (!plain || !/프로그램에서/.test(raw)) text = '문서가 손상됐거나 지원하지 않는 형식이라 열 수 없어요. ‘프로그램에서 열기’로 열어 보세요.';
      retry = false;
    } else {
      // 괄호 안 영어 문장(Office가 낸 원문)은 자세히로
      text = text.replace(/\s*\((?=[^()]*[A-Za-z]{3,}\s+[A-Za-z]{2,})[^()가-힣]*\)/g, '').trim();
    }
    text = text || raw;
    // 자세히: 엔진이 낸 원래 글 (보여 준 글과 같으면 빼고)
    const detail = [raw !== text ? raw : null, err?.detail].filter(Boolean).join('\n');
    return { text, detail: detail || null, retry };
  }
  function showDetail(title, detail) {
    confirmDialog({ title, body: h('pre', { class: 'err-detail', text: detail }), ok: '닫기', cancel: null });
  }

  function showError(err, prefix) {
    const f = friendlyError(err);
    const message = prefix ? `${prefix}: ${f.text}` : f.text;
    if (err.code === 'HANCOM_SETUP') toast(message, { type: 'warn', action: { label: '설정 열기', run: () => bridge.call('home.openSettings', 'hancom') } });
    else toast(message, { type: 'error', ...(f.detail ? { action: { label: '자세히', run: () => showDetail('오류 자세히', f.detail) } } : {}) });
  }

  // ── 확인 창 ──────────────────────────────────────────

  /**
   * body: 문자열 | 노드 배열. alt: 세 번째 단추 이름 (누르면 'alt'), cancel: null이면 취소 단추 없음
   * → Promise<true | false | 'alt'>
   */
  // 확인 창은 하나뿐이라 대기열로 하나씩 띄운다 (앞 창이 닫힌 뒤 다음 창)
  const dialogQueue = window.OKDialogQueue.createDialogQueue();
  function confirmDialog(opts) { return dialogQueue.run(() => showConfirm(opts)); }

  function showConfirm({ title, body, ok = '확인', cancel = '취소', alt = null, tone = 'info' }) {
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

  // ── 다리 · 시작 ──────────────────────────────────────

  // home.js ↔ panel.js 사이 호출은 이 다리로만 한다. 이 뷰에 없는 이름은 main을 거쳐 다른 뷰로 (VIEW_ROUTES)
  const bridge = window.OKBridge.createBridge({ relay: (name, args) => K.view.call(name, args) });
  // 다른 뷰가 보낸 이름: 이 뷰가 provide한 것만 부른다
  K.on.viewCalled(({ name, args } = {}) => bridge.receive(name, Array.isArray(args) ? args : []));

  // init 결과를 받을 함수들: 등록 차례(home → panel)대로 부른다
  const initFns = [];
  function onInit(fn) { initFns.push(fn); }

  async function start() {
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
    for (const fn of initFns) {
      const r = fn(init);
      if (r && typeof r.then === 'function') await r;
    }
  }

  // home.js·panel.js가 단추·구독을 다 건 뒤(defer 스크립트는 DOMContentLoaded 전에 모두 돈다) 시작한다
  document.addEventListener('DOMContentLoaded', start, { once: true });

  globalThis.OKCore = Object.freeze({
    K, MD, I18N, view: VIEW, $, $$, h, icon, spinner, call, sleep,
    pathKey, extOf, baseName, dirName, KIND_BY_EXT, kindStyle, ENGINE_LABEL, ENGINE_SHORT, timeText, dayDiff,
    whenText, relText, setBusy, clearBusy, toast, friendlyError, showDetail, showError, confirmDialog, state,
    el, bridge, onInit,
  });
})();
