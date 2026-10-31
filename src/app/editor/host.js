// host.js — 한글 편집 창: 내장 rhwp-studio(iframe)를 SDK로 띄우고, 저장·닫기·바깥 변경을 앱(preload)과 잇는다.
const api = window.officeKitEditor;
// 시험용 고리 (main이 ?test=1로 열 때만)
const TEST = new URLSearchParams(location.search).get('test') === '1';
const $ = (id) => document.getElementById(id);

// 원래 형식으로 저장: hwp·hwt는 HWP, hwpx는 HWPX
const FORMAT_OF = { hwp: 'hwp', hwt: 'hwp', hwpx: 'hwpx' };
const DIRTY_POLL_MS = 800;
// 자동 저장: 고친 내용이 있으면 10분마다 복구 사본으로 (원본 파일은 그대로)
const AUTOSAVE_MS = 10 * 60 * 1000;
// 저장 뒤 이 시간 안에 studio가 '바뀜'을 다시 표시하면 저장한 내용과 같은지 확인한다
const RECHECK_MS = 4000;

const state = {
  editor: null,
  name: '', ext: '', format: 'hwp', readOnly: false,
  loaded: false, failed: null,
  dirty: false, saving: null, reloading: false, missing: false,
  // 저장하는 동안 고친 내용이 파일에 안 들어갔는데 studio는 '저장됨'으로 아는 경우 → 다음 저장까지 '저장 안 됨'
  lateEdits: false,
  // 마지막 저장 { format, sha, at } — 저장 직후 studio가 늦게 '바뀜'을 표시하면 내용이 같은지 견준다
  saved: null,
  poll: null, pollTimer: 0,
  bannerKind: null,
};

const log = (message) => { try { api?.log(String(message)); } catch {} };

/** main 응답 { ok, value } 풀기 */
function unwrap(res) {
  if (res && res.ok) return res.value;
  const err = new Error(res?.error || '알 수 없는 오류가 났습니다');
  err.code = res?.code ?? null;
  throw err;
}

// ── 화면 ────────────────────────────────────────────────

function icon(name) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'i');
  const use = document.createElementNS('http://www.w3.org/2000/svg', 'use');
  use.setAttribute('href', `#i-${name}`);
  svg.append(use);
  return svg;
}

function setDocInfo(doc) {
  if (doc.name) state.name = String(doc.name);
  if (doc.ext) {
    state.ext = String(doc.ext).toLowerCase();
    state.format = FORMAT_OF[state.ext] ?? 'hwp';
  }
  if (typeof doc.readOnly === 'boolean') state.readOnly = doc.readOnly;
  $('doc-name').textContent = state.name;
  $('doc-name').title = state.name;
  $('doc-kind').textContent = state.ext === 'hwt' ? '한글 서식' : '한글';
  $('doc-readonly').hidden = !state.readOnly;
  document.title = `${state.name} - 한글 편집`;
}

function renderState() {
  const badge = $('doc-state');
  let s; let text;
  if (state.failed) { s = 'error'; text = '열 수 없음'; }
  else if (!state.loaded) { s = 'loading'; text = '불러오는 중'; }
  else if (state.saving) { s = 'saving'; text = '저장 중…'; }
  else if (state.dirty) { s = 'dirty'; text = '저장 안 됨'; }
  else { s = 'saved'; text = '저장됨'; }
  badge.dataset.state = s;
  badge.textContent = text;
  const ready = state.loaded && !state.failed;
  $('btn-save').disabled = !ready || !!state.saving;
  $('btn-save-as').disabled = !ready || !!state.saving;
  $('btn-print').disabled = !ready || !!state.pdfBusy;
}

function setDirty(dirty) {
  dirty = !!dirty;
  if (dirty === state.dirty) return;
  state.dirty = dirty;
  renderState();
  try { api?.reportDirty(dirty); } catch {}
}

/** 알림 띠. actions: [{ label, primary, onClick }] */
function showBanner(kind, text, actions = []) {
  state.bannerKind = kind;
  const banner = $('banner');
  banner.dataset.kind = kind === 'error' ? 'error' : 'warn';
  $('banner-text').textContent = text;
  const box = $('banner-actions');
  box.replaceChildren(...actions.map((a) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = `btn sm${a.primary ? ' primary' : ''}`;
    b.textContent = a.label;
    b.addEventListener('click', () => a.onClick());
    return b;
  }));
  banner.hidden = false;
}

function hideBanner() {
  state.bannerKind = null;
  $('banner').hidden = true;
}

function toast(text, type = 'success') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  const msg = document.createElement('span');
  msg.textContent = text;
  el.append(icon(type === 'error' ? 'alert' : type === 'info' ? 'info' : 'check'), msg);
  $('toasts').append(el);
  setTimeout(() => { el.classList.add('out'); setTimeout(() => el.remove(), 200); }, 2600);
}

function setLoading(text) {
  $('loading').hidden = !text;
  if (text) $('loading-text').textContent = text;
}

/** 편집기를 쓸 수 없을 때 (이유를 한국어로) */
function fail(reason, detail = '', title = '편집기를 열 수 없습니다') {
  state.failed = reason || title;
  stopDirtyPoll();
  setLoading(null);
  $('failure-title').textContent = title;
  $('failure-reason').textContent = reason || '';
  $('failure-detail').textContent = detail || '';
  $('failure-detail').hidden = !detail;
  $('failure').hidden = false;
  renderState();
  try { api?.reportDirty(false); } catch {}
  log(`실패: ${reason}${detail ? ` (${detail})` : ''}`);
}

// ── 바뀜 확인 ───────────────────────────────────────────
// studio는 사람이 고친 것을 이벤트로 알려 주지 않으므로 상태를 짧게 묻는다 (가벼운 호출)

function refreshDirty() {
  if (!state.editor || !state.loaded || state.failed) return Promise.resolve(state.dirty);
  state.poll ??= (async () => {
    try {
      const ctx = await state.editor.commands.context();
      if (ctx && typeof ctx.isDirty === 'boolean' && !state.saving && !state.reloading) {
        let dirty = ctx.isDirty;
        // 입력 직후에 저장하면 studio가 저장 뒤에 늦게 '바뀜'을 다시 표시한다 → 내용이 저장한 것과 같으면 '저장됨'으로 되돌린다
        if (dirty && !state.lateEdits && state.saved && Date.now() - state.saved.at < RECHECK_MS) dirty = !(await sameAsSaved());
        if (!state.saving && !state.reloading) setDirty(dirty || state.lateEdits);
      }
    } catch {} finally { state.poll = null; }
    return state.dirty;
  })();
  return state.poll;
}

function startDirtyPoll() {
  stopDirtyPoll();
  state.pollTimer = setInterval(() => { void refreshDirty(); }, DIRTY_POLL_MS);
}
function stopDirtyPoll() { clearInterval(state.pollTimer); state.pollTimer = 0; }

// ── 불러오기 · 저장 ─────────────────────────────────────

async function loadInto(doc) {
  const res = await state.editor.loadFile(doc.bytes, doc.name, { skipUnsavedGuard: true, suppressDialogs: true });
  state.loaded = true;
  state.missing = false;
  state.lateEdits = false;
  state.saved = null;
  state.dirty = true; // 아래 setDirty가 꼭 알리도록
  setDirty(false);
  renderState();
  return res;
}

async function reloadFromDisk(message) {
  if (state.reloading || !state.editor) return false;
  state.reloading = true;
  try {
    const doc = unwrap(await api.getDocument());
    setDocInfo(doc);
    await loadInto(doc);
    hideBanner();
    toast(message, 'info');
    return true;
  } catch (err) {
    showBanner('error', `다시 불러오지 못했습니다: ${err.message}`, [{ label: '닫기', onClick: hideBanner }]);
    return false;
  } finally {
    state.reloading = false;
    renderState();
  }
}

async function sha256Hex(bytes) {
  const d = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  return Array.from(d, (b) => b.toString(16).padStart(2, '0')).join('');
}
const sameBytes = (a, b) => a instanceof Uint8Array && a.length === b.length && a.every((x, i) => x === b[i]);

/** 편집기의 지금 내용이 저장한 바이트(format 형식, SHA-256 sha)와 다른지. 내보내기는 같은 내용이면 같은 바이트다 */
async function changedSince(bytes, format, sha) {
  try {
    // 문서 형식 그대로면 studio가 지금 내용의 SHA-256을 준다 (바이트를 옮기지 않아 가볍다)
    const st = await state.editor.getDocumentState();
    if (st?.format === format && typeof st.documentSha256 === 'string') return st.documentSha256 !== sha;
  } catch {}
  const again = format === 'hwpx' ? await state.editor.exportHwpx() : await state.editor.exportHwp();
  return !sameBytes(again, bytes);
}

/** 저장 직후 studio가 늦게 '바뀜'으로 표시했을 때: 지금 내용이 저장한 것과 같으면 다시 '저장됨'으로 → true */
async function sameAsSaved() {
  const saved = state.saved;
  try {
    const st = await state.editor.getDocumentState();
    if (state.saving || st?.format !== saved.format || st.documentSha256 !== saved.sha) return false;
    await state.editor.notifySaved(state.name);
    // notifySaved 사이에 친 글이 있으면 studio는 방금 '저장됨'으로 알게 됐으므로 여기서 기억한다
    const after = await state.editor.getDocumentState();
    if (after?.documentSha256 !== saved.sha) { state.lateEdits = true; return false; }
    return true;
  } catch { return false; }
}

/** 저장. ext를 주면 다른 이름으로 저장(그 형식으로 내보냄) → 성공 여부 */
function save({ ext } = {}) {
  if (!state.loaded || state.failed || !state.editor) return Promise.resolve(false);
  if (state.saving) return state.saving;
  const as = !!ext;
  const format = as ? ext : state.format;
  closeMenu();
  state.saving = (async () => {
    renderState();
    try {
      const bytes = format === 'hwpx' ? await state.editor.exportHwpx() : await state.editor.exportHwp();
      if (!(bytes instanceof Uint8Array) || bytes.length === 0) throw new Error('내보낸 문서가 비어 있습니다');
      const res = unwrap(as ? await api.saveAs(bytes, format) : await api.save(bytes));
      if (res?.canceled) return false;
      if (as) setDocInfo(res);
      // studio의 '바뀜' 표시를 지운다 (실패해도 파일은 이미 저장됨)
      try { await state.editor.notifySaved(state.name); } catch (err) { log(`notifySaved 실패: ${err.message}`); }
      state.missing = false;
      // 내보낸 뒤 ~ notifySaved 사이에 친 글은 파일에 없는데 studio는 '저장됨'으로 안다 → 지금 내용과 저장한 내용을 견준다.
      // notifySaved 뒤에 고친 것은 studio가 스스로 '바뀜'으로 표시하므로 여기서 놓치는 틈이 없다
      let late; let sha = null;
      try {
        sha = await sha256Hex(bytes);
        late = await changedSince(bytes, format, sha);
      } catch (err) {
        log(`저장 뒤 확인 실패: ${err.message}`);
        late = true; // 모르면 '저장 안 됨'으로 (닫을 때 묻게)
      }
      state.lateEdits = late;
      state.saved = sha && !late ? { format, sha, at: Date.now() } : null;
      state.dirty = !late;
      setDirty(late);
      if (state.bannerKind) hideBanner();
      if (late) toast('저장했습니다. 저장하는 동안 고친 내용은 아직 저장되지 않았습니다', 'info');
      else toast(as ? `다른 이름으로 저장했습니다: ${state.name}` : '저장했습니다');
      return true;
    } catch (err) {
      showBanner('error', `저장하지 못했습니다: ${err.message}`, [{ label: '닫기', onClick: hideBanner }]);
      return false;
    } finally {
      state.saving = null;
      renderState();
    }
  })();
  return state.saving;
}

// ── 인쇄 · PDF로 저장 · 자동 저장 ──────────────────────

/** 지금 내용으로 PDF: op 'print' | 'save' (저장 안 한 내용도 들어간다) */
async function makePdf(op) {
  if (!state.loaded || state.failed || !state.editor || state.pdfBusy) return false;
  closeMenu();
  state.pdfBusy = true;
  renderState();
  try {
    const bytes = await state.editor.exportHwp();
    const r = unwrap(await api.pdf(bytes, op));
    if (r?.canceled) return false;
    if (op === 'save') toast(`PDF로 저장했습니다: ${r.name ?? ''}`);
    else if (r?.printed) toast('인쇄를 보냈습니다');
    else if (r?.opened) toast('PDF 앱에서 인쇄하세요', 'info');
    else if (r?.reason && r.reason !== 'cancelled') toast(`인쇄하지 못했습니다: ${r.reason}`, 'error');
    return true;
  } catch (err) {
    showBanner('error', `${op === 'save' ? 'PDF로 저장하지' : '인쇄하지'} 못했습니다: ${err.message}`, [{ label: '닫기', onClick: hideBanner }]);
    return false;
  } finally {
    state.pdfBusy = false;
    renderState();
  }
}

/** 10분마다: 고친 내용이 있으면 복구 사본으로 (원본 형식 그대로, 원본 파일은 건드리지 않음) */
async function autosave() {
  if (!state.loaded || state.failed || !state.editor || state.saving || state.reloading || !state.dirty || state.autosaving) return;
  state.autosaving = true;
  try {
    const bytes = state.format === 'hwpx' ? await state.editor.exportHwpx() : await state.editor.exportHwp();
    unwrap(await api.save(bytes, { reason: 'auto' }));
    log('자동 저장 사본을 남겼습니다');
  } catch (err) { log(`자동 저장 실패: ${err.message}`); } finally { state.autosaving = false; }
}

/** 다시 열었을 때 복구 사본이 파일보다 새로우면 묻는다 */
function offerDraft(draft) {
  const when = new Date(draft.mtimeMs).toLocaleString('ko-KR');
  showBanner('draft', `저장하지 않고 닫힌 편집 내용이 있습니다 (${when} 자동 저장). 복구할까요?`, [
    { label: '복구', primary: true, onClick: async () => {
      try {
        const d = unwrap(await api.readDraft());
        await state.editor.loadFile(d.bytes, state.name, { skipUnsavedGuard: true, suppressDialogs: true });
        state.saved = null;
        state.lateEdits = true; // 복구한 내용은 아직 파일에 없다 → '저장 안 됨'
        setDirty(true);
        hideBanner();
        toast('자동 저장된 내용을 복구했습니다. 저장하면 문서에 들어갑니다', 'info');
      } catch (err) { showBanner('error', `복구하지 못했습니다: ${err.message}`, [{ label: '닫기', onClick: hideBanner }]); }
    } },
    { label: '버리기', onClick: async () => { try { await api.discardDraft(); } catch {} hideBanner(); } },
  ]);
}

// ── 다른 이름으로 저장 메뉴 ─────────────────────────────

function openMenu() {
  const menu = $('save-as-menu');
  menu.hidden = false;
  $('btn-save-as').setAttribute('aria-expanded', 'true');
  menu.querySelector(`[data-ext="${state.format}"]`)?.focus();
}
function closeMenu() {
  $('save-as-menu').hidden = true;
  $('btn-save-as').setAttribute('aria-expanded', 'false');
}

function wireUi() {
  $('btn-save').addEventListener('click', () => { void save(); });
  $('btn-save-as').addEventListener('click', (e) => {
    e.stopPropagation();
    if ($('save-as-menu').hidden) openMenu(); else closeMenu();
  });
  for (const item of document.querySelectorAll('#save-as-menu [data-ext]')) {
    item.addEventListener('click', () => { void save({ ext: item.dataset.ext }); });
  }
  document.querySelector('#save-as-menu [data-pdf]')?.addEventListener('click', () => { void makePdf('save'); });
  $('btn-print').addEventListener('click', () => { void makePdf('print'); });
  document.addEventListener('click', (e) => { if (!e.target.closest?.('.menu-wrap')) closeMenu(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); });
  window.addEventListener('blur', closeMenu);
  $('btn-close').addEventListener('click', () => api?.close());
  $('btn-panel').addEventListener('click', () => api?.panel?.());
  $('btn-failure-close').addEventListener('click', () => api?.close());
}

// ── 앱과 잇기 ───────────────────────────────────────────

function wireApp() {
  // 다른 곳(다른 프로그램·AI)에서 파일이 바뀜
  api.onExternalChange(async (info) => {
    if (!state.loaded || state.failed) return;
    if (info?.missing) {
      state.missing = true;
      showBanner('missing', '문서 파일이 없어졌습니다 (옮겨졌거나 지워짐). 저장하면 같은 자리에 다시 만듭니다.', [{ label: '확인', onClick: hideBanner }]);
      return;
    }
    state.missing = false;
    if (state.saving) await state.saving;
    const dirty = await refreshDirty();
    // 주 창의 '원래대로'는 문구를 따로 (파일 내용이 그대로여도 알려 온다)
    const restore = info?.reason === 'restore';
    if (!dirty) { await reloadFromDisk(restore ? '원래대로 되돌린 내용을 다시 불러왔습니다' : '다른 곳에서 바뀐 내용을 다시 불러왔습니다'); return; }
    showBanner('changed', restore
      ? '주 창에서 문서를 원래대로 되돌렸습니다. 다시 불러오면 여기서 고친 내용은 사라지고, 내 편집을 유지한 채 저장하면 되돌린 내용을 덮어씁니다.'
      : '다른 곳에서 이 문서가 바뀌었습니다. 다시 불러오면 여기서 고친 내용은 사라지고, 내 편집을 유지한 채 저장하면 바깥에서 바뀐 내용을 덮어씁니다.', [
      { label: '다시 불러오기', primary: true, onClick: () => { void reloadFromDisk(restore ? '원래대로 되돌린 내용을 다시 불러왔습니다' : '바뀐 내용을 다시 불러왔습니다'); } },
      { label: '내 편집 유지', onClick: hideBanner },
    ]);
  });

  // 창을 닫기 전: 저장 안 한 변경이 있는지
  api.onRequestClose(async () => {
    if (!state.loaded || state.failed) return false;
    if (state.saving) await state.saving;
    return refreshDirty();
  });

  // 단축키(Ctrl+S, Ctrl+Shift+S)와 닫기 확인창의 '저장'
  api.onCommand(async (command) => {
    if (command === 'save') return save();
    if (command === 'save-as') { if (!state.loaded || state.failed) return false; openMenu(); return true; }
    if (command === 'print') return makePdf('print');
    return false;
  });
}

// ── 시작 ─────────────────────────────────────────────────

async function boot() {
  wireUi();
  renderState();
  if (!api) { fail('편집 창을 앱과 연결하지 못했습니다.'); return; }
  wireApp();

  let doc;
  try { doc = unwrap(await api.getDocument()); } catch (err) { fail(err.message); return; }
  setDocInfo(doc);
  if (!doc.available) { fail(doc.reason || '한글 편집기 파일이 없습니다.'); return; }

  let sdk;
  try { sdk = await import('/sdk/index.js'); } catch (err) {
    fail('한글 편집기 연결 파일(SDK)을 불러오지 못했습니다. Office Kit를 다시 설치하세요.', err?.message);
    return;
  }

  const started = performance.now();
  try {
    state.editor = await sdk.createEditor($('editor'), { studioUrl: doc.studioUrl });
  } catch (err) {
    fail('한글 편집기를 시작하지 못했습니다. 편집기 파일이 빠졌거나 손상됐을 수 있습니다.', err?.message);
    return;
  }
  log(`편집기 준비 ${Math.round(performance.now() - started)}ms`);

  setLoading('문서를 여는 중…');
  try {
    const res = await loadInto(doc);
    log(`문서 열림: ${res?.pageCount ?? '?'}쪽, ${Math.round(performance.now() - started)}ms`);
  } catch (err) {
    fail('이 문서를 편집기로 열지 못했습니다. 손상됐거나 지원하지 않는 내용이 있을 수 있습니다.', err?.message, '문서를 열 수 없습니다');
    return;
  }
  setLoading(null);
  try {
    // 에이전트 명령으로 바뀐 것은 바로 알린다 (사람이 고친 것은 주기적으로 확인)
    state.editor.onDocumentChanged(() => { void refreshDirty(); });
  } catch {}
  startDirtyPoll();
  setInterval(() => { void autosave(); }, AUTOSAVE_MS);
  if (doc.draft) offerDraft(doc.draft);
  try { state.editor.element.focus(); } catch {}
}

if (TEST) {
  window.__editorTest = {
    state: () => ({
      loaded: state.loaded, failed: state.failed, dirty: state.dirty, saving: !!state.saving, reloading: state.reloading,
      name: state.name, ext: state.ext, format: state.format, readOnly: state.readOnly, missing: state.missing, lateEdits: state.lateEdits,
      banner: state.bannerKind, bannerText: state.bannerKind ? $('banner-text').textContent : null,
      badge: $('doc-state').textContent, title: document.title,
    }),
    refreshDirty: () => refreshDirty(),
    save: (ext) => save(ext ? { ext } : {}),
    autosave: () => autosave(),
    makePdf: (op) => makePdf(op),
    pageCount: () => state.editor?.pageCount(),
    pageSvg: (page = 0) => state.editor?.getPageSvg(page),
    context: () => state.editor?.commands.context(),
    execute: (id, params) => state.editor?.commands.execute(id, params),
    focusEditor: () => { state.editor?.element.focus(); return document.activeElement === state.editor?.element; },
    editorRect: () => {
      const r = $('editor').getBoundingClientRect();
      return { x: r.x, y: r.y, width: r.width, height: r.height };
    },
    bannerAction: (label) => {
      const b = [...$('banner-actions').querySelectorAll('button')].find((x) => x.textContent === label);
      b?.click();
      return !!b;
    },
  };
}

boot().catch((err) => fail('편집 창을 준비하지 못했습니다.', err?.message));
