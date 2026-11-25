// home.js — 홈 탭(?view=home: 빈 화면·최근 문서·새 문서·도구·설정)과 문서 탭(?view=doc: 그 문서의 미리보기·문서 작업) 화면.
// 열기는 main이 새 탭으로 연다 (홈은 문서를 품지 않는다). AI 대화(panel.js, 다른 뷰)는 bridge.call('chat.…')로만 부른다 (main 중계).
'use strict';
(() => {
  const {
    K, I18N, $, $$, h, icon, spinner, call, sleep, pathKey,
    extOf, baseName, dirName, KIND_BY_EXT, kindStyle, ENGINE_LABEL, ENGINE_SHORT, whenText, relText, setBusy,
    clearBusy, toast, friendlyError, showError, confirmDialog, state, el: coreEl, bridge, onInit, view,
  } = OKCore;
  // AI 패널 뷰에서는 아무것도 등록하지 않는다 (panel.js 몫)
  if (view === 'panel') return;

  // home 칸 (panel 칸은 읽지도 쓰지도 않는다)
  Object.assign(state, {
    preview: null,
    previewSeq: 0,
    rendering: null,        // 렌더링 중인 문서 경로
    pendingRefresh: null,
    zoom: 'auto',
    smokeSent: false,
  });

  const el = {
    app: coreEl.app, workspace: $('#workspace'), viewer: $('#viewer'),
    empty: $('#empty'), preview: $('#preview'), zoombar: $('#zoombar'), refreshing: $('#refreshing'),
    recent: $('#recent'), recentList: $('#recent-list'),
    docinfo: $('#docinfo'), docKind: $('#doc-kind'), docName: $('#doc-name'), docEngine: $('#doc-engine'),
    stEngine: $('#st-engine'), stPages: $('#st-pages'), stSaved: $('#st-saved'), stEditor: $('#st-editor'),
    btnEdit: $('#btn-edit'),
    overlay: $('#drop-overlay'),
    settings: $('#dlg-settings'), confirm: coreEl.confirm,
  };

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
      const pin = h('button', {
        class: `icon-btn sm recent-pin${r.pinned ? ' on' : ''}`, title: r.pinned ? '고정 풀기' : '목록 맨 위에 고정',
        'aria-label': `${r.name} ${r.pinned ? '고정 풀기' : '고정'}`, 'aria-pressed': String(!!r.pinned),
        onclick: async (e) => { e.stopPropagation(); try { renderRecent(await call(K.recent.pin(r.path, !r.pinned))); } catch (err) { showError(err); } },
      }, icon('pin'));
      return h('li', { class: `recent-item${r.exists ? '' : ' missing'}${r.pinned ? ' pinned' : ''}` }, open, pin, remove);
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
    closeVersions();
    // 열려 있는 문서를 다시 연 경우: 미리보기는 그대로 두고 바뀌었을 때만 새로 그린다
    if (doc && state.doc && pathKey(doc.path) === pathKey(state.doc.path)) {
      state.doc = doc;
      renderTopbar();
      renderStatus();
      refreshPreview({ ifChanged: true, reason: 'reopen' });
      return;
    }
    // 다른 문서가 열리면 개인정보 찾기 결과를 비우고 창을 닫는다
    pii.items = []; pii.path = null;
    try { const pd = $('#dlg-pii'); if (pd?.open) pd.close(); } catch {}
    state.doc = doc;
    state.preview = null;
    state.zoom = 'auto';
    state.previewSeq++;
    state.pendingRefresh = null;
    document.title = doc ? `${doc.name} - Office Kit` : 'Office Kit';
    renderTopbar();
    renderStatus();
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
    // PDF 문서는 이미 PDF다
    const isPdf = !!d && extOf(d.path) === 'pdf';
    $('#btn-pdf').disabled = !d || isPdf;
    $('#btn-pdf').title = isPdf ? '이미 PDF 문서예요' : 'PDF로 저장';
    const restore = $('#btn-restore');
    restore.disabled = !d || !d.backup || !!state.restoring;
    restore.title = d?.backup ? '마지막 변경 전으로 되돌리기 (옆 ▾: 버전 목록)' : '이 문서의 백업이 없어요';
    $('#btn-versions').disabled = restore.disabled;
    if (restore.disabled) closeVersions();
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
      else if (p.kind === 'pdf') pages = !p.pages ? '' : d.kind === 'PowerPoint' ? `슬라이드 ${p.pages}장` : `${p.pages}쪽`;
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
    // PDF: 앱 안에서는 못 고치므로 Word로 열어 고치게 한다 (Word가 PDF를 편집 가능한 문서로 바꾼다)
    const pdf = !!d && extOf(d.path) === 'pdf';
    b.querySelector('.lbl').textContent = pdf ? 'Word로 고치기' : '직접 편집';
    if (pdf) {
      b.hidden = false; b.disabled = false;
      b.classList.remove('active');
      b.title = 'Word에서 PDF를 편집할 수 있는 문서로 바꿔 열어요 (모양이 조금 달라질 수 있어요). 고친 뒤 Word에서 .docx나 PDF로 저장하세요.';
      return;
    }
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
    if (err.code === 'PDF_EDIT') { toast(err.message, { type: 'info', timeout: 9000 }); return null; }
    if (err.code === 'EDITOR_MISSING') { setEditorAvailable(p, false, err.message); renderEditButton(); return showEditorMissing(err.message, p); }
    if (OPEN_IN_APP_CODES.has(err.code)) {
      toast(err.message, { type: 'info', action: { label: '프로그램에서 열기', run: actOpenInApp } });
      return null;
    }
    showError(err, '편집 창을 열지 못했어요');
    return null;
  }

  async function actPdfInWord() {
    const d = state.doc;
    if (!d) return;
    const yes = await confirmDialog({
      title: 'Word로 열어 고칠까요?',
      body: [
        'Word가 PDF를 편집할 수 있는 문서로 바꿔서 열어요. Word가 묻는 확인 창에서 ‘확인’을 눌러 주세요.',
        '표·그림이 많은 PDF는 모양이 조금 달라질 수 있어요. 원본 PDF는 그대로 두고, 고친 뒤 Word에서 ‘다른 이름으로 저장’으로 .docx나 PDF로 저장하세요.',
      ],
      ok: 'Word로 열기',
    });
    if (!yes || state.doc !== d) return;
    setBusy('edit', 'Word 여는 중…');
    try {
      await call(K.doc.openInWord());
      toast('Word에서 열고 있어요. 확인 창이 뜨면 ‘확인’을 눌러 주세요.', { type: 'info', timeout: 8000 });
    } catch (err) { showError(err); } finally { clearBusy('edit'); }
  }

  async function actEdit() {
    if (!state.doc) return;
    if (extOf(state.doc.path) === 'pdf') { actPdfInWord(); return; }
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
    // Office 편집 탭이 많이 열렸을 때의 메모리 안내(r.notice)는 main이 탭 창에 대화상자로 보인다 (이 문서 탭은 편집 탭으로 바뀐다)
  }

  function onDocEditor(ev) {
    if (!ev || !state.doc || pathKey(ev.path) !== pathKey(state.doc.path)) return;
    // aiBusy: main이 AI 대화 시작·끝에 보낸다 (패널이 중계한 편집 상태에는 없으므로 있을 때만 바꾼다)
    state.doc = { ...state.doc, editor: { open: !!ev.open, dirty: !!ev.dirty }, ...(typeof ev.aiBusy === 'boolean' ? { aiBusy: ev.aiBusy } : {}) };
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
    const busyText = had ? '미리보기 새로 고치는 중…' : '미리보기 만드는 중…';
    setBusy('preview', busyText);
    // 프로그램으로 PDF를 만들면 오래 걸린다 (처음 한 번): 몇 초 지나면 무엇을 기다리는지와 걸린 시간을 보여 준다
    const started = Date.now();
    const kind = state.doc.kind;
    const slow = setInterval(() => {
      const sec = Math.round((Date.now() - started) / 1000);
      if (sec < 3 || seq !== state.previewSeq) return;
      const by = { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' }[kind];
      const text = by ? `${by}로 미리보기 PDF를 만드는 중… ${sec}초 (처음 한 번만 오래 걸려요)` : `${busyText.replace(/…$/, '')}… ${sec}초`;
      if (state.busy.has('preview')) setBusy('preview', text);
      const label = $('.sk-label', el.preview);
      if (label) label.replaceChildren(spinner(), text);
    }, 1000);
    let result = null;
    try {
      const res = await call(K.doc.preview({ ifChanged }));
      if (seq !== state.previewSeq || !state.doc || pathKey(state.doc.path) !== pathKey(path)) return;
      if (res.kind === 'stale') return; // 늦게 끝난 다른 문서의 미리보기
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
      clearInterval(slow);
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
    const f = friendlyError(err);
    const card = h('div', { class: 'error-card', role: 'alert' },
      h('div', { class: 'err-icon' }, icon('alert')),
      h('h3', { text: '미리보기를 만들지 못했어요' }),
      h('p', { text: f.text }),
      // 엔진이 낸 원래 글(오류 코드·경로·영어)은 접어 둔다
      f.detail ? h('details', { class: 'err-more' }, h('summary', { text: '자세히' }), h('pre', { class: 'err-detail', text: f.detail })) : null,
      h('div', { class: 'btn-row' },
        // 암호·손상은 다시 해도 같다 → 다시 시도 없이
        f.retry ? h('button', { class: 'btn primary', onclick: () => refreshPreview({ reason: 'manual' }) }, icon('refresh'), '다시 시도') : null,
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
    sheets.forEach((s, i) => tabs.append(h('button', { class: 'sheet-tab', role: 'tab', text: s.name || `시트 ${i + 1}`, 'data-no-i18n': !!s.name, onclick: () => show(i) })));
    view.append(tabs, scroll);
    show(0);
    return view;
  }

  // ── 문서 작업 ────────────────────────────────────────

  // 버전(백업)을 남긴 까닭 → 목록에 보일 말
  // ── 버전·바뀐 내용·되돌리기 (doc-actions.js — 홈 종류 뷰와 AI 패널이 같이 쓴다) ──
  const versionLabel = (v) => OKDocActions.versionLabel(v);

  /** 버전으로 되돌리기. file을 주지 않으면 마지막 변경 전으로 */
  function actRestore(file) {
    if (state.doc) OKDocActions.restore(state.doc.path, file);
  }

  /** file: 버전(백업) 파일 이름. 그 버전 → 지금 사이에 바뀐 글을 보여 준다 */
  function showDiff(file, opts) {
    if (state.doc) OKDocActions.showDiff(state.doc.path, file, opts);
  }

  // 되돌리는 동안·끝난 뒤 상단 막대를 맞춘다 (끝나면 main이 돌려준 문서 정보로)
  document.addEventListener('okdoc:restore', (e) => {
    const { path, phase, doc } = e.detail ?? {};
    if (!state.doc || pathKey(path) !== pathKey(state.doc.path)) return;
    if (phase === 'end' && doc) updateDoc(doc);
    else renderTopbar();
  });

  // ── 도구 메뉴 ────────────────────────────────────────

  function closeTools() {
    const pop = $('#tools-pop');
    if (pop.hidden) return;
    pop.hidden = true;
    $('#btn-tools').setAttribute('aria-expanded', 'false');
  }

  function toggleTools() {
    const pop = $('#tools-pop');
    const btn = $('#btn-tools');
    if (!pop.hidden) { closeTools(); return; }
    closeVersions();
    $('#tool-pii').disabled = !state.doc;
    $('#tool-pii').title = state.doc ? '' : '문서를 먼저 열어 주세요';
    pop.hidden = false;
    const r = btn.getBoundingClientRect();
    pop.style.top = `${Math.round(r.bottom + 6)}px`;
    pop.style.left = `${Math.round(Math.max(12, Math.min(r.right - pop.offsetWidth, window.innerWidth - pop.offsetWidth - 12)))}px`;
    btn.setAttribute('aria-expanded', 'true');
    ($('#tools-pop .tool-item:not(:disabled)'))?.focus();
  }

  // ── 개인정보 찾아 가리기 ─────────────────────────────

  const pii = { items: [], busy: false, path: null };

  async function openPii() {
    closeTools();
    if (!state.doc) { toast('문서를 먼저 열어 주세요.', { type: 'info' }); return; }
    const d0 = state.doc;
    const dlg = $('#dlg-pii');
    const body = $('#pii-body');
    const btn = $('#pii-mask');
    $('#pii-sub').textContent = d0.name;
    $('#pii-stats').replaceChildren();
    $('#pii-note').textContent = '원본은 그대로 두고 가린 사본을 만들어요. 그림 속 글자는 찾지 못해요.';
    btn.hidden = false;
    btn.disabled = true;
    body.replaceChildren(h('div', { class: 'diff-msg' }, spinner(), ' 찾는 중…'));
    if (!dlg.open) dlg.showModal();
    let r;
    try { r = await call(K.pii.scan()); } catch (err) {
      body.replaceChildren(h('div', { class: 'diff-msg' }, h('div', { class: 'inline-error', text: err.message })));
      return;
    }
    if (!dlg.open || !state.doc || pathKey(state.doc.path) !== pathKey(d0.path)) return;
    pii.path = r.path; // 찾은 문서 — 가리기 요청에 함께 보내 같은 문서에만 적용
    pii.items = r.items.map((it) => ({ ...it, on: true }));
    if (!pii.items.length) {
      body.replaceChildren(h('div', { class: 'diff-msg', text: `개인정보를 찾지 못했어요 (글자 ${r.chars.toLocaleString()}자 확인). 그림 속 글자나 띄어 쓴 번호는 찾지 못하니 직접 확인해 주세요.` }));
      return;
    }
    const counts = {};
    for (const it of pii.items) counts[it.label] = (counts[it.label] ?? 0) + it.count;
    $('#pii-stats').replaceChildren(...Object.entries(counts).map(([k, n]) => h('span', { class: 'pill warn', text: `${k} ${n}` })));
    $('#pii-note').textContent = `가린 사본: ${baseName(r.output)} (원본 옆)`;
    const pdfOnly = extOf(d0.path) === 'pdf';
    if (pdfOnly) $('#pii-note').textContent = 'PDF는 글자를 바꿀 수 없어 찾기만 돼요. 가린 사본은 한글·Word·PowerPoint·Excel 문서에서 만들 수 있어요.';
    const all = h('input', { type: 'checkbox', checked: true, 'aria-label': '모두 고르기' });
    const boxes = pii.items.map((it) => {
      const cb = h('input', { type: 'checkbox', checked: true, 'aria-label': `${it.label} ${it.value} 가리기` });
      cb.addEventListener('change', () => { it.on = cb.checked; sync(); });
      return cb;
    });
    const sync = () => {
      const n = pii.items.filter((x) => x.on).length;
      all.checked = n === pii.items.length;
      all.indeterminate = n > 0 && n < pii.items.length;
      btn.disabled = !n || pii.busy || pdfOnly;
      btn.lastChild.textContent = n && !pdfOnly ? `가린 사본 만들기 (${n}가지)` : '가린 사본 만들기';
    };
    all.addEventListener('change', () => { pii.items.forEach((it, i) => { it.on = all.checked; boxes[i].checked = all.checked; }); sync(); });
    body.replaceChildren(h('table', { class: 'diff-table' },
      h('thead', {}, h('tr', {}, h('th', {}, all), h('th', { text: '종류' }), h('th', { text: '찾은 값' }), h('th', { text: '가린 모양' }), h('th', { text: '앞뒤 글' }))),
      h('tbody', {}, ...pii.items.map((it, i) => h('tr', { class: 'pii-row' },
        h('td', {}, boxes[i]),
        h('td', { text: it.count > 1 ? `${it.label} ×${it.count}` : it.label }),
        h('td', { class: 'val', text: it.value }),
        h('td', { class: 'masked', text: it.masked }),
        h('td', { class: 'ctx', text: it.context }))))));
    sync();
  }

  async function actPiiMask() {
    const values = pii.items.filter((x) => x.on).map((x) => x.value);
    if (!values.length || pii.busy) return;
    const btn = $('#pii-mask');
    pii.busy = true;
    btn.disabled = true;
    setBusy('pii', '가린 사본 만드는 중…');
    let r;
    try { r = await call(K.pii.mask({ values, path: pii.path })); } catch (err) { showError(err); return; } finally {
      pii.busy = false;
      clearBusy('pii');
      btn.disabled = false;
    }
    btn.hidden = true;
    const left = r.remaining.length;
    $('#pii-stats').replaceChildren(h('span', { class: `pill ${left ? 'warn' : 'on'}`, text: left ? `남은 것 ${left}가지` : '모두 가림' }));
    $('#pii-body').replaceChildren(h('div', { class: 'pii-result' },
      h('p', { text: `가린 사본을 만들었어요: ${baseName(r.output)}` }),
      h('p', { class: 'help', text: `${r.items.reduce((n, it) => n + it.count, 0)}곳을 가렸어요. 원본은 그대로예요.` }),
      left ? h('div', { class: 'warn-note' }, icon('alert'), h('span', { text: `가리지 못하고 남은 것: ${r.remaining.map((x) => x.value).join(', ')} — 글자가 여러 조각으로 나뉘어 저장된 경우예요. 사본을 열어 직접 고쳐 주세요.` })) : null,
      h('div', { class: 'btn-row' },
        h('button', { class: 'btn primary', text: '사본 열기', onclick: () => { $('#dlg-pii').close(); openPath(r.output); } }),
        h('button', { class: 'btn', text: '폴더에서 보기', onclick: () => call(K.shell.showItem(r.output)).catch(showError) }))));
    $('#pii-note').textContent = '사본을 열어 가려진 모양을 꼭 확인해 주세요.';
  }

  // ── PDF 도구 ─────────────────────────────────────────

  const pdft = { action: 'merge', files: [], busy: false };
  const PDF_ACTION_HINT = {
    merge: '위에서부터 차례로 하나의 PDF로 합쳐요. ↑↓로 순서를 바꿀 수 있어요.',
    extract: '고른 쪽만 새 PDF로 저장해요 (첫 번째 PDF). 예: 1-3, 5',
    delete: '고른 쪽을 뺀 나머지를 새 PDF로 저장해요 (첫 번째 PDF).',
    rotate: '고른 쪽을 돌려 새 PDF로 저장해요. 쪽을 비우면 모든 쪽.',
  };

  function openPdfTools() {
    closeTools();
    // 열린 문서가 PDF면 미리 넣어 둔다
    if (!pdft.files.length && state.doc && extOf(state.doc.path) === 'pdf') {
      pdft.files.push({ path: state.doc.path, name: state.doc.name, pages: null, error: null });
    }
    renderPdfTools();
    const dlg = $('#dlg-pdf');
    if (!dlg.open) dlg.showModal();
  }

  function renderPdfTools() {
    const a = pdft.action;
    $$('#pdf-action button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.action === a)));
    const single = a !== 'merge';
    $('#pdf-list').replaceChildren(...pdft.files.map((f, i) => h('li', { class: 'pdf-item' },
      h('span', { class: 'pdf-name', text: `${i + 1}. ${f.name}`, title: f.path }),
      f.error ? h('span', { class: 'pdf-err', text: f.error }) : h('span', { class: 'pdf-pages', text: f.pages ? `${f.pages}쪽` : '' }),
      single && i > 0 ? h('span', { class: 'pdf-pages', text: '(쓰지 않음)' }) : null,
      h('button', { class: 'icon-btn sm', title: '위로', 'aria-label': '위로', disabled: i === 0 || pdft.busy, onclick: () => { pdft.files.splice(i - 1, 0, pdft.files.splice(i, 1)[0]); renderPdfTools(); } }, icon('chevron-down', 'i flip')),
      h('button', { class: 'icon-btn sm', title: '아래로', 'aria-label': '아래로', disabled: i === pdft.files.length - 1 || pdft.busy, onclick: () => { pdft.files.splice(i + 1, 0, pdft.files.splice(i, 1)[0]); renderPdfTools(); } }, icon('chevron-down')),
      h('button', { class: 'icon-btn sm', title: '빼기', 'aria-label': `${f.name} 빼기`, disabled: pdft.busy, onclick: () => { pdft.files.splice(i, 1); renderPdfTools(); } }, icon('x')))));
    $('#pdf-hint').textContent = PDF_ACTION_HINT[a];
    $('#pdf-opts').hidden = a === 'merge';
    $('#pdf-pages-label').textContent = a === 'rotate' ? '쪽 (비우면 모두)' : '쪽';
    $('#pdf-degrees').hidden = a !== 'rotate';
    const ok = pdft.files.filter((f) => !f.error);
    $('#pdf-run').disabled = pdft.busy || (a === 'merge' ? ok.length < 2 : !ok.length);
  }

  async function actPdfRun() {
    if (pdft.busy) return;
    const files = pdft.files.filter((f) => !f.error).map((f) => f.path);
    pdft.busy = true;
    renderPdfTools();
    setBusy('pdftools', 'PDF 만드는 중…');
    try {
      const r = await call(K.pdftools.run({ action: pdft.action, inputs: pdft.action === 'merge' ? files : files.slice(0, 1), pages: $('#pdf-pages').value, degrees: $('#pdf-degrees').value }));
      if (r?.canceled) return;
      $('#pdf-note').textContent = `저장했어요: ${baseName(r.output)} (${r.pages}쪽)`;
      toast(`PDF를 저장했어요: ${baseName(r.output)} (${r.pages}쪽)`, { type: 'success', action: { label: '열기', run: () => { $('#dlg-pdf').close(); openPath(r.output); } } });
    } catch (err) { showError(err); } finally {
      pdft.busy = false;
      clearBusy('pdftools');
      renderPdfTools();
    }
  }

  // ── 한글 → HWPX 일괄 변환 ────────────────────────────

  const batch = { sources: [], outDir: '', running: false, done: false, scanSeq: 0, count: 0 };
  const BATCH_PILL = { ok: ['같음', 'on'], check: ['확인 필요', 'warn'], skipped: ['건너뜀', 'off'], fail: ['실패', 'err'] };

  function openBatch() {
    closeTools();
    const dlg = $('#dlg-batch');
    if (!batch.running && batch.done) resetBatch();
    renderBatchSetup();
    if (!dlg.open) dlg.showModal();
  }

  function resetBatch() {
    batch.done = false;
    $('#batch-setup').hidden = false;
    $('#batch-run').hidden = true;
    $('#batch-rows').replaceChildren();
    $('#batch-bar').style.width = '0';
    $('#batch-stats').replaceChildren();
    $('#batch-report').hidden = true;
    $('#batch-start').hidden = false;
    $('#batch-cancel').textContent = '닫기';
  }

  function batchAdd(paths) {
    const have = new Set(batch.sources.map(pathKey));
    for (const p of paths) if (p && !have.has(pathKey(p))) { batch.sources.push(p); have.add(pathKey(p)); }
    renderBatchSetup();
  }

  async function renderBatchSetup() {
    $('#batch-sources').replaceChildren(...batch.sources.map((p) => h('span', { class: 'batch-src', title: p },
      h('span', { text: p }),
      h('button', { class: 'icon-btn sm', 'aria-label': `${baseName(p)} 빼기`, onclick: () => { batch.sources = batch.sources.filter((x) => x !== p); renderBatchSetup(); } }, icon('x')))));
    const outDir = $('input[name="batch-out"][value="dir"]').checked;
    $('#batch-outdir-desc').textContent = batch.outDir ? batch.outDir : '폴더 구조는 그대로 두고 한곳에 모아요.';
    const start = $('#batch-start');
    if (!batch.sources.length) { start.disabled = true; start.lastChild.textContent = '변환 시작'; $('#batch-note').textContent = ''; return; }
    const seq = ++batch.scanSeq;
    $('#batch-note').textContent = '파일 세는 중…';
    let r;
    try { r = await call(K.batch.scan({ inputs: batch.sources, recursive: $('#batch-recursive').checked })); } catch (err) { $('#batch-note').textContent = err.message; return; }
    if (seq !== batch.scanSeq) return;
    batch.count = r.count;
    $('#batch-note').textContent = r.count ? `HWP 파일 ${r.count.toLocaleString()}개${r.truncated ? ' (너무 많아 앞부분만)' : ''}` : 'HWP 파일이 없어요';
    start.disabled = !r.count || (outDir && !batch.outDir);
    start.lastChild.textContent = r.count ? `변환 시작 (${r.count.toLocaleString()}개)` : '변환 시작';
  }

  async function actBatchStart() {
    if (batch.running || !batch.count) return;
    const outDir = $('input[name="batch-out"][value="dir"]').checked ? batch.outDir : '';
    batch.running = true;
    $('#batch-setup').hidden = true;
    $('#batch-run').hidden = false;
    $('#batch-rows').replaceChildren();
    $('#batch-bar').style.width = '0';
    $('#batch-start').hidden = true;
    $('#batch-cancel').textContent = '멈추기';
    $('#batch-note').textContent = '바꾸는 중…';
    setBusy('batch', 'HWPX로 바꾸는 중…');
    let r = null;
    try {
      r = await call(K.batch.start({ inputs: batch.sources, recursive: $('#batch-recursive').checked, overwrite: $('#batch-overwrite').checked, outDir }));
    } catch (err) { showError(err); } finally {
      batch.running = false;
      batch.done = true;
      clearBusy('batch');
      $('#batch-cancel').textContent = '닫기';
    }
    if (!r) { $('#batch-note').textContent = ''; return; }
    const s = r.summary;
    $('#batch-bar').style.width = '100%';
    $('#batch-stats').replaceChildren(...['ok', 'check', 'skipped', 'fail'].filter((k) => s[k]).map((k) => h('span', { class: `pill ${BATCH_PILL[k][1]}`, text: `${BATCH_PILL[k][0]} ${s[k]}` })));
    $('#batch-note').textContent = s.cancelled ? `멈췄어요 (${r.results.length}/${s.total}개 처리)` : `${s.total}개 끝났어요${s.check ? ' · ‘확인 필요’는 열어서 원본과 견줘 보세요' : ''}`;
    $('#batch-report').hidden = !r.results.length;
  }

  function onBatchProgress(ev) {
    if (!ev?.result) return;
    const x = ev.result;
    const [txt, cls] = BATCH_PILL[x.status] ?? [x.status, 'off'];
    const open = x.status === 'ok' || x.status === 'check'
      ? h('button', { class: 'link-btn', text: baseName(x.out), title: `${x.out} 열기`, onclick: () => { $('#dlg-batch').close(); openPath(x.out); } })
      : h('span', { text: baseName(x.src), title: x.src });
    const dir = x.src.slice(0, x.src.length - baseName(x.src).length).replace(/[\\/]$/, '');
    $('#batch-rows').append(h('tr', {},
      h('td', {}, h('span', { class: `pill ${cls}`, text: txt })),
      h('td', { class: 'file', title: x.src }, open, h('span', { class: 'dir', text: dir })),
      h('td', { class: 'memo', text: (x.notes ?? []).join(', ') || (x.pages ? `${x.pages[1]}쪽 · 표 ${x.tables?.[1] ?? 0}개` : '') })));
    $('#batch-bar').style.width = `${Math.round(((ev.index + 1) / ev.total) * 100)}%`;
    // 끝난 뒤 늦게 온 진행 알림이 결과 문구를 덮지 않게
    if (batch.running) $('#batch-note').textContent = `${ev.index + 1} / ${ev.total}`;
    const rows = $('#batch-rows').parentElement.parentElement;
    rows.scrollTop = rows.scrollHeight;
  }

  async function actBatchCancel() {
    if (batch.running) { await call(K.batch.cancel()).catch(() => {}); $('#batch-note').textContent = '지금 파일까지 하고 멈춰요…'; return; }
    $('#dlg-batch').close();
  }

  // ── 버전 목록 (원래대로 ▾, doc-actions.js) ───────────

  function closeVersions() { OKDocActions.closeVersions(); }

  function toggleVersions() {
    if (state.doc) OKDocActions.openVersions(state.doc.path);
  }

  async function actOpenInApp() {
    if (!state.doc) return;
    setBusy('openapp', '프로그램 여는 중…');
    try {
      const r = await call(K.doc.openInApp());
      if (r?.readOnly) {
        toast(r.readOnlyReason === 'writePassword'
          ? '쓰기 암호가 걸린 문서라 읽기 전용으로 열었어요. 고치려면 그 프로그램에서 문서를 닫고 다시 열어 쓰기 암호를 넣으세요.'
          : '다른 프로그램이 이 문서를 쓰고 있거나 읽기 전용 파일이라 읽기 전용으로 열었어요.', { type: 'warn' });
      } else {
        toast('프로그램에서 열었어요. 저장하면 여기 미리보기도 자동으로 바뀌어요.', { type: 'success' });
      }
    } catch (err) { showError(err); } finally { clearBusy('openapp'); }
  }

  async function actSavePdf() {
    if (!state.doc) return;
    try {
      // 저장 대화상자가 떠 있는 동안에는 'PDF 만드는 중'을 보이지 않는다 (자리를 고른 뒤에 만든다)
      const pick = await call(K.doc.savePdf({ step: 'pick' }));
      if (pick?.canceled) return;
      setBusy('pdf', 'PDF 만드는 중…');
      const r = await call(K.doc.savePdf({ step: 'make' }));
      // 한컴오피스 없이 내장 엔진(rhwp)으로 만든 한글 PDF는 그렇다고 알린다 (한컴과 글꼴·줄바꿈이 조금 다를 수 있음)
      toast(`PDF로 저장했어요${r.note ? ' (내장 엔진)' : ''}: ${baseName(r.output)}`, { type: 'success', action: { label: '폴더에서 보기', run: () => call(K.shell.showItem(r.output)).catch(showError) } });
    } catch (err) { showError(err); } finally { clearBusy('pdf'); }
  }

  async function actPrint() {
    if (!state.doc) { toast('문서를 먼저 열어 주세요.', { type: 'info' }); return; }
    setBusy('print', '인쇄 준비 중…');
    try {
      const r = await call(K.doc.print());
      if (r?.opened) toast('PDF 프로그램에서 인쇄해 주세요.', { type: 'info' });
      else if (r?.printed) toast('인쇄를 보냈어요', { type: 'success' });
      else if (r?.reason && r.reason !== 'cancelled') toast(`인쇄하지 못했어요: ${r.reason}`, { type: 'warn' });
    } catch (err) { showError(err); } finally { clearBusy('print'); }
  }

  async function actShowFolder() {
    try { await call(K.doc.showInFolder()); } catch (err) { showError(err); }
  }

  async function actCloseDoc() {
    try { await call(K.doc.close()); } catch (err) { showError(err); }
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
  const nd = { format: 'hwpx', start: 'blank', edit: true, folder: '', autoName: '', busy: false, opening: false, seq: 0, timer: 0, templates: {}, template: null };
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
    nd.start = start === 'ai' || start === 'template' ? start : 'blank';
    nd.opening = true;
    try {
      const info = await call(K.doc.newInfo({ format: nd.format }));
      nd.folder = info.folder;
      nd.autoName = info.name;
      nd.templates = info.templates ?? {};
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
    // 서식 파일: 이 형식의 서식 (없으면 고를 수 없음)
    const tpls = nd.templates[nd.format] ?? [];
    if (!tpls.length && nd.start === 'template') nd.start = 'blank';
    if (!tpls.some((t) => t.id === nd.template)) nd.template = tpls[0]?.id ?? null;
    const tplBox = $('#new-template-box');
    tplBox.hidden = nd.start !== 'template';
    tplBox.replaceChildren(...tpls.map((t) => h('button', {
      class: 'tpl-item', type: 'button', role: 'radio', 'aria-checked': String(t.id === nd.template), disabled: busy,
      onclick: () => { nd.template = t.id; renderNewDoc(); },
    }, h('span', { class: 'tpl-name', text: t.name }), h('span', { class: 'tpl-desc', text: t.desc }))));
    const f = NEW_FORMATS[nd.format];
    for (const r of $$('input[name="new-format"]', ne.dlg)) { r.checked = r.value === nd.format; r.disabled = busy; }
    for (const r of $$('input[name="new-start"]', ne.dlg)) { r.checked = r.value === nd.start; r.disabled = busy || (r.value === 'ai' && !ai.ok) || (r.value === 'template' && !tpls.length); }
    $('#new-start-tpl').classList.toggle('disabled', !tpls.length);
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
    bridge.call('chat.forgetIdle', { path: check.path });
    // AI로 시작하면 main이 그 탭 창 AI 패널에 'start' 이벤트를 먼저 보내고 패널이 대화를 만든다
    try {
      const edit = newDocEditable(nd.format) && nd.edit;
      const res = await call(K.doc.create({ format: nd.format, folder: nd.folder, name: ne.name.value, start: nd.start, request: ai ? request : undefined, template: nd.start === 'template' ? nd.template : undefined, edit }));
      nd.busy = false;
      ne.dlg.close();
      const path = res.doc?.path ?? check.path;
      if (ai && !res.chat?.chatId) toast(`${baseName(path)}을(를) 만들었지만 AI를 시작하지 못했어요`, { type: 'warn' });
      else toast(ai ? `${baseName(path)}을(를) 만들고 AI에게 부탁했어요` : `새 문서를 만들었어요: ${baseName(path)}`, { type: 'success' });
      // 편집 창은 main이 열었다 (못 열었으면 이유만 알린다. 문서는 만들어졌음)
      // (메모리 안내는 main이 탭 창에 대화상자로 보인다)
      if (res.editorError) editorError(res.editorError, path);
    } catch (err) {
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
    if (tab === 'open') return loadAssoc();
    return Promise.resolve();
  }

  function renderSettingsView() {
    const s = state.settings;
    $$('input[name="previewMode"]', el.settings).forEach((r) => { r.checked = r.value === (s.previewMode ?? 'auto'); });
    $$('input[name="openAction"]', el.settings).forEach((r) => { r.checked = r.value === (s.openAction ?? 'edit'); });
    $$('#theme-seg button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.theme === (s.theme ?? 'system'))));
    $$('#lang-seg button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.lang === (s.lang ?? 'system'))));
    // 언어는 다시 시작해야 바뀐다 (시작할 때의 설정과 다르면 다시 시작 단추)
    $('#lang-restart').hidden = (s.lang ?? 'system') === I18N.setting;
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

  /** AI 목록 받기 (새 문서 창의 'AI에게 부탁하기'·설정의 AI 고르기). AI 패널은 따로 받는다 */
  async function loadProviders(refresh = false) {
    try {
      state.providers = await call(K.ai.providers({ refresh }));
      state.providerError = null;
    } catch (err) {
      state.providers = null;
      state.providerError = err.message;
    }
    if (ne.dlg.open) renderNewDoc();
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
        if (await saveSettings({ provider: it.id })) {
          await loadProviders(false);
          // 모든 탭 창의 AI 패널도 다시 받게 한다
          bridge.call('chat.loadProviders', false);
          renderProviderChoices(false);
        }
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

  // ── 파일 열기 · 기본 프로그램 ─────────────────────────

  async function loadAssoc() {
    const list = $('#assoc-list');
    const btn = $('#btn-assoc-settings');
    list.replaceChildren(h('div', { class: 'assoc-row' }, h('div', { class: 'loading-line' }, spinner(), '확인하는 중…')));
    let st;
    try { st = await call(K.assoc.status()); } catch (err) {
      list.replaceChildren(h('div', { class: 'assoc-row' }, h('div', { class: 'inline-error', text: err.message })));
      return;
    }
    btn.disabled = !st.supported || !st.packaged;
    btn.title = st.reason ?? '';
    if (!st.supported) { list.replaceChildren(h('div', { class: 'assoc-row' }, h('div', { class: 'help', text: st.reason }))); return; }
    const rows = st.groups.map((g) => h('div', { class: 'assoc-row' },
      h('span', { class: 'assoc-kind', text: g.kind }),
      h('span', { class: 'assoc-types' }, ...g.types.map((t) => h('span', {
        class: `assoc-type${t.isDefault ? ' on' : ''}`,
        title: t.isDefault ? `${t.desc}(.${t.ext})를 Office Kit로 열어요` : `${t.desc}(.${t.ext})는 다른 프로그램으로 열려요`,
      }, t.isDefault ? icon('check') : null, `.${t.ext}`)))));
    const n = st.groups.reduce((a, g) => a + g.types.filter((t) => t.isDefault).length, 0);
    const head = h('div', { class: 'assoc-row' },
      h('span', { class: 'assoc-kind', text: '지금' }),
      h('span', { class: 'int-desc', text: !st.packaged ? st.reason
        : !st.registered || !st.matches ? '기본 프로그램 후보 등록이 빠져 있어요. 아래 단추를 누르면 다시 등록하고 Windows 설정을 열어요.'
          : n ? `초록색 형식은 Office Kit로 열려요 (${n}개).` : '아직 Office Kit로 여는 형식이 없어요.' }));
    list.replaceChildren(head, ...rows);
  }

  async function actAssocSettings() {
    const btn = $('#btn-assoc-settings');
    btn.disabled = true;
    try {
      await call(K.assoc.openSettings());
      toast('Windows 설정에서 바꿀 형식을 눌러 Office Kit를 고른 뒤, 여기서 ‘다시 확인’을 눌러 주세요.', { type: 'info', timeout: 10000 });
    } catch (err) { showError(err); } finally { btn.disabled = false; }
    loadAssoc();
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
        '등록하면 한글 프로그램으로 원본과 똑같은 PDF 저장과 정확한 미리보기를 할 수 있어요. 문서 내용은 어디로도 보내지 않습니다.',
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
    // PDF 미리보기(iframe)는 그 위로 들어온 끌기를 혼자 받아 버린다 (창에 drop이 오지 않음).
    // 끌기는 다른 창(탐색기)에서 시작하므로: 이 창이 포커스를 잃거나 마우스가 창 밖으로 나가면 iframe이 마우스를 받지 않게 했다가
    // (끌기가 어디로 들어와도 이 창이 받는다), 창 안에서 마우스가 움직이면 되돌린다 (끌기 중에는 mousemove가 오지 않는다)
    let passing = false;
    const passThrough = (on) => { if (passing !== on) { passing = on; document.body.classList.toggle('drop-ready', on); } };
    window.addEventListener('blur', () => passThrough(true));
    document.addEventListener('mouseout', (e) => { if (!e.relatedTarget) passThrough(true); });
    window.addEventListener('mousemove', () => passThrough(false), { passive: true });
    window.addEventListener('focus', () => { if (!document.body.classList.contains('dragging')) setTimeout(() => passThrough(false), 0); });
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
      // 일괄 변환 창이 열려 있으면 파일·폴더를 변환 목록에 넣는다
      if ($('#dlg-batch').open) { if (!batch.running) batchAdd(paths); return; }
      if ($('#dlg-pdf').open) {
        const have = new Set(pdft.files.map((f) => pathKey(f.path)));
        for (const p of paths) if (extOf(p) === 'pdf' && !have.has(pathKey(p))) pdft.files.push({ path: p, name: baseName(p), pages: null, error: null });
        renderPdfTools();
        return;
      }
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
    if (state.smoke.newdoc) {
      // AI 선택지가 제 모양으로 보이게 AI 확인을 기다린다
      await Promise.race([state.providersReady, sleep(8000)]);
      await openNewDoc({ start: state.smoke.newdoc === '1' ? 'blank' : state.smoke.newdoc });
      if (state.smoke.newdoc === 'ai') ne.request.value = '주간 회의록 양식 만들어줘';
      await sleep(300);
    }
    if (state.smoke.settings) {
      try { await openSettings(state.smoke.settings); } catch {}
      await sleep(300);
    }
    if (state.smoke.tool === 'pii' || state.smoke.tool === 'pii-mask') {
      await openPii();
      if (state.smoke.tool === 'pii-mask') await actPiiMask();
      await sleep(300);
    }
    if (state.smoke.tool === 'print') { await actPrint(); await sleep(300); }
    if (state.smoke.tool === 'pdftools') {
      openPdfTools();
      if (state.smoke.batch) { pdft.files.push({ path: state.smoke.batch, name: baseName(state.smoke.batch), pages: null, error: null }); }
      renderPdfTools();
      await sleep(300);
    }
    if (state.smoke.tool === 'batch') {
      openBatch();
      if (state.smoke.batch) {
        batchAdd([state.smoke.batch]);
        for (let i = 0; i < 50 && !batch.count; i++) await sleep(100);
        await actBatchStart();
      }
      await sleep(300);
    }
    if (state.smoke.diff) {
      try {
        const v = await call(K.doc.versions({ path: state.doc?.path }));
        const first = v.versions.find((x) => x.reason === 'open');
        if (first) await showDiff(first.file, { label: versionLabel(first), time: first.time });
      } catch (err) { console.error('[smoke] diff', err?.message ?? err); }
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
    $('#btn-restore').addEventListener('click', () => actRestore());
    $('#btn-versions').addEventListener('click', (e) => { e.stopPropagation(); toggleVersions(); });
    $('#ver-open-backups').addEventListener('click', () => { closeVersions(); call(K.shell.openBackups()).catch(showError); });
    // 바깥을 누르거나 Esc → 닫기
    document.addEventListener('mousedown', (e) => {
      if (!$('#versions-pop').hidden && !e.target.closest('#versions-pop, #btn-versions')) closeVersions();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !$('#versions-pop').hidden) { e.preventDefault(); closeVersions(); $('#btn-versions').focus(); }
    });
    window.addEventListener('resize', closeVersions);
    el.btnEdit.addEventListener('click', actEdit);
    el.stEditor.addEventListener('click', actEdit);
    $('#btn-openapp').addEventListener('click', actOpenInApp);
    $('#btn-pdf').addEventListener('click', actSavePdf);
    $('#btn-folder').addEventListener('click', actShowFolder);
    $('#btn-close-doc').addEventListener('click', actCloseDoc);
    $('#btn-settings').addEventListener('click', () => openSettings('ai'));
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

    // 설정 창
    $$('.nav-item', el.settings).forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.tab)));
    $('#settings-close').addEventListener('click', () => el.settings.close());
    el.settings.addEventListener('click', (e) => { if (e.target === el.settings) el.settings.close(); });
    $$('input[name="openAction"]', el.settings).forEach((r) => r.addEventListener('change', () => { if (r.checked) saveSettings({ openAction: r.value }); }));
    $('#btn-assoc-settings').addEventListener('click', actAssocSettings);
    $('#diff-close').addEventListener('click', () => $('#dlg-diff').close());
    $('#btn-tools').addEventListener('click', (e) => { e.stopPropagation(); toggleTools(); });
    $('#tool-pii').addEventListener('click', openPii);
    $('#tool-batch').addEventListener('click', openBatch);
    $('#tool-pdf').addEventListener('click', openPdfTools);
    $$('#pdf-action button').forEach((b) => b.addEventListener('click', () => { pdft.action = b.dataset.action; renderPdfTools(); }));
    $('#pdf-add').addEventListener('click', async () => {
      try {
        const add = await call(K.pdftools.pick());
        const have = new Set(pdft.files.map((f) => pathKey(f.path)));
        for (const f of add) if (!have.has(pathKey(f.path))) pdft.files.push(f);
        renderPdfTools();
      } catch (err) { showError(err); }
    });
    $('#pdf-run').addEventListener('click', actPdfRun);
    $('#pdf-close').addEventListener('click', () => $('#dlg-pdf').close());
    $('#pdf-cancel').addEventListener('click', () => $('#dlg-pdf').close());
    document.addEventListener('click', (e) => { if (!$('#tools-pop').hidden && !e.target.closest('#tools-pop, #btn-tools')) closeTools(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('#tools-pop').hidden) { e.preventDefault(); closeTools(); $('#btn-tools').focus(); } });
    $('#pii-close').addEventListener('click', () => $('#dlg-pii').close());
    $('#pii-cancel').addEventListener('click', () => $('#dlg-pii').close());
    $('#pii-mask').addEventListener('click', actPiiMask);
    $('#batch-close').addEventListener('click', () => { if (!batch.running) $('#dlg-batch').close(); });
    $('#dlg-batch').addEventListener('cancel', (e) => { if (batch.running) e.preventDefault(); });
    $('#batch-cancel').addEventListener('click', actBatchCancel);
    $('#batch-start').addEventListener('click', actBatchStart);
    $('#batch-report').addEventListener('click', async () => {
      try { const r = await call(K.batch.saveReport()); if (r?.output) toast(`보고서를 저장했어요: ${baseName(r.output)}`, { type: 'success', action: { label: '폴더에서 보기', run: () => call(K.shell.showItem(r.output)).catch(showError) } }); } catch (err) { showError(err); }
    });
    $('#batch-add-files').addEventListener('click', async () => { try { batchAdd(await call(K.batch.pick({ kind: 'files' }))); } catch (err) { showError(err); } });
    $('#batch-add-folder').addEventListener('click', async () => { try { batchAdd(await call(K.batch.pick({ kind: 'folder' }))); } catch (err) { showError(err); } });
    $('#batch-recursive').addEventListener('change', () => renderBatchSetup());
    $$('input[name="batch-out"]').forEach((r) => r.addEventListener('change', async () => {
      if (r.value === 'dir' && r.checked && !batch.outDir) {
        try { batch.outDir = (await call(K.batch.pickOut())) ?? ''; } catch (err) { showError(err); }
        if (!batch.outDir) $('input[name="batch-out"][value="same"]').checked = true;
      }
      renderBatchSetup();
    }));
    $('#diff-cancel').addEventListener('click', () => $('#dlg-diff').close());
    $('#btn-assoc-refresh').addEventListener('click', () => loadAssoc());
    $$('input[name="previewMode"]', el.settings).forEach((r) => r.addEventListener('change', async () => {
      if (!r.checked) return;
      if (await saveSettings({ previewMode: r.value }) && state.doc) {
        toast({ exact: '설치된 프로그램으로 미리보기를 다시 만들어요', text: '글자만 보는 미리보기로 바꿨어요' }[r.value] ?? '빠른 미리보기로 다시 만들어요', { type: 'info' });
        refreshPreview({ reason: 'manual' });
      }
    }));
    $$('#theme-seg button').forEach((b) => b.addEventListener('click', () => saveSettings({ theme: b.dataset.theme })));
    $$('#lang-seg button').forEach((b) => b.addEventListener('click', () => saveSettings({ lang: b.dataset.lang })));
    $('#btn-lang-restart').addEventListener('click', () => call(K.relaunch()).catch(showError));
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
      refreshPreview({ ifChanged: true, reason: ev.by === 'editor' ? 'editor' : state.doc.aiBusy ? 'ai' : 'watch' });
    });
    K.on.docEditor(onDocEditor);
    K.on.menu((cmd) => {
      if (cmd === 'open') pickFile();
      else if (cmd === 'new') openNewDoc();
      else if (cmd === 'edit') { if (state.doc) actEdit(); }
      else if (cmd === 'refresh') { if (state.doc) refreshPreview({ reason: 'manual' }); }
      else if (cmd === 'settings') openSettings('ai');
      else if (cmd === 'pii') openPii();
      else if (cmd === 'batch') openBatch();
      else if (cmd === 'pdftools') openPdfTools();
      else if (cmd === 'print') actPrint();
    });
    K.on.batchProgress(onBatchProgress);
    K.on.toast((t) => {
      toast(String(t?.message ?? ''), { type: t?.type ?? 'info' });
      // 확인용 캡처: 인수로 받은 파일을 못 열었으면 알림이 뜬 화면을 찍고 실패로 끝낸다
      if (state.smoke && t?.type === 'error' && !state.doc) smokeSignal({ ok: false, kind: 'open-error', error: String(t?.message ?? '') });
    });
  }

  // AI 패널(다른 뷰)이 부르는 이름 — main이 VIEW_ROUTES대로 중계한다 (ipc-channels.mjs)
  bridge.provide('home.pickFile', pickFile);
  bridge.provide('home.openNewDoc', openNewDoc);
  // 새 문서 창이 열려 있을 때만 다시 그린다 (AI 확인이 끝났을 때 AI 선택지를 고치려고)
  bridge.provide('home.renderNewDoc', () => { if (ne.dlg.open) loadProviders(false); });
  // 그 문서를 보는 문서 탭에만 온다 (main이 경로로 고른다)
  bridge.provide('home.refreshPreview', (opts = {}) => {
    if (opts.path && (!state.doc || pathKey(opts.path) !== pathKey(state.doc.path))) return;
    refreshPreview(opts);
  });
  bridge.provide('home.docEditor', onDocEditor);
  // 패널은 settings:set을 못 부른다: 대신 저장하고, AI를 바꿨으면 패널들이 목록을 다시 받게 한다
  bridge.provide('home.saveSettings', async (patch) => {
    const ok = await saveSettings(patch);
    if (ok && patch && Object.hasOwn(patch, 'provider')) { await loadProviders(false); bridge.call('chat.loadProviders', false); }
    return ok;
  });
  bridge.provide('home.openSettings', openSettings);

  onInit((init) => {
    state.providersReady = loadProviders(false);
    renderRecent(init.recent);
    if (init.doc) onDocOpened(init.doc);
    else { renderTopbar(); renderStatus(); }
    // 열 파일 없이 시작한 확인용 캡처는 빈 화면을 찍는다
    if (state.smoke && !init.doc && !init.opening) setTimeout(() => smokeSignal({ ok: true, kind: 'empty' }), 600);
  });

  bindUi();
  bindEvents();
})();
