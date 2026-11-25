// doc-actions.js — 버전 목록·바뀐 내용·되돌리기 화면 (홈 종류 뷰와 AI 패널 뷰가 같이 쓴다, 판정 5).
// 문서는 경로로 받는다: main의 doc:versions/diff/restore는 그 경로가 이 뷰의 문서이거나 그 탭 창의 AI 대상일 때만 허용한다.
// 대화상자 요소(#dlg-diff·#versions-pop·#dlg-confirm)는 index.html에 있어 어느 뷰에서도 쓸 수 있다.
// 되돌리기 앞뒤로 document에 'okdoc:restore' 이벤트 { path, phase: 'start'|'end', doc }를 낸다 (home.js가 상단 막대를 고친다)
'use strict';
(() => {
  const {
    K, $, h, spinner, call, pathKey, whenText, setBusy, clearBusy, toast, showError, confirmDialog, state, bridge,
  } = OKCore;

  const VERSION_LABEL = { open: '처음 연 상태', editor: '직접 편집 저장 전', ai: 'AI 대화 전', restore: '원래대로 전' };
  const versionLabel = (v) => VERSION_LABEL[v?.reason] ?? '백업';

  /** 이 뷰가 아는 그 문서 정보 (홈 종류 뷰 = 그 탭의 문서, 패널 = AI 대상). 다른 문서면 null */
  const docOf = (path) => (state.doc && pathKey(state.doc.path) === pathKey(path) ? state.doc : null);

  /** main이 경로 검사로 거절한 것(열린 탭의 문서가 아님)은 쉬운 말로 */
  function explain(err) {
    if (err?.code === 'NOT_OPEN_TAB') return Object.assign(new Error('이 문서는 열린 탭이 아니에요'), { code: err.code });
    return err;
  }

  const emit = (detail) => document.dispatchEvent(new CustomEvent('okdoc:restore', { detail }));

  // ── 되돌리기 ─────────────────────────────────────────

  /** 버전으로 되돌리기. file을 주지 않으면 마지막 변경 전으로 (지금과 다른 가장 최근 버전) */
  async function restore(path, file) {
    const d = docOf(path);
    // 백업이 없다고 알고 있는 문서(홈 종류 뷰)는 하지 않는다
    if (!d || (Object.hasOwn(d, 'backup') && !d.backup) || state.restoring || state.restoreAsking) return;
    state.restoreAsking = true; // 확인 창이 떠 있는 동안에도 다시 누름을 무시한다
    try { await restoreFlow(path, file, d); } finally { state.restoreAsking = false; }
  }

  async function restoreFlow(path, file, d) {
    closeVersions();
    let target;
    try {
      const v = await call(K.doc.versions({ path }));
      target = typeof file === 'string' ? v.versions.find((x) => x.file === file) : v.versions.find((x) => x.file === v.undo);
      if (typeof file !== 'string' && !target) {
        // 바뀐 것이 없다. 편집 창에 저장 안 한 내용이 있을 때만 (그것을 버릴지 묻게) 지금 버전으로 되돌린다
        if (!d.editor?.dirty) { toast('되돌릴 변경이 없어요. 문서가 마지막으로 남긴 버전과 같아요.', { type: 'info' }); return; }
        target = v.versions[0];
      }
      if (!target) { toast('그 버전을 찾을 수 없어요. 버전 목록을 다시 열어 주세요.', { type: 'warn' }); return; }
    } catch (err) { showError(explain(err)); return; }
    // AI가 고치는 중이면 먼저 멈춘다 (main이 멈추고 끝날 때까지 기다린 뒤 되돌린다)
    const aiBusy = !!docOf(path)?.aiBusy;
    const what = `${whenText(target.time)} · ${versionLabel(target)}`;
    const yes = await confirmDialog({
      title: typeof file === 'string' ? '이 버전으로 되돌릴까요?' : '마지막 변경 전으로 되돌릴까요?',
      body: [
        aiBusy ? '진행 중인 AI 작업을 멈추고 되돌립니다.' : null,
        `${what} 버전으로 되돌립니다.`,
        '지금 상태도 버전 목록에 ‘원래대로 전’으로 남겨 두니, 다시 돌아올 수 있어요.',
        !d.editor?.open ? null : d.editor.dirty
          ? '편집 창에 저장하지 않은 내용이 있어요. 되돌린 뒤 편집 창에서 다시 불러올지, 내 편집을 유지할지 고를 수 있어요.'
          : '열려 있는 편집 창도 되돌린 내용으로 다시 불러와요.',
      ].filter(Boolean),
      ok: aiBusy ? '멈추고 되돌리기' : '되돌리기', tone: 'danger',
    });
    if (!yes || !docOf(path)) return;
    if (aiBusy) bridge.call('chat.markStopping', { path });
    state.restoring = true;
    emit({ path, phase: 'start' });
    setBusy('restore', aiBusy ? 'AI 작업을 멈추고 되돌리는 중…' : '되돌리는 중…');
    let doc = null;
    try {
      const r = await call(K.doc.restore({ path, file: target.file }));
      doc = r.doc ?? null;
      const editorNote = !r.editorNotified ? '' : docOf(path)?.editor?.dirty ? '. 편집 창에서 다시 불러올지 골라 주세요' : '. 편집 창도 다시 불러와요';
      toast(r.stoppedAi > 0 ? `AI 작업을 멈추고 ${what} 버전으로 되돌렸어요${editorNote}` : `${what} 버전으로 되돌렸어요${editorNote}`, { type: 'success' });
      // main이 이 문서의 AI 대화를 끝냈다 → 대화창에 알린다 (AI가 되돌리기 전 내용을 기억하지 않게)
      bridge.call('chat.restored', { path, newChat: !!r.newChat });
      // 그 문서의 미리보기(문서 탭)를 새로 그린다
      bridge.call('home.refreshPreview', { path, reason: 'manual' });
    } catch (err) { showError(explain(err)); } finally {
      state.restoring = false;
      clearBusy('restore');
      emit({ path, phase: 'end', doc });
    }
  }

  // ── 바뀐 내용 (버전 이후 · AI 대화 전 이후) ──────────

  const DIFF_STAT = { changed: ['바뀜', 'warn'], added: ['더함', 'on'], removed: ['지움', 'err'] };

  /** 글자 단위로 바뀐 곳을 표시한 문단 */
  function segText(segs, side) {
    return segs.filter((s) => s.t === '=' || s.t === side).map((s) => (s.t === '=' ? document.createTextNode(s.s) : h(side === '-' ? 'del' : 'ins', { text: s.s })));
  }
  const lineText = (t) => (t.trim() ? t : h('span', { class: 'diff-empty-line', text: '(빈 줄)' }));
  const diffLine = (cls, sign, ...kids) => h('div', { class: `diff-line ${cls}` }, h('span', { class: 'diff-sign', text: sign }), h('span', { class: 'diff-text' }, ...kids));

  function diffTextNodes(d) {
    return d.hunks.map((hk) => h('div', { class: 'diff-hunk' },
      hk.where ? h('div', { class: 'diff-where', text: hk.where }) : null,
      ...hk.before.map((t) => diffLine('ctx', '', t)),
      ...hk.items.flatMap((it) => (it.type === 'mod'
        ? [diffLine('mod del', '−', ...segText(it.segs, '-')), diffLine('mod add', '+', ...segText(it.segs, '+'))]
        : it.type === 'del' ? [diffLine('del', '−', lineText(it.old))] : [diffLine('add', '+', lineText(it.new))])),
      ...hk.after.map((t) => diffLine('ctx', '', t))));
  }

  function diffCellNodes(d) {
    const out = [];
    if (d.sheetsAdded.length) out.push(h('p', { class: 'help', text: `새 시트: ${d.sheetsAdded.join(', ')}` }));
    if (d.sheetsRemoved.length) out.push(h('p', { class: 'help', text: `없어진 시트: ${d.sheetsRemoved.join(', ')}` }));
    if (d.cells.length) {
      out.push(h('table', { class: 'diff-table' },
        h('thead', {}, h('tr', {}, h('th', { text: '칸' }), h('th', { text: '전' }), h('th', { text: '후' }))),
        h('tbody', {}, ...d.cells.map((c) => h('tr', {},
          h('td', { class: 'addr', text: `${c.sheet}!${c.cell}` }),
          h('td', { class: c.old ? 'old' : '', text: c.old || '(빈칸)', 'data-no-i18n': !!c.old }),
          h('td', { class: c.new ? 'new' : '', text: c.new || '(빈칸)', 'data-no-i18n': !!c.new }))))));
    }
    return out;
  }

  /** path 문서의 file 버전(백업) → 지금 사이에 바뀐 글을 보여 준다 */
  async function showDiff(path, file, { label = '이 버전', time = null, ai = false } = {}) {
    if (!docOf(path)) return;
    const dlg = $('#dlg-diff');
    const body = $('#diff-body');
    $('#diff-title').textContent = ai ? 'AI가 바꾼 내용' : '바뀐 내용';
    $('#diff-sub').textContent = `${label}${time ? ` (${whenText(time)})` : ''} → 지금`;
    $('#diff-stats').replaceChildren();
    body.replaceChildren(h('div', { class: 'diff-msg' }, spinner(), ' 비교하는 중…'));
    const restoreBtn = $('#diff-restore');
    restoreBtn.disabled = true;
    restoreBtn.lastChild.textContent = ai ? '이 대화 전으로 되돌리기' : '이 버전으로 되돌리기';
    restoreBtn.onclick = () => { dlg.close(); restore(path, file); };
    if (!dlg.open) dlg.showModal();
    let d;
    try { d = await call(K.doc.diff({ path, file })); } catch (err) {
      body.replaceChildren(h('div', { class: 'diff-msg' }, h('div', { class: 'inline-error', text: explain(err).message })));
      if (err.code === 'NO_VERSION' || err.code === 'NOT_OPEN_TAB') return;
      restoreBtn.disabled = false;
      return;
    }
    // 그 사이 다른 문서로 바뀌었으면 그리지 않는다 (같은 문서의 정보가 새로 고쳐진 것은 괜찮다)
    if (!dlg.open || !docOf(path)) return;
    if (!ai) $('#diff-sub').textContent = `${label} (${whenText(d.time)}) → 지금`;
    $('#diff-stats').replaceChildren(...Object.entries(DIFF_STAT).filter(([k]) => d.stats[k] > 0)
      .map(([k, [txt, cls]]) => h('span', { class: `pill ${cls}`, text: `${txt} ${d.stats[k]}${d.kind === 'cells' ? '칸' : '곳'}` })));
    restoreBtn.disabled = false;
    if (d.same) {
      body.replaceChildren(h('div', { class: 'diff-msg', text: '바뀐 글이 없어요. 서식·그림만 바뀌었거나 같은 내용이에요.' }));
      return;
    }
    const nodes = d.kind === 'cells' ? diffCellNodes(d) : diffTextNodes(d);
    if (d.truncated) nodes.push(h('p', { class: 'help', text: '바뀐 곳이 많아 앞부분만 보여요.' }));
    if (d.approximate) nodes.unshift(h('p', { class: 'help', text: '너무 많이 바뀌어서 대략 비교했어요.' }));
    body.replaceChildren(...nodes);
    body.scrollTop = 0;
  }

  // ── 버전 목록 (원래대로 ▾) ───────────────────────────

  function closeVersions() {
    const pop = $('#versions-pop');
    if (!pop || pop.hidden) return;
    pop.hidden = true;
    $('#btn-versions')?.setAttribute('aria-expanded', 'false');
  }

  /** path 문서의 버전 목록을 열고 닫는다 (#btn-versions 아래) */
  async function openVersions(path) {
    const pop = $('#versions-pop');
    const btn = $('#btn-versions');
    if (!pop.hidden) { closeVersions(); return; }
    if (!docOf(path)) return;
    let v;
    try { v = await call(K.doc.versions({ path })); } catch (err) { showError(explain(err)); return; }
    const item = (x) => h('li', { class: `ver-item${x.file === v.undo ? ' is-undo' : ''}` },
      h('div', { class: 'ver-text' },
        h('span', { class: 'ver-label', text: versionLabel(x) }),
        h('span', { class: 'ver-time', text: whenText(x.time) }),
        x.file === v.undo ? h('span', { class: 'ver-tag', text: '마지막 변경 전 (원래대로 버튼)' }) : null),
      x.current ? h('span', { class: 'ver-now', text: '지금 상태' })
        : h('span', { class: 'ver-btns' },
          h('button', { class: 'btn sm ghost', text: '바뀐 내용', title: '이 버전 이후 바뀐 글 보기', onclick: () => { closeVersions(); showDiff(path, x.file, { label: versionLabel(x), time: x.time }); } }),
          h('button', { class: 'btn sm', text: '되돌리기', title: '이 버전으로 되돌리기', onclick: () => restore(path, x.file) })));
    $('#ver-list').replaceChildren(...(v.versions.length ? v.versions.map(item) : [h('li', { class: 'ver-empty', text: '아직 남긴 버전이 없어요' })]));
    const r = btn.getBoundingClientRect();
    pop.hidden = false;
    pop.style.top = `${Math.round(r.bottom + 6)}px`;
    pop.style.left = `${Math.round(Math.max(12, Math.min(r.right - pop.offsetWidth, window.innerWidth - pop.offsetWidth - 12)))}px`;
    btn.setAttribute('aria-expanded', 'true');
    ($('#ver-list button') ?? $('#ver-open-backups')).focus();
  }

  globalThis.OKDocActions = Object.freeze({ openVersions, closeVersions, showDiff, restore, versionLabel });
})();
