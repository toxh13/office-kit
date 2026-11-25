// target-chips.js — AI 패널의 "대상 칩": 고칠 문서 1개 + 읽기 재료들 + 새 파일 출력 폴더.
// model()은 화면 값(열린 문서 목록·고른 재료)을 칩 모양으로 정리하고, targetOf()는 ai:send의 target을 만든다.
// render()는 DOM만 그린다(h·icon·콜백은 호출하는 쪽이 넘긴다).
'use strict';
(() => {
  const key = (p) => String(p ?? '').replace(/\//g, '\\').toLowerCase();
  const folderOf = (p) => {
    const s = String(p ?? '');
    const i = Math.max(s.lastIndexOf('\\'), s.lastIndexOf('/'));
    return i > 0 ? s.slice(0, i) : null;
  };

  function model({ info, materials, outDir }) {
    const open = Array.isArray(info?.open) ? info.open : [];
    const edit = info?.edit ?? null;
    const editKey = edit ? key(edit.path) : null;
    const latest = new Map(open.map((m) => [key(m.path), m]));
    const seen = new Set();
    const read = [];
    const take = (m, removable) => {
      const k = key(m.path);
      if (k === editKey || seen.has(k) || !latest.has(k)) return;
      seen.add(k);
      read.push({ ...latest.get(k), removable });
    };
    const materialOnly = !edit;
    // 재료 전용이면 지금 문서가 뺄 수 없는 첫 재료
    if (materialOnly) { const cur = open.find((m) => m.current); if (cur) take(cur, false); }
    for (const m of materials ?? []) take(m, true);
    const addable = open
      .filter((m) => key(m.path) !== editKey && !seen.has(key(m.path)))
      .map((m) => ({ ...m, disabled: !!m.busy }));
    return {
      edit,
      read,
      addable,
      materialOnly,
      narrowed: read.some((m) => m.zone3),
      outDir: outDir ?? (materialOnly && read[0] ? folderOf(read[0].path) : null),
      busyMaterial: read.find((m) => m.busy) ?? null,
    };
  }

  function targetOf(m) {
    return { edit: m.edit ? m.edit.path : null, read: m.read.map((x) => x.path) };
  }

  // root별 마지막 그림 값 — 펼친 목록·Esc 닫기에 쓴다
  const states = new WeakMap();

  function render(root, m, opts) {
    const { h, icon, onRemove, onAdd, onOpenMenu, onPickOutDir } = opts;
    const st = states.get(root) ?? { open: false };
    states.set(root, st);
    st.m = m; st.opts = opts;
    const redraw = (focusMenuBtn) => {
      render(root, st.m, st.opts);
      if (focusMenuBtn) root.querySelector('.tc-add')?.focus();
    };
    const badge = (text) => h('span', { class: 'tc-badge', text });
    const chips = [];
    if (m.edit) {
      chips.push(h('span', { class: 'tc-chip tc-edit' }, h('span', { class: 'tc-name', text: m.edit.name }), h('span', { class: 'tc-role', text: '고칠 문서' })));
    }
    for (const r of m.read) {
      const parts = [h('span', { class: 'tc-name', text: r.name }), h('span', { class: 'tc-role', text: '읽기 재료' })];
      if (r.zone3) parts.push(badge('인터넷 파일'));
      if (r.busy) parts.push(badge('AI 작업 중'));
      if (r.removable) {
        parts.push(h('button', { class: 'tc-x', type: 'button', 'aria-label': `${r.name} 재료에서 빼기`, onclick: () => onRemove(r.path) }, icon('x')));
      }
      chips.push(h('span', { class: 'tc-chip tc-read' }, ...parts));
    }
    const addBtn = h('button', {
      class: 'tc-add', type: 'button', 'aria-haspopup': 'menu', 'aria-expanded': st.open ? 'true' : 'false',
      onclick: () => { st.open = !st.open; if (st.open) onOpenMenu(); redraw(false); },
    }, '+ 열린 탭 추가');
    const row = [...chips, addBtn];
    const nodes = [h('div', { class: 'tc-row' }, ...row)];
    if (st.open) {
      const items = m.addable.length
        ? m.addable.map((a) => {
          const props = { class: 'tc-item', type: 'button', role: 'menuitem', onclick: () => { st.open = false; onAdd(a.path); redraw(true); } };
          if (a.disabled) { props['aria-disabled'] = 'true'; props.onclick = (e) => e.preventDefault(); }
          return h('button', props, h('span', { text: a.name }), a.disabled ? h('span', { class: 'tc-badge', text: 'AI 작업 중' }) : null);
        })
        : [h('div', { class: 'tc-empty', text: '추가할 수 있는 열린 문서가 없어요' })];
      nodes.push(h('div', {
        class: 'tc-menu', role: 'menu',
        onkeydown: (e) => { if (e.key === 'Escape') { e.stopPropagation(); st.open = false; redraw(true); } },
      }, ...items));
    }
    if (m.narrowed) nodes.push(h('div', { class: 'tc-note', text: '인터넷 재료가 있어 읽기 범위를 좁혔습니다' }));
    if (m.materialOnly) {
      nodes.push(h('div', { class: 'tc-note' },
        h('span', { text: '고칠 문서 없이 새 파일만 만들어요' }),
        m.outDir ? h('span', { class: 'tc-out', text: m.outDir }) : null,
        h('button', { class: 'link-btn', type: 'button', onclick: () => onPickOutDir() }, '출력 폴더 바꾸기')));
    }
    root.replaceChildren(...nodes);
    if (st.open) root.querySelector('.tc-item:not([aria-disabled])')?.focus?.();
  }

  globalThis.OKTargetChips = Object.freeze({ model, targetOf, render });
})();
