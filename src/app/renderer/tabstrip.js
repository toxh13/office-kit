// tabstrip.js — 탭 창 맨 위 탭 줄 화면. main이 보낸 탭 목록(okTabs.onList)을 그리고, 누름·키보드를 main에 넘긴다.
// 접근성: role=tablist / role=tab + aria-selected, 로빙 tabindex(활성 0, 나머지 -1),
//   (활성 탭이 없으면 첫 탭이 0), ←/→/Home/End = 초점 이동(활성은 그대로), Enter/Space = 활성, Delete = 그 탭 닫기(홈 제외).
'use strict';

// ── 순수 함수 (시험이 DOM 없이 불러 쓴다) ──
(function (g) {
  /** 키 하나 뒤 초점이 갈 탭 자리. ←/→는 끝에서 돌아가고, Home=0, End=끝, 그 밖은 그대로 */
  function focusAfterKey(index, key, count) {
    if (!(count > 0)) return index;
    switch (key) {
      case 'ArrowLeft': return (index - 1 + count) % count;
      case 'ArrowRight': return (index + 1) % count;
      case 'Home': return 0;
      case 'End': return count - 1;
      default: return index;
    }
  }
  /** 로빙 tabindex 0을 받을 탭 자리: 활성 탭, 활성이 없거나 목록에 없으면 첫 탭 (탭이 없으면 -1) */
  function tabStopIndex(ids, active) {
    if (!ids.length) return -1;
    const at = ids.indexOf(active);
    return at >= 0 ? at : 0;
  }
  /** 활성 탭으로 스크롤할지: 처음이거나 활성 탭이 바뀌었거나 탭이 늘었을 때만 (같은 목록을 다시 그릴 때는 사용자가 옮긴 위치 그대로) */
  function needsReveal(prev, next) {
    return !prev || prev.active !== next.active || next.count > prev.count;
  }
  g.OKTabstrip = { focusAfterKey, tabStopIndex, needsReveal };
})(globalThis);

// ── 화면 ──
if (typeof document !== 'undefined') {
  (() => {
    const K = window.okTabs;
    const { focusAfterKey, tabStopIndex, needsReveal } = globalThis.OKTabstrip;
    const list = document.getElementById('tabs');
    const aiToggle = document.getElementById('ai-toggle');

    // 화면 언어: 한국어 원문을 사전으로 바꾼다 (탭 이름 = 파일 이름은 그대로)
    // tr: 탭 이름 칸(번역하지 않는 곳)에 붙이는 상태 글을 직접 옮길 때 쓴다
    const tr = (() => {
      const d = K?.i18n; const lib = window.OfficeKitI18n;
      if (!d || !lib || !['en', 'ja'].includes(d.lang)) return (s) => s;
      document.documentElement.lang = d.lang;
      const c = lib.compile(d.dict, d.lang);
      lib.translateDom(document.documentElement, c, { skip: '.tab-name' });
      return c.t;
    })();

    let state = { tabs: [], active: null, panelCollapsed: false };
    let tabStop = -1; // 키보드 Tab으로 들어올 탭 자리 (render가 정한다)

    /** 요소 만들기 (글은 textContent로만 넣는다) */
    function h(tag, props = {}, ...children) {
      const el = document.createElement(tag);
      for (const [k, v] of Object.entries(props)) {
        if (v === undefined || v === null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : String(v));
      }
      for (const c of children) if (c) el.append(c);
      return el;
    }
    const srOnly = (text) => h('span', { class: 'sr-only', text });

    const tabButtons = () => [...list.querySelectorAll('[role="tab"]')];
    const isHome = (tab) => tab.type === 'home';

    function activate(tab) {
      if (!K) return;
      if (isHome(tab)) void K.home(); else void K.activate(tab.id);
    }
    function close(tab) {
      if (!K || isHome(tab)) return;
      void K.close(tab.id);
    }

    function onKey(e, tab, index) {
      if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) {
        e.preventDefault();
        const buttons = tabButtons();
        buttons[focusAfterKey(index, e.key, buttons.length)]?.focus();
      } else if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        activate(tab);
      } else if (e.key === 'Delete') {
        e.preventDefault();
        close(tab);
      }
    }

    function renderTab(tab, index) {
      const selected = tab.id === state.active;
      const name = isHome(tab) ? '홈' : String(tab.title ?? '');
      const cls = ['tab', selected && 'active', tab.dirty && 'dirty', tab.loaded === false && 'lazy', tab.zone3 && 'zone3', tab.busy && 'busy']
        .filter(Boolean).join(' ');
      // 마우스를 올리면 이름과 상태(저장 안 한 변경·인터넷 파일)를 함께 보인다 (점·표시만으로는 뜻을 모르므로).
      // 탭 이름 칸은 번역하지 않으므로 상태 글은 여기서 옮긴다
      const states = [tab.dirty && '저장 안 한 변경 있음', tab.zone3 && '인터넷 파일'].filter(Boolean).map(tr);
      const tip = states.length ? `${name} (${states.join(', ')})` : name;
      const button = h('button', {
        type: 'button', role: 'tab', id: 'tab-' + tab.id, class: cls,
        'aria-selected': selected ? 'true' : 'false', tabindex: index === tabStop ? '0' : '-1',
        'data-id': tab.id,
        // 아직 불러오지 않은 탭을 불러오는 중 (지난 세션 복원)
        'aria-busy': tab.loading ? 'true' : null,
        onclick: () => activate(tab),
        onkeydown: (e) => onKey(e, tab, index),
        // 가운데 단추로 닫기 (홈 제외)
        onauxclick: (e) => { if (e.button === 1) { e.preventDefault(); close(tab); } },
      },
      tab.dirty ? h('span', { class: 'dirty-dot', 'aria-hidden': 'true', text: '●' }) : null,
      isHome(tab) ? h('span', { class: 'tab-home', text: name }) : h('span', { class: 'tab-name', 'data-no-i18n': true, title: tip, text: name }),
      tab.dirty ? srOnly('저장 안 한 변경 있음') : null,
      tab.loaded === false ? srOnly('아직 불러오지 않음') : null,
      tab.zone3 ? srOnly('인터넷 파일') : null);
      const wrap = h('div', { class: isHome(tab) ? 'tab-wrap home' : 'tab-wrap', role: 'presentation' }, button);
      if (!isHome(tab)) {
        // 닫기 단추는 마우스용 (키보드는 Delete)
        wrap.append(h('button', {
          type: 'button', class: 'tab-close', tabindex: '-1', 'aria-label': `‘${name}’ 닫기`,
          onclick: (e) => { e.stopPropagation(); close(tab); },
        }, h('span', { 'aria-hidden': 'true', text: '×' })));
      }
      return wrap;
    }

    let revealed = null; // 마지막으로 활성 탭을 보이게 한 때의 { active, count }
    function render() {
      // 다시 그려도 초점은 같은 탭에 (닫혀 없어졌으면 활성 탭으로)
      const focusedId = document.activeElement?.closest?.('[role="tab"]')?.dataset.id ?? null;
      tabStop = tabStopIndex(state.tabs.map((t) => t.id), state.active);
      list.replaceChildren(...state.tabs.map(renderTab));
      if (focusedId !== null) {
        const again = tabButtons().find((b) => b.dataset.id === focusedId)
          ?? tabButtons().find((b) => b.getAttribute('aria-selected') === 'true');
        again?.focus();
      }
      aiToggle.setAttribute('aria-pressed', state.panelCollapsed ? 'false' : 'true');
      // 탭이 많아 줄이 넘치면 활성 탭이 보이게 (가로로만, 가까운 쪽으로). 활성 탭이 바뀌거나 탭이 늘 때만
      const shown = { active: state.active, count: state.tabs.length };
      if (needsReveal(revealed, shown)) tabButtons().find((b) => b.getAttribute('aria-selected') === 'true')?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
      revealed = shown;
    }

    aiToggle.addEventListener('click', () => { if (K) void K.togglePanel(); });

    if (K) {
      K.onList((data) => {
        if (!data || !Array.isArray(data.tabs)) return;
        state = { tabs: data.tabs, active: data.active ?? null, panelCollapsed: !!data.panelCollapsed };
        render();
      });
      K.ready();
    }
  })();
}
