// bridge.js — 화면 진입점(home.js · panel.js) 사이의 이름 붙은 호출.
// 한쪽이 provide한 이름을 다른 쪽이 call한다. 탭 창에서는 홈 뷰와 AI 패널 뷰가 따로라서, 이 화면에 없는 이름은
// relay(name, args)로 main에 넘긴다 (ipc-channels.mjs VIEW_ROUTES대로 다른 뷰에 view:called로 간다).
// relay가 없으면: 없는 이름 call은 undefined, 이름마다 로그 한 번.
'use strict';
(() => {
  function createBridge({ log = console.warn, relay = null } = {}) {
    const fns = new Map();
    const missed = new Set();
    return Object.freeze({
      provide(name, fn) {
        if (typeof fn !== 'function') throw new TypeError(`bridge: ${name} is not a function`);
        if (fns.has(name)) throw new Error(`bridge: ${name} is already provided`);
        fns.set(name, fn);
      },
      call(name, ...args) {
        const fn = fns.get(name);
        if (fn) return fn(...args);
        // 다른 뷰에 있는 이름: main을 거쳐 넘긴다 (값은 돌려받지 않는다)
        if (relay) { relay(name, args); return undefined; }
        if (!missed.has(name)) { missed.add(name); log(`[bridge] 이 화면에 없는 이름: ${name}`); }
        return undefined;
      },
      has(name) { return fns.has(name); },
      /** 다른 뷰가 보낸 이름(view:called): 이 화면이 provide한 것만 부른다 (없으면 무시, 다시 넘기지 않는다) */
      receive(name, args = []) {
        const fn = fns.get(name);
        if (fn) fn(...args);
      },
    });
  }

  globalThis.OKBridge = Object.freeze({ createBridge });
})();
