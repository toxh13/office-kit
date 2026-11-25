// bridge.js — 화면 진입점(home.js · panel.js) 사이의 이름 붙은 호출.
// 한쪽이 provide한 이름을 다른 쪽이 call한다. 이 이름 목록이 M3에서 IPC(탭 창을 거치는 이벤트)로 바뀔 목록이다.
// 한 진입점만 있는 화면(M3의 패널 뷰 단독)에서도 던지지 않게: 없는 이름 call은 undefined, 이름마다 로그 한 번.
'use strict';
(() => {
  function createBridge({ log = console.warn } = {}) {
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
        if (!missed.has(name)) { missed.add(name); log(`[bridge] 이 화면에 없는 이름: ${name}`); }
        return undefined;
      },
      has(name) { return fns.has(name); },
    });
  }

  globalThis.OKBridge = Object.freeze({ createBridge });
})();
