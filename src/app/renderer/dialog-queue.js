// dialog-queue.js — 확인 창을 하나씩 차례로 띄우는 대기열.
// 확인 창 요소(<dialog>)는 하나뿐이라, 두 창이 겹치면 한 번 누름에 두 동작이 실행될 수 있다.
// run(fn)은 앞 작업이 (실패해도) 끝난 뒤에 fn을 시작하고, fn의 결과로 끝나는 Promise를 돌려준다.
'use strict';
(() => {
  function createDialogQueue() {
    let tail = Promise.resolve();
    return {
      run(fn) {
        const result = tail.then(() => fn());
        tail = result.then(() => {}, () => {});
        return result;
      },
    };
  }
  globalThis.OKDialogQueue = Object.freeze({ createDialogQueue });
})();
