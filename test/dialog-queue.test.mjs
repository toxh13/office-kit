// dialog-queue.js 단위 시험 — 확인 창이 겹쳐도 하나씩 차례로
import { test } from 'node:test';
import assert from 'node:assert/strict';
import '../src/app/renderer/dialog-queue.js';

const { createDialogQueue } = globalThis.OKDialogQueue;

test('두 번째 작업은 첫 번째가 끝난 뒤에 시작한다', async () => {
  const q = createDialogQueue();
  const log = [];
  let release;
  const first = q.run(() => new Promise((res) => { log.push('1 시작'); release = () => { log.push('1 끝'); res('a'); }; }));
  const second = q.run(async () => { log.push('2 시작'); return 'b'; });
  await new Promise((r) => setImmediate(r));
  assert.deepEqual(log, ['1 시작']);
  release();
  assert.equal(await first, 'a');
  assert.equal(await second, 'b');
  assert.deepEqual(log, ['1 시작', '1 끝', '2 시작']);
});

test('첫 번째가 throw해도 두 번째는 실행된다', async () => {
  const q = createDialogQueue();
  const first = q.run(() => { throw new Error('실패'); });
  const second = q.run(async () => 'ok');
  await assert.rejects(first, /실패/);
  assert.equal(await second, 'ok');
});
