import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createSessionKeeper, promoteCrashed, restorePlan, SESSION_DEBOUNCE_MS } from '../src/app/tab-session.mjs';

// 가짜 설정: 쓴 횟수를 센다
const fakeSettings = (data = {}) => ({
  data, writes: 0,
  get(k) { return this.data[k] ?? null; },
  set(p) { this.writes++; Object.assign(this.data, p); },
});
// 가짜 타이머 큐
const fakeTimers = () => {
  const q = new Map(); let n = 0;
  return {
    setTimer: (fn, ms) => { q.set(++n, { fn, ms }); return n; },
    clearTimer: (id) => { q.delete(id); },
    get pending() { return q.size; },
    run() { const all = [...q.values()]; q.clear(); all.forEach((t) => t.fn()); },
  };
};
const mk = (o = {}) => {
  const settings = fakeSettings(o.data); const timers = fakeTimers();
  const keeper = createSessionKeeper({ settings, explorer: o.explorer ?? false, setTimer: timers.setTimer, clearTimer: timers.clearTimer, now: () => 111 });
  return { settings, timers, keeper };
};
const win = (n = 0) => ({ bounds: { x: n, y: 0, width: 800, height: 600 }, maximized: false, snapshot: { active: 0, tabs: [] } });

test('update: 1초 안의 여러 변경은 settings.set 한 번 (D9)', () => {
  const { settings, timers, keeper } = mk();
  keeper.update(1, win()); keeper.update(1, win(1)); keeper.update(1, win(2));
  assert.equal(settings.writes, 0);
  assert.equal(SESSION_DEBOUNCE_MS, 1000);
  timers.run();
  assert.equal(settings.writes, 1);
  const s = settings.data.tabSession;
  assert.equal(s.v, 1);
  assert.equal(s.at, 111);
  assert.equal(s.explorer, false);
  assert.equal(s.finalized, false);
  assert.equal(s.windows.length, 1);
  assert.equal(s.windows[0].bounds.x, 2);
});

test('update: 창은 만든 차례대로 저장한다', () => {
  const { settings, timers, keeper } = mk();
  keeper.update(5, win(5)); keeper.update(3, win(3)); keeper.update(5, win(6));
  timers.run();
  assert.deepEqual(settings.data.tabSession.windows.map((w) => w.bounds.x), [6, 3]);
});

test('remove: 여러 창 중 하나를 닫으면 목록에서 빠진다 (D11)', () => {
  const { settings, timers, keeper } = mk();
  keeper.update(1, win(1)); keeper.update(2, win(2)); timers.run();
  keeper.remove(1);
  assert.equal(timers.pending, 1);
  timers.run();
  assert.deepEqual(settings.data.tabSession.windows.map((w) => w.bounds.x), [2]);
});

test('finalize: 즉시 쓰고 lastTabSession도 확정, 그 뒤 update·remove는 저장 안 함', () => {
  const { settings, timers, keeper } = mk();
  keeper.update(1, win(1)); keeper.update(2, win(2));
  keeper.finalize();
  assert.equal(timers.pending, 0, '기다리던 디바운스를 버린다');
  assert.equal(settings.writes, 1);
  assert.equal(settings.data.lastTabSession.windows.length, 2);
  assert.equal(settings.data.lastTabSession.finalized, true);
  assert.equal(settings.data.tabSession.finalized, true);
  keeper.remove(2); keeper.update(1, { bounds: null, maximized: false, snapshot: { active: 0, tabs: [] } });
  assert.equal(timers.pending, 0);
  assert.equal(settings.writes, 1);
  assert.equal(settings.data.tabSession.windows.length, 2);
});

test('explorer 세션: tabSession만 쓰고 lastTabSession은 덮지 않는다 (D11)', () => {
  const old = { v: 1, windows: [{ tabs: [] }] };
  const { settings, keeper } = mk({ explorer: true, data: { lastTabSession: old } });
  keeper.update(1, win());
  keeper.finalize();
  assert.equal(settings.data.lastTabSession, old);
  assert.equal(settings.data.tabSession.explorer, true);
  assert.equal(settings.data.tabSession.finalized, true);
});

test('resume: 끝내기 취소 뒤 다시 저장한다', () => {
  const { settings, timers, keeper } = mk();
  keeper.update(1, win(1));
  keeper.finalize();
  keeper.update(2, win(2));
  assert.equal(timers.pending, 0);
  keeper.resume();
  assert.equal(timers.pending, 1);
  timers.run();
  assert.equal(settings.data.tabSession.finalized, false);
  assert.deepEqual(settings.data.tabSession.windows.map((w) => w.bounds.x), [1]);
  keeper.update(2, win(2)); timers.run();
  assert.equal(settings.data.tabSession.windows.length, 2);
});

test('promoteCrashed: 확정 안 된 평소 세션만 지난 세션으로 올린다', () => {
  const a = fakeSettings({ tabSession: { v: 1, explorer: false, finalized: false, windows: [] } });
  assert.equal(promoteCrashed(a), true);
  assert.equal(a.data.lastTabSession, a.data.tabSession);
  const b = fakeSettings({ tabSession: { v: 1, explorer: true, finalized: false, windows: [] } });
  assert.equal(promoteCrashed(b), false);
  assert.equal(b.data.lastTabSession, undefined);
  const c = fakeSettings({ tabSession: { v: 1, explorer: false, finalized: true, windows: [] } });
  assert.equal(promoteCrashed(c), false);
  assert.equal(promoteCrashed(fakeSettings()), false);
});

test('restorePlan: 없는 파일은 빼고 missing에, 없는 활성 탭은 오른쪽 → 왼쪽 → 홈', () => {
  const s = { v: 1, windows: [{ bounds: null, maximized: false, active: 2, tabs: [
    { type: 'editor', path: 'C:\\문서\\a.hwpx' }, { type: 'doc', path: 'C:\\문서\\없음.pdf' }, { type: 'editor', path: 'C:\\문서\\c.docx' }] }] };
  const p = restorePlan(s, { exists: (f) => !f.includes('없음') });
  assert.deepEqual(p.missing, ['C:\\문서\\없음.pdf']);
  assert.deepEqual(p.windows[0].tabs.map((t) => t.path), ['C:\\문서\\a.hwpx', 'C:\\문서\\c.docx']);
  assert.equal(p.windows[0].active, 2, '없어진 활성 탭(2번) 자리 → 오른쪽 이웃 c(새 2번)');
});

test('restorePlan: 오른쪽 이웃이 없으면 왼쪽, 다 없으면 홈, 탭 없는 창도 남긴다', () => {
  const t = (n) => ({ type: 'doc', path: 'C:\\문서\\' + n });
  const none = (f) => f.includes('x');
  const left = restorePlan({ v: 1, windows: [{ active: 2, tabs: [t('a'), t('x1')] }] }, { exists: (f) => !f.includes('x') });
  assert.equal(left.windows[0].active, 1);
  const home = restorePlan({ v: 1, windows: [{ active: 1, tabs: [t('x1')] }] }, { exists: (f) => !none(f) });
  assert.equal(home.windows.length, 1);
  assert.equal(home.windows[0].active, 0);
  assert.deepEqual(home.windows[0].tabs, []);
  const kept = restorePlan({ v: 1, windows: [{ active: 1, tabs: [t('a'), t('b')] }] }, { exists: () => true });
  assert.equal(kept.windows[0].active, 1);
  assert.deepEqual(kept.missing, []);
});

test('restorePlan: 버전이 다르거나 모양이 틀리면 빈 계획', () => {
  const empty = { windows: [], missing: [] };
  assert.deepEqual(restorePlan(null), empty);
  assert.deepEqual(restorePlan({ v: 2, windows: [] }), empty);
  assert.deepEqual(restorePlan({ v: 1, windows: 'x' }), empty);
});
