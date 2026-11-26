import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createTabList, tabCommandOf, docKey } from '../src/app/tabs.mjs';

const seq = () => { let n = 0; return () => 't' + (++n); };
const doc = (path, o = {}) => ({ type: 'doc', path, title: path, ...o });

test('docKey: 윈도우에서는 소문자 절대 경로', () => {
  if (process.platform === 'win32') assert.equal(docKey('C:\\A\\B.TXT'), 'c:\\a\\b.txt');
  else assert.equal(typeof docKey('/a/B'), 'string');
});

test('createTabList: 홈 탭 하나가 0번·활성, 홈은 닫을 수 없다', () => {
  const l = createTabList({ newId: seq() });
  assert.deepEqual(l.tabs.map((t) => t.id), ['home']);
  assert.equal(l.active.id, 'home');
  assert.deepEqual(l.get('home'), { id: 'home', type: 'home', title: '홈', path: null, convertedFrom: null, dirty: false, loaded: true, zone3: false });
  assert.throws(() => l.remove('home'), { code: 'HOME_FIXED' });
});

test('add: 활성 탭 오른쪽에 넣고 활성, remove: 오른쪽 이웃 → 없으면 왼쪽', () => {
  const l = createTabList({ newId: seq() });
  const a = l.add(doc('C:\\a.txt'));
  const b = l.add(doc('C:\\b.txt'));
  assert.deepEqual(l.tabs.map((t) => t.id), ['home', a.id, b.id]);
  assert.equal(l.active.id, b.id);
  l.activate(a.id);
  const c = l.add(doc('C:\\c.txt'));
  assert.deepEqual(l.tabs.map((t) => t.id), ['home', a.id, c.id, b.id]);
  assert.equal(l.remove(c.id).active.id, b.id);
  const r = l.remove(b.id);
  assert.equal(r.removed.id, b.id);
  assert.equal(r.active.id, a.id);
  const d = l.add(doc('C:\\d.txt'), { activate: false, at: 0 });
  assert.equal(l.tabs[0].id, d.id);
  assert.equal(l.active.id, a.id);
});

test('find: 대소문자 다른 경로와 convertedFrom(원본)으로도 찾는다', () => {
  const l = createTabList({ newId: seq() });
  const t = l.add({ type: 'editor', path: 'C:\\변환\\목록.xlsx', convertedFrom: 'C:\\문서\\목록.csv', title: '목록.csv → .xlsx' });
  assert.equal(l.find('c:\\문서\\목록.CSV')?.id, process.platform === 'win32' ? t.id : undefined);
  assert.equal(l.find('C:\\변환\\목록.xlsx')?.id, t.id);
  assert.equal(l.find('C:\\문서\\다른.csv'), null);
});

test('step·at: 돌아가며 이동, Ctrl+9 = 마지막, 없는 번호는 null', () => {
  const l = createTabList({ newId: seq() });
  const [a, b, c] = ['a', 'b', 'c'].map((n) => l.add(doc('C:\\' + n)));
  l.activate('home');
  assert.equal(l.step(-1).id, c.id);
  assert.equal(l.step(1).id, 'home');
  assert.equal(l.at(2).id, a.id);
  assert.equal(l.at(9).id, c.id);
  assert.equal(l.at(7), null);
  assert.equal(b.type, 'doc');
});

test('replace: 문서 탭을 같은 자리 편집 탭으로 (활성 유지)', () => {
  const l = createTabList({ newId: seq() });
  const a = l.add(doc('C:\\a.txt'));
  const b = l.add(doc('C:\\b.txt'));
  l.activate(a.id);
  const e = l.replace(a.id, { type: 'editor', path: 'C:\\a.txt', title: 'a' });
  assert.deepEqual(l.tabs.map((t) => t.id), ['home', e.id, b.id]);
  assert.equal(l.active.id, e.id);
  assert.equal(l.get(a.id), null);
  l.activate(b.id);
  const e2 = l.replace(e.id, { type: 'editor', path: 'C:\\a.txt', title: 'a2' });
  assert.equal(l.active.id, b.id);
  assert.equal(l.tabs[1].id, e2.id);
});

test('update: id·type 바꾸기는 무시', () => {
  const l = createTabList({ newId: seq() });
  const a = l.add(doc('C:\\a.txt'));
  const u = l.update(a.id, { id: 'x', type: 'editor', dirty: true });
  assert.equal(u.id, a.id);
  assert.equal(u.type, 'doc');
  assert.equal(u.dirty, true);
});

test('snapshot: 홈 빼고 type·path·convertedFrom, active는 홈 포함 차례', () => {
  const l = createTabList({ newId: seq() });
  l.add({ type: 'editor', path: 'C:\\a.xlsx', convertedFrom: 'C:\\a.csv', title: 'a' });
  l.add(doc('C:\\b.txt'));
  assert.deepEqual(l.snapshot(), {
    active: 2,
    tabs: [{ type: 'editor', path: 'C:\\a.xlsx', convertedFrom: 'C:\\a.csv' }, { type: 'doc', path: 'C:\\b.txt' }],
  });
});

test('change 이벤트: add·remove·activate·update·replace마다 한 번', () => {
  const l = createTabList({ newId: seq() });
  let n = 0;
  l.on('change', () => n++);
  const a = l.add(doc('C:\\a.txt')); assert.equal(n, 1);
  l.activate('home'); assert.equal(n, 2);
  l.update(a.id, { dirty: true }); assert.equal(n, 3);
  const e = l.replace(a.id, { type: 'editor', path: 'C:\\a.txt', title: 'a' }); assert.equal(n, 4);
  l.remove(e.id); assert.equal(n, 5);
});

test('tabCommandOf: Ctrl+T/W/Tab/Shift+Tab/1~9/N, Alt·Shift+W·Shift+N·keyUp은 null', () => {
  const k = (code, o = {}) => tabCommandOf({ type: 'keyDown', code, control: true, ...o });
  assert.deepEqual(k('KeyT'), { cmd: 'home' });
  assert.deepEqual(k('KeyW'), { cmd: 'close' });
  assert.deepEqual(k('Tab'), { cmd: 'next' });
  assert.deepEqual(k('Tab', { shift: true }), { cmd: 'prev' });
  assert.deepEqual(k('Digit9'), { cmd: 'index', n: 9 });
  assert.deepEqual(k('Digit1'), { cmd: 'index', n: 1 });
  assert.deepEqual(k('KeyN'), { cmd: 'new-window' });
  assert.equal(k('KeyW', { shift: true }), null);
  assert.equal(k('KeyN', { shift: true }), null);
  assert.equal(k('KeyT', { alt: true }), null);
  assert.equal(k('Digit0'), null);
  assert.equal(tabCommandOf({ type: 'keyUp', code: 'KeyT', control: true }), null);
  assert.equal(tabCommandOf({ type: 'keyDown', code: 'KeyT' }), null);
});

test('tabCommandOf: 누르고 있어 되풀이되는 Ctrl+W·Ctrl+N은 무시, 탭 옮기기는 되풀이 그대로 (F8)', () => {
  const k = (code, o = {}) => tabCommandOf({ type: 'keyDown', code, control: true, isAutoRepeat: true, ...o });
  assert.equal(k('KeyW'), null, '탭이 줄줄이 닫히지 않게');
  assert.equal(k('KeyN'), null, '창이 여럿 뜨지 않게');
  assert.deepEqual(k('Tab'), { cmd: 'next' });
  assert.deepEqual(k('Tab', { shift: true }), { cmd: 'prev' });
  assert.deepEqual(k('KeyW', { isAutoRepeat: false }), { cmd: 'close' });
});
