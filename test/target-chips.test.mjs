import test from 'node:test';
import assert from 'node:assert/strict';
import '../src/app/renderer/target-chips.js';

const { model, targetOf } = globalThis.OKTargetChips;
const M = (path, o = {}) => ({ path, name: path.split('\\').pop(), kind: 'Word', zone3: false, busy: false, dirty: false, current: false, ...o });
test('고칠 문서 + 재료: 재료는 뺄 수 있고, 추가 목록에서 이미 고른 것·고칠 문서는 빠진다', () => {
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\명단.xlsx'), M('D:\\자료\\표.xlsx')] };
  const m = model({ info, materials: [info.open[0]], outDir: null });
  assert.equal(m.edit.path, 'C:\\문서\\공문.hwpx');
  assert.deepEqual(m.read.map((x) => [x.path, x.removable]), [['D:\\자료\\명단.xlsx', true]]);
  assert.deepEqual(m.addable.map((x) => x.path), ['D:\\자료\\표.xlsx']);
  assert.equal(m.materialOnly, false);
  assert.deepEqual(targetOf(m), { edit: 'C:\\문서\\공문.hwpx', read: ['D:\\자료\\명단.xlsx'] });
});
test('재료 전용: 지금 문서(PDF)는 뺄 수 없는 첫 재료, 출력 폴더 기본=첫 재료 폴더, 고르면 그 폴더', () => {
  const info = { edit: null, open: [M('D:\\자료\\보고.pdf', { current: true, kind: 'PDF' })] };
  let m = model({ info, materials: [], outDir: null });
  assert.equal(m.materialOnly, true);
  assert.deepEqual(m.read.map((x) => [x.path, x.removable]), [['D:\\자료\\보고.pdf', false]]);
  assert.equal(m.outDir, 'D:\\자료');
  m = model({ info, materials: [], outDir: 'E:\\결과' });
  assert.equal(m.outDir, 'E:\\결과');
  assert.deepEqual(targetOf(m), { edit: null, read: ['D:\\자료\\보고.pdf'] });
});
test('인터넷 재료가 있으면 narrowed, AI 작업 중 재료는 추가 목록에서 disabled, 골라 둔 것은 busyMaterial', () => {
  const busy = M('D:\\자료\\고치는중.docx', { busy: true });
  const net = M('D:\\받은 파일\\안내.pdf', { zone3: true });
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [busy, net] };
  const m = model({ info, materials: [net, busy], outDir: null });
  assert.equal(m.narrowed, true);
  assert.equal(m.busyMaterial.path, busy.path);
  assert.equal(model({ info, materials: [], outDir: null }).addable.find((x) => x.path === busy.path).disabled, true);
});
test('닫힌 문서는 재료에서 사라지고, 대소문자만 다른 경로는 같은 문서(최신 info 값을 쓴다)', () => {
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\명단.xlsx')] };
  const m = model({ info, materials: [M('D:\\자료\\닫힘.xlsx'), M('d:\\자료\\명단.XLSX')], outDir: null });
  assert.deepEqual(m.read.map((x) => x.path), ['D:\\자료\\명단.xlsx']);
});

// ── render: 가짜 DOM으로 접근성 속성과 초점 동작을 확인한다 ──
const { render } = globalThis.OKTargetChips;
function fakeEnv() {
  const focused = [];
  const h = (tag, props = {}, ...children) => {
    const n = { tag, props: props ?? {}, children: children.flat().filter((c) => c !== null && c !== undefined && c !== false) };
    n.focus = () => focused.push(n);
    return n;
  };
  const walk = (n, fn) => { fn(n); for (const c of n.children ?? []) if (typeof c === 'object') walk(c, fn); };
  const cls = (n) => String(n.props?.class ?? '').split(' ');
  const root = {
    kids: [],
    replaceChildren(...k) { this.kids = k; },
    contains() { return false; },
    find(pred) { const r = []; for (const k of this.kids) walk(k, (n) => { if (pred(n)) r.push(n); }); return r; },
    querySelector(sel) {
      const pred = sel === '.tc-add' ? (n) => cls(n).includes('tc-add')
        : (n) => cls(n).includes('tc-item') && n.props['aria-disabled'] === undefined;
      return this.find(pred)[0] ?? null;
    },
  };
  return { h, icon: (name) => ({ tag: 'svg', props: { class: name }, children: [] }), root, focused };
}
const noop = () => {};
const cbs = (env) => ({ h: env.h, icon: env.icon, onRemove: noop, onAdd: noop, onOpenMenu: noop, onPickOutDir: noop });
const addBtn = (env) => env.root.find((n) => String(n.props.class ?? '').includes('tc-add'))[0];

test('render: ✕ 버튼에 재료 이름이 든 aria-label, 추가 버튼은 aria-expanded가 열림 상태를 따른다', () => {
  const env = fakeEnv();
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\명단.xlsx'), M('D:\\자료\\표.xlsx')] };
  render(env.root, model({ info, materials: [info.open[0]], outDir: null }), cbs(env));
  assert.equal(env.root.find((n) => n.props.class === 'tc-x')[0].props['aria-label'], '명단.xlsx 재료에서 빼기');
  assert.equal(addBtn(env).props['aria-expanded'], 'false');
  assert.equal(addBtn(env).props['aria-haspopup'], undefined, '메뉴 역할을 쓰지 않는다 (단추 목록) — 의도한 변경');
  addBtn(env).props.onclick();
  assert.equal(addBtn(env).props['aria-expanded'], 'true');
  const list = env.root.find((n) => String(n.props.class ?? '') === 'tc-menu')[0];
  assert.equal(addBtn(env).props['aria-controls'], list.props.id, '펼친 목록을 가리킨다');
  assert.equal(list.props.role, 'group');
  assert.equal(list.props['aria-label'], '추가할 열린 문서');
});

test('render: AI 작업 중인 후보는 aria-disabled, 후보가 없으면 안내 글', () => {
  const env = fakeEnv();
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\고치는중.docx', { busy: true }), M('D:\\자료\\표.xlsx')] };
  render(env.root, model({ info, materials: [], outDir: null }), cbs(env));
  addBtn(env).props.onclick();
  const items = env.root.find((n) => String(n.props.class ?? '') === 'tc-item');
  assert.deepEqual(items.map((n) => n.props['aria-disabled']), ['true', undefined]);
  const env2 = fakeEnv();
  render(env2.root, model({ info: { edit: info.edit, open: [] }, materials: [], outDir: null }), cbs(env2));
  addBtn(env2).props.onclick();
  assert.ok(env2.root.find((n) => n.props.class === 'tc-empty').some((n) => n.props.text === '추가할 수 있는 열린 문서가 없어요'));
});

test('render: 목록이 열린 채 다시 그려도 focus()를 부르지 않는다 (처음 열 때만 목록으로 초점)', () => {
  const env = fakeEnv();
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\표.xlsx')] };
  const m = model({ info, materials: [], outDir: null });
  const o = cbs(env);
  render(env.root, m, o);
  addBtn(env).props.onclick();
  assert.equal(env.focused.length, 1);
  assert.equal(env.focused[0].props.class, 'tc-item');
  render(env.root, m, o);
  render(env.root, m, o);
  assert.equal(env.focused.length, 1);
});

test('render: Esc는 목록을 닫고 추가 버튼으로 초점을 돌린다', () => {
  const env = fakeEnv();
  const info = { edit: M('C:\\문서\\공문.hwpx', { current: true }), open: [M('D:\\자료\\표.xlsx')] };
  render(env.root, model({ info, materials: [], outDir: null }), cbs(env));
  addBtn(env).props.onclick();
  env.root.find((n) => String(n.props.class ?? '') === 'tc-menu')[0].props.onkeydown({ key: 'Escape', stopPropagation: noop });
  assert.equal(env.root.find((n) => String(n.props.class ?? '') === 'tc-menu').length, 0);
  assert.equal(env.focused.at(-1).props.class, 'tc-add');
});

test('render: 재료를 빼고 다시 그리면 다음 칩의 ✕로, 마지막이면 앞 칩, 다 빠지면 추가 버튼으로 초점 (A11y)', () => {
  const env = fakeEnv();
  const a = M('D:\\자료\\가.xlsx'); const b = M('D:\\자료\\나.xlsx'); const c = M('D:\\자료\\다.xlsx');
  const edit = M('C:\\문서\\공문.hwpx', { current: true });
  let materials = [a, b, c];
  const o = { ...cbs(env), onRemove: (p) => { materials = materials.filter((x) => x.path !== p); } };
  const draw = () => render(env.root, model({ info: { edit, open: [a, b, c] }, materials, outDir: null }), o);
  const xOf = (name) => env.root.find((n) => n.props.class === 'tc-x' && n.props['aria-label'].startsWith(name))[0];
  draw();
  xOf('나.xlsx').props.onclick();
  draw(); // 화면이 다시 그린다
  assert.equal(env.focused.at(-1).props['aria-label'], '다.xlsx 재료에서 빼기', '다음 칩의 ✕');
  xOf('다.xlsx').props.onclick();
  draw();
  assert.equal(env.focused.at(-1).props['aria-label'], '가.xlsx 재료에서 빼기', '마지막이면 앞 칩');
  xOf('가.xlsx').props.onclick();
  draw();
  assert.equal(env.focused.at(-1).props.class, 'tc-add', '다 빠지면 추가 버튼');
});

test('render: 목록에서 고르고 다시 그리면 추가 버튼에 초점, 초점이 칩 줄 밖으로 나가면 목록이 닫힌다 (A11y)', () => {
  const env = fakeEnv();
  const edit = M('C:\\문서\\공문.hwpx', { current: true });
  const t = M('D:\\자료\\표.xlsx');
  let materials = [];
  const o = { ...cbs(env), onAdd: (p) => { materials = [t].filter((x) => x.path === p); } };
  const draw = () => render(env.root, model({ info: { edit, open: [t] }, materials, outDir: null }), o);
  draw();
  addBtn(env).props.onclick();
  env.root.find((n) => String(n.props.class ?? '') === 'tc-item')[0].props.onclick();
  draw();
  assert.equal(env.focused.at(-1).props.class, 'tc-add');
  // 다시 열고 초점이 밖(다른 단추)으로 나가면 닫힌다. 다시 그려 사라진 단추에서 나간 것(relatedTarget 없음)은 그대로
  materials = [];
  draw();
  addBtn(env).props.onclick();
  const menu = () => env.root.find((n) => String(n.props.class ?? '') === 'tc-menu');
  menu()[0].props.onfocusout({ relatedTarget: null });
  assert.equal(menu().length, 1, '다시 그리기로 초점이 사라진 것은 닫지 않는다');
  menu()[0].props.onfocusout({ relatedTarget: { outside: true } });
  assert.equal(menu().length, 0, '밖으로 나가면 닫는다');
  assert.equal(addBtn(env).props['aria-expanded'], 'false');
});
