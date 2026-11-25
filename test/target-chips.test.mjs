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
