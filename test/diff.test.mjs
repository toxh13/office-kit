// 변경 내용 보기 (src/core/diff.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, copyFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { myers, diffUnits, diffSheets, charSegments, diffDocuments } from '../src/core/diff.mjs';
import { createDocument, replaceText, setCells } from '../src/core/index.mjs';

// 편집 순서를 따라가면 a가 b가 되는지
function apply(a, b, ops) {
  const out = [];
  for (const o of ops) {
    if (o[0] === '=') { assert.equal(a[o[1]], b[o[2]]); out.push(a[o[1]]); } else if (o[0] === '+') out.push(b[o[1]]);
  }
  return out;
}

test('myers: 무작위 배열에서 늘 올바른 최소 편집', () => {
  let seed = 7;
  const rnd = (n) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  for (let t = 0; t < 300; t++) {
    const a = Array.from({ length: rnd(15) }, () => 'abcd'[rnd(4)]);
    const b = Array.from({ length: rnd(15) }, () => 'abcd'[rnd(4)]);
    const ops = myers(a, b);
    assert.deepEqual(apply(a, b, ops), b);
    assert.equal(ops.filter((o) => o[0] === '-').length, a.length - ops.filter((o) => o[0] === '=').length);
  }
  assert.equal(myers(['a', 'b'], ['c', 'd'], { maxD: 1 }), null);
});

test('diffUnits: 바뀐 문단은 글자 단위로, 더하고 지운 문단은 따로', () => {
  const U = (s) => s.split('\n').map((text) => ({ text }));
  const d = diffUnits(U('제목\n첫째 문단입니다.\n둘째 문단\n셋째\n끝'), U('제목\n첫째 문단이에요.\n셋째\n새 문단\n끝'));
  assert.deepEqual(d.stats, { added: 1, removed: 1, changed: 1 });
  const items = d.hunks.flatMap((h) => h.items);
  const mod = items.find((x) => x.type === 'mod');
  assert.equal(mod.old, '첫째 문단입니다.');
  assert.deepEqual(mod.segs.filter((s) => s.t !== '=').map((s) => s.t + s.s), ['-입니다', '+이에요']);
  assert.ok(items.some((x) => x.type === 'del' && x.old === '둘째 문단'));
  assert.ok(items.some((x) => x.type === 'add' && x.new === '새 문단'));
  assert.equal(diffUnits(U('a\nb'), U('a\nb')).hunks.length, 0);
  assert.equal(diffUnits(U('a\nb'), U('a\n\n \nb')).hunks.length, 0, '빈 줄만 더한 것은 바뀐 것으로 보지 않는다');
  assert.deepEqual(charSegments('가나다', '가다').map((s) => s.t + s.s), ['=가', '-나', '=다']);
});

test('diffSheets: 칸 주소와 시트 추가·삭제', () => {
  const d = diffSheets(
    [{ name: '매출', rows: [['항목', '금액'], ['A', '100']] }, { name: '옛 시트', rows: [] }],
    [{ name: '매출', rows: [['항목', '금액'], ['A', '120'], ['B', '50']] }, { name: '새 시트', rows: [] }],
  );
  assert.deepEqual(d.cells.map((c) => `${c.type} ${c.sheet}!${c.cell} ${c.old}→${c.new}`), ['mod 매출!B2 100→120', 'add 매출!A3 →B', 'add 매출!B3 →50']);
  assert.deepEqual(d.sheetsAdded, ['새 시트']);
  assert.deepEqual(d.sheetsRemoved, ['옛 시트']);
});

test('diffDocuments: 실제 hwpx·docx·xlsx 파일', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-diff-'));
  try {
    for (const ext of ['hwpx', 'docx']) {
      const a = join(dir, `a.${ext}`); const b = join(dir, `b.${ext}`);
      await createDocument(a, '# 회의록\n\n참석자는 김철수입니다.\n\n다음 회의는 월요일.');
      copyFileSync(a, b);
      await replaceText(b, [['김철수', '이영희'], ['월요일', '화요일']], { engine: 'file' });
      const d = await diffDocuments(a, b);
      assert.equal(d.kind, 'text');
      assert.equal(d.same, false);
      assert.equal(d.stats.changed, 2, `${ext}: ${JSON.stringify(d.hunks)}`);
      assert.equal((await diffDocuments(a, a)).same, true);
    }
    const a = join(dir, 'a.xlsx'); const b = join(dir, 'b.xlsx');
    await createDocument(a, '| 항목 | 금액 |\n|---|---|\n| 사과 | 100 |');
    copyFileSync(a, b);
    await setCells(b, { B2: 150 }, { engine: 'file' });
    const d = await diffDocuments(a, b);
    assert.equal(d.kind, 'cells');
    assert.deepEqual(d.cells.map((c) => `${c.cell}:${c.old}→${c.new}`), ['B2:100→150']);
    await assert.rejects(diffDocuments(a, join(dir, 'a.docx')), /같은 형식/);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 5 }); }
});
