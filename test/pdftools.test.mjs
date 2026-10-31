// PDF 도구 (src/core/pdftools.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { parsePages, mergePdfs, extractPages, deletePages, rotatePages, pdfPageCount } from '../src/core/pdftools.mjs';
import { readPdf } from '../src/core/pdffile.mjs';
import { ROOT } from './helpers/fixtures.mjs';

const SAMPLE = join(ROOT, 'test', 'fixtures', 'pdf', '신청서.pdf'); // 2쪽

async function blankPdf(path, n) {
  const d = await PDFDocument.create();
  for (let i = 0; i < n; i++) d.addPage([300 + i, 400]); // 쪽마다 너비가 달라 순서를 알 수 있다
  writeFileSync(path, await d.save());
}
const widths = async (p) => (await PDFDocument.load(readFileSync(p))).getPages().map((x) => Math.round(x.getWidth()));

test('parsePages: 범위·낱쪽·끝까지·전체, 틀리면 한국어 오류', () => {
  assert.deepEqual(parsePages('1-3,5', 6), [0, 1, 2, 4]);
  assert.deepEqual(parsePages('5-', 6), [4, 5]);
  assert.deepEqual(parsePages('-2', 6), [0, 1]);
  assert.deepEqual(parsePages('', 3), [0, 1, 2]);
  assert.deepEqual(parsePages('전체', 2), [0, 1]);
  assert.deepEqual(parsePages('3-1', 3), [2, 1, 0], '거꾸로 적으면 거꾸로');
  assert.deepEqual(parsePages(' 2 쪽 ', 3), [1]);
  assert.throws(() => parsePages('7', 6), /범위를 벗어났습니다/);
  assert.throws(() => parsePages('a-b', 6), /알아볼 수 없습니다/);
});

test('합치기·쪽 뽑기·쪽 지우기·돌리기, 원본은 그대로', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-pdft-'));
  try {
    const a = join(dir, 'a.pdf'); const b = join(dir, 'b.pdf');
    await blankPdf(a, 3); await blankPdf(b, 2);
    const aBefore = readFileSync(a);
    const m = await mergePdfs([a, b, SAMPLE], join(dir, '합친.pdf'));
    assert.equal(m.pages, 7);
    assert.deepEqual((await widths(m.output)).slice(0, 5), [300, 301, 302, 300, 301]);
    assert.match((await readPdf(m.output)).text, /장비 대여 신청서/, '글자도 그대로');
    const e = await extractPages(m.output, '2,4-5', join(dir, '뽑은.pdf'));
    assert.equal(e.pages, 3);
    assert.deepEqual(await widths(e.output), [301, 300, 301]);
    const d = await deletePages(a, '1', join(dir, '지운.pdf'));
    assert.deepEqual(await widths(d.output), [301, 302]);
    const r = await rotatePages(a, 90, '2', join(dir, '돌린.pdf'));
    const rot = (await PDFDocument.load(readFileSync(r.output))).getPages().map((p) => p.getRotation().angle);
    assert.deepEqual(rot, [0, 90, 0]);
    assert.equal(await pdfPageCount(a), 3);
    assert.deepEqual(readFileSync(a), aBefore, '원본은 그대로');
    await assert.rejects(mergePdfs([a], join(dir, 'x.pdf')), /둘 이상/);
    await assert.rejects(extractPages(a, '1', a), /원본과 다른 이름/);
    await assert.rejects(deletePages(a, '1-3', join(dir, 'y.pdf')), /모든 쪽을 지울 수는 없습니다/);
    await assert.rejects(rotatePages(a, 45, '', join(dir, 'z.pdf')), /90, 180, 270/);
    writeFileSync(join(dir, 'bad.pdf'), 'nope');
    await assert.rejects(extractPages(join(dir, 'bad.pdf'), '1', join(dir, 'w.pdf')), /PDF 파일이 아니거나 손상됐습니다|PDF를 읽지 못했습니다/);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 5 }); }
});
