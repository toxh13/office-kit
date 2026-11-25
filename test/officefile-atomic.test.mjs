// officefile-atomic.test.mjs — 파일 엔진(docx/pptx/xlsx) 저장이 실패해도 원본은 그대로다.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { chmodSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { tempDir } from './helpers/fixtures.mjs';
import { csvCellValue, replaceInFile } from '../src/core/officefile.mjs';

const tmp = tempDir('officekit-atomic-');
after(() => tmp.cleanup());

test('파일 엔진 저장이 실패해도 원본 docx는 그대로다', async () => {
  const src = tmp.path('원본.docx');
  writeFileSync(src, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new TextRun('홍길동 안녕')] })] }] })));
  const out = tmp.path('결과.docx');
  writeFileSync(out, 'ORIGINAL-BYTES');
  chmodSync(out, 0o444);
  try {
    await assert.rejects(
      () => replaceInFile(src, 'docx', [['홍길동', '김철수']], out),
      (e) => /다른 프로그램|권한/.test(e.message) && e.message.includes('결과.docx') && Boolean(e.cause),
    );
  } finally { chmodSync(out, 0o666); }
  assert.equal(readFileSync(out, 'utf8'), 'ORIGINAL-BYTES');
  assert.deepEqual(readdirSync(tmp.dir).filter((f) => f.endsWith('.tmp')), []);
});

test('CSV를 엑셀로 바꿀 때 앞자리 0과 소수 끝자리는 문자열로 지킨다', () => {
  assert.equal(csvCellValue('01012345678'), '01012345678');
  assert.equal(csvCellValue('007'), '007');
  assert.equal(csvCellValue('1.50'), '1.50');
  assert.equal(csvCellValue('0'), 0);
  assert.equal(csvCellValue('-12'), -12);
  assert.equal(csvCellValue('3.14'), 3.14);
  assert.equal(csvCellValue('1234567890123456'), '1234567890123456');
  assert.equal(csvCellValue(''), '');
  assert.equal(csvCellValue('-0'), '-0');
});
