// PDF 열기·읽기 (src/core/pdffile.mjs · index.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readPdf } from '../src/core/pdffile.mjs';
import { readDocument, documentInfo, renderPreview, OPENABLE_EXTS, kindOf } from '../src/core/index.mjs';
import { diffDocuments } from '../src/core/diff.mjs';
import { maskDocument, scanDocument } from '../src/core/pii.mjs';
import { ROOT } from './helpers/fixtures.mjs';

const SAMPLE = join(ROOT, 'test', 'fixtures', 'pdf', '신청서.pdf'); // examples/신청서.hwp를 내장 엔진으로 만든 PDF

/** 글자 한 줄(영문)만 든 아주 작은 PDF, text 없으면 빈 쪽 */
function tinyPdf(text) {
  const stream = text ? `BT /F1 24 Tf 72 720 Td (${text}) Tj ET` : '';
  const objs = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let out = '%PDF-1.4\n';
  const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offs.map((n) => `${String(n).padStart(10, '0')} 00000 n \n`).join('')}`;
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

test('PDF는 열 수 있는 형식, 종류 이름은 PDF', () => {
  assert.ok(OPENABLE_EXTS.includes('pdf'));
  assert.equal(kindOf('a.pdf'), 'PDF');
});

test('readPdf: 쪽마다 글자, 한글 그대로', async () => {
  const r = await readPdf(SAMPLE);
  assert.equal(r.pages, 2);
  assert.equal(r.scanned, false);
  assert.match(r.text, /^--- 쪽 1 ---\n장비 대여 신청서/);
  assert.match(r.text, /--- 쪽 2 ---\n유의사항/);
  const d = await readDocument(SAMPLE);
  assert.equal(d.kind, 'PDF');
  assert.equal(d.engine, 'pdf');
  assert.equal(d.pages, 2);
  const info = await documentInfo(SAMPLE);
  assert.equal(info.pages, 2);
  assert.equal(info.fileEngine, true);
});

test('글자 없는 PDF는 scanned, 망가진 파일은 한국어 오류', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-pdf-'));
  try {
    writeFileSync(join(dir, 'empty.pdf'), tinyPdf(''));
    writeFileSync(join(dir, 'hello.pdf'), tinyPdf('Hello Office Kit'));
    writeFileSync(join(dir, 'bad.pdf'), 'not a pdf at all');
    assert.equal((await readPdf(join(dir, 'empty.pdf'))).scanned, true);
    const h = await readPdf(join(dir, 'hello.pdf'));
    assert.match(h.text, /Hello Office Kit/);
    assert.equal(h.scanned, false);
    await assert.rejects(readPdf(join(dir, 'bad.pdf')), /PDF 파일이 아니거나 손상됐습니다/);
    // 바뀐 내용 보기도 PDF끼리 된다
    writeFileSync(join(dir, 'hello2.pdf'), tinyPdf('Hello Office Kit 2'));
    const diff = await diffDocuments(join(dir, 'hello.pdf'), join(dir, 'hello2.pdf'));
    assert.equal(diff.stats.changed, 1);
    // 미리보기는 사본을 내장 PDF 뷰어로
    const p = await renderPreview(join(dir, 'hello.pdf'), join(dir, 'preview'));
    assert.equal(p.kind, 'pdf');
    assert.notEqual(p.file, join(dir, 'hello.pdf'));
    // 개인정보: 찾기는 되고 가린 사본은 안 됨
    writeFileSync(join(dir, 'pii.pdf'), tinyPdf('Mail hong@example.com'));
    assert.equal((await scanDocument(join(dir, 'pii.pdf'))).items[0]?.type, 'email');
    await assert.rejects(maskDocument(join(dir, 'pii.pdf')), /PDF는 글자를 바꿀 수 없어/);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 5 }); }
});
