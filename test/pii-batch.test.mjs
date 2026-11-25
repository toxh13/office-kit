// 개인정보 찾기·가리기 (src/core/pii.mjs) · 한글 → HWPX 일괄 변환 (src/core/batch.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { findPii, groupPii, maskDocument, scanDocument, maskedPathFor, piiTargetCheck } from '../src/core/pii.mjs';
import { collectHwpFiles, convertHwpBatch, outputFor, batchReportCsv, planBatchOutputs } from '../src/core/batch.mjs';
import { createDocument, readDocument } from '../src/core/index.mjs';

const EXAMPLES = join(import.meta.dirname, '..', 'examples');
const tmp = () => mkdtempSync(join(tmpdir(), 'officekit-pii-'));
const clean = (d) => rmSync(d, { recursive: true, force: true, maxRetries: 5 });

const SAMPLE = [
  '성명: 홍길동, 주민등록번호: 900101-1234567',
  '외국인 등록번호 850505-5123456',
  '휴대폰 010-1234-5678 / 010 9876 5432 / 01011112222',
  '사무실 (02) 123-4567, 팩스 031-765-4321',
  '메일 gildong.hong@example.co.kr',
  '카드 4111-1111-1111-1111',
  '입금 계좌: 국민은행 123456-01-234567',
  '여권 M12345678',
  '면허 11-12-123456-12',
  // 아닌 것: 날짜·금액·문서 번호·전화 아닌 숫자
  '2026-10-20 회의, 금액 1,234,567원, 문서번호 2026-123-4567, 카드 아님 1234-5678-9012-3456, 코드 12345678901',
].join('\n');

test('findPii: 종류별로 찾고 가린다', () => {
  const f = findPii(SAMPLE);
  const by = (t) => f.filter((x) => x.type === t).map((x) => `${x.value}→${x.masked}`);
  assert.deepEqual(by('rrn'), ['900101-1234567→900101-1******']);
  assert.deepEqual(by('frn'), ['850505-5123456→850505-5******']);
  assert.deepEqual(by('mobile'), ['010-1234-5678→010-****-5678', '010 9876 5432→010 **** 5432', '01011112222→010****2222']);
  assert.deepEqual(by('phone'), ['(02) 123-4567→(02) ***-4567', '031-765-4321→031-***-4321']);
  assert.deepEqual(by('email'), ['gildong.hong@example.co.kr→gi**********@example.co.kr']);
  assert.deepEqual(by('card'), ['4111-1111-1111-1111→4111-****-****-1111']);
  assert.deepEqual(by('account'), ['123456-01-234567→123***-**-***567']);
  assert.deepEqual(by('passport'), ['M12345678→M******78']);
  assert.deepEqual(by('license'), ['11-12-123456-12→11-12-******-12']);
  const last = SAMPLE.split('\n').at(-1);
  assert.deepEqual(findPii(last), [], `오탐: ${JSON.stringify(findPii(last))}`);
  // 가린 글은 다시 찾히지 않는다
  assert.deepEqual(findPii(f.map((x) => x.masked).join('\n')), []);
  // 하이픈 없는 13자리는 검증 숫자가 맞을 때만
  assert.equal(findPii('8001011234561').length, 0);
  assert.equal(findPii('8001011234560')[0]?.type, 'rrn');
  // 종류 고르기
  assert.deepEqual([...new Set(findPii(SAMPLE, { types: ['email'] }).map((x) => x.type))], ['email']);
  const g = groupPii(findPii('010-1234-5678 그리고 다시 010-1234-5678'), '010-1234-5678 그리고 다시 010-1234-5678');
  assert.equal(g.length, 1);
  assert.equal(g[0].count, 2);
});

test('piiTargetCheck: 찾을 때의 문서와 지금 문서가 다르면 PII_DOC_CHANGED', () => {
  const MSG = '문서가 바뀌었습니다. 개인정보 찾기를 다시 하세요.';
  assert.throws(() => piiTargetCheck('C:\\Users\\user\\a.docx', 'C:\\Users\\user\\b.docx'), (e) => e.code === 'PII_DOC_CHANGED' && e.message === MSG && e.expected === true);
  assert.throws(() => piiTargetCheck(undefined, 'C:\\Users\\user\\a.docx'), (e) => e.code === 'PII_DOC_CHANGED' && e.message === MSG);
  assert.throws(() => piiTargetCheck('', 'C:\\Users\\user\\a.docx'), (e) => e.code === 'PII_DOC_CHANGED');
  assert.doesNotThrow(() => piiTargetCheck('C:\\Users\\user\\a.docx', 'C:\\Users\\user\\a.docx'));
  assert.doesNotThrow(() => piiTargetCheck('C:\\Users\\user\\A.DOCX', 'c:\\users\\user\\a.docx'));
});

test('maskDocument: hwpx·docx·xlsx 가린 사본, 원본은 그대로', async () => {
  const dir = tmp();
  try {
    for (const ext of ['hwpx', 'docx', 'xlsx']) {
      const src = join(dir, `명단.${ext}`);
      await createDocument(src, ext === 'xlsx'
        ? '| 이름 | 연락처 | 주민번호 |\n|---|---|---|\n| 홍길동 | 010-1234-5678 | 900101-1234567 |\n| 김철수 | 010-2222-3333 | 851212-2345678 |'
        : '# 명단\n\n홍길동 010-1234-5678 900101-1234567\n\n| 이름 | 메일 |\n|---|---|\n| 김철수 | kim@example.com |');
      const before = readFileSync(src);
      const scan = await scanDocument(src);
      assert.ok(scan.items.length >= 2, `${ext}: ${JSON.stringify(scan.items)}`);
      const r = await maskDocument(src);
      assert.equal(r.output, join(dir, `명단_가림.${ext}`));
      assert.deepEqual(r.remaining, [], `${ext}: 남은 것 ${JSON.stringify(r.remaining)}`);
      assert.ok(r.items.every((it) => it.count >= 1), `${ext}: ${JSON.stringify(r.items)}`);
      const text = (await readDocument(r.output)).text;
      assert.doesNotMatch(text, /010-1234-5678|900101-1234567/);
      assert.match(text, /010-\*\*\*\*-5678/);
      assert.deepEqual(readFileSync(src), before, '원본은 바뀌지 않는다');
      assert.equal(maskedPathFor(src), join(dir, `명단_가림 (2).${ext}`));
      // 일부만 고르기
      const one = await maskDocument(src, { output: join(dir, `하나.${ext}`), values: ['010-1234-5678'] });
      const t1 = (await readDocument(one.output)).text;
      assert.doesNotMatch(t1, /010-1234-5678/);
      assert.match(t1, /900101-1234567/);
    }
    await assert.rejects(maskDocument(join(dir, '명단.docx'), { output: join(dir, '명단.docx') }), /덮어쓰지/);
  } finally { clean(dir); }
});

test('일괄 변환: 같은 이름 결과는 덮어쓰지 않고 (2), (3)을 붙인다', async () => {
  const dir = tmp();
  try {
    for (const f of ['A', 'B', 'C']) mkdirSync(join(dir, f));
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, 'A', '보고서.hwp'));
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, 'B', '보고서.hwp'));
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, 'C', '보고서.hwp'));
    const { items } = collectHwpFiles([join(dir, 'A', '보고서.hwp'), join(dir, 'B', '보고서.hwp'), join(dir, 'C', '보고서.hwp')]);
    const out = join(dir, '모음');
    const r = await convertHwpBatch(items, { outDir: out, overwrite: true });
    assert.deepEqual(r.results.map((x) => x.status), ['ok', 'ok', 'ok'], JSON.stringify(r.results));
    assert.equal(new Set(r.results.map((x) => x.out.toLowerCase())).size, 3);
    assert.ok(existsSync(join(out, '보고서.hwpx')));
    assert.ok(existsSync(join(out, '보고서 (2).hwpx')));
    assert.ok(existsSync(join(out, '보고서 (3).hwpx')));
    assert.ok(r.results[1].notes.includes('이름이 겹쳐 ‘보고서 (2).hwpx’로 저장'));
    assert.ok(r.results[2].notes.includes('이름이 겹쳐 ‘보고서 (3).hwpx’로 저장'));
    assert.ok(!r.results[0].notes.some((n) => n.includes('이름이 겹쳐')));
  } finally { clean(dir); }
});

test('일괄 변환: 디스크에 이미 있는 다른 \'보고서 (2).hwpx\'는 덮어쓰지 않고 (3)을 붙인다', async () => {
  const dir = tmp();
  try {
    for (const f of ['A', 'B']) mkdirSync(join(dir, f));
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, 'A', '보고서.hwp'));
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, 'B', '보고서.hwp'));
    const out = join(dir, '모음');
    mkdirSync(out);
    const unrelated = join(out, '보고서 (2).HWPX'); // 대소문자만 다른 이름도 같은 파일로 본다 (Windows)
    writeFileSync(unrelated, '상관없는 파일');
    const { items } = collectHwpFiles([join(dir, 'A', '보고서.hwp'), join(dir, 'B', '보고서.hwp')]);
    const plan = planBatchOutputs(items, out);
    const second = process.platform === 'win32' ? '보고서 (3).hwpx' : '보고서 (2).hwpx';
    assert.deepEqual(plan.map((p) => p.out), [join(out, '보고서.hwpx'), join(out, second)]);
    assert.deepEqual(plan.map((p) => p.renamed), [false, true]);
    const r = await convertHwpBatch(items, { outDir: out, overwrite: true });
    assert.deepEqual(r.results.map((x) => x.status), ['ok', 'ok'], JSON.stringify(r.results));
    assert.equal(r.results[1].out, join(out, second));
    assert.equal(readFileSync(unrelated, 'utf8'), '상관없는 파일', '다른 파일은 그대로');
    assert.ok(r.results[1].notes.includes(`이름이 겹쳐 ‘${second}’로 저장`));
  } finally { clean(dir); }
});

test('일괄 변환: 결과 이름 계획 — 원래 이름은 있어도 그대로 (덮어쓰기는 overwrite가 정한다)', () => {
  const dir = tmp();
  try {
    const a = join(dir, 'a.hwp');
    writeFileSync(join(dir, 'a.hwpx'), 'old');
    writeFileSync(join(dir, 'a (2).hwpx'), 'other');
    const plan = planBatchOutputs([{ src: a, root: dir }], undefined);
    assert.deepEqual(plan, [{ out: join(dir, 'a.hwpx'), renamed: false }]);
  } finally { clean(dir); }
});

test('일괄 변환: 폴더(하위 포함)의 hwp → hwpx, 확인 결과·건너뛰기·보고서', async () => {
  const dir = tmp();
  try {
    mkdirSync(join(dir, '하위', '더 깊이'), { recursive: true });
    copyFileSync(join(EXAMPLES, '신청서.hwp'), join(dir, '신청서.hwp'));
    copyFileSync(join(EXAMPLES, '신청서_채움.hwp'), join(dir, '하위', '더 깊이', '채움.hwp'));
    writeFileSync(join(dir, '하위', '깨진.hwp'), 'not a hwp');
    writeFileSync(join(dir, '~$임시.hwp'), 'x');
    const { items } = collectHwpFiles([dir]);
    assert.equal(items.length, 3);
    assert.equal(collectHwpFiles([dir], { recursive: false }).items.length, 1);
    const progress = [];
    const r = await convertHwpBatch(items, { onProgress: (p) => progress.push(p.index) });
    assert.deepEqual(progress, [0, 1, 2]);
    assert.deepEqual({ ok: r.summary.ok, fail: r.summary.fail }, { ok: 2, fail: 1 }, JSON.stringify(r.results));
    assert.ok(existsSync(join(dir, '신청서.hwpx')));
    assert.ok(existsSync(join(dir, '하위', '더 깊이', '채움.hwpx')));
    assert.ok(!existsSync(join(dir, '하위', '깨진.hwpx')));
    const ok = r.results.find((x) => x.src.endsWith('신청서.hwp'));
    assert.deepEqual(ok.pages, [2, 2]);
    // 다시 하면 건너뛴다, 덮어쓰기면 다시 만든다
    const again = await convertHwpBatch(items.slice(0, 1));
    assert.equal(again.results[0].status, 'skipped');
    assert.equal((await convertHwpBatch(items.slice(0, 1), { overwrite: true })).results[0].status, 'ok');
    // 다른 폴더에 모으기: 하위 폴더 구조 유지
    const out = join(dir, '모음');
    assert.equal(outputFor(items.find((x) => x.src.endsWith('채움.hwp')), out), join(out, '하위', '더 깊이', '채움.hwpx'));
    const csv = batchReportCsv(r.results);
    assert.ok(csv.startsWith('﻿"결과","원본","HWPX"'));
    assert.match(csv, /"같음"/);
    assert.match(csv, /"실패"/);
    // 취소
    let n = 0;
    const c = await convertHwpBatch(items, { overwrite: true, isCancelled: () => n++ >= 1 });
    assert.equal(c.summary.cancelled, true);
    assert.equal(c.results.length, 1);
  } finally { clean(dir); }
});
