// excel.test.mjs — Excel 전용 살펴보기·편집 (excel_inspect / excel_edit, docs/FEATURES.md 'Excel').
// 파일 엔진(exceljs) 시험은 늘 돈다. OFFICEKIT_COM_TESTS=1이면 파일 엔진 결과를 Excel로 열어
// (경고를 끈 Excel은 고쳐야 하는 파일을 열지 못한다) 복구 없이 받아들이는지도 확인하고, Excel(COM) 엔진 시험도 한다.
import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync, writeFileSync, copyFileSync } from 'node:fs';
import { inflateSync } from 'node:zlib';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { excelEdit, excelInspect, inspectText, normalizeOperations, EXCEL_OPS, _formula } from '../src/core/excel.mjs';
import { pageImage } from '../src/core/render.mjs';
import { COM_TESTS, ROOT, tempDir, isPng, xlsxWithChart } from './helpers/fixtures.mjs';

const tmp = tempDir('officekit-excel-');
after(() => tmp.cleanup());

const VT = ExcelJS.ValueType;
const COM = { timeout: 600_000 };

// 시험용 통합 문서: '매출'(머리글 + 5개 지점 + 합계 수식), '목록'(매출을 가리키는 수식)
async function makeBook(path) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('매출');
  ws.addRows([['지점', '1월', '2월', '합계'], ['서울', 10, 20], ['부산', 30, 40], ['대구', 50, 60], ['인천', 70, 80], ['광주', 90, 100]]);
  for (let r = 2; r <= 6; r++) ws.getCell(`D${r}`).value = { formula: `SUM(B${r}:C${r})` };
  ws.getCell('A7').value = '합계';
  for (const c of ['B', 'C', 'D']) ws.getCell(`${c}7`).value = { formula: `SUM(${c}2:${c}6)` };
  const ls = wb.addWorksheet('목록');
  ls.getCell('A1').value = '총합';
  ls.getCell('B1').value = { formula: '매출!D7' };
  ls.getCell('B2').value = { formula: "SUM('매출'!B2:B6)" };
  ls.getCell('B3').value = { formula: '매출!A3' };
  await wb.xlsx.writeFile(path);
  return path;
}
let seq = 0;
const fresh = async (name = 'book') => makeBook(tmp.path(`${name}-${++seq}.xlsx`));
// exceljs는 표 열 안의 자식 요소(계산 열 수식)를 잘못 읽으므로 빼고 읽는다 (excel.mjs의 forExceljs와 같음)
async function load(path) {
  const zip = await JSZip.loadAsync(readFileSync(path));
  for (const f of Object.keys(zip.files).filter((x) => /^xl\/tables\/.+\.xml$/.test(x))) {
    zip.file(f, (await zip.file(f).async('string')).replace(/<(tableColumn)\b([^>]*?)(?<!\/)>[\s\S]*?<\/\1>/g, '<$1$2/>'));
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await zip.generateAsync({ type: 'nodebuffer' }));
  return wb;
}
const formulaOf = (cell) => (cell.type === VT.Formula ? `=${cell.formula}` : null);
const zipText = async (path, name) => (await JSZip.loadAsync(readFileSync(path))).file(name)?.async('string');

/** 파일 엔진으로 고치고 exceljs로 다시 읽기 */
async function fileEdit(ops, { name = 'f', src } = {}) {
  src ??= await fresh(name);
  const out = tmp.path(`${name}-${++seq}-out.xlsx`);
  const r = await excelEdit(src, ops, { output: out, engine: 'file' });
  assert.equal(r.engine, 'file');
  assert.equal(r.results.length, ops.length);
  for (const line of r.results) assert.match(line, /[가-힣]/);
  return { src, out, r, wb: await load(out) };
}

/** OFFICEKIT_COM_TESTS=1이면 Excel이 복구 없이 여는지 확인 */
async function excelAccepts(path) {
  if (!COM_TESTS) return null;
  const info = await excelInspect(path, { engine: 'app' });
  assert.equal(info.engine, 'excel');
  assert.ok(info.sheets.length >= 1);
  return info;
}

// ── PNG 읽기 (8비트 RGB/RGBA) — 쪽 그림에서 색 세기 ──────────
function decodePng(buf) {
  let pos = 8; const idat = []; let w; let h; let ct;
  while (pos < buf.length) {
    const len = buf.readUInt32BE(pos); const type = buf.toString('ascii', pos + 4, pos + 8);
    const data = buf.subarray(pos + 8, pos + 8 + len);
    if (type === 'IHDR') { w = data.readUInt32BE(0); h = data.readUInt32BE(4); ct = data[9]; assert.equal(data[8], 8, '8비트 PNG만'); }
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    pos += 12 + len;
  }
  const bpp = { 0: 1, 2: 3, 4: 2, 6: 4 }[ct];
  const raw = inflateSync(Buffer.concat(idat));
  const stride = w * bpp; const out = Buffer.alloc(h * stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)];
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[y * stride + x - bpp] : 0;
      const b = y ? out[(y - 1) * stride + x] : 0;
      const c = x >= bpp && y ? out[(y - 1) * stride + x - bpp] : 0;
      let v = raw[y * (stride + 1) + 1 + x];
      if (f === 1) v += a; else if (f === 2) v += b; else if (f === 3) v += (a + b) >> 1;
      else if (f === 4) { const p = a + b - c; const pa = Math.abs(p - a); const pb = Math.abs(p - b); const pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      out[y * stride + x] = v & 255;
    }
  }
  return { width: w, height: h, bpp, data: out };
}
function countPixels(img, pred) {
  let n = 0;
  for (let i = 0; i < img.data.length; i += img.bpp) if (pred(img.data[i], img.data[i + 1], img.data[i + 2])) n++;
  return n;
}
const near = (r, g, b, hex, tol = 10) => Math.abs(r - parseInt(hex.slice(0, 2), 16)) <= tol && Math.abs(g - parseInt(hex.slice(2, 4), 16)) <= tol
  && Math.abs(b - parseInt(hex.slice(4, 6), 16)) <= tol;

// ── 수식 참조 옮기기 (단위) ───────────────────────────────

describe('수식 참조 옮기기', () => {
  const { shiftFormula, mover, slideFormula, renameInFormula, dropSheetInFormula, dropTablesInFormula, scanRefs } = _formula;
  const S = (f, mv, home = 'S') => shiftFormula(f, home, 'S', mv);

  test('행 넣기·지우기: 범위는 늘고 줄며, 지운 칸은 #REF!', () => {
    const ins = mover('row', 3, 2, false); const del = mover('row', 3, 2, true);
    assert.equal(S('SUM(A1:A10)', ins), 'SUM(A1:A12)');
    assert.equal(S('SUM(A1:A10)', del), 'SUM(A1:A8)');
    assert.equal(S('A3+B4*$C$5', ins), 'A5+B6*$C$7');
    assert.equal(S('A3+B4*$C$5', del), '#REF!+#REF!*$C$3');
    assert.equal(S('SUM(3:4)+SUM(5:9)+SUM(A:A)', del), 'SUM(#REF!)+SUM(3:7)+SUM(A:A)');
  });

  test('열 넣기·지우기', () => {
    assert.equal(S('VLOOKUP(B2,$D$2:$F$20,2,FALSE)', mover('col', 2, 1, false)), 'VLOOKUP(C2,$E$2:$G$20,2,FALSE)');
    assert.equal(S('VLOOKUP(B2,$D$2:$F$20,2,FALSE)', mover('col', 2, 1, true)), 'VLOOKUP(#REF!,$C$2:$E$20,2,FALSE)');
    assert.equal(S('SUM(B:B)+A1', mover('col', 2, 1, true)), 'SUM(#REF!)+A1');
  });

  test('문자열·함수 이름·구조적 참조·다른 시트·다른 파일은 그대로', () => {
    const ins = mover('row', 3, 2, false);
    assert.equal(S('LOG10(A5)+"A5"&T1+Table1[Col]+T1', ins), 'LOG10(A7)+"A5"&T1+Table1[Col]+T1');
    assert.equal(S("'My Sheet'!A5+S!A5+매출!B10", ins, 'X'), "'My Sheet'!A5+S!A7+매출!B10");
    assert.equal(S('[1]S!A5+A5+1E+10', ins), '[1]S!A5+A7+1E+10');
    assert.equal(S('#REF!A5+A5', ins), '#REF!A5+A7');
    const toks = scanRefs("SUM('시트 1'!$A$1:B2,매출!C:C,3:4)");
    assert.deepEqual(toks.map((t) => [t.sheet, t.ref.kind]), [['시트 1', 'area'], ['매출', 'cols'], [null, 'rows']]);
  });

  test('공유 수식 풀기·시트 이름 바꾸기·시트·표 지우기', () => {
    assert.equal(slideFormula('A1*$B$1+SUM(C1:C3)+B$2+$C3+D:D', 2, 1), 'B3*$B$1+SUM(D3:D5)+C$2+$C5+E:E');
    assert.equal(renameInFormula("S!A1+'s'!B2+SS!A1+A1", 'S', 'New Name'), "'New Name'!A1+'New Name'!B2+SS!A1+A1");
    assert.equal(renameInFormula('S!A1', 'S', '매출'), '매출!A1');
    assert.equal(dropSheetInFormula("S!A1+'s'!B2:C3+X!A1", 'S'), '#REF!+#REF!+X!A1');
    assert.equal(dropTablesInFormula('SUM(표1[수량])+표10[x]+"표1[a]"', ['표1']), 'SUM(#REF!)+표10[x]+"표1[a]"');
  });
});

// ── 작업 검증 (두 엔진 공통) ──────────────────────────────

describe('작업 검증', () => {
  const bad = (ops, re) => assert.throws(() => normalizeOperations(ops), re);

  test('한국어 오류: 몇 번째 작업인지와 이유', () => {
    bad([{ op: 'set_cells', cells: { A1: 1 } }, { op: 'nope' }], /2번째 작업\(nope\).*알 수 없는 작업/);
    bad([{ op: 'set_cells', cells: { 'A0': 1 } }], /셀 주소가 올바르지 않습니다: A0/);
    bad([{ op: 'format', range: 'A1', color: '빨강빨강' }], /색이 올바르지 않습니다/);
    bad([{ op: 'format', range: 'A1' }], /바꿀 서식이 없습니다/);
    bad([{ op: 'merge', range: 'B2' }], /두 칸 이상/);
    bad([{ op: 'add_sheet', name: 'a/b' }], /쓸 수 없는 글자/);
    bad([{ op: 'add_sheet', name: 'x'.repeat(32) }], /31자/);
    bad([{ op: 'sort', range: 'A1:B5', by: 'D' }], /범위 A1:B5 안에 없습니다/);
    bad([{ op: 'conditional_format', range: 'A1:A3', rule: { type: 'between', value: 1 } }], /value2/);
    bad([{ op: 'data_validation', range: 'A1', type: 'whole' }], /min 또는 max/);
    bad([{ op: 'data_validation', range: 'A1', list: ['a,b'] }], /쉼표/);
    bad([{ op: 'named_range', name: 'A1', range: 'S!A1' }], /이름으로 쓸 수 없습니다/);
    bad([{ op: 'column_width', columns: 'A:C', width: 300 }], /0~255/);
    bad([], /비어 있습니다/);
  });

  test('정리: 주소 대문자, 열 문자→번호, 색 #RRGGBB', () => {
    const [a, b, c, d] = normalizeOperations([
      { op: 'set_cells', cells: { b2: 1 } }, { op: 'insert_columns', at: 'c', count: 2 },
      { op: 'format', range: 'a1:b2', fill: '#abc', color: 'red' }, { op: 'named_range', name: '합계', range: "'내 시트'!B2:C3" },
    ]);
    assert.deepEqual(a.cells, { B2: 1 });
    assert.deepEqual([b.at, b.columns], [3, 'C:D']);
    assert.deepEqual([c.range, c.fill, c.color], ['A1:B2', 'AABBCC', 'FF0000']);
    assert.deepEqual([d.sheet, d.area], ['내 시트', '$B$2:$C$3']);
  });

  test('작업 표: 파일 엔진에서 ×인 작업', () => {
    assert.deepEqual(Object.entries(EXCEL_OPS).filter(([, v]) => !v.file).map(([k]) => k).sort(), ['add_chart', 'copy_sheet', 'recalculate', 'sort']);
  });

  test('수식 문법 확인: Excel이 열지 못하는 파일이 되는 수식은 미리 거절 (두 엔진 공통)', () => {
    const f = (v) => normalizeOperations([{ op: 'set_cells', cells: { A1: v } }]);
    bad([{ op: 'set_cells', cells: { F1: '=SUMM((1' } }], /F1: 수식이 올바르지 않습니다: =SUMM\(\(1 — 여는 괄호.*앞에 ' 를 붙이세요/);
    bad([{ op: 'set_cells', cells: { A1: '=SUM(A1;A2)' } }], /쉼표\(,\)로 구분하세요/);
    bad([{ op: 'set_cells', cells: { A1: '="abc' } }], /큰따옴표\("\)가 닫히지 않았습니다/);
    bad([{ op: 'set_cells', cells: { A1: "='내 시트!A1" } }], /작은따옴표/);
    bad([{ op: 'set_cells', cells: { A1: '=A1+' } }], /연산자/);
    bad([{ op: 'set_cells', cells: { A1: '=SUM(A1:)' } }], /범위\(:\)의 앞이나 뒤가 비어/);
    bad([{ op: 'set_cells', cells: { A1: '=' } }], /빈 수식/);
    bad([{ op: 'set_cells', cells: { A1: '=SUM(1))' } }], /닫는 괄호/);
    bad([{ op: 'set_range', range: 'A1', values: [['=1+(2']] }], /values\[0\]\[0\]: 수식이 올바르지 않습니다/);
    bad([{ op: 'conditional_format', range: 'A1:A3', rule: { type: 'greater_than', value: '=bad((' } }], /rule\.value: 수식이 올바르지 않습니다/);
    bad([{ op: 'data_validation', range: 'A1', list: '=$Z$1:' }], /list: 수식이 올바르지 않습니다/);
    bad([{ op: 'data_validation', range: 'A1', list: ['5" 모니터', '7'] }], /큰따옴표\("\)를 쓸 수 없습니다.*셀 범위/);
    // 맞는 수식은 그대로
    for (const ok of ['=SUM(A1,A2)', '={1,2;3,4}', '=IF(A1=";",1,2)', "='내 시트'!A1+표1[[#This Row],[수량]]", '=A1%', '=-A1', '=SUMM(1)', '=SUM(A:A)', '=IF(A1,,1)', '="it\'s"']) {
      assert.equal(f(ok)[0].cells.A1, ok);
    }
    // 글자는 확인하지 않음
    assert.equal(f("'=SUM((")[0].cells.A1, "'=SUM((");
    assert.equal(normalizeOperations([{ op: 'conditional_format', range: 'A1', rule: { type: 'text_contains', value: '=(' } }])[0].rule.value, '=(');
  });

  test('모르는 인수는 거절 (조용히 버리지 않음)', () => {
    bad([{ op: 'conditional_format', range: 'A1:A3', rule: { type: 'greater_than', value: 1, color: '#00FF00' } }], /모르는 인수입니다: color — rule에 쓸 수 있는 인수: type, value, value2, fill, font_color/);
    bad([{ op: 'add_table', range: 'A1:B3', header: false, total_row: true }], /모르는 인수입니다: header, total_row — add_table에 쓸 수 있는 인수: sheet, range, name, style/);
    bad([{ op: 'recalculate', sheet: 1 }], /모르는 인수입니다: sheet/);
    bad([{ op: 'add_chart', range: 'A1:B3', type: 'column', series_in: 'diag' }], /series_in/);
    // 명령줄에서 받던 다른 이름은 그대로
    assert.equal(normalizeOperations([{ op: 'column_width', range: 'A:B', width: 10 }])[0].columns, 'A:B');
    assert.equal(normalizeOperations([{ op: 'add_chart', range: 'A1:B3', type: 'pie', series_in: 'rows' }])[0].series_in, 'rows');
  });
});

// ── 파일 엔진 (exceljs) ─────────────────────────────────

describe('파일 엔진: 시트 관리', () => {
  test('add_sheet(after) · move_sheet · rename_sheet(수식 참조도) · delete_sheet(#REF!)', async () => {
    const { out, wb, r } = await fileEdit([
      { op: 'add_sheet', name: '요약', after: '매출' },
      { op: 'add_sheet', name: '임시' },
      { op: 'move_sheet', sheet: '임시', position: 1 },
      { op: 'rename_sheet', sheet: '매출', name: '매출 2026' },
      { op: 'delete_sheet', sheet: '임시' },
    ]);
    assert.deepEqual(wb.worksheets.map((w) => w.name), ['매출 2026', '요약', '목록']);
    const ls = wb.getWorksheet('목록');
    assert.equal(formulaOf(ls.getCell('B1')), "='매출 2026'!D7");
    assert.equal(formulaOf(ls.getCell('B2')), "=SUM('매출 2026'!B2:B6)");
    assert.match(r.results[3], /'매출' → '매출 2026'/);
    // 활성 시트는 그대로 첫 시트(매출 2026)
    assert.equal((await excelInspect(out, { engine: 'file' })).activeSheet, '매출 2026');
    await excelAccepts(out);
  });

  test('지운 시트를 가리키던 수식은 #REF!, 보이는 시트가 하나면 못 지움', async () => {
    const { wb, out } = await fileEdit([{ op: 'delete_sheet', sheet: '매출' }]);
    assert.equal(formulaOf(wb.getWorksheet('목록').getCell('B1')), '=#REF!');
    await excelAccepts(out);
    await assert.rejects(fileEdit([{ op: 'delete_sheet', sheet: 1 }, { op: 'delete_sheet', sheet: 1 }]), /2번째 작업\(delete_sheet\) 실패: 보이는 시트가 하나뿐/);
    await assert.rejects(fileEdit([{ op: 'add_sheet', name: '목록' }]), /같은 이름의 시트가 이미 있습니다: 목록/);
    await assert.rejects(fileEdit([{ op: 'rename_sheet', sheet: '없는시트', name: 'x' }]), /시트 없음: 없는시트 \(있는 시트: '매출', '목록'\)/);
  });

  test('시트는 번호(1부터)로도', async () => {
    const { wb } = await fileEdit([{ op: 'set_cells', sheet: 2, cells: { C1: '둘째' } }]);
    assert.equal(wb.getWorksheet('목록').getCell('C1').value, '둘째');
  });
});

describe('파일 엔진: 셀', () => {
  test('set_cells: 숫자·글자(숫자처럼 보여도 글자)·수식·true/false·null·작은따옴표', async () => {
    const { wb, out, r } = await fileEdit([{ op: 'set_cells', sheet: '매출', cells: {
      F1: 120, F2: '120', F3: '=F1*2', F4: true, B2: null, F5: "'=글자", F6: '1/2',
    } }]);
    const ws = wb.getWorksheet('매출');
    assert.equal(ws.getCell('F1').value, 120);
    assert.equal(ws.getCell('F2').value, '120');
    assert.equal(formulaOf(ws.getCell('F3')), '=F1*2');
    assert.equal(ws.getCell('F4').value, true);
    assert.equal(ws.getCell('B2').value, null);
    assert.equal(ws.getCell('F5').value, '=글자');
    assert.equal(ws.getCell('F6').value, '1/2');
    assert.ok(r.notes.some((n) => /다시 계산/.test(n)));
    // 다시 열 때 계산하도록
    assert.match(await zipText(out, 'xl/workbook.xml'), /fullCalcOnLoad="1"/);
    if (COM_TESTS) {
      const { readDocument } = await import('../src/core/index.mjs');
      const rows = (await readDocument(out, { engine: 'app' })).sheets[0].rows;
      assert.equal(rows[2][5], '240');
      assert.equal(rows[6][1], '240', '합계가 다시 계산됨 (서울 1월 10을 지움)');
    }
  });

  test('set_range: 2차원 붙여 넣기, 짧은 행은 빈 칸', async () => {
    const { wb } = await fileEdit([{ op: 'set_range', sheet: '목록', range: 'D2', values: [['이름', '점수', '비고'], ['가', 90], ['나', '=E3+1', '메모']] }]);
    const ws = wb.getWorksheet('목록');
    assert.deepEqual([ws.getCell('D2').value, ws.getCell('E3').value, ws.getCell('F3').value, formulaOf(ws.getCell('E4'))], ['이름', 90, null, '=E3+1']);
  });

  test('clear: contents(서식 유지) · formats(값 유지) · all', async () => {
    const { wb } = await fileEdit([
      { op: 'format', sheet: '매출', range: 'A1:D2', bold: true, fill: '#FFFF00' },
      { op: 'clear', sheet: '매출', range: 'A1:B1' },
      { op: 'clear', sheet: '매출', range: 'C1', what: 'formats' },
      { op: 'clear', sheet: '매출', range: 'A2', what: 'all' },
    ]);
    const ws = wb.getWorksheet('매출');
    assert.equal(ws.getCell('A1').value, null);
    assert.equal(ws.getCell('A1').font?.bold, true);
    assert.equal(ws.getCell('C1').value, '2월');
    assert.ok(!ws.getCell('C1').font?.bold);
    assert.equal(ws.getCell('A2').value, null);
    assert.ok(!ws.getCell('A2').fill || ws.getCell('A2').fill.pattern !== 'solid');
  });

  test("clear 'all'은 메모도 지우고, 메모 없는 칸을 지워도 빈 메모를 만들지 않음", async () => {
    const src = await fresh('note');
    const wb = await load(src);
    wb.getWorksheet('매출').getCell('A2').note = '지울 메모';
    wb.getWorksheet('매출').getCell('A3').note = '남길 메모';
    await wb.xlsx.writeFile(src);
    const { wb: got, out } = await fileEdit([{ op: 'clear', sheet: '매출', range: 'A2', what: 'all' }, { op: 'clear', sheet: '목록', range: 'A1', what: 'all' }], { src });
    assert.equal(got.getWorksheet('매출').getCell('A2').note, undefined);
    assert.ok(got.getWorksheet('매출').getCell('A3').note);
    assert.ok(!got.getWorksheet('목록').getCell('A1').note);
    await excelAccepts(out);
  });

  test('Excel 2010 뒤의 함수는 파일에 _xlfn. 접두사로 (없으면 #NAME?), 동적 배열 함수는 거절', async () => {
    const { wb, out } = await fileEdit([{ op: 'set_cells', sheet: '목록', cells: { D1: '=STDEV.S(1,2,3)', D2: '=IFERROR(IFNA(1,2),0)', D3: '=CONCAT("a",IFS(TRUE,"b"))', D4: '="IFNA("&SUM(1)' } }]);
    const ls = wb.getWorksheet('목록');
    assert.deepEqual(['D1', 'D2', 'D3', 'D4'].map((a) => formulaOf(ls.getCell(a))),
      ['=_xlfn.STDEV.S(1,2,3)', '=IFERROR(_xlfn.IFNA(1,2),0)', '=_xlfn.CONCAT("a",_xlfn.IFS(TRUE,"b"))', '="IFNA("&SUM(1)']);
    const info = await excelInspect(out, { engine: 'file' });
    assert.ok(info.formulas.some((f) => f.formula === '=STDEV.S(1,2,3)'), '살펴보기에는 접두사 없이');
    await assert.rejects(fileEdit([{ op: 'set_cells', cells: { F1: '=SORT(A2:A6)' } }]), /SORT 함수\(동적 배열·LAMBDA\)는 파일 엔진으로 쓸 수 없습니다/);
    if (COM_TESTS) {
      const { readDocument } = await import('../src/core/index.mjs');
      const rows = (await readDocument(out, { engine: 'app' })).sheets[1].rows;
      assert.deepEqual([rows[0][3], rows[1][3]], ['1', '1']);
    }
  });

  test('병합된 칸의 일부에는 쓰지 못함', async () => {
    await assert.rejects(fileEdit([{ op: 'merge', sheet: '매출', range: 'A1:B1' }, { op: 'set_cells', sheet: '매출', cells: { B1: 'x' } }]),
      /2번째 작업\(set_cells\) 실패: B1는 병합된 칸의 일부라 쓸 수 없습니다 — 병합 영역의 왼쪽 위 칸 A1에 쓰세요/);
  });
});

describe('파일 엔진: 행·열 넣기·지우기 (수식·병합·조건부 서식·입력 제한·표·이름 정의도 옮김)', () => {
  const setup = [
    { op: 'merge', sheet: '목록', range: 'D5:E6' },
    { op: 'conditional_format', sheet: '매출', range: 'B2:C6', rule: { type: 'greater_than', value: 50 } },
    { op: 'data_validation', sheet: '매출', range: 'B2:C6', type: 'whole', min: 0 },
    { op: 'named_range', name: '일월', range: '매출!B2:B6' },
    { op: 'add_table', sheet: '목록', range: 'A10:B12', name: '목록표' },
  ];

  test('insert_rows: 같은 시트·다른 시트 수식, 이름 정의, 조건부 서식, 입력 제한이 늘어남', async () => {
    const { wb, out } = await fileEdit([...setup, { op: 'insert_rows', sheet: '매출', at: 4, count: 2 }]);
    const ws = wb.getWorksheet('매출');
    assert.equal(ws.getCell('A6').value, '대구');
    assert.equal(formulaOf(ws.getCell('D6')), '=SUM(B6:C6)');
    assert.equal(formulaOf(ws.getCell('B9')), '=SUM(B2:B8)');
    assert.equal(formulaOf(wb.getWorksheet('목록').getCell('B1')), '=매출!D9');
    assert.equal(formulaOf(wb.getWorksheet('목록').getCell('B3')), '=매출!A3');
    assert.deepEqual(ws.conditionalFormattings.map((c) => c.ref), ['B2:C8']);
    assert.ok(ws.dataValidations.model.B4 && ws.dataValidations.model.C5, '넣은 행에도 입력 제한');
    assert.match(await zipText(out, 'xl/workbook.xml'), /<definedName name="일월">매출!\$B\$2:\$B\$8<\/definedName>/);
    await excelAccepts(out);
  });

  test('delete_rows: 지운 행을 가리키던 수식은 #REF!, 표·병합도 옮김', async () => {
    const { wb, out } = await fileEdit([...setup, { op: 'delete_rows', sheet: '목록', at: 2, count: 2 }, { op: 'delete_rows', sheet: '매출', at: 3 }]);
    const ls = wb.getWorksheet('목록');
    assert.equal(formulaOf(ls.getCell('B1')), '=매출!D6');
    assert.equal(ls.getCell('B2').value, null, '목록 2~3행이 지워짐');
    assert.deepEqual(Object.values(ls._merges).map((m) => m.range), ['D3:E4']);
    assert.equal(ls.getTable('목록표').table.tableRef, 'A8:B10');
    assert.equal(formulaOf(wb.getWorksheet('매출').getCell('B6')), '=SUM(B2:B5)');
    await excelAccepts(out);
  });

  test('insert_columns · delete_columns: 수식·표 열(머리글)도 맞춤', async () => {
    const { wb, out } = await fileEdit([...setup,
      { op: 'insert_columns', sheet: '매출', at: 'B' },
      { op: 'insert_columns', sheet: '목록', at: 'B', count: 1 },
      { op: 'delete_columns', sheet: '매출', at: 'D' }]);
    const ws = wb.getWorksheet('매출');
    assert.deepEqual([ws.getCell('A1').value, ws.getCell('B1').value, ws.getCell('C1').value, ws.getCell('D1').value], ['지점', null, '1월', '합계']);
    assert.equal(formulaOf(ws.getCell('D2')), '=SUM(C2:C2)', '2월 열을 지우면 범위가 줄어듦 (Excel과 같음)');
    assert.equal(formulaOf(ws.getCell('C7')), '=SUM(C2:C6)');
    const t = wb.getWorksheet('목록').getTable('목록표').table;
    assert.equal(t.tableRef, 'A10:C12');
    assert.equal(t.columns.length, 3);
    assert.equal(wb.getWorksheet('목록').getCell('B10').value, t.columns[1].name);
    assert.match(await zipText(out, 'xl/workbook.xml'), /일월">매출!\$C\$2:\$C\$6</);
    await excelAccepts(out);
  });
});

describe('파일 엔진: 서식·크기·병합·틀 고정', () => {
  test('format: 글꼴·크기·색·채우기·표시 형식·정렬·줄 바꿈·테두리', async () => {
    const { wb, out } = await fileEdit([
      { op: 'format', sheet: '매출', range: 'A1:D1', bold: true, italic: true, underline: true, font: 'Arial', size: 14, color: '#FFFFFF', fill: '#1F4E78', align: 'center', valign: 'center', wrap: true, border: 'thin' },
      { op: 'format', sheet: '매출', range: 'B2:D7', number_format: '#,##0', border: 'outline' },
      { op: 'format', sheet: '매출', range: 'A7', fill: 'none' },
    ]);
    const ws = wb.getWorksheet('매출');
    const a1 = ws.getCell('A1');
    assert.equal(a1.font.bold, true); assert.equal(a1.font.italic, true); assert.ok(a1.font.underline);
    assert.equal(a1.font.name, 'Arial'); assert.equal(a1.font.size, 14); assert.equal(a1.font.color.argb, 'FFFFFFFF');
    assert.equal(a1.fill.fgColor.argb, 'FF1F4E78');
    assert.deepEqual([a1.alignment.horizontal, a1.alignment.vertical, a1.alignment.wrapText], ['center', 'middle', true]);
    assert.equal(a1.border.left.style, 'thin');
    assert.equal(ws.getCell('B2').numFmt, '#,##0');
    assert.equal(ws.getCell('B2').border.top.style, 'thin');
    assert.ok(!ws.getCell('C3').border?.top, 'outline은 바깥쪽만');
    assert.equal(ws.getCell('D7').border.right.style, 'thin');
    await excelAccepts(out);
  });

  test('column_width · row_height (숫자, auto 근사)', async () => {
    const { wb } = await fileEdit([
      { op: 'set_cells', sheet: '목록', cells: { A5: '아주 긴 한글 제목이 들어간 칸입니다' } },
      { op: 'column_width', sheet: '목록', columns: 'A', width: 'auto' },
      { op: 'column_width', sheet: '목록', columns: 'B:C', width: 20 },
      { op: 'row_height', sheet: '목록', rows: '2:3', height: 30 },
      { op: 'format', sheet: '목록', range: 'A6', wrap: true },
      { op: 'set_cells', sheet: '목록', cells: { A6: '첫 줄\n둘째 줄\n셋째 줄' } },
      { op: 'row_height', sheet: '목록', rows: 6, height: 'auto' },
    ]);
    const ls = wb.getWorksheet('목록');
    assert.ok(ls.getColumn(1).width > 30, `너비 ${ls.getColumn(1).width}`);
    assert.equal(ls.getColumn(3).width, 20);
    assert.equal(ls.getRow(3).height, 30);
    assert.ok(ls.getRow(6).height >= 40, `높이 ${ls.getRow(6).height}`);
  });

  test('merge(왼쪽 위 값만 남김) · unmerge · 걸친 병합은 오류', async () => {
    const { wb, r } = await fileEdit([{ op: 'merge', sheet: '매출', range: 'A1:B2' }, { op: 'merge', sheet: '매출', range: 'C1:D1' }, { op: 'unmerge', sheet: '매출', range: 'D1' }]);
    const ws = wb.getWorksheet('매출');
    assert.deepEqual(Object.values(ws._merges).map((m) => m.range), ['A1:B2']);
    assert.equal(ws.getCell('A1').value, '지점');
    assert.match(r.results[0], /값이 있던 칸 3개는 지워짐/);
    await assert.rejects(fileEdit([{ op: 'merge', sheet: '매출', range: 'A1:B2' }, { op: 'merge', sheet: '매출', range: 'B2:C3' }]), /이미 병합된 영역 A1:B2가 범위 밖으로/);
  });

  test('freeze · 해제', async () => {
    const { wb } = await fileEdit([{ op: 'freeze', sheet: '매출', cell: 'B2' }, { op: 'freeze', sheet: '목록', cell: 'A3' }, { op: 'freeze', sheet: '목록', cell: null }]);
    const v = wb.getWorksheet('매출').views[0];
    assert.deepEqual([v.state, v.xSplit, v.ySplit, v.topLeftCell], ['frozen', 1, 1, 'B2']);
    assert.notEqual(wb.getWorksheet('목록').views[0]?.state, 'frozen');
    const info = await excelInspect((await fileEdit([{ op: 'freeze', cell: 'C4' }])).out, { engine: 'file' });
    assert.equal(info.sheets[0].freeze, 'C4');
  });
});

describe('파일 엔진: 필터·표·조건부 서식·입력 제한·이름 정의', () => {
  test('autofilter 걸기·풀기, 빈 범위·표와 겹침은 오류', async () => {
    const { wb } = await fileEdit([{ op: 'autofilter', sheet: '매출', range: 'A1:D6' }]);
    assert.equal(wb.getWorksheet('매출').autoFilter, 'A1:D6');
    const off = await fileEdit([{ op: 'autofilter', sheet: '매출', range: 'A1:D6' }, { op: 'autofilter', sheet: '매출', range: null }]);
    assert.ok(!off.wb.getWorksheet('매출').autoFilter);
    await assert.rejects(fileEdit([{ op: 'autofilter', sheet: '매출', range: 'H1:J5' }]), /비어 있는 범위/);
    await assert.rejects(fileEdit([{ op: 'add_table', sheet: '매출', range: 'A1:D6' }, { op: 'autofilter', sheet: '매출', range: 'A1:B3' }]), /표 '표1'와 겹칩니다/);
  });

  test('add_table: 이름·스타일·머리글(빈 칸은 열N), Excel이 받아들임', async () => {
    const { wb, out, r } = await fileEdit([
      { op: 'set_cells', sheet: '매출', cells: { E1: null } },
      { op: 'add_table', sheet: '매출', range: 'A1:E6', name: '매출표', style: 'TableStyleLight9' },
      { op: 'add_table', sheet: '목록', range: 'D1:E3' },
      { op: 'set_cells', sheet: '매출', cells: { B1: '일월' } },
    ]);
    const t = wb.getWorksheet('매출').getTable('매출표').table;
    assert.equal(t.tableRef, 'A1:E6');
    assert.equal(t.style.theme, 'TableStyleLight9');
    assert.deepEqual(t.columns.map((c) => c.name), ['지점', '일월', '2월', '합계', '열1']);
    assert.equal(wb.getWorksheet('매출').getCell('E1').value, '열1');
    assert.equal(wb.getWorksheet('목록').getTable('표1').table.tableRef, 'D1:E3');
    assert.match(r.results[2], /표 만듦: 표1/);
    const tableXml = await zipText(out, 'xl/tables/table1.xml');
    assert.match(tableXml, /headerRowCount="1"/);
    const info = await excelAccepts(out);
    if (info) assert.deepEqual(info.sheets[0].tables.map((x) => [x.name, x.range]), [['매출표', 'A1:E6']]);
    await assert.rejects(fileEdit([{ op: 'add_table', sheet: '매출', range: 'A1:D6', style: '없는스타일' }]), /표 스타일 이름이 올바르지 않습니다/);
    await assert.rejects(fileEdit([{ op: 'merge', sheet: '매출', range: 'A1:B1' }, { op: 'add_table', sheet: '매출', range: 'A1:D6' }]), /병합된 칸이 있는 범위/);
  });

  test('불러온 표의 머리글 행 유지 (exceljs가 headerRowCount 생략을 머리글 없음으로 읽는 문제)', async () => {
    const { out } = await fileEdit([{ op: 'add_table', sheet: '매출', range: 'A1:D6', name: 'T' }]);
    const zip = await JSZip.loadAsync(readFileSync(out));
    const name = Object.keys(zip.files).find((f) => /^xl\/tables\/.+\.xml$/.test(f));
    zip.file(name, (await zip.file(name).async('string')).replace(/\sheaderRowCount="1"/, '')); // Excel은 기본값을 생략한다
    const src = tmp.path('excel-made-table.xlsx');
    writeFileSync(src, await zip.generateAsync({ type: 'nodebuffer' }));
    const again = await fileEdit([{ op: 'set_cells', sheet: '목록', cells: { A9: 1 } }], { src });
    assert.doesNotMatch(await zipText(again.out, name), /headerRowCount="0"/);
  });

  test('conditional_format: 여덟 가지 규칙', async () => {
    const { wb, out } = await fileEdit([
      { op: 'conditional_format', sheet: '매출', range: 'B2:B6', rule: { type: 'greater_than', value: 40 } },
      { op: 'conditional_format', sheet: '매출', range: 'C2:C6', rule: { type: 'less_than', value: 50, fill: '#C6EFCE', font_color: '#006100' } },
      { op: 'conditional_format', sheet: '매출', range: 'B2:C6', rule: { type: 'between', value: 20, value2: 60 } },
      { op: 'conditional_format', sheet: '매출', range: 'B2:B6', rule: { type: 'equal', value: 30 } },
      { op: 'conditional_format', sheet: '매출', range: 'A2:A6', rule: { type: 'text_contains', value: '산"' } },
      { op: 'conditional_format', sheet: '매출', range: 'A2:A6', rule: { type: 'duplicate' } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'color_scale' } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'data_bar', fill: '#63BE7B' } },
    ]);
    const rules = wb.getWorksheet('매출').conditionalFormattings.flatMap((c) => c.rules.map((x) => [c.ref, x.type, x.operator ?? null]));
    assert.deepEqual(rules.map((x) => x[1]), ['cellIs', 'cellIs', 'cellIs', 'cellIs', 'expression', 'expression', 'colorScale', 'dataBar']);
    assert.deepEqual(rules.slice(0, 4).map((x) => x[2]), ['greaterThan', 'lessThan', 'between', 'equal']);
    const cfs = wb.getWorksheet('매출').conditionalFormattings;
    assert.equal(cfs[4].rules[0].formulae[0], 'NOT(ISERROR(SEARCH("산""",A2)))');
    assert.equal(cfs[5].rules[0].formulae[0], 'COUNTIF($A$2:$A$6,A2)>1');
    assert.equal((await excelInspect(out, { engine: 'file' })).sheets[0].conditionalFormats, 8);
    const info = await excelAccepts(out);
    if (info) assert.equal(info.sheets[0].conditionalFormats, 8);
  });

  test('data_validation: 목록·정수·소수, 겹치는 기존 제한은 바뀜', async () => {
    const { wb, out } = await fileEdit([
      { op: 'data_validation', sheet: '목록', range: 'C1:C10', list: ['예', '아니오'] },
      { op: 'data_validation', sheet: '목록', range: 'D1:D10', type: 'whole', min: 1, max: 10 },
      { op: 'data_validation', sheet: '목록', range: 'E1:E3', type: 'decimal', max: 0.5 },
      { op: 'data_validation', sheet: '목록', range: 'C5:C6', list: '=$A$1:$A$3' },
    ]);
    const dv = wb.getWorksheet('목록').dataValidations.model;
    assert.deepEqual([dv.C1.type, dv.C1.formulae[0]], ['list', '"예,아니오"']);
    assert.deepEqual([dv.C5.type, dv.C5.formulae[0]], ['list', '$A$1:$A$3']);
    assert.deepEqual([dv.D3.type, dv.D3.operator, dv.D3.formulae], ['whole', 'between', [1, 10]]);
    assert.deepEqual([dv.E2.type, dv.E2.operator, dv.E2.formulae], ['decimal', 'lessThanOrEqual', [0.5]]);
    assert.equal(dv.C7.formulae[0], '"예,아니오"');
    await excelAccepts(out);
  });

  test('named_range: 추가·바꾸기, 시트 이름에 띄어쓰기', async () => {
    const { out } = await fileEdit([
      { op: 'rename_sheet', sheet: '목록', name: '내 목록' },
      { op: 'named_range', name: '범위1', range: "'내 목록'!A1:B3" },
      { op: 'named_range', name: '범위2', range: 'A1' },
      { op: 'named_range', name: '범위2', range: '매출!C2:C6' },
    ]);
    const wbXml = await zipText(out, 'xl/workbook.xml');
    assert.match(wbXml, /<definedName name="범위1">'내 목록'!\$A\$1:\$B\$3<\/definedName>/);
    assert.match(wbXml, /<definedName name="범위2">매출!\$C\$2:\$C\$6<\/definedName>/);
    const info = await excelInspect(out, { engine: 'file' });
    assert.deepEqual(info.names.map((n) => n.name), ['범위1', '범위2']);
    const com = await excelAccepts(out);
    if (com) assert.deepEqual(com.names.map((n) => n.name).sort(), ['범위1', '범위2']);
  });
});

describe('파일 엔진: 못 하는 것·지키는 것', () => {
  test('×인 작업은 Excel이 필요하다고 알려 줌 (아무것도 저장하지 않음)', async () => {
    const src = await fresh('need');
    const before = readFileSync(src);
    for (const op of [{ op: 'sort', range: 'A1:D6', by: 'B' }, { op: 'add_chart', range: 'A1:C6', type: 'column' }, { op: 'copy_sheet', sheet: 1, name: 'x' }, { op: 'recalculate' }]) {
      await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { Z1: 1 } }, op], { engine: 'file' }),
        (e) => new RegExp(`할 수 없는 작업입니다: ${op.op}`).test(e.message) && /Excel이 설치돼 있어야|Excel이 있으니 engine:'auto' 또는 'app'/.test(e.message));
    }
    assert.ok(before.equals(readFileSync(src)));
  });

  test('한 번에 저장: 중간 작업이 실패하면 원본·출력 모두 그대로', async () => {
    const src = await fresh('atomic');
    const before = readFileSync(src);
    const out = tmp.path('atomic-out.xlsx');
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '바뀜' } }, { op: 'add_sheet', name: '새' }, { op: 'set_cells', sheet: '없음', cells: { A1: 1 } }], { output: out, engine: 'file' }),
      /3번째 작업\(set_cells\) 실패: 시트 없음: 없음 .*저장하지 않았습니다/);
    assert.equal(existsSync(out), false);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '바뀜' } }, { op: 'merge', range: 'A1:B2' }, { op: 'set_cells', cells: { B2: 1 } }], { engine: 'file' }), /3번째 작업/);
    assert.ok(before.equals(readFileSync(src)), '원본이 바뀜');
  });

  test('차트 등 보존할 수 없는 내용이 있으면 고치지 않음', async () => {
    const src = tmp.path('chart.xlsx');
    writeFileSync(src, await xlsxWithChart());
    const before = readFileSync(src);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', sheet: '매출', cells: { C1: 1 } }], { engine: 'file' }), /사라지는 내용이 있어.*차트/);
    assert.ok(before.equals(readFileSync(src)));
    const info = await excelInspect(src, { engine: 'file' });
    assert.equal(info.sheets[0].charts.length, 1);
    assert.match(inspectText(info), /차트 1개/);
    assert.match(inspectText(info), /Excel로만 고칩니다: 차트/);
  });

  test('형식: xlsm·xls·ods는 파일 엔진으로 못 고침, 다른 형식으로 저장 못 함', async () => {
    const src = await fresh('fmt');
    const xlsm = tmp.path('macro.xlsm'); copyFileSync(src, xlsm);
    await assert.rejects(excelEdit(xlsm, [{ op: 'set_cells', cells: { A1: 1 } }], { engine: 'file' }), /xlsm\(매크로 문서\)/);
    const xls = tmp.path('old.xls'); writeFileSync(xls, 'x');
    await assert.rejects(excelEdit(xls, [{ op: 'set_cells', cells: { A1: 1 } }], { engine: 'file' }), /\.xls 파일을 고칠 수 없습니다/);
    await assert.rejects(excelInspect(xls, { engine: 'file' }), /Excel 없이 열 수 없습니다/);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: 1 } }], { engine: 'file', output: tmp.path('x.xlsm') }), /같은 형식\(\.xlsx\)으로만/);
    await assert.rejects(excelEdit(tmp.path('없는.xlsx'), [{ op: 'recalculate' }]), /파일이 없습니다/);
    const docx = tmp.path('a.docx'); writeFileSync(docx, 'x');
    await assert.rejects(excelEdit(docx, [{ op: 'recalculate' }]), /엑셀 문서/);
  });

  test('손대지 않은 내용 유지: 숨긴 시트·그림·메모·문서 안 링크·자동 필터 조건·기본 글꼴', async () => {
    // Excel이 만든 것처럼: 숨긴 시트, 그림, 메모, location 링크, 필터 조건, 기본 글꼴 맑은 고딕
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('데이터');
    ws.addRows([['이름', '점수'], ['가', 1], ['나', 2], ['다', 3]]);
    ws.getCell('A2').note = '메모입니다';
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
    ws.addImage(wb.addImage({ buffer: png, extension: 'png' }), 'D2:E5');
    ws.autoFilter = 'A1:B4';
    wb.addWorksheet('숨김', { state: 'hidden' }).getCell('A1').value = '비밀';
    const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
    let sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
    sheet = sheet.replace(/<autoFilter ref="A1:B4"\/>/, '<autoFilter ref="A1:B4"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>');
    sheet = sheet.replace(/<pageMargins/, '<hyperlinks><hyperlink ref="A3" location="\'숨김\'!A1" display="나"/></hyperlinks><pageMargins');
    zip.file('xl/worksheets/sheet1.xml', sheet);
    const styles = await zip.file('xl/styles.xml').async('string');
    zip.file('xl/styles.xml', styles.replace(/(<fonts[^>]*>\s*<font>)[\s\S]*?(<\/font>)/, '$1<sz val="11"/><color theme="1"/><name val="맑은 고딕"/><family val="2"/><scheme val="minor"/>$2'));
    const src = tmp.path('keep.xlsx');
    writeFileSync(src, await zip.generateAsync({ type: 'nodebuffer' }));

    const { out, wb: got } = await fileEdit([{ op: 'set_cells', sheet: '데이터', cells: { C1: '추가' } }, { op: 'insert_rows', sheet: '데이터', at: 1 }], { src });
    assert.equal(got.getWorksheet('숨김').state, 'hidden');
    assert.equal(got.getWorksheet('데이터').getCell('A3').note?.texts?.[0]?.text ?? got.getWorksheet('데이터').getCell('A3').note, '메모입니다');
    assert.equal(got.getWorksheet('데이터').getImages().length, 1);
    assert.equal(got.getWorksheet('데이터').getImages()[0].range.tl.nativeRow, 2, '그림도 한 행 내려감');
    const outSheet = await zipText(out, 'xl/worksheets/sheet1.xml');
    assert.match(outSheet, /<hyperlink ref="A4" location="'숨김'!A1"\/>/);
    assert.doesNotMatch(outSheet, /<hyperlink[^>]*r:id/);
    assert.match(await zipText(out, 'xl/styles.xml'), /<fonts[^>]*><font><sz val="11"\/><color theme="1"\/><name val="맑은 고딕"\/>/);
    // 행을 넣지 않은 시트의 필터 조건은 그대로 (행을 넣으면 Excel처럼 범위만 옮긴다)
    const { out: out2 } = await fileEdit([{ op: 'set_cells', sheet: '데이터', cells: { C1: '추가' } }], { src });
    assert.match(await zipText(out2, 'xl/worksheets/sheet1.xml'), /<autoFilter ref="A1:B4"><filterColumn colId="1">/);
    await excelAccepts(out);
    await excelAccepts(out2);
  });
});

describe('파일 엔진: Excel이 만든 표·낡은 수식 결과·표 머리글 수식', () => {
  // Excel이 만든 것 같은 표: 계산 열(<calculatedColumnFormula>)이 있고 autoFilter에 filterColumn이 없음, 필터 단추를 끈 표(autoFilter 없음)
  async function excelMadeTables() {
    const { out } = await fileEdit([
      { op: 'set_cells', sheet: '매출', cells: { E1: '비고', E2: 'a' } },
      { op: 'add_table', sheet: '매출', range: 'A1:E6', name: '매출표' },
      { op: 'add_table', sheet: '목록', range: 'D1:E3', name: '단추없음' },
    ], { name: 'xlmade' });
    const zip = await JSZip.loadAsync(readFileSync(out));
    for (const name of Object.keys(zip.files).filter((f) => /^xl\/tables\/.+\.xml$/.test(f))) {
      let xml = await zip.file(name).async('string');
      if (xml.includes('매출표')) {
        xml = xml.replace(/<autoFilter ref="([^"]*)">[\s\S]*?<\/autoFilter>/, '<autoFilter ref="$1"/>')
          .replace(/<tableColumn id="4" name="합계"\/>/, '<tableColumn id="4" name="합계"><calculatedColumnFormula>SUM(B2:C2)</calculatedColumnFormula></tableColumn>');
      } else xml = xml.replace(/<autoFilter\b[\s\S]*?<\/autoFilter>/, '');
      zip.file(name, xml);
    }
    const src = tmp.path(`xlmade-${++seq}.xlsx`);
    writeFileSync(src, await zip.generateAsync({ type: 'nodebuffer' }));
    return src;
  }
  const tableXml = async (path, name) => {
    const zip = await JSZip.loadAsync(readFileSync(path));
    for (const f of Object.keys(zip.files).filter((x) => /^xl\/tables\/.+\.xml$/.test(x))) {
      const xml = await zip.file(f).async('string');
      if (xml.includes(`name="${name}"`)) return xml;
    }
    return null;
  };

  test('계산 열이 있는 표: 다른 시트를 고쳐도 열이 사라지지 않고, 필터 단추·계산 열 수식이 남음', async () => {
    const src = await excelMadeTables();
    const { out } = await fileEdit([{ op: 'set_cells', sheet: '목록', cells: { A9: 1 } }], { src });
    const xml = await tableXml(out, '매출표');
    assert.match(xml, /<tableColumns count="5">/);
    assert.match(xml, /<tableColumn id="4" name="합계"><calculatedColumnFormula>SUM\(B2:C2\)<\/calculatedColumnFormula><\/tableColumn><tableColumn id="5" name="비고"\/>/);
    assert.doesNotMatch(xml, /hiddenButton="1"/, '필터 단추가 숨겨짐');
    assert.doesNotMatch(await tableXml(out, '단추없음'), /<autoFilter/, '필터 단추를 끈 표에 자동 필터가 생김');
    // 고친 파일을 파일 엔진이 다시 열고 고칠 수 있다
    const again = await fileEdit([{ op: 'set_cells', sheet: '목록', cells: { A10: 2 } }], { src: out });
    assert.match(await tableXml(again.out, '매출표'), /<tableColumns count="5">[\s\S]*SUM\(B2:C2\)/);
    await excelAccepts(out);
  });

  test('계산 열 수식은 행·열을 넣고 지우면 표의 첫 데이터 행 기준으로 맞춤', async () => {
    const src = await excelMadeTables();
    const calc = async (ops) => /<calculatedColumnFormula>([^<]*)</.exec(await tableXml((await fileEdit(ops, { src })).out, '매출표'))?.[1] ?? null;
    assert.equal(await calc([{ op: 'insert_rows', sheet: '매출', at: 1, count: 2 }]), 'SUM(B4:C4)', '표 위에 넣으면 함께 내려감');
    assert.equal(await calc([{ op: 'insert_rows', sheet: '매출', at: 2 }]), 'SUM(B2:C2)', '첫 데이터 행 자리에 넣으면 기준은 그대로');
    assert.equal(await calc([{ op: 'delete_rows', sheet: '매출', at: 2 }]), 'SUM(B2:C2)', '첫 데이터 행을 지우면 다음 행이 기준');
    assert.equal(await calc([{ op: 'insert_columns', sheet: '매출', at: 'C' }]), 'SUM(B2:D2)');
    assert.equal(await calc([{ op: 'delete_columns', sheet: '매출', at: 'B', count: 2 }]), null, '참조하던 열을 지우면 계산 열 표시만 뺌');
    const { out, wb } = await fileEdit([{ op: 'insert_rows', sheet: '매출', at: 2 }, { op: 'insert_columns', sheet: '매출', at: 'C' }, { op: 'insert_rows', sheet: '매출', at: 5 }], { src });
    assert.match(await tableXml(out, '매출표'), /ref="A1:F8"[\s\S]*<tableColumns count="6">/);
    // 표 안에 넣은 행의 계산 열에는 Excel처럼 수식이 채워짐
    assert.deepEqual(['E2', 'E5'].map((a) => formulaOf(wb.getWorksheet('매출').getCell(a))), ['=SUM(B2:D2)', '=SUM(B5:D5)']);
    assert.equal(wb.getWorksheet('매출').getCell('B5').value, null);
    await excelAccepts(out);
  });

  test('값이 바뀌면 수식 칸의 옛 결과를 비움 (읽는 쪽이 틀린 숫자를 보지 않게)', async () => {
    const src = await fresh('stale');
    const wb = await load(src);
    const ws = wb.getWorksheet('매출');
    ws.getCell('B7').value = { formula: 'SUM(B2:B6)', result: 250 };
    ws.getCell('C7').value = { formula: 'SUM(C2:C6)', result: 300 };
    await wb.xlsx.writeFile(src);
    const { wb: got, r } = await fileEdit([{ op: 'delete_rows', sheet: '매출', at: 2 }], { src });
    const b6 = got.getWorksheet('매출').getCell('B6');
    assert.equal(formulaOf(b6), '=SUM(B2:B5)');
    assert.equal(b6.value.result, undefined, '옛 결과(250)가 남음');
    assert.ok(r.notes.some((n) => /옛 결과를 비웠습니다/.test(n)), r.notes.join(' / '));
    const { readDocument } = await import('../src/core/index.mjs');
    const rows = (await readDocument((await fileEdit([{ op: 'set_cells', sheet: '매출', cells: { B2: 1000 } }], { src })).out, { engine: 'file' })).sheets[0].rows;
    assert.notEqual(rows[6][1], '250');
    // 서식만 바꾸면 결과를 그대로 둠
    const fmt = await fileEdit([{ op: 'format', sheet: '매출', range: 'A1', bold: true }], { src });
    assert.equal(fmt.wb.getWorksheet('매출').getCell('B7').value.result, 250);
  });

  test('add_table: 머리글 칸의 수식은 저장된 결과를 글자로, 결과가 없으면 거절', async () => {
    const src = await fresh('hdrf');
    const wb = await load(src);
    wb.getWorksheet('목록').getCell('D1').value = { formula: '40+1', result: 41 };
    wb.getWorksheet('목록').getCell('E1').value = '이름';
    await wb.xlsx.writeFile(src);
    const { wb: got } = await fileEdit([{ op: 'add_table', sheet: '목록', range: 'D1:E3' }], { src });
    assert.equal(got.getWorksheet('목록').getCell('D1').value, '41');
    await assert.rejects(fileEdit([{ op: 'set_cells', sheet: '목록', cells: { D1: '=40+1' } }, { op: 'add_table', sheet: '목록', range: 'D1:E3' }], { src }),
      /2번째 작업\(add_table\) 실패: 표 머리글 칸\(D1\)에 수식이 있습니다 — 값으로 바꾼 뒤 만드세요/);
  });

  test('파일 엔진이 쓴 수식이 잘못되면 저장하지 않음 (Excel이 열지 못하는 파일 방지)', async () => {
    const src = await fresh('badf');
    const before = readFileSync(src);
    for (const ops of [[{ op: 'set_cells', cells: { F1: '=SUMM((1' } }], [{ op: 'set_cells', cells: { F1: '=SUM(A1;A2)' } }],
      [{ op: 'data_validation', range: 'F1:F3', list: ['5" 모니터'] }], [{ op: 'data_validation', range: 'F1', list: '=$Z$1:' }]]) {
      await assert.rejects(excelEdit(src, ops, { engine: 'file' }), /오류: .*저장하지 않았습니다/);
    }
    assert.ok(before.equals(readFileSync(src)));
  });
});

describe('살펴보기 글: 큰 통합 문서는 줄임', () => {
  test('시트 100개·시트마다 표/차트 20개까지 자세히, 나머지는 개수와 이름만', () => {
    const sheets = Array.from({ length: 130 }, (_, i) => ({
      index: i + 1, name: `S${i + 1}`, visible: 'visible', usedRange: 'A1:B2', rows: 2, cols: 2,
      tables: Array.from({ length: i === 0 ? 25 : 1 }, (_, k) => ({ name: `T${i}_${k}`, range: 'A1:B2' })), charts: [],
      ...(i >= 100 ? { brief: true, tables: [] } : {}),
    }));
    const text = inspectText({ engine: 'excel', format: 'xlsx', sheets, names: [], formulas: [], formulaCount: 0 });
    assert.match(text, /시트 130개/);
    assert.match(text, /1\. 'S1' — .*표 25개/);
    assert.match(text, /… 외 표 5개/);
    assert.ok(!text.includes('T0_20'));
    assert.match(text, /… 외 시트 30개 \(처음 100개 시트만 자세히, 수식 수도 그 시트들만 셈\): 101\. 'S101'/);
    assert.ok(text.split('\n').length < 250, `줄 수 ${text.split('\n').length}`);
  });
});

describe('Excel을 확인하지 못할 때 알림', { skip: process.platform !== 'win32' && 'Windows만' }, () => {
  test('프로그램 확인이 실패하면 그 사실과 대안을 알려 줌', async () => {
    const { execFileSync } = await import('node:child_process');
    const xls = tmp.path('probe.xls'); writeFileSync(xls, 'x');
    const code = `import(${JSON.stringify(pathToFileURL(join(ROOT, 'src', 'core', 'excel.mjs')).href)}).then((m) => m.excelEdit(${JSON.stringify(xls)}, [{ op: 'recalculate' }], { engine: 'app' }))`
      + '.then(() => console.log("성공"), (e) => console.log(e.message))';
    const out = execFileSync(process.execPath, ['-e', code], { env: { ...process.env, OFFICEKIT_POWERSHELL: join(tmp.dir, 'none.exe') }, encoding: 'utf8' });
    assert.match(out, /engine:'app'으로 고칠 수 없습니다 — Excel을 쓸 수 있는지 확인하지 못했습니다 \(.+\) — 대신 office_set_cells/);
  });
});

describe('파일 엔진: csv (값만)', () => {
  test('값 쓰기·행열 넣고 지우기, 인코딩(BOM)·구분자 유지', async () => {
    const src = tmp.path('a.csv');
    writeFileSync(src, '\uFEFF이름,점수\r\n가,1\r\n"나,다",2\r\n');
    const r = await excelEdit(src, [
      { op: 'set_cells', cells: { C1: '비고', C2: 'x"y' } },
      { op: 'set_range', range: 'A5', values: [['라', 4]] },
      { op: 'insert_rows', at: 2 },
      { op: 'delete_columns', at: 'B' },
    ], { engine: 'file' });
    assert.equal(r.engine, 'file');
    const text = readFileSync(src, 'utf8');
    assert.ok(text.startsWith('\uFEFF'));
    assert.equal(text.slice(1), '이름,비고\r\n,\r\n가,"x""y"\r\n"나,다",\r\n,\r\n라,\r\n');
    const info = await excelInspect(src);
    assert.deepEqual([info.engine, info.sheets[0].rows, info.sheets[0].cols], ['file', 6, 2]);
  });

  test("engine 'auto'의 csv 값 작업은 Excel이 있어도 파일 엔진 (BOM 없는 UTF-8·'0123' 그대로)", async () => {
    const src = tmp.path('nobom.csv');
    writeFileSync(src, '이름,점수,코드\n김철수,90,0123\n이영희,85,1/2\n');
    const r = await excelEdit(src, [{ op: 'set_cells', cells: { B3: 100 } }]);
    assert.equal(r.engine, 'file');
    assert.equal(readFileSync(src, 'utf8'), '이름,점수,코드\n김철수,90,0123\n이영희,100,1/2\n');
  });

  test('서식·수식·시트 작업은 csv로 저장할 수 없음', async () => {
    const src = tmp.path('b.csv');
    writeFileSync(src, 'a,b\n1,2\n');
    await assert.rejects(excelEdit(src, [{ op: 'format', range: 'A1', bold: true }], { engine: 'file' }), /csv 파일은 값만 저장되므로 할 수 없는 작업입니다: format/);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '=1+1' } }], { engine: 'file' }), /csv 파일에는 수식을 저장할 수 없습니다/);
    await assert.rejects(excelEdit(src, [{ op: 'add_sheet', name: 'x' }], { engine: 'file' }), /csv/);
    assert.equal(readFileSync(src, 'utf8'), 'a,b\n1,2\n');
  });
});

describe('파일 엔진: 살펴보기', () => {
  test('시트·사용 범위·표·병합·필터·조건부 서식·이름·수식 예시', async () => {
    const { out } = await fileEdit([
      { op: 'merge', sheet: '목록', range: 'D1:E1' }, { op: 'autofilter', sheet: '매출', range: 'A1:D6' },
      { op: 'add_table', sheet: '목록', range: 'A10:B12', name: '목록T' }, { op: 'named_range', name: '합', range: '매출!D7' },
      { op: 'conditional_format', sheet: '매출', range: 'B2:B6', rule: { type: 'data_bar' } }, { op: 'freeze', sheet: '매출', cell: 'A2' },
    ]);
    const info = await excelInspect(out);
    assert.equal(info.engine, 'file');
    assert.equal(info.activeSheet, '매출');
    const [s1, s2] = info.sheets;
    assert.deepEqual([s1.index, s1.name, s1.visible, s1.usedRange, s1.rows, s1.cols, s1.freeze, s1.autoFilter, s1.conditionalFormats],
      [1, '매출', 'visible', 'A1:D7', 7, 4, 'A2', 'A1:D6', 1]);
    assert.deepEqual([s2.merges, s2.tables.map((t) => t.name)], [1, ['목록T']]);
    assert.deepEqual(info.names, [{ name: '합', refersTo: '=매출!$D$7', scope: null }]);
    assert.equal(info.formulaCount, 11);
    assert.deepEqual(info.formulas[0], { sheet: '매출', cell: 'D2', formula: '=SUM(B2:C2)' });
    const text = inspectText(info);
    for (const s of ['시트 2개', "활성 시트: '매출'", '틀 고정 A2', '자동 필터 A1:D6', '병합 1개', '표 목록T: A10:B12', '합 = 매출!$D$7', '수식 11개']) assert.ok(text.includes(s), `${s} 없음:\n${text}`);
  });
});

describe('MCP 도구 excel_inspect / excel_edit', () => {
  test('등록·한국어 설명·호출', async () => {
    const client = new Client({ name: 'excel-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [join(ROOT, 'src', 'mcp', 'server.mjs')], cwd: tmp.dir, stderr: 'pipe' }));
    try {
      const { tools } = await client.listTools();
      const edit = tools.find((t) => t.name === 'excel_edit');
      const inspect = tools.find((t) => t.name === 'excel_inspect');
      assert.ok(edit && inspect);
      for (const op of Object.keys(EXCEL_OPS)) assert.ok(edit.description.includes(`"${op}"`), `설명에 ${op} 없음`);
      assert.match(edit.description, /Excel이 설치돼 있어야 함/);
      assert.match(inspect.description, /[가-힣]/);
      const src = await fresh('mcp');
      const call = (name, args) => client.callTool({ name, arguments: args }, undefined, { timeout: 300_000 });
      const ok = await call('excel_edit', { path: src, engine: 'file', operations: [{ op: 'set_cells', sheet: '매출', cells: { F1: '메모' } }, { op: 'freeze', sheet: '매출', cell: 'B2' }] });
      assert.ok(!ok.isError, ok.content[0].text);
      assert.match(ok.content[0].text, /1\. 셀 1개 씀[\s\S]*2\. 틀 고정 B2[\s\S]*저장: .*mcp-.*\.xlsx \(엔진: 파일/);
      const info = await call('excel_inspect', { path: src, engine: 'file' });
      assert.match(info.content[0].text, /틀 고정 B2/);
      const bad = await call('excel_edit', { path: src, operations: [{ op: 'nope' }] });
      assert.equal(bad.isError, true);
      assert.match(bad.content[0].text, /알 수 없는 작업\(op\)/);
      // 모르는 인수·범위 밖 값·빈 목록도 한국어로
      const extra = await call('excel_edit', { path: src, engine: 'file', operations: [{ op: 'conditional_format', range: 'A1:A3', rule: { type: 'greater_than', value: 1, color: '#00FF00' } }] });
      assert.equal(extra.isError, true);
      assert.match(extra.content[0].text, /모르는 인수입니다: color — rule에 쓸 수 있는 인수: type, value, value2, fill, font_color/);
      const extraOp = await call('excel_edit', { path: src, engine: 'file', operations: [{ op: 'add_table', range: 'A1:B3', header: false }] });
      assert.match(extraOp.content[0].text, /모르는 인수입니다: header — add_table에 쓸 수 있는 인수: sheet, range, name, style/);
      for (const args of [{ operations: [] }, { operations: [{ op: 'insert_rows', at: 0 }] }]) {
        const r = await call('excel_edit', { path: src, engine: 'file', ...args });
        assert.equal(r.isError, true);
        assert.doesNotMatch(r.content[0].text, /Too small|expected/, r.content[0].text);
        assert.match(r.content[0].text, /[가-힣]/);
      }
      const fail = await call('excel_edit', { path: src, engine: 'file', operations: [{ op: 'set_cells', sheet: '없는 시트', cells: { A1: 1 } }] });
      assert.equal(fail.isError, true);
      assert.match(fail.content[0].text, /1번째 작업\(set_cells\) 실패: 시트 없음/);
      const exists = await call('excel_edit', { path: src, engine: 'file', output: src.replace('.xlsx', '-x.xlsx'), operations: [{ op: 'set_cells', cells: { A1: 1 } }] });
      assert.ok(!exists.isError);
      const again = await call('excel_edit', { path: src, engine: 'file', output: src.replace('.xlsx', '-x.xlsx'), operations: [{ op: 'set_cells', cells: { A1: 1 } }] });
      assert.match(again.content[0].text, /이미 있는 파일입니다/);
    } finally { await client.close(); }
  });
});

// ── Excel(COM) 엔진 — OFFICEKIT_COM_TESTS=1 일 때만 ────────────

describe('Excel 엔진', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  const app = async (ops, { name = 'app', src, output } = {}) => {
    src ??= await fresh(name);
    const out = output ?? tmp.path(`${name}-${++seq}-app.xlsx`);
    const r = await excelEdit(src, ops, { output: out, engine: 'app' });
    assert.equal(r.engine, 'excel');
    assert.equal(r.results.length, ops.length);
    return { src, out, r, wb: await load(out) };
  };

  test('셀·서식·병합·틀 고정·필터·표·조건부 서식·입력 제한·이름·너비·높이를 한 번에', COM, async () => {
    const { out, wb, r } = await app([
      { op: 'set_cells', sheet: '매출', cells: { F1: 120, F2: '120', F3: '=F1*2', F4: true, F5: "'=글자", F6: '1/2', B2: null } },
      { op: 'set_range', sheet: '목록', range: 'D2', values: [['이름', '점수'], ['가', 90], ['나', '=E3+1']] },
      { op: 'format', sheet: '매출', range: 'A1:D1', bold: true, font: 'Arial', size: 14, color: '#FFFFFF', fill: '#1F4E78', align: 'center', valign: 'center', wrap: true, border: 'thin', underline: true, italic: true },
      { op: 'format', sheet: '매출', range: 'B2:D7', number_format: '#,##0', border: 'outline' },
      { op: 'column_width', sheet: '매출', columns: 'A', width: 15 },
      { op: 'column_width', sheet: '매출', columns: 'B:C', width: 'auto' },
      { op: 'row_height', sheet: '매출', rows: '1', height: 30 },
      { op: 'row_height', sheet: '매출', rows: '2:3', height: 'auto' },
      { op: 'merge', sheet: '목록', range: 'A5:B6' },
      { op: 'unmerge', sheet: '목록', range: 'A5' },
      { op: 'merge', sheet: '목록', range: 'A8:C8' },
      { op: 'freeze', sheet: '매출', cell: 'B2' },
      { op: 'autofilter', sheet: '매출', range: 'A1:D6' },
      { op: 'add_table', sheet: '목록', range: 'D2:E4', name: '점수표', style: 'TableStyleLight9' },
      { op: 'conditional_format', sheet: '매출', range: 'B2:B6', rule: { type: 'greater_than', value: 40 } },
      { op: 'conditional_format', sheet: '매출', range: 'C2:C6', rule: { type: 'less_than', value: 50 } },
      { op: 'conditional_format', sheet: '매출', range: 'B2:C6', rule: { type: 'between', value: 20, value2: 60 } },
      { op: 'conditional_format', sheet: '매출', range: 'B2:B6', rule: { type: 'equal', value: 30 } },
      { op: 'conditional_format', sheet: '매출', range: 'A2:A6', rule: { type: 'text_contains', value: '산' } },
      { op: 'conditional_format', sheet: '매출', range: 'A2:A6', rule: { type: 'duplicate' } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'color_scale' } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'data_bar' } },
      { op: 'data_validation', sheet: '목록', range: 'G1:G5', list: ['예', '아니오'] },
      { op: 'data_validation', sheet: '목록', range: 'H1:H5', type: 'whole', min: 1, max: 10 },
      { op: 'data_validation', sheet: '목록', range: 'I1:I5', type: 'decimal', min: 0.5 },
      { op: 'named_range', name: '일월', range: '매출!B2:B6' },
      { op: 'clear', sheet: '목록', range: 'B2', what: 'all' },
    ]);
    const ws = wb.getWorksheet('매출');
    assert.deepEqual([ws.getCell('F1').value, ws.getCell('F2').value, ws.getCell('F4').value, ws.getCell('F5').value, ws.getCell('F6').value],
      [120, '120', true, '=글자', '1/2'], '글자는 글자 그대로');
    assert.equal(formulaOf(ws.getCell('F3')), '=F1*2');
    assert.equal(ws.getCell('F3').result, 240);
    assert.equal(ws.getCell('B2').value, null);
    const a1 = ws.getCell('A1');
    assert.deepEqual([a1.font.bold, a1.font.name, a1.font.size, a1.fill.fgColor.argb, a1.alignment.horizontal], [true, 'Arial', 14, 'FF1F4E78', 'center']);
    assert.equal(ws.getCell('B2').numFmt, '#,##0');
    assert.ok(Math.abs(ws.getColumn(1).width - 15) < 1, `너비 ${ws.getColumn(1).width}`);
    assert.equal(ws.getRow(1).height, 30);
    const ls = wb.getWorksheet('목록');
    assert.deepEqual(Object.values(ls._merges).map((m) => m.range), ['A8:C8']);
    assert.equal(ls.getCell('B2').value, null);
    assert.equal(ls.getCell('D2').value, '이름');
    const info = await excelInspect(out, { engine: 'app' });
    const [s1, s2] = info.sheets;
    assert.deepEqual([s1.freeze, s1.conditionalFormats], ['B2', 8]);
    assert.match(s1.autoFilter, /^A1:D[67]$/, 'Excel은 이어진 행까지 필터 범위를 넓히기도 한다');
    assert.deepEqual(s2.tables.map((t) => [t.name, t.range, t.style]), [['점수표', 'D2:E4', 'TableStyleLight9']]);
    assert.equal(s2.merges, 1);
    assert.ok(s2.validations > 0);
    assert.deepEqual(info.names.map((n) => [n.name, n.refersTo]), [['일월', '=매출!$B$2:$B$6']]);
    // 파일 엔진으로 읽어도 같다
    const fi = await excelInspect(out, { engine: 'file' });
    assert.deepEqual([fi.sheets[0].freeze, fi.sheets[0].autoFilter, fi.sheets[0].conditionalFormats, fi.sheets[1].merges], ['B2', 'A1:D6', 8, 1]);
    assert.match(r.results.at(-1), /B2 모두 지움/);
  });

  test('시트 관리: add·copy·move·rename(수식도)·delete', COM, async () => {
    const { wb, r } = await app([
      { op: 'add_sheet', name: '요약', after: '매출' },
      { op: 'copy_sheet', sheet: '매출', name: '매출 사본' },
      { op: 'move_sheet', sheet: '목록', position: 1 },
      { op: 'rename_sheet', sheet: '매출', name: '매출 2026' },
      { op: 'delete_sheet', sheet: '요약' },
    ]);
    assert.deepEqual(wb.worksheets.map((w) => w.name), ['목록', '매출 2026', '매출 사본']);
    assert.equal(formulaOf(wb.getWorksheet('목록').getCell('B1')), "='매출 2026'!D7");
    assert.equal(wb.getWorksheet('매출 사본').getCell('A2').value, '서울');
    assert.match(r.results[1], /시트 복사: '매출' → '매출 사본'/);
  });

  test('행·열 넣기·지우기: Excel이 수식을 옮긴 결과가 파일 엔진과 같다', COM, async () => {
    const ops = [
      { op: 'insert_rows', sheet: '매출', at: 4, count: 2 },
      { op: 'delete_rows', sheet: '매출', at: 2 },
      { op: 'insert_columns', sheet: '매출', at: 'C' },
      { op: 'delete_columns', sheet: '목록', at: 'A' },
      { op: 'named_range', name: '일월', range: '매출!B2:B6' },
    ];
    const src = await fresh('splice');
    const a = await app(ops, { src: (() => { const p = tmp.path('splice-a.xlsx'); copyFileSync(src, p); return p; })() });
    const f = await fileEdit(ops, { src: (() => { const p = tmp.path('splice-f.xlsx'); copyFileSync(src, p); return p; })() });
    for (const [sheet, cells] of [['매출', ['A1', 'D1', 'E5', 'B8', 'C8', 'E8', 'E2']], ['목록', ['A1', 'A2', 'A3']]]) {
      for (const c of cells) {
        const x = a.wb.getWorksheet(sheet).getCell(c); const y = f.wb.getWorksheet(sheet).getCell(c);
        // Excel은 따옴표가 필요 없는 시트 이름의 따옴표를 뺀다
        const norm = (v) => (typeof v === 'string' ? v.replace(/'([\w.\u00A0-\uFFFF]+)'!/g, '$1!') : v);
        assert.equal(norm(formulaOf(y) ?? y.value), norm(formulaOf(x) ?? x.value), `${sheet}!${c}`);
      }
    }
  });

  test('정렬 (머리글 제외, 내림차순)', COM, async () => {
    const { wb, r } = await app([{ op: 'sort', sheet: '매출', range: 'A1:C6', by: 'B', order: 'desc', header: true }]);
    assert.deepEqual([2, 3, 4, 5, 6].map((n) => wb.getWorksheet('매출').getCell(`A${n}`).value), ['광주', '인천', '대구', '부산', '서울']);
    assert.match(r.results[0], /B열 내림차순 \(첫 행 머리글 제외\)/);
  });

  test('recalculate · 틀 고정 해제 · 필터 해제', COM, async () => {
    const { out } = await app([{ op: 'freeze', sheet: '매출', cell: 'B2' }, { op: 'autofilter', sheet: '매출', range: 'A1:D6' }]);
    const { wb, r } = await app([{ op: 'freeze', sheet: '매출', cell: null }, { op: 'autofilter', sheet: '매출', range: null }, { op: 'recalculate' }], { src: out });
    assert.ok(!wb.getWorksheet('매출').autoFilter);
    assert.notEqual(wb.getWorksheet('매출').views[0]?.state, 'frozen');
    assert.equal(wb.getWorksheet('매출').getCell('D7').result, 550);
    assert.equal(r.results[2], '모두 다시 계산함');
  });

  test('한 번에 저장: 중간 작업이 실패하면 저장하지 않음 (한국어 오류)', COM, async () => {
    const src = await fresh('app-atomic');
    const before = readFileSync(src);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '바뀜' } }, { op: 'add_sheet', name: '새' }, { op: 'set_cells', sheet: '없음', cells: { A1: 1 } }], { engine: 'app' }),
      /3번째 작업\(set_cells\) 실패: 시트 없음: 없음 — 저장하지 않았습니다/);
    assert.ok(before.equals(readFileSync(src)));
    await assert.rejects(excelEdit(src, [{ op: 'merge', range: 'A1:B2' }, { op: 'set_cells', cells: { B2: 1 } }], { engine: 'app' }),
      /2번째 작업\(set_cells\) 실패: B2는 병합된 칸의 일부라 쓸 수 없습니다 — 병합 영역의 왼쪽 위 칸 A1에 쓰세요/);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '=SUM(' } }], { engine: 'app' }), /수식이 올바르지 않습니다/);
    await assert.rejects(excelEdit(src, [{ op: 'add_sheet', name: '목록' }], { engine: 'app' }), /같은 이름의 시트가 이미 있습니다/);
    assert.ok(before.equals(readFileSync(src)));
  });

  test('차트와 조건부 서식을 그림으로 확인', COM, async () => {
    const { out, r } = await app([
      { op: 'add_chart', sheet: '매출', range: 'A1:C6', type: 'column', title: '지점별 매출', at: 'A9', width: 360, height: 220 },
      { op: 'add_chart', sheet: '목록', range: 'A1:B3', type: 'pie' },
    ]);
    assert.match(r.results[0], /차트 만듦: 세로 막대형 'Chart 1' \(원본 A1:C6, 위치 A9, 계열 2개\(열마다\), 제목 '지점별 매출'\)/);
    assert.match(r.results[1], /계열 1개\(열마다\)/);
    const info = await excelInspect(out, { engine: 'app' });
    assert.deepEqual(info.sheets[0].charts.map((c) => [c.name, c.title]), [['Chart 1', '지점별 매출']]);
    assert.equal(info.sheets[1].charts.length, 1);
    const img = await pageImage(out, 1, { dpi: 96 });
    assert.ok(isPng(img.png));
    writeFileSync(tmp.path('chart.png'), img.png);
    const px = decodePng(img.png);
    // 막대(파랑·빨강 계열)가 그려졌는지
    const bars = countPixels(px, (rr, g, b) => (b > 140 && b - rr > 60) || (rr > 150 && rr - g > 70 && rr - b > 70));
    assert.ok(bars > 1500, `막대 픽셀 ${bars}`);

    const cf = await app([
      { op: 'conditional_format', sheet: '매출', range: 'B2:C6', rule: { type: 'greater_than', value: 45 } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'data_bar', fill: '#63BE7B' } },
    ]);
    const cfImg = decodePng((await pageImage(cf.out, 1, { dpi: 96 })).png);
    assert.ok(countPixels(cfImg, (rr, g, b) => near(rr, g, b, 'FFC7CE', 6)) > 400, '강조 채우기(#FFC7CE)가 보이지 않음');
    assert.ok(countPixels(cfImg, (rr, g, b) => near(rr, g, b, '63BE7B', 25)) > 100, '데이터 막대(#63BE7B)가 보이지 않음');
  });

  test('파일 엔진 결과도 Excel이 그대로 연다 (모든 ○ 작업)', COM, async () => {
    const { out } = await fileEdit([
      { op: 'add_sheet', name: '요약' }, { op: 'set_range', sheet: '요약', range: 'A1', values: [['이름', '점수'], ['가', 3], ['나', 5]] },
      { op: 'set_cells', sheet: '요약', cells: { D1: '=SUM(B2:B3)', D2: '=매출!D7' } },
      { op: 'format', sheet: '매출', range: 'A1:D1', bold: true, fill: '#DDEBF7', border: 'thin' },
      { op: 'column_width', sheet: '요약', columns: 'A:B', width: 'auto' }, { op: 'row_height', sheet: '매출', rows: '1', height: 24 },
      { op: 'merge', sheet: '목록', range: 'D1:E1' }, { op: 'freeze', sheet: '매출', cell: 'A2' },
      { op: 'autofilter', sheet: '요약', range: 'A1:B3' }, { op: 'add_table', sheet: '목록', range: 'A10:B12', name: '목록T' },
      { op: 'conditional_format', sheet: '매출', range: 'B2:C6', rule: { type: 'greater_than', value: 45 } },
      { op: 'conditional_format', sheet: '매출', range: 'D2:D6', rule: { type: 'data_bar' } },
      { op: 'conditional_format', sheet: '매출', range: 'A2:A6', rule: { type: 'duplicate' } },
      { op: 'data_validation', sheet: '요약', range: 'C2:C3', list: ['예', '아니오'] },
      { op: 'named_range', name: '합계', range: '매출!D7' },
      { op: 'insert_rows', sheet: '매출', at: 3 }, { op: 'delete_columns', sheet: '목록', at: 'C' },
      { op: 'rename_sheet', sheet: '매출', name: '실적' }, { op: 'move_sheet', sheet: '요약', position: 1 },
      { op: 'clear', sheet: '목록', range: 'A1', what: 'all' },
    ]);
    const com = await excelInspect(out, { engine: 'app' });
    const file = await excelInspect(out, { engine: 'file' });
    assert.deepEqual(com.sheets.map((s) => s.name), ['요약', '실적', '목록']);
    for (const k of ['freeze', 'autoFilter', 'merges', 'conditionalFormats']) {
      assert.deepEqual(com.sheets.map((s) => s[k]), file.sheets.map((s) => s[k]), k);
    }
    assert.deepEqual(com.sheets[2].tables.map((t) => [t.name, t.range]), file.sheets[2].tables.map((t) => [t.name, t.range]));
    assert.deepEqual(com.names.map((n) => n.refersTo), ['=실적!$D$8']);
    const { readDocument } = await import('../src/core/index.mjs');
    const rows = (await readDocument(out, { engine: 'app' })).sheets[0].rows;
    assert.deepEqual([rows[0][3], rows[1][3]], ['8', '550']);
  });

  test('글자는 Excel이 고른 표시 형식 없이 글자로 · 병합된 칸 일부에 붙여 넣기는 거절', COM, async () => {
    const { wb } = await app([
      { op: 'set_cells', sheet: '목록', cells: { F1: '1/2', F2: '15%', F3: '2026-10-18', F4: '1e5', F5: '0123' } },
      { op: 'set_range', sheet: '목록', range: 'G1', values: [['1/2', '15%', '이름'], [3, '0123', '=1+1']] },
    ]);
    const ls = wb.getWorksheet('목록');
    for (const a of ['F1', 'F2', 'F3', 'F4', 'F5', 'G1', 'H1', 'H2']) {
      assert.equal(typeof ls.getCell(a).value, 'string', a);
      assert.ok(!ls.getCell(a).numFmt || ls.getCell(a).numFmt === 'General', `${a} 표시 형식 ${ls.getCell(a).numFmt}`);
    }
    assert.equal(ls.getCell('G2').value, 3);
    const src = await fresh('app-merge');
    const before = readFileSync(src);
    await assert.rejects(excelEdit(src, [{ op: 'merge', sheet: '매출', range: 'A1:D1' }, { op: 'set_range', sheet: '매출', range: 'A1', values: [['제목', '부제']] }], { engine: 'app' }),
      /2번째 작업\(set_range\) 실패: B1는 병합된 칸의 일부라 쓸 수 없습니다/);
    assert.ok(before.equals(readFileSync(src)));
  });

  test('표 이름이 이미 있으면 거절 (Excel이 몰래 이름을 바꾸지 않게)', COM, async () => {
    const src = await fresh('app-tname');
    await assert.rejects(excelEdit(src, [
      { op: 'add_table', sheet: '매출', range: 'A1:D6', name: '매출표' },
      { op: 'add_table', sheet: '목록', range: 'D1:E3', name: '매출표' },
    ], { engine: 'app' }), /2번째 작업\(add_table\) 실패: 이미 있는 표·이름입니다: 매출표/);
    await assert.rejects(excelEdit(src, [{ op: 'named_range', name: '합', range: '매출!D7' }, { op: 'add_table', sheet: '목록', range: 'D1:E3', name: '합' }], { engine: 'app' }),
      /이미 있는 표·이름입니다: 합/);
  });

  test('차트: 세로로 긴 범위는 열마다 계열, 제목을 안 주면 자리 표시 제목(Chart Title) 없음', COM, async () => {
    const src = tmp.path('tall.xlsx');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('지점');
    ws.addRow(['지점', '1월', '2월', '3월']);
    for (let i = 0; i < 20; i++) ws.addRow([`지점${i + 1}`, 100 + i * 7, 120 + i * 3, 90 + i * 11]);
    await wb.xlsx.writeFile(src);
    const { out, r } = await app([
      { op: 'add_chart', range: 'A1:D21', type: 'column', at: 'F2' },
      { op: 'add_chart', range: 'A1:D21', type: 'line', at: 'F20', series_in: 'rows', title: '행마다' },
    ], { src });
    assert.match(r.results[0], /계열 3개\(열마다\), 제목 없음/);
    assert.match(r.results[1], /계열 20개\(행마다\), 제목 '행마다'/);
    const info = await excelInspect(out, { engine: 'app' });
    assert.deepEqual(info.sheets[0].charts.map((c) => c.title), [null, '행마다']);
    assert.doesNotMatch(inspectText(info), /Chart Title|차트 제목/);
  });

  test('다른 프로그램이 열어 둔 파일을 제자리에 고치면 오류 (조용히 저장 안 하는 일 없음)', COM, async () => {
    const src = await fresh('app-locked');
    const before = readFileSync(src);
    const { spawn } = await import('node:child_process');
    const ps = spawn('powershell.exe', ['-NoProfile', '-Command', `$f=[System.IO.File]::Open('${src}','Open','Read','Read'); Start-Sleep 120; $f.Close()`]);
    try {
      await new Promise((ok) => setTimeout(ok, 2500));
      await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: '바뀜' } }], { engine: 'app' }), /다른 프로그램에서 열려 있어/);
      assert.ok(before.equals(readFileSync(src)));
    } finally { ps.kill(); }
  });

  test('BOM 없는 UTF-8 csv를 Excel로 고쳐도 글자가 깨지지 않고 인코딩·줄바꿈 유지', COM, async () => {
    const src = tmp.path('nobom-app.csv');
    writeFileSync(src, '이름,점수,메모\n김철수,90,"좋음, 매우"\n이영희,85,보통\n');
    const r = await excelEdit(src, [{ op: 'sort', range: 'A1:C3', by: 'B', header: true }], { engine: 'auto' });
    assert.equal(r.engine, 'excel');
    assert.equal(readFileSync(src, 'utf8'), '이름,점수,메모\n이영희,85,보통\n김철수,90,"좋음, 매우"\n');
  });

  test('여러 시트 문서를 csv로: 고친 시트를 저장, 여러 시트를 고치면 거절', COM, async () => {
    const src = await fresh('app-tocsv');
    const out = tmp.path('tocsv.csv');
    const r = await excelEdit(src, [{ op: 'set_cells', sheet: '목록', cells: { A5: '추가' } }], { engine: 'app', output: out });
    assert.ok(r.notes.includes("csv에는 시트 '목록'의 값만 저장됩니다"), r.notes.join(' / '));
    assert.match(readFileSync(out, 'utf8'), /^﻿총합,[\s\S]*추가/);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', sheet: '목록', cells: { A5: 'x' } }, { op: 'set_cells', cells: { A9: 'y' } }], { engine: 'app', output: tmp.path('tocsv2.csv') }),
      /csv에는 시트 하나의 값만 저장되는데 작업이 여러 시트\('목록', '매출'\)를 고칩니다/);
  });

  test('Excel의 영어 오류도 한국어로', COM, async () => {
    const src = await fresh('app-errs');
    const before = readFileSync(src);
    const cases = [
      [[{ op: 'merge', sheet: '매출', range: 'A2:A3' }, { op: 'sort', sheet: '매출', range: 'A1:D6', by: 'B', header: true }], /2번째 작업\(sort\) 실패: 병합된 칸이 있는 범위는 정렬할 수 없습니다/],
      [[{ op: 'merge', sheet: '매출', range: 'A2:B3' }, { op: 'clear', sheet: '매출', range: 'A2' }], /2번째 작업\(clear\) 실패: 병합된 영역의 일부만 지울 수 없습니다/],
      [[{ op: 'add_table', sheet: '매출', range: 'A1:D6' }, { op: 'delete_rows', sheet: '매출', at: 1 }], /2번째 작업\(delete_rows\) 실패: 표 '.+'의 머리글 행\(1행\)은 지울 수 없습니다/],
      [[{ op: 'set_cells', sheet: '매출', cells: { A1048576: 1 } }, { op: 'insert_rows', sheet: '매출', at: 3 }], /2번째 작업\(insert_rows\) 실패: 시트 끝에 있는 내용이 밀려나므로 행을 넣을 수 없습니다/],
    ];
    for (const [ops, re] of cases) await assert.rejects(excelEdit(src, ops, { engine: 'app' }), re);
    await assert.rejects(excelEdit(src, [{ op: 'set_cells', cells: { A1: 1 } }], { engine: 'app', output: tmp.path('com1.xlsx') }), /문서를 저장하지 못했습니다/);
    assert.ok(before.equals(readFileSync(src)));
  });

  test('csv를 Excel로 고치면 UTF-8 csv로 저장', COM, async () => {
    const src = tmp.path('excel.csv');
    writeFileSync(src, '\uFEFF이름,점수\r\n가,1\r\n나,2\r\n');
    const r = await excelEdit(src, [{ op: 'set_cells', cells: { C1: '합', C2: '=B2*10' } }, { op: 'sort', range: 'A1:C3', by: 'B', order: 'desc', header: true }], { engine: 'app' });
    assert.equal(r.engine, 'excel');
    const text = readFileSync(src, 'utf8').replace(/^\uFEFF/, '');
    assert.match(text, /^이름,점수,합\r?\n나,2,\r?\n가,1,10/);
  });
});
