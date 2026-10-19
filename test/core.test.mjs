// core.test.mjs — 문서 엔진·쪽 그림·명령줄 테스트.
// 기본은 파일 엔진만 쓴다. Word/PowerPoint/Excel COM 테스트는 OFFICEKIT_COM_TESTS=1 일 때만 (한컴 COM은 쓰지 않음).
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { execFile } from 'node:child_process';
import { createRequire } from 'node:module';
import { inflateSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import {
  createDocument, readDocument, documentInfo, listFields, fillFields, replaceText, setCells, convertDocument,
} from '../src/core/index.mjs';
import { pageImage } from '../src/core/render.mjs';
import { MANUSCRIPT, COM_TESTS, ROOT, tempDir, makePdf, isPng, pngSize, xlsxWithChart, zipParts, zipText } from './helpers/fixtures.mjs';
import { Document, Packer, Paragraph, Header, TextRun, Tab, Table, TableRow, TableCell } from 'docx';

const FILE = { engine: 'file' };
const tmp = tempDir();
after(() => tmp.cleanup());

const textOf = async (path) => (await readDocument(path, FILE)).text;

// ── 다섯 형식 공통: 만들기·읽기·칸·바꾸기·채우기 ─────────────

for (const ext of ['hwp', 'hwpx', 'docx', 'pptx', 'xlsx']) {
  describe(`파일 엔진 .${ext}`, () => {
    const src = tmp.path(`doc.${ext}`);
    const rep = tmp.path(`rep.${ext}`);

    before(async () => { await createDocument(src, MANUSCRIPT); });

    test('만들기·읽기', async () => {
      assert.ok(existsSync(src));
      const r = await readDocument(src, FILE);
      assert.equal(r.engine, ext.startsWith('hwp') ? 'rhwp' : 'file');
      for (const word of ['분기 보고', '매출', '비용', '신규 고객 확보']) assert.ok(r.text.includes(word), `${word} 없음:\n${r.text}`);
      if (ext.startsWith('hwp')) { assert.equal(r.pages, 2); assert.deepEqual(r.tables[0][1], ['매출', '120']); }
      if (ext === 'docx') assert.deepEqual(r.tables[0][0], ['항목', '값']);
      if (ext === 'pptx') assert.equal(r.slides.length, 2);
      if (ext === 'xlsx') assert.equal(r.sheets[0].name, 'Sheet1');
    });

    test('칸 목록에 담당자', async () => {
      const fields = await listFields(src, FILE);
      const f = fields.find((x) => x.name === '담당자');
      assert.ok(f, JSON.stringify(fields));
      assert.equal(f.kind, ext.startsWith('hwp') ? '누름틀' : 'placeholder');
    });

    test('찾아 바꾸기 (다른 파일로 저장, 원본 유지)', async () => {
      const r = await replaceText(src, [['매출', '수익'], { find: '비용', replace: '지출' }], { output: rep, engine: 'file' });
      assert.ok(r.count >= 4, `count=${r.count}`);
      const after = await textOf(rep);
      assert.ok(after.includes('수익') && after.includes('지출'));
      assert.ok(!after.includes('매출') && !after.includes('비용'));
      assert.ok((await textOf(src)).includes('매출'), '원본이 바뀜');
    });

    test('칸 채우기 (제자리 저장)', async () => {
      const r = await fillFields(rep, { 담당자: '홍길동', 없는칸: 'x' }, FILE);
      assert.deepEqual(r.filled, ['담당자']);
      assert.deepEqual(r.missing, ['없는칸']);
      const after = await textOf(rep);
      assert.ok(after.includes('홍길동'), after);
      assert.ok(!after.includes('{{담당자}}'));
    });
  });
}

// ── 엑셀 셀 ─────────────────────────────────────────────

describe('엑셀 셀 쓰기', () => {
  const src = tmp.path('cells.xlsx');
  before(async () => { await createDocument(src, MANUSCRIPT); });

  test('값과 수식', async () => {
    const out = tmp.path('cells-out.xlsx');
    const r = await setCells(src, { D1: 5, D2: '=D1*2', E1: '메모' }, { sheet: 'Sheet1', output: out, engine: 'file' });
    assert.equal(r.updated, 3);
    const rows = (await readDocument(out, FILE)).sheets[0].rows;
    assert.equal(rows[0][3], '5');
    assert.equal(rows[1][3], '=D1*2');
    assert.equal(rows[0][4], '메모');
  });

  test('병합된 칸은 값을 한 번만 (왼쪽 위 칸), 시트 머리에 병합 범위 표시', async () => {
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('표');
    ws.addRows([['2027 예산 현황', '', ''], ['부서', '예산', '집행']]);
    ws.mergeCells('A1:C1');
    const p = tmp.path('merged.xlsx');
    await wb.xlsx.writeFile(p);
    const d = await readDocument(p, FILE);
    assert.deepEqual(d.sheets[0].rows[0], ['2027 예산 현황', '', '']);
    assert.deepEqual(d.sheets[0].merges, ['A1:C1']);
    assert.match(d.text, /^--- 시트 표 --- \(병합: A1:C1 — 값은 왼쪽 위 칸에만\)\n2027 예산 현황\t\t\n/);
  });

  test('없는 시트·엑셀 아닌 파일은 한국어 오류', async () => {
    await assert.rejects(setCells(src, { A1: 1 }, { sheet: '없음', engine: 'file', output: tmp.path('x.xlsx') }), /시트 없음/);
    const docx = tmp.path('cells.docx');
    await createDocument(docx, '# 표');
    await assert.rejects(setCells(docx, { A1: 1 }, FILE), /엑셀 파일에서만/);
  });
});

// ── hwp ↔ hwpx 변환 ─────────────────────────────────────

describe('한글 변환 (프로그램 없이)', () => {
  test('hwp → hwpx → hwp', async () => {
    const hwp = tmp.path('conv.hwp');
    await createDocument(hwp, MANUSCRIPT);
    const hwpx = tmp.path('conv.hwpx');
    const a = await convertDocument(hwp, hwpx);
    assert.ok(['rhwp', 'hwp'].includes(a.engine));
    assert.equal(readFileSync(hwpx).subarray(0, 2).toString(), 'PK', 'hwpx는 ZIP이어야 함');
    const back = tmp.path('conv-back.hwp');
    await convertDocument(hwpx, back);
    const t1 = await textOf(hwp);
    assert.equal(await textOf(hwpx), t1);
    assert.equal(await textOf(back), t1);
    const info = await documentInfo(back);
    assert.equal(info.pages, 2);
    assert.deepEqual(info.fields, ['담당자']);
  });

  test('없는 파일은 한국어 오류', async () => {
    await assert.rejects(convertDocument(tmp.path('없음.hwp'), tmp.path('a.hwpx')), /파일이 없습니다/);
    await assert.rejects(readDocument(tmp.path('없음.docx')), /파일이 없습니다/);
  });
});

// ── 쪽 그림 ──────────────────────────────────────────────

describe('쪽 그림 pageImage', () => {
  for (const ext of ['hwp', 'hwpx']) {
    test(`.${ext} (rhwp + resvg)`, async () => {
      const f = tmp.path(`img.${ext}`);
      await createDocument(f, MANUSCRIPT);
      const r = await pageImage(f, 1, { dpi: 96 });
      assert.ok(isPng(r.png));
      assert.equal(r.engine, 'rhwp');
      assert.equal(r.page, 1);
      assert.equal(r.pages, 2);
      const { width, height } = pngSize(r.png);
      assert.ok(Math.abs(width - 794) <= 2 && Math.abs(height - 1123) <= 2, `${width}x${height}`); // A4 @96dpi
      const big = await pageImage(f, 2, { dpi: 192 });
      assert.ok(Math.abs(pngSize(big.png).width - 1588) <= 3);
      await assert.rejects(pageImage(f, 3), /범위를 벗어났습니다.*전체 2쪽/);
      await assert.rejects(pageImage(f, 0), /범위/);
    });
  }

  test('.pdf (Windows.Data.Pdf, 프로그램 불필요)', { skip: process.platform !== 'win32' && 'Windows 전용' }, async () => {
    const pdf = tmp.path('gen.pdf');
    writeFileSync(pdf, makePdf(3));
    const r = await pageImage(pdf, 2, { dpi: 96 });
    assert.ok(isPng(r.png));
    assert.equal(r.engine, 'pdf');
    assert.equal(r.pages, 3);
    assert.equal(r.page, 2);
    const { width, height } = pngSize(r.png);
    assert.ok(Math.abs(width - 793) <= 3 && Math.abs(height - 1123) <= 3, `${width}x${height}`); // 595x842pt
    const hi = await pageImage(pdf, 1, { dpi: 144 });
    assert.ok(Math.abs(pngSize(hi.png).width - 1190) <= 4, JSON.stringify(pngSize(hi.png)));
    await assert.rejects(pageImage(pdf, 4), /범위를 벗어났습니다.*전체 3쪽/);
  });

  test('없는 파일', async () => {
    await assert.rejects(pageImage(tmp.path('없음.hwp')), /파일이 없습니다/);
  });

  // PowerPoint는 쪽 = 슬라이드 번호 (숨긴 슬라이드 포함). 범위·암호는 PowerPoint를 띄우기 전에 거른다
  test('.pptx 범위·암호 확인 (PowerPoint 없이)', async () => {
    const f = await slideDeck(tmp.path('range.pptx'));
    await assert.rejects(pageImage(f, 4), /범위를 벗어났습니다: 4쪽 \(전체 3쪽/); // 숨긴 슬라이드 2도 센다
    await assert.rejects(pageImage(f, 0), /범위를 벗어났습니다: 0쪽 \(전체 3쪽/);
    const odp = tmp.path('range.odp');
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation');
    zip.file('content.xml', '<office:document-content><office:body><office:presentation><draw:page draw:name="1"/><draw:page draw:name="2">'
      + '<presentation:notes><draw:page-thumbnail/></presentation:notes></draw:page></office:presentation></office:body></office:document-content>');
    writeFileSync(odp, await zip.generateAsync({ type: 'nodebuffer' }));
    await assert.rejects(pageImage(odp, 3), /범위를 벗어났습니다: 3쪽 \(전체 2쪽/);
    // 열기 암호(암호화된 꾸러미)는 보이지 않는 암호 창에서 멈추므로 바로 거절
    const locked = tmp.path('locked.pptx');
    writeFileSync(locked, Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), Buffer.alloc(504),
      Buffer.from('EncryptedPackage', 'utf16le'), Buffer.alloc(64)]));
    await assert.rejects(pageImage(locked, 1), /암호가 걸린 문서는 열 수 없습니다/);
  });
});

// 슬라이드 3장 (1 빨강, 2 초록·숨김, 3 파랑 배경). 숨긴 슬라이드는 PDF에서 빠지므로 쪽 그림 번호 확인용
async function slideDeck(out) {
  const { default: JSZip } = await import('jszip');
  await createDocument(out, '# 하나\n# 둘 (숨김)\n# 셋');
  const zip = await JSZip.loadAsync(readFileSync(out));
  for (const [n, color] of [[1, 'FF0000'], [2, '00FF00'], [3, '0000FF']]) {
    const name = `ppt/slides/slide${n}.xml`;
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<p:cSld\b[^>]*>/, (m) => `${m}<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`);
    if (n === 2) xml = xml.replace(/<p:sld\b/, '<p:sld show="0"');
    zip.file(name, xml);
  }
  writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer' }));
  return out;
}

// PNG 한 점의 색 → 'rrggbb' (비월 주사 없는 PNG: 회색·RGB·팔레트(1~8비트)·RGBA, 필터 풀기).
// PowerPoint가 내보낸 그림은 팔레트 PNG일 때가 많다
function pixelAt(png, x, y) {
  let pos = 8; const idat = []; let w; let depth; let ct; let plte;
  while (pos < png.length) {
    const len = png.readUInt32BE(pos); const type = png.toString('ascii', pos + 4, pos + 8);
    const data = png.subarray(pos + 8, pos + 8 + len);
    if (type === 'IHDR') { w = data.readUInt32BE(0); depth = data[8]; ct = data[9]; assert.equal(data[12], 0, '비월 주사 PNG는 못 읽음'); }
    else if (type === 'PLTE') plte = data;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    pos += 12 + len;
  }
  const bits = depth * { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ct];
  assert.ok(bits <= 32 && depth <= 8, `PNG 형식 ${ct}/${depth}비트`);
  const bpp = Math.max(1, bits / 8);
  const stride = Math.ceil((w * bits) / 8);
  const raw = inflateSync(Buffer.concat(idat));
  let prev = Buffer.alloc(stride); let row;
  for (let r = 0; r <= y; r++) {
    const f = raw[r * (stride + 1)]; row = Buffer.alloc(stride);
    for (let i = 0; i < stride; i++) {
      const a = i >= bpp ? row[i - bpp] : 0; const b = prev[i]; const c = i >= bpp ? prev[i - bpp] : 0;
      let v = raw[r * (stride + 1) + 1 + i];
      if (f === 1) v += a; else if (f === 2) v += b; else if (f === 3) v += (a + b) >> 1;
      else if (f === 4) { const p = a + b - c; const pa = Math.abs(p - a); const pb = Math.abs(p - b); const pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      row[i] = v & 255;
    }
    prev = row;
  }
  const hex = (vals) => vals.map((v) => v.toString(16).padStart(2, '0')).join('');
  if (depth < 8) { // 팔레트·회색: 한 바이트에 여러 점
    const at = x * depth; const v = (row[at >> 3] >> (8 - depth - (at & 7))) & ((1 << depth) - 1);
    return ct === 3 ? hex([...plte.subarray(v * 3, v * 3 + 3)]) : hex(Array(3).fill(Math.round((v * 255) / ((1 << depth) - 1))));
  }
  if (ct === 3) return hex([...plte.subarray(row[x] * 3, row[x] * 3 + 3)]);
  if (ct === 0 || ct === 4) return hex(Array(3).fill(row[x * bpp]));
  return hex([row[x * bpp], row[x * bpp + 1], row[x * bpp + 2]]);
}

// ── 바꾸기·채우기 세부 동작 (파일 엔진) ─────────────────────

const ENGLISH = '# 보고서\nAI 도입 계획입니다.\nThe main goal was said by email.\n';

describe('바꾸기는 영문 대소문자를 구분 (모든 형식 같음)', () => {
  for (const ext of ['hwp', 'hwpx', 'docx', 'pptx', 'xlsx']) {
    test(`.${ext}`, async () => {
      const src = tmp.path(`case.${ext}`);
      await createDocument(src, ENGLISH);
      const r = await replaceText(src, [['AI', '인공지능'], ['ai', 'X']], { output: tmp.path(`case-out.${ext}`), engine: 'file' });
      assert.deepEqual(r.counts, [1, 3], JSON.stringify(r));
      assert.equal(r.count, 4);
      const t = await textOf(tmp.path(`case-out.${ext}`));
      assert.ok(t.includes('인공지능 도입'), t);
      assert.ok(t.includes('The mXn goal was sXd by emXl.'), t);
    });
  }
});

describe('docx: 여러 조각(run)에 걸친 자리표시자', () => {
  test('탭 사이에 쪼개진 {{담당자}} → 그 조각만 고치고 탭·서식은 그대로', async () => {
    const src = tmp.path('tab.docx');
    writeFileSync(src, await Packer.toBuffer(new Document({ sections: [{ children: [
      new Paragraph({ children: [new TextRun({ text: '성명', bold: true }), new TextRun({ children: [new Tab()] }),
        new TextRun('{{담'), new TextRun('당자}}'), new TextRun({ children: [new Tab()] }), new TextRun('(서명)')] }),
    ] }] })));
    assert.equal(await textOf(src), '성명\t{{담당자}}\t(서명)');
    const out = tmp.path('tab-out.docx');
    const r = await fillFields(src, { 담당자: '홍길동' }, { output: out, engine: 'file' });
    assert.deepEqual(r.filled, ['담당자']);
    assert.equal(await textOf(out), '성명\t홍길동\t(서명)');
    const xml = await zipText(out, 'word/document.xml');
    // 굵게는 '성명' 조각에만
    const runs = [...xml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((m) => m[1]);
    assert.ok(runs.find((x) => x.includes('성명')).includes('<w:b/>'));
    assert.ok(!runs.find((x) => x.includes('홍길동')).includes('<w:b/>'), xml);
  });

  test('한 문단에 여러 번, 조각 경계에 걸친 바꾸기', async () => {
    const src = tmp.path('runs.docx');
    writeFileSync(src, await Packer.toBuffer(new Document({ sections: [{ children: [
      new Paragraph({ children: [new TextRun('가나'), new TextRun({ text: '다라가', italics: true }), new TextRun('나다')] }),
    ] }] })));
    const r = await replaceText(src, [['가나다', '①']], { output: tmp.path('runs-out.docx'), engine: 'file' });
    assert.equal(r.count, 2);
    assert.equal(await textOf(tmp.path('runs-out.docx')), '①라①');
  });
});

describe('칸 채우기: 머리글·띄어 쓴 자리표시자', () => {
  test('docx 머리글 {{작성자}}, 본문 {{ 담당자 }}', async () => {
    const src = tmp.path('header.docx');
    writeFileSync(src, await Packer.toBuffer(new Document({ sections: [{
      headers: { default: new Header({ children: [new Paragraph('작성자: {{작성자}}')] }) },
      children: [new Paragraph('본문 {{ 담당자 }} 입니다'), new Paragraph('날짜: {{날짜}}')],
    }] })));
    const names = (await listFields(src, FILE)).map((f) => f.name).sort();
    assert.deepEqual(names, ['날짜', '담당자', '작성자']);
    const out = tmp.path('header-out.docx');
    const r = await fillFields(src, { 작성자: '김선호', 담당자: '홍길동', 날짜: '2026-10-17', 없음: 'x' }, { output: out, engine: 'file' });
    assert.deepEqual(r.filled, ['작성자', '담당자', '날짜']);
    assert.deepEqual(r.missing, ['없음']);
    const header = (await zipParts(out)).find((f) => /^word\/header\d*\.xml$/.test(f));
    assert.match(await zipText(out, header), /김선호/);
    const body = await textOf(out);
    assert.ok(body.includes('본문 홍길동 입니다') && !body.includes('{{'), body);
  });

  test('hwp 본문 {{ 이름 }} (띄어 씀)', async () => {
    const src = tmp.path('space.hwp');
    await createDocument(src, '# 제목\n신청인 {{ 이름 }} 입니다');
    const r = await fillFields(src, { 이름: '이영희' }, { output: tmp.path('space-out.hwpx'), engine: 'file' });
    assert.deepEqual(r.filled, ['이름']);
    assert.ok((await textOf(tmp.path('space-out.hwpx'))).includes('신청인 이영희 입니다'));
  });

  test('결과는 원본과 같은 형식으로만 저장 (다른 확장자는 거부, 파일 안 만듦)', async () => {
    const src = tmp.path('fmt.docx');
    await createDocument(src, MANUSCRIPT);
    for (const bad of ['evil.cmd', 'note.txt', 'x.hwpx']) {
      await assert.rejects(fillFields(src, { 없는칸: 'x' }, { output: tmp.path(bad), engine: 'file' }), /저장할 수 없습니다/);
      assert.ok(!existsSync(tmp.path(bad)), bad);
    }
    await assert.rejects(replaceText(src, [['매출', 'x']], { output: tmp.path('evil.bat'), engine: 'file' }), /저장할 수 없습니다/);
    assert.ok(!existsSync(tmp.path('evil.bat')));
  });
});

describe('xlsx 파일 엔진 쓰기는 차트·그림을 지우지 않음', () => {
  const charted = (p) => zipParts(p).then((parts) => ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml'].every((f) => parts.includes(f)));
  let src;
  before(async () => { src = tmp.path('chart.xlsx'); writeFileSync(src, await xlsxWithChart()); });

  test('셀 쓰기 (값·수식·새 행·글자)', async () => {
    const out = tmp.path('chart-cells.xlsx');
    const r = await setCells(src, { B3: 25, C2: '=B2*2', A9: '새 행', D1: '  앞뒤 공백 ' }, { output: out, engine: 'file' });
    assert.equal(r.updated, 4);
    assert.ok(await charted(out));
    const rows = (await readDocument(out, FILE)).sheets[0].rows;
    assert.equal(rows[2][1], '25');
    assert.equal(rows[1][2], '=B2*2');
    assert.equal(rows[8][0], '새 행');
    assert.equal(rows[0][3], '  앞뒤 공백 ');
    assert.match(await zipText(out, 'xl/workbook.xml'), /fullCalcOnLoad="1"/);
  });

  test('찾아 바꾸기·칸 채우기', async () => {
    const out = tmp.path('chart-rep.xlsx');
    const r = await replaceText(src, [['월', '달']], { output: out, engine: 'file' });
    assert.equal(r.count, 3); // 공유 문자열 '1월','2월','월' 각각 한 셀
    assert.ok(await charted(out));
    const f = await fillFields(src, { 담당자: '홍길동' }, { output: tmp.path('chart-fill.xlsx'), engine: 'file' });
    assert.deepEqual(f.filled, ['담당자']);
    assert.ok(await charted(tmp.path('chart-fill.xlsx')));
    assert.equal((await readDocument(tmp.path('chart-fill.xlsx'), FILE)).sheets[0].rows[3][1], '홍길동');
  });
});

describe('docx 표 속 표', () => {
  test('바깥 표 모양 유지, 안쪽 표는 칸 글자로', async () => {
    const cell = (children) => new TableCell({ children });
    const inner = new Table({ rows: [new TableRow({ children: [cell([new Paragraph('안1')]), cell([new Paragraph('안2')])] })] });
    const outer = new Table({ rows: [
      new TableRow({ children: [cell([new Paragraph('바깥1'), inner, new Paragraph('')]), cell([new Paragraph('바깥2 {{담당자}}')])] }),
      new TableRow({ children: [cell([new Paragraph('바깥3')]), cell([new Paragraph('바깥4')])] }),
    ] });
    const src = tmp.path('nested.docx');
    writeFileSync(src, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('앞 문단'), outer, new Paragraph('뒤 문단')] }] })));
    const d = await readDocument(src, FILE);
    assert.equal(d.tables.length, 1);
    assert.equal(d.tables[0].length, 2);
    assert.match(d.tables[0][0][0], /^바깥1\n안1 \| 안2/);
    assert.deepEqual(d.tables[0][1], ['바깥3', '바깥4']);
    assert.match(d.text, /바깥3 \| 바깥4/);
    assert.ok(d.text.startsWith('앞 문단') && d.text.endsWith('뒤 문단'), d.text);
  });
});

describe('손상된 파일·csv·hwt', () => {
  test('빈 docx/xlsx는 한국어 오류', async () => {
    for (const ext of ['docx', 'pptx', 'xlsx']) {
      const p = tmp.path(`empty.${ext}`);
      writeFileSync(p, '');
      await assert.rejects(readDocument(p, FILE), /손상되었거나 .*문서가 아닙니다/);
    }
  });

  test('csv는 프로그램 없이 읽음 (UTF-8 BOM, CP949, 따옴표)', async () => {
    const a = tmp.path('utf8.csv');
    writeFileSync(a, '\ufeff이름,메모\n홍길동,"쉼표, 있음"\n"김""철수""","줄\n바꿈"\n', 'utf8');
    const r = await readDocument(a);
    assert.equal(r.engine, 'file');
    assert.deepEqual(r.sheets[0].rows, [['이름', '메모'], ['홍길동', '쉼표, 있음'], ['김"철수"', '줄\n바꿈']]);
    const b = tmp.path('cp949.csv');
    writeFileSync(b, Buffer.from([0xc0, 0xcc, 0xb8, 0xa7, 0x2c, 0xb0, 0xaa, 0x0d, 0x0a, 0xc8, 0xab, 0xb1, 0xe6, 0xb5, 0xbf, 0x2c, 0x31]));
    assert.deepEqual((await readDocument(b)).sheets[0].rows, [['이름', '값'], ['홍길동', '1']]);
    assert.deepEqual((await listFields(a)).length, 0);
  });

  test('.hwt (한글 서식 파일) 읽기·칸·쪽 그림·저장', async () => {
    const src = tmp.path('form.hwt');
    const h = tmp.path('form-src.hwp');
    await createDocument(h, MANUSCRIPT);
    writeFileSync(src, readFileSync(h));
    const d = await readDocument(src);
    assert.equal(d.engine, 'rhwp');
    assert.ok(d.text.includes('분기 보고'));
    assert.deepEqual((await listFields(src)).map((f) => f.name), ['담당자']);
    assert.ok(isPng((await pageImage(src, 1, { dpi: 48 })).png));
    const r = await fillFields(src, { 담당자: '박민수' }, { output: tmp.path('form-out.hwt'), engine: 'file' });
    assert.deepEqual(r.filled, ['담당자']);
    assert.ok((await textOf(tmp.path('form-out.hwt'))).includes('분기 보고'));
    const p = await (await import('../src/core/index.mjs')).renderPreview(src, tmp.path('pv'));
    assert.equal(p.kind, 'svg');
  });
});

describe('변환: 지원하지 않는 형식은 프로그램을 띄우기 전에 거부', () => {
  test('docx → hwpx / exe', async () => {
    const src = tmp.path('conv.docx');
    await createDocument(src, MANUSCRIPT);
    const t = Date.now();
    await assert.rejects(convertDocument(src, tmp.path('conv.hwpx')), /Word 문서는 \.hwpx 형식으로 바꿀 수 없습니다 \(가능: .*pdf/);
    await assert.rejects(convertDocument(src, tmp.path('conv.exe')), /바꿀 수 없습니다/);
    await assert.rejects(convertDocument(tmp.path('conv-src.hwp'), tmp.path('x.xlsx')), /파일이 없습니다/);
    assert.ok(Date.now() - t < 5000, `${Date.now() - t}ms`);
  });
});

// PowerShell을 못 쓰는 PC(실행 정책 등): 프로그램 확인이 실패해도 파일 엔진 작업은 돼야 한다
describe('PowerShell 없이 (OFFICEKIT_POWERSHELL=없는 경로)', () => {
  const env = { OFFICEKIT_POWERSHELL: 'C:\\없는폴더\\powershell.exe' };
  test('읽기·바꾸기·미리보기는 파일 엔진으로', async () => {
    await createDocument(tmp.path('nops.docx'), MANUSCRIPT);
    await createDocument(tmp.path('nops.hwp'), MANUSCRIPT);
    let r = await cli(['text', 'nops.hwp'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /분기 보고/);
    r = await cli(['replace', 'nops.docx', '매출', '수익'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /2곳 바꿈 \(file\)/);
    r = await cli(['convert', 'nops.docx', 'nops.pdf'], { env });
    assert.equal(r.code, 1);
    assert.match(r.stderr, /^오류: .*확인하지 못해/m);
    const code = `import('${pathToFileURL(join(ROOT, 'src/core/index.mjs')).href}').then(async (m) => {
      const p = await m.renderPreview('nops.docx', 'pv');
      console.log(JSON.stringify({ kind: p.kind, docKind: p.docKind, has: p.text.includes('수익') }));
    })`;
    const out = await new Promise((done) => execFile(process.execPath, ['-e', code], { cwd: tmp.dir, env: { ...process.env, ...env }, encoding: 'utf8' },
      (err, stdout, stderr) => done({ err, stdout, stderr })));
    assert.ok(!out.err, out.stderr);
    assert.deepEqual(JSON.parse(out.stdout.trim().split('\n').pop()), { kind: 'text', docKind: 'Word', has: true });
  });
});

// ── 명령줄 ───────────────────────────────────────────────

const CLI = join(ROOT, 'src', 'cli.mjs');
function cli(args, { exe = process.execPath, env = {} } = {}) {
  return new Promise((done) => {
    execFile(exe, [CLI, ...args], { cwd: tmp.dir, env: { ...process.env, ...env }, encoding: 'utf8', timeout: 120_000, windowsHide: true },
      (err, stdout, stderr) => done({ code: err ? err.code ?? 1 : 0, stdout, stderr }));
  });
}

describe('명령줄 office-kit', () => {
  before(() => writeFileSync(tmp.path('원고.md'), MANUSCRIPT, 'utf8'));

  test('help', async () => {
    const r = await cli(['help']);
    assert.equal(r.code, 0);
    assert.match(r.stdout, /사용법: office-kit/);
  });

  test('new → fields → fill → replace → text → convert → image', async () => {
    let r = await cli(['new', '원고.md', 'cli.hwp']);
    assert.equal(r.code, 0, r.stderr);
    r = await cli(['new', '원고.md', 'cli.hwp']);
    assert.equal(r.code, 1);
    assert.match(r.stderr, /이미 있는 파일/);
    r = await cli(['fields', 'cli.hwp']);
    assert.match(r.stdout, /담당자\t\[누름틀\]/);
    writeFileSync(tmp.path('값.json'), '﻿{"담당자":"김영희"}', 'utf8');
    r = await cli(['fill', 'cli.hwp', '값.json', '-o', 'out/cli-fill.hwpx']);
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /채움: 담당자/);
    r = await cli(['replace', 'out/cli-fill.hwpx', '매출', '수익']);
    assert.match(r.stdout, /2곳 바꿈/);
    r = await cli(['text', 'out/cli-fill.hwpx']);
    assert.ok(r.stdout.includes('김영희') && r.stdout.includes('수익 증가'), r.stdout);
    r = await cli(['text', 'out/cli-fill.hwpx', '--out', 'out/t.txt']);
    assert.ok(readFileSync(tmp.path('out/t.txt'), 'utf8').includes('김영희'));
    r = await cli(['convert', 'out/cli-fill.hwpx', 'out/cli.hwp']);
    assert.equal(r.code, 0, r.stderr);
    r = await cli(['image', 'out/cli.hwp', '2', '-o', 'out/p2.png', '--dpi', '72']);
    assert.equal(r.code, 0, r.stderr);
    assert.ok(isPng(readFileSync(tmp.path('out/p2.png'))));
  });

  test('cells (JSON 문자열 인수)', async () => {
    await createDocument(tmp.path('cli.xlsx'), MANUSCRIPT);
    const r = await cli(['cells', 'cli.xlsx', '{"B1":"합계","B2":"=1+2"}', '--sheet', 'Sheet1', '--engine', 'file']);
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /셀 2개/);
  });

  test('오류는 한국어 + 종료 코드 1', async () => {
    for (const args of [['없는명령'], ['text', '없음.hwp'], ['image', 'cli.hwp', '9'], ['fill', 'cli.hwp'], ['text', 'cli.hwp', '--engine', 'x'], ['setup', '어딘가']]) {
      const r = await cli(args);
      assert.equal(r.code, 1, `${args.join(' ')} → ${r.code}`);
      assert.match(r.stderr, /^오류: .*[가-힣]/, r.stderr);
    }
  });

  // 실제 설정(Claude Desktop, Codex, 명령줄 shim)은 건드리지 않도록 경로 환경 변수를 임시 폴더로 바꾼다.
  // Claude Code는 claude 명령이 실제 ~/.claude.json을 고치므로 여기서 실행하지 않는다.
  test('setup claude-desktop / codex / cli (임시 폴더)', { skip: process.platform !== 'win32' && 'Windows 전용' }, async (t) => {
    let integrations;
    try { integrations = await import('../src/setup/integrations.mjs'); } catch { return t.skip('src/setup 없음'); }
    const sys = process.env.SystemRoot ?? 'C:\\Windows';
    const env = {
      APPDATA: tmp.path('setup/appdata'), LOCALAPPDATA: tmp.path('setup/local'), CODEX_HOME: tmp.path('setup/codex'),
      PATH: `${sys}\\System32;${sys}\\System32\\WindowsPowerShell\\v1.0`,
    };
    const desktop = join(env.APPDATA, 'Claude', 'claude_desktop_config.json');
    let r = await cli(['setup', 'claude-desktop'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /Claude Desktop: 연결했습니다/);
    const cfg = JSON.parse(readFileSync(desktop, 'utf8'));
    assert.deepEqual(cfg.mcpServers['office-kit'].args, [join(ROOT, 'src', 'mcp', 'server.mjs')]);
    r = await cli(['setup', 'claude-desktop'], { env });
    assert.match(r.stdout, /이미 연결돼 있습니다/);
    r = await cli(['setup', 'codex'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(readFileSync(join(env.CODEX_HOME, 'config.toml'), 'utf8'), /\[mcp_servers\.office-kit\]/);
    r = await cli(['setup', 'cli'], { env });
    assert.equal(r.code, 0, r.stderr);
    const shim = join(env.LOCALAPPDATA, 'Microsoft', 'WindowsApps', 'office-kit.cmd');
    assert.ok(readFileSync(shim, 'utf8').includes('cli.mjs'));
    r = await cli(['status'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /연결: Claude Desktop 연결됨, Claude Code [^,]+, Codex 연결됨, 명령줄 연결됨/);
    for (const target of ['claude-desktop', 'codex', 'cli']) {
      r = await cli(['setup', target, '--remove'], { env });
      assert.equal(r.code, 0, r.stderr);
      assert.match(r.stdout, /연결을 해제했습니다/);
    }
    assert.ok(!JSON.parse(readFileSync(desktop, 'utf8')).mcpServers?.['office-kit']);
    assert.ok(!existsSync(shim));
    assert.equal(typeof integrations.mcpLaunchSpec, 'function');
  });

  test('status: 다른 경로로 등록된 연결·읽을 수 없는 설정 (임시 폴더)', { skip: process.platform !== 'win32' && 'Windows 전용' }, async (t) => {
    try { await import('../src/setup/integrations.mjs'); } catch { return t.skip('src/setup 없음'); }
    const sys = process.env.SystemRoot ?? 'C:\\Windows';
    const env = {
      APPDATA: tmp.path('stale/appdata'), LOCALAPPDATA: tmp.path('stale/local'), CODEX_HOME: tmp.path('stale/codex'),
      PATH: `${sys}\\System32;${sys}\\System32\\WindowsPowerShell\\v1.0`,
    };
    const desktop = join(env.APPDATA, 'Claude', 'claude_desktop_config.json');
    const { mkdirSync } = await import('node:fs');
    mkdirSync(join(env.APPDATA, 'Claude'), { recursive: true });
    writeFileSync(desktop, JSON.stringify({ mcpServers: { 'office-kit': { command: 'C:\\Old\\Office Kit.exe', args: ['x'] } } }));
    let r = await cli(['status'], { env });
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /Claude Desktop 연결됨\(다른 경로 — 다시 연결 필요\)/);
    writeFileSync(desktop, '{ 이건 JSON 아님');
    r = await cli(['status'], { env });
    assert.match(r.stdout, /Claude Desktop 확인 필요 \(.*[가-힣]/);
  });

  // Electron 실행 파일로도 동작해야 한다 (설치본의 명령줄 연결)
  let electron;
  try { electron = createRequire(import.meta.url)('electron'); } catch {}
  test('Electron + ELECTRON_RUN_AS_NODE=1', { skip: !(typeof electron === 'string' && existsSync(electron)) && 'electron 없음' }, async () => {
    const r = await cli(['text', 'cli.hwp'], { exe: electron, env: { ELECTRON_RUN_AS_NODE: '1' } });
    assert.equal(r.code, 0, r.stderr);
    assert.ok(r.stdout.includes('분기 보고'), r.stdout);
    const bad = await cli(['text', '없음.hwp'], { exe: electron, env: { ELECTRON_RUN_AS_NODE: '1' } });
    assert.equal(bad.code, 1);
    assert.match(bad.stderr, /파일이 없습니다/);
  });
});

// ── 설치된 프로그램 (COM) — OFFICEKIT_COM_TESTS=1 일 때만 ──────

describe('프로그램 엔진 (Word/PowerPoint/Excel)', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  for (const ext of ['docx', 'pptx', 'xlsx']) {
    test(`.${ext} 읽기·바꾸기·PDF·쪽 그림`, { timeout: 300_000 }, async () => {
      const src = tmp.path(`com.${ext}`);
      await createDocument(src, MANUSCRIPT);
      const r = await readDocument(src, { engine: 'app' });
      assert.notEqual(r.engine, 'file');
      assert.ok(r.text.includes('매출'), r.text);
      const out = tmp.path(`com-rep.${ext}`);
      const rep = await replaceText(src, [['매출', '수익']], { output: out, engine: 'app' });
      assert.ok(rep.count >= 2, `count=${rep.count}`);
      const after = await textOf(out);
      assert.ok(after.includes('수익') && !after.includes('매출'), after);
      const pdf = tmp.path(`com-${ext}.pdf`);
      await convertDocument(src, pdf);
      assert.equal(readFileSync(pdf).subarray(0, 4).toString(), '%PDF');
      const img = await pageImage(src, 1);
      assert.ok(isPng(img.png));
      // PowerPoint는 PDF를 거치지 않고 슬라이드를 바로 그린다 (숨긴 슬라이드 때문에)
      assert.match(img.engine, ext === 'pptx' ? /^powerpoint$/ : /\+pdf$/);
    });
  }

  test('.pptx 쪽 그림 = 슬라이드 번호 (숨긴 슬라이드 포함, PDF 안 거침)', { timeout: 300_000 }, async () => {
    const running = () => new Promise((done) => execFile('tasklist', ['/FI', 'IMAGENAME eq POWERPNT.EXE', '/NH'], { encoding: 'utf8' },
      (_e, out) => done(String(out).includes('POWERPNT.EXE'))));
    const before = await running();
    const src = await slideDeck(tmp.path('hidden.pptx'));
    // 숨긴 슬라이드 2도 2쪽, 3쪽은 슬라이드 3 (PDF였다면 2쪽짜리라 슬라이드 3이 2쪽)
    const p2 = await pageImage(src, 2, { dpi: 24 });
    assert.equal(p2.engine, 'powerpoint');
    assert.deepEqual([p2.page, p2.pages], [2, 3]);
    assert.deepEqual(pngSize(p2.png), { width: 320, height: 180 }); // 960×540pt @24dpi (기본 크기에서 줄임)
    assert.deepEqual([p2.width, p2.height], [320, 180]);
    assert.equal(pixelAt(p2.png, 4, 4), '00ff00');
    const p3 = await pageImage(src, 3, { dpi: 24 });
    assert.equal(pixelAt(p3.png, 4, 4), '0000ff');
    const p1 = await pageImage(src, 1, { dpi: 96 });
    assert.deepEqual(pngSize(p1.png), { width: 1280, height: 720 });
    assert.equal(pixelAt(p1.png, 4, 4), 'ff0000');
    // dpi가 기본 크기보다 크면 늘리지 않는다 (같은 슬라이드는 다시 띄우지 않고 기억한 그림을 씀)
    const t = Date.now();
    const big = await pageImage(src, 1, { dpi: 200 });
    assert.deepEqual(pngSize(big.png), { width: 1280, height: 720 });
    assert.ok(Date.now() - t < 2000, `${Date.now() - t}ms`);
    await assert.rejects(pageImage(src, 4), /범위를 벗어났습니다: 4쪽 \(전체 3쪽/);
    // .ppt는 미리 셀 수 없어 PowerPoint가 알려 준 수로 같은 문구
    const ppt = tmp.path('hidden.ppt');
    await convertDocument(src, ppt);
    const q2 = await pageImage(ppt, 2, { dpi: 24 });
    assert.deepEqual([q2.page, q2.pages], [2, 3]);
    assert.equal(pixelAt(q2.png, 4, 4), '00ff00');
    await assert.rejects(pageImage(ppt, 5), /범위를 벗어났습니다: 5쪽 \(전체 3쪽/);
    if (!before) {
      await new Promise((r) => setTimeout(r, 1000));
      assert.equal(await running(), false, '숨은 POWERPNT.EXE가 남음');
    }
  });

  // 프로그램은 임시 폴더에 PDF를 쓰고 미리보기 폴더로 옮긴다 (MSIX 앱 안에서는 가상화된 %APPDATA%를 Word가 못 본다)
  test('.docx 미리보기 PDF는 미리보기 폴더에 생긴다', { timeout: 180_000 }, async () => {
    const src = tmp.path('com-preview.docx');
    await createDocument(src, MANUSCRIPT);
    const outDir = tmp.path('com-pv');
    const p = await (await import('../src/core/index.mjs')).renderPreview(src, outDir);
    assert.equal(p.kind, 'pdf');
    assert.equal(dirname(p.file), outDir);
    assert.equal(readFileSync(p.file).subarray(0, 4).toString(), '%PDF');
  });

  test('.docx 칸 채우기 (Word)', { timeout: 120_000 }, async () => {
    const src = tmp.path('com-fill.docx');
    await createDocument(src, MANUSCRIPT);
    const r = await fillFields(src, { 담당자: '박민수', 없는칸: 'x' }, { engine: 'app' });
    assert.deepEqual(r.filled, ['담당자']);
    assert.deepEqual(r.missing, ['없는칸']);
    assert.ok((await textOf(src)).includes('박민수'));
  });

  // 파일 엔진과 같은 규칙: 대소문자 구분, 특수 문자(* ? ~ ^)는 글자 그대로, 쌍마다 개수
  for (const ext of ['docx', 'pptx', 'xlsx']) {
    test(`.${ext} 바꾸기 대소문자·특수 문자 (프로그램)`, { timeout: 180_000 }, async () => {
      const src = tmp.path(`com-case.${ext}`);
      await createDocument(src, `${ENGLISH}가격 a*b? ~x^p\n`);
      const out = tmp.path(`com-case-out.${ext}`);
      const r = await replaceText(src, [['AI', '인공지능'], ['a*b?', '별~표^&'], ['^p', '[캐럿]']], { output: out, engine: 'app' });
      assert.deepEqual(r.counts, [1, 1, 1], JSON.stringify(r));
      const t = await textOf(out);
      assert.ok(t.includes('The main goal was said by email.'), t);
      assert.ok(t.includes('가격 별~표^& ~x[캐럿]'), t);
    });
  }

  test('파일 엔진으로 쓴 xlsx를 Excel이 그대로 읽음', { timeout: 120_000 }, async () => {
    const src = tmp.path('com-chart.xlsx');
    writeFileSync(src, await xlsxWithChart());
    await setCells(src, { B3: 25, C2: '=B2*2', A9: '새 행' }, { engine: 'file' });
    const rows = (await readDocument(src, { engine: 'app' })).sheets[0].rows;
    assert.equal(rows[2][1], '25');
    assert.equal(rows[1][2], '20');
    assert.equal(rows[8][0], '새 행');
  });

  test('Excel 읽기: 한 번에 읽고, 잘리면 알림', { timeout: 180_000 }, async () => {
    const big = tmp.path('big.csv');
    // BOM이 없으면 Excel은 csv를 CP949로 읽는다
    writeFileSync(big, `﻿${Array.from({ length: 10_050 }, (_, i) => `${i + 1},행${i + 1}`).join('\n')}`);
    const t = Date.now();
    const r = await readDocument(big, { engine: 'app' });
    assert.ok(Date.now() - t < 60_000, `${Date.now() - t}ms`);
    assert.equal(r.truncated, true);
    assert.equal(r.sheets[0].totalRows, 10_050);
    assert.equal(r.sheets[0].rows.length, 10_000);
    assert.deepEqual(r.sheets[0].rows[9_999], ['10000', '행10000']);
    assert.match(r.text, /잘림: 전체 10050행×2열 중 10000행×2열만 읽음/);
  });

  test('시간 초과면 숨은 Office도 끝낸다', { timeout: 180_000 }, async () => {
    const { comCall } = await import('../src/core/com.mjs');
    const huge = tmp.path('huge.csv');
    writeFileSync(huge, Array.from({ length: 12_000 }, (_, i) => Array.from({ length: 20 }, (_, c) => (c ? i * c : `r${i}`)).join(',')).join('\n'));
    const running = () => new Promise((done) => execFile('tasklist', ['/FI', 'IMAGENAME eq EXCEL.EXE', '/NH'], { encoding: 'utf8' },
      (_e, out) => done(String(out).includes('EXCEL.EXE'))));
    const before = await running();
    await assert.rejects(comCall({ app: 'excel', op: 'text', path: huge }, { timeoutMs: 6000 }), /응답하지 않아 작업을 멈췄습니다/);
    await new Promise((r) => setTimeout(r, 1500));
    if (!before) assert.equal(await running(), false, '숨은 EXCEL.EXE가 남음');
    // 버려진 잠금(뮤텍스) 뒤에도 다음 요청이 된다
    const r = await readDocument(tmp.path('com.xlsx'), { engine: 'app' });
    assert.ok(r.text.includes('매출'));
  });
});
