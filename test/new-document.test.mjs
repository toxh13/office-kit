// new-document.test.mjs — 새 문서: 빈 문서 만들기(다섯 형식)와 새 파일 이름 도우미.
// 기본은 파일 엔진만 쓴다. Word/PowerPoint/Excel로 여는 시험은 OFFICEKIT_COM_TESTS=1 일 때만 (한컴 COM은 쓰지 않음).
import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import {
  BLANK_FORMATS, DEFAULT_DOC_NAME, createBlankDocument, cleanDocName, fileNameProblem, freeDocName, planBlankDocument,
} from '../src/core/create.mjs';
import { readDocument, renderPreview, replaceText, convertDocument } from '../src/core/index.mjs';
import { Hwp } from '../src/core/hwp.mjs';
import { pageImage } from '../src/core/render.mjs';
import { COM_TESTS, tempDir, isPng } from './helpers/fixtures.mjs';

const tmp = tempDir('officekit-newdoc-');
after(() => tmp.cleanup());

const FILE = { engine: 'file' };
const zipText = async (path, name) => (await JSZip.loadAsync(readFileSync(path))).file(name)?.async('string');

// ── 빈 문서 (파일 엔진) ─────────────────────────────────────

describe('빈 문서 만들기', () => {
  test('형식 목록: 한글(hwpx)이 기본', () => {
    assert.deepEqual([...BLANK_FORMATS], ['hwpx', 'hwp', 'docx', 'pptx', 'xlsx']);
  });

  for (const ext of ['hwpx', 'hwp']) {
    test(`.${ext}: 글 없는 한 쪽, 제목을 주면 첫 문단`, async () => {
      const blank = await createBlankDocument(tmp.path(`빈.${ext}`));
      const h = Hwp.open(blank);
      try {
        assert.equal(h.pageCount, 1);
        assert.equal(h.text().trim(), '');
        assert.equal(h.tables().length, 0);
      } finally { h.close(); }
      const titled = await createBlankDocument(tmp.path(`제목.${ext}`), { title: '주간 회의록' });
      const t = await readDocument(titled, FILE);
      assert.equal(t.text.trim(), '주간 회의록');
      // 미리보기(rhwp SVG)도 그려진다
      const prev = await renderPreview(blank, tmp.path('preview'));
      assert.equal(prev.kind, 'svg');
      assert.equal(prev.total, 1);
      assert.ok(isPng((await pageImage(blank, 1)).png));
    });
  }

  test('.docx: 빈 문단 하나, 지은이 없음, 맑은 고딕', async () => {
    const p = await createBlankDocument(tmp.path('빈.docx'));
    assert.equal((await readDocument(p, FILE)).text.trim(), '');
    const body = await zipText(p, 'word/document.xml');
    assert.equal((body.match(/<w:p\b/g) ?? []).length, 1);
    const core = await zipText(p, 'docProps/core.xml');
    assert.ok(!/Un-named|<dc:creator>/.test(core), core);
    assert.ok((await zipText(p, 'word/styles.xml')).includes('맑은 고딕'));
    const titled = await createBlankDocument(tmp.path('제목.docx'), { title: '보고서 & 계획' });
    assert.equal((await readDocument(titled, FILE)).text.trim(), '보고서 & 계획');
    assert.ok((await zipText(titled, 'docProps/core.xml')).includes('<dc:title>보고서 &amp; 계획</dc:title>'));
  });

  test('.pptx: 제목 슬라이드 하나 (빈 가운데 제목·부제목 개체 틀)', async () => {
    const p = await createBlankDocument(tmp.path('빈.pptx'));
    const r = await readDocument(p, FILE);
    assert.deepEqual(r.slides, ['']);
    const zip = await JSZip.loadAsync(readFileSync(p));
    assert.deepEqual(Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)), ['ppt/slides/slide1.xml']);
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    assert.match(slide, /<p:ph\b[^>]*type="ctrTitle"/);
    assert.match(slide, /<p:ph\b[^>]*type="subTitle"/);
    assert.ok(slide.includes('name="제목 1"') && slide.includes('name="부제목 2"'), '개체 이름');
    // 레이아웃: 제목 슬라이드(type=title) — PowerPoint·ppt_inspect가 title로 알아본다
    const layouts = await Promise.all(Object.keys(zip.files).filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).map((n) => zip.file(n).async('string')));
    assert.ok(layouts.some((x) => /<p:sldLayout type="title"/.test(x) && x.includes('제목 슬라이드') && x.includes('제목을 입력하세요')));
    // 레이아웃 3개: PowerPoint '새 슬라이드'가 제목 슬라이드 다음에 고르는 '제목 및 내용'(본문 = 내용 개체 틀)까지
    assert.equal(layouts.length, 3);
    assert.ok(layouts.some((x) => /<p:sldLayout type="obj"/.test(x) && x.includes('제목 및 내용') && /<p:ph\b[^>]*type="obj"/.test(x)));
    assert.ok(layouts.some((x) => /<p:sldLayout type="blank"/.test(x) && x.includes('빈 화면')));
    // 갤러리 순서: 제목 슬라이드 → 제목 및 내용 → 빈 화면
    const master = await zip.file('ppt/slideMasters/slideMaster1.xml').async('string');
    const rels = await zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels').async('string');
    const target = Object.fromEntries([...rels.matchAll(/Id="(rId\d+)"[^>]*Target="\.\.\/slideLayouts\/(slideLayout\d+\.xml)"/g)].map((m) => [m[1], m[2]]));
    const names = await Promise.all([...master.matchAll(/<p:sldLayoutId\b[^>]*r:id="(rId\d+)"/g)]
      .map(async (m) => /<p:cSld name="([^"]+)"/.exec(await zip.file(`ppt/slideLayouts/${target[m[1]]}`).async('string'))[1]));
    assert.deepEqual(names, ['제목 슬라이드', '제목 및 내용', '빈 화면']);
    const core = await zip.file('docProps/core.xml').async('string');
    assert.ok(!core.includes('PptxGenJS'), core);
    const titled = await createBlankDocument(tmp.path('제목.pptx'), { title: '회사 소개' });
    assert.deepEqual((await readDocument(titled, FILE)).slides, ['회사 소개']);
  });

  test('.xlsx: 빈 시트 하나 (Sheet1), 지은이 없음', async () => {
    const p = await createBlankDocument(tmp.path('빈.xlsx'));
    const r = await readDocument(p, FILE);
    assert.equal(r.sheets.length, 1);
    assert.equal(r.sheets[0].name, 'Sheet1');
    assert.equal(r.sheets[0].rows.length, 0);
    const core = await zipText(p, 'docProps/core.xml');
    assert.ok(!/Unknown|ExcelJS/.test(core), core);
    const titled = await createBlankDocument(tmp.path('제목.xlsx'), { title: '가계부' });
    assert.deepEqual((await readDocument(titled, FILE)).sheets[0].rows, [['가계부']]);
  });

  test('이미 있는 파일은 덮어쓰지 않는다 (overwrite=true일 때만)', async () => {
    const p = tmp.path('있음.docx');
    writeFileSync(p, 'keep');
    await assert.rejects(createBlankDocument(p), (e) => e.code === 'EXISTS' && /이미 있습니다/.test(e.message));
    assert.equal(readFileSync(p, 'utf8'), 'keep');
    await createBlankDocument(p, { overwrite: true });
    assert.notEqual(readFileSync(p, 'utf8'), 'keep');
    assert.equal((await readDocument(p, FILE)).text.trim(), '');
  });

  test('지원하지 않는 형식·없는 폴더는 오류, 파일을 남기지 않는다', async () => {
    await assert.rejects(createBlankDocument(tmp.path('a.txt')), /빈 문서를 만들 수 없는 형식입니다: \.txt/);
    await assert.rejects(createBlankDocument(tmp.path('a.doc')), /\.doc/);
    assert.ok(!existsSync(tmp.path('a.txt')));
    await assert.rejects(createBlankDocument(tmp.path('없는 폴더/a.docx')), (e) => e.code === 'ENOENT');
  });
});

// ── 새 파일 이름 도우미 ─────────────────────────────────────

describe('새 파일 이름', () => {
  test('fileNameProblem: 쓸 수 있는 이름', () => {
    for (const n of ['새 문서', '보고서 (2)', '2026.3분기 보고', '회의록_최종-v2', 'a', '한글.문서', '…', 'con-report', 'aux1', 'console', 'x'.repeat(200)]) {
      assert.equal(fileNameProblem(n), null, n);
    }
  });

  test('fileNameProblem: 빈 이름', () => {
    for (const n of ['', '   ', '.', '..', ' . ', null, undefined]) assert.equal(fileNameProblem(n)?.code, 'EMPTY', String(n));
  });

  test('fileNameProblem: Windows에서 못 쓰는 문자는 무엇인지 알려 준다', () => {
    for (const c of ['<', '>', ':', '"', '/', '\\', '|', '?', '*']) {
      const p = fileNameProblem(`보고${c}서`);
      assert.equal(p?.code, 'BAD_CHARS', c);
      assert.ok(p.message.endsWith(c), p.message);
    }
    assert.equal(fileNameProblem('a?b*c?').message, '파일 이름에 쓸 수 없는 문자가 있어요: ? *');
    assert.match(fileNameProblem('탭\t이름').message, /제어 문자/);
    assert.equal(fileNameProblem('..\\..\\escape').code, 'BAD_CHARS');
  });

  test('fileNameProblem: 예약 이름 (대소문자·확장자 무관)', () => {
    for (const n of ['CON', 'con', 'Prn', 'AUX', 'nul', 'COM1', 'com9', 'LPT1', 'lpt9', 'nul.txt', 'CON.보고서', 'COM¹', 'CONIN$']) {
      const p = fileNameProblem(n);
      assert.equal(p?.code, 'RESERVED', n);
      assert.match(p.message, /Windows가 예약한 이름\(.+\)이라 쓸 수 없어요/);
    }
  });

  test('fileNameProblem: 앞뒤 공백, 너무 긴 이름', () => {
    assert.equal(fileNameProblem(' 앞 공백').code, 'SPACES');
    assert.equal(fileNameProblem('뒤 공백 ').code, 'SPACES');
    assert.equal(fileNameProblem('가'.repeat(201)).code, 'TOO_LONG');
  });

  test('cleanDocName: 앞뒤 공백과 같은 확장자를 뗀다', () => {
    assert.equal(cleanDocName('  보고서  ', 'hwpx'), '보고서');
    assert.equal(cleanDocName('보고서.hwpx', 'hwpx'), '보고서');
    assert.equal(cleanDocName('보고서.HWPX', 'hwpx'), '보고서');
    assert.equal(cleanDocName('보고서.docx', 'hwpx'), '보고서.docx');
    assert.equal(cleanDocName('보고서 .hwpx ', 'hwpx'), '보고서');
    assert.equal(cleanDocName(undefined, 'hwpx'), '');
  });

  test('freeDocName: 겹치지 않는 이름', () => {
    const taken = (names) => (p) => names.some((n) => p.toLowerCase().endsWith(`\\${n}`.toLowerCase()) || p.toLowerCase().endsWith(`/${n}`.toLowerCase()));
    assert.equal(DEFAULT_DOC_NAME, '새 문서');
    assert.equal(freeDocName('C:\\d', 'hwpx', undefined, taken([])), '새 문서');
    assert.equal(freeDocName('C:\\d', 'hwpx', undefined, taken(['새 문서.hwpx'])), '새 문서 (2)');
    assert.equal(freeDocName('C:\\d', 'hwpx', undefined, taken(['새 문서.hwpx', '새 문서 (2).hwpx'])), '새 문서 (3)');
    // 형식이 다르면 같은 이름이어도 된다
    assert.equal(freeDocName('C:\\d', 'docx', undefined, taken(['새 문서.hwpx'])), '새 문서');
    // 이미 번호가 붙은 이름은 그다음 번호부터
    assert.equal(freeDocName('C:\\d', 'docx', '보고서 (2)', taken(['보고서 (2).docx'])), '보고서 (3)');
    assert.equal(freeDocName('C:\\d', 'docx', '보고서', taken(['보고서.docx', '보고서 (2).docx', '보고서 (3).docx'])), '보고서 (4)');
  });

  test('freeDocName: 실제 폴더 (대소문자 무시)', () => {
    const dir = tmp.path('이름들');
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, '새 문서.pptx'), '');
    writeFileSync(join(dir, '새 문서 (2).PPTX'), '');
    assert.equal(freeDocName(dir, 'pptx'), '새 문서 (3)');
    assert.equal(freeDocName(dir, 'xlsx'), '새 문서');
  });

  test('planBlankDocument: 통과하면 경로·이름을 돌려준다', () => {
    const dir = tmp.path('계획');
    mkdirSync(dir, { recursive: true });
    const plan = planBlankDocument({ folder: dir, name: ' 주간 회의록.hwpx ', format: 'hwpx' });
    assert.deepEqual(plan, { path: join(dir, '주간 회의록.hwpx'), name: '주간 회의록', ext: 'hwpx', folder: dir });
    assert.equal(planBlankDocument({ folder: `${dir}\\`, name: 'a', format: 'docx' }).path, join(dir, 'a.docx'));
  });

  test('planBlankDocument: 오류마다 code와 한국어 문장', () => {
    const dir = tmp.path('계획2');
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, '있는 문서.xlsx'), '');
    const code = (args, opts) => { try { planBlankDocument(args, opts); return null; } catch (e) { assert.match(e.message, /[가-힣]/); return e.code; } };
    assert.equal(code({ folder: dir, name: 'a', format: 'txt' }), 'BAD_FORMAT');
    assert.equal(code({ folder: dir, name: 'a', format: 'HWPX' }), null); // 대문자 형식도 받아 준다
    assert.equal(code({ folder: dir, name: 'a' }), 'BAD_FORMAT');
    assert.equal(code({ folder: 'relative\\dir', name: 'a', format: 'docx' }), 'NO_FOLDER');
    assert.equal(code({ folder: '', name: 'a', format: 'docx' }), 'NO_FOLDER');
    assert.equal(code({ folder: 42, name: 'a', format: 'docx' }), 'NO_FOLDER');
    assert.equal(code({ folder: join(dir, '없음'), name: 'a', format: 'docx' }), 'NO_FOLDER');
    assert.equal(code({ folder: join(dir, '있는 문서.xlsx'), name: 'a', format: 'docx' }), 'NO_FOLDER'); // 파일은 폴더가 아님
    assert.equal(code({ folder: dir, name: '', format: 'docx' }), 'EMPTY');
    assert.equal(code({ folder: dir, name: 'a/b', format: 'docx' }), 'BAD_CHARS');
    assert.equal(code({ folder: dir, name: 'LPT1', format: 'docx' }), 'RESERVED');
    assert.equal(code({ folder: dir, name: 'x'.repeat(201), format: 'docx' }), 'TOO_LONG');
    assert.equal(code({ folder: dir, name: 'x'.repeat(200), format: 'docx' }, { exists: () => false }), dir.length + 206 > 259 ? 'PATH_TOO_LONG' : null);
    assert.equal(code({ folder: 'C:\\', name: 'x'.repeat(200), format: 'docx' }, { exists: () => false }), null);
    assert.equal(code({ folder: dir, name: '있는 문서', format: 'xlsx' }), 'EXISTS');
    assert.equal(code({ folder: dir, name: '있는 문서.XLSX', format: 'xlsx' }), 'EXISTS');
    assert.equal(code({ folder: dir, name: '있는 문서', format: 'docx' }), null);
  });

  test('planBlankDocument → createBlankDocument: 계획한 자리에 만든다', async () => {
    const dir = tmp.path('계획3');
    mkdirSync(dir, { recursive: true });
    for (const format of BLANK_FORMATS) {
      const name = freeDocName(dir, format);
      const plan = planBlankDocument({ folder: dir, name, format });
      await createBlankDocument(plan.path);
      assert.ok(existsSync(plan.path), plan.path);
      assert.equal(freeDocName(dir, format), '새 문서 (2)');
      assert.throws(() => planBlankDocument({ folder: dir, name, format }), (e) => e.code === 'EXISTS');
    }
  });
});

// ── 설치된 프로그램으로 열기 (COM) — OFFICEKIT_COM_TESTS=1 일 때만 ──

describe('빈 문서를 Word/PowerPoint/Excel이 연다', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  test('.docx: Word로 읽기·제자리 저장·PDF', { timeout: 300_000 }, async () => {
    const { wordInspect } = await import('../src/core/word.mjs');
    const p = await createBlankDocument(tmp.path('com 빈.docx'));
    const info = await wordInspect(p);
    assert.equal(info.paragraphTotal, 1);
    assert.equal(info.pages, 1);
    assert.equal(info.characters, 0);
    const titled = await createBlankDocument(tmp.path('com 제목.docx'), { title: '분기 보고' });
    const r = await readDocument(titled, { engine: 'app' });
    assert.equal(r.engine, 'word');
    assert.ok(r.text.includes('분기 보고'), r.text);
    // Word가 고쳐 제자리에 저장할 수 있다
    const rep = await replaceText(titled, [['분기', '연간']], { engine: 'app' });
    assert.equal(rep.count, 1);
    assert.equal((await readDocument(titled, FILE)).text.trim(), '연간 보고');
    const pdf = tmp.path('com 빈 docx.pdf');
    await convertDocument(p, pdf);
    assert.equal(readFileSync(pdf).subarray(0, 4).toString(), '%PDF');
  });

  test('.pptx: PowerPoint가 제목 슬라이드로 알아본다', { timeout: 300_000 }, async () => {
    const { pptInspect } = await import('../src/core/ppt.mjs');
    const p = await createBlankDocument(tmp.path('com 빈.pptx'));
    const info = await pptInspect(p);
    assert.equal(info.slideCount, 1);
    assert.equal(info.slideWidth, 960); // 16:9
    const s = info.slides[0];
    assert.equal(s.key, 'title');
    assert.equal(s.layout, '제목 슬라이드');
    assert.deepEqual(s.shapes.map((x) => [x.name, x.placeholder]), [['제목 1', 3], ['부제목 2', 4]]);
    assert.ok(info.layouts.some((l) => l.key === 'blank'));
    assert.deepEqual(info.layouts.map((l) => l.name), ['제목 슬라이드', '제목 및 내용', '빈 화면']);
    const titled = await createBlankDocument(tmp.path('com 제목.pptx'), { title: '분기 보고' });
    assert.equal((await pptInspect(titled)).slides[0].title, '분기 보고');
    const rep = await replaceText(titled, [['분기', '연간']], { engine: 'app' });
    assert.equal(rep.count, 1);
    assert.deepEqual((await readDocument(titled, FILE)).slides, ['연간 보고']);
    assert.ok(isPng((await pageImage(p, 1)).png));
  });

  test('.xlsx: Excel로 살펴보기·셀 쓰기·PDF', { timeout: 300_000 }, async () => {
    const { excelInspect } = await import('../src/core/excel.mjs');
    const { setCells } = await import('../src/core/index.mjs');
    const p = await createBlankDocument(tmp.path('com 빈.xlsx'));
    const info = await excelInspect(p, { engine: 'app' });
    assert.equal(info.engine, 'excel');
    assert.deepEqual(info.sheets.map((x) => [x.name, x.rows]), [['Sheet1', 0]]);
    // 글자·숫자·참거짓·수식을 섞어 한 번에 (office_set_cells 설명의 예와 같은 모양: {"A1":"제목","B2":120})
    await setCells(p, { A1: '항목', B1: 100, A2: '비율', B2: 3.5, C1: true, C2: '=B1*2' }, { engine: 'app' });
    assert.deepEqual((await readDocument(p, FILE)).sheets[0].rows, [['항목', '100', 'true'], ['비율', '3.5', '200']]);
    const pdf = tmp.path('com 빈 xlsx.pdf');
    await convertDocument(await createBlankDocument(tmp.path('com 제목.xlsx'), { title: '가계부' }), pdf);
    assert.equal(readFileSync(pdf).subarray(0, 4).toString(), '%PDF');
  });
});
