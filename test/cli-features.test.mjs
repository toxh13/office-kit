// cli-features.test.mjs — 명령줄 inspect / edit (형식별 전용 기능, docs/FEATURES.md).
// 한글(rhwp)·Excel 파일 엔진은 늘 돈다. Word/PowerPoint/Excel(COM)은 OFFICEKIT_COM_TESTS=1 일 때만 (한컴 COM은 쓰지 않음).
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { execFile } from 'node:child_process';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { createDocument, Hwp } from '../src/core/index.mjs';
import { MANUSCRIPT, COM_TESTS, ROOT, tempDir, isPng } from './helpers/fixtures.mjs';

const CLI = join(ROOT, 'src', 'cli.mjs');
const tmp = tempDir('officekit-cli-');
after(() => tmp.cleanup());

function cli(args, { timeout = 120_000 } = {}) {
  return new Promise((done) => {
    execFile(process.execPath, [CLI, ...args], { cwd: tmp.dir, encoding: 'utf8', timeout, windowsHide: true, maxBuffer: 32 * 1024 * 1024 },
      (err, stdout, stderr) => done({ code: err ? err.code ?? 1 : 0, stdout, stderr }));
  });
}
const ok = (r) => assert.equal(r.code, 0, `${r.stderr}\n${r.stdout}`);
const fails = (r, re) => { assert.equal(r.code, 1, r.stdout); assert.match(r.stderr, re); };
const json = (name, v) => { writeFileSync(tmp.path(name), JSON.stringify(v), 'utf8'); return name; };
const bytes = (name) => readFileSync(tmp.path(name));
const copy = (from, to) => { copyFileSync(tmp.path(from), tmp.path(to)); return to; };

// ── 도움말 ───────────────────────────────────────────────

describe('도움말', () => {
  test('help: inspect·edit 안내', async () => {
    const r = await cli(['help']);
    ok(r);
    assert.match(r.stdout, /inspect <파일> \[--offset n\] \[--limit n\] \[--slide n\] \[--json\]/);
    assert.match(r.stdout, /edit <파일> <작업\.json> \[-o 출력\] \[--overwrite\]/);
    assert.match(r.stdout, /--overwrite/);
  });

  test('help edit: 네 형식의 작업 목록과 규칙', async () => {
    const r = await cli(['help', 'edit']);
    ok(r);
    for (const s of ['Word (docx', 'PowerPoint (pptx', 'Excel (xlsx', '한글 (hwp', 'insert_toc', 'track_changes', 'set_notes', 'add_chart',
      'conditional_format', 'sort×', 'recalculate×', 'table_add_row', 'page_setup', '한 번만 저장', 'FEATURES.md']) {
      assert.ok(r.stdout.includes(s), `${s} 없음:\n${r.stdout}`);
    }
  });
});

// ── 한글 (rhwp, 프로그램 없이) ───────────────────────────

describe('한글 inspect / edit (rhwp)', () => {
  before(async () => {
    await createDocument(tmp.path('보고.hwpx'), MANUSCRIPT);
    ok(await cli(['image', '보고.hwpx', '1', '-o', '그림.png', '--dpi', '48']));
  });

  test('inspect: 쪽 설정·문단·표·누름틀 번호', async () => {
    const r = await cli(['inspect', '보고.hwpx']);
    ok(r);
    assert.match(r.stdout, /한글\(\.hwpx\) 2쪽/);
    assert.match(r.stdout, /\[쪽 설정\]\n {2}구역 1: A4 세로/);
    assert.match(r.stdout, /1\. 분기 보고/);
    assert.match(r.stdout, /표 1 — 문단 \d+, 3행×2열, 첫 행: 항목 \| 값/);
    assert.match(r.stdout, /담당자 = \(빈 칸\)/);
  });

  test('inspect --json --offset --limit: 문단 나눠 보기', async () => {
    const r = await cli(['inspect', '보고.hwpx', '--json', '--offset', '1', '--limit', '3']);
    ok(r);
    const d = JSON.parse(r.stdout);
    assert.deepEqual(d.paragraphs.map((p) => p.n), [2, 3, 4]);
    assert.equal(d.paragraphs[0].text, '2026년 3분기');
    assert.equal(d.tables[0].rows, 3);
    assert.ok(d.paragraphCount >= 10);
  });

  test('edit: 작업을 순서대로 적용한 뒤 한 번 저장 (-o, 원본은 그대로)', async () => {
    const before = bytes('보고.hwpx');
    const ops = json('한글작업.json', [
      { op: 'insert_paragraph', text: '맺음말을 덧붙입니다', at: { after_paragraph: 5 } },
      { op: 'format_text', find: '매출', all: true, bold: true, color: '#c00000' },
      { op: 'table_set_cell', table: 1, row: 2, col: 2, text: '150' },
      { op: 'table_add_row', table: 1, values: ['합계', '230'] },
      { op: 'set_header', text: '분기 보고 — 내부용' },
      { op: 'set_footer', text: '', page_numbers: true },
      { op: 'fill', values: { 담당자: '김영희' } },
      { op: 'insert_image', image_path: '그림.png', at: 'end', width_mm: 30 },
    ]);
    const r = await cli(['edit', '보고.hwpx', ops, '-o', 'out/보고-수정.hwpx']);
    ok(r);
    assert.match(r.stdout, /^1\. 문단 넣기 → 문단 6: 맺음말을 덧붙입니다/m);
    assert.match(r.stdout, /^4\. 표 1에 행 추가 → 4행/m);
    assert.match(r.stdout, /^8\. 그림 넣기 → 문단 \d+: 그림\.png 30×[\d.]+mm/m);
    assert.match(r.stdout, /저장: .*보고-수정\.hwpx \(내장 엔진\(rhwp\), \d+쪽\)/);
    assert.deepEqual(bytes('보고.hwpx'), before, '원본이 바뀌면 안 됨');

    const h = Hwp.open(tmp.path('out/보고-수정.hwpx'));
    try {
      const d = h.inspect({ limit: 1000 });
      assert.equal(d.paragraphs[5].text, '맺음말을 덧붙입니다');
      assert.equal(d.tables[0].rows, 4);
      assert.equal(d.counts.pictures, 1);
      assert.ok(d.headers.some((x) => x.text.includes('분기 보고 — 내부용')), JSON.stringify(d.headers));
      assert.ok(d.footers.length >= 1);
      assert.equal(d.fields.find((f) => f.name === '담당자')?.value, '김영희');
      const t = h.text();
      assert.ok(t.includes('150') && t.includes('합계'), t);
    } finally { h.close(); }
    // 눈으로 볼 수 있게 그려지는지
    const img = await cli(['image', 'out/보고-수정.hwpx', '1', '-o', 'out/보고-수정.png', '--dpi', '48']);
    ok(img);
    assert.ok(isPng(bytes('out/보고-수정.png')));
  });

  test('edit: 하나라도 실패하면 아무것도 저장하지 않음', async () => {
    copy('보고.hwpx', '실패.hwpx');
    const before = bytes('실패.hwpx');
    const ops = json('실패작업.json', [
      { op: 'set_paragraph_text', paragraph: 2, text: '바뀌면 안 됨' },
      { op: 'table_set_cell', table: 9, row: 1, col: 1, text: 'x' },
    ]);
    fails(await cli(['edit', '실패.hwpx', ops]), /2번째 작업\(table_set_cell\) 실패: 표 번호.*9.*— 저장하지 않았습니다/);
    assert.deepEqual(bytes('실패.hwpx'), before);
    fails(await cli(['edit', '실패.hwpx', ops, '-o', 'out/실패-결과.hwpx']), /저장하지 않았습니다/);
    assert.ok(!existsSync(tmp.path('out/실패-결과.hwpx')));
  });

  test('edit: 입력 오류는 몇 번째 작업이 왜 틀렸는지 한국어로', async () => {
    let r = await cli(['edit', '보고.hwpx', '[{"op":"add_heading","text":"제목"},{"op":"nope"}]']);
    fails(r, /2번째 작업\(nope\) 입력 오류: 알 수 없는 작업\(op\)입니다: nope \(가능: add_paragraph, .*page_setup\)/);
    r = await cli(['edit', '보고.hwpx', '[{"op":"table_set_cell","table":"첫째","row":1,"col":1,"text":"x"}]']);
    fails(r, /1번째 작업\(table_set_cell\) 입력 오류: 'table' .*number/);
    fails(await cli(['edit', '보고.hwpx', '{"이름":"값"}']), /작업 목록은 \[\{"op"/);
    fails(await cli(['edit', '보고.hwpx', '[]']), /작업 목록이 비어 있습니다/);
    fails(await cli(['edit', '보고.hwpx', '없는파일.json']), /작업 목록 파일이 없습니다/);
    fails(await cli(['edit', '보고.hwpx', '[{"op":']), /JSON으로 읽을 수 없습니다/);
    fails(await cli(['edit', '보고.hwpx']), /인수가 부족합니다/);
  });

  test('edit: 작업 목록은 JSON 문자열·{operations}·작업 하나도 됨, -o 없으면 원본에 저장', async () => {
    copy('보고.hwpx', '제자리.hwpx');
    let r = await cli(['edit', '제자리.hwpx', '{"op":"replace","find":"비용","replace":"지출"}']);
    ok(r);
    assert.match(r.stdout, /1\. 바꾸기 '비용' → '지출': 2곳/);
    assert.match(r.stdout, /저장: .*제자리\.hwpx/);
    r = await cli(['edit', '제자리.hwpx', json('묶음.json', { operations: [{ op: 'add_heading', text: '부록', level: 2 }] })]);
    ok(r);
    const h = Hwp.open(tmp.path('제자리.hwpx'));
    try { const t = h.text(); assert.ok(t.includes('지출 절감') && t.includes('부록') && !t.includes('비용'), t); } finally { h.close(); }
  });

  test('edit: -o 에 이미 있는 다른 파일은 --overwrite 가 있어야 덮어씀, 저장 형식 확인', async () => {
    writeFileSync(tmp.path('있음.hwpx'), 'x');
    const op = '[{"op":"add_paragraph","text":"끝"}]';
    fails(await cli(['edit', '보고.hwpx', op, '-o', '있음.hwpx']), /이미 있는 파일입니다: .*있음\.hwpx \(덮어쓰려면 --overwrite\)/);
    assert.equal(readFileSync(tmp.path('있음.hwpx'), 'utf8'), 'x');
    ok(await cli(['edit', '보고.hwpx', op, '-o', '있음.hwpx', '--overwrite']));
    assert.ok(bytes('있음.hwpx').length > 1000);
    ok(await cli(['edit', '보고.hwpx', op, '-o', 'out/형식바꿈.hwp']));
    const h = Hwp.open(tmp.path('out/형식바꿈.hwp'));
    try { assert.equal(h.format.toLowerCase(), 'hwp'); } finally { h.close(); }
    fails(await cli(['edit', '보고.hwpx', op, '-o', 'out/x.docx']), /\.hwp 또는 \.hwpx로만 저장/);
  });

  test('--engine app 은 한글 전용 기능에 쓸 수 없음 (한컴 COM을 부르지 않음)', async () => {
    fails(await cli(['inspect', '보고.hwpx', '--engine', 'app']), /내장 엔진\(rhwp\)으로만/);
    fails(await cli(['edit', '보고.hwpx', '[{"op":"add_page_break"}]', '--engine', 'app']), /내장 엔진\(rhwp\)으로만/);
  });
});

// ── Excel 파일 엔진 (Excel 없이) ─────────────────────────

describe('Excel inspect / edit (파일 엔진)', () => {
  before(async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('성적');
    ws.addRows([['이름', '점수'], ['가', 90], ['나', 75], ['다', 82]]);
    ws.getCell('C2').value = { formula: 'B2*2' };
    await wb.xlsx.writeFile(tmp.path('성적.xlsx'));
  });

  test('inspect: 시트·사용 범위·수식', async () => {
    const r = await cli(['inspect', '성적.xlsx']);
    ok(r);
    assert.match(r.stdout, /엔진: 파일/);
    assert.match(r.stdout, /1\. '성적' — 사용 범위 A1:C4 \(4행×3열\)/);
    assert.match(r.stdout, /'성적'!C2 =?B2\*2/);
    const d = JSON.parse((await cli(['inspect', '성적.xlsx', '--json'])).stdout);
    assert.equal(d.sheets[0].name, '성적');
  });

  test('edit --engine file: 행 넣기·셀·서식·틀 고정·조건부 서식·시트 추가', async () => {
    const ops = json('엑셀작업.json', [
      { op: 'insert_rows', at: 1 },
      { op: 'set_cells', cells: { A1: '3학년 성적', D3: '=B3+1' } },
      { op: 'format', range: 'A2:C2', bold: true, fill: '#FFF2CC' },
      { op: 'freeze', cell: 'A3' },
      { op: 'conditional_format', range: 'B3:B5', rule: { type: 'greater_than', value: 80 } },
      { op: 'add_sheet', name: '요약' },
    ]);
    const r = await cli(['edit', '성적.xlsx', ops, '-o', 'out/성적-수정.xlsx', '--engine', 'file']);
    ok(r);
    assert.match(r.stdout, /^1\. 1행 앞에 1행 넣음/m);
    assert.match(r.stdout, /^6\. 시트 추가: '요약'/m);
    assert.match(r.stdout, /저장: .*성적-수정\.xlsx \(파일 엔진\(exceljs\)\)/);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp.path('out/성적-수정.xlsx'));
    const ws = wb.getWorksheet('성적');
    assert.equal(ws.getCell('A1').value, '3학년 성적');
    assert.equal(ws.getCell('A3').value, '가');
    assert.equal(ws.getCell('C3').value.formula, 'B3*2', '수식 참조가 함께 밀려야 함');
    assert.equal(ws.getCell('A2').font?.bold, true);
    assert.equal(ws.views[0].state, 'frozen');
    assert.equal(ws.views[0].ySplit, 2);
    assert.ok(wb.getWorksheet('요약'));
    const after = await cli(['inspect', 'out/성적-수정.xlsx']);
    assert.match(after.stdout, /틀 고정 A3/);
    assert.match(after.stdout, /조건부 서식 1개/);
  });

  test('edit --engine file: Excel이 필요한 작업은 거절하고 저장하지 않음 (명령줄 옵션 표기)', async () => {
    copy('성적.xlsx', '정렬.xlsx');
    const before = bytes('정렬.xlsx');
    const r = await cli(['edit', '정렬.xlsx', '[{"op":"set_cells","cells":{"A1":"x"}},{"op":"sort","range":"A1:B4","by":"B"}]', '--engine', 'file']);
    fails(r, /할 수 없는 작업입니다: sort/);
    assert.doesNotMatch(r.stderr, /engine:'/);
    assert.deepEqual(bytes('정렬.xlsx'), before);
    fails(await cli(['edit', '정렬.xlsx', '[{"op":"merge"}]', '--engine', 'file']), /1번째 작업\(merge\)/);
  });

  test('edit --json: 결과를 JSON으로', async () => {
    copy('성적.xlsx', '제이슨.xlsx');
    const r = await cli(['edit', '제이슨.xlsx', '[{"op":"column_width","columns":"A:B","width":14}]', '--engine', 'file', '--json']);
    ok(r);
    const d = JSON.parse(r.stdout);
    assert.equal(d.format, 'excel');
    assert.equal(d.engine, 'file');
    assert.ok(d.output.endsWith('제이슨.xlsx'));
    assert.equal(d.results.length, 1);
    assert.match(d.results[0], /^1\. 열 너비/);
  });
});

// ── 형식·엔진·옵션 확인 (프로그램 없이) ─────────────────

describe('형식·엔진·옵션 확인', () => {
  before(async () => {
    await createDocument(tmp.path('문서.docx'), MANUSCRIPT);
    await createDocument(tmp.path('발표.pptx'), MANUSCRIPT);
    writeFileSync(tmp.path('메모.txt'), '글', 'utf8');
  });

  test('형식별 기능이 없는 파일', async () => {
    fails(await cli(['inspect', '메모.txt']), /형식별 기능을 쓸 수 없는 파일입니다: 메모\.txt/);
    fails(await cli(['edit', '메모.txt', '[{"op":"x"}]']), /형식별 기능을 쓸 수 없는 파일/);
    fails(await cli(['inspect', '없는.docx']), /파일이 없습니다/);
  });

  test('Word·PowerPoint 전용 기능은 --engine file 불가', async () => {
    fails(await cli(['inspect', '문서.docx', '--engine', 'file']), /Word 전용 기능은 설치된 Word로만/);
    fails(await cli(['edit', '발표.pptx', '[{"op":"add_slide"}]', '--engine', 'file']), /PowerPoint 전용 기능은 설치된 PowerPoint로만/);
  });

  test('정수 옵션 확인', async () => {
    fails(await cli(['inspect', '보고없음.hwpx', '--offset', '-1']), /--offset 은 0 이상의 정수|파일이 없습니다/);
    await createDocument(tmp.path('옵션.hwpx'), MANUSCRIPT);
    fails(await cli(['inspect', '옵션.hwpx', '--offset', '-1']), /--offset 은 0 이상의 정수여야 합니다/);
    fails(await cli(['inspect', '옵션.hwpx', '--limit', '0']), /--limit 은 1 이상의 정수여야 합니다/);
    fails(await cli(['inspect', '옵션.hwpx', '--limit', '5000']), /--limit 은 1000 이하/);
    fails(await cli(['inspect', '옵션.hwpx', '--slide']), /--slide 뒤에 값이 필요합니다/);
    const r = await cli(['inspect', '옵션.hwpx', '--slide', '2']);
    ok(r);
    assert.match(r.stderr, /--slide는 PowerPoint 문서에만/);
  });
});

// ── 앱 AI 지침·도구 이름 (형식별 도구) ──────────────────

describe('앱 AI: 형식별 도구 안내', () => {
  test('지침: 열린 문서 형식의 inspect → edit, 단순 글자는 office_replace/fill, 확인은 office_page_image', async () => {
    const { systemPrompt } = await import('../src/app/ai.mjs');
    const cases = { '보고.docx': 'word', '발표.pptx': 'ppt', '표.xlsx': 'excel', '신청서.hwpx': 'hwp' };
    for (const [name, t] of Object.entries(cases)) {
      const s = systemPrompt({ filePath: `C:\\문서\\${name}` });
      assert.match(s, new RegExp(`이 문서의 형식별 도구: ${t}_inspect → ${t}_edit`), name);
      assert.match(s, /office_replace \/ office_fill/);
      assert.match(s, /office_page_image/);
      assert.match(s, /아무것도 저장되지 않으니/);
    }
    assert.match(systemPrompt({ filePath: 'C:\\문서\\발표.pptx' }), /발표자 노트/);
    const none = systemPrompt({});
    for (const t of ['hwp', 'word', 'ppt', 'excel']) assert.ok(none.includes(`${t}_inspect·${t}_edit`), none);
  });

  test('도구 이름을 한국어로', async () => {
    const { toolLabel } = await import('../src/app/ai.mjs');
    assert.equal(toolLabel('mcp__office-kit__word_edit'), 'Word 편집');
    assert.equal(toolLabel('mcp__office-kit__ppt_edit'), '슬라이드 편집');
    assert.equal(toolLabel('mcp__office-kit__excel_edit'), '시트 편집');
    for (const t of ['hwp', 'word', 'ppt', 'excel']) assert.equal(toolLabel(`mcp__office-kit__${t}_inspect`), '문서 구조 살펴보기');
  });

  test('화면: 편집 도구 뒤 미리보기 새로 고침, 문서 종류별 추천 문장', () => {
    const src = readFileSync(join(ROOT, 'src', 'app', 'renderer', 'app.js'), 'utf8');
    const edit = src.match(/const EDIT_TOOLS = new Set\(\[([\s\S]*?)\]\)/)?.[1] ?? '';
    for (const t of ['hwp_edit', 'word_edit', 'ppt_edit', 'excel_edit']) assert.ok(edit.includes(`'${t}'`), `${t}가 EDIT_TOOLS에 없음`);
    for (const s of ['발표자 노트 달아줘', '슬라이드 하나 추가해줘', '차트 만들어줘', '머리글 굵게 하고 틀 고정해줘', '제목 스타일 정리해줘', '목차 넣어줘', '머리말 넣어줘']) {
      assert.ok(src.includes(`'${s}'`), s);
    }
    for (const t of ['word_edit', 'ppt_edit', 'excel_edit', 'word_inspect', 'ppt_inspect', 'excel_inspect', 'hwp_inspect']) {
      assert.match(src, new RegExp(`${t}: \\['[^']+ 중', '[^']+'\\]`), t);
    }
  });
});

// ── COM (Word / PowerPoint / Excel) ──────────────────────

describe('COM inspect / edit — OFFICEKIT_COM_TESTS=1 일 때만', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  const LONG = { timeout: 600_000 };
  const comCli = (args) => cli(args, { timeout: 540_000 });

  before(async () => {
    await createDocument(tmp.path('com.docx'), MANUSCRIPT);
    await createDocument(tmp.path('com.pptx'), MANUSCRIPT);
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('매출').addRows([['월', '매출', '비용'], ['1월', 120, 80], ['2월', 150, 90], ['3월', 90, 70]]);
    await wb.xlsx.writeFile(tmp.path('com.xlsx'));
  });

  test('Word: inspect → edit → inspect', LONG, async () => {
    let r = await comCli(['inspect', 'com.docx', '--limit', '5']);
    ok(r);
    assert.match(r.stdout, /Word 문서: \d+쪽/);
    assert.match(r.stdout, /문단 1~5 \/ 전체 \d+개/);
    const before = bytes('com.docx');
    r = await comCli(['edit', 'com.docx', json('워드작업.json', [
      { op: 'insert_paragraph', text: '들어가는 말', at: { after_paragraph: 1 }, style: 'normal' },
      { op: 'format_text', find: '매출', all: true, bold: true, highlight: 'yellow' },
      { op: 'set_footer', page_numbers: true },
      { op: 'page_setup', orientation: 'landscape' },
    ]), '-o', 'out/com-수정.docx']);
    ok(r);
    assert.match(r.stdout, /^1\. /m);
    assert.match(r.stdout, /^4\. /m);
    assert.match(r.stdout, /저장: .*com-수정\.docx \(Word\)/);
    assert.deepEqual(bytes('com.docx'), before);
    r = await comCli(['inspect', 'out/com-수정.docx', '--json']);
    ok(r);
    const d = JSON.parse(r.stdout);
    assert.equal(d.sections[0].orientation, 1, '가로');
    assert.ok(d.paragraphs.some((p) => p.i === 2 && p.text.startsWith('들어가는 말')), JSON.stringify(d.paragraphs.slice(0, 3)));
    // 실패하면 저장하지 않음
    fails(await comCli(['edit', 'com.docx', '[{"op":"insert_paragraph","text":"x"},{"op":"table_set_cell","table":7,"row":1,"col":1,"text":"x"}]']),
      /2번째 작업\(table_set_cell\) 실패: 표 번호 7.*저장하지 않았습니다/);
    assert.deepEqual(bytes('com.docx'), before);
  });

  test('PowerPoint: add_slide·set_notes → inspect', LONG, async () => {
    const r = await comCli(['edit', 'com.pptx', json('슬라이드작업.json', [
      { op: 'add_slide', layout: 'title_content', title: '다음 분기', body: ['신규 고객', { text: '기업 10곳', level: 2 }], notes: '일정 설명' },
      { op: 'set_notes', slide: 1, text: '인사 후 시작' },
    ]), '-o', 'out/com-수정.pptx']);
    ok(r);
    assert.match(r.stdout, /^1\. 슬라이드 3 추가/m);
    assert.match(r.stdout, /저장: .*com-수정\.pptx \(PowerPoint\)/);
    const one = await comCli(['inspect', 'out/com-수정.pptx', '--slide', '3']);
    ok(one);
    assert.match(one.stdout, /\[슬라이드 3\]/);
    assert.match(one.stdout, /제목: 다음 분기/);
    assert.match(one.stdout, /노트: 일정 설명/);
    const d = JSON.parse((await comCli(['inspect', 'out/com-수정.pptx', '--json'])).stdout);
    assert.equal(d.slideCount, 3);
    assert.equal(d.slides[0].notes, '인사 후 시작');
  });

  test('Excel (Excel로): add_chart·sort → inspect', LONG, async () => {
    const r = await comCli(['edit', 'com.xlsx', json('차트작업.json', [
      { op: 'sort', range: 'A1:C4', by: 'B', order: 'desc', header: true },
      { op: 'add_chart', range: 'A1:C4', type: 'column', title: '월별 매출', at: 'E2' },
    ]), '-o', 'out/com-수정.xlsx']);
    ok(r);
    assert.match(r.stdout, /저장: .*com-수정\.xlsx \(Excel\)/);
    const i = await comCli(['inspect', 'out/com-수정.xlsx', '--engine', 'app']);
    ok(i);
    assert.match(i.stdout, /엔진: Excel/);
    assert.match(i.stdout, /차트 1개/);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp.path('out/com-수정.xlsx'));
    assert.equal(wb.getWorksheet('매출').getCell('A2').value, '2월', '매출 큰 순서');
  });
});

// ── HWPX 일괄 변환 · 개인정보 ──────────────────────────────

describe('hwpx · pii', () => {
  test('hwpx: 폴더의 hwp → hwpx, 보고서, 다시 하면 건너뜀', async () => {
    copyFileSync(join(ROOT, 'examples', '신청서.hwp'), tmp.path('일괄.hwp'));
    let r = await cli(['hwpx', tmp.path('일괄.hwp'), '--report', '결과.csv']);
    ok(r);
    assert.match(r.stdout, /\[1\/1\] 같음/);
    assert.match(r.stdout, /완료: 같음 1/);
    assert.ok(existsSync(tmp.path('일괄.hwpx')));
    assert.match(readFileSync(tmp.path('결과.csv'), 'utf8'), /같음/);
    r = await cli(['hwpx', tmp.path('일괄.hwp')]);
    ok(r);
    assert.match(r.stdout, /건너뜀 1/);
    fails(await cli(['hwpx', tmp.path('없는폴더')]), /찾지 못했습니다/);
  });

  test('pii: 찾기 → --mask 가린 사본', async () => {
    await createDocument(tmp.path('연락처.docx'), '# 연락처\n\n홍길동 010-1234-5678 / hong@example.com');
    let r = await cli(['pii', '연락처.docx']);
    ok(r);
    assert.match(r.stdout, /휴대전화 {2}010-1234-5678 → 010-\*{4}-5678/);
    assert.match(r.stdout, /이메일/);
    r = await cli(['pii', '연락처.docx', '--mask', '--types', 'mobile']);
    ok(r);
    assert.match(r.stdout, /가린 사본: .*연락처_가림\.docx/);
    assert.match(r.stdout, /남은 것이 없습니다/);
    const t = await cli(['text', '연락처_가림.docx']);
    assert.match(t.stdout, /010-\*{4}-5678/);
    assert.match(t.stdout, /hong@example\.com/);
    fails(await cli(['pii', '연락처.docx', '--types', 'xyz']), /알 수 없는 종류/);
  });
});
