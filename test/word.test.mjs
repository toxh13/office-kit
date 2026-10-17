// word.test.mjs — Word 전용 도구 (word_inspect / word_edit).
// 입력 확인·MCP 등록은 Word 없이 돈다. 실제 Word(COM) 시험은 OFFICEKIT_COM_TESTS=1 일 때만.
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmodSync, copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { Resvg } from '@resvg/resvg-js';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, PageOrientation, SectionType,
} from 'docx';
import { COM_TESTS, ROOT, tempDir, isPng, zipText } from './helpers/fixtures.mjs';
import { createDocument } from '../src/core/index.mjs';
import { pageImage } from '../src/core/render.mjs';
import { wordInspect, wordEdit, formatWordInspect, normalizeWordOps, WORD_OPS } from '../src/core/word.mjs';

const SERVER = join(ROOT, 'src', 'mcp', 'server.mjs');
const tmp = tempDir('officekit-word-');
after(() => tmp.cleanup());

// 문단 번호: 1 제목1, 2 가운데 줄, 3 제목2, 4·5 글머리표, 6~14 표(3행×2열 + 행 끝 표시), 15 본문, 16 굵은 줄, 17 마지막
const DOC = `# 분기 보고
>> 2026년 3분기
## 요약
- 매출 증가
- 비용 절감
| 항목 | 값 |
|---|---|
| 매출 | 120 |
| 비용 | 80 |
본문 문단입니다. 매출이 늘었습니다.
**신규 고객 확보**
마지막 문단`;

const PNG = tmp.path('blue.png');
writeFileSync(PNG, new Resvg('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200"><rect width="400" height="200" fill="#3366cc"/></svg>').render().asPng());

// 구역 2개(세로·가로), 표 앞 제목, 세로 병합 표, 직접 쓴 번호가 있는 문서. 문단 번호:
// 1 제목 '1. 개요', 2 본문(매출 3번), 3 제목 '표 제목'(표 바로 앞), 4~9 표1(2행×2열), 10 '표 뒤 문단 매출',
// 11~25 표2: 1행 [지역 | 상·하반기(가로 병합)] 2행 [서울(세로 병합) | 50 | 60] 3행 [(병합) | 55 | 65] 4행 [부산 | 30 | 35]
// (18 = 서울 칸의 아랫부분), 26 '1. 첫째', 27 '- 둘째', 28 '셋째', 29 '구역 끝 문단'(구역 나누기), 30 제목 '부록', 31 '부록 본문'
async function makeRich(path) {
  const p = (t) => new Paragraph({ children: [new TextRun(t)] });
  const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
  const cell = (t, o = {}) => new TableCell({ children: [new Paragraph(String(t))], ...o });
  const doc = new Document({
    sections: [
      {
        children: [
          h1('1. 개요'), p('본문 하나: 매출 매출 매출'), h1('표 제목'),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [['가', '나'], ['다', '라']].map((r) => new TableRow({ children: r.map((t) => cell(t)) })) }),
          p('표 뒤 문단 매출'),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({ children: [cell('지역'), cell('상·하반기', { columnSpan: 2 })] }),
              new TableRow({ children: [cell('서울', { rowSpan: 2 }), cell('50'), cell('60')] }),
              new TableRow({ children: [cell('55'), cell('65')] }),
              new TableRow({ children: [cell('부산'), cell('30'), cell('35')] }),
            ],
          }),
          p('1. 첫째'), p('- 둘째'), p('셋째'), p('구역 끝 문단'),
        ],
      },
      { properties: { type: SectionType.NEXT_PAGE, page: { size: { orientation: PageOrientation.LANDSCAPE } } }, children: [h1('부록'), p('부록 본문')] },
    ],
  });
  writeFileSync(path, await Packer.toBuffer(doc));
}

// 다른 프로그램이 파일을 읽기 공유로 열어 둔 상태 (사용자가 Word로 열어 둔 것과 같음)
function holdOpen(path) {
  const ps = spawn('powershell', ['-NoProfile', '-Command', `$s=[IO.File]::Open('${path}','Open','Read','Read'); Start-Sleep 600; $s.Close()`], { stdio: 'ignore' });
  return { ready: new Promise((r) => setTimeout(r, 3000)), close: () => ps.kill() };
}

// 모든 작업의 올바른 예 (입력 확인 시험과 MCP 스키마 시험에 같이 쓴다)
const SAMPLE = {
  insert_paragraph: { text: '새 문단\n둘째 줄', at: { after_paragraph: 1 }, style: 'heading2', format: { bold: true, color: 'red' } },
  set_paragraph_text: { paragraph: 2, text: '바뀐 글' },
  delete_paragraph: { find: '마지막', all: true },
  apply_style: { paragraphs: [1, 2], style: 'Heading 1' },
  format_text: { find: '매출', all: true, italic: true, highlight: 'yellow', size: 12, font: '맑은 고딕' },
  format_paragraph: { paragraph: 3, align: 'center', line_spacing: 1.5, space_after: 6, first_line_indent_mm: 5, keep_with_next: true },
  set_list: { paragraphs: [4, 5], type: 'number' },
  insert_table: { rows: [['a', 'b'], [1, null]], at: 'start', style: 'grid', header_bold: false, autofit: true },
  table_set_cell: { table: 1, row: 1, col: 2, text: 'x' },
  table_add_row: { table: 1, values: ['c', 'd'], after_row: 0 },
  table_delete_row: { table: 1, row: 2 },
  insert_image: { image_path: PNG, at: 'end', width_mm: 40, caption: '파란 상자' },
  insert_page_break: { at: { before_paragraph: 3 } },
  set_header: { text: '머리말', align: 'right' },
  set_footer: { text: '{page} / {pages}', page_numbers: true },
  page_setup: { orientation: 'landscape', paper: 'a4', margins_mm: { top: 20, left: 15 } },
  add_comment: { find: '매출', text: '메모', author: '검토자' },
  delete_comments: { all: true },
  track_changes: { on: true },
  accept_revisions: {},
  reject_revisions: {},
  insert_toc: { at: 'start', levels: 2, title: '목차' },
  update_fields: {},
  add_footnote: { find: '매출', text: '각주' },
  insert_hyperlink: { find: '매출', url: 'https://example.com' },
};

describe('입력 확인 (Word 없이)', () => {
  test('모든 작업의 올바른 예가 통과하고 값이 정리된다', () => {
    assert.deepEqual(Object.keys(SAMPLE).sort(), [...WORD_OPS].sort());
    const ops = normalizeWordOps(WORD_OPS.map((op) => ({ op, ...SAMPLE[op] })));
    const by = Object.fromEntries(ops.map((o) => [o.op, o]));
    assert.equal(by.insert_paragraph.text, '새 문단\v둘째 줄', '줄바꿈은 문단 안 줄바꿈');
    assert.equal(by.insert_paragraph.format.color, '#ff0000');
    assert.equal(by.insert_table.autofit, 'content');
    assert.deepEqual(by.insert_table.rows, [['a', 'b'], ['1', '']]);
    assert.equal(by.page_setup.paper, 'A4');
    assert.equal(by.insert_image.caption_label, '그림');
    assert.equal(by.insert_image.image_path, PNG);
  });

  test('잘못된 입력은 몇 번째 작업인지 한국어로', () => {
    const bad = (ops, re) => assert.throws(() => normalizeWordOps(ops), re);
    bad([], /비어 있습니다/);
    bad([{ op: 'update_fields' }, { op: 'nope' }], /2번째 작업\(nope\) 입력 오류: 알 수 없는 작업/);
    bad([{ op: 'set_paragraph_text', paragraph: 0, text: 'x' }], /paragraph/);
    bad([{ op: 'apply_style', paragraph: 1, find: 'a', style: 'normal' }], /함께 쓸 수 없습니다/);
    bad([{ op: 'apply_style', paragraphs: [5, 2], style: 'normal' }], /시작\(5\)이 끝\(2\)보다/);
    bad([{ op: 'format_text', find: '매출' }], /바꿀 서식이 없습니다/);
    bad([{ op: 'format_text', find: '매출', color: '파랑' }], /#rrggbb/);
    bad([{ op: 'format_text', find: 'x'.repeat(256), bold: true }], /255자/);
    bad([{ op: 'insert_paragraph', text: 'x', at: { after: 3 } }], /at은/);
    bad([{ op: 'insert_table', rows: [] }], /2차원 배열/);
    bad([{ op: 'insert_image', image_path: tmp.path('없음.png') }], /그림 파일이 없습니다/);
    bad([{ op: 'insert_image', image_path: SERVER }], /png, jpg/);
    bad([{ op: 'insert_hyperlink', find: 'a', url: 'javascript:alert(1)' }], /http/);
    bad([{ op: 'delete_comments' }], /all:true/);
    bad([{ op: 'set_footer' }], /page_numbers/);
    bad([{ op: 'page_setup', paper: 'A9' }], /paper/);
    bad([{ op: 'track_changes' }], /on/);
  });

  test('Word 문서가 아니면 거부', async () => {
    const hwp = tmp.path('a.hwp');
    writeFileSync(hwp, 'x');
    await assert.rejects(wordEdit(hwp, [{ op: 'update_fields' }]), /Word 문서\(\.docx\/\.doc\/\.rtf\/\.odt\)만/);
    await assert.rejects(wordInspect(hwp), /hwp_edit/);
    await assert.rejects(wordInspect(tmp.path('없음.docx')), /파일이 없습니다/);
    const docx = tmp.path('v.docx');
    writeFileSync(docx, 'x');
    await assert.rejects(wordEdit(docx, [{ op: 'update_fields' }], { output: tmp.path('v.pdf') }), /office_convert/);
    await assert.rejects(wordInspect(docx, { limit: 0 }), /limit/);
  });

  test('살펴보기 글: 필드 코드·특수 문자·다음 위치', () => {
    const t = formatWordInspect({
      path: 'x.docx', pages: 3, words: 10, characters: 40, paragraphTotal: 30, tableCount: 0, offset: 0, limit: 2, sectionCount: 1,
      sections: [{ index: 1, width: 210, height: 297, top: 25, bottom: 25, left: 30, right: 30, orientation: 0,
        footer: '- \x13 PAGE \x15 / \x13 NUMPAGES \x15 -', footerFields: [33, 26], header: '', headerFields: [] }],
      trackRevisions: false, revisionCount: 0, revisions: [], commentCount: 0, comments: [], counts: {},
      styles: [{ name: '제목 1', key: 'heading1' }],
      paragraphs: [{ i: 1, style: '제목 1', key: 'heading1', text: '가\v나\f' }, { i: 2, style: '표준', text: '/', images: 1, list: String.fromCodePoint(0xf0b7), listType: 2 }],
    });
    assert.match(t, /A4 세로 210×297mm/);
    assert.match(t, /꼬리말: '- \{page\} \/ \{pages\} -' \[쪽 번호, 전체 쪽 수\]/);
    assert.match(t, /1 \[제목 1·heading1\] 가↵나\[쪽 나누기\]/);
    assert.match(t, /2 \[표준 · 글머리표 •\] \[그림\]/);
    assert.match(t, /offset=2/);
  });

  test('살펴보기 글: 병합된 칸·구역 나누기·지움 표시·메모 목록 안내', () => {
    const base = { path: 'x.docx', pages: 1, words: 1, characters: 1, paragraphTotal: 50, tableCount: 0, sectionCount: 0, sections: [],
      trackRevisions: true, revisionCount: 1, revisions: [], counts: {}, styles: [] };
    const t = formatWordInspect({ ...base, offset: 0, limit: 3, commentCount: 150, commentsListed: true,
      comments: Array.from({ length: 100 }, (_, i) => ({ index: i + 1, author: 'a', text: `메모 ${i + 1}`, scope: 'x' })),
      paragraphs: [{ i: 1, merged: true, table: 2 }, { i: 2, style: '표준', text: '끝 문단', sectionBreak: true }, { i: 3, style: '표준', text: '지운 문단', deleted: true }] });
    assert.match(t, /1 \[표2 병합된 칸 — 위 칸에 세로로 합쳐진 자리라 글을 넣을 수 없음\]/);
    assert.match(t, /2 \[표준\] 끝 문단 \[구역 나누기\]/);
    assert.doesNotMatch(t, /쪽 나누기/);
    assert.match(t, /3 \[표준\] \[지움 표시됨 — 변경 내용 수락 전\] 지운 문단/);
    assert.match(t, /\(100개만 표시\)/);
    const t2 = formatWordInspect({ ...base, revisionCount: 0, offset: 3, limit: 3, commentCount: 150, commentsListed: false, comments: [], paragraphs: [] });
    assert.match(t2, /메모 150개 \(목록은 offset=0일 때만 보여 줌\)/);
    assert.doesNotMatch(t2, /개만 표시/);
  });

  test('그림 파일은 내용(파일 머리)까지 확인', () => {
    const fake = tmp.path('fake.png');
    writeFileSync(fake, 'not an image');
    assert.throws(() => normalizeWordOps([{ op: 'insert_image', image_path: fake }]), /그림 파일이 아니거나 깨졌습니다: fake\.png/);
    const jpgNamedPng = tmp.path('photo.png');
    writeFileSync(jpgNamedPng, Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 0x4a, 0x46]));
    assert.equal(normalizeWordOps([{ op: 'insert_image', image_path: jpgNamedPng }])[0].image_path, jpgNamedPng, '확장자와 달라도 그림이면 받음');
  });

  test('읽기 전용 파일에 그대로 저장하려 하면 Word를 띄우기 전에 거절', { skip: process.platform !== 'win32' }, async () => {
    const ro = tmp.path('ro.docx');
    writeFileSync(ro, 'x');
    chmodSync(ro, 0o444);
    try {
      const t0 = Date.now();
      await assert.rejects(wordEdit(ro, [{ op: 'update_fields' }]), /읽기 전용이거나 다른 프로그램\(Word 등\)에서 열려 있어.*output으로 다른 경로/);
      assert.ok(Date.now() - t0 < 2000);
      const src = tmp.path('ro-src.docx');
      const out = tmp.path('ro-out.docx');
      writeFileSync(src, 'x');
      writeFileSync(out, 'x');
      chmodSync(out, 0o444);
      try {
        await assert.rejects(wordEdit(src, [{ op: 'update_fields' }], { output: out }), /저장할 파일이 읽기 전용입니다: .*ro-out\.docx/);
      } finally { chmodSync(out, 0o666); }
    } finally { chmodSync(ro, 0o666); }
  });
});

describe('MCP 등록 (Word 없이)', () => {
  let client;
  before(async () => {
    client = new Client({ name: 'word-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [SERVER], cwd: tmp.dir, stderr: 'pipe' }));
  });
  after(async () => { await client?.close(); });
  const textOf = (r) => r.content.filter((c) => c.type === 'text').map((c) => c.text).join('\n');

  test('tools/list에 word_inspect·word_edit, 설명에 모든 작업', async () => {
    const { tools } = await client.listTools();
    const names = tools.map((t) => t.name);
    assert.ok(names.includes('word_inspect') && names.includes('word_edit'), names.join(', '));
    const edit = tools.find((t) => t.name === 'word_edit');
    const inspect = tools.find((t) => t.name === 'word_inspect');
    assert.match(inspect.description, /[가-힣]/);
    assert.match(edit.description, /저장하지 않는다/);
    const schema = JSON.stringify(edit.inputSchema);
    for (const op of WORD_OPS) {
      assert.ok(edit.description.includes(`'${op}'`), `설명에 ${op}`);
      assert.ok(schema.includes(`"${op}"`), `스키마에 ${op}`);
    }
    assert.deepEqual(Object.keys(inspect.inputSchema.properties).sort(), ['limit', 'offset', 'path']);
    assert.equal(edit.annotations?.destructiveHint, true);
  });

  test('스키마가 모든 작업의 예를 받아들이고, Word 문서가 아니면 거부', async () => {
    writeFileSync(tmp.path('m.xlsx'), 'x');
    const r = await client.callTool({ name: 'word_edit', arguments: { path: 'm.xlsx', operations: WORD_OPS.map((op) => ({ op, ...SAMPLE[op] })) } });
    assert.equal(r.isError, true);
    assert.match(textOf(r), /Word 문서\(\.docx\/\.doc\/\.rtf\/\.odt\)만 다룹니다/);
    const r2 = await client.callTool({ name: 'word_edit', arguments: { path: 'm.xlsx', operations: [{ op: 'bogus' }] } });
    assert.equal(r2.isError, true);
    assert.match(textOf(r2), /^오류: /);
    const r3 = await client.callTool({ name: 'word_inspect', arguments: { path: '없는.docx' } });
    assert.match(textOf(r3), /파일이 없습니다/);
  });

  test('저장 경로 확장자 확인', async () => {
    writeFileSync(tmp.path('m.docx'), 'x');
    const r = await client.callTool({ name: 'word_edit', arguments: { path: 'm.docx', output: 'm.pdf', operations: [{ op: 'update_fields' }] } });
    assert.equal(r.isError, true);
    assert.match(textOf(r), /\.docx\/\.doc\/\.rtf\/\.odt/);
  });

  test('입력 검증 오류도 한국어', async () => {
    writeFileSync(tmp.path('k.docx'), 'x');
    const err = async (operations) => {
      const r = await client.callTool({ name: 'word_edit', arguments: { path: 'k.docx', operations } });
      assert.equal(r.isError, true);
      return textOf(r);
    };
    assert.match(await err([{ op: 'bogus' }]), /알 수 없는 작업\(op\)입니다: bogus \(가능: insert_paragraph, /);
    assert.match(await err([{ op: 'set_paragraph_text', paragraph: 0, text: 'x' }]), /1 이상이어야 합니다 \(번호는 1부터\)/);
    assert.match(await err([{ op: 'format_text', find: '매출', highlight: 'purple' }]), /highlight는 yellow\|/);
    assert.match(await err([{ op: 'insert_table', rows: 'x' }]), /예상 타입은 array/);
    for (const msg of [await err([{ op: 'set_list', paragraph: 1, type: 'roman' }]), await err([{ op: 'table_add_row' }])]) {
      assert.doesNotMatch(msg, /Invalid|Expected|Too small|Required/, msg);
    }
  });
});

// ── 설치된 Word (COM) ────────────────────────────────────

describe('Word (COM) — OFFICEKIT_COM_TESTS=1 일 때만', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  const T = { timeout: 600_000 };
  const base = tmp.path('base.docx');
  const rich = tmp.path('rich.docx');
  const fresh = (name, from = base) => { const p = tmp.path(name); copyFileSync(from, p); return p; };
  const para = (info, text) => info.paragraphs.find((p) => p.text === text);
  const at = (info, i) => info.paragraphs.find((p) => p.i === i);
  before(async () => { await createDocument(base, DOC); await makeRich(rich); });

  test('word_inspect: 문단·표·목록·스타일 키, 나눠 보기', T, async () => {
    const info = await wordInspect(base);
    assert.equal(info.paragraphTotal, 17);
    assert.equal(info.tableCount, 1);
    assert.equal(info.paragraphs.length, 14, '행 끝 표시 3개는 목록에서 빠진다');
    assert.equal(info.paragraphs[0].key, 'heading1');
    assert.equal(info.paragraphs[2].key, 'heading2');
    assert.equal(para(info, '매출 증가').listType, 2);
    assert.deepEqual(info.tables[0], { ...info.tables[0], index: 1, rows: 3, cols: 2, paragraph: 6, first: ['항목', '값'] });
    const cell = para(info, '120');
    assert.deepEqual([cell.i, cell.table, cell.row, cell.col], [10, 1, 2, 2]);
    assert.equal(info.sections[0].width, 210);
    assert.ok(info.styles.some((s) => s.key === 'heading1'));
    const text = formatWordInspect(info);
    assert.match(text, /표1 2행2열\] 120/);
    assert.match(text, /쓰인 스타일/);

    const page = await wordInspect(base, { offset: 10, limit: 3 });
    assert.deepEqual(page.paragraphs.map((p) => p.i), [12, 13], '11번은 행 끝 표시');
    assert.match(formatWordInspect(page), /offset=13/);
  });

  test('문단 작업: 넣기·글 바꾸기·지우기·스타일·글자/문단 서식·목록', T, async () => {
    const src = fresh('p.docx');
    const out = tmp.path('p-out.docx');
    const r = await wordEdit(src, [
      { op: 'insert_paragraph', text: '제목 뒤 문단', at: { after_paragraph: 1 } },
      { op: 'insert_paragraph', text: '맨 앞', at: 'start', style: 'title' },
      { op: 'insert_paragraph', text: '끝 제목', style: 'heading2' },
      { op: 'insert_paragraph', text: '굵은 줄 다음', at: { after_text: '신규 고객' }, format: { color: '#ff0000' } },
      { op: 'set_paragraph_text', paragraph: 4, text: '바뀐 부제' },
      { op: 'apply_style', paragraph: 4, style: 'subtitle' },
      { op: 'apply_style', find: '마지막 문단', style: 'quote' },
      { op: 'format_text', find: '매출', all: true, bold: true, highlight: 'yellow', color: 'blue' },
      { op: 'format_paragraph', paragraphs: [2, 3], align: 'center', space_after: 12, line_spacing: 1.5 },
      { op: 'set_list', paragraphs: [6, 7], type: 'number' },
      { op: 'delete_paragraph', find: '본문 문단' },
    ], { output: out });
    assert.equal(r.lines.length, 12, r.lines.join('\n'));
    assert.match(r.lines[0], /^1\. 문단 넣음: 2번/);
    assert.match(r.lines[7], /'매출' 3곳/);
    assert.match(r.lines.at(-1), /저장: .*p-out\.docx/);
    assert.ok(readFileSync(src).equals(readFileSync(base)), 'output을 주면 원본은 그대로');

    const info = await wordInspect(out);
    const keys = info.paragraphs.slice(0, 4).map((p) => [p.text, p.key]);
    assert.deepEqual(keys, [['맨 앞', 'title'], ['분기 보고', 'heading1'], ['제목 뒤 문단', 'normal'], ['바뀐 부제', 'subtitle']]);
    assert.equal(para(info, '매출 증가').listType, 3, '번호 목록');
    assert.equal(para(info, '비용 절감').list, '2.');
    assert.equal(para(info, '본문 문단입니다. 매출이 늘었습니다.'), undefined);
    assert.equal(para(info, '마지막 문단').key, 'quote');
    assert.equal(info.paragraphs.at(-1).text, '끝 제목');
    assert.equal(info.paragraphs.at(-1).key, 'heading2');
    const bold = para(info, '신규 고객 확보').i;
    assert.equal(para(info, '굵은 줄 다음').i, bold + 1);
    const xml = await zipText(out, 'word/document.xml');
    assert.match(xml, /<w:highlight w:val="yellow"\/>/);
    assert.match(xml, /<w:color w:val="FF0000"\/>/i);
    assert.match(xml, /<w:spacing[^>]*w:line="360"/, '1.5줄');
  });

  test('표 작업: 넣기·칸·행 추가·행 지우기', T, async () => {
    const src = fresh('t.docx');
    const r = await wordEdit(src, [
      { op: 'insert_table', rows: [['이름', '점수'], ['가', 90], ['나', 85.5]], at: { after_paragraph: 5 } },
      { op: 'table_set_cell', table: 2, row: 3, col: 2, text: '95' },
      { op: 'table_add_row', table: 2, values: ['합계', '215'] },
      { op: 'table_add_row', table: 1, after_row: 0, values: ['머리', '칸'] },
      { op: 'table_delete_row', table: 1, row: 3 },
      { op: 'insert_table', rows: [['끝 표']], style: 'light_grid', header_bold: false, autofit: 'content' },
    ]);
    assert.match(r.lines[0], /표 1 넣음: 3행×2열, 6번 문단부터/);
    assert.match(r.lines[1], /'80' → '95'/);
    const info = await wordInspect(src);
    assert.equal(info.tableCount, 3);
    assert.deepEqual(info.tables.map((t) => t.rows), [3, 4, 1]);
    assert.deepEqual(info.tables[0].first, ['머리', '칸']);
    assert.ok(!info.paragraphs.some((p) => p.text === '가'), '지운 행');
    assert.ok(para(info, '95') && para(info, '합계'));
    // 새 표와 원래 표 사이에는 빈 문단이 남아 두 표가 합쳐지지 않는다
    const gap = info.paragraphs.find((p) => p.i === info.tables[1].paragraph - 1);
    assert.equal(gap?.table, undefined);
  });

  test('그림·쪽 나누기·머리말·꼬리말·쪽 설정', T, async () => {
    const src = fresh('g.docx');
    const r = await wordEdit(src, [
      { op: 'insert_image', image_path: PNG, at: { after_paragraph: 1 }, width_mm: 50, caption: '파란 상자' },
      { op: 'insert_image', image_path: PNG },
      { op: 'insert_page_break', at: { after_text: '신규 고객' } },
      { op: 'set_header', text: '머리말 글', align: 'right' },
      { op: 'set_footer', text: '{page} / {pages}', page_numbers: true },
      { op: 'page_setup', orientation: 'landscape', paper: 'A4', margins_mm: { top: 20, bottom: 20, left: 15, right: 15 } },
    ]);
    assert.match(r.lines[0], /그림 넣음: 2번 문단, 50mm×25mm, 캡션 3번 문단/);
    const info = await wordInspect(src);
    assert.equal(info.counts.inlineShapes, 2);
    const cap = info.paragraphs.find((p) => p.i === 3);
    assert.equal(cap.key, 'caption');
    assert.match(cap.text, /^그림 1\. 파란 상자$/);
    assert.equal(info.paragraphs.find((p) => p.i === 2).images, 1);
    const brk = para(info, '신규 고객 확보').i + 1;
    assert.equal(info.paragraphs.find((p) => p.i === brk).text, '\f');
    const s = info.sections[0];
    assert.deepEqual([s.orientation, s.width, s.height, s.top, s.left], [1, 297, 210, 20, 15]);
    assert.equal(s.header, '머리말 글');
    assert.deepEqual([...s.footerFields].sort(), [26, 33]);
    assert.match(formatWordInspect(info), /꼬리말: '\{page\} \/ \{pages\}'/);
    assert.ok(info.pages >= 2);
    const img = await pageImage(src, 1);
    assert.ok(isPng(img.png));
  });

  test('메모·변경 내용 추적·각주·링크·목차·필드', T, async () => {
    const src = fresh('c.docx');
    await wordEdit(src, [
      { op: 'insert_toc', at: 'start', levels: 2, title: '목차' },
      { op: 'add_comment', find: '비용 절감', text: '확인 필요', author: '검토자' },
      { op: 'add_comment', find: '매출', all: true, text: '매출 메모' },
      { op: 'delete_comments', index: 1 },
      { op: 'add_footnote', find: '2026년 3분기', text: '7~9월 기준' },
      { op: 'insert_hyperlink', find: '신규 고객', url: 'https://example.com/plan' },
      { op: 'track_changes', on: true },
      { op: 'insert_paragraph', text: '추적 문단' },
      { op: 'update_fields' },
    ]);
    let info = await wordInspect(src);
    assert.deepEqual([info.counts.toc, info.counts.footnotes, info.counts.hyperlinks, info.commentCount], [1, 1, 1, 3]);
    assert.ok(info.comments.some((c) => c.author === '검토자' && c.text === '확인 필요' && c.scope === '비용 절감'));
    assert.equal(info.trackRevisions, true);
    assert.ok(info.revisionCount >= 1);
    assert.deepEqual([info.paragraphs[0].text, info.paragraphs[0].key], ['목차', 'toc_heading']);
    assert.ok(info.paragraphs.some((p) => /^분기 보고\t\d+$/.test(p.text)), '목차 항목');

    await wordEdit(src, [{ op: 'accept_revisions' }, { op: 'track_changes', on: false }, { op: 'delete_comments', all: true }]);
    info = await wordInspect(src);
    assert.deepEqual([info.revisionCount, info.commentCount, info.trackRevisions], [0, 0, false]);
    assert.ok(para(info, '추적 문단'), '수락한 문단은 남는다');

    const r = await wordEdit(src, [{ op: 'track_changes', on: true }, { op: 'insert_paragraph', text: '거부될 문단' }, { op: 'reject_revisions' }]);
    assert.match(r.lines[2], /모두 거부/);
    info = await wordInspect(src);
    assert.equal(para(info, '거부될 문단'), undefined);
  });

  test('하나라도 실패하면 아무것도 저장하지 않는다 (바이트 같음)', T, async () => {
    const src = fresh('a.docx');
    const before = readFileSync(src);
    const cases = [
      [[{ op: 'insert_paragraph', text: '들어가면 안 됨' }, { op: 'apply_style', paragraph: 1, style: '없는 스타일' }], /2번째 작업\(apply_style\) 실패: 문서에 없는 스타일.*저장하지 않았습니다/],
      [[{ op: 'format_text', find: '매출', bold: true }, { op: 'table_set_cell', table: 1, row: 9, col: 1, text: 'x' }], /2번째 작업\(table_set_cell\) 실패: 9행 1열 칸이 없습니다/],
      [[{ op: 'set_paragraph_text', paragraph: 99, text: 'x' }], /문단 번호 99이\(가\) 범위를 벗어났습니다 \(전체 17개/],
      [[{ op: 'delete_paragraph', paragraph: 8 }], /행 끝 표시/],
      [[{ op: 'insert_hyperlink', find: '없는 글', url: 'https://a.b' }], /찾는 글자가 없습니다: '없는 글'/],
    ];
    for (const [ops, re] of cases) {
      await assert.rejects(wordEdit(src, ops), re);
      assert.ok(readFileSync(src).equals(before), `원본이 바뀌면 안 됨: ${ops.map((o) => o.op).join(',')}`);
    }
  });

  test('다른 Word 형식으로 저장 (rtf), 원본은 그대로', T, async () => {
    const src = fresh('f.docx');
    const out = tmp.path('f.rtf');
    const r = await wordEdit(src, [{ op: 'insert_paragraph', text: 'RTF 문단', style: 'List Bullet' }], { output: out });
    assert.match(r.lines[0], /17|18/);
    assert.ok(readFileSync(out).subarray(0, 5).toString() === '{\\rtf');
    assert.ok(readFileSync(src).equals(readFileSync(base)));
    const info = await wordInspect(out);
    assert.equal(info.paragraphs.at(-1).key, 'list_bullet');
  });

  test('fixture: 구역 나누기·병합된 칸이 살펴보기에 보인다', T, async () => {
    const info = await wordInspect(rich);
    assert.equal(info.paragraphTotal, 31);
    assert.deepEqual(info.sections.map((s) => s.orientation), [0, 1]);
    assert.equal(at(info, 29).sectionBreak, true);
    assert.equal(at(info, 29).text, '구역 끝 문단');
    assert.deepEqual([at(info, 18).merged, at(info, 18).table], [true, 2]);
    assert.equal(at(info, 3).text, '표 제목');
    assert.equal(at(info, 4).text, '가');
    const text = formatWordInspect(info);
    assert.match(text, /29 \[[^\]]*\] 구역 끝 문단 \[구역 나누기\]/);
    assert.match(text, /18 \[표2 병합된 칸/);
  });

  test('delete_paragraph: 구역 나누기를 품은 문단·표 바로 앞 문단도 문단째 지운다', T, async () => {
    const src = fresh('del.docx', rich);
    const r = await wordEdit(src, [
      { op: 'delete_paragraph', find: '구역 끝 문단' },
      { op: 'delete_paragraph', paragraph: 3 },
    ]);
    assert.match(r.lines[0], /^1\. 29번 문단 지움 \(29번의 구역 나누기는 앞 문단 끝에 남겨 쪽 설정 유지\) · 전체 30문단/);
    assert.match(r.lines[1], /^2\. 3번 문단 지움 · 전체 29문단/);
    const info = await wordInspect(src);
    assert.equal(info.paragraphTotal, 29);
    assert.deepEqual(info.sections.map((s) => s.orientation), [0, 1], '구역 2는 가로 그대로');
    assert.equal(at(info, 3).text, '가', '표 제목이 사라지고 표가 바로 옴');
    const end = para(info, '셋째');
    assert.equal(end.sectionBreak, true, '구역 나누기는 앞 문단으로');
    assert.equal(para(info, '구역 끝 문단'), undefined);
    const p2 = await pageImage(src, 2);
    assert.ok(p2.width > p2.height, '2쪽(부록)은 가로');
    const p1 = await pageImage(src, 1);
    assert.ok(p1.width < p1.height, '1쪽은 세로');
  });

  test('delete_paragraph: 지울 수 없는 자리는 이유를 알리고, 병합된 칸 아랫부분은 거절', T, async () => {
    const src = fresh('del2.docx', rich);
    const r = await wordEdit(src, [{ op: 'delete_paragraph', paragraph: 4 }]);
    assert.match(r.lines[0], /4번 문단은 지울 수 없는 자리\(표 칸의 유일한 문단\)라 글만 비움 \(번호 그대로\) · 전체 31문단/);
    const before = readFileSync(src);
    for (const op of [{ op: 'delete_paragraph', paragraph: 18 }, { op: 'set_paragraph_text', paragraph: 18, text: 'x' }, { op: 'insert_paragraph', text: 'x', at: { after_paragraph: 18 } }]) {
      await assert.rejects(wordEdit(src, [op]), /18번 문단은 세로로 병합된 칸의 아랫부분.*저장하지 않았습니다/);
    }
    assert.ok(readFileSync(src).equals(before));
  });

  test('세로 병합 표: 행 지우기·넣기가 그 행만 다루고 실제 번호를 알린다', T, async () => {
    const d2 = fresh('m-del2.docx', rich);
    let r = await wordEdit(d2, [{ op: 'table_delete_row', table: 2, row: 2 }]);
    assert.match(r.lines[0], /표 2 2행 지움 \(남은 3행\) — 세로로 병합된 칸의 글 '서울'은 남은 행에 둠/);
    let info = await wordInspect(d2);
    const cells = (i) => info.paragraphs.filter((p) => p.table === 2 && p.row === i).map((p) => p.text);
    assert.equal(info.tables[1].rows, 3);
    assert.deepEqual(cells(2), ['서울', '55', '65']);
    assert.deepEqual(cells(3), ['부산', '30', '35']);

    const d3 = fresh('m-del3.docx', rich);
    r = await wordEdit(d3, [{ op: 'table_delete_row', table: 2, row: 3 }]);
    assert.match(r.lines[0], /표 2 3행 지움 \(남은 3행\)/);
    info = await wordInspect(d3);
    assert.deepEqual(cells(2), ['서울', '50', '60']);
    assert.deepEqual(cells(3), ['부산', '30', '35']);

    const a = fresh('m-add.docx', rich);
    r = await wordEdit(a, [
      { op: 'table_add_row', table: 2, after_row: 2 },
      { op: 'table_add_row', table: 2, after_row: 4, values: [null, 'x', 'y'] },
      { op: 'table_add_row', table: 2, after_row: 0, values: ['머리', '반기'] },
    ]);
    assert.match(r.lines[0], /표 2에 행 추가: 새 3행 \(전체 5행\)/);
    assert.match(r.lines[1], /새 5행 \(전체 6행\)/);
    assert.match(r.lines[2], /새 1행 \(전체 7행\)/);
    info = await wordInspect(a);
    assert.deepEqual(cells(1), ['머리', '반기']);
    assert.deepEqual(cells(3), ['서울', '50', '60']);
    assert.deepEqual(cells(4), ['', ''], '병합 안에 든 새 행 (1열은 서울 칸)');
    assert.deepEqual(cells(5), ['55', '65']);
    assert.deepEqual(cells(6), ['', 'x', 'y']);
    assert.deepEqual(cells(7), ['부산', '30', '35']);

    const bad = fresh('m-bad.docx', rich);
    const before = readFileSync(bad);
    await assert.rejects(wordEdit(bad, [{ op: 'table_add_row', table: 2, after_row: 2, values: ['a', 'b', 'c'] }]), /새 3행 1열은 위 칸과 세로로 병합돼 값을 넣을 수 없습니다 — values의 그 자리는 null/);
    await assert.rejects(wordEdit(bad, [{ op: 'table_add_row', table: 2, after_row: 0, values: ['a', 'b', 'c'] }]), /값이 3개인데 새 행은 칸이 2개입니다/);
    assert.ok(readFileSync(bad).equals(before));
  });

  test('변경 내용 추적 중 지우기는 표시만 하고 번호를 그대로 둔다고 알린다', T, async () => {
    const src = fresh('track-del.docx', rich);
    const r = await wordEdit(src, [{ op: 'track_changes', on: true }, { op: 'delete_paragraph', paragraph: 2 }, { op: 'table_delete_row', table: 1, row: 2 }]);
    assert.match(r.lines[1], /2번 문단을 변경 내용으로 지움 표시 \(번호는 그대로 — 수락하면 사라짐\) · 전체 31문단/);
    assert.match(r.lines[2], /표 1 2행을 변경 내용으로 지움 표시/);
    const info = await wordInspect(src);
    assert.equal(at(info, 2).deleted, true);
    assert.match(formatWordInspect(info), /2 \[[^\]]*\] \[지움 표시됨 — 변경 내용 수락 전\] 본문 하나/);
  });

  test('find: 목차 안은 건너뛰고, all의 문단 작업은 문단마다 한 번', T, async () => {
    const src = fresh('find.docx', rich);
    const r = await wordEdit(src, [
      { op: 'insert_toc', at: 'start' },
      { op: 'insert_hyperlink', find: '표 제목', all: true, url: 'https://example.com/t' },
      { op: 'insert_paragraph', text: '개요 뒤', at: { after_text: '1. 개요' } },
      { op: 'format_paragraph', find: '매출', all: true, space_after: 3 },
      { op: 'apply_style', find: '매출', all: true, style: 'normal' },
      { op: 'format_text', find: '매출', all: true, bold: true, highlight: 'green' },
    ]);
    assert.match(r.lines[1], /'표 제목' 1곳/, '목차 항목은 빼고 본문 제목만');
    assert.match(r.lines[3], /문단 2개/);
    assert.match(r.lines[4], /포함 문단 2곳/);
    assert.match(r.lines[5], /'매출' 4곳/);
    const info = await wordInspect(src);
    const i = info.paragraphs.findIndex((p) => p.text === '개요 뒤');
    assert.equal(info.paragraphs[i - 1].key, 'heading1', '목차 항목이 아니라 본문 제목 뒤');
    const xml = await zipText(src, 'word/document.xml');
    assert.equal((xml.match(/<w:highlight w:val="green"\/>/g) ?? []).length >= 2, true);
    // 목차에만 있는 글(제목 + 탭 + 쪽 번호)은 따로 알린다
    await assert.rejects(wordEdit(src, [{ op: 'format_text', find: '개요\t', bold: true }]), /찾는 글자가 목차 안에만 있습니다: '개요/);
  });

  test('insert_toc: 제목 문단이 없으면 영어 안내 대신 한국어 오류', T, async () => {
    const plain = tmp.path('plain.docx');
    await createDocument(plain, '본문 하나\n본문 둘');
    const before = readFileSync(plain);
    await assert.rejects(wordEdit(plain, [{ op: 'insert_toc', at: 'start' }]), /목차에 넣을 제목 문단\(heading1~heading3 스타일\)이 없습니다/);
    assert.ok(readFileSync(plain).equals(before));
    const r = await wordEdit(plain, [{ op: 'apply_style', paragraph: 1, style: 'heading1' }, { op: 'insert_toc', at: 'start' }]);
    assert.match(r.lines[1], /목차 넣음/);
  });

  test('set_list: 직접 쓴 번호·기호가 빠지면 알리고, 한 목록으로 번호를 매긴다', T, async () => {
    const src = fresh('list.docx', rich);
    const r = await wordEdit(src, [{ op: 'set_list', paragraphs: [26, 28], type: 'number' }]);
    assert.match(r.lines[0], /26~28번 문단 번호 목록으로 \(앞에 직접 쓴 번호·기호는 자동 번호로 바뀌어 글에서 빠짐: 26번 '1\. ', 27번 '- '\)/);
    const info = await wordInspect(src);
    assert.deepEqual([26, 27, 28].map((i) => [at(info, i).list, at(info, i).text]), [['1.', '첫째'], ['2.', '둘째'], ['3.', '셋째']]);
  });

  test('format_text: 문단 범위면 목록 번호도 고르게, all은 많은 곳도 빠르게', T, async () => {
    const big = tmp.path('big.docx');
    await createDocument(big, Array.from({ length: 300 }, (_, i) => `${i + 1}번째 줄: 매출 지표와 비용 지표, 매출 지표`).join('\n'));
    const r = await wordEdit(big, [{ op: 'format_text', find: '매출 지표', all: true, bold: true, highlight: 'yellow' }]);
    assert.match(r.lines[0], /'매출 지표' 600곳/);
    assert.ok(r.results[0].ms < 90_000, `600곳 서식에 ${r.results[0].ms}ms`);
    const xml = await zipText(big, 'word/document.xml');
    assert.ok((xml.match(/<w:highlight w:val="yellow"\/>/g) ?? []).length >= 600);

    const src = fresh('fl.docx', rich);
    await wordEdit(src, [{ op: 'set_list', paragraphs: [26, 28], type: 'number' }, { op: 'format_text', paragraphs: [26, 28], color: '#c00000' }]);
    const doc = await zipText(src, 'word/document.xml');
    const marks = [...doc.matchAll(/<w:p\b[^>]*>(?:(?!<\/w:p>).)*?<w:numPr>(?:(?!<\/w:p>).)*?<\/w:p>/gs)].map((m) => /<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr>(?:(?!<\/w:rPr>).)*<w:color w:val="C00000"\/>/is.test(m[0]));
    assert.deepEqual(marks.slice(-3), [true, true, true], '세 문단의 문단 표시(번호 모양) 모두 같은 색');
  });

  test('읽기 공유로 열려 있는 문서에 그대로 저장하려 하면 바로 알린다 (저장 창에서 멈추지 않음)', T, async () => {
    const src = fresh('busy.docx', rich);
    const before = readFileSync(src);
    const h = holdOpen(src);
    try {
      await h.ready;
      const t0 = Date.now();
      await assert.rejects(wordEdit(src, [{ op: 'insert_paragraph', text: 'x' }]), /읽기 전용이거나 다른 프로그램\(Word 등\)에서 열려 있어/);
      assert.ok(Date.now() - t0 < 120_000);
      // 다른 경로로는 저장된다
      const out = tmp.path('busy-out.docx');
      const r = await wordEdit(src, [{ op: 'insert_paragraph', text: 'x' }], { output: out });
      assert.match(r.lines.at(-1), /busy-out\.docx/);
    } finally { h.close(); }
    assert.ok(readFileSync(src).equals(before));
  });

  test('MCP word_inspect / word_edit 실제 호출', T, async () => {
    const client = new Client({ name: 'word-com-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [SERVER], cwd: tmp.dir, stderr: 'pipe' }));
    try {
      const call = (name, args) => client.callTool({ name, arguments: args }, undefined, { timeout: 600_000 });
      const textOf = (r) => r.content.map((c) => c.text).join('\n');
      fresh('mcp.docx');
      const i = await call('word_inspect', { path: 'mcp.docx', limit: 5 });
      assert.ok(!i.isError, textOf(i));
      assert.match(textOf(i), /문단 1~5 \/ 전체 17개/);
      const e = await call('word_edit', { path: 'mcp.docx', output: 'mcp-out.docx', operations: [
        { op: 'insert_paragraph', text: 'MCP 문단', at: 'start', style: 'heading1' },
        { op: 'insert_image', image_path: 'blue.png', width_mm: 30 },
      ] });
      assert.ok(!e.isError, textOf(e));
      assert.match(textOf(e), /1\. 문단 넣음: 1번/);
      assert.ok(existsSync(tmp.path('mcp-out.docx')));
      const again = await call('word_edit', { path: 'mcp.docx', output: 'mcp-out.docx', operations: [{ op: 'update_fields' }] });
      assert.equal(again.isError, true);
      assert.match(textOf(again), /이미 있는 파일/);
      const bad = await call('word_edit', { path: 'mcp.docx', operations: [{ op: 'delete_paragraph', paragraph: 50 }] });
      assert.equal(bad.isError, true);
      assert.match(textOf(bad), /1번째 작업\(delete_paragraph\) 실패: 문단 번호 50/);
    } finally { await client.close(); }
  });
});
