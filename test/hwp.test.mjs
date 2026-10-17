// hwp.test.mjs — 한글 전용 기능 (hwp_inspect·hwp_edit 확장 작업) 테스트. 한컴 프로그램(COM) 없이 rhwp만 쓴다.
// 작업마다 .hwp와 .hwpx 두 형식으로 저장한 뒤 다시 열어 내용·서식을 확인하고, 쪽 그림(SVG/PNG)도 그려 본다.
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Resvg } from '@resvg/resvg-js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createDocument } from '../src/core/index.mjs';
import { Hwp, imageInfo } from '../src/core/hwp.mjs';
import { pageImage } from '../src/core/render.mjs';
import { applyHwpOp, inspectText } from '../src/mcp/tools/hwp.mjs';
import { MANUSCRIPT, ROOT, tempDir, isPng } from './helpers/fixtures.mjs';

const tmp = tempDir('officekit-hwp-');
after(() => tmp.cleanup());

const FORMATS = ['hwp', 'hwpx'];
const PNG = new Resvg('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200"><rect width="400" height="200" fill="#3366cc"/><circle cx="200" cy="100" r="60" fill="#ffcc00"/></svg>').render().asPng();
const JPG = Buffer.from('/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAEAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDF/tP/AIlX2D7DZf8AXx5X7772fvZ/D6UUUV87hvt/4mfoVLr6n//Z', 'base64');
const GIF = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
const BMP = Buffer.from('Qk22AAAAAAAAADYAAAAoAAAACAAAAAQAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAKCjI/ygoyP8oKMj/KCjI/ygoyP8oKMj/KCjI/ygoyP8oKMj/KCjI/wDX//8A1///ANf//wDX//8oKMj/KCjI/ygoyP8oKMj/ANf//wDX//8A1///ANf//ygoyP8oKMj/KCjI/ygoyP8oKMj/KCjI/ygoyP8oKMj/KCjI/ygoyP8=', 'base64');

// 원고 → 문단 1 '분기 보고', 2 '2026년 3분기', 3 '요약', 4 '• 매출 증가', 5 '• 비용 절감', 6 [표 1], 7 '담당자: [누름틀]', 8 '', 9 '다음 계획', 10 '신규 고객 확보'
const src = {};
before(async () => {
  for (const ext of FORMATS) { src[ext] = tmp.path(`원본.${ext}`); await createDocument(src[ext], MANUSCRIPT); }
  writeFileSync(tmp.path('그림.png'), PNG);
  writeFileSync(tmp.path('그림.jpg'), JPG);
  writeFileSync(tmp.path('그림.gif'), GIF);
  writeFileSync(tmp.path('그림.bmp'), BMP);
});

let seq = 0;
/** ops를 적용해 .hwp·.hwpx 둘 다로 저장 → { lines, outs: {hwp, hwpx} } */
function edit(from, ops) {
  const h = Hwp.open(from);
  const lines = [];
  try {
    for (const op of ops) lines.push(applyHwpOp(h, op, { resolvePath: (p) => join(tmp.dir, p) }));
    const base = tmp.path(`out-${++seq}`);
    const outs = {};
    for (const ext of FORMATS) outs[ext] = h.save(`${base}.${ext}`);
    return { lines, outs };
  } finally { h.close(); }
}

/** 다시 열어서 fn(h) */
function reopen(path, fn) {
  const h = Hwp.open(path);
  try { return fn(h); } finally { h.close(); }
}

const texts = (h) => h.paragraphs().map((p) => p.text);
// rhwp 원본 호출 결과 확인 (JSON 문자열, 실패면 ok:false)
const call0 = (r) => { const j = typeof r === 'string' ? JSON.parse(r) : r; assert.notEqual(j?.ok, false, JSON.stringify(j)); return j; };
const charAt = (h, n, offset = 0) => { const { sec, para } = h.locate(n); return JSON.parse(h.raw.getCharPropertiesAt(sec, para, offset)); };
const paraAt = (h, n) => { const { sec, para } = h.locate(n); return JSON.parse(h.raw.getParaPropertiesAt(sec, para)); };
const styleAt = (h, n) => { const { sec, para } = h.locate(n); return JSON.parse(h.raw.getStyleAt(sec, para)).name; };

/** 쪽 SVG의 글자 조각 [{x, y, t}] (화면용으로 열어서) */
function svgOf(path, page = 0) {
  const h = Hwp.openForView(path);
  try { return h.svg(page); } finally { h.close(); }
}
const svgTexts = (svg) => [...svg.matchAll(/<text x="([\d.]+)" y="([\d.]+)"[^>]*>([^<]*)<\/text>/g)].map((m) => ({ x: Number(m[1]), y: Number(m[2]), t: m[3] }));
const svgHeight = (svg) => Number(/height="([\d.]+)"/.exec(svg)?.[1] ?? 0);

/** 모든 출력을 쪽 그림(PNG)으로 그려 본다 */
async function renders(outs) {
  for (const ext of FORMATS) {
    const r = await pageImage(outs[ext], 1, { dpi: 48 });
    assert.ok(isPng(r.png), `${ext} 쪽 그림`);
  }
}

// ── 살펴보기 ─────────────────────────────────────────────

for (const ext of FORMATS) {
  describe(`hwp_inspect (.${ext})`, () => {
    test('쪽·구역·문단·표·누름틀·스타일', () => reopen(src[ext], (h) => {
      const r = h.inspect();
      assert.equal(r.format, ext);
      assert.equal(r.pages, 2);
      assert.equal(r.sections.length, 1);
      assert.deepEqual([r.sections[0].paper, r.sections[0].orientation, r.sections[0].widthMm, r.sections[0].heightMm], ['A4', 'portrait', 210, 297]);
      assert.equal(r.sections[0].margins.left, 30);
      assert.equal(r.paragraphCount, 10);
      assert.deepEqual(r.paragraphs.map((p) => p.n), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
      assert.equal(r.paragraphs[0].text, '분기 보고');
      assert.equal(r.paragraphs[0].align, '가운데');
      assert.equal(r.paragraphs[0].size, 18);
      assert.equal(r.paragraphs[0].bold, true);
      assert.deepEqual(r.paragraphs[5].objects, ['표 1(3행×2열)']);
      assert.deepEqual(r.tables, [{ n: 1, paragraph: 6, rows: 3, cols: 2, firstRow: ['항목', '값'] }]);
      assert.deepEqual(r.fields.map((f) => [f.name, f.paragraph]), [['담당자', 7]]);
      assert.ok(r.styles.includes('바탕글') && r.styles.includes('본문') && r.styles.includes('개요 1'));
      assert.deepEqual([r.headers, r.footers], [[], []]);
      const t = inspectText(r, { name: `원본.${ext}` });
      for (const s of ['[쪽 설정]', 'A4 세로 210×297mm', '[문단] 1~10 / 전체 10개', '  1. 분기 보고', '표 1 — 문단 6, 3행×2열, 첫 행: 항목 | 값', '담당자 = (빈 칸) — 문단 7', '[스타일] 바탕글']) {
        assert.ok(t.includes(s), `inspect 글에 '${s}' 없음:\n${t}`);
      }
    }));

    test('표·누름틀 목록은 보여 준 문단 범위 안 것만 (전체 개수는 머리 줄에)', () => reopen(src[ext], (h) => {
      const r = h.inspect({ offset: 0, limit: 5 });
      assert.equal(r.tables.length, 1, '원자료는 전체 목록');
      assert.equal(r.tables[0].firstRow, undefined, '범위 밖 표는 첫 행을 읽지 않음');
      const t = inspectText(r);
      assert.match(t, /표 1개 · 누름틀 1개/);
      assert.match(t, /\[표\] 문단 1~5 안 0개 \/ 전체 1개/);
      assert.match(t, /뒤쪽 표 1개: 표 1\(문단 6\) — 보려면 offset=\(문단 번호-1\)/);
      assert.match(t, /\[누름틀\] 문단 1~5 안 0개 \/ 전체 1개/);
      assert.ok(!t.includes('담당자 ='), t);
      const t2 = inspectText(h.inspect({ offset: 6, limit: 4 }));
      assert.match(t2, /앞쪽 표 1개: 표 1~1 \(문단 6~6\)/);
      assert.match(t2, /담당자 = \(빈 칸\) — 문단 7/);
    }));

    test('offset/limit로 나눠 보기', () => reopen(src[ext], (h) => {
      const r = h.inspect({ offset: 8, limit: 5 });
      assert.deepEqual(r.paragraphs.map((p) => p.n), [9, 10]);
      const r2 = h.inspect({ offset: 0, limit: 3 });
      assert.deepEqual(r2.paragraphs.map((p) => p.n), [1, 2, 3]);
      assert.match(inspectText(r2), /다음 문단은 offset=3/);
    }));
  });
}

test('hwp_inspect 글: 큰 서식 문서도 표 50개·누름틀 100개까지만, 긴 첫 행·값은 자름', () => {
  const tables = Array.from({ length: 70 }, (_, i) => ({ n: i + 1, paragraph: i + 1, rows: 2, cols: 14, firstRow: Array.from({ length: 10 }, () => '아주 긴 칸 글'.repeat(10)) }));
  const fields = Array.from({ length: 300 }, (_, i) => ({ name: `칸${i + 1}`, value: '값'.repeat(300), cell: true, paragraph: 1 + (i % 70) }));
  const r = {
    format: 'hwp', pages: 9, sections: [], headers: [], footers: [], styles: ['바탕글'], counts: { pictures: 0, footnotes: 0 },
    paragraphCount: 500, offset: 0, limit: 150, paragraphs: [], tables: [...tables, { n: 71, paragraph: 400, rows: 1, cols: 1 }], fields,
  };
  const t = inspectText(r);
  assert.match(t, /표 71개 · 누름틀 300개/);
  assert.match(t, /\[표\] 이 범위 안 70개 \/ 전체 71개/);
  assert.equal((t.match(/^ {2}표 \d+ — /gm) ?? []).length, 50);
  assert.match(t, /… 외 20개 — limit을 줄이거나 offset을 옮겨 보세요/);
  assert.match(t, /… \(\+4칸\)/);
  assert.match(t, /뒤쪽 표 1개: 표 71\(문단 400\)/);
  assert.equal((t.match(/^ {2}칸\d+ = /gm) ?? []).length, 100);
  assert.match(t, /… 외 200개 — 이름 전체는 office_fields로 보세요/);
  // 문서가 더 커져도 길이가 늘지 않음 (표 50·누름틀 100 상한)
  assert.ok(t.length < 25000, `너무 김: ${t.length}자`);
  const bigger = inspectText({ ...r, tables: [...tables, ...tables], fields: [...fields, ...fields] });
  assert.ok(Math.abs(bigger.length - t.length) < 50, `${bigger.length} vs ${t.length}`);
});

// ── 편집 작업: 두 원본 형식 × 두 저장 형식 ─────────────────

for (const ext of FORMATS) {
  describe(`hwp_edit 문단 (원본 .${ext})`, () => {
    test('insert_paragraph: start / after / before / after_text / end, 스타일·본보기·서식·여러 줄', async () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'set_header', text: '머리말 유지 확인' },
        { op: 'insert_paragraph', text: '맨 앞', at: 'start', style: '본문' },
        { op: 'insert_paragraph', text: '첫째 줄\n둘째 줄', at: { after_paragraph: 4 }, format: { bold: true, color: '#0055aa' }, align: 'center' },
        { op: 'insert_paragraph', text: '표 뒤', at: { after_text: '120' } },
        { op: 'insert_paragraph', text: '제목처럼', at: { before_paragraph: 3 }, like_paragraph: 2 },
        { op: 'insert_paragraph', text: '맨 끝' },
      ]);
      assert.match(lines[1], /문단 1: 맨 앞/);
      assert.match(lines[2], /문단 5~6: .*\+2/);
      assert.match(lines[3], /문단 10: 표 뒤/);
      assert.match(lines[4], /문단 3: 제목처럼/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.deepEqual(texts(h), ['맨 앞', '분기 보고', '제목처럼', '2026년 3분기', '요약', '첫째 줄', '둘째 줄', '• 매출 증가', '• 비용 절감', '',
            '표 뒤', '담당자: ', '', '다음 계획', '신규 고객 확보', '맨 끝'], e);
          assert.equal(styleAt(h, 1), '본문');
          assert.equal(styleAt(h, 16), '바탕글');
          for (const n of [6, 7]) {
            assert.equal(charAt(h, n).bold, true);
            assert.equal(charAt(h, n).textColor, '#0055aa');
            assert.equal(paraAt(h, n).alignment, 'center');
          }
          // like_paragraph 2(제목) → 18pt 굵게 가운데
          assert.equal(charAt(h, 3).fontSize, 1800);
          assert.equal(charAt(h, 3).bold, true);
          assert.equal(paraAt(h, 3).alignment, 'center');
          // 기본은 바탕글 모양 (제목 뒤에 넣어도 제목 모양을 물려받지 않음)
          assert.equal(charAt(h, 11).bold, false);
          // 첫 문단 앞에 넣어도 구역 정의(쪽 설정)·머리말이 그대로
          assert.equal(h.pageSetups()[0].paper, 'A4');
          assert.equal(h.headerFooter('header')[0]?.text, '머리말 유지 확인');
          assert.equal(h.tables()[0].paragraph, 10);
          assert.equal(h.fields()[0].name, '담당자');
        });
      }
      await renders(outs);
    });

    test('set_paragraph_text: 모양 유지, 여러 줄, 개체가 든 문단은 거부', () => {
      const { outs } = edit(src[ext], [
        { op: 'set_paragraph_text', paragraph: 1, text: '새 제목' },
        { op: 'set_paragraph_text', find: '신규 고객', text: '첫 줄\n둘째 줄' },
      ]);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.equal(texts(h)[0], '새 제목');
          assert.equal(charAt(h, 1).fontSize, 1800);
          assert.equal(paraAt(h, 1).alignment, 'center');
          assert.deepEqual(texts(h).slice(9), ['첫 줄', '둘째 줄']);
          assert.equal(charAt(h, 11).bold, charAt(h, 10).bold);
        });
      }
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'set_paragraph_text', paragraph: 7, text: 'x' }), /누름틀.*replace.*fill/);
        assert.throws(() => applyHwpOp(h, { op: 'set_paragraph_text', paragraph: 6, text: 'x' }), /표.*table_set_cell/);
        assert.throws(() => applyHwpOp(h, { op: 'set_paragraph_text', paragraph: 11, text: 'x' }), /문단 번호가 범위를 벗어났습니다: 11 \(전체 10개/);
        assert.throws(() => applyHwpOp(h, { op: 'set_paragraph_text', text: 'x' }), /대상 문단을 paragraph, find 중 하나로만/);
        assert.throws(() => applyHwpOp(h, { op: 'delete_paragraph', paragraph: 1, find: '요약' }), /paragraph, find 중 하나로만 지정하세요 \(함께 줌: paragraph, find\)/);
        assert.throws(() => applyHwpOp(h, { op: 'format_paragraph', align: 'left' }), /paragraph, paragraphs:\[처음,끝\], find 중 하나로만/);
      } finally { h.close(); }
    });

    test('delete_paragraph: 번호·찾기, 첫 문단(머리말 유지), 표가 든 문단', async () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'set_footer', text: '', page_numbers: true },
        { op: 'delete_paragraph', paragraph: 1 },
        { op: 'delete_paragraph', find: '비용 절감' },
        { op: 'delete_paragraph', paragraph: 4 },
      ]);
      assert.equal(lines[1], '문단 1 지움 — 뒤 문단 번호 -1');
      // 표가 하나뿐이라 뒤 표 번호 안내는 없음
      assert.equal(lines[3], '문단 4 지움 (표 1도 함께 지움) — 뒤 문단 번호 -1');
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.deepEqual(texts(h), ['2026년 3분기', '요약', '• 매출 증가', '담당자: ', '', '다음 계획', '신규 고객 확보']);
          assert.equal(h.tables().length, 0);
          assert.equal(h.headerFooter('footer').length, 1, '꼬리말이 첫 문단과 함께 사라짐');
          assert.equal(h.pageSetups()[0].margins.left, 30);
          assert.equal(paraAt(h, 1).alignment, 'center');
        });
      }
      await renders(outs);
      const h = Hwp.create();
      try {
        h.addParagraph('하나뿐');
        assert.throws(() => h.deleteParagraph(1), /하나뿐인 문단은 지울 수 없습니다/);
      } finally { h.close(); }
    });

    test('format_text: 첫 번째·전부, 본문과 표 칸, 글꼴', () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'format_text', find: '매출', all: true, bold: true, italic: true, underline: true, size: 14, color: '#cc0000', font: '맑은 고딕' },
        { op: 'format_text', find: '비용', strike: true, highlight: '#ffff00' },
      ]);
      assert.match(lines[0], /'매출' 2곳/);
      assert.match(lines[1], /'비용' 1곳/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          const c = charAt(h, 4, 2); // '• 매출 증가'의 '매'
          assert.deepEqual([c.bold, c.italic, c.underline, c.fontSize, c.textColor, c.fontFamily], [true, true, true, 1400, '#cc0000', '맑은 고딕'], e);
          assert.equal(charAt(h, 4, 5).bold, false, '찾은 글 뒤는 그대로');
          const t = h.tables()[0];
          const cell = JSON.parse(h.raw.getCellCharPropertiesAt(t.section, t.paraIndex, t.controlIndex, 2, 0, 0)); // 2행 1열 '매출'
          assert.deepEqual([cell.bold, cell.textColor], [true, '#cc0000']);
          const s = charAt(h, 5, 2); // '• 비용 절감'
          assert.deepEqual([s.strikethrough, s.shadeColor], [true, '#ffff00']);
          const other = JSON.parse(h.raw.getCellCharPropertiesAt(t.section, t.paraIndex, t.controlIndex, 4, 0, 0)); // 표 안 '비용'은 첫 번째가 아님
          assert.equal(other.strikethrough, false);
        });
      }
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'format_text', find: '없는 글', bold: true }), /찾지 못했습니다/);
        assert.throws(() => applyHwpOp(h, { op: 'format_text', find: '매출' }), /바꿀 글자 서식이 없습니다/);
        assert.throws(() => applyHwpOp(h, { op: 'format_text', find: '매출', color: 'red' }), /#rrggbb/);
      } finally { h.close(); }
    });

    test('format_paragraph: 정렬·줄 간격·여백(mm)·간격·범위', () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'format_paragraph', paragraphs: [4, 5], align: 'right', line_spacing: 2, left_indent_mm: 10, first_line_indent_mm: 5 },
        { op: 'format_paragraph', find: '신규 고객', line_spacing: '130%', space_before: 6, page_break_before: true },
        { op: 'format_paragraph', paragraph: 3, first_line_indent_mm: -8, space_after: 4, keep_with_next: true },
      ]);
      // 결과 줄은 한국어 이름으로
      assert.equal(lines[0], '문단 서식 문단 4~5: 정렬 오른쪽, 줄 간격 200%, 왼쪽 여백 10mm, 첫 줄 들여쓰기 5mm');
      assert.equal(lines[1], '문단 서식 문단 10: 줄 간격 130%, 문단 위 6pt, 앞에서 쪽 나눔');
      assert.equal(lines[2], '문단 서식 문단 3: 내어쓰기 8mm, 문단 아래 4pt, 다음 문단과 함께');
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          for (const n of [4, 5]) {
            const p = paraAt(h, n);
            assert.equal(p.alignment, 'right');
            assert.equal(p.lineSpacing, 200);
            assert.ok(Math.abs(p.marginLeft - 37.8) < 0.2, `왼쪽 여백 10mm = 37.8px: ${p.marginLeft}`);
            assert.ok(Math.abs(p.indent - 18.9) < 0.2, `들여쓰기 5mm: ${p.indent}`);
          }
          assert.equal(paraAt(h, 3).alignment, 'justify', '범위 밖은 그대로');
          assert.equal(paraAt(h, 10).lineSpacing, 130);
          assert.equal(paraAt(h, 10).pageBreakBefore, true);
          assert.equal(h.pageCount, 3);
        });
      }
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'format_paragraph', paragraph: 1 }), /바꿀 문단 서식이 없습니다/);
        assert.throws(() => applyHwpOp(h, { op: 'format_paragraph', paragraphs: [3, 2], align: 'left' }), /거꾸로/);
        assert.throws(() => applyHwpOp(h, { op: 'format_paragraph', paragraph: 1, line_spacing: 9 }), /줄 간격/);
      } finally { h.close(); }
    });

    test('apply_style: 이름·영문 이름, 문단·글자 모양이 스타일 것으로', async () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'apply_style', find: '신규 고객', style: '개요 1' },
        { op: 'apply_style', paragraphs: [4, 5], style: 'Body' },
      ]);
      assert.match(lines[0], /'개요 1' 적용: 문단 10/);
      assert.match(lines[1], /'본문' 적용: 문단 4~5/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.equal(styleAt(h, 10), '개요 1');
          assert.equal(charAt(h, 10).bold, false, '스타일 글자 모양(굵게 아님)');
          assert.equal(paraAt(h, 10).headType, 'Outline');
          assert.deepEqual([styleAt(h, 4), styleAt(h, 5)], ['본문', '본문']);
          assert.ok(Math.abs(paraAt(h, 4).marginLeft - 20) < 0.2, '본문 스타일 왼쪽 여백 15pt');
        });
      }
      await renders(outs);
      const h = Hwp.open(src[ext]);
      try { assert.throws(() => applyHwpOp(h, { op: 'apply_style', paragraph: 1, style: '없는 스타일' }), /문서에 없습니다 \(있는 스타일: 바탕글, 본문/); } finally { h.close(); }
    });
  });

  describe(`hwp_edit 표 (원본 .${ext})`, () => {
    test('table_set_cell / table_add_row / table_delete_row', async () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'table_set_cell', table: 1, row: 2, col: 2, text: '150' },
        { op: 'table_add_row', table: 1, values: ['이익', 70] },
        { op: 'table_add_row', table: 1, after_row: 1, values: ['첫 행 아래'] },
        { op: 'table_add_row', table: 1, after_row: 0, values: ['맨 위', '0'] },
        { op: 'table_delete_row', table: 1, row: 4 },
        { op: 'table_set_cell', table: 1, row: 1, col: 2, text: '두 줄\n칸' },
      ]);
      assert.match(lines[1], /→ 4행: 이익 \| 70/);
      assert.match(lines[2], /→ 2행/);
      assert.match(lines[3], /→ 1행/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          const t = h.tables()[0];
          assert.deepEqual([t.rows, t.cols], [5, 2]);
          assert.deepEqual(h.readTable(t), [['맨 위', '두 줄\n칸'], ['항목', '값'], ['첫 행 아래', ''], ['비용', '80'], ['이익', '70']], e);
        });
      }
      await renders(outs);
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'table_set_cell', table: 2, row: 1, col: 1, text: 'x' }), /표 번호가 범위를 벗어났습니다: 2 \(표 1개/);
        assert.throws(() => applyHwpOp(h, { op: 'table_set_cell', table: 1, row: 4, col: 1, text: 'x' }), /행 번호가 범위를 벗어났습니다: 4 \(전체 3행\)/);
        assert.throws(() => applyHwpOp(h, { op: 'table_add_row', table: 1, values: ['a', 'b', 'c'] }), /2열/);
      } finally { h.close(); }
    });

    test('병합 칸: 읽기·쓰기는 칸이 덮는 자리 기준', () => {
      const h = Hwp.open(src[ext]);
      try {
        const t = h.tables()[0];
        h.raw.mergeTableCells(t.section, t.paraIndex, t.controlIndex, 1, 0, 2, 0); // '매출'·'비용' 세로 병합
        assert.deepEqual(h.readTable(t), [['항목', '값'], ['매출\n비용', '120'], ['', '80']]);
        h.setCell(1, 3, 1, '합침'); // 덮인 자리 → 병합 칸
        h.setCell(1, 3, 2, '90');
        assert.deepEqual(h.readTable(t), [['항목', '값'], ['합침', '120'], ['', '90']]);
        for (const e of FORMATS) {
          const out = h.save(tmp.path(`merged-${ext}.${e}`));
          reopen(out, (r) => assert.deepEqual(r.readTable(r.tables()[0]), [['항목', '값'], ['합침', '120'], ['', '90']]));
        }
      } finally { h.close(); }
    });

    test('table_add_row: 세로 병합 칸 사이에 넣으면 그 열 값은 거절(표 그대로), 빈 값이면 병합 안내', async () => {
      // 2~3행 2열('120'·'80')을 세로 병합한 표
      const h = Hwp.open(src[ext]);
      let merged;
      try {
        const t = h.tables()[0];
        h.raw.mergeTableCells(t.section, t.paraIndex, t.controlIndex, 1, 1, 2, 1);
        merged = h.save(tmp.path(`vmerge-${ext}.${ext}`));
      } finally { h.close(); }
      const r = Hwp.open(merged);
      try {
        const before = r.readTable(r.tables()[0]);
        assert.throws(() => applyHwpOp(r, { op: 'table_add_row', table: 1, after_row: 2, values: ['이익', '40'] }),
          /표 1의 새 3행에서 2열은 위 칸과 세로로 병합된 자리라 값을 넣을 수 없습니다: '40'/);
        assert.deepEqual(r.readTable(r.tables()[0]), before, '거절하면 표를 바꾸지 않음');
      } finally { r.close(); }
      const { lines, outs } = edit(merged, [
        { op: 'table_add_row', table: 1, after_row: 2, values: ['이익', ''] },
        { op: 'table_add_row', table: 1, after_row: 1, values: ['병합 위', '5'] },
        { op: 'table_add_row', table: 1, values: ['끝', '9'] },
      ]);
      assert.equal(lines[0], '표 1에 행 추가 → 3행: 이익 | (병합) (2열은 위 칸과 병합된 자리)');
      assert.equal(lines[1], '표 1에 행 추가 → 2행: 병합 위 | 5');
      for (const e of FORMATS) {
        reopen(outs[e], (h2) => {
          assert.deepEqual(h2.readTable(h2.tables()[0]), [['항목', '값'], ['병합 위', '5'], ['매출', '120\n80'], ['이익', ''], ['비용', ''], ['끝', '9']], e);
        });
        const page = svgTexts(svgOf(outs[e], 0)).map((x) => x.t).join('');
        assert.match(page, /병합위5매출120.*이익.*비용끝9/, `${e} 그려진 표`);
      }
      await renders(outs);
    });

    test('표가 맨 앞인 문서: 앞에 넣기는 거부, 표 문단 지우기는 머리말 유지', () => {
      const h = Hwp.create();
      try {
        h.addTable([['가', '나'], ['다', '라']]);
        h.addParagraph('본문');
        h.setHeaderFooter('header', '머리');
        assert.equal(h.tables()[0].paragraph, 1);
        assert.throws(() => h.insertParagraph('x', { at: 'start' }), /첫 문단에 표이\(가\) 있어/);
        assert.match(applyHwpOp(h, { op: 'delete_paragraph', paragraph: 1 }), /표 1\(2행×2열\)|표/);
        assert.equal(h.tables().length, 0);
        for (const e of FORMATS) {
          reopen(h.save(tmp.path(`tablefirst-${ext}.${e}`)), (r) => {
            assert.deepEqual(texts(r).filter(Boolean), ['본문']);
            assert.equal(r.headerFooter('header')[0]?.text, '머리');
          });
        }
      } finally { h.close(); }
    });
  });

  describe(`hwp_edit 머리말·각주·그림·쪽 (원본 .${ext})`, () => {
    test('set_header / set_footer: 글·정렬·쪽 번호, 다시 정하기, 지우기', async () => {
      const { outs } = edit(src[ext], [
        { op: 'set_header', text: '임시 머리말' },
        { op: 'set_header', text: '분기 보고서 — 대외비', align: 'right' },
        { op: 'set_footer', text: '쪽 {page} / {pages}' },
      ]);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.deepEqual(h.headerFooter('header').map((x) => x.text), ['분기 보고서 — 대외비']);
          assert.equal(h.headerFooter('footer').length, 1);
          const { sec } = h.locate(1);
          assert.equal(JSON.parse(h.raw.getParaPropertiesInHf(sec, true, 0, 0)).alignment, 'right');
        });
        // 2쪽 아래쪽에 '쪽 2 / 2'
        const svg = svgOf(outs[e], 1);
        const bottom = svgTexts(svg).filter((x) => x.y > svgHeight(svg) * 0.85).map((x) => x.t).join('');
        assert.match(bottom, /쪽2\/2/, `${e} 꼬리말 쪽 번호: ${bottom}`);
        assert.ok(svgTexts(svgOf(outs[e], 0)).some((x) => x.t === '대'), '머리말이 그려짐');
      }
      await renders(outs);
      const { outs: removed } = edit(outs.hwp, [{ op: 'set_header', text: '' }, { op: 'set_footer', text: '', page_numbers: true }]);
      for (const e of FORMATS) {
        reopen(removed[e], (h) => {
          assert.equal(h.headerFooter('header').length, 0);
          assert.equal(h.headerFooter('footer').length, 1);
          assert.equal(JSON.parse(h.raw.getParaPropertiesInHf(0, false, 0, 0)).alignment, 'center', '쪽 번호만 있으면 가운데');
        });
        const svg = svgOf(removed[e], 0);
        assert.match(svgTexts(svg).filter((x) => x.y > svgHeight(svg) * 0.85).map((x) => x.t).join(''), /-1-/);
      }
    });

    test('add_footnote: 찾은 글 뒤에 각주', () => {
      const { lines, outs } = edit(src[ext], [{ op: 'add_footnote', find: '비용 절감', text: '인건비 포함' }]);
      assert.match(lines[0], /각주 1\) '비용 절감' 뒤 \(문단 5\)/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.equal(h.inspect().counts.footnotes, 1);
          assert.deepEqual(h.inspect().paragraphs[4].objects, ['각주']);
          const info = JSON.parse(h.raw.getFootnoteInfo(0, 4, 0));
          assert.match(info.texts.join(''), /인건비 포함/);
        });
        const page = svgTexts(svgOf(outs[e], 0)).map((x) => x.t).join('');
        assert.match(page, /1\)인건비포함/);
      }
      const h = Hwp.open(src[ext]);
      try { assert.throws(() => applyHwpOp(h, { op: 'add_footnote', find: '120', text: 'x' }), /표 안에만 있습니다/); } finally { h.close(); }
    });

    test('insert_image: png·jpg·gif·bmp, 크기·정렬, 본문보다 넓으면 줄임', async () => {
      const { lines, outs } = edit(src[ext], [
        { op: 'insert_image', image_path: '그림.png', at: { after_paragraph: 3 }, width_mm: 60 },
        { op: 'insert_image', image_path: '그림.jpg', at: { after_text: '비용 절감' }, height_mm: 10, align: 'left' },
        { op: 'insert_image', image_path: '그림.gif', at: 'end', width_mm: 5 },
        { op: 'insert_image', image_path: '그림.bmp', at: { before_paragraph: 1 }, width_mm: 20, align: 'right' },
        { op: 'insert_image', image_path: '그림.png', width_mm: 400 },
      ]);
      assert.match(lines[0], /문단 4: 그림\.png 60×30mm/);
      assert.match(lines[1], /20×10mm/);
      assert.match(lines[4], /150×75mm \(본문 너비에 맞춰 줄임\)/);
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          const r = h.inspect();
          assert.equal(r.counts.pictures, 5);
          assert.deepEqual(r.paragraphs[0].objects, ['그림']);
          assert.equal(r.paragraphs[0].align, '오른쪽');
          assert.equal(r.paragraphs[4].align, '가운데');
          assert.equal(texts(h)[1], '분기 보고');
          const pic = JSON.parse(h.raw.getPictureProperties(0, 4, 0));
          assert.equal(pic.treatAsChar, true, '글자처럼 취급');
          assert.ok(Math.abs(pic.width / (7200 / 25.4) - 60) < 0.5);
        });
        assert.equal((svgOf(outs[e], 0).match(/<image/g) ?? []).length, 3, `${e} 1쪽 그림 3개가 그려짐`);
      }
      await renders(outs);
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'insert_image', image_path: '없는.png' }, { resolvePath: (p) => join(tmp.dir, p) }), /그림 파일이 없습니다/);
        writeFileSync(tmp.path('가짜.png'), 'not an image');
        assert.throws(() => applyHwpOp(h, { op: 'insert_image', image_path: '가짜.png' }, { resolvePath: (p) => join(tmp.dir, p) }), /그림 형식을 알 수 없습니다/);
      } finally { h.close(); }
    });

    test('insert_image: 크기를 안 준 긴 그림은 본문 높이에 맞춰 줄여 한 쪽 안에 그려짐', async () => {
      writeFileSync(tmp.path('긴그림.png'), new Resvg('<svg xmlns="http://www.w3.org/2000/svg" width="200" height="4000"><rect width="200" height="4000" fill="#264653"/><rect y="1900" width="200" height="200" fill="#e9c46a"/></svg>').render().asPng());
      const { lines, outs } = edit(src[ext], [
        { op: 'insert_image', image_path: '긴그림.png', at: { after_paragraph: 5 } },
        { op: 'insert_image', image_path: '긴그림.png', at: 'end', height_mm: 400 },
      ]);
      // A4 297mm - 위 20 - 아래 15 - 머리말 15 - 꼬리말 15 = 232mm, 줄 여유 2mm
      assert.match(lines[0], /긴그림\.png 11\.5×230mm \(본문 높이에 맞춰 줄임\)/);
      assert.match(lines[1], /11\.5×230mm \(본문 높이에 맞춰 줄임\)/);
      for (const e of FORMATS) {
        const h = Hwp.openForView(outs[e]);
        try {
          let seen = 0;
          for (let p = 0; p < h.pageCount; p++) {
            const svg = h.svg(p);
            for (const m of svg.matchAll(/<image x="[\d.]+" y="([\d.]+)" width="[\d.]+" height="([\d.]+)"/g)) {
              seen++;
              assert.ok(Number(m[1]) + Number(m[2]) <= svgHeight(svg), `${e} ${p + 1}쪽 그림이 쪽 아래로 넘침`);
            }
          }
          assert.equal(seen, 2, `${e} 그림 2개가 그려짐`);
        } finally { h.close(); }
      }
      await renders(outs);
    });

    test('page_setup: 용지·방향·여백(mm)', async () => {
      const { lines, outs } = edit(src[ext], [{ op: 'page_setup', paper: 'A3', orientation: 'landscape', margins_mm: { top: 10, bottom: 12, left: 15, right: 16, header: 8, footer: 9 } }]);
      assert.equal(lines[0], '쪽 설정: A3, 가로, 여백(mm) 위 10, 아래 12, 왼쪽 15, 오른쪽 16, 머리말 8, 꼬리말 9');
      for (const e of FORMATS) {
        reopen(outs[e], (h) => {
          assert.deepEqual(h.pageSetups()[0], { section: 0, paper: 'A3', orientation: 'landscape', widthMm: 420, heightMm: 297,
            margins: { top: 10, bottom: 12, left: 15, right: 16, header: 8, footer: 9, gutter: 0 } });
        });
        const svg = svgOf(outs[e], 0);
        assert.ok(Number(/width="([\d.]+)"/.exec(svg)[1]) > svgHeight(svg), '가로 쪽');
      }
      await renders(outs);
      const h = Hwp.open(src[ext]);
      try {
        assert.throws(() => applyHwpOp(h, { op: 'page_setup' }), /바꿀 쪽 설정이 없습니다/);
        assert.throws(() => applyHwpOp(h, { op: 'page_setup', margins_mm: { left: 120, right: 120 } }), /본문 자리가 남지 않습니다/);
        assert.throws(() => applyHwpOp(h, { op: 'page_setup', paper: 'A4', section: 2 }), /구역 번호가 범위를 벗어났습니다/);
      } finally { h.close(); }
    });
  });
}

// ── 구역이 둘인 문서: 번호는 구역을 이어서 센다 ─────────────

describe('여러 구역', () => {
  let two;
  before(async () => {
    const z = await JSZip.loadAsync(readFileSync(src.hwpx));
    const s0 = await z.file('Contents/section0.xml').async('string');
    z.file('Contents/section1.xml', s0.replace('분기 보고', '둘째 구역'));
    const hpf = await z.file('Contents/content.hpf').async('string');
    z.file('Contents/content.hpf', hpf
      .replace('<opf:item id="settings"', '<opf:item id="section1" href="Contents/section1.xml" media-type="application/xml"/><opf:item id="settings"')
      .replace('</opf:spine>', '<opf:itemref idref="section1" linear="yes"/></opf:spine>'));
    two = tmp.path('두구역.hwpx');
    writeFileSync(two, await z.generateAsync({ type: 'nodebuffer' }));
  });

  test('번호·표·누름틀이 구역을 이어 센다', () => reopen(two, (h) => {
    assert.equal(h.raw.getSectionCount(), 2);
    const r = h.inspect();
    assert.equal(r.paragraphCount, 20);
    assert.equal(r.paragraphs[10].text, '둘째 구역');
    assert.equal(r.paragraphs[10].section, 2);
    assert.deepEqual(h.locate(11), { sec: 1, para: 0 });
    assert.equal(h.numberOf(1, 5), 16);
    assert.deepEqual(r.tables.map((t) => t.paragraph), [6, 16]);
    assert.deepEqual(r.fields.map((f) => f.paragraph), [7, 17]);
    assert.match(inspectText(r), /── 구역 2 ──\n {2}11\. 둘째 구역/);
  }));

  test('둘째 구역 첫 문단 앞에 넣기·지우기, 둘째 표, 구역 하나만 쪽 설정', () => {
    const { outs } = edit(two, [
      { op: 'insert_paragraph', text: '구역2 맨 앞', at: { before_paragraph: 11 } },
      { op: 'delete_paragraph', paragraph: 13 },
      { op: 'table_set_cell', table: 2, row: 1, col: 1, text: '둘째 표' },
      { op: 'page_setup', section: 2, orientation: 'landscape' },
      { op: 'set_footer', text: '{page}', page_numbers: true },
    ]);
    for (const e of FORMATS) {
      reopen(outs[e], (h) => {
        assert.equal(h.raw.getSectionCount(), 2);
        assert.deepEqual(texts(h).slice(10, 13), ['구역2 맨 앞', '둘째 구역', '요약']);
        assert.equal(h.readTable(h.tables()[1])[0][0], '둘째 표');
        assert.equal(h.readTable(h.tables()[0])[0][0], '항목');
        assert.deepEqual(h.pageSetups().map((s) => s.orientation), ['portrait', 'landscape']);
        assert.equal(h.headerFooter('footer').length, 2, '모든 구역에 꼬리말');
      });
    }
  });
});

// ── 기존 작업과 섞어 쓰기 ─────────────────────────────────

describe('기존 작업과 함께', () => {
  test('빈 새 문서: insert_paragraph 뒤 add_paragraph가 앞 글을 덮지 않음, add_paragraph 글꼴', () => {
    const h = Hwp.create();
    try {
      applyHwpOp(h, { op: 'insert_paragraph', text: '첫 문단' });
      applyHwpOp(h, { op: 'add_paragraph', text: '둘째', style: { font: '맑은 고딕', size: 12 } });
      applyHwpOp(h, { op: 'add_table', rows: [['a', 'b']] });
      applyHwpOp(h, { op: 'insert_paragraph', text: '표 다음', at: { after_paragraph: 3 } });
      assert.deepEqual(texts(h).filter(Boolean), ['첫 문단', '둘째', '표 다음']);
      assert.equal(charAt(h, 2).fontFamily, '맑은 고딕');
      assert.equal(charAt(h, 2).fontSize, 1200);
    } finally { h.close(); }
  });

  test('끝 빈 문단 재사용: 다른 위치에 넣어도 add_table 뒤 add_paragraph 사이에 빈 문단이 남지 않음', () => {
    const h = Hwp.create();
    try {
      applyHwpOp(h, { op: 'add_paragraph', text: 'A' });
      applyHwpOp(h, { op: 'add_table', rows: [['a', 'b']] });
      applyHwpOp(h, { op: 'insert_paragraph', text: 'B', at: { after_paragraph: 1 } });
      applyHwpOp(h, { op: 'set_paragraph_text', paragraph: 1, text: 'A2' });
      applyHwpOp(h, { op: 'add_paragraph', text: 'C' });
      assert.deepEqual(texts(h), ['A2', 'B', '', 'C']);
      assert.equal(h.tables()[0].paragraph, 3);
      // 끝 빈 문단 뒤에 넣었으면 그 빈 문단은 더는 재사용하지 않음 (넣은 글을 덮지 않음)
      applyHwpOp(h, { op: 'add_table', rows: [['c', 'd']] });
      applyHwpOp(h, { op: 'insert_paragraph', text: 'D', at: { after_paragraph: h.paragraphCount } });
      applyHwpOp(h, { op: 'add_paragraph', text: 'E' });
      assert.deepEqual(texts(h), ['A2', 'B', '', 'C', '', '', 'D', 'E']);
      // 끝 빈 문단 자체에 글을 쓰면 다음 add_paragraph는 새 문단
      applyHwpOp(h, { op: 'add_page_break' });
      applyHwpOp(h, { op: 'set_paragraph_text', paragraph: h.paragraphCount, text: 'F' });
      applyHwpOp(h, { op: 'add_paragraph', text: 'G' });
      assert.deepEqual(texts(h).slice(-2), ['F', 'G']);
    } finally { h.close(); }
  });

  test('fill: 같은 이름 누름틀은 모두 채움', () => {
    const h = Hwp.create();
    try {
      h.addField('담당자', { label: '1장 담당' });
      h.addParagraph('본문');
      h.addField('담당자', { label: '2장 담당' });
      h.addField('기한', { label: '기한' });
      h.addField('담당자', { label: '3장 담당' });
      assert.equal(applyHwpOp(h, { op: 'fill', values: { 담당자: '김선호', 기한: '10월' } }), '채우기: 담당자(3곳), 기한');
      assert.deepEqual(h.fields().map((f) => [f.name, f.value]), [['담당자', '김선호'], ['담당자', '김선호'], ['기한', '10월'], ['담당자', '김선호']]);
      for (const e of FORMATS) {
        reopen(h.save(tmp.path(`fill-all.${e}`)), (r) => assert.deepEqual(r.fields().map((f) => f.value), ['김선호', '김선호', '10월', '김선호'], e));
        assert.equal((svgTexts(svgOf(tmp.path(`fill-all.${e}`))).map((x) => x.t).join('').match(/김선호/g) ?? []).length, 3, `${e} 그려진 값`);
      }
      assert.deepEqual(h.fill({ 기한: '11월' }), ['기한']);
    } finally { h.close(); }
  });

  test('누름틀이 표 칸 안이면 hwp_inspect에 (표 칸 안)', () => {
    const h = Hwp.open(src.hwp);
    try {
      const t = h.tables()[0];
      call0(h.raw.insertClickHereFieldInCell(t.section, t.paraIndex, t.controlIndex, 3, 0, 3, false, '칸입력', '', '칸필드', true));
      const f = h.inspect().fields.find((x) => x.name === '칸필드');
      assert.deepEqual([f.cell, f.paragraph], [true, 6]);
      assert.match(inspectText(h.inspect()), /칸필드 = \(빈 칸\) — 문단 6 \(표 칸 안\)/);
      assert.equal(h.inspect().fields.find((x) => x.name === '담당자').cell, false);
    } finally { h.close(); }
  });

  test('저장: 임시 파일에 쓴 뒤 바꿔 끼움 (제자리 저장, 임시 파일 안 남음)', () => {
    const dir = tmp.path('저장');
    mkdirSync(dir, { recursive: true });
    const target = join(dir, '제자리.hwp');
    copyFileSync(src.hwp, target);
    const h = Hwp.open(target);
    try {
      applyHwpOp(h, { op: 'insert_paragraph', text: '제자리 저장', at: 'start' });
      assert.equal(h.save(target), target);
    } finally { h.close(); }
    assert.deepEqual(readdirSync(dir), ['제자리.hwp']);
    reopen(target, (r) => assert.equal(texts(r)[0], '제자리 저장'));
  });

  test('imageInfo', () => {
    assert.deepEqual(imageInfo(PNG), { ext: 'png', width: 400, height: 200 });
    assert.deepEqual(imageInfo(JPG), { ext: 'jpg', width: 8, height: 4 });
    assert.deepEqual(imageInfo(GIF), { ext: 'gif', width: 1, height: 1 });
    assert.deepEqual(imageInfo(BMP), { ext: 'bmp', width: 8, height: 4 });
    assert.throws(() => imageInfo(Buffer.from('hello world, not an image')), /그림 형식을 알 수 없습니다/);
  });
});

// ── MCP 도구: hwp_inspect, hwp_edit 한 번에 저장·하나라도 실패하면 저장 안 함 ──

describe('MCP hwp_inspect / hwp_edit', () => {
  let client;
  before(async () => {
    client = new Client({ name: 'office-kit-hwp-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [join(ROOT, 'src', 'mcp', 'server.mjs')], cwd: tmp.dir, stderr: 'pipe' }));
  });
  after(async () => { await client?.close(); });
  const call = (name, args) => client.callTool({ name, arguments: args }, undefined, { timeout: 120_000 });
  const textOf = (r) => r.content.filter((c) => c.type === 'text').map((c) => c.text).join('\n');

  test('도구 설명에 모든 작업', async () => {
    const { tools } = await client.listTools();
    const edit = tools.find((t) => t.name === 'hwp_edit');
    for (const op of ['add_paragraph', 'insert_paragraph', 'set_paragraph_text', 'delete_paragraph', 'format_text', 'format_paragraph', 'apply_style',
      'table_set_cell', 'table_add_row', 'table_delete_row', 'set_header', 'set_footer', 'add_footnote', 'insert_image', 'page_setup']) {
      assert.ok(edit.description.includes(op), `설명에 ${op} 없음`);
      assert.ok(JSON.stringify(edit.inputSchema).includes(op), `스키마에 ${op} 없음`);
    }
    assert.match(tools.find((t) => t.name === 'hwp_inspect').description, /문단 번호/);
  });

  test('입력 오류는 한국어, 모르는 인수는 거절', async () => {
    const bad = async (args, re, tool = 'hwp_edit') => {
      const r = await call(tool, { path: '원본.hwp', ...args });
      assert.equal(r.isError, true, textOf(r));
      assert.match(textOf(r), re);
      assert.ok(!/Invalid|expected|Too (small|big)/.test(textOf(r)), `영어 메시지: ${textOf(r)}`);
    };
    await bad({ operations: [{ op: 'make_coffee' }] }, /알 수 없는 작업\(op\)입니다: make_coffee\. 가능: add_paragraph, .*page_setup/);
    await bad({ operations: [{ op: 'insert_paragraph', text: 'x', at: {} }] }, /위치\(at\)는 'end'/);
    await bad({ operations: [{ op: 'insert_paragraph', text: 'x', align: 'centre' }] }, /정렬은 left·center·right·justify·distribute 중 하나/);
    await bad({ operations: [{ op: 'add_heading', text: 'x', level: 5 }] }, /level은 1~3 정수/);
    await bad({ operations: [] }, /operations에 작업이 하나 이상/);
    await bad({ operations: [{ op: 'table_set_cell', table: 1, row: 1, text: 'x' }] }, /값이 필요합니다 \(숫자\)/);
    await bad({ operations: [{ op: 'insert_image', image_path: '그림.png', width: 30 }] }, /모르는 인수입니다: width — insert_image에 쓸 수 있는 인수: image_path, at, width_mm/);
    await bad({ operations: [{ op: 'format_text', find: '매출', bol: true, color: '#ff0000' }] }, /모르는 인수입니다: bol/);
    await bad({ operations: [{ op: 'insert_paragraph', text: 'x', format: { colour: '#ff0000' } }] }, /모르는 인수입니다: colour — format에/);
    await bad({ operations: [{ op: 'page_setup', margins_mm: { topp: 10 } }] }, /모르는 인수입니다: topp — margins_mm에/);
    await bad({ limit: 0 }, /limit은 1~1000 사이 정수/, 'hwp_inspect');
    await bad({ limit: 5000 }, /limit은 1~1000 사이 정수/, 'hwp_inspect');
    await bad({ offset: -1 }, /offset은 0 이상의 정수/, 'hwp_inspect');
    await bad({ operations: [{ op: 'insert_image', image_path: '.' }] }, /그림 파일이 아니라 폴더입니다: .+/);
  });

  test('실패하면 새로 만든 출력 폴더도 남기지 않음', async () => {
    const r = await call('hwp_edit', { path: '원본.hwp', output: '새폴더/안쪽/x.hwp', operations: [
      { op: 'insert_paragraph', text: 'x' }, { op: 'table_set_cell', table: 9, row: 1, col: 1, text: 'x' },
    ] });
    assert.equal(r.isError, true);
    assert.ok(!existsSync(tmp.path('새폴더')), '빈 폴더가 남음');
    const ok = await call('hwp_edit', { path: '원본.hwp', output: '새폴더/안쪽/x.hwp', operations: [{ op: 'insert_paragraph', text: 'x' }] });
    assert.ok(!ok.isError, textOf(ok));
    assert.ok(existsSync(tmp.path('새폴더/안쪽/x.hwp')));
  });

  test('hwp_inspect', async () => {
    const r = await call('hwp_inspect', { path: '원본.hwp', limit: 5 });
    assert.ok(!r.isError, textOf(r));
    assert.match(textOf(r), /원본\.hwp — 한글\(\.hwp\) 2쪽/);
    assert.match(textOf(r), /\[문단\] 1~5 \/ 전체 10개/);
    const bad = await call('hwp_inspect', { path: '없음.hwp' });
    assert.equal(bad.isError, true);
    assert.match(textOf(bad), /파일이 없습니다/);
  });

  for (const ext of FORMATS) {
    test(`.${ext}: 한 번에 저장, 실패하면 아무것도 저장 안 함`, async () => {
      const path = `원본.${ext}`;
      const before = readFileSync(tmp.path(path));
      const mtime = statSync(tmp.path(path)).mtimeMs;
      const bad = await call('hwp_edit', { path, operations: [
        { op: 'insert_paragraph', text: '들어가면 안 됨', at: 'start' },
        { op: 'table_set_cell', table: 5, row: 1, col: 1, text: 'x' },
      ] });
      assert.equal(bad.isError, true);
      assert.match(textOf(bad), /2번째 작업\(table_set_cell\) 실패: 표 번호가 범위를 벗어났습니다.*저장하지 않았습니다/);
      assert.ok(readFileSync(tmp.path(path)).equals(before), '원본이 바뀜');
      assert.equal(statSync(tmp.path(path)).mtimeMs, mtime);
      const bad2 = await call('hwp_edit', { path, output: `실패.${ext}`, operations: [{ op: 'insert_image', image_path: '없는.png' }] });
      assert.equal(bad2.isError, true);
      assert.ok(!existsSync(tmp.path(`실패.${ext}`)));

      const ok = await call('hwp_edit', { path, output: `mcp-결과.${ext}`, operations: [
        { op: 'insert_paragraph', text: '머리 문단', at: 'start', style: '본문' },
        { op: 'format_text', find: '매출', bold: true },
        { op: 'table_add_row', table: 1, values: ['이익', '40'] },
        { op: 'insert_image', image_path: '그림.png', at: { after_paragraph: 1 }, width_mm: 30 },
        { op: 'set_footer', text: '', page_numbers: true },
      ] });
      assert.ok(!ok.isError, textOf(ok));
      const out = textOf(ok);
      assert.match(out, /1\. 문단 넣기 → 문단 1: 머리 문단/);
      assert.match(out, /4\. 그림 넣기 → 문단 2/);
      assert.match(out, /저장: .*mcp-결과\./);
      reopen(tmp.path(`mcp-결과.${ext}`), (h) => {
        assert.equal(texts(h)[0], '머리 문단');
        assert.deepEqual(h.readTable(h.tables()[0]).at(-1), ['이익', '40']);
        assert.equal(h.inspect().counts.pictures, 1);
      });
      const img = await call('office_page_image', { path: `mcp-결과.${ext}`, page: 1, dpi: 48 });
      assert.ok(!img.isError, textOf(img));
      assert.ok(img.content.some((c) => c.type === 'image'));
    });
  }
});
