// tools/word.mjs — Word 전용 MCP 도구 (docs/FEATURES.md 'Word'). 설치된 Word를 COM으로 조작한다.
import { z } from 'zod';
import { extOf } from '../../core/index.mjs';
import {
  wordInspect, wordEdit, formatWordInspect, WORD_EXTS, WORD_OPS, STYLE_KEYS, TABLE_STYLE_KEYS, HIGHLIGHTS, ALIGNS, PAPERS,
} from '../../core/word.mjs';

// zod 메시지를 한국어로: 스키마마다 한국어 오류 함수를 붙인다 (MCP SDK 입력 검증이 이 메시지를 쓴다)
const KO = z.locales.ko().localeError;
const E = { error: (iss) => KO(iss) };
const msg = (m) => ({ error: () => m });
const str = () => z.string(E);
const num = () => z.number(E);
const bool = () => z.boolean(E);
const opt = (s) => s.optional();

const idx = num().int(msg('정수여야 합니다')).min(1, msg('1 이상이어야 합니다 (번호는 1부터)'));
const cellValue = z.union([str(), num(), bool(), z.null(E)], msg('칸 값은 글자·숫자·true/false·null이어야 합니다'));

const atSchema = z.union([
  z.enum(['end', 'start'], E),
  z.object({ after_paragraph: idx }, E),
  z.object({ before_paragraph: idx }, E),
  z.object({ after_text: str().min(1, E).max(255, E) }, E),
], msg("at은 'end' | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'찾을 글'} 중 하나여야 합니다"))
  .optional().describe("위치: 'end'(기본) | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'찾을 글'}");

const findSchema = str().min(1, E).max(255, E).describe('찾을 글자 (대소문자 구분, 본문에서. 목차 안은 건너뜀)');
const allSchema = opt(bool()).describe('true면 찾은 곳 전부(5000곳까지), 생략하면 첫 번째만');
const paragraphSchema = opt(idx).describe('문단 번호 (word_inspect 번호, 1부터)');
const paragraphsSchema = opt(z.array(idx, E).length(2, E)).describe('[시작, 끝] 문단 번호');
const textSchema = z.union([str(), num()], msg('글자(문자열)여야 합니다'));
const styleSchema = str().min(1, E).describe(`스타일 키(${STYLE_KEYS.join(', ')}) 또는 문서 안 스타일 이름`);
const alignSchema = z.enum(ALIGNS, E);

// 글자 서식 (format_text와 insert_paragraph.format 공통)
const fontFields = {
  bold: opt(bool()),
  italic: opt(bool()),
  underline: opt(bool()),
  strike: opt(bool()).describe('취소선'),
  color: opt(str()).describe("글자색 '#rrggbb'"),
  highlight: opt(z.union([z.enum(HIGHLIGHTS, E), bool()], msg(`highlight는 ${HIGHLIGHTS.join('|')} 또는 true/false여야 합니다`)))
    .describe('형광펜 (true = yellow, none/false = 없앰)'),
  size: opt(num().positive(E)).describe('글자 크기(pt)'),
  font: opt(str().min(1, E)).describe('글꼴 이름 (한글·영문 모두에 적용)'),
  superscript: opt(bool()).describe('위 첨자'),
  subscript: opt(bool()).describe('아래 첨자'),
};

const paraFields = {
  align: opt(alignSchema),
  line_spacing: opt(num().positive(E)).describe('줄 간격 배수 (1, 1.15, 1.5, 2 …)'),
  line_spacing_pt: opt(num().positive(E)).describe('줄 간격 고정값(pt)'),
  space_before: opt(num().min(0, E)).describe('문단 앞 간격(pt)'),
  space_after: opt(num().min(0, E)).describe('문단 뒤 간격(pt)'),
  first_line_indent_mm: opt(num()).describe('첫 줄 들여쓰기(mm, 음수면 내어쓰기)'),
  left_indent_mm: opt(num()).describe('왼쪽 들여쓰기(mm)'),
  right_indent_mm: opt(num()).describe('오른쪽 들여쓰기(mm)'),
  keep_with_next: opt(bool()).describe('다음 문단과 같은 쪽에'),
  page_break_before: opt(bool()).describe('이 문단 앞에서 쪽 나누기'),
};

const sectionSchema = opt(idx).describe('구역 번호 (생략하면 모든 구역)');
const op = (name, shape, what) => z.object({ op: z.literal(name, E), ...shape }, E).describe(what);

const wordOp = z.discriminatedUnion('op', [
  op('insert_paragraph', { text: textSchema, at: atSchema, style: opt(styleSchema), format: opt(z.object(fontFields, E)) }, '문단 넣기'),
  op('set_paragraph_text', { paragraph: idx, text: textSchema }, '문단 글만 바꾸기 (스타일 유지)'),
  op('delete_paragraph', { paragraph: paragraphSchema, find: opt(findSchema), all: allSchema }, '문단 지우기'),
  op('apply_style', { paragraph: paragraphSchema, paragraphs: paragraphsSchema, find: opt(findSchema), all: allSchema, style: styleSchema }, '스타일 적용'),
  op('format_text', { find: opt(findSchema), all: allSchema, paragraph: paragraphSchema, paragraphs: paragraphsSchema, ...fontFields }, '글자 서식'),
  op('format_paragraph', { paragraph: paragraphSchema, paragraphs: paragraphsSchema, find: opt(findSchema), all: allSchema, ...paraFields }, '문단 서식'),
  op('set_list', { paragraph: paragraphSchema, paragraphs: paragraphsSchema, type: z.enum(['bullet', 'number', 'none'], E) }, '글머리표·번호 목록'),
  op('insert_table', {
    rows: z.array(z.array(cellValue, E).min(1, E), E).min(1, E).describe('[["머리1","머리2"],["값1","값2"]]'),
    at: atSchema,
    style: opt(str().min(1, E)).describe(`표 스타일 키(${TABLE_STYLE_KEYS.join(', ')}) 또는 문서 안 표 스타일 이름. 생략하면 격자 테두리`),
    header_bold: opt(bool()).describe('첫 행 굵게·제목 행 반복 (기본 true)'),
    autofit: opt(z.union([z.enum(['window', 'content', 'fixed'], E), bool()], msg("autofit은 'window' | 'content' | 'fixed' 이어야 합니다")))
      .describe("'window'(쪽 너비, 기본) | 'content'(내용에 맞춤) | 'fixed'"),
  }, '표 넣기'),
  op('table_set_cell', { table: idx, row: idx, col: idx, text: textSchema }, '표 칸 고치기'),
  op('table_add_row', {
    table: idx,
    values: opt(z.array(cellValue, E)).describe('열 순서대로 칸 값 (null = 비워 둠, 세로로 병합된 칸 자리는 null)'),
    after_row: opt(num().int(E).min(0, E)).describe('이 행 뒤에 (0 = 맨 위, 생략하면 맨 아래)'),
  }, '표 행 추가'),
  op('table_delete_row', { table: idx, row: idx }, '표 행 지우기'),
  op('insert_image', {
    image_path: str().min(1, E).describe('그림 파일 (png·jpg·jpeg·gif·bmp)'),
    at: atSchema,
    width_mm: opt(num().positive(E)),
    height_mm: opt(num().positive(E)),
    caption: opt(str()).describe("그림 아래 캡션 ('그림 N. 캡션', 번호 자동)"),
    caption_label: opt(str().min(1, E)).describe("캡션 머리말 (기본 '그림')"),
    align: opt(alignSchema).describe('그림 문단 정렬 (기본 center)'),
  }, '그림 넣기'),
  op('insert_page_break', { at: atSchema }, '쪽 나누기'),
  op('set_header', { text: opt(textSchema), align: opt(alignSchema), page_numbers: opt(bool()), section: sectionSchema }, '머리말'),
  op('set_footer', { text: opt(textSchema), align: opt(alignSchema), page_numbers: opt(bool()), section: sectionSchema }, '꼬리말'),
  op('page_setup', {
    orientation: opt(z.enum(['portrait', 'landscape'], E)),
    paper: opt(str().min(1, E)).describe(`용지: ${PAPERS.join(' | ')}`),
    margins_mm: opt(z.object({ top: opt(num().min(0, E)), bottom: opt(num().min(0, E)), left: opt(num().min(0, E)), right: opt(num().min(0, E)) }, E)),
    section: sectionSchema,
  }, '쪽 설정'),
  op('add_comment', { find: findSchema, text: textSchema, all: allSchema, author: opt(str().min(1, E)) }, '메모 달기'),
  op('delete_comments', { all: opt(bool()), index: opt(idx).describe('메모 번호 (word_inspect 번호)') }, '메모 지우기'),
  op('track_changes', { on: bool() }, '변경 내용 추적 켜기/끄기'),
  op('accept_revisions', {}, '변경 내용 모두 수락'),
  op('reject_revisions', {}, '변경 내용 모두 거부'),
  op('insert_toc', {
    at: atSchema,
    levels: opt(num().int(E).min(1, E).max(9, E)).describe('제목 몇 수준까지 (기본 3)'),
    title: opt(str()).describe("목차 위 제목 (예: '목차')"),
  }, '목차 넣기'),
  op('update_fields', {}, '목차·쪽 번호 등 필드 새로 고침'),
  op('add_footnote', { find: findSchema, text: textSchema, all: allSchema }, '각주 달기 (찾은 글 뒤에 각주 번호)'),
  op('insert_hyperlink', { find: findSchema, url: str().min(1, E).describe('http(s)://, mailto:, ftp:// 주소'), all: allSchema }, '찾은 글에 링크 걸기'),
], {
  error: (iss) => (iss.code === 'invalid_union' && iss.note === 'No matching discriminator'
    ? `알 수 없는 작업(op)입니다${typeof iss.input?.op === 'string' ? `: ${iss.input.op}` : ''} (가능: ${WORD_OPS.join(', ')})`
    : KO(iss)),
});

const EDIT_HELP = [
  'Word 문서(docx·doc·rtf·odt)를 설치된 Word로 편집한다. operations를 순서대로 모두 적용하고 마지막에 한 번만 저장한다. '
    + '하나라도 실패하면 아무것도 저장하지 않는다 (몇 번째 작업이 왜 실패했는지 알려 줌).',
  '먼저 word_inspect로 문단·표 번호를 확인한다. 번호는 1부터. 앞 작업이 문단을 넣거나 지우면 뒤 작업은 바뀐 번호를 쓴다 (결과 줄에 새 번호와 전체 문단 수가 나온다).',
  "공통: at = 'end'(기본) | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'찾을 글'} · find = 대소문자 구분, 첫 번째만 (all:true면 전부, 5000곳까지), 목차 안은 건너뜀 · "
    + '글 안의 줄바꿈(\\n)은 문단 안 줄바꿈 · 색은 #rrggbb · 길이는 mm, 글자 크기·간격은 pt',
  `스타일 키: ${STYLE_KEYS.join(', ')} (한국어 Word에서도 동작). 그 밖의 문자열은 문서 안 스타일 이름(word_inspect '쓰인 스타일').`,
  '작업 종류:',
  "  {op:'insert_paragraph', text, at?, style?, format?:{bold, italic, underline, strike, color, highlight, size, font, superscript, subscript}}  "
    + '(style 생략 시 Word에서 Enter를 친 것처럼: 제목 뒤면 본문 스타일, 목록 뒤면 목록 계속)',
  "  {op:'set_paragraph_text', paragraph, text}  문단 글만 바꿈 (스타일 유지)",
  "  {op:'delete_paragraph', paragraph | find, all?}  변경 내용 추적 중이면 지움 표시만 (번호 그대로). 구역 나누기·표 앞 문단은 앞 문단과 합쳐 지움",
  "  {op:'apply_style', paragraph | paragraphs:[from,to] | find, all?, style}",
  "  {op:'format_text', find | paragraph | paragraphs, all?, bold, italic, underline, strike, color, highlight(yellow|green|cyan|pink|red|blue|gray|none…), size, font, superscript, subscript}",
  "  {op:'format_paragraph', paragraph | paragraphs | find, align(left|center|right|justify|distribute), line_spacing(배수), line_spacing_pt, space_before, space_after, "
    + 'first_line_indent_mm, left_indent_mm, right_indent_mm, keep_with_next, page_break_before}',
  "  {op:'set_list', paragraph | paragraphs:[from,to], type:'bullet'|'number'|'none'}",
  `  {op:'insert_table', rows:[["머리1","머리2"],["값1","값2"]], at?, style?(${TABLE_STYLE_KEYS.join('|')} 또는 이름), header_bold?(기본 true), autofit?('window'|'content'|'fixed')}`,
  "  {op:'table_set_cell', table, row, col, text}",
  "  {op:'table_add_row', table, values?:[...](열 순서, 세로로 병합된 칸 자리는 null), after_row?(0=맨 위, 생략=맨 아래)}",
  "  {op:'table_delete_row', table, row}",
  "  {op:'insert_image', image_path(png·jpg·jpeg·gif·bmp), at?, width_mm?, height_mm?(하나만 주면 비율 유지, 둘 다 없으면 쪽 너비에 맞춰 줄임), caption?, caption_label?('그림'), align?(기본 center)}",
  "  {op:'insert_page_break', at?}",
  "  {op:'set_header' | 'set_footer', text?, align?, page_numbers?, section?}  text 안의 {page}=쪽 번호, {pages}=전체 쪽 수. "
    + 'page_numbers:true면 {page}가 없을 때 끝에 붙임 (꼬리말은 가운데 정렬). 머리말·꼬리말 내용을 통째로 바꾼다',
  `  {op:'page_setup', orientation?('portrait'|'landscape'), paper?(${PAPERS.join('|')}), margins_mm?:{top,bottom,left,right}, section?}`,
  "  {op:'add_comment', find, text, all?, author?}",
  "  {op:'delete_comments', all:true | index}",
  "  {op:'track_changes', on:true|false}  켠 뒤의 작업은 변경 내용으로 기록된다",
  "  {op:'accept_revisions'} / {op:'reject_revisions'}  변경 내용 모두 수락/거부",
  "  {op:'insert_toc', at?(보통 'start'), levels?(기본 3), title?}  제목(heading1~) 스타일로 목차를 만든다",
  "  {op:'update_fields'}  목차·쪽 번호·캡션 번호 등 새로 고침",
  "  {op:'add_footnote', find, text, all?}",
  "  {op:'insert_hyperlink', find, url, all?}",
  'output을 생략하면 원본에 저장한다 (다른 이름이면 .docx/.doc/.rtf/.odt). 결과는 office_page_image로 눈으로 확인할 수 있다.',
].join('\n');

/** server.mjs가 부른다. ctx: 공통 도우미 (tool, text, mustExist, outPath, 스키마 등) */
export function register(server, ctx) {
  const { tool, text, abs, mustExist, outPath, pathSchema, outputSchema, overwriteSchema } = ctx;

  const mustWord = (full) => {
    if (!WORD_EXTS.has(extOf(full))) {
      throw new Error('Word 문서(.docx/.doc/.rtf/.odt)만 다룹니다 — 한글은 hwp_inspect/hwp_edit, PowerPoint는 ppt_inspect/ppt_edit, Excel은 excel_inspect/excel_edit');
    }
    return full;
  };

  server.registerTool('word_inspect', {
    title: 'Word 문서 구조 보기',
    description: 'Word 문서(docx·doc·rtf·odt)의 구조를 설치된 Word로 살펴본다: 쪽 수, 단어 수, 구역별 쪽 설정(용지·방향·여백), 머리말·꼬리말, '
      + '문단 목록(번호·스타일·글 200자·표 안 위치·목록 여부), 표 목록(번호·행×열·첫 행), 메모, 변경 내용 추적 상태·개수, '
      + '목차·그림·콘텐츠 컨트롤·책갈피·각주·링크 개수, 쓰인 스타일. word_edit에 쓸 번호를 여기서 확인한다. '
      + "긴 문서는 문단을 limit개씩 보여 주고 '[계속] offset=N'을 알려 준다.",
    inputSchema: {
      path: pathSchema('Word 문서 경로 (.docx/.doc/.rtf/.odt)'),
      offset: num().int(E).min(0, E).default(0).describe('건너뛸 문단 수 (0부터, 다음 쪽을 볼 때 결과의 offset)'),
      limit: num().int(E).min(1, E).max(1000, E).default(200).describe('보여 줄 문단 수 (기본 200)'),
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, offset = 0, limit = 200 }) => {
    const full = mustWord(mustExist(path));
    return text(formatWordInspect(await wordInspect(full, { offset, limit })));
  }));

  server.registerTool('word_edit', {
    title: 'Word 문서 편집',
    description: EDIT_HELP,
    inputSchema: {
      path: pathSchema('Word 문서 경로 (.docx/.doc/.rtf/.odt)'),
      operations: z.array(wordOp, E).min(1, E).describe('순서대로 적용할 작업 목록'),
      output: outputSchema,
      overwrite: overwriteSchema,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, operations, output, overwrite = false }) => {
    const full = mustWord(mustExist(path));
    if (output && !WORD_EXTS.has(extOf(abs(output)))) throw new Error('저장 경로는 .docx/.doc/.rtf/.odt 여야 합니다 (PDF는 office_convert)');
    const out = outPath(output, { source: full, allow: WORD_EXTS, overwrite });
    // 그림 경로도 서버 작업 폴더 기준
    const ops = operations.map((op) => (op.op === 'insert_image' && op.image_path ? { ...op, image_path: mustExist(op.image_path) } : op));
    const r = await wordEdit(full, ops, { output: out });
    return text(r.lines.join('\n'));
  }));
}
