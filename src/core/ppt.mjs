// ppt.mjs — PowerPoint 전용 기능: 살펴보기(pptInspect)와 편집(pptEdit). 설치된 PowerPoint(COM)가 필요하다.
// 작업 실행은 scripts/com/ppt-ops.ps1 (office.ps1이 불러옴). 규칙은 docs/FEATURES.md.
import { existsSync, openSync, readSync, closeSync, statSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, extname, join, resolve } from 'node:path';
import JSZip from 'jszip';
import { z } from 'zod';
import { comCall, comAvailableFor } from './com.mjs';

export const PPT_EXTS = ['pptx', 'ppt', 'odp'];
export const LAYOUT_KEYS = ['title', 'title_content', 'section', 'two_content', 'comparison', 'title_only', 'blank'];
const IMAGE_EXTS = ['png', 'jpg', 'jpeg', 'gif', 'bmp'];
const THEME_EXTS = ['thmx', 'potx'];
// 살펴보기: MCP 기본 슬라이드 수, 슬라이드마다 보여 줄 도형 수, 나눠 볼 때 글자 수 한도
export const INSPECT_LIMIT = 30;
const MAX_SHAPES = 60;
const MAX_SHAPES_ONE = 500;
const INSPECT_BUDGET = 60_000;

const extOf = (p) => extname(p).slice(1).toLowerCase();
const list = (v) => (v == null ? [] : [].concat(v));
const short = (t, n = 40) => { const s = String(t ?? '').replace(/\s+/g, ' ').trim(); return s.length > n ? `${s.slice(0, n)}…` : s; };

export const NEED_POWERPOINT = 'PowerPoint가 설치돼 있어야 합니다 — 대신 office_read / office_replace / office_fill / office_create는 쓸 수 있습니다';

// ── 작업 스키마 (MCP 도구와 명령줄이 함께 쓴다) ─────────────────
// zod 메시지를 한국어로: 스키마마다 한국어 오류 함수를 붙인다 (MCP SDK 입력 검증도 이 메시지를 쓴다)
const KO = z.locales.ko().localeError;
const E = { error: (iss) => KO(iss) };
const msg = (m) => ({ error: () => m });
/** zod 스키마에 붙이는 한국어 오류 (MCP 도구 입력 스키마도 같이 쓴다) */
export const zodKo = E;

const int1 = (what) => z.number(E).int(E).min(1, E).describe(what);
const slideNo = int1('슬라이드 번호 (1부터, ppt_inspect의 번호)');
const slideOrAll = z.union([z.number(E).int(E).min(1, E), z.literal('all', E)], msg("슬라이드 번호(1부터) 또는 'all'이어야 합니다"))
  .describe("슬라이드 번호(1부터) 또는 'all'(전체)");
const shapeRef = z.union([z.string(E).min(1, E), z.number(E).int(E).min(1, E)], msg('도형 이름(문자열) 또는 번호(1부터)여야 합니다'))
  .describe('도형 이름 또는 번호(1부터) — ppt_inspect의 도형 목록 참고');
const PCT = /^\s*-?\d+(?:\.\d+)?\s*%\s*$/;
const PCT_MSG = msg("pt 숫자 또는 '10%' 같은 비율 문자열이어야 합니다");
const axis = (what) => (/^위 |높이/.test(what) ? '높이' : '너비');
const pos = (what) => z.union([z.number(E), z.string(E).regex(PCT, PCT_MSG)], PCT_MSG)
  .describe(`${what} (pt, 또는 '10%'처럼 슬라이드 ${axis(what)} 비율)`);
// 너비·높이: 0이나 음수('-10%'·'0%' 포함)는 거절
const size = (what) => z.union([z.number(E), z.string(E).regex(PCT, PCT_MSG)], PCT_MSG)
  .refine((v) => (typeof v === 'number' ? v : parseFloat(v)) > 0, msg("0보다 커야 합니다 (pt 또는 '10%')"))
  .describe(`${what} (pt, 또는 '10%'처럼 슬라이드 ${axis(what)} 비율)`);
const color = (what = '색') => z.string(E).regex(/^#?[0-9a-fA-F]{6}$/, msg("색은 '#rrggbb' 형식이어야 합니다 (예: '#1f4e79')")).describe(`${what} '#rrggbb'`);
const text = (what) => z.string(E).describe(what);
const bullet = z.union([z.string(E), z.strictObject({ text: z.string(E), level: z.number(E).int(E).min(1, E).max(9, E).optional().describe('수준 1~9 (기본 1)') }, E)],
  msg('글머리표 항목은 문자열 또는 {text, level}이어야 합니다'));
const bulletsSchema = z.union([z.string(E), z.array(bullet, E)], msg('문자열(줄마다 문단) 또는 [{text, level}] 배열이어야 합니다'))
  .describe('문자열(줄마다 한 문단) 또는 ["항목", {text:"하위 항목", level:2}] 글머리표 배열');
const fontOpts = {
  bold: z.boolean(E).optional(), italic: z.boolean(E).optional(), underline: z.boolean(E).optional(),
  color: color('글자색').optional(), size: z.number(E).positive(E).max(4000, E).optional().describe('글자 크기(pt)'),
  font: z.string(E).min(1, E).optional().describe('글꼴 이름 (예: 맑은 고딕)'),
};
const op = (name, shape, what) => z.strictObject({ op: z.literal(name, E), ...shape }, E).describe(what);

export const pptOpSchema = z.discriminatedUnion('op', [
  op('add_slide', {
    layout: z.string(E).min(1, E).optional().describe(`레이아웃 키 ${LAYOUT_KEYS.join('|')} (기본 title_content) 또는 ppt_inspect에 나온 레이아웃 이름`),
    position: int1('넣을 위치 (1부터, 기본: 맨 뒤)').optional(),
    title: text('제목').optional(), body: bulletsSchema.optional(), notes: text('발표자 노트').optional(),
  }, '슬라이드 추가'),
  op('delete_slide', { slide: slideNo }, '슬라이드 지우기 (뒤 슬라이드 번호가 하나씩 당겨짐)'),
  op('duplicate_slide', { slide: slideNo, position: int1('복제본 위치 (기본: 원본 바로 뒤)').optional() }, '슬라이드 복제'),
  op('move_slide', { slide: slideNo, to: int1('옮길 위치 (1부터)') }, '슬라이드 순서 바꾸기'),
  op('set_title', { slide: slideNo, text: text('제목 글') }, '제목 개체 틀 채우기'),
  op('set_body', { slide: slideNo, bullets: bulletsSchema, placeholder: int1('몇 번째 내용 개체 틀인지 (기본 1, 두 콘텐츠·비교 레이아웃의 오른쪽은 2)').optional() }, '본문 개체 틀 채우기 (글머리표)'),
  op('set_notes', { slide: slideNo, text: text('발표자 노트') }, '발표자 노트 쓰기 (기존 노트를 바꿈)'),
  op('set_text', { slide: slideNo, shape: shapeRef, text: text('새 글 (줄바꿈 = 새 문단)') }, '도형 글 바꾸기'),
  op('add_textbox', {
    slide: slideNo, text: text('글'), left: pos('왼쪽 위치'), top: pos('위 위치'), width: size('너비'), height: size('높이 (글이 더 길면 글에 맞춰 늘어남)'),
    size: fontOpts.size, bold: fontOpts.bold, color: fontOpts.color, font: fontOpts.font,
    align: z.enum(['left', 'center', 'right', 'justify'], E).optional().describe('문단 정렬'),
  }, '글상자 추가'),
  op('add_image', {
    slide: slideNo, image_path: z.string(E).min(1, E).describe('그림 파일 경로 (png·jpg·jpeg·gif·bmp)'),
    left: pos('왼쪽 위치'), top: pos('위 위치'), width: size('너비 (생략하면 원래 비율)').optional(), height: size('높이 (생략하면 원래 비율)').optional(),
  }, '그림 넣기 (너비·높이 중 하나만 주면 비율 유지, 둘 다 없으면 원래 크기 — 슬라이드를 넘으면 줄임)'),
  op('add_table', {
    slide: slideNo, rows: z.array(z.array(z.union([z.string(E), z.number(E)], E), E).min(1, E), E).min(1, E).describe('[["머리1","머리2"],["값1","값2"]]'),
    left: pos('왼쪽 위치'), top: pos('위 위치'), width: size('너비'), height: size('높이 (생략하면 내용에 맞춤)').optional(),
    header_bold: z.boolean(E).optional().describe('첫 행 굵게 (기본 true)'),
  }, '표 넣기'),
  op('table_set_cell', { slide: slideNo, shape: shapeRef, row: int1('행 (1부터)'), col: int1('열 (1부터)'), text: text('칸 글') }, '표 칸 고치기'),
  op('add_shape', {
    slide: slideNo, shape: z.enum(['rectangle', 'rounded_rectangle', 'oval', 'arrow', 'line'], E).describe('도형 종류 (arrow = 오른쪽 화살표, line은 (left,top)에서 (left+width, top+height)까지)'),
    left: pos('왼쪽 위치'), top: pos('위 위치'), width: pos('너비'), height: pos('높이'),
    fill: z.union([color(), z.literal('none', E)], msg("채우기는 '#rrggbb' 또는 'none'이어야 합니다")).optional().describe("채우기 '#rrggbb' 또는 'none'"),
    line: z.union([color(), z.literal('none', E)], msg("선 색은 '#rrggbb' 또는 'none'이어야 합니다")).optional().describe("테두리·선 색 '#rrggbb' 또는 'none'"),
    text: text('도형 안 글 (line 제외)').optional(),
    color: color('글자색 (생략하면 채우기에 맞춰: 채우기 없음·밝은 채우기는 어두운 글자)').optional(),
  }, '도형 넣기'),
  op('add_chart', {
    slide: slideNo, type: z.enum(['column', 'bar', 'line', 'pie'], E).describe('세로 막대 column | 가로 막대 bar | 꺾은선 line | 원형 pie(첫 계열만)'),
    categories: z.array(z.union([z.string(E), z.number(E)], E), E).min(1, E).describe('항목 이름 (가로축)'),
    series: z.array(z.strictObject({ name: z.string(E).describe('계열 이름'), values: z.array(z.number(E), E).describe('항목마다 값 (categories와 같은 개수)') }, E), E).min(1, E).describe('[{name, values}]'),
    left: pos('왼쪽 위치'), top: pos('위 위치'), width: size('너비'), height: size('높이'),
    title: text('차트 제목 (생략하면 제목 없음)').optional(),
  }, '차트 넣기 (데이터는 차트 안 Excel 데이터로 저장돼 나중에 PowerPoint에서 "데이터 편집" 가능)'),
  op('format_text', {
    slide: slideNo.optional().describe('슬라이드 번호 (생략하면 전체)'), find: z.string(E).min(1, E).describe('찾을 글 (대소문자 구분)'),
    all: z.boolean(E).optional().describe('true면 모두, 기본은 첫 번째만'), ...fontOpts,
  }, '글자 서식 (표·그룹 안 글 포함)'),
  op('delete_shape', { slide: slideNo, shape: shapeRef }, '도형 지우기'),
  op('set_background', { slide: slideOrAll, color: color('배경색') }, "배경색 (all이면 모든 슬라이드와 마스터)"),
  op('set_transition', {
    slide: slideOrAll, effect: z.enum(['none', 'fade', 'push', 'wipe', 'split', 'cover'], E).describe('화면 전환 효과'),
    duration_s: z.number(E).positive(E).max(60, E).optional().describe('전환 시간(초)'),
  }, '화면 전환'),
  op('hide_slide', { slide: slideNo, hidden: z.boolean(E).optional().describe('true 숨김(기본) / false 다시 보이기') }, '슬라이드 숨기기 (쇼에서 건너뜀)'),
  op('apply_theme', { theme_path: z.string(E).min(1, E).describe('테마 파일 경로 (.thmx 또는 .potx)') }, '테마 적용 (모든 슬라이드)'),
  op('set_slide_size', { preset: z.enum(['16:9', '4:3'], E).describe("'16:9'(960×540pt) 또는 '4:3'(720×540pt)") },
    '슬라이드 크기 (좁히면 내용을 비율에 맞춰 줄이고, 넓히면 내용 크기는 그대로 두고 가운데로 옮김 — 양옆이 빔)'),
], { error: (iss) => (iss.code === 'invalid_union' && iss.note === 'No matching discriminator'
  ? `알 수 없는 작업(op)입니다. 가능: ${OP_NAMES.join(', ')}` : KO(iss)) });

export const OP_NAMES = ['add_slide', 'delete_slide', 'duplicate_slide', 'move_slide', 'set_title', 'set_body', 'set_notes', 'set_text',
  'add_textbox', 'add_image', 'add_table', 'table_set_cell', 'add_shape', 'add_chart', 'format_text', 'delete_shape',
  'set_background', 'set_transition', 'hide_slide', 'apply_theme', 'set_slide_size'];

/** 작업 설명 (MCP 도구 설명·명령줄 도움말) */
export const PPT_OPS_HELP = [
  '작업 종류 (슬라이드·도형·행·열 번호는 1부터, ppt_inspect 결과의 번호와 이름을 쓴다. 위치·크기는 pt 또는 \'10%\' 같은 슬라이드 비율):',
  `  {op:"add_slide", layout?(${LAYOUT_KEYS.join('|')} 또는 레이아웃 이름, 기본 title_content), position?, title?, body?(문자열 또는 [{text, level}]), notes?}`,
  '  {op:"delete_slide", slide}   {op:"duplicate_slide", slide, position?}   {op:"move_slide", slide, to}',
  '  {op:"set_title", slide, text}   {op:"set_body", slide, bullets(문자열 또는 [{text, level}]), placeholder?(두 번째 내용 칸이면 2)}',
  '  {op:"set_notes", slide, text}   {op:"set_text", slide, shape(이름 또는 번호), text}',
  '  {op:"add_textbox", slide, text, left, top, width, height(글이 더 길면 늘어남), size?, bold?, color?("#rrggbb"), align?(left|center|right|justify), font?}',
  '  {op:"add_image", slide, image_path(png·jpg·gif·bmp), left, top, width?, height?}',
  '  {op:"add_table", slide, rows:[["머리1","머리2"],["값1","값2"]], left, top, width, height?, header_bold?(기본 true)}',
  '  {op:"table_set_cell", slide, shape, row, col, text}',
  '  {op:"add_shape", slide, shape(rectangle|rounded_rectangle|oval|arrow|line), left, top, width, height, fill?("#rrggbb"|"none"), line?, text?, color?(글자색)}',
  '  {op:"add_chart", slide, type(column|bar|line|pie), categories:["1분기",...], series:[{name, values:[...]}], left, top, width, height, title?}',
  '  {op:"format_text", slide?(없으면 전체), find, all?(기본 첫 번째만), bold?, italic?, underline?, color?, size?, font?}  (대소문자 구분)',
  '  {op:"delete_shape", slide, shape}   {op:"set_background", slide|"all", color}',
  '  {op:"set_transition", slide|"all", effect(none|fade|push|wipe|split|cover), duration_s?}   {op:"hide_slide", slide, hidden?(기본 true)}',
  '  {op:"apply_theme", theme_path(.thmx|.potx)}   {op:"set_slide_size", preset("16:9"|"4:3")}  (좁히면 내용을 줄이고, 넓히면 내용 크기 그대로 가운데로)',
].join('\n');

const where = (path) => (path.length && path.join('.') !== 'op' ? `'${path.join('.')}' ` : '');

// ── 파일 첫 바이트 확인 ───────────────────────────────────
const CFB = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]); // 옛 Office 파일(복합 문서) = 암호화된 OOXML 또는 .ppt
const ZIP = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
const IMAGE_SIGS = [[0x89, 0x50, 0x4e, 0x47], [0xff, 0xd8, 0xff], [0x47, 0x49, 0x46, 0x38], [0x42, 0x4d]]; // png jpg gif bmp

function headBytes(path, n = 8) {
  let fd;
  try {
    fd = openSync(path, 'r');
    const b = Buffer.alloc(n);
    return b.subarray(0, readSync(fd, b, 0, n, 0));
  } catch { return Buffer.alloc(0); } finally { if (fd !== undefined) closeSync(fd); }
}
const startsWith = (b, sig) => b.length >= sig.length && sig.every((x, i) => b[i] === x);
const isImageFile = (path) => IMAGE_SIGS.some((s) => startsWith(headBytes(path, 8), s));

/** 작업 목록 검사 → 정리된 작업 (경로는 절대 경로로). 틀리면 몇 번째 작업이 왜 틀렸는지 한국어로 */
export function validatePptOps(operations, { base = process.cwd() } = {}) {
  if (!Array.isArray(operations) || !operations.length) throw new Error('operations에 작업이 하나 이상 있어야 합니다');
  return operations.map((raw, i) => {
    const label = `${i + 1}번째 작업${raw && typeof raw === 'object' && typeof raw.op === 'string' ? `(${raw.op})` : ''}`;
    const r = pptOpSchema.safeParse(raw, { error: (iss) => KO(iss) });
    if (!r.success) {
      const why = r.error.issues.map((iss) => `${where(iss.path)}${iss.message}`).join('; ');
      throw new Error(`${label}: ${why}`);
    }
    const o = { ...r.data };
    const bad = (m) => { throw new Error(`${label}: ${m}`); };
    if (o.image_path) {
      o.image_path = resolve(base, o.image_path);
      if (!IMAGE_EXTS.includes(extOf(o.image_path))) bad(`그림은 ${IMAGE_EXTS.join('·')} 파일만 넣을 수 있습니다: ${o.image_path}`);
      if (!existsSync(o.image_path)) bad(`그림 파일이 없습니다: ${o.image_path}`);
      // 확장자만 맞는 다른 파일은 PowerPoint가 빈 그림 틀로 넣어 버린다
      if (!isImageFile(o.image_path)) bad(`그림 파일이 아니거나 손상됐습니다: ${o.image_path}`);
    }
    if (o.theme_path) {
      o.theme_path = resolve(base, o.theme_path);
      if (!THEME_EXTS.includes(extOf(o.theme_path))) bad(`테마는 .thmx 또는 .potx 파일이어야 합니다: ${o.theme_path}`);
      if (!existsSync(o.theme_path)) bad(`테마 파일이 없습니다: ${o.theme_path}`);
      const head = headBytes(o.theme_path, 8);
      if (startsWith(head, CFB)) bad(`암호가 걸린 테마 파일은 쓸 수 없습니다 — 암호를 푼 뒤 다시 시도하세요: ${o.theme_path}`);
      if (!startsWith(head, ZIP)) bad(`테마 파일이 아니거나 손상됐습니다: ${o.theme_path}`);
    }
    if (o.op === 'add_chart') {
      const n = o.categories.length;
      const off = o.series.filter((s) => s.values.length !== n);
      if (off.length) bad(`계열 값 개수가 항목(categories) ${n}개와 다릅니다: ${off.map((s) => `${s.name}(${s.values.length}개)`).join(', ')}`);
      o.categories = o.categories.map(String);
    }
    if (o.op === 'add_table') o.rows = o.rows.map((row) => row.map(String));
    if (o.op === 'format_text' && !['bold', 'italic', 'underline', 'color', 'size', 'font'].some((k) => o[k] !== undefined)) {
      bad('바꿀 서식(bold, italic, underline, color, size, font)을 하나 이상 주세요');
    }
    if (o.op === 'add_shape' && o.shape !== 'line') {
      for (const k of ['width', 'height']) if (typeof o[k] === 'number' && o[k] <= 0) bad(`${k}는 0보다 커야 합니다`);
      for (const k of ['width', 'height']) if (typeof o[k] === 'string' && parseFloat(o[k]) <= 0) bad(`${k}는 0보다 커야 합니다`);
    }
    if (o.op === 'add_shape' && o.shape === 'line' && o.text) bad('선(line)에는 글을 넣을 수 없습니다');
    if (o.op === 'add_shape' && o.color && o.text == null) bad('color(글자색)는 text와 함께 써야 합니다 (선 색은 line)');
    if (o.color) o.color = `#${o.color.replace(/^#/, '').toLowerCase()}`;
    return o;
  });
}

// ── 오류 ─────────────────────────────────────────────────

const ERRORS = {
  slide_range: ([n, count]) => `슬라이드 ${n}번이 없습니다 (지금 ${count}장)`,
  position_range: ([n, max]) => `위치 ${n}이(가) 범위를 벗어났습니다 (1~${max})`,
  shape_range: ([s, n, count]) => `슬라이드 ${s}에 ${n}번 도형이 없습니다 (도형 ${count}개)`,
  shape_not_found: ([s, ...name]) => `슬라이드 ${s}에 '${name.join('|')}' 도형이 없습니다 (ppt_inspect로 도형 이름을 확인하세요)`,
  not_table: ([s, ...name]) => `슬라이드 ${s}의 '${name.join('|')}'은(는) 표가 아닙니다`,
  is_table: ([s, ...name]) => `슬라이드 ${s}의 '${name.join('|')}'은(는) 표입니다 — table_set_cell로 칸을 고치세요`,
  no_text: ([s, ...name]) => `슬라이드 ${s}의 '${name.join('|')}' 도형에는 글을 넣을 수 없습니다`,
  cell_range: ([r, c, rows, cols]) => `${r}행 ${c}열 칸이 없습니다 (표 ${rows}행×${cols}열)`,
  no_title: ([s]) => `슬라이드 ${s}에 제목 개체 틀이 없습니다 — set_text(도형 이름) 또는 add_textbox를 쓰세요`,
  no_body: ([s, k, n]) => (Number(n) === 0
    ? `슬라이드 ${s}에 본문(내용) 개체 틀이 없습니다 — set_text 또는 add_textbox를 쓰세요`
    : `슬라이드 ${s}에 ${k}번째 내용 개체 틀이 없습니다 (${n}개 있음)`),
  no_notes: ([s]) => `슬라이드 ${s}의 노트 페이지에 노트 칸이 없습니다`,
  not_found: (find) => `찾는 글자가 없습니다: '${find.join('|')}'`,
  layout_not_found: ([name, ...names]) => `레이아웃 '${name}'이(가) 없습니다 (키: ${LAYOUT_KEYS.join(', ')} / 이 문서의 레이아웃: ${names.join('|')})`,
  bad_size: ([what, ...v]) => `${what} 값이 올바르지 않습니다: ${v.join('|')} (pt 숫자 또는 '10%')`,
  size_not_positive: ([what, ...v]) => `${what} 값은 0보다 커야 합니다: ${v.join('|')}`,
  bad_color: (v) => `색이 올바르지 않습니다: ${v.join('|')} ('#rrggbb')`,
  no_file: (p) => `파일이 없습니다: ${p.join('|')}`,
  image_failed: (m) => `그림을 넣지 못했습니다: ${m.join('|')}`,
  chart_failed: (m) => `차트를 만들지 못했습니다: ${m.join('|')}`,
  theme_failed: (m) => `테마를 적용하지 못했습니다: ${m.join('|')}`,
  line_text: () => '선(line)에는 글을 넣을 수 없습니다',
  bad_shape: ([k]) => `알 수 없는 도형 종류: ${k}`,
  bad_chart: ([k]) => `알 수 없는 차트 종류: ${k}`,
  bad_effect: ([k]) => `알 수 없는 화면 전환: ${k}`,
  bad_preset: ([k]) => `알 수 없는 슬라이드 크기: ${k}`,
  unknown_op: ([k]) => `알 수 없는 작업: ${k}`,
};

function codeMessage(s) {
  const m = String(s).match(/^E\|([a-z_]+)\|?([\s\S]*)$/);
  if (!m) return null;
  const args = m[2] === '' ? [] : m[2].split('|');
  return ERRORS[m[1]] ? ERRORS[m[1]](args) : `PowerPoint 작업 실패: ${m[1]} ${args.join(' ')}`.trim();
}

// PowerPoint가 영어로 돌려주는 흔한 COM 오류
const COM_MESSAGES = [
  [/specified value is out of range|value (?:is )?out of range/i, '값이 허용 범위를 벗어났습니다 (위치·크기·번호를 확인하세요)'],
];
const comMessage = (m) => {
  const t = String(m).trim();
  const hit = COM_MESSAGES.find(([re]) => re.test(t));
  return hit ? `PowerPoint 오류: ${hit[1]}` : `PowerPoint 오류: ${t}`;
};

/** office.ps1이 돌려준 오류 → 한국어 (몇 번째 작업이 왜 실패했는지). output: 저장하려던 경로 (저장 실패 안내용) */
export function pptErrorMessage(message, { saving = true, output } = {}) {
  const s = String(message ?? '').replace(/^PowerPoint 작업 실패:\s*/, '');
  const f = s.match(/opfail\|(\d+)\|([^|]*)\|([\s\S]*)$/);
  if (f) {
    const inner = codeMessage(f[3]) ?? comMessage(f[3]);
    return `${f[1]}번째 작업(${f[2]}) 실패: ${inner}${saving ? ' — 저장하지 않았습니다' : ''}`;
  }
  const c = s.match(/E\|[a-z_]+[\s\S]*$/);
  if (c) return codeMessage(c[0]);
  // 모든 작업은 됐지만 저장(Save/SaveAs)에서 실패
  if (/Presentation\.Save(?:As|CopyAs)?\b|while PowerPoint was saving/i.test(s)) {
    return `문서를 저장하지 못했습니다${output ? `: ${output}` : ''} — 저장할 파일이 다른 프로그램에서 열려 있거나 쓰기 권한이 없을 수 있습니다`;
  }
  return String(message ?? '');
}

// ── 결과 줄 ───────────────────────────────────────────────

const EFFECT_NAMES = { none: '없음', fade: '페이드', push: '밀어내기', wipe: '닦아내기', split: '나누기', cover: '덮기' };
const SHAPE_NAMES = { rectangle: '사각형', rounded_rectangle: '둥근 사각형', oval: '타원', arrow: '화살표', line: '선' };
const CHART_NAMES = { column: '세로 막대', bar: '가로 막대', line: '꺾은선', pie: '원형' };
const slidesText = (r) => (r.all ? `모든 슬라이드(${list(r.slides).length}장)` : `슬라이드 ${list(r.slides).join(', ')}`);
const pt = (v) => Math.round(Number(v) * 10) / 10;

function describe(r, o) {
  switch (r.op) {
    case 'add_slide': return `슬라이드 ${r.slide} 추가 (레이아웃: ${r.layout}${r.created ? ' — 문서에 없어서 새로 만듦' : ''}${o.title ? `, 제목: ${short(o.title)}` : ''}) — 전체 ${r.count}장`;
    case 'delete_slide': return `슬라이드 ${r.slide} 지움 — 전체 ${r.count}장 (뒤 슬라이드 번호가 하나씩 당겨짐)`;
    case 'duplicate_slide': return `슬라이드 ${r.from} 복제 → 새 슬라이드 ${r.slide} — 전체 ${r.count}장`;
    case 'move_slide': return `슬라이드 ${r.from} → ${r.slide}번 자리로 옮김`;
    case 'set_title': return `슬라이드 ${r.slide} 제목: ${short(o.text)}`;
    case 'set_body': return `슬라이드 ${r.slide} 본문('${r.shape}') ${r.lines}문단`;
    case 'set_notes': return `슬라이드 ${r.slide} 발표자 노트 ${String(o.text).length}자`;
    case 'set_text': return `슬라이드 ${r.slide} 도형 '${r.shape}' 글: ${short(o.text)}`;
    case 'add_textbox': return `슬라이드 ${r.slide}에 글상자 '${r.shape}' 추가 (위치 ${pt(r.left)},${pt(r.top)} 크기 ${pt(r.width)}×${pt(r.height)}pt`
      + `${r.grown ? ` — 글이 길어 높이를 ${pt(r.asked)}pt에서 늘림` : ''})`;
    case 'add_image': return `슬라이드 ${r.slide}에 그림 '${r.shape}' 추가 (${basename(o.image_path)}, 위치 ${pt(r.left)},${pt(r.top)} 크기 ${pt(r.width)}×${pt(r.height)}pt)`;
    case 'add_table': return `슬라이드 ${r.slide}에 표 '${r.shape}' 추가 (${r.rows}행×${r.cols}열)`;
    case 'table_set_cell': return `슬라이드 ${r.slide} 표 '${r.shape}' ${r.row}행 ${r.col}열: ${short(o.text)}`;
    case 'add_shape': return `슬라이드 ${r.slide}에 ${SHAPE_NAMES[o.shape] ?? o.shape} '${r.shape}' 추가`;
    case 'add_chart': return `슬라이드 ${r.slide}에 ${CHART_NAMES[o.type] ?? o.type} 차트 '${r.shape}' 추가 (항목 ${r.categories}개, 계열 ${r.series}개${o.type === 'pie' && r.series > 1 ? ' — 원형은 첫 계열만 그림' : ''})`;
    case 'format_text': return `'${short(o.find)}' ${r.count}곳 서식 바꿈 (슬라이드 ${list(r.slides).join(', ')})`;
    case 'delete_shape': return `슬라이드 ${r.slide} 도형 '${r.shape}' 지움`;
    case 'set_background': return `배경색 ${o.color}: ${slidesText(r)}`;
    case 'set_transition': return `화면 전환 ${EFFECT_NAMES[o.effect] ?? o.effect}${o.duration_s && o.effect !== 'none' ? ` ${o.duration_s}초` : ''}: ${slidesText(r)}`;
    case 'hide_slide': return r.hidden
      ? `슬라이드 ${r.slide} 숨김 (슬라이드 쇼·office_convert PDF에서는 빠짐. office_page_image는 슬라이드 번호 그대로 보여 줌)`
      : `슬라이드 ${r.slide} 다시 보이게 함`;
    case 'apply_theme': return `테마 적용: ${basename(o.theme_path)} (디자인 '${r.design}', 레이아웃 ${r.layouts}개)`;
    case 'set_slide_size': {
      const was = Number(r.oldWidth);
      const how = !was || Math.abs(was - r.width) < 0.5 ? ' — 크기 그대로'
        : was < r.width ? ' — 넓어진 만큼 기존 내용은 크기 그대로 가운데로 옮겨짐(양옆이 빔)' : ' — 기존 내용은 비율에 맞춰 줄어듦';
      return `슬라이드 크기 ${o.preset} (${pt(r.width)}×${pt(r.height)}pt)${how}`;
    }
    default: return `${r.op} 완료`;
  }
}

// ── 공개 함수 ─────────────────────────────────────────────

function mustBePpt(path) {
  path = resolve(path);
  if (!existsSync(path)) throw new Error(`파일이 없습니다: ${path}`);
  if (!PPT_EXTS.includes(extOf(path))) throw new Error(`PowerPoint 문서(.pptx/.ppt/.odp)가 아닙니다: ${path}`);
  return path;
}

async function needPowerPoint(path) {
  if (!(await comAvailableFor(extOf(path)))) throw new Error(NEED_POWERPOINT);
}

// ── 암호 확인 ─────────────────────────────────────────────
// 암호가 걸린 문서는 PowerPoint가 보이지 않는 암호 창에서 멈춘다 (DisplayAlerts로도 안 막힘) → 열기 전에 거른다.

export const LOCKED_OPEN = '암호가 걸린 문서는 열 수 없습니다 — 암호를 푼 뒤 다시 시도하세요';
export const LOCKED_PPT = '암호가 걸린 .ppt 문서는 열 수 없습니다 (열기 암호 또는 쓰기 암호) — PowerPoint에서 암호를 푼 뒤 다시 시도하세요';
export const LOCKED_WRITE = '쓰기 암호가 걸린 문서는 바로 고칠 수 없습니다 — output(명령줄은 -o)으로 다른 파일에 저장하세요';
const BIG_FILE = 1024 * 1024 * 1024; // 이보다 크면 내용 확인을 건너뛴다
const MODIFY_VERIFIER = /<(\w+:)?modifyVerifier\b[^>]*?(?:\/>|>[\s\S]*?<\/\1modifyVerifier>)/;

// OOXML 꾸러미의 본문(presentation.xml) 부분 이름
async function mainPartName(zip) {
  const rels = await zip.file('_rels/.rels')?.async('string');
  const target = rels?.match(/<Relationship\b[^>]*Type="[^"]*\/officeDocument"[^>]*>/)?.[0].match(/Target="([^"]+)"/)?.[1];
  return target ? target.replace(/^\//, '') : 'ppt/presentation.xml';
}

// .ppt: 'Current User' 스트림의 headerToken 0xF3D1C4DF = 암호화됨 (쓰기 암호만 있어도 기본 암호로 암호화된다)
function pptBinaryEncrypted(buf) {
  const token = Buffer.from([0x14, 0x00, 0x00, 0x00, 0xdf, 0xc4, 0xd1, 0xf3]);
  for (let i = buf.indexOf(token); i >= 0; i = buf.indexOf(token, i + 1)) {
    if (i >= 6 && buf[i - 6] === 0xf6 && buf[i - 5] === 0x0f) return true; // CurrentUserAtom 레코드 머리 (recType 0x0FF6)
  }
  return false;
}

/**
 * 암호 확인 → 'open'(열기 암호) | 'ppt'(.ppt 암호, 열기·쓰기 구분 못 함) | 'write'(쓰기 암호, pptx) | null
 * 내용을 못 읽는 파일은 null (PowerPoint가 알려 준다).
 */
export async function pptProtection(path) {
  const head = headBytes(path, 8);
  let size = 0;
  try { size = statSync(path).size; } catch { return null; }
  if (size > BIG_FILE) return null;
  if (startsWith(head, CFB)) {
    const buf = await readFile(path);
    if (buf.includes(Buffer.from('EncryptedPackage', 'utf16le'))) return 'open'; // 암호화된 pptx·odp 꾸러미
    return pptBinaryEncrypted(buf) ? 'ppt' : null;
  }
  if (!startsWith(head, ZIP)) return null;
  let zip;
  try { zip = await JSZip.loadAsync(await readFile(path)); } catch { return null; }
  const manifest = await zip.file('META-INF/manifest.xml')?.async('string'); // odp
  if (manifest && /<(?:\w+:)?encryption-data\b/.test(manifest)) return 'open';
  const main = await zip.file(await mainPartName(zip))?.async('string');
  if (main && MODIFY_VERIFIER.test(main)) return 'write';
  return null;
}

/** 열 수 없는 암호 문서면 한국어 오류. inPlace(원본에 저장)면 쓰기 암호도 거절 */
async function refuseLocked(path, { inPlace = false } = {}) {
  const lock = await pptProtection(path);
  if (lock === 'open') throw new Error(LOCKED_OPEN);
  if (lock === 'ppt') throw new Error(LOCKED_PPT);
  if (lock === 'write' && inPlace) throw new Error(LOCKED_WRITE);
  return lock;
}

// 쓰기 암호 pptx → 쓰기 암호만 뺀 임시 사본 (PowerPoint가 쓰기 암호 창 없이 열도록). verifier: 원본의 쓰기 암호 요소
async function writableCopy(path) {
  const zip = await JSZip.loadAsync(await readFile(path));
  const name = await mainPartName(zip);
  const xml = await zip.file(name).async('string');
  const verifier = xml.match(MODIFY_VERIFIER)[0];
  zip.file(name, xml.replace(MODIFY_VERIFIER, ''));
  const dir = mkdtempSync(join(tmpdir(), 'officekit-ppt-'));
  const file = join(dir, basename(path));
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  return { dir, file, verifier };
}

// 결과 pptx에 원본의 쓰기 암호를 다시 건다 (p:presentation의 extLst 앞, 없으면 끝)
async function putVerifier(output, verifier) {
  const zip = await JSZip.loadAsync(await readFile(output));
  const name = await mainPartName(zip);
  let xml = await zip.file(name).async('string');
  if (MODIFY_VERIFIER.test(xml)) return;
  const ext = xml.match(/<p:extLst>(?:(?!<p:extLst>)[\s\S])*<\/p:extLst>\s*<\/p:presentation>\s*$/);
  const at = ext ? ext.index : xml.lastIndexOf('</p:presentation>');
  if (at < 0) throw new Error('presentation.xml 끝을 찾지 못함');
  xml = xml.slice(0, at) + verifier + xml.slice(at);
  zip.file(name, xml);
  writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

const samePath = (a, b) => (process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b);

/**
 * PowerPoint 문서 살펴보기 → { slideWidth, slideHeight, slideCount, designs, layouts:[{design, index, name, type, key}],
 *   offset, limit, slides:[{index, layout, key, title, hidden, effect, duration, notes, shapeCount,
 *   shapes:[{index, name, type, placeholder, left, top, width, height, text, rows, cols, chart, items}]}] }
 * slide를 주면 그 슬라이드만 자세히 (layouts·크기는 그대로). 아니면 offset(건너뛸 장 수)부터 limit장 (limit 생략 = 끝까지).
 * 슬라이드마다 도형은 앞 60개만 (slide를 주면 500개), 나머지는 shapeCount로 개수만.
 */
export async function pptInspect(path, { slide, offset = 0, limit } = {}) {
  path = mustBePpt(path);
  await refuseLocked(path);
  await needPowerPoint(path);
  offset = Math.max(0, Math.trunc(Number(offset) || 0));
  limit = limit == null ? null : Math.max(1, Math.trunc(Number(limit) || 1));
  const req = { app: 'ppt', op: 'inspect', path, maxShapes: slide ? MAX_SHAPES_ONE : MAX_SHAPES };
  if (slide) req.slide = slide;
  else { if (offset) req.offset = offset; if (limit) req.limit = limit; }
  let res;
  try {
    res = await comCall(req, { timeoutMs: 300_000 });
  } catch (e) { throw new Error(pptErrorMessage(e.message, { saving: false })); }
  const info = res.inspect ?? {};
  return {
    path,
    slideWidth: Number(info.slideWidth), slideHeight: Number(info.slideHeight), slideCount: Number(info.slideCount),
    designs: list(info.designs).map(String),
    layouts: list(info.layouts).map((l) => ({ ...l, key: l.key || null })),
    slide: slide ?? null, offset: slide ? null : offset, limit: slide ? null : limit,
    slides: list(info.slides).map((s) => {
      const shapes = list(s.shapes);
      return { ...s, key: s.key || null, shapes, shapeCount: Number(s.shapeCount ?? shapes.length) };
    }),
  };
}

/**
 * PowerPoint 편집: operations를 순서대로 모두 적용한 뒤 한 번만 저장한다. 하나라도 실패하면 저장하지 않는다.
 * 쓰기 암호가 걸린 pptx는 다른 파일(output)로만 저장한다 (결과가 pptx면 같은 쓰기 암호를 다시 건다).
 * → { engine:'powerpoint', results:[한국어 한 줄...], notes:[참고 글], output, raw:[작업별 결과] }
 */
export async function pptEdit(path, operations, { output } = {}) {
  path = mustBePpt(path);
  output = resolve(output ?? path);
  if (!PPT_EXTS.includes(extOf(output))) throw new Error(`결과는 .pptx/.ppt/.odp로만 저장할 수 있습니다: ${output}`);
  const ops = validatePptOps(operations);
  const lock = await refuseLocked(path, { inPlace: samePath(output, path) });
  await needPowerPoint(path);
  const copy = lock === 'write' ? await writableCopy(path) : null;
  let res;
  try {
    res = await comCall({ app: 'ppt', op: 'edit', path: copy?.file ?? path, operations: ops, save: output, saveFormat: extOf(output) }, { timeoutMs: 300_000 });
  } catch (e) {
    throw new Error(pptErrorMessage(e.message, { output }));
  } finally {
    if (copy) rmSync(copy.dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
  const notes = [];
  if (copy) {
    if (extOf(output) === 'pptx') {
      try {
        await putVerifier(output, copy.verifier);
        notes.push('참고: 원본처럼 결과 파일에도 쓰기 암호가 걸려 있습니다 (다시 고칠 때도 output으로 다른 파일에 저장)');
      } catch (e) { notes.push(`참고: 결과 파일에 원본의 쓰기 암호를 다시 걸지 못했습니다 (${e.message})`); }
    } else {
      notes.push(`참고: 결과(.${extOf(output)}) 파일에는 원본의 쓰기 암호가 없습니다`);
    }
  }
  const raw = list(res.results);
  return { engine: 'powerpoint', results: raw.map((r, i) => `${i + 1}. ${describe(r, ops[i])}`), notes, output, raw };
}

// ── 살펴보기 글 ───────────────────────────────────────────

const SHAPE_TYPES = { 1: '도형', 3: '차트', 5: '자유형', 6: '그룹', 7: 'OLE 개체', 9: '선', 10: '연결된 OLE 개체', 11: '연결된 그림',
  12: '컨트롤', 13: '그림', 14: '개체 틀', 15: 'WordArt', 16: '미디어', 17: '글상자', 19: '표', 24: 'SmartArt', 28: '3D 모델' };
const PLACEHOLDERS = { 1: '제목', 2: '본문', 3: '가운데 제목', 4: '부제목', 5: '세로 제목', 6: '세로 본문', 7: '내용', 8: '차트', 9: '그림',
  10: '미디어', 11: '조직도', 12: '표', 13: '슬라이드 번호', 14: '머리글', 15: '바닥글', 16: '날짜', 17: '세로 내용', 18: '그림' };
const EFFECTS = { 0: null, 3849: 'fade', 1793: 'fade(검게)', 3852: 'push', 3853: 'push', 3854: 'push', 3855: 'push',
  2817: 'wipe', 2818: 'wipe', 2819: 'wipe', 2820: 'wipe', 3585: 'split', 3586: 'split', 3587: 'split', 3588: 'split',
  1281: 'cover', 1282: 'cover', 1283: 'cover', 1284: 'cover', 1285: 'cover', 1286: 'cover', 1287: 'cover', 1288: 'cover' };
// XlChartType 중 흔한 것
const CHART_KINDS = { 51: '세로 막대', 52: '누적 세로 막대', 57: '가로 막대', 58: '누적 가로 막대', 4: '꺾은선', 65: '표식 꺾은선',
  5: '원형', '-4120': '도넛', 1: '영역', '-4169': '분산형', 15: '거품형', '-4151': '방사형' };
const ratio = (w, h) => (Math.abs(w / h - 16 / 9) < 0.01 ? '16:9' : Math.abs(w / h - 4 / 3) < 0.01 ? '4:3' : Math.abs(w / h - 16 / 10) < 0.01 ? '16:10' : null);
const round = (v) => Math.round(Number(v) || 0);

export function effectName(code) {
  if (!code) return null;
  return EFFECTS[code] ?? `기타(${code})`;
}

function slideLines(s, single) {
  const eff = effectName(s.effect);
  const out = ['', `[슬라이드 ${s.index}] 레이아웃: ${s.layout}${s.key ? ` [${s.key}]` : ''}${s.hidden ? ', 숨김' : ''}${eff ? `, 화면 전환: ${eff}${s.duration ? ` ${s.duration}초` : ''}` : ''}`];
  if (s.title) out.push(`  제목: ${s.title}`);
  if (s.notes) out.push(`  노트: ${s.notes.replace(/\n/g, ' / ')}`);
  const total = Number(s.shapeCount ?? s.shapes.length);
  out.push(`  도형 ${total}개:`);
  for (const sh of s.shapes) {
    let kind = SHAPE_TYPES[sh.type] ?? `종류 ${sh.type}`;
    if (sh.placeholder) kind = `개체 틀(${PLACEHOLDERS[sh.placeholder] ?? sh.placeholder})`;
    if (sh.rows) kind = `${sh.placeholder ? `${kind} ` : ''}표 ${sh.rows}행×${sh.cols}열`;
    if (sh.chart) kind = `${sh.placeholder ? `${kind} ` : ''}차트(${CHART_KINDS[sh.chart] ?? sh.chart})`;
    if (sh.items) kind += ` ${sh.items}개 묶음`;
    const txt = sh.text ? `: ${String(sh.text).replace(/\n/g, ' / ')}` : '';
    out.push(`    ${sh.index}. '${sh.name}' ${kind} @ ${round(sh.left)},${round(sh.top)} ${round(sh.width)}×${round(sh.height)}pt${txt}`);
  }
  if (total > s.shapes.length) {
    out.push(`    …외 ${total - s.shapes.length}개${single ? '' : ` (slide=${s.index}로 이 슬라이드만 보면 ${MAX_SHAPES_ONE}개까지 나옴)`}`);
  }
  return out;
}

/** pptInspect 결과 → AI가 읽기 좋은 한국어 글. 나눠 보기(limit)면 글자 수 한도 안에서 끊고 이어 볼 offset을 알려 준다 */
export function formatPptInspect(info) {
  const r = ratio(info.slideWidth, info.slideHeight);
  const out = [
    `${basename(info.path ?? '')} — 슬라이드 ${info.slideCount}장, 크기 ${round(info.slideWidth)}×${round(info.slideHeight)}pt${r ? ` (${r})` : ''}, 디자인: ${info.designs.join(', ')}`,
    '',
    '[쓸 수 있는 레이아웃] add_slide의 layout에 키 또는 이름을 쓴다',
  ];
  for (const l of info.layouts) out.push(`  - ${l.name}${l.key ? ` [${l.key}]` : ''}${info.designs.length > 1 ? ` (디자인: ${l.design})` : ''}`);
  const missing = LAYOUT_KEYS.filter((k) => !info.layouts.some((l) => l.key === k));
  if (missing.length) out.push(`  (없는 키 ${missing.join(', ')}는 add_slide에서 쓰면 표준 레이아웃을 새로 만든다)`);
  const paging = info.limit != null;
  const budget = paging ? INSPECT_BUDGET : Infinity;
  let used = out.join('\n').length;
  let last = null;
  for (const s of info.slides) {
    const lines = slideLines(s, Boolean(info.slide));
    const size = lines.join('\n').length + 1;
    if (last !== null && used + size > budget) break; // 한 장은 늘 보여 준다
    out.push(...lines);
    used += size;
    last = s.index;
  }
  const first = info.slides[0]?.index;
  if (!info.slide && !info.slides.length && info.offset && info.slideCount) {
    out.push('', `(슬라이드는 ${info.slideCount}장뿐입니다 — offset은 ${info.slideCount - 1} 이하로)`);
  } else if (!info.slide && last !== null && (last < info.slideCount || first > 1)) {
    out.push('', `(슬라이드 ${first}~${last}만 보였습니다 (전체 ${info.slideCount}장)${last < info.slideCount ? ` — 이어서 보려면 offset=${last}` : ''})`);
  }
  return out.join('\n');
}
