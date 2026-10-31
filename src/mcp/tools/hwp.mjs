// tools/hwp.mjs — 한글(hwp/hwpx) 전용 MCP 도구. 프로그램 없이 rhwp로 살펴보고 편집한다 (docs/FEATURES.md).
import { existsSync, rmdirSync, statSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { z } from 'zod';
import { Hwp, extOf, isHwp } from '../../core/index.mjs';
import { PLACEHOLDER } from '../../core/officefile.mjs';

// ── 입력 검사 (zod) — 오류는 한국어, 작업 객체는 엄격(모르는 인수 이름은 거절) ───
// 스키마마다 한국어 오류 함수를 붙인다: MCP SDK의 입력 검증과 명령줄(hwpOps)이 같은 메시지를 쓴다
const KO_LOCALE = z.locales.ko().localeError;
const TYPE_KO = { number: '숫자', string: '글자', boolean: 'true 또는 false', array: '배열', object: '객체' };
// 빠진 인수는 '값이 필요합니다'로, 나머지는 zod 한국어 로캘 그대로
const KO = (iss) => (iss.code === 'invalid_type' && iss.input === undefined
  ? `값이 필요합니다 (${TYPE_KO[iss.expected] ?? iss.expected})` : KO_LOCALE(iss));
const E = { error: (iss) => KO(iss) };
const msg = (m) => ({ error: () => m });

const str = () => z.string(E);
const text1 = () => z.string(E).min(1, msg('빈 글은 안 됩니다'));
const bool = () => z.boolean(E);
const scalar = z.union([z.string(E), z.number(E)], msg('글자 또는 숫자여야 합니다'));
// 번호(1부터)·크기의 범위 검사는 core(Hwp)가 한국어 오류로 한다 — 스키마는 숫자인지만 본다
const num = z.number(E);

/** 엄격한 객체: 모르는 인수 이름이면 쓸 수 있는 인수를 알려 준다 */
function strict(shape, what) {
  const keys = Object.keys(shape).filter((k) => k !== 'op');
  return z.strictObject(shape, {
    error: (iss) => (iss.code === 'unrecognized_keys'
      ? `모르는 인수입니다: ${iss.keys.join(', ')} — ${what}에 쓸 수 있는 인수: ${keys.join(', ') || '없음'}`
      : KO(iss)),
  });
}
const op = (name, shape, what) => strict({ op: z.literal(name, E), ...shape }, name).describe(what);

const ALIGNS = ['left', 'center', 'right', 'justify', 'distribute'];
const alignSchema = z.enum(ALIGNS, msg(`정렬은 ${ALIGNS.join('·')} 중 하나여야 합니다`));

const styleSchema = strict({
  bold: bool().optional(),
  italic: bool().optional(),
  underline: bool().optional(),
  size: num.positive(msg('글자 크기(pt)는 0보다 커야 합니다')).optional().describe('글자 크기(pt)'),
  color: str().optional().describe("글자색 '#rrggbb'"),
  font: str().optional().describe('글꼴 이름'),
  align: z.enum(['left', 'center', 'right', 'justify'], msg('정렬은 left·center·right·justify 중 하나여야 합니다')).optional(),
}, 'style').optional();

// 글자 서식 (format_text와 insert_paragraph.format 공통)
const charFields = {
  bold: bool().optional(),
  italic: bool().optional(),
  underline: bool().optional(),
  strike: bool().optional().describe('취소선'),
  size: num.optional().describe('글자 크기(pt)'),
  color: str().optional().describe("글자색 '#rrggbb'"),
  font: str().optional().describe('글꼴 이름 (예: 맑은 고딕)'),
  highlight: str().optional().describe("글자 배경색(형광펜) '#rrggbb'"),
  superscript: bool().optional().describe('위 첨자'),
  subscript: bool().optional().describe('아래 첨자'),
};
const CHAR_KEYS = Object.keys(charFields);

const AT_HELP = "'end'(기본) | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'찾을 글'}";
const atSchema = z.union([
  z.enum(['end', 'start'], E),
  z.strictObject({ after_paragraph: num }, E),
  z.strictObject({ before_paragraph: num }, E),
  z.strictObject({ after_text: text1() }, E),
], msg(`위치(at)는 ${AT_HELP} 중 하나여야 합니다`)).optional().describe(`넣을 위치: ${AT_HELP}`);
const rangeSchema = z.tuple([num, num], msg('paragraphs는 [처음, 끝] 두 숫자여야 합니다')).optional().describe('문단 범위 [처음, 끝] (1부터, 끝 포함)');
const paraNo = num.optional().describe('문단 번호 (1부터, hwp_inspect의 번호)');
const findText = text1().optional().describe('이 글이 든 첫 본문 문단 (대소문자 구분)');

const OPS = [
  // 기존 (문서 끝에 붙이기)
  op('add_paragraph', { text: str(), style: styleSchema }, '문서 끝에 문단 추가'),
  op('add_heading', { text: str(), level: z.number(E).int(msg('level은 1~3 정수여야 합니다')).min(1, msg('level은 1~3 정수여야 합니다')).max(3, msg('level은 1~3 정수여야 합니다')).default(1) }, '문서 끝에 제목 추가 (1~3단계)'),
  op('add_table', { rows: z.array(z.array(scalar, E).min(1, msg('행마다 칸이 하나 이상 있어야 합니다')), E).min(1, msg('rows에 행이 하나 이상 있어야 합니다')) }, '문서 끝에 표 추가. 첫 행은 굵게'),
  op('add_page_break', {}, '문서 끝에 쪽 나누기'),
  op('add_field', { name: text1(), label: str().optional() }, '문서 끝에 누름틀 추가 ("라벨: [누름틀]")'),
  op('replace', { find: text1(), replace: str() }, '문서 전체에서 찾아 바꾸기 (영문 대소문자 구분)'),
  op('fill', { values: z.record(str(), scalar, E) }, '누름틀 채우기 (같은 이름은 모두, 없는 이름은 {{이름}} 자리표시자를 바꿈)'),
  // 문단
  op('insert_paragraph', {
    text: str().describe('문단 글 (줄바꿈마다 문단 하나)'), at: atSchema,
    style: str().optional().describe("문서 스타일 이름 (예: '바탕글', '본문', '개요 1'). 생략하면 '바탕글'"),
    like_paragraph: num.optional().describe('이 번호 문단의 모양(스타일·문단·글자 모양)을 그대로 따름'),
    format: strict(charFields, 'format').optional().describe('글자 서식 (format_text와 같음)'),
    align: alignSchema.optional(),
  }, '원하는 위치에 문단 넣기'),
  op('set_paragraph_text', { paragraph: paraNo, find: findText, text: str() }, '문단 글 바꾸기 (모양 유지)'),
  op('delete_paragraph', { paragraph: paraNo, find: findText }, '문단 지우기'),
  op('format_text', { find: text1(), all: bool().optional(), ...charFields }, '글자 서식'),
  op('format_paragraph', {
    paragraph: paraNo, paragraphs: rangeSchema, find: findText,
    align: alignSchema.optional(),
    line_spacing: z.union([z.number(E), z.string(E)], msg("줄 간격은 배수(1.6) 또는 '160%'여야 합니다")).optional().describe("줄 간격: 배수(1.6) 또는 '160%'"),
    left_indent_mm: num.optional().describe('왼쪽 여백(mm)'),
    right_indent_mm: num.optional().describe('오른쪽 여백(mm)'),
    first_line_indent_mm: num.optional().describe('첫 줄 들여쓰기(mm, 음수면 내어쓰기)'),
    space_before: num.optional().describe('문단 위 간격(pt)'),
    space_after: num.optional().describe('문단 아래 간격(pt)'),
    keep_with_next: bool().optional(),
    page_break_before: bool().optional().describe('문단 앞에서 쪽 나누기'),
  }, '문단 서식'),
  op('apply_style', { paragraph: paraNo, paragraphs: rangeSchema, find: findText, style: text1() }, '문서 스타일 적용'),
  // 표
  op('table_set_cell', { table: num, row: num, col: num, text: scalar }, '기존 표 칸 글 바꾸기'),
  op('table_add_row', { table: num, after_row: num.optional(), values: z.array(scalar, E).optional() }, '표에 행 추가'),
  op('table_delete_row', { table: num, row: num }, '표 행 지우기'),
  // 머리말·각주·그림·쪽
  op('set_header', { text: str(), page_numbers: bool().optional(), align: alignSchema.optional() }, '머리말'),
  op('set_footer', { text: str(), page_numbers: bool().optional(), align: alignSchema.optional() }, '꼬리말'),
  op('add_footnote', { find: text1(), text: str() }, '각주'),
  op('insert_image', {
    image_path: text1(), at: atSchema,
    width_mm: num.optional(), height_mm: num.optional(),
    align: alignSchema.optional(), description: str().optional(),
  }, '그림 넣기'),
  op('page_setup', {
    orientation: z.enum(['portrait', 'landscape'], msg('orientation은 portrait 또는 landscape여야 합니다')).optional(),
    paper: z.enum(['A4', 'A3', 'A5', 'B4', 'B5', 'Letter', 'Legal'], msg('paper는 A4·A3·A5·B4·B5·Letter·Legal 중 하나여야 합니다')).optional(),
    margins_mm: strict({
      top: num.optional(), bottom: num.optional(), left: num.optional(), right: num.optional(),
      header: num.optional(), footer: num.optional(), gutter: num.optional(),
    }, 'margins_mm').optional(),
    section: num.optional().describe('구역 번호 (생략하면 모든 구역)'),
  }, '쪽 설정'),
];
export const HWP_OP_NAMES = OPS.map((o) => o.shape.op.value);

export const hwpOp = z.discriminatedUnion('op', OPS, {
  error: (iss) => (iss.code === 'invalid_union' && iss.note === 'No matching discriminator'
    ? `알 수 없는 작업(op)입니다${typeof iss.input?.op === 'string' ? `: ${iss.input.op}` : ''}. 가능: ${HWP_OP_NAMES.join(', ')}`
    : KO(iss)),
});

const short = (t, n = 40) => { const s = String(t ?? '').replace(/\s+/g, ' ').trim(); return s.length > n ? `${s.slice(0, n)}…` : s; };

// 대상 문단: paragraph | paragraphs:[처음,끝] | find(글이 든 첫 본문 문단). paragraphs는 범위를 받는 작업만
function targetRange(h, op) {
  const keys = op.op === 'set_paragraph_text' || op.op === 'delete_paragraph' ? ['paragraph', 'find'] : ['paragraph', 'paragraphs', 'find'];
  const given = keys.filter((k) => op[k] !== undefined);
  if (given.length !== 1) {
    throw new Error(`대상 문단을 ${keys.map((k) => (k === 'paragraphs' ? 'paragraphs:[처음,끝]' : k)).join(', ')} 중 하나로만 지정하세요${given.length > 1 ? ` (함께 줌: ${given.join(', ')})` : ''}`);
  }
  if (op.paragraph !== undefined) return [op.paragraph, op.paragraph];
  if (op.paragraphs !== undefined) return op.paragraphs;
  const n = h.findParagraph(op.find);
  return [n, n];
}
const rangeText = ([a, b]) => (a === b ? `문단 ${a}` : `문단 ${a}~${b}`);

const CHAR_KO = { bold: '굵게', italic: '기울임', underline: '밑줄', strike: '취소선', superscript: '위 첨자', subscript: '아래 첨자' };
function charSummary(f) {
  const out = [];
  for (const [k, v] of Object.entries(f)) {
    if (v === undefined) continue;
    if (CHAR_KO[k]) out.push(v ? CHAR_KO[k] : `${CHAR_KO[k]} 없앰`);
    else if (k === 'size') out.push(`${v}pt`);
    else if (k === 'color') out.push(`글자색 ${v}`);
    else if (k === 'font') out.push(v);
    else if (k === 'highlight') out.push(`배경 ${v}`);
  }
  return out.join(', ');
}
const pick = (o, keys) => Object.fromEntries(keys.filter((k) => o[k] !== undefined).map((k) => [k, o[k]]));

const ALIGN_KO = { left: '왼쪽', center: '가운데', right: '오른쪽', justify: '양쪽', distribute: '배분' };
const MARGIN_KO = { top: '위', bottom: '아래', left: '왼쪽', right: '오른쪽', header: '머리말', footer: '꼬리말', gutter: '제본' };
function paraSummary(f) {
  const out = [];
  if (f.align) out.push(`정렬 ${ALIGN_KO[f.align] ?? f.align}`);
  if (f.line_spacing !== undefined) {
    const s = String(f.line_spacing).trim();
    out.push(`줄 간격 ${s.endsWith('%') ? s : `${Math.round(Number(s) * 100)}%`}`);
  }
  if (f.left_indent_mm !== undefined) out.push(`왼쪽 여백 ${f.left_indent_mm}mm`);
  if (f.right_indent_mm !== undefined) out.push(`오른쪽 여백 ${f.right_indent_mm}mm`);
  if (f.first_line_indent_mm !== undefined) {
    const v = Number(f.first_line_indent_mm);
    out.push(v < 0 ? `내어쓰기 ${-v}mm` : `첫 줄 들여쓰기 ${v}mm`);
  }
  if (f.space_before !== undefined) out.push(`문단 위 ${f.space_before}pt`);
  if (f.space_after !== undefined) out.push(`문단 아래 ${f.space_after}pt`);
  if (f.keep_with_next !== undefined) out.push(f.keep_with_next ? '다음 문단과 함께' : '다음 문단과 함께 안 함');
  if (f.page_break_before !== undefined) out.push(f.page_break_before ? '앞에서 쪽 나눔' : '앞 쪽 나눔 없앰');
  return out.join(', ');
}

/**
 * 작업 하나 적용 → 결과 한 줄 (한국어). 실패하면 예외.
 * ctx.resolvePath: 그림 경로를 절대 경로로 (MCP 서버는 작업 폴더 기준)
 */
export function applyHwpOp(h, op, { resolvePath = (p) => resolve(p) } = {}) {
  switch (op.op) {
    case 'add_paragraph': h.addParagraph(op.text, op.style ?? {}); return `문단 추가: ${op.text.slice(0, 40)}`;
    case 'add_heading': h.addHeading(op.text, op.level ?? 1); return `제목${op.level ?? 1} 추가: ${op.text.slice(0, 40)}`;
    case 'add_table': {
      const rows = op.rows.map((r) => r.map((v) => String(v)));
      h.addTable(rows);
      return `표 추가: ${rows.length}행×${Math.max(...rows.map((r) => r.length))}열`;
    }
    case 'add_page_break': h.addPageBreak(); return '쪽 나누기';
    case 'add_field': h.addField(op.name, { label: op.label }); return `누름틀 추가: ${op.name}`;
    case 'replace': return `바꾸기 '${op.find}' → '${op.replace}': ${h.replace(op.find, op.replace, { caseSensitive: true })}곳`;
    case 'fill': {
      const own = new Set(h.fields().map((f) => f.name));
      // {{ 이름 }}처럼 띄어 쓴 자리표시자도 문서에 쓰인 표기 그대로 바꾼다
      const spell = new Map();
      for (const m of h.text().matchAll(PLACEHOLDER)) { if (!spell.has(m[1])) spell.set(m[1], new Set()); spell.get(m[1]).add(m[0]); }
      const filled = []; const missing = [];
      for (const [name, value] of Object.entries(op.values)) {
        if (own.has(name)) {
          h.fill({ [name]: String(value) });
          const k = h.lastFillCounts?.[name] ?? 1;
          filled.push(k > 1 ? `${name}(${k}곳)` : name);
          continue;
        }
        let n = 0;
        for (const raw of spell.get(name) ?? []) n += h.replace(raw, String(value), { caseSensitive: true });
        (n > 0 ? filled : missing).push(name);
      }
      return `채우기: ${filled.join(', ') || '없음'}${missing.length ? ` / 못 찾음: ${missing.join(', ')}` : ''}`;
    }

    case 'insert_paragraph': {
      const r = h.insertParagraph(op.text, { at: op.at, style: op.style, like: op.like_paragraph, format: op.format, align: op.align });
      const last = r.paragraph + r.count - 1;
      return `문단 넣기 → ${rangeText([r.paragraph, last])}: ${short(op.text) || '(빈 문단)'} (뒤 문단 번호 +${r.count})`;
    }
    case 'set_paragraph_text': {
      const [n] = targetRange(h, op);
      const count = h.setParagraphText(n, op.text);
      return `문단 ${n} 글 바꿈: ${short(op.text) || '(빈 글)'}${count > 1 ? ` (${rangeText([n + 1, n + count - 1])} 새로 생김, 뒤 문단 번호 +${count - 1})` : ''}`;
    }
    case 'delete_paragraph': {
      const [n] = targetRange(h, op);
      const before = h.tables();
      const gone = before.map((t, i) => (t.paragraph === n ? i + 1 : 0)).filter(Boolean);
      const with_ = h.deleteParagraph(n);
      // 함께 지운 개체 설명의 '표'/'표 2개'를 표 번호로
      const what = with_ ? with_.split(', ').map((w) => (gone.length && /^표( \d+개)?$/.test(w) ? `표 ${gone.join('·')}` : w)).join(', ') : '';
      const later = gone.length && before.length > gone.at(-1) ? `, 뒤 표 번호 -${gone.length}` : '';
      return `문단 ${n} 지움${what ? ` (${what}도 함께 지움)` : ''} — 뒤 문단 번호 -1${later}`;
    }
    case 'format_text': {
      const f = pick(op, CHAR_KEYS);
      const count = h.formatText(op.find, f, { all: !!op.all });
      return `글자 서식 '${short(op.find, 30)}' ${count}곳: ${charSummary(f)}`;
    }
    case 'format_paragraph': {
      const range = targetRange(h, op);
      const f = pick(op, ['align', 'line_spacing', 'left_indent_mm', 'right_indent_mm', 'first_line_indent_mm', 'space_before', 'space_after', 'keep_with_next', 'page_break_before']);
      h.formatParagraphs(range[0], range[1], f);
      return `문단 서식 ${rangeText(range)}: ${paraSummary(f)}`;
    }
    case 'apply_style': {
      const range = targetRange(h, op);
      const name = h.applyStyle(range[0], range[1], op.style);
      return `스타일 '${name}' 적용: ${rangeText(range)}`;
    }
    case 'table_set_cell': h.setCell(op.table, op.row, op.col, String(op.text)); return `표 ${op.table} ${op.row}행 ${op.col}열 = ${short(op.text) || '(빈 칸)'}`;
    case 'table_add_row': {
      const r = h.addTableRow(op.table, { afterRow: op.after_row, values: (op.values ?? []).map(String) });
      const vals = (op.values ?? []).map((v, j) => (r.merged.includes(j + 1) ? '(병합)' : short(v, 15)));
      const note = r.merged.length ? ` (${r.merged.map((c) => `${c}열`).join('·')}은 위 칸과 병합된 자리)` : '';
      return `표 ${op.table}에 행 추가 → ${r.row}행${vals.length ? `: ${vals.join(' | ')}` : ''}${note}`;
    }
    case 'table_delete_row': h.deleteTableRow(op.table, op.row); return `표 ${op.table}의 ${op.row}행 지움 — 아래 행 번호 -1`;
    case 'set_header':
    case 'set_footer': {
      const kind = op.op === 'set_header' ? 'header' : 'footer';
      const what = kind === 'header' ? '머리말' : '꼬리말';
      const r = h.setHeaderFooter(kind, op.text, { pageNumbers: !!op.page_numbers, align: op.align });
      if (r === 'removed') return `${what} 지움`;
      return `${what}: ${short(op.text) || '(글 없음)'}${op.page_numbers ? ' + 쪽 번호' : ''}`;
    }
    case 'add_footnote': {
      const r = h.addFootnote(op.find, op.text);
      return `각주 ${r.number}) '${short(op.find, 20)}' 뒤 (문단 ${r.paragraph}): ${short(op.text)}`;
    }
    case 'insert_image': {
      const full = resolvePath(op.image_path);
      if (!existsSync(full)) throw new Error(`그림 파일이 없습니다: ${full}`);
      if (statSync(full).isDirectory()) throw new Error(`그림 파일이 아니라 폴더입니다: ${full}`);
      const r = h.insertImage(full, { at: op.at, widthMm: op.width_mm, heightMm: op.height_mm, align: op.align ?? 'center', description: op.description ?? basename(full) });
      const fit = r.fit === 'height' ? ' (본문 높이에 맞춰 줄임)' : r.shrunk ? ' (본문 너비에 맞춰 줄임)' : '';
      return `그림 넣기 → 문단 ${r.paragraph}: ${basename(full)} ${r.widthMm}×${r.heightMm}mm${fit} (뒤 문단 번호 +1)`;
    }
    case 'page_setup': {
      const n = h.pageSetup({ orientation: op.orientation, paper: op.paper, margins_mm: op.margins_mm, section: op.section });
      const parts = [op.paper, op.orientation ? (op.orientation === 'landscape' ? '가로' : '세로') : null,
        op.margins_mm ? `여백(mm) ${Object.entries(op.margins_mm).filter(([, v]) => v !== undefined).map(([k, v]) => `${MARGIN_KO[k] ?? k} ${v}`).join(', ')}` : null].filter(Boolean);
      return `쪽 설정: ${parts.join(', ')}${op.section ? ` (구역 ${op.section})` : n > 1 ? ` (구역 ${n}개 모두)` : ''}`;
    }
    default: throw new Error(`알 수 없는 작업: ${op.op}`);
  }
}

// ── hwp_inspect 결과 글 ──────────────────────────────────

const ORIENT_KO = { portrait: '세로', landscape: '가로' };
const MAX_TABLES = 50; // 한 번에 보여 줄 표
const MAX_FIELDS = 100; // 한 번에 보여 줄 누름틀
const MAX_OTHER = 10; // 뒤쪽 표 번호 안내

/** Hwp#inspect() 결과 → 사람·AI가 읽는 한국어 글 */
export function inspectText(r, { name = '' } = {}) {
  const L = [];
  L.push(`${name ? `${name} — ` : ''}한글(.${r.format || 'hwp'}) ${r.pages}쪽 · 구역 ${r.sections.length}개 · 문단 ${r.paragraphCount}개 · 표 ${r.tables.length}개`
    + ` · 누름틀 ${r.fields.length}개 · 그림 ${r.counts.pictures}개 · 각주 ${r.counts.footnotes}개`
    + `${r.counts.shapes ? ` · 도형 ${r.counts.shapes}개` : ''}${r.counts.endnotes ? ` · 미주 ${r.counts.endnotes}개` : ''}${r.counts.equations ? ` · 수식 ${r.counts.equations}개` : ''}`);
  L.push('', '[쪽 설정]');
  for (const s of r.sections) {
    const m = s.margins;
    L.push(`  구역 ${s.section + 1}: ${s.paper} ${ORIENT_KO[s.orientation]} ${s.widthMm}×${s.heightMm}mm · 여백(mm) 위 ${m.top}, 아래 ${m.bottom}, 왼쪽 ${m.left}, 오른쪽 ${m.right}, 머리말 ${m.header}, 꼬리말 ${m.footer}${m.gutter ? `, 제본 ${m.gutter}` : ''}`);
  }
  L.push('', '[머리말·꼬리말] (저장된 쪽 번호 자리는 빈칸으로 보일 수 있음)');
  const hf = (list, what) => {
    if (!list.length) { L.push(`  ${what}: 없음`); return; }
    for (const x of list) L.push(`  ${what}(${r.sections.length > 1 ? `구역 ${x.section + 1}, ` : ''}${x.applyTo}): ${x.text.trim() || '(빈 글)'}`);
  };
  hf(r.headers, '머리말'); hf(r.footers, '꼬리말');
  const first = r.paragraphs[0]?.n ?? r.offset + 1;
  const last = r.paragraphs.at(-1)?.n ?? r.offset;
  L.push('', `[문단] ${r.paragraphs.length ? `${first}~${last}` : '없음'} / 전체 ${r.paragraphCount}개 (이 번호를 hwp_edit의 paragraph·at에 그대로 씀)`);
  let sec = 0;
  for (const p of r.paragraphs) {
    if (r.sections.length > 1 && p.section !== sec) { sec = p.section; L.push(`  ── 구역 ${sec} ──`); }
    const look = [p.style, p.align, [p.font, p.size ? `${p.size}pt` : '', p.bold ? '굵게' : '', p.italic ? '기울임' : ''].filter(Boolean).join(' ')];
    if (p.leftIndentMm) look.push(`왼쪽 여백 ${p.leftIndentMm}mm`);
    if (p.lineSpacing && p.lineSpacing !== '160%') look.push(`줄 간격 ${p.lineSpacing}`);
    const text = p.text.replace(/\r?\n/g, ' ');
    const body = [text, p.objects.length ? `[${p.objects.join(', ')}]` : ''].filter(Boolean).join(' ') || '(빈 문단)';
    L.push(`  ${p.n}. ${body}  ‹${look.filter(Boolean).join(' · ')}›`);
  }
  if (last < r.paragraphCount) L.push(`  … 다음 문단은 offset=${last} 로 다시 살펴보세요`);
  // 표·누름틀은 이번에 보여 준 문단 범위 안 것만 (큰 서식 문서가 결과를 넘치게 하지 않도록), 그래도 많으면 자른다
  const lo = r.offset ?? 0;
  const hi = lo + (r.limit ?? Infinity);
  const inRange = (para) => para > lo && para <= hi;
  const span = r.paragraphs.length ? `문단 ${first}~${last}` : '이 범위';
  if (r.tables.length) {
    const win = r.tables.filter((t) => inRange(t.paragraph));
    const rest = r.tables.filter((t) => !inRange(t.paragraph));
    L.push('', `[표] ${span} 안 ${win.length}개 / 전체 ${r.tables.length}개 (hwp_edit의 table 번호, 행·열은 1부터)`);
    for (const t of win.slice(0, MAX_TABLES)) {
      const row = t.firstRow ?? [];
      const more = t.cols > row.length ? ` … (+${t.cols - row.length}칸)` : '';
      L.push(`  표 ${t.n} — 문단 ${t.paragraph}, ${t.rows}행×${t.cols}열, 첫 행: ${row.map((c) => short(c, 20)).join(' | ')}${more}`);
    }
    if (win.length > MAX_TABLES) L.push(`  … 외 ${win.length - MAX_TABLES}개 — limit을 줄이거나 offset을 옮겨 보세요`);
    // 범위 밖 표: 앞쪽은 번호 범위만, 뒤쪽은 가까운 것부터 문단 번호와 함께 (offset=문단 번호-1로 가서 봄)
    const before = rest.filter((t) => t.paragraph <= lo);
    const after = rest.filter((t) => t.paragraph > hi);
    if (before.length) L.push(`  앞쪽 표 ${before.length}개: 표 ${before[0].n}~${before.at(-1).n} (문단 ${before[0].paragraph}~${before.at(-1).paragraph})`);
    if (after.length) {
      const list = after.slice(0, MAX_OTHER).map((t) => `표 ${t.n}(문단 ${t.paragraph})`).join(', ');
      L.push(`  뒤쪽 표 ${after.length}개: ${list}${after.length > MAX_OTHER ? ` … 외 ${after.length - MAX_OTHER}개` : ''} — 보려면 offset=(문단 번호-1)`);
    }
  }
  if (r.fields.length) {
    // 위치를 모르는 누름틀은 첫 범위에서 보여 준다
    const win = r.fields.filter((f) => (f.paragraph ? inRange(f.paragraph) : lo === 0));
    L.push('', `[누름틀] ${span} 안 ${win.length}개 / 전체 ${r.fields.length}개 (fill로 채움)`);
    for (const f of win.slice(0, MAX_FIELDS)) L.push(`  ${f.name} = ${f.value ? short(f.value, 60) : '(빈 칸)'}${f.paragraph ? ` — 문단 ${f.paragraph}${f.cell ? ' (표 칸 안)' : ''}` : ''}`);
    if (win.length > MAX_FIELDS) L.push(`  … 외 ${win.length - MAX_FIELDS}개 — 이름 전체는 office_fields로 보세요`);
    if (win.length < r.fields.length) L.push(`  다른 범위의 누름틀 ${r.fields.length - win.length}개 — offset을 옮겨 보거나 office_fields로 이름 전체를 보세요`);
  }
  L.push('', `[스타일] ${r.styles.join(', ')}`);
  return L.join('\n');
}

// 파일 경로의 아직 없는 상위 폴더들 (깊은 것부터)
function missingDirs(file) {
  const out = [];
  for (let d = dirname(file); !existsSync(d); d = dirname(d)) {
    out.push(d);
    if (dirname(d) === d) break;
  }
  return out;
}

/** server.mjs가 부른다. ctx: 공통 도우미 (tool, text, mustExist, outPath, 스키마 등) */
export function register(server, ctx) {
  const { tool, text, koreanError, abs, mustExist, outPath, pathSchema, outputSchema, overwriteSchema, HWP_EXTS } = ctx;
  // 공통 스키마와 같은 설명에 한국어 입력 오류를 붙인 것
  const pathKo = z.string(E).min(1, msg('문서 경로가 비어 있습니다')).describe(pathSchema('한글 문서 경로 (.hwp/.hwpx/.hwt)').description);
  const outputKo = z.string(E).optional().describe(outputSchema.description);
  const overwriteKo = z.boolean(E).default(false).describe(overwriteSchema.description);
  const LIMIT_MSG = msg('limit은 1~1000 사이 정수여야 합니다');
  const OFFSET_MSG = msg('offset은 0 이상의 정수여야 합니다');

  server.registerTool('hwp_inspect', {
    title: '한글 문서 살펴보기',
    description: '한글(hwp/hwpx/hwt) 문서 구조를 프로그램 없이(rhwp) 살펴본다: 쪽 수, 구역별 쪽 설정(용지·방향·여백 mm), 머리말·꼬리말, '
      + '문단 목록(번호·글 200자·스타일·정렬·글꼴·표/그림/누름틀 표시), 표 목록(번호·행×열·첫 행), 누름틀, 스타일 이름. '
      + '문단 번호는 모든 구역의 본문 문단을 문서 순서대로 1부터 센 것(표 칸·머리말·각주 안 문단은 세지 않음)이며 hwp_edit에 그대로 쓴다. '
      + '긴 문서는 offset(건너뛸 문단 수)/limit로 나눠 본다. 표·누름틀 목록도 그 문단 범위 안 것만 나온다(전체 개수와 다른 범위의 표 번호는 함께 알려 줌).',
    inputSchema: {
      path: pathKo,
      offset: z.number(OFFSET_MSG).int(OFFSET_MSG).min(0, OFFSET_MSG).default(0).describe('건너뛸 문단 수 (0부터)'),
      limit: z.number(LIMIT_MSG).int(LIMIT_MSG).min(1, LIMIT_MSG).max(1000, LIMIT_MSG).default(150)
        .describe('보여 줄 문단 수 (표·누름틀 목록도 이 문단 범위 안 것만)'),
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, offset = 0, limit = 150 }) => {
    const full = mustExist(path);
    if (!isHwp(extOf(full))) throw new Error('hwp_inspect는 한글 문서(.hwp/.hwpx/.hwt)만 다룹니다');
    const h = Hwp.open(full);
    try { return text(inspectText(h.inspect({ offset, limit }), { name: basename(full) })); } finally { h.close(); }
  }));

  server.registerTool('hwp_edit', {
    title: '한글 문서 구조 편집',
    description: '한글(hwp/hwpx) 문서를 프로그램 없이(rhwp) 편집한다. operations를 순서대로 적용하고 마지막에 한 번만 저장한다. '
      + '하나라도 실패하면 아무것도 저장하지 않는다. 번호(문단·표·행·열)는 모두 1부터이며 hwp_inspect 결과 번호를 쓴다. '
      + '앞 작업이 문단을 넣거나 지우면 뒤 작업은 바뀐 번호를 쓴다(결과 줄에 새 번호를 알려 줌).\n'
      + '문서 끝에 붙이기:\n'
      + '  {op:"add_paragraph", text, style?:{bold, italic, underline, size(pt), color("#rrggbb"), font, align("left"|"center"|"right"|"justify")}}\n'
      + '  {op:"add_heading", text, level(1~3)}\n'
      + '  {op:"add_table", rows:[["머리1","머리2"],["값1","값2"]]}  (첫 행 굵게)\n'
      + '  {op:"add_page_break"}\n'
      + '  {op:"add_field", name, label?}  (누름틀, "label: [칸]")\n'
      + '바꾸기·채우기:\n'
      + '  {op:"replace", find, replace}  (영문 대소문자 구분)\n'
      + '  {op:"fill", values:{이름: 값}}  (누름틀 — 같은 이름은 모두 채움, 없으면 {{이름}} 자리표시자)\n'
      + '문단 (at: "end"(기본) | "start" | {after_paragraph:n} | {before_paragraph:n} | {after_text:"찾을 글"}):\n'
      + '  {op:"insert_paragraph", text, at?, style?("바탕글"·"본문"·"개요 1" 등 문서 스타일, 기본 바탕글), like_paragraph?(n: 그 문단 모양 따르기), format?{글자 서식}, align?}  (줄바꿈마다 문단 하나)\n'
      + '  {op:"set_paragraph_text", paragraph | find, text}  (모양 유지. 표·그림·누름틀이 든 문단은 안 됨 → replace·fill·table_set_cell)\n'
      + '  {op:"delete_paragraph", paragraph | find}  (표가 든 문단이면 표도 지움)\n'
      + '  {op:"format_text", find, all?, bold, italic, underline, strike, size(pt), color("#rrggbb"), font, highlight("#rrggbb"), superscript, subscript}  (본문·표 칸, 대소문자 구분, 기본 첫 번째만)\n'
      + '  {op:"format_paragraph", paragraph | paragraphs:[처음,끝] | find, align("left"|"center"|"right"|"justify"|"distribute"), line_spacing(1.6 또는 "160%"), left_indent_mm, right_indent_mm, first_line_indent_mm(음수=내어쓰기), space_before(pt), space_after(pt), keep_with_next, page_break_before}\n'
      + '  {op:"apply_style", paragraph | paragraphs:[처음,끝] | find, style}  (문단·글자 모양이 스타일 것으로 바뀜)\n'
      + '표 (table: 본문 표 번호):\n'
      + '  {op:"table_set_cell", table, row, col, text}  (병합 칸은 덮는 자리 아무거나)\n'
      + '  {op:"table_add_row", table, after_row?(0=맨 위, 생략=맨 아래), values?:[왼쪽부터]}  (세로로 병합된 칸 사이에 넣으면 그 열은 병합에 덮여 값을 넣을 수 없음 → 그 자리는 "")\n'
      + '  {op:"table_delete_row", table, row}\n'
      + '머리말·각주·그림·쪽:\n'
      + '  {op:"set_header"|"set_footer", text, page_numbers?, align?}  (text 안 {page}=쪽 번호, {pages}=전체 쪽 수. page_numbers:true면 끝에 쪽 번호, 글이 없으면 "- 1 -". 빈 text = 지우기. 모든 구역에 적용)\n'
      + '  {op:"add_footnote", find, text}  (본문 글 첫 번째 뒤에 각주)\n'
      + '  {op:"insert_image", image_path(png·jpg·gif·bmp), at?, width_mm?, height_mm?, align?(기본 center), description?}  (그림 하나만 든 새 문단, 크기를 안 주면 원래 크기. 본문보다 넓거나 높으면 본문 안에 맞춰 줄임)\n'
      + '  {op:"page_setup", orientation?("portrait"|"landscape"), paper?("A4"|"A3"|"A5"|"B4"|"B5"|"Letter"|"Legal"), margins_mm?{top,bottom,left,right,header,footer,gutter}, section?}  (한글 방식: 본문은 위 여백+머리말 아래에서 시작)\n'
      + '작업 객체에 없는 인수 이름은 거절한다(오타 방지). output을 생략하면 원본을 덮어쓴다(.hwp 또는 .hwpx로 저장).',
    inputSchema: {
      path: pathKo,
      operations: z.array(hwpOp, E).min(1, msg('operations에 작업이 하나 이상 있어야 합니다')).describe('순서대로 적용할 작업 목록'),
      output: outputKo,
      overwrite: overwriteKo,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, operations, output, overwrite = false }) => {
    const full = mustExist(path);
    if (!isHwp(extOf(full))) throw new Error('hwp_edit는 한글 문서(.hwp/.hwpx/.hwt)만 다룹니다');
    if (output && !HWP_EXTS.has(extOf(abs(output)))) throw new Error('저장 경로는 .hwp 또는 .hwpx 여야 합니다');
    // 저장 경로 확인은 먼저 하되(outPath가 폴더도 만듦), 저장하지 못하면 새로 만든 빈 폴더를 지운다
    const created = missingDirs(output ? abs(output) : full);
    const out = outPath(output, { source: full, allow: HWP_EXTS, overwrite });
    const done = [];
    let saved = false;
    try {
      const h = Hwp.open(full);
      try {
        for (const [i, op] of operations.entries()) {
          try { done.push(`${i + 1}. ${applyHwpOp(h, op, { resolvePath: mustExist })}`); } catch (e) {
            throw new Error(`${i + 1}번째 작업(${op.op}) 실패: ${koreanError(e)} — 저장하지 않았습니다`);
          }
        }
        h.save(out);
        saved = true;
        done.push(`저장: ${out} (${h.pageCount}쪽)`);
      } finally { h.close(); }
    } finally {
      if (!saved) for (const d of created) { try { rmdirSync(d); } catch {} }
    }
    return text(done.join('\n'));
  }));
}
