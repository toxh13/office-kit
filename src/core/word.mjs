// word.mjs — Word 전용 살펴보기(inspect)·편집(edit). 설치된 Word를 COM으로 조작한다 (scripts/com/word-ops.ps1).
// 작업은 순서대로 모두 적용한 뒤 한 번만 저장하고, 하나라도 실패하면 저장하지 않는다. (docs/FEATURES.md 'Word')
import { accessSync, closeSync, constants, existsSync, openSync, readSync, rmSync, statSync } from 'node:fs';
import { extname, resolve, basename, dirname, join } from 'node:path';
import { comCall, probe } from './com.mjs';

export const WORD_EXTS = new Set(['docx', 'doc', 'rtf', 'odt']);
export const IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp']);
/** 지역 이름 대신 쓰는 스타일 키 (WdBuiltinStyle로 바꿔서 한국어 Word에서도 동작) */
export const STYLE_KEYS = ['normal', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5',
  'heading6', 'heading7', 'heading8', 'heading9', 'quote', 'list_bullet', 'list_number', 'caption', 'no_spacing', 'toc_heading'];
export const TABLE_STYLE_KEYS = ['grid', 'plain', 'light_grid', 'light_shading', 'light_list', 'medium_shading1'];
export const HIGHLIGHTS = ['yellow', 'green', 'bright_green', 'cyan', 'turquoise', 'pink', 'blue', 'red', 'dark_blue', 'teal',
  'dark_green', 'violet', 'dark_red', 'dark_yellow', 'gray', 'dark_gray', 'black', 'white', 'none'];
export const ALIGNS = ['left', 'center', 'right', 'justify', 'distribute'];
export const PAPERS = ['A4', 'A3', 'A5', 'B4', 'B5', 'Letter', 'Legal'];
export const WORD_OPS = ['insert_paragraph', 'set_paragraph_text', 'delete_paragraph', 'apply_style', 'format_text',
  'format_paragraph', 'set_list', 'insert_table', 'table_set_cell', 'table_add_row', 'table_delete_row', 'insert_image',
  'insert_page_break', 'set_header', 'set_footer', 'page_setup', 'add_comment', 'delete_comments', 'track_changes',
  'accept_revisions', 'reject_revisions', 'insert_toc', 'update_fields', 'add_footnote', 'insert_hyperlink'];

export const NO_WORD = 'Word가 설치돼 있어야 합니다 — 대신 office_replace / office_fill / office_create는 쓸 수 있습니다';

const extOf = (p) => extname(p).slice(1).toLowerCase();
const samePath = (a, b) => (process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b);

async function requireWord() {
  if (process.platform !== 'win32') throw new Error(NO_WORD);
  const p = await probe();
  if (!p.word) throw new Error(p.error ? `${NO_WORD} (설치된 프로그램 확인 실패: ${p.error})` : NO_WORD);
}

function wordFile(path) {
  const full = resolve(String(path ?? ''));
  if (!path || !existsSync(full)) throw new Error(`파일이 없습니다: ${full}`);
  if (statSync(full).isDirectory()) throw new Error(`파일이 아니라 폴더입니다: ${full}`);
  if (!WORD_EXTS.has(extOf(full))) {
    throw new Error(`Word 문서(.docx/.doc/.rtf/.odt)만 다룹니다: ${basename(full)} — 한글은 hwp_edit, PowerPoint는 ppt_edit, Excel은 excel_edit`);
  }
  return full;
}

// ── 입력 확인·정리 ───────────────────────────────────────

const COLOR_NAMES = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#00b050', blue: '#0070c0', yellow: '#ffff00',
  orange: '#ffc000', purple: '#7030a0', gray: '#808080', grey: '#808080', navy: '#002060' };
const FONT_KEYS = ['bold', 'italic', 'underline', 'strike', 'color', 'highlight', 'size', 'font', 'superscript', 'subscript'];
const PARA_KEYS = ['align', 'line_spacing', 'line_spacing_pt', 'space_before', 'space_after', 'first_line_indent_mm',
  'left_indent_mm', 'right_indent_mm', 'keep_with_next', 'page_break_before'];

class OpError extends Error {}
const fail = (msg) => { throw new OpError(msg); };
const isInt = (v, min = 1) => Number.isInteger(v) && v >= min;
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
// 줄바꿈은 문단 안 줄바꿈(Shift+Enter)으로: 문단 번호가 예상대로 유지된다
const wordText = (v) => String(v ?? '').replace(/\r\n|\r|\n/g, '\v');

function color(v, name) {
  const s = String(v).trim().toLowerCase();
  if (COLOR_NAMES[s]) return COLOR_NAMES[s];
  let m = s.match(/^#?([0-9a-f]{6})$/);
  if (m) return `#${m[1]}`;
  m = s.match(/^#?([0-9a-f])([0-9a-f])([0-9a-f])$/);
  if (m) return `#${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}`;
  return fail(`${name} 색은 '#rrggbb' 형식이어야 합니다: ${v}`);
}

function checkFind(v, name = 'find') {
  if (typeof v !== 'string' || !v) fail(`${name}(찾을 글자)가 필요합니다`);
  if (v.length > 255) fail(`${name}는 255자까지만 됩니다`);
  if (/[\r\n]/.test(v)) fail(`${name}에는 줄바꿈을 넣을 수 없습니다 (한 문단 안의 글자만 찾습니다)`);
  return v;
}

function checkAt(at) {
  if (at === undefined || at === null) return undefined;
  if (at === 'end' || at === 'start') return at;
  if (typeof at === 'object' && !Array.isArray(at)) {
    if (at.after_paragraph !== undefined) { if (!isInt(at.after_paragraph)) fail('at.after_paragraph는 1 이상의 정수여야 합니다'); return { after_paragraph: at.after_paragraph }; }
    if (at.before_paragraph !== undefined) { if (!isInt(at.before_paragraph)) fail('at.before_paragraph는 1 이상의 정수여야 합니다'); return { before_paragraph: at.before_paragraph }; }
    if (at.after_text !== undefined) return { after_text: checkFind(at.after_text, 'at.after_text') };
  }
  return fail(`at은 'end' | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'글'} 중 하나여야 합니다`);
}

// paragraph | paragraphs | find 중 하나 (allowed: 허용하는 선택자)
function checkTarget(op, out, allowed) {
  const given = allowed.filter((k) => op[k] !== undefined && op[k] !== null);
  if (given.length === 0) fail(`${allowed.join(' 또는 ')} 중 하나가 필요합니다`);
  if (given.length > 1) fail(`${given.join('와 ')}는 함께 쓸 수 없습니다 (하나만)`);
  const k = given[0];
  if (k === 'paragraph') { if (!isInt(op.paragraph)) fail('paragraph는 1 이상의 정수여야 합니다'); out.paragraph = op.paragraph; }
  if (k === 'paragraphs') {
    const p = op.paragraphs;
    if (!Array.isArray(p) || p.length !== 2 || !isInt(p[0]) || !isInt(p[1])) fail('paragraphs는 [시작, 끝] 문단 번호여야 합니다 (예: [3, 7])');
    if (p[0] > p[1]) fail(`paragraphs 시작(${p[0]})이 끝(${p[1]})보다 큽니다`);
    out.paragraphs = [p[0], p[1]];
  }
  if (k === 'find') { out.find = checkFind(op.find); if (op.all !== undefined) out.all = Boolean(op.all); }
}

function checkFont(src, out) {
  let n = 0;
  for (const k of ['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript']) {
    if (src[k] === undefined || src[k] === null) continue;
    if (typeof src[k] !== 'boolean') fail(`${k}는 true/false여야 합니다`);
    out[k] = src[k]; n++;
  }
  if (src.color != null) { out.color = color(src.color, 'color'); n++; }
  if (src.highlight != null) {
    const h = src.highlight;
    if (typeof h === 'boolean') out.highlight = h;
    else if (HIGHLIGHTS.includes(String(h).toLowerCase())) out.highlight = String(h).toLowerCase();
    else fail(`highlight는 ${HIGHLIGHTS.join('|')} 또는 true/false여야 합니다: ${h}`);
    n++;
  }
  if (src.size != null) { if (!isNum(src.size) || src.size < 1 || src.size > 1638) fail(`size(pt)는 1~1638이어야 합니다: ${src.size}`); out.size = src.size; n++; }
  if (src.font != null) { if (typeof src.font !== 'string' || !src.font.trim()) fail('font는 글꼴 이름이어야 합니다'); out.font = src.font.trim(); n++; }
  return n;
}

function checkAlign(v) {
  const a = String(v).toLowerCase();
  if (!ALIGNS.includes(a)) fail(`align은 ${ALIGNS.join('|')} 중 하나여야 합니다: ${v}`);
  return a;
}

function checkPara(src, out) {
  let n = 0;
  if (src.align != null) { out.align = checkAlign(src.align); n++; }
  if (src.line_spacing != null) { if (!isNum(src.line_spacing) || src.line_spacing < 0.25 || src.line_spacing > 10) fail(`line_spacing(줄 배수)은 0.25~10이어야 합니다: ${src.line_spacing}`); out.line_spacing = src.line_spacing; n++; }
  if (src.line_spacing_pt != null) { if (!isNum(src.line_spacing_pt) || src.line_spacing_pt < 1 || src.line_spacing_pt > 1584) fail(`line_spacing_pt는 1~1584pt여야 합니다: ${src.line_spacing_pt}`); out.line_spacing_pt = src.line_spacing_pt; n++; }
  for (const k of ['space_before', 'space_after']) {
    if (src[k] == null) continue;
    if (!isNum(src[k]) || src[k] < 0 || src[k] > 1584) fail(`${k}(pt)는 0~1584여야 합니다: ${src[k]}`);
    out[k] = src[k]; n++;
  }
  for (const k of ['first_line_indent_mm', 'left_indent_mm', 'right_indent_mm']) {
    if (src[k] == null) continue;
    if (!isNum(src[k]) || Math.abs(src[k]) > 500) fail(`${k}는 -500~500mm여야 합니다: ${src[k]}`);
    out[k] = src[k]; n++;
  }
  for (const k of ['keep_with_next', 'page_break_before']) {
    if (src[k] == null) continue;
    if (typeof src[k] !== 'boolean') fail(`${k}는 true/false여야 합니다`);
    out[k] = src[k]; n++;
  }
  return n;
}

// 파일 머리(시그니처)로 그림 종류 확인 — 이름만 .png인 글 파일은 Word가 깨진 그림 상자로 넣는다
function imageKind(file) {
  const b = Buffer.alloc(8);
  let n = 0;
  const fd = openSync(file, 'r');
  try { n = readSync(fd, b, 0, 8, 0); } finally { closeSync(fd); }
  if (n >= 4 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (n >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpg';
  if (n >= 4 && b.toString('latin1', 0, 4) === 'GIF8') return 'gif';
  if (n >= 2 && b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
  return null;
}

function checkText(v, name = 'text') {
  if (typeof v !== 'string' && typeof v !== 'number') fail(`${name}(글자)가 필요합니다`);
  return wordText(v);
}
const checkStyle = (v) => { if (typeof v !== 'string' || !v.trim()) fail('style(스타일 키 또는 이름)이 필요합니다'); return v.trim(); };
const checkTable = (op, out) => { if (!isInt(op.table)) fail('table(표 번호, 1부터)이 필요합니다'); out.table = op.table; };

function normalizeOp(op) {
  if (!op || typeof op !== 'object' || Array.isArray(op)) fail('작업은 {op: ...} 객체여야 합니다');
  if (!WORD_OPS.includes(op.op)) fail(`알 수 없는 작업입니다: ${op.op ?? '(op 없음)'} (가능: ${WORD_OPS.join(', ')})`);
  const out = { op: op.op };
  switch (op.op) {
    case 'insert_paragraph': {
      out.text = checkText(op.text);
      out.at = checkAt(op.at);
      if (op.style != null) out.style = checkStyle(op.style);
      if (op.format != null) {
        if (typeof op.format !== 'object' || Array.isArray(op.format)) fail('format은 {bold, color, ...} 객체여야 합니다');
        out.format = {};
        checkFont(op.format, out.format);
      }
      break;
    }
    case 'set_paragraph_text':
      if (!isInt(op.paragraph)) fail('paragraph(문단 번호, 1부터)가 필요합니다');
      out.paragraph = op.paragraph; out.text = checkText(op.text);
      break;
    case 'delete_paragraph': checkTarget(op, out, ['paragraph', 'find']); break;
    case 'apply_style': checkTarget(op, out, ['paragraph', 'paragraphs', 'find']); out.style = checkStyle(op.style); break;
    case 'format_text':
      checkTarget(op, out, ['find', 'paragraph', 'paragraphs']);
      if (!checkFont(op, out)) fail(`바꿀 서식이 없습니다 (${FONT_KEYS.join(', ')} 중 하나 이상)`);
      break;
    case 'format_paragraph':
      checkTarget(op, out, ['paragraph', 'paragraphs', 'find']);
      if (!checkPara(op, out)) fail(`바꿀 문단 서식이 없습니다 (${PARA_KEYS.join(', ')} 중 하나 이상)`);
      break;
    case 'set_list':
      checkTarget(op, out, ['paragraph', 'paragraphs']);
      if (!['bullet', 'number', 'none'].includes(op.type)) fail("type은 'bullet' | 'number' | 'none' 이어야 합니다");
      out.type = op.type;
      break;
    case 'insert_table': {
      const rows = op.rows;
      if (!Array.isArray(rows) || !rows.length || !rows.every((r) => Array.isArray(r))) fail('rows는 [["머리1","머리2"],["값1","값2"]] 형태의 2차원 배열이어야 합니다');
      const cols = Math.max(...rows.map((r) => r.length));
      if (cols < 1) fail('rows에 칸이 없습니다');
      if (rows.length > 500 || cols > 63 || rows.length * cols > 5000) fail(`표가 너무 큽니다: ${rows.length}행×${cols}열 (최대 500행, 63열, 5000칸)`);
      out.rows = rows.map((r) => r.map((v) => {
        if (v === null || v === undefined) return '';
        if (typeof v !== 'string' && typeof v !== 'number' && typeof v !== 'boolean') fail('표 칸 값은 글자나 숫자여야 합니다');
        return wordText(v);
      }));
      out.at = checkAt(op.at);
      if (op.style != null) out.style = checkStyle(op.style);
      if (op.header_bold != null) { if (typeof op.header_bold !== 'boolean') fail('header_bold는 true/false여야 합니다'); out.header_bold = op.header_bold; }
      if (op.autofit != null) {
        const a = op.autofit === true ? 'content' : op.autofit === false ? 'fixed' : op.autofit;
        if (!['window', 'content', 'fixed'].includes(a)) fail("autofit은 'window'(쪽 너비) | 'content'(내용) | 'fixed' 이어야 합니다");
        out.autofit = a;
      }
      break;
    }
    case 'table_set_cell':
      checkTable(op, out);
      if (!isInt(op.row) || !isInt(op.col)) fail('row와 col(1부터)이 필요합니다');
      out.row = op.row; out.col = op.col; out.text = checkText(op.text);
      break;
    case 'table_add_row':
      checkTable(op, out);
      if (op.after_row != null) { if (!isInt(op.after_row, 0)) fail('after_row는 0 이상의 정수여야 합니다 (0 = 맨 위)'); out.after_row = op.after_row; }
      if (op.values != null) {
        if (!Array.isArray(op.values)) fail('values는 칸 값 배열이어야 합니다');
        out.values = op.values.map((v) => (v === null || v === undefined ? null : wordText(v)));
      }
      break;
    case 'table_delete_row':
      checkTable(op, out);
      if (!isInt(op.row)) fail('row(행 번호, 1부터)가 필요합니다');
      out.row = op.row;
      break;
    case 'insert_image': {
      if (typeof op.image_path !== 'string' || !op.image_path.trim()) fail('image_path(그림 파일 경로)가 필요합니다');
      const img = resolve(op.image_path.trim().replace(/^"(.*)"$/, '$1'));
      if (!IMAGE_EXTS.has(extOf(img))) fail(`그림은 ${[...IMAGE_EXTS].join(', ')} 파일만 넣을 수 있습니다: ${basename(img)}`);
      if (!existsSync(img) || statSync(img).isDirectory()) fail(`그림 파일이 없습니다: ${img}`);
      if (!imageKind(img)) fail(`그림 파일이 아니거나 깨졌습니다: ${basename(img)} (png·jpg·gif·bmp 파일 머리가 아님)`);
      out.image_path = img;
      out.at = checkAt(op.at);
      for (const k of ['width_mm', 'height_mm']) {
        if (op[k] == null) continue;
        if (!isNum(op[k]) || op[k] <= 0 || op[k] > 1000) fail(`${k}는 0보다 크고 1000mm 이하여야 합니다: ${op[k]}`);
        out[k] = op[k];
      }
      if (op.caption != null) { out.caption = checkText(op.caption, 'caption'); out.caption_label = typeof op.caption_label === 'string' && op.caption_label.trim() ? op.caption_label.trim() : '그림'; }
      if (op.align != null) out.align = checkAlign(op.align);
      break;
    }
    case 'insert_page_break': out.at = checkAt(op.at); break;
    case 'set_header':
    case 'set_footer':
      if (op.text != null) out.text = checkText(op.text);
      if (op.page_numbers != null) { if (typeof op.page_numbers !== 'boolean') fail('page_numbers는 true/false여야 합니다'); out.page_numbers = op.page_numbers; }
      if (out.text === undefined && !out.page_numbers) fail('text나 page_numbers:true가 필요합니다');
      if (op.align != null) out.align = checkAlign(op.align);
      if (op.section != null) { if (!isInt(op.section)) fail('section은 1 이상의 정수여야 합니다'); out.section = op.section; }
      break;
    case 'page_setup': {
      let n = 0;
      if (op.orientation != null) { if (!['portrait', 'landscape'].includes(op.orientation)) fail("orientation은 'portrait'(세로) | 'landscape'(가로) 이어야 합니다"); out.orientation = op.orientation; n++; }
      if (op.paper != null) {
        const p = PAPERS.find((x) => x.toLowerCase() === String(op.paper).toLowerCase());
        if (!p) fail(`paper는 ${PAPERS.join('|')} 중 하나여야 합니다: ${op.paper}`);
        out.paper = p; n++;
      }
      if (op.margins_mm != null) {
        const m = op.margins_mm;
        if (typeof m !== 'object' || Array.isArray(m)) fail('margins_mm는 {top, bottom, left, right} 객체여야 합니다');
        out.margins_mm = {};
        for (const k of ['top', 'bottom', 'left', 'right']) {
          if (m[k] == null) continue;
          if (!isNum(m[k]) || m[k] < 0 || m[k] > 500) fail(`margins_mm.${k}는 0~500mm여야 합니다: ${m[k]}`);
          out.margins_mm[k] = m[k]; n++;
        }
      }
      if (op.section != null) { if (!isInt(op.section)) fail('section은 1 이상의 정수여야 합니다'); out.section = op.section; }
      if (!n) fail('orientation, paper, margins_mm 중 하나 이상이 필요합니다');
      break;
    }
    case 'add_comment':
      out.find = checkFind(op.find); out.text = checkText(op.text);
      if (op.all !== undefined) out.all = Boolean(op.all);
      if (op.author != null) { if (typeof op.author !== 'string' || !op.author.trim()) fail('author는 이름 글자여야 합니다'); out.author = op.author.trim(); }
      break;
    case 'delete_comments':
      if (op.index != null) { if (!isInt(op.index)) fail('index(메모 번호, 1부터)는 1 이상의 정수여야 합니다'); out.index = op.index; }
      else if (op.all === true) out.all = true;
      else fail('all:true(모두) 또는 index(메모 번호)가 필요합니다');
      break;
    case 'track_changes':
      if (typeof op.on !== 'boolean') fail('on(true/false)이 필요합니다');
      out.on = op.on;
      break;
    case 'accept_revisions': case 'reject_revisions': case 'update_fields': break;
    case 'insert_toc':
      out.at = checkAt(op.at);
      if (op.levels != null) { if (!isInt(op.levels) || op.levels > 9) fail('levels는 1~9여야 합니다'); out.levels = op.levels; }
      if (op.title != null) out.title = checkText(op.title, 'title');
      break;
    case 'add_footnote':
      out.find = checkFind(op.find); out.text = checkText(op.text);
      if (op.all !== undefined) out.all = Boolean(op.all);
      break;
    case 'insert_hyperlink': {
      out.find = checkFind(op.find);
      const url = typeof op.url === 'string' ? op.url.trim() : '';
      if (!/^(https?:\/\/|mailto:|ftp:\/\/)\S+$/i.test(url)) fail('url은 http://, https://, mailto:, ftp:// 로 시작해야 합니다');
      out.url = url;
      if (op.all !== undefined) out.all = Boolean(op.all);
      break;
    }
    default: break;
  }
  return out;
}

/** 작업 목록 확인·정리. 틀리면 "n번째 작업(op): 이유" 한국어 오류 */
export function normalizeWordOps(operations) {
  if (!Array.isArray(operations) || !operations.length) throw new Error('operations(작업 목록)가 비어 있습니다');
  if (operations.length > 500) throw new Error(`작업이 너무 많습니다: ${operations.length}개 (한 번에 500개까지)`);
  return operations.map((op, i) => {
    try { return normalizeOp(op); } catch (e) {
      if (e instanceof OpError) throw new Error(`${i + 1}번째 작업(${op?.op ?? '?'}) 입력 오류: ${e.message}`);
      throw e;
    }
  });
}

// ── 결과·오류를 한국어로 ─────────────────────────────────

const ERRORS = [
  [/^paragraph out of range: (-?\d+) of (\d+)$/, (m) => `문단 번호 ${m[1]}이(가) 범위를 벗어났습니다 (전체 ${m[2]}개, 1부터 — 앞 작업이 문단 수를 바꿨을 수 있음)`],
  [/^table out of range: (-?\d+) of (\d+)$/, (m) => `표 번호 ${m[1]}이(가) 범위를 벗어났습니다 (표 ${m[2]}개)`],
  [/^row out of range: (-?\d+) of (\d+)$/, (m) => `행 번호 ${m[1]}이(가) 범위를 벗어났습니다 (전체 ${m[2]}행)`],
  [/^cell out of range: (-?\d+),(-?\d+) of (\d+)x(\d+)$/, (m) => `${m[1]}행 ${m[2]}열 칸이 없습니다 (표 ${m[3]}행×${m[4]}열, 병합된 칸이면 열 수가 다를 수 있음)`],
  [/^too many values: (\d+) of (\d+)$/, (m) => `값이 ${m[1]}개인데 표는 ${m[2]}열입니다`],
  [/^comment out of range: (-?\d+) of (\d+)$/, (m) => `메모 번호 ${m[1]}이(가) 범위를 벗어났습니다 (메모 ${m[2]}개)`],
  [/^section out of range: (-?\d+) of (\d+)$/, (m) => `구역 번호 ${m[1]}이(가) 범위를 벗어났습니다 (구역 ${m[2]}개)`],
  [/^not found: ([\s\S]*)$/, (m) => `찾는 글자가 없습니다: '${m[1]}' (본문에서 대소문자 구분)`],
  [/^style not found: ([\s\S]*)$/, (m) => `문서에 없는 스타일입니다: '${m[1]}' (키 ${STYLE_KEYS.join(', ')} 또는 word_inspect의 스타일 이름을 쓰세요)`],
  [/^row end mark: (\d+)$/, (m) => `${m[1]}번 문단은 표의 행 끝 표시라 고칠 수 없습니다 (행은 table_add_row / table_delete_row로)`],
  [/^merged cell: (\d+)$/, (m) => `${m[1]}번 문단은 세로로 병합된 칸의 아랫부분(위 칸에 합쳐진 자리)이라 고칠 수 없습니다 — 병합된 칸의 첫 행 문단(word_inspect의 '표n r행c열')을 쓰세요`],
  [/^merged rows: (\d+)$/, (m) => `${m[1]}행은 모든 칸이 위·아래 행과 세로로 병합돼 있어 그 행만 따로 지우거나 그 자리에 행을 넣을 수 없습니다 (Word에서 병합을 먼저 푸세요)`],
  [/^row cells: (\d+) of (\d+)$/, (m) => `값이 ${m[1]}개인데 새 행은 칸이 ${m[2]}개입니다 (가로로 병합된 칸이 있는 행을 본뜸) — 값 개수를 칸 수에 맞추세요`],
  [/^merged cell value: (\d+),(\d+)$/, (m) => `새 ${m[1]}행 ${m[2]}열은 위 칸과 세로로 병합돼 값을 넣을 수 없습니다 — values의 그 자리는 null로 두세요`],
  [/^inserted at: (\d+)$/, (m) => `새 행이 예상과 다른 자리(${m[1]}행)에 들어갔습니다 — 병합된 칸을 word_inspect로 확인하세요`],
  [/^too many: (\d+)$/, (m) => `찾은 곳이 ${m[1]}곳이 넘습니다 — find를 더 좁히거나 paragraphs 범위로 나눠 주세요`],
  [/^only in toc: ([\s\S]*)$/, (m) => `찾는 글자가 목차 안에만 있습니다: '${m[1]}' (목차는 update_fields 때 다시 만들어지므로 본문 제목을 고치세요)`],
  [/^no headings: (\d+)$/, (m) => `목차에 넣을 제목 문단(heading1~heading${m[1]} 스타일)이 없습니다 — 먼저 apply_style로 제목을 정하세요 (같은 작업 목록이면 insert_toc보다 앞에)`],
  [/^image not found: ([\s\S]*)$/, (m) => `그림 파일이 없습니다: ${m[1]}`],
  [/^find too long$/, () => '찾을 글자는 255자까지만 됩니다'],
  [/^bad value: no format given$/, () => '바꿀 서식이 없습니다'],
  [/^bad value: ([\s\S]*)$/, (m) => `값이 올바르지 않습니다: ${m[1]}`],
  [/^unknown op: ([\s\S]*)$/, (m) => `알 수 없는 작업입니다: ${m[1]}`],
  [/range cannot be deleted|cannot be deleted/i, () => '이 부분은 지울 수 없습니다 (목차·필드 결과나 보호된 영역일 수 있음)'],
  [/requested member of the collection does not exist/i, () => '대상이 없습니다 (번호를 word_inspect로 확인하세요)'],
  [/document is (?:locked|protected)|protected document|문서가 보호/i, () => '보호된 문서라 고칠 수 없습니다'],
  // Word(COM)가 직접 낸 영어 오류
  [/cannot insert text into a table cell which has been merged/i, () => '세로로 병합된 칸의 아랫부분에는 글을 넣을 수 없습니다 — 병합된 칸의 첫 행 문단을 쓰세요'],
  [/could not apply the style/i, () => '이 자리에는 그 스타일을 적용할 수 없습니다 (표 스타일·글자 스타일인지, 문단 스타일 이름이 맞는지 확인하세요)'],
  [/too large for the page/i, () => '여백·들여쓰기가 쪽 너비보다 큽니다 — 값을 줄이세요'],
  [/\bvalue out of range\b/i, () => '값이 Word가 받는 범위를 벗어났습니다'],
];

// 저장 단계 오류 (작업 번호 없음): com.mjs가 'Word 작업 실패: …'로 감싸 보낸다
const READ_ONLY = '문서가 읽기 전용이거나 다른 프로그램(Word 등)에서 열려 있어 그 자리에 저장할 수 없습니다 — 문서를 닫거나 output으로 다른 경로를 주세요 (저장하지 않았습니다)';
const SAVE_ERRORS = [
  [/(?:^|: )read only$/, () => READ_ONLY],
  [/in use by another process|being used by another|used by another process/i, () => '저장할 파일이 다른 프로그램에서 열려 있어 저장하지 못했습니다 — 그 파일을 닫거나 다른 output 경로를 주세요 (저장하지 않았습니다)'],
  [/can.t save this file because it.s read-only|file is read-only|read-only file/i, () => '저장할 파일이 읽기 전용이거나 다른 프로그램에서 열려 있어 저장하지 못했습니다 — 다른 output 경로를 주세요 (저장하지 않았습니다)'],
];

function koreanOpError(detail) {
  const s = String(detail).trim();
  for (const [re, fmt] of ERRORS) { const m = s.match(re); if (m) return fmt(m); }
  return s;
}

const short = (s, n = 40) => {
  const t = String(s ?? '').replace(/\v/g, '↵').replace(/\f/g, '[쪽 나누기]');
  return t.length > n ? `${t.slice(0, n)}…` : t;
};
const quote = (s, n) => `'${short(s, n)}'`;
const span = (a, b) => (a === b ? `${a}번 문단` : `${a}~${b}번 문단`);
const targetOf = (op) => (op.paragraphs ? span(op.paragraphs[0], op.paragraphs[1]) : `${op.paragraph}번 문단`);
const PROP_NAMES = { bold: '굵게', italic: '기울임', underline: '밑줄', strike: '취소선', color: '색', highlight: '형광펜', size: '크기(pt)',
  font: '글꼴', superscript: '위 첨자', subscript: '아래 첨자', align: '정렬', line_spacing: '줄 간격(배)', line_spacing_pt: '줄 간격(pt)',
  space_before: '문단 앞(pt)', space_after: '문단 뒤(pt)', first_line_indent_mm: '첫 줄 들여쓰기(mm)', left_indent_mm: '왼쪽 들여쓰기(mm)',
  right_indent_mm: '오른쪽 들여쓰기(mm)', keep_with_next: '다음 문단과 함께', page_break_before: '앞에서 쪽 나누기' };
const props = (op, keys) => keys.filter((k) => op[k] !== undefined)
  .map((k) => `${PROP_NAMES[k]} ${op[k] === true ? '켬' : op[k] === false ? '끔' : op[k]}`).join(', ');
const MM = (v) => `${v}mm`;
// PowerShell JSON은 원소 하나짜리 배열을 값 하나로 줄 수 있다
const arr = (v) => (v === undefined || v === null ? [] : Array.isArray(v) ? v : [v]);
const TRACKED = ' (변경 내용으로 기록 — 수락하기 전까지 번호 그대로)';
// 지울 수 없어 글만 비운 이유 (word-ops.ps1 WdDeletePara)
const CLEARED_WHY = {
  cell: '표 칸의 유일한 문단',
  last: '문서 마지막 문단인데 앞이 표·구역 나누기라 합칠 수 없음',
  section: '구역 나누기가 든 문단인데 앞이 표·구역 나누기라 합칠 수 없음',
  table: '표 바로 앞 문단인데 앞에 합칠 문단이 없음',
  kept: 'Word가 문단 표시를 남김',
};

function describe(op, r) {
  const total = ` · 전체 ${r.paragraphs}문단`;
  switch (op.op) {
    case 'insert_paragraph': return `문단 넣음: ${r.paragraph}번 [${r.style}] ${quote(op.text)}${total}`;
    case 'set_paragraph_text': return `${r.paragraph}번 문단 글 바꿈: ${quote(r.old, 30)} → ${quote(op.text, 30)}`;
    case 'delete_paragraph': {
      const parts = [];
      const deleted = arr(r.deleted);
      const sections = arr(r.sectionKept);
      if (deleted.length) parts.push(`${deleted.join(', ')}번 문단 지움${sections.length ? ` (${sections.join(', ')}번의 구역 나누기는 앞 문단 끝에 남겨 쪽 설정 유지)` : ''}`);
      for (const c of arr(r.cleared)) parts.push(`${c.paragraph}번 문단은 지울 수 없는 자리(${CLEARED_WHY[c.why] ?? c.why})라 글만 비움 (번호 그대로)`);
      const tracked = arr(r.tracked);
      if (tracked.length) parts.push(`${tracked.join(', ')}번 문단을 변경 내용으로 지움 표시 (번호는 그대로 — 수락하면 사라짐)`);
      const trackedText = arr(r.trackedText);
      if (trackedText.length) parts.push(`${trackedText.join(', ')}번 문단은 글만 변경 내용으로 지움 표시 (구역 나누기·표 앞·칸 끝 문단 표시는 남김, 번호 그대로)`);
      return `${parts.join(', ')}${total}`;
    }
    case 'apply_style': return `스타일 '${r.style}'${r.character ? '(글자 스타일)' : ''} 적용: ${op.find ? `${quote(op.find)}${r.character ? '' : ' 포함 문단'} ${r.count}곳` : span(r.from, r.to)}`;
    case 'format_text': {
      return `글자 서식 (${props(op, FONT_KEYS)}): ${op.find ? `${quote(op.find)} ${r.count}곳` : targetOf(op)}`;
    }
    case 'format_paragraph': {
      return `문단 서식 (${props(op, PARA_KEYS)}): 문단 ${r.count}개${op.find ? ` (${quote(op.find)} 포함 문단)` : ''}`;
    }
    case 'set_list': {
      const removed = arr(r.removed);
      const note = removed.length
        ? ` (앞에 직접 쓴 번호·기호는 자동 번호로 바뀌어 글에서 빠짐: ${removed.slice(0, 5).map((x) => `${x.paragraph}번 ${quote(x.removed, 10)}`).join(', ')}${removed.length > 5 ? ` 외 ${removed.length - 5}곳` : ''})`
        : '';
      return `${span(r.from, r.to)} ${{ bullet: '글머리표 목록으로', number: '번호 목록으로', none: '목록 해제' }[r.type]}${note}`;
    }
    case 'insert_table': return `표 ${r.table} 넣음: ${r.rows}행×${r.cols}열, ${r.paragraph}번 문단부터 (표 다음 문단 ${r.after}번)${total}`;
    case 'table_set_cell': return `표 ${r.table} ${r.row}행 ${r.col}열: ${quote(r.old, 30)} → ${quote(op.text, 30)}`;
    case 'table_add_row': return `표 ${r.table}에 행 추가: 새 ${r.row}행 (전체 ${r.rows}행)${r.tracked ? ' (변경 내용으로 기록)' : ''}${total}`;
    case 'table_delete_row': {
      if (r.tableDeleted) return `표 ${r.table}의 마지막 남은 행이라 표 전체를 지움${r.tracked ? TRACKED : ''}${total}`;
      if (r.tracked) return `표 ${r.table} ${r.row}행을 변경 내용으로 지움 표시 (수락하면 사라짐 — 그전까지 행 번호 그대로, 전체 ${r.rows}행)${total}`;
      const kept = arr(r.kept);
      return `표 ${r.table} ${r.row}행 지움 (남은 ${r.rows}행)${kept.length ? ` — 세로로 병합된 칸의 글 ${kept.map((k) => quote(k, 20)).join(', ')}은 남은 행에 둠` : ''}${total}`;
    }
    case 'insert_image': return `그림 넣음: ${r.paragraph}번 문단, ${MM(r.width)}×${MM(r.height)}${r.caption ? `, 캡션 ${r.caption}번 문단` : ''}${total}`;
    case 'insert_page_break': return `쪽 나누기: ${r.paragraph}번 문단 (다음 내용은 새 쪽에서 시작)${total}`;
    case 'set_header':
    case 'set_footer': {
      const what = op.op === 'set_header' ? '머리말' : '꼬리말';
      const shown = String(r.text ?? '').replace(/\{pages\}/g, '[전체 쪽 수]').replace(/\{page\}/g, '[쪽 번호]');
      return `${what} 설정 (구역 ${r.sections}/${r.total}개): ${quote(shown, 60)}`;
    }
    case 'page_setup': return `쪽 설정 (구역 ${r.sections}개): ${r.orientation === 1 ? '가로' : '세로'} ${r.width}×${r.height}mm, 여백 위 ${r.top}·아래 ${r.bottom}·왼쪽 ${r.left}·오른쪽 ${r.right}mm`;
    case 'add_comment': return `메모 달기: ${quote(op.find)} ${r.count}곳 (메모 모두 ${r.total}개)`;
    case 'delete_comments': return `메모 ${r.deleted}개 지움 (남은 메모 ${r.total}개)`;
    case 'track_changes': return `변경 내용 추적 ${r.on ? '켬 (이후 작업은 변경 내용으로 기록됨)' : '끔'}`;
    case 'accept_revisions': return `변경 내용 ${r.count}개 모두 수락${r.left ? ` (남은 ${r.left}개)` : ''}${total}`;
    case 'reject_revisions': return `변경 내용 ${r.count}개 모두 거부${r.left ? ` (남은 ${r.left}개)` : ''}${total}`;
    case 'insert_toc': return `목차 넣음: ${r.title ? `제목 ${r.title}번 문단, ` : ''}${r.paragraph}번 문단부터 (목차 ${r.entries}문단, 제목 1~${r.levels}수준)${total}`;
    case 'update_fields': return `필드 새로 고침: 필드 ${r.fields}개, 목차 ${r.tocs}개${r.skipped ? ` (바깥 파일·웹·다른 프로그램을 불러오는 필드 ${r.skipped}개는 안전을 위해 건너뜀)` : ''}`;
    case 'add_footnote': return `각주 달기: ${quote(op.find)} ${r.count}곳 (각주 모두 ${r.total}개)`;
    case 'insert_hyperlink': return `링크 걸기: ${quote(op.find)} ${r.count}곳 → ${op.url}`;
    default: return op.op;
  }
}

// ── 살펴보기 ─────────────────────────────────────────────

/** Word 문서 구조 → 원자료 (scripts/com/word-ops.ps1 Invoke-WordInspect). offset: 건너뛸 문단 수(0부터), limit: 보여 줄 문단 수 */
export async function wordInspect(path, { offset = 0, limit = 200 } = {}) {
  const full = wordFile(path);
  if (!Number.isInteger(offset) || offset < 0) throw new Error(`offset은 0 이상의 정수여야 합니다: ${offset}`);
  if (!Number.isInteger(limit) || limit < 1 || limit > 1000) throw new Error(`limit은 1~1000이어야 합니다: ${limit}`);
  await requireWord();
  const res = await comCall({ app: 'word', op: 'inspect', path: full, offset, limit }, { timeoutMs: 300_000 });
  return { path: full, ...res.inspect };
}

const PAPER_SIZES = [['A4', 210, 297], ['A3', 297, 420], ['A5', 148, 210], ['B4', 257, 364], ['B5', 182, 257], ['Letter', 215.9, 279.4], ['Legal', 215.9, 355.6]];
function paperName(w, h) {
  const [a, b] = [Math.min(w, h), Math.max(w, h)];
  const hit = PAPER_SIZES.find(([, x, y]) => Math.abs(a - x) <= 1 && Math.abs(b - y) <= 1);
  return hit ? hit[0] : '사용자 지정';
}
const FIELD_NAMES = { 33: '쪽 번호', 26: '전체 쪽 수', 31: '날짜', 32: '시각', 13: '목차', 12: '순서(SEQ)' };
const REVISION_NAMES = { 1: '삽입', 2: '삭제', 3: '서식', 8: '스타일', 10: '문단 서식', 14: '이동(원래 자리)', 15: '이동(새 자리)' };
const PRIVATE_USE = new RegExp(`[${String.fromCodePoint(0xe000)}-${String.fromCodePoint(0xf8ff)}]`, 'g');
const LIST_NAMES = { 2: '글머리표', 3: '번호', 4: '번호(개요)', 5: '번호(혼합)', 6: '그림 글머리표', 1: '번호(목록 필드)' };

// 머리말·꼬리말의 필드 코드(PAGE, NUMPAGES)를 {page}·{pages}로 (set_header/set_footer와 같은 표기)
function fieldCodes(t) {
  return String(t ?? '').replace(/\x13\s*([A-Za-z]+)[^\x13\x15]*\x15/g, (_m, code) => {
    const c = code.toUpperCase();
    return c === 'PAGE' ? '{page}' : c === 'NUMPAGES' ? '{pages}' : `{${c}}`;
  });
}

// 문단 글의 특수 문자를 읽기 쉽게
function showText(t, images) {
  let s = String(t ?? '');
  if (images && /^[\/\x01\s]*$/.test(s)) return images > 1 ? `[그림 ${images}개]` : '[그림]';
  s = s.replace(/[\x13\x14\x15\x07]/g, '').replace(/\f/g, '[쪽 나누기]').replace(/\x0b/g, '↵').replace(/\t/g, '⇥')
    .replace(/\x02/g, '[각주]').replace(/\x05/g, '').replace(/\x01/g, '[그림]').replace(/\x0e/g, '[단 나누기]');
  if (images) s += ` [그림 ${images}개]`;
  return s;
}

/** wordInspect 결과 → AI가 읽을 한국어 글 */
export function formatWordInspect(info) {
  const L = [];
  const yes = (v) => (v ? '켜짐' : '꺼짐');
  L.push(`${basename(info.path ?? info.name ?? '')} — Word 문서: ${info.pages}쪽, 단어 ${info.words}개, 글자 ${info.characters}자, 문단 ${info.paragraphTotal}개, 표 ${info.tableCount}개`);
  for (const s of info.sections ?? []) {
    L.push(`구역 ${s.index}: ${paperName(s.width, s.height)} ${s.orientation === 1 ? '가로' : '세로'} ${s.width}×${s.height}mm, 여백 위 ${s.top}·아래 ${s.bottom}·왼쪽 ${s.left}·오른쪽 ${s.right}mm`
      + (s.firstPage ? ', 첫 쪽 머리말·꼬리말 따로' : ''));
    for (const k of ['header', 'footer']) {
      const name = k === 'header' ? '머리말' : '꼬리말';
      const fields = (s[`${k}Fields`] ?? []).map((t) => FIELD_NAMES[t] ?? `필드 ${t}`);
      const linked = s[`${k}Linked`] ? ' (앞 구역과 같음)' : '';
      const txt = showText(fieldCodes(s[k]));
      if (txt || fields.length) L.push(`  ${name}: ${txt ? quote(txt, 80) : '(글 없음)'}${fields.length ? ` [${fields.join(', ')}]` : ''}${linked}`);
      if (s.firstPage && s[`first${k}`]) L.push(`  첫 쪽 ${name}: ${quote(showText(fieldCodes(s[`first${k}`])), 80)}`);
    }
  }
  if (info.sectionCount > (info.sections?.length ?? 0)) L.push(`(구역 ${info.sectionCount}개 중 ${info.sections.length}개만 표시)`);
  L.push(`변경 내용 추적: ${yes(info.trackRevisions)}, 변경 ${info.revisionCount}개`);
  for (const r of info.revisions ?? []) L.push(`  - ${REVISION_NAMES[r.type] ?? '변경'} ${quote(showText(r.text), 40)}${r.author ? ` (${r.author})` : ''}`);
  if (info.revisionCount > (info.revisions?.length ?? 0)) L.push(`  (${info.revisions.length}개만 표시)`);
  const comments = arr(info.comments);
  L.push(`메모 ${info.commentCount}개${info.commentsListed === false && info.commentCount ? ' (목록은 offset=0일 때만 보여 줌)' : ''}`);
  for (const c of comments) L.push(`  ${c.index}. ${c.author || '작성자 없음'}: ${quote(showText(c.text), 80)} ← ${quote(showText(c.scope), 30)}${c.done ? ' (해결됨)' : ''}`);
  if (info.commentsListed !== false && info.commentCount > comments.length) L.push(`  (${comments.length}개만 표시)`);
  const c = { toc: 0, inlineShapes: 0, shapes: 0, contentControls: 0, bookmarks: 0, footnotes: 0, endnotes: 0, hyperlinks: 0, fields: 0, ...info.counts };
  L.push(`개수: 목차 ${c.toc} · 그림(글 안) ${c.inlineShapes} · 떠 있는 개체 ${c.shapes} · 콘텐츠 컨트롤 ${c.contentControls} · 책갈피 ${c.bookmarks}`
    + ` · 각주 ${c.footnotes} · 미주 ${c.endnotes} · 링크 ${c.hyperlinks} · 필드 ${c.fields}`);
  const styles = info.styles ?? [];
  if (styles.length) {
    L.push(`쓰인 스타일 (본문): ${styles.map((s) => `${s.name}${s.key ? `(${s.key})` : ''}${s.character ? '[글자 스타일]' : ''}`).join(', ')}`);
  }
  if (info.tableCount) {
    L.push(`표 ${info.tableCount}개:`);
    for (const t of info.tables ?? []) {
      L.push(`  표 ${t.index}: ${t.rows}행×${t.cols}열${t.uniform ? '' : '(병합된 칸 있음)'}, ${t.paragraph}번 문단부터${t.style ? `, 스타일 ${t.style}` : ''}`
        + `, 첫 행 [${(t.first ?? []).map((x) => showText(x)).join(' | ')}]`);
    }
    if (info.tableCount > (info.tables?.length ?? 0)) L.push(`  (${info.tables.length}개만 표시)`);
  }
  const paras = info.paragraphs ?? [];
  const from = info.offset + 1;
  const to = Math.min(info.offset + info.limit, info.paragraphTotal);
  if (from > info.paragraphTotal) {
    L.push(`문단: offset=${info.offset}이 전체 ${info.paragraphTotal}개를 넘었습니다`);
  } else {
    L.push(`문단 ${from}~${to} / 전체 ${info.paragraphTotal}개 (번호는 편집 작업에 그대로 씀. 표의 행 끝 표시는 번호만 차지하고 목록에서 뺌):`);
    for (const p of paras) {
      if (p.merged) {
        L.push(`  ${p.i} [표${p.table} 병합된 칸 — 위 칸에 세로로 합쳐진 자리라 글을 넣을 수 없음]`);
        continue;
      }
      const tags = [p.key ? `${p.style}·${p.key}` : p.style];
      if (p.table !== undefined) tags.push(`표${p.table} ${p.row}행${p.col}열`);
      // 기호 글꼴의 글머리표(사용자 영역 문자)는 •로
      if (p.list !== undefined) tags.push(`${LIST_NAMES[p.listType] ?? '목록'} ${String(p.list).replace(PRIVATE_USE, '•')}`.trim());
      const shown = showText(p.text, p.images);
      // 구역 나누기는 쪽 나누기와 달리 쪽 설정(방향·여백)을 나눈다
      L.push(`  ${p.i} [${tags.join(' · ')}] ${p.deleted ? '[지움 표시됨 — 변경 내용 수락 전] ' : ''}${shown || (p.sectionBreak ? '' : '(빈 문단)')}`
        + `${p.sectionBreak ? `${shown ? ' ' : ''}[구역 나누기]` : ''}${p.cut ? ` …(전체 ${p.cut}자)` : ''}`);
    }
    if (to < info.paragraphTotal) L.push(`[계속] 다음 문단은 offset=${to} 로 다시 살펴보세요.`);
  }
  return L.join('\n');
}

// ── 편집 ─────────────────────────────────────────────────

// Word가 연 문서 옆에 만드는 소유자 파일 이름: 이름(확장자 앞)이 6자 이하면 '~$'+이름, 7자면 첫 글자, 8자 이상이면 첫 두 글자를 바꾼다
function ownerFiles(file) {
  const name = basename(file);
  const stem = name.length - extname(name).length;
  const cand = stem <= 6 ? [`~$${name}`] : stem === 7 ? [`~$${name.slice(1)}`] : [`~$${name.slice(2)}`];
  return cand.map((n) => join(dirname(file), n));
}
// 이번 호출이 시작된 뒤 생긴 작은 소유자 파일만 지운다 (사용자의 Word가 쓰는 중이면 지워지지 않음)
function removeOwnerFiles(files, since) {
  for (const f of new Set(files)) {
    for (const o of ownerFiles(f)) {
      try {
        const st = statSync(o);
        if (st.size < 4096 && st.mtimeMs >= since - 2000) rmSync(o);
      } catch { /* 없음·사용 중 */ }
    }
  }
}

/**
 * Word 문서 편집: operations를 순서대로 적용하고 한 번만 저장한다. 하나라도 실패하면 저장하지 않는다.
 * output 생략 = 원본에 저장. → { output, engine: 'word', results: [원자료], lines: [한국어 한 줄씩] }
 */
export async function wordEdit(path, operations, { output } = {}) {
  const full = wordFile(path);
  const ops = normalizeWordOps(operations);
  let out = full;
  if (output) {
    out = resolve(String(output));
    if (!WORD_EXTS.has(extOf(out))) throw new Error(`결과는 Word 문서(.docx/.doc/.rtf/.odt)로만 저장할 수 있습니다: ${basename(out)} — PDF는 office_convert로`);
    if (samePath(out, full)) out = full;
  }
  // 읽기 전용 파일에 그대로 저장하면 Word가 보이지 않는 '다른 이름으로 저장' 창에서 멈춘다: Word를 띄우기 전에 거절
  if (existsSync(out)) {
    try { accessSync(out, constants.W_OK); } catch {
      throw new Error(out === full ? READ_ONLY : `저장할 파일이 읽기 전용입니다: ${out} — 다른 output 경로를 주세요`);
    }
  }
  await requireWord();
  // 찾은 곳 전부(all)를 다루는 작업은 큰 문서에서 오래 걸린다
  const timeoutMs = Math.min(900_000, 300_000 + ops.length * 2_000 + ops.filter((o) => o.all).length * 120_000);
  const started = Date.now();
  let res;
  try {
    res = await comCall({ app: 'word', op: 'edit', path: full, operations: ops, save: out, saveFormat: extOf(out) }, { timeoutMs });
  } catch (e) {
    const msg = String(e?.message ?? e);
    const m = msg.match(/op (\d+) \(([a-z_]*)\): ([\s\S]*)$/);
    if (m) throw new Error(`${m[1]}번째 작업(${m[2]}) 실패: ${koreanOpError(m[3])} — 저장하지 않았습니다`);
    for (const [re, fmt] of SAVE_ERRORS) { const s = msg.match(re); if (s) throw new Error(fmt(s)); }
    // 시간 초과로 Word를 끝내면 그 Word가 만든 '~$' 소유자 파일이 남는다
    if (/응답하지 않아/.test(msg)) removeOwnerFiles([full, out], started);
    throw e;
  }
  const results = Array.isArray(res.results) ? res.results : res.results ? [res.results] : [];
  const lines = ops.map((op, i) => `${i + 1}. ${describe(op, results[i] ?? {})}`);
  lines.push(`저장: ${res.saved ?? out}`);
  return { output: res.saved ?? out, engine: 'word', results, lines };
}
