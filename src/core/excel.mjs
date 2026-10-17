// excel.mjs — Excel 전용 살펴보기·편집 (docs/FEATURES.md 'Excel').
//   app  엔진: 설치된 Excel (COM, scripts/com/excel-ops.ps1) — 모든 작업, 원본 그대로
//   file 엔진: exceljs (Excel 없이) — 표에서 ○인 작업만. 보존하지 못하는 내용(차트·피벗 등)이 있으면 고치지 않는다
// 작업은 순서대로 모두 적용한 뒤 한 번만 저장한다. 하나라도 실패하면 아무것도 저장하지 않는다.
import { existsSync, readFileSync, renameSync, rmSync, statSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { comCall, comAvailableFor, probe } from './com.mjs';
import { parseCsv } from './officefile.mjs';

// ── 기본 ─────────────────────────────────────────────────

export const EXCEL_EXTS = ['xlsx', 'xlsm', 'xls', 'csv', 'ods'];
const APP_SAVE = new Set(['xlsx', 'xlsm', 'xls', 'csv', 'ods']);
const MAX_ROW = 1048576;
const MAX_COL = 16384;
const COM_TIMEOUT = 300_000;
const VT = ExcelJS.ValueType;

const extOf = (p) => extname(p).slice(1).toLowerCase();
const list = (v) => (v == null ? [] : [].concat(v));
const q = (s) => `'${s}'`;

/** 작업 목록: file = 파일 엔진(Excel 없이) 가능 여부 */
export const EXCEL_OPS = {
  add_sheet: { file: true, label: '시트 추가' },
  rename_sheet: { file: true, label: '시트 이름 바꾸기' },
  delete_sheet: { file: true, label: '시트 지우기' },
  copy_sheet: { file: false, label: '시트 복사' },
  move_sheet: { file: true, label: '시트 옮기기' },
  set_cells: { file: true, label: '셀 값·수식' },
  set_range: { file: true, label: '2차원 값 붙여 넣기' },
  clear: { file: true, label: '지우기' },
  insert_rows: { file: true, label: '행 넣기' },
  delete_rows: { file: true, label: '행 지우기' },
  insert_columns: { file: true, label: '열 넣기' },
  delete_columns: { file: true, label: '열 지우기' },
  format: { file: true, label: '서식' },
  column_width: { file: true, label: '열 너비' },
  row_height: { file: true, label: '행 높이' },
  merge: { file: true, label: '병합' },
  unmerge: { file: true, label: '병합 풀기' },
  freeze: { file: true, label: '틀 고정' },
  sort: { file: false, label: '정렬' },
  autofilter: { file: true, label: '자동 필터' },
  add_table: { file: true, label: '표(ListObject)' },
  add_chart: { file: false, label: '차트' },
  conditional_format: { file: true, label: '조건부 서식' },
  data_validation: { file: true, label: '입력 제한' },
  named_range: { file: true, label: '이름 정의' },
  recalculate: { file: false, label: '다시 계산' },
};
// csv로 저장할 때 쓸 수 있는 작업 (csv에는 값만 남는다)
const CSV_OPS = new Set(['set_cells', 'set_range', 'clear', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'sort', 'recalculate']);
// 그중 파일 엔진(csvEdit)이 하는 것 (수식 값은 Excel이 계산해야 하므로 빼고)
const CSV_FILE_OPS = new Set(['set_cells', 'set_range', 'clear', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns']);
const isFormula = (v) => typeof v === 'string' && v.startsWith('=');
const hasFormulaValue = (o) => (o.op === 'set_cells' && Object.values(o.cells).some(isFormula))
  || (o.op === 'set_range' && o.values.some((r) => r.some(isFormula)));

const NEED_EXCEL_TAIL = '— 대신 office_set_cells / office_replace / office_fill / office_create는 쓸 수 있습니다';

// ── 주소 ─────────────────────────────────────────────────

const colNum = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const addr = (c, r) => `${colName(c)}${r}`;

/** 'A1' → {col,row} */
function parseCell(s) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(s ?? '').trim());
  if (!m) return null;
  const col = colNum(m[1]); const row = Number(m[2]);
  if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return null;
  return { col, row };
}

/** 'A1' | 'A1:B2' | 'A:C' | '3:5' → {c1,r1,c2,r2} (전체 열·행은 1~끝) */
function parseArea(s) {
  const t = String(s ?? '').trim().replace(/\$/g, '');
  let m = /^([A-Za-z]{1,3})(\d{1,7})(?::([A-Za-z]{1,3})(\d{1,7}))?$/.exec(t);
  if (m) {
    const a = { c1: colNum(m[1]), r1: Number(m[2]), c2: colNum(m[3] ?? m[1]), r2: Number(m[4] ?? m[2]) };
    return okArea(norm(a));
  }
  m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
  if (m) return okArea(norm({ c1: colNum(m[1]), r1: 1, c2: colNum(m[2]), r2: MAX_ROW, cols: true }));
  m = /^(\d{1,7}):(\d{1,7})$/.exec(t);
  if (m) return okArea(norm({ c1: 1, r1: Number(m[1]), c2: MAX_COL, r2: Number(m[2]), rows: true }));
  return null;
}
function norm(a) {
  return { ...a, c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2), r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2) };
}
const okArea = (a) => (a.c1 >= 1 && a.c2 <= MAX_COL && a.r1 >= 1 && a.r2 <= MAX_ROW ? a : null);
function areaText(a) {
  if (a.cols) return `${colName(a.c1)}:${colName(a.c2)}`;
  if (a.rows) return `${a.r1}:${a.r2}`;
  return a.c1 === a.c2 && a.r1 === a.r2 ? addr(a.c1, a.r1) : `${addr(a.c1, a.r1)}:${addr(a.c2, a.r2)}`;
}
const absText = (a) => (a.cols ? `$${colName(a.c1)}:$${colName(a.c2)}` : a.rows ? `$${a.r1}:$${a.r2}`
  : a.c1 === a.c2 && a.r1 === a.r2 ? `$${colName(a.c1)}$${a.r1}` : `$${colName(a.c1)}$${a.r1}:$${colName(a.c2)}$${a.r2}`);
const intersects = (a, b) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;
const contains = (outer, a) => a.c1 >= outer.c1 && a.c2 <= outer.c2 && a.r1 >= outer.r1 && a.r2 <= outer.r2;
const cellsIn = (a) => (a.c2 - a.c1 + 1) * (a.r2 - a.r1 + 1);

// 시트 이름을 수식에 쓸 때: 영문·숫자·밑줄·한글 등만 있고 숫자로 시작하지 않으며 셀 주소처럼 보이지 않으면 따옴표 없이
function sheetPrefix(name) {
  const plain = /^[A-Za-z_\u00A0-\uFFFF][\w.\u00A0-\uFFFF]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[RrCc](\d|$)/.test(name);
  return `${plain ? name : `'${name.replace(/'/g, "''")}'`}!`;
}
const sameName = (a, b) => a != null && b != null && String(a).toLowerCase() === String(b).toLowerCase();

// ── 수식 속 참조 ──────────────────────────────────────────
// 문자열·구조적 참조([...])·오류 값을 건너뛰고 A1, $A$1:B2, A:C, 3:5, 시트!A1, '시트 이름'!A1 을 찾는다.

const WORDCH = /[\w.$\u00A0-\uFFFF]/;
const RE_AREA = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/y;
const RE_COLS = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const RE_ROWS = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;
const RE_SHEET = /([A-Za-z_\u00A0-\uFFFF][\w.\u00A0-\uFFFF]*)!/y;
const RE_ERROR = /#(?:REF!|N\/A|NAME\?|DIV\/0!|NULL!|NUM!|VALUE!|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)/y;

function boundaryAfter(f, i) {
  const ch = f[i];
  return ch === undefined || !(WORDCH.test(ch) || ch === '(' || ch === '[' || ch === '!' || ch === ':');
}

function matchRef(f, i) {
  RE_AREA.lastIndex = i;
  let m = RE_AREA.exec(f);
  if (m && boundaryAfter(f, RE_AREA.lastIndex)) {
    const r = { kind: 'area', c1: colNum(m[2]), r1: Number(m[4]), a: [!!m[1], !!m[3], !!m[5], !!m[7]] };
    if (m[6]) { r.c2 = colNum(m[6]); r.r2 = Number(m[8]); } else { r.c2 = r.c1; r.r2 = r.r1; r.single = true; r.a[2] = r.a[0]; r.a[3] = r.a[1]; }
    if (r.c1 <= MAX_COL && r.c2 <= MAX_COL && r.r1 >= 1 && r.r2 >= 1 && r.r1 <= MAX_ROW && r.r2 <= MAX_ROW) return { end: RE_AREA.lastIndex, ref: r };
  }
  RE_COLS.lastIndex = i;
  m = RE_COLS.exec(f);
  if (m && boundaryAfter(f, RE_COLS.lastIndex)) {
    const r = { kind: 'cols', c1: colNum(m[2]), c2: colNum(m[4]), a: [!!m[1], false, !!m[3], false] };
    if (r.c1 <= MAX_COL && r.c2 <= MAX_COL) return { end: RE_COLS.lastIndex, ref: r };
  }
  RE_ROWS.lastIndex = i;
  m = RE_ROWS.exec(f);
  if (m && boundaryAfter(f, RE_ROWS.lastIndex)) {
    const r = { kind: 'rows', r1: Number(m[2]), r2: Number(m[4]), a: [false, !!m[1], false, !!m[3]] };
    if (r.r1 >= 1 && r.r2 >= 1 && r.r1 <= MAX_ROW && r.r2 <= MAX_ROW) return { end: RE_ROWS.lastIndex, ref: r };
  }
  return null;
}

function skipString(f, i) {
  for (let j = i + 1; j < f.length; j++) {
    if (f[j] === '"') { if (f[j + 1] === '"') { j++; continue; } return j + 1; }
  }
  return f.length;
}
function skipBracket(f, i) {
  let depth = 0;
  for (let j = i; j < f.length; j++) {
    if (f[j] === '[') depth++;
    else if (f[j] === ']') { depth--; if (depth === 0) return j + 1; }
    else if (f[j] === "'" && f[j + 1] !== undefined) j++; // 구조적 참조 안의 ' 이스케이프
  }
  return f.length;
}
function readQuoted(f, i) {
  let name = '';
  for (let j = i + 1; j < f.length; j++) {
    if (f[j] === "'") { if (f[j + 1] === "'") { name += "'"; j++; continue; } return { name, end: j + 1 }; }
    name += f[j];
  }
  return { name, end: f.length };
}

/** 수식의 참조 목록: [{start, end, prefixEnd, sheet(이름|null), ref, external}] */
export function scanRefs(f) {
  const out = [];
  let i = 0;
  const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') { i = skipString(f, i); continue; }
    if (ch === '[') {
      i = skipBracket(f, i);
      // [1]시트!A1 같은 다른 파일 참조는 건드리지 않는다
      let j = i;
      if (f[j] === "'") j = readQuoted(f, j).end;
      else { RE_SHEET.lastIndex = j; const s = RE_SHEET.exec(f); if (s) j = RE_SHEET.lastIndex - 1; }
      if (f[j] === '!') { const r = matchRef(f, j + 1); i = r ? r.end : j + 1; }
      continue;
    }
    if (ch === '#') {
      RE_ERROR.lastIndex = i;
      const e = RE_ERROR.exec(f);
      if (e) {
        i = RE_ERROR.lastIndex;
        if (e[0] === '#REF!') { const r = matchRef(f, i); if (r) i = r.end; }
        continue;
      }
      i++; continue;
    }
    if (ch === "'") {
      const qd = readQuoted(f, i);
      if (f[qd.end] === '!') {
        const r = matchRef(f, qd.end + 1);
        if (r) { out.push({ start: i, prefixEnd: qd.end + 1, end: r.end, sheet: qd.name, ref: r.ref, external: qd.name.startsWith('[') }); i = r.end; continue; }
      }
      i = qd.end; continue;
    }
    if (WORDCH.test(ch) || ch === '\\') {
      RE_SHEET.lastIndex = i;
      const s = RE_SHEET.exec(f);
      if (s) {
        const r = matchRef(f, RE_SHEET.lastIndex);
        if (r) { out.push({ start: i, prefixEnd: RE_SHEET.lastIndex, end: r.end, sheet: s[1], ref: r.ref }); i = r.end; continue; }
        i = RE_SHEET.lastIndex; continue;
      }
      const r = matchRef(f, i);
      if (r) { out.push({ start: i, prefixEnd: i, end: r.end, sheet: null, ref: r.ref }); i = r.end; continue; }
      while (i < n && (WORDCH.test(f[i]) || f[i] === '\\')) i++;
      continue;
    }
    i++;
  }
  return out;
}

function refText(r) {
  const d = (b) => (b ? '$' : '');
  if (r.kind === 'cols') return `${d(r.a[0])}${colName(r.c1)}:${d(r.a[2])}${colName(r.c2)}`;
  if (r.kind === 'rows') return `${d(r.a[1])}${r.r1}:${d(r.a[3])}${r.r2}`;
  const one = `${d(r.a[0])}${colName(r.c1)}${d(r.a[1])}${r.r1}`;
  return r.single ? one : `${one}:${d(r.a[2])}${colName(r.c2)}${d(r.a[3])}${r.r2}`;
}

/** 참조마다 fn(token) → 새 참조 글자(문자열) | undefined(그대로) 로 바꾼 수식 */
function mapRefs(f, fn) {
  if (typeof f !== 'string' || !f) return f;
  const toks = scanRefs(f);
  if (!toks.length) return f;
  let out = ''; let last = 0;
  for (const t of toks) {
    if (t.external) continue;
    const rep = fn(t);
    if (rep === undefined) continue;
    out += f.slice(last, t.start) + rep;
    last = t.end;
  }
  return out + f.slice(last);
}

// 행·열을 넣거나 지울 때 위치 옮기기. axis 'row'|'col', at부터 count개
function mover(axis, at, count, del) {
  const d1 = at + count - 1;
  const idx = (i) => (del ? (i < at ? i : i > d1 ? i - count : null) : (i >= at ? i + count : i));
  const span = (a, b) => {
    if (!del) return [a >= at ? a + count : a, b >= at ? b + count : b];
    if (a >= at && b <= d1) return null;
    const na = a < at ? a : a > d1 ? a - count : at;
    const nb = b < at ? b : b > d1 ? b - count : at - 1;
    return na <= nb ? [na, nb] : null;
  };
  return { axis, at, count, del, idx, span };
}

// 참조 하나 옮기기 → 새 ref | null(지워짐)
function moveRef(r, mv) {
  const out = { ...r, a: [...r.a] };
  if (mv.axis === 'row') {
    if (r.kind === 'cols') return out;
    const s = mv.span(r.r1, r.r2);
    if (!s) return null;
    [out.r1, out.r2] = s;
    if (out.single && out.r1 !== out.r2) out.single = false;
  } else {
    if (r.kind === 'rows') return out;
    const s = mv.span(r.c1, r.c2);
    if (!s) return null;
    [out.c1, out.c2] = s;
  }
  if (out.r1 > MAX_ROW || out.c1 > MAX_COL) return null;
  out.r2 = Math.min(out.r2 ?? 0, MAX_ROW); out.c2 = Math.min(out.c2 ?? 0, MAX_COL);
  if (r.kind === 'cols') { delete out.r1; delete out.r2; }
  if (r.kind === 'rows') { delete out.c1; delete out.c2; }
  return out;
}

/** target 시트의 행·열이 바뀔 때 수식 고치기 (home: 수식이 있는 시트 — 시트 이름 없는 참조의 주인) */
function shiftFormula(f, home, target, mv) {
  return mapRefs(f, (t) => {
    if (!sameName(t.sheet ?? home, target)) return undefined;
    const nr = moveRef(t.ref, mv);
    const prefix = f.slice(t.start, t.prefixEnd);
    return nr ? prefix + refText(nr) : `${prefix}#REF!`;
  });
}

/** 공유 수식 풀기: 기준 칸에서 (dr, dc)만큼 떨어진 칸의 수식 */
function slideFormula(f, dr, dc) {
  if (!dr && !dc) return f;
  return mapRefs(f, (t) => {
    const r = { ...t.ref, a: [...t.ref.a] };
    if (r.kind !== 'rows') { if (!r.a[0]) r.c1 += dc; if (!r.a[2]) r.c2 += dc; }
    if (r.kind !== 'cols') { if (!r.a[1]) r.r1 += dr; if (!r.a[3]) r.r2 += dr; }
    const bad = (r.kind !== 'rows' && (r.c1 < 1 || r.c2 < 1 || r.c1 > MAX_COL || r.c2 > MAX_COL))
      || (r.kind !== 'cols' && (r.r1 < 1 || r.r2 < 1 || r.r1 > MAX_ROW || r.r2 > MAX_ROW));
    return f.slice(t.start, t.prefixEnd) + (bad ? '#REF!' : refText(r));
  });
}

// 표 이름[...] / 표 이름 → #REF! (지운 시트에 있던 표)
function dropTablesInFormula(f, names) {
  if (typeof f !== 'string' || !f || !names.length) return f;
  const set = new Set(names.map((n) => n.toLowerCase()));
  let out = ''; let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') { const j = skipString(f, i); out += f.slice(i, j); i = j; continue; }
    if (ch === "'") { const j = readQuoted(f, i).end; out += f.slice(i, j); i = j; continue; }
    if (/[A-Za-z_\\\u00A0-\uFFFF]/.test(ch) && !(i > 0 && WORDCH.test(f[i - 1]))) {
      let j = i;
      while (j < f.length && (WORDCH.test(f[j]) || f[j] === '\\')) j++;
      const word = f.slice(i, j);
      if (set.has(word.toLowerCase()) && f[j] !== '!' && f[j] !== '(') {
        out += '#REF!';
        i = f[j] === '[' ? skipBracket(f, j) : j;
        continue;
      }
      out += word; i = j; continue;
    }
    out += ch; i++;
  }
  return out;
}

function renameInFormula(f, from, to) {
  return mapRefs(f, (t) => (t.sheet != null && sameName(t.sheet, from) ? sheetPrefix(to) + refText(t.ref) : undefined));
}
function dropSheetInFormula(f, name) {
  return mapRefs(f, (t) => (t.sheet != null && sameName(t.sheet, name) ? '#REF!' : undefined));
}

// ── 값·색 ────────────────────────────────────────────────

const NAMED_COLORS = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '00B050', blue: '0070C0', yellow: 'FFFF00', orange: 'FFC000',
  gray: '808080', grey: '808080', purple: '7030A0', navy: '002060', lightgray: 'D9D9D9', lightgrey: 'D9D9D9',
};
function color(v, what) {
  const s = String(v ?? '').trim();
  const named = NAMED_COLORS[s.toLowerCase()];
  if (named) return named;
  let m = /^#?([0-9a-f]{6})$/i.exec(s);
  if (m) return m[1].toUpperCase();
  m = /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(s);
  if (m) return `${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}`.toUpperCase();
  throw new Error(`${what} 색이 올바르지 않습니다: ${s} ('#RRGGBB' 형식, 예: '#FF0000')`);
}

function cellValue(v, where) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') { if (!Number.isFinite(v)) throw new Error(`${where}: 숫자가 올바르지 않습니다`); return v; }
  if (typeof v === 'string') { if (v.startsWith('=')) checkFormula(v, where, { cell: true }); return v; }
  if (typeof v === 'boolean') return v;
  throw new Error(`${where}: 값은 글자·숫자·true/false·null만 됩니다`);
}

/**
 * '=…' 수식의 가벼운 문법 확인 (두 엔진 공통). 파일 엔진은 수식을 그대로 쓰므로 Excel이 열지 못하는 파일이 되지 않게
 * 괄호·따옴표 짝, 세미콜론 구분자, 연산자·':'로 끝나는 수식을 미리 거른다. 함수 이름은 보지 않는다 (틀리면 #NAME?).
 */
export function checkFormula(f, where = '수식', { cell = false } = {}) {
  const bad = (why) => { throw new Error(`${where}: 수식이 올바르지 않습니다: ${f} — ${why}${cell ? " (글자 그대로 쓰려면 앞에 ' 를 붙이세요)" : ''}`); };
  const body = String(f).slice(1);
  if (!body.trim()) bad('빈 수식입니다');
  let plain = ''; // 문자열·따옴표 이름·대괄호를 한 글자로 줄인 수식
  let depth = 0; let brace = false;
  for (let i = 0; i < body.length;) {
    const ch = body[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1; let closed = false;
      while (j < body.length) {
        if (body[j] === ch) { if (body[j + 1] === ch) { j += 2; continue; } closed = true; j++; break; }
        j++;
      }
      if (!closed) bad(ch === '"' ? '큰따옴표(")가 닫히지 않았습니다' : "작은따옴표(')가 닫히지 않았습니다");
      plain += ch === '"' ? 'S' : 'Q'; i = j; continue;
    }
    if (ch === '[') {
      let d = 0; let j = i; let closed = false;
      for (; j < body.length; j++) {
        if (body[j] === "'") { j++; continue; }
        if (body[j] === '[') d++;
        else if (body[j] === ']') { d--; if (d === 0) { closed = true; j++; break; } }
      }
      if (!closed) bad('대괄호([ ])의 짝이 맞지 않습니다');
      plain += 'B'; i = j; continue;
    }
    if (ch === '(') depth++;
    else if (ch === ')') { depth--; if (depth < 0) bad('닫는 괄호 )가 여는 괄호보다 많습니다'); }
    else if (ch === '{') { if (brace) bad('중괄호({ })의 짝이 맞지 않습니다'); brace = true; }
    else if (ch === '}') { if (!brace) bad('중괄호({ })의 짝이 맞지 않습니다'); brace = false; }
    else if (ch === ';' && !brace) bad('함수 인수는 쉼표(,)로 구분하세요 (세미콜론 ;은 배열 상수 {1,2;3,4} 안에서만 씀)');
    plain += ch; i++;
  }
  if (depth > 0) bad('여는 괄호 (와 닫는 괄호 )의 수가 맞지 않습니다');
  if (brace) bad('중괄호({ })의 짝이 맞지 않습니다');
  const t = plain.trim();
  if (/[+\-*/^&=<>:,]$/.test(t)) bad('수식이 연산자나 \':\'·\',\'로 끝납니다');
  if (/^[*/^&=<>:,)]/.test(t)) bad('수식이 연산자로 시작합니다');
  if (/:\s*[),]|[(,]\s*:/.test(t)) bad('범위(:)의 앞이나 뒤가 비어 있습니다');
  return f;
}

// ── 작업 검증·정리 (두 엔진 공통) ──────────────────────────

const SHEET_BAD = /[\\/?*[\]:]/;
function sheetName(v, what = '시트 이름') {
  const s = String(v ?? '');
  if (!s.trim()) throw new Error(`${what}이(가) 비어 있습니다`);
  if (s.length > 31) throw new Error(`${what}은(는) 31자까지입니다: ${s}`);
  if (SHEET_BAD.test(s)) throw new Error(`${what}에 쓸 수 없는 글자(\\ / ? * [ ] :)가 있습니다: ${s}`);
  if (s.startsWith("'") || s.endsWith("'")) throw new Error(`${what}은(는) 작은따옴표로 시작하거나 끝날 수 없습니다: ${s}`);
  if (s.toLowerCase() === 'history') throw new Error(`'History'는 Excel이 쓰는 이름이라 시트 이름으로 쓸 수 없습니다`);
  return s;
}
function sheetRef(v) {
  if (v === undefined || v === null || v === '') return undefined;
  if (typeof v === 'number') { if (!Number.isInteger(v) || v < 1) throw new Error(`시트 번호는 1부터입니다: ${v}`); return v; }
  return String(v);
}
// 표·이름 정의에 쓰는 이름
function definedNameOk(s) {
  return s.length <= 255 && /^[A-Za-z_\\\u00A0-\uFFFF][\w.\\\u00A0-\uFFFF]*$/.test(s)
    && !/^[A-Za-z]{1,3}\d{1,7}$/.test(s) && !/^[RrCc]$/.test(s) && !/^[Rr]\d*[Cc]\d*$/.test(s);
}
function areaArg(v, what = 'range', { single = false, multi = false } = {}) {
  const a = parseArea(v);
  if (!a) throw new Error(`${what}가 올바른 범위가 아닙니다: ${v ?? '(없음)'} (예: 'A1', 'A1:C10', 'A:C', '2:5')`);
  if (single && !(a.c1 === a.c2 && a.r1 === a.r2)) throw new Error(`${what}에는 셀 하나를 쓰세요: ${v}`);
  if (multi && a.c1 === a.c2 && a.r1 === a.r2) throw new Error(`${what}는 두 칸 이상이어야 합니다: ${v}`);
  return a;
}
function intArg(v, what, min = 1, max = Infinity) {
  const n = Number(v);
  if (!Number.isInteger(n) || n < min || n > max) throw new Error(`${what}는 ${min}${max < Infinity ? `~${max}` : ' 이상'}의 정수여야 합니다: ${v}`);
  return n;
}
function colArg(v, what = 'at') {
  if (typeof v === 'number') return intArg(v, what, 1, MAX_COL);
  const s = String(v ?? '').trim().replace(/\$/g, '');
  if (/^\d+$/.test(s)) return intArg(Number(s), what, 1, MAX_COL);
  if (/^[A-Za-z]{1,3}$/.test(s) && colNum(s) <= MAX_COL) return colNum(s);
  throw new Error(`${what}는 열 문자('C') 또는 열 번호(3)여야 합니다: ${v}`);
}
function linesArg(v, axis) {
  const s = String(v ?? '').trim().replace(/\$/g, '');
  if (axis === 'col') {
    const m = /^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$/.exec(s);
    if (!m) throw new Error(`columns는 'A' 또는 'A:C' 형식이어야 합니다: ${v}`);
    const a = colNum(m[1]); const b = colNum(m[2] ?? m[1]);
    if (Math.max(a, b) > MAX_COL) throw new Error(`열이 범위를 벗어났습니다: ${v}`);
    return { from: Math.min(a, b), to: Math.max(a, b) };
  }
  const m = /^(\d{1,7})(?::(\d{1,7}))?$/.exec(s);
  if (!m) throw new Error(`rows는 '5' 또는 '1:3' 형식이어야 합니다: ${v}`);
  const a = Number(m[1]); const b = Number(m[2] ?? m[1]);
  if (Math.min(a, b) < 1 || Math.max(a, b) > MAX_ROW) throw new Error(`행이 범위를 벗어났습니다: ${v}`);
  return { from: Math.min(a, b), to: Math.max(a, b) };
}

const ALIGN = ['general', 'left', 'center', 'right', 'justify', 'center_across', 'fill'];
const VALIGN = ['top', 'center', 'bottom', 'justify'];
const BORDERS = ['thin', 'medium', 'thick', 'none', 'outline'];
const CHART_TYPES = ['column', 'bar', 'line', 'pie', 'scatter', 'area'];
const CF_TYPES = ['greater_than', 'less_than', 'between', 'equal', 'text_contains', 'duplicate', 'color_scale', 'data_bar'];
const TABLE_STYLE = /^TableStyle(Light([1-9]|1\d|2[0-1])|Medium([1-9]|1\d|2[0-8])|Dark([1-9]|1[01]))$/;

function oneOf(v, allowed, what) {
  if (!allowed.includes(v)) throw new Error(`${what}는 ${allowed.map((x) => `'${x}'`).join(', ')} 중 하나여야 합니다: ${v}`);
  return v;
}
const has = (o, k) => o != null && Object.prototype.hasOwnProperty.call(o, k) && o[k] !== undefined;

// 작업마다 받는 인수 (모르는 인수는 거절한다 — 조용히 버리면 AI가 원하지 않은 결과를 성공으로 받는다)
const FMT_KEYS = ['bold', 'italic', 'underline', 'font', 'size', 'color', 'fill', 'number_format', 'align', 'valign', 'wrap', 'border'];
export const EXCEL_OP_KEYS = {
  add_sheet: ['name', 'after'], rename_sheet: ['sheet', 'name'], delete_sheet: ['sheet'], copy_sheet: ['sheet', 'name'],
  move_sheet: ['sheet', 'position'], set_cells: ['sheet', 'cells'], set_range: ['sheet', 'range', 'values'],
  clear: ['sheet', 'range', 'what'], insert_rows: ['sheet', 'at', 'count'], delete_rows: ['sheet', 'at', 'count'],
  insert_columns: ['sheet', 'at', 'count'], delete_columns: ['sheet', 'at', 'count'], format: ['sheet', 'range', ...FMT_KEYS],
  column_width: ['sheet', 'columns', 'width'], row_height: ['sheet', 'rows', 'height'], merge: ['sheet', 'range'], unmerge: ['sheet', 'range'],
  freeze: ['sheet', 'cell'], sort: ['sheet', 'range', 'by', 'order', 'header'], autofilter: ['sheet', 'range'],
  add_table: ['sheet', 'range', 'name', 'style'], add_chart: ['sheet', 'range', 'type', 'title', 'at', 'width', 'height', 'series_in'],
  conditional_format: ['sheet', 'range', 'rule'], data_validation: ['sheet', 'range', 'list', 'type', 'min', 'max'],
  named_range: ['name', 'range', 'sheet'], recalculate: [],
};
export const EXCEL_RULE_KEYS = ['type', 'value', 'value2', 'fill', 'font_color'];
// 명령줄에서 예전부터 받던 다른 이름
const KEY_ALIASES = { column_width: ['range', 'auto'], row_height: ['range', 'auto'] };
function checkKeys(obj, allowed, what, quiet = []) {
  const extra = Object.keys(obj).filter((k) => !allowed.includes(k) && !quiet.includes(k) && obj[k] !== undefined);
  if (extra.length) throw new Error(`모르는 인수입니다: ${extra.join(', ')} — ${what}에 쓸 수 있는 인수: ${allowed.join(', ') || '(없음)'}`);
}

/** 작업 하나 검증·정리 → 정리된 작업 (오류는 한국어) */
function normalizeOp(raw) {
  if (!raw || typeof raw !== 'object') throw new Error('작업은 {op: "...", ...} 형식이어야 합니다');
  const kind = raw.op;
  if (!EXCEL_OPS[kind]) throw new Error(`알 수 없는 작업입니다: ${kind ?? '(op 없음)'} (가능: ${Object.keys(EXCEL_OPS).join(', ')})`);
  checkKeys(raw, EXCEL_OP_KEYS[kind], kind, ['op', ...(KEY_ALIASES[kind] ?? [])]);
  const op = { op: kind };
  const withSheet = () => { const s = sheetRef(raw.sheet); if (s !== undefined) op.sheet = s; };
  switch (kind) {
    case 'add_sheet':
      op.name = sheetName(raw.name);
      if (has(raw, 'after')) op.after = sheetRef(raw.after);
      break;
    case 'rename_sheet':
      op.sheet = sheetRef(raw.sheet); if (op.sheet === undefined) throw new Error('sheet(바꿀 시트)가 필요합니다');
      op.name = sheetName(raw.name);
      break;
    case 'delete_sheet':
      op.sheet = sheetRef(raw.sheet); if (op.sheet === undefined) throw new Error('sheet(지울 시트)가 필요합니다');
      break;
    case 'copy_sheet':
      op.sheet = sheetRef(raw.sheet); if (op.sheet === undefined) throw new Error('sheet(복사할 시트)가 필요합니다');
      op.name = sheetName(raw.name, '새 시트 이름');
      break;
    case 'move_sheet':
      op.sheet = sheetRef(raw.sheet); if (op.sheet === undefined) throw new Error('sheet(옮길 시트)가 필요합니다');
      op.position = intArg(raw.position, 'position');
      break;
    case 'set_cells': {
      withSheet();
      const cells = raw.cells;
      if (!cells || typeof cells !== 'object' || Array.isArray(cells) || !Object.keys(cells).length) throw new Error('cells: { "A1": 값 } 형식으로 하나 이상 주세요');
      op.cells = {};
      for (const [k, v] of Object.entries(cells)) {
        const c = parseCell(k);
        if (!c) throw new Error(`셀 주소가 올바르지 않습니다: ${k} (예: A1, B12 — 범위는 set_range)`);
        op.cells[addr(c.col, c.row)] = cellValue(v, k);
      }
      break;
    }
    case 'set_range': {
      withSheet();
      const a = areaArg(raw.range);
      if (a.cols || a.rows) throw new Error(`range에는 시작 셀을 쓰세요: ${raw.range} (예: 'A1')`);
      let rows = raw.values;
      if (!Array.isArray(rows) || !rows.length) throw new Error('values: [[...], [...]] 2차원 배열이 필요합니다');
      if (!rows.some(Array.isArray)) rows = [rows];
      op.values = rows.map((r, i) => list(r).map((v, j) => cellValue(v, `values[${i}][${j}]`)));
      const nc = Math.max(1, ...op.values.map((r) => r.length));
      if (a.r1 + op.values.length - 1 > MAX_ROW || a.c1 + nc - 1 > MAX_COL) throw new Error('붙여 넣을 값이 시트 끝을 넘습니다');
      if (op.values.length * nc > 1_000_000) throw new Error('한 번에 붙여 넣을 수 있는 칸은 100만 개까지입니다');
      op.range = addr(a.c1, a.r1);
      break;
    }
    case 'clear':
      withSheet();
      op.range = areaText(areaArg(raw.range));
      op.what = raw.what === undefined ? 'contents' : oneOf(raw.what, ['all', 'contents', 'formats'], 'what');
      break;
    case 'insert_rows': case 'delete_rows': {
      withSheet();
      op.at = intArg(raw.at, 'at(행 번호)', 1, MAX_ROW);
      op.count = raw.count === undefined ? 1 : intArg(raw.count, 'count', 1, MAX_ROW);
      if (op.at + op.count - 1 > MAX_ROW) throw new Error('행 범위가 시트 끝을 넘습니다');
      break;
    }
    case 'insert_columns': case 'delete_columns': {
      withSheet();
      op.at = colArg(raw.at, 'at(열)');
      op.count = raw.count === undefined ? 1 : intArg(raw.count, 'count', 1, MAX_COL);
      if (op.at + op.count - 1 > MAX_COL) throw new Error('열 범위가 시트 끝을 넘습니다');
      op.columns = `${colName(op.at)}:${colName(op.at + op.count - 1)}`;
      break;
    }
    case 'format': {
      withSheet();
      op.range = areaText(areaArg(raw.range));
      let any = false;
      for (const k of ['bold', 'italic', 'underline', 'wrap']) {
        if (has(raw, k)) { if (typeof raw[k] !== 'boolean') throw new Error(`${k}는 true/false여야 합니다`); op[k] = raw[k]; any = true; }
      }
      if (has(raw, 'font')) { if (!String(raw.font).trim()) throw new Error('font(글꼴 이름)가 비어 있습니다'); op.font = String(raw.font); any = true; }
      if (has(raw, 'size')) { const s = Number(raw.size); if (!(s >= 1 && s <= 409)) throw new Error(`size(pt)는 1~409입니다: ${raw.size}`); op.size = s; any = true; }
      if (has(raw, 'color')) { op.color = color(raw.color, '글자'); any = true; }
      if (raw.fill === null || (typeof raw.fill === 'string' && /^(none|없음|transparent)$/i.test(raw.fill))) { op.fill = 'none'; any = true; }
      else if (has(raw, 'fill')) { op.fill = color(raw.fill, '채우기'); any = true; }
      if (has(raw, 'number_format')) { op.number_format = String(raw.number_format); any = true; }
      if (has(raw, 'align')) { op.align = oneOf(raw.align, ALIGN, 'align'); any = true; }
      if (has(raw, 'valign')) { op.valign = oneOf(raw.valign === 'middle' ? 'center' : raw.valign, VALIGN, 'valign'); any = true; }
      if (has(raw, 'border')) { op.border = oneOf(raw.border, BORDERS, 'border'); any = true; }
      if (!any) throw new Error('바꿀 서식이 없습니다 (bold, italic, underline, font, size, color, fill, number_format, align, valign, wrap, border)');
      break;
    }
    case 'column_width': {
      withSheet();
      const c = linesArg(raw.columns ?? raw.range, 'col');
      op.columns = `${colName(c.from)}:${colName(c.to)}`;
      const w = raw.width ?? (raw.auto ? 'auto' : undefined);
      if (w === 'auto') op.width = 'auto';
      else { const n = Number(w); if (!(n >= 0 && n <= 255)) throw new Error(`width는 0~255(글자 수) 또는 'auto'여야 합니다: ${w}`); op.width = n; }
      break;
    }
    case 'row_height': {
      withSheet();
      const r = linesArg(raw.rows ?? raw.range, 'row');
      op.rows = `${r.from}:${r.to}`;
      const h = raw.height ?? (raw.auto ? 'auto' : undefined);
      if (h === 'auto') op.height = 'auto';
      else { const n = Number(h); if (!(n >= 0 && n <= 409)) throw new Error(`height는 0~409(pt) 또는 'auto'여야 합니다: ${h}`); op.height = n; }
      break;
    }
    case 'merge':
      withSheet();
      op.range = areaText(areaArg(raw.range, 'range', { multi: true }));
      break;
    case 'unmerge':
      withSheet();
      op.range = areaText(areaArg(raw.range));
      break;
    case 'freeze':
      withSheet();
      if (raw.cell === null || raw.cell === undefined || raw.cell === '' || raw.cell === false) op.cell = null;
      else { const c = parseCell(raw.cell); if (!c) throw new Error(`cell이 올바른 셀 주소가 아닙니다: ${raw.cell} (예: 'B2', 해제는 null)`); op.cell = addr(c.col, c.row); }
      break;
    case 'sort': {
      withSheet();
      const a = areaArg(raw.range, 'range', { multi: true });
      op.range = areaText(a);
      const by = String(raw.by ?? '').trim().replace(/\$/g, '');
      if (!/^[A-Za-z]{1,3}$/.test(by)) throw new Error(`by는 정렬 기준 열 문자여야 합니다: ${raw.by ?? '(없음)'} (예: 'B')`);
      if (colNum(by) < a.c1 || colNum(by) > a.c2) throw new Error(`정렬 기준 열 ${by.toUpperCase()}가 범위 ${op.range} 안에 없습니다`);
      op.by = by.toUpperCase();
      op.order = raw.order === undefined ? 'asc' : oneOf(raw.order, ['asc', 'desc'], 'order');
      if (has(raw, 'header')) op.header = !!raw.header;
      break;
    }
    case 'autofilter':
      withSheet();
      op.range = raw.range === null || raw.range === undefined || raw.range === '' ? null : areaText(areaArg(raw.range));
      break;
    case 'add_table': {
      withSheet();
      const a = areaArg(raw.range);
      if (a.cols || a.rows) throw new Error(`표 범위는 'A1:C10'처럼 주세요: ${raw.range}`);
      op.range = areaText(a);
      if (has(raw, 'name')) { const nm = String(raw.name); if (!definedNameOk(nm)) throw new Error(`표 이름으로 쓸 수 없습니다: ${nm} (글자나 _로 시작, 띄어쓰기 없이, 셀 주소 모양 불가)`); op.name = nm; }
      if (has(raw, 'style')) op.style = String(raw.style);
      break;
    }
    case 'add_chart': {
      withSheet();
      const a = areaArg(raw.range, 'range', { multi: true });
      op.range = areaText(a);
      op.type = oneOf(raw.type ?? 'column', CHART_TYPES, 'type');
      if (has(raw, 'title')) op.title = String(raw.title);
      if (has(raw, 'at')) { const c = parseCell(raw.at); if (!c) throw new Error(`at은 차트 왼쪽 위 셀이어야 합니다: ${raw.at}`); op.at = addr(c.col, c.row); }
      for (const k of ['width', 'height']) {
        if (has(raw, k)) { const n = Number(raw[k]); if (!(n >= 20 && n <= 4000)) throw new Error(`${k}(pt)는 20~4000입니다: ${raw[k]}`); op[k] = n; }
      }
      if (has(raw, 'series_in')) op.series_in = oneOf(raw.series_in, ['columns', 'rows'], 'series_in');
      break;
    }
    case 'conditional_format': {
      withSheet();
      op.range = areaText(areaArg(raw.range));
      const r = raw.rule;
      if (!r || typeof r !== 'object' || Array.isArray(r)) throw new Error('rule: {type, value?, value2?, fill?, font_color?}가 필요합니다');
      checkKeys(r, EXCEL_RULE_KEYS, 'rule');
      const rule = { type: oneOf(r.type, CF_TYPES, 'rule.type') };
      const val = (k) => {
        const v = r[k];
        if (typeof v === 'number' && Number.isFinite(v)) return v;
        if (typeof v === 'string' && v !== '') return rule.type !== 'text_contains' && v.startsWith('=') ? checkFormula(v, `rule.${k}`) : v;
        throw new Error(`rule.${k}(비교할 값)가 필요합니다`);
      };
      if (['greater_than', 'less_than', 'equal', 'between', 'text_contains'].includes(rule.type)) rule.value = val('value');
      if (rule.type === 'between') rule.value2 = val('value2');
      if (rule.type === 'text_contains') rule.value = String(rule.value);
      if (has(r, 'fill')) rule.fill = color(r.fill, '채우기');
      if (has(r, 'font_color')) rule.font_color = color(r.font_color, '글자');
      op.rule = rule;
      break;
    }
    case 'data_validation': {
      withSheet();
      op.range = areaText(areaArg(raw.range));
      if (has(raw, 'list')) {
        if (typeof raw.list === 'string') {
          if (!raw.list.startsWith('=')) throw new Error("list는 ['가','나'] 목록이거나 '=$E$1:$E$5' 같은 범위여야 합니다");
          op.list = checkFormula(raw.list, 'list');
        } else {
          const items = list(raw.list).map((v) => String(v));
          if (!items.length) throw new Error('list가 비어 있습니다');
          if (items.some((s) => s.includes(','))) throw new Error('목록 항목에는 쉼표(,)를 쓸 수 없습니다 — 항목이 많거나 쉼표가 있으면 셀 범위(\'=$E$1:$E$5\')를 쓰세요');
          if (items.some((s) => s.includes('"'))) throw new Error('목록 항목에는 큰따옴표(")를 쓸 수 없습니다 — 항목을 시트 칸에 적고 셀 범위(\'=$E$1:$E$5\')를 쓰세요');          if (items.join(',').length > 255) throw new Error('목록 글자가 255자를 넘습니다 — 셀 범위(\'=$E$1:$E$5\')를 쓰세요');
          op.list = items;
        }
      } else {
        op.type = oneOf(raw.type, ['whole', 'decimal'], 'type(list가 없을 때)');
        for (const k of ['min', 'max']) {
          if (has(raw, k)) { const n = Number(raw[k]); if (!Number.isFinite(n)) throw new Error(`${k}가 숫자가 아닙니다: ${raw[k]}`); if (op.type === 'whole' && !Number.isInteger(n)) throw new Error(`정수 제한의 ${k}는 정수여야 합니다: ${raw[k]}`); op[k] = n; }
        }
        if (!has(op, 'min') && !has(op, 'max')) throw new Error('min 또는 max가 필요합니다');
        if (has(op, 'min') && has(op, 'max') && op.min > op.max) throw new Error(`min(${op.min})이 max(${op.max})보다 큽니다`);
      }
      break;
    }
    case 'named_range': {
      const nm = String(raw.name ?? '');
      if (!definedNameOk(nm)) throw new Error(`이름으로 쓸 수 없습니다: ${nm || '(없음)'} (글자나 _로 시작, 띄어쓰기 없이, 셀 주소 모양 불가)`);
      op.name = nm;
      const text = String(raw.range ?? '').trim().replace(/^=/, '');
      const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text);
      const a = parseArea(m ? m[3] : text);
      if (!a) throw new Error(`range는 '시트!A1:B2' 형식이어야 합니다: ${raw.range ?? '(없음)'}`);
      op.area = absText(a);
      const sh = m ? (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]) : sheetRef(raw.sheet);
      if (sh !== undefined) op.sheet = sh;
      break;
    }
    case 'recalculate':
      break;
    default:
      break;
  }
  return op;
}

/** 작업 목록 검증·정리. 오류: "N번째 작업(op): ..." */
export function normalizeOperations(operations) {
  const ops = list(operations);
  if (!ops.length) throw new Error('operations가 비어 있습니다');
  return ops.map((raw, i) => {
    try { return normalizeOp(raw); } catch (e) {
      throw new Error(`${i + 1}번째 작업(${raw?.op ?? '?'}) 오류: ${e.message} — 저장하지 않았습니다`);
    }
  });
}

// ── 결과 줄 ──────────────────────────────────────────────

const CHART_KO = { column: '세로 막대형', bar: '가로 막대형', line: '꺾은선형', pie: '원형', scatter: '분산형', area: '영역형' };
const ALIGN_KO = { general: '일반', left: '왼쪽', center: '가운데', right: '오른쪽', justify: '양쪽', center_across: '선택 영역 가운데', fill: '채우기' };
const VALIGN_KO = { top: '위', center: '가운데', bottom: '아래', justify: '양쪽' };
const BORDER_KO = { thin: '가는 선', medium: '보통 선', thick: '굵은 선', none: '없음', outline: '바깥쪽만' };
const CLEAR_KO = { all: '모두', contents: '내용', formats: '서식' };
function cfText(r) {
  const v = (x) => (typeof x === 'string' && !x.startsWith('=') ? `'${x}'` : x);
  switch (r.type) {
    case 'greater_than': return `${v(r.value)}보다 큼`;
    case 'less_than': return `${v(r.value)}보다 작음`;
    case 'between': return `${v(r.value)}~${v(r.value2)} 사이`;
    case 'equal': return `${v(r.value)}과(와) 같음`;
    case 'text_contains': return `'${r.value}' 포함`;
    case 'duplicate': return '중복 값';
    case 'color_scale': return '색조(빨강-노랑-초록)';
    case 'data_bar': return `데이터 막대${r.fill ? ` #${r.fill}` : ''}`;
    default: return r.type;
  }
}
const on = (d) => (d?.sheet ? ` (시트 ${q(d.sheet)})` : '');

function describe(op, d = {}) {
  switch (op.op) {
    case 'add_sheet': return `시트 추가: ${q(d.sheet ?? op.name)}${d.index ? ` (${d.index}번째)` : ''}`;
    case 'rename_sheet': return `시트 이름 바꿈: ${q(d.from ?? op.sheet)} → ${q(d.sheet ?? op.name)}`;
    case 'delete_sheet': return `시트 지움: ${q(d.sheet ?? op.sheet)}`;
    case 'copy_sheet': return `시트 복사: ${q(d.from ?? op.sheet)} → ${q(d.sheet ?? op.name)}${d.index ? ` (${d.index}번째)` : ''}`;
    case 'move_sheet': return `시트 옮김: ${q(d.sheet ?? op.sheet)} → ${op.position}번째`;
    case 'set_cells': return `셀 ${Object.keys(op.cells).length}개 씀: ${Object.keys(op.cells).slice(0, 8).join(', ')}${Object.keys(op.cells).length > 8 ? ' …' : ''}${on(d)}`;
    case 'set_range': {
      const nc = Math.max(1, ...op.values.map((r) => r.length));
      const end = addr(parseCell(op.range).col + nc - 1, parseCell(op.range).row + op.values.length - 1);
      return `${op.range}:${end}에 ${op.values.length}행×${nc}열 붙여 넣음${on(d)}`;
    }
    case 'clear': return `${op.range} ${CLEAR_KO[op.what]} 지움${on(d)}`;
    case 'insert_rows': return `${op.at}행 앞에 ${op.count}행 넣음 — 원래 ${op.at}행부터는 ${op.at + op.count}행부터로 밀림${on(d)}`;
    case 'delete_rows': return `${op.at}${op.count > 1 ? `~${op.at + op.count - 1}` : ''}행 지움 — 그 아래 행은 ${op.count}행씩 당겨짐${on(d)}`;
    case 'insert_columns': return `${colName(op.at)}열 앞에 ${op.count}열 넣음 — 원래 ${colName(op.at)}열부터는 ${colName(op.at + op.count)}열부터로 밀림${on(d)}`;
    case 'delete_columns': return `${op.columns.replace(/^(\w+):\1$/, '$1')}열 지움 — 그 오른쪽 열은 ${op.count}열씩 당겨짐${on(d)}`;
    case 'format': {
      const parts = [];
      if (has(op, 'bold')) parts.push(op.bold ? '굵게' : '굵게 해제');
      if (has(op, 'italic')) parts.push(op.italic ? '기울임' : '기울임 해제');
      if (has(op, 'underline')) parts.push(op.underline ? '밑줄' : '밑줄 해제');
      if (op.font) parts.push(`글꼴 ${op.font}`);
      if (op.size) parts.push(`${op.size}pt`);
      if (op.color) parts.push(`글자색 #${op.color}`);
      if (op.fill) parts.push(op.fill === 'none' ? '채우기 없음' : `채우기 #${op.fill}`);
      if (has(op, 'number_format')) parts.push(`표시 형식 ${op.number_format}`);
      if (op.align) parts.push(`가로 ${ALIGN_KO[op.align]}`);
      if (op.valign) parts.push(`세로 ${VALIGN_KO[op.valign]}`);
      if (has(op, 'wrap')) parts.push(op.wrap ? '줄 바꿈' : '줄 바꿈 해제');
      if (op.border) parts.push(`테두리 ${BORDER_KO[op.border]}`);
      return `서식 ${op.range}: ${parts.join(', ')}${on(d)}`;
    }
    case 'column_width': return `열 너비 ${op.columns}: ${op.width === 'auto' ? `내용에 맞춤${d.approx ? '(근사)' : ''}` : op.width}${on(d)}`;
    case 'row_height': return `행 높이 ${op.rows}: ${op.height === 'auto' ? `내용에 맞춤${d.approx ? '(근사)' : ''}` : `${op.height}pt`}${on(d)}`;
    case 'merge': return `병합 ${op.range}${d.discarded ? ` — 왼쪽 위 값만 남고 값이 있던 칸 ${d.discarded}개는 지워짐` : ''}${on(d)}`;
    case 'unmerge': return `병합 풂 ${op.range}${on(d)}`;
    case 'freeze': {
      if (!op.cell) return `틀 고정 해제${on(d)}`;
      const c = parseCell(op.cell);
      const parts = [c.row > 1 ? `위 ${c.row - 1}행` : '', c.col > 1 ? `왼쪽 ${c.col - 1}열` : ''].filter(Boolean);
      return parts.length ? `틀 고정 ${op.cell} (${parts.join('·')})${on(d)}` : `틀 고정 해제 (A1)${on(d)}`;
    }
    case 'sort': return `정렬 ${op.range}: ${op.by}열 ${op.order === 'desc' ? '내림차순' : '오름차순'}${has(op, 'header') ? (op.header ? ' (첫 행 머리글 제외)' : ' (첫 행 포함)') : ' (머리글은 Excel이 판단)'}${on(d)}`;
    case 'autofilter': return op.range ? `자동 필터 ${d.range ?? op.range}${on(d)}` : `자동 필터 해제${on(d)}`;
    case 'add_table': return `표 만듦: ${d.name ?? op.name} (${d.range ?? op.range}, ${d.style ?? op.style ?? 'TableStyleMedium2'})${on(d)}`;
    case 'add_chart': {
      const series = d.series ? `, 계열 ${d.series}개(${d.seriesIn === 'rows' ? '행마다' : '열마다'})` : '';
      const title = op.title ? `, 제목 '${op.title}'` : d.title ? `, 제목 '${d.title}'(자동)` : ', 제목 없음';
      return `차트 만듦: ${CHART_KO[op.type]}${d.name ? ` '${d.name}'` : ''} (원본 ${op.range}${d.at ? `, 위치 ${d.at}` : ''}${series}${title})${on(d)}`;
    }
    case 'conditional_format': return `조건부 서식 ${op.range}: ${cfText(op.rule)}${d.note ? ` — ${d.note}` : ''}${on(d)}`;
    case 'data_validation': {
      if (op.list) return `입력 제한 ${op.range}: 목록 ${typeof op.list === 'string' ? op.list : `(${op.list.join(', ')})`}${on(d)}`;
      const kind = op.type === 'whole' ? '정수' : '소수';
      const rng = has(op, 'min') && has(op, 'max') ? `${op.min}~${op.max}` : has(op, 'min') ? `${op.min} 이상` : `${op.max} 이하`;
      return `입력 제한 ${op.range}: ${kind} ${rng}${on(d)}`;
    }
    case 'named_range': return `이름 정의: ${op.name} = ${d.refersTo ? String(d.refersTo).replace(/^=/, '') : op.area}`;
    case 'recalculate': return '모두 다시 계산함';
    default: return op.op;
  }
}

// ── 오류 (COM 코드 → 한국어) ───────────────────────────────

const COM_CODES = {
  E_SHEET: (x) => `시트 없음: ${x}`,
  E_RANGE: (x) => `범위가 올바르지 않습니다: ${x}`,
  E_SHEET_EXISTS: (x) => `같은 이름의 시트가 이미 있습니다: ${x}`,
  E_SHEET_NAME: (x) => `시트 이름으로 쓸 수 없습니다: ${x}`,
  E_LAST_SHEET: (x) => `보이는 시트가 하나뿐이라 지울 수 없습니다: ${x}`,
  E_POSITION: (x, n) => `position이 범위를 벗어났습니다: ${x} (1~${n})`,
  E_MERGED_PART: (x, tl) => `${x}는 병합된 칸의 일부라 쓸 수 없습니다 — 병합 영역의 왼쪽 위 칸 ${tl}에 쓰세요`,
  E_FORMULA: (x) => `수식이 올바르지 않습니다: ${x} (함수 이름은 영어, 인수 구분은 쉼표(,))`,
  E_TABLE_OVERLAP: (x) => `표 '${x}'와 겹칩니다`,
  E_TABLE_MERGED: (x) => `병합된 칸이 있는 범위에는 표를 만들 수 없습니다: ${x}`,
  E_TABLE_NAME: (x) => `표 이름으로 쓸 수 없거나 이미 있는 이름입니다: ${x}`,
  E_TABLE_STYLE: (x) => `표 스타일이 없습니다: ${x} (예: TableStyleMedium2, TableStyleLight9)`,
  E_MERGE_OVERLAP: (x) => `이미 병합된 영역 ${x}가 범위 밖으로 걸쳐 있습니다 — 범위를 넓히거나 먼저 unmerge 하세요`,
  E_SORT_KEY: (x) => `정렬 기준 열 ${x}가 범위 안에 없습니다`,
  E_NAME: (x) => `이름으로 쓸 수 없습니다: ${x}`,
  E_UNKNOWN_OP: (x) => `알 수 없는 작업입니다: ${x}`,
  E_TABLE_EXISTS: (x) => `이미 있는 표·이름입니다: ${x}`,
  E_TABLE_HEADER: (t, r) => `표 '${t}'의 머리글 행(${r}행)은 지울 수 없습니다 — 표 전체를 지우거나 데이터 행만 지우세요`,
  E_TABLE_TOTALS: (t, r) => `표 '${t}'의 요약 행(${r}행)은 지울 수 없습니다`,
  E_MERGE_PARTIAL: (x) => `병합된 영역의 일부만 지울 수 없습니다(${x}) — 범위에 병합 영역 전체를 넣으세요`,
  E_SORT_MERGED: (x) => `병합된 칸이 있는 범위는 정렬할 수 없습니다: ${x} — 병합을 풀거나 범위를 바꾸세요`,
  E_PUSH_OFF: (x) => `시트 끝에 있는 내용이 밀려나므로 ${x === 'col' ? '열' : '행'}을 넣을 수 없습니다`,
  E_CSV_SHEETS: (a, b) => `csv에는 시트 하나의 값만 저장되는데 작업이 여러 시트('${a}', '${b}')를 고칩니다 — 시트마다 따로 저장하거나 xlsx로 저장하세요`,
};
// Excel이 영어로 내는 흔한 오류 → 한국어
const EXCEL_MESSAGES = [
  [/merged cells need to be the same size/i, '병합된 칸의 크기가 서로 달라 할 수 없습니다 — 범위 안의 병합을 풀거나 범위를 바꾸세요'],
  [/push (?:non-empty cells|objects) off|can't insert new cells|Insert property of the Range class/i, '시트 끝에 있는 내용(값·서식·개체)이 밀려나므로 넣을 수 없습니다'],
  [/can't do that to a merged cell|change part of a merged cell/i, '병합된 칸의 일부에는 할 수 없습니다 — 범위에 병합 영역 전체를 넣으세요'],
  [/change part of an array/i, '배열 수식의 일부만 바꿀 수 없습니다 — 배열 수식 범위 전체를 고치세요'],
  [/Delete property of the Range class/i, 'Excel이 이 행·열을 지우지 못했습니다 (표·병합·배열 수식의 일부인지 확인하세요)'],
  [/Sort method of the Range class|sort reference is not valid/i, 'Excel이 정렬하지 못했습니다 (범위와 기준 열을 확인하세요)'],
  [/Save(?:As)? (?:property|method) of the Workbook class|save failed/i, '문서를 저장하지 못했습니다 (저장할 파일 이름·폴더·쓰기 권한을 확인하고, 그 파일이 다른 프로그램에서 열려 있지 않은지 확인하세요)'],
  [/0x800A03EC/i, 'Excel이 이 값이나 범위를 받아들이지 않았습니다 (주소·수식·목록을 확인하세요)'],
];
function excelMessage(s) {
  for (const [re, ko] of EXCEL_MESSAGES) if (re.test(s)) return ko;
  return null;
}
function comCodeText(s) {
  const parts = String(s).split('|');
  const fmt = COM_CODES[parts[0]];
  if (fmt) return fmt(...parts.slice(1));
  return excelMessage(String(s)) ?? `Excel 오류: ${String(s).trim()}`;
}
function opError(i, op, msg) {
  return new Error(`${i}번째 작업(${op}) 실패: ${msg} — 저장하지 않았습니다`);
}
function comEditError(e) {
  const msg = e?.message ?? '';
  const m = /opfail\|(\d+)\|([^|]*)\|([\s\S]*)$/.exec(msg);
  if (m) return opError(Number(m[1]), m[2], comCodeText(m[3]));
  // 작업 밖(저장·csv 시트 확인 등)에서 난 오류
  const code = /\b(E_[A-Z_]+(?:\|[\s\S]*)?)$/.exec(msg);
  if (code && COM_CODES[code[1].split('|')[0]]) return new Error(`${comCodeText(code[1])} — 저장하지 않았습니다`);
  const ko = excelMessage(msg);
  return ko ? new Error(ko) : e;
}

// ── 엔진 고르기 ──────────────────────────────────────────

async function excelReady(ext) {
  try { return await comAvailableFor(ext); } catch { return false; }
}
// why: '… 수 없습니다'처럼 끝난 문장. 뒤에 이유(Excel 없음 / 확인 실패)와 대안을 붙인다
async function needExcel(why, ext) {
  const p = await probe().catch((e) => ({ error: e.message }));
  if (p.error) return `${why} — Excel을 쓸 수 있는지 확인하지 못했습니다 (${p.error}) ${NEED_EXCEL_TAIL}`;
  // engine:'file'을 골랐지만 Excel이 있는 경우
  if (ext && p.excel && (await excelReady(ext))) return `${why} — 이 PC에는 Excel이 있으니 engine:'auto' 또는 'app'으로 하세요`;
  return `${why} — Excel이 설치돼 있어야 합니다 ${NEED_EXCEL_TAIL}`;
}

function mustExist(path) {
  if (!existsSync(path)) throw new Error(`파일이 없습니다: ${path}`);
  return resolve(path);
}
function excelOnly(path) {
  const ext = extOf(path);
  if (!EXCEL_EXTS.includes(ext)) throw new Error(`엑셀 문서(${EXCEL_EXTS.map((x) => `.${x}`).join(', ')})만 다룹니다: ${basename(path)}`);
  return ext;
}

// ── 파일 엔진: 읽기·보존 확인 ─────────────────────────────

const decodeXml = (s) => s.replace(/&(amp|lt|gt|quot|apos);|&#(\d+);|&#x([0-9a-f]+);/gi,
  (m, n, d, h) => ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[n] ?? String.fromCodePoint(d ? Number(d) : parseInt(h, 16))));
const encodeXml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const attrOf = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? decodeXml(m[1]) : undefined; };
const resolvePart = (base, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = dirname(base).split('/').filter((p) => p && p !== '.');
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
  return parts.join('/');
};
const relsPath = (part) => `${dirname(part)}/_rels/${basename(part)}.rels`.replace(/^\.\//, '');
async function relsOf(zip, part) {
  const f = zip.file(relsPath(part));
  if (!f) return [];
  const xml = await f.async('string');
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
    id: attrOf(m[0], 'Id'), type: (attrOf(m[0], 'Type') ?? '').split('/').pop(), target: attrOf(m[0], 'Target') ?? '',
    external: attrOf(m[0], 'TargetMode') === 'External',
  })).map((r) => ({ ...r, part: r.external ? null : resolvePart(part, r.target) }));
}

function damaged(path, cause) {
  return new Error(`엑셀 파일이 손상되었거나 암호가 걸려 있어 열 수 없습니다: ${basename(path)}`, { cause });
}

// 원본 ZIP에서 읽어 두는 것: 시트 부품, 이름 정의 원문, 표 머리글 여부, 차트 수, 보존할 수 없는 내용
async function rawBook(zip, path) {
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!wbXml) throw damaged(path);
  const rels = await relsOf(zip, 'xl/workbook.xml');
  const byId = new Map(rels.map((r) => [r.id, r]));
  const sheets = [...wbXml.matchAll(/<(?:\w+:)?sheet\b[^>]*\/?>/g)].map((m) => {
    const id = /\sr:id="([^"]*)"|\s\w+:id="([^"]*)"/.exec(m[0]);
    const rel = byId.get(id?.[1] ?? id?.[2]) ?? {};
    return { name: attrOf(m[0], 'name') ?? '', part: rel.part, kind: rel.type };
  });
  const names = [];
  const dn = /<(?:\w+:)?definedNames>([\s\S]*?)<\/(?:\w+:)?definedNames>/.exec(wbXml);
  if (dn) {
    for (const m of dn[1].matchAll(/<(?:\w+:)?definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?definedName>)/g)) {
      const attrs = m[1];
      const local = attrOf(attrs, 'localSheetId');
      const rest = attrs.replace(/\s(name|localSheetId)="[^"]*"/g, '');
      names.push({ name: attrOf(attrs, 'name') ?? '', local: local === undefined ? null : Number(local), attrs: rest, text: decodeXml(m[2] ?? '') });
    }
  }
  const files = Object.keys(zip.files);
  const lost = new Set();
  const flag = (re, label) => { if (files.some((f) => re.test(f))) lost.add(label); };
  flag(/^xl\/(charts|chartsheets)\//, '차트');
  flag(/^xl\/(pivotTables|pivotCache)\//, '피벗 테이블');
  flag(/^xl\/(slicers|slicerCaches)\//, '슬라이서');
  flag(/^xl\/(timelines|timelineCaches)\//, '시간 표시 막대');
  flag(/^xl\/externalLinks\//, '다른 파일 참조');
  flag(/^xl\/vbaProject\.bin$/, '매크로(VBA)');
  flag(/^xl\/threadedComments\//, '스레드 메모');
  flag(/^xl\/model\//, '데이터 모델');
  flag(/^xl\/(queryTables\/|connections\.xml$)/, '외부 데이터 연결');
  flag(/^xl\/(ctrlProps|activeX)\//, '양식 컨트롤·ActiveX');
  flag(/^xl\/embeddings\//, '포함된 개체(OLE)');
  flag(/^xl\/richData\//, '셀 안 그림·데이터 형식');
  if (/<(?:\w+:)?workbookProtection\b/.test(wbXml)) lost.add('통합 문서 보호');
  if (/<(?:\w+:)?fileSharing\b/.test(wbXml)) lost.add('쓰기 암호');
  if (sheets.some((s) => s.kind && s.kind !== 'worksheet')) lost.add('차트 시트·매크로 시트');
  const stylesXml = (await zip.file('xl/styles.xml')?.async('string')) ?? '';
  if (/<(?:\w+:)?tableStyles\b[^>]*\scount="[1-9]/.test(stylesXml)) lost.add('사용자 지정 표 스타일');
  if (/<(?:\w+:)?indexedColors\b/.test(stylesXml)) lost.add('사용자 지정 색 팔레트');
  // 기본 글꼴(Normal 스타일): exceljs는 늘 Calibri로 쓰므로 저장할 때 원래 것으로 돌려놓는다
  const font0 = /<(?:\w+:)?fonts\b[^>]*>\s*(<(?:\w+:)?font\b[\s\S]*?<\/(?:\w+:)?font>)/.exec(stylesXml)?.[1] ?? null;

  const internalLinks = new Set();
  const SUPPORTED_CF = new Set(['expression', 'cellIs', 'top10', 'aboveAverage', 'dataBar', 'colorScale', 'iconSet', 'containsText',
    'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod']);
  for (const s of sheets) {
    if (s.kind !== 'worksheet' || !s.part || !zip.file(s.part)) continue;
    const xml = await zip.file(s.part).async('string');
    s.xml = xml;
    if (/<x14:sparklineGroups\b/.test(xml)) lost.add('스파크라인');
    if (/<x14:dataValidations\b/.test(xml)) lost.add('다른 시트를 참조하는 입력 제한');
    if (/<(?:\w+:)?legacyDrawingHF\b/.test(xml)) lost.add('머리글·바닥글 그림');
    if (/<(?:\w+:)?(oleObjects|controls)\b/.test(xml)) lost.add('양식 컨트롤·포함된 개체');
    if (/<(?:\w+:)?protectedRanges\b/.test(xml)) lost.add('편집 허용 범위');
    if (/<(?:\w+:)?c\b[^>]*\scm="/.test(xml)) lost.add('동적 배열 수식');
    for (const m of xml.matchAll(/<(?:\w+:)?cfRule\b[^>]*\stype="(\w+)"/g)) if (!SUPPORTED_CF.has(m[1])) lost.add('조건부 서식(중복 값·시작 문자 등)');
    for (const m of xml.matchAll(/<x14:cfRule\b[^>]*\stype="(\w+)"/g)) if (!['dataBar', 'iconSet'].includes(m[1])) lost.add('조건부 서식(확장)');
    for (const m of xml.matchAll(/<(?:\w+:)?dataValidation\b[^>]*\stype="date"[^>]*>([\s\S]*?)<\/(?:\w+:)?dataValidation>/g)) {
      if ([...m[1].matchAll(/<(?:\w+:)?formula\d>([\s\S]*?)<\//g)].some((f) => !/^-?[\d.]+$/.test(f[1].trim()))) lost.add('날짜 입력 제한(수식)');
    }
    for (const m of xml.matchAll(/<(?:\w+:)?hyperlink\b[^>]*>/g)) {
      const loc = attrOf(m[0], 'location');
      if (!loc) continue;
      internalLinks.add(loc);
      if (attrOf(m[0], 'r:id')) lost.add('바깥 파일 안 위치로 가는 링크');
    }
    // 차트·도형
    s.charts = [];
    for (const r of await relsOf(zip, s.part)) {
      if (r.type !== 'drawing' || !r.part || !zip.file(r.part)) continue;
      const dxml = await zip.file(r.part).async('string');
      const drels = await relsOf(zip, r.part);
      const chartIds = new Set(drels.filter((x) => x.type === 'chart').map((x) => x.id));
      for (const g of dxml.matchAll(/<xdr:graphicFrame\b[\s\S]*?<\/xdr:graphicFrame>/g)) {
        const rid = /<c:chart\b[^>]*\sr:id="([^"]*)"/.exec(g[0]);
        if (rid && chartIds.has(rid[1])) s.charts.push({ name: attrOf(/<xdr:cNvPr\b[^>]*>/.exec(g[0])?.[0] ?? '', 'name') ?? '' });
        else lost.add('SmartArt 등 그래픽 개체');
      }
      if (/<xdr:(sp|cxnSp|grpSp)\b/.test(dxml)) lost.add('도형·글상자');
      if (/<mc:AlternateContent\b/.test(dxml)) lost.add('최신 차트·도형');
    }
    for (const r of await relsOf(zip, s.part)) {
      if (r.type !== 'table' || !r.part || !zip.file(r.part)) continue;
      const txml = await zip.file(r.part).async('string');
      if (/totalsRowFunction="custom"/.test(txml)) lost.add('표 요약 행 사용자 수식');
    }
  }
  // 표: exceljs가 잘못 읽는 것을 원문에서 — 머리글 행 수(headerRowCount 생략 = 머리글 있음), 열 목록
  // (<calculatedColumnFormula>가 든 열 뒤의 열을 exceljs가 버린다), 필터 단추 숨김, 자동 필터가 없는 표
  const tables = new Map();
  for (const f of files.filter((x) => /^xl\/tables\/[^/]+\.xml$/.test(x))) {
    const xml = await zip.file(f).async('string');
    const tag = /<(?:\w+:)?table\b[^>]*>/.exec(xml)?.[0] ?? '';
    const name = attrOf(tag, 'name');
    if (!name) continue;
    const cols = [];
    const tc = /<(?:\w+:)?tableColumns\b[^>]*>([\s\S]*?)<\/(?:\w+:)?tableColumns>/.exec(xml)?.[1] ?? '';
    for (const m of tc.matchAll(/<(?:\w+:)?tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?tableColumn>)/g)) {
      const inner = m[2] ?? '';
      const calc = /<(?:\w+:)?calculatedColumnFormula\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?calculatedColumnFormula>/.exec(inner);
      const rest = inner.replace(/<(?:\w+:)?(calculatedColumnFormula|totalsRowFormula)\b[\s\S]*?<\/(?:\w+:)?\1>/g, '').trim();
      if (rest) lost.add('표 열 확장 정보(XML 연결 등)');
      cols.push({
        name: attrOf(m[1], 'name') ?? '', totalsRowFunction: attrOf(m[1], 'totalsRowFunction'), totalsRowLabel: attrOf(m[1], 'totalsRowLabel'),
        calc: calc ? { attrs: calc[1].replace(/\s+xmlns(?::\w+)?="[^"]*"/g, ''), formula: decodeXml(calc[2]) } : null,
      });
    }
    const af = /<(?:\w+:)?autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?autoFilter>)/.exec(xml)?.[0];
    const hidden = new Set();
    for (const m of (af ?? '').matchAll(/<(?:\w+:)?filterColumn\b[^>]*>/g)) if (attrOf(m[0], 'hiddenButton') === '1') hidden.add(Number(attrOf(m[0], 'colId')));
    tables.set(name, { header: attrOf(tag, 'headerRowCount') !== '0', cols, autoFilter: !!af, hidden });
  }
  return { sheets, names, lost: [...lost], tables, wbXml, font0, internalLinks };
}

// exceljs가 잘못 읽는 것 바로잡기
function fixAfterLoad(wb, raw) {
  const wsRaw = raw.sheets.filter((s) => s.kind === 'worksheet');
  wb.worksheets.forEach((ws, i) => {
    for (const t of Object.values(ws.tables ?? {})) {
      const m = t.table ?? t.model;
      const info = raw.tables.get(m?.name);
      if (!m || !info) continue;
      m.headerRow = info.header;
      // 열 목록은 원문대로 (계산 열 수식은 저장할 때 postProcess가 다시 넣는다)
      if (info.cols.length) {
        m.columns = info.cols.map((c, k) => ({
          name: c.name, totalsRowFunction: c.totalsRowFunction, totalsRowLabel: c.totalsRowLabel,
          filterButton: info.autoFilter && !info.hidden.has(k),
          ...(c.calc ? { _calc: { ...c.calc } } : {}),
        }));
      }
      if (!info.autoFilter) m._noAutoFilter = true;
    }
    // 정수·소수 제한의 셀 참조 수식은 숫자로 바뀌어 NaN이 된다 → 원문으로
    const xml = wsRaw[i]?.xml ?? '';
    for (const m of xml.matchAll(/<(?:\w+:)?dataValidation\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?dataValidation>/g)) {
      const first = (attrOf(m[1], 'sqref') ?? '').split(/\s+/)[0]?.split(':')[0];
      const dv = first && ws.dataValidations.model[first];
      if (!dv?.formulae) continue;
      const rawF = [...m[2].matchAll(/<(?:\w+:)?formula(\d)>([\s\S]*?)<\/(?:\w+:)?formula\d>/g)].map((f) => decodeXml(f[2]));
      dv.formulae = dv.formulae.map((v, k) => (typeof v === 'number' && Number.isNaN(v) && rawF[k] !== undefined ? rawF[k] : v));
    }
    // 문서 안 링크(location만 있는 hyperlink)는 exceljs가 버린다 → 셀 값에 다시 붙이고 저장할 때 location으로 되돌린다
    for (const m of xml.matchAll(/<(?:\w+:)?hyperlink\b[^>]*>/g)) {
      const loc = attrOf(m[0], 'location');
      const ref = attrOf(m[0], 'ref');
      if (!loc || !ref || attrOf(m[0], 'r:id')) continue;
      const a = parseArea(ref);
      if (!a) continue;
      const cell = ws.getCell(a.r1, a.c1);
      const v = cell.value;
      const text = typeof v === 'string' ? v : v?.richText ? v.richText.map((r) => r.text).join('') : null;
      if (text === null && v != null) throw new OpError(`문서 안 링크가 걸린 ${cell.address} 칸은 글자가 아니라 파일 엔진으로 보존할 수 없습니다 — Excel로 하세요`);
      cell.value = { text: text ?? attrOf(m[0], 'display') ?? loc, hyperlink: loc, ...(attrOf(m[0], 'tooltip') ? { tooltip: attrOf(m[0], 'tooltip') } : {}) };
    }
    // 내용 없는 숨긴 행·서식 행은 exceljs가 버린다 → 높이를 주어 남긴다
    const dflt = ws.properties?.defaultRowHeight || 15;
    ws._rows.forEach((row) => {
      if (!row || row.height || row._cells.some(Boolean)) return;
      if (row.hidden || (row.style && Object.keys(row.style).length)) row.height = dflt;
    });
  });
}

// 공유 수식은 기준 칸을 고치거나 옮기면 깨지므로 처음부터 칸마다의 수식으로 푼다
function unshareFormulas(wb) {
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== 'object' || cell.type !== VT.Formula) return;
        if (v.sharedFormula) {
          const master = ws.findCell(v.sharedFormula);
          const mf = master?.value?.formula;
          if (!mf) return;
          const a = parseCell(v.sharedFormula); const b = parseCell(cell.address);
          cell.value = { formula: slideFormula(mf, b.row - a.row, b.col - a.col), result: v.result };
        }
      });
    });
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (cell.type === VT.Formula && v?.shareType === 'shared') cell.value = { formula: v.formula, result: v.result };
      });
    });
  }
}

async function loadBook(path) {
  const buf = readFileSync(path);
  let zip;
  try { zip = await JSZip.loadAsync(buf); } catch (e) { throw damaged(path, e); }
  const raw = await rawBook(zip, path);
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(await forExceljs(buf)); } catch (e) { throw damaged(path, e); }
  return { wb, zip, raw };
}

// exceljs는 <tableColumn> 안에 자식 요소(계산 열 수식 등)가 있으면 그 뒤의 열을 버리고, 자동 필터 열이 더 많으면 읽다가 멈춘다
// → 자식 요소를 뺀 사본을 읽힌다 (계산 열 수식은 rawBook이 원문에서 읽어 두었다가 저장할 때 다시 넣는다)
async function forExceljs(buf) {
  const zip = await JSZip.loadAsync(buf);
  let changed = false;
  for (const part of Object.keys(zip.files).filter((x) => /^xl\/tables\/[^/]+\.xml$/.test(x))) {
    const xml = await zip.file(part).async('string');
    const out = xml.replace(/<((?:\w+:)?tableColumn)\b([^>]*?)(?<!\/)>[\s\S]*?<\/\1>/g, (_m, tag, attrs) => `<${tag}${attrs}/>`);
    if (out !== xml) { zip.file(part, out); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'nodebuffer' }) : buf;
}

// ── 파일 엔진: 살펴보기 ──────────────────────────────────

function usedBounds(ws) {
  let b = null;
  ws._rows.forEach((row) => {
    if (!row) return;
    row._cells.forEach((cell) => {
      if (!cell) return;
      const r = row.number; const c = cell.col;
      if (!b) b = { c1: c, r1: r, c2: c, r2: r };
      else { b.c1 = Math.min(b.c1, c); b.c2 = Math.max(b.c2, c); b.r1 = Math.min(b.r1, r); b.r2 = Math.max(b.r2, r); }
    });
  });
  return b;
}

function filterRef(af) {
  if (!af) return null;
  if (typeof af === 'string') return af;
  const s = (x) => (typeof x === 'string' ? x : addr(x.column ?? x.col, x.row));
  return af.from && af.to ? `${s(af.from)}:${s(af.to)}` : null;
}

// 활성 시트 번호(0부터): bookViews의 activeTab, 없으면 tabSelected인 시트, 없으면 첫 시트
function activeIndex(wb) {
  const t = wb.views?.[0]?.activeTab;
  if (Number.isInteger(t) && t >= 0 && t < wb.worksheets.length) return t;
  const sel = wb.worksheets.findIndex((w) => w.views?.[0]?.tabSelected);
  return sel >= 0 ? sel : 0;
}

function freezeOf(ws) {
  const v = ws.views?.[0];
  if (!v || v.state !== 'frozen') return null;
  const xs = v.xSplit || 0; const ys = v.ySplit || 0;
  if (!xs && !ys) return null;
  return addr(xs + 1, ys + 1);
}

async function fileInspect(path) {
  const { wb, raw } = await loadBook(path);
  const wsRaw = raw.sheets.filter((s) => s.kind === 'worksheet');
  const formulas = []; let formulaCount = 0;
  const sheets = wb.worksheets.map((ws, i) => {
    const b = usedBounds(ws);
    const cfCount = (ws.conditionalFormattings ?? []).reduce((n, cf) => n + (cf.rules?.length ?? 0), 0);
    const dvSet = new Set(Object.values(ws.dataValidations?.model ?? {}).filter(Boolean));
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type !== VT.Formula) return;
      formulaCount++;
      if (formulas.length < 20) formulas.push({ sheet: ws.name, cell: cell.address, formula: `=${plainFormula(cell.formula)}` });
    }));
    return {
      index: i + 1, name: ws.name, visible: ws.state === 'hidden' ? 'hidden' : ws.state === 'veryHidden' ? 'veryHidden' : 'visible',
      usedRange: b ? areaText(b) : null, rows: b ? b.r2 - b.r1 + 1 : 0, cols: b ? b.c2 - b.c1 + 1 : 0,
      freeze: freezeOf(ws),
      tables: Object.values(ws.tables ?? {}).map((t) => { const m = t.table ?? t.model; return { name: m.name, range: m.tableRef, style: m.style?.theme ?? null, totals: !!m.totalsRow }; }),
      charts: wsRaw[i]?.charts ?? [],
      merges: Object.keys(ws._merges ?? {}).length,
      autoFilter: filterRef(ws.autoFilter),
      conditionalFormats: cfCount,
      validations: dvSet.size,
    };
  });
  const activeIdx = activeIndex(wb);
  const visibleNames = raw.names.filter((n) => !/\shidden="(1|true)"/.test(n.attrs));
  return {
    engine: 'file',
    activeSheet: wb.worksheets[activeIdx]?.name ?? wb.worksheets[0]?.name ?? null,
    sheets,
    names: visibleNames.map((n) => ({ name: n.name, refersTo: `=${n.text}`, scope: n.local === null ? null : raw.sheets[n.local]?.name ?? null })),
    hiddenNames: raw.names.length - visibleNames.length,
    formulas, formulaCount,
    fileEngineLimits: raw.lost,
  };
}

async function csvInspect(path) {
  const { rows } = readCsvFile(path);
  const nc = rows.reduce((n, r) => Math.max(n, r.length), 0);
  const name = basename(path, extname(path));
  return {
    engine: 'file', activeSheet: name,
    sheets: [{ index: 1, name, visible: 'visible', usedRange: rows.length && nc ? `A1:${addr(nc, rows.length)}` : null, rows: rows.length, cols: nc,
      freeze: null, tables: [], charts: [], merges: 0, autoFilter: null, conditionalFormats: 0, validations: 0 }],
    names: [], hiddenNames: 0, formulas: [], formulaCount: 0,
  };
}

// ── 파일 엔진: 편집 ──────────────────────────────────────

class OpError extends Error {}
const fail = (msg) => { throw new OpError(msg); };

function sheetOf(st, ref, { what = '시트' } = {}) {
  const all = st.wb.worksheets;
  const names = () => all.map((w) => q(w.name)).join(', ');
  if (ref === undefined || ref === null || ref === '') return all[0] ?? fail('시트가 없습니다');
  if (typeof ref === 'number') return all[ref - 1] ?? fail(`${what} 번호가 범위를 벗어났습니다: ${ref} (1~${all.length})`);
  const hit = all.find((w) => sameName(w.name, ref));
  if (hit) return hit;
  if (/^\d+$/.test(ref)) return all[Number(ref) - 1] ?? fail(`${what} 번호가 범위를 벗어났습니다: ${ref} (1~${all.length})`);
  return fail(`${what} 없음: ${ref} (있는 시트: ${names()})`);
}

function setOrder(st, ordered) {
  ordered.forEach((ws, i) => { ws.orderNo = i; });
}

const allCells = function* (ws) {
  for (const row of ws._rows) {
    if (!row) continue;
    for (const cell of row._cells) if (cell) yield cell;
  }
};

// 모든 수식(셀·조건부 서식·입력 제한·이름 정의)에 fn(수식, 수식이 있는 시트 이름) 적용
function rewriteAllFormulas(st, fn) {
  for (const ws of st.wb.worksheets) {
    const home = ws.name;
    for (const cell of allCells(ws)) {
      if (cell.type !== VT.Formula) continue;
      const v = cell.value;
      if (!v?.formula) continue;
      const nf = fn(v.formula, home);
      if (nf !== v.formula) cell.value = { ...v, formula: nf };
    }
    for (const cf of ws.conditionalFormattings ?? []) {
      for (const rule of cf.rules ?? []) if (rule.formulae) rule.formulae = rule.formulae.map((f) => (typeof f === 'string' ? fn(f, home) : f));
    }
    const seen = new Set();
    for (const dv of Object.values(ws.dataValidations?.model ?? {})) {
      if (!dv || seen.has(dv)) continue;
      seen.add(dv);
      if (dv.formulae) dv.formulae = dv.formulae.map((f) => (typeof f === 'string' ? fn(f, home) : f));
    }
    for (const { m } of tablesOf(ws)) for (const c of m.columns ?? []) if (c._calc) c._calc.formula = fn(c._calc.formula, home);
  }
  for (const n of st.names) n.text = fn(n.text, n.ws ? n.ws.name : null);
}

// 표 계산 열 수식 옮기기: 수식은 표의 첫 데이터 행 기준이라, 그 행에 행을 넣거나 그 행을 지우면 기준을 다시 맞춘다
function shiftCalcFormulas(st, target, mv) {
  for (const w of st.wb.worksheets) {
    for (const { m, area } of tablesOf(w)) {
      const anchor = area && sameName(w.name, target) && mv.axis === 'row' ? area.r1 + (m.headerRow === false ? 0 : 1) : null;
      for (const c of m.columns ?? []) {
        if (!c._calc) continue;
        let f = c._calc.formula;
        if (anchor && !mv.del && mv.at === anchor) f = slideFormula(shiftFormula(f, w.name, target, mv), -mv.count, 0);
        else {
          if (anchor && mv.del && mv.at <= anchor && anchor <= mv.at + mv.count - 1) f = slideFormula(f, mv.at + mv.count - anchor, 0);
          f = shiftFormula(f, w.name, target, mv);
        }
        if (/#REF!/.test(f)) delete c._calc; else c._calc.formula = f;
      }
    }
  }
}

function markFormulaChange(st) { st.recalc = true; }

// Excel 2010 뒤에 생긴 함수는 파일 안에 '_xlfn.' 접두사를 붙여 적어야 한다 (없으면 Excel이 #NAME?으로 보임)
const XLFN = new Set(('ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR '
  + 'BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST COMBINA CONCAT CONFIDENCE.NORM '
  + 'CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT '
  + 'F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR '
  + 'FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH '
  + 'IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.SNGL NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST '
  + 'NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI '
  + 'POISSON.DIST QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ RRI SEC SECH SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T '
  + 'T.DIST.RT T.INV T.INV.2T T.TEST TEXTAFTER TEXTBEFORE TEXTJOIN UNICHAR UNICODE VALUETOTEXT VAR.P VAR.S WEBSERVICE WEIBULL.DIST '
  + 'WORKDAY.INTL XLOOKUP XMATCH XOR Z.TEST').split(' '));
// 여러 칸으로 넘치는 동적 배열·LAMBDA 계열은 파일 안에 넘침 정보·매개변수 접두사가 따로 있어야 해서 파일 엔진으로는 쓰지 않는다
const XLFN_EXCEL_ONLY = new Set(('FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPCOLS WRAPROWS TAKE DROP '
  + 'CHOOSECOLS CHOOSEROWS EXPAND MUNIT MODE.MULT LET LAMBDA MAP REDUCE SCAN BYROW BYCOL MAKEARRAY GROUPBY PIVOTBY').split(' '));

/** 파일 엔진이 쓰는 수식(= 없이): 새 함수에 _xlfn. 접두사, 동적 배열·LAMBDA 함수는 거절 */
function fileFormula(f) {
  let out = '';
  for (let i = 0; i < f.length;) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } j++; break; } j++; }
      out += f.slice(i, j); i = j; continue;
    }
    if (ch === '[') { const j = skipBracket(f, i); out += f.slice(i, j); i = j; continue; }
    if (/[A-Za-z_]/.test(ch) && !(i > 0 && /[\w.$ -￿]/.test(f[i - 1]))) {
      let j = i;
      while (j < f.length && /[\w. -￿]/.test(f[j])) j++;
      const word = f.slice(i, j);
      if (/^\s*\(/.test(f.slice(j))) {
        const up = word.toUpperCase();
        if (XLFN_EXCEL_ONLY.has(up)) fail(`${up} 함수(동적 배열·LAMBDA)는 파일 엔진으로 쓸 수 없습니다 — Excel이 설치된 PC에서 engine:'app'으로 하세요`);
        if (XLFN.has(up)) { out += `_xlfn.${up}`; i = j; continue; }
      }
      out += word; i = j; continue;
    }
    out += ch; i++;
  }
  return out;
}
// 살펴보기에 보일 때는 접두사를 뺀다
const plainFormula = (f) => String(f ?? '').replace(/_xlfn\.(?:_xlws\.)?|_xlpm\./gi, '');
// 값이 바뀌어 수식 결과가 낡을 수 있는 작업
const VALUE_OPS = new Set(['set_cells', 'set_range', 'clear', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns',
  'delete_sheet', 'rename_sheet', 'merge', 'named_range', 'add_table']);

function putValue(st, cell, v) {
  if (v === null) { cell.value = null; return; }
  if (typeof v === 'string') {
    if (v.startsWith('=')) {
      if (v.length < 2) fail('빈 수식입니다');
      cell.value = { formula: fileFormula(v.slice(1)) };
      markFormulaChange(st);
      return;
    }
    cell.value = v.startsWith("'") ? v.slice(1) : v;
    return;
  }
  cell.value = v;
}

function checkWritable(ws, a) {
  const c = ws.findCell(a.r1, a.c1);
  if (c?.isMerged && c.master !== c) fail(`${c.address}는 병합된 칸의 일부라 쓸 수 없습니다 — 병합 영역의 왼쪽 위 칸 ${c.master.address}에 쓰세요`);
}

// 범위 안의 셀 (전체 행·열이면 사용 범위까지만)
function* cellsOf(ws, a, { create = true } = {}) {
  let { c1, c2, r1, r2 } = a;
  if (a.cols || a.rows || cellsIn(a) > 2_000_000) {
    const b = usedBounds(ws);
    if (!b) return;
    c1 = Math.max(c1, b.c1); c2 = Math.min(c2, b.c2); r1 = Math.max(r1, b.r1); r2 = Math.min(r2, b.r2);
    if (c1 > c2 || r1 > r2) return;
  }
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = create ? ws.getCell(r, c) : ws.findCell(r, c);
      if (cell) yield cell;
    }
  }
}

function mergesOf(ws) {
  return Object.values(ws._merges ?? {}).map((m) => parseArea(m.range ?? m.shortRange ?? m.toString())).filter(Boolean);
}
function tablesOf(ws) {
  return Object.values(ws.tables ?? {}).map((t) => ({ t, m: t.table ?? t.model })).map((x) => ({ ...x, area: parseArea(x.m.tableRef) }));
}
function tableOverlap(ws, a) {
  const hit = tablesOf(ws).find((x) => x.area && intersects(x.area, a));
  return hit?.m.name ?? null;
}

const solidFill = (hex) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${hex}` }, bgColor: { indexed: 64 } });

function patchStyle(style, op) {
  const s = { ...(style ?? {}) };
  if (['bold', 'italic', 'underline', 'font', 'size', 'color'].some((k) => has(op, k))) {
    const f = { ...(s.font ?? {}) };
    if (has(op, 'bold')) f.bold = op.bold;
    if (has(op, 'italic')) f.italic = op.italic;
    if (has(op, 'underline')) { if (op.underline) f.underline = true; else delete f.underline; }
    if (op.font) { f.name = op.font; delete f.scheme; delete f.family; delete f.charset; }
    if (op.size) f.size = op.size;
    if (op.color) f.color = { argb: `FF${op.color}` };
    s.font = f;
  }
  if (op.fill) s.fill = op.fill === 'none' ? { type: 'pattern', pattern: 'none' } : solidFill(op.fill);
  if (has(op, 'number_format')) s.numFmt = op.number_format;
  if (op.align || op.valign || has(op, 'wrap')) {
    const al = { ...(s.alignment ?? {}) };
    if (op.align) { if (op.align === 'general') delete al.horizontal; else al.horizontal = op.align === 'center_across' ? 'centerContinuous' : op.align; }
    if (op.valign) al.vertical = op.valign === 'center' ? 'middle' : op.valign;
    if (has(op, 'wrap')) al.wrapText = op.wrap;
    s.alignment = al;
  }
  return s;
}

function borderStyle(style, sides) {
  const s = { ...(style ?? {}) };
  s.border = { ...(s.border ?? {}), ...sides };
  return s;
}

// 'auto' 너비·높이 근사: 글자 수(한글 등 넓은 글자는 2칸)와 글자 크기로
const wide = (ch) => /[\u1100-\u11FF\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/.test(ch);
function textWidth(t) { let n = 0; for (const ch of String(t)) n += wide(ch) ? 2 : 1; return n; }
function displayText(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (cell.type === VT.Formula) return v.result == null ? '' : typeof v.result === 'object' ? String(v.result.error ?? '') : String(v.result);
  if (cell.type === VT.RichText) return v.richText.map((r) => r.text).join('');
  if (cell.type === VT.Hyperlink) return String(v.text ?? '');
  if (cell.type === VT.Number) {
    const fmt = cell.numFmt ?? '';
    if (/%/.test(fmt)) return `${(v * 100).toFixed(/0\.(0+)%/.exec(fmt)?.[1].length ?? 0)}%`;
    const dec = /0\.(0+)/.exec(fmt)?.[1].length;
    const s = dec !== undefined ? v.toFixed(dec) : String(v);
    return /#,##/.test(fmt) ? s.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : s;
  }
  if (cell.type === VT.Date) return 'yyyy-mm-dd';
  return String(cell.text ?? v);
}

// ── 행·열 넣기·지우기 (exceljs spliceRows는 병합·수식·조건부 서식·표·그림 위치를 옮기지 않으므로 직접) ──

function moveSqref(sqref, mv) {
  const out = [];
  for (const part of String(sqref ?? '').split(/\s+/).filter(Boolean)) {
    const a = parseArea(part);
    if (!a) { out.push(part); continue; }
    const r = moveArea(a, mv);
    if (r) out.push(areaText(r));
  }
  return out.join(' ');
}
function moveArea(a, mv) {
  const out = { ...a };
  if (mv.axis === 'row') {
    if (a.cols) return out;
    const s = mv.span(a.r1, a.r2); if (!s) return null;
    [out.r1, out.r2] = [s[0], Math.min(s[1], MAX_ROW)];
    if (out.r1 > MAX_ROW) return null;
  } else {
    if (a.rows) return out;
    const s = mv.span(a.c1, a.c2); if (!s) return null;
    [out.c1, out.c2] = [s[0], Math.min(s[1], MAX_COL)];
    if (out.c1 > MAX_COL) return null;
  }
  return out;
}

function spliceSheet(st, ws, axis, at, count, del) {
  const mv = mover(axis, at, count, del);
  const target = ws.name;
  const lo = at; const hi = at + count - 1;
  const inBand = (a) => (axis === 'row' ? [a.r1, a.r2] : [a.c1, a.c2]);
  shiftCalcFormulas(st, target, mv);

  // 표: 머리글 행을 지우거나 표 안에 열을 넣고 지우는 것은 Excel처럼 다룬다
  const tableFix = [];
  for (const { m, area } of tablesOf(ws)) {
    if (!area) continue;
    const [b1, b2] = inBand(area);
    if (del && lo <= b1 && hi >= b2) { tableFix.push({ m, drop: true }); continue; }
    if (axis === 'row' && del && m.headerRow !== false && lo <= b1 && hi >= b1) fail(`표 '${m.name}'의 머리글 행(${b1}행)은 지울 수 없습니다 — 표 전체를 지우거나 데이터 행만 지우세요`);
    if (axis === 'row' && del && m.totalsRow && lo <= b2 && hi >= b2) fail(`표 '${m.name}'의 요약 행은 지울 수 없습니다`);
    if (axis === 'col') {
      const left = area.c1; const right = area.c2;
      if (del) {
        const from = Math.max(lo, left); const to = Math.min(hi, right);
        if (from <= to) tableFix.push({ m, delCols: [from - left, to - from + 1] });
      } else if (lo > left && lo <= right) tableFix.push({ m, insCols: [lo - left, count] });
    }
  }
  for (const f of tableFix) if (f.drop) delete ws.tables[f.m.name];

  const m = ws.model;
  const sheetFormula = (f) => shiftFormula(f, target, target, mv);
  const relocate = (cm, row, col) => {
    const out = { ...cm, address: addr(col, row) };
    if (out.type === VT.Merge) return { address: out.address, type: VT.Null, style: out.style };
    if (out.type === VT.Formula) {
      if (out.formula) out.formula = sheetFormula(out.formula);
      if (out.ref) { const ra = parseArea(out.ref); const nr = ra && moveArea(ra, mv); if (nr) out.ref = areaText(nr); }
    }
    return out;
  };
  const rows = [];
  for (const r of m.rows ?? []) {
    if (!r) continue;
    if (axis === 'row') {
      const n = mv.idx(r.number);
      if (n === null || n > MAX_ROW) continue;
      rows.push({ ...r, number: n, cells: r.cells.map((c) => { const p = parseCell(c.address); return relocate(c, n, p.col); }) });
    } else {
      const cells = [];
      for (const c of r.cells) {
        const p = parseCell(c.address);
        const nc = mv.idx(p.col);
        if (nc === null || nc > MAX_COL) continue;
        cells.push(relocate(c, r.number, nc));
      }
      cells.sort((x, y) => parseCell(x.address).col - parseCell(y.address).col);
      const cols = cells.map((c) => parseCell(c.address).col);
      rows.push({ ...r, cells, min: cols.length ? Math.min(...cols) : r.min, max: cols.length ? Math.max(...cols) : r.max });
    }
  }
  m.rows = rows.sort((x, y) => x.number - y.number);
  if (axis === 'col' && m.cols) {
    const per = [];
    for (const c of m.cols) for (let k = c.min; k <= c.max; k++) { const n = mv.idx(k); if (n !== null && n <= MAX_COL) per.push({ ...c, min: n, max: n }); }
    m.cols = per.sort((x, y) => x.min - y.min);
  }
  m.mergeCells = (m.merges ?? []).map((s) => { const a = parseArea(s); const r = a && moveArea(a, mv); return r && cellsIn(r) > 1 ? areaText(r) : null; }).filter(Boolean);
  const dv = {};
  const dvSeen = new Set();
  for (const [k, v] of Object.entries(m.dataValidations ?? {})) {
    if (!v) continue;
    if (!dvSeen.has(v)) { dvSeen.add(v); if (v.formulae) v.formulae = v.formulae.map((f) => (typeof f === 'string' ? sheetFormula(f) : f)); }
    const a = parseArea(k);
    const r = a && moveArea(a, mv);
    if (r) dv[areaText(r)] = v;
  }
  // exceljs는 입력 제한을 칸마다 따로 갖는다: 넣은 줄의 앞뒤 칸이 같은 제한이면 넣은 칸에도 (Excel처럼 범위가 늘어남)
  if (!del && at > 1) {
    for (const [k, v] of Object.entries(dv)) {
      const p = parseCell(k);
      if (!p || (axis === 'row' ? p.row : p.col) !== at - 1) continue;
      const after = axis === 'row' ? addr(p.col, at + count) : addr(at + count, p.row);
      if (dv[after] !== v) continue;
      for (let n = 0; n < count; n++) dv[axis === 'row' ? addr(p.col, at + n) : addr(at + n, p.row)] = v;
    }
  }
  m.dataValidations = dv;
  m.conditionalFormattings = (m.conditionalFormattings ?? []).map((cf) => ({
    ...cf, ref: moveSqref(cf.ref, mv),
    rules: cf.rules.map((rule) => (rule.formulae ? { ...rule, formulae: rule.formulae.map((f) => (typeof f === 'string' ? sheetFormula(f) : f)) } : rule)),
  })).filter((cf) => cf.ref);
  const af = filterRef(m.autoFilter);
  if (af) { const a = parseArea(af); const r = a && moveArea(a, mv); m.autoFilter = r ? areaText(r) : null; }
  m.tables = (m.tables ?? []).filter((t) => !tableFix.some((f) => f.drop && f.m.name === t.name)).map((t) => {
    const out = { ...t };
    for (const key of ['tableRef', 'autoFilterRef']) {
      if (!out[key]) continue;
      const a = parseArea(out[key]); const r = a && moveArea(a, mv);
      if (r) out[key] = areaText(r);
    }
    const fx = tableFix.find((f) => f.m.name === t.name);
    if (fx?.delCols) { out.columns = [...out.columns]; out.columns.splice(fx.delCols[0], fx.delCols[1]); }
    if (fx?.insCols) {
      out.columns = [...out.columns];
      const used = new Set(out.columns.map((c) => c.name));
      const fresh = [];
      for (let k = 1; fresh.length < fx.insCols[1]; k++) { const nm = `열${k}`; if (!used.has(nm)) { used.add(nm); fresh.push({ name: nm, filterButton: true }); } }
      out.columns.splice(fx.insCols[0], 0, ...fresh);
      out._fillHeader = true;
    }
    return out;
  });
  for (const md of m.media ?? []) {
    if (md.type !== 'image' || !md.range?.tl) continue;
    for (const end of ['tl', 'br']) {
      const an = md.range[end];
      if (!an) continue;
      const key = axis === 'row' ? 'nativeRow' : 'nativeCol'; const off = axis === 'row' ? 'nativeRowOff' : 'nativeColOff';
      const n1 = an[key] + 1;
      const nn = mv.idx(n1);
      if (nn === null) { an[key] = lo - 1; an[off] = 0; } else an[key] = nn - 1;
    }
  }
  ws._merges = {};
  ws.model = m;
  for (const t of Object.values(ws.tables)) {
    const tm = t.table ?? t.model;
    if (tm?._fillHeader && tm.headerRow !== false) {
      const a = parseArea(tm.tableRef);
      tm.columns.forEach((c, k) => { const cell = ws.getCell(a.r1, a.c1 + k); if (cell.value == null || cell.value === '') cell.value = c.name; });
      delete tm._fillHeader;
    }
    // 표 안에 넣은 행의 계산 열에는 Excel처럼 그 열의 수식을 채운다
    const ta = tm && parseArea(tm.tableRef);
    if (ta && axis === 'row' && !del) {
      const anchor = ta.r1 + (tm.headerRow === false ? 0 : 1);
      const last = ta.r2 - (tm.totalsRow ? 1 : 0);
      tm.columns.forEach((c, k) => {
        if (!c._calc) return;
        for (let r = Math.max(at, anchor); r <= Math.min(at + count - 1, last); r++) {
          const cell = ws.getCell(r, ta.c1 + k);
          if (cell.value == null || cell.value === '') cell.value = { formula: slideFormula(c._calc.formula, r - anchor, 0) };
        }
      });
    }
  }
  // 새 행·열은 위 행·왼쪽 열의 서식을 따른다 (Excel 기본)
  if (!del && at > 1) {
    if (axis === 'row') {
      const src = ws.getRow(at - 1);
      for (let k = 0; k < count; k++) {
        const dst = ws.getRow(at + k);
        if (src.style && Object.keys(src.style).length) dst.style = { ...src.style };
        src.eachCell({ includeEmpty: true }, (cell, c) => { if (cell.style && Object.keys(cell.style).length) dst.getCell(c).style = cell.style; });
      }
    } else {
      const srcCol = at - 1;
      for (const row of ws._rows) {
        if (!row) continue;
        const cell = row.findCell(srcCol);
        if (!cell?.style || !Object.keys(cell.style).length) continue;
        for (let k = 0; k < count; k++) row.getCell(at + k).style = cell.style;
      }
    }
  }
  // 다른 시트 수식과 이름 정의
  for (const other of st.wb.worksheets) {
    if (other === ws) continue;
    for (const cell of allCells(other)) {
      if (cell.type !== VT.Formula || !cell.value?.formula) continue;
      const nf = shiftFormula(cell.value.formula, other.name, target, mv);
      if (nf !== cell.value.formula) cell.value = { ...cell.value, formula: nf };
    }
    for (const cf of other.conditionalFormattings ?? []) for (const rule of cf.rules ?? []) if (rule.formulae) rule.formulae = rule.formulae.map((f) => (typeof f === 'string' ? shiftFormula(f, other.name, target, mv) : f));
  }
  for (const n of st.names) n.text = shiftFormula(n.text, null, target, mv);
  markFormulaChange(st);
  st.touched.add(ws);
}

// ── 파일 엔진: 작업 하나 ─────────────────────────────────

const DEFAULT_CF_FILL = 'FFC7CE';
const DEFAULT_CF_FONT = '9C0006';

function cfValue(v) {
  if (typeof v === 'number') return String(v);
  if (v.startsWith('=')) return fileFormula(v.slice(1));
  return `"${v.replace(/"/g, '""')}"`;
}

function fileOp(st, op) {
  const { wb } = st;
  switch (op.op) {
    case 'add_sheet': {
      if (wb.worksheets.some((w) => sameName(w.name, op.name))) fail(`같은 이름의 시트가 이미 있습니다: ${op.name}`);
      const order = wb.worksheets;
      const after = has(op, 'after') ? sheetOf(st, op.after) : order[order.length - 1];
      const ws = wb.addWorksheet(op.name);
      const next = order.filter((w) => w !== ws);
      next.splice(next.indexOf(after) + 1, 0, ws);
      setOrder(st, next);
      return { sheet: ws.name, index: next.indexOf(ws) + 1 };
    }
    case 'rename_sheet': {
      const ws = sheetOf(st, op.sheet);
      const from = ws.name;
      if (!sameName(from, op.name) && wb.worksheets.some((w) => sameName(w.name, op.name))) fail(`같은 이름의 시트가 이미 있습니다: ${op.name}`);
      if (from !== op.name) {
        rewriteAllFormulas(st, (f) => renameInFormula(f, from, op.name));
        ws.name = op.name;
        st.renamed.set(ws, st.renamed.get(ws) ?? from);
      }
      return { from, sheet: ws.name };
    }
    case 'delete_sheet': {
      const ws = sheetOf(st, op.sheet);
      const visible = wb.worksheets.filter((w) => (w.state ?? 'visible') === 'visible');
      if ((ws.state ?? 'visible') === 'visible' && visible.length <= 1) fail(`보이는 시트가 하나뿐이라 지울 수 없습니다: ${ws.name}`);
      const name = ws.name;
      const gone = tablesOf(ws).map((x) => x.m.name);
      st.names = st.names.filter((n) => n.ws !== ws);
      wb.removeWorksheet(ws.id);
      rewriteAllFormulas(st, (f) => dropTablesInFormula(dropSheetInFormula(f, name), gone));
      markFormulaChange(st);
      setOrder(st, wb.worksheets);
      return { sheet: name };
    }
    case 'move_sheet': {
      const ws = sheetOf(st, op.sheet);
      const order = wb.worksheets;
      if (op.position > order.length) fail(`position이 범위를 벗어났습니다: ${op.position} (1~${order.length})`);
      const next = order.filter((w) => w !== ws);
      next.splice(op.position - 1, 0, ws);
      setOrder(st, next);
      return { sheet: ws.name, position: op.position };
    }
    case 'set_cells': {
      const ws = sheetOf(st, op.sheet);
      for (const [a, v] of Object.entries(op.cells)) {
        const p = parseCell(a);
        checkWritable(ws, { c1: p.col, r1: p.row });
        putValue(st, ws.getCell(a), v);
      }
      st.touched.add(ws);
      return { sheet: ws.name, count: Object.keys(op.cells).length };
    }
    case 'set_range': {
      const ws = sheetOf(st, op.sheet);
      const s = parseCell(op.range);
      op.values.forEach((row, i) => row.forEach((v, j) => checkWritable(ws, { c1: s.col + j, r1: s.row + i })));
      const nc = Math.max(1, ...op.values.map((r) => r.length));
      op.values.forEach((row, i) => {
        for (let j = 0; j < nc; j++) putValue(st, ws.getCell(s.row + i, s.col + j), j < row.length ? row[j] : null);
      });
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'clear': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      for (const mg of mergesOf(ws)) {
        if (intersects(a, mg) && !contains(a, mg)) fail(`병합된 영역 ${areaText(mg)}의 일부만 지울 수 없습니다 — 범위에 병합 영역 전체를 넣으세요`);
      }
      if (op.what !== 'contents') {
        // 범위 안의 병합은 풀고(Excel과 같음), 범위 안에만 있는 조건부 서식은 지운다
        for (const mg of mergesOf(ws)) if (contains(a, mg)) ws.unMergeCells(areaText(mg));
        ws.conditionalFormattings = (ws.conditionalFormattings ?? []).filter((cf) => {
          const parts = String(cf.ref).split(/\s+/).map(parseArea).filter(Boolean);
          return !(parts.length && parts.every((p) => contains(a, p)));
        });
      }
      for (const cell of cellsOf(ws, a, { create: false })) {
        if (op.what !== 'formats') { if (!(cell.isMerged && cell.master !== cell)) cell.value = null; }
        if (op.what !== 'contents') { cell.style = {}; if (op.what === 'all' && cell._comment) cell._comment = undefined; }
      }
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'insert_rows': case 'delete_rows': {
      const ws = sheetOf(st, op.sheet);
      if (op.op === 'insert_rows') {
        const b = usedBounds(ws);
        if (b && b.r2 >= op.at && b.r2 + op.count > MAX_ROW) fail('시트 끝에 있는 내용이 밀려나므로 행을 넣을 수 없습니다');
      }
      spliceSheet(st, ws, 'row', op.at, op.count, op.op === 'delete_rows');
      return { sheet: ws.name };
    }
    case 'insert_columns': case 'delete_columns': {
      const ws = sheetOf(st, op.sheet);
      if (op.op === 'insert_columns') {
        const b = usedBounds(ws);
        if (b && b.c2 >= op.at && b.c2 + op.count > MAX_COL) fail('시트 끝에 있는 내용이 밀려나므로 열을 넣을 수 없습니다');
      }
      spliceSheet(st, ws, 'col', op.at, op.count, op.op === 'delete_columns');
      return { sheet: ws.name };
    }
    case 'format': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      const plain = { ...op }; delete plain.border;
      const touchesStyle = Object.keys(plain).some((k) => !['op', 'sheet', 'range'].includes(k));
      if (touchesStyle) {
        if (a.cols) for (let c = a.c1; c <= a.c2; c++) { const col = ws.getColumn(c); col.style = patchStyle(col.style, op); }
        if (a.rows) for (let r = a.r1; r <= a.r2; r++) { const row = ws.getRow(r); row.style = patchStyle(row.style, op); }
        for (const cell of cellsOf(ws, a)) cell.style = patchStyle(cell.style, op);
      }
      if (op.border) {
        for (const cell of cellsOf(ws, a)) {
          if (op.border === 'none') { const s = { ...(cell.style ?? {}) }; s.border = {}; cell.style = s; continue; }
          const w = op.border === 'outline' ? 'thin' : op.border;
          const p = parseCell(cell.address);
          const sides = {};
          if (op.border !== 'outline' || p.row === a.r1) sides.top = { style: w };
          if (op.border !== 'outline' || p.row === a.r2) sides.bottom = { style: w };
          if (op.border !== 'outline' || p.col === a.c1) sides.left = { style: w };
          if (op.border !== 'outline' || p.col === a.c2) sides.right = { style: w };
          if (Object.keys(sides).length) cell.style = borderStyle(cell.style, sides);
        }
      }
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'column_width': {
      const ws = sheetOf(st, op.sheet);
      const [f, t] = op.columns.split(':').map(colNum);
      for (let c = f; c <= t; c++) {
        const col = ws.getColumn(c);
        if (op.width !== 'auto') { col.width = op.width; continue; }
        let w = 0;
        col.eachCell({ includeEmpty: false }, (cell) => {
          if (cell.isMerged) return;
          const size = cell.font?.size ?? 11;
          w = Math.max(w, textWidth(displayText(cell)) * (size / 11) * (cell.font?.bold ? 1.1 : 1));
        });
        col.width = w ? Math.min(255, Math.round((w + 2) * 10) / 10) : 8.43;
      }
      st.touched.add(ws);
      return { sheet: ws.name, approx: op.width === 'auto' };
    }
    case 'row_height': {
      const ws = sheetOf(st, op.sheet);
      const [f, t] = op.rows.split(':').map(Number);
      for (let r = f; r <= t; r++) {
        const row = ws.getRow(r);
        if (op.height !== 'auto') { row.height = op.height; continue; }
        let h = 0;
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (cell.isMerged) return;
          const size = cell.font?.size ?? 11;
          const text = displayText(cell);
          let lines = text.split('\n').length;
          if (cell.alignment?.wrapText) {
            const width = ws.getColumn(c).width ?? 8.43;
            lines = text.split('\n').reduce((n, ln) => n + Math.max(1, Math.ceil(textWidth(ln) * (size / 11) / Math.max(1, width))), 0);
          }
          h = Math.max(h, lines * size * 1.36 + 1.5);
        });
        row.height = Math.max(15, Math.round(h * 4) / 4) || 15;
      }
      st.touched.add(ws);
      return { sheet: ws.name, approx: op.height === 'auto' };
    }
    case 'merge': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      const t = tableOverlap(ws, a);
      if (t) fail(`표 '${t}'와 겹치는 범위는 병합할 수 없습니다`);
      for (const mg of mergesOf(ws)) {
        if (!intersects(a, mg)) continue;
        if (!contains(a, mg)) fail(`이미 병합된 영역 ${areaText(mg)}가 범위 밖으로 걸쳐 있습니다 — 범위를 넓히거나 먼저 unmerge 하세요`);
        ws.unMergeCells(areaText(mg));
      }
      let discarded = 0;
      for (const cell of cellsOf(ws, a, { create: false })) {
        const p = parseCell(cell.address);
        if ((p.row !== a.r1 || p.col !== a.c1) && cell.value != null && cell.value !== '') { discarded++; cell.value = null; }
      }
      ws.mergeCells(op.range);
      st.touched.add(ws);
      return { sheet: ws.name, discarded };
    }
    case 'unmerge': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      for (const mg of mergesOf(ws)) if (intersects(a, mg)) ws.unMergeCells(areaText(mg));
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'freeze': {
      const ws = sheetOf(st, op.sheet);
      const base = { ...(ws.views?.[0] ?? {}) };
      delete base.xSplit; delete base.ySplit; delete base.topLeftCell; delete base.activePane;
      const c = op.cell ? parseCell(op.cell) : null;
      if (c && (c.col > 1 || c.row > 1)) ws.views = [{ ...base, state: 'frozen', xSplit: c.col - 1, ySplit: c.row - 1, topLeftCell: op.cell }];
      else ws.views = [{ ...base, state: 'normal' }];
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'autofilter': {
      const ws = sheetOf(st, op.sheet);
      if (!op.range) { ws.autoFilter = null; st.filterSet.add(ws); syncFilterName(st, ws, null); return { sheet: ws.name }; }
      const a = parseArea(op.range);
      const t = tableOverlap(ws, a);
      if (t) fail(`표 '${t}'와 겹칩니다 — 표에는 자체 필터가 있습니다`);
      let any = false;
      for (const cell of cellsOf(ws, a, { create: false })) if (cell.value != null && cell.value !== '') { any = true; break; }
      if (!any) fail(`비어 있는 범위에는 자동 필터를 걸 수 없습니다: ${op.range}`);
      ws.autoFilter = areaText(a.cols || a.rows ? { ...a, ...usedClip(ws, a) } : a);
      st.filterSet.add(ws);
      syncFilterName(st, ws, parseArea(ws.autoFilter));
      return { sheet: ws.name, range: ws.autoFilter };
    }
    case 'add_table': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      const t = tableOverlap(ws, a);
      if (t) fail(`표 '${t}'와 겹칩니다`);
      if (mergesOf(ws).some((mg) => intersects(a, mg))) fail(`병합된 칸이 있는 범위에는 표를 만들 수 없습니다: ${op.range}`);
      const style = op.style ?? 'TableStyleMedium2';
      if (!TABLE_STYLE.test(style)) fail(`표 스타일 이름이 올바르지 않습니다: ${style} (예: TableStyleMedium2, TableStyleLight9, TableStyleDark1)`);
      const taken = new Set([...wb.worksheets.flatMap((w) => tablesOf(w).map((x) => x.m.name.toLowerCase())), ...st.names.map((n) => n.name.toLowerCase())]);
      let name = op.name;
      if (name && taken.has(name.toLowerCase())) fail(`이미 있는 표·이름입니다: ${name}`);
      if (!name) { for (let k = 1; ; k++) { if (!taken.has(`표${k}`)) { name = `표${k}`; break; } } }
      // 머리글: 빈 칸·중복은 Excel처럼 '열N'으로 채운다
      const used = new Set();
      const columns = [];
      for (let c = a.c1; c <= a.c2; c++) {
        const cell = ws.getCell(a.r1, c);
        // 머리글은 글자여야 한다: 수식은 Excel처럼 저장된 결과를 글자로 남기고, 결과가 없으면 거절
        if (cell.type === VT.Formula && (cell.value?.result == null || typeof cell.value.result === 'object')) {
          fail(`표 머리글 칸(${cell.address})에 수식이 있습니다 — 값으로 바꾼 뒤 만드세요`);
        }
        let h = displayText(cell).trim() || (typeof cell.value === 'string' ? cell.value : '');
        if (!h || used.has(h.toLowerCase())) {
          const baseName = h || '열';
          for (let k = h ? 2 : 1; ; k++) { const cand = `${baseName}${k}`; if (!used.has(cand.toLowerCase())) { h = cand; break; } }
        }
        used.add(h.toLowerCase());
        if (cell.value !== h) cell.value = h;
        columns.push({ name: h, filterButton: true });
      }
      if (ws.autoFilter && (() => { const f = parseArea(filterRef(ws.autoFilter)); return f && intersects(f, a); })()) { ws.autoFilter = null; st.filterSet.add(ws); }
      const ref = areaText(a);
      const model = {
        name, displayName: name, tableRef: ref, autoFilterRef: ref, headerRow: true, totalsRow: false,
        style: { theme: style, showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
        columns,
      };
      ws.tables[name] = newTable(ws, model);
      st.touched.add(ws);
      return { sheet: ws.name, name, range: ref, style };
    }
    case 'conditional_format': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      const ref = areaText(explicit(a));
      const r = op.rule;
      const tl = addr(parseArea(ref).c1, parseArea(ref).r1);
      const style = {};
      const fillHex = has(r, 'fill') ? r.fill : has(r, 'font_color') ? null : DEFAULT_CF_FILL;
      const fontHex = has(r, 'font_color') ? r.font_color : has(r, 'fill') ? null : DEFAULT_CF_FONT;
      if (fillHex) style.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: `FF${fillHex}` } };
      if (fontHex) style.font = { color: { argb: `FF${fontHex}` } };
      const priority = (ws.conditionalFormattings ?? []).reduce((n, cf) => Math.max(n, ...cf.rules.map((x) => x.priority || 0)), 0) + 1;
      let rule; let note;
      switch (r.type) {
        case 'greater_than': rule = { type: 'cellIs', operator: 'greaterThan', formulae: [cfValue(r.value)], style }; break;
        case 'less_than': rule = { type: 'cellIs', operator: 'lessThan', formulae: [cfValue(r.value)], style }; break;
        case 'equal': rule = { type: 'cellIs', operator: 'equal', formulae: [cfValue(r.value)], style }; break;
        case 'between': rule = { type: 'cellIs', operator: 'between', formulae: [cfValue(r.value), cfValue(r.value2)], style }; break;
        case 'text_contains':
          rule = { type: 'expression', formulae: [`NOT(ISERROR(SEARCH("${r.value.replace(/"/g, '""')}",${tl})))`], style };
          note = '수식 규칙으로 저장';
          break;
        case 'duplicate':
          rule = { type: 'expression', formulae: [`COUNTIF(${absText(parseArea(ref))},${tl})>1`], style };
          note = '수식 규칙(COUNTIF)으로 저장';
          break;
        case 'color_scale':
          rule = { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
            color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }] };
          break;
        case 'data_bar':
          rule = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: `FF${r.fill ?? '638EC6'}` }, gradient: true };
          break;
        default: fail(`알 수 없는 조건부 서식: ${r.type}`);
      }
      rule.priority = priority;
      ws.addConditionalFormatting({ ref, rules: [rule] });
      st.touched.add(ws);
      return { sheet: ws.name, note };
    }
    case 'data_validation': {
      const ws = sheetOf(st, op.sheet);
      const a = parseArea(op.range);
      const target = explicit(a);
      const model = ws.dataValidations.model;
      // 겹치는 기존 제한은 범위 밖 부분만 남긴다
      for (const key of Object.keys(model)) {
        const k = parseArea(key);
        if (!k || !model[key] || !intersects(k, target)) continue;
        const v = model[key];
        delete model[key];
        if (contains(target, k)) continue;
        if (cellsIn(k) > 200_000) fail(`기존 입력 제한 ${key}가 너무 넓어 나눌 수 없습니다 — Excel로 하세요`);
        for (let r = k.r1; r <= k.r2; r++) for (let c = k.c1; c <= k.c2; c++) if (!(r >= target.r1 && r <= target.r2 && c >= target.c1 && c <= target.c2)) model[addr(c, r)] = v;
      }
      let dv;
      if (op.list) {
        const f = typeof op.list === 'string' ? fileFormula(op.list.slice(1)) : `"${op.list.join(',')}"`;
        dv = { type: 'list', allowBlank: true, formulae: [f], showErrorMessage: true };
      } else {
        const both = has(op, 'min') && has(op, 'max');
        dv = { type: op.type, allowBlank: true, showErrorMessage: true,
          operator: both ? 'between' : has(op, 'min') ? 'greaterThanOrEqual' : 'lessThanOrEqual',
          formulae: both ? [op.min, op.max] : [has(op, 'min') ? op.min : op.max] };
      }
      model[areaText(target)] = dv;
      st.touched.add(ws);
      return { sheet: ws.name };
    }
    case 'named_range': {
      const ws = sheetOf(st, op.sheet);
      const text = `${sheetPrefix(ws.name)}${op.area}`;
      const existing = st.names.find((n) => n.ws === null && sameName(n.name, op.name));
      if (tablesOf(ws).some((x) => sameName(x.m.name, op.name)) || wb.worksheets.some((w) => tablesOf(w).some((x) => sameName(x.m.name, op.name)))) fail(`표 이름과 같은 이름은 쓸 수 없습니다: ${op.name}`);
      if (existing) existing.text = text;
      else st.names.push({ name: op.name, ws: null, attrs: '', text });
      return { refersTo: `=${text}` };
    }
    default:
      fail(`파일 엔진으로는 할 수 없는 작업입니다: ${op.op}`);
  }
  return {};
}

// 전체 행·열 범위를 사용 범위로 줄이기 (자동 필터)
function usedClip(ws, a) {
  const b = usedBounds(ws) ?? { c1: a.c1, r1: a.r1, c2: a.c1, r2: a.r1 };
  const out = { c1: a.c1, c2: a.c2, r1: a.r1, r2: a.r2, cols: false, rows: false };
  if (a.cols) { out.r1 = Math.max(1, b.r1); out.r2 = Math.max(b.r2, out.r1); }
  if (a.rows) { out.c1 = Math.max(1, b.c1); out.c2 = Math.max(b.c2, out.c1); }
  return out;
}
// 'A:A' → 'A1:A1048576' (파일 안에서는 전체 열·행도 셀 범위로 적는다)
const explicit = (a) => ({ c1: a.c1, r1: a.r1, c2: a.c2, r2: a.r2 });

// exceljs Table 클래스 (불러온 표처럼 모델만으로 만든다: 셀 값을 다시 쓰지 않게)
let TableClass;
function newTable(ws, model) {
  TableClass ??= new ExcelJS.Workbook().addWorksheet('x').addTable({ name: 'x', ref: 'A1', columns: [{ name: 'a' }], rows: [] }).constructor;
  const t = new TableClass();
  t.worksheet = ws;
  t.model = model;
  return t;
}

// Excel이 시트마다 두는 숨은 이름 _xlnm._FilterDatabase를 자동 필터 범위에 맞춘다
function syncFilterName(st, ws, area) {
  const i = st.names.findIndex((n) => n.ws === ws && n.name === '_xlnm._FilterDatabase');
  if (i < 0) return;
  if (area) st.names[i].text = `${sheetPrefix(ws.name)}${absText(area)}`;
  else st.names.splice(i, 1);
}

// 표 머리글과 열 이름 맞추기 (다르면 Excel이 파일을 고친다)
function syncTableHeaders(st) {
  for (const ws of st.wb.worksheets) {
    for (const { m, area } of tablesOf(ws)) {
      if (!area || m.headerRow === false) continue;
      const used = new Set();
      m.columns.forEach((col, k) => {
        const cell = ws.getCell(area.r1, area.c1 + k);
        const shown = cell.type === VT.Formula ? '' : displayText(cell).trim();
        let h = shown;
        if (!h || used.has(h.toLowerCase())) {
          const base = h || '열';
          for (let n = h ? 2 : 1; ; n++) { const cand = `${base}${n}`; if (!used.has(cand.toLowerCase())) { h = cand; break; } }
        }
        used.add(h.toLowerCase());
        if (shown !== h) cell.value = h;
        col.name = h;
      });
    }
  }
}

// ── 파일 엔진: 저장 ──────────────────────────────────────

function namesXml(st) {
  const order = st.wb.worksheets;
  const items = [];
  for (const n of st.names) {
    let local = '';
    if (n.ws) { const i = order.indexOf(n.ws); if (i < 0) continue; local = ` localSheetId="${i}"`; }
    items.push(`<definedName name="${encodeXml(n.name)}"${local}${n.attrs}>${encodeXml(n.text)}</definedName>`);
  }
  return items.length ? `<definedNames>${items.join('')}</definedNames>` : '';
}

async function sheetParts(zip) {
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const rels = await relsOf(zip, 'xl/workbook.xml');
  const byId = new Map(rels.map((r) => [r.id, r]));
  return [...wbXml.matchAll(/<(?:\w+:)?sheet\b[^>]*\/?>/g)].map((m) => {
    const id = /\sr:id="([^"]*)"|\s\w+:id="([^"]*)"/.exec(m[0]);
    return { name: attrOf(m[0], 'name') ?? '', part: byId.get(id?.[1] ?? id?.[2])?.part };
  });
}

async function postProcess(st, buf) {
  const zip = await JSZip.loadAsync(buf);
  let wbXml = await zip.file('xl/workbook.xml').async('string');
  const dn = namesXml(st);
  if (/<definedNames>[\s\S]*?<\/definedNames>|<definedNames\/>/.test(wbXml)) wbXml = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>|<definedNames\/>/, dn);
  else if (dn) wbXml = wbXml.replace('</sheets>', `</sheets>${dn}`);
  zip.file('xl/workbook.xml', wbXml);

  // 걸러 둔 조건(자동 필터 안의 filterColumn 등)은 손대지 않은 시트에서 그대로 살린다
  const outSheets = await sheetParts(zip);
  const origSheets = st.raw.sheets.filter((s) => s.kind === 'worksheet');
  for (const ws of st.wb.worksheets) {
    if (st.filterSet.has(ws) || st.spliced.has(ws)) continue;
    const origName = st.renamed.get(ws) ?? ws.name;
    const orig = origSheets.find((s) => s.name === origName && s.xml);
    const out = outSheets.find((s) => s.name === ws.name);
    if (!orig || !out?.part || !zip.file(out.part)) continue;
    const full = /<(?:\w+:)?autoFilter\b[^>]*[^/]>[\s\S]*?<\/(?:\w+:)?autoFilter>/.exec(orig.xml)?.[0];
    if (!full) continue;
    let xml = await zip.file(out.part).async('string');
    const leaf = /<autoFilter\b[^>]*\/>/.exec(xml);
    if (leaf && attrOf(leaf[0], 'ref') === attrOf(full, 'ref')) {
      const ns = /^<(\w+):/.exec(full)?.[1];
      const plain = ns ? full.replace(new RegExp(`<(/?)${ns}:`, 'g'), '<$1') : full;
      xml = xml.replace(leaf[0], plain);
      zip.file(out.part, xml);
    }
  }
  // 문서 안 링크: exceljs는 location 링크에도 바깥 링크(r:id)를 붙인다 → location만 남긴다
  if (st.raw.internalLinks.size) {
    for (const sp of outSheets) {
      if (!sp.part || !zip.file(sp.part)) continue;
      let xml = await zip.file(sp.part).async('string');
      if (!/<hyperlink\b/.test(xml)) continue;
      const rp = relsPath(sp.part);
      let relsXml = zip.file(rp) ? await zip.file(rp).async('string') : '';
      const rels = await relsOf(zip, sp.part);
      let changed = false;
      xml = xml.replace(/<hyperlink\b[^>]*\/>/g, (tag) => {
        const rid = attrOf(tag, 'r:id');
        const rel = rid && rels.find((r) => r.id === rid);
        const target = rel?.target;
        const loc = attrOf(tag, 'location') ?? (target && st.raw.internalLinks.has(target) ? target : null);
        if (!rid || !loc || !(st.raw.internalLinks.has(loc) || target === loc)) return tag;
        changed = true;
        relsXml = relsXml.replace(new RegExp(`<Relationship\\b[^>]*\\sId="${rid}"[^>]*/>`), '');
        const rest = tag.replace(/\s(r:id|location)="[^"]*"/g, '').replace(/\s*\/>$/, '');
        return `${rest} location="${encodeXml(loc)}"/>`;
      });
      if (changed) { zip.file(sp.part, xml); if (relsXml) zip.file(rp, relsXml); }
    }
  }
  // 표: 계산 열 수식 되살리기, 원래 자동 필터가 없던 표(필터 단추 끔)는 autoFilter 빼기, 열 수가 범위 너비와 같은지 확인
  const tableModels = new Map();
  for (const w of st.wb.worksheets) for (const { m } of tablesOf(w)) tableModels.set(String(m.name).toLowerCase(), m);
  for (const part of Object.keys(zip.files).filter((x) => /^xl\/tables\/[^/]+\.xml$/.test(x))) {
    let xml = await zip.file(part).async('string');
    const tag = /<table\b[^>]*>/.exec(xml)?.[0] ?? '';
    const name = attrOf(tag, 'name') ?? '';
    const m = tableModels.get(name.toLowerCase());
    if (m) {
      if (m._noAutoFilter) xml = xml.replace(/<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/, '');
      let k = 0;
      xml = xml.replace(/<tableColumn\b([^>]*?)\s*\/>/g, (whole, attrs) => {
        const calc = m.columns?.[k++]?._calc;
        if (!calc || /#REF!/.test(calc.formula)) return whole;
        return `<tableColumn${attrs}><calculatedColumnFormula${calc.attrs}>${encodeXml(calc.formula)}</calculatedColumnFormula></tableColumn>`;
      });
      zip.file(part, xml);
    }
    const ref = parseArea(attrOf(tag, 'ref') ?? '');
    const count = (xml.match(/<tableColumn\b/g) ?? []).length;
    if (!ref || count !== ref.c2 - ref.c1 + 1) throw new Error(`표 '${name}'의 열 수(${count})가 범위 ${attrOf(tag, 'ref')}와 맞지 않습니다`);
  }
  // 기본 글꼴
  if (st.raw.font0 && zip.file('xl/styles.xml')) {
    const sx = await zip.file('xl/styles.xml').async('string');
    const m = /(<fonts\b[^>]*>\s*)(<font\b[\s\S]*?<\/font>|<font\/>)/.exec(sx);
    if (m) {
      const ns = /^<(\w+):/.exec(st.raw.font0)?.[1];
      const f0 = ns ? st.raw.font0.replace(new RegExp(`<(/?)${ns}:`, 'g'), '<$1') : st.raw.font0;
      zip.file('xl/styles.xml', sx.replace(m[0], m[1] + f0));
    }
  }
  // 사용자 지정 문서 속성(민감도 레이블 등)
  const custom = st.zip.file('docProps/custom.xml');
  if (custom && !zip.file('docProps/custom.xml')) {
    zip.file('docProps/custom.xml', await custom.async('string'));
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!ct.includes('/docProps/custom.xml')) ct = ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
    zip.file('[Content_Types].xml', ct);
    let rels = await zip.file('_rels/.rels').async('string');
    if (!rels.includes('docProps/custom.xml')) rels = rels.replace('</Relationships>', '<Relationship Id="rIdOkCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>');
    zip.file('_rels/.rels', rels);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function atomicWrite(output, buf) {
  const tmp = join(dirname(output), `.~${basename(output)}.${process.pid}.${Date.now()}.tmp`);
  try {
    writeFileSync(tmp, buf);
    renameSync(tmp, output);
  } catch (e) {
    rmSync(tmp, { force: true });
    if (e.code === 'EBUSY' || e.code === 'EPERM' || e.code === 'EACCES') throw new Error(`파일이 다른 프로그램에서 열려 있거나 쓰기 권한이 없어 저장하지 못했습니다: ${output}`);
    throw e;
  }
}

async function fileEdit(path, ops, output) {
  const { wb, zip, raw } = await loadBook(path);
  if (raw.lost.length) {
    throw new Error(`이 파일에는 Excel 없이 고치면 사라지는 내용이 있어 파일 엔진으로는 고치지 않습니다: ${raw.lost.join(', ')} `
      + `— Excel이 설치된 PC에서 engine:'app'으로 하세요`);
  }
  fixAfterLoad(wb, raw);
  unshareFormulas(wb);
  const order0 = wb.worksheets;
  const wsRaw = raw.sheets.filter((s) => s.kind === 'worksheet');
  const st = {
    wb, zip, raw,
    names: raw.names.map((n) => ({ ...n, ws: n.local === null ? null : order0[wsRaw.indexOf(raw.sheets[n.local])] ?? null }))
      .filter((n) => n.local === null || n.ws),
    renamed: new Map(), touched: new Set(), filterSet: new Set(), spliced: new Set(), recalc: false,
  };
  const activeBefore = order0[activeIndex(wb)] ?? order0[0];
  const details = [];
  for (const [i, op] of ops.entries()) {
    try {
      const d = fileOp(st, op);
      if (['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns'].includes(op.op)) st.spliced.add(sheetOf(st, d.sheet));
      if (VALUE_OPS.has(op.op) && !(op.op === 'clear' && op.what === 'formats')) st.stale = true;
      details.push(d);
    } catch (e) {
      if (e instanceof OpError) throw opError(i + 1, op.op, e.message);
      throw opError(i + 1, op.op, `처리하지 못했습니다: ${e.message}`);
    }
  }
  syncTableHeaders(st);
  // 파일 엔진은 계산하지 않으므로 값이 바뀌었으면 수식 칸에 저장된 옛 결과를 비운다 (읽는 쪽이 틀린 숫자를 보지 않게.
  // Excel은 fullCalcOnLoad로 열 때 다시 계산한다)
  if (st.stale) {
    for (const ws of wb.worksheets) {
      for (const cell of allCells(ws)) {
        if (cell.type !== VT.Formula) continue;
        const v = cell.value;
        if (v && typeof v === 'object' && v.result !== undefined) { const { result, ...rest } = v; cell.value = rest; st.cleared = true; }
      }
    }
  }
  // 활성 시트 유지 (지워졌으면 처음 보이는 시트)
  const order = wb.worksheets;
  const active = order.includes(activeBefore) ? activeBefore : order.find((w) => (w.state ?? 'visible') === 'visible') ?? order[0];
  wb.views = wb.views?.length ? [...wb.views] : [{ x: 0, y: 0, width: 12000, height: 24000, visibility: 'visible' }];
  wb.views[0] = { ...wb.views[0], activeTab: order.indexOf(active), firstSheet: 0 };
  for (const w of order) if (w.views?.[0] && !!w.views[0].tabSelected !== (w === active)) w.views[0] = { ...w.views[0], tabSelected: w === active };
  wb.calcProperties = { ...(wb.calcProperties ?? {}), fullCalcOnLoad: true };
  let buf;
  try { buf = await postProcess(st, await wb.xlsx.writeBuffer()); } catch (e) {
    throw new Error(`저장할 파일을 만들지 못했습니다: ${e.message} — 저장하지 않았습니다`);
  }
  // 다시 읽히는지 확인한 뒤에만 쓴다
  try { await new ExcelJS.Workbook().xlsx.load(await forExceljs(buf)); } catch (e) {
    throw new Error(`만든 파일을 다시 읽지 못해 저장하지 않았습니다: ${e.message}`);
  }
  atomicWrite(output, buf);
  return { details, recalc: st.recalc || !!st.cleared, cleared: !!st.cleared };
}

// ── csv (파일 엔진: 값만) ─────────────────────────────────

function readCsvFile(path) {
  const { text, encoding } = decodeCsv(readFileSync(path));
  const first = text.slice(0, text.search(/\r?\n|$/));
  const sep = first.includes(',') ? ',' : first.includes('\t') ? '\t' : first.includes(';') ? ';' : ',';
  const eol = /\r\n/.test(text) ? '\r\n' : '\n';
  const trailing = /\r?\n$/.test(text);
  return { rows: parseCsv(text, sep), sep, eol, trailing, encoding };
}

function csvCell(v, sep) {
  const s = String(v ?? '');
  return s.includes(sep) || /["\r\n]/.test(s) || /^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function csvEdit(path, ops, output) {
  const f = readCsvFile(path);
  let rows = f.rows.map((r) => [...r]);
  const set = (r, c, v) => { while (rows.length < r) rows.push([]); const row = rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; };
  const asText = (v, where) => {
    if (v === null) return '';
    if (typeof v === 'string' && v.startsWith('=')) fail(`csv 파일에는 수식을 저장할 수 없습니다(${where}) — Excel로 하면 계산 결과가 저장됩니다`);
    if (typeof v === 'string') return v.startsWith("'") ? v.slice(1) : v;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  };
  const details = [];
  for (const [i, op] of ops.entries()) {
    try {
      if (op.sheet !== undefined && op.sheet !== 1 && op.sheet !== '1' && !sameName(op.sheet, basename(path, extname(path)))) fail(`시트 없음: ${op.sheet} (csv에는 시트가 하나뿐입니다)`);
      switch (op.op) {
        case 'set_cells':
          for (const [a, v] of Object.entries(op.cells)) { const p = parseCell(a); set(p.row, p.col, asText(v, a)); }
          break;
        case 'set_range': {
          const s = parseCell(op.range);
          const nc = Math.max(1, ...op.values.map((r) => r.length));
          op.values.forEach((row, r) => { for (let c = 0; c < nc; c++) set(s.row + r, s.col + c, asText(c < row.length ? row[c] : null, `values[${r}][${c}]`)); });
          break;
        }
        case 'clear': {
          if (op.what === 'formats') fail('csv에는 서식이 없습니다');
          const a = parseArea(op.range);
          for (let r = a.r1; r <= Math.min(a.r2, rows.length); r++) for (let c = a.c1; c <= Math.min(a.c2, rows[r - 1].length); c++) rows[r - 1][c - 1] = '';
          break;
        }
        case 'insert_rows': rows.splice(op.at - 1, 0, ...Array.from({ length: op.count }, () => [])); break;
        case 'delete_rows': rows.splice(op.at - 1, op.count); break;
        case 'insert_columns': rows = rows.map((r) => { const x = [...r]; if (x.length >= op.at) x.splice(op.at - 1, 0, ...Array(op.count).fill('')); return x; }); break;
        case 'delete_columns': rows = rows.map((r) => { const x = [...r]; x.splice(op.at - 1, op.count); return x; }); break;
        default:
          fail(EXCEL_OPS[op.op].file
            ? 'csv 파일은 값만 저장되므로 이 작업은 할 수 없습니다 (서식·표·시트 등은 xlsx에서)'
            : `Excel이 있어야 하는 작업입니다 ${NEED_EXCEL_TAIL}`);
      }
      details.push({});
    } catch (e) {
      if (e instanceof OpError) throw opError(i + 1, op.op, e.message);
      throw e;
    }
  }
  while (rows.length && rows[rows.length - 1].every((x) => x === '')) rows.pop();
  const nc = rows.reduce((n, r) => Math.max(n, r.length), 0);
  const text = rows.map((r) => { const x = [...r]; while (x.length < nc) x.push(''); return x.map((v) => csvCell(v, f.sep)).join(f.sep); }).join(f.eol) + (f.trailing ? f.eol : '');
  const bom = f.encoding !== 'utf8';
  atomicWrite(output, Buffer.concat([bom ? Buffer.from([0xef, 0xbb, 0xbf]) : Buffer.alloc(0), Buffer.from(text, 'utf8')]));
  return { details, note: f.encoding === 'cp949' ? 'CP949(EUC-KR) csv를 UTF-8(BOM)로 저장했습니다' : null };
}

// ── 프로그램 엔진 (Excel COM) ─────────────────────────────

const comOps = (ops) => ops.map((op) => ({ ...op }));

async function appInspect(path) {
  const res = await comCall({ app: 'excel', op: 'inspect', path }, { timeoutMs: COM_TIMEOUT });
  const i = res.inspect ?? {};
  return {
    engine: 'excel', activeSheet: i.activeSheet ?? null,
    sheets: list(i.sheets).map((s) => ({ ...s, tables: list(s.tables), charts: list(s.charts) })).sort((a, b) => a.index - b.index),
    names: list(i.names), hiddenNames: i.hiddenNames ?? 0, formulas: list(i.formulas), formulaCount: i.formulaCount ?? 0,
  };
}

const BOM = Buffer.from([0xef, 0xbb, 0xbf]);
// csv 인코딩: 'utf8bom' | 'utf8' | 'cp949' 와 글
function decodeCsv(buf) {
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return { text: buf.subarray(3).toString('utf8'), encoding: 'utf8bom' };
  try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(buf), encoding: 'utf8' }; }
  catch { return { text: new TextDecoder('euc-kr').decode(buf), encoding: 'cp949' }; }
}

async function appEdit(path, ops, output, { csvSheets = null } = {}) {
  const ext = extOf(path);
  const outExt = extOf(output);
  const notes = [];
  let work = null;
  const workDir = () => (work ??= mkdtempSync(join(tmpdir(), 'officekit-xl-')));
  try {
    // BOM 없는 csv는 Excel이 ANSI(CP949)로 읽어 글자가 깨진다 → UTF-8(BOM) 사본을 같은 이름(=시트 이름)으로 만들어 연다
    let src = path; let encoding = null; let lf = false;
    if (ext === 'csv') {
      const dec = decodeCsv(readFileSync(path));
      encoding = dec.encoding;
      lf = /\n/.test(dec.text) && !/\r\n/.test(dec.text);
      if (encoding !== 'utf8bom') {
        src = join(workDir(), basename(path));
        writeFileSync(src, Buffer.concat([BOM, Buffer.from(dec.text, 'utf8')]));
      }
    }
    // csv는 Excel이 UTF-8로 저장하도록 다른 이름으로 저장한 뒤 옮긴다 (그대로 저장하면 ANSI csv가 된다)
    const viaTemp = outExt === 'csv';
    const target = viaTemp ? join(workDir(), `out-${Date.now()}.csv`) : output;
    const inPlace = !viaTemp && src === path && resolve(output) === resolve(path);
    const before = inPlace ? statSync(path).mtimeMs : 0;
    let res;
    try {
      res = await comCall({ app: 'excel', op: 'edit', path: src, operations: comOps(ops), save: target, saveFormat: outExt, ...(csvSheets ? { csvSheets } : {}) },
        { timeoutMs: COM_TIMEOUT });
    } catch (e) { throw comEditError(e); }
    if (inPlace && statSync(path).mtimeMs === before) {
      throw new Error(`Excel이 문서를 저장하지 못했습니다 (문서가 다른 프로그램에서 열려 있는지 확인하세요): ${path}`);
    }
    if (viaTemp) {
      let buf = readFileSync(target);
      // BOM 없던 UTF-8 csv는 BOM 없이 (원래 인코딩 유지), CP949는 UTF-8(BOM)으로 바뀜을 알린다
      if (encoding === 'utf8' && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) buf = buf.subarray(3);
      if (encoding === 'cp949') notes.push('CP949(EUC-KR) csv를 UTF-8(BOM)로 저장했습니다');
      // 줄바꿈도 원래대로 (Excel은 CRLF로 쓴다)
      if (lf) buf = Buffer.from(buf.toString('utf8').replace(/\r\n/g, '\n'), 'utf8');
      atomicWrite(output, buf);
    }
    return { details: list(res.results), notes };
  } finally {
    if (work) rmSync(work, { recursive: true, force: true });
  }
}

// ── 공개 함수 ────────────────────────────────────────────

/**
 * 엑셀 문서 살펴보기 → { engine, activeSheet, sheets[], names[], hiddenNames, formulas[], formulaCount }
 * engine: 'auto'(xlsx·xlsm·csv는 파일 엔진, 그 밖은 Excel) | 'app' | 'file'
 */
export async function excelInspect(path, { engine = 'auto' } = {}) {
  path = mustExist(path);
  const ext = excelOnly(path);
  const fileOk = ['xlsx', 'xlsm', 'csv'].includes(ext);
  if (engine === 'file' && !fileOk) throw new Error(await needExcel(`.${ext} 파일은 Excel 없이 열 수 없습니다`));
  if (engine === 'app' || (engine === 'auto' && !fileOk)) {
    if (!(await excelReady(ext))) throw new Error(await needExcel(`.${ext} 파일을 Excel로 살펴볼 수 없습니다`));
    return { ...(await appInspect(path)), path, format: ext };
  }
  const info = ext === 'csv' ? await csvInspect(path) : await fileInspect(path);
  return { ...info, path, format: ext };
}

/**
 * 엑셀 편집: operations를 순서대로 모두 적용한 뒤 한 번만 저장. 하나라도 실패하면 저장하지 않는다.
 * → { engine, results: [작업별 한 줄], output, notes: [] }
 */
export async function excelEdit(path, operations, { output, engine = 'auto' } = {}) {
  path = mustExist(path);
  const ext = excelOnly(path);
  output = resolve(output ?? path);
  const outExt = extOf(output);
  const ops = normalizeOperations(operations);
  const notes = [];

  if (outExt === 'csv') {
    const bad = ops.filter((o) => !CSV_OPS.has(o.op)).map((o) => o.op);
    if (bad.length) throw new Error(`csv 파일은 값만 저장되므로 할 수 없는 작업입니다: ${[...new Set(bad)].join(', ')} — xlsx로 저장하세요 (output을 .xlsx로, Excel 필요)`);
    if (ops.some((o) => ['add_sheet', 'copy_sheet', 'delete_sheet', 'move_sheet', 'rename_sheet'].includes(o.op))) throw new Error('csv에는 시트가 하나뿐입니다');
  }

  // csv → csv의 값 작업은 Excel이 있어도 파일 엔진으로: Excel은 csv를 열 때 '0123'·'1/2' 같은 값을 숫자·날짜로 바꾸고
  // BOM 없는 UTF-8을 ANSI로 읽는다. 파일 엔진은 인코딩·구분자·줄바꿈을 그대로 지킨다
  const csvByFile = engine === 'auto' && ext === 'csv' && outExt === 'csv' && ops.every((o) => CSV_FILE_OPS.has(o.op) && !hasFormulaValue(o));
  const appOk = engine === 'file' || csvByFile ? false : await excelReady(ext);
  let eng;
  if (engine === 'app') {
    if (!appOk) throw new Error(await needExcel('engine:\'app\'으로 고칠 수 없습니다'));
    eng = 'app';
  } else if (appOk) eng = 'app';
  else {
    // 파일 엔진: xlsx(csv) 같은 형식으로만, ×인 작업은 못 함
    const why = engine === 'file' ? '파일 엔진(Excel 없이)으로는' : 'Excel 없이는';
    if (ext === 'xlsm') throw new Error(await needExcel(`${why} .xlsm(매크로 문서)을 고칠 수 없습니다 (매크로가 사라짐)`, ext));
    if (!['xlsx', 'csv'].includes(ext)) throw new Error(await needExcel(`${why} .${ext} 파일을 고칠 수 없습니다`, ext));
    if (outExt !== ext) throw new Error(await needExcel(`${why} 같은 형식(.${ext})으로만 저장할 수 있습니다 (요청: .${outExt})`, ext));
    const need = [...new Set(ops.filter((o) => !EXCEL_OPS[o.op].file).map((o) => o.op))];
    if (need.length) throw new Error(await needExcel(`${why} 할 수 없는 작업입니다: ${need.join(', ')}`, ext));
    eng = 'file';
  }
  if (eng === 'app') {
    if (!APP_SAVE.has(outExt)) throw new Error(`결과를 .${outExt || '(확장자 없음)'} 파일로 저장할 수 없습니다 (가능: ${[...APP_SAVE].join(', ')})`);
    // 여러 시트 문서를 csv로: 작업이 고친 시트를 저장한다 (시트를 안 주면 첫 시트)
    const csvSheets = outExt === 'csv' && ext !== 'csv' ? ops.filter((o) => o.op !== 'recalculate').map((o) => o.sheet ?? 1) : null;
    const { details, notes: appNotes } = await appEdit(path, ops, output, { csvSheets: csvSheets?.length ? csvSheets : null });
    notes.push(...appNotes);
    if (outExt === 'csv' && ext !== 'csv') {
      const sheet = details.find((d) => d?.sheet)?.sheet;
      notes.push(sheet ? `csv에는 시트 '${sheet}'의 값만 저장됩니다` : 'csv에는 활성 시트의 값만 저장됩니다');
    }
    return { engine: 'excel', results: ops.map((op, i) => describe(op, details[i] ?? {})), output, notes };
  }
  if (ext === 'csv') {
    const r = csvEdit(path, ops, output);
    if (r.note) notes.push(r.note);
    return { engine: 'file', results: ops.map((op, i) => describe(op, r.details[i] ?? {})), output, notes };
  }
  const { details, recalc, cleared } = await fileEdit(path, ops, output);
  if (recalc) {
    notes.push(cleared ? '파일 엔진은 계산하지 않아 수식 칸에 저장돼 있던 옛 결과를 비웠습니다 — Excel에서 열 때 다시 계산됩니다'
      : '수식 결과는 Excel에서 열 때 다시 계산됩니다 (파일 엔진은 계산하지 않음)');
  }
  return { engine: 'file', results: ops.map((op, i) => describe(op, details[i] ?? {})), output, notes };
}

// ── 살펴보기 결과를 글로 ──────────────────────────────────

const VIS_KO = { hidden: '숨김', veryHidden: '완전히 숨김' };
const MAX_SHEETS = 100; // 자세히 보여 줄 시트 수
const MAX_OBJS = 20; // 시트마다 보여 줄 표·차트 수
export function inspectText(info) {
  const lines = [];
  const name = info.path ? basename(info.path) : '';
  lines.push(`${name} — Excel (.${info.format ?? ''}, 엔진: ${info.engine === 'excel' ? 'Excel' : '파일'})`);
  lines.push(`시트 ${info.sheets.length}개${info.activeSheet ? ` · 활성 시트: ${q(info.activeSheet)}` : ''}`);
  const shown = info.sheets.filter((s) => !s.brief).slice(0, MAX_SHEETS);
  for (const s of shown) {
    const bits = [];
    const nTables = s.tableCount ?? s.tables?.length ?? 0;
    const nCharts = s.chartCount ?? s.charts?.length ?? 0;
    bits.push(s.usedRange ? `사용 범위 ${s.usedRange} (${s.rows}행×${s.cols}열)` : '비어 있음');
    if (s.freeze) bits.push(`틀 고정 ${s.freeze}`);
    if (s.autoFilter) bits.push(`자동 필터 ${s.autoFilter}`);
    if (s.merges) bits.push(`병합 ${s.merges < 0 ? '알 수 없음' : `${s.merges}개`}`);
    if (s.conditionalFormats) bits.push(`조건부 서식 ${s.conditionalFormats}개`);
    if (s.validations) bits.push('입력 제한 있음');
    if (nCharts) bits.push(`차트 ${nCharts}개`);
    if (nTables) bits.push(`표 ${nTables}개`);
    lines.push(`${s.index}. ${q(s.name)}${VIS_KO[s.visible] ? ` (${VIS_KO[s.visible]})` : ''} — ${bits.join(', ')}`);
    for (const t of (s.tables ?? []).slice(0, MAX_OBJS)) lines.push(`   표 ${t.name}: ${t.range}${t.style ? `, ${t.style}` : ''}${t.totals ? ', 요약 행' : ''}`);
    if (nTables > MAX_OBJS) lines.push(`   … 외 표 ${nTables - MAX_OBJS}개`);
    for (const c of (s.charts ?? []).slice(0, MAX_OBJS)) lines.push(`   차트 ${c.name || '(이름 없음)'}${c.title ? `: '${c.title}'` : ''}`);
    if (nCharts > MAX_OBJS) lines.push(`   … 외 차트 ${nCharts - MAX_OBJS}개`);
  }
  const rest = info.sheets.filter((s) => !shown.includes(s));
  if (rest.length) {
    const names = rest.slice(0, 200).map((s) => `${s.index}. ${q(s.name)}${VIS_KO[s.visible] ? `(${VIS_KO[s.visible]})` : ''}`).join(', ');
    lines.push(`… 외 시트 ${rest.length}개 (처음 ${MAX_SHEETS}개 시트만 자세히${rest.some((s) => s.brief) ? ', 수식 수도 그 시트들만 셈' : ''}): ${names}${rest.length > 200 ? ' …' : ''}`);
  }
  if (info.names?.length) {
    lines.push(`이름 정의 ${info.names.length}개${info.hiddenNames ? ` (숨은 이름 ${info.hiddenNames}개 제외)` : ''}:`);
    for (const n of info.names.slice(0, 50)) lines.push(`   ${n.name}${n.scope ? ` [시트 ${n.scope}]` : ''} = ${String(n.refersTo).replace(/^=/, '')}`);
    if (info.names.length > 50) lines.push(`   … 외 ${info.names.length - 50}개`);
  } else lines.push('이름 정의 없음');
  if (info.formulaCount) {
    lines.push(`수식 ${info.formulaCount}개${info.formulas.length < info.formulaCount ? ` (예시 ${info.formulas.length}개)` : ''}:`);
    for (const f of info.formulas) lines.push(`   ${q(f.sheet)}!${f.cell} ${f.formula}`);
  } else lines.push('수식 없음');
  if (info.fileEngineLimits?.length) {
    lines.push(`주의: Excel 없이(파일 엔진) 고치면 사라질 내용이 있어 이 파일은 Excel로만 고칩니다: ${info.fileEngineLimits.join(', ')}`);
  }
  return lines.join('\n');
}

/** 시험용: 수식 참조 도우미 */
export const _formula = { scanRefs, shiftFormula, slideFormula, renameInFormula, dropSheetInFormula, dropTablesInFormula, mover };
