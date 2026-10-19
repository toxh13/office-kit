// hwp.mjs — 한컴오피스 없이 HWP/HWPX를 읽고 쓰는 래퍼.
// HOP(golbin/hop)이 내부에서 쓰는 rhwp 엔진(@rhwp/core, WASM)을 Node에서 직접 호출한다.
//
// 번호 규칙 (inspect·편집 공통, 모두 1부터):
//   문단 n — 모든 구역의 본문 문단을 문서 순서대로 이어 센 번호 (구역 1의 문단 다음에 구역 2의 문단).
//           표 칸·머리말·꼬리말·각주 안 문단은 세지 않는다. 표가 든 문단도 문단 하나로 센다.
//   표 n   — 본문에 놓인 표를 문서 순서대로 센 번호 (표 안의 표는 제외). 행·열도 1부터, 병합 칸은 왼쪽 위 칸 번호로.
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs';
import { extname } from 'node:path';
import { createRequire } from 'node:module';
import { initSync, HwpDocument, version as engineVersion } from '@rhwp/core';

const require = createRequire(import.meta.url);
initSync({ module: readFileSync(require.resolve('@rhwp/core/rhwp_bg.wasm')) });

export { engineVersion, HwpDocument };

// rhwp는 대부분 JSON 문자열을 돌려주고, 실패하면 {"ok":false,...}나 예외를 낸다.
function call(result, what) {
  if (typeof result !== 'string') return result;
  let parsed;
  try { parsed = JSON.parse(result); } catch { return result; }
  if (parsed && parsed.ok === false) throw new Error(`${what} 실패: ${parsed.error ?? result}`);
  return parsed;
}

// 엔진 예외(문자열·JsValue)도 한국어 작업 이름을 붙여 Error로
function engine(fn, what) {
  try { return call(fn(), what); } catch (e) {
    if (e instanceof Error && e.message.startsWith(`${what} 실패`)) throw e;
    throw new Error(`${what} 실패: ${e?.message ?? e}`);
  }
}

const ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify', distribute: 'distribute' };
const ALIGN_KO = { left: '왼쪽', center: '가운데', right: '오른쪽', justify: '양쪽', distribute: '배분' };
const BASE_CHAR = { bold: false, italic: false, underline: false, strikethrough: false, fontSize: 1000, textColor: '#000000' };

// 단위: HWPUNIT = 1/7200인치. 문단 여백·들여쓰기·문단 간격은 rhwp에 HWPUNIT의 2배 값으로 넘긴다 (1pt = 200)
const HU_PER_MM = 7200 / 25.4;
const HU_PER_PX = 7200 / 96; // 엔진이 돌려주는 문단 간격 값(px, 96dpi) → HWPUNIT
const CELL_PAD_TB = Math.round(HU_PER_MM); // 새 표 칸 안 위·아래 여백 1mm
const PARA_PER_MM = HU_PER_MM * 2;
const PARA_PER_PT = 200;
const mmOf = (hu) => Math.round((hu / HU_PER_MM) * 10) / 10;
const pxToMm = (px) => Math.round(((px * 25.4) / 96) * 10) / 10;

export const PAPERS = { A3: [297, 420], A4: [210, 297], A5: [148, 210], B4: [257, 364], B5: [182, 257], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };

// 구역 첫 문단에 있어도 되는(지우면 안 되는) 구조 개체 — 구역·단 정의, 머리말·꼬리말, 쪽 번호 위치 등
const STRUCTURAL = new Set(['secd', 'cold', 'head', 'foot', 'pgnp', 'pghd', 'nwno', 'pgct', 'bokm', 'atno', 'tcps']);
const CONTROL_KO = { tbl: '표', gso: '그림·도형', fn: '각주', en: '미주', eqed: '수식', '': '누름틀·필드' };

// 사용자 친화 스타일 {bold, italic, underline, size(pt), color, font} → rhwp 글자 속성 (기본값 위에 덮어씀)
function charProps(style = {}) {
  const p = { ...BASE_CHAR };
  for (const k of ['bold', 'italic', 'underline', 'strikethrough']) if (style[k] !== undefined) p[k] = !!style[k];
  if (style.size) p.fontSize = Math.round(style.size * 100);
  if (style.color) p.textColor = style.color;
  return p;
}

function color(c, what = '글자색') {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(c ?? '').trim());
  if (!m) throw new Error(`${what}은 '#rrggbb' 형식이어야 합니다: ${c}`);
  return `#${m[1].toLowerCase()}`;
}

const posInt = (v, what) => {
  if (!Number.isInteger(v) || v < 1) throw new Error(`${what}는 1 이상의 정수여야 합니다: ${v}`);
  return v;
};

/** 그림 바이트에서 형식·픽셀 크기 읽기 (png·jpg·gif·bmp) */
export function imageInfo(buf) {
  const b = Buffer.from(buf);
  if (b.length > 24 && b.readUInt32BE(0) === 0x89504e47) return { ext: 'png', width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
  if (b.length > 10 && b.toString('latin1', 0, 4) === 'GIF8') return { ext: 'gif', width: b.readUInt16LE(6), height: b.readUInt16LE(8) };
  if (b.length > 26 && b.toString('latin1', 0, 2) === 'BM') return { ext: 'bmp', width: Math.abs(b.readInt32LE(18)), height: Math.abs(b.readInt32LE(22)) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      if (marker === 0xff) { i++; continue; }
      const len = b.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
        return { ext: 'jpg', width: b.readUInt16BE(i + 7), height: b.readUInt16BE(i + 5) };
      }
      i += 2 + len;
    }
  }
  throw new Error('그림 형식을 알 수 없습니다 (png·jpg·jpeg·gif·bmp만 넣을 수 있습니다)');
}

export class Hwp {
  constructor(doc) { this.doc = doc; }

  /** 빈 문서 만들기 */
  static create() {
    const doc = HwpDocument.createEmpty();
    call(doc.createBlankDocument(), '빈 문서 생성');
    const h = new Hwp(doc);
    h._fresh = true; // 첫 빈 문단을 재사용하기 위한 표시
    return h;
  }

  /** 파일 열기 (.hwp / .hwpx / .hml). 암호 문서는 { password } */
  static open(path, { password } = {}) {
    const bytes = readFileSync(path);
    return new Hwp(password ? HwpDocument.openWithPassword(bytes, password) : new HwpDocument(bytes));
  }

  /**
   * 화면(그림)용으로 열기. 줄 배치 정보(lineseg)가 빈 hwpx(rhwp가 만든 hwpx 등)는
   * 표 다음 문단이 겹쳐 그려지므로 hwp로 한 번 거쳐 다시 연다. 표 다음 문단은 표의 바깥 아래 여백만큼 띄운다 (_gapAfterTables).
   * 파일은 바꾸지 않는다.
   */
  static openForView(path, opts) {
    let h = Hwp.open(path, opts);
    if (extname(path).toLowerCase() !== '.hwpx') { h._gapAfterTables(); return h; }
    let empty = 0;
    try { empty = JSON.parse(h.doc.getValidationWarnings()).count ?? 0; } catch {}
    if (empty) {
      try {
        const v = new Hwp(new HwpDocument(h.toBytes('hwp')));
        h.close();
        h = v;
      } catch {}
    }
    h._relayoutObjects();
    h._gapAfterTables();
    return h;
  }

  // 표 다음 문단: 한글은 표의 바깥 아래 여백(기본 1mm)만큼 띄우는데 rhwp 0.8.7은 그 여백을 무시하고 표 바로 밑에 붙여 그린다
  // (다음 문단 글자가 표 아래 테두리에 닿거나 겹침). 화면용 사본에서만 다음 문단의 '문단 위 간격'에 그 여백을 더한다.
  // rhwp는 문단 위 간격을 절반으로 그리므로(값 1000 → 6.7px = 500 HWPUNIT) 두 배를 더한다. 파일은 바꾸지 않는다.
  _gapAfterTables() {
    let list = [];
    try { list = this.tables(); } catch { return; }
    if (!list.length) return;
    this._batch(() => {
      for (const t of list) {
        try {
          const tp = JSON.parse(this.doc.getTableProperties(t.section, t.paraIndex, t.controlIndex));
          const gap = Number(tp.outerBottom) || 0;
          if (tp.treatAsChar || tp.textWrap !== 'TopAndBottom' || gap <= 0) continue;
          const next = t.paraIndex + 1;
          if (next >= this.doc.getParagraphCount(t.section)) continue;
          const pp = JSON.parse(this.doc.getParaPropertiesAt(t.section, next, 0));
          const before = Math.round((Number(pp.spacingBefore) || 0) * HU_PER_PX);
          this.doc.applyParaFormat(t.section, next, JSON.stringify({ spacingBefore: before + 2 * gap }));
        } catch {}
      }
    });
  }

  // 글 없이 그림·수식만 든 문단은 줄 정보가 비어 있으면 그려지지 않는다 → 화면용 사본에서 한 글자 넣었다 지워 다시 배치
  _relayoutObjects() {
    const targets = [...this._controlMap()].filter(([, list]) => list.some((c) => c.id === 'gso' || c.id === 'eqed')).map(([g]) => g);
    if (!targets.length) return;
    this._batch(() => {
      for (const g of targets) {
        try {
          const { sec, para } = this.locate(g + 1);
          call(this.doc.insertText(sec, para, 0, ' '), '배치');
          call(this.doc.deleteText(sec, para, 0, 1), '배치');
        } catch {}
      }
    });
  }

  /** 엔진 원본 객체 — 래퍼에 없는 기능(약 400개 메서드)은 여기서 직접 호출 */
  get raw() { return this.doc; }

  info() { return call(this.doc.getDocumentInfo(), '문서 정보'); }
  get pageCount() { return this.doc.pageCount(); }
  get format() { return this.doc.getSourceFormat(); }

  /** 전체 텍스트 (표 셀 포함, 문서 순서대로) */
  text() {
    return call(this.doc.getScanItems(), '텍스트 추출')
      .map((it) => it.text ?? '')
      .join('')
      .replace(/\r\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  _paraText(s, p) {
    const len = this.doc.getParagraphLength(s, p);
    const raw = len ? this.doc.getTextRange(s, p, 0, len) : '';
    try { const j = JSON.parse(raw); return typeof j === 'string' ? j : (j.text ?? raw); } catch { return raw; }
  }

  /** 본문 문단 목록 [{section, index, n, text}] (표 셀 제외). n = 1부터 센 문서 전체 문단 번호 */
  paragraphs() {
    const out = [];
    let n = 0;
    for (let s = 0; s < this.doc.getSectionCount(); s++) {
      for (let p = 0; p < this.doc.getParagraphCount(s); p++) out.push({ section: s, index: p, n: ++n, text: this._paraText(s, p) });
    }
    return out;
  }

  /** 페이지 하나의 텍스트 */
  pageText(page) {
    const t = this.doc.getPageText(page);
    try { return JSON.parse(t); } catch { return t; }
  }

  /** 검색: [{sec, para, charOffset, length, cellContext?}] */
  search(query, { caseSensitive = false, includeCells = true } = {}) {
    return call(this.doc.searchAllText(query, caseSensitive, includeCells), '검색');
  }

  /** 모두 바꾸기 → 바꾼 개수. 영문 대소문자를 구분한다 (Word·Excel·PowerPoint와 같게) */
  replace(find, replacement, { caseSensitive = true } = {}) {
    return call(this.doc.replaceAll(find, replacement, caseSensitive), '바꾸기').count ?? 0;
  }

  // ── 번호 ↔ 위치 ─────────────────────────────────────────

  _counts() {
    const out = [];
    for (let s = 0; s < this.doc.getSectionCount(); s++) out.push(this.doc.getParagraphCount(s));
    return out;
  }

  /** 문서 전체 본문 문단 수 */
  get paragraphCount() { return this._counts().reduce((a, b) => a + b, 0); }

  /** 문단 번호(1부터) → { sec, para } */
  locate(n) {
    const counts = this._counts();
    const total = counts.reduce((a, b) => a + b, 0);
    if (!Number.isInteger(n) || n < 1 || n > total) throw new Error(`문단 번호가 범위를 벗어났습니다: ${n} (전체 ${total}개, 1부터)`);
    let k = n - 1;
    for (let s = 0; s < counts.length; s++) {
      if (k < counts[s]) return { sec: s, para: k };
      k -= counts[s];
    }
    throw new Error(`문단 ${n}을 찾지 못했습니다`);
  }

  /** { sec, para } → 문단 번호(1부터) */
  numberOf(sec, para) {
    const counts = this._counts();
    let n = para + 1;
    for (let s = 0; s < sec; s++) n += counts[s];
    return n;
  }

  // 본문 개체 목록을 문단별로 { 전체 문단 index(0부터): [{id, desc, index}] }
  _controlMap() {
    const map = new Map();
    let list = [];
    try { list = JSON.parse(this.doc.getControls()); } catch {}
    for (const c of list) {
      if (c.list !== 0) continue;
      if (!map.has(c.para)) map.set(c.para, []);
      map.get(c.para).push({ id: String(c.ctrlId ?? '').trim(), desc: c.userDesc || '', index: c.controlIndex });
    }
    return map;
  }

  // 문단 안의 '내용' 개체 (표·그림·누름틀·각주 등). 구조 개체(구역 정의·머리말 등)는 뺀다
  _visibleControls(sec, para, map = this._controlMap()) {
    return (map.get(this.numberOf(sec, para) - 1) ?? []).filter((c) => !STRUCTURAL.has(c.id));
  }

  static _describe(controls) {
    const n = new Map();
    for (const c of controls) { const k = c.id === 'gso' ? (c.desc || '그림·도형') : (CONTROL_KO[c.id] ?? (c.desc || c.id)); n.set(k, (n.get(k) ?? 0) + 1); }
    return [...n].map(([k, v]) => (v > 1 ? `${k} ${v}개` : k)).join(', ');
  }

  // 찾을 글 → 일치 목록 (대소문자 구분). cells=false면 본문 문단만
  _matches(find, { cells = true } = {}) {
    if (typeof find !== 'string' || !find) throw new Error('찾을 글(find)이 비어 있습니다');
    const all = call(this.doc.searchAllText(find, true, true), '검색') ?? [];
    return cells ? all : all.filter((m) => !m.cellContext);
  }

  _firstMatch(find, { cells = true } = {}) {
    const ms = this._matches(find, { cells });
    if (ms.length) return ms[0];
    if (!cells && this._matches(find).length) throw new Error(`'${find}'은(는) 표 안에만 있습니다 — 본문 글을 지정하세요`);
    throw new Error(`'${find}'을(를) 찾지 못했습니다 (대소문자 구분)`);
  }

  /** 글이 든 첫 본문 문단 번호 */
  findParagraph(find) {
    const m = this._firstMatch(find, { cells: false });
    return this.numberOf(m.sec, m.para);
  }

  // ── 스타일 ───────────────────────────────────────────────

  /** 문서 스타일 목록 [{id, name, englishName, type(0 문단, 1 글자), paraShapeId, charShapeId}] */
  styles() { return call(this.doc.getStyleList(), '스타일 목록'); }

  /** 스타일 찾기: 문서 스타일 이름(예: '본문', '개요 1') 또는 영문 이름('Body', 'Outline 1') */
  findStyle(name) {
    const list = this.styles();
    const key = String(name ?? '').trim();
    const norm = (s) => String(s ?? '').replace(/\s+/g, '').toLowerCase();
    const hit = list.find((s) => s.name === key) ?? list.find((s) => s.englishName === key)
      ?? list.find((s) => norm(s.name) === norm(key) || norm(s.englishName) === norm(key));
    if (!hit) throw new Error(`스타일 '${key}'이(가) 문서에 없습니다 (있는 스타일: ${list.map((s) => s.name).join(', ')})`);
    return hit;
  }

  // 새 문단 모양: 스타일(기본 '바탕글') 또는 본보기 문단(like)의 문단 모양·글자 모양
  _look({ style, like } = {}) {
    if (style !== undefined && style !== null && style !== '') {
      const st = this.findStyle(style);
      if (st.type === 1) { const base = this.styles()[0]; return { styleId: base.id, psid: base.paraShapeId, csid: st.charShapeId }; }
      return { styleId: st.id, psid: st.paraShapeId, csid: st.charShapeId };
    }
    if (like !== undefined && like !== null) {
      const { sec, para } = this.locate(like);
      return {
        styleId: call(this.doc.getStyleAt(sec, para), '스타일 조회').id,
        psid: call(this.doc.getParaPropertiesAt(sec, para), '문단 모양 조회').paraShapeId,
        csid: call(this.doc.getCharPropertiesAt(sec, para, 0), '글자 모양 조회').charShapeId,
      };
    }
    const base = this.styles()[0];
    return { styleId: base.id, psid: base.paraShapeId, csid: base.charShapeId };
  }

  _applyLook(sec, para, look) {
    // applyStyle은 줄 배치 정보를 지우므로(hwpx 저장 시 그림 문단이 안 보임) 스타일이 다를 때만 부른다
    let cur;
    try { cur = JSON.parse(this.doc.getStyleAt(sec, para)).id; } catch {}
    if (cur !== look.styleId) engine(() => this.doc.applyStyle(sec, para, look.styleId), '스타일 적용');
    engine(() => this.doc.setParaShapeId(sec, para, look.psid), '문단 모양');
    const len = this.doc.getParagraphLength(sec, para);
    engine(() => this.doc.setCharShapeId(sec, para, 0, Math.max(len, 1), look.csid), '글자 모양');
  }

  // ── 글자·문단 서식 ───────────────────────────────────────

  /**
   * 바꿀 글자 서식만 담은 rhwp 속성.
   * f: { bold, italic, underline, strike, size(pt), color('#rrggbb'), font, highlight('#rrggbb'), superscript, subscript }
   */
  charMods(f = {}) {
    const p = {};
    for (const k of ['bold', 'italic', 'underline', 'superscript', 'subscript']) if (f[k] !== undefined && f[k] !== null) p[k] = !!f[k];
    if (f.strike !== undefined && f.strike !== null) p.strikethrough = !!f.strike;
    if (f.size !== undefined && f.size !== null) {
      if (!(Number(f.size) > 0 && Number(f.size) <= 4096)) throw new Error(`글자 크기(pt)가 올바르지 않습니다: ${f.size}`);
      p.fontSize = Math.round(Number(f.size) * 100);
    }
    if (f.color) p.textColor = color(f.color);
    if (f.highlight) p.shadeColor = color(f.highlight, '형광펜 색');
    if (f.font) {
      const id = this.doc.findOrCreateFontId(String(f.font));
      if (id < 0) throw new Error(`글꼴을 등록하지 못했습니다: ${f.font}`);
      p.fontId = id;
    }
    if (!Object.keys(p).length) throw new Error('바꿀 글자 서식이 없습니다 (bold, italic, underline, strike, size, color, font, highlight 중 하나 이상)');
    return p;
  }

  /** 글자 서식: find와 같은 글(본문·표 칸)에 적용 → 적용한 곳 수. all=false면 첫 번째만 */
  formatText(find, format, { all = false } = {}) {
    const json = JSON.stringify(this.charMods(format));
    const ms = this._matches(find);
    if (!ms.length) throw new Error(`'${find}'을(를) 찾지 못했습니다 (대소문자 구분)`);
    const targets = all ? ms : ms.slice(0, 1);
    this._batch(() => {
      for (const m of targets) {
        const end = m.charOffset + m.length;
        if (m.cellContext) {
          const c = m.cellContext;
          engine(() => this.doc.applyCharFormatInCell(m.sec, c.parentPara, c.ctrlIdx, c.cellIdx, c.cellPara, m.charOffset, end, json), '글자 서식');
        } else {
          engine(() => this.doc.applyCharFormat(m.sec, m.para, m.charOffset, end, json), '글자 서식');
        }
      }
    });
    return targets.length;
  }

  /**
   * 문단 서식 속성. f: { align, line_spacing(배수 1.6 또는 '160%'), left_indent_mm, right_indent_mm,
   * first_line_indent_mm(+ 들여쓰기 / - 내어쓰기), space_before(pt), space_after(pt), keep_with_next, page_break_before }
   */
  paraMods(f = {}) {
    const p = {};
    if (f.align) {
      if (!ALIGN[f.align]) throw new Error(`정렬은 left·center·right·justify·distribute 중 하나여야 합니다: ${f.align}`);
      p.alignment = ALIGN[f.align];
    }
    if (f.line_spacing !== undefined && f.line_spacing !== null) {
      const s = String(f.line_spacing).trim();
      const pct = s.endsWith('%') ? Number(s.slice(0, -1)) : Number(s) * 100;
      if (!(pct >= 50 && pct <= 500)) throw new Error(`줄 간격이 올바르지 않습니다: ${f.line_spacing} (배수 0.5~5 또는 '160%')`);
      p.lineSpacing = Math.round(pct);
      p.lineSpacingType = 'Percent';
    }
    const mm = (v, what) => {
      const n = Number(v);
      if (!Number.isFinite(n) || Math.abs(n) > 500) throw new Error(`${what}(mm)이 올바르지 않습니다: ${v}`);
      return Math.round(n * PARA_PER_MM);
    };
    if (f.left_indent_mm !== undefined && f.left_indent_mm !== null) p.marginLeft = mm(f.left_indent_mm, '왼쪽 여백');
    if (f.right_indent_mm !== undefined && f.right_indent_mm !== null) p.marginRight = mm(f.right_indent_mm, '오른쪽 여백');
    if (f.first_line_indent_mm !== undefined && f.first_line_indent_mm !== null) p.indent = mm(f.first_line_indent_mm, '첫 줄 들여쓰기');
    const pt = (v, what) => {
      const n = Number(v);
      if (!(n >= 0 && n <= 1000)) throw new Error(`${what}(pt)이 올바르지 않습니다: ${v}`);
      return Math.round(n * PARA_PER_PT);
    };
    if (f.space_before !== undefined && f.space_before !== null) p.spacingBefore = pt(f.space_before, '문단 위 간격');
    if (f.space_after !== undefined && f.space_after !== null) p.spacingAfter = pt(f.space_after, '문단 아래 간격');
    if (f.keep_with_next !== undefined && f.keep_with_next !== null) p.keepWithNext = !!f.keep_with_next;
    if (f.page_break_before !== undefined && f.page_break_before !== null) p.pageBreakBefore = !!f.page_break_before;
    if (!Object.keys(p).length) throw new Error('바꿀 문단 서식이 없습니다 (align, line_spacing, left_indent_mm, first_line_indent_mm 등 하나 이상)');
    return p;
  }

  // 여러 번 고칠 때 매번 쪽을 다시 나누지 않도록 묶는다 (끝나면 한 번에 다시 나눔)
  _batch(fn) {
    if (this._inBatch) return fn();
    let started = false;
    try { call(this.doc.beginBatch(), '일괄 편집'); started = true; } catch {}
    this._inBatch = started;
    try { return fn(); } finally {
      if (started) { this._inBatch = false; try { this.doc.endBatch(); } catch {} }
    }
  }

  _range(from, to = from) {
    posInt(from, '문단 번호'); posInt(to, '문단 번호');
    if (to < from) throw new Error(`문단 범위가 거꾸로입니다: ${from}~${to}`);
    this.locate(to);
    return Array.from({ length: to - from + 1 }, (_, i) => this.locate(from + i));
  }

  /** 문단 서식: 문단 from~to (1부터) → 바꾼 문단 수 */
  formatParagraphs(from, to, format) {
    const json = JSON.stringify(this.paraMods(format));
    const targets = this._range(from, to);
    this._batch(() => { for (const { sec, para } of targets) engine(() => this.doc.applyParaFormat(sec, para, json), '문단 서식'); });
    return targets.length;
  }

  /** 스타일 적용: 문단 from~to에 문서 스타일을 입힌다 (문단 모양·글자 모양을 스타일 것으로 바꿈) → 스타일 이름 */
  applyStyle(from, to, style) {
    const st = this.findStyle(style);
    const targets = this._range(from, to);
    this._batch(() => {
      for (const { sec, para } of targets) {
        if (st.type === 1) {
          const len = this.doc.getParagraphLength(sec, para);
          if (len) engine(() => this.doc.setCharShapeId(sec, para, 0, len, st.charShapeId), '글자 스타일');
        } else {
          this._applyLook(sec, para, { styleId: st.id, psid: st.paraShapeId, csid: st.charShapeId });
        }
      }
    });
    return st.name;
  }

  // ── 문단 넣기·고치기·지우기 ─────────────────────────────

  // 구역 sec의 문단 para 앞에 빈 문단을 만든다
  _newBefore(sec, para) {
    if (para > 0) {
      engine(() => this.doc.insertParagraph(sec, para), '문단 넣기');
      return { sec, para };
    }
    // 구역 첫 문단에는 구역 정의·머리말 개체가 있어야 한다 → 첫 문단을 맨 앞에서 갈라 원래 글을 다음 문단으로 보낸다
    const vis = this._visibleControls(sec, 0);
    if (vis.length) {
      throw new Error(`${sec ? `구역 ${sec + 1}의 ` : ''}첫 문단에 ${Hwp._describe(vis)}이(가) 있어 그 앞에는 문단을 넣을 수 없습니다 — after_paragraph로 다른 위치를 지정하세요`);
    }
    engine(() => this.doc.splitParagraph(sec, 0, 0), '문단 넣기');
    return { sec, para: 0 };
  }

  // 구역 sec의 문단 para 뒤에 빈 문단을 만든다
  _newAfter(sec, para) {
    engine(() => this.doc.insertParagraph(sec, para + 1), '문단 넣기');
    return { sec, para: para + 1 };
  }

  // ── 문서 끝 빈 문단 재사용 ──────────────────────────────
  // add_table·add_page_break·add_field는 끝에 빈 문단을 남기고 _fresh로 표시한다 → 다음 '끝에 붙이기'가 그 문단을 쓴다.
  // 다른 작업은 그 문단(마지막 문단)을 건드렸을 때만 표시를 지운다.

  _isLast(sec, para) {
    const last = this.doc.getSectionCount() - 1;
    return sec === last && para === this.doc.getParagraphCount(last) - 1;
  }

  // 표시된 끝 빈 문단을 쓸 수 있으면 그 위치, 아니면 null (표시는 지움)
  _takeFresh() {
    if (!this._fresh) return null;
    this._fresh = false;
    const sec = this.doc.getSectionCount() - 1;
    const para = this.doc.getParagraphCount(sec) - 1;
    // 혹시 글·개체가 들어갔으면 덮어쓰지 않는다
    if (this.doc.getParagraphLength(sec, para) !== 0 || this._visibleControls(sec, para).length) return null;
    return { sec, para };
  }

  /**
   * 빈 문단을 만들고 위치를 돌려준다.
   * at: 'end'(기본) | 'start' | {after_paragraph:n} | {before_paragraph:n} | {after_text:'글'} (표 칸 안 글이면 그 표 뒤)
   */
  newParagraphAt(at = 'end') {
    if (at === undefined || at === null || at === 'end') {
      const reuse = this._takeFresh();
      if (reuse) return reuse;
      const sec = this.doc.getSectionCount() - 1;
      const count = this.doc.getParagraphCount(sec);
      engine(() => this.doc.insertParagraph(sec, count), '문단 넣기');
      return { sec, para: count };
    }
    let pos;
    if (at === 'start') pos = this._newBefore(0, 0);
    else if (typeof at === 'object' && at.after_paragraph !== undefined) { const l = this.locate(at.after_paragraph); pos = this._newAfter(l.sec, l.para); }
    else if (typeof at === 'object' && at.before_paragraph !== undefined) { const l = this.locate(at.before_paragraph); pos = this._newBefore(l.sec, l.para); }
    else if (typeof at === 'object' && at.after_text !== undefined) { const m = this._firstMatch(at.after_text); pos = this._newAfter(m.sec, m.para); }
    else throw new Error(`위치(at)를 알 수 없습니다: ${JSON.stringify(at)} ('end', 'start', {after_paragraph:n}, {before_paragraph:n}, {after_text:'글'})`);
    // 새 문단이 맨 끝이면 표시된 빈 문단은 더는 끝이 아니다
    if (this._isLast(pos.sec, pos.para)) this._fresh = false;
    return pos;
  }

  /**
   * 원하는 위치에 문단 넣기. 줄바꿈(\n)마다 문단을 하나씩 만든다.
   * opts: { at, style(문서 스타일 이름, 기본 '바탕글'), like(모양을 따라 할 문단 번호), format(글자 서식), align }
   * → { paragraph: 첫 새 문단 번호, count }
   */
  insertParagraph(text = '', { at, style, like, format, align } = {}) {
    const lines = String(text ?? '').split(/\r?\n/);
    const look = this._look({ style, like });
    const mods = format && Object.keys(format).length ? JSON.stringify(this.charMods(format)) : null;
    const paraJson = align ? JSON.stringify(this.paraMods({ align })) : null;
    const first = this.newParagraphAt(at);
    let { sec, para } = first;
    lines.forEach((line, i) => {
      if (i > 0) para = this._newAfter(sec, para).para;
      if (line) engine(() => this.doc.insertText(sec, para, 0, line), '글 넣기');
      this._applyLook(sec, para, look);
      const len = this.doc.getParagraphLength(sec, para);
      if (mods && len) engine(() => this.doc.applyCharFormat(sec, para, 0, len, mods), '글자 서식');
      if (paraJson) engine(() => this.doc.applyParaFormat(sec, para, paraJson), '문단 서식');
    });
    return { paragraph: this.numberOf(first.sec, first.para), count: lines.length };
  }

  /** 문단 글 바꾸기 (문단·글자 모양 유지). 줄바꿈(\n)이 있으면 같은 모양의 문단을 뒤에 잇는다 → 만든 문단 수 */
  setParagraphText(n, text) {
    const { sec, para } = this.locate(n);
    const wasLast = this._isLast(sec, para);
    const vis = this._visibleControls(sec, para);
    if (vis.length) {
      throw new Error(`문단 ${n}에는 ${Hwp._describe(vis)}이(가) 있어 글 전체를 바꿀 수 없습니다 — 일부만 고치려면 replace, 누름틀은 fill, 표 칸은 table_set_cell을 쓰세요`);
    }
    const lines = String(text ?? '').split(/\r?\n/);
    // 글을 다 지우면 글자 모양이 기본으로 돌아가므로 첫 글자 모양을 기억해 두었다가 새 글에 입힌다
    const csid = call(this.doc.getCharPropertiesAt(sec, para, 0), '글자 모양 조회').charShapeId;
    const len = this.doc.getParagraphLength(sec, para);
    if (len) engine(() => this.doc.deleteText(sec, para, 0, len), '글 지우기');
    if (lines[0]) {
      engine(() => this.doc.insertText(sec, para, 0, lines[0]), '글 넣기');
      if (csid !== undefined) engine(() => this.doc.setCharShapeId(sec, para, 0, this.doc.getParagraphLength(sec, para), csid), '글자 모양');
    }
    if (lines.length > 1) {
      const look = this._look({ like: n });
      let cur = para;
      for (const line of lines.slice(1)) {
        cur = this._newAfter(sec, cur).para;
        if (line) engine(() => this.doc.insertText(sec, cur, 0, line), '글 넣기');
        this._applyLook(sec, cur, look);
      }
    }
    if (wasLast) this._fresh = false;
    return lines.length;
  }

  /** 문단 지우기 (표·그림이 든 문단이면 함께 지움) → 함께 지운 개체 설명 ('' = 없음) */
  deleteParagraph(n) {
    const { sec, para } = this.locate(n);
    const count = this.doc.getParagraphCount(sec);
    if (count === 1) throw new Error(`구역 ${sec + 1}의 하나뿐인 문단은 지울 수 없습니다 — 글을 비우려면 set_paragraph_text로 빈 글을 넣으세요`);
    const vis = this._visibleControls(sec, para);
    if (this._isLast(sec, para)) this._fresh = false;
    if (para > 0) {
      engine(() => this.doc.deleteParagraph(sec, para), '문단 지우기');
    } else {
      // 구역 첫 문단을 통째로 지우면 구역 정의·머리말·꼬리말이 함께 사라진다
      // → 글과 내용 개체만 지우고 다음 문단을 끌어올린다
      const bad = vis.filter((c) => !['tbl', 'gso', 'fn', 'en', 'eqed'].includes(c.id));
      if (bad.length) throw new Error(`문단 ${n}(구역 첫 문단)에 ${Hwp._describe(bad)}이(가) 있어 지울 수 없습니다 — replace로 글을 고치세요`);
      for (const c of [...vis].sort((a, b) => b.index - a.index)) {
        if (c.id === 'tbl') engine(() => this.doc.deleteTableControl(sec, 0, c.index), '표 지우기');
        else if (c.id === 'fn' || c.id === 'en') engine(() => this.doc.deleteFootnote(sec, 0, c.index), '각주 지우기');
        else if (c.id === 'eqed') engine(() => this.doc.deleteEquationControl(sec, 0, c.index), '수식 지우기');
        else {
          try { call(this.doc.deletePictureControl(sec, 0, c.index), '그림 지우기'); } catch {
            engine(() => this.doc.deleteShapeControl(sec, 0, c.index), '도형 지우기');
          }
        }
      }
      const len = this.doc.getParagraphLength(sec, 0);
      if (len) engine(() => this.doc.deleteText(sec, 0, 0, len), '글 지우기');
      const look = this._look({ like: this.numberOf(sec, 1) });
      const nextLen = this.doc.getParagraphLength(sec, 1);
      const runs = nextLen ? this.doc.getCharShapeRuns(sec, 1, 0, nextLen) : null;
      engine(() => this.doc.mergeParagraph(sec, 1), '문단 지우기');
      engine(() => this.doc.applyStyle(sec, 0, look.styleId), '스타일 적용');
      engine(() => this.doc.setParaShapeId(sec, 0, look.psid), '문단 모양');
      if (runs) engine(() => this.doc.setCharShapeRuns(sec, 0, 0, nextLen, runs), '글자 모양');
    }
    return Hwp._describe(vis);
  }

  // ── 쓰기 (문서 끝에 붙이기) ──────────────────────────────

  // 마지막 구역 끝에 새 빈 문단을 만들고 그 위치를 돌려준다.
  _newParagraph() {
    const reuse = this._takeFresh();
    if (reuse) return reuse;
    const sec = this.doc.getSectionCount() - 1;
    const last = this.doc.getParagraphCount(sec) - 1;
    const r = call(this.doc.splitParagraph(sec, last, this.doc.getParagraphLength(sec, last)), '문단 추가');
    return { sec, para: r.paraIdx };
  }

  // charProps + 글꼴(font) — 글꼴은 문서 글꼴 목록에 등록한 번호로 건다
  _charProps(style = {}) {
    const p = charProps(style);
    if (style.font) {
      const id = this.doc.findOrCreateFontId(String(style.font));
      if (id >= 0) p.fontId = id;
    }
    return p;
  }

  /**
   * 문단 추가.
   * style: { bold, italic, underline, size(pt), color('#rrggbb'), font, align('left'|'center'|'right'|'justify') }
   */
  addParagraph(text = '', style = {}) {
    const { sec, para } = this._newParagraph();
    if (text) call(this.doc.insertText(sec, para, 0, text), '텍스트 삽입');
    const len = this.doc.getParagraphLength(sec, para);
    if (len) call(this.doc.applyCharFormat(sec, para, 0, len, JSON.stringify(this._charProps(style))), '글자 서식');
    call(this.doc.applyParaFormat(sec, para, JSON.stringify({ alignment: ALIGN[style.align] ?? 'justify' })), '문단 서식');
    return { section: sec, index: para };
  }

  /** 제목 (level 1~3 → 18/15/13pt 굵게) */
  addHeading(text, level = 1, style = {}) {
    const size = { 1: 18, 2: 15, 3: 13 }[level] ?? 12;
    return this.addParagraph(text, { bold: true, size, ...style });
  }

  /**
   * 표 추가. rows = [['이름','나이'], ['홍길동','30']]
   * options: { headerBold: true }
   */
  addTable(rows, { headerBold = true } = {}) {
    const rowCount = rows.length;
    const colCount = Math.max(...rows.map((r) => r.length));
    const { sec, para } = this._newParagraph();
    // 앞 문단이 가운데 정렬이면 표 칸까지 물려받으므로 기본 정렬로 되돌린다
    call(this.doc.applyParaFormat(sec, para, JSON.stringify({ alignment: 'justify' })), '문단 서식');
    const t = call(this.doc.createTable(sec, para, 0, rowCount, colCount), '표 만들기');
    // 칸 안 위·아래 여백: 엔진 기본(0.5mm)은 글자가 테두리에 붙어 보인다 → 1mm (왼쪽·오른쪽 1.8mm는 그대로)
    try { this.doc.setTableProperties(sec, t.paraIdx, t.controlIdx, JSON.stringify({ paddingTop: CELL_PAD_TB, paddingBottom: CELL_PAD_TB })); } catch {}
    rows.forEach((row, r) => row.forEach((value, c) => {
      const text = String(value ?? '');
      if (!text) return;
      const cell = r * colCount + c;
      call(this.doc.insertTextInCell(sec, t.paraIdx, t.controlIdx, cell, 0, 0, text), '셀 입력');
      // 셀은 앞 문단의 글자 모양을 물려받으므로 기본 서식으로 덮어쓴다
      call(this.doc.applyCharFormatInCell(sec, t.paraIdx, t.controlIdx, cell, 0, 0, text.length,
        JSON.stringify(charProps({ bold: headerBold && r === 0 }))), '셀 서식');
    }));
    // 표 다음 문단은 엔진이 만들어 둔 빈 문단을 재사용
    this._fresh = this.doc.getParagraphLength(sec, this.doc.getParagraphCount(sec) - 1) === 0;
    return { section: sec, paraIndex: t.paraIdx, controlIndex: t.controlIdx };
  }

  /** 쪽 나누기 */
  addPageBreak() {
    const sec = this.doc.getSectionCount() - 1;
    const last = this.doc.getParagraphCount(sec) - 1;
    call(this.doc.insertPageBreak(sec, last, this.doc.getParagraphLength(sec, last)), '쪽 나누기');
    this._fresh = this.doc.getParagraphLength(sec, this.doc.getParagraphCount(sec) - 1) === 0;
  }

  // ── 표 ──────────────────────────────────────────────────

  // 표의 칸 배치: grid[행][열] = 칸 번호(병합 칸은 덮는 자리마다 같은 번호), cells[칸] = {row, col, rowSpan, colSpan}
  _grid({ section, paraIndex, controlIndex }) {
    const dim = call(this.doc.getTableDimensions(section, paraIndex, controlIndex), '표 크기');
    const grid = Array.from({ length: dim.rowCount }, () => Array(dim.colCount).fill(undefined));
    const cells = [];
    for (let i = 0; i < dim.cellCount; i++) {
      let info;
      try { info = JSON.parse(this.doc.getCellInfo(section, paraIndex, controlIndex, i)); } catch { info = null; }
      if (!info || info.row === undefined) {
        info = { row: Math.floor(i / dim.colCount), col: i % dim.colCount, rowSpan: 1, colSpan: 1 };
      }
      cells[i] = info;
      for (let r = info.row; r < Math.min(dim.rowCount, info.row + (info.rowSpan || 1)); r++) {
        for (let c = info.col; c < Math.min(dim.colCount, info.col + (info.colSpan || 1)); c++) grid[r][c] ??= i;
      }
    }
    return { rows: dim.rowCount, cols: dim.colCount, grid, cells };
  }

  _cellText({ section, paraIndex, controlIndex }, cell) {
    const n = this.doc.getCellParagraphCount(section, paraIndex, controlIndex, cell);
    const parts = [];
    for (let p = 0; p < n; p++) {
      const len = this.doc.getCellParagraphLength(section, paraIndex, controlIndex, cell, p);
      let t = len ? this.doc.getTextInCell(section, paraIndex, controlIndex, cell, p, 0, len) : '';
      try { const j = JSON.parse(t); t = typeof j === 'string' ? j : (j.text ?? t); } catch {}
      parts.push(t);
    }
    return parts.join('\n');
  }

  /** 표 읽기 → 2차원 배열 (병합 칸은 왼쪽 위 자리에만 글, 덮인 자리는 ''). 위치는 addTable 반환값 또는 tables() 결과 */
  readTable(t) {
    const g = this._grid(t);
    return g.grid.map((row, r) => row.map((cell, c) => {
      if (cell === undefined) return '';
      const info = g.cells[cell];
      return info.row === r && info.col === c ? this._cellText(t, cell) : '';
    }));
  }

  // 첫 행 글 (살펴보기용): 앞 maxCells칸, 칸마다 maxChars자까지. 병합으로 덮인 자리는 ''
  _firstRow(t, { maxCells = 10, maxChars = 60 } = {}) {
    try {
      const g = this._grid(t);
      return (g.grid[0] ?? []).slice(0, maxCells).map((cell, c) => {
        if (cell === undefined || g.cells[cell].row !== 0 || g.cells[cell].col !== c) return '';
        const s = this._cellText(t, cell);
        return s.length > maxChars ? `${s.slice(0, maxChars)}…` : s;
      });
    } catch { return []; }
  }

  /** 문서 안의 본문 표 위치 [{section, paraIndex, controlIndex, rows, cols, paragraph}] (문서 순서, 표 안 표 제외) */
  tables() {
    const found = [];
    const counts = this._counts();
    const starts = counts.map((_, s) => counts.slice(0, s).reduce((a, b) => a + b, 0));
    let list = [];
    try { list = JSON.parse(this.doc.getControls()); } catch {}
    for (const c of list) {
      if (c.list !== 0 || String(c.ctrlId ?? '').trim() !== 'tbl') continue;
      let s = counts.length - 1;
      while (s > 0 && c.para < starts[s]) s--;
      const p = c.para - starts[s];
      let dim;
      try { dim = JSON.parse(this.doc.getTableDimensions(s, p, c.controlIndex)); } catch { continue; }
      if (!dim || dim.ok === false || dim.rowCount === undefined) continue;
      found.push({ section: s, paraIndex: p, controlIndex: c.controlIndex, rows: dim.rowCount, cols: dim.colCount, paragraph: c.para + 1 });
    }
    return found;
  }

  _table(n) {
    const ts = this.tables();
    if (!Number.isInteger(n) || n < 1 || n > ts.length) throw new Error(`표 번호가 범위를 벗어났습니다: ${n} (표 ${ts.length}개, 1부터)`);
    return ts[n - 1];
  }

  _setCellText(t, cell, text) {
    const { section: s, paraIndex: p, controlIndex: c } = t;
    const n = this.doc.getCellParagraphCount(s, p, c, cell);
    if (n > 1) {
      const lastLen = this.doc.getCellParagraphLength(s, p, c, cell, n - 1);
      engine(() => this.doc.deleteRangeInCell(s, p, c, cell, 0, 0, n - 1, lastLen), '칸 비우기');
    }
    const len = this.doc.getCellParagraphLength(s, p, c, cell, 0);
    if (len) engine(() => this.doc.deleteTextInCell(s, p, c, cell, 0, 0, len), '칸 비우기');
    String(text ?? '').split(/\r?\n/).forEach((line, i) => {
      if (i > 0) engine(() => this.doc.splitParagraphInCell(s, p, c, cell, i - 1, this.doc.getCellParagraphLength(s, p, c, cell, i - 1)), '칸 줄 나누기');
      if (line) engine(() => this.doc.insertTextInCell(s, p, c, cell, i, 0, line), '칸 입력');
    });
  }

  /** 표 n의 (행, 열) 칸 글 바꾸기 (모두 1부터, 칸 모양 유지) */
  setCell(n, row, col, text) {
    const t = this._table(n);
    const g = this._grid(t);
    if (!Number.isInteger(row) || row < 1 || row > g.rows) throw new Error(`표 ${n}의 행 번호가 범위를 벗어났습니다: ${row} (전체 ${g.rows}행)`);
    if (!Number.isInteger(col) || col < 1 || col > g.cols) throw new Error(`표 ${n}의 열 번호가 범위를 벗어났습니다: ${col} (전체 ${g.cols}열)`);
    const cell = g.grid[row - 1][col - 1];
    if (cell === undefined) throw new Error(`표 ${n}의 ${row}행 ${col}열 칸을 찾지 못했습니다`);
    this._setCellText(t, cell, text);
  }

  /**
   * 표 n에 행 추가. afterRow: 0이면 맨 위, 생략하면 맨 아래. values: 왼쪽부터 칸 글
   * → { row: 새 행 번호, merged: 위아래 병합 칸에 덮여 새 칸이 없는 열 번호들 }
   * 병합에 덮인 열에 값을 주면 표를 바꾸기 전에 거절한다 (값이 조용히 사라지지 않게).
   */
  addTableRow(n, { afterRow, values = [] } = {}) {
    const t = this._table(n);
    const { section: s, paraIndex: p, controlIndex: c } = t;
    const before = this._grid(t);
    const after = afterRow ?? before.rows;
    if (!Number.isInteger(after) || after < 0 || after > before.rows) throw new Error(`after_row가 범위를 벗어났습니다: ${after} (0~${before.rows})`);
    if (values.length > before.cols) throw new Error(`값이 ${values.length}개인데 표 ${n}은 ${before.cols}열입니다`);
    // 위아래로 병합된 칸 사이에 넣으면 그 병합이 새 행까지 늘어나 새 행의 그 열에는 칸이 없다
    const spanned = (j) => after > 0 && after < before.rows && before.grid[after - 1][j] !== undefined && before.grid[after - 1][j] === before.grid[after][j];
    const merged = Array.from({ length: before.cols }, (_, j) => j).filter(spanned);
    const blocked = (cols) => cols.filter((j) => String(values[j] ?? '') !== '');
    const refuse = (cols) => {
      throw new Error(`표 ${n}의 새 ${after + 1}행에서 ${cols.map((j) => `${j + 1}열`).join('·')}은 위 칸과 세로로 병합된 자리라 값을 넣을 수 없습니다: `
        + `${cols.map((j) => `'${values[j]}'`).join(', ')} — 그 자리 값을 ''로 비우거나 병합이 끝난 행 뒤(after_row)에 넣으세요`);
    };
    if (blocked(merged).length) refuse(blocked(merged));
    if (after === 0) engine(() => this.doc.insertTableRow(s, p, c, 0, false), '행 추가');
    else engine(() => this.doc.insertTableRow(s, p, c, after - 1, true), '행 추가');
    const g = this._grid(t);
    const r0 = after; // 새 행 (0부터)
    const own = (j) => { const cell = g.grid[r0]?.[j]; return cell !== undefined && g.cells[cell].row === r0 && g.cells[cell].col === j; };
    // 엔진이 예상과 다르게 병합을 늘렸으면 (문서는 아직 저장 전) 값을 버리지 않고 거절
    const lost = blocked(values.map((_, j) => j).filter((j) => !own(j)));
    if (lost.length) refuse(lost);
    // 모양 본보기: 바로 아래 행(새 행에 밀려난 행), 맨 끝에 붙였으면 바로 위 행
    const ref = r0 + 1 < g.rows ? r0 + 1 : r0 - 1;
    values.forEach((v, j) => {
      const cell = g.grid[r0]?.[j];
      if (!own(j)) return;
      const text = String(v ?? '');
      this._setCellText(t, cell, text);
      const refCell = g.grid[ref]?.[j];
      if (refCell === undefined || refCell === cell) return;
      try {
        const csid = JSON.parse(this.doc.getCellCharPropertiesAt(s, p, c, refCell, 0, 0)).charShapeId;
        const psid = JSON.parse(this.doc.getCellParaPropertiesAt(s, p, c, refCell, 0)).paraShapeId;
        const paras = this.doc.getCellParagraphCount(s, p, c, cell);
        for (let k = 0; k < paras; k++) {
          if (psid !== undefined) call(this.doc.setCellParaShapeId(s, p, c, cell, k, psid), '칸 문단 모양');
          const len = this.doc.getCellParagraphLength(s, p, c, cell, k);
          if (csid !== undefined && len) call(this.doc.setCharShapeIdInCell(s, p, c, cell, k, 0, len, csid), '칸 글자 모양');
        }
      } catch {}
    });
    return { row: r0 + 1, merged: Array.from({ length: g.cols }, (_, j) => j).filter((j) => !own(j)).map((j) => j + 1) };
  }

  /** 표 n의 행 지우기 (1부터) */
  deleteTableRow(n, row) {
    const t = this._table(n);
    const dim = call(this.doc.getTableDimensions(t.section, t.paraIndex, t.controlIndex), '표 크기');
    if (!Number.isInteger(row) || row < 1 || row > dim.rowCount) throw new Error(`표 ${n}의 행 번호가 범위를 벗어났습니다: ${row} (전체 ${dim.rowCount}행)`);
    if (dim.rowCount === 1) throw new Error(`표 ${n}의 하나뿐인 행은 지울 수 없습니다 — 표를 지우려면 delete_paragraph로 표가 든 문단(${t.paragraph})을 지우세요`);
    engine(() => this.doc.deleteTableRow(t.section, t.paraIndex, t.controlIndex, row - 1), '행 지우기');
  }

  // ── 머리말·꼬리말 ───────────────────────────────────────

  /** 머리말/꼬리말 글: [{section, applyTo('양쪽'|'짝수쪽'|'홀수쪽'), text}] ({쪽}·{전체쪽}은 쪽 번호 자리) */
  headerFooter(kind) {
    const isH = kind === 'header';
    const out = [];
    for (let s = 0; s < this.doc.getSectionCount(); s++) {
      for (const applyTo of [0, 1, 2]) {
        let r;
        try { r = JSON.parse(this.doc.getHeaderFooter(s, isH, applyTo)); } catch { continue; }
        if (!r?.exists) continue;
        const text = String(r.text ?? '').replace(/\u0015/g, '{쪽}').replace(/\u0016/g, '{전체쪽}').replace(/\u0017/g, '{파일 이름}');
        out.push({ section: s, applyTo: r.label ?? ['양쪽', '짝수쪽', '홀수쪽'][applyTo], text });
      }
    }
    return out;
  }

  // 머리말/꼬리말 글을 모두 비운다 (모양은 첫 글자 것 유지). 못 비우면 false
  _clearHeaderFooter(s, isH) {
    try {
      const info = () => JSON.parse(this.doc.getHeaderFooter(s, isH, 0));
      for (let i = info().paraCount - 1; i > 0; i--) call(this.doc.mergeParagraphInHeaderFooter(s, isH, 0, i), '머리말 비우기');
      for (let k = 0; k < 3; k++) {
        const len = JSON.parse(this.doc.getHeaderFooterParaInfo(s, isH, 0, 0)).charCount ?? 0;
        if (!len) return info().paraCount === 1;
        call(this.doc.deleteTextInHeaderFooter(s, isH, 0, 0, 0, len), '머리말 비우기');
      }
    } catch {}
    return false;
  }

  /**
   * 머리말(kind='header')/꼬리말('footer') 정하기 — 모든 구역에 같은 글 (홀수·짝수쪽 따로 있던 것은 지움).
   * text 안의 {page}는 쪽 번호, {pages}는 전체 쪽 수. pageNumbers=true면 글 끝에 쪽 번호(글이 없으면 '- 1 -' 모양).
   * text가 비고 pageNumbers도 없으면 머리말/꼬리말을 지운다. → 'set' | 'removed'
   */
  setHeaderFooter(kind, text = '', { pageNumbers = false, align } = {}) {
    const isH = kind === 'header';
    let body = String(text ?? '');
    if (pageNumbers && !/\{pages?\}/.test(body)) body = body ? `${body} {page}` : '- {page} -';
    const sections = this.doc.getSectionCount();
    const exists = (s, a) => { try { return JSON.parse(this.doc.getHeaderFooter(s, isH, a)).exists; } catch { return false; } };
    if (!body) {
      for (let s = 0; s < sections; s++) for (const a of [0, 1, 2]) if (exists(s, a)) engine(() => this.doc.deleteHeaderFooter(s, isH, a), '머리말 지우기');
      return 'removed';
    }
    // 쪽 번호만 있는 머리말/꼬리말은 따로 정하지 않으면 가운데
    if (!align && !body.replace(/\{pages?\}/g, '').replace(/[\s\-–—/|·.()[\]]/g, '')) align = 'center';
    const paraJson = align ? JSON.stringify(this.paraMods({ align })) : null;
    for (let s = 0; s < sections; s++) {
      for (const a of [1, 2]) if (exists(s, a)) engine(() => this.doc.deleteHeaderFooter(s, isH, a), '머리말 지우기');
      if (exists(s, 0) && !this._clearHeaderFooter(s, isH)) engine(() => this.doc.deleteHeaderFooter(s, isH, 0), '머리말 지우기');
      if (!exists(s, 0)) engine(() => this.doc.createHeaderFooter(s, isH, 0), isH ? '머리말 만들기' : '꼬리말 만들기');
      body.split(/\r?\n/).forEach((line, i) => {
        let off = 0;
        if (i > 0) engine(() => this.doc.splitParagraphInHeaderFooter(s, isH, 0, i - 1, JSON.parse(this.doc.getHeaderFooterParaInfo(s, isH, 0, i - 1)).charCount), '머리말 줄 나누기');
        for (const part of line.split(/(\{pages?\})/)) {
          if (!part) continue;
          if (part === '{page}' || part === '{pages}') {
            off = engine(() => this.doc.insertFieldInHf(s, isH, 0, i, off, part === '{page}' ? 1 : 2), '쪽 번호 넣기').charOffset;
          } else {
            off = engine(() => this.doc.insertTextInHeaderFooter(s, isH, 0, i, off, part), '머리말 글 넣기').charOffset;
          }
        }
        if (paraJson) engine(() => this.doc.applyParaFormatInHf(s, isH, 0, i, paraJson), '머리말 정렬');
      });
    }
    return 'set';
  }

  // ── 각주·그림·쪽 설정 ───────────────────────────────────

  /** 각주: find(본문 글, 첫 번째) 바로 뒤에 각주를 달고 내용을 넣는다 → { number, paragraph } */
  addFootnote(find, text) {
    const m = this._firstMatch(find, { cells: false });
    const r = engine(() => this.doc.insertFootnote(m.sec, m.para, m.charOffset + m.length), '각주 넣기');
    const para = r.paraIdx ?? m.para;
    const info = engine(() => this.doc.getFootnoteInfo(m.sec, para, r.controlIdx), '각주 조회');
    const body = String(text ?? '').replace(/\r?\n/g, ' ');
    if (body) engine(() => this.doc.insertTextInFootnote(m.sec, para, r.controlIdx, 0, info.totalTextLen ?? 0, body), '각주 글 넣기');
    return { number: r.footnoteNumber ?? info.number, paragraph: this.numberOf(m.sec, para) };
  }

  // 구역 본문 크기(mm) { width, height } — 한컴 방식: 본문 높이 = 쪽 높이 - 위·아래 여백 - 머리말·꼬리말 자리
  _bodySizeMm(sec) {
    const pd = call(this.doc.getPageDef(sec), '쪽 설정 조회');
    const [w, h] = pd.landscape ? [pd.height, pd.width] : [pd.width, pd.height];
    return {
      width: (w - pd.marginLeft - pd.marginRight - (pd.marginGutter ?? 0)) / HU_PER_MM,
      height: (h - pd.marginTop - pd.marginBottom - pd.marginHeader - pd.marginFooter) / HU_PER_MM,
    };
  }

  /**
   * 그림 넣기 (글자처럼 취급하는 그림 하나만 든 새 문단).
   * src: 파일 경로 또는 Buffer. opts: { at, widthMm, heightMm, align(기본 'center'), description }
   * 크기를 안 주면 96dpi 원래 크기. 본문보다 넓거나 높으면 본문 안에 들어가게 줄인다 (가로세로 비율 유지)
   * → { paragraph, widthMm, heightMm, shrunk, fit('width'|'height'|null: 무엇에 맞춰 줄였는지) }
   */
  insertImage(src, { at, widthMm, heightMm, align = 'center', description = '' } = {}) {
    const bytes = Buffer.isBuffer(src) || src instanceof Uint8Array ? Buffer.from(src) : readFileSync(src);
    const img = imageInfo(bytes);
    if (!img.width || !img.height) throw new Error('그림의 가로·세로 크기를 읽지 못했습니다');
    const ratio = img.height / img.width;
    let w; let h;
    if (widthMm && heightMm) { w = widthMm; h = heightMm; } else if (widthMm) { w = widthMm; h = widthMm * ratio; } else if (heightMm) { h = heightMm; w = heightMm / ratio; } else { w = (img.width * 25.4) / 96; h = w * ratio; }
    if (!(w > 0 && h > 0) || w > 2000 || h > 2000) throw new Error(`그림 크기(mm)가 올바르지 않습니다: ${w}×${h}`);
    const paraJson = JSON.stringify(this.paraMods({ align }));
    const pos = this.newParagraphAt(at);
    let fit = null;
    const body = this._bodySizeMm(pos.sec);
    // 높이는 줄 간격 여유를 조금 남긴다 (꽉 채우면 문단이 본문보다 길어져 다음 쪽으로 밀림)
    const maxW = body.width; const maxH = body.height - 2;
    if (maxW > 10 && w > maxW) { h *= maxW / w; w = maxW; fit = 'width'; }
    if (maxH > 10 && h > maxH) { w *= maxH / h; h = maxH; fit = 'height'; }
    this._applyLook(pos.sec, pos.para, this._look({}));
    engine(() => this.doc.applyParaFormat(pos.sec, pos.para, paraJson), '문단 서식');
    const r = engine(() => this.doc.insertPicture(pos.sec, pos.para, 0, '', bytes, Math.round(w * HU_PER_MM), Math.round(h * HU_PER_MM),
      img.width, img.height, img.ext, String(description ?? '')), '그림 넣기');
    engine(() => this.doc.setPictureProperties(pos.sec, r.paraIdx ?? pos.para, r.controlIdx ?? 0, JSON.stringify({ treatAsChar: true })), '그림 배치');
    return { paragraph: this.numberOf(pos.sec, r.paraIdx ?? pos.para), widthMm: Math.round(w * 10) / 10, heightMm: Math.round(h * 10) / 10, shrunk: !!fit, fit };
  }

  /** 구역 쪽 설정 [{section, paper, orientation, widthMm, heightMm, margins: {top,bottom,left,right,header,footer,gutter}}] (mm) */
  pageSetups() {
    const out = [];
    for (let s = 0; s < this.doc.getSectionCount(); s++) {
      const pd = call(this.doc.getPageDef(s), '쪽 설정 조회');
      const a = Math.min(pd.width, pd.height) / HU_PER_MM; const b = Math.max(pd.width, pd.height) / HU_PER_MM;
      const paper = Object.entries(PAPERS).find(([, [x, y]]) => Math.abs(x - a) < 1.5 && Math.abs(y - b) < 1.5)?.[0] ?? '사용자 지정';
      const [w, h] = pd.landscape ? [pd.height, pd.width] : [pd.width, pd.height];
      out.push({
        section: s, paper, orientation: pd.landscape ? 'landscape' : 'portrait', widthMm: mmOf(w), heightMm: mmOf(h),
        margins: { top: mmOf(pd.marginTop), bottom: mmOf(pd.marginBottom), left: mmOf(pd.marginLeft), right: mmOf(pd.marginRight),
          header: mmOf(pd.marginHeader), footer: mmOf(pd.marginFooter), gutter: mmOf(pd.marginGutter) },
      });
    }
    return out;
  }

  /**
   * 쪽 설정. opts: { orientation('portrait'|'landscape'), paper('A4'|'A3'|'A5'|'B4'|'B5'|'Letter'|'Legal'),
   * margins_mm: {top, bottom, left, right, header, footer, gutter}, section(1부터, 생략하면 모든 구역) }
   * 한글 방식: 위·아래 여백은 머리말·꼬리말 자리와 별개다 (본문은 위 여백 + 머리말 아래에서 시작).
   */
  pageSetup({ orientation, paper, margins_mm: margins, section } = {}) {
    if (orientation === undefined && paper === undefined && margins === undefined) throw new Error('바꿀 쪽 설정이 없습니다 (orientation, paper, margins_mm 중 하나 이상)');
    if (orientation !== undefined && !['portrait', 'landscape'].includes(orientation)) throw new Error(`방향은 portrait 또는 landscape여야 합니다: ${orientation}`);
    let size;
    if (paper !== undefined) {
      const key = Object.keys(PAPERS).find((k) => k.toLowerCase() === String(paper).toLowerCase());
      if (!key) throw new Error(`용지 크기를 알 수 없습니다: ${paper} (가능: ${Object.keys(PAPERS).join(', ')})`);
      size = PAPERS[key];
    }
    const total = this.doc.getSectionCount();
    if (section !== undefined && (!Number.isInteger(section) || section < 1 || section > total)) throw new Error(`구역 번호가 범위를 벗어났습니다: ${section} (전체 ${total}개)`);
    const KEYS = { top: 'marginTop', bottom: 'marginBottom', left: 'marginLeft', right: 'marginRight', header: 'marginHeader', footer: 'marginFooter', gutter: 'marginGutter' };
    for (const k of Object.keys(margins ?? {})) {
      if (!KEYS[k]) throw new Error(`여백 이름을 알 수 없습니다: ${k} (top, bottom, left, right, header, footer, gutter)`);
      const v = Number(margins[k]);
      if (!(v >= 0 && v <= 300)) throw new Error(`여백 ${k}(mm)이 올바르지 않습니다: ${margins[k]}`);
    }
    const targets = section ? [section - 1] : Array.from({ length: total }, (_, i) => i);
    for (const s of targets) {
      const pd = call(this.doc.getPageDef(s), '쪽 설정 조회');
      const next = { ...pd };
      if (size) { next.width = Math.round(size[0] * HU_PER_MM); next.height = Math.round(size[1] * HU_PER_MM); }
      if (orientation) next.landscape = orientation === 'landscape';
      for (const [k, v] of Object.entries(margins ?? {})) next[KEYS[k]] = Math.round(Number(v) * HU_PER_MM);
      const [w, h] = next.landscape ? [next.height, next.width] : [next.width, next.height];
      if (w - next.marginLeft - next.marginRight < 10 * HU_PER_MM
        || h - next.marginTop - next.marginBottom - next.marginHeader - next.marginFooter < 10 * HU_PER_MM) {
        throw new Error('여백이 너무 커서 본문 자리가 남지 않습니다');
      }
      engine(() => this.doc.setPageDef(s, JSON.stringify(next)), '쪽 설정');
    }
    return targets.length;
  }

  // ── 살펴보기 ────────────────────────────────────────────

  /**
   * 문서 구조 요약 (hwp_inspect). opts: { offset: 건너뛸 문단 수, limit: 보여 줄 문단 수 }
   * → { format, pages, sections:[쪽 설정], paragraphCount, offset, limit, paragraphs:[{n, section, text, style, align, font, size, bold, italic, objects}],
   *     tables:[{n, paragraph, rows, cols, firstRow?}], fields:[{name, value, cell, paragraph}], styles:[이름], headers, footers, counts }
   * tables·fields는 문서 전체 목록이다. 표의 firstRow(앞 10칸, 칸마다 60자)는 문단 범위(offset < 문단 ≤ offset+limit) 안 표에만 붙는다.
   * 누름틀 값은 200자까지.
   */
  inspect({ offset = 0, limit = 150 } = {}) {
    const map = this._controlMap();
    const counts = { pictures: 0, shapes: 0, footnotes: 0, endnotes: 0, equations: 0 };
    for (const list of map.values()) {
      for (const c of list) {
        if (c.id === 'gso') { if (c.desc === '그림') counts.pictures++; else counts.shapes++; }
        else if (c.id === 'fn') counts.footnotes++;
        else if (c.id === 'en') counts.endnotes++;
        else if (c.id === 'eqed') counts.equations++;
      }
    }
    const tables = this.tables();
    const tableNo = new Map(tables.map((t, i) => [`${t.paragraph}:${t.controlIndex}`, i + 1]));
    const styleName = new Map(this.styles().map((s) => [s.id, s.name]));
    const total = this.paragraphCount;
    const paragraphs = [];
    let n = 0;
    for (let s = 0; s < this.doc.getSectionCount() && paragraphs.length < limit; s++) {
      for (let p = 0; p < this.doc.getParagraphCount(s) && paragraphs.length < limit; p++) {
        n++;
        if (n <= offset) continue;
        let style = ''; let para = {}; let ch = {};
        try { const st = JSON.parse(this.doc.getStyleAt(s, p)); style = st.name ?? styleName.get(st.id) ?? ''; } catch {}
        try { para = JSON.parse(this.doc.getParaPropertiesAt(s, p)); } catch {}
        try { ch = JSON.parse(this.doc.getCharPropertiesAt(s, p, 0)); } catch {}
        const objects = (map.get(n - 1) ?? []).filter((c) => !STRUCTURAL.has(c.id)).map((c) => {
          if (c.id === 'tbl') { const k = tableNo.get(`${n}:${c.index}`); const t = tables[k - 1]; return k ? `표 ${k}(${t.rows}행×${t.cols}열)` : '표'; }
          return c.id === 'gso' ? (c.desc || '그림·도형') : (CONTROL_KO[c.id] ?? (c.desc || c.id));
        });
        const text = this._paraText(s, p);
        paragraphs.push({
          n, section: s + 1, text: text.length > 200 ? `${text.slice(0, 200)}…` : text, style,
          align: ALIGN_KO[para.alignment] ?? para.alignment ?? '', leftIndentMm: para.marginLeft ? pxToMm(para.marginLeft) : 0,
          lineSpacing: para.lineSpacingType === 'Percent' ? `${Math.round(para.lineSpacing)}%` : undefined,
          font: ch.fontFamily ?? '', size: ch.fontSize ? ch.fontSize / 100 : undefined, bold: !!ch.bold, italic: !!ch.italic, objects,
        });
      }
    }
    let fields = [];
    try {
      fields = this.fields().map((f) => {
        const value = String(f.value ?? '');
        return {
          name: f.name, value: value.length > 200 ? `${value.slice(0, 200)}…` : value,
          // rhwp는 칸 안 누름틀도 cellField:false로 줄 때가 있어 위치 경로로도 본다
          cell: !!f.cellField || (f.location?.path ?? []).some((x) => x?.type === 'cell'),
          paragraph: f.location ? this.numberOf(f.location.sectionIndex ?? 0, f.location.paraIndex ?? 0) : undefined,
        };
      });
    } catch {}
    const inRange = (para) => para > offset && para <= offset + limit;
    return {
      format: String(this.format ?? '').toLowerCase(), pages: this.pageCount, sections: this.pageSetups(),
      paragraphCount: total, offset, limit, paragraphs,
      tables: tables.map((t, i) => {
        const o = { n: i + 1, paragraph: t.paragraph, rows: t.rows, cols: t.cols };
        if (inRange(t.paragraph)) o.firstRow = this._firstRow(t);
        return o;
      }),
      fields, styles: this.styles().map((s) => s.name), headers: this.headerFooter('header'), footers: this.headerFooter('footer'), counts,
    };
  }

  // ── 누름틀(필드): 서식 문서 채우기 ───────────────────────

  /** 누름틀 목록 [{fieldId, name, guide, value, ...}] */
  fields() { return call(this.doc.getFieldList(), '필드 목록'); }

  /**
   * 누름틀 채우기: fill({ 성명: '홍길동', 날짜: '2026-10-17' }) → 채운 이름 목록.
   * 같은 이름 누름틀이 여럿이면 모두 채운다 (한컴 PutFieldText와 같게). 채운 칸 수는 lastFillCounts에.
   */
  fill(values) {
    const done = [];
    this.lastFillCounts = {};
    for (const [name, value] of Object.entries(values)) {
      const ids = this.fields().filter((f) => f.name === name && f.fieldId !== undefined).map((f) => f.fieldId);
      if (!ids.length) call(this.doc.setFieldValueByName(name, String(value)), `필드 '${name}' 채우기`);
      for (const id of ids) call(this.doc.setFieldValue(id, String(value)), `필드 '${name}' 채우기`);
      this.lastFillCounts[name] = Math.max(ids.length, 1);
      done.push(name);
    }
    return done;
  }

  /** 새 문단에 누름틀 추가 (label이 있으면 "label: [누름틀]" 형태) */
  addField(name, { label, guide = name } = {}) {
    const { sec, para } = this._newParagraph();
    let offset = 0;
    if (label) offset = call(this.doc.insertText(sec, para, 0, `${label}: `), '텍스트 삽입').charOffset;
    const len = this.doc.getParagraphLength(sec, para);
    if (len) call(this.doc.applyCharFormat(sec, para, 0, len, JSON.stringify(charProps())), '글자 서식');
    call(this.doc.applyParaFormat(sec, para, JSON.stringify({ alignment: 'justify' })), '문단 서식');
    // 누름틀 뒤에서 문단을 나누면 누름틀이 두 문단에 걸쳐 쪼개지므로, 다음 빈 문단을 먼저 만들어 둔다
    call(this.doc.splitParagraph(sec, para, offset), '문단 추가');
    this._fresh = true;
    call(this.doc.insertClickHereField(sec, para, offset, guide, '', name, true), '누름틀 삽입');
    return { section: sec, index: para };
  }

  // ── 출력 ────────────────────────────────────────────────

  /** 페이지를 SVG 문자열로 렌더링 (미리보기용) */
  svg(page = 0) { return this.doc.renderPageSvg(page); }

  /** 바이트로 내보내기: 'hwp' | 'hwpx' | 'hml' | 'hwt'(서식 파일: 원래 형식 그대로, 보통 hwp) */
  toBytes(format = 'hwp') {
    if (format === 'hwt') format = String(this.format).toLowerCase() === 'hwpx' ? 'hwpx' : 'hwp';
    if (format === 'hwp') return this.doc.exportHwp();
    if (format === 'hwpx') return this.doc.exportHwpx();
    if (format === 'hml') return this.doc.exportHml();
    throw new Error(`지원하지 않는 형식: ${format}`);
  }

  /** 저장. 형식은 확장자로 결정 (.hwp / .hwpx / .hwt / .hml) */
  save(path) {
    const format = extname(path).slice(1).toLowerCase() || 'hwp';
    const bytes = this.toBytes(format);
    // 쓰는 도중 끊겨도 원본이 망가지지 않게: 같은 폴더 임시 파일에 다 쓴 뒤 이름을 바꿔 끼운다
    const tmp = `${path}.${process.pid}.${Date.now()}.tmp`;
    try {
      writeFileSync(tmp, bytes);
    } catch (e) {
      try { unlinkSync(tmp); } catch {}
      throw e;
    }
    // Windows: 백신·미리보기가 잠깐 잡고 있으면 이름 바꾸기가 실패한다 → 몇 번 다시 해 보고, 그래도 안 되면 바로 쓴다
    for (let i = 0; ; i++) {
      try { renameSync(tmp, path); return path; } catch (e) {
        if (i < 4 && ['EPERM', 'EACCES', 'EBUSY'].includes(e.code)) {
          try { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 50 * (i + 1)); } catch {}
          continue;
        }
        try { unlinkSync(tmp); } catch {}
        if (!['EPERM', 'EACCES', 'EBUSY', 'EXDEV'].includes(e.code)) throw e;
        writeFileSync(path, bytes);
        return path;
      }
    }
  }

  close() { this.doc.free(); }
}
