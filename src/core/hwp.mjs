// hwp.mjs — 한컴오피스 없이 HWP/HWPX를 읽고 쓰는 래퍼.
// HOP(golbin/hop)이 내부에서 쓰는 rhwp 엔진(@rhwp/core, WASM)을 Node에서 직접 호출한다.
import { readFileSync, writeFileSync } from 'node:fs';
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

const ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify', distribute: 'distribute' };
const BASE_CHAR = { bold: false, italic: false, underline: false, strikethrough: false, fontSize: 1000, textColor: '#000000' };

// 사용자 친화 스타일 {bold, italic, underline, size(pt), color, font} → rhwp 글자 속성
function charProps(style = {}) {
  const p = { ...BASE_CHAR };
  for (const k of ['bold', 'italic', 'underline', 'strikethrough']) if (style[k] !== undefined) p[k] = !!style[k];
  if (style.size) p.fontSize = Math.round(style.size * 100);
  if (style.color) p.textColor = style.color;
  if (style.font) p.fontFamily = style.font;
  return p;
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
   * 표 다음 문단이 겹쳐 그려지므로 hwp로 한 번 거쳐 다시 연다. 파일은 바꾸지 않는다.
   */
  static openForView(path, opts) {
    const h = Hwp.open(path, opts);
    if (extname(path).toLowerCase() !== '.hwpx') return h;
    let empty = 0;
    try { empty = JSON.parse(h.doc.getValidationWarnings()).count ?? 0; } catch {}
    if (!empty) return h;
    try {
      const v = new Hwp(new HwpDocument(h.toBytes('hwp')));
      h.close();
      return v;
    } catch { return h; }
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

  /** 본문 문단 목록 [{section, index, text}] (표 셀 제외) */
  paragraphs() {
    const out = [];
    for (let s = 0; s < this.doc.getSectionCount(); s++) {
      for (let p = 0; p < this.doc.getParagraphCount(s); p++) {
        const len = this.doc.getParagraphLength(s, p);
        const raw = len ? this.doc.getTextRange(s, p, 0, len) : '';
        let text = raw;
        try { const j = JSON.parse(raw); text = typeof j === 'string' ? j : (j.text ?? raw); } catch {}
        out.push({ section: s, index: p, text });
      }
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

  // ── 쓰기 ────────────────────────────────────────────────

  // 마지막 구역 끝에 새 빈 문단을 만들고 그 위치를 돌려준다.
  _newParagraph() {
    const sec = this.doc.getSectionCount() - 1;
    const last = this.doc.getParagraphCount(sec) - 1;
    if (this._fresh) { this._fresh = false; return { sec, para: last }; }
    const r = call(this.doc.splitParagraph(sec, last, this.doc.getParagraphLength(sec, last)), '문단 추가');
    return { sec, para: r.paraIdx };
  }

  /**
   * 문단 추가.
   * style: { bold, italic, underline, size(pt), color('#rrggbb'), font, align('left'|'center'|'right'|'justify') }
   */
  addParagraph(text = '', style = {}) {
    const { sec, para } = this._newParagraph();
    if (text) call(this.doc.insertText(sec, para, 0, text), '텍스트 삽입');
    const len = this.doc.getParagraphLength(sec, para);
    if (len) call(this.doc.applyCharFormat(sec, para, 0, len, JSON.stringify(charProps(style))), '글자 서식');
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

  /** 표 읽기 → 2차원 배열. 위치는 addTable 반환값 또는 tables() 결과 */
  readTable({ section, paraIndex, controlIndex }) {
    const dim = call(this.doc.getTableDimensions(section, paraIndex, controlIndex), '표 크기');
    const rows = [];
    for (let r = 0; r < dim.rowCount; r++) {
      const row = [];
      for (let c = 0; c < dim.colCount; c++) {
        const cell = r * dim.colCount + c;
        if (cell >= dim.cellCount) { row.push(''); continue; }
        const n = this.doc.getCellParagraphCount(section, paraIndex, controlIndex, cell);
        const parts = [];
        for (let p = 0; p < n; p++) {
          const len = this.doc.getCellParagraphLength(section, paraIndex, controlIndex, cell, p);
          let t = len ? this.doc.getTextInCell(section, paraIndex, controlIndex, cell, p, 0, len) : '';
          try { const j = JSON.parse(t); t = typeof j === 'string' ? j : (j.text ?? t); } catch {}
          parts.push(t);
        }
        row.push(parts.join('\n'));
      }
      rows.push(row);
    }
    return rows;
  }

  /** 문서 안의 모든 표 위치 [{section, paraIndex, controlIndex, rows, cols}] */
  tables() {
    const found = [];
    for (let s = 0; s < this.doc.getSectionCount(); s++) {
      for (let p = 0; p < this.doc.getParagraphCount(s); p++) {
        for (let c = 0; c < 16; c++) {
          let dim;
          try { dim = JSON.parse(this.doc.getTableDimensions(s, p, c)); } catch { break; }
          if (!dim || dim.ok === false || dim.rowCount === undefined) break;
          found.push({ section: s, paraIndex: p, controlIndex: c, rows: dim.rowCount, cols: dim.colCount });
        }
      }
    }
    return found;
  }

  // ── 누름틀(필드): 서식 문서 채우기 ───────────────────────

  /** 누름틀 목록 [{fieldId, name, guide, value, ...}] */
  fields() { return call(this.doc.getFieldList(), '필드 목록'); }

  /** 누름틀 채우기: fill({ 성명: '홍길동', 날짜: '2026-10-17' }) → 채운 이름 목록 */
  fill(values) {
    const done = [];
    for (const [name, value] of Object.entries(values)) {
      call(this.doc.setFieldValueByName(name, String(value)), `필드 '${name}' 채우기`);
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
    writeFileSync(path, this.toBytes(format));
    return path;
  }

  close() { this.doc.free(); }
}
