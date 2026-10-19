// create.mjs — 마크다운 원고 하나로 hwp/hwpx/docx/pptx/xlsx 새 문서를 만든다. 빈 문서·새 파일 이름 도우미도 여기.
import { existsSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { extname, isAbsolute, join, resolve } from 'node:path';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, PageBreak,
} from 'docx';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Hwp } from './hwp.mjs';

/**
 * 원고 문법
 *   # 제목 / ## 소제목 / ### 작은 제목     >> 가운데 줄     **굵은 한 줄**
 *   - 글머리표                               | 표 | 행 |  (|---| 줄 무시)
 *   라벨: {{필드}}  또는 {{필드}}  (누름틀/자리표시자)
 *   ---page---  (쪽 나누기 / 새 슬라이드)
 */
export function parseMarkdown(src) {
  const lines = src.replace(/^﻿/, '').replace(/\r\n/g, '\n').split('\n');
  const blocks = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/^\s*\|/.test(line)) {
      const rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        const cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
        if (!cells.every((c) => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      i--;
      blocks.push({ type: 'table', rows });
      continue;
    }
    let m;
    if ((m = line.match(/^(#{1,3})\s+(.*)$/))) blocks.push({ type: 'heading', level: m[1].length, text: m[2] });
    else if (/^\s*-{3}\s*page\s*-{3}\s*$/i.test(line)) blocks.push({ type: 'pagebreak' });
    else if ((m = line.match(/^(?:(.*?):\s*)?\{\{\s*(.+?)\s*\}\}\s*$/))) blocks.push({ type: 'field', name: m[2], label: m[1]?.trim() || null });
    else if ((m = line.match(/^\*\*(.+)\*\*$/))) blocks.push({ type: 'para', text: m[1], bold: true });
    else if ((m = line.match(/^>>\s?(.*)$/))) blocks.push({ type: 'para', text: m[1], center: true });
    else if ((m = line.match(/^\s*[-*]\s+(.*)$/))) blocks.push({ type: 'bullet', text: m[1] });
    else blocks.push({ type: 'para', text: line });
  }
  return blocks;
}

function toHwp(blocks) {
  const h = Hwp.create();
  for (const b of blocks) {
    if (b.type === 'heading') h.addHeading(b.text, b.level, b.level === 1 ? { align: 'center' } : {});
    else if (b.type === 'table') h.addTable(b.rows);
    else if (b.type === 'pagebreak') h.addPageBreak();
    else if (b.type === 'field') h.addField(b.name, { label: b.label ?? undefined });
    else if (b.type === 'bullet') h.addParagraph(`• ${b.text}`);
    else h.addParagraph(b.text, { bold: b.bold, align: b.center ? 'center' : undefined });
  }
  return h;
}

async function toDocx(blocks) {
  const children = [];
  const H = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  for (const b of blocks) {
    if (b.type === 'heading') children.push(new Paragraph({ text: b.text, heading: H[b.level], alignment: b.level === 1 ? AlignmentType.CENTER : undefined }));
    else if (b.type === 'pagebreak') children.push(new Paragraph({ children: [new PageBreak()] }));
    else if (b.type === 'field') children.push(new Paragraph({ text: `${b.label ? `${b.label}: ` : ''}{{${b.name}}}` }));
    else if (b.type === 'bullet') children.push(new Paragraph({ text: b.text, bullet: { level: 0 } }));
    else if (b.type === 'table') {
      const cols = Math.max(...b.rows.map((r) => r.length));
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: b.rows.map((r, ri) => new TableRow({
          children: Array.from({ length: cols }, (_, c) => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: r[c] ?? '', bold: ri === 0 })] })],
          })),
        })),
      }));
    } else children.push(new Paragraph({ alignment: b.center ? AlignmentType.CENTER : undefined, children: [new TextRun({ text: b.text, bold: b.bold })] }));
  }
  const doc = new Document({ styles: { default: { document: { run: { font: '맑은 고딕', language: KO_LANG } } } }, sections: [{ children }] });
  return Packer.toBuffer(doc);
}

async function toPptx(blocks) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  const font = '맑은 고딕';
  // 슬라이드 단위로 묶기: # 제목 또는 ---page--- 가 새 슬라이드를 연다
  const slides = [];
  let cur = null;
  for (const b of blocks) {
    if (b.type === 'pagebreak') { cur = null; continue; }
    if ((b.type === 'heading' && b.level === 1) || !cur) { cur = { title: null, items: [] }; slides.push(cur); }
    if (b.type === 'heading' && !cur.title) cur.title = b.text;
    else cur.items.push(b);
  }
  for (const s of slides) {
    const slide = pptx.addSlide();
    let y = 0.4;
    if (s.title) { slide.addText(s.title, { x: 0.6, y, w: 12, h: 1, fontSize: 30, bold: true, fontFace: font }); y += 1.2; }
    const texts = [];
    const flush = () => {
      if (!texts.length) return;
      const h = Math.min(0.45 * texts.length + 0.3, 7.2 - y);
      slide.addText(texts.splice(0), { x: 0.6, y, w: 12, h, fontSize: 18, fontFace: font, valign: 'top' });
      y += h + 0.1;
    };
    for (const b of s.items) {
      if (b.type === 'table') {
        flush();
        const rows = b.rows.map((r, ri) => r.map((t) => ({ text: t, options: { bold: ri === 0, fontFace: font, fontSize: 14 } })));
        slide.addTable(rows, { x: 0.6, y, w: 12, border: { type: 'solid', pt: 1, color: '999999' } });
        y += 0.4 * rows.length + 0.2;
      } else if (b.type === 'heading') {
        texts.push({ text: b.text, options: { bold: true, fontSize: 22, breakLine: true } });
      } else if (b.type === 'bullet') {
        texts.push({ text: b.text, options: { bullet: true, breakLine: true } });
      } else if (b.type === 'field') {
        texts.push({ text: `${b.label ? `${b.label}: ` : ''}{{${b.name}}}`, options: { breakLine: true } });
      } else if (b.text) {
        texts.push({ text: b.text, options: { bold: b.bold, align: b.center ? 'center' : 'left', breakLine: true } });
      }
    }
    flush();
  }
  return pptx.write({ outputType: 'nodebuffer' });
}

async function toXlsx(blocks) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (const b of blocks) {
    if (b.type === 'table') {
      b.rows.forEach((r, ri) => {
        const row = ws.addRow(r.map((v) => (v !== '' && !Number.isNaN(Number(v)) ? Number(v) : v)));
        if (ri === 0) row.font = { bold: true };
      });
      ws.addRow([]);
    } else if (b.type === 'heading') ws.addRow([b.text]).font = { bold: true, size: 14 };
    else if (b.type === 'field') ws.addRow([b.label ?? b.name, `{{${b.name}}}`]);
    else if (b.type !== 'pagebreak' && b.text) ws.addRow([b.text]);
  }
  ws.columns.forEach((col) => { col.width = 16; });
  return wb.xlsx.writeBuffer();
}

/** 원고(마크다운 문자열)로 새 문서 만들기. 형식은 출력 확장자로 결정 */
export async function createDocument(output, markdown) {
  const ext = output.split('.').pop().toLowerCase();
  const blocks = parseMarkdown(markdown);
  if (ext === 'hwp' || ext === 'hwpx') { toHwp(blocks).save(output); return output; }
  const build = { docx: toDocx, pptx: toPptx, xlsx: toXlsx }[ext];
  if (!build) throw new Error(`새로 만들 수 없는 형식: .${ext} (hwp, hwpx, docx, pptx, xlsx 가능)`);
  writeFileSync(output, Buffer.from(await build(blocks)));
  return output;
}

// ── 빈 문서 ──────────────────────────────────────────────

/** 빈 문서로 만들 수 있는 형식 (앞이 기본) */
export const BLANK_FORMATS = Object.freeze(['hwpx', 'hwp', 'docx', 'pptx', 'xlsx']);
// 문서 기본 언어: 한국어 (Word·편집 창의 맞춤법·글자 언어 표시가 '영어(미국)'으로 나오지 않게)
const KO_LANG = { value: 'ko-KR', eastAsia: 'ko-KR' };
const FONT = '맑은 고딕';

// 만든 꾸러미(zip)의 XML 몇 개를 고친다 (라이브러리가 채우는 영어 기본값 정리)
async function patchZip(buf, edits) {
  const zip = await JSZip.loadAsync(buf);
  for (const [name, fn] of Object.entries(edits)) {
    const f = zip.file(name);
    if (f) zip.file(name, fn(await f.async('string')));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function blankBytes(ext, title) {
  if (ext === 'hwp' || ext === 'hwpx') {
    const h = Hwp.create();
    try {
      if (title) h.addHeading(title, 1, { align: 'center' });
      return Buffer.from(h.toBytes(ext));
    } finally { h.close(); }
  }
  if (ext === 'docx') {
    // 지은이는 비워 둔다 (Word가 저장할 때 사용자 이름을 넣는다)
    const children = title
      ? [new Paragraph({ text: title, heading: HeadingLevel.TITLE }), new Paragraph({})]
      : [new Paragraph({})];
    const doc = new Document({
      creator: '', lastModifiedBy: '', title: title || undefined,
      styles: { default: { document: { run: { font: FONT, language: KO_LANG } } } },
      sections: [{ children }],
    });
    return Buffer.from(await Packer.toBuffer(doc));
  }
  if (ext === 'pptx') {
    // 제목 슬라이드 하나: 비어 있는 제목·부제목 개체 틀 (PowerPoint에서 '제목을 입력하세요'로 보인다)
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = ''; pptx.company = ''; pptx.subject = ''; pptx.title = title || '';
    pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
    pptx.defineSlideMaster({
      title: '제목 슬라이드',
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 1, y: 2.1, w: 11.33, h: 1.6, fontSize: 44, align: 'center', valign: 'bottom', fontFace: FONT }, text: '제목을 입력하세요' } },
        { placeholder: { options: { name: 'subtitle', type: 'body', x: 1, y: 3.9, w: 11.33, h: 1.1, fontSize: 22, align: 'center', valign: 'top', fontFace: FONT, color: '595959' }, text: '부제목을 입력하세요' } },
      ],
    });
    // PowerPoint의 '새 슬라이드'가 제목 슬라이드 다음에 고르는 '제목 및 내용' (본문 개체 틀이 있는 레이아웃)
    pptx.defineSlideMaster({
      title: '제목 및 내용',
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 0.92, y: 0.4, w: 11.5, h: 1.45, fontSize: 40, valign: 'middle', fontFace: FONT }, text: '제목을 입력하세요' } },
        { placeholder: { options: { name: 'body', type: 'body', x: 0.92, y: 1.99, w: 11.5, h: 4.76, fontSize: 28, valign: 'top', fontFace: FONT }, text: '텍스트를 입력하세요' } },
      ],
    });
    const slide = pptx.addSlide({ masterName: '제목 슬라이드' });
    slide.addText(title || '', { placeholder: 'title', lang: 'ko-KR' });
    slide.addText('', { placeholder: 'subtitle', lang: 'ko-KR' });
    // PowerPoint의 새 프레젠테이션처럼 고친다: 개체 틀을 가운데 제목·부제목으로, 레이아웃을 '제목 슬라이드'·'빈 화면' 종류로,
    // 개체 이름('Text 0' 둘)을 '제목 1'·'부제목 2'로 (ppt_inspect가 레이아웃을 title로 알아본다)
    const ph = (xml) => xml.replace(/(<p:ph\b[^>]*?\btype=")title"/g, '$1ctrTitle"').replace(/(<p:ph\b[^>]*?\btype=")body"/g, '$1subTitle"');
    const layout = (xml) => {
      if (xml.includes('<p:cSld name="제목 슬라이드"')) return ph(xml).replace(/<p:sldLayout\b(?![^>]*\btype=)/, '<p:sldLayout type="title"');
      if (xml.includes('<p:cSld name="DEFAULT"')) return xml.replace('<p:cSld name="DEFAULT"', '<p:cSld name="빈 화면"').replace(/<p:sldLayout\b(?![^>]*\btype=)/, '<p:sldLayout type="blank"');
      // 본문 개체 틀은 글·표·차트·그림을 넣을 수 있는 '내용'(obj) 개체 틀로
      if (xml.includes('<p:cSld name="제목 및 내용"')) {
        return xml.replace(/(<p:ph\b[^>]*?\btype=")body"/g, '$1obj"').replace(/<p:sldLayout\b(?![^>]*\btype=)/, '<p:sldLayout type="obj"');
      }
      return xml;
    };
    // 레이아웃 갤러리 순서를 PowerPoint 새 프레젠테이션처럼: 제목 슬라이드 → 제목 및 내용 → 빈 화면
    // (pptxgenjs는 빈 화면(DEFAULT)을 늘 맨 앞(rId1)에 둔다)
    const order = (xml) => xml.replace(/<p:sldLayoutIdLst>([\s\S]*?)<\/p:sldLayoutIdLst>/, (all, inner) => {
      const ids = [...inner.matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"[^>]*\/>/g)].map((m) => m[1]);
      if (ids.length !== 3) return all;
      const rids = ['rId2', 'rId3', 'rId1'];
      return `<p:sldLayoutIdLst>${ids.map((id, i) => `<p:sldLayoutId id="${id}" r:id="${rids[i]}"/>`).join('')}</p:sldLayoutIdLst>`;
    });
    const names = ['제목 1', '부제목 2'];
    return patchZip(await pptx.write({ outputType: 'nodebuffer' }), {
      'ppt/slideLayouts/slideLayout1.xml': layout,
      'ppt/slideLayouts/slideLayout2.xml': layout,
      'ppt/slideLayouts/slideLayout3.xml': layout,
      'ppt/slideMasters/slideMaster1.xml': order,
      'ppt/slides/slide1.xml': (xml) => ph(xml).replace(/name="Text 0"/g, (m) => (names.length ? `name="${names.shift()}"` : m)),
    });
  }
  if (ext === 'xlsx') {
    const wb = new ExcelJS.Workbook();
    wb.creator = ''; wb.lastModifiedBy = '';
    if (title) wb.title = title;
    const ws = wb.addWorksheet('Sheet1');
    if (title) { ws.getCell('A1').value = title; ws.getCell('A1').font = { name: FONT, bold: true, size: 14 }; }
    // 지은이가 비면 ExcelJS가 'Unknown'을 넣으므로 뺀다
    return patchZip(Buffer.from(await wb.xlsx.writeBuffer()), {
      'docProps/core.xml': (xml) => xml.replace(/<dc:creator>Unknown<\/dc:creator>|<cp:lastModifiedBy>Unknown<\/cp:lastModifiedBy>/g, ''),
    });
  }
  throw new Error(`빈 문서를 만들 수 없는 형식입니다: .${ext || '(확장자 없음)'} (${BLANK_FORMATS.join(', ')} 가능)`);
}

/**
 * 빈 문서 만들기. 형식은 출력 확장자로 정한다 (hwpx·hwp·docx·pptx·xlsx).
 * title이 있으면 제목으로 넣는다 (한글·Word 첫 문단, PowerPoint 제목 슬라이드, Excel A1).
 * 이미 있는 파일은 overwrite=true일 때만 덮어쓴다 (아니면 code 'EXISTS' 오류).
 */
export async function createBlankDocument(output, { title, overwrite = false } = {}) {
  output = resolve(output);
  const ext = extname(output).slice(1).toLowerCase();
  const t = typeof title === 'string' ? title.trim() : '';
  const bytes = await blankBytes(ext, t);
  if (overwrite) { writeFileSync(output, bytes); return output; }
  try {
    writeFileSync(output, bytes, { flag: 'wx' }); // 없을 때만 만든다
  } catch (err) {
    if (err.code === 'EEXIST') throw Object.assign(new Error(`같은 이름의 파일이 이미 있습니다: ${output}`), { code: 'EXISTS' });
    try { rmSync(output, { force: true }); } catch {} // 쓰다 만 파일
    throw err;
  }
  return output;
}

// ── 새 파일 이름 ─────────────────────────────────────────

export const DEFAULT_DOC_NAME = '새 문서';
const BAD_CHARS = /[<>:"/\\|?*\u0000-\u001f]/g;
// Windows 예약 이름 (확장자가 붙어도 안 됨: con.hwpx)
const RESERVED = /^(con|prn|aux|nul|conin\$|conout\$|com[0-9¹²³]|lpt[0-9¹²³])$/i;
const NAME_MAX = 200;   // 확장자 뺀 이름 길이
const PATH_MAX = 259;   // Office가 열 수 있는 전체 경로 길이

/** 사용자가 쓴 이름 다듬기: 앞뒤 공백을 떼고, 고른 형식의 확장자를 붙여 썼으면 뗀다 */
export function cleanDocName(name, ext) {
  let n = String(name ?? '').trim();
  if (ext && n.toLowerCase().endsWith(`.${String(ext).toLowerCase()}`)) n = n.slice(0, -(String(ext).length + 1)).trim();
  return n;
}

/** 파일 이름(확장자 뺀 것)의 문제 → { code, message } 또는 null */
export function fileNameProblem(name) {
  const n = String(name ?? '');
  if (!n.trim() || /^[.\s]+$/.test(n)) return { code: 'EMPTY', message: '파일 이름을 입력해 주세요' };
  const bad = [...new Set(n.match(BAD_CHARS) ?? [])];
  if (bad.length) {
    const shown = bad.map((c) => (c.charCodeAt(0) < 32 ? '제어 문자' : c));
    return { code: 'BAD_CHARS', message: `파일 이름에 쓸 수 없는 문자가 있어요: ${[...new Set(shown)].join(' ')}` };
  }
  if (n !== n.trim()) return { code: 'SPACES', message: '파일 이름 앞뒤에 공백을 둘 수 없어요' };
  const stem = n.split('.')[0].trimEnd();
  if (RESERVED.test(stem)) return { code: 'RESERVED', message: `Windows가 예약한 이름(${stem.toUpperCase()})이라 쓸 수 없어요` };
  if (n.length > NAME_MAX) return { code: 'TOO_LONG', message: `파일 이름이 너무 길어요 (${NAME_MAX}자 이하)` };
  return null;
}

/** 폴더 안에서 겹치지 않는 이름: '새 문서' → '새 문서 (2)' → '새 문서 (3)' … */
export function freeDocName(folder, ext, base = DEFAULT_DOC_NAME, exists = existsSync) {
  const taken = (n) => exists(join(folder, `${n}.${ext}`));
  if (!taken(base)) return base;
  // '새 문서 (2)'에서 다시 찾으면 '새 문서 (3)'부터
  const m = /^(.*) \((\d+)\)$/.exec(base);
  const stem = m ? m[1] : base;
  for (let i = m ? Number(m[2]) + 1 : 2; i < 10_000; i++) if (!taken(`${stem} (${i})`)) return `${stem} (${i})`;
  return `${stem} (${Date.now()})`;
}

/**
 * 새 문서 만들 자리 확인 → { path, name, ext, folder }. 문제가 있으면 한국어 오류 (code 붙음):
 * BAD_FORMAT · NO_FOLDER · EMPTY · BAD_CHARS · SPACES · RESERVED · TOO_LONG · PATH_TOO_LONG · EXISTS
 */
export function planBlankDocument({ folder, name, format } = {}, { exists = existsSync } = {}) {
  const err = (code, message) => Object.assign(new Error(message), { code });
  const ext = String(format ?? '').toLowerCase();
  if (!BLANK_FORMATS.includes(ext)) throw err('BAD_FORMAT', '지원하지 않는 형식입니다');
  if (typeof folder !== 'string' || !folder.trim() || folder.length > 4096 || !isAbsolute(folder)) throw err('NO_FOLDER', '저장할 폴더를 골라 주세요');
  const dir = resolve(folder);
  let isDir = false;
  try { isDir = statSync(dir).isDirectory(); } catch {}
  if (!isDir) throw err('NO_FOLDER', `저장할 폴더를 찾을 수 없어요: ${dir}`);
  const n = cleanDocName(name, ext);
  const problem = fileNameProblem(n);
  if (problem) throw err(problem.code, problem.message);
  const path = join(dir, `${n}.${ext}`);
  if (path.length > PATH_MAX) throw err('PATH_TOO_LONG', '경로가 너무 길어요. 더 짧은 이름이나 다른 폴더를 골라 주세요');
  if (exists(path)) throw err('EXISTS', `이 폴더에 같은 이름의 파일이 이미 있어요: ${n}.${ext}`);
  return { path, name: n, ext, folder: dir };
}
