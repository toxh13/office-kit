// create.mjs — 마크다운 원고 하나로 hwp/hwpx/docx/pptx/xlsx 새 문서를 만든다.
import { writeFileSync } from 'node:fs';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, PageBreak,
} from 'docx';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
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
  const doc = new Document({ styles: { default: { document: { run: { font: '맑은 고딕' } } } }, sections: [{ children }] });
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
