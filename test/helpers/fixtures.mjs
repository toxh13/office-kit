// fixtures.mjs — 테스트용 원고·임시 폴더·PDF·PNG 도우미
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..');

/** COM(Word/PowerPoint/Excel) 테스트는 OFFICEKIT_COM_TESTS=1 일 때만. 한컴은 쓰지 않는다 */
export const COM_TESTS = process.env.OFFICEKIT_COM_TESTS === '1' && process.platform === 'win32';

export const MANUSCRIPT = `# 분기 보고
>> 2026년 3분기
## 요약
- 매출 증가
- 비용 절감
| 항목 | 값 |
|---|---|
| 매출 | 120 |
| 비용 | 80 |
담당자: {{담당자}}
---page---
# 다음 계획
**신규 고객 확보**`;

export function tempDir(prefix = 'officekit-test-') {
  const dir = mkdtempSync(join(tmpdir(), prefix));
  return { dir, path: (name) => join(dir, name), cleanup: () => rmSync(dir, { recursive: true, force: true }) };
}

/** 글꼴 Helvetica + 파란 사각형이 있는 최소 PDF (pages쪽, A4 pt 단위) */
export function makePdf(pages = 2, { width = 595, height = 842 } = {}) {
  const bodies = {};
  const kids = [];
  let next = 4;
  for (let i = 0; i < pages; i++) {
    const pageId = next++; const contentId = next++;
    kids.push(pageId);
    const stream = `BT /F1 36 Tf 72 ${height - 120} Td (Page ${i + 1}) Tj ET\n0 0 1 rg 72 ${height - 320} 200 100 re f\n`;
    bodies[pageId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentId} 0 R >>`;
    bodies[contentId] = `<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}endstream`;
  }
  bodies[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  bodies[2] = `<< /Type /Pages /Kids [${kids.map((k) => `${k} 0 R`).join(' ')}] /Count ${pages} >>`;
  bodies[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
  let out = '%PDF-1.4\n';
  const offsets = [];
  for (let id = 1; id < next; id++) {
    offsets[id] = Buffer.byteLength(out, 'latin1');
    out += `${id} 0 obj\n${bodies[id]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${next}\n0000000000 65535 f \n${offsets.slice(1).map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
  out += `trailer\n<< /Size ${next} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

const PNG_SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
export const isPng = (buf) => Buffer.isBuffer(buf) && buf.length > 100 && buf.subarray(0, 8).equals(PNG_SIG);
export const pngSize = (buf) => ({ width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) });

/** 차트가 붙은 xlsx (시트 '매출', B4에 {{담당자}}) — 파일 엔진 쓰기가 차트·그림 부품을 지우지 않는지 확인용 */
export async function xlsxWithChart() {
  const { default: ExcelJS } = await import('exceljs');
  const { default: JSZip } = await import('jszip');
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('매출').addRows([['월', '매출'], ['1월', 10], ['2월', 20], ['담당', '{{담당자}}']]);
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  zip.file('xl/charts/chart1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>'매출'!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'매출'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`);
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/chart" Target="../charts/chart1.xml"/></Relationships>`);
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('</worksheet>', `<drawing xmlns:r="${R}" r:id="rIdDr1"/></worksheet>`));
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const rels = zip.file(relsPath) ? await zip.file(relsPath).async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  zip.file(relsPath, rels.replace('</Relationships>', `<Relationship Id="rIdDr1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    + '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** zip 안의 파일 이름 목록 */
export async function zipParts(path) {
  const { default: JSZip } = await import('jszip');
  const { readFileSync } = await import('node:fs');
  return Object.keys((await JSZip.loadAsync(readFileSync(path))).files);
}

/** zip 안 파일 하나의 글자 */
export async function zipText(path, name) {
  const { default: JSZip } = await import('jszip');
  const { readFileSync } = await import('node:fs');
  return (await JSZip.loadAsync(readFileSync(path))).file(name).async('string');
}
