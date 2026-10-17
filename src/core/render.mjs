// render.mjs — 문서 한 쪽을 PNG로 그린다 (AI가 결과를 눈으로 확인하는 용도).
//   hwp/hwpx/hwt: rhwp SVG → resvg (시스템 글꼴, 프로그램 불필요)
//   pdf     : Windows.Data.Pdf (office.ps1 app:'pdf'), 프로그램 불필요
//   그 외   : 설치된 프로그램으로 PDF 변환 → 위 PDF 경로
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Resvg } from '@resvg/resvg-js';
import { Hwp } from './hwp.mjs';
import { comCall } from './com.mjs';
import { convertDocument, extOf, kindOf, isHwp } from './index.mjs';

const FONT = { loadSystemFonts: true, defaultFontFamily: '맑은 고딕', sansSerifFamily: '맑은 고딕' };

function checkPage(page, pages) {
  if (!Number.isInteger(page) || page < 1 || page > pages) {
    throw new Error(`쪽 번호가 범위를 벗어났습니다: ${page}쪽 (전체 ${pages}쪽, 1부터 시작)`);
  }
}

// 변환한 PDF는 원본이 바뀌기 전까지 재사용한다 (Word/PowerPoint 변환이 느리므로)
let workDir;
const pdfCache = new Map();
function scratch() {
  if (!workDir) {
    workDir = mkdtempSync(join(tmpdir(), 'officekit-render-'));
    process.once('exit', () => rmSync(workDir, { recursive: true, force: true }));
  }
  return workDir;
}

async function pdfFor(path) {
  const st = statSync(path);
  const key = `${path}|${st.mtimeMs}|${st.size}`;
  const hit = pdfCache.get(key);
  if (hit && existsSync(hit.file)) return hit;
  const file = join(scratch(), `${pdfCache.size + 1}-${Date.now()}.pdf`);
  const res = await convertDocument(path, file);
  const entry = { file, engine: res.engine };
  pdfCache.set(key, entry);
  if (pdfCache.size > 8) {
    const [oldKey, old] = pdfCache.entries().next().value;
    pdfCache.delete(oldKey);
    rmSync(old.file, { force: true });
  }
  return entry;
}

async function renderPdf(pdf, page, dpi) {
  if (process.platform !== 'win32') throw new Error('PDF 쪽 그림은 Windows에서만 만들 수 있습니다');
  const out = join(scratch(), `p${page}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`);
  try {
    const res = await comCall({ app: 'pdf', op: 'pageImage', path: pdf, page, dpi, output: out }, { timeoutMs: 60_000 });
    checkPage(page, res.pages);
    if (!res.output || !existsSync(out)) throw new Error('PDF 쪽을 그림으로 바꾸지 못했습니다');
    return { png: readFileSync(out), pages: res.pages, width: res.width, height: res.height };
  } finally {
    rmSync(out, { force: true });
  }
}

/**
 * 문서 한 쪽을 PNG로 → { png: Buffer, page, pages, engine, width, height }
 * page는 1부터. dpi 96 = 원래 크기.
 */
export async function pageImage(path, page = 1, { dpi = 110 } = {}) {
  if (!existsSync(path)) throw new Error(`파일이 없습니다: ${path}`);
  path = resolve(path);
  page = Number(page);
  dpi = Math.min(Math.max(Number(dpi) || 110, 24), 600);
  const ext = extOf(path);

  if (isHwp(ext)) {
    const h = Hwp.openForView(path);
    let svg, pages;
    try {
      pages = h.pageCount;
      checkPage(page, pages);
      svg = h.svg(page - 1);
    } finally { h.close(); }
    const img = new Resvg(svg, { font: FONT, background: 'white', fitTo: { mode: 'zoom', value: dpi / 96 } }).render();
    return { png: img.asPng(), page, pages, engine: 'rhwp', width: img.width, height: img.height };
  }

  if (ext === 'pdf') return { ...(await renderPdf(path, page, dpi)), page, engine: 'pdf' };

  let pdf;
  try { pdf = await pdfFor(path); } catch (e) {
    throw new Error(`${kindOf(path)} 문서를 그림으로 보려면 PDF 변환이 필요합니다: ${e.message}`);
  }
  return { ...(await renderPdf(pdf.file, page, dpi)), page, engine: `${pdf.engine}+pdf` };
}
