// render.mjs — 문서 한 쪽을 PNG로 그린다 (AI가 결과를 눈으로 확인하는 용도).
//   hwp/hwpx/hwt: rhwp SVG → resvg (시스템 글꼴, 프로그램 불필요)
//   pdf     : Windows.Data.Pdf (office.ps1 app:'pdf'), 프로그램 불필요
//   PowerPoint: 슬라이드를 바로 그림으로 (office.ps1 'slideImage' = Slide.Export). PDF는 숨긴 슬라이드를 빼서
//              쪽 번호가 슬라이드 번호와 어긋나므로 거치지 않는다 → 쪽 = 슬라이드 번호(숨긴 슬라이드 포함)
//   그 외   : 설치된 프로그램으로 PDF 변환 → 위 PDF 경로
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Resvg } from '@resvg/resvg-js';
import JSZip from 'jszip';
import { Hwp } from './hwp.mjs';
import { comCall, comAppFor, comAvailableFor } from './com.mjs';
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

// ── PowerPoint 슬라이드 ──────────────────────────────────

// pptx·odp 꾸러미에서 슬라이드 수(숨긴 슬라이드 포함)와 너비(pt)를 읽는다 → PowerPoint를 띄우기 전에 범위 확인.
// .ppt(복합 문서)나 못 읽는 파일은 null (그때는 PowerPoint가 알려 준 수를 쓴다)
async function pptFacts(path) {
  try {
    const zip = await JSZip.loadAsync(readFileSync(path));
    const rels = await zip.file('_rels/.rels')?.async('string');
    const main = rels?.match(/<Relationship\b[^>]*Type="[^"]*\/officeDocument"[^>]*>/)?.[0].match(/Target="([^"]+)"/)?.[1];
    const pres = await zip.file(main ? main.replace(/^\//, '') : 'ppt/presentation.xml')?.async('string');
    if (pres) {
      const cx = Number(pres.match(/<(?:\w+:)?sldSz\b[^>]*\bcx="(\d+)"/)?.[1]);
      return { count: (pres.match(/<(?:\w+:)?sldId\b/g) ?? []).length, widthPt: cx > 0 ? cx / 12700 : null };
    }
    const content = await zip.file('content.xml')?.async('string'); // odp
    if (content) return { count: (content.match(/<draw:page[\s>]/g) ?? []).length, widthPt: null };
  } catch { /* 꾸러미가 아님 */ }
  return null;
}

// 같은 파일·슬라이드를 다시 그릴 때 PowerPoint를 또 띄우지 않도록 원래 크기 PNG를 기억한다
const slideCache = new Map();
const SLIDE_CACHE_MAX = 24;

// Slide.Export는 크기를 받지 않아(office.ps1 'slideImage') 기본 크기(보통 96dpi, 960pt 너비 → 1280px)로 나온다.
// 더 작은 dpi는 여기서 줄인다. 더 큰 dpi는 늘려도 글자가 선명해지지 않으므로 기본 크기 그대로 둔다.
function shrink(png, width, height, targetWidth) {
  if (!(targetWidth > 0) || targetWidth >= width - 1) return { png, width, height };
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${width}" height="${height}" `
    + `viewBox="0 0 ${width} ${height}"><image width="${width}" height="${height}" xlink:href="data:image/png;base64,${png.toString('base64')}"/></svg>`;
  const img = new Resvg(svg, { background: 'white', fitTo: { mode: 'width', value: Math.max(1, Math.round(targetWidth)) } }).render();
  return { png: img.asPng(), width: img.width, height: img.height };
}

async function slideImage(path, page, dpi) {
  const ext = extOf(path);
  const { pptProtection, LOCKED_OPEN, LOCKED_PPT } = await import('./ppt.mjs');
  // 암호 문서는 PowerPoint가 보이지 않는 암호 창에서 멈추므로 열기 전에 거른다 (쓰기 암호 pptx는 읽기 전용으로 열려서 됨)
  const lock = await pptProtection(path);
  if (lock === 'open') throw new Error(LOCKED_OPEN);
  if (lock === 'ppt') throw new Error(LOCKED_PPT);
  const facts = await pptFacts(path);
  if (facts) checkPage(page, facts.count);
  else if (!Number.isInteger(page) || page < 1) throw new Error(`쪽 번호가 범위를 벗어났습니다: ${page}쪽 (슬라이드 번호는 1부터 시작)`);
  if (!(await comAvailableFor(ext))) throw new Error(`${kindOf(path)} 문서를 그림으로 보려면 PowerPoint가 설치돼 있어야 합니다`);

  const st = statSync(path);
  const key = `${path}|${st.mtimeMs}|${st.size}|${page}`;
  let hit = slideCache.get(key);
  if (!hit) {
    const out = join(scratch(), `s${page}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`);
    try {
      let res;
      try {
        res = await comCall({ app: 'ppt', op: 'slideImage', path, page, output: out }, { timeoutMs: 180_000 });
      } catch (e) {
        // .ppt처럼 미리 못 센 문서: PowerPoint의 범위 오류 → 같은 한국어 문구
        const total = String(e.message).match(/valid range of 1 to (\d+)/i)?.[1];
        if (total) checkPage(page, Number(total));
        if (/Slides\.Item|out of range/i.test(e.message)) throw new Error(`쪽 번호가 범위를 벗어났습니다: ${page}쪽 (슬라이드 번호는 1부터 시작)`);
        throw e;
      }
      if (!existsSync(out)) throw new Error('슬라이드를 그림으로 바꾸지 못했습니다');
      const png = readFileSync(out);
      hit = { png, pages: Number(res.slides), width: png.readUInt32BE(16), height: png.readUInt32BE(20) };
    } finally {
      rmSync(out, { force: true });
    }
    slideCache.set(key, hit);
    if (slideCache.size > SLIDE_CACHE_MAX) slideCache.delete(slideCache.keys().next().value);
  }
  checkPage(page, hit.pages);
  // 원하는 너비: 슬라이드 너비(pt)를 알면 그것으로, 모르면 기본 크기를 96dpi로 보고 계산
  const target = facts?.widthPt ? (facts.widthPt * dpi) / 72 : (hit.width * dpi) / 96;
  return { ...shrink(hit.png, hit.width, hit.height, target), pages: hit.pages };
}

/**
 * 문서 한 쪽을 PNG로 → { png: Buffer, page, pages, engine, width, height }
 * page는 1부터. dpi 96 = 원래 크기.
 * PowerPoint는 쪽 = 슬라이드 번호(숨긴 슬라이드 포함), pages = 슬라이드 수. 그림은 PowerPoint 기본 크기까지만 (dpi가 더 크면 기본 크기).
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

  if (comAppFor(ext) === 'ppt') return { ...(await slideImage(path, page, dpi)), page, engine: 'powerpoint' };

  let pdf;
  try { pdf = await pdfFor(path); } catch (e) {
    throw new Error(`${kindOf(path)} 문서를 그림으로 보려면 PDF 변환이 필요합니다: ${e.message}`);
  }
  return { ...(await renderPdf(pdf.file, page, dpi)), page, engine: `${pdf.engine}+pdf` };
}
