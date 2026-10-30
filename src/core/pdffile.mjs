// pdffile.mjs — PDF 글자 읽기 (unpdf: pdf.js를 묶은 MIT 꾸러미, Node 안에서 동작 — 인터넷·프로그램 불필요).
//   쪽마다 "--- 쪽 n ---" 머리를 붙여 한 글로 돌려준다. 그림으로만 된 쪽(스캔본)은 글자가 없다 → scanned로 알린다.
//   쪽 그림은 render.mjs(Windows.Data.Pdf), PDF를 Word로 바꾸기는 index.mjs convertDocument(Word COM).
import { readFileSync } from 'node:fs';

let loaded = null;
const unpdf = async () => (loaded ??= await import('unpdf'));

/** PDF 오류 → 한국어 */
function pdfError(path, e) {
  const msg = String(e?.message ?? e);
  if (e?.name === 'PasswordException' || /password/i.test(msg)) return Object.assign(new Error('암호가 걸린 PDF라 글자를 읽을 수 없습니다'), { code: 'ENCRYPTED' });
  if (e?.name === 'InvalidPDFException' || /Invalid PDF|PDF header/i.test(msg)) return Object.assign(new Error(`PDF 파일이 아니거나 손상됐습니다: ${path}`), { code: 'BAD_FORMAT' });
  return new Error(`PDF를 읽지 못했습니다: ${msg}`);
}

/**
 * PDF 글자 → { text, pages, pageTexts: [쪽마다 글], scanned }
 * maxPages: 이만큼만 읽는다 (아주 긴 PDF에서 시간을 아끼려고. 기본 2000쪽)
 */
export async function readPdf(path, { maxPages = 2000 } = {}) {
  const { getDocumentProxy, extractText } = await unpdf();
  let doc;
  try { doc = await getDocumentProxy(new Uint8Array(readFileSync(path))); } catch (e) { throw pdfError(path, e); }
  try {
    const pages = doc.numPages;
    let pageTexts;
    if (pages <= maxPages) {
      const r = await extractText(doc, { mergePages: false });
      pageTexts = r.text.map((t) => String(t ?? ''));
    } else {
      pageTexts = [];
      for (let i = 1; i <= maxPages; i++) {
        const page = await doc.getPage(i);
        const c = await page.getTextContent();
        pageTexts.push(c.items.map((it) => `${it.str ?? ''}${it.hasEOL ? '\n' : ''}`).join(''));
      }
    }
    pageTexts = pageTexts.map((t) => t.replace(/\r\n?/g, '\n').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim());
    const chars = pageTexts.reduce((n, t) => n + t.replace(/\s/g, '').length, 0);
    const text = pageTexts.map((t, i) => `--- 쪽 ${i + 1} ---\n${t}`).join('\n\n');
    return { text, pages, pageTexts, scanned: pages > 0 && chars < pages * 5, ...(pages > maxPages ? { truncated: true } : {}) };
  } catch (e) {
    throw e?.code ? e : pdfError(path, e);
  } finally {
    try { await doc.destroy(); } catch {}
  }
}

/** 쪽 수만 */
export async function pdfPages(path) {
  const { getDocumentProxy } = await unpdf();
  let doc;
  try { doc = await getDocumentProxy(new Uint8Array(readFileSync(path))); } catch (e) { throw pdfError(path, e); }
  try { return doc.numPages; } finally { try { await doc.destroy(); } catch {} }
}
