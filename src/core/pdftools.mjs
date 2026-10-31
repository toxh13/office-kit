// pdftools.mjs — PDF 합치기 · 쪽 뽑기 · 쪽 지우기 · 돌리기 (pdf-lib, MIT — 인터넷·프로그램 없이).
//   원본은 바꾸지 않고 output에 새로 쓴다. 쪽 번호는 1부터, "1-3,5,8-" 같은 꼴 ("8-"은 8쪽부터 끝까지).
import { readFileSync, writeFileSync, renameSync, rmSync } from 'node:fs';
import { resolve } from 'node:path';
import { PDFDocument, degrees as deg } from 'pdf-lib';

const MAX_INPUTS = 200;

function pdfError(path, e) {
  const msg = String(e?.message ?? e);
  if (e?.name === 'EncryptedPDFError' || /encrypt/i.test(msg)) return new Error(`암호가 걸린 PDF는 다룰 수 없습니다: ${path}`);
  if (/No PDF header|Failed to parse|Expected/i.test(msg)) return new Error(`PDF 파일이 아니거나 손상됐습니다: ${path}`);
  return new Error(`PDF를 읽지 못했습니다 (${path}): ${msg}`);
}

async function load(path) {
  let bytes;
  try { bytes = readFileSync(path); } catch { throw new Error(`파일을 찾을 수 없습니다: ${path}`); }
  try { return await PDFDocument.load(bytes, { updateMetadata: false }); } catch (e) { throw pdfError(path, e); }
}

async function save(doc, output) {
  const bytes = await doc.save({ useObjectStreams: true });
  // 다 쓴 뒤 이름을 바꿔 끼운다 (중간에 끊겨도 있던 파일이 반쯤 망가지지 않게)
  const part = `${output}.${process.pid}.${Date.now()}.part`;
  try { writeFileSync(part, bytes); renameSync(part, output); } catch (e) { try { rmSync(part, { force: true }); } catch {} throw e; }
  return { output, pages: doc.getPageCount(), bytes: bytes.length };
}

/**
 * 쪽 지정 → 0부터 센 쪽 번호 배열 (적은 순서 그대로, 같은 쪽 여러 번 가능). spec 비면 모든 쪽
 * "1-3,5" / "2" / "7-" / "-3"(1~3) / "all"
 */
export function parsePages(spec, total) {
  const s = String(spec ?? '').replace(/\s+/g, '').replace(/쪽/g, '');
  if (!s || /^(all|전체|모두)$/i.test(s)) return Array.from({ length: total }, (_, i) => i);
  const out = [];
  for (const part of s.split(',')) {
    if (!part) continue;
    const m = /^(\d*)-(\d*)$/.exec(part) ?? /^(\d+)$/.exec(part);
    if (!m) throw new Error(`쪽 번호를 알아볼 수 없습니다: "${part}" (예: 1-3,5,8-)`);
    const range = m.length === 3;
    const a = range ? (m[1] ? Number(m[1]) : 1) : Number(m[1]);
    const b = range ? (m[2] ? Number(m[2]) : total) : a;
    if (a < 1 || b < 1 || a > total || b > total) throw new Error(`쪽 번호가 범위를 벗어났습니다: "${part}" (1~${total}쪽)`);
    const step = a <= b ? 1 : -1;
    for (let p = a; step > 0 ? p <= b : p >= b; p += step) out.push(p - 1);
  }
  if (!out.length) throw new Error('쪽을 하나 이상 골라 주세요');
  return out;
}

const sameAsInput = (output, inputs) => inputs.some((p) => resolve(p).toLowerCase() === resolve(output).toLowerCase());
function checkOutput(output, inputs) {
  if (!output || !/\.pdf$/i.test(output)) throw new Error('결과는 .pdf 파일로 저장합니다');
  if (sameAsInput(output, inputs)) throw new Error('원본과 다른 이름으로 저장하세요 (원본은 그대로 둡니다)');
}

/** 합치기: inputs 순서대로 → { output, pages, bytes } */
export async function mergePdfs(inputs, output) {
  inputs = [].concat(inputs ?? []).map((p) => resolve(p));
  if (inputs.length < 2) throw new Error('합칠 PDF를 둘 이상 골라 주세요');
  if (inputs.length > MAX_INPUTS) throw new Error(`한 번에 ${MAX_INPUTS}개까지 합칠 수 있습니다`);
  output = resolve(output);
  checkOutput(output, inputs);
  const out = await PDFDocument.create();
  for (const p of inputs) {
    const src = await load(p);
    for (const page of await out.copyPages(src, src.getPageIndices())) out.addPage(page);
  }
  return save(out, output);
}

/** 쪽 뽑기 (나누기): 고른 쪽만 새 PDF로 */
export async function extractPages(input, spec, output) {
  input = resolve(input); output = resolve(output);
  checkOutput(output, [input]);
  const src = await load(input);
  const idx = parsePages(spec, src.getPageCount());
  const out = await PDFDocument.create();
  for (const page of await out.copyPages(src, idx)) out.addPage(page);
  return save(out, output);
}

/** 쪽 지우기: 고른 쪽을 뺀 나머지 */
export async function deletePages(input, spec, output) {
  input = resolve(input); output = resolve(output);
  checkOutput(output, [input]);
  const src = await load(input);
  const total = src.getPageCount();
  const drop = new Set(parsePages(spec, total));
  if (drop.size >= total) throw new Error('모든 쪽을 지울 수는 없습니다');
  const keep = src.getPageIndices().filter((i) => !drop.has(i));
  const out = await PDFDocument.create();
  for (const page of await out.copyPages(src, keep)) out.addPage(page);
  return save(out, output);
}

/** 돌리기: 고른 쪽(비면 모든 쪽)을 90·180·270도 (시계 방향) */
export async function rotatePages(input, degrees, spec, output) {
  input = resolve(input); output = resolve(output);
  checkOutput(output, [input]);
  const d = Number(degrees);
  if (![90, 180, 270, -90].includes(d)) throw new Error('90, 180, 270도만 돌릴 수 있습니다');
  const doc = await load(input);
  const idx = new Set(parsePages(spec, doc.getPageCount()));
  doc.getPages().forEach((page, i) => {
    if (!idx.has(i)) return;
    page.setRotation(deg((((page.getRotation().angle + d) % 360) + 360) % 360));
  });
  return save(doc, output);
}

/** 쪽 수 */
export async function pdfPageCount(path) {
  return (await load(resolve(path))).getPageCount();
}
