// index.mjs — 형식과 설치된 프로그램을 보고 엔진을 골라 문서 작업을 한다.
//   app  엔진: 설치된 한컴/Word/PowerPoint/Excel (원본 그대로, PDF 가능, 느림)
//   file 엔진: rhwp·ZIP/XML·exceljs (프로그램 불필요, 빠름). 한글→PDF는 rhwp 명령줄 프로그램(hwppdf.mjs)
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, renameSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { extname, basename, join, resolve } from 'node:path';
import { Hwp } from './hwp.mjs';
import { readDocx, readPptx, readXlsx, readCsv, replaceInFile, placeholders, setXlsxCells, PLACEHOLDER } from './officefile.mjs';
import { comCall, comAppFor, comAvailableFor, probe } from './com.mjs';
import { createDocument, parseMarkdown } from './create.mjs';
import { hwpToPdf, hwpPdfStatus, HWP_PDF_NOTE } from './hwppdf.mjs';
import { readPdf } from './pdffile.mjs';

export { createDocument, parseMarkdown, probe, Hwp, hwpPdfStatus, HWP_PDF_NOTE };

// 한글 계열 (hwt는 한글 서식 파일, rhwp로 열린다)
const HWP_EXTS = new Set(['hwp', 'hwpx', 'hwt']);
export const isHwp = (ext) => HWP_EXTS.has(ext);
// 프로그램 없이 읽을 수 있는 형식 / 고칠 수 있는 형식
const READ_EXTS = new Set([...HWP_EXTS, 'docx', 'pptx', 'xlsx', 'csv', 'pdf']);
const WRITE_EXTS = new Set([...HWP_EXTS, 'docx', 'pptx', 'xlsx']);
export const OPENABLE_EXTS = ['hwp', 'hwpx', 'hwt', 'docx', 'doc', 'rtf', 'odt', 'pptx', 'ppt', 'odp', 'xlsx', 'xls', 'xlsm', 'csv', 'ods', 'pdf'];

// 프로그램별로 저장(변환)할 수 있는 형식 — office.ps1과 같게
export const SAVE_TARGETS = {
  hwp: ['hwp', 'hwpx', 'hwt', 'pdf', 'docx', 'txt', 'html', 'odt'],
  word: ['docx', 'pdf', 'doc', 'rtf', 'txt', 'html', 'odt'],
  ppt: ['pptx', 'pdf', 'ppt', 'odp'],
  excel: ['xlsx', 'xlsm', 'xls', 'csv', 'ods', 'pdf'],
};
const RHWP_TARGETS = ['hwp', 'hwpx', 'hwt', 'hml'];

export const extOf = (p) => extname(p).slice(1).toLowerCase();

const KIND = { hwp: '한글', hwpx: '한글', hwt: '한글', docx: 'Word', doc: 'Word', rtf: 'Word', odt: 'Word',
  pptx: 'PowerPoint', ppt: 'PowerPoint', odp: 'PowerPoint', xlsx: 'Excel', xls: 'Excel', xlsm: 'Excel', csv: 'Excel', ods: 'Excel', pdf: 'PDF' };
export const kindOf = (p) => KIND[extOf(p)] ?? '알 수 없음';

const HANCOM_NO_MODULE = '한컴오피스는 있지만 자동화 보안 모듈이 등록되지 않아 한글 프로그램으로 작업할 수 없습니다. '
  + '설정에서 "한글 자동화 연결"(명령줄: office-kit setup hancom)을 먼저 해주세요.';

function mustExist(path) {
  if (!existsSync(path)) throw new Error(`파일이 없습니다: ${path}`);
  return resolve(path);
}

const samePath = (a, b) => (process.platform === 'win32' ? resolve(a).toLowerCase() === resolve(b).toLowerCase() : resolve(a) === resolve(b));

// ── 인터넷에서 받은 파일 (Windows '웹에서 받은 파일' 표시, Zone.Identifier) ──────────
// Word·Excel·PowerPoint를 COM으로 열면 보호된 보기를 거치지 않는다 → 인터넷(3)·제한(4) 영역 파일은 자동으로 열지 않는다
/** 파일의 보안 영역: 0 내 컴퓨터 · 1 인트라넷 · 2 신뢰 · 3 인터넷 · 4 제한 (표시가 없으면 0) */
export function securityZone(path) {
  if (process.platform !== 'win32') return 0;
  try {
    const t = readFileSync(`${resolve(path)}:Zone.Identifier`, 'utf8');
    const m = /ZoneId\s*=\s*(\d)/i.exec(t);
    return m ? Number(m[1]) : 0;
  } catch { return 0; }
}
export const isFromInternet = (path) => securityZone(path) >= 3;
export const UNTRUSTED_APP_MESSAGE = '인터넷에서 받은 문서라 안전을 위해 Office 프로그램으로 자동으로 열지 않습니다 (자동으로 열면 보호된 보기를 거치지 않아요). '
  + '믿을 수 있는 문서면 파일 속성에서 \'차단 해제\'를 고른 뒤 다시 하거나, \'프로그램에서 열기\'로 보호된 보기에서 확인하세요.';

// 프로그램 엔진을 못 쓰는 이유
async function noAppReason(path, write) {
  const ext = extOf(path);
  const app = comAppFor(ext);
  if (!app) return `지원하지 않는 형식입니다: .${ext}`;
  const p = await probe();
  if (p.error) return `설치된 프로그램을 확인하지 못했습니다: ${p.error}`;
  if (app === 'hwp' && p.hancom && !p.hancomSecurityModule) return HANCOM_NO_MODULE;
  return `${kindOf(path)} 프로그램이 설치돼 있지 않아 .${ext} 파일을 ${write ? '고칠' : '열'} 수 없습니다`;
}

/** engine: 'auto' | 'app' | 'file'. 쓰기는 프로그램 우선, 읽기는 파일 엔진 우선 */
async function pickEngine(path, { engine = 'auto', write = false } = {}) {
  const ext = extOf(path);
  const fileOk = (write ? WRITE_EXTS : READ_EXTS).has(ext);
  if (engine === 'file') {
    if (!fileOk) throw new Error(`.${ext}는 프로그램 없이 ${write ? '고칠' : '다룰'} 수 없습니다`);
    return 'file';
  }
  // 인터넷에서 받은 파일: 프로그램(COM)으로 열지 않고 파일 엔진으로 (파일 엔진으로 못 하면 거절)
  if (!HWP_EXTS.has(ext) && isFromInternet(path)) {
    if (engine !== 'app' && fileOk) return 'file';
    throw new Error(UNTRUSTED_APP_MESSAGE);
  }
  // 파일 엔진으로 읽을 수 있으면 프로그램 확인(PowerShell)도 하지 않는다
  if (engine !== 'app' && !write && fileOk) return 'file';
  if (await comAvailableFor(ext)) return 'app';
  if (engine !== 'app' && fileOk) return 'file';
  throw new Error(await noAppReason(path, write));
}

// 결과를 원본과 다른 형식으로 저장해도 되는지. 파일 엔진은 같은 형식(한글은 hwp·hwpx·hwt끼리)만
function checkOutput(path, output, eng, { sameKindOnly = false } = {}) {
  const from = extOf(path); const to = extOf(output);
  if (from === to) return;
  if (HWP_EXTS.has(from) && HWP_EXTS.has(to) && (eng !== 'app' || SAVE_TARGETS.hwp.includes(to))) return;
  if (!sameKindOnly && eng === 'app' && (SAVE_TARGETS[comAppFor(from)] ?? []).includes(to)) return;
  throw new Error(`결과를 .${to || '(확장자 없음)'} 파일로 저장할 수 없습니다. 원본과 같은 .${from}로 저장하고, 다른 형식이 필요하면 변환(convert)을 쓰세요`);
}

const sum = (a) => a.reduce((n, x) => n + (Number(x) || 0), 0);
const list = (v) => (v == null ? [] : [].concat(v));

// ── 읽기 ─────────────────────────────────────────────────

// Word 본문: \r 문단 끝, \x07 표 칸 끝, \f 쪽 나누기, \x0b 줄바꿈 (칸 끝 \r\x07 바로 뒤의 \r\x07 은 행 끝)
const cleanWord = (t) => (t ?? '').replace(/\r\x07\r\x07/g, '\n').replace(/\r\x07/g, '\t').replace(/\x07/g, '\t')
  .replace(/\r/g, '\n').replace(/\x0b/g, '\n').replace(/\f/g, '\n');

function sheetText(s) {
  const head = s.truncated
    ? `--- 시트 ${s.name} (잘림: 전체 ${s.totalRows}행×${s.totalCols}열 중 ${s.rows.length}행×${s.rows[0]?.length ?? 0}열만 읽음) ---`
    : `--- 시트 ${s.name} ---`;
  return `${head}\n${s.rows.map((r) => r.join('\t')).join('\n')}`;
}

/** 문서 내용 읽기 → { engine, kind, text, tables?, slides?, sheets?, truncated? } */
export async function readDocument(path, opts = {}) {
  path = mustExist(path);
  const ext = extOf(path);
  const engine = await pickEngine(path, opts);
  if (engine === 'app') {
    const res = await comCall({ app: comAppFor(ext), op: 'text', path });
    const out = { engine: res.engine, kind: kindOf(path), text: res.engine === 'word' ? cleanWord(res.text) : String(res.text ?? '') };
    if (res.engine === 'powerpoint') out.text = out.text.replace(/^--- slide (\d+) ---$/gm, '--- 슬라이드 $1 ---');
    if (res.sheets) {
      out.sheets = list(res.sheets).map((s) => ({
        name: String(s.name), rows: list(s.rows).map((r) => list(r).map((v) => String(v ?? ''))),
        ...(s.truncated ? { truncated: true, totalRows: s.totalRows, totalCols: s.totalCols } : {}),
      }));
      out.text = out.sheets.map(sheetText).join('\n\n');
      if (out.sheets.some((s) => s.truncated)) out.truncated = true;
    }
    return out;
  }
  if (HWP_EXTS.has(ext)) {
    const h = Hwp.open(path);
    try { return { engine: 'rhwp', kind: '한글', text: h.text(), tables: h.tables().map((t) => h.readTable(t)), pages: h.pageCount }; }
    finally { h.close(); }
  }
  if (ext === 'pdf') {
    const r = await readPdf(path);
    return { engine: 'pdf', kind: 'PDF', text: r.text, pages: r.pages, ...(r.scanned ? { scanned: true } : {}), ...(r.truncated ? { truncated: true } : {}) };
  }
  const r = await { docx: readDocx, pptx: readPptx, xlsx: readXlsx, csv: readCsv }[ext](path);
  return { engine: 'file', kind: kindOf(path), ...r };
}

/** 파일 기본 정보 */
export async function documentInfo(path) {
  path = mustExist(path);
  const ext = extOf(path);
  const apps = await probe();
  const info = { path, name: basename(path), kind: kindOf(path), format: ext, size: statSync(path).size,
    appAvailable: await comAvailableFor(ext), fileEngine: READ_EXTS.has(ext), installed: apps };
  if (HWP_EXTS.has(ext)) {
    const h = Hwp.open(path);
    try { Object.assign(info, { pages: h.pageCount, tables: h.tables().length, fields: h.fields().map((f) => f.name) }); }
    finally { h.close(); }
  }
  if (ext === 'pdf') {
    const r = await readPdf(path);
    Object.assign(info, { pages: r.pages, scanned: r.scanned });
  }
  return info;
}

// ── 쓰기 ─────────────────────────────────────────────────

function normalizePairs(pairs) {
  return pairs.map((p) => (Array.isArray(p) ? p : [p.find, p.replace]))
    .filter(([f]) => f).map(([f, r]) => [String(f), String(r ?? '')]);
}

/**
 * 찾아 바꾸기. pairs: [[찾을말, 바꿀말], ...] 또는 [{find, replace}]
 * 영문 대소문자를 구분한다 (모든 형식·엔진 같음). → { engine, count, counts: [쌍마다 바꾼 수], output }
 */
export async function replaceText(path, pairs, { output, engine } = {}) {
  path = mustExist(path);
  output = resolve(output ?? path);
  pairs = normalizePairs(pairs);
  const ext = extOf(path);
  const eng = await pickEngine(path, { engine, write: true });
  checkOutput(path, output, eng);
  if (eng === 'app') {
    const res = await comCall({ app: comAppFor(ext), op: 'replace', path, pairs, save: output, saveFormat: extOf(output) });
    const counts = list(res.counts).map(Number);
    return { engine: res.engine, count: res.count ?? sum(counts), counts, output };
  }
  if (HWP_EXTS.has(ext)) {
    const h = Hwp.open(path);
    try {
      const counts = pairs.map(([f, r]) => h.replace(f, r, { caseSensitive: true }));
      h.save(output);
      return { engine: 'rhwp', count: sum(counts), counts, output };
    } finally { h.close(); }
  }
  const counts = await replaceInFile(path, ext, pairs, output);
  return { engine: 'file', count: sum(counts), counts, output };
}

// {{이름}} 자리표시자 → [{ name, value, kind, raw: [문서에 실제로 쓰인 표기...] }]
async function placeholderList(path, ext, opts = {}) {
  if (['docx', 'pptx', 'xlsx', 'csv'].includes(ext)) return placeholders(path, ext);
  let text = '';
  if (HWP_EXTS.has(ext)) { const h = Hwp.open(path); try { text = h.text(); } finally { h.close(); } }
  else text = (await readDocument(path, opts).catch(() => null))?.text ?? '';
  const found = new Map();
  for (const m of text.matchAll(PLACEHOLDER)) {
    if (!found.has(m[1])) found.set(m[1], new Set());
    found.get(m[1]).add(m[0]);
  }
  return [...found].map(([name, raw]) => ({ name, value: '', kind: 'placeholder', raw: [...raw] }));
}

/** 채울 수 있는 칸: 한글 누름틀, Word 콘텐츠 컨트롤·책갈피, {{자리표시자}} (머리글·바닥글 포함) */
export async function listFields(path, opts = {}) {
  path = mustExist(path);
  const ext = extOf(path);
  let fields = [];
  if (HWP_EXTS.has(ext)) {
    const h = Hwp.open(path);
    try { fields = h.fields().map((f) => ({ name: f.name, value: f.value, kind: '누름틀' })); } finally { h.close(); }
  } else if (comAppFor(ext) === 'word' && opts.engine !== 'file' && (await comAvailableFor(ext))) {
    fields = list((await comCall({ app: 'word', op: 'fields', path })).fields);
  }
  const known = new Set(fields.map((f) => f.name));
  for (const f of await placeholderList(path, ext, opts)) if (!known.has(f.name)) { known.add(f.name); fields.push(f); }
  return fields;
}

/** 칸 채우기: values = { 이름: 값 } → { filled, missing, output } */
export async function fillFields(path, values, { output, engine } = {}) {
  path = mustExist(path);
  output = resolve(output ?? path);
  const ext = extOf(path);
  // 채우기 결과는 원본과 같은 형식으로만 (한글은 hwp·hwpx·hwt끼리)
  checkOutput(path, output, 'file', { sameKindOnly: true });
  const names = Object.keys(values);
  const filled = new Set();
  let current = path;

  if (HWP_EXTS.has(ext)) {
    const own = new Set((await listFields(path)).filter((f) => f.kind === '누름틀').map((f) => f.name));
    const mine = Object.fromEntries(names.filter((n) => own.has(n)).map((n) => [n, String(values[n])]));
    if (Object.keys(mine).length) {
      const eng = await pickEngine(path, { engine, write: true });
      if (eng === 'app') await comCall({ app: 'hwp', op: 'fill', path, values: mine, save: output, saveFormat: extOf(output) });
      else { const h = Hwp.open(path); try { h.fill(mine); h.save(output); } finally { h.close(); } }
      Object.keys(mine).forEach((n) => filled.add(n));
      current = output;
    }
  } else if (comAppFor(ext) === 'word' && engine !== 'file' && (await comAvailableFor(ext))) {
    const res = await comCall({ app: 'word', op: 'fill', path, values, save: output, saveFormat: extOf(output) });
    list(res.filled).forEach((n) => filled.add(n));
    current = output;
  }

  // 남은 이름은 {{이름}} 자리표시자. 문서에 실제로 쓰인 표기({{ 이름 }} 등)를 찾아 그대로 바꾸고,
  // 바뀐 수가 있는 이름만 채움으로 센다
  const rest = names.filter((n) => !filled.has(n));
  if (rest.length) {
    let spell;
    if (READ_EXTS.has(ext)) spell = new Map((await placeholderList(current, ext)).map((f) => [f.name, f.raw]));
    else spell = new Map(rest.map((n) => [n, [`{{${n}}}`, `{{ ${n} }}`]])); // 프로그램으로만 열리는 형식: 흔한 두 표기
    const pairs = []; const owner = [];
    for (const n of rest) for (const raw of spell.get(n) ?? []) { pairs.push([raw, String(values[n])]); owner.push(n); }
    if (pairs.length) {
      const r = await replaceText(current, pairs, { output, engine });
      current = output;
      if (r.counts.length === pairs.length) r.counts.forEach((c, i) => { if (c > 0) filled.add(owner[i]); });
      else if (r.count > 0) owner.forEach((n) => filled.add(n));
    }
  }
  // 아무것도 쓰지 않았어도 output을 지정했으면 그 자리에 결과 파일이 있어야 한다
  if (!samePath(current, output)) {
    if (extOf(current) === extOf(output)) copyFileSync(current, output);
    else await convertDocument(current, output); // hwp ↔ hwpx
  }
  return { filled: names.filter((n) => filled.has(n)), missing: names.filter((n) => !filled.has(n)), output };
}

/** 엑셀 셀 쓰기: cells = { A1: 값, B2: '=SUM(A1:A3)' } */
export async function setCells(path, cells, { sheet, output, engine } = {}) {
  path = mustExist(path);
  output = resolve(output ?? path);
  const ext = extOf(path);
  if (comAppFor(ext) !== 'excel') throw new Error('셀 쓰기는 엑셀 파일에서만 됩니다');
  const eng = await pickEngine(path, { engine, write: true });
  checkOutput(path, output, eng);
  if (eng === 'app') {
    const res = await comCall({ app: 'excel', op: 'setCells', path, sheet, cells, save: output, saveFormat: extOf(output) });
    return { engine: 'excel', updated: res.updated, output };
  }
  return { engine: 'file', updated: await setXlsxCells(path, cells, { sheet, output }), output };
}

/**
 * 다른 형식으로 저장 (pdf 포함). hwp↔hwpx는 프로그램 없이도 된다.
 * 한글 → pdf: 한컴오피스 + 자동화 보안 모듈이 있으면 한컴으로(원본 그대로), 없으면 내장 엔진(rhwp, note에 안내)
 * → { engine, output, pages?, note? }
 */
export async function convertDocument(path, output) {
  path = mustExist(path);
  output = resolve(output);
  const from = extOf(path);
  const to = extOf(output);
  const app = comAppFor(from);
  const appTargets = SAVE_TARGETS[app] ?? [];
  const rhwpOk = HWP_EXTS.has(from) && RHWP_TARGETS.includes(to);
  const rhwpPdf = HWP_EXTS.has(from) && to === 'pdf';
  // 프로그램을 띄우기 전에 형식부터 확인한다
  if (!rhwpOk && !rhwpPdf && !appTargets.includes(to)) {
    const can = [...new Set([...appTargets, ...(HWP_EXTS.has(from) ? RHWP_TARGETS : [])])].filter((t) => t !== from);
    throw new Error(`${kindOf(path)} 문서는 .${to || '(확장자 없음)'} 형식으로 바꿀 수 없습니다${can.length ? ` (가능: ${can.join(', ')})` : ''}`);
  }
  if (appTargets.includes(to) && (await comAvailableFor(from))) {
    await comCall({ app, op: 'convert', path, save: output, saveFormat: to }, { timeoutMs: 180_000 });
    return { engine: app, output };
  }
  if (rhwpOk) {
    const h = Hwp.open(path); try { h.save(output); } finally { h.close(); }
    return { engine: 'rhwp', output };
  }
  if (rhwpPdf) {
    const r = await hwpToPdf(path, output);
    return { engine: 'rhwp', output: r.output, pages: r.pages, note: HWP_PDF_NOTE };
  }
  const apps = await probe();
  if (apps.error) throw new Error(`설치된 프로그램을 확인하지 못해 .${to}로 변환할 수 없습니다: ${apps.error}`);
  if (app === 'hwp' && apps.hancom && !apps.hancomSecurityModule) {
    throw new Error('한컴오피스는 있지만 자동화 보안 모듈이 등록되지 않아 변환할 수 없습니다. 설정에서 "한글 자동화 연결"을 먼저 해주세요.');
  }
  throw new Error(`${kindOf(path)} 프로그램이 없어 .${to}로 변환할 수 없습니다`);
}

/**
 * 미리보기·쪽 그림용 PDF (설치된 프로그램). Excel은 시트마다 한 쪽 너비에 맞춘다 (넓은 표·차트가 여러 쪽으로 잘리지 않게).
 * 쪽 설정은 임시 사본에만 바꾸고 저장하지 않는다 — 원본 파일은 건드리지 않는다.
 * '파일 → PDF로 저장'(convertDocument)은 문서의 인쇄 설정을 그대로 쓴다.
 */
export async function previewPdf(path, output) {
  path = mustExist(path);
  if (comAppFor(extOf(path)) !== 'excel' || !(await comAvailableFor(extOf(path)))) return convertDocument(path, output);
  const work = mkdtempSync(join(tmpdir(), 'officekit-xlpdf-'));
  try {
    const copy = join(work, basename(path));
    copyFileSync(path, copy);
    await comCall({ app: 'excel', op: 'convert', path: copy, save: resolve(output), saveFormat: 'pdf', fitWidth: true }, { timeoutMs: 180_000 });
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
  return { engine: 'excel', output: resolve(output) };
}

/**
 * 미리보기 만들기.
 *   mode 'auto' : 한글은 rhwp SVG(빠름), 나머지는 프로그램 PDF
 *   mode 'exact': 설치된 프로그램으로 PDF (가장 정확)
 * → { kind: 'pdf', file } | { kind: 'svg', pages: [svg...] } | { kind: 'text', docKind, ...readDocument }
 */
export async function renderPreview(path, outDir, { mode = 'auto', maxPages = 40 } = {}) {
  path = mustExist(path);
  const ext = extOf(path);
  // PDF는 그대로 내장 PDF 뷰어로 (사본 — 보는 동안 원본을 잡고 있지 않게)
  if (ext === 'pdf') {
    mkdirSync(outDir, { recursive: true });
    const file = join(outDir, `${basename(path)}.${Date.now()}.pdf`);
    copyFileSync(path, file);
    return { kind: 'pdf', file, engine: 'pdf' };
  }
  const hwp = HWP_EXTS.has(ext);
  // 인터넷에서 받은 Word·Excel·PowerPoint 문서는 프로그램으로 미리 그리지 않는다 (보호된 보기 없이 열리므로)
  const untrusted = !hwp && isFromInternet(path);
  const appOk = !untrusted && await comAvailableFor(ext);
  if (untrusted && !READ_EXTS.has(ext)) {
    return { kind: 'text', engine: 'file', docKind: kindOf(path), text: '', note: '인터넷에서 받은 옛 형식 문서라 안전을 위해 미리보기를 만들지 않았어요. \'프로그램에서 열기\'를 누르면 보호된 보기로 열려요.' };
  }
  if (appOk && (mode === 'exact' || !hwp)) {
    mkdirSync(outDir, { recursive: true });
    const file = join(outDir, `${basename(path)}.${Date.now()}.pdf`);
    // 프로그램(COM)은 따로 떠서, MSIX 앱(스토어판 Claude Desktop 등) 안에서 실행될 때 가상화된
    // %APPDATA% 폴더(outDir)를 보지 못한다 → 임시 폴더에 PDF를 받은 뒤 옮긴다
    const work = mkdtempSync(join(tmpdir(), 'officekit-preview-'));
    try {
      const pdf = join(work, 'preview.pdf');
      await previewPdf(path, pdf);
      try { renameSync(pdf, file); } catch { copyFileSync(pdf, file); }
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
    return { kind: 'pdf', file, engine: comAppFor(ext) };
  }
  if (hwp) {
    const h = Hwp.openForView(path);
    try {
      const n = Math.min(h.pageCount, maxPages);
      return { kind: 'svg', engine: 'rhwp', total: h.pageCount, pages: Array.from({ length: n }, (_, i) => h.svg(i)) };
    } finally { h.close(); }
  }
  const doc = await readDocument(path);
  return { ...doc, kind: 'text', docKind: doc.kind, ...(untrusted ? { note: '인터넷에서 받은 문서라 안전을 위해 글자 미리보기로 보여요 (Office 프로그램으로 그리지 않음).' } : {}) };
}

/** 설치된 프로그램 창으로 열기 (사람이 직접 확인·편집) */
export async function openInApp(path) {
  path = mustExist(path);
  const ext = extOf(path);
  if (!(await comAvailableFor(ext, { interactive: true }))) throw new Error(`${kindOf(path)} 프로그램이 설치돼 있지 않습니다`);
  // 사람이 고치도록 읽기·쓰기로 연다. 쓰기 암호가 있거나 다른 프로그램이 쥐고 있으면 읽기 전용 (readOnlyReason)
  const r = await comCall({ app: comAppFor(ext), op: 'open', path, visible: true });
  return { opened: path, readOnly: !!r?.readOnly, ...(r?.readOnly ? { readOnlyReason: r.readOnlyReason === 'write password' ? 'writePassword' : 'inUse' } : {}) };
}
