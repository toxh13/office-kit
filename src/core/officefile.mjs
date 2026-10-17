// officefile.mjs — 프로그램 없이 docx/pptx/xlsx/csv를 읽고 쓰는 파일 엔진.
// 쓰기는 ZIP 안의 XML을 직접 고친다 (차트·그림·매크로 등 모르는 부분은 그대로 남긴다).
// xlsx 읽기는 exceljs를 쓴다.
import { readFileSync, writeFileSync } from 'node:fs';
import { basename, extname } from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';

const ENT = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
const decode = (s) => s.replace(/&(amp|lt|gt|quot|apos);|&#(\d+);|&#x([0-9a-f]+);/gi,
  (m, _n, d, h) => ENT[m] ?? String.fromCodePoint(d ? Number(d) : parseInt(h, 16)));
const encode = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const LABEL = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' };
function damaged(path, cause) {
  const label = LABEL[extname(path).slice(1).toLowerCase()];
  return new Error(`문서 파일이 손상되었거나 ${label ? `${label} ` : ''}문서가 아닙니다: ${basename(path)}`, { cause });
}

async function loadZip(path) {
  const buf = readFileSync(path);
  try { return await JSZip.loadAsync(buf); } catch (e) { throw damaged(path, e); }
}
async function partText(zip, name, path) {
  const f = zip.file(name);
  if (!f) throw damaged(path);
  return f.async('string');
}
const saveZip = async (zip, output) => writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

function numbered(zip, re) {
  return Object.keys(zip.files).filter((f) => re.test(f))
    .sort((a, b) => Number(a.match(/(\d+)\.xml$/)?.[1] ?? 0) - Number(b.match(/(\d+)\.xml$/)?.[1] ?? 0));
}

// 같은 태그가 안에 또 있어도(표 속 표, 글상자 속 문단) 짝을 맞춰 가장 바깥 요소만 [시작, 끝)으로
function outer(xml, tag) {
  const re = new RegExp(`<(/?)${tag}(?=[\\s>/])[^>]*>`, 'g');
  const out = [];
  let depth = 0; let start = 0;
  for (const m of xml.matchAll(re)) {
    if (!m[1] && m[0].endsWith('/>')) { if (!depth) out.push([m.index, m.index + m[0].length]); }
    else if (m[1]) { if (depth && --depth === 0) out.push([start, m.index + m[0].length]); }
    else if (depth++ === 0) start = m.index;
  }
  return out;
}
const inner = (el) => (el.endsWith('/>') ? '' : el.slice(el.indexOf('>') + 1, el.lastIndexOf('</')));

// 문단 하나의 글자 (w:t / a:t, 탭·줄바꿈 포함)
function paragraphText(xml, ns) {
  let out = '';
  const re = new RegExp(`<${ns}:t(?:\\s[^>]*)?>([\\s\\S]*?)</${ns}:t>|<${ns}:tab/>|<${ns}:(?:br|cr)(?=[\\s>/])[^>]*>`, 'g');
  for (const m of xml.matchAll(re)) out += m[1] !== undefined ? decode(m[1]) : m[0].includes('tab') ? '\t' : '\n';
  return out;
}
const paragraphs = (xml, ns) => outer(xml, `${ns}:p`).map(([s, e]) => paragraphText(xml.slice(s, e), ns));

// ── 읽기 ─────────────────────────────────────────────────

// 표 → 행 → 칸. 칸 안의 표는 "a | b" 줄로 그 칸 글자에 들어간다
function tableRows(tbl) {
  const body = inner(tbl);
  return outer(body, 'w:tr').map(([s, e]) => {
    const row = inner(body.slice(s, e));
    return outer(row, 'w:tc').map(([a, b]) => blockText(inner(row.slice(a, b))));
  });
}

// 문단과 표를 문서 순서대로 글자로
function blockText(xml) {
  const out = [];
  let pos = 0;
  for (const [s, e] of outer(xml, 'w:tbl')) {
    out.push(...paragraphs(xml.slice(pos, s), 'w'));
    out.push(tableRows(xml.slice(s, e)).map((r) => r.join(' | ')).join('\n'));
    pos = e;
  }
  out.push(...paragraphs(xml.slice(pos), 'w'));
  return out.join('\n');
}

// 글상자는 새 형식(mc:Choice)과 옛 형식(mc:Fallback)이 같은 글을 두 번 담고 있어 읽을 때 하나는 뺀다
const dropFallback = (xml) => xml.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, '');

export async function readDocx(path) {
  const zip = await loadZip(path);
  const xml = dropFallback(await partText(zip, 'word/document.xml', path));
  const tables = outer(xml, 'w:tbl').map(([s, e]) => tableRows(xml.slice(s, e)));
  return { text: blockText(xml), tables };
}

export async function readPptx(path) {
  const zip = await loadZip(path);
  const slides = [];
  for (const f of numbered(zip, /^ppt\/slides\/slide\d+\.xml$/)) {
    const xml = dropFallback(await zip.file(f).async('string'));
    slides.push(paragraphs(xml, 'a').filter(Boolean).join('\n'));
  }
  return { text: slides.map((s, n) => `--- 슬라이드 ${n + 1} ---\n${s}`).join('\n\n'), slides };
}

function cellText(cell) {
  const v = cell.value;
  if (v && typeof v === 'object' && 'formula' in v) return v.result === undefined ? `=${v.formula}` : String(v.result);
  return cell.text ?? '';
}

const sheetsText = (sheets) => sheets.map((s) => `--- 시트 ${s.name} ---\n${s.rows.map((r) => r.join('\t')).join('\n')}`).join('\n\n');

export async function readXlsx(path) {
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(readFileSync(path)); } catch (e) { throw damaged(path, e); }
  const sheets = wb.worksheets.map((ws) => {
    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) vals.push(cellText(row.getCell(c)));
      rows[r - 1] = vals;
    });
    return { name: ws.name, rows: Array.from(rows, (r) => r ?? []) };
  });
  return { text: sheetsText(sheets), sheets };
}

// ── csv ──────────────────────────────────────────────────
// 엑셀이 저장한 csv는 UTF-8(BOM) 또는 CP949. UTF-8로 읽히지 않으면 CP949로 본다.
function decodeText(buf) {
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.subarray(3).toString('utf8');
  if (buf[0] === 0xff && buf[1] === 0xfe) return new TextDecoder('utf-16le').decode(buf.subarray(2));
  if (buf[0] === 0xfe && buf[1] === 0xff) return new TextDecoder('utf-16be').decode(buf.subarray(2));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch {}
  try { return new TextDecoder('euc-kr').decode(buf); } catch { return buf.toString('latin1'); }
}

export function parseCsv(text, sep) {
  if (!sep) {
    const first = text.slice(0, text.search(/\r?\n|$/));
    sep = first.includes(',') ? ',' : first.includes('\t') ? '\t' : first.includes(';') ? ';' : ',';
  }
  const rows = [];
  let row = []; let cell = ''; let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false; } else cell += ch;
    } else if (ch === '"' && cell === '') quoted = true;
    else if (ch === sep) { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

export async function readCsv(path) {
  const rows = parseCsv(decodeText(readFileSync(path)));
  const sheets = [{ name: basename(path, extname(path)), rows }];
  return { text: sheetsText(sheets), sheets };
}

// ── 바꾸기 ───────────────────────────────────────────────
// 글자가 여러 조각(run)으로 쪼개져 있어도 찾을 수 있도록 문단(block) 단위로 이어 붙여 찾는다.
// 고치는 것은 일치한 부분이 걸친 조각뿐이다: 바꿀 말은 일치가 시작된 조각에 넣고,
// 뒤 조각에서는 일치한 부분만 지운다. 탭·줄바꿈·다른 조각과 그 서식은 그대로 둔다.
//   block: 'w:p' | 'a:p' | 'si' | 'is'   text: 'w:t' | 'a:t' | 't'   skip: 무시할 요소 (rPh 읽는 법 등)
//   weight(n): n번째 block의 일치 수에 곱할 값 (공유 문자열을 쓰는 셀 수)
function replaceInXml(xml, { block, text, skip, preserve = false }, pairs, weight = () => 1) {
  const parts = [`<(/?)${block}(?=[\\s>/])[^>]*>`, `(<${text}(?:\\s[^>]*)?>)([\\s\\S]*?)(</${text}>)`];
  if (skip) parts.push(`<${skip}(?=[\\s>])[\\s\\S]*?</${skip}>`);
  const re = new RegExp(parts.join('|'), 'g');
  const runs = []; const stack = []; let blocks = 0;
  for (const m of xml.matchAll(re)) {
    if (m[2] !== undefined) {
      if (stack.length) runs.push({ at: m.index, end: m.index + m[0].length, open: m[2], text: decode(m[3]), close: m[4], block: stack.at(-1) });
    } else if (m[1] === '/') stack.pop();
    else if (m[1] === '' && !m[0].endsWith('/>')) stack.push(blocks++);
  }
  const counts = pairs.map(() => 0);
  const byBlock = new Map();
  for (const r of runs) { if (!byBlock.has(r.block)) byBlock.set(r.block, []); byBlock.get(r.block).push(r); }
  for (const [b, list] of byBlock) {
    pairs.forEach(([find, repl], k) => {
      const joined = list.map((r) => r.text).join('');
      if (!find || !joined.includes(find)) return;
      // 글자마다 어느 조각에 속하는지
      const owner = new Array(joined.length);
      let p = 0;
      list.forEach((r, i) => { for (let j = 0; j < r.text.length; j++) owner[p++] = i; });
      const next = list.map(() => '');
      let hits = 0;
      for (let i = 0; i < joined.length;) {
        if (joined.startsWith(find, i)) { next[owner[i]] += repl; i += find.length; hits++; } else { next[owner[i]] += joined[i]; i++; }
      }
      list.forEach((r, i) => { if (r.text !== next[i]) { r.text = next[i]; r.changed = true; } });
      counts[k] += hits * weight(b);
    });
  }
  // 바뀐 조각만 다시 쓴다
  let out = ''; let pos = 0;
  for (const r of runs) {
    if (!r.changed) continue;
    let open = r.open;
    if (preserve && !/xml:space=/.test(open) && /^\s|\s$/.test(r.text)) open = open.replace(/>$/, ' xml:space="preserve">');
    out += xml.slice(pos, r.at) + open + encode(r.text) + r.close;
    pos = r.end;
  }
  return { xml: out + xml.slice(pos), counts, changed: pos > 0 };
}

const DOCX_PARTS = (zip) => Object.keys(zip.files).filter((f) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(f));
const PPTX_PARTS = (zip) => Object.keys(zip.files).filter((f) => /^ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml$/.test(f));
const DOCX_TEXT = { block: 'w:p', text: 'w:t', preserve: true };
const PPTX_TEXT = { block: 'a:p', text: 'a:t' };

async function replaceInZip(path, files, spec, pairs, output) {
  const zip = await loadZip(path);
  const counts = pairs.map(() => 0);
  for (const f of files(zip)) {
    const r = replaceInXml(await zip.file(f).async('string'), spec, pairs);
    r.counts.forEach((n, i) => { counts[i] += n; });
    if (r.changed) zip.file(f, r.xml);
  }
  await saveZip(zip, output);
  return counts;
}

// ── xlsx 부품 ────────────────────────────────────────────

// 루트 요소의 접두어 ('x:' 등, 보통은 없음)
const prefixOf = (xml, root) => xml.match(new RegExp(`<(\\w+:)?${root}[\\s>]`))?.[1] ?? '';
const attr = (tag, name) => tag.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];

/** 시트 목록 [{ name, part, kind: 'worksheet'|'chartsheet'|... }] (통합 문서 순서) */
async function xlsxSheets(zip, path) {
  const wb = await partText(zip, 'xl/workbook.xml', path);
  const rels = await partText(zip, 'xl/_rels/workbook.xml.rels', path);
  const target = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id'); const t = attr(m[0], 'Target') ?? ''; const type = attr(m[0], 'Type') ?? '';
    target[id] = { part: t.startsWith('/') ? t.slice(1) : `xl/${t}`.replace(/\/[^/]+\/\.\.\//g, '/'), kind: type.split('/').pop() };
  }
  const p = prefixOf(wb, 'workbook');
  return [...wb.matchAll(new RegExp(`<${p}sheet\\b[^>]*>`, 'g'))].map((m) => {
    const id = m[0].match(/\sr:id="([^"]*)"|\s\w+:id="([^"]*)"/);
    const t = target[id?.[1] ?? id?.[2]] ?? {};
    return { name: decode(attr(m[0], 'name') ?? ''), part: t.part, kind: t.kind };
  });
}

// 공유 문자열 번호별로 그것을 쓰는 셀 수
function sharedRefs(sheetXmls) {
  const refs = new Map();
  for (const xml of sheetXmls) {
    for (const m of xml.matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)) {
      if (!/\st="s"/.test(m[1]) || !m[2]) continue;
      const v = m[2].match(/<(?:\w+:)?v>(\d+)<\/(?:\w+:)?v>/);
      if (v) refs.set(Number(v[1]), (refs.get(Number(v[1])) ?? 0) + 1);
    }
  }
  return refs;
}

// 다음에 Excel로 열 때 수식을 모두 다시 계산하게 한다 (계산 순서 캐시도 버림)
async function forceRecalc(zip, path, { dropChain = false } = {}) {
  let wb = await partText(zip, 'xl/workbook.xml', path);
  const p = prefixOf(wb, 'workbook');
  const calc = wb.match(new RegExp(`<${p}calcPr\\b[^>]*?/?>`));
  if (calc) {
    const tag = /\sfullCalcOnLoad="/.test(calc[0]) ? calc[0].replace(/\sfullCalcOnLoad="[^"]*"/, ' fullCalcOnLoad="1"')
      : calc[0].replace(/\s*(\/?)>$/, ' fullCalcOnLoad="1"$1>');
    wb = wb.replace(calc[0], tag);
  } else {
    // calcPr는 정해진 순서 자리에 넣어야 한다 (뒤에 올 요소 앞, 없으면 끝)
    const after = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing',
      'fileRecoveryPr', 'webPublishObjects', 'extLst'].map((t) => wb.search(new RegExp(`<${p}${t}[\\s>/]`))).filter((i) => i >= 0);
    const at = after.length ? Math.min(...after) : wb.lastIndexOf(`</${p}workbook>`);
    wb = `${wb.slice(0, at)}<${p}calcPr fullCalcOnLoad="1"/>${wb.slice(at)}`;
  }
  zip.file('xl/workbook.xml', wb);
  if (dropChain && zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/g, ''));
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, ''));
  }
}

async function replaceInXlsx(path, pairs, output) {
  const zip = await loadZip(path);
  const sheets = (await xlsxSheets(zip, path)).filter((s) => s.kind === 'worksheet' && s.part && zip.file(s.part));
  const xmls = await Promise.all(sheets.map((s) => zip.file(s.part).async('string')));
  const counts = pairs.map(() => 0);
  const add = (r) => r.counts.forEach((n, i) => { counts[i] += n; });
  const sst = 'xl/sharedStrings.xml';
  if (zip.file(sst)) {
    const xml = await zip.file(sst).async('string');
    const p = prefixOf(xml, 'sst');
    const refs = sharedRefs(xmls);
    const r = replaceInXml(xml, { block: `${p}si`, text: `${p}t`, skip: `${p}rPh`, preserve: true }, pairs, (n) => refs.get(n) ?? 0);
    add(r);
    if (r.changed) zip.file(sst, r.xml);
  }
  // 셀 안에 직접 든 글자 (inlineStr)
  sheets.forEach((s, i) => {
    const p = prefixOf(xmls[i], 'worksheet');
    const r = replaceInXml(xmls[i], { block: `${p}is`, text: `${p}t`, skip: `${p}rPh`, preserve: true }, pairs);
    add(r);
    if (r.changed) zip.file(s.part, r.xml);
  });
  if (counts.some(Boolean)) await forceRecalc(zip, path);
  await saveZip(zip, output);
  return counts;
}

/** 찾아 바꾸기 → 쌍마다 바꾼 수 [n, ...] */
export async function replaceInFile(path, ext, pairs, output = path) {
  if (ext === 'docx') return replaceInZip(path, DOCX_PARTS, DOCX_TEXT, pairs, output);
  if (ext === 'pptx') return replaceInZip(path, PPTX_PARTS, PPTX_TEXT, pairs, output);
  if (ext === 'xlsx') return replaceInXlsx(path, pairs, output);
  throw new Error(`파일 엔진이 지원하지 않는 형식: .${ext}`);
}

// ── {{자리표시자}} ────────────────────────────────────────
export const PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;

// 바꾸기와 같은 단위(문단마다 이어 붙인 글자)에서 찾아야 그대로 바꿀 수 있다
function blockTexts(xml, { block, text, skip }) {
  const out = [];
  const parts = [`<(/?)${block}(?=[\\s>/])[^>]*>`, `<${text}(?:\\s[^>]*)?>([\\s\\S]*?)</${text}>`];
  if (skip) parts.push(`<${skip}(?=[\\s>])[\\s\\S]*?</${skip}>`);
  const stack = [];
  for (const m of xml.matchAll(new RegExp(parts.join('|'), 'g'))) {
    if (m[2] !== undefined) { if (stack.length) out[stack.at(-1)] += decode(m[2]); }
    else if (m[1] === '/') stack.pop();
    else if (m[1] === '' && !m[0].endsWith('/>')) { stack.push(out.length); out.push(''); }
  }
  return out;
}

async function searchableTexts(path, ext) {
  if (ext === 'csv') return [(await readCsv(path)).text];
  const zip = await loadZip(path);
  const texts = [];
  const scan = async (files, spec) => { for (const f of files) texts.push(...blockTexts(await zip.file(f).async('string'), spec)); };
  if (ext === 'docx') await scan(DOCX_PARTS(zip), DOCX_TEXT);
  else if (ext === 'pptx') await scan(PPTX_PARTS(zip), PPTX_TEXT);
  else if (ext === 'xlsx') {
    if (zip.file('xl/sharedStrings.xml')) {
      const xml = await zip.file('xl/sharedStrings.xml').async('string');
      const p = prefixOf(xml, 'sst');
      texts.push(...blockTexts(xml, { block: `${p}si`, text: `${p}t`, skip: `${p}rPh` }));
    }
    for (const s of await xlsxSheets(zip, path)) {
      if (s.kind !== 'worksheet' || !zip.file(s.part)) continue;
      const xml = await zip.file(s.part).async('string');
      const p = prefixOf(xml, 'worksheet');
      texts.push(...blockTexts(xml, { block: `${p}is`, text: `${p}t`, skip: `${p}rPh` }));
    }
  } else throw new Error(`파일 엔진이 지원하지 않는 형식: .${ext}`);
  return texts;
}

/** {{이름}} 자리표시자 목록 (본문·머리글·바닥글·각주·슬라이드 노트 포함) → [{ name, value, kind, raw: [실제 표기...] }] */
export async function placeholders(path, ext) {
  const found = new Map();
  for (const t of await searchableTexts(path, ext)) {
    for (const m of t.matchAll(PLACEHOLDER)) {
      if (!found.has(m[1])) found.set(m[1], new Set());
      found.get(m[1]).add(m[0]);
    }
  }
  return [...found].map(([name, raw]) => ({ name, value: '', kind: 'placeholder', raw: [...raw] }));
}

// ── xlsx 셀 쓰기 ─────────────────────────────────────────

const colNum = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

function parseAddr(addr) {
  const m = String(addr).replace(/\$/g, '').toUpperCase().match(/^([A-Z]{1,3})(\d{1,7})$/);
  if (!m || colNum(m[1]) > 16384 || Number(m[2]) < 1 || Number(m[2]) > 1048576) throw new Error(`셀 주소가 올바르지 않습니다: ${addr} (예: A1, B12)`);
  return { col: colNum(m[1]), row: Number(m[2]), ref: `${m[1]}${Number(m[2])}` };
}

function cellXml(p, ref, style, value) {
  const s = style ? ` s="${style}"` : '';
  if (value === null || value === undefined || value === '') return `<${p}c r="${ref}"${s}/>`;
  if (typeof value === 'number' && Number.isFinite(value)) return `<${p}c r="${ref}"${s}><${p}v>${value}</${p}v></${p}c>`;
  if (typeof value === 'boolean') return `<${p}c r="${ref}"${s} t="b"><${p}v>${value ? 1 : 0}</${p}v></${p}c>`;
  const str = String(value);
  if (str.startsWith('=') && str.length > 1) return `<${p}c r="${ref}"${s}><${p}f>${encode(str.slice(1))}</${p}f></${p}c>`;
  const keep = /^\s|\s$/.test(str) ? ' xml:space="preserve"' : '';
  return `<${p}c r="${ref}"${s} t="inlineStr"><${p}is><${p}t${keep}>${encode(str)}</${p}t></${p}is></${p}c>`;
}

// 시트 XML 하나에 셀 값을 쓴다 (행·셀이 없으면 순서에 맞춰 만든다)
function setCellInSheet(xml, addr, value) {
  const p = prefixOf(xml, 'worksheet');
  const { col, row, ref } = parseAddr(addr);
  if (new RegExp(`<${p}sheetData\\s*/>`).test(xml)) xml = xml.replace(new RegExp(`<${p}sheetData\\s*/>`), `<${p}sheetData></${p}sheetData>`);
  const dataOpen = xml.match(new RegExp(`<${p}sheetData\\b[^>]*>`));
  const dataEnd = xml.indexOf(`</${p}sheetData>`);
  if (!dataOpen || dataEnd < 0) throw new Error('시트 구조를 읽을 수 없습니다');
  const dataStart = dataOpen.index + dataOpen[0].length;
  const data = xml.slice(dataStart, dataEnd);

  const rowRe = new RegExp(`<${p}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}row>)`, 'g');
  let newData;
  let found = null; let insertAt = data.length;
  const odd = () => new Error('이 엑셀 파일은 행·셀 번호가 빠져 있어 프로그램 없이 셀을 쓸 수 없습니다 (Excel이 있으면 engine "app")');
  for (const m of data.matchAll(rowRe)) {
    const r = Number(attr(m[1], 'r'));
    if (!r) throw odd();
    if (r === row) { found = m; break; }
    if (r > row) { insertAt = m.index; break; }
  }
  if (!found) {
    newData = `${data.slice(0, insertAt)}<${p}row r="${row}">${cellXml(p, ref, null, value)}</${p}row>${data.slice(insertAt)}`;
  } else {
    const attrs = found[1].replace(/\sspans="[^"]*"/, ''); // 칸 범위 힌트는 바뀔 수 있어 뺀다
    const cells = found[2] ?? '';
    const cellRe = new RegExp(`<${p}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}c>)`, 'g');
    let cell = null; let at = cells.length;
    for (const m of cells.matchAll(cellRe)) {
      const c = attr(m[1], 'r');
      if (!c) throw odd();
      const n = colNum(c.replace(/\d+/g, ''));
      if (n === col) { cell = m; break; }
      if (n > col) { at = m.index; break; }
    }
    let body;
    if (cell) {
      if (/<(?:\w+:)?f\b[^>]*\bt="(?:shared|array)"[^>]*\bref="/.test(cell[2] ?? '')) {
        throw new Error(`${ref} 셀은 여러 셀이 함께 쓰는 수식의 기준 셀이라 프로그램 없이 바꿀 수 없습니다 (Excel이 있으면 engine 'app')`);
      }
      body = cells.slice(0, cell.index) + cellXml(p, ref, attr(cell[1], 's'), value) + cells.slice(cell.index + cell[0].length);
    } else {
      // 새 셀은 그 행의 서식을 따른다
      const rowStyle = /\scustomFormat="1"/.test(found[1]) ? attr(found[1], 's') : null;
      body = cells.slice(0, at) + cellXml(p, ref, rowStyle, value) + cells.slice(at);
    }
    newData = `${data.slice(0, found.index)}<${p}row${attrs}>${body}</${p}row>${data.slice(found.index + found[0].length)}`;
  }
  xml = xml.slice(0, dataStart) + newData + xml.slice(dataEnd);

  // 사용 범위(dimension) 넓히기
  const dim = xml.match(new RegExp(`<${p}dimension\\b[^>]*\\sref="([^"]*)"[^>]*/>`));
  if (dim) {
    const [a, b = a] = dim[1].split(':');
    try {
      const x = parseAddr(a); const y = parseAddr(b);
      const ref2 = `${colName(Math.min(x.col, col))}${Math.min(x.row, row)}:${colName(Math.max(y.col, col))}${Math.max(y.row, row)}`;
      xml = xml.replace(dim[0], dim[0].replace(/\sref="[^"]*"/, ` ref="${ref2}"`));
    } catch { /* 이상한 범위는 그대로 */ }
  }
  return xml;
}

/** 엑셀 셀 쓰기 (파일 엔진). 차트·그림·피벗 등은 그대로 남는다 */
export async function setXlsxCells(path, cells, { sheet, output = path } = {}) {
  const zip = await loadZip(path);
  const all = await xlsxSheets(zip, path);
  const ws = sheet ? all.find((s) => s.name === sheet) : all.find((s) => s.kind === 'worksheet');
  if (!ws) throw new Error(`시트 없음: ${sheet ?? '(일반 시트가 없습니다)'}`);
  if (ws.kind !== 'worksheet' || !ws.part || !zip.file(ws.part)) throw new Error(`셀을 쓸 수 없는 시트입니다(차트 시트 등): ${ws.name}`);
  let xml = await zip.file(ws.part).async('string');
  for (const [addr, value] of Object.entries(cells)) xml = setCellInSheet(xml, addr, value);
  zip.file(ws.part, xml);
  await forceRecalc(zip, path, { dropChain: true });
  await saveZip(zip, output);
  return Object.keys(cells).length;
}
