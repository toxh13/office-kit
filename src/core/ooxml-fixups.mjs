// ooxml-fixups.mjs — Office 편집기(x2t)가 내보낸 docx·xlsx·pptx를 저장 직전에 고친다 (Word·Excel·PowerPoint가 경고 없이 열게).
//   pptx: 노트·유인물 마스터가 다른 마스터와 테마 부분을 같이 쓰면 사본을 따로 준다 (없으면 PowerPoint 복구 창 — 필수)
//   xlsx: 어떤 관계도 가리키지 않는 xl/theme/*.xml을 지운다 (x2t가 남기는 theme.xml)
//         x14 확장에만 남은 보통 조건부 서식(cellIs 등)을 기본 조건부 서식으로 되돌린다 (Excel 없는 파일 엔진이 고칠 수 있게)
//   docx: 아래 행이 이어 받지 않는 <w:vMerge w:val="restart"/>를 지운다 (편집기에서 더한 표 행 → Word가 세로 병합으로 봄)
// 원본에서 옮겨 오기 (덮어쓸 원래 파일을 주면 — 편집기에 화면이 없는 것만, 사용자 편집은 건드리지 않음):
//   공통: 테마 글꼴 묶음(스크립트별 글꼴·빈 ea) — x2t가 지워 한글이 맑은 고딕 대신 바탕·Arial로 보임
//   docx: w:themeFontLang(eastAsia), 기본 문단 스타일의 wordWrap·autoSpaceDE·autoSpaceDN, w:compat 항목, 내용 컨트롤 종류(w:text 등)
//   xlsx: 테마 글꼴 이름(맑은 고딕 → Calibri로 바뀐 것), 원본에 없던 스레드 메모 → 원래 메모
//   pptx: 구역(p14:sectionLst)
// 깨뜨리지 않는다: 실패하면 원래 바이트를 그대로 돌려주고 error를 채운다 (저장은 막지 않음). 원본 옮기기만 실패하면 그것만 빼고 고친다.
// 고친 결과는 다시 열어 확인한다: 부분 목록, 손대지 않은 부분의 CRC·크기, 고친 XML의 짝, 고칠 것이 남지 않았는지.
// 부분은 한도 안에서만 푼다 (부분마다 64MB, 모두 256MB) — 압축 폭탄으로 주 프로세스 메모리를 다 쓰지 않게.
import { inflateRawSync } from 'node:zlib';
import JSZip from 'jszip';

/** 고치기가 있는 형식 */
export const FIXUP_EXTS = ['docx', 'xlsx', 'pptx'];
/** 푸는 한도: 부분 하나 · 한 꾸러미 모두 */
export const PART_MAX_BYTES = 64 * 1024 * 1024;
export const TOTAL_MAX_BYTES = 256 * 1024 * 1024;

const LABEL = { docx: 'Word', xlsx: 'Excel', pptx: 'PowerPoint' };
const THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml';
// 마스터 종류 (앞일수록 테마 주인): 내용 형식 끝 · 기본 경로
const MASTERS = [
  { kind: '슬라이드 마스터', ct: 'presentationml.slideMaster+xml', path: /^ppt\/slideMasters\/[^/]+\.xml$/i },
  { kind: '노트 마스터', ct: 'presentationml.notesMaster+xml', path: /^ppt\/notesMasters\/[^/]+\.xml$/i },
  { kind: '유인물 마스터', ct: 'presentationml.handoutMaster+xml', path: /^ppt\/handoutMasters\/[^/]+\.xml$/i },
];
const REL_TYPE = (name) => new RegExp(`/relationships/${name}$`);
const IS_THEME_REL = REL_TYPE('theme');
const IS_MAIN_REL = REL_TYPE('officeDocument');
// docx 표를 고칠 부분 (주 문서 + 이것들)
const WORD_STORIES = ['header', 'footer', 'footnotes', 'endnotes', 'comments'].map(REL_TYPE);

const fail = (msg) => { throw new Error(msg); };
const lower = (s) => s.toLowerCase();
const num = (name) => Number(/(\d+)\.xml$/i.exec(name)?.[1] ?? 0);

// ── XML 글자 ─────────────────────────────────────────────

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const unescapeAttr = (s) => s.replace(/&(amp|lt|gt|quot|apos);/g, (_, e) => ENT[e]);
const escapeAttr = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

function attrsOf(tag) {
  const out = {};
  for (const m of tag.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) out[m[1]] = unescapeAttr(m[2] ?? m[3]);
  return out;
}

// 한도 안에서 부분 바이트 풀기. 적힌 크기가 한도를 넘으면 풀지 않고, 거짓 크기여도 풀면서 한도에서 멈춘다
const STORE = '\x00\x00';
const DEFLATE = '\x08\x00';
function newBudget() { return { used: 0 }; }
async function partBytes(f, name, budget) {
  const d = f._data;
  let out;
  if (d && d.compressedContent instanceof Uint8Array && d.compression && typeof d.uncompressedSize === 'number') {
    if (d.uncompressedSize > PART_MAX_BYTES) fail(`부분이 너무 큼: ${name}`);
    const raw = d.compressedContent;
    if (d.compression.magic === STORE) out = raw;
    else if (d.compression.magic === DEFLATE) {
      try { out = inflateRawSync(raw, { maxOutputLength: PART_MAX_BYTES }); } catch (err) {
        fail(err?.code === 'ERR_BUFFER_TOO_LARGE' || err instanceof RangeError ? `부분이 너무 큼: ${name}` : `부분을 풀지 못함: ${name}`);
      }
    } else fail(`모르는 압축 방식: ${name}`);
  } else {
    out = await f.async('uint8array'); // 이 모듈이 넣은 것 (이미 메모리에 있음)
  }
  if (out.length > PART_MAX_BYTES) fail(`부분이 너무 큼: ${name}`);
  if (budget) {
    budget.used += out.length;
    if (budget.used > TOTAL_MAX_BYTES) fail('꾸러미를 풀면 너무 큼');
  }
  return out;
}

// UTF-8만 고친다 (BOM은 그대로 남김). UTF-16·잘못된 UTF-8이면 실패
async function readXml(zip, name, budget) {
  const f = zip.file(name);
  if (!f) return null;
  const b = await partBytes(f, name, budget);
  if ((b[0] === 0xff && b[1] === 0xfe) || (b[0] === 0xfe && b[1] === 0xff)) fail(`UTF-16 XML은 고치지 않음: ${name}`);
  return new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(b);
}

/** 여는·닫는 태그 짝이 맞는지 (고친 XML이 깨지지 않았는지 보는 정도) */
export function xmlBalanced(xml) {
  const stack = [];
  // 마지막 '<' 갈래: 태그가 되지 못한 꺾쇠 (잘린 태그)
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<(\/?)([^\s/>!?<]+)(?:[^<>"']|"[^"]*"|'[^']*')*?(\/?)>|</g;
  let rootSeen = false;
  for (const m of xml.matchAll(re)) {
    if (m[0] === '<') return false;
    if (!m[2]) continue; // 주석·CDATA·처리 지시
    if (m[1]) { if (stack.pop() !== m[2]) return false; continue; }
    if (!stack.length && rootSeen) return false; // 뿌리 요소가 둘
    rootSeen = true;
    if (!m[3]) stack.push(m[2]);
  }
  return rootSeen && stack.length === 0;
}

// ── 꾸러미(OPC) 경로·관계·내용 형식 ─────────────────────────

const dirOf = (part) => part.slice(0, Math.max(0, part.lastIndexOf('/')));
const baseOf = (part) => part.slice(part.lastIndexOf('/') + 1);
const relsPathOf = (part) => (part ? `${dirOf(part) ? `${dirOf(part)}/` : ''}_rels/${baseOf(part)}.rels` : '_rels/.rels');
function sourceOfRels(relsPath) {
  const m = /^(?:(.*)\/)?_rels\/([^/]*)\.rels$/.exec(relsPath); // '_rels/.rels' = 꾸러미 뿌리 ('')
  if (!m) return null;
  return m[2] ? `${m[1] ? `${m[1]}/` : ''}${m[2]}` : '';
}

// 관계의 Target → 꾸러미 안 부분 이름 (앞 '/' 없음). 밖을 가리키면 null
function resolveTarget(source, target) {
  let t = target.split('#')[0];
  if (!t || /^[a-z][a-z0-9+.-]*:/i.test(t)) return null;
  try { t = decodeURIComponent(t); } catch { /* 그대로 */ }
  const segs = t.startsWith('/') ? [] : dirOf(source).split('/').filter(Boolean);
  for (const s of t.split('/')) {
    if (!s || s === '.') continue;
    if (s === '..') { if (!segs.length) return null; segs.pop(); } else segs.push(s);
  }
  return segs.join('/');
}

// from 부분에서 to 부분을 가리키는 상대 경로 ('../theme/theme2.xml')
function relativeTarget(from, to) {
  const a = dirOf(from).split('/').filter(Boolean);
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && lower(a[i]) === lower(b[i])) i++;
  return [...Array(a.length - i).fill('..'), ...b.slice(i)].join('/');
}

function parseRels(xml, source) {
  const out = [];
  for (const m of xml.matchAll(/<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/g)) {
    const a = attrsOf(m[0]);
    const external = a.TargetMode === 'External';
    out.push({ tag: m[0], index: m.index, id: a.Id, type: a.Type ?? '', target: a.Target ?? '', external,
      part: external ? null : resolveTarget(source, a.Target ?? '') });
  }
  return out;
}

function setTarget(xml, rel, target) {
  const tag = rel.tag.replace(/(\sTarget\s*=\s*)(?:"[^"]*"|'[^']*')/, (_, pre) => `${pre}"${escapeAttr(target)}"`);
  if (tag === rel.tag || xml.slice(rel.index, rel.index + rel.tag.length) !== rel.tag) fail(`관계를 고치지 못함: ${rel.id}`);
  return xml.slice(0, rel.index) + tag + xml.slice(rel.index + rel.tag.length);
}

// [Content_Types].xml
async function readTypes(zip, budget) {
  const xml = await readXml(zip, '[Content_Types].xml', budget);
  if (xml == null) fail('[Content_Types].xml 없음');
  const overrides = new Map();
  for (const m of xml.matchAll(/<Override\b(?:[^>"']|"[^"]*"|'[^']*')*?(?:\/>|>\s*<\/Override>)/g)) {
    const a = attrsOf(m[0]);
    if (a.PartName) overrides.set(lower(a.PartName.replace(/^\//, '')), { tag: m[0], type: a.ContentType ?? '' });
  }
  const defaults = new Map();
  for (const m of xml.matchAll(/<Default\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/g)) {
    const a = attrsOf(m[0]);
    if (a.Extension) defaults.set(lower(a.Extension), a.ContentType ?? '');
  }
  return {
    xml, overrides, defaults, changed: false,
    typeOf(part) { return this.overrides.get(lower(part))?.type ?? this.defaults.get(lower(part.split('.').pop())) ?? null; },
    add(part, type) {
      if (this.overrides.has(lower(part))) fail(`내용 형식이 이미 있음: ${part}`);
      const at = this.xml.lastIndexOf('</Types>');
      if (at < 0) fail('[Content_Types].xml 끝을 찾지 못함');
      const tag = `<Override PartName="/${escapeAttr(part)}" ContentType="${escapeAttr(type)}"/>`;
      this.xml = this.xml.slice(0, at) + tag + this.xml.slice(at);
      this.overrides.set(lower(part), { tag, type });
      this.changed = true;
    },
    remove(part) {
      const o = this.overrides.get(lower(part));
      if (!o) return;
      const at = this.xml.indexOf(o.tag);
      if (at < 0) fail(`내용 형식을 지우지 못함: ${part}`);
      this.xml = this.xml.slice(0, at) + this.xml.slice(at + o.tag.length);
      this.overrides.delete(lower(part));
      this.changed = true;
    },
  };
}

// 고칠 계획: 바뀐 부분(새 바이트)·지운 부분. 마지막에 한꺼번에 zip에 넣는다
class Edit {
  constructor(zip, budget = newBudget()) {
    this.zip = zip;
    this.budget = budget;
    this.names = new Map(Object.keys(zip.files).filter((n) => !zip.files[n].dir).map((n) => [lower(n), n]));
    this.put = new Map();      // 이름 → Uint8Array
    this.removed = new Set();
    this.fixes = [];
    this.cache = new Map();    // 풀어 둔 XML (같은 부분을 여러 번 풀지 않게)
  }
  has(part) { const n = this.names.get(lower(part)); return !!n && !this.removed.has(n); }
  name(part) { return this.names.get(lower(part)) ?? part; }
  async xml(part) {
    const n = this.name(part);
    if (this.put.has(n)) return new TextDecoder('utf-8', { ignoreBOM: true }).decode(this.put.get(n));
    if (!this.has(n)) return null;
    if (!this.cache.has(n)) this.cache.set(n, await readXml(this.zip, n, this.budget));
    return this.cache.get(n);
  }
  async bytes(part) {
    const n = this.name(part);
    return this.put.get(n) ?? partBytes(this.zip.file(n), n, this.budget);
  }
  write(part, data) {
    const n = this.name(part);
    this.put.set(n, typeof data === 'string' ? new TextEncoder().encode(data) : data);
    this.names.set(lower(n), n);
    this.removed.delete(n);
  }
  remove(part) { const n = this.name(part); this.removed.add(n); this.put.delete(n); }
  get changed() { return this.put.size > 0 || this.removed.size > 0; }
  // 한 단계가 실패하면 그 단계 앞으로 되돌린다
  snapshot(types) {
    return { put: new Map(this.put), removed: new Set(this.removed), names: new Map(this.names), fixes: this.fixes.length,
      types: types && { xml: types.xml, overrides: new Map(types.overrides), changed: types.changed } };
  }
  restore(s, types) {
    this.put = s.put; this.removed = s.removed; this.names = s.names; this.fixes.length = s.fixes;
    if (types && s.types) { types.xml = s.types.xml; types.overrides = s.types.overrides; types.changed = s.types.changed; }
  }
}

async function mainPart(edit) {
  const rels = await edit.xml('_rels/.rels');
  return rels ? parseRels(rels, '').find((r) => IS_MAIN_REL.test(r.type) && r.part)?.part ?? null : null;
}

// ── pptx: 마스터마다 자기 테마 ─────────────────────────────

async function mastersOf(edit, types) {
  const found = new Map();
  for (const [i, m] of MASTERS.entries()) {
    for (const [part, o] of types.overrides) if (o.type.endsWith(m.ct)) found.set(part, { order: i, kind: m.kind });
    for (const n of edit.names.values()) if (m.path.test(n) && edit.has(n) && !found.has(lower(n))) found.set(lower(n), { order: i, kind: m.kind });
  }
  return [...found].filter(([p]) => edit.has(p)).map(([p, v]) => ({ part: edit.name(p), ...v }))
    .sort((a, b) => a.order - b.order || num(a.part) - num(b.part) || a.part.localeCompare(b.part));
}

// 테마를 다른 마스터와 같이 쓰는 마스터 → [{ master, kind, relsPath, rel, theme, owner, ownerKind }] (첫 주인은 빼고)
async function sharedThemes(edit, types) {
  const owner = new Map();
  const shared = [];
  for (const m of await mastersOf(edit, types)) {
    const relsPath = relsPathOf(m.part);
    const xml = await edit.xml(relsPath);
    if (xml == null) continue;
    for (const rel of parseRels(xml, m.part)) {
      if (rel.external || !rel.part || !IS_THEME_REL.test(rel.type)) continue;
      const key = lower(rel.part);
      if (!owner.has(key)) owner.set(key, m);
      else shared.push({ master: m.part, kind: m.kind, relsPath, rel, theme: rel.part, owner: owner.get(key).part, ownerKind: owner.get(key).kind });
    }
  }
  return shared;
}

function freeThemeName(edit, types, dir) {
  for (let n = 1; ; n++) {
    const name = `${dir ? `${dir}/` : ''}theme${n}.xml`;
    if (!edit.has(name) && !types.overrides.has(lower(name)) && !edit.has(relsPathOf(name))) return name;
  }
}

async function splitSharedThemes(edit, types) {
  // 한 마스터의 관계 파일에 여러 번 손댈 수 있으니 하나씩 다시 읽으며
  for (let guard = 0; guard < 64; guard++) {
    const [s] = await sharedThemes(edit, types);
    if (!s) return;
    if (!edit.has(s.theme)) fail(`${s.kind}가 쓰는 테마 부분이 없음: ${s.theme}`);
    const fresh = freeThemeName(edit, types, dirOf(edit.name(s.theme)));
    edit.write(fresh, await edit.bytes(s.theme));
    const themeRels = relsPathOf(edit.name(s.theme));
    if (edit.has(themeRels)) edit.write(relsPathOf(fresh), await edit.bytes(themeRels)); // 같은 폴더라 상대 경로 그대로
    const rels = await edit.xml(s.relsPath);
    const target = s.rel.target.startsWith('/') ? `/${fresh}` : relativeTarget(s.master, fresh);
    edit.write(s.relsPath, setTarget(rels, s.rel, target));
    types.add(fresh, types.typeOf(s.theme) || THEME_CT);
    edit.fixes.push(`${s.kind} 테마 나눔: ${edit.name(s.theme)} → ${fresh} (${s.master})`);
  }
  fail('마스터 테마 나누기가 끝나지 않음');
}

// ── xlsx: 아무도 안 쓰는 테마 ──────────────────────────────

// 모든 관계가 가리키는 부분 (소문자). 관계 파일을 못 읽으면 실패
async function referencedParts(edit) {
  const out = new Set();
  for (const n of edit.names.values()) {
    if (!edit.has(n)) continue;
    const source = sourceOfRels(n);
    if (source == null) continue;
    for (const r of parseRels(await edit.xml(n), source)) if (r.part) out.add(lower(r.part));
  }
  return out;
}

async function unusedXlsxThemes(edit) {
  const book = await mainPart(edit);
  if (!book || !edit.has(book) || !edit.has(relsPathOf(book))) return []; // 판단할 수 없음 → 손대지 않음
  const used = await referencedParts(edit);
  if (!used.size) return [];
  const themeDir = lower(`${dirOf(book) ? `${dirOf(book)}/` : ''}theme/`);
  return [...edit.names.values()].filter((n) => edit.has(n) && lower(n).startsWith(themeDir)
    && !n.slice(themeDir.length).includes('/') && /\.xml$/i.test(n) && !used.has(lower(n)));
}

async function dropUnusedThemes(edit, types) {
  for (const n of await unusedXlsxThemes(edit)) {
    edit.remove(n);
    if (edit.has(relsPathOf(n))) edit.remove(relsPathOf(n));
    types.remove(n);
    edit.fixes.push(`쓰지 않는 테마 지움: ${n}`);
  }
}

// ── docx: 혼자 있는 세로 병합 시작 ──────────────────────────

const WTOKEN = /<(\/?)w:(tbl|tr|tc|tcPr|trPr|tcPrChange|trPrChange|gridBefore|gridSpan|vMerge|cellMerge|cellIns|cellDel|ins|del)(?=[\s/>])((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
const NESTING = new Set(['tbl', 'tr', 'tc', 'tcPr', 'trPr', 'tcPrChange', 'trPrChange']);
const valOf = (attrs) => /\bw:val\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(attrs)?.slice(1).find((v) => v !== undefined);

/**
 * 표의 행마다, 아래 행이 이어 받지 않는(continue가 없는) vMerge restart를 지운다.
 * 변경 추적이 걸린 표·모르는 값이 있는 표는 그대로 둔다.
 * @returns {{ xml: string, removed: number }}
 */
export function dropLoneVMergeRestarts(xml) {
  const stack = [];   // 열린 요소 이름 (NESTING만)
  const tables = [];  // 열린 표: { rows: [{ gridBefore, cells: [{ span, vm }] }], skip }
  const cuts = [];
  let open = null;    // 닫는 태그가 따로 있는 <w:vMerge ...></w:vMerge>
  const table = () => tables.at(-1);
  const row = () => table()?.rows.at(-1);
  const cell = () => row()?.cells.at(-1);
  const top = (n = 1) => stack.at(-n);
  for (const m of xml.matchAll(WTOKEN)) {
    const [tag, close, name, attrs, self] = m;
    if (close) {
      if (name === 'vMerge') { if (open) { open.end = m.index + tag.length; open = null; } continue; }
      if (!NESTING.has(name)) continue;
      if (stack.pop() !== name) fail('표 XML 짝이 맞지 않음');
      if (name === 'tbl') cuts.push(...loneRestarts(tables.pop()));
      continue;
    }
    const t = table();
    if (name === 'tbl') { if (!self) { tables.push({ rows: [], skip: false }); stack.push(name); } continue; }
    if (!t) { if (NESTING.has(name) && !self) stack.push(name); continue; }
    if (name === 'tr') t.rows.push({ gridBefore: 0, cells: [] });
    else if (name === 'tc') { if (row()) row().cells.push({ span: 1, vm: null }); else t.skip = true; }
    else if (name === 'tcPrChange' || name === 'trPrChange' || name === 'cellMerge' || name === 'cellIns' || name === 'cellDel') t.skip = true;
    else if ((name === 'ins' || name === 'del') && (top() === 'trPr' || top() === 'tcPr')) t.skip = true;
    else if (name === 'gridBefore' && top() === 'trPr' && top(2) === 'tr') {
      const v = Number(valOf(attrs) ?? 0);
      if (Number.isInteger(v) && v >= 0) row().gridBefore = v; else t.skip = true;
    } else if (name === 'gridSpan' && top() === 'tcPr' && top(2) === 'tc') {
      const v = Number(valOf(attrs) ?? 1);
      if (Number.isInteger(v) && v >= 1) cell().span = v; else t.skip = true;
    } else if (name === 'vMerge' && top() === 'tcPr' && top(2) === 'tc') {
      const v = valOf(attrs) ?? 'continue';
      if (v !== 'restart' && v !== 'continue') t.skip = true;
      const vm = { kind: v, start: m.index, end: self ? m.index + tag.length : -1 };
      if (cell()) cell().vm = vm; else t.skip = true;
      if (!self) open = vm;
    }
    if (NESTING.has(name) && !self) stack.push(name);
  }
  if (stack.length || tables.length || open) fail('표 XML 짝이 맞지 않음');
  if (!cuts.length) return { xml, removed: 0 };
  cuts.sort((a, b) => a.start - b.start);
  let out = '';
  let at = 0;
  for (const c of cuts) {
    if (c.start < at || c.end <= c.start) fail('세로 병합 위치가 겹침');
    out += xml.slice(at, c.start);
    at = c.end;
  }
  return { xml: out + xml.slice(at), removed: cuts.length };
}

function loneRestarts(t) {
  if (t.skip) return [];
  const cols = t.rows.map((r) => {
    let c = r.gridBefore;
    return r.cells.map((cell) => { const from = c; c += cell.span; return { ...cell, from, to: c }; });
  });
  const out = [];
  for (const [i, cells] of cols.entries()) {
    for (const c of cells) {
      if (c.vm?.kind !== 'restart') continue;
      if (c.vm.end < 0) return []; // 닫히지 않은 요소: 손대지 않음
      const below = cols[i + 1] ?? [];
      const continued = below.some((d) => d.from <= c.from && c.from < d.to && d.vm?.kind === 'continue');
      if (!continued) out.push(c.vm);
    }
  }
  return out;
}

async function wordStories(edit) {
  const main = await mainPart(edit);
  if (!main || !edit.has(main)) return [];
  const parts = [main];
  const rels = await edit.xml(relsPathOf(main));
  if (rels) for (const r of parseRels(rels, main)) if (r.part && edit.has(r.part) && WORD_STORIES.some((re) => re.test(r.type))) parts.push(edit.name(r.part));
  return [...new Set(parts)];
}

async function dropLoneRestarts(edit) {
  for (const part of await wordStories(edit)) {
    const xml = await edit.xml(part);
    if (!xml.includes('w:vMerge')) continue;
    const r = dropLoneVMergeRestarts(xml);
    if (!r.removed) continue;
    edit.write(part, r.xml);
    edit.fixes.push(`혼자 있는 세로 병합 시작 ${r.removed}개 지움: ${part}`);
  }
}

// ── XML 조각 다루기 (원본 옮기기·조건부 서식용) ─────────────────

const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const TAG = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(\/?)([^\s/>!?<]+)(?:[^<>"']|"[^"]*"|'[^']*')*?(\/?)>/g;
const localOf = (q) => q.slice(q.indexOf(':') + 1);
const splice = (s, a, b, ins = '') => s.slice(0, a) + ins + s.slice(b);

/** 첫 <qname …>…</qname> 또는 <qname …/> (같은 이름이 안에 또 있어도 짝을 맞춤) → { start, openEnd, closeStart, end, tag, attrs, self } | null */
function findEl(xml, qname, from = 0, to = xml.length) {
  const re = new RegExp(`<(/?)${reEsc(qname)}(?=[\\s/>])(?:[^>"']|"[^"]*"|'[^']*')*?(/?)>`, 'g');
  re.lastIndex = from;
  let depth = 0; let el = null;
  for (let m; (m = re.exec(xml)) && m.index < to;) {
    if (!m[1]) {
      if (depth === 0) {
        el = { start: m.index, openEnd: m.index + m[0].length, tag: m[0], attrs: attrsOf(m[0]) };
        if (m[2]) return { ...el, closeStart: el.openEnd, end: el.openEnd, self: true };
      }
      if (!m[2]) depth++;
    } else {
      if (!el) fail(`XML 짝이 맞지 않음: ${qname}`);
      if (--depth === 0) return { ...el, closeStart: m.index, end: m.index + m[0].length, self: false };
    }
  }
  if (el) fail(`XML 짝이 맞지 않음: ${qname}`);
  return null;
}
function findAll(xml, qname, from = 0, to = xml.length) {
  const out = [];
  for (let at = from, el; (el = findEl(xml, qname, at, to)) && el.end <= to; at = el.end) out.push(el);
  return out;
}
/** 조각 [from, to) 바로 아래 요소들 → [{ name, start, end }] */
function childrenOf(xml, from, to) {
  const out = []; let depth = 0; let cur = null;
  const re = new RegExp(TAG.source, 'g');
  re.lastIndex = from;
  for (let m; (m = re.exec(xml)) && m.index < to;) {
    if (!m[2]) continue;
    if (m[1]) {
      if (--depth < 0) fail('XML 짝이 맞지 않음');
      if (depth === 0 && cur) { cur.end = m.index + m[0].length; out.push(cur); cur = null; }
      continue;
    }
    if (depth === 0) {
      if (m[3]) { out.push({ name: m[2], start: m.index, end: m.index + m[0].length }); continue; }
      cur = { name: m[2], start: m.index };
    }
    if (!m[3]) depth++;
  }
  if (depth !== 0) fail('XML 짝이 맞지 않음');
  return out;
}
/** <a/> → <a></a> (안에 넣을 수 있게) */
function opened(xml, el) {
  if (!el.self) return { xml, el };
  const open = el.tag.replace(/\s*\/>$/, '>');
  const close = `</${/^<([^\s/>]+)/.exec(el.tag)[1]}>`;
  const x = splice(xml, el.start, el.end, open + close);
  return { xml: x, el: { ...el, openEnd: el.start + open.length, closeStart: el.start + open.length, end: el.start + open.length + close.length, self: false, tag: open } };
}
/** 부모 안에 순서표(order: 이름 끝부분)대로 자식 넣기 → 새 XML */
function insertOrdered(xml, parent, child, local, order) {
  const want = order.indexOf(local);
  const rank = (n) => { const i = order.indexOf(localOf(n)); return i < 0 ? Infinity : i; };
  const kids = childrenOf(xml, parent.openEnd, parent.closeStart);
  const before = want < 0 ? null : kids.find((k) => rank(k.name) > want);
  return splice(xml, before ? before.start : parent.closeStart, before ? before.start : parent.closeStart, child);
}
const textOf = (xml) => unescapeText(xml.replace(/<[^>]*>/g, ''));
function unescapeText(s) {
  return s.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-f]+|#\d+);/gi, (_, e) => (e[0] === '#'
    ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENT[e]));
}
const escapeText = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// 관계 종류로 부분 찾기 (source 부분의 관계에서)
async function relParts(e, source, type) {
  const rels = await e.xml(relsPathOf(source));
  if (!rels) return [];
  return parseRels(rels, source).filter((r) => !r.external && r.part && REL_TYPE(type).test(r.type) && e.has(r.part)).map((r) => ({ ...r, part: e.name(r.part) }));
}

// ── 원본에서 옮겨 오기 ────────────────────────────────────

// 테마 글꼴 묶음: { name, major, minor, scheme(XML) }
function themeFonts(xml) {
  const root = /<a:theme\b(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(xml);
  const fs = findEl(xml, 'a:fontScheme');
  if (!root || !fs || fs.self) return null;
  const latin = (q) => {
    const f = findEl(xml, q, fs.openEnd, fs.closeStart);
    return f && !f.self ? findEl(xml, 'a:latin', f.openEnd, f.closeStart)?.attrs.typeface ?? null : null;
  };
  return { name: attrsOf(root[0]).name ?? '', major: latin('a:majorFont'), minor: latin('a:minorFont'), fs, scheme: xml.slice(fs.start, fs.end) };
}
async function themesOf(e) {
  const out = [];
  for (const n of e.names.values()) {
    if (!e.has(n) || !/(^|\/)theme\/theme\d*\.xml$/i.test(n)) continue;
    const xml = await e.xml(n);
    const t = themeFonts(xml);
    if (t) out.push({ part: n, xml, ...t });
  }
  return out;
}

/** 테마 이름·제목/본문 라틴 글꼴이 같은 원본 테마가 있으면 글꼴 묶음을 원본 것으로 (편집기에서 테마 글꼴을 바꿨으면 손대지 않음) */
async function carryThemeFonts(edit, orig) {
  const before = await themesOf(orig);
  for (const t of await themesOf(edit)) {
    const same = before.filter((o) => o.name === t.name && o.major === t.major && o.minor === t.minor);
    const o = same.find((x) => lower(x.part) === lower(t.part)) ?? same[0];
    if (!o || o.scheme === t.scheme) continue;
    edit.write(t.part, splice(t.xml, t.fs.start, t.fs.end, o.scheme));
    edit.fixes.push(`테마 글꼴 묶음 원본대로: ${t.part}`);
  }
}

// docx: 주 문서와 그 관계 부분
async function wordPart(e, type) {
  const main = await mainPart(e);
  if (!main || !e.has(main)) return null;
  return (await relParts(e, main, type))[0]?.part ?? null;
}

const SETTINGS_ORDER = ['writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries',
  'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData', 'embedTrueTypeFonts',
  'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges', 'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors', 'activeWritingStyle', 'proofState',
  'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter', 'stylePaneSortMethod', 'documentType', 'mailMerge',
  'revisionView', 'trackRevisions', 'doNotTrackMoves', 'doNotTrackFormatting', 'documentProtection', 'autoFormatOverride',
  'styleLockTheme', 'styleLockQFSet', 'defaultTabStop', 'autoHyphenation', 'consecutiveHyphenLimit', 'hyphenationZone',
  'doNotHyphenateCaps', 'showEnvelope', 'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle', 'evenAndOddHeaders',
  'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing',
  'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl',
  'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema',
  'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly',
  'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults',
  'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping',
  'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType',
  'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];

/** w:themeFontLang: x2t가 eastAsia를 zh-CN으로 바꾼다 → 원본 값 (원본에 없으면 문서 기본 언어의 eastAsia) */
async function carryThemeFontLang(edit, orig) {
  const part = await wordPart(edit, 'settings');
  if (!part) return;
  let xml = await edit.xml(part);
  const cur = findEl(xml, 'w:themeFontLang');
  const oPart = orig && await wordPart(orig, 'settings');
  const oXml = oPart ? await orig.xml(oPart) : null;
  const o = oXml ? findEl(oXml, 'w:themeFontLang') : null;
  let want = null;
  if (o) want = oXml.slice(o.start, o.end);
  else if (cur) {
    // 원본이 없거나 그 값이 없으면: 문서 기본 언어(docDefaults w:lang)의 eastAsia로
    const sPart = await wordPart(edit, 'styles');
    const sx = sPart ? await edit.xml(sPart) : null;
    const dd = sx ? findEl(sx, 'w:docDefaults') : null;
    const lang = dd && !dd.self ? findEl(sx, 'w:lang', dd.openEnd, dd.closeStart) : null;
    const ea = lang?.attrs['w:eastAsia'];
    if (!ea || cur.attrs['w:eastAsia'] === ea) return;
    want = cur.tag.replace(/(\sw:eastAsia\s*=\s*)(?:"[^"]*"|'[^']*')/, (_, p) => `${p}"${escapeAttr(ea)}"`);
    if (want === cur.tag) want = cur.tag.replace(/\s*\/>$/, ` w:eastAsia="${escapeAttr(ea)}"/>`);
  }
  if (!want || (cur && xml.slice(cur.start, cur.end) === want)) return;
  if (cur) xml = splice(xml, cur.start, cur.end, want);
  else {
    const root = findEl(xml, 'w:settings');
    if (!root || root.self) return;
    xml = insertOrdered(xml, root, want, 'themeFontLang', SETTINGS_ORDER);
  }
  edit.write(part, xml);
  edit.fixes.push(`문서 글꼴 언어 원본대로: ${want}`);
}

const COMPAT_ORDER = ['useSingleBorderforContiguousCells', 'wpJustification', 'noTabHangInd', 'noLeading', 'spaceForUL', 'noColumnBalance',
  'balanceSingleByteDoubleByteWidth', 'noExtraLineSpacing', 'doNotLeaveBackslashAlone', 'ulTrailSpace', 'doNotExpandShiftReturn',
  'spacingInWholePoints', 'lineWrapLikeWord6', 'printBodyTextBeforeHeader', 'printColBlack', 'wpSpaceWidth', 'showBreaksInFrames',
  'subFontBySize', 'suppressBottomSpacing', 'suppressTopSpacing', 'suppressSpacingAtTopOfPage', 'suppressTopSpacingWP',
  'suppressSpBfAfterPgBrk', 'swapBordersFacingPages', 'convMailMergeEsc', 'truncateFontHeightsLikeWP6', 'mwSmallCaps',
  'usePrinterMetrics', 'doNotSuppressParagraphBorders', 'wrapTrailSpaces', 'footnoteLayoutLikeWW8', 'shapeLayoutLikeWW8',
  'alignTablesRowByRow', 'forgetLastTabAlignment', 'adjustLineHeightInTable', 'autoSpaceLikeWord95', 'noSpaceRaiseLower',
  'doNotUseHTMLParagraphAutoSpacing', 'layoutRawTableWidth', 'layoutTableRowsApart', 'useWord97LineBreakRules',
  'doNotBreakWrappedTables', 'doNotSnapToGridInCell', 'selectFldWithFirstOrLastChar', 'applyBreakingRules', 'doNotWrapTextWithPunct',
  'doNotUseEastAsianBreakRules', 'useWord2002TableStyleRules', 'growAutofit', 'useFELayout', 'useNormalStyleForList',
  'doNotUseIndentAsNumberingTabStop', 'useAltKinsokuLineBreakRules', 'allowSpaceOfSameStyleInTable', 'doNotSuppressIndentation',
  'doNotAutofitConstrainedTables', 'autofitToFirstFixedWidthCell', 'underlineTabInNumList', 'displayHangulFixedWidth',
  'splitPgBreakAndParaMark', 'doNotVertAlignCellWithSp', 'doNotBreakConstrainedForcedTable', 'doNotVertAlignInTxbx',
  'useAnsiKerningPairs', 'cachedColBandSize', 'compatSetting'];

/** w:compat: 원본에 있고 저장본에 없는 항목(표 줄 높이·밑줄 간격 등, 편집기에 화면이 없는 것)을 되살린다 */
async function carryCompat(edit, orig) {
  const part = await wordPart(edit, 'settings');
  const oPart = await wordPart(orig, 'settings');
  if (!part || !oPart) return;
  let xml = await edit.xml(part);
  const oXml = await orig.xml(oPart);
  const oc = findEl(oXml, 'w:compat');
  let c = findEl(xml, 'w:compat');
  if (!oc || oc.self || !c) return;
  ({ xml, el: c } = opened(xml, c));
  const key = (x, k) => (k.name === 'w:compatSetting' ? `cs:${attrsOf(x.slice(k.start, k.end))['w:name'] ?? ''}` : k.name);
  const have = new Set(childrenOf(xml, c.openEnd, c.closeStart).map((k) => key(xml, k)));
  const add = childrenOf(oXml, oc.openEnd, oc.closeStart).filter((k) => !have.has(key(oXml, k)));
  if (!add.length) return;
  for (const k of add) {
    c = findEl(xml, 'w:compat');
    xml = insertOrdered(xml, c, oXml.slice(k.start, k.end), localOf(k.name), COMPAT_ORDER);
  }
  edit.write(part, xml);
  edit.fixes.push(`호환 설정 ${add.length}개 원본대로: ${add.map((k) => key(oXml, k).replace(/^cs:/, '')).join(', ')}`);
}

const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr',
  'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
  'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
  'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const STYLE_ORDER = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed',
  'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'];
// 한국어 Word 기본값: 영어 단어 단위 줄 바꿈 끔, 한글·영문/숫자 사이 자동 간격 끔
const NORMAL_KEEP = ['wordWrap', 'autoSpaceDE', 'autoSpaceDN'];
const isOn = (v) => v === undefined || ['1', 'true', 'on'].includes(String(v).toLowerCase());

function defaultParaStyle(xml) {
  const styles = findAll(xml, 'w:style');
  const isPara = (s) => (s.attrs['w:type'] ?? 'paragraph') === 'paragraph';
  return styles.find((s) => isPara(s) && s.attrs['w:default'] !== undefined && isOn(s.attrs['w:default']))
    ?? styles.find((s) => isPara(s) && !s.self && /<w:name\s+w:val\s*=\s*["']Normal["']/.test(xml.slice(s.openEnd, s.closeStart))) ?? null;
}

/** 기본 문단 스타일(Normal): 원본의 wordWrap·autoSpaceDE·autoSpaceDN이 빠졌으면 되살린다 (한글·숫자 사이 띄어 보임, 단어 단위 줄 바꿈) */
async function carryNormalStyle(edit, orig) {
  const part = await wordPart(edit, 'styles');
  const oPart = await wordPart(orig, 'styles');
  if (!part || !oPart) return;
  const oXml = await orig.xml(oPart);
  const os = defaultParaStyle(oXml);
  const op = os && !os.self ? findEl(oXml, 'w:pPr', os.openEnd, os.closeStart) : null;
  if (!op || op.self) return;
  const want = childrenOf(oXml, op.openEnd, op.closeStart).filter((k) => NORMAL_KEEP.includes(localOf(k.name)) && k.name.startsWith('w:'));
  if (!want.length) return;
  let xml = await edit.xml(part);
  let s = defaultParaStyle(xml);
  if (!s) return;
  ({ xml, el: s } = opened(xml, s));
  let p = findEl(xml, 'w:pPr', s.openEnd, s.closeStart);
  if (!p) {
    xml = insertOrdered(xml, s, '<w:pPr></w:pPr>', 'pPr', STYLE_ORDER);
    s = defaultParaStyle(xml);
    p = findEl(xml, 'w:pPr', s.openEnd, s.closeStart);
  }
  ({ xml, el: p } = opened(xml, p));
  const have = new Set(childrenOf(xml, p.openEnd, p.closeStart).map((k) => k.name));
  const add = want.filter((k) => !have.has(k.name));
  if (!add.length) return;
  for (const k of add) {
    s = defaultParaStyle(xml);
    p = findEl(xml, 'w:pPr', s.openEnd, s.closeStart);
    xml = insertOrdered(xml, p, oXml.slice(k.start, k.end), localOf(k.name), PPR_ORDER);
  }
  edit.write(part, xml);
  edit.fixes.push(`기본 문단 스타일 원본대로: ${add.map((k) => localOf(k.name)).join(', ')}`);
}

// 내용 컨트롤 종류 (없으면 서식 있는 텍스트)
const SDT_TYPES = ['w:text', 'w:dropDownList', 'w:comboBox', 'w:date', 'w14:checkbox'];
const SDT_ANY_TYPE = [...SDT_TYPES, 'w:richText', 'w:docPartObj', 'w:docPartList', 'w:picture', 'w:equation', 'w:citation', 'w:group',
  'w:bibliography', 'w15:repeatingSection', 'w15:repeatingSectionItem', 'w15:webExtensionLinked', 'w15:webExtensionCreated'];

async function sdtTypes(e) {
  const out = new Map();
  for (const part of await wordStories(e)) {
    const xml = await e.xml(part);
    if (!xml.includes('<w:sdtPr')) continue;
    for (const pr of findAll(xml, 'w:sdtPr')) {
      if (pr.self) continue;
      const id = findEl(xml, 'w:id', pr.openEnd, pr.closeStart)?.attrs['w:val'];
      if (id === undefined) continue;
      const kids = childrenOf(xml, pr.openEnd, pr.closeStart);
      const t = kids.find((k) => SDT_TYPES.includes(k.name));
      if (t) out.set(id, xml.slice(t.start, t.end));
    }
  }
  return out;
}

/** 내용 컨트롤: 원본이 일반 텍스트(w:text)·목록·날짜·확인란이었는데 종류가 빠졌으면 (서식 있는 텍스트가 됨) 되살린다 */
async function carrySdtTypes(edit, orig) {
  const before = await sdtTypes(orig);
  if (!before.size) return;
  let n = 0;
  for (const part of await wordStories(edit)) {
    let xml = await edit.xml(part);
    if (!xml.includes('<w:sdtPr')) continue;
    const w14 = /\sxmlns:w14\s*=/.test(xml.slice(0, 4096));
    let changed = false;
    for (let at = 0, pr; (pr = findEl(xml, 'w:sdtPr', at)); at = pr.end) {
      if (pr.self) continue;
      const id = findEl(xml, 'w:id', pr.openEnd, pr.closeStart)?.attrs['w:val'];
      const want = id === undefined ? null : before.get(id);
      if (!want || (want.startsWith('<w14:') && !w14)) continue;
      if (childrenOf(xml, pr.openEnd, pr.closeStart).some((k) => SDT_ANY_TYPE.includes(k.name))) continue;
      xml = splice(xml, pr.closeStart, pr.closeStart, want);
      pr = { ...pr, end: pr.end + want.length };
      changed = true; n++;
    }
    if (changed) edit.write(part, xml);
  }
  if (n) edit.fixes.push(`내용 컨트롤 종류 ${n}개 원본대로`);
}

// pptx: 구역
const SECTIONS_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
function sldIds(xml) {
  const l = findEl(xml, 'p:sldIdLst');
  return l && !l.self ? findAll(xml, 'p:sldId', l.openEnd, l.closeStart).map((s) => s.attrs.id).filter(Boolean) : [];
}

/** 구역(p14:sectionLst): x2t가 지운다 → 원본 구역을 지금 슬라이드 순서에 맞춰 되살린다 (새 슬라이드는 앞 슬라이드의 구역으로) */
async function carrySections(edit, orig) {
  const part = await mainPart(edit);
  const oPart = await mainPart(orig);
  if (!part || !oPart || !edit.has(part) || !orig.has(oPart)) return;
  let xml = await edit.xml(part);
  if (xml.includes('sectionLst')) return;
  const oXml = await orig.xml(oPart);
  const ext = findAll(oXml, 'p:ext').find((x) => x.attrs.uri === SECTIONS_URI);
  if (!ext || ext.self) return;
  const extXml = oXml.slice(ext.start, ext.end);
  const lst = /<(\w+):sectionLst\b/.exec(extXml);
  if (!lst) return;
  const px = lst[1];
  const sl = findEl(extXml, `${px}:sectionLst`);
  let head = extXml.slice(0, sl.openEnd);
  if (!new RegExp(`\\sxmlns:${reEsc(px)}\\s*=`).test(extXml)) {
    // 원본이 접두어를 뿌리에 선언했으면 구역 목록에 옮겨 적는다
    const rootTag = findEl(oXml, 'p:presentation')?.tag ?? '';
    const uri = attrsOf(rootTag)[`xmlns:${px}`];
    if (!uri) return;
    head = extXml.slice(0, sl.start) + sl.tag.replace(/\s*>$/, ` xmlns:${px}="${escapeAttr(uri)}">`);
  }
  const sections = findAll(extXml, `${px}:section`, sl.openEnd, sl.closeStart).map((s) => ({
    tag: s.tag.replace(/\s*\/>$/, '>'),
    ids: s.self ? [] : findAll(extXml, `${px}:sldId`, s.openEnd, s.closeStart).map((x) => x.attrs.id),
  }));
  const owner = new Map();
  sections.forEach((s, i) => s.ids.forEach((id) => owner.set(id, i)));
  const ids = sldIds(xml);
  if (!sections.length || !ids.some((id) => owner.has(id))) { edit.fixes.push('구역: 원본 슬라이드와 겹치지 않아 되살리지 않음'); return; }
  // 슬라이드마다 구역: 원본에 있던 것은 그 구역, 새 것은 앞 슬라이드의 구역 (맨 앞이면 다음 아는 슬라이드의 구역)
  const firstKnown = owner.get(ids.find((id) => owner.has(id)));
  let cur = firstKnown;
  const of = ids.map((id) => (cur = owner.has(id) ? owner.get(id) : cur));
  for (let i = 1; i < of.length; i++) {
    if (of[i] < of[i - 1]) { edit.fixes.push('구역: 슬라이드 순서가 구역을 넘나들어 되살리지 않음'); return; }
  }
  const lists = sections.map(() => []);
  ids.forEach((id, i) => lists[of[i]].push(id));
  const body = sections.map((s, i) => `${s.tag}<${px}:sldIdLst>${lists[i].map((id) => `<${px}:sldId id="${escapeAttr(id)}"/>`).join('')}</${px}:sldIdLst></${px}:section>`).join('');
  const newExt = head + body + extXml.slice(sl.closeStart);
  const el = findEl(xml, 'p:extLst');
  if (el && !el.self) xml = splice(xml, el.closeStart, el.closeStart, newExt);
  else if (el) xml = splice(xml, el.start, el.end, `<p:extLst>${newExt}</p:extLst>`);
  else {
    const root = findEl(xml, 'p:presentation');
    if (!root || root.self) return;
    xml = splice(xml, root.closeStart, root.closeStart, `<p:extLst>${newExt}</p:extLst>`);
  }
  edit.write(part, xml);
  edit.fixes.push(`구역 ${sections.length}개 원본대로`);
}

// xlsx
async function workbookParts(e, type) {
  const book = await mainPart(e);
  if (!book || !e.has(book)) return [];
  return relParts(e, book, type);
}
async function sheetMap(e) {
  // 시트 이름 → 부분
  const book = await mainPart(e);
  if (!book || !e.has(book)) return new Map();
  const xml = await e.xml(book);
  const rels = new Map((await relParts(e, book, 'worksheet')).map((r) => [r.id, r.part]));
  const out = new Map();
  for (const m of xml.matchAll(/<(?:\w+:)?sheet\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/g)) {
    const a = attrsOf(m[0]);
    const part = rels.get(a['r:id']);
    if (a.name && part) out.set(a.name, part);
  }
  return out;
}

/** 셀 글꼴: x2t가 테마 글꼴(scheme)을 쓰는 글꼴 이름을 테마 라틴 글꼴(Calibri)로 바꾼다 → 원본 이름(맑은 고딕)·charset으로 */
async function carryCellFonts(edit, orig) {
  const [st] = await workbookParts(edit, 'styles');
  const [ost] = await workbookParts(orig, 'styles');
  const [th] = await workbookParts(edit, 'theme');
  if (!st || !ost || !th) return;
  const theme = themeFonts(await edit.xml(th.part));
  if (!theme) return;
  const oXml = await orig.xml(ost.part);
  const ofs = findEl(oXml, 'fonts');
  if (!ofs || ofs.self) return;
  // 원본에서 scheme마다 처음 나오는 글꼴 = 그 scheme의 이름
  const rep = {};
  for (const f of findAll(oXml, 'font', ofs.openEnd, ofs.closeStart)) {
    if (f.self) continue;
    const scheme = findEl(oXml, 'scheme', f.openEnd, f.closeStart)?.attrs.val;
    const name = findEl(oXml, 'name', f.openEnd, f.closeStart)?.attrs.val;
    const charset = findEl(oXml, 'charset', f.openEnd, f.closeStart);
    if ((scheme === 'minor' || scheme === 'major') && name && !rep[scheme]) rep[scheme] = { name, charset: charset ? oXml.slice(charset.start, charset.end) : null };
  }
  const latin = { minor: theme.minor, major: theme.major };
  if (!['minor', 'major'].some((k) => rep[k] && latin[k] && rep[k].name !== latin[k])) return;
  let xml = await edit.xml(st.part);
  const fonts = findEl(xml, 'fonts');
  if (!fonts || fonts.self) return;
  let n = 0;
  let out = xml.slice(fonts.openEnd, fonts.closeStart);
  const list = findAll(out, 'font');
  for (const f of list.reverse()) {
    if (f.self) continue;
    const scheme = findEl(out, 'scheme', f.openEnd, f.closeStart)?.attrs.val;
    const r = rep[scheme];
    const name = findEl(out, 'name', f.openEnd, f.closeStart);
    if (!r || !name || name.attrs.val !== latin[scheme] || r.name === latin[scheme]) continue;
    let tag = `<name val="${escapeAttr(r.name)}"/>`;
    if (r.charset && !findEl(out, 'charset', f.openEnd, f.closeStart)) tag += r.charset;
    out = splice(out, name.start, name.end, tag);
    n++;
  }
  if (!n) return;
  xml = xml.slice(0, fonts.openEnd) + out + xml.slice(fonts.closeStart);
  edit.write(st.part, xml);
  edit.fixes.push(`셀 글꼴 ${n}개 이름 원본대로: ${['minor', 'major'].filter((k) => rep[k]).map((k) => rep[k].name).join(', ')}`);
}

/** 메모: 원본에 스레드 메모가 없었는데 x2t가 메모를 스레드 메모로 바꿨으면 → 원래 메모(작성자·글)로 되돌리고 스레드 메모 부분을 지운다 */
async function carryNotes(edit, orig, types) {
  const isThreaded = (n) => /^xl\/threadedComments\//i.test(n);
  if ([...orig.names.values()].some((n) => orig.has(n) && isThreaded(n))) return; // 원본에도 스레드 메모가 있으면 그대로
  const threadedParts = [...edit.names.values()].filter((n) => edit.has(n) && isThreaded(n) && /\.xml$/i.test(n));
  if (!threadedParts.length) return;
  // 사람 목록
  const persons = await workbookParts(edit, 'person');
  const who = new Map();
  for (const p of persons) {
    for (const m of (await edit.xml(p.part)).matchAll(/<(?:\w+:)?person\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/g)) {
      const a = attrsOf(m[0]);
      if (a.id) who.set(a.id, a.displayName ?? '');
    }
  }
  // 스레드 메모: id → 작성자 이름 (답글이 있으면 되돌리지 않음 — 답글은 원래 메모로 나타낼 수 없다)
  const thread = new Map();
  for (const n of threadedParts) {
    for (const m of (await edit.xml(n)).matchAll(/<(?:\w+:)?threadedComment\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/g)) {
      const a = attrsOf(m[0]);
      if (a.parentId) { edit.fixes.push('메모: 답글이 있어 스레드 메모를 그대로 둠'); return; }
      if (a.id) thread.set(a.id.toUpperCase(), who.get(a.personId) ?? '');
    }
  }
  // 원본 메모: 시트 이름 → (셀 → { author, text })
  const oSheets = await sheetMap(orig);
  const origNotes = new Map();
  for (const [name, sp] of oSheets) {
    const [c] = await relParts(orig, sp, 'comments');
    if (!c) continue;
    const x = await orig.xml(c.part);
    const authors = findAll(x, 'author').map((a) => (a.self ? '' : unescapeText(x.slice(a.openEnd, a.closeStart))));
    const m = new Map();
    for (const cm of findAll(x, 'comment')) {
      const t = cm.self ? null : findEl(x, 'text', cm.openEnd, cm.closeStart);
      if (cm.attrs.ref && t) m.set(cm.attrs.ref, { author: authors[Number(cm.attrs.authorId)] ?? '', text: x.slice(t.start, t.end) });
    }
    origNotes.set(name, m);
  }
  const sheets = await sheetMap(edit);
  let n = 0;
  for (const [name, sp] of sheets) {
    const relsPath = relsPathOf(sp);
    let rels = await edit.xml(relsPath);
    if (!rels) continue;
    // 시트 → 스레드 메모 관계 지우기
    for (const r of parseRels(rels, sp).filter((r) => /\/threadedComment$/.test(r.type)).reverse()) rels = rels.slice(0, r.index) + rels.slice(r.index + r.tag.length);
    if (rels !== await edit.xml(relsPath)) edit.write(relsPath, rels);
    const [c] = await relParts(edit, sp, 'comments');
    if (!c) continue;
    let x = await edit.xml(c.part);
    const before = origNotes.get(name) ?? new Map();
    const authors = findAll(x, 'author').map((a) => (a.self ? '' : unescapeText(x.slice(a.openEnd, a.closeStart))));
    const names = [];
    const idOf = (s) => { let i = names.indexOf(s); if (i < 0) { names.push(s); i = names.length - 1; } return i; };
    const list = findEl(x, 'commentList');
    if (!list || list.self) continue;
    let body = '';
    let at = list.openEnd;
    for (const cm of findAll(x, 'comment', list.openEnd, list.closeStart)) {
      body += x.slice(at, cm.start);
      at = cm.end;
      let author = authors[Number(cm.attrs.authorId)] ?? '';
      const tc = /^tc=(\{[0-9a-f-]+\})$/i.exec(author);
      if (tc) author = thread.get(tc[1].toUpperCase()) ?? thread.get(String(cm.attrs['xr:uid'] ?? '').toUpperCase()) ?? '';
      const t = cm.self ? null : findEl(x, 'text', cm.openEnd, cm.closeStart);
      let text = t ? x.slice(t.start, t.end) : '<text/>';
      if (t && tc) text = stripThreadPrefix(text, author);
      const o = before.get(cm.attrs.ref);
      if (o && t && textOf(o.text) === textOf(text)) { text = o.text; author = o.author; }
      const open = cm.tag.replace(/\s*\/>$/, '>').replace(/(\sauthorId\s*=\s*)(?:"[^"]*"|'[^']*')/, (_, p) => `${p}"${idOf(author)}"`);
      body += `${open}${text}</comment>`;
      n++;
    }
    body += x.slice(at, list.closeStart);
    const au = findEl(x, 'authors');
    const authorsXml = `<authors>${names.map((s) => `<author>${escapeText(s)}</author>`).join('')}</authors>`;
    x = x.slice(0, list.openEnd) + body + x.slice(list.closeStart);
    x = au ? splice(x, au.start, au.end, authorsXml) : x;
    edit.write(c.part, x);
  }
  // 스레드 메모·사람 부분과 그 관계·내용 형식 지우기
  const book = await mainPart(edit);
  let wrels = await edit.xml(relsPathOf(book));
  for (const r of parseRels(wrels, book).filter((r) => /\/person$/.test(r.type)).reverse()) wrels = wrels.slice(0, r.index) + wrels.slice(r.index + r.tag.length);
  edit.write(relsPathOf(book), wrels);
  for (const part of [...threadedParts, ...persons.map((p) => p.part)]) {
    edit.remove(part);
    if (edit.has(relsPathOf(part))) edit.remove(relsPathOf(part));
    types.remove(part);
  }
  edit.fixes.push(`스레드 메모 ${thread.size}개를 원래 메모로 되돌림 (메모 ${n}개)`);
}

// x2t가 원래 메모에 붙이는 '<작성자>:' 줄과 앞뒤 줄 바꿈을 뗀다
function stripThreadPrefix(text, author) {
  const t = findEl(text, 'text');
  const runs = findAll(text, 'r', t.openEnd, t.closeStart);
  if (runs.length < 2 || !author) return text;
  const first = text.slice(runs[0].start, runs[0].end);
  if (textOf(first).trim() !== `${author}:`) return text;
  let rest = text.slice(runs[1].start, t.closeStart);
  const tt = findEl(rest, 't');
  if (tt && !tt.self) {
    const s = unescapeText(rest.slice(tt.openEnd, tt.closeStart)).replace(/^\r?\n/, '');
    const last = runs.length === 2;
    const v = last ? s.replace(/\r?\n$/, '') : s;
    rest = splice(rest, tt.openEnd, tt.closeStart, escapeText(v));
  }
  if (runs.length > 2) {
    // 마지막 run 끝 줄 바꿈
    const lastRun = findAll(rest, 'r').at(-1);
    const lt = lastRun && findEl(rest, 't', lastRun.openEnd, lastRun.closeStart);
    if (lt && !lt.self) rest = splice(rest, lt.openEnd, lt.closeStart, escapeText(unescapeText(rest.slice(lt.openEnd, lt.closeStart)).replace(/\r?\n$/, '')));
  }
  return text.slice(0, t.openEnd) + rest + text.slice(t.closeStart);
}

// x14 조건부 서식 → 기본 조건부 서식 (Excel 2007 형식으로 나타낼 수 있는 것만: 수식·서식만 있는 규칙, 다른 시트 참조 없음)
const CF_MAIN_TYPES = new Set(['cellIs', 'expression', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'top10', 'aboveAverage',
  'timePeriod', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'duplicateValues', 'uniqueValues']);
const X14_CF_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
const SHEET_AFTER_CF = ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture',
  'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

async function mainConditionalFormats(edit) {
  const [st] = await workbookParts(edit, 'styles');
  if (!st) return;
  const sheets = await workbookParts(edit, 'worksheet');
  const dxfs = [];
  let styles = await edit.xml(st.part);
  const dx = findEl(styles, 'dxfs');
  let next = dx && !dx.self ? findAll(styles, 'dxf', dx.openEnd, dx.closeStart).length : 0;
  let moved = 0;
  for (const { part } of sheets) {
    let xml = await edit.xml(part);
    if (!xml.includes('x14:conditionalFormattings')) continue;
    const root = findEl(xml, 'worksheet');
    if (!root || root.self) continue;
    const top = childrenOf(xml, root.openEnd, root.closeStart).find((k) => k.name === 'extLst');
    const exts = top ? findEl(xml, 'extLst', top.start) : null;
    if (!exts || exts.self) continue;
    const ext = findAll(xml, 'ext', exts.openEnd, exts.closeStart).find((e) => e.attrs.uri?.toUpperCase() === X14_CF_URI.toUpperCase());
    if (!ext || ext.self) continue;
    const cfs = findEl(xml, 'x14:conditionalFormattings', ext.openEnd, ext.closeStart);
    if (!cfs || cfs.self) continue;
    const mains = [];
    let inner = '';
    let at = cfs.openEnd;
    for (const cf of findAll(xml, 'x14:conditionalFormatting', cfs.openEnd, cfs.closeStart)) {
      inner += xml.slice(at, cf.start);
      at = cf.end;
      const whole = xml.slice(cf.start, cf.end);
      const sq = cf.self ? null : findEl(xml, 'xm:sqref', cf.openEnd, cf.closeStart);
      const sqref = sq ? unescapeText(xml.slice(sq.openEnd, sq.closeStart)).trim() : '';
      if (!sqref || cf.attrs.pivot !== undefined) { inner += whole; continue; }
      const keep = [];
      const rules = [];
      for (const r of findAll(xml, 'x14:cfRule', cf.openEnd, cf.closeStart)) {
        const kids = r.self ? [] : childrenOf(xml, r.openEnd, r.closeStart);
        const fs = kids.filter((k) => k.name === 'xm:f').map((k) => xml.slice(k.start, k.end).replace(/^<xm:f\b[^>]*>|<\/xm:f>$/g, ''));
        const dxf = kids.find((k) => k.name === 'x14:dxf');
        const plain = kids.every((k) => k.name === 'xm:f' || k.name === 'x14:dxf');
        const local = fs.every((f) => !/[![]/.test(unescapeText(f).replace(/"[^"]*"/g, '')));
        if (!CF_MAIN_TYPES.has(r.attrs.type) || !plain || !local || kids.filter((k) => k.name === 'x14:dxf').length > 1) { keep.push(xml.slice(r.start, r.end)); continue; }
        let dxfId = '';
        if (dxf) {
          const dx = findEl(xml, 'x14:dxf', dxf.start);
          const d = dx.self ? '<dxf/>' : `<dxf>${xml.slice(dx.openEnd, dx.closeStart)}</dxf>`;
          dxfs.push(d);
          dxfId = ` dxfId="${next++}"`;
        }
        const attrs = Object.entries(r.attrs).filter(([k]) => k !== 'id' && k !== 'type' && !k.includes(':'))
          .map(([k, v]) => ` ${k}="${escapeAttr(v)}"`).join('');
        rules.push(`<cfRule type="${escapeAttr(r.attrs.type)}"${dxfId}${attrs}>${fs.map((f) => `<formula>${f}</formula>`).join('')}</cfRule>`);
      }
      if (rules.length) mains.push(`<conditionalFormatting sqref="${escapeAttr(sqref)}">${rules.join('')}</conditionalFormatting>`);
      if (!rules.length) { inner += whole; continue; }
      moved += rules.length;
      if (keep.length) {
        const ruleEls = findAll(xml, 'x14:cfRule', cf.openEnd, cf.closeStart);
        inner += xml.slice(cf.start, ruleEls[0].start) + keep.join('') + xml.slice(ruleEls.at(-1).end, cf.end);
      }
    }
    if (!mains.length) continue;
    inner += xml.slice(at, cfs.closeStart);
    // 확장 안 정리: 비면 x14:conditionalFormattings → ext → extLst 차례로 뺀다
    let tail;
    if (inner.trim()) tail = xml.slice(exts.start, cfs.openEnd) + inner + xml.slice(cfs.closeStart, exts.end);
    else {
      const extInner = xml.slice(ext.openEnd, cfs.start) + xml.slice(cfs.end, ext.closeStart);
      const restExts = xml.slice(exts.openEnd, ext.start) + (extInner.trim() ? xml.slice(ext.start, ext.openEnd) + extInner + xml.slice(ext.closeStart, ext.end) : '') + xml.slice(ext.end, exts.closeStart);
      tail = restExts.trim() ? xml.slice(exts.start, exts.openEnd) + restExts + xml.slice(exts.closeStart, exts.end) : '';
    }
    xml = xml.slice(0, exts.start) + tail + xml.slice(exts.end);
    // 기본 조건부 서식 자리: 마지막 기본 conditionalFormatting 뒤, 없으면 그 뒤에 오는 요소 앞
    const r2 = findEl(xml, 'worksheet');
    const kids = childrenOf(xml, r2.openEnd, r2.closeStart);
    const lastCf = kids.filter((k) => k.name === 'conditionalFormatting').at(-1);
    const after = kids.find((k) => SHEET_AFTER_CF.includes(k.name));
    const pos = lastCf ? lastCf.end : after ? after.start : r2.closeStart;
    xml = splice(xml, pos, pos, mains.join(''));
    edit.write(part, xml);
  }
  if (!moved) return;
  // 서식(dxf)을 styles.xml <dxfs> 끝에
  const d = findEl(styles, 'dxfs');
  if (d) {
    const total = (d.self ? 0 : findAll(styles, 'dxf', d.openEnd, d.closeStart).length) + dxfs.length;
    const open = d.tag.replace(/\s*\/?>$/, '>').replace(/(\scount\s*=\s*)(?:"[^"]*"|'[^']*')/, `$1"${total}"`);
    const body = d.self ? dxfs.join('') : styles.slice(d.openEnd, d.closeStart) + dxfs.join('');
    styles = splice(styles, d.start, d.end, `${/\scount\s*=/.test(open) ? open : open.replace(/>$/, ` count="${total}">`)}${body}</dxfs>`);
  } else {
    const ss = findEl(styles, 'styleSheet');
    if (!ss || ss.self) fail('styles.xml 모양이 다름');
    styles = insertOrdered(styles, ss, `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`, 'dxfs',
      ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']);
  }
  if (dxfs.length) edit.write(st.part, styles);
  edit.fixes.push(`확장에만 있던 조건부 서식 ${moved}개를 기본 조건부 서식으로`);
}

// 원본이 있을 때 · 없을 때 하는 단계 (조건부 서식 되돌리기는 원본과 상관없이)
const CARRY = {
  docx: [['테마 글꼴', carryThemeFonts], ['글꼴 언어', carryThemeFontLang], ['기본 문단 스타일', carryNormalStyle], ['호환 설정', carryCompat], ['내용 컨트롤', carrySdtTypes]],
  xlsx: [['테마 글꼴', carryThemeFonts], ['셀 글꼴', carryCellFonts], ['메모', carryNotes], ['조건부 서식', mainConditionalFormats]],
  pptx: [['테마 글꼴', carryThemeFonts], ['구역', carrySections]],
};
const NO_ORIGINAL = { docx: [['글꼴 언어', carryThemeFontLang]], xlsx: [['조건부 서식', mainConditionalFormats]], pptx: [] };

/** 원본에서 옮겨 오기 (단계마다 실패하면 그 단계만 되돌림) → 경고 목록 */
async function carryOver(edit, types, ext, orig) {
  const warnings = [];
  const steps = orig ? CARRY[ext] : NO_ORIGINAL[ext];
  for (const [label, step] of steps) {
    const snap = edit.snapshot(types);
    try { await step(edit, orig, types); } catch (err) {
      edit.restore(snap, types);
      warnings.push(`원본에서 옮기기(${label}) 실패: ${err?.message ?? err}`);
    }
  }
  return warnings;
}

// ── 공통: 고칠 것 찾기 · 고치기 · 확인 ────────────────────────

async function issuesIn(edit, ext) {
  const types = await readTypes(edit.zip, edit.budget);
  if (ext === 'pptx') return (await sharedThemes(edit, types)).map((s) => `${s.kind}가 ${s.ownerKind}와 테마를 같이 씀: ${s.theme} (${s.master}, ${s.owner})`);
  if (ext === 'xlsx') return (await unusedXlsxThemes(edit)).map((n) => `쓰지 않는 테마: ${n}`);
  const out = [];
  for (const part of await wordStories(edit)) {
    const xml = await edit.xml(part);
    const n = xml.includes('w:vMerge') ? dropLoneVMergeRestarts(xml).removed : 0;
    if (n) out.push(`혼자 있는 세로 병합 시작 ${n}개: ${part}`);
  }
  return out;
}

const normExt = (ext) => String(ext ?? '').replace(/^\./, '').toLowerCase();
const toBuffer = (b) => (Buffer.isBuffer(b) ? b : b instanceof ArrayBuffer ? Buffer.from(b) : ArrayBuffer.isView(b) ? Buffer.from(b.buffer, b.byteOffset, b.byteLength) : null);
const sameBytes = (a, b) => a.length === b.length && Buffer.from(a.buffer, a.byteOffset, a.byteLength).equals(Buffer.from(b.buffer, b.byteOffset, b.byteLength));

/**
 * 고칠 것이 남아 있는지만 본다 (고치지 않음). 시험·검증용.
 * @param {Uint8Array|ArrayBuffer} bytes
 * @param {string} ext  'docx' | 'xlsx' | 'pptx'
 * @returns {Promise<string[]>}  남은 문제 (없으면 빈 배열). 꾸러미를 못 읽으면 예외
 */
export async function findOoxmlIssues(bytes, ext) {
  const e = normExt(ext);
  if (!FIXUP_EXTS.includes(e)) return [];
  return issuesIn(new Edit(await JSZip.loadAsync(toBuffer(bytes))), e);
}

// 손대지 않은 부분은 압축된 그대로 옮겨지므로 CRC·크기로 비교한다 (값을 모르면 풀어서 비교)
const metaOf = (f) => (typeof f?._data?.crc32 === 'number' && typeof f._data.uncompressedSize === 'number'
  ? `${f._data.crc32 >>> 0}:${f._data.uncompressedSize}` : null);

async function verify(edit, out, ext) {
  const zip = await JSZip.loadAsync(out);
  const outNames = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const want = [...edit.names.values()].filter((n) => edit.has(n));
  if (outNames.length !== want.length || want.some((n) => !zip.files[n] || zip.files[n].dir)) fail('고친 꾸러미의 부분 목록이 다름');
  for (const n of want) {
    const f = zip.file(n);
    if (edit.put.has(n)) {
      const b = await f.async('uint8array');
      if (!sameBytes(b, edit.put.get(n))) fail(`고친 부분이 다르게 들어감: ${n}`);
      if (/\.(xml|rels)$/i.test(n) && !xmlBalanced(new TextDecoder('utf-8', { ignoreBOM: true }).decode(b))) fail(`고친 XML이 깨짐: ${n}`);
      continue;
    }
    const a = metaOf(edit.zip.file(n)); const b = metaOf(f);
    if (a && b) { if (a !== b) fail(`손대지 않은 부분이 바뀜: ${n}`); continue; }
    if (!sameBytes(await partBytes(edit.zip.file(n), n, edit.budget), await partBytes(f, n, null))) fail(`손대지 않은 부분이 바뀜: ${n}`);
  }
  // 모든 관계가 있는 부분을 가리키는지 (고친 관계 파일만)
  const check = new Edit(zip);
  for (const n of edit.put.keys()) {
    const source = sourceOfRels(n);
    if (source == null) continue;
    for (const r of parseRels(await check.xml(n), source)) if (r.part && !check.has(r.part)) fail(`관계가 없는 부분을 가리킴: ${n} → ${r.target}`);
  }
  const left = await issuesIn(check, ext);
  if (left.length) fail(`고친 뒤에도 남음: ${left.join(', ')}`);
}

/**
 * 편집기가 내보낸 OOXML 바이트를 저장 직전에 고친다. 저장할 바이트에만 쓴다 (입력은 바꾸지 않음).
 * 어떤 실패든 원래 바이트를 그대로 돌려주고 error를 채운다 — 저장은 막지 않는다.
 * 원본에서 옮기기가 실패하면 그것만 빼고 다시 고친다 (warnings에 남김).
 * @param {Uint8Array|ArrayBuffer} bytes  편집기가 내보낸 꾸러미
 * @param {string} ext  'docx' | 'xlsx' | 'pptx' ('.PPTX'도 됨). 다른 형식은 그대로 돌려준다
 * @param {{ original?: Uint8Array|ArrayBuffer|null }} [o]  덮어쓸(편집기가 열었던) 원래 파일 — 주면 x2t가 잃는 것을 옮겨 온다
 * @returns {Promise<{ bytes: Buffer, changed: boolean, fixes: string[], warnings: string[], error: string | null, notify: boolean }>}
 *   bytes    저장할 바이트 (고쳤으면 새 꾸러미, 아니면 입력과 같은 내용)
 *   changed  고쳤으면 true
 *   fixes    고친 것 (한국어 한 줄씩, 로그용)
 *   warnings 원본에서 옮기기 일부를 못 한 것 (로그용 — 저장·필수 고치기는 됨)
 *   error    고치기 실패 (한국어 한 문장, 로그·알림 띠용). 실패해도 bytes(원래 것)는 저장해도 된다
 *   notify   true면 알림 띠로 알린다 (pptx 실패: PowerPoint가 복구 창을 띄울 수 있음)
 */
export async function fixOoxml(bytes, ext, { original = null } = {}) {
  const e = normExt(ext);
  const input = toBuffer(bytes);
  if (!input) throw new TypeError('bytes는 Uint8Array·ArrayBuffer여야 합니다');
  const done = (o) => ({ bytes: input, changed: false, fixes: [], warnings: [], error: null, notify: false, ...o });
  if (!FIXUP_EXTS.includes(e)) return done({});
  const warnings = [];
  let orig = null;
  if (original != null) {
    try {
      const ob = toBuffer(original);
      if (!ob) throw new TypeError('바이트가 아님');
      orig = new Edit(await JSZip.loadAsync(ob));
    } catch (err) { warnings.push(`원본을 읽지 못해 원본에서 옮기기는 건너뜀: ${err?.message ?? err}`); }
  }
  try {
    let r;
    try { r = await fixOnce(input, e, orig, true); } catch (err) {
      // 원본 옮기기·조건부 서식이 확인을 통과하지 못하면 필수 고치기만 다시
      warnings.push(`원본에서 옮기기를 빼고 다시 고침: ${err?.message ?? err}`);
      r = await fixOnce(input, e, null, false);
    }
    warnings.push(...r.warnings);
    return r.changed ? done({ bytes: r.bytes, changed: true, fixes: r.fixes, warnings }) : done({ warnings });
  } catch (err) {
    const cause = err?.message ?? String(err);
    const error = e === 'pptx'
      ? `PowerPoint에서 경고 없이 열리도록 고치지 못해 편집기가 만든 그대로 저장했습니다. PowerPoint에서 열 때 복구 창이 뜰 수 있습니다 (${cause})`
      : `${LABEL[e]}에서 경고 없이 열리도록 고치지 못해 편집기가 만든 그대로 저장했습니다 (${cause})`;
    return done({ error, notify: e === 'pptx', warnings });
  }
}

async function fixOnce(input, e, orig, carry) {
  const edit = new Edit(await JSZip.loadAsync(input));
  const types = await readTypes(edit.zip, edit.budget);
  const warnings = carry ? await carryOver(edit, types, e, orig) : [];
  if (e === 'pptx') await splitSharedThemes(edit, types);
  else if (e === 'xlsx') await dropUnusedThemes(edit, types);
  else await dropLoneRestarts(edit);
  if (!edit.changed) return { changed: false, warnings };
  if (types.changed) edit.write('[Content_Types].xml', types.xml);
  for (const n of edit.removed) edit.zip.remove(n);
  for (const [n, data] of edit.put) edit.zip.file(n, data, { binary: true, createFolders: false });
  const out = await edit.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  // 확인은 원래 꾸러미와 비교하므로 새로 읽은 것으로
  const before = new Edit(await JSZip.loadAsync(input));
  before.put = edit.put; before.removed = edit.removed; before.names = new Map([...edit.names]);
  await verify(before, out, e);
  return { changed: true, bytes: out, fixes: edit.fixes, warnings };
}
