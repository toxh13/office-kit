// SPDX-License-Identifier: AGPL-3.0-only
// allfonts.mjs — 설치된 글꼴로 편집기의 AllFonts.js(sdkjs 9.4, __all_fonts_js_version__ = 2)를 만든다.
//   글꼴 목록(__fonts_files·__fonts_infos), 유니코드 대체 범위(__fonts_ranges), 글꼴 선택 이진(g_fonts_selection_bin).
//
// 이 파일은 ONLYOFFICE core의 AllFontsGen 규칙을 JavaScript로 옮기고 고친 것이다 (수정판):
//   DesktopEditor/fontengine/ApplicationFontsWorker.cpp (SaveAllFontsJS·CSymbolSimpleChecker2·SaveWebFonts)
//   DesktopEditor/fontengine/ApplicationFonts.cpp (CFontList::Add·ToBuffer)
//   원본: Copyright (C) Ascensio System SIA, 2009-2026, AGPL-3.0-only + ONLYOFFICE 추가 조항. https://github.com/ONLYOFFICE/core/tree/v9.4.0.129
//   라이선스 본문: LICENSES/AGPL-3.0.txt, 추가 조항: LICENSES/ONLYOFFICE-additional-terms.txt (저장소 뿌리)
//   수정: 2026-10-18 Office Kit (toxh13) — FreeType 대신 직접 읽기(sfnt.mjs), 한국어 이름 별칭, 한글 대체 글꼴을 맑은 고딕으로,
//         선택 이진의 경로를 웹 파일 번호로 (sdkjs ListMap이 맞게). profile: 'onlyoffice'면 원본과 같은 결과를 낸다.
import { readdirSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { faceOffsets, freetypeName, openFileSource, readFace } from './sfnt.mjs';

export const ALL_FONTS_JS_VERSION = 2;
/** 생성 규칙이 바뀌면 올린다 (캐시 키에 들어감) */
export const GENERATOR_VERSION = 1;
/** AllFontsGen이 웹 글꼴 앞 32바이트에 XOR하는 열쇠 (sdkjs guidOdttf) */
export const FONT_XOR_KEY = Buffer.from([0xA0, 0x66, 0xD6, 0x20, 0x14, 0x96, 0x47, 0xFA, 0x95, 0x69, 0xB8, 0x50, 0xB0, 0x41, 0x49, 0x48]);
/** 100MB 넘는 글꼴은 AllFontsGen도 건너뛴다 */
export const MAX_FONT_FILE_SIZE = 100000000;

// AllFontsGen SET_FONT_PRIORITY 순서 (ApplicationFontsWorker.cpp, _MAC 부분 제외)
export const ONLYOFFICE_PRIORITY = [
  'ASCW3', 'Arial', 'Times New Roman', 'Tahoma', 'Cambria', 'Calibri', 'Verdana', 'Georgia', 'Open Sans', 'Liberation Sans',
  'Helvetica', 'Nimbus Sans L', 'DejaVu Sans', 'Liberation Serif', 'Trebuchet MS', 'Courier New', 'Carlito', 'Segoe UI',
  'SimSun', 'MS Gothic', 'Nirmala UI', 'Batang', 'MS Mincho', 'Wingdings', 'Microsoft JhengHei', 'Microsoft JhengHei UI',
  'Microsoft YaHei', 'PMingLiU', 'MingLiU', 'DFKai-SB', 'FangSong', 'KaiTi', 'SimKai', 'SimHei', 'Meiryo',
  'Malgun Gothic', 'Nanum Gothic', 'NanumGothic', 'Noto Sans KR', 'TakaoGothic',
];
// Office Kit: 한글은 맑은 고딕, 한자는 바탕(한국식 자형)이 중국·일본 글꼴보다 먼저
export const OFFICEKIT_PRIORITY = (() => {
  const l = ONLYOFFICE_PRIORITY.filter((n) => n !== 'Malgun Gothic' && n !== 'Batang');
  l.splice(l.indexOf('SimSun'), 0, 'Malgun Gothic', 'Batang');
  return l;
})();

/**
 * 한국어 이름 별칭: [한국어 이름, 영어 가족 후보...]. 설치된 글꼴 어디에도 그 한국어 이름이 없을 때만
 * 처음 찾은 후보 가족의 선택 레코드 names에 덧붙인다 → 문서의 '함초롬바탕' 등이 이름으로 찾아진다.
 */
export const KOREAN_ALIASES = [
  ['맑은 고딕', 'Malgun Gothic'], ['맑은고딕', 'Malgun Gothic'],
  ['바탕', 'Batang'], ['바탕체', 'BatangChe'], ['궁서', 'Gungsuh'], ['궁서체', 'GungsuhChe'],
  ['굴림', 'Gulim'], ['굴림체', 'GulimChe'], ['돋움', 'Dotum'], ['돋움체', 'DotumChe'],
  ['새굴림', 'New Gulim', 'Gulim'],
  ['함초롬바탕', 'HCR Batang', 'Batang'], ['함초롬돋움', 'HCR Dotum', 'Malgun Gothic', 'Dotum'],
  ['한컴바탕', 'Haansoft Batang', 'HCR Batang', 'Batang'], ['한컴돋움', 'Haansoft Dotum', 'HCR Dotum', 'Malgun Gothic', 'Dotum'],
  ['나눔고딕', 'NanumGothic', 'Nanum Gothic'], ['나눔명조', 'NanumMyeongjo', 'Nanum Myeongjo'],
  ['나눔바른고딕', 'NanumBarunGothic'], ['나눔스퀘어', 'NanumSquare'],
  ['본고딕', 'Noto Sans KR', 'Source Han Sans KR'], ['본명조', 'Noto Serif KR', 'Source Han Serif KR'],
];

/** 설정 묶음. officekit = 실제 앱, onlyoffice = AllFontsGen 9.4와 같은 결과 (기준 시험용) */
export const PROFILES = {
  officekit: { priority: OFFICEKIT_PRIORITY, allNames: true, aliases: KOREAN_ALIASES, selectionPath: 'id' },
  onlyoffice: { priority: ONLYOFFICE_PRIORITY, allNames: false, aliases: [], selectionPath: 'path' },
};

const FONT_EXT = /\.(ttf|ttc|otf|otc)$/i;
const slash = (p) => p.replace(/\\/g, '/');

/** C++ std::wstring '<' (UTF-16 코드 단위 비교) */
const wcmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
/** g_sort_font_names: ASCII 대문자만 소문자로 바꿔 비교 */
const lowerAscii = (s) => s.replace(/[A-Z]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 32));
export const compareFontNames = (a, b) => wcmp(lowerAscii(a), lowerAscii(b));

/** IsEastAsianScript (ApplicationFontsWorker.cpp) */
export function isEastAsianCode(v) {
  return (v >= 0x3100 && v <= 0x312F) || (v >= 0x31A0 && v <= 0x31BF) || (v >= 0x4E00 && v <= 0x9FEA) || (v >= 0x3400 && v <= 0x4DB5)
    || (v >= 0x20000 && v <= 0x2A6D6) || (v >= 0x2A700 && v <= 0x2B734) || (v >= 0x2B740 && v <= 0x2B81D) || (v >= 0x2B820 && v <= 0x2CEA1)
    || (v >= 0x2CEB0 && v <= 0x2EBE0) || (v >= 0xF900 && v <= 0xFAFF) || (v >= 0x2F800 && v <= 0x2FA1F) || (v >= 0x2F00 && v <= 0x2FDF)
    || (v >= 0x2E80 && v <= 0x2EFF) || (v >= 0x31C0 && v <= 0x31EF) || (v >= 0x2FF0 && v <= 0x2FFF) || (v >= 0x1100 && v <= 0x11FF)
    || (v >= 0xA960 && v <= 0xA97F) || (v >= 0xD7B0 && v <= 0xD7FF) || (v >= 0x3130 && v <= 0x318F) || (v >= 0xFF00 && v <= 0xFFEF)
    || (v >= 0xAC00 && v <= 0xD7AF) || (v >= 0x3040 && v <= 0x309F) || (v >= 0x1B100 && v <= 0x1B12F) || (v >= 0x1B000 && v <= 0x1B0FF)
    || (v >= 0x3190 && v <= 0x319F) || (v >= 0x30A0 && v <= 0x30FF) || (v >= 0x31F0 && v <= 0x31FF) || (v >= 0xA4D0 && v <= 0xA4FF)
    || (v >= 0x16F00 && v <= 0x16F9F) || (v >= 0x1B170 && v <= 0x1B2FF) || (v >= 0x17000 && v <= 0x187EC) || (v >= 0x18800 && v <= 0x18AFF)
    || (v >= 0xA000 && v <= 0xA48F) || (v >= 0xA490 && v <= 0xA4CF);
}
export const isEastAsianName = (s) => { for (const ch of s) if (isEastAsianCode(ch.codePointAt(0))) return true; return false; };
const HANGUL = /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/;

// CFontInfo::IsStyleRegular 등 (스타일 이름이 '표준' 이름인지 — 같은 이름·굵기·기울임이 겹칠 때 바꿔 넣을지)
const STYLE_R = new Set(['regular', 'normal', 'roman', 'standard']);
const STYLE_I = new Set(['italic', 'oblique']);
const STYLE_B = new Set(['bold']);
const STYLE_BI = new Set(['bold italic', 'bolditalic', 'bold oblique', 'boldoblique', 'italic bold', 'oblique bold']);
const isStd = (style, b, i) => {
  const s = (style || '').toLowerCase();
  if (b && i) return STYLE_BI.has(s);
  if (b) return STYLE_B.has(s);
  if (i) return STYLE_I.has(s);
  return STYLE_R.has(s);
};

/**
 * 폴더들에서 글꼴 파일 목록 → [{ path, size, mtimeMs }] (경로를 '/'로 바꾼 문자열 순서 = AllFontsGen std::sort)
 * recursive: 하위 폴더까지 (기본 true — AllFontsGen과 같게)
 */
export function listFontFiles(dirs, { recursive = true } = {}) {
  const out = [];
  const seen = new Set();
  const walk = (dir, depth) => {
    let ents = [];
    try { ents = readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of ents) {
      const p = join(dir, e.name);
      if (e.isDirectory()) { if (recursive && depth < 8 && e.name !== '.git') walk(p, depth + 1); continue; }
      if (!FONT_EXT.test(e.name)) continue;
      let st; try { st = statSync(p); } catch { continue; }
      if (!st.isFile() || st.size < 12 || st.size > MAX_FONT_FILE_SIZE) continue;
      const key = slash(resolve(p)).toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ path: resolve(p), size: st.size, mtimeMs: Math.floor(st.mtimeMs) });
    }
  };
  for (const d of dirs) if (d) walk(resolve(d), 0);
  out.sort((a, b) => wcmp(slash(a.path), slash(b.path)));
  return out;
}

/** (SHORT)(v * (1000 / upem)) — C 형 변환처럼 0 쪽으로 자르고 16비트로 감는다 */
const scaleShort = (v, upem) => { const x = upem ? Math.trunc(v * (1000 / upem)) : v; return ((x & 0xffff) << 16) >> 16; };

/**
 * 글꼴 파일 하나 → CFontInfo 같은 레코드들 (CFontList::Add(FT_Library…) 규칙). 비트맵·Type1 등은 []
 * file: { path, size } (또는 시험용 { path, size, source }), opts.allNames: 다른 언어 이름을 모두(Office Kit) / 하나만(9.4)
 */
export function readFontFile(file, opts = {}) {
  const src = file.source ?? openFileSource(file.path);
  const recs = [];
  try {
    const offs = faceOffsets(src);
    offs.forEach((off, faceIndex) => {
      let f;
      try { f = readFace(src, off); } catch { return; }
      // 크기를 바꿀 수 없는(비트맵) 글꼴은 뺀다. TrueType(glyf) = 1, OpenType(CFF) = 2
      const format = f.hasGlyf ? 1 : f.isCFF ? 2 : 0;
      if (!format) return;
      const os2 = f.os2;
      let bold; let italic;
      if (os2) { italic = (os2.fsSelection & 512) || (os2.fsSelection & 1) ? 1 : 0; bold = os2.fsSelection & 32 ? 1 : 0; }
      else { bold = f.macStyle & 1 ? 1 : 0; italic = f.macStyle & 2 ? 1 : 0; }
      let family = freetypeName(f.names, 1) || '';
      const style = freetypeName(f.names, 2) || '';
      let cp1 = os2 ? os2.ulCodePageRange[0] : 0; const cp2 = os2 ? os2.ulCodePageRange[1] : 0;
      for (const cm of f.charmaps) {
        if (cm.platformID !== 3) continue;
        const bit = { 0: 0x80000000, 2: 0x00020000, 3: 0x00040000, 4: 0x00100000, 5: 0x00080000, 6: 0x00200000 }[cm.encodingID];
        if (bit) cp1 = (cp1 | bit) >>> 0;
      }
      // 다른 언어 가족 이름 (이름 ID 1, 유니코드·Windows 표만). 영어 이름이 깨졌으면('?') 첫 다른 이름을 가족 이름으로
      let badAscii = family.includes('?');
      const localized = [];
      for (const n of f.names) {
        if (n.nameID !== 1 || n.raw.length === 0) continue;
        if (!(n.platformID === 0 || (n.platformID === 3 && (n.encodingID === 0 || n.encodingID === 1 || n.encodingID === 10)))) continue;
        const v = n.value;
        if (v.includes(family) || family.includes(v)) continue;
        if (badAscii && localized.length === 0) { family = v; badAscii = false; continue; }
        if (!localized.some((x) => x.v === v)) localized.push({ v, ko: (n.languageID & 0x3ff) === 0x012 || HANGUL.test(v) });
      }
      if (!family) return;
      let names;
      if (opts.allNames) {
        // 한국어 이름을 앞으로 (편집기의 _ea 미리보기·인터페이스 이름이 names[0]을 쓴다)
        names = [...localized.filter((x) => x.ko), ...localized.filter((x) => !x.ko)].map((x) => x.v);
      } else {
        names = localized.length ? [localized[0].v] : []; // 9.4.0.129는 다른 이름을 하나만 쓴다
      }
      const upem = f.unitsPerEm;
      const rec = {
        file, faceIndex, name: family, style, names, bold, italic,
        fixed: f.post && f.post.isFixedPitch ? 1 : 0,
        panose: os2 ? os2.panose : new Array(10).fill(0),
        ur: os2 ? os2.ulUnicodeRange.map((x) => x >>> 0) : [0, 0, 0, 0],
        cp: [cp1 >>> 0, cp2 >>> 0],
        weight: os2 ? os2.usWeightClass : 0, width: os2 ? os2.usWidthClass : 0, familyClass: os2 ? os2.sFamilyClass : 0, format,
        avg: os2 ? scaleShort(os2.xAvgCharWidth, upem) : 0, asc: os2 ? scaleShort(os2.sTypoAscender, upem) : 0,
        desc: os2 ? scaleShort(os2.sTypoDescender, upem) : 0, gap: os2 ? scaleShort(os2.sTypoLineGap, upem) : 0,
        xh: os2 ? scaleShort(os2.sxHeight, upem) : 0, caph: os2 ? scaleShort(os2.sCapHeight, upem) : 0,
        fsType: os2 ? os2.fsType : 0,
        symbolic: f.symbolic,
        // FreeType face->family_name (ASCW3 판정용, 원래 ASCII 이름)
        ftFamily: freetypeName(f.names, 1) || '',
        // 별칭을 덧붙이기 전의 이름 (미리보기 _ea 이름은 글꼴 자신의 이름만)
        ownNames: names,
        // 미리보기 그리기(resvg)가 이 얼굴을 찾을 이름들: 표기 가족(16) → 가족(1) → 다른 언어 이름
        cssFamilies: [...new Set([
          ...f.names.filter((n) => n.nameID === 16 && n.value && (n.platformID !== 3 || (n.languageID & 0x3ff) === 9)).map((n) => n.value),
          ...f.names.filter((n) => n.nameID === 1 && n.value && (n.platformID !== 3 || (n.languageID & 0x3ff) === 9)).map((n) => n.value),
          family, ...names,
        ])].filter((s) => s && !/[\u0000-\u001f]/.test(s)),
      };
      recs.push(rec);
    });
  } finally {
    if (!file.source) src.close?.();
  }
  return recs;
}

/** 얼굴 하나의 모든 cmap 문자 (AllFontsGen CApplicationFontsSymbols::CheckSymbols: 모든 cmap 표를 차례로) */
function faceCodes(file, faceIndex) {
  const src = file.source ?? openFileSource(file.path);
  try {
    const off = faceOffsets(src)[faceIndex];
    if (off == null) return { charmaps: [] };
    const f = readFace(src, off, { codes: 'all' });
    return { charmaps: f.charmaps, ftFamily: freetypeName(f.names, 1) || '' };
  } catch {
    return { charmaps: [] };
  } finally {
    if (!file.source) src.close?.();
  }
}

/**
 * 대체 범위 계산 (SaveAllFontsJS의 CSymbolSimpleChecker2·Apply2를 그대로 — 임시 버퍼를 덜 지우는 원본 동작까지 같게)
 * families: arrFonts 순서(이름순)의 [{ name, slots: { R, B, I, BI } }], slot = { file, faceIndex } | null
 * priority: 표준 우선순위 이름 목록 → [시작, 끝, infos 번호, …]
 */
export function computeRanges(families, priority, sizeOf) {
  const MAX = 0x10FFFF;
  const nCount = families.length;
  const std = new Map();
  priority.forEach((n, i) => { if (!std.has(n)) std.set(n, i + 1); });
  const pr = families.map((fam) => {
    if (std.has(fam.name)) return { name: fam.name, priority: std.get(fam.name) };
    let size = 0;
    for (const k of ['R', 'B', 'I', 'BI']) if (fam.slots[k]) size = Math.max(size, sizeOf(fam.slots[k].file));
    return { name: fam.name, priority: size };
  });
  const sorted = [...pr].sort((a, b) => (a.priority !== b.priority ? a.priority - b.priority : wcmp(a.name, b.name)));
  const rank = new Map(sorted.map((p, i) => [p.name, i + 1]));
  const sumPriority = sorted.length + 1;
  const index = new Map(families.map((f, i) => [f.name, i]));

  const symbols = new Int32Array(MAX + 1);
  const tmp = new Uint8Array(MAX + 1);
  let min = MAX + 1; let max = -1;
  const check = (code, style) => {
    if (code > MAX) return;
    if (code > max) max = code;
    if (code < min) min = code;
    tmp[code] |= style;
  };
  for (let fi = 0; fi < nCount; fi++) {
    const fam = families[fi];
    const nPriority = rank.get(fam.name);
    let mask = 0;
    for (const [k, style] of [['R', 1], ['B', 2], ['I', 4], ['BI', 8]]) {
      const slot = fam.slots[k];
      if (!slot) continue;
      mask |= style;
      const { charmaps, ftFamily } = faceCodes(slot.file, slot.faceIndex);
      const isASC = ftFamily === 'ASCW3';
      for (const cm of charmaps) for (const c of cm.codes || []) if (!isASC || c < 35 || c > 255) check(c, style);
    }
    // Apply2(mask, sumPriority)
    let t = min; const last = max + 1;
    while (t < last) {
      if (tmp[t] !== mask) { t++; continue; }
      const first = t; let count = 1; t++;
      while (t < last && tmp[t] === mask) { t++; count++; }
      let p;
      if (nPriority !== 1) {
        if (first > 0x4DFF && first < 0x9FFF) p = count > 30 ? nPriority : nPriority + sumPriority;
        else p = count > 10 ? nPriority : nPriority + sumPriority;
      } else p = nPriority;
      for (let i = first; i < first + count; i++) if (symbols[i] === 0 || symbols[i] > p) symbols[i] = p;
    }
    // 원본은 m_pTmpSymbols 맨 앞부터 (max - min)바이트만 지운다 — 그대로 따라 한다
    if (max >= min) tmp.fill(0, 0, max - min);
    min = MAX + 1; max = -1;
  }
  const ranges = [];
  let cur = 0; let start = 0;
  for (let i = 0; i < MAX + 1; i++) {
    const v = symbols[i];
    const test = v > sumPriority ? v - sumPriority : v;
    if (cur === test) continue;
    if (cur !== 0) ranges.push(start, i - 1, index.get(sorted[cur - 1].name));
    cur = test; start = i;
  }
  if (cur !== 0) ranges.push(start, MAX - 1, index.get(sorted[cur - 1].name));
  return ranges;
}

/** 선택 이진 레코드 하나 (ToBuffer, 버전 2) */
function selectionRecord(rec, path) {
  const str = (s) => { const b = Buffer.from(s, 'utf8'); const l = Buffer.alloc(4); l.writeInt32LE(b.length); return [l, b]; };
  const parts = [Buffer.alloc(4)];
  parts.push(...str(rec.name));
  const nn = Buffer.alloc(4); nn.writeInt32LE(rec.names.length); parts.push(nn);
  for (const n of rec.names) parts.push(...str(n));
  parts.push(...str(path));
  const tail = Buffer.alloc(4 * 4 + 4 + 10 + 6 * 4 + 10 * 2 + 2);
  let o = 0;
  const i32 = (v) => { tail.writeInt32LE(v | 0, o); o += 4; };
  const u32 = (v) => { tail.writeUInt32LE(v >>> 0, o); o += 4; };
  const s16 = (v) => { tail.writeUInt16LE(v & 0xffff, o); o += 2; };
  i32(rec.faceIndex); i32(rec.italic); i32(rec.bold); i32(rec.fixed);
  i32(10); for (let k = 0; k < 10; k++) tail[o++] = (rec.panose[k] || 0) & 0xff;
  rec.ur.forEach(u32); rec.cp.forEach(u32);
  s16(rec.weight); s16(rec.width); s16(rec.familyClass); s16(rec.format);
  [rec.avg, rec.asc, rec.desc, rec.gap, rec.xh, rec.caph].forEach(s16);
  s16(rec.fsType);
  const buf = Buffer.concat([...parts, tail]);
  buf.writeInt32LE(buf.length, 0);
  return buf;
}

const escName = (s) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\n\r]/g, '');

/**
 * AllFonts.js 만들기
 * files: listFontFiles() 결과 (또는 { path, size, source? } 배열 — source는 bufferSource 등, 시험용)
 * opts: { profile: 'officekit'|'onlyoffice', priority, allNames, aliases,
 *         selectionPath: 'id' | 'path' | (file) => string,   // 선택 이진에 쓰는 경로
 *         filesMode: 'web' | 'path' }                        // __fonts_files: 웹 번호("000") 또는 실제 경로
 * → { js, files:[{ id, path, size, mtimeMs }](웹 번호 순), infos, ranges, selection:[레코드], thumbs:[미리보기 정보], stats }
 */
export function generateAllFonts(files, opts = {}) {
  const prof = PROFILES[opts.profile ?? 'officekit'] ?? PROFILES.officekit;
  const priority = opts.priority ?? prof.priority;
  const allNames = opts.allNames ?? prof.allNames;
  const aliases = opts.aliases ?? prof.aliases;
  const selectionPath = opts.selectionPath ?? prof.selectionPath;
  const filesMode = opts.filesMode ?? 'web';
  const t0 = Date.now();

  // 1. 글꼴 읽기 + CFontList::Add (같은 이름·굵기·기울임이면, 먼저 것의 스타일 이름이 표준이 아닐 때만 바꿔 넣는다)
  const list = [];
  let skippedFiles = 0;
  for (const file of files) {
    let recs = [];
    try { recs = readFontFile(file, { allNames }); } catch { recs = []; }
    if (!recs.length) { skippedFiles++; continue; }
    for (const r of recs) {
      const i = list.findIndex((x) => x.name === r.name && x.bold === r.bold && x.italic === r.italic);
      if (i < 0) { list.push(r); continue; }
      if (!isStd(list[i].style, list[i].bold, list[i].italic)) list[i] = r;
    }
  }

  // 2. 한국어 별칭: 그 이름을 가진 글꼴이 없으면 후보 가족의 names에 덧붙인다
  const aliasAdded = [];
  if (aliases.length) {
    const known = new Set();
    for (const r of list) { known.add(r.name); r.names.forEach((n) => known.add(n)); }
    const byName = new Map();
    for (const r of list) { if (!byName.has(r.name)) byName.set(r.name, []); byName.get(r.name).push(r); }
    for (const [alias, ...targets] of aliases) {
      if (known.has(alias)) continue;
      const target = targets.find((t) => byName.has(t));
      if (!target) continue;
      for (const r of byName.get(target)) if (!r.names.includes(alias)) r.names = [...r.names, alias];
      known.add(alias);
      aliasAdded.push([alias, target]);
    }
  }

  // 3. 파일 목록 (처음 나온 순서) + 웹 번호 (경로 순서로 000, 001 …)
  const fileIndex = new Map(); const fileList = [];
  for (const r of list) if (!fileIndex.has(r.file)) { fileIndex.set(r.file, fileList.length); fileList.push(r.file); }
  const digits = fileList.length >= 1000 ? 4 : 3;
  const byPath = [...fileList].sort((a, b) => wcmp(slash(a.path), slash(b.path)));
  const idOf = new Map(byPath.map((f, i) => [f, String(i).padStart(digits, '0')]));

  // 4. 가족별 R/I/B/BI (CFontInfoJS) + 이름순 정렬 (g_sort_font_names)
  const fams = new Map(); const famOrder = [];
  for (const r of list) {
    const k = r.bold && r.italic ? 'BI' : r.bold ? 'B' : r.italic ? 'I' : 'R';
    if (!fams.has(r.name)) { fams.set(r.name, { name: r.name, slots: { R: null, I: null, B: null, BI: null }, styles: {} }); famOrder.push(r.name); }
    const fam = fams.get(r.name);
    if (!fam.slots[k] || !isStd(fam.styles[k], k === 'B' || k === 'BI', k === 'I' || k === 'BI')) { fam.slots[k] = r; fam.styles[k] = r.style; }
  }
  const names = [...famOrder].sort(compareFontNames);
  const families = names.map((n) => fams.get(n));
  const slotOut = (s) => (s ? [fileIndex.get(s.file), s.faceIndex] : [-1, -1]);
  const infos = families.map((f) => [f.name, ...slotOut(f.slots.R), ...slotOut(f.slots.I), ...slotOut(f.slots.B), ...slotOut(f.slots.BI)]);

  // 5. 대체 범위
  const t1 = Date.now();
  const ranges = computeRanges(families, priority, (file) => file.size ?? 0);
  const t2 = Date.now();

  // 6. 선택 이진 (목록 순서)
  const pathFor = (file) => (typeof selectionPath === 'function' ? selectionPath(file)
    : selectionPath === 'id' ? idOf.get(file) : slash(file.path));
  const head = Buffer.alloc(4); head.writeInt32LE(list.length);
  const selection = Buffer.concat([head, ...list.map((r) => selectionRecord(r, pathFor(r.file)))]);

  // 7. AllFonts.js 글 (AllFontsGen과 같은 모양, UTF-8 BOM)
  const fontsFiles = fileList.map((f) => (filesMode === 'web' ? idOf.get(f) : slash(f.path)));
  let js = `\uFEFFwindow["__all_fonts_js_version__"] = ${ALL_FONTS_JS_VERSION};\n\n`;
  js += fontsFiles.length ? `window["__fonts_files"] = [\n${fontsFiles.map((f) => `"${f}"`).join(',\n')}\n];\n\n` : 'window["__fonts_files"] = []; \n\n';
  js += `window["__fonts_infos"] = [\n${infos.map((x) => `["${escName(x[0])}",${x.slice(1).join(',')}]`).join(',\n')}${infos.length ? '\n' : ''}];\n\n`;
  js += `window["__fonts_ranges"] = [${ranges.length ? `\n${ranges.join(',')}` : ''}];\n\n`;
  js += `window["g_fonts_selection_bin"] = "${selection.toString('base64')}";\n`;

  // 8. 미리보기 정보 (SaveThumbnails: R → I → B → BI 얼굴, _ea는 names[0]이 동아시아 이름일 때)
  const altEA = new Map();
  for (const r of list) {
    const own = r.ownNames ?? r.names;
    if (own.length && !isEastAsianName(r.name) && isEastAsianName(own[0]) && !altEA.has(r.name)) altEA.set(r.name, own[0]);
  }
  const thumbs = families.map((f) => {
    const s = f.slots.R || f.slots.I || f.slots.B || f.slots.BI;
    return {
      name: f.name, eaName: altEA.get(f.name) ?? null, path: s.file.path, faceIndex: s.faceIndex, symbolic: !!s.symbolic,
      css: s.cssFamilies ?? [f.name], weight: s.weight || 400, italic: !!s.italic,
    };
  });

  const hangul = (() => { for (let i = 0; i < ranges.length; i += 3) if (ranges[i] <= 0xAC00 && ranges[i + 1] >= 0xAC00) return infos[ranges[i + 2]][0]; return null; })();
  return {
    js,
    files: byPath.map((f) => ({ id: idOf.get(f), path: f.path, size: f.size ?? 0, mtimeMs: f.mtimeMs ?? 0 })),
    fontsFiles, infos, ranges, selection: list, thumbs,
    stats: {
      inputFiles: files.length, skippedFiles, files: fileList.length, faces: list.length, families: infos.length, ranges: ranges.length / 3,
      hangulFallback: hangul, aliases: aliasAdded, readMs: t1 - t0, rangesMs: t2 - t1, totalMs: Date.now() - t0,
    },
  };
}
