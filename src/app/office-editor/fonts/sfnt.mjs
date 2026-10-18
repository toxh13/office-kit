// SPDX-License-Identifier: AGPL-3.0-only
// sfnt.mjs — 글꼴 파일(TrueType·OpenType·TTC)에서 AllFonts.js에 필요한 표만 읽는다. 의존성 없음.
//   ONLYOFFICE AllFontsGen이 FreeType으로 얻는 값(이름·OS/2·cmap 순회)을 같은 규칙으로 흉내 낸다.
//   Office Kit 글꼴 생성기의 일부 (Copyright (C) 2026 toxh13). 원본 규칙: ONLYOFFICE core (Ascensio System SIA), FreeType.
import { closeSync, fstatSync, openSync, readSync } from 'node:fs';

/** Buffer를 읽기 원천으로 */
export function bufferSource(buf) {
  return { size: buf.length, read: (off, len) => buf.subarray(off, Math.min(buf.length, off + len)) };
}

/** 파일을 읽기 원천으로 (필요한 표만 읽는다). 다 쓰면 close() */
export function openFileSource(path) {
  const fd = openSync(path, 'r');
  const size = fstatSync(fd).size;
  return {
    size,
    read(off, len) {
      len = Math.max(0, Math.min(len, size - off));
      const b = Buffer.allocUnsafe(len);
      let got = 0;
      while (got < len) { const n = readSync(fd, b, got, len - got, off + got); if (!n) break; got += n; }
      return got === len ? b : b.subarray(0, got);
    },
    close() { try { closeSync(fd); } catch {} },
  };
}

/** 'ttc' | 'sfnt' | null */
export function fontKind(src) {
  if (src.size < 12) return null;
  const h = src.read(0, 4);
  const tag = h.toString('latin1');
  if (tag === 'ttcf') return 'ttc';
  if (tag === 'OTTO' || tag === 'true' || h.readUInt32BE(0) === 0x00010000) return 'sfnt';
  return null;
}

/** 얼굴(face) 시작 위치들. 모르는 형식이면 [] */
export function faceOffsets(src) {
  if (Buffer.isBuffer(src)) src = bufferSource(src);
  const kind = fontKind(src);
  if (kind === 'sfnt') return [0];
  if (kind !== 'ttc') return [];
  const n = src.read(8, 4).readUInt32BE(0);
  if (n < 1 || n > 4096 || 12 + 4 * n > src.size) return [];
  const b = src.read(12, 4 * n);
  return Array.from({ length: n }, (_, i) => b.readUInt32BE(4 * i));
}

// Mac Roman 128~255 (이름 표 플랫폼 1)
const MAC_ROMAN = 'ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ';
/** UTF-16BE 바이트 → 문자열 (홀수 끝 바이트는 버림) */
export const utf16be = (raw) => { const c = Buffer.from(raw.subarray(0, raw.length & ~1)); return c.swap16().toString('utf16le'); };

function readNames(t) {
  const out = [];
  if (!t || t.length < 6) return out;
  const count = t.readUInt16BE(2);
  const so = t.readUInt16BE(4);
  for (let i = 0; i < count; i++) {
    const r = 6 + 12 * i;
    if (r + 12 > t.length) break;
    const platformID = t.readUInt16BE(r); const encodingID = t.readUInt16BE(r + 2);
    const languageID = t.readUInt16BE(r + 4); const nameID = t.readUInt16BE(r + 6);
    const len = t.readUInt16BE(r + 8); const off = t.readUInt16BE(r + 10);
    if (so + off + len > t.length) continue;
    const raw = t.subarray(so + off, so + off + len);
    let value = null;
    if (platformID === 0 || (platformID === 3 && (encodingID === 0 || encodingID === 1 || encodingID === 10))) value = utf16be(raw);
    else if (platformID === 1 && encodingID === 0) value = [...raw].map((c) => (c < 128 ? String.fromCharCode(c) : MAC_ROMAN[c - 128])).join('');
    out.push({ platformID, encodingID, languageID, nameID, raw, value });
  }
  return out;
}

/**
 * FreeType tt_face_get_name와 같은 규칙으로 이름 하나 고르기 → ASCII 문자열 (32~127 밖은 '?') | null
 * Windows 영어 이름을 먼저, 없으면 Mac 영어·로마자, 그다음 유니코드 플랫폼
 */
export function freetypeName(names, nameID) {
  let foundWin = -1; let isEnglish = false; let foundAppleEnglish = -1; let foundAppleRoman = -1; let foundUnicode = -1;
  names.forEach((r, n) => {
    if (r.nameID !== nameID || r.raw.length === 0) return;
    if (r.platformID === 0 || r.platformID === 2) foundUnicode = n;
    else if (r.platformID === 1) {
      if (r.languageID === 0) foundAppleEnglish = n;
      else if (r.encodingID === 0) foundAppleRoman = n;
    } else if (r.platformID === 3) {
      if (foundWin === -1 || (r.languageID & 0x3ff) === 0x009) {
        if (r.encodingID === 0 || r.encodingID === 1 || r.encodingID === 10) { isEnglish = (r.languageID & 0x3ff) === 0x009; foundWin = n; }
      }
    }
  });
  const foundApple = foundAppleEnglish >= 0 ? foundAppleEnglish : foundAppleRoman;
  const ascii16 = (raw) => { let s = ''; for (let i = 0; i + 1 < raw.length; i += 2) { const c = raw.readUInt16BE(i); if (c === 0) break; s += c < 32 || c > 127 ? '?' : String.fromCharCode(c); } return s; };
  const ascii8 = (raw) => { let s = ''; for (const c of raw) { if (c === 0) break; s += c < 32 || c > 127 ? '?' : String.fromCharCode(c); } return s; };
  if (foundWin >= 0 && !(foundApple >= 0 && !isEnglish)) return ascii16(names[foundWin].raw);
  if (foundApple >= 0) return ascii8(names[foundApple].raw);
  if (foundUnicode >= 0) return ascii16(names[foundUnicode].raw);
  return null;
}

// ── cmap ──
// FreeType가 싣는 부속 표 형식 (14는 변형 선택자: 문자 순회에 안 씀)
const CMAP_FORMATS = new Set([0, 2, 4, 6, 8, 10, 12, 13, 14]);

/**
 * FreeType ttcmap.c의 기본 검사(FT_VALIDATE_DEFAULT) — 통과 못 한 부속 표는 FreeType가 버린다 ('broken cmap sub-table ignored').
 * 예: 길이 칸(16비트)이 넘친 Mac 한국어 형식 6 표
 */
function cmapValid(t, at, format) {
  const L = t.length;
  const u16 = (o) => t.readUInt16BE(o); const u32 = (o) => t.readUInt32BE(o);
  try {
    if (format === 0) {
      if (at + 4 > L) return false;
      const len = u16(at + 2);
      return !(at + len > L || len < 262);
    }
    if (format === 2) {
      if (at + 4 > L) return false;
      const len = u16(at + 2);
      if (at + len > L || len < 6 + 512) return false;
      let maxSubs = 0;
      for (let n = 0; n < 256; n++) maxSubs = Math.max(maxSubs, u16(at + 6 + 2 * n) >> 3);
      const subs = at + 518; const glyphIds = subs + (maxSubs + 1) * 8;
      if (glyphIds > L) return false;
      for (let n = 0; n <= maxSubs; n++) {
        const p = subs + 8 * n;
        const count = u16(p + 2); const offset = u16(p + 6);
        if (count === 0) continue;
        if (offset !== 0) { const ids = p + 6 + offset; if (ids < glyphIds || ids + count * 2 > at + len) return false; }
      }
      return true;
    }
    if (format === 4) {
      if (at + 4 > L) return false;
      let len = u16(at + 2);
      if (at + len > L) len = L - at;
      if (len < L - at) len = L - at;
      if (len < 16) return false;
      const segs = u16(at + 6) >> 1;
      if (len < 16 + segs * 2 * 4) return false;
      const ends = at + 14; const starts = at + 16 + segs * 2; const deltas = starts + segs * 2; const offs = deltas + segs * 2; const glyphIds = offs + segs * 2;
      for (let n = 0; n < segs; n++) {
        const start = u16(starts + 2 * n); const end = u16(ends + 2 * n); const offset = u16(offs + 2 * n);
        if (start > end) return false;
        if (offset && offset !== 0xffff) {
          const p = offs + 2 * n + offset;
          if (!(n === segs - 1 && start === 0xffff && end === 0xffff)) { if (p < glyphIds || p + (end - start + 1) * 2 > L) return false; }
        } else if (offset === 0xffff) {
          if (n !== segs - 1 || !(start === 0xffff && end === 0xffff)) return false;
        }
      }
      return true;
    }
    if (format === 6) {
      if (at + 10 > L) return false;
      const len = u16(at + 2); const count = u16(at + 8);
      return !(at + len > L || len < 10 + count * 2);
    }
    if (format === 10) {
      if (at + 20 > L) return false;
      const len = u32(at + 4); const count = u32(at + 16);
      return !(len > L - at || len < 20 || Math.floor((len - 20) / 2) < count);
    }
    if (format === 12 || format === 13) {
      if (at + 16 > L) return false;
      const len = u32(at + 4); const groups = u32(at + 12);
      if (len > L - at || len < 16 || Math.floor((len - 16) / 12) < groups) return false;
      let last = 0;
      for (let n = 0; n < groups; n++) {
        const p = at + 16 + 12 * n;
        const start = u32(p); const end = u32(p + 4); const sid = u32(p + 8);
        if (start > end) return false;
        if (n > 0 && start <= last) return false;
        if (format === 12 && sid + (end - start) > 0xffffffff) return false;
        last = end;
      }
      return true;
    }
    if (format === 8 || format === 14) return at + 10 <= L;
  } catch {
    return false;
  }
  return false;
}

/** cmap 부속 표 머리 목록 (FreeType가 싣는 것만) → [{ platformID, encodingID, format, at }] (at = cmap 표 안 위치) */
function cmapHeaders(t) {
  if (!t || t.length < 4) return [];
  const n = t.readUInt16BE(2);
  const out = [];
  for (let i = 0; i < n; i++) {
    const r = 4 + 8 * i;
    if (r + 8 > t.length) break;
    const at = t.readUInt32BE(r + 4);
    if (!at || at > t.length - 2) continue;
    const format = t.readUInt16BE(at);
    if (!CMAP_FORMATS.has(format) || !cmapValid(t, at, format)) continue;
    out.push({ platformID: t.readUInt16BE(r), encodingID: t.readUInt16BE(r + 2), format, at });
  }
  return out;
}

/**
 * 부속 표 하나의 문자 코드들 (글리프 번호가 0이 아니고 numGlyphs보다 작은 것만 — FreeType FT_Get_Next_Char와 같게)
 * → number[] (올림차순이 아닐 수 있음) | null(읽지 못함)
 */
function cmapCodes(t, h, numGlyphs) {
  const at = h.at; const fmt = h.format; const L = t.length;
  const out = [];
  const ok = (g) => g !== 0 && g < numGlyphs;
  const u16 = (o) => (o + 2 <= L ? t.readUInt16BE(o) : 0);
  if (fmt === 0) {
    for (let c = 0; c < 256; c++) if (at + 6 + c < L && ok(t[at + 6 + c])) out.push(c);
  } else if (fmt === 2) {
    // 상위 바이트 표 (CJK 옛 인코딩): FreeType tt_cmap2_char_index·char_next를 그대로 따라 한다
    const s16 = (o) => (o + 2 <= L ? t.readInt16BE(o) : 0);
    const keys = at + 6; const subs = at + 518;
    const subOf = (cc) => {
      const lo = cc & 0xff; const hi = cc >> 8;
      if (hi === 0) return u16(keys + 2 * lo) === 0 ? subs : null;
      const s = subs + (u16(keys + 2 * hi) & ~7);
      return s === subs ? null : s;
    };
    const index0 = () => { // tt_cmap2_char_index(0)
      const sub = subOf(0); if (sub === null) return 0;
      const start = u16(sub); const count = u16(sub + 2); const delta = s16(sub + 4); const offset = u16(sub + 6);
      const idx = 0 - start;
      if (idx >= 0 && idx < count && offset !== 0) { const g = u16(sub + 6 + offset + 2 * idx); return g ? (g + delta) & 0xffff : 0; }
      return 0;
    };
    const next = (from) => { // tt_cmap2_char_next → [code, gid] | null
      let cc = from + 1;
      while (cc < 0x10000) {
        const sub = subOf(cc);
        if (sub !== null) {
          const start = u16(sub); const count = u16(sub + 2); const delta = s16(sub + 4); const offset = u16(sub + 6);
          let lo = cc & 0xff;
          if (lo >= start + count && cc <= 0xff) { cc = 0x100; continue; }
          if (offset === 0) {
            if (cc === 0x100) return null;
          } else {
            let pos;
            if (lo < start) { lo = start; pos = 0; } else pos = lo - start;
            let p = sub + 6 + offset + 2 * pos;
            cc = (cc & ~0xff) + lo;
            for (; pos < count; pos++, cc++, p += 2) {
              const idx = u16(p);
              if (idx) { const g = (idx + delta) & 0xffff; if (g) return [cc, g]; }
            }
            if (count) cc--;
          }
        }
        if (cc <= 0xff) cc++; else cc = (cc & ~0xff) + 0x100;
      }
      return null;
    };
    // FT_Get_First_Char / FT_Get_Next_Char (numGlyphs 넘는 글리프는 건너뛴다)
    let code = 0; let g = index0();
    if (!ok(g)) { let r; do { r = next(code); if (r) [code, g] = r; } while (r && g >= numGlyphs); if (!r) g = 0; }
    let guard = 0;
    while (g && guard++ < 0x20000) {
      out.push(code);
      let r; do { r = next(code); if (r) [code, g] = r; } while (r && g >= numGlyphs);
      if (!r) break;
    }
  } else if (fmt === 4) {
    const segX2 = u16(at + 6);
    const ends = at + 14; const starts = ends + segX2 + 2; const deltas = starts + segX2; const ros = deltas + segX2;
    if (ros + segX2 > L) return null;
    for (let s = 0; s < segX2; s += 2) {
      const end = t.readUInt16BE(ends + s); const start = t.readUInt16BE(starts + s);
      const delta = t.readInt16BE(deltas + s); const ro = t.readUInt16BE(ros + s);
      if (start > end) continue;
      for (let c = start; c <= end && c < 0xffff; c++) {
        let g;
        if (ro === 0) g = (c + delta) & 0xffff;
        else if (ro === 0xffff) g = 0; // 깨진 마지막 구간
        else {
          const ga = ros + s + ro + 2 * (c - start);
          if (ga + 2 > L) break;
          g = t.readUInt16BE(ga);
          if (g !== 0) g = (g + delta) & 0xffff;
        }
        if (ok(g)) out.push(c);
      }
    }
  } else if (fmt === 6) {
    const first = u16(at + 6); const n = u16(at + 8);
    for (let i = 0; i < n; i++) if (ok(u16(at + 10 + 2 * i))) out.push(first + i);
  } else if (fmt === 10) {
    if (at + 20 > L) return null;
    const first = t.readUInt32BE(at + 12); const n = t.readUInt32BE(at + 16);
    for (let i = 0; i < n && at + 20 + 2 * i + 2 <= L; i++) if (ok(t.readUInt16BE(at + 20 + 2 * i))) out.push(first + i);
  } else if (fmt === 12 || fmt === 13) {
    if (at + 16 > L) return null;
    const n = t.readUInt32BE(at + 12);
    for (let i = 0; i < n; i++) {
      const r = at + 16 + 12 * i;
      if (r + 12 > L) break;
      const s = t.readUInt32BE(r); const e = Math.min(t.readUInt32BE(r + 4), 0x10ffff); const g0 = t.readUInt32BE(r + 8);
      if (e < s || e - s > 0x10ffff) continue;
      for (let c = s; c <= e; c++) if (ok(fmt === 12 ? g0 + (c - s) : g0)) out.push(c);
    }
  } else {
    return null; // 8, 14: 순회하지 않음
  }
  return out;
}

const isUnicodeCmap = (h) => h.platformID === 0 || (h.platformID === 3 && (h.encodingID === 1 || h.encodingID === 10));
const isUcs4Cmap = (h) => (h.platformID === 3 && h.encodingID === 10) || (h.platformID === 0 && (h.encodingID === 4 || h.encodingID === 6));

/** FreeType find_unicode_charmap: 뒤에서부터 UCS-4, 없으면 UCS-2 유니코드 표 → 머리 | null */
function pickUnicode(heads) {
  const cand = heads.filter((h) => h.format !== 14);
  for (let i = cand.length - 1; i >= 0; i--) if (isUnicodeCmap(cand[i]) && isUcs4Cmap(cand[i])) return cand[i];
  for (let i = cand.length - 1; i >= 0; i--) if (isUnicodeCmap(cand[i])) return cand[i];
  return null;
}

const i16 = (b, o) => (o + 2 <= b.length ? b.readInt16BE(o) : 0);
const u16 = (b, o) => (o + 2 <= b.length ? b.readUInt16BE(o) : 0);

/**
 * 얼굴 하나 읽기. src = Buffer 또는 openFileSource(). 깨진 글꼴이면 throw.
 * opts.codes = 'all'   → charmaps[i].codes (FreeType가 싣는 모든 부속 표의 문자 — AllFontsGen의 범위 계산용)
 *            = 'unicode' → unicode.codes (FreeType 기본 표)
 * → { tables, unitsPerEm, macStyle, numGlyphs, names, os2, post, isCFF, hasGlyf, charmaps, unicode }
 */
export function readFace(src, off = 0, opts = {}) {
  if (Buffer.isBuffer(src)) src = bufferSource(src);
  const dir = src.read(off, 12);
  if (dir.length < 12) throw new Error('bad offset');
  const n = dir.readUInt16BE(4);
  if (n === 0 || n > 512) throw new Error('bad table count');
  const recs = src.read(off + 12, 16 * n);
  if (recs.length < 16 * n) throw new Error('bad table directory');
  const dirTables = {};
  for (let i = 0; i < n; i++) {
    const r = 16 * i;
    const tag = recs.toString('latin1', r, r + 4);
    const o = recs.readUInt32BE(r + 8); const len = recs.readUInt32BE(r + 12);
    if (o + len <= src.size && !dirTables[tag]) dirTables[tag] = { off: o, len };
  }
  const get = (tag, max = Infinity) => (dirTables[tag] ? src.read(dirTables[tag].off, Math.min(dirTables[tag].len, max)) : null);
  const head = get('head', 54);
  if (!head || head.length < 46) throw new Error('no head');
  const maxp = get('maxp', 6);
  if (!maxp || maxp.length < 6) throw new Error('no maxp');
  const unitsPerEm = u16(head, 18);
  const macStyle = u16(head, 44);
  const numGlyphs = u16(maxp, 4);
  const names = readNames(get('name'));
  const os2t = get('OS/2', 100);
  // FreeType는 78바이트보다 짧은 OS/2를 버전 0으로 본다 (v0도 68바이트 이상이면 읽는다)
  const os2 = os2t && os2t.length >= 68 ? (() => {
    const version = u16(os2t, 0);
    const v1 = version >= 1 && os2t.length >= 86;
    const v2 = version >= 2 && os2t.length >= 96;
    return {
      version,
      xAvgCharWidth: i16(os2t, 2), usWeightClass: u16(os2t, 4), usWidthClass: u16(os2t, 6), fsType: u16(os2t, 8),
      sFamilyClass: i16(os2t, 30), panose: [...os2t.subarray(32, 42)],
      ulUnicodeRange: [0, 1, 2, 3].map((i) => os2t.readUInt32BE(42 + 4 * i)),
      fsSelection: u16(os2t, 62),
      sTypoAscender: i16(os2t, 68), sTypoDescender: i16(os2t, 70), sTypoLineGap: i16(os2t, 72),
      ulCodePageRange: v1 ? [os2t.readUInt32BE(78), os2t.readUInt32BE(82)] : [0, 0],
      sxHeight: v2 ? i16(os2t, 86) : 0, sCapHeight: v2 ? i16(os2t, 88) : 0,
    };
  })() : null;
  const postt = get('post', 16);
  const post = postt && postt.length >= 16 ? { isFixedPitch: postt.readUInt32BE(12) } : null;
  const cmapt = get('cmap');
  const heads = cmapHeaders(cmapt);
  const charmaps = heads.map((h) => ({ platformID: h.platformID, encodingID: h.encodingID, format: h.format }));
  const uni = pickUnicode(heads);
  const face = {
    tables: Object.keys(dirTables), unitsPerEm, macStyle, numGlyphs, names, os2, post,
    isCFF: !!(dirTables['CFF '] || dirTables.CFF2), hasGlyf: !!(dirTables.glyf && dirTables.loca),
    charmaps,
    unicode: uni ? { platformID: uni.platformID, encodingID: uni.encodingID, format: uni.format } : null,
    symbolic: !uni && heads.some((h) => h.platformID === 3 && h.encodingID === 0),
  };
  if (opts.codes === 'all') {
    heads.forEach((h, i) => { let c = null; try { c = cmapCodes(cmapt, h, numGlyphs); } catch { c = null; } charmaps[i].codes = c || []; });
  } else if (opts.codes === 'unicode' && uni) {
    try { face.unicode.codes = cmapCodes(cmapt, uni, numGlyphs) || []; } catch { face.unicode.codes = []; }
  }
  return face;
}
