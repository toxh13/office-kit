// Office 편집기 글꼴 시험 (src/app/office-editor/fonts) — AllFonts.js 만들기·미리보기 그림·캐시·글꼴 내주기
//   기준 시험: 묶음 core-fonts(편집기 묶음의 editor/fonts, XOR 풀어서)로 만든 AllFonts.js가 묶음에 든 AllFonts.js
//   (ONLYOFFICE allfontsgen 결과)와 바이트까지 같아야 한다. 묶음 캐시가 없으면 건너뛴다.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  ensureFonts, stopFontsWork, createObfuscatedStream, obfuscateFontBytes, fontsCacheKey, defaultFontDirs, FONT_XOR_KEY,
} from '../src/app/office-editor/fonts/index.mjs';
import {
  generateAllFonts, listFontFiles, computeRanges, compareFontNames, OFFICEKIT_PRIORITY, ONLYOFFICE_PRIORITY,
} from '../src/app/office-editor/fonts/allfonts.mjs';
import { bufferSource, faceOffsets, freetypeName, readFace } from '../src/app/office-editor/fonts/sfnt.mjs';
import {
  blankSprites, decodeSpriteBin, downsampleAlpha, encodeSpriteBin, renderThumbnails, thumbFileName, THUMB_FILES, thumbGeometry,
} from '../src/app/office-editor/fonts/thumbs.mjs';

const ROOT = resolve(import.meta.dirname, '..');
const work = mkdtempSync(join(tmpdir(), 'officekit-fonts-'));
after(async () => { await stopFontsWork(); rmSync(work, { recursive: true, force: true }); });

// ── 도우미 ─────────────────────────────────────────────

/** AllFonts.js 읽기 → { files, infos, ranges, sel:[레코드] } */
function parseAllFonts(src) {
  src = src.replace(/^\uFEFF/, '');
  const grab = (name) => { const m = new RegExp(`window\\["${name}"\\] = \\[([\\s\\S]*?)\\];`).exec(src); return m ? JSON.parse(`[${m[1]}]`) : null; };
  const b = Buffer.from(/window\["g_fonts_selection_bin"\] = "([^"]*)"/.exec(src)[1], 'base64');
  let o = 0;
  const i32 = () => { const v = b.readInt32LE(o); o += 4; return v; };
  const str = () => { const n = i32(); const s = b.toString('utf8', o, o + n); o += n; return s; };
  const sel = [];
  for (let n = i32(), k = 0; k < n; k++) {
    const start = o; const len = i32();
    const r = { name: str(), names: [] };
    for (let c = i32(), j = 0; j < c; j++) r.names.push(str());
    r.path = str(); r.face = i32(); r.italic = i32(); r.bold = i32(); r.fixed = i32();
    const panoseLen = i32(); o += panoseLen;
    r.ur = [0, 1, 2, 3].map(() => { const v = b.readUInt32LE(o); o += 4; return v; });
    r.cp = [0, 1].map(() => { const v = b.readUInt32LE(o); o += 4; return v; });
    r.weight = b.readUInt16LE(o); r.format = b.readInt16LE(o + 6);
    r.metrics = [8, 10, 12, 14, 16, 18].map((d) => b.readInt16LE(o + d));
    r.fsType = b.readUInt16LE(o + 20);
    r.len = len;
    o = start + len;
    sel.push(r);
  }
  return { version: /__all_fonts_js_version__"\] = (\d+)/.exec(src)[1], files: grab('__fonts_files'), infos: grab('__fonts_infos'), ranges: grab('__fonts_ranges'), sel };
}

/** 범위 → 코드에서 글꼴 이름 */
const rangeFont = (out, code) => { for (let i = 0; i < out.ranges.length; i += 3) if (out.ranges[i] <= code && out.ranges[i + 1] >= code) return out.infos[out.ranges[i + 2]][0]; return null; };

/**
 * 시험용 작은 TrueType 글꼴 (글자 모양은 비어 있음).
 * cmaps: [{ p, e, fmt: 4|6|12, codes:[…], lengthOverride? }]
 */
function buildFont({
  family, style = 'Regular', localized = [], bold = false, italic = false, weight = 400, upem = 2048, numGlyphs: ng,
  cmaps = [{ p: 3, e: 1, fmt: 4, codes: [0x20, 0x41, 0x42, 0x43] }], cp1 = 1, fixed = false, os2Version = 4, nameRecords = null,
}) {
  const t = {};
  // 글리프 번호가 numGlyphs보다 작아야 FreeType가 문자로 센다
  const numGlyphs = ng ?? Math.min(65535, Math.max(300, ...cmaps.map((c) => c.codes.length + 2)));
  const head = Buffer.alloc(54); head.writeUInt32BE(0x00010000, 0); head.writeUInt32BE(0x5F0F3CF5, 12); head.writeUInt16BE(upem, 18);
  head.writeUInt16BE((bold ? 1 : 0) | (italic ? 2 : 0), 44); t.head = head;
  const maxp = Buffer.alloc(6); maxp.writeUInt32BE(0x00005000, 0); maxp.writeUInt16BE(numGlyphs, 4); t.maxp = maxp;
  t.hhea = Buffer.alloc(36); t.hhea.writeUInt32BE(0x00010000, 0);
  const post = Buffer.alloc(32); post.writeUInt32BE(0x00030000, 0); post.writeUInt32BE(fixed ? 1 : 0, 12); t.post = post;
  const os2 = Buffer.alloc(os2Version >= 2 ? 96 : 78 + 8);
  os2.writeUInt16BE(os2Version, 0); os2.writeInt16BE(1000, 2); os2.writeUInt16BE(weight, 4); os2.writeUInt16BE(5, 6);
  os2.writeUInt16BE(8, 8); os2.writeInt16BE(0x0805, 30); [2, 11, 6, 4, 2, 2, 2, 2, 2, 4].forEach((v, i) => { os2[32 + i] = v; });
  os2.writeUInt32BE(0x80000003, 42);
  os2.writeUInt16BE((italic ? 1 : 0) | (bold ? 32 : 0) | (!bold && !italic ? 64 : 0), 62);
  os2.writeInt16BE(1638, 68); os2.writeInt16BE(-410, 70); os2.writeInt16BE(307, 72);
  if (os2Version >= 1) os2.writeUInt32BE(cp1 >>> 0, 78);
  if (os2Version >= 2) { os2.writeInt16BE(1062, 86); os2.writeInt16BE(1467, 88); }
  t['OS/2'] = os2;
  // name
  const recs = nameRecords ?? [
    [3, 1, 0x409, 1, family], [3, 1, 0x409, 2, style],
    ...localized.map(([lang, v]) => [3, 1, lang, 1, v]),
  ];
  const strs = recs.map((r) => Buffer.from(r[4], 'utf16le').swap16());
  const nameHead = Buffer.alloc(6 + 12 * recs.length);
  nameHead.writeUInt16BE(0, 0); nameHead.writeUInt16BE(recs.length, 2); nameHead.writeUInt16BE(6 + 12 * recs.length, 4);
  let so = 0;
  recs.forEach((r, i) => {
    const o = 6 + 12 * i;
    nameHead.writeUInt16BE(r[0], o); nameHead.writeUInt16BE(r[1], o + 2); nameHead.writeUInt16BE(r[2], o + 4); nameHead.writeUInt16BE(r[3], o + 6);
    nameHead.writeUInt16BE(strs[i].length, o + 8); nameHead.writeUInt16BE(so, o + 10); so += strs[i].length;
  });
  t.name = Buffer.concat([nameHead, ...strs]);
  // cmap
  const subs = cmaps.map((c) => {
    const codes = [...c.codes].sort((a, b) => a - b);
    if (c.fmt === 4) {
      const segs = []; // [start, end, delta]
      let g = 1;
      for (let i = 0; i < codes.length;) {
        let j = i; while (j + 1 < codes.length && codes[j + 1] === codes[j] + 1) j++;
        segs.push([codes[i], codes[j], (g - codes[i]) & 0xffff]); g += j - i + 1; i = j + 1;
      }
      segs.push([0xffff, 0xffff, 1]);
      const n = segs.length; const b = Buffer.alloc(16 + 8 * n);
      b.writeUInt16BE(4, 0); b.writeUInt16BE(b.length, 2); b.writeUInt16BE(2 * n, 6);
      segs.forEach(([s, e, d], i) => { b.writeUInt16BE(e, 14 + 2 * i); b.writeUInt16BE(s, 16 + 2 * n + 2 * i); b.writeUInt16BE(d, 16 + 4 * n + 2 * i); });
      return b;
    }
    if (c.fmt === 6) {
      const first = codes[0]; const count = codes.at(-1) - first + 1; const b = Buffer.alloc(10 + 2 * count);
      b.writeUInt16BE(6, 0); b.writeUInt16BE(c.lengthOverride ?? b.length, 2); b.writeUInt16BE(first, 6); b.writeUInt16BE(count, 8);
      codes.forEach((code, i) => b.writeUInt16BE(i + 1, 10 + 2 * (code - first)));
      return b;
    }
    // 12
    const b = Buffer.alloc(16 + 12 * codes.length);
    b.writeUInt16BE(12, 0); b.writeUInt32BE(b.length, 4); b.writeUInt32BE(codes.length, 12);
    codes.forEach((code, i) => { b.writeUInt32BE(code, 16 + 12 * i); b.writeUInt32BE(code, 20 + 12 * i); b.writeUInt32BE(i + 1, 24 + 12 * i); });
    return b;
  });
  const cmHead = Buffer.alloc(4 + 8 * subs.length); cmHead.writeUInt16BE(subs.length, 2);
  let off = cmHead.length;
  subs.forEach((s, i) => { cmHead.writeUInt16BE(cmaps[i].p, 4 + 8 * i); cmHead.writeUInt16BE(cmaps[i].e, 6 + 8 * i); cmHead.writeUInt32BE(off, 8 + 8 * i); off += s.length; });
  t.cmap = Buffer.concat([cmHead, ...subs]);
  t.glyf = Buffer.alloc(4); t.loca = Buffer.alloc((numGlyphs + 1) * 2);
  // sfnt
  const tags = Object.keys(t).sort();
  const dir = Buffer.alloc(12 + 16 * tags.length); dir.writeUInt32BE(0x00010000, 0); dir.writeUInt16BE(tags.length, 4);
  const bodies = []; let pos = dir.length;
  tags.forEach((tag, i) => {
    const data = t[tag]; const pad = Buffer.alloc((4 - (data.length % 4)) % 4);
    dir.write(tag.padEnd(4, ' '), 12 + 16 * i, 'latin1'); dir.writeUInt32BE(pos, 20 + 16 * i); dir.writeUInt32BE(data.length, 24 + 16 * i);
    bodies.push(data, pad); pos += data.length + pad.length;
  });
  return Buffer.concat([dir, ...bodies]);
}
const mem = (path, buf) => ({ path, size: buf.length, source: bufferSource(buf) });
const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

/** 편집기 묶음(글꼴 포함)을 찾는다: W1 캐시(OFFICEKIT_OFFICE_EDITOR_CACHE, tmp/office-editor-build) → 판정 실험 묶음 */
function findBundleEditor() {
  const cands = [];
  for (const c of [process.env.OFFICEKIT_OFFICE_EDITOR_CACHE, join(ROOT, 'tmp', 'office-editor-build')].filter(Boolean)) {
    try { for (const e of readdirSync(c)) cands.push(join(c, e, 'editor')); } catch {}
    cands.push(join(c, 'editor'));
  }
  cands.push(join(ROOT, 'tmp', 'spike-filex', 'ui', 'editor'));
  return cands.find((d) => existsSync(join(d, 'sdkjs', 'common', 'AllFonts.js')) && existsSync(join(d, 'fonts', '000')));
}
const bundle = findBundleEditor();

/** 묶음 글꼴(앞 32바이트 XOR)을 풀어 메모리 원천으로, 원래 경로는 묶음 선택 이진에 적힌 것 */
function bundleInputs(ref) {
  const paths = []; for (const r of ref.sel) if (!paths.includes(r.path)) paths.push(r.path);
  assert.equal(paths.length, ref.files.length, '묶음 파일 수');
  const files = ref.files.map((id, i) => {
    const b = Buffer.from(readFileSync(join(bundle, 'fonts', id)));
    for (let k = 0; k < 32 && k < b.length; k++) b[k] ^= FONT_XOR_KEY[k % 16];
    return mem(paths[i], b);
  });
  return files.sort((x, y) => (x.path < y.path ? -1 : x.path > y.path ? 1 : 0));
}

// ── 기준 시험 ───────────────────────────────────────────

test('기준: 묶음 core-fonts로 만든 AllFonts.js가 묶음 것(ONLYOFFICE allfontsgen 9.4)과 바이트까지 같다', { skip: !bundle && '편집기 묶음 캐시가 없음 (npm run office-editor:build)' }, () => {
  const refText = readFileSync(join(bundle, 'sdkjs', 'common', 'AllFonts.js'), 'utf8');
  const ref = parseAllFonts(refText);
  const files = bundleInputs(ref);
  const out = generateAllFonts(files, { profile: 'onlyoffice', selectionPath: (f) => f.path });
  const mine = parseAllFonts(out.js);
  assert.equal(mine.version, '2');
  assert.deepEqual(mine.files, ref.files, '__fonts_files');
  assert.deepEqual(mine.infos, ref.infos, '__fonts_infos');
  assert.deepEqual(mine.ranges, ref.ranges, '__fonts_ranges');
  assert.deepEqual(mine.sel, ref.sel, 'g_fonts_selection_bin');
  assert.equal(out.js, refText, 'AllFonts.js 전체');
  assert.equal(out.stats.files, ref.files.length);
});

test('기준: Office Kit 설정도 한국어 글꼴이 없으면 같은 목록·범위, 선택 이진 경로만 웹 번호', { skip: !bundle && '편집기 묶음 캐시가 없음' }, () => {
  const ref = parseAllFonts(readFileSync(join(bundle, 'sdkjs', 'common', 'AllFonts.js'), 'utf8'));
  const out = generateAllFonts(bundleInputs(ref), { profile: 'officekit' });
  const mine = parseAllFonts(out.js);
  assert.deepEqual(mine.files, ref.files);
  assert.deepEqual(mine.infos, ref.infos);
  assert.deepEqual(mine.ranges, ref.ranges);
  assert.equal(mine.sel.length, ref.sel.length);
  // sdkjs ListMap[파일 번호]가 맞도록 선택 레코드 경로 = 그 얼굴의 웹 파일 번호
  mine.sel.forEach((r, i) => {
    assert.deepEqual({ ...r, path: '', len: 0 }, { ...ref.sel[i], path: '', len: 0 });
    assert.match(r.path, /^\d{3}$/);
  });
  const fileOfFace = new Map(); // (번호, 얼굴) → 가족
  ref.infos.forEach((x) => { for (let k = 1; k < 9; k += 2) if (x[k] >= 0) fileOfFace.set(`${mine.files[x[k]]}/${x[k + 1]}`, x[0]); });
  for (const r of mine.sel) assert.equal(fileOfFace.get(`${r.path}/${r.face}`), r.name, `${r.name} → ${r.path}`);
});

// ── 글꼴 읽기 (sfnt.mjs) ────────────────────────────────

test('sfnt: FreeType 이름 고르기 (Windows 영어 우선, ASCII 밖은 ?), 얼굴 값 읽기', () => {
  const buf = buildFont({ family: 'Test Sans', localized: [[0x412, '시험 산스']], bold: true, weight: 700, fixed: true });
  assert.deepEqual(faceOffsets(buf), [0]);
  const f = readFace(buf, 0, { codes: 'all' });
  assert.equal(freetypeName(f.names, 1), 'Test Sans');
  assert.equal(freetypeName(f.names, 2), 'Regular');
  assert.equal(f.os2.usWeightClass, 700);
  assert.equal(f.unitsPerEm, 2048);
  assert.ok(f.hasGlyf);
  assert.deepEqual(f.charmaps[0].codes, [0x20, 0x41, 0x42, 0x43]);
  // 영어 이름이 없으면 첫 Windows 이름을 ASCII로 (한글은 '?')
  const ko = readFace(buildFont({ family: 'x', nameRecords: [[3, 1, 0x412, 1, '한글체'], [3, 1, 0x412, 2, 'Regular']] }));
  assert.equal(freetypeName(ko.names, 1), '???');
});

test('sfnt: FreeType처럼 깨진 cmap 부속 표는 버린다 (길이가 넘친 형식 6)', () => {
  const codes = range(0x20, 0x7e);
  const ok = readFace(buildFont({ family: 'A', cmaps: [{ p: 1, e: 3, fmt: 6, codes }, { p: 3, e: 1, fmt: 4, codes }] }), 0, { codes: 'all' });
  assert.equal(ok.charmaps.length, 2);
  const bad = readFace(buildFont({ family: 'A', cmaps: [{ p: 1, e: 3, fmt: 6, codes, lengthOverride: 20 }, { p: 3, e: 1, fmt: 4, codes }] }), 0, { codes: 'all' });
  assert.deepEqual(bad.charmaps.map((c) => [c.platformID, c.encodingID]), [[3, 1]]);
  // 형식 12(UCS-4)가 있으면 유니코드 기본 표는 그것
  const u4 = readFace(buildFont({ family: 'B', cmaps: [{ p: 3, e: 1, fmt: 4, codes: [0x41] }, { p: 3, e: 10, fmt: 12, codes: [0x41, 0x1F600] }] }), 0, { codes: 'unicode' });
  assert.equal(u4.unicode.format, 12);
  assert.deepEqual(u4.unicode.codes, [0x41, 0x1F600]);
});

// ── 목록 만들기 규칙 (allfonts.mjs) ──────────────────────

test('목록: 굵게·기울임 칸, 같은 이름·스타일은 표준 스타일 이름이 이김, 이름순(ASCII만 소문자로)', () => {
  const files = [
    mem('/f/a-book.ttf', buildFont({ family: 'Alpha', style: 'Book' })),
    mem('/f/a-bold.ttf', buildFont({ family: 'Alpha', style: 'Bold', bold: true, weight: 700 })),
    mem('/f/a-it.ttf', buildFont({ family: 'Alpha', style: 'Italic', italic: true })),
    mem('/f/a-reg.ttf', buildFont({ family: 'Alpha', style: 'Regular' })), // 'Book'을 바꿔 넣는다
    mem('/f/a-reg2.ttf', buildFont({ family: 'Alpha', style: 'Regular' })), // 'Regular'는 그대로 둔다
    mem('/f/b.ttf', buildFont({ family: 'beta' })),
    mem('/f/c.ttf', buildFont({ family: 'Gamma', bold: true, italic: true, style: 'Bold Italic', weight: 700 })),
  ].sort((x, y) => (x.path < y.path ? -1 : 1));
  const out = generateAllFonts(files, { profile: 'onlyoffice' });
  assert.deepEqual(out.infos.map((x) => x[0]), ['Alpha', 'beta', 'Gamma']);
  const alpha = out.infos[0];
  const fileOf = (i) => out.files.find((f) => f.id === out.fontsFiles[i]).path;
  assert.equal(fileOf(alpha[1]), '/f/a-reg.ttf', 'R');
  assert.equal(fileOf(alpha[3]), '/f/a-it.ttf', 'I');
  assert.equal(fileOf(alpha[5]), '/f/a-bold.ttf', 'B');
  assert.equal(alpha[7], -1, 'BI 없음');
  assert.deepEqual(out.infos[2].slice(1, 7), [-1, -1, -1, -1, -1, -1]);
  assert.equal(out.stats.faces, 5);
  assert.ok(compareFontNames('Asana Math', 'ASCW3') < 0 && compareFontNames('mry_KacstQurn', 'Mukti Narrow') < 0);
});

test('목록: 선택 이진 v2 레코드 (이름·다른 언어 이름·지표 1000 단위·기호 글꼴 코드 페이지)', () => {
  const files = [
    mem('/f/k.ttf', buildFont({ family: 'Kor Font', localized: [[0x412, '한국 글꼴'], [0x411, '日本語名']] })),
    mem('/f/s.ttf', buildFont({ family: 'Sym Font', cp1: 0, os2Version: 0, cmaps: [{ p: 3, e: 0, fmt: 4, codes: range(0xF020, 0xF0FF) }] })),
  ];
  const strict = parseAllFonts(generateAllFonts(files, { profile: 'onlyoffice' }).js);
  const k = strict.sel.find((r) => r.name === 'Kor Font');
  assert.deepEqual(k.names, ['한국 글꼴'], '9.4: 다른 언어 이름은 하나');
  assert.deepEqual(k.metrics, [488, 799, -200, 149, 518, 716], 'Math.trunc(v * (1000 / upem))');
  assert.equal(k.format, 1);
  assert.equal(k.len, k.len | 0);
  const s = strict.sel.find((r) => r.name === 'Sym Font');
  assert.equal(s.cp[0], 0x80000000, 'cmap (3,0)이면 기호 코드 페이지');
  const ok = parseAllFonts(generateAllFonts(files, { profile: 'officekit' }).js);
  assert.deepEqual(ok.sel.find((r) => r.name === 'Kor Font').names, ['한국 글꼴', '日本語名'], 'Office Kit: 모든 다른 언어 이름, 한국어 먼저');
});

test('범위: 짧은 구간(10자 이하)은 벌점, 1순위 글꼴(ASCW3 또는 Arial)은 예외, 마지막 끝은 0x10FFFE', () => {
  const fam = (name, file) => ({ name, slots: { R: { file, faceIndex: 0 }, I: null, B: null, BI: null } });
  const arial = mem('/f/arial.ttf', buildFont({ family: 'Arial', cmaps: [{ p: 3, e: 1, fmt: 4, codes: [...range(0x41, 0x5a), 0x100] }] }));
  const big = mem('/f/big.ttf', buildFont({ family: 'Big', cmaps: [{ p: 3, e: 10, fmt: 12, codes: [...range(0xfe, 0x120), 0x10FFFE] }] }));
  const fams = [fam('Arial', arial), fam('Big', big)];
  const r = computeRanges(fams, ONLYOFFICE_PRIORITY, (f) => f.size);
  const at = (c) => { for (let i = 0; i < r.length; i += 3) if (r[i] <= c && r[i + 1] >= c) return fams[r[i + 2]].name; return null; };
  assert.equal(at(0x41), 'Arial');
  assert.equal(at(0x100), 'Arial', 'Arial이 1순위(ASCW3 없음)면 한 글자 구간도 이긴다');
  assert.equal(at(0x101), 'Big');
  assert.equal(at(0x10FFFE), 'Big');
  assert.equal(r.at(-2), 0x10FFFE);
});

test('한국어: 한글 대체는 맑은 고딕, 별칭(맑은고딕·함초롬바탕)은 선택 이진 names에만, _ea 미리보기는 글꼴 자기 이름', () => {
  const hangul = [...range(0x1100, 0x11ff), ...range(0x3131, 0x318e), ...range(0xac00, 0xd7a3)];
  const files = [
    mem('/w/arial.ttf', buildFont({ family: 'Arial', cmaps: [{ p: 3, e: 1, fmt: 4, codes: range(0x20, 0x7e) }] })),
    mem('/w/batang.ttf', buildFont({ family: 'Batang', localized: [[0x412, '바탕']], cmaps: [{ p: 3, e: 1, fmt: 4, codes: [...range(0x20, 0x7e), ...hangul, ...range(0x4e00, 0x4fff)] }] })),
    mem('/w/malgun.ttf', buildFont({ family: 'Malgun Gothic', localized: [[0x412, '맑은 고딕']], cmaps: [{ p: 3, e: 1, fmt: 4, codes: [...range(0x20, 0x7e), ...hangul] }] })),
  ];
  const ok = generateAllFonts(files, { profile: 'officekit' });
  assert.equal(ok.stats.hangulFallback, 'Malgun Gothic');
  assert.equal(rangeFont(ok, 0xac00), 'Malgun Gothic');
  assert.equal(rangeFont(ok, 0x3131), 'Malgun Gothic');
  assert.equal(rangeFont(ok, 0x4e00), 'Batang', '한자는 바탕');
  const strict = generateAllFonts(files, { profile: 'onlyoffice' });
  assert.equal(rangeFont(strict, 0xac00), 'Batang', 'ONLYOFFICE 원래 순서는 바탕');
  const malgun = ok.selection.find((r) => r.name === 'Malgun Gothic');
  assert.deepEqual(malgun.names, ['맑은 고딕', '맑은고딕', '함초롬돋움', '한컴돋움'], '자기 이름 먼저, 그다음 별칭 (HCR Dotum이 없으면 함초롬돋움도)');
  assert.ok(ok.selection.find((r) => r.name === 'Batang').names.includes('함초롬바탕'), 'HCR Batang이 없으면 바탕이 함초롬바탕');
  assert.deepEqual(ok.stats.aliases.find((a) => a[0] === '함초롬바탕'), ['함초롬바탕', 'Batang']);
  assert.deepEqual(ok.infos.map((x) => x[0]), ['Arial', 'Batang', 'Malgun Gothic'], '별칭은 글꼴 목록에 따로 넣지 않는다');
  assert.equal(ok.thumbs.find((t) => t.name === 'Malgun Gothic').eaName, '맑은 고딕');
  assert.equal(ok.thumbs.find((t) => t.name === 'Arial').eaName, null);
  // 그 이름을 가진 글꼴이 있으면 별칭을 붙이지 않는다
  const withHcr = generateAllFonts([...files, mem('/w/hcr.ttf', buildFont({ family: 'HCR Batang', localized: [[0x412, '함초롬바탕']] }))], { profile: 'officekit' });
  assert.ok(!withHcr.selection.find((r) => r.name === 'Batang').names.includes('함초롬바탕'));
});

const WIN_FONTS = process.platform === 'win32' ? defaultFontDirs()[0] : null;
const hasWin = (n) => !!WIN_FONTS && existsSync(join(WIN_FONTS, n));
test('Windows 글꼴: 맑은 고딕·바탕·굴림 실제 파일로 한국어 이름·한글 대체', { skip: !(hasWin('malgun.ttf') && hasWin('batang.ttc') && hasWin('gulim.ttc')) && 'Windows 한국어 글꼴이 없음' }, () => {
  const files = listFontFiles([WIN_FONTS], { recursive: false })
    .filter((f) => /[\\/](malgun\.ttf|malgunbd\.ttf|batang\.ttc|gulim\.ttc|arial\.ttf|times\.ttf)$/i.test(f.path));
  const out = generateAllFonts(files, { profile: 'officekit' });
  const names = out.infos.map((x) => x[0]);
  for (const n of ['Malgun Gothic', 'Batang', 'BatangChe', 'Gungsuh', 'Gulim', 'Dotum']) assert.ok(names.includes(n), n);
  assert.equal(out.stats.hangulFallback, 'Malgun Gothic');
  const sel = parseAllFonts(out.js).sel;
  const has = (en, ko) => assert.ok(sel.some((r) => r.name === en && r.names.includes(ko)), `${en} ← ${ko}`);
  has('Malgun Gothic', '맑은 고딕'); has('Batang', '바탕'); has('Gulim', '굴림'); has('Dotum', '돋움'); has('Gungsuh', '궁서');
  const m = out.infos.find((x) => x[0] === 'Malgun Gothic');
  assert.ok(m[1] >= 0 && m[5] >= 0, '맑은 고딕 보통·굵게');
});

// ── 미리보기 그림 (thumbs.mjs) ───────────────────────────

test('미리보기: 파일 이름·크기가 편집기(ComboBoxFonts)와 같고 .bin 압축이 되돌려진다', () => {
  assert.deepEqual(THUMB_FILES.slice(0, 5), ['fonts_thumbnail.png.bin', 'fonts_thumbnail@1.25x.png.bin', 'fonts_thumbnail@1.5x.png.bin', 'fonts_thumbnail@1.75x.png.bin', 'fonts_thumbnail@2x.png.bin']);
  assert.equal(thumbFileName(1.5, true, false), 'fonts_thumbnail_ea@1.5x.png');
  assert.deepEqual([1, 1.25, 1.5, 1.75, 2].map((s) => [thumbGeometry(s).width, thumbGeometry(s).row]), [[300, 28], [375, 35], [450, 42], [525, 49], [600, 56]]);
  const a = new Uint8Array(300 * 28 * 3);
  for (let i = 0; i < a.length; i += 97) a[i] = (i % 255) + 1;
  a.fill(0, 1000, 1700); // 255보다 긴 0
  const bin = encodeSpriteBin(a, 300, 28, 3);
  assert.deepEqual([bin.readUInt32BE(0), bin.readUInt32BE(4), bin.readUInt32BE(8)], [300, 28, 3]);
  const d = decodeSpriteBin(bin);
  assert.deepEqual(Buffer.from(d.alpha), Buffer.from(a));
  // 빈 그림: 모두 0, 줄 수 = 글꼴 수
  const blank = blankSprites(5);
  assert.equal(blank.size, 10);
  const b2 = decodeSpriteBin(blank.get('fonts_thumbnail_ea@2x.png.bin'));
  assert.deepEqual([b2.width, b2.row, b2.count], [600, 56, 5]);
  assert.ok(b2.alpha.every((v) => v === 0));
  // 줄이기: 꽉 찬 칸은 꽉 찬 그대로, 빈 칸은 빈 그대로
  const full = new Uint8Array(600 * 56).fill(255);
  assert.ok(downsampleAlpha(full, 600, 56, 375, 35).every((v) => v === 255));
  assert.ok(downsampleAlpha(new Uint8Array(600 * 56), 600, 56, 300, 28).every((v) => v === 0));
});

const realFont = ['arial.ttf', 'malgun.ttf'].every(hasWin) ? { latin: join(WIN_FONTS, 'arial.ttf'), hangul: join(WIN_FONTS, 'malgun.ttf') } : null;
test('미리보기: resvg로 글꼴 이름을 그린다 (_ea는 한국어 이름)', { skip: !realFont && 'Arial·맑은 고딕 글꼴 파일이 없음' }, async () => {
  const thumbs = [
    { name: 'Arial', eaName: null, path: realFont.latin, faceIndex: 0, symbolic: false, css: ['Arial'], weight: 400, italic: false },
    { name: 'Malgun Gothic', eaName: '맑은 고딕', path: realFont.hangul, faceIndex: 0, symbolic: false, css: ['Malgun Gothic'], weight: 400, italic: false },
  ];
  const out = join(work, 'thumbs');
  const written = await renderThumbnails(thumbs, { outDir: out, mode: 'full', fallback: realFont });
  assert.equal(written.length, 10);
  const ink = (name, row) => { const d = decodeSpriteBin(readFileSync(join(out, name))); const n = d.width * d.row; return d.alpha.subarray(row * n, (row + 1) * n).reduce((s, v) => s + v, 0); };
  for (const f of THUMB_FILES) { assert.ok(ink(f, 0) > 1000, `${f} 0줄`); assert.ok(ink(f, 1) > 1000, `${f} 1줄`); }
  // 한국어 이름 줄은 영어 이름 줄과 다르다
  const en = decodeSpriteBin(readFileSync(join(out, 'fonts_thumbnail@2x.png.bin')));
  const ea = decodeSpriteBin(readFileSync(join(out, 'fonts_thumbnail_ea@2x.png.bin')));
  const rowOf = (d, r) => Buffer.from(d.alpha.subarray(r * 600 * 56, (r + 1) * 600 * 56));
  assert.ok(rowOf(en, 0).equals(rowOf(ea, 0)), 'Arial은 같은 그림');
  assert.ok(!rowOf(en, 1).equals(rowOf(ea, 1)), '맑은 고딕은 _ea가 한국어 이름');
});

// ── 캐시 · ensureFonts · 글꼴 내주기 ─────────────────────

/** 시험 글꼴 폴더 (작은 가짜 글꼴) */
function makeFontDir(name, fonts) {
  const d = join(work, name); mkdirSync(d, { recursive: true });
  for (const [file, opts] of fonts) writeFileSync(join(d, file), buildFont(opts));
  return d;
}

test('ensureFonts: worker에서 만들고 캐시를 다시 쓴다, W3 서버 약속(allFontsJsPath·thumbnailsDir·fontFileByIndex)', async () => {
  const fontDir = makeFontDir('fonts-a', [
    ['a.ttf', { family: 'Alpha' }], ['b.ttf', { family: 'Beta', localized: [[0x412, '베타']] }],
    ['readme.txt', { family: 'x' }],
  ]);
  writeFileSync(join(fontDir, 'readme.txt'), '글꼴 아님');
  const userData = join(work, 'ud');
  const logs = [];
  const t0 = Date.now();
  const stale = join(userData, 'office-editor', 'fonts', 'old-0000000000000000');
  mkdirSync(stale, { recursive: true });
  const f = await ensureFonts({ userData, editorVersion: '9.4.0.129/x', fontDirs: [fontDir], log: (m) => logs.push(m), thumbDelayMs: 0 });
  assert.ok(!existsSync(stale), '지난 실행의 캐시는 지운다');
  assert.ok(Date.now() - t0 < 20000);
  assert.match(f.key, /^9\.4\.0\.129_x-[0-9a-f]{16}$/);
  assert.equal(f.dir, join(userData, 'office-editor', 'fonts', f.key));
  assert.ok(logs.some((m) => /글꼴 목록: 파일 2개/.test(m)), logs.join('\n'));
  // AllFonts.js
  const js = parseAllFonts(readFileSync(f.allFontsJsPath, 'utf8'));
  assert.deepEqual(js.files, ['000', '001']);
  assert.deepEqual(js.infos.map((x) => x[0]), ['Alpha', 'Beta']);
  assert.deepEqual(js.sel.map((r) => r.path), ['000', '001']);
  // 미리보기 10개가 늘 있다 (빈 그림부터)
  for (const n of THUMB_FILES) assert.ok(existsSync(join(f.thumbnailsDir, n)), n);
  assert.equal(decodeSpriteBin(readFileSync(join(f.thumbnailsDir, 'fonts_thumbnail.png.bin'))).count, 2);
  assert.equal(f.thumbnailPath('fonts_thumbnail@2x.png.bin'), join(f.thumbnailsDir, 'fonts_thumbnail@2x.png.bin'));
  assert.equal(f.thumbnailPath('../fonts.json'), null);
  // 글꼴 번호 → 실제 파일 (절대 경로 문자열)
  assert.equal(f.fontFileByIndex(0), join(fontDir, 'a.ttf'));
  assert.equal(f.fontFileByIndex('001'), join(fontDir, 'b.ttf'));
  for (const bad of [2, -1, 1.5, '1a', '../a', null, undefined, {}]) assert.equal(f.fontFileByIndex(bad), null, String(bad));
  // resolve: 편집기 경로 → 파일
  assert.equal(f.resolve('sdkjs/common/AllFonts.js').file, f.allFontsJsPath);
  assert.equal(f.resolve('sdkjs/common/Images/fonts_thumbnail_ea@1.5x.png.bin').contentType, 'application/octet-stream');
  assert.equal(f.resolve('fonts/001').size, statSync(join(fontDir, 'b.ttf')).size);
  for (const bad of ['fonts/9', 'fonts/../x', 'sdkjs/common/Images/x.png', 'sdkjs/common/AllFonts.json', 'index.html']) assert.equal(f.resolve(bad), null, bad);
  // XOR 스트림 = 원본의 앞 32바이트만 XOR
  const orig = readFileSync(join(fontDir, 'b.ttf'));
  const chunks = []; for await (const c of f.resolve('fonts/001').open()) chunks.push(c);
  assert.deepEqual(Buffer.concat(chunks), obfuscateFontBytes(orig));
  assert.deepEqual(obfuscateFontBytes(obfuscateFontBytes(orig)), orig);
  // 미리보기 그리기가 끝나면 full
  assert.equal(await f.thumbnails, 'full');
  assert.equal(f.thumbnailsStage(), 'full');
  // 다시 부르면 같은 캐시 (다시 만들지 않음)
  const mtime = statSync(f.allFontsJsPath).mtimeMs;
  const g = await ensureFonts({ userData, editorVersion: '9.4.0.129/x', fontDirs: [fontDir] });
  assert.equal(g.dir, f.dir);
  assert.equal(statSync(g.allFontsJsPath).mtimeMs, mtime);
  // 글꼴 폴더가 바뀌면 새 캐시 (이 프로세스가 내주고 있는 옛 캐시는 남긴다)
  writeFileSync(join(fontDir, 'c.ttf'), buildFont({ family: 'Gamma' }));
  const h = await ensureFonts({ userData, editorVersion: '9.4.0.129/x', fontDirs: [fontDir], thumbnails: false });
  assert.notEqual(h.key, f.key);
  assert.deepEqual(parseAllFonts(readFileSync(h.allFontsJsPath, 'utf8')).infos.map((x) => x[0]), ['Alpha', 'Beta', 'Gamma']);
  assert.ok(existsSync(f.allFontsJsPath), '쓰고 있는 캐시는 지우지 않는다');
  // 그 뒤에 글꼴 파일이 바뀌면 내주지 않는다
  writeFileSync(join(fontDir, 'a.ttf'), Buffer.concat([readFileSync(join(fontDir, 'a.ttf')), Buffer.alloc(4)]));
  assert.equal(h.fontFileByIndex(0), null);
});

test('캐시 키: 편집기 버전·파일 크기·수정 시각이 바뀌면 달라진다', () => {
  const listing = [{ path: 'C:\\Windows\\Fonts\\a.ttf', size: 10, mtimeMs: 1 }];
  const k = fontsCacheKey({ editorVersion: '9.4', listing });
  assert.equal(k, fontsCacheKey({ editorVersion: '9.4', listing: [{ ...listing[0] }] }));
  assert.notEqual(k, fontsCacheKey({ editorVersion: '9.5', listing }));
  assert.notEqual(k, fontsCacheKey({ editorVersion: '9.4', listing: [{ ...listing[0], size: 11 }] }));
  assert.notEqual(k, fontsCacheKey({ editorVersion: '9.4', listing: [{ ...listing[0], mtimeMs: 2 }] }));
  assert.notEqual(k, fontsCacheKey({ editorVersion: '9.4', listing, profile: 'onlyoffice' }));
  assert.match(fontsCacheKey({ editorVersion: '../..', listing }), /^\.\._\.\.-[0-9a-f]{16}$/);
});

test('ensureFonts: 글꼴 캐시 폴더가 없으면 한국어 오류', async () => {
  await assert.rejects(() => ensureFonts({ fontDirs: [] }), { code: 'FONTS_ARGS', message: /userData/ });
});

test('XOR 스트림: 파일이 없으면 오류를 낸다, 32바이트보다 짧은 파일도 된다', async () => {
  const p = join(work, 'short.bin'); writeFileSync(p, Buffer.from([1, 2, 3]));
  const chunks = []; for await (const c of createObfuscatedStream(p)) chunks.push(c);
  assert.deepEqual(Buffer.concat(chunks), Buffer.from([1 ^ 0xA0, 2 ^ 0x66, 3 ^ 0xD6]));
  await assert.rejects(async () => { for await (const _ of createObfuscatedStream(join(work, 'nope.ttf'))) { /* 없음 */ } }, { code: 'ENOENT' });
});

test('W3 서버와 잇기: setOfficeFonts(손잡이)로 AllFonts.js·미리보기·XOR 글꼴을 내준다', async (t) => {
  const srvMod = await import('../src/app/editor-server.mjs');
  if (typeof srvMod.startEditorServer !== 'function') { t.skip('editor-server 없음'); return; }
  const fontDir = makeFontDir('fonts-srv', [['z.ttf', { family: 'Zeta' }]]);
  const f = await ensureFonts({ userData: join(work, 'ud-srv'), editorVersion: 'dev', fontDirs: [fontDir], thumbnails: false });
  const office = join(work, 'office-bundle');
  mkdirSync(join(office, 'editor', 'sdkjs', 'common'), { recursive: true });
  writeFileSync(join(office, 'index.html'), '<!doctype html>');
  writeFileSync(join(office, 'editor', 'sdkjs', 'common', 'AllFonts.js'), '/* 묶음 것 */');
  const srv = await srvMod.startEditorServer({ root: work, dirs: { office }, officeFonts: f });
  t.after(() => srv.close());
  if (!srv.officeBase || typeof srv.setOfficeFonts !== 'function') { t.skip(`office 서버 없음: ${srv.officeError}`); return; }
  const get = async (rel) => { const r = await fetch(`${srv.officeBase}${rel}`); return { status: r.status, body: Buffer.from(await r.arrayBuffer()) }; };
  const all = await get('editor/sdkjs/common/AllFonts.js');
  assert.equal(all.status, 200);
  assert.ok(all.body.toString('utf8').includes('"Zeta"'));
  const thumb = await get('editor/sdkjs/common/Images/fonts_thumbnail@1.5x.png.bin');
  assert.equal(thumb.status, 200);
  assert.equal(decodeSpriteBin(thumb.body).count, 1);
  const font = await get('editor/fonts/000');
  assert.equal(font.status, 200);
  assert.deepEqual(font.body, obfuscateFontBytes(readFileSync(join(fontDir, 'z.ttf'))));
  assert.equal((await get('editor/fonts/001')).status, 404);
});

test('OFFICEKIT 설정: 우선순위는 맑은 고딕·바탕을 중국·일본 글꼴 앞으로만 옮긴다', () => {
  assert.equal(OFFICEKIT_PRIORITY.length, ONLYOFFICE_PRIORITY.length);
  assert.ok(OFFICEKIT_PRIORITY.indexOf('Malgun Gothic') < OFFICEKIT_PRIORITY.indexOf('SimSun'));
  assert.ok(OFFICEKIT_PRIORITY.indexOf('Batang') < OFFICEKIT_PRIORITY.indexOf('SimSun'));
  assert.deepEqual(OFFICEKIT_PRIORITY.slice(0, 18), ONLYOFFICE_PRIORITY.slice(0, 18));
});

