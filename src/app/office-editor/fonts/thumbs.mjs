// SPDX-License-Identifier: AGPL-3.0-only
// thumbs.mjs — 글꼴 목록 미리보기 그림 sdkjs/common/Images/fonts_thumbnail[_ea][@1.25x…@2x].png(.bin)
//   한 줄(28px × 배율)에 글꼴 하나, __fonts_infos 순서. 웹 편집기는 .bin(알파 + 0 길이 압축)만 읽는다 (ComboBoxFonts.js).
// 그림 크기·글자 위치·.bin 형식은 ONLYOFFICE core ApplicationFontsWorker.cpp SaveThumbnails
// (Copyright (C) Ascensio System SIA, 2009-2026, AGPL-3.0-only)를 옮긴 것이고, 그리기는 resvg로 바꿨다.
// 수정: 2026-10-18 Office Kit (toxh13).
import { existsSync, mkdirSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { crc32, deflateSync } from 'node:zlib';

export const THUMB_SCALES = [1, 1.25, 1.5, 1.75, 2];

/** 파일 이름 (C++과 같은 규칙: 100 → .png, 200 → @2x.png, 150 → @1.5x.png, 125 → @1.25x.png) */
export function thumbFileName(scale, ea = false, bin = true) {
  const n = Math.trunc(scale * 100 + 0.5);
  let s = `fonts_thumbnail${ea ? '_ea' : ''}`;
  if (n === 100) s += '.png';
  else if (n % 100 === 0) s += `@${n / 100}x.png`;
  else if (n % 10 === 0) s += `@${Math.trunc(n / 100)}.${Math.trunc(n / 10) % 10}x.png`;
  else s += `@${Math.trunc(n / 100)}.${n % 100}x.png`;
  return bin ? `${s}.bin` : s;
}
/** 만드는 그림 파일 10개 (.bin) */
export const THUMB_FILES = [false, true].flatMap((ea) => THUMB_SCALES.map((s) => thumbFileName(s, ea)));
const THUMB_RE = /^fonts_thumbnail(_ea)?(@(1\.25|1\.5|1\.75|2)x)?\.png(\.bin)?$/;
export const isThumbName = (name) => THUMB_RE.test(name);

/** 배율별 크기 (C++: dpi = 96×배율, 글자 14pt, x = 5mm, 기준선 = 줄 아래 − 2mm) */
export function thumbGeometry(scale) {
  const dpi = 96 * scale;
  const row = Math.trunc(28 * scale);
  return { width: Math.trunc(300 * scale), row, fontPx: (14 * dpi) / 72, x: (5 * dpi) / 25.4, baseline: row - (2 * dpi) / 25.4 };
}

/** 알파 값 → .bin (머리 12바이트: 폭·줄 높이·줄 수, 그다음 알파. 0은 [0, 길이(1~255)]로 묶음) */
export function encodeSpriteBin(alpha, width, row, count) {
  const n = width * row * count;
  const out = Buffer.allocUnsafe(12 + n + Math.ceil(n / 255) * 2 + 16);
  out.writeUInt32BE(width, 0); out.writeUInt32BE(row, 4); out.writeUInt32BE(count, 8);
  let o = 12; let i = 0;
  while (i < n) {
    const a = alpha[i];
    out[o++] = a;
    if (a !== 0) { i++; continue; }
    const i0 = i; const limit = Math.min(i0 + 255, n);
    while (i < limit) { i++; if (i >= limit || alpha[i] !== 0) break; }
    out[o++] = i - i0;
  }
  return Buffer.from(out.subarray(0, o));
}

/** .bin → { width, row, count, alpha } (시험·확인용) */
export function decodeSpriteBin(buf) {
  const width = buf.readUInt32BE(0); const row = buf.readUInt32BE(4); const count = buf.readUInt32BE(8);
  const alpha = new Uint8Array(width * row * count);
  let i = 12; let p = 0;
  while (i < buf.length && p < alpha.length) {
    const v = buf[i++];
    if (v === 0) p += buf[i++]; else alpha[p++] = v;
  }
  return { width, row, count, alpha };
}

/** 빈(투명) 미리보기 .bin 10개 → Map(이름 → Buffer). 그림이 다 되기 전에 내준다 */
export function blankSprites(count) {
  const m = new Map();
  for (const ea of [false, true]) for (const s of THUMB_SCALES) {
    const g = thumbGeometry(s);
    m.set(thumbFileName(s, ea), encodeSpriteBin(new Uint8Array(g.width * g.row * Math.max(count, 0)), g.width, g.row, count));
  }
  return m;
}

/** RGBA → PNG (확인용) */
export function encodePng(width, height, rgba) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) { raw[y * (width * 4 + 1)] = 0; Buffer.from(rgba.buffer, rgba.byteOffset + y * width * 4, width * 4).copy(raw, y * (width * 4 + 1) + 1); }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, { level: 6 })), chunk('IEND', Buffer.alloc(0))]);
}

/** .bin 하나 → 흰 바탕 검은 글씨 PNG (사람 눈으로 확인용) */
export function spriteBinToPng(buf, { rows } = {}) {
  const { width, row, count, alpha } = decodeSpriteBin(buf);
  const n = Math.min(count, rows ?? count);
  const rgba = Buffer.alloc(width * row * n * 4);
  for (let i = 0; i < width * row * n; i++) { const g = 255 - alpha[i]; rgba[4 * i] = g; rgba[4 * i + 1] = g; rgba[4 * i + 2] = g; rgba[4 * i + 3] = 255; }
  return encodePng(width, row * n, rgba);
}

const xml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const cssList = (names) => names.filter((n) => n && !/['"\\]/.test(n)).map((n) => `'${xml(n)}'`).join(', ');

/**
 * 면적 평균으로 줄이기 (알파 한 칸 sw×sh → dw×dh). 큰 배율(2x)로 한 번만 그리고 나머지 배율은 이것으로 만든다
 * — resvg는 <text>마다 글꼴을 다시 읽어 글자 배치를 하므로(큰 CJK 글꼴은 20~40ms) 그리기 횟수를 줄인다
 */
export function downsampleAlpha(src, sw, sh, dw, dh) {
  const out = new Uint8Array(dw * dh);
  // 글자가 있는 상자만 계산한다 (나머지는 투명)
  let minX = sw; let maxX = -1; let minY = sh; let maxY = -1;
  for (let y = 0; y < sh; y++) {
    const row = y * sw;
    for (let x = 0; x < sw; x++) if (src[row + x]) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
  }
  if (maxX < 0) return out;
  const kx = kernel(sw, dw); const ky = kernel(sh, dh);
  const dx0 = Math.max(0, Math.floor(minX / kx.f) - 1); const dx1 = Math.min(dw - 1, Math.ceil((maxX + 1) / kx.f));
  const dy0 = Math.max(0, Math.floor(minY / ky.f) - 1); const dy1 = Math.min(dh - 1, Math.ceil((maxY + 1) / ky.f));
  const tmp = new Float32Array(dw * sh);
  for (let y = minY; y <= maxY; y++) {
    const row = y * sw; const trow = y * dw;
    for (let x = dx0; x <= dx1; x++) {
      let acc = 0; const b = x * kx.K;
      for (let k = 0; k < kx.K; k++) { const wv = kx.w[b + k]; if (wv) acc += src[row + kx.idx[b + k]] * wv; }
      tmp[trow + x] = acc;
    }
  }
  for (let y = dy0; y <= dy1; y++) {
    const b = y * ky.K; const orow = y * dw;
    for (let x = dx0; x <= dx1; x++) {
      let acc = 0;
      for (let k = 0; k < ky.K; k++) { const wv = ky.w[b + k]; if (wv) acc += tmp[ky.idx[b + k] * dw + x] * wv; }
      out[orow + x] = acc >= 255 ? 255 : Math.round(acc);
    }
  }
  return out;
}

// 줄이기 가중치 표 (원래 길이 → 새 길이), 한 번 만들고 다시 쓴다
const KERNELS = new Map();
function kernel(s, d) {
  const key = `${s}>${d}`;
  if (KERNELS.has(key)) return KERNELS.get(key);
  const f = s / d; const K = Math.ceil(f) + 1;
  const idx = new Int32Array(d * K); const w = new Float32Array(d * K);
  for (let x = 0; x < d; x++) {
    const x0 = x * f; const x1 = x0 + f; let k = 0;
    for (let sx = Math.floor(x0); sx < Math.ceil(x1) && sx < s; sx++, k++) { idx[x * K + k] = sx; w[x * K + k] = (Math.min(x1, sx + 1) - Math.max(x0, sx)) / f; }
  }
  const kk = { K, idx, w, f };
  KERNELS.set(key, kk);
  return kk;
}

/** 이름 그리기에 쓸 기본 글꼴 이름들 (줄 끝 대체용) */
const FALLBACK_FAMILIES = ['Arial', 'Segoe UI', 'Liberation Sans', 'DejaVu Sans', 'Malgun Gothic', 'NanumGothic', 'Noto Sans KR', 'Gulim'];

/**
 * 미리보기 그림 만들기 → 쓴 파일 이름 목록 (그만두면 null)
 * thumbs: generateAllFonts().thumbs ([{ name, eaName, path, faceIndex, symbolic, css, weight, italic }], __fonts_infos 순서)
 * opts: {
 *   outDir,                          // .bin을 쓸 폴더
 *   mode: 'full' | 'plain',          // full = 글꼴마다 자기 글꼴로, plain = 모두 기본 글꼴(영어 이름)로 — 빠름
 *   fallback: { latin, hangul, mono } // 기본·대체 글꼴 파일 (없으면 thumbs에서 Arial·Malgun Gothic·Courier New를 찾는다)
 *   batchBytes = 48MB, batchRows = 48, yieldMs = 0, shouldStop() → true면 그만, onProgress(done, total)
 *   png: false                       // true면 .png도 쓴다 (확인용)
 *   Resvg                            // @resvg/resvg-js의 Resvg (넘기지 않으면 불러온다)
 * }
 */
export async function renderThumbnails(thumbs, opts) {
  const { outDir, mode = 'full', batchBytes = 48 * 1024 * 1024, batchRows = 48, yieldMs = 0, shouldStop = () => false, onProgress = () => {} } = opts;
  const Resvg = opts.Resvg ?? (await import('@resvg/resvg-js')).Resvg;
  const count = thumbs.length;
  const geo = THUMB_SCALES.map(thumbGeometry);
  const big = geo.at(-1); // 2x: 600 × 56
  const find = (name) => thumbs.find((t) => t.name === name)?.path;
  const fb = {
    latin: opts.fallback?.latin ?? find('Arial') ?? find('Segoe UI') ?? find('Liberation Sans') ?? find('DejaVu Sans'),
    hangul: opts.fallback?.hangul ?? find('Malgun Gothic') ?? find('NanumGothic') ?? find('Noto Sans KR') ?? find('Gulim'),
    mono: opts.fallback?.mono ?? find('Courier New') ?? find('Liberation Mono') ?? find('DejaVu Sans Mono'),
  };
  const fbFiles = [...new Set([fb.latin, fb.hangul].filter(Boolean))];
  const sizeOf = (p) => { try { return statSync(p).size; } catch { return 0; } };
  const alpha = { false: geo.map((g) => new Uint8Array(g.width * g.row * count)), true: geo.map((g) => new Uint8Array(g.width * g.row * count)) };

  // 묶음 나누기: 한 번에 싣는 글꼴 파일 크기 합이 batchBytes를 넘지 않게
  const fileOf = (t) => (mode !== 'full' ? null : t.symbolic ? fb.mono : t.path);
  const batches = [];
  let cur = []; let curFiles = new Set(); let curBytes = 0;
  for (let i = 0; i < count; i++) {
    const file = fileOf(thumbs[i]);
    const add = file && !curFiles.has(file) ? sizeOf(file) : 0;
    if (cur.length && (cur.length >= batchRows || curBytes + add > batchBytes)) { batches.push({ rows: cur, files: [...curFiles] }); cur = []; curFiles = new Set(); curBytes = 0; }
    cur.push(i);
    if (file && !curFiles.has(file)) { curFiles.add(file); curBytes += add; }
  }
  if (cur.length) batches.push({ rows: cur, files: [...curFiles] });

  let done = 0;
  for (const b of batches) {
    if (shouldStop()) return null;
    // 칸: 글꼴마다 2x 한 줄 (+ _ea 이름이 따로 있으면 한 줄 더). plain은 영어 이름만
    const cells = [];
    for (const i of b.rows) {
      const t = thumbs[i];
      cells.push({ i, ea: false, text: t.name, t });
      if (mode === 'full' && t.eaName) cells.push({ i, ea: true, text: t.eaName, t });
    }
    const H = big.row * cells.length;
    const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${big.width}" height="${H}"><defs>`];
    cells.forEach((c, n) => parts.push(`<clipPath id="c${n}"><rect x="0" y="${n * big.row}" width="${big.width}" height="${big.row}"/></clipPath>`));
    parts.push('</defs>');
    cells.forEach((c, n) => {
      const t = c.t;
      let fam; let weight = 400; let style = 'normal';
      if (mode !== 'full') fam = cssList(FALLBACK_FAMILIES);
      else if (t.symbolic) fam = cssList(['Courier New', 'Liberation Mono', 'DejaVu Sans Mono', ...FALLBACK_FAMILIES]);
      else { fam = cssList([...(t.css ?? [t.name]), ...FALLBACK_FAMILIES]); weight = t.weight || 400; style = t.italic ? 'italic' : 'normal'; }
      parts.push(`<text clip-path="url(#c${n})" x="${big.x.toFixed(3)}" y="${(n * big.row + big.baseline).toFixed(3)}" font-size="${big.fontPx.toFixed(3)}" font-family="${fam}" font-weight="${weight}" font-style="${style}" fill="#000">${xml(c.text)}</text>`);
    });
    parts.push('</svg>');
    let pixels = null;
    try {
      // 대체 글꼴을 먼저 실어야 빠진 글자를 그것으로 채운다 (resvg는 실은 순서대로 찾는다)
      const files = mode === 'full' ? [...fbFiles, ...b.files.filter((f) => !fbFiles.includes(f))] : fbFiles;
      const r = new Resvg(parts.join(''), { font: { loadSystemFonts: false, fontFiles: files, defaultFontFamily: 'Arial', serifFamily: 'Arial', sansSerifFamily: 'Arial', monospaceFamily: 'Courier New' }, fitTo: { mode: 'original' }, logLevel: 'off' });
      pixels = r.render().pixels;
    } catch {
      pixels = null; // 이 묶음은 빈 칸으로 둔다
    }
    if (pixels) {
      const cell = new Uint8Array(big.width * big.row);
      cells.forEach((c, n) => {
        for (let p = 0, q = n * big.width * big.row * 4 + 3; p < cell.length; p++, q += 4) cell[p] = pixels[q];
        geo.forEach((g, k) => {
          const a = k === geo.length - 1 ? cell : downsampleAlpha(cell, big.width, big.row, g.width, g.row);
          const base = c.i * g.width * g.row;
          alpha[c.ea][k].set(a, base);
          if (!c.ea && !(mode === 'full' && c.t.eaName)) alpha.true[k].set(a, base); // _ea 이름이 따로 없으면 같은 그림
        });
      });
    }
    done += b.rows.length;
    onProgress(done, count);
    if (yieldMs) await new Promise((ok) => setTimeout(ok, yieldMs));
  }

  mkdirSync(outDir, { recursive: true });
  const written = [];
  for (const ea of [false, true]) {
    THUMB_SCALES.forEach((s, k) => {
      const g = geo[k];
      const name = thumbFileName(s, ea);
      writeAtomic(join(outDir, name), encodeSpriteBin(alpha[ea][k], g.width, g.row, count));
      written.push(name);
      if (opts.png) {
        const a = alpha[ea][k]; const rgba = Buffer.alloc(a.length * 4);
        for (let i = 0; i < a.length; i++) rgba[4 * i + 3] = a[i];
        writeAtomic(join(outDir, thumbFileName(s, ea, false)), encodePng(g.width, g.row * count, rgba));
      }
    });
  }
  return written;
}

/** 파일을 임시 이름으로 쓴 뒤 바꿔 넣기 (Windows에서 잠깐 잠긴 파일은 몇 번 다시) */
export function writeAtomic(path, data) {
  const tmp = `${path}.${process.pid}.${Math.random().toString(36).slice(2, 8)}.tmp`;
  writeFileSync(tmp, data);
  for (let n = 0; ; n++) {
    try { renameSync(tmp, path); return; } catch (err) {
      if (n >= 5 || !['EPERM', 'EBUSY', 'EACCES'].includes(err.code)) { try { rmSync(tmp, { force: true }); } catch {} throw err; }
      const until = Date.now() + 50 * (n + 1); while (Date.now() < until) { /* 잠깐 기다림 */ }
    }
  }
}

/** 빈 그림 10개를 폴더에 쓰기 (이미 있으면 두기) */
export function writeBlankSprites(outDir, count) {
  mkdirSync(outDir, { recursive: true });
  for (const [name, buf] of blankSprites(count)) if (!existsSync(join(outDir, name))) writeAtomic(join(outDir, name), buf);
}
