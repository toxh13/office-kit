// third-party.mjs — 편집기 묶음 안 기계어(WebAssembly)·JS 라이브러리의 제3자 고지 (LICENSES/office-editor-third-party-native.txt).
//
// ONLYOFFICE가 내는 고지(editor/3rd-Party.txt, editor/license/)는 편집기 안 JavaScript만 다룬다. 그런데
//   x2t/x2t.wasm  (ONLYOFFICE core v9.4.0.129를 emscripten으로 빌드 — filex scripts/x2t/steps.sh)
//   editor/sdkjs/common/*/engine/*.wasm  (sdkjs 글꼴·zlib 엔진: FreeType·HarfBuzz·libpng·libjpeg·libtiff·CxImage·Brotli·zlib)
//   editor/web-apps/vendor/es6-promise, editor/web-apps/vendor/fetch  (라이선스 파일이 없는 MIT 라이브러리)
// 안에 든 라이브러리의 라이선스 상당수(BSD·FTL·IJG·libtiff·MIT·Apache)는 이진 배포에도 고지를 함께 내라고 한다.
// 그래서 그 고지를 고정한 주소(태그·커밋)에서 받아 SHA-256을 맞춘 뒤 한 파일로 엮어 저장소에 둔다 (빌드는 받지 않음 — 오프라인).
//
//   node scripts/office-editor/third-party.mjs --write   주소에서 받아 SHA-256 확인 → LICENSES/office-editor-third-party-native.txt
//   node scripts/office-editor/third-party.mjs --check   저장소의 파일이 지금 목록으로 만든 것과 같은지 (받지 않음: 파일 안 원문으로 다시 엮음)
//
// 목록을 바꾸면 --write로 다시 만들고, 묶음을 다시 빌드한다 (build-office-editor가 이 파일을 묶음에 THIRD-PARTY-NATIVE.txt로 넣는다).
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
/** 저장소 안 고지 파일 · 묶음 안 이름 */
export const THIRD_PARTY_FILE = join(ROOT, 'LICENSES', 'office-editor-third-party-native.txt');
export const THIRD_PARTY_BUNDLE_PATH = 'THIRD-PARTY-NATIVE.txt';

const CORE = 'https://raw.githubusercontent.com/ONLYOFFICE/core/a016fc2868b9f99f59dd30aa2088e68bbe5104af'; // v9.4.0.129
const GH = 'https://raw.githubusercontent.com';
const EMSDK = `${GH}/emscripten-core/emscripten/4.0.11`;

const X2T = 'x2t/x2t.wasm';
const SDKJS = 'editor/sdkjs/common/libfont/engine/fonts.wasm, editor/sdkjs/common/zlib/engine/zlib.wasm';

/**
 * 고지 목록. from/to: 원문에서 이 줄(정규식)부터 · 이 줄 앞까지만 (긴 README·소스 머리에서 라이선스 부분만)
 * @type {readonly {id:string, name:string, version:string, where:string, license:string, url:string, sha256:string, from?:RegExp, to?:RegExp, note?:string}[]}
 */
export const THIRD_PARTY = Object.freeze([
  // ── x2t: core 밖에서 받아 빌드하는 것 (core Common/3dParty/*/fetch.py·make.py, filex upstream/onlyoffice.json) ──
  { id: 'boost', name: 'Boost (headers, date_time, regex)', version: '1.84.0', where: X2T, license: 'BSL-1.0', url: `${GH}/boostorg/boost/boost-1.84.0/LICENSE_1_0.txt`, sha256: 'c9bff75738922193e67fa726fa225535870d2aa1059f91452c411736284ad566' },
  { id: 'icu', name: 'ICU (emscripten port)', version: 'release-68-2', where: X2T, license: 'ICU (Unicode) and third-party terms listed inside', url: `${GH}/unicode-org/icu/release-68-2/icu4c/LICENSE`, sha256: '7915b19db903070778581ae05d8bf4ea241b34a05deb51ca4f5cbb15ea1cbba3' },
  { id: 'harfbuzz', name: 'HarfBuzz', version: '894a1f72ee93', where: `${X2T}, ${SDKJS}`, license: 'MIT (Old MIT)', url: `${GH}/harfbuzz/harfbuzz/894a1f72ee93a1fd8dc1d9218cb3fd8f048be29a/COPYING`, sha256: 'ba8f810f2455c2f08e2d56bb49b72f37fcf68f1f4fade38977cfd7372050ad64' },
  { id: 'brotli', name: 'Brotli', version: 'a47d7475063e', where: `${X2T}, ${SDKJS}`, license: 'MIT', url: `${GH}/google/brotli/a47d7475063eb223c87632eed806c0070e70da29/LICENSE`, sha256: '3d180008e36922a4e8daec11c34c7af264fed5962d07924aea928c38e8663c94' },
  { id: 'hyphen', name: 'Hyphen', version: 'dcd90f0056cf', where: X2T, license: 'MPL-1.1 / GPL-2.0-or-later / LGPL-2.1-or-later (tri-license)', url: `${GH}/hunspell/hyphen/dcd90f0056cfbcbd2d389b7a13db45ddf5b0b98e/COPYING`, sha256: '706294d8b12dd1b11d229f0d8e2a869e20f1622c8bacd18efb62f999e56e65e9' },
  { id: 'hyphen-mpl', name: 'Hyphen — MPL 1.1 text', version: 'dcd90f0056cf', where: X2T, license: 'MPL-1.1', url: `${GH}/hunspell/hyphen/dcd90f0056cfbcbd2d389b7a13db45ddf5b0b98e/COPYING.MPL`, sha256: '53692a2ed6c6a2c6ec9b32dd0b820dfae91e0a1fcdf625ca9ed0bdf8705fcc4f' },
  { id: 'gumbo', name: 'Gumbo HTML parser', version: 'aa91b27b02c0', where: X2T, license: 'Apache-2.0', url: `${GH}/google/gumbo-parser/aa91b27b02c0c80c482e24348a457ed7c3c088e0/COPYING`, sha256: 'c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4', note: 'Gumbo has no NOTICE file. ONLYOFFICE changes src/tag.c (isspace cast).' },
  { id: 'katana', name: 'Katana CSS parser', version: 'be6df458d454', where: X2T, license: 'MIT', url: `${GH}/jasenhuang/katana-parser/be6df458d4540eee375c513958dcb862a391cdd1/LICENSE`, sha256: 'ec400e80cb6f41880c97723d722cfcba44273918f488db567f94c1262776beb7' },
  { id: 'md4c', name: 'MD4C', version: '481fbfbdf72d', where: X2T, license: 'MIT', url: `${GH}/mity/md4c/481fbfbdf72daab2912380d62bb5f2187d438408/LICENSE.md`, sha256: 'd30937367d5413e7eaa218b1640b8946ff76fd34d97152f6979fd96169d5d0fc' },
  { id: 'glm', name: 'GLM', version: '33b4a621a697', where: X2T, license: 'MIT (or The Happy Bunny License)', url: `${GH}/g-truc/glm/33b4a621a697a305bc3a7610d290677b96beb181/copying.txt`, sha256: '62d2d642c7d054d4fb4c9b42faad617d6c88fcd91e317f8035aa9f277cc159c3' },
  { id: 'mdds', name: 'mdds', version: '0783158939c6', where: X2T, license: 'MIT', url: `${GH}/kohei-us/mdds/0783158939c6ce4b0b1b89e345ab983ccb0f0ad0/LICENSE`, sha256: '559cdfad8d02644df807f40fa43e9895c021d516f2e1c87bda406dbd1c69e4b5' },
  { id: 'librevenge', name: 'librevenge', version: 'becd044b519a', where: X2T, license: 'MPL-2.0 (or LGPL-2.1-or-later)', url: `${GH}/Distrotech/librevenge/becd044b519ab83893ad6398e3cbb499a7f0aaf4/COPYING.MPL`, sha256: 'fab3dd6bdab226f1c08630b1dd917e11fcb4ec5e1e020e2c16f83a0a13863e85', note: 'The same MPL-2.0 text also covers libodfgen (8ef8c171ebe3) and libetonyek (cb396b4a9453); their source is at https://github.com/Distrotech/libodfgen and https://github.com/LibreOffice/libetonyek at those commits.' },
  // ── x2t: ONLYOFFICE core 안에 들어 있는 것 (core 저장소의 그 폴더 라이선스 파일) ──
  { id: 'freetype', name: 'FreeType', version: '2.10.4 (core DesktopEditor/freetype-2.10.4)', where: `${X2T}, ${SDKJS}`, license: 'FTL (FreeType License; chosen over GPL-2.0)', url: `${GH}/freetype/freetype/VER-2-10-4/docs/FTL.TXT`, sha256: '08c135755dd589039470f1fdbb400daaabaaa50d0b366d19cebff4d22986baa1', note: 'Portions of this software are copyright (C) 2020 The FreeType Project (www.freetype.org). All rights reserved.' },
  { id: 'cximage', name: 'CxImage', version: 'core DesktopEditor/cximage', where: `${X2T}, ${SDKJS}`, license: 'CxImage license (zlib-like)', url: `${CORE}/DesktopEditor/cximage/CxImage/license.txt`, sha256: 'f50c9c12604dafcd15bc88b2e76fbf2e4009c18702669b463e839eb985897034' },
  { id: 'libpng', name: 'libpng', version: 'core DesktopEditor/cximage/png', where: `${X2T}, ${SDKJS}`, license: 'libpng License', url: `${CORE}/DesktopEditor/cximage/png/LICENSE`, sha256: '693fd4bca4178b4b4f11c8043ceb9b68b76658465e692d18e3e8b4920d2fbe66' },
  { id: 'libjpeg', name: 'IJG libjpeg', version: 'core DesktopEditor/cximage/jpeg', where: `${X2T}, ${SDKJS}`, license: 'IJG License', url: `${CORE}/DesktopEditor/cximage/jpeg/README`, sha256: '5fa97edc0ce7dfdcfb3744d705ebd2665190a7f77388747ffaa75c16f10fcfb2', from: /^LEGAL ISSUES\s*$/, to: /^REFERENCES\s*$/, note: 'This software is based in part on the work of the Independent JPEG Group.' },
  { id: 'libtiff', name: 'LibTIFF', version: 'core DesktopEditor/cximage/tiff', where: `${X2T}, ${SDKJS}`, license: 'libtiff License', url: `${CORE}/DesktopEditor/cximage/tiff/tiff.h`, sha256: 'b54d7ac2bb124b3df10aa7a50001e2c6369998621481dd32e94526bb019d2dbd', from: /^\/\*\s*$/, to: /^#ifndef _TIFF_/ },
  { id: 'jasper', name: 'JasPer', version: 'core DesktopEditor/cximage/jasper', where: X2T, license: 'JasPer License Version 2.0', url: `${CORE}/DesktopEditor/cximage/jasper/LICENSE`, sha256: 'edf9f4257cc33a1f396f9b062ca4a01dca7c79747c7d85b8947e603e5166760e' },
  { id: 'jbig', name: 'JBIG-KIT', version: 'core DesktopEditor/cximage/jbig', where: X2T, license: 'GPL-2.0-or-later', url: `${CORE}/DesktopEditor/cximage/jbig/copying.txt`, sha256: '91df39d1816bfb17a4dda2d3d2c83b1f6f2d38d53e53e41e8f97ad5ac46a0cad' },
  { id: 'libpsd', name: 'libpsd', version: 'core DesktopEditor/cximage/libpsd', where: X2T, license: 'LGPL-2.1-or-later', url: `${CORE}/DesktopEditor/cximage/libpsd/COPYING`, sha256: 'd245807f90032872d1438d741ed21e2490e1175dc8aa3afa5ddb6c8e529b58e5' },
  { id: 'libmng', name: 'libmng', version: 'core DesktopEditor/cximage/mng', where: X2T, license: 'libmng License (zlib-like)', url: `${CORE}/DesktopEditor/cximage/mng/LICENSE.txt`, sha256: 'd2ab2151d732a827473845c82df86a90cbe62860701d18cce0750f078b10acc1' },
  { id: 'openjpeg', name: 'OpenJPEG', version: '2.4.0 (core DesktopEditor/raster/Jp2/openjpeg)', where: X2T, license: 'BSD-2-Clause', url: `${CORE}/DesktopEditor/raster/Jp2/openjpeg/openjpeg-2.4.0/LICENSE`, sha256: 'a6af136f3e15038a666b61f376612a07d9a4e48cb7c01adbf3e33b3f14ab49b6' },
  { id: 'zlib', name: 'zlib', version: '1.2.11 (core OfficeUtils/src/zlib-1.2.11)', where: `${X2T}, ${SDKJS}`, license: 'Zlib', url: `${CORE}/OfficeUtils/src/zlib-1.2.11/README`, sha256: '7960b6b1cc63e619abb77acaea5427159605afee8c8b362664f4effc7d7f7d15', from: /^Copyright notice:/ },
  { id: 'libxml2', name: 'libxml2', version: 'core DesktopEditor/xml/libxml2', where: X2T, license: 'MIT', url: `${CORE}/DesktopEditor/xml/libxml2/Copyright`, sha256: '7b5c35a856959625fca9bcedb4f13258a37910a0822aedf29166d3d5ea786951' },
  { id: 'agg', name: 'Anti-Grain Geometry', version: '2.4 (core DesktopEditor/agg-2.4)', where: X2T, license: 'Modified BSD (AGG 2.4)', url: `${CORE}/DesktopEditor/agg-2.4/copying`, sha256: 'f580c4e73d78e6f4c1c17ab9042bf630b0b4dabf3680d0c850cf14f910fdf616' },
  { id: 'cryptopp', name: 'Crypto++', version: 'core Common/3dParty/cryptopp', where: X2T, license: 'BSL-1.0 (library) / public domain (compilation)', url: `${CORE}/Common/3dParty/cryptopp/License.txt`, sha256: '78e4010b682cb94187fe0b57e50116d0ba271ef81104d1ddb35c80c3d81e3169' },
  { id: 'xpdf', name: 'Xpdf (PDF reader core: xpdf, fofi, goo, splash)', version: '4.03 (core PdfFile/lib)', where: X2T, license: 'GPL-2.0 or GPL-3.0 (Xpdf 4.x; used here under GPL-3.0, compatible with AGPL-3.0)', url: `${CORE}/PdfFile/lib/xpdf/XRef.h`, sha256: '9415df2cffee2e44cf2ee31ee94d1f65baa97375f7513d94913d578ac249ca3e', to: /^#ifndef XREF_H/, note: 'Copyright 1996-2021 Glyph & Cog, LLC. Xpdf 4.03 is distributed under the GNU GPL version 2 or version 3; ONLYOFFICE modified it. GPL-3.0 text: https://www.gnu.org/licenses/gpl-3.0.txt; GPL-2.0 text: the JBIG-KIT section above.' },
  { id: 'djvulibre', name: 'DjVuLibre (libdjvu)', version: '3.5 (core DjVuFile/libdjvu)', where: X2T, license: 'GPL-2.0-or-later', url: `${CORE}/DjVuFile/libdjvu/DjVuDocument.h`, sha256: '95fc15a4e1be58212175b693aa027d64f016a9a806dbe7d81b5d40772e54a39b', to: /^(?!\/\/C-)/, note: 'GPL-2.0 text: see the JBIG-KIT section above (copying.txt is the GNU GPL version 2).' },
  { id: 'urw-fonts', name: 'URW++ base-14 Type 1 fonts (Nimbus Sans, Nimbus Roman, Nimbus Mono, Standard Symbols L, Dingbats)', version: 'core PdfFile/Resources/Fonts', where: X2T, license: 'AGPL-3.0 with a font embedding exception', url: `${CORE}/PdfFile/Resources/Fonts/LICENSE`, sha256: '772b1b47eead8722156abaed483154e1dd1137e7035290550cea6444192a2e3f', note: 'x2t.wasm embeds these 14 fonts (Copyright URW++, 1997-2006) for PDF output. The AGPL-3.0 text (the COPYING file next to this LICENSE) is byte-for-byte LICENSES/AGPL-3.0.txt.' },
  { id: 'adobe-cmap', name: 'Adobe CMap resources', version: 'core PdfFile/Resources/CMap', where: X2T, license: 'BSD-3-Clause', url: `${CORE}/PdfFile/Resources/CMap/CMap/LICENSE.md`, sha256: 'feb8b068a417681821823ef2e8995072aa111121fa11f8664038f78717f4497c' },
  { id: 'openssl', name: 'OpenSSL (headers only — no OpenSSL code is linked)', version: 'OpenSSL_1_1_1f', where: X2T, license: 'OpenSSL / SSLeay', url: `${GH}/openssl/openssl/OpenSSL_1_1_1f/LICENSE`, sha256: 'c32913b33252e71190af2066f08115c69bc9fddadf3bf29296e20c835389841c' },
  // ── emscripten 4.0.11 런타임 (x2t·sdkjs wasm 모두) ──
  { id: 'emscripten', name: 'Emscripten runtime', version: '4.0.11', where: `${X2T} and x2t/x2t.js`, license: 'MIT / University of Illinois/NCSA (dual)', url: `${EMSDK}/LICENSE`, sha256: '620a78084fc7ca97c0b5dea9abf891f3ffcadfdbf305276f099c9c4e12fc1d86' },
  { id: 'musl', name: 'musl libc (emscripten system/lib/libc/musl)', version: 'emscripten 4.0.11', where: X2T, license: 'MIT', url: `${EMSDK}/system/lib/libc/musl/COPYRIGHT`, sha256: 'f9bc4423732350eb0b3f7ed7e91d530298476f8fec0c6c427a1c04ade22655af' },
  { id: 'libcxx', name: 'libc++ / libc++abi (emscripten system/lib)', version: 'emscripten 4.0.11', where: X2T, license: 'Apache-2.0 WITH LLVM-exception', url: `${EMSDK}/system/lib/libcxx/LICENSE.TXT`, sha256: '539dd7aed86e8a4f12cbdd0e6c50c189c7d74847e4fecc64ce2c6ee3a01da38b' },
  // ── web-apps의 라이선스 파일 없는 JS ──
  { id: 'es6-promise', name: 'es6-promise', version: '4.2.8', where: 'editor/web-apps/vendor/es6-promise/es6-promise.auto.min.js', license: 'MIT', url: `${GH}/stefanpenner/es6-promise/v4.2.8/LICENSE`, sha256: 'da852e222c72bc04a0a9046aac35f2e5d33775a8ea9f1d8f59b1cfd195ef9bb0' },
  { id: 'whatwg-fetch', name: 'whatwg-fetch (GitHub fetch polyfill)', version: '3.x', where: 'editor/web-apps/vendor/fetch/fetch.umd.js', license: 'MIT', url: `${GH}/github/fetch/v3.0.0/LICENSE`, sha256: '05a4adff1d99954f18ed2ef89d47b4567cdba599635ddbe9393a95b9d48b2643' },
]);

const HEAD = [
  'Third-party notices for native and unlicensed code in the Office Kit office editor bundle',
  '========================================================================================',
  '',
  'Office Kit ships the ONLYOFFICE Docs 9.4.0.129 editor (via filex-office-editor 0.3.1) in resources/office-editor/.',
  'The notices ONLYOFFICE ships (editor/3rd-Party.txt, editor/license/) cover the editor\'s JavaScript libraries and',
  'Hunspell. This file adds the notices for code compiled into the WebAssembly binaries and two JavaScript files',
  'that have no license file:',
  '',
  '  x2t/x2t.wasm, x2t/x2t.js   ONLYOFFICE core v9.4.0.129 (commit a016fc2868b9), built with emscripten 4.0.11 by',
  '                             filex-office-editor scripts/x2t (filex release v9.4.0.129+2). The list follows the',
  '                             core qmake projects that build links (graphics, PdfFile, DjVuFile, HtmlFile2, IWorkFile,',
  '                             CryptoPPLib, ...) and core\'s 3DPARTY.md; a library listed here may be only partly linked.',
  '  editor/sdkjs/common/*/engine/*.wasm   sdkjs font and zlib engines (FreeType, HarfBuzz, Brotli, CxImage, libpng,',
  '                             libjpeg, libtiff, zlib), built by ONLYOFFICE.',
  '  editor/web-apps/vendor/es6-promise, editor/web-apps/vendor/fetch.',
  '',
  'ONLYOFFICE core itself is AGPL-3.0 with additional terms (editor/LICENSE.txt). Source for every item: the',
  'pinned URL given in its section (the license text below was fetched from it and checked with the SHA-256 shown),',
  'and the corresponding source package office-editor-source-<version>.zip attached to each Office Kit release.',
  '',
  'Generated by scripts/office-editor/third-party.mjs — do not edit by hand.',
  '',
];

const RULE = '-'.repeat(100);

/** 원문 → 이 파일에 넣을 부분 (from 줄부터, to 줄 앞까지, 줄바꿈 LF, BOM 뺌) */
export function excerpt(entry, text) {
  let lines = text.replace(/^﻿/, '').replace(/\r\n?/g, '\n').split('\n');
  if (entry.from) {
    const i = lines.findIndex((l) => entry.from.test(l));
    if (i < 0) throw new Error(`${entry.id}: 시작 줄을 찾지 못함 (${entry.from})`);
    lines = lines.slice(i);
  }
  if (entry.to) {
    const j = lines.findIndex((l, k) => k > 0 && entry.to.test(l));
    if (j < 0) throw new Error(`${entry.id}: 끝 줄을 찾지 못함 (${entry.to})`);
    lines = lines.slice(0, j);
  }
  return lines.join('\n').replace(/\s+$/, '');
}

/** texts: id → 원문 (받은 그대로) */
export function renderThirdParty(texts) {
  const out = [...HEAD, 'Contents', ''];
  THIRD_PARTY.forEach((e, i) => out.push(`  ${String(i + 1).padStart(2)}. ${e.name} — ${e.license}`));
  out.push('');
  THIRD_PARTY.forEach((e, i) => {
    const t = texts[e.id];
    if (typeof t !== 'string') throw new Error(`${e.id}: 원문이 없음`);
    out.push(RULE, `${i + 1}. ${e.name}`, RULE,
      `Version:  ${e.version}`, `Used in:  ${e.where}`, `License:  ${e.license}`, `Source:   ${e.url}`, `SHA-256:  ${e.sha256} (of the file at Source)`);
    if (e.note) out.push(`Note:     ${e.note}`);
    out.push('', excerpt(e, t), '');
  });
  return `${out.join('\n')}\n`;
}

const sha = (b) => createHash('sha256').update(b).digest('hex');

async function fetchAll() {
  const texts = {};
  for (const e of THIRD_PARTY) {
    const r = await fetch(e.url);
    if (!r.ok) throw new Error(`${e.id}: ${e.url} → HTTP ${r.status}`);
    const b = Buffer.from(await r.arrayBuffer());
    const got = sha(b);
    if (got !== e.sha256) throw new Error(`${e.id}: SHA-256이 다름 (${got}) — ${e.url}`);
    texts[e.id] = b.toString('utf8');
  }
  return texts;
}

/** 저장소 파일에서 각 항목의 머리(판·주소·SHA-256)가 지금 목록과 맞는지 → 틀린 것 목록 */
export function checkThirdPartyFile(text = readFileSync(THIRD_PARTY_FILE, 'utf8')) {
  const bad = [];
  for (const e of THIRD_PARTY) {
    if (!text.includes(`Source:   ${e.url}\nSHA-256:  ${e.sha256} `)) bad.push(e.id);
  }
  return bad;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const arg = process.argv[2];
  if (arg === '--write') {
    const text = renderThirdParty(await fetchAll());
    writeFileSync(THIRD_PARTY_FILE, text);
    console.log(`${THIRD_PARTY_FILE}: ${THIRD_PARTY.length}개, ${text.length}바이트`);
  } else if (arg === '--check') {
    const bad = checkThirdPartyFile();
    if (bad.length) { console.error(`목록과 다름: ${bad.join(', ')} — node scripts/office-editor/third-party.mjs --write`); process.exit(1); }
    console.log('맞음');
  } else {
    console.log('node scripts/office-editor/third-party.mjs --write | --check');
  }
}
