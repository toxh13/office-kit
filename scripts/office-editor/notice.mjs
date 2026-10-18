// notice.mjs — vendor/office-editor/NOTICE-OfficeKit.md (날짜 붙은 수정 고지, 설계 8장 2번).
// ONLYOFFICE 추가 조항 2: 수정판은 수정했다는 것과 날짜, ONLYOFFICE 원본을 바탕으로 했다는 것을 눈에 띄게 적는다.
import { EDITOR_LOCK, FILEX, ONLYOFFICE, UI_ZIP, X2T } from './pins.mjs';
import { NOTICE_DATED, RUNTIME_REPLACED } from './rules.mjs';

const OFFICE_KIT_REPO = 'https://github.com/toxh13/office-kit';
const mb = (n) => `${(n / 1048576).toFixed(1)}MB`;
const num = (n) => n.toLocaleString('en-US');
const cell = (s) => String(s).replace(/\|/g, '\\|').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * @param {object} o
 * @param {string} o.builtAt               ISO 시각
 * @param {string} o.officeKitVersion      package.json version
 * @param {{id,what,why,files,bytes}[]} o.removed
 * @param {{name,date,subject,sha256}[]} o.patches
 * @param {{changed:{path,sha256}[]}|null} o.rebuilt
 * @param {number} o.fileCount
 * @param {number} o.totalBytes
 */
export function renderNotice(o) {
  const day = o.builtAt.slice(0, 10);
  const srcZip = `office-editor-source-${o.officeKitVersion}.zip`;
  const removedFiles = o.removed.reduce((s, r) => s + r.files, 0);
  const removedBytes = o.removed.reduce((s, r) => s + r.bytes, 0);
  const lines = [
    '# Office Kit 수정 고지 — vendor/office-editor',
    '',
    `이 폴더는 Office Kit의 문서·표·슬라이드 편집 창이 쓰는 편집기 묶음이다. **${ONLYOFFICE.product} ${ONLYOFFICE.version}`,
    `(빌드 ${ONLYOFFICE.build}, ${ONLYOFFICE.vendor})** 을 바탕으로 한 **${FILEX.name} ${FILEX.version} (BRF Tech)** 의`,
    `릴리스 묶음(\`ui.zip\`)을 Office Kit이 고친 **수정판**이다. 원본과 다르게 동작할 수 있다.`,
    '',
    `- 수정한 곳: Office Kit (${OFFICE_KIT_REPO}) ${o.officeKitVersion}`,
    `- 수정 날짜: ${NOTICE_DATED} (덜어 내기 규칙)${o.patches.length ? `, filex 패치는 아래 표의 날짜` : ''}`,
    `- 이 묶음을 만든 날: ${day}`,
    `- 남은 파일: ${num(o.fileCount)}개, ${mb(o.totalBytes)} (이 고지·VERSION.json 빼고)`,
    '',
    '## 바탕이 된 것',
    '',
    '| 부분 | 판 | 라이선스 | 고지·라이선스 파일 |',
    '|---|---|---|---|',
    `| ONLYOFFICE Docs 편집기 (web-apps, sdkjs) | ${ONLYOFFICE.build} — Document Server 이미지 \`${ONLYOFFICE.image}\` | AGPL-3.0 + 추가 조항 (\`editor/LICENSE.txt\` 끝부분). 코드가 아닌 것(그림·아이콘·글)은 CC BY-SA 4.0 | \`editor/LICENSE.txt\`, \`editor/3rd-Party.txt\`, \`editor/license/\` |`,
    `| x2t (WebAssembly 변환기) | ONLYOFFICE core ${X2T.core.tag} + CryptPad·BRF Tech 패치 (filex 릴리스 \`${X2T.release}\`) | AGPL-3.0 (패치는 AGPL-3.0-or-later) | \`NOTICE\` |`,
    `| ${FILEX.name} (앱 페이지, Document Server 대신 답하는 다리, 변환 워커) | ${FILEX.version} — 커밋 \`${FILEX.commit}\` | AGPL-3.0-or-later (\`src/locks.ts\`만 AGPL-3.0-only) | \`LICENSE\`, \`NOTICE\`, \`editor/filex/CHANGES.txt\` |`,
    `| 편집기 안 제3자 JavaScript (Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO 등) | — | MIT·BSD·Apache-2.0 등 | \`editor/3rd-Party.txt\`, \`editor/license/\` |`,
    `| x2t·sdkjs WebAssembly 안 제3자 코드 (Boost, ICU, FreeType, HarfBuzz, libpng, libjpeg, libtiff, OpenJPEG, libxml2, Xpdf, DjVuLibre, emscripten 런타임 등)와 라이선스 파일 없는 es6-promise·fetch | — | BSL-1.0·ICU·FTL·MIT·BSD·Apache-2.0·IJG·GPL/LGPL 등 | \`THIRD-PARTY-NATIVE.txt\` (Office Kit이 더함) |`,
    `| x2t.wasm 안 PDF 기본 14 글꼴 (URW++ Nimbus Sans·Roman·Mono, Standard Symbols L, Dingbats — ONLYOFFICE core \`PdfFile/Resources/Fonts\`) | — | AGPL-3.0 + PDF·PostScript 포함 예외 | \`THIRD-PARTY-NATIVE.txt\` |`,
    '',
    `filex가 ONLYOFFICE 편집기 파일에 한 일(HTML 안 스크립트를 파일로 옮김, 이름 바꿈, 뺀 것, 다리로 바꾼`,
    `\`web-apps/vendor/socketio/socket.io.min.js\` 등)은 \`editor/filex/CHANGES.txt\`에 있다. 편집기 파일마다의 SHA-256은`,
    `filex \`${EDITOR_LOCK.path}\`(커밋 \`${FILEX.commit.slice(0, 12)}\`)에 있고, Office Kit은 묶음을 만들 때 파일마다 이것과 맞춰 본다.`,
    '',
    '## Office Kit이 바꾼 것',
    '',
    `### 1. 파일 빼기 (${NOTICE_DATED})`,
    '',
    '파일을 **통째로 뺐을 뿐**, 남은 파일은 한 바이트도 고치지 않았다 (줄이기·압축 다시 하기 없음, 저작권·라이선스 머리글 그대로).',
    '',
    '| 뺀 것 | 파일 | 크기 | 까닭 |',
    '|---|---:|---:|---|',
    ...o.removed.map((r) => `| ${cell(r.what)} | ${num(r.files)} | ${mb(r.bytes)} | ${cell(r.why)} |`),
    `| **모두** | **${num(removedFiles)}** | **${mb(removedBytes)}** | |`,
    '',
    '빼지 않은 것: `LICENSE`, `NOTICE`, `editor/LICENSE.txt`, `editor/3rd-Party.txt`, `editor/license/`, `editor/filex/CHANGES.txt`, 모든 파일 머리글. 더한 것: 이 고지, `VERSION.json`, `THIRD-PARTY-NATIVE.txt`(기계어·JS 제3자 고지).',
    '',
    '### 2. filex 패치',
    '',
  ];
  if (!o.patches.length) {
    lines.push(`없음 — ${FILEX.name} ${FILEX.version} 릴리스의 앱 페이지와 다리를 그대로 쓴다.`);
  } else {
    lines.push(
      `filex 소스(커밋 \`${FILEX.commit.slice(0, 12)}\`)에 아래 패치를 얹어 앱 페이지를 다시 빌드했다 (filex \`scripts/build-app.mjs\`).`,
      '패치는 AGPL-3.0-or-later이고 Office Kit 저장소 `patches/office-editor/filex/`에 있다.',
      '',
      '| 패치 | 날짜 | 내용 | SHA-256 (LF) |',
      '|---|---|---|---|',
      ...o.patches.map((p) => `| \`${cell(p.name)}\` | ${p.date || NOTICE_DATED} | ${cell(p.subject || '—')} | \`${p.sha256.slice(0, 16)}…\` |`),
      '',
    );
    if (o.rebuilt?.changed?.length) {
      lines.push('다시 빌드해서 릴리스와 달라진 파일:', '', ...o.rebuilt.changed.map((c) => `- \`${c.path}\``), '');
    }
  }
  lines.push(
    '',
    '### 3. 실행 중에 바꿔 내주는 것 (파일은 그대로)',
    '',
    'Office Kit은 이 묶음을 이 PC 안에서만(127.0.0.x) 내준다. 아래 주소는 묶음의 파일 대신 그 자리에서 만든 것을 내준다.',
    '',
    ...RUNTIME_REPLACED.map((r) => `- \`${r.path}\`: ${r.what}`),
    '',
    '## 상표',
    '',
    `ONLYOFFICE는 ${ONLYOFFICE.vendor}의 상표다. Office Kit은 ${ONLYOFFICE.vendor}·BRF Tech와 관계가 없고 그들이 보증하거나 지원하지 않는다.`,
    '편집기 안의 ONLYOFFICE 로고와 \'정보\' 화면은 라이선스 추가 조항에 따라 그대로 두었다. Office Kit의 이름·창 제목·기능 이름에는 ONLYOFFICE를 쓰지 않는다.',
    '',
    '## 소스 받기',
    '',
    `- 이 묶음을 다시 만드는 방법: Office Kit 저장소의 \`scripts/build-office-editor.mjs\`·\`scripts/office-editor/\` (\`npm run office-editor:build\`)`,
    `- Office Kit 릴리스마다 붙는 소스 꾸러미 \`${srcZip}\`: ${OFFICE_KIT_REPO}/releases/tag/v${o.officeKitVersion}`,
    `- ${FILEX.name}: ${FILEX.repo}/tree/${FILEX.commit} (릴리스 묶음 ${UI_ZIP.url}, SHA-256 \`${UI_ZIP.sha256}\`)`,
    ...Object.entries(ONLYOFFICE.sources).map(([k, v]) => `- ONLYOFFICE ${k}: ${v}`),
    `- x2t 빌드 방법: ${X2T.recipe} (core \`${X2T.core.commit}\`, build_tools \`${X2T.buildTools.commit}\`)`,
    '',
    '## Modification notice (English)',
    '',
    `This directory is a **modified version** of the ${FILEX.name} ${FILEX.version} release bundle (ui.zip, BRF Tech),`,
    `which is based on **${ONLYOFFICE.product} ${ONLYOFFICE.version} (build ${ONLYOFFICE.build}) by ${ONLYOFFICE.vendor}**.`,
    `It was modified by Office Kit (${OFFICE_KIT_REPO}) on ${NOTICE_DATED}${o.patches.length ? ' and on the patch dates listed above' : ''}; this copy was built on ${day}.`,
    `Changes: whole files were removed (${num(removedFiles)} files: editor languages other than Korean and English,`,
    'formula languages other than Korean and English, the mobile/embed/forms apps, and the editor/fonts folder with its licenses);',
    'no remaining file was altered, and all copyright and license notices are kept.',
    o.patches.length
      ? `The filex app page was rebuilt from its source with the patches listed above (${o.patches.map((p) => p.name).join(', ')}).`
      : 'The filex app page and bridge are used as released.',
    'At run time, Office Kit serves a font list, font thumbnails and font files generated from the fonts installed on the computer',
    'in place of AllFonts.js, fonts_thumbnail*.png and editor/fonts/*.',
    `ONLYOFFICE is a trademark of ${ONLYOFFICE.vendor}; Office Kit is not affiliated with or endorsed by ${ONLYOFFICE.vendor} or BRF Tech.`,
    'x2t/x2t.wasm still contains the 14 URW++ base-14 Type 1 fonts that ONLYOFFICE core embeds for PDF output (AGPL-3.0 with a font embedding exception).',
    'Office Kit added THIRD-PARTY-NATIVE.txt: notices for the libraries compiled into the WebAssembly files and for es6-promise and fetch.',
    `License texts: editor/LICENSE.txt (AGPL-3.0 with additional terms), LICENSE and NOTICE (filex), editor/3rd-Party.txt, editor/license/, THIRD-PARTY-NATIVE.txt.`,
    `Corresponding source: ${srcZip} attached to ${OFFICE_KIT_REPO}/releases/tag/v${o.officeKitVersion}, and the upstream links above.`,
    '',
  );
  return lines.join('\n');
}
