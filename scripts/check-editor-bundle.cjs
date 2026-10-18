// check-editor-bundle.cjs — electron-builder beforePack: 한글 편집기 묶음(vendor/)이 없으면 설치 파일을 만들지 않는다.
//   electron-builder를 바로 실행해도(npm run dist를 거치지 않아도) 편집기 없는 설치 파일이 나가지 않게.
//   설치 파일에는 들어가지 않는다 (build.files에 scripts/com만 있음).
const { existsSync, readdirSync, statSync } = require('node:fs');
const { join, resolve } = require('node:path');

const isFile = (p) => { try { return statSync(p).isFile(); } catch { return false; } };

/** 빠진 것 목록 (없으면 빈 배열) */
function missingParts(root) {
  const studio = join(root, 'vendor', 'rhwp-studio');
  const sdk = join(root, 'vendor', 'rhwp-editor-sdk');
  const out = [];
  if (!isFile(join(studio, 'index.html'))) out.push('vendor/rhwp-studio/index.html');
  let wasm = false;
  try { wasm = readdirSync(join(studio, 'assets')).some((f) => /^rhwp_bg.*\.wasm$/i.test(f)); } catch {}
  if (!wasm) out.push('vendor/rhwp-studio/assets/rhwp_bg*.wasm');
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) {
    if (!isFile(join(sdk, f))) out.push(`vendor/rhwp-editor-sdk/${f}`);
  }
  return out;
}

async function checkEditorBundle(context) {
  const root = resolve(context?.packager?.projectDir ?? context?.projectDir ?? join(__dirname, '..'));
  const missing = missingParts(root);
  if (missing.length) {
    throw new Error(`vendor/ 한글 편집기 묶음이 없습니다 — 먼저 npm run editor:build (빠진 것: ${missing.join(', ')})`);
  }
  if (!existsSync(join(root, 'vendor', 'rhwp-studio', 'VERSION.json'))) {
    console.warn('[editor] vendor/rhwp-studio/VERSION.json이 없습니다 — npm run editor:build로 만든 묶음인지 확인하세요');
  }
}

module.exports = checkEditorBundle;
module.exports.default = checkEditorBundle;
module.exports.missingParts = missingParts;
