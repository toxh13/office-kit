// check-editor-bundle.cjs — electron-builder beforePack: 편집기 묶음(vendor/)이 없으면 설치 파일을 만들지 않는다.
//   electron-builder를 바로 실행해도(npm run dist를 거치지 않아도) 편집기 없는 설치 파일이 나가지 않게.
//   - 한글 편집기: vendor/rhwp-studio, vendor/rhwp-editor-sdk (npm run editor:build)
//   - Office 편집기: vendor/office-editor (npm run office-editor:build) — 라이선스 파일·수정 고지·VERSION.json까지 (배포 관문)
//   설치 파일에는 들어가지 않는다 (build.files에 scripts/com만 있음).
const { existsSync, readdirSync, readFileSync, statSync } = require('node:fs');
const { join, resolve } = require('node:path');

const isFile = (p) => { try { return statSync(p).isFile(); } catch { return false; } };

/** 한글 편집기에서 빠진 것 목록 (없으면 빈 배열) */
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

// Office 편집기 묶음에 꼭 있어야 하는 파일 (docs/OFFICE-EDITOR-DESIGN.md 4장 CI·배포, 8장 9번)
const OFFICE_FILES = [
  'index.html',
  'x2t/x2t.wasm',
  'x2t/x2t.js',
  'filex/app.js',
  'editor/web-apps/apps/api/documents/api.js',
  'editor/LICENSE.txt',
  'editor/3rd-Party.txt',
  'editor/filex/CHANGES.txt',
  'LICENSE',
  'NOTICE',
  'NOTICE-OfficeKit.md',
  'THIRD-PARTY-NATIVE.txt',
  'VERSION.json',
];

/** Office 편집기에서 빠지거나 깨진 것 목록 (없으면 빈 배열) */
function missingOfficeParts(root) {
  const dir = join(root, 'vendor', 'office-editor');
  const out = [];
  for (const f of OFFICE_FILES) if (!isFile(join(dir, f))) out.push(`vendor/office-editor/${f}`);
  let licenses = false;
  try { licenses = readdirSync(join(dir, 'editor', 'license')).length > 0; } catch {}
  if (!licenses) out.push('vendor/office-editor/editor/license/');
  if (out.length) return out;
  // VERSION.json에 적힌 파일이 모두 그 크기로 있는지 (반쯤 복사된 묶음을 막음 — 해시는 --verify가 봄)
  let v;
  try { v = JSON.parse(readFileSync(join(dir, 'VERSION.json'), 'utf8')); } catch { return ['vendor/office-editor/VERSION.json (읽을 수 없음)']; }
  if (v?.name !== 'office-editor' || !Array.isArray(v.files) || !v.files.length) return ['vendor/office-editor/VERSION.json (모양이 다름)'];
  // 고지·편집기 아래 법적 고지 줄의 소스 주소는 package.json 판을 따른다 → 판을 올린 뒤 묶음을 다시 만들지 않았으면 막는다
  let pkgVersion = null;
  try { pkgVersion = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version; } catch {}
  if (pkgVersion && v.officeKitVersion !== pkgVersion) {
    return [`vendor/office-editor/VERSION.json (officeKitVersion ${v.officeKitVersion} ≠ package.json ${pkgVersion} — npm run office-editor:build로 다시 만드세요)`];
  }
  for (const f of v.files) {
    let size = -1;
    try { size = statSync(join(dir, ...String(f.path).split('/'))).size; } catch {}
    if (size !== f.bytes) out.push(`vendor/office-editor/${f.path}${size < 0 ? '' : ' (크기가 다름)'}`);
    if (out.length >= 10) break;
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
  const office = missingOfficeParts(root);
  if (office.length) {
    throw new Error(`vendor/office-editor Office 편집기 묶음이 없거나 깨졌습니다 — 먼저 npm run office-editor:build (깨졌으면 -- --force) (빠진 것: ${office.join(', ')})`);
  }
}

module.exports = checkEditorBundle;
module.exports.default = checkEditorBundle;
module.exports.missingParts = missingParts;
module.exports.missingOfficeParts = missingOfficeParts;
module.exports.OFFICE_FILES = OFFICE_FILES;
