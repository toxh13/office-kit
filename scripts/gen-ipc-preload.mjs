// gen-ipc-preload.mjs — 각 preload 파일의 채널 블록을 src/app/ipc-channels.mjs 표에서 만들어 바꿔 쓴다.
//   node scripts/gen-ipc-preload.mjs          블록을 다시 쓴다 (파일의 줄 끝을 따른다)
//   node scripts/gen-ipc-preload.mjs --check  쓰지 않고, 표와 다른 파일 이름을 찍고 exit 1
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { BLOCK_END, BLOCK_START, PRELOADS, preloadBlock } from '../src/app/ipc-channels.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const check = process.argv.includes('--check');

const stale = [];
for (const [rel, kinds] of Object.entries(PRELOADS)) {
  const file = join(ROOT, rel);
  const text = readFileSync(file, 'utf8');
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const a = text.indexOf(BLOCK_START);
  const b = text.indexOf(BLOCK_END);
  if (a < 0 || b < a) {
    console.error(`${rel}: 자동 생성 블록 표시(${BLOCK_START})가 없다`);
    process.exit(1);
  }
  const next = text.slice(0, a) + preloadBlock(kinds, eol) + text.slice(b + BLOCK_END.length);
  if (next === text) continue;
  stale.push(rel);
  if (!check) writeFileSync(file, next);
}

if (check && stale.length) {
  console.error('표와 다른 preload:\n' + stale.map((r) => '  ' + r).join('\n') + '\nnode scripts/gen-ipc-preload.mjs 로 다시 만든다');
  process.exit(1);
}
if (!check) console.log(stale.length ? `다시 씀: ${stale.join(', ')}` : '바꿀 것 없음');
