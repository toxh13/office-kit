#!/usr/bin/env node
// ai-smoke.mjs — AI 중계(src/app/ai.mjs)를 손으로 확인하는 실행기. 이벤트를 그대로 찍는다.
//   node scripts/dev/ai-smoke.mjs --detect
//   node scripts/dev/ai-smoke.mjs claude "OK라고만 답해"
//   node scripts/dev/ai-smoke.mjs codex "이 문서 첫 문장을 읽어 줘" --file tmp/s.hwpx --raw tmp/ai/codex.jsonl
// 선택: --file 문서  --resume 세션id  --model 모델  --cwd 폴더  --raw 원문.jsonl  --no-mcp  --cancel-after 밀리초
import { appendFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { detectProviders, startChat } from '../../src/app/ai.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const argv = process.argv.slice(2);
const opt = (name) => { const i = argv.indexOf(name); return i >= 0 ? argv.splice(i, 2)[1] : undefined; };
const flag = (name) => { const i = argv.indexOf(name); if (i >= 0) argv.splice(i, 1); return i >= 0; };

if (flag('--detect')) {
  console.log(JSON.stringify(await detectProviders(), null, 2));
  process.exit(0);
}

const file = opt('--file');
const sessionId = opt('--resume');
const model = opt('--model');
const cwd = opt('--cwd');
const raw = opt('--raw');
const cancelAfter = Number(opt('--cancel-after') ?? 0);
const noMcp = flag('--no-mcp');
const [provider, ...rest] = argv;
const prompt = rest.join(' ');
if (!provider || !prompt) {
  console.error('사용법: node scripts/dev/ai-smoke.mjs <claude|codex> "<할 말>" [--file 문서] [--resume id] [--model m] [--raw out.jsonl] [--no-mcp] [--cancel-after ms]');
  process.exit(2);
}

if (raw) { mkdirSync(dirname(resolve(raw)), { recursive: true }); writeFileSync(raw, ''); }
const mcp = noMcp ? undefined : { command: process.execPath, args: [resolve(root, 'src/mcp/server.mjs')], env: {} };

let inText = false;
const t0 = Date.now();
const stamp = () => `[${((Date.now() - t0) / 1000).toFixed(1)}s]`;
const chat = startChat({
  provider, prompt, sessionId, model, cwd, mcp,
  filePath: file ? resolve(file) : undefined,
  onRaw: raw ? (line) => appendFileSync(raw, `${line}\n`) : undefined,
  onEvent(e) {
    if (e.type === 'text') { if (!inText) { process.stdout.write(`${stamp()} 글: `); inText = true; } process.stdout.write(e.text); return; }
    if (inText) { process.stdout.write('\n'); inText = false; }
    // 터미널에서는 지울 수 없으므로 알리기만 (스트림이 끊겨 다시 받은 경우)
    if (e.type === 'text_reset') { console.log(stamp(), `↺ 끊긴 글 ${e.drop}자 지움 → 지금까지 글: ${JSON.stringify(e.text.slice(-200))}`); return; }
    console.log(stamp(), JSON.stringify(e).slice(0, 600));
  },
});
if (cancelAfter) setTimeout(() => { console.log(stamp(), '→ cancel()'); chat.cancel(); }, cancelAfter);
const res = await chat.done;
if (inText) process.stdout.write('\n');
console.log(stamp(), '결과:', JSON.stringify({ ...res, text: res.text.slice(0, 300) }));
process.exit(res.status === 'ok' ? 0 : 1);
