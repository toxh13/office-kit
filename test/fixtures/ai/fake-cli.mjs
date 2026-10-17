// fake-cli.mjs — 시험용 가짜 claude/codex. 받은 인수·입력을 기록하고, 기록된 JSONL을 잘게 쪼개 내보낸다.
//   FAKE_RECORD  받은 것(argv, stdin, mcp 설정)을 적을 파일
//   FAKE_FIXTURE 내보낼 JSONL, FAKE_CHUNK 조각 크기(바이트), FAKE_EXIT 종료 코드, FAKE_STDERR stderr 내용
//   FAKE_HANG=1  손자 프로세스를 띄우고 끝나지 않는다 (중단 시험)
import { readFileSync, writeFileSync } from 'node:fs';
import { spawn } from 'node:child_process';

const argv = process.argv.slice(2);
const env = process.env;
// node --test 가 이 파일을 시험으로 집어도 바로 끝나게
if (!env.FAKE_RECORD && !env.FAKE_FIXTURE && !env.FAKE_HANG) process.exit(0);
let stdin = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (d) => { stdin += d; });
process.stdin.on('end', main);

function record(extra = {}) {
  const rec = { argv, stdin, cwd: process.cwd(), ...extra };
  const i = argv.indexOf('--mcp-config');
  if (i >= 0) { rec.mcpConfigPath = argv[i + 1]; rec.mcpConfig = JSON.parse(readFileSync(argv[i + 1], 'utf8')); }
  if (env.FAKE_RECORD) writeFileSync(env.FAKE_RECORD, JSON.stringify(rec), 'utf8');
}

async function main() {
  if (env.FAKE_HANG) {
    const gc = spawn(process.execPath, ['-e', 'setInterval(() => {}, 1000)'], { stdio: 'ignore' });
    record({ grandchild: gc.pid });
    process.stdout.write(`${JSON.stringify({ type: 'system', subtype: 'init', session_id: 'hang-1', mcp_servers: [] })}\n`);
    setInterval(() => {}, 1000);
    return;
  }
  record();
  if (env.FAKE_FIXTURE) {
    const bytes = readFileSync(env.FAKE_FIXTURE);
    const n = Number(env.FAKE_CHUNK || 5);
    for (let o = 0; o < bytes.length; o += n) {
      process.stdout.write(bytes.subarray(o, o + n));
      await new Promise((r) => setImmediate(r));
    }
  }
  if (env.FAKE_STDERR) process.stderr.write(env.FAKE_STDERR);
  process.exitCode = Number(env.FAKE_EXIT || 0);
}
