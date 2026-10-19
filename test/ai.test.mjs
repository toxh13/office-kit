// ai.test.mjs — AI 중계(src/app/ai.mjs) 시험. 실제 CLI·네트워크 없이 기록된 JSONL과 가짜 CLI로 확인한다.
//   fixtures/ai/claude-*.jsonl, codex-ratelimit.jsonl : 실제 기록 (개인 정보 지움)
//   fixtures/ai/*.synthetic.jsonl : 지금 기록할 수 없는 경우(로그인 안 됨, codex 한도 초과 중의 도구 호출)를 형식대로 만든 것
//   fixtures/ai/claude-abort-*.jsonl : 실제 claude.exe(2.1.301)를 가짜 로컬 API에 붙여, 스트림을 중간에 끊어 기록한 것
//     fallback = 오류 이벤트로 끊김 → 스트림 없는 요청으로 다시, retry = 연결 끊김 → 스트림으로 다시,
//     resume = 둘째 덩어리에서 끊김 → "이어서 답하라" 요청
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  createLineDecoder, createClaudeParser, createCodexParser, claudeArgs, codexArgs, friendlyError, summarize,
  toolLabel, shortToolName, systemPrompt, startChat, runCli, detectProviders, resolveClaude, resolveCodex,
  cmdShimTarget, findOnPath, PROVIDERS,
} from '../src/app/ai.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const FIX = join(here, 'fixtures', 'ai');
const FAKE = join(FIX, 'fake-cli.mjs');
const fixture = (name) => join(FIX, name);
const bytesOf = (name) => readFileSync(fixture(name));

const unhandled = [];
process.on('unhandledRejection', (e) => unhandled.push(e));

/** 바이트 조각들 → 줄 → 파서 → 이벤트 */
function parseChunks(chunks, makeParser) {
  const events = [];
  const parser = makeParser((e) => events.push(e));
  const dec = createLineDecoder((line) => parser.line(JSON.parse(line)));
  for (const c of chunks) dec.write(c);
  dec.end();
  return { events, st: parser.state };
}
const whole = (name, makeParser) => parseChunks([bytesOf(name)], makeParser);
/** 화면처럼 글 조각을 이어 붙이고, text_reset 이 오면 끝에서 drop 글자를 지운다 */
function textOf(events) {
  let s = '';
  for (const e of events) {
    if (e.type === 'text') s += e.text;
    else if (e.type === 'text_reset') {
      assert.ok(e.drop > 0 && e.drop <= s.length, `drop ${e.drop} / 글 ${s.length}`);
      s = s.slice(0, s.length - e.drop);
      assert.equal(s, e.text, 'text_reset.text 는 지운 뒤의 전체 글');
    }
  }
  return s;
}

function splitEvery(buf, n) {
  const out = [];
  for (let i = 0; i < buf.length; i += n) out.push(buf.subarray(i, i + n));
  return out;
}

function tempDir() { return mkdtempSync(join(tmpdir(), 'officekit-ai-test-')); }

// ── 줄 나누기 ────────────────────────────────────────────

test('줄 나누기: 줄 경계와 한글(3바이트) 중간에서 잘린 조각을 이어 붙인다', () => {
  const src = Buffer.from('{"a":"가나다"}\n{"b":"라마"}\r\n{"c":"끝"}', 'utf8');
  for (const n of [1, 2, 4, 5, 7, 64]) {
    const lines = [];
    const dec = createLineDecoder((l) => lines.push(l));
    for (const c of splitEvery(src, n)) dec.write(c);
    dec.end(); // 마지막 줄은 줄바꿈 없이 끝남
    assert.deepEqual(lines, ['{"a":"가나다"}', '{"b":"라마"}', '{"c":"끝"}'], `조각 ${n}바이트`);
  }
});

test('줄 나누기: 한 글자를 정확히 둘로 자른 경우 (U+FFFD 없이)', () => {
  const src = Buffer.from('{"t":"한"}\n', 'utf8');
  const at = src.indexOf(Buffer.from('한')) + 1; // 3바이트 글자의 첫 바이트 뒤
  const lines = [];
  const dec = createLineDecoder((l) => lines.push(l));
  dec.write(src.subarray(0, at));
  dec.write(src.subarray(at, at + 1));
  dec.write(src.subarray(at + 1));
  dec.end();
  assert.deepEqual(lines, ['{"t":"한"}']);
  assert.ok(!lines[0].includes('\uFFFD'));
});

// ── Claude ───────────────────────────────────────────────

test('Claude 실제 기록: 짧은 답', () => {
  const { events, st } = whole('claude-text.jsonl', createClaudeParser);
  assert.equal(st.ok, true);
  assert.equal(st.error, null);
  assert.equal(st.sessionId, 'aaaaaaaa-0000-4000-8000-000000000001');
  assert.equal(st.text, 'OK');
  assert.equal(textOf(events), 'OK');
  assert.ok(events.every((e) => e.type !== 'text' || e.delta === true));
  const status = events.find((e) => e.type === 'status');
  assert.equal(status.sessionId, st.sessionId);
});

test('Claude 실제 기록: office-kit 도구 호출 → tool / tool_result / 글', () => {
  const { events, st } = whole('claude-tool.jsonl', createClaudeParser);
  assert.equal(st.ok, true);
  const tool = events.find((e) => e.type === 'tool');
  assert.equal(tool.tool, 'office_read');
  assert.equal(tool.name, '문서 읽기');
  assert.match(tool.input.path, /s\.hwpx$/);
  const res = events.find((e) => e.type === 'tool_result');
  assert.equal(res.id, tool.id);
  assert.equal(res.ok, true);
  assert.match(res.summary, /분기 보고/);
  // 도구는 한 번만, 결과는 도구 다음
  assert.equal(events.filter((e) => e.type === 'tool').length, 1);
  assert.ok(events.indexOf(tool) < events.indexOf(res));
  // 조각을 이어 붙이면 최종 text 와 같다
  assert.equal(textOf(events), st.text);
  assert.match(st.text, /분기 보고/);
});

test('Claude 실제 기록: 1바이트씩 받아도 한 번에 받은 것과 이벤트가 같다', () => {
  const a = whole('claude-tool.jsonl', createClaudeParser).events;
  for (const n of [1, 3, 17]) {
    const b = parseChunks(splitEvery(bytesOf('claude-tool.jsonl'), n), createClaudeParser).events;
    assert.deepEqual(b, a, `조각 ${n}바이트`);
  }
});

test('Claude 실제 기록: 이어 하기(--resume)는 같은 세션 id', () => {
  const { st } = whole('claude-resume.jsonl', createClaudeParser);
  assert.equal(st.ok, true);
  assert.equal(st.sessionId, 'aaaaaaaa-0000-4000-8000-000000000002');
  assert.match(st.text, /분기 보고/);
});

test('Claude: 로그인 안 됨 안내는 글로 내보내지 않고 오류로', () => {
  const { events, st } = whole('claude-login.synthetic.jsonl', createClaudeParser);
  assert.equal(st.ok, false);
  assert.match(st.error, /Not logged in/);
  assert.equal(textOf(events), '');
  const msg = friendlyError('claude', st.error);
  assert.match(msg, /로그인이 필요합니다/);
  assert.match(msg, /'claude'/);
});

test('Claude: 조각 글(partial)이 없는 옛 형식이면 완성된 메시지 글을 한 번 내보낸다', () => {
  const lines = [
    { type: 'system', subtype: 'init', session_id: 's1', mcp_servers: [{ name: 'office-kit', status: 'failed' }] },
    { type: 'assistant', message: { id: 'm1', content: [{ type: 'text', text: '읽어 볼게요.' }, { type: 'tool_use', id: 't1', name: 'mcp__office-kit__office_info', input: { path: 'a.hwp' } }] } },
    { type: 'user', message: { content: [{ type: 'tool_result', tool_use_id: 't1', is_error: true, content: [{ type: 'text', text: '파일이 없습니다: a.hwp' }] }] } },
    { type: 'assistant', message: { id: 'm2', content: [{ type: 'text', text: '파일이 없네요.' }] } },
    { type: 'assistant', message: { id: 'sub', content: [{ type: 'text', text: '하위 에이전트' }] }, parent_tool_use_id: 'x' },
    { type: 'result', subtype: 'success', is_error: false, result: '파일이 없네요.', session_id: 's1' },
  ];
  const events = [];
  const p = createClaudeParser((e) => events.push(e));
  lines.forEach((l) => p.line(l));
  assert.match(events[0].message, /연결하지 못했습니다/);
  assert.equal(p.state.text, '읽어 볼게요.\n\n파일이 없네요.');
  assert.equal(textOf(events), p.state.text);
  assert.deepEqual(events.find((e) => e.type === 'tool_result'), { type: 'tool_result', id: 't1', ok: false, summary: '파일이 없습니다: a.hwp' });
  assert.equal(events.find((e) => e.type === 'tool').name, '문서 정보 보기');
});

test('Claude: 오류 결과(result is_error)', () => {
  const p = createClaudeParser(() => {});
  p.line({ type: 'result', subtype: 'error_during_execution', is_error: true, errors: ['API Error: 529 overloaded'], session_id: 's' });
  assert.equal(p.state.ok, false);
  assert.match(friendlyError('claude', p.state.error), /서버가 바쁩니다/);
});

// ── Claude: 스트림이 중간에 끊긴 경우 ─────────────────────

const ABORTS = [
  // [기록, 최종 글, 지울 글자 수(끊긴 조각 + 앞 빈 줄, JS 문자열 길이)]
  ['claude-abort-fallback.jsonl', '완전한 답입니다.', '첫 시도에서 끊긴 반쪽 답'.length],
  ['claude-abort-retry.jsonl', '완전한 답입니다.', '첫 시도에서 끊긴 반쪽 답'.length],
  ['claude-abort-resume.jsonl', '끝난 첫 덩어리.\n\n완전한 답입니다.', '\n\n첫 시도에서 끊긴 반쪽 답'.length],
];

/** 기록 → 이벤트. old=true 면 abandoned_blocks 표시를 빼서 옛 CLI 처럼 */
function parseAbort(name, { old = false, chunk } = {}) {
  let bytes = bytesOf(name);
  if (old) {
    bytes = Buffer.from(bytes.toString('utf8').split('\n').map((l) => {
      if (!l.includes('abandoned_blocks')) return l;
      const o = JSON.parse(l);
      delete o.abandoned_blocks;
      return JSON.stringify(o);
    }).join('\n'), 'utf8');
    assert.ok(!bytes.includes('abandoned_blocks'));
  }
  return parseChunks(chunk ? splitEvery(bytes, chunk) : [bytes], createClaudeParser);
}

for (const [name, want, drop] of ABORTS) {
  test(`Claude 실제 기록(스트림 끊김, ${name}): 끊긴 조각은 지우고 최종 글은 CLI 답과 같다`, () => {
    assert.ok(bytesOf(name).includes('abandoned_blocks'), '기록에 abandoned_blocks 표시가 있어야 함');
    const { events, st } = parseAbort(name);
    assert.equal(st.ok, true);
    assert.equal(st.error, null);
    assert.equal(st.text, want);
    assert.equal(textOf(events), want);
    const resets = events.filter((e) => e.type === 'text_reset');
    assert.equal(resets.length, 1);
    assert.equal(resets[0].drop, drop);
    // 표시가 있으면 다시 받기 전에 바로 지운다 → 끊긴 글이 화면에 남아 있는 동안 새 글이 붙지 않는다
    const i = events.indexOf(resets[0]);
    assert.ok(!events.slice(i).some((e) => e.type === 'text' && e.text.includes('끊긴')));
    assert.ok(events.slice(i).some((e) => e.type === 'text' && e.text === '완전한 답입니다.'));
    // 1바이트씩 받아도 같다
    assert.deepEqual(parseAbort(name, { chunk: 1 }).events, events);
  });

  test(`Claude(abandoned_blocks 없는 옛 CLI, ${name}): 다른 메시지가 오면 끊긴 조각을 지운다`, () => {
    const { events, st } = parseAbort(name, { old: true });
    assert.equal(st.ok, true);
    assert.equal(st.text, want);
    assert.equal(textOf(events), want);
    assert.deepEqual(events.filter((e) => e.type === 'text_reset').map((e) => e.drop), [drop]);
  });
}

test('Claude: 정상 기록에서는 지우기(text_reset)가 없다', () => {
  for (const name of ['claude-text.jsonl', 'claude-tool.jsonl', 'claude-resume.jsonl']) {
    assert.ok(!whole(name, createClaudeParser).events.some((e) => e.type === 'text_reset'), name);
  }
});

test('Claude: 확정된 글·도구는 그대로, 끊긴 메시지 뒤 새 메시지의 글만 남는다 (여러 메시지)', () => {
  const ev = (event, extra = {}) => ({ type: 'stream_event', event, session_id: 's', parent_tool_use_id: null, ...extra });
  const delta = (index, text) => ev({ type: 'content_block_delta', index, delta: { type: 'text_delta', text } });
  const lines = [
    ev({ type: 'message_start', message: { id: 'm1' } }),
    delta(0, '읽어 볼게요.'),
    { type: 'assistant', message: { id: 'm1', content: [{ type: 'text', text: '읽어 볼게요.' }] } },
    { type: 'assistant', message: { id: 'm1', content: [{ type: 'tool_use', id: 't1', name: 'mcp__office-kit__office_read', input: {} }] } },
    { type: 'user', message: { content: [{ type: 'tool_result', tool_use_id: 't1', content: '본문' }] } },
    ev({ type: 'message_start', message: { id: 'm2' } }),
    delta(0, '첫 줄은 '),
    delta(0, '반쪽'),
    // 표시 없이 끊김 → 새 시도
    ev({ type: 'message_start', message: { id: 'm3' } }),
    delta(0, '첫 줄은 "본문"입니다.'),
    { type: 'assistant', message: { id: 'm3', content: [{ type: 'text', text: '첫 줄은 "본문"입니다.' }] } },
    { type: 'result', subtype: 'success', is_error: false, result: '첫 줄은 "본문"입니다.' },
  ];
  const events = [];
  const p = createClaudeParser((e) => events.push(e));
  lines.forEach((l) => p.line(l));
  assert.equal(p.state.text, '읽어 볼게요.\n\n첫 줄은 "본문"입니다.');
  assert.equal(textOf(events), p.state.text);
  assert.deepEqual(events.filter((e) => e.type === 'text_reset'), [{ type: 'text_reset', text: '읽어 볼게요.', drop: '\n\n첫 줄은 반쪽'.length }]);
  assert.equal(events.filter((e) => e.type === 'tool').length, 1);
});

// ── Codex ────────────────────────────────────────────────

test('Codex 실제 기록: 사용 한도 초과 → 한국어 안내 + 다시 쓸 수 있는 때', () => {
  const { events, st } = whole('codex-ratelimit.jsonl', createCodexParser);
  assert.equal(st.ok, false);
  assert.equal(st.sessionId, 'bbbbbbbb-0000-7000-8000-000000000001');
  assert.match(st.error, /usage limit/);
  assert.equal(textOf(events), '');
  const msg = friendlyError('codex', st.error);
  assert.match(msg, /사용 한도에 걸렸습니다/);
  assert.match(msg, /Oct 22nd, 2026 3:14 PM/);
});

test('Codex(형식대로 만든 기록): 도구 시작·성공·실패, 글 사이 빈 줄', () => {
  const { events, st } = whole('codex-tool.synthetic.jsonl', createCodexParser);
  assert.equal(st.ok, true);
  assert.equal(st.sessionId, 'bbbbbbbb-0000-7000-8000-000000000002');
  const tools = events.filter((e) => e.type === 'tool');
  assert.deepEqual(tools.map((t) => [t.id, t.tool, t.name]), [['item_2', 'office_read', '문서 읽기'], ['item_4', 'office_page_image', '쪽 이미지로 확인']]);
  assert.equal(tools[0].input.path, 'C:\\docs\\보고서 s.hwpx');
  const results = events.filter((e) => e.type === 'tool_result');
  assert.deepEqual(results.map((r) => [r.id, r.ok]), [['item_2', true], ['item_4', false]]);
  assert.match(results[0].summary, /분기 보고/);
  assert.match(results[1].summary, /범위를 벗어났습니다/);
  assert.equal(st.text, '문서를 읽어 볼게요.\n\n첫 줄은 **분기 보고** 입니다. 문서는 고치지 않았어요.');
  assert.equal(textOf(events), st.text);
  assert.deepEqual(st.usage, { input_tokens: 12034, cached_input_tokens: 8064, output_tokens: 118 });
});

test('Codex: 조각나 들어와도 같은 이벤트, 글이 늘어나는 item.updated 는 늘어난 만큼만', () => {
  const a = whole('codex-tool.synthetic.jsonl', createCodexParser).events;
  const b = parseChunks(splitEvery(bytesOf('codex-tool.synthetic.jsonl'), 2), createCodexParser).events;
  assert.deepEqual(b, a);
  const events = [];
  const p = createCodexParser((e) => events.push(e));
  p.line({ type: 'item.updated', item: { id: 'm', type: 'agent_message', text: '안녕' } });
  p.line({ type: 'item.completed', item: { id: 'm', type: 'agent_message', text: '안녕하세요' } });
  assert.deepEqual(events.map((e) => e.text), ['안녕', '하세요']);
});

test('Codex: turn.completed 없이 끝나면 마지막 오류를 쓴다', () => {
  const p = createCodexParser(() => {});
  p.line({ type: 'thread.started', thread_id: 't' });
  p.line({ type: 'error', message: 'stream disconnected before completion: error sending request' });
  assert.equal(p.state.ok, false);
  assert.match(friendlyError('codex', p.state.lastError), /네트워크/);
});

// ── 인수 ─────────────────────────────────────────────────

test('claude 인수: 내장 도구 끔, office-kit 만 허용, 위험한 옵션 없음, 할 말은 인수에 없음', () => {
  const args = claudeArgs({ sessionId: 'sid', model: 'sonnet', mcpConfigPath: 'C:\\t\\mcp.json', filePath: 'C:\\문서\\보고서.hwpx' });
  const at = (k) => args[args.indexOf(k) + 1];
  assert.equal(args[0], '-p');
  assert.equal(at('--output-format'), 'stream-json');
  assert.ok(args.includes('--verbose') && args.includes('--include-partial-messages') && args.includes('--strict-mcp-config'));
  assert.equal(at('--tools'), '');
  assert.equal(at('--permission-mode'), 'dontAsk');
  assert.equal(at('--allowedTools'), 'mcp__office-kit');
  assert.equal(at('--mcp-config'), 'C:\\t\\mcp.json');
  assert.equal(at('--resume'), 'sid');
  assert.equal(at('--model'), 'sonnet');
  assert.match(at('--append-system-prompt'), /보고서\.hwpx/);
  assert.ok(!args.some((a) => /dangerously|bypassPermissions|acceptEdits/i.test(a)));
  // 여러 값을 받는 옵션 뒤에 위치 인수(할 말)가 붙지 않는다
  assert.ok(args.every((a, i) => i === 0 || a.startsWith('-') || args[i - 1].startsWith('-')));
  assert.ok(!claudeArgs({}).includes('--resume'));
  assert.ok(!claudeArgs({}).includes('--allowedTools'));
});

test('codex 인수: 읽기 전용, 셸 끔, MCP 주입(TOML 따옴표), 자동 승인, 이어 하기 꼴', () => {
  const mcp = { command: 'C:\\Program Files\\Office Kit\\Office Kit.exe', args: ['C:\\앱 폴더\\src\\mcp\\server.mjs'], env: { ELECTRON_RUN_AS_NODE: '1' } };
  const args = codexArgs({ cwd: 'C:\\work', mcp, model: 'gpt-x' });
  const cfg = args.filter((_, i) => args[i - 1] === '-c');
  const val = (k) => cfg.find((c) => c.startsWith(`${k}=`))?.slice(k.length + 1);
  assert.deepEqual(args.slice(0, 1), ['exec']);
  assert.deepEqual(args.slice(-3), ['-C', 'C:\\work', '-']);
  assert.ok(args.includes('--json') && args.includes('--skip-git-repo-check') && args.includes('--ignore-user-config'));
  assert.equal(val('sandbox_mode'), '"read-only"');
  assert.equal(val('approval_policy'), '"never"');
  assert.equal(val('features.shell_tool'), 'false');
  assert.equal(val('features.unified_exec'), 'false');
  // TOML 기본 문자열은 JSON 문자열과 같은 꼴 → JSON 으로 되읽어 원래 값이 나와야 한다
  assert.equal(JSON.parse(val('mcp_servers.office-kit.command')), mcp.command);
  assert.deepEqual(JSON.parse(val('mcp_servers.office-kit.args')), mcp.args);
  assert.equal(val('mcp_servers.office-kit.env'), '{ "ELECTRON_RUN_AS_NODE" = "1" }');
  assert.equal(val('mcp_servers.office-kit.default_tools_approval_mode'), '"approve"');
  assert.equal(args[args.indexOf('-m') + 1], 'gpt-x');
  assert.ok(!args.some((a) => /dangerously|danger-full-access|workspace-write/i.test(a)));

  const resume = codexArgs({ sessionId: 'thread-1', mcp });
  assert.deepEqual(resume.slice(0, 2), ['exec', 'resume']);
  assert.deepEqual(resume.slice(-2), ['thread-1', '-']);
  assert.ok(!resume.includes('-C') && !resume.includes('-s')); // resume 에는 없는 옵션
});

test('지침: 문서 경로·종류, office-kit 도구만, 확인은 office_page_image, 짧은 한국어', () => {
  const s = systemPrompt({ filePath: 'C:\\문서\\신청서.hwpx' });
  assert.match(s, /C:\\문서\\신청서\.hwpx \(한글 문서, \.hwpx\)/);
  assert.match(s, /office-kit 도구/);
  assert.match(s, /office_page_image/);
  assert.match(s, /요청하지 않은/);
  assert.match(s, /짧은 한국어/);
  assert.match(systemPrompt({}), /현재 열린 문서: 없음/);
});

// ── 보조 함수 ────────────────────────────────────────────

test('도구 이름·요약', () => {
  assert.equal(shortToolName('mcp__office-kit__office_page_image'), 'office_page_image');
  assert.equal(shortToolName('mcp__office_kit__hwp_edit'), 'hwp_edit');
  assert.equal(toolLabel('mcp__office-kit__office_replace'), '찾아 바꾸기');
  assert.equal(toolLabel('mystery_tool'), 'mystery_tool');
  assert.equal(summarize([{ type: 'text', text: '가\n\n나' }, { type: 'image', data: 'x' }]), '가 나 [이미지]');
  // Claude CLI가 그림 결과에 붙이는 내부 경로는 보이지 않게
  assert.equal(summarize('[Image: source: C:\\Users\\me\\.claude\\projects\\x\\a.png] 2쪽'), '[이미지] 2쪽');
  const long = summarize('한'.repeat(500), 10);
  assert.equal([...long].length, 10);
  assert.ok(long.endsWith('…'));
});

test('오류 안내 한국어로', () => {
  assert.match(friendlyError('claude', 'spawn C:\\x\\claude.exe ENOENT'), /명령을 찾지 못했습니다.*설치 방법/);
  assert.match(friendlyError('codex', 'Error: 401 Unauthorized'), /codex login/);
  assert.match(friendlyError('claude', "You've hit your limit · resets 5pm (Asia/Seoul)"), /사용 한도.*5pm \(Asia\/Seoul\)/);
  assert.match(friendlyError('claude', 'Claude AI usage limit reached|1792251000'), /사용 한도.*다시 쓸 수 있는 때: 2026/);
  assert.match(friendlyError('claude', 'No conversation found with session ID: abc'), /이전 대화를 찾지 못했습니다/);
  assert.match(friendlyError('codex', 'getaddrinfo ENOTFOUND api.openai.com'), /네트워크/);
  assert.match(friendlyError('codex', 'something odd happened\nmore'), /오류가 났습니다: something odd happened$/);
  assert.match(friendlyError('codex', '', { exitCode: 3 }), /종료 코드 3/);
});

// ── 실행 파일 찾기 (가짜 폴더, 실제 설정은 건드리지 않음) ─────

test('실행 파일 찾기: npm .cmd 껍데기, codex 네이티브 exe 우선, 없으면 node + codex.js', { skip: process.platform !== 'win32' }, () => {
  const dir = tempDir();
  try {
    const home = join(dir, 'home');
    const npm = join(dir, 'AppData', 'Roaming', 'npm');
    const pkg = join(npm, 'node_modules', '@openai', 'codex');
    mkdirSync(join(pkg, 'bin'), { recursive: true });
    mkdirSync(home, { recursive: true });
    writeFileSync(join(pkg, 'bin', 'codex.js'), '');
    writeFileSync(join(npm, 'codex.cmd'), '@ECHO off\r\n"%_prog%"  "%dp0%\\node_modules\\@openai\\codex\\bin\\codex.js" %*\r\n');
    const nodeDir = join(dir, 'nodejs');
    mkdirSync(nodeDir);
    writeFileSync(join(nodeDir, 'node.exe'), '');
    const env = { PATH: [npm, nodeDir].join(';'), USERPROFILE: home, APPDATA: join(dir, 'AppData', 'Roaming') };

    assert.equal(cmdShimTarget(join(npm, 'codex.cmd')), join(pkg, 'bin', 'codex.js'));
    assert.equal(findOnPath('codex', env), join(npm, 'codex.cmd'));
    // 네이티브 exe 가 없으면 PATH 의 node 로 codex.js
    assert.deepEqual(resolveCodex(env), { command: join(nodeDir, 'node.exe'), args: [join(pkg, 'bin', 'codex.js')], env: {} });
    // 네이티브 exe 가 있으면 그것을 바로
    const triple = { x64: 'x86_64-pc-windows-msvc', arm64: 'aarch64-pc-windows-msvc' }[process.arch];
    const exe = join(pkg, 'node_modules', '@openai', `codex-win32-${process.arch}`, 'vendor', triple, 'bin', 'codex.exe');
    mkdirSync(dirname(exe), { recursive: true });
    writeFileSync(exe, '');
    const r = resolveCodex(env);
    assert.equal(r.command, exe);
    assert.equal(r.env.CODEX_MANAGED_PACKAGE_ROOT, pkg);

    // claude: ~/.local/bin/claude.exe
    assert.equal(resolveClaude(env), null);
    mkdirSync(join(home, '.local', 'bin'), { recursive: true });
    writeFileSync(join(home, '.local', 'bin', 'claude.exe'), '');
    assert.deepEqual(resolveClaude(env), { command: join(home, '.local', 'bin', 'claude.exe'), args: [], env: {} });
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('codex 찾기: PATH 의 독립 설치본이 %APPDATA%\\npm 에 남은 옛 설치본보다 먼저', { skip: process.platform !== 'win32' }, () => {
  const dir = tempDir();
  try {
    const appdata = join(dir, 'AppData', 'Roaming');
    const triple = { x64: 'x86_64-pc-windows-msvc', arm64: 'aarch64-pc-windows-msvc' }[process.arch];
    // npm 패키지(네이티브 exe 포함) 하나 만들기
    const makePkg = (pkg) => {
      const exe = join(pkg, 'node_modules', '@openai', `codex-win32-${process.arch}`, 'vendor', triple, 'bin', 'codex.exe');
      mkdirSync(dirname(exe), { recursive: true });
      mkdirSync(join(pkg, 'bin'), { recursive: true });
      writeFileSync(exe, '');
      writeFileSync(join(pkg, 'bin', 'codex.js'), '');
      return exe;
    };
    const staleExe = makePkg(join(appdata, 'npm', 'node_modules', '@openai', 'codex')); // 옛 npm 설치본 (PATH 에는 없음)
    const standalone = join(dir, 'standalone');
    mkdirSync(standalone);
    writeFileSync(join(standalone, 'codex.exe'), '');
    const base = { USERPROFILE: join(dir, 'home'), APPDATA: appdata };

    // 1) PATH 에 진짜 codex.exe → 그것을 그대로 (npm 표시 환경 변수 없음)
    const env1 = { ...base, PATH: standalone };
    assert.equal(findOnPath('codex', env1), join(standalone, 'codex.exe'));
    assert.deepEqual(resolveCodex(env1), { command: join(standalone, 'codex.exe'), args: [], env: {} });

    // 2) PATH 에 아무것도 없으면 (앱이 PATH 를 늦게 물려받은 경우) %APPDATA%\npm 설치본
    const r2 = resolveCodex({ ...base, PATH: '' });
    assert.equal(r2.command, staleExe);
    assert.equal(r2.env.CODEX_MANAGED_BY_NPM, '1');

    // 3) PATH 의 .cmd 가 다른 곳의 패키지를 가리키면 (pnpm 등) 그 패키지의 exe, %APPDATA%\npm 것이 아니라
    const pnpm = join(dir, 'pnpm');
    const pnpmPkg = join(pnpm, 'global', '5', 'node_modules', '@openai', 'codex');
    const pnpmExe = makePkg(pnpmPkg);
    writeFileSync(join(pnpm, 'codex.cmd'), '@ECHO off\r\n"%~dp0\\global\\5\\node_modules\\@openai\\codex\\bin\\codex.js" %*\r\n');
    const r3 = resolveCodex({ ...base, PATH: pnpm });
    assert.equal(r3.command, pnpmExe);
    assert.equal(r3.env.CODEX_MANAGED_PACKAGE_ROOT, pnpmPkg);

    // 4) PATH 의 .cmd 를 해석할 수 없으면 %APPDATA%\npm 설치본으로
    const broken = join(dir, 'broken');
    mkdirSync(broken);
    writeFileSync(join(broken, 'codex.cmd'), '@ECHO off\r\nREM 알 수 없는 꼴\r\n');
    assert.equal(resolveCodex({ ...base, PATH: broken }).command, staleExe);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('설치 안 됨: 두 AI 모두 available=false + 설치 안내', async () => {
  const dir = tempDir();
  try {
    const env = { PATH: '', USERPROFILE: dir, APPDATA: join(dir, 'Roaming'), HOME: dir };
    const list = await detectProviders({ env });
    assert.deepEqual(list.map((p) => [p.id, p.available]), [['claude', false], ['codex', false]]);
    assert.equal(list[0].name, PROVIDERS.claude.name);
    assert.match(list[0].name, /내 Claude Code 로그인 사용/);
    assert.match(list[1].name, /내 ChatGPT 로그인 사용/);
    for (const p of list) assert.match(p.hint, /설치 방법/);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

// ── 끝까지 실행 (가짜 CLI) ───────────────────────────────

function fakeChat(provider, fixtureName, extra = {}) {
  const dir = tempDir();
  const recordPath = join(dir, 'record.json');
  const events = [];
  const chat = startChat({
    provider,
    prompt: '첫 줄을 "그대로" 알려 줘 — 한글 확인',
    filePath: 'C:\\문서\\보고서 s.hwpx',
    cwd: dir,
    mcp: { command: process.execPath, args: ['C:\\앱\\server.mjs'], env: { ELECTRON_RUN_AS_NODE: '1' } },
    cli: { command: process.execPath, args: [FAKE] },
    env: { ...process.env, FAKE_RECORD: recordPath, FAKE_FIXTURE: fixtureName ? fixture(fixtureName) : '', FAKE_CHUNK: '3', ...extra },
    onEvent: (e) => events.push(e),
    ...extra.opts,
  });
  return { chat, events, dir, record: () => JSON.parse(readFileSync(recordPath, 'utf8')) };
}

test('startChat(claude, 가짜 CLI): 할 말은 stdin 으로, MCP 설정 임시 파일은 끝나면 지움', async () => {
  const { chat, events, dir, record } = fakeChat('claude', 'claude-tool.jsonl');
  try {
    const res = await chat.done;
    assert.equal(res.status, 'ok');
    assert.equal(res.sessionId, 'aaaaaaaa-0000-4000-8000-000000000002');
    const rec = record();
    assert.equal(rec.stdin, '첫 줄을 "그대로" 알려 줘 — 한글 확인');
    assert.ok(!rec.argv.some((a) => a.includes('첫 줄을')));
    assert.deepEqual(rec.mcpConfig.mcpServers['office-kit'], { type: 'stdio', command: process.execPath, args: ['C:\\앱\\server.mjs'], env: { ELECTRON_RUN_AS_NODE: '1' } });
    assert.ok(!existsSync(rec.mcpConfigPath), '임시 MCP 설정이 남아 있음');
    assert.equal(resolve(rec.cwd), resolve(dir));
    assert.equal(events.at(-1).type, 'done');
    assert.equal(events.at(-1).text, res.text);
    assert.equal(textOf(events), res.text);
    assert.equal(events.filter((e) => e.type === 'done' || e.type === 'error').length, 1);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('startChat(claude, 가짜 CLI): 스트림이 끊겼다 다시 받은 기록 → done.text 는 CLI 답과 같다', async () => {
  const { chat, events, dir } = fakeChat('claude', 'claude-abort-fallback.jsonl');
  try {
    const res = await chat.done;
    assert.equal(res.status, 'ok');
    assert.equal(res.text, '완전한 답입니다.');
    assert.equal(events.at(-1).type, 'done');
    assert.equal(events.at(-1).text, res.text);
    assert.equal(textOf(events), res.text);
    assert.equal(events.filter((e) => e.type === 'text_reset').length, 1);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('startChat(codex, 가짜 CLI): 지침이 맨 위, 이어 하기는 exec resume', async () => {
  const a = fakeChat('codex', 'codex-tool.synthetic.jsonl');
  try {
    const res = await a.chat.done;
    assert.equal(res.status, 'ok');
    const rec = a.record();
    assert.ok(rec.stdin.startsWith('너는 Office Kit 앱 안에서'));
    assert.match(rec.stdin, /---\n\n첫 줄을 "그대로" 알려 줘/);
    assert.equal(rec.argv.at(-1), '-');
    assert.deepEqual(a.events.filter((e) => e.type === 'tool_result').map((e) => e.ok), [true, false]);
  } finally { rmSync(a.dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }

  const b = fakeChat('codex', 'codex-tool.synthetic.jsonl', { opts: { sessionId: 'thread-9' } });
  try {
    await b.chat.done;
    const rec = b.record();
    assert.deepEqual(rec.argv.slice(0, 2), ['exec', 'resume']);
    assert.deepEqual(rec.argv.slice(-2), ['thread-9', '-']);
    assert.equal(rec.stdin, '(현재 열린 문서: C:\\문서\\보고서 s.hwpx)\n\n첫 줄을 "그대로" 알려 줘 — 한글 확인');
  } finally { rmSync(b.dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('startChat: 실제 한도 초과 기록 + 종료 코드 1 → error 이벤트 하나, done 은 거부되지 않음', async () => {
  const { chat, events, dir } = fakeChat('codex', 'codex-ratelimit.jsonl', { FAKE_EXIT: '1', FAKE_STDERR: 'Reading prompt from stdin...\n' });
  try {
    const res = await chat.done;
    assert.equal(res.status, 'error');
    assert.match(res.error, /사용 한도/);
    const last = events.at(-1);
    assert.equal(last.type, 'error');
    assert.match(last.detail, /usage limit/);
    assert.equal(events.filter((e) => e.type === 'done').length, 0);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('startChat: 결과 없이 끝나면 stderr 로 안내', async () => {
  const { chat, dir } = fakeChat('claude', null, { FAKE_EXIT: '1', FAKE_STDERR: 'Error: Invalid API key\n' });
  try {
    const res = await chat.done;
    assert.equal(res.status, 'error');
    assert.match(res.error, /로그인이 필요합니다/);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('startChat: 실행 파일이 없으면 한국어 오류, 잘못된 입력도 거부 없이 끝남', async () => {
  const events = [];
  const res = await startChat({ provider: 'claude', prompt: 'hi', cwd: tmpdir(), cli: { command: join(tmpdir(), 'no-such-dir', 'claude.exe') }, onEvent: (e) => events.push(e) }).done;
  assert.equal(res.status, 'error');
  assert.match(res.error, /명령을 찾지 못했습니다/);
  assert.equal(events.at(-1).type, 'error');

  assert.match((await startChat({ provider: 'gemini', prompt: 'hi' }).done).error, /알 수 없는 AI/);
  assert.match((await startChat({ provider: 'codex', prompt: '  ' }).done).error, /비어 있습니다/);
});

test('cancel(): 프로세스 트리(손자까지) 종료, done 은 cancelled 로 끝남', async () => {
  const dir = tempDir();
  const recordPath = join(dir, 'record.json');
  const events = [];
  let ready;
  const started = new Promise((r) => { ready = r; });
  const chat = startChat({
    provider: 'claude', prompt: '오래 걸리는 일', cwd: dir,
    cli: { command: process.execPath, args: [FAKE] },
    env: { ...process.env, FAKE_HANG: '1', FAKE_RECORD: recordPath },
    onEvent: (e) => { events.push(e); if (e.type === 'status' && e.sessionId) ready(); },
  });
  try {
    await started;
    const { grandchild } = JSON.parse(readFileSync(recordPath, 'utf8'));
    assert.ok(grandchild > 0);
    assert.doesNotThrow(() => process.kill(grandchild, 0), '손자 프로세스가 떠 있어야 함');
    const t0 = Date.now();
    const res = await chat.cancel();
    assert.equal(res.status, 'cancelled');
    assert.equal(res.cancelled, true);
    assert.equal(res.sessionId, 'hang-1');
    assert.ok(Date.now() - t0 < 5000, '5초 안에 끝나야 함');
    assert.deepEqual(events.at(-1), { type: 'done', sessionId: 'hang-1', text: '', cancelled: true });
    assert.equal(await chat.done, res);
    // 손자까지 죽었는지 (taskkill 은 비동기라 잠깐 기다림)
    let alive = true;
    for (let i = 0; i < 50 && alive; i++) {
      try { process.kill(grandchild, 0); await new Promise((r) => setTimeout(r, 100)); } catch { alive = false; }
    }
    assert.equal(alive, false, '손자 프로세스가 남아 있음');
    // 두 번 불러도 괜찮다
    assert.equal(await chat.cancel(), res);
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('runCli: onEvent 가 예외를 던져도 대화는 끝까지 간다', async () => {
  const dir = tempDir();
  try {
    const res = await runCli({
      provider: 'claude', command: process.execPath, args: [FAKE], cwd: dir,
      env: { ...process.env, FAKE_FIXTURE: fixture('claude-text.jsonl'), FAKE_CHUNK: '1' },
      makeParser: createClaudeParser,
      onEvent: () => { throw new Error('화면 오류'); },
    }).done;
    assert.equal(res.status, 'ok');
    assert.equal(res.text, 'OK');
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }); }
});

test('처리되지 않은 거부(unhandled rejection) 없음', async () => {
  await new Promise((r) => setTimeout(r, 50));
  assert.deepEqual(unhandled, []);
});
