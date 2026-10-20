// mcp.test.mjs — MCP 서버를 실제로 띄워(stdio) SDK 클라이언트로 도구를 호출한다.
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { MANUSCRIPT, COM_TESTS, ROOT, tempDir, isPng, pngSize } from './helpers/fixtures.mjs';

const SERVER = join(ROOT, 'src', 'mcp', 'server.mjs');
const CLI = join(ROOT, 'src', 'cli.mjs');
// 서버가 등록하는 도구: 공통 office_* + 형식마다(src/mcp/tools/*.mjs) 살펴보기·편집 한 쌍.
// 도구를 더하거나 빼면 여기도 고쳐야 한다 (빠지거나 늘어난 도구를 잡는다).
const COMMON_TOOLS = ['office_status', 'office_info', 'office_read', 'office_fields', 'office_fill', 'office_replace',
  'office_set_cells', 'office_convert', 'office_create', 'office_page_image', 'office_find_pii', 'office_mask_pii', 'office_hwpx_batch'];
const FORMATS = ['hwp', 'word', 'ppt', 'excel'];
const TOOLS = [...COMMON_TOOLS, ...FORMATS.flatMap((f) => [`${f}_inspect`, `${f}_edit`])];
const HANGUL = /[가-힣]/;

const tmp = tempDir('officekit-mcp-');

async function connect(args, env) {
  const client = new Client({ name: 'office-kit-test', version: '1.0.0' });
  const transport = new StdioClientTransport({ command: process.execPath, args, cwd: tmp.dir, stderr: 'pipe', ...(env ? { env: { ...process.env, ...env } } : {}) });
  const errors = [];
  transport.stderr?.on('data', (d) => errors.push(String(d)));
  await client.connect(transport);
  return { client, errors };
}

let client;
let serverLog;
before(async () => { ({ client, errors: serverLog } = await connect([SERVER])); });
after(async () => { await client?.close(); tmp.cleanup(); });

const call = (name, args = {}) => client.callTool({ name, arguments: args }, undefined, { timeout: 180_000 });
const textOf = (r) => r.content.filter((c) => c.type === 'text').map((c) => c.text).join('\n');
async function ok(name, args) {
  const r = await call(name, args);
  assert.ok(!r.isError, `${name} 실패: ${textOf(r)}`);
  return textOf(r);
}
async function fails(name, args, re) {
  const r = await call(name, args);
  assert.equal(r.isError, true, `${name}가 실패해야 함: ${textOf(r)}`);
  assert.match(textOf(r), re);
  return textOf(r);
}

describe('도구 목록', () => {
  test(`${TOOLS.length}개 도구 (공통 ${COMMON_TOOLS.length} + 형식별 ${FORMATS.length * 2}), 한국어 설명`, async () => {
    const { tools } = await client.listTools();
    assert.deepEqual(tools.map((t) => t.name).sort(), [...TOOLS].sort());
    for (const t of tools) {
      assert.match(t.title ?? '', HANGUL, `${t.name} 제목`);
      assert.match(t.description ?? '', HANGUL, `${t.name} 설명`);
      assert.equal(t.inputSchema.type, 'object', t.name);
      for (const [k, v] of Object.entries(t.inputSchema.properties ?? {})) assert.match(v.description ?? '', HANGUL, `${t.name}.${k} 설명`);
    }
    // 형식별 도구: 모듈(src/mcp/tools/*.mjs)마다 한 쌍. 살펴보기는 읽기 전용, 편집은 path + operations 필수
    const modules = readdirSync(join(ROOT, 'src', 'mcp', 'tools')).filter((f) => f.endsWith('.mjs')).map((f) => f.slice(0, -4));
    assert.deepEqual(modules.sort(), [...FORMATS].sort(), '도구 모듈과 FORMATS가 다름');
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));
    for (const f of FORMATS) {
      const inspect = byName[`${f}_inspect`];
      const edit = byName[`${f}_edit`];
      assert.equal(inspect.annotations?.readOnlyHint, true, `${f}_inspect readOnlyHint`);
      assert.ok(inspect.inputSchema.required?.includes('path'), `${f}_inspect path`);
      assert.equal(edit.annotations?.destructiveHint, true, `${f}_edit destructiveHint`);
      for (const k of ['path', 'operations']) assert.ok(edit.inputSchema.required?.includes(k), `${f}_edit ${k}`);
    }
    const create = tools.find((t) => t.name === 'office_create');
    assert.match(create.description, /---page---/);
    assert.match(create.description, /\{\{필드\}\}/);
    const edit = tools.find((t) => t.name === 'hwp_edit');
    assert.ok(JSON.stringify(edit.inputSchema).includes('add_table'));
  });

  test('office_status', async () => {
    const t = await ok('office_status');
    assert.match(t, /설치된 프로그램/);
    assert.match(t, /한글 자동화 보안 모듈/);
  });
});

describe('한글 문서 흐름', () => {
  test('office_create (작업 폴더 기준 상대 경로) + 덮어쓰기 거부', async () => {
    const t = await ok('office_create', { output: 'report.hwp', markdown: MANUSCRIPT });
    assert.match(t, /report\.hwp/);
    assert.ok(existsSync(tmp.path('report.hwp')));
    await fails('office_create', { output: 'report.hwp', markdown: '# x' }, /이미 있는 파일/);
    await ok('office_create', { output: 'sub/dir/new.hwpx', markdown: '# 새 문서', overwrite: true });
    assert.ok(existsSync(tmp.path('sub/dir/new.hwpx')));
  });

  test('office_info / office_fields', async () => {
    const info = await ok('office_info', { path: 'report.hwp' });
    assert.match(info, /한글/);
    assert.match(info, /"pages": 2/);
    const fields = await ok('office_fields', { path: tmp.path('report.hwp') });
    assert.match(fields, /담당자 \[누름틀\]/);
  });

  test('office_read: 구조 + 본문', async () => {
    const t = await ok('office_read', { path: 'report.hwp' });
    assert.match(t, /끝까지 읽음/);
    assert.match(t, /\[구조\]/);
    assert.match(t, /표 1 \(3행×2열\)/);
    assert.match(t, /분기 보고/);
  });

  test('office_replace → office_fill (다른 파일로 저장)', async () => {
    let t = await ok('office_replace', { path: 'report.hwp', replacements: [{ find: '매출', replace: '수익' }], output: 'filled.hwpx' });
    assert.match(t, /2곳 바꿈/);
    t = await ok('office_fill', { path: 'filled.hwpx', values: { 담당자: '홍길동', 없는칸: '?' } });
    assert.match(t, /채움 1개: 담당자/);
    assert.match(t, /못 찾음 1개: 없는칸/);
    const read = await ok('office_read', { path: 'filled.hwpx' });
    assert.ok(read.includes('홍길동') && read.includes('수익 증가') && !read.includes('매출'), read);
    assert.ok((await ok('office_read', { path: 'report.hwp' })).includes('매출'), '원본이 바뀜');
  });

  test('office_convert hwpx → hwp', async () => {
    const t = await ok('office_convert', { path: 'filled.hwpx', output: 'filled.hwp' });
    assert.match(t, /변환 완료/);
    assert.ok(existsSync(tmp.path('filled.hwp')));
    await fails('office_convert', { path: 'filled.hwp', output: 'filled.hwp' }, /다른 경로/);
  });

  test('office_page_image: 그림 + 설명', async () => {
    const r = await call('office_page_image', { path: 'filled.hwp', page: 2, dpi: 72 });
    assert.ok(!r.isError, textOf(r));
    const img = r.content.find((c) => c.type === 'image');
    assert.equal(img.mimeType, 'image/png');
    const png = Buffer.from(img.data, 'base64');
    assert.ok(isPng(png));
    assert.ok(Math.abs(pngSize(png).width - 595) <= 2, JSON.stringify(pngSize(png)));
    assert.match(textOf(r), /2\/2쪽/);
    await fails('office_page_image', { path: 'filled.hwp', page: 5 }, /범위를 벗어났습니다/);
  });

  test('hwp_edit: 여러 작업을 한 번에 저장', async () => {
    const t = await ok('hwp_edit', {
      path: 'report.hwp',
      output: 'edited.hwpx',
      operations: [
        { op: 'add_page_break' },
        { op: 'add_heading', text: '추가 장', level: 2 },
        { op: 'add_paragraph', text: '굵은 파란 문단', style: { bold: true, color: '#0000ff', align: 'center' } },
        { op: 'add_table', rows: [['구분', '수량'], ['사과', 3]] },
        { op: 'add_field', name: '서명', label: '서명' },
        { op: 'replace', find: '비용', replace: '지출' },
        { op: 'fill', values: { 담당자: '이영희', 서명: '김철수', 없는칸: 'x' } },
      ],
    });
    assert.match(t, /2곳/);
    assert.match(t, /채우기: 담당자, 서명 \/ 못 찾음: 없는칸/);
    assert.match(t, /저장: .*edited\.hwpx \(3쪽\)/);
    const read = await ok('office_read', { path: 'edited.hwpx' });
    for (const w of ['추가 장', '굵은 파란 문단', '사과', '이영희', '김철수', '지출 절감']) assert.ok(read.includes(w), `${w} 없음\n${read}`);
    assert.match(read, /표 2개/);
    const fields = await ok('office_fields', { path: 'edited.hwpx' });
    assert.match(fields, /서명 \[누름틀\] 현재값: 김철수/);
    const img = await call('office_page_image', { path: 'edited.hwpx', page: 3 });
    assert.ok(!img.isError, textOf(img));
  });

  test('hwp_edit: 잘못된 입력은 저장하지 않음', async () => {
    await fails('hwp_edit', { path: 'report.hwp', operations: [{ op: 'add_heading', text: 'x', level: 9 }] }, /입력값이 올바르지 않습니다/);
    await fails('hwp_edit', { path: 'report.hwp', operations: [{ op: 'explode' }] }, /입력값이 올바르지 않습니다/);
    await fails('hwp_edit', { path: 'report.hwp', operations: [{ op: 'add_page_break' }], output: 'x.docx' }, /\.hwp 또는 \.hwpx/);
    assert.ok(!existsSync(tmp.path('x.docx')));
  });
});

describe('다른 형식 (파일 엔진)', () => {
  test('docx 만들기·읽기·바꾸기·채우기', async () => {
    await ok('office_create', { output: 'memo.docx', markdown: MANUSCRIPT });
    let t = await ok('office_replace', { path: 'memo.docx', replacements: [{ find: '비용', replace: '지출' }], engine: 'file' });
    assert.match(t, /2곳 바꿈 \(엔진: file\)/);
    t = await ok('office_fields', { path: 'memo.docx', engine: 'file' });
    assert.match(t, /담당자 \[placeholder\]/);
    t = await ok('office_fill', { path: 'memo.docx', values: { 담당자: '최수진' }, engine: 'file' });
    assert.match(t, /채움 1개/);
    const read = await ok('office_read', { path: 'memo.docx' });
    assert.ok(read.includes('최수진') && read.includes('지출 절감'), read);
    assert.match(read, /표 1개/);
  });

  test('xlsx 셀 쓰기 + 잘못된 주소', async () => {
    await ok('office_create', { output: 'book.xlsx', markdown: MANUSCRIPT });
    const t = await ok('office_set_cells', { path: 'book.xlsx', cells: { D1: 10, D2: '=D1*3' }, sheet: 'Sheet1', engine: 'file' });
    assert.match(t, /셀 2개/);
    const read = await ok('office_read', { path: 'book.xlsx' });
    assert.match(read, /시트 1개: Sheet1/);
    assert.ok(read.includes('=D1*3'));
    await fails('office_set_cells', { path: 'book.xlsx', cells: { '가1': 1 }, engine: 'file' }, /셀 주소가 올바르지 않습니다/);
  });

  test('office_read: 긴 문서는 offset으로 나눠 읽고 잘림 없음', async () => {
    const lines = Array.from({ length: 2000 }, (_, i) => `${i + 1}번째 문단입니다. 긴 문서 나눠 읽기 확인용 글자.`);
    await ok('office_create', { output: 'long.docx', markdown: lines.join('\n') });
    const first = await ok('office_read', { path: 'long.docx', max_chars: 20000 });
    const total = Number(first.match(/전체 (\d+)자/)[1]);
    assert.ok(total > 60000, `total=${total}`);
    let offset = 0;
    let body = '';
    for (let guard = 0; guard < 10; guard++) {
      const t = guard === 0 ? first : await ok('office_read', { path: 'long.docx', offset, max_chars: 20000 });
      body += t.split('[본문]\n')[1].split('\n\n[계속]')[0];
      const next = t.match(/다음: offset=(\d+)/);
      if (!next) { assert.match(t, /끝까지 읽음/); break; }
      offset = Number(next[1]);
    }
    assert.equal(body.length, total);
    assert.ok(body.includes('2000번째 문단입니다'));
    await fails('office_read', { path: 'long.docx', offset: total + 10 }, /offset/);
  });
});

describe('저장 경로 제한 (AI가 아무 데나 쓰지 못하게)', () => {
  test('문서가 아닌 확장자는 거부하고 폴더도 만들지 않음', async () => {
    await ok('office_create', { output: 'guard.docx', markdown: MANUSCRIPT });
    await fails('office_fill', { path: 'guard.docx', values: { 없는칸: 'x' }, output: '../evil.cmd', engine: 'file' }, /\.cmd 파일로는 저장할 수 없습니다/);
    assert.ok(!existsSync(join(tmp.dir, '..', 'evil.cmd')));
    await fails('office_replace', { path: 'guard.docx', replacements: [{ find: '매출', replace: 'x' }], output: 'deep/new/x.ps1', engine: 'file' }, /\.ps1 파일로는/);
    assert.ok(!existsSync(tmp.path('deep')), '폴더가 만들어짐');
    await fails('office_convert', { path: 'guard.docx', output: 'run.bat' }, /\.bat 파일로는/);
    await fails('office_create', { output: 'x.html', markdown: '# x' }, /\.html 파일로는/);
    await fails('hwp_edit', { path: 'report.hwp', operations: [{ op: 'add_page_break' }], output: 'x.cmd' }, /\.hwp 또는 \.hwpx/);
  });

  test('이미 있는 다른 파일은 overwrite=true일 때만 덮어씀', async () => {
    writeFileSync(tmp.path('victim.docx'), '사용자의 다른 파일');
    await fails('office_fill', { path: 'guard.docx', values: { 담당자: 'x' }, output: 'victim.docx', engine: 'file' }, /이미 있는 파일.*overwrite=true/);
    assert.equal(readFileSync(tmp.path('victim.docx'), 'utf8'), '사용자의 다른 파일');
    await ok('office_fill', { path: 'guard.docx', values: { 담당자: 'x' }, output: 'victim.docx', overwrite: true, engine: 'file' });
    await fails('office_convert', { path: 'report.hwp', output: 'filled.hwpx' }, /이미 있는 파일/);
    await ok('office_convert', { path: 'report.hwp', output: 'filled.hwpx', overwrite: true });
    // 원본 자리에 그대로 쓰는 것은 허용
    await ok('office_replace', { path: 'guard.docx', replacements: [{ find: '매출', replace: '수익' }], engine: 'file' });
  });

  test('변환할 수 없는 형식은 프로그램을 띄우기 전에 한국어로 알림', async () => {
    await fails('office_convert', { path: 'guard.docx', output: 'guard.hwpx' }, /Word 문서는 \.hwpx 형식으로 바꿀 수 없습니다/);
    const { tools } = await client.listTools();
    assert.match(tools.find((t) => t.name === 'office_convert').description, /한글\(hwp\/hwpx\)로 바꿀 수 없다/);
  });

  test('OFFICEKIT_WRITE_ROOTS: 허락한 폴더 안에만 저장', async () => {
    const root = tmp.path('allowed');
    const { client: c } = await connect([SERVER], { OFFICEKIT_WRITE_ROOTS: root });
    const run = (name, args) => c.callTool({ name, arguments: args }, undefined, { timeout: 60_000 });
    try {
      let r = await run('office_create', { output: 'allowed/a.docx', markdown: MANUSCRIPT });
      assert.ok(!r.isError, textOf(r));
      r = await run('office_replace', { path: 'allowed/a.docx', replacements: [{ find: '매출', replace: '수익' }], engine: 'file' });
      assert.ok(!r.isError, textOf(r));
      for (const [name, args] of [
        ['office_replace', { path: 'allowed/a.docx', replacements: [{ find: '수익', replace: 'x' }], output: 'outside.docx', engine: 'file' }],
        ['office_replace', { path: 'allowed/a.docx', replacements: [{ find: '수익', replace: 'x' }], output: 'allowed/../outside.docx', engine: 'file' }],
        ['office_replace', { path: 'guard.docx', replacements: [{ find: '수익', replace: 'x' }], engine: 'file' }],
        ['office_create', { output: 'outside.hwp', markdown: '# x' }],
        ['office_convert', { path: 'allowed/a.docx', output: 'outside.pdf' }],
        ['hwp_edit', { path: 'report.hwp', operations: [{ op: 'add_page_break' }] }],
      ]) {
        r = await run(name, args);
        assert.equal(r.isError, true, `${name} ${JSON.stringify(args)} 가 실패해야 함`);
        assert.match(textOf(r), /폴더 안에만 저장할 수 있습니다/);
      }
      assert.ok(!existsSync(tmp.path('outside.docx')) && !existsSync(tmp.path('outside.hwp')));
      r = await run('office_read', { path: 'guard.docx' });
      assert.ok(!r.isError, '읽기는 막지 않음');
    } finally { await c.close(); }
  });
});

describe('개인정보 · HWPX 일괄 변환', () => {
  test('office_find_pii → office_mask_pii: 가린 사본, 원본 그대로', async () => {
    await ok('office_create', { output: 'pii.docx', markdown: '# 명단\n\n홍길동 010-1234-5678, 900101-1234567\n\n메일 hong@example.com' });
    const found = await ok('office_find_pii', { path: 'pii.docx' });
    assert.match(found, /3가지/);
    assert.match(found, /900101-1\*{6}/);
    const masked = await ok('office_mask_pii', { path: 'pii.docx', types: ['rrn', 'mobile'] });
    assert.match(masked, /pii_가림\.docx/);
    assert.match(masked, /남지 않은 것을 확인/);
    const read = await ok('office_read', { path: 'pii_가림.docx' });
    assert.ok(read.includes('010-****-5678') && read.includes('hong@example.com') && !read.includes('900101-1234567'), read);
    assert.match(await ok('office_read', { path: 'pii.docx' }), /900101-1234567/);
    await fails('office_mask_pii', { path: 'pii.docx', output: 'pii.docx' }, /덮어쓰|이미 있는 파일/);
  });

  test('office_hwpx_batch: 폴더의 hwp → hwpx + 보고서', async () => {
    mkdirSync(tmp.path('batch/sub'), { recursive: true });
    copyFileSync(join(ROOT, 'examples', '신청서.hwp'), tmp.path('batch/a.hwp'));
    copyFileSync(join(ROOT, 'examples', '신청서_채움.hwp'), tmp.path('batch/sub/b.hwp'));
    const t = await ok('office_hwpx_batch', { paths: ['batch'], report: 'batch/결과.csv' });
    assert.match(t, /2개: 같음 2/);
    assert.ok(existsSync(tmp.path('batch/a.hwpx')) && existsSync(tmp.path('batch/sub/b.hwpx')));
    assert.match(readFileSync(tmp.path('batch/결과.csv'), 'utf8'), /같음/);
    assert.match(await ok('office_hwpx_batch', { paths: ['batch/a.hwp'] }), /건너뜀 1/);
  });
});

describe('hwp_edit 세부', () => {
  test('{{ 이름 }} 띄어 쓴 자리표시자 채우기, 대소문자 구분 바꾸기', async () => {
    await ok('office_create', { output: 'space.hwpx', markdown: ['# 제목', '신청인 {{ 이름 }} 입니다', 'AI said the main idea.'].join('\n') });
    const t = await ok('hwp_edit', { path: 'space.hwpx', operations: [
      { op: 'fill', values: { 이름: '최수진' } },
      { op: 'replace', find: 'AI', replace: '인공지능' },
    ] });
    assert.match(t, /채우기: 이름/);
    assert.match(t, /'AI' → '인공지능': 1곳/);
    const read = await ok('office_read', { path: 'space.hwpx' });
    assert.ok(read.includes('신청인 최수진 입니다') && read.includes('인공지능 said the main idea.'), read);
  });
});

describe('오류 처리', () => {
  test('없는 파일', async () => {
    for (const [name, args] of [
      ['office_read', { path: 'nope.hwp' }],
      ['office_info', { path: 'nope.docx' }],
      ['office_fill', { path: 'nope.hwp', values: { a: 'b' } }],
      ['office_page_image', { path: 'nope.hwp' }],
      ['hwp_edit', { path: 'nope.hwp', operations: [{ op: 'add_page_break' }] }],
    ]) await fails(name, args, /파일이 없습니다/);
  });

  test('필수 인수 누락은 한국어 검증 오류', async () => {
    const t = await fails('office_read', {}, /입력값이 올바르지 않습니다/);
    assert.match(t, /'path' 값이 필요합니다/);
  });

  test("검증 오류의 영어 경로 꼬리(' at 경로')는 '(위치: 경로)'로", async () => {
    let t = await fails('hwp_edit', { path: 'report.hwp', operations: [{ op: 'explode' }] }, /알 수 없는 작업\(op\)/);
    assert.match(t, /\(위치: operations\[0\]\.op\)$/);
    assert.doesNotMatch(t, / at operations/);
    t = await fails('office_page_image', { path: 'report.hwp', page: 0 }, /값이 너무 작습니다 \(1 이상\) \(위치: page\)$/);
    assert.doesNotMatch(t, /Too small| at page/);
    // 문제가 여럿이면 줄마다
    const { koreanError } = await import('../src/mcp/server.mjs');
    const msg = koreanError({ message: 'MCP error -32602: Input validation error: Invalid arguments for tool ppt_edit: '
      + '잘못된 입력 at operations[0].slide\nInvalid option: expected one of "a"|"b" at operations[1].effect\n'
      + 'Too small: expected string to have >=1 characters at values.칸 이름\nUnrecognized key: "extra" at operations[2]\n경로 없는 문제' });
    assert.equal(msg, '입력값이 올바르지 않습니다 (ppt_edit): 잘못된 입력 (위치: operations[0].slide)\n'
      + '다음 중 하나여야 합니다: "a", "b" (위치: operations[1].effect)\n너무 짧습니다 (1자 이상) (위치: values.칸 이름)\n'
      + '알 수 없는 인수: "extra" (위치: operations[2])\n경로 없는 문제');
  });

  test('stdout은 프로토콜 전용 (서버 로그는 stderr)', () => {
    assert.ok(serverLog.join('').includes('MCP 서버 시작'));
  });
});

describe('office-kit mcp (명령줄로 시작)', () => {
  test('같은 도구 목록', async () => {
    const { client: c2 } = await connect([CLI, 'mcp']);
    try {
      const { tools } = await c2.listTools();
      assert.deepEqual(tools.map((t) => t.name).sort(), [...TOOLS].sort());
      const r = await c2.callTool({ name: 'office_read', arguments: { path: 'report.hwp' } });
      assert.ok(!r.isError && textOf(r).includes('분기 보고'));
    } finally { await c2.close(); }
  });
});

describe('프로그램 엔진 (Word) — OFFICEKIT_COM_TESTS=1 일 때만', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  test('docx 기본(auto) 바꾸기·쪽 그림', { timeout: 300_000 }, async () => {
    await ok('office_create', { output: 'com.docx', markdown: MANUSCRIPT });
    const t = await ok('office_replace', { path: 'com.docx', replacements: [{ find: '매출', replace: '수익' }] });
    assert.match(t, /엔진: word/);
    const r = await call('office_page_image', { path: 'com.docx', page: 1 });
    assert.ok(!r.isError, textOf(r));
    assert.match(textOf(r), /word\+pdf/);
    const pdf = await ok('office_convert', { path: 'com.docx', output: 'com.pdf' });
    assert.match(pdf, /변환 완료/);
    assert.equal(readFileSync(tmp.path('com.pdf')).subarray(0, 4).toString(), '%PDF');
  });
});
