// hwp-pdf.test.mjs — 한컴오피스 없이 한글(hwp·hwpx·hwt) → PDF (vendor/rhwp-cli/rhwp.exe, src/core/hwppdf.mjs).
//   core·명령줄·MCP 세 길 모두. 쪽 수는 rhwp 화면용 쪽 수와 같아야 하고, PDF 안 글자(ToUnicode)로 한국어를 찾을 수 있어야 한다.
//   rhwp.exe가 없으면(npm run rhwp-cli:build 전) 변환 시험은 건너뛴다 — CI에서는 먼저 만들므로 없으면 실패.
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { isAbsolute, join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { convertDocument, createDocument, Hwp, hwpPdfStatus } from '../src/core/index.mjs';
import { hwpToPdf, RHWP_CLI_DEFAULT, fallbackFontArgs } from '../src/core/hwppdf.mjs';
import { comAvailableFor } from '../src/core/com.mjs';
import { pageImage } from '../src/core/render.mjs';
import { MANUSCRIPT, ROOT, tempDir, isPng } from './helpers/fixtures.mjs';
import { readPdf, squash } from './helpers/pdftext.mjs';

const require = createRequire(import.meta.url);
const HAVE_CLI = existsSync(RHWP_CLI_DEFAULT);
const SKIP = !HAVE_CLI && 'vendor/rhwp-cli/rhwp.exe 없음 (npm run rhwp-cli:build)';
const tmp = tempDir('officekit-hwppdf-test-');
// 한국어 글꼴이 없는 PC(예: GitHub의 Windows Server 빌드 기계)에서는 PDF에 한글 글자가 들어가지 않는다 → 한글 글자 확인만 건너뛴다
const FONT_DIR = join(process.env.SystemRoot ?? process.env.windir ?? 'C:\Windows', 'Fonts');
const HAS_KO_FONT = ['malgun.ttf', 'batang.ttc', 'gulim.ttc', 'HANBatang.ttf', 'HMKMRHD.TTF'].some((f) => existsSync(join(FONT_DIR, f)));
const koOk = (cond, msg) => { if (HAS_KO_FONT) assert.ok(cond, msg); };
after(() => tmp.cleanup());

const sha = (p) => createHash('sha256').update(readFileSync(p)).digest('hex');
const pagesOf = (p) => { const h = Hwp.openForView(p); try { return h.pageCount; } finally { h.close(); } };
const ICON = join(ROOT, 'build', 'icon.png');

// 여러 쪽 시험 문서: 표·머리말·꼬리말(쪽 번호)·그림·한국어
async function makeMulti(out) {
  let md = '# 사업 결과 보고서\n\n>> 정보화담당관실\n\n';
  for (let s = 1; s <= 4; s++) {
    md += `## ${s}. 추진 과제 ${s}\n\n`;
    for (let p = 0; p < 6; p++) md += `대한민국 국민은 법 앞에 평등하다 시험 문단 ${s}-${p}. 가나다라마바사 아자차카타파하 한글 문서를 PDF로 바꾸는 시험입니다.\n\n`;
    md += '| 구분 | 계획 | 실적 |\n|---|---|---|\n';
    for (let r = 1; r <= 4; r++) md += `| 항목${s}-${r} | ${r * 100} | ${r * 95} |\n`;
    md += '\n';
  }
  await createDocument(out, md);
  const h = Hwp.open(out);
  try {
    h.setHeaderFooter('header', '대외비 머리말');
    h.setHeaderFooter('footer', '', { pageNumbers: true });
    h.insertImage(ICON, { at: { after_paragraph: 2 }, widthMm: 40 });
    h.save(out);
  } finally { h.close(); }
}

describe('vendor/rhwp-cli 묶음 확인 (beforePack)', () => {
  const { missingRhwpCliParts, RHWP_CLI_FILES } = require('../scripts/check-editor-bundle.cjs');
  test('파일·VERSION.json 크기를 본다', () => {
    const root = tmp.path('proj');
    const dir = join(root, 'vendor', 'rhwp-cli');
    assert.deepEqual(missingRhwpCliParts(root), RHWP_CLI_FILES.map((f) => `vendor/rhwp-cli/${f}`));
    mkdirSync(dir, { recursive: true });
    const files = [];
    for (const f of RHWP_CLI_FILES.filter((f) => f !== 'VERSION.json')) {
      writeFileSync(join(dir, f), `내용 ${f}`);
      files.push({ path: f, bytes: readFileSync(join(dir, f)).length });
    }
    writeFileSync(join(dir, 'VERSION.json'), JSON.stringify({ name: 'rhwp-cli', files }));
    assert.deepEqual(missingRhwpCliParts(root), []);
    writeFileSync(join(dir, 'rhwp.exe'), 'x');
    assert.deepEqual(missingRhwpCliParts(root), ['vendor/rhwp-cli/rhwp.exe (크기가 다름)']);
    writeFileSync(join(dir, 'VERSION.json'), '{}');
    assert.match(missingRhwpCliParts(root)[0], /모양이 다름/);
  });
  test('만들어 둔 vendor/rhwp-cli는 통과한다', { skip: SKIP }, () => {
    assert.deepEqual(missingRhwpCliParts(ROOT), []);
  });
  test('빌드 스크립트 --help · --verify', { skip: SKIP }, async () => {
    const run = (args) => new Promise((done) => execFile(process.execPath, [join(ROOT, 'scripts', 'build-rhwp-cli.mjs'), ...args],
      { encoding: 'utf8', timeout: 120_000, windowsHide: true }, (err, stdout, stderr) => done({ code: err ? err.code ?? 1 : 0, stdout, stderr })));
    let r = await run(['--help']);
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /rhwp v0\.8\.7/);
    r = await run(['--verify']);
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /확인 완료/);
    r = await run(['--nope']);
    assert.equal(r.code, 1);
    assert.match(r.stderr, /모르는 옵션/);
  });
});

describe('한글 → PDF (core, 한컴오피스 없이)', { skip: SKIP }, () => {
  before(async () => {
    if (process.env.CI) assert.ok(HAVE_CLI, 'CI에서는 npm run rhwp-cli:build로 rhwp.exe를 먼저 만들어야 합니다');
    await createDocument(tmp.path('보고.hwp'), MANUSCRIPT);
    await createDocument(tmp.path('보고.hwpx'), MANUSCRIPT);
    await makeMulti(tmp.path('여러 쪽.hwp'));
    await createDocument(tmp.path('가로.hwp'), '# 가로 문서\n\n가로 방향 A4 문서입니다.\n\n| 가 | 나 |\n|---|---|\n| 1 | 2 |\n');
    const h = Hwp.open(tmp.path('가로.hwp'));
    try { h.pageSetup({ orientation: 'landscape' }); h.save(tmp.path('가로.hwp')); } finally { h.close(); }
  });

  const CASES = [
    ['보고.hwp', ['분기보고', '2026년3분기', '매출증가', '신규고객확보']],
    ['보고.hwpx', ['분기보고', '비용절감', '다음계획']],
    ['여러 쪽.hwp', ['대외비머리말', '대한민국국민은법앞에평등하다', '항목4-4', '추진과제4']],
    [join(ROOT, 'examples', '신청서.hwp'), ['장비대여신청서', '위와같이장비대여를신청합니다']],
    [join(ROOT, 'examples', '신청서.hwpx'), ['장비대여신청서', '반납이늦어지면']],
    // 누름틀에 채운 값은 PDF에 나온다 (빈 누름틀의 안내문은 한컴 인쇄처럼 나오지 않음)
    [join(ROOT, 'examples', '신청서_채움.hwpx'), ['성명:홍길동', '학번:20261234', '010-0000-0000']],
  ];
  for (const [name, words] of CASES) {
    test(`${name.split(/[\\/]/).pop()}: 쪽 수·한국어 글자·글꼴`, async () => {
      const src = isAbsolute(name) ? name : tmp.path(name);
      const out = tmp.path(`${name.split(/[\\/]/).pop()}.pdf`);
      const before = sha(src);
      const r = await hwpToPdf(src, out);
      assert.equal(r.engine, 'rhwp');
      assert.equal(sha(src), before, '원본은 그대로');
      const pdf = readPdf(readFileSync(out));
      const want = pagesOf(src);
      assert.equal(pdf.pages.length, want, '쪽 수 = rhwp 쪽 수');
      assert.equal(r.pages, want);
      const text = squash(pdf.text);
      for (const w of words) koOk(text.includes(w), `PDF 글자에 '${w}' 없음: ${text.slice(0, 200)} · 글꼴: ${pdf.fonts.join(',') || '없음'} · 대체: ${fallbackFontArgs().join(' ')}`);
      // 글꼴은 쓴 글자만 넣은 부분 글꼴(이름 앞 'ABCDEF+')로 들어 있다
      assert.ok(pdf.fonts.length > 0 && pdf.fonts.every((f) => /^[A-Z]{6}\+/.test(f)), pdf.fonts.join(','));
      for (const p of pdf.pages) assert.ok(Math.abs(p.width - 595.28) < 1 && Math.abs(p.height - 841.86) < 1, `A4 세로: ${p.width}×${p.height}`);
    });
  }

  test('여러 쪽 문서: 쪽 번호 꼬리말이 쪽마다', async () => {
    const pdf = readPdf(readFileSync(tmp.path('여러 쪽.hwp.pdf')));
    assert.ok(pdf.pages.length >= 2, `${pdf.pages.length}쪽`);
    pdf.pages.forEach((p, i) => koOk(squash(p.text).includes(`-${i + 1}-`), `${i + 1}쪽 꼬리말: ${squash(p.text).slice(-20)}`));
  });

  test('가로 문서는 가로 쪽', async () => {
    const out = tmp.path('가로.pdf');
    await hwpToPdf(tmp.path('가로.hwp'), out);
    const [p] = readPdf(readFileSync(out)).pages;
    assert.ok(Math.abs(p.width - 841.86) < 1 && Math.abs(p.height - 595.28) < 1, `${p.width}×${p.height}`);
  });

  test('.hwt(서식 파일)도 된다', async () => {
    const hwt = tmp.path('서식.hwt');
    writeFileSync(hwt, readFileSync(tmp.path('보고.hwp')));
    const out = tmp.path('서식.pdf');
    await hwpToPdf(hwt, out);
    koOk(squash(readPdf(readFileSync(out)).text).includes('분기보고'));
  });

  test('convertDocument: 한컴이 없으면 내장 엔진 + 안내, 공백·한글 경로, 덮어쓰기', async (t) => {
    if (await comAvailableFor('hwp')) return t.skip('한컴오피스 + 보안 모듈이 있으면 한컴으로 변환한다 (한컴은 시험에서 쓰지 않음)');
    const dir = tmp.path('결과 폴더');
    mkdirSync(dir, { recursive: true });
    const out = join(dir, '보고 (최종).pdf');
    writeFileSync(out, '옛 파일');
    const r = await convertDocument(tmp.path('보고.hwpx'), out);
    assert.equal(r.engine, 'rhwp');
    assert.equal(r.output, out);
    assert.equal(r.pages, pagesOf(tmp.path('보고.hwpx')));
    assert.match(r.note, /한컴오피스 없이 내장 엔진\(rhwp\)/);
    assert.equal(readFileSync(out).subarray(0, 5).toString('latin1'), '%PDF-');
    assert.deepEqual(readdirSync(dir), ['보고 (최종).pdf'], '임시 파일이 남지 않음');
  });

  test('만든 PDF를 그림으로 (Windows.Data.Pdf)', { skip: process.platform !== 'win32' && 'Windows 전용' }, async () => {
    const img = await pageImage(tmp.path('여러 쪽.hwp.pdf'), 2, { dpi: 48 });
    assert.ok(isPng(img.png));
    assert.equal(img.pages, pagesOf(tmp.path('여러 쪽.hwp')));
    assert.ok(img.height > img.width);
  });

  test('시간 제한을 넘기면 멈추고 결과를 만들지 않는다', async () => {
    const out = tmp.path('늦음.pdf');
    const before = readdirSync(tmpdir()).filter((f) => f.startsWith('officekit-hwppdf-')).length;
    await assert.rejects(hwpToPdf(tmp.path('여러 쪽.hwp'), out, { timeoutMs: 1 }), /시간이 너무 오래 걸려 멈췄습니다/);
    assert.ok(!existsSync(out));
    const left = readdirSync(tmpdir()).filter((f) => f.startsWith('officekit-hwppdf-')).length;
    assert.ok(left <= before, '임시 폴더를 지움');
  });

  test('프로그램이 없거나 실패하면 한국어 오류', async () => {
    const saved = process.env.OFFICEKIT_RHWP_CLI;
    try {
      process.env.OFFICEKIT_RHWP_CLI = tmp.path('없음\\rhwp.exe');
      assert.equal(hwpPdfStatus().ok, false);
      await assert.rejects(convertDocument(tmp.path('보고.hwp'), tmp.path('x.pdf')), /내장 PDF 변환 프로그램이 없습니다/);
      // 다른 프로그램(node)을 rhwp 자리에 두면 종료 코드가 0이 아니다 → 변환 실패 문구
      process.env.OFFICEKIT_RHWP_CLI = process.execPath;
      await assert.rejects(hwpToPdf(tmp.path('보고.hwp'), tmp.path('x.pdf')), /한글 문서를 PDF로 바꾸지 못했습니다 \(내장 엔진 rhwp\)/);
      assert.ok(!existsSync(tmp.path('x.pdf')));
    } finally {
      if (saved === undefined) delete process.env.OFFICEKIT_RHWP_CLI; else process.env.OFFICEKIT_RHWP_CLI = saved;
    }
  });

  test('깨진 파일·없는 파일', async () => {
    const bad = tmp.path('깨짐.hwp');
    writeFileSync(bad, Buffer.alloc(4096, 7));
    await assert.rejects(hwpToPdf(bad, tmp.path('깨짐.pdf')), /한글 문서를 열지 못해 PDF로 바꿀 수 없습니다/);
    await assert.rejects(convertDocument(tmp.path('없음.hwp'), tmp.path('없음.pdf')), /파일이 없습니다/);
  });
});

// ── 명령줄 · MCP ─────────────────────────────────────────

function cli(args, env = {}) {
  return new Promise((done) => {
    execFile(process.execPath, [join(ROOT, 'src', 'cli.mjs'), ...args], { cwd: tmp.dir, env: { ...process.env, ...env }, encoding: 'utf8', timeout: 180_000, windowsHide: true },
      (err, stdout, stderr) => done({ code: err ? err.code ?? 1 : 0, stdout, stderr }));
  });
}

describe('명령줄 office-kit convert x.hwp x.pdf', { skip: SKIP }, () => {
  before(async () => { await createDocument(tmp.path('cli.hwp'), MANUSCRIPT); });

  test('내장 엔진으로 변환 + 안내 (PowerShell 없이도)', async (t) => {
    if (await comAvailableFor('hwp')) return t.skip('한컴오피스 + 보안 모듈이 있으면 한컴으로 변환한다');
    for (const env of [{}, { OFFICEKIT_POWERSHELL: 'C:\\없는폴더\\powershell.exe' }]) {
      const out = join('출력 폴더', `cli-${Object.keys(env).length}.pdf`);
      const r = await cli(['convert', 'cli.hwp', out], env);
      assert.equal(r.code, 0, r.stderr);
      assert.match(r.stdout, /변환 완료 \(내장 엔진\(rhwp\), 2쪽\)/);
      assert.match(r.stdout, /참고: 한컴오피스 없이/);
      koOk(squash(readPdf(readFileSync(tmp.path(out))).text).includes('분기보고'));
    }
  });

  test('status에 한글→PDF 엔진', async () => {
    const r = await cli(['status']);
    assert.equal(r.code, 0, r.stderr);
    assert.match(r.stdout, /한글→PDF: (한컴오피스|내장 엔진\(rhwp)/);
  });

  test('help에 한글→pdf 안내', async () => {
    const r = await cli(['help']);
    assert.match(r.stdout, /한글→pdf는 한컴오피스\(\+자동화 연결\)가 있으면 한컴으로, 없으면 내장 엔진\(rhwp\)으로/);
  });
});

describe('MCP office_convert (ELECTRON_RUN_AS_NODE와 같은 Node 실행)', { skip: SKIP }, () => {
  let client;
  before(async () => {
    await createDocument(tmp.path('mcp.hwpx'), MANUSCRIPT);
    client = new Client({ name: 'office-kit-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [join(ROOT, 'src', 'mcp', 'server.mjs')], cwd: tmp.dir, stderr: 'pipe' }));
  });
  after(async () => { await client?.close(); });
  const call = (name, args = {}) => client.callTool({ name, arguments: args }, undefined, { timeout: 180_000 });
  const textOf = (r) => r.content.filter((c) => c.type === 'text').map((c) => c.text).join('\n');

  test('hwpx → pdf', async (t) => {
    if (await comAvailableFor('hwp')) return t.skip('한컴오피스 + 보안 모듈이 있으면 한컴으로 변환한다');
    const r = await call('office_convert', { path: 'mcp.hwpx', output: 'mcp/결과.pdf' });
    assert.ok(!r.isError, textOf(r));
    assert.match(textOf(r), /변환 완료 \(엔진: 내장\(rhwp\), 2쪽\)/);
    assert.match(textOf(r), /참고: 한컴오피스 없이/);
    const pdf = readPdf(readFileSync(tmp.path('mcp/결과.pdf')));
    assert.equal(pdf.pages.length, pagesOf(tmp.path('mcp.hwpx')));
    koOk(squash(pdf.text).includes('신규고객확보'));
    // 이미 있으면 overwrite 없이는 거부
    const again = await call('office_convert', { path: 'mcp.hwpx', output: 'mcp/결과.pdf' });
    assert.equal(again.isError, true);
    assert.match(textOf(again), /이미 있는 파일/);
  });

  test('office_status·도구 설명', async () => {
    assert.match(textOf(await call('office_status')), /한글→PDF: (한컴오피스로|내장 엔진\(rhwp\)으로)/);
    const { tools } = await client.listTools();
    assert.match(tools.find((x) => x.name === 'office_convert').description, /한글→pdf: 한컴오피스 \+ 자동화 연결이 있으면 한컴으로.*없으면 내장 엔진\(rhwp/);
  });
});

// 실제 앱의 'PDF로 저장' (저장 대화상자는 자동화할 수 없어 OFFICEKIT_SMOKE_PDF로 경로를 미리 정함). 창이 잠깐 뜨므로 OFFICEKIT_EDITOR_SMOKE=1일 때만
const wantSmoke = process.env.OFFICEKIT_EDITOR_SMOKE === '1';
const electronExe = !wantSmoke ? '' : (() => { try { return require('electron'); } catch { return ''; } })();
test('실제 앱: PDF로 저장 단추 → 내장 엔진 PDF + 알림', { skip: (SKIP || !(wantSmoke && electronExe && existsSync(electronExe))) && 'OFFICEKIT_EDITOR_SMOKE=1 과 rhwp.exe가 있을 때만 (창이 잠깐 뜬다)' }, async (t) => {
  if (await comAvailableFor('hwp')) return t.skip('한컴오피스 + 보안 모듈이 있으면 한컴으로 변환한다');
  const doc = tmp.path('앱 시험.hwp');
  await makeMulti(doc);
  const out = tmp.path('앱 결과/앱 시험.pdf');
  mkdirSync(join(tmp.dir, '앱 결과'), { recursive: true });
  const env = { ...process.env, OFFICEKIT_SMOKE: tmp.path('app.png'), OFFICEKIT_SMOKE_PDF: out, OFFICEKIT_USER_DATA: tmp.path('ud') };
  delete env.ELECTRON_RUN_AS_NODE;
  for (const k of ['OFFICEKIT_SMOKE_EDITOR', 'OFFICEKIT_AI_MODULE', 'OFFICEKIT_RHWP_CLI']) delete env[k];
  const r = await new Promise((done) => execFile(electronExe, ['.', doc], { cwd: ROOT, env, encoding: 'utf8', timeout: 240_000, windowsHide: true },
    (err, stdout, stderr) => done({ code: err ? err.code ?? 1 : 0, stdout, stderr })));
  const line = r.stdout.split(/\r?\n/).find((l) => l.startsWith('[smoke] '));
  assert.ok(line, `${r.stdout}\n${r.stderr}`);
  const info = JSON.parse(line.slice('[smoke] '.length));
  assert.equal(r.code, 0, line);
  assert.equal(info.pdf?.ok, true, line);
  assert.match(info.pdf.toast, /PDF로 저장했어요 \(내장 엔진\): 앱 시험\.pdf/);
  const pdf = readPdf(readFileSync(out));
  assert.equal(pdf.pages.length, pagesOf(doc));
  koOk(squash(pdf.text).includes('대외비머리말'));
});

// 시험이 끝난 뒤 rhwp.exe가 남아 있지 않은지 (시간 제한으로 끊은 것 포함)
test('rhwp.exe 프로세스가 남지 않는다', { skip: (SKIP || process.platform !== 'win32') && '해당 없음' }, async () => {
  const out = await new Promise((done) => execFile('tasklist', ['/FI', 'IMAGENAME eq rhwp.exe', '/FO', 'CSV', '/NH'], { encoding: 'utf8', windowsHide: true },
    (err, stdout) => done(String(stdout ?? ''))));
  // 다른 곳(사용자)이 띄운 rhwp.exe가 있을 수 있으므로 우리 vendor 경로 것만 보면 좋겠지만 tasklist는 경로를 안 준다 → 수만 확인
  assert.ok(!/"rhwp\.exe"/i.test(out), `남은 rhwp.exe: ${out.trim()}`);
});
