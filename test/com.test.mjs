// com.test.mjs — core/com.mjs: 조사, 오류 문구, 공유 Office 잠금(Local\OfficeKitCom)과 시간 제한, 쓰기 암호 문서.
// 잠금 시험은 Windows PowerShell만 쓴다 (Office를 띄우지 않음). 쓰기 암호·시간 초과 시험은 OFFICEKIT_COM_TESTS=1 일 때만 (한컴 COM은 쓰지 않음).
import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { execFile, spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { josa, koreanComError, comCall, LOCK_WAIT_MS } from '../src/core/com.mjs';
import { convertDocument, createDocument, readDocument, replaceText } from '../src/core/index.mjs';
import { COM_TESTS, MANUSCRIPT, tempDir } from './helpers/fixtures.mjs';

const WIN = process.platform === 'win32';
const POWERSHELL = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const tmp = tempDir('officekit-com-test-');
after(() => tmp.cleanup());
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── 조사 ─────────────────────────────────────────────────

describe('조사 (josa)', () => {
  test('프로그램 이름은 읽는 소리대로', () => {
    assert.equal(josa('Word'), 'Word가');
    assert.equal(josa('Excel'), 'Excel이');
    assert.equal(josa('PowerPoint'), 'PowerPoint가');
    assert.equal(josa('한글'), '한글이');
    assert.equal(josa('PDF 변환'), 'PDF 변환이');
    assert.equal(josa('Office Kit'), 'Office Kit가');
  });
  test('받침·ㄹ 받침·숫자·약어', () => {
    assert.equal(josa('문서', '을/를'), '문서를');
    assert.equal(josa('표', '은/는'), '표는');
    assert.equal(josa('파일', '으로/로'), '파일로');
    assert.equal(josa('문단', '으로/로'), '문단으로');
    assert.equal(josa('슬라이드 3', '이/가'), '슬라이드 3이');
    assert.equal(josa('슬라이드 2', '이/가'), '슬라이드 2가');
    assert.equal(josa('HTML'), 'HTML이');
    assert.equal(josa('CSV', '과/와'), 'CSV와');
  });
  test('모르면 둘 다', () => {
    assert.equal(josa('…'), '…이(가)');
    assert.equal(josa('…', '으로/로'), '…(으)로');
  });
});

// ── 오류 문구 ─────────────────────────────────────────────

describe('COM 오류 → 한국어', () => {
  const IN_USE = '문서가 다른 프로그램에서 열려 있어 작업할 수 없습니다';
  test('다른 프로그램이 쥐고 있는 파일 (Office가 쓰는 여러 표현)', () => {
    for (const msg of [
      'This file is in use by another application or user. (C:\\a\\b.docx)',
      'The process cannot access the file because it is being used by another process.',
      "'b.xlsx' is locked for editing by '김선호'.",
      'save failed: C:\\a\\b.pptx (The file is in use by another process.)',
      'open failed: C:\\a\\b.docx (Word cannot save this file because it is already open elsewhere.)',
      'Your changes could not be saved because of a sharing violation.',
      'locked for editing: C:\\a\\b.xlsx',
    ]) assert.equal(koreanComError(msg, 'word'), IN_USE, msg);
  });
  test("편집 제한 문서의 'locked for editing'은 다른 프로그램 탓이 아니다", () => {
    assert.equal(koreanComError('This method or property is not available because the document is locked for editing.', 'word'),
      '편집이 제한된(보호된) 문서라 고칠 수 없습니다');
  });
  test('쓰기 암호 문서를 그 자리에 저장하려 함', () => {
    assert.match(koreanComError('write protected: C:\\a\\보고서.docx', 'word'), /^쓰기 암호가 걸린 문서라 그 자리에 저장할 수 없습니다: 보고서\.docx — output/);
  });
  test('작업별 오류는 형식 모듈이 풀도록 그대로 넘긴다 (번호를 잃지 않음)', () => {
    assert.equal(koreanComError('opfail|3|add_image|The process cannot access the file because it is being used by another process.', 'ppt'),
      'PowerPoint 작업 실패: opfail|3|add_image|The process cannot access the file because it is being used by another process.');
    assert.equal(koreanComError('op 2 (insert_image): password is incorrect', 'word'), 'Word 작업 실패: op 2 (insert_image): password is incorrect');
  });
  test('나머지는 전과 같이', () => {
    assert.equal(koreanComError('open failed: C:\\x.docx (The password is incorrect.)', 'word'), '암호가 걸린 문서는 열 수 없습니다');
    assert.equal(koreanComError('unsupported format: zzz', 'excel'), '이 프로그램으로는 .zzz 형식으로 저장할 수 없습니다');
    assert.equal(koreanComError('Something odd', 'excel'), 'Excel 작업 실패: Something odd');
    assert.equal(koreanComError('', 'ppt'), 'PowerPoint 작업 실패: 알 수 없는 오류');
  });
});

// ── 공유 Office 잠금 ──────────────────────────────────────

/**
 * 다른 Office Kit 프로세스(앱 미리보기·AI의 MCP 서버)처럼 공유 잠금을 잡고 ms 동안 놓지 않는다.
 * → { until: 놓는 시각, done: 끝날 때까지 } | null (다른 Office 작업이 잠금을 2분 넘게 쥐고 있음)
 */
async function holdLock(ms) {
  const ready = tmp.path(`held-${Date.now()}.txt`);
  const script = [
    "$m = New-Object System.Threading.Mutex($false, 'Local\\OfficeKitCom')",
    'try { $got = $m.WaitOne(120000) } catch [System.Threading.AbandonedMutexException] { $got = $true }',
    'if (-not $got) { exit 2 }',
    `[System.IO.File]::WriteAllText('${ready}', 'held')`,
    `Start-Sleep -Milliseconds ${ms}`,
    '$m.ReleaseMutex(); $m.Dispose()',
  ].join('; ');
  const ps = spawn(POWERSHELL, ['-NoProfile', '-NonInteractive', '-Command', script], { windowsHide: true, stdio: 'ignore' });
  const done = new Promise((r) => ps.on('exit', r));
  const t = Date.now();
  while (!existsSync(ready)) {
    if (ps.exitCode !== null || Date.now() - t > 130_000) { ps.kill(); return null; }
    await sleep(50);
  }
  return { until: Date.now() + ms, done };
}

// 잠금을 잡은 뒤 곧바로 끝나는 요청: 지원하지 않는 형식은 Word를 띄우기 전에 거절된다
const quickAfterLock = () => ({ app: 'word', op: 'convert', path: tmp.path('lock.docx'), save: tmp.path('lock.zzz'), saveFormat: 'zzz' });

describe('공유 Office 잠금과 시간 제한', { skip: !WIN && 'Windows만' }, () => {
  test('잠금 대기 상한 기본값은 15분', () => assert.equal(LOCK_WAIT_MS, 15 * 60_000));

  test('다른 프로세스의 Office 작업을 기다린 시간은 시간 제한에 넣지 않는다', { timeout: 300_000 }, async (t) => {
    const holder = await holdLock(6000);
    if (!holder) return t.skip('다른 Office 작업이 잠금을 오래 쥐고 있음');
    const start = Date.now();
    // 시간 제한 2.5초 < 잠금을 기다리는 약 6초: 예전에는 여기서 '응답하지 않아' 시간 초과였다
    await assert.rejects(comCall(quickAfterLock(), { timeoutMs: 2500 }), (e) => {
      assert.match(e.message, /\.zzz 형식으로 저장할 수 없습니다/);
      return true;
    });
    const waited = Date.now() - start;
    assert.ok(Date.now() >= holder.until - 300, `잠금이 풀리기 전에 끝남 (${waited}ms)`);
    await holder.done;
  });

  test('잠금을 너무 오래 기다리면 따로 알리고 멈춘다', { timeout: 300_000 }, async (t) => {
    const holder = await holdLock(10_000);
    if (!holder) return t.skip('다른 Office 작업이 잠금을 오래 쥐고 있음');
    const start = Date.now();
    await assert.rejects(comCall(quickAfterLock(), { timeoutMs: 60_000, lockTimeoutMs: 2000 }), (e) => {
      assert.match(e.message, /^다른 Office 작업이 끝나기를 기다리다 멈췄습니다 \(2초 넘게 기다림\)/);
      // word.mjs는 '응답하지 않아'를 보고 소유자 파일을 지운다: 기다리기만 한 요청에는 그 말을 쓰지 않는다
      assert.doesNotMatch(e.message, /응답하지 않아/);
      return true;
    });
    assert.ok(Date.now() - start < 8000, `${Date.now() - start}ms`);
    assert.ok(Date.now() < holder.until, '잠금이 풀린 뒤에야 멈춤');
    await holder.done;
    // 버려진 대기 뒤에도 다음 요청은 잠금을 잡는다
    await assert.rejects(comCall(quickAfterLock(), { timeoutMs: 20_000 }), /\.zzz 형식으로 저장할 수 없습니다/);
  });
});

// ── 설치된 프로그램 (COM) — OFFICEKIT_COM_TESTS=1 일 때만 ──────

const officeRunning = (image) => new Promise((done) => execFile('tasklist', ['/FI', `IMAGENAME eq ${image}`, '/NH'], { encoding: 'utf8' },
  (_e, out) => done(String(out).toUpperCase().includes(image))));

// 쓰기 암호('modify')를 건 사본을 프로그램으로 만든다. 경로는 꼭 [string]으로 (PSObject로 싼 경로를 받으면 Word SaveAs2가 멈춘다)
const PROTECT_PS = `param([string]$Dir)
$ErrorActionPreference = 'Stop'
$m = New-Object System.Threading.Mutex($false, 'Local\\OfficeKitCom')
try { [void]$m.WaitOne() } catch [System.Threading.AbandonedMutexException] { }
try {
  [string]$src = Join-Path $Dir 'plain.docx'; [string]$dst = Join-Path $Dir 'wpw.docx'
  $w = New-Object -ComObject Word.Application
  try { $w.DisplayAlerts = 0; $d = $w.Documents.Open($src, $false, $false, $false); $d.WritePassword = 'modify'; $d.SaveAs2($dst, 16); $d.Close($false) }
  finally { $w.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) }
  [string]$src = Join-Path $Dir 'plain.xlsx'; [string]$dst = Join-Path $Dir 'wpw.xlsx'
  $x = New-Object -ComObject Excel.Application
  try { $x.DisplayAlerts = $false; $wb = $x.Workbooks.Open($src); $wb.SaveAs($dst, 51, [Type]::Missing, 'modify'); $wb.Close($false) }
  finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
  [string]$src = Join-Path $Dir 'plain.pptx'; [string]$dst = Join-Path $Dir 'wpw.pptx'
  $p = New-Object -ComObject PowerPoint.Application
  try { $s = $p.Presentations.Open($src, 0, 0, 0); $s.WritePassword = 'modify'; $s.SaveAs($dst, 24); $s.Close() }
  finally { $p.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p) }
} finally { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); $m.ReleaseMutex(); $m.Dispose() }
`;

const WRITE_PW = { docx: ['word/settings.xml', /<w:writeProtection\b[^>]*w:hash=/], xlsx: ['xl/workbook.xml', /<fileSharing\b[^>]*hashValue=/],
  pptx: ['ppt/presentation.xml', /<p:modifyVerifier\b/] };
async function hasWritePassword(file) {
  const [part, re] = WRITE_PW[file.split('.').pop()];
  return re.test(await (await JSZip.loadAsync(readFileSync(file))).file(part).async('string'));
}

describe('프로그램 엔진 — OFFICEKIT_COM_TESTS=1 일 때만', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  test('시간 제한은 잠금을 잡은 뒤부터 재고, 넘으면 숨은 Office도 끝낸다', { timeout: 300_000 }, async (t) => {
    const huge = tmp.path('lock-huge.csv');
    writeFileSync(huge, Array.from({ length: 12_000 }, (_, i) => Array.from({ length: 20 }, (_, c) => (c ? i * c : `r${i}`)).join(',')).join('\n'));
    const before = await officeRunning('EXCEL.EXE');
    const holder = await holdLock(5000);
    if (!holder) return t.skip('다른 Office 작업이 잠금을 오래 쥐고 있음');
    const start = Date.now();
    await assert.rejects(comCall({ app: 'excel', op: 'text', path: huge }, { timeoutMs: 5000 }),
      /Excel이 응답하지 않아 작업을 멈췄습니다 \(5초 초과\)/);
    // 잠금 약 5초 + 일한 5초: 잠금을 기다린 시간이 들어갔다면 5초쯤에서 끝났을 것
    assert.ok(Date.now() - start >= 9000, `${Date.now() - start}ms`);
    await holder.done;
    await sleep(1500);
    if (!before) assert.equal(await officeRunning('EXCEL.EXE'), false, '숨은 EXCEL.EXE가 남음');
  });

  describe('쓰기 암호가 걸린 문서', () => {
    const dir = tmp.path('wpw');
    const file = (name) => join(dir, name);
    let ready = false;

    test('준비: 쓰기 암호 사본 만들기 (Word·Excel·PowerPoint)', { timeout: 240_000 }, async () => {
      mkdirSync(dir, { recursive: true });
      for (const ext of ['docx', 'xlsx', 'pptx']) await createDocument(file(`plain.${ext}`), MANUSCRIPT);
      const script = tmp.path('protect.ps1');
      writeFileSync(script, PROTECT_PS, 'ascii');
      await new Promise((done, fail) => execFile(POWERSHELL, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, dir],
        { timeout: 180_000, windowsHide: true }, (e, _o, se) => (e ? fail(new Error(`${e.message} ${se}`)) : done())));
      for (const ext of ['docx', 'xlsx', 'pptx']) assert.ok(await hasWritePassword(file(`wpw.${ext}`)), `wpw.${ext}에 쓰기 암호가 없음`);
      ready = true;
    });

    for (const ext of ['docx', 'xlsx', 'pptx']) {
      test(`.${ext} → PDF 변환이 암호 창에서 멈추지 않는다`, { timeout: 200_000 }, async (t) => {
        if (!ready) return t.skip('준비 실패');
        const out = file(`wpw-${ext}.pdf`);
        const start = Date.now();
        await convertDocument(file(`wpw.${ext}`), out);
        assert.ok(Date.now() - start < 60_000, `${Date.now() - start}ms`);
        assert.equal(readFileSync(out).subarray(0, 5).toString('latin1'), '%PDF-');
        assert.ok(await hasWritePassword(file(`wpw.${ext}`)), '원본의 쓰기 암호가 그대로');
      });

      test(`.${ext} 제자리 저장은 바로 한국어로 거절, 다른 파일로는 저장된다`, { timeout: 200_000 }, async (t) => {
        if (!ready) return t.skip('준비 실패');
        const src = file(`wpw.${ext}`);
        const before = readFileSync(src);
        let start = Date.now();
        await assert.rejects(replaceText(src, [['매출', '수익']], { engine: 'app' }), /^Error: 쓰기 암호가 걸린 문서라 그 자리에 저장할 수 없습니다: wpw\./);
        assert.ok(Date.now() - start < 15_000, `${Date.now() - start}ms`);
        assert.ok(before.equals(readFileSync(src)), '원본이 바뀜');
        const out = file(`wpw-out.${ext}`);
        start = Date.now();
        const r = await replaceText(src, [['매출', '수익']], { engine: 'app', output: out });
        assert.ok(Date.now() - start < 60_000, `${Date.now() - start}ms`);
        assert.equal(r.count, 2);
        // PowerPoint는 읽기 전용으로 열어 다른 이름으로 저장해도 쓰기 암호를 남긴다 → 같게. Word·Excel은 남기지 않는다.
        assert.equal(await hasWritePassword(out), ext === 'pptx');
        assert.match((await readDocument(out, { engine: 'app' })).text, /수익/);
      });
    }
  });
});
