// 문서·표·슬라이드 편집 창 e2e (설계 docs/OFFICE-EDITOR-DESIGN.md 10장) — 실제 앱 + CDP (scripts/dev/office-e2e.mjs)
//   1) 늘: 도우미 시험 (주소 가르기, 굵게·수식 읽기, PowerShell 도우미가 ASCII인지)
//   2) OFFICEKIT_EDITOR_SMOKE=1 이고 vendor/office-editor 묶음이 있을 때만: 실제 Electron 창이 뜨고 닫힌다 (몇 분)
//      형식 3 × 만든 방법(createDocument · '새 문서' doc:new edit:true · Office COM) → 열기 → 한국어 입력·굵게·표 행/수식·새 슬라이드
//      → Ctrl+S → readDocument, 밖으로 나간 요청 0건, WebRTC 후보 0개, 자동 저장 → 사본만, 바깥 변경 띠, 원래대로, 멈춤 뒤 복구, 닫기 확인
//   3) OFFICEKIT_COM_TESTS=1 (Word·Excel·PowerPoint가 있을 때): Office로 만든 문서도 쓰고, 저장한 파일을 COM 엄격 열기(strict-open.ps1)
//      COM 없이도 test/fixtures/office-editor/office.{docx,xlsx,pptx}(Office로 만든 고정 파일)가 있으면 그것으로 세 번째 방법을 돈다
//   보고서·화면: OFFICEKIT_SMOKE_OUT=<폴더> 이면 거기 (아니면 임시 폴더, 끝나면 지움)
//   실제 MS IME 시험은 사람이 한다: docs/QA-OFFICE-EDITOR.md
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import {
  docxParagraphBold, FORMATS, isLocalUrl, METHODS, PS_HELPER, PS_MAKE_OFFICE, pptxRunBold, runOfficeE2E, xlsxCell,
} from '../scripts/dev/office-e2e.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// ── 1) 도우미 ─────────────────────────────────────────────

test('isLocalUrl: 이 PC 주소·내장 주소만 안쪽으로 본다', () => {
  for (const u of ['http://127.0.0.1:1234/office-host/host.html', 'http://127.0.0.3:9/office/x/index.html', 'blob:http://127.0.0.3:5/abc', 'data:text/plain,a', 'about:blank', 'http://localhost:8080/', 'http://[::1]:80/']) {
    assert.equal(isLocalUrl(u), true, u);
  }
  for (const u of ['https://example.com/', 'http://10.0.0.1/', 'https://www.onlyoffice.com/', 'ws://stun.l.google.com:19302/', 'http://127.0.0.1.nip.io/', 'not a url']) {
    assert.equal(isLocalUrl(u), false, u);
  }
});

test('docxParagraphBold: 그 글이 든 문단의 글자 조각이 모두 굵은지', () => {
  const xml = '<w:body>'
    + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>굵은 </w:t></w:r><w:r><w:rPr><w:b w:val="1"/></w:rPr><w:t>문단</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>반만 </w:t></w:r><w:r><w:t>굵음</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>보통 &amp; 글</w:t></w:r></w:p></w:body>';
  assert.equal(docxParagraphBold(xml, '굵은 문단'), true);
  assert.equal(docxParagraphBold(xml, '반만 굵음'), false);
  assert.equal(docxParagraphBold(xml, '보통 & 글'), false);
  assert.equal(docxParagraphBold(xml, '없는 글'), null);
});

test('xlsxCell · pptxRunBold: 저장한 꾸러미에서 굵게·수식 읽기', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'ok-office-e2e-'));
  try {
    const x = new JSZip();
    x.file('[Content_Types].xml', '<Types/>');
    x.file('xl/workbook.xml', '<workbook><sheets><sheet name="예산" sheetId="1" r:id="rId1"/></sheets></workbook>');
    x.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    x.file('xl/sharedStrings.xml', '<sst><si><t>세종</t></si><si><t>제주</t></si></sst>');
    x.file('xl/styles.xml', '<styleSheet><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts><cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="1" applyFont="1"/></cellXfs></styleSheet>');
    x.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="3"><c r="A3" t="s"><v>0</v></c></row>'
      + '<row r="15"><c r="A15" s="1" t="s"><v>1</v></c><c r="B15"><v>10</v></c><c r="E15"><f>SUM(B15:D15)</f><v>60</v></c></row></sheetData></worksheet>');
    const xlsx = join(dir, 'a.xlsx');
    writeFileSync(xlsx, await x.generateAsync({ type: 'nodebuffer' }));
    assert.deepEqual(await xlsxCell(xlsx, '제주'), { ref: 'A15', row: 15, bold: true, formula: 'SUM(B15:D15)', sumValue: '60' });
    assert.equal((await xlsxCell(xlsx, '세종')).bold, false);
    assert.equal(await xlsxCell(xlsx, '없음'), null);

    const p = new JSZip();
    p.file('ppt/slides/slide1.xml', '<p:sld><a:r><a:rPr lang="ko-KR" b="1"/><a:t> — 판정 편집</a:t></a:r><a:r><a:rPr lang="ko-KR"/><a:t>보통 글</a:t></a:r></p:sld>');
    const pptx = join(dir, 'a.pptx');
    writeFileSync(pptx, await p.generateAsync({ type: 'nodebuffer' }));
    assert.equal(await pptxRunBold(pptx, ' — 판정 편집'), true);
    assert.equal(await pptxRunBold(pptx, '보통 글'), false);
    assert.equal(await pptxRunBold(pptx), true, '글을 안 주면 첫 슬라이드 첫 조각 (제목)');
    assert.equal(await pptxRunBold(pptx, '제목 — 판정 편집 뒤'), true, '그 글의 일부인 조각도 찾는다');
    assert.equal(await pptxRunBold(pptx, '없는 글'), null);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('PowerShell 도우미는 ASCII만 (Windows PowerShell 5.1은 BOM 없는 파일을 ANSI로 읽는다), 키 입력을 보내지 않는다', () => {
  for (const [name, s] of Object.entries({ PS_HELPER, PS_MAKE_OFFICE })) {
    assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(s), `${name}에 ASCII 밖 글자`);
    assert.doesNotMatch(s, /SendKeys|SendInput|keybd_event|mouse_event|SetForegroundWindow/i, name);
  }
  // 사용자 Office를 건드리지 않는다: 떠 있으면 거절, 공유 잠금
  assert.match(PS_MAKE_OFFICE, /already running/);
  assert.match(PS_MAKE_OFFICE, /Local\\OfficeKitCom/);
});

// ── 2) 실제 앱 e2e (선택) ──────────────────────────────────

const want = process.env.OFFICEKIT_EDITOR_SMOKE === '1';
const com = process.env.OFFICEKIT_COM_TESTS === '1';
// Office로 만든 고정 파일 (있으면 COM 없이도 세 번째 방법을 돈다)
const FIXTURES = join(ROOT, 'test', 'fixtures', 'office-editor');
const officeFixtures = FORMATS.every((ext) => existsSync(join(FIXTURES, `office.${ext}`)));
const methods = com || officeFixtures ? METHODS : METHODS.filter((m) => m !== 'office');
const electronOk = (() => { try { return existsSync(join(ROOT, 'node_modules', 'electron', 'path.txt')); } catch { return false; } })();
const ready = want && electronOk && existsSync(join(ROOT, 'vendor', 'office-editor', 'VERSION.json'));
const why = !want ? 'OFFICEKIT_EDITOR_SMOKE=1 일 때만 (실제 창이 뜬다, 몇 분)'
  : !electronOk ? 'electron이 설치돼 있지 않음' : 'vendor/office-editor 묶음이 없음 (npm run office-editor:build)';

test(`실제 앱 e2e: 형식 3 × 만든 방법 ${methods.length} → 열기·고치기·Ctrl+S·확인, 네트워크·WebRTC·자동 저장·바깥 변경·원래대로·닫기${com ? ' + COM 엄격 열기' : ''}`,
  { skip: !ready && why, timeout: 40 * 60_000 }, async (t) => {
    const keepOut = process.env.OFFICEKIT_SMOKE_OUT ? resolve(process.env.OFFICEKIT_SMOKE_OUT, 'office-e2e') : null;
    const out = keepOut ?? mkdtempSync(join(tmpdir(), 'ok-office-e2e-'));
    const lines = [];
    let report;
    try {
      report = await runOfficeE2E({
        out, formats: FORMATS, methods, com: com ? true : false, flows: true, log: (s) => lines.push(s),
        ...(!com && officeFixtures ? { officeDocs: FIXTURES } : {}),
      });
    } finally {
      if (!keepOut) rmSync(out, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
    }
    t.diagnostic(`${report.checks.filter((c) => c.ok).length}/${report.checks.length} 통과${keepOut ? ` · 보고서 ${join(out, 'report.json')}` : ''}`);
    for (const n of report.notes) t.diagnostic(n);
    // 문서마다·흐름마다 하위 시험으로 (실패한 확인과 그 정보를 그대로 보여 준다)
    const groups = [...new Set(report.checks.map((c) => c.group))];
    for (const g of groups) {
      await t.test(g, () => {
        const fails = report.checks.filter((c) => c.group === g && !c.ok);
        assert.deepEqual(fails.map((c) => `${c.name} — ${JSON.stringify(c.info).slice(0, 600)}`), []);
      });
    }
    // 빠진 것 없이 돌았는지
    const ran = new Set(report.docs.filter((d) => !d.skipped).map((d) => d.id));
    for (const ext of FORMATS) {
      for (const m of methods) assert.ok(ran.has(`${m}-${ext}`), `${m}-${ext}가 돌지 않음: ${report.docs.find((d) => d.id === `${m}-${ext}`)?.skipped ?? ''}`);
    }
    assert.ok(report.checks.some((c) => c.group === 'flow-docx'), '흐름 시험이 돌았다');
    if (com) assert.ok(report.checks.some((c) => c.name.startsWith('COM 엄격 열기')), 'COM 엄격 열기가 돌았다');
    assert.equal(report.ok, true, lines.filter((l) => l.startsWith('FAIL')).join('\n'));
  });
