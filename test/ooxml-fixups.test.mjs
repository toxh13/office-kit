// ooxml-fixups.test.mjs — core/ooxml-fixups.mjs: Office 편집기(x2t)가 내보낸 꾸러미를 저장 직전에 고치기.
// 고정 파일 test/fixtures/office-editor/ = 판정 실험(tmp/spike-judge/out/j1)이 편집기에서 저장한 그대로의 바이트:
//   blank-raw.pptx  'createBlankDocument'(새 문서) pptx를 편집·저장 — 고치지 않으면 PowerPoint 복구 창에서 멈춤
//   judge-raw.docx  createDocument docx에 표 행을 더해 저장 — 새 행에 혼자 있는 vMerge restart 3개
//   judge-raw.xlsx  createDocument xlsx를 편집·저장 — 아무도 안 쓰는 xl/theme/theme.xml
//   carry/<이름>.orig.*·.saved.*  Office로 만든 한국어 문서와 그것을 편집기에서 저장한 것 (원본에서 옮겨 오기)
// Office 엄격 열기(scripts/dev/strict-open.ps1)는 OFFICEKIT_COM_TESTS=1 일 때만 (한컴 COM은 쓰지 않음).
import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { fixOoxml, findOoxmlIssues, dropLoneVMergeRestarts, xmlBalanced, FIXUP_EXTS, PART_MAX_BYTES } from '../src/core/ooxml-fixups.mjs';
import { excelEdit } from '../src/core/excel.mjs';
import { COM_TESTS, ROOT, tempDir } from './helpers/fixtures.mjs';

const FIX = (name) => readFileSync(join(ROOT, 'test', 'fixtures', 'office-editor', name));
const PPTX = FIX('blank-raw.pptx');
const DOCX = FIX('judge-raw.docx');
const XLSX = FIX('judge-raw.xlsx');
const THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const tmp = tempDir('officekit-fixups-test-');
after(() => tmp.cleanup());

// 꾸러미 → { names(순서 그대로), part(name) → Buffer, text(name) }
async function unzip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files);
  const parts = new Map();
  for (const n of names) if (!zip.files[n].dir) parts.set(n, Buffer.from(await zip.file(n).async('uint8array')));
  return { names, parts, part: (n) => parts.get(n), text: (n) => parts.get(n)?.toString('utf8') };
}
// 고정 파일을 조금 바꾼 꾸러미
async function mutate(bytes, fn) {
  const zip = await JSZip.loadAsync(bytes);
  await fn(zip, async (n) => zip.file(n).async('string'));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
// 손대지 않은 부분은 바이트 그대로인지
function assertUntouched(before, afterPkg, skip) {
  for (const [n, b] of before.parts) {
    if (skip.includes(n)) continue;
    assert.ok(afterPkg.parts.has(n), `빠진 부분: ${n}`);
    assert.ok(afterPkg.part(n).equals(b), `바뀐 부분: ${n}`);
  }
}

// ── pptx: 노트 마스터 테마 나누기 ───────────────────────────

describe('pptx — 마스터마다 자기 테마 (필수)', () => {
  test('새 문서 pptx: 노트 마스터가 슬라이드 마스터의 theme1을 같이 씀 → theme2로 나눈다', async () => {
    assert.deepEqual(await findOoxmlIssues(PPTX, 'pptx'),
      ['노트 마스터가 슬라이드 마스터와 테마를 같이 씀: ppt/theme/theme1.xml (ppt/notesMasters/notesMaster1.xml, ppt/slideMasters/slideMaster1.xml)']);
    const copy = Buffer.from(PPTX);
    const r = await fixOoxml(PPTX, 'pptx');
    assert.ok(PPTX.equals(copy), '입력 바이트를 바꾸면 안 됨');
    assert.equal(r.error, null);
    assert.equal(r.notify, false);
    assert.equal(r.changed, true);
    assert.deepEqual(r.fixes, ['노트 마스터 테마 나눔: ppt/theme/theme1.xml → ppt/theme/theme2.xml (ppt/notesMasters/notesMaster1.xml)']);
    assert.ok(Buffer.isBuffer(r.bytes));

    const a = await unzip(PPTX); const b = await unzip(r.bytes);
    assert.deepEqual(b.names, [...a.names, 'ppt/theme/theme2.xml', 'ppt/theme/_rels/theme2.xml.rels'], '새 부분은 뒤에, 폴더 항목 없이');
    assert.ok(b.part('ppt/theme/theme2.xml').equals(a.part('ppt/theme/theme1.xml')));
    assert.ok(b.part('ppt/theme/_rels/theme2.xml.rels').equals(a.part('ppt/theme/_rels/theme1.xml.rels')));
    assert.equal(b.text('ppt/notesMasters/_rels/notesMaster1.xml.rels'),
      a.text('ppt/notesMasters/_rels/notesMaster1.xml.rels').replace('Target="../theme/theme1.xml"', 'Target="../theme/theme2.xml"'));
    assert.equal(b.text('[Content_Types].xml'),
      a.text('[Content_Types].xml').replace('</Types>', `<Override PartName="/ppt/theme/theme2.xml" ContentType="${THEME_CT}"/></Types>`));
    // 슬라이드 마스터·발표 자료는 theme1 그대로
    assert.match(b.text('ppt/slideMasters/_rels/slideMaster1.xml.rels'), /Target="\.\.\/theme\/theme1\.xml"/);
    assertUntouched(a, b, ['[Content_Types].xml', 'ppt/notesMasters/_rels/notesMaster1.xml.rels']);
    assert.deepEqual(await findOoxmlIssues(r.bytes, 'pptx'), []);
  });

  test('한 번 더 고쳐도 그대로 (이미 고친 파일·고칠 것 없는 파일은 바이트 그대로)', async () => {
    const once = await fixOoxml(PPTX, 'pptx');
    const twice = await fixOoxml(once.bytes, 'pptx');
    assert.equal(twice.changed, false);
    assert.equal(twice.error, null);
    assert.deepEqual(twice.fixes, []);
    assert.ok(twice.bytes.equals(once.bytes));
  });

  test('받는 형식: Buffer·Uint8Array·ArrayBuffer, 확장자는 점·대문자도', async () => {
    const u8 = new Uint8Array(PPTX.length + 8);
    u8.set(PPTX, 4);
    const view = u8.subarray(4, 4 + PPTX.length); // 더 큰 버퍼 안의 일부 (IPC로 오는 모양)
    const ab = PPTX.buffer.slice(PPTX.byteOffset, PPTX.byteOffset + PPTX.length);
    const want = (await fixOoxml(PPTX, 'pptx')).bytes;
    for (const [input, ext] of [[view, '.PPTX'], [ab, 'Pptx']]) {
      const r = await fixOoxml(input, ext);
      assert.equal(r.changed, true, ext);
      assert.ok(r.bytes.equals(want), ext);
    }
    await assert.rejects(fixOoxml('nope', 'pptx'), TypeError);
  });

  test('theme2가 이미 있으면 다음 빈 번호 (theme3)', async () => {
    const input = await mutate(PPTX, async (z) => { z.file('ppt/theme/theme2.xml', '<a:theme xmlns:a="x"/>'); });
    const r = await fixOoxml(input, 'pptx');
    assert.equal(r.error, null);
    const b = await unzip(r.bytes);
    assert.match(b.text('ppt/notesMasters/_rels/notesMaster1.xml.rels'), /Target="\.\.\/theme\/theme3\.xml"/);
    assert.equal(b.text('ppt/theme/theme2.xml'), '<a:theme xmlns:a="x"/>');
    assert.match(b.text('[Content_Types].xml'), /<Override PartName="\/ppt\/theme\/theme3\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.theme\+xml"\/><\/Types>$/);
  });

  test('유인물 마스터도 같이 쓰면 각자 사본 (노트 → theme2, 유인물 → theme3), 절대 경로 Target은 절대 경로로', async () => {
    const input = await mutate(PPTX, async (z, text) => {
      z.file('ppt/handoutMasters/handoutMaster1.xml', '<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');
      z.file('ppt/handoutMasters/_rels/handoutMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/theme" Target="/ppt/theme/theme1.xml"/></Relationships>`);
      z.file('[Content_Types].xml', (await text('[Content_Types].xml')).replace('</Types>',
        '<Override PartName="/ppt/handoutMasters/handoutMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"/></Types>'));
    });
    assert.equal((await findOoxmlIssues(input, 'pptx')).length, 2);
    const r = await fixOoxml(input, 'pptx');
    assert.equal(r.error, null);
    assert.deepEqual(r.fixes, [
      '노트 마스터 테마 나눔: ppt/theme/theme1.xml → ppt/theme/theme2.xml (ppt/notesMasters/notesMaster1.xml)',
      '유인물 마스터 테마 나눔: ppt/theme/theme1.xml → ppt/theme/theme3.xml (ppt/handoutMasters/handoutMaster1.xml)',
    ]);
    const b = await unzip(r.bytes);
    assert.match(b.text('ppt/handoutMasters/_rels/handoutMaster1.xml.rels'), /Target="\/ppt\/theme\/theme3\.xml"/);
    assert.deepEqual(await findOoxmlIssues(r.bytes, 'pptx'), []);
  });

  test('같이 쓰는 테마 부분이 없으면: 원래 바이트 + 오류 + 알림 띠', async () => {
    const input = await mutate(PPTX, async (z) => { z.remove('ppt/theme/theme1.xml'); });
    const r = await fixOoxml(input, 'pptx');
    assert.equal(r.changed, false);
    assert.ok(r.bytes.equals(input));
    assert.equal(r.notify, true);
    assert.deepEqual(r.fixes, []);
    assert.match(r.error, /^PowerPoint에서 경고 없이 열리도록 고치지 못해 편집기가 만든 그대로 저장했습니다\. PowerPoint에서 열 때 복구 창이 뜰 수 있습니다 \(노트 마스터가 쓰는 테마 부분이 없음: ppt\/theme\/theme1\.xml\)$/);
  });

  test('관계 Target이 작은따옴표·다른 속성 순서여도 고친다', async () => {
    const input = await mutate(PPTX, async (z, text) => {
      z.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', (await text('ppt/notesMasters/_rels/notesMaster1.xml.rels'))
        .replace(/<Relationship [^>]*\/>/, `<Relationship Target='../theme/theme1.xml' Id='rId1' Type='${REL}/theme' />`));
    });
    const r = await fixOoxml(input, 'pptx');
    assert.equal(r.error, null);
    assert.match((await unzip(r.bytes)).text('ppt/notesMasters/_rels/notesMaster1.xml.rels'), /<Relationship Target="\.\.\/theme\/theme2\.xml" Id='rId1'/);
  });
});

// ── xlsx: 아무도 안 쓰는 테마 ───────────────────────────────

describe('xlsx — 쓰지 않는 테마 지우기', () => {
  test('x2t가 남긴 xl/theme/theme.xml만 지우고 theme1.xml(통합 문서가 씀)은 둔다', async () => {
    assert.deepEqual(await findOoxmlIssues(XLSX, 'xlsx'), ['쓰지 않는 테마: xl/theme/theme.xml']);
    const r = await fixOoxml(XLSX, 'xlsx');
    assert.equal(r.error, null);
    assert.equal(r.changed, true);
    assert.deepEqual(r.fixes, ['쓰지 않는 테마 지움: xl/theme/theme.xml']);
    const a = await unzip(XLSX); const b = await unzip(r.bytes);
    assert.deepEqual(b.names, a.names.filter((n) => n !== 'xl/theme/theme.xml'));
    assertUntouched(a, b, ['xl/theme/theme.xml']); // [Content_Types].xml도 그대로 (theme.xml은 Override가 없었음)
    assert.deepEqual(await findOoxmlIssues(r.bytes, 'xlsx'), []);
    const again = await fixOoxml(r.bytes, 'xlsx');
    assert.equal(again.changed, false);
    assert.ok(again.bytes.equals(r.bytes));
  });

  test('지우는 테마의 Override·관계 파일도 같이 지운다', async () => {
    const input = await mutate(XLSX, async (z, text) => {
      z.file('xl/theme/_rels/theme.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
      z.file('[Content_Types].xml', (await text('[Content_Types].xml')).replace('</Types>', `<Override PartName="/xl/theme/theme.xml" ContentType="${THEME_CT}"/></Types>`));
    });
    const r = await fixOoxml(input, 'xlsx');
    assert.equal(r.error, null);
    const b = await unzip(r.bytes);
    assert.ok(!b.parts.has('xl/theme/theme.xml') && !b.parts.has('xl/theme/_rels/theme.xml.rels'));
    assert.doesNotMatch(b.text('[Content_Types].xml'), /theme\/theme\.xml/);
    assert.match(b.text('[Content_Types].xml'), /PartName="\/xl\/theme\/theme1\.xml"/);
  });

  test('관계가 가리키는 테마는 이름과 상관없이 남긴다 (theme.xml을 쓰고 theme1.xml이 남는 경우)', async () => {
    const input = await mutate(XLSX, async (z, text) => {
      z.file('xl/_rels/workbook.xml.rels', (await text('xl/_rels/workbook.xml.rels')).replace('Target="theme/theme1.xml"', 'Target="theme/theme.xml"'));
    });
    const r = await fixOoxml(input, 'xlsx');
    assert.deepEqual(r.fixes, ['쓰지 않는 테마 지움: xl/theme/theme1.xml']);
    const b = await unzip(r.bytes);
    assert.ok(b.parts.has('xl/theme/theme.xml') && !b.parts.has('xl/theme/theme1.xml'));
    assert.doesNotMatch(b.text('[Content_Types].xml'), /theme1\.xml/);
  });

  test('꾸러미 뿌리 관계(_rels/.rels)가 가리키는 것도 쓰는 것으로 본다', async () => {
    const input = await mutate(XLSX, async (z, text) => {
      z.file('_rels/.rels', (await text('_rels/.rels')).replace('</Relationships>', `<Relationship Id="rIdT" Type="${REL}/theme" Target="xl/theme/theme.xml"/></Relationships>`));
    });
    const r = await fixOoxml(input, 'xlsx');
    assert.equal(r.changed, false);
    assert.equal(r.error, null);
  });

  test('통합 문서 관계 파일이 없으면 판단하지 않고 그대로', async () => {
    const input = await mutate(XLSX, async (z) => { z.remove('xl/_rels/workbook.xml.rels'); });
    const r = await fixOoxml(input, 'xlsx');
    assert.equal(r.changed, false);
    assert.equal(r.error, null);
    assert.ok(r.bytes.equals(input));
  });
});

// ── docx: 혼자 있는 세로 병합 시작 ──────────────────────────

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = '<w:vMerge w:val="restart"/>';
const C = '<w:vMerge/>';
const span = (n) => `<w:gridSpan w:val="${n}"/>`;
const tc = (text, pr = '') => `<w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/>${pr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const tr = (cells, pr = '') => `<w:tr>${pr ? `<w:trPr>${pr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
const tbl = (...rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>${rows.join('')}</w:tbl>`;
const doc = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body}<w:p/></w:body></w:document>`;
const vmergeCount = (xml) => (xml.match(/<w:vMerge\b/g) ?? []).length;

describe('docx — 혼자 있는 vMerge restart 지우기', () => {
  test('편집기에서 더한 표 행: restart 3개를 지우고 나머지 XML은 한 글자도 바꾸지 않는다', async () => {
    assert.deepEqual(await findOoxmlIssues(DOCX, 'docx'), ['혼자 있는 세로 병합 시작 3개: word/document.xml']);
    const r = await fixOoxml(DOCX, 'docx');
    assert.equal(r.error, null);
    assert.equal(r.changed, true);
    assert.deepEqual(r.fixes, ['혼자 있는 세로 병합 시작 3개 지움: word/document.xml']);
    const a = await unzip(DOCX); const b = await unzip(r.bytes);
    assert.deepEqual(b.names, a.names);
    const before = a.text('word/document.xml');
    assert.equal(before.split('<w:vMerge w:val="restart" />').length - 1, 3);
    assert.equal(b.text('word/document.xml'), before.split('<w:vMerge w:val="restart" />').join(''));
    assertUntouched(a, b, ['word/document.xml']);
    assert.deepEqual(await findOoxmlIssues(r.bytes, 'docx'), []);
  });

  test('진짜 세로 병합(restart 아래 continue)은 남긴다', () => {
    const xml = doc(tbl(tr([tc('a', R), tc('b')]), tr([tc('', C), tc('c')]), tr([tc('', C), tc('d')])));
    assert.deepEqual(dropLoneVMergeRestarts(xml), { xml, removed: 0 });
  });

  test('마지막 행·아래가 이어 받지 않는 restart는 지운다 (tcPr의 다른 속성은 그대로)', () => {
    const r = dropLoneVMergeRestarts(doc(tbl(tr([tc('a', R), tc('b')]), tr([tc('x'), tc('y', R)]))));
    assert.equal(r.removed, 2);
    assert.equal(r.xml, doc(tbl(tr([tc('a'), tc('b')]), tr([tc('x'), tc('y')]))));
  });

  test('열 위치는 gridSpan·gridBefore로 센다', () => {
    // 두 칸짜리 병합: 아래 행도 두 칸 continue → 남김
    const merged = doc(tbl(tr([tc('a', span(2) + R), tc('b')]), tr([tc('', span(2) + C), tc('c')])));
    assert.equal(dropLoneVMergeRestarts(merged).removed, 0);
    // 1열의 restart 아래: 0~1열을 차지한 보통 칸 → 혼자 (2열의 continue와 헷갈리지 않음)
    const lone = doc(tbl(tr([tc('a'), tc('b', R), tc('c')]), tr([tc('x', span(2)), tc('', C)])));
    assert.equal(dropLoneVMergeRestarts(lone).removed, 1);
    // gridBefore=1 이면 아래 행 첫 칸이 1열 → 이어 받음
    const before1 = doc(tbl(tr([tc('a'), tc('b', R)]), tr([tc('', C)], '<w:gridBefore w:val="1"/>')));
    assert.equal(dropLoneVMergeRestarts(before1).removed, 0);
    const before0 = doc(tbl(tr([tc('a'), tc('b', R)]), tr([tc('', C)])));
    assert.equal(dropLoneVMergeRestarts(before0).removed, 1);
  });

  test('표 속 표: 안쪽 표는 따로 본다', () => {
    const inner = tbl(tr([tc('i', R)]));
    const xml = doc(tbl(tr([tc('o', R)]), tr([`<w:tc><w:tcPr>${C}</w:tcPr>${inner}<w:p/></w:tc>`])));
    const r = dropLoneVMergeRestarts(xml);
    assert.equal(r.removed, 1);
    assert.equal(vmergeCount(r.xml), 2); // 바깥 restart·continue는 남음
    assert.ok(r.xml.includes(tbl(tr([tc('i')]))));
  });

  test('닫는 태그가 따로 있는 vMerge도 통째로 지운다', () => {
    const r = dropLoneVMergeRestarts(doc(tbl(tr([tc('a', '<w:vMerge w:val="restart"></w:vMerge>')]))));
    assert.equal(r.removed, 1);
    assert.equal(r.xml, doc(tbl(tr([tc('a')]))));
  });

  test('변경 추적이 걸린 표·모르는 값이 있는 표는 손대지 않는다', () => {
    const tracked = doc(tbl(tr([tc('a', R + '<w:tcPrChange w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z"><w:tcPr/></w:tcPrChange>')])));
    assert.equal(dropLoneVMergeRestarts(tracked).removed, 0);
    const rowIns = doc(tbl(tr([tc('a', R)], '<w:ins w:id="2" w:author="x" w:date="2026-01-01T00:00:00Z"/>')));
    assert.equal(dropLoneVMergeRestarts(rowIns).removed, 0);
    const odd = doc(tbl(tr([tc('a', R)]), tr([tc('b', '<w:vMerge w:val="sideways"/>')])));
    assert.equal(dropLoneVMergeRestarts(odd).removed, 0);
    // 문단 안의 변경 추적(w:ins)은 상관없다
    const runIns = doc(tbl(tr([`<w:tc><w:tcPr>${R}</w:tcPr><w:p><w:ins w:id="3" w:author="x"><w:r><w:t>z</w:t></w:r></w:ins></w:p></w:tc>`])));
    assert.equal(dropLoneVMergeRestarts(runIns).removed, 1);
  });

  test('짝이 맞지 않는 표 XML은 예외 → fixOoxml은 원래 바이트 + 오류 (알림 띠 없음)', async () => {
    assert.throws(() => dropLoneVMergeRestarts(`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tr><w:tc><w:tcPr>${R}</w:tcPr></w:tr></w:tbl></w:body></w:document>`), /짝이 맞지 않음/);
    const input = await mutate(DOCX, async (z, text) => {
      z.file('word/document.xml', (await text('word/document.xml')).replace('</w:tbl>', '</w:tr></w:tbl>'));
    });
    const r = await fixOoxml(input, 'docx');
    assert.equal(r.changed, false);
    assert.ok(r.bytes.equals(input));
    assert.equal(r.notify, false);
    assert.match(r.error, /^Word에서 경고 없이 열리도록 고치지 못해 편집기가 만든 그대로 저장했습니다 \(표 XML 짝이 맞지 않음\)$/);
  });

  test('머리글의 표도 고친다 (주 문서 관계로 찾음)', async () => {
    const input = await mutate(DOCX, async (z, text) => {
      z.file('word/header9.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}">${tbl(tr([tc('h', R)]))}<w:p/></w:hdr>`);
      z.file('word/_rels/document.xml.rels', (await text('word/_rels/document.xml.rels'))
        .replace('</Relationships>', `<Relationship Id="rIdH9" Type="${REL}/header" Target="header9.xml"/></Relationships>`));
      z.file('[Content_Types].xml', (await text('[Content_Types].xml'))
        .replace('</Types>', '<Override PartName="/word/header9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
    });
    const r = await fixOoxml(input, 'docx');
    assert.equal(r.error, null);
    assert.deepEqual(r.fixes, ['혼자 있는 세로 병합 시작 3개 지움: word/document.xml', '혼자 있는 세로 병합 시작 1개 지움: word/header9.xml']);
    assert.equal(vmergeCount((await unzip(r.bytes)).text('word/header9.xml')), 0);
  });
});

// ── 공통 ─────────────────────────────────────────────────

describe('실패해도 깨뜨리지 않는다', () => {
  test('꾸러미가 아니면 원래 바이트 + 오류 (pptx만 알림 띠)', async () => {
    const junk = Buffer.from('PK\u0003\u0004 이것은 zip이 아님');
    for (const ext of FIXUP_EXTS) {
      const r = await fixOoxml(junk, ext);
      assert.equal(r.changed, false, ext);
      assert.ok(r.bytes.equals(junk), ext);
      assert.equal(r.notify, ext === 'pptx', ext);
      assert.match(r.error, /고치지 못해 편집기가 만든 그대로 저장했습니다/, ext);
    }
    const cut = await fixOoxml(PPTX.subarray(0, PPTX.length - 100), 'pptx'); // 끝이 잘린 zip
    assert.equal(cut.changed, false);
    assert.equal(cut.notify, true);
    assert.ok(cut.bytes.equals(PPTX.subarray(0, PPTX.length - 100)));
  });

  test('[Content_Types].xml이 없으면 오류', async () => {
    const input = await mutate(PPTX, async (z) => { z.remove('[Content_Types].xml'); });
    const r = await fixOoxml(input, 'pptx');
    assert.ok(r.bytes.equals(input));
    assert.match(r.error, /\[Content_Types\]\.xml 없음/);
  });

  test('UTF-16 관계 파일은 고치지 않는다 (글자 깨짐 방지)', async () => {
    const input = await mutate(PPTX, async (z, text) => {
      const s = await text('ppt/notesMasters/_rels/notesMaster1.xml.rels');
      z.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]));
    });
    const r = await fixOoxml(input, 'pptx');
    assert.equal(r.changed, false);
    assert.match(r.error, /UTF-16/);
  });

  test('고치기가 없는 형식은 그대로 (오류 아님)', async () => {
    for (const ext of ['hwp', 'hwpx', 'pdf', '', undefined]) {
      const r = await fixOoxml(PPTX, ext);
      assert.deepEqual({ ...r, bytes: null }, { bytes: null, changed: false, fixes: [], warnings: [], error: null, notify: false });
      assert.ok(r.bytes.equals(PPTX));
    }
    assert.deepEqual(await findOoxmlIssues(PPTX, 'hwp'), []);
  });

  test('xmlBalanced: 짝 확인', async () => {
    for (const [name, buf] of [['pptx', PPTX], ['docx', DOCX], ['xlsx', XLSX]]) {
      const p = await unzip(buf);
      for (const [n, b] of p.parts) if (/\.(xml|rels)$/.test(n)) assert.ok(xmlBalanced(b.toString('utf8')), `${name} ${n}`);
    }
    assert.ok(xmlBalanced('<?xml version="1.0"?><!-- <x> --><a b="1 > 0"><![CDATA[<y>]]>1 > 0<c/></a>'));
    for (const bad of ['<a><b></a>', '<a>', '<a/><b/>', '<a></a></a>', '<a><b</a>', '', 'text']) assert.equal(xmlBalanced(bad), false, bad);
  });
});

// ── Office 엄격 열기 — OFFICEKIT_COM_TESTS=1 일 때만 ─────────

const POWERSHELL = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const STRICT = join(ROOT, 'scripts', 'dev', 'strict-open.ps1');
const OFFICE_IMAGES = ['WINWORD.EXE', 'EXCEL.EXE', 'POWERPNT.EXE'];

// strict-open.ps1 → { code, r } (code 0 깨끗이 열림, 1 멈춤·실패, 3 바쁨)
function strictOpen(file, timeoutSec = 60) {
  const out = `${file}.strict.json`;
  return new Promise((done) => execFile(POWERSHELL,
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', STRICT, '-Path', file, '-Out', out, '-TimeoutSec', String(timeoutSec), '-LockWaitSec', '240'],
    { windowsHide: true, timeout: (timeoutSec + 300) * 1000 },
    (err) => {
      let r = null;
      try { r = JSON.parse(readFileSync(out, 'utf8')); } catch { /* 결과 없음 */ }
      done({ code: err ? (err.code ?? -1) : 0, r });
    }));
}
// 지금 떠 있는 Office 프로세스 PID들
const officePids = () => new Promise((done) => execFile('tasklist', ['/FO', 'CSV', '/NH'], { encoding: 'utf8', windowsHide: true }, (_e, out) => {
  const pids = new Set();
  for (const line of String(out).split(/\r?\n/)) {
    const [image, pid] = line.split('","').map((s) => s.replace(/"/g, ''));
    if (OFFICE_IMAGES.includes(String(image).toUpperCase())) pids.add(Number(pid));
  }
  done(pids);
}));

// ── 원본에서 옮겨 오기 ─────────────────────────────────────
// test/fixtures/office-editor/carry/<이름>.orig.* = Word·Excel(COM)로 만든 한국어 문서, <이름>.saved.* = 그것을 편집기에서 고쳐 저장한 바이트
// (검증 tmp/verify-fidelity: 고치기 전에는 Word가 한글을 바탕으로·'3 분기'처럼 띄워 그리고, Excel은 행이 낮아지고 차트·그림이 찌그러짐)

const CARRY = (name) => FIX(`carry/${name}`);
const fontSchemeOf = (xml) => xml.match(/<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/)[0];

describe('원본에서 옮겨 오기 — docx (테마 글꼴·글꼴 언어·한국어 줄 바꿈·호환 설정·내용 컨트롤)', () => {
  const orig = CARRY('minutes.orig.docx');
  const saved = CARRY('minutes.saved.docx');

  test('x2t가 잃은 것만 원본대로 (맑은 고딕, 글자 단위 줄 바꿈, 한글·숫자 붙여 씀, 표 줄 높이), 나머지 부분은 바이트 그대로', async () => {
    const r = await fixOoxml(saved, 'docx', { original: orig });
    assert.equal(r.error, null);
    assert.deepEqual(r.warnings, []);
    assert.equal(r.changed, true);
    assert.deepEqual(r.fixes, [
      '테마 글꼴 묶음 원본대로: word/theme/theme1.xml',
      '문서 글꼴 언어 원본대로: <w:themeFontLang w:val="en-US" w:eastAsia="ko-KR"/>',
      '기본 문단 스타일 원본대로: wordWrap, autoSpaceDE, autoSpaceDN',
      '호환 설정 5개 원본대로: w:spaceForUL, w:doNotLeaveBackslashAlone, w:adjustLineHeightInTable, differentiateMultirowTableHeaders, useWord2013TrackBottomHyphenation',
    ]);
    const a = await unzip(orig); const s = await unzip(saved); const b = await unzip(r.bytes);
    // 테마 글꼴 묶음: 한글(Hang) 맑은 고딕, 빈 ea (x2t는 ea=Arial, 스크립트별 글꼴 없음)
    assert.doesNotMatch(fontSchemeOf(s.text('word/theme/theme1.xml')), /script="Hang"/);
    assert.equal(fontSchemeOf(b.text('word/theme/theme1.xml')), fontSchemeOf(a.text('word/theme/theme1.xml')));
    assert.match(fontSchemeOf(b.text('word/theme/theme1.xml')), /<a:font script="Hang" typeface="맑은 고딕"\/>/);
    // 테마 글꼴 언어: zh-CN → ko-KR
    assert.match(s.text('word/settings.xml'), /w:eastAsia="zh-CN"/);
    assert.equal(b.text('word/settings.xml').match(/<w:themeFontLang\b[^>]*>/)[0], '<w:themeFontLang w:val="en-US" w:eastAsia="ko-KR"/>');
    // 기본 문단 스타일: CT_PPrBase 순서대로 pBdr 뒤에
    assert.match(b.text('word/styles.xml'), /<w:style w:type="paragraph" w:styleId="\d+" w:default="1"><w:name w:val="Normal"\/><w:qFormat\/><w:pPr><w:widowControl w:val="false" \/><w:pBdr><\/w:pBdr><w:wordWrap w:val="0"\/><w:autoSpaceDE w:val="0"\/><w:autoSpaceDN w:val="0"\/><w:spacing \/><w:ind \/><\/w:pPr><\/w:style>/);
    // 호환 설정: 원본 순서대로 끼우고, 저장본에 있던 것은 그대로
    const compat = b.text('word/settings.xml').match(/<w:compat>[\s\S]*?<\/w:compat>/)[0];
    assert.equal(compat.replace(/<w:compatSetting[^>]*\/>/g, '').replace(/ w:val="true" \//g, '/'),
      '<w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/></w:compat>');
    for (const n of ['compatibilityMode', 'differentiateMultirowTableHeaders', 'useWord2013TrackBottomHyphenation']) assert.match(compat, new RegExp(`w:name="${n}"`));
    assert.ok(xmlBalanced(b.text('word/styles.xml')) && xmlBalanced(b.text('word/settings.xml')));
    assertUntouched(s, b, ['word/theme/theme1.xml', 'word/settings.xml', 'word/styles.xml']);
    // 이미 옮긴 파일을 다시 고치면 그대로
    const again = await fixOoxml(r.bytes, 'docx', { original: orig });
    assert.equal(again.changed, false);
    assert.ok(again.bytes.equals(r.bytes));
  });

  test('원본이 없으면 (새 이름으로 저장 등) 글꼴 언어만: 문서 기본 언어(docDefaults w:lang)의 eastAsia로', async () => {
    const r = await fixOoxml(saved, 'docx');
    assert.deepEqual(r.fixes, ['문서 글꼴 언어 원본대로: <w:themeFontLang w:val="en-US" w:eastAsia="ko-KR" />']);
    // 문서 기본 언어가 없으면 손대지 않는다 (judge-raw.docx)
    assert.equal((await fixOoxml(DOCX, 'docx')).fixes.some((f) => f.includes('글꼴 언어')), false);
  });

  test('편집기에서 테마 글꼴을 바꿨으면 (제목·본문 라틴 글꼴이 다르면) 테마는 저장본 그대로', async () => {
    const changed = await mutate(saved, async (z, read) => {
      z.file('word/theme/theme1.xml', (await read('word/theme/theme1.xml')).replace(/(<a:minorFont><a:latin typeface=")[^"]*"/, '$1Cambria"'));
    });
    const r = await fixOoxml(changed, 'docx', { original: orig });
    assert.equal(r.error, null);
    assert.ok(!r.fixes.some((f) => f.startsWith('테마 글꼴')), r.fixes.join('\n'));
    assert.ok(r.fixes.some((f) => f.startsWith('기본 문단 스타일')));
  });

  test('원본이 꾸러미가 아니면: 옮기기 없이 필수 고치기만 + 경고', async () => {
    const plain = await fixOoxml(DOCX, 'docx');
    const r = await fixOoxml(DOCX, 'docx', { original: Buffer.from('not a zip') });
    assert.equal(r.error, null);
    assert.deepEqual(r.fixes, plain.fixes);
    assert.match(r.warnings[0], /^원본을 읽지 못해 원본에서 옮기기는 건너뜀/);
    assert.ok(r.bytes.equals(plain.bytes));
  });

  test('내용 컨트롤: 원본이 일반 텍스트(w:text)였는데 종류가 빠졌으면 되살린다 (w:id로 맞춤)', async () => {
    const sdt = (id, type) => `<w:sdt><w:sdtPr><w:alias w:val="담당자" /><w:id w:val="${id}" />${type}</w:sdtPr><w:sdtContent><w:p><w:r><w:t>홍길동</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
    const withSdt = (types) => mutate(DOCX, async (z, read) => {
      z.file('word/document.xml', (await read('word/document.xml')).replace('<w:body>', `<w:body>${types.map(([id, t]) => sdt(id, t)).join('')}`));
    });
    const original = await withSdt([['11', '<w:text/>'], ['12', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>'], ['13', '<w:text/>']]);
    const editor = await withSdt([['11', '<w:rPr></w:rPr>'], ['12', ''], ['13', '<w:richText/>'], ['99', '']]);
    const r = await fixOoxml(editor, 'docx', { original });
    assert.equal(r.error, null);
    assert.ok(r.fixes.includes('내용 컨트롤 종류 2개 원본대로'), r.fixes.join('\n'));
    const doc = (await unzip(r.bytes)).text('word/document.xml');
    assert.ok(doc.includes('<w:id w:val="11" /><w:rPr></w:rPr><w:text/></w:sdtPr>'), '11: 일반 텍스트로');
    assert.ok(doc.includes('<w:id w:val="12" /><w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr>'), '12: 날짜로');
    assert.ok(doc.includes('<w:id w:val="13" /><w:richText/></w:sdtPr>'), '13: 저장본에 종류가 있으면 그대로');
    assert.ok(doc.includes('<w:id w:val="99" /></w:sdtPr>'), '99: 원본에 없는 컨트롤은 그대로');
  });
});

describe('원본에서 옮겨 오기 — xlsx (셀 글꼴·메모·조건부 서식)', () => {
  const orig = CARRY('sales.orig.xlsx');
  const saved = CARRY('sales.saved.xlsx');

  test('sales: 셀 글꼴 맑은 고딕(+charset), 스레드 메모 → 원래 메모, 확장에만 있던 cellIs → 기본 조건부 서식', async () => {
    const r = await fixOoxml(saved, 'xlsx', { original: orig });
    assert.equal(r.error, null);
    assert.deepEqual(r.warnings, []);
    assert.deepEqual(r.fixes, [
      '테마 글꼴 묶음 원본대로: xl/theme/theme1.xml',
      '셀 글꼴 2개 이름 원본대로: 맑은 고딕',
      '스레드 메모 1개를 원래 메모로 되돌림 (메모 1개)',
      '확장에만 있던 조건부 서식 1개를 기본 조건부 서식으로',
    ]);
    const a = await unzip(orig); const s = await unzip(saved); const b = await unzip(r.bytes);
    // 셀 글꼴: x2t가 테마 라틴 글꼴(Calibri)로 바꾼 이름을 원본 이름으로 (행 높이 17.4pt → 차트·그림이 찌그러지지 않게)
    assert.match(s.text('xl/styles.xml'), /<name val="Calibri"\/><scheme val="minor"\/>/);
    assert.equal((b.text('xl/styles.xml').match(/<name val="맑은 고딕"\/><charset val="129"\/><scheme val="minor"\/>/g) ?? []).length, 2);
    assert.doesNotMatch(b.text('xl/styles.xml'), /Calibri/);
    // 메모: 스레드 메모·사람 부분과 그 관계·내용 형식이 없고, 원래 메모는 원본 작성자·글(굵게 Tahoma) 그대로
    assert.ok(s.names.some((n) => /threadedComments/.test(n)));
    assert.ok(!b.names.some((n) => /threadedComments|persons/.test(n)), b.names.join(', '));
    assert.doesNotMatch(b.text('[Content_Types].xml'), /threadedcomments|person\+xml/i);
    assert.doesNotMatch(b.text('xl/_rels/workbook.xml.rels'), /relationships\/person/);
    assert.doesNotMatch(b.text('xl/worksheets/_rels/sheet1.xml.rels'), /threadedComment/);
    const cm = b.text('xl/comments1.xml');
    assert.match(cm, /<authors><author>이작성<\/author><\/authors>/);
    assert.ok(cm.includes(a.text('xl/comments1.xml').match(/<text>[\s\S]*?<\/text>/)[0]), cm);
    assert.doesNotMatch(cm, /tc=\{|이작성:/);
    // 조건부 서식: 기본 conditionalFormatting(서식은 styles.xml dxfs로), 확장은 비어서 뺀다
    const sheet = b.text('xl/worksheets/sheet1.xml');
    assert.ok(sheet.includes('</sheetData><conditionalFormatting sqref="D2:D4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1"><formula>150</formula></cfRule></conditionalFormatting><printOptions'), sheet);
    assert.doesNotMatch(sheet, /x14:conditionalFormatting|<extLst>/);
    assert.match(b.text('xl/styles.xml'), /<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF9999"\/><bgColor rgb="FFFF9999"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/);
    for (const n of b.names) if (/\.(xml|rels)$/.test(n)) assert.ok(xmlBalanced(b.text(n)), n);
    // 다시 고쳐도 그대로
    const again = await fixOoxml(r.bytes, 'xlsx', { original: orig });
    assert.equal(again.changed, false);
  });

  test('sales: 고치기 전에는 Excel 없는 파일 엔진이 거절하고, 고친 뒤에는 고친다 (set_cells·insert_rows)', async () => {
    const before = join(tmp.dir, 'sales-before.xlsx');
    writeFileSync(before, saved);
    await assert.rejects(excelEdit(before, [{ op: 'set_cells', sheet: '매출', cells: { B5: 10 } }], { engine: 'file' }), /스레드 메모, 조건부 서식\(확장\)/);
    const p = join(tmp.dir, 'sales-after.xlsx');
    writeFileSync(p, (await fixOoxml(saved, 'xlsx', { original: orig })).bytes);
    await excelEdit(p, [{ op: 'set_cells', sheet: '매출', cells: { B5: 10 } }], { engine: 'file' });
    await excelEdit(p, [{ op: 'insert_rows', sheet: '매출', at: 3, count: 1 }], { engine: 'file' });
    const x = await unzip(readFileSync(p));
    assert.match(x.text('xl/worksheets/sheet1.xml'), /<conditionalFormatting sqref="D2:D5">/, '행 삽입 뒤 조건부 서식 범위도 밀림');
  });

  test('budget: 데이터 막대·아이콘·색조는 그대로 두고 cellIs·containsText·expression만 옮긴다 (서식 3개 → dxfs)', async () => {
    const r = await fixOoxml(CARRY('budget.saved.xlsx'), 'xlsx', { original: CARRY('budget.orig.xlsx') });
    assert.equal(r.error, null);
    assert.ok(r.fixes.includes('확장에만 있던 조건부 서식 3개를 기본 조건부 서식으로'), r.fixes.join('\n'));
    assert.ok(r.fixes.includes('셀 글꼴 4개 이름 원본대로: 맑은 고딕'));
    const b = await unzip(r.bytes);
    const sheet = b.text('xl/worksheets/sheet1.xml');
    assert.match(sheet, /<cfRule type="containsText" dxfId="1" priority="4" operator="containsText" stopIfTrue="1" text="초과"><formula>NOT\(ISERROR\(SEARCH\(&quot;초과&quot;,F4\)\)\)<\/formula><\/cfRule>/);
    assert.match(sheet, /<cfRule type="expression" dxfId="2" priority="5" stopIfTrue="1"><formula>\$E4&gt;1<\/formula><\/cfRule>/);
    assert.equal((sheet.match(/<x14:cfRule\b/g) ?? []).length, 1, '데이터 막대만 확장에 남음');
    assert.match(sheet, /<x14:cfRule type="dataBar"/);
    assert.match(sheet, /<cfRule type="iconSet" priority="6">/);
    assert.match(b.text('xl/styles.xml'), /<dxfs count="3"><dxf><fill>[\s\S]*<\/dxf><dxf><font><b\/><i val="0"\/><color indexed="2"\/><\/font><\/dxf><dxf><fill>/);
  });

  test('원본에도 스레드 메모가 있으면 메모는 그대로, 다른 시트를 가리키는 조건부 서식은 확장에 남긴다', async () => {
    const o2 = await mutate(orig, async (z) => { z.file('xl/threadedComments/threadedComment1.xml', '<ThreadedComments/>'); });
    const s2 = await mutate(saved, async (z, read) => {
      z.file('xl/worksheets/sheet1.xml', (await read('xl/worksheets/sheet1.xml')).replace('<xm:f>150</xm:f>', "<xm:f>'다른 시트'!$A$1</xm:f>"));
    });
    const r = await fixOoxml(s2, 'xlsx', { original: o2 });
    assert.equal(r.error, null);
    assert.ok(!r.fixes.some((f) => f.includes('메모') || f.includes('조건부 서식')), r.fixes.join('\n'));
    const b = await unzip(r.bytes);
    assert.ok(b.names.includes('xl/threadedComments/threadedComment1.xml'));
    assert.match(b.text('xl/worksheets/sheet1.xml'), /<x14:cfRule type="cellIs"/);
  });
});

// carry/chart.orig.xlsx = 막대 차트(항목 A4:A9, 값 B4:B9·G4:G9) 문서, 편집기에서 9행에 행을 넣고 저장한 것:
//   chart-rowhead.saved.xlsx  행 머리글 → '위의 행 삽입': 값 범위만 B4:B10으로 늘고 항목은 A4:A9 (Excel에서 7번째 막대 이름이 빔)
//   chart-cell.saved.xlsx     셀 하나 고르고 행 삽입: 범위가 둘 다 그대로 (마지막 행 '홍보담당관'이 차트에서 빠짐)
const chartRefs = (xml) => [...xml.matchAll(/<c:(cat|val)>[\s\S]*?<c:f>([^<]*)<\/c:f>[\s\S]*?<c:ptCount val="(\d+)"/g)].map((m) => `${m[1]} ${m[2]} ${m[3]}`);
const lastCatLabel = (xml) => [...xml.match(/<c:cat>[\s\S]*?<\/c:cat>/)[0].matchAll(/<c:v>([^<]*)<\/c:v>/g)].at(-1)[1];

describe('xlsx — 차트 범위 (편집기에서 데이터 범위 안에 행을 넣은 뒤)', () => {
  const orig = CARRY('chart.orig.xlsx');
  const want = ['cat 집행현황!$A$4:$A$10 7', 'val 집행현황!$B$4:$B$10 7', 'cat 집행현황!$A$4:$A$10 7', 'val 집행현황!$G$4:$G$10 7'];

  test('값 범위만 늘었으면 항목 범위를 값에 맞춘다 (원본 없이도) + 항목 캐시도 시트 값으로', async () => {
    const saved = CARRY('chart-rowhead.saved.xlsx');
    assert.deepEqual(chartRefs((await unzip(saved)).text('xl/charts/chart1.xml')),
      ['cat 집행현황!$A$4:$A$9 6', 'val 집행현황!$B$4:$B$10 7', 'cat 집행현황!$A$4:$A$9 6', 'val 집행현황!$G$4:$G$10 7']);
    for (const original of [null, orig]) {
      const r = await fixOoxml(saved, 'xlsx', { original });
      assert.equal(r.error, null);
      assert.deepEqual(r.warnings, []);
      assert.deepEqual(r.fixes.filter((f) => f.startsWith('차트')), [
        '차트 범위 맞춤: xl/charts/chart1.xml 집행현황!$A$4:$A$9 → 집행현황!$A$4:$A$10',
        '차트 범위 맞춤: xl/charts/chart1.xml 집행현황!$A$4:$A$9 → 집행현황!$A$4:$A$10',
      ]);
      const chart = (await unzip(r.bytes)).text('xl/charts/chart1.xml');
      assert.deepEqual(chartRefs(chart), want);
      assert.equal(lastCatLabel(chart), '홍보담당관');
      assert.ok(xmlBalanced(chart));
      // 손대지 않은 부분은 그대로 (시트·스타일 등)
      assertUntouched(await unzip(saved), await unzip(r.bytes), ['xl/charts/chart1.xml', ...(original ? ['xl/theme/theme1.xml', 'xl/styles.xml'] : [])]);
      assert.equal((await fixOoxml(r.bytes, 'xlsx', { original })).fixes.filter((f) => f.startsWith('차트')).length, 0, '다시 고쳐도 그대로');
    }
  });

  test('범위가 그대로면 원본 항목 글을 따라 범위를 옮긴다 (원본이 있을 때만)', async () => {
    const saved = CARRY('chart-cell.saved.xlsx');
    const before = (await unzip(saved)).text('xl/charts/chart1.xml');
    assert.deepEqual(chartRefs(before), ['cat 집행현황!$A$4:$A$9 6', 'val 집행현황!$B$4:$B$9 6', 'cat 집행현황!$A$4:$A$9 6', 'val 집행현황!$G$4:$G$9 6']);
    // 원본이 없으면 판단할 근거가 없다 → 그대로
    const bare = await fixOoxml(saved, 'xlsx');
    assert.equal(bare.fixes.filter((f) => f.startsWith('차트')).length, 0);
    const r = await fixOoxml(saved, 'xlsx', { original: orig });
    assert.equal(r.error, null);
    assert.deepEqual(r.warnings, []);
    const chart = (await unzip(r.bytes)).text('xl/charts/chart1.xml');
    assert.deepEqual(chartRefs(chart), want);
    assert.equal(lastCatLabel(chart), '홍보담당관');
    // 계열 이름(B3·G3)은 그 자리 그대로
    assert.match(chart, /<c:tx><c:strRef><c:f>집행현황!\$B\$3<\/c:f>/);
  });

  test('원본에도 같은 범위 글로 있던 길이 차이는 사용자가 만든 것 → 그대로, 자리가 어긋나면 경고만', async () => {
    const odd = await mutate(CARRY('chart-rowhead.saved.xlsx'), async (z, read) => {
      z.file('xl/charts/chart1.xml', (await read('xl/charts/chart1.xml')).replaceAll('집행현황!$A$4:$A$9', '집행현황!$A$5:$A$9'));
    });
    const r = await fixOoxml(odd, 'xlsx');
    assert.equal(r.error, null);
    assert.equal(r.fixes.filter((f) => f.startsWith('차트')).length, 0);
    assert.equal(r.warnings.length, 2);
    assert.match(r.warnings[0], /^차트 항목·값 범위 길이가 달라도 자리가 어긋나 그대로 둠: xl\/charts\/chart1\.xml/);
    // 원본(같은 범위 글) = 저장본 그 자체 → 이미 있던 차이
    const same = await fixOoxml(CARRY('chart-rowhead.saved.xlsx'), 'xlsx', { original: CARRY('chart-rowhead.saved.xlsx') });
    assert.equal(same.fixes.filter((f) => f.startsWith('차트')).length, 0);
  });

  test('범위 글 읽기: 따옴표 시트 이름·한 칸·여러 영역', async () => {
    const { parseChartRef } = await import('../src/core/ooxml-fixups.mjs');
    assert.deepEqual(parseChartRef("'My ''Data'''!$B$4:$B$10"), { sheet: "My 'Data'", token: "'My ''Data'''", c1: 2, c2: 2, r1: 4, r2: 10 });
    assert.deepEqual(parseChartRef('Sheet1!$AA$3'), { sheet: 'Sheet1', token: 'Sheet1', c1: 27, c2: 27, r1: 3, r2: 3 });
    assert.equal(parseChartRef('(Sheet1!$A$1:$A$3,Sheet1!$A$5:$A$6)'), null);
    assert.equal(parseChartRef('이름정의'), null);
  });
});

describe('원본에서 옮겨 오기 — pptx (테마 글꼴·구역)', () => {
  // blank-raw.pptx: 슬라이드 256·257. 원본에 구역 둘(앞·뒤)을 넣은 것
  const SECTIONS = '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    + '<p14:section name="도입" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section>'
    + '<p14:section name="계획" id="{00000000-0000-0000-0000-000000000002}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section>'
    + '</p14:sectionLst></p:ext></p:extLst>';
  const origP = () => mutate(PPTX, async (z, read) => {
    z.file('ppt/presentation.xml', (await read('ppt/presentation.xml')).replace('</p:presentation>', `${SECTIONS}</p:presentation>`));
  });
  const slides = (ids) => async (z, read) => {
    const x = await read('ppt/presentation.xml');
    const list = ids.map((id, i) => `<p:sldId id="${id}" r:id="${id === '256' ? 'rId3' : id === '257' ? 'rId4' : `rId${90 + i}`}"></p:sldId>`).join('');
    z.file('ppt/presentation.xml', x.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${list}</p:sldIdLst>`));
  };
  const sectionsOf = (xml) => [...xml.matchAll(/<p14:section name="([^"]+)"[^>]*><p14:sldIdLst>([\s\S]*?)<\/p14:sldIdLst>/g)]
    .map((m) => `${m[1]}:${[...m[2].matchAll(/id="(\d+)"/g)].map((x) => x[1]).join(',')}`);

  test('x2t가 지운 구역을 되살린다 (테마 나누기도 함께)', async () => {
    const r = await fixOoxml(PPTX, 'pptx', { original: await origP() });
    assert.equal(r.error, null);
    assert.ok(r.fixes.includes('구역 2개 원본대로'), r.fixes.join('\n'));
    assert.ok(r.fixes.some((f) => f.startsWith('노트 마스터 테마 나눔')));
    const x = (await unzip(r.bytes)).text('ppt/presentation.xml');
    assert.deepEqual(sectionsOf(x), ['도입:256', '계획:257']);
    assert.match(x, /<\/p:defaultTextStyle><p:extLst><p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"><p14:sectionLst xmlns:p14=/);
    assert.ok(xmlBalanced(x));
    assert.deepEqual(await findOoxmlIssues(r.bytes, 'pptx'), []);
  });

  test('새 슬라이드는 앞 슬라이드의 구역으로, 지운 슬라이드는 빠지고 빈 구역은 남는다', async () => {
    const o = await origP();
    let r = await fixOoxml(await mutate(PPTX, slides(['256', '300', '257'])), 'pptx', { original: o });
    assert.deepEqual(sectionsOf((await unzip(r.bytes)).text('ppt/presentation.xml')), ['도입:256,300', '계획:257']);
    r = await fixOoxml(await mutate(PPTX, slides(['300', '257'])), 'pptx', { original: o });
    assert.deepEqual(sectionsOf((await unzip(r.bytes)).text('ppt/presentation.xml')), ['도입:', '계획:300,257'], '맨 앞 새 슬라이드는 다음 아는 슬라이드의 구역');
  });

  test('슬라이드 순서가 구역을 넘나들거나 원본과 겹치지 않으면 되살리지 않는다', async () => {
    const o = await origP();
    let r = await fixOoxml(await mutate(PPTX, slides(['257', '256'])), 'pptx', { original: o });
    assert.ok(r.fixes.includes('구역: 슬라이드 순서가 구역을 넘나들어 되살리지 않음'), r.fixes.join('\n'));
    assert.doesNotMatch((await unzip(r.bytes)).text('ppt/presentation.xml'), /sectionLst/);
    r = await fixOoxml(await mutate(PPTX, slides(['400', '401'])), 'pptx', { original: o });
    assert.ok(r.fixes.includes('구역: 원본 슬라이드와 겹치지 않아 되살리지 않음'), r.fixes.join('\n'));
  });
});

describe('부분 크기 한도 (압축 폭탄으로 주 프로세스 메모리를 다 쓰지 않게)', () => {
  // 중앙 디렉터리·로컬 머리글의 '풀린 크기' 칸 바꾸기
  function setDeclaredSize(zipBytes, name, size) {
    const b = Buffer.from(zipBytes);
    const nm = Buffer.from(name);
    for (let i = 0; i < b.length - 46; i++) {
      if (b.readUInt32LE(i) === 0x02014b50 && b.readUInt16LE(i + 28) === nm.length && b.subarray(i + 46, i + 46 + nm.length).equals(nm)) b.writeUInt32LE(size, i + 24);
      if (b.readUInt32LE(i) === 0x04034b50 && b.readUInt16LE(i + 26) === nm.length && b.subarray(i + 30, i + 30 + nm.length).equals(nm)) b.writeUInt32LE(size, i + 22);
    }
    return b;
  }
  let bomb;
  test('준비: 풀면 64MB를 넘는 word/document.xml', async () => {
    assert.equal(PART_MAX_BYTES, 64 * 1024 * 1024);
    bomb = await mutate(DOCX, async (z) => { z.file('word/document.xml', Buffer.alloc(PART_MAX_BYTES + 1024, 0x20)); });
    assert.ok(bomb.length < 2 * 1024 * 1024, `압축하면 작다 (${bomb.length})`);
  });

  test('적힌 크기가 한도를 넘으면 풀지 않는다 → 원래 바이트 + 오류', async () => {
    const r = await fixOoxml(bomb, 'docx', { original: bomb });
    assert.equal(r.changed, false);
    assert.ok(r.bytes.equals(bomb));
    assert.match(r.error, /부분이 너무 큼: word\/document\.xml/);
  });

  test('크기를 속여도 (작게 적어도) 풀면서 한도에서 멈춘다', async () => {
    const liar = setDeclaredSize(bomb, 'word/document.xml', 100);
    const r = await fixOoxml(liar, 'docx');
    assert.equal(r.changed, false);
    assert.ok(r.bytes.equals(liar));
    assert.match(r.error, /부분이 너무 큼: word\/document\.xml/);
  });
});

describe('Office 엄격 열기 (경고 모두 켬·복구 끔·감시 타이머) — OFFICEKIT_COM_TESTS=1 일 때만', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  const dir = tmp.path(`strict-${Date.now()}`); // 실행마다 새 경로: PowerPoint는 끝낸 적 있는 경로를 '심각한 오류' 목록에 둔다
  mkdirSync(dir, { recursive: true });
  let before = new Set();
  const write = (name, bytes) => { const p = join(dir, name); writeFileSync(p, bytes); return p; };
  const skipIfBusy = (t, { code, r }) => { if (code === 3) { t.skip(`Office가 바쁨: ${r?.error}`); return true; } return false; };

  test('준비: 지금 떠 있는 Office', async () => { before = await officePids(); });

  test('고친 pptx(새 문서)는 PowerPoint가 경고 없이 연다', { timeout: 600_000 }, async (t) => {
    const res = await strictOpen(write('fixed.pptx', (await fixOoxml(PPTX, 'pptx')).bytes));
    if (skipIfBusy(t, res)) return;
    assert.equal(res.code, 0, JSON.stringify(res.r));
    assert.equal(res.r.hung, false);
    assert.equal(res.r.slides.length, 2);
    assert.ok(res.r.slides.flatMap((s) => s.shapes).some((s) => s.text === '한글 입력 확인'));
  });

  test('고치지 않은 pptx(새 문서)는 PowerPoint 복구 창에서 멈춘다 → 감시 타이머가 그 PowerPoint만 끝냄', { timeout: 600_000 }, async (t) => {
    const res = await strictOpen(write('unfixed.pptx', PPTX), 30);
    if (skipIfBusy(t, res)) return;
    assert.equal(res.code, 1, JSON.stringify(res.r));
    assert.equal(res.r.hung, true);
    assert.equal(res.r.opened, false);
    assert.equal(res.r.watchdog, 'killed');
    // 멈춘 창의 글 (UI Automation): '... can attempt to repair the presentation' / 한국어 UI는 '복구'
    const texts = (res.r.prompt ?? []).flatMap((w) => w.texts ?? []).join('\n');
    if (texts) assert.match(texts, /repair|복구/i, texts);
    else t.diagnostic(`멈춘 창 글을 읽지 못함: ${res.r.promptError ?? '창 없음'}`);
  });

  test('고친 docx: Word가 경고 없이 열고 표가 고르다 (Rows.Item 됨), 고치기 전은 세로 병합으로 봄', { timeout: 600_000 }, async (t) => {
    const fixed = await strictOpen(write('fixed.docx', (await fixOoxml(DOCX, 'docx')).bytes));
    if (skipIfBusy(t, fixed)) return;
    assert.equal(fixed.code, 0, JSON.stringify(fixed.r));
    const [table] = fixed.r.tables;
    assert.deepEqual({ rows: table.rows, uniform: table.uniform, rowsItem: table.rowsItem }, { rows: 4, uniform: true, rowsItem: true });
    assert.ok(table.cells.includes('r4c1=총무팀') && table.cells.includes('r4c3=3,300'), table.cells.join(' '));
    // 고치기 전: Word가 새 행을 세로 병합으로 봐서 Office Kit의 표 작업(Rows.Item)이 실패한다
    const raw = await strictOpen(write('raw.docx', DOCX));
    if (skipIfBusy(t, raw)) return;
    assert.equal(raw.code, 0, JSON.stringify(raw.r));
    assert.equal(raw.r.tables[0].uniform, false);
    assert.equal(raw.r.tables[0].rowsItem, false);
  });

  test('고친 xlsx: Excel이 경고 없이 열고 수식이 그대로', { timeout: 600_000 }, async (t) => {
    const res = await strictOpen(write('fixed.xlsx', (await fixOoxml(XLSX, 'xlsx')).bytes));
    if (skipIfBusy(t, res)) return;
    assert.equal(res.code, 0, JSON.stringify(res.r));
    assert.ok(res.r.sheets[0].rows.some((row) => row.includes('{=SUM(B15:D15)}')), res.r.sheets[0].rows.join('\n'));
  });

  test('원본에서 옮긴 docx·xlsx(메모·조건부 서식)·pptx(구역)도 경고 없이 열린다', { timeout: 900_000 }, async (t) => {
    const cases = [
      ['carried.docx', await fixOoxml(CARRY('minutes.saved.docx'), 'docx', { original: CARRY('minutes.orig.docx') })],
      ['carried-sales.xlsx', await fixOoxml(CARRY('sales.saved.xlsx'), 'xlsx', { original: CARRY('sales.orig.xlsx') })],
      ['carried-budget.xlsx', await fixOoxml(CARRY('budget.saved.xlsx'), 'xlsx', { original: CARRY('budget.orig.xlsx') })],
    ];
    const sec = await mutate(PPTX, async (z, read) => {
      z.file('ppt/presentation.xml', (await read('ppt/presentation.xml')).replace('</p:presentation>', '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="도입" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section><p14:section name="계획" id="{00000000-0000-0000-0000-000000000002}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>'));
    });
    cases.push(['carried.pptx', await fixOoxml(PPTX, 'pptx', { original: sec })]);
    for (const [name, r] of cases) {
      assert.equal(r.error, null, name);
      assert.equal(r.changed, true, name);
      const res = await strictOpen(write(name, r.bytes));
      if (skipIfBusy(t, res)) return;
      assert.equal(res.code, 0, `${name}: ${JSON.stringify(res.r)}`);
      assert.equal(res.r.hung, false, name);
    }
  });

  test('차트 범위를 맞춘 xlsx: Excel에서 7번째 막대에 항목 이름이 붙는다 (고치기 전은 빈 이름)', { timeout: 600_000 }, async () => {
    const ps = join(dir, 'chart-labels.ps1');
    writeFileSync(ps, `param([string]$Path, [string]$Out)
$m = New-Object System.Threading.Mutex($false, 'Local\\OfficeKitCom'); try { [void]$m.WaitOne() } catch [System.Threading.AbandonedMutexException] { }
$x = New-Object -ComObject Excel.Application
try {
  $x.DisplayAlerts = $false
  $wb = $x.Workbooks.Open($Path, 0, $true)
  $res = @(foreach ($s in $wb.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection()) { $xv = @($s.XValues); @{ n = @($s.Values).Count; labels = @($xv | ForEach-Object { [string]$_ }) } })
  $wb.Close($false)
  [System.IO.File]::WriteAllText($Out, (ConvertTo-Json -InputObject $res -Depth 4 -Compress), (New-Object System.Text.UTF8Encoding($false)))
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers(); $m.ReleaseMutex() }
`, 'ascii');
    const labels = (file) => new Promise((done, fail) => execFile(POWERSHELL, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps, file, `${file}.json`],
      { windowsHide: true, timeout: 300_000 }, (err, _o, se) => (err ? fail(new Error(`${err.message} ${se}`)) : done(JSON.parse(readFileSync(`${file}.json`, 'utf8'))))));
    const raw = await labels(write('chart-raw.xlsx', CARRY('chart-rowhead.saved.xlsx')));
    assert.deepEqual(raw.map((s) => [s.n, s.labels.length, s.labels.at(-1)]), [[7, 7, ''], [7, 7, '']]);
    const fixed = await labels(write('chart-fixed.xlsx', (await fixOoxml(CARRY('chart-rowhead.saved.xlsx'), 'xlsx', { original: CARRY('chart.orig.xlsx') })).bytes));
    assert.deepEqual(fixed.map((s) => [s.n, s.labels.at(-1)]), [[7, '홍보담당관'], [7, '홍보담당관']]);
    const followed = await labels(write('chart-followed.xlsx', (await fixOoxml(CARRY('chart-cell.saved.xlsx'), 'xlsx', { original: CARRY('chart.orig.xlsx') })).bytes));
    assert.deepEqual(followed[0].labels, ['총무과', '기획예산과', '정보통신과', '시설관리과', '민원봉사과', '감사담당관', '홍보담당관']);
  });

  test('끝: 시험이 띄운 Office가 남지 않았다', { timeout: 60_000 }, async () => {
    let left = [];
    for (let i = 0; i < 20; i++) {
      left = [...await officePids()].filter((p) => !before.has(p));
      if (!left.length) break;
      await new Promise((r) => setTimeout(r, 500));
    }
    assert.deepEqual(left, [], `남은 Office PID: ${left.join(', ')}`);
  });
});
