// ppt.test.mjs — PowerPoint 전용 기능 (ppt_inspect / ppt_edit, docs/FEATURES.md 'PowerPoint').
// 작업 검사·오류 문구·MCP 등록은 항상, PowerPoint(COM) 시험은 OFFICEKIT_COM_TESTS=1 일 때만.
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { copyFileSync, existsSync, readdirSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import {
  validatePptOps, pptErrorMessage, pptEdit, pptInspect, formatPptInspect, effectName, pptProtection,
  OP_NAMES, PPT_OPS_HELP, LAYOUT_KEYS, LOCKED_OPEN, LOCKED_PPT, LOCKED_WRITE,
} from '../src/core/ppt.mjs';
import { createDocument, readDocument } from '../src/core/index.mjs';
import { pageImage } from '../src/core/render.mjs';
import { MANUSCRIPT, COM_TESTS, ROOT, tempDir, isPng, pngSize } from './helpers/fixtures.mjs';

const tmp = tempDir('officekit-ppt-');
after(() => tmp.cleanup());

const ICON = join(ROOT, 'build', 'icon.png');
const SERVER = join(ROOT, 'src', 'mcp', 'server.mjs');
const FAKE_PNG = tmp.path('fake.png');
writeFileSync(FAKE_PNG, 'not an image');

// 쓰기 암호(modifyVerifier)가 걸린 pptx: PowerPoint가 Presentation.WritePassword로 쓰는 요소 그대로
const VERIFIER = '<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="14" '
  + 'spinCount="100000" saltData="NwtmvC7if0X5YLnw+kkhFw==" hashData="jXvlR6kDTRmFuD4UvmYAfc7G7G/ZgaqqQTF3XvRkmuPbxX6Q6TzAU5QmJ4dzMtBin/j9dFRmJlUZoM+UnWbFTQ=="/>';
async function withVerifier(src, out) {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(readFileSync(src));
  const xml = await zip.file('ppt/presentation.xml').async('string');
  zip.file('ppt/presentation.xml', xml.replace('</p:presentation>', `${VERIFIER}</p:presentation>`));
  writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer' }));
}
// 복합 문서(CFB) 머리 + 내용 조각
const CFB_HEAD = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const cfb = (...parts) => Buffer.concat([CFB_HEAD, Buffer.alloc(504), ...parts, Buffer.alloc(64)]);

// 모든 작업을 한 번씩 (COM 시험과 검사 시험이 함께 쓴다). 슬라이드 번호는 MANUSCRIPT로 만든 2장짜리 문서 기준
const ALL_OPS = [
  { op: 'add_slide', layout: 'title', position: 1, title: '연간 보고', body: '부제목', notes: '표지 노트' }, // 1 → 3장
  { op: 'add_slide', layout: 'title_content', title: '요약', body: [{ text: '첫째 항목' }, { text: '하위 항목', level: 2 }, '셋째 항목'] }, // 4
  { op: 'add_slide', layout: 'two_content', title: '비교' }, // 5
  { op: 'set_body', slide: 5, bullets: ['오른쪽 내용'], placeholder: 2 },
  { op: 'set_title', slide: 5, text: '비교 (수정)' },
  { op: 'set_notes', slide: 4, text: '요약 노트\n둘째 줄' },
  { op: 'set_text', slide: 2, shape: 'Text 2', text: '담당자: 김철수' },
  { op: 'add_textbox', slide: 4, text: '비율 글상자', left: '10%', top: '80%', width: '50%', height: 30, size: 18, bold: true, color: '#C00000', align: 'center', font: '맑은 고딕' },
  { op: 'add_image', slide: 4, image_path: ICON, left: '75%', top: '55%', width: 100 },
  { op: 'add_table', slide: 3, rows: [['이름', '점수'], ['가', 1]], left: 60, top: 200, width: 300 },
  { op: 'table_set_cell', slide: 3, shape: 3, row: 2, col: 2, text: '1점' },
  { op: 'table_set_cell', slide: 2, shape: 'Table 0', row: 2, col: 2, text: '999' },
  { op: 'add_shape', slide: 3, shape: 'rounded_rectangle', left: 450, top: 200, width: 200, height: 80, fill: '#2E75B6', line: 'none', text: '도형 글' },
  { op: 'add_shape', slide: 3, shape: 'line', left: 60, top: 480, width: 800, height: 0, line: '#FF0000' },
  { op: 'add_slide', layout: 'blank' }, // 6
  { op: 'add_chart', slide: 6, type: 'column', categories: ['1분기', '2분기', '3분기'], series: [{ name: '매출', values: [10, 20, 30] }, { name: '비용', values: [5, 8, 12] }], left: '5%', top: '5%', width: '90%', height: '85%', title: '분기 실적' },
  { op: 'format_text', find: '매출', all: true, color: '#FF0000', bold: true },
  { op: 'format_text', slide: 4, find: '첫째', italic: true, underline: true, size: 28, font: '맑은 고딕' },
  { op: 'delete_shape', slide: 3, shape: 'Text 1' },
  { op: 'set_background', slide: 3, color: '#FFF2CC' },
  { op: 'set_transition', slide: 'all', effect: 'fade', duration_s: 1 },
  { op: 'set_transition', slide: 2, effect: 'push' },
  { op: 'hide_slide', slide: 5 },
  { op: 'duplicate_slide', slide: 4 }, // 요약 복제 → 5, 비교 → 6, 차트 → 7
  { op: 'move_slide', slide: 5, to: 7 }, // 복제본을 맨 뒤로: 비교 5, 차트 6, 복제본 7
  { op: 'add_slide', layout: 'title_only', title: '지울 슬라이드' }, // 8
  { op: 'delete_slide', slide: 8 },
];

// ── 작업 검사 (PowerPoint 없이) ──────────────────────────

describe('ppt 작업 검사', () => {
  test('모든 작업이 스키마를 통과하고 설명에 나온다', () => {
    const extra = [{ op: 'apply_theme', theme_path: ICON.replace(/png$/, 'thmx') }, { op: 'set_slide_size', preset: '4:3' }];
    const seen = new Set([...ALL_OPS, ...extra].map((o) => o.op));
    assert.deepEqual([...seen].sort(), [...OP_NAMES].sort(), '시험 작업이 모든 op를 덮어야 함');
    const ops = validatePptOps(ALL_OPS);
    assert.equal(ops.length, ALL_OPS.length);
    for (const name of OP_NAMES) assert.ok(PPT_OPS_HELP.includes(`"${name}"`), name);
    for (const key of LAYOUT_KEYS) assert.ok(PPT_OPS_HELP.includes(key), key);
  });

  test('정리: 색·경로·표 값', () => {
    const [box, img, table, chart] = validatePptOps([
      { op: 'add_textbox', slide: 1, text: 'x', left: 0, top: '5 %', width: '50%', height: 10, color: 'ABCDEF' },
      { op: 'add_image', slide: 1, image_path: ICON, left: 0, top: 0 },
      { op: 'add_table', slide: 1, rows: [['a', 1], [2.5]], left: 0, top: 0, width: 100 },
      { op: 'add_chart', slide: 1, type: 'pie', categories: [2024, 2025], series: [{ name: 's', values: [1, 2] }], left: 0, top: 0, width: 10, height: 10 },
    ]);
    assert.equal(box.color, '#abcdef');
    assert.equal(img.image_path, ICON);
    assert.deepEqual(table.rows, [['a', '1'], ['2.5']]);
    assert.deepEqual(chart.categories, ['2024', '2025']);
  });

  const bad = [
    [[{ op: 'explode' }], /1번째 작업\(explode\): 알 수 없는 작업\(op\)입니다. 가능: add_slide/],
    [[{ op: 'add_slide' }, { op: 'add_slide', layout: 'title', foo: 1 }], /2번째 작업\(add_slide\): 인식할 수 없는 키: "foo"/],
    [[{ op: 'add_textbox', slide: 1, text: 'x', left: 'abc', top: 0, width: 10, height: 10 }], /'left' pt 숫자 또는 '10%'/],
    [[{ op: 'set_title', slide: 0, text: 'x' }], /'slide' .*1 이상/],
    [[{ op: 'set_background', slide: 'everything', color: '#ff0000' }], /슬라이드 번호\(1부터\) 또는 'all'/],
    [[{ op: 'add_shape', slide: 1, shape: 'oval', left: 0, top: 0, width: 10, height: 10, fill: 'red' }], /'fill' 색은 '#rrggbb'/],
    [[{ op: 'add_chart', slide: 1, type: 'column', categories: ['a', 'b'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: 9, height: 9 }], /계열 값 개수가 항목\(categories\) 2개와 다릅니다: s\(1개\)/],
    [[{ op: 'add_image', slide: 1, image_path: join(tmp.dir, 'none.png'), left: 0, top: 0 }], /그림 파일이 없습니다/],
    [[{ op: 'add_image', slide: 1, image_path: join(ROOT, 'package.json'), left: 0, top: 0 }], /png·jpg·jpeg·gif·bmp 파일만/],
    [[{ op: 'apply_theme', theme_path: ICON }], /\.thmx 또는 \.potx/],
    [[{ op: 'format_text', find: '가' }], /바꿀 서식/],
    [[{ op: 'add_shape', slide: 1, shape: 'line', left: 0, top: 0, width: 10, height: 0, text: '글' }], /선\(line\)에는 글을 넣을 수 없습니다/],
    [[{ op: 'add_shape', slide: 1, shape: 'oval', left: 0, top: 0, width: 0, height: 10 }], /width는 0보다 커야/],
    [[{ op: 'set_body', slide: 1, bullets: [{ text: 'a', level: 12 }] }], /'bullets\.0\.level' .*9 이하/],
    [[], /작업이 하나 이상/],
    // 0·음수 비율 크기
    [[{ op: 'add_image', slide: 1, image_path: ICON, left: 0, top: 0, width: '0%' }], /'width' 0보다 커야 합니다/],
    [[{ op: 'add_table', slide: 1, rows: [['a']], left: 0, top: 0, width: '-50%' }], /'width' 0보다 커야 합니다/],
    [[{ op: 'add_chart', slide: 1, type: 'pie', categories: ['a'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: '-5%', height: 10 }], /'width' 0보다 커야/],
    [[{ op: 'add_textbox', slide: 1, text: 'x', left: 0, top: 0, width: 10, height: -3 }], /'height' 0보다 커야/],
    // 확장자만 그림인 파일
    [[{ op: 'add_image', slide: 1, image_path: FAKE_PNG, left: 0, top: 0 }], /그림 파일이 아니거나 손상됐습니다/],
    [[{ op: 'add_shape', slide: 1, shape: 'oval', left: 0, top: 0, width: 10, height: 10, color: '#ff0000' }], /color\(글자색\)는 text와 함께/],
  ];
  for (const [ops, re] of bad) {
    test(`거부: ${re.source.slice(0, 40)}`, () => assert.throws(() => validatePptOps(ops), re));
  }

  test('pptEdit: PowerPoint를 띄우기 전에 형식·경로 확인', async () => {
    const doc = tmp.path('not-ppt.docx');
    writeFileSync(doc, 'x');
    await assert.rejects(pptEdit(doc, [{ op: 'add_slide' }]), /PowerPoint 문서\(\.pptx\/\.ppt\/\.odp\)가 아닙니다/);
    await assert.rejects(pptEdit(tmp.path('none.pptx'), [{ op: 'add_slide' }]), /파일이 없습니다/);
    const p = tmp.path('fake.pptx');
    writeFileSync(p, 'x');
    await assert.rejects(pptEdit(p, [{ op: 'add_slide' }], { output: tmp.path('out.pdf') }), /\.pptx\/\.ppt\/\.odp로만 저장/);
    await assert.rejects(pptEdit(p, [{ op: 'nope' }]), /1번째 작업\(nope\)/);
    await assert.rejects(pptInspect(doc), /PowerPoint 문서/);
  });

  test('테마 파일: 암호·가짜 거절', () => {
    const enc = tmp.path('enc.thmx');
    writeFileSync(enc, cfb(Buffer.from('EncryptedPackage', 'utf16le')));
    assert.throws(() => validatePptOps([{ op: 'apply_theme', theme_path: enc }]), /암호가 걸린 테마 파일은 쓸 수 없습니다/);
    const fake = tmp.path('fake.potx');
    writeFileSync(fake, 'not a theme');
    assert.throws(() => validatePptOps([{ op: 'apply_theme', theme_path: fake }]), /테마 파일이 아니거나 손상됐습니다/);
  });

  test('암호 문서는 PowerPoint를 띄우기 전에 거절 (보이지 않는 암호 창에서 멈추지 않게)', async () => {
    // 열기 암호 pptx(암호화된 꾸러미), .ppt 암호(Current User headerToken), 암호화된 odp
    const openPw = tmp.path('lock-open.pptx');
    writeFileSync(openPw, cfb(Buffer.from('EncryptionInfo', 'utf16le'), Buffer.from('EncryptedPackage', 'utf16le')));
    const pptPw = tmp.path('lock.ppt');
    writeFileSync(pptPw, cfb(Buffer.from([0x00, 0x00, 0xf6, 0x0f, 0x18, 0x00, 0x00, 0x00, 0x14, 0x00, 0x00, 0x00, 0xdf, 0xc4, 0xd1, 0xf3])));
    const pptPlain = tmp.path('plain.ppt');
    writeFileSync(pptPlain, cfb(Buffer.from([0x00, 0x00, 0xf6, 0x0f, 0x18, 0x00, 0x00, 0x00, 0x14, 0x00, 0x00, 0x00, 0x5f, 0xc0, 0x91, 0xe3])));
    const { default: JSZip } = await import('jszip');
    const odp = new JSZip();
    odp.file('META-INF/manifest.xml', '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>');
    const odpPw = tmp.path('lock.odp');
    writeFileSync(odpPw, await odp.generateAsync({ type: 'nodebuffer' }));
    const plain = tmp.path('lock-plain.pptx');
    await createDocument(plain, MANUSCRIPT);
    const writePw = tmp.path('lock-write.pptx');
    await withVerifier(plain, writePw);

    assert.equal(await pptProtection(openPw), 'open');
    assert.equal(await pptProtection(pptPw), 'ppt');
    assert.equal(await pptProtection(pptPlain), null);
    assert.equal(await pptProtection(odpPw), 'open');
    assert.equal(await pptProtection(plain), null);
    assert.equal(await pptProtection(writePw), 'write');

    const fast = async (p, re) => { const t0 = Date.now(); await assert.rejects(p, re); assert.ok(Date.now() - t0 < 5000, `${Date.now() - t0}ms`); };
    await fast(pptInspect(openPw), (e) => e.message === LOCKED_OPEN);
    await fast(pptEdit(openPw, [{ op: 'add_slide' }], { output: tmp.path('lock-out.pptx') }), (e) => e.message === LOCKED_OPEN);
    await fast(pptInspect(pptPw), (e) => e.message === LOCKED_PPT);
    await fast(pptInspect(odpPw), (e) => e.message === LOCKED_OPEN);
    // 쓰기 암호: 원본에 바로 저장은 거절 (다른 파일로는 된다 — COM 시험)
    await fast(pptEdit(writePw, [{ op: 'add_slide' }]), (e) => e.message === LOCKED_WRITE);
    if (process.platform === 'win32') await fast(pptEdit(writePw, [{ op: 'add_slide' }], { output: writePw.toUpperCase() }), (e) => e.message === LOCKED_WRITE);
    assert.equal(existsSync(tmp.path('lock-out.pptx')), false);
  });

  test('PowerShell 오류 → 한국어 (몇 번째 작업, 저장 안 함)', () => {
    assert.equal(pptErrorMessage('PowerPoint 작업 실패: opfail|3|set_title|E|slide_range|9|4'),
      '3번째 작업(set_title) 실패: 슬라이드 9번이 없습니다 (지금 4장) — 저장하지 않았습니다');
    assert.equal(pptErrorMessage('opfail|1|set_text|E|shape_not_found|2|제목 | 1'),
      "1번째 작업(set_text) 실패: 슬라이드 2에 '제목 | 1' 도형이 없습니다 (ppt_inspect로 도형 이름을 확인하세요) — 저장하지 않았습니다");
    assert.match(pptErrorMessage('opfail|2|add_chart|COM 구성 요소 호출에서 HRESULT E_FAIL'), /^2번째 작업\(add_chart\) 실패: PowerPoint 오류: COM 구성 요소/);
    assert.match(pptErrorMessage('PowerPoint 작업 실패: E|slide_range|5|2', { saving: false }), /^슬라이드 5번이 없습니다 \(지금 2장\)$/);
    assert.equal(pptErrorMessage('문서를 열지 못했습니다: x'), '문서를 열지 못했습니다: x');
    assert.equal(pptErrorMessage('opfail|1|add_textbox|The specified value is out of range.'),
      '1번째 작업(add_textbox) 실패: PowerPoint 오류: 값이 허용 범위를 벗어났습니다 (위치·크기·번호를 확인하세요) — 저장하지 않았습니다');
    assert.match(pptErrorMessage('opfail|1|add_textbox|E|size_not_positive|width|-10%'), /width 값은 0보다 커야 합니다: -10%/);
    assert.equal(pptErrorMessage('PowerPoint 작업 실패: Presentation.SaveAs : An error occurred while PowerPoint was saving the file.', { output: 'C:\\a\\b.pptx' }),
      '문서를 저장하지 못했습니다: C:\\a\\b.pptx — 저장할 파일이 다른 프로그램에서 열려 있거나 쓰기 권한이 없을 수 있습니다');
  });

  test('살펴보기 글', () => {
    const t = formatPptInspect({
      path: 'C:/a/발표.pptx', slideWidth: 960, slideHeight: 540, slideCount: 1, designs: ['Office 테마'],
      layouts: [{ design: 'Office 테마', index: 1, name: '제목 슬라이드', key: 'title' }, { design: 'Office 테마', index: 2, name: '사용자 지정', key: null }],
      slides: [{ index: 1, layout: '제목 슬라이드', key: 'title', title: '안녕', hidden: true, effect: 3849, duration: 1.5, notes: '노트\n둘째',
        shapes: [{ index: 1, name: '제목 1', type: 14, placeholder: 3, left: 120, top: 88.4, width: 720, height: 188, text: '안녕' },
          { index: 2, name: '표 3', type: 19, left: 10, top: 10, width: 300, height: 80, rows: 3, cols: 2, text: '이름 | 점수' },
          { index: 3, name: '차트 4', type: 3, left: 0, top: 0, width: 400, height: 300, chart: 51 }] }],
    });
    assert.match(t, /발표\.pptx — 슬라이드 1장, 크기 960×540pt \(16:9\)/);
    assert.match(t, /- 제목 슬라이드 \[title\]/);
    assert.match(t, /없는 키 title_content, .*새로 만든다/);
    assert.match(t, /\[슬라이드 1\] 레이아웃: 제목 슬라이드 \[title\], 숨김, 화면 전환: fade 1\.5초/);
    assert.match(t, /노트: 노트 \/ 둘째/);
    assert.match(t, /1\. '제목 1' 개체 틀\(가운데 제목\) @ 120,88 720×188pt: 안녕/);
    assert.match(t, /2\. '표 3' 표 3행×2열 @ 10,10 300×80pt: 이름 \| 점수/);
    assert.match(t, /3\. '차트 4' 차트\(세로 막대\)/);
    assert.equal(effectName(0), null);
    assert.equal(effectName(3855), 'push');
    assert.doesNotMatch(t, /이어서 보려면|보였습니다/, '나눠 보기가 아니면 끝 안내 없음');
  });

  test('살펴보기 글: 나눠 보기·도형 수 제한·글자 수 한도', () => {
    const base = { path: 'C:/a/큰.pptx', slideWidth: 960, slideHeight: 540, slideCount: 150, designs: ['Office Theme'], layouts: [], slide: null };
    const slide = (index, extra = {}) => ({ index, layout: 'DEFAULT', key: 'blank', title: `제목 ${index}`, hidden: false, effect: 0, duration: 0, notes: '',
      shapes: [{ index: 1, name: 'Title 1', type: 14, placeholder: 1, left: 0, top: 0, width: 10, height: 10, text: `제목 ${index}` }], shapeCount: 1, ...extra });
    let t = formatPptInspect({ ...base, offset: 0, limit: 30, slides: Array.from({ length: 30 }, (_, i) => slide(i + 1)) });
    assert.match(t, /\[슬라이드 30\]/);
    assert.match(t, /\(슬라이드 1~30만 보였습니다 \(전체 150장\) — 이어서 보려면 offset=30\)$/);
    t = formatPptInspect({ ...base, offset: 140, limit: 30, slides: Array.from({ length: 10 }, (_, i) => slide(i + 141)) });
    assert.match(t, /\(슬라이드 141~150만 보였습니다 \(전체 150장\)\)$/);
    t = formatPptInspect({ ...base, offset: 150, limit: 30, slides: [] });
    assert.match(t, /슬라이드는 150장뿐입니다 — offset은 149 이하로/);
    // 도형이 많은 슬라이드: 앞 60개만 + '…외 N개'
    const many = Array.from({ length: 60 }, (_, i) => ({ index: i + 1, name: `Box ${i + 1}`, type: 17, left: 0, top: 0, width: 1, height: 1 }));
    t = formatPptInspect({ ...base, slideCount: 1, offset: 0, limit: 30, slides: [slide(1, { shapes: many, shapeCount: 75 })] });
    assert.match(t, /도형 75개:/);
    assert.match(t, /…외 15개 \(slide=1로 이 슬라이드만 보면 500개까지 나옴\)/);
    // 글자 수 한도(약 6만 자)를 넘으면 그 앞에서 끊고 이어 볼 offset을 알려 준다
    const long = 'ㄱ'.repeat(4000);
    t = formatPptInspect({ ...base, offset: 0, limit: 30, slides: Array.from({ length: 30 }, (_, i) => slide(i + 1, { notes: long, shapes: [] , shapeCount: 0 })) });
    assert.ok(t.length <= 61_000, `${t.length}자`);
    const shown = t.match(/\[슬라이드 \d+\]/g).length;
    assert.ok(shown < 30, `${shown}장`);
    assert.match(t, new RegExp(`이어서 보려면 offset=${shown}\\)$`));
  });

  test('ppt-ops.ps1: ASCII만, 모든 작업을 처리', () => {
    const src = readFileSync(join(ROOT, 'scripts', 'com', 'ppt-ops.ps1'));
    assert.ok(src.every((b) => b < 0x80), 'PowerShell 5.1은 BOM 없는 스크립트를 ANSI로 읽는다');
    const text = src.toString('latin1');
    for (const name of OP_NAMES) assert.match(text, new RegExp(`'${name}' \\{`), name);
    for (const fn of ['Invoke-PptInspect', 'Invoke-PptEdit']) assert.match(text, new RegExp(`function ${fn}\\(\\$pres, \\$app, \\$req\\)`));
  });
});

// ── MCP 등록 (PowerPoint 없이) ───────────────────────────

describe('MCP ppt 도구', () => {
  let client;
  const call = (name, args) => client.callTool({ name, arguments: args }, undefined, { timeout: 400_000 });
  const textOf = (r) => r.content.filter((c) => c.type === 'text').map((c) => c.text).join('\n');
  before(async () => {
    mkdirSync(tmp.path('mcp'), { recursive: true });
    client = new Client({ name: 'office-kit-ppt-test', version: '1.0.0' });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [SERVER], cwd: tmp.path('mcp'), stderr: 'pipe' }));
  });
  after(async () => { await client?.close(); });

  test('ppt_inspect·ppt_edit 등록, 한국어 설명, 모든 작업이 스키마에', async () => {
    const { tools } = await client.listTools();
    const inspect = tools.find((t) => t.name === 'ppt_inspect');
    const edit = tools.find((t) => t.name === 'ppt_edit');
    assert.ok(inspect && edit, tools.map((t) => t.name).join(', '));
    assert.match(inspect.description, /레이아웃/);
    assert.equal(inspect.annotations?.readOnlyHint, true);
    assert.match(edit.description, /하나라도 실패하면 아무것도 저장하지 않는다/);
    const schema = JSON.stringify(edit.inputSchema);
    for (const name of OP_NAMES) {
      assert.ok(edit.description.includes(`"${name}"`), `설명에 ${name}`);
      assert.ok(schema.includes(`"const":"${name}"`), `스키마에 ${name}`);
    }
    assert.deepEqual(edit.inputSchema.required.sort(), ['operations', 'path']);
    assert.deepEqual(Object.keys(inspect.inputSchema.properties).sort(), ['limit', 'offset', 'path', 'slide']);
    assert.match(inspect.description, /30장까지\(limit\)/);
    assert.match(edit.description, /add_chart를 하는 동안 PowerPoint 창이/);
    assert.match(edit.description, /쪽 번호 = 슬라이드 번호, 숨긴 슬라이드도 포함/);
  });

  test('잘못된 입력은 한국어 오류', async () => {
    writeFileSync(tmp.path('mcp/a.pptx'), 'x');
    writeFileSync(tmp.path('mcp/a.docx'), 'x');
    let r = await call('ppt_edit', { path: 'a.pptx', operations: [{ op: 'explode' }] });
    assert.equal(r.isError, true);
    assert.match(textOf(r), /알 수 없는 작업\(op\)입니다/);
    r = await call('ppt_edit', { path: 'a.pptx', operations: [{ op: 'add_textbox', slide: 1, text: 'x', left: 'abc', top: 0, width: 1, height: 1 }] });
    assert.match(textOf(r), /pt 숫자 또는 '10%'/);
    r = await call('ppt_edit', { path: 'a.docx', operations: [{ op: 'add_slide' }] });
    assert.match(textOf(r), /ppt_edit는 PowerPoint 문서/);
    r = await call('ppt_edit', { path: 'a.pptx', operations: [{ op: 'add_slide' }], output: 'b.pdf' });
    assert.match(textOf(r), /저장 경로는 \.pptx, \.ppt 또는 \.odp/);
    r = await call('ppt_inspect', { path: 'a.docx' });
    assert.match(textOf(r), /ppt_inspect는 PowerPoint 문서/);
    // SDK 입력 검증도 한국어 (영어 'Too small' 없이)
    for (const [name, args, re] of [
      ['ppt_inspect', { path: 'a.pptx', slide: 0 }, /1 이상/],
      ['ppt_inspect', { path: 'a.pptx', limit: 500 }, /200 이하/],
      ['ppt_edit', { path: 'a.pptx', operations: [] }, /1개 이상/],
      ['ppt_edit', { path: 'a.pptx', operations: [{ op: 'add_textbox', slide: 1, text: 'x', left: 0, top: 0, width: '-10%', height: 10 }] }, /0보다 커야 합니다/],
    ]) {
      r = await call(name, args);
      assert.equal(r.isError, true, JSON.stringify(args));
      assert.match(textOf(r), re);
      assert.doesNotMatch(textOf(r), /Too small|Too big|expected/);
    }
    writeFileSync(tmp.path('mcp/fake.png'), 'not an image');
    r = await call('ppt_edit', { path: 'a.pptx', operations: [{ op: 'add_image', slide: 1, image_path: 'fake.png', left: 0, top: 0 }] });
    assert.match(textOf(r), /그림 파일이 아니거나 손상됐습니다/);
  });

  test('ppt_edit → ppt_inspect (PowerPoint)', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만', timeout: 600_000 }, async () => {
    await createDocument(tmp.path('mcp/deck.pptx'), MANUSCRIPT);
    let r = await call('ppt_edit', { path: 'deck.pptx', output: 'deck-out.pptx', operations: [
      { op: 'add_slide', layout: 'title_content', title: 'MCP 슬라이드', body: ['가', { text: '나', level: 2 }] },
      { op: 'add_image', slide: 3, image_path: ICON, left: '80%', top: '70%', height: '20%' },
    ] });
    assert.ok(!r.isError, textOf(r));
    assert.match(textOf(r), /1\. 슬라이드 3 추가/);
    assert.match(textOf(r), /저장: .*deck-out\.pptx/);
    r = await call('ppt_inspect', { path: 'deck-out.pptx', slide: 3 });
    assert.ok(!r.isError, textOf(r));
    assert.match(textOf(r), /\[슬라이드 3\] 레이아웃: .*\[title_content\]/);
    assert.match(textOf(r), /제목: MCP 슬라이드/);
    assert.match(textOf(r), /그림 @ \d+,378 \d+×108pt/); // 높이 20% = 108pt, 위 70% = 378pt
    assert.doesNotMatch(textOf(r), /\[슬라이드 1\]/);
  });
});

// ── 설치된 PowerPoint (COM) ──────────────────────────────

// 차트 데이터용 Excel (PowerPoint가 띄운 'EXCEL.EXE /Automation -Embedding /K')이 남았는지
function chartExcels() {
  return new Promise((done) => execFile('powershell.exe', ['-NoProfile', '-Command',
    "@(Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" | Where-Object { $_.CommandLine -match '-Embedding /K' } | ForEach-Object { $_.ProcessId }) -join ','"],
  { encoding: 'utf8', windowsHide: true }, (_e, out) => done(String(out ?? '').trim().split(',').filter(Boolean))));
}

async function slideXmls(path) {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(readFileSync(path));
  const names = Object.keys(zip.files).filter((n) => /^ppt\/(slides\/slide|charts\/chart)\d+\.xml$/.test(n));
  return Object.fromEntries(await Promise.all(names.map(async (n) => [n, await zip.file(n).async('string')])));
}

function findThemeFile() {
  for (const dir of ['C:/Program Files/Microsoft Office/root/Document Themes 16', 'C:/Program Files (x86)/Microsoft Office/root/Document Themes 16']) {
    if (!existsSync(dir)) continue;
    const names = readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.thmx'));
    const pick = names.find((n) => /^ion\.thmx$/i.test(n)) ?? names.find((n) => !/office/i.test(n)) ?? names[0];
    if (pick) return join(dir, pick);
  }
  return null;
}

describe('PowerPoint 편집 (COM)', { skip: !COM_TESTS && 'OFFICEKIT_COM_TESTS=1 일 때만' }, () => {
  const deck = () => tmp.path('deck.pptx');
  before(async () => { await createDocument(deck(), MANUSCRIPT); });

  test('모든 작업 → 살펴보기로 확인 + 쪽 그림', { timeout: 900_000 }, async () => {
    const excelBefore = await chartExcels();
    const out = tmp.path('all-out.pptx');
    const r = await pptEdit(deck(), ALL_OPS, { output: out });
    assert.equal(r.results.length, ALL_OPS.length);
    const lines = r.results.join('\n');
    assert.match(lines, /^1\. 슬라이드 1 추가 \(레이아웃: .+, 제목: 연간 보고\) — 전체 3장/m);
    assert.match(lines, /^16\. 슬라이드 6에 세로 막대 차트 '.+' 추가 \(항목 3개, 계열 2개\)/m);
    assert.match(lines, /^17\. '매출' \d+곳 서식 바꿈/m);
    assert.match(lines, /^24\. 슬라이드 4 복제 → 새 슬라이드 5 — 전체 7장/m);
    assert.match(lines, /^27\. 슬라이드 8 지움 — 전체 7장/m);
    // 차트 데이터용 Excel을 남기지 않는다
    assert.deepEqual((await chartExcels()).filter((p) => !excelBefore.includes(p)), [], '차트 데이터 Excel이 남음');

    const info = await pptInspect(out);
    assert.equal(info.slideCount, 7);
    assert.equal(info.slideWidth, 960);
    const [s1, s2, s3, s4, s5, s6, s7] = info.slides;
    assert.equal(s1.key, 'title');
    assert.equal(s1.title, '연간 보고');
    assert.equal(s1.notes, '표지 노트');
    assert.ok(s1.shapes.some((s) => s.placeholder === 4 && s.text === '부제목'));
    assert.equal(effectName(s1.effect), 'fade');
    assert.equal(s1.duration, 1);
    assert.equal(effectName(s2.effect), 'push');
    assert.ok(s2.shapes.some((s) => s.name === 'Text 2' && s.text === '담당자: 김철수'));
    assert.ok(!s3.shapes.some((s) => s.name === 'Text 1'), 'delete_shape');
    const table = s3.shapes.find((s) => s.rows === 2);
    assert.ok(table && table.cols === 2 && table.text === '이름 | 점수', JSON.stringify(s3.shapes));
    assert.ok(s3.shapes.some((s) => s.type === 1 && s.text === '도형 글'));
    assert.ok(s3.shapes.some((s) => s.type === 9), '선');
    assert.equal(s4.key, 'title_content');
    assert.equal(s4.notes, '요약 노트\n둘째 줄');
    const box = s4.shapes.find((s) => s.text === '비율 글상자');
    assert.ok(box, JSON.stringify(s4.shapes));
    assert.equal(box.left, 96); // 10% × 960
    assert.equal(box.top, 432); // 80% × 540
    assert.equal(box.width, 480);
    const pic = s4.shapes.find((s) => s.type === 13);
    assert.equal(pic?.width, 100);
    assert.equal(pic.left, 720);
    assert.ok(s4.shapes.some((s) => s.placeholder === 7 && s.text === '첫째 항목\n하위 항목\n셋째 항목'));
    assert.equal(s5.title, '비교 (수정)');
    assert.equal(s5.hidden, true);
    assert.equal(s5.key, 'two_content');
    assert.deepEqual(s5.shapes.filter((s) => s.placeholder === 7).map((s) => s.text ?? ''), ['', '오른쪽 내용']);
    const chart = s6.shapes.find((s) => s.chart);
    assert.equal(chart?.chart, 51);
    assert.equal(chart.text, '분기 실적');
    assert.equal(s7.title, '요약', '복제본이 맨 뒤로');
    assert.equal(s7.notes, '요약 노트\n둘째 줄');

    // 파일 안 확인: 서식·배경·숨김·차트 계열
    const xml = await slideXmls(out);
    const slides = Object.entries(xml).filter(([n]) => n.includes('/slides/')).map(([, x]) => x);
    assert.ok(slides.some((x) => /<p:bg>[\s\S]*FFF2CC/i.test(x)), '배경색');
    const runs = slides.flatMap((x) => [...x.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((m) => m[1]));
    const sales = runs.filter((r) => r.includes('<a:t>매출</a:t>'));
    assert.ok(sales.length >= 2, `매출 글자 조각 ${sales.length}개`);
    for (const run of sales) assert.match(run, /<a:rPr[^>]* b="1"[\s\S]*srgbClr val="FF0000"/, '매출 빨간 굵게');
    const firstRun = runs.find((r) => r.includes('<a:t>첫째</a:t>'));
    assert.match(firstRun ?? '', /<a:rPr[^>]* sz="2800"[^>]* i="1"|<a:rPr[^>]* i="1"[^>]* sz="2800"/, '첫째 서식');
    assert.ok(slides.some((x) => /<p:sld [^>]*show="0"/.test(x)), '숨김');
    assert.ok(slides.some((x) => /<p:push dir="u"\/>/.test(x)), 'push 전환');
    const charts = Object.entries(xml).filter(([n]) => n.includes('/charts/')).map(([, x]) => x);
    assert.equal(charts.length, 1);
    assert.equal((charts[0].match(/<c:ser>/g) ?? []).length, 2, '차트 계열 2개 (견본 3개가 남지 않음)');
    assert.ok(charts[0].includes('<c:v>매출</c:v>') && charts[0].includes('<c:v>3분기</c:v>'));

    // 글 (파일 엔진): 표 칸
    const t = (await readDocument(out, { engine: 'file' })).text;
    for (const s of ['999', '1점', '김철수', '오른쪽 내용']) assert.ok(t.includes(s), s);

    // 쪽 그림 (숨긴 슬라이드는 PDF에서 빠질 수 있음)
    const dir = process.env.OFFICEKIT_PPT_RENDER_DIR;
    if (dir) mkdirSync(dir, { recursive: true });
    const first = await pageImage(out, 1, { dpi: 60 });
    for (let p = 1; p <= first.pages; p++) {
      const img = p === 1 ? first : await pageImage(out, p, { dpi: 60 });
      assert.ok(isPng(img.png));
      const { width, height } = pngSize(img.png);
      assert.ok(Math.abs(width / height - 16 / 9) < 0.02, `${width}×${height}`);
      if (dir) writeFileSync(join(dir, `all-${p}.png`), img.png);
    }
  });

  test('하나라도 실패하면 저장하지 않음 (원본·출력 그대로)', { timeout: 600_000 }, async () => {
    const before = readFileSync(deck());
    const out = tmp.path('fail-out.pptx');
    await assert.rejects(pptEdit(deck(), [
      { op: 'add_slide', title: '새 슬라이드' },
      { op: 'set_title', slide: 99, text: 'x' },
    ], { output: out }), /^Error: 2번째 작업\(set_title\) 실패: 슬라이드 99번이 없습니다 \(지금 3장\) — 저장하지 않았습니다$/);
    assert.equal(existsSync(out), false);
    const excelBefore = await chartExcels();
    await assert.rejects(pptEdit(deck(), [
      { op: 'add_chart', slide: 1, type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: 300, height: 200 },
      { op: 'delete_slide', slide: 1 },
      { op: 'set_text', slide: 1, shape: '없는 도형', text: 'x' },
    ]), /3번째 작업\(set_text\) 실패: 슬라이드 1에 '없는 도형' 도형이 없습니다/);
    assert.ok(readFileSync(deck()).equals(before), '원본이 바뀜');
    assert.deepEqual((await chartExcels()).filter((p) => !excelBefore.includes(p)), [], '차트 데이터 Excel이 남음');
  });

  test('작업별 오류 문구', { timeout: 600_000 }, async () => {
    const cases = [
      [{ op: 'set_title', slide: 1, text: 'x' }, /슬라이드 1에 제목 개체 틀이 없습니다 — set_text/],
      [{ op: 'table_set_cell', slide: 1, shape: 'Text 0', row: 1, col: 1, text: 'x' }, /'Text 0'은\(는\) 표가 아닙니다/],
      [{ op: 'table_set_cell', slide: 1, shape: 'Table 0', row: 9, col: 1, text: 'x' }, /9행 1열 칸이 없습니다 \(표 3행×2열\)/],
      [{ op: 'set_text', slide: 1, shape: 'Table 0', text: 'x' }, /표입니다 — table_set_cell/],
      [{ op: 'format_text', find: '없는 글자', bold: true }, /찾는 글자가 없습니다: '없는 글자'/],
      [{ op: 'add_slide', layout: '없는 레이아웃' }, /레이아웃 '없는 레이아웃'이\(가\) 없습니다 \(키: title,.* 이 문서의 레이아웃: DEFAULT\)/],
      [{ op: 'move_slide', slide: 1, to: 5 }, /위치 5이\(가\) 범위를 벗어났습니다 \(1~2\)/],
      [{ op: 'set_text', slide: 1, shape: 9, text: 'x' }, /슬라이드 1에 9번 도형이 없습니다 \(도형 4개\)/],
    ];
    for (const [op, re] of cases) {
      await assert.rejects(pptEdit(deck(), [op], { output: tmp.path('err-out.pptx') }), (e) => {
        assert.match(e.message, new RegExp(`^1번째 작업\\(${op.op}\\) 실패: `));
        assert.match(e.message, re);
        return true;
      });
    }
    assert.equal(existsSync(tmp.path('err-out.pptx')), false);
  });

  test('쓰기 암호 pptx → 다른 파일로 저장 (쓰기 암호 유지) + 글상자 높이·도형 글자색·슬라이드 크기 안내', { timeout: 900_000 }, async () => {
    const src = tmp.path('wpw.pptx');
    await withVerifier(deck(), src);
    const before = readFileSync(src);
    const out = tmp.path('wpw-out.pptx');
    const r = await pptEdit(src, [
      { op: 'add_slide', layout: 'blank' }, // 3
      { op: 'add_textbox', slide: 3, text: '높이 30', left: 40, top: 40, width: 200, height: 30, size: 14 },
      { op: 'add_textbox', slide: 3, text: '아주 긴 글상자 글입니다. 높이 20pt에는 다 들어가지 않으므로 상자가 늘어나야 합니다.', left: 40, top: 120, width: 200, height: 20, size: 14 },
      { op: 'add_shape', slide: 3, shape: 'rectangle', left: 300, top: 40, width: 200, height: 80, fill: 'none', line: '#C00000', text: '대외비' },
      { op: 'add_shape', slide: 3, shape: 'rectangle', left: 300, top: 140, width: 200, height: 80, fill: '#FFF2CC', text: '밝은 채우기' },
      { op: 'add_shape', slide: 3, shape: 'oval', left: 300, top: 240, width: 200, height: 80, fill: 'none', text: '초록 글', color: '#00B050' },
      { op: 'add_image', slide: 3, image_path: ICON, left: 600, top: 40, height: '20%' },
      { op: 'set_slide_size', preset: '4:3' },
      { op: 'set_slide_size', preset: '16:9' },
    ], { output: out });
    assert.ok(readFileSync(src).equals(before), '원본이 바뀜');
    assert.match(r.results[1], /크기 200×30pt\)$/);
    assert.match(r.results[2], /글이 길어 높이를 20pt에서 늘림/);
    assert.match(r.results[7], /슬라이드 크기 4:3 \(720×540pt\) — 기존 내용은 비율에 맞춰 줄어듦/);
    assert.match(r.results[8], /슬라이드 크기 16:9 \(960×540pt\) — 넓어진 만큼 기존 내용은 크기 그대로 가운데로/);
    assert.deepEqual(r.notes, ['참고: 원본처럼 결과 파일에도 쓰기 암호가 걸려 있습니다 (다시 고칠 때도 output으로 다른 파일에 저장)']);
    assert.equal(await pptProtection(out), 'write');
    // 쓰기 암호 문서도 살펴보기는 된다 (읽기 전용으로 열림)
    const info = await pptInspect(out, { slide: 3 });
    const sh = info.slides[0].shapes;
    assert.equal(sh.find((s) => s.text === '높이 30')?.height * 4 / 3, 30, '4:3→16:9를 거친 뒤에도 비율대로'); // 4:3로 줄며 0.75배
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(readFileSync(out));
    const slide3 = await zip.file('ppt/slides/slide3.xml').async('string');
    const runOf = (t) => slide3.match(new RegExp(`<a:r>((?:(?!</a:r>)[\\s\\S])*<a:t>${t}</a:t>)`))?.[1] ?? '';
    assert.match(runOf('대외비'), /schemeClr val="tx1"/, '채우기 없음 → 본문 글자색');
    assert.match(runOf('밝은 채우기'), /schemeClr val="dk1"/, '밝은 채우기 → 어두운 글자');
    assert.match(runOf('초록 글'), /srgbClr val="00B050"/);
    const presXml = await zip.file('ppt/presentation.xml').async('string');
    assert.equal((presXml.match(/<p:modifyVerifier /g) ?? []).length, 1);
  });

  test('큰 문서 나눠 보기 (offset·limit)·슬라이드마다 도형 수 제한', { timeout: 900_000 }, async () => {
    const { default: PptxGenJS } = await import('pptxgenjs');
    const p = new PptxGenJS();
    for (let i = 1; i <= 35; i++) {
      const s = p.addSlide();
      s.addText(`슬라이드 ${i}`, { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 28 });
      if (i === 2) for (let k = 0; k < 64; k++) s.addText(`칸 ${k + 1}`, { x: 0.2 + (k % 8) * 1.2, y: 1.2 + Math.floor(k / 8) * 0.5, w: 1.1, h: 0.4, fontSize: 10 });
    }
    const big = tmp.path('big.pptx');
    writeFileSync(big, await p.write({ outputType: 'nodebuffer' }));
    const a = await pptInspect(big, { limit: 30 });
    assert.equal(a.slideCount, 35);
    assert.deepEqual(a.slides.map((s) => s.index), Array.from({ length: 30 }, (_, i) => i + 1));
    assert.equal(a.slides[1].shapeCount, 65);
    assert.equal(a.slides[1].shapes.length, 60);
    const ta = formatPptInspect(a);
    assert.match(ta, /…외 5개 \(slide=2로 이 슬라이드만 보면/);
    assert.match(ta, /\(슬라이드 1~30만 보였습니다 \(전체 35장\) — 이어서 보려면 offset=30\)$/);
    const b = await pptInspect(big, { offset: 30, limit: 30 });
    assert.deepEqual(b.slides.map((s) => s.index), [31, 32, 33, 34, 35]);
    assert.match(formatPptInspect(b), /\(슬라이드 31~35만 보였습니다 \(전체 35장\)\)$/);
    const c = await pptInspect(big, { slide: 2 });
    assert.equal(c.slides[0].shapes.length, 65);
    assert.doesNotMatch(formatPptInspect(c), /…외|보였습니다/);
  });

  test('테마(.thmx·.potx)·레이아웃 이름·슬라이드 크기', { timeout: 900_000 }, async (t) => {
    const theme = findThemeFile();
    if (!theme) { t.skip('Office 테마 파일(.thmx)이 없음'); return; }
    const out = tmp.path('theme-out.pptx');
    const r = await pptEdit(deck(), [
      { op: 'apply_theme', theme_path: theme },
      { op: 'add_slide', layout: 'comparison', title: '비교', body: '왼쪽 머리' },
      { op: 'set_slide_size', preset: '4:3' },
    ], { output: out });
    assert.match(r.results[0], /^1\. 테마 적용: .+\.thmx/);
    assert.doesNotMatch(r.results[1], /새로 만듦/, '테마에 있는 비교 레이아웃을 써야 함');
    assert.match(r.results[2], /슬라이드 크기 4:3 \(720×540pt\)/);
    const info = await pptInspect(out);
    assert.equal(info.slideWidth, 720);
    assert.equal(info.slideHeight, 540);
    for (const key of LAYOUT_KEYS) assert.ok(info.layouts.some((l) => l.key === key), `레이아웃 키 ${key}`);
    assert.equal(info.slides[2].key, 'comparison');
    // 내용도 4:3에 맞춰 줄어든다
    assert.ok(info.slides[0].shapes[0].width < 864, JSON.stringify(info.slides[0].shapes[0]));
    const img = await pageImage(out, 3, { dpi: 60 });
    const { width, height } = pngSize(img.png);
    assert.ok(Math.abs(width / height - 4 / 3) < 0.02, `${width}×${height}`);
    if (process.env.OFFICEKIT_PPT_RENDER_DIR) writeFileSync(join(process.env.OFFICEKIT_PPT_RENDER_DIR, 'theme-3.png'), img.png);

    // .potx: 같은 꾸러미를 서식 파일 형식으로 → 다른 문서에 적용, 이름으로 레이아웃 고르기
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(readFileSync(out));
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace('presentationml.presentation.main+xml', 'presentationml.template.main+xml'));
    const potx = tmp.path('theme.potx');
    writeFileSync(potx, await zip.generateAsync({ type: 'nodebuffer' }));
    const other = tmp.path('other.pptx');
    copyFileSync(deck(), other);
    const named = info.layouts.find((l) => l.key === 'title_only').name;
    const r2 = await pptEdit(other, [{ op: 'apply_theme', theme_path: potx }, { op: 'add_slide', layout: named, title: '이름으로' }]);
    assert.match(r2.results[0], /테마 적용: theme\.potx \(디자인 '/);
    assert.match(r2.results[1], new RegExp(`레이아웃: ${named.replace(/[()]/g, '\\$&')}, 제목: 이름으로`));
    const info2 = await pptInspect(other, { slide: 3 });
    assert.deepEqual(info2.designs, info.designs);
    assert.equal(info2.slides.length, 1);
    assert.equal(info2.slides[0].key, 'title_only');
  });
});
