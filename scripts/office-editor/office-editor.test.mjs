// office-editor.test.mjs — Office 편집기 묶음 만들기 스크립트 시험 (인터넷 없이)
//   node --test scripts/office-editor/*.test.mjs
// 캐시(tmp/office-editor-build)에 받아 둔 ui.zip 풀이가 있으면 실제 파일 목록으로 덜어 내기 수도 확인한다.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { after, describe, test } from 'node:test';
import { fileURLToPath } from 'node:url';

import { FetchError, fetchPinned } from './fetch.mjs';
import { listPatches, patchesSha256, sparseDirsFor } from './filex-source.mjs';
import { ADDED, BRIDGE, CHANGES, bridgeBanner, makeLockEditor, parseLock, verifyUi } from './lock.mjs';
import { renderNotice } from './notice.mjs';
import { EDITOR_LOCK, FILEX, UI_ZIP, X2T, cachePaths, pinsSha256 } from './pins.mjs';
import { DROP_RULES, GENERATED, NOTICE_DATED, REQUIRED, dropRuleFor, partition, rulesSha256 } from './rules.mjs';
import { THIRD_PARTY, THIRD_PARTY_BUNDLE_PATH, THIRD_PARTY_FILE, checkThirdPartyFile, excerpt, renderThirdParty } from './third-party.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const require = createRequire(import.meta.url);
const sha = (b) => createHash('sha256').update(b).digest('hex');
const scratch = mkdtempSync(join(tmpdir(), 'officekit-oe-test-'));
after(() => rmSync(scratch, { recursive: true, force: true }));

// ── 덜어 내기 규칙 ──
describe('rules', () => {
  test('언어·수식 언어는 ko·en만 남긴다', () => {
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/main/locale/de.json')?.id, 'locales');
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/main/locale/ko.json'), null);
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/main/locale/en.json'), null);
    assert.equal(dropRuleFor('editor/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/de_desc.json')?.id, 'formula-lang');
    assert.equal(dropRuleFor('editor/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/ko_desc.json'), null);
    assert.equal(dropRuleFor('editor/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/en.json'), null);
  });
  test('휴대폰·embed·forms 앱과 묶음 글꼴을 뺀다', () => {
    assert.equal(dropRuleFor('editor/web-apps/apps/presentationeditor/mobile/index.html')?.id, 'mobile-embed-forms');
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/forms/app.js')?.id, 'mobile-embed-forms');
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/embed/index.html')?.id, 'mobile-embed-forms');
    assert.equal(dropRuleFor('editor/fonts/000')?.id, 'fonts');
    assert.equal(dropRuleFor('editor/core-fonts-licenses/dejavu/LICENSE.txt')?.id, 'fonts');
    assert.equal(dropRuleFor('editor/web-apps/apps/documenteditor/main/index.html'), null);
    assert.equal(dropRuleFor('editor/sdkjs/common/AllFonts.js'), null);
    assert.equal(dropRuleFor('x2t/x2t.wasm'), null);
  });
  test('라이선스 파일은 어떤 규칙에도 맞지 않는다', () => {
    for (const p of ['LICENSE', 'NOTICE', 'editor/LICENSE.txt', 'editor/3rd-Party.txt', 'editor/license/jQuery.license.txt', 'editor/filex/CHANGES.txt']) {
      assert.equal(dropRuleFor(p), null, p);
    }
  });
  test('partition: 꼭 있어야 할 파일이 빠지면 멈춘다', () => {
    const all = [...REQUIRED, 'editor/license/a.txt', 'editor/web-apps/apps/documenteditor/main/locale/de.json'];
    const { kept, removed } = partition(all);
    assert.equal(kept.length, all.length - 1);
    assert.deepEqual(removed.get('locales'), ['editor/web-apps/apps/documenteditor/main/locale/de.json']);
    assert.throws(() => partition(all.filter((p) => p !== 'x2t/x2t.wasm')), /x2t\/x2t\.wasm/);
    assert.throws(() => partition(all.filter((p) => !p.startsWith('editor/license/'))), /editor\/license/);
    assert.throws(() => partition([...all, 'VERSION.json']), /VERSION\.json/);
  });
  test('규칙·고정 값 해시는 일정하다', () => {
    assert.match(rulesSha256(), /^[0-9a-f]{64}$/);
    assert.equal(rulesSha256(), rulesSha256());
    assert.match(pinsSha256(), /^[0-9a-f]{64}$/);
    assert.match(NOTICE_DATED, /^\d{4}-\d{2}-\d{2}$/);
  });

  const ui = cachePaths(join(ROOT, 'tmp', 'office-editor-build')).ui;
  test('실제 ui.zip 풀이에서 설계 표와 같은 만큼 뺀다', { skip: !existsSync(join(ui, 'index.html')) && '캐시에 ui.zip 풀이가 없음' }, () => {
    const files = [];
    (function walk(d) {
      for (const e of readdirSync(d, { withFileTypes: true })) {
        const p = join(d, e.name);
        if (e.isDirectory()) walk(p); else files.push(relative(ui, p).split(sep).join('/'));
      }
    })(ui);
    assert.equal(files.length, UI_ZIP.files);
    const { kept, removed } = partition(files.sort());
    const count = Object.fromEntries(DROP_RULES.map((r) => [r.id, removed.get(r.id).length]));
    assert.deepEqual(count, { locales: 264, 'formula-lang': 68, 'mobile-embed-forms': 611, fonts: 215 });
    assert.equal(kept.length, 1475);
    const bytes = kept.reduce((s, p) => s + statSync(join(ui, ...p.split('/'))).size, 0);
    assert.ok(Math.abs(bytes / 1048576 - 213.7) < 1, `남는 크기 ${(bytes / 1048576).toFixed(1)}MB`);
  });
});

// ── lock 확인 ──
function fakeBundle() {
  const ed = {
    'web-apps/apps/api/documents/api.js': Buffer.from('api'),
    [BRIDGE]: Buffer.from('original socket.io'),
    [CHANGES]: Buffer.from('filex changes\n'),
    'sdkjs/word/sdk-all.js': Buffer.from('word'),
  };
  const lock = { onlyoffice: {}, files: Object.fromEntries(Object.entries(ed).map(([p, b]) => [p, `${sha(b)} ${b.length}`])) };
  const x2tJs = Buffer.from('x2t js');
  const x2tWasm = Buffer.from('x2t wasm');
  const x2t = { 'x2t.js': sha(x2tJs), 'x2t.wasm': sha(x2tWasm) };
  const zip = new Map([
    ['editor/web-apps/apps/api/documents/api.js', ed['web-apps/apps/api/documents/api.js']],
    [`editor/${BRIDGE}`, Buffer.from(`${bridgeBanner()} (scripts/build-app.mjs).\n bridge`)],
    [`editor/${CHANGES}`, Buffer.concat([ed[CHANGES], Buffer.from('\nReplaced by this project\n')])],
    ['editor/sdkjs/word/sdk-all.js', ed['sdkjs/word/sdk-all.js']],
    ['editor/themes.json', Buffer.from(ADDED['themes.json'])],
    ['editor/plugins.json', Buffer.from(ADDED['plugins.json'])],
    ['x2t/x2t.js', x2tJs], ['x2t/x2t.wasm', x2tWasm],
    ['index.html', Buffer.from('<html>')], ['filex/app.js', Buffer.from('app')], ['filex/app.css', Buffer.from('css')],
    ['filex/x2t-worker.js', Buffer.from('w')], ['LICENSE', Buffer.from('agpl')], ['NOTICE', Buffer.from('notice')],
  ]);
  return { lock, zip, x2t, ed };
}
const check = (zip, lock, x2t) => verifyUi({
  files: new Map([...zip].map(([p, b]) => [p, { bytes: b.length, sha256: sha(b) }])),
  read: (p) => zip.get(p),
  lock,
  x2t,
});

describe('lock', () => {
  test('filex가 바꾼 것만 다르면 통과', () => {
    const { lock, zip, x2t } = fakeBundle();
    const r = check(zip, lock, x2t);
    assert.equal(r.editorFiles, 4);
    assert.equal(r.bridge.path, `editor/${BRIDGE}`);
    assert.equal(r.changes.appendedBytes, '\nReplaced by this project\n'.length);
    assert.deepEqual(r.added, ['editor/plugins.json', 'editor/themes.json']);
  });
  const broken = [
    ['편집기 파일이 바뀜', (z) => z.set('editor/sdkjs/word/sdk-all.js', Buffer.from('w0rd')), /~ editor\/sdkjs\/word\/sdk-all\.js/],
    ['편집기 파일이 빠짐', (z) => z.delete('editor/sdkjs/word/sdk-all.js'), /- editor\/sdkjs\/word\/sdk-all\.js/],
    ['lock에 없는 편집기 파일', (z) => z.set('editor/sdkjs/evil.js', Buffer.from('x')), /\+ editor\/sdkjs\/evil\.js/],
    ['다리가 아니라 원래 socket.io', (z, ed) => z.set(`editor/${BRIDGE}`, ed[BRIDGE]), /원래 socket\.io/],
    ['다리 머리글이 다름', (z) => z.set(`editor/${BRIDGE}`, Buffer.from('/* something else */')), /머리글/],
    ['CHANGES.txt 앞부분이 다름', (z) => z.set(`editor/${CHANGES}`, Buffer.from('other changes\nmore')), /CHANGES\.txt/],
    ['빈 목록이 비어 있지 않음', (z) => z.set('editor/themes.json', Buffer.from('{"themes":[1]}\n')), /themes\.json/],
    ['x2t가 다름', (z) => z.set('x2t/x2t.wasm', Buffer.from('other')), /x2t\.wasm/],
    ['모르는 맨 위 파일', (z) => z.set('evil.js', Buffer.from('x')), /\+ evil\.js/],
    ['앱 파일이 빠짐', (z) => z.delete('filex/app.js'), /- filex\/app\.js/],
  ];
  for (const [name, mutate, re] of broken) {
    test(`멈춤: ${name}`, () => {
      const { lock, zip, x2t, ed } = fakeBundle();
      mutate(zip, ed);
      assert.throws(() => check(zip, lock, x2t), re);
    });
  }
  test('parseLock: 고정 SHA-256과 모양을 확인한다', () => {
    const { lock } = fakeBundle();
    const buf = Buffer.from(JSON.stringify(lock));
    assert.equal(Object.keys(parseLock(buf, { pinned: null }).files).length, 4);
    assert.throws(() => parseLock(buf), /고정한 것과 다릅니다/);
    const bad = Buffer.from(JSON.stringify({ files: { '../x': `${'0'.repeat(64)} 1` } }));
    assert.throws(() => parseLock(bad, { pinned: null }), /경로가 이상/);
    const cached = cachePaths(join(ROOT, 'tmp', 'office-editor-build')).lock;
    if (existsSync(cached)) assert.equal(Object.keys(parseLock(readFileSync(cached)).files).length, EDITOR_LOCK.files);
  });
  test('makeLockEditor: lock 그대로의 폴더를 만든다', () => {
    const { lock, zip, ed } = fakeBundle();
    const ui = join(scratch, 'ui');
    for (const [p, b] of zip) { mkdirSync(dirname(join(ui, p)), { recursive: true }); writeFileSync(join(ui, p), b); }
    const out = join(scratch, 'lock-editor');
    const r = makeLockEditor({ uiDir: ui, lock, outDir: out, socketIo: ed[BRIDGE] });
    assert.equal(r.files, 4);
    for (const [p, b] of Object.entries(ed)) assert.deepEqual(readFileSync(join(out, p)), b, p);
    assert.equal(existsSync(join(out, 'themes.json')), false);
    assert.throws(() => makeLockEditor({ uiDir: ui, lock, outDir: out, socketIo: Buffer.from('nope') }), /socket\.io/);
  });
});

// ── 받기 ──
describe('fetch', () => {
  const body = Buffer.alloc(300_000);
  for (let i = 0; i < body.length; i++) body[i] = (i * 7) % 251;
  const pin = { bytes: body.length, sha256: sha(body) };
  const state = { mode: 'ok', hits: [] };
  const server = createServer((req, res) => {
    const range = /^bytes=(\d+)-$/.exec(req.headers.range || '');
    state.hits.push(range ? Number(range[1]) : 0);
    if (state.mode === 'cut' && state.hits.length === 1) {
      // 처음 한 번은 절반만 보내고 끊는다
      res.writeHead(200, { 'content-length': body.length });
      res.write(body.subarray(0, 120_000), () => setTimeout(() => res.destroy(), 200));
      return;
    }
    if (state.mode === 'wrong') { res.writeHead(200, { 'content-length': body.length }); res.end(Buffer.alloc(body.length)); return; }
    if (state.mode === '404') { res.writeHead(404); res.end(); return; }
    if (range && state.mode !== 'norange') {
      const from = Number(range[1]);
      res.writeHead(206, { 'content-range': `bytes ${from}-${body.length - 1}/${body.length}`, 'content-length': body.length - from });
      res.end(body.subarray(from));
      return;
    }
    res.writeHead(200, { 'content-length': body.length });
    res.end(body);
  });
  let url;
  const ready = new Promise((r) => server.listen(0, '127.0.0.1', () => { url = `http://127.0.0.1:${server.address().port}/ui.zip`; r(); }));
  after(() => server.close());
  const dest = (n) => join(scratch, 'dl', `${n}.zip`);
  const get = (n, extra = {}) => fetchPinned({ url, dest: dest(n), ...pin, label: 'ui.zip', retryMs: 1, ...extra });

  test('받고, 다음에는 캐시를 쓴다', async () => {
    await ready;
    state.mode = 'ok'; state.hits = [];
    assert.equal(await get('a'), 'download');
    assert.deepEqual(readFileSync(dest('a')), body);
    assert.equal(await get('a'), 'cache');
    assert.equal(state.hits.length, 1);
  });
  test('받다 만 파일에서 이어받는다', async () => {
    await ready;
    state.mode = 'ok'; state.hits = [];
    mkdirSync(dirname(dest('b')), { recursive: true });
    writeFileSync(`${dest('b')}.partial`, body.subarray(0, 100_000));
    assert.equal(await get('b'), 'download');
    assert.deepEqual(state.hits, [100_000]);
    assert.deepEqual(readFileSync(dest('b')), body);
  });
  test('끊기면 다시 해서 이어받는다', async () => {
    await ready;
    state.mode = 'cut'; state.hits = [];
    assert.equal(await get('c'), 'download');
    assert.equal(state.hits.length, 2);
    assert.ok(state.hits[1] > 0, '두 번째는 이어받기');
    assert.deepEqual(readFileSync(dest('c')), body);
  });
  test('서버가 이어받기를 모르면 처음부터 받는다', async () => {
    await ready;
    state.mode = 'norange'; state.hits = [];
    writeFileSync(`${dest('d')}.partial`, Buffer.from('garbage-prefix'));
    assert.equal(await get('d'), 'download');
    assert.deepEqual(readFileSync(dest('d')), body);
  });
  test('SHA-256이 다르면 파일을 지우고 멈춘다', async () => {
    await ready;
    state.mode = 'wrong'; state.hits = [];
    await assert.rejects(get('e'), (e) => e instanceof FetchError && /SHA-256/.test(e.message));
    assert.equal(existsSync(dest('e')), false);
    assert.equal(existsSync(`${dest('e')}.partial`), false);
  });
  test('캐시 파일이 고정 값과 다르면 다시 받는다', async () => {
    await ready;
    state.mode = 'ok';
    writeFileSync(dest('f'), Buffer.alloc(body.length));
    assert.equal(await get('f'), 'download');
    assert.deepEqual(readFileSync(dest('f')), body);
  });
  test('3번 다시 해도 안 되면 멈춘다 (HTTP 404)', async () => {
    await ready;
    state.mode = '404'; state.hits = [];
    await assert.rejects(get('g'), (e) => e instanceof FetchError && /HTTP 404/.test(e.message));
    assert.equal(state.hits.length, 4);
  });
  test('--offline: 캐시에 없으면 받지 않고 멈춘다', async () => {
    await ready;
    state.hits = [];
    await assert.rejects(get('h', { offline: true }), /--offline/);
    assert.equal(state.hits.length, 0);
  });
});

// ── 패치 목록 ──
describe('patches', () => {
  test('이름 순, 줄바꿈과 상관없는 해시, 머리글의 제목·날짜', () => {
    const dir = join(scratch, 'patches');
    mkdirSync(dir, { recursive: true });
    const p2 = [
      'From 0000000000000000000000000000000000000000 Mon Sep 17 00:00:00 2001',
      'From: Office Kit <noreply@example.invalid>',
      'Date: Sat, 17 Oct 2026 10:00:00 +0900',
      'Subject: [PATCH 2/2] file.save에 reason 더하기',
      '',
      'diff --git a/src/app/main.ts b/src/app/main.ts',
      '--- a/src/app/main.ts',
      '+++ b/src/app/main.ts',
      '@@ -1 +1 @@',
      '-a',
      '+b',
      '',
    ].join('\n');
    const p1 = 'diff --git a/tests/x.test.ts b/tests/x.test.ts\n--- a/tests/x.test.ts\n+++ b/tests/x.test.ts\n@@ -1 +1 @@\n-a\n+b\n';
    writeFileSync(join(dir, '0002-reason.patch'), p2.replace(/\n/g, '\r\n'));
    writeFileSync(join(dir, '0001-test.patch'), p1);
    writeFileSync(join(dir, 'README.md'), 'not a patch');
    const list = listPatches(dir);
    assert.deepEqual(list.map((p) => p.name), ['0001-test.patch', '0002-reason.patch']);
    assert.equal(list[1].subject, 'file.save에 reason 더하기');
    assert.equal(list[1].date, '2026-10-17');
    assert.equal(list[1].sha256, sha(Buffer.from(p2)));
    assert.deepEqual(list[1].touched, ['src/app/main.ts']);
    assert.deepEqual(sparseDirsFor(list), ['app', 'scripts', 'src', 'tests', 'upstream']);
    assert.match(patchesSha256(list), /^[0-9a-f]{64}$/);
    assert.equal(patchesSha256([]), null);
    assert.deepEqual(listPatches(join(scratch, 'nope')), []);
  });
});

// ── 수정 고지 ──
describe('notice', () => {
  test('수정판·날짜·원 개발사·소스 주소를 적는다', () => {
    const text = renderNotice({
      builtAt: '2026-10-18T00:00:00.000Z',
      officeKitVersion: '9.9.9',
      removed: [{ id: 'fonts', what: '글꼴', why: '시스템 글꼴', files: 2, bytes: 2048 }],
      patches: [{ name: '0001-x.patch', date: '2026-10-17', subject: '시험', sha256: 'a'.repeat(64) }],
      rebuilt: { changed: [{ path: 'filex/app.js' }] },
      fileCount: 10,
      totalBytes: 1048576,
    });
    for (const s of ['THIRD-PARTY-NATIVE.txt', 'URW++', '수정판', 'Ascensio System SIA', NOTICE_DATED, '2026-10-17', 'office-editor-source-9.9.9.zip', FILEX.commit, '0001-x.patch', 'filex/app.js', 'Modification notice', 'editor/LICENSE.txt', X2T.release]) {
      assert.ok(text.includes(s), s);
    }
  });
});

// ── 소스 꾸러미 (make-office-source.mjs) ──
describe('make-office-source', () => {
  test('MANIFEST dirtyFiles: porcelain -z의 첫 칸 공백을 잘라 먹지 않고, 이름 바꿈은 새 이름만', async () => {
    const { porcelainPaths, checkLegalDoc } = await import('../make-office-source.mjs');
    assert.deepEqual(porcelainPaths(' M LICENSE\0?? docs/a b.md\0R  new.txt\0old.txt\0 D gone.mjs\0'), ['LICENSE', 'docs/a b.md', 'new.txt', 'gone.mjs']);
    assert.deepEqual(porcelainPaths(''), []);
    // 대화상자 글이 고정 값·고지 연결(기계어 제3자 고지 포함)과 맞는다
    assert.deepEqual(checkLegalDoc(readFileSync(join(ROOT, 'docs', 'office-editor-legal-ko.md'), 'utf8')), []);
  });
});

// ── 기계어(WebAssembly) 제3자 고지 ──
describe('third-party native notices', () => {
  test('저장소의 LICENSES 파일이 목록(주소·SHA-256)과 맞고, 묶음이 만드는 파일 이름에 있다', () => {
    assert.deepEqual(checkThirdPartyFile(), []);
    assert.ok(GENERATED.includes(THIRD_PARTY_BUNDLE_PATH));
    const ids = new Set(THIRD_PARTY.map((e) => e.id));
    assert.equal(ids.size, THIRD_PARTY.length, 'id가 겹치지 않는다');
    // 이진 배포에도 고지를 함께 내라는 것들, x2t 안 URW++ 글꼴, 라이선스 파일 없는 JS
    for (const id of ['boost', 'icu', 'freetype', 'harfbuzz', 'libpng', 'libjpeg', 'libtiff', 'zlib', 'openjpeg', 'libxml2', 'xpdf', 'djvulibre', 'urw-fonts', 'emscripten', 'musl', 'libcxx', 'es6-promise', 'whatwg-fetch']) assert.ok(ids.has(id), id);
    for (const e of THIRD_PARTY) {
      assert.match(e.url, /^https:\/\/raw\.githubusercontent\.com\/[^/]+\/[^/]+\/(?:[0-9a-f]{40}|[\w.+-]*\d[\w.+-]*)\//, `${e.id}: 고정한 태그·커밋 주소`);
      assert.match(e.sha256, /^[0-9a-f]{64}$/, e.id);
    }
    const text = readFileSync(THIRD_PARTY_FILE, 'utf8');
    assert.match(text, /This software is based in part on the work of the Independent JPEG Group/);
    assert.match(text, /The FreeType Project/);
    assert.match(text, /As a special exception, permission is granted to include these font/);
  });
  test('원문에서 고지 부분만 잘라 엮는다 (BOM·CRLF 정리)', () => {
    const e = { id: 'x', from: /^LEGAL$/, to: /^REFS$/ };
    assert.equal(excerpt(e, '﻿intro\r\nLEGAL\r\nkeep me\r\nREFS\r\nlater'), 'LEGAL\nkeep me');
    assert.throws(() => excerpt({ id: 'y', from: /^NONE$/ }, 'a'), /시작 줄/);
    const texts = Object.fromEntries(THIRD_PARTY.map((t) => [t.id, `text of ${t.id}\n`]));
    // 자르기 규칙이 있는 것: 그 시작·끝 줄을 넣은 원문
    for (const t of THIRD_PARTY) {
      if (!t.from && !t.to) continue;
      const start = t.from ? { libjpeg: 'LEGAL ISSUES', libtiff: '/*', zlib: 'Copyright notice:' }[t.id] : '';
      const end = t.to ? { libjpeg: 'REFERENCES', libtiff: '#ifndef _TIFF_', xpdf: '#ifndef XREF_H', djvulibre: 'not a //C- line' }[t.id] : '';
      texts[t.id] = `${t.from ? '' : '//C- head\n'}${start}\nbody\n${end}\nrest\n`;
    }
    const out = renderThirdParty(texts);
    assert.deepEqual(checkThirdPartyFile(out), []);
    assert.throws(() => renderThirdParty({}), /원문이 없음/);
  });
});

// ── beforePack 확인 ──
describe('check-editor-bundle', () => {
  const { missingOfficeParts, OFFICE_FILES } = require('../check-editor-bundle.cjs');
  test('Office 묶음 파일·라이선스·VERSION.json을 본다', () => {
    const root = join(scratch, 'proj');
    const dir = join(root, 'vendor', 'office-editor');
    assert.ok(missingOfficeParts(root).includes('vendor/office-editor/VERSION.json'));
    const files = [];
    for (const f of [...OFFICE_FILES.filter((f) => f !== 'VERSION.json'), 'editor/license/a.txt']) {
      mkdirSync(dirname(join(dir, f)), { recursive: true });
      writeFileSync(join(dir, f), f);
      files.push({ path: f, bytes: Buffer.byteLength(f) });
    }
    writeFileSync(join(dir, 'VERSION.json'), JSON.stringify({ name: 'office-editor', files }));
    assert.deepEqual(missingOfficeParts(root), []);
    // package.json 판과 묶음의 officeKitVersion이 다르면 (판을 올리고 묶음을 다시 만들지 않음 → 고지의 소스 주소가 옛 판) 막는다
    writeFileSync(join(root, 'package.json'), JSON.stringify({ version: '9.9.9' }));
    writeFileSync(join(dir, 'VERSION.json'), JSON.stringify({ name: 'office-editor', officeKitVersion: '9.9.8', files }));
    assert.match(missingOfficeParts(root)[0], /officeKitVersion 9.9.8 ≠ package.json 9.9.9/);
    writeFileSync(join(dir, 'VERSION.json'), JSON.stringify({ name: 'office-editor', officeKitVersion: '9.9.9', files }));
    assert.deepEqual(missingOfficeParts(root), []);
    writeFileSync(join(dir, 'filex/app.js'), 'changed size!');
    assert.deepEqual(missingOfficeParts(root), ['vendor/office-editor/filex/app.js (크기가 다름)']);
    rmSync(join(dir, 'editor', 'license'), { recursive: true });
    assert.ok(missingOfficeParts(root).includes('vendor/office-editor/editor/license/'));
  });
  const real = join(ROOT, 'vendor', 'office-editor', 'VERSION.json');
  test('만들어 둔 vendor/office-editor는 통과한다', { skip: !existsSync(real) && 'vendor/office-editor 없음' }, () => {
    assert.deepEqual(missingOfficeParts(ROOT), []);
  });
});
