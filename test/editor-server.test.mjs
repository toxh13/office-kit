// editor-server.mjs 단위 시험 — 임시 폴더의 가짜 편집기 묶음으로만
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request } from 'node:http';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { linkSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  editorAvailable, editorDirs, FONT_XOR_KEY, hostCsp, isUnsafePort, listenSafe, mimeOf, officeCsp, officeEditorAvailable,
  officeEditorVersion, officeHostCsp, resolveRequestFile, splitUrlPath, startEditorServer, studioCsp, UNSAFE_PORTS, xorFontHead,
} from '../src/app/editor-server.mjs';
import { atomicWrite, bytesMatchExt, EDITABLE_EXTS, OFFICE_EXTS } from '../src/app/editor-window.mjs';

let base; let root; let srv; let fontsDir;
const SECRET = '비밀 내용';
// 포트 → 주소 (studio 서버는 127.0.0.2, office 서버는 127.0.0.3)
const hostOf = new Map();
const track = (s) => { hostOf.set(s.studioPort, s.studioHost); if (s.officePort) hostOf.set(s.officePort, s.officeHost); return s; };
const VERSION_JSON = JSON.stringify({ filex: 'v0.3.1', files: { 'index.html': 1 } });
const FONT_BYTES = Buffer.from(Array.from({ length: 100 }, (_, i) => (i * 7 + 3) & 255));

function put(path, text = 'x') {
  mkdirSync(join(path, '..'), { recursive: true });
  writeFileSync(path, text);
}

function makeTree(dir) {
  const studio = join(dir, 'vendor', 'rhwp-studio');
  put(join(studio, 'index.html'), '<!doctype html><title>studio</title>');
  put(join(studio, 'assets', 'index-abc.js'), 'export const a = 1;');
  put(join(studio, 'assets', 'mod.mjs'), 'export {};');
  put(join(studio, 'assets', 'style.css'), 'body{}');
  put(join(studio, 'assets', 'rhwp_bg-123.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(studio, 'rhwp_bg.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(studio, 'manifest.webmanifest'), '{}');
  put(join(studio, 'fonts', 'a.woff2'), 'w');
  put(join(studio, 'sw.js'), 'self.x=1');
  put(join(studio, 'registerSW.js'), 'navigator.serviceWorker.register("./sw.js")');
  put(join(studio, 'workbox-dcde9eb3.js'), 'w');
  put(join(studio, '.secret'), SECRET);
  mkdirSync(join(studio, 'empty'), { recursive: true });
  const sdk = join(dir, 'vendor', 'rhwp-editor-sdk');
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) put(join(sdk, f), 'export {};');
  const host = join(dir, 'src', 'app', 'editor');
  put(join(host, 'host.html'), '<!doctype html><title>host</title>');
  put(join(host, 'host.js'), 'console.log(1)');
  put(join(host, 'host.css'), 'body{}');
  put(join(host, 'editor-preload.cjs'), 'require("electron")');
  // Office 묶음 (vendor/office-editor 흉내)
  const office = join(dir, 'vendor', 'office-editor');
  put(join(office, 'index.html'), '<!doctype html><title>office app</title>');
  put(join(office, 'filex', 'app.js'), 'console.log("app")');
  put(join(office, 'filex', 'app.css'), 'body{}');
  put(join(office, 'editor', 'web-apps', 'apps', 'documenteditor', 'main', 'index.html'), '<!doctype html><title>de</title>');
  put(join(office, 'editor', 'sdkjs', 'common', 'AllFonts.js'), 'window["__fonts_files"] = ["000"]; // 묶음');
  put(join(office, 'editor', 'sdkjs', 'common', 'Images', 'fonts_thumbnail.png'), 'BUNDLE-PNG');
  put(join(office, 'editor', 'sdkjs', 'common', 'Images', 'fonts_thumbnail@2x.png.bin'), 'BUNDLE-BIN2');
  put(join(office, 'editor', 'sdkjs', 'common', 'Images', 'themes_thumbnail.png'), 'THEMES');
  put(join(office, 'editor', 'fonts', '000'), 'BUNDLE-FONT-000');
  put(join(office, 'editor', 'LICENSE.txt'), 'AGPL');
  put(join(office, 'editor', 'plugins.json'), '{}');
  put(join(office, 'x2t', 'x2t.js'), 'var x2t;');
  put(join(office, 'x2t', 'x2t.wasm'), Buffer.from([0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]));
  put(join(office, 'VERSION.json'), VERSION_JSON);
  put(join(office, '.secret'), SECRET);
  put(join(office, 'preload.cjs'), SECRET);
  const officeHost = join(dir, 'src', 'app', 'office-editor');
  put(join(officeHost, 'host.html'), '<!doctype html><title>office host</title>');
  put(join(officeHost, 'host.js'), 'console.log(2)');
  put(join(officeHost, 'host.css'), 'body{}');
  put(join(officeHost, 'office-preload.cjs'), SECRET);
  put(join(officeHost, 'fonts', 'index.mjs'), SECRET);
  put(join(officeHost, 'fonts', 'notes.txt'), SECRET);
  put(join(officeHost, 'helper.mjs'), SECRET);
  return { studio, sdk, host, office, officeHost };
}

/** W5 글꼴 색인 흉내: AllFonts.js·미리보기 폴더·글꼴 파일 (번호 0 = a.ttf, 1 = 문서, 2 = 없는 파일, 3 = 오류, 4 = 짧은 글꼴, 5 = 상대 경로) */
function makeFonts(dir) {
  const allFontsJsPath = join(dir, 'AllFonts.js');
  put(allFontsJsPath, 'window["__fonts_files"] = ["000"]; // 이 PC');
  const thumbnailsDir = join(dir, 'thumbs');
  put(join(thumbnailsDir, 'fonts_thumbnail.png'), 'PC-PNG');
  put(join(dir, 'sys', 'a.ttf'), FONT_BYTES);
  put(join(dir, 'sys', 'tiny.otf'), Buffer.from([1, 2, 3]));
  put(join(dir, 'sys', '문서.docx'), SECRET);
  const files = [join(dir, 'sys', 'a.ttf'), join(dir, 'sys', '문서.docx'), join(dir, 'sys', 'none.ttf'), null, join(dir, 'sys', 'tiny.otf'), 'a.ttf'];
  return {
    allFontsJsPath, thumbnailsDir,
    fontFileByIndex(i) {
      if (i === 3) throw new Error('색인 오류');
      return files[i] ?? null;
    },
  };
}

/** 주소를 고치지 않고 그대로 보내는 요청 (fetch는 '..'을 정리해 버린다) */
function raw(port, path, { method = 'GET', headers = {} } = {}) {
  return new Promise((ok, fail) => {
    const req = request({ host: hostOf.get(port) ?? '127.0.0.1', port, path, method, headers: { ...headers } }, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => ok({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
    });
    req.on('error', fail);
    req.end();
  });
}

before(async () => {
  base = mkdtempSync(join(tmpdir(), 'okeditor-'));
  root = join(base, 'app');
  const dirs = makeTree(root);
  put(join(base, 'secret.txt'), SECRET);
  put(join(root, 'vendor', 'secret.txt'), SECRET);
  // 묶음 안에서 밖을 가리키는 정션/링크 (관리자 권한 없이 되는 정션)
  mkdirSync(join(base, 'outside'), { recursive: true });
  put(join(base, 'outside', 'secret.txt'), SECRET);
  symlinkSync(join(base, 'outside'), join(dirs.studio, 'link'), 'junction');
  try { symlinkSync(join(base, 'secret.txt'), join(dirs.studio, 'filelink.txt'), 'file'); } catch {}
  symlinkSync(join(base, 'outside'), join(dirs.office, 'link'), 'junction');
  fontsDir = join(base, 'fonts');
  srv = track(await startEditorServer({ root }));
});

after(async () => {
  await srv?.close();
  rmSync(base, { recursive: true, force: true });
});

test('주소: host는 127.0.0.1, studio는 다른 사이트 127.0.0.2 (렌더러 프로세스를 나누려고)', () => {
  assert.match(srv.origin, /^http:\/\/127\.0\.0\.1:\d+$/);
  if (process.platform === 'win32') assert.equal(srv.studioHost, '127.0.0.2');
  assert.equal(srv.studioOrigin, `http://${srv.studioHost}:${srv.studioPort}`);
  assert.notEqual(srv.origin, srv.studioOrigin);
  assert.equal(srv.hostUrl, `${srv.origin}/host/host.html`);
  assert.equal(srv.studioUrl, `${srv.studioOrigin}/studio/index.html?chrome=embed`);
});

test('studio: 파일과 MIME', async () => {
  const p = srv.studioPort;
  let r = await raw(p, '/studio/index.html');
  assert.equal(r.status, 200);
  assert.equal(r.headers['content-type'], 'text/html; charset=utf-8');
  assert.equal(r.headers['x-content-type-options'], 'nosniff');
  // studio CSP: 자기 출처만, 바깥 연결(fetch·WebSocket·WebTransport) 없음, host 창 안에서만
  const csp = r.headers['content-security-policy'];
  assert.equal(csp, studioCsp(srv.origin));
  assert.match(csp, /default-src 'none'/);
  assert.match(csp, /connect-src 'self' data: blob:(;|$)/);
  assert.match(csp, /script-src 'self' 'wasm-unsafe-eval'(;|$)/);
  assert.ok(csp.endsWith(`frame-ancestors ${srv.origin}`), csp);
  assert.doesNotMatch(csp, /https?:\/\/(?!127\.0\.0\.1)/);
  r = await raw(p, '/studio/');
  assert.equal(r.status, 200);
  assert.match(r.body.toString(), /studio/);
  r = await raw(p, '/studio/assets/rhwp_bg-123.wasm');
  assert.equal(r.status, 200);
  assert.equal(r.headers['content-type'], 'application/wasm');
  assert.equal(r.body.length, 8);
  r = await raw(p, '/studio/assets/index-abc.js?v=1');
  assert.equal(r.headers['content-type'], 'text/javascript; charset=utf-8');
  assert.equal(r.headers['content-security-policy'], undefined);
  r = await raw(p, '/studio/assets/mod.mjs');
  assert.equal(r.headers['content-type'], 'text/javascript; charset=utf-8');
  r = await raw(p, '/studio/assets/style.css');
  assert.equal(r.headers['content-type'], 'text/css; charset=utf-8');
  r = await raw(p, '/studio/fonts/a.woff2');
  assert.equal(r.headers['content-type'], 'font/woff2');
  r = await raw(p, '/studio/manifest.webmanifest');
  assert.equal(r.headers['content-type'], 'application/manifest+json; charset=utf-8');
});

test('host: CSP는 자기 출처 + studio 틀만', async () => {
  const r = await raw(srv.port, '/host/host.html');
  assert.equal(r.status, 200);
  const csp = r.headers['content-security-policy'];
  assert.equal(csp, hostCsp(srv.studioOrigin));
  assert.match(csp, /default-src 'none'/);
  assert.match(csp, /script-src 'self'/);
  assert.match(csp, new RegExp(`frame-src ${srv.studioOrigin.replace(/\./g, '\\.')}(;|$)`));
  assert.match(csp, /frame-ancestors 'none'/);
  assert.doesNotMatch(csp, /unsafe/);
  const js = await raw(srv.port, '/host/host.js');
  assert.equal(js.status, 200);
  assert.equal(js.headers['content-type'], 'text/javascript; charset=utf-8');
  const sdk = await raw(srv.port, '/sdk/index.js');
  assert.equal(sdk.status, 200);
  assert.equal(sdk.headers['content-type'], 'text/javascript; charset=utf-8');
});

test('출처마다 맡은 폴더만', async () => {
  assert.equal((await raw(srv.port, '/studio/index.html')).status, 404);
  assert.equal((await raw(srv.studioPort, '/host/host.html')).status, 404);
  assert.equal((await raw(srv.studioPort, '/sdk/index.js')).status, 404);
});

test('HEAD·304·허용 안 된 방식', async () => {
  const h = await raw(srv.studioPort, '/studio/rhwp_bg.wasm', { method: 'HEAD' });
  assert.equal(h.status, 200);
  assert.equal(h.headers['content-length'], '8');
  assert.equal(h.body.length, 0);
  const again = await raw(srv.studioPort, '/studio/rhwp_bg.wasm', { headers: { 'If-None-Match': h.headers.etag } });
  assert.equal(again.status, 304);
  for (const method of ['POST', 'PUT', 'DELETE', 'OPTIONS']) {
    const r = await raw(srv.studioPort, '/studio/index.html', { method });
    assert.equal(r.status, 405, method);
    assert.equal(r.headers.allow, 'GET, HEAD');
  }
});

test('Host 머리글이 다르면 거절 (DNS 다시 묶기)', async () => {
  const r = await raw(srv.studioPort, '/studio/index.html', { headers: { Host: 'evil.example' } });
  assert.equal(r.status, 403);
  const r2 = await raw(srv.studioPort, '/studio/index.html', { headers: { Host: `localhost:${srv.studioPort}` } });
  assert.equal(r2.status, 403);
  if (srv.studioHost !== '127.0.0.1') {
    const r3 = await raw(srv.studioPort, '/studio/index.html', { headers: { Host: `127.0.0.1:${srv.studioPort}` } });
    assert.equal(r3.status, 403);
  }
});

test('studio 주소 후보: 127.0.0.2를 못 쓰면 127.0.0.1의 다른 포트로', async () => {
  const s = track(await startEditorServer({ root, studioHosts: ['203.0.113.7', '127.0.0.1'] }));
  try {
    assert.equal(s.studioHost, '127.0.0.1');
    assert.equal(s.studioOrigin, `http://127.0.0.1:${s.studioPort}`);
    assert.notEqual(s.port, s.studioPort);
    assert.equal((await raw(s.studioPort, '/studio/index.html')).status, 200);
  } finally { await s.close(); }
});

test('막는 이름은 8.3 짧은 이름으로 불러도 404 (실제 긴 이름으로 다시 확인)', async (t) => {
  if (process.platform !== 'win32') { t.skip('Windows만'); return; }
  const studio = editorDirs(root).studio;
  put(join(studio, 'service-worker.js'), 'self.x=1');
  put(join(studio, '.hiddenfolder', 'secret.txt'), SECRET);
  const short = (p) => spawnSync('cmd', ['/d', '/s', '/c', `"for %I in ("${p}") do @echo %~snxI"`], { encoding: 'utf8', windowsVerbatimArguments: true }).stdout.trim();
  const sw = short(join(studio, 'service-worker.js'));
  const hidden = short(join(studio, '.hiddenfolder'));
  if (!/~/.test(sw) || !/~/.test(hidden)) { t.skip(`이 드라이브는 8.3 이름을 만들지 않음 (${sw}, ${hidden})`); return; }
  for (const path of [`/studio/${sw}`, `/studio/${hidden}/secret.txt`]) {
    const r = await raw(srv.studioPort, path);
    assert.equal(r.status, 404, path);
    assert.ok(!r.body.toString().includes(SECRET), path);
  }
});

test('서비스 워커·숨김 파일·폴더 목록은 404', async () => {
  for (const path of ['/sw.js', '/studio/sw.js', '/studio/registerSW.js', '/studio/REGISTERSW.JS', '/studio/workbox-dcde9eb3.js',
    '/studio/.secret', '/studio/assets', '/studio/assets/', '/studio/empty/', '/studio', '/', '/index.html', '/vendor/secret.txt',
    '/host/editor-preload.cjs', '/host/', '/sdk/']) {
    const r = await raw(path.startsWith('/studio') ? srv.studioPort : srv.port, path);
    assert.equal(r.status, 404, path);
    assert.ok(!r.body.toString().includes(SECRET), path);
  }
});

test('경로 넘기 시도는 모두 404', async () => {
  const attempts = [
    '/studio/../secret.txt', '/studio/../../secret.txt', '/studio/../../../secret.txt',
    '/studio/..%2fsecret.txt', '/studio/..%2F..%2Fsecret.txt', '/studio/%2e%2e/%2e%2e/secret.txt', '/studio/%2E%2E%2F%2E%2E%2Fsecret.txt',
    '/studio/..%5c..%5csecret.txt', '/studio/assets\\..\\..\\secret.txt', '/studio/%5c..%5c..%5csecret.txt',
    '/studio/%252e%252e/secret.txt', '/studio//secret.txt', '/studio/./index.html', '/studio/assets/./../index.html',
    `/studio/${encodeURIComponent(join(base, 'secret.txt'))}`, '/studio/C:%5cWindows%5cwin.ini', '/studio/C:/Windows/win.ini',
    '/studio/%2fetc%2fpasswd', '/studio/index.html%00.js', '/studio/index.html::$DATA', '/studio/index.html.', '/studio/index.html%20',
    '/studio/NUL', '/studio/con.txt', '/studio/COM1', '/studio/link/secret.txt', '/studio/filelink.txt', '/studio/%E0%A4%A',
    '/host/../vendor/secret.txt', '/sdk/../../secret.txt',
  ];
  for (const path of attempts) {
    for (const port of [srv.port, srv.studioPort]) {
      const r = await raw(port, path);
      assert.ok([400, 404].includes(r.status), `${path} → ${r.status}`);
      assert.ok(!r.body.toString().includes(SECRET), path);
    }
  }
});

test('splitUrlPath: 조각 확인', () => {
  assert.deepEqual(splitUrlPath('/studio/assets/a.js?x=1#y'), ['studio', 'assets', 'a.js']);
  assert.deepEqual(splitUrlPath('/studio/'), ['studio', '']);
  assert.deepEqual(splitUrlPath('/studio/%ED%95%9C%EA%B8%80.png'), ['studio', '한글.png']);
  for (const bad of ['', 'studio/a', '//a', '/a/../b', '/a/%2e%2e/b', '/a/..%2fb', '/a/b%5cc', '/a\\b', '/a/.git', '/a/b:c', '/a/aux', '/a/b.', '/a/%00', '/a/%zz']) {
    assert.equal(splitUrlPath(bad), null, bad);
  }
});

test('resolveRequestFile: 폴더 밖으로 못 나간다', () => {
  const routes = { studio: editorDirs(root).studio };
  assert.ok(resolveRequestFile('/studio/index.html', routes));
  assert.equal(resolveRequestFile('/studio/link/secret.txt', routes), null);
  assert.equal(resolveRequestFile('/other/index.html', routes), null);
  assert.equal(resolveRequestFile('/studio/hasOwnProperty', routes), null);
  assert.equal(resolveRequestFile('/constructor/x', routes), null);
});

test('mimeOf: 모르는 확장자는 octet-stream', () => {
  assert.equal(mimeOf('a.wasm'), 'application/wasm');
  assert.equal(mimeOf('A.JS'), 'text/javascript; charset=utf-8');
  assert.equal(mimeOf('a.exe'), 'application/octet-stream');
});

test('editorAvailable: 묶음이 다 있으면 true, 없으면 한국어 이유', () => {
  assert.deepEqual(editorAvailable(root), { available: true, reason: null });
  const empty = join(base, 'empty-app');
  mkdirSync(empty, { recursive: true });
  let r = editorAvailable(empty);
  assert.equal(r.available, false);
  assert.match(r.reason, /한글 편집기 파일이 없습니다/);
  // 설치본 사용자에게는 개발 안내를 보이지 않는다
  assert.match(r.reason, /다시 설치하세요\.$/);
  assert.doesNotMatch(r.reason, /개발 중/);
  assert.match(editorAvailable(empty, {}, { dev: true }).reason, /개발 중이면 편집기 묶음\(vendor\)을 먼저 만드세요/);
  // studio는 있고 SDK가 없을 때
  const half = join(base, 'half-app');
  const d = makeTree(half);
  rmSync(d.sdk, { recursive: true, force: true });
  r = editorAvailable(half);
  assert.equal(r.available, false);
  assert.match(r.reason, /연결 파일이 없습니다/);
  // wasm이 없을 때
  const nowasm = join(base, 'nowasm-app');
  const d2 = makeTree(nowasm);
  rmSync(join(d2.studio, 'rhwp_bg.wasm'));
  rmSync(join(d2.studio, 'assets', 'rhwp_bg-123.wasm'));
  r = editorAvailable(nowasm);
  assert.match(r.reason, /rhwp_bg\.wasm/);
  // 폴더 바꾸기
  assert.equal(editorAvailable(empty, { studio: d.studio, sdk: editorDirs(root).sdk, host: d.host }).available, true);
});

test('close: 닫으면 더 받지 않는다', async () => {
  const s = track(await startEditorServer({ root }));
  assert.equal((await raw(s.studioPort, '/studio/index.html')).status, 200);
  assert.equal((await raw(s.officePort, `/office/${s.officeKey}/index.html`)).status, 200);
  await s.close();
  await s.close(); // 두 번 불러도 괜찮다
  await assert.rejects(raw(s.studioPort, '/studio/index.html'));
  await assert.rejects(raw(s.port, '/host/host.html'));
  await assert.rejects(raw(s.officePort, `/office/${s.officeKey}/index.html`));
});

test('편집 창 도우미: 형식 확인과 원자적 쓰기', async () => {
  assert.deepEqual(EDITABLE_EXTS, ['hwp', 'hwpx', 'hwt']);
  assert.deepEqual(OFFICE_EXTS, ['docx', 'xlsx', 'pptx']);
  const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0]);
  const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0]);
  assert.equal(bytesMatchExt(ole, 'hwp'), true);
  assert.equal(bytesMatchExt(ole, 'hwt'), true);
  assert.equal(bytesMatchExt(zip, 'hwpx'), true);
  assert.equal(bytesMatchExt(zip, 'hwp'), false);
  assert.equal(bytesMatchExt(ole, 'hwpx'), false);
  assert.equal(bytesMatchExt(new Uint8Array(2), 'hwp'), false);
  assert.equal(bytesMatchExt('PK', 'hwpx'), false);
  for (const ext of OFFICE_EXTS) {
    assert.equal(bytesMatchExt(zip, ext), true, ext);
    assert.equal(bytesMatchExt(ole, ext), false, ext);
  }

  const dir = join(base, 'write');
  mkdirSync(dir, { recursive: true });
  const target = join(dir, '문서.hwpx');
  writeFileSync(target, 'old');
  await atomicWrite(target, Buffer.from('새 내용'));
  assert.equal(readFileSync(target, 'utf8'), '새 내용');
  // 있는 파일은 내용만 덮어 복사 → 하드 링크(같은 파일)도 함께 바뀐다 (이름 바꾸기면 링크가 끊긴다)
  const linked = join(dir, '링크.hwpx');
  linkSync(target, linked);
  await atomicWrite(target, Buffer.from('세 번째'));
  assert.equal(readFileSync(linked, 'utf8'), '세 번째');
  await atomicWrite(join(dir, 'new.hwp'), Buffer.from('x'));
  assert.equal(readFileSync(join(dir, 'new.hwp'), 'utf8'), 'x');
  assert.deepEqual(readdirSync(dir).filter((f) => f.endsWith('.tmp')), []);
  await assert.rejects(atomicWrite(join(dir, 'no-such-dir', 'a.hwp'), Buffer.from('x')), { code: 'ENOENT' });
});

// ── 문서·표·슬라이드 편집기 (office) ─────────────────────────

const ob = () => `/office/${srv.officeKey}`;

test('office: 다른 사이트 127.0.0.3, 주소는 /office/<판>/ (판 = VERSION.json 해시)', () => {
  if (process.platform === 'win32') assert.equal(srv.officeHost, '127.0.0.3');
  assert.equal(srv.officeOrigin, `http://${srv.officeHost}:${srv.officePort}`);
  assert.notEqual(srv.officeOrigin, srv.origin);
  assert.notEqual(srv.officeOrigin, srv.studioOrigin);
  assert.equal(srv.officeVersion, createHash('sha256').update(VERSION_JSON).digest('hex').slice(0, 16));
  assert.equal(srv.officeVersion, officeEditorVersion(editorDirs(root).office));
  assert.equal(srv.officeBase, `${srv.officeOrigin}/office/${srv.officeKey}/`);
  assert.match(srv.officeKey, new RegExp(`^${srv.officeVersion}-[0-9a-f]{32}$`), '판 + 실행마다 새 열쇠 (주소를 미리 알 수 없게)');
  assert.equal(srv.officeUrl, `${srv.officeBase}index.html`);
  assert.equal(srv.officeHostUrl, `${srv.origin}/office-host/host.html`);
  assert.equal(srv.officeError, null);
  assert.equal(officeEditorVersion(join(base, 'no-such')), 'dev');
});

test('office: 열쇠 없는 주소 /office/<판>/는 없다 (설치본마다 같은 주소로 글꼴 목록·글꼴 파일을 읽지 못하게)', async () => {
  for (const p of ['/index.html', '/editor/sdkjs/common/AllFonts.js', '/VERSION.json']) {
    assert.equal((await raw(srv.officePort, `/office/${srv.officeVersion}${p}`)).status, 404, p);
    assert.equal((await raw(srv.officePort, `/office/${srv.officeVersion}-${'0'.repeat(32)}${p}`)).status, 404, p);
  }
  assert.equal((await raw(srv.officePort, `${ob()}/VERSION.json`)).status, 200);
});

test('office: HTML은 sandbox CSP(opaque origin), 나머지 파일은 CORS·CORP를 연다', async () => {
  let r = await raw(srv.officePort, `${ob()}/index.html`);
  assert.equal(r.status, 200);
  assert.equal(r.headers['content-type'], 'text/html; charset=utf-8');
  const csp = r.headers['content-security-policy'];
  assert.equal(csp, officeCsp(srv.officeBase));
  for (const part of ["default-src 'none'", `script-src ${srv.officeBase} 'unsafe-eval' 'wasm-unsafe-eval'`, `style-src ${srv.officeBase} 'unsafe-inline'`,
    `img-src ${srv.officeBase} data: blob:`, `font-src ${srv.officeBase} data:`, `connect-src ${srv.officeBase} blob:`, 'worker-src blob:',
    `frame-src ${srv.officeBase}`, "object-src 'none'", "base-uri 'none'", "form-action 'none'", 'sandbox allow-scripts']) {
    assert.ok(csp.split('; ').includes(part), part);
  }
  assert.doesNotMatch(csp, /allow-same-origin/);
  assert.doesNotMatch(csp, /https?:\/\/(?!127\.0\.0\.)/);
  assert.equal(r.headers['cross-origin-resource-policy'], 'cross-origin');
  assert.equal(r.headers['x-content-type-options'], 'nosniff');
  r = await raw(srv.officePort, `${ob()}/`);
  assert.equal(r.status, 200);
  assert.match(r.body.toString(), /office app/);
  r = await raw(srv.officePort, `${ob()}/editor/web-apps/apps/documenteditor/main/index.html`);
  assert.equal(r.headers['content-security-policy'], officeCsp(srv.officeBase));
  // 나머지: sandbox 틀(origin null)이 읽으므로 ACAO *, CORP cross-origin, 직접 열어도 스크립트가 돌지 않게 sandbox
  for (const [path, type] of [['/filex/app.js', 'text/javascript; charset=utf-8'], ['/x2t/x2t.wasm', 'application/wasm'],
    ['/filex/app.css', 'text/css; charset=utf-8'], ['/editor/plugins.json', 'application/json; charset=utf-8']]) {
    r = await raw(srv.officePort, `${ob()}${path}`, { headers: { Origin: 'null' } });
    assert.equal(r.status, 200, path);
    assert.equal(r.headers['content-type'], type, path);
    assert.equal(r.headers['access-control-allow-origin'], '*', path);
    assert.equal(r.headers['cross-origin-resource-policy'], 'cross-origin', path);
    assert.equal(r.headers['content-security-policy'], "default-src 'none'; sandbox", path);
    assert.match(r.headers['cache-control'], /immutable/, path);
  }
});

test('office: 판이 다르거나 맡지 않은 경로는 404, 출처마다 맡은 폴더만', async () => {
  for (const path of ['/office/0000000000000000/index.html', '/office/dev/index.html', `${ob()}`, '/office/', '/office',
    `${ob()}/.secret`, `${ob()}/preload.cjs`, `${ob()}/link/secret.txt`, `${ob()}/editor/`, `${ob()}/x2t`,
    '/studio/index.html', '/host/host.html', '/office-host/host.html', '/sdk/index.js']) {
    const r = await raw(srv.officePort, path);
    assert.equal(r.status, 404, path);
    assert.ok(!r.body.toString().includes(SECRET), path);
  }
  for (const port of [srv.port, srv.studioPort]) {
    assert.equal((await raw(port, `${ob()}/index.html`)).status, 404);
  }
  assert.equal((await raw(srv.studioPort, '/office-host/host.html')).status, 404);
});

test('office: 경로 넘기 시도·Host 머리글·허용 안 된 방식', async () => {
  const attempts = ['/../secret.txt', '/..%2f..%2fsecret.txt', '/%2e%2e/%2e%2e/secret.txt', '/editor/..%5c..%5csecret.txt',
    '/editor/fonts/..%2f..%2f..%2fsecret.txt', '/editor/fonts/000%00', '/NUL', '/index.html.', '/index.html::$DATA',
    `/${encodeURIComponent(join(base, 'secret.txt'))}`, '/editor//LICENSE.txt'];
  for (const a of attempts) {
    const r = await raw(srv.officePort, `${ob()}${a}`);
    assert.ok([400, 404].includes(r.status), `${a} → ${r.status}`);
    assert.ok(!r.body.toString().includes(SECRET), a);
  }
  assert.equal((await raw(srv.officePort, `${ob()}/index.html`, { headers: { Host: 'evil.example' } })).status, 403);
  assert.equal((await raw(srv.officePort, `${ob()}/index.html`, { headers: { Host: `localhost:${srv.officePort}` } })).status, 403);
  for (const method of ['POST', 'PUT', 'OPTIONS']) {
    assert.equal((await raw(srv.officePort, `${ob()}/index.html`, { method })).status, 405, method);
  }
});

test('office-host: CSP는 자기 출처 + office 묶음 틀만, Node 쪽 파일(*.mjs·*.cjs·fonts/)은 내주지 않는다', async () => {
  const r = await raw(srv.port, '/office-host/host.html');
  assert.equal(r.status, 200);
  const csp = r.headers['content-security-policy'];
  assert.equal(csp, officeHostCsp(srv.officeBase));
  assert.match(csp, /default-src 'none'/);
  assert.match(csp, /script-src 'self'(;|$)/);
  assert.ok(csp.includes(`frame-src ${srv.officeBase};`), csp);
  assert.match(csp, /frame-ancestors 'none'/);
  assert.doesNotMatch(csp, /unsafe/);
  assert.equal(r.headers['cross-origin-resource-policy'], 'same-site');
  assert.equal(r.headers['access-control-allow-origin'], undefined);
  assert.equal((await raw(srv.port, '/office-host/host.js')).status, 200);
  assert.equal((await raw(srv.port, '/office-host/host.css')).status, 200);
  for (const path of ['/office-host/office-preload.cjs', '/office-host/fonts/index.mjs', '/office-host/fonts/notes.txt',
    '/office-host/FONTS/notes.txt', '/office-host/helper.mjs', '/office-host/HELPER.MJS', '/office-host/', '/office-host/../editor/host.html']) {
    const x = await raw(srv.port, path);
    assert.equal(x.status, 404, path);
    assert.ok(!x.body.toString().includes(SECRET), path);
  }
  // 한글 host CSP는 그대로 (studio 틀만)
  assert.equal((await raw(srv.port, '/host/host.html')).headers['content-security-policy'], hostCsp(srv.studioOrigin));
});

test('office 글꼴: 색인이 없으면 묶음 것 그대로, 색인을 걸면 이 PC 글꼴 목록·미리보기·글꼴 파일(앞 32바이트 XOR)', async () => {
  const s = track(await startEditorServer({ root }));
  try {
    const o = (p, opts) => raw(s.officePort, `/office/${s.officeKey}${p}`, opts);
    assert.match((await o('/editor/sdkjs/common/AllFonts.js')).body.toString(), /묶음/);
    let r = await o('/editor/fonts/000');
    assert.equal(r.status, 200);
    assert.equal(r.body.toString(), 'BUNDLE-FONT-000', '묶음 글꼴은 이미 XOR된 것이라 그대로');
    assert.equal((await o('/editor/sdkjs/common/Images/fonts_thumbnail.png')).body.toString(), 'BUNDLE-PNG');

    const fonts = makeFonts(fontsDir);
    s.setOfficeFonts(fonts);
    r = await o('/editor/sdkjs/common/AllFonts.js');
    assert.equal(r.status, 200);
    assert.match(r.body.toString(), /이 PC/);
    assert.equal(r.headers['content-type'], 'text/javascript; charset=utf-8');
    assert.equal(r.headers['cache-control'], 'no-cache', '글꼴 목록은 바뀔 수 있다');
    assert.equal(r.headers['access-control-allow-origin'], '*');
    // 미리보기: 만든 것이 있으면 그것, 없으면 묶음 것
    assert.equal((await o('/editor/sdkjs/common/Images/fonts_thumbnail.png')).body.toString(), 'PC-PNG');
    assert.equal((await o('/editor/sdkjs/common/Images/fonts_thumbnail@2x.png.bin')).body.toString(), 'BUNDLE-BIN2');
    assert.equal((await o('/editor/sdkjs/common/Images/themes_thumbnail.png')).body.toString(), 'THEMES');
    // 미리보기가 나중에 생기면 바로 그것을
    put(join(fonts.thumbnailsDir, 'fonts_thumbnail@2x.png.bin'), 'PC-BIN2');
    assert.equal((await o('/editor/sdkjs/common/Images/fonts_thumbnail@2x.png.bin')).body.toString(), 'PC-BIN2');
    // 글꼴 파일: 번호 → 실제 파일, 앞 32바이트만 XOR (ONLYOFFICE가 기대하는 모양), 원본은 그대로
    for (const n of ['0', '000', '00000']) {
      r = await o(`/editor/fonts/${n}`);
      assert.equal(r.status, 200, n);
      assert.equal(r.headers['content-type'], 'application/octet-stream');
      assert.equal(r.headers['access-control-allow-origin'], '*');
      assert.equal(r.headers['cross-origin-resource-policy'], 'cross-origin');
      assert.equal(r.headers['content-length'], String(FONT_BYTES.length));
      assert.deepEqual(r.body, xorFontHead(FONT_BYTES));
      assert.deepEqual(r.body.subarray(32), FONT_BYTES.subarray(32));
      assert.equal(r.body[0], FONT_BYTES[0] ^ FONT_XOR_KEY[0]);
      assert.equal(r.body[31], FONT_BYTES[31] ^ FONT_XOR_KEY[15]);
    }
    assert.deepEqual(readFileSync(join(fontsDir, 'sys', 'a.ttf')), FONT_BYTES, '원본 글꼴 파일은 고치지 않는다');
    assert.deepEqual(xorFontHead(xorFontHead(FONT_BYTES)), FONT_BYTES);
    // 32바이트보다 짧은 글꼴
    r = await o('/editor/fonts/4');
    assert.deepEqual(r.body, Buffer.from([1 ^ FONT_XOR_KEY[0], 2 ^ FONT_XOR_KEY[1], 3 ^ FONT_XOR_KEY[2]]));
    const head = await o('/editor/fonts/0', { method: 'HEAD' });
    assert.equal(head.status, 200);
    assert.equal(head.headers['content-length'], String(FONT_BYTES.length));
    assert.equal(head.body.length, 0);
    assert.equal((await o('/editor/fonts/0', { headers: { 'If-None-Match': head.headers.etag } })).status, 304);
    // 색인이 걸리면 묶음 글꼴은 쓰지 않는다. 글꼴이 아닌 파일·없는 파일·오류·상대 경로·이상한 번호는 404
    for (const n of ['1', '2', '3', '5', '99', '000000', 'abc', '1e3', '-1', '0x1', '0.1']) {
      const x = await o(`/editor/fonts/${n}`);
      assert.equal(x.status, 404, n);
      assert.ok(!x.body.toString().includes(SECRET), n);
    }
    // 색인을 떼면 다시 묶음 것
    s.setOfficeFonts(null);
    assert.equal((await o('/editor/fonts/000')).body.toString(), 'BUNDLE-FONT-000');
    // 처음부터 색인을 넘길 수도 있다
    const s2 = track(await startEditorServer({ root, officeFonts: makeFonts(fontsDir) }));
    try {
      assert.match((await raw(s2.officePort, `/office/${s2.officeKey}/editor/sdkjs/common/AllFonts.js`)).body.toString(), /이 PC/);
    } finally { await s2.close(); }
  } finally { await s.close(); }
});

test('office 주소 후보: 127.0.0.3을 못 쓰면 127.0.0.1의 다른 포트로, 아예 못 띄워도 한글 편집은 된다', async () => {
  const s = track(await startEditorServer({ root, officeHosts: ['203.0.113.7', '127.0.0.1'] }));
  try {
    assert.equal(s.officeHost, '127.0.0.1');
    assert.equal(s.officeOrigin, `http://127.0.0.1:${s.officePort}`);
    assert.notEqual(s.port, s.officePort);
    assert.equal((await raw(s.officePort, `/office/${s.officeKey}/index.html`)).status, 200);
  } finally { await s.close(); }
  const s2 = track(await startEditorServer({ root, officeHosts: ['203.0.113.7'] }));
  try {
    assert.equal(s2.officeOrigin, null);
    assert.equal(s2.officeUrl, null);
    assert.match(s2.officeError, /문서 편집기 서버를 시작하지 못했습니다/);
    assert.equal((await raw(s2.studioPort, '/studio/index.html')).status, 200);
  } finally { await s2.close(); }
});

test('Chromium 금지 포트: 목록, listenSafe는 다시 받고, 편집기 서버 셋 다 피한다', async () => {
  for (const p of [1719, 1720, 1723, 2049, 3659, 4045, 4190, 5060, 5061, 6000, 6566, 6665, 6666, 6667, 6668, 6669, 6679, 6697, 10080, 21, 22, 25, 110, 143, 995]) {
    assert.equal(isUnsafePort(p), true, p);
  }
  for (const p of [80, 443, 2161, 3982, 4017, 8080, 9357, 49152]) assert.equal(isUnsafePort(p), false, p);
  assert.ok(UNSAFE_PORTS.size >= 80);
  // 받은 포트를 금지로 보는 판정 → 다른 포트를 받는다 (금지된 포트는 끝내 쓰지 않는다)
  const rejected = [];
  const server = createServer();
  const port = await listenSafe(server, '127.0.0.1', { isUnsafe: (p) => { if (rejected.length < 2) { rejected.push(p); return true; } return false; } });
  try {
    assert.equal(rejected.length, 2);
    assert.ok(!rejected.includes(port), `${port} ∉ ${rejected}`);
    assert.equal(server.address().port, port);
  } finally { await new Promise((r) => server.close(r)); }
  // 늘 금지면 몇 번 뒤 포기
  await assert.rejects(listenSafe(createServer(), '127.0.0.1', { isUnsafe: () => true, tries: 3 }), { code: 'NO_SAFE_PORT' });
  // 한글 host·studio·office 서버 모두
  const seen = [];
  const s = track(await startEditorServer({ root, isUnsafe: (p) => { if (seen.length < 5) { seen.push(p); return true; } return false; } }));
  try {
    assert.equal(seen.length, 5);
    for (const p of [s.port, s.studioPort, s.officePort]) assert.ok(!seen.includes(p), `${p} ∉ ${seen}`);
    assert.equal((await raw(s.studioPort, '/studio/index.html')).status, 200);
    assert.equal((await raw(s.officePort, `/office/${s.officeKey}/index.html`)).status, 200);
  } finally { await s.close(); }
});

test('officeEditorAvailable: 묶음·호스트가 다 있으면 true, 없으면 한국어 이유', () => {
  assert.deepEqual(officeEditorAvailable(root), { available: true, reason: null });
  const empty = join(base, 'empty-office');
  mkdirSync(empty, { recursive: true });
  let r = officeEditorAvailable(empty);
  assert.equal(r.available, false);
  assert.match(r.reason, /문서 편집기 파일이 없습니다 \(vendor\/office-editor\)/);
  assert.match(r.reason, /다시 설치하세요\.$/);
  assert.doesNotMatch(r.reason, /ONLYOFFICE/i);
  assert.match(officeEditorAvailable(empty, {}, { dev: true }).reason, /office-editor:build/);
  for (const [f, re] of [[['x2t', 'x2t.wasm'], /변환 엔진/], [['VERSION.json'], /VERSION\.json/], [['editor', 'LICENSE.txt'], /LICENSE\.txt/]]) {
    const dir = join(base, `office-${f.at(-1)}`);
    const d = makeTree(dir);
    rmSync(join(d.office, ...f));
    r = officeEditorAvailable(dir);
    assert.equal(r.available, false, f.join('/'));
    assert.match(r.reason, re);
  }
  const nohost = join(base, 'office-nohost');
  const d = makeTree(nohost);
  rmSync(join(d.officeHost, 'office-preload.cjs'));
  assert.match(officeEditorAvailable(nohost).reason, /편집 창 파일이 없습니다 \(src\/app\/office-editor\/office-preload\.cjs\)/);
  // 폴더 바꾸기, 한글 편집기 확인과는 따로
  assert.equal(officeEditorAvailable(empty, { office: editorDirs(root).office, officeHost: editorDirs(root).officeHost }).available, true);
  assert.equal(editorAvailable(empty).available, false);
});
