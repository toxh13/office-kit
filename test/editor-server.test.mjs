// editor-server.mjs 단위 시험 — 임시 폴더의 가짜 편집기 묶음으로만
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { request } from 'node:http';
import { spawnSync } from 'node:child_process';
import { linkSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  editorAvailable, editorDirs, hostCsp, mimeOf, resolveRequestFile, splitUrlPath, startEditorServer, studioCsp,
} from '../src/app/editor-server.mjs';
import { atomicWrite, bytesMatchExt, EDITABLE_EXTS } from '../src/app/editor-window.mjs';

let base; let root; let srv;
const SECRET = '비밀 내용';
// 포트 → 주소 (studio 서버는 127.0.0.2)
const hostOf = new Map();
const track = (s) => { hostOf.set(s.studioPort, s.studioHost); return s; };

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
  return { studio, sdk, host };
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
  await s.close();
  await s.close(); // 두 번 불러도 괜찮다
  await assert.rejects(raw(s.studioPort, '/studio/index.html'));
  await assert.rejects(raw(s.port, '/host/host.html'));
});

test('편집 창 도우미: 형식 확인과 원자적 쓰기', async () => {
  assert.deepEqual(EDITABLE_EXTS, ['hwp', 'hwpx', 'hwt']);
  const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0]);
  const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0]);
  assert.equal(bytesMatchExt(ole, 'hwp'), true);
  assert.equal(bytesMatchExt(ole, 'hwt'), true);
  assert.equal(bytesMatchExt(zip, 'hwpx'), true);
  assert.equal(bytesMatchExt(zip, 'hwp'), false);
  assert.equal(bytesMatchExt(ole, 'hwpx'), false);
  assert.equal(bytesMatchExt(new Uint8Array(2), 'hwp'), false);
  assert.equal(bytesMatchExt('PK', 'hwpx'), false);

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
