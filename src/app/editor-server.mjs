// editor-server.mjs — 한글 편집기(rhwp-studio) 정적 파일 서버. 127.0.0.1 전용, GET/HEAD만, 인터넷 없이 동작.
//   편집 창(host)과 편집기(studio)는 다른 사이트(127.0.0.1 / 127.0.0.2)로 띄운다
//   → studio 안 스크립트가 host의 preload API에 손대지 못하고, 렌더러 프로세스도 따로 쓴다 (사이트 격리).
//   host 서버:   127.0.0.1  /host/*  → src/app/editor      /sdk/* → vendor/rhwp-editor-sdk
//   studio 서버: 127.0.0.2  /studio/* → vendor/rhwp-studio  (127.0.0.2를 못 쓰는 PC면 127.0.0.1의 다른 포트)
import { createServer } from 'node:http';
import { createReadStream, readdirSync, realpathSync, statSync } from 'node:fs';
import { extname, isAbsolute, join, relative, resolve, sep } from 'node:path';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.cjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.webmanifest': 'application/manifest+json; charset=utf-8',
  '.wasm': 'application/wasm',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.bmp': 'image/bmp',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.ttc': 'font/collection',
  '.txt': 'text/plain; charset=utf-8',
  '.ts': 'text/plain; charset=utf-8',
};
export const mimeOf = (file) => MIME[extname(file).toLowerCase()] ?? 'application/octet-stream';

// 서비스 워커(PWA)는 앱 안에서 쓰지 않는다 — 있으면 404. *.cjs(편집 창 preload 등 Node 쪽 파일)도 내주지 않는다
const BLOCKED_NAME = /^(sw\.js|service-worker\.js|registersw\.js|workbox-[\w.-]*\.js|.*\.cjs)$/i;
// Windows 장치 이름 (CON, NUL, COM1 …) — 파일 대신 장치가 열린다
const WIN_DEVICE = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³]|conin\$|conout\$|clock\$)(\..*)?$/i;

/** 앱 폴더 기준 편집기 폴더 (바꿔 끼울 수 있다: 시험·임시 묶음) */
export function editorDirs(root, overrides = {}) {
  const base = resolve(root ?? '.');
  return {
    studio: resolve(overrides.studio ?? join(base, 'vendor', 'rhwp-studio')),
    sdk: resolve(overrides.sdk ?? join(base, 'vendor', 'rhwp-editor-sdk')),
    host: resolve(overrides.host ?? join(base, 'src', 'app', 'editor')),
  };
}

const isFile = (p) => { try { return statSync(p).isFile(); } catch { return false; } };

/**
 * 편집기 묶음이 있는지 → { available, reason }. reason은 한국어 (없으면 null)
 * root = 앱 폴더, overrides = { studio, sdk, host } 폴더 바꾸기
 */
export function editorAvailable(root, overrides = {}, { dev = false } = {}) {
  const d = editorDirs(root, overrides);
  // 개발 안내는 개발 실행에서만 (설치본 사용자에게는 다시 설치하라는 말만)
  const reinstall = `설치가 손상됐을 수 있으니 Office Kit를 다시 설치하세요.${dev ? ' (개발 중이면 편집기 묶음(vendor)을 먼저 만드세요)' : ''}`;
  if (!isFile(join(d.studio, 'index.html'))) {
    return { available: false, reason: `한글 편집기 파일이 없습니다 (vendor/rhwp-studio). ${reinstall}` };
  }
  let wasm = isFile(join(d.studio, 'rhwp_bg.wasm'));
  if (!wasm) {
    try { wasm = readdirSync(join(d.studio, 'assets')).some((f) => /^rhwp_bg.*\.wasm$/i.test(f)); } catch {}
  }
  if (!wasm) return { available: false, reason: `한글 편집기의 문서 엔진(rhwp_bg.wasm)이 없습니다. ${reinstall}` };
  for (const f of ['index.js', 'transport.js', 'document-agent-contract.js']) {
    if (!isFile(join(d.sdk, f))) return { available: false, reason: `한글 편집기 연결 파일이 없습니다 (vendor/rhwp-editor-sdk/${f}). ${reinstall}` };
  }
  for (const f of ['host.html', 'host.js', 'host.css']) {
    if (!isFile(join(d.host, f))) return { available: false, reason: `편집 창 파일이 없습니다 (src/app/editor/${f}). ${reinstall}` };
  }
  return { available: true, reason: null };
}

/**
 * 요청 주소 → 경로 조각 배열. 위험하면 null
 * - '..', '.', 점으로 시작하는 이름, 빈 조각('//'), 역슬래시·콜론·제어 문자, 끝 점/공백, 장치 이름을 거른다
 * - 조각마다 한 번만 풀어 쓴다 (%2f로 만든 '/'도 거부)
 * - 끝이 '/'면 마지막 조각이 '' (폴더 → index.html)
 */
export function splitUrlPath(raw) {
  if (typeof raw !== 'string' || !raw.startsWith('/') || raw.length > 2048) return null;
  const cut = raw.search(/[?#]/);
  const path = cut >= 0 ? raw.slice(0, cut) : raw;
  if (path.includes('\\')) return null;
  const parts = path.split('/').slice(1);
  const out = [];
  for (let i = 0; i < parts.length; i++) {
    let s;
    try { s = decodeURIComponent(parts[i]); } catch { return null; }
    if (s === '') {
      if (i === parts.length - 1 && i > 0) { out.push(''); continue; }
      return null;
    }
    if (s.startsWith('.')) return null;
    if (/[\\/:*?"<>|\u0000-\u001f\u007f]/.test(s)) return null;
    if (/[. ]$/.test(s)) return null;
    if (WIN_DEVICE.test(s)) return null;
    out.push(s);
  }
  return out;
}

const inside = (root, p) => {
  const r = relative(root, p);
  return !!r && !r.startsWith('..') && !isAbsolute(r);
};

/**
 * 경로 조각 → 실제 파일. 폴더 밖(심볼릭 링크·정션 포함)이거나 파일이 아니면 null
 * routes: { 접두어: 폴더 }
 * 막는 이름(서비스 워커·*.cjs·점 이름·장치 이름)은 실제 긴 이름에도 다시 확인한다 (8.3 짧은 이름 SERVIC~1.JS 등으로 돌아가지 못하게)
 */
export function resolveRequestFile(rawUrl, routes) {
  const parts = splitUrlPath(rawUrl);
  if (!parts || parts.length < 2) return null;
  const [prefix, ...rest] = parts;
  if (!Object.hasOwn(routes, prefix)) return null;
  if (rest.at(-1) === '') rest[rest.length - 1] = 'index.html';
  if (BLOCKED_NAME.test(rest.at(-1))) return null;
  const root = routes[prefix];
  let realRoot; let real;
  try { realRoot = realpathSync.native(root); } catch { return null; }
  try { real = realpathSync.native(join(root, ...rest)); } catch { return null; }
  if (!inside(realRoot, real)) return null;
  const segs = relative(realRoot, real).split(sep);
  if (segs.some((x) => x.startsWith('.') || WIN_DEVICE.test(x)) || BLOCKED_NAME.test(segs.at(-1))) return null;
  let st;
  try { st = statSync(real); } catch { return null; }
  if (!st.isFile()) return null;
  return { file: real, stat: st };
}

/** host 페이지 CSP: 자기 출처 스크립트·스타일만, 틀(iframe)은 studio 출처만 */
export function hostCsp(studioOrigin) {
  return [
    "default-src 'none'", "script-src 'self'", "style-src 'self'", "img-src 'self' data:", "font-src 'self'",
    "connect-src 'self'", `frame-src ${studioOrigin}`, `child-src ${studioOrigin}`, "worker-src 'none'",
    "object-src 'none'", "base-uri 'none'", "form-action 'none'", "frame-ancestors 'none'",
  ].join('; ');
}

/**
 * studio 페이지 CSP: 자기 출처 파일만 (WASM·글꼴·그림 포함), 바깥 연결(fetch·WebSocket·WebTransport 등)은 모두 거절.
 * 문서를 그리다 끼어든 스크립트가 있어도 문서 내용을 밖으로 보내지 못하게
 */
export function studioCsp(hostOrigin) {
  return [
    "default-src 'none'", "script-src 'self' 'wasm-unsafe-eval'", "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data: blob:", "font-src 'self' data: blob:", "connect-src 'self' data: blob:",
    "worker-src 'self' blob:", "frame-src 'self' blob: data:", "media-src 'self' blob: data:", "manifest-src 'none'",
    "object-src 'none'", "base-uri 'self'", "form-action 'none'", `frame-ancestors ${hostOrigin}`,
  ].join('; ');
}

// studio 주소 후보: 127.0.0.2 = host와 다른 사이트 → Chromium이 렌더러 프로세스를 나눈다
export const STUDIO_HOSTS = ['127.0.0.2', '127.0.0.1'];

function listen(server, address = '127.0.0.1') {
  return new Promise((ok, fail) => {
    server.once('error', fail);
    server.listen(0, address, () => { server.off('error', fail); ok(server.address().port); });
  });
}

function makeServer(routes, extraHeaders) {
  let port = 0; let address = '127.0.0.1';
  const server = createServer((req, res) => {
    const base = {
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer',
      'Cross-Origin-Resource-Policy': 'same-site',
      'Cache-Control': 'no-cache',
    };
    const end = (status, text = '') => {
      res.writeHead(status, { ...base, 'Content-Type': 'text/plain; charset=utf-8', 'Content-Length': Buffer.byteLength(text) });
      res.end(req.method === 'HEAD' ? undefined : text);
    };
    try {
      // 다른 이름으로 들어온 요청(DNS 다시 묶기 등)은 거절
      if (req.headers.host !== `${address}:${port}`) return end(403, '허용되지 않은 요청입니다');
      if (req.method !== 'GET' && req.method !== 'HEAD') {
        res.setHeader('Allow', 'GET, HEAD');
        return end(405, '허용되지 않은 방식입니다');
      }
      const hit = resolveRequestFile(req.url, routes);
      if (!hit) return end(404, '찾을 수 없습니다');
      const { file, stat } = hit;
      const etag = `W/"${stat.size.toString(16)}-${Math.floor(stat.mtimeMs).toString(16)}"`;
      const headers = { ...base, ETag: etag, ...extraHeaders(file, req.url) };
      if (req.headers['if-none-match'] === etag) { res.writeHead(304, headers); return res.end(); }
      res.writeHead(200, { ...headers, 'Content-Type': mimeOf(file), 'Content-Length': stat.size });
      if (req.method === 'HEAD') return res.end();
      const stream = createReadStream(file);
      stream.on('error', () => res.destroy());
      stream.pipe(res);
    } catch {
      if (!res.headersSent) end(500, '서버 오류'); else res.destroy();
    }
  });
  server.on('clientError', (_err, socket) => { try { socket.destroy(); } catch {} });
  server.keepAliveTimeout = 5000;
  return {
    server,
    /** 주소 후보를 차례로 시도 → { port, address } */
    async start(addresses = ['127.0.0.1']) {
      let last;
      for (const a of addresses) {
        try { port = await listen(server, a); address = a; return { port, address }; } catch (err) { last = err; }
      }
      throw last;
    },
  };
}

function closeServer(server) {
  return new Promise((ok) => {
    try { server.closeAllConnections?.(); } catch {}
    server.close(() => ok());
  });
}

/**
 * 편집기 서버 시작 → { origin, port, studioOrigin, studioPort, studioHost, hostUrl, studioUrl, dirs, close() }
 * origin(port) = host 서버 (127.0.0.1, /host, /sdk), studioOrigin(studioHost:studioPort) = studio 서버 (/studio)
 * studioHosts: studio 주소 후보 (앞에서부터 시도)
 * dirs: { studio, sdk, host } 폴더 바꾸기 (기본은 root 아래 vendor/·src/app/editor)
 */
export async function startEditorServer({ root, dirs, studioHosts = STUDIO_HOSTS } = {}) {
  const d = editorDirs(root, dirs);
  let origin = ''; let studioOrigin = '';
  const isHtml = (file) => /\.html?$/i.test(file);
  const hostSrv = makeServer({ host: d.host, sdk: d.sdk }, (file) => (isHtml(file) ? { 'Content-Security-Policy': hostCsp(studioOrigin) } : {}));
  // studio는 host 창 안에서만, 바깥 연결 없이
  const studioSrv = makeServer({ studio: d.studio }, (file) => (isHtml(file) ? { 'Content-Security-Policy': studioCsp(origin) } : {}));
  try {
    const { port } = await hostSrv.start(['127.0.0.1']);
    const { port: studioPort, address: studioHost } = await studioSrv.start(studioHosts);
    origin = `http://127.0.0.1:${port}`;
    studioOrigin = `http://${studioHost}:${studioPort}`;
    let closed = null;
    return {
      origin, port, studioOrigin, studioPort, studioHost, dirs: d,
      hostUrl: `${origin}/host/host.html`,
      // chrome=embed: 열기·저장 같은 파일 메뉴를 studio가 숨기고, 저장은 이 앱이 맡는다
      studioUrl: `${studioOrigin}/studio/index.html?chrome=embed`,
      close() {
        closed ??= Promise.all([closeServer(hostSrv.server), closeServer(studioSrv.server)]).then(() => {});
        return closed;
      },
    };
  } catch (err) {
    await Promise.all([hostSrv.server, studioSrv.server].map((s) => (s.listening ? closeServer(s) : null)));
    throw Object.assign(new Error(`한글 편집기 서버를 시작하지 못했습니다: ${err.message}`), { code: 'EDITOR_SERVER' });
  }
}
