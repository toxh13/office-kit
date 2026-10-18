// editor-server.mjs — 편집기 정적 파일 서버. 127.0.0.x 전용, GET/HEAD만, 인터넷 없이 동작.
//   편집 창(host)과 편집기는 다른 사이트로 띄운다
//   → 편집기 안 스크립트가 host의 preload API에 손대지 못하고, 렌더러 프로세스도 따로 쓴다 (사이트 격리).
//   host 서버:   127.0.0.1  /host/* → src/app/editor  /sdk/* → vendor/rhwp-editor-sdk  /office-host/* → src/app/office-editor
//   studio 서버: 127.0.0.2  /studio/* → vendor/rhwp-studio      (한글 편집기)
//   office 서버: 127.0.0.3  /office/<판>/* → vendor/office-editor (문서·표·슬라이드 편집기) + 설치 글꼴 경로
//   127.0.0.2·127.0.0.3을 못 쓰는 PC면 127.0.0.1의 다른 포트로 (출처만 나뉜다)
import { createServer } from 'node:http';
import { createServer as createNetServer } from 'node:net';
import { createHash, randomBytes } from 'node:crypto';
import { createReadStream, readdirSync, readFileSync, realpathSync, statSync } from 'node:fs';
import { Transform } from 'node:stream';
import { basename, extname, isAbsolute, join, relative, resolve, sep } from 'node:path';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.cjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
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

/**
 * Chromium이 막는 포트 (net/base/port_util.cc kRestrictedPorts). 서버가 이 포트를 받으면 편집 창이 ERR_UNSAFE_PORT로 못 연다.
 * 이 PC처럼 임시 포트가 낮은 번호(2161, 3982, 4017 …)에서 나오면 실제로 걸린다
 */
export const UNSAFE_PORTS = new Set([
  1, 7, 9, 11, 13, 15, 17, 19, 20, 21, 22, 23, 25, 37, 42, 43, 53, 69, 77, 79, 87, 95, 101, 102, 103, 104, 109, 110, 111, 113, 115,
  117, 119, 123, 135, 137, 139, 143, 161, 179, 389, 427, 465, 512, 513, 514, 515, 526, 530, 531, 532, 540, 548, 554, 556, 563, 587,
  601, 636, 989, 990, 993, 995, 1719, 1720, 1723, 2049, 3659, 4045, 4190, 5060, 5061, 6000, 6566, 6665, 6666, 6667, 6668, 6669,
  6679, 6697, 10080,
]);
export const isUnsafePort = (port) => UNSAFE_PORTS.has(port);

/** 앱 폴더 기준 편집기 폴더 (바꿔 끼울 수 있다: 시험·임시 묶음) */
export function editorDirs(root, overrides = {}) {
  const base = resolve(root ?? '.');
  return {
    studio: resolve(overrides.studio ?? join(base, 'vendor', 'rhwp-studio')),
    sdk: resolve(overrides.sdk ?? join(base, 'vendor', 'rhwp-editor-sdk')),
    host: resolve(overrides.host ?? join(base, 'src', 'app', 'editor')),
    office: resolve(overrides.office ?? join(base, 'vendor', 'office-editor')),
    officeHost: resolve(overrides.officeHost ?? join(base, 'src', 'app', 'office-editor')),
  };
}

const isFile = (p) => { try { return statSync(p).isFile(); } catch { return false; } };

/**
 * 한글 편집기 묶음이 있는지 → { available, reason }. reason은 한국어 (없으면 null)
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
 * 문서·표·슬라이드 편집기(docx·xlsx·pptx) 묶음이 있는지 → { available, reason }
 * overrides = { office, officeHost } 폴더 바꾸기
 */
export function officeEditorAvailable(root, overrides = {}, { dev = false } = {}) {
  const d = editorDirs(root, overrides);
  const reinstall = `설치가 손상됐을 수 있으니 Office Kit를 다시 설치하세요.${dev ? ' (개발 중이면 편집기 묶음을 먼저 만드세요: npm run office-editor:build)' : ''}`;
  if (!isFile(join(d.office, 'index.html'))) {
    return { available: false, reason: `문서 편집기 파일이 없습니다 (vendor/office-editor). ${reinstall}` };
  }
  for (const f of ['x2t/x2t.wasm', 'x2t/x2t.js']) {
    if (!isFile(join(d.office, ...f.split('/')))) return { available: false, reason: `문서 편집기의 변환 엔진이 없습니다 (vendor/office-editor/${f}). ${reinstall}` };
  }
  for (const f of ['editor/LICENSE.txt', 'VERSION.json']) {
    if (!isFile(join(d.office, ...f.split('/')))) return { available: false, reason: `문서 편집기 묶음이 온전하지 않습니다 (vendor/office-editor/${f}). ${reinstall}` };
  }
  for (const f of ['host.html', 'host.js', 'host.css', 'office-preload.cjs']) {
    if (!isFile(join(d.officeHost, f))) return { available: false, reason: `편집 창 파일이 없습니다 (src/app/office-editor/${f}). ${reinstall}` };
  }
  return { available: true, reason: null };
}

/**
 * office 묶음의 판 이름 (주소 /office/<판>/에 들어간다 — 묶음이 바뀌면 주소가 바뀌어 캐시가 섞이지 않는다).
 * VERSION.json 내용의 SHA-256 앞 16자리. 읽을 수 없으면 'dev'
 */
export function officeEditorVersion(dir) {
  try {
    const buf = readFileSync(join(dir, 'VERSION.json'));
    return createHash('sha256').update(buf).digest('hex').slice(0, 16);
  } catch { return 'dev'; }
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
 * 폴더 + 경로 조각(splitUrlPath가 거른 것) → 실제 파일. 폴더 밖(심볼릭 링크·정션 포함)이거나 파일이 아니면 null
 * 막는 이름(서비스 워커·*.cjs·점 이름·장치 이름)은 실제 긴 이름에도 다시 확인한다 (8.3 짧은 이름 SERVIC~1.JS 등으로 돌아가지 못하게)
 */
function resolveInDir(root, rest) {
  if (!rest.length) return null;
  rest = [...rest];
  if (rest.at(-1) === '') rest[rest.length - 1] = 'index.html';
  if (BLOCKED_NAME.test(rest.at(-1))) return null;
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

/**
 * 경로 조각 → 실제 파일. 폴더 밖(심볼릭 링크·정션 포함)이거나 파일이 아니면 null
 * routes: { 접두어: 폴더 }
 */
export function resolveRequestFile(rawUrl, routes) {
  const parts = splitUrlPath(rawUrl);
  if (!parts || parts.length < 2) return null;
  const [prefix, ...rest] = parts;
  if (!Object.hasOwn(routes, prefix)) return null;
  return resolveInDir(routes[prefix], rest);
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

/** Office host 페이지 CSP: 자기 출처 스크립트·스타일만, 틀은 office 묶음 주소(officeBase = http://127.0.0.3:C/office/<판>/)만 */
export function officeHostCsp(officeBase) {
  return [
    "default-src 'none'", "script-src 'self'", "style-src 'self'", "img-src 'self' data:", "font-src 'self'",
    "connect-src 'self'", `frame-src ${officeBase}`, `child-src ${officeBase}`, "worker-src 'none'",
    "object-src 'none'", "base-uri 'none'", "form-action 'none'", "frame-ancestors 'none'",
  ].join('; ');
}

/**
 * office 묶음 HTML CSP (설계 7장): 묶음 주소(pkg)의 파일만, 바깥 연결 없음, sandbox → opaque origin.
 * 'unsafe-eval'은 편집기가 꼭 쓴다 (템플릿·new Function) → 네트워크 없는 opaque origin 틀 안에 가둔다
 */
export function officeCsp(pkg) {
  return [
    "default-src 'none'", `script-src ${pkg} 'unsafe-eval' 'wasm-unsafe-eval'`, `style-src ${pkg} 'unsafe-inline'`,
    `img-src ${pkg} data: blob:`, `font-src ${pkg} data:`, `media-src ${pkg} blob:`, `connect-src ${pkg} blob:`,
    'worker-src blob:', `frame-src ${pkg}`, `child-src ${pkg}`, "object-src 'none'", "manifest-src 'none'",
    "base-uri 'none'", "form-action 'none'", 'sandbox allow-scripts',
  ].join('; ');
}

// ONLYOFFICE 글꼴 파일 모양: 앞 32바이트를 이 열쇠로 XOR (allfontsgen이 쓰는 값)
export const FONT_XOR_KEY = Buffer.from('a066d620149647fa9569b850b0414948', 'hex');
/** 글꼴 바이트 앞 32바이트 XOR (사본을 돌려준다) — 시험·확인용 */
export function xorFontHead(buf) {
  const out = Buffer.from(buf);
  for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= FONT_XOR_KEY[i % 16];
  return out;
}
function xorHeadStream() {
  let seen = 0;
  return new Transform({
    transform(chunk, _enc, cb) {
      if (seen >= 32) { seen += chunk.length; cb(null, chunk); return; }
      const c = Buffer.from(chunk);
      for (let i = 0; i < c.length && seen + i < 32; i++) c[i] ^= FONT_XOR_KEY[(seen + i) % 16];
      seen += c.length;
      cb(null, c);
    },
  });
}
// 글꼴로 내줄 수 있는 실제 파일 (글꼴 색인이 망가져도 문서 같은 다른 파일은 내주지 않게)
const FONT_FILE = /\.(ttf|ttc|otf|otc)$/i;
const FONT_MAX = 256 * 1024 * 1024;
const THUMB_NAME = /^fonts_thumbnail[\w@.-]*\.png(\.bin)?$/i;
// office 묶음의 고지 파일 중 MIME 표에 없는 것 (글로 내준다)
const NOTICE_TEXT = /^(?:[\w.-]+\.md|LICENSE|NOTICE|COPYING|AUTHORS)$/i;

// office 묶음 쪽 경로 (/office/<판>/ 다음)
const ALLFONTS_REL = 'editor/sdkjs/common/AllFonts.js';
const THUMBS_REL = 'editor/sdkjs/common/Images';

// 주소 후보: 127.0.0.2·127.0.0.3 = host와 다른 사이트 → Chromium이 렌더러 프로세스를 나눈다
export const STUDIO_HOSTS = ['127.0.0.2', '127.0.0.1'];
export const OFFICE_HOSTS = ['127.0.0.3', '127.0.0.1'];

function listenOnce(server, address, port = 0) {
  return new Promise((ok, fail) => {
    server.once('error', fail);
    server.listen(port, address, () => { server.off('error', fail); ok(server.address().port); });
  });
}

/**
 * listen(0)으로 포트를 받되 Chromium 금지 포트면 다시 받는다 → 포트.
 * 금지 포트는 잠깐 붙잡아 두어(같은 번호를 다시 받지 않게) 다 받은 뒤 놓는다
 */
export async function listenSafe(server, address = '127.0.0.1', { isUnsafe = isUnsafePort, tries = 20 } = {}) {
  const held = [];
  try {
    for (let i = 0; i < tries; i++) {
      const port = await listenOnce(server, address);
      if (!isUnsafe(port)) return port;
      await new Promise((ok) => server.close(() => ok()));
      const hold = createNetServer();
      try { await listenOnce(hold, address, port); held.push(hold); } catch {}
    }
    throw Object.assign(new Error(`쓸 수 있는 포트를 받지 못했습니다 (${address})`), { code: 'NO_SAFE_PORT' });
  } finally {
    for (const h of held) { try { h.close(); } catch {} }
  }
}

const BASE_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'Cross-Origin-Resource-Policy': 'same-site',
  'Cache-Control': 'no-cache',
};

/**
 * 정적 서버 하나. find(url) → { file, stat, headers?, xor? } | null
 *   headers: 이 응답에 더할 머리글 (기본 머리글을 덮어쓴다), xor: 앞 32바이트 XOR (글꼴)
 */
function makeServer(find, { isUnsafe } = {}) {
  let port = 0; let address = '127.0.0.1';
  const server = createServer((req, res) => {
    const end = (status, text = '') => {
      res.writeHead(status, { ...BASE_HEADERS, 'Content-Type': 'text/plain; charset=utf-8', 'Content-Length': Buffer.byteLength(text) });
      res.end(req.method === 'HEAD' ? undefined : text);
    };
    try {
      // 다른 이름으로 들어온 요청(DNS 다시 묶기 등)은 거절
      if (req.headers.host !== `${address}:${port}`) return end(403, '허용되지 않은 요청입니다');
      if (req.method !== 'GET' && req.method !== 'HEAD') {
        res.setHeader('Allow', 'GET, HEAD');
        return end(405, '허용되지 않은 방식입니다');
      }
      const hit = find(req.url);
      if (!hit) return end(404, '찾을 수 없습니다');
      const { file, stat } = hit;
      const etag = `W/"${stat.size.toString(16)}-${Math.floor(stat.mtimeMs).toString(16)}"`;
      const headers = { ...BASE_HEADERS, ETag: etag, ...(hit.headers ?? {}) };
      if (req.headers['if-none-match'] === etag) { res.writeHead(304, headers); return res.end(); }
      res.writeHead(200, { ...headers, 'Content-Type': hit.type ?? mimeOf(file), 'Content-Length': stat.size });
      if (req.method === 'HEAD') return res.end();
      const stream = createReadStream(file, { flags: 'r' });
      stream.on('error', () => res.destroy());
      if (hit.xor) {
        const x = xorHeadStream();
        x.on('error', () => res.destroy());
        stream.pipe(x).pipe(res);
      } else {
        stream.pipe(res);
      }
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
        try { port = await listenSafe(server, a, isUnsafe ? { isUnsafe } : {}); address = a; return { port, address }; } catch (err) { last = err; }
      }
      throw last;
    },
  };
}

function closeServer(server) {
  return new Promise((ok) => {
    if (!server.listening) { ok(); return; }
    try { server.closeAllConnections?.(); } catch {}
    server.close(() => ok());
  });
}

/** 글꼴 색인 확인 → { allFontsJsPath, thumbnailsDir, fontFileByIndex } | null */
function normalizeFonts(f) {
  if (!f || typeof f !== 'object') return null;
  return {
    allFontsJsPath: typeof f.allFontsJsPath === 'string' && isAbsolute(f.allFontsJsPath) ? f.allFontsJsPath : null,
    thumbnailsDir: typeof f.thumbnailsDir === 'string' && isAbsolute(f.thumbnailsDir) ? f.thumbnailsDir : null,
    fontFileByIndex: typeof f.fontFileByIndex === 'function' ? f.fontFileByIndex : null,
  };
}

const fileStat = (p) => { try { const st = statSync(p); return st.isFile() ? st : null; } catch { return null; } };

/**
 * 편집기 서버 시작 → {
 *   origin, port, studioOrigin, studioPort, studioHost, hostUrl, studioUrl,            (한글)
 *   officeOrigin, officePort, officeHost, officeVersion, officeKey, officeBase, officeUrl, officeHostUrl, officeError,  (문서·표·슬라이드)
 *   dirs, setOfficeFonts(fonts), close()
 * }
 * origin(port) = host 서버 (127.0.0.1, /host, /sdk, /office-host)
 * studioOrigin = studio 서버 (/studio), officeOrigin = office 서버 (/office/<판>-<실행마다 새 열쇠>/ + 글꼴 경로)
 *   officeVersion = 묶음 판 (VERSION.json 해시 — 글꼴 캐시 이름에 씀), officeKey = 주소에 쓰는 '<판>-<열쇠>'.
 *   열쇠가 없으면 같은 판을 설치한 모든 PC에서 주소가 같아, 포트를 훑는 다른 프로그램·웹 페이지가 글꼴 목록·글꼴 파일을 읽을 수 있다
 * studioHosts·officeHosts: 주소 후보 (앞에서부터 시도)
 * dirs: { studio, sdk, host, office, officeHost } 폴더 바꾸기
 * officeFonts: 글꼴 색인 { allFontsJsPath, thumbnailsDir, fontFileByIndex(i) → 절대 경로 } (나중에 setOfficeFonts로도)
 * office 서버를 못 띄워도 한글 편집은 된다 (officeOrigin = null, officeError = 한국어 이유)
 */
export async function startEditorServer({
  root, dirs, studioHosts = STUDIO_HOSTS, officeHosts = OFFICE_HOSTS, officeFonts = null, isUnsafe,
} = {}) {
  const d = editorDirs(root, dirs);
  let origin = ''; let studioOrigin = ''; let officeBase = '';
  let fonts = normalizeFonts(officeFonts);
  const officeVersion = officeEditorVersion(d.office);
  const officeKey = `${officeVersion}-${randomBytes(16).toString('hex')}`;
  const isHtml = (file) => /\.html?$/i.test(file);

  // host 서버: 한글 host·SDK, Office host (Office host 폴더의 Node 쪽 파일(*.mjs·fonts/)은 내주지 않는다)
  const hostSrv = makeServer((url) => {
    const parts = splitUrlPath(url);
    if (!parts || parts.length < 2) return null;
    const [prefix, ...rest] = parts;
    if (prefix === 'host' || prefix === 'sdk') {
      const hit = resolveInDir(prefix === 'host' ? d.host : d.sdk, rest);
      if (hit && isHtml(hit.file)) hit.headers = { 'Content-Security-Policy': hostCsp(studioOrigin) };
      return hit;
    }
    if (prefix === 'office-host') {
      if (rest[0]?.toLowerCase() === 'fonts' || /\.mjs$/i.test(rest.at(-1) ?? '')) return null;
      const hit = resolveInDir(d.officeHost, rest);
      if (!hit || /\.mjs$/i.test(hit.file)) return null;
      if (isHtml(hit.file)) hit.headers = { 'Content-Security-Policy': officeHostCsp(officeBase) };
      return hit;
    }
    return null;
  }, { isUnsafe });

  // studio는 host 창 안에서만, 바깥 연결 없이
  const studioSrv = makeServer((url) => {
    const hit = resolveRequestFile(url, { studio: d.studio });
    if (hit && isHtml(hit.file)) hit.headers = { 'Content-Security-Policy': studioCsp(origin) };
    return hit;
  }, { isUnsafe });

  // office 묶음: sandbox 틀(opaque origin)이 읽으므로 CORS·CORP를 연다. 글꼴 경로는 글꼴 색인에서
  const officeHeaders = (file, { cache = 'public, max-age=31536000, immutable' } = {}) => (isHtml(file)
    ? { 'Content-Security-Policy': officeCsp(officeBase), 'Cross-Origin-Resource-Policy': 'cross-origin', 'Cache-Control': 'no-cache' }
    : {
      'Access-Control-Allow-Origin': '*', 'Cross-Origin-Resource-Policy': 'cross-origin',
      'Content-Security-Policy': "default-src 'none'; sandbox", 'Cache-Control': cache,
    });
  const officeSrv = makeServer((url) => {
    const parts = splitUrlPath(url);
    if (!parts || parts.length < 3 || parts[0] !== 'office' || parts[1] !== officeKey) return null;
    const rest = parts.slice(2);
    const rel = rest.join('/');
    const live = { cache: 'no-cache' }; // 글꼴 색인에서 오는 파일은 바뀔 수 있다 (미리보기는 뒤에서 만든다)
    // 글꼴 파일: /office/<판>/editor/fonts/<번호> → 설치된 글꼴을 읽기 전용으로 열고 앞 32바이트만 XOR (복사하지 않는다)
    if (rest.length === 3 && rest[0] === 'editor' && rest[1] === 'fonts') {
      if (!fonts?.fontFileByIndex) {
        // 글꼴 색인이 없으면 묶음에 든 글꼴(이미 XOR된 것) 그대로
        const hit = resolveInDir(d.office, rest);
        if (hit) hit.headers = officeHeaders(hit.file);
        return hit;
      }
      if (!/^\d{1,5}$/.test(rest[2])) return null;
      let p;
      try { p = fonts.fontFileByIndex(Number(rest[2])); } catch { return null; }
      if (typeof p !== 'string' || !isAbsolute(p) || !FONT_FILE.test(p)) return null;
      const st = fileStat(p);
      if (!st || st.size > FONT_MAX) return null;
      return { file: p, stat: st, xor: true, type: 'application/octet-stream', headers: officeHeaders(p, live) };
    }
    // 글꼴 목록: 이 PC 글꼴로 만든 AllFonts.js
    if (rel === ALLFONTS_REL && fonts?.allFontsJsPath) {
      const st = fileStat(fonts.allFontsJsPath);
      if (st && /\.js$/i.test(fonts.allFontsJsPath)) return { file: fonts.allFontsJsPath, stat: st, headers: officeHeaders(fonts.allFontsJsPath, live) };
    }
    // 글꼴 미리보기 그림: 만들어졌으면 그것, 아니면 묶음 것
    if (rest.length === 5 && rest.slice(0, 4).join('/') === THUMBS_REL && THUMB_NAME.test(rest[4]) && fonts?.thumbnailsDir) {
      const p = join(fonts.thumbnailsDir, basename(rest[4]));
      const st = fileStat(p);
      if (st) return { file: p, stat: st, headers: officeHeaders(p, live) };
    }
    const hit = resolveInDir(d.office, rest);
    if (hit) {
      hit.headers = officeHeaders(hit.file);
      // 고지 파일(NOTICE-OfficeKit.md, 확장자 없는 LICENSE·NOTICE)은 글로 ('정보 · 라이선스' 보기 틀이 내려받기로 막히지 않게)
      if (NOTICE_TEXT.test(basename(hit.file))) hit.type = 'text/plain; charset=utf-8';
    }
    return hit;
  }, { isUnsafe });

  try {
    const { port } = await hostSrv.start(['127.0.0.1']);
    const { port: studioPort, address: studioHost } = await studioSrv.start(studioHosts);
    origin = `http://127.0.0.1:${port}`;
    studioOrigin = `http://${studioHost}:${studioPort}`;
    let office = { officeOrigin: null, officePort: null, officeHost: null, officeBase: null, officeUrl: null, officeError: null };
    try {
      const { port: officePort, address: officeHost } = await officeSrv.start(officeHosts);
      const officeOrigin = `http://${officeHost}:${officePort}`;
      officeBase = `${officeOrigin}/office/${officeKey}/`;
      office = { officeOrigin, officePort, officeHost, officeBase, officeUrl: `${officeBase}index.html`, officeError: null };
    } catch (err) {
      office.officeError = `문서 편집기 서버를 시작하지 못했습니다: ${err.message}`;
    }
    let closed = null;
    return {
      origin, port, studioOrigin, studioPort, studioHost, dirs: d,
      hostUrl: `${origin}/host/host.html`,
      // chrome=embed: 열기·저장 같은 파일 메뉴를 studio가 숨기고, 저장은 이 앱이 맡는다
      studioUrl: `${studioOrigin}/studio/index.html?chrome=embed`,
      ...office,
      officeVersion,
      officeKey,
      officeHostUrl: `${origin}/office-host/host.html`,
      /** 글꼴 색인 바꾸기 (null = 묶음 글꼴) */
      setOfficeFonts(f) { fonts = normalizeFonts(f); },
      get officeFonts() { return fonts; },
      close() {
        closed ??= Promise.all([closeServer(hostSrv.server), closeServer(studioSrv.server), closeServer(officeSrv.server)]).then(() => {});
        return closed;
      },
    };
  } catch (err) {
    await Promise.all([hostSrv.server, studioSrv.server, officeSrv.server].map((s) => closeServer(s)));
    throw Object.assign(new Error(`한글 편집기 서버를 시작하지 못했습니다: ${err.message}`), { code: 'EDITOR_SERVER' });
  }
}
