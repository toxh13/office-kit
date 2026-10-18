// editor-window.mjs — 편집 창. 문서마다 창 하나, 내장 편집기를 로컬 서버(127.0.0.x)에서 띄운다 (인터넷 없이).
//   종류(kind) 둘: 'hwp'(한글 hwp·hwpx·hwt → rhwp-studio), 'office'(docx·xlsx·pptx → vendor/office-editor)
//   Electron 모듈은 주입받는다 (main.mjs가 부르고, 시험 하니스도 같은 길로 부른다).
import { execFile } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import {
  closeSync, copyFileSync, existsSync, fsyncSync, mkdirSync, mkdtempSync, openSync, readdirSync, readFileSync, readSync, renameSync,
  rmSync, statSync, watch, writeSync,
} from 'node:fs';
import { isIP } from 'node:net';
import { tmpdir, userInfo } from 'node:os';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { inflateRawSync } from 'node:zlib';
import { editorAvailable, editorDirs, officeEditorAvailable, startEditorServer } from './editor-server.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** 한글 편집 창으로 여는 형식 (hwt는 저장할 때 HWP로 내보낸다). main.mjs가 예전부터 쓰는 이름이라 한글만 담는다 */
export const EDITABLE_EXTS = ['hwp', 'hwpx', 'hwt'];
/** 문서·표·슬라이드 편집 창으로 여는 형식 */
export const OFFICE_EXTS = ['docx', 'xlsx', 'pptx'];
/** 종류별 형식 */
export const EDITOR_KINDS = Object.freeze({ hwp: EDITABLE_EXTS, office: OFFICE_EXTS });
/** 편집 창으로 열 수 있는 모든 형식 */
export const ALL_EDITABLE_EXTS = [...EDITABLE_EXTS, ...OFFICE_EXTS];
/** 파일 경로(또는 확장자) → 'hwp' | 'office' | null */
export function editorKindOf(pathOrExt) {
  const s = String(pathOrExt ?? '');
  const ext = (s.includes('.') ? extname(s).slice(1) : s).toLowerCase();
  if (EDITABLE_EXTS.includes(ext)) return 'hwp';
  if (OFFICE_EXTS.includes(ext)) return 'office';
  return null;
}

const SAVE_AS_EXTS = ['hwp', 'hwpx'];
const PARTITIONS = { hwp: 'officekit-editor', office: 'officekit-office' }; // 'persist:'가 아님 → 저장소·서비스 워커가 디스크에 남지 않는다
/** 창 이름 ('ONLYOFFICE'는 쓰지 않는다 — 설계 8장) */
const LABELS = { hwp: '한글 편집', docx: '문서 편집', xlsx: '표 편집', pptx: '슬라이드 편집' };
const MAX_BYTES = 512 * 1024 * 1024;
// 편집기가 쓰는 권한만. 한글: 복사(클립보드 쓰기), 설치된 글꼴 읽기(옵션 창에서 사용자가 누를 때). Office: 복사만.
// 나머지(클립보드 읽기 API·카메라·위치·알림·화면 공유 …)는 거절. 메뉴의 '붙이기'는 enableDeprecatedPaste로 된다
const ALLOWED_PERMISSIONS = {
  hwp: new Set(['clipboard-sanitized-write', 'local-fonts']),
  office: new Set(['clipboard-sanitized-write']),
};
const REPLY_TIMEOUT_MS = 3000;
const SAVE_TIMEOUT_MS = 180_000;
// 아무 곳에도 이어지지 않는 주소: 맞춤법 사전 내려받기·프록시를 여기로 보내 인터넷에 나가지 못하게 한다
const DEAD_URL = 'http://127.0.0.1:9/';
const DEAD_PROXY = 'socks5://127.0.0.1:9';
/** 열린 Office 편집 창이 이보다 많으면 메모리 안내 (창마다 약 0.5GB) */
const OFFICE_WINDOWS_NOTICE = 4;
/** 자동 저장 복구 사본 폴더 (userData 아래) */
export const OFFICE_DRAFTS_DIR = 'office-drafts';
/** 편집기 설정 값 하나의 크기 한도 (JSON 바이트) */
export const STATE_VALUE_MAX = 8 * 1024;
/** 인쇄: 한 창에서 이 간격 안에 다시 부르면 거절 (편집기 틀이 PDF 앱을 연달아 띄우지 못하게) */
const PRINT_MIN_INTERVAL_MS = 3000;

const OOXML_MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};
/** 형식 설명 ('Word·Excel·PowerPoint'는 형식 설명에만 쓴다) */
const TYPE_NAMES = {
  docx: 'Word 문서', xlsx: 'Excel 통합 문서', pptx: 'PowerPoint 프레젠테이션',
  dotx: 'Word 서식 파일', xltx: 'Excel 서식 파일', potx: 'PowerPoint 서식 파일',
  pdf: 'PDF 문서', odt: 'ODF 텍스트 문서', ods: 'ODF 스프레드시트', odp: 'ODF 프레젠테이션',
  ott: 'ODF 텍스트 서식', ots: 'ODF 스프레드시트 서식', otp: 'ODF 프레젠테이션 서식',
  rtf: 'RTF 문서', txt: '텍스트 파일', csv: 'CSV 파일', html: '웹 페이지', htm: '웹 페이지', epub: 'EPUB 전자책', fb2: 'FB2 전자책',
  png: 'PNG 그림', jpg: 'JPEG 그림', jpeg: 'JPEG 그림', bmp: 'BMP 그림', gif: 'GIF 그림',
};
/** '다른 형식으로 내려받기'로 쓸 수 있는 형식 (실행 파일·매크로 문서 등은 받지 않는다) */
export const DOWNLOAD_EXTS = Object.keys(TYPE_NAMES);
const ZIP_TYPES = new Set(['docx', 'xlsx', 'pptx', 'dotx', 'xltx', 'potx', 'odt', 'ods', 'odp', 'ott', 'ots', 'otp', 'epub']);

const extOf = (p) => extname(p).slice(1).toLowerCase();
const keyOf = (p) => (process.platform === 'win32' ? resolve(p).toLowerCase() : resolve(p));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fail = (message, code) => Object.assign(new Error(message), { expected: true }, code ? { code } : {});
const labelOf = (kind, ext) => (kind === 'office' ? LABELS[ext] ?? '문서 편집' : LABELS.hwp);

// 내보낸 바이트가 형식과 맞는지 (HWP·HWT = OLE 복합 문서, HWPX·OOXML = ZIP)
const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP = [0x50, 0x4b, 0x03, 0x04];
const startsWith = (bytes, sig) => bytes.length >= sig.length && sig.every((b, i) => bytes[i] === b);
export function bytesMatchExt(bytes, ext) {
  if (!(bytes instanceof Uint8Array)) return false;
  return ext === 'hwpx' || OFFICE_EXTS.includes(ext) ? startsWith(bytes, ZIP) : startsWith(bytes, OLE);
}

// ── ZIP 목록 읽기 (OOXML 확인용: 중앙 디렉터리만 읽는다) ─────────────

/** read(위치, 길이) → Buffer 로 ZIP 중앙 디렉터리 → Map(소문자 이름 → { name, method, comp, local }) | null */
function zipDirectory(size, read) {
  if (size < 22) return null;
  const tailLen = Math.min(size, 0xffff + 22);
  const tail = read(size - tailLen, tailLen);
  let eocd = -1;
  for (let i = tail.length - 22; i >= 0; i--) if (tail.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) return null;
  let count = tail.readUInt16LE(eocd + 10);
  let cdSize = tail.readUInt32LE(eocd + 12);
  let cdOff = tail.readUInt32LE(eocd + 16);
  if (count === 0xffff || cdSize === 0xffffffff || cdOff === 0xffffffff) {
    const loc = eocd - 20; // ZIP64 위치 기록
    if (loc < 0 || tail.readUInt32LE(loc) !== 0x07064b50) return null;
    const z64 = Number(tail.readBigUInt64LE(loc + 8));
    if (z64 + 56 > size) return null;
    const z = read(z64, 56);
    if (z.readUInt32LE(0) !== 0x06064b50) return null;
    count = Number(z.readBigUInt64LE(32)); cdSize = Number(z.readBigUInt64LE(40)); cdOff = Number(z.readBigUInt64LE(48));
  }
  if (count > 200_000 || cdSize > 64 * 1024 * 1024 || cdOff + cdSize > size) return null;
  const cd = read(cdOff, cdSize);
  const out = new Map();
  for (let p = 0, n = 0; n < count; n++) {
    if (p + 46 > cd.length || cd.readUInt32LE(p) !== 0x02014b50) return null;
    const nameLen = cd.readUInt16LE(p + 28); const extraLen = cd.readUInt16LE(p + 30); const commentLen = cd.readUInt16LE(p + 32);
    const name = cd.toString('utf8', p + 46, p + 46 + nameLen);
    out.set(name.toLowerCase(), { name, method: cd.readUInt16LE(p + 10), comp: cd.readUInt32LE(p + 20), local: cd.readUInt32LE(p + 42) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/** 작은 항목 하나 풀기 (관계 파일용, 1MB까지) → 문자열 | null */
function zipText(size, read, entry) {
  try {
    if (!entry || entry.comp > 1024 * 1024 || entry.local + 30 > size) return null;
    const h = read(entry.local, 30);
    if (h.readUInt32LE(0) !== 0x04034b50) return null;
    const start = entry.local + 30 + h.readUInt16LE(26) + h.readUInt16LE(28);
    const data = read(start, entry.comp);
    if (entry.method === 0) return data.toString('utf8');
    if (entry.method === 8) return inflateRawSync(data, { maxOutputLength: 4 * 1024 * 1024 }).toString('utf8');
  } catch {}
  return null;
}

const MAIN_PART = { docx: 'word/document.xml', xlsx: 'xl/workbook.xml', pptx: 'ppt/presentation.xml' };
const MAIN_DIR = { docx: 'word/', xlsx: 'xl/', pptx: 'ppt/' };

/** OOXML 꾸러미 확인: ZIP 서명 + [Content_Types].xml + 주 부분 → null(맞음) | 한국어 이유 */
function ooxmlProblem(size, read, ext, head) {
  const what = `${TYPE_NAMES[ext]}(.${ext})`;
  if (!startsWith(head, ZIP)) return `${what} 형식이 아닙니다 (ZIP 꾸러미가 아님)`;
  let dir;
  try { dir = zipDirectory(size, read); } catch { dir = null; }
  if (!dir) return `${what} 꾸러미가 손상됐습니다 (목록을 읽을 수 없음)`;
  if (!dir.has('[content_types].xml')) return `${what} 꾸러미에 [Content_Types].xml이 없습니다`;
  if (dir.has(MAIN_PART[ext])) return null;
  // 주 부분 이름이 다른 꾸러미: _rels/.rels가 가리키는 곳
  const rels = zipText(size, read, dir.get('_rels/.rels')) ?? '';
  for (const m of rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const tag = m[0];
    if (!/Type="[^"]*\/officeDocument"/.test(tag)) continue;
    const target = /Target="([^"]+)"/.exec(tag)?.[1]?.replace(/^\//, '').toLowerCase();
    if (target && target.startsWith(MAIN_DIR[ext]) && dir.has(target)) return null;
  }
  return `${what}의 본문(${MAIN_PART[ext]})이 없습니다`;
}

/** 바이트가 그 OOXML 형식인지 → null(맞음) | 한국어 이유 */
export function checkOoxmlBytes(bytes, ext) {
  if (!(bytes instanceof Uint8Array) || !OFFICE_EXTS.includes(ext)) return '알 수 없는 형식입니다';
  const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const read = (pos, len) => buf.subarray(pos, pos + len);
  return ooxmlProblem(buf.length, read, ext, buf.subarray(0, 8));
}

/**
 * 열기 전에 거르기 (Office 문서) → 문제 없으면 null, 있으면 { code, message } (한국어)
 * - 512MB 넘음, 암호 걸린 OOXML(복합 문서 안 EncryptionInfo), 옛 형식을 이름만 바꾼 것, 꾸러미 손상
 */
export function officeOpenProblem(path, { maxBytes = MAX_BYTES } = {}) {
  const ext = extOf(path);
  let st;
  try { st = statSync(path); } catch { return { code: 'NOT_FOUND', message: `파일을 찾을 수 없습니다: ${basename(path)}` }; }
  if (st.size > maxBytes) return { code: 'TOO_BIG', message: '문서가 너무 커서 편집기로 열 수 없습니다 (512MB 넘음)' };
  if (st.size === 0) return { code: 'BAD_FORMAT', message: `‘${basename(path)}’ 파일이 비어 있어 편집할 수 없습니다.` };
  let fd;
  try { fd = openSync(path, 'r'); } catch (err) {
    return { code: err?.code === 'EBUSY' ? 'LOCKED' : 'READ_FAILED', message: `‘${basename(path)}’ 파일을 읽을 수 없습니다. 다른 프로그램이 쓰고 있으면 닫고 다시 시도하세요.` };
  }
  try {
    const read = (pos, len) => {
      const b = Buffer.alloc(len);
      let got = 0;
      while (got < len) { const n = readSync(fd, b, got, len - got, pos + got); if (!n) break; got += n; }
      return b.subarray(0, got);
    };
    const head = read(0, 8);
    if (startsWith(head, OLE)) {
      // 복합 문서: 암호 걸린 OOXML(EncryptionInfo·EncryptedPackage 흐름) 또는 옛 형식(.doc·.xls·.ppt)
      const all = read(0, st.size);
      if (all.includes(Buffer.from('EncryptionInfo', 'utf16le')) || all.includes(Buffer.from('EncryptedPackage', 'utf16le'))) {
        return { code: 'ENCRYPTED', message: '암호가 걸린 문서는 직접 편집할 수 없습니다. 프로그램에서 여세요.' };
      }
      return { code: 'OLD_FORMAT', message: `‘${basename(path)}’ 파일은 이름만 .${ext}이고 내용은 옛 형식(.doc·.xls·.ppt 등)이라 직접 편집할 수 없습니다. 프로그램에서 여세요.` };
    }
    const why = ooxmlProblem(st.size, read, ext, head);
    return why ? { code: 'BAD_FORMAT', message: `‘${basename(path)}’ 파일을 편집할 수 없습니다: ${why}.` } : null;
  } catch (err) {
    return { code: 'READ_FAILED', message: `‘${basename(path)}’ 파일을 읽지 못했습니다: ${err?.message ?? err}` };
  } finally {
    try { closeSync(fd); } catch {}
  }
}

/** 자동 저장 복구 사본 경로: userData/office-drafts/<문서 경로 sha1>.<확장자> (경로는 Windows면 소문자로 맞춘다) */
export function draftPathFor(userData, docPath) {
  return join(userData, OFFICE_DRAFTS_DIR, `${createHash('sha1').update(keyOf(docPath)).digest('hex')}.${extOf(docPath)}`);
}

function hashOf(buf) { return createHash('sha1').update(buf).digest('hex'); }
function statOf(p) {
  try { const st = statSync(p); return { mtimeMs: st.mtimeMs, size: st.size, readOnly: (st.mode & 0o200) === 0 }; } catch { return null; }
}

/** 파일 오류 → 한국어 */
function saveError(err, name) {
  const code = err?.code ?? '';
  if (['EBUSY', 'EPERM', 'EACCES'].includes(code)) {
    return fail(`‘${name}’ 파일을 다른 프로그램이 쓰고 있거나 쓸 권한이 없어 저장하지 못했습니다. 그 프로그램을 닫고 다시 저장하거나, 다른 이름으로 저장하세요.`, 'LOCKED');
  }
  if (code === 'ENOSPC') return fail('디스크 공간이 부족해 저장하지 못했습니다.', 'NO_SPACE');
  if (code === 'ENOENT') return fail('문서가 있던 폴더를 찾을 수 없어 저장하지 못했습니다. 다른 이름으로 저장하세요.', 'NOT_FOUND');
  if (code === 'EROFS') return fail('읽기 전용 드라이브라 저장할 수 없습니다. 다른 이름으로 저장하세요.', 'READ_ONLY');
  return fail(`저장하지 못했습니다: ${err?.message ?? err}`, code || 'SAVE_FAILED');
}

/**
 * 같은 폴더 임시 파일에 다 쓰고(fsync) 문서 자리에 옮긴다 (백신·미리보기가 잠깐 잡으면 몇 번 다시).
 * - 새 파일: 이름 바꾸기
 * - 있는 파일: 내용만 덮어 복사 → 파일의 보안 설정(ACL)·하드 링크·추가 스트림이 그대로 남는다
 *   (이름 바꾸기는 폴더의 기본 권한을 물려받아 '나만 읽기' 같은 설정이 풀린다). 복사가 끝내 안 되면 이름 바꾸기로
 */
export async function atomicWrite(path, bytes, { retries = 8 } = {}) {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomBytes(4).toString('hex')}.tmp`);
  const fd = openSync(tmp, 'wx');
  try {
    for (let off = 0; off < bytes.length;) off += writeSync(fd, bytes, off, bytes.length - off);
    fsyncSync(fd);
  } catch (err) {
    closeSync(fd);
    rmSync(tmp, { force: true });
    throw err;
  }
  closeSync(fd);
  const busy = (err) => ['EBUSY', 'EPERM', 'EACCES'].includes(err?.code);
  try {
    if (existsSync(path)) {
      for (let i = 0; ; i++) {
        try { copyFileSync(tmp, path); return; } catch (err) {
          if (busy(err) && i < retries) { await sleep(60 * (i + 1)); continue; }
          // 숨김 속성 등으로 덮어 복사가 안 되거나 도중에 끊겼으면 이름 바꾸기로 (그래도 안 되면 처음 오류)
          try { renameSync(tmp, path); return; } catch { throw err; }
        }
      }
    }
    for (let i = 0; ; i++) {
      try { renameSync(tmp, path); return; } catch (err) {
        if (!busy(err) || i >= retries) throw err;
        await sleep(60 * (i + 1));
      }
    }
  } finally {
    rmSync(tmp, { force: true });
  }
}

/** 사용자 이름 다듬기: 제어 문자 빼고 64자까지 (없으면 '') */
export function cleanUserName(v) {
  return typeof v === 'string' ? v.replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 64) : '';
}

/** reg export 결과(.reg, UTF-16LE)에서 "UserName" 값 */
export function parseRegUserName(buf) {
  if (!Buffer.isBuffer(buf)) return '';
  const text = buf[0] === 0xff && buf[1] === 0xfe ? buf.toString('utf16le', 2) : buf.toString('utf8');
  const m = /^"UserName"="((?:[^"\\]|\\.)*)"\s*$/m.exec(text);
  return m ? cleanUserName(m[1].replace(/\\(.)/g, '$1')) : '';
}

/**
 * 편집기에 넘길 사용자 이름 (변경 내용 추적·메모·마지막 수정한 사람에 들어간다):
 * Office에 적힌 이름(HKCU\Software\Microsoft\Office\Common\UserInfo\UserName) → 없으면 Windows 계정 이름.
 * reg export는 UTF-16으로 써서 한글 이름도 깨지지 않는다
 */
export function readOfficeUserName({ timeoutMs = 3000 } = {}) {
  const fallback = () => { try { return cleanUserName(userInfo().username); } catch { return ''; } };
  if (process.platform !== 'win32') return Promise.resolve(fallback());
  return new Promise((res) => {
    let dir;
    try { dir = mkdtempSync(join(tmpdir(), 'officekit-user-')); } catch { res(fallback()); return; }
    const file = join(dir, 'user.reg');
    execFile('reg.exe', ['export', 'HKCU\\Software\\Microsoft\\Office\\Common\\UserInfo', file, '/y'], { windowsHide: true, timeout: timeoutMs }, (err) => {
      let name = '';
      try { if (!err) name = parseRegUserName(readFileSync(file)); } catch {}
      try { rmSync(dir, { recursive: true, force: true }); } catch {}
      res(name || fallback());
    });
  });
}

/** 브라우저로 넘기면 안 되는 주소: 이 PC·사설망 (편집기 서버 주소 포함) */
export function isPrivateHost(hostname) {
  let h = String(hostname ?? '').toLowerCase().replace(/^\[|\]$/g, '');
  if (!h || h === 'localhost' || h.endsWith('.localhost') || h === '0.0.0.0') return true;
  const dotted = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(h);
  if (dotted) h = dotted[1];
  const hex = /^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(h);
  if (hex) { const a = parseInt(hex[1], 16); const b = parseInt(hex[2], 16); h = [a >> 8, a & 255, b >> 8, b & 255].join('.'); }
  if (isIP(h) === 4) {
    const [a, b] = h.split('.').map(Number);
    return a === 127 || a === 10 || a === 0 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31)
      || (a === 192 && b === 168) || (a === 100 && b >= 64 && b <= 127);
  }
  if (isIP(h) === 6) return h === '::1' || h === '::' || /^f[cd]/.test(h) || /^fe[89ab]/.test(h);
  return false;
}

/** 파일 이름으로 쓸 수 있게 (경로 구분자·금지 문자 → _) */
function safeFileName(name, fallback) {
  let s = String(name ?? '').replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_').replace(/[. ]+$/, '').trim();
  if (s.length > 180) s = s.slice(0, 180);
  return s || fallback;
}

/**
 * 편집 창 관리자 (한글·Office 두 종류).
 * @param {object} o
 * @param {import('electron').App} o.app
 * @param {typeof import('electron').BrowserWindow} o.BrowserWindow
 * @param {import('electron').IpcMain} o.ipcMain
 * @param {typeof import('electron').session} o.session
 * @param {import('electron').Dialog} o.dialog
 * @param {import('electron').Shell} o.shell
 * @param {string} o.appRoot              앱 폴더 (vendor/, src/app/editor 기준)
 * @param {string} [o.userData]           Office 자동 저장 사본(office-drafts/)·글꼴 캐시 기준 폴더
 * @param {() => ({ backup(path): any }) | { backup(path): any }} [o.getBackups]  저장 전에 원본 백업
 * @param {(path: string) => void} [o.onSaved]      저장(다른 이름 포함) 뒤
 * @param {(path: string) => void} [o.onDownloaded] 다른 형식으로 내려받기 뒤
 * @param {(path: string, dirty: boolean) => void} [o.onDirtyChange]
 * @param {{ studio?, sdk?, host?, office?, officeHost? }} [o.dirs]  편집기 폴더 바꾸기 (시험·임시 묶음)
 * @param {{ get(key): any, set(key, value): any }} [o.officeState]  Office 편집기 설정 보관 (없으면 이번 실행 동안만)
 * @param {{ ensureFonts(o): Promise<object> }} [o.fonts]  글꼴 모듈 (기본: ./office-editor/fonts/index.mjs)
 * @param {{ applyFixups(ext, bytes, o): any }} [o.fixups]    저장 뒤 고치기 모듈 (기본: ../core/ooxml-fixups.mjs). o.original = 덮어쓸 원래 파일 바이트
 * @param {string | (() => string | Promise<string>)} [o.userName]  편집기 사용자 이름 (기본: Office에 적힌 이름 → Windows 계정 이름)
 * @param {object} [o.nativeTheme]  창 배경색 고르기용
 * @param {string} [o.icon]
 * @param {boolean} [o.testHooks]   host 페이지에 시험용 고리를 연다 (?test=1)
 * @param {(...args) => void} [o.log]
 * @param {{ maxBytes?: number, fontsTimeoutMs?: number, printIntervalMs?: number }} [o.limits]  시험용 한도 (기본 512MB, 글꼴 준비 20초, 인쇄 간격 3초)
 */
export function createEditorManager({
  app, BrowserWindow, ipcMain, session, dialog, shell, appRoot, userData, getBackups, onSaved, onDownloaded, onDirtyChange,
  dirs, officeState, fonts: fontsModule, fixups: fixupsModule, nativeTheme, icon, testHooks = false, userName,
  log = (...a) => console.log('[editor]', ...a), limits = {},
} = {}) {
  const root = resolve(appRoot ?? join(here, '..', '..'));
  const dirOverrides = { ...(dirs ?? {}) };
  const maxBytes = Number.isFinite(limits.maxBytes) ? limits.maxBytes : MAX_BYTES;
  const fontsTimeoutMs = Number.isFinite(limits.fontsTimeoutMs) ? limits.fontsTimeoutMs : 20_000;
  const printIntervalMs = Number.isFinite(limits.printIntervalMs) ? limits.printIntervalMs : PRINT_MIN_INTERVAL_MS;
  const entries = new Map();    // 문서 키 → 창 정보
  const byWc = new Map();       // webContents id → 창 정보
  const pending = new Map();    // 요청 id → { resolve, timer, entry }
  const memoryState = new Map(); // officeState가 없을 때 편집기 설정
  let server = null; let serverStarting = null;
  const sessionsReady = new Map(); // 종류 → 세션 준비 Promise
  let quitting = false;
  let seq = 0;
  let fontsPromise = null; let fontsWarning = null;
  let userNamePromise = null;

  const dev = !app?.isPackaged;
  const available = (kind = 'hwp') => (kind === 'office'
    ? officeEditorAvailable(root, dirOverrides, { dev })
    : editorAvailable(root, dirOverrides, { dev }));
  const officeHostDir = () => editorDirs(root, dirOverrides).officeHost;

  // ── 서버 · 세션 ───────────────────────────────────────

  async function ensureServer() {
    if (server) return server;
    serverStarting ??= startEditorServer({ root, dirs: dirOverrides })
      .then((s) => { server = s; return s; })
      .finally(() => { serverStarting = null; });
    return serverStarting;
  }

  /** 그 종류의 창이 불러와도 되는 주소인지 (세션 요청 막기) */
  function allowedUrl(url, kind, srv = server) {
    if (/^(blob|data):/i.test(url) || /^(devtools|chrome-devtools):/i.test(url) || url === 'about:blank' || url === 'about:srcdoc') return true;
    if (!srv) return false;
    let u;
    try { u = new URL(url); } catch { return false; }
    if (u.protocol !== 'http:') return false;
    if (kind === 'office') {
      return (u.origin === srv.origin && u.pathname.startsWith('/office-host/'))
        || (!!srv.officeOrigin && u.origin === srv.officeOrigin && u.pathname.startsWith(`/office/${srv.officeKey}/`));
    }
    return u.origin === srv.origin || u.origin === srv.studioOrigin;
  }

  // 세션(요청 막기·프록시)을 다 갖춘 뒤에 창을 연다. 종류마다 세션이 따로다 (한글 창은 Office 출처에, Office 창은 studio에 못 닿는다)
  function setupSession(srv, kind) {
    if (!sessionsReady.has(kind)) {
      const p = prepareSession(srv, kind);
      sessionsReady.set(kind, p);
      p.catch(() => sessionsReady.delete(kind)); // 실패하면 다음 창에서 다시
    }
    return sessionsReady.get(kind);
  }

  async function prepareSession(srv, kind) {
    const ses = session.fromPartition(PARTITIONS[kind]);
    // 맞춤법 사전: 끄기만 해도 Chromium이 인터넷에서 사전(ko-3-0.bdic, 11MB)을 받으므로 받을 곳을 죽은 주소로
    try { ses.setSpellCheckerDictionaryDownloadURL(DEAD_URL); } catch {}
    try { ses.setSpellCheckerEnabled(false); } catch {}
    // HTTP 밖의 연결(WebRTC TURN-TCP·WebTransport 등)은 요청 막기에 안 잡히므로 죽은 프록시로 보낸다.
    // 바로 가는 것은 우리 서버 두 곳(주소:포트)뿐: Chromium이 늘 바로 보내는 이 PC 주소(127.0.0.0/8) 규칙은 <-loopback>으로 끈다
    // (끄지 않으면 편집기 틀이 TURN-TCP로 이 PC의 다른 서비스 포트에 닿는다). 서버가 다시 뜨면 포트가 바뀌므로 dispose에서 다시 만든다
    const own = kind === 'office' ? [srv.origin, srv.officeOrigin] : [srv.origin, srv.studioOrigin];
    const bypass = ['<-loopback>', ...new Set(own.filter(Boolean).map((o) => new URL(o).host))].join(';');
    await ses.setProxy({ mode: 'fixed_servers', proxyRules: DEAD_PROXY, proxyBypassRules: bypass });
    // 내 서버 밖으로 나가는 요청은 모두 막는다 (CDN 글꼴·외부 그림 등)
    ses.webRequest.onBeforeRequest({ urls: ['<all_urls>'] }, (details, cb) => {
      const ok = allowedUrl(details.url, kind, server ?? srv);
      if (!ok) log('막은 요청:', details.url.slice(0, 200));
      cb({ cancel: !ok });
    });
    const allowed = ALLOWED_PERMISSIONS[kind];
    ses.setPermissionRequestHandler((wc, permission, cb, details) => {
      const ok = allowed.has(permission) && byWc.get(wc?.id)?.kind === kind;
      if (!ok) log('거절한 권한:', permission, details?.requestingUrl ?? '');
      cb(ok);
    });
    ses.setPermissionCheckHandler((wc, permission) => allowed.has(permission) && (!wc || byWc.get(wc.id)?.kind === kind));
    try { ses.setDevicePermissionHandler(() => false); } catch {}
    try { ses.setDisplayMediaRequestHandler((_req, cb) => cb({})); } catch {}
    ses.on('will-download', (e) => e.preventDefault());
  }

  // ── 글꼴 · 저장 뒤 고치기 (Office) ──────────────────────

  async function loadFontsModule() {
    if (fontsModule) return fontsModule;
    return import(pathToFileURL(join(here, 'office-editor', 'fonts', 'index.mjs')).href);
  }

  /** 이 PC 글꼴 목록을 준비해 서버에 건다 (첫 Office 창에서 한 번, 실패하면 다음 창에서 다시) */
  function ensureOfficeFonts(srv) {
    fontsPromise ??= (async () => {
      const t0 = Date.now();
      try {
        const m = await loadFontsModule();
        const f = await m.ensureFonts({ userData: userData ?? join(tmpdir(), 'officekit-editor'), editorVersion: srv.officeVersion });
        srv.setOfficeFonts(f);
        fontsWarning = null;
        log(`글꼴 목록 준비 (${Date.now() - t0}ms)`);
        return true;
      } catch (err) {
        log('글꼴 목록을 만들지 못했습니다:', err?.message ?? err);
        fontsWarning = '이 PC의 글꼴 목록을 만들지 못해 일부 글자가 다른 글꼴이나 □로 보일 수 있습니다.';
        fontsPromise = null;
        return false;
      }
    })();
    return fontsPromise;
  }

  /**
   * 저장 뒤 고치기 → { bytes, applied, warning }. 실패하면 원래 바이트 (저장은 막지 않는다, pptx만 알림).
   * original: 편집기가 열었던 문서의 지금 바이트 — x2t가 잃는 것(테마 글꼴·한국어 줄 바꿈·구역·메모 …)을 옮겨 온다
   */
  async function runFixups(ext, bytes, original = null) {
    const warn = ext === 'pptx'
      ? '저장은 했지만 PowerPoint 호환 고치기를 하지 못했습니다. PowerPoint에서 열 때 복구 안내가 나올 수 있습니다.'
      : null;
    try {
      const m = fixupsModule ?? await import(pathToFileURL(join(here, '..', 'core', 'ooxml-fixups.mjs')).href);
      // applyFixups(ext, bytes, { original }) → { bytes, applied[] } 또는 fixOoxml(bytes, ext, { original }) → { bytes, fixes[], warnings[], error, notify }
      let r;
      if (typeof m.applyFixups === 'function') r = await m.applyFixups(ext, bytes, { original });
      else if (typeof m.fixOoxml === 'function') r = await m.fixOoxml(bytes, ext, { original });
      else throw new Error('고치기 모듈에 applyFixups가 없습니다');
      const out = r?.bytes;
      if (!(out instanceof Uint8Array)) throw new Error('고치기 결과가 바이트가 아닙니다');
      const applied = (Array.isArray(r.applied) ? r.applied : Array.isArray(r.fixes) ? r.fixes : []).map(String);
      if (out !== bytes) {
        const why = checkOoxmlBytes(out, ext);
        if (why) throw new Error(`고친 결과가 형식에 맞지 않습니다: ${why}`);
      }
      if (applied.length) log(`저장 뒤 고치기 (${ext}):`, applied.join(', '));
      if (Array.isArray(r.warnings) && r.warnings.length) log(`저장 뒤 고치기 경고 (${ext}):`, r.warnings.map(String).join(' / '));
      // 모듈이 실패를 알려 주면 (원래 바이트를 돌려줌): 로그, pptx(또는 notify)면 알림 띠
      if (typeof r.error === 'string' && r.error) {
        log(`저장 뒤 고치기 실패 (${ext}):`, r.error);
        return { bytes: out, applied, warning: r.notify === true || ext === 'pptx' ? r.error : null };
      }
      return { bytes: out, applied, warning: null };
    } catch (err) {
      log(`저장 뒤 고치기 실패 (${ext}):`, err?.message ?? err);
      return { bytes, applied: [], warning: warn };
    }
  }

  // ── IPC ──────────────────────────────────────────────

  /** 보낸 쪽이 편집 창의 host 페이지(최상위 틀)인지 → 그 창 정보 */
  function senderEntry(e) {
    const entry = byWc.get(e.sender?.id);
    if (!entry || !entry.win || entry.win.isDestroyed()) return null;
    const frame = e.senderFrame;
    if (!frame || frame !== entry.win.webContents.mainFrame) return null;
    try {
      const u = new URL(frame.url);
      if (u.origin !== server?.origin) return null;
      if (!u.pathname.startsWith(entry.kind === 'office' ? '/office-host/' : '/host/')) return null;
    } catch { return null; }
    return entry;
  }

  function handle(channel, fn, { kinds = ['hwp', 'office'] } = {}) {
    ipcMain.handle(channel, async (e, ...args) => {
      try {
        const entry = senderEntry(e);
        if (!entry || !kinds.includes(entry.kind)) throw fail('허용되지 않은 요청입니다');
        return { ok: true, value: await fn(entry, ...args) };
      } catch (err) {
        if (!err?.expected) log(`[${channel}]`, err);
        return { ok: false, error: err?.message ?? String(err), code: err?.code ?? null };
      }
    });
  }
  function on(channel, fn) {
    ipcMain.on(channel, (e, ...args) => {
      const entry = senderEntry(e);
      if (!entry) return;
      try { fn(entry, ...args); } catch (err) { log(`[${channel}]`, err); }
    });
  }

  handle('editor:get-document', async (entry) => {
    const st = statOf(entry.path);
    if (!st) throw fail(`문서를 찾을 수 없습니다: ${basename(entry.path)}`, 'NOT_FOUND');
    if (st.size > maxBytes) throw fail('문서가 너무 커서 편집기로 열 수 없습니다 (512MB 넘음)', 'TOO_BIG');
    if (entry.kind === 'office') {
      const srv = await ensureServer();
      // 글꼴 목록이 준비된 뒤에 편집기 틀을 만들게 (첫 창만 약 1.5초). 너무 오래 걸리면 묶음 글꼴로 먼저 연다
      let timer;
      const late = new Promise((r) => { timer = setTimeout(() => r(false), fontsTimeoutMs); });
      const ready = await Promise.race([ensureOfficeFonts(srv), late]);
      clearTimeout(timer);
      if (!ready && !fontsWarning) log('글꼴 목록 준비가 늦어 묶음 글꼴로 먼저 엽니다');
    }
    const buf = readFileSync(entry.path);
    entry.known = { mtimeMs: st.mtimeMs, size: buf.length, hash: hashOf(buf), missing: false };
    const ext = extOf(entry.path);
    const base = {
      kind: entry.kind, name: basename(entry.path), ext, size: buf.length, mtimeMs: st.mtimeMs, readOnly: st.readOnly,
      bytes: new Uint8Array(buf), // 복사본 (Buffer 풀의 다른 내용이 함께 가지 않게)
    };
    if (entry.kind !== 'office') {
      const ok = available('hwp');
      return { ...base, available: ok.available, reason: ok.reason, studioUrl: server.studioUrl };
    }
    const ok = available('office');
    const draft = draftOffer(entry, st);
    // 복구 / 버리기를 고르기 전에 창을 닫아도 사본을 지우지 않는다 (다음에 다시 묻는다)
    entry.draftPending = !!draft;
    return {
      ...base, available: ok.available, reason: ok.reason,
      appUrl: server.officeUrl, locale: 'ko-KR', mime: OOXML_MIME[ext] ?? '', title: labelOf('office', ext),
      theme: nativeTheme?.shouldUseDarkColors ? 'dark' : 'light',
      userName: await editorUserName(),
      fontsWarning,
      draft,
      about: aboutInfo(),
    };
  });

  // 자동 저장 복구 사본 (op: 'read' → { bytes, size, mtimeMs } | 'discard' → { discarded: true })
  handle('editor:draft', async (entry, op) => {
    if (op === 'discard') { entry.draftPending = false; removeDraft(entry.path); return { discarded: true }; }
    if (op !== 'read') throw fail('알 수 없는 요청입니다', 'BAD_OP');
    const d = readDraft(entry);
    if (!d) throw fail('복구할 자동 저장 사본이 없습니다', 'NO_DRAFT');
    entry.draftPending = false; // 복구함 → 편집 창이 저장 안 한 내용으로 들고 있다 (닫을 때 저장 여부를 묻는다)
    return d;
  }, { kinds: ['office'] });

  handle('editor:save', async (entry, bytes, opts) => {
    if (entry.kind === 'office') {
      if (opts?.reason === 'auto') return saveDraft(entry, bytes);
      return saveTo(entry, entry.path, bytes, extOf(entry.path));
    }
    return saveTo(entry, entry.path, bytes, extOf(entry.path) === 'hwpx' ? 'hwpx' : 'hwp');
  });

  handle('editor:save-as', async (entry, bytes, ext) => {
    const office = entry.kind === 'office';
    const cur = extOf(entry.path);
    if (office) {
      // 같은 형식만 (docx → docx). 다른 형식은 '내려받기'(editor:download)
      if (ext !== cur) throw fail(`${TYPE_NAMES[cur]}(.${cur})로만 다른 이름으로 저장할 수 있습니다. 다른 형식은 ‘다른 형식으로 내려받기’를 쓰세요.`, 'UNSUPPORTED');
    } else if (!SAVE_AS_EXTS.includes(ext)) {
      throw fail('한글 문서(.hwp, .hwpx)로만 저장할 수 있습니다');
    }
    checkBytes(entry, bytes, ext);
    const stem = basename(entry.path).replace(/\.[^.]+$/, '');
    const target = await askTarget(entry, {
      title: '다른 이름으로 저장', defaultPath: join(dirname(entry.path), `${stem}.${ext}`), ext,
      filterName: office ? `${TYPE_NAMES[ext]} (.${ext})` : ext === 'hwpx' ? '한글 문서 (HWPX)' : '한글 문서 (HWP)',
    });
    if (!target) return { canceled: true };
    const other = entries.get(keyOf(target));
    if (other && other !== entry) throw fail('그 문서는 다른 편집 창에서 열려 있습니다. 그 창을 닫고 다시 저장하세요.', 'IN_USE');
    const oldPath = entry.path;
    const res = await saveTo(entry, target, bytes, ext);
    if (keyOf(target) !== entry.key) {
      rekey(entry, target);
      // 편집 내용은 새 파일에 저장됐다. 단 '복구 / 버리기'에 답하지 않은 사본은 옛 파일 몫으로 남긴다
      if (office && !entry.draftPending) removeDraft(oldPath);
      entry.draftPending = false;
    }
    return res;
  });

  // 다른 형식으로 내려받기 (PDF 등) → main의 저장 대화상자
  handle('editor:download', async (entry, name, bytes, mime) => {
    void mime;
    const fallback = `${basename(entry.path).replace(/\.[^.]+$/, '')}.pdf`;
    let file = safeFileName(name, fallback);
    let ext = extOf(file);
    if (!ext) { ext = 'pdf'; file = `${file}.pdf`; }
    if (!DOWNLOAD_EXTS.includes(ext)) throw fail(`이 형식(.${ext})으로는 내려받을 수 없습니다.`, 'UNSUPPORTED');
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) throw fail('내려받을 내용이 비어 있습니다');
    if (bytes.length > maxBytes) throw fail('내용이 너무 커서 저장할 수 없습니다 (512MB 넘음)', 'TOO_BIG');
    if (!downloadBytesOk(bytes, ext)) throw fail(`편집기가 내보낸 내용이 .${ext} 형식이 아니라 저장하지 않았습니다`, 'BAD_FORMAT');
    const target = await askTarget(entry, {
      title: '다른 형식으로 내려받기', defaultPath: join(dirname(entry.path), file), ext, filterName: `${TYPE_NAMES[ext]} (.${ext})`,
    });
    if (!target) return { canceled: true };
    if (entries.has(keyOf(target))) throw fail('편집 창에서 열려 있는 문서에는 내려받을 수 없습니다. 다른 이름을 고르세요.', 'IN_USE');
    const res = await writeFile(target, bytes);
    log(`내려받기: ${target} (${bytes.length}바이트)`);
    try { onDownloaded?.(target); } catch (err) { log('onDownloaded 오류:', err); }
    return res;
  }, { kinds: ['office'] });

  // 인쇄: 편집기가 만든 PDF → 임시 파일 → 기본 PDF 앱으로 (PDF 앱이 없으면 저장 대화상자)
  // 편집기 틀이 만든 PDF가 이 창 밖(기본 PDF 앱)으로 나가는 유일한 길이므로: 한 번에 하나, 3초 간격, 창마다 처음 한 번 묻는다
  handle('editor:print', async (entry, pdf) => {
    if (!(pdf instanceof Uint8Array) || pdf.length === 0) throw fail('인쇄할 내용이 비어 있습니다');
    if (pdf.length > maxBytes) throw fail('인쇄할 내용이 너무 큽니다 (512MB 넘음)', 'TOO_BIG');
    if (!downloadBytesOk(pdf, 'pdf')) throw fail('편집기가 만든 인쇄 파일이 PDF가 아닙니다', 'BAD_FORMAT');
    if (entry.printing || Date.now() - (entry.lastPrintAt ?? 0) < printIntervalMs) {
      throw fail('인쇄를 준비하는 중입니다. 잠시 뒤 다시 시도하세요.', 'BUSY');
    }
    entry.printing = true;
    try { return await printPdf(entry, pdf); } finally { entry.printing = false; entry.lastPrintAt = Date.now(); }
  }, { kinds: ['office'] });

  async function printPdf(entry, pdf) {
    if (!entry.printAllowed) {
      const { response } = await dialog.showMessageBox(entry.win, {
        type: 'question', title: labelOfEntry(entry),
        message: '인쇄용 PDF를 기본 PDF 앱으로 열까요?',
        detail: 'PDF 앱에서 인쇄합니다. 이 창에서는 다시 묻지 않습니다.',
        buttons: ['열기', '취소'], defaultId: 0, cancelId: 1, noLink: true,
      });
      if (response !== 0) return { canceled: true };
      entry.printAllowed = true;
    }
    const stem = safeFileName(basename(entry.path).replace(/\.[^.]+$/, ''), '문서');
    cleanPrintDir();
    const dir = join(printRoot(), randomBytes(6).toString('hex'));
    mkdirSync(dir, { recursive: true });
    const file = join(dir, `${stem}.pdf`);
    await atomicWrite(file, Buffer.from(pdf.buffer, pdf.byteOffset, pdf.byteLength));
    const err = await shell?.openPath?.(file);
    if (typeof err === 'string' && !err) return { opened: true, path: file };
    log('PDF 앱으로 열지 못했습니다:', err);
    const target = await askTarget(entry, {
      title: '인쇄용 PDF 저장', defaultPath: join(dirname(entry.path), `${stem}.pdf`), ext: 'pdf', filterName: 'PDF 문서 (.pdf)',
    });
    if (!target) return { canceled: true };
    if (entries.has(keyOf(target))) throw fail('편집 창에서 열려 있는 문서에는 저장할 수 없습니다. 다른 이름을 고르세요.', 'IN_USE');
    return { opened: false, ...(await writeFile(target, pdf)) };
  }

  // 편집기 설정 보관 (op: 'get' | 'set')
  handle('editor:state', async (_entry, op, key, value) => {
    if (typeof key !== 'string' || !/^[\w.:-]{1,128}$/.test(key)) throw fail('잘못된 설정 이름입니다', 'BAD_KEY');
    if (op === 'get') {
      const v = officeState ? await officeState.get(key) : memoryState.get(key);
      return v === undefined ? null : v;
    }
    if (op !== 'set') throw fail('알 수 없는 요청입니다', 'BAD_OP');
    let json;
    try { json = JSON.stringify(value ?? null); } catch { throw fail('저장할 수 없는 설정 값입니다', 'BAD_VALUE'); }
    if (typeof json !== 'string') throw fail('저장할 수 없는 설정 값입니다', 'BAD_VALUE');
    if (Buffer.byteLength(json) > STATE_VALUE_MAX) throw fail('설정 값이 너무 큽니다 (8KB 넘음)', 'TOO_BIG');
    const v = JSON.parse(json);
    if (officeState) await officeState.set(key, v);
    else if (v === null) memoryState.delete(key);
    else if (memoryState.size < 256 || memoryState.has(key)) memoryState.set(key, v);
    return null;
  }, { kinds: ['office'] });

  on('editor:dirty', (entry, dirty) => {
    dirty = dirty === true;
    if (entry.dirty === dirty) return;
    entry.dirty = dirty;
    updateTitle(entry);
    try { onDirtyChange?.(entry.path, dirty); } catch {}
  });
  on('editor:log', (entry, message) => log(`(${basename(entry.path)})`, String(message ?? '').slice(0, 2000)));
  on('editor:close', (entry) => { entry.win.close(); });
  on('editor:reply', (entry, msg) => {
    const p = pending.get(msg?.id);
    if (!p || p.entry !== entry) return;
    pending.delete(msg.id);
    clearTimeout(p.timer);
    p.resolve(typeof msg.value === 'boolean' ? msg.value : null);
  });

  /** host 페이지에 묻고 답(boolean | null)을 기다린다. 시간이 지나거나 창이 없어지면 null */
  function ask(entry, channel, payload = {}, timeoutMs = REPLY_TIMEOUT_MS) {
    if (!entry.win || entry.win.isDestroyed()) return Promise.resolve(null);
    const id = `${++seq}`;
    return new Promise((res) => {
      const timer = setTimeout(() => { pending.delete(id); res(null); }, timeoutMs);
      pending.set(id, { resolve: res, timer, entry });
      try { entry.win.webContents.send(channel, { ...payload, id }); } catch { clearTimeout(timer); pending.delete(id); res(null); }
    });
  }

  // ── 저장 ─────────────────────────────────────────────

  function checkBytes(entry, bytes, ext) {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) throw fail('저장할 내용이 비어 있습니다');
    if (bytes.length > maxBytes) throw fail('문서가 너무 커서 저장할 수 없습니다 (512MB 넘음)');
    if (entry.kind === 'office') {
      const why = checkOoxmlBytes(bytes, ext);
      if (why) throw fail(`편집기가 내보낸 내용이 올바른 ${TYPE_NAMES[ext] ?? ''}(.${ext})가 아니라 저장하지 않았습니다: ${why}`, 'BAD_FORMAT');
      return;
    }
    if (!bytesMatchExt(bytes, ext)) throw fail(`편집기가 내보낸 내용이 .${ext} 형식이 아니라 저장하지 않았습니다`, 'BAD_FORMAT');
  }

  function downloadBytesOk(bytes, ext) {
    if (ZIP_TYPES.has(ext)) return startsWith(bytes, ZIP);
    if (ext === 'pdf') return startsWith(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d]); // %PDF-
    if (ext === 'png') return startsWith(bytes, [0x89, 0x50, 0x4e, 0x47]);
    if (ext === 'jpg' || ext === 'jpeg') return startsWith(bytes, [0xff, 0xd8]);
    return true;
  }

  function backupIfExists(path) {
    if (!getBackups || !statOf(path)) return null;
    try {
      const store = typeof getBackups === 'function' ? getBackups() : getBackups;
      const b = store?.backup?.(path);
      return b?.file ?? b?.path ?? null;
    } catch (err) {
      log('백업 실패:', err?.message ?? err);
      return null;
    }
  }

  /** 저장 대화상자 → 절대 경로 | null (확장자가 없으면 붙이고, 붙인 이름이 이미 있으면 다시 묻는다) */
  async function askTarget(entry, { title, defaultPath, ext, filterName }) {
    const r = await dialog.showSaveDialog(entry.win, {
      title, defaultPath, filters: [{ name: filterName, extensions: [ext] }], properties: ['showOverwriteConfirmation'],
    });
    if (r.canceled || !r.filePath) return null;
    let target = resolve(r.filePath);
    if (extOf(target) !== ext) {
      target = `${target}.${ext}`;
      // 덮어쓰기 확인은 사용자가 적은 이름으로 받았으므로, 확장자를 붙인 이름이 이미 있으면 다시 묻는다
      if (statOf(target)) {
        const { response } = await dialog.showMessageBox(entry.win, {
          type: 'warning', title,
          message: `‘${basename(target)}’ 파일이 이미 있습니다. 바꿀까요?`,
          detail: '바꾸면 원래 파일의 내용은 백업 폴더에만 남습니다.',
          buttons: ['바꾸기', '취소'], defaultId: 1, cancelId: 1, noLink: true,
        });
        if (response !== 0) return null;
      }
    }
    return target;
  }

  /** 편집 중인 문서가 아닌 파일 쓰기 (내려받기·인쇄 PDF 저장): 있던 파일은 백업 → { saved: true, name, ext, size, path } */
  async function writeFile(target, bytes) {
    const name = basename(target);
    if (statOf(target)?.readOnly) throw fail(`‘${name}’ 파일은 읽기 전용이라 저장할 수 없습니다. 다른 이름을 고르세요.`, 'READ_ONLY');
    backupIfExists(target);
    const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    try { await atomicWrite(target, buf); } catch (err) { throw saveError(err, name); }
    return { saved: true, name, ext: extOf(target), size: buf.length, path: target };
  }

  async function saveTo(entry, target, bytes, ext) {
    checkBytes(entry, bytes, ext);
    const name = basename(target);
    const before = statOf(target);
    if (before?.readOnly) throw fail(`‘${name}’ 파일은 읽기 전용이라 저장할 수 없습니다. 다른 이름으로 저장하세요.`, 'READ_ONLY');
    if (entry.saving) throw fail('저장하는 중입니다. 잠시 뒤 다시 시도하세요.', 'BUSY');
    entry.saving = true;
    try {
      // Office: 형식 확인 → 저장 뒤 고치기(편집기가 열었던 문서의 지금 바이트를 원본으로) → 백업 → 쓰기
      let out = bytes; let applied = []; let warning = null;
      if (entry.kind === 'office') ({ bytes: out, applied, warning } = await runFixups(ext, bytes, readOriginal(entry.path)));
      // 덮어쓰기 전의 파일을 백업해 둔다 (실수로 저장해도 백업 폴더에서 되찾게)
      const backup = backupIfExists(target);
      const buf = Buffer.from(out.buffer, out.byteOffset, out.byteLength);
      try { await atomicWrite(target, buf); } catch (err) { throw saveError(err, name); }
      const st = statOf(target);
      if (keyOf(target) === entry.key) entry.known = { mtimeMs: st?.mtimeMs ?? 0, size: buf.length, hash: hashOf(buf), missing: false };
      // 정상 저장 → 복구 사본은 필요 없다 (복구 / 버리기를 아직 고르지 않았으면 남긴다)
      if (entry.kind === 'office' && keyOf(target) === entry.key && !entry.draftPending) removeDraft(target);
      log(`저장: ${target} (${buf.length}바이트)`);
      try { onSaved?.(target); } catch (err) { log('onSaved 오류:', err); }
      const res = { name, ext: extOf(target), size: buf.length, mtimeMs: st?.mtimeMs ?? null, backup };
      return entry.kind === 'office' ? { ...res, fixups: applied, warning } : res;
    } finally {
      entry.saving = false;
    }
  }

  /** 저장 뒤 고치기의 원본: 편집기가 열었던 문서 파일의 지금 바이트 (없거나 너무 크면 null) */
  function readOriginal(path) {
    try {
      const st = statSync(path);
      if (!st.isFile() || st.size === 0 || st.size > maxBytes) return null;
      return readFileSync(path);
    } catch { return null; }
  }

  /** 편집기 사용자 이름 (한 번만 읽는다, 1.5초 안에 못 읽으면 Windows 계정 이름) */
  async function editorUserName() {
    if (typeof userName === 'string') return cleanUserName(userName);
    userNamePromise ??= Promise.resolve().then(() => (typeof userName === 'function' ? userName() : readOfficeUserName()))
      .then(cleanUserName).catch(() => '');
    let timer;
    const late = new Promise((r) => { timer = setTimeout(() => r(null), 1500); });
    const v = await Promise.race([userNamePromise, late]);
    clearTimeout(timer);
    if (v) return v;
    try { return cleanUserName(userInfo().username); } catch { return ''; }
  }

  // ── 자동 저장 (Office) ────────────────────────────────
  // 사용자 파일을 몰래 덮어쓰지 않는다: 편집기의 자동 저장(reason:'auto')은 userData/office-drafts/ 사본으로만

  async function saveDraft(entry, bytes) {
    checkBytes(entry, bytes, extOf(entry.path));
    if (!userData) return { draft: false };
    if (entry.drafting) throw fail('자동 저장하는 중입니다.', 'BUSY');
    entry.drafting = true;
    try {
      const file = draftPathFor(userData, entry.path);
      mkdirSync(dirname(file), { recursive: true });
      const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      try { await atomicWrite(file, buf); } catch (err) { throw saveError(err, basename(file)); }
      const st = statOf(file);
      log(`자동 저장 사본: ${basename(entry.path)} → ${file} (${buf.length}바이트)`);
      return { draft: true, size: buf.length, mtimeMs: st?.mtimeMs ?? null };
    } finally {
      entry.drafting = false;
    }
  }

  function removeDraft(docPath) {
    if (!userData) return;
    try { rmSync(draftPathFor(userData, docPath), { force: true }); } catch (err) { log('복구 사본을 지우지 못했습니다:', err?.message ?? err); }
  }

  /** 복구 사본이 파일보다 새로우면 → { size, mtimeMs } (host가 '복구 / 버리기' 띠를 띄운다) */
  function draftOffer(entry, st) {
    if (!userData) return null;
    const d = statOf(draftPathFor(userData, entry.path));
    if (!d || d.size === 0 || d.size > maxBytes || d.mtimeMs <= st.mtimeMs) return null;
    return { size: d.size, mtimeMs: d.mtimeMs };
  }

  /** 복구 사본 읽기 → { bytes, size, mtimeMs } | null (형식이 맞지 않으면 null) */
  function readDraft(entry) {
    if (!userData) return null;
    const file = draftPathFor(userData, entry.path);
    const d = statOf(file);
    if (!d || d.size === 0 || d.size > maxBytes) return null;
    try {
      const bytes = new Uint8Array(readFileSync(file));
      if (checkOoxmlBytes(bytes, extOf(entry.path))) return null;
      return { bytes, size: bytes.length, mtimeMs: d.mtimeMs };
    } catch { return null; }
  }

  /** '정보 · 라이선스'용: { appVersion, bundle(VERSION.json, 파일 목록 뺌), sources } */
  let bundleInfo;
  function aboutInfo() {
    if (bundleInfo === undefined) {
      bundleInfo = null;
      try {
        const v = JSON.parse(readFileSync(join(editorDirs(root, dirOverrides).office, 'VERSION.json'), 'utf8'));
        if (v && typeof v === 'object' && !Array.isArray(v)) {
          const { files: _files, ...rest } = v;
          bundleInfo = rest;
        }
      } catch {}
    }
    let appVersion = null;
    try { appVersion = app?.getVersion?.() ?? null; } catch {}
    return { appVersion, bundle: bundleInfo, sources: Array.isArray(bundleInfo?.sources) ? bundleInfo.sources : [] };
  }

  // ── 인쇄 임시 파일 ────────────────────────────────────

  const printRoot = () => join(tmpdir(), 'officekit-print');
  /** 한 시간 넘은 인쇄용 PDF 지우기 (PDF 앱이 아직 잡고 있으면 다음에) */
  function cleanPrintDir(maxAgeMs = 60 * 60 * 1000) {
    let names = [];
    try { names = readdirSync(printRoot()); } catch { return; }
    for (const n of names) {
      const p = join(printRoot(), n);
      try { if (Date.now() - statSync(p).mtimeMs > maxAgeMs) rmSync(p, { recursive: true, force: true }); } catch {}
    }
  }

  // 다른 이름으로 저장 뒤: 이 창은 새 파일을 편집한다
  function rekey(entry, target) {
    stopWatch(entry);
    entries.delete(entry.key);
    entry.path = target;
    entry.key = keyOf(target);
    entries.set(entry.key, entry);
    const buf = readFileSync(target);
    const st = statOf(target);
    entry.known = { mtimeMs: st?.mtimeMs ?? 0, size: buf.length, hash: hashOf(buf), missing: false };
    updateTitle(entry);
    startWatch(entry);
  }

  // ── 바깥 변경 감시 ────────────────────────────────────
  // 편집기가 새 파일로 바꿔치기하는 경우도 잡도록 폴더를 지켜본다. 내용(해시)이 같으면 알리지 않는다 (자기 저장 포함)

  /**
   * force: 내용이 같아도 알린다 ('원래대로'를 눌렀는데 편집 창에 저장 안 한 편집이 있을 때 → 편집 창이 다시 불러올지 묻는다)
   * reason: 편집 창 안내 문구용 ('restore' = 주 창의 원래대로)
   */
  function checkExternal(entry, { force = false, reason = null } = {}) {
    if (!entry.win || entry.win.isDestroyed() || entry.saving || !entry.known) return false;
    const st = statOf(entry.path);
    if (!st) {
      if (entry.known.missing) return false;
      entry.known = { ...entry.known, missing: true };
      entry.win.webContents.send('editor:external-change', { missing: true, size: 0, mtimeMs: null });
      return true;
    }
    if (!force && !entry.known.missing && st.mtimeMs === entry.known.mtimeMs && st.size === entry.known.size) return false;
    let hash;
    try { hash = hashOf(readFileSync(entry.path)); } catch { return false; } // 쓰는 중이면 다음 알림에서
    const { hash: before, missing: wasMissing } = entry.known;
    entry.known = { mtimeMs: st.mtimeMs, size: st.size, hash, missing: false };
    if (!force && hash === before && !wasMissing) return false;
    entry.win.webContents.send('editor:external-change', { missing: false, size: st.size, mtimeMs: st.mtimeMs, reason });
    return true;
  }

  function startWatch(entry) {
    const name = basename(entry.path).toLowerCase();
    try {
      entry.watcher = watch(dirname(entry.path), { persistent: false }, (_ev, file) => {
        if (file && String(file).toLowerCase() !== name) return;
        clearTimeout(entry.watchTimer);
        entry.watchTimer = setTimeout(() => checkExternal(entry), 600);
      });
      entry.watcher.on('error', () => {});
    } catch { entry.watcher = null; }
  }
  function stopWatch(entry) {
    clearTimeout(entry.watchTimer);
    try { entry.watcher?.close(); } catch {}
    entry.watcher = null;
  }

  // ── 창 ───────────────────────────────────────────────

  const labelOfEntry = (entry) => labelOf(entry.kind, extOf(entry.path));
  function titleOf(entry) { return `${entry.dirty ? '* ' : ''}${basename(entry.path)} - ${labelOfEntry(entry)} (Office Kit)`; }
  function updateTitle(entry) { if (entry.win && !entry.win.isDestroyed()) entry.win.setTitle(titleOf(entry)); }

  function focus(entry) {
    const w = entry.win;
    if (!w || w.isDestroyed()) return;
    if (w.isMinimized()) w.restore();
    w.show();
    w.focus();
  }

  /** 그 종류의 편집기 틀(iframe)이 가도 되는 주소인지 */
  function isEditorUrl(url, kind) {
    try {
      const u = new URL(url);
      if (kind === 'office') return !!server?.officeOrigin && u.origin === server.officeOrigin && u.pathname.startsWith(`/office/${server.officeKey}/`);
      return u.origin === server?.studioOrigin && u.pathname.startsWith('/studio/');
    } catch { return false; }
  }

  /** 닫기 흐름: 저장 안 한 변경이 있으면 저장 / 저장 안 함 / 취소 → 닫혔으면 true */
  async function requestClose(entry) {
    if (!entry.win || entry.win.isDestroyed()) return true;
    if (entry.closing) return entry.closing;
    entry.closing = (async () => {
      const answer = await ask(entry, 'editor:request-close');
      const dirty = answer ?? entry.dirty;
      if (dirty) {
        focus(entry);
        const { response } = await dialog.showMessageBox(entry.win, {
          type: 'warning', title: labelOfEntry(entry),
          message: `‘${basename(entry.path)}’의 바뀐 내용을 저장할까요?`,
          detail: '저장하지 않으면 마지막으로 저장한 뒤 고친 내용이 사라집니다.',
          buttons: ['저장', '저장 안 함', '취소'], defaultId: 0, cancelId: 2, noLink: true,
        });
        if (response === 2) return false;
        if (response === 0) {
          const saved = await ask(entry, 'editor:command', { command: 'save' }, SAVE_TIMEOUT_MS);
          if (saved !== true) return false; // 저장 실패는 창에 알림 띠로 보인다
        }
      }
      if (!entry.win || entry.win.isDestroyed()) return true;
      const closed = new Promise((r) => entry.win.once('closed', () => r(true)));
      entry.allowClose = true;
      entry.cleanClose = true; // 정상 닫기 → 복구 사본을 지운다
      entry.win.close();
      return closed;
    })();
    try { return await entry.closing; } finally { entry.closing = null; }
  }

  // ── 문서 안 링크 ─────────────────────────────────────
  // 사람이 누른 링크든 스크립트가 연 창이든, 이 PC·사설망 주소는 열지 않고 나머지는 물어본 뒤 기본 브라우저로

  let linkAsking = false;
  async function openLink(entry, url) {
    let u;
    try { u = new URL(String(url)); } catch { return false; }
    if (!['http:', 'https:', 'mailto:'].includes(u.protocol) || u.username || u.password) { log('열지 않은 주소:', String(url).slice(0, 200)); return false; }
    const local = u.protocol !== 'mailto:' && (isPrivateHost(u.hostname)
      || (server && [server.origin, server.studioOrigin, server.officeOrigin].includes(u.origin)));
    if (local) { log('열지 않은 주소(이 PC·사설망):', u.href.slice(0, 200)); return false; }
    if (linkAsking || !entry.win || entry.win.isDestroyed()) return false; // 한 번에 하나만 묻는다
    linkAsking = true;
    try {
      const { response } = await dialog.showMessageBox(entry.win, {
        type: 'question', title: labelOfEntry(entry),
        message: u.protocol === 'mailto:' ? '이 주소로 메일을 쓸까요?' : '이 주소를 기본 브라우저로 열까요?',
        detail: u.href.length > 500 ? `${u.href.slice(0, 500)}…` : u.href,
        buttons: ['열기', '취소'], defaultId: 0, cancelId: 1, noLink: true,
      });
      if (response !== 0) return false;
      await shell?.openExternal(u.href);
      return true;
    } catch (err) {
      log('링크를 열지 못했습니다:', err?.message ?? err);
      return false;
    } finally {
      linkAsking = false;
    }
  }

  /** 편집기가 연 빈 창: 주소를 바꾸려는 순간 막고 그 주소를 openLink로 넘긴 뒤 창을 없앤다 */
  function catchLinkWindow(entry, child) {
    const cwc = child.webContents;
    let taken = false;
    const take = (e, url) => {
      if (!url || url === 'about:blank') return;
      try { e?.preventDefault?.(); } catch {}
      if (taken) return;
      taken = true;
      void openLink(entry, url);
      setImmediate(() => { if (!child.isDestroyed()) child.destroy(); });
    };
    cwc.on('will-navigate', (e, url) => take(e, e?.url ?? url));
    cwc.on('will-frame-navigate', (e) => take(e, e?.url));
    cwc.on('will-redirect', (e, url) => take(e, e?.url ?? url));
    cwc.setWindowOpenHandler(() => ({ action: 'deny' }));
    // 아무 데도 가지 않는 빈 창은 곧 없앤다
    const timer = setTimeout(() => { if (!child.isDestroyed()) child.destroy(); }, 5000);
    child.once('closed', () => clearTimeout(timer));
  }

  function createWindow(entry) {
    const dark = !!nativeTheme?.shouldUseDarkColors;
    const office = entry.kind === 'office';
    const partition = PARTITIONS[entry.kind];
    const win = new BrowserWindow({
      width: 1280, height: 900, minWidth: 720, minHeight: 520,
      title: titleOf(entry), show: false,
      backgroundColor: dark ? '#15171c' : '#f3f4f6',
      autoHideMenuBar: true,
      ...(icon ? { icon } : {}),
      webPreferences: {
        preload: office ? join(officeHostDir(), 'office-preload.cjs') : join(here, 'editor', 'editor-preload.cjs'),
        contextIsolation: true, sandbox: true, nodeIntegration: false, nodeIntegrationInSubFrames: false,
        partition,
        webviewTag: false, spellcheck: false, navigateOnDragDrop: false,
        // 편집기 메뉴·오른쪽 클릭의 '붙이기'는 execCommand('paste')를 쓴다 (Ctrl+V는 이것 없이도 된다)
        enableDeprecatedPaste: true,
        devTools: !app?.isPackaged,
      },
    });
    entry.win = win;
    const wc = win.webContents;
    byWc.set(wc.id, entry);
    // 앱 메뉴 단축키(Ctrl+O, F5 …)가 이 창에서 주 창 일을 하지 않게
    win.removeMenu();
    win.once('ready-to-show', () => win.show());
    win.on('page-title-updated', (e) => e.preventDefault());

    // WebRTC가 프록시를 거치지 않는 UDP(STUN)로 밖에 나가지 못하게 (TCP는 세션의 죽은 프록시로 간다)
    try { wc.setWebRTCIPHandlingPolicy('disable_non_proxied_udp'); } catch {}

    // 새 창 막기: 웹 주소는 물어본 뒤 기본 브라우저로.
    // 문서 링크는 빈 창(about:blank)을 연 뒤 그 창의 주소를 바꾸기도 하므로, 빈 창만 숨겨서 열고 이동을 가로챈다
    wc.setWindowOpenHandler(({ url }) => {
      if (url === 'about:blank') {
        return {
          action: 'allow',
          overrideBrowserWindowOptions: {
            show: false, skipTaskbar: true, width: 200, height: 100,
            webPreferences: {
              partition, sandbox: true, contextIsolation: true, nodeIntegration: false,
              javascript: false, webviewTag: false, spellcheck: false, devTools: false,
            },
          },
        };
      }
      void openLink(entry, url);
      return { action: 'deny' };
    });
    wc.on('did-create-window', (child) => {
      try { child.webContents.setWebRTCIPHandlingPolicy('disable_non_proxied_udp'); } catch {}
      catchLinkWindow(entry, child);
    });
    wc.on('will-navigate', (e) => e.preventDefault());
    wc.on('will-frame-navigate', (e) => {
      if (e.isMainFrame) { e.preventDefault(); return; }
      const u = e.url;
      if (u === 'about:blank' || u === 'about:srcdoc' || isEditorUrl(u, entry.kind)) return;
      e.preventDefault();
    });
    wc.on('will-attach-webview', (e) => e.preventDefault());
    // 편집기의 beforeunload(저장 안 함 경고)는 무시하고 닫기 확인은 여기서 한다
    wc.on('will-prevent-unload', (e) => e.preventDefault());
    // 단축키: 편집기(iframe) 안에 커서가 있어도 먼저 받는다
    wc.on('before-input-event', (e, input) => {
      if (input.type !== 'keyDown') return;
      const code = input.code || '';
      // 개발 중에만: F12 / Ctrl+Shift+I 개발자 도구 (메뉴를 뗐으므로)
      if (!app?.isPackaged && (code === 'F12' || ((input.control || input.meta) && input.shift && code === 'KeyI'))) {
        e.preventDefault();
        wc.toggleDevTools();
        return;
      }
      if (!(input.control || input.meta) || input.alt) return;
      if (code === 'KeyS') {
        e.preventDefault();
        void ask(entry, 'editor:command', { command: input.shift ? 'save-as' : 'save' }, SAVE_TIMEOUT_MS);
      } else if (code === 'KeyW' && !input.shift) {
        e.preventDefault();
        win.close();
      }
    });
    wc.on('render-process-gone', (_e, d) => {
      log('편집 창 화면 종료:', d.reason);
      if (d.reason === 'clean-exit' || win.isDestroyed()) return;
      entry.dirty = false;
      updateTitle(entry);
      try { onDirtyChange?.(entry.path, false); } catch {}
      dialog.showMessageBox(win, {
        type: 'error', title: labelOfEntry(entry), noLink: true, buttons: ['확인'],
        message: '편집 창이 예기치 않게 멈췄습니다.',
        detail: office
          ? '마지막으로 저장한 문서를 다시 엽니다. 자동 저장된 편집 내용이 있으면 다시 열린 창에서 복구할 수 있습니다.'
          : '마지막으로 저장한 문서를 다시 엽니다. 저장하지 않은 내용은 사라졌을 수 있습니다.',
      }).then(() => { if (!win.isDestroyed()) wc.loadURL(hostUrl(entry.kind)); }).catch(() => {});
    });

    win.on('close', (e) => {
      if (entry.allowClose) return;
      e.preventDefault();
      void requestClose(entry).then((closed) => {
        if (!closed) quitting = false;
        else if (quitting) app?.quit();
      });
    });
    win.on('closed', () => {
      stopWatch(entry);
      byWc.delete(wc.id);
      if (entries.get(entry.key) === entry) entries.delete(entry.key);
      for (const [id, p] of pending) if (p.entry === entry) { clearTimeout(p.timer); pending.delete(id); p.resolve(null); }
      // 정상 닫기 → 복구 사본을 지운다. 단 '복구 / 버리기'에 아직 답하지 않았으면 남긴다 (다음에 열 때 다시 묻는다)
      if (office && entry.cleanClose && !entry.draftPending) removeDraft(entry.path);
      entry.win = null;
    });

    wc.loadURL(hostUrl(entry.kind)).catch((err) => log('편집 창을 열지 못했습니다:', err?.message ?? err));
    return win;
  }

  const hostUrl = (kind) => `${kind === 'office' ? server.officeHostUrl : server.hostUrl}${testHooks ? '?test=1' : ''}`;

  /** 문서 경로 확인 → { path: 절대 경로, kind } */
  function editablePath(p) {
    if (typeof p !== 'string' || !p.trim() || p.length > 4096) throw fail('잘못된 파일 경로입니다', 'BAD_PATH');
    if (!isAbsolute(p)) throw fail('파일의 전체 경로가 필요합니다', 'BAD_PATH');
    p = resolve(p);
    let st;
    try { st = statSync(p); } catch { throw fail(`파일을 찾을 수 없습니다: ${basename(p)}`, 'NOT_FOUND'); }
    if (!st.isFile()) throw fail('파일이 아닙니다', 'BAD_PATH');
    const kind = editorKindOf(p);
    if (!kind) {
      throw fail(`한글(.hwp, .hwpx, .hwt)·Word(.docx)·Excel(.xlsx)·PowerPoint(.pptx) 문서만 편집할 수 있습니다 (.${extOf(p) || '확장자 없음'})`, 'UNSUPPORTED');
    }
    return { path: p, kind };
  }

  const openOfficeCount = () => [...entries.values()].filter((e) => e.kind === 'office' && e.win && !e.win.isDestroyed()).length;

  if (app) {
    app.on('before-quit', () => { quitting = true; });
    app.on('will-quit', () => { void server?.close(); cleanPrintDir(); });
  }

  return {
    /**
     * 문서를 편집 창으로 연다 (이미 열려 있으면 그 창을 앞으로) → { path, kind, reused, available, reason, notice }
     * 편집기 묶음이 없으면 창 안에 한국어 이유를 보여 준다 (available:false).
     * Office: 암호 문서·옛 형식·손상·512MB 넘음은 창을 열기 전에 한국어 오류 (code ENCRYPTED·OLD_FORMAT·BAD_FORMAT·TOO_BIG).
     * notice: Office 편집 창이 많이 열렸을 때 메모리 안내 (없으면 null)
     */
    async openEditor(path) {
      const { path: p, kind } = editablePath(path);
      const ok = available(kind);
      const existing = entries.get(keyOf(p));
      if (existing?.win && !existing.win.isDestroyed()) {
        focus(existing);
        return { path: existing.path, kind: existing.kind, reused: true, ...ok, notice: null };
      }
      if (kind === 'office') {
        const bad = officeOpenProblem(p, { maxBytes });
        if (bad) throw fail(bad.message, bad.code);
      }
      const srv = await ensureServer();
      if (kind === 'office' && !srv.officeOrigin) throw fail(srv.officeError ?? '문서 편집기 서버를 시작하지 못했습니다.', 'EDITOR_SERVER');
      await setupSession(srv, kind);
      if (kind === 'office') void ensureOfficeFonts(srv); // 미리 시작 (get-document가 기다린다)
      const again = entries.get(keyOf(p)); // 서버를 기다리는 사이 같은 문서를 또 열었으면
      if (again?.win && !again.win.isDestroyed()) { focus(again); return { path: again.path, kind: again.kind, reused: true, ...ok, notice: null }; }
      const entry = {
        path: p, key: keyOf(p), kind, win: null, dirty: false, known: null, saving: false, drafting: false,
        allowClose: false, cleanClose: false, closing: null, watcher: null, watchTimer: null,
        draftPending: false, printing: false, lastPrintAt: 0, printAllowed: false,
      };
      entries.set(entry.key, entry);
      try { createWindow(entry); } catch (err) { entries.delete(entry.key); throw err; }
      startWatch(entry);
      const count = kind === 'office' ? openOfficeCount() : 0;
      const notice = count > OFFICE_WINDOWS_NOTICE
        ? `문서 편집 창이 ${count}개 열려 있습니다. 창마다 메모리를 약 0.5GB 쓰므로, 쓰지 않는 창은 닫아 주세요.`
        : null;
      return { path: p, kind, reused: false, ...ok, notice };
    },
    /** 그 문서가 편집 창에 열려 있는지 */
    isOpen(path) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return !!(e?.win && !e.win.isDestroyed());
    },
    /** 그 문서에 저장 안 한 변경이 있는지 (host가 알려 준 마지막 상태) */
    isDirty(path) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return !!(e?.win && !e.win.isDestroyed() && e.dirty);
    },
    /**
     * 다른 곳(AI·다른 프로그램)에서 파일을 바꿨을 때 알려 주기 → 편집 창에 알렸으면 true (내용이 같으면 false)
     * force: 내용이 같아도 알린다, reason: 'restore'(원래대로) 등 편집 창 안내 문구용
     */
    notifyExternalChange(path, { force = false, reason = null } = {}) {
      const e = typeof path === 'string' ? entries.get(keyOf(path)) : null;
      return e ? checkExternal(e, { force: force === true, reason: reason === 'restore' ? 'restore' : null }) : false;
    },
    /** 열린 편집 창 문서 경로들 */
    list() { return [...entries.values()].filter((e) => e.win && !e.win.isDestroyed()).map((e) => e.path); },
    /** 편집기 묶음 확인 (kind: 'hwp' 기본 | 'office') → { available, reason } */
    available,
    /** 파일 경로 → 'hwp' | 'office' | null */
    kindOf: editorKindOf,
    /** 그 문서의 자동 저장 복구 사본 경로 (userData가 없으면 null) */
    draftPath: (path) => (userData && typeof path === 'string' ? draftPathFor(userData, path) : null),
    /**
     * 모든 편집 창 닫기 → 다 닫혔으면 true (사용자가 '취소'하면 false)
     * force: 묻지 않고 닫는다 (저장 안 한 변경은 사라짐 — Office 복구 사본은 남긴다)
     */
    async closeAll({ force = false } = {}) {
      const list = [...entries.values()].filter((e) => e.win && !e.win.isDestroyed());
      if (force) {
        for (const e of list) { e.allowClose = true; e.win.destroy(); }
        return true;
      }
      for (const e of list) {
        if (!(await requestClose(e))) return false;
      }
      return true;
    },
    /** 서버 주소 (진단용) */
    serverInfo: () => (server
      ? { origin: server.origin, studioOrigin: server.studioOrigin, officeOrigin: server.officeOrigin, officeVersion: server.officeVersion, officeKey: server.officeKey }
      : null),
    /** 서버 닫기 (창을 다 닫은 뒤) */
    // 서버를 다시 띄우면 포트가 바뀌므로 세션(프록시 건너뛰기 주소:포트)도 다시 만든다
    async dispose() { await server?.close(); server = null; fontsPromise = null; sessionsReady.clear(); },
  };
}
