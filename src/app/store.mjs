// store.mjs — 앱 설정(JSON), 최근 파일, 원본 백업. Electron 없이도 동작하도록 폴더를 주입받는다.
import { createHash } from 'node:crypto';
import {
  copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync, utimesSync, writeFileSync,
} from 'node:fs';
import { basename, join, resolve } from 'node:path';

export const DEFAULT_SETTINGS = Object.freeze({
  provider: null,          // 'claude' | 'codex' | null(자동)
  models: {},              // AI별 모델 { claude: 'sonnet', codex: 'gpt-5' }. 없으면 기본 모델
  previewMode: 'auto',     // 'auto'(빠르게) | 'exact'(설치된 프로그램)
  openAction: 'edit',      // 문서를 열면: 'edit'(편집 창도 바로 — 한글·docx·xlsx·pptx) | 'preview'(미리보기만)
  theme: 'system',         // 'system' | 'light' | 'dark'
  lang: 'system',          // 화면 언어 'system'(Windows 언어) | 'ko' | 'en' | 'ja'
  chatWidth: 400,
  chatCollapsed: false,
  window: null,            // { x, y, width, height, maximized }
  recent: [],              // [{ path, openedAt }]
  officeEditor: { settings: {} }, // 문서·표·슬라이드 편집기가 보관하는 설정 (편집기의 state.get/set, 값마다 8KB)
});

const RECENT_MAX = 10;
const PINNED_MAX = 20;
/** 고정은 PINNED_MAX개, 고정 안 한 것은 RECENT_MAX개까지 (순서는 그대로) */
function keepWithin(list) {
  const pinned = new Set(list.filter((r) => r.pinned).slice(0, PINNED_MAX));
  const recent = new Set(list.filter((r) => !r.pinned).slice(0, RECENT_MAX));
  return list.filter((r) => pinned.has(r) || recent.has(r));
}
/** Office 편집기 설정: 이름 규칙·값 크기(JSON 바이트)·개수 한도 */
export const OFFICE_STATE_KEY = /^[\w.:-]{1,128}$/;
export const OFFICE_STATE_VALUE_MAX = 8 * 1024;
export const OFFICE_STATE_MAX_KEYS = 128;
// 백업(문서 버전): 모두 합쳐 이만큼까지 (개수·바이트). 문서마다 몫은 main이 maxPerSource로 (최근 30개)
export const BACKUP_KEEP = 200;
export const BACKUP_MAX_BYTES = 2 * 1024 * 1024 * 1024;
/** 백업을 만든 까닭 → 버전 목록에 보일 말 (화면도 같은 이름을 쓴다) */
export const BACKUP_REASONS = ['open', 'editor', 'ai', 'restore'];

// 같은 경로인지 (Windows는 대소문자 무시)
export const samePath = (a, b) => process.platform === 'win32'
  ? resolve(a).toLowerCase() === resolve(b).toLowerCase()
  : resolve(a) === resolve(b);

/** 임시 파일에 쓴 뒤 바꿔치기 (쓰다 끊겨도 원본이 깨지지 않게) */
export function writeJsonAtomic(path, data) {
  const tmp = `${path}.${process.pid}.tmp`;
  writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  // 백신 등이 잠깐 잡고 있으면 실패하므로 몇 번 다시 시도
  for (let i = 0; ; i++) {
    try { renameSync(tmp, path); return; } catch (err) {
      if (i >= 4) { rmSync(tmp, { force: true }); throw err; }
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 30 * (i + 1));
    }
  }
}

function readJson(path, fallback) {
  try { return JSON.parse(readFileSync(path, 'utf8').replace(/^﻿/, '')); } catch { return fallback; }
}

/** 모델 이름으로 어느 AI 것인지 짐작 (모르면 null) */
export function guessModelProvider(model) {
  const m = String(model ?? '').trim().toLowerCase();
  if (/^(claude|sonnet|opus|haiku)\b/.test(m)) return 'claude';
  if (/^(gpt|o\d|codex)/.test(m)) return 'codex';
  return null;
}

/** 옛 설정(model 하나)을 AI별 모델로 옮긴다. 어느 AI 것인지 모르면 버린다 (엉뚱한 AI에 넘기지 않게) */
function migrate(data) {
  const models = {};
  if (data.models && typeof data.models === 'object' && !Array.isArray(data.models)) {
    for (const [k, v] of Object.entries(data.models)) if (typeof v === 'string' && v.trim()) models[k] = v.trim();
  }
  if (typeof data.model === 'string' && data.model.trim()) {
    const owner = guessModelProvider(data.model) ?? (typeof data.provider === 'string' ? data.provider : null);
    if (owner && !models[owner]) models[owner] = data.model.trim();
  }
  delete data.model;
  data.models = models;
  if (!Array.isArray(data.recent)) data.recent = [];
  data.officeEditor = { settings: cleanOfficeState(data.officeEditor?.settings) };
  return data;
}

/** 저장된 Office 편집기 설정 걸러 내기 (이름 규칙·크기에 맞는 것만, 자기 속성으로) */
function cleanOfficeState(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  const out = [];
  for (const [k, v] of Object.entries(raw)) {
    if (!OFFICE_STATE_KEY.test(k) || v === null || v === undefined) continue;
    let json;
    try { json = JSON.stringify(v); } catch { continue; }
    if (typeof json === 'string' && Buffer.byteLength(json) <= OFFICE_STATE_VALUE_MAX) out.push([k, v]);
  }
  // Object.fromEntries는 '__proto__' 같은 이름도 자기 속성으로 둔다 (프로토타입을 바꾸지 않음)
  return Object.fromEntries(out.slice(-OFFICE_STATE_MAX_KEYS));
}

/** 설정 저장소. dir = app.getPath('userData') */
export function createSettings(dir) {
  mkdirSync(dir, { recursive: true });
  const file = join(dir, 'settings.json');
  let data = migrate({ ...DEFAULT_SETTINGS, ...readJson(file, {}) });

  const save = () => writeJsonAtomic(file, data);
  return {
    file,
    get: (key) => (key ? data[key] : { ...data }),
    /** 일부만 바꾸기 */
    set(patch) { data = { ...data, ...patch }; save(); return { ...data }; },

    // ── 최근 파일 ───────────────────────────────────────
    /** 고정한 것이 먼저, 그다음 최근에 연 것 */
    recent() {
      const list = data.recent.map((r) => ({ ...r, pinned: r.pinned === true, name: basename(r.path), exists: existsSync(r.path) }));
      return [...list.filter((r) => r.pinned), ...list.filter((r) => !r.pinned)];
    },
    /** 연 문서를 맨 앞에 (고정은 그대로). 고정 안 한 것은 최근 RECENT_MAX개, 고정은 PINNED_MAX개까지 */
    addRecent(path) {
      const old = data.recent.find((r) => samePath(r.path, path));
      const all = [{ path, openedAt: Date.now(), ...(old?.pinned ? { pinned: true } : {}) }, ...data.recent.filter((r) => !samePath(r.path, path))];
      data.recent = keepWithin(all);
      save();
    },
    /** 목록에 고정 · 풀기 (목록에 없으면 넣는다) */
    pinRecent(path, on) {
      let list = data.recent;
      if (!list.some((r) => samePath(r.path, path))) list = [{ path, openedAt: Date.now() }, ...list];
      if (on && list.filter((r) => r.pinned && !samePath(r.path, path)).length >= PINNED_MAX) throw new Error(`고정은 ${PINNED_MAX}개까지 할 수 있어요`);
      list = list.map((r) => {
        if (!samePath(r.path, path)) return r;
        const { pinned: _old, ...rest } = r;
        return on ? { ...rest, pinned: true } : rest;
      });
      data.recent = keepWithin(list);
      save();
    },
    removeRecent(path) { data.recent = data.recent.filter((r) => !samePath(r.path, path)); save(); },
    /** 지우기: 고정한 것은 남긴다 */
    clearRecent() { data.recent = data.recent.filter((r) => r.pinned); save(); },

    // ── Office 편집기 설정 (편집 창의 editor:state) ─────────
    /** 값 하나 → 값(사본) | null */
    officeGet(key) {
      if (typeof key !== 'string' || !OFFICE_STATE_KEY.test(key)) return null;
      const all = data.officeEditor.settings;
      return Object.hasOwn(all, key) ? structuredClone(all[key]) : null;
    },
    /**
     * 값 하나 보관 (null = 지우기). 같은 값이면 파일을 다시 쓰지 않는다.
     * 개수가 한도를 넘으면 오래된 것부터 지운다 → 바뀌었으면 true
     */
    officeSet(key, value) {
      if (typeof key !== 'string' || !OFFICE_STATE_KEY.test(key)) throw new Error('잘못된 설정 이름입니다');
      const all = data.officeEditor.settings;
      const entries = Object.entries(all).filter(([k]) => k !== key);
      if (value === null || value === undefined) {
        if (!Object.hasOwn(all, key)) return false;
      } else {
        const json = JSON.stringify(value);
        if (typeof json !== 'string' || Buffer.byteLength(json) > OFFICE_STATE_VALUE_MAX) throw new Error('설정 값이 너무 큽니다 (8KB 넘음)');
        if (Object.hasOwn(all, key) && JSON.stringify(all[key]) === json) return false;
        entries.push([key, JSON.parse(json)]);
      }
      data = { ...data, officeEditor: { ...data.officeEditor, settings: Object.fromEntries(entries.slice(-OFFICE_STATE_MAX_KEYS)) } };
      save();
      return true;
    },
  };
}

// ── 백업 ─────────────────────────────────────────────────

const pad = (n) => String(n).padStart(2, '0');
export function stamp(d = new Date()) {
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

export function hashFile(path) { return createHash('sha1').update(readFileSync(path)).digest('hex'); }

/**
 * 원본 백업 = 문서 버전. dir = userData/backups
 * 파일 이름은 <yyyyMMdd-HHmmss>-<원래이름>, 원래 경로·까닭(reason)·내용 해시는 index.json에 적어 둔다.
 * keep: 모두 합쳐 몇 개까지, maxBytes: 모두 합쳐 몇 바이트까지 (방금 것·protect는 넘어도 남긴다)
 */
export function createBackups(dir, { keep = BACKUP_KEEP, maxBytes = BACKUP_MAX_BYTES } = {}) {
  mkdirSync(dir, { recursive: true });
  const indexFile = join(dir, 'index.json');
  const load = () => readJson(indexFile, []).filter((e) => e && e.file && e.source && existsSync(join(dir, e.file)));
  const store = (list) => writeJsonAtomic(indexFile, list);

  // 마지막으로 쓴 시각 (같은 내용이라 다시 쓴 백업은 usedAt이 새로워진다)
  const lastUse = (e) => Math.max(e.time ?? 0, e.usedAt ?? 0);
  // 최근 것부터. 같은 시각이면 나중에 추가한 것이 앞 (정렬은 안정적이므로 뒤집은 뒤 정렬)
  const forSource = (list, source) => list.filter((e) => samePath(e.source, source)).reverse().sort((a, b) => lastUse(b) - lastUse(a));
  const nameOf = (x) => (typeof x === 'string' ? x : x?.file);

  return {
    dir,
    /**
     * 원본 복사. 같은 내용의 백업이 이미 있으면 새로 만들지 않고 그것을 돌려준다 → { file, path, time, source, reused }
     * protect: 정리할 때 지우면 안 되는 백업(파일 이름 또는 항목) — 열린 문서의 '원래대로' 기준 등
     * maxPerSource: 이 문서의 백업은 최근 n개만 (protect·방금 것은 남기고 개수에 센다). 0이면 따로 세지 않음
     * reason: 왜 만드는지 ('open' 문서를 열 때 | 'editor' 편집 창 저장 전 | 'ai' AI 대화 전 | 'restore' 원래대로 전).
     *   같은 내용의 백업을 다시 쓰면 처음 까닭을 그대로 둔다 (그 내용이 처음 생긴 때를 말하므로)
     */
    backup(source, now = new Date(), { protect = [], maxPerSource = 0, reason = null } = {}) {
      source = resolve(source);
      const list = load();
      const hash = hashFile(source);
      let entry = forSource(list, source).find((e) => e.hash === hash);
      const reused = !!entry;
      if (entry) {
        // 다시 쓴 백업은 가장 최근 것으로 쳐서 정리 순서에서 뒤로 보낸다
        entry.usedAt = now.getTime();
      } else {
        let file = `${stamp(now)}-${basename(source)}`;
        for (let i = 2; existsSync(join(dir, file)); i++) file = `${stamp(now)}-${i}-${basename(source)}`;
        copyFileSync(source, join(dir, file));
        entry = { file, source, time: now.getTime(), hash, size: statSync(source).size, ...(BACKUP_REASONS.includes(reason) ? { reason } : {}) };
        list.push(entry);
      }
      // 오래 안 쓴 것부터 지워 keep개만 남긴다. 방금 것과 protect는 남긴다
      const keepSet = new Set([entry.file, ...[].concat(protect).map(nameOf).filter(Boolean)]);
      // 이 문서 몫: 자주 저장해도 다른 문서의 백업을 밀어내지 않게 최근 maxPerSource개만
      if (maxPerSource > 0) {
        let kept = 0;
        for (const e of forSource(list, source)) {
          if (keepSet.has(e.file) || kept < maxPerSource) { kept++; continue; }
          try { rmSync(join(dir, e.file), { force: true }); } catch { continue; }
          list.splice(list.indexOf(e), 1);
        }
      }
      list.sort((a, b) => lastUse(a) - lastUse(b));
      for (let i = 0; list.length > keep && i < list.length;) {
        if (keepSet.has(list[i].file)) { i++; continue; }
        // 다른 프로그램이 잡고 있어 못 지우면 목록에 남겨 두고 다음에 다시 지운다
        try { rmSync(join(dir, list[i].file), { force: true }); } catch { i++; continue; }
        list.splice(i, 1);
      }
      // 바이트 한도: 큰 문서를 자주 저장해도 디스크를 다 쓰지 않게 (오래 안 쓴 것부터)
      let bytes = list.reduce((n, e) => n + (e.size ?? 0), 0);
      for (let i = 0; bytes > maxBytes && i < list.length;) {
        if (keepSet.has(list[i].file)) { i++; continue; }
        try { rmSync(join(dir, list[i].file), { force: true }); } catch { i++; continue; }
        bytes -= list[i].size ?? 0;
        list.splice(i, 1);
      }
      store(list);
      return { ...entry, path: join(dir, entry.file), reused };
    },
    /** 이 문서의 가장 최근 백업 */
    latest(source) {
      const e = forSource(load(), resolve(source))[0];
      return e ? { ...e, path: join(dir, e.file) } : null;
    },
    /** 이 문서의 백업(버전)들, 최근에 쓴 것부터 */
    list(source) { return forSource(load(), resolve(source)).map((e) => ({ ...e, path: join(dir, e.file) })); },
    /** 파일 이름으로 이 문서의 백업 하나 (없으면 null) */
    find(source, file) {
      const e = forSource(load(), resolve(source)).find((x) => x.file === file);
      return e ? { ...e, path: join(dir, e.file) } : null;
    },
    /** 백업을 원래 자리에 덮어쓴다. entry 생략 시 가장 최근 것 */
    restore(source, entry) {
      source = resolve(source);
      const e = entry ?? this.latest(source);
      if (!e) throw new Error('이 문서의 백업이 없습니다');
      const from = join(dir, e.file);
      if (!existsSync(from)) throw new Error('백업 파일을 찾을 수 없습니다');
      copyFileSync(from, source);
      // 복사는 백업의 옛 수정 시각을 그대로 옮긴다 → 지금 바뀐 파일로 보이게 시각을 새로 찍는다 (동기화·백업 프로그램용)
      try { const now = new Date(); utimesSync(source, now, now); } catch {}
      return { ...e, path: from };
    },
  };
}

// ── 미리보기 PDF 기억 ────────────────────────────────────

export const PREVIEW_CACHE_KEEP = 40;
export const PREVIEW_CACHE_MAX_BYTES = 512 * 1024 * 1024;
// 만드는 방법이 바뀌면 올린다 → 옛 PDF는 쓰지 않는다 (2: Excel 시트마다 한 쪽 너비)
export const PREVIEW_CACHE_VERSION = 2;

/**
 * 프로그램(Word·Excel·PowerPoint·한컴)으로 만든 미리보기 PDF를 기억한다. dir = userData/preview-cache
 * 열쇠 = 문서 경로·크기·수정 시각·미리보기 방식(엔진) → 파일이 그대로면 다시 변환하지 않는다 (5~25초 걸리는 변환).
 * 경로는 해시로만 적는다. 최근에 쓴 것부터 keep개·maxBytes까지.
 */
export function createPreviewCache(dir, { keep = PREVIEW_CACHE_KEEP, maxBytes = PREVIEW_CACHE_MAX_BYTES } = {}) {
  const indexFile = join(dir, 'index.json');
  const load = () => readJson(indexFile, []).filter((e) => e && typeof e.key === 'string' && existsSync(join(dir, `${e.key}.pdf`)));
  const keyOf = ({ path, size, mtimeMs, mode }) => createHash('sha1')
    .update(`${PREVIEW_CACHE_VERSION}|${process.platform === 'win32' ? resolve(path).toLowerCase() : resolve(path)}|${size}|${mtimeMs}|${mode}`).digest('hex');
  return {
    dir,
    /** → { file(기억해 둔 PDF), engine, pages } | null. 쓴 시각을 새로 적는다 */
    get(what) {
      const k = keyOf(what);
      const list = load();
      const e = list.find((x) => x.key === k);
      if (!e) return null;
      e.usedAt = Date.now();
      try { writeJsonAtomic(indexFile, list); } catch {}
      return { file: join(dir, `${k}.pdf`), engine: e.engine ?? null, pages: e.pages ?? null };
    },
    /** PDF를 복사해 기억한다. 한도를 넘으면 오래 안 쓴 것부터 지운다 */
    put(what, file, { engine = null, pages = null } = {}) {
      mkdirSync(dir, { recursive: true });
      const k = keyOf(what);
      copyFileSync(file, join(dir, `${k}.pdf`));
      const list = load().filter((x) => x.key !== k);
      list.push({ key: k, engine, pages, size: statSync(join(dir, `${k}.pdf`)).size, usedAt: Date.now() });
      list.sort((a, b) => (a.usedAt ?? 0) - (b.usedAt ?? 0));
      let bytes = list.reduce((n, e) => n + (e.size ?? 0), 0);
      while (list.length > 1 && (list.length > keep || bytes > maxBytes)) {
        const old = list[0];
        try { rmSync(join(dir, `${old.key}.pdf`), { force: true }); } catch { break; }
        bytes -= old.size ?? 0;
        list.shift();
      }
      writeJsonAtomic(indexFile, list);
    },
    clear() { rmSync(dir, { recursive: true, force: true }); },
  };
}
