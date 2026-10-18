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
  theme: 'system',         // 'system' | 'light' | 'dark'
  chatWidth: 400,
  chatCollapsed: false,
  window: null,            // { x, y, width, height, maximized }
  recent: [],              // [{ path, openedAt }]
  officeEditor: { settings: {} }, // 문서·표·슬라이드 편집기가 보관하는 설정 (편집기의 state.get/set, 값마다 8KB)
});

const RECENT_MAX = 10;
/** Office 편집기 설정: 이름 규칙·값 크기(JSON 바이트)·개수 한도 */
export const OFFICE_STATE_KEY = /^[\w.:-]{1,128}$/;
export const OFFICE_STATE_VALUE_MAX = 8 * 1024;
export const OFFICE_STATE_MAX_KEYS = 128;
const BACKUP_KEEP = 50;

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
    recent() {
      return data.recent.map((r) => ({ ...r, name: basename(r.path), exists: existsSync(r.path) }));
    },
    addRecent(path) {
      data.recent = [{ path, openedAt: Date.now() }, ...data.recent.filter((r) => !samePath(r.path, path))].slice(0, RECENT_MAX);
      save();
    },
    removeRecent(path) { data.recent = data.recent.filter((r) => !samePath(r.path, path)); save(); },
    clearRecent() { data.recent = []; save(); },

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

function hashFile(path) { return createHash('sha1').update(readFileSync(path)).digest('hex'); }

/**
 * 원본 백업. dir = userData/backups
 * 파일 이름은 <yyyyMMdd-HHmmss>-<원래이름>, 원래 경로는 index.json에 적어 둔다.
 */
export function createBackups(dir, { keep = BACKUP_KEEP } = {}) {
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
     */
    backup(source, now = new Date(), { protect = [], maxPerSource = 0 } = {}) {
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
        entry = { file, source, time: now.getTime(), hash, size: statSync(source).size };
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
      store(list);
      return { ...entry, path: join(dir, entry.file), reused };
    },
    /** 이 문서의 가장 최근 백업 */
    latest(source) {
      const e = forSource(load(), resolve(source))[0];
      return e ? { ...e, path: join(dir, e.file) } : null;
    },
    list(source) { return forSource(load(), resolve(source)).map((e) => ({ ...e, path: join(dir, e.file) })); },
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
