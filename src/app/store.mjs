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
});

const RECENT_MAX = 10;
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
  return data;
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
     */
    backup(source, now = new Date(), { protect = [] } = {}) {
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
