// batch.mjs — 한글(.hwp) 문서를 HWPX로 한꺼번에 바꾸고, 바꾼 파일을 다시 열어 원본과 내용이 같은지 확인한다.
//   변환은 convertDocument (한컴오피스+자동화 연결이 있으면 한컴, 없으면 내장 엔진 rhwp). 원본은 건드리지 않는다.
//   확인: 쪽 수·표 수·글(공백 무시)을 원본과 견준다 → 'ok' 같음 | 'check' 다른 곳 있음(직접 확인 필요) | 'skipped' | 'fail'
import { existsSync, mkdirSync, readdirSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, join, relative, resolve } from 'node:path';
import { randomBytes } from 'node:crypto';
import { convertDocument, Hwp } from './index.mjs';
import { diffUnits } from './diff.mjs';

export const BATCH_SOURCE_EXTS = ['hwp'];
export const BATCH_MAX_FILES = 5000;
const SKIP_DIRS = new Set(['$recycle.bin', 'system volume information', 'node_modules', '.git']);

const isSource = (name) => BATCH_SOURCE_EXTS.includes(extname(name).slice(1).toLowerCase()) && !name.startsWith('~$') && !name.startsWith('.~');
const keyOf = (p) => (process.platform === 'win32' ? p.toLowerCase() : p);

/**
 * 파일·폴더 목록에서 바꿀 한글 파일 찾기 → { items: [{ src, root }], truncated }
 * root: 폴더로 고른 경우 그 폴더 (다른 폴더에 모을 때 하위 폴더 구조를 그대로 두려고), 파일로 고르면 그 파일의 폴더
 */
export function collectHwpFiles(inputs, { recursive = true, limit = BATCH_MAX_FILES } = {}) {
  const items = []; const seen = new Set();
  let truncated = false;
  const add = (src, root) => {
    const k = keyOf(src);
    if (seen.has(k)) return;
    if (items.length >= limit) { truncated = true; return; }
    seen.add(k); items.push({ src, root });
  };
  const walk = (dir, root, depth) => {
    let entries;
    try { entries = readdirSync(dir, { withFileTypes: true }); } catch { return; }
    entries.sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    for (const e of entries) {
      if (truncated) return;
      const p = join(dir, e.name);
      if (e.isFile() && isSource(e.name)) add(p, root);
      else if (e.isDirectory() && recursive && depth < 32 && !SKIP_DIRS.has(e.name.toLowerCase()) && !e.name.startsWith('.')) walk(p, root, depth + 1);
    }
  };
  for (const raw of [].concat(inputs ?? [])) {
    if (typeof raw !== 'string' || !raw.trim()) continue;
    const p = resolve(raw);
    let st;
    try { st = statSync(p); } catch { continue; }
    if (st.isDirectory()) walk(p, p, 0);
    else if (st.isFile() && isSource(basename(p))) add(p, dirname(p));
  }
  return { items, truncated };
}

/** 결과 경로: 원본 옆(같은 이름 .hwpx) 또는 outDir 아래 (고른 폴더 기준 하위 경로 유지) */
export function outputFor({ src, root }, outDir) {
  const name = `${basename(src, extname(src))}.hwpx`;
  if (!outDir) return join(dirname(src), name);
  const rel = relative(root, dirname(src));
  return join(resolve(outDir), rel.startsWith('..') ? '' : rel, name);
}

const norm = (t) => String(t ?? '').replace(/\s+/g, ' ').trim();

function snapshot(path) {
  const h = Hwp.open(path);
  try { return { pages: h.pageCount, tables: h.tables().length, text: h.text() }; } finally { h.close(); }
}

/** 원본과 바꾼 파일 비교 → { same, pages: [전, 후], tables: [전, 후], textDiff: 다른 곳 수, notes: [한국어] } */
export function verifyConversion(src, out) {
  const a = snapshot(src);
  let b;
  try { b = snapshot(out); } catch (err) { return { same: false, reopen: false, notes: [`바꾼 파일을 다시 열지 못했어요: ${err?.message ?? err}`] }; }
  const notes = [];
  let textDiff = 0;
  if (norm(a.text) !== norm(b.text)) {
    const lines = (t) => String(t).split('\n').map((x) => ({ text: x.replace(/\s+/g, ' ').trim() }));
    const d = diffUnits(lines(a.text), lines(b.text));
    textDiff = d.stats.added + d.stats.removed + d.stats.changed;
    // 공백·빈 줄만 다르면 diffUnits가 0곳으로 본다
    if (textDiff) notes.push(`글이 다른 곳 ${textDiff}곳`);
  }
  if (a.pages !== b.pages) notes.push(`쪽 수 ${a.pages} → ${b.pages}`);
  if (a.tables !== b.tables) notes.push(`표 ${a.tables}개 → ${b.tables}개`);
  return { same: !notes.length, reopen: true, pages: [a.pages, b.pages], tables: [a.tables, b.tables], textDiff, notes };
}

/**
 * 한꺼번에 바꾸기. items: collectHwpFiles().items
 * opts: outDir(없으면 원본 옆), overwrite(같은 이름 .hwpx가 있으면 덮어쓰기), onProgress({ index, total, result }), isCancelled()
 * → { results: [{ src, out, status, engine?, notes, pages?, tables?, error? }], summary: { ok, check, skipped, fail, total, cancelled } }
 */
export async function convertHwpBatch(items, { outDir, overwrite = false, onProgress, isCancelled, convert = convertDocument } = {}) {
  const results = [];
  let cancelled = false;
  for (let i = 0; i < items.length; i++) {
    if (isCancelled?.()) { cancelled = true; break; }
    const { src } = items[i];
    const out = outputFor(items[i], outDir);
    let r;
    if (existsSync(out) && !overwrite) {
      r = { src, out, status: 'skipped', notes: ['같은 이름의 HWPX 파일이 이미 있어 건너뛰었어요'] };
    } else {
      // 같은 폴더의 임시 이름으로 만든 뒤 확인하고 바꿔 끼운다 (중간에 끊겨도 있던 파일이 반쯤 망가지지 않게)
      const tmp = join(dirname(out), `~officekit-${randomBytes(4).toString('hex')}.hwpx`);
      try {
        mkdirSync(dirname(out), { recursive: true });
        const c = await convert(src, tmp);
        const v = verifyConversion(src, tmp);
        if (!v.reopen) throw new Error(v.notes[0]);
        renameSync(tmp, out);
        r = { src, out, status: v.same ? 'ok' : 'check', engine: c?.engine ?? null, pages: v.pages, tables: v.tables, notes: v.notes };
      } catch (err) {
        try { rmSync(tmp, { force: true }); } catch {}
        const msg = String(err?.message ?? err);
        const note = /EBUSY|EPERM|EACCES/.test(err?.code ?? msg) ? '저장하지 못했어요 — 같은 이름의 파일이 열려 있거나 쓰기 권한이 없어요'
          : /UNSUPPORTED_FILE_FORMAT|유효하지 않은 파일|지원하지 않는 포맷/i.test(msg) ? '한글 문서가 아니거나 손상된 파일이에요 (또는 지원하지 않는 옛 한글 형식)'
            : /암호|password|encrypt|배포용|distribut/i.test(msg) ? '암호가 걸렸거나 배포용 문서라 바꿀 수 없어요'
              : msg;
        r = { src, out, status: 'fail', notes: [note], error: msg };
      }
    }
    results.push(r);
    try { onProgress?.({ index: i, total: items.length, result: r }); } catch {}
  }
  const count = (s) => results.filter((r) => r.status === s).length;
  return { results, summary: { ok: count('ok'), check: count('check'), skipped: count('skipped'), fail: count('fail'), total: items.length, cancelled } };
}

const STATUS_KO = { ok: '같음', check: '확인 필요', skipped: '건너뜀', fail: '실패' };
const ENGINE_KO = { rhwp: '내장 엔진', hwp: '한컴오피스' };

/** 결과 보고서 (CSV, 엑셀에서 바로 열리게 UTF-8 BOM) */
export function batchReportCsv(results) {
  const q = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const head = ['결과', '원본', 'HWPX', '쪽 수(원본)', '쪽 수(HWPX)', '표(원본)', '표(HWPX)', '변환', '메모'];
  const rows = results.map((r) => [STATUS_KO[r.status] ?? r.status, r.src, r.out, r.pages?.[0] ?? '', r.pages?.[1] ?? '', r.tables?.[0] ?? '', r.tables?.[1] ?? '', ENGINE_KO[r.engine] ?? r.engine ?? '', (r.notes ?? []).join(' / ')]);
  return `﻿${[head, ...rows].map((row) => row.map(q).join(',')).join('\r\n')}\r\n`;
}

export function writeBatchReport(path, results) {
  writeFileSync(path, batchReportCsv(results), 'utf8');
  return path;
}

export { STATUS_KO as BATCH_STATUS_KO };
