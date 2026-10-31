#!/usr/bin/env node
// merge.mjs — 번역 조각(parts/en-1.json …)과 손본 글(fixes.<lang>.json)을 사전(en.json·ja.json)으로 합친다.
//   strings.json에 있는 글만 넣고 (소스에서 사라진 글은 뺀다), 다음은 사전에 넣지 않는다:
//   - 조사·읽는 소리 같은 한두 글자 조각 (com.mjs의 조사 고르기 표 — 화면 글이 아니고, 자리 글을 잘못 바꿀 수 있다)
//   - 코드·토큰(정규식, XML, {쪽} 같은 머리말 자리) — 화면에 나와도 사용자가 그대로 써야 하는 글
//   사용: node scripts/i18n/merge.mjs
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { skipKey } from './rules.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const DIR = join(ROOT, 'src', 'app', 'i18n');
const strings = JSON.parse(readFileSync(join(DIR, 'strings.json'), 'utf8'));
const HAN = /[가-힣]/;

const placeholders = (s) => (s.match(/\{\d+\}/g) ?? []).sort().join(',');
const tokens = (s) => (s.match(/\{(?!\d+\})[^{}]+\}/g) ?? []).sort().join(',');

let bad = 0;
for (const lang of ['en', 'ja']) {
  const parts = readdirSync(join(DIR, 'parts')).filter((f) => new RegExp(`^${lang}-\\d+\\.json$`).test(f)).sort();
  const all = {};
  for (const f of parts) Object.assign(all, JSON.parse(readFileSync(join(DIR, 'parts', f), 'utf8')));
  const fixFile = join(DIR, 'parts', `fixes.${lang}.json`);
  const fixes = existsSync(fixFile) ? JSON.parse(readFileSync(fixFile, 'utf8')) : {};
  Object.assign(all, fixes);
  const out = {};
  const keys = [...new Set([...strings, ...Object.keys(fixes)])].sort((a, b) => a.localeCompare(b, 'ko'));
  for (const k of keys) {
    const v = all[k];
    if (typeof v !== 'string') { if (strings.includes(k)) { console.log(`${lang}: 번역 없음 ${JSON.stringify(k)}`); bad++; } continue; }
    if (skipKey(k)) continue;
    if (placeholders(k) !== placeholders(v)) { console.log(`${lang}: 자리 표시가 다름 ${JSON.stringify(k)} → ${JSON.stringify(v)}`); bad++; continue; }
    if (tokens(k) !== tokens(v)) { console.log(`${lang}: {토큰}이 다름 ${JSON.stringify(k)} → ${JSON.stringify(v)}`); bad++; continue; }
    if (lang === 'ja' && HAN.test(v)) { console.log(`ja: 한글이 남음 ${JSON.stringify(k)} → ${JSON.stringify(v)}`); }
    if (!v.trim()) continue;
    out[k] = v;
  }
  writeFileSync(join(DIR, `${lang}.json`), `${JSON.stringify(out, null, 1)}\n`, 'utf8');
  console.log(`${lang}: ${Object.keys(out).length}개`);
}
process.exit(bad ? 1 : 0);
