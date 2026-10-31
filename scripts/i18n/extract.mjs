#!/usr/bin/env node
// extract.mjs — 화면에 나올 수 있는 한국어 글을 소스에서 뽑아 src/app/i18n/strings.json 으로 (번역할 목록).
//   JS: acorn으로 글자·템플릿 글자 (템플릿의 ${…} 자리는 {0} {1} … 으로). 주석·console.log·log() 안 글은 뺀다.
//   HTML: 글 노드와 title·placeholder·aria-label 속성.
//   사용: node scripts/i18n/extract.mjs [--check]  (--check: 사전(en·ja)에 빠진 글이 있으면 실패)
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as acorn from 'acorn';
import { skipKey } from './rules.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const OUT = join(ROOT, 'src', 'app', 'i18n', 'strings.json');

/** 번역할 파일 (앱 화면에 글이 나오는 곳) */
export const SOURCES = [
  'src/app/renderer/index.html', 'src/app/renderer/app.js',
  'src/app/main.mjs', 'src/app/editor-window.mjs', 'src/app/store.mjs',
  'src/app/editor/host.html', 'src/app/editor/host.js',
  'src/app/office-editor/host.html', 'src/app/office-editor/host.js',
  'src/setup/fileassoc.mjs', 'src/setup/integrations.mjs', 'src/setup/hancom.mjs',
  'src/core/index.mjs', 'src/core/create.mjs', 'src/core/templates.mjs', 'src/core/pii.mjs', 'src/core/batch.mjs',
  'src/core/diff.mjs', 'src/core/pdffile.mjs', 'src/core/pdftools.mjs', 'src/core/hwppdf.mjs', 'src/core/com.mjs',
  'src/core/officefile.mjs', 'src/core/hwp.mjs', 'src/core/render.mjs',
  'src/app/ai.mjs',
];

const HAN = /[가-힣]/;
// 화면에 나오지 않는 호출 (진단 로그)
const SKIP_CALLEES = new Set(['log', 'debug', 'warn', 'error', 'info']);

function calleeName(node) {
  const c = node?.callee;
  if (!c) return '';
  if (c.type === 'Identifier') return c.name;
  if (c.type === 'MemberExpression' && !c.computed) return c.property.name;
  return '';
}

/** 템플릿 글 → '{0}' 자리 표시 키 */
function templateKey(node) {
  let out = '';
  node.quasis.forEach((q, i) => { out += q.value.cooked ?? q.value.raw; if (i < node.expressions.length) out += `{${i}}`; });
  return out;
}

function walk(node, parents, onString) {
  if (!node || typeof node.type !== 'string') return;
  const inLog = parents.some((p) => p.type === 'CallExpression' && SKIP_CALLEES.has(calleeName(p))
    && (p.callee.type === 'Identifier' ? p.callee.name === 'log' : ['console'].includes(p.callee.object?.name) || p.callee.object?.name === 'log'));
  if (!inLog) {
    if (node.type === 'Literal' && typeof node.value === 'string' && HAN.test(node.value)) onString(node.value);
    if (node.type === 'TemplateLiteral') {
      const key = templateKey(node);
      if (HAN.test(key)) onString(key);
    }
  }
  for (const k of Object.keys(node)) {
    if (k === 'loc' || k === 'start' || k === 'end') continue;
    const v = node[k];
    if (Array.isArray(v)) for (const x of v) walk(x, [...parents, node], onString);
    else if (v && typeof v.type === 'string') walk(v, [...parents, node], onString);
  }
}

function fromJs(src, onString) {
  const ast = acorn.parse(src, { ecmaVersion: 'latest', sourceType: 'module', allowHashBang: true, allowReturnOutsideFunction: true });
  walk(ast, [], onString);
}

// translate.js의 richKey와 같은 규칙
const richKey = (html) => html.replace(/<(\/?)b\b[^>]*>/gi, '<$1b>').replace(/\s+/g, ' ').trim();

function fromHtml(src, onString) {
  let body = src.replace(/<!--[\s\S]*?-->/g, '').replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '');
  // <b>가 섞인 문장 (data-i18n-rich): 안쪽 HTML 하나로, 조각은 뽑지 않는다
  body = body.replace(/(<(\w+)[^>]*\sdata-i18n-rich[^>]*>)([\s\S]*?)(<\/\2>)/g, (_m, open, _tag, inner, close) => { onString(richKey(inner)); return `${open}${close}`; });
  for (const m of body.matchAll(/>([^<>]+)</g)) {
    const t = m[1].replace(/\s+/g, ' ').trim();
    if (HAN.test(t)) onString(t);
  }
  for (const m of body.matchAll(/\s(?:title|placeholder|aria-label|alt)="([^"]*)"/g)) if (HAN.test(m[1])) onString(m[1].trim());
}

const norm = (s) => s.replace(/\r\n/g, '\n');

export function extractAll() {
  const all = new Map(); // 글 → [파일]
  for (const rel of SOURCES) {
    const file = join(ROOT, rel);
    const src = readFileSync(file, 'utf8');
    const add = (s) => {
      const k = norm(s);
      if (!k.trim()) return;
      if (!all.has(k)) all.set(k, new Set());
      all.get(k).add(rel);
    };
    try {
      if (rel.endsWith('.html')) fromHtml(src, add); else fromJs(src, add);
    } catch (e) { throw new Error(`${rel}: ${e.message}`); }
  }
  return all;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const all = extractAll();
  const list = [...all.keys()].sort((a, b) => a.localeCompare(b, 'ko'));
  if (process.argv.includes('--check')) {
    let missing = 0;
    for (const lang of ['en', 'ja']) {
      const dict = JSON.parse(readFileSync(join(ROOT, 'src', 'app', 'i18n', `${lang}.json`), 'utf8'));
      const miss = list.filter((k) => !Object.hasOwn(dict, k) && !skipKey(k));
      if (miss.length) { missing += miss.length; console.log(`${lang}: 빠진 글 ${miss.length}개 (예: ${JSON.stringify(miss.slice(0, 3))})`); }
    }
    process.exit(missing ? 1 : 0);
  }
  writeFileSync(OUT, `${JSON.stringify(list, null, 1)}\n`, 'utf8');
  console.log(`${list.length}개 → ${relative(ROOT, OUT)}`);
}
