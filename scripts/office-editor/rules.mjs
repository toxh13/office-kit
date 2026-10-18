// rules.mjs — Office 편집기 묶음 덜어 내기 규칙 (docs/OFFICE-EDITOR-DESIGN.md 4장 '덜어 내기').
//
// 경로는 ui.zip 안 경로 그대로 ('editor/web-apps/...', 'x2t/x2t.wasm', 'index.html').
// 파일을 통째로 빼기만 한다 — 남는 파일은 한 바이트도 고치지 않는다 (줄이기·압축 다시 하기 없음).
// 규칙을 바꾸면 NOTICE_DATED도 바꾼다 (NOTICE-OfficeKit.md의 수정 날짜).
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

/** 덜어 내기 규칙을 마지막으로 바꾼 날 (수정 고지의 날짜) */
export const NOTICE_DATED = '2026-10-18';

const KEEP_LANGS = ['ko', 'en'];
const APPS = ['documenteditor', 'spreadsheeteditor', 'presentationeditor'];

/** 빼는 것: { id, what(한국어), why, test(path) } — 위에서부터 처음 맞는 규칙 하나로 센다 */
export const DROP_RULES = Object.freeze([
  {
    id: 'locales',
    what: '편집기 언어 파일 (ko·en 말고)',
    why: '한국어 화면 + 영어 대비',
    test: (p) => {
      const m = /^editor\/web-apps\/apps\/.+\/locale\/([^/]+)\.json$/.exec(p);
      return !!m && !KEEP_LANGS.includes(m[1]);
    },
  },
  {
    id: 'formula-lang',
    what: '수식 언어 파일 (ko·en 말고)',
    why: '한국어 화면 + 영어 대비',
    test: (p) => {
      const m = /^editor\/web-apps\/apps\/[^/]+\/main\/resources\/formula-lang\/([^/]+)$/.exec(p);
      return !!m && !/^(ko|en)(_desc)?\.json$/.test(m[1]);
    },
  },
  {
    id: 'mobile-embed-forms',
    what: '휴대폰·embed·forms 앱 (web-apps/apps/*/mobile|embed|forms)',
    why: '데스크톱 편집 창만 쓴다',
    test: (p) => /^editor\/web-apps\/apps\/[^/]+\/(mobile|embed|forms)\//.test(p),
  },
  {
    id: 'fonts',
    what: '묶음 글꼴 editor/fonts와 그 라이선스 (editor/core-fonts-licenses)',
    why: '실행 중에 이 PC에 설치된 글꼴을 쓴다 (묶음에는 한글 글꼴도 없다). 설치 파일에 편집용 글꼴 파일은 넣지 않는다 — 다만 x2t.wasm 안에 ONLYOFFICE core가 넣어 둔 PDF 기본 14 글꼴(URW++ Nimbus Sans·Roman·Mono, Standard Symbols L, Dingbats — AGPL-3.0 + PDF 포함 예외)은 들어 있다',
    test: (p) => /^editor\/(fonts|core-fonts-licenses)\//.test(p),
  },
]);

/** 어떤 규칙도 빼면 안 되는 것 (라이선스·고지) */
export const PROTECTED = Object.freeze([
  'LICENSE',
  'NOTICE',
  'editor/LICENSE.txt',
  'editor/3rd-Party.txt',
  'editor/filex/CHANGES.txt',
]);
export const PROTECTED_DIRS = Object.freeze(['editor/license/']);

/** 덜어 낸 뒤 꼭 있어야 하는 것 */
export const REQUIRED = Object.freeze([
  ...PROTECTED,
  'index.html',
  'filex/app.js',
  'filex/app.css',
  'filex/x2t-worker.js',
  'x2t/x2t.js',
  'x2t/x2t.wasm',
  'editor/themes.json',
  'editor/filex/storage.js',
  'editor/web-apps/apps/api/documents/api.js',
  'editor/web-apps/vendor/socketio/socket.io.min.js',
  'editor/sdkjs/common/AllFonts.js',
  ...APPS.flatMap((a) => [
    `editor/web-apps/apps/${a}/main/index.html`,
    ...KEEP_LANGS.map((l) => `editor/web-apps/apps/${a}/main/locale/${l}.json`),
  ]),
  ...KEEP_LANGS.flatMap((l) => [
    `editor/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/${l}.json`,
    `editor/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/${l}_desc.json`,
  ]),
]);

/** 이 스크립트가 묶음에 더하는 파일 (묶음 안 파일과 이름이 겹치면 안 됨) */
export const GENERATED = Object.freeze(['VERSION.json', 'NOTICE-OfficeKit.md', 'THIRD-PARTY-NATIVE.txt']);

/** 실행 중에 Office Kit이 다른 내용으로 내주는 파일 (파일은 그대로 두고 서버가 바꿔 내줌 — 5장 글꼴) */
export const RUNTIME_REPLACED = Object.freeze([
  { path: 'editor/sdkjs/common/AllFonts.js', what: '이 PC에 설치된 글꼴로 만든 글꼴 목록' },
  { path: 'editor/sdkjs/common/Images/fonts_thumbnail*.png(.bin)', what: '그 글꼴로 다시 그린 글꼴 미리보기 그림' },
  { path: 'editor/fonts/<번호>', what: '이 PC의 글꼴 파일 (복사하지 않고 읽기만, 앞 32바이트를 ONLYOFFICE 방식으로 XOR해서 보냄)' },
]);

const isProtected = (p) => PROTECTED.includes(p) || PROTECTED_DIRS.some((d) => p.startsWith(d));

/** 이 경로를 빼는 규칙 (없으면 null). 보호 파일에 규칙이 맞으면 오류 */
export function dropRuleFor(p) {
  const rule = DROP_RULES.find((r) => r.test(p));
  if (rule && isProtected(p)) throw new Error(`덜어 내기 규칙 '${rule.id}'이(가) 지우면 안 되는 파일(${p})에 맞습니다 — rules.mjs를 고치세요.`);
  return rule || null;
}

/**
 * paths(묶음 안 경로 목록)를 남길 것과 뺄 것으로 나눈다.
 * @returns {{ kept: string[], removed: Map<string, string[]> }}  removed: 규칙 id → 경로들
 */
export function partition(paths) {
  const kept = [];
  const removed = new Map(DROP_RULES.map((r) => [r.id, []]));
  for (const p of paths) {
    const rule = dropRuleFor(p);
    if (rule) removed.get(rule.id).push(p);
    else kept.push(p);
  }
  const keptSet = new Set(kept);
  const missing = REQUIRED.filter((p) => !keptSet.has(p));
  if (!kept.some((p) => p.startsWith('editor/license/'))) missing.push('editor/license/*');
  if (missing.length) throw new Error(`덜어 낸 묶음에 꼭 있어야 할 파일이 없습니다: ${missing.join(', ')}`);
  const clash = GENERATED.filter((g) => keptSet.has(g));
  if (clash.length) throw new Error(`묶음에 이미 ${clash.join(', ')}이(가) 있습니다 — 이 스크립트가 만드는 파일과 이름이 겹칩니다.`);
  return { kept, removed };
}

/** 규칙 해시: 이 파일 내용 (줄바꿈 LF로 맞춤 — Windows·CI 체크아웃이 달라도 같게) */
export function rulesSha256() {
  const text = readFileSync(fileURLToPath(import.meta.url), 'utf8').replace(/\r\n/g, '\n');
  return createHash('sha256').update(text).digest('hex');
}
