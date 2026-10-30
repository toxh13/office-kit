#!/usr/bin/env node
// server.mjs — office-kit MCP 서버 (stdio). Claude/Codex가 문서 도구를 쓰게 한다.
// stdout은 프로토콜 전용이므로 로그는 모두 stderr로 보낸다.
import { existsSync, mkdirSync, readFileSync, realpathSync } from 'node:fs';
import { basename, delimiter, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import {
  readDocument, documentInfo, listFields, fillFields, replaceText, setCells, convertDocument,
  createDocument, probe, Hwp, extOf, kindOf, isHwp, OPENABLE_EXTS, hwpPdfStatus,
} from '../core/index.mjs';
import { PLACEHOLDER } from '../core/officefile.mjs';
import { pageImage } from '../core/render.mjs';
import { PII_TYPES, scanDocument, maskDocument, maskedPathFor } from '../core/pii.mjs';
import { collectHwpFiles, convertHwpBatch, outputFor, writeBatchReport, BATCH_STATUS_KO } from '../core/batch.mjs';
import { register as registerHwp } from './tools/hwp.mjs';
import { register as registerWord } from './tools/word.mjs';
import { register as registerPpt } from './tools/ppt.mjs';
import { register as registerExcel } from './tools/excel.mjs';

const TOOL_MODULES = [registerHwp, registerWord, registerPpt, registerExcel];

const here = dirname(fileURLToPath(import.meta.url));
const VERSION = (() => {
  try { return JSON.parse(readFileSync(resolve(here, '../../package.json'), 'utf8')).version; } catch { return '0.0.0'; }
})();

const log = (...a) => console.error('[office-kit]', ...a);

// ── 공통 ─────────────────────────────────────────────────

/** 경로: 절대 경로 또는 서버 작업 폴더(process.cwd()) 기준 상대 경로 */
const abs = (p) => resolve(process.cwd(), String(p).trim().replace(/^"(.*)"$/, '$1'));

function mustExist(p) {
  const full = abs(p);
  if (!existsSync(full)) throw new Error(`파일이 없습니다: ${full}`);
  return full;
}

const samePath = (a, b) => (process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b);

// 저장할 수 있는 형식: 문서 형식만 (실행 파일·스크립트 확장자로는 쓰지 않는다)
const EDIT_EXTS = new Set(['hwp', 'hwpx', 'hwt', 'docx', 'doc', 'rtf', 'odt', 'pptx', 'ppt', 'odp', 'xlsx', 'xls', 'xlsm', 'csv', 'ods']);
const CONVERT_EXTS = new Set([...EDIT_EXTS, 'pdf', 'txt', 'hml']);
const CREATE_EXTS = new Set(['hwp', 'hwpx', 'docx', 'pptx', 'xlsx']);
const HWP_EXTS = new Set(['hwp', 'hwpx', 'hwt']);

// 아직 없는 경로는 있는 상위 폴더까지 실제 경로(정션·링크를 푼 것)로 바꿔 비교한다
function realOf(p) {
  let cur = p; const rest = [];
  while (!existsSync(cur)) { const up = dirname(cur); if (up === cur) break; rest.unshift(basename(cur)); cur = up; }
  try { cur = realpathSync.native(cur); } catch {}
  return join(cur, ...rest);
}
const inside = (p, root) => { const r = relative(root, p); return r === '' || (!r.startsWith('..') && !isAbsolute(r)); };

/** OFFICEKIT_WRITE_ROOTS: 저장을 허락한 폴더들 (경로 구분자로 여러 개). 앱 안 AI 대화가 문서 폴더만 쓰게 할 때 */
export function writeRoots(env = process.env) {
  return (env.OFFICEKIT_WRITE_ROOTS ?? '').split(delimiter).map((s) => s.trim()).filter(Boolean).map((r) => realOf(resolve(r)));
}

/**
 * 저장 경로 확인: 형식(확장자), 허락된 폴더, 이미 있는 다른 파일 덮어쓰기.
 * 모두 통과한 뒤에만 폴더를 만든다.
 */
function outPath(output, { source, allow, overwrite = false }) {
  const full = output ? abs(output) : source;
  const ext = extOf(full);
  if (!allow.has(ext)) throw new Error(`.${ext || '(확장자 없음)'} 파일로는 저장할 수 없습니다 (가능: ${[...allow].join(', ')})`);
  const roots = writeRoots();
  if (roots.length && !roots.some((r) => inside(realOf(full), r))) {
    throw new Error(`이 대화에서는 다음 폴더 안에만 저장할 수 있습니다: ${roots.join(', ')} (요청한 경로: ${full})`);
  }
  if (existsSync(full) && !(source && samePath(full, source)) && !overwrite) {
    throw new Error(`이미 있는 파일입니다: ${full} (덮어쓰려면 overwrite=true)`);
  }
  mkdirSync(dirname(full), { recursive: true });
  return full;
}

/** Node/시스템 오류를 한국어로 */
export function koreanError(e) {
  const msg = e?.message ?? String(e);
  const code = e?.code;
  if (code === 'EBUSY' || code === 'EPERM') return `파일이 다른 프로그램에서 열려 있거나 쓰기 권한이 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'EACCES') return `파일에 접근할 권한이 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'ENOENT') return `파일이나 폴더가 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'EISDIR') return `파일이 아니라 폴더입니다: ${e.path ?? ''}`.trim();
  const v = msg.match(/Input validation error:\s*(?:Invalid arguments for tool (\S+):\s*)?([\s\S]*)$/i);
  if (v) {
    // SDK는 문제마다 한 줄 '메시지 at 경로'(경로는 operations[0].op 꼴, 없으면 안 붙음)로 준다.
    // 경로는 맨 끝 ' at ' 뒤 → '(위치: 경로)'. zod 영어 메시지 중 흔한 것은 우리말로
    const count = (n, u) => (u === 'characters' ? `${n}자` : u === 'bytes' ? `${n}바이트` : `${n}개`);
    const detail = v[2].split('\n').map((line) => {
      const at = line.match(/^(.*) at (.+)$/);
      const [text, p] = at ? [at[1], at[2]] : [line, null];
      let m = text.match(/^Invalid input: expected (\w+), received undefined$/);
      if (m) return `${p ? `'${p}' ` : ''}값이 필요합니다 (${m[1]})`;
      m = text.match(/^Invalid input: expected (\w+), received (\w+)$/);
      if (m) return `${p ? `'${p}' ` : ''}값의 형식이 틀립니다 (${m[1]} 필요, ${m[2]} 받음)`;
      const ko = text
        .replace(/^Too small: expected \w+ to have >=?(\S+) (\w+)$/, (_m, n, u) => `${u === 'characters' ? '너무 짧습니다' : '너무 적습니다'} (${count(n, u)} 이상)`)
        .replace(/^Too big: expected \w+ to have <=?(\S+) (\w+)$/, (_m, n, u) => `${u === 'characters' ? '너무 깁니다' : '너무 많습니다'} (${count(n, u)} 이하)`)
        .replace(/^Too small: expected \w+ to be (>=?)(\S+)$/, (_m, op, n) => `값이 너무 작습니다 (${n} ${op === '>' ? '초과' : '이상'})`)
        .replace(/^Too big: expected \w+ to be (<=?)(\S+)$/, (_m, op, n) => `값이 너무 큽니다 (${n} ${op === '<' ? '미만' : '이하'})`)
        .replace(/^Invalid option: expected one of (.+)$/, (_m, list) => `다음 중 하나여야 합니다: ${list.replace(/\|/g, ', ')}`)
        .replace(/^Unrecognized keys?: (.+)$/, (_m, keys) => `알 수 없는 인수: ${keys}`)
        .replace(/^Invalid input$/, '값이 올바르지 않습니다');
      return p ? `${ko} (위치: ${p})` : ko;
    }).join('\n');
    return `입력값이 올바르지 않습니다${v[1] ? ` (${v[1]})` : ''}: ${detail}`;
  }
  return msg.replace(/^MCP error -?\d+:\s*/, '');
}

const text = (t) => ({ content: [{ type: 'text', text: t }] });
const json = (o) => JSON.stringify(o, null, 2);

// 실패는 isError + 한국어 메시지로 돌려준다
function tool(fn) {
  return async (...args) => {
    try { return await fn(...args); } catch (e) {
      log('도구 오류:', process.env.OFFICEKIT_DEBUG ? e?.stack ?? e : koreanError(e));
      return { content: [{ type: 'text', text: `오류: ${koreanError(e)}` }], isError: true };
    }
  };
}

const engineSchema = z.enum(['auto', 'app', 'file']).default('auto')
  .describe("'auto'(기본: 읽기는 파일 엔진, 쓰기는 설치된 프로그램 우선) | 'app'(설치된 한컴/Word/PowerPoint/Excel) | 'file'(프로그램 없이 직접 파일 처리)");
const pathSchema = (what = '문서 파일 경로') => z.string().min(1).describe(`${what} (절대 경로 또는 작업 폴더 기준 상대 경로)`);
const outputSchema = z.string().optional().describe('저장할 경로 (원본과 같은 형식). 생략하면 원본 파일을 덮어씀');
const overwriteSchema = z.boolean().default(false).describe('output이 이미 있는 다른 파일일 때 덮어쓸지 (기본 false: 거부)');
const scalar = z.union([z.string(), z.number()]);

const SUPPORTED = 'hwp, hwpx, hwt, docx, pptx, xlsx, csv(읽기) (프로그램이 있으면 doc, ppt, xls, rtf, odt 등도)';

// 원고 문법 (src/core/create.mjs parseMarkdown 과 같음)
const MARKDOWN_HELP = [
  '원고 문법 (한 줄 = 한 덩어리):',
  '  # 제목 / ## 소제목 / ### 작은 제목  (# 제목은 가운데 정렬. pptx는 # 마다 새 슬라이드)',
  '  >> 가운데 정렬 문단',
  '  **굵은 한 줄**  (줄 전체가 ** 로 감싸진 경우)',
  '  - 글머리표 항목',
  '  | 칸1 | 칸2 |  표 (연속된 | 줄이 한 표, 첫 행은 굵게, |---| 구분줄은 무시)',
  '  라벨: {{필드}}  또는  {{필드}}  → 한글(hwp/hwpx)은 누름틀, 그 외 형식은 {{필드}} 자리표시자. 나중에 office_fill로 채운다',
  '  ---page---  쪽 나누기 (pptx는 새 슬라이드)',
  '  그 밖의 줄은 일반 문단, 빈 줄은 빈 문단. 기울임·링크·번호 목록 등 다른 마크다운 문법은 지원하지 않음',
].join('\n');

// 긴 글자를 자를 때 서로게이트 쌍이 갈라지지 않게
function cut(t, start, end) {
  if (end < t.length && /[\uD800-\uDBFF]/.test(t[end - 1] ?? '')) end -= 1;
  return { slice: t.slice(start, end), end };
}

// 표·시트·슬라이드 구조를 정해진 글자 수 안에서 요약
const widthOf = (rows) => rows.reduce((n, r) => Math.max(n, r?.length ?? 0), 0);
function structureOf(doc, budget) {
  const parts = [];
  let used = 0;
  if (doc.pages) parts.push(`쪽 수: ${doc.pages}`);
  if (doc.sheets?.length) {
    parts.push(`시트 ${doc.sheets.length}개: ${doc.sheets.map((s) => `${s.name}(${s.rows.length}행×${widthOf(s.rows)}열`
      + `${s.truncated ? `, 잘림: 전체 ${s.totalRows}행×${s.totalCols}열` : ''})`).join(', ')}`);
  }
  if (doc.slides?.length) parts.push(`슬라이드 ${doc.slides.length}장`);
  if (doc.tables?.length) {
    const shown = [];
    for (const [i, t] of doc.tables.entries()) {
      const s = `표 ${i + 1} (${t.length}행×${widthOf(t)}열): ${JSON.stringify(t)}`;
      if (used + s.length > budget) break;
      shown.push(s); used += s.length;
    }
    parts.push(`표 ${doc.tables.length}개${shown.length < doc.tables.length ? ` (글자 수 제한으로 ${shown.length}개만 표시)` : ''}`);
    parts.push(...shown);
  }
  return parts.join('\n');
}

// ── 서버 ─────────────────────────────────────────────────

export function createServer() {
  const server = new McpServer({ name: 'office-kit', version: VERSION });
  // 입력 검증 오류 등 SDK가 만드는 오류 메시지도 한국어로
  const origError = server.createToolError?.bind(server);
  if (origError) server.createToolError = (msg) => ({ content: [{ type: 'text', text: `오류: ${koreanError({ message: msg })}` }], isError: true });

  server.registerTool('office_status', {
    title: '설치 상태 확인',
    description: '이 PC에 설치된 문서 프로그램(한컴오피스·Word·PowerPoint·Excel), 한글 자동화 보안 모듈 등록 여부, 어떤 엔진으로 작업하는지 알려준다. '
      + '작업 전에 무엇이 가능한지(PDF 변환, 원본 서식 유지 편집 등) 확인할 때 먼저 호출한다.',
    annotations: { readOnlyHint: true },
  }, tool(async () => {
    const p = await probe();
    let hancom = null;
    try { hancom = await (await import('../setup/hancom.mjs')).hancomSecurityStatus(); } catch {}
    const yes = (v) => (v ? '있음' : '없음');
    const hwpApp = p.hancom && p.hancomSecurityModule;
    const hwpPdf = hwpPdfStatus();
    const lines = [
      `office-kit ${VERSION} · 작업 폴더: ${process.cwd()}`,
      `설치된 프로그램: 한컴오피스 ${yes(p.hancom)}, Word ${yes(p.word)}, PowerPoint ${yes(p.powerpoint)}, Excel ${yes(p.excel)}`,
      `한글 자동화 보안 모듈: ${p.hancomSecurityModule ? '등록됨' : '등록 안 됨'}`
        + (p.hancom && !p.hancomSecurityModule ? ' → 한글 문서는 rhwp(파일 엔진)로 처리. 한컴오피스로 처리하려면 Office Kit 설정의 "한글 자동화 연결"' : ''),
      '',
      '엔진 안내:',
      '- 파일 엔진(기본 읽기): hwp/hwpx는 rhwp, docx/pptx는 XML 직접 처리, xlsx는 exceljs. 프로그램 없이 빠르게 동작',
      `- 프로그램 엔진(기본 쓰기): 설치된 프로그램을 COM으로 조작해 원본 서식을 그대로 유지. 한글 ${hwpApp ? '가능' : '불가'}, Word ${p.word ? '가능' : '불가'}, PowerPoint ${p.powerpoint ? '가능' : '불가'}, Excel ${p.excel ? '가능' : '불가'}`,
      '- hwp↔hwpx 변환과 한글 쪽 그림(office_page_image)은 프로그램 없이 가능',
      `- 한글→PDF: ${hwpApp ? '한컴오피스로 (원본 그대로)' : hwpPdf.ok ? '내장 엔진(rhwp)으로, 한컴오피스 없이 — 앱 미리보기와 같은 모양 (글꼴·줄바꿈이 한컴과 조금 다를 수 있음)' : `불가 (${hwpPdf.reason})`}`,
      '- Word/PowerPoint/Excel의 PDF 변환과 쪽 그림은 해당 프로그램이 필요',
      `- 다루는 형식: ${SUPPORTED}`,
    ];
    if (hancom?.registered) lines.splice(3, 0, `  보안 모듈 파일: ${hancom.dllPath}${hancom.dllExists ? '' : ' (파일 없음)'}`);
    if (hancom?.warning) lines.splice(3, 0, `  주의: ${hancom.warning}`);
    return text(lines.join('\n'));
  }));

  server.registerTool('office_info', {
    title: '문서 정보',
    description: '문서 파일의 기본 정보(종류, 크기, 쪽 수, 표·누름틀 개수, 사용할 수 있는 엔진)를 돌려준다. 내용은 office_read로 읽는다.',
    inputSchema: { path: pathSchema() },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path }) => {
    const info = await documentInfo(mustExist(path));
    const head = `${info.name} — ${info.kind} (.${info.format}), ${info.size.toLocaleString()}바이트`
      + (info.pages ? `, ${info.pages}쪽` : '');
    return text(`${head}\n${json(info)}`);
  }));

  server.registerTool('office_read', {
    title: '문서 읽기',
    description: '문서(hwp, hwpx, docx, pptx, xlsx, pdf 등)의 글자를 읽는다. PDF는 쪽마다 "--- 쪽 n ---"으로 나누고, 글자가 없는 스캔본이면 office_page_image로 쪽 그림을 본다. 긴 문서는 max_chars만큼 잘라 주고 전체 길이와 다음 offset을 알려주므로, '
      + "'다음: offset=N'이 보이면 그 offset으로 다시 호출해 나머지를 읽는다. 첫 부분(offset=0)에는 표·시트·슬라이드 구조 요약도 붙는다.",
    inputSchema: {
      path: pathSchema(),
      offset: z.number().int().min(0).default(0).describe('읽기 시작할 글자 위치 (0부터)'),
      max_chars: z.number().int().min(1000).max(200000).default(40000).describe('한 번에 돌려줄 최대 글자 수'),
      engine: engineSchema,
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, offset = 0, max_chars = 40000, engine = 'auto' }) => {
    const full = mustExist(path);
    const doc = await readDocument(full, { engine });
    const all = doc.text ?? '';
    const total = all.length;
    if (offset > total) throw new Error(`offset(${offset})이 전체 글자 수(${total})보다 큽니다`);
    const { slice, end } = cut(all, offset, Math.min(offset + max_chars, total));
    const more = end < total;
    const lines = [
      `${basename(full)} — ${doc.kind} (엔진: ${doc.engine})`,
      `글자 ${offset}~${end} / 전체 ${total}자${more ? ` — 아직 남음. 다음: offset=${end}` : doc.truncated ? '' : ' (끝까지 읽음)'}`,
    ];
    if (doc.scanned) lines.push('(이 PDF는 글자가 거의 없습니다 — 스캔한 문서일 수 있습니다. office_page_image로 쪽 그림을 보고 읽으세요)');
    // 엑셀 프로그램으로 읽은 큰 시트는 앞부분만 온다
    if (doc.truncated && doc.sheets) {
      lines.push(`(잘림: 문서가 커서 일부만 읽었습니다 — ${(doc.sheets ?? []).filter((s) => s.truncated)
        .map((s) => `시트 ${s.name}: 전체 ${s.totalRows}행×${s.totalCols}열 중 ${s.rows.length}행×${widthOf(s.rows)}열`).join(', ')})`);
    }
    if (offset === 0) {
      const s = structureOf(doc, Math.min(20000, Math.floor(max_chars / 2)));
      if (s) lines.push('', '[구조]', s);
    }
    lines.push('', '[본문]', slice);
    if (more) lines.push('', `[계속] 다음 부분은 offset=${end} 로 다시 읽으세요.`);
    return text(lines.join('\n'));
  }));

  server.registerTool('office_fields', {
    title: '채울 칸 목록',
    description: '서식 문서에서 채울 수 있는 칸 목록: 한글 누름틀, Word 콘텐츠 컨트롤·책갈피, {{이름}} 형태의 자리표시자. 결과의 이름을 office_fill에 넘긴다.',
    inputSchema: { path: pathSchema(), engine: engineSchema },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, engine = 'auto' }) => {
    const fields = await listFields(mustExist(path), { engine });
    if (!fields.length) return text('채울 칸(누름틀·자리표시자)이 없습니다.');
    return text(`채울 칸 ${fields.length}개:\n${fields.map((f) => `- ${f.name} [${f.kind ?? ''}]${f.value ? ` 현재값: ${String(f.value).slice(0, 60)}` : ''}`).join('\n')}`);
  }));

  server.registerTool('office_fill', {
    title: '칸 채우기',
    description: '누름틀·콘텐츠 컨트롤·책갈피·{{이름}} 자리표시자(본문·표·머리글·바닥글, {{ 이름 }}처럼 띄어 써도 됨)를 이름으로 찾아 값을 채운다. '
      + '이름은 office_fields로 확인. output을 생략하면 원본을 덮어쓴다 (output은 원본과 같은 형식, 한글은 hwp·hwpx끼리 가능).',
    inputSchema: {
      path: pathSchema(),
      values: z.record(z.string(), scalar).describe('{ "칸 이름": "값" }'),
      output: outputSchema,
      overwrite: overwriteSchema,
      engine: engineSchema,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, values, output, overwrite = false, engine = 'auto' }) => {
    const full = mustExist(path);
    const vals = Object.fromEntries(Object.entries(values).map(([k, v]) => [k, String(v)]));
    if (!Object.keys(vals).length) throw new Error('채울 값이 없습니다');
    const r = await fillFields(full, vals, { output: outPath(output, { source: full, allow: EDIT_EXTS, overwrite }), engine });
    return text([
      `채움 ${r.filled.length}개: ${r.filled.join(', ') || '없음'}`,
      r.missing.length ? `못 찾음 ${r.missing.length}개: ${r.missing.join(', ')} (office_fields로 이름 확인)` : null,
      `저장: ${r.output}`,
    ].filter(Boolean).join('\n'));
  }));

  server.registerTool('office_replace', {
    title: '찾아 바꾸기',
    description: '문서 전체(본문·표·머리글 등)에서 글자를 찾아 바꾼다. 영문 대소문자를 구분한다(모든 형식 같음). 여러 쌍을 한 번에 처리. '
      + '서식은 최대한 유지. output을 생략하면 원본을 덮어쓴다.',
    inputSchema: {
      path: pathSchema(),
      replacements: z.array(z.object({ find: z.string().min(1), replace: z.string() })).min(1).describe('[{ "find": "찾을 말", "replace": "바꿀 말" }]'),
      output: outputSchema,
      overwrite: overwriteSchema,
      engine: engineSchema,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, replacements, output, overwrite = false, engine = 'auto' }) => {
    const full = mustExist(path);
    const r = await replaceText(full, replacements, { output: outPath(output, { source: full, allow: EDIT_EXTS, overwrite }), engine });
    return text(`${r.count}곳 바꿈 (엔진: ${r.engine})${r.count === 0 ? ' — 찾는 글자가 없습니다' : ''}\n저장: ${r.output}`);
  }));

  server.registerTool('office_set_cells', {
    title: '엑셀 셀 쓰기',
    description: "엑셀(xlsx 등) 셀에 값이나 수식을 쓴다. cells: { \"A1\": \"제목\", \"B2\": 120, \"C2\": \"=SUM(B2:B9)\" } ('='로 시작하면 수식). "
      + 'sheet를 생략하면 첫 시트. output을 생략하면 원본을 덮어쓴다.',
    inputSchema: {
      path: pathSchema('엑셀 파일 경로'),
      cells: z.record(z.string(), scalar).describe('{ "셀 주소": 값 }'),
      sheet: z.string().optional().describe('시트 이름 (생략하면 첫 시트)'),
      output: outputSchema,
      overwrite: overwriteSchema,
      engine: engineSchema,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, cells, sheet, output, overwrite = false, engine = 'auto' }) => {
    const full = mustExist(path);
    const bad = Object.keys(cells).filter((a) => !/^\$?[A-Za-z]{1,3}\$?\d{1,7}$/.test(a));
    if (bad.length) throw new Error(`셀 주소가 올바르지 않습니다: ${bad.join(', ')} (예: A1, B12)`);
    const r = await setCells(full, cells, { sheet, output: outPath(output, { source: full, allow: EDIT_EXTS, overwrite }), engine });
    return text(`셀 ${r.updated}개 씀 (엔진: ${r.engine})\n저장: ${r.output}`);
  }));

  server.registerTool('office_convert', {
    title: '형식 변환',
    description: '문서를 다른 형식으로 저장한다. 출력 형식은 output 확장자로 정한다. 가능한 형식:\n'
      + '  한글(hwp/hwpx/hwt) → hwp, hwpx, hml, pdf (프로그램 없이) / docx, odt, txt (한컴오피스 + 자동화 연결 필요)\n'
      + '    한글→pdf: 한컴오피스 + 자동화 연결이 있으면 한컴으로(원본 그대로), 없으면 내장 엔진(rhwp: 앱 미리보기와 같은 모양, 글자 선택·검색 가능, '
      + '글꼴·줄바꿈은 한컴과 조금 다를 수 있음)\n'
      + '  Word(docx/doc/rtf/odt) → pdf, docx, doc, rtf, odt, txt\n'
      + '  PowerPoint(pptx/ppt/odp) → pdf, pptx, ppt, odp\n'
      + '  Excel(xlsx/xls/xlsm/csv/ods) → pdf, xlsx, xlsm, xls, csv, ods\n'
      + 'Word·PowerPoint·Excel 문서는 한글(hwp/hwpx)로 바꿀 수 없다. 한글 이외는 해당 프로그램이 설치돼 있어야 한다. '
      + '원본은 그대로 두고, 이미 있는 파일은 overwrite=true일 때만 덮어쓴다.',
    inputSchema: {
      path: pathSchema('원본 문서 경로'),
      output: z.string().min(1).describe('저장할 경로 (확장자로 형식 결정)'),
      overwrite: overwriteSchema,
    },
  }, tool(async ({ path, output, overwrite = false }) => {
    const full = mustExist(path);
    if (samePath(abs(output), full)) throw new Error('원본과 다른 경로를 지정하세요');
    const out = outPath(output, { allow: CONVERT_EXTS, overwrite });
    const r = await convertDocument(full, out);
    const how = r.note ? `엔진: 내장(${r.engine})` : `엔진: ${r.engine}`;
    return text(`변환 완료 (${how}${r.pages ? `, ${r.pages}쪽` : ''}): ${r.output}${r.note ? `\n참고: ${r.note}` : ''}`);
  }));

  server.registerTool('office_create', {
    title: '새 문서 만들기',
    description: `마크다운 비슷한 원고로 새 문서를 만든다. 형식은 output 확장자로 정한다: .hwp .hwpx .docx .pptx .xlsx\n${MARKDOWN_HELP}\n`
      + '이미 있는 파일은 overwrite=true일 때만 덮어쓴다. 만든 뒤 office_page_image로 모양을 확인할 수 있다.',
    inputSchema: {
      output: z.string().min(1).describe('만들 파일 경로 (.hwp/.hwpx/.docx/.pptx/.xlsx)'),
      markdown: z.string().describe('원고 내용'),
      overwrite: z.boolean().default(false).describe('같은 이름의 파일이 있으면 덮어쓸지'),
    },
  }, tool(async ({ output, markdown, overwrite = false }) => {
    const out = outPath(output, { allow: CREATE_EXTS, overwrite });
    await createDocument(out, markdown);
    return text(`새 ${kindOf(out)} 문서를 만들었습니다: ${out}`);
  }));

  server.registerTool('office_page_image', {
    title: '쪽 그림 보기',
    description: '문서의 한 쪽(슬라이드·시트)을 PNG 그림으로 돌려준다. 수정 결과를 눈으로 확인할 때 쓴다. '
      + 'hwp/hwpx/hwt와 pdf는 프로그램 없이 되고, docx/pptx/xlsx는 설치된 프로그램으로 PDF로 바꾼 뒤 그린다.',
    inputSchema: {
      path: pathSchema(),
      page: z.number().int().min(1).default(1).describe('쪽 번호 (1부터)'),
      dpi: z.number().int().min(48).max(200).default(110).describe('해상도 (96 = 원래 크기, 기본 110)'),
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, page = 1, dpi = 110 }) => {
    const full = mustExist(path);
    const r = await pageImage(full, page, { dpi });
    return {
      content: [
        { type: 'image', data: r.png.toString('base64'), mimeType: 'image/png' },
        { type: 'text', text: `${basename(full)} ${r.page}/${r.pages}쪽 (엔진: ${r.engine}${r.width ? `, ${r.width}×${r.height}px` : ''})` },
      ],
    };
  }));

  const piiTypes = z.array(z.enum(Object.keys(PII_TYPES))).optional()
    .describe(`찾을 종류 (생략하면 모두): ${Object.entries(PII_TYPES).map(([k, v]) => `${k}=${v}`).join(', ')}`);

  server.registerTool('office_find_pii', {
    title: '개인정보 찾기',
    description: '문서에서 개인정보(주민·외국인등록번호, 휴대전화·전화번호, 이메일, 카드번호, 계좌번호(계좌·은행 같은 말이 앞에 있을 때), '
      + '여권번호, 운전면허번호)를 글자 규칙으로 찾는다. 인터넷·AI 없이 동작하고 파일을 바꾸지 않는다. '
      + '같은 값은 한 번만, 몇 번 나오는지(count)와 가릴 모양(masked)을 함께 돌려준다. 그림 속 글자는 찾지 못한다.',
    inputSchema: { path: pathSchema(), types: piiTypes },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, types }) => {
    const full = mustExist(path);
    const r = await scanDocument(full, { types });
    if (!r.items.length) return text(`개인정보를 찾지 못했어요 (${basename(full)}, 글자 ${r.chars}자 확인)`);
    return text(`${basename(full)}에서 ${r.items.length}가지(모두 ${r.total}번) 찾음\n${json(r.items.map(({ label, value, masked, count, context }) => ({ label, value, masked, count, context })))}`);
  }));

  server.registerTool('office_mask_pii', {
    title: '개인정보 가린 사본 만들기',
    description: '문서의 개인정보를 가린 사본을 만든다 (예: 900101-1234567 → 900101-1******, 010-1234-5678 → 010-****-5678). '
      + '원본은 바꾸지 않고 output(생략하면 원본 옆 "<이름>_가림.<확장자>")에 원본과 같은 형식으로 저장한 뒤, 사본을 다시 읽어 남은 것(remaining)을 알려 준다. '
      + 'values로 가릴 값만 고를 수 있다 (office_find_pii의 value). 한글·Word·PowerPoint·Excel(xlsx)은 프로그램 없이 된다.',
    inputSchema: {
      path: pathSchema(),
      output: z.string().optional().describe('저장할 경로 (원본과 같은 형식, 원본과 다른 이름). 생략하면 "<이름>_가림.<확장자>"'),
      types: piiTypes,
      values: z.array(z.string().min(1)).optional().describe('가릴 값만 고르기 (생략하면 찾은 것 모두)'),
      overwrite: overwriteSchema,
    },
    annotations: { destructiveHint: false },
  }, tool(async ({ path, output, types, values, overwrite = false }) => {
    const full = mustExist(path);
    const out = outPath(output ?? maskedPathFor(full), { allow: new Set([extOf(full)]), overwrite });
    if (samePath(out, full)) throw new Error('원본에 바로 덮어쓰지 않아요. 다른 이름을 지정하세요');
    const r = await maskDocument(full, { output: out, types, values });
    const lines = r.items.map((it) => `- ${it.label}: ${it.value} → ${it.masked} (${it.count}곳)`);
    const left = r.remaining.length
      ? `\n주의: 가리지 못하고 남은 것 ${r.remaining.length}가지 — ${r.remaining.map((x) => x.value).join(', ')} (글자가 여러 조각으로 나뉘어 저장된 경우 등. 직접 확인하세요)`
      : '\n확인: 가린 사본을 다시 읽어 고른 개인정보가 남지 않은 것을 확인함';
    return text(`가린 사본 저장: ${r.output}\n${lines.join('\n')}${left}`);
  }));

  server.registerTool('office_hwpx_batch', {
    title: '한글 → HWPX 일괄 변환',
    description: '여러 한글(.hwp) 문서를 HWPX로 한꺼번에 바꾸고, 바꾼 파일을 다시 열어 쪽 수·표 수·글이 원본과 같은지 확인한다. '
      + 'paths에 파일이나 폴더를 준다 (폴더는 하위 폴더까지). 결과는 원본 옆 같은 이름 .hwpx (outDir를 주면 그 폴더에 하위 구조 그대로). '
      + '원본은 바꾸지 않고, 같은 이름 .hwpx가 이미 있으면 overwrite=true일 때만 덮어쓴다. report에 .csv 경로를 주면 결과 보고서를 쓴다. '
      + '한컴오피스+자동화 연결이 있으면 한컴으로, 없으면 내장 엔진(rhwp)으로 바꾼다.',
    inputSchema: {
      paths: z.array(z.string().min(1)).min(1).max(200).describe('바꿀 .hwp 파일이나 폴더 경로들'),
      outDir: z.string().optional().describe('결과를 모을 폴더 (생략하면 원본 옆)'),
      overwrite: overwriteSchema,
      report: z.string().optional().describe('결과 보고서(.csv)를 쓸 경로 (선택)'),
    },
    annotations: { destructiveHint: false },
  }, tool(async ({ paths, outDir, overwrite = false, report }) => {
    const { items, truncated } = collectHwpFiles(paths.map(abs));
    if (!items.length) throw new Error('바꿀 한글(.hwp) 파일을 찾지 못했어요');
    const dir = outDir ? abs(outDir) : undefined;
    // 저장할 곳이 모두 허락된 폴더 안인지 먼저 확인 (하나라도 밖이면 시작하지 않는다)
    for (const it of items) outPath(outputFor(it, dir), { allow: new Set(['hwpx']), overwrite: true });
    const reportPath = report ? outPath(report, { allow: new Set(['csv']), overwrite: true }) : null;
    const r = await convertHwpBatch(items, { outDir: dir, overwrite });
    if (reportPath) writeBatchReport(reportPath, r.results);
    const s = r.summary;
    const lines = r.results.map((x) => `- [${BATCH_STATUS_KO[x.status]}] ${x.src} → ${x.out}${x.notes?.length ? ` (${x.notes.join(', ')})` : ''}`);
    return text(`한글 → HWPX ${s.total}개: 같음 ${s.ok}, 확인 필요 ${s.check}, 건너뜀 ${s.skipped}, 실패 ${s.fail}${truncated ? ' (파일이 많아 앞부분만)' : ''}`
      + `${reportPath ? `\n보고서: ${reportPath}` : ''}\n${lines.join('\n')}`);
  }));

  // 형식별 전용 도구 (src/mcp/tools/*.mjs)
  const ctx = {
    z, scalar, tool, text, json, log, abs, mustExist, outPath, koreanError,
    pathSchema, outputSchema, overwriteSchema, engineSchema, EDIT_EXTS, CONVERT_EXTS, HWP_EXTS,
  };
  for (const register of TOOL_MODULES) register(server, ctx);

  return server;
}

/** stdio로 서버 시작 (node src/mcp/server.mjs 또는 office-kit mcp) */
export async function startServer() {
  // 다른 모듈의 console.log가 프로토콜 출력(stdout)을 망치지 않도록 stderr로 돌린다
  console.log = console.info = console.debug = (...a) => console.error(...a);
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stdin.on('end', () => process.exit(0));
  log(`MCP 서버 시작 (v${VERSION}, 작업 폴더 ${process.cwd()}, 형식: ${OPENABLE_EXTS.join(' ')})`);
  return server;
}

function isMain() {
  if (!process.argv[1]) return false;
  try {
    const a = realpathSync(process.argv[1]);
    const b = realpathSync(fileURLToPath(import.meta.url));
    return process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b;
  } catch { return false; }
}

if (isMain()) {
  startServer().catch((e) => { log('시작 실패:', e?.stack ?? e); process.exit(1); });
}
