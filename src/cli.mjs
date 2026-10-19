#!/usr/bin/env node
// cli.mjs — 명령줄 도구 office-kit.
// node src/cli.mjs ... 또는 Electron 실행 파일 + ELECTRON_RUN_AS_NODE=1 로 실행한다.
import { existsSync, mkdirSync, readFileSync, realpathSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const APP_ROOT = resolve(here, '..');
const VERSION = (() => {
  try { return JSON.parse(readFileSync(resolve(APP_ROOT, 'package.json'), 'utf8')).version; } catch { return '0.0.0'; }
})();

// 엔진 모듈은 필요할 때만 불러온다 (help가 빠르도록)
const core = () => import('./core/index.mjs');

const HELP = `office-kit ${VERSION} — 한글·Word·PowerPoint·Excel 문서 도구

사용법: office-kit <명령> [인수] [옵션]

  status                              설치된 프로그램·한글 보안 모듈·연결 상태
  info <파일>                         문서 기본 정보 (--json: JSON으로)
  text <파일> [--out 파일]            문서 글자 출력 (--out: 파일로 저장)
  fields <파일>                       채울 칸(누름틀·자리표시자) 목록
  fill <파일> <값.json> [-o 출력]     칸 채우기. 값은 {"이름":"값"} JSON 파일 또는 JSON 문자열
  replace <파일> <찾을말> <바꿀말> [-o 출력]
                                      찾아 바꾸기 (영문 대소문자 구분)
  cells <파일> <셀.json> [--sheet 시트] [-o 출력]
                                      엑셀 셀 쓰기. {"A1":"제목","B2":120,"C2":"=SUM(B2:B9)"}
  convert <입력> <출력>               형식 변환 (pdf, hwp↔hwpx, docx 등 — 출력 확장자로 결정)
                                      한글→pdf는 한컴오피스(+자동화 연결)가 있으면 한컴으로, 없으면 내장 엔진(rhwp)으로
  new <원고.md> <출력> [--force]      원고로 새 문서 만들기 (.hwp .hwpx .docx .pptx .xlsx)
  image <파일> [쪽] [-o 그림.png] [--dpi 110]
                                      한 쪽을 PNG 그림으로 저장 (PowerPoint는 쪽 = 슬라이드 번호, 숨긴 슬라이드 포함)
  inspect <파일> [--offset n] [--limit n] [--slide n] [--json]
                                      형식별 구조 보기: 문단·표·스타일(Word·한글), 슬라이드·도형(PowerPoint),
                                      시트·표·차트(Excel). edit에 쓸 번호를 여기서 확인
                                      (--offset/--limit: Word·한글은 문단, PowerPoint는 슬라이드 나눠 보기, --slide: 슬라이드 한 장만)
  edit <파일> <작업.json> [-o 출력] [--overwrite]
                                      형식별 편집: 작업 목록을 순서대로 모두 적용한 뒤 한 번만 저장
                                      (하나라도 실패하면 저장 안 함). 작업 종류: office-kit help edit
  mcp                                 MCP 서버 실행 (stdio, Claude/Codex 연결용)
  setup <대상> [--remove]             연결 설정. 대상: claude-desktop | claude-code | codex | cli | hancom
                                      (hancom: --dll <경로> 로 보안 모듈 파일 지정, --download 로 한컴 공식 배포본 받기)
  help [edit]                         이 도움말 (edit: 형식별 편집 작업 목록)

공통 옵션:
  -o, --out <경로>    결과를 저장할 경로 (생략하면 원본을 덮어씀)
  --engine <방식>     auto(기본) | app(설치된 프로그램) | file(프로그램 없이 파일 직접 처리)
  --overwrite         edit: -o 경로에 이미 다른 파일이 있어도 덮어씀

원고 문법 (new):
  # 제목 / ## 소제목 / ### 작은 제목    >> 가운데 줄    **굵은 한 줄**    - 글머리표
  | 칸 | 칸 | 표 (|---| 줄 무시)         라벨: {{필드}}  (한글은 누름틀, 그 외는 자리표시자)
  ---page---  쪽 나누기 (pptx는 새 슬라이드)

예:
  office-kit text 보고서.hwp
  office-kit fill 신청서.hwpx 값.json -o 신청서_채움.hwpx
  office-kit replace 계약서.docx 갑 을
  office-kit convert 보고서.hwp 보고서.hwpx
  office-kit convert 보고서.hwp 보고서.pdf
  office-kit image 보고서.hwp 2 -o 2쪽.png
  office-kit inspect 발표.pptx --slide 2
  office-kit edit 보고서.docx 작업.json -o 보고서_수정.docx
`;

// ── 인수 해석 ────────────────────────────────────────────

const VALUED = { '-o': 'out', '--out': 'out', '--output': 'out', '--sheet': 'sheet', '--engine': 'engine', '--dpi': 'dpi', '--dll': 'dll',
  '--offset': 'offset', '--limit': 'limit', '--slide': 'slide' };
const FLAGS = { '--remove': 'remove', '--download': 'download', '--json': 'json', '--force': 'force', '--overwrite': 'overwrite',
  '-h': 'help', '--help': 'help', '-v': 'version', '--version': 'version' };

export function parseArgs(argv) {
  const pos = []; const opt = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--') { pos.push(...argv.slice(i + 1)); break; }
    const eq = a.match(/^(--[\w-]+)=(.*)$/s);
    if (eq && VALUED[eq[1]]) { opt[VALUED[eq[1]]] = eq[2]; continue; }
    if (VALUED[a]) {
      if (i + 1 >= argv.length) throw new Error(`${a} 뒤에 값이 필요합니다`);
      opt[VALUED[a]] = argv[++i]; continue;
    }
    if (FLAGS[a]) { opt[FLAGS[a]] = true; continue; }
    if (/^--?[A-Za-z]/.test(a)) throw new Error(`알 수 없는 옵션: ${a} (도움말: office-kit help)`);
    pos.push(a);
  }
  if (opt.engine && !['auto', 'app', 'file'].includes(opt.engine)) throw new Error(`--engine 은 auto, app, file 중 하나여야 합니다: ${opt.engine}`);
  return { pos, opt };
}

function need(pos, n, usage) {
  if (pos.length < n) throw new Error(`인수가 부족합니다. 사용법: office-kit ${usage}`);
}

function fileArg(p) {
  const full = resolve(p);
  if (!existsSync(full)) throw new Error(`파일이 없습니다: ${full}`);
  return full;
}

/** JSON 파일 경로 또는 JSON 문자열 → 값 */
function readJson(arg, what) {
  let src = arg;
  if (!/^\s*[{[]/.test(arg)) {
    const full = resolve(arg);
    if (!existsSync(full)) throw new Error(`${what} 파일이 없습니다: ${full}`);
    src = readFileSync(full, 'utf8').replace(/^﻿/, '');
  }
  try { return JSON.parse(src); } catch (e) { throw new Error(`${what}을(를) JSON으로 읽을 수 없습니다: ${e.message}`); }
}

/** JSON 파일 경로 또는 JSON 문자열 → 객체 */
function jsonArg(arg, what) {
  const v = readJson(arg, what);
  if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Error(`${what}은(는) {"이름": 값} 형태의 JSON 객체여야 합니다`);
  return v;
}

/** 작업 목록: [{op…}] 배열, {operations:[…]}, 또는 작업 하나 {op…} */
function opsArg(arg) {
  const v = readJson(arg, '작업 목록');
  const ops = Array.isArray(v) ? v : Array.isArray(v?.operations) ? v.operations : (v && typeof v === 'object' && typeof v.op === 'string') ? [v] : null;
  if (!ops) throw new Error('작업 목록은 [{"op": "…", …}, …] 배열(또는 {"operations": […]})이어야 합니다 — office-kit help edit');
  if (!ops.length) throw new Error('작업 목록이 비어 있습니다');
  return ops;
}

function intOpt(v, name, min) {
  if (v === undefined) return undefined;
  if (!/^\d+$/.test(String(v)) || Number(v) < min) throw new Error(`${name} 은 ${min} 이상의 정수여야 합니다: ${v}`);
  return Number(v);
}

const outOf = (opt, fallback) => {
  const out = opt.out ? resolve(opt.out) : fallback;
  if (out) mkdirSync(dirname(out), { recursive: true });
  return out;
};
const print = (s = '') => process.stdout.write(`${s}\n`);

function koreanError(e) {
  const code = e?.code;
  if (code === 'EBUSY' || code === 'EPERM') return `파일이 다른 프로그램에서 열려 있거나 쓰기 권한이 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'EACCES') return `파일에 접근할 권한이 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'ENOENT') return `파일이나 폴더가 없습니다: ${e.path ?? ''}`.trim();
  if (code === 'EISDIR') return `파일이 아니라 폴더입니다: ${e.path ?? ''}`.trim();
  return e?.message ?? String(e);
}

// ── setup (다른 모듈: src/setup/*.mjs) ───────────────────

async function setupModule(name) {
  try { return await import(`./setup/${name}`); } catch (e) {
    if (e?.code === 'ERR_MODULE_NOT_FOUND' && String(e.message).includes(name)) throw new Error(`설정 모듈이 없습니다: src/setup/${name}`);
    throw e;
  }
}

const where = (r) => (typeof r === 'string' ? r : r?.configPath ?? r?.path ?? '');

const TARGETS = {
  'claude-desktop': { name: 'Claude Desktop', on: 'connectClaudeDesktop', off: 'disconnectClaudeDesktop' },
  'claude-code': { name: 'Claude Code', on: 'connectClaudeCode', off: 'disconnectClaudeCode' },
  codex: { name: 'Codex', on: 'connectCodex', off: 'disconnectCodex' },
  cli: { name: '명령줄 office-kit', on: 'installCliShim', off: 'uninstallCliShim' },
};

async function setup(target, opt) {
  if (target === 'hancom') return setupHancom(opt);
  const t = TARGETS[target];
  if (!t) throw new Error(`알 수 없는 연결 대상: ${target ?? '(없음)'} — claude-desktop, claude-code, codex, cli, hancom 중 하나`);
  const m = await setupModule('integrations.mjs');
  const spec = m.mcpLaunchSpec({ execPath: process.execPath, appRoot: APP_ROOT });
  const r = opt.remove ? await m[t.off]() : await m[t.on](spec);
  const changed = r?.changed !== false;
  const msg = opt.remove ? (changed ? '연결을 해제했습니다' : '연결돼 있지 않습니다') : (changed ? '연결했습니다' : '이미 연결돼 있습니다');
  const info = where(r);
  print(`${t.name}: ${msg}${info ? ` (${info})` : ''}`);
  if (r?.backup) print(`원래 설정 백업: ${r.backup}`);
  if (r?.warning) print(`주의: ${r.warning}`);
  if (changed && target !== 'cli') print('적용하려면 해당 프로그램을 다시 시작하세요.');
  if (changed && target === 'cli' && !opt.remove && typeof m.cliShimStatus === 'function' && !m.cliShimStatus(spec).onPath) {
    print('주의: 이 폴더가 PATH에 없어 새 창에서 office-kit 명령이 바로 잡히지 않을 수 있습니다.');
  }
}

async function setupHancom(opt) {
  const m = await setupModule('hancom.mjs');
  const st = await m.hancomSecurityStatus();
  if (opt.remove) {
    if (typeof m.removeHancomSecurity !== 'function') throw new Error('한글 보안 모듈 등록 해제는 지원하지 않습니다');
    const r = await m.removeHancomSecurity();
    print(r?.changed === false ? '한글 자동화 보안 모듈이 등록돼 있지 않습니다.' : '한글 자동화 보안 모듈 등록을 해제했습니다. (DLL 파일은 남겨 둡니다)');
    return;
  }
  if (!st.hancomInstalled) throw new Error('한컴오피스가 설치돼 있지 않습니다');
  if (st.registered && st.dllExists && !opt.dll && !opt.download) {
    print(`이미 등록돼 있습니다: ${st.dllPath}`);
    if (st.warning) print(`주의: ${st.warning}`);
    return;
  }
  let dllPath = opt.dll ? resolve(opt.dll) : null;
  if (!dllPath) {
    if (!opt.download) {
      throw new Error('보안 모듈 파일(DLL)이 필요합니다. --download 로 한컴 공식 배포본을 받거나 --dll <경로> 로 지정하세요');
    }
    // 앱(설정 화면)과 같은 곳에 받는다: %APPDATA%\office-kit\hancom
    const dest = resolve(process.env.APPDATA ?? APP_ROOT, 'office-kit', 'hancom');
    print('한컴 공식 배포본을 받는 중...');
    const got = await m.downloadHancomSecurityModule(dest);
    dllPath = typeof got === 'string' ? got : got?.dllPath;
    if (!dllPath) throw new Error('보안 모듈을 받지 못했습니다');
  }
  const after = await m.setupHancomSecurity({ dllPath });
  print(`한글 자동화 보안 모듈을 등록했습니다: ${after?.dllPath ?? dllPath}`);
  if (after?.warning) print(`주의: ${after.warning}`);
}

// ── 형식별 전용 기능 (inspect / edit) — docs/FEATURES.md ──

const FORMAT_OF = { hwp: 'hwp', hwpx: 'hwp', hwt: 'hwp', docx: 'word', doc: 'word', rtf: 'word', odt: 'word',
  pptx: 'ppt', ppt: 'ppt', odp: 'ppt', xlsx: 'excel', xlsm: 'excel', xls: 'excel', csv: 'excel', ods: 'excel' };
const FORMAT_NAME = { hwp: '한글', word: 'Word', ppt: 'PowerPoint', excel: 'Excel' };
const HWP_SAVE = new Set(['hwp', 'hwpx', 'hwt']);
const ENGINE_NAME = { rhwp: '내장 엔진(rhwp)', word: 'Word', powerpoint: 'PowerPoint', excel: 'Excel', file: '파일 엔진(exceljs)' };
const extOfPath = (p) => extname(p).slice(1).toLowerCase();
const samePath = (a, b) => (process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b);

function formatOf(full) {
  const f = FORMAT_OF[extOfPath(full)];
  if (!f) {
    throw new Error(`형식별 기능을 쓸 수 없는 파일입니다: ${basename(full)} `
      + '(가능: hwp hwpx hwt · docx doc rtf odt · pptx ppt odp · xlsx xlsm xls csv ods)');
  }
  return f;
}

// 엔진: 한글은 rhwp만, Word·PowerPoint는 설치된 프로그램만, Excel은 둘 다
function checkEngine(format, engine) {
  if (!engine || engine === 'auto') return;
  if (format === 'hwp' && engine === 'app') throw new Error('한글 전용 기능은 내장 엔진(rhwp)으로만 합니다 — --engine app은 쓸 수 없습니다');
  if ((format === 'word' || format === 'ppt') && engine === 'file') {
    const n = FORMAT_NAME[format];
    throw new Error(`${n} 전용 기능은 설치된 ${n}로만 합니다 — --engine file은 쓸 수 없습니다 (프로그램 없이: text, replace, fill, new)`);
  }
}

const hwpTools = () => import('./mcp/tools/hwp.mjs');

// core 오류·참고 글의 MCP 표기(engine:'app', output)를 명령줄 옵션 표기로
const cliText = (s) => String(s)
  .replace(/engine:\s*'(\w+)'\s*또는\s*'(\w+)'/g, '--engine $1 또는 $2')
  .replace(/engine:\s*'(\w+)'/g, '--engine $1')
  .replace(/output을 (\.\w+)로/g, '-o 를 $1 파일로')
  .replace(/output으로 다른 파일에/g, '-o 로 다른 파일에');
function cliWording(e) {
  if (!(e instanceof Error)) return e;
  e.message = cliText(e.message);
  return e;
}

// 살펴보기 글의 안내(MCP 인수 표기 slide=n, offset=n)를 명령줄 옵션으로.
// 문서 글자는 건드리지 않도록 core가 붙이는 안내 문구에만 맞춘다
const INSPECT_HINTS = [
  [/\(slide=(\d+)로 이 슬라이드만 보면/g, '(--slide $1 로 이 슬라이드만 보면'], // PowerPoint 도형이 잘린 슬라이드
  [/ — 이어서 보려면 offset=(\d+)\)$/gm, ' — 이어서 보려면 --offset $1)'],
  [/ — offset은 (\d+) 이하로\)$/gm, ' — --offset 은 $1 이하로)'],
  [/\(목록은 offset=0일 때만 보여 줌\)/g, '(목록은 --offset 0일 때만 보여 줌)'], // Word 메모
  [/^문단: offset=(\d+)이 전체/gm, '문단: --offset $1 이 전체'],
  [/다음 문단은 offset=(\d+) 로 다시 살펴보세요/g, '다음 문단은 --offset $1 로 다시 살펴보세요'], // Word·한글
  [/ — limit을 줄이거나 offset을 옮겨 보세요$/gm, ' — --limit 을 줄이거나 --offset 을 옮겨 보세요'],
  [/ — 보려면 offset=\(문단 번호-1\)$/gm, ' — 보려면 --offset (문단 번호-1)'],
];
const cliInspectText = (t) => INSPECT_HINTS.reduce((s, [re, to]) => s.replace(re, to), t);

/** 한글 작업 확인 (MCP hwp_edit와 같은 스키마, 오류는 한국어) */
async function hwpOps(ops) {
  const [{ z }, { hwpOp }] = await Promise.all([import('zod'), hwpTools()]);
  const ko = z.locales.ko().localeError;
  const names = hwpOp.options.map((o) => o.shape.op.value);
  return ops.map((raw, i) => {
    const r = hwpOp.safeParse(raw, { error: (iss) => ko(iss) });
    if (r.success) return r.data;
    const name = typeof raw?.op === 'string' ? raw.op : null;
    const why = r.error.issues.map((iss) => (iss.code === 'invalid_union' && iss.discriminator === 'op'
      ? `알 수 없는 작업(op)입니다${name ? `: ${name}` : ''} (가능: ${names.join(', ')})`
      : `${iss.path.length ? `'${iss.path.join('.')}' ` : ''}${iss.message}`)).join('; ');
    throw new Error(`${i + 1}번째 작업${name ? `(${name})` : ''} 입력 오류: ${why} — 저장하지 않았습니다`);
  });
}

/** 살펴보기 → { format, data(원자료), text(한국어 글) } */
async function inspectDoc(full, opt) {
  const format = formatOf(full);
  checkEngine(format, opt.engine);
  const offset = intOpt(opt.offset, '--offset', 0);
  const limit = intOpt(opt.limit, '--limit', 1);
  const slide = intOpt(opt.slide, '--slide', 1);
  if (limit !== undefined && limit > 1000) throw new Error(`--limit 은 1000 이하여야 합니다: ${limit}`);
  const paging = offset !== undefined || limit !== undefined;
  if (paging && !['hwp', 'word', 'ppt'].includes(format)) {
    process.stderr.write('참고: --offset/--limit은 Word·한글(문단)·PowerPoint(슬라이드) 문서에만 씁니다\n');
  }
  if (slide !== undefined && format !== 'ppt') process.stderr.write('참고: --slide는 PowerPoint 문서에만 씁니다\n');
  if (slide !== undefined && paging && format === 'ppt') process.stderr.write('참고: --slide를 주면 --offset/--limit은 쓰지 않습니다\n');
  if (format === 'hwp') {
    const [{ Hwp }, { inspectText }] = await Promise.all([core(), hwpTools()]);
    const h = Hwp.open(full);
    try {
      const data = h.inspect({ offset: offset ?? 0, limit: limit ?? 150 });
      return { format, data, text: cliInspectText(inspectText(data, { name: basename(full) })) };
    } finally { h.close(); }
  }
  if (format === 'word') {
    const { wordInspect, formatWordInspect } = await import('./core/word.mjs');
    const data = await wordInspect(full, { offset: offset ?? 0, limit: limit ?? 200 });
    return { format, data, text: cliInspectText(formatWordInspect(data)) };
  }
  if (format === 'ppt') {
    // --offset: 건너뛸 슬라이드 수, --limit: 볼 슬라이드 수 (둘 다 없으면 모든 슬라이드)
    const { pptInspect, formatPptInspect } = await import('./core/ppt.mjs');
    const data = await pptInspect(full, slide !== undefined ? { slide } : { offset: offset ?? 0, limit });
    return { format, data, text: cliInspectText(formatPptInspect(data)) };
  }
  const { excelInspect, inspectText } = await import('./core/excel.mjs');
  const data = await excelInspect(full, { engine: opt.engine ?? 'auto' });
  return { format, data, text: inspectText(data) };
}

/** 편집 → { format, engine, output, lines(작업별 한 줄, 번호 붙음), notes, pages? } */
async function editDoc(src, out, ops, opt) {
  const format = formatOf(src);
  checkEngine(format, opt.engine);
  if (format === 'hwp') {
    if (!HWP_SAVE.has(extOfPath(out))) throw new Error(`한글 문서는 .hwp 또는 .hwpx로만 저장할 수 있습니다: ${basename(out)} (다른 형식은 convert)`);
    const valid = await hwpOps(ops);
    const [{ Hwp }, { applyHwpOp }] = await Promise.all([core(), hwpTools()]);
    const h = Hwp.open(src);
    try {
      const lines = [];
      for (const [i, op] of valid.entries()) {
        try { lines.push(`${i + 1}. ${applyHwpOp(h, op)}`); } catch (e) {
          throw new Error(`${i + 1}번째 작업(${op.op}) 실패: ${koreanError(e)} — 저장하지 않았습니다`);
        }
      }
      mkdirSync(dirname(out), { recursive: true });
      h.save(out);
      return { format, engine: 'rhwp', output: out, lines, notes: [], pages: h.pageCount };
    } finally { h.close(); }
  }
  mkdirSync(dirname(out), { recursive: true });
  if (format === 'word') {
    const { wordEdit } = await import('./core/word.mjs');
    const r = await wordEdit(src, ops, { output: out });
    return { format, engine: 'word', output: r.output, lines: r.lines.filter((l) => !l.startsWith('저장: ')), notes: [] };
  }
  if (format === 'ppt') {
    const { pptEdit } = await import('./core/ppt.mjs');
    const r = await pptEdit(src, ops, { output: out });
    // 참고 글(예: 쓰기 암호 유지)은 '참고: '를 붙여 출력하므로 core가 붙인 머리는 뗀다
    const notes = (r.notes ?? []).map((n) => cliText(String(n).replace(/^참고:\s*/, '')));
    return { format, engine: 'powerpoint', output: r.output, lines: r.results, notes };
  }
  const { excelEdit } = await import('./core/excel.mjs');
  const r = await excelEdit(src, ops, { output: out, engine: opt.engine ?? 'auto' });
  return { format, engine: r.engine, output: r.output, lines: r.results.map((l, i) => `${i + 1}. ${l}`), notes: r.notes ?? [] };
}

/** office-kit help edit */
async function editHelp() {
  const [{ WORD_OPS }, { OP_NAMES }, { EXCEL_OPS }, { hwpOp }] = await Promise.all([
    import('./core/word.mjs'), import('./core/ppt.mjs'), import('./core/excel.mjs'), hwpTools()]);
  const wrap = (names) => {
    const out = []; let line = '';
    for (const n of names) {
      if (line && line.length + n.length + 1 > 96) { out.push(line); line = ''; }
      line = line ? `${line} ${n}` : n;
    }
    if (line) out.push(line);
    return out.map((l) => `    ${l}`).join('\n');
  };
  const excel = Object.entries(EXCEL_OPS).map(([k, v]) => (v.file ? k : `${k}×`));
  return `형식별 편집: office-kit edit <파일> <작업.json> [-o 출력] [--overwrite] [--engine auto|app|file]

작업.json — 순서대로 적용할 작업 배열 (또는 {"operations": [...]}). JSON 문자열을 바로 써도 된다.
  [{"op":"insert_paragraph","text":"맺음말","at":{"after_paragraph":3},"style":"heading2"},
   {"op":"format_text","find":"중요","all":true,"bold":true,"color":"#c00000"}]

- 모두 적용한 뒤 한 번만 저장한다. 하나라도 실패하면 아무것도 저장하지 않는다 (몇 번째 작업이 왜 실패했는지 알려 줌).
- 번호(문단·표·슬라이드·행·열)는 1부터. office-kit inspect 결과의 번호를 쓴다.
  앞 작업이 문단·슬라이드를 넣거나 지우면 뒤 작업은 바뀐 번호를 쓴다 (결과 줄에 새 번호가 나온다).
- 위치 at: "end"(기본) | "start" | {"after_paragraph":n} | {"before_paragraph":n} | {"after_text":"찾을 글"}
- find: 대소문자 구분, 첫 번째만 ("all": true면 전부). 그림·테마 경로는 지금 폴더 기준.
- -o 를 생략하면 원본에 저장한다. -o 경로에 다른 파일이 있으면 --overwrite 를 붙여야 덮어쓴다.

Word (docx doc rtf odt) — 설치된 Word 필요:
${wrap(WORD_OPS)}
PowerPoint (pptx ppt odp) — 설치된 PowerPoint 필요:
${wrap(OP_NAMES)}
Excel (xlsx xlsm xls csv ods) — Excel 우선, 없으면 파일 엔진(xlsx·csv만, ×는 Excel 필요):
${wrap(excel)}
한글 (hwp hwpx hwt) — 내장 엔진(rhwp), 한컴오피스 없이:
${wrap(hwpOp.options.map((o) => o.shape.op.value))}

작업마다 인수: ${resolve(APP_ROOT, 'docs', 'FEATURES.md')}`;
}

// ── 명령 ─────────────────────────────────────────────────

const yes = (v) => (v ? '있음' : '없음');

const COMMANDS = {
  async status() {
    const { probe, hwpPdfStatus } = await core();
    const p = await probe();
    print(`office-kit ${VERSION}`);
    if (p.error) print(`주의: 설치된 프로그램을 확인하지 못했습니다 (${p.error}). 프로그램 없이 되는 작업만 합니다.`);
    print(`설치된 프로그램: 한컴오피스 ${yes(p.hancom)}, Word ${yes(p.word)}, PowerPoint ${yes(p.powerpoint)}, Excel ${yes(p.excel)}`);
    print(`한글 자동화 보안 모듈: ${p.hancomSecurityModule ? '등록됨' : '등록 안 됨'}`
      + (p.hancom && !p.hancomSecurityModule ? ' (한글 문서는 rhwp로 처리 — 한컴으로 처리하려면 office-kit setup hancom)' : ''));
    const pdf = hwpPdfStatus();
    print(`한글→PDF: ${p.hancom && p.hancomSecurityModule ? '한컴오피스' : pdf.ok ? '내장 엔진(rhwp, 한컴오피스 없이)' : `불가 — ${pdf.reason}`}`);
    try {
      const h = await (await setupModule('hancom.mjs')).hancomSecurityStatus();
      if (h.registered) print(`  보안 모듈 파일: ${h.dllPath}${h.dllExists ? '' : ' (파일 없음)'}`);
      if (h.warning) print(`  주의: ${h.warning}`);
    } catch { /* 설정 모듈이 없으면 생략 */ }
    try {
      const m = await setupModule('integrations.mjs');
      const s = await m.integrationStatus(m.mcpLaunchSpec({ execPath: process.execPath, appRoot: APP_ROOT }));
      // 상태 객체: { connected, matches, available, error, detail } (앱 설정 화면과 같은 기준)
      const on = (v) => {
        if (!v || typeof v !== 'object') return v ? '연결됨' : '안 됨';
        if (v.error) return `확인 필요 (${v.error})`;
        if (v.available === false) return '프로그램 없음';
        const connected = v.connected ?? v.configured ?? v.installed ?? (v.exists && v.ours);
        if (connected && v.matches === false) return '연결됨(다른 경로 — 다시 연결 필요)';
        if (connected) return '연결됨';
        return v.exists && v.detail ? `안 됨 (${v.detail})` : '안 됨';
      };
      print(`연결: Claude Desktop ${on(s.claudeDesktop)}, Claude Code ${on(s.claudeCode)}, Codex ${on(s.codex)}, 명령줄 ${on(s.cli)}`);
    } catch { /* 설정 모듈이 없으면 연결 상태는 생략 */ }
  },

  async info([file], opt) {
    need([file], 1, 'info <파일>');
    const { documentInfo } = await core();
    const i = await documentInfo(fileArg(file));
    if (opt.json) return print(JSON.stringify(i, null, 2));
    print(`이름: ${i.name}`);
    print(`종류: ${i.kind} (.${i.format}), ${i.size.toLocaleString()}바이트`);
    if (i.pages !== undefined) print(`쪽: ${i.pages}, 표: ${i.tables}, 누름틀: ${i.fields?.length ? i.fields.join(', ') : '없음'}`);
    print(`프로그램 엔진: ${i.appAvailable ? '가능' : '불가'}, 파일 엔진: ${i.fileEngine ? '가능' : '불가'}`);
  },

  async text([file], opt) {
    need([file], 1, 'text <파일> [--out 파일]');
    const { readDocument } = await core();
    const doc = await readDocument(fileArg(file), { engine: opt.engine });
    if (opt.out) { writeFileSync(outOf(opt), doc.text, 'utf8'); print(`저장: ${resolve(opt.out)} (${doc.text.length}자)`); }
    else print(doc.text);
  },

  async fields([file], opt) {
    need([file], 1, 'fields <파일>');
    const { listFields } = await core();
    const fields = await listFields(fileArg(file), { engine: opt.engine });
    if (opt.json) return print(JSON.stringify(fields, null, 2));
    if (!fields.length) return print('채울 칸이 없습니다.');
    for (const f of fields) print(`${f.name}\t[${f.kind ?? ''}]${f.value ? `\t${f.value}` : ''}`);
  },

  async fill([file, values], opt) {
    need([file, values].filter(Boolean), 2, 'fill <파일> <값.json> [-o 출력]');
    const { fillFields } = await core();
    const src = fileArg(file);
    const vals = Object.fromEntries(Object.entries(jsonArg(values, '값')).map(([k, v]) => [k, String(v)]));
    const r = await fillFields(src, vals, { output: outOf(opt, src), engine: opt.engine });
    print(`채움: ${r.filled.join(', ') || '없음'}`);
    if (r.missing.length) print(`못 찾음: ${r.missing.join(', ')}`);
    print(`저장: ${r.output}`);
  },

  async replace([file, find, repl], opt) {
    need([file, find, repl].filter((v) => v !== undefined), 3, 'replace <파일> <찾을말> <바꿀말> [-o 출력]');
    if (!find) throw new Error('찾을 말이 비어 있습니다');
    const { replaceText } = await core();
    const src = fileArg(file);
    const r = await replaceText(src, [[find, repl]], { output: outOf(opt, src), engine: opt.engine });
    print(`${r.count}곳 바꿈 (${r.engine})`);
    print(`저장: ${r.output}`);
  },

  async cells([file, cells], opt) {
    need([file, cells].filter(Boolean), 2, 'cells <파일> <셀.json> [--sheet 시트] [-o 출력]');
    const { setCells } = await core();
    const src = fileArg(file);
    const r = await setCells(src, jsonArg(cells, '셀'), { sheet: opt.sheet, output: outOf(opt, src), engine: opt.engine });
    print(`셀 ${r.updated}개 씀 (${r.engine})`);
    print(`저장: ${r.output}`);
  },

  async convert([input, output]) {
    need([input, output].filter(Boolean), 2, 'convert <입력> <출력>');
    const { convertDocument } = await core();
    const src = fileArg(input);
    const out = resolve(output);
    if (src.toLowerCase() === out.toLowerCase()) throw new Error('입력과 다른 출력 경로를 지정하세요');
    mkdirSync(dirname(out), { recursive: true });
    const r = await convertDocument(src, out);
    print(`변환 완료 (${ENGINE_NAME[r.engine] ?? r.engine}${r.pages ? `, ${r.pages}쪽` : ''}): ${r.output}`);
    if (r.note) print(`참고: ${r.note}`);
  },

  async new([manuscript, output], opt) {
    need([manuscript, output].filter(Boolean), 2, 'new <원고.md> <출력>');
    const { createDocument } = await core();
    const md = readFileSync(fileArg(manuscript), 'utf8');
    const out = resolve(output);
    if (existsSync(out) && !opt.force) throw new Error(`이미 있는 파일입니다: ${out} (덮어쓰려면 --force)`);
    mkdirSync(dirname(out), { recursive: true });
    await createDocument(out, md);
    print(`만들었습니다: ${out}`);
  },

  async image([file, page = '1'], opt) {
    need([file].filter(Boolean), 1, 'image <파일> [쪽] [-o 그림.png]');
    if (!/^\d+$/.test(page)) throw new Error(`쪽 번호는 1 이상의 정수여야 합니다: ${page}`);
    const { pageImage } = await import('./core/render.mjs');
    const src = fileArg(file);
    const r = await pageImage(src, Number(page), { dpi: opt.dpi ? Number(opt.dpi) : undefined });
    const out = outOf(opt) ?? resolve(`${basename(src, extname(src))}-p${r.page}.png`);
    writeFileSync(out, r.png);
    print(`${r.page}/${r.pages}쪽 저장 (${r.engine}${r.width ? `, ${r.width}×${r.height}px` : ''}): ${out}`);
  },

  async inspect([file], opt) {
    need([file].filter(Boolean), 1, 'inspect <파일> [--offset n] [--limit n] [--slide n] [--json]');
    const r = await inspectDoc(fileArg(file), opt).catch((e) => { throw cliWording(e); });
    print(opt.json ? JSON.stringify(r.data, null, 2) : r.text);
  },

  async edit([file, ops], opt) {
    need([file, ops].filter(Boolean), 2, 'edit <파일> <작업.json> [-o 출력] [--overwrite]');
    const src = fileArg(file);
    formatOf(src);
    const list = opsArg(ops);
    let out = src;
    if (opt.out) {
      out = resolve(opt.out);
      if (samePath(out, src)) out = src;
      else if (existsSync(out) && !opt.overwrite) throw new Error(`이미 있는 파일입니다: ${out} (덮어쓰려면 --overwrite)`);
    }
    const r = await editDoc(src, out, list, opt).catch((e) => { throw cliWording(e); });
    if (opt.json) {
      return print(JSON.stringify({ format: r.format, engine: r.engine, output: r.output, results: r.lines, notes: r.notes,
        ...(r.pages ? { pages: r.pages } : {}) }, null, 2));
    }
    for (const l of r.lines) print(l);
    for (const n of r.notes) print(`참고: ${n}`);
    print(`저장: ${r.output} (${ENGINE_NAME[r.engine] ?? r.engine}${r.pages ? `, ${r.pages}쪽` : ''})`);
  },

  async mcp() {
    const { startServer } = await import('./mcp/server.mjs');
    await startServer();
    return 'keep-alive';
  },

  async setup([target], opt) { await setup(target, opt); },

  async help([topic]) {
    if (topic === 'edit' || topic === 'inspect') return print(await editHelp());
    print(HELP);
  },
};

export async function main(argv = process.argv.slice(2)) {
  try {
    const { pos, opt } = parseArgs(argv);
    if (opt.version) { print(VERSION); return 0; }
    const [cmd = 'help', ...rest] = pos;
    if (opt.help || cmd === '--help') { print(HELP); return 0; }
    const run = Object.hasOwn(COMMANDS, cmd) ? COMMANDS[cmd] : null;
    if (!run) throw new Error(`알 수 없는 명령: ${cmd} (도움말: office-kit help)`);
    await run(rest, opt);
    return 0;
  } catch (e) {
    process.stderr.write(`오류: ${koreanError(e)}\n`);
    if (process.env.OFFICEKIT_DEBUG) process.stderr.write(`${e?.stack ?? ''}\n`);
    return 1;
  }
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
  const code = await main();
  if (code) process.exitCode = code;
}
