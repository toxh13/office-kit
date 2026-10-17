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
  new <원고.md> <출력> [--force]      원고로 새 문서 만들기 (.hwp .hwpx .docx .pptx .xlsx)
  image <파일> [쪽] [-o 그림.png] [--dpi 110]
                                      한 쪽을 PNG 그림으로 저장
  mcp                                 MCP 서버 실행 (stdio, Claude/Codex 연결용)
  setup <대상> [--remove]             연결 설정. 대상: claude-desktop | claude-code | codex | cli | hancom
                                      (hancom: --dll <경로> 로 보안 모듈 파일 지정, --download 로 한컴 공식 배포본 받기)
  help                                이 도움말

공통 옵션:
  -o, --out <경로>    결과를 저장할 경로 (생략하면 원본을 덮어씀)
  --engine <방식>     auto(기본) | app(설치된 프로그램) | file(프로그램 없이 파일 직접 처리)

원고 문법 (new):
  # 제목 / ## 소제목 / ### 작은 제목    >> 가운데 줄    **굵은 한 줄**    - 글머리표
  | 칸 | 칸 | 표 (|---| 줄 무시)         라벨: {{필드}}  (한글은 누름틀, 그 외는 자리표시자)
  ---page---  쪽 나누기 (pptx는 새 슬라이드)

예:
  office-kit text 보고서.hwp
  office-kit fill 신청서.hwpx 값.json -o 신청서_채움.hwpx
  office-kit replace 계약서.docx 갑 을
  office-kit convert 보고서.hwp 보고서.hwpx
  office-kit image 보고서.hwp 2 -o 2쪽.png
`;

// ── 인수 해석 ────────────────────────────────────────────

const VALUED = { '-o': 'out', '--out': 'out', '--output': 'out', '--sheet': 'sheet', '--engine': 'engine', '--dpi': 'dpi', '--dll': 'dll' };
const FLAGS = { '--remove': 'remove', '--download': 'download', '--json': 'json', '--force': 'force',
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

/** JSON 파일 경로 또는 JSON 문자열 → 객체 */
function jsonArg(arg, what) {
  let src = arg;
  if (!/^\s*[{[]/.test(arg)) {
    const full = resolve(arg);
    if (!existsSync(full)) throw new Error(`${what} 파일이 없습니다: ${full}`);
    src = readFileSync(full, 'utf8').replace(/^﻿/, '');
  }
  let v;
  try { v = JSON.parse(src); } catch (e) { throw new Error(`${what}을(를) JSON으로 읽을 수 없습니다: ${e.message}`); }
  if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Error(`${what}은(는) {"이름": 값} 형태의 JSON 객체여야 합니다`);
  return v;
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

// ── 명령 ─────────────────────────────────────────────────

const yes = (v) => (v ? '있음' : '없음');

const COMMANDS = {
  async status() {
    const { probe } = await core();
    const p = await probe();
    print(`office-kit ${VERSION}`);
    if (p.error) print(`주의: 설치된 프로그램을 확인하지 못했습니다 (${p.error}). 프로그램 없이 되는 작업만 합니다.`);
    print(`설치된 프로그램: 한컴오피스 ${yes(p.hancom)}, Word ${yes(p.word)}, PowerPoint ${yes(p.powerpoint)}, Excel ${yes(p.excel)}`);
    print(`한글 자동화 보안 모듈: ${p.hancomSecurityModule ? '등록됨' : '등록 안 됨'}`
      + (p.hancom && !p.hancomSecurityModule ? ' (한글 문서는 rhwp로 처리, PDF 변환 불가 — office-kit setup hancom)' : ''));
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
    print(`변환 완료 (${r.engine}): ${r.output}`);
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

  async mcp() {
    const { startServer } = await import('./mcp/server.mjs');
    await startServer();
    return 'keep-alive';
  },

  async setup([target], opt) { await setup(target, opt); },

  async help() { print(HELP); },
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
