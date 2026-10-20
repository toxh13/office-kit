// pii.mjs — 문서에서 개인정보(주민등록번호·전화번호·이메일·카드·계좌·여권·운전면허)를 찾아 가린 사본을 만든다.
//   인터넷·AI 없이 글자 규칙으로 찾는다. 원본은 바꾸지 않고 '<이름>_가림.<확장자>'에 쓴 뒤,
//   가린 사본을 다시 읽어 남은 것이 없는지 확인한다 (글자가 여러 조각으로 나뉘어 저장돼 못 바꾼 것 등).
//   찾기는 완벽하지 않다 (사진·그림 속 글자, 띄어 쓴 번호 등) — 화면·보고서에 그렇게 적는다.
import { existsSync } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { readDocument, replaceText, extOf } from './index.mjs';

/** 찾는 종류 (화면 순서) */
export const PII_TYPES = {
  rrn: '주민등록번호',
  frn: '외국인등록번호',
  mobile: '휴대전화',
  phone: '전화번호',
  email: '이메일',
  card: '카드번호',
  account: '계좌번호',
  passport: '여권번호',
  license: '운전면허번호',
};
export const PII_DEFAULT_TYPES = Object.keys(PII_TYPES);

const digits = (s) => s.replace(/\D/g, '');
/** 숫자만 *로 (구분 기호는 그대로) — 앞 keep자리와 뒤 tail자리 숫자는 남긴다 */
function maskDigits(s, keep, tail = 0) {
  const total = digits(s).length;
  let n = 0;
  return s.replace(/\d/g, (d) => { n++; return n <= keep || n > total - tail ? d : '*'; });
}

function luhn(num) {
  let sum = 0; let alt = false;
  for (let i = num.length - 1; i >= 0; i--) {
    let d = Number(num[i]);
    if (alt) { d *= 2; if (d > 9) d -= 9; }
    sum += d; alt = !alt;
  }
  return sum % 10 === 0;
}

// 주민등록번호 검증 숫자 (2020년 10월 이전 번호). 하이픈 없는 13자리 숫자를 다른 번호와 가르는 데만 쓴다
function rrnChecksum(d) {
  if (d.length !== 13) return false;
  const w = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5];
  const sum = w.reduce((n, x, i) => n + x * Number(d[i]), 0);
  return (11 - (sum % 11)) % 10 === Number(d[12]);
}

const near = (text, index, re, span = 24) => re.test(text.slice(Math.max(0, index - span), index));
const ACCOUNT_WORDS = /(계좌|통장|은행|예금주|입금|송금|account)/i;

// 각 규칙: re(전역), type(값 → 종류|null), mask(값 → 가린 글)
const RULES = [
  { // 주민·외국인 등록번호: 생년월일 6자리 - 성별 1자리 + 6자리. 하이픈 없이 13자리면 검증 숫자가 맞을 때만
    re: /(?<![\d-])(\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(\s?[-–]\s?)?([1-8])\d{6}(?![\d-])/g,
    type: (m) => (!m[4] && !rrnChecksum(digits(m[0])) ? null : Number(m[5]) >= 5 ? 'frn' : 'rrn'),
    mask: (v) => maskDigits(v, 7),
  },
  { // 휴대전화 010-1234-5678 / 01012345678 / 010 1234 5678
    re: /(?<![\d-])01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}(?![\d-])/g,
    type: () => 'mobile',
    mask: (v) => maskMiddleGroup(v),
  },
  { // 유선 전화 02-123-4567, 031-123-4567, 070-1234-5678 (구분 기호가 있을 때만)
    re: /(?<![\d-])(?:\(0(?:2|[3-6][1-5]|70)\)\s?|0(?:2|[3-6][1-5]|70)[-.\s])\d{3,4}[-.\s]\d{4}(?![\d-])/g,
    type: () => 'phone',
    mask: (v) => maskMiddleGroup(v),
  },
  { // 이메일
    re: /(?<![\w.+-])[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9.-]+\.[A-Za-z]{2,}(?![\w-])/g,
    type: () => 'email',
    mask: (v) => { const [id, host] = v.split('@'); return `${id.slice(0, Math.min(2, id.length - 1 || 1))}${'*'.repeat(Math.max(3, id.length - 2))}@${host}`; },
  },
  { // 카드번호 1234-5678-9012-3456 (Luhn 검사)
    re: /(?<![\d-])(?:\d{4}[-\s]){3}\d{4}(?![\d-])|(?<!\d)\d{16}(?!\d)/g,
    type: (m) => (luhn(digits(m[0])) ? 'card' : null),
    mask: (v) => maskDigits(v, 4, 4),
  },
  { // 여권번호 M12345678 / M123A4567
    re: /(?<![A-Za-z0-9])[MSRODG](?:\d{8}|\d{3}[A-Z]\d{4})(?![A-Za-z0-9])/g,
    type: () => 'passport',
    mask: (v) => `${v.slice(0, 1)}${'*'.repeat(v.length - 3)}${v.slice(-2)}`,
  },
  { // 운전면허 11-12-123456-12
    re: /(?<![\d-])(?:1[1-9]|2[0-8])-\d{2}-\d{6}-\d{2}(?![\d-])/g,
    type: () => 'license',
    mask: (v) => maskDigits(v, 4, 2),
  },
  { // 계좌번호: 숫자 묶음 3~5개(하이픈), 합쳐 10~16자리, 앞에 '계좌·은행·입금' 같은 말이 있을 때만
    re: /(?<![\d-])\d{2,6}(?:-\d{2,7}){2,4}(?![\d-])/g,
    type: (m, text) => { const n = digits(m[0]).length; return n >= 10 && n <= 16 && near(text, m.index, ACCOUNT_WORDS, 30) ? 'account' : null; },
    mask: (v) => maskDigits(v, 3, 3),
  },
];

/** 가운데 묶음만 가리기 (010-1234-5678 → 010-****-5678, 01012345678 → 010****5678) */
function maskMiddleGroup(v) {
  const parts = v.split(/([-.\s)]+)/);
  const groups = parts.filter((p, i) => i % 2 === 0 && /\d/.test(p));
  if (groups.length >= 3) {
    let gi = 0;
    return parts.map((p, i) => (i % 2 === 0 && /\d/.test(p) ? (gi++ === groups.length - 2 ? p.replace(/\d/g, '*') : p) : p)).join('');
  }
  return maskDigits(v, 3, 4);
}

/**
 * 글에서 개인정보 찾기 → [{ type, label, value, masked, index }] (앞에서부터, 겹치면 먼저 찾은 규칙이 이긴다)
 * types: 찾을 종류 (기본: 모두)
 */
export function findPii(text, { types = PII_DEFAULT_TYPES } = {}) {
  text = String(text ?? '');
  const want = new Set(types);
  const taken = [];
  const out = [];
  const overlaps = (s, e) => taken.some(([a, b]) => s < b && e > a);
  for (const rule of RULES) {
    for (const m of text.matchAll(rule.re)) {
      const type = rule.type(m, text);
      if (!type || !want.has(type)) continue;
      const s = m.index; const e = s + m[0].length;
      if (overlaps(s, e)) continue;
      taken.push([s, e]);
      out.push({ type, label: PII_TYPES[type], value: m[0], masked: rule.mask(m[0]), index: s });
    }
  }
  return out.sort((a, b) => a.index - b.index);
}

/** 같은 값끼리 묶기 → [{ type, label, value, masked, count, context }] */
export function groupPii(found, text = '') {
  const map = new Map();
  for (const f of found) {
    const k = `${f.type}\u0000${f.value}`;
    const g = map.get(k);
    if (g) { g.count++; continue; }
    const s = Math.max(0, f.index - 12); const e = Math.min(text.length, f.index + f.value.length + 12);
    const context = text ? `${s > 0 ? '…' : ''}${text.slice(s, e).replace(/\s+/g, ' ')}${e < text.length ? '…' : ''}` : '';
    map.set(k, { type: f.type, label: f.label, value: f.value, masked: f.masked, count: 1, context });
  }
  return [...map.values()];
}

async function documentText(path) {
  const r = await readDocument(path);
  // 시트는 칸 사이를 탭으로 (한 칸에 든 번호가 옆 칸과 붙지 않게)
  return String(r.text ?? '');
}

/** 문서에서 찾기 → { items: groupPii 결과, total, text길이 } */
export async function scanDocument(path, { types } = {}) {
  path = resolve(path);
  const text = await documentText(path);
  const found = findPii(text, { types });
  return { path, items: groupPii(found, text), total: found.length, chars: text.length };
}

/** 가린 사본 기본 이름: <이름>_가림.<확장자> (있으면 _가림 (2) …) */
export function maskedPathFor(path) {
  const dir = dirname(path); const ext = extname(path); const stem = basename(path, ext);
  let out = join(dir, `${stem}_가림${ext}`);
  for (let i = 2; existsSync(out); i++) out = join(dir, `${stem}_가림 (${i})${ext}`);
  return out;
}

const FILE_ENGINE_EXTS = new Set(['hwp', 'hwpx', 'hwt', 'docx', 'pptx', 'xlsx']);

/**
 * 가린 사본 만들기. values: 가릴 값 목록 (생략하면 types의 모든 것). output: 결과 경로 (원본과 달라야 함)
 * → { output, replaced, items: [{ value, masked, count }], remaining: [남은 값 groupPii], engine }
 */
export async function maskDocument(path, { output, types, values } = {}) {
  path = resolve(path);
  output = resolve(output ?? maskedPathFor(path));
  if (output.toLowerCase() === path.toLowerCase()) throw new Error('원본에 바로 덮어쓰지 않아요. 다른 이름으로 저장하세요');
  if (extOf(output) !== extOf(path)) throw new Error(`가린 사본은 원본과 같은 형식(.${extOf(path)})으로 저장해요`);
  const scan = await scanDocument(path, { types });
  const pick = values ? scan.items.filter((it) => values.includes(it.value)) : scan.items;
  if (!pick.length) throw new Error('가릴 개인정보가 없어요');
  // 긴 값부터 바꾼다 (짧은 값이 긴 값의 일부일 때 먼저 바뀌어 긴 값을 못 찾지 않게)
  const pairs = [...pick].sort((a, b) => b.value.length - a.value.length).map((it) => [it.value, it.masked]);
  const r = await replaceText(path, pairs, { output, engine: FILE_ENGINE_EXTS.has(extOf(path)) ? 'file' : 'auto' });
  const counts = new Map(pairs.map(([v], i) => [v, r.counts?.[i] ?? 0]));
  // 확인: 가린 사본을 다시 읽어 고른 값이 남았는지
  const after = await scanDocument(output, { types: [...new Set(pick.map((it) => it.type))] });
  const chosen = new Set(pick.map((it) => it.value));
  const remaining = after.items.filter((it) => chosen.has(it.value));
  return {
    output, engine: r.engine, replaced: r.count ?? 0,
    items: pick.map((it) => ({ type: it.type, label: it.label, value: it.value, masked: it.masked, count: counts.get(it.value) ?? 0 })),
    remaining,
  };
}
