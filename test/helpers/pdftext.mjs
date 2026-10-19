// pdftext.mjs — 시험용 작은 PDF 읽개: 쪽 수·쪽 크기·글자(ToUnicode로 푼 것)·글꼴.
//   rhwp(svg2pdf/pdf-writer)가 만드는 모양만 다룬다: 압축 안 한 객체 표, FlateDecode 스트림, Type0(Identity-H) 글꼴,
//   쪽 → Form XObject(Do) 안에 글자. 글자 사이 띄어쓰기는 PDF에 없을 수 있으므로 비교할 때는 공백을 빼고 본다.
import { inflateSync } from 'node:zlib';

function parseObjects(buf) {
  const s = buf.toString('latin1');
  const objs = new Map();
  const re = /(\d+) 0 obj\s*/g;
  let m;
  while ((m = re.exec(s))) {
    const id = Number(m[1]);
    const start = re.lastIndex;
    // 사전 끝 찾기 (<< >> 짝)
    let i = start; let depth = 0; let dict = '';
    if (s.startsWith('<<', i)) {
      for (; i < s.length; i++) {
        if (s.startsWith('<<', i)) { depth++; i++; continue; }
        if (s.startsWith('>>', i)) { depth--; i++; if (depth === 0) { i++; break; } continue; }
      }
      dict = s.slice(start, i);
    } else {
      const end = s.indexOf('endobj', start);
      objs.set(id, { dict: s.slice(start, end).trim(), stream: null });
      re.lastIndex = end;
      continue;
    }
    let stream = null;
    const after = s.slice(i, i + 20).match(/^\s*stream\r?\n/);
    if (after) {
      const dataStart = i + after[0].length;
      let len = dict.match(/\/Length (\d+)(?! \d+ R)/)?.[1];
      const ref = dict.match(/\/Length (\d+) 0 R/)?.[1];
      if (!len && ref) len = objs.get(Number(ref))?.dict;
      len = Number(len);
      if (!(len >= 0)) len = s.indexOf('endstream', dataStart) - dataStart;
      stream = buf.subarray(dataStart, dataStart + len);
      re.lastIndex = dataStart + len;
    } else re.lastIndex = i;
    objs.set(id, { dict, stream });
  }
  return objs;
}

const ref = (d, key) => Number(d.match(new RegExp(`/${key}\\s*(\\d+) 0 R`))?.[1]);
const data = (o) => (o?.stream ? (/\/FlateDecode/.test(o.dict) ? inflateSync(o.stream) : o.stream) : Buffer.alloc(0));

// "<< /Font << /fo0 8 0 R >> /XObject << /xo0 6 0 R >> >>" 안의 이름 → 객체 번호
function subDict(d, key, objs) {
  let m = d.match(new RegExp(`/${key}\\s*<<([^<>]*)>>`));
  let body = m?.[1];
  if (!body) {
    const r = ref(d, key);
    if (r) body = objs.get(r)?.dict.replace(/^<<|>>$/g, '');
  }
  const out = new Map();
  for (const [, name, id] of (body ?? '').matchAll(/\/([^\s/<>\[\]]+)\s*(\d+) 0 R/g)) out.set(name, Number(id));
  return out;
}

function resourcesOf(d, objs) {
  const inline = d.match(/\/Resources\s*(<<[\s\S]*>>)/)?.[1];
  if (inline) return inline;
  const r = ref(d, 'Resources');
  return r ? objs.get(r)?.dict ?? '' : '';
}

function hex(u) { return String.fromCodePoint(...(u.match(/.{4}/g) ?? []).map((h) => parseInt(h, 16))); }

function cmapOf(o) {
  const text = data(o).toString('latin1');
  const map = new Map();
  for (const block of text.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const [, a, b] of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) map.set(parseInt(a, 16), hex(b));
  }
  for (const block of text.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const [, a, b, c] of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) {
      const lo = parseInt(a, 16); const hi = parseInt(b, 16); const base = parseInt(c, 16);
      for (let k = lo; k <= hi; k++) map.set(k, String.fromCodePoint(base + k - lo));
    }
  }
  return map;
}

/** PDF 읽기 → { pages: [{ width, height, text }], fonts: [BaseFont...], text } */
export function readPdf(buf) {
  const objs = parseObjects(buf);
  const fonts = new Set();
  const cmaps = new Map();
  const fontInfo = (id) => {
    if (cmaps.has(id)) return cmaps.get(id);
    const d = objs.get(id)?.dict ?? '';
    const base = d.match(/\/BaseFont\s*\/([^\s/<>\[\]]+)/)?.[1];
    if (base) fonts.add(base);
    const tu = ref(d, 'ToUnicode');
    const info = { twoByte: /\/Type0/.test(d), map: tu ? cmapOf(objs.get(tu)) : new Map() };
    cmaps.set(id, info);
    return info;
  };
  const decode = (font, bytes) => {
    if (!font) return '';
    let out = '';
    const step = font.twoByte ? 2 : 1;
    for (let i = 0; i + step <= bytes.length; i += step) {
      const code = step === 2 ? (bytes[i] << 8) | bytes[i + 1] : bytes[i];
      out += font.map.get(code) ?? '';
    }
    return out;
  };
  const strBytes = (tok) => {
    if (tok.startsWith('<')) return Buffer.from(tok.slice(1, -1).replace(/\s/g, ''), 'hex');
    // (리터럴) — 이스케이프 몇 가지만
    const raw = tok.slice(1, -1).replace(/\\([nrtbf()\\]|[0-7]{1,3})/g, (_m, e) => (/^[0-7]/.test(e) ? String.fromCharCode(parseInt(e, 8))
      : ({ n: '\n', r: '\r', t: '\t', b: '\b', f: '\f' }[e] ?? e)));
    return Buffer.from(raw, 'latin1');
  };
  const run = (content, res, depth = 0) => {
    if (depth > 8) return '';
    const fontsHere = subDict(res, 'Font', objs);
    const xobjs = subDict(res, 'XObject', objs);
    let font = null; let out = '';
    const tokens = content.toString('latin1').match(/<[0-9a-fA-F\s]*>|\((?:\\.|[^\\)])*\)|\[|\]|\/[^\s/<>\[\]()]+|[^\s/<>\[\]()]+/g) ?? [];
    const stack = [];
    for (const t of tokens) {
      if (t === 'Tf') { const name = stack[stack.length - 2]?.slice(1); const id = fontsHere.get(name); font = id ? fontInfo(id) : null; stack.length = 0; continue; }
      if (t === 'Tj' || t === "'" || t === '"') { const s = stack.pop(); if (s && /^[<(]/.test(s)) out += decode(font, strBytes(s)); stack.length = 0; continue; }
      if (t === 'TJ') { out += stack.filter((x) => /^[<(]/.test(x)).map((x) => decode(font, strBytes(x))).join(''); stack.length = 0; continue; }
      if (t === 'ET' || t === 'T*' || t === 'Td' || t === 'TD') { if (t === 'ET') out += '\n'; stack.length = 0; continue; }
      if (t === 'Do') {
        const id = xobjs.get(stack.pop()?.slice(1));
        const o = id ? objs.get(id) : null;
        if (o && /\/Subtype\s*\/Form/.test(o.dict)) out += run(data(o), resourcesOf(o.dict, objs), depth + 1);
        stack.length = 0; continue;
      }
      if (t === '[' || t === ']') { if (t === '[') stack.length = 0; continue; }
      stack.push(t);
    }
    return out;
  };
  // 쪽 순서: /Pages의 /Kids를 따라간다
  const catalog = [...objs.values()].find((o) => /\/Type\s*\/Catalog/.test(o.dict));
  const pageIds = [];
  const walk = (id) => {
    const d = objs.get(id)?.dict ?? '';
    if (/\/Type\s*\/Pages/.test(d)) {
      for (const [, k] of (d.match(/\/Kids\s*\[([^\]]*)\]/)?.[1] ?? '').matchAll(/(\d+) 0 R/g)) walk(Number(k));
    } else if (/\/Type\s*\/Page\b/.test(d)) pageIds.push(id);
  };
  walk(ref(catalog?.dict ?? '', 'Pages'));
  const pages = pageIds.map((id) => {
    const d = objs.get(id).dict;
    const box = (d.match(/\/MediaBox\s*\[([^\]]*)\]/)?.[1] ?? '').trim().split(/\s+/).map(Number);
    const res = resourcesOf(d, objs);
    const contents = [];
    const one = ref(d, 'Contents');
    if (one) contents.push(one);
    else for (const [, k] of (d.match(/\/Contents\s*\[([^\]]*)\]/)?.[1] ?? '').matchAll(/(\d+) 0 R/g)) contents.push(Number(k));
    const text = contents.map((c) => run(data(objs.get(c)), res)).join('');
    return { width: box[2] - box[0], height: box[3] - box[1], text };
  });
  return { pages, fonts: [...fonts], text: pages.map((p) => p.text).join('\f') };
}

/** 공백을 뺀 글자 (띄어쓰기가 없는 PDF와 비교하려고) */
export const squash = (s) => String(s).replace(/\s+/g, '');
