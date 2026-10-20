// diff.mjs — 두 문서(같은 문서의 옛 버전과 지금)에서 바뀐 글을 찾는다. '변경 내용 보기'용.
//   한글·Word: 문단(줄) 단위, PowerPoint: 슬라이드마다 줄 단위, Excel·CSV: 칸 단위.
//   글만 비교한다 (글꼴·색·그림·서식 변화는 나오지 않는다). 프로그램이 필요 없는 형식은 인터넷·프로그램 없이 된다.
import { readDocument, extOf } from './index.mjs';

const MAX_D = 4000;        // 이보다 많이 다르면 가운데를 통째로 바뀐 것으로 본다 (시간·메모리 한도)
const MAX_ITEMS = 400;     // 화면에 보일 바뀐 줄·칸 수 한도
const CONTEXT = 1;         // 바뀐 곳 앞뒤로 보일 같은 줄 수
const CHAR_DIFF_MAX = 600; // 이보다 긴 줄은 글자 단위로 비교하지 않는다

/**
 * Myers 차이 (O((N+M)D)). a, b: 비교할 값 배열 → [['=', i, j] | ['-', i] | ['+', j]] (a·b 순서대로)
 * 다른 곳이 maxD보다 많으면 null
 */
export function myers(a, b, { maxD = MAX_D, eq = (x, y) => x === y } = {}) {
  const n = a.length; const m = b.length;
  const max = n + m;
  const off = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace = [];
  let found = -1;
  for (let d = 0; d <= Math.min(max, maxD); d++) {
    trace.push(v.slice(off - d - 1, off + d + 2));
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[off + k - 1] < v[off + k + 1]) ? v[off + k + 1] : v[off + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && eq(a[x], b[y])) { x++; y++; }
      v[off + k] = x;
      if (x >= n && y >= m) { found = d; break; }
    }
    if (found >= 0) break;
  }
  if (found < 0) return null;
  // 거꾸로 따라가며 편집 순서를 만든다
  const ops = [];
  let x = n; let y = m;
  for (let d = found; d > 0; d--) {
    const pv = trace[d]; // d번째 단계를 시작할 때의 v (k = -d-1 .. d+1 → 인덱스 k + d + 1)
    const at = (k) => pv[k + d + 1];
    const k = x - y;
    const prevK = k === -d || (k !== d && at(k - 1) < at(k + 1)) ? k + 1 : k - 1;
    const prevX = at(prevK);
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) { x--; y--; ops.push(['=', x, y]); }
    if (x === prevX) { y--; ops.push(['+', y]); } else { x--; ops.push(['-', x]); }
  }
  while (x > 0 && y > 0) { x--; y--; ops.push(['=', x, y]); }
  return ops.reverse();
}

/** 앞뒤 같은 부분을 떼고 가운데만 Myers로 (AI 수정은 대개 한두 곳이라 빠르다). 너무 다르면 가운데를 통째로 바꾼 것으로 */
function diffSeq(a, b, opts) {
  let s = 0;
  while (s < a.length && s < b.length && a[s] === b[s]) s++;
  let e = 0;
  while (e < a.length - s && e < b.length - s && a[a.length - 1 - e] === b[b.length - 1 - e]) e++;
  const ops = [];
  for (let i = 0; i < s; i++) ops.push(['=', i, i]);
  const am = a.slice(s, a.length - e); const bm = b.slice(s, b.length - e);
  const mid = myers(am, bm, opts);
  if (mid) for (const o of mid) ops.push(o[0] === '=' ? ['=', o[1] + s, o[2] + s] : o[0] === '-' ? ['-', o[1] + s] : ['+', o[1] + s]);
  else {
    for (let i = 0; i < am.length; i++) ops.push(['-', i + s]);
    for (let j = 0; j < bm.length; j++) ops.push(['+', j + s]);
  }
  for (let i = 0; i < e; i++) ops.push(['=', a.length - e + i, b.length - e + i]);
  return { ops, approximate: !mid };
}

// 글자 단위 비교 → [{ t: '=' | '-' | '+', s }] (붙어 있는 같은 종류는 합친다)
export function charSegments(oldText, newText) {
  const a = [...oldText]; const b = [...newText];
  const { ops } = diffSeq(a, b, { maxD: 400 });
  const segs = [];
  const push = (t, ch) => { const last = segs.at(-1); if (last?.t === t) last.s += ch; else segs.push({ t, s: ch }); };
  for (const o of ops) push(o[0], o[0] === '+' ? b[o[1]] : a[o[1]]);
  return segs;
}

// 바꾼 줄로 짝지을 만큼 비슷한가 (같은 글자가 절반 넘게 남음)
function similar(a, b) {
  if (!a || !b) return false;
  if (a.length > CHAR_DIFF_MAX || b.length > CHAR_DIFF_MAX) return false;
  const same = charSegments(a, b).filter((x) => x.t === '=').reduce((n, x) => n + [...x.s].length, 0);
  return same * 2 >= Math.max([...a].length, [...b].length);
}

/**
 * 줄(문단) 비교. units: [{ text, where? }] → { hunks, stats, truncated, approximate }
 * hunk = { where, before: [줄], items: [{ type: 'del'|'add'|'mod', old?, new?, segs? }], after: [줄] }
 */
export function diffUnits(oldUnits, newUnits) {
  const a = oldUnits.map((u) => u.text); const b = newUnits.map((u) => u.text);
  const { ops, approximate } = diffSeq(a, b);
  const stats = { added: 0, removed: 0, changed: 0 };
  const hunks = [];
  let items = 0; let truncated = false;
  let i = 0;
  while (i < ops.length) {
    if (ops[i][0] === '=') { i++; continue; }
    // 바뀐 덩어리 [i, j)
    let j = i;
    while (j < ops.length && ops[j][0] !== '=') j++;
    const dels = ops.slice(i, j).filter((o) => o[0] === '-').map((o) => o[1]);
    const adds = ops.slice(i, j).filter((o) => o[0] === '+').map((o) => o[1]);
    const out = [];
    // 순서대로 짝지어 비슷하면 '바뀜', 아니면 지움/더함
    let p = 0; let q = 0;
    while (p < dels.length || q < adds.length) {
      const o = dels[p]; const nw = adds[q];
      if (p < dels.length && q < adds.length && similar(a[o], b[nw])) {
        out.push({ type: 'mod', old: a[o], new: b[nw], segs: charSegments(a[o], b[nw]) }); stats.changed++; p++; q++;
      } else if (p < dels.length && (q >= adds.length || dels.length - p >= adds.length - q)) {
        // 빈 줄을 더하고 지운 것은 보이지 않는다 (문단 사이 간격일 뿐)
        if (a[o].trim()) { out.push({ type: 'del', old: a[o] }); stats.removed++; }
        p++;
      } else {
        if (b[nw].trim()) { out.push({ type: 'add', new: b[nw] }); stats.added++; }
        q++;
      }
    }
    if (!out.length) { i = j; continue; }
    // 앞뒤 같은 줄 (빈 줄은 건너뛰고 보이는 것만)
    const ctx = (from, step) => {
      const lines = [];
      for (let k = from; k >= 0 && k < ops.length && lines.length < CONTEXT; k += step) {
        if (ops[k][0] !== '=') break;
        const t = b[ops[k][2]];
        if (t.trim()) lines.push(t);
      }
      return step < 0 ? lines.reverse() : lines;
    };
    const firstNew = adds.length ? adds[0] : null;
    const where = (firstNew !== null ? newUnits[firstNew]?.where : oldUnits[dels[0]]?.where) ?? null;
    if (items < MAX_ITEMS) {
      const keep = out.slice(0, MAX_ITEMS - items);
      if (keep.length < out.length) truncated = true;
      items += keep.length;
      hunks.push({ where, before: ctx(i - 1, -1), items: keep, after: ctx(j, 1) });
    } else truncated = true;
    i = j;
  }
  return { hunks, stats, truncated, approximate };
}

// ── 칸(Excel·CSV) ────────────────────────────────────────

const colName = (c) => { let s = ''; for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };

/** 시트 비교. sheets: [{ name, rows }] → { cells: [{ sheet, cell, old, new, type }], sheetsAdded, sheetsRemoved, stats, truncated } */
export function diffSheets(oldSheets, newSheets) {
  const byName = (list) => new Map(list.map((s) => [s.name, s]));
  const A = byName(oldSheets); const B = byName(newSheets);
  const cells = []; const stats = { added: 0, removed: 0, changed: 0 };
  let truncated = false;
  const sheetsAdded = newSheets.filter((s) => !A.has(s.name)).map((s) => s.name);
  const sheetsRemoved = oldSheets.filter((s) => !B.has(s.name)).map((s) => s.name);
  for (const s of newSheets) {
    const o = A.get(s.name);
    if (!o) continue;
    const rows = Math.max(o.rows.length, s.rows.length);
    for (let r = 0; r < rows; r++) {
      const ro = o.rows[r] ?? []; const rn = s.rows[r] ?? [];
      const cols = Math.max(ro.length, rn.length);
      for (let c = 0; c < cols; c++) {
        const before = String(ro[c] ?? ''); const after = String(rn[c] ?? '');
        if (before === after) continue;
        const type = !before ? 'add' : !after ? 'del' : 'mod';
        stats[{ add: 'added', del: 'removed', mod: 'changed' }[type]]++;
        if (cells.length < MAX_ITEMS) cells.push({ sheet: s.name, cell: `${colName(c)}${r + 1}`, old: before, new: after, type });
        else truncated = true;
      }
    }
  }
  return { cells, sheetsAdded, sheetsRemoved, stats, truncated };
}

// ── 문서 ─────────────────────────────────────────────────

const SHEET_EXTS = new Set(['xlsx', 'xlsm', 'xls', 'csv', 'ods']);

async function unitsOf(path) {
  const r = await readDocument(path);
  if (Array.isArray(r.slides)) {
    return r.slides.flatMap((s, n) => String(s).split('\n').map((text) => ({ text, where: `슬라이드 ${n + 1}` })));
  }
  return String(r.text ?? '').replace(/\r\n?/g, '\n').split('\n').map((text) => ({ text }));
}

/**
 * 옛 파일 → 새 파일에서 바뀐 글. 두 파일은 같은 형식이어야 한다.
 * → { kind: 'text', hunks, stats, truncated, approximate } | { kind: 'cells', cells, sheetsAdded, sheetsRemoved, stats, truncated }
 *   stats = { added, removed, changed }, same = 바뀐 글이 없음
 */
export async function diffDocuments(oldPath, newPath) {
  const ext = extOf(newPath);
  if (extOf(oldPath) !== ext) throw new Error('같은 형식의 문서끼리만 비교할 수 있습니다');
  if (SHEET_EXTS.has(ext)) {
    const [a, b] = await Promise.all([readDocument(oldPath), readDocument(newPath)]);
    const d = diffSheets(a.sheets ?? [], b.sheets ?? []);
    return { kind: 'cells', ...d, same: !d.cells.length && !d.sheetsAdded.length && !d.sheetsRemoved.length };
  }
  const [a, b] = await Promise.all([unitsOf(oldPath), unitsOf(newPath)]);
  const d = diffUnits(a, b);
  return { kind: 'text', ...d, same: !d.hunks.length };
}
