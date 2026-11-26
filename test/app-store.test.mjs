// store.mjs 단위 시험 (임시 폴더에서만)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, existsSync, readdirSync, rmSync, statSync, utimesSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createSettings, createBackups, createPreviewCache, createPreviewFiles, writeJsonAtomic, stamp, DEFAULT_SETTINGS, guessModelProvider } from '../src/app/store.mjs';

const tmp = () => mkdtempSync(join(tmpdir(), 'okstore-'));

test('설정: 기본값, 저장, 다시 읽기', () => {
  const dir = tmp();
  const s = createSettings(dir);
  assert.equal(s.get('previewMode'), 'auto');
  s.set({ previewMode: 'exact', chatWidth: 420 });
  const again = createSettings(dir);
  assert.equal(again.get('previewMode'), 'exact');
  assert.equal(again.get('chatWidth'), 420);
  assert.equal(again.get('theme'), DEFAULT_SETTINGS.theme);
  assert.equal(again.get('tabSession'), null);
  assert.equal(again.get('lastTabSession'), null);
  assert.deepEqual(readdirSync(dir).filter((f) => f.endsWith('.tmp')), []);
  rmSync(dir, { recursive: true, force: true });
});

test('설정: 깨진 JSON이면 기본값', () => {
  const dir = tmp();
  writeFileSync(join(dir, 'settings.json'), '{ 깨짐');
  const s = createSettings(dir);
  assert.equal(s.get('previewMode'), 'auto');
  assert.deepEqual(s.get('recent'), []);
});

test('최근 파일: 중복 제거(대소문자 무시), 10개, 없는 파일 표시', () => {
  const dir = tmp();
  const s = createSettings(dir);
  const files = Array.from({ length: 12 }, (_, i) => { const p = join(dir, `f${i}.hwp`); writeFileSync(p, 'x'); return p; });
  files.forEach((f) => s.addRecent(f));
  s.addRecent(files[3].toUpperCase());
  const r = s.recent();
  assert.equal(r.length, 10);
  assert.equal(r[0].path, files[3].toUpperCase());
  assert.equal(r.filter((x) => x.path.toLowerCase() === files[3].toLowerCase()).length, 1);
  rmSync(files[11]);
  assert.equal(s.recent().find((x) => x.path === files[11]).exists, false);
  s.removeRecent(files[11]);
  assert.equal(s.recent().length, 9);
  s.clearRecent();
  assert.equal(s.recent().length, 0);
});

test('백업: 이름 형식, 같은 내용은 재사용, 50개 유지, 복원', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { keep: 3 });
  const doc = join(dir, '보고서.hwp');
  writeFileSync(doc, 'v1');
  const e1 = b.backup(doc, new Date(2026, 9, 17, 9, 5, 7));
  assert.equal(e1.file, '20261017-090507-보고서.hwp');
  assert.equal(e1.reused, false);
  const e1b = b.backup(doc, new Date(2026, 9, 17, 9, 5, 7));
  assert.equal(e1b.reused, true);
  writeFileSync(doc, 'v2');
  const e2 = b.backup(doc, new Date(2026, 9, 17, 9, 5, 7)); // 같은 초 → 번호 붙음
  assert.match(e2.file, /^20261017-090507-2-보고서\.hwp$/);
  assert.equal(b.latest(doc).file, e2.file);
  writeFileSync(doc, 'v3');
  b.restore(doc, e1);
  assert.equal(readFileSync(doc, 'utf8'), 'v1');
  // keep=3 → 오래된 것부터 지움
  for (let i = 0; i < 4; i++) { writeFileSync(doc, `x${i}`); b.backup(doc, new Date(2026, 9, 18, 0, 0, i)); }
  const files = readdirSync(join(dir, 'backups')).filter((f) => f !== 'index.json');
  assert.equal(files.length, 3);
  assert.equal(b.list(doc).length, 3);
  assert.equal(existsSync(e1.path), false);
});

test('백업: 되돌린 파일은 지금 시각으로 바뀐 파일로 보인다 (백업의 옛 수정 시각을 옮기지 않음)', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'));
  const doc = join(dir, '보고서.docx');
  writeFileSync(doc, 'v1');
  const old = new Date(2020, 0, 1);
  utimesSync(doc, old, old);
  const e1 = b.backup(doc); // Windows 복사(CopyFileW)는 옛 수정 시각까지 옮긴다
  writeFileSync(doc, 'v2');
  const before = Date.now();
  b.restore(doc, e1);
  assert.equal(readFileSync(doc, 'utf8'), 'v1');
  assert.ok(statSync(doc).mtimeMs >= before - 2000, `mtime=${new Date(statSync(doc).mtimeMs).toISOString()}`);
  rmSync(dir, { recursive: true, force: true });
});

test('백업: 없는 문서 복원은 한국어 오류', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'));
  assert.throws(() => b.restore(join(dir, '없음.hwp')), /백업이 없습니다/);
});

test('원자적 쓰기 · 시각 문자열', () => {
  const dir = tmp();
  const f = join(dir, 'a.json');
  writeJsonAtomic(f, { a: 1 });
  writeJsonAtomic(f, { a: 2 });
  assert.deepEqual(JSON.parse(readFileSync(f, 'utf8')), { a: 2 });
  assert.equal(stamp(new Date(2026, 0, 2, 3, 4, 5)), '20260102-030405');
});

// ── 검토에서 나온 문제 회귀 시험 ──────────────────────────

const read = (p) => readFileSync(p, 'utf8');

test('백업: 같은 내용을 다시 쓰면 usedAt이 새로워져 먼저 지워지지 않는다', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { keep: 3 });
  const doc = join(dir, '양식.hwp'); writeFileSync(doc, 'ORIGINAL');
  let t = Date.UTC(2026, 9, 1);
  const first = b.backup(doc, new Date(t += 60000));
  for (const n of ['a', 'b']) { const p = join(dir, `${n}.hwp`); writeFileSync(p, n); b.backup(p, new Date(t += 60000)); }
  const again = b.backup(doc, new Date(t += 60000));        // 내용 그대로 다시 열기
  assert.equal(again.reused, true);
  assert.equal(again.file, first.file);
  assert.equal(again.time, first.time);                     // 만든 시각은 그대로
  const c = join(dir, 'c.hwp'); writeFileSync(c, 'c');
  b.backup(c, new Date(t += 60000));                        // 4개째 → 가장 오래 안 쓴 a가 지워져야 한다
  assert.equal(existsSync(again.path), true);
  assert.equal(b.list(join(dir, 'a.hwp')).length, 0);
});

test('백업: 원래대로 순서(지금 상태 백업 → 복원)에서 되돌릴 백업이 지워지지 않는다 (keep=3)', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { keep: 3 });
  const doc = join(dir, 'template.hwp'); writeFileSync(doc, 'ORIGINAL');
  let t = Date.UTC(2026, 9, 1);
  b.backup(doc, new Date(t += 60000));
  for (const n of ['a', 'b']) { const p = join(dir, `${n}.hwp`); writeFileSync(p, n); b.backup(p, new Date(t += 60000)); }
  const opened = b.backup(doc, new Date(t += 60000));
  writeFileSync(doc, 'EDITED BY AI');
  b.backup(doc, new Date(t += 60000), { protect: [opened.file] }); // main.restore()와 같은 순서
  b.restore(doc, opened);
  assert.equal(read(doc), 'ORIGINAL');
  assert.deepEqual(b.list(doc).map((e) => read(e.path)).sort(), ['EDITED BY AI', 'ORIGINAL']);
});

test('백업: 기본 200개, 한 달 전 연 문서를 다시 열고 원래대로 해도 원본이 남는다', () => {
  const dir = tmp();
  const docs = join(dir, 'docs'); mkdirSync(docs);
  const b = createBackups(join(dir, 'backups'));
  const X = join(docs, '보고서.hwp'); writeFileSync(X, 'ORIGINAL');
  const day = 86400000; const t0 = Date.now() - 30 * day;
  b.backup(X, new Date(t0));
  for (let i = 1; i <= 199; i++) { const f = join(docs, `other${i}.docx`); writeFileSync(f, `o${i}`); b.backup(f, new Date(t0 + i * 60000)); }
  const openEntry = b.backup(X);
  assert.equal(openEntry.reused, true);
  writeFileSync(X, 'EDITED BY AI');
  b.backup(X, new Date(), { protect: [openEntry.file] });
  assert.equal(existsSync(openEntry.path), true);
  b.restore(X, openEntry);
  assert.equal(read(X), 'ORIGINAL');
  assert.equal(readdirSync(join(dir, 'backups')).filter((f) => f !== 'index.json').length, 200);
});

test('버전: 까닭(reason)을 적고, 같은 내용을 다시 쓰면 처음 까닭 그대로 · find로 하나 찾기', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'));
  const doc = join(dir, '보고서.docx'); writeFileSync(doc, 'A');
  let t = Date.UTC(2026, 9, 1);
  const opened = b.backup(doc, new Date(t += 1000), { reason: 'open' });
  assert.equal(opened.reason, 'open');
  assert.equal(b.backup(doc, new Date(t += 1000), { reason: 'ai' }).reason, 'open', '내용이 같으면 새로 만들지 않음');
  writeFileSync(doc, 'B');
  const ai = b.backup(doc, new Date(t += 1000), { reason: 'ai' });
  writeFileSync(doc, 'C');
  b.backup(doc, new Date(t += 1000), { reason: 'editor' });
  writeFileSync(doc, 'D');
  b.backup(doc, new Date(t += 1000), { reason: 'restore' });
  b.backup(doc, new Date(t += 1000), { reason: '모름' }); // 모르는 까닭은 적지 않음 (D는 이미 있음)
  assert.deepEqual(b.list(doc).map((e) => [read(e.path), e.reason]), [['D', 'restore'], ['C', 'editor'], ['B', 'ai'], ['A', 'open']]);
  assert.equal(b.find(doc, ai.file).reason, 'ai');
  assert.equal(b.find(doc, 'none.docx'), null);
  assert.equal(b.find(join(dir, '다른.docx'), ai.file), null, '다른 문서의 버전은 찾지 않음');
});

test('버전: 모두 합친 바이트 한도 (maxBytes) — 오래 안 쓴 것부터, 방금 것·protect는 남김', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { maxBytes: 250 });
  let t = Date.UTC(2026, 9, 1);
  const made = [];
  for (const n of ['a', 'b', 'c']) { const p = join(dir, `${n}.hwp`); writeFileSync(p, n.repeat(100)); made.push(b.backup(p, new Date(t += 1000), { protect: made.slice(0, 1) })); }
  // 300바이트 > 250 → 보호된 a는 남고 b가 지워짐
  assert.deepEqual(made.map((e) => existsSync(e.path)), [true, false, true]);
});

test('백업: protect 없이도 방금 만든·다시 쓴 백업은 지우지 않고, protect가 많으면 keep을 넘겨 남긴다', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { keep: 2 });
  let t = Date.UTC(2026, 9, 1);
  const made = ['p', 'q', 'r'].map((n) => { const p = join(dir, `${n}.hwp`); writeFileSync(p, n); return b.backup(p, new Date(t += 1000), { protect: [] }); });
  assert.equal(b.list(join(dir, 'p.hwp')).length, 0);
  const s = join(dir, 's.hwp'); writeFileSync(s, 's');
  b.backup(s, new Date(t += 1000), { protect: made.slice(1) }); // q, r 보호 + 방금 s → 3개
  assert.equal(readdirSync(join(dir, 'backups')).filter((f) => f !== 'index.json').length, 3);
});

test('백업: 되돌린 뒤 다시 열면 같은 내용의 옛 백업을 다시 쓴다 (사본을 늘리지 않음)', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'));
  const doc = join(dir, 'x.docx'); writeFileSync(doc, 'A');
  let t = Date.UTC(2026, 9, 1);
  const e1 = b.backup(doc, new Date(t += 1000));
  writeFileSync(doc, 'B');
  b.backup(doc, new Date(t += 1000), { protect: [e1.file] });
  b.restore(doc, e1);
  const again = b.backup(doc, new Date(t += 1000));
  assert.equal(again.reused, true);
  assert.equal(again.file, e1.file);
  assert.equal(b.latest(doc).file, e1.file);
  assert.equal(b.list(doc).length, 2);
});

test('설정: 옛 model 하나는 주인 AI의 모델로 옮기고, 모르면 버린다', () => {
  const cases = [
    [{ model: 'sonnet', provider: null }, { claude: 'sonnet' }],
    [{ model: 'gpt-5', provider: 'claude' }, { codex: 'gpt-5' }],
    [{ model: 'my-model', provider: 'codex' }, { codex: 'my-model' }],
    [{ model: 'my-model', provider: null }, {}],
    [{ model: '  ', models: { claude: 'opus' } }, { claude: 'opus' }],
    [{ models: ['bad'] }, {}],
  ];
  for (const [saved, want] of cases) {
    const dir = tmp();
    writeFileSync(join(dir, 'settings.json'), JSON.stringify(saved));
    const s = createSettings(dir);
    assert.deepEqual(s.get('models'), want, JSON.stringify(saved));
    assert.equal(s.get('model'), undefined);
  }
  assert.equal(guessModelProvider('claude-sonnet-4-5'), 'claude');
  assert.equal(guessModelProvider('o3'), 'codex');
  assert.equal(guessModelProvider('llama'), null);
  assert.deepEqual(DEFAULT_SETTINGS.models, {});
});

test('백업: maxPerSource — 편집 창에서 자주 저장해도 그 문서 몫만 최근 n개, 다른 문서·원래대로 기준은 남는다', () => {
  const dir = tmp();
  const b = createBackups(join(dir, 'backups'), { keep: 50 });
  let t = Date.UTC(2026, 9, 1);
  const others = ['a', 'b', 'c'].map((n) => { const p = join(dir, `${n}.hwp`); writeFileSync(p, n); return b.backup(p, new Date(t += 1000)); });
  const doc = join(dir, '편집.hwp'); writeFileSync(doc, 'ORIGINAL');
  const opened = b.backup(doc, new Date(t += 1000)); // 열 때 만든 '원래대로' 기준
  for (let i = 1; i <= 10; i++) {
    b.backup(doc, new Date(t += 1000), { protect: [opened.file], maxPerSource: 3 });
    writeFileSync(doc, `저장 ${i}`);
  }
  const mine = b.list(doc);
  assert.equal(mine.length, 4);                               // 최근 3개 + 보호된 기준
  assert.ok(mine.some((e) => e.file === opened.file));
  assert.equal(read(mine[0].path), '저장 9');                  // 가장 최근 = 마지막 저장 직전 내용
  for (const o of others) assert.equal(existsSync(o.path), true);
  // maxPerSource를 안 주면 예전처럼 따로 세지 않는다
  b.backup(doc, new Date(t += 1000));
  assert.equal(b.list(doc).length, 5);
});

test('미리보기 기억: 경로·크기·수정 시각·방식이 같을 때만 다시 쓰고, 개수·바이트 한도', () => {
  const dir = tmp();
  const c = createPreviewCache(join(dir, 'preview-cache'), { keep: 2, maxBytes: 1000 });
  const pdf = (name, text) => { const p = join(dir, name); writeFileSync(p, text); return p; };
  const what = { path: join(dir, '표.xlsx'), size: 10, mtimeMs: 1000, mode: 'auto' };
  assert.equal(c.get(what), null);
  c.put(what, pdf('a.pdf', '%PDF-a'), { engine: 'excel', pages: 2 });
  const hit = c.get({ ...what, path: what.path.toUpperCase() }); // Windows: 대소문자 무시
  assert.deepEqual([read(hit.file), hit.engine, hit.pages], ['%PDF-a', 'excel', 2]);
  assert.equal(c.get({ ...what, mtimeMs: 2000 }), null, '파일이 바뀌면 안 씀');
  assert.equal(c.get({ ...what, size: 11 }), null);
  assert.equal(c.get({ ...what, mode: 'exact' }), null, '방식이 다르면 따로');
  // 경로는 해시로만 남는다
  assert.doesNotMatch(read(join(dir, 'preview-cache', 'index.json')), /표\.xlsx/);
  // 개수 한도: 오래 안 쓴 것부터
  c.put({ ...what, path: join(dir, 'b.docx') }, pdf('b.pdf', '%PDF-b'));
  c.get(what); // a를 다시 씀 → b가 더 오래 안 씀
  c.put({ ...what, path: join(dir, 'c.docx') }, pdf('c.pdf', '%PDF-c'));
  assert.ok(c.get(what));
  assert.equal(c.get({ ...what, path: join(dir, 'b.docx') }), null);
  // 바이트 한도
  c.put({ ...what, path: join(dir, 'big.docx') }, pdf('big.pdf', 'x'.repeat(999)));
  assert.equal(readdirSync(join(dir, 'preview-cache')).filter((f) => f.endsWith('.pdf')).length, 1);
});

test('미리보기 기억: 같은 밀리초 안에 다시 쓴 것도 나중에 쓴 것으로 친다 (시계 고정)', () => {
  const dir = tmp();
  const realNow = Date.now;
  Date.now = () => 5000; // 모든 일이 같은 밀리초에
  try {
    const c = createPreviewCache(join(dir, 'preview-cache'), { keep: 2, maxBytes: 1e6 });
    const pdf = (name) => { const p = join(dir, name); writeFileSync(p, `%PDF-${name}`); return p; };
    const what = (n) => ({ path: join(dir, n), size: 10, mtimeMs: 1000, mode: 'auto' });
    c.put(what('a.docx'), pdf('a.pdf'));
    c.put(what('b.docx'), pdf('b.pdf'));
    assert.ok(c.get(what('a.docx'))); // a를 다시 씀 → b가 가장 오래 안 씀
    c.put(what('c.docx'), pdf('c.pdf'));
    assert.ok(c.get(what('a.docx')), '방금 쓴 a는 남는다');
    assert.equal(c.get(what('b.docx')), null, 'b가 밀려난다');
  } finally {
    Date.now = realNow;
  }
});

test('최근 문서 고정: 고정한 것이 먼저, 지우기·밀려남에도 남는다', async () => {
  const { mkdtempSync, rmSync } = await import('node:fs');
  const { tmpdir } = await import('node:os');
  const { join } = await import('node:path');
  const dir = mkdtempSync(join(tmpdir(), 'officekit-pin-'));
  try {
    const s = createSettings(dir);
    const p = (n) => join(dir, `문서${n}.docx`);
    for (let i = 1; i <= 3; i++) s.addRecent(p(i));
    s.pinRecent(p(1), true);
    assert.deepEqual(s.recent().map((r) => r.name), ['문서1.docx', '문서3.docx', '문서2.docx']);
    assert.equal(s.recent()[0].pinned, true);
    // 고정 안 한 것이 많아져도 고정은 남는다
    for (let i = 10; i < 25; i++) s.addRecent(p(i));
    const names = s.recent().map((r) => r.name);
    assert.equal(names[0], '문서1.docx');
    assert.equal(names.length, 11, '고정 1 + 최근 10');
    // 다시 열어도 고정은 그대로
    s.addRecent(p(1));
    assert.equal(s.recent()[0].pinned, true);
    // 목록 지우기: 고정은 남김
    s.clearRecent();
    assert.deepEqual(s.recent().map((r) => r.name), ['문서1.docx']);
    s.pinRecent(p(1), false);
    assert.equal(s.recent()[0].pinned, false);
    assert.equal(createSettings(dir).recent()[0].pinned, false, '파일에 저장됨');
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('restore는 백업을 원자적으로 복사한다 (내용 = 백업, 임시 파일 없음)', () => {
  const dir = tmp();
  const doc = join(dir, '문서.hwp');
  writeFileSync(doc, 'v1');
  const b = createBackups(join(dir, 'backups'));
  const e1 = b.backup(doc, new Date(2026, 9, 17, 9, 5, 7));
  writeFileSync(doc, 'v2-바뀐 내용');
  b.restore(doc, e1);
  assert.equal(readFileSync(doc, 'utf8'), 'v1');
  assert.deepEqual(readdirSync(dir).filter((f) => f.endsWith('.tmp')), []);
  rmSync(dir, { recursive: true, force: true });
});

test('미리보기 폴더: 늦게 끝난 미리보기는 지금 문서의 미리보기 폴더를 지우지 않는다', () => {
  const dir = tmp();
  const files = createPreviewFiles(dir);
  // A 미리보기를 시작했는데 오래 걸리는 사이 B를 열어 B 미리보기가 먼저 끝났다
  // A의 PDF는 미리보기 폴더 밖에 둔다: B를 놓을 때의 폴더 정리가 아니라 늦은 결과를 버리는 쪽이 지웠음을 확인하려고
  const elsewhere = tmp();
  const rawA = join(elsewhere, '가.docx.1.pdf');
  const rawB = join(dir, '나.docx.2.pdf');
  writeFileSync(rawA, '%PDF-A');
  writeFileSync(rawB, '%PDF-B');
  const b = files.settle(rawB, 'C:\\Users\\user\\문서\\나.docx', { current: true });
  assert.ok(b && existsSync(b.file), 'B 미리보기 PDF가 놓인다');
  assert.equal(readFileSync(b.file, 'utf8'), '%PDF-B');
  // 이제 A가 끝났다 (지금 문서는 B)
  const a = files.settle(rawA, 'C:\\Users\\user\\문서\\가.docx', { current: false });
  assert.equal(a, null, '늦은 결과는 둘 곳이 없다');
  assert.ok(existsSync(b.file), 'B의 PDF는 남아 있어야 한다');
  assert.equal(existsSync(rawA), false, '늦은 A의 PDF는 지운다');
  assert.equal(readdirSync(dir).length, 1, 'B 폴더 하나만 남는다');
  rmSync(dir, { recursive: true, force: true });
  rmSync(elsewhere, { recursive: true, force: true });
});

test('previewFiles: keep 목록 문서마다 가장 새 PDF를 남긴다 (여러 탭·창)', () => {
  const dir = tmp();
  const files = createPreviewFiles(dir);
  const a = 'C:\\Users\\user\\문서\\가.docx';
  const b = 'C:\\Users\\user\\문서\\나.pdf';
  const c = 'C:\\Users\\user\\문서\\다.xlsx';
  let n = 0;
  const raw = (doc) => { const f = join(dir, `${doc.split('\\').at(-1)}.${++n}.pdf`); writeFileSync(f, `%PDF-${n}`); return f; };
  // a 사본 2개, b 1개, c 1개 (각 탭·창의 미리보기)
  const a1 = files.tidy(raw(a), a);
  const a2 = files.tidy(raw(a), a);
  const b1 = files.tidy(raw(b), b);
  const c1 = files.tidy(raw(c), c);
  const stray = join(dir, '지난 실행.pdf');
  writeFileSync(stray, '%PDF-old');
  files.clean([a, b]);
  assert.equal(existsSync(a2.file), true, 'a의 가장 새 PDF는 남긴다');
  assert.equal(existsSync(b1.file), true, 'b의 PDF는 남긴다');
  assert.equal(existsSync(a1.file), false, 'a의 옛 PDF는 지운다');
  assert.equal(existsSync(c1.file), false, '열려 있지 않은 c의 PDF는 지운다');
  assert.equal(existsSync(stray), true, '주인 모를 파일은 실행 중에는 지우지 않는다 (다른 탭이 만드는 중일 수 있다) — 의도한 변경 (F3)');
  assert.equal(readdirSync(dir).length, 3);
  rmSync(stray);
  // 새 미리보기를 놓을 때도 keep 문서들의 PDF는 남긴다 (자기 문서는 늘 남긴다)
  const b2 = files.settle(raw(b), b, { keep: [a] });
  assert.ok(b2 && existsSync(b2.file));
  assert.equal(existsSync(a2.file), true, '다른 탭(a)의 미리보기는 그대로');
  assert.equal(existsSync(b1.file), false, 'b의 옛 PDF는 지운다');
  // 대소문자만 다른 경로도 같은 문서 (Windows)
  if (process.platform === 'win32') {
    files.clean([a.toUpperCase(), b]);
    assert.equal(existsSync(a2.file), true);
  }
  files.clean(null);
  assert.deepEqual(readdirSync(dir), [], 'null이면 모두 지운다');
  rmSync(dir, { recursive: true, force: true });
});

test('previewFiles: 다른 탭이 만드는 중인 PDF(주인 없음)는 실행 중 정리에서 지우지 않고, 시작·끝낼 때(null)만 지운다 (F3)', () => {
  const dir = tmp();
  const files = createPreviewFiles(dir);
  const a = 'C:\\Users\\user\\문서\\가.docx';
  // 다른 탭이 미리보기 PDF를 만드는 중 (아직 settle 전 — 주인을 모른다)
  const producing = join(dir, '나.docx.7.pdf');
  writeFileSync(producing, '%PDF-half');
  const raw = join(dir, '가.docx.8.pdf');
  writeFileSync(raw, '%PDF-A');
  const pdf = files.settle(raw, a);
  assert.ok(pdf && existsSync(pdf.file));
  assert.equal(existsSync(producing), true, '만드는 중인 PDF는 남긴다');
  files.clean([a]);
  assert.equal(existsSync(producing), true, '실행 중 정리도 주인 있는 것만 지운다');
  files.clean(null);
  assert.deepEqual(readdirSync(dir), [], '시작·끝낼 때는 남은 것까지 모두 지운다');
  rmSync(dir, { recursive: true, force: true });
});
