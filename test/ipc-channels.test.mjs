// IPC 채널 표(src/app/ipc-channels.mjs) 시험 — 표 모양, 보낸 쪽 검사(allows), 원본 글자 대조
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as acorn from 'acorn';
import { allows, BLOCK_END, BLOCK_START, channelSpec, channelsFor, HUB_KINDS, IPC, PRELOADS, preloadBlock, VIEW_KINDS } from '../src/app/ipc-channels.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => readFileSync(join(ROOT, rel), 'utf8');

test('표: 모든 줄의 call·kinds가 올바르고, tabstrip은 아직 비어 있다', () => {
  assert.deepEqual(VIEW_KINDS, ['home', 'panel', 'tabstrip', 'hwp', 'office']);
  assert.deepEqual(HUB_KINDS, ['home', 'panel']);
  assert.ok(Object.isFrozen(IPC), '표는 얼려 둔다');
  const open = [];
  for (const [ch, spec] of Object.entries(IPC)) {
    assert.ok(['invoke', 'on', 'send'].includes(spec.call), `${ch}: call`);
    assert.ok(Array.isArray(spec.kinds) && spec.kinds.length > 0, `${ch}: kinds가 비어 있음`);
    for (const k of spec.kinds) assert.ok(VIEW_KINDS.includes(k), `${ch}: 모르는 종류 ${k}`);
    assert.ok(Object.isFrozen(spec) && Object.isFrozen(spec.kinds), `${ch}: 얼려 둔다`);
    if (spec.open) open.push(ch);
  }
  assert.deepEqual(open.sort(), ['editor:i18n', 'i18n:get']);
  assert.deepEqual(channelsFor(['tabstrip']), { invoke: [], on: [], send: [] });
  // 주 창 send 8개, 편집 창 send 3개
  assert.deepEqual(channelsFor(HUB_KINDS).send, ['ai:event', 'batch:progress', 'doc:changed', 'doc:editor', 'doc:opened', 'menu', 'panel', 'toast']);
  assert.deepEqual(channelsFor(['hwp']).send, ['editor:command', 'editor:external-change', 'editor:request-close']);
  assert.deepEqual(channelsFor(['office']).on, ['editor:close', 'editor:dirty', 'editor:i18n', 'editor:log', 'editor:panel', 'editor:reply']);
  const hub = channelsFor(HUB_KINDS);
  assert.deepEqual(hub.on, ['i18n:get', 'smoke:ready']);
  assert.deepEqual(hub.invoke, [...hub.invoke].sort(), '정렬');
});

test('allows: 편집 창은 주 창 채널을, 주 창은 편집 창 채널을 못 부른다', () => {
  assert.equal(allows('ai:send', ['office']), false);
  assert.equal(allows('app:init', ['hwp']), false);
  assert.equal(allows('editor:save', HUB_KINDS), false);
  assert.equal(allows('editor:hwp-pdf', ['office']), false);
  assert.equal(allows('editor:download', ['hwp']), false);
  assert.equal(allows('editor:save', ['hwp']), true);
  assert.equal(allows('ai:send', HUB_KINDS), true);
  assert.equal(allows('i18n:get', null), true);
  assert.equal(allows('ai:send', null), false);
});

test('channelSpec: 없는 채널·다른 호출 종류는 오류', () => {
  assert.throws(() => channelSpec('doc:nope', 'invoke'), /IPC 표에 없는 채널/);
  assert.throws(() => channelSpec('doc:opened', 'invoke'), /IPC 호출 종류가 다름/);
  assert.equal(channelSpec('doc:opened', 'send'), IPC['doc:opened']);
});

test('원본 대조: main.mjs·editor-window.mjs가 등록·보내는 채널 글자는 모두 표에 같은 호출 종류로 있다', () => {
  // 앞이 '.'이나 글자가 아닌 곳의 이름만 (ses.on('will-download'), wc.on('will-navigate') 같은 것은 IPC가 아님)
  const grab = (src, re, call) => [...src.matchAll(re)].map((m) => [m[1], call]);
  const main = read('src/app/main.mjs');
  const ew = read('src/app/editor-window.mjs');
  const used = [
    ...grab(main, /(?<![.\w])handle\('([^']+)'/g, 'invoke'),
    ...grab(main, /(?<![.\w])onIpc\('([^']+)'/g, 'on'),
    ...grab(main, /(?<![.\w])send\('([^']+)'/g, 'send'),
    ...grab(ew, /(?<![.\w])handle\('([^']+)'/g, 'invoke'),
    ...grab(ew, /(?<![.\w])(?:on|onIpc)\('([^']+)'/g, 'on'),
    ...grab(ew, /(?<![.\w])ask\(\w+, '([^']+)'/g, 'send'),
    ...grab(ew, /(?<![.\w])sendToEditor\(\w+, '([^']+)'/g, 'send'),
    ...grab(ew, /webContents\.send\('([^']+)'/g, 'send'),
  ];
  assert.ok(used.length > 70, `찾은 채널이 너무 적다: ${used.length}`);
  for (const [ch, call] of used) {
    assert.ok(IPC[ch], `IPC 표에 없는 채널: ${ch}`);
    assert.equal(IPC[ch].call, call, `${ch}: 호출 종류`);
  }
  // 손으로 ipcMain.on/handle에 채널 글자를 직접 넘기지 않는다 (표 검사를 건너뛰므로)
  assert.doesNotMatch(main, /ipcMain\.(?:on|handle)\('/);
  assert.doesNotMatch(ew, /ipcMain\.(?:on|handle)\('/);
  // main.mjs는 webContents.send('채널')을 직접 부르지 않는다 (send 도우미가 표를 먼저 본다)
  assert.doesNotMatch(main, /webContents\.send\('/);
  // 표의 모든 채널은 두 파일 중 한 곳에 나온다 (안 쓰는 줄 없음)
  const seen = new Set(used.map(([ch]) => ch));
  assert.deepEqual(Object.keys(IPC).filter((ch) => !seen.has(ch)), []);
});

// 자동 생성 블록: BLOCK_START부터 BLOCK_END까지 (줄 끝은 \n으로 맞춘다)
function blockOf(rel) {
  const text = read(rel).replace(/\r\n/g, '\n');
  const a = text.indexOf(BLOCK_START);
  const b = text.indexOf(BLOCK_END);
  assert.ok(a >= 0 && b > a, `${rel}: 자동 생성 블록이 없다`);
  return { block: text.slice(a, b + BLOCK_END.length), outside: text.slice(0, a) + text.slice(b + BLOCK_END.length) };
}

test('preload 블록이 표와 같다 (gen-ipc-preload --check)', () => {
  for (const [rel, kinds] of Object.entries(PRELOADS)) {
    assert.equal(blockOf(rel).block, preloadBlock(kinds), `${rel}: 블록이 표와 다르다`);
  }
  const r = spawnSync(process.execPath, ['scripts/gen-ipc-preload.mjs', '--check'], { cwd: ROOT, encoding: 'utf8' });
  assert.equal(r.status, 0, r.stdout + r.stderr);
});

test('preloadBlock: 이름 정렬·줄 끝 지정', () => {
  const b = preloadBlock(['hwp'], '\r\n');
  assert.ok(b.startsWith(BLOCK_START + '\r\n') && b.endsWith('\r\n' + BLOCK_END));
  assert.equal(b.split('\r\n').length, 3);
  assert.match(b, /send: new Set\(\['editor:command', 'editor:external-change', 'editor:request-close'\]\)/);
});

test('preload 채널 글자 = 표 블록 (빠짐·남음 없음)', () => {
  for (const [rel, kinds] of Object.entries(PRELOADS)) {
    const { outside } = blockOf(rel);
    const used = new Set([
      ...[...outside.matchAll(/'([a-z0-9]+:[\w-]+)'/gi)].map((m) => m[1]),
      ...[...outside.matchAll(/\b(?:invoke|subscribe|answer|sendTo|sendSync)\('([^']+)'/g)].map((m) => m[1]),
    ]);
    const c = channelsFor(kinds);
    assert.deepEqual([...used].sort(), [...c.invoke, ...c.on, ...c.send].sort(), rel);
  }
});

test('홈·패널 종류 = 그 채널을 쓰는 화면 파일', () => {
  // preload.cjs의 exposeInMainWorld 두 번째 인수를 걸어 API 경로 → 채널 표를 만든다
  const ast = acorn.parse(read('src/app/preload.cjs'), { ecmaVersion: 'latest' });
  const HELPERS = new Set(['invoke', 'subscribe', 'answer', 'sendTo', 'sendSync']);
  const children = (node) => Object.entries(node)
    .filter(([k]) => k !== 'type' && k !== 'start' && k !== 'end')
    .flatMap(([, v]) => (Array.isArray(v) ? v : [v]))
    .filter((c) => c && typeof c.type === 'string');
  const channelsIn = (node, out = []) => {
    if (node.type === 'CallExpression' && node.callee.type === 'Identifier' && HELPERS.has(node.callee.name)
      && node.arguments[0]?.type === 'Literal' && typeof node.arguments[0].value === 'string') out.push(node.arguments[0].value);
    for (const c of children(node)) channelsIn(c, out);
    return out;
  };
  const apiPaths = new Map(); // 'doc.pick' → ['doc:pick']
  const walkObject = (obj, prefix) => {
    for (const p of obj.properties) {
      const name = p.key.name ?? p.key.value;
      const path = prefix ? prefix + '.' + name : name;
      if (p.value.type === 'ObjectExpression') walkObject(p.value, path);
      else apiPaths.set(path, channelsIn(p.value));
    }
  };
  let exposed = null;
  (function find(node) {
    if (node.type === 'CallExpression' && node.callee.property?.name === 'exposeInMainWorld') exposed = node.arguments[1];
    for (const c of children(node)) find(c);
  })(ast);
  assert.equal(exposed?.type, 'ObjectExpression', 'exposeInMainWorld 인수를 못 찾음');
  walkObject(exposed, '');
  assert.ok(apiPaths.get('doc.pick')?.includes('doc:pick') && apiPaths.get('on.aiEvent')?.includes('ai:event') && apiPaths.get('smokeReady')?.includes('smoke:ready'));

  // 화면 파일이 쓰는 API 경로 (core.js → 홈·패널 둘 다, home.js → 홈, panel.js → 패널)
  const FILES = { 'src/app/renderer/core.js': ['home', 'panel'], 'src/app/renderer/home.js': ['home'], 'src/app/renderer/panel.js': ['panel'] };
  const usedKinds = new Map(); // 채널 → Set(종류)
  for (const [rel, kinds] of Object.entries(FILES)) {
    const paths = new Set([...read(rel).matchAll(/\bK\.([\w.]+)/g)].map((m) => m[1]));
    for (const [path, chans] of apiPaths) {
      if (!paths.has(path)) continue;
      for (const ch of chans) {
        if (!usedKinds.has(ch)) usedKinds.set(ch, new Set());
        for (const k of kinds) usedKinds.get(ch).add(k);
      }
    }
  }
  const all = new Set([...apiPaths.values()].flat());
  assert.ok(all.size > 40, `노출 채널이 너무 적다: ${all.size}`);
  for (const ch of all) {
    if (IPC[ch].open) continue;
    const want = IPC[ch].kinds.filter((k) => HUB_KINDS.includes(k)).sort();
    const got = [...(usedKinds.get(ch) ?? [])].sort();
    assert.ok(got.length > 0, `${ch}: 쓰는 화면 파일이 없다 (죽은 노출)`);
    assert.deepEqual(got, want, `${ch}: 표의 종류와 쓰는 화면 파일이 다르다`);
  }
});
