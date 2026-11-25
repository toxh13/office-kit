// IPC 채널 표(src/app/ipc-channels.mjs) 시험 — 표 모양, 보낸 쪽 검사(allows), 원본 글자 대조
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { allows, channelSpec, channelsFor, HUB_KINDS, IPC, VIEW_KINDS } from '../src/app/ipc-channels.mjs';

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
  // 표의 모든 채널은 두 파일 중 한 곳에 나온다 (안 쓰는 줄 없음)
  const seen = new Set(used.map(([ch]) => ch));
  assert.deepEqual(Object.keys(IPC).filter((ch) => !seen.has(ch)), []);
});
