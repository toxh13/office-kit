import test from 'node:test';
import assert from 'node:assert/strict';
import { makeFakeElectron, installFakeElectron, ELECTRON_EXPORTS } from './helpers/fake-electron.mjs';

const tick = () => new Promise((r) => setImmediate(r));

test('BaseWindow: contentView에 뷰 붙이기·떼기, 다시 붙이면 맨 위', () => {
  const f = makeFakeElectron();
  const w = new f.BaseWindow({ width: 800 });
  const a = new f.WebContentsView({ webPreferences: { preload: 'x.cjs' } });
  const b = new f.WebContentsView({});
  w.contentView.addChildView(a); w.contentView.addChildView(b); w.contentView.addChildView(a);
  assert.deepEqual(w.contentView.children, [b, a]);
  w.contentView.removeChildView(b);
  assert.deepEqual(w.contentView.children, [a]);
  assert.equal(a.opts.webPreferences.preload, 'x.cjs');
  assert.deepEqual(f.views, [a, b]);
});

test('BaseWindow: close는 막을 수 있고, 안 막으면 closed·getAllWindows·fromId에서 빠진다', () => {
  const f = makeFakeElectron();
  const w = new f.BaseWindow();
  let closed = 0, block = true;
  w.on('close', (e) => { if (block) e.preventDefault(); });
  w.on('closed', () => { closed += 1; });
  w.close();
  assert.equal(w.isDestroyed(), false);
  assert.deepEqual(f.BaseWindow.getAllWindows(), [w]);
  block = false;
  w.close();
  assert.equal(closed, 1);
  assert.equal(w.isDestroyed(), true);
  assert.equal(f.BaseWindow.fromId(w.id), null);
  assert.deepEqual(f.BaseWindow.getAllWindows(), []);
});

test('webContents: close()는 destroyed를 한 번만, setBackgroundThrottling 기록, 포커스', () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  let n = 0;
  wc.on('destroyed', () => { n += 1; });
  wc.close(); wc.close();
  assert.equal(n, 1);
  assert.equal(wc.isDestroyed(), true);
  wc.setBackgroundThrottling(false);
  assert.equal(wc.backgroundThrottling, false);
  assert.equal(wc.isFocused(), false);
  wc.focus();
  assert.equal(wc.isFocused(), true);
});

test('keyDown: before-input-event에서 preventDefault하면 prevented', () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  wc.on('before-input-event', (e, input) => { if (input.key === 'w') e.preventDefault(); });
  assert.equal(f.keyDown(wc, { type: 'keyDown', key: 'w', control: true }).prevented, true);
  assert.equal(f.keyDown(wc, { type: 'keyDown', key: 'q', control: true }).prevented, false);
});

test('crash: render-process-gone에 reason', () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  const got = [];
  wc.on('render-process-gone', (_e, d) => got.push(d.reason));
  f.crash(wc, 'oom');
  f.crash(wc);
  assert.deepEqual(got, ['oom', 'crashed']);
});

test('dialog: 답을 넣은 차례대로, 없으면 기본값', async () => {
  const f = makeFakeElectron();
  f.dialogAnswers.push(0, 1);
  assert.equal((await f.dialog.showMessageBox(null, { a: 1 })).response, 0);
  assert.equal((await f.dialog.showMessageBox(null, {})).response, 1);
  assert.equal((await f.dialog.showMessageBox(null, {})).response, 2);
  assert.equal(f.dialogs.length, 3);
  assert.deepEqual(await f.dialog.showSaveDialog(null, { t: 1 }), { canceled: true });
  assert.deepEqual(f.saveDialogs, [{ t: 1 }]);
  f.saveDialogAnswers.push({ canceled: false, filePath: 'a' });
  assert.equal((await f.dialog.showSaveDialog(null, {})).filePath, 'a');
  assert.deepEqual(await f.dialog.showOpenDialog(null, {}), { canceled: true, filePaths: [] });
  f.openDialogAnswers.push({ canceled: false, filePaths: ['b'] });
  assert.deepEqual((await f.dialog.showOpenDialog(null, {})).filePaths, ['b']);
  f.dialog.showErrorBox('제목', '내용');
  assert.equal(f.errorBox, '제목: 내용');
});

test('IPC: invokeAs의 senderFrame이 그대로 핸들러에 간다, emitSync는 returnValue', async () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  f.ipcMain.handle('x', (e) => e.senderFrame.url);
  assert.equal(await f.invokeAs(f.evOf(wc, { url: 'file:///sub' }), 'x'), 'file:///sub');
  f.ipcMain.on('y', (e) => { e.returnValue = 7; });
  assert.equal(f.emitSync(wc, 'y'), 7);
  assert.throws(() => f.ipcMain.handle('x', () => 1), /이미 등록된 채널/);
  await assert.rejects(f.invokeFrom(wc, 'nope'), /없는 채널/);
});

test('invoke: 저장 앞에 mouseDown을 보내고 noUserAct면 보내지 않는다', async () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  const seen = [];
  wc.on('input-event', (_e, ev) => seen.push(ev.type));
  f.ipcMain.handle('editor:save', () => 'ok');
  assert.equal(await f.invoke(wc, 'editor:save'), 'ok');
  assert.deepEqual(seen, ['mouseDown']);
  f.noUserAct = true;
  await f.invoke(wc, 'editor:save');
  assert.deepEqual(seen, ['mouseDown']);
});

test('commandAnswer: 없으면 답하지 않고, 있으면 editor:reply로 그 값', async () => {
  const f = makeFakeElectron();
  const wc = new f.WebContentsView({}).webContents;
  const replies = [];
  f.ipcMain.on('editor:reply', (_e, r) => replies.push(r));
  wc.send('editor:command', { id: 1, command: 'save' });
  await tick(); await tick();
  assert.deepEqual(replies, []);
  assert.deepEqual(f.commands, ['save']);
  f.commandAnswer = async () => 'yes';
  wc.send('editor:command', { id: 2, command: 'save' });
  await tick(); await tick();
  assert.deepEqual(replies, [{ id: 2, value: 'yes' }]);
  assert.deepEqual(f.commands, ['save', 'save']);
  // 닫기 질문은 replies를 따른다
  f.replies.set(wc.id, true);
  wc.send('editor:request-close', { id: 3 });
  await tick(); await tick();
  assert.deepEqual(replies.at(-1), { id: 3, value: true });
});

test('BrowserWindow: BaseWindow이기도 하고, 만들 때 browser-window-created, destroy는 wc destroyed를 보내지 않음', () => {
  const f = makeFakeElectron();
  const created = [];
  f.app.on('browser-window-created', (_e, w) => created.push(w));
  const base = new f.BaseWindow();
  const w = new f.BrowserWindow({ title: 't' });
  assert.ok(w instanceof f.BaseWindow);
  assert.deepEqual(created, [w]);
  assert.equal(w.getTitle(), 't');
  assert.deepEqual(f.BrowserWindow.getAllWindows(), [w]);
  assert.deepEqual(f.BaseWindow.getAllWindows(), [base, w]);
  assert.deepEqual(f.windows, [base, w]);
  let destroyed = 0;
  w.webContents.on('destroyed', () => { destroyed += 1; });
  w.destroy();
  assert.equal(destroyed, 0);
  assert.equal(w.isDestroyed(), true);
});

test('app·session·상태: 호출마다 따로, 세션은 이름별 Map', () => {
  const f = makeFakeElectron({ version: '9.9.9' });
  const g = makeFakeElectron();
  assert.equal(f.app.getVersion(), '9.9.9');
  const s = f.session.fromPartition('p');
  assert.equal(f.session.fromPartition('p'), s);
  assert.equal(f.sessions.get('p'), s);
  s.setSpellCheckerEnabled(true);
  assert.equal(s.spellcheck, true);
  f.session.defaultSession.setSpellCheckerDictionaryDownloadURL('u');
  assert.equal(f.defaultDictionaryUrl, 'u');
  assert.equal(new g.BaseWindow().id, 1);
  assert.equal(new f.BaseWindow().id, 1);
  assert.equal(typeof f.app.getPath('temp'), 'string');
});

test('installFakeElectron: electron 불러오기가 가짜로, 두 번째는 교체만', async () => {
  const f = makeFakeElectron();
  installFakeElectron(f);
  const mod = await import('electron');
  assert.equal(mod.default, f);
  for (const n of ELECTRON_EXPORTS) assert.equal(mod[n], f[n], n);
  const g = makeFakeElectron();
  installFakeElectron(g);
  assert.equal(globalThis.__okFakeElectron, g);
});
