// renderer-split.test.mjs — 화면 코드 나누기(core.js · home.js · panel.js + bridge.js)의 경계 시험.
// 진입점 사이 직접 호출은 bridge.js의 이름 붙은 호출로만 한다 (M3에서 이 이름들이 IPC로 바뀐다).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as acorn from 'acorn';
import '../src/app/renderer/bridge.js';
import { VIEW_ROUTES } from '../src/app/ipc-channels.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const RENDERER = join(ROOT, 'src', 'app', 'renderer');
const read = (name) => readFileSync(join(RENDERER, name), 'utf8');
const ENTRY = ['core.js', 'doc-actions.js', 'home.js', 'panel.js'];

// 표준 JS·DOM 전역 + 앞서 불러온 스크립트가 만든 전역
const ALLOWED = new Set([
  'Array', 'Boolean', 'Date', 'Error', 'TypeError', 'Map', 'Set', 'WeakMap', 'Math', 'Number', 'Object', 'Promise', 'String',
  'URL', 'JSON', 'RegExp', 'Symbol', 'Infinity', 'NaN', 'undefined', 'globalThis', 'console',
  'window', 'document', 'ResizeObserver', 'requestAnimationFrame', 'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval',
  'location', 'URLSearchParams', 'CustomEvent',
  'OKCore', 'OKBridge', 'OKTargetChips', 'OKMarkdown', 'OKDialogQueue', 'OKDocActions', 'OfficeKitI18n',
]);

// ── 작은 범위 분석: 자유 변수 찾기 ──
function patternNames(p, out = []) {
  if (!p) return out;
  if (p.type === 'Identifier') out.push(p.name);
  else if (p.type === 'ObjectPattern') for (const pr of p.properties) patternNames(pr.type === 'RestElement' ? pr.argument : pr.value, out);
  else if (p.type === 'ArrayPattern') for (const e of p.elements) patternNames(e, out);
  else if (p.type === 'RestElement') patternNames(p.argument, out);
  else if (p.type === 'AssignmentPattern') patternNames(p.left, out);
  return out;
}
const isNode = (v) => v && typeof v.type === 'string';
const children = (node) => Object.keys(node).flatMap((k) => (k === 'loc' ? [] : [].concat(node[k]).filter(isNode)));

/** 블록 범위 선언 (let/const/class/function) */
function blockDecls(body, s) {
  for (const st of body) {
    if (st.type === 'VariableDeclaration' && st.kind !== 'var') st.declarations.forEach((d) => patternNames(d.id, []).forEach((n) => s.add(n)));
    else if ((st.type === 'FunctionDeclaration' || st.type === 'ClassDeclaration') && st.id) s.add(st.id.name);
  }
}
/** 함수 범위 var (중첩 함수 안은 빼고) */
function varDecls(node, s) {
  if (!isNode(node) || /Function/.test(node.type)) return;
  if (node.type === 'VariableDeclaration' && node.kind === 'var') node.declarations.forEach((d) => patternNames(d.id, []).forEach((n) => s.add(n)));
  for (const c of children(node)) varDecls(c, s);
}

/** 소스 → { ast, free: Map(자유 변수 이름 → 줄 번호들) } */
function freeVariables(src) {
  const ast = acorn.parse(src, { ecmaVersion: 'latest', locations: true });
  const free = new Map();
  const use = (node, chain) => {
    if (chain.some((s) => s.has(node.name))) return;
    if (!free.has(node.name)) free.set(node.name, []);
    free.get(node.name).push(node.loc.start.line);
  };
  const pattern = (p, chain) => {
    if (!p) return;
    if (p.type === 'AssignmentPattern') { pattern(p.left, chain); walk(p.right, chain); }
    else if (p.type === 'ObjectPattern') p.properties.forEach((pr) => { if (pr.type === 'RestElement') pattern(pr.argument, chain); else { if (pr.computed) walk(pr.key, chain); pattern(pr.value, chain); } });
    else if (p.type === 'ArrayPattern') p.elements.forEach((e) => pattern(e, chain));
    else if (p.type === 'RestElement') pattern(p.argument, chain);
    else if (p.type !== 'Identifier') walk(p, chain); // 대입 대상 a.b
  };
  const fn = (node, chain) => {
    const s = new Set();
    if (node.type === 'FunctionExpression' && node.id) s.add(node.id.name);
    if (node.type !== 'ArrowFunctionExpression') s.add('arguments');
    node.params.forEach((p) => patternNames(p, []).forEach((n) => s.add(n)));
    const c2 = [s, ...chain];
    node.params.forEach((p) => pattern(p, c2));
    if (node.body.type === 'BlockStatement') { blockDecls(node.body.body, s); varDecls(node.body, s); node.body.body.forEach((x) => walk(x, c2)); }
    else walk(node.body, c2);
  };
  function walk(node, chain) {
    if (!isNode(node)) return;
    switch (node.type) {
      case 'Identifier': use(node, chain); return;
      case 'FunctionDeclaration': case 'FunctionExpression': case 'ArrowFunctionExpression': fn(node, chain); return;
      case 'ClassDeclaration': case 'ClassExpression': { const s = new Set(node.id ? [node.id.name] : []); walk(node.superClass, chain); walk(node.body, [s, ...chain]); return; }
      case 'MethodDefinition': case 'PropertyDefinition': case 'Property': if (node.computed) walk(node.key, chain); walk(node.value, chain); return;
      case 'MemberExpression': walk(node.object, chain); if (node.computed) walk(node.property, chain); return;
      case 'LabeledStatement': walk(node.body, chain); return;
      case 'BreakStatement': case 'ContinueStatement': case 'MetaProperty': return;
      case 'VariableDeclarator': pattern(node.id, chain); walk(node.init, chain); return;
      case 'AssignmentExpression': pattern(node.left, chain); walk(node.right, chain); return;
      case 'Program': case 'BlockStatement': case 'StaticBlock': {
        const s = new Set(); blockDecls(node.body, s);
        if (node.type === 'Program') varDecls(node, s);
        node.body.forEach((x) => walk(x, [s, ...chain])); return;
      }
      case 'ForStatement': case 'ForInStatement': case 'ForOfStatement': {
        const s = new Set(); const head = node.init ?? node.left;
        if (head?.type === 'VariableDeclaration' && head.kind !== 'var') head.declarations.forEach((d) => patternNames(d.id, []).forEach((n) => s.add(n)));
        for (const k of ['init', 'test', 'update', 'left', 'right', 'body']) walk(node[k], [s, ...chain]);
        return;
      }
      case 'CatchClause': { const s = new Set(patternNames(node.param, [])); pattern(node.param, [s, ...chain]); walk(node.body, [s, ...chain]); return; }
      case 'SwitchStatement': {
        walk(node.discriminant, chain);
        const s = new Set(); node.cases.forEach((c) => blockDecls(c.consequent, s));
        node.cases.forEach((c) => { walk(c.test, [s, ...chain]); c.consequent.forEach((x) => walk(x, [s, ...chain])); });
        return;
      }
    }
    for (const c of children(node)) walk(c, chain);
  }
  walk(ast, []);
  return { ast, free };
}

/** 'use strict'; (() => { … })(); 의 IIFE 몸통 최상위 선언 이름 */
function iifeDecls(ast) {
  const st = ast.body.find((s) => s.type === 'ExpressionStatement' && s.expression.type === 'CallExpression');
  const names = new Set();
  for (const s of st.expression.callee.body.body) {
    if (s.type === 'VariableDeclaration') s.declarations.forEach((d) => patternNames(d.id, []).forEach((n) => names.add(n)));
    else if (s.id) names.add(s.id.name);
  }
  return names;
}

test('범위 분석 자체: 매개변수·구조 분해·catch·블록은 선언, 속성 이름은 아님', () => {
  const { free } = freeVariables('const { a, b: [c] } = x; function f(d = a, ...e) { try { g(d, e, c); } catch (err) { err.y; } { let q = 1; q.z; } return { k: v, w }; } h.i;');
  assert.deepEqual([...free.keys()].sort(), ['g', 'h', 'v', 'w', 'x']);
});

test('bridge: provide 두 번은 오류, 없는 이름 call은 undefined + 이름마다 로그 한 번', () => {
  const logs = []; const b = OKBridge.createBridge({ log: (m) => logs.push(m) });
  b.provide('chat.new', () => 1);  assert.throws(() => b.provide('chat.new', () => 2));
  assert.equal(b.call('chat.new'), 1);
  assert.equal(b.call('home.edit', 'x'), undefined);
  assert.equal(b.call('home.edit'), undefined);
  assert.equal(logs.length, 1);
  assert.equal(b.has('chat.new'), true);
});

test('bridge: 인수를 그대로 넘기고 돌려준 값을 그대로 준다, 없는 이름이 둘이면 로그 둘', () => {
  const logs = []; const b = OKBridge.createBridge({ log: (m) => logs.push(m) });
  b.provide('chat.of', (p, q) => ({ p, q }));
  assert.deepEqual(b.call('chat.of', 'C:\\문서\\a.docx', 2), { p: 'C:\\문서\\a.docx', q: 2 });
  assert.equal(b.has('home.edit'), false);
  b.call('home.edit'); b.call('home.restore'); b.call('home.edit');
  assert.equal(logs.length, 2);
  assert.ok(Object.isFrozen(OKBridge));
});

test('index.html: 스크립트 차례와 defer, app.js 없음', () => {
  const html = read('index.html');
  const scripts = [...html.matchAll(/<script\b([^>]*)>/g)].map((m) => m[1]);
  const srcs = scripts.map((a) => /\bsrc="([^"]+)"/.exec(a)?.[1]);
  assert.deepEqual(srcs, ['../i18n/translate.js', 'md.js', 'dialog-queue.js', 'target-chips.js', 'bridge.js', 'core.js', 'doc-actions.js', 'home.js', 'panel.js']);
  for (const a of scripts) assert.match(a, /\bdefer\b/, a);
  assert.equal(existsSync(join(RENDERER, 'app.js')), false);
});

test('K.init은 core.js에서 한 번만', () => {
  assert.equal(read('core.js').match(/\bK\.init\(/g)?.length ?? 0, 1);
  assert.equal(read('home.js').match(/\bK\.init\(/g)?.length ?? 0, 0);
  assert.equal(read('panel.js').match(/\bK\.init\(/g)?.length ?? 0, 0);
});

test('home.js·panel.js·core.js는 서로의 이름을 직접 쓰지 않는다 (자유 변수 = 허용 전역뿐)', () => {
  const tops = new Set();
  for (const f of ENTRY) {
    const { ast, free } = freeVariables(read(f));
    const bad = [...free].filter(([n]) => !ALLOWED.has(n)).map(([n, lines]) => n + ' (' + lines.join(', ') + '줄)');
    assert.deepEqual(bad, [], f + '의 자유 변수');
    for (const n of iifeDecls(ast)) tops.add(n);
  }
  assert.deepEqual([...tops].filter((n) => ALLOWED.has(n)), []);
});

test('home.js: doc:editor의 aiBusy를 문서 탭 state.doc에 합친다 (AI가 고친 것을 바깥 변경으로 알리지 않게)', () => {
  const home = read('home.js');
  assert.match(home, /typeof ev\.aiBusy === 'boolean'/);
  assert.match(home, /state\.doc\.aiBusy \? 'ai' : 'watch'/);
});

test('state 칸 주인: home.js는 panel 칸을, panel.js는 home 칸을 쓰지 않는다', () => {
  assert.doesNotMatch(read('home.js'), /state\.(chats|targetInfo|targetSeq|demo)\b/);
  assert.doesNotMatch(read('panel.js'), /state\.(preview|previewSeq|rendering|pendingRefresh|zoom|panel|smokeSent)\b/);
});

/** 주석은 빼고 코드의 bridge.provide('…')·bridge.call('…')만 센다 (이름은 글자 그대로여야 한다) */
function bridgeNames(src, verb) {
    const out = new Set();
    const visit = (node) => {
      if (!isNode(node)) return;
      const c = node.type === 'CallExpression' ? node.callee : null;
      if (c?.type === 'MemberExpression' && c.object.name === 'bridge' && c.property.name === verb) {
        assert.equal(node.arguments[0]?.type, 'Literal', 'bridge.' + verb + '의 이름은 글자 그대로');
        out.add(node.arguments[0].value);
      }
      for (const ch of children(node)) visit(ch);
    };
    visit(acorn.parse(src, { ecmaVersion: 'latest' }));
    return out;
}

test('다리: home.js·panel.js의 bridge.call 이름은 같은 파일·core.js에서 provide했거나 VIEW_ROUTES에 있다', () => {
  const [core, docActions, home, panel] = ENTRY.map(read);
  const homeGives = bridgeNames(home, 'provide');
  const panelGives = bridgeNames(panel, 'provide');
  const coreGives = bridgeNames(core, 'provide');
  assert.ok(homeGives.size > 0 && panelGives.size > 0);
  for (const n of homeGives) assert.match(n, /^home\.\w+$/, n);
  for (const n of panelGives) assert.match(n, /^chat\.\w+$/, n);
  const routed = (n) => Object.hasOwn(VIEW_ROUTES, n);
  for (const n of bridgeNames(home, 'call')) assert.ok(homeGives.has(n) || coreGives.has(n) || routed(n), 'home.js가 부르는 ' + n + '을 아무도 받지 않음');
  for (const n of bridgeNames(panel, 'call')) assert.ok(panelGives.has(n) || coreGives.has(n) || routed(n), 'panel.js가 부르는 ' + n + '을 아무도 받지 않음');
  // core.js·doc-actions.js는 홈·패널 어느 뷰에서도 돌므로 중계 표에 있는 이름만 부른다
  for (const n of [...bridgeNames(core, 'call'), ...bridgeNames(docActions, 'call')]) assert.ok(routed(n), 'core.js·doc-actions.js가 부르는 ' + n + '이 VIEW_ROUTES에 없음');
  // 중계 표의 이름은 받는 쪽 파일이 provide한다 (home.* → home.js, chat.* → panel.js)
  for (const [n, r] of Object.entries(VIEW_ROUTES)) assert.ok((r.to === 'home' ? homeGives : panelGives).has(n), n + '을 받는 쪽이 provide하지 않음');
  assert.equal(coreGives.size, 0);
});

test('다리: 값을 돌려받는 이름(chat.of·chat.current)은 부르지 않는다', () => {
  for (const f of ENTRY) {
    const calls = bridgeNames(read(f), 'call');
    assert.equal(calls.has('chat.of'), false, f);
    assert.equal(calls.has('chat.current'), false, f);
  }
});

test('옆 창·대화창 배치 코드 없음: home.js에 onPanel·panel-mode·applyChatLayout, index.html에 옆 창 단추·#splitter 없음', () => {
  const home = read('home.js');
  assert.doesNotMatch(home, /\bonPanel\b|panel-mode|applyChatLayout/);
  const html = read('index.html');
  assert.doesNotMatch(html, /btn-panel-full|btn-panel-close|id="splitter"/);
});

test('bridge: 없는 이름은 relay가 있으면 relay(name, args)로 넘기고 undefined (로그 없음)', () => {
  const logs = []; const relayed = [];
  const b = OKBridge.createBridge({ log: (m) => logs.push(m), relay: (name, args) => relayed.push([name, args]) });
  b.provide('home.pickFile', () => 'here');
  assert.equal(b.call('home.pickFile'), 'here');
  assert.equal(b.call('chat.markStopping', { path: 'C:\\문서\\a.docx' }), undefined);
  assert.deepEqual(relayed, [['chat.markStopping', [{ path: 'C:\\문서\\a.docx' }]]]);
  assert.deepEqual(logs, []);
});

test('화면 뷰 종류: core.js가 ?view=를 OKCore.view·body[data-view]로, home.js는 패널에서·panel.js는 패널 밖에서 아무것도 하지 않는다', () => {
  const core = read('core.js');
  assert.match(core, /new URLSearchParams\(location\.search\)\.get\('view'\)/);
  assert.match(core, /document\.body\.dataset\.view\s*=/);
  assert.match(read('home.js'), /if \(view === 'panel'\) return;/);
  assert.match(read('panel.js'), /if \(view !== 'panel'\) return;/);
  const css = readFileSync(join(RENDERER, 'style.css'), 'utf8');
  assert.match(css, /\[data-view="panel"\]/);
  assert.doesNotMatch(css, /panel-mode|\.panel-only/);
});

test('main.mjs: 전역 current·watcher·pendingFile·rendererReady 선언 없음', () => {
  // 열린 문서·파일 감시·대기 파일·화면 준비는 보낸 쪽 허브(hub)마다 둔다 (탭 창이 여러 개가 될 때 섞이지 않게)
  const main = readFileSync(join(ROOT, 'src', 'app', 'main.mjs'), 'utf8');
  assert.doesNotMatch(main, /^\s{2}let (current|watcher|pendingFile|rendererReady)\b/m);
});
