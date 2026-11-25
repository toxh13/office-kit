// ipc-channels.mjs — IPC 채널 표 하나: 채널마다 호출 종류와 그 채널을 쓸 수 있는 뷰 종류.
// main.mjs·editor-window.mjs는 등록·보내기 때 이 표로 채널을 확인하고, 요청마다 보낸 쪽 종류를 검사한다.
// 주 창(home·panel)의 종류는 그 채널을 실제로 쓰는 화면 파일로 정한다: core.js = home·panel, home.js = home, panel.js = panel.

/** 뷰 종류. tabstrip에 열린 채널은 아직 없다 (M3가 더함) */
export const VIEW_KINDS = Object.freeze(['home', 'panel', 'tabstrip', 'hwp', 'office']);
/** 지금 주 창 (홈 화면과 AI 대화가 한 화면에 있다) */
export const HUB_KINDS = Object.freeze(['home', 'panel']);

const HOME = ['home'];
const PANEL = ['panel'];
const HUB = ['home', 'panel'];
const EDIT = ['hwp', 'office'];

// call: 'invoke'(ipcMain.handle ← ipcRenderer.invoke) | 'on'(ipcMain.on ← ipcRenderer.send/sendSync) | 'send'(webContents.send → ipcRenderer.on)
// open: 보낸 쪽을 가리지 않는 공개 채널 (앱에 든 번역 사전)
const TABLE = {
  // ── 주 창: 시작·언어 ──
  'app:init': { call: 'invoke', kinds: HUB },
  'app:relaunch': { call: 'invoke', kinds: HOME },
  'i18n:get': { call: 'on', kinds: HUB, open: true },
  'smoke:ready': { call: 'on', kinds: HOME },

  // ── 주 창: 문서 ──
  'doc:pick': { call: 'invoke', kinds: HOME },
  'doc:open': { call: 'invoke', kinds: HOME },
  'doc:close': { call: 'invoke', kinds: HOME },
  'doc:preview': { call: 'invoke', kinds: HOME },
  'doc:restore': { call: 'invoke', kinds: HOME },
  'doc:versions': { call: 'invoke', kinds: HOME },
  'doc:diff': { call: 'invoke', kinds: HOME },
  'doc:openInApp': { call: 'invoke', kinds: HOME },
  'doc:openInWord': { call: 'invoke', kinds: HOME },
  'doc:savePdf': { call: 'invoke', kinds: HOME },
  'doc:print': { call: 'invoke', kinds: HOME },
  'doc:showInFolder': { call: 'invoke', kinds: HOME },
  'doc:edit': { call: 'invoke', kinds: HOME },
  'doc:newInfo': { call: 'invoke', kinds: HOME },
  'doc:newCheck': { call: 'invoke', kinds: HOME },
  'doc:pickFolder': { call: 'invoke', kinds: HUB },
  'doc:new': { call: 'invoke', kinds: HOME },

  // ── 주 창: 최근 목록·설정 ──
  'recent:list': { call: 'invoke', kinds: HOME },
  'recent:remove': { call: 'invoke', kinds: HOME },
  'recent:clear': { call: 'invoke', kinds: HOME },
  'recent:pin': { call: 'invoke', kinds: HOME },
  'settings:get': { call: 'invoke', kinds: HOME },
  'settings:set': { call: 'invoke', kinds: HOME },

  // ── 주 창: AI 대화 ──
  'ai:providers': { call: 'invoke', kinds: PANEL },
  'ai:send': { call: 'invoke', kinds: PANEL },
  'ai:cancel': { call: 'invoke', kinds: PANEL },
  'ai:materials': { call: 'invoke', kinds: PANEL },
  'ai:reset': { call: 'invoke', kinds: PANEL },

  // ── 주 창: 연결·한컴·기본 프로그램 ──
  'integrations:status': { call: 'invoke', kinds: HOME },
  'integrations:set': { call: 'invoke', kinds: HOME },
  'hancom:status': { call: 'invoke', kinds: HOME },
  'hancom:setup': { call: 'invoke', kinds: HOME },
  'hancom:remove': { call: 'invoke', kinds: HOME },
  'assoc:status': { call: 'invoke', kinds: HOME },
  'assoc:openSettings': { call: 'invoke', kinds: HOME },

  // ── 주 창: 도구 (개인정보·일괄 변환·PDF) ──
  'pii:scan': { call: 'invoke', kinds: HOME },
  'pii:mask': { call: 'invoke', kinds: HOME },
  'batch:pick': { call: 'invoke', kinds: HOME },
  'batch:pickOut': { call: 'invoke', kinds: HOME },
  'batch:scan': { call: 'invoke', kinds: HOME },
  'batch:start': { call: 'invoke', kinds: HOME },
  'batch:cancel': { call: 'invoke', kinds: HOME },
  'batch:saveReport': { call: 'invoke', kinds: HOME },
  'pdftools:pick': { call: 'invoke', kinds: HOME },
  'pdftools:run': { call: 'invoke', kinds: HOME },

  // ── 주 창: 옆 창 (M3에서 지움: 탭 창 AI 패널 뷰로 바뀜) ──
  'panel:full': { call: 'invoke', kinds: HOME },
  'panel:close': { call: 'invoke', kinds: HOME },

  // ── 주 창: 탐색기·브라우저 ──
  'shell:showItem': { call: 'invoke', kinds: HOME },
  'shell:openBackups': { call: 'invoke', kinds: HOME },
  'shell:openExternal': { call: 'invoke', kinds: PANEL },

  // ── 주 창에 보내기 ──
  'doc:opened': { call: 'send', kinds: HOME },
  'doc:changed': { call: 'send', kinds: HOME },
  'doc:editor': { call: 'send', kinds: HOME },
  'ai:event': { call: 'send', kinds: PANEL },
  'menu': { call: 'send', kinds: HOME },
  'toast': { call: 'send', kinds: HOME },
  'batch:progress': { call: 'send', kinds: HOME },
  // M3에서 지움: 탭 창 AI 패널 뷰로 바뀜
  'panel': { call: 'send', kinds: HOME },

  // ── 편집 창 (한글 = hwp, 문서·표·슬라이드 = office) ──
  'editor:get-document': { call: 'invoke', kinds: EDIT },
  'editor:draft': { call: 'invoke', kinds: EDIT },
  'editor:save': { call: 'invoke', kinds: EDIT },
  'editor:save-as': { call: 'invoke', kinds: EDIT },
  'editor:hwp-pdf': { call: 'invoke', kinds: ['hwp'] },
  'editor:download': { call: 'invoke', kinds: ['office'] },
  'editor:print': { call: 'invoke', kinds: ['office'] },
  'editor:state': { call: 'invoke', kinds: ['office'] },
  'editor:dirty': { call: 'on', kinds: EDIT },
  'editor:log': { call: 'on', kinds: EDIT },
  'editor:close': { call: 'on', kinds: EDIT },
  'editor:reply': { call: 'on', kinds: EDIT },
  'editor:panel': { call: 'on', kinds: EDIT },
  'editor:i18n': { call: 'on', kinds: EDIT, open: true },
  'editor:external-change': { call: 'send', kinds: EDIT },
  'editor:request-close': { call: 'send', kinds: EDIT },
  'editor:command': { call: 'send', kinds: EDIT },
};

/** 채널 → { call, kinds, open? } (얼려 둔다) */
export const IPC = Object.freeze(Object.fromEntries(Object.entries(TABLE).map(([ch, spec]) => [
  ch, Object.freeze({ ...spec, kinds: Object.freeze([...spec.kinds]) }),
])));

/** 표의 한 줄. 없는 채널이나 호출 종류가 다르면 던진다 (등록·보내기 때 바로 드러나게) */
export function channelSpec(channel, call) {
  const spec = Object.hasOwn(IPC, channel) ? IPC[channel] : null;
  if (!spec) throw new Error('IPC 표에 없는 채널: ' + channel);
  if (spec.call !== call) throw new Error('IPC 호출 종류가 다름: ' + channel);
  return spec;
}

/** 보낸 쪽 종류(senderKinds, 모르면 null)가 이 채널을 쓸 수 있는지. 열린 채널은 늘 true */
export function allows(channel, senderKinds) {
  const spec = Object.hasOwn(IPC, channel) ? IPC[channel] : null;
  if (!spec) return false;
  if (spec.open) return true;
  if (!Array.isArray(senderKinds)) return false;
  return spec.kinds.some((k) => senderKinds.includes(k));
}

/** 이 종류들이 쓸 수 있는 채널 (호출 종류마다, 이름 정렬). 열린 채널도 kinds에 없으면 빠진다 */
export function channelsFor(kinds) {
  const out = { invoke: [], on: [], send: [] };
  for (const [ch, spec] of Object.entries(IPC)) {
    if (spec.kinds.some((k) => kinds.includes(k))) out[spec.call].push(ch);
  }
  for (const list of Object.values(out)) list.sort();
  return out;
}

/** preload 파일 → 그 파일이 노출하는 뷰 종류 (블록은 scripts/gen-ipc-preload.mjs가 채운다) */
export const PRELOADS = Object.freeze({
  'src/app/preload.cjs': Object.freeze(['home', 'panel']),
  'src/app/editor/editor-preload.cjs': Object.freeze(['hwp']),
  'src/app/office-editor/office-preload.cjs': Object.freeze(['office']),
});

export const BLOCK_START = '// <ipc-channels> 자동 생성: node scripts/gen-ipc-preload.mjs (손으로 고치지 않는다)';
export const BLOCK_END = '// </ipc-channels>';

/** preload 파일에 들어갈 채널 블록 (BLOCK_START ~ BLOCK_END, 가운데 한 줄). 줄 끝은 eol로 */
export function preloadBlock(kinds, eol = '\n') {
  const c = channelsFor(kinds);
  const set = (list) => `new Set([${list.map((ch) => `'${ch}'`).join(', ')}])`;
  return [
    BLOCK_START,
    `const CHANNELS = { invoke: ${set(c.invoke)}, on: ${set(c.on)}, send: ${set(c.send)} };`,
    BLOCK_END,
  ].join(eol);
}
