// panel.js — AI 패널 뷰(?view=panel): 대상 칩·AI 고르기·대화 그리기·보내기·멈추기·AI 이벤트·패널 너비 조절.
// 대화 대상(state.doc)은 main이 ai:target으로 민다 (그 탭 창의 활성 탭 문서). 홈·문서 탭(home.js, 다른 뷰)은 bridge.call('home.…')로만 부른다 (main 중계).
'use strict';
(() => {
  const {
    K, MD, $, h, icon, spinner, call, pathKey, setBusy, clearBusy,
    toast, showError, confirmDialog, state, bridge, onInit, view,
  } = OKCore;
  // 홈·문서 탭 뷰에서는 아무것도 등록하지 않는다 (home.js 몫)
  if (view !== 'panel') return;

  // panel 칸 (home 칸은 읽지도 쓰지도 않는다)
  Object.assign(state, {
    chats: new Map(),       // 문서 경로 → { items, busy, chatId, materials, outDir }
    targetInfo: null,       // 마지막 ai:materials 결과 { edit, open }
    targetSeq: 0,
    demo: false,
  });

  const el = {
    chatBody: $('#chat-body'), chatLog: $('#chat-log'), provider: $('#provider'),
    composer: $('#composer'), targetChips: $('#target-chips'), prompt: $('#prompt'), send: $('#btn-send'), composerHint: $('#composer-hint'),
    editorWarn: $('#editor-warn'), resizer: $('#panel-resizer'),
  };

  // AI 도구 이름 → 화면에 보일 말
  const TOOL_LABEL = {
    office_status: ['상태 확인 중', '상태 확인'],
    office_info: ['문서 정보 보는 중', '문서 정보 보기'],
    office_read: ['문서 읽는 중', '문서 읽기'],
    office_fields: ['채울 칸 찾는 중', '채울 칸 찾기'],
    office_fill: ['칸 채우는 중', '칸 채우기'],
    office_replace: ['바꾸는 중', '찾아 바꾸기'],
    office_set_cells: ['셀 고치는 중', '셀 고치기'],
    office_convert: ['변환 중', '형식 바꾸기'],
    office_create: ['새 문서 만드는 중', '새 문서 만들기'],
    office_page_image: ['결과 확인 중', '결과 확인'],
    office_find_pii: ['개인정보 찾는 중', '개인정보 찾기'],
    office_mask_pii: ['개인정보 가리는 중', '개인정보 가린 사본'],
    office_hwpx_batch: ['HWPX로 바꾸는 중', 'HWPX 일괄 변환'],
    hwp_edit: ['한글 편집 중', '한글 편집'],
    word_edit: ['Word 편집 중', 'Word 편집'],
    ppt_edit: ['슬라이드 편집 중', '슬라이드 편집'],
    excel_edit: ['시트 편집 중', '시트 편집'],
    hwp_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    word_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    ppt_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
    excel_inspect: ['문서 구조 살피는 중', '문서 구조 살펴보기'],
  };
  const EDIT_TOOLS = new Set(['office_fill', 'office_replace', 'office_set_cells', 'hwp_edit', 'office_create', 'office_convert',
    'word_edit', 'ppt_edit', 'excel_edit']);
  // mcp__office-kit__office_read, office-kit.office_read 같은 이름도 받아 준다
  const toolBase = (name) => String(name ?? '').split(/__|[./:]/).filter(Boolean).pop() ?? '';
  /** part = { tool, name } (tool: office_read 같은 원래 이름, name: ai.mjs가 붙인 표시 이름) */
  function toolLabel(part, done) {
    const base = toolBase(part.tool || part.name);
    const l = TOOL_LABEL[base];
    if (l) return done ? l[1] : l[0];
    const label = part.name && part.name !== base ? part.name : base || '도구';
    return done ? label : `${label} 중`;
  }

  // ── AI 대화 ──────────────────────────────────────────

  function chatFor(path) {
    const k = pathKey(path);
    let c = state.chats.get(k);
    if (!c) { c = { items: [], busy: false, chatId: null, materials: [], outDir: null }; state.chats.set(k, c); }
    return c;
  }

  // ── AI 대상 칩 (고칠 문서 + 읽기 재료) ─────────────────

  /** 열린 문서 목록을 새로 받아 칩을 다시 그린다. 받는 사이 문서가 바뀌었으면 버린다. */
  async function refreshTargetInfo() {
    const seq = ++state.targetSeq;
    if (!state.doc) { state.targetInfo = null; renderTargetChips(); return; }
    try {
      const info = await call(K.ai.materials());
      if (seq !== state.targetSeq) return;
      state.targetInfo = info;
    } catch { return; }
    renderTargetChips();
  }

  function targetModel() {
    const chat = chatCurrent();
    if (!state.targetInfo || !chat) return null;
    return OKTargetChips.model({ info: state.targetInfo, materials: chat.materials, outDir: chat.outDir });
  }

  function renderTargetChips() {
    const root = el.targetChips;
    const m = state.doc && !chatBlocker() ? targetModel() : null;
    root.hidden = !m;
    if (!m) { root.replaceChildren(); return; }
    OKTargetChips.render(root, m, {
      h, icon,
      onRemove(path) {
        const chat = chatCurrent();
        const gone = chat.materials.find((x) => pathKey(x.path) === pathKey(path));
        chat.materials = chat.materials.filter((x) => pathKey(x.path) !== pathKey(path));
        // 인터넷 재료를 뺀 뒤에는 main이 다음 보내기에서 세션을 지운다
        if (gone?.zone3 && chat.items.length) chat.items.push({ role: 'note', text: '자료가 바뀌어 다음 메시지부터 새 대화로 시작해요.' });
        renderChat();
      },
      onAdd(path) {
        const chat = chatCurrent();
        const m2 = state.targetInfo?.open.find((x) => pathKey(x.path) === pathKey(path));
        if (m2 && !chat.materials.some((x) => pathKey(x.path) === pathKey(path))) chat.materials = [...chat.materials, m2];
        renderChat();
      },
      onOpenMenu() { refreshTargetInfo(); },
      async onPickOutDir() {
        const chat = chatCurrent();
        const cur = targetModel();
        let folder = null;
        try { folder = await call(K.doc.pickFolder(cur?.outDir ?? undefined)); } catch (err) { showError(err); return; }
        if (folder && chatCurrent() === chat) { chat.outDir = folder; renderTargetChips(); }
      },
    });
  }

  async function loadProviders(refresh = false) {
    try {
      state.providers = await call(K.ai.providers({ refresh }));
      state.providerError = null;
    } catch (err) {
      state.providers = null;
      state.providerError = err.message;
    }
    renderProviderSelect();
    renderChat();
    bridge.call('home.renderNewDoc');
  }

  function renderProviderSelect() {
    const sel = el.provider;
    const p = state.providers;
    sel.replaceChildren();
    if (!p || !p.providers.length) {
      sel.append(h('option', { value: '', text: state.providerError ? 'AI 없음' : '확인 중…' }));
      sel.disabled = true;
      return;
    }
    for (const it of p.providers) {
      sel.append(h('option', { value: it.id, text: it.available ? it.short : `${it.short} (설치 안 됨)`, title: it.name, disabled: !it.available }));
    }
    sel.value = p.selected ?? '';
    sel.disabled = !p.selected || chatCurrent()?.busy;
  }

  const chatCurrent = () => (state.doc ? chatFor(state.doc.path) : null);

  /** 대화창을 쓸 수 없는 이유 카드 (없으면 null) */
  function chatBlocker() {
    if (state.demo) return null;
    if (!state.doc) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('file')),
        h('h3', { text: '문서를 먼저 열어 주세요' }),
        h('p', { text: '문서를 열면 AI가 내용을 읽고 요약하거나, 글자를 바꾸고, 빈칸을 채워 드려요. 고친 결과는 왼쪽 미리보기에 바로 보여요.' }),
        h('p', { text: '새 문서를 만들면서 AI에게 내용을 부탁할 수도 있어요.' }),
        h('div', { class: 'btn-row' },
          h('button', { class: 'btn sm', onclick: () => bridge.call('home.pickFile') }, icon('file-open'), '파일 열기'),
          h('button', { class: 'btn sm', id: 'chat-new-doc', onclick: () => bridge.call('home.openNewDoc', { start: 'ai' }) }, icon('file-new'), '새 문서')));
    }
    if (state.providerError) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('alert')),
        h('h3', { text: 'AI 기능을 쓸 수 없어요' }),
        h('p', { text: state.providerError }),
        h('button', { class: 'btn sm', onclick: () => loadProviders(true) }, icon('refresh'), '다시 확인'));
    }
    if (!state.providers) return h('div', { class: 'loading-line' }, spinner(), 'AI 프로그램을 확인하는 중…');
    if (!state.providers.selected) {
      return h('div', { class: 'chat-card' },
        h('div', { class: 'c-icon' }, icon('sparkle')),
        h('h3', { text: 'AI 프로그램이 필요해요' }),
        h('p', { text: 'Office Kit는 이 컴퓨터에 설치하고 로그인한 Claude Code 또는 Codex로 대화합니다. 아래 안내대로 하나를 설치한 뒤 다시 확인을 눌러 주세요.' }),
        h('div', { class: 'hint-list' }, ...state.providers.providers.map((p) =>
          h('div', { class: 'hint-item' }, h('b', { text: p.name }), h('span', { text: p.hint || '설치되어 있지 않아요.' })))),
        h('button', { class: 'btn sm', onclick: () => loadProviders(true) }, icon('refresh'), '다시 확인'));
    }
    return null;
  }

  // 첫 화면 추천 문장: 열린 문서 종류에 맞춰 (형식별 전용 기능)
  const SUGGESTS = {
    한글: ['이 문서 요약해줘', '표 내용 정리해줘', '머리말 넣어줘', '맞춤법 고쳐줘'],
    Word: ['이 문서 요약해줘', '제목 스타일 정리해줘', '목차 넣어줘', '맞춤법 고쳐줘'],
    PowerPoint: ['발표자 노트 달아줘', '슬라이드 하나 추가해줘', '이 발표 요약해줘', 'PDF로 저장해줘'],
    Excel: ['차트 만들어줘', '머리글 굵게 하고 틀 고정해줘', '이 표 요약해줘', 'PDF로 저장해줘'],
    PDF: ['이 문서 요약해줘', '중요한 날짜·금액만 뽑아줘', '표 내용 정리해줘', '개인정보가 있는지 찾아줘'],
  };
  const SUGGESTS_DEFAULT = ['이 문서 요약해줘', '맞춤법 고쳐줘', '표 내용 정리해줘', 'PDF로 저장해줘'];
  const suggestsFor = (doc) => SUGGESTS[doc?.kind] ?? SUGGESTS_DEFAULT;

  /** 고른 AI가 로그인되지 않은 것 같으면 안내 (막지는 않음) */
  function loginWarning() {
    const p = state.providers?.providers.find((x) => x.id === state.providers.selected);
    if (!p || p.loggedIn !== false || !p.hint) return null;
    return h('div', { class: 'warn-note' }, icon('alert'), h('span', { text: p.hint }));
  }

  function renderChat() {
    const blocker = chatBlocker();
    const chat = chatCurrent();
    const log = el.chatLog;
    if (blocker) {
      log.replaceChildren(blocker);
    } else if (!chat.items.length) {
      log.replaceChildren(h('div', { class: 'chat-welcome' },
        h('div', { class: 'w-icon' }, icon('sparkle')),
        h('h3', { text: '무엇을 도와드릴까요?' }),
        h('p', { text: state.doc.kind === 'PDF'
          ? `「${state.doc.name}」 문서를 읽고 요약·정리해 드려요. PDF는 내용을 바꿀 수는 없어요.`
          : `「${state.doc.name}」 문서를 읽고 고쳐 드려요. 바꾼 내용은 왼쪽 미리보기에 바로 반영되고, 언제든 원래대로 되돌릴 수 있어요.` }),
        loginWarning(),
        h('div', { class: 'suggests' }, ...suggestsFor(state.doc).map((s) => h('button', { class: 'suggest', onclick: () => sendPrompt(s) }, icon('sparkle'), s)))));
    } else {
      log.replaceChildren(...chat.items.map(messageNode));
    }
    const off = !!blocker;
    el.composer.classList.toggle('disabled', off);
    el.prompt.disabled = off;
    el.prompt.placeholder = off ? (state.doc ? 'AI 프로그램을 먼저 준비해 주세요' : '문서를 열면 대화할 수 있어요') : '문서에 대해 물어보거나 고칠 내용을 적어 주세요';
    $('#btn-new-chat').disabled = off || !chat?.items.length || chat?.busy;
    renderSendButton();
    renderTargetChips();
    renderEditorWarn();
    scrollChat(true);
  }

  /** 대화창 위 경고: 편집 탭에 저장 안 한 내용이 있을 때 */
  function renderEditorWarn() {
    el.editorWarn.hidden = !(state.doc?.editor?.open && state.doc.editor.dirty);
  }

  /** AI 대상 문서의 편집 탭으로 (문서 탭이면 편집 탭으로 바뀐다) */
  function showEditor() {
    call(K.ai.showEditor()).catch(showError);
  }

  function renderSendButton() {
    const chat = chatCurrent();
    const busy = !!chat?.busy;
    el.send.classList.toggle('stop', busy);
    el.send.replaceChildren(icon(busy ? 'stop' : 'send'));
    el.send.title = busy ? '멈추기' : '보내기 (Enter)';
    el.send.setAttribute('aria-label', busy ? '멈추기' : '보내기');
    el.send.disabled = !busy && (el.prompt.disabled || !el.prompt.value.trim());
    el.composerHint.textContent = busy ? 'AI가 작업 중이에요 · 멈추려면 ■ 버튼' : 'Enter 보내기 · Shift+Enter 줄바꿈';
    el.provider.disabled = busy || !state.providers?.selected;
  }

  function messageNode(item) {
    if (item.role === 'user') {
      return h('div', { class: 'msg user' }, h('div', { class: 'bubble', text: item.text }));
    }
    if (item.role === 'note') return h('div', { class: 'chat-divider', role: 'note', text: item.text });
    const body = h('div', { class: 'msg-body' });
    fillAssistant(body, item);
    item.node = body;
    return h('div', { class: 'msg assistant' }, h('div', { class: 'avatar', 'aria-hidden': 'true' }, icon('sparkle')), body);
  }

  function fillAssistant(body, item) {
    const nodes = [];
    let chips = null;
    for (const part of item.parts) {
      if (part.type === 'tool') {
        if (!chips) { chips = h('div', { class: 'tools' }); nodes.push(chips); }
        const st = part.state;
        const chip = h('span', { class: `chip ${st}`, title: part.summary || toolBase(part.tool || part.name) },
          st === 'run' ? spinner() : icon(st === 'ok' ? 'check' : st === 'fail' ? 'alert' : 'x'),
          h('span', { class: 'chip-label', text: toolLabel(part, st !== 'run') + (st === 'fail' ? ' 실패' : st === 'stop' ? ' (멈춤)' : '') }));
        chips.append(chip);
      } else if (part.text.trim()) {
        chips = null;
        nodes.push(h('div', { class: 'md', html: MD.render(part.text) }));
      }
    }
    // 마지막이 돌고 있는 도구면 칩이 이미 진행 중임을 보여 준다
    const lastPart = item.parts.at(-1);
    if (!item.done && !(lastPart?.type === 'tool' && lastPart.state === 'run' && !item.stopping)) {
      nodes.push(h('div', { class: 'thinking' }, h('span', { class: 'dots' }, h('i'), h('i'), h('i')), item.status || '생각하는 중…'));
    }
    if (item.error) nodes.push(h('div', { class: 'msg-error', role: 'alert', title: item.errorDetail || null }, icon('alert'), h('span', { text: item.error })));
    if (item.canceled && !item.error) nodes.push(h('div', { class: 'msg-note', text: '멈췄어요.' }));
    else if (item.done && item.stopping && !item.error) nodes.push(h('div', { class: 'msg-note', text: '멈추기 전에 작업이 끝났어요.' }));
    if (item.done && !item.error && !item.canceled && !item.parts.some((p) => p.type === 'text' && p.text.trim())) {
      nodes.push(h('div', { class: 'msg-note', text: '작업을 마쳤어요.' }));
    }
    if (item.changed) {
      const ch = item.changed;
      const here = () => state.doc && pathKey(state.doc.path) === pathKey(ch.path);
      nodes.push(h('div', { class: 'msg-change' },
        h('span', { class: 'msg-change-text' }, icon('check'), '이 대화에서 문서를 고쳤어요'),
        h('button', { class: 'btn sm', text: '바뀐 내용 보기', onclick: () => { if (here()) OKDocActions.showDiff(ch.path, ch.before, { label: 'AI 대화 전', ai: true }); } }),
        h('button', { class: 'btn sm ghost', text: '이 대화 전으로 되돌리기', onclick: () => { if (here()) OKDocActions.restore(ch.path, ch.before); } })));
    }
    body.replaceChildren(...nodes);
  }

  let chatFrame = 0;
  function updateAssistant(item) {
    if (!item.node || !item.node.isConnected) return;
    if (chatFrame) return;
    // rAF는 창이 가려지면 멈추므로 타이머로 모아서 그린다
    chatFrame = setTimeout(() => {
      chatFrame = 0;
      const nearBottom = isNearBottom();
      const chat = chatCurrent();
      if (chat) for (const it of chat.items) if (it.role === 'assistant' && it.dirty && it.node?.isConnected) { it.dirty = false; fillAssistant(it.node, it); }
      if (nearBottom) scrollChat(true);
    }, 33);
  }

  const isNearBottom = () => el.chatBody.scrollHeight - el.chatBody.scrollTop - el.chatBody.clientHeight < 80;
  function scrollChat(force) {
    if (force || isNearBottom()) el.chatBody.scrollTop = el.chatBody.scrollHeight;
  }

  /**
   * 대화 한 번 준비: 사용자 말풍선과 답 자리를 넣고 '작업 중'으로 둔다 → { chat, msg }.
   * 대화 번호를 받으면 chatStarted, 못 받으면 chatFailed (번호보다 먼저 온 이벤트는 chat.pending이 받는다)
   */
  function queueChat(path, text) {
    const chat = chatFor(path);
    const msg = { role: 'assistant', parts: [], done: false, status: '생각하는 중…', chatId: null };
    chat.items.push({ role: 'user', text }, msg);
    chat.busy = true;
    chat.pending = msg;
    chat.stopRequested = false;
    chat.cancelSent = null;
    setBusy(`ai:${pathKey(path)}`, 'AI가 작업 중…');
    if (chatCurrent() === chat) renderChat();
    return { chat, msg };
  }

  function chatStarted(chat, msg, chatId) {
    msg.chatId = chatId;
    if (chat.pending === msg) chat.pending = null;
    // 이미 끝났으면('end'가 먼저 옴) 대화 번호를 남기지 않는다
    if (!msg.done) chat.chatId = chatId;
    // 보내는 동안 멈추기를 눌렀으면 지금 멈춘다
    if (chat.stopRequested && !msg.done) cancelChatId(chat, chatId);
  }

  function chatFailed(chat, msg, path, err) {
    if (chat.pending === msg) chat.pending = null;
    msg.done = true;
    msg.error = err?.message || 'AI 대화를 시작하지 못했어요';
    chat.busy = false;
    clearBusy(`ai:${pathKey(path)}`);
    if (err?.code === 'NO_PROVIDER' || err?.code === 'MODULE_MISSING') loadProviders(true);
    if (chatCurrent() === chat) renderChat();
  }

  /** 편집 창에 저장 안 한 내용이 있을 때 보내기 전에 묻는다 → 'send' | 'editor' | 'cancel' */
  async function askEditorDirty() {
    const r = await confirmDialog({
      title: '편집 창에 저장하지 않은 내용이 있어요',
      tone: 'warn',
      body: [
        'AI는 저장된 파일을 고칩니다. 지금 보내면 편집 창에서 고친 내용은 AI가 보지 못하고, 나중에 편집 창에서 저장하면 AI가 고친 내용을 덮어쓸 수 있어요.',
        '편집 창에서 먼저 저장(Ctrl+S)한 뒤 보내 주세요.',
      ],
      ok: '편집 창 보기', alt: '그대로 보내기', cancel: '취소',
    });
    return r === true ? 'editor' : r === 'alt' ? 'send' : 'cancel';
  }

  async function sendPrompt(text, { editorOk = false, materialsOk = false, saveMaterials = false } = {}) {
    text = String(text ?? '').trim();
    if (!text || !state.doc || chatBlocker()) return;
    const path = state.doc.path;
    if (chatFor(path).busy) return;
    // 편집 창에 저장 안 한 내용이 있으면 먼저 묻는다 (입력한 글은 그대로 둔다)
    if (!editorOk && state.doc.editor?.dirty) {
      const choice = await askEditorDirty();
      if (choice === 'editor') { showEditor(); return; }
      if (choice !== 'send' || !state.doc || pathKey(state.doc.path) !== pathKey(path) || chatFor(path).busy) return;
      editorOk = true;
    }
    // 보내기 직전에 열린 문서 목록을 새로 받아 대상을 다시 만든다 (받지 못하면 마지막 값)
    try { state.targetInfo = await call(K.ai.materials()); } catch {}
    if (!state.doc || pathKey(state.doc.path) !== pathKey(path) || chatFor(path).busy) return;
    const tm = targetModel();
    if (!tm) { toast('대화 대상을 확인하지 못했어요. 잠시 뒤 다시 시도해 주세요', { type: 'warn' }); return; }
    if (tm.busyMaterial) {
      toast(`‘${tm.busyMaterial.name}’은(는) AI가 고치는 중이라 재료로 넘길 수 없어요`, { type: 'warn' });
      renderTargetChips();
      return;
    }
    el.prompt.value = '';
    autoGrow();
    const { chat, msg } = queueChat(path, text);
    try {
      const sent = await call(K.ai.send({
        prompt: text, editorOk, target: OKTargetChips.targetOf(tm),
        outDir: tm.materialOnly ? tm.outDir : undefined, materialsOk, saveMaterials,
      }));
      chatStarted(chat, msg, sent.chatId);
    } catch (err) {
      if (err.code === 'MATERIAL_DIRTY') {
        // 저장 안 한 재료: 대화를 빼고 글을 되돌린 뒤 묻는다 (EDITOR_DIRTY와 같은 방식)
        dropQueuedChat(path, { chat, msg });
        renderChat();
        if (!el.prompt.value.trim()) { el.prompt.value = text; autoGrow(); }
        const r = await confirmDialog({
          title: '저장 안 한 재료가 있어요', tone: 'warn', body: [err.message],
          ok: '저장하고 보내기', alt: '그대로 보내기', cancel: '취소',
        });
        if (r === true) sendPrompt(text, { editorOk, saveMaterials: true });
        else if (r === 'alt') sendPrompt(text, { editorOk, materialsOk: true });
        return;
      }
      if (['MATERIAL_BUSY', 'MATERIAL_NOT_OPEN', 'TARGET_CHANGED', 'MATERIAL_SAVE_FAILED', 'BAD_OUT_DIR', 'AI_BUSY'].includes(err.code)) {
        chatFailed(chat, msg, path, err);
        refreshTargetInfo();
        return;
      }
      if (err.code !== 'EDITOR_DIRTY') { chatFailed(chat, msg, path, err); return; }
      // 화면이 편집 창 상태를 아직 몰랐던 경우: 대화를 빼고 글을 되돌린 뒤 묻는다
      dropQueuedChat(path, { chat, msg });
      renderChat();
      if (!el.prompt.value.trim()) { el.prompt.value = text; autoGrow(); }
      if (state.doc && pathKey(state.doc.path) === pathKey(path)) {
        state.doc = { ...state.doc, editor: { open: true, dirty: true } };
        renderEditorWarn();
        bridge.call('home.docEditor', { path, open: true, dirty: true });
      }
      const choice = await askEditorDirty();
      if (choice === 'editor') showEditor();
      else if (choice === 'send') sendPrompt(text, { editorOk: true });
    }
  }

  function findMessage(chat, chatId) {
    for (let i = chat.items.length - 1; i >= 0; i--) {
      const it = chat.items[i];
      if (it.role === 'assistant' && it.chatId === chatId) return it;
    }
    // 응답보다 이벤트가 먼저 온 경우 (그 사이 멈추기를 눌렀으면 바로 멈춘다)
    if (chat.pending) {
      const m = chat.pending;
      m.chatId = chatId;
      chat.pending = null;
      if (chat.stopRequested) cancelChatId(chat, chatId);
      return m;
    }
    return null;
  }

  function onAiEvent(ev) {
    if (!ev || !ev.path) return;
    // main이 시작한 대화(새 문서 AI 시작): 모르는 대화 번호면 대화를 만든다 (다른 이벤트보다 먼저 온다)
    if (ev.type === 'start') {
      const c = chatFor(ev.path);
      if (c.items.some((i) => i.chatId === ev.chatId)) return;
      const { msg } = queueChat(ev.path, String(ev.prompt ?? ''));
      chatStarted(c, msg, ev.chatId);
      return;
    }
    const chat = chatFor(ev.path);
    const msg = findMessage(chat, ev.chatId);
    const isCurrent = state.doc && pathKey(state.doc.path) === pathKey(ev.path);
    if (!msg) return;
    switch (ev.type) {
      case 'text': {
        const last = msg.parts.at(-1);
        const text = String(ev.text ?? '');
        if (ev.delta === false) {
          if (last?.type === 'text' && text.startsWith(last.text)) last.text = text;
          else msg.parts.push({ type: 'text', text });
        } else if (last?.type === 'text') last.text += text;
        else msg.parts.push({ type: 'text', text });
        msg.status = '답하는 중…';
        break;
      }
      case 'text_reset': {
        // 스트림이 끊겨 AI가 다시 쓰는 경우: 이미 보인 글 끝에서 drop 글자만큼 지운다
        let drop = Math.max(0, Number(ev.drop) || 0);
        for (let i = msg.parts.length - 1; i >= 0 && drop > 0; i--) {
          const p = msg.parts[i];
          if (p.type !== 'text') continue;
          const cut = Math.min(drop, p.text.length);
          p.text = p.text.slice(0, p.text.length - cut);
          drop -= cut;
          if (!p.text) msg.parts.splice(i, 1);
        }
        break;
      }
      case 'tool': {
        const part = { type: 'tool', id: ev.id, tool: String(ev.tool ?? ''), name: String(ev.name ?? ''), state: 'run', summary: '' };
        msg.parts.push(part);
        msg.status = `${toolLabel(part, false)}…`;
        break;
      }
      case 'tool_result': {
        const part = msg.parts.find((p) => p.type === 'tool' && p.id === ev.id && p.state === 'run')
          ?? msg.parts.find((p) => p.type === 'tool' && p.id === ev.id);
        if (part) {
          part.state = ev.ok === false ? 'fail' : 'ok';
          part.summary = String(ev.summary ?? '');
          if (part.state === 'ok' && EDIT_TOOLS.has(toolBase(part.tool || part.name))) bridge.call('home.refreshPreview', { path: ev.path, ifChanged: true, reason: 'ai' });
        }
        msg.status = '생각하는 중…';
        break;
      }
      case 'status':
        if (ev.message) msg.status = String(ev.message);
        break;
      case 'done':
        msg.done = true;
        if (ev.cancelled) msg.canceled = true;
        if (ev.text && !msg.parts.some((p) => p.type === 'text' && p.text.trim())) msg.parts.push({ type: 'text', text: String(ev.text) });
        break;
      case 'error':
        msg.error = String(ev.message ?? '오류가 났습니다');
        msg.errorDetail = ev.detail ? String(ev.detail).slice(-600) : '';
        msg.done = true;
        break;
      case 'end':
        msg.done = true;
        // AI가 이 대화에서 문서를 고쳤다 → 바뀐 내용 보기 · 이 대화 전으로 되돌리기
        if (ev.changed && typeof ev.before === 'string') msg.changed = { before: ev.before, path: ev.path };
        for (const p of msg.parts) if (p.type === 'tool' && p.state === 'run') p.state = msg.canceled ? 'stop' : 'ok';
        if (chat.chatId === ev.chatId) chat.chatId = null;
        chat.busy = false;
        clearBusy(`ai:${pathKey(ev.path)}`);
        refreshTargetInfo();
        // 그 문서를 보는 문서 탭의 미리보기를 새로 그린다 (main이 경로로 고른다)
        bridge.call('home.refreshPreview', { path: ev.path, ifChanged: true, reason: 'ai' });
        if (isCurrent) {
          renderChat();
          return;
        }
        break;
      default:
        return;
    }
    // 멈추기를 누른 뒤에는 다른 상태 글로 덮지 않는다
    if (msg.stopping && !msg.done) msg.status = '멈추는 중…';
    msg.dirty = true;
    if (isCurrent) updateAssistant(msg);
  }

  /** 멈추기 요청만 표시한다. '멈췄어요'는 AI가 실제로 멈췄다고 알려 올 때(done.cancelled) 붙인다 */
  function markStopping(chat) {
    chat.stopRequested = true;
    const msg = chat.items.findLast((i) => i.role === 'assistant');
    if (msg && !msg.done) { msg.stopping = true; msg.status = '멈추는 중…'; msg.dirty = true; updateAssistant(msg); }
  }

  /** 대화 하나에 멈추기를 한 번만 보낸다 */
  async function cancelChatId(chat, chatId) {
    if (!chatId || chat.cancelSent === chatId) return;
    chat.cancelSent = chatId;
    try { await call(K.ai.cancel(chatId)); } catch (err) { showError(err); }
  }

  async function stopChat() {
    const chat = chatCurrent();
    if (!chat?.busy) return;
    markStopping(chat);
    // 아직 대화 번호를 못 받았으면 받는 즉시 sendPrompt가 멈춘다
    if (chat.chatId) await cancelChatId(chat, chat.chatId);
  }

  async function newChat() {
    const chat = chatCurrent();
    if (!chat || chat.busy) return;
    try { await call(K.ai.reset()); } catch {}
    chat.items = [];
    renderChat();
    el.prompt.focus();
  }

  function autoGrow() {
    const t = el.prompt;
    t.style.height = 'auto';
    t.style.height = `${Math.min(t.scrollHeight, 180)}px`;
    renderSendButton();
  }

  // 확인용: 가짜 대화 내용
  function injectDemoChat() {
    if (!state.doc) return;
    state.demo = true;
    const chat = chatFor(state.doc.path);
    chat.items = [
      { role: 'user', text: '담당자 칸에 김민지 넣고, 표 숫자 단위도 맞춰줘' },
      { role: 'assistant', done: true, parts: [
        { type: 'tool', id: '1', tool: 'office_read', name: '문서 읽기', state: 'ok' },
        { type: 'tool', id: '2', tool: 'office_fill', name: '칸 채우기', state: 'ok' },
        { type: 'tool', id: '3', tool: 'office_replace', name: '찾아 바꾸기', state: 'ok' },
        { type: 'text', text: '**담당자** 칸에 `김민지`를 넣었어요.\n\n- 표의 금액을 모두 **만 원** 단위로 맞췄어요 (2곳)\n- 나머지 서식은 그대로 두었어요\n\n왼쪽 미리보기에서 바뀐 부분을 확인해 보세요.' },
      ], changed: { before: '', path: state.doc.path } },
      { role: 'user', text: '결과 한번 확인해 줘' },
      { role: 'assistant', done: false, status: '결과 확인 중…', parts: [
        { type: 'tool', id: '4', tool: 'office_page_image', name: '쪽 이미지로 확인', state: 'run' },
      ] },
    ];
    chat.busy = true;
    renderChat();
    renderProviderSelect();
  }

  /** 문서를 만들지 못했을 때 미리 넣은 대화를 뺀다 */
  function dropQueuedChat(path, { chat, msg }) {
    const i = chat.items.indexOf(msg);
    if (i > 0) chat.items.splice(i - 1, 2);
    if (chat.pending === msg) chat.pending = null;
    chat.busy = false;
    clearBusy(`ai:${pathKey(path)}`);
    if (!chat.items.length) state.chats.delete(pathKey(path));
    if (chatCurrent() === chat) renderChat();
  }

  /** 같은 경로에 새 문서를 다시 만들 때: 쉬고 있는 예전 대화를 잊는다 */
  function forgetIdleChat(path) {
    const old = state.chats.get(pathKey(path));
    if (old && !old.busy) state.chats.delete(pathKey(path));
  }

  // ── 연결 ─────────────────────────────────────────────

  function bindUi() {
    $('#btn-new-chat').addEventListener('click', newChat);
    $('#editor-warn-show').addEventListener('click', showEditor);
    setupResizer();

    // 대화 입력
    el.prompt.addEventListener('input', autoGrow);
    el.prompt.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && e.keyCode !== 229) {
        e.preventDefault();
        if (!chatCurrent()?.busy) sendPrompt(el.prompt.value);
      }
    });
    el.send.addEventListener('click', () => (chatCurrent()?.busy ? stopChat() : sendPrompt(el.prompt.value)));
    el.provider.addEventListener('change', async () => {
      if (await bridge.call('home.saveSettings', { provider: el.provider.value || null })) loadProviders(false);
    });
    // 답변 속 링크는 기본 브라우저로
    el.chatLog.addEventListener('click', (e) => {
      const a = e.target.closest('a.ext');
      if (!a) return;
      e.preventDefault();
      call(K.shell.openExternal(a.dataset.href)).catch(showError);
    });
  }

  /** main이 민 AI 대상 (그 탭 창의 활성 탭 문서, 홈이면 null) */
  function setTarget(doc) {
    const before = state.doc?.path ?? null;
    state.doc = doc ?? null;
    if (pathKey(before) !== pathKey(state.doc?.path)) state.targetInfo = null;
    // 확인용 캡처: 대상이 생기면 가짜 대화를 넣는다
    if (state.smoke?.chat && state.doc && !state.demo) { injectDemoChat(); return; }
    renderProviderSelect();
    renderChat();
    refreshTargetInfo();
  }

  // ── 패널 너비 (왼쪽 가장자리, role=separator) ── 끌기·←/→ 키. main이 탭 창 배치를 바꾸고 설정에 남긴다
  let panelWidth = 400;
  async function setPanelWidth(w) {
    try { panelWidth = (await call(K.tabs.panelWidth(Math.round(w)))).width ?? panelWidth; } catch {}
    el.resizer.setAttribute('aria-valuenow', String(panelWidth));
  }
  function setupResizer() {
    const r = el.resizer;
    let startX = 0; let startW = 0; let dragging = false;
    r.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      // 뷰 자체가 움직이므로 화면 좌표(screenX)로 잰다
      dragging = true; startX = e.screenX; startW = panelWidth;
      r.setPointerCapture(e.pointerId);
      document.body.classList.add('resizing');
      e.preventDefault();
    });
    r.addEventListener('pointermove', (e) => {
      if (!dragging) return;
      void setPanelWidth(startW + (startX - e.screenX));
    });
    const end = (e) => {
      if (!dragging) return;
      dragging = false;
      try { r.releasePointerCapture(e.pointerId); } catch {}
      document.body.classList.remove('resizing');
    };
    r.addEventListener('pointerup', end);
    r.addEventListener('pointercancel', end);
    r.addEventListener('dblclick', () => setPanelWidth(400));
    r.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      e.preventDefault();
      // 왼쪽 가장자리: ← = 넓게, → = 좁게
      void setPanelWidth(panelWidth + (e.key === 'ArrowLeft' ? 24 : -24));
    });
  }

  function bindEvents() {
    K.on.aiEvent(onAiEvent);
    K.on.aiTarget((t) => setTarget(t?.doc ?? null));
  }

  // 홈·문서 탭(다른 뷰)이 부르는 이름 — main이 VIEW_ROUTES대로 중계한다 (ipc-channels.mjs)
  bridge.provide('chat.markStopping', ({ path } = {}) => { if (path) markStopping(chatFor(path)); });
  bridge.provide('chat.restored', ({ path, newChat: fresh } = {}) => {
    // main이 이 문서의 AI 대화를 끝냈다 → 대화창에 알린다 (AI가 되돌리기 전 내용을 기억하지 않게)
    const c = path ? state.chats.get(pathKey(path)) : null;
    if (!fresh || !c?.items.length) return;
    c.items.push({ role: 'note', text: '문서를 되돌려서, 다음 메시지부터 새 대화로 시작해요.' });
    if (chatCurrent() === c) renderChat();
  });
  bridge.provide('chat.forgetIdle', ({ path } = {}) => { if (path) forgetIdleChat(path); });
  bridge.provide('chat.loadProviders', (refresh) => loadProviders(!!refresh));

  onInit((init) => {
    panelWidth = Number(init.settings?.chatWidth) || 400;
    el.resizer.setAttribute('aria-valuenow', String(panelWidth));
    state.providersReady = loadProviders(false);
    setTarget(init.doc ?? null);
  });

  bindUi();
  bindEvents();
})();
