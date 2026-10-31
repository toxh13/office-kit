// translate.js — Office Kit 화면 번역기 (한국어 원문 → 영어·일본어). 의존성 없음.
//   사전: { "한국어 원문": "번역" }. 원문에 {0} {1} … 자리가 있으면 틀(패턴)로 맞춘다 — 자리에 든 글도 다시 번역한다.
//   화면(DOM)은 바뀌는 대로 번역하고 (MutationObserver), data-no-i18n 아래(문서 내용·AI 답·파일 이름)는 건드리지 않는다.
//   data-i18n-rich 요소는 <b>가 섞인 문장을 통째로 (키: 안쪽 HTML, 공백 하나로) — 조각으로 옮기면 어순이 깨진다.
//   일반 <script>로도, ES 모듈 import로도 쓴다 (globalThis.OfficeKitI18n).
(function (g) {
  const HAN = /[가-힣]/;
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  /** 사전 → { t(글), lang } */
  function compile(dict, lang = 'ko') {
    const exact = new Map();
    const pats = [];
    for (const [k, v] of Object.entries(dict || {})) {
      if (typeof v !== 'string' || !v) continue;
      if (/\{\d+\}/.test(k)) {
        const parts = k.split(/\{(\d+)\}/);
        const lit = parts.filter((_, i) => i % 2 === 0).join('');
        if (!HAN.test(lit)) continue; // 한국어 고정 글이 없는 틀은 무엇이든 맞으므로 쓰지 않는다
        let re = '^'; const order = [];
        parts.forEach((p, i) => { if (i % 2 === 0) re += esc(p); else { re += '([\\s\\S]+?)'; order.push(Number(p)); } });
        try { pats.push({ re: new RegExp(`${re}$`), order, out: v, lit: lit.length }); } catch {}
      } else exact.set(k, v);
    }
    pats.sort((a, b) => b.lit - a.lit); // 고정 글이 긴(더 구체적인) 틀 먼저
    const cache = new Map();

    function t(s, depth = 0) {
      if (lang === 'ko' || typeof s !== 'string' || !s || depth > 5 || !HAN.test(s)) return s;
      if (exact.has(s)) return exact.get(s);
      if (!depth && cache.has(s)) return cache.get(s);
      const lead = /^\s*/.exec(s)[0]; const tail = /\s*$/.exec(s)[0];
      const core = s.slice(lead.length, s.length - tail.length);
      let out = s;
      const flat = core.replace(/\s+/g, ' '); // HTML 글은 줄바꿈·들여쓰기가 섞여 있다
      if (exact.has(core)) out = lead + exact.get(core) + tail;
      else if (exact.has(flat)) out = lead + exact.get(flat) + tail;
      else {
        for (const p of pats) {
          const m = p.re.exec(core);
          if (!m) continue;
          let r = p.out;
          p.order.forEach((n, i) => { r = r.split(`{${n}}`).join(t(m[i + 1], depth + 1)); });
          out = lead + r + tail;
          break;
        }
      }
      if (!depth) { if (cache.size > 5000) cache.clear(); cache.set(s, out); }
      return out;
    }
    return { t, lang };
  }

  const ATTRS = ['title', 'placeholder', 'aria-label', 'alt'];
  const RICH = '[data-i18n-rich]';
  /** 문장 HTML 키 (<b>만 남기고 공백 정리) — extract.mjs와 같은 규칙 */
  const richKey = (html) => html.replace(/<(\/?)b\b[^>]*>/gi, '<$1b>').replace(/\s+/g, ' ').trim();
  const SKIP = '[data-no-i18n], script, style, textarea, [contenteditable="true"]';

  /** 화면 번역: root 아래 글·속성을 지금 바꾸고, 앞으로 바뀌는 것도 바꾼다 → 멈추는 함수. opts.skip: 더 뺄 곳(문서 내용 등) 선택자 */
  function translateDom(root, tr, opts = {}) {
    if (!root || !tr || tr.lang === 'ko') return () => {};
    const doc = root.ownerDocument || root;
    const skipSel = opts.skip ? `${SKIP}, ${opts.skip}` : SKIP;
    const skipped = (el) => !!(el && el.closest && el.closest(skipSel));
    // 속성(안내 글·placeholder)은 입력 칸·편집 영역에서도 바꾼다 — 내용만 그대로
    const attrSel = opts.skip ? `[data-no-i18n], ${opts.skip}` : '[data-no-i18n]';
    const attrSkipped = (el) => !!(el && el.closest && el.closest(attrSel));
    const doText = (node) => {
      const v = node.nodeValue;
      if (!v || !HAN.test(v) || skipped(node.parentElement)) return;
      const out = tr.t(v);
      if (out !== v) node.nodeValue = out;
    };
    const doAttrs = (el) => {
      if (attrSkipped(el)) return;
      for (const a of ATTRS) {
        const v = el.getAttribute && el.getAttribute(a);
        if (v && HAN.test(v)) { const out = tr.t(v); if (out !== v) el.setAttribute(a, out); }
      }
    };
    // 번역문에서 <b>…</b>만 요소로 만들고 나머지는 글로 (사전 글을 HTML로 넣지 않는다)
    const doRich = (el) => {
      if (el.dataset.i18nDone) return true;
      const key = richKey(el.innerHTML);
      const out = tr.t(key);
      if (!out || out === key) return false;
      el.textContent = '';
      out.split(/(<b>[\s\S]*?<\/b>)/).forEach((part) => {
        const m = /^<b>([\s\S]*)<\/b>$/.exec(part);
        if (m) { const b = doc.createElement('b'); b.textContent = m[1]; el.appendChild(b); } else if (part) el.appendChild(doc.createTextNode(part));
      });
      el.dataset.i18nDone = '1';
      return true;
    };
    const doTree = (n) => {
      if (n.nodeType === 3) { doText(n); return; }
      if (n.nodeType !== 1 && n.nodeType !== 9 && n.nodeType !== 11) return;
      if (n.nodeType === 1) { doAttrs(n); if (skipped(n)) return; }
      if (n.querySelectorAll) for (const el of n.querySelectorAll(RICH)) if (!skipped(el)) doRich(el);
      if (n.nodeType === 1 && n.matches(RICH)) doRich(n);
      const w = doc.createTreeWalker(n, 1 | 4, { acceptNode: (x) => { if (x.nodeType === 1 && x.matches && x.matches(skipSel)) { doAttrs(x); return 2; } return 1; } });
      let x = w.nextNode();
      while (x) { if (x.nodeType === 3) doText(x); else doAttrs(x); x = w.nextNode(); }
    };
    // 창 제목: 바뀔 때만 쓴다 (같은 값을 다시 쓰면 제목 감시가 또 불려 끝없이 돈다)
    const doTitle = () => { const v = doc.title; if (!v || !HAN.test(v)) return; const out = tr.t(v); if (out !== v) doc.title = out; };
    doTree(root);
    doTitle();
    const mo = new MutationObserver((list) => {
      for (const m of list) {
        if (m.type === 'characterData') doText(m.target);
        else if (m.type === 'attributes') doAttrs(m.target);
        else for (const n of m.addedNodes) doTree(n);
      }
    });
    mo.observe(root, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ATTRS });
    // 창 제목 (document.title은 head의 글이라 따로)
    const titleEl = doc.querySelector && doc.querySelector('title');
    let tmo = null;
    if (titleEl) {
      tmo = new MutationObserver(doTitle);
      tmo.observe(titleEl, { childList: true, characterData: true, subtree: true });
    }
    return () => { mo.disconnect(); if (tmo) tmo.disconnect(); };
  }

  /** Windows·브라우저 언어 → 'ko' | 'en' | 'ja' */
  function pickLang(locale) {
    const l = String(locale || '').toLowerCase();
    if (l.startsWith('ko')) return 'ko';
    if (l.startsWith('ja')) return 'ja';
    return 'en';
  }

  g.OfficeKitI18n = { compile, translateDom, pickLang, LANGS: ['ko', 'en', 'ja'] };
})(typeof globalThis !== 'undefined' ? globalThis : window);
