// md.js — AI 답변용 안전한 마크다운 일부만 그린다.
// 모든 글자를 먼저 이스케이프한 뒤 굵게·인라인 코드·코드 블록·목록·표·줄바꿈만 태그로 바꾼다.
'use strict';
(() => {
  const esc = (s) => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  // 링크 (클릭은 app.js가 기본 브라우저로 넘긴다). 주소는 이스케이프 전 원문에서 찾는다
  // 주소 글자는 ASCII만 (따옴표·꺾쇠·대괄호 제외) → "https://a.com에서" 의 한글이 붙지 않는다
  const URL_CHARS = "[A-Za-z0-9\\-._~:/?#@!$&()*+,;=%]";
  const URL_NP = "[A-Za-z0-9\\-._~:/?#@!$&*+,;=%]"; // 괄호 빼고
  const LINK_RE = new RegExp(
    `\\[([^\\[\\]\\n]{1,300})\\]\\((https?:\\/\\/(?:${URL_NP}|\\(${URL_NP}*\\))+)\\)` // [글](주소) — 주소 속 괄호는 한 겹까지
    + `|<(https?:\\/\\/${URL_CHARS}+)>` //                             <주소>
    + `|\\bhttps?:\\/\\/${URL_CHARS}+`, //                             그냥 주소
    'g');
  const link = (url, text) => `<a class="ext" href="#" data-href="${esc(url)}">${esc(text)}</a>`;

  /** 그냥 주소 끝의 문장 부호는 링크에서 뺀다. 닫는 괄호는 짝이 없을 때만 */
  function splitTrail(url) {
    let end = url.length;
    for (;;) {
      const c = url[end - 1];
      if (/[.,;:!?*]/.test(c)) { end--; continue; }
      if (c === ')') {
        const s = url.slice(0, end);
        if ((s.match(/\)/g)?.length ?? 0) > (s.match(/\(/g)?.length ?? 0)) { end--; continue; }
      }
      break;
    }
    return [url.slice(0, end), url.slice(end)];
  }

  // 링크 자리는 이스케이프·굵게 처리 동안 사용자 영역 글자로 비워 두었다가 마지막에 끼운다
  const HOLD_A = '';
  const HOLD_B = '';

  function inline(raw) {
    return String(raw).split(/(`[^`\n]+`)/g).map((part) => {
      if (/^`[^`\n]+`$/.test(part)) return `<code>${esc(part.slice(1, -1))}</code>`;
      const links = [];
      const hold = (html) => `${HOLD_A}${links.push(html) - 1}${HOLD_B}`;
      const src = part.replace(/[]/g, '').replace(LINK_RE, (m, label, mdUrl, angleUrl) => {
        if (mdUrl) return hold(link(mdUrl, label));
        if (angleUrl) return hold(link(angleUrl, angleUrl));
        const [u, trail] = splitTrail(m);
        return hold(link(u, u)) + trail;
      });
      let t = esc(src);
      t = t.replace(/\*\*(?=\S)([^*\n]*?\S)\*\*/g, '<strong>$1</strong>')
        .replace(/__(?=\S)([^_\n]*?\S)__/g, '<strong>$1</strong>');
      return t.replace(/(\d+)/g, (_, n) => links[Number(n)] ?? '');
    }).join('');
  }

  const LIST = /^(\s*)([-*•]|\d{1,3}[.)])\s+(.*)$/;
  const TABLE = /^\s*\|.*\|\s*$/;
  const isSep = (cells) => cells.length && cells.every((c) => /^:?-{2,}:?$/.test(c.trim()));
  const cellsOf = (line) => line.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());

  function render(src) {
    const lines = String(src ?? '').replace(/\r\n?/g, '\n').split('\n');
    const out = [];
    let para = [];
    const flush = () => {
      if (para.length) out.push(`<p>${para.map(inline).join('<br>')}</p>`);
      para = [];
    };

    for (let i = 0; i < lines.length;) {
      const line = lines[i];

      // ```코드 블록``` (스트리밍 중 닫히지 않았으면 끝까지 코드로)
      if (/^\s*```/.test(line)) {
        flush();
        const buf = [];
        i++;
        while (i < lines.length && !/^\s*```\s*$/.test(lines[i])) buf.push(lines[i++]);
        i++;
        out.push(`<pre class="code"><code>${esc(buf.join('\n'))}</code></pre>`);
        continue;
      }
      if (!line.trim()) { flush(); i++; continue; }

      const hm = line.match(/^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/);
      if (hm) { flush(); out.push(`<p class="h h${Math.min(hm[1].length, 3)}">${inline(hm[2])}</p>`); i++; continue; }

      if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) { flush(); out.push('<hr>'); i++; continue; }

      // 표
      if (TABLE.test(line)) {
        flush();
        const rows = [];
        while (i < lines.length && TABLE.test(lines[i])) rows.push(cellsOf(lines[i++]));
        const head = rows.length > 1 && isSep(rows[1]) ? rows[0] : null;
        const body = rows.filter((r, n) => !isSep(r) && !(head && n === 0));
        const cols = Math.max(...rows.map((r) => r.length));
        const tr = (r, tag) => `<tr>${Array.from({ length: cols }, (_, c) => `<${tag}>${inline(r[c] ?? '')}</${tag}>`).join('')}</tr>`;
        out.push(`<table>${head ? `<thead>${tr(head, 'th')}</thead>` : ''}<tbody>${body.map((r) => tr(r, 'td')).join('')}</tbody></table>`);
        continue;
      }

      // 목록 (- * • 또는 1. 1))
      const lm = line.match(LIST);
      if (lm) {
        flush();
        const ordered = /\d/.test(lm[2]);
        const start = ordered ? parseInt(lm[2], 10) : 1;
        const items = [];
        while (i < lines.length) {
          const m = lines[i].match(LIST);
          if (m && /\d/.test(m[2]) === ordered) { items.push({ lines: [m[3]], sub: m[1].length >= 2 }); i++; continue; }
          // 들여 쓴 이어지는 줄
          if (!m && items.length && /^\s{2,}\S/.test(lines[i])) { items[items.length - 1].lines.push(lines[i].trim()); i++; continue; }
          // 다른 종류의 들여 쓴 목록은 같은 목록에 붙인다
          if (m && m[1].length >= 2 && items.length) { items.push({ lines: [m[3]], sub: true }); i++; continue; }
          break;
        }
        const tag = ordered ? 'ol' : 'ul';
        const startAttr = ordered && start !== 1 ? ` start="${start}"` : '';
        out.push(`<${tag}${startAttr}>${items.map((it) => `<li${it.sub ? ' class="sub"' : ''}>${it.lines.map(inline).join('<br>')}</li>`).join('')}</${tag}>`);
        continue;
      }

      // 인용
      if (/^\s*>\s?/.test(line)) {
        flush();
        const buf = [];
        while (i < lines.length && /^\s*>\s?/.test(lines[i])) buf.push(lines[i++].replace(/^\s*>\s?/, ''));
        out.push(`<blockquote>${buf.map(inline).join('<br>')}</blockquote>`);
        continue;
      }

      para.push(line);
      i++;
    }
    flush();
    return out.join('');
  }

  window.OKMarkdown = Object.freeze({ render, escape: esc });
})();
