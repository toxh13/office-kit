// 탭 목록 모델(순수 로직). Electron·fs 없이 탭 순서·활성·단축키 표만 다룬다.
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import { resolve } from 'node:path';

// 같은 문서인지 가르는 열쇠. 윈도우는 대소문자를 구분하지 않는다.
export function docKey(p) {
  const r = resolve(p);
  return process.platform === 'win32' ? r.toLowerCase() : r;
}

const HOME = () => ({ id: 'home', type: 'home', title: '홈', path: null, convertedFrom: null, dirty: false, loaded: true, zone3: false });

export function createTabList({ newId = randomUUID } = {}) {
  const em = new EventEmitter();
  const list = [HOME()];
  let activeId = 'home';
  const changed = () => em.emit('change');
  const idx = (id) => list.findIndex((t) => t.id === id);
  const need = (id) => {
    const i = idx(id);
    if (i < 0) throw Object.assign(new Error('없는 탭입니다'), { code: 'NO_TAB' });
    return i;
  };
  const make = (tab) => ({ title: '', path: null, convertedFrom: null, dirty: false, loaded: true, zone3: false, ...tab, id: tab.id ?? newId() });

  return {
    get tabs() { return list.slice(); },
    get active() { return list[idx(activeId)]; },
    get(id) { return list.find((t) => t.id === id) ?? null; },
    add(tab, { at, activate = true } = {}) {
      const t = make(tab);
      const pos = at ?? idx(activeId) + 1;
      list.splice(Math.max(0, Math.min(pos, list.length)), 0, t);
      if (activate) activeId = t.id;
      changed();
      return t;
    },
    remove(id) {
      const i = need(id);
      if (list[i].type === 'home') throw Object.assign(new Error('홈 탭은 닫을 수 없습니다'), { code: 'HOME_FIXED' });
      const [removed] = list.splice(i, 1);
      if (activeId === id) activeId = (list[i] ?? list[i - 1]).id;
      changed();
      return { removed, active: list[idx(activeId)] };
    },
    activate(id) {
      need(id);
      activeId = id;
      changed();
      return list[idx(id)];
    },
    update(id, patch) {
      const i = need(id);
      const { id: _i, type: _t, ...rest } = patch;
      list[i] = { ...list[i], ...rest };
      changed();
      return list[i];
    },
    replace(id, tab) {
      const i = need(id);
      const t = make(tab);
      list[i] = t;
      if (activeId === id) activeId = t.id;
      changed();
      return t;
    },
    find(path) {
      const k = docKey(path);
      return list.find((t) => (t.path && docKey(t.path) === k) || (t.convertedFrom && docKey(t.convertedFrom) === k)) ?? null;
    },
    step(dir) {
      const n = list.length;
      activeId = list[(idx(activeId) + dir + n) % n].id;
      changed();
      return list[idx(activeId)];
    },
    at(n) {
      if (n === 9) return list[list.length - 1];
      return n >= 1 && n <= 8 ? list[n - 1] ?? null : null;
    },
    snapshot() {
      return {
        active: idx(activeId),
        tabs: list.filter((t) => t.type !== 'home').map((t) => {
          const o = { type: t.type, path: t.path };
          if (t.convertedFrom) o.convertedFrom = t.convertedFrom;
          return o;
        }),
      };
    },
    on: (ev, fn) => em.on(ev, fn),
  };
}

// 탭 단축키 표. 키보드 배열과 무관하게 input.code로 판정한다.
export function tabCommandOf(input) {
  if (!input || input.type !== 'keyDown') return null;
  if (!(input.control || input.meta) || input.alt) return null;
  const code = input.code;
  if (code === 'Tab') return { cmd: input.shift ? 'prev' : 'next' };
  if (input.shift) return null;
  if (code === 'KeyT') return { cmd: 'home' };
  if (code === 'KeyW') return { cmd: 'close' };
  if (code === 'KeyN') return { cmd: 'new-window' };
  const m = /^Digit([1-9])$/.exec(code ?? '');
  return m ? { cmd: 'index', n: Number(m[1]) } : null;
}
