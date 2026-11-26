// 탭 세션 저장 규칙(순수 로직). 설정에는 지금 세션(tabSession)과 지난 세션(lastTabSession)을 둔다.
// 모양: { v:1, at, explorer, finalized, windows:[{ bounds, maximized, active, tabs:[{ type, path, … }] }] }
// active는 홈을 0번으로 센 자리이고, tabs에는 홈이 없다.
import { existsSync } from 'node:fs';

export const SESSION_DEBOUNCE_MS = 1000;

export function createSessionKeeper({ settings, explorer = false, delayMs = SESSION_DEBOUNCE_MS, setTimer = setTimeout, clearTimer = clearTimeout, now = Date.now }) {
  const wins = new Map(); // 창 번호 → 창 값 (만든 차례 유지)
  let timer = null;
  let closing = false;

  const build = (finalized) => ({
    v: 1, at: now(), explorer, finalized,
    windows: [...wins.values()].map((w) => ({
      bounds: w.bounds ?? null,
      maximized: !!w.maximized,
      active: w.snapshot?.active ?? 0,
      tabs: w.snapshot?.tabs ?? [],
    })),
  });
  const cancel = () => { if (timer != null) { clearTimer(timer); timer = null; } };
  const schedule = () => {
    cancel();
    timer = setTimer(() => {
      timer = null;
      // 나중에 쓰는 것이라 던지면 잡을 곳이 없다 (디스크 문제 등) — 다음 변경 때 다시 쓴다
      try { settings.set({ tabSession: build(false) }); } catch (err) { console.error('[session] 세션을 적지 못했습니다:', err?.message ?? err); }
    }, delayMs);
  };

  return {
    update(winId, win) {
      if (closing) return;
      wins.set(winId, win);
      schedule();
    },
    remove(winId) {
      if (closing) return;
      if (wins.delete(winId)) schedule();
    },
    finalize() {
      cancel();
      closing = true;
      const s = build(true);
      // 한 번의 set으로 쓴다. 탐색기로 연 세션은 지난 세션을 덮지 않는다.
      settings.set(explorer ? { tabSession: s } : { tabSession: s, lastTabSession: s });
    },
    resume() {
      if (!closing) return;
      closing = false;
      schedule();
    },
  };
}

/** 시작할 때 한 번: 확정 못 한 평소 세션(강제 종료)은 지난 세션으로 올린다. */
export function promoteCrashed(settings) {
  const s = settings.get('tabSession');
  if (!s || s.finalized !== false || s.explorer !== false) return false;
  settings.set({ lastTabSession: s });
  return true;
}

/**
 * 저장된 세션 → 복원 계획. 없는 파일 탭은 빼고 missing에 모은다.
 * 바꿔 연 탭(convertedFrom)의 path는 시작 때 지우는 임시 사본이므로 원본으로 판정하고, path도 원본으로 바꾼다 (열 때 다시 바꾼다)
 */
export function restorePlan(session, { exists = existsSync } = {}) {
  const empty = () => ({ windows: [], missing: [] });
  if (!session || session.v !== 1 || !Array.isArray(session.windows)) return empty();
  const missing = [];
  const windows = [];
  for (const w of session.windows) {
    if (!w || typeof w !== 'object') continue;
    const tabs = Array.isArray(w.tabs) ? w.tabs : [];
    const keep = tabs.map((t) => {
      const src = t?.convertedFrom || t?.path;
      const ok = !src || exists(src);
      if (!ok) missing.push(src);
      return ok;
    });
    const kept = tabs.filter((_, i) => keep[i]).map((t) => (t?.convertedFrom ? { ...t, path: t.convertedFrom } : t));
    const a = Number.isInteger(w.active) && w.active >= 0 && w.active <= tabs.length ? w.active : 0;
    let active = 0;
    if (a > 0) {
      const before = keep.slice(0, a - 1).filter(Boolean).length; // 활성 자리 앞에 남은 탭 수
      if (keep[a - 1]) active = before + 1;
      else if (keep.slice(a).some(Boolean)) active = before + 1; // 오른쪽 이웃이 그 자리로 온다
      else active = before; // 왼쪽 이웃(없으면 0 = 홈)
    }
    windows.push({ bounds: w.bounds ?? null, maximized: !!w.maximized, active, tabs: kept });
  }
  return { windows, missing };
}
