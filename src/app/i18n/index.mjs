// index.mjs — main 프로세스의 화면 번역: 사전 읽기, 언어 고르기, 알림 창(dialog)·메뉴 글 바꾸기.
//   화면(renderer)은 같은 translate.js로 DOM을 바꾸고, 여기서는 Electron이 직접 그리는 글(알림 창·메뉴·창 제목)을 바꾼다.
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import './translate.js';

const here = dirname(fileURLToPath(import.meta.url));
const I18N = globalThis.OfficeKitI18n;

export const LANGS = I18N.LANGS;
export const LANG_SETTINGS = ['system', ...LANGS];
export const pickLang = I18N.pickLang;

/** 언어 설정('system' | 'ko' | 'en' | 'ja') + Windows 언어 목록 → 실제 언어 */
export function resolveLang(setting, systemLocales = []) {
  if (LANGS.includes(setting)) return setting;
  const first = [systemLocales].flat().find((l) => typeof l === 'string' && l);
  return pickLang(first || 'ko');
}

const dictCache = new Map();
/** 언어 → 사전 { 한국어: 번역 } (한국어는 빈 사전) */
export function loadDictionary(lang) {
  if (lang === 'ko' || !LANGS.includes(lang)) return {};
  if (!dictCache.has(lang)) {
    let dict = {};
    try { dict = JSON.parse(readFileSync(join(here, `${lang}.json`), 'utf8')); } catch {}
    dictCache.set(lang, dict);
  }
  return dictCache.get(lang);
}

/** 언어 → { lang, dict, t } */
export function createI18n(lang) {
  const dict = loadDictionary(lang);
  const { t } = I18N.compile(dict, lang);
  return { lang, dict, t };
}

const str = (t, v) => (typeof v === 'string' ? t(v) : v);

/** 알림 창 옵션의 글을 바꾼 사본 */
export function localizeDialogOptions(o, t) {
  if (!o || typeof o !== 'object') return o;
  const out = { ...o };
  for (const k of ['title', 'message', 'detail', 'checkboxLabel', 'buttonLabel', 'nameFieldLabel']) if (k in out) out[k] = str(t, out[k]);
  if (Array.isArray(out.buttons)) out.buttons = out.buttons.map((b) => str(t, b));
  if (Array.isArray(out.filters)) out.filters = out.filters.map((f) => (f && typeof f === 'object' ? { ...f, name: str(t, f.name) } : f));
  return out;
}

/**
 * Electron dialog → 글을 번역해 보이는 dialog (같은 모양). getT: 지금 번역 함수 (언어가 바뀌어도 따라가게)
 * 인수는 (창?, 옵션) — 마지막 객체 인수가 옵션이다
 */
export function localizedDialog(dialog, getT) {
  const wrap = (name) => (...args) => {
    const t = getT();
    const i = args.length - 1;
    if (i >= 0 && args[i] && typeof args[i] === 'object' && !('webContents' in args[i])) args[i] = localizeDialogOptions(args[i], t);
    return dialog[name](...args);
  };
  return {
    showMessageBox: wrap('showMessageBox'), showMessageBoxSync: wrap('showMessageBoxSync'),
    showOpenDialog: wrap('showOpenDialog'), showOpenDialogSync: wrap('showOpenDialogSync'),
    showSaveDialog: wrap('showSaveDialog'), showSaveDialogSync: wrap('showSaveDialogSync'),
    showErrorBox: (title, content) => { const t = getT(); return dialog.showErrorBox(str(t, title), str(t, content)); },
    showCertificateTrustDialog: (...a) => dialog.showCertificateTrustDialog(...a),
  };
}

/** 메뉴 틀(template)의 글을 바꾼 사본 */
export function localizeMenu(template, t) {
  return template.map((item) => {
    if (!item || typeof item !== 'object') return item;
    const out = { ...item };
    for (const k of ['label', 'sublabel', 'toolTip']) if (k in out) out[k] = str(t, out[k]);
    if (Array.isArray(out.submenu)) out.submenu = localizeMenu(out.submenu, t);
    return out;
  });
}

/** 편집기 언어: 한글 편집기(rhwp-studio)는 ko·en, Office 편집기는 ko-KR·en-US만 있다 → 일본어는 영어로 */
export const studioLang = (lang) => (lang === 'ko' ? 'ko' : 'en');
export const officeLocale = (lang) => (lang === 'ko' ? 'ko-KR' : 'en-US');

/** AI에게 답할 언어를 알려 줄 이름 */
export const LANG_NAMES = { ko: '한국어', en: 'English', ja: '日本語' };
