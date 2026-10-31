// 화면 언어 (src/app/i18n): 번역기·사전·알림 창/메뉴 번역·언어 고르기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  createI18n, loadDictionary, localizeDialogOptions, localizedDialog, localizeMenu, resolveLang, officeLocale, studioLang,
} from '../src/app/i18n/index.mjs';
import { extractAll } from '../scripts/i18n/extract.mjs';
import { skipKey } from '../scripts/i18n/rules.mjs';
import { systemPrompt } from '../src/app/ai.mjs';

const ROOT = join(import.meta.dirname, '..');
const { compile } = globalThis.OfficeKitI18n;
const HAN = /[가-힣]/;

test('사전: 화면에 나오는 한국어 글이 영어·일본어 사전에 모두 있다 (자리 표시도 같게)', () => {
  const keys = [...extractAll().keys()].filter((k) => !skipKey(k));
  for (const lang of ['en', 'ja']) {
    const dict = loadDictionary(lang);
    const missing = keys.filter((k) => !Object.hasOwn(dict, k));
    assert.deepEqual(missing.slice(0, 5), [], `${lang}: 빠진 글 ${missing.length}개 — node scripts/i18n/merge.mjs 로 다시 만드세요`);
    for (const [k, v] of Object.entries(dict)) {
      const ph = (s) => (s.match(/\{\d+\}/g) ?? []).sort().join();
      assert.equal(ph(v), ph(k), `${lang}: 자리 표시가 다름 ${k}`);
    }
  }
});

test('번역기: 그대로 맞추기, 앞뒤 빈칸·줄바꿈, 자리 틀 (자리 글도 다시 번역)', () => {
  const tr = compile({
    '문서 열기': 'Open document', '저장하지 못했습니다: {0}': "Couldn't save: {0}", '파일을 찾을 수 없습니다': 'File not found',
    '{0}쪽': '{0} pages', '오늘 {0}': 'Today {0}', '{0}{1} 버전으로 되돌렸어요{2}': 'Reverted to the {0}{1} version{2}',
  }, 'en');
  assert.equal(tr.t('문서 열기'), 'Open document');
  assert.equal(tr.t('  문서\n    열기 '), '  Open document ', 'HTML의 줄바꿈·들여쓰기');
  assert.equal(tr.t('저장하지 못했습니다: 파일을 찾을 수 없습니다'), "Couldn't save: File not found");
  assert.equal(tr.t('저장하지 못했습니다: 보고서.hwp 잠김'), "Couldn't save: 보고서.hwp 잠김", '모르는 자리 글은 그대로');
  assert.equal(tr.t('12쪽'), '12 pages');
  assert.equal(tr.t('Hello'), 'Hello', '한국어가 없으면 그대로');
  assert.equal(compile({ '문서 열기': 'x' }, 'ko').t('문서 열기'), '문서 열기', '한국어 화면은 바꾸지 않는다');
  // 한국어 고정 글이 없는 틀은 무엇이든 맞으므로 쓰지 않는다
  assert.equal(compile({ '{0} {1}': '{1} {0}' }, 'en').t('가 나'), '가 나');
});

test('main 번역: 알림 창 옵션·메뉴·언어 고르기', () => {
  const t = (s) => ({ '확인': 'OK', '취소': 'Cancel', '파일': 'File', '끝내기': 'Quit', 'PDF 문서': 'PDF document', '저장할까요?': 'Save?' }[s] ?? s);
  const o = localizeDialogOptions({ title: '저장할까요?', buttons: ['확인', '취소'], filters: [{ name: 'PDF 문서', extensions: ['pdf'] }], defaultPath: 'C:\\문서.pdf' }, t);
  assert.deepEqual(o, { title: 'Save?', buttons: ['OK', 'Cancel'], filters: [{ name: 'PDF document', extensions: ['pdf'] }], defaultPath: 'C:\\문서.pdf' });
  const calls = [];
  const fake = { showMessageBox: (...a) => { calls.push(a); return Promise.resolve({ response: 0 }); } };
  const win = { webContents: {} };
  localizedDialog(fake, () => t).showMessageBox(win, { message: '저장할까요?', buttons: ['확인'] });
  assert.equal(calls[0][0], win, '창은 그대로');
  assert.deepEqual(calls[0][1], { message: 'Save?', buttons: ['OK'] });
  const menu = localizeMenu([{ label: '파일', submenu: [{ label: '끝내기', role: 'quit' }, { type: 'separator' }] }], t);
  assert.deepEqual(menu, [{ label: 'File', submenu: [{ label: 'Quit', role: 'quit' }, { type: 'separator' }] }]);

  assert.equal(resolveLang('system', ['ja-JP', 'en-US']), 'ja');
  assert.equal(resolveLang('system', ['ko-KR']), 'ko');
  assert.equal(resolveLang('system', ['de-DE']), 'en', '지원하지 않는 언어는 영어');
  assert.equal(resolveLang('ko', ['en-US']), 'ko', '고른 언어가 먼저');
  assert.equal(resolveLang('bogus', []), 'ko');
  assert.deepEqual([studioLang('ja'), officeLocale('ja'), officeLocale('ko')], ['en', 'en-US', 'ko-KR'], '편집기에 일본어 메뉴가 없어 영어로');
  const en = createI18n('en');
  assert.equal(en.t('파일 열기…'), 'Open file…');
  assert.equal(createI18n('ko').t('파일 열기…'), '파일 열기…');
});

test('AI 지침: 고른 언어로 답하게 한다 (한국어는 예전 그대로)', () => {
  assert.match(systemPrompt({ lang: 'ko' }), /답변은 짧은 한국어로 한다/);
  assert.match(systemPrompt({ lang: 'en' }), /답변은 짧은 English\(영어\)로 한다/);
  assert.match(systemPrompt({ lang: 'ja' }), /日本語\(일본어\)/);
});

test('일본어 사전에 한국어가 남지 않는다 (지침처럼 번역하지 않는 글 빼고)', () => {
  const ja = loadDictionary('ja');
  const left = Object.entries(ja).filter(([k, v]) => HAN.test(v) && v !== k).map(([k]) => k);
  assert.deepEqual(left, []);
  // <b>가 섞인 문장 (data-i18n-rich)도 통째로 들어 있다
  const html = readFileSync(join(ROOT, 'src', 'app', 'renderer', 'index.html'), 'utf8');
  const rich = [...html.matchAll(/data-i18n-rich[^>]*>([\s\S]*?)<\/p>/g)].map((m) => m[1].replace(/\s+/g, ' ').trim());
  assert.ok(rich.length >= 4);
  for (const k of rich) assert.ok(Object.hasOwn(ja, k) && Object.hasOwn(loadDictionary('en'), k), k.slice(0, 30));
});
