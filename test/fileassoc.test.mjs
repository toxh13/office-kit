// 기본 프로그램 후보 등록 (src/setup/fileassoc.mjs · build/installer.nsh)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  ASSOC_TYPES, ASSOC_TYPE_NAMES, APP_DESCRIPTION, OPEN_VERB, ASSOC_GROUPS, PROGID_PREFIX, DEFAULT_APPS_URI, registrationEntries, summarizeStatus, assocStatus, registerFileTypes, unregisterForTest,
} from '../src/setup/fileassoc.mjs';
import { OPENABLE_EXTS } from '../src/core/index.mjs';

const EXE = 'C:\\Users\\김\\AppData\\Local\\Programs\\Office Kit\\Office Kit.exe';

test('등록 형식 = 열 수 있는 형식 = 설치 프로그램 목록', () => {
  assert.deepEqual(Object.keys(ASSOC_TYPES).sort(), [...OPENABLE_EXTS].sort());
  assert.deepEqual(ASSOC_GROUPS.flatMap((g) => g.exts).sort(), [...OPENABLE_EXTS].sort());
  const nsh = readFileSync(join(import.meta.dirname, '..', 'build', 'installer.nsh'), 'utf8');
  const block = /!macro OfficeKitForEachType MACRO([\s\S]*?)!macroend/.exec(nsh)[1];
  const rows = [...block.matchAll(/\$\{MACRO\} (\w+) "([^"]+)" "([^"]+)" "([^"]+)"/g)];
  ['ko', 'en', 'ja'].forEach((lang, i) => {
    const fromNsh = Object.fromEntries(rows.map((m) => [m[1], m[2 + i]]));
    assert.deepEqual(fromNsh, ASSOC_TYPE_NAMES[lang], `installer.nsh의 형식·이름(${lang})이 fileassoc.mjs와 같아야 한다`);
  });
  // 오른쪽 클릭 메뉴 이름·기본 앱 설명도 세 언어가 같아야 한다
  for (const table of [OPEN_VERB, APP_DESCRIPTION]) assert.ok(nsh.includes(`OfficeKitPick "${table.ko}" "${table.en}" "${table.ja}"`));
  // 기본 연결을 덮어쓰는 키(.ext의 기본값)는 쓰지 않는다
  assert.doesNotMatch(nsh, /WriteRegStr HKCU "\$\{OK_CLASSES\}\\\.\$\{EXT\}" ""/);
});

test('레지스트리 값: ProgID·OpenWithProgids·Capabilities, .확장자 기본값은 건드리지 않음', () => {
  const e = registrationEntries(EXE);
  const get = (key, name) => e.find((x) => x.key === key && x.name === name)?.value;
  assert.equal(get('Software\\Classes\\OfficeKit.docx\\shell\\open\\command', ''), `"${EXE}" "%1"`);
  assert.equal(get('Software\\Classes\\.docx\\OpenWithProgids', 'OfficeKit.docx'), '');
  assert.equal(get('Software\\OfficeKit\\Capabilities\\FileAssociations', '.hwpx'), 'OfficeKit.hwpx');
  assert.equal(get('Software\\RegisteredApplications', 'Office Kit'), 'Software\\OfficeKit\\Capabilities');
  assert.ok(!e.some((x) => /^Software\\Classes\\\.[a-z]+$/.test(x.key) && x.name === ''), '.확장자 키의 기본값(기본 프로그램)은 쓰지 않는다');
  assert.ok(!e.some((x) => /FileExts|UserChoice/i.test(x.key)), 'UserChoice는 쓰지 않는다');
  assert.match(DEFAULT_APPS_URI, /^ms-settings:defaultapps\?registeredAppUser=Office%20Kit$/);
});

test('상태 요약: UserChoiceLatest가 UserChoice보다 앞서고, 경로가 다르면 matches=false', () => {
  const types = Object.fromEntries(Object.keys(ASSOC_TYPES).map((x) => [x, { listed: `${PROGID_PREFIX}${x}`, command: `"${EXE}" "%1"`, user: '', latest: '' }]));
  types.docx.user = 'Word.Document.12';
  types.docx.latest = 'OfficeKit.docx';
  types.pptx.user = 'officekit.pptx';
  types.xlsx.user = 'Excel.Sheet.12';
  const st = summarizeStatus({ registered: 'Software\\OfficeKit\\Capabilities', types }, EXE);
  assert.equal(st.registered, true);
  assert.equal(st.matches, true);
  const by = Object.fromEntries(st.types.map((t) => [t.ext, t]));
  assert.equal(by.docx.isDefault, true);
  assert.equal(by.pptx.isDefault, true, 'ProgID는 대소문자를 가리지 않는다');
  assert.equal(by.xlsx.isDefault, false);
  assert.equal(by.hwp.progId, null);
  assert.equal(summarizeStatus({ registered: '', types }, 'D:\\다른\\Office Kit.exe').matches, false);
  assert.equal(summarizeStatus({ registered: '', types: {} }, EXE).registered, false);
});

test('실제 레지스트리: 시험용 키에 등록 → 다시 읽기 (한글 경로 포함)', { skip: process.platform !== 'win32' }, async () => {
  const root = `Software\\OfficeKitTest-${process.pid}`;
  try {
    let st = await assocStatus({ exe: EXE, root });
    assert.equal(st.registered, false);
    assert.equal(st.matches, false);
    const n = await registerFileTypes({ exe: EXE, root });
    assert.ok(n > 15 * 5);
    st = await assocStatus({ exe: EXE, root });
    assert.equal(st.registered, true);
    assert.equal(st.matches, true, '열기 명령이 한글이 든 경로 그대로 읽혀야 한다');
    assert.equal((await assocStatus({ exe: 'C:\\x\\Office Kit.exe', root })).matches, false);
  } finally {
    await unregisterForTest({ root });
  }
  await assert.rejects(unregisterForTest({ root: 'Software' }));
});

test('Office Kit가 기본 프로그램이면 다른 프로그램으로 열기 (프로그램에서 열기가 자기를 다시 열지 않게)', async () => {
  const { pickOpenCommand, parseOpenCommand } = await import('../src/setup/fileassoc.mjs');
  const r = String.raw;
  const ok = r`"C:\Users\a\AppData\Local\Programs\Office Kit\Office Kit.exe" "%1"`;
  // Windows가 쓰는 명령이 Office Kit가 아니면 그대로
  assert.equal(pickOpenCommand({ effective: r`"C:\Hnc\Hwp.exe" "%1"`, candidates: [] }), r`"C:\Hnc\Hwp.exe" "%1"`);
  // Office Kit면 다른 후보 중 이름난 프로그램 먼저 (Word > 한컴 > 그 밖)
  assert.equal(pickOpenCommand({ effective: ok, candidates: [
    { id: 'LibreOffice.Docx', command: r`"C:\LO\swriter.exe" -o "%1"` },
    { id: 'HWord.Document.docx.130', command: r`"C:\Hnc\Hword.exe" "%1"` },
    { id: 'Word.Document.12', command: r`"C:\Office\WINWORD.EXE" /n "%1"` },
  ] }), r`"C:\Office\WINWORD.EXE" /n "%1"`);
  assert.equal(pickOpenCommand({ effective: r`C:\Windows\system32\OpenWith.exe "%1"`, candidates: [{ id: 'X', command: ok }] }), null);
  // 명령 → 실행 파일과 인수 (%1 → 경로, 나머지 자리 지움, 환경 변수 풀기)
  assert.deepEqual(parseOpenCommand(r`"C:\Hnc\Hwp.exe" "%1"`, r`D:\문서\a b.hwp`), { file: r`C:\Hnc\Hwp.exe`, args: r`"D:\문서\a b.hwp"` });
  assert.deepEqual(parseOpenCommand(r`%PF%\Viewer.exe /open "%1" %*`, r`D:\x.pdf`, { PF: r`C:\Program Files` }), { file: r`C:\Program Files\Viewer.exe`, args: r`/open "D:\x.pdf"` },
    '따옴표 없는 경로에 빈칸이 있어도 .exe까지가 실행 파일');
  assert.deepEqual(parseOpenCommand(r`"%PF%\Viewer.exe" /open "%1" %*`, r`D:\x.pdf`, { PF: r`C:\Program Files` }), { file: r`C:\Program Files\Viewer.exe`, args: r`/open "D:\x.pdf"` });
  assert.deepEqual(parseOpenCommand(r`"C:\a\b.exe"`, r`D:\x.pdf`), { file: r`C:\a\b.exe`, args: r`"D:\x.pdf"` }, '%1이 없으면 끝에 붙인다');
});
