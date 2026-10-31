<!--
office-editor-legal-en.md — Office 편집 창 '정보 · 라이선스' 대화상자 글 (영어. 화면 언어가 영어·일본어일 때). ONLYOFFICE 추가 조항 3의 화면 고지.
office-editor-legal-ko.md와 같은 내용·같은 문법·같은 고정 값으로 쓴다 (한쪽을 고치면 다른 쪽도).
시험(test/editor-integration.test.mjs)이 이 주석을 떼고 src/app/office-editor/legal-en.json을 만든다 (OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs).
-->
# About · Licenses

This editor window (Document editor · Spreadsheet editor · Slide editor) is based on **ONLYOFFICE Docs**.

## Original developer

- The original developer of the editor is **Ascensio System SIA** (ONLYOFFICE). [www.onlyoffice.com](https://www.onlyoffice.com)
- Version in use: ONLYOFFICE Docs Community Edition 9.4.0 (build 9.4.0.129), source tag `v9.4.0.129`
- The bridge code that runs the editor files in the browser and the app page are BRF Tech's filex-office-editor v0.3.1.

## This may be a modified version

- The editor in Office Kit is not the original as is; **it may be a modified version**.
- Changes: what filex-office-editor changed when packaging the editor files; files Office Kit does not use were removed (other-language files, mobile screens, the font folder of the editor package); while editing, the font list is replaced with the fonts on this PC.
- Dated lists: [Modification notice (NOTICE-OfficeKit.md)](local:NOTICE-OfficeKit.md) · [filex-office-editor notice (NOTICE)](local:NOTICE)
- The logos inside the editor and the editor's own "About" screen are left unchanged.

## Licenses

- The editor, the converter (x2t), and filex-office-editor are licensed under the **GNU Affero General Public License version 3 (AGPL-3.0)**. The ONLYOFFICE parts carry additional terms 1–5 (section 7).
- This is free software. You can redistribute and modify it under this license, and it comes with **no warranty**.
- [Full AGPL v3 text and ONLYOFFICE additional terms (LICENSE.txt)](local:editor/LICENSE.txt)
- [filex-office-editor license (LICENSE)](local:LICENSE)
- [Third-party components in the editor (3rd-Party.txt)](local:editor/3rd-Party.txt) — Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO, and more (MIT · BSD · Apache-2.0)
- [Third-party components in the converter and font engine (WebAssembly) (THIRD-PARTY-NATIVE.txt)](local:THIRD-PARTY-NATIVE.txt) — Boost, ICU, FreeType, HarfBuzz, libpng, libjpeg, libtiff, OpenJPEG, libxml2, Xpdf, DjVuLibre, the emscripten runtime, URW++ fonts, and more
- Non-code material such as images, icons, and help text is **CC BY-SA 4.0**. Author: Ascensio System SIA (ONLYOFFICE). [CC BY-SA 4.0 legal code](https://creativecommons.org/licenses/by-sa/4.0/legalcode)
- Office Kit's own code (the window, saving, the font list, and so on) is MIT, except that the filex patches are AGPL-3.0-or-later and the font list generator is AGPL-3.0-only (+ ONLYOFFICE additional terms).

## Source code

- Editor source package for this version (Office Kit {{appVersion}}): [office-editor-source-{{appVersion}}.zip](https://github.com/toxh13/office-kit/releases/download/v{{appVersion}}/office-editor-source-{{appVersion}}.zip)
- Office Kit source: [github.com/toxh13/office-kit (v{{appVersion}})](https://github.com/toxh13/office-kit/tree/v{{appVersion}})
- filex-office-editor v0.3.1: [github.com/BRF-Tech/filex-office-editor](https://github.com/BRF-Tech/filex-office-editor/tree/v0.3.1)
- ONLYOFFICE original source (tag `v9.4.0.129`): [sdkjs](https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129) · [web-apps](https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129) · [core](https://github.com/ONLYOFFICE/core/tree/v9.4.0.129) · [build_tools](https://github.com/ONLYOFFICE/build_tools/tree/v9.4.0.129)
- The x2t converter is built from ONLYOFFICE core `v9.4.0.129` plus CryptPad's WebAssembly patches and BRF Tech's patches. How to build it is in `UPSTREAM.md` in the source package.

## Trademarks

- ONLYOFFICE® and ONLYOFFICE Docs™ are trademarks of Ascensio System SIA. Office Kit is not affiliated with Ascensio System SIA and is not endorsed or sponsored by it. [Trademark policy](https://www.onlyoffice.com/trademark-policy)
- Microsoft Word, Excel, and PowerPoint are trademarks of Microsoft. Office Kit uses these names only to refer to file formats.

## Fonts

- Fonts used for editing are not distributed with the app. The editor only reads the fonts installed on this PC.
- However, the x2t converter contains the standard 14 PDF fonts that ONLYOFFICE core includes (URW++ Nimbus Sans · Nimbus Roman · Nimbus Mono, Standard Symbols L, Dingbats). These are free fonts under AGPL-3.0 with an exception that allows embedding them in PDF and PostScript documents. [Notice (THIRD-PARTY-NATIVE.txt)](local:THIRD-PARTY-NATIVE.txt)
