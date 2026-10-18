<!--
office-editor-legal-ko.md — Office 편집 창 '정보 · 라이선스' 대화상자 글 (한국어). ONLYOFFICE 추가 조항 3의 화면 고지.

읽는 쪽
  - 이 문서는 실행 중에 읽지 않는다. 시험(test/editor-integration.test.mjs)이 맨 위의 이 주석 덩어리를 떼고 블록으로 바꿔
    src/app/office-editor/legal-ko.json을 만든다 (OFFICEKIT_UPDATE_LEGAL=1 node --test test/editor-integration.test.mjs).
    office 편집 창 host(host.js)가 그 JSON을 같은 폴더에서 읽고, {{appVersion}}을 앱 버전(예 0.1.3)으로 바꿔 그린다.
  - host는 아래 문법만 그린다. HTML은 쓰지 않는다 (글은 모두 글자로 이스케이프).
문법 (이것만)
  - '# ' 제목 1개, '## ' 소제목, 빈 줄로 나뉜 문단, '- ' 목록(한 단계), **굵게**, `코드`, [글](대상)
  - 대상이 https:// → 앱의 바깥 링크 흐름(묻고 기본 브라우저). window.open 쓰지 않음.
  - 대상이 local:<경로> → vendor/office-editor/<경로> 파일 (office 출처의 /office/<ver>/<경로>)을 대화상자 안 보기 틀에 글로 보여 준다.
    쓰는 경로: NOTICE-OfficeKit.md, NOTICE, LICENSE, editor/LICENSE.txt, editor/3rd-Party.txt, THIRD-PARTY-NATIVE.txt (host.js NOTICE_FILES에 있는 것만)
자리표시: {{appVersion}} 하나뿐. 나머지 고정 값(9.4.0.129, v0.3.1 등)은 글에 그대로 적는다 —
  올릴 때 scripts/make-office-source.mjs의 PINS와 함께 바꾼다 (그 스크립트의 checkLegalDoc이 맞는지 확인한다).
'ONLYOFFICE'는 사실을 말할 때만 쓴다 (기반·원 개발사·상표). 창·기능 이름으로 쓰지 않는다.
-->
# 정보 · 라이선스

이 편집 창(문서 편집 · 표 편집 · 슬라이드 편집)은 **ONLYOFFICE Docs**를 바탕으로 합니다.

## 원 개발사

- 편집기의 원 개발사는 **Ascensio System SIA**(ONLYOFFICE)입니다. [www.onlyoffice.com](https://www.onlyoffice.com)
- 쓰는 판: ONLYOFFICE Docs Community Edition 9.4.0 (빌드 9.4.0.129), 소스 태그 `v9.4.0.129`
- 편집기 파일을 브라우저 안에서 돌리는 다리 코드와 앱 페이지는 BRF Tech의 filex-office-editor v0.3.1입니다.

## 수정된 판일 수 있습니다

- Office Kit에 든 편집기는 원본 그대로가 아니라 **수정된 판일 수 있습니다**.
- 바뀐 것: filex-office-editor가 편집기 파일을 묶으며 바꾼 것, Office Kit이 쓰지 않는 파일(다른 나라 말 파일, 휴대폰용 화면, 편집기 묶음의 글꼴 폴더)을 덜어 낸 것, 편집 중에 글꼴 목록을 이 PC의 글꼴로 바꿔 내주는 것.
- 날짜가 붙은 목록: [수정 고지 (NOTICE-OfficeKit.md)](local:NOTICE-OfficeKit.md) · [filex-office-editor 고지 (NOTICE)](local:NOTICE)
- 편집기 안의 로고와 편집기 자체의 '정보' 화면은 바꾸지 않고 그대로 둡니다.

## 라이선스

- 편집기, 변환기(x2t), filex-office-editor는 **GNU Affero General Public License 3판(AGPL-3.0)** 을 따릅니다. ONLYOFFICE 부분에는 추가 조항(7조) 1~5가 붙습니다.
- 자유 소프트웨어입니다. 이 라이선스에 따라 다시 배포하거나 고칠 수 있고, **아무런 보증이 없습니다**.
- [AGPL 3판 전문과 ONLYOFFICE 추가 조항 (LICENSE.txt)](local:editor/LICENSE.txt)
- [filex-office-editor 라이선스 (LICENSE)](local:LICENSE)
- [편집기 안 제3자 구성 요소 (3rd-Party.txt)](local:editor/3rd-Party.txt) — Bootstrap, jQuery, RequireJS, Underscore, Backbone, Socket.IO 등 (MIT · BSD · Apache-2.0)
- [변환기·글꼴 엔진(WebAssembly) 안 제3자 구성 요소 (THIRD-PARTY-NATIVE.txt)](local:THIRD-PARTY-NATIVE.txt) — Boost, ICU, FreeType, HarfBuzz, libpng, libjpeg, libtiff, OpenJPEG, libxml2, Xpdf, DjVuLibre, emscripten 런타임, URW++ 글꼴 등
- 그림 · 아이콘 · 설명 글처럼 코드가 아닌 것은 **CC BY-SA 4.0**입니다. 저작자: Ascensio System SIA (ONLYOFFICE). [CC BY-SA 4.0 전문](https://creativecommons.org/licenses/by-sa/4.0/legalcode)
- 창, 저장, 글꼴 목록 등 Office Kit 자신의 코드는 MIT입니다. 단 filex 패치는 AGPL-3.0-or-later, 글꼴 목록 생성기는 AGPL-3.0-only(+ ONLYOFFICE 추가 조항)입니다.

## 소스 코드

- 이 판(Office Kit {{appVersion}})의 편집기 소스 꾸러미: [office-editor-source-{{appVersion}}.zip](https://github.com/toxh13/office-kit/releases/download/v{{appVersion}}/office-editor-source-{{appVersion}}.zip)
- Office Kit 소스: [github.com/toxh13/office-kit (v{{appVersion}})](https://github.com/toxh13/office-kit/tree/v{{appVersion}})
- filex-office-editor v0.3.1: [github.com/BRF-Tech/filex-office-editor](https://github.com/BRF-Tech/filex-office-editor/tree/v0.3.1)
- ONLYOFFICE 원본 소스 (태그 `v9.4.0.129`): [sdkjs](https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129) · [web-apps](https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129) · [core](https://github.com/ONLYOFFICE/core/tree/v9.4.0.129) · [build_tools](https://github.com/ONLYOFFICE/build_tools/tree/v9.4.0.129)
- 변환기 x2t는 ONLYOFFICE core `v9.4.0.129`에 CryptPad의 WebAssembly 패치와 BRF Tech의 패치를 더해 만든 것입니다. 만드는 방법은 소스 꾸러미의 `UPSTREAM.md`에 있습니다.

## 상표

- ONLYOFFICE®와 ONLYOFFICE Docs™는 Ascensio System SIA의 상표입니다. Office Kit은 Ascensio System SIA와 관계가 없고, 보증이나 후원을 받지 않습니다. [상표 정책](https://www.onlyoffice.com/trademark-policy)
- Microsoft Word, Excel, PowerPoint는 Microsoft의 상표입니다. Office Kit은 파일 형식을 가리킬 때만 이 이름을 씁니다.

## 글꼴

- 편집할 때 쓰는 글꼴은 함께 배포하지 않습니다. 편집기는 이 PC에 설치된 글꼴을 읽기만 합니다.
- 다만 변환기 x2t 안에는 ONLYOFFICE core가 넣어 둔 PDF 기본 14 글꼴(URW++ Nimbus Sans · Nimbus Roman · Nimbus Mono, Standard Symbols L, Dingbats)이 들어 있습니다. 이 글꼴은 AGPL-3.0에 PDF·PostScript 문서에 넣어도 된다는 예외가 붙은 자유 글꼴입니다. [고지 (THIRD-PARTY-NATIVE.txt)](local:THIRD-PARTY-NATIVE.txt)
