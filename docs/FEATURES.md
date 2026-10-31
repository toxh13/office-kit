# 형식별 전용 기능

Word·PowerPoint·Excel·한글 각각의 고유 기능을 AI(MCP)와 명령줄에서 쓰게 하는 도구 모음.
공통 도구(`office_read`, `office_replace`, `office_fill`, `office_convert`, `office_create`, `office_page_image` 등)는 그대로 두고,
형식마다 **살펴보기(inspect)** 와 **편집(edit)** 도구를 더했다. 이 문서는 **지금 코드가 하는 것**을 적는다 (한계는 각 절 끝).

## 한눈에

| 형식 | 살펴보기 | 편집 | 필요한 것 | 작업 수 |
|---|---|---|---|---|
| Word (docx·doc·rtf·odt) | `word_inspect` | `word_edit` | 설치된 Word (COM) | 25 |
| PowerPoint (pptx·ppt·odp) | `ppt_inspect` | `ppt_edit` | 설치된 PowerPoint (COM) | 21 |
| Excel (xlsx·xlsm·xls·csv·ods) | `excel_inspect` | `excel_edit` | Excel 우선. 없으면 파일 엔진(exceljs)으로 22개 (xlsx·csv만) | 26 |
| 한글 (hwp·hwpx·hwt) | `hwp_inspect` | `hwp_edit` | 없음 — 내장 엔진 rhwp만 (한컴 COM은 안 씀) | 21 (기존 7 + 새 14) |

쓰는 순서:
- [ ] `…_inspect` 로 번호·이름 확인 (문단·표·슬라이드·도형·시트)
- [ ] `…_edit` 에 작업을 순서대로 모아 **한 번에** 보낸다
- [ ] `office_page_image` 로 결과를 눈으로 본다 (PowerPoint는 쪽 번호 = 슬라이드 번호, 숨긴 슬라이드 포함)

- MCP 도구는 모두 22개: 공통 14개 + 형식별 8개 (`src/mcp/server.mjs` + `src/mcp/tools/*.mjs`).
  공통 중 `office_find_pii`·`office_mask_pii`(개인정보 찾기·가린 사본, `src/core/pii.mjs`)와
  `office_hwpx_batch`(한글 → HWPX 일괄 변환 + 확인, `src/core/batch.mjs`), `office_pdf_tools`(PDF 합치기·쪽 뽑기·쪽 지우기·돌리기,
  `src/core/pdftools.mjs`)는 형식에 묶이지 않은 도구다.
- 명령줄: `office-kit inspect <파일>`, `office-kit edit <파일> <작업.json>` — 형식에 맞는 core 함수로 보낸다 (아래 '명령줄').
- 앱 안 AI: 열린 문서 형식의 `…_inspect` → `…_edit` 를 먼저 쓰도록 안내받는다 (`src/app/ai.mjs` 지침).
- 앱: 새 문서를 만들고, 한글 문서와 Word·Excel·PowerPoint 형식(docx·xlsx·pptx) 문서는 **직접 편집**(앱 안 편집 창)으로도 고친다 — 아래 '앱에서'.

---

## 공통 규칙

**입력과 저장**
- 편집 입력: `{ path, operations: [...], output?, overwrite? }` (Excel은 `engine?`도).
- 작업을 **순서대로 모두 적용한 뒤 한 번만 저장**한다. 하나라도 실패하면 **아무것도 저장하지 않는다**
  ("N번째 작업(op) 실패: 이유 — 저장하지 않았습니다"). 실패 시 원본이 바이트 그대로인지 네 형식 모두 시험한다.
- `output` 생략 = 원본에 저장. `output`이 이미 있는 다른 파일이면 `overwrite: true`(명령줄 `--overwrite`)일 때만 덮어쓴다.
- MCP는 `OFFICEKIT_WRITE_ROOTS`가 있으면 그 폴더 밖에는 저장하지 않는다 (명령줄은 이 규칙 없음).
- 결과: 작업마다 한 줄 한국어 요약 (+ `참고:` 줄) + `저장: 경로`.

**번호와 위치**
- 번호는 모두 **1부터** (문단 1, 슬라이드 1, 표 1, 행 1). 살펴보기 결과의 번호를 그대로 쓴다.
- 앞 작업이 번호를 바꾸면(삽입·삭제) 뒤 작업은 **바뀐 번호**를 쓴다 — 결과 줄에 새 번호를 알려 준다.
- 위치 `at`: `'end'`(기본) \| `'start'` \| `{ after_paragraph: n }` \| `{ before_paragraph: n }` \| `{ after_text: '찾을 글' }` (Word·한글).
- 글자 찾기 `find`: 대소문자 구분, 첫 번째만. `all: true`면 전부. **못 찾으면 실패**(저장 안 함).

**단위·색·그림**
- Word·한글 쪽 설정·들여쓰기는 **mm**, 글자 크기·문단 간격은 **pt**.
- PowerPoint 위치·크기는 **pt** 또는 `'10%'` 같은 슬라이드 비율 문자열.
- Excel 차트 크기는 pt, 열 너비는 글자 수, 행 높이는 pt.
- 색: `'#rrggbb'` (Word는 `#rgb`와 red·blue 같은 색 이름 몇 개도 받음).
- 그림 `image_path`: png·jpg·jpeg·gif·bmp. Word·PowerPoint는 **파일 머리(시그니처)** 도 확인한다 (이름만 .png인 파일 거절).
  읽기만 하므로 어느 폴더든 된다. 상대 경로는 MCP 서버 작업 폴더·명령줄 현재 폴더 기준.

**스타일 키 (Word)**
- `normal`, `title`, `subtitle`, `heading1`~`heading9`, `quote`, `list_bullet`, `list_number`, `caption`, `no_spacing`, `toc_heading`
  → WdBuiltinStyle 번호로 바꿔서 UI 언어와 상관없이 동작. 영어 이름('Heading 1')도 받는다. 그 밖의 문자열은 문서 안 스타일 이름.

**입력 오류**
- 오류는 한국어. 스키마는 네 형식 모두 zod 한국어 로캘(`z.locales.ko()`), 그다음 core가 범위·조합을 한국어로 확인한다.
- 모르는 인수 이름:

| 형식 | 모르는 인수 | 확인하는 곳 |
|---|---|---|
| PowerPoint | 거절 | `pptOpSchema` (core, MCP·명령줄 공통) |
| Excel | 거절 ("모르는 인수입니다: x — 쓸 수 있는 인수: …") | MCP 스키마 + core `checkKeys` |
| 한글 | 거절 (같은 형식의 안내) | `hwpOp` (tools/hwp.mjs, MCP·명령줄 공통) |
| Word | **무시** (조용히 버림. 서식이 하나도 안 남으면 "바꿀 서식이 없습니다") | MCP 스키마 + core `normalizeWordOps` |

**프로그램이 없으면** 대안을 알려 준다: "Word가 설치돼 있어야 합니다 — 대신 office_replace / office_fill / office_create는 쓸 수 있습니다".

---

## 엔진과 시간 제한

| 형식 | 살펴보기 | 편집 |
|---|---|---|
| Word | Word (읽기 전용으로 엶) | Word |
| PowerPoint | PowerPoint (읽기 전용으로 엶) | PowerPoint |
| Excel | `auto`: xlsx·xlsm·csv는 파일 엔진, xls·ods는 Excel | `auto`: Excel이 있으면 Excel, 없으면 파일 엔진. csv→csv 값 작업(수식 없음)은 Excel이 있어도 파일 엔진 |
| 한글 | rhwp | rhwp |

한글 → PDF (`office_convert`, 앱 **PDF로 저장**, `office-kit convert`): 한컴오피스 + 자동화 보안 모듈이 있으면 한컴, 없으면 rhwp 명령줄 프로그램
(`vendor/rhwp-cli/rhwp.exe`, 미리보기와 같은 모양, 글자 선택·검색 가능). 결과 줄에 `엔진: 내장(rhwp)`과 한컴과 다를 수 있다는 안내가 붙는다.

시간 제한 (한 번의 COM 호출, `src/core/*.mjs`):

| 호출 | 한도 |
|---|---|
| Word 살펴보기 | 300초 |
| Word 편집 | 300초 + 작업당 2초 + `all:true` 작업당 120초 (최대 900초) |
| PowerPoint 살펴보기·편집 | 300초 |
| Excel 살펴보기·편집 (Excel) | 300초 |
| 형식 변환·PDF (`office_convert`, Word·Excel `office_page_image`의 PDF 단계) | 180초 |
| PowerPoint 슬라이드 그림 (`office_page_image`) | 180초 |
| 한글 → PDF, 한컴오피스 없이 (`office_convert`, rhwp 명령줄 프로그램 — COM 아님) | 60초 + 쪽당 5초 (최대 20분) |
| 한글·Excel 파일 엔진 | 한도 없음 (Node 안에서 처리) |

- 한도를 넘으면 PowerShell과 그 호출이 띄운 Office 프로세스를 끝낸다. Word 편집은 그때 남은 `~$` 소유자 파일도 지운다.
- 잠금(다른 작업이 끝나길 기다리는 시간)과 한도의 관계는 [`ARCHITECTURE.md`](ARCHITECTURE.md) '시간 초과와 잠금'.

---

## 살펴보기 출력 한도

AI 대화가 넘치지 않게 한 번에 보여 주는 양을 정해 두었다. 넘치면 **이어 볼 위치**를 알려 준다.

| 도구 | 나눠 보기 | 한 번에 보이는 것 |
|---|---|---|
| `word_inspect(path, offset?, limit?)` | 문단 `offset`(건너뛸 수, 0부터) · `limit`(기본 200, 최대 1000). 끝에 `[계속] 다음 문단은 offset=N` | 문단 글 200자, 구역 20개, 표 50개(첫 행 12칸·칸마다 60자), 쓰인 스타일 80개, 메모 100개(**offset=0일 때만**), 변경 내용 20개(60자), 머리말·꼬리말 200자 |
| `ppt_inspect(path, slide?, offset?, limit?)` | `slide` = 한 장만 자세히. 아니면 `offset`(건너뛸 장 수) · `limit`(기본 30, 최대 200). 글이 약 60,000자를 넘으면 그 앞 슬라이드까지만 보이고 `이어서 보려면 offset=N` | 슬라이드마다 도형 60개 (`slide`를 주면 500개), 도형·제목 글 200자 |
| `excel_inspect(path, engine?)` | 없음 (셀 값은 `office_read`) | 자세히는 처음 시트 100개 (나머지는 이름만 200개까지), 시트마다 표·차트 20개, 이름 정의 50개, 수식 예시 20개 |
| `hwp_inspect(path, offset?, limit?)` | 문단 `offset` · `limit`(기본 150, 최대 1000). 끝에 `… 다음 문단은 offset=N` | 문단 글 200자. 표·누름틀은 **그 문단 범위 안 것만**: 표 50개, 누름틀 100개. 범위 밖 표는 앞쪽은 번호 범위, 뒤쪽은 10개까지 `표 n(문단 m)` |

- 명령줄 `--json`은 원자료를 그대로 준다: 한글은 표·누름틀 전체 목록(첫 행은 범위 안 표만), Excel 파일 엔진은 모든 시트
  (Excel로 읽으면 101번째 시트부터는 이름만).
- 명령줄 PowerPoint 살펴보기는 `--limit`을 주지 않으면 (`--offset`부터) 모든 슬라이드를 글자 한도 없이 보여 준다 (아래 '명령줄').

---

## Word — `word_edit` 작업

| op | 인자 | 하는 일 |
|---|---|---|
| `insert_paragraph` | `text, at?, style?, format?` | 문단 넣기. `format`은 format_text와 같은 글자 서식. style 생략 시 Word에서 Enter를 친 것처럼(제목 뒤면 다음 스타일, 목록은 이어짐, 굵게·형광펜 등은 안 이어짐). 글 안 줄바꿈은 문단 안 줄바꿈(Shift+Enter) |
| `set_paragraph_text` | `paragraph, text` | 문단 글만 바꾸기 (스타일 유지) |
| `delete_paragraph` | `paragraph` \| `find`, `all?` | 문단 지우기 (아래 '지우기 규칙') |
| `apply_style` | `paragraph` \| `paragraphs:[from,to]` \| `find`, `all?`, `style` | 스타일 적용 (글자 스타일이면 찾은 글에만) |
| `format_text` | `find` \| `paragraph` \| `paragraphs`, `all?`, `bold, italic, underline, strike, color, highlight, size, font, superscript, subscript` | 글자 서식. highlight: yellow·green·bright_green·cyan·turquoise·pink·blue·red·dark_blue·teal·dark_green·violet·dark_red·dark_yellow·gray·dark_gray·black·white·none (true = yellow) |
| `format_paragraph` | `paragraph` \| `paragraphs` \| `find`, `all?`, `align(left\|center\|right\|justify\|distribute), line_spacing(배수), line_spacing_pt, space_before, space_after, first_line_indent_mm(음수=내어쓰기), left_indent_mm, right_indent_mm, keep_with_next, page_break_before` | 문단 서식 |
| `set_list` | `paragraph` \| `paragraphs:[from,to]`, `type: 'bullet'\|'number'\|'none'` | 글머리표·번호 목록 (범위를 목록 하나로) |
| `insert_table` | `rows, at?, style?, header_bold?(기본 true), autofit?('window'(기본)\|'content'\|'fixed')` | 표 넣기. style 키: grid·plain·light_grid·light_shading·light_list·medium_shading1 또는 문서 안 표 스타일 이름 (생략 = 격자 테두리) |
| `table_set_cell` | `table, row, col, text` | 칸 고치기 |
| `table_add_row` | `table, values?, after_row?(0=맨 위, 생략=맨 아래)` | 행 추가 (세로 병합 자리 값은 `null`) |
| `table_delete_row` | `table, row` | 행 지우기 (마지막 남은 행이면 표 전체) |
| `insert_image` | `image_path, at?, width_mm?, height_mm?, caption?, caption_label?('그림'), align?(기본 center)` | 그림 넣기 (하나만 주면 비율 유지, 둘 다 없으면 쪽 너비에 맞춰 줄임). 캡션은 그림 아래 '그림 N. 캡션'(번호 자동) |
| `insert_page_break` | `at?` | 쪽 나누기 |
| `set_header` / `set_footer` | `text?, align?, page_numbers?, section?` | 머리말·꼬리말. text 안 `{page}`=쪽 번호, `{pages}`=전체 쪽 수. `page_numbers:true`면 `{page}`가 없을 때 끝에 붙임. 내용을 통째로 바꾼다 |
| `page_setup` | `orientation?, paper?('A4'\|'A3'\|'A5'\|'B4'\|'B5'\|'Letter'\|'Legal'), margins_mm?{top,bottom,left,right}, section?` | 쪽 설정 (section 생략 = 모든 구역) |
| `add_comment` | `find, text, all?, author?` | 메모 달기 |
| `delete_comments` | `all: true` \| `index` | 메모 지우기 |
| `track_changes` | `on` | 변경 내용 추적 켜기/끄기 (켠 뒤의 작업이 변경으로 기록됨) |
| `accept_revisions` / `reject_revisions` | — | 변경 내용 모두 수락/거부 |
| `insert_toc` | `at?, levels?(1~9, 기본 3), title?` | 목차 넣기 (heading 스타일 기준) |
| `update_fields` | — | 목차·쪽 번호·캡션 번호 등 필드 새로 고침 |
| `add_footnote` | `find, text, all?` | 각주 |
| `insert_hyperlink` | `find, url(http(s)://·mailto:·ftp://), all?` | 링크 걸기 |

입력 범위: 한 번에 작업 500개 · `find` 255자(줄바꿈 불가) · 글자 크기 1~1638pt · 줄 배수 0.25~10 · `line_spacing_pt` 1~1584 ·
문단 간격 0~1584pt · 들여쓰기 ±500mm · 여백 0~500mm · 그림 1000mm 이하 · 표 500행·63열·5000칸까지.

**word_inspect** → 쪽 수, 단어·글자 수, 구역별 쪽 설정(용지·방향·여백, 첫 쪽 머리말 따로 여부), 머리말·꼬리말(필드는 `{page}`·`{pages}`로),
변경 내용 추적 상태·개수·목록, 메모 목록, 목차·그림·떠 있는 개체·콘텐츠 컨트롤·책갈피·각주·미주·링크·필드 개수,
쓰인 스타일(키와 함께), 표 목록(번호·행×열·병합 여부·시작 문단·스타일·첫 행), 문단 목록(번호·스타일(키)·글 200자·표 안 위치·목록 종류).
- 세로 병합의 아랫부분 문단은 `[표n 병합된 칸 — …글을 넣을 수 없음]`으로 보인다.
- 구역 나누기가 든 문단은 `[구역 나누기]`, 추적 중 지운 문단은 `[지움 표시됨 — 변경 내용 수락 전]`.
- 작은 문서 약 6~8초, 48쪽(문단 2,100개) 약 14초.

번호 규칙:
- 문단 번호는 Word의 `Paragraphs(n)` 그대로라 **표의 행 끝 표시도 번호를 하나씩 차지한다** (목록에서는 빼고 번호만 건너뜀).
- 표 번호는 본문 맨 바깥 표만 센다.

지우기 규칙 (`delete_paragraph`):
- 보통 문단은 그냥 지운다.
- 구역 나누기가 든 문단·표 바로 앞 문단·문서 마지막 문단은 **앞 문단과 합쳐 지운다** (구역 나누기는 앞 문단 끝에 남겨 쪽 설정 유지).
- 합칠 앞 문단이 없을 때만 **글만 비우고** 이유를 결과 줄에 적는다 (표 칸의 유일한 문단, 앞이 표·구역 나누기 등). 번호는 그대로.
- 변경 내용 추적 중이면 **지움 표시만** 한다 (수락 전까지 번호 그대로).
- 행 끝 표시는 거절(→ `table_add_row`/`table_delete_row`), 세로 병합의 아랫부분도 거절.

동작 메모:
- `find`는 본문에서만 찾고 **목차 안 결과는 건너뛴다**. 목차 안에만 있으면 "찾는 글자가 목차 안에만 있습니다"로 실패.
  `all:true`는 5000곳까지, 넘으면 그 작업 전체가 실패한다 (find를 좁히거나 paragraphs로 나눔).
- `format_text`를 `paragraph`/`paragraphs`로 주면 문단 표시까지 포함한다 → 목록 번호도 같은 서식이 된다.
- `set_list`: 글 앞에 직접 쓴 번호·기호('1.', '-')는 Word가 자동 번호로 바꾸며 글에서 빼므로 결과 줄에 알린다.
- `insert_toc`: 제목(heading1~) 문단이 없으면 한국어로 실패한다 (같은 작업 목록이면 `apply_style`을 앞에).
- `insert_toc`·`update_fields`는 필드를 새로 고치는 동안 변경 내용 추적을 잠시 끈다 (새로 고친 결과는 변경으로 기록되지 않음).
- `table_add_row`·`table_delete_row`는 세로 병합이 있는 표에서도 된다. 병합된 자리의 값은 `null`로 둔다.
  병합의 맨 윗행을 지우면 그 칸 글은 남은 행에 옮긴다. 모든 칸이 위·아래와 병합된 행은 거절.
- `at`이 표 칸 안 문단이면 `insert_paragraph`·`insert_image`는 그 칸 안에, `insert_table`·`insert_toc`·`insert_page_break`는 표 앞뒤에 넣는다.
- 원본에 그대로 저장할 때 파일이 **읽기 전용이거나 다른 프로그램에서 열려 있으면 Word를 띄우기 전에(또는 작업 전에) 거절**한다
  (그대로 두면 Word가 보이지 않는 '다른 이름으로 저장' 창에서 멈춤).

한계:
- `set_header`/`set_footer`: 기본(primary) 머리말·꼬리말만 고친다. 첫 쪽·홀짝 머리말은 `word_inspect`에 보이기만 한다.
- 표 안의 표(중첩 표)는 번호가 없어 고를 수 없다.
- '쓰인 스타일'은 본문에서 찾은 스타일 이름만 보여 준다 (개수 없음 — 큰 문서에서 문단마다 세면 너무 느림).
- `all:true`가 필드 사이를 오래 건너야 하는 문서(찾은 곳 사이에 필드 수백 개)는 느려진다.
- 시험한 Word는 UI 언어가 영어(1033)다. 스타일은 언어와 상관없는 번호로 다루지만 한국어 Word에서는 아직 확인하지 않았다.
- 메모·변경 내용은 PDF 쪽 그림에 나오지 않는다 → `word_inspect`로 확인.

---

## PowerPoint — `ppt_edit` 작업

| op | 인자 | 하는 일 |
|---|---|---|
| `add_slide` | `layout?('title'\|'title_content'(기본)\|'section'\|'two_content'\|'comparison'\|'title_only'\|'blank' 또는 ppt_inspect의 레이아웃 이름), position?, title?, body?(문자열 또는 [{text, level}]), notes?` | 슬라이드 추가. 마스터에 그 키의 레이아웃이 없으면 표준 레이아웃을 새로 만든다 (결과 줄에 알림) |
| `delete_slide` | `slide` | 슬라이드 지우기 |
| `duplicate_slide` | `slide, position?` | 복제 (기본: 바로 뒤) |
| `move_slide` | `slide, to` | 순서 바꾸기 |
| `set_title` / `set_body` | `slide, text` / `slide, bullets, placeholder?` | 제목·본문 개체 틀 채우기. placeholder: 몇 번째 내용 칸 (두 콘텐츠·비교의 오른쪽 = 2). 내용 → 본문 → 부제목 개체 틀 순서로 센다 |
| `set_notes` | `slide, text` | 발표자 노트 (기존 노트를 바꿈) |
| `set_text` | `slide, shape(이름 또는 번호), text` | 도형 글 바꾸기 (줄바꿈 = 새 문단) |
| `add_textbox` | `slide, text, left, top, width, height, size?, bold?, color?, align?(left\|center\|right\|justify), font?` | 글상자. 높이는 그대로 두고 글이 더 길 때만 늘린다 (결과 줄에 알림) |
| `add_image` | `slide, image_path, left, top, width?, height?` | 그림 (하나만 주면 비율 유지, 둘 다 없으면 원래 크기, 슬라이드를 넘으면 줄임) |
| `add_table` | `slide, rows, left, top, width, height?, header_bold?(기본 true)` | 표 |
| `table_set_cell` | `slide, shape, row, col, text` | 표 칸 |
| `add_shape` | `slide, shape('rectangle'\|'rounded_rectangle'\|'oval'\|'arrow'\|'line'), left, top, width, height, fill?, line?, text?, color?` | 도형. fill·line은 '#rrggbb' 또는 'none'. `color`는 글자색(text와 함께) — 생략하면 채우기 없음은 tx1, 밝은 채우기는 dk1. line은 (left,top)→(left+width, top+height), 글 불가 |
| `add_chart` | `slide, type('column'\|'bar'\|'line'\|'pie'), categories, series[{name, values}], left, top, width, height, title?` | 차트 (values 개수 = categories 개수). 데이터는 차트 안 Excel 데이터로 저장 → PowerPoint에서 "데이터 편집" 가능 |
| `format_text` | `slide?(없으면 전체), find, all?, bold?, italic?, underline?, color?, size?, font?` | 글자 서식 (표·그룹 안 글 포함, 서식 하나 이상) |
| `delete_shape` | `slide, shape` | 도형 지우기 |
| `set_background` | `slide\|'all', color` | 배경색 ('all'이면 슬라이드 마스터도 → 나중에 넣는 슬라이드도 같은 색) |
| `set_transition` | `slide\|'all', effect('none'\|'fade'\|'push'\|'wipe'\|'split'\|'cover'), duration_s?(60초까지)` | 화면 전환 (fade=FadeSmoothly, push=PushUp, wipe=WipeLeft, split=SplitVerticalOut, cover=CoverLeft) |
| `hide_slide` | `slide, hidden?(기본 true)` | 숨기기 (쇼에서 건너뜀) |
| `apply_theme` | `theme_path(.thmx\|.potx)` | 테마 적용 (암호 없는 ZIP 꾸러미만) |
| `set_slide_size` | `preset('16:9'(960×540pt)\|'4:3'(720×540pt))` | 슬라이드 크기. **좁히면** 내용을 비율에 맞춰 줄여 가운데로, **넓히면** 내용 크기는 그대로 가운데로 옮긴다 (양옆이 빔) |

**ppt_inspect** → 슬라이드 크기(pt·비율), 디자인, 쓸 수 있는 레이아웃(이름과 키 — 레이아웃 이름은 지역화되므로 종류로 키를 붙임, 없는 키는 따로 알림),
슬라이드마다 번호·레이아웃·제목·숨김·화면 전환·발표자 노트, 도형 목록(번호·이름·종류·개체 틀 종류·위치·크기(pt)·글 200자·표 행×열·차트 종류·그룹 항목 수).
결과를 `.ppt`·`.odp`로 저장해도 된다 (`SaveAs`).

쪽 그림 (`office_page_image`):
- PowerPoint 문서는 PDF를 거치지 않고 슬라이드를 바로 그림으로 만든다 (`office.ps1` `slideImage` = `Slide.Export`, 읽기 전용으로 엶).
  → **쪽 번호 = 슬라이드 번호** (숨긴 슬라이드 포함), 전체 쪽 수 = 슬라이드 수.
- 그림 크기는 PowerPoint 기본 크기(보통 960pt 너비 → 1280px)까지. 더 작은 dpi는 줄이고, 더 큰 dpi는 기본 크기 그대로.
- 같은 파일·슬라이드는 다시 그리지 않고 기억해 둔 그림을 쓴다 (파일이 바뀌면 다시 그림).

암호:
- 열기 암호(암호화된 pptx·odp)와 암호가 걸린 .ppt는 **열기 전에 거절**한다 (PowerPoint가 보이지 않는 암호 창에서 멈추므로).
- 쓰기 암호만 걸린 pptx: 살펴보기는 된다. 편집은 원본 자리에는 안 되고 `output`(명령줄 `-o`)으로 다른 파일에만.
  결과가 .pptx면 같은 쓰기 암호를 다시 건다(`참고:` 줄), .ppt·.odp면 암호가 빠진다(`참고:` 줄).

한계:
- `add_chart`: PowerPoint는 창 없는 프레젠테이션에 차트를 못 넣어서 작업 동안 **최소화된 PowerPoint 창**이 열린다
  (사용자가 PowerPoint를 쓰고 있으면 그 안에 잠깐 보임). 차트 데이터용 Excel도 잠깐 떴다 숨겨진다.
  끝나면 그 호출이 띄운 Excel만 정리한다. 시간 초과로 끊기면 그 Excel이 남을 수 있다 (com.mjs는 PowerPoint PID만 앎).
- `add_chart` 원형(pie)은 첫 계열만 그린다 (결과 줄에 알림).
- 그룹 안 도형은 `format_text`로는 서식을 바꿀 수 있지만 `set_text`·`delete_shape`로 고를 수 없다 (맨 위 도형만).
- `office_convert`로 만든 PDF에는 숨긴 슬라이드가 없다 (PowerPoint의 PDF 저장 동작).
- 쓰기 암호만 걸린 **.ppt**도 거절한다 (열기 암호와 구별하지 못함 — 살펴보기·쪽 그림 포함).
- 쓰기 암호가 걸린 문서(Word·Excel·PowerPoint)는 변환·쪽 그림·다른 이름 저장은 읽기 전용으로 열어 처리한다.
  그 자리에 저장(output 없이 편집)은 바로 거절한다 → output으로 다른 파일에 저장. (.xls·ODF는 미리 확인하지 못해 멈출 수 있음)

---

## Excel — `excel_edit` 작업

○ = 파일 엔진(Excel 없이)도 됨, × = Excel 필요.

| op | 인자 | 하는 일 | 파일 엔진 |
|---|---|---|---|
| `add_sheet` | `name, after?` | 시트 추가 (이름 31자, `\ / ? * [ ] :` 불가) | ○ |
| `rename_sheet` | `sheet, name` | 이름 바꾸기 (다른 시트 수식 참조도 고침) | ○ |
| `delete_sheet` | `sheet` | 지우기 (지운 시트 참조는 #REF!, 보이는 마지막 시트는 거절) | ○ |
| `copy_sheet` | `sheet, name` | 복사 (원본 바로 뒤) | × |
| `move_sheet` | `sheet, position` | 순서 바꾸기 | ○ |
| `set_cells` | `sheet?, cells{A1: 값 또는 '=수식', null=지우기}` | 셀 값·수식 | ○ |
| `set_range` | `sheet?, range('A1' 시작 셀), values[[...]]` | 2차원 값 붙여 넣기 (100만 칸까지) | ○ |
| `clear` | `sheet?, range, what?('contents'(기본)\|'formats'\|'all')` | 지우기 | ○ |
| `insert_rows` / `delete_rows` | `sheet?, at(행 번호), count?(기본 1)` | 행 (수식·병합·표·조건부 서식·입력 제한·이름 정의·필터·그림·메모·링크도 함께 옮김) | ○ |
| `insert_columns` / `delete_columns` | `sheet?, at('C' 또는 3), count?` | 열 (위와 같음) | ○ |
| `format` | `sheet?, range, bold, italic, underline, font, size(1~409), color, fill('#RRGGBB'\|'none'), number_format, align(general\|left\|center\|right\|justify\|center_across\|fill), valign(top\|center\|bottom\|justify), wrap, border('thin'\|'medium'\|'thick'\|'none'\|'outline')` | 서식 | ○ |
| `column_width` | `sheet?, columns('A' 또는 'A:C'), width(0~255)\|'auto'` | 열 너비 | ○ ('auto'는 근사) |
| `row_height` | `sheet?, rows('5' 또는 '1:3'), height(0~409pt)\|'auto'` | 행 높이 | ○ ('auto'는 근사) |
| `merge` / `unmerge` | `sheet?, range` | 병합 (두 칸 이상, 버린 값 개수를 알려 줌) / 풀기 | ○ |
| `freeze` | `sheet?, cell('B2'\|null)` | 틀 고정 ('A2' = 1행만, null = 해제). Excel 엔진: 화면 없는 Excel은 틀 고정을 못 거므로 .xlsx·.xlsm은 Excel이 저장한 사본의 시트 보기에 직접 써서 옮긴다 (차트가 있어도 됨). 다른 형식(.xls·.ods)으로 저장하면 안 될 수 있다 (한국어 오류) | ○ |
| `sort` | `sheet?, range, by(열 문자, 범위 안), order?('asc'(기본)\|'desc'), header?` | 정렬 (header 생략 = Excel이 판단, 병합 칸 있으면 거절) | × |
| `autofilter` | `sheet?, range\|null` | 자동 필터 | ○ |
| `add_table` | `sheet?, range, name?, style?(기본 TableStyleMedium2)` | 표(ListObject), 첫 행은 늘 머리글 | ○ |
| `add_chart` | `sheet?, range` 또는 `values, categories?`, `type('column'\|'bar'\|'line'\|'pie'\|'scatter'\|'area'), title?, at?(셀), width?(pt, 기본 360), height?(pt, 기본 216), series_in?('columns'\|'rows')` | 차트. 첫 행·첫 열은 이름표. series_in 생략 = Excel처럼(행이 열보다 많으면 columns). 결과 줄에 계열 수와 방향. title 생략 = 제목 없음 (계열 하나면 그 이름). 떨어진 열: `range:'A3:A10,H3:H10'`(같은 행들에 나란히, Ctrl로 고른 것처럼) 또는 `categories:'A3:A10', values:['H3:H10', …]`(값 한 줄 = 계열 하나, 첫 칸이 글자면 계열 이름, categories와 칸 수가 같게, series_in 없이) | × |
| `conditional_format` | `sheet?, range, rule{type('greater_than'\|'less_than'\|'between'\|'equal'\|'text_contains'\|'duplicate'\|'color_scale'\|'data_bar'), value?, value2?, fill?, font_color?}` | 조건부 서식 (기본 강조 '#FFC7CE' 배경·'#9C0006' 글자, value에 '=$B$1' 수식도) | ○ |
| `data_validation` | `sheet?, range, list?(['예','아니오'] 또는 '=$E$1:$E$5')` 또는 `type('whole'\|'decimal'), min?, max?` | 입력 제한. 목록 항목에 쉼표·큰따옴표 불가, 합쳐 255자까지 | ○ |
| `named_range` | `name, range('시트!A1:B2'), sheet?(range에 시트가 없을 때)` | 이름 정의 (통합 문서 범위, 있으면 바꿈, 표 이름과 같으면 거절) | ○ |
| `recalculate` | — | 다시 계산 | × |

엔진 (`engine`):
- `'auto'`(기본) = Excel이 있으면 Excel(모든 작업, 원본 그대로), 없으면 파일 엔진.
  단 **csv → csv 값 작업(수식 없음)은 Excel이 있어도 파일 엔진** — 인코딩·구분자·줄바꿈을 지키고 '0123' 같은 값이 숫자로 안 바뀌게.
- `'app'` = Excel 필수. `'file'` = Excel 없이 (× 작업 불가).
- 파일 엔진은 xlsx·csv만, **같은 형식으로만** 저장한다. xlsm은 살펴보기만 (고치면 매크로가 사라짐). xls·ods는 Excel 필요.
- 시트는 이름 또는 번호(1부터), 생략하면 첫 시트. 결과 마지막 줄에 쓴 엔진이 나온다.

**excel_inspect(path, engine?)** → 시트마다 이름·번호·숨김·사용 범위·행×열·틀 고정·표·차트·병합 수·자동 필터·조건부 서식 수·입력 제한 여부,
이름 정의(시트 범위면 `[시트 이름]`), 활성 시트, 수식 개수와 예시, 파일 엔진으로 고칠 수 없는 내용(있으면 `주의:` 줄).
`auto`는 xlsx·xlsm·csv를 파일 엔진으로(빠름), xls·ods는 Excel로 읽는다.

값·수식 규칙:
- 문자열은 늘 글자로 저장한다 ('120', '1/2'도 글자 — Excel 엔진은 앞에 ' 를 붙여 씀). 숫자는 JSON 숫자로.
- `'='`로 시작하면 수식(영어 함수 이름, 쉼표 구분). 글자 그대로 '='로 시작하려면 앞에 `'`.
- 수식 문법(괄호·따옴표 짝, `;` 구분자, 연산자로 끝남 등)은 **두 엔진 모두** 미리 확인한다. 함수 이름은 보지 않는다 (틀리면 #NAME?).
- Excel 엔진은 `Formula2`(동적 배열)로, 파일 엔진은 일반 수식으로 쓴다. 파일 엔진은 Excel 2010 뒤 함수에 `_xlfn.`을 붙이고,
  넘치는 동적 배열·LAMBDA 계열(FILTER, SORT, UNIQUE, SEQUENCE, LET, LAMBDA …)은 거절한다.
- 병합 영역의 왼쪽 위가 아닌 칸에 쓰면 두 엔진 모두 실패한다.

csv:
- csv로 저장할 때는 값 작업만 된다: `set_cells`, `set_range`, `clear`, 행·열 넣기/지우기 (+ Excel로는 `sort`·`recalculate`). 시트·서식·표 작업은 거절.
- 파일 엔진: BOM·구분자·줄바꿈을 지킨다. 수식은 저장할 수 없다(거절). CP949 csv는 UTF-8(BOM)로 바꿔 쓰고 `참고:`로 알린다.
- Excel 엔진: BOM 없는 UTF-8 csv는 BOM 없이, LF 줄바꿈도 그대로. CP949는 UTF-8(BOM)로(`참고:`).
  xlsx 등 → csv는 작업이 고친 시트를 저장하고, 여러 시트를 고치면 거절한다.

표:
- `add_table`은 이미 있는 표·이름이면 거절(두 엔진), 병합 칸·다른 표와 겹치면 거절.
  파일 엔진은 표 스타일 이름 형식(TableStyleLight1~21·Medium1~28·Dark1~11)을 확인하고, 머리글 칸 수식은 저장된 결과를 글자로 바꾸며 결과가 없으면 거절.
- `delete_rows`로 표의 머리글 행·요약 행만 지우는 것은 두 엔진 모두 거절한다 (표 전체를 지우면 표도 사라짐).
- 파일 엔진은 표의 계산 열·필터 단추를 지키고, 표 안에 넣은 행에 계산 열 수식을 채운다.

한계 (파일 엔진):
- `sort`·`add_chart`·`copy_sheet`·`recalculate`는 Excel이 있어야 한다 (한국어 오류, 이 PC에 Excel이 있으면 `engine:'auto'`를 권함).
- 아래 내용이 있는 파일은 **고치지 않는다** (아무것도 쓰지 않음, `excel_inspect`가 `주의:`로 알려 줌):
  차트·차트 시트, 피벗, 슬라이서·시간 표시 막대, 스파크라인, 외부 파일 참조·외부 데이터 연결, 데이터 모델, VBA(xlsm),
  스레드 메모, 도형·글상자·SmartArt·최신 차트, OLE·양식 컨트롤·ActiveX, 셀 안 그림, 동적 배열 수식,
  사용자 지정 표 스타일·색 팔레트, 통합 문서 보호·쓰기 암호·편집 허용 범위, 지원하지 않는 조건부 서식(중복 값·시작 문자 등, 확장 규칙),
  다른 시트를 참조하는 입력 제한, 수식을 쓰는 날짜 입력 제한, 머리글·바닥글 그림, 바깥 파일 안 위치로 가는 링크, 표 요약 행 사용자 수식.
- 수식 결과는 계산하지 않는다: 값·구조를 바꾸면 수식 칸에 저장된 옛 결과를 비우고(`참고:`), `fullCalcOnLoad`를 켜서 Excel이 열 때 다시 계산.
- 'auto' 너비·높이는 글자 수·크기로 어림한다. text_contains·duplicate는 수식 규칙(SEARCH·COUNTIF)으로 저장된다(보기는 같음). 데이터 막대는 옛 모양.

한계 (공통):
- `named_range`는 통합 문서 범위만 (시트 범위 이름은 못 만듦). `add_table`은 첫 행을 늘 머리글로 본다.
- inspect의 입력 제한은 있음/없음만. Excel은 자동 필터 범위를 이어진 데이터까지 넓혀 보고할 수 있다.

---

## 한글 — `hwp_edit` 작업 (rhwp)

기존 7개 (문서 끝에 붙이기·바꾸기·채우기):

| op | 인자 | 하는 일 |
|---|---|---|
| `add_paragraph` | `text, style?{bold, italic, underline, size, color, font, align(left\|center\|right\|justify)}` | 문서 끝에 문단 |
| `add_heading` | `text, level?(1~3, 기본 1)` | 문서 끝에 제목 (18·15·13pt 굵게) |
| `add_table` | `rows` | 문서 끝에 표 (첫 행 굵게) |
| `add_page_break` | — | 문서 끝에 쪽 나누기 |
| `add_field` | `name, label?` | 문서 끝에 누름틀 ("label: [누름틀]") |
| `replace` | `find, replace` | 문서 전체 찾아 바꾸기 (대소문자 구분) |
| `fill` | `values{이름: 값}` | 누름틀 채우기 — **같은 이름은 모두**(결과에 `이름(N곳)`). 없는 이름은 `{{이름}}` 자리표시자(`{{ 이름 }}`도)를 바꿈 |

새로 더한 14개:

| op | 인자 | 하는 일 |
|---|---|---|
| `insert_paragraph` | `text, at?, style?(문서 스타일 이름, 기본 '바탕글'), like_paragraph?(n), format?, align?` | 원하는 위치에 문단 (줄바꿈마다 문단 하나). style·like_paragraph가 없으면 앞 문단 모양을 따르지 않는다 |
| `set_paragraph_text` | `paragraph` \| `find`, `text` | 문단 글 바꾸기 (모양 유지, 줄바꿈이면 같은 모양 문단을 뒤에) |
| `delete_paragraph` | `paragraph` \| `find` | 문단 지우기 (표·그림이 든 문단이면 함께) |
| `format_text` | `find, all?, bold, italic, underline, strike, size, color, font, highlight('#rrggbb'), superscript, subscript` | 글자 서식 (본문·표 칸) |
| `format_paragraph` | `paragraph` \| `paragraphs:[처음,끝]` \| `find`, `align(left\|center\|right\|justify\|distribute), line_spacing(1.6 또는 '160%', 50~500%), left_indent_mm, right_indent_mm, first_line_indent_mm(음수=내어쓰기), space_before, space_after(pt), keep_with_next, page_break_before` | 문단 서식 |
| `apply_style` | `paragraph` \| `paragraphs` \| `find`, `style` | 스타일 (문단·글자 모양이 스타일 것으로 바뀜 — 한컴 방식, 직접 서식은 사라짐) |
| `table_set_cell` | `table, row, col, text` | 기존 표 칸 (병합 칸은 덮는 자리 아무거나) |
| `table_add_row` | `table, after_row?(0=맨 위, 생략=맨 아래), values?` | 행 추가. 세로 병합에 덮이는 자리에 값을 주면 **거절** (그 자리는 `''`) |
| `table_delete_row` | `table, row` | 행 지우기 (하나뿐인 행은 거절 → 표가 든 문단을 `delete_paragraph`) |
| `set_header` / `set_footer` | `text, page_numbers?, align?` | 머리말·꼬리말 (`{page}`·`{pages}`, `page_numbers:true`면 끝에 쪽 번호 — 빈 text면 '- 1 -' 모양, 빈 text만 = 지우기, 모든 구역에 적용, 쪽 번호만 있으면 가운데) |
| `add_footnote` | `find, text` | 본문 글 첫 번째 뒤에 각주 (표 안에만 있으면 거절) |
| `insert_image` | `image_path, at?, width_mm?, height_mm?, align?(기본 center), description?` | 그림 하나만 든 새 문단. 크기를 안 주면 원래 크기(96dpi). 본문보다 **넓거나 높으면** 비율을 지키며 본문 안에 맞춰 줄임 (결과 줄에 알림) |
| `page_setup` | `orientation?, paper?('A4'\|'A3'\|'A5'\|'B4'\|'B5'\|'Letter'\|'Legal'), margins_mm?{top,bottom,left,right,header,footer,gutter}(0~300), section?` | 쪽 설정 (한컴 방식: 본문은 위 여백 + 머리말 아래에서 시작) |

- `find`(paragraph 대신)는 그 글이 든 **첫 본문 문단**을 고른다. 번호 범위는 core가 한국어 오류로 확인한다.
- 저장은 .hwp·.hwpx (.hwt로 주면 원본 형식 그대로). 결과 마지막 줄에 쪽 수.

**hwp_inspect** → 쪽 수, 구역별 쪽 설정(용지·방향·여백 mm), 머리말·꼬리말,
문단 목록(번호·글 200자·스타일·정렬·글꼴·크기·표/그림/누름틀 표시), 표 목록(번호·문단·행×열·첫 행), 누름틀(값·문단), 스타일 이름, 그림·각주·도형 개수.
표·누름틀 목록은 이번 문단 범위 안 것만 (위 '살펴보기 출력 한도').

번호 규칙: 문단은 **모든 구역의 본문 문단을 문서 순서대로 1부터** 센다 (구역 2는 구역 1 끝에서 이어짐).
표 칸·머리말·각주 안 문단은 세지 않는다. 표가 든 문단은 문단 하나로 세고 `[표 n(r행×c열)]`로 보인다. 표는 바깥 표만 문서 순서대로.

한계:
- `add_footnote`: 각주 글이 붙는 자리의 글자 모양을 따른다 (굵은 글 뒤면 굵은 각주 — rhwp에 각주 안 서식 기능이 없음).
- `set_paragraph_text`: 표·그림·누름틀·각주가 든 문단은 거절 (→ replace·fill·table_set_cell).
- `insert_paragraph`를 'start'나 구역 첫 문단 앞에 넣을 때, 그 문단에 표·그림·누름틀이 있으면 거절 (구역 정의를 옮길 수 없음).
- `delete_paragraph`: 구역에 하나 남은 문단, 누름틀 등이 든 구역 첫 문단은 거절.
- `set_header`/`set_footer`: 저장 후 다시 열면 rhwp가 쪽 번호 자리를 빈칸으로 보고한다 (`hwp_inspect`에 빈칸, 그린 쪽에는 번호가 나옴). 모든 쪽 머리말을 넣으면 홀짝 머리말은 지운다.
- rhwp의 hwpx 저장은 줄 배치 정보를 빼먹어 그림만 든 문단이 미리보기에 안 그려질 수 있다 → `Hwp.openForView`가 화면용 사본에서만 다시 배치.
- 첫 쪽 제목 바로 아래에 본문 높이에 맞춘 큰 그림을 넣으면 rhwp 배치에서는 1쪽에 남아 아래 여백을 약 14mm 넘긴다 (잘리지는 않음.
  문서 중간에 넣으면 다음 쪽으로 감).
- 한컴오피스에서 어떻게 보이는지는 확인하지 못했다 (한컴 COM 못 씀).
- `format_text` 등 `find`는 머리말·꼬리말·각주 안 글을 찾지 않는다. 표 안의 표는 번호가 없다.
- 한컴 보안 모듈이 등록되면 나중에 한컴 COM 경로를 더할 수 있다 (지금은 rhwp만).

---

## 앱에서 — 새 문서 · 직접 편집

| 기능 | 어디서 | 형식 | 필요한 것 |
|---|---|---|---|
| 새 문서 (빈 문서) | 상단 `새 문서` (Ctrl+N), 빈 화면, 대화창 | hwpx · hwp · docx · pptx · xlsx | 없음 — 내장 엔진으로 만든다 |
| 새 문서 + AI가 채우기 | 새 문서 창 '시작 방법' → AI에게 내용 부탁하기 | 위와 같음 | Claude Code 또는 Codex |
| 직접 편집 (한글) | 상단 `직접 편집` (Ctrl+E), 새 문서 창 '만든 뒤 바로 편집 창 열기' (기본 켜짐) | hwp · hwpx · hwt | 없음 — 앱에 든 rhwp-studio, 인터넷 없이 |
| 직접 편집 (문서 · 표 · 슬라이드) | 위와 같음 → 형식에 따라 '문서 편집' · '표 편집' · '슬라이드 편집' 창 | docx · xlsx · pptx | 없음 — 앱에 든 편집기(ONLYOFFICE Docs 기반, AGPL), 인터넷 없이, MS Office 없이 |

- 직접 편집 저장: hwp·hwt → HWP, hwpx → HWPX (원래 형식). 다른 이름으로 저장은 .hwpx / .hwp. 덮어쓰기 전 내용은 문서 버전('직접 편집 저장 전')으로 남는다 (문서마다 최근 30개, 원래대로 ▾에서 되돌림).
  저장하는 동안 친 글은 '저장 안 됨'으로 남아 닫을 때 묻는다. 파일의 보안 설정(나만 읽기 등)은 저장해도 그대로다.
- 문서 안 링크(http·https·mailto)를 누르면 "기본 브라우저로 열까요?"를 물은 뒤 연다. 이 PC·사설망 주소(127.x, 192.168.x 등)는 열지 않는다.
- 원래대로: 편집 창에 저장 안 한 내용이 있으면 파일이 이미 원본과 같아도 편집 창이 다시 불러오기 / 내 편집 유지를 묻는다.
- AI와 함께: AI 도구는 **저장된 파일**을 고친다. 편집 창에 저장 안 한 내용이 있으면 앱이 보내기 전에 묻고(편집 창 보기 / 그대로 보내기 / 취소),
  AI가 파일을 고치면 편집 창이 다시 불러온다 (편집 창에서 고친 것이 있으면 다시 불러오기 / 내 편집 유지를 묻는다).
- Word·Excel·PowerPoint 형식은 아래 '문서 · 표 · 슬라이드 직접 편집'.
- 한계: 편집기는 rhwp-studio v0.8.7 그대로다. 한컴오피스와 모양·저장 결과가 다를 수 있고, 기본 설치에는 웹 글꼴이 없어 PC에 있는 글꼴로 보인다.
  - 병합된 칸이 많은 표에 줄·칸을 넣으면 표 모양이 흐트러질 수 있어요 (Ctrl+Z로 되돌리기). 흐트러진 채 저장하면 파일에도 그대로 남는다.
    복잡한 서식 문서(공문 양식 등)는 '프로그램에서 열기'로 고친다. (rhwp 0.8.7 `insertTableRow`의 병합 줄 처리 — 위로 알릴 문제)
  - 일부 문서는 편집 창에서 표 바로 다음 문단이 표의 마지막 줄 위에 겹쳐 그려진다 (rhwp-studio 0.8.7 화면 그리기 문제 — 저장한 파일과 주 창 미리보기는 그대로).
- 구조(두 사이트, 네트워크 차단, IPC, 저장·감시·닫기 흐름): [`ARCHITECTURE.md`](ARCHITECTURE.md) '한글 편집 창'.

### 문서 · 표 · 슬라이드 직접 편집 (docx · xlsx · pptx)

| 할 일 | 어떻게 |
|---|---|
| 열기 | docx·xlsx·pptx를 연 뒤 `직접 편집` (Ctrl+E). 문서마다 창 하나 — 제목 `<파일> - 문서 편집 / 표 편집 / 슬라이드 편집 (Office Kit)` |
| 저장 | `Ctrl+S` 또는 위 막대 `저장` → 원래 형식 그대로. 덮어쓰기 전 백업, 미리보기 새로 고침 |
| 다른 형식 | 편집기의 '다른 형식으로 내려받기'(PDF 등) → 저장 대화상자. '다른 이름으로 저장'은 같은 형식으로만 |
| 인쇄 | PDF로 만들어 기본 PDF 앱으로 연다 |
| 자동 저장 | 원본에 쓰지 않는다 → 복구 사본. 다음에 열 때 '자동 저장된 편집 내용이 있습니다 — 복구 / 버리기' |
| 정보 · 라이선스 | 위 막대 → 원 개발사(Ascensio System SIA), 수정판일 수 있다는 안내, AGPL 전문·추가 조항, 이 판의 소스 주소 |

- 글꼴: 이 PC에 설치된 글꼴을 쓴다 (설치 파일에 글꼴 없음). '맑은 고딕'·'바탕' 같은 한국어 이름으로도 찾고, Calibri·Arial 문단 안 한글은 맑은 고딕으로 보인다.
- 저장 뒤 고치기 (`src/core/ooxml-fixups.mjs`, 저장할 바이트에만): pptx 노트 마스터 테마를 따로 줌(없으면 PowerPoint 복구 창), xlsx 안 쓰는 테마 지움,
  docx 표에 더한 행의 혼자 있는 세로 병합 표시 지움(`word_edit` 표 작업이 깨지지 않게). 고치기가 실패해도 저장은 된다 (pptx는 알림 띠).
- 바깥 변경(AI·다른 프로그램)·원래대로·닫기(저장 / 저장 안 함 / 취소)·AI 보내기 전 확인·문서 안 링크: 한글 편집 창과 같다.
- 네트워크: 편집하는 동안 앱 밖으로 나가는 요청을 모두 막는다 (웹 요청·WebSocket·WebRTC·맞춤법 사전).
- 직접 편집하지 않는 것: 암호 걸린 문서("암호가 걸린 문서는 직접 편집할 수 없습니다"), 512MB 넘는 문서, 매크로 문서(docm·xlsm·pptm), 옛 형식(doc·xls·ppt·odt·ods·odp) → '프로그램에서 열기'.
- 한계:
  - Word·Excel·PowerPoint와 같지 않다. 저장하면 변환기(x2t)가 파일 꾸러미를 다시 쓴다 (`styles.xml`이 커짐, 한글·영문 사이 간격, 부분 순서).
    변경 추적·메모·차트·SmartArt·필드·포함 개체가 든 문서는 저장 뒤 확인하거나 '프로그램에서 열기'로.
  - 함께 편집·맞춤법 검사 없음. 편집기 화면은 원래 모양·로고 그대로이고 일부 안내는 영어다.
  - 상태 막대의 '모든 변경 사항이 저장되었습니다'는 편집기 안 기록 이야기 — 파일 저장 여부는 위 막대.
  - 편집 창마다 메모리 약 0.5GB.
- 라이선스: 편집기 묶음(`vendor/office-editor`)은 AGPL-3.0이다. 소스·고지·상표 규칙: [`LICENSES-OFFICE-EDITOR.md`](LICENSES-OFFICE-EDITOR.md).
  구조(세 출처, SDK 약속, 저장 흐름, 보안): [`ARCHITECTURE.md`](ARCHITECTURE.md) '문서·표·슬라이드 편집 창', 설계: [`OFFICE-EDITOR-DESIGN.md`](OFFICE-EDITOR-DESIGN.md).

---

## 명령줄

```
office-kit inspect <파일> [--offset n] [--limit n] [--slide n] [--engine auto|app|file] [--json]
office-kit edit <파일> <작업.json> [-o 출력] [--overwrite] [--engine auto|app|file] [--json]
office-kit help edit        # 형식별 작업 이름 목록 (help inspect도 같음)
```

- 파일 확장자로 형식을 골라 core 함수를 그대로 부른다 (MCP 도구와 같은 결과 글). 한글 작업은 MCP `hwp_edit`와 같은 스키마(`hwpOp`)로 먼저 확인한다.
- `작업.json`: 작업 배열, `{"operations": [...]}`, 작업 하나 `{"op": ...}` 모두 됨. JSON 문자열을 바로 써도 된다.
- `-o` 생략 = 원본에 저장. `-o`가 이미 있는 다른 파일이면 `--overwrite`가 있어야 덮어쓴다. `-o`의 폴더는 만들어 준다.
- 출력: 작업마다 `N. 결과`, `참고: …`(PowerPoint 쓰기 암호 안내 포함), 마지막에 `저장: 경로 (엔진[, N쪽])`.
  `--json`: inspect는 원자료, edit는 `{format, engine, output, results, notes, pages?}`.
- 살펴보기 글의 이어 보기 안내는 명령줄 옵션으로 바꿔 보여 준다 (`offset=30` → `--offset 30`, `slide=3` → `--slide 3`). 오류 글의 `engine:'app'`·`output`도 `--engine app`·`-o`로.
- 오류는 stderr에 `오류: …`, 종료 코드 1. `OFFICEKIT_DEBUG=1`이면 스택도.

| 형식 | inspect 옵션 | `--engine` | edit 저장 형식 |
|---|---|---|---|
| 한글 | `--offset/--limit` 문단 (기본 150, 최대 1000) | `app` 거절 | .hwp·.hwpx·.hwt |
| Word | `--offset/--limit` 문단 (기본 200, 최대 1000) | `file` 거절 | .docx·.doc·.rtf·.odt |
| PowerPoint | `--slide n` 또는 `--offset/--limit` 슬라이드 (`--limit` 없으면 끝까지, 있으면 MCP처럼 약 60,000자에서 끊음). `--slide`를 주면 `--offset/--limit`은 안 씀 | `file` 거절 | .pptx·.ppt·.odp |
| Excel | `--engine`만 (`--offset/--limit`은 "참고:"만 찍고 무시) | auto·app·file | Excel이면 xlsx·xlsm·xls·csv·ods, 파일 엔진이면 원본과 같은 형식 |

- 도움말(`help edit`)은 작업 이름만 보여 준다. 인수는 이 문서.

---

## 구현 위치

| 형식 | COM | JS (core) | MCP | 시험 |
|---|---|---|---|---|
| Word | `scripts/com/word-ops.ps1` (`Invoke-WordInspect/Edit`, 도우미 `Wd…`) | `src/core/word.mjs` (`wordInspect`, `formatWordInspect`, `wordEdit`, `normalizeWordOps`) | `src/mcp/tools/word.mjs` | `test/word.test.mjs` |
| PowerPoint | `scripts/com/ppt-ops.ps1` (`Invoke-PptInspect/Edit`, 도우미 `…-Ppt…`) | `src/core/ppt.mjs` (`pptInspect`, `formatPptInspect`, `pptEdit`, `pptOpSchema`, `PPT_OPS_HELP`) | `src/mcp/tools/ppt.mjs` | `test/ppt.test.mjs` |
| Excel | `scripts/com/excel-ops.ps1` (`Invoke-ExcelInspect/Edit`, 도우미 `Xl…`) | `src/core/excel.mjs` (`excelInspect`, `inspectText`, `excelEdit`, exceljs 파일 엔진 포함) | `src/mcp/tools/excel.mjs` | `test/excel.test.mjs` |
| 한글 | — | `src/core/hwp.mjs` (`Hwp#inspect`, `insertParagraph` 등) | `src/mcp/tools/hwp.mjs` (`hwpOp`, `applyHwpOp`, `inspectText`) | `test/hwp.test.mjs` |
| 명령줄 | — | `src/cli.mjs` (`inspect`, `edit`, `help edit`) | — | `test/cli-features.test.mjs` |

- COM 쪽 약속(점 소싱, 요청·결과 모양, 오류 코드, 원자적 편집, 시간 초과)은 [`ARCHITECTURE.md`](ARCHITECTURE.md).
- COM 시험은 `OFFICEKIT_COM_TESTS=1`일 때만 돈다 (한컴 COM은 쓰지 않음).
