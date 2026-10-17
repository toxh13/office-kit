// tools/excel.mjs — Excel 전용 MCP 도구: excel_inspect, excel_edit (docs/FEATURES.md 'Excel')
import { z } from 'zod';
import { extOf } from '../../core/index.mjs';
import { EXCEL_EXTS, EXCEL_OPS, excelEdit, excelInspect, inspectText } from '../../core/excel.mjs';

// zod 메시지를 한국어로: 스키마마다 한국어 오류 함수를 붙인다 (MCP SDK 입력 검증이 이 메시지를 쓴다)
const KO = z.locales.ko().localeError;
const E = { error: (iss) => KO(iss) };
const msg = (m) => ({ error: () => m });
// 모르는 인수는 거절한다 (조용히 버리면 AI가 원하지 않은 결과를 성공으로 받는다)
const strict = (shape, what) => z.strictObject(shape, {
  error: (iss) => (iss.code === 'unrecognized_keys'
    ? `모르는 인수입니다: ${iss.keys.join(', ')} — ${what}에 쓸 수 있는 인수: ${Object.keys(shape).filter((k) => k !== 'op').join(', ') || '(없음)'}`
    : KO(iss)),
});

const str = () => z.string(E);
const num = () => z.number(E);
const bool = () => z.boolean(E);
const int1 = () => z.number(E).int(E).min(1, E);
const value = z.union([z.string(E), z.number(E), z.boolean(E), z.null(E)], msg('값은 글자·숫자·true/false·null만 됩니다'));
const sheet = z.union([str().min(1, E), int1()], msg('시트는 이름(문자열) 또는 번호(1부터)여야 합니다'));
const sheetOpt = sheet.optional().describe('시트 이름 또는 번호(1부터). 생략하면 첫 시트');
const range = str().min(1, E).describe("범위 'A1:C10' (전체 열 'A:C', 전체 행 '2:5'도 됨)");
const color = str().describe("색 '#RRGGBB'");
const col = z.union([str(), z.number(E).int(E)], msg("열 문자('C') 또는 열 번호(3)여야 합니다"));

const op = (name, shape, what) => strict({ op: z.literal(name, E), ...shape }, name).describe(`${what}${EXCEL_OPS[name].file ? '' : ' (Excel 필요)'}`);

const excelOp = z.discriminatedUnion('op', [
  op('add_sheet', { name: str().min(1, E).describe('새 시트 이름 (31자까지, \\ / ? * [ ] : 불가)'), after: sheet.optional().describe('이 시트 뒤에 (생략하면 맨 뒤)') }, '시트 추가'),
  op('rename_sheet', { sheet, name: str().min(1, E) }, '시트 이름 바꾸기 (다른 시트 수식의 참조도 함께 바뀜)'),
  op('delete_sheet', { sheet }, '시트 지우기 (그 시트를 가리키던 수식은 #REF!)'),
  op('copy_sheet', { sheet, name: str().min(1, E).describe('복사본 이름') }, '시트 복사 (원본 바로 뒤에)'),
  op('move_sheet', { sheet, position: int1().describe('옮길 자리 (1부터)') }, '시트 순서 바꾸기'),
  op('set_cells', { sheet: sheetOpt, cells: z.record(str(), value, E).describe('{ "A1": "제목", "B2": 120, "C2": "=SUM(B2:B9)" } — 숫자는 따옴표 없이, \'=\'로 시작하면 수식, null은 지우기') }, '셀 값·수식 쓰기'),
  op('set_range', { sheet: sheetOpt, range: str().min(1, E).describe("붙여 넣을 왼쪽 위 셀 'A1'"), values: z.array(z.array(value, E), E).min(1, E).describe('[["이름","점수"],["가",90]] (값 규칙은 set_cells와 같음)') }, '2차원 값 붙여 넣기'),
  op('clear', { sheet: sheetOpt, range, what: z.enum(['all', 'contents', 'formats'], E).optional().describe("'contents'(기본: 값·수식만) | 'formats'(서식만) | 'all'") }, '지우기'),
  op('insert_rows', { sheet: sheetOpt, at: int1().describe('이 행 앞에 넣음'), count: int1().optional().describe('기본 1') }, '행 넣기'),
  op('delete_rows', { sheet: sheetOpt, at: int1().describe('지울 첫 행'), count: int1().optional().describe('기본 1') }, '행 지우기'),
  op('insert_columns', { sheet: sheetOpt, at: col.describe("이 열 앞에 넣음 ('C' 또는 3)"), count: int1().optional() }, '열 넣기'),
  op('delete_columns', { sheet: sheetOpt, at: col.describe("지울 첫 열 ('C' 또는 3)"), count: int1().optional() }, '열 지우기'),
  op('format', {
    sheet: sheetOpt, range,
    bold: bool().optional(), italic: bool().optional(), underline: bool().optional(),
    font: str().optional().describe('글꼴 이름'), size: num().optional().describe('글자 크기(pt)'),
    color: color.optional().describe("글자색 '#RRGGBB'"), fill: str().nullable().optional().describe("배경색 '#RRGGBB', 없애려면 'none'"),
    number_format: str().optional().describe("표시 형식 (예: '#,##0', '0.0%', 'yyyy-mm-dd', '@'(텍스트))"),
    align: z.enum(['general', 'left', 'center', 'right', 'justify', 'center_across', 'fill'], E).optional(),
    valign: z.enum(['top', 'center', 'bottom', 'justify'], E).optional(),
    wrap: bool().optional().describe('자동 줄 바꿈'),
    border: z.enum(['thin', 'medium', 'thick', 'none', 'outline'], E).optional().describe("모든 칸 테두리 굵기 | 'none' | 'outline'(바깥쪽만)"),
  }, '서식'),
  op('column_width', { sheet: sheetOpt, columns: str().describe("'A' 또는 'A:C'"), width: z.union([num(), z.literal('auto', E)], msg("width는 숫자(글자 수) 또는 'auto'여야 합니다")).describe("글자 수 기준 너비 또는 'auto'(내용에 맞춤, 파일 엔진은 근사)") }, '열 너비'),
  op('row_height', { sheet: sheetOpt, rows: z.union([str(), z.number(E).int(E)], msg("rows는 '5' 또는 '1:3'이어야 합니다")).describe("'5' 또는 '1:3'"), height: z.union([num(), z.literal('auto', E)], msg("height는 숫자(pt) 또는 'auto'여야 합니다")).describe("pt 또는 'auto'") }, '행 높이'),
  op('merge', { sheet: sheetOpt, range }, '병합 (왼쪽 위 값만 남음)'),
  op('unmerge', { sheet: sheetOpt, range }, '병합 풀기'),
  op('freeze', { sheet: sheetOpt, cell: str().nullable().describe("이 셀의 위·왼쪽을 고정 (예: 'B2' = 1행·A열 고정, 'A2' = 1행만). null이면 해제") }, '틀 고정'),
  op('sort', { sheet: sheetOpt, range, by: str().describe("기준 열 문자 (예: 'B')"), order: z.enum(['asc', 'desc'], E).optional(), header: bool().optional().describe('첫 행이 머리글이면 true (생략하면 Excel이 판단)') }, '정렬'),
  op('autofilter', { sheet: sheetOpt, range: str().nullable().describe("범위 (null이면 해제)") }, '자동 필터'),
  op('add_table', { sheet: sheetOpt, range, name: str().optional().describe('표 이름 (띄어쓰기 없이, 이미 있는 표·이름은 안 됨)'), style: str().optional().describe("표 스타일 (기본 'TableStyleMedium2', 예: 'TableStyleLight9')") }, '표(ListObject) 만들기 — 첫 행은 머리글'),
  op('add_chart', {
    sheet: sheetOpt, range: range.describe("원본 데이터 범위 (첫 행·첫 열은 이름표) 예: 'A1:C6'"),
    type: z.enum(['column', 'bar', 'line', 'pie', 'scatter', 'area'], E).describe('column 세로 막대 | bar 가로 막대 | line 꺾은선 | pie 원 | scatter 분산 | area 영역'),
    title: str().optional().describe('차트 제목 (생략하면 제목 없음 — 계열이 하나면 그 계열 이름)'), at: str().optional().describe('차트 왼쪽 위 셀 (생략하면 데이터 오른쪽)'),
    width: num().optional().describe('너비(pt, 기본 360)'), height: num().optional().describe('높이(pt, 기본 216)'),
    series_in: z.enum(['columns', 'rows'], E).optional().describe("계열 방향: 'columns'(열마다 계열 하나, 첫 열이 항목) | 'rows'(행마다). 생략하면 Excel처럼 행이 열보다 많으면 columns, 아니면 rows"),
  }, '차트 만들기'),
  op('conditional_format', {
    sheet: sheetOpt, range,
    rule: strict({
      type: z.enum(['greater_than', 'less_than', 'between', 'equal', 'text_contains', 'duplicate', 'color_scale', 'data_bar'], E),
      value: z.union([str(), num()], msg('value는 글자 또는 숫자여야 합니다')).optional().describe("비교 값 (between은 아래값, '=$B$1' 같은 수식도 됨)"),
      value2: z.union([str(), num()], msg('value2는 글자 또는 숫자여야 합니다')).optional().describe('between의 위값'),
      fill: color.optional().describe("강조 배경색 (data_bar는 막대 색). 생략하면 연한 빨강 '#FFC7CE'"),
      font_color: color.optional().describe("강조 글자색 (생략하면 '#9C0006')"),
    }, 'rule'),
  }, '조건부 서식'),
  op('data_validation', {
    sheet: sheetOpt, range,
    list: z.union([z.array(z.union([str(), num()], msg('목록 항목은 글자 또는 숫자여야 합니다')), E), str()], msg("list는 ['예','아니오'] 또는 '=$E$1:$E$5'여야 합니다")).optional()
      .describe("드롭다운 목록 ['예','아니오'] (항목에 쉼표·큰따옴표 불가) 또는 범위 '=$E$1:$E$5'"),
    type: z.enum(['whole', 'decimal'], E).optional().describe('list가 없을 때: whole 정수 | decimal 소수'),
    min: num().optional(), max: num().optional(),
  }, '입력 제한(데이터 유효성 검사)'),
  op('named_range', {
    name: str().min(1, E).describe('이름 (글자나 _로 시작, 띄어쓰기 없이)'),
    range: str().min(1, E).describe("'시트!A1:B2' — 띄어쓰기가 있는 시트 이름은 작은따옴표로 감쌈: 'My Sheet'!A1:B2"),
    sheet: sheet.optional().describe('range에 시트 이름이 없을 때 쓸 시트 (생략하면 첫 시트)'),
  }, '이름 정의 (이미 있으면 바꿈)'),
  op('recalculate', {}, '모든 수식 다시 계산'),
], {
  error: (iss) => (iss.code === 'invalid_union' && iss.discriminator === 'op'
    ? `알 수 없는 작업(op)입니다 — 가능: ${Object.keys(EXCEL_OPS).join(', ')}` : KO(iss)),
});

const OPS_HELP = [
  '작업 종류 (×는 Excel이 설치돼 있어야 함 — 없으면 파일 엔진(exceljs)으로 나머지를 처리):',
  '  {op:"add_sheet", name, after?} / {op:"rename_sheet", sheet, name} / {op:"delete_sheet", sheet} / {op:"copy_sheet", sheet, name} × / {op:"move_sheet", sheet, position}',
  '  {op:"set_cells", sheet?, cells:{"A1":"글", "B2":120, "C2":"=SUM(B2:B9)"}}  숫자는 따옴표 없이, \'=\'로 시작하면 수식(영어 함수 이름, 쉼표 구분), null은 지우기, 글자 그대로 \'=\'로 시작하려면 앞에 \'',
  '    글자는 늘 글자로 저장(\'0123\'·\'1/2\'·\'2026-10-18\'도 글자). 날짜는 "=DATE(2026,10,18)"이나 날짜 일련번호를 쓰고 format의 number_format:"yyyy-mm-dd"로, 백분율은 0.15 + number_format:"0%"',
  '  {op:"set_range", sheet?, range:"A1", values:[["이름","점수"],["가",90]]}  (병합된 칸은 왼쪽 위 칸에만 쓸 수 있음)',
  '  {op:"clear", sheet?, range, what?:"contents"(기본)|"formats"|"all"}',
  '  {op:"insert_rows"|"delete_rows", sheet?, at:5, count?} / {op:"insert_columns"|"delete_columns", sheet?, at:"C", count?}  (수식·병합·표·조건부 서식·이름 정의도 함께 옮겨짐)',
  '  {op:"format", sheet?, range, bold, italic, underline, font, size, color:"#RRGGBB", fill:"#RRGGBB"|"none", number_format:"#,##0", align:"left|center|right|justify", valign:"top|center|bottom", wrap, border:"thin|medium|thick|none|outline"}',
  '  {op:"column_width", sheet?, columns:"A:C", width:15|"auto"} / {op:"row_height", sheet?, rows:"1:3", height:24|"auto"}  (파일 엔진의 auto는 근사)',
  '  {op:"merge"|"unmerge", sheet?, range} / {op:"freeze", sheet?, cell:"B2"|null}',
  '  {op:"sort", sheet?, range, by:"B", order:"asc|desc", header?} ×',
  '  {op:"autofilter", sheet?, range|null} / {op:"add_table", sheet?, range, name?, style?}',
  '  {op:"add_chart", sheet?, range, type:"column|bar|line|pie|scatter|area", title?, at?:"E2", width?, height?, series_in?:"columns|rows"} ×  (첫 행·첫 열은 이름표, 계열 방향은 생략하면 Excel처럼 정함)',
  '  {op:"conditional_format", sheet?, range, rule:{type:"greater_than|less_than|between|equal|text_contains|duplicate|color_scale|data_bar", value?, value2?, fill?, font_color?}}',
  '  {op:"data_validation", sheet?, range, list:["예","아니오"]|"=$E$1:$E$5"} 또는 {..., type:"whole|decimal", min?, max?}',
  '  {op:"named_range", name, range:"시트!A1:B2"} / {op:"recalculate"} ×',
].join('\n');

/** server.mjs가 부른다. ctx: 공통 도우미 (tool, text, mustExist, outPath, 스키마 등) */
export function register(server, ctx) {
  const { tool, text, mustExist, outPath, pathSchema, outputSchema, overwriteSchema } = ctx;
  const ALLOW = new Set(EXCEL_EXTS);
  const excelPath = (path) => {
    const full = mustExist(path);
    if (!EXCEL_EXTS.includes(extOf(full))) throw new Error(`엑셀 문서(${EXCEL_EXTS.map((e) => `.${e}`).join(', ')})만 다룹니다: ${full}`);
    return full;
  };

  server.registerTool('excel_inspect', {
    title: '엑셀 구조 살펴보기',
    description: '엑셀 문서(xlsx, xlsm, xls, csv, ods)의 구조를 알려준다: 시트마다 이름·번호·숨김·사용 범위·행×열·틀 고정·표·차트·병합 수·'
      + '자동 필터·조건부 서식 수·입력 제한, 이름 정의, 활성 시트, 수식 예시(최대 20개). excel_edit 전에 시트 이름과 범위를 확인할 때 쓴다. '
      + '셀 값 자체는 office_read로 읽는다.',
    inputSchema: {
      path: pathSchema('엑셀 파일 경로'),
      engine: z.enum(['auto', 'app', 'file'], E).default('auto')
        .describe("'auto'(기본: xlsx·xlsm·csv는 Excel 없이 바로 읽고, xls·ods는 Excel로) | 'app'(설치된 Excel로) | 'file'(Excel 없이)"),
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, engine = 'auto' }) => {
    const info = await excelInspect(excelPath(path), { engine });
    return text(inspectText(info));
  }));

  server.registerTool('excel_edit', {
    title: '엑셀 편집',
    description: '엑셀 문서를 고친다: 시트 관리, 셀 값·수식, 행·열 넣기/지우기, 서식, 너비·높이, 병합, 틀 고정, 정렬, 자동 필터, 표, 차트, '
      + '조건부 서식, 입력 제한, 이름 정의, 다시 계산. operations를 순서대로 모두 적용한 뒤 한 번만 저장하고, 하나라도 실패하면 아무것도 저장하지 않는다. '
      + '앞 작업이 행·열·시트 순서를 바꾸면 뒤 작업은 바뀐 위치를 쓴다. 시트 이름·범위는 excel_inspect로 먼저 확인한다. '
      + '설치된 Excel이 있으면 Excel로(모든 작업, 원본 그대로), 없으면 파일 엔진(exceljs)으로 ×가 아닌 작업만 한다 '
      + '(파일 엔진은 xlsx·csv만, 차트·피벗 등 보존할 수 없는 내용이 있는 파일은 고치지 않음, csv는 값만). '
      + 'csv의 값 작업(수식 없이)은 Excel이 있어도 파일 엔진으로 한다 (인코딩·구분자를 지키고 \'0123\' 같은 값이 숫자로 바뀌지 않게).\n'
      + `${OPS_HELP}\n`
      + '모르는 인수는 거절한다. output을 생략하면 원본을 덮어쓴다. 결과는 office_page_image로 눈으로 확인할 수 있다.',
    inputSchema: {
      path: pathSchema('엑셀 파일 경로'),
      operations: z.array(excelOp, E).min(1, msg('operations에 작업을 하나 이상 주세요')).describe('순서대로 적용할 작업 목록'),
      output: outputSchema,
      overwrite: overwriteSchema,
      engine: z.enum(['auto', 'app', 'file'], E).default('auto')
        .describe("'auto'(기본: Excel이 있으면 Excel, 없으면 파일 엔진 — csv의 값 작업은 늘 파일 엔진) | 'app'(Excel 필수) | 'file'(Excel 없이, ×인 작업은 못 함)"),
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, operations, output, overwrite = false, engine = 'auto' }) => {
    const full = excelPath(path);
    const out = outPath(output, { source: full, allow: ALLOW, overwrite });
    const r = await excelEdit(full, operations, { output: out, engine });
    return text([
      ...r.results.map((line, i) => `${i + 1}. ${line}`),
      ...r.notes.map((n) => `참고: ${n}`),
      `저장: ${r.output} (엔진: ${r.engine === 'excel' ? 'Excel' : '파일(exceljs)'})`,
    ].join('\n'));
  }));
}
