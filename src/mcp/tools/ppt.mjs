// tools/ppt.mjs — PowerPoint 전용 MCP 도구 (docs/FEATURES.md 'PowerPoint'). 설치된 PowerPoint를 COM으로 조작한다.
import { z } from 'zod';
import { extOf } from '../../core/index.mjs';
import { pptInspect, pptEdit, formatPptInspect, pptOpSchema, PPT_OPS_HELP, PPT_EXTS, INSPECT_LIMIT, zodKo as E } from '../../core/ppt.mjs';

const PPT_SAVE = new Set(PPT_EXTS);

/** server.mjs가 부른다. ctx: 공통 도우미 (tool, text, mustExist, outPath, 스키마 등) */
export function register(server, ctx) {
  const { tool, text, abs, mustExist, outPath, pathSchema, outputSchema, overwriteSchema } = ctx;

  const mustBePpt = (path, name) => {
    const full = mustExist(path);
    if (!PPT_SAVE.has(extOf(full))) throw new Error(`${name}는 PowerPoint 문서(.pptx/.ppt/.odp)만 다룹니다`);
    return full;
  };

  server.registerTool('ppt_inspect', {
    title: 'PowerPoint 구조 살펴보기',
    description: 'PowerPoint 문서(pptx/ppt/odp)의 구조를 본다: 슬라이드 크기, 슬라이드마다 번호·레이아웃·제목·숨김·화면 전환·발표자 노트, '
      + '도형 목록(번호·이름·종류·개체 틀 종류·위치·크기(pt)·글 200자), 쓸 수 있는 레이아웃 이름(add_slide의 layout 키와 짝). '
      + `ppt_edit 전에 슬라이드·도형 번호와 이름을 확인할 때 쓴다. 한 번에 슬라이드 ${INSPECT_LIMIT}장까지(limit) 보여 주고, 더 있으면 이어 볼 offset을 알려 준다. `
      + '슬라이드마다 도형은 60개까지 (slide를 주면 그 슬라이드만 자세히, 도형 500개까지). 암호가 걸린 문서는 열 수 없다. PowerPoint가 설치돼 있어야 한다.',
    inputSchema: {
      path: pathSchema('PowerPoint 문서 경로 (.pptx/.ppt/.odp)'),
      slide: z.number(E).int(E).min(1, E).optional().describe('이 슬라이드만 자세히 (1부터, 생략하면 offset부터 limit장)'),
      offset: z.number(E).int(E).min(0, E).optional().describe('앞에서 건너뛸 슬라이드 수 (기본 0 — 31번째부터 보려면 30)'),
      limit: z.number(E).int(E).min(1, E).max(200, E).optional().describe(`보여 줄 슬라이드 수 (기본 ${INSPECT_LIMIT}, 최대 200)`),
    },
    annotations: { readOnlyHint: true },
  }, tool(async ({ path, slide, offset = 0, limit = INSPECT_LIMIT }) => {
    const full = mustBePpt(path, 'ppt_inspect');
    return text(formatPptInspect(await pptInspect(full, { slide, offset, limit })));
  }));

  server.registerTool('ppt_edit', {
    title: 'PowerPoint 편집',
    description: 'PowerPoint 문서(pptx/ppt/odp)를 설치된 PowerPoint로 편집한다. operations를 순서대로 모두 적용한 뒤 한 번만 저장하고, '
      + '하나라도 실패하면 아무것도 저장하지 않는다(몇 번째 작업이 왜 실패했는지 알려 준다). 앞 작업이 슬라이드를 넣거나 지우면 '
      + '뒤 작업은 바뀐 번호를 쓴다(결과 줄에 새 번호가 나온다). 먼저 ppt_inspect로 번호·도형 이름·레이아웃을 확인하고, '
      + '끝나면 office_page_image로 모양을 확인한다 (쪽 번호 = 슬라이드 번호, 숨긴 슬라이드도 포함).\n'
      + `${PPT_OPS_HELP}\n`
      + 'add_chart를 하는 동안 PowerPoint 창이 잠깐 최소화돼 보일 수 있다. 암호가 걸린 문서는 열 수 없고, 쓰기 암호만 걸린 pptx는 output으로 다른 파일에만 저장한다.\n'
      + 'output을 생략하면 원본을 덮어쓴다 (.pptx/.ppt/.odp로 저장).',
    inputSchema: {
      path: pathSchema('PowerPoint 문서 경로 (.pptx/.ppt/.odp)'),
      operations: z.array(pptOpSchema, E).min(1, E).describe('순서대로 적용할 작업 목록'),
      output: outputSchema,
      overwrite: overwriteSchema,
    },
    annotations: { destructiveHint: true },
  }, tool(async ({ path, operations, output, overwrite = false }) => {
    const full = mustBePpt(path, 'ppt_edit');
    if (output && !PPT_SAVE.has(extOf(abs(output)))) throw new Error('저장 경로는 .pptx, .ppt 또는 .odp 여야 합니다');
    const out = outPath(output, { source: full, allow: PPT_SAVE, overwrite });
    // 그림·테마 경로도 작업 폴더 기준
    const ops = operations.map((op) => ({
      ...op,
      ...(op.image_path ? { image_path: abs(op.image_path) } : {}),
      ...(op.theme_path ? { theme_path: abs(op.theme_path) } : {}),
    }));
    const r = await pptEdit(full, ops, { output: out });
    return text([...r.results, ...(r.notes ?? []), `저장: ${r.output}`].join('\n'));
  }));
}
