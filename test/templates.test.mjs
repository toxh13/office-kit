// 새 문서 서식 파일 (src/core/templates.mjs): 모든 서식이 모든 형식으로 만들어지고 다시 읽힌다
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { TEMPLATES, templatesFor, templateMarkdown } from '../src/core/templates.mjs';
import { createDocument } from '../src/core/create.mjs';
import { readDocument } from '../src/core/index.mjs';

test('서식 목록: 형식마다 하나 이상, 이름·설명은 한국어', () => {
  for (const f of ['hwpx', 'hwp', 'docx', 'pptx', 'xlsx']) assert.ok(templatesFor(f).length > 0, f);
  for (const t of TEMPLATES) {
    assert.match(t.name, /[가-힣]/);
    assert.match(t.desc, /[가-힣]/);
  }
  assert.equal(templateMarkdown('minutes', 'pptx'), null, '그 형식이 아닌 서식은 없음');
  assert.equal(templateMarkdown('없는 서식', 'docx'), null);
});

test('모든 서식 × 형식으로 만들고 다시 읽기', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'officekit-tpl-'));
  try {
    for (const t of TEMPLATES) {
      for (const f of t.formats) {
        const out = join(dir, `${t.id}.${f}`);
        await createDocument(out, templateMarkdown(t.id, f));
        const d = await readDocument(out, { engine: 'file' });
        assert.ok(String(d.text ?? '').trim().length > 0, `${t.id}.${f} 글이 비었음`);
      }
    }
  } finally { rmSync(dir, { recursive: true, force: true, maxRetries: 5 }); }
});
