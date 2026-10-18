// SPDX-License-Identifier: AGPL-3.0-only
// jobs.mjs — 글꼴 캐시 폴더를 채우는 일 (worker.mjs에서, 또는 시험에서 바로 부른다). Office Kit 글꼴 생성기의 일부.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { GENERATOR_VERSION, generateAllFonts } from './allfonts.mjs';
import { renderThumbnails, writeAtomic, writeBlankSprites } from './thumbs.mjs';

/** 캐시 폴더 안 이름들 */
export const CACHE_FILES = {
  allFonts: 'AllFonts.js',
  manifest: 'fonts.json',
  thumbs: 'thumbs.json',
  images: 'Images',
  // Images/ 한 폴더: 빈 그림(바로) → plain(기본 글꼴, 약 1초) → full(글꼴마다, 약 6초)로 파일을 바꿔 넣는다
  stages: ['blank', 'plain', 'full'],
  stageDone: (stage) => `${stage}.done.json`,
};

/**
 * job 'allfonts': { outDir, listing:[{path,size,mtimeMs}], profile, key, editorVersion }
 * → { stats } — outDir에 AllFonts.js, fonts.json, thumbs.json, Images/*.bin(빈 그림)을 쓴다
 */
export function runAllFontsJob({ outDir, listing, profile = 'officekit', key = '', editorVersion = '' }) {
  const out = generateAllFonts(listing, { profile });
  mkdirSync(outDir, { recursive: true });
  writeFileSync(join(outDir, CACHE_FILES.allFonts), out.js);
  writeFileSync(join(outDir, CACHE_FILES.thumbs), JSON.stringify(out.thumbs));
  writeBlankSprites(join(outDir, CACHE_FILES.images), out.thumbs.length);
  writeFileSync(join(outDir, CACHE_FILES.stageDone('blank')), JSON.stringify({ at: Date.now() }));
  const manifest = {
    v: 1, generator: GENERATOR_VERSION, key, editorVersion, profile, createdAt: new Date().toISOString(),
    files: out.files, stats: out.stats,
  };
  // fonts.json을 마지막에 쓴다 — 이것이 있어야 완성된 캐시로 본다
  writeFileSync(join(outDir, CACHE_FILES.manifest), JSON.stringify(manifest));
  return { stats: out.stats };
}

/**
 * job 'thumbs': { dir, stages = ['plain', 'full'], yieldMs }
 * 이미 끝난 단계는 건너뛴다. 그림은 Images/에 바꿔 넣고, 끝나면 <stage>.done.json
 */
export async function runThumbsJob({ dir, stages = ['plain', 'full'], yieldMs = 2 }, { shouldStop = () => false, onStage = () => {} } = {}) {
  const thumbs = JSON.parse(readFileSync(join(dir, CACHE_FILES.thumbs), 'utf8'));
  const done = [];
  for (const stage of stages) {
    if (existsSync(join(dir, CACHE_FILES.stageDone(stage)))) { done.push(stage); continue; }
    const t0 = Date.now();
    const written = await renderThumbnails(thumbs, {
      outDir: join(dir, CACHE_FILES.images), mode: stage === 'full' ? 'full' : 'plain', shouldStop, yieldMs: stage === 'full' ? yieldMs : 0,
    });
    if (!written) return { stopped: true, done };
    writeAtomic(join(dir, CACHE_FILES.stageDone(stage)), JSON.stringify({ at: Date.now(), ms: Date.now() - t0, files: written.length }));
    done.push(stage);
    onStage(stage);
  }
  return { stopped: false, done };
}
