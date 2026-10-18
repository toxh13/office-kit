// SPDX-License-Identifier: AGPL-3.0-only
// worker.mjs — 글꼴 일을 worker_thread에서 한다 (주 프로세스가 멈추지 않게). Office Kit 글꼴 생성기의 일부.
//   job 'allfonts': 글꼴 목록 → AllFonts.js·fonts.json·thumbs.json·빈 미리보기 (outDir)
//   job 'thumbs'  : thumbs.json → 미리보기 그림 plain(기본 글꼴) → full(글꼴마다)
import { parentPort, workerData } from 'node:worker_threads';
import { runAllFontsJob, runThumbsJob } from './jobs.mjs';

const post = (m) => parentPort?.postMessage(m);
let stop = false;
parentPort?.on('message', (m) => { if (m && m.type === 'stop') stop = true; });

(async () => {
  try {
    if (workerData.job === 'allfonts') post({ type: 'done', result: runAllFontsJob(workerData) });
    else if (workerData.job === 'thumbs') {
      const result = await runThumbsJob(workerData, { shouldStop: () => stop, onStage: (stage) => post({ type: 'stage', stage }) });
      post({ type: 'done', result });
    } else throw new Error(`모르는 일: ${workerData.job}`);
  } catch (err) {
    post({ type: 'error', message: err?.message ?? String(err), stack: err?.stack });
  }
})();
