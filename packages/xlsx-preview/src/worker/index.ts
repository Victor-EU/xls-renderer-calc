/**
 * Loading a workbook without blocking the page.
 *
 * The main-thread path is synchronous once it starts: reading the XML,
 * evaluating and building the layout all run to completion in one go. On the
 * real corpus's largest workbook — 138,421 formulas — that is long enough for
 * the tab to stop responding, which for a *preview* is the wrong trade at any
 * duration.
 *
 * Nothing about the pipeline needed to change to fix that. `loadXlsx` already
 * takes an `ArrayBuffer` and returns data; the only thing standing in the way
 * was that the returned document held an ExcelJS workbook, a live engine and a
 * method, none of which can cross a `postMessage`. Once `layout.ts` and
 * `snapshot.ts` made the result plain data, this file became a few dozen lines
 * of message plumbing.
 *
 *     import { createPreviewWorker } from '@xlscalc/xlsx-preview/worker';
 *
 *     const preview = createPreviewWorker();
 *     const doc = await preview.load(await file.arrayBuffer());
 *     // <ExcelView doc={doc} sheet={0} /> — same component, same props.
 *
 * The default construction uses `new URL('./entry.js', import.meta.url)`, which
 * Vite and webpack both recognise and bundle. Anything else — a different
 * bundler, a CSP that forbids blob workers, a same-origin worker you already
 * own — can pass its own `Worker` instead.
 */

import type { BindOptions, RenderModel } from '../bind.js';
import { fromSnapshot, type RestoredDocument } from '../snapshot.js';
import type { WorkerRequest, WorkerResponse } from './protocol.js';

export interface LoadInWorkerOptions extends BindOptions {
  /**
   * Hand the buffer to the worker rather than copying it. Faster for a large
   * file, and it leaves the caller's `ArrayBuffer` detached and unusable —
   * which is why it is off by default. Turn it on when the bytes came straight
   * from a `File` and nothing else is going to read them.
   */
  transfer?: boolean;
}

export interface PreviewWorker {
  /** Read, evaluate and lay out a workbook off the main thread. */
  load(buf: ArrayBuffer, opts?: LoadInWorkerOptions): Promise<RestoredDocument>;
  /**
   * Cells the given cell reads from, traced through the engine that is still
   * alive in the worker. Requires a prior `load` on the same worker.
   */
  precedents(
    sheet: number,
    row: number,
    col: number,
    limit?: number,
  ): Promise<Array<{ sheet: number; row: number; col: number }>>;
  terminate(): void;
}

/** Minimal structural type, so this file does not require DOM lib types. */
interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: { data: WorkerResponse }) => void): void;
  addEventListener(type: 'error', listener: (e: unknown) => void): void;
  terminate(): void;
}

export function createPreviewWorker(worker?: WorkerLike): PreviewWorker {
  const w = worker ?? defaultWorker();
  const pending = new Map<number, { resolve(v: WorkerResponse): void; reject(e: Error): void }>();
  let nextId = 1;

  w.addEventListener('message', (e) => {
    const res = e.data;
    const slot = pending.get(res.id);
    if (!slot) return;
    pending.delete(res.id);
    if (res.kind === 'error') slot.reject(new Error(res.message));
    else slot.resolve(res);
  });

  // A worker that dies takes every outstanding request with it. Without this,
  // a failed load is a promise that never settles and a spinner that never
  // stops — the most annoying possible way to report an error.
  w.addEventListener('error', (e) => {
    const message = errorText(e);
    for (const slot of pending.values()) slot.reject(new Error(message));
    pending.clear();
  });

  const send = (req: WorkerRequest, transfer?: Transferable[]): Promise<WorkerResponse> =>
    new Promise((resolve, reject) => {
      pending.set(req.id, { resolve, reject });
      w.postMessage(req, transfer);
    });

  return {
    async load(buf, opts = {}) {
      const { transfer, ...bind } = opts;
      const res = await send(
        { id: nextId++, kind: 'load', buf, opts: bind },
        transfer ? [buf] : undefined,
      );
      if (res.kind !== 'load') throw new Error(`unexpected reply: ${res.kind}`);
      return fromSnapshot(res.snapshot);
    },
    async precedents(sheet, row, col, limit) {
      const res = await send({
        id: nextId++,
        kind: 'precedents',
        sheet,
        row,
        col,
        ...(limit === undefined ? {} : { limit }),
      });
      if (res.kind !== 'precedents') throw new Error(`unexpected reply: ${res.kind}`);
      return res.cells;
    },
    terminate: () => w.terminate(),
  };
}

function defaultWorker(): WorkerLike {
  if (typeof Worker === 'undefined') {
    throw new Error('no Worker constructor here — pass one to createPreviewWorker()');
  }
  // Written as one literal expression on purpose. Vite and webpack both detect
  // `new Worker(new URL(…, import.meta.url))` *syntactically* and bundle the
  // entry as a separate chunk; hiding the constructor behind a variable defeats
  // that and produces a worker that 404s at runtime, in the build only.
  return new Worker(new URL('./entry.js', import.meta.url), {
    type: 'module',
  }) as unknown as WorkerLike;
}

function errorText(e: unknown): string {
  if (e && typeof e === 'object' && 'message' in e) return String((e as { message: unknown }).message);
  return 'the preview worker failed';
}

export type { RestoredDocument, RenderModel };
