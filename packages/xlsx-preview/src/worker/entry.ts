/**
 * The Worker body.
 *
 * Everything expensive happens here: unzipping, the XML read, the ExcelJS
 * styling parse, evaluation, and building the layout. The main thread receives
 * a snapshot — plain data — and does nothing but paint it.
 *
 * The engine stays *here* rather than being sent across, because sending it
 * would mean rebuilding it, and rebuilding it means re-parsing every formula on
 * the thread we just moved the work off. The one thing that genuinely needs it
 * afterwards — tracing precedents when a user clicks a cell — is a question
 * asked over the same channel instead.
 *
 * This module is a worker entry point: importing it for its exports does
 * nothing useful, and it installs a message handler on import.
 */

import { loadXlsx, type PreviewDocument } from '../parse.js';
import { toSnapshot } from '../snapshot.js';
import type { WorkerRequest, WorkerResponse } from './protocol.js';

declare const self: {
  postMessage(message: unknown): void;
  onmessage: ((e: { data: WorkerRequest }) => void) | null;
};

let doc: PreviewDocument | undefined;

self.onmessage = (e) => {
  const req = e.data;
  void handle(req).then(
    (res) => self.postMessage(res),
    (err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      const res: WorkerResponse = { id: req.id, kind: 'error', message };
      self.postMessage(res);
    },
  );
};

async function handle(req: WorkerRequest): Promise<WorkerResponse> {
  switch (req.kind) {
    case 'load': {
      doc = await loadXlsx(req.buf, req.opts);
      return { id: req.id, kind: 'load', snapshot: toSnapshot(doc) };
    }
    case 'precedents': {
      if (!doc) throw new Error('no workbook has been loaded in this worker');
      const cells = doc.model.engine.precedents(req.sheet, req.row, req.col, req.limit ?? 500);
      return { id: req.id, kind: 'precedents', cells };
    }
  }
}
