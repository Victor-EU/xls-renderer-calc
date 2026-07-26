/**
 * The quickstart from the package README, kept here so it is typechecked.
 *
 * A README example that no longer compiles is worse than no example: it is the
 * first thing a new user runs and the last thing anyone re-reads. `npm run
 * typecheck` covers this file, so the snippet and the API cannot drift.
 */

import { useEffect, useState } from 'react';
import { loadXlsx, type PreviewDocument } from '@xlscalc/xlsx-preview';
import { ExcelView } from '@xlscalc/xlsx-preview/view';
import { createPreviewWorker } from '@xlscalc/xlsx-preview/worker';

// ── The whole integration ────────────────────────────────────────────────────

export function Preview({ file }: { file: File }) {
  const [doc, setDoc] = useState<PreviewDocument | null>(null);

  useEffect(() => {
    void file.arrayBuffer().then(loadXlsx).then(setDoc);
  }, [file]);

  if (!doc) return null;
  return <ExcelView doc={doc} sheet={0} />;
}

// ── The same thing, off the main thread ──────────────────────────────────────

const preview = createPreviewWorker();

export function OffThreadPreview({ file }: { file: File }) {
  const [doc, setDoc] = useState<Awaited<ReturnType<typeof preview.load>> | null>(null);

  useEffect(() => {
    void file.arrayBuffer().then((buf) => preview.load(buf)).then(setDoc);
  }, [file]);

  if (!doc) return null;
  return <ExcelView doc={doc} sheet={0} />;
}

// ── Reading what happened ────────────────────────────────────────────────────

export function summarise(doc: PreviewDocument): string {
  const { formulas, computed, cached, unsupported, circular } = doc.model.report.stats;
  return `${computed} computed, ${cached} from the file, ${unsupported + circular} refused, of ${formulas}`;
}
