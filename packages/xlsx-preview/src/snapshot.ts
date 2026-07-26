/**
 * A loaded document, flattened so it can cross a `postMessage`.
 *
 * `PreviewDocument` holds three things structured clone refuses: an ExcelJS
 * `Workbook`, the live engine `Workbook`, and `model.cell` — a function. That is
 * fine on the main thread and fatal off it, and it is the reason the Worker path
 * needs this module rather than a cast.
 *
 * What survives is deliberate. The values, the gaps, the audit findings and the
 * full appearance of every sheet all cross. The engine does not: reviving one
 * would mean re-parsing every formula on the main thread, which is the work we
 * moved off it. Anything that genuinely needs the engine — precedents, say —
 * asks the Worker, which still has it (see worker/index.ts).
 *
 * Errors are the one value that needs encoding. `ExcelError` is a class, and a
 * cloned class instance arrives as a plain object with the prototype gone, so
 * `isErr()` would quietly answer false and a `#REF!` would render as `[object
 * Object]`. They are written as `{ err: '#REF!' }` and rebuilt through the
 * engine's own `errorOf`, so the revived value is the *same* singleton the
 * engine would have produced.
 */

import { errorOf, type ErrorKind, type EvalStats, type Gap, type Scalar } from '@xlscalc/formula-engine';
import { isErr } from '@xlscalc/formula-engine';
import type { HardcodedFinding } from './audit.js';
import type { FileFacts, OverlayCell, RenderModel, RenderSource } from './bind.js';
import { layoutKey, type SheetLayout } from './layout.js';
import type { RawWorkbook } from './ooxml.js';
import type { PreviewDocument, SheetInfo } from './parse.js';

/** A `Scalar` with errors written as data. */
export type ScalarSnapshot = number | string | boolean | null | { err: ErrorKind };

export interface OverlayCellSnapshot {
  value: ScalarSnapshot;
  provenance: OverlayCell['provenance'];
  cached?: ScalarSnapshot;
  reason?: string;
  formula?: string;
}

export interface PreviewSnapshot {
  raw: RawWorkbook;
  layouts: SheetLayout[];
  sheets: SheetInfo[];
  sheetNames: string[];
  facts: FileFacts;
  report: {
    stats: EvalStats;
    gaps: Gap[];
    mismatches: Array<{ address: string; computed: ScalarSnapshot; cached: ScalarSnapshot }>;
  };
  hardcoded: Array<
    Omit<HardcodedFinding, 'stated' | 'expected'> & {
      stated: ScalarSnapshot;
      expected: ScalarSnapshot;
    }
  >;
  /** Per sheet: `layoutKey(row, col)` → the cell's overlay. */
  cells: Array<Map<number, OverlayCellSnapshot>>;
  parseMs: number;
  layoutMs: number;
  evalMs: number;
  stylesError?: string;
}

/** A document rebuilt from a snapshot: everything a render needs, minus the engine. */
export interface RestoredDocument extends RenderSource {
  raw: RawWorkbook;
  sheets: SheetInfo[];
  model: RenderModel;
  parseMs: number;
  layoutMs: number;
  evalMs: number;
  stylesError?: string;
}

export const encodeScalar = (v: Scalar): ScalarSnapshot => (isErr(v) ? { err: v.kind } : v);

export const decodeScalar = (v: ScalarSnapshot): Scalar =>
  v !== null && typeof v === 'object' ? (errorOf(v.err) ?? null) : v;

export function toSnapshot(doc: PreviewDocument): PreviewSnapshot {
  const cells = doc.raw.sheets.map((sheet, index) => {
    const out = new Map<number, OverlayCellSnapshot>();
    for (const cell of sheet.cells) {
      const overlay = doc.model.cell(index, cell.row, cell.col);
      if (overlay) out.set(layoutKey(cell.row, cell.col), encodeOverlay(overlay));
    }
    return out;
  });

  return {
    raw: doc.raw,
    layouts: doc.layouts,
    sheets: doc.sheets,
    sheetNames: doc.model.sheetNames,
    facts: doc.model.facts,
    report: {
      stats: doc.model.report.stats,
      gaps: doc.model.report.gaps,
      mismatches: doc.model.report.mismatches.map((m) => ({
        address: m.address,
        computed: encodeScalar(m.computed),
        cached: encodeScalar(m.cached),
      })),
    },
    hardcoded: doc.model.hardcoded.map((h) => ({
      ...h,
      stated: encodeScalar(h.stated),
      expected: encodeScalar(h.expected),
    })),
    cells,
    parseMs: doc.parseMs,
    layoutMs: doc.layoutMs,
    evalMs: doc.evalMs,
    ...(doc.stylesError === undefined ? {} : { stylesError: doc.stylesError }),
  };
}

export function fromSnapshot(snap: PreviewSnapshot): RestoredDocument {
  const model: RenderModel = {
    facts: snap.facts,
    sheetNames: snap.sheetNames,
    report: {
      stats: snap.report.stats,
      gaps: snap.report.gaps,
      mismatches: snap.report.mismatches.map((m) => ({
        address: m.address,
        computed: decodeScalar(m.computed),
        cached: decodeScalar(m.cached),
      })),
    },
    hardcoded: snap.hardcoded.map((h) => ({
      ...h,
      stated: decodeScalar(h.stated),
      expected: decodeScalar(h.expected),
    })),
    cell(sheet, row, col) {
      const found = snap.cells[sheet]?.get(layoutKey(row, col));
      return found ? decodeOverlay(found) : undefined;
    },
  };

  return {
    raw: snap.raw,
    layouts: snap.layouts,
    sheets: snap.sheets,
    model,
    parseMs: snap.parseMs,
    layoutMs: snap.layoutMs,
    evalMs: snap.evalMs,
    ...(snap.stylesError === undefined ? {} : { stylesError: snap.stylesError }),
  };
}

function encodeOverlay(c: OverlayCell): OverlayCellSnapshot {
  return {
    value: encodeScalar(c.value),
    provenance: c.provenance,
    ...(c.cached === undefined ? {} : { cached: encodeScalar(c.cached) }),
    ...(c.reason === undefined ? {} : { reason: c.reason }),
    ...(c.formula === undefined ? {} : { formula: c.formula }),
  };
}

function decodeOverlay(c: OverlayCellSnapshot): OverlayCell {
  return {
    value: decodeScalar(c.value),
    provenance: c.provenance,
    ...(c.cached === undefined ? {} : { cached: decodeScalar(c.cached) }),
    ...(c.reason === undefined ? {} : { reason: c.reason }),
    ...(c.formula === undefined ? {} : { formula: c.formula }),
  };
}
