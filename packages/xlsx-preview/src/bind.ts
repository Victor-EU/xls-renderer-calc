/**
 * Binding the file to the engine, and the value overlay that comes back.
 *
 * The overlay is a side table, never a mutation of the parsed workbook. That is
 * what makes three things possible at once: every rendered number knows where it
 * came from, the file's own cached value survives next to the computed one so
 * the two can be compared (an audit finding, not a nicety), and the engine stays
 * swappable because nothing shares mutable state with it.
 */

import { ERR, errorOf, Workbook, type Scalar } from '@xlscalc/formula-engine';
import type { EvalReport, Provenance } from '@xlscalc/formula-engine';
import { findHardcoded, type HardcodedFinding } from './audit.js';
import type { SheetLayout } from './layout.js';
import { hasCachedValue, type RawCell, type RawWorkbook } from './ooxml.js';

export interface OverlayCell {
  value: Scalar;
  provenance: Provenance;
  /** The value the file carried, when it had one and it differs from ours. */
  cached?: Scalar;
  /** Why an `unsupported` cell could not be computed — shown verbatim in the ⚠ tooltip. */
  reason?: string;
  formula?: string;
}

export interface FileFacts {
  /** Formula cells the file never had a value for — the recalc trap, counted exactly. */
  uncached: number;
  formulas: number;
  arrayFormulas: number;
  sharedFormulas: number;
  date1904: boolean;
  fullCalcOnLoad: boolean;
  hasCalcChain: boolean;
}

/**
 * What a renderer actually needs from a model: the values, and the one fact
 * (`date1904`) that decides how a serial becomes a date.
 *
 * It is separate from `PreviewModel` because of the Worker. A live `Workbook`
 * cannot cross a `postMessage`, so the off-thread path returns a model that has
 * everything below and no `engine`. Typing the renderer against this instead of
 * against the full model is what lets one component serve both paths.
 */
export interface RenderModel {
  report: EvalReport;
  facts: FileFacts;
  sheetNames: string[];
  /** Literal cells whose neighbours' formulas disagree with them — see audit.ts. */
  hardcoded: HardcodedFinding[];
  cell(sheet: number, row: number, col: number): OverlayCell | undefined;
}

export interface PreviewModel extends RenderModel {
  engine: Workbook;
}

/**
 * The minimum a renderer takes: how each sheet looks, and what each cell holds.
 * Both `PreviewDocument` and the Worker's revived result satisfy it.
 */
export interface RenderSource {
  layouts: SheetLayout[];
  model: RenderModel;
}

export interface BindOptions {
  /** Frozen clock for NOW/TODAY, so a render is reproducible. */
  now?: number;
  /**
   * Show the file's own value for volatile cells when it has one. A preview is
   * meant to show what the generator produced; recomputing NOW() would make the
   * preview disagree with the file for a reason that is not a finding.
   */
  preferCachedVolatile?: boolean;
}

export function bind(raw: RawWorkbook, opts: BindOptions = {}): PreviewModel {
  const wb = new Workbook({
    date1904: raw.date1904,
    ...(opts.now === undefined ? {} : { now: opts.now }),
  });

  for (const sheet of raw.sheets) wb.addSheet(sheet.name);

  const facts: FileFacts = {
    uncached: 0,
    formulas: 0,
    arrayFormulas: 0,
    sharedFormulas: 0,
    date1904: raw.date1904,
    fullCalcOnLoad: raw.fullCalcOnLoad,
    hasCalcChain: raw.hasCalcChain,
  };

  raw.sheets.forEach((sheet, index) => {
    for (const cell of sheet.cells) {
      if (cell.f) {
        facts.formulas++;
        if (cell.arrayFormula) facts.arrayFormulas++;
        if (cell.fromShared) facts.sharedFormulas++;
        const cached = decodeCached(cell, raw.sharedStrings);
        if (cached === undefined) facts.uncached++;
        if (cached === undefined) wb.setFormula(index, cell.row, cell.col, cell.f);
        else wb.setFormula(index, cell.row, cell.col, cell.f, cached);
      } else {
        const v = decodeValue(cell, raw.sharedStrings);
        if (v !== undefined) wb.setValue(index, cell.row, cell.col, v);
      }
    }
  });

  for (const dn of raw.definedNames) {
    wb.defineName(dn.name, dn.formula, dn.sheet);
  }

  const report = wb.evaluateAll();

  // Collected after evaluation so the inferred formulas can actually be run.
  const literals: Array<{ sheet: number; row: number; col: number; value: Scalar }> = [];
  raw.sheets.forEach((sheet, index) => {
    for (const cell of sheet.cells) {
      if (cell.f) continue;
      const v = decodeValue(cell, raw.sharedStrings);
      if (typeof v === 'number') literals.push({ sheet: index, row: cell.row, col: cell.col, value: v });
    }
  });
  const hardcoded = findHardcoded(wb, literals);

  if (opts.preferCachedVolatile !== false) {
    for (const sheetIndex of raw.sheets.keys()) {
      const sheet = raw.sheets[sheetIndex]!;
      for (const cell of sheet.cells) {
        if (!cell.f) continue;
        const rec = wb.record(sheetIndex, cell.row, cell.col);
        if (rec?.provenance === 'volatile' && rec.cached !== undefined) rec.value = rec.cached;
      }
    }
  }

  return {
    engine: wb,
    report,
    facts,
    hardcoded,
    sheetNames: raw.sheets.map((s) => s.name),
    cell(sheet, row, col) {
      const rec = wb.record(sheet, row, col);
      if (!rec) return undefined;
      const out: OverlayCell = { value: rec.value, provenance: rec.provenance };
      if (rec.cached !== undefined) out.cached = rec.cached;
      if (rec.reason !== undefined) out.reason = rec.reason;
      const f = wb.formulaAt(sheet, row, col);
      if (f !== undefined) out.formula = f;
      return out;
    },
  };
}

/**
 * The cached value of a formula cell, or `undefined` when the file never had
 * one. This is the distinction the whole recalc story turns on.
 *
 * *Whether* there is one is `hasCachedValue`, which lives in ooxml.ts and is
 * shared with `inspectXlsx` so the two cannot answer differently. This decodes
 * it. `decodeValue` needs no special case for the empty-string result: its `str`
 * branch already defaults an absent `<v>` to `''`.
 */
function decodeCached(cell: RawCell, strings: string[]): Scalar | undefined {
  return hasCachedValue(cell) ? decodeValue(cell, strings) : undefined;
}

function decodeValue(cell: RawCell, strings: string[]): Scalar | undefined {
  switch (cell.t) {
    case 's': {
      const i = Number(cell.v);
      return Number.isFinite(i) ? (strings[i] ?? '') : '';
    }
    case 'inlineStr':
      return cell.is ?? '';
    case 'str':
      return cell.v ?? '';
    case 'b':
      return cell.v === '1' || cell.v === 'true';
    case 'e':
      return errorOf(cell.v ?? '') ?? ERR.value;
    case 'd': {
      // ISO-8601 cell values are rare but legal; keep them as text rather than
      // guessing a serial, which would be a silently wrong number.
      return cell.v ?? '';
    }
    default: {
      if (cell.v === undefined || cell.v === '') return undefined;
      const n = Number(cell.v);
      return Number.isFinite(n) ? n : cell.v;
    }
  }
}
