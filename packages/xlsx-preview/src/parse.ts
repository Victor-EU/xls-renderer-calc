/**
 * Loading a workbook: two passes over the same bytes, each doing what it is best at.
 *
 *   ExcelJS  → styles, fonts, fills, borders, merges, column widths, panes.
 *   ooxml.ts → formulas, the raw `t` attribute, shared-formula translation.
 *
 * ExcelJS is kept because its style fidelity is proven (this is the renderer the
 * original spike validated cell by cell against a real financial model) and
 * because rewriting `styles.xml` handling would be a regression risk with no
 * upside. It is *not* trusted for formula values, because it drops the one
 * attribute the recalc story depends on.
 *
 * Both passes get their own copy of the buffer: ExcelJS detaches the ArrayBuffer
 * it is handed, and a detached buffer read later yields zero bytes rather than
 * an error — a failure that unit tests with small fixtures do not catch.
 *
 * The two passes are also *independent in failure*, which they were not until a
 * real corpus said otherwise. Three of ten real workbooks made ExcelJS throw —
 * comments filed under `xl/comments/` instead of `xl/comments1.xml`, a missing
 * content-type part, the sort of thing a generator emits and Excel forgives —
 * and the whole load failed. Our own reader parsed all three without complaint.
 * Losing the styling of a workbook we can compute perfectly, and showing the
 * user nothing at all, is the wrong trade in a viewer whose entire promise is
 * that it will not mislead you: a document with default fonts is honest, an
 * empty screen is not. So a styling failure is now recorded and rendered
 * around, and only a failure of the *formula* reader is fatal.
 */

import ExcelJS from 'exceljs';
import { format as numFormat } from 'numfmt';
import { setNumberFormatter } from '@xlscalc/formula-engine';
import { bind, type BindOptions, type PreviewModel, type RenderSource } from './bind.js';
import { blankLayout, layoutSheet, type SheetLayout } from './layout.js';
import { readXlsx, type RawWorkbook } from './ooxml.js';

// TEXT() asks the host for a number formatter rather than shipping its own; the
// renderer already carries numfmt, so wire it up once here.
setNumberFormatter((code, value) => numFormat(code, value as never));

export interface SheetInfo {
  name: string;
  /** Index into both `styled.worksheets` and the engine's sheet list. */
  index: number;
  rows: number;
  cols: number;
  cells: number;
  formulas: number;
  uncached: number;
}

export interface PreviewDocument extends RenderSource {
  /**
   * The ExcelJS parse. Kept for hosts that want something we do not model, but
   * nothing in this package reads it after `layouts` is built — that is the
   * point of layout.ts, and a renderer should take `layouts` instead.
   */
  styled: ExcelJS.Workbook;
  raw: RawWorkbook;
  model: PreviewModel;
  /** How each sheet looks, as plain data. Index matches `sheets`. */
  layouts: SheetLayout[];
  sheets: SheetInfo[];
  parseMs: number;
  /** Time spent reading appearance out of the styled parse. */
  layoutMs: number;
  evalMs: number;
  /**
   * Set when the styling pass failed and the document is being shown with
   * default formatting. The values are unaffected — they come from the other
   * pass — but the user is told, because a preview that silently drops the
   * author's formatting is making a claim it has not earned.
   */
  stylesError?: string;
}

export interface LoadOptions extends BindOptions {}

export async function loadXlsx(buf: ArrayBuffer, opts: LoadOptions = {}): Promise<PreviewDocument> {
  const t0 = now();

  const styled = new ExcelJS.Workbook();
  let stylesError: string | undefined;
  try {
    await styled.xlsx.load(buf.slice(0));
  } catch (e) {
    // Degrade to default styling rather than failing the load. See the note at
    // the top of this file: the numbers come from the other pass and are fine.
    stylesError = e instanceof Error ? e.message : String(e);
  }
  const raw = readXlsx(buf.slice(0));
  const parseMs = now() - t0;

  const t1 = now();
  const model = bind(raw, opts);
  const evalMs = now() - t1;

  // Appearance is read once, here, rather than on every render. The renderer
  // used to walk the ExcelJS worksheet itself, which meant one `getCell()` per
  // coordinate per React render; a wide sheet made that tens of thousands of
  // lookups per keystroke. Doing it at load also makes the result plain data,
  // which is what lets the whole pipeline move into a Worker.
  const t2 = now();
  const layouts: SheetLayout[] = raw.sheets.map((sheet, index) => {
    // Matched by name, with an index fallback, because the engine's sheet order
    // (workbook.xml) and ExcelJS's can disagree — and a silent off-by-one here
    // would render one sheet's numbers over another's formatting.
    const ws = styled.getWorksheet(sheet.name) ?? styled.worksheets[index];
    const minRows = maxOf(sheet.cells, (c) => c.row);
    const minCols = maxOf(sheet.cells, (c) => c.col);
    return ws
      ? layoutSheet(ws, { name: sheet.name, minRows, minCols })
      : blankLayout(sheet.name, minRows, minCols);
  });
  const layoutMs = now() - t2;

  const sheets: SheetInfo[] = raw.sheets.map((sheet, index) => {
    let formulas = 0;
    let uncached = 0;
    for (const cell of sheet.cells) {
      if (!cell.f) continue;
      formulas++;
      const rec = model.engine.record(index, cell.row, cell.col);
      if (rec?.cached === undefined) uncached++;
    }
    const layout = layouts[index]!;
    return {
      name: sheet.name,
      index,
      rows: layout.rows,
      cols: layout.cols,
      cells: sheet.cells.length,
      formulas,
      uncached,
    };
  });

  return {
    styled,
    raw,
    model,
    layouts,
    sheets,
    parseMs: Math.round(parseMs),
    layoutMs: Math.round(layoutMs),
    evalMs: Math.round(evalMs),
    ...(stylesError === undefined ? {} : { stylesError }),
  };
}

function maxOf<T>(items: T[], pick: (t: T) => number): number {
  let max = 0;
  for (const item of items) {
    const v = pick(item);
    if (v > max) max = v;
  }
  return max;
}

const now = (): number =>
  typeof performance !== 'undefined' ? performance.now() : Date.now();
