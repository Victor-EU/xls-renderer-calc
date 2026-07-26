/**
 * The visual shape of a sheet, as plain data.
 *
 * This module exists to draw a line. Everything about *how a workbook looks* —
 * column widths, merges, frozen panes, fonts, fills, borders, number-format
 * codes — is read out of ExcelJS here, once, at load time, and handed on as
 * ordinary objects, arrays and Maps. Nothing downstream of this file knows that
 * ExcelJS exists.
 *
 * Three things follow from that, and each of them was a real problem before:
 *
 *   1. A renderer no longer takes an `ExcelJS.Worksheet` as a prop, so a
 *      consumer no longer has to install ExcelJS, import its types, or work out
 *      that the worksheet must be looked up by *name* with an index fallback.
 *   2. The result is structured-cloneable, which is what lets the whole load run
 *      in a Worker. A live `ExcelJS.Workbook` cannot cross a postMessage; this
 *      can.
 *   3. The work happens once per load instead of once per render. The previous
 *      renderer called `worksheet.getCell()` for every coordinate on every
 *      React render, which for a wide sheet is tens of thousands of lookups per
 *      keystroke.
 *
 * Styles are deduplicated into a table and referenced by index, because that is
 * how spreadsheets are actually shaped: a financial model has tens of thousands
 * of cells and perhaps forty distinct formats. Storing a resolved style object
 * per cell would be most of the memory of the document for none of the
 * information.
 */

import type ExcelJS from 'exceljs';
import { MAX_COLS } from '@xlscalc/formula-engine';
import { resolveColor } from './theme.js';

/** A run of rich text, with its own formatting, already resolved to CSS values. */
export interface RichRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  size?: number;
  name?: string;
}

/**
 * Cell content that is presentation rather than value: rich text and
 * hyperlinks. These carry no number, so they never come from the engine.
 */
export type CellContentData =
  | { kind: 'rich'; runs: RichRun[] }
  /**
   * A hyperlink. `href` is the target exactly as the file wrote it; `unsafe`
   * says whether this library is willing to put it in an anchor — see
   * `isSafeHref`. When it is set, a renderer must show the text and not the
   * link.
   */
  | { kind: 'link'; text: string; href: string; unsafe?: true };

/** Schemes a hyperlink out of a stranger's workbook may navigate to. */
const SAFE_SCHEMES = new Set(['http', 'https', 'mailto']);

/**
 * Is this hyperlink target something we will hand to the browser?
 *
 * A `.xlsx` can name any target it likes, and `javascript:alert(document.cookie)`
 * is a legal one. Escaping the URL — which this library did, carefully, at every
 * insertion point — stops it breaking out of the attribute and does nothing at
 * all about the scheme: the anchor still runs script in the host's origin on
 * click. For a viewer whose entire pitch is "open an untrusted spreadsheet
 * without it leaving your machine", one click is one too many.
 *
 * So the check is an allowlist, not a denylist, and it is deliberately strict:
 *
 *   - Control characters are stripped before the scheme is read, because
 *     browsers strip them too — `java\nscript:` is `javascript:` by the time it
 *     is dispatched, and a naive `startsWith` misses that.
 *   - A target with no scheme at all is refused as well. In a workbook those are
 *     relative file paths (`../Q3 actuals.xlsx`), which mean nothing in a
 *     browser preview, and `//evil.example` is a protocol-relative jump to
 *     another origin wearing the same clothes.
 */
export function isSafeHref(target: string): boolean {
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(target.replace(/[\u0000-\u0020\u007f]/g, ''));
  return scheme !== null && SAFE_SCHEMES.has(scheme[1]!.toLowerCase());
}

export interface CellFont {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  size?: number;
  name?: string;
  /** CSS colour, already resolved through the theme and indexed palettes. */
  color?: string;
}

/** Excel's border weights, kept as names so a renderer decides how to draw them. */
export type BorderStyle =
  | 'hair'
  | 'thin'
  | 'dotted'
  | 'dashed'
  | 'medium'
  | 'thick'
  | 'double';

export interface CellBorder {
  style: BorderStyle;
  color?: string;
}

export interface CellAlign {
  horizontal?: 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
  wrap?: boolean;
}

/** One entry in a sheet's deduplicated style table. Index 0 is always empty. */
export interface CellStyle {
  /** The Excel number-format code, e.g. `#,##0.00;[Red](#,##0.00)`. */
  numFmt?: string;
  font?: CellFont;
  /** Solid pattern fills only, as a CSS colour. Gradients are not represented. */
  fill?: string;
  align?: CellAlign;
  top?: CellBorder;
  right?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
}

export interface MergeRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface SheetLayout {
  name: string;
  /** 1-based extents; a renderer walks `1..rows` and `1..cols`. */
  rows: number;
  cols: number;
  /** Pixel widths and heights, 1-based (index 0 is a placeholder). */
  colWidths: number[];
  rowHeights: number[];
  /** Frozen pane split, in whole rows/columns. Zero means no freeze. */
  frozenRows: number;
  frozenCols: number;
  merges: MergeRange[];
  /** Deduplicated styles; `styles[0]` is the empty style. */
  styles: CellStyle[];
  /** Cell → index into `styles`. Absent means `styles[0]`. */
  styleAt: Map<number, number>;
  /** Cell → rich text or hyperlink, for the few cells that have either. */
  content: Map<number, CellContentData>;
}

/**
 * Pack a row/column pair into one number. Same layout as the engine's `cellId`
 * with the sheet bits left at zero, because a layout is per-sheet — a Map keyed
 * this way is far smaller and faster than one keyed `"12:4"`, and it survives
 * structured clone, which an object with numeric keys does not do as cheaply.
 */
export const layoutKey = (row: number, col: number): number => row * MAX_COLS + col;

/**
 * Merged ranges, indexed for rendering: which cell starts a span, and which
 * cells are swallowed by one and must not be emitted at all.
 *
 * Shared rather than derived per renderer, because a `<td>` emitted for a
 * covered cell shifts every cell after it on that row — a whole sheet slides
 * sideways, and it looks like a data bug rather than a markup one.
 */
export interface MergeMap {
  span(row: number, col: number): { rows: number; cols: number } | undefined;
  covered(row: number, col: number): boolean;
}

export function mergeMap(layout: SheetLayout): MergeMap {
  const spans = new Map<number, { rows: number; cols: number }>();
  const hidden = new Set<number>();
  for (const { top, left, bottom, right } of layout.merges) {
    spans.set(layoutKey(top, left), { rows: bottom - top + 1, cols: right - left + 1 });
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        if (r !== top || c !== left) hidden.add(layoutKey(r, c));
      }
    }
  }
  return {
    span: (row, col) => spans.get(layoutKey(row, col)),
    covered: (row, col) => hidden.has(layoutKey(row, col)),
  };
}

/**
 * Cumulative pixel offsets, so a frozen row or column can be positioned with
 * `position: sticky`. Both arrays are 1-based and hold the offset *before* the
 * given row or column.
 *
 * `upto` bounds the work to what is actually being drawn. A renderer that has
 * clamped its extent has no use for the offset of row 900,000, and building it
 * anyway is a million-element array per render on exactly the sheets that can
 * least afford one.
 */
export function paneOffsets(
  layout: SheetLayout,
  upto?: { rows: number; cols: number },
): { left: number[]; top: number[] } {
  const nCols = Math.min(layout.cols, upto?.cols ?? layout.cols);
  const nRows = Math.min(layout.rows, upto?.rows ?? layout.rows);
  const left = [0, 0];
  for (let c = 1; c <= nCols; c++) left[c + 1] = left[c]! + (layout.colWidths[c] ?? 0);
  const top = [0, 0];
  for (let r = 1; r <= nRows; r++) top[r + 1] = top[r]! + (layout.rowHeights[r] ?? 0);
  return { left, top };
}

/**
 * How much of a sheet a renderer will actually draw.
 *
 * Both renderers paint every cell of the extent into the document — there is no
 * virtualisation, and adding it is a different piece of work — so the cost is
 * `rows × cols` whatever the sheet contains. That is fine for a financial model
 * and fatal for a stray: `ws.rowCount` is the last row with *anything* in it, so
 * one accidental keystroke in the bottom-right corner of a sheet declares a
 * 1,048,576 × 16,384 grid. Seventeen billion cells is not a slow render, it is
 * an out-of-memory crash — measured, at 160 bytes of HTML per cell — and a
 * viewer that dies on a file is worse than one that admits a limit.
 *
 * So the extent is bounded and the renderer *says* it was bounded. Truncating in
 * silence would be the same failure as a guessed number: the user sees a
 * complete-looking sheet that is not the sheet.
 */
export interface RenderLimits {
  maxRows: number;
  maxCols: number;
  /** Ceiling on `rows × cols`, which is what actually costs. */
  maxCells: number;
}

/**
 * Roughly 150,000 cells: a second of work and tens of megabytes of document, at
 * the outer edge of what a browser will hold without complaint. Every real
 * corpus sheet is far below it; a host that knows its files are bigger can pass
 * its own.
 */
export const DEFAULT_RENDER_LIMITS: RenderLimits = {
  maxRows: 20_000,
  maxCols: 512,
  maxCells: 150_000,
};

export interface RenderExtent {
  rows: number;
  cols: number;
  /** True when the sheet is larger than what is being drawn. */
  truncated: boolean;
}

export function renderExtent(layout: SheetLayout, limits?: Partial<RenderLimits>): RenderExtent {
  const lim = { ...DEFAULT_RENDER_LIMITS, ...limits };
  const cols = Math.max(1, Math.min(layout.cols, lim.maxCols));
  const rows = Math.max(1, Math.min(layout.rows, lim.maxRows, Math.floor(lim.maxCells / cols)));
  return { rows, cols, truncated: rows < layout.rows || cols < layout.cols };
}

/** Thousands separators without asking the platform what locale it is in. */
const grouped = (n: number): string => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

/** What the renderers put above a grid they had to cut short. Shared, so they agree. */
export function truncationNote(layout: SheetLayout, extent: RenderExtent): string | undefined {
  if (!extent.truncated) return undefined;
  const parts: string[] = [];
  if (extent.rows < layout.rows) parts.push(`${grouped(extent.rows)} of ${grouped(layout.rows)} rows`);
  if (extent.cols < layout.cols) parts.push(`${grouped(extent.cols)} of ${grouped(layout.cols)} columns`);
  return `Showing ${parts.join(' and ')} — the grid is drawn in full, so the rest is not rendered.`;
}

// Excel stores widths in characters and heights in points; browsers want pixels.
// The constants are the ones the original spike measured against real files.
const colPx = (w?: number, hidden?: boolean): number =>
  hidden ? 0 : Math.round((w ?? 8.43) * 7 + 5);
const rowPx = (h?: number, hidden?: boolean): number =>
  hidden ? 0 : Math.round((h ?? 15) * (96 / 72));

const BORDER_STYLES = new Set<string>([
  'hair',
  'thin',
  'dotted',
  'dashed',
  'medium',
  'thick',
  'double',
]);

export interface LayoutOptions {
  /** Sheet name to record, when it should come from the workbook rather than ExcelJS. */
  name?: string;
  /**
   * Floor for the grid extent. The styled parse and our own reader can disagree
   * about how big a sheet is — and where they do, the cells we can *compute*
   * must still be reachable, so the larger extent wins.
   */
  minRows?: number;
  minCols?: number;
}

/** Read one worksheet's appearance into plain data. */
export function layoutSheet(ws: ExcelJS.Worksheet, opts: LayoutOptions = {}): SheetLayout {
  const cols = Math.max(1, ws.columnCount, opts.minCols ?? 0);
  const rows = Math.max(1, ws.rowCount, opts.minRows ?? 0);
  const name = opts.name;

  const colWidths: number[] = [0];
  for (let c = 1; c <= cols; c++) {
    const col = ws.getColumn(c);
    colWidths.push(colPx(col.width, col.hidden));
  }
  // `findRow` rather than `getRow`: the extent runs to the last row with
  // *anything* in it, so a stray in the bottom corner makes this loop a million
  // iterations, and `getRow` materialises a Row object for every one of them.
  // `findRow` returns undefined for a row that does not exist, which is also the
  // right answer — it has Excel's default height.
  const rowHeights: number[] = [0];
  for (let r = 1; r <= rows; r++) {
    const row = ws.findRow(r);
    rowHeights.push(rowPx(row?.height, row?.hidden));
  }

  const frozen = (ws.views ?? []).find((v) => v.state === 'frozen') as
    | { xSplit?: number; ySplit?: number }
    | undefined;

  const merges: MergeRange[] = [];
  for (const range of (ws.model as unknown as { merges?: string[] }).merges ?? []) {
    const m = decodeRange(range);
    if (m) merges.push(m);
  }

  const styles: CellStyle[] = [{}];
  const byKey = new Map<string, number>();
  const styleAt = new Map<number, number>();
  const content = new Map<number, CellContentData>();

  // `includeEmpty` on both loops so a cell that carries only formatting — a
  // filled header band with no text — keeps its fill. Skipping those was never
  // an option: they are half of what makes a rendered sheet recognisable.
  ws.eachRow({ includeEmpty: true }, (row, r) => {
    row.eachCell({ includeEmpty: true }, (cell, c) => {
      const style = readStyle(cell);
      const id = intern(style, styles, byKey);
      if (id !== 0) styleAt.set(layoutKey(r, c), id);
      const rich = readContent(cell.value);
      if (rich) content.set(layoutKey(r, c), rich);
    });
  });

  return {
    name: name ?? ws.name,
    rows,
    cols,
    colWidths,
    rowHeights,
    frozenRows: frozen?.ySplit ?? 0,
    frozenCols: frozen?.xSplit ?? 0,
    merges,
    styles,
    styleAt,
    content,
  };
}

/**
 * A sheet with Excel's default sizing and no formatting at all.
 *
 * This is the degraded path, and it has to exist. Three of ten real workbooks
 * made ExcelJS throw — a comment part filed under the wrong name, a missing
 * content type, the sort of thing a generator emits and Excel forgives — while
 * our own reader parsed every one of them. The load already survives that and
 * records `stylesError`; without a layout to fall back on, the *renderer* would
 * still show an empty screen, which is the failure the whole arrangement exists
 * to avoid. A document in default fonts is honest. A blank page is not.
 */
export function blankLayout(name: string, rows: number, cols: number): SheetLayout {
  const width = colPx();
  const height = rowPx();
  return {
    name,
    rows: Math.max(1, rows),
    cols: Math.max(1, cols),
    colWidths: [0, ...Array<number>(Math.max(1, cols)).fill(width)],
    rowHeights: [0, ...Array<number>(Math.max(1, rows)).fill(height)],
    frozenRows: 0,
    frozenCols: 0,
    merges: [],
    styles: [{}],
    styleAt: new Map(),
    content: new Map(),
  };
}

function readStyle(cell: ExcelJS.Cell): CellStyle {
  const out: CellStyle = {};

  if (cell.numFmt) out.numFmt = cell.numFmt;

  const f = (cell.font ?? {}) as Record<string, unknown>;
  const font: CellFont = {};
  if (f.bold) font.bold = true;
  if (f.italic) font.italic = true;
  if (f.underline) font.underline = true;
  if (typeof f.size === 'number') font.size = f.size;
  if (typeof f.name === 'string') font.name = f.name;
  const color = resolveColor(f.color);
  if (color) font.color = color;
  if (Object.keys(font).length > 0) out.font = font;

  const fill = cell.fill as { type?: string; pattern?: string; fgColor?: unknown } | undefined;
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const bg = resolveColor(fill.fgColor);
    if (bg) out.fill = bg;
  }

  const a = cell.alignment ?? {};
  const align: CellAlign = {};
  if (a.horizontal) align.horizontal = a.horizontal;
  if (a.vertical) align.vertical = a.vertical;
  if (a.wrapText) align.wrap = true;
  if (Object.keys(align).length > 0) out.align = align;

  const b = cell.border ?? {};
  const top = readBorder(b.top);
  const right = readBorder(b.right);
  const bottom = readBorder(b.bottom);
  const left = readBorder(b.left);
  if (top) out.top = top;
  if (right) out.right = right;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;

  return out;
}

function readBorder(side: unknown): CellBorder | undefined {
  if (!side || typeof side !== 'object') return undefined;
  const s = side as { style?: string; color?: unknown };
  if (!s.style) return undefined;
  // An unrecognised weight still draws a line — Excel keeps adding names, and a
  // border we cannot classify is better rendered thin than dropped.
  const style = (BORDER_STYLES.has(s.style) ? s.style : 'thin') as BorderStyle;
  const color = resolveColor(s.color);
  return color ? { style, color } : { style };
}

function readContent(value: unknown): CellContentData | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const o = value as Record<string, unknown>;
  if ('richText' in o) {
    const runs = (o.richText as Array<{ text: string; font?: Record<string, unknown> }>).map((r) => {
      const f = (r.font ?? {}) as Record<string, unknown>;
      const run: RichRun = { text: r.text };
      if (f.bold) run.bold = true;
      if (f.italic) run.italic = true;
      if (typeof f.size === 'number') run.size = f.size;
      if (typeof f.name === 'string') run.name = f.name;
      const color = resolveColor(f.color);
      if (color) run.color = color;
      return run;
    });
    return { kind: 'rich', runs };
  }
  if ('hyperlink' in o) {
    // Judged here rather than in a renderer, because this is where the file's
    // bytes become our data: a `SheetLayout` is exported, structured-cloned into
    // a Worker and handed to third-party renderers, and none of them should have
    // to know that a spreadsheet can carry a `javascript:` URL.
    const href = String(o.hyperlink);
    const text = String(o.text ?? o.hyperlink);
    return isSafeHref(href) ? { kind: 'link', text, href } : { kind: 'link', text, href, unsafe: true };
  }
  return undefined;
}

/**
 * Return the index of an equal style, adding it if new. Keying on the JSON of a
 * style built with a fixed property order makes equality structural without a
 * deep-compare per cell.
 */
function intern(style: CellStyle, styles: CellStyle[], byKey: Map<string, number>): number {
  const key = JSON.stringify(style);
  if (key === '{}') return 0;
  const found = byKey.get(key);
  if (found !== undefined) return found;
  const id = styles.length;
  styles.push(style);
  byKey.set(key, id);
  return id;
}

function decodeRange(range: string): MergeRange | undefined {
  const [a, b] = range.split(':');
  const m1 = /^([A-Z]+)(\d+)$/.exec(a ?? '');
  const m2 = /^([A-Z]+)(\d+)$/.exec(b ?? a ?? '');
  if (!m1 || !m2) return undefined;
  const col = (s: string): number => s.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  return {
    top: Number(m1[2]),
    left: col(m1[1]!),
    bottom: Number(m2[2]),
    right: col(m2[1]!),
  };
}
