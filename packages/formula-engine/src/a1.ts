/**
 * A1 addressing. Everything public in this engine is **1-based**, exactly like
 * Excel: `A1` is `{row: 1, col: 1}`. That choice removes a whole class of
 * off-by-one bugs at the OOXML, ExcelJS and `ROW()`/`INDEX()` boundaries, which
 * are all 1-based too. The only 0-based indices in the codebase are inside
 * `Matrix`, and they never escape it.
 */

export const MAX_ROWS = 1_048_576; // 2^20
export const MAX_COLS = 16_384; // 2^14 — column XFD

/** Bit budget: sheet | row(20) | col(14). Max id ≈ 2^43, well inside 2^53. */
const COL_BITS = 14;
const ROW_BITS = 20;
const CELLS_PER_SHEET = 2 ** (ROW_BITS + COL_BITS);

/** Pack (sheet, row, col) into one number so hot maps are keyed by int. */
export function cellId(sheet: number, row: number, col: number): number {
  return sheet * CELLS_PER_SHEET + row * MAX_COLS + col;
}

export function unpackId(id: number): { sheet: number; row: number; col: number } {
  const sheet = Math.floor(id / CELLS_PER_SHEET);
  const rest = id - sheet * CELLS_PER_SHEET;
  const row = Math.floor(rest / MAX_COLS);
  return { sheet, row, col: rest - row * MAX_COLS };
}

/** 1 → "A", 26 → "Z", 27 → "AA", 16384 → "XFD". */
export function colName(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" → 1, "xfd" → 16384. Returns 0 for anything out of range. */
export function colNumber(name: string): number {
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const ch = name.charCodeAt(i) & ~32; // upper-case
    if (ch < 65 || ch > 90) return 0;
    n = n * 26 + (ch - 64);
    if (n > MAX_COLS) return 0;
  }
  return n;
}

/** One corner of a reference, carrying Excel's `$` anchors for translation. */
export interface CellRef {
  col: number;
  row: number;
  absCol: boolean;
  absRow: boolean;
  /** Whole-column (`A:A`) has no row anchor; whole-row (`1:1`) has no column. */
  colOnly?: boolean;
  rowOnly?: boolean;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;
const COL_ONLY_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_ONLY_RE = /^(\$?)([0-9]{1,7})$/;

/** Parse `$B$7` / `B7`. Returns undefined if it is not a legal A1 address. */
export function parseCellRef(text: string): CellRef | undefined {
  const m = CELL_RE.exec(text);
  if (!m) return undefined;
  const col = colNumber(m[2]!);
  const row = Number(m[4]!);
  if (col === 0 || row < 1 || row > MAX_ROWS) return undefined;
  return { col, row, absCol: m[1] === '$', absRow: m[3] === '$' };
}

/** Parse a bare column (`A`, `$XFD`) — only meaningful inside `A:A`. */
export function parseColOnly(text: string): CellRef | undefined {
  const m = COL_ONLY_RE.exec(text);
  if (!m) return undefined;
  const col = colNumber(m[2]!);
  if (col === 0) return undefined;
  return { col, row: 1, absCol: m[1] === '$', absRow: false, colOnly: true };
}

/** Parse a bare row (`7`, `$7`) — only meaningful inside `1:1`. */
export function parseRowOnly(text: string): CellRef | undefined {
  const m = ROW_ONLY_RE.exec(text);
  if (!m) return undefined;
  const row = Number(m[2]!);
  if (row < 1 || row > MAX_ROWS) return undefined;
  return { col: 1, row, absCol: false, absRow: m[1] === '$', rowOnly: true };
}

export function formatCellRef(r: CellRef): string {
  if (r.colOnly) return `${r.absCol ? '$' : ''}${colName(r.col)}`;
  if (r.rowOnly) return `${r.absRow ? '$' : ''}${r.row}`;
  return `${r.absCol ? '$' : ''}${colName(r.col)}${r.absRow ? '$' : ''}${r.row}`;
}

/** "Income Statement" → "'Income Statement'"; quoting is mandatory with spaces. */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseCellRef(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

export function unquoteSheetName(text: string): string {
  return text.startsWith("'") ? text.slice(1, -1).replace(/''/g, "'") : text;
}

/** "Sheet1!B7" for diagnostics and overlay keys. */
export function formatAddress(sheetName: string, row: number, col: number): string {
  return `${quoteSheetName(sheetName)}!${colName(col)}${row}`;
}
