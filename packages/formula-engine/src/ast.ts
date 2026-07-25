import type { CellRef } from './a1.js';
import { formatCellRef, quoteSheetName } from './a1.js';
import { ERR, type ExcelError } from './values.js';

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';

export type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: ExcelError }
  /** `sheet` undefined means "the sheet the formula lives on". */
  | { k: 'ref'; sheet?: string; a: CellRef; b?: CellRef }
  | { k: 'name'; name: string; sheet?: string }
  | { k: 'fn'; name: string; args: Node[]; pos: number }
  | { k: 'bin'; op: BinOp; l: Node; r: Node }
  | { k: 'un'; op: '-' | '+'; v: Node }
  | { k: 'pct'; v: Node }
  | { k: 'array'; rows: Node[][] }
  /** Excel's space operator: the overlap of two ranges. */
  | { k: 'isect'; l: Node; r: Node }
  /** Excel's comma operator inside parens: `(A1:A2,C1:C2)`. */
  | { k: 'union'; parts: Node[] }
  /** Implicit-intersection marker written by modern Excel: `@A1:A9`. */
  | { k: 'at'; v: Node };

/** Depth-first walk, parents before children. */
export function walk(node: Node, visit: (n: Node) => void): void {
  visit(node);
  switch (node.k) {
    case 'fn':
      for (const a of node.args) walk(a, visit);
      return;
    case 'bin':
    case 'isect':
      walk(node.l, visit);
      walk(node.r, visit);
      return;
    case 'un':
    case 'pct':
    case 'at':
      walk(node.v, visit);
      return;
    case 'array':
      for (const row of node.rows) for (const c of row) walk(c, visit);
      return;
    case 'union':
      for (const p of node.parts) walk(p, visit);
      return;
    default:
      return;
  }
}

/**
 * Shift relative (non-`$`) coordinates by a row/column delta.
 *
 * This is what makes shared formulas correct: OOXML stores `<f t="shared">`
 * once on a master cell and every sibling reuses it by `si`, offset by its own
 * position. A renderer that hands back the master formula verbatim reports
 * every sibling as computing the master's numbers.
 *
 * Coordinates pushed out of the sheet become `#REF!`, exactly as Excel does.
 */
export function translate(node: Node, dRow: number, dCol: number, maxRow: number, maxCol: number): Node {
  const shiftCorner = (r: CellRef): CellRef | 'ref-error' => {
    const nextRow = r.absRow ? r.row : r.row + dRow;
    const nextCol = r.absCol ? r.col : r.col + dCol;
    if (!r.colOnly && (nextRow < 1 || nextRow > maxRow)) return 'ref-error';
    if (!r.rowOnly && (nextCol < 1 || nextCol > maxCol)) return 'ref-error';
    return { ...r, row: r.colOnly ? r.row : nextRow, col: r.rowOnly ? r.col : nextCol };
  };

  const rec = (n: Node): Node => {
    switch (n.k) {
      case 'ref': {
        const a = shiftCorner(n.a);
        const b = n.b ? shiftCorner(n.b) : undefined;
        if (a === 'ref-error' || b === 'ref-error') return { k: 'err', v: ERR.ref };
        return b ? { k: 'ref', sheet: n.sheet, a, b } : { k: 'ref', sheet: n.sheet, a };
      }
      case 'fn':
        return { ...n, args: n.args.map(rec) };
      case 'bin':
        return { ...n, l: rec(n.l), r: rec(n.r) };
      case 'isect':
        return { ...n, l: rec(n.l), r: rec(n.r) };
      case 'un':
      case 'pct':
      case 'at':
        return { ...n, v: rec(n.v) };
      case 'array':
        return { ...n, rows: n.rows.map((row) => row.map(rec)) };
      case 'union':
        return { ...n, parts: n.parts.map(rec) };
      default:
        return n;
    }
  };
  return rec(node);
}

/** Render an AST back to formula text — used in tooltips, tests and write-back. */
export function unparse(node: Node): string {
  switch (node.k) {
    case 'num':
      return String(node.v);
    case 'str':
      return `"${node.v.replace(/"/g, '""')}"`;
    case 'bool':
      return node.v ? 'TRUE' : 'FALSE';
    case 'err':
      return node.v.kind;
    case 'ref': {
      const prefix = node.sheet ? `${quoteSheetName(node.sheet)}!` : '';
      return prefix + formatCellRef(node.a) + (node.b ? `:${formatCellRef(node.b)}` : '');
    }
    case 'name':
      return (node.sheet ? `${quoteSheetName(node.sheet)}!` : '') + node.name;
    case 'fn':
      return `${node.name}(${node.args.map(unparse).join(',')})`;
    case 'bin':
      return `${unparse(node.l)}${node.op}${unparse(node.r)}`;
    case 'un':
      return `${node.op}${unparse(node.v)}`;
    case 'pct':
      return `${unparse(node.v)}%`;
    case 'array':
      return `{${node.rows.map((r) => r.map(unparse).join(',')).join(';')}}`;
    case 'isect':
      return `${unparse(node.l)} ${unparse(node.r)}`;
    case 'union':
      return `(${node.parts.map(unparse).join(',')})`;
    case 'at':
      return `@${unparse(node.v)}`;
  }
}
