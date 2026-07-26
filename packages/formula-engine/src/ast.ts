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

/**
 * Binding powers, in Microsoft's documented order: reference operators,
 * negation, `%`, `^`, `*` `/`, `+` `-`, `&`, comparison.
 *
 * These live here rather than in the parser because `unparse` needs exactly the
 * same table to be the parser's inverse. When they were two tables, they drifted
 * — `unparse` emitted `bin` nodes with no parentheses at all, so a shared
 * formula `O19*(1-$E$17)` came back as `O19*1-$E$17` and computed a plausible
 * number two and a half times too large. Keeping one table makes that class of
 * bug unrepresentable.
 */
export const BP_CMP = 10;
export const BP_CONCAT = 20;
export const BP_ADD = 30;
export const BP_MUL = 40;
export const BP_POW = 50;
/** Postfix `%`. */
export const BP_PCT = 60;
/** Prefix `-`/`+` — above `^` and `%`, which is why `-2^2` is 4. */
export const BP_UNARY = 70;
/** The space operator, tightest of the value operators. */
export const BP_ISECT = 90;

export const INFIX_BP: Record<BinOp, number> = {
  '=': BP_CMP,
  '<>': BP_CMP,
  '<': BP_CMP,
  '>': BP_CMP,
  '<=': BP_CMP,
  '>=': BP_CMP,
  '&': BP_CONCAT,
  '+': BP_ADD,
  '-': BP_ADD,
  '*': BP_MUL,
  '/': BP_MUL,
  '^': BP_POW,
};

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

/**
 * Render an AST back to formula text — used in tooltips, in tests, and, load
 * bearingly, to reconstruct every sibling of a shared formula.
 *
 * The grammar has no parenthesis node: `(1-A1)` parses to the same subtree as
 * `1-A1`, and the brackets exist only in how it is attached to its parent. So
 * `unparse` has to *re-derive* them from precedence. It emits a bracket exactly
 * when a child binds more loosely than its position allows, which is both
 * correct and minimal — no defensive parentheses around things that do not need
 * them, because the output is read by people in tooltips.
 *
 * `parse(unparse(ast))` is asserted to deep-equal `ast` over every formula in
 * both corpora, which is the property that was silently false before.
 */
export function unparse(node: Node): string {
  return emit(node, 0);
}

function emit(node: Node, minBp: number): string {
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
      return `${node.name}(${node.args.map((a) => emit(a, 0)).join(',')})`;
    case 'array':
      return `{${node.rows.map((r) => r.map((c) => emit(c, 0)).join(',')).join(';')}}`;
    case 'union':
      // Already delimited by its own parentheses, which are part of the syntax.
      return `(${node.parts.map((p) => emit(p, 0)).join(',')})`;
    case 'bin': {
      const bp = INFIX_BP[node.op];
      // Left-associative, `^` included, so the right operand needs one more —
      // mirroring `parseExpr(bp + 1)` on the other side.
      return paren(bp, minBp, `${emit(node.l, bp)}${node.op}${emit(node.r, bp + 1)}`);
    }
    case 'un':
      return paren(BP_UNARY, minBp, `${node.op}${emit(node.v, BP_UNARY)}`);
    case 'pct':
      return paren(BP_PCT, minBp, `${emit(node.v, BP_PCT)}%`);
    case 'isect':
      return paren(BP_ISECT, minBp, `${emit(node.l, BP_ISECT)} ${emit(node.r, BP_ISECT + 1)}`);
    case 'at':
      return paren(BP_UNARY, minBp, `@${emit(node.v, BP_UNARY)}`);
  }
}

const paren = (bp: number, minBp: number, text: string): string =>
  bp < minBp ? `(${text})` : text;
