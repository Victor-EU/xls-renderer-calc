/**
 * The evaluator.
 *
 * Two invariants hold everywhere below.
 *
 * **Errors are values.** `#DIV/0!` propagates through operators and into
 * functions that don't trap it, and lands in the cell as a computed result.
 * Refusals are different: they `throw Unsupported`, unwind the whole cell, and
 * land as a ⚠ with a reason. Nothing converts one into the other.
 *
 * **A reference is only dereferenced when someone asks for a value.** That is
 * what lets `SUM(A1:A3)` ignore text in the range while `SUM("5")` coerces it,
 * and what makes `ROW(B7)`, `COLUMNS(A:C)` and `OFFSET` expressible at all.
 */

import type { CellRef } from './a1.js';
import { MAX_COLS, MAX_ROWS, colName } from './a1.js';
import type { Node } from './ast.js';
import { toBool, toNumber, toText, compare } from './coerce.js';
import { Unsupported } from './errors.js';
import { getFunction } from './functions/registry.js';
import { isMissing } from './parser.js';
import {
  ERR,
  ExcelError,
  Matrix,
  RefValue,
  isErr,
  isMatrix,
  isRef,
  type EvalValue,
  type Scalar,
} from './values.js';

/** A comma-union of areas: `(A1:A2,C1:C2)`. Only aggregation functions accept it. */
export class RefUnion {
  constructor(readonly areas: RefValue[]) {}
}

/** What the interpreter needs from the workbook. Keeps this module storage-agnostic. */
export interface EvalHost {
  sheetIndex(name: string): number | undefined;
  sheetName(index: number): string;
  /**
   * Current value of a cell. Implementations must throw `Unsupported` when the
   * cell itself could not be computed — an unsupported precedent makes every
   * dependent unsupported too, which is the difference between an honest ⚠ and
   * a subtotal that is silently missing one of its inputs.
   */
  cellValue(sheet: number, row: number, col: number): Scalar;
  /** Bounds for whole-column/row references, so `SUM(A:A)` is not 1,048,576 reads. */
  usedRange(sheet: number): { rows: number; cols: number };
  definedName(name: string, sheet: number): RefValue | Scalar | undefined;
  /** Serial for NOW(), read once per evaluation run so results are reproducible. */
  now(): number;
  readonly date1904: boolean;
}

export interface FnContext {
  host: EvalHost;
  sheet: number;
  row: number;
  col: number;
  /**
   * In array mode a multi-cell reference derefs to a Matrix instead of doing
   * implicit intersection — the difference between `SUMPRODUCT(A1:A3*B1:B3)`
   * and `A1:A3*B1:B3` typed into a plain cell.
   */
  arrayMode: boolean;
  /** Set when the cell touched NOW/TODAY/RAND — recorded as `volatile` provenance. */
  volatile: boolean;
}

const MAX_RANGE_CELLS = 1_000_000;

export function evaluate(node: Node, ctx: FnContext): EvalValue {
  switch (node.k) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return node.v;

    case 'ref':
      return resolveRef(node, ctx);

    case 'name': {
      if (isMissing(node)) return null;
      const sheet = node.sheet ? ctx.host.sheetIndex(node.sheet) : ctx.sheet;
      if (sheet === undefined) return ERR.ref;
      const found = ctx.host.definedName(node.name, sheet);
      // A name that is genuinely absent from the workbook is Excel's own
      // `#NAME?` — we read every defined name, so absence is a fact, not a gap.
      return found === undefined ? ERR.name : found;
    }

    case 'array':
      return Matrix.of(
        node.rows.map((row) =>
          row.map((cell) => {
            const v = evaluate(cell, ctx);
            return isMatrix(v) || isRef(v) ? ERR.value : v;
          }),
        ),
      );

    case 'un': {
      const v = deref(evaluate(node.v, ctx), ctx);
      if (node.op === '+') return mapNumeric(v, (n) => n);
      return mapNumeric(v, (n) => -n);
    }

    case 'pct':
      return mapNumeric(deref(evaluate(node.v, ctx), ctx), (n) => n / 100);

    case 'bin':
      return binary(node.op, evaluate(node.l, ctx), evaluate(node.r, ctx), ctx);

    case 'at': {
      const v = evaluate(node.v, ctx);
      return isRef(v) ? implicitIntersect(v, ctx) : deref(v, { ...ctx, arrayMode: false });
    }

    case 'isect': {
      const l = evaluate(node.l, ctx);
      const r = evaluate(node.r, ctx);
      if (!isRef(l) || !isRef(r)) return ERR.value;
      if (l.sheet !== r.sheet) return ERR.value;
      const top = Math.max(l.top, r.top);
      const left = Math.max(l.left, r.left);
      const bottom = Math.min(l.bottom, r.bottom);
      const right = Math.min(l.right, r.right);
      if (top > bottom || left > right) return ERR.null; // Excel's #NULL! is exactly this
      return new RefValue(l.sheet, top, left, bottom, right);
    }

    case 'union': {
      const areas: RefValue[] = [];
      for (const part of node.parts) {
        const v = evaluate(part, ctx);
        if (isRef(v)) areas.push(v);
        else if (v instanceof RefUnion) areas.push(...v.areas);
        else return ERR.value;
      }
      return new RefUnion(areas) as unknown as EvalValue;
    }

    case 'fn':
      return callFunction(node, ctx);
  }
}

function callFunction(node: Node & { k: 'fn' }, ctx: FnContext): EvalValue {
  const def = getFunction(node.name);
  if (!def) {
    throw new Unsupported(
      'FN_UNKNOWN',
      `${node.name}() is not implemented by this engine`,
      node.name,
    );
  }
  if (node.args.length < def.minArgs || (def.maxArgs >= 0 && node.args.length > def.maxArgs)) {
    return ERR.value; // Excel reports a formula error for the wrong arity
  }
  if (def.volatile) ctx.volatile = true;

  if (def.lazy) return def.lazy(node.args, ctx, evaluate);

  const argCtx = def.arrayArgs ? { ...ctx, arrayMode: true } : ctx;
  const args = node.args.map((a) => (isMissing(a) ? MISSING : evaluate(a, argCtx)));
  return def.call(args, ctx);
}

/** Sentinel for an omitted argument — distinct from blank, which is a value. */
export const MISSING = Symbol('missing-arg') as unknown as EvalValue;
export const isMissingArg = (v: EvalValue): boolean => (v as unknown) === MISSING;

function resolveRef(node: Node & { k: 'ref' }, ctx: FnContext): EvalValue {
  const sheet = node.sheet === undefined ? ctx.sheet : ctx.host.sheetIndex(node.sheet);
  if (sheet === undefined) return ERR.ref;

  const used = ctx.host.usedRange(sheet);
  const a = node.a;
  const b = node.b ?? node.a;

  // Whole-column and whole-row references are clamped to the used range. Excel
  // effectively does the same; expanding them literally is 1M cells per edge.
  const top = a.colOnly ? 1 : Math.min(a.row, b.row);
  const bottom = a.colOnly ? Math.max(1, used.rows) : Math.max(a.row, b.row);
  const left = a.rowOnly ? 1 : Math.min(a.col, b.col);
  const right = a.rowOnly ? Math.max(1, used.cols) : Math.max(a.col, b.col);

  if (top > MAX_ROWS || left > MAX_COLS) return ERR.ref;
  const rect = new RefValue(sheet, top, left, Math.min(bottom, MAX_ROWS), Math.min(right, MAX_COLS));
  if (rect.cellCount > MAX_RANGE_CELLS) {
    throw new Unsupported('LIMIT', `range ${describeRect(rect, ctx)} covers too many cells`);
  }
  return rect;
}

export function describeRect(r: RefValue, ctx: FnContext): string {
  const name = ctx.host.sheetName(r.sheet);
  const a = `${colName(r.left)}${r.top}`;
  const b = `${colName(r.right)}${r.bottom}`;
  return `${name}!${a}${a === b ? '' : `:${b}`}`;
}

/** Read every cell of a reference into a Matrix. */
export function materialize(ref: RefValue, ctx: FnContext): Matrix {
  const data: Scalar[] = new Array(ref.cellCount);
  let i = 0;
  for (let r = ref.top; r <= ref.bottom; r++) {
    for (let c = ref.left; c <= ref.right; c++) data[i++] = ctx.host.cellValue(ref.sheet, r, c);
  }
  return new Matrix(ref.rows, ref.cols, data);
}

/**
 * Turn a value into something a scalar-hungry operator can use.
 * A single cell yields its value; a range yields a Matrix in array mode and an
 * implicit intersection otherwise (legacy Excel semantics — this engine renders
 * saved files, so it never spills).
 */
export function deref(v: EvalValue, ctx: FnContext): Scalar | Matrix {
  if (isRef(v)) {
    if (v.isSingleCell) return ctx.host.cellValue(v.sheet, v.top, v.left);
    return ctx.arrayMode ? materialize(v, ctx) : implicitIntersect(v, ctx);
  }
  if (v instanceof RefUnion) return ERR.value;
  return v;
}

/** Collapse a value all the way to one scalar, for functions that need exactly one. */
export function scalarOf(v: EvalValue, ctx: FnContext): Scalar {
  const d = deref(v, { ...ctx, arrayMode: false });
  return isMatrix(d) ? (d.size === 0 ? ERR.value : d.at(0, 0)) : d;
}

function implicitIntersect(ref: RefValue, ctx: FnContext): Scalar {
  if (ref.isSingleCell) return ctx.host.cellValue(ref.sheet, ref.top, ref.left);
  // A single row spanning our column, or a single column spanning our row.
  if (ref.rows === 1 && ctx.col >= ref.left && ctx.col <= ref.right) {
    return ctx.host.cellValue(ref.sheet, ref.top, ctx.col);
  }
  if (ref.cols === 1 && ctx.row >= ref.top && ctx.row <= ref.bottom) {
    return ctx.host.cellValue(ref.sheet, ctx.row, ref.left);
  }
  return ERR.value;
}

function mapNumeric(v: Scalar | Matrix, f: (n: number) => number): EvalValue {
  if (isMatrix(v)) {
    return new Matrix(
      v.rows,
      v.cols,
      v.data.map((x) => {
        const n = toNumber(x);
        return isErr(n) ? n : f(n);
      }),
    );
  }
  const n = toNumber(v);
  return isErr(n) ? n : f(n);
}

const COMPARISONS: Record<string, (c: number) => boolean> = {
  '=': (c) => c === 0,
  '<>': (c) => c !== 0,
  '<': (c) => c < 0,
  '>': (c) => c > 0,
  '<=': (c) => c <= 0,
  '>=': (c) => c >= 0,
};

function binary(op: string, lRaw: EvalValue, rRaw: EvalValue, ctx: FnContext): EvalValue {
  const l = deref(lRaw, ctx);
  const r = deref(rRaw, ctx);

  if (isMatrix(l) || isMatrix(r)) return broadcast(op, l, r, ctx);
  return scalarBinary(op, l, r);
}

function scalarBinary(op: string, l: Scalar, r: Scalar): Scalar {
  if (isErr(l)) return l;
  if (isErr(r)) return r;

  if (op === '&') {
    const a = toText(l);
    if (isErr(a)) return a;
    const b = toText(r);
    return isErr(b) ? b : a + b;
  }

  const cmp = COMPARISONS[op];
  if (cmp) {
    const c = compare(l, r);
    return isErr(c) ? c : cmp(c);
  }

  const a = toNumber(l);
  if (isErr(a)) return a;
  const b = toNumber(r);
  if (isErr(b)) return b;

  switch (op) {
    case '+':
      return a + b;
    case '-':
      return a - b;
    case '*':
      return a * b;
    case '/':
      return b === 0 ? ERR.div0 : a / b;
    case '^': {
      const p = Math.pow(a, b);
      // Excel refuses results that are not real numbers, e.g. (-8)^(1/3).
      return Number.isFinite(p) ? p : ERR.num;
    }
    default:
      return ERR.value;
  }
}

function broadcast(op: string, l: Scalar | Matrix, r: Scalar | Matrix, ctx: FnContext): EvalValue {
  const lm = isMatrix(l) ? l : Matrix.scalar(l);
  const rm = isMatrix(r) ? r : Matrix.scalar(r);
  const rows = Math.max(lm.rows, rm.rows);
  const cols = Math.max(lm.cols, rm.cols);
  if (rows * cols > MAX_RANGE_CELLS) {
    throw new Unsupported('LIMIT', 'array operation covers too many cells');
  }
  const out: Scalar[] = new Array(rows * cols);
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const a = pick(lm, i, j);
      const b = pick(rm, i, j);
      out[i * cols + j] = scalarBinary(op, a, b);
    }
  }
  void ctx;
  return new Matrix(rows, cols, out);
}

/** Excel broadcasts a single row down and a single column across; else #N/A. */
function pick(m: Matrix, i: number, j: number): Scalar {
  const r = m.rows === 1 ? 0 : i;
  const c = m.cols === 1 ? 0 : j;
  if (r >= m.rows || c >= m.cols) return ERR.na;
  return m.at(r, c);
}

export { toBool, toNumber, toText, compare, isErr, ExcelError, ERR, Matrix, RefValue };
export type { CellRef, Scalar, EvalValue };
