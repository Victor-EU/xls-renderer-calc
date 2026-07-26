/**
 * Shared argument handling.
 *
 * The rule almost every aggregate depends on, stated once here rather than
 * re-derived per function:
 *
 *   `SUM(A1:A3)` where A1 is the text "5"  → 0   (text inside a range is invisible)
 *   `SUM("5")`                             → 5   (a direct argument is coerced)
 *   `SUM("a")`                             → #VALUE!
 *   `SUM(A1:A3)` where A1 is TRUE          → 0   (booleans invisible in ranges)
 *   `SUM(TRUE)`                            → 1
 *
 * Getting this backwards produces subtly wrong subtotals rather than errors,
 * which is precisely the failure this project exists to avoid, so every aggregate
 * goes through `collectNumbers` instead of rolling its own loop.
 */

import { toNumber } from '../coerce.js';
import { deref, materialize, RefUnion, isMissingArg, type FnContext } from '../interpreter.js';
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
} from '../values.js';

export interface FlatItem {
  v: Scalar;
  /** True when the value was written directly as an argument, not read from a range. */
  direct: boolean;
}

/** Walk every scalar an argument list contributes, tagging its origin. */
export function* flatten(args: EvalValue[], ctx: FnContext): Generator<FlatItem> {
  for (const arg of args) {
    if (isMissingArg(arg)) {
      yield { v: null, direct: true };
      continue;
    }
    if (arg instanceof RefUnion) {
      for (const area of arg.areas) yield* fromMatrix(materialize(area, ctx));
      continue;
    }
    if (isRef(arg)) {
      yield* fromMatrix(materialize(arg, ctx));
      continue;
    }
    if (isMatrix(arg)) {
      yield* fromMatrix(arg);
      continue;
    }
    yield { v: arg, direct: true };
  }
}

function* fromMatrix(m: Matrix): Generator<FlatItem> {
  for (const v of m.values()) yield { v, direct: false };
}

/**
 * Numbers for an aggregate, following Excel's direct-vs-range rule.
 * Returns the first error encountered — errors are never skipped, they win.
 */
export function collectNumbers(args: EvalValue[], ctx: FnContext): number[] | ExcelError {
  const out: number[] = [];
  for (const arg of args) {
    // Fast path for the overwhelmingly common case: a range argument. Going
    // through `flatten` would materialise a Matrix and allocate one wrapper
    // object per cell, which is O(cells) garbage per call — and a model with a
    // running total (`SUM($B$1:B500)` on every row) makes that O(n²) overall.
    // Measured on a 45,000-formula workbook this path is the difference between
    // 14 seconds and well under a second.
    if (isRef(arg)) {
      const err = readRangeNumbers(arg, ctx, out);
      if (err) return err;
      continue;
    }
    if (arg instanceof RefUnion) {
      for (const area of arg.areas) {
        const err = readRangeNumbers(area, ctx, out);
        if (err) return err;
      }
      continue;
    }
    for (const { v, direct } of flatten([arg], ctx)) {
      if (isErr(v)) return v;
      if (direct) {
        if (v === null) continue; // an omitted argument contributes nothing
        const n = toNumber(v);
        if (isErr(n)) return n;
        out.push(n);
      } else if (typeof v === 'number') {
        out.push(v);
      }
    }
  }
  return out;
}

/** Append every number in a range to `out`; returns the first error found. */
function readRangeNumbers(ref: RefValue, ctx: FnContext, out: number[]): ExcelError | undefined {
  const { sheet, top, bottom, left, right } = ref;
  for (let r = top; r <= bottom; r++) {
    for (let c = left; c <= right; c++) {
      const v = ctx.host.cellValue(sheet, r, c);
      if (typeof v === 'number') out.push(v);
      else if (isErr(v)) return v;
    }
  }
  return undefined;
}

/**
 * Numbers for the "A" variants (AVERAGEA, COUNTA-adjacent, MAXA…), which do see
 * text as 0 and booleans as 1/0 inside ranges.
 */
export function collectNumbersA(args: EvalValue[], ctx: FnContext): number[] | ExcelError {
  const out: number[] = [];
  for (const { v } of flatten(args, ctx)) {
    if (isErr(v)) return v;
    if (v === null) continue;
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    else out.push(0); // text counts as zero
  }
  return out;
}

/** One scalar from an argument, with references dereferenced and arrays collapsed. */
export function arg1(args: EvalValue[], i: number, ctx: FnContext): Scalar {
  const a = args[i];
  if (a === undefined || isMissingArg(a)) return null;
  const d = deref(a, { ...ctx, arrayMode: false });
  if (isMatrix(d)) return d.size === 0 ? ERR.value : d.at(0, 0);
  return d;
}

/** A number from an argument, or the error that stopped it. */
export function num(args: EvalValue[], i: number, ctx: FnContext): number | ExcelError {
  return toNumber(arg1(args, i, ctx));
}

/** A number from an optional argument, with a default when omitted or blank. */
export function optNum(
  args: EvalValue[],
  i: number,
  ctx: FnContext,
  fallback: number,
): number | ExcelError {
  const a = args[i];
  if (a === undefined || isMissingArg(a)) return fallback;
  const s = arg1(args, i, ctx);
  if (s === null) return fallback;
  return toNumber(s);
}

/** Present an argument as a rectangle of values, whatever form it arrived in. */
export function asMatrix(v: EvalValue, ctx: FnContext): Matrix | ExcelError {
  if (isMissingArg(v)) return ERR.value;
  if (isRef(v)) return materialize(v, ctx);
  if (isMatrix(v)) return v;
  if (v instanceof RefUnion) return ERR.value;
  // An argument that *is* an error propagates; it does not become a
  // one-cell range containing an error. The difference decides what a
  // reference to a deleted sheet does: `COUNTA(Gone!A:A)` must be #REF!, and
  // wrapping the error in a matrix made it count the error as one non-empty
  // item and return 1 — a number, confidently, for a sheet that is not there.
  // A cell that merely *holds* an error is unaffected: it arrives through
  // materialize() as an element of the range, never as the bare argument.
  if (isErr(v)) return v;
  return Matrix.scalar(v);
}

/** The reference behind an argument, for functions with genuine ref semantics. */
export function asRef(v: EvalValue): RefValue | undefined {
  return isRef(v) ? v : undefined;
}

/**
 * Smallest and largest of a list, without spreading it into an argument list.
 *
 * `Math.min(...ns)` is the obvious spelling and it is a landmine: the spread
 * becomes one JS argument per element, and past roughly 125,000 arguments V8
 * throws `RangeError: Maximum call stack size exceeded`. `=MIN(A:A)` down a
 * column of a transaction export is an ordinary formula over an ordinary sheet,
 * and because a RangeError is not an `Unsupported` it escaped `evaluateAll` and
 * failed the whole workbook — the file rendered as nothing at all.
 */
export const minOf = (ns: number[]): number => {
  let out = ns[0]!;
  for (let i = 1; i < ns.length; i++) if (ns[i]! < out) out = ns[i]!;
  return out;
};
export const maxOf = (ns: number[]): number => {
  let out = ns[0]!;
  for (let i = 1; i < ns.length; i++) if (ns[i]! > out) out = ns[i]!;
  return out;
};

/** Excel truncates toward zero wherever an integer argument is required. */
export const trunc = (n: number): number => (n < 0 ? Math.ceil(n) : Math.floor(n));

/** Guard for the many functions that reject non-finite results as #NUM!. */
export const finiteOr = (n: number, e: ExcelError = ERR.num): number | ExcelError =>
  Number.isFinite(n) ? n : e;
