/**
 * Lookups.
 *
 * Approximate match (`VLOOKUP`'s omitted 4th argument, `MATCH` type 1) is the
 * most misimplemented behaviour in Excel: it returns the *largest value ≤ the
 * key* and is only defined on sorted data, because Excel implements it as a
 * binary search. On unsorted data Excel returns whatever its probe sequence
 * happens to land on — a number, but a meaningless one.
 *
 * So this engine implements the sorted contract exactly and *refuses* when the
 * data is not sorted, rather than reproducing an arbitrary probe order or, worse,
 * silently scanning linearly and disagreeing with Excel. A ⚠ saying "this lookup
 * needs sorted data" is information; a plausible wrong row is a liability.
 */

import { compare } from '../coerce.js';
import { Unsupported } from '../errors.js';
import { isMissingArg, scalarOf, type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, isRef, Matrix, RefValue, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix, asRef, num, optNum, trunc } from './util.js';
import { wildcardToRegExp } from './text.js';

export function registerLookup(): void {
  fn('VLOOKUP', 3, 4, (args, ctx) => lookupTable(args, ctx, 'v'));
  fn('HLOOKUP', 3, 4, (args, ctx) => lookupTable(args, ctx, 'h'));

  fn('MATCH', 2, 3, (args, ctx) => {
    const key = scalarOf(args[0]!, ctx);
    if (isErr(key)) return key;
    const arr = asMatrix(args[1]!, ctx);
    if (isErr(arr)) return arr;
    const type = optNum(args, 2, ctx, 1);
    if (isErr(type)) return type;
    const vals = [...arr.values()];
    const i = matchIndex(key, vals, trunc(type), 'MATCH');
    return isErr(i) ? i : i + 1;
  });

  fn('INDEX', 2, 3, (args, ctx) => {
    const r = num(args, 1, ctx);
    if (isErr(r)) return r;
    const c = optNum(args, 2, ctx, 0);
    if (isErr(c)) return c;

    const row = trunc(r);
    const col = trunc(c);
    if (row < 0 || col < 0) return ERR.value;

    // A whole-column or whole-row reference names far more cells than we
    // materialise, and INDEX addresses it by position. Resolve the position
    // against the extent the formula wrote and read that one cell, rather than
    // against the clamped rectangle — where a row below the used range looks
    // out of range and becomes #REF!, which is what a real board pack hit.
    // Reading a single cell is also strictly cheaper than materialising.
    const ref = asRef(args[0]!);
    if (ref?.isClamped) {
      const spot = locate(ref.declaredRows, ref.declaredCols, row, col, args[2] === undefined);
      if (spot === 'out-of-range') return ERR.ref;
      if (spot !== 'array') {
        return ctx.host.cellValue(ref.sheet, ref.top + spot.row - 1, ref.left + spot.col - 1);
      }
      // An array result from a whole-column reference would be a million cells.
      // Fall through to the clamped rectangle, which holds every value there is
      // — the remainder is empty by definition of the used range.
    }

    const arr = asMatrix(args[0]!, ctx);
    if (isErr(arr)) return arr;

    // A single-row or single-column source lets INDEX take one index only.
    if (args[2] === undefined && arr.rows === 1 && arr.cols > 1) {
      return row === 0 ? arr : pick(arr, 1, row);
    }
    if (args[2] === undefined && arr.cols === 1) {
      return row === 0 ? arr : pick(arr, row, 1);
    }
    if (row === 0 && col === 0) return arr;
    if (row === 0) return sliceColumn(arr, col);
    if (col === 0) return sliceRow(arr, row);
    return pick(arr, row, col);
  });

  fn('LOOKUP', 2, 3, (args, ctx) => {
    const key = scalarOf(args[0]!, ctx);
    if (isErr(key)) return key;
    const vector = asMatrix(args[1]!, ctx);
    if (isErr(vector)) return vector;

    // Two-argument LOOKUP searches the first row/column of an array and returns
    // from the last one.
    const searchVals = vectorOf(vector, 'first');
    const result = args[2] === undefined ? vectorOf(vector, 'last') : asMatrix(args[2]!, ctx);
    if (isErr(result)) return result;
    const resultVals = isErr(result) ? [] : [...(result as Matrix).values()];

    const i = matchIndex(key, searchVals, 1, 'LOOKUP');
    if (isErr(i)) return i;
    return resultVals[i] ?? ERR.na;
  });

  fn('XLOOKUP', 3, 6, (args, ctx) => {
    const key = scalarOf(args[0]!, ctx);
    if (isErr(key)) return key;
    const lookupArr = asMatrix(args[1]!, ctx);
    if (isErr(lookupArr)) return lookupArr;
    const returnArr = asMatrix(args[2]!, ctx);
    if (isErr(returnArr)) return returnArr;
    const matchMode = optNum(args, 4, ctx, 0);
    if (isErr(matchMode)) return matchMode;
    const searchMode = optNum(args, 5, ctx, 1);
    if (isErr(searchMode)) return searchMode;

    const vals = [...lookupArr.values()];
    const idx = xmatchIndex(key, vals, trunc(matchMode), trunc(searchMode));
    if (isErr(idx)) {
      // `XLOOKUP(k,a,b,,0)` writes the 4th argument as omitted, which is not the
      // same as supplying a blank fallback: Excel still answers #N/A. Testing
      // `!== undefined` alone treated the sentinel as a supplied value.
      if (idx.kind === '#N/A' && args[3] !== undefined && !isMissingArg(args[3])) {
        return scalarOf(args[3], ctx);
      }
      return idx;
    }
    // The return array may be a whole row or column wider than the lookup one.
    if (returnArr.rows === lookupArr.rows && returnArr.cols === lookupArr.cols) {
      return returnArr.data[idx] ?? ERR.na;
    }
    if (lookupArr.cols === 1) return sliceRow(returnArr, idx + 1);
    if (lookupArr.rows === 1) return sliceColumn(returnArr, idx + 1);
    return ERR.value;
  });

  fn('XMATCH', 2, 4, (args, ctx) => {
    const key = scalarOf(args[0]!, ctx);
    if (isErr(key)) return key;
    const arr = asMatrix(args[1]!, ctx);
    if (isErr(arr)) return arr;
    const matchMode = optNum(args, 2, ctx, 0);
    if (isErr(matchMode)) return matchMode;
    const searchMode = optNum(args, 3, ctx, 1);
    if (isErr(searchMode)) return searchMode;
    const i = xmatchIndex(key, [...arr.values()], trunc(matchMode), trunc(searchMode));
    return isErr(i) ? i : i + 1;
  });

  fn('ROW', 0, 1, (args, ctx) => {
    if (args.length === 0) return ctx.row;
    const r = args[0]!;
    return isRef(r) ? r.top : ERR.value;
  });
  fn('COLUMN', 0, 1, (args, ctx) => {
    if (args.length === 0) return ctx.col;
    const r = args[0]!;
    return isRef(r) ? r.left : ERR.value;
  });
  // ROWS and COLUMNS report the extent the formula named, not the rectangle we
  // chose to read. `ROWS(A:A)` is 1048576 in Excel however empty the column is.
  fn('ROWS', 1, 1, (args, ctx) => {
    const ref = asRef(args[0]!);
    if (ref) return ref.declaredRows;
    const m = asMatrix(args[0]!, ctx);
    return isErr(m) ? m : m.rows;
  });
  fn('COLUMNS', 1, 1, (args, ctx) => {
    const ref = asRef(args[0]!);
    if (ref) return ref.declaredCols;
    const m = asMatrix(args[0]!, ctx);
    return isErr(m) ? m : m.cols;
  });
  fn('AREAS', 1, 1, (args) => (isRef(args[0]!) ? 1 : ERR.value));

  fn('TRANSPOSE', 1, 1, (args, ctx) => {
    const m = asMatrix(args[0]!, ctx);
    if (isErr(m)) return m;
    const out: Scalar[] = new Array(m.size);
    for (let r = 0; r < m.rows; r++) {
      for (let c = 0; c < m.cols; c++) out[c * m.rows + r] = m.at(r, c);
    }
    return new Matrix(m.cols, m.rows, out);
  });
}

function lookupTable(args: EvalValue[], ctx: FnContext, dir: 'v' | 'h'): Scalar | Matrix {
  const key = scalarOf(args[0]!, ctx);
  if (isErr(key)) return key;
  const index = num(args, 2, ctx);
  if (isErr(index)) return index;
  const k = trunc(index);
  const approx = args[3] === undefined ? true : toLogical(scalarOf(args[3]!, ctx));
  if (isErr(approx)) return approx;
  const who = dir === 'v' ? 'VLOOKUP' : 'HLOOKUP';

  // Fast path: a range argument only needs its search line read, plus the one
  // result cell. Materialising the whole table first is a full extra copy of
  // it per call, which is what makes a lookup-per-row model quadratic in
  // allocation as well as in comparisons.
  const ref = args[1]!;
  if (isRef(ref)) {
    const limit = dir === 'v' ? ref.cols : ref.rows;
    if (k < 1 || k > limit) return ERR.ref;
    const line = readLine(ref, ctx, dir);
    const i = matchIndex(key, line, approx ? 1 : 0, who);
    if (isErr(i)) return i;
    return dir === 'v'
      ? ctx.host.cellValue(ref.sheet, ref.top + i, ref.left + k - 1)
      : ctx.host.cellValue(ref.sheet, ref.top + k - 1, ref.left + i);
  }

  const table = asMatrix(ref, ctx);
  if (isErr(table)) return table;
  const limit = dir === 'v' ? table.cols : table.rows;
  if (k < 1 || k > limit) return ERR.ref;
  const line = dir === 'v' ? columnOf(table, 0) : rowOf(table, 0);
  const i = matchIndex(key, line, approx ? 1 : 0, who);
  if (isErr(i)) return i;
  return dir === 'v' ? table.at(i, k - 1) : table.at(k - 1, i);
}

/** The first column (VLOOKUP) or first row (HLOOKUP) of a range. */
function readLine(ref: RefValue, ctx: FnContext, dir: 'v' | 'h'): Scalar[] {
  const out: Scalar[] = [];
  if (dir === 'v') {
    for (let r = ref.top; r <= ref.bottom; r++) out.push(ctx.host.cellValue(ref.sheet, r, ref.left));
  } else {
    for (let c = ref.left; c <= ref.right; c++) out.push(ctx.host.cellValue(ref.sheet, ref.top, c));
  }
  return out;
}

const toLogical = (v: Scalar): boolean | ExcelError => {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (isErr(v)) return v;
  return v.toUpperCase() !== 'FALSE';
};

/**
 * `type` follows MATCH: 0 exact (wildcards allowed), 1 largest ≤ key in
 * ascending data, -1 smallest ≥ key in descending data.
 */
function matchIndex(key: Scalar, vals: Scalar[], type: number, who: string): number | ExcelError {
  if (type === 0) {
    if (typeof key === 'string' && /[*?]/.test(key)) {
      const re = wildcardToRegExp(key, 'i');
      for (let i = 0; i < vals.length; i++) {
        const v = vals[i]!;
        if (typeof v === 'string' && re.test(v)) return i;
      }
      return ERR.na;
    }
    // Exact match is equality, not a criteria expression.
    //
    // Routing it through `makeCriteria` — the *IF family's miniature language —
    // meant a key that happens to start with a comparison operator stopped being
    // a key: `MATCH(">5",A1:A3,0)` looked for the first value above 5 instead of
    // for the literal text ">5", and `MATCH("<>",…,0)` matched the first
    // non-empty cell instead of answering #N/A. Bucket labels in a real model
    // read exactly like that — ">90 days", "<1 yr" — and the wrong row is a
    // number, which is the failure mode with no symptom. `compare` is also
    // stricter in the right direction: it keeps `MATCH("5",…,0)` off a numeric 5,
    // where the criteria parser coerced the text and matched.
    for (let i = 0; i < vals.length; i++) if (compare(vals[i]!, key) === 0) return i;
    return ERR.na;
  }

  assertSorted(vals, type > 0 ? 'ascending' : 'descending', who);

  let best = -1;
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i]!;
    if (v === null) continue;
    const c = compare(v, key);
    if (isErr(c)) continue;
    if (type > 0 ? c <= 0 : c >= 0) best = i;
    else break;
  }
  return best < 0 ? ERR.na : best;
}

/**
 * Excel's approximate match is a binary search, so it is only meaningful on
 * sorted data. Refuse rather than return the arbitrary value Excel's probe order
 * would produce on unsorted input.
 */
function assertSorted(vals: Scalar[], order: 'ascending' | 'descending', who: string): void {
  let prev: Scalar = null;
  let started = false;
  for (const v of vals) {
    if (v === null || isErr(v)) continue;
    if (started) {
      const c = compare(prev, v);
      if (isErr(c)) continue;
      if (order === 'ascending' ? c > 0 : c < 0) {
        throw new Unsupported(
          'ARG_UNSUPPORTED',
          `${who} approximate match requires ${order} data; this range is not sorted, ` +
            'and Excel’s result there depends on its internal probe order',
          who,
        );
      }
    }
    prev = v;
    started = true;
  }
}

function xmatchIndex(
  key: Scalar,
  vals: Scalar[],
  matchMode: number,
  searchMode: number,
): number | ExcelError {
  const order = searchMode === -1 ? [...vals.keys()].reverse() : [...vals.keys()];

  if (matchMode === 2) {
    const re = typeof key === 'string' ? wildcardToRegExp(key, 'i') : undefined;
    for (const i of order) {
      const v = vals[i]!;
      if (re ? typeof v === 'string' && re.test(v) : compare(v, key) === 0) return i;
    }
    return ERR.na;
  }

  if (matchMode === 0) {
    for (const i of order) {
      const c = compare(vals[i]!, key);
      if (!isErr(c) && c === 0) return i;
    }
    return ERR.na;
  }

  // -1 = exact or next smaller, 1 = exact or next larger.
  let best = -1;
  let bestVal: Scalar = null;
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i]!;
    if (v === null || isErr(v)) continue;
    const c = compare(v, key);
    if (isErr(c)) continue;
    if (c === 0) return i;
    const eligible = matchMode < 0 ? c < 0 : c > 0;
    if (!eligible) continue;
    if (best < 0) {
      best = i;
      bestVal = v;
      continue;
    }
    const better = compare(v, bestVal);
    if (isErr(better)) continue;
    if (matchMode < 0 ? better > 0 : better < 0) {
      best = i;
      bestVal = v;
    }
  }
  return best < 0 ? ERR.na : best;
}

const columnOf = (m: Matrix, c: number): Scalar[] =>
  Array.from({ length: m.rows }, (_, r) => m.at(r, c));
const rowOf = (m: Matrix, r: number): Scalar[] =>
  Array.from({ length: m.cols }, (_, c) => m.at(r, c));

const vectorOf = (m: Matrix, which: 'first' | 'last'): Scalar[] => {
  if (m.rows === 1 || m.cols >= m.rows) {
    return which === 'first' ? rowOf(m, 0) : rowOf(m, m.rows - 1);
  }
  return which === 'first' ? columnOf(m, 0) : columnOf(m, m.cols - 1);
};

/**
 * Resolve INDEX's one-or-two indices to a 1-based position in a rows x cols
 * source, mirroring the shorthand rules applied to a materialised matrix: a
 * single index counts columns across a single row, rows down a single column,
 * and otherwise selects a whole row.
 */
function locate(
  rows: number,
  cols: number,
  row: number,
  col: number,
  singleIndex: boolean,
): { row: number; col: number } | 'array' | 'out-of-range' {
  let r = row;
  let c = col;
  if (singleIndex) {
    if (rows === 1 && cols > 1) {
      r = 1;
      c = row;
    } else if (cols === 1) {
      r = row;
      c = 1;
    }
  }
  if (r === 0 || c === 0) return 'array';
  if (r > rows || c > cols) return 'out-of-range';
  return { row: r, col: c };
}

const pick = (m: Matrix, row: number, col: number): Scalar =>
  row > m.rows || col > m.cols ? ERR.ref : m.at(row - 1, col - 1);

const sliceRow = (m: Matrix, row: number): Matrix | Scalar =>
  row > m.rows ? ERR.ref : new Matrix(1, m.cols, rowOf(m, row - 1));

const sliceColumn = (m: Matrix, col: number): Matrix | Scalar =>
  col > m.cols ? ERR.ref : new Matrix(m.rows, 1, columnOf(m, col - 1));
