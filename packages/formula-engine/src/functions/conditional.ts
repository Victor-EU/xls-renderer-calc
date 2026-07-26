/**
 * The `*IF` family.
 *
 * Their criteria argument is a miniature expression language — `">100"`,
 * `"<>x"`, `"North*"` — so these are a parsing task, not a one-line filter.
 * Everything routes through `makeCriteria` so SUMIF, COUNTIFS and MAXIFS cannot
 * drift apart in how they read `">="`.
 */

import { compare, textToNumber, toText } from '../coerce.js';
import { type FnContext } from '../interpreter.js';
import { MAX_COLS, MAX_ROWS } from '../a1.js';
import { ERR, ExcelError, isErr, Matrix, RefValue, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix, asRef, maxOf, minOf } from './util.js';
import { wildcardToRegExp } from './text.js';

export type Predicate = (v: Scalar) => boolean;

const OPERATORS = ['<=', '>=', '<>', '=', '<', '>'] as const;

export function makeCriteria(crit: Scalar): Predicate {
  if (crit === null) return (v) => v === null || v === '';
  if (typeof crit !== 'string') return (v) => compare(v, crit) === 0;

  let op: (typeof OPERATORS)[number] | '' = '';
  let rest = crit;
  for (const candidate of OPERATORS) {
    if (crit.startsWith(candidate)) {
      op = candidate;
      rest = crit.slice(candidate.length);
      break;
    }
  }

  // `"="` and `""` both mean "empty", which is not the same as "equals nothing".
  if ((op === '=' || op === '') && rest === '') return (v) => v === null || v === '';
  if (op === '<>' && rest === '') return (v) => v !== null && v !== '';

  const asNumber = textToNumber(rest);
  const target: Scalar =
    asNumber !== undefined
      ? asNumber
      : /^(TRUE|FALSE)$/i.test(rest)
        ? rest.toUpperCase() === 'TRUE'
        : rest;

  if (op === '' || op === '=' || op === '<>') {
    const negate = op === '<>';
    // Wildcards only apply to text equality tests.
    if (typeof target === 'string' && /[*?]/.test(rest)) {
      const re = wildcardToRegExp(rest, 'i');
      return (v) => {
        const t = toText(v);
        const hit = !isErr(t) && v !== null && re.test(t);
        return negate ? !hit : hit;
      };
    }
    return (v) => {
      const c = compare(v, target);
      const hit = !isErr(c) && c === 0;
      return negate ? !hit : hit;
    };
  }

  return (v) => {
    // Blank cells do not satisfy ordering comparisons in the *IF family.
    if (v === null) return false;
    // Nor do values of a different type from the criterion. This is where the
    // *IF family parts company with the `>` operator: Excel's general ordering
    // ranks number < text < boolean, so `="text">0` really is TRUE, but
    // `COUNTIF(range,">0")` counts only numbers. Without this line every label
    // and every "" in a column satisfies ">0", which silently inflates
    // conditional sums over any range that is not purely numeric — a cohort
    // triangle, a column with a header, a lookup table with a footnote.
    if (typeClass(v) !== typeClass(target)) return false;
    const c = compare(v, target);
    if (isErr(c)) return false;
    switch (op) {
      case '<':
        return c < 0;
      case '<=':
        return c <= 0;
      case '>':
        return c > 0;
      default:
        return c >= 0;
    }
  };
}

/** Excel's criteria type classes. Ordering only compares within one of them. */
function typeClass(v: Scalar): 'number' | 'text' | 'boolean' | 'other' {
  if (typeof v === 'number') return 'number';
  if (typeof v === 'string') return 'text';
  if (typeof v === 'boolean') return 'boolean';
  return 'other';
}

/**
 * Run `compute` once per element of an array-valued criteria argument.
 *
 * `COUNTIF(A1:A5,A1:A5)` is not a typo — handed a range where it expects one
 * criterion, Excel evaluates the function once per element and returns an
 * array. That is the whole basis of the distinct-count idiom,
 * `SUMPRODUCT(1/COUNTIF(range,range))`, which is common enough in real
 * spreadsheets that collapsing the criteria to its first cell turns a correct
 * formula into a plausible wrong number rather than an error.
 *
 * The all-scalar case — every criteria argument a single value — takes the
 * fast path and allocates nothing extra, which is what a 24,000-formula sheet
 * of ordinary SUMIFS actually does.
 */
function overCriteria(
  ctx: FnContext,
  criteria: EvalValue[],
  compute: (crits: Scalar[]) => Scalar,
): EvalValue {
  const mats: Matrix[] = [];
  for (const c of criteria) {
    const m = asMatrix(c, ctx);
    if (isErr(m)) return m;
    mats.push(m);
  }
  const first = (m: Matrix): Scalar => (m.size === 0 ? null : m.data[0]!);
  const shaped = mats.find((m) => m.size > 1);
  if (!shaped) return compute(mats.map(first));

  const out: Scalar[] = new Array(shaped.size);
  for (let i = 0; i < shaped.size; i++) {
    out[i] = compute(mats.map((m) => (m.size > 1 ? (m.data[i] ?? null) : first(m))));
  }
  return new Matrix(shaped.rows, shaped.cols, out);
}

/**
 * Excel's sum/average range is a *corner*, not a range.
 *
 * `SUMIF(C1:P56, x, P1:P56)` does not sum the 56 cells of P1:P56. Excel takes
 * the top-left of the third argument and reshapes it to the criteria range's
 * dimensions — here 56 rows by 14 columns, so P1:AC56 — and pairs the two
 * position by position. Authors reach this by dragging a formula sideways
 * until the criteria range grows and the sum range does not, which is common
 * enough that Excel made it work rather than an error.
 *
 * Reading the values alone cannot reproduce it, because the extra columns
 * were never fetched. Indexing a 56-cell matrix by a position that runs to
 * 784 silently finds nothing and contributes nothing, which is how a board
 * pack came to show 0 where Excel shows 89,263 — a plausible number, no
 * warning, on 71 cells. So this resolves the *reference* and re-reads it at
 * the shape Excel would use.
 *
 * It applies to the whole family, not just SUMIF. SUMIFS, AVERAGEIFS, MAXIFS
 * and MINIFS take their aggregate range *first* and were reading it as-is, so
 * `SUMIFS(E1,C1:D3,1)` aggregated the single cell E1 and answered 10 where
 * Excel reshapes E1 to E1:F3 and answers 170. Same bug, four functions, and
 * the four that hid it longest are the ones a real model reaches for most.
 */
function align(arg: EvalValue, range: Matrix, ctx: FnContext): Matrix | ExcelError {
  const ref = asRef(arg);
  if (!ref || (ref.rows === range.rows && ref.cols === range.cols)) return asMatrix(arg, ctx);
  const bottom = ref.top + range.rows - 1;
  const right = ref.left + range.cols - 1;
  if (bottom > MAX_ROWS || right > MAX_COLS) return ERR.ref;
  return asMatrix(new RefValue(ref.sheet, ref.top, ref.left, bottom, right), ctx);
}

export function registerConditional(): void {
  fn('COUNTIF', 2, 2, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    return overCriteria(ctx, [args[1]!], (crits) => {
      const pred = makeCriteria(crits[0]!);
      let n = 0;
      for (const v of range.values()) if (pred(v)) n++;
      return n;
    });
  });

  fn('SUMIF', 2, 3, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    const target = args[2] === undefined ? range : align(args[2]!, range, ctx);
    if (isErr(target)) return target;
    return overCriteria(ctx, [args[1]!], (crits) => {
      const pred = makeCriteria(crits[0]!);
      let total = 0;
      for (let i = 0; i < range.size; i++) {
        if (!pred(range.data[i]!)) continue;
        const v = target.data[i];
        if (typeof v === 'number') total += v;
        else if (isErr(v)) return v;
      }
      return total;
    });
  });

  fn('AVERAGEIF', 2, 3, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    const pred = makeCriteria(scalarCriteria(args[1]!, ctx));
    const target = args[2] === undefined ? range : align(args[2]!, range, ctx);
    if (isErr(target)) return target;
    let total = 0;
    let n = 0;
    for (let i = 0; i < range.size; i++) {
      if (!pred(range.data[i]!)) continue;
      const v = target.data[i];
      if (typeof v === 'number') {
        total += v;
        n++;
      } else if (isErr(v)) return v;
    }
    return n === 0 ? ERR.div0 : total / n;
  });

  fn('COUNTIFS', 2, -1, (args, ctx) =>
    overCriteria(ctx, criteriaArgs(args, 0), (crits) => {
      const hits = multiMatch(args, ctx, 0, crits);
      return isErr(hits) ? hits : hits.length;
    }),
  );

  fn('SUMIFS', 3, -1, (args, ctx) => aggregateIfs(args, ctx, 'sum'));
  fn('AVERAGEIFS', 3, -1, (args, ctx) => aggregateIfs(args, ctx, 'average'));
  fn('MAXIFS', 3, -1, (args, ctx) => aggregateIfs(args, ctx, 'max'));
  fn('MINIFS', 3, -1, (args, ctx) => aggregateIfs(args, ctx, 'min'));
}

function scalarCriteria(v: EvalValue, ctx: FnContext): Scalar {
  const m = asMatrix(v, ctx);
  if (isErr(m)) return m;
  return m.size === 0 ? null : m.at(0, 0);
}

/**
 * Indices satisfying every (range, criteria) pair, starting at `from`.
 * All criteria ranges must have the same shape — Excel returns #VALUE! otherwise.
 */
function criteriaArgs(args: EvalValue[], from: number): EvalValue[] {
  const out: EvalValue[] = [];
  for (let i = from; i + 1 < args.length; i += 2) out.push(args[i + 1]!);
  return out;
}

function multiMatch(
  args: EvalValue[],
  ctx: FnContext,
  from: number,
  crits?: Scalar[],
): number[] | ExcelError {
  const pairs: Array<{ m: Matrix; p: Predicate }> = [];
  let k = 0;
  for (let i = from; i + 1 < args.length; i += 2, k++) {
    const m = asMatrix(args[i]!, ctx);
    if (isErr(m)) return m;
    const crit = crits ? crits[k]! : scalarCriteria(args[i + 1]!, ctx);
    pairs.push({ m, p: makeCriteria(crit) });
  }
  if (pairs.length === 0) return ERR.value;
  const size = pairs[0]!.m.size;
  if (pairs.some((x) => x.m.size !== size)) return ERR.value;

  const out: number[] = [];
  for (let i = 0; i < size; i++) {
    if (pairs.every((x) => x.p(x.m.data[i]!))) out.push(i);
  }
  return out;
}

function aggregateIfs(
  args: EvalValue[],
  ctx: FnContext,
  mode: 'sum' | 'average' | 'max' | 'min',
): Scalar {
  // The aggregate range is reshaped to the first criteria range, exactly as
  // SUMIF's third argument is — see `align`.
  const first = asMatrix(args[1]!, ctx);
  if (isErr(first)) return first;
  const target = align(args[0]!, first, ctx);
  if (isErr(target)) return target;
  const idx = multiMatch(args, ctx, 1);
  if (!Array.isArray(idx)) return idx;

  const picked: number[] = [];
  for (const i of idx) {
    const v = target.data[i];
    if (isErr(v)) return v;
    if (typeof v === 'number') picked.push(v);
  }
  switch (mode) {
    case 'sum':
      return picked.reduce((a, b) => a + b, 0);
    case 'average':
      return picked.length ? picked.reduce((a, b) => a + b, 0) / picked.length : ERR.div0;
    case 'max':
      return picked.length ? maxOf(picked) : 0;
    default:
      return picked.length ? minOf(picked) : 0;
  }
}
