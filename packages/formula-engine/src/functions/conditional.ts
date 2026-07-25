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
import { ERR, ExcelError, isErr, Matrix, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix } from './util.js';
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

export function registerConditional(): void {
  fn('COUNTIF', 2, 2, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    const pred = makeCriteria(scalarCriteria(args[1]!, ctx));
    let n = 0;
    for (const v of range.values()) if (pred(v)) n++;
    return n;
  });

  fn('SUMIF', 2, 3, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    const pred = makeCriteria(scalarCriteria(args[1]!, ctx));
    const target = args[2] === undefined ? range : asMatrix(args[2]!, ctx);
    if (isErr(target)) return target;
    let total = 0;
    for (let i = 0; i < range.size; i++) {
      if (!pred(range.data[i]!)) continue;
      // The sum range is addressed positionally from its top-left corner, so a
      // shorter one is read as if it extended to match — Excel's actual behaviour.
      const v = target.data[i];
      if (typeof v === 'number') total += v;
      else if (isErr(v)) return v;
    }
    return total;
  });

  fn('AVERAGEIF', 2, 3, (args, ctx) => {
    const range = asMatrix(args[0]!, ctx);
    if (isErr(range)) return range;
    const pred = makeCriteria(scalarCriteria(args[1]!, ctx));
    const target = args[2] === undefined ? range : asMatrix(args[2]!, ctx);
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

  fn('COUNTIFS', 2, -1, (args, ctx) => {
    const hits = multiMatch(args, ctx, 0);
    return isErr(hits) ? hits : hits.length;
  });

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
function multiMatch(args: EvalValue[], ctx: FnContext, from: number): number[] | ExcelError {
  const pairs: Array<{ m: Matrix; p: Predicate }> = [];
  for (let i = from; i + 1 < args.length; i += 2) {
    const m = asMatrix(args[i]!, ctx);
    if (isErr(m)) return m;
    pairs.push({ m, p: makeCriteria(scalarCriteria(args[i + 1]!, ctx)) });
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
  const target = asMatrix(args[0]!, ctx);
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
      return picked.length ? Math.max(...picked) : 0;
    default:
      return picked.length ? Math.min(...picked) : 0;
  }
}
