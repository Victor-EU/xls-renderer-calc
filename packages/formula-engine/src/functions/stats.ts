import { textToNumber } from '../coerce.js';
import { type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix, collectNumbers, collectNumbersA, flatten, num, optNum, trunc } from './util.js';
import { excelRound } from './math.js';

export function registerStats(): void {
  const overNumbers = (name: string, f: (ns: number[]) => Scalar, variant: 'plain' | 'a' = 'plain'): void =>
    fn(name, 1, -1, (args, ctx) => {
      const ns = variant === 'a' ? collectNumbersA(args, ctx) : collectNumbers(args, ctx);
      return isErr(ns) ? ns : f(ns);
    });

  overNumbers('AVERAGE', (ns) => (ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : ERR.div0));
  overNumbers('AVERAGEA', (ns) => (ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : ERR.div0), 'a');
  // MIN/MAX over an empty set are 0 in Excel, not an error.
  overNumbers('MIN', (ns) => (ns.length ? Math.min(...ns) : 0));
  overNumbers('MAX', (ns) => (ns.length ? Math.max(...ns) : 0));
  overNumbers('MINA', (ns) => (ns.length ? Math.min(...ns) : 0), 'a');
  overNumbers('MAXA', (ns) => (ns.length ? Math.max(...ns) : 0), 'a');
  overNumbers('MEDIAN', (ns) => {
    if (!ns.length) return ERR.num;
    const s = [...ns].sort((a, b) => a - b);
    const mid = s.length >> 1;
    return s.length % 2 ? s[mid]! : (s[mid - 1]! + s[mid]!) / 2;
  });

  overNumbers('STDEV.S', (ns) => spread(ns, true, true));
  overNumbers('STDEV', (ns) => spread(ns, true, true));
  overNumbers('STDEV.P', (ns) => spread(ns, false, true));
  overNumbers('STDEVP', (ns) => spread(ns, false, true));
  overNumbers('VAR.S', (ns) => spread(ns, true, false));
  overNumbers('VAR', (ns) => spread(ns, true, false));
  overNumbers('VAR.P', (ns) => spread(ns, false, false));
  overNumbers('VARP', (ns) => spread(ns, false, false));

  /** COUNT sees numbers only; COUNTA sees everything except a truly empty cell. */
  fn('COUNT', 1, -1, (args, ctx) => {
    let n = 0;
    for (const { v, direct } of flatten(args, ctx)) {
      if (typeof v === 'number') n++;
      else if (!direct || v === null || isErr(v)) continue;
      else if (typeof v === 'boolean') n++;
      else if (textToNumber(v) !== undefined) n++;
    }
    return n;
  });

  fn('COUNTA', 1, -1, (args, ctx) => {
    let n = 0;
    for (const { v } of flatten(args, ctx)) if (v !== null) n++;
    return n;
  });

  fn('COUNTBLANK', 1, 1, (args, ctx) => {
    let n = 0;
    for (const { v } of flatten(args, ctx)) if (v === null || v === '') n++;
    return n;
  });

  fn('LARGE', 2, 2, (args, ctx) => nth(args, ctx, 'desc'));
  fn('SMALL', 2, 2, (args, ctx) => nth(args, ctx, 'asc'));

  fn('RANK', 2, 3, (args, ctx) => rankImpl(args, ctx));
  fn('RANK.EQ', 2, 3, (args, ctx) => rankImpl(args, ctx));

  fn('PERCENTILE', 2, 2, (args, ctx) => percentile(args, ctx, 'inclusive'));
  fn('PERCENTILE.INC', 2, 2, (args, ctx) => percentile(args, ctx, 'inclusive'));
  fn('PERCENTILE.EXC', 2, 2, (args, ctx) => percentile(args, ctx, 'exclusive'));
  fn('QUARTILE', 2, 2, (args, ctx) => quartile(args, ctx, 'inclusive'));
  fn('QUARTILE.INC', 2, 2, (args, ctx) => quartile(args, ctx, 'inclusive'));
  fn('QUARTILE.EXC', 2, 2, (args, ctx) => quartile(args, ctx, 'exclusive'));

  fn('CORREL', 2, 2, (args, ctx) => pairwise(args, ctx, (xs, ys) => correl(xs, ys)));
  fn('SLOPE', 2, 2, (args, ctx) => pairwise(args, ctx, (ys, xs) => slope(xs, ys)));
  fn('INTERCEPT', 2, 2, (args, ctx) =>
    pairwise(args, ctx, (ys, xs) => {
      const m = slope(xs, ys);
      if (isErr(m)) return m;
      return mean(ys) - m * mean(xs);
    }),
  );
  fn('RSQ', 2, 2, (args, ctx) =>
    pairwise(args, ctx, (xs, ys) => {
      const r = correl(xs, ys);
      return isErr(r) ? r : r * r;
    }),
  );
  fn('FORECAST', 3, 3, (args, ctx) => {
    const x = num(args, 0, ctx);
    if (isErr(x)) return x;
    return pairwise(args.slice(1), ctx, (ys, xs) => {
      const m = slope(xs, ys);
      if (isErr(m)) return m;
      return mean(ys) + m * (x - mean(xs));
    });
  });
}

const mean = (ns: number[]): number => ns.reduce((a, b) => a + b, 0) / ns.length;

function spread(ns: number[], sample: boolean, root: boolean): Scalar {
  const n = ns.length;
  if (n === 0 || (sample && n < 2)) return ERR.div0;
  const m = mean(ns);
  const ss = ns.reduce((a, b) => a + (b - m) * (b - m), 0);
  const v = ss / (sample ? n - 1 : n);
  return root ? Math.sqrt(v) : v;
}

function nth(args: EvalValue[], ctx: FnContext, dir: 'asc' | 'desc'): Scalar {
  const ns = collectNumbers([args[0]!], ctx);
  if (isErr(ns)) return ns;
  const k = num(args, 1, ctx);
  if (isErr(k)) return k;
  const i = trunc(k);
  if (i < 1 || i > ns.length) return ERR.num;
  const sorted = [...ns].sort((a, b) => (dir === 'asc' ? a - b : b - a));
  return sorted[i - 1]!;
}

function rankImpl(args: EvalValue[], ctx: FnContext): Scalar {
  const x = num(args, 0, ctx);
  if (isErr(x)) return x;
  const ns = collectNumbers([args[1]!], ctx);
  if (isErr(ns)) return ns;
  const order = optNum(args, 2, ctx, 0);
  if (isErr(order)) return order;
  const sorted = [...ns].sort((a, b) => (order === 0 ? b - a : a - b));
  const i = sorted.indexOf(x);
  return i < 0 ? ERR.na : i + 1;
}

function percentileOf(sorted: number[], p: number, mode: 'inclusive' | 'exclusive'): Scalar {
  const n = sorted.length;
  if (n === 0) return ERR.num;
  if (mode === 'inclusive') {
    if (p < 0 || p > 1) return ERR.num;
    const pos = p * (n - 1);
    const lo = Math.floor(pos);
    const hi = Math.ceil(pos);
    return sorted[lo]! + (sorted[hi]! - sorted[lo]!) * (pos - lo);
  }
  if (p <= 0 || p >= 1 || p < 1 / (n + 1) || p > n / (n + 1)) return ERR.num;
  const pos = p * (n + 1) - 1;
  const lo = Math.floor(pos);
  const hi = Math.min(n - 1, Math.ceil(pos));
  return sorted[lo]! + (sorted[hi]! - sorted[lo]!) * (pos - lo);
}

function percentile(args: EvalValue[], ctx: FnContext, mode: 'inclusive' | 'exclusive'): Scalar {
  const ns = collectNumbers([args[0]!], ctx);
  if (isErr(ns)) return ns;
  const p = num(args, 1, ctx);
  if (isErr(p)) return p;
  return percentileOf([...ns].sort((a, b) => a - b), p, mode);
}

function quartile(args: EvalValue[], ctx: FnContext, mode: 'inclusive' | 'exclusive'): Scalar {
  const q = num(args, 1, ctx);
  if (isErr(q)) return q;
  const k = trunc(q);
  if (k < 0 || k > 4) return ERR.num;
  const ns = collectNumbers([args[0]!], ctx);
  if (isErr(ns)) return ns;
  return percentileOf([...ns].sort((a, b) => a - b), k / 4, mode);
}

/** Pair two ranges positionally, dropping pairs where either side is non-numeric. */
function pairwise(
  args: EvalValue[],
  ctx: FnContext,
  f: (a: number[], b: number[]) => Scalar,
): Scalar {
  const ma = asMatrix(args[0]!, ctx);
  if (isErr(ma)) return ma;
  const mb = asMatrix(args[1]!, ctx);
  if (isErr(mb)) return mb;
  if (ma.size !== mb.size) return ERR.na;
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i < ma.size; i++) {
    const a = ma.data[i]!;
    const b = mb.data[i]!;
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (typeof a === 'number' && typeof b === 'number') {
      xs.push(a);
      ys.push(b);
    }
  }
  if (xs.length === 0) return ERR.div0;
  return f(xs, ys);
}

function correl(xs: number[], ys: number[]): number | ExcelError {
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < xs.length; i++) {
    const dx = xs[i]! - mx;
    const dy = ys[i]! - my;
    sxy += dx * dy;
    sxx += dx * dx;
    syy += dy * dy;
  }
  const d = Math.sqrt(sxx * syy);
  return d === 0 ? ERR.div0 : sxy / d;
}

function slope(xs: number[], ys: number[]): number | ExcelError {
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i]! - mx) * (ys[i]! - my);
    sxx += (xs[i]! - mx) ** 2;
  }
  return sxx === 0 ? ERR.div0 : sxy / sxx;
}
