import { textToNumber, toBool } from '../coerce.js';
import { type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { arg1, asMatrix, collectNumbers, collectNumbersA, flatten, num, optNum, trunc } from './util.js';
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
  const forecast = (args: EvalValue[], ctx: FnContext): Scalar => {
    const x = num(args, 0, ctx);
    if (isErr(x)) return x;
    return pairwise(args.slice(1), ctx, (ys, xs) => {
      const m = slope(xs, ys);
      if (isErr(m)) return m;
      return mean(ys) + m * (x - mean(xs));
    });
  };
  fn('FORECAST', 3, 3, forecast);
  // The 2016 rename. Both names ship in current Excel and a generator picks
  // whichever its training data favoured, so refusing one is a coin flip.
  fn('FORECAST.LINEAR', 3, 3, forecast);

  // The normal distribution. Reached for by anything that sizes a buffer from a
  // service level — safety stock, VaR, capacity planning — which is to say by
  // every operations model, and by no finance model, which is exactly how a
  // function library ends up shaped like its own test set.
  fn('NORM.S.DIST', 2, 2, (args, ctx) => {
    const z = num(args, 0, ctx);
    if (isErr(z)) return z;
    const cumulative = toBool(arg1(args, 1, ctx));
    if (isErr(cumulative)) return cumulative;
    return cumulative ? normalCdf(z) : normalPdf(z);
  });

  fn('NORMSDIST', 1, 1, (args, ctx) => {
    const z = num(args, 0, ctx);
    return isErr(z) ? z : normalCdf(z);
  });

  fn('NORM.DIST', 4, 4, (args, ctx) => {
    const x = num(args, 0, ctx);
    if (isErr(x)) return x;
    const mu = num(args, 1, ctx);
    if (isErr(mu)) return mu;
    const sigma = num(args, 2, ctx);
    if (isErr(sigma)) return sigma;
    if (sigma <= 0) return ERR.num;
    const cumulative = toBool(arg1(args, 3, ctx));
    if (isErr(cumulative)) return cumulative;
    const z = (x - mu) / sigma;
    return cumulative ? normalCdf(z) : normalPdf(z) / sigma;
  });

  const inverse = (args: EvalValue[], ctx: FnContext): Scalar => {
    const p = num(args, 0, ctx);
    if (isErr(p)) return p;
    if (p <= 0 || p >= 1) return ERR.num;
    return normalInv(p);
  };
  fn('NORM.S.INV', 1, 1, inverse);
  fn('NORMSINV', 1, 1, inverse);

  fn('NORM.INV', 3, 3, (args, ctx) => {
    const p = num(args, 0, ctx);
    if (isErr(p)) return p;
    const mu = num(args, 1, ctx);
    if (isErr(mu)) return mu;
    const sigma = num(args, 2, ctx);
    if (isErr(sigma)) return sigma;
    if (sigma <= 0) return ERR.num;
    if (p <= 0 || p >= 1) return ERR.num;
    return mu + sigma * normalInv(p);
  });
}

const SQRT_2PI = Math.sqrt(2 * Math.PI);

const normalPdf = (z: number): number => Math.exp((-z * z) / 2) / SQRT_2PI;

/**
 * The standard normal CDF, by Graeme West's double-precision refinement of
 * Hart's rational approximation. Accurate to about 1e-15 across the whole
 * range, which matters because the tail is where a service level actually
 * lives: a 99.9% target is 3.09 standard deviations out, and a cheap
 * approximation that is fine near the mean is visibly wrong there.
 */
function normalCdf(z: number): number {
  const a = Math.abs(z);
  let tail: number;
  if (a > 37) {
    tail = 0;
  } else {
    const e = Math.exp((-a * a) / 2);
    if (a < 7.07106781186547) {
      let n = 3.52624965998911e-2 * a + 0.700383064443688;
      n = n * a + 6.37396220353165;
      n = n * a + 33.912866078383;
      n = n * a + 112.079291497871;
      n = n * a + 221.213596169931;
      n = n * a + 220.206867912376;
      let d = 8.83883476483184e-2 * a + 1.75566716318264;
      d = d * a + 16.064177579207;
      d = d * a + 86.7807322029461;
      d = d * a + 296.564248779674;
      d = d * a + 637.333633378831;
      d = d * a + 793.826512519948;
      d = d * a + 440.413735824752;
      tail = (e * n) / d;
    } else {
      let cf = a + 0.65;
      cf = a + 4 / cf;
      cf = a + 3 / cf;
      cf = a + 2 / cf;
      cf = a + 1 / cf;
      tail = e / cf / 2.506628274631;
    }
  }
  return z > 0 ? 1 - tail : tail;
}

/**
 * The inverse: Acklam's rational approximation (about 1.15e-9 relative) plus one
 * Halley step against the CDF above, which takes it to full double precision.
 */
function normalInv(p: number): number {
  const A = [
    -3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2,
    -3.066479806614716e1, 2.506628277459239,
  ];
  const B = [
    -5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1,
    -1.328068155288572e1,
  ];
  const C = [
    -7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734,
    4.374664141464968, 2.938163982698783,
  ];
  const D = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const LOW = 0.02425;

  let x: number;
  if (p < LOW) {
    const q = Math.sqrt(-2 * Math.log(p));
    x =
      (((((C[0]! * q + C[1]!) * q + C[2]!) * q + C[3]!) * q + C[4]!) * q + C[5]!) /
      ((((D[0]! * q + D[1]!) * q + D[2]!) * q + D[3]!) * q + 1);
  } else if (p <= 1 - LOW) {
    const q = p - 0.5;
    const r = q * q;
    x =
      ((((((A[0]! * r + A[1]!) * r + A[2]!) * r + A[3]!) * r + A[4]!) * r + A[5]!) * q) /
      (((((B[0]! * r + B[1]!) * r + B[2]!) * r + B[3]!) * r + B[4]!) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x =
      -(((((C[0]! * q + C[1]!) * q + C[2]!) * q + C[3]!) * q + C[4]!) * q + C[5]!) /
      ((((D[0]! * q + D[1]!) * q + D[2]!) * q + D[3]!) * q + 1);
  }

  const e = normalCdf(x) - p;
  const u = e * SQRT_2PI * Math.exp((x * x) / 2);
  return x - u / (1 + (x * u) / 2);
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
