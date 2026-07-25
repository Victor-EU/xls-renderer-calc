import { ERR, ExcelError, isErr, Matrix, type EvalValue, type Scalar } from '../values.js';
import { Unsupported } from '../errors.js';
import { type FnContext } from '../interpreter.js';
import { fn, refuse } from './registry.js';
import { asMatrix, collectNumbers, finiteOr, num, optNum, trunc } from './util.js';

/**
 * Excel rounds the *decimal* value, not the binary double.
 *
 * `ROUND(2.675, 2)` is 2.68 in Excel. In IEEE-754, 2.675 is really
 * 2.67499999999999982…, so `Math.round(2.675 * 100) / 100` gives 2.67 — a
 * one-cent disagreement that shows up in any model with rounded currency.
 * Rounding the 15-significant-digit decimal string instead reproduces Excel.
 */
export function excelRound(n: number, digits: number, mode: 'half' | 'up' | 'down'): number {
  if (!Number.isFinite(n) || n === 0) return n === 0 ? 0 : n;
  const neg = n < 0;
  const a = Math.abs(n);

  const [mant, expPart] = a.toExponential(14).split('e');
  const exp = Number(expPart);
  const ds = mant!.replace('.', ''); // exactly 15 significant digits
  const keep = exp + digits + 1; // how many of those digits survive

  if (keep >= 15) return n; // already finer than the requested precision

  const decide = (firstDropped: number, restNonZero: boolean): boolean => {
    if (mode === 'down') return false;
    if (mode === 'up') return firstDropped > 0 || restNonZero;
    return firstDropped >= 5;
  };

  if (keep <= 0) {
    // Everything is dropped; the result is 0 or one unit in the last place.
    const first = keep === 0 ? Number(ds[0]) : 0;
    const rest = keep === 0 ? /[1-9]/.test(ds.slice(1)) : true;
    const up = decide(first, rest);
    const v = up ? Math.pow(10, -digits) : 0;
    return neg ? -v : v;
  }

  const kept = ds.slice(0, keep);
  const first = Number(ds[keep]);
  const rest = /[1-9]/.test(ds.slice(keep + 1));
  let value = Number(kept);
  if (decide(first, rest)) value += 1;

  const scaled = value * Math.pow(10, exp - keep + 1);
  return neg ? -scaled : scaled;
}

const unary = (name: string, f: (n: number) => number | ExcelError): void =>
  fn(name, 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    return isErr(n) ? n : f(n);
  });

export function registerMath(): void {
  fn('SUM', 1, -1, (args, ctx) => {
    const ns = collectNumbers(args, ctx);
    return isErr(ns) ? ns : ns.reduce((a, b) => a + b, 0);
  });

  fn('PRODUCT', 1, -1, (args, ctx) => {
    const ns = collectNumbers(args, ctx);
    if (isErr(ns)) return ns;
    return ns.length === 0 ? 0 : ns.reduce((a, b) => a * b, 1);
  });

  fn('SUMSQ', 1, -1, (args, ctx) => {
    const ns = collectNumbers(args, ctx);
    return isErr(ns) ? ns : ns.reduce((a, b) => a + b * b, 0);
  });

  unary('ABS', Math.abs);
  unary('SIGN', Math.sign);
  unary('INT', Math.floor); // INT floors — it does NOT truncate: INT(-2.5) = -3
  unary('EXP', (n) => finiteOr(Math.exp(n)));
  unary('SQRT', (n) => (n < 0 ? ERR.num : Math.sqrt(n)));
  unary('LN', (n) => (n <= 0 ? ERR.num : Math.log(n)));
  unary('LOG10', (n) => (n <= 0 ? ERR.num : Math.log10(n)));
  unary('EVEN', (n) => (n === 0 ? 0 : Math.sign(n) * 2 * Math.ceil(Math.abs(n) / 2)));
  unary('ODD', (n) => {
    if (n === 0) return 1;
    const a = Math.abs(n);
    const r = Math.floor(a) % 2 === 1 && a === Math.floor(a) ? a : 2 * Math.floor(a / 2) + 1;
    return Math.sign(n) * (r < a ? r + 2 : r);
  });

  fn('TRUNC', 1, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const d = optNum(args, 1, ctx, 0);
    if (isErr(d)) return d;
    return excelRound(n, trunc(d), 'down');
  });

  fn('ROUND', 2, 2, (args, ctx) => roundImpl(args, ctx, 'half'));
  fn('ROUNDUP', 2, 2, (args, ctx) => roundImpl(args, ctx, 'up'));
  fn('ROUNDDOWN', 2, 2, (args, ctx) => roundImpl(args, ctx, 'down'));

  fn('MROUND', 2, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const m = num(args, 1, ctx);
    if (isErr(m)) return m;
    if (m === 0) return 0;
    if (Math.sign(n) !== Math.sign(m) && n !== 0) return ERR.num;
    return excelRound(n / m, 0, 'half') * m;
  });

  fn('CEILING', 2, 2, (args, ctx) => stepRound(args, ctx, 'ceil'));
  fn('FLOOR', 2, 2, (args, ctx) => stepRound(args, ctx, 'floor'));

  fn('CEILING.MATH', 1, 3, (args, ctx) => mathStep(args, ctx, 'ceil'));
  fn('FLOOR.MATH', 1, 3, (args, ctx) => mathStep(args, ctx, 'floor'));

  fn('MOD', 2, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const d = num(args, 1, ctx);
    if (isErr(d)) return d;
    if (d === 0) return ERR.div0;
    // Excel's MOD takes the sign of the divisor: MOD(-3,2) = 1, not -1.
    return n - d * Math.floor(n / d);
  });

  fn('QUOTIENT', 2, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const d = num(args, 1, ctx);
    if (isErr(d)) return d;
    return d === 0 ? ERR.div0 : trunc(n / d);
  });

  fn('POWER', 2, 2, (args, ctx) => {
    const b = num(args, 0, ctx);
    if (isErr(b)) return b;
    const e = num(args, 1, ctx);
    if (isErr(e)) return e;
    const p = Math.pow(b, e);
    return Number.isFinite(p) ? p : ERR.num;
  });

  fn('LOG', 1, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const base = optNum(args, 1, ctx, 10);
    if (isErr(base)) return base;
    if (n <= 0 || base <= 0 || base === 1) return ERR.num;
    return Math.log(n) / Math.log(base);
  });

  fn('PI', 0, 0, () => Math.PI);

  for (const [name, f] of [
    ['SIN', Math.sin],
    ['COS', Math.cos],
    ['TAN', Math.tan],
    ['SINH', Math.sinh],
    ['COSH', Math.cosh],
    ['TANH', Math.tanh],
  ] as const) {
    unary(name, (n) => finiteOr(f(n)));
  }
  unary('ASIN', (n) => (n < -1 || n > 1 ? ERR.num : Math.asin(n)));
  unary('ACOS', (n) => (n < -1 || n > 1 ? ERR.num : Math.acos(n)));
  unary('ATAN', Math.atan);
  unary('DEGREES', (n) => (n * 180) / Math.PI);
  unary('RADIANS', (n) => (n * Math.PI) / 180);
  fn('ATAN2', 2, 2, (args, ctx) => {
    const x = num(args, 0, ctx);
    if (isErr(x)) return x;
    const y = num(args, 1, ctx);
    if (isErr(y)) return y;
    if (x === 0 && y === 0) return ERR.div0;
    return Math.atan2(y, x); // Excel takes (x, y), the reverse of most languages
  });

  fn('FACT', 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const k = trunc(n);
    if (k < 0 || k > 170) return ERR.num;
    let out = 1;
    for (let i = 2; i <= k; i++) out *= i;
    return out;
  });

  fn('COMBIN', 2, 2, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const k = num(args, 1, ctx);
    if (isErr(k)) return k;
    const N = trunc(n);
    const K = trunc(k);
    if (N < 0 || K < 0 || K > N) return ERR.num;
    let out = 1;
    for (let i = 1; i <= K; i++) out = (out * (N - K + i)) / i;
    return excelRound(out, 10, 'half');
  });

  fn('GCD', 1, -1, (args, ctx) => intFold(args, ctx, (a, b) => (b === 0 ? a : gcd2(a, b)), 0));
  fn('LCM', 1, -1, (args, ctx) => intFold(args, ctx, (a, b) => (a === 0 || b === 0 ? 0 : (a / gcd2(a, b)) * b), 1));

  fn(
    'SUMPRODUCT',
    1,
    -1,
    (args, ctx) => {
      const mats: Matrix[] = [];
      for (const a of args) {
        const m = asMatrix(a, ctx);
        if (isErr(m)) return m;
        mats.push(m);
      }
      const rows = mats[0]!.rows;
      const cols = mats[0]!.cols;
      if (mats.some((m) => m.rows !== rows || m.cols !== cols)) return ERR.value;
      let total = 0;
      for (let i = 0; i < rows * cols; i++) {
        let p = 1;
        for (const m of mats) {
          const v = m.data[i]!;
          if (isErr(v)) return v;
          // SUMPRODUCT treats text and blanks as zero rather than erroring.
          p *= typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
        }
        total += p;
      }
      return total;
    },
    { arrayArgs: true },
  );

  fn('RAND', 0, 0, () => Math.random(), { volatile: true });
  fn(
    'RANDBETWEEN',
    2,
    2,
    (args, ctx) => {
      const lo = num(args, 0, ctx);
      if (isErr(lo)) return lo;
      const hi = num(args, 1, ctx);
      if (isErr(hi)) return hi;
      if (lo > hi) return ERR.num;
      return Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) + Math.ceil(lo);
    },
    { volatile: true },
  );

  fn('SUBTOTAL', 2, -1, (args, ctx) => {
    const code = num(args, 0, ctx);
    if (isErr(code)) return code;
    const k = trunc(code);
    if (k >= 101 && k <= 111) {
      // 101–111 exclude manually hidden rows. The engine has no view state, so
      // answering would mean assuming nothing is hidden — and a subtotal that
      // silently includes rows Excel excludes is exactly the failure mode this
      // project exists to prevent.
      throw new Unsupported(
        'ARG_UNSUPPORTED',
        `SUBTOTAL(${k},…) excludes hidden rows, which this engine cannot see`,
        'SUBTOTAL',
      );
    }
    const rest = args.slice(1);
    const ns = collectNumbers(rest, ctx);
    if (isErr(ns)) return ns;
    return subtotalOf(k, ns);
  });

  refuse('AGGREGATE', 'AGGREGATE() ignores hidden rows and nested results, which this engine cannot see');
  refuse('INDIRECT', 'INDIRECT() builds references at evaluation time, which breaks the dependency graph');
  refuse('OFFSET', 'OFFSET() builds references at evaluation time, which breaks the dependency graph');
}

function roundImpl(args: EvalValue[], ctx: FnContext, mode: 'half' | 'up' | 'down'): Scalar {
  const n = num(args, 0, ctx);
  if (isErr(n)) return n;
  const d = num(args, 1, ctx);
  if (isErr(d)) return d;
  return excelRound(n, trunc(d), mode);
}

function stepRound(args: EvalValue[], ctx: FnContext, dir: 'ceil' | 'floor'): Scalar {
  const n = num(args, 0, ctx);
  if (isErr(n)) return n;
  const s = num(args, 1, ctx);
  if (isErr(s)) return s;
  if (s === 0) return dir === 'floor' ? ERR.div0 : 0;
  if (n !== 0 && Math.sign(n) !== Math.sign(s)) return ERR.num;
  const q = n / s;
  return (dir === 'ceil' ? Math.ceil(q) : Math.floor(q)) * s;
}

function mathStep(args: EvalValue[], ctx: FnContext, dir: 'ceil' | 'floor'): Scalar {
  const n = num(args, 0, ctx);
  if (isErr(n)) return n;
  const sig = optNum(args, 1, ctx, 1);
  if (isErr(sig)) return sig;
  const mode = optNum(args, 2, ctx, 0);
  if (isErr(mode)) return mode;
  if (sig === 0) return 0;
  const s = Math.abs(sig);
  const q = n / s;
  // `mode` non-zero rounds negatives away from zero instead of toward it.
  const awayFromZero = n < 0 && mode !== 0;
  const up = dir === 'ceil' ? !awayFromZero : awayFromZero;
  return (up ? Math.ceil(q) : Math.floor(q)) * s;
}

function intFold(
  args: EvalValue[],
  ctx: FnContext,
  step: (a: number, b: number) => number,
  seed: number,
): Scalar {
  const ns = collectNumbers(args, ctx);
  if (isErr(ns)) return ns;
  let acc = seed;
  for (const raw of ns) {
    const v = trunc(raw);
    if (v < 0) return ERR.num;
    acc = step(acc, v);
  }
  return acc;
}

const gcd2 = (a: number, b: number): number => {
  let x = Math.abs(a);
  let y = Math.abs(b);
  while (y) {
    const t = y;
    y = x % y;
    x = t;
  }
  return x;
};

function subtotalOf(code: number, ns: number[]): Scalar {
  switch (code) {
    case 1:
      return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : ERR.div0;
    case 2:
    case 3:
      return ns.length;
    case 4:
      return ns.length ? Math.max(...ns) : 0;
    case 5:
      return ns.length ? Math.min(...ns) : 0;
    case 6:
      return ns.reduce((a, b) => a * b, 1);
    case 9:
      return ns.reduce((a, b) => a + b, 0);
    case 7:
    case 8:
    case 10:
    case 11:
      return varianceLike(code, ns);
    default:
      return ERR.value;
  }
}

function varianceLike(code: number, ns: number[]): Scalar {
  const n = ns.length;
  if (n === 0) return ERR.div0;
  const mean = ns.reduce((a, b) => a + b, 0) / n;
  const ss = ns.reduce((a, b) => a + (b - mean) * (b - mean), 0);
  const sample = code === 7 || code === 10;
  if (sample && n < 2) return ERR.div0;
  const v = ss / (sample ? n - 1 : n);
  return code === 7 || code === 8 ? Math.sqrt(v) : v;
}
