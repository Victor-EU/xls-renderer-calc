/**
 * Financial functions.
 *
 * These are the ones a generated model is most likely to reach for, and the ones
 * where LibreOffice — our oracle — is least reliable as a stand-in for Excel
 * (§12.5 of the design). Every function here therefore also carries hand-checked
 * fixtures in financial.test.ts with values verified against Excel's documented
 * definitions, not just against LO.
 *
 * The annuity identity all of PV/FV/PMT/NPER/RATE share:
 *
 *   pv·(1+r)^n + pmt·(1 + r·type)·((1+r)^n − 1)/r + fv = 0
 */

import { type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix, collectNumbers, num, optNum, trunc } from './util.js';


const annuityFactor = (r: number, n: number, type: number): number =>
  r === 0 ? n : ((Math.pow(1 + r, n) - 1) / r) * (1 + r * type);

export function registerFinancial(): void {
  fn('PMT', 3, 5, (args, ctx) => {
    const a = annuityArgs(args, ctx, ['rate', 'nper', 'pv', 'fv', 'type']);
    if (isErr(a)) return a;
    const { rate: r, nper: n, pv, fv, type } = a;
    if (n === 0) return ERR.num;
    if (r === 0) return -(pv + fv) / n;
    return -(pv * Math.pow(1 + r, n) + fv) / annuityFactor(r, n, type);
  });

  fn('PV', 3, 5, (args, ctx) => {
    const a = annuityArgs(args, ctx, ['rate', 'nper', 'pmt', 'fv', 'type']);
    if (isErr(a)) return a;
    const { rate: r, nper: n, pmt, fv, type } = a;
    if (r === 0) return -(fv + pmt * n);
    return -(fv + pmt * annuityFactor(r, n, type)) / Math.pow(1 + r, n);
  });

  fn('FV', 3, 5, (args, ctx) => {
    const a = annuityArgs(args, ctx, ['rate', 'nper', 'pmt', 'pv', 'type']);
    if (isErr(a)) return a;
    const { rate: r, nper: n, pmt, pv, type } = a;
    if (r === 0) return -(pv + pmt * n);
    return -(pv * Math.pow(1 + r, n) + pmt * annuityFactor(r, n, type));
  });

  fn('NPER', 3, 5, (args, ctx) => {
    const a = annuityArgs(args, ctx, ['rate', 'pmt', 'pv', 'fv', 'type']);
    if (isErr(a)) return a;
    const { rate: r, pmt, pv, fv, type } = a;
    if (pmt === 0 && r === 0) return ERR.num;
    if (r === 0) return -(pv + fv) / pmt;
    const adj = pmt * (1 + r * type);
    const ratio = (adj - fv * r) / (pv * r + adj);
    if (!(ratio > 0)) return ERR.num;
    return Math.log(ratio) / Math.log(1 + r);
  });

  fn('RATE', 3, 6, (args, ctx) => {
    const nper = num(args, 0, ctx);
    if (isErr(nper)) return nper;
    const pmt = num(args, 1, ctx);
    if (isErr(pmt)) return pmt;
    const pv = num(args, 2, ctx);
    if (isErr(pv)) return pv;
    const fv = optNum(args, 3, ctx, 0);
    if (isErr(fv)) return fv;
    const type = optNum(args, 4, ctx, 0);
    if (isErr(type)) return type;
    const guess = optNum(args, 5, ctx, 0.1);
    if (isErr(guess)) return guess;

    const f = (r: number): number =>
      r === 0
        ? pv + pmt * nper + fv
        : pv * Math.pow(1 + r, nper) + pmt * annuityFactor(r, nper, type) + fv;
    const root = solve(f, guess);
    return root === undefined ? ERR.num : root;
  });

  fn('NPV', 2, -1, (args, ctx) => {
    const rate = num(args, 0, ctx);
    if (isErr(rate)) return rate;
    if (rate === -1) return ERR.div0;
    const flows = collectNumbers(args.slice(1), ctx);
    if (isErr(flows)) return flows;
    // NPV discounts the FIRST value by one period — it is an end-of-period
    // convention, which is why models pair it with a separate initial outlay.
    let total = 0;
    for (let i = 0; i < flows.length; i++) total += flows[i]! / Math.pow(1 + rate, i + 1);
    return total;
  });

  fn('IRR', 1, 2, (args, ctx) => {
    const flows = collectNumbers([args[0]!], ctx);
    if (isErr(flows)) return flows;
    const guess = optNum(args, 1, ctx, 0.1);
    if (isErr(guess)) return guess;
    if (!hasSignChange(flows)) return ERR.num;
    const npv = (r: number): number =>
      flows.reduce((acc, c, i) => acc + c / Math.pow(1 + r, i), 0);
    const root = solve(npv, guess);
    return root === undefined ? ERR.num : root;
  });

  fn('MIRR', 3, 3, (args, ctx) => {
    const flows = collectNumbers([args[0]!], ctx);
    if (isErr(flows)) return flows;
    const finance = num(args, 1, ctx);
    if (isErr(finance)) return finance;
    const reinvest = num(args, 2, ctx);
    if (isErr(reinvest)) return reinvest;
    const n = flows.length;
    if (n < 2) return ERR.div0;
    let negPv = 0;
    let posFv = 0;
    for (let i = 0; i < n; i++) {
      const c = flows[i]!;
      if (c < 0) negPv += c / Math.pow(1 + finance, i);
      else posFv += c * Math.pow(1 + reinvest, n - 1 - i);
    }
    if (negPv === 0 || posFv === 0) return ERR.div0;
    return Math.pow(-posFv / negPv, 1 / (n - 1)) - 1;
  });

  fn('XNPV', 3, 3, (args, ctx) => {
    const parts = datedFlows(args, ctx);
    if (isErr(parts)) return parts;
    const rate = num(args, 0, ctx);
    if (isErr(rate)) return rate;
    const { values, dates } = parts;
    const d0 = dates[0]!;
    let total = 0;
    for (let i = 0; i < values.length; i++) {
      total += values[i]! / Math.pow(1 + rate, (dates[i]! - d0) / 365);
    }
    return total;
  });

  fn('XIRR', 2, 3, (args, ctx) => {
    const parts = datedFlows([null as unknown as EvalValue, args[0]!, args[1]!], ctx);
    if (isErr(parts)) return parts;
    const guess = optNum(args, 2, ctx, 0.1);
    if (isErr(guess)) return guess;
    const { values, dates } = parts;
    if (!hasSignChange(values)) return ERR.num;
    const d0 = dates[0]!;
    const xnpv = (r: number): number =>
      values.reduce((acc, v, i) => acc + v / Math.pow(1 + r, (dates[i]! - d0) / 365), 0);
    const root = solve(xnpv, guess);
    return root === undefined ? ERR.num : root;
  });

  fn('IPMT', 4, 6, (args, ctx) => periodicSplit(args, ctx, 'interest'));
  fn('PPMT', 4, 6, (args, ctx) => periodicSplit(args, ctx, 'principal'));

  fn('CUMIPMT', 6, 6, (args, ctx) => cumulative(args, ctx, 'interest'));
  fn('CUMPRINC', 6, 6, (args, ctx) => cumulative(args, ctx, 'principal'));

  fn('SLN', 3, 3, (args, ctx) => {
    const cost = num(args, 0, ctx);
    if (isErr(cost)) return cost;
    const salvage = num(args, 1, ctx);
    if (isErr(salvage)) return salvage;
    const life = num(args, 2, ctx);
    if (isErr(life)) return life;
    return life === 0 ? ERR.div0 : (cost - salvage) / life;
  });

  fn('SYD', 4, 4, (args, ctx) => {
    const cost = num(args, 0, ctx);
    if (isErr(cost)) return cost;
    const salvage = num(args, 1, ctx);
    if (isErr(salvage)) return salvage;
    const life = num(args, 2, ctx);
    if (isErr(life)) return life;
    const per = num(args, 3, ctx);
    if (isErr(per)) return per;
    if (life <= 0 || per < 1 || per > life) return ERR.num;
    return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
  });

  fn('DDB', 4, 5, (args, ctx) => {
    const cost = num(args, 0, ctx);
    if (isErr(cost)) return cost;
    const salvage = num(args, 1, ctx);
    if (isErr(salvage)) return salvage;
    const life = num(args, 2, ctx);
    if (isErr(life)) return life;
    const period = num(args, 3, ctx);
    if (isErr(period)) return period;
    const factor = optNum(args, 4, ctx, 2);
    if (isErr(factor)) return factor;
    if (life <= 0 || period < 1) return ERR.num;

    let book = cost;
    let charge = 0;
    for (let p = 1; p <= Math.ceil(period); p++) {
      charge = Math.min((book * factor) / life, book - salvage);
      if (charge < 0) charge = 0;
      book -= charge;
    }
    return charge;
  });

  fn('EFFECT', 2, 2, (args, ctx) => {
    const nominal = num(args, 0, ctx);
    if (isErr(nominal)) return nominal;
    const periods = num(args, 1, ctx);
    if (isErr(periods)) return periods;
    const n = trunc(periods);
    if (nominal <= 0 || n < 1) return ERR.num;
    return Math.pow(1 + nominal / n, n) - 1;
  });

  fn('NOMINAL', 2, 2, (args, ctx) => {
    const effective = num(args, 0, ctx);
    if (isErr(effective)) return effective;
    const periods = num(args, 1, ctx);
    if (isErr(periods)) return periods;
    const n = trunc(periods);
    if (effective <= 0 || n < 1) return ERR.num;
    return (Math.pow(effective + 1, 1 / n) - 1) * n;
  });

  fn('RRI', 3, 3, (args, ctx) => {
    const nper = num(args, 0, ctx);
    if (isErr(nper)) return nper;
    const pv = num(args, 1, ctx);
    if (isErr(pv)) return pv;
    const fv = num(args, 2, ctx);
    if (isErr(fv)) return fv;
    if (nper <= 0) return ERR.num;
    // Growing from nothing to nothing is a rate of zero, not an error. Excel
    // special-cases only this one corner: `pv = 0` with a non-zero `fv` really
    // is #NUM!, and `fv = 0` with a non-zero `pv` needs no special case at all
    // because the formula already gives -1. Derived from 735 RRI calls in a
    // business plan saved by Excel itself, where every one of those three
    // shapes appears; LibreOffice returns #VALUE! here and is the outlier.
    if (pv === 0) return fv === 0 ? 0 : ERR.num;
    return Math.pow(fv / pv, 1 / nper) - 1;
  });

  fn('PDURATION', 3, 3, (args, ctx) => {
    const rate = num(args, 0, ctx);
    if (isErr(rate)) return rate;
    const pv = num(args, 1, ctx);
    if (isErr(pv)) return pv;
    const fv = num(args, 2, ctx);
    if (isErr(fv)) return fv;
    if (rate <= 0 || pv <= 0 || fv <= 0) return ERR.num;
    return (Math.log(fv) - Math.log(pv)) / Math.log(1 + rate);
  });

}

interface Annuity {
  rate: number;
  nper: number;
  pmt: number;
  pv: number;
  fv: number;
  type: number;
}

/** Read the positional annuity arguments by role, defaulting fv and type to 0. */
function annuityArgs(
  args: EvalValue[],
  ctx: FnContext,
  order: Array<keyof Annuity>,
): Annuity | ExcelError {
  const out: Annuity = { rate: 0, nper: 0, pmt: 0, pv: 0, fv: 0, type: 0 };
  for (let i = 0; i < order.length; i++) {
    const key = order[i]!;
    const optional = key === 'fv' || key === 'type';
    const v = optional ? optNum(args, i, ctx, 0) : num(args, i, ctx);
    if (isErr(v)) return v;
    out[key] = key === 'type' ? (v === 0 ? 0 : 1) : v;
  }
  return out;
}

function periodicSplit(args: EvalValue[], ctx: FnContext, want: 'interest' | 'principal'): Scalar {
  const rate = num(args, 0, ctx);
  if (isErr(rate)) return rate;
  const per = num(args, 1, ctx);
  if (isErr(per)) return per;
  const nper = num(args, 2, ctx);
  if (isErr(nper)) return nper;
  const pv = num(args, 3, ctx);
  if (isErr(pv)) return pv;
  const fv = optNum(args, 4, ctx, 0);
  if (isErr(fv)) return fv;
  const typeRaw = optNum(args, 5, ctx, 0);
  if (isErr(typeRaw)) return typeRaw;
  const type = typeRaw === 0 ? 0 : 1;
  if (per < 1 || per > nper) return ERR.num;

  const payment =
    rate === 0 ? -(pv + fv) / nper : -(pv * Math.pow(1 + rate, nper) + fv) / annuityFactor(rate, nper, type);

  // Balance at the start of the requested period.
  const balance =
    rate === 0
      ? pv + payment * (per - 1)
      : pv * Math.pow(1 + rate, per - 1) + payment * annuityFactor(rate, per - 1, type);

  let interest: number;
  if (type === 1 && per === 1) interest = 0;
  else interest = -balance * rate;

  return want === 'interest' ? interest : payment - interest;
}

function cumulative(args: EvalValue[], ctx: FnContext, want: 'interest' | 'principal'): Scalar {
  const rate = num(args, 0, ctx);
  if (isErr(rate)) return rate;
  const nper = num(args, 1, ctx);
  if (isErr(nper)) return nper;
  const pv = num(args, 2, ctx);
  if (isErr(pv)) return pv;
  if (rate <= 0 || nper <= 0 || pv <= 0) return ERR.num;
  const start = num(args, 3, ctx);
  if (isErr(start)) return start;
  const end = num(args, 4, ctx);
  if (isErr(end)) return end;
  if (start < 1 || end < start) return ERR.num;
  let total = 0;
  for (let p = trunc(start); p <= trunc(end); p++) {
    const part = periodicSplit(
      [args[0]!, p, args[1]!, args[2]!, 0, args[5]!],
      ctx,
      want,
    );
    if (isErr(part)) return part;
    if (typeof part !== 'number') return ERR.value;
    total += part;
  }
  return total;
}

function datedFlows(
  args: EvalValue[],
  ctx: FnContext,
): { values: number[]; dates: number[] } | ExcelError {
  const vm = asMatrix(args[1]!, ctx);
  if (isErr(vm)) return vm;
  const dm = asMatrix(args[2]!, ctx);
  if (isErr(dm)) return dm;
  if (vm.size !== dm.size) return ERR.num;

  const values: number[] = [];
  const dates: number[] = [];
  for (let i = 0; i < vm.size; i++) {
    const v = vm.data[i]!;
    const d = dm.data[i]!;
    if (isErr(v)) return v;
    if (isErr(d)) return d;
    if (typeof v !== 'number' || typeof d !== 'number') return ERR.value;
    values.push(v);
    dates.push(Math.floor(d));
  }
  if (values.length === 0) return ERR.num;
  if (dates.some((d) => d < dates[0]!)) return ERR.num; // Excel requires the first date to be earliest
  return { values, dates };
}

const hasSignChange = (ns: number[]): boolean =>
  ns.some((n) => n > 0) && ns.some((n) => n < 0);

/**
 * Newton's method with a bisection fallback.
 *
 * Excel's own IRR/RATE are iterative and stop at 1e-7 over at most 20 steps,
 * returning #NUM! when they miss. Matching that convergence exactly is not
 * possible from outside, so this solver aims tighter and then rounds — and the
 * oracle harness carries fixtures for the shapes where the two could disagree.
 */
function solve(f: (x: number) => number, guess: number): number | undefined {
  const bracket = findBracket(f, guess);
  if (bracket) return bisect(f, bracket[0], bracket[1]);

  // No sign change anywhere we looked — fall back to Newton, which can still
  // find a root the scan stepped over.
  let x = guess;
  for (let i = 0; i < 200; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) return undefined;
    if (y === 0) return x;
    const h = Math.max(1e-8, Math.abs(x) * 1e-8);
    const dy = (f(x + h) - f(x - h)) / (2 * h);
    if (!Number.isFinite(dy) || dy === 0) return undefined;
    const next = x - y / dy;
    if (!Number.isFinite(next)) return undefined;
    if (Math.abs(next - x) <= 1e-15 * Math.max(1, Math.abs(next))) return next;
    x = next;
  }
  return undefined;
}

/** Scan outward from the guess for a sign change, staying above rate = -100%. */
function findBracket(f: (x: number) => number, guess: number): [number, number] | undefined {
  const probes = [guess, 0, -0.5, 0.5, -0.9, 1, -0.99, 5, -0.999, 20, -0.9999, 100, 1000, 1e5];
  const seen: Array<{ x: number; y: number }> = [];
  for (const x of probes) {
    if (x <= -1) continue;
    const y = f(x);
    if (!Number.isFinite(y)) continue;
    if (y === 0) return [x, x];
    seen.push({ x, y });
  }
  seen.sort((a, b) => a.x - b.x);
  for (let i = 1; i < seen.length; i++) {
    if (seen[i - 1]!.y * seen[i]!.y < 0) return [seen[i - 1]!.x, seen[i]!.x];
  }
  return undefined;
}

/** Bisection to the limit of double precision — slower than Newton, exact. */
function bisect(f: (x: number) => number, a: number, b: number): number | undefined {
  if (a === b) return a;
  let lo = a;
  let hi = b;
  let flo = f(lo);
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (mid === lo || mid === hi) return mid; // adjacent doubles: as tight as it gets
    const fm = f(mid);
    if (!Number.isFinite(fm)) return undefined;
    if (fm === 0) return mid;
    if (flo * fm < 0) {
      hi = mid;
    } else {
      lo = mid;
      flo = fm;
    }
  }
  return (lo + hi) / 2;
}
