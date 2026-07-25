import { toBool } from '../coerce.js';
import { scalarOf, type FnContext } from '../interpreter.js';
import { ERR, isErr, type EvalValue, type Scalar } from '../values.js';
import { fn, lazyFn } from './registry.js';
import { flatten, num } from './util.js';

export function registerLogical(): void {
  /**
   * IF must not evaluate the branch it does not take.
   *
   * `IF(A1=0,"",1/A1)` is the canonical guard, and an eager IF divides by zero
   * on exactly the input the guard exists to protect against — turning a clean
   * blank into `#DIV/0!`.
   */
  lazyFn('IF', 2, 3, (nodes, ctx, ev) => {
    const cond = toBool(scalarOf(ev(nodes[0]!, ctx), ctx));
    if (isErr(cond)) return cond;
    const branch = cond ? nodes[1] : nodes[2];
    if (!branch) return false; // omitted FALSE branch yields FALSE, not blank
    return ev(branch, ctx);
  });

  lazyFn('IFS', 2, -1, (nodes, ctx, ev) => {
    if (nodes.length % 2 !== 0) return ERR.na;
    for (let i = 0; i < nodes.length; i += 2) {
      const cond = toBool(scalarOf(ev(nodes[i]!, ctx), ctx));
      if (isErr(cond)) return cond;
      if (cond) return ev(nodes[i + 1]!, ctx);
    }
    return ERR.na;
  });

  /**
   * IFERROR is lazy so that an *unsupported* fallback cannot poison a cell whose
   * primary value computed cleanly — and vice versa.
   */
  lazyFn('IFERROR', 2, 2, (nodes, ctx, ev) => {
    const v = ev(nodes[0]!, ctx);
    const s = scalarOf(v, ctx);
    return isErr(s) ? ev(nodes[1]!, ctx) : v;
  });

  lazyFn('IFNA', 2, 2, (nodes, ctx, ev) => {
    const v = ev(nodes[0]!, ctx);
    const s = scalarOf(v, ctx);
    return isErr(s) && s.kind === '#N/A' ? ev(nodes[1]!, ctx) : v;
  });

  lazyFn('CHOOSE', 2, -1, (nodes, ctx, ev) => {
    const i = scalarOf(ev(nodes[0]!, ctx), ctx);
    const n = num([i], 0, ctx);
    if (isErr(n)) return n;
    const k = Math.trunc(n);
    if (k < 1 || k >= nodes.length) return ERR.value;
    return ev(nodes[k]!, ctx);
  });

  lazyFn('SWITCH', 3, -1, (nodes, ctx, ev) => {
    const subject = scalarOf(ev(nodes[0]!, ctx), ctx);
    if (isErr(subject)) return subject;
    let i = 1;
    for (; i + 1 < nodes.length; i += 2) {
      const candidate = scalarOf(ev(nodes[i]!, ctx), ctx);
      if (isErr(candidate)) return candidate;
      if (looseEqual(subject, candidate)) return ev(nodes[i + 1]!, ctx);
    }
    // A trailing odd argument is the default.
    return i < nodes.length ? ev(nodes[i]!, ctx) : ERR.na;
  });

  fn('AND', 1, -1, (args, ctx) => boolFold(args, ctx, 'and'));
  fn('OR', 1, -1, (args, ctx) => boolFold(args, ctx, 'or'));
  fn('XOR', 1, -1, (args, ctx) => boolFold(args, ctx, 'xor'));

  fn('NOT', 1, 1, (args, ctx) => {
    const b = toBool(scalarOf(args[0]!, ctx));
    return isErr(b) ? b : !b;
  });

  fn('TRUE', 0, 0, () => true);
  fn('FALSE', 0, 0, () => false);
}

function looseEqual(a: Scalar, b: Scalar): boolean {
  if (typeof a === 'string' && typeof b === 'string') return a.toUpperCase() === b.toUpperCase();
  return a === b;
}

/**
 * AND/OR/XOR see booleans and numbers; text and blanks inside a range are
 * invisible, while text written directly as an argument is an error. With no
 * logical values at all, Excel returns #VALUE! rather than a default.
 */
function boolFold(args: EvalValue[], ctx: FnContext, mode: 'and' | 'or' | 'xor'): Scalar {
  let seen = 0;
  let trues = 0;
  for (const { v, direct } of flatten(args, ctx)) {
    if (isErr(v)) return v;
    if (typeof v === 'boolean') {
      seen++;
      if (v) trues++;
    } else if (typeof v === 'number') {
      seen++;
      if (v !== 0) trues++;
    } else if (direct && v !== null) {
      const b = toBool(v);
      if (isErr(b)) return b;
      seen++;
      if (b) trues++;
    }
  }
  if (seen === 0) return ERR.value;
  if (mode === 'and') return trues === seen;
  if (mode === 'or') return trues > 0;
  return trues % 2 === 1;
}
