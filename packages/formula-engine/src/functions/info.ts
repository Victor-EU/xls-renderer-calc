import { scalarOf } from '../interpreter.js';
import { ERR, isErr, isRef, type Scalar } from '../values.js';
import { fn, refuse } from './registry.js';
import { num, trunc } from './util.js';

export function registerInfo(): void {
  const predicate = (name: string, test: (v: Scalar) => boolean): void =>
    fn(name, 1, 1, (args, ctx) => test(scalarOf(args[0]!, ctx)));

  predicate('ISBLANK', (v) => v === null);
  predicate('ISNUMBER', (v) => typeof v === 'number');
  predicate('ISTEXT', (v) => typeof v === 'string');
  predicate('ISNONTEXT', (v) => typeof v !== 'string');
  predicate('ISLOGICAL', (v) => typeof v === 'boolean');
  predicate('ISERROR', (v) => isErr(v));
  predicate('ISERR', (v) => isErr(v) && v.kind !== '#N/A');
  predicate('ISNA', (v) => isErr(v) && v.kind === '#N/A');

  fn('ISREF', 1, 1, (args) => isRef(args[0]!));
  fn('NA', 0, 0, () => ERR.na);

  fn('ISEVEN', 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    return isErr(n) ? n : trunc(n) % 2 === 0;
  });
  fn('ISODD', 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    return isErr(n) ? n : Math.abs(trunc(n) % 2) === 1;
  });

  /** N and T are the two legacy coercion helpers, both still common in models. */
  fn('N', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    return 0; // text and blank both become 0 — N never errors on text
  });

  fn('T', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    return typeof v === 'string' ? v : '';
  });

  fn('TYPE', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (typeof v === 'number' || v === null) return 1;
    if (typeof v === 'string') return 2;
    if (typeof v === 'boolean') return 4;
    return 16; // error
  });

  fn('ERROR.TYPE', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (!isErr(v)) return ERR.na;
    const order = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];
    const i = order.indexOf(v.kind);
    return i < 0 ? ERR.na : i + 1;
  });

  refuse('CELL', 'CELL() reports workbook and formatting state this engine does not model');
  refuse('INFO', 'INFO() reports environment state this engine does not model');
  refuse('ISFORMULA', 'ISFORMULA() needs cell metadata the evaluator does not carry');
}
