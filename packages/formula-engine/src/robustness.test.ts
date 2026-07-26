/**
 * One cell must never take the workbook with it.
 *
 * Every case below used to throw something that was neither `Unsupported` nor
 * `ParseError`, which meant it unwound `evaluateAll` and the *entire* file
 * rendered as nothing — the worst outcome available to a viewer whose promise is
 * that it will not mislead you. Each is an ordinary formula over an ordinary
 * sheet, not a fuzzing artefact.
 *
 * The tests come in pairs on purpose. One asserts the specific fix (the value
 * Excel gives), the other asserts the floor underneath it: whatever else breaks,
 * the neighbouring cells still compute.
 */

import { describe, expect, it } from 'vitest';
import { register } from './functions/registry.js';
import { ERR } from './values.js';
import { Workbook } from './workbook.js';

/** A sheet with a lookup table in D1:E3 and a numeric column in A. */
function fixture(rows = 3): Workbook {
  const wb = new Workbook({ now: 45000 });
  wb.addSheet('S');
  for (let r = 1; r <= rows; r++) wb.setValue(0, r, 1, r);
  for (let r = 1; r <= 3; r++) {
    wb.setValue(0, r, 4, r);
    wb.setValue(0, r, 5, 'xyz'[r - 1]!);
  }
  return wb;
}

/** Put `formula` in B1 and a control in B2, evaluate, report both. */
function withControl(wb: Workbook, formula: string): { value: unknown; control: unknown; gaps: string[] } {
  wb.setFormula(0, 1, 2, formula);
  wb.setFormula(0, 2, 2, '=1+1');
  const report = wb.evaluateAll();
  return {
    value: wb.record(0, 1, 2)!.value,
    control: wb.record(0, 2, 2)!.value,
    gaps: report.gaps.map((g) => g.code),
  };
}

describe('an omitted trailing argument', () => {
  // `IF(A1,,2)` made the parser produce a sentinel for "argument not written",
  // which is a real distinction worth keeping — but the sentinel is a Symbol,
  // and any function that asked for its *value* without checking first got one.
  // `=VLOOKUP(A1,T,2,)` threw `v.toUpperCase is not a function`.
  it('reads as blank rather than throwing', () => {
    const cases: Array<[string, unknown]> = [
      ['=VLOOKUP(1,D1:E3,2,)', 'x'],
      ['=CONCATENATE("a",,"b")', 'ab'],
      ['=SUBSTITUTE("abc","b",)', 'ac'],
      ['=DAYS360(1,2,)', 1],
      ['=FIXED(1.5,1,)', '1.5'],
    ];
    for (const [formula, expected] of cases) {
      const { value, control } = withControl(fixture(), formula);
      expect(value, formula).toEqual(expected);
      expect(control, `${formula} — the rest of the sheet`).toBe(2);
    }
  });

  it('is not a supplied fallback: XLOOKUP still answers #N/A', () => {
    // `XLOOKUP(k,a,b,,0)` writes the 4th argument as omitted. Treating the
    // sentinel as a value stored a Symbol in the cell, where it survived all the
    // way to `structuredClone` and rejected the Worker's whole load.
    const { value } = withControl(fixture(), '=XLOOKUP(99,D1:D3,E1:E3,,0)');
    expect(value).toBe(ERR.na);
  });
});

describe('an aggregate over a long column', () => {
  // `Math.min(...ns)` is one JS argument per element; past ~125,000 V8 throws
  // RangeError. A transaction export is longer than that, and `=MIN(A:A)` over
  // one is not an unusual formula.
  it('does not spread the range into an argument list', () => {
    const wb = fixture(200_000);
    wb.setFormula(0, 1, 2, '=MIN(A:A)');
    wb.setFormula(0, 2, 2, '=MAX(A:A)');
    wb.setFormula(0, 3, 2, '=SUBTOTAL(4,A:A)');
    wb.setFormula(0, 4, 2, '=MAXIFS(A:A,A:A,">0")');
    wb.evaluateAll();
    expect(wb.record(0, 1, 2)!.value).toBe(1);
    expect(wb.record(0, 2, 2)!.value).toBe(200_000);
    expect(wb.record(0, 3, 2)!.value).toBe(200_000);
    expect(wb.record(0, 4, 2)!.value).toBe(200_000);
  });
});

describe('a date function handed a serial off the calendar', () => {
  // Both walk one day at a time. An out-of-range serial — a mistyped cell, a
  // misread column — was tens of millions of iterations, each allocating a Date.
  // It read as a hang, which is the one failure a viewer cannot explain.
  it('refuses instead of counting to it', () => {
    const started = Date.now();
    const { value, control } = withControl(fixture(), '=NETWORKDAYS(1,50000000)');
    expect(value).toBe(ERR.num);
    expect(control).toBe(2);
    expect(Date.now() - started).toBeLessThan(1000);
  });

  it('refuses a day count that walks off it', () => {
    const { value } = withControl(fixture(), '=WORKDAY(1,1000000000)');
    expect(value).toBe(ERR.num);
  });
});

// A function that fails the way a real defect does — not an Excel error, not a
// refusal, just a thrown TypeError from somewhere nobody expected one.
register({
  name: 'BOOM',
  minArgs: 0,
  maxArgs: 0,
  call: () => {
    throw new TypeError('v.toUpperCase is not a function');
  },
});

describe('an unexpected failure', () => {
  it('is refused at the cell and reported as INTERNAL, not thrown', () => {
    // The specific crashes above are fixed; the floor underneath them is what
    // this asserts. A defect we have not met yet must still cost one cell, and
    // must be loud about it in `gaps` — an engine bug should never read as a
    // deliberate refusal.
    const wb = fixture();
    const { value, control, gaps } = withControl(wb, '=BOOM()');
    expect(value).toBe(null);
    expect(wb.record(0, 1, 2)!.provenance).toBe('unsupported');
    expect(wb.record(0, 1, 2)!.reason).toContain('TypeError');
    expect(gaps).toContain('INTERNAL');
    expect(control, 'the rest of the sheet still computes').toBe(2);
  });

  it('cannot fail the audit pass either', () => {
    // `tryEvaluate` runs speculative formulas on behalf of the hardcoded-cell
    // detector. It used to rethrow anything unexpected, so a defect reached
    // there would fail a load that had already succeeded.
    const wb = fixture();
    wb.evaluateAll();
    expect(wb.tryEvaluate(0, 1, 2, '=BOOM()')).toHaveProperty('unsupported');
  });
});
