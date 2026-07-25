import { describe, expect, it } from 'vitest';
import { Workbook } from './workbook.js';

/**
 * Performance guards, by formula *shape* rather than by formula count.
 *
 * The numbers that matter to a preview are the first row's: a generated model is
 * mostly local arithmetic, and 25,000 such formulas evaluate in well under a
 * tenth of a second.
 *
 * The other two shapes are quadratic by construction — a running total reads
 * every earlier row, a per-row full-table lookup scans the whole table — so they
 * cost O(rows²) cell reads no matter who evaluates them, Excel included. They
 * are here to keep that cost from regressing further, not because it can be
 * removed. The thresholds are deliberately loose so the suite does not go red on
 * a slower machine; a real regression is an order of magnitude, not 20%.
 */
function build(rows: number, shapes: string[]): Workbook {
  const wb = new Workbook();
  wb.addSheet('Model');
  for (let r = 1; r <= rows; r++) {
    wb.setValue(0, r, 1, r * 1.5);
    shapes.forEach((tpl, i) => {
      wb.setFormula(0, r, 2 + i, tpl.replace(/#/g, String(r)).replace(/@/g, String(rows)));
    });
  }
  return wb;
}

const time = (wb: Workbook): { ms: number; formulas: number } => {
  const t = performance.now();
  const report = wb.evaluateAll();
  return { ms: Math.round(performance.now() - t), formulas: report.stats.formulas };
};

describe('scale', () => {
  it('evaluates 25,000 local-arithmetic formulas quickly', () => {
    const r = time(build(5000, ['=A#*1.1', '=IF(B#>100,B#*0.2,0)', '=ROUND(C#/#,2)', '=IFERROR(C#/B#,0)', '=MAX(0,E#-1)']));
    console.log(`  local arithmetic    ${r.formulas} formulas  ${r.ms} ms`);
    expect(r.formulas).toBe(25000);
    expect(r.ms).toBeLessThan(2000);
  }, 120000);

  it('survives a running total over 5,000 rows without exhausting memory', () => {
    const r = time(build(5000, ['=SUM($A$1:A#)']));
    console.log(`  running total       ${r.formulas} formulas  ${r.ms} ms  (12.5M cell reads)`);
    expect(r.ms).toBeLessThan(10000);
  }, 120000);

  it('does a full-table lookup per row without copying the table each time', () => {
    const r = time(build(2000, ['=VLOOKUP(A#,$A$1:$A$@,1,FALSE)']));
    console.log(`  full-table VLOOKUP  ${r.formulas} formulas  ${r.ms} ms  (4M comparisons)`);
    expect(r.ms).toBeLessThan(5000);
  }, 120000);

  it('keeps a 5,000-row dependency chain off the call stack', () => {
    // Row-to-row growth is as deep as the model is long. A recursive evaluator
    // overflows here; the iterative DFS does not.
    const wb = new Workbook();
    wb.addSheet('Model');
    wb.setValue(0, 1, 1, 100);
    for (let r = 2; r <= 5000; r++) wb.setFormula(0, r, 1, `=A${r - 1}*1.01`);
    const report = wb.evaluateAll();
    expect(report.stats.computed).toBe(4999);
    expect(report.stats.unsupported).toBe(0);
    // Iterated multiplication and pow() differ in the last few bits at 1e23, so
    // compare relatively — the same rule the oracle comparator uses.
    const got = wb.record(0, 5000, 1)!.value as number;
    const want = 100 * 1.01 ** 4999;
    expect(Math.abs(got - want) / want).toBeLessThan(1e-12);
  }, 120000);
});
