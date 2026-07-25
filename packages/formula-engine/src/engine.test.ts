import { describe, expect, it } from 'vitest';
import { ERR, isErr, type Scalar } from './values.js';
import { Workbook } from './workbook.js';

/**
 * A tiny fluent harness: `calc('=1+1')` builds a one-cell workbook, evaluates it
 * and returns the value. `sheet()` builds a grid when a formula needs context.
 */
function calc(formula: string, grid: Record<string, Scalar | string> = {}): Scalar {
  const wb = new Workbook({ now: 45000 });
  wb.addSheet('S');
  for (const [addr, v] of Object.entries(grid)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
    const col = m[1]!.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
    const row = Number(m[2]);
    if (typeof v === 'string' && v.startsWith('=')) wb.setFormula(0, row, col, v);
    else wb.setValue(0, row, col, v as Scalar);
  }
  wb.setFormula(0, 100, 26, formula);
  wb.evaluateAll();
  return wb.record(0, 100, 26)!.value;
}

const provenance = (formula: string, grid: Record<string, Scalar | string> = {}): string => {
  const wb = new Workbook({ now: 45000 });
  wb.addSheet('S');
  for (const [addr, v] of Object.entries(grid)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
    const col = m[1]!.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
    if (typeof v === 'string' && v.startsWith('=')) wb.setFormula(0, Number(m[2]), col, v);
    else wb.setValue(0, Number(m[2]), col, v as Scalar);
  }
  wb.setFormula(0, 100, 26, formula);
  wb.evaluateAll();
  return wb.record(0, 100, 26)!.provenance;
};

const near = (v: Scalar, expected: number): void => {
  expect(typeof v).toBe('number');
  expect(Math.abs((v as number) - expected)).toBeLessThan(1e-9 * Math.max(1, Math.abs(expected)));
};

describe('arithmetic and precedence', () => {
  it('computes the two precedence traps', () => {
    expect(calc('=-2^2')).toBe(4);
    expect(calc('=2^3^2')).toBe(64);
  });

  it('divides by zero into a value, not a failure', () => {
    expect(calc('=1/0')).toBe(ERR.div0);
    expect(provenance('=1/0')).toBe('computed');
  });

  it('applies percent as a postfix operator', () => {
    expect(calc('=50%')).toBe(0.5);
    expect(calc('=200*5%')).toBe(10);
  });
});

describe('coercion', () => {
  it('coerces numeric text in arithmetic but not in comparison', () => {
    expect(calc('="5"+1')).toBe(6);
    expect(calc('="5"=5')).toBe(false);
  });

  it('coerces booleans in arithmetic but not in comparison', () => {
    expect(calc('=TRUE+1')).toBe(2);
    expect(calc('=TRUE=1')).toBe(false);
  });

  it('errors on non-numeric text, including the empty string', () => {
    expect(calc('="a"+1')).toBe(ERR.value);
    expect(calc('=""+1')).toBe(ERR.value);
  });

  it('ranks types for cross-type comparison', () => {
    expect(calc('="a">1')).toBe(true); // text outranks number
    expect(calc('=TRUE>"z"')).toBe(true); // boolean outranks text
  });

  it('compares text case-insensitively', () => {
    expect(calc('="a"="A"')).toBe(true);
    expect(calc('=EXACT("a","A")')).toBe(false); // …except EXACT
  });
});

describe('the empty-cell rule', () => {
  it('treats an empty cell as 0 in arithmetic and as both 0 and "" in comparison', () => {
    expect(calc('=A1+1', {})).toBe(1);
    expect(calc('=A1=0', {})).toBe(true);
    expect(calc('=A1=""', {})).toBe(true);
    expect(calc('=A1=" "', {})).toBe(false);
  });

  it("keeps the sample model's guard clean instead of dividing by zero", () => {
    // IF(G9=0,"",F9/G9-1) over a column with no prior-year data.
    expect(calc('=IF(G9=0,"",F9/G9-1)', { F9: 100 })).toBe('');
  });

  it('distinguishes an empty cell from an empty string', () => {
    expect(calc('=ISBLANK(A1)', {})).toBe(true);
    expect(calc('=ISBLANK(A1)', { A1: '' })).toBe(false);
    expect(calc('=COUNT(A1:A2)', { A1: 1 })).toBe(1);
  });
});

describe('aggregates and the direct-vs-range rule', () => {
  it('ignores text and booleans inside a range but coerces them as arguments', () => {
    expect(calc('=SUM(A1:A3)', { A1: 1, A2: '5', A3: true })).toBe(1);
    expect(calc('=SUM("5")')).toBe(5);
    expect(calc('=SUM(TRUE)')).toBe(1);
    expect(calc('=SUM("a")')).toBe(ERR.value);
  });

  it('propagates an error out of a range rather than skipping it', () => {
    expect(calc('=SUM(A1:A2)', { A1: 1, A2: '=1/0' })).toBe(ERR.div0);
  });

  it('returns 0 for MIN/MAX over nothing but #DIV/0! for AVERAGE', () => {
    expect(calc('=MAX(A1:A3)', {})).toBe(0);
    expect(calc('=AVERAGE(A1:A3)', {})).toBe(ERR.div0);
  });
});

describe('laziness', () => {
  it('does not evaluate the branch IF did not take', () => {
    expect(calc('=IF(TRUE,1,1/0)')).toBe(1);
  });

  it('lets IFERROR trap a computed error', () => {
    expect(calc('=IFERROR(1/0,"fallback")')).toBe('fallback');
    expect(calc('=IFNA(1/0,"fallback")')).toBe(ERR.div0); // IFNA traps only #N/A
  });
});

describe('rounding', () => {
  it('rounds the decimal value, not the binary double', () => {
    expect(calc('=ROUND(2.675,2)')).toBe(2.68);
    expect(calc('=ROUND(-2.675,2)')).toBe(-2.68);
    expect(calc('=ROUND(1.005,2)')).toBe(1.01);
  });

  it('rounds half away from zero', () => {
    expect(calc('=ROUND(2.5,0)')).toBe(3);
    expect(calc('=ROUND(-2.5,0)')).toBe(-3);
  });

  it('rounds to the left of the decimal point', () => {
    expect(calc('=ROUND(1234.5,-2)')).toBe(1200);
    expect(calc('=ROUNDUP(1234.5,-2)')).toBe(1300);
  });

  it('floors INT rather than truncating it', () => {
    expect(calc('=INT(-2.5)')).toBe(-3);
    expect(calc('=TRUNC(-2.5)')).toBe(-2);
  });

  it('gives MOD the sign of the divisor', () => {
    expect(calc('=MOD(-3,2)')).toBe(1);
    expect(calc('=MOD(3,-2)')).toBe(-1);
  });
});

describe('references', () => {
  it('reads cross-sheet references', () => {
    const wb = new Workbook();
    wb.addSheet('Income Statement');
    wb.addSheet('Summary');
    wb.setValue(0, 4, 6, 1827.6);
    wb.setFormula(1, 1, 1, "='Income Statement'!F4*2");
    wb.evaluateAll();
    near(wb.record(1, 1, 1)!.value, 3655.2);
  });

  it('sums a range and a whole column identically', () => {
    const grid = { A1: 1, A2: 2, A3: 3 };
    expect(calc('=SUM(A1:A3)', grid)).toBe(6);
    expect(calc('=SUM(A:A)', grid)).toBe(6);
  });

  it('intersects two ranges with the space operator', () => {
    expect(calc('=SUM(A1:C1 B1:B3)', { A1: 1, B1: 2, C1: 3, B2: 4 })).toBe(2);
  });
});

describe('lookup', () => {
  const table = { A1: 1, B1: 'one', A2: 5, B2: 'five', A3: 9, B3: 'nine' };

  it('does exact and approximate VLOOKUP', () => {
    expect(calc('=VLOOKUP(5,A1:B3,2,FALSE)', table)).toBe('five');
    expect(calc('=VLOOKUP(7,A1:B3,2,TRUE)', table)).toBe('five'); // largest <= 7
    expect(calc('=VLOOKUP(0,A1:B3,2,TRUE)', table)).toBe(ERR.na);
  });

  it('refuses approximate match over unsorted data instead of guessing', () => {
    expect(provenance('=VLOOKUP(7,A1:B3,2,TRUE)', { A1: 9, B1: 'nine', A2: 1, B2: 'one' })).toBe(
      'unsupported',
    );
  });

  it('does INDEX/MATCH', () => {
    expect(calc('=INDEX(B1:B3,MATCH(9,A1:A3,0))', table)).toBe('nine');
  });
});

describe('conditional aggregates', () => {
  const grid = { A1: 'North', B1: 10, A2: 'South', B2: 20, A3: 'Northeast', B3: 30 };

  it('reads comparison criteria', () => {
    expect(calc('=SUMIF(B1:B3,">15")', grid)).toBe(50);
    expect(calc('=COUNTIF(B1:B3,"<=20")', grid)).toBe(2);
  });

  it('reads wildcard criteria', () => {
    expect(calc('=SUMIF(A1:A3,"North*",B1:B3)', grid)).toBe(40);
    expect(calc('=COUNTIF(A1:A3,"North")', grid)).toBe(1); // exact, no implicit wildcard
  });

  it('reads multi-criteria form', () => {
    expect(calc('=SUMIFS(B1:B3,A1:A3,"North*",B1:B3,">15")', grid)).toBe(30);
  });
});

describe('dates', () => {
  it('reproduces the 1900 leap-year bug', () => {
    expect(calc('=DATE(1900,1,1)')).toBe(1);
    expect(calc('=DATE(1900,3,1)')).toBe(61); // 60 is the phantom 29 Feb 1900
    expect(calc('=DATE(2026,7,25)')).toBe(46228);
  });

  it('extracts date parts', () => {
    expect(calc('=YEAR(46228)')).toBe(2026);
    expect(calc('=MONTH(46228)')).toBe(7);
    expect(calc('=DAY(46228)')).toBe(25);
  });

  it('does month arithmetic', () => {
    expect(calc('=EOMONTH(DATE(2026,1,31),1)')).toBe(calc('=DATE(2026,2,28)'));
    expect(calc('=EDATE(DATE(2026,1,31),1)')).toBe(calc('=DATE(2026,2,28)'));
  });
});

describe('financial', () => {
  it('computes an annuity payment', () => {
    near(calc('=PMT(0.05/12,360,-300000)'), 1610.4648690364193);
  });

  it('computes NPV with the end-of-period convention', () => {
    near(calc('=NPV(0.1,A1:A3)', { A1: 100, A2: 100, A3: 100 }), 248.68519909972893);
  });

  it('computes IRR', () => {
    near(calc('=IRR(A1:A4)', { A1: -1000, A2: 400, A3: 400, A4: 400 }), 0.09701025740327285);
  });

  it('splits a payment into interest and principal', () => {
    const i = calc('=IPMT(0.05,1,10,-1000)') as number;
    const p = calc('=PPMT(0.05,1,10,-1000)') as number;
    near(i, 50);
    near(i + p, calc('=PMT(0.05,10,-1000)') as number);
  });
});

describe('the honesty layer', () => {
  it('marks an unknown function unsupported rather than guessing a value', () => {
    expect(provenance('=NOTAFUNCTION(1)')).toBe('unsupported');
  });

  it('refuses INDIRECT and OFFSET because they break the dependency graph', () => {
    expect(provenance('=INDIRECT("A1")')).toBe('unsupported');
    expect(provenance('=OFFSET(A1,1,1)')).toBe('unsupported');
  });

  it('propagates unsupported to every dependent — a subtotal is never silently short', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setValue(0, 1, 1, 10);
    wb.setFormula(0, 2, 1, '=INDIRECT("A1")');
    wb.setFormula(0, 3, 1, '=SUM(A1:A2)');
    wb.evaluateAll();
    expect(wb.record(0, 3, 1)!.provenance).toBe('unsupported');
    expect(wb.record(0, 3, 1)!.value).not.toBe(10);
  });

  it('detects a cycle without iterating it', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setFormula(0, 1, 1, '=B1+1');
    wb.setFormula(0, 1, 2, '=A1+1');
    const report = wb.evaluateAll();
    expect(report.stats.circular).toBe(2);
    expect(wb.record(0, 1, 1)!.provenance).toBe('circular');
  });

  it('detects a self-reference', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setFormula(0, 1, 1, '=A1+1');
    expect(wb.evaluateAll().stats.circular).toBe(1);
  });

  it('marks volatile cells so the cached-value diff can exclude them', () => {
    expect(provenance('=TODAY()')).toBe('volatile');
  });

  it('reports gaps with a reason and a sample address', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setFormula(0, 1, 1, '=XLOOKUP2(1)');
    const report = wb.evaluateAll();
    expect(report.gaps[0]!.subject).toBe('XLOOKUP2');
    expect(report.gaps[0]!.sample).toBe('S!A1');
  });
});

describe('computed-versus-cached diff', () => {
  it('flags a stated total that disagrees with the sum of its parts', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setValue(0, 1, 1, 100);
    wb.setValue(0, 2, 1, 200);
    wb.setFormula(0, 3, 1, '=SUM(A1:A2)', 500); // the file claims 500
    const report = wb.evaluateAll();
    expect(report.stats.mismatched).toBe(1);
    expect(report.mismatches[0]!.computed).toBe(300);
    expect(report.mismatches[0]!.cached).toBe(500);
  });

  it('accepts LibreOffice-truncated precision as agreement', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setValue(0, 1, 1, 1);
    wb.setValue(0, 2, 1, 3);
    wb.setFormula(0, 3, 1, '=A1/A2', 0.333333333333333); // 15 digits, as LO writes
    expect(wb.evaluateAll().stats.mismatched).toBe(0);
  });
});

describe('error values are values', () => {
  it('keeps a computed Excel error as a result, not a gap', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setFormula(0, 1, 1, '=1/0');
    const report = wb.evaluateAll();
    expect(report.stats.unsupported).toBe(0);
    expect(report.stats.computed).toBe(1);
    expect(isErr(wb.record(0, 1, 1)!.value)).toBe(true);
  });
});
