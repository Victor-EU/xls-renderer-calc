import { describe, expect, it } from 'vitest';
import { MAX_COLS, MAX_ROWS } from './a1.js';
import { translate, unparse } from './ast.js';
import { isUnsupported, ParseError } from './errors.js';
import { parseFormula } from './parser.js';

const round = (f: string) => unparse(parseFormula(f));

describe('precedence', () => {
  it('binds unary minus tighter than ^ (Excel: -2^2 = 4)', () => {
    expect(round('-2^2')).toBe('-2^2');
    const n = parseFormula('-2^2');
    // structure, not text: the negation must be the *base* of the power
    expect(n.k).toBe('bin');
    expect(n.k === 'bin' && n.op).toBe('^');
    expect(n.k === 'bin' && n.l.k).toBe('un');
  });

  it('makes ^ left-associative (Excel: 2^3^2 = 64)', () => {
    const n = parseFormula('2^3^2');
    expect(n.k === 'bin' && n.l.k).toBe('bin'); // (2^3)^2, not 2^(3^2)
  });

  it('orders * / above + -', () => {
    const n = parseFormula('1+2*3');
    expect(n.k === 'bin' && n.op).toBe('+');
    expect(n.k === 'bin' && n.r.k === 'bin' && n.r.op).toBe('*');
  });

  it('puts & between arithmetic and comparison', () => {
    const n = parseFormula('1+2&"x"="3x"');
    expect(n.k === 'bin' && n.op).toBe('=');
    expect(n.k === 'bin' && n.l.k === 'bin' && n.l.op).toBe('&');
  });

  it('applies postfix % after unary minus', () => {
    expect(round('-2%')).toBe('-2%');
    expect(round('50%%')).toBe('50%%');
  });

  it('allows a unary operator on the right of an infix one', () => {
    expect(round('2^-1')).toBe('2^-1');
    expect(round('1--1')).toBe('1--1');
  });
});

describe('references', () => {
  it('parses relative, absolute and mixed anchors', () => {
    expect(round('$B$4')).toBe('$B$4');
    expect(round('B$4')).toBe('B$4');
    expect(round('$B4')).toBe('$B4');
  });

  it('parses ranges and normalises reversed corners', () => {
    expect(round('B4:E4')).toBe('B4:E4');
    expect(round('E4:B2')).toBe('B2:E4');
  });

  it('carries the $ with the coordinate when it reverses a range', () => {
    // The anchor belongs to the column or row, not to the corner it was written
    // in. Swapping the numbers alone left each `$` behind, which names the same
    // rectangle today and different cells the moment a shared formula moves it:
    // one row and column over, `$B2:A$1` went to `$A2:C$2` instead of `B$1:$B3`.
    expect(round('$B2:A$1')).toBe('A$1:$B2');
    const moved = unparse(translate(parseFormula('$B2:A$1'), 1, 1, MAX_ROWS, MAX_COLS));
    expect(moved).toBe('B$1:$B3');
  });

  it('parses quoted cross-sheet references', () => {
    expect(round("'Income Statement'!F4")).toBe("'Income Statement'!F4");
    expect(round('Sheet1!A1:B2')).toBe('Sheet1!A1:B2');
  });

  it('parses whole-column and whole-row references', () => {
    expect(round('A:A')).toBe('A:A');
    expect(round('1:1')).toBe('1:1');
    expect(round('SUM(B:C)')).toBe('SUM(B:C)');
  });

  it('treats a name-shaped token that is a legal address as a reference', () => {
    expect(parseFormula('LOG10').k).toBe('ref'); // column LOG, row 10 — Excel agrees
    expect(parseFormula('LOG10(2)').k).toBe('fn'); // …unless it is called
  });

  it('treats non-address identifiers as defined names', () => {
    expect(parseFormula('TaxRate').k).toBe('name');
    expect(parseFormula('XFE1').k).toBe('name'); // past the last column
    expect(parseFormula('A1048577').k).toBe('name'); // past the last row
  });

  it('parses the intersection operator', () => {
    const n = parseFormula('SUM(A1:C1 B1:B3)');
    expect(n.k === 'fn' && n.args[0]!.k).toBe('isect');
  });

  it('parses a union inside parentheses', () => {
    const n = parseFormula('SUM((A1:A2,C1:C2))');
    expect(n.k === 'fn' && n.args[0]!.k).toBe('union');
  });
});

describe('literals', () => {
  it('parses numbers in every written form', () => {
    expect(round('1')).toBe('1');
    expect(round('.5')).toBe('0.5');
    expect(round('1e-3')).toBe('0.001');
    expect(round('1E3')).toBe('1000');
  });

  it('unescapes doubled quotes in strings', () => {
    const n = parseFormula('"a""b"');
    expect(n.k === 'str' && n.v).toBe('a"b');
  });

  it('parses booleans and error literals', () => {
    expect(parseFormula('TRUE').k).toBe('bool');
    expect(parseFormula('#DIV/0!').k).toBe('err');
    expect(parseFormula('#N/A').k).toBe('err');
  });

  it('parses array constants', () => {
    expect(round('{1,2;3,4}')).toBe('{1,2;3,4}');
    expect(round('{-1,"a";TRUE,#N/A}')).toBe('{-1,"a";TRUE,#N/A}');
  });
});

describe('functions', () => {
  it('strips the _xlfn namespace Excel writes into files', () => {
    const n = parseFormula('_xlfn.XLOOKUP(1,A:A,B:B)');
    expect(n.k === 'fn' && n.name).toBe('XLOOKUP');
  });

  it('keeps omitted arguments positional', () => {
    const n = parseFormula('IF(A1,,2)');
    expect(n.k === 'fn' && n.args.length).toBe(3);
  });

  it('accepts a zero-argument call', () => {
    expect(round('TODAY()')).toBe('TODAY()');
  });
});

describe('refusals — these must never become values', () => {
  const rejects = [
    ['[1]Sheet1!A1', 'external workbook'],
    ['Table1[Revenue]', 'structured table'],
    ['Sheet1:Sheet3!A1', '3D reference'],
    ['{=SUM(A1:A3)}', 'legacy array formula'],
  ] as const;

  for (const [formula, what] of rejects) {
    it(`refuses ${what}`, () => {
      try {
        parseFormula(formula);
        throw new Error(`expected ${formula} to be refused`);
      } catch (e) {
        expect(isUnsupported(e)).toBe(true);
      }
    });
  }

  it('reports a position for malformed text', () => {
    try {
      parseFormula('1+');
      throw new Error('expected a parse error');
    } catch (e) {
      expect(e).toBeInstanceOf(ParseError);
      expect((e as ParseError).annotate()).toContain('^');
    }
  });
});

/**
 * `unparse` is the parser's inverse, and this is the property that matters:
 * re-parsing its output must give back the same tree.
 *
 * It was silently false. The grammar has no parenthesis node, so `(1-A1)` and
 * `1-A1` parse identically and the brackets survive only in how the subtree
 * hangs off its parent — which `unparse` used to discard. Nothing noticed,
 * because `unparse` was mostly read by humans in tooltips. Then a shared
 * formula `O19*(1-$E$17)` was reconstructed for its siblings as `O19*1-$E$17`,
 * and a real workbook rendered numbers two and a half times too large with no
 * error anywhere. Text equality would not have caught it: the wrong output is
 * perfectly stable under further round-trips. Only the tree tells the truth.
 */
describe('unparse is the parser inverse', () => {
  const CASES = [
    'O19*(1-$E$17)',
    'A1*(B1+C1)/(D1-E1)',
    '(1+2)*3',
    '1-(2-3)',
    '2/(3*4)',
    '-(2^2)',
    '(1+2)%',
    '(A1=B1)*2',
    'A1&(B1&C1)',
    '2^(3^2)',
    '(A1+B1)^2',
    'SUM((1+2)*3,-(A1-B1))',
    'IF((A1>0)*(B1>0),(A1+B1)/2,0)',
    '-2^2',
    '2^3^2',
    '1+2*3',
    "('My Sheet'!A1+2)*3",
    '{1,2;3,4}',
    '(A1:A5 B1:D1)',
    'SUM(A1:A5,(B1+C1))',
  ];

  for (const f of CASES) {
    it(`round-trips ${f}`, () => {
      const once = parseFormula(f);
      expect(parseFormula(unparse(once))).toEqual(once);
    });
  }

  it('restores the brackets a shared formula depends on', () => {
    // The exact regression: translate one column right, then read it back.
    const shifted = translate(parseFormula('O19*(1-$E$17)'), 0, 1, MAX_ROWS, MAX_COLS);
    expect(unparse(shifted)).toBe('P19*(1-$E$17)');
  });

  it('adds no brackets that are not needed', () => {
    // The output is read by people. Defensive parentheses everywhere would be
    // correct and unusable.
    expect(unparse(parseFormula('1+2*3-4/5'))).toBe('1+2*3-4/5');
    expect(unparse(parseFormula('SUM(A1:A5)+1'))).toBe('SUM(A1:A5)+1');
  });
});
