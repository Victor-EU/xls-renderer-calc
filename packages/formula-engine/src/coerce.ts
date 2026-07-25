/**
 * Type coercion — where Excel compatibility actually lives.
 *
 * Function *count* is a vanity metric. The behaviours below are what separate an
 * engine that agrees with Excel from one that quietly doesn't:
 *
 *   "5"+1   → 6        arithmetic coerces numeric text
 *   "5"=5   → FALSE    comparison does not: different types never compare equal
 *   TRUE+1  → 2        booleans coerce in arithmetic
 *   TRUE=1  → FALSE    …but not in comparison
 *   ""+1    → #VALUE!  an empty *string* is not an empty *cell*
 *   blank+1 → 1        an empty *cell* is 0 in arithmetic
 *   blank=0 → TRUE     …and equals both 0 and "" in comparison
 *
 * That last pair is not academic. A real generated model contains
 * `IF(G9=0,"",F9/G9-1)` over a column with no prior-year data: with the blank
 * rule, `G9=0` is TRUE and the cell is a clean blank; without it, the guard
 * fails and twelve cells render `#DIV/0!` that Excel renders empty.
 */

import { ERR, ExcelError, isErr, type Scalar } from './values.js';

/** Text Excel accepts as a number in arithmetic context. */
const NUMERIC_TEXT =
  /^[\s ]*([+-]?)[\s ]*\$?[\s ]*((?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?[\s ]*(%?)[\s ]*$/;

/**
 * Parse text the way Excel's implicit coercion does. Returns undefined when the
 * text is not numeric — callers turn that into `#VALUE!` rather than 0, because
 * a silent 0 is a wrong number and a `#VALUE!` is a visible one.
 */
export function textToNumber(text: string): number | undefined {
  const m = NUMERIC_TEXT.exec(text);
  if (!m) return undefined;
  const digits = m[2]!;
  if (digits === '' || digits === '.') return undefined;
  const n = Number(`${m[1] === '-' ? '-' : ''}${digits.replace(/,/g, '')}${m[3] ? `e${m[3]}` : ''}`);
  if (!Number.isFinite(n)) return undefined;
  return m[4] === '%' ? n / 100 : n;
}

/**
 * Excel's "General" rendering of a number *as text* — the conversion `&` and
 * most text functions use. 15 significant digits, trailing zeros dropped, and
 * scientific notation at the same thresholds Excel uses.
 */
export function numberToText(n: number): string {
  if (Number.isNaN(n)) return 'NaN';
  if (!Number.isFinite(n)) return n > 0 ? 'Infinity' : '-Infinity';
  if (n === 0) return '0';

  const abs = Math.abs(n);
  const exp = Math.floor(Math.log10(abs));

  // Thresholds measured against the oracle rather than assumed. Positional
  // rendering holds from 1E-14 up to 1E15 inclusive; 1E-15 and 1E16 are the
  // first values on each side to switch. See the `text` suite in
  // tools/oracle/suites.py, which brackets both boundaries.
  if (exp >= 16 || exp <= -15) {
    // 1.23456789012346E+17 — mantissa trimmed, exponent unpadded (Excel style)
    const [mantissa, e] = n.toExponential(14).split('e');
    const trimmed = mantissa!.includes('.') ? mantissa!.replace(/0+$/, '').replace(/\.$/, '') : mantissa!;
    const sign = e!.startsWith('-') ? '-' : '+';
    return `${trimmed}E${sign}${Math.abs(Number(e))}`;
  }

  // Round to 15 significant digits, then render positionally.
  // 15 significant digits at the smallest positional exponent needs 29 places.
  const decimals = Math.min(30, Math.max(0, 14 - exp));
  let s = n.toFixed(decimals);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s === '-0' ? '0' : s;
}

/** Arithmetic coercion. Errors pass through; non-numeric text becomes #VALUE!. */
export function toNumber(v: Scalar): number | ExcelError {
  if (typeof v === 'number') return v;
  if (v === null) return 0; // blank cell
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  const n = textToNumber(v);
  return n === undefined ? ERR.value : n;
}

/** Text coercion, as used by `&`, TEXT-family functions and criteria matching. */
export function toText(v: Scalar): string | ExcelError {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v;
  return numberToText(v);
}

/** Logical coercion, as used by IF/AND/OR. Text is only coercible if it spells a boolean. */
export function toBool(v: Scalar): boolean | ExcelError {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (isErr(v)) return v;
  const up = v.trim().toUpperCase();
  if (up === 'TRUE') return true;
  if (up === 'FALSE') return false;
  return ERR.value;
}

/**
 * Excel's comparison ordering.
 *
 * Returns -1 / 0 / 1, or an ExcelError if either side is one. Rules:
 *   - blank adapts to the other side (0 against a number, "" against text,
 *     FALSE against a boolean, equal against another blank);
 *   - across types the order is number < text < boolean, so `"5"=5` is FALSE
 *     and `TRUE>"z"` is TRUE;
 *   - text compares case-insensitively, as Excel does.
 */
export function compare(a: Scalar, b: Scalar): number | ExcelError {
  if (isErr(a)) return a;
  if (isErr(b)) return b;

  if (a === null && b === null) return 0;
  if (a === null) a = blankAs(b);
  if (b === null) b = blankAs(a);

  const ra = rankOf(a);
  const rb = rankOf(b);
  if (ra !== rb) return ra < rb ? -1 : 1;

  if (typeof a === 'number' && typeof b === 'number') return a < b ? -1 : a > b ? 1 : 0;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;

  const sa = String(a).toUpperCase();
  const sb = String(b).toUpperCase();
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

function blankAs(other: Scalar): Scalar {
  if (typeof other === 'string') return '';
  if (typeof other === 'boolean') return false;
  return 0;
}

function rankOf(v: Scalar): number {
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  return 3; // boolean
}

/** True when a scalar counts as a number for aggregation (COUNT, AVERAGE...). */
export const isNumeric = (v: Scalar): v is number => typeof v === 'number';
