/**
 * Turning an overlay value into render-ready content.
 *
 * Number formatting is `numfmt` (MIT), the same library the original spike
 * proved out — it is what makes formula cells render as `$1,827.6`, `(168.9)`
 * in red and `59.0%` instead of raw floats. The spike also showed why this layer
 * matters: the drop-in canvas renderer it was measured against loses number
 * formats on formula cells specifically, which in a financial model is nearly
 * every meaningful number.
 *
 * The one rule this module adds: a cell the engine could not compute renders ⚠,
 * never a number. There is no code path here from `unsupported` to a value.
 */

import { format as numFormat, formatColor } from 'numfmt';
import { isErr, numberToText, serialToDate, type Scalar } from '@xlscalc/formula-engine';
import type { OverlayCell } from './bind.js';
import type { CellContentData, RichRun } from './layout.js';

export type { RichRun };

export type CellContent =
  | { kind: 'empty' }
  /** The engine declined to compute this cell; `reason` says why. */
  | { kind: 'unsupported'; reason: string }
  | { kind: 'error'; text: string }
  | { kind: 'text'; text: string; color?: string; numeric?: boolean }
  | { kind: 'rich'; runs: RichRun[] }
  | { kind: 'link'; text: string; href: string };

/** The eight colour names an Excel format code can name, e.g. `[Red]`. */
const NUMFMT_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00b0c0',
  green: '#008000',
  magenta: '#ff00ff',
  red: '#c0392b',
  white: '#ffffff',
  yellow: '#c8a400',
};

const DATE_CODE = /(?:^|[^\\"])[ymdhs]/i;

function isDateFormat(code: string | undefined): boolean {
  if (!code || code.toLowerCase() === 'general') return false;
  return DATE_CODE.test(code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, ''));
}

/** Format a number through an Excel number-format code. */
function formatNumber(value: number, numFmt: string | undefined, date1904: boolean): {
  text: string;
  color?: string;
} {
  if (!numFmt || numFmt.toLowerCase() === 'general') {
    return { text: generalText(value) };
  }
  try {
    // numfmt takes a JS Date directly for date codes, so serial conversion
    // happens here at the render boundary and nowhere inside evaluation.
    const input = isDateFormat(numFmt) ? serialToDate(value, { date1904 }) : value;
    const text = numFormat(numFmt, input as never);
    const named = formatColor(numFmt, value);
    return named ? { text, color: NUMFMT_COLORS[named] } : { text };
  } catch {
    return { text: generalText(value) };
  }
}

/**
 * Excel's General display: 15 significant digits, trailing zeros dropped.
 *
 * This is the engine's `numberToText`, not a second implementation of it. There
 * were two, and the one here was wrong: it stripped trailing zeros off the
 * output of `toPrecision(15)` without noticing that the output may be in
 * exponential form, where the last character is part of the *exponent*. So a
 * cell holding 1e20 rendered as `1.00000000000000e+2` — a confidently wrong
 * number, eighteen orders of magnitude out, with no warning attached. Every
 * exponent ending in a zero was affected: 1e20, 1e-10, 3e-20.
 *
 * Delegating also makes the renderer and the engine agree by construction, which
 * matters beyond this bug: `plainText` decides whether a computed value differs
 * from the file's cached one, and two General renderers that disagree would
 * report mismatches that are only a formatting difference.
 */
function generalText(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  return numberToText(value);
}

export interface ResolveOptions {
  numFmt?: string;
  date1904?: boolean;
  /**
   * Rich text or a hyperlink for this cell, from the layout. Both are
   * presentation rather than computation — they carry no number — so they win
   * over the value and never come from the engine.
   */
  content?: CellContentData;
}

export function resolveContent(cell: OverlayCell | undefined, opts: ResolveOptions = {}): CellContent {
  if (opts.content) return opts.content;

  if (!cell) return { kind: 'empty' };

  if (cell.provenance === 'unsupported' || cell.provenance === 'circular') {
    return { kind: 'unsupported', reason: cell.reason ?? 'this cell could not be computed' };
  }

  return fromScalar(cell.value, opts);
}

export function fromScalar(value: Scalar, opts: ResolveOptions = {}): CellContent {
  if (value === null || value === '') return { kind: 'empty' };
  if (isErr(value)) return { kind: 'error', text: value.kind };
  if (typeof value === 'boolean') return { kind: 'text', text: value ? 'TRUE' : 'FALSE' };
  if (typeof value === 'number') {
    const { text, color } = formatNumber(value, opts.numFmt, opts.date1904 ?? false);
    return color ? { kind: 'text', text, color, numeric: true } : { kind: 'text', text, numeric: true };
  }
  return { kind: 'text', text: String(value) };
}

/** Plain text for a value, for tooltips and the audit diff. */
export function plainText(value: Scalar): string {
  if (value === null) return '';
  if (isErr(value)) return value.kind;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return generalText(value);
  return value;
}
