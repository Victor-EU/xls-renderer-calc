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
import { isErr, serialToDate, type Scalar } from '@xlscalc/formula-engine';
import type { OverlayCell } from '../bind.js';
import { resolveColor } from './theme.js';

export type RichRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  size?: number;
  name?: string;
};

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

/** Excel's General display: 15 significant digits, trailing zeros dropped. */
function generalText(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  const s = value.toPrecision(15);
  return s.includes('.') ? s.replace(/0+$/, '').replace(/\.$/, '') : s;
}

export interface ResolveOptions {
  numFmt?: string;
  date1904?: boolean;
  /** ExcelJS value, used only for rich text and hyperlinks, which carry no number. */
  styled?: unknown;
}

export function resolveContent(cell: OverlayCell | undefined, opts: ResolveOptions = {}): CellContent {
  const styled = opts.styled;

  // Rich text and hyperlinks are presentation, not computation; they come from
  // the styled parse and never from the engine.
  if (styled && typeof styled === 'object') {
    const o = styled as Record<string, unknown>;
    if ('richText' in o) {
      const runs = (o.richText as Array<{ text: string; font?: Record<string, unknown> }>).map((r) => {
        const f = (r.font ?? {}) as Record<string, unknown>;
        return {
          text: r.text,
          bold: f.bold as boolean | undefined,
          italic: f.italic as boolean | undefined,
          size: f.size as number | undefined,
          name: f.name as string | undefined,
          color: resolveColor(f.color),
        };
      });
      return { kind: 'rich', runs };
    }
    if ('hyperlink' in o) {
      return { kind: 'link', text: String(o.text ?? o.hyperlink), href: String(o.hyperlink) };
    }
  }

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
