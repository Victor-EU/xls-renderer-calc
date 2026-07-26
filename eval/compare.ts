/**
 * Shared judging machinery for the harnesses.
 *
 * The oracle, the synthetic corpus and the real one ask different questions, but
 * they must decide *agreement* the same way — otherwise a bug could be a finding
 * in one harness and a pass in another, and the scoreboards would not be
 * comparable. This module is that single definition.
 *
 * It is not yet the whole of it: `agrees` and `show` are still duplicated in
 * each harness, which is exactly the drift this file exists to prevent and is
 * worth collapsing. `oracleCannotAnswer` is defined here and used by all of
 * them, because it was found by a CI run that broke two harnesses at once.
 */

import { colName } from '../packages/formula-engine/src/a1.js';
import { isErr, type Scalar } from '../packages/formula-engine/src/values.js';
import { readXlsx, type RawCell } from '../packages/xlsx-preview/src/ooxml.js';

/**
 * Read a workbook's values into an address → value map.
 *
 * Deliberately a separate decoder from `bind.ts`. The oracle should not be read
 * through the same code path as the thing it is judging.
 */
export function readOracle(buf: ArrayBuffer): Map<string, Scalar> {
  const raw = readXlsx(buf);
  const out = new Map<string, Scalar>();
  for (const sheet of raw.sheets) {
    for (const cell of sheet.cells) {
      const v = decode(cell, raw.sharedStrings);
      if (v === undefined) continue;
      out.set(`${sheet.name}!${colName(cell.col)}${cell.row}`, v);
    }
  }
  return out;
}

export function decode(cell: RawCell, strings: string[]): Scalar | undefined {
  if (cell.t === 'inlineStr') return cell.is ?? '';
  if (cell.v === undefined || cell.v === '') return cell.t === 'str' ? '' : undefined;
  switch (cell.t) {
    case 's': {
      const i = Number(cell.v);
      return Number.isFinite(i) ? (strings[i] ?? '') : '';
    }
    case 'str':
      return cell.v;
    case 'b':
      return cell.v === '1';
    case 'e':
      return { kind: cell.v } as unknown as Scalar;
    default: {
      const n = Number(cell.v);
      return Number.isFinite(n) ? n : cell.v;
    }
  }
}

/** What LibreOffice answers for a function the running version does not implement. */
export const UNKNOWN_FUNCTION = '#NAME?';

/**
 * Did the oracle fail to answer because it does not know the function?
 *
 * Which functions LibreOffice implements depends on which LibreOffice. The
 * author's machine runs 26.2 and knows `XLOOKUP`; a GitHub runner's
 * `libreoffice-calc` is 24.2 and does not — XLOOKUP arrived in 24.8 — so it
 * writes `#NAME?` into the fixture and every probe using it reads as a
 * disagreement. Five of them turned up as MISMATCHes on the first CI run, which
 * is the loudest signal these harnesses have, reporting a version difference as
 * false confidence.
 *
 * An oracle has to be able to report an absence of evidence. Without this the
 * build fails on any environment older than the author's, and the reflex fix is
 * to delete the probe — which is how a suite quietly stops covering the newest
 * thing in it.
 *
 * Narrow on purpose: the oracle said `#NAME?` and we produced something else. If
 * we answer `#NAME?` too that is agreement and stays a match; if we refused
 * outright that is `unsupported` and is bucketed before this. Callers must
 * exclude the result from accuracy rather than count it as a pass — a probe
 * nobody graded must not be able to flatter the score.
 */
export function oracleCannotAnswer(oracleErrorKind: string | undefined, ours: Scalar): boolean {
  if (oracleErrorKind !== UNKNOWN_FUNCTION) return false;
  return !(isErr(ours) && ours.kind === UNKNOWN_FUNCTION);
}

/** The oracle's value as an error kind, if it is one. */
export const errorKind = (v: Scalar | undefined): string | undefined =>
  typeof v === 'object' && v !== null ? (v as { kind?: string }).kind : undefined;

export function agrees(ours: Scalar, want: Scalar): boolean {
  if (isErr(ours)) {
    const kind = (want as { kind?: string })?.kind;
    return typeof kind === 'string' && ours.kind === kind;
  }
  if (typeof want === 'object' && want !== null) return false; // oracle is an error, we are not
  if (typeof ours === 'number' && typeof want === 'number') {
    // Relative tolerance: a 360-row amortisation accumulates rounding, and a
    // difference in the fifteenth digit is agreement, not a finding.
    return Math.abs(ours - want) <= 1e-9 * Math.max(1, Math.abs(ours), Math.abs(want));
  }
  if (ours === null) return want === null || want === '' || want === 0;
  return ours === want;
}

export function show(v: Scalar | undefined): string {
  if (v === undefined) return '(absent)';
  if (v === null) return '(blank)';
  if (isErr(v)) return v.kind;
  if (typeof v === 'object') return String((v as { kind?: string }).kind ?? v);
  if (typeof v === 'string') return JSON.stringify(clip(v, 60));
  return String(v);
}

export const clip = (s: string, n: number): string => (s.length <= n ? s : `${s.slice(0, n - 1)}…`);

export const pct = (a: number, b: number): string =>
  b === 0 ? 'n/a' : `${((a / b) * 100).toFixed(1)}%`;

export function firstFunction(formula: string): string | undefined {
  const m = /(?<![A-Za-z0-9_.])([A-Z][A-Z0-9._]*)\s*\(/.exec(formula);
  return m?.[1];
}
