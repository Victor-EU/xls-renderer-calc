/**
 * `findHardcoded`, and the locality it was missing.
 *
 * This detector makes the most specific claim in the library: it names a cell,
 * quotes the number in it, and says what the number should have been. That is
 * the point — "D6 states 1261.2; SUM(D4:D5) is 1221.2" is actionable in a way
 * that "something looks off" is not. It is also why a false positive here costs
 * more than a miss: acting on one means editing a cell that was correct.
 *
 * It had no test, and it was wrong. Every formula cell in the literal's row or
 * column counted as a "neighbour", so any shape shared by two of them was
 * extrapolated over the whole line. On the demo model in this repository that
 * produced 33 findings and not one of them was real. The suite below pins both
 * halves: the shapes it must still catch, and the extrapolations it must refuse.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { bind, readXlsx, describeFinding } from '../src/index.js';
import { minimalXlsx } from './minimal-xlsx.js';

const num = (ref: string, v: number): string => `<c r="${ref}"><v>${v}</v></c>`;
const fx = (ref: string, formula: string): string => `<c r="${ref}"><f>${formula}</f><v></v></c>`;
const txt = (ref: string, s: string): string => `<c r="${ref}" t="inlineStr"><is><t>${s}</t></is></c>`;

/** Two input rows, so every `SUM(x4:x5)` below has something real to add. */
const INPUTS =
  `<row r="4">${num('B4', 10) + num('C4', 10) + num('D4', 10) + num('E4', 10) + num('F4', 10)}</row>` +
  `<row r="5">${num('B5', 20) + num('C5', 20) + num('D5', 20) + num('E5', 20) + num('F5', 20)}</row>`;

const findings = (rows: string) => bind(readXlsx(minimalXlsx({ rows }))).hardcoded;

const demo = (name: string) => {
  const path = new URL(`../../../apps/demo/public/${name}`, import.meta.url);
  const buf = readFileSync(path);
  return bind(readXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer));
};

describe('the shapes it must catch', () => {
  it('finds a typed-in total at the end of a row of totals', () => {
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + fx('D6', 'SUM(D4:D5)') + num('E6', 99)}</row>`);
    expect(f).toHaveLength(1);
    expect(f[0]!.address).toBe('E6');
    expect(f[0]!.stated).toBe(99);
    expect(f[0]!.expected).toBe(30);
    // The specificity is the product. Keep it.
    expect(describeFinding(f[0]!)).toContain('SUM(E4:E5)');
  });

  it('finds one sitting in the middle of the band, filling its own gap', () => {
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + num('D6', 99) + fx('E6', 'SUM(E4:E5)')}</row>`);
    expect(f.map((x) => x.address)).toEqual(['D6']);
  });

  it('finds two adjacent hardcoded cells, which bridge for each other', () => {
    // Neither breaks the band, because the only thing between the formulas is
    // the other number being judged in the same pass.
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + num('D6', 99) + num('E6', 98) + fx('F6', 'SUM(F4:F5)')}</row>`);
    expect(f.map((x) => x.address).sort()).toEqual(['D6', 'E6']);
  });

  it('says nothing when the typed-in number is right', () => {
    expect(findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + fx('D6', 'SUM(D4:D5)') + num('E6', 30)}</row>`)).toEqual([]);
  });
});

describe('the extrapolations it must refuse', () => {
  it('will not call two cells at opposite ends of a column a pattern', () => {
    // The exact shape of the 33 false findings: `B6` and `B18` share "sum of
    // the two above", eleven unrelated rows apart, and the rule was applied to
    // every literal in the column — including a header row of date serials.
    const f = findings(
      `<row r="3">${num('B3', 46022)}</row>` +
        `<row r="4">${num('B4', 412)}</row>` +
        `<row r="5">${num('B5', -168.9)}</row>` +
        `<row r="6">${fx('B6', 'B4+B5')}</row>` +
        `<row r="16">${num('B16', 12.1)}</row>` +
        `<row r="17">${num('B17', 22.6)}</row>` +
        `<row r="18">${fx('B18', 'B16+B17')}</row>`,
    );
    expect(f).toEqual([]);
  });

  it('will not reach across a cell it does not understand', () => {
    // `C6` is a formula of a different shape. That makes B6 and D6 two
    // neighbourhoods, not one, and a rule from either must not cross it.
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'B6*2') + fx('D6', 'SUM(D4:D5)') + num('E6', 99)}</row>`);
    expect(f).toEqual([]);
  });

  it('will not reach across a label or an empty cell', () => {
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + txt('C6', 'note') + fx('D6', 'SUM(D4:D5)') + num('E6', 99)}</row>`);
    expect(f).toEqual([]);
  });

  it('will not flag the seed cell in front of a run of formulas', () => {
    // An opening balance, a period-zero column, the last month of actuals
    // before a projection starts: the cell before a run is where a model seeds
    // a series, and derived cells come *after* the inputs they derive from. On
    // the real corpus this one idiom was most of what survived the other rules.
    const f = findings(
      `<row r="4">${num('B4', 100) + num('C4', 0) + num('D4', 0) + num('E4', 0)}</row>` +
        `<row r="5">${num('B5', 5) + fx('C5', 'B5*2') + fx('D5', 'C5*2') + fx('E5', 'D5*2')}</row>`,
    );
    expect(f).toEqual([]);
  });

  it('will not flag a number that is merely on the same line as the band', () => {
    // The band stops at D6; F6 is past its edge with a hole in between.
    const f = findings(INPUTS + `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + fx('D6', 'SUM(D4:D5)') + num('F6', 99)}</row>`);
    expect(f).toEqual([]);
  });
});

describe('a parallel data series is not a missing formula', () => {
  // `B:D` are this period, computed. `E` is last year, hand-entered top to
  // bottom — the layout of almost every management report, and the reason two
  // false findings survived the first fix.
  const PRIOR_YEAR =
    `<row r="4">${num('B4', 10) + num('C4', 10) + num('D4', 10)}</row>` +
    `<row r="5">${num('B5', 20) + num('C5', 20) + num('D5', 20)}</row>` +
    `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + fx('D6', 'SUM(D4:D5)') + num('E6', 1502)}</row>` +
    `<row r="7">${fx('B7', 'B6*2') + fx('C7', 'C6*2') + fx('D7', 'D6*2') + num('E7', 300)}</row>`;

  it('says nothing about either cell of the hand-entered column', () => {
    // E7 is refused because its neighbours read a computed cell and it would
    // read a raw one; E6 because its neighbours read real inputs and it would
    // read two cells that do not exist. Different bits, same conclusion.
    expect(findings(PRIOR_YEAR)).toEqual([]);
  });

  it('still reports the same cell when the column really is part of the block', () => {
    const f = findings(
      `<row r="4">${num('B4', 10) + num('C4', 10) + num('D4', 10) + num('E4', 10)}</row>` +
        `<row r="5">${num('B5', 20) + num('C5', 20) + num('D5', 20) + num('E5', 20)}</row>` +
        `<row r="6">${fx('B6', 'SUM(B4:B5)') + fx('C6', 'SUM(C4:C5)') + fx('D6', 'SUM(D4:D5)') + fx('E6', 'SUM(E4:E5)')}</row>` +
        `<row r="7">${fx('B7', 'B6*2') + fx('C7', 'C6*2') + fx('D7', 'D6*2') + num('E7', 300)}</row>`,
    );
    // Same position, same stated value; the only change is that E now belongs
    // to the block. The detector must not have gone blind to its last column.
    expect(f.map((x) => x.address)).toEqual(['E7']);
    expect(f[0]!.expected).toBe(60);
  });
});

describe('the demo workbooks in this repository', () => {
  it('reports the one hardcoded total in hardcoded-total.xlsx', () => {
    const f = demo('hardcoded-total.xlsx').hardcoded;
    expect(f).toHaveLength(1);
    expect(f[0]!.address).toBe('E6');
    expect(describeFinding(f[0]!)).toContain('SUM(E4:E5)');
  });

  it('reports nothing on financial-model-nocache.xlsx, which has nothing to report', () => {
    // This file used to produce 33 findings, all false. It is a hand-built
    // model: rows 3-5 are its inputs and a date header, and the formulas below
    // them are correct. There is no hardcoded cell in it.
    expect(demo('financial-model-nocache.xlsx').hardcoded).toEqual([]);
  });

  it('reports nothing on formula-tour.xlsx either', () => {
    expect(demo('formula-tour.xlsx').hardcoded).toEqual([]);
  });
});
