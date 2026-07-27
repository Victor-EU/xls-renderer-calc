/**
 * "Did the file already have a value for this cell?"
 *
 * The question the recalc story turns on, and the one an agent-generated
 * workbook answers *no* to for every formula it contains — which is the case
 * this project exists for. Two entry points ask it: `inspectXlsx`, cheaply,
 * before deciding whether to render a file at all, and `bind`, during the load.
 * They disagreed. `bind` honoured the `t="str"` refinement and `inspect` tested
 * `<v>` alone, so every `IF(...,"",...)` — a formula that appears in essentially
 * every finished model — was counted as never computed by the tool you run
 * first.
 *
 * Neither number was alarming on its own, which is what made it worth a test
 * rather than just a fix: the failure was two plausible answers to one question,
 * and the wrong one arriving earlier and cheaper. Nobody compares them by hand.
 */

import { describe, expect, it } from 'vitest';
import { hasCachedValue, inspectXlsx, loadXlsx, type RawCell } from '../src/index.js';
import { minimalXlsx } from './minimal-xlsx.js';

const cell = (over: Partial<RawCell>): RawCell => ({ row: 1, col: 1, f: 'X()', ...over });

describe('hasCachedValue', () => {
  it('reads the type attribute, not just the value element', () => {
    // Excel writes `t` whenever it writes a result, so the attribute is the
    // evidence that the formula ran and the empty `<v>` is only what it
    // returned. Without `t` there is nothing saying anything ever ran.
    expect(hasCachedValue(cell({ t: 'str', v: '' }))).toBe(true); // computed to ""
    expect(hasCachedValue(cell({ t: 'str' }))).toBe(true); // ditto, `<v>` elided
    expect(hasCachedValue(cell({ v: '' }))).toBe(false); // never computed
    expect(hasCachedValue(cell({}))).toBe(false); // ditto, no `<v>` at all
  });

  it('treats any non-empty value as carried, whatever its type', () => {
    for (const c of [
      cell({ v: '4' }),
      cell({ t: 's', v: '0' }),
      cell({ t: 'b', v: '1' }),
      cell({ t: 'e', v: '#DIV/0!' }),
      cell({ t: 'str', v: 'text' }),
    ]) {
      expect(hasCachedValue(c), JSON.stringify(c)).toBe(true);
    }
  });

  it('does not count an empty result of any other type as carried', () => {
    // An error or a boolean always writes its value out; an empty `<v>` under
    // those types is a malformed cell, not a computed emptiness. Only `str` has
    // a genuine empty result to express.
    for (const t of ['s', 'b', 'e', 'inlineStr', 'd']) {
      expect(hasCachedValue(cell({ t, v: '' })), t).toBe(false);
    }
  });
});

/** One sheet holding every shape at once, so one file settles the question. */
const MIXED = minimalXlsx({
  rows:
    `<row r="1">` +
    `<c r="A1" t="str"><f>IF(TRUE,"","x")</f><v></v></c>` + // computed to ""
    `<c r="B1"><f>SUM(D1:D1)</f></c>` + //                    never computed, no <v>
    `<c r="C1"><f>1+1</f><v></v></c>` + //                    never computed, empty <v>
    `<c r="D1"><f>2+2</f><v>4</v></c>` + //                   carried
    `<c r="E1"><v>10</v></c>` + //                            a literal, not a formula
    `</row>`,
});

describe('the two entry points count the same file the same way', () => {
  it('agrees on how many formulas have no value', async () => {
    const inspected = inspectXlsx(MIXED);
    const loaded = await loadXlsx(MIXED);

    // B1 and C1. Not A1: it ran, and returned "".
    expect(inspected.uncached).toBe(2);
    expect(inspected.uncached).toBe(loaded.model.facts.uncached);
    expect(inspected.formulas).toBe(loaded.model.facts.formulas);
  });

  it('counts formulas, not cells', () => {
    const inspected = inspectXlsx(MIXED);
    expect(inspected.formulas).toBe(4); // E1 is a literal
    expect(inspected.sheets[0]!.cells).toBe(5);
    expect(inspected.sheets.reduce((n, s) => n + s.uncached, 0)).toBe(inspected.uncached);
  });

  it('backs the count with the value each cell actually ends up with', async () => {
    const { model } = await loadXlsx(MIXED);
    const at = (col: number) => model.cell(0, 1, col)!;

    // A1 had a value, and it was the empty string — so it is not a hole to fill.
    expect(at(1).cached).toBe('');
    expect(at(1).value).toBe('');

    // C1 had none, and we supplied one. That is the whole product.
    expect(at(3).cached).toBeUndefined();
    expect(at(3).provenance).toBe('computed');
    expect(at(3).value).toBe(2);
  });
});

describe('a workbook no one ever calculated', () => {
  // What an agent writes: formulas, inputs, and not one computed value. Every
  // other renderer shows this file as a grid of blanks.
  const GENERATED = minimalXlsx({
    rows:
      `<row r="1"><c r="A1"><v>100</v></c><c r="B1"><v>1.2</v></c></row>` +
      `<row r="2"><c r="A2"><f>A1*B1</f></c><c r="B2"><f>A2-A1</f></c></row>`,
  });

  it('is recognised as entirely uncomputed before it is loaded', () => {
    const inspected = inspectXlsx(GENERATED);
    expect(inspected.formulas).toBe(2);
    expect(inspected.uncached).toBe(inspected.formulas);
    expect(inspected.fullyCovered).toBe(true);
  });

  it('gets its numbers from us, and says so', async () => {
    const { model } = await loadXlsx(GENERATED);
    expect(model.facts.uncached).toBe(2);

    const a2 = model.cell(0, 2, 1)!;
    expect(a2.value).toBeCloseTo(120, 10);
    expect(a2.provenance).toBe('computed');
    // Nothing to compare against, so nothing can be flagged as disagreeing —
    // the mismatch audit is inert on a file like this, by construction.
    expect(a2.cached).toBeUndefined();
    expect(model.report.mismatches).toEqual([]);
    expect(model.report.stats.cached).toBe(0);
  });
});
