/**
 * The oracle gate.
 *
 * Every probe is evaluated twice: once by LibreOffice, which wrote its answers
 * into `build/expected.json`, and once by this engine. Each comparison lands in
 * exactly one bucket:
 *
 *   match          the two agree
 *   unsupported    we refused to answer, and said why  — acceptable, tracked
 *   divergence     they disagree and we deliberately follow Excel — declared
 *   MISMATCH       they disagree and we did not know   — the failure that matters
 *
 * The last bucket is the whole point. A wrong number rendered confidently into a
 * financial model is invisible and unfalsifiable by eye, so its count is a hard
 * gate at zero rather than a metric to watch.
 *
 * Run `python3 tools/oracle/generate.py` to (re)build the fixtures; without them
 * this file skips rather than failing, so `npm test` works on a machine with no
 * LibreOffice.
 */

import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { colNumber } from '../../packages/formula-engine/src/a1.js';
import { isErr, type Scalar } from '../../packages/formula-engine/src/values.js';
import { Workbook } from '../../packages/formula-engine/src/workbook.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const BUILD = join(HERE, 'build');

type Expected =
  | { t: 'blank' }
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'other'; v: string };

interface Spec {
  grid: Record<string, number | string | boolean | null>;
  probeColumn: number;
  sheet: string;
  suites: Record<string, string[]>;
}

interface Divergence {
  oracle: string;
  ours?: string;
  why: string;
}

type Bucket = 'match' | 'unsupported' | 'divergence' | 'no-oracle' | 'mismatch';

interface Outcome {
  suite: string;
  formula: string;
  bucket: Bucket;
  oracle: string;
  ours: string;
  note?: string;
}

const haveFixtures = existsSync(join(BUILD, 'spec.json')) && existsSync(join(BUILD, 'expected.json'));

/**
 * The fixtures are read out here, not inside the `describe`, and the difference
 * is not stylistic.
 *
 * `describe.skipIf(cond)` still *executes* its callback — it collects the tests
 * and then marks the results skipped. So a `readFileSync` in the body runs even
 * when the condition says to skip, and this file threw `ENOENT` on every machine
 * that had never run `generate.py`: which is every fresh clone, and flatly
 * contradicted the promise in the header above. It went unnoticed for as long as
 * everyone who ran the suite happened to have the fixtures already.
 */
const fixtures = haveFixtures
  ? {
      spec: JSON.parse(readFileSync(join(BUILD, 'spec.json'), 'utf8')) as Spec,
      expected: JSON.parse(readFileSync(join(BUILD, 'expected.json'), 'utf8')) as Record<
        string,
        Expected[]
      >,
      divergences: JSON.parse(readFileSync(join(HERE, 'divergences.json'), 'utf8')) as Record<
        string,
        Divergence
      >,
    }
  : undefined;

// One visible skipped test rather than an empty file, so a run says *why* the
// oracle scored nothing instead of reporting nothing at all.
if (!fixtures) {
  it.skip('LibreOffice oracle — run `python3 tools/oracle/generate.py` to build the fixtures', () => {});
}

describe.skipIf(!fixtures)('LibreOffice oracle', () => {
  const { spec, expected, divergences } = fixtures ?? {
    spec: { sheet: 'Oracle', suites: {} } as Spec,
    expected: {} as Record<string, Expected[]>,
    divergences: {} as Record<string, Divergence>,
  };

  const outcomes: Outcome[] = [];

  for (const [suite, formulas] of Object.entries(spec.suites)) {
    const wb = new Workbook({ now: 46228.5 });
    wb.addSheet(spec.sheet);
    for (const [addr, value] of Object.entries(spec.grid)) {
      if (value === null) continue; // a blank cell must stay absent
      const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
      wb.setValue(0, Number(m[2]), colNumber(m[1]!), value as Scalar);
    }
    formulas.forEach((f, i) => wb.setFormula(0, i + 1, spec.probeColumn, f));
    wb.evaluateAll();

    formulas.forEach((formula, i) => {
      const rec = wb.record(0, i + 1, spec.probeColumn)!;
      const want = expected[suite]![i]!;
      const declared = divergences[`${suite}:${formula}`];

      const ours =
        rec.provenance === 'unsupported' || rec.provenance === 'circular'
          ? `⚠ ${rec.reason ?? rec.provenance}`
          : show(rec.value);
      const oracle = showExpected(want);

      let bucket: Bucket;
      if (rec.provenance === 'unsupported' || rec.provenance === 'circular') {
        bucket = 'unsupported';
      } else if (noOracle(want, rec.value)) {
        // The oracle does not know this function. Not a disagreement — an
        // absence of evidence, and it must not be scored as either.
        bucket = 'no-oracle';
      } else if (agrees(rec.value, want)) {
        bucket = declared ? 'mismatch' : 'match';
      } else {
        bucket = declared ? 'divergence' : 'mismatch';
      }

      const outcome: Outcome = { suite, formula, bucket, oracle, ours };
      if (declared) {
        outcome.note =
          bucket === 'mismatch'
            ? `declared divergence no longer holds — ${declared.why}`
            : declared.why;
      }
      outcomes.push(outcome);
    });
  }

  const bySuite = new Map<string, Outcome[]>();
  for (const o of outcomes) {
    const list = bySuite.get(o.suite) ?? [];
    list.push(o);
    bySuite.set(o.suite, list);
  }

  it('reports the scoreboard', () => {
    const rows: string[] = [
      '',
      '| suite | probes | match | unsupported | divergence | no oracle | MISMATCH |',
      '|---|---:|---:|---:|---:|---:|---:|',
    ];
    for (const [suite, list] of bySuite) {
      const n = (b: Bucket): number => list.filter((o) => o.bucket === b).length;
      rows.push(
        `| ${suite} | ${list.length} | ${n('match')} | ${n('unsupported')} | ${n('divergence')} | ${n('no-oracle')} | ${n('mismatch')} |`,
      );
    }
    const total = (b: Bucket): number => outcomes.filter((o) => o.bucket === b).length;
    rows.push(
      `| **total** | **${outcomes.length}** | **${total('match')}** | **${total('unsupported')}** | **${total('divergence')}** | **${total('no-oracle')}** | **${total('mismatch')}** |`,
    );

    // Coverage is about us: how many probes we were willing to answer.
    // Accuracy is about the ones somebody could grade, so `no-oracle` comes out
    // of the denominator rather than counting as a pass — an ungraded probe
    // must not be able to improve the score.
    const answered = outcomes.length - total('unsupported');
    const graded = answered - total('no-oracle');
    const accuracy = graded === 0 ? 0 : ((total('match') + total('divergence')) / graded) * 100;
    rows.push('');
    rows.push(
      `coverage ${((answered / outcomes.length) * 100).toFixed(1)}% · ` +
        `accuracy ${accuracy.toFixed(1)}% of ${graded} graded · false confidence ${total('mismatch')}`,
    );

    const ungraded = outcomes.filter((o) => o.bucket === 'no-oracle');
    if (ungraded.length) {
      rows.push(
        '',
        `Not graded — this LibreOffice does not implement these (it answered #NAME?).`,
        `A newer one would; see \`noOracle\`. These probes prove nothing either way:`,
      );
      for (const o of ungraded) rows.push(`  ${o.suite}: ${o.formula}  →  ours ${o.ours}`);
    }

    const unsupported = outcomes.filter((o) => o.bucket === 'unsupported');
    if (unsupported.length) {
      rows.push('', 'Refused (⚠ in the render, never a number):');
      for (const o of unsupported) rows.push(`  ${o.suite}: ${o.formula}  →  ${o.ours}`);
    }

    const bad = outcomes.filter((o) => o.bucket === 'mismatch');
    if (bad.length) {
      rows.push('', 'MISMATCHES:');
      for (const o of bad) {
        rows.push(`  ${o.suite}: ${o.formula}`);
        rows.push(`      oracle ${o.oracle}`);
        rows.push(`      ours   ${o.ours}${o.note ? `   (${o.note})` : ''}`);
      }
    }
    console.log(rows.join('\n'));
    expect(outcomes.length).toBeGreaterThan(0);
  });

  it('renders no number that disagrees with the oracle — hard gate', () => {
    const bad = outcomes.filter((o) => o.bucket === 'mismatch');
    expect(
      bad.map((o) => `${o.suite}: ${o.formula}  oracle=${o.oracle}  ours=${o.ours}`),
      'every disagreement must be either fixed or declared in divergences.json',
    ).toEqual([]);
  });
});

/**
 * Did the oracle simply not know this function?
 *
 * LibreOffice answers `#NAME?` for a function it does not implement, and which
 * functions those are depends on the *version* of LibreOffice doing the
 * answering. The author's machine runs 26.2 and knows `XLOOKUP`; the
 * `libreoffice-calc` on a GitHub runner is 24.2 and does not — XLOOKUP arrived
 * in 24.8. So five probes that score clean locally arrived in CI as five
 * MISMATCHes, which is the harness's loudest possible signal, reporting a
 * version difference as false confidence.
 *
 * That is a real defect in the harness, not a CI problem to route around. A
 * differential oracle has to be able to say *I cannot answer this one*, or every
 * environment older than the author's fails the build for the wrong reason —
 * and, worse, the reflex fix is to delete the probe, which is how a test suite
 * quietly stops covering the newest thing it has.
 *
 * The condition is deliberately narrow: the oracle said `#NAME?` and we produced
 * an actual value. If we answer `#NAME?` too, that is agreement and stays a
 * match. If we refuse, that is `unsupported` and is bucketed before this. And
 * because `no-oracle` is excluded from accuracy rather than counted as success,
 * it can never flatter the score — a probe nobody graded is reported as
 * ungraded, on its own line.
 */
function noOracle(want: Expected, ours: Scalar): boolean {
  const oracleSaidNAME = want.t === 'err' && want.v === '#NAME?';
  const weAnswered = !(isErr(ours) && ours.kind === '#NAME?');
  return oracleSaidNAME && weAnswered;
}

/** Relative epsilon for floats; exact for everything else. */
function agrees(ours: Scalar, want: Expected): boolean {
  switch (want.t) {
    case 'blank':
      // openpyxl cannot distinguish a formula that produced "" from one that
      // produced nothing, so both satisfy a blank oracle.
      return ours === null || ours === '';
    case 'num':
      if (typeof ours !== 'number') return false;
      return Math.abs(ours - want.v) <= 1e-9 * Math.max(1, Math.abs(ours), Math.abs(want.v));
    case 'str':
      return ours === want.v;
    case 'bool':
      return ours === want.v;
    case 'err':
      return isErr(ours) && ours.kind === want.v;
    default:
      return false;
  }
}

const show = (v: Scalar): string => {
  if (v === null) return '(blank)';
  if (isErr(v)) return v.kind;
  if (typeof v === 'string') return JSON.stringify(v);
  return String(v);
};

const showExpected = (e: Expected): string => (e.t === 'blank' ? '(blank)' : JSON.stringify(e.v));
