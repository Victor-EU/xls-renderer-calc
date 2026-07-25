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

type Bucket = 'match' | 'unsupported' | 'divergence' | 'mismatch';

interface Outcome {
  suite: string;
  formula: string;
  bucket: Bucket;
  oracle: string;
  ours: string;
  note?: string;
}

const haveFixtures = existsSync(join(BUILD, 'spec.json')) && existsSync(join(BUILD, 'expected.json'));

describe.skipIf(!haveFixtures)('LibreOffice oracle', () => {
  const spec: Spec = JSON.parse(readFileSync(join(BUILD, 'spec.json'), 'utf8'));
  const expected: Record<string, Expected[]> = JSON.parse(
    readFileSync(join(BUILD, 'expected.json'), 'utf8'),
  );
  const divergences: Record<string, Divergence> = JSON.parse(
    readFileSync(join(HERE, 'divergences.json'), 'utf8'),
  );

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
      '| suite | probes | match | unsupported | divergence | MISMATCH |',
      '|---|---:|---:|---:|---:|---:|',
    ];
    for (const [suite, list] of bySuite) {
      const n = (b: Bucket): number => list.filter((o) => o.bucket === b).length;
      rows.push(
        `| ${suite} | ${list.length} | ${n('match')} | ${n('unsupported')} | ${n('divergence')} | ${n('mismatch')} |`,
      );
    }
    const total = (b: Bucket): number => outcomes.filter((o) => o.bucket === b).length;
    rows.push(
      `| **total** | **${outcomes.length}** | **${total('match')}** | **${total('unsupported')}** | **${total('divergence')}** | **${total('mismatch')}** |`,
    );

    const attempted = outcomes.length - total('unsupported');
    const accuracy = attempted === 0 ? 0 : ((total('match') + total('divergence')) / attempted) * 100;
    rows.push('');
    rows.push(
      `coverage ${(((outcomes.length - total('unsupported')) / outcomes.length) * 100).toFixed(1)}% · ` +
        `accuracy ${accuracy.toFixed(1)}% · false confidence ${total('mismatch')}`,
    );

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
