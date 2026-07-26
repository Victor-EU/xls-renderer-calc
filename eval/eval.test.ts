/**
 * The corpus gate.
 *
 * Ten whole workbooks, each loaded twice: once as the generator emitted it —
 * formulas with an empty `<v>`, which is the artefact this project exists to
 * render — and once after LibreOffice has recalculated every cell. Then every
 * formula cell in the file is compared.
 *
 * This differs from `tools/oracle` in what it is testing. That harness probes
 * one formula at a time against a shared grid, and answers "is this function
 * right". This one loads a real model through the real reader and answers a
 * question no probe suite can: does the *system* — zip reader, shared-formula
 * translation, cross-sheet binding, evaluation order, the function library —
 * produce the same workbook a spreadsheet application does.
 *
 * Buckets, in the order they are tested:
 *
 *   unsupported   we refused, and said why. Correct for INDIRECT and friends;
 *                 a coverage gap everywhere else. Never a wrong number.
 *   circular      we detected a cycle. On the M10 Circular sheet this is the
 *                 only correct answer.
 *   volatile      depends on NOW/TODAY, so it cannot agree with an oracle
 *                 computed at a different moment. Excluded, not excused.
 *   no-oracle     LibreOffice produced nothing for the cell, so there is
 *                 nothing to compare against.
 *   match         the two agree.
 *   inherited     they disagree, but re-running *our* formula against *the
 *                 oracle's* inputs reproduces the oracle's answer — so the
 *                 disagreement came in from upstream and this cell computed
 *                 correctly. Almost all of these trace to the oracle's own
 *                 precision: LibreOffice writes at most 15 significant digits
 *                 into the file, and a model with a ROUND sitting on a
 *                 half-way boundary amplifies the sixteenth digit into a
 *                 visible tenth. Counted, never a gate.
 *   divergence    they disagree and we deliberately follow Excel. Declared in
 *                 divergences.json, and the gate is symmetric: a declared
 *                 divergence that starts matching also fails.
 *   MISMATCH      they disagree and we did not know. Hard gate at zero.
 *
 * Build the fixtures with `python3 eval/build.py` (needs LibreOffice). Without
 * them this file skips rather than failing.
 */

import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { beforeAll, describe, expect, it } from 'vitest';
import { colName } from '../packages/formula-engine/src/a1.js';
import { isErr, type Scalar } from '../packages/formula-engine/src/values.js';
import { Workbook } from '../packages/formula-engine/src/workbook.js';
import { loadXlsx } from '../packages/xlsx-preview/src/parse.js';
import { readXlsx, type RawCell } from '../packages/xlsx-preview/src/ooxml.js';
import { errorKind, oracleCannotAnswer } from './compare.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const BUILD = join(HERE, 'build');
const RECALC = join(BUILD, 'recalc');

const MODELS = [
  ['m01_budget', 'Annual budget and variance'],
  ['m02_valuation_dcf', 'DCF valuation and sensitivity'],
  ['m03_lbo', 'LBO with debt schedule'],
  ['m04_three_statement', 'Linked three-statement model'],
  ['m05_workflow_approvals', 'Approval workflow, IF-heavy'],
  ['m06_sales_data', '2,000 transactions with SUMIFS'],
  ['m07_cohort_retention', 'Cohort retention triangle'],
  ['m08_loan_amortization', '360-month amortisation'],
  ['m09_inventory_planning', 'Inventory planning, 150 SKUs'],
  ['m10_edge_cases', 'Semantics, errors and refusals'],
] as const;

/** Frozen clock, so NOW/TODAY are at least reproducible run to run. */
const NOW = 46228.5; // 2026-07-25 12:00

type Bucket =
  | 'match'
  | 'unsupported'
  | 'circular'
  | 'volatile'
  | 'no-oracle'
  | 'inherited'
  | 'divergence'
  | 'mismatch';

interface Outcome {
  model: string;
  sheet: string;
  address: string;
  formula: string;
  bucket: Bucket;
  ours: string;
  oracle: string;
  reason?: string;
  subject?: string;
  note?: string;
}

interface Divergence {
  why: string;
  confidence: 'excel-documented' | 'oracle-limit' | 'unverified';
}

const present = MODELS.filter(
  ([name]) => existsSync(join(BUILD, `${name}.xlsx`)) && existsSync(join(RECALC, `${name}.xlsx`)),
);

describe.skipIf(present.length === 0)('eval corpus', () => {
  const divergences: Record<string, Divergence> = Object.fromEntries(
    Object.entries(
      existsSync(join(HERE, 'divergences.json'))
        ? (JSON.parse(readFileSync(join(HERE, 'divergences.json'), 'utf8')) as Record<string, Divergence>)
        : {},
      // Keys beginning with `_` are prose for the reader, not declarations.
    ).filter(([k]) => !k.startsWith('_')),
  );

  const outcomes: Outcome[] = [];
  const timings: Array<{ model: string; label: string; cells: number; formulas: number; parseMs: number; evalMs: number }> = [];
  const byModel = new Map<string, Outcome[]>();

  beforeAll(async () => {
  for (const [name, label] of present) {
    const emitted = new Uint8Array(readFileSync(join(BUILD, `${name}.xlsx`))).buffer;
    const recalced = new Uint8Array(readFileSync(join(RECALC, `${name}.xlsx`))).buffer;

    const doc = await loadXlsx(emitted, { now: NOW, preferCachedVolatile: false });
    const oracle = readOracle(recalced);
    // A workbook holding the oracle's answers as plain values, used to ask
    // whether a disagreement started here or arrived from upstream.
    const asOracle = new Workbook({ now: NOW });
    for (const s2 of doc.raw.sheets) asOracle.addSheet(s2.name);
    doc.raw.sheets.forEach((s2, si) => {
      for (const c of s2.cells) {
        const v = oracle.get(`${s2.name}!${colName(c.col)}${c.row}`);
        if (v !== undefined) asOracle.setValue(si, c.row, c.col, v);
      }
    });

    let cells = 0;
    let formulas = 0;
    doc.raw.sheets.forEach((sheet, sheetIndex) => {
      cells += sheet.cells.length;
      for (const cell of sheet.cells) {
        if (!cell.f) continue;
        formulas++;
        const address = `${colName(cell.col)}${cell.row}`;
        const rec = doc.model.engine.record(sheetIndex, cell.row, cell.col);
        const want = oracle.get(`${sheet.name}!${address}`);
        const key = `${name}:${sheet.name}!${address}`;
        const declared = divergences[key];

        const outcome: Outcome = {
          model: name,
          sheet: sheet.name,
          address,
          formula: cell.f,
          bucket: 'match',
          ours: '',
          oracle: want === undefined ? '(absent)' : show(want),
        };
        if (rec === undefined) {
          outcome.bucket = 'mismatch';
          outcome.ours = '(no record)';
          outcomes.push(outcome);
          continue;
        }

        outcome.ours =
          rec.provenance === 'unsupported' || rec.provenance === 'circular'
            ? `⚠ ${rec.reason ?? rec.provenance}`
            : show(rec.value);
        if (rec.reason) outcome.reason = rec.reason;
        const engineSheet = doc.model.engine.sheet(sheetIndex);
        const fault = engineSheet?.formulas.get(engineSheet.key(cell.row, cell.col))?.fault;
        if (fault?.subject) outcome.subject = fault.subject;

        if (rec.provenance === 'circular') outcome.bucket = 'circular';
        else if (rec.provenance === 'unsupported') outcome.bucket = 'unsupported';
        else if (rec.provenance === 'volatile') outcome.bucket = 'volatile';
        // Absent from the oracle, or present as a `#NAME?` this LibreOffice
        // could not evaluate — both are "nobody graded this", not a finding.
        else if (want === undefined || oracleCannotAnswer(errorKind(want), rec.value)) {
          outcome.bucket = 'no-oracle';
        }
        else if (agrees(rec.value, want)) outcome.bucket = declared ? 'mismatch' : 'match';
        else if (declared) outcome.bucket = 'divergence';
        else {
          const redone = asOracle.tryEvaluate(sheetIndex, cell.row, cell.col, cell.f);
          outcome.bucket =
            'value' in redone && agrees(redone.value, want) ? 'inherited' : 'mismatch';
          if (outcome.bucket === 'inherited') {
            outcome.note = 'our formula over the oracle\'s inputs reproduces the oracle';
          }
        }

        if (declared) {
          outcome.note =
            outcome.bucket === 'mismatch'
              ? `declared divergence no longer holds — ${declared.why}`
              : declared.why;
        }
        outcomes.push(outcome);
      }
    });

    timings.push({
      model: name,
      label,
      cells,
      formulas,
      parseMs: doc.parseMs,
      evalMs: doc.evalMs,
    });
  }

  for (const o of outcomes) {
    const list = byModel.get(o.model) ?? [];
    list.push(o);
    byModel.set(o.model, list);
  }
  }, 300_000);

  it('writes the scoreboard', () => {
    const n = (list: Outcome[], b: Bucket): number => list.filter((o) => o.bucket === b).length;
    const rows: string[] = [];
    rows.push('# Eval corpus report');
    rows.push('');
    rows.push('Generated by `eval/eval.test.ts`. Do not edit by hand — run `npm run eval`.');
    rows.push('');
    rows.push('Each workbook is loaded as emitted (no cached values) and compared cell by');
    rows.push('cell against the same workbook after LibreOffice recalculated it.');
    rows.push('');
    rows.push('| model | formulas | match | inherited | refused | circular | volatile | no oracle | divergence | **MISMATCH** |');
    rows.push('|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|');
    for (const [model, label] of present) {
      const list = byModel.get(model) ?? [];
      rows.push(
        `| \`${model}\`<br>${label} | ${list.length} | ${n(list, 'match')} | ${n(list, 'inherited')} | ` +
          `${n(list, 'unsupported')} | ${n(list, 'circular')} | ${n(list, 'volatile')} | ` +
          `${n(list, 'no-oracle')} | ${n(list, 'divergence')} | **${n(list, 'mismatch')}** |`,
      );
    }
    const t = (b: Bucket): number => n(outcomes, b);
    rows.push(
      `| **total** | **${outcomes.length}** | **${t('match')}** | **${t('inherited')}** | ` +
        `**${t('unsupported')}** | **${t('circular')}** | **${t('volatile')}** | ` +
        `**${t('no-oracle')}** | **${t('divergence')}** | **${t('mismatch')}** |`,
    );
    rows.push('');

    const compared = t('match') + t('divergence') + t('mismatch') + t('inherited');
    const answerable = outcomes.length - t('unsupported') - t('circular');
    rows.push(
      `**coverage** ${pct(answerable, outcomes.length)} · ` +
        `**accuracy** ${pct(t('match') + t('divergence') + t('inherited'), compared)} · ` +
        `**false confidence** ${t('mismatch')}`,
    );
    rows.push('');
    rows.push(
      'Coverage is the share of formula cells the engine was willing to answer. ' +
        'Accuracy is how often it agreed with LibreOffice where both produced a value. ' +
        'False confidence is the count that matters: a number rendered confidently ' +
        'that disagrees with the oracle and was not declared.',
    );

    // --- refusals, bucketed by function, because that is the backlog ---------
    const refusals = new Map<string, { count: number; reason: string; models: Set<string> }>();
    for (const o of outcomes) {
      if (o.bucket !== 'unsupported') continue;
      const subject = o.subject ?? firstFunction(o.formula) ?? '(other)';
      const entry = refusals.get(subject) ?? { count: 0, reason: o.reason ?? '', models: new Set() };
      entry.count++;
      entry.models.add(o.model);
      refusals.set(subject, entry);
    }
    if (refusals.size) {
      rows.push('');
      rows.push('## What the engine refused');
      rows.push('');
      rows.push('| subject | cells | models | reason |');
      rows.push('|---|---:|---|---|');
      for (const [subject, e] of [...refusals].sort((a, b) => b[1].count - a[1].count)) {
        rows.push(
          `| \`${subject}\` | ${e.count} | ${[...e.models].map((m) => m.slice(0, 3)).join(' ')} | ${e.reason} |`,
        );
      }
    }

    // --- declared divergences ------------------------------------------------
    const declaredSeen = outcomes.filter((o) => o.bucket === 'divergence');
    if (declaredSeen.length) {
      const conf = (o: Outcome): string =>
        divergences[`${o.model}:${o.sheet}!${o.address}`]?.confidence ?? '?';
      rows.push('');
      rows.push(`## Declared divergences (${declaredSeen.length})`);
      rows.push('');
      rows.push('Cells where we disagree with LibreOffice on purpose. The gate is');
      rows.push('symmetric — if one of these starts matching, that fails too.');
      rows.push('');
      rows.push('| cell | confidence | formula | oracle | ours |');
      rows.push('|---|---|---|---|---|');
      for (const o of declaredSeen) {
        rows.push(
          `| \`${o.model.slice(0, 3)} ${o.sheet}!${o.address}\` | ${conf(o)} | ` +
            `\`${clip(o.formula, 46)}\` | \`${clip(o.oracle, 26)}\` | \`${clip(o.ours, 26)}\` |`,
        );
      }
      const unverified = declaredSeen.filter((o) => conf(o) === 'unverified');
      if (unverified.length) {
        rows.push('');
        rows.push(
          `**${unverified.length} of these are marked \`unverified\`** — we believe we follow Excel, ` +
            'but the cell has not been run in Excel itself. See `eval/divergences.json` for what ' +
            'would settle each one.',
        );
      }
    }

    // --- mismatches ----------------------------------------------------------
    const bad = outcomes.filter((o) => o.bucket === 'mismatch');
    if (bad.length) {
      rows.push('');
      rows.push(`## Mismatches (${bad.length})`);
      rows.push('');
      for (const o of bad.slice(0, 400)) {
        rows.push(`- \`${o.model}\` **${o.sheet}!${o.address}**  \`${clip(o.formula, 110)}\``);
        rows.push(`  - oracle \`${o.oracle}\` · ours \`${o.ours}\`${o.note ? ` — ${o.note}` : ''}`);
      }
      if (bad.length > 400) rows.push(`- …and ${bad.length - 400} more`);
    }

    // --- timings -------------------------------------------------------------
    rows.push('');
    rows.push('## Load and evaluation time');
    rows.push('');
    rows.push('| model | cells | formulas | parse | evaluate |');
    rows.push('|---|---:|---:|---:|---:|');
    for (const t2 of timings) {
      rows.push(
        `| \`${t2.model}\` | ${t2.cells.toLocaleString()} | ${t2.formulas.toLocaleString()} | ` +
          `${t2.parseMs} ms | ${t2.evalMs} ms |`,
      );
    }

    writeFileSync(join(HERE, 'REPORT.md'), rows.join('\n') + '\n');
    console.log(rows.slice(0, 40).join('\n'));
    console.log(`\nfull report → eval/REPORT.md`);
    expect(outcomes.length).toBeGreaterThan(0);
  });

  it('renders no number that disagrees with the oracle — hard gate', () => {
    const bad = outcomes.filter((o) => o.bucket === 'mismatch');
    expect(
      bad
        .slice(0, 40)
        .map((o) => `${o.model} ${o.sheet}!${o.address} ${clip(o.formula, 80)}\n` +
          `      oracle=${o.oracle}\n      ours  =${o.ours}`),
      'every disagreement must be fixed or declared in eval/divergences.json',
    ).toEqual([]);
  });

  it('refuses every circular reference on the M10 Circular sheet', () => {
    const sheet = outcomes.filter((o) => o.model === 'm10_edge_cases' && o.sheet === 'Circular');
    if (sheet.length === 0) return;

    // The doctrine is "never a number", not "carries a particular label". A cell
    // inside the cycle is `circular`; one merely downstream of it is
    // `unsupported` with the offending cell named. Both are correct refusals,
    // and insisting on the first would be testing the label rather than the
    // promise.
    const refused = (b: Bucket): boolean => b === 'circular' || b === 'unsupported';

    // B22 is the control: no cycle anywhere in its history, so it must compute.
    const cyclic = sheet.filter((o) => o.address !== 'B22');
    expect(
      cyclic.filter((o) => !refused(o.bucket)).map((o) => `${o.address} → ${o.ours}`),
      'a cell in or downstream of a cycle must never render a number',
    ).toEqual([]);

    const control = sheet.find((o) => o.address === 'B22');
    expect(control?.bucket, 'poisoning must not spread to cells with no cycle upstream').toBe(
      'match',
    );
  });
});

/** Cached values out of the recalculated file, keyed `Sheet!A1`. */
function readOracle(buf: ArrayBuffer): Map<string, Scalar> {
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

/**
 * Deliberately a separate decoder from `bind.ts`. The oracle should not be read
 * through the same code path as the thing it is judging.
 */
function decode(cell: RawCell, strings: string[]): Scalar | undefined {
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
      return cell.v === '1' || cell.v === 'true';
    case 'e':
      return { kind: cell.v } as unknown as Scalar;
    default: {
      const n = Number(cell.v);
      return Number.isFinite(n) ? n : cell.v;
    }
  }
}

function agrees(ours: Scalar, want: Scalar): boolean {
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

function show(v: Scalar | undefined): string {
  if (v === undefined) return '(absent)';
  if (v === null) return '(blank)';
  if (isErr(v)) return v.kind;
  if (typeof v === 'object') return String((v as { kind?: string }).kind ?? v);
  if (typeof v === 'string') return JSON.stringify(clip(v, 60));
  return String(v);
}

const clip = (s: string, n: number): string => (s.length <= n ? s : `${s.slice(0, n - 1)}…`);
const pct = (a: number, b: number): string => (b === 0 ? 'n/a' : `${((a / b) * 100).toFixed(1)}%`);

function firstFunction(formula: string): string | undefined {
  const m = /(?<![A-Za-z0-9_.])([A-Z][A-Z0-9._]*)\s*\(/.exec(formula);
  return m?.[1];
}
