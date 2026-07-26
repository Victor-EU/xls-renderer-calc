/**
 * Did the oracle and the corpus actually run, or did they quietly not exist?
 *
 * Both suites are wrapped in `describe.skipIf(!haveFixtures)`, which is the right
 * behaviour for `npm test` on a machine without LibreOffice — but it means a
 * missing fixture directory produces a *green run with zero tests in it*, not a
 * failure. On a release that is the difference between a gate and a formality:
 * the whole claim of this project is that its numbers are checked against
 * something, and the two harnesses that check them are exactly the ones that
 * vanish silently.
 *
 * Note what a skipped suite looks like in the report: `numTotalTests: 0`, and
 * `numPendingTests: 0` as well. It does not register as "skipped" — it registers
 * as nothing at all, which is why this checks per file for tests that ran rather
 * than checking a pending count.
 *
 *   node scripts/check-harnesses-ran.mjs harness.json
 */

import { readFileSync } from 'node:fs';

const REQUIRED = [
  ['tools/oracle/oracle.test.ts', 'the LibreOffice oracle'],
  ['eval/eval.test.ts', 'the ten-workbook synthetic corpus'],
];

const report = JSON.parse(readFileSync(process.argv[2] ?? 'harness.json', 'utf8'));
let failed = false;

for (const [file, what] of REQUIRED) {
  const result = report.testResults?.find((r) => r.name.replace(/\\/g, '/').endsWith(file));
  const tests = result?.assertionResults ?? [];
  const passed = tests.filter((t) => t.status === 'passed').length;

  if (passed === 0) {
    console.error(`::error::${what} ran no tests — its fixtures were not built, so it skipped`);
    failed = true;
  } else {
    console.log(`  ok  ${what}: ${passed} of ${tests.length} passed`);
  }
}

if (report.numFailedTests > 0) {
  console.error(`::error::${report.numFailedTests} harness tests failed`);
  failed = true;
}

process.exit(failed ? 1 : 0);
