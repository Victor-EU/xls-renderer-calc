import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { implementedFunctions, installStandardLibrary, refusedFunctions } from './functions/index.js';

/**
 * The published capability floor, kept in `CAPABILITY.md` at the repo root.
 *
 * It is generated from the registry rather than written by hand, because a
 * hand-maintained list of what an engine supports drifts within a week and then
 * actively misleads — which for a fail-loud engine is the one thing the
 * documentation must not do.
 *
 * It also has a job beyond documentation. Because we control the generator that
 * writes these workbooks, this list can be published into its prompt as an
 * allowlist, so the renderer never meets a formula it cannot compute *by
 * construction*. Widening the floor is then a deliberate act — implement,
 * oracle-test, regenerate — rather than a bug report from a user staring at a ⚠.
 *
 *   npm test                        asserts the file matches the registry
 *   UPDATE_CAPABILITY=1 npm test    regenerates it
 */

const HERE = dirname(fileURLToPath(import.meta.url));
const PATH = join(HERE, '..', '..', '..', 'CAPABILITY.md');

function render(): string {
  installStandardLibrary();
  const implemented = implementedFunctions();
  const refused = refusedFunctions();
  const wrap = (names: string[]): string => {
    const lines: string[] = [];
    for (let i = 0; i < names.length; i += 8) lines.push(names.slice(i, i + 8).join(' '));
    return lines.join('\n');
  };

  return `# Capability floor

Generated from the function registry by \`capability.test.ts\`. Do not edit by
hand — run \`UPDATE_CAPABILITY=1 npm test\`.

Anything outside this list renders ⚠ with a stated reason. It is never guessed
at, and never silently treated as zero.

## Implemented (${implemented.length})

\`\`\`
${wrap(implemented)}
\`\`\`

## Refused, deliberately (${refused.length})

These are recognised Excel functions the engine declines to evaluate. The
distinction from an unknown name matters for the coverage backlog: an unknown
name might be a typo in the generated model, while a refusal is our own gap with
a stated reason.

\`\`\`
${wrap(refused)}
\`\`\`

Broadly they fall into three groups:

- **Dynamic references** — \`INDIRECT\`, \`OFFSET\`. They build references at
  evaluation time, which a static dependency graph cannot express.
- **Spilling results** — \`FILTER\`, \`SORT\`, \`UNIQUE\`, \`SEQUENCE\`, \`MMULT\`,
  \`TREND\` and friends. They change the shape of the grid rather than one cell's
  value. A saved file already contains whatever the writer spilled, so refusing
  costs nothing and prevents a half-right render.
- **State the engine cannot see** — \`AGGREGATE\` (hidden rows), \`CELL\`/\`INFO\`
  (workbook and environment), \`GETPIVOTDATA\` (pivot cache), \`WEBSERVICE\`/\`RTD\`
  (network).

Two further refusals are conditional rather than whole-function, so they do not
appear in the list above:

- \`SUBTOTAL(101–111)\` — the hidden-row variants only; 1–11 are implemented.
- **Approximate-match lookups over unsorted data** (\`VLOOKUP\`/\`HLOOKUP\` with the
  4th argument omitted or TRUE, \`MATCH\` type 1 or -1). The sorted contract is
  implemented exactly; unsorted input is refused, because Excel answers it from
  its binary search's probe order and that answer is arbitrary.
`;
}

describe('capability floor', () => {
  it('matches CAPABILITY.md', () => {
    const expected = render();
    if (process.env.UPDATE_CAPABILITY === '1' || !existsSync(PATH)) {
      writeFileSync(PATH, expected);
      return;
    }
    expect(readFileSync(PATH, 'utf8'), 'run UPDATE_CAPABILITY=1 npm test to regenerate').toBe(
      expected,
    );
  });
});
