/**
 * Where `@xlscalc/*` resolves to *inside this repository*.
 *
 * Published, these packages resolve through their `exports` map to `dist/`.
 * That is the path consumers take and the one `tools/verify/package-smoke.mjs`
 * exercises against real Node ESM resolution.
 *
 * In-repo we deliberately take a different path — straight to `src` — so a
 * change in the engine shows up in the demo on the next hot reload and in
 * `npm test` without a build step in between. The cost of that convenience is
 * that neither the demo nor the test suite proves the published artefact works,
 * which is exactly why the smoke test exists and runs in `npm run verify`.
 */

import { fileURLToPath } from 'node:url';

const at = (path: string): string => fileURLToPath(new URL(`../${path}`, import.meta.url));

export const devAliases = [
  {
    find: /^@xlscalc\/formula-engine$/,
    replacement: at('packages/formula-engine/src/index.ts'),
  },
  {
    find: /^@xlscalc\/xlsx-preview$/,
    replacement: at('packages/xlsx-preview/src/index.ts'),
  },
  {
    find: /^@xlscalc\/xlsx-preview\/view$/,
    replacement: at('packages/xlsx-preview/src/view/index.ts'),
  },
  {
    find: /^@xlscalc\/xlsx-preview\/view\/style\.css$/,
    replacement: at('packages/xlsx-preview/src/view/style.css'),
  },
  {
    find: /^@xlscalc\/xlsx-preview\/worker$/,
    replacement: at('packages/xlsx-preview/src/worker/index.ts'),
  },
  {
    find: /^@xlscalc\/xlsx-preview\/worker\/entry$/,
    replacement: at('packages/xlsx-preview/src/worker/entry.ts'),
  },
];
