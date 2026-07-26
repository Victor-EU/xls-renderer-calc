import { defineConfig } from 'vitest/config';
import { devAliases } from './tools/dev-aliases.js';

export default defineConfig({
  // The suite runs against `src`, not `dist`, so a failing test points at a line
  // you can edit. What that leaves uncovered — that the *published* artefact
  // resolves and imports — is covered by `npm run smoke` instead. See
  // tools/dev-aliases.ts.
  resolve: { alias: devAliases },
  test: {
    include: [
      'packages/*/src/**/*.test.ts',
      'packages/*/test/**/*.test.ts',
      'tools/oracle/*.test.ts',
      'eval/*.test.ts',
      // Matches nothing in a clone: the real-workbook corpus and its harness are
      // other people's confidential files and are not published. See the README.
      'eval/real/*.test.ts',
    ],
    environment: 'node',
  },
});
