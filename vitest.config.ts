import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['packages/*/src/**/*.test.ts', 'packages/*/test/**/*.test.ts', 'tools/oracle/*.test.ts'],
    environment: 'node',
  },
});
