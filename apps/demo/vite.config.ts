import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';
import { devAliases } from '../../tools/dev-aliases.js';

export default defineConfig({
  plugins: [react()],
  // The demo consumes the packages as TypeScript sources rather than build
  // output, so a change in the engine shows up on the next hot reload. It
  // imports them by their *published* specifiers all the same — including
  // `@xlscalc/xlsx-preview/view/style.css` — so the demo is written exactly as
  // a consumer's app would be, and only the resolution differs.
  resolve: { alias: devAliases },
  optimizeDeps: { exclude: ['@xlscalc/formula-engine', '@xlscalc/xlsx-preview'] },
});
