import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [react()],
  // The packages are consumed as TypeScript sources rather than build output, so
  // a change in the engine shows up in the app on the next hot reload.
  optimizeDeps: { exclude: ['@xlscalc/formula-engine', '@xlscalc/xlsx-preview'] },
});
