import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

export default defineConfig({
  plugins: [react(), viteSingleFile()],
  resolve: { alias: { exceljs: 'exceljs/dist/exceljs.min.js' } },
  define: { 'process.env': {}, global: 'globalThis' },
  build: { target: 'es2020', chunkSizeWarningLimit: 4000, assetsInlineLimit: 100000000 },
});
