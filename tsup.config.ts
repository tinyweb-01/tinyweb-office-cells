import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs', 'esm'],
  dts: true,
  clean: true,
  sourcemap: true,
  external: ['jszip', 'fast-xml-parser'],
  outDir: 'dist',
  splitting: false,
  treeshake: true,
});
