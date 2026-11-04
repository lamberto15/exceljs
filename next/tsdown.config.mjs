import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: {
    index: 'src/index.ts', lite: 'src/lite.ts', csv: 'src/csv/index.ts', json: 'src/json/index.ts',
    crypto: 'src/crypto/index.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  platform: 'neutral',
  target: 'es2022',
  minify: false,
  sourcemap: false,
  clean: true,
  unbundle: true,
});
