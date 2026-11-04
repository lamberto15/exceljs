import { defineConfig } from 'vitest/config';

// CI runners (Windows especially) are 3-6x slower than local machines on fixture-heavy tests.
export default defineConfig({ test: { include: ['test/**/*.test.ts'], testTimeout: 30_000 } });
