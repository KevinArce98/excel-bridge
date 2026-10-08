import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    testTimeout: 30_000,
    typecheck: {
      enabled: true,
      include: ['tests/types/**/*.test-d.ts'],
      tsconfig: 'tests/types/tsconfig.json',
    },
  },
});
