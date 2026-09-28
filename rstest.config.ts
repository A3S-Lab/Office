import { resolve } from 'node:path';
import { withRslibConfig } from '@rstest/adapter-rslib';
import { defineConfig, type ProjectConfig } from '@rstest/core';

// Security-sensitive DOM code (HTML sanitization) needs a spec-compliant DOM;
// happy-dom silently lets DOMPurify return unsafe markup.
const JSDOM_TESTS = '**/*.jsdom.test.ts';

const shared: ProjectConfig = {
  extends: withRslibConfig({ libId: 'library' }),
  resolve: {
    alias: {
      '@a3s-lab/office/core': resolve(import.meta.dirname, 'src/core.ts'),
      '@a3s-lab/office/react': resolve(import.meta.dirname, 'src/react.tsx'),
    },
  },
  setupFiles: ['./rstest.setup.ts'],
  testTimeout: 10_000,
  tools: {
    rspack: {
      module: {
        rules: [{ test: /\.base64$/i, type: 'asset/source' }],
      },
    },
  },
};

export default defineConfig({
  coverage: {
    provider: 'v8',
  },
  pool: {
    maxWorkers: 2,
  },
  projects: [
    {
      ...shared,
      name: 'happy-dom',
      testEnvironment: 'happy-dom',
      exclude: ['visual-tests/**', JSDOM_TESTS],
    },
    {
      ...shared,
      name: 'jsdom',
      testEnvironment: 'jsdom',
      include: [JSDOM_TESTS],
      exclude: ['visual-tests/**'],
    },
  ],
});
