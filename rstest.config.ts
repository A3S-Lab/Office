import { withRslibConfig } from '@rstest/adapter-rslib';
import { defineConfig } from '@rstest/core';
import { resolve } from 'node:path';

export default defineConfig({
  environment: 'happy-dom',
  extends: withRslibConfig({ libId: 'library' }),
  resolve: {
    alias: {
      '@a3s-lab/office/core': resolve(import.meta.dirname, 'src/core.ts'),
      '@a3s-lab/office/react': resolve(import.meta.dirname, 'src/react.tsx'),
    },
  },
  setupFiles: ['./rstest.setup.ts'],
});
