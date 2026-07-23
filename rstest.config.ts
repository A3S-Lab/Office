import { withRslibConfig } from '@rstest/adapter-rslib';
import { defineConfig } from '@rstest/core';

export default defineConfig({
  environment: 'happy-dom',
  extends: withRslibConfig({ libId: 'library' }),
  setupFiles: ['./rstest.setup.ts'],
});
