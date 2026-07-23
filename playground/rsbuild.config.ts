import { pluginReact } from '@rsbuild/plugin-react';
import { defineConfig } from '@rsbuild/core';

export default defineConfig({
  html: {
    template: './index.html',
  },
  output: {
    cleanDistPath: true,
    distPath: {
      root: '../playground-dist',
    },
  },
  plugins: [pluginReact()],
  root: import.meta.dirname,
  source: {
    entry: {
      index: './src/main.tsx',
    },
  },
});
