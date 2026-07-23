import { pluginReact } from '@rsbuild/plugin-react';
import { defineConfig } from '@rsbuild/core';

const assetPrefix = process.env.A3S_OFFICE_PLAYGROUND_BASE ?? '/';

export default defineConfig({
  html: {
    template: './index.html',
  },
  output: {
    assetPrefix,
    cleanDistPath: true,
    copy: [
      {
        from: '../node_modules/@embedpdf/pdfium/dist/pdfium.wasm',
        to: 'pdfium.wasm',
      },
      {
        from: '../node_modules/@embedpdf/pdfium/LICENSE.pdfium',
        to: 'pdfium.LICENSE.txt',
      },
      {
        from: '../node_modules/pptxgenjs/dist/pptxgen.bundle.js',
        to: 'pptxgen.bundle.js',
      },
      {
        from: '../node_modules/pptxgenjs/LICENSE',
        to: 'pptxgen.LICENSE.txt',
      },
    ],
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
