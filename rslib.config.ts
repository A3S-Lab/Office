import { pluginReact } from '@rsbuild/plugin-react';
import { defineConfig } from '@rslib/core';

export default defineConfig({
  source: {
    entry: {
      core: './src/core.ts',
      index: './src/index.ts',
      react: './src/react.tsx',
      styles: './src/styles/index.css',
      vue: './src/vue.ts',
      'web-component': './src/web-component.tsx',
    },
  },
  lib: [
    {
      dts: true,
      format: 'esm',
    },
  ],
  output: {
    copy: [
      {
        from: './node_modules/@embedpdf/pdfium/dist/pdfium.wasm',
        to: 'pdfium.wasm',
      },
      {
        from: './node_modules/@embedpdf/pdfium/LICENSE.pdfium',
        to: 'pdfium.LICENSE.txt',
      },
      {
        from: './node_modules/pptxgenjs/dist/pptxgen.bundle.js',
        to: 'pptxgen.bundle.js',
      },
      {
        from: './node_modules/pptxgenjs/LICENSE',
        to: 'pptxgen.LICENSE.txt',
      },
    ],
    target: 'web',
  },
  plugins: [pluginReact()],
});
