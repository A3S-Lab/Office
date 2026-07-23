import { pluginReact } from '@rsbuild/plugin-react';
import { defineConfig } from '@rslib/core';

export default defineConfig({
  lib: [
    {
      dts: true,
      format: 'esm',
      id: 'library',
      output: {
        copy: [
          {
            from: './node_modules/@embedpdf/pdfium/dist/pdfium.wasm',
            to: 'pdfium.wasm',
          },
          {
            from: './generated/office-kernel.wasm',
            to: 'office-kernel.wasm',
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
    },
    {
      autoExternal: false,
      dts: false,
      format: 'iife',
      id: 'office-kernel-worker',
      output: {
        target: 'web',
      },
      source: {
        entry: {
          'office-kernel.worker':
            './src/internal/kernel/office-kernel.worker.ts',
        },
      },
    },
  ],
  plugins: [pluginReact()],
});
