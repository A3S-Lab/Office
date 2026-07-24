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
            from: './node_modules/@embedpdf/fonts-sc/fonts/NotoSansHans-Regular.otf',
            to: 'noto-sans-hans-regular.otf',
          },
          {
            from: './node_modules/@embedpdf/fonts-latin/fonts/NotoSans-Regular.ttf',
            to: 'noto-sans-regular.ttf',
          },
          {
            from: './node_modules/@embedpdf/fonts-arabic/fonts/NotoNaskhArabic-Regular.ttf',
            to: 'noto-naskh-arabic-regular.ttf',
          },
          {
            from: './node_modules/@embedpdf/fonts-hebrew/fonts/NotoSansHebrew-Regular.ttf',
            to: 'noto-sans-hebrew-regular.ttf',
          },
          {
            from: './node_modules/@embedpdf/fonts-sc/LICENSE',
            to: 'noto-sans-hans.LICENSE.txt',
          },
          {
            from: './node_modules/@embedpdf/fonts-latin/LICENSE',
            to: 'noto-sans.LICENSE.txt',
          },
          {
            from: './node_modules/@embedpdf/fonts-arabic/LICENSE',
            to: 'noto-naskh-arabic.LICENSE.txt',
          },
          {
            from: './node_modules/@embedpdf/fonts-hebrew/LICENSE',
            to: 'noto-sans-hebrew.LICENSE.txt',
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
