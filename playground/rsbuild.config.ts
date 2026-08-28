import { pluginReact } from '@rsbuild/plugin-react';
import { defineConfig } from '@rsbuild/core';
import { resolve } from 'node:path';
import {
  playgroundBaseFromSiteBase,
  siteBaseFromEnvironment,
} from '../website/site-paths';

const assetPrefix = playgroundBaseFromSiteBase(siteBaseFromEnvironment());

export default defineConfig(({ command, env }) => {
  const testKitModule = resolve(import.meta.dirname, 'src/testkit.tsx');
  const testKitImplementation = resolve(
    import.meta.dirname,
    command === 'dev' || env === 'development'
      ? 'src/testkit-runtime.tsx'
      : 'src/testkit-noop.tsx',
  );

  return {
    html: {
      template: './index.html',
    },
    output: {
      assetPrefix,
      cleanDistPath: true,
      copy: [
        {
          from: '../node_modules/@embedpdf/pdfium/LICENSE.pdfium',
          to: 'pdfium.LICENSE.txt',
        },
        {
          from: '../node_modules/pptxgenjs/LICENSE',
          to: 'pptxgen.LICENSE.txt',
        },
        {
          from: './generated/a3s-office-skill.tar.gz',
          to: 'downloads/a3s-office-skill.tar.gz',
        },
        {
          from: '../crates/cli/skills/a3s-office',
          to: 'downloads/a3s-office-skill',
        },
      ],
      distPath: {
        root: '../playground-dist/playground',
      },
    },
    plugins: [pluginReact()],
    resolve: {
      alias: {
        [testKitModule]: testKitImplementation,
      },
    },
    root: import.meta.dirname,
    server: {
      host: '127.0.0.1',
    },
    source: {
      entry: {
        index: './src/main.tsx',
      },
    },
    tools: {
      rspack: {
        module: {
          rules: [{ test: /\.base64$/i, type: 'asset/source' }],
        },
      },
    },
  };
});
