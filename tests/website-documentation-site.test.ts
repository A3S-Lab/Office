import { access, readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  DOCUMENTATION_DEFAULT_LANGUAGE,
  DOCUMENTATION_DEFAULT_VERSION,
  DOCUMENTATION_LOCALES,
  DOCUMENTATION_REQUIRED_ROUTES,
  DOCUMENTATION_VERSIONS,
} from '../website/documentation-site';

const documentationRoot = path.resolve(import.meta.dirname, '../docs');

function homepageComponentHref(version: string, component: string): string {
  const extension =
    version === 'latest' ||
    version === '0.13.0' ||
    version === '0.12.0' ||
    version === '0.11.0' ||
    version === '0.10.0' ||
    version === '0.9.2' ||
    version === '0.9.1' ||
    version === '0.9.0' ||
    version === '0.8.1' ||
    version === '0.8.0' ||
    version === '0.7.3'
      ? 'html'
      : 'mdx';
  return `./components/${component}.${extension}`;
}

test('uses Simplified Chinese and latest as stable documentation defaults', () => {
  expect(DOCUMENTATION_DEFAULT_LANGUAGE).toBe('zh');
  expect(DOCUMENTATION_LOCALES.map(({ lang }) => lang)).toEqual(['zh', 'en']);
  expect(DOCUMENTATION_DEFAULT_VERSION).toBe('latest');
  expect(DOCUMENTATION_VERSIONS).toEqual([
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
    '0.7.1',
    '0.7.0',
    '0.6.0',
    '0.5.0',
    '0.4.0',
    '0.3.0',
    '0.2.0',
    '0.1.0',
  ]);
});

test('keeps every public route available in every language and version', async () => {
  for (const version of DOCUMENTATION_VERSIONS) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      for (const route of DOCUMENTATION_REQUIRED_ROUTES) {
        await expect(
          access(path.join(documentationRoot, version, lang, route)),
        ).resolves.toBeUndefined();
      }
    }
  }
});

test('keeps published release homepages frozen and visibly versioned', async () => {
  for (const version of [
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
    '0.7.1',
    '0.7.0',
    '0.6.0',
    '0.5.0',
    '0.4.0',
    '0.3.0',
    '0.2.0',
    '0.1.0',
  ]) {
    const [chinese, english] = await Promise.all([
      readFile(path.join(documentationRoot, `${version}/zh/index.mdx`), 'utf8'),
      readFile(path.join(documentationRoot, `${version}/en/index.mdx`), 'utf8'),
    ]);

    expect(chinese).toContain(`# A3S Office ${version} 文档`);
    expect(chinese).toContain('冻结文档');
    expect(english).toContain(`# A3S Office ${version} documentation`);
    expect(english.toLowerCase()).toContain('frozen documentation');
  }
});

test('removes broken online Playground actions from frozen homepages', async () => {
  for (const version of DOCUMENTATION_VERSIONS.filter(
    (candidate) => candidate !== 'latest',
  )) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const homepage = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(homepage).not.toContain('<PlaygroundLink>');
      expect(homepage).not.toContain('在线体验');
      expect(homepage).not.toContain('Open the Playground');
    }
  }
});

test('uses deployable HTML targets in current release homepage actions', async () => {
  for (const version of [
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const homepage = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(homepage).not.toMatch(/href="[^"]+\.mdx(?:[?#][^"]*)?"/);
      expect(homepage).toContain('href="./guide/index.html"');
      expect(homepage).toContain('href="./components/collaboration.html"');
      expect(homepage).toContain(
        'href="./components/collaboration-server.html"',
      );
    }
  }
});

test('publishes real-time collaboration as a bilingual first-class capability', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
    '0.7.1',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [homepage, componentIndex, componentNavigation, collaboration] =
        await Promise.all([
          readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration.mdx'),
            'utf8',
          ),
        ]);

      expect(homepage).toContain(
        homepageComponentHref(version, 'collaboration'),
      );
      expect(componentIndex).toContain('./collaboration.mdx');
      expect(componentNavigation).toContain('"collaboration"');
      expect(collaboration).toContain('Yjs');
      expect(collaboration).toContain('Awareness');
      expect(collaboration).toContain('Spreadsheet');
      expect(collaboration).toContain('Presentation');
      expect(collaboration).toContain('PDF');
    }
  }
});

test('publishes the runnable collaboration backend in latest releases', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [homepage, componentIndex, componentNavigation, server] =
        await Promise.all([
          readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration-server.mdx'),
            'utf8',
          ),
        ]);

      expect(homepage).toContain(
        homepageComponentHref(version, 'collaboration-server'),
      );
      expect(componentIndex).toContain('./collaboration-server.mdx');
      expect(componentNavigation).toContain('"collaboration-server"');
      expect(server).toContain('a3s-boot');
      expect(server).toContain('WebSocket');
      expect(server).toContain('Yrs');
    }
  }
});

test('documents ephemeral native agent presence in current releases', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, server, cli]) {
        expect(source).toContain('--actor-name');
        expect(source).toContain('outbound-awareness');
        expect(source).toContain('receive-awareness');
        expect(source).toContain('peer-left');
      }
    }
  }
});

test('documents durable Document comments in current releases', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const document of [collaboration, cli]) {
        expect(document).toContain('document-comment-create');
        expect(document).toContain('document-comment-reply');
        expect(document).toContain('document-comment-set-resolved');
        expect(document).toContain('document-comment-delete');
      }
      expect(collaboration).toContain('startUtf16');
      expect(collaboration).toContain('expectedTextId');
      expect(server).toContain('actorName');
      expect(server).toContain('comment');
      expect(server).toContain('FORBIDDEN');
    }
  }
});

test('documents attributed Document suggestions and native typed mutations', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, document, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, document]) {
        expect(source).toContain("mode: 'suggest'");
      }
      expect(collaboration).toContain('document.change-decisions');
      expect(server).toContain('FORBIDDEN');
      expect(server).toContain('Document `suggest`');
      expect(cli).toContain('document-suggestion-create');
      expect(cli).toContain('document-suggestion-decide');
      expect(cli).toContain('v3');
      expect(cli).toContain('changeDecisions');
    }
  }
});

test('documents collaborative character-formatting revisions', async () => {
  for (const version of ['latest', '0.13.0', '0.12.0', '0.11.0', '0.10.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [document, collaboration, server] = await Promise.all([
        readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
      ]);

      for (const source of [document, collaboration, server]) {
        expect(source).toContain('formatting');
        expect(source).toContain('w:rPrChange');
      }
      expect(collaboration).toContain('changeKind: "formatting"');
      expect(server).toContain('document-suggestion-create');
      expect(server).toContain('document-suggestion-decide');
    }
  }
});

test('documents collaborative paragraph-formatting revisions', async () => {
  for (const version of ['latest', '0.13.0', '0.12.0', '0.11.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [document, collaboration, server, architecture, roadmap] =
        await Promise.all([
          readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration.mdx'),
            'utf8',
          ),
          readFile(
            path.join(localeRoot, 'components/collaboration-server.mdx'),
            'utf8',
          ),
          readFile(
            path.join(localeRoot, 'browser-editor-architecture.md'),
            'utf8',
          ),
          readFile(path.join(localeRoot, 'editor-quality-roadmap.md'), 'utf8'),
        ]);

      for (const source of [
        document,
        collaboration,
        server,
        architecture,
        roadmap,
      ]) {
        expect(source).toContain('paragraph-formatting');
        expect(source).toContain('w:pPrChange');
      }
      expect(collaboration).toContain('changeKind: "paragraph-formatting"');
      expect(collaboration).toContain('word-paragraph-formatting-revision.acl');
      expect(server).toContain('A3S Boot');
    }
  }
});

test('documents atomic native Spreadsheet cell batches', async () => {
  for (const version of [
    'latest',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, spreadsheet, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/spreadsheet.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, spreadsheet, cli]) {
        expect(source).toContain('spreadsheet-batch-cells');
        expect(source).toContain('nextCell: null');
      }
    }
  }
});

test('documents maximum sparse spreadsheets and cancellable imports in 0.12.0', async () => {
  for (const version of ['latest', '0.13.0', '0.12.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [guide, spreadsheet] = await Promise.all([
        readFile(path.join(localeRoot, 'guide/index.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/spreadsheet.mdx'), 'utf8'),
      ]);

      for (const value of [
        '1,048,576',
        '16,384',
        'dataValidationRanges',
        'cellProtectionRanges',
      ]) {
        expect(spreadsheet).toContain(value);
      }
      for (const value of ['AbortSignal', 'reading', 'finalizing']) {
        expect(guide).toContain(value);
      }
    }
  }
});

test('documents WPS four-direction cell fill in 0.13.0', async () => {
  for (const version of ['latest', '0.13.0']) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      'Fill Down',
      'Fill Right',
      'Fill Up',
      'Fill Left',
      'Cmd/Ctrl+D',
      'Cmd/Ctrl+R',
      '50,000 target cells',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('publishes reproducible 100k Document performance evidence', async () => {
  for (const version of ['latest', '0.13.0']) {
    const [document, architecture] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(
          documentationRoot,
          version,
          'en/browser-editor-architecture.md',
        ),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '100,000 text paragraphs',
      '100,000 table rows',
      '0.397 s',
      '120.0 FPS',
      '70.4 MiB',
      '53.9 ms',
    ]) {
      expect(document).toContain(evidence);
      expect(architecture).toContain(evidence);
    }
    for (const command of [
      'bun run performance:large-documents',
      'bun run performance:large-document-edits',
      'bun run test:e2e:large-documents',
    ]) {
      expect(document).toContain(command);
    }
    expect(document).toContain(
      '../browser-editor-architecture.md#current-100000-unit-evidence',
    );
  }
});

test('keeps the 0.8.0 native suggestion limitation frozen', async () => {
  for (const { lang } of DOCUMENTATION_LOCALES) {
    const cli = await readFile(
      path.join(documentationRoot, '0.8.0', lang, 'cli-reference.md'),
      'utf8',
    );

    expect(cli).toContain('document-suggestion-*');
    expect(cli).toContain('document-change-decision-*');
  }
});
