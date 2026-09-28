import AxeBuilder from '@axe-core/playwright';
import { expect, type Page, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

type AxeEditorKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

interface AxeFixture {
  kind: AxeEditorKind;
  open: (page: Page) => Promise<void>;
  ready: (page: Page) => Promise<void>;
  /** Limit axe to editor chrome; playground shell contrast is out of scope. */
  include?: string[];
  /**
   * Regions that still lack native semantics by design until Phase 3.3/4
   * closes (canvas grid, slide stage). Chrome outside these nodes must still
   * be axe-clean; excluding them is scoped evidence, not a severity waiver.
   */
  exclude?: string[];
}

const fixtures: AxeFixture[] = [
  {
    kind: 'document',
    open: openDocumentFixture,
    ready: waitForDocumentFixture,
    include: ['.work-document-editor'],
  },
  {
    kind: 'markdown',
    open: (page) =>
      page
        .getByRole('button', {
          name: '# 产品说明 MD · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page
        .locator('.work-markdown-editor .ProseMirror')
        .waitFor({ state: 'attached' });
    },
    include: ['.work-markdown-editor'],
  },
  {
    kind: 'spreadsheet',
    open: (page) =>
      page
        .getByRole('button', {
          name: '季度执行计划 XLSX · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page
        .locator('.work-spreadsheet-canvas > .fortune-container')
        .waitFor();
    },
    include: ['.work-spreadsheet-editor'],
    // Fortune canvas has no grid ARIA until the A3S-owned virtual grid (Phase 4).
    exclude: ['.fortune-container', '.work-spreadsheet-canvas canvas'],
  },
  {
    kind: 'presentation',
    open: (page) =>
      page
        .getByRole('button', {
          name: '业务策略汇报 PPTX · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page.locator('.work-slide-canvas.interactive').waitFor();
    },
    include: ['.work-presentation-editor'],
    // Slide canvas remains a visual surface; SR/keyboard navigation uses the
    // Presentation object list beside the stage (Phase 3.3).
    exclude: ['.work-slide-canvas'],
  },
  {
    kind: 'pdf',
    open: openPdfFixture,
    ready: waitForPdfFixture,
    include: ['.work-pdf-viewer'],
    // EmbedPDF page rasters and its scroll host are out of chrome scope until PDF/UA.
    exclude: ['.work-pdf-embed', '.work-pdf-native-viewer'],
  },
];

test.describe('Office editor accessibility (axe)', () => {
  for (const fixture of fixtures) {
    test(`${fixture.kind} chrome has no serious or critical axe violations`, async ({
      page,
    }) => {
      await page.goto('/playground/');
      await fixture.open(page);
      await fixture.ready(page);

      let builder = new AxeBuilder({ page }).withTags([
        'wcag2a',
        'wcag2aa',
        'wcag21a',
        'wcag21aa',
      ]);
      for (const selector of fixture.include ?? []) {
        builder = builder.include(selector);
      }
      for (const selector of fixture.exclude ?? []) {
        builder = builder.exclude(selector);
      }

      const results = await builder.analyze();
      const blocking = results.violations.filter((violation) =>
        ['serious', 'critical'].includes(violation.impact ?? ''),
      );

      expect(
        blocking,
        blocking
          .map(
            (violation) =>
              `${violation.id} (${violation.impact}): ${violation.help}`,
          )
          .join('\n'),
      ).toEqual([]);
    });
  }
});
