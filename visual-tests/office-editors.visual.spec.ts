import { expect, test, type Page } from '@playwright/test';
import { jsPDF } from 'jspdf';

const visualDifferenceProbe =
  process.env.A3S_OFFICE_VISUAL_DIFFERENCE_PROBE === '1';

type VisualEditorKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

interface VisualFixture {
  kind: VisualEditorKind;
  open: (page: Page) => Promise<void>;
  ready: (page: Page) => Promise<void>;
}

const fixtures: VisualFixture[] = [
  {
    kind: 'document',
    open: (page) =>
      page
        .getByRole('button', {
          name: '新项目方案 DOCX · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page
        .locator('.ProseMirror[data-pagination-state="ready"]')
        .waitFor();
    },
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
      await page.locator('.work-markdown-editor .ProseMirror').waitFor();
    },
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
      await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
        'data-presentation-geometry-state',
        'idle',
      );
    },
  },
  {
    kind: 'pdf',
    open: async (page) => {
      await page
        .locator('input[aria-label="打开 Office 或 PDF 文件"]')
        .setInputFiles({
          name: 'visual-fixture.pdf',
          mimeType: 'application/pdf',
          buffer: visualPdf(),
        });
    },
    ready: async (page) => {
      await page.locator('.work-pdf-embed[data-ready="true"]').waitFor({
        timeout: 30_000,
      });
    },
  },
];

test.describe('Office editor visual contracts', () => {
  test.describe.configure({ mode: 'serial' });

  for (const fixture of fixtures) {
    test(`${fixture.kind} editor`, async ({ page }, testInfo) => {
      await page.goto('/');
      await fixture.open(page);
      await fixture.ready(page);
      await stabilizeVisualSurface(page);
      await verifySharedEditorGeometry(
        page,
        fixture.kind,
        testInfo.project.name,
      );
      await expect(page).toHaveScreenshot(`${fixture.kind}.png`);
    });
  }
});

async function stabilizeVisualSurface(page: Page): Promise<void> {
  await page.addStyleTag({
    content: `
      *,
      *::before,
      *::after {
        caret-color: transparent !important;
        scroll-behavior: auto !important;
      }

      * {
        scrollbar-width: none !important;
      }

      *::-webkit-scrollbar,
      .playground-toast {
        display: none !important;
      }

      ${
        visualDifferenceProbe
          ? `
            html::after {
              content: '';
              position: fixed;
              inset: 0;
              z-index: 2147483647;
              box-sizing: border-box;
              border: 12px solid #ff00ff;
              pointer-events: none;
            }
          `
          : ''
      }
    `,
  });
  await page.evaluate(async () => {
    await document.fonts.ready;
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });
  });
}

async function verifySharedEditorGeometry(
  page: Page,
  kind: VisualEditorKind,
  projectName: string,
): Promise<void> {
  const shell = page.locator(`.work-editor-shell.${kind}`);
  const header = shell.locator('.work-editor-header');
  const editorHost = shell.locator('.playground-editor-host');
  await expect(shell).toBeVisible();
  await expect(header).toBeVisible();
  await expect(editorHost).toBeVisible();
  await expect(
    page.getByRole('button', { name: '返回办公首页' }),
  ).toBeVisible();
  await expect(
    page.getByRole('button', { name: '打开 AI 助手' }),
  ).toBeVisible();

  const geometry = await page.evaluate(() => {
    const shell = document.querySelector<HTMLElement>('.work-editor-shell');
    const header = shell?.querySelector<HTMLElement>('.work-editor-header');
    const host = shell?.querySelector<HTMLElement>('.playground-editor-host');
    if (!shell || !header || !host) {
      throw new Error('The shared Office editor shell is incomplete.');
    }
    const shellRect = shell.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const hostRect = host.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      documentScrollWidth: document.documentElement.scrollWidth,
      shell: {
        left: shellRect.left,
        top: shellRect.top,
        width: shellRect.width,
        height: shellRect.height,
      },
      header: {
        left: headerRect.left,
        top: headerRect.top,
        width: headerRect.width,
        height: headerRect.height,
      },
      host: {
        left: hostRect.left,
        top: hostRect.top,
        width: hostRect.width,
        height: hostRect.height,
      },
    };
  });

  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.viewportWidth + 1,
  );
  expect(geometry.shell.left).toBeCloseTo(0, 0);
  expect(geometry.shell.top).toBeCloseTo(0, 0);
  expect(geometry.shell.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.header.left).toBeCloseTo(0, 0);
  expect(geometry.header.top).toBeCloseTo(0, 0);
  expect(geometry.header.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.header.height).toBeGreaterThanOrEqual(48);
  expect(geometry.header.height).toBeLessThanOrEqual(56);
  expect(geometry.host.left).toBeCloseTo(0, 0);
  expect(geometry.host.top).toBeCloseTo(geometry.header.height, 0);
  expect(geometry.host.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.host.height).toBeGreaterThan(600);

  if (projectName === 'compact-768') {
    await expect(
      page.getByRole('button', { name: '展开办公侧边栏' }),
    ).toBeVisible();
  }
}

function visualPdf(): Buffer {
  const pdf = new jsPDF({
    compress: true,
    format: 'a4',
    orientation: 'portrait',
    unit: 'pt',
  });
  pdf.setCreationDate(new Date('2026-01-01T00:00:00.000Z'));
  pdf.setFileId('A3S0FF1CE00000000000000000000001');
  pdf.setProperties({
    author: 'A3S Lab',
    creator: 'A3S Office visual tests',
    subject: 'Deterministic PDF editor fixture',
    title: 'A3S Office',
  });
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(24);
  pdf.text('A3S Office', 72, 96);
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.text('PDF editor visual fixture', 72, 124);
  pdf.setDrawColor(40, 103, 216);
  pdf.setFillColor(238, 244, 255);
  pdf.roundedRect(72, 158, 451, 92, 8, 8, 'FD');
  pdf.setTextColor(34, 52, 82);
  pdf.text(
    'Typed toolbar, PDFium canvas, annotations, search, and save.',
    92,
    194,
  );
  pdf.text(
    'This page is generated in memory by the visual regression test.',
    92,
    218,
  );
  return Buffer.from(pdf.output('arraybuffer'));
}
