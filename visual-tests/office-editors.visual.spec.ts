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
      const editor = page.locator(
        '.ProseMirror[data-pagination-state="ready"]',
      );
      await editor.waitFor();
      await expect(editor).toHaveAttribute('data-pagination-engine', 'wasm');
      await expect(editor).toHaveAttribute(
        'data-pagination-text-engine',
        'wasm',
      );
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

test.describe('Office editor context menu contracts', () => {
  test.describe.configure({ mode: 'serial' });

  for (const kind of ['document', 'spreadsheet', 'presentation'] as const) {
    test(`${kind} context menu`, async ({ page }) => {
      const fixture = fixtures.find((candidate) => candidate.kind === kind);
      if (!fixture) throw new Error(`Missing ${kind} visual fixture.`);

      await page.goto('/');
      await fixture.open(page);
      await fixture.ready(page);
      await openEditorContextMenu(page, kind);

      const menu = page.getByRole('menu');
      await expect(menu).toBeVisible();
      await expect(menu.getByRole('menuitem').first()).toBeFocused();

      const geometry = await menu.evaluate((element) => {
        const rect = element.getBoundingClientRect();
        return {
          position: getComputedStyle(element).position,
          viewportWidth: document.documentElement.clientWidth,
          viewportHeight: document.documentElement.clientHeight,
          left: rect.left,
          top: rect.top,
          right: rect.right,
          bottom: rect.bottom,
          width: rect.width,
          height: rect.height,
        };
      });
      expect(geometry.position).toBe('fixed');
      expect(geometry.left).toBeGreaterThanOrEqual(8);
      expect(geometry.top).toBeGreaterThanOrEqual(8);
      expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
      expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
      expect(geometry.width).toBeGreaterThanOrEqual(200);
      expect(geometry.width).toBeLessThanOrEqual(320);
      expect(geometry.height).toBeGreaterThan(40);

      await page.keyboard.press('Escape');
      await expect(menu).toBeHidden();
    });
  }
});

test('document comments align with their review rail', async ({
  page,
}, testInfo) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await page
    .locator('.work-document-editable .ProseMirror')
    .evaluate((root) => {
      const paragraph = root.querySelectorAll('p')[1];
      const text = paragraph?.firstChild;
      if (!(text instanceof Text)) {
        throw new Error('Document comment fixture text is unavailable.');
      }
      const range = document.createRange();
      range.setStart(text, 0);
      range.setEnd(text, Math.min(12, text.length));
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
    });
  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '添加批注' }).click();
  const dialog = page.getByRole('dialog');
  await dialog
    .getByRole('textbox', { name: '批注内容' })
    .fill('这里需要补充可衡量的验收标准。');
  await dialog.getByRole('button', { name: '添加批注' }).click();

  const panel = page.getByRole('complementary', { name: '批注审阅' });
  const mark = page.locator('[data-document-comment]');
  const connector = page.locator('.work-document-comment-connectors path');
  await expect(panel).toBeVisible();
  await expect(mark).toHaveCount(1);
  await expect(mark).toHaveClass(/is-active-comment/);
  await expect(connector).toHaveCount(1);
  await expect(page.getByText('这里需要补充可衡量的验收标准。')).toBeVisible();

  const geometry = await page.evaluate(() => {
    const panel = document.querySelector<HTMLElement>(
      '.work-document-comments-panel',
    );
    const pageElement = document.querySelector<HTMLElement>(
      '.work-document-page',
    );
    if (!panel || !pageElement) {
      throw new Error('Document review geometry is unavailable.');
    }
    const panelRect = panel.getBoundingClientRect();
    const pageRect = pageElement.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      panelLeft: panelRect.left,
      panelRight: panelRect.right,
      panelWidth: panelRect.width,
      pageRight: pageRect.right,
    };
  });
  if (testInfo.project.name === 'compact-768') {
    expect(geometry.panelLeft).toBeGreaterThanOrEqual(
      geometry.viewportWidth - geometry.panelWidth - 16,
    );
  } else {
    expect(geometry.panelLeft).toBeGreaterThanOrEqual(geometry.pageRight - 1);
  }
  expect(geometry.panelRight).toBeLessThanOrEqual(geometry.viewportWidth + 1);
  expect(geometry.panelWidth).toBeGreaterThanOrEqual(270);

  await stabilizeVisualSurface(page);
  await expect(page).toHaveScreenshot('document-comments.png');
});

test('component guide provides framework-specific examples', async ({
  page,
}) => {
  await page.goto('/#guide');
  await expect(
    page.getByRole('heading', { name: '组件接入', level: 1 }),
  ).toBeVisible();
  await expect(page.getByRole('tab', { name: 'React' })).toBeVisible();

  await page.getByRole('tab', { name: 'Vue' }).click();
  await expect(page.locator('.playground-framework-example pre')).toContainText(
    'v-model:content',
  );

  await page.getByRole('tab', { name: 'Web Component' }).click();
  await expect(page.locator('.playground-framework-example pre')).toContainText(
    'defineA3SOfficeElements',
  );
});

test('PDF workspace card uses a single, legible file mark', async ({
  page,
}) => {
  await page.goto('/');
  const card = page.getByRole('button', {
    name: 'PDF 编辑器 查看、批注并保存 PDF',
  });
  const sheet = card.locator('.template-document-sheet');
  const mark = sheet.locator('.template-pdf-mark');

  await expect(card).toBeVisible();
  await expect(mark).toHaveText('PDF');
  await expect(sheet.locator('svg')).toHaveCount(0);
  await expect(mark).toHaveCSS('color', 'rgb(255, 255, 255)');
  await expect(mark).toHaveCSS('background-color', 'rgb(200, 64, 64)');
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

  if (kind !== 'pdf') {
    const startTab = page.getByRole('tab', { name: '开始' });
    await expect(startTab).toHaveAttribute('aria-selected', 'true');
    await expect(startTab).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
    await expect(
      page.getByRole('button', { name: '文件', exact: true }),
    ).toHaveCount(0);
  }
}

async function openEditorContextMenu(
  page: Page,
  kind: 'document' | 'spreadsheet' | 'presentation',
): Promise<void> {
  if (kind === 'document') {
    await page
      .locator('.work-document-editable .ProseMirror')
      .evaluate((editor) => {
        const text = editor.querySelector('h1')?.firstChild;
        if (!(text instanceof Text)) {
          throw new Error('Document context-menu text is unavailable.');
        }
        const range = document.createRange();
        range.setStart(text, 0);
        range.setEnd(text, Math.min(4, text.length));
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        (editor as HTMLElement).focus();
        document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
        editor.dispatchEvent(
          new MouseEvent('contextmenu', {
            bubbles: true,
            cancelable: true,
            clientX: 420,
            clientY: 520,
          }),
        );
      });
    return;
  }

  const target =
    kind === 'spreadsheet'
      ? page.locator('.work-spreadsheet-canvas')
      : page.locator('.work-slide-canvas.interactive');
  await target.click({
    button: 'right',
    position: kind === 'spreadsheet' ? { x: 280, y: 180 } : undefined,
  });
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
