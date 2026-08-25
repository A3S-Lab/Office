import { expect, type Page } from '@playwright/test';

const visualDifferenceProbe =
  process.env.A3S_OFFICE_VISUAL_DIFFERENCE_PROBE === '1';

export async function openDocumentFixture(page: Page): Promise<void> {
  await page
    .getByRole('button', {
      name: '新项目方案 DOCX · 本次会话',
    })
    .click();
}

export async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}

export async function waitForDocumentFixture(page: Page): Promise<void> {
  const editor = page.locator('.ProseMirror[data-pagination-state="ready"]');
  await editor.waitFor();
  await expect(editor).toHaveAttribute('data-pagination-engine', 'wasm');
  await expect(editor).toHaveAttribute('data-pagination-text-engine', 'wasm');
}

export async function stabilizeVisualSurface(page: Page): Promise<void> {
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
