import { expect, type Locator, type Page } from '@playwright/test';

const visualDifferenceProbe =
  process.env.A3S_OFFICE_VISUAL_DIFFERENCE_PROBE === '1';

/** Choose an OfficeSelect option by `data-value` or accessible name. */
export async function chooseOfficeSelectOption(
  page: Page,
  combobox: Locator,
  choice: string | { label: string | RegExp },
): Promise<void> {
  await combobox.click();
  if (typeof choice === 'object') {
    await page.getByRole('option', { name: choice.label }).click();
    return;
  }
  const byValue = page.locator(`[role="option"][data-value="${choice}"]`);
  try {
    await byValue.first().waitFor({ state: 'visible', timeout: 2_000 });
    await byValue.first().click();
  } catch {
    await page.getByRole('option', { name: choice }).click();
  }
}

export async function expectOfficeSelectValue(
  combobox: Locator,
  value: string,
): Promise<void> {
  await expect(combobox).toHaveAttribute('data-selected-value', value);
}

/** Close an open OfficeSelect listbox by toggling its trigger (avoids Escape closing dialogs). */
export async function closeOfficeSelect(combobox: Locator): Promise<void> {
  if ((await combobox.getAttribute('aria-expanded')) === 'true') {
    await combobox.click();
  }
}

/** Set an OfficeColorPicker value through the custom hex field. */
export async function chooseOfficeColor(
  page: Page,
  trigger: Locator,
  color: string,
): Promise<void> {
  await trigger.click();
  const custom = page.getByRole('textbox', { name: '自定义颜色值' });
  await custom.fill(color);
  await page.getByRole('button', { name: '应用自定义颜色' }).click();
}

export async function expectOfficeColorValue(
  trigger: Locator,
  value: string,
): Promise<void> {
  await expect(trigger).toHaveAttribute('data-selected-value', value);
}

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
