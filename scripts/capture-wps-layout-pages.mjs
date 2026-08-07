import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { pathToFileURL } from 'node:url';
import { chromium } from '@playwright/test';

const { values } = parseArgs({
  options: {
    'a3s-output': { type: 'string' },
    'base-url': { type: 'string', default: 'http://127.0.0.1:4175' },
    chromium: { type: 'string' },
    docx: { type: 'string' },
    'layout-output': { type: 'string' },
    pdf: { type: 'string' },
    'wps-output': { type: 'string' },
  },
  strict: true,
});

const requiredOptions = [
  'a3s-output',
  'docx',
  'layout-output',
  'pdf',
  'wps-output',
];
for (const option of requiredOptions) {
  if (!values[option]) throw new Error(`Missing required --${option} option.`);
}

const baseUrl = values['base-url'];
const docxPath = path.resolve(values.docx);
const pdfPath = path.resolve(values.pdf);
const a3sOutputPath = path.resolve(values['a3s-output']);
const wpsOutputPath = path.resolve(values['wps-output']);
const layoutOutputPath = path.resolve(values['layout-output']);
const chromiumPath =
  values.chromium ??
  process.env.A3S_OFFICE_VISUAL_CHROMIUM_EXECUTABLE ??
  process.env.AGENT_BROWSER_EXECUTABLE_PATH;
const PDF_REFERENCE_EMBED_TIMEOUT_MS = 20_000;

await Promise.all(
  [a3sOutputPath, wpsOutputPath, layoutOutputPath].map((outputPath) =>
    mkdir(path.dirname(outputPath), { recursive: true }),
  ),
);

const browser = await chromium.launch({
  ...(chromiumPath ? { executablePath: path.resolve(chromiumPath) } : {}),
  args: ['--no-sandbox', '--allow-file-access-from-files'],
});

const browserErrors = [];
const captureErrors = (page) => {
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
};

try {
  const context = await browser.newContext({
    colorScheme: 'light',
    deviceScaleFactor: 1,
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    viewport: { width: 1400, height: 1400 },
  });

  const documentPage = await context.newPage();
  captureErrors(documentPage);
  await documentPage.goto(`${baseUrl}/?e2e=wps-layout-capture-document`);
  await documentPage
    .locator("input[type='file'][accept*='.docx']")
    .setInputFiles(docxPath);
  const editor = documentPage.locator(
    '.ProseMirror[data-pagination-state="ready"]',
  );
  await editor.waitFor();
  await setDocumentZoom(documentPage, 100);
  await settle(documentPage);

  const documentSheet = documentPage.locator('.work-document-page');
  const documentBox = await documentSheet.boundingBox();
  assertPageSize(documentBox, 'A3S document page');
  await documentSheet.screenshot({
    animations: 'disabled',
    caret: 'hide',
    path: a3sOutputPath,
  });

  const a3sLayout = await documentPage.evaluate(() => {
    const page = document.querySelector('.work-document-page');
    if (!(page instanceof HTMLElement)) {
      throw new Error('A3S document page was not found.');
    }
    const pageBox = page.getBoundingClientRect();
    return Array.from(
      page.querySelectorAll(
        '.work-document-section > p, .work-document-section > .tableWrapper, table tr',
      ),
    ).map((element) => {
      const box = element.getBoundingClientRect();
      const style = getComputedStyle(element);
      return {
        kind: element.matches('.tableWrapper')
          ? 'table'
          : element.tagName.toLowerCase(),
        text: element.textContent?.trim() ?? '',
        x: box.x - pageBox.x,
        y: box.y - pageBox.y,
        width: box.width,
        height: box.height,
        lineHeight: style.lineHeight,
        marginTop: style.marginTop,
        marginBottom: style.marginBottom,
      };
    });
  });
  const a3sRunLayout = await documentPage.evaluate(() => {
    const page = document.querySelector('.work-document-page');
    if (!(page instanceof HTMLElement)) {
      throw new Error('A3S document page was not found.');
    }
    const pageBox = page.getBoundingClientRect();
    return Array.from(
      page.querySelectorAll('[data-office-word-line-height-factor]'),
    ).map((element) => {
      const box = element.getBoundingClientRect();
      const style = getComputedStyle(element);
      return {
        text: element.textContent?.trim() ?? '',
        x: box.x - pageBox.x,
        y: box.y - pageBox.y,
        width: box.width,
        height: box.height,
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        lineHeight: style.lineHeight,
        position: style.position,
        top: style.top,
        wordLineHeightFactor:
          element.getAttribute('data-office-word-line-height-factor') ?? '',
        wordSnapToGrid:
          element.getAttribute('data-office-word-snap-to-grid') ?? '',
      };
    });
  });

  const pdfPage = await context.newPage();
  captureErrors(pdfPage);
  await pdfPage.goto(`${baseUrl}/?e2e=wps-layout-capture-pdf`);
  await pdfPage
    .locator("input[type='file'][accept='.pdf,application/pdf']")
    .setInputFiles(pdfPath);
  let referenceRenderer = 'embedpdf';
  let referenceFallback;
  let wpsBox;
  try {
    await pdfPage.waitForFunction(
      () =>
        document.querySelector('.work-pdf-embed[data-ready="true"]') ||
        document.querySelector('.work-pdf-state[role="alert"]'),
      undefined,
      { timeout: PDF_REFERENCE_EMBED_TIMEOUT_MS },
    );
    if (await pdfPage.locator('.work-pdf-state[role="alert"]').count()) {
      throw new Error('The embedded PDF viewer rejected the WPS reference.');
    }
    await pdfPage.evaluate(async () => {
      const host = document.querySelector('embedpdf-container');
      if (!host) throw new Error('PDF viewer host was not found.');
      const registry = await host.registry;
      const documentId = registry
        .getPlugin('document-manager')
        .provides()
        .getActiveDocumentId();
      const zoom = registry
        .getPlugin('zoom')
        .provides()
        .forDocument(documentId);
      zoom.requestZoom(1.334);
      await new Promise((resolve) => setTimeout(resolve, 1_500));
      const renderedPage = Array.from(
        host.shadowRoot.querySelectorAll('div'),
      ).find((element) => {
        const box = element.getBoundingClientRect();
        return (
          element.getAttribute('style')?.includes('touch-action: none') &&
          box.width > 790 &&
          box.width < 800 &&
          box.height > 1_120 &&
          box.height < 1_125
        );
      });
      if (!renderedPage)
        throw new Error('Rendered WPS PDF page was not found.');
      renderedPage.dataset.wpsReferencePage = 'true';
    });
    const wpsSheet = pdfPage.locator('[data-wps-reference-page="true"]');
    wpsBox = await wpsSheet.boundingBox();
    assertPageSize(wpsBox, 'WPS reference page');
    await pdfPage.screenshot({
      animations: 'disabled',
      caret: 'hide',
      clip: {
        x: wpsBox.x,
        y: wpsBox.y,
        width: 794,
        height: 1123,
      },
      path: wpsOutputPath,
    });
  } catch (error) {
    const diagnostics = await pdfPage.evaluate(() => ({
      bodyText: document.body.innerText.slice(-2_000),
      state: document
        .querySelector('.work-pdf-state')
        ?.outerHTML.slice(0, 2_000),
    }));
    referenceRenderer = 'chromium-native-pdf';
    referenceFallback = {
      cause: error instanceof Error ? error.message : String(error),
      diagnostics,
    };
    console.warn(
      `Embedded PDF capture failed; using Chromium native PDF rendering: ${JSON.stringify(referenceFallback)}`,
    );
    wpsBox = await captureNativePdfReference(browser, pdfPath, wpsOutputPath);
  }

  await writeFile(
    layoutOutputPath,
    `${JSON.stringify(
      {
        a3sPage: documentBox,
        wpsPage: wpsBox,
        referenceRenderer,
        ...(referenceFallback ? { referenceFallback } : {}),
        a3sLayout,
        a3sRunLayout,
        pagination: {
          engine: await editor.getAttribute('data-pagination-engine'),
          pages: await editor.getAttribute('data-pagination-pages'),
          textEngine: await editor.getAttribute('data-pagination-text-engine'),
        },
        browserErrors,
      },
      null,
      2,
    )}\n`,
    'utf8',
  );
  if (browserErrors.length) {
    throw new Error(
      `Browser errors were captured: ${browserErrors.join(' | ')}`,
    );
  }
} finally {
  await browser.close();
}

async function captureNativePdfReference(browser, pdfFile, outputPath) {
  const nativeContext = await browser.newContext({
    colorScheme: 'light',
    deviceScaleFactor: 1,
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    viewport: { width: 1200, height: 1400 },
  });
  try {
    const page = await nativeContext.newPage();
    await page.goto(pathToFileURL(pdfFile).href, {
      timeout: 30_000,
      waitUntil: 'load',
    });
    await page.waitForTimeout(2_000);
    await page.mouse.click(31, 28);
    await page.waitForTimeout(1_000);
    const pageBox = { x: 203, y: 58, width: 794, height: 1123 };
    await page.screenshot({
      animations: 'disabled',
      caret: 'hide',
      clip: pageBox,
      path: outputPath,
    });
    return pageBox;
  } finally {
    await nativeContext.close();
  }
}

async function setDocumentZoom(page, target) {
  const slider = page.locator('.work-document-footer [role="slider"]').last();
  await slider.waitFor();
  await slider.focus();
  for (let attempt = 0; attempt < 40; attempt += 1) {
    const current = Number(await slider.getAttribute('aria-valuenow'));
    if (current === target) return;
    if (!Number.isFinite(current)) {
      throw new Error('Document zoom does not expose aria-valuenow.');
    }
    await slider.press(current < target ? 'ArrowRight' : 'ArrowLeft');
  }
  throw new Error(`Unable to set document zoom to ${target}.`);
}

async function settle(page) {
  await page.evaluate(async () => {
    await document.fonts.ready;
    await new Promise((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(resolve)),
    );
  });
}

function assertPageSize(box, label) {
  if (!box) throw new Error(`${label} could not be measured.`);
  if (Math.abs(box.width - 794) > 1 || Math.abs(box.height - 1123) > 1) {
    throw new Error(
      `${label} is ${box.width}x${box.height}; expected an A4 794x1123 CSS pixel page.`,
    );
  }
}
