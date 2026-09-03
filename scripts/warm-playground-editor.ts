import { chromium } from 'playwright';

const baseUrl = process.env.A3S_TESTKIT_BASE_URL ?? 'http://127.0.0.1:3000';
const browserExecutable = process.env.AGENT_BROWSER_EXECUTABLE_PATH;

const editorTargets = [
  { id: 'project-brief', kind: 'document' },
  { id: 'quarterly-plan', kind: 'spreadsheet' },
  { id: 'strategy-deck', kind: 'presentation' },
  { id: 'blank-markdown', kind: 'markdown' },
] as const;

const pageUrl = `${baseUrl.replace(/\/$/, '')}/playground/`;

const browser = await chromium.launch(
  browserExecutable ? { executablePath: browserExecutable } : undefined,
);
const page = await browser.newPage({
  locale: 'zh-CN',
  viewport: { width: 1280, height: 800 },
});

try {
  for (const target of editorTargets) {
    const marker = `a3s-office.playground.${target.kind}.module-preload-ready`;
    let warmed = false;

    for (let attempt = 0; attempt < 3 && !warmed; attempt += 1) {
      await page.goto(pageUrl, { waitUntil: 'domcontentloaded' });
      const card = page.locator(`button[data-template-id="${target.id}"]`);
      await card.waitFor({ state: 'visible' });
      await card.hover();

      try {
        await page.waitForFunction(
          (entryName) => performance.getEntriesByName(entryName).length > 0,
          marker,
          { timeout: 20_000 },
        );
        warmed = true;
      } catch {
        // The first on-demand dev compilation can trigger an HMR reload. The
        // next attempt revisits the card after that compilation has settled.
      }
    }

    if (!warmed) {
      throw new Error(`Timed out warming Playground editor: ${target.kind}`);
    }
  }
} finally {
  await browser.close();
}
