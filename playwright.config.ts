import { defineConfig } from '@playwright/test';

const externalBaseUrl = process.env.A3S_OFFICE_VISUAL_BASE_URL;
const baseURL = externalBaseUrl ?? 'http://127.0.0.1:4175';
const chromiumExecutablePath =
  process.env.A3S_OFFICE_VISUAL_CHROMIUM_EXECUTABLE;

export default defineConfig({
  testDir: './visual-tests',
  outputDir: './test-results/visual',
  fullyParallel: false,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: process.env.CI
    ? [['line'], ['html', { outputFolder: 'playwright-report', open: 'never' }]]
    : 'line',
  snapshotPathTemplate:
    '{testDir}/__snapshots__/{platform}/{projectName}/{arg}{ext}',
  expect: {
    timeout: 10_000,
    toHaveScreenshot: {
      animations: 'disabled',
      caret: 'hide',
      maxDiffPixelRatio: 0.005,
      scale: 'css',
      threshold: 0.2,
    },
  },
  use: {
    baseURL,
    colorScheme: 'light',
    deviceScaleFactor: 1,
    launchOptions: chromiumExecutablePath
      ? { executablePath: chromiumExecutablePath }
      : undefined,
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    screenshot: 'only-on-failure',
    timezoneId: 'Asia/Shanghai',
    trace: 'retain-on-failure',
  },
  projects: [
    {
      name: 'desktop-1280',
      use: { viewport: { width: 1280, height: 800 } },
    },
    {
      name: 'compact-768',
      use: { viewport: { width: 768, height: 800 } },
    },
  ],
  webServer: externalBaseUrl
    ? undefined
    : {
        command: 'bun run playground:preview',
        reuseExistingServer: !process.env.CI,
        timeout: 30_000,
        url: baseURL,
      },
});
