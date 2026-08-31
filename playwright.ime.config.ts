import { defineConfig } from '@playwright/test';

const externalBaseUrl = process.env.A3S_OFFICE_VISUAL_BASE_URL;
const baseURL = externalBaseUrl ?? 'http://127.0.0.1:4175';

export default defineConfig({
  testDir: './visual-tests',
  testMatch: 'controlled-ime.webkit.spec.ts',
  timeout: 60_000,
  outputDir: './test-results/ime-webkit',
  fullyParallel: false,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: process.env.CI
    ? [
        ['line'],
        [
          'html',
          { outputFolder: 'playwright-report-ime-webkit', open: 'never' },
        ],
      ]
    : 'line',
  use: {
    baseURL,
    browserName: 'webkit',
    colorScheme: 'light',
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    screenshot: 'only-on-failure',
    timezoneId: 'Asia/Shanghai',
    trace: 'retain-on-failure',
    viewport: { width: 1280, height: 800 },
  },
  webServer: externalBaseUrl
    ? undefined
    : {
        command: 'bun run playground:preview',
        reuseExistingServer: !process.env.CI,
        timeout: 30_000,
        url: baseURL,
      },
});
