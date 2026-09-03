import { defineConfig } from '@playwright/test';

const externalBaseUrl = process.env.A3S_OFFICE_TESTKIT_BASE_URL;
const baseURL = externalBaseUrl ?? 'http://127.0.0.1:3000';
const chromiumExecutablePath =
  process.env.A3S_OFFICE_TESTKIT_CHROMIUM_EXECUTABLE;

export default defineConfig({
  testDir: './visual-tests',
  testMatch: '**/office-testkit-ui.functional.spec.ts',
  timeout: 120_000,
  outputDir: './test-results/testkit',
  fullyParallel: false,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: process.env.CI ? [['line'], ['html', { open: 'never' }]] : 'line',
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
      name: 'testkit-desktop',
      use: { viewport: { width: 1280, height: 800 } },
    },
    {
      name: 'testkit-compact',
      use: { viewport: { width: 768, height: 800 } },
    },
  ],
  webServer: externalBaseUrl
    ? undefined
    : {
        command: 'bun run playground',
        reuseExistingServer: !process.env.CI,
        timeout: 120_000,
        url: `${baseURL}/playground/`,
      },
});
