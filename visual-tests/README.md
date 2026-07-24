# Office visual contracts

This suite verifies the real Playground surfaces for Document, Markdown,
Spreadsheet, Presentation, and PDF at 1280 × 800 and 768 × 800 CSS pixels.
Every case checks shared-shell geometry, page-level overflow, primary
navigation, and a committed screenshot.

The baselines are grouped by operating system and generated with the Chromium
revision pinned by `@playwright/test`. CI and the GitHub Pages deployment use
the committed Linux baselines. Local macOS runs use the committed Darwin
baselines, so browser and operating-system text rasterization differences do
not hide real layout regressions.

Build the Playground and run the visual gate:

```bash
bun run playground:build
bunx playwright install chromium
bun run playground:visual
```

Only update baselines for an intentional, reviewed UI change:

```bash
bun run playground:visual:update
```

Review the ten changed PNG files for the current platform before accepting an
update. A baseline update is not a substitute for checking keyboard access,
responsive behavior, or editor-specific functional tests.
