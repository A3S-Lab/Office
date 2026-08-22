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

## Traditional Office Writer page-layout parity

The deterministic Traditional Office layout contract is separate from the
ordinary UI snapshots. It generates a fixed DOCX fixture, exports a one-page
PDF through the locally installed reference Writer automation server, captures
the A3S and Traditional Office pages at 794 by 1123 CSS pixels, and compares
both semantic landmarks and bounded pixel differences. It never updates a
golden from the A3S output.

On Windows, build and serve the Playground, set a Chromium executable when a
Playwright-managed browser is unavailable, then run:

```powershell
$env:A3S_OFFICE_VISUAL_CHROMIUM_EXECUTABLE = 'C:\path\to\chrome.exe'
bun run playground:build
bun run playground:preview
bun run test:office-layout
```

The result bundle is written below `.a3s-test/` and includes the reference
Office version and hashes, both page PNGs, browser layout evidence, and the
comparison report. The gate requires a one-page A4 result, no browser errors,
at most a one-pixel landmark delta, no more than a 2% thresholded pixel
difference, and mean absolute channel error no greater than 1.0.
