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

## Focused WebKit IME contract

The controlled rich-text regression uses the Playwright-pinned WebKit engine
without expanding the screenshot matrix. It dispatches a complete composition
lifecycle in `DocumentEditor`, proves Pinyin pre-edit changes are not published,
expects exactly one committed Chinese publication, and reopens the controlled
fixture to verify persistence:

```bash
bunx playwright install webkit
bun run playground:ime:webkit
```

CI runs this focused gate separately from the Chromium visual contracts.

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

## Focused Writer text-effect contract

The native outline, shadow, emboss, and imprint functional test opens the
public **文字效果** template, checks the four real CSS projections, and selects
the valid outline-shadow pair. It then verifies independently checked advanced
font controls, conflict-safe emboss authoring, live preview, one-step Undo,
focus restoration, and retained Worker/WASM layout eligibility at desktop and
390 px widths:

```bash
bun run playground:visual:document-legacy-text-effects
```

This Playwright test owns visual evidence. The companion A3S Test ACL remains a
local semantic and diagnostics gate and is intentionally absent from GitHub
Actions and Pages.

## Focused Spreadsheet multi-key sort contract

The functional regression opens the public quarterly-pipeline workbook,
captures a seven-column range, configures two sort levels, and verifies stable
row movement, relative-formula translation, one-step Undo, invoker/grid focus,
accessibility, and empty console and page-error diagnostics at desktop and
compact widths:

```bash
bun run playground:visual:spreadsheet-custom-sort
```

The companion A3S Test ACL exercises the same public workflow through the
local release gate and remains intentionally absent from GitHub Actions and
Pages.

## Focused Spreadsheet custom-list sort contract

The functional regression uses the public quarterly workbook to author and
locally save a bounded status sequence, apply it as the primary key, keep a
value-based secondary key, and verify formula translation plus one-step Undo.
It then opens the responsive preference manager, verifies the read-only
built-ins and initial keyboard focus, edits one sequence, creates and reorders
another, deletes a staged list, and remounts the editor to prove the exact final
preference order persisted. Playwright runs the same path at desktop and compact
widths, captures the manager dialog, and rejects browser console or page errors:

```bash
bun run playground:visual:spreadsheet-custom-list-sort
```

The companion local-only A3S Test ACL additionally rejects a duplicate list,
manages stored sequences, reloads the whole Playground, verifies the saved
preference order, and records interactive accessibility evidence. Persistence
is explicit through the Playground's typed local store and never modifies
workbook content.

## Focused Spreadsheet appearance sort contract

The functional regression opens a workbook with direct cell and font colors
plus conditional-format icons, applies all three appearance kinds across four
levels, and verifies same-column color priority, relative-formula translation,
one-step Undo, invoker/grid
focus, responsive containment, and empty browser diagnostics at desktop and
compact widths:

```bash
bun run playground:visual:spreadsheet-appearance-sort
```

The companion local-only A3S Test ACL records interactive accessibility and
diagnostic evidence for the same complete workflow.

## Focused Spreadsheet sort-range warning contract

The functional regression selects only the quarterly workbook's formula
column, verifies both exact and expanded Traditional Office-style warning choices, applies the
expanded Custom Sort, and confirms formula rebasing, one-step Undo, invoker
focus, responsive containment, and empty browser diagnostics at desktop and
compact widths:

```bash
bun run playground:visual:spreadsheet-sort-range
```

The companion A3S Test ACL adds accessibility-tree evidence and executes the
same local release contract without adding A3S Test to GitHub Actions or Pages.

## Focused Spreadsheet owned-range sort contract

The functional regression starts from one cell inside a native table or active
worksheet AutoFilter, requires the exact complete owner, locks the structural
header and top-to-bottom direction, and checks responsive dialog containment.
The table path verifies relative-formula translation and one-step Undo; the
AutoFilter path applies a real native value selection and proves hidden-row
ownership follows the sorted record and its Undo at desktop and compact widths:

```bash
bun run playground:visual:spreadsheet-owned-range-sort
```

The companion local-only A3S Test ACL records the table-owned warning, locked
controls, focus restoration, accessibility, and browser diagnostics without
adding A3S Test to GitHub Actions or Pages.
