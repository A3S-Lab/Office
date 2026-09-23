# WPS Editor Parity Plan (First Principles)

Last reviewed: 2026-09-22

This plan is the execution contract for bringing A3S Office editors to
**product-enough parity with WPS daily workflows**, not a mandate to clone
every WPS ribbon control, pixel, or enterprise service.

It complements [`ROADMAP.md`](../../../ROADMAP.md) (capability inventory and
ordered R0→R6 work) and
[`editor-quality-roadmap.md`](editor-quality-roadmap.md) (quality gates).
When those documents disagree with this plan on priority, **this plan wins for
WPS-facing daily workflows**; ROADMAP still owns deep format / revision /
calculation slices.

## First-principles review

| Question | Answer |
| --- | --- |
| Core mission? | Embeddable, format-native editors that open real Office files, finish common daily work without silent data loss, and leave identity / storage / collab relay / AI to the host. |
| Does “match every WPS feature” serve that mission? | **No.** WPS ships desktop shells, cloud accounts, templates, macros, and long-tail UI that conflict with the browser embed boundary. Cloning them weakens the architecture. |
| Real problem? | Hosts and Chinese-locale users judge parity by **Writer / Spreadsheet / Presentation daily loops** plus trustworthy DOCX/XLSX/PPTX/PDF round trips—not by decorative art borders or every COM command ID. |
| Simpler alternative? | Keep WPS as a **local reference probe** (`wps-probe`, `wps-ui-probe`, `wps-fields-probe`). Prove A3S behavior with a3s-test Web/CDP ACL suites and Playwright baselines. |

### Explicit non-goals

- Pixel-perfect layout against WPS Writer pagination.
- Shipping WPS / Microsoft Office as a product runtime or CI dependency.
- Windows CUA Driver 0.10.0 profiles (locked **unsupported**; use A3S Test
  Web/CDP for browser evidence).
- Macros, ActiveX, VBA, or silent approximation of unsupported DrawingML /
  SmartArt / mail-merge graphs.
- Treating Markdown as a WPS clone (Markdown is an A3S differentiator).

### Success bar (product-enough WPS parity)

A surface is WPS-parity-enough when all of the following are true:

1. **Daily loop:** create → edit → format → save native bytes → reopen → same
   observable results (a3s-test ACL + empty console/page-error diagnostics).
2. **Shortcut / chrome contract:** primary WPS-familiar shortcuts and ribbon
   groups used in that loop keep focus and emit one undoable controlled update.
3. **Fail-closed fidelity:** unsupported structures diagnose; no silent
   clobber of identities (R0 corpus remains permanent).
4. **Evidence location:** suites listed in
   `scripts/office-editor-matrix.json`; artifacts under `.a3s-test/`.

## Surface map

| WPS product | A3S Office surface | Formats | Parity posture |
| --- | --- | --- | --- |
| WPS 文字 | Document (Writer) | DOCX, HTML, TXT | Primary fidelity track |
| WPS 表格 | Spreadsheet | XLSX, XLS, ODS, CSV | Calculation + daily grid workflows |
| WPS 演示 | Presentation | PPTX | Scene graph + slideshow daily workflows |
| WPS PDF | PDF | PDF | Annotate / organize first; native edit later |
| *(none)* | Markdown | MD | Maintain as differentiator; no WPS clone |

## Ordered delivery (do not reorder without exit evidence)

### Track A — Permanent WPS daily-loop gates (P0)

Already authored as deterministic ACL suites:

| Surface | Suite | Workflow |
| --- | --- | --- |
| Writer | `tests/e2e/word-wps-daily-loop.acl` | Type → bold / grow font / center → find-replace → insert 2×2 table → export DOCX → reopen |
| Spreadsheet | `tests/e2e/spreadsheet-wps-daily-loop.acl` | Enter values → SUM → bold → insert row → export XLSX → reopen |
| Presentation | `tests/e2e/presentation-wps-daily-loop.acl` | New slide → edit title → insert shape → export PPTX → reopen |

Run (Windows uses Web/CDP; CUA stays fail-closed):

```bash
bun run test:e2e:wps-daily:writer
bun run test:e2e:wps-daily:spreadsheet
bun run test:e2e:wps-daily:presentation
# or
bun run test:e2e:wps-daily
```

Operator wiring: suites must remain in `office-editor-matrix.json` so
`bun run office:ops -- check|gate <surface>` includes them.

**Next daily-loop expansions (only after the three gates stay green):**

1. Writer: comments inside the same save/reopen daily loop (**done** —
   `word-wps-daily-loop.acl` records `WPSDAILY review note` and asserts it after
   DOCX reopen).
2. Spreadsheet: AutoFilter toggle + one sort, then save/reopen (**done** —
   `spreadsheet-wps-daily-loop.acl` builds an isolated Score column, enables
   AutoFilter, confirms the owned-range descending sort dialog, and asserts
   descending order plus restored AutoFilter chrome after XLSX reopen).
3. Presentation: duplicate slide + notes text, then save/reopen (**done** —
   `presentation-wps-daily-loop.acl` edits speaker notes, duplicates with
   Ctrl+D, and asserts notes on both slides after PPTX reopen).
4. PDF: annotate or page-organize → save → reopen (**done** —
   `pdf-wps-daily-loop.acl` inserts a blank page, downloads the PDF, and asserts
   page count 5 after reopen; matrix-wired + `bun run test:e2e:wps-daily:pdf`).

### Track B — R0 close (P0, format trust)

Follow `ROADMAP.md` § Remaining roadmap → R0. Prefer **corpus growth** and
review/property admission over decorative PDF art-border treadmills. Exit when
`bun run test:corpus:r0` and representative Traditional Office / WPS-authored
fixtures reopen without unreported loss.

### Track C — R1 Writer daily leftovers (P0/P1)

After R0 exit evidence exists: structural/object compare, bounded DrawingML
expansion with fixtures, deeper fields, content-control forms (checkbox
`0.349.0`, drop-down list `0.351.0`, combo box `0.352.0`, date `0.353.0`,
repeating section `0.355.0`, passive custom-XML data binding `0.356.0`,
single-paragraph body-level block controls `0.357.0`, multi-paragraph
plain-run block bodies `0.368.0`, one-level nested leaf block controls
`0.369.0`), then mail merge
(`MERGEFIELD` `0.350.0`; bounded host `DocumentMailMergeSource` preview runners
`0.354.0`; bounded batch DOCX generation `0.370.0`; bounded recipient-filter UI
`0.371.0`; two-level nested block controls `0.372.0`). Depth ≥ 3 block nesting
remains fail-closed among R1 leftovers.

### Track D — R2 Spreadsheet (P0/P1)

Versioned formula corpus vs Traditional Office/Excel (`bun run test:corpus:r2`
through `0.311.0` bounded `UNIQUE` spill, `0.311.0` bounded `TRANSPOSE` spill,
`0.311.0` bounded `SEQUENCE` spill,
`0.311.0` locale-guessing `VALUE`,
`0.311.0` 1904 formula dates,
`0.311.0` bounded `NUMBERVALUE`, plus `0.375.0`
`SUMPRODUCT`/`PRODUCT`/`VALUE` and earlier approximate
`VLOOKUP`/`MATCH`/`HLOOKUP`, `AVERAGEIF`/`AVERAGEIFS`, exact lookups, conditional
aggregates, date/`LEFT`/volatility gates), pivot style +
diagnostic / slicer fail-closed corpus (`0.360.0`–`0.361.0`), in-app
slicer-style multi-select report filters (`0.376.0`), CF precedence
corpus (`0.362.0` / `bun run test:corpus:r2-cf`), bounded worksheet-image
transform round trips (`0.364.0`), remaining native XLSX slicer packages /
timeline / pivot-chart depth, then virtual-grid ownership replacing Fortune
redraw for large books.

### Track E — R3 Presentation (P1)

Broader native animation preservation, safe media relationships, master/layout
authoring, shape/chart fidelity with fixtures.

### Track F — R4 PDF workbench (P1/P2)

Native text/image/link editing, OCR/conversion provider contracts, then
e-sign / true redaction / protection.

### Track G — Reference probes (local only)

Keep WPS COM probes as optional developer evidence:

- `bun run office:ops -- wps-probe --connector`
- `bun run office:ops -- wps-fields-probe --profile numeric|common --json`
- `bun run office:ops -- wps-ui-probe --profile shell|fields|all --json`

Never gate CI on an installed WPS binary.

## Test strategy (a3s-test + CUA)

| Layer | Tool | Role |
| --- | --- | --- |
| Deterministic E2E | a3s-test ACL (`tests/e2e/*.acl`) | Primary interaction contract |
| Exploratory / CUA-style browser | `a3s-test agent` / `office:ops -- a3s agent` | Observe → one typed action → observe |
| Native GUI CUA | `a3s cua certification` | Linux/mac where certified; **Windows unsupported** |
| Pixel baseline | Playwright visual-tests | Supplemental desktop/compact baselines |
| Format trust | `bun run test:corpus:r0` | Permanent no-clobber gate |

Rules:

1. Promote agent discoveries to the smallest ACL suite; do not leave critical
   paths only in chat transcripts.
2. Prefer role / label / testid targets over brittle CSS when possible.
3. Capture accessibility + console + page_errors on daily loops; non-empty
   diagnostics fail the gate.
4. Do not invent Windows CUA success claims while the locked driver profile is
   unsupported.

## Verification checklist for “parity complete”

Do **not** mark this plan complete until every item has current-tree evidence:

- [x] Track A suites authored and matrix-wired (`word` / `spreadsheet` /
      `presentation` / `pdf` WPS daily-loop ACLs + `bun run test:e2e:wps-daily*`).
- [x] `a3s-test check` admits all four daily-loop suites (syntax/contract).
- [x] Track A: four WPS daily-loop ACL suites green via
      `office:ops -- a3s run … --cdp-port` on Windows (Writer / Spreadsheet /
      Presentation / PDF create→edit→save→reopen; empty console and page-errors).
- [x] Matrix + `office:ops check` list those suites for writer / spreadsheet /
      presentation / pdf.
- [x] R0 no-clobber corpus gate green (`bun run test:corpus:r0`, 131/131 on
      2026-09-22 including complete reviewable tblPrChange/sectPrChange property
      subsets through tblpPr and pgBorders, bounded rPrChange/pPrChange
      formatting revisions, plus paragraph-mark insertion, paragraph-break
      split, paired text-move, and companion move-range bookmark revision
      identity). Writer PDF tagged bootstrap through `0.333.0` (URI links,
      `/Link`+`/OBJR`, `/Figure`+MCID, `/Table`→`/TR`→`/TH`|`/TD`); remaining
      art borders / full PDF/UA explicitly deferred past R0.
- [x] R0 exit evidence per ROADMAP exit criteria (corpus gate + tagged PDF
      bootstrap + fail-closed active content); continuous corpus growth stays
      a permanent gate, not an open R0 blocker.
- [ ] R1–R4 ordered leftovers either gated green or explicitly deferred with
      fail-closed diagnostics (no silent gaps in claimed workflows). Same-shape
      table cell Compare lands through `0.334.0`; simple-text paragraph-break
      split/merge Compare lands through `0.335.0`; soft-break Compare lands
      through `0.336.0`; DrawingML parallelogram / hexagon text-box presets land
      through `0.337.0` / `0.338.0`; bounded document-property fields
      (`FILENAME` through `PRINTDATE`) land through `0.339.0`–`0.348.0`;
      checkbox content controls land in `0.349.0`; bounded `MERGEFIELD` lands
      in `0.350.0`; drop-down list content controls land in `0.351.0`; combo box
      content controls land in `0.352.0`; date content controls land in
      `0.353.0`; bounded mail-merge host preview runners land in `0.354.0`;
      bounded repeating-section content controls land in `0.355.0`; bounded
      custom-XML content-control data bindings land in `0.356.0`;
      single-paragraph body-level block content controls land in `0.357.0`;
      R2 formula-compatibility corpus bootstrap lands in `0.358.0`
      (`bun run test:corpus:r2`); volatility/`SUBTOTAL` expansion lands in
      `0.359.0`; pivot style catalog + calculated-field diagnostic split lands
      in `0.360.0`; daily text `LEFT`/`RIGHT`/`LEN`/`MID` lands in `0.361.0`;
      CF precedence corpus lands in `0.362.0` (`bun run test:corpus:r2-cf`);
      daily date `DATE`/`YEAR`/`MONTH`/`DAY` lands in `0.363.0`; bounded
      worksheet-image 90-degree / flip transforms land in `0.364.0`; daily
      `SUMIF`/`COUNTIF`/exact `VLOOKUP` land in `0.365.0`; multi-criteria
      `SUMIFS`/`COUNTIFS` land in `0.366.0`; `INDEX`/`MATCH`/exact `HLOOKUP`
      land in `0.367.0`; multi-paragraph plain-run block bodies land in
      `0.368.0`; one-level nested leaf block controls land in `0.369.0`;
      bounded mail-merge batch DOCX generation lands in `0.370.0`; bounded
      recipient-filter UI lands in `0.371.0`; two-level nested block controls
      land in `0.372.0`; daily `AVERAGEIF`/`AVERAGEIFS` land in `0.373.0`;
      approximate ascending `VLOOKUP`/`MATCH`/`HLOOKUP` land in `0.374.0`;
      `SUMPRODUCT`/`PRODUCT`/`VALUE` land in `0.375.0`; in-app slicer-style
      multi-select report filters land in `0.376.0`; bounded `NUMBERVALUE` lands in `0.311.0`; 1904 `DATE`/`YEAR`/`MONTH`/`DAY` land in `0.311.0`; locale-guessing `VALUE` lands in `0.311.0`; bounded `SEQUENCE` spill lands in `0.311.0`; bounded `TRANSPOSE` spill lands in `0.311.0`; bounded `UNIQUE` spill lands in `0.311.0`; depth ≥ 3 block nesting, native XLSX
      slicer/timeline/pivot-chart depth, Fortune virtual-grid ownership, and
      remaining R2–R4 remain.
- [x] Docs (this file, ROADMAP executive view, e2e README) describe the same
      bar in English.

Until then, keep delivering the next ordered slice and leave the long-running
goal active.

### Environment notes (2026-09-22)

- Locked Windows CUA remains unsupported; use A3S Test Web/CDP.
- Local release gate pins `a3s-test 1.0.1` (matches `crates/test` workspace).
- On Windows, `bun run test:e2e:wps-daily*` uses
  `office:ops -- a3s run … --cdp-port --command-timeout-ms 120000` so the
  compiled CDP adapter owns Chrome launch (`--no-sandbox`) and per-command
  waits match the web-gate deadline (avoids Fortune Sheet cold-start flakes
  under the 30s default). Do not point `A3S_TEST_AGENT_BROWSER` at native
  `agent-browser` when a CDP port is set; that bypasses the adapter and fails
  with `DevToolsActivePort`. Native binary override:
  `A3S_TEST_AGENT_BROWSER_NATIVE`.
- Linux/mac CI may still use `scripts/run-a3s-test-web-gate.sh` with `jq` on
  `PATH`.
