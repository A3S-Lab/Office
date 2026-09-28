# Accessibility Conformance Report (VPAT® 2.5)

**Product:** `@a3s-lab/office` (A3S Office editors)  
**Version:** package version under test in this repository  
**Report date:** 2026-09-26  
**Contact:** A3S Lab  

## Evaluation methods

- Automated: `@axe-core/playwright` WCAG 2.0/2.1 A/AA tags in
  `visual-tests/office-editors.a11y.spec.ts` (runs in `bun run playground:visual`).
  Analysis is scoped to each editor chrome root; canvas/page-raster surfaces
  that are not yet semantic are excluded with documented reasons in that suite.
- Manual / product design: keyboard paths for ribbons, dialogs, Presentation
  object list, Spreadsheet sheet bar, and live-region active-cell announcements.
- Locales exercised: `zh-CN` (default) and `en-US` (`?locale=en-US`,
  `bun run playground:visual:en-us`).

## Applicable standards

| Standard | Level |
| --- | --- |
| WCAG 2.1 | Level A and AA (editor chrome in scope) |
| Section 508 (revised) | Mapped via WCAG 2.1 A/AA for in-scope UI |
| EN 301 549 | Mapped via WCAG 2.1 A/AA for in-scope UI |

## Terms

| Term | Meaning |
| --- | --- |
| Supports | Functionality meets the criterion without known defects in scope. |
| Partially Supports | Some functionality meets the criterion; defects or exclusions are listed. |
| Does Not Support | Majority of functionality does not meet the criterion. |
| Not Applicable | Criterion is not relevant to this product. |

## Product overview

Five embeddable editors: Document (Writer), Spreadsheet, Presentation,
Markdown, and PDF. Hosts set `locale` / `messages`, keyboard chrome, dialogs,
and status regions. Spreadsheet cells and Presentation/PDF page contents are
drawn on canvas or raster surfaces; semantic alternatives exist where noted.

## WCAG 2.1 checklist (summary)

Criteria not listed are **Not Applicable** for this ACR (time-based media,
captions for prerecorded video, etc.).

| Criterion | Conformance | Remarks |
| --- | --- | --- |
| 1.1.1 Non-text Content | Partially Supports | Chrome icons have accessible names. PDF page rasters and EmbedPDF canvases are decorative/out of chrome scope until PDF/UA; excluded from the axe gate with documented selectors. |
| 1.3.1 Info and Relationships | Partially Supports | Ribbons, dialogs, and lists expose roles/names. Spreadsheet sheet strip uses a named group of sheet buttons (not ARIA tablist) so per-sheet option menus remain valid. Fortune grid cells lack grid semantics until the Phase 4 virtual grid. |
| 1.3.2 Meaningful Sequence | Supports | Document/Markdown reading order follows DOM; Presentation object list provides linear access beside the canvas. |
| 1.4.3 Contrast (Minimum) | Supports | Ribbon group labels, status/save text, Presentation object-type labels, and speaker-notes captions use `var(--a3s-muted)` (AA on panel backgrounds). Verified by axe on editor chrome. |
| 1.4.4 Resize Text | Supports | Chrome remains usable at common zoom levels; compact viewports covered by the 768 CSS-px project. |
| 1.4.10 Reflow | Partially Supports | Chrome reflows; large canvases may scroll. |
| 1.4.11 Non-text Contrast | Partially Supports | Controls meet AA in chrome; canvas glyphs are not evaluated. |
| 1.4.13 Content on Hover or Focus | Supports | Tooltips/popovers dismiss with Escape where implemented. |
| 2.1.1 Keyboard | Partially Supports | Ribbons, dialogs, sheet bar, Presentation object list, and PDF toolbar are keyboard operable. Spreadsheet cell caret remains canvas-bound until Phase 4. |
| 2.1.2 No Keyboard Trap | Supports | Dialogs restore focus; Escape closes menus. |
| 2.4.3 Focus Order | Supports | Focus moves in a consistent order through chrome. |
| 2.4.6 Headings and Labels | Supports | Commands and fields resolve through the message catalog (`zh-CN` / `en-US`). |
| 2.4.7 Focus Visible | Supports | Focus rings on chrome controls. |
| 2.5.3 Label in Name | Supports | Visible labels match accessible names for chrome controls. |
| 3.1.1 Language of Page | Partially Supports | Host sets document language; editors accept BCP 47 `locale`. Playground defaults to `zh-CN` and supports `en-US`. |
| 3.2.1 On Focus | Supports | Focus alone does not change context unexpectedly. |
| 3.2.2 On Input | Supports | Changing controls does not force unexpected navigation. |
| 3.3.1 Error Identification | Supports | Dialog validation surfaces named errors. |
| 3.3.2 Labels or Instructions | Supports | Fields expose labels/descriptions. |
| 4.1.1 Parsing | Supports | Generated DOM is well-formed for chrome. |
| 4.1.2 Name, Role, Value | Partially Supports | Chrome controls expose name/role/value. Canvas cells/slides/PDF pages rely on live regions or adjacent lists rather than full widget semantics. |
| 4.1.3 Status Messages | Supports | Spreadsheet active-cell and filter/format-painter status use a polite live region; save status is announced via labeled outputs. |

## Known gaps (tracked)

1. **Spreadsheet virtual grid (Phase 4)** — Opt-in A3S
   `SpreadsheetVirtualGrid` (`virtualGrid` / playground `?virtualGrid=1`) owns
   viewport paint, `role="grid"` / `aria-activedescendant`, and literal cell
   edits; Playwright axe covers the grid root. Default path remains Fortune
   (dual-mount command port) until merges/CF paint and command-port exit land.
2. **Presentation slide canvas** — objects are reached via
   `PresentationObjectList`, not canvas hit-testing for SR.
3. **PDF page rasters** — EmbedPDF images/canvas excluded from chrome axe until
   tagged PDF / PDF/UA work.
4. **npm Trusted Publisher (0.1)** — release process owner-blocked; unrelated to
   a11y runtime.

## Evidence commands

```bash
bun run playground:build
bun run playground:visual -- visual-tests/office-editors.a11y.spec.ts
bun run playground:visual:en-us
```

## Legal disclaimer

This ACR describes the product as evaluated on the stated date using the
methods above. It is not a warranty of compliance in every host embedding or
assistive-technology combination. Re-evaluate after material UI changes.
