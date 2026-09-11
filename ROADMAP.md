# A3S Office / Traditional Office Capability Gap Roadmap

Last reviewed: 2026-09-10

This roadmap compares the current `main` branch of A3S Office with the public
capability surface of Traditional Office. It is a prioritization tool, not a
claim that the products have the same architecture or that every Traditional
Office service belongs in this library.

Traditional Office availability varies by operating system, region,
application version, and subscription. The comparison therefore uses
capability families rather than
plan-specific limits. A3S Office is a browser-first, embeddable editor whose
host application intentionally owns identity, storage, permissions,
collaboration transport, and AI providers.

**Product-enough** is not “finish every Traditional Office row.” The bar is a
host-embeddable, format-native editor suite: open real Office files, complete
common daily workflows without silent data loss, keep unsupported paths
fail-closed and diagnosed, and leave identity / storage / collab relay / AI to
the host. Completing R0→R6 in full is optional long-horizon work; shipping
product-enough means closing the ordered remaining work below through R0 exit
plus the R1 / R2 / R3 / R4 main paths already called out.

## Status and priority

| Mark | Meaning |
| --- | --- |
| **Supported** | The editable path exists and is covered by deterministic tests or native round-trip evidence. |
| **Partial** | A useful subset exists, but Traditional Office has broader editing, rendering, or format coverage. |
| **Gap** | There is no product-grade editable path yet. |
| **Host-owned** | A typed integration boundary exists or is intended; A3S Office should not ship the corresponding backend/service. |
| **Not planned** | The capability conflicts with the browser security model or product boundary. |

| Priority | Rule |
| --- | --- |
| **P0** | Prevent data loss, security regressions, broken daily editing, or materially incorrect native-file output. |
| **P1** | Close common Traditional Office workflows that materially improve adoption. |
| **P2** | Add advanced professional workflows after the underlying model is stable. |
| **P3** | Long-tail compatibility, ecosystem, or convenience work. |

## Executive view

| Surface | A3S Office today | Largest Traditional Office gaps | Direction |
| --- | --- | --- | --- |
| Shared shell | Five lazy browser surfaces, responsive ribbon and dialogs, keyboard/focus contracts, typed host ports | Bundled cloud workspace, live co-editing service, native desktop/mobile shells, large template marketplace | Keep services host-owned; deepen reusable integration protocols |
| Writer | Strongest surface: structured editing, sections, tables, images, bounded editable text boxes and typed straight/elbow/curved connectors with line/arrow styles, equations, comments, revisions, deterministic bounded compare/combine, fields, typed TOC and native index authoring, notes, page chrome, source-backed DOCX, and WPS-imported page/font/grid parity evidence | Complete routed/custom DrawingML objects, mail merge, structural/object compare, tables of figures/authorities, exact desktop pagination | Remains the first fidelity track |
| Spreadsheet | Workbook editing, formulas, recalculation, native Tables/ListObjects, bounded structured-reference formulas, automatic calculated-column fill, complete common totals-row authoring, charts, pivots, conditional formatting, validation, protection, comments, print setup | Advanced pivots/slicers, external data, macros/add-ins, advanced analysis, and broader formula parity | Stabilize calculation and native workbook semantics before adding more UI |
| Presentation | Editable scene graph, masters/layouts, text, shapes, images, tables, charts, groups, comments, transitions, composable bounded Work entrance/exit animations with native PPTX timing-tree round trips, slideshow/presenter view | Broader native animation preservation, emphasis/motion paths, media, broad shape/SmartArt fidelity, full master authoring, richer views and video export | Build expressiveness on the typed scene graph |
| PDF | PDFium rendering, search, forms, annotations, navigation, bounded long-file rendering, save, and Worker-backed insert/delete/rotate/reorder/extract/merge/split page organization | Native text/object editing, document-level catalog rewriting, conversion, OCR, signatures, redaction, compression | Evolve from viewer/annotator into an optional PDF workbench |
| Markdown | GFM source, visual editing, split preview, tables, task lists, links, images, code | No direct Traditional Office equivalent | Maintain as an A3S differentiator |
| Automation | Rust CLI, MCP server, Office Skill, bounded typed mutations | Office-style macro/add-in ecosystem | Prefer deterministic, auditable automation; do not execute Office macros |

## Remaining roadmap (product-enough bar)

Tip reviewed against `@a3s-lab/office@0.154.0`. Use this section as the active
backlog; the comparison tables below remain the gap inventory.

### Cross-surface permanent gates

These are already the default contract. Do not trade them away for breadth:

1. No silent clobber of critical identities or relationships; unsupported
   structures stay fail-closed with diagnostics.
2. One user intent produces one controlled update and one Undo record.
3. Shared ribbon / shortcut / focus contracts across the five surfaces, with
   a3s-test and Playwright evidence.
4. Typed host ports for Yjs, presence, and agents—without shipping cloud
   accounts, storage, or model providers inside this package.

### Surface enough criteria

| Surface | Enough when | Remaining priority (ordered) |
| --- | --- | --- |
| Document (Writer) | Report, contract, academic, and review workflows finish inside the embedded editor; DOCX round trips stay trustworthy | 1) Close remaining R0 revision / move / compare blockers; 2) R1 daily paths that still force leaving the editor; keep pagination predictable rather than pixel-cloning desktop Word |
| Spreadsheet | Common finance / operations / reporting workbooks calculate, filter, and export XLSX with bounded diagnostics | 1) Versioned formula-compatibility corpus (locale, arrays, volatility, deps); 2) pivot / slicer depth; 3) replace remaining Fortune redraw ownership with an A3S virtual grid |
| Presentation | Typical sales / training / classroom decks keep visual hierarchy and animation / media intent | 1) Broader native PPTX animation preservation; 2) safe audio / video relationships; 3) deeper master / layout authoring |
| PDF | Annotate, form-fill, and page-organize with verified save / reopen | 1) Native text / image / link editing; 2) OCR and conversion provider contracts; 3) e-sign, true redaction, protection |
| Markdown | Reliable GFM visual editing (A3S differentiator) | Maintain; do not expand via Office long-tail commands |

### Ordered remaining work

Do the next slice in this order. Skip a row only when a higher row’s exit
evidence already exists in-repo.

#### 1. R0 close (P0, active)

Attribute-free paragraph `CT_Empty` glyph admission inside revision /
paragraph-break / move bodies is a **completed** family through `0.129.0`
(soft breaks, tabs, hyphens, `w:cr`, `w:lastRenderedPageBreak`, page-number
and long/short date-field glyphs). Do not grow that treadmill without a
review-workflow fixture that still fails closed today.

Still required for R0 exit:

1. Drawing / picture admission for whole-paragraph mark and paragraph-break
   bodies is **complete** through `0.135.0` (picture-only bodies through
   `0.135.0`; untracked supported inline DrawingML picture siblings beside mark
   wrappers through `0.134.0`; supported inline DrawingML pictures inside
   wrappers through `0.133.0`; safe relationship-bound external hyperlinks
   through `0.132.0`; untracked text-only sibling runs through `0.130.0`) with
   fail-closed ineligible shapes.
2. Move-range companion family is **largely complete** through `0.140.0`
   (one-level nested-table companions with move ancestry ≤2 tables through
   `0.140.0`; simple SDT-wrapped companion `w:move*Range*` bookmarks through
   `0.139.0`; multi-cell table-spanning companion `w:move*Range*` bookmarks with
   untracked sibling-cell text through `0.138.0`; Compare / same-document
   text-only move-range generation through `0.137.0`; single-cell table
   companion `w:move*Range*` bookmarks for supported text-only moves through
   `0.136.0`; cross-section companions through `0.131.0`). Remaining
   fail-closed edges: deeper nesting, a nested table beside the move, and
   SDT+nested-table combinations.
3. Remaining complex numbering edge cases that still normalize or diagnose as
   opaque. Omitted `w:ilvl` defaults to OOXML level `0` when `w:numId` and a
   supported `w:numberingChange` are present through `0.143.0`; opaque sibling
   `ST_NumberFormat` values in multi-level `w:original` (while the current
   `w:ilvl` stays common nfc 0–4) are admitted through `0.141.0`; missing
   `numId`, current-level bullet and picture formats, and malformed originals
   stay fail-closed.
4. Broader relationship-free `w:tblPr` / `w:trPr` / `w:tcPr` / `w:sectPr`
   reviewable subsets only when each property has a bounded fixture and
   accept/reject contract. Relationship-free attribute-only `w:tblpPr` priors
   in `w:tblPrChange` (known anchors/specs, bounded twips, optional FromText)
   are admitted through `0.142.0`; invalid or relationship-bound `tblpPr`
   stays fail-closed (`w:printerSettings`, header/footer refs, and nested
   cell/row ins/del stay permanently opaque).
5. Note-adjacent empty glyphs family is **complete** through `0.147.0`
   (attribute-free empty `w:separator` and `w:continuationSeparator` through
   `0.147.0`; attribute-free empty `w:annotationRef` through `0.146.0`;
   attribute-free empty `w:endnoteRef` through `0.145.0`; attribute-free empty
   `w:footnoteRef` through `0.144.0`). Attributed `footnoteRef` /
   `endnoteRef` / `annotationRef` and `footnoteReference` /
   `endnoteReference` with id stay fail-closed until dedicated fixtures admit
   them.
6. Browser document PDF output toward searchable text, broader vector paint,
   and tagged / PDF/UA structure without a second layout model. PDF bookmarks
   include Writer outline-level paragraphs (`p[data-office-outline-level]`)
   alongside h1–h6 with level nesting through `0.148.0`; Writer underlines
   paint as vector paths (`single` / `double` / `thick`) through `0.149.0`;
   Writer text highlights paint as vector fill strips through `0.150.0`;
   Writer paragraph borders paint as vector strokes (top/left/bottom/right;
   single/thick/double/dashed/dotted) through `0.151.0`; Writer `between` and
   `bar` paragraph border edges paint as vector strokes through `0.152.0`;
   catalog `/MarkInfo` and vector-run `/Span` ActualText deepen the tagged
   bootstrap through `0.153.0`; a catalog-linked StructTreeRoot stub
   (Document + outline-derived H1–H6/P with empty `/K`) lands through
   `0.154.0`; ParentTree / MCID content links for vector-run Spans land
   through `0.155.0`; Writer `wave` / `doubleWave` paragraph borders paint as
   explicit sine polylines through `0.156.0`; Writer `threeDEmboss` /
   `threeDEngrave` / `inset` / `outset` paragraph borders paint as dual-tone
   highlight+shadow offsets through `0.157.0`; Writer `zigZag` /
   `zigZagStitch` art paragraph borders paint as chevron polylines and catalog
   `/ViewerPreferences << /DisplayDocTitle true >>` lands through `0.158.0`;
   Writer `sawtooth` / `sharksTeeth` art paragraph borders paint as triangular
   teeth through `0.159.0`; Writer `triangles` / `triangle1` / `triangle2` art
   paragraph borders paint as closed isosceles triangles through `0.160.0`;
   Writer `ovals` / `rings` art paragraph borders paint as ellipse motifs
   through `0.161.0`; Writer `marquee` / `marqueeToothed` art paragraph borders
   paint as rectangle motifs through `0.162.0`; Writer `moons` art paragraph
   borders paint as crescent motifs through `0.163.0`; Writer
   `basicBlackSquares` / `basicWhiteSquares` art paragraph borders paint as
   discrete square stamps through `0.164.0`; Writer `basicBlackDots` /
   `basicWhiteDots` art paragraph borders paint as discrete circular stamps
   through `0.165.0`; Writer `basicBlackDashes` / `basicWhiteDashes` art
   paragraph borders paint as discrete dash stamps through `0.166.0`; Writer
   `basicThinLines` art paragraph borders paint as parallel hairlines through
   `0.167.0`; Writer `basicWideInline` / `basicWideMidline` /
   `basicWideOutline` art paragraph borders paint as thick geometric rails
   through `0.168.0`; remaining decorative art
   borders and full PDF/UA certification remain open.
7. Continuous no-clobber corpus growth for representative Traditional Office
   fixtures (diagnose every intentional normalization).

R0 exit criteria (unchanged in substance): representative fixtures reopen
without unreported data loss; edited native structures retain identity;
malformed inputs fail closed; pagination and export have deterministic
structural and visual evidence.

#### 2. R1 — Writer daily-work parity (P0/P1, next after R0 exit)

Already complete and must stay gated: typed TOC, native index, bounded
compare/combine for same-layout paragraphs/headings, field-authoring slice,
WPS UI-reference receipts, Writer shortcut evidence, quarter-turn picture
transforms, isolated WPS connector subset.

Remaining, in order:

1. Structural / object compare and richer mixed-content paragraph-mark /
   paragraph-break follow-ups that still leave review workflows incomplete.
2. Typed DrawingML shape model expansion only after each branch has native
   fixtures (no silent WordArt / chart / SmartArt approximations).
3. Field instructions beyond the bounded authoring slice (nested fields,
   document properties, mail fields) with independent contracts.
4. Content-control depth beyond inline plain/rich text (block, form,
   data-binding) with fail-closed fixtures.
5. Mail merge only after fields and an explicit host data-source contract
   exist.

#### 3. R2 — Spreadsheet calculation and analysis (P0/P1)

1. Versioned formula-compatibility corpus vs Traditional Office/Excel
   (locale, errors, arrays, volatility, dates, dependency updates).
2. Pivot depth: caches, grouping, calculated fields, slicers/timelines,
   styles, pivot charts.
3. Advanced sort/filter and conditional-format precedence gaps that still
   diverge from common Excel workflows.
4. Drawing/image fidelity before specialist analysis tools.
5. Continue replacing Fortune-owned redraw with A3S virtual-grid ownership
   for large workbooks.

Totals-row authoring, automatic calculated-column fill, and bounded
structured-reference calculation stay permanent gates.

#### 4. R3 — Presentation expressiveness (P1)

1. Broader native PPTX animation preservation (emphasis, motion paths,
   trigger-on-object) with round-trip-safe unsupported states.
2. Safe audio/video relationships and deterministic playback policy.
3. Visual master/layout editing and dedicated sorter/master workflows.
4. Shapes, connectors, effects, SmartArt fallback, and chart fidelity only
   with fixtures.
5. Slideshow evidence for timings, media, presenter tools, and printing.

Composable bounded entrance/exit effects with PPTX timing-tree round trips
remain a completed slice.

#### 5. R4 — PDF workbench beyond page organization (P1/P2)

Page organization is the first completed workbench slice and stays gated.

1. Native text / image / link editing with strict font and content-stream
   fallbacks.
2. Provider contracts for OCR and conversion (confidence + geometry in typed
   results).
3. E-sign, watermark, true redaction, protection, and optimization with
   byte-level safety tests.

#### 6. R5 — Host collaboration and AI references (P1)

Transport-neutral Yjs/Yrs replicas, host-channel sync, Awareness, and the
shared participant roster are in place. Remaining:

1. Map comments, revisions, selections, cursors, permissions, and conflicts
   onto the controlled editor model for host relays.
2. Provider-neutral examples for writing, formula, slide, and PDF AI tasks
   using immutable context and bounded mutations.
3. Keep production relay topology, sharing, and ACL enforcement host-owned.

#### 7. R6 — Enterprise and long-tail (P2/P3, after enough bar)

Defer until the product-enough bar above is met:

- Encrypted-file providers, signature validation display, accessibility
  metadata depth, localization/RTL depth, policy-driven passive macro
  handling.
- Legacy format conversion and specialist features only with authoritative
  fixtures.
- VBA / ActiveX / untrusted add-in **execution** stays outside the browser
  editor permanently.

### Explicitly out of product-enough scope

Do not schedule these as substitutes for R0–R4 exit work:

- Macro / ActiveX execution inside the browser editor
- Suite-grade proofing dictionaries, grammar engines, or bundled translation
- Template / asset marketplace (host-extensible catalogs only)
- Full DrawingML / SmartArt / chart authoring parity without per-branch
  fixtures
- Mail merge without a host data-source contract
- Pixel-perfect desktop Word pagination as a gate (predictable shared kernel
  pagination is the bar)

## Shared product and platform comparison

| Capability | Traditional Office public surface | A3S Office status | Priority / decision |
| --- | --- | --- | --- |
| Writer, Spreadsheet, Presentation, and PDF in one product | Integrated suite | **Supported** as lazy browser editors; Markdown is additional | Maintain |
| Microsoft Office file families | DOC/DOCX/DOCM, XLS/XLSX/XLSM, PPT/PPTX/PPTM and related formats vary by platform | **Partial**: DOCX; XLS/XLSX/ODS/CSV import; XLSX export; PPTX; PDF; macro-enabled semantics are not executable | P0 round-trip safety; P2 broader legacy conversion |
| Browser embedding | Traditional Office offers web and native applications | **Supported**, with React, Vue, Web Components, and Core entry points | A3S differentiator |
| Native Windows/macOS/Linux/iOS/Android applications | Product-owned native applications | **Host-owned**; A3S ships a browser component and native automation engine, not app shells | No core editor work |
| Cloud storage and multi-device sync | Integrated cloud documents and sync | **Host-owned** through controlled values, file actions, and repository ports | Publish reference adapters, not a storage service |
| Real-time co-editing, presence, reactions, sharing, permissions | Integrated collaboration workspace | **Partial**: browser/native Yjs replicas, a bounded host-channel sync adapter, typed browser and native Awareness presence, remote caret/selection overlays, a shared five-editor participant roster, and resumable CLI/MCP coding-agent event loops are supported; production relay topology, reactions, sharing, and ACL enforcement remain incomplete or host-owned | P1 richer native mutations and host integration adapters |
| Version history and recovery | Integrated cloud workflow | **Host-owned**; controlled revisions and repository contracts exist | Keep persistence outside the package |
| AI writing, document analysis, slide creation, formulas, and PDF chat | Integrated Traditional Office AI features on supported plans | **Host-owned** typed agent ports and immutable selection context; no bundled model/provider | Maintain provider-neutral boundary |
| Templates and asset marketplace | Large Traditional Office template catalog | **Partial**: small built-in starter templates, no marketplace | P3; host-extensible template catalog |
| Responsive and keyboard-accessible editing | Varies by Traditional Office platform | **Supported** as an explicit A3S contract across desktop and compact Web | P0 regression gate |
| IME composition in controlled rich-text editors | Native IME composition across desktop and web surfaces | **Supported** for Document, visual Markdown, and Presentation text: pre-edit text remains local, mid-composition host replacements are deferred, and the committed value publishes once; component and WebKit gates cover the contract | Maintain / P0 regression gate |
| Encryption, permissions, and document protection | File encryption plus view/edit permissions | **Partial**: XLSX protection semantics and package-security validation; permissions are host-owned; broad Office encryption is absent | P0 fail safely; P2 encrypted-file provider |
| Digital signatures | Signing workflows vary by document type and platform | **Gap**; package signatures are invalidated by editing and deliberately omitted | P2 validation/display; signing must use an explicit trusted provider |
| VBA, ActiveX, and Office add-ins | Desktop ecosystem capability | **Not planned for execution**; active content is never run and macro-free export omits it | P0 security boundary; consider passive quarantine/export policy only |
| Deterministic agent automation | Traditional Office exposes product automation through its own ecosystem | **Supported** through the Rust CLI, MCP, and Office Skill | Expand only with bounded, testable mutations |

## Writer comparison

| Capability | A3S Office status | Main gap or boundary | Priority |
| --- | --- | --- | --- |
| Text entry, character and paragraph formatting, styles, clipboard, format painter, undo/redo | **Supported**, including independent native `w:rFonts` ASCII, high ANSI, East Asian, and complex-script slots with exact direct/theme identity, mixed-run segmentation, one script-aware `Cmd/Ctrl+D` dialog, and one-step Undo; complete bounded Office 2010 OpenType ligature, numeral-form, numeral-spacing, style-set, and contextual-alternate controls shared with equations; independent native `w:lang` Latin, East Asian, and bidi proofing slots plus explicit `w:noProof` inclusion/exclusion; mutually exclusive native `w:caps` / `w:smallCaps`; signed native `w:spacing` and `w:position` with explicit zero; all five native `w:em` emphasis values with explicit `none` and direct-format clearing; native `w:vanish` hidden text with explicit visible resets, the standard `Cmd/Ctrl+Shift+H` shortcut, and an editing-only dotted reveal view; independent native `w:outline`, `w:shadow`, `w:emboss`, and `w:imprint` effects with explicit false resets and conflict-safe authoring; native `w:bdr` character borders with 25 visible line styles plus `nil` and `none`, direct/theme colors, bounded width/spacing, shadow, and frame semantics; all 18 native `w:u` values with direct or theme color identity and explicit resets; independent native `w:strike` / `w:dstrike` state with explicit resets; standard shortcuts where Traditional Office defines them; page chrome; formatting revisions; and exact DOCX reopen | Long-tail Word style and typography effects outside the declared typed models can still normalize | Maintain / P0 fidelity |
| Bullets, numbering, nesting, restarts, RTL lists | **Supported** for common editable structures, including bounded ordered-list style/start revisions, atomic review, and common single-level plus bounded multi-level native `w:numberingChange` round trips (omitted `w:ilvl` as OOXML default level `0` when `w:numId` + supported `w:numberingChange` through `0.143.0`; opaque sibling `ST_NumberFormat` values in `w:original` when current `w:ilvl` stays common nfc 0–4 through `0.141.0`) | Exotic numbering pictures, current-level bullet/picture numbering revisions, and some inherited list metadata remain compatibility work | P1 |
| Page size, orientation, columns, section/page breaks | **Supported** per section | Browser line breaking and pagination are not yet desktop-engine exact | P0 |
| Page margins (`w:pgMar`, `mirrorMargins`, `gutterAtTop`, `rtlGutter`) | **Supported** on current `main`: all seven native twip values, signed top/bottom overlap semantics, strict universal measures, header/footer distances, top/left/right gutters, facing-page swaps, multi-section inheritance, diagnostics, editing, PDF preview, and exact DOCX export | The continuous editing surface projects the first-page horizontal origin; physical-page PDF preview is authoritative for facing-page placement | Maintain / P0 regression gate |
| Headers, footers, first/even/default variants, page numbers | **Supported** | Complex fields and application-specific placement settings remain partial | P0 |
| Page borders (`w:pgBorders`) | **Supported** on current `main`: four ordered edges, 197 styles, theme/direct colors, page/text offsets, first/not-first display, front/back order, strict namespaces, diagnostics, exact DOCX export | Art borders and document-wide compatibility modifiers use bounded browser approximations | Maintain / P0 regression gate |
| Document grid and script-aware typography | **Supported** for section grid, run snap overrides, bundled/host/imported font choices, four native script-font slots, theme/style inheritance, mixed-script spans, and Traditional Office layout fixtures | Font substitution, missing glyphs, and browser shaping can still alter pagination | P0 |
| Tables, merges, sizing, margins, styles, row pagination, nested tables | **Partial**, with broad editable geometry, style inheritance, Word-style shared-edge border paint for rectangular and colspan/rowspan occupancy grids, opaque body `w:tblpPr` float round-trip, and reviewable relationship-free attribute-only `w:tblpPr` in `w:tblPrChange` through `0.142.0` | Full ECMA style matrix, editable floating placement UI, formulas, and advanced table tools are incomplete | P0/P1 |
| Inline and floating pictures, crop, wrap contour, layer, identity, alt text, quarter-turn transforms | **Partial**, with strong DrawingML picture support plus editable 90-degree rotation and horizontal/vertical reflection | Arbitrary-angle transforms, broad shapes, connectors, text boxes, WordArt, SmartArt, charts, and unsupported drawings normalize | P0 safe preservation; P1 editable drawings |
| Editable text boxes and bounded shape geometry | **Partial**, with isolated native WordprocessingML text boxes editable through inline/floating layout, five explicit presets (rectangle, rounded rectangle, ellipse, diamond, triangle), bounded geometry, fill, outline, padding, vertical alignment, WPS `mc:AlternateContent` import, and native DOCX round trips; isolated straight WPS VML connectors now have a typed endpoint/color/width/line-style/arrow model and native DrawingML export | Mixed paragraphs, routed or VML-only non-straight connectors, arbitrary shapes, WordArt, SmartArt, charts, malformed bodies, and unsupported DrawingML branches remain explicit compatibility boundaries | P1 editable drawings |
| OMML equations | **Partial**, with a large bounded structured model and strict/transitional import/export; relationship-free unsupported roots preserve as atomic native OMML | Relationship-bound, spoofed, misplaced, or over-limit OMML still flatten; editable coverage of remaining branches is P2 | P0 no-clobber; P2 coverage |
| Comments, replies, resolved state, anchors, modern IDs | **Supported** for editable review records and safe source preservation | Reactions, people sidecars, live presence, and server synchronization are absent/host-owned | P1 protocol |
| Track changes review | **Partial**: body-text insertions/deletions plus bounded whole-paragraph mark insertion/deletion (including multi-wrapper text-only bodies, soft breaks, tabs, carriage returns, last-rendered page breaks, page-number and long/short date-field glyphs, non-breaking and soft hyphens, relationship-free internal hyperlinks, relationship-free bookmarks, empty/`rPr`-only untracked sibling runs, and untracked text-only sibling runs), character-, paragraph-formatting, ordered-list numbering, **table-formatting** (`w:tblPrChange` with prior `w:jc` and/or `w:tblW` and/or `w:tblInd` and/or `w:tblCellMar` and/or `w:tblLayout` and/or `w:bidiVisual` and/or solid `w:shd`; live track-changes), **row-formatting** (`w:trPrChange` with prior `w:cantSplit`, `w:tblHeader`, and/or `w:trHeight` and/or `w:hidden` and/or `w:jc` and/or `w:gridBefore` and/or `w:gridAfter` and/or `w:wBefore` and/or `w:wAfter` and/or `w:cnfStyle` and/or `w:divId` and/or `w:tblCellSpacing`; live track-changes), **cell-formatting** (`w:tcPrChange` with prior `w:vAlign` and/or solid `w:shd` and/or `w:tcMar` and/or `w:tcW` and/or `w:noWrap` and/or `w:textDirection` and/or `w:tcFitText` and/or `w:hideMark` and/or `w:cnfStyle` and/or `w:hMerge` and/or `w:vMerge` and/or `w:gridSpan`; live track-changes), **section-formatting** (`w:sectPrChange` with prior orientation-only or complete `w:pgSz`, complete `w:pgMar`, `w:paperSrc`, equal-width `w:cols`, `w:titlePg`, `w:rtlGutter`, and/or bounded `w:docGrid`; live track-changes), text-only move revisions with soft breaks / internal hyperlinks / relationship-free bookmarks and companion `w:move*Range*` bookmarks (including cross-paragraph sandwich, cross-section destination sides, single-cell table enclosures, multi-cell tables whose sibling cells hold only untracked text-only content, simple SDT wrappers with `w:sdtPr` chrome, and one-level nested tables with move ancestry ≤2), and eligible text-only paragraph-break merge/split revisions; native `w:pPr/w:rPr/w:ins`/`w:del`, `w:rPrChange`, `w:pPrChange`, `w:numberingChange`, `w:tblPrChange` (alignment/preferred-width/indent/cell-margins/layout), `w:trPrChange` (cantSplit/tblHeader/trHeight/hidden/jc), `w:tcPrChange` (vAlign/solid-shd/tcMar/tcW/noWrap), `w:sectPrChange` (orientation/page-geometry/page-margins/paper-source/equal-width-cols), `w:moveFrom`, and `w:moveTo` round trips; broader relationship-free `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange` / `w:sectPrChange` opaque metadata round-trip; navigation, atomic accept/reject, immutable collaboration audit, and long-list virtualization; ineligible isolated paragraph-break shapes stay `docx.revisions.paragraph-break`; unpaired move-range markers stay `docx.revisions.move-range` | Attribute-free paragraph `CT_Empty` glyphs in revision bodies are admitted through `0.129.0`; untracked text-only sibling runs beside whole-paragraph marks are admitted through `0.130.0`; cross-section companion `w:move*Range*` bookmarks for supported text-only moves are admitted through `0.131.0`; safe relationship-bound external hyperlinks through `0.132.0`; supported inline DrawingML pictures through `0.133.0`; untracked picture siblings through `0.134.0`; picture-only bodies through `0.135.0`; single-cell table companion `w:move*Range*` bookmarks through `0.136.0`; Compare / same-document text-only move-range generation through `0.137.0`; multi-cell table-spanning companions with untracked sibling-cell text through `0.138.0`; simple SDT-wrapped companions through `0.139.0`; one-level nested-table companions (move ancestry ≤2) through `0.140.0`; opaque sibling `ST_NumberFormat` values in multi-level numbering `w:original` (current `w:ilvl` common nfc 0–4) through `0.141.0`; relationship-free attribute-only `w:tblpPr` priors in `w:tblPrChange` (known anchors/specs, bounded twips, optional FromText) through `0.142.0`; omitted `w:ilvl` as OOXML default level `0` when `w:numId` + supported `w:numberingChange` through `0.143.0`; attribute-free empty `w:footnoteRef` in whole-paragraph mark, paragraph-break, and text-move revision bodies through `0.144.0`; attribute-free empty `w:endnoteRef` in the same revision bodies through `0.145.0`; attribute-free empty `w:annotationRef` in the same revision bodies through `0.146.0`; attribute-free empty `w:separator` and `w:continuationSeparator` in the same revision bodies through `0.147.0`. Still open: deeper nested-table / beside-move nested-table / SDT+nested move-range edges; rich/relationship-bound moves; attributed `footnoteRef` / `endnoteRef` / `annotationRef` / `footnoteReference` / `endnoteReference` with id; remaining complex numbering edge cases (missing `numId`, current-level bullet/picture, and malformed originals stay fail-closed); remaining table/row/cell/section property revisions not yet reviewable | P0 |
| Bookmarks, links, captions, cross-references, citations, bibliography, footnotes/endnotes | **Partial**, with native identity and editable common paths | Wider field instructions, tables of authorities/figures, citation styles, and reference dialogs remain incomplete | P1 |
| Table of contents and outline authoring | **Supported**: shared semantic-heading/native-outline model plus typed insert/customize/refresh, levels 1–9, hyperlinks, live page numbers, alignment, four leader styles, stable paragraph-identity targets, one-step Undo, and native DOCX `TOC` round trips | Custom style-to-level mappings, tables of figures, and deeper TOC style formatting remain open | Maintain / P1 fidelity |
| Native index authoring | **Supported**: primary/secondary `XE` entries, cross-references, bold/italic page intent, stable marker targets, merged page numbers, typed insert/customize/refresh, 1–4 columns, indented/run-in layouts, four leader styles, one-step Undo, and native DOCX `XE`/`INDEX` round trips | Entry ranges, custom index types, letter-heading formats, authorities, and locale-specific collation controls remain open | Maintain / P1 fidelity |
| Fields | **Partial**: PAGE, NUMPAGES, SECTION, SECTIONPAGES, DATE, TIME, NUMWORDS, NUMCHARS, safe REF paths, and bookmark-backed PAGEREF with WPS numeric switches and bounded diagnostics | Broader Word field grammar, nested fields, switches outside the deterministic common subset, mail fields, document properties, and tables of figures/authorities remain incomplete | P1/P2 |
| Mail merge | **Gap** | No data-source mapping, recipient filtering, preview, or batch generation | P2 |
| Compare/combine documents | **Partial**: same-layout paragraphs/headings compare into deterministic insertion, deletion, character-formatting, paragraph-formatting, and bounded text-move revisions within one section; exact text-only whole-paragraph changes round-trip with native paragraph-mark records; reviewed-copy combine requires an exact reject-all baseline match | Changed complex structures, isolated paragraph-break merge/split or mixed-content paragraph-mark revisions, cross-section/range move generation, layout changes, and multi-copy structural conflicts remain explicit fail-closed boundaries | P1 |
| Content controls and forms | **Partial**: inline plain-text and rich-text controls are editable with aliases, tags, bounded locks, multiline text, appearance/color, and direct-paragraph DOCX `w:sdt` round trips; static controls in note/comment preservation remain separately bounded | Data binding, repeating sections, date/dropdown/picture/form controls, block controls, nested or relationship-bound semantics, placeholders, and active behavior are intentionally limited | P1/P2 |
| Spelling, grammar, language, translation | **Partial**: native Latin/East Asian/bidi `w:lang` metadata, explicit `w:noProof` authoring and round trips, browser spellcheck, and host actions | No suite-grade proofing dictionaries, grammar engine, or bundled translation service | Host/provider-owned; P2 adapter |
| Find, navigation, physical thumbnails, long-document review | **Supported** with transferable Worker import, bounded auxiliary-pane virtualization, model-level text/table-row NodeView windows for eligible structurally plain large DOCX files, and model-boundary keyboard navigation | Rich-feature giant DOCX fallback paths still need the same bounded body rendering and broader performance fixtures | P0 |
| PDF output | **Partial**: live browser pages share editor pagination; Latin/Latin-1 vector text clears under measured runs and paints Helvetica at the same geometry for search/copy; optional host `registerWorkPdfCjkFont` TrueType face enables searchable CJK vector text (fail-soft without it); title/`lang`/heading outline bootstrap including Writer outline-level paragraphs (`p[data-office-outline-level]`) alongside h1–h6 with level nesting through `0.148.0`; Writer underlines paint as vector paths (`single` / `double` / `thick`) through `0.149.0`; Writer text highlights paint as vector fill strips through `0.150.0`; Writer paragraph borders paint as vector strokes (top/left/bottom/right; single/thick/double/dashed/dotted) through `0.151.0`; Writer `between` and `bar` paragraph border edges paint as vector strokes through `0.152.0`; catalog `/MarkInfo` and vector-run `/Span` ActualText deepen the tagged bootstrap through `0.153.0`; catalog-linked StructTreeRoot stub (Document + outline-derived H1–H6/P with empty `/K`) through `0.154.0`; ParentTree / MCID content links for vector-run Spans through `0.155.0`; Writer `wave` / `doubleWave` paragraph borders paint as explicit sine polylines through `0.156.0`; Writer `threeDEmboss` / `threeDEngrave` / `inset` / `outset` paragraph borders paint as dual-tone highlight+shadow offsets through `0.157.0`; Writer `zigZag` / `zigZagStitch` art borders paint as chevron polylines and catalog `/DisplayDocTitle` through `0.158.0`; Writer `sawtooth` / `sharksTeeth` art borders paint as triangular teeth through `0.159.0`; Writer `triangles` / `triangle1` / `triangle2` art borders paint as closed isosceles triangles through `0.160.0`; Writer `ovals` / `rings` art borders paint as ellipse motifs through `0.161.0`; Writer `marquee` / `marqueeToothed` art borders paint as rectangle motifs through `0.162.0` | Full PDF/UA certification, bundled CJK fonts, remaining decorative art borders, and broader vector object fidelity remain gaps | P0/P1 |
| DOCX no-clobber round trip | **Partial but strong**: safe source-only parts and stable identities are preserved under bounded rules | Unsupported, relationship-bound, ambiguous, active, or changed structures may normalize and must stay diagnosed | P0 continuous work |

## Spreadsheet comparison

| Capability | A3S Office status | Main gap or boundary | Priority |
| --- | --- | --- | --- |
| Cell editing, multiple sheets, selection, keyboard navigation, search, clipboard, format painter, four-direction fill, copy from above, undo/redo | **Supported**, including Traditional Office `Cmd/Ctrl+D` and `Cmd/Ctrl+R` fill habits; exact `Ctrl+'` formula copying without relative-reference translation; cached-value-only `Ctrl+Shift+'`; target-style preservation; native relative-formula and style fill propagation; sparse-row-safe execution; maximum-dimension XLSX sheets with sparse virtual editing; viewport-bounded Canvas painting; streamed SheetJS Worker rows; an allocation-bounded transferable plain-OOXML cursor; and clone-free adoption of authenticated frozen matrices for eligible reserved-ID imports | Rich or structurally stateful imports still use the complete Fortune remount path; replace the remaining Fortune-owned redraw boundary with an A3S-owned virtual grid and continue large-workbook performance matrices | P0 |
| Formula entry and dependency-aware recalculation | **Partial** | Traditional Office/Excel function breadth, array/dynamic formulas, volatile/external semantics, and calculation parity are broader | P0 |
| Calculation settings and cached results | **Supported** for explicit modes and bounded calculation metadata | Complex data tables and compatibility calculation paths need broader parity | P0 |
| Cell formatting, number formats, alignment, wrapping, borders, merge/center | **Partial**, with explicit General, Number, CNY Currency, Accounting, Percentage, Short Date, Time, Scientific, Fraction, and Text presets; seven Traditional Office/Excel number-format shortcuts; static local-date `Ctrl+;` and minute-precision local-time `Ctrl+Shift+;` entry at the active cell through one native value/format batch using the workbook's 1900 or 1904 epoch; Traditional Office Grow/Shrink Font aliases with mixed-cell stepping; `Ctrl+2`/`Ctrl+3`/`Ctrl+4` Bold/Italic/Underline aliases; `Cmd/Ctrl+Shift+F` and `Cmd/Ctrl+Shift+P` deep links into the shared Font tab with exact family/size focus and a luminance-safe native-style preview; Automatic Color and No Fill direct-style resets; Outside/Clear Borders shortcuts; independent diagonal-down, diagonal-up, and crossed borders with exact OOXML flag round trips; exact single, double, single-accounting, and double-accounting underline controls; six Traditional Office text-orientation choices with accurate Fortune `rt` / stacked `tr='3'` and OOXML 0–180 / 255 round trips; 17 grouped Traditional Office built-in cell styles; direct XLSX font, fill, alignment, wrap, rotation, border, and number-format XF round trips; all 17 non-solid OOXML pattern fills and native linear/path gradient fills with Format Cells authoring plus viewport-only Canvas rendering; native shared/inline XLSX rich-text runs with exact whitespace, selected-run font formatting, direct insertion/deletion, and bounded authenticated formatted-HTML paste; exact source-level 1900/1904 date-system and typed numeric-serial retention; and conditional preservation of theme, indexed, automatic, and tint identities for direct and rich-run font, solid-fill, native pattern-fill foreground/background, native gradient stops, and border colors | Custom conditional sections, broader locale/currency choices, the complete themed style gallery, advanced differential or gradient border cases, simultaneous disjoint rich-text edits, non-font inline objects, and semantic colors in differential or conflicting-palette cases can normalize | P0/P1 |
| Sort and AutoFilter | **Partial** with shared dense/sparse current-region discovery, Traditional Office-style exact-or-expand warnings for partial sort selections, filter toggling, hidden-row state, keyboard menus, common text/value/numeric/blank conditions, exactly two same-column conditions joined by AND or OR, native OOXML `*`/`?`/`~` wildcard expressions, negative prefix/suffix operators, tie-aware Top/Bottom item and percentage filters, locally evaluated above/below-average and complete native relative-date, month, and quarter dynamic families on typed 1900/1904-system date cells, multi-column row ownership, manual-hide preservation, native XLSX criteria round trips with imported dynamic-filter recomputation, and stable value/custom-list/effective-color/conditional-icon sorting across top-to-bottom rows or left-to-right complete columns; offline Simplified Chinese pinyin/stroke comparison with optional case sensitivity and lexical numeric text; seven read-only built-in month/weekday sequences; bounded host-stored/session sequences through a typed store; an accessible preference manager for atomic user-sequence creation, editing, deletion, and reordering; exact table/AutoFilter owner expansion with totals exclusion, locked structural headers and direction, apply-time owner reauthentication, typed-filter recomputation, and opaque native hidden-row remapping; top-to-bottom header retention; top/bottom or left/right appearance placement; up to 64 ordered keys; blanks last for value orders; axis-aware relative-formula translation; and one-step Undo | Broader locale collation and large aggregate/rank or sort Worker/WASM offload remain incomplete | P1 |
| Native spreadsheet tables | **Partial**: Insert Table and grid-scoped `Cmd/Ctrl+T`, semantic ID-keyed ListObjects, contextual Table Design, 60 built-in OOXML styles, sparse-safe Convert to Range, row/column reconciliation, supported filters, native XLSX table parts, browser/Yjs convergence, bounded structured-reference calculation, validated calculated-column fill, and complete common totals-row authoring with native functions, labels, custom formulas, and exact XLSX metadata | Slicers, external/query tables, and broader table integrations remain incomplete | P0/P1 |
| Pivot tables | **Partial**: rows, columns, report filters, common aggregations, and native XML paths | Calculated fields/items, grouping, slicers, timelines, pivot charts, style controls, and broader cache semantics are incomplete | P1 |
| Charts | **Partial**: common, combination, scatter, bubble, radar, axes, labels, trendlines, and error bars | 3D/specialized charts, full formatting, drawing interactions, and every Excel chart extension are incomplete | P1/P2 |
| Conditional formatting | **Partial**: common comparisons, scales, data bars, icon sets, editable rules, bounded local formula rules with relative/absolute and cross-sheet references, blank-cell scans, stop-if-true ordering, and compact maximum-sheet ranges | Formula-function breadth, large-range worker evaluation, and advanced visuals still need more coverage | P1 |
| Data validation | **Partial**: common list, numeric, date, text-length, and bounded local custom-formula rules; compact maximum-sheet ranges; local dependent dropdowns through bounded `=INDIRECT(...)` text/cell references; workbook-local named ranges and one-row/one-column sources; complete common input and error-alert settings; blank/dropdown behavior; custom titles/messages; Stop blocking plus explicit Warning/Information confirmation branches for direct and formula-bar edits; proposed-value substitution with relative-reference anchoring; and exact native XLSX round trips | Broader formula-function coverage, dependent formulas beyond the bounded `INDIRECT` grammar, paste/object batch validation, and wider rule-specific browser affordances remain incomplete | P1 |
| Named ranges, print areas/titles, page setup, breaks, headers/footers | **Supported** for common workbook print workflows | Advanced print scaling, repeating content, and device-specific output need deeper parity | P1 |
| Freeze panes, row/column structure, clear modes, sheet lifecycle | **Supported** for common Traditional Office workflows, including row/column hide and unhide menu actions, grid-scoped `Cmd/Ctrl+9`, `Cmd/Ctrl+0`, `Cmd/Ctrl+Shift+9`, and `Cmd/Ctrl+Shift+0`, plus pre-allocation bounds of 10,000 rows or 1,000 columns per visibility intent | Grouping/outlining and advanced view state remain incomplete | P1 |
| Comments | **Partial**: legacy comments are editable | Rich comment formatting, threaded conversations, mentions, and live collaboration are incomplete/host-owned | P1 |
| Images and drawings | **Partial**: worksheet images round-trip under bounded geometry | Crop/rotation/flip can normalize; shapes, connectors, SmartArt, and unsupported drawing frames are gaps | P1/P2 |
| Protection | **Partial**: workbook/worksheet protection plus compact locked and passwordless editable ranges | Encryption, enterprise permissions, and every protection option are incomplete | P1/P2 |
| External data, connections, queries, macros, add-ins | **Gap** | No Power Query-like pipeline, data connections, VBA execution, or add-in runtime | P2; macro execution not planned |
| Advanced analysis | **Gap** for solver, scenarios, goal seek, consolidation, and equivalent specialist tools | Requires explicit bounded models rather than opaque workbook mutation | P2 |
| XLS/XLSX/ODS/CSV interoperability and PDF output | **Partial** | XLSX is authoritative; ODS/legacy conversion and exact print fidelity can normalize | P0/P1 |

## Presentation comparison

| Capability | A3S Office status | Main gap or boundary | Priority |
| --- | --- | --- | --- |
| Slide creation, duplication, ordering, thumbnails, canvas editing | **Supported** | Dedicated sorter/outline workflows are less complete than Traditional Office | P1 |
| Typed text, shapes, images, tables, charts, links, alt text | **Partial** | Shape catalog, connectors, text effects, SmartArt, equations, and broad DrawingML effects are incomplete | P0 no-clobber; P1 coverage |
| Multi-selection, grouping, move/resize/rotate, guides | **Supported** for the typed scene graph | Deep nested groups and exotic transforms need continued native parity tests | P0/P1 |
| Masters, layouts, placeholders, inherited backgrounds/artwork | **Partial** import/export support | Full visual master/layout authoring UI and all placeholder inheritance rules are incomplete | P1 |
| Slide transitions and timings | **Partial**: fade, push, wipe, split, cut; click/automatic advance | Traditional Office exposes a much broader transition catalog and controls | P2 |
| Object animations and triggers | **Partial** | The Work model supports composable entrance/exit cues, eight effects, three trigger modes, bounded non-overlapping timing, collaboration, playback, and native PPTX round trips for the supported subset; broader PPTX preservation, emphasis, motion paths, trigger-on-object, and a full timeline remain gaps | P1 |
| Audio, video, recording, and background music | **Gap** | No media timeline, playback policy, poster frame, trim, or relationship-safe export | P1/P2 |
| Speaker notes, slideshow, current-slide start, presenter view, timer | **Supported** | Rehearsed timings, recording, ink/laser tools, and multi-display edge cases are partial | P1 |
| Comments and review | **Supported** for slide comments and responsive review UI | Threads, mentions, assignments, live presence, and cloud synchronization are host-owned/incomplete | P1 protocol |
| Charts | **Partial**, sharing the typed chart model with Spreadsheet | Full PowerPoint chart workbook/editing and advanced formatting are incomplete | P1 |
| Views: normal, sorter, notes, reading, master | **Partial** | Normal editing, thumbnails, notes, slideshow, and presenter flows exist; dedicated sorter and master editing remain gaps | P1 |
| Print layouts and PDF | **Partial**: slides, notes, and 2/3/6-slide handouts are modeled | Exact Traditional Office print options, headers/footers, and vector fidelity need more coverage | P1 |
| Video export | **Gap** | Requires deterministic animation/media rendering and a host/native encoder | P3/provider-owned |
| PPTX round trip | **Partial** | The bounded appear/disappear, fade-in/out, fly-in/out, and zoom-in/out subset round-trips through native entrance/exit timing trees; unsupported drawings, SmartArt, broader animations, media, and extensions may normalize | P0 continuous work |

## PDF comparison

| Capability | A3S Office status | Main gap or boundary | Priority |
| --- | --- | --- | --- |
| Render, zoom, navigation, thumbnails, long-file windowing | **Supported** with PDFium | Continue platform/provider and very-large-file testing | P0 |
| Text search and selection evidence | **Supported** for browser search and bounded native text-layer evidence | Reading order, complex scripts, tagged structure, and accessibility extraction need deeper parity | P1 |
| Annotations and appearance controls | **Supported** for common annotations with history and save | Broader annotation types, replies, stamps, measurement, and collaboration remain incomplete | P1 |
| Form filling | **Supported** for common interactive forms | Form creation, calculation scripts, signatures, XFA, and complex appearance regeneration are incomplete | P1/P2 |
| Save edited annotations/forms | **Supported** | Full incremental-update/signature preservation requires stronger native guarantees | P0 |
| Edit existing text, images, links, and objects | **Gap** | Needs font matching, content-stream editing, reflow policy, and safe fallback behavior | P1 |
| Insert, delete, rotate, reorder, extract, merge, and split pages | **Supported with boundaries** | Dedicated Worker and Blob-level Undo/Redo are implemented; signed/encrypted files fail closed, risky catalog structures constrain destructive mutations, and page-only exports intentionally omit document-level objects with diagnostics | Maintain / P1 catalog fidelity |
| Compress and optimize | **Gap** | Requires bounded image/font/object optimization with quality controls | P2 |
| Convert PDF to/from DOCX/XLSX/PPTX/images | **Gap** as a general product workflow | Document-to-PDF exists; reverse conversion needs authoritative layout/OCR providers | P2/provider-assisted |
| OCR scanned documents | **Gap** | Should be an explicit provider interface with language, confidence, geometry, and privacy contracts | P1 provider boundary |
| Electronic and digital signatures | **Gap** | Needs trusted identity/certificate providers and signed-byte-range rules | P1 e-sign; P2 digital signatures |
| Watermark, redact, protect, and sanitize | **Gap** beyond package safety | Redaction must remove underlying content; visual covering alone is unacceptable | P1/P2 |
| AI summarize, explain, and translate | **Host-owned** | Use typed page/text evidence and host model policy; do not bundle a provider | P1 reference integration |

## Ordered delivery roadmap

The sequence below is capability-driven. It intentionally has no date promise;
a phase exits only when its evidence is complete. For the active backlog and
product-enough cut line, prefer
[Remaining roadmap (product-enough bar)](#remaining-roadmap-product-enough-bar)
above this section.

### R0 — Native fidelity and no-clobber baseline (P0, active)

- Continue complex DOCX/OOXML boundary work: section geometry, compatibility
  settings, drawings, revisions, fields, tables, equations, themes, and strict
  namespaces.
- Treat `w:pgBorders` support as the first completed section-decoration slice;
  keep its malformed, theme, multi-section, preview, and exact-export tests as
  permanent gates.
- Treat `w:pgMar`, document-wide mirror/top-gutter settings, and per-section
  `w:rtlGutter` as a completed section-geometry slice; retain exact native
  twips, signed overlap, strict-measure, malformed-input, inheritance,
  physical-page preview, editing, diagnostics, and round-trip tests as gates.
- Treat exact `w:pgSz` width, height, orientation, and paper code plus
  printer-specific `w:paperSrc` tray codes as a completed page-geometry slice;
  retain strict/transitional exact-measure, malformed and namespace-spoofed
  input, inheritance, custom-size editing, bounded tiny-page projection,
  per-page PDF geometry, diagnostics, and exact round-trip tests as gates.
- Treat per-section live page metrics as a completed page-geometry slice;
  mixed custom-size and orientation transitions now paginate in the shared
  JS/Rust kernel, render as variable-size live sheets, and retain exact
  per-page geometry in thumbnails and PDF capture.
- Treat bounded character, paragraph-formatting, ordered-list numbering,
  whole-paragraph mark insertion/deletion, imported text-only move revisions,
  and same-section Compare move inference as
  completed revision slices:
  retain accept/reject inversion, exact old-property snapshots, atomic
  multi-paragraph decisions, strict/transitional `w:rPrChange` and
  `w:pPrChange`, common single-level and bounded multi-level `w:numberingChange`, browser/Yrs
  convergence, malformed-input diagnostics, and Playground A3S Test coverage
  as permanent gates. Move pairs additionally require strict/transitional
  `w:moveFrom`/`w:moveTo` identity matching, atomic editor decisions, native
  export/reopen, immutable collaboration audit, and a diagnosed rich/range/
  relationship-bound fallback. Compare inference now admits deterministic unique
  lexical matches across simple paragraphs or headings within one section;
  whole-paragraph records additionally require matching text-only body and
  `w:pPr/w:rPr` metadata while allowing distinct native IDs. Eligible paragraph-break merge/split revisions now admit soft
  breaks, tabs, carriage returns, last-rendered page breaks, page-number and long/short date-field glyphs, non-breaking and soft hyphens, relationship-free internal
  hyperlinks, relationship-free bookmarks, and empty/rPr-only runs. Untracked
  text-only sibling runs beside whole-paragraph marks are admitted through
  `0.130.0`. Cross-section companion `w:move*Range*` bookmarks for supported
  text-only moves are admitted through `0.131.0`. Safe relationship-bound
  external hyperlinks in whole-paragraph mark and paragraph-break bodies are
  admitted through `0.132.0`. Supported inline DrawingML pictures in
  whole-paragraph mark and paragraph-break bodies are admitted through
  `0.133.0`. Untracked supported inline DrawingML picture siblings beside
  whole-paragraph mark wrappers are admitted through `0.134.0`. Picture-only
  paragraph bodies for whole-paragraph mark wrappers and paragraph-break
  neighbors are admitted through `0.135.0`. Single-cell table companion
  `w:move*Range*` bookmarks for supported text-only moves are admitted through
  `0.136.0`. Compare / same-document text-only moves generate companion
  `w:move*Range*` bookmarks on export through `0.137.0`. Multi-cell
  table-spanning companion `w:move*Range*` bookmarks with untracked
  sibling-cell text are admitted through `0.138.0`. Simple SDT-wrapped
  companion `w:move*Range*` bookmarks (at most one `w:sdt` containing the move,
  with `w:sdtPr` chrome) are admitted through `0.139.0`. One-level nested-table
  companion `w:move*Range*` bookmarks (move ancestry ≤2 tables) are admitted
  through `0.140.0`. Opaque sibling `ST_NumberFormat` values in multi-level
  numbering `w:original` (while the current `w:ilvl` stays common nfc 0–4) are
  admitted through `0.141.0`. Omitted `w:ilvl` defaults to OOXML level `0`
  when `w:numId` and a supported `w:numberingChange` are present through
  `0.143.0`; missing `numId`, current-level bullet and picture formats, and
  malformed originals stay fail-closed. Attribute-free empty `w:footnoteRef`
  glyphs in whole-paragraph mark, paragraph-break, and text-move revision
  bodies are admitted through `0.144.0`; attribute-free empty `w:endnoteRef`
  glyphs in the same revision bodies are admitted through `0.145.0`;
  attribute-free empty `w:annotationRef` glyphs in the same revision bodies
  are admitted through `0.146.0`; attribute-free empty `w:separator` and
  `w:continuationSeparator` glyphs in the same revision bodies are admitted
  through `0.147.0`. The note-adjacent empty glyphs family is complete through
  `0.147.0`; attributed `footnoteRef` / `endnoteRef` / `annotationRef` and
  `footnoteReference` / `endnoteReference` with id stay fail-closed. Floating anchors, empty or malformed drawings, unresolved
  embeds, residual deeper / beside-move / SDT+nested move-range edges, and
  remaining complex numbering edge cases remain explicit follow-up work.
  Relationship-free `w:tblPrChange` with a prior `w:jc`
  and/or `w:tblW` and/or `w:tblInd` and/or `w:tblCellMar` and/or `w:tblLayout`   and/or `w:bidiVisual` and/or solid direct-color `w:shd` and/or `w:tblLook` and/or `w:tblOverlap` and/or relationship-free `w:tblStyle` and/or dxa `w:tblCellSpacing` and/or relationship-free `w:tblStyleColBandSize` and/or relationship-free `w:tblStyleRowBandSize` and/or direct-color `w:tblBorders` and/or relationship-free `w:tblCaption` and/or relationship-free `w:tblDescription` and/or relationship-free attribute-only `w:tblpPr`
  snapshot is reviewable as `table-formatting` with accept/reject, live
  track-changes for alignment/preferred-width/indent/default-cell-margin/layout-mode/bidiVisual/solid-fill/tblLook/tblOverlap/tblStyle/tblCellSpacing/tblStyleColBandSize/tblStyleRowBandSize/tblBorders/tblCaption/tblDescription/tblpPr
  edits, and native export.   Relationship-free `w:trPrChange` with a prior `w:cantSplit`,
  `w:tblHeader`, and/or `w:trHeight` and/or `w:hidden` and/or `w:jc` and/or
  `w:gridBefore` and/or `w:gridAfter` and/or `w:wBefore` and/or `w:wAfter` and/or
  `w:cnfStyle` and/or `w:divId` and/or `w:tblCellSpacing` snapshot is reviewable as `row-formatting` with accept/reject, live
  track-changes for
  cantSplit/repeat-header/row-height/hidden/alignment/gridBefore/gridAfter/width-before/width-after/cnfStyle/divId/tblCellSpacing edits, and native export.
  Relationship-free `w:tcPrChange`
  with a prior `w:vAlign` and/or solid direct-color `w:shd` and/or `w:tcMar`
  and/or `w:tcW` and/or `w:noWrap` and/or `w:textDirection` and/or `w:tcFitText` and/or `w:hideMark` and/or `w:cnfStyle` and/or `w:hMerge` and/or `w:vMerge` and/or `w:gridSpan` and/or direct-color `w:tcBorders` snapshot is reviewable as `cell-formatting`
  with accept/reject, live track-changes for
  vertical-align/solid-fill/cell-margin/preferred-width/noWrap/textDirection/tcFitText/hideMark/cnfStyle/hMerge/vMerge/gridSpan/tcBorders edits, and
  native export.   Relationship-free
  `w:sectPrChange` with a prior orientation-only or complete `w:pgSz`
  (width/height with optional orientation/code), complete seven-edge
  `w:pgMar`, and/or `w:paperSrc` snapshot is reviewable as `section-formatting`
  with accept/reject, live track-changes for orientation, page-geometry,
  page-margin, paper-source, differentFirstPage (`w:titlePg`), rtlGutter,
  bounded `w:docGrid` (`w:type`/`w:linePitch`/`w:charSpace`), bounded `w:lnNumType`
  (`w:countBy`/`w:start`/`w:distance`/`w:restart`), bounded `w:pgNumType`
  (`w:fmt`/`w:start`/`w:chapStyle`/`w:chapSep`), `w:formProt`, `w:vAlign`, `w:noEndnote`, `w:textDirection`, `w:bidi`, bounded `w:footnotePr`, bounded `w:endnotePr`, `w:type` (`breakAfter`), and bounded `w:pgBorders` edits, and native export.
  `w:printerSettings` stays permanently opaque (required `r:id` / `CT_Rel`);
  `headerReference`/`footerReference` stay relationship-bound.
  Equal-width `w:cols` (num/space/sep, no unequal `w:col` children) and
  unequal-width `w:cols` (`w:equalWidth` off with bounded `w:col` children) are
  also reviewable as `section-formatting` with live column edits. Broader
  relationship-free
  `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, and `w:sectPrChange`
  records still round-trip as opaque metadata (not yet reviewable).
  Multi-wrapper text-only whole-paragraph mark bodies that share the mark
  author and date are admitted, including soft breaks, tabs, carriage returns, last-rendered page breaks, page-number and long/short date-field glyphs, non-breaking and
  soft hyphens, relationship-free internal hyperlinks, relationship-free bookmarks,
  and empty/`w:rPr`-only
  untracked sibling runs. Companion same-section `w:move*Range*`
  bookmarks around supported text moves (including cross-paragraph sandwich)
  round-trip with the move; unpaired range markers report
  `docx.revisions.move-range`. Soft breaks, tabs, carriage returns, last-rendered page breaks, page-number and long/short date-field glyphs, non-breaking and soft hyphens, relationship-free internal
  hyperlinks, and relationship-free ordinary bookmarks inside move wrappers
  are also admitted.
- Treat native all-caps and small-caps as a completed character-effect slice:
  retain one mutually exclusive typed state, semantic source text, standard
  shortcuts, body/header/footer parity, formatting revision and Format Painter
  support, exact `w:caps` / `w:smallCaps` import-export-reopen, and
  browser-authoritative measurement for case-shaped paragraphs as permanent
  gates.
- Treat native script-specific fonts as a completed typography slice: retain
  independent `ascii`, `hAnsi`, `eastAsia`, and `cs` slots; exact direct and
  theme identities; style inheritance; bounded mixed-script segmentation;
  Latin, East Asian, and complex-text mixed states in the shared `Cmd/Ctrl+D`
  dialog; Follow style; all-text Home font intent; body/page-chrome/note parity;
  Format Painter; formatting revisions; one-step Undo; strict/transitional and
  malformed-input rejection; diagnostics; and a real
  import-edit-export-reopen DOCX cycle as permanent gates. Keep browser-resolved
  families separate from native source identity so substitution cannot rewrite
  untouched theme references.
- Treat Office 2010 OpenType run typography as a completed typography slice:
  retain all 16 exact `w14:ligatures` combinations, default/lining/old-style
  `w14:numForm`, default/proportional/tabular `w14:numSpacing`, canonical
  `w14:stylisticSets` IDs 1-20, and explicit `w14:cntxtAlts` enable/reset
  values. Keep independent inheritance and mixed-selection safety in the
  shared `Cmd/Ctrl+D` dialog, partial per-run edits, one-step Undo, Format
  Painter, formatting revisions and reject restoration, body/page-chrome/note
  parity, exact namespace and `mc:Ignorable` export, reopen, dedicated
  diagnostics, bounded CSS projection, browser-authoritative measurement for
  active shaping, and the shared equation/text semantic model as permanent
  gates.
- Treat native character spacing as a completed typography slice: retain signed
  `w:spacing` values from -31,680 through 31,680 twips, explicit zero, mixed
  selection safety, `Cmd/Ctrl+D`, body/header/footer parity, Format Painter,
  formatting revisions, one-step Undo, strict-namespace import, exact DOCX
  export/reopen, and Worker/WASM layout eligibility as permanent gates.
- Treat native character horizontal scaling as a completed typography slice:
  retain exact `w:w` integers from 1 through 600 percent, empty-element 100
  percent defaults, explicit 100 percent resets, independent mixed-selection
  safety in the shared `Cmd/Ctrl+D` dialog, body/header/footer parity, Format
  Painter, formatting revisions, one-step Undo with spacing and position,
  strict/transitional namespace and malformed-input rejection, exact DOCX
  export/reopen, and browser-authoritative measurement for non-default scale as
  permanent gates. Do not invent a dedicated shortcut.
- Treat native pair kerning as a completed typography slice: retain exact
  `w:kern` thresholds from 0 through 3,277 half-points, explicit zero for all
  font sizes, style inheritance, and disabled behavior when the property is
  absent throughout the hierarchy. Keep effective `w:sz >= w:kern` rendering,
  independent mixed-selection safety, direct-format clearing, all editable
  Word stories, Format Painter, formatting revisions, one-step Undo, strict
  namespace and malformed-input rejection, exact DOCX export/reopen,
  diagnostics, and Worker/WASM effective-state parity as permanent gates. Use
  the shared `Cmd/Ctrl+D` dialog rather than inventing a dedicated shortcut.
- Treat native East Asian emphasis marks as a completed typography slice:
  retain all five `w:em` values, explicit `none`, style inheritance, independent
  mixed-selection safety and direct-format clearing in the shared `Cmd/Ctrl+D`
  dialog, all editable Word stories, Format Painter, formatting revisions,
  one-step Undo, strict/transitional namespace and malformed-input rejection,
  diagnostics, exact DOCX export/reopen, canonical CSS projection, and
  browser-authoritative line measurement for visible out-of-line marks as
  permanent gates. Do not invent a dedicated shortcut.
- Treat native hidden text as a completed typography slice: retain inherited,
  hidden, and explicit-visible states; exact `w:vanish` style resolution;
  mixed-selection safety in the shared `Cmd/Ctrl+D` dialog; the standard
  `Cmd/Ctrl+Shift+H` toggle; all editable Word stories; Format Painter;
  formatting revisions; one-step Undo; strict/transitional and malformed-input
  handling; exact DOCX export/reopen; and an editing-only dotted reveal view as
  permanent gates. Preview and PDF output must always suppress hidden text,
  unchanged comment XML remains source-preserved, and affected paragraphs stay
  on browser-authoritative measurement.
- Treat native outline, shadow, emboss, and imprint as a completed typography
  slice: retain independent inherited, enabled, and explicit-false states;
  permit outline plus shadow; reject or atomically clear conflicts involving
  emboss or imprint; preserve every editable Word story, style inheritance,
  Format Painter, formatting revisions, one-step Undo, strict/transitional and
  malformed-input handling, diagnostics, exact DOCX export/reopen, and the
  shared `Cmd/Ctrl+D` dialog as permanent gates. Keep the CSS and PDF
  projections bounded and paint-only so eligible paragraphs remain on
  Worker/WASM layout. Do not invent a dedicated shortcut.
- Treat native character borders as a completed typography slice: retain all
  25 visible `w:bdr` line styles plus explicit `nil` and `none`; direct and
  theme colors with tint/shade; exact 2–96 eighth-point widths; 0–31 point
  spacing; shadow and frame flags; style inheritance; mixed-safe Follow style,
  no-border, and editable modes in the shared `Cmd/Ctrl+D` dialog; the Home
  toggle; every editable Word story; Format Painter; formatting revisions;
  one-step Undo; strict/transitional and malformed-input handling;
  diagnostics; and exact DOCX export/reopen as permanent gates. Visible
  borders must use browser-authoritative measurement, explicit resets may stay
  on Worker/WASM, and CSS/PDF paint remains a bounded approximation. Do not
  invent a dedicated shortcut.
- Treat native character baseline position as a completed typography slice:
  retain signed `w:position` values from -3,168 through 3,168 half-points,
  explicit zero, mixed-selection safety, the shared `Cmd/Ctrl+D` dialog,
  body/header/footer parity, Format Painter, formatting revisions, one-step
  Undo, exact strict measures, namespace and duplicate rejection, native
  DOCX export/reopen, subscript/superscript precedence, and the explicit
  browser-measurement fallback as permanent gates.
- Make unsupported semantics explicit in compatibility reports; never attach
  relationship-bound or namespace-spoofed data to a regenerated identity.
- Move browser document PDF output toward searchable text, vector content, and
  tagged/accessibility structure without creating a second layout model.
  **Partial**: Latin vector paint + title/`lang`/heading outline bootstrap are
  in place; PDF bookmarks include Writer outline-level paragraphs
  (`p[data-office-outline-level]`) alongside h1–h6 with level nesting through
  `0.148.0`; Writer underlines paint as vector paths through `0.149.0`; Writer
  text highlights paint as vector fill strips through `0.150.0`; Writer
  paragraph borders paint as vector strokes through `0.151.0`; Writer
  `between` and `bar` paragraph border edges paint as vector strokes through
  `0.152.0`; catalog `/MarkInfo` and vector-run `/Span` ActualText deepen the
  tagged bootstrap through `0.153.0`; a catalog-linked StructTreeRoot stub
  (Document + outline-derived H1–H6/P with empty `/K`) lands through
  `0.154.0`; ParentTree / MCID content links for vector-run Spans land through
  `0.155.0`; Writer `wave` / `doubleWave` paragraph borders paint as explicit
  sine polylines through `0.156.0`; Writer `threeDEmboss` / `threeDEngrave` /
  `inset` / `outset` paragraph borders paint as dual-tone highlight+shadow
  offsets through `0.157.0`; Writer `zigZag` / `zigZagStitch` art paragraph
  borders paint as chevron polylines and catalog `/DisplayDocTitle` lands
  through `0.158.0`; Writer `sawtooth` / `sharksTeeth` art paragraph borders
  paint as triangular teeth through `0.159.0`; Writer `triangles` /
  `triangle1` / `triangle2` art paragraph borders paint as closed isosceles
  triangles through `0.160.0`; Writer `ovals` / `rings` art paragraph borders
  paint as ellipse motifs through `0.161.0`; Writer `marquee` /
  `marqueeToothed` art paragraph borders paint as rectangle motifs through
  `0.162.0`; Writer `moons` art paragraph borders paint as crescent motifs
  through `0.163.0`; Writer `basicBlackSquares` / `basicWhiteSquares` art
  paragraph borders paint as discrete square stamps through `0.164.0`; Writer
  `basicBlackDots` / `basicWhiteDots` art paragraph borders paint as discrete
  circular stamps through `0.165.0`; Writer `basicBlackDashes` / `basicWhiteDashes` art
  paragraph borders paint as discrete dash stamps through `0.166.0`; Writer
  `basicThinLines` art paragraph borders paint as parallel hairlines through
  `0.167.0`; Writer `basicWideInline` / `basicWideMidline` /
  `basicWideOutline` art paragraph borders paint as thick geometric rails
  through `0.168.0`; optional host
  `registerWorkPdfCjkFont` TrueType enables searchable CJK vector text; full
  PDF/UA certification, remaining decorative art borders, and bundled CJK fonts
  remain open.

Exit criteria: representative Traditional Office/Word fixtures reopen without
unreported data loss; edited native structures retain identity; malformed
inputs fail closed;
pagination and export have deterministic structural and visual evidence.

### R1 — Writer daily-work parity (P0/P1)

The typed Table of Contents slice is complete: insert, selected-block
customization, navigation, explicit refresh, native cached entries, and DOCX
reopen all share the heading outline and pagination resolver.

The native index slice is complete: primary and secondary markers,
cross-references, page emphasis, duplicate-page merging, click navigation,
selected-entry/block customization, explicit all-index refresh, bounded cached
rows, and DOCX reopen all share one typed marker graph and the pagination field
resolver.

The bounded compare/combine slice is complete for same-layout paragraphs and
headings. Compare imports DOCX, HTML, or TXT and produces reviewable text,
format, and bounded same-section cross-paragraph move revisions; Combine admits
a reviewed copy only when rejecting every imported revision exactly reconstructs
the current baseline. Exact text-only whole-paragraph changes now export and
reopen with native paragraph-mark records. Complex structural changes, isolated
paragraph-break merge/split or mixed-content paragraph-mark revisions,
cross-section/range move generation, and multi-copy conflicts remain follow-up
work rather than silent approximations. See
[Remaining roadmap (product-enough bar)](#remaining-roadmap-product-enough-bar)
for the ordered R1 leftovers after R0 exit.
- Keep the bounded Writer picture transform slice native: quarter-turn rotation
  and horizontal/vertical reflection round-trip through DrawingML `a:xfrm`;
  arbitrary angles remain an explicit diagnostic boundary.
- Extend the bounded text-box preset model to a typed DrawingML shape/connector
  model before expanding to WordArt, charts, and SmartArt. The isolated WPS
  straight, elbow, and curved connector subset now has independent native
  fixtures, typed SVG/DrawingML projections, and editor contracts. Keep
  arbitrary routing, mixed paragraphs, and unsupported branches fail-closed
  until each has its own native fixture and editor contract.
- The bounded field-authoring slice is complete: typed insert/edit settings for
  page, section, and bookmark page-reference formats, WPS numeric switches,
  date/time presets, bookmark targets, hyperlinks, MERGEFORMAT preservation,
  responsive discovery, and A3S Test desktop/390px phone evidence. Keep
  nested fields, document properties, mail fields, and unsupported switches
  fail-closed until each has a native contract.
- The local WPS UI-reference slice is complete: typed `shell`, `fields`, and
  `all` COM profiles record the real Writer window, Ribbon/status shell, field
  command bars, command IDs, and WPS build metadata as deterministic JSON.
  These receipts inform accessible command naming and discoverable field
  workflows; A3S Test Web/CDP remains the primary browser UI contract and the
  locked Windows CUA profile remains unsupported.
- The Writer WPS shortcut slice now pairs the existing ACL with desktop and
  compact visual evidence for format copy/paste, text-case menu discovery,
  atomic undo, paragraph formatting, spelling/review toggles, and editor-focus
  preservation.
- Extend common field instructions and reference workflows beyond this bounded
  authoring slice only with independent native fixtures and typed contracts.
- Extend the bounded content-control slice beyond direct inline text and rich
  text only after block, form, and data-binding semantics have independent
  typed models and fail-closed native fixtures.
- Add mail merge only after fields and data-source contracts are stable.

Exit criteria: the most common Traditional Office Writer report, contract,
academic-paper, and review workflows can be completed without leaving the
embedded editor.

### R2 — Spreadsheet calculation and analysis parity (P0/P1)

- Build a versioned formula-compatibility corpus against Traditional
  Office/Excel behavior, including locale, errors, arrays, volatility, dates,
  and dependency updates.
- Treat common native totals-row authoring as supported: retain per-column
  aggregate/label/custom-formula controls, filtered-row-aware `SUBTOTAL`,
  dense/sparse reconciliation, manual-cell authority, Yjs convergence, and
  exact XLSX totals metadata as permanent gates. Continue with slicers,
  external/query-table boundaries, and a broader structured-reference formula
  corpus. Automatic calculated-column fill is now
  supported for newly inserted body rows with manual-exception and conflict
  safeguards.
- Complete advanced sort/filter and conditional-format precedence.
- Expand pivot caches, grouping, calculated fields, slicers/timelines, styles,
  and pivot charts.
- Improve drawing/image fidelity before adding specialist analysis tools.

Exit criteria: common finance, operations, reporting, and pivot-analysis
workbooks recalculate and round-trip with explicit, bounded diagnostics.

### R3 — Presentation expressiveness (P1)

- Treat composable bounded entrance/exit effects, object-centric authoring,
  browser playback, collaboration, and native PPTX timing-tree round trips as
  a completed slice.
- Extend the animation model with broader native PPTX preservation, emphasis,
  motion paths, trigger-on-object, and round-trip-safe unsupported states.
- Add safe audio/video relationships and deterministic playback policy.
- Add visual master/layout editing and dedicated sorter/master workflows.
- Expand shapes, connectors, effects, SmartArt fallback, and chart fidelity.
- Extend slideshow evidence to timings, media, presenter tools, and printing.

Exit criteria: a typical Traditional Office sales, training, or classroom deck
retains its visual hierarchy and can be presented without losing
animation/media intent.

### R4 — PDF workbench (P1/P2, active)

- Treat page organization as the first completed PDF-workbench slice: retain
  insert, delete, rotate, reorder, extract, merge, and split; one Blob
  replacement and one Undo record per mutation; native-history-first toolbar
  ordering; dedicated-Worker execution; 256 MiB primary, 128 MiB merge, and
  4,096-page bounds; fail-closed signed/encrypted/risky-structure handling; and
  PDFium plus independent-parser save/reopen evidence as permanent gates.
- Add native text/image/link editing with strict font and content-stream
  fallbacks.
- Add provider contracts for OCR and conversion; keep confidence and geometry
  in the typed result.
- Add e-sign, watermark, true redaction, protection, and optimization with
  explicit security acceptance tests.

Exit criteria: save/reopen validates every mutation against PDFium and a second
independent parser; redaction and signatures meet byte-level safety rules.

### R5 — Host collaboration and AI reference integrations (P1)

- The transport-neutral Yjs/Yrs document layer now includes durable native
  replicas plus bounded, resumable CLI and MCP event cursors. Checkpoint gaps
  produce an explicit full-state reset instead of silently dropping history.
- The browser Core now exposes an identity-bound, bounded host-channel adapter
  for Yjs v1 state-vector/update synchronization and an ephemeral Awareness
  contract for typed actors, modes, activity, and format-specific locations.
  Typed incremental origins survive live transport without entering canonical
  document state. All five editors now project one accessible participant
  roster across edit and preview chrome.
- The native JSONL host transport now carries bounded, typed, ephemeral
  Awareness state for agents and projects format-specific selections and
  locations into the existing editor canvases without embedding accounts,
  storage, or authorization.
- Map comments, revisions, selections, cursors, permissions, and conflict
  handling onto the existing controlled editor model.
- Publish provider-neutral examples for writing, formula, slide, and PDF AI
  tasks using immutable context and bounded mutations.

Exit criteria: two independent hosts can implement collaboration and AI without
scraping the UI or depending on private editor internals.

### R6 — Enterprise and long-tail compatibility (P2/P3)

- Add encrypted-file provider support, signature validation, accessibility
  metadata, localization/RTL depth, and policy-driven passive macro handling.
- Expand legacy format conversion and specialist Office features only when
  authoritative fixtures and no-clobber rules exist.
- Keep VBA, ActiveX, and untrusted add-in execution outside the browser editor.

## Definition of done for a roadmap row

A visible command or “Supported” claim is complete only when all applicable
items below exist:

1. A typed canonical model and bounded validation rules.
2. Executable desktop and compact-Web UI; no placeholder command.
3. One user intent produces one controlled update and one undo record.
4. Import, edit, export, reopen, and structural assertions for native files.
5. Strict/transitional namespace, malformed, duplicate, spoofed, oversized,
   and unsafe relationship tests where the format permits them.
6. Compatibility diagnostics for every intentional normalization.
7. Keyboard, focus restoration, accessible naming, and touch-target evidence.
8. Traditional Office/Word/Excel/PowerPoint or PDF reference fixtures where
   visual/layout fidelity matters.
9. TypeScript, unit/component tests, lint, Rust/WASM, and production builds pass.

## UI evidence and regression gates

The shared editor shell now has a dedicated A3S Test Kit UI gate. Every active
editor publishes an explicit component boundary with typed facts, while the
development Playground exposes bounded style/layout/motion snapshots and
state diffs through `@a3s-lab/testkit@0.6.2`. The gate exercises desktop and
compact interactions for the home, Writer, Spreadsheet, Presentation,
Markdown, and PDF surfaces, and verifies focus, dialog geometry, responsive
conditions, accessibility evidence, screenshots, and clean browser
diagnostics. The standalone `a3s-test` runner is the primary interaction
authority; Playwright validates the live bridge and supplemental pixel
assertions. The Commander operator also exposes the bounded A3S Test agent
lifecycle and reports the locked CUA Driver/MCP certification matrix before
native GUI work. Production preview intentionally aliases the SDK to a no-op, so this
dev-only evidence is kept separate from the static release gate and is chained
by `bun run test:e2e`.

The gate is a permanent P0 regression requirement for new editor-shell or
responsive UI work. A feature is not considered UI-complete until it has a
stable semantic target, a component owner, a typed state transition, and
desktop/compact evidence in `tests/e2e/office-testkit-ui.acl` or the matching
Playwright Test Kit spec.

## Evidence and source baseline

A3S Office status is derived from the current repository, especially
[README.md](README.md), [PRODUCT.md](PRODUCT.md), the compatibility diagnostics,
and deterministic suites under [`tests/`](tests/) and [`tests/e2e/`](tests/e2e/).

The Traditional Office baseline uses public format specifications and mature
desktop/web suite documentation rather than one product release:

- [ECMA-376 Office Open XML](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [Word help and learning](https://support.microsoft.com/word)
- [Excel help and learning](https://support.microsoft.com/excel)
- [PowerPoint help and learning](https://support.microsoft.com/powerpoint)
- [Office help and learning](https://support.microsoft.com/office)
- [Acrobat feature overview](https://www.adobe.com/acrobat/features.html)

Review this baseline whenever a Traditional Office capability materially
changes or an A3S Office release changes a row's status.
