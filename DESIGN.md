# A3S Office product home

The product home follows the same light product-system language as A3S Cloud
while remaining separate from the versioned documentation site and the
Playground.

## Visual decisions

- **Canvas:** white is the primary canvas. Full-width chapters alternate
  between neutral white, violet-tinted workflow, green-tinted collaboration,
  and blue-tinted action surfaces within a 1,440px product frame.
- **Palette:** A3S blue (`#1264ff`) leads the brand and primary actions. Green
  (`#0c8f69`), violet (`#7157c9`), and amber (`#a96906`) distinguish product
  states without turning the page into a dark developer-tool interface.
- **Dark surfaces:** dark navy appears only inside the installation command,
  where it communicates executable code. Product copy, diagrams, editor
  chrome, collaboration, and calls to action stay on light surfaces.
- **Type:** Geist and Geist Mono, inherited from the documentation shell. Large
  headings are left aligned with restrained tracking. Monospace is limited to
  commands, component names, protocol labels, and measured state.
- **Hero visual:** a light surface map communicates the product boundary:
  editor surfaces, shared state, collaborators, and the A3S Boot host
  contract. The hero never mounts a full editor runtime.
- **Composition:** the first viewport moves from product promise to the surface
  map, then to a five-surface contract bar. The editor section is a deliberate
  chapter sequence: Document, Markdown, Spreadsheet, Presentation, and PDF.
  Each chapter pairs the real editor's interaction language with a lightweight
  UI/UX motion preview and a documentation entry point. The complete editable
  editors remain in Playground.
- **Motion:** each chapter preview animates one recognizable interaction, such
  as a revision/comment, source-to-preview sync, cell selection, an entrance
  cue, or page annotation. A shared pause control stops every preview;
  `prefers-reduced-motion` disables authored motion without hiding content and
  exposes that system-controlled state on the disabled motion control.
  A sticky current-chapter meter keeps the reader oriented through the long
  sequence without introducing a second navigation model. The heading, pause
  control, chapter index, and meter stay painted while the heavy preview stack
  uses `content-visibility: auto` with viewport-calibrated intrinsic reserves.
  This preserves deep-link scroll stability without making a returning user
  click into an unpainted control.
- **Preview boundary:** the chapter header carries a persistent
  `Illustrative UI/UX · preview only` note, and simulated window status values
  use `Sample`/`示意` wording. Each chapter exposes a collapsed, non-modal
  recovery disclosure with base-aware links to the component documentation and
  Playground. These cues make sample presence, autosave, and rendering states
  explicit without adding noise to the healthy path.
- **Runtime boundary:** the product home imports only the preview components.
  It does not load DocumentEditor, MarkdownEditor, SpreadsheetEditor,
  PresentationEditor, or PdfViewer runtimes, keeping the first paint small and
  leaving editing state and large document work to Playground.

## Responsive behavior

The hero switches to one column below 960px, and the product window hides its
secondary host panel below 520px. Assurance items scroll horizontally on
narrow screens, while workflow, collaboration peers, editor entries, and final
actions become single-column reading paths. The global A3S navigation remains
the source of truth for Product, Docs, and Playground wayfinding.

## Content contract

The latest-capability rail, collaboration links, A3S Boot backend route,
Playground route, and editor names are stable product contracts. New homepage
treatments must keep these destinations base-aware so Rspress can rewrite them
for root and GitHub Pages deployments. Product-only CSS belongs under
`website/product-theme/`; documentation typography and syntax highlighting
remain owned by `website/theme/`.

The five chapter anchors (`#editor-chapter-document`,
`#editor-chapter-markdown`, `#editor-chapter-spreadsheet`,
`#editor-chapter-presentation`, and `#editor-chapter-pdf`) are part of the
homepage contract. Keep their labels, preview captions, and documentation
links synchronized when an editor is renamed or its route changes. The
`data-current-chapter` value and progress meter are derived from those same
anchors; do not add a parallel chapter registry.
