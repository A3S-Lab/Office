# A3S Office product home

The product home follows the same light product-system language as A3S Cloud
while remaining separate from the versioned documentation site and the
Playground.

## Visual decisions

- **Canvas:** a cool light canvas (`#f5f7fb`) frames white editor surfaces and
  the same centered product grid used by A3S Flow. Sections use whitespace and
  restrained surface changes for rhythm; decorative divider rules stay out of
  the product home.
- **Palette:** one blue accent (`#1456f0`) owns brand emphasis, links, active
  controls, and primary actions. Green is reserved for semantic sync/success
  status, while neutral slate tones carry copy, borders, and editor chrome.
  This keeps the page recognizable as one product instead of a color-coded
  collection of editors.
- **Light-first surfaces:** the product home has no dark installation panel.
  Code and protocol examples keep their own documentation syntax theme, while
  product copy, diagrams, editor chrome, collaboration, and calls to action
  remain readable on light surfaces.
- **Type:** Geist and Geist Mono, inherited from the documentation shell. Large
  headings are left aligned with restrained tracking. Monospace is limited to
  commands, component names, protocol labels, and measured state.
- **Hero visual:** a light surface map communicates the product boundary:
  editor surfaces, shared state, collaborators, and the A3S Boot host
  contract. The hero never mounts a full editor runtime.
- **Composition:** the first viewport moves from the product promise to a
  poker-hand stack of five real editor captures, then directly into the editor
  chapter sequence: Document,
  Markdown, Spreadsheet, Presentation, and PDF. Each chapter pairs the real
  editor's interaction language with a lightweight UI/UX motion preview and a
  documentation entry point. The complete editable editors remain in
  Playground.
- **Motion:** each chapter preview animates one recognizable interaction, such
  as a revision/comment, source-to-preview sync, cell selection, an entrance
  cue, or page annotation. The collaboration diagram adds a restrained state
  packet and service-status pulse so the backend path reads as live without
  competing with the editor chapters. A shared pause control stops every
  chapter preview;
  `prefers-reduced-motion` disables authored motion without hiding content and
  exposes that system-controlled state on the disabled motion control.
  A sticky current-chapter meter keeps the reader oriented through the long
  sequence without introducing a second navigation model. The heading, pause
  control, chapter index, and meter stay painted while the heavy preview stack
  uses `content-visibility: auto` with viewport-calibrated intrinsic reserves.
  This preserves deep-link scroll stability without making a returning user
  click into an unpainted control.
- **Preview boundary:** the chapter and collaboration surfaces use captures
  produced by the visual-test suite. Capture labels, implementation metadata,
  and recovery instructions are intentionally omitted from the healthy path so
  the editor UI remains the visual focus. The surrounding headings and links
  provide the only necessary context and recovery route.
- **Collaboration scene:** the homepage includes a compact three-peer scene
  using the real Document editor capture for Member A, Member B, and A3S Agent.
  Colored cursors and selections move over each capture while state packets
  travel to an A3S Boot room node. The scene is illustrative and does not mount
  a live editor runtime; the collaboration and backend links lead to the full
  implementation documentation.
- **Runtime boundary:** the product home imports only the preview components.
  It does not load DocumentEditor, MarkdownEditor, SpreadsheetEditor,
  PresentationEditor, or PdfViewer runtimes, keeping the first paint small and
  leaving editing state and large document work to Playground.
- **Playground shell:** Playground is an editing workspace, not another
  marketing surface. It has no global top navigation and consumes the full
  dynamic viewport. The A3S Office logo and name live in the sidebar as the
  product-home link; product navigation stays in the documentation shell and
  is not duplicated above the editor.

## Responsive behavior

The hero switches to one column below 1,040px, and compact chapter/window
controls collapse below 720px (with a final 420px pass for very narrow phones).
The product window hides its secondary host panel on compact screens.
The chapter index scrolls horizontally when needed, while workflow,
collaboration peers, editor entries, and final actions become
single-column reading paths. Chapter deep links reserve space for both the
global navigation and sticky progress meter, so the destination heading is not
hidden behind chrome. The global A3S navigation remains the source of truth for
Product, Docs, and Playground wayfinding.

## Content contract

Collaboration links, the A3S Boot backend route, the Playground route, and
editor names are stable product contracts. Latest capabilities remain
discoverable in the documentation homes, README, and Playground; the product
home does not repeat them in a separate rail. New homepage treatments must keep
these destinations base-aware so Rspress can rewrite them for root and GitHub
Pages deployments. Product-only CSS belongs under
`website/product-theme/`; documentation typography and syntax highlighting
remain owned by `website/theme/`.

The five chapter anchors (`#editor-chapter-document`,
`#editor-chapter-markdown`, `#editor-chapter-spreadsheet`,
`#editor-chapter-presentation`, and `#editor-chapter-pdf`) are part of the
homepage contract. Keep their labels, preview captions, and documentation
links synchronized when an editor is renamed or its route changes. The
`data-current-chapter` value and progress meter are derived from those same
anchors; do not add a parallel chapter registry.
