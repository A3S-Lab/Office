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
- **Hero visual:** a light editor-system window communicates actual product
  boundaries: editor surfaces, shared state, collaborators, and the A3S Boot
  host contract. It is synthetic UI chrome, not a claim about real users or
  production scale.
- **Composition:** the first viewport moves from product promise to the editor
  system, then to a five-surface contract bar. Workflow, collaboration, editor
  surfaces, and the final action use distinct chapter colors and layouts.
- **Motion:** the editor system has one short arrival transition. Hover and
  focus-visible states remain functional, and `prefers-reduced-motion`
  disables authored motion without hiding content.

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
