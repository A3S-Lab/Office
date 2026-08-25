# A3S Office documentation home

The documentation home uses a Cloud-inspired product-system direction while
remaining a documentation surface rather than a marketing page.

## Visual decisions

- **Canvas:** a cool white/blue canvas (`#f7faff`) with soft radial light
  fields. Content is contained to 1,440px and keeps the existing A3S UI
  navigation and documentation tree.
- **Accent:** electric blue (`#1264ff`) is the single page accent. Navy is
  reserved for collaboration and installation surfaces so code and shared
  state read as infrastructure.
- **Type:** Geist and Geist Mono, inherited from the documentation shell. Large
  headlines are left aligned and split into a dark promise plus a blue
  collaboration phrase.
- **Hero visual:** a code-native editor system window communicates actual
  product boundaries: editor surfaces, shared state, collaborators, and the
  A3S Boot host contract. It is synthetic UI chrome, not a claim about real
  users or production scale.
- **Composition:** the first viewport moves from promise to system visual, then
  to a five-surface assurance bar. Workflow, collaboration, editor surfaces,
  and the final call to action each use a different layout family.
- **Motion:** only small hover lifts and focus-visible states are used. The
  page respects `prefers-reduced-motion` and does not depend on animation for
  comprehension.

## Responsive behavior

The documentation sidebar remains available for wayfinding on desktop and
collapses through the existing A3S UI mobile navigation. The hero switches to
a single column below 960px; the product window hides its secondary host panel
below 520px. Assurance items scroll horizontally on narrow screens, while
workflow and editor entries become single-column reading lists.

## Content contract

The latest-capability rail, collaboration links, Playground route, and editor
names are stable documentation contracts. New homepage treatments must keep
those routes relative (`./...`) so Rspress can rewrite them for root and
GitHub Pages deployments.
