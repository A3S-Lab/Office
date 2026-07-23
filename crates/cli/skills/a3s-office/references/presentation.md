# Presentation Workflows

Use stable one-based paths returned by `outline` or `get`, such as `/slide[1]`,
`/slide[1]/shape[1]`, and table row/cell paths.

## Inspect

```bash
a3s use office native get deck.pptx '/slide[1]' --depth 3 --json
a3s use office native query deck.pptx 'shape:contains("Roadmap")' --json
a3s use office native view deck.pptx outline --json
a3s use office native view deck.pptx annotated --limit 200 --json
a3s use office native view deck.pptx issues --json
```

## Create and Edit

```bash
a3s use office native create deck.pptx --json
a3s use office native add deck.pptx / --type slide --text 'Roadmap' --json
a3s use office native add deck.pptx '/slide[1]' --type shape --text 'Q3' --json
a3s use office native add deck.pptx '/slide[1]' --type table --rows 3 --columns 4 --json
a3s use office native add deck.pptx '/slide[1]' --type picture --input diagram.png --alt 'System architecture' --json
a3s use office native set deck.pptx '/slide[1]/shape[1]' --text 'Updated roadmap' --json
a3s use office native set deck.pptx '/slide[1]' --find Draft --replace Final --json
a3s use office native set deck.pptx '/slide[1]/notes' --find internal --replace confidential --json
a3s use office native set deck.pptx '/slide[1]/shape[1]/paragraph[1]/run[1]' --bold true --underline double --script subscript --text-case all-caps --highlight cyan --language zh-CN --font-family 'Aptos Display' --font-size 20 --text-color AA2200 --json
a3s use office native set deck.pptx '/slide[1]/shape[1]/paragraph[1]' --align center --json
a3s use office native set deck.pptx '/slide[1]/shape[1]' --url https://example.com/slides --tooltip 'Open slides' --json
a3s use office native set deck.pptx '/slide[1]/shape[1]/hyperlink' --location 'slide[2]' --tooltip 'Next slide' --json
a3s use office native query deck.pptx hyperlink --json
a3s use office native remove deck.pptx '/slide[1]/shape[1]/hyperlink' --json
a3s use office native add deck.pptx '/slide[1]' --type comment --author Alice --initials AL --text 'Rework this slide' --x-emu 914400 --y-emu 457200 --json
a3s use office native set deck.pptx '/slide[1]/comment[1]' --author Bob --initials BO --text 'Reviewed' --x-emu 1828800 --y-emu 914400 --json
a3s use office native query deck.pptx comment --json
a3s use office native remove deck.pptx '/slide[1]/comment[1]' --json
a3s use office native swap deck.pptx '/slide[1]' '/slide[2]' --json
```

General replacement is literal by default and may span rich-text runs. Add
`--regex` for Rust regex with capture expansion. `/` covers all slide and notes
text; a slide, object, table-cell, paragraph, or run path narrows the visible
slide content, while `/slide[N]/notes` targets speaker notes explicitly. A
slide scope does not implicitly alter its notes. Zero matches are an unchanged
success.

Character formatting targets a run path; alignment targets its paragraph.
This works for shape and table-cell text returned by a semantic read. The typed
subset covers bold, italic, `none`/single/double underline,
baseline/superscript/subscript, `none`/small-caps/all-caps display case, the
same portable 17-value highlight palette as Word, one conservative BCP-47
language tag, font family, exact centipoint size, RGB text color, and horizontal
alignment. `highlight=none` removes the run highlight. Presentation single and
double strikethrough are rejected explicitly rather than ignored. Shape-wide
propagation, theme authoring, fills, outlines, and effects remain incomplete.

Native Presentation hyperlinks are shape-wide click links. Target a shape or
its returned `/hyperlink` child path and use an optional tooltip. External
targets accept only absolute HTTP, HTTPS, or mailto URIs without credentials.
Internal targets accept an existing `slide[N]` or `/slide[N]` semantic path and
write the standard slide-jump action. A link may switch between external and
internal targets without leaking its old relationship. Separate display text
is unsupported because the shape retains its existing text. Reads and previews
never fetch external targets.

Legacy slide comments return `/slide[N]/comment[M]` paths. Add requires an
author and plain text, accepts optional initials, and optionally accepts both
`--x-emu` and `--y-emu`; coordinates are rejected unless both are present.
Updates may change author, initials, text, or the complete position. The native
engine manages the shared author list and per-author indexes and removes a
slide's comment part when the slide is removed. Modern PowerPoint threaded
comments, replies, resolved state, writable dates, and rich bodies are not yet
native.

Basic table rows, cells, and virtual columns support bounded structural edits.
Merged cells, rich table styles, relationship-owning object copies, advanced
charts/media, animations, transitions, and theme fidelity remain incomplete.

## Verify and Preview

```bash
a3s use office native validate deck.pptx --json
a3s use office native view deck.pptx issues --type format --json
a3s use office native view deck.pptx svg --output deck.svg --json
a3s use office native view deck.pptx screenshot --output deck.png --json
a3s use office native watch deck.pptx --port 0
```

SVG, HTML, PNG, and live watch output are semantic previews. They do not prove
PowerPoint layout, animation, morph, font, or theme fidelity. Watch refreshes
the full saved preview and does not provide slide-scoped editing or annotations.
