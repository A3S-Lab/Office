# Word Workflows

Use stable one-based semantic paths such as `/body/p[1]`,
`/body/tbl[1]/tr[1]/tc[1]`, and header/footer paths returned by `get` or
`outline`. Discover paths instead of guessing them.

## Inspect

```bash
a3s use office native get report.docx /body --depth 2 --json
a3s use office native query report.docx 'p[style=Heading1]' --json
a3s use office native view report.docx text --json
a3s use office native view report.docx annotated --limit 200 --json
a3s use office native view report.docx issues --type content --json
```

## Create and Edit

```bash
a3s use office native create report.docx --json
a3s use office native add report.docx /body --type paragraph --text 'Summary' --json
a3s use office native add report.docx /body --type table --rows 2 --columns 3 --json
a3s use office native set report.docx '/body/p[1]' --text 'Updated summary' --json
a3s use office native set report.docx /body --find Draft --replace Final --json
a3s use office native set report.docx / --find 'Q([1-4]) 2025' --replace 'Q$1 2026' --regex --json
a3s use office native set report.docx '/body/p[1]/r[1]' --bold true --italic false --underline double --script superscript --strikethrough true --double-strikethrough false --text-case small-caps --highlight yellow --language en-US --font-family Aptos --font-size 14 --text-color 123456 --json
a3s use office native set report.docx '/body/p[1]' --align center --json
a3s use office native add report.docx '/body/p[1]' --type hyperlink --url https://example.com/report --display 'Open report' --tooltip 'A3S report' --json
a3s use office native set report.docx '/body/p[1]/hyperlink[1]' --location section_1 --display 'Jump to section' --json
a3s use office native set report.docx '/header[1]/p[1]' --url https://example.com/header --display 'Header link' --json
a3s use office native query report.docx hyperlink --json
a3s use office native remove report.docx '/body/p[1]/hyperlink[1]' --json
a3s use office native add report.docx '/body/p[1]' --type comment --author Alice --initials AL --text 'Please reword this' --json
a3s use office native set report.docx '/comments/comment[1]' --author Bob --initials BO --text 'Reviewed' --json
a3s use office native query report.docx comment --json
a3s use office native remove report.docx '/comments/comment[1]' --json
a3s use office native add report.docx /body --type picture --input chart.png --alt 'Quarterly revenue chart' --json
a3s use office native move report.docx '/body/p[2]' --before '/body/p[1]' --json
```

General replacement is case-sensitive and literal by default. `--regex`
enables Rust regular expressions with `$1`/`$name` capture expansion. `/`
includes the main document, headers, footers, footnotes, endnotes, and legacy
comments. Use `/body`, a header/footer, paragraph, run, table/cell, hyperlink,
or comment path to keep the edit narrower. Matches may span runs; replacement
text uses the first matched run's formatting and does not flatten later runs.
Zero matches return an unchanged success.

Character formatting targets a run returned by `get --depth 2`. Paragraph
alignment targets the paragraph itself. Supported typed properties are bold,
italic, `none`/single/double underline, baseline/superscript/subscript, explicit
single and double strikethrough, `none`/small-caps/all-caps display case, a
portable highlight, one conservative BCP-47 language tag, font family, RGB text
color, and font size. Word sizes must be exact half-point increments. The
highlight palette is `none`, `black`, `blue`, `cyan`, `dark-blue`, `dark-cyan`,
`dark-gray`, `dark-green`, `dark-magenta`, `dark-red`, `dark-yellow`, `green`,
`light-gray`, `magenta`, `red`, `white`, or `yellow`. Language mutation writes
the primary `w:lang` slot and preserves existing East Asian and complex-script
slots. Advanced styles, inheritance, scheme shading, extended underline
variants, character spacing, per-script language mutation, and RTL mutation
remain outside the typed native subset.

Hyperlink creation accepts a body, header, or footer paragraph path and returns
a stable `/hyperlink[N]` child path for update or removal. Each part stores and
cleans up its own external relationship. External targets accept only absolute
HTTP, HTTPS, or mailto URIs without credentials. Internal targets are 1-40
character Word bookmark names beginning with a letter or underscore. Display
text and tooltips are optional. Reads and rendering keep external links inert;
they never fetch them.

Legacy comments accept a main-document paragraph or run parent and return a
stable `/comments/comment[N]` path. The comment is range-anchored in the owner;
semantic reads report the containing paragraph as `anchoredTo`. Add requires an
author and plain text; initials are optional. Update author, initials, or text
through the comment path, and remove through the ordinary `remove` command.
Removing the owning paragraph or run also removes its comment. Replies,
resolved state, writable dates, rich comment bodies, `commentsExtended.xml`,
and header/footer anchors are not yet native.

Use `copy`, `swap`, and `remove` only with paths returned by a fresh read.
Identity-bearing copies, cross-parent ownership migration, and rich structures
outside the documented native subset fail closed.

`add-part --type header|footer|chart` creates a typed part carrier and
relationship; it does not make a chart or header visible by itself.

## Merge and Verify

```bash
a3s use office native merge template.docx report.docx --data @report.json --json
a3s use office native validate report.docx --json
a3s use office native view report.docx issues --json
a3s use office native view report.docx html --output report.html --json
a3s use office native view report.docx svg --output report.svg --json
a3s use office native watch report.docx --port 0
```

Tracked changes, comment replies/resolution and modern metadata, complete
fields/forms, TOC updates, equations,
advanced styles, and layout-accurate pagination are not yet complete native
capabilities. HTML, SVG, and live watch are semantic previews, not Word
pagination. Watch follows saved disk revisions and is not an interactive Word
editor. Use the compatibility route only after checking provider readiness and
explaining the boundary.
