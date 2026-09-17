# R0 no-clobber corpus

Permanent R0 gate for representative DOCX round trips.

## Run

```bash
bun run test:corpus:r0
```

## Contract

Each fixture must prove:

1. Import → light/no-op edit path → export → reopen does not silently drop
   critical identities (bookmarks, revision authors/text, review comments,
   internal and external links, headers/footers, footnotes, endnotes, text
   content controls, PAGE/PAGEREF/NUMPAGES/SECTION/SECTIONPAGES/NUMWORDS/
   NUMCHARS/DATE/TIME fields (including DATE `\@` format switches), XE/INDEX
   index identities, TOC options/entries, table cells).
2. Every intentional normalization appears in `compatibility.issues` with a
   stable code.
3. Active content (VBA, package signatures) stays fail-closed on export.

Current fixtures cover report bookmarks/external links, internal
bookmark-anchor hyperlinks, review track changes, review comments, default
headers/footers, footnotes, endnotes, text content controls, body PAGE,
bookmark-backed PAGEREF, NUMPAGES, SECTION, SECTIONPAGES, NUMWORDS, NUMCHARS,
DATE (bare and `\@` format), and TIME fields, XE index entries and INDEX field
columns, TOC level/hyperlink options and cached entry titles, contract tables,
duplicate-bookmark identity diagnostics, and active-content fail-closed
export.

Grow fixtures here when a real Traditional Office sample still fails the
contract. Do not grow decorative PDF art-border coverage in place of this gate.
