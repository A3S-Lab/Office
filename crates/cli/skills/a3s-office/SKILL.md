---
name: a3s-office
description: Inspect, create, edit, validate, merge, render, and live-preview .docx, .xlsx, and .pptx files through the A3S Use native CLI or standard MCP server, with explicit fallback to its pinned OfficeCLI compatibility route for unpromoted operations. Use when an agent needs to work with Microsoft Office OOXML documents, automate Word, Spreadsheet, or Presentation changes, diagnose document issues, preview saved changes, or choose between native and compatibility Office capabilities.
---

# A3S Office

Use A3S Use as the application boundary for Office documents. Prefer the
in-process native engine and its typed operations. Use the compatibility route
only when the requested operation is not yet native.

Use the host surface that is already available:

- In an A3S Code `use` worker, call the available
  `mcp__use_office__*` tools directly. The host has already started the native
  MCP server and owns its lifecycle; do not run shell commands or install a
  provider.
- If a requested operation is absent from the native tools, use an available
  `mcp__use_office_compat__*` tool only as an explicit compatibility fallback.
  If that route is absent, report the missing capability instead of installing,
  repairing, or falling back to a shell.
- In a CLI-only agent host, use the `a3s use office native ...` commands below.

## Workflow

1. Identify and inspect the document before changing it.

   In an A3S Code `use` worker, begin with
   `mcp__use_office__office_validate`, then open a session and use
   `mcp__use_office__office_view`, `office_get`, or `office_query` as needed.
   In a CLI-only host, use:

   ```bash
   a3s use office native validate "$FILE" --json
   a3s use office native view "$FILE" annotated --limit 200 --json
   a3s use office native view "$FILE" outline --json
   a3s use office native view "$FILE" issues --json
   ```

2. In a CLI-only agent host, load the format reference relevant to the task.
   An A3S Code `use` worker cannot read Skill reference files; rely on this
   guidance and the available MCP tool schemas instead.

   - Read [references/word.md](references/word.md) for `.docx`.
   - Read [references/spreadsheet.md](references/spreadsheet.md) for `.xlsx`.
   - Read [references/presentation.md](references/presentation.md) for `.pptx`.

3. Prefer a typed `office native` operation. Use `--output` for a distinct
   result when the command supports it; otherwise work on an intentional copy.
   Use one atomic `batch` for dependent changes. For general find/replace, use
   `set <file> <scope> --find ... --replace ...`; use literal mode unless regex
   captures are actually required.

4. Verify the result with `validate`, a targeted `get` or `query`, and
   `view ... issues`. Use HTML, SVG, or screenshot only as a semantic preview.

5. Report the exact output path and any remaining issue records. Do not claim
   Microsoft Office layout fidelity from a semantic preview.

For an iterative visual loop, run the foreground watch in a separate process:

```bash
a3s use office native watch "$FILE" --port 0
```

Open the authenticated loopback URL printed at startup. The watch refreshes
after on-disk saves, including separate native CLI mutations. It is read-only,
does not expose OfficeCLI's resident protocol, and does not observe unsaved MCP
session changes until `office_save`. Stop it with Ctrl+C; use `--timeout-ms`
when an agent must bound its lifetime.

## Choose the Surface

- In an A3S Code `use` worker, use `mcp__use_office__*` and keep the returned
  Office session ID stable until the document is saved and closed.
- Use `mcp__use_office_compat__*` only when the native route lacks the requested
  operation and the compatibility tools are actually present.
- Use `a3s use office native ... --json` for local automation and scripts.
- Use `a3s use mcp serve office-native` for typed, stateful agent sessions.
  Read [references/mcp.md](references/mcp.md) before using its session tools.
- Use the typed Rust API when embedding Office behavior in Rust.
- Use `a3s use office ...` only for an operation absent from the native route.
  Check `a3s use office doctor --json` first. The first real compatibility
  command prepares the pinned OfficeCLI provider when policy allows it; help,
  version, doctor, Skills, and native commands remain non-installing.

`a3s-use` accepts the same arguments when the umbrella `a3s` executable is not
available.

## Safety Rules

- Do not invoke LibreOffice, Microsoft Office, Python, Node.js, or .NET as an
  implicit runtime. LibreOffice is only an optional external CI oracle.
- Do not use `raw-set` when a typed operation exists. When raw XML replacement
  is unavoidable, inspect the exact part, preserve its root QName, write to a
  distinct output, and validate the result.
- Do not evaluate formulas through a shell or general-purpose script runtime.
  Native cell-formula writes validate and store the expression but do not
  compute a cached value implicitly. When fresh results are required, use
  `office native recalculate` or the typed
  `recalculate-spreadsheet-formulas` mutation. The closed native function
  registry must reject unsupported functions instead of falling back to code
  execution.
- Treat dynamic-array spill children as read-only calculated output. Find and
  edit or remove the formula anchor whose `formulaRef` contains the child;
  recalculation, cache writes, spill cleanup, and every sibling mutation in
  the batch roll back together on failure.
- Treat external OOXML relationships as inert. Do not fetch linked resources
  while inspecting or rendering a document. Native hyperlink writes accept
  only absolute HTTP, HTTPS, or mailto URIs without embedded credentials.
- Preserve no-clobber behavior. Do not add `--force` or replace an existing
  destination unless the user explicitly authorizes replacement.
- Treat a zero-match replacement as an unchanged successful receipt, not as a
  claimed edit. Keep Spreadsheet scopes narrow when only selected cells should
  change; the engine protects shared-string aliases outside the scope.
- Merge Spreadsheet cells only through a normalized cell/range path. Use
  `--merge-cells false` only with the exact existing range reported by
  `mergeCell` query results; do not approximate a destructive unmerge sweep.
  Merges that overlap another merge or a Spreadsheet table must fail closed.
- Treat Spreadsheet data validation as worksheet structure, not formula
  execution. Query `dataValidation` first, use the returned stable path for
  update/remove, and keep every rule area disjoint. Use typed list, comparison,
  or custom rules; do not work around an overlap or formula/type error with raw
  XML. Validate and read back one covered cell after mutation.
- Treat Spreadsheet conditional formatting as ordered worksheet structure, not
  evaluated styling. Query `conditionalFormatting` first, use the returned
  `/Sheet/cf[N]` path, preserve rule priority and `stopIfTrue`, and use only the
  closed classic/data-bar/color-scale/icon-set fields. Do not mutate a node
  whose semantic readback reports `nativeMutable=false`, bypass a shared-range
  or unknown-content error with raw XML, or claim that a semantic preview proves
  Excel's rendered result.
- Treat Spreadsheet defined names as scoped workbook identities. Query
  `namedrange` first and use the returned `@name` plus `@scope` path for
  update/remove. Do not edit `_xlnm.*` or `Slicer_*` names, collide with a table
  name, add a formula-bar leading `=`, or use raw XML to bypass a typed
  identity/ref error. Defined-name mutations store the definition and request
  recalculation; supported names referenced by cell formulas are resolved only
  during an explicit native recalculation pass.
- Import CSV or TSV only through the bounded typed import. Use exactly one
  regular source file or `--stdin`, make `--format` explicit for stdin, and
  inspect the target worksheet, `/Sheet/autofilter`, and `/Sheet/freeze` before
  enabling `--header`: header mode intentionally replaces the worksheet filter
  range and canonical frozen pane in one transaction. Explicit empty fields
  clear existing cells; missing trailing fields in ragged rows do not. Treat
  inferred formulas as parsed but not implicitly recalculated; run the explicit
  native pass when cached values are required. Never bypass a malformed-quote,
  formula-syntax, range, type, or unknown-view error with `raw-set`.
- Treat Spreadsheet AutoFilters as typed worksheet or table structure. Query
  `autofilter` or `filtercolumn` first and inspect `nativeMutable`; use the
  stable `/Sheet/autofilter` or `/Sheet/table[N]` path for updates. Every
  `--filter` is one strict JSON object containing a unique zero-based `column`
  and closed `criteria`. Repeated values replace the complete criterion list;
  use `--clear-filters` explicitly to clear it. Do not use raw XML to flatten
  imported date groups, color/icon filters, extensions, or embedded sort state.
- Sort Spreadsheet records only with `office native sort`. Use `/Sheet` to
  auto-detect the used range or an explicit `/Sheet/A1:D100`; supply ordered
  absolute `--key` columns and explicit header/case flags. Equal keys are
  stable, numbers precede text, and blanks stay last. A partial-column range
  moves only those cells, not whole rows. Verify `/Sheet/sort` and its key
  children after sorting. Removing `/Sheet/sort` clears metadata only and does
  not restore the old physical order. Do not bypass failures for formulas,
  totals rows, intersecting merges, pivots, unknown sort state, partial
  table/AutoFilter overlap, or non-lossless drawings with raw XML.
- Treat Spreadsheet ListObject tables as owned worksheet structures. Query
  `table` first and use the returned `/Sheet/table[N]` path for set/remove. The
  final range includes enabled header and totals rows; provide exactly one
  unique `--table-column` per range column and leave at least one data row. Do
  not overlap another table, a merge, or a worksheet AutoFilter, and do not use
  raw XML to bypass `nativeMutable=false` or an unknown-content/relationship
  error. Table criteria use the same typed filter-column values as worksheet
  AutoFilters. Table set automatically rewrites common explicit structured
  references and provably owned table-local column references when aliases or
  position-mapped columns change. Do not bypass
  `use.office.spreadsheet_table_formula_rewrite_unsupported` for unsafe local
  geometry or ownership, or `use.office.spreadsheet_table_referenced` when
  removal is blocked. Exact mutable table or data ranges without totals rows
  can use the separate physical sort contract; unsupported embedded/imported
  sort state remains non-mutable.
- Keep the default OfficeCLI compatibility route separate from the native
  engine. Do not depend on OfficeCLI's private resident protocol.

## Native Boundaries

The native engine currently owns safe OPC/ZIP admission, semantic reads,
bounded annotated and issue analysis, common typed mutations, atomic batches,
scoped literal/regex replacement, typed bold, italic, underline, vertical
script, font, size, RGB, and alignment formatting, Word/Spreadsheet single
strikethrough, Word double strikethrough, and Word/Presentation display case,
portable highlight, and primary-language formatting. It also owns typed inert
hyperlinks, typed legacy comments, typed Spreadsheet number formats, solid or
cleared fills, cardinal and diagonal borders, vertical alignment, wrapping,
rotation, indentation, shrink-to-fit, reading order, and exact merged-cell
editing. It owns typed Spreadsheet list, whole, decimal, date, time,
text-length, and custom data-validation rules over disjoint A1 ranges, with
stable rule paths and sparse cell readback. It also owns typed Spreadsheet
comparison/formula/text/statistical/date conditional-format rules, data bars,
two/three-color scales, standard 3/4/5-icon sets, differential fill/font/bold,
typed thresholds, stable paths, semantic queries, and exact canonical replay.
It owns workbook-global and
worksheet-local Spreadsheet defined names with stable scoped paths, typed
add/set/remove, semantic readback, and exact replay. It owns typed Spreadsheet
formula parsing, bounded dependency graphs, a closed typed function registry,
read-only calculation, atomic cached-value and dynamic-array spill writeback,
CLI/MCP/batch recalculation, and exact replay. It owns typed Spreadsheet
CSV/TSV import with bounded strict parsing, typed cell inference, explicit
empty-cell semantics, optional header AutoFilter/frozen-pane setup, semantic
`/Sheet/freeze` state, and exact canonical replay. It owns typed Spreadsheet
worksheet and table AutoFilters with closed value, comparison, top/bottom, and
dynamic criteria, stable filter paths, add/set/remove, and exact replay. It
owns stable ordered multi-key Spreadsheet physical sorting over an explicit or
auto-detected used range, persisted `/Sheet/sort` and `/Sheet/sort/key[N]`
state, record-bound metadata movement, metadata-only removal, and exact replay.
It owns typed Spreadsheet ListObject names, ranges, column identities,
header/totals state, filter criteria, built-in styles, stable table/column
paths, add/set/remove, and exact replay. It also
owns template merge, constrained XML access,
deterministic all-format HTML/SVG, Browser-injected semantic screenshots, and
authenticated loopback live watch for saved files.
Hyperlinks cover Word body/header/footer paragraphs and bookmarks, Spreadsheet
cells or bounded ranges and internal locations, and external Presentation shape
clicks or internal jumps to existing slides. Remaining boundaries include
modern threaded comments, replies/resolution, writable comment dates,
rich comment bodies, Word header/footer comment anchors,
gradient/pattern/theme fills, advanced x14 conditional-format visuals, named
styles, complete Excel function/structured-reference/external-workbook formula
compatibility, formula-bearing or table-totals sorting, table calculated
columns/totals functions, date-group/color/icon filters and unsupported
embedded/imported sort-state variants, custom table styles, query
tables/external data, advanced charts, pivots, and media,
interactive preview editing/annotations, and full Office layout fidelity. Fail
closed or use the explicit compatibility route rather than inventing
unsupported native behavior.
