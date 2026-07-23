# Spreadsheet Workflows

Use stable worksheet and A1 paths such as `/Sheet1`, `/Sheet1/A1`, and
`/Sheet1/A1:C20`. Preserve value types instead of writing every value as text.

## Contents

- [Inspect](#inspect)
- [Values and Formulas](#values-and-formulas)
- [Delimited Import and Frozen Panes](#delimited-import-and-frozen-panes)
- [Cell Text Formatting](#cell-text-formatting)
- [Cell Presentation Formatting](#cell-presentation-formatting)
- [Merged Cells](#merged-cells)
- [AutoFilters](#autofilters)
- [Sorting](#sorting)
- [Tables](#tables)
- [Data Validation](#data-validation)
- [Conditional Formatting](#conditional-formatting)
- [Named Ranges](#named-ranges)
- [Structure](#structure)
- [Verify](#verify)

## Inspect

```bash
a3s use office native get workbook.xlsx /Sheet1 --depth 2 --json
a3s use office native query workbook.xlsx 'cell[formula]' --json
a3s use office native view workbook.xlsx stats --json
a3s use office native view workbook.xlsx annotated --limit 200 --json
a3s use office native view workbook.xlsx issues --type content --json
```

## Values and Formulas

```bash
a3s use office native set workbook.xlsx /Sheet1/A1 --text 'Revenue' --json
a3s use office native set workbook.xlsx /Sheet1/A1:C20 --find Draft --replace Final --json
a3s use office native set workbook.xlsx /Sheet1/B1 --number 42.5 --json
a3s use office native set workbook.xlsx /Sheet1/C1 --boolean true --json
a3s use office native set workbook.xlsx /Sheet1/D1 --formula 'SUM(B1:B12)' --json
a3s use office native recalculate workbook.xlsx --output calculated.xlsx --json
a3s use office native set workbook.xlsx /Sheet1/E1 --url https://example.com/data --display Data --tooltip 'Open data' --json
a3s use office native set workbook.xlsx /Sheet1/F1 --location 'Sheet1!B2' --display B2 --json
a3s use office native set workbook.xlsx /Sheet1/G2:H4 --url https://example.com/range --display Range --json
a3s use office native query workbook.xlsx hyperlink --json
a3s use office native remove workbook.xlsx /Sheet1/E1/hyperlink --json
a3s use office native add workbook.xlsx /Sheet1/B2 --type comment --author Alice --text 'Check this formula' --json
a3s use office native set workbook.xlsx /Sheet1/B2/comment --author Bob --text 'Formula checked' --json
a3s use office native query workbook.xlsx comment --json
a3s use office native remove workbook.xlsx /Sheet1/B2/comment --json
```

General replacement accepts `/`, one worksheet, one cell, or a rectangular A1
range and edits string cells only. Literal matching is the default; add
`--regex` for Rust regex and capture expansion. A scoped edit of a shared rich
string clones and redirects only selected cells when references exist outside
the scope. Rich runs and unknown XML survive, and phonetic text is excluded.
Numeric, boolean, formula, and error values are not coerced. Zero matches are
reported as an unchanged success.

Formula writes remove one optional leading `=`, parse the body with bounded
Excel operator/reference syntax, store the normalized formula text, invalidate
stale calculation caches, and request recalculation. They do not calculate
implicitly. A syntax failure returns
`use.office.spreadsheet_formula_invalid` with byte and character offsets and
leaves the document unchanged.

Run `office native recalculate` in place or with `--output` to build the
dependency graph, calculate supported formulas, and atomically write typed
cached values and dynamic-array spills. The same operation is available as the
`recalculate-spreadsheet-formulas` batch/MCP mutation and as read-only or
writeback Rust APIs. Supported functions are `SUM`, `AVERAGE`, `MIN`, `MAX`,
`COUNT`, `COUNTA`, `ABS`, `SQRT`, `POWER`, `MOD`, `ROUND`, `IF`, `IFERROR`,
`AND`, `OR`, `NOT`, `CONCAT`, `CONCATENATE`, `ROW`, `COLUMN`, `SEQUENCE`,
`TRANSPOSE`, `PI`, and `NA`. Cross-sheet ranges, scoped names, typed errors,
array broadcasting, spill references, and ordinary Excel operators are
supported. ListObject structured references resolve a table `name` or
`displayName`: `Sales[Qty]` selects one data column,
`Sales[[Qty]:[Price]]` selects a contiguous data-column range, and `#All`,
`#Data`, `#Headers`, or `#Totals` selects structural rows. `Sales[@Qty]`,
`Sales[[#This Row],[Qty]]`, and table-local `[@Qty]` select the current data
row; table-local forms require the formula cell to be inside the inferred
table.

Spill children are read-only; update or remove the anchor instead. A blocked
spill produces typed `#SPILL!`, while formula error values such as `#DIV/0!`
remain typed cell results. Circular dependencies, unsupported or qualified
functions, missing tables, columns, or requested header/totals rows, disjoint
or non-canonical structured-reference forms, and external-workbook reads fail
with stable errors and leave the complete mutation batch unchanged. No shell,
script runtime, or external workbook is invoked. Limits are 8,192 formula
characters, depth 128 across both AST and nested named-reference resolution,
8,192 AST nodes, 100,000 reference areas per value, 100,000 graph formulas,
1,000,000 dependency edges, 1,000,000 graph reference visits, 100,000
materialized cells per array or function call, 100,000 cumulative spill
children per pass, 200,000 OOXML cell writes, and 1 MiB per text result. All
formula text results together are limited to 8 MiB per pass. Check
`formula_not_evaluated` and `formula_eval_error` issue records after the pass.

Semantic cell reads expose string-valued `formulaCached` on formula anchors and
`valuePresent` on every cell. A recalculated anchor reports
`formulaCached=true`; a formula stored without `<v>` reports `false`. Spill
children contain cached values but no independent `formula` field.

Exact replay accepts canonical formula storage and canonical array anchors only
when the array result is natively cached. It fails closed with
`use.office.dump_unsupported` for non-reproducible physical storage such as
explicit `t="normal"` formulas and uncached or malformed array anchors.

Hyperlinks target one cell or a bounded rectangular range. A missing single
cell is auto-created; a range link neither creates cells nor rewrites their
contents. External targets accept only absolute HTTP, HTTPS, or mailto URIs
without credentials; internal targets are workbook locations such as
`Sheet1!B2`. Display text and tooltips are optional. Single-cell links use a
cell `/hyperlink` child path; range links use a stable worksheet
`/hyperlink[N]` path returned by the mutation or a query. Update through the
cell/range or returned hyperlink path, and remove through the hyperlink path.
Overlapping hyperlink ranges fail with `use.office.hyperlink_range_conflict`.
Reads and previews never fetch external targets.

Classic cell comments use stable `/SheetName/A1/comment` paths and may be added
to an otherwise blank cell. Add requires an author and plain text. Update the
author or text through the comment path; Spreadsheet rejects separate initials
and slide coordinates instead of ignoring them. Native removal also cleans up
the matching VML note shape and removes unused comment/VML parts. Threaded
comments, replies, writable dates, and rich bodies are not yet native.

## Delimited Import and Frozen Panes

Import one bounded UTF-8 CSV or TSV source into an existing worksheet:

```bash
# .tsv and .tab infer TSV; every other file extension defaults to CSV.
a3s use office native import workbook.xlsx /Sheet1 source.csv \
  --header \
  --start-cell B2 \
  --json

# Stdin is bounded too. State the format instead of relying on its CSV default.
a3s use office native import workbook.xlsx /Sheet1 \
  --stdin \
  --format tsv \
  --output imported.xlsx \
  --json
```

Supply exactly one positional source, `--file <source>`, or `--stdin`. Files
must be regular, non-symlink files. One request accepts at most 8 MiB and a
100,000-cell rectangular target within Excel's row and column bounds. The
parser accepts a leading UTF-8 BOM, CRLF, quoted delimiters, embedded quoted
newlines, and doubled quotes. It rejects unclosed quotes, quotes inside
unquoted fields, and non-boundary content after a closing quote rather than
guessing.

An explicit empty field clears an existing target cell value while retaining
its unrelated style and extension content. A missing trailing field in a
ragged source row leaves that target cell unchanged, and a blank target is not
materialized just to represent emptiness. Import infers leading-`=` formulas,
finite numbers, booleans, ISO dates/times, and otherwise text. Dates honor the
workbook's 1900/1904 date system and receive the canonical native date number
format. Inferred formulas pass the same bounded syntax parser as direct cell
writes, are stored, and are marked for recalculation. Import does not calculate
them implicitly; run `office native recalculate` when fresh cached values are
required.

`--header` treats the first imported row as headers. In the same atomic
transaction it adds or replaces the worksheet AutoFilter over the imported
extent and adds or replaces one canonical frozen pane below the header. Inspect
existing `/Sheet1/autofilter` and `/Sheet1/freeze` state first when importing
into a populated worksheet.

Read or remove the frozen pane through its stable semantic path:

```bash
a3s use office native get workbook.xlsx /Sheet1/freeze --json
a3s use office native query workbook.xlsx frozen-pane --json
a3s use office native remove workbook.xlsx /Sheet1/freeze --json
```

Rust, versioned batch, and standard MCP can set a canonical pane independently:

```json
{
  "operation": "set-spreadsheet-frozen-pane",
  "sheet": "/Sheet1",
  "pane": {
    "frozenRows": 1,
    "frozenColumns": 0,
    "topLeftCell": "A2"
  }
}
```

`topLeftCell` must be below and to the right of every frozen split. Imported
split panes, vendor attributes, unknown children, or unsupported view state
remain readable with `nativeMutable=false` and fail closed on set/remove.
Strict and transitional SpreadsheetML are preserved.

## Cell Text Formatting

```bash
a3s use office native set workbook.xlsx /Sheet1/A1:C1 --bold true --underline double --script superscript --strikethrough true --font-family Aptos --font-size 11.5 --text-color 0066CC --align center --json
a3s use office native set workbook.xlsx /Sheet1/A2 --text 'Total' --bold true --align right --json
```

Formatting accepts one cell or a bounded rectangular range and may be combined
with one content write. It auto-creates empty styled cells and deduplicates
OOXML font and cell-style records. The native typed subset covers bold, italic,
`none`/single/double underline, baseline/superscript/subscript text, explicit
single strikethrough, font family, point size, RGB text color, and horizontal
alignment. Run-only text case, highlight, language, and double strikethrough
fail atomically with `use.office.spreadsheet_run_format_unsupported`; they are
not silently flattened into a cell style. Use the separate cell-presentation
options below for non-text properties and the typed conditional-format contract
for rule-based differential styling. Named styles and gradient/pattern/theme
fills remain outside the native subset.

## Cell Presentation Formatting

```bash
a3s use office native set workbook.xlsx /Sheet1/A1:C3 --number-format currency --fill FFF2CC --border-all thin --border-color 808080 --border-bottom double --border-bottom-color 000000 --vertical-align center --wrap-text true --json
a3s use office native set workbook.xlsx /Sheet1/D1 --number 0.125 --bold true --number-format percent --fill 0066CC --text-rotation 45 --indent 1 --shrink-to-fit false --reading-order rtl --json
a3s use office native set workbook.xlsx /Sheet1/E1 --border-diagonal slant-dash-dot --border-diagonal-color FF0000 --border-diagonal-up true --border-diagonal-down false --json
a3s use office native set workbook.xlsx /Sheet1/A1:C3 --fill none --wrap-text false --reading-order context --json
```

Cell presentation accepts one cell or a bounded rectangular range and may be
combined atomically with one content write, text formatting, a hyperlink, and
merged-cell state.
Use `--number-format` for an explicit Excel format code or one of `general`,
`number`, `currency`, `accounting`, `percent`, `scientific`, `text`, `date`,
`time`, or `datetime`. Codes may contain at most four sections and must keep
quotes and square brackets balanced.

`--fill` accepts `none` or exactly six hexadecimal RGB digits.
`--border-all` (alias `--border`) sets all four cardinal line styles;
`--border-color` supplies their default RGB color. Override one side with
`--border-left`, `--border-right`, `--border-top`, or `--border-bottom` and the
matching `-color` option. Use `none` to clear a line. Styles are `thin`,
`medium`, `thick`, `double`, `dashed`, `dotted`, `dash-dot`, `dash-dot-dot`,
`hair`, `medium-dashed`, `medium-dash-dot`, `medium-dash-dot-dot`, and
`slant-dash-dot`. The shared diagonal line uses `--border-diagonal` and
`--border-diagonal-color`; explicitly select its direction with
`--border-diagonal-up` and `--border-diagonal-down`. A color option requires a
non-`none` style in the same command.
`--vertical-align` accepts `top`, `center`, `bottom`, `justify`, or
`distributed`. `--wrap-text` and `--shrink-to-fit` require an explicit boolean;
`--text-rotation` accepts 0–180 or 255 for stacked text; `--indent` accepts
0–255; and `--reading-order` accepts `context`, `ltr`, or `rtl`. Explicit
`none`, `false`, zero, and `context` values clear or reset the corresponding
property instead of being treated as omitted.

The writer preserves unknown style and border data and deduplicates
number-format, fill, border, and cell-style records. Invalid values, an empty
cell-format or border object, a bad
target kind, or any other mutation failure rolls back the complete in-memory
batch before save. Verify with a targeted `get`; HTML/SVG expose observed
values as inert `data-*` attributes but remain semantic previews rather than
Excel layout evidence.

## Merged Cells

```bash
# Merge, optionally composing content and formatting in the same atomic set.
a3s use office native set workbook.xlsx /Sheet1/A1:C1 --text 'Quarter' --bold true --merge-cells true --json

# Inspect the anchor, a blank covered cell, the exact range, and stable merges.
a3s use office native get workbook.xlsx /Sheet1/A1 --json
a3s use office native get workbook.xlsx /Sheet1/B1 --json
a3s use office native get workbook.xlsx /Sheet1/A1:C1 --json
a3s use office native query workbook.xlsx mergeCell --json

# Unmerge only the exact existing range.
a3s use office native set workbook.xlsx /Sheet1/A1:C1 --merge-cells false --json
```

Paths are case-insensitive and their A1 endpoints are normalized. Repeating an
exact merge is an unchanged success. A non-identical overlapping merge fails
with `use.office.spreadsheet_merge_overlap`; any range intersecting a
ListObject table fails with `use.office.spreadsheet_merge_table_overlap`.
Unmerge is deliberately precise: an absent exact range is unchanged, while a
range that intersects but does not exactly equal an existing merge fails with
`use.office.spreadsheet_merge_not_exact` and reports `validRanges`. Unmerge
those returned ranges individually; there is no destructive sweep operation.

Semantic query results expose `/SheetName/mergeCell[N]` with a normalized
`ref`. Covered cell reads expose the normalized range in `format.merge` and
whether the cell is the top-left anchor in `format.mergeAnchor`. A blank
covered cell is returned virtually without creating every covered cell in
`sheetData`. An exact range read reports `format.merge=true`; an unmerged range
reports `false`. HTML/SVG use inert `data-merge` and `data-merge-anchor`
attributes only and do not claim Excel layout fidelity.

Merge/unmerge participates in normal batch rollback, strict/transitional OOXML
preservation, and replay dump. Unknown merge collection attributes and
extension children are retained. If removing the final merge would also delete
unknown collection data, the operation fails with
`use.office.spreadsheet_merge_unknown_content` instead of discarding it. This
is merged-cell structure support, not complete rich Spreadsheet or OfficeCLI
parity.

## AutoFilters

Use the typed worksheet lifecycle instead of editing `<autoFilter>` directly.
Every repeated `--filter` value is one strict JSON object with a unique
zero-based column offset and one closed criterion:

```bash
a3s use office native add workbook.xlsx /Sheet1 \
  --type auto-filter \
  --range A1:C20 \
  --filter '{"column":0,"criteria":{"type":"values","values":["Open","Closed"],"includeBlanks":true}}' \
  --filter '{"column":2,"criteria":{"type":"greater-than","value":"100"}}' \
  --json

a3s use office native get workbook.xlsx /Sheet1/autofilter --depth 2 --json
a3s use office native query workbook.xlsx \
  'filtercolumn[criteriaType=greater-than]' --json

# Omitted range is preserved; repeated filters replace the complete list.
a3s use office native set workbook.xlsx /Sheet1/autofilter \
  --range B2:D30 \
  --filter '{"column":1,"criteria":{"type":"dynamic","kind":"this-month"}}' \
  --json

# Clearing criteria is explicit and preserves the worksheet AutoFilter range.
a3s use office native set workbook.xlsx /Sheet1/autofilter \
  --clear-filters --json
a3s use office native remove workbook.xlsx /Sheet1/autofilter --json
```

Criteria types are `values` with optional `includeBlanks`; `equals`,
`not-equals`, `contains`, `does-not-contain`, `begins-with`, `ends-with`,
`greater-than`, `greater-than-or-equal`, `less-than`, and
`less-than-or-equal`; `between`/`not-between` with `lower` and `upper`;
`blanks`/`non-blanks`; `top`/`bottom` with `count` 1–500;
`top-percent`/`bottom-percent` with `percent` 1–100; and `dynamic` with a closed
average/date/month/quarter kind such as `above-average`, `today`,
`this-month`, `year-to-date`, `quarter1`, or `month12`. Values are XML-safe,
bounded, and wildcard literals are escaped rather than interpreted as an
unrequested pattern.

Batch, standard MCP, and Rust use one complete value:

```json
{
  "operation": "add-spreadsheet-auto-filter",
  "sheet": "/Sheet1",
  "filter": {
    "range": "A1:C20",
    "columns": [
      {
        "column": 0,
        "criteria": {
          "type": "values",
          "values": ["Open", "Closed"],
          "includeBlanks": true
        }
      },
      {
        "column": 2,
        "criteria": {"type": "between", "lower": "10", "upper": "100"}
      }
    ]
  }
}
```

Use `set-spreadsheet-auto-filter` with a complete `filter` value for batch or
MCP replacement. A worksheet accepts one AutoFilter; its range may not
intersect a table or merged range. Semantic nodes expose
`/SheetName/autofilter`, `/SheetName/autofilter/filterColumn[N]`, value
children, normalized `ref`, criterion metadata, and `nativeMutable`. Table
AutoFilters use `/SheetName/table[N]/autofilter` and the identical column
value. Do not mutate a node with `nativeMutable=false`. Imported date-group,
color/icon, extension, unknown-content, and sort-state forms fail closed.
Filter definitions and physical row sorting are separate typed capabilities;
the sort workflow below does not flatten an unsupported imported AutoFilter.

## Sorting

Use the dedicated command to physically reorder Spreadsheet records by one or
more ordered keys:

```bash
a3s use office native sort workbook.xlsx /Sheet1/A1:D100 \
  --key B:desc \
  --key C:asc \
  --header true \
  --case-sensitive false \
  --json

# A worksheet path auto-detects the smallest used cell rectangle.
a3s use office native sort workbook.xlsx /Sheet1 \
  --key B:ascending \
  --header true \
  --case-sensitive false \
  --output sorted.xlsx \
  --json
```

Each repeated key is an absolute column from `A` through `XFD`, optionally
followed by `asc`/`ascending` or `desc`/`descending`. Direction defaults to
ascending. Keys must be unique and lie inside the selected range; the first key
has highest precedence. One request accepts 1–64 keys and at most 100,000
selected cells. `header` defaults to false and keeps exactly the first selected
row fixed when true. `caseSensitive` defaults to false.

Sorting is stable. Numeric and boolean cell values compare numerically before
text, blanks stay last for both ascending and descending keys, and records equal
on all keys retain their source order. Case-insensitive comparison treats text
that differs only by case as equal, so stable source order decides the result.
Sorting an explicit partial-column range moves only the selected cell fragments;
cells outside those columns and destination row properties remain fixed. Sparse
records can create sparse destination rows.

Batch, standard MCP, replay, and Rust use the same complete value:

```json
{
  "operation": "sort-spreadsheet-range",
  "path": "/Sheet1/A1:D100",
  "sort": {
    "keys": [
      {"column": "B", "direction": "descending"},
      {"column": "C", "direction": "ascending"}
    ],
    "header": true,
    "caseSensitive": false
  }
}
```

The physical permutation and metadata changes are one atomic mutation. Exact
mutable ListObject and worksheet-AutoFilter ranges are supported when `header`
matches their structure; their exact data ranges are also supported with
`header=false`. Partial intersections fail closed. A table totals row, formulas
anywhere in the workbook, an intersecting merge, an owned pivot table, unknown
existing sort state, or a drawing that cannot follow one record losslessly
rejects the operation before save.

Hyperlinks, classic/threaded comment refs, VML note anchors, data-validation
areas, conditional-format areas, protected ranges, ignored errors, and
supported drawing anchors follow the records. Chart caches are cleared. Strict
and transitional SpreadsheetML, prefixed XML, sparse gaps, and destination row
properties are preserved; the worksheet used dimension is recomputed.

Successful sorting persists a worksheet-level sort state. Inspect it and its
ordered children through stable semantic paths:

```bash
a3s use office native get workbook.xlsx /Sheet1/sort --depth 1 --json
a3s use office native query workbook.xlsx 'sortkey[direction=descending]' --json

# This removes metadata only. It never restores the previous physical order.
a3s use office native remove workbook.xlsx /Sheet1/sort --json
```

Only simple supported row-sort state is mutable. Imported column sorting,
extended conditions, unknown attributes or children, inconsistent condition
ranges, and unsupported embedded table/AutoFilter state remain readable with
`nativeMutable=false`; do not bypass that boundary with `raw-set`.

## Tables

Use the closed ListObject lifecycle for an Excel table rather than editing its
worksheet relationship or `xl/tables` part directly:

```bash
# The final range includes the header and any enabled totals row.
a3s use office native add workbook.xlsx /Sheet1 \
  --type table \
  --name Sales \
  --range F1:H4 \
  --table-column Name \
  --table-column Qty \
  --table-column Price \
  --filter '{"column":1,"criteria":{"type":"top","count":10}}' \
  --style medium:4 \
  --show-row-stripes true \
  --json

# Discover the stable table and exact column identities.
a3s use office native query workbook.xlsx 'table[name=Sales]' --json
a3s use office native get workbook.xlsx '/Sheet1/table[1]' --depth 1 --json

# CLI set preserves omitted fields. Repeated columns replace the complete list.
a3s use office native set workbook.xlsx '/Sheet1/table[1]' \
  --name Inventory \
  --display-name InventoryView \
  --range B2:D6 \
  --table-column Item \
  --table-column Units \
  --table-column Cost \
  --totals-row true \
  --style dark:2 \
  --show-row-stripes false \
  --show-column-stripes true \
  --json

# Replace all table criteria, or clear them explicitly.
a3s use office native set workbook.xlsx '/Sheet1/table[1]' \
  --filter '{"column":0,"criteria":{"type":"contains","value":"West"}}' \
  --json
a3s use office native set workbook.xlsx '/Sheet1/table[1]' \
  --clear-filters --json

a3s use office native remove workbook.xlsx '/Sheet1/table[1]' --json
```

Table `set` rewrites common structured references when `name`, effective
`displayName`, or position-mapped column names change. The audit covers cell
formulas, workbook defined names, conditional-format and data-validation
formulas, charts, and formula carriers in table parts. String literals and
external-workbook references remain unchanged. Table-local forms such as
`[@Qty]` are rewritten only with provable ListObject ownership; an unsafe local
rewrite or local reference across a range/header/totals-row change fails with
`use.office.spreadsheet_table_formula_rewrite_unsupported`. Removing a table
still targeted by a structured reference fails with
`use.office.spreadsheet_table_referenced`. Both failures roll back the complete
mutation.

Provide exactly one non-empty, case-insensitively unique column name for every
range column. Table `name` and optional `displayName` use Excel identifier
grammar, are limited to 255 characters, may not resemble A1/R1C1 references,
and share a case-insensitive workbook namespace with other table and defined
names. Use `--display-name none` to clear a distinct display name during CLI
set. A workbook accepts at most 65,536 tables.

`--header-row` and `--totals-row` require booleans. The final range must retain
at least one data row after enabled structural rows. An enabled header stamps
the ordered column names into its cells and creates a table-owned AutoFilter;
an enabled totals row is excluded from that filter range. The table range may
not intersect another table, a merged range, or a worksheet-level AutoFilter.
Every failure rolls back the full in-memory batch.

`--style` accepts `none`, `light:<1-21>`, `medium:<1-28>`, or `dark:<1-11>`.
New tables default to `medium:2` with row stripes. `--show-first-column`,
`--show-last-column`, `--show-row-stripes`, and `--show-column-stripes` require
booleans. `none` omits the style carrier and therefore requires every display
flag to be false.

Batch, standard MCP, and Rust use the complete typed value:

```json
{
  "operation": "add-spreadsheet-table",
  "sheet": "/Sheet1",
  "table": {
    "name": "Sales",
    "range": "F1:H4",
    "columns": [
      {"name": "Name"},
      {"name": "Qty"},
      {"name": "Price"}
    ],
    "filters": [
      {"column": 1, "criteria": {"type": "top", "count": 10}}
    ],
    "headerRow": true,
    "totalsRow": false,
    "style": {"family": "medium", "number": 4},
    "showFirstColumn": false,
    "showLastColumn": false,
    "showRowStripes": true,
    "showColumnStripes": false
  }
}
```

Use `set-spreadsheet-table` with a stable `path` and a complete `table` value
for batch or MCP replacement; use ordinary `remove` for deletion. Semantic
nodes expose `/SheetName/table[N]`, children such as
`/SheetName/table[N]/column[M]`, normalized `ref`, style and structural flags,
the table-owned AutoFilter and criteria, and `nativeMutable`. Do not set a table
whose readback reports `nativeMutable=false`.

The writer owns the table part, content type, worksheet relationship,
`tableParts`, header cells, and table AutoFilter while preserving strict or
transitional OOXML and supported unknown root/style/extension content. It fails
closed for unknown column metadata, unsafe relationship graphs, or collection
data that cannot be retained. Exact replay, CLI atomic batches, and standard MCP
sessions use the same contract. Calculated columns, totals labels/functions,
date-group/color/icon filters, unsupported embedded/imported sort state, custom
table styles, query tables, external data, slicers, and pivot integration are
not yet native. Exact mutable table or data ranges without totals rows can be
physically sorted through the separate sort contract.

## Data Validation

Use one typed rule for one or more disjoint areas on the same worksheet:

```bash
# Add a list rule. Inline comma-separated values are quoted by the writer.
a3s use office native add workbook.xlsx /Sheet1 \
  --type data-validation \
  --validation-type list \
  --range A2:A20 \
  --range C2:C20 \
  --formula1 'Draft,Review,Approved' \
  --prompt-title Status \
  --prompt 'Choose a workflow state' \
  --error-title 'Invalid status' \
  --error-message 'Choose one of the listed states' \
  --error-style stop \
  --json

# Discover the stable rule and read an otherwise blank validated cell.
a3s use office native query workbook.xlsx 'dataValidation[type=list]' --json
a3s use office native get workbook.xlsx '/Sheet1/dataValidation[1]' --json
a3s use office native get workbook.xlsx /Sheet1/C3 --json

# Replace the rule through its stable path. Omitted CLI fields are preserved.
a3s use office native set workbook.xlsx '/Sheet1/dataValidation[1]' \
  --validation-type whole \
  --range B2:B50 \
  --operator between \
  --formula1 18 \
  --formula2 120 \
  --allow-blank false \
  --error-style warning \
  --json

# Remove exactly one discovered rule.
a3s use office native remove workbook.xlsx '/Sheet1/dataValidation[1]' --json
```

Rule types are `list`, `whole`, `decimal`, `date`, `time`, `text-length`, and
`custom`. List and custom rules do not accept an operator or `formula2`.
Comparison rules require `between`, `not-between`, `equal`, `not-equal`,
`greater-than`, `greater-than-or-equal`, `less-than`, or
`less-than-or-equal`; the first two require `formula2` and the others reject
it. `--error-style` accepts `stop`, `warning`, or `information`. Blank,
input-message, error-message, and list-dropdown flags default to true for a new
A3S rule. Only list rules accept `--in-cell-dropdown false`.

For a list, `formula1` may be inline comma-separated text, an A1 source such as
`=$H$2:$H$5`, or a defined name. Embedded double quotes in an inline list are
rejected; use cells as the list source instead. Date inputs in valid
`YYYY-MM-DD` form from 1900 through 9999 become serial dates under the
workbook's declared 1900 or 1904 date system. Time inputs in `HH:MM` or
`HH:MM:SS` form become day fractions. Range, defined-name, dynamic spill, and
function sources such as `INDIRECT(...)` remain formulas. Other formula text is
stored after removing one optional leading `=`; data-validation rule predicates
are not executed by either the validation writer or the cell-formula
recalculation pass.

Each rule accepts 1–1,024 normalized rectangular A1 areas and a worksheet
accepts at most 65,534 rules. Formula fields are limited to 255 characters;
titles to 32, prompt text to 255, and error text to 225. Areas must not overlap
inside one rule or across rules. An invalid range, operator/formula mismatch,
XML-forbidden character, or overlap rolls back the whole batch. Use `none` or
an empty CLI value to clear optional text or `formula2`; use `--operator none`
to clear an operator when changing to a list or custom rule.

The native batch and standard MCP payload is the same closed value:

```json
{
  "operation": "add-data-validation",
  "sheet": "/Sheet1",
  "validation": {
    "type": "date",
    "ranges": ["D2:D100"],
    "operator": "between",
    "formula1": "2026-01-01",
    "formula2": "2026-12-31",
    "allowBlank": false,
    "showInput": true,
    "showError": true,
    "errorStyle": "stop",
    "inCellDropdown": true
  }
}
```

Use `set-data-validation` with `path` and a complete `validation` value for a
typed batch replacement. Use ordinary `remove` for deletion. A rule appears as
`/SheetName/dataValidation[N]` with normalized `ref`, type, operator, formulas,
messages, and flags. Covered cells expose `dataValidation` and
`validationType`; requesting a covered blank cell returns a virtual cell and
does not populate the worksheet. HTML/SVG emit only inert `data-validation`
metadata.

Updates preserve unknown rule attributes. Replacement fails if unknown rule
children would be lost, and final removal fails if unknown collection data
would be discarded. Strict/transitional OOXML, atomic batch rollback, and
exact replay are supported. This capability does not add table calculated
columns/totals functions, date-group/color/icon filters, unsupported imported
sort-state variants, charts, pivots, data-validation predicate execution, or
Excel layout fidelity.

## Conditional Formatting

Use a closed rule family rather than a generic style or formula property map:

```bash
# Highlight values above 80 with a differential fill and font style.
a3s use office native add workbook.xlsx /Sheet1 \
  --type conditional-format \
  --rule-type cell-is \
  --range A2:A20 \
  --operator greater-than \
  --formula1 80 \
  --fill C6EFCE \
  --text-color 006100 \
  --bold true \
  --stop-if-true true \
  --json

# Add the three broad visual forms.
a3s use office native add workbook.xlsx /Sheet1 \
  --type conditional-format --rule-type data-bar --range B2:B20 \
  --color 638EC6 --min min --max number:100 --show-value true --json
a3s use office native add workbook.xlsx /Sheet1 \
  --type conditional-format --rule-type color-scale --range C2:C20 \
  --min-color F8696B --midpoint percentile:50 --mid-color FFEB84 \
  --max-color 63BE7B --json
a3s use office native add workbook.xlsx /Sheet1 \
  --type conditional-format --rule-type icon-set --range D2:D20 \
  --icon-set 3-traffic-lights-1 --reverse true --show-value false --json

# Discover stable rules and partially update one through the CLI adapter.
a3s use office native query workbook.xlsx \
  'conditionalFormatting[type=iconSet]' --json
a3s use office native get workbook.xlsx '/Sheet1/cf[1]' --json
a3s use office native set workbook.xlsx '/Sheet1/cf[1]' \
  --formula1 90 --fill FFEB9C --stop-if-true false --json

# Remove exactly one rule.
a3s use office native remove workbook.xlsx '/Sheet1/cf[2]' --json
```

Classic `--rule-type` values are `cell-is`, `formula`, `contains-text`,
`not-contains-text`, `begins-with`, `ends-with`, `top`, `bottom`,
`above-average`, `below-average`, `duplicate-values`, `unique-values`,
`contains-blanks`, `not-contains-blanks`, `contains-errors`,
`not-contains-errors`, and `time-period`. Formula rules use `--formula` or
`--formula1`. Text predicates use `--text`. Top/bottom rules use `--rank`,
`--percent`, and `--bottom`; average rules use `--above`, `--equal-average`, and
`--std-dev`; time rules use `--period`. Classic rules accept only `--fill`,
`--text-color`, and `--bold` as differential formatting. Use `none` to clear a
differential color or bold state during a partial CLI update.

`cell-is` supports `between`, `not-between`, `equal`, `not-equal`,
`greater-than`, `greater-than-or-equal`, `less-than`, and
`less-than-or-equal`. Between and not-between require `--formula2`; the other
operators reject it. Formula bodies omit a leading `=`. A3S stores them and
does not evaluate whether a rule matches.

Visual rule types are `data-bar`, `color-scale`, and `icon-set`. Thresholds use
`min`, `max`, `number:<value>`, `percent:<0..100>`,
`percentile:<0..100>`, or `formula:<expression>`. A data bar accepts
`--color`, `--min`, `--max`, `--show-value`, and optional 0–100
`--min-length`/`--max-length`. A color scale becomes three-color only when both
`--midpoint` and `--mid-color` are present; set either to `none` to return to a
two-color scale. An icon set accepts a standard 3/4/5-icon name and one repeated
`--threshold` per icon. Omit thresholds to generate evenly spaced percent
defaults. Visual rules reject `--fill`, `--text-color`, and `--bold` rather than
ignoring them.

The native batch and standard MCP value is fully typed:

```json
{
  "operation": "add-conditional-format",
  "sheet": "/Sheet1",
  "conditionalFormat": {
    "ranges": ["A2:A20", "C2:C20"],
    "stopIfTrue": true,
    "rule": {
      "type": "cellIs",
      "operator": "greaterThan",
      "formula1": "80",
      "format": {
        "fill": {"red": 198, "green": 239, "blue": 206},
        "fontColor": {"red": 0, "green": 97, "blue": 0},
        "bold": true
      }
    }
  }
}
```

Use `set-conditional-format` with a stable `path` and a complete
`conditionalFormat` value for batch or MCP replacement; only the CLI adapter
merges omitted fields. Use ordinary `remove` for deletion. Each rule accepts
1–1,024 internally disjoint normalized A1 areas; a worksheet accepts at most
65,534 rules. Different rules may overlap and remain ordered by unique
priority. `stopIfTrue` preserves Excel's later-rule suppression semantics.

Semantic nodes use `/SheetName/cf[N]`, expose the normalized `ref`, priority,
rule-specific fields, and `nativeMutable`, and support
`conditionalFormatting[type=...]` selectors. Unsupported extension-only rules
remain readable with `nativeMutable=false`; never replace them through raw XML
to bypass a typed mutation failure. Unknown attributes and strict/transitional
SpreadsheetML are preserved. Unknown child or collection content that cannot
survive a set/remove operation fails closed. Imported multi-rule carriers share
one range: keep that range unchanged when updating one child rule.

Canonical replay, atomic rollback, CLI, and standard MCP are supported. This
conditional-format feature does not evaluate rule formulas or reproduce
Excel's rendered appearance, and it does not support x14-only negative data-bar
axes/colors, custom icon sets, table/chart/pivot formatting, or complete
OfficeCLI/Spreadsheet parity.

## Named Ranges

Use a scoped defined name when formulas or validation rules need a stable
workbook identifier:

```bash
# Workbook-global name: the A1 ref must include its worksheet.
a3s use office native add workbook.xlsx / \
  --type named-range \
  --name Revenue \
  --ref 'Sheet1!$A$2:$A$20' \
  --scope workbook \
  --comment 'Workbook revenue' \
  --json

# A worksheet parent defaults to local scope and qualifies a bare A1 ref.
a3s use office native add workbook.xlsx /Sheet1 \
  --type named-range \
  --name Status \
  --ref A2:A20 \
  --json

# Discover and address the complete scoped identity.
a3s use office native query workbook.xlsx 'namedrange[name=Revenue]' --json
a3s use office native get workbook.xlsx \
  '/namedrange[@name=Revenue][@scope=workbook]' --json

# CLI set preserves omitted fields; use `none` to clear the comment.
a3s use office native set workbook.xlsx \
  '/namedrange[@name=Status][@scope=Sheet1]' \
  --name WorkflowStatus \
  --ref B2:B20 \
  --volatile false \
  --json

a3s use office native remove workbook.xlsx \
  '/namedrange[@name=Revenue][@scope=workbook]' --json
```

Canonical paths carry both `@name` and `@scope`. Name-only paths such as
`/namedrange[Revenue]`, `@name` paths, and one-based positional paths remain
compatible, but an unscoped path fails with
`use.office.spreadsheet_named_range_ambiguous` when the same name exists in
multiple scopes. Use the canonical path returned by add/query for update and
remove. Path values are percent-encoded when a name or worksheet requires it.
If the worksheet is literally named `workbook`, use `--scope
worksheet:workbook`; semantic readback and its canonical path use the same
escaped scope so it cannot collide with the global identity.

Names may contain Unicode letters and digits plus underscore, period, and
backslash, must begin with a letter, underscore, or backslash, may not resemble
A1/R1C1 notation, and are limited to 255 characters. Refs omit the formula-bar
leading `=` and are limited to 8,192 characters. A workbook-scoped bare A1 ref
is invalid; a local bare A1 ref is qualified with its scope worksheet. Simple
qualified refs must target an existing sheet. Cross-workbook refs requiring an
external-link part are not native. Comments are limited to 255 characters.

The identity is case-insensitively unique by `(name, scope)`. A defined name
also may not collide with a ListObject table `name` or `displayName`. Do not
edit or remove `_xlnm.*` print/filter definitions or `Slicer_*` sentinels;
manage the owning typed feature instead. `--volatile true` maps to the OOXML
defined-name function flag and requests recalculation. The name mutation does
not itself calculate anything; supported names referenced by cell formulas are
resolved by an explicit native recalculation pass.

Batch, standard MCP, and Rust use one complete typed value for add/set and
ordinary typed `remove` for deletion. The writer preserves strict/transitional
SpreadsheetML and unknown attributes. Unknown collection or child content
fails closed when an edit cannot retain it. Exact replay includes supported
defined names. This remains defined-name lifecycle support, not external-link
authoring or complete Spreadsheet parity.

## Structure

```bash
a3s use office native add workbook.xlsx / --type sheet --name Data --json
a3s use office native insert-rows workbook.xlsx /Sheet1 2 --count 3 --json
a3s use office native delete-columns workbook.xlsx /Sheet1 C --count 1 --json
a3s use office native rename-sheet workbook.xlsx /Sheet1 Summary --json
a3s use office native copy-sheet workbook.xlsx /Summary 'Summary Copy' --json
a3s use office native move-sheet workbook.xlsx /Data 1 --json
a3s use office native add workbook.xlsx /Sheet1/A1 --type picture --input chart.png --alt 'Sales chart' --json
```

Supported structural edits rewrite bounded A1 references and related metadata.
Pivot-table changes, unsafe 3D references, x14-only conditional-format
extensions, full chart authoring, and complete Excel formula compatibility
remain outside the native subset and fail closed where safety cannot be proven.

## Verify

```bash
a3s use office native validate workbook.xlsx --json
a3s use office native view workbook.xlsx issues --limit 200 --json
a3s use office native view workbook.xlsx html --output workbook.html --json
a3s use office native view workbook.xlsx svg --output workbook.svg --json
a3s use office native watch workbook.xlsx --port 0
```

HTML, SVG, and screenshots are sparse semantic previews, not Excel layout or
print fidelity. Watch reloads saved revisions; it does not provide inline cell
editing or trigger formula recalculation.
