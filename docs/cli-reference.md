# A3S Office CLI reference

Office is moving to an A3S-owned Rust engine for Word, Spreadsheet, and
Presentation documents. The native engine now includes bounded package
admission, byte-preserving XML, content types, a safe relationship graph,
stable selectors, semantic `get`, `query`, `text`, `outline`, and `stats`
reads, bounded `issues` reports, safe blank-document creation, and
loss-preserving text assignment plus scoped literal/regex replacement for Word,
Spreadsheet, and Presentation. Matches may span rich-text runs without
flattening their formatting. Typed rich-text mutation covers bold, italic,
`none`/single/double underline, baseline/superscript/subscript, font family,
exact centipoint font size, RGB text color, and horizontal alignment. Word and
Spreadsheet also support an explicit single-strikethrough boolean;
Presentation rejects that property with a typed error instead of silently
dropping it. Word and Presentation additionally share a portable 17-color
highlight palette, display-only text case, and conservative BCP-47 primary
language tags. Word alone supports explicit double strikethrough; Spreadsheet
and Presentation reject it through format-specific typed errors.
Word and Presentation apply character properties to run paths and alignment to
paragraph paths; Spreadsheet applies the same contract to cells or bounded A1
ranges,
creating and deduplicating OOXML font and cell-style records when necessary. It
also exposes a separate typed Spreadsheet cell-presentation contract for
number formats, solid RGB fill or explicit fill removal, cardinal and diagonal
borders, vertical alignment, wrap text, rotation, indentation, shrink-to-fit,
and reading order. These properties apply to a cell or bounded rectangular
range, compose atomically with content and text formatting, preserve unknown
style data, and deduplicate number-format, fill, border, and cell-style
records. Spreadsheet merged cells use a separate typed contract: rectangular
A1 ranges are normalized, exact repeated merges are idempotent, and unmerge
requires one exact existing range. Geometric overlaps and ListObject table
intersections fail closed. Semantic reads expose stable `mergeCell` nodes and
anchor metadata without materializing every blank covered cell. A third typed
Spreadsheet contract owns data-validation rules for list, whole-number,
decimal, date, time, text-length, and custom-formula constraints. One rule may
target multiple disjoint normalized A1 ranges and carries typed comparison
operators, prompt/error messages, alert style, blank handling, and
list-dropdown state. Inline lists are quoted safely, ISO dates and clock times
are normalized to Spreadsheet serial values, and overlapping rule ranges fail
atomically. Semantic reads expose stable `/Sheet/dataValidation[N]` nodes,
annotate observed and virtual blank cells, and project inert metadata into
HTML/SVG. Native add/set/remove, batch, exact replay, CLI, Rust, and standard
MCP use the same closed contract while retaining strict/transitional
SpreadsheetML and unknown attributes. It stores validation formulas but does
not evaluate them.

A separate typed Spreadsheet conditional-format contract owns comparison and
formula rules; text, rank, average, duplicate/unique, blank/error, and time
predicates; data bars; two- or three-color scales; and standard three-, four-,
or five-icon sets. Classic rules apply a differential solid fill, font color,
and bold state. Visual rules use typed thresholds, colors, and display flags.
One rule can cover multiple disjoint A1 areas and can stop later rule
evaluation. Semantic reads expose stable `/Sheet/cf[N]` nodes and selectors
such as `conditionalFormatting[type=dataBar]`. Rust, versioned batch, CLI,
standard MCP, exact replay, and the Office Skill share the closed value. The
writer deduplicates differential formats, preserves strict or transitional
SpreadsheetML plus unknown attributes, and fails closed when unsupported child
or collection content would be lost. This milestone does not evaluate rule
formulas, reproduce Excel rendering, or cover x14-only data-bar axes/colors,
table/chart/pivot conditional formatting, or complete Spreadsheet parity.

Native Spreadsheet defined names use a separate typed contract for
workbook-global and worksheet-local scopes. Stable selectors include both name
and scope, while compatibility selectors remain available when the name is
unambiguous. `worksheet:workbook` disambiguates a local scope when the
worksheet itself is literally named `workbook`. The writer validates Excel identifier and reference limits,
qualifies a bare A1 range only for a local scope, enforces case-insensitive
uniqueness by `(name, scope)`, rejects ListObject table-name collisions, and
protects `_xlnm.*` and `Slicer_*` names owned by other Office features.
Semantic get/query, ordinary typed remove, batch, exact replay, CLI, Rust, and
standard MCP share the same value. Strict/transitional SpreadsheetML and
unknown defined-name attributes are retained; unknown collection or child
content fails closed when it cannot be preserved. This contract owns
defined-name lifecycle, not external-link authoring or complete Spreadsheet
parity; supported names participate when an explicit native cell-formula
recalculation references them.

Native Spreadsheet AutoFilters use a closed typed contract shared by worksheet
filters and ListObject tables. One value owns a normalized rectangular A1
range plus at most one criterion for each zero-based range column. Supported
criteria are exact value sets with optional blanks; equality, text, and ordered
comparisons; between/not-between; blanks/non-blanks; top/bottom count or
percentage; and closed dynamic average, relative-date, month, and quarter
families. Worksheet lifecycle uses stable `/Sheet/autofilter` and
`/Sheet/autofilter/filterColumn[N]` paths; table filters appear below
`/Sheet/table[N]/autofilter`. Rust, versioned batch, CLI, standard MCP, exact
replay, and the Office Skill use the same value. The strict/transitional writer
sorts columns deterministically, escapes wildcard literals, rejects duplicate
or out-of-range columns, protects table/merge geometry, and fails closed for
imported date-group items, color/icon filters, embedded sort state,
extensions, comments, or unknown attributes. Physical row sorting is owned by
the separate typed sort contract below; an imported AutoFilter with unsupported
embedded sort state remains readable but non-mutable.

Native Spreadsheet sorting physically reorders a worksheet range through one
stable, ordered multi-key contract. A worksheet path auto-detects its used
range; callers may instead supply `/Sheet/A1:D100`. Keys are absolute `A:XFD`
columns inside that range, the first selected row stays fixed only when
`header=true`, text comparison is case-insensitive by default, numbers sort
before text, blanks always remain last, and records equal on every key retain
their source order. Partial-column sorts move only selected cells; cells outside
the range and destination row properties stay fixed. The editor persists a
worksheet SpreadsheetML `sortState`, exposed as `/Sheet/sort` and ordered
`/Sheet/sort/key[N]` nodes. Removing `/Sheet/sort` deletes metadata only and
never reverses the physical row order.

Rust, versioned batch/replay, CLI, standard MCP, and the Office Skill share the
same sort value. Sorts accept 1–64 unique keys and at most 100,000 selected
cells. Exact mutable ListObject or worksheet-AutoFilter ranges, or their exact
data ranges, are supported. Table totals rows, formulas anywhere in the
workbook, intersecting merges, pivots, unknown existing sort state, and drawing
anchors that cannot follow one record losslessly fail closed. Hyperlinks,
comments/VML notes, data validation, conditional formatting, protected ranges,
ignored errors, and supported drawing anchors move with their records; chart
caches are cleared and the worksheet used dimension is recomputed after the
physical change.

Native Spreadsheet delimited import accepts bounded UTF-8 CSV or TSV from a
regular file or stdin and writes it into an existing worksheet from an explicit
A1 start cell. The parser supports a leading BOM, CRLF, quoted delimiters,
embedded newlines, and doubled quotes; malformed quote state fails atomically.
One request is limited to 8 MiB and a 100,000-cell rectangular extent. Explicit
empty fields clear existing target cells, while missing trailing fields in a
ragged row leave those cells unchanged.

Typed inference stores formulas, finite numbers, booleans, ISO dates/times, and
text without a Python, Node.js, OfficeCLI, or spreadsheet-application runtime.
Dates honor the workbook's 1900/1904 system and receive a native date number
format. Inferred formulas pass the same bounded native syntax parser as direct
cell writes; malformed expressions fail the complete import. Import does not
implicitly calculate formulas; run the explicit native recalculation command
or mutation when fresh cached results are required. Header mode atomically
installs the worksheet AutoFilter and a canonical frozen pane below the header.
Frozen pane state is readable at `/Sheet/freeze`, is set through the typed
batch/Rust/MCP contract, and is removed through the ordinary typed `remove`
mutation. Strict/transitional SpreadsheetML and unknown view content are
preserved; unsupported pane content reports `nativeMutable=false` and fails
closed on mutation.

Native Spreadsheet formula calculation builds a deterministic bounded
dependency graph across worksheets, ranges, spills, and workbook- or
worksheet-scoped names. The closed built-in registry implements
`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `SQRT`, `POWER`,
`MOD`, `ROUND`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `CONCAT`,
`CONCATENATE`, `ROW`, `COLUMN`, `SEQUENCE`, `TRANSPOSE`, `PI`, and `NA`.
Operators and typed blank, number, text, boolean, and Spreadsheet error values
participate in scalar or rectangular array calculation.

The read-only Rust calculation API leaves package bytes unchanged. The editor,
versioned batch, replay, CLI `office native recalculate`, and standard MCP
`recalculate-spreadsheet-formulas` mutation atomically write typed OOXML
caches, canonical array anchors, spill children, and calculated-workbook
metadata. Spill children are read-only; edit or remove their formula anchor.
Exact replay accepts canonical formula storage and natively cached array
anchors. It fails closed for physical distinctions typed mutations cannot
reproduce, including explicit `t="normal"` storage and uncached or malformed
array anchors.
ListObject structured references resolve table names or display names.
`Sales[Qty]` and `Sales[[Qty]:[Price]]` select data rows; `#All`, `#Data`,
`#Headers`, and `#Totals` select structural rows; and `Sales[@Qty]`,
`Sales[[#This Row],[Qty]]`, or table-local `[@Qty]` select the current data row.
Table-local forms require the formula cell to be inside the inferred table.
Missing tables, columns, or requested structural rows, disjoint columns,
non-canonical forms, cycles, unsupported or qualified functions, and
external-workbook reads fail with stable typed errors and roll back the whole
batch. The engine never fetches an external workbook or falls back to a shell
or script runtime. One formula is limited to 8,192 characters, depth
128 (including nested named-reference resolution), and 8,192 AST nodes; one
reference value to 100,000 areas; one graph to 100,000 formulas, 1,000,000
edges, and 1,000,000 formula-cell reference visits; one materialized array or
function call to 100,000 cells; and one text result to 1 MiB.
One calculation pass is also limited to 100,000 cumulative spill children and
200,000 OOXML cell writes, plus 8 MiB of cumulative text-result bytes.

Native Spreadsheet tables use a separate closed ListObject contract. Add and
set own the workbook-wide `name`, optional distinct `displayName`, final
rectangular A1 range, one exact column identity per range column, header/totals
row state, typed filter criteria, built-in light/medium/dark style identity,
and first/last-column plus row/column-stripe flags; ordinary typed `remove`
owns deletion. When a header is enabled, its names are stamped into the first
row and the table-owned AutoFilter range excludes an enabled totals row. The
`set` lifecycle keeps common structured references consistent: changed table
names/display names and position-mapped column identities are rewritten across
cell formulas, defined names, conditional formats, data validations, charts,
and table formulas without touching string literals or external-workbook
references. Table-local references are rewritten only when their ListObject
context is provable. Unsafe local-reference geometry changes fail with
`use.office.spreadsheet_table_formula_rewrite_unsupported`, and `remove`
fails with `use.office.spreadsheet_table_referenced` while a live structured
reference still targets the table. All checks and rewrites share the table
mutation's atomic rollback boundary.
The editor rejects
Excel-identifier and A1/R1C1 name errors,
case-insensitive table/defined-name collisions, duplicate columns, missing data
rows, table/merge/worksheet-AutoFilter overlap, and unsafe relationship graphs.
Semantic reads expose stable `/Sheet/table[N]` and child column paths. Rust,
versioned batch, CLI, standard MCP, exact replay, and the Office Skill share the
same value while strict/transitional SpreadsheetML and supported unknown root
or style data are retained. Imported calculated columns, totals functions,
date-group/color/icon filters, unsupported embedded sort state, custom styles,
query tables, and external data remain explicit gaps and fail closed when a
lossless typed mutation cannot be proved. Exact mutable table and data ranges
can still be physically sorted through the separate sort contract.

The engine also creates, updates, reads,
queries, and removes typed hyperlinks. Word owns
external HTTP/HTTPS/mailto links and internal bookmark targets in body,
header, and footer paragraphs, with display text and tooltips; Spreadsheet
owns external links and internal workbook locations on cells or bounded
rectangular ranges, with display text and tooltips, auto-creating a missing
single linked cell; Presentation owns external shape-wide links and internal
jumps to existing slides, with optional tooltips. External targets reject
embedded credentials, active or relative schemes, and malformed URIs; semantic rendering
keeps every relationship inert and never fetches it. The same typed engine
creates, updates, reads, queries, and removes classic Office comments. Word
comments anchor to a main-document paragraph or run and expose stable
`/comments/comment[N]` paths. Spreadsheet comments are classic cell notes with
an author table, VML note drawing, and `/SheetName/A1/comment` paths, including
notes on otherwise blank cells. Presentation comments use legacy per-slide
comment and shared-author parts, optional EMU coordinates, and
`/slide[N]/comment[M]` paths. Removing an owning Word node, Spreadsheet cell or
range, or Presentation slide also removes its owned comment resources. This is
plain legacy-comment scope, not complete modern threaded-comment parity:
replies, resolved state, writable dates, rich bodies, Word header/footer
anchors, Spreadsheet threaded comments, and modern PowerPoint threaded
comments remain outside the typed contract. The engine can safely inspect
existing XML parts and replace non-OPC-metadata XML parts while preserving the
root QName and validating the final document. Known chart, header, and footer
part carriers can be created together with their content type and owner
relationship. The typed mutation layer also adds and removes Word paragraphs
and basic table/row/cell structures,
creates real Presentation DrawingML tables, appends grid-conformant rows,
fills underfull rows, edits cell text, and exposes table columns as stable
virtual paths. Columns can be inserted, resized in EMUs, removed, moved,
copied, or swapped while the grid, every row, and the graphic-frame width stay
in sync. Tables and structurally safe rows/cells can also be removed. It
upserts typed Spreadsheet text, number, boolean, and formula cells, removes
cells and bounded cell ranges, structurally inserts or deletes rows and columns,
and adds, removes, renames, reorders, or copies worksheets. It also exposes
typed `move`, `copy`, and `swap` mutations with zero-based or path-relative
placement. Word supports same-parent paragraph, table, row, cell, and run
moves and swaps plus identity-free paragraph, table, row, and run copies.
Spreadsheet supports worksheet arrangement and dense plain-row arrangement with
row/cell renumbering. Presentation supports slide arrangement, same-slide
top-level object movement/swaps, layout-only slide copies, and relationship-free
plain-shape copies with fresh non-visual identities. Cross-parent ownership
migration, formula-bearing or reference-rich row arrangement, identity-bearing
Word copies, table-cell copies, and relationship-owning Presentation copies fail
closed before save. Worksheet copy
clones the owned OPC relationship subgraph while preserving shared workbook
resources; removal garbage-collects only unshared descendants. Structural
Spreadsheet edits rewrite affected A1 formulas, defined names, worksheet
metadata, tables, comments, VML notes, drawing anchors, and chart references;
unsupported pivot and 3D-reference cases fail closed before save. Presentation
slides, text shapes, and basic tables also support native add/remove. PNG, JPEG,
and GIF can be embedded as real Word inline pictures, Spreadsheet one-cell
drawing anchors, and Presentation slide pictures; semantic reads and
reference-aware removal use the same cross-format `Picture` contract. Saves are
atomic and reject a changed source revision instead of overwriting another
writer.
Cross-format template merge replaces `{{key}}` text in Word document and
auxiliary text parts, Spreadsheet string cells, and Presentation slides and
notes while preserving split-run formatting and reporting unresolved keys.

The native document engine does not require Microsoft Office, LibreOffice,
OfficeCLI, Python, Node.js, or .NET. LibreOffice may be used only by optional
CI interoperability checks and is never part of document execution. Optional
PNG screenshot output requires the `browser` feature and a ready A3S Browser
provider because it captures the native semantic HTML through the existing
Browser contract.

The explicit `office native` CLI exposes in-process blank creation, reads,
typed add/set/remove/move/copy/swap, scoped literal/regex replacement,
rich-text, exact Spreadsheet merged-cell, stable Spreadsheet physical sorting
with persisted sort state, bounded CSV/TSV import with typed inference and
header filter/freeze behavior, explicit Spreadsheet formula recalculation,
worksheet/table AutoFilter, data-validation, conditional-format, defined-name,
ListObject table, hyperlink, and legacy-comment operations, constrained raw
XML access,
known typed part carriers, exact replay artifacts for a constrained canonical
subset, visible PNG/JPEG/GIF pictures, and atomic mutation batches, plus
dependency-free template merge and semantic rendering today. HTML and SVG are
available for Word, Spreadsheet, and Presentation; bounded annotated and issue
views are available for all three formats; and Browser-injected PNG screenshots
are available for all three formats. An authenticated, loopback-only foreground
watch provides full saved-revision refresh for all three formats without a
resident pipe or mutation endpoint. `mcp serve office-native` exposes the same
editor, annotated/issue analysis, and screenshot composition through typed
standard MCP tools and bounded in-memory sessions.
The packaged `a3s-use-office` Skill exposes the same product boundaries to
agents without starting OfficeCLI. Discover its metadata with
`office skills list`, read only its `SKILL.md` with
`office skills get a3s-use-office`, append its four format/MCP references with
`--full`, or locate the installed directory with `office skills path`. The
capability snapshot binds the Skill path and lowercase SHA-256 so a resident
host can verify the bytes before loading them. Resident Code hosts receive the
native engine as canonical route `use/office` targeting `office-native`; a ready
OfficeCLI installation is projected separately as `use/office-compat` targeting
`office-compat`.
Other `0.1.x` commands and the default `mcp serve office` target still use a
compatibility backend pinned to OfficeCLI `1.0.136`. This is a migration
boundary, not a native-promotion claim. The default routes will be promoted
only after mutation, fidelity, rendering, compatibility, and cross-application
interoperability gates pass. The first real compatibility CLI command prepares
that pinned provider when first-use policy allows. In Code, the native Office
worker requests `office_install_compat` through parent confirmation only when
the requested operation is outside the native surface.

```bash
# Inspect without downloading anything.
a3s use office doctor --json

# Create by extension; an existing destination is never overwritten.
a3s-office create report.docx --json
a3s-office create workbook.xlsx --json
a3s-office create deck.pptx --json

# Read without OfficeCLI, Microsoft Office, or LibreOffice.
a3s-office get report.docx /body --depth 2 --json
a3s-office query report.docx 'p[style=Heading1]' --json
a3s-office view report.xlsx stats --json
a3s-office view report.docx issues --json
a3s-office view workbook.xlsx issues --type formula_not_evaluated --limit 20 --json
a3s-office view report.docx html --output report.html --json
a3s-office view workbook.xlsx html --output workbook.html --json
a3s-office view report.docx svg --output report.svg --json
a3s-office view workbook.xlsx svg --output workbook.svg --json
a3s-office view deck.pptx svg --output deck.svg --json
a3s-office view report.docx screenshot --output report.png --timeout-ms 30000 --json
a3s-office watch deck.pptx --port 0
a3s-office validate deck.pptx --json

# Inspect a safely parsed XML part inline or export its original bytes.
a3s-office raw report.docx /word/document.xml --json
a3s-office raw report.docx /word/document.xml --output document.xml --json

# Replace one existing XML part; --output is an optional Office save-as target.
a3s-office raw-set report.docx /word/document.xml --input document.xml --output updated.docx --json

# Create known part carriers and receive their owner relationship IDs.
a3s-office add-part report.docx / --type header --json
a3s-office add-part report.docx / --type chart --json
a3s-office add-part workbook.xlsx /Sheet1 --type chart --json
a3s-office add-part deck.pptx '/slide[1]' --type chart --json

# Replace text in place or save to a separate OOXML document.
a3s-office set report.docx /body/p[1] --text 'Updated' --json
a3s-office set report.xlsx /Sheet1/B2 --text '42' --output updated.xlsx --json

# Find and replace within a semantic scope. Literal matching is the default;
# --regex enables Rust regular expressions and $name/$1 capture expansion.
a3s-office set report.docx /body --find 'Q1 2025' --replace 'Q1 2026' --json
a3s-office set report.docx / --find 'Q([1-4]) 2025' --replace 'Q$1 2026' --regex --json
a3s-office set workbook.xlsx /Sheet1/A1:C20 --find Draft --replace Final --json
a3s-office set deck.pptx '/slide[1]/notes' --find internal --replace confidential --json

# Apply typed text formatting. Word and Presentation character properties use
# run paths; their alignment uses paragraph paths. Word sizes must be exact
# half-point increments. Spreadsheet ranges may combine content and formatting.
# Strikethrough is native for Word and Spreadsheet, but not Presentation.
a3s-office set report.docx '/body/p[1]/r[1]' --bold true --italic false --underline double --script superscript --strikethrough true --double-strikethrough false --text-case small-caps --highlight yellow --language en-US --font-family Aptos --font-size 14 --text-color 123456 --json
a3s-office set report.docx '/body/p[1]' --align center --json
a3s-office set workbook.xlsx /Sheet1/A1:C1 --bold true --underline single --script baseline --strikethrough false --font-size 11.5 --text-color 0066CC --align center --json
a3s-office set deck.pptx '/slide[1]/shape[1]/paragraph[1]/run[1]' --italic true --underline double --script subscript --text-case all-caps --highlight cyan --language zh-CN --font-family 'Aptos Display' --font-size 20 --json
a3s-office set deck.pptx '/slide[1]/shape[1]/paragraph[1]' --align center --json

# Apply Spreadsheet cell presentation independently or in the same atomic set
# as content/text formatting. Fill and border colors accept six-digit RGB.
a3s-office set workbook.xlsx /Sheet1/A1:C3 --number-format currency --fill FFF2CC --border-all thin --border-color 808080 --border-bottom double --border-bottom-color 000000 --vertical-align center --wrap-text true --text-rotation 0 --indent 1 --shrink-to-fit false --reading-order ltr --json
a3s-office set workbook.xlsx /Sheet1/D1 --number 0.125 --bold true --number-format percent --fill 0066CC --json
a3s-office set workbook.xlsx /Sheet1/E1 --border-diagonal slant-dash-dot --border-diagonal-color FF0000 --border-diagonal-up true --border-diagonal-down false --json

# Merge one normalized Spreadsheet range, or unmerge the exact same range.
# Content, text format, cell format, hyperlink, and merge state can share one
# atomic set command.
a3s-office set workbook.xlsx /Sheet1/A1:C1 --text 'Quarter' --bold true --merge-cells true --json
a3s-office set workbook.xlsx /Sheet1/A1:C1 --merge-cells false --json

# Import a bounded UTF-8 CSV/TSV source. Header mode also installs the
# worksheet AutoFilter and canonical frozen pane in the same transaction.
a3s-office import workbook.xlsx /Sheet1 source.csv --header --start-cell A1 --json
a3s-office import workbook.xlsx /Sheet1 --stdin --format tsv --output imported.xlsx --json
a3s-office get workbook.xlsx /Sheet1/freeze --json

# Add, inspect, replace, clear, and remove one worksheet AutoFilter. Each
# --filter is a strict JSON object with a zero-based column and typed criteria.
a3s-office add workbook.xlsx /Sheet1 --type auto-filter --range A1:C20 --filter '{"column":0,"criteria":{"type":"values","values":["Open","Closed"],"includeBlanks":true}}' --filter '{"column":2,"criteria":{"type":"greater-than","value":"100"}}' --json
a3s-office query workbook.xlsx 'filtercolumn[criteriaType=greater-than]' --json
a3s-office get workbook.xlsx /Sheet1/autofilter --depth 2 --json
a3s-office set workbook.xlsx /Sheet1/autofilter --range B2:D30 --filter '{"column":1,"criteria":{"type":"dynamic","kind":"this-month"}}' --json
a3s-office set workbook.xlsx /Sheet1/autofilter --clear-filters --json
a3s-office remove workbook.xlsx /Sheet1/autofilter --json

# Physically sort selected records by ordered absolute columns. A worksheet
# path auto-detects its used range. Remove the resulting semantic sort node to
# clear only persisted metadata; the physical row order remains unchanged.
a3s-office sort workbook.xlsx /Sheet1/A1:D100 --key B:desc --key C:asc --header true --case-sensitive false --json
a3s-office get workbook.xlsx /Sheet1/sort --depth 1 --json
a3s-office remove workbook.xlsx /Sheet1/sort --json

# Add, inspect, replace, and remove one native Spreadsheet ListObject table.
# The range is final and includes enabled header/totals rows. Repeat
# --table-column exactly once for every range column. Table --filter values use
# the same strict zero-based contract as worksheet AutoFilters.
a3s-office add workbook.xlsx /Sheet1 --type table --name Sales --range F1:H4 --table-column Name --table-column Qty --table-column Price --filter '{"column":1,"criteria":{"type":"top","count":10}}' --style medium:4 --json
a3s-office query workbook.xlsx 'table[name=Sales]' --json
a3s-office get workbook.xlsx '/Sheet1/table[1]' --depth 1 --json
a3s-office set workbook.xlsx '/Sheet1/table[1]' --name Inventory --display-name InventoryView --range B2:D6 --table-column Item --table-column Units --table-column Cost --totals-row true --style dark:2 --show-row-stripes false --show-column-stripes true --json
a3s-office remove workbook.xlsx '/Sheet1/table[1]' --json

# Add, inspect, update, and remove typed Spreadsheet data validation. Repeated
# --range values form one rule over disjoint areas; set preserves omitted fields.
a3s-office add workbook.xlsx /Sheet1 --type data-validation --validation-type list --range A2:A20 --range C2:C20 --formula1 'Draft,Review,Approved' --prompt-title Status --prompt 'Choose a workflow state' --error-title 'Invalid status' --error-message 'Choose a listed state' --json
a3s-office query workbook.xlsx 'dataValidation[type=list]' --json
a3s-office get workbook.xlsx /Sheet1/C3 --json
a3s-office set workbook.xlsx '/Sheet1/dataValidation[1]' --validation-type whole --range B2:B50 --operator between --formula1 18 --formula2 120 --allow-blank false --error-style warning --json
a3s-office remove workbook.xlsx '/Sheet1/dataValidation[1]' --json

# Add, query, partially update, and remove native Spreadsheet conditional
# formats. Rule-specific options are closed and validated before atomic save.
a3s-office add workbook.xlsx /Sheet1 --type conditional-format --rule-type cell-is --range A2:A20 --operator greater-than --formula1 80 --fill C6EFCE --text-color 006100 --bold true --json
a3s-office add workbook.xlsx /Sheet1 --type conditional-format --rule-type data-bar --range B2:B20 --color 638EC6 --min min --max number:100 --json
a3s-office add workbook.xlsx /Sheet1 --type conditional-format --rule-type color-scale --range C2:C20 --min-color F8696B --midpoint percentile:50 --mid-color FFEB84 --max-color 63BE7B --json
a3s-office add workbook.xlsx /Sheet1 --type conditional-format --rule-type icon-set --range D2:D20 --icon-set 3-traffic-lights-1 --reverse true --json
a3s-office query workbook.xlsx 'conditionalFormatting[type=iconSet]' --json
a3s-office set workbook.xlsx '/Sheet1/cf[1]' --formula1 90 --fill FFEB9C --stop-if-true true --json
a3s-office remove workbook.xlsx '/Sheet1/cf[2]' --json

# Add, inspect, update, and remove typed Spreadsheet defined names. A sheet
# parent defaults to local scope and qualifies a bare A1 ref automatically.
a3s-office add workbook.xlsx / --type named-range --name Revenue --ref 'Sheet1!$A$2:$A$20' --scope workbook --comment 'Workbook revenue' --json
a3s-office add workbook.xlsx /Sheet1 --type named-range --name Status --ref A2:A20 --json
a3s-office query workbook.xlsx 'namedrange[scope=Sheet1]' --json
a3s-office get workbook.xlsx '/namedrange[@name=Revenue][@scope=workbook]' --json
a3s-office set workbook.xlsx '/namedrange[@name=Status][@scope=Sheet1]' --name WorkflowStatus --ref B2:B20 --volatile false --json
a3s-office remove workbook.xlsx '/namedrange[@name=Revenue][@scope=workbook]' --json

# Add or update inert hyperlinks. External targets accept only absolute
# HTTP/HTTPS/mailto URIs without credentials. Word internal targets are bookmark
# names; Spreadsheet internal targets are workbook locations; Presentation
# internal targets are existing slide[N] paths. Presentation keeps shape text.
a3s-office add report.docx '/body/p[1]' --type hyperlink --url https://example.com/report --display 'Open report' --tooltip 'A3S report' --json
a3s-office set report.docx '/body/p[1]/hyperlink[1]' --location section_1 --display 'Jump to section' --json
a3s-office set report.docx '/header[1]/p[1]' --url https://example.com/header --display 'Header link' --json
a3s-office set workbook.xlsx /Sheet1/A1 --location 'Sheet1!B2' --display B2 --json
a3s-office set workbook.xlsx /Sheet1/B2:C3 --url https://example.com/range --display Range --json
a3s-office set deck.pptx '/slide[1]/shape[1]' --url https://example.com/slides --tooltip 'Open slides' --json
a3s-office set deck.pptx '/slide[1]/shape[1]/hyperlink' --location 'slide[2]' --tooltip 'Next slide' --json
a3s-office query report.docx hyperlink --json
a3s-office remove report.docx '/body/p[1]/hyperlink[1]' --json

# Add, update, discover, and remove classic Office comments. Word uses a body
# paragraph or run anchor; Spreadsheet uses one cell; Presentation uses a slide
# and optionally accepts a complete x/y EMU coordinate pair.
a3s-office add report.docx '/body/p[1]' --type comment --author Alice --initials AL --text 'Please reword this' --json
a3s-office set report.docx '/comments/comment[1]' --author Bob --initials BO --text 'Reviewed' --json
a3s-office add workbook.xlsx /Sheet1/B2 --type comment --author Alice --text 'Check this formula' --json
a3s-office add deck.pptx '/slide[1]' --type comment --author Alice --initials AL --text 'Rework this slide' --x-emu 914400 --y-emu 457200 --json
a3s-office query deck.pptx comment --json
a3s-office remove workbook.xlsx /Sheet1/B2/comment --json

# Preserve Spreadsheet value types. Formula writes validate and store the
# expression; explicit recalculation computes and writes cached values.
a3s-office set workbook.xlsx /Sheet1/A1 --number 42.5 --json
a3s-office set workbook.xlsx /Sheet1/B1 --boolean true --json
a3s-office set workbook.xlsx /Sheet1/C1 --formula 'SUM(A1:B1)' --json
a3s-office recalculate workbook.xlsx --output calculated.xlsx --json

# Set or remove a bounded rectangular range atomically.
a3s-office set workbook.xlsx /Sheet1/A2:C4 --number 0 --json
a3s-office remove workbook.xlsx /Sheet1/B3:C4 --json

# Edit worksheet structure and ordering without invoking OfficeCLI.
a3s-office insert-rows workbook.xlsx /Sheet1 2 --count 3 --json
a3s-office delete-columns workbook.xlsx /Sheet1 B --count 2 --json
a3s-office rename-sheet workbook.xlsx /Sheet1 'Q1 Data' --json
a3s-office move-sheet workbook.xlsx '/Q1 Data' 1 --json
a3s-office copy-sheet workbook.xlsx '/Q1 Data' 'Q1 Copy' --position 2 --json

# Add and remove native document structures.
a3s-office add report.docx /body --type paragraph --text 'Summary' --json
a3s-office add report.docx /body --type table --rows 2 --columns 3 --json
a3s-office add report.docx '/body/tbl[1]' --type row --columns 3 --json
a3s-office add report.docx '/body/tbl[1]/tr[3]' --type cell --text 'Total' --json
a3s-office add workbook.xlsx / --type sheet --name Data --json
a3s-office add deck.pptx / --type slide --text 'Results' --json
a3s-office add deck.pptx '/slide[1]' --type shape --text '42%' --json
a3s-office add deck.pptx '/slide[1]' --type table --rows 3 --columns 2 --json
a3s-office set deck.pptx '/slide[1]/table[1]/tr[1]/tc[1]' --text 'Metric' --json
a3s-office add deck.pptx '/slide[1]/table[1]' --type row --columns 2 --json
a3s-office add deck.pptx '/slide[1]/table[1]' --type column --index 1 --text 'Q2' --json
a3s-office set deck.pptx '/slide[1]/table[1]/col[2]' --width-emu 2000000 --json
a3s-office move deck.pptx '/slide[1]/table[1]/col[1]' --after '/slide[1]/table[1]/col[2]' --json
a3s-office remove deck.pptx '/slide[1]/table[1]/col[3]' --json
a3s-office remove deck.pptx '/slide[1]/table[1]/tr[4]' --json
a3s-office remove workbook.xlsx /Data --json
a3s-office remove deck.pptx '/slide[1]/shape[2]' --json

# Embed bounded PNG, JPEG, or GIF data as a real DrawingML picture. Supplying
# one dimension preserves the source aspect ratio; supplying both is explicit.
a3s-office add report.docx /body --type picture --input logo.png --name Logo --alt 'A3S logo' --width 320 --json
a3s-office add workbook.xlsx /Sheet1/B2 --type picture --input chart.jpeg --width 480 --height 270 --json
a3s-office add deck.pptx '/slide[1]' --type picture --input photo.gif --json
a3s-office remove deck.pptx '/slide[1]/picture[1]' --json

# Arrange supported semantic nodes. --index is zero-based; --before and --after
# resolve stable pre-mutation paths. A copy defaults to immediately after its
# source.
a3s-office move report.docx '/body/p[3]' --before '/body/p[1]' --json
a3s-office copy workbook.xlsx '/Q1 Data' --name 'Q1 Copy' --after '/Q1 Data' --json
a3s-office swap deck.pptx '/slide[1]' '/slide[3]' --json

# Merge a template into a separate output. JSON may be inline, @file, or an
# existing .json path. Existing outputs require an explicit --force.
a3s-office merge template.docx report.docx --data @report.json --json
a3s-office merge template.xlsx report.xlsx --data '{"quarter":"Q3"}' --json

# Apply a bounded, versioned mutation document atomically.
a3s-office batch deck.pptx --input mutations.json --json

# Dump the exactly replayable root subset, then replay it into a native blank.
a3s-office dump report.docx --output report.replay.json --json
a3s-office create restored.docx --json
a3s-office batch restored.docx --input report.replay.json --json

# Optional compatibility pre-warm. The following compatibility commands also
# prepare this pinned provider on first use.
a3s install use/office
a3s use office get report.docx /body --json
a3s use office batch report.xlsx --input updates.json --json

# Launch the explicit native standard MCP preview. No OfficeCLI is consulted.
a3s-office mcp

# Launch the current compatibility standard MCP server.
a3s use mcp serve office
```

The native MCP process exposes 12 document tools plus the confirmed
`office_install_compat` compatibility installer: `office_validate`,
`office_create`, `office_open`, `office_list`, `office_get`, `office_query`,
`office_view`, `office_raw_xml`, `office_apply_batch`,
`office_merge_template`, `office_save`, and `office_close`. It accepts no shell
command string and defines no A3S RPC dialect; stdio carries only standard MCP.
Each process owns at most 64 sessions. Mutation batches are atomic in memory,
limited to 10,000 mutations and 8 MiB of JSON, and remain unsaved until
`office_save`. Results are limited to 8 MiB, raw XML responses to 1 MiB, and
queries to at most 1,000 returned nodes. `office_close` rejects dirty sessions
unless the caller saves or explicitly sets `discard=true`. `office_view`
accepts `html` and `svg` for all three formats in addition to text, annotated,
outline, and statistics. The typed `annotated` view flattens stable semantic
paths, node types, text, styles, and observed formatting; its `limit` is 1
through 1,000 and defaults to 200. The `issues` view accepts an optional typed
`issueType` and a `limit` from 1 through 1,000, defaulting to 200. It also
accepts `screenshot` for all three formats; that mode requires a no-clobber
local `output` ending in `.png` and accepts an optional `timeoutMs` from 1
through 120,000.

Native issue analysis is conservative and read-only. It currently reports
missing picture alternative text, missing or incompatible internal part
relationships, formulas without cached results, formulas that reference a
missing worksheet, cached or explicit formula errors, and low contrast between
explicit RGB run text and its own shape fill. Reports include stable category,
subtype, severity, path, context, and suggestion fields; filtering happens
before the bounded result window, and `count`, `returned`, and `truncated`
remain explicit. It does not infer text overflow, object overlap, theme or
inherited colors, or Microsoft Office layout behavior. A clean report is not a
full fidelity or interoperability certification.

Native render artifacts are deterministic, standalone, and network-free. They
contain no timestamp or source path, escape document text and attributes, carry
stable semantic paths as `data-path`, and embed only validated internal
PNG/JPEG/GIF parts as `data:` URLs. HTML declares a restrictive CSP and uses a
sparse observed-cell representation instead of expanding large Spreadsheet
gaps. Each render is bounded to 16 MiB while it is composed. CLI `--output`
publishes through an atomic no-clobber file operation; inline MCP output remains
subject to the stricter 8 MiB structured-result limit. These are semantic
previews, not a Microsoft Office layout-fidelity claim. Screenshot mode stages
the same deterministic HTML privately, opens its `file://` URL through the
existing `PageRenderer`, and validates one regular PNG plus its size and
SHA-256 receipt before atomic no-clobber publication. It defaults to a 30-second
deadline, caps the deadline at 120 seconds, and caps the PNG at 64 MiB. It does
not fetch external relationships or consult OfficeCLI.

`office native watch <file>` renders the same bounded all-format HTML, binds
only `127.0.0.1`, selects an ephemeral port by default, and prints a URL with a
fresh 256-bit capability token. Every page, status response, and standard SSE
stream requires that token or its HttpOnly same-site cookie and validates the
exact loopback `Host`. The wrapper runs its own fixed script while the document
preview stays in a sandboxed iframe under the renderer's script-free CSP.
Atomic saves from another `office native` process trigger a full refresh.
Transient missing or invalid revisions leave the last valid preview visible,
publish a typed error state, and retry until recovery. The server is read-only:
it has no mutation/RPC endpoint, never opens an external relationship, and
does not observe unsaved `office-native` MCP sessions until `office_save`.
`--timeout-ms` bounds automated runs; otherwise Ctrl+C stops the foreground
server. Interactive editing, selection/mark overlays, and layout goldens remain
open.

Native batch input is an ordinary JSON document, not an RPC protocol. The
current schema is:

```json
{
  "schemaVersion": 1,
  "mutations": [
    {
      "operation": "replace-text",
      "path": "/body",
      "replacement": {
        "find": "Q([1-4]) 2025",
        "replace": "Q$1 2026",
        "mode": "regex"
      }
    },
    {
      "operation": "set-text",
      "path": "/body/p[1]",
      "text": "Updated"
    },
    {
      "operation": "add-paragraph",
      "parent": "/body",
      "text": "Summary"
    },
    {
      "operation": "remove",
      "path": "/body/p[2]"
    },
    {
      "operation": "move",
      "path": "/body/p[3]",
      "position": {
        "kind": "before",
        "path": "/body/p[1]"
      }
    }
  ]
}
```

The whole batch rolls back if any mutation fails. Inputs are limited to 8 MiB
and 10,000 mutations. The version 1 mutation set is `replace-text`, `set-text`,
`set-text-format`, `set-cell-format`, `add-data-validation`,
`set-data-validation`, `add-conditional-format`, `set-conditional-format`,
`add-named-range`, `set-named-range`, `add-spreadsheet-table`,
`set-spreadsheet-table`, `add-spreadsheet-auto-filter`,
`set-spreadsheet-auto-filter`, `sort-spreadsheet-range`,
`import-spreadsheet-delimited`, `set-spreadsheet-frozen-pane`, `merge-cells`,
`unmerge-cells`, `recalculate-spreadsheet-formulas`,
`set-hyperlink`, `set-comment`, `set-table-column-width`,
`set-cell-value`, `add-paragraph`,
`add-table`, `add-table-row`, `add-table-column`, `add-table-cell`,
`add-comment`, `add-worksheet`, `insert-rows`, `delete-rows`, `insert-columns`,
`delete-columns`, `rename-worksheet`, `move-worksheet`, `copy-worksheet`,
`move`, `copy`, `swap`, `replace-xml-part`, `add-part`, `add-slide`, `add-shape`,
`add-image`, and `remove`.

A data-validation mutation uses the same atomic document and standard MCP
payload. Remove a rule with the ordinary typed `remove` mutation:

```json
{
  "operation": "add-data-validation",
  "sheet": "/Sheet1",
  "validation": {
    "type": "whole",
    "ranges": ["B2:B50", "D2:D50"],
    "operator": "between",
    "formula1": "18",
    "formula2": "120",
    "allowBlank": false,
    "errorStyle": "warning"
  }
}
```

Each rule accepts 1–1,024 disjoint ranges and each worksheet accepts at most
65,534 rules. Formulas are limited to 255 characters; prompt/error titles to
32, prompts to 255, and error messages to 225. List and custom rules reject
operators and `formula2`; comparison rules require an operator and require
`formula2` only for `between` or `notBetween`. Invalid or overlapping input
rolls back the complete in-memory batch.

A conditional-format mutation uses one complete closed rule value. CLI `set`
can merge omitted options, while batch and standard MCP set replace the complete
value:

```json
{
  "operation": "add-conditional-format",
  "sheet": "/Sheet1",
  "conditionalFormat": {
    "ranges": ["A2:A20"],
    "stopIfTrue": true,
    "rule": {
      "type": "cellIs",
      "operator": "greaterThan",
      "formula1": "80",
      "format": {
        "fill": {"red": 198, "green": 239, "blue": 206},
        "bold": true
      }
    }
  }
}
```

Closed classic predicates plus data bars, two/three-color scales, and standard
3/4/5-icon sets are supported. Threshold, range, priority, shared-range, and
loss-preservation failures roll back the whole batch. Rule formulas are stored,
not evaluated, and the semantic preview is not Excel rendering evidence.

A named-range mutation uses one complete scoped value. Deletion reuses the
ordinary typed `remove` mutation:

```json
{
  "operation": "add-named-range",
  "namedRange": {
    "name": "Revenue",
    "ref": "'Sheet1'!$A$2:$A$20",
    "scope": "workbook",
    "comment": "Workbook revenue",
    "volatile": false
  }
}
```

Use `set-named-range` with a stable `path` and a complete `namedRange` value.
Names are limited to 255 characters, refs to 8,192, comments to 255, and a
workbook to 65,536 defined names. Workbook-scoped bare A1 refs, leading `=`,
cross-workbook refs without external-link parts, reserved Office-managed names,
duplicate `(name, scope)` identities, and ListObject table-name collisions fail
atomically. Use the explicit scope `worksheet:workbook` for a local name on a
worksheet literally named `workbook`.

A worksheet AutoFilter mutation uses one complete range-and-columns value.
Table `filters` use the same column objects:

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

Use `set-spreadsheet-auto-filter` with `/Sheet/autofilter` and a complete
`filter`; ordinary typed `remove` deletes it. CLI `set` preserves an omitted
range, replaces all criteria when one or more `--filter` objects are supplied,
and uses `--clear-filters` for an explicit empty criterion list. Column offsets
are zero-based and unique inside the range. Imported date-group/color/icon
filters, embedded sort state, extensions, and unknown content are readable but
`nativeMutable=false`. Use the separate physical sort mutation for supported
ranges; it does not flatten an unsupported imported AutoFilter.

A Spreadsheet physical sort mutation owns an ordered, stable multi-key value:

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

The path may instead be `/Sheet1` to auto-detect the used range. Keys are
unique absolute `A:XFD` columns within that range; 1–64 keys and at most
100,000 selected cells are accepted. Numbers precede text, blanks remain last
in both directions, and rows equal across every key retain source order. A
partial-column range moves only its cells, preserving cells outside the range
and destination row properties. Exact mutable table and worksheet-AutoFilter
ranges, or their exact data ranges, are supported. Formulas, totals rows,
intersecting merges, pivots, unknown sort state, and non-lossless drawing
movement fail the whole batch before save.

Read persisted metadata at `/Sheet1/sort` and ordered
`/Sheet1/sort/key[N]`. An ordinary `remove` of `/Sheet1/sort` removes only the
SpreadsheetML sort state; it does not restore the prior physical row order.
Supported record-bound hyperlinks, comments/VML notes, validations,
conditional formatting, protected ranges, ignored errors, and drawing anchors
follow the row permutation. The worksheet used dimension is recomputed and
chart caches are cleared. Exact replay emits the same
`sort-spreadsheet-range` mutation.

A Spreadsheet delimited import embeds bounded content in the typed batch or MCP
request; filesystem paths remain a CLI-only concern:

```json
{
  "operation": "import-spreadsheet-delimited",
  "sheet": "/Sheet1",
  "import": {
    "content": "Name,Amount,Date\nAlpha,42,2026-07-17",
    "format": "csv",
    "header": true,
    "startCell": "A1"
  }
}
```

`format` is `csv` or `tsv`. Header mode replaces the worksheet AutoFilter range
and its canonical frozen pane in the same transaction, so inspect those nodes
before importing into a populated sheet. Set a pane independently with
`set-spreadsheet-frozen-pane` and remove it through `/Sheet/freeze`:

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

A Spreadsheet table mutation uses one complete ListObject value. CLI `set`
preserves omitted fields; batch, Rust, and standard MCP replacements supply the
complete table:

```json
{
  "operation": "add-spreadsheet-table",
  "sheet": "/Sheet1",
  "table": {
    "name": "Sales",
    "range": "A1:C4",
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

Use `set-spreadsheet-table` with a stable `path` and complete `table`; use
ordinary typed `remove` for deletion. Styles are `none`, light 1–21, medium
1–28, or dark 1–11. `none` requires all style flags to be false. The final range
must leave at least one data row after enabled header and totals rows and must
not intersect another table, a merge, or a worksheet-level AutoFilter. Table
filters require an enabled header; `--clear-filters` clears their criteria
without removing the table-owned AutoFilter range.

An image mutation uses the same versioned batch boundary:

```json
{
  "operation": "add-image",
  "parent": "/Sheet1/B2",
  "image": {
    "format": "png",
    "data": "<base64>",
    "name": "Logo",
    "altText": "A3S logo",
    "widthPx": 320
  }
}
```

The native engine validates decoded format signatures and structure, bounds
bytes and pixel dimensions, infers a missing axis from the source aspect ratio,
and inserts a visible OOXML picture object. Word uses an inline DrawingML run,
Spreadsheet uses a one-cell drawing anchor, and Presentation uses a slide
picture. Removal deletes the XML object and an unused image relationship, then
deletes the media part and content-type declaration only when no relationship
anywhere in the package still targets it. Direct CLI image inputs must be
regular, non-symlink files no larger than 64 MiB. Normal CLI output never
contains image data; `createdImage` and batch `createdImages` receipts contain
only paths, owner/media parts, relationship ID, format, and final dimensions.

OOXML SVG image embedding is not implemented yet; the all-format SVG semantic
preview is an output format and does not alter the package. Correct OOXML SVG
image support requires an SVG part plus a raster fallback rather than treating
SVG as an ordinary bitmap.

Image replacement, crop, rotation, effects, compression controls, floating
Word wrapping, Spreadsheet two-cell sizing, and rich layout rendering also
remain outside this bounded add/read/remove milestone.

`office native dump` produces a stricter versioned batch artifact, also as
ordinary JSON:

```json
{
  "format": "a3s.office.native-replay",
  "schemaVersion": 1,
  "documentKind": "word",
  "scope": "/",
  "base": "blank",
  "baseSha256": "<sha256-of-the-uncompressed-blank-part-map>",
  "resultSha256": "<sha256-of-the-uncompressed-result-part-map>",
  "mutations": []
}
```

The first dump scope is the complete document (`/`). It accepts only content
that current typed mutations can reproduce byte-for-byte at the OOXML part-map
level: plain Word paragraphs and rectangular tables, Spreadsheet worksheets,
typed defined names, typed cells, typed worksheet/table AutoFilters, typed
ListObject tables, stable physical row order with supported typed sort state,
canonical frozen panes and import date styles, merged ranges, typed
data-validation rules, canonical typed conditional-format rules, and natively
recalculable formula caches and canonical cached dynamic-array spills;
plus Presentation slides with plain one-run text shapes and canonical basic
tables.
Headers, notes,
media, custom or non-canonical table styling, rich text, non-canonical package resources,
and every other lossy case fail with
`use.office.dump_unsupported`; nothing is silently flattened or omitted.

Replay requires the exact A3S blank template identified by `baseSha256`.
`batch` checks that precondition before mutation and checks `resultSha256`
afterward. A failed result check restores the original in-memory package before
any save. Dump files are limited to 8 MiB and 10,000 mutations, refuse to
overwrite an existing path, and use a 1 MiB inline-output limit. This is a
portable Office batch artifact, not RPC and not a universal action envelope.

`office native merge` opens a `.docx`, `.xlsx`, or `.pptx` template, performs a
single-pass replacement, validates the resulting OPC/semantic document, and
atomically writes a separate output. The template and output may not identify
the same file. The output is no-clobber by default; `--force` is the only way to
replace an existing destination, and it never authorizes modifying the template
in place.

Merge data must be a JSON object. Literal top-level keys take precedence over
flattened nested paths: `{"user.name":"literal","user":{"name":"nested"}}`
resolves `{{user.name}}` to `literal`. Nested objects use dot paths and arrays
use bracket paths such as `{{items[0].name}}`. Replacement is deliberately
single pass, so a value containing `{{another.key}}` remains literal and is
reported as unresolved rather than recursively substituted. Results include the
replacement count, sorted used keys, sorted unresolved placeholders, and sorted
changed OOXML parts.

Word merge covers the main document, headers, footers, footnotes, endnotes, and
comments. Presentation merge covers slides and notes. Spreadsheet merge covers
inline strings, direct `t="str"` values, and referenced shared rich strings;
shared-string replacements are counted per referencing cell and phonetic runs
are left untouched. A resolved placeholder in a numeric, boolean, error, or
otherwise unsupported cell fails closed instead of coercing the cell type.

Data files are limited to 8 MiB and must be regular, non-symlink files. Flattened
data is additionally bounded by entry count, nesting depth, key length, and
total bytes. XML-forbidden replacement characters fail the whole in-memory
transaction before any output is created. This native path does not invoke
OfficeCLI, Microsoft Office, LibreOffice, Python, Node.js, or .NET.

General text replacement is separate from template merge. `replace-text` uses
an explicit semantic path and either case-sensitive, non-overlapping literal
matching or a linear-time Rust regular expression. Word `/` covers the main
document plus headers, footers, footnotes, endnotes, and comments; narrower
body, header/footer, paragraph, run, table, cell, hyperlink, and comment paths
stay within their source part. Spreadsheet accepts `/`, a worksheet, or one
cell/rectangular range and edits only string cells. A scoped shared-string edit
clones the rich shared-string item and redirects selected cells when other
cells still reference the original. Presentation accepts `/`,
slide/object/text paths, and `/slide[N]/notes`; slide scopes do not implicitly
include notes. Phonetic Spreadsheet text is never changed.

One operation accepts at most 64 KiB of find expression, 1 MiB of replacement
text, 100,000 semantic matches, 64 MiB of expanded replacement text, and a
100,000-cell Spreadsheet scope. Regex matches must consume text. Results report
`matchCount`, `changed`, and sorted `changedParts`; zero matches are a
successful unchanged result. Batch results add these receipts under
`textReplacements`. All replacements are single pass, preserve split-run
ownership by assigning new text to the first matched run, retain unknown XML,
support strict and transitional OOXML, and participate in normal batch rollback
and post-mutation validation.

Raw replacement is also available inside the same atomic batch:

```json
{
  "operation": "replace-xml-part",
  "part": "/word/document.xml",
  "xml": "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/></w:body></w:document>"
}
```

Only existing XML parts may be replaced. `[Content_Types].xml`, root and
part-level relationship files, binary parts, DTDs, external entities, non-UTF-8
input, and replacement roots with a different local name or namespace are
rejected. Direct `raw-set` input is limited to 8 MiB. Inline `raw` output is
limited to 1 MiB; use its `--output` option to export larger original part bytes
without modifying the Office package. Raw export refuses to overwrite an
existing destination. Every replacement runs through the normal semantic and
OPC post-mutation validation, and any failure rolls back the whole batch.

Typed part creation is also batchable:

```json
{
  "operation": "add-part",
  "parent": "/slide[1]",
  "type": "chart"
}
```

The batch result keeps the existing ordered `paths` ledger, adds `swaps`
receipts containing the post-mutation `first` and `second` paths, and adds
`createdParts` receipts containing `part`, `ownerPart`, `relationshipId`, and
`type`. Text replacement receipts are reported separately under
`textReplacements` so a successful zero-match operation remains distinguishable
from a content change. Word supports chart, header, and footer carriers at `/`;
Spreadsheet supports chart carriers under a worksheet; Presentation supports
chart carriers under a slide. A carrier is a valid blank XML part with content
type and owner relationship. It is not visible in document layout until a typed
operation or explicit XML replacement references the returned relationship ID
from the owner XML.

Typed formatting is batchable through the same public mutation contract:

```json
{
  "operation": "set-text-format",
  "path": "/body/p[1]/r[1]",
  "format": {
    "bold": true,
    "underline": "double",
    "script": "superscript",
    "strikethrough": true,
    "doubleStrikethrough": false,
    "textCase": "small-caps",
    "highlight": "yellow",
    "language": "en-US",
    "fontFamily": "Aptos",
    "fontSizeCentipoints": 1150,
    "textColor": { "red": 0, "green": 102, "blue": 204 }
  }
}
```

`fontSizeCentipoints` is an integer count of 1/100 point; the CLI accepts the
equivalent point value through `--font-size`. `underline` accepts `none`,
`single`, or `double`, and `script` accepts `baseline`, `superscript`, or
`subscript`. `strikethrough` is supported by Word and Spreadsheet;
Presentation rejects it before mutation. `textCase`, `highlight`, and
`language` apply to Word and Presentation runs. `doubleStrikethrough` applies
only to Word. The portable highlight palette is `none`, the six bright colors,
black/white, and the dark or light gray/color variants exposed by the typed
schema. An empty `format` object, invalid BCP-47 shape, unknown properties,
invalid RGB components, and unsupported target/property combinations fail the
whole batch. Spreadsheet style records are cloned and deduplicated without
replacing unrelated style children or attributes.

Spreadsheet cell presentation uses its own typed mutation rather than adding
non-text properties to `set-text-format`:

```json
{
  "operation": "set-cell-format",
  "path": "/Sheet1/A1:C3",
  "format": {
    "numberFormat": "currency",
    "fill": {
      "kind": "solid",
      "color": { "red": 255, "green": 242, "blue": 204 }
    },
    "border": {
      "left": {
        "kind": "line",
        "style": "thin",
        "color": { "red": 128, "green": 128, "blue": 128 }
      },
      "right": { "kind": "line", "style": "thin" },
      "top": { "kind": "line", "style": "thin" },
      "bottom": { "kind": "line", "style": "double" },
      "diagonalUp": false,
      "diagonalDown": false
    },
    "verticalAlignment": "center",
    "wrapText": true,
    "textRotation": 0,
    "indent": 1,
    "shrinkToFit": false,
    "readingOrder": "left-to-right"
  }
}
```

`numberFormat` accepts an explicit Excel format code or `general`, `number`,
`currency`, `accounting`, `percent`, `scientific`, `text`, `date`, `time`, or
`datetime`. Format codes are limited to 255 characters, four sections, and
balanced quotes/brackets. Fill is either `none` or one solid 24-bit RGB color.
Each border side is an explicit `none` or `line` value. Lines accept `thin`,
`medium`, `thick`, `double`, `dashed`, `dotted`, `dashDot`, `dashDotDot`,
`hair`, `mediumDashed`, `mediumDashDot`, `mediumDashDotDot`, or
`slantDashDot`, plus an optional 24-bit RGB color. The shared diagonal line is
controlled independently from `diagonalUp` and `diagonalDown`. Vertical
alignment accepts `top`, `center`, `bottom`, `justify`, or
`distributed`; rotation accepts 0–180 or 255 for stacked text; indentation is
0–255; and reading order is `context`, `left-to-right`, or `right-to-left`.
Unknown fields, empty format objects, invalid values, and non-Spreadsheet
targets fail the whole batch. Native semantic HTML/SVG previews expose the
observed values as inert `data-*` attributes; they do not claim Excel layout
fidelity.

Spreadsheet merged cells are batchable through two closed mutations:

```json
{
  "operation": "merge-cells",
  "path": "/Sheet1/A1:C1"
}
```

Use `unmerge-cells` with the exact same path to remove the merge. Range order
and case are normalized. An exact repeated merge and an absent exact unmerge
are unchanged successes. A partial overlap, a range intersecting a Spreadsheet
table, or any unmerge range that intersects but does not exactly equal an
existing merge fails the complete batch. The latter error reports
`validRanges`; callers must unmerge each exact range rather than request a
destructive sweep. Strict and
transitional OOXML are retained, unknown `mergeCells` data is preserved, and
removing the final merge fails closed if deleting its container would discard
unknown attributes or children. Semantic cell reads expose `merge` and
`mergeAnchor`; range reads expose `merge=true|false`; and `mergeCell` queries
return stable nodes. HTML/SVG carry the same facts only as inert attributes.

Hyperlinks use the same typed batch contract and remain inert data:

```json
{
  "operation": "set-hyperlink",
  "path": "/body/p[1]",
  "hyperlink": {
    "target": {
      "kind": "external",
      "uri": "https://example.com/report"
    },
    "display": "Open report",
    "tooltip": "A3S report"
  }
}
```

Use `{ "kind": "internal", "location": "section_1" }` for a Word
bookmark, `Sheet1!B2` for a Spreadsheet location, or `/slide[2]` for a
Presentation slide jump. Word accepts body, header, and footer paragraph paths
when adding and the returned hyperlink path when updating. Spreadsheet accepts
one cell or a bounded rectangular range; only a missing single cell is created,
and a range link does not rewrite cell contents. Presentation accepts a shape
or its hyperlink path and does not accept separate display text. Removing a
hyperlink or its owning paragraph, cell/range, shape, or slide garbage-collects
only an unused hyperlink or slide relationship. Strict and transitional OOXML
namespaces are preserved.

Legacy comments use their own typed batch variants rather than generic
properties:

```json
{
  "operation": "add-comment",
  "parent": "/slide[1]",
  "comment": {
    "author": "Alice",
    "text": "Review this slide",
    "initials": "AL",
    "position": { "xEmu": 914400, "yEmu": 457200 }
  }
}
```

Use `set-comment` with a partial `update` object and the stable returned path.
Word accepts author, initials, and plain text on main-document comments;
Spreadsheet accepts author and plain text for classic cell notes; Presentation
also accepts a complete `position`. Removing a comment uses the ordinary typed
`remove` mutation. Unknown OOXML attributes and extension nodes survive these
updates, and strict/transitional relationship and root dialects are retained.
Modern threaded comments and replies, resolved state, writable comment dates,
rich comment bodies, Word header/footer anchors, and Spreadsheet threaded
comments are intentionally not represented yet.

Typed Spreadsheet content values use an explicit nested type, for example:

```json
{
  "operation": "set-cell-value",
  "path": "/Sheet1/C1",
  "value": {
    "type": "formula",
    "expression": "SUM(A1:B1)"
  }
}
```

Formula mutation removes one optional leading `=`, parses the bounded body into
a source-spanned typed AST, stores the original normalized OOXML formula text,
and marks the workbook for recalculation. It does not implicitly calculate the
workbook. The parser covers scalar and error literals, Excel operator
precedence, function calls and omitted arguments, parentheses and array
constants, names and structured references, A1 cell/row/column references with
quoted, 3D, or external qualifiers, and range/intersection/union operators.
Invalid syntax returns `use.office.spreadsheet_formula_invalid` with zero-based
UTF-8 byte and character offsets before any package mutation. Structural edits
continue to rewrite supported A1 references without evaluating formulas.

`NativeOfficeDocument::formula_dependency_graph` and
`calculate_spreadsheet_formulas` provide read-only graph and calculation
results. `NativeOfficeEditor::recalculate_spreadsheet_formulas`, the
`recalculate-spreadsheet-formulas` batch/MCP mutation, exact replay, and the
CLI command above atomically persist supported cached values and spills. Excel
formula breadth beyond the closed native registry, disjoint or non-canonical
structured-reference forms, qualified functions, external-workbook
calculation, and full cross-application conformance remain rich-Spreadsheet
delivery gates.

The native package, semantic, and editor APIs are available directly to Rust
callers:

```rust,no_run
use a3s_office::{
    NativeOfficeComment, NativeOfficeCommentUpdate, NativeOfficeDocument,
    NativeOfficeEditor, NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment,
    NativeOfficeHyperlink,
    NativeOfficeInsertPosition, NativeOfficePackage, NativeOfficeRenderFormat,
    NativeOfficeReplayArtifact, NativeOfficeRgbColor, NativeOfficeTextCase,
    NativeOfficeTextFormat, NativeOfficeTextScript, NativeOfficeUnderline,
    NativeSpreadsheetBorder, NativeSpreadsheetBorderLine, NativeSpreadsheetBorderStyle,
    NativeSpreadsheetCellFormat, NativeSpreadsheetFill,
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
    NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope,
    NativeSpreadsheetReadingOrder, NativeSpreadsheetSort,
    NativeSpreadsheetSortKey, NativeSpreadsheetTable,
    NativeSpreadsheetTableStyle, NativeSpreadsheetVerticalAlignment,
};

# async fn inspect() -> Result<(), Box<dyn std::error::Error>> {
let mut package = NativeOfficePackage::open("report.docx").await?;
let document_xml = package.part("/word/document.xml")?;
println!("{} bytes", document_xml.len());

// Format engines will use this API without dropping unknown OOXML parts.
package.save().await?;

let blank = NativeOfficePackage::create("blank.xlsx").await?;
println!("created {:?}", blank.kind());

let mut workbook = NativeOfficeEditor::create("styled.xlsx").await?;
workbook.set_cell_format(
    "/Sheet1/A1:C3",
    NativeSpreadsheetCellFormat {
        number_format: Some("currency".into()),
        fill: Some(NativeSpreadsheetFill::Solid {
            color: NativeOfficeRgbColor::new(0xFF, 0xF2, 0xCC),
        }),
        border: Some(NativeSpreadsheetBorder {
            bottom: Some(NativeSpreadsheetBorderLine::Line {
                style: NativeSpreadsheetBorderStyle::Double,
                color: Some(NativeOfficeRgbColor::new(0x80, 0x80, 0x80)),
            }),
            ..NativeSpreadsheetBorder::default()
        }),
        vertical_alignment: Some(NativeSpreadsheetVerticalAlignment::Center),
        wrap_text: Some(true),
        reading_order: Some(NativeSpreadsheetReadingOrder::LeftToRight),
        ..NativeSpreadsheetCellFormat::default()
    },
)?;
workbook.merge_cells("/Sheet1/A1:C3")?;
let validation_path = workbook.add_data_validation(
    "/Sheet1",
    NativeSpreadsheetDataValidation::new(
        NativeSpreadsheetDataValidationType::Whole,
        "D2:D50",
        "18",
    )
    .with_operator(NativeSpreadsheetDataValidationOperator::Between)
    .with_formula2("120")
    .with_allow_blank(false)
    .with_error_message(
        NativeSpreadsheetDataValidationErrorStyle::Warning,
        "Age outside range",
        "Enter an age from 18 through 120",
    ),
)?;
println!("created {validation_path}");
let named_range_path = workbook.add_named_range(
    NativeSpreadsheetNamedRange::new("Revenue", "'Sheet1'!$A$2:$A$20")
        .with_scope(NativeSpreadsheetNamedRangeScope::Workbook)
        .with_comment("Workbook revenue"),
)?;
println!("created {named_range_path}");
let table_path = workbook.add_spreadsheet_table(
    "/Sheet1",
    NativeSpreadsheetTable::new("Sales", "G1:I4", ["Name", "Qty", "Price"])
        .with_style(NativeSpreadsheetTableStyle::Medium { number: 4 }),
)?;
println!("created {table_path}");
let sort_path = workbook.sort_spreadsheet_range(
    "/Sheet1/G1:I4",
    NativeSpreadsheetSort::new(vec![
        NativeSpreadsheetSortKey::descending("H"),
        NativeSpreadsheetSortKey::ascending("G"),
    ])
    .with_header(true),
)?;
println!("sorted records at {sort_path}");
workbook.save().await?;

let document = NativeOfficeDocument::open("report.docx").await?;
println!("{}", document.text_view().text);
let html = document.render(NativeOfficeRenderFormat::Html)?;
println!("{} {} bytes", html.sha256, html.byte_length);
let headings = document.query("p[style=Heading1]")?;
println!("{} heading(s)", headings.len());

let mut editor = NativeOfficeEditor::open("report.docx").await?;
let raw = editor.raw_xml_part("/word/document.xml")?;
println!("{} {}", raw.part, raw.sha256);
let header = editor.add_part("/", a3s_office::NativeOfficePartType::Header)?;
println!("{} {}", header.part, header.relationship_id);
editor.set_text("/body/p[1]", "Updated")?;
editor.set_text_format(
    "/body/p[1]/r[1]",
    NativeOfficeTextFormat {
        bold: Some(true),
        underline: Some(NativeOfficeUnderline::Double),
        script: Some(NativeOfficeTextScript::Superscript),
        strikethrough: Some(true),
        double_strikethrough: Some(false),
        text_case: Some(NativeOfficeTextCase::SmallCaps),
        highlight: Some(NativeOfficeHighlightColor::Yellow),
        language: Some("en-US".into()),
        font_family: Some("Aptos".into()),
        font_size_centipoints: Some(1400),
        text_color: Some(NativeOfficeRgbColor::new(0x12, 0x34, 0x56)),
        ..NativeOfficeTextFormat::default()
    },
)?;
editor.set_text_format(
    "/body/p[1]",
    NativeOfficeTextFormat {
        alignment: Some(NativeOfficeHorizontalAlignment::Center),
        ..NativeOfficeTextFormat::default()
    },
)?;
let hyperlink = NativeOfficeHyperlink::external("https://example.com/report")?
    .with_display("Open report")
    .with_tooltip("A3S report");
let hyperlink_path = editor.set_hyperlink("/body/p[1]", hyperlink)?;
println!("created {hyperlink_path}");
let comment_path = editor.add_comment(
    "/body/p[1]",
    NativeOfficeComment::new("Alice", "Please review this paragraph")?
        .with_initials("AL"),
)?;
editor.set_comment(
    &comment_path,
    NativeOfficeCommentUpdate {
        text: Some("Reviewed".into()),
        ..NativeOfficeCommentUpdate::default()
    },
)?;
let added = editor.add_paragraph("/body", "Summary")?;
let moved = editor.move_node(
    added,
    None,
    Some(NativeOfficeInsertPosition::at_index(0)),
)?;
let copied = editor.copy_node(&moved, None, None, None)?;
let swapped = editor.swap_nodes(moved, copied)?;
editor.remove(swapped.second)?;
let table = editor.add_table("/body", 2, 3)?;
editor.set_text(format!("{table}/tr[1]/tc[1]"), "Name")?;
editor.save().await?;

let mut template = NativeOfficeEditor::open("template.docx").await?;
let merge = template.merge_template(&serde_json::json!({
    "customer": {"name": "A3S Lab"}
}))?;
println!("{} replacement(s)", merge.replaced_count);
template.save_as_new("merged.docx").await?;

let replay = NativeOfficeReplayArtifact::dump(&editor.snapshot()?, "/")?;
let mut restored = NativeOfficeEditor::create("restored.docx").await?;
restored.apply_replay(&replay)?;
restored.save().await?;
# Ok(())
# }
```

Managed compatibility installation accepts only approved HTTPS release
origins, bounds the download, verifies the publisher SHA-256, stages outside
the active version, and activates atomically. Compatibility execution sets
`OFFICECLI_SKIP_UPDATE=1` so upgrades remain explicit A3S operations.

The native engine does not copy OfficeCLI's private resident protocol. The
explicit `office-native` target now provides typed in-process sessions over its
own standard MCP surface. Until the default-route promotion, a lost
compatibility response can return
`use.office.outcome_unknown`; callers must not retry it automatically.

See [Native Office Engine](native-office-engine.md) for the complete
requirements, compatibility scope, safety invariants, delivery gates, and
migration plan.
