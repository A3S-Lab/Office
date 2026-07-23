# Native Office Engine

## Decision

A3S Use will own its Office document engine. The steady-state runtime must read,
create, modify, validate, and render `.docx`, `.xlsx`, and `.pptx` without an
OfficeCLI, Microsoft Office, LibreOffice, Python, Node.js, or .NET dependency.

LibreOffice is not a fallback engine, renderer, formula evaluator, conversion
service, or installation prerequisite. CI may invoke it only as an optional
external interoperability oracle; shipped binaries and normal tests must work
without it.

The `a3s use office` route, typed Rust API, standard MCP server, and packaged
Skill remain stable product surfaces. OfficeCLI `1.0.136` is a temporary
compatibility backend and a black-box behavior baseline while native coverage
is built. It is not the target architecture.

The native implementation is written in Rust from the OOXML/OPC specifications
and public compatibility behavior. Do not mechanically translate upstream C#
source. Synthetic fixtures and differential black-box tests may be used to
measure compatibility.

## Product scope

The native engine must provide these common capabilities across all three
formats:

- safe OPC/ZIP package loading and atomic persistence;
- loss-preserving round trips for unknown parts, relationships, elements, and
  attributes;
- create, semantic view, get, query, set, add, remove, move, copy, and swap;
- bounded, scoped literal and regular-expression text replacement that
  preserves rich-text run ownership;
- typed bold, italic, underline, vertical-script, strikethrough where supported,
  portable highlight, text case, primary language, font-family, exact-size, RGB
  text-color, and alignment mutation without generic property maps;
- typed inert hyperlinks with format-specific external and internal targets;
- typed legacy comments with format-specific anchors, authors, and positions;
- typed selectors with stable, one-based document paths;
- transactional batch mutation and explicit partial-apply compatibility mode;
- template merge and replayable dump/batch documents;
- raw part access, constrained XML mutation, part creation, and validation;
- open, save, close, revision tracking, and conflict detection;
- text, bounded annotated, outline, statistics, bounded typed issue, HTML, SVG,
  and screenshot views;
- standard MCP tools and a first-party Office Skill backed by the same typed engine.

The compatibility command baseline is OfficeCLI `1.0.136`, commit
`4ba79f0b984e141f57f58d4398ba2df29e8187e8`. Product parity covers document
operations, not OfficeCLI's installer, updater, agent-configuration writers, or
private resident-pipe protocol.

### Word

Word coverage includes paragraphs, runs, styles, numbering, tables, sections,
headers and footers, hyperlinks, bookmarks, images, shapes, text boxes, fields,
TOC, equations, comments, footnotes and endnotes, content controls, form fields,
revisions, charts, OLE preservation, document properties, international fonts,
BCP-47 language tags, and RTL layout.

### Spreadsheet

Spreadsheet coverage includes workbooks, worksheets, cells, ranges, row and
column mutation, formulas, defined names, styles and number formats, merged
cells, tables,
sorting, filters, validation, conditional formatting, hyperlinks, drawings,
images, charts, pivot tables and caches, slicers, sparklines, comments, OLE
preservation, and CSV/TSV import.

The formula subsystem uses a bounded parser, deterministic dependency graph,
native recalculation engine, dynamic-array spilling, reference rewriting, and
a typed closed function registry. Formula values are never evaluated by a
shell or general-purpose script runtime.

### Presentation

Presentation coverage includes slides, masters, layouts, themes, placeholders,
shapes, text, groups, connectors, tables, images, charts, audio and video, OLE
preservation, 3D model parts, equations, diagrams, notes, comments, animations,
transitions, hyperlink actions, morph metadata, and slide zoom.

## Architecture

```text
CLI / standard MCP / Skill / typed Rust API
                    |
             Office command layer
                    |
       session + transaction coordinator
                    |
   selector / semantic model / validation
        /              |              \
     Word          Spreadsheet     Presentation
                        |
             formula + pivot engines
        \              |              /
       relationships / media / charts / themes
                    |
        loss-preserving XML part model
                    |
         safe OPC/ZIP package kernel
                    |
              .docx/.xlsx/.pptx
```

The package kernel owns archive safety, canonical part names, document-kind
identification, bounded memory admission, unknown-part preservation, and atomic
save. It does not contain format-specific selectors or mutations.

Format engines expose typed operations. CLI and MCP payloads are boundary
adapters and must parse into those operations; they are not the domain model.
There is no cross-domain action envelope and no A3S JSON-RPC protocol.

Rendering is separate from document mutation. Semantic HTML/SVG renderers
produce deterministic artifacts inside `a3s-use-office`, which remains
browser-independent. The root `a3s-use` facade implements screenshots by
injecting the existing `a3s-use-browser` `PageRenderer` contract instead of
embedding a second browser runtime in the Office engine.

## Safety and fidelity invariants

1. Opening a document never downloads a relationship target or executes a
   macro, OLE payload, formula, field, or embedded script.
2. Archive entry count, archive bytes, expanded bytes, part bytes, and
   compression ratio are bounded before semantic parsing.
3. Absolute, traversal, control-character, symbolic-link, encrypted, and
   case-ambiguous part names are rejected.
4. DTDs and external XML entities are rejected. Namespace prefixes, unknown
   attributes, `mc:AlternateContent`, and untouched parts survive round trips.
5. External relationships are data. Native hyperlink writes accept only
   absolute HTTP, HTTPS, or mailto URIs without embedded credentials. Opening,
   inspecting, and rendering never fetch them; network access requires a
   separate, explicit policy and remains disabled by default.
6. Mutation batches are atomic by default. A compatibility caller must opt in
   to continue-on-error behavior and receives an applied-operation ledger.
7. Save writes a synchronized temporary package in the destination directory
   and atomically replaces the target. A changed source revision causes a
   conflict instead of silently overwriting another writer.
8. Ambiguous mutation outcomes retain `use.office.outcome_unknown` and are
   never retried automatically.
9. Replay dump never emits a lossy approximation. A dump is accepted only when
   replaying its typed mutations from the recorded blank-template fingerprint
   reproduces the complete uncompressed OOXML part map byte-for-byte.
10. Template merge never modifies its template in place. It validates all
    replacements transactionally and creates a distinct no-clobber output unless
    the caller explicitly authorizes destination replacement with `--force`.
11. Render output is bounded while it is composed, contains no source path or
    timestamp, never fetches an external relationship, and never emits document
    text as executable markup, style, or script.
12. General find/replace is single pass and path-scoped. Zero matches are
    reported as an unchanged success; a scoped Spreadsheet replacement never
    mutates cells outside the requested worksheet or A1 range through a shared
    string alias.

## Delivery gates

Native Office is promoted by evidence, not by the presence of command names.

### Gate 0 — Package kernel

- Detect Word, Spreadsheet, and Presentation packages.
- Enforce archive and part safety limits.
- Preserve every unknown part byte-for-byte while a known part changes.
- Save atomically and remain `Send + Sync`.

Status: implemented in the native engine; it is a public foundation API and is
not yet the default CLI provider.

### Gate 1 — Native read

- Loss-preserving XML and relationship graph.
- Common selector parser and stable paths.
- Text, bounded annotated, outline, statistics, bounded issues, get, and query
  for all three formats.
- Fixtures open without repair in Microsoft Office and LibreOffice.

Status: implementation in progress. Loss-preserving XML, UTF-8/UTF-16 safety,
content types, relationship resolution, common selectors, stable `get` paths,
and semantic text/annotated/outline/statistics reads are implemented for all
three formats. Annotated reads return at most 200 entries by default and 1,000
when explicitly requested; individual text and format fields are also bounded
and truncated on UTF-8 boundaries. The explicit `a3s-office` CLI
exercises them without an external provider. Gate 1 remains unpromoted until
the cross-application repair-dialog corpus passes and remaining rich read nodes
are covered.

### Gate 2 — Native basic mutation

- Create documents and perform core text, table, cell, sheet, slide, shape,
  hyperlink, and image mutations.
- Atomic batch, save conflict detection, raw access, validation, merge, and dump.
- Untouched-part and untouched-subtree fidelity gates.

Status: implementation in progress. Native creation now writes minimal Word,
Spreadsheet, and Presentation packages selected by extension with atomic,
no-clobber persistence. Native `set-text` edits existing Word
paragraph/run/cell content and Presentation shape text; Spreadsheet cell paths
upsert missing ordered rows and cells and maintain worksheet dimensions.
Spreadsheet writes preserve explicit text, finite-number, boolean, and formula
types. Formula writes strip an optional leading `=`, enforce Excel's 8192
character bound, parse the normalized body into a bounded source-spanned typed
AST, and mark the workbook for recalculation. Formula writes and import do not
implicitly calculate the workbook. The parser covers literals, Excel operator
precedence, calls and omitted arguments, parentheses and array constants, names
and structured references, qualified A1 cell/row/column references, and
range/intersection/union operators. Syntax errors report stable zero-based
UTF-8 byte and character offsets before mutation.

The native calculation subsystem builds a deterministic graph across
worksheets, ranges, spills, and workbook- or worksheet-scoped names.
`NativeOfficeDocument::formula_dependency_graph`,
`calculate_spreadsheet_formulas`, and the registry-aware calculation variant
are read-only. The closed default registry implements `SUM`, `AVERAGE`, `MIN`,
`MAX`, `COUNT`, `COUNTA`, `ABS`, `SQRT`, `POWER`, `MOD`, `ROUND`, `IF`,
`IFERROR`, `AND`, `OR`, `NOT`, `CONCAT`, `CONCATENATE`, `ROW`, `COLUMN`,
`SEQUENCE`, `TRANSPOSE`, `PI`, and `NA`, together with ordinary operators,
typed errors, scalar/array broadcasting, spill references, and dynamic-array
results.

`NativeOfficeEditor::recalculate_spreadsheet_formulas` and
`NativeOfficeMutation::RecalculateSpreadsheetFormulas` atomically calculate
and write typed cached values, canonical array anchors, spill children, and
calculated-workbook metadata. The mutation is exposed by replay, standard MCP
as `recalculate-spreadsheet-formulas`, and CLI as
`office native recalculate <workbook> [--output <workbook>]`. Spill children
are read-only; replacing or removing an anchor clears its old spill. Failures
including cycles, unsupported or qualified functions, unsupported
structured-reference forms, external-workbook references, overlapping formula
storage, and invalid OOXML roll back the complete batch. Spreadsheet errors
such as `#DIV/0!` and `#SPILL!` remain typed cell values. The engine never
fetches external workbooks or invokes a scripting fallback.

ListObject structured references resolve table names or display names.
`Table[Column]` selects one data column and `Table[[First]:[Last]]` selects a
contiguous data-column range. `#All`, `#Data`, `#Headers`, and `#Totals` select
structural rows. `Table[@Column]`, `Table[[#This Row],[Column]]`, and
table-local `[@Column]` select the current data row; table-local forms require
the formula cell to be inside the inferred table. Missing tables, columns, or
requested header/totals rows, disjoint columns, and non-canonical forms remain
fail-closed.

Exact replay accepts canonical formula storage and canonical array anchors
only when their native cached results are present. It rejects physical storage
that typed mutations cannot reproduce byte-for-byte, including explicit
`t="normal"` formulas and uncached or malformed array anchors, with
`use.office.dump_unsupported`.

Semantic Spreadsheet cell nodes expose `formulaCached` for formula anchors and
`valuePresent` for physical cached or inline values. This distinguishes an
explicit empty string from an absent cell value and lets issue analysis report
only formulas that truly lack an OOXML cache.

Formula calculation is bounded to 8,192 formula characters, depth 128, and
8,192 AST nodes, with the same depth bound applied across nested named
references; 100,000 reference areas per value; 100,000 formula cells,
1,000,000 graph edges, and 1,000,000 formula-cell reference visits; 100,000
materialized cells in one array or function call and 100,000 cumulative spill
children in one pass; 200,000 OOXML cell writes; and 1 MiB for one UTF-8 text
result. Formula text is additionally limited to 8 MiB cumulatively per
calculation pass. Cell set/remove accepts normalized A1 rectangular ranges of
at most 100,000 cells and rolls back the whole operation on error.
Native Spreadsheet structure operations insert or delete at most 10,000 rows or
columns, rename worksheets, reorder worksheets, and copy a worksheet after its
source or at an explicit one-based position. Copy assigns new worksheet and
table identities, clones owned table/comment/VML/drawing/chart/media OPC
subgraphs, preserves explicitly shared workbook resources, rewrites copied
self-references and local defined names, and leaves every source part unchanged.
Worksheet removal deletes only unshared descendants, removes local definitions,
shifts workbook indexes, and changes surviving formulas that target the deleted
sheet to `#REF!`. Formula and structural mutations discard stale calculation
chains and request a full application recalculation.

Native Spreadsheet physical sorting is also implemented as one stable ordered
multi-key mutation with persisted worksheet sort state. It is deliberately
separate from worksheet/table filter-definition lifecycle and from formula
calculation.

Native bounded CSV/TSV import is implemented through typed Rust, versioned
batch, dedicated CLI, and standard MCP surfaces. One import targets an existing
worksheet at an explicit A1 start cell, accepts at most 8 MiB of UTF-8 and a
100,000-cell rectangular extent, and parses BOM, CRLF, quoted delimiters,
embedded newlines, and doubled quotes without a scripting runtime. Unclosed,
misplaced, or trailing quote content fails the whole transaction instead of
being guessed. Ragged missing fields preserve cells beyond that source row;
explicit empty fields clear an existing target cell without materializing a new
blank cell.

Import inference writes formulas, finite numbers, booleans, ISO dates/times, or
text as typed cell values. Dates honor the workbook's 1900/1904 date system and
receive a canonical date number format. Formulas pass the bounded cell-formula
parser, are stored, and are marked for recalculation. Import does not calculate
them implicitly; callers can append the native recalculation mutation to the
same atomic batch. Header mode atomically adds or replaces the worksheet
AutoFilter over the imported extent and installs one canonical frozen pane
below the header. Frozen panes have typed set/remove, semantic `/Sheet/freeze`
readback, selectors, versioned batch and MCP payloads, and exact replay.
Unsupported or vendor-extended pane content is readable with
`nativeMutable=false` and fails closed on mutation. Tests cover
strict/transitional SpreadsheetML, malformed input rollback, typed values,
explicit empty cells, filter/freeze lifecycle, exact replay, file/stdin CLI,
and an unsaved/save/close/reopen standard MCP session without OfficeCLI.

Native `replace-text` is implemented through one typed Rust, batch, CLI, and
standard MCP contract. Literal mode performs case-sensitive, non-overlapping
substring matching. Regex mode uses Rust's linear-time regular-expression
engine and expands `$1` and `$name` captures; matches that consume no text are
rejected. Replacement is single pass and may span multiple OOXML text runs.
The inserted value belongs to the first matched run, unmatched text retains its
original run, and later covered runs are emptied without deleting their
formatting or unknown extension content.

Word `/` covers `document.xml`, referenced or unreferenced header/footer text
parts, footnotes, endnotes, and legacy comments. Narrower body, header/footer,
paragraph, run, table/row/cell, hyperlink, and comment scopes remain inside one
source part. Spreadsheet accepts `/`, one worksheet, one cell, or a normalized
rectangular range and changes only shared, inline, or direct string values.
When selected and unselected cells reference the same shared rich string, the
engine clones the complete `si` item, preserves unknown children and rich runs,
redirects only selected cells, and leaves phonetic `rPh` text unchanged.
Presentation accepts `/`, a slide or supported object/text descendant, and a
related `/slide[N]/notes` path. A slide scope excludes speaker notes unless the
notes path or root is selected.

Find expressions are limited to 64 KiB, replacement input to 1 MiB, semantic
matches to 100,000, expanded replacement output to 64 MiB, and Spreadsheet
range or observed-scope cells to 100,000. Receipts contain the scope,
literal/regex mode, `matchCount`, `changed`, and sorted `changedParts`; batch
receipts are additive under `textReplacements`. Zero matches do not dirty an
in-place document. XML-forbidden output, invalid regex, limit overflow, path or
cell-type errors, and failed post-mutation validation roll back the complete
batch. Tests cover strict and transitional Word, Spreadsheet, and Presentation,
split runs, unknown XML, partial shared-string aliases, notes, zero matches,
rollback, the CLI on all formats, and a complete unsaved/save/close standard
MCP lifecycle with an unusable OfficeCLI provider path.

Native `set-text-format` is implemented through one typed Rust, batch, CLI,
and standard MCP contract. It supports explicit bold and italic state,
`none`/single/double underline, baseline/superscript/subscript, font family,
integer centipoint size, 24-bit RGB text color, and
left/center/right/justify alignment. Word and Spreadsheet additionally support
an explicit single-strikethrough boolean. Presentation rejects
`strikethrough` with `use.office.presentation_strikethrough_unsupported` before
changing a package. Word and Presentation additionally share display-only
`none`/small-caps/all-caps, a closed 17-color highlight palette, and a
conservatively validated BCP-47 primary-language tag. Word alone supports
double strikethrough; Presentation and Spreadsheet reject that format rather
than ignoring it. Word and Presentation character properties target semantic
run paths; alignment targets their paragraph paths. Word accepts only sizes
divisible by 50 centipoints because WordprocessingML stores half-points.
Spreadsheet accepts single cells and bounded rectangular ranges, auto-vivifies
empty styled cells, creates `xl/styles.xml` and its workbook relationship when
absent, clones the base font/cell XF without dropping unknown properties, and
deduplicates derived font and style records. Transitional and strict OOXML
namespaces are retained. All three formats provide normalized semantic
readback and use normal batch rollback and post-mutation validation. Tests
cover explicit off/baseline/none values, the complete portable highlight
mapping, invalid language and format combinations, unknown style attributes,
strict OOXML, CLI execution without OfficeCLI, and a complete standard MCP
session. Advanced named styles, inheritance, arbitrary or scheme highlights,
extended underline variants, per-script Word languages, character spacing,
and arbitrary property maps remain outside this text-format milestone.
Spreadsheet borders use the separate cell-format contract below.

Native `set-cell-format` is a separate typed Rust, batch, CLI, and standard MCP
contract for Spreadsheet presentation properties that are not text-run
formatting. One cell or a bounded rectangular range accepts an Excel number
format, explicit solid RGB fill or fill removal, cardinal and diagonal borders,
vertical alignment, wrap-text state, text rotation, indentation, shrink-to-fit
state, and reading order. The CLI can compose these properties with one typed
content write, text formatting, and a hyperlink in the editor's existing
atomic batch; a failure rolls back every mutation before save. Word and
Presentation reject this mutation with
`use.office.mutation_type_unsupported` rather than ignoring its properties.

Number formats accept an explicit code or the normalized aliases `general`,
`number`, `currency`, `accounting`, `percent`, `scientific`, `text`, `date`,
`time`, and `datetime`. Codes are limited to 255 Unicode scalar values, at most
four semicolon-separated sections, balanced quotes and square brackets, and
XML-safe characters. The style writer reuses built-in IDs where possible,
deduplicates custom `numFmt` records, maintains the collection count, and
retains unrelated records. Fill is a closed `none` or solid 24-bit RGB value.
Borders provide explicit left, right, top, bottom, and shared diagonal line
updates. A line is either `none` or one of all 13 SpreadsheetML line styles,
with an optional 24-bit RGB color; diagonal-up and diagonal-down flags are
independent explicit booleans. The CLI provides `--border-all`/`--border` and
`--border-color` cardinal shorthands plus per-side style/color flags. A
per-side flag overrides the all-sides shorthand, and an explicit `none` clears
that line. Vertical alignment is `top`, `center`, `bottom`, `justify`, or `distributed`;
rotation is 0 through 180 or 255 for stacked text; indentation is 0 through
255; and reading order is contextual, left-to-right, or right-to-left.

The writer clones and deduplicates `fills`, `borders`, and `cellXfs`, sets only
the relevant apply flags, and merges alignment or border updates without
dropping unknown XF, alignment, border, color, extension, start/end, or
vertical/horizontal-border data. A line update is a complete typed value, so a
`line` without `color` clears the prior explicit color while omitted sides are
preserved. Explicit false, zero, contextual reading order, fill removal, and
border removal are intentional changes. Strict and transitional OOXML dialects
continue through the existing loss-preserving style path. Semantic reads
normalize observed border styles, RGB colors, and diagonal flags alongside the
other properties, and HTML/SVG semantic previews expose them as inert `data-*`
attributes. Tests cover stable JSON, every native line style, range writes,
style/number/fill/border deduplication, explicit clearing, unknown style and
border preservation, invalid-value rollback, standard MCP schema conversion,
native CLI and batch execution, wrong-document rejection, and an unusable
OfficeCLI provider. Gradient/pattern/theme fills, named styles, locale-derived
formats, and Excel layout fidelity remain outside this cell-format milestone;
conditional formatting uses the separate typed contract below.

Native `add-data-validation` and `set-data-validation` form a closed typed
Rust, versioned batch, CLI, and standard MCP contract. Ordinary typed `remove`
deletes one rule through its stable `/SheetName/dataValidation[N]` path. The
CLI adds a rule beneath a worksheet with `--type data-validation`, repeated
`--range` values, `--validation-type`, and `--formula1`; `set` replaces the
complete domain value while its CLI adapter first carries forward unspecified
fields from the existing semantic node. Data-validation options cannot be
combined with unrelated `set` content, format, hyperlink, merge, width, or
comment options.

The closed rule families are list, whole number, decimal, date, time,
text length, and custom formula. List and custom rules reject comparison
operators and `formula2`. Whole, decimal, date, time, and text-length rules
require one of eight typed comparison operators; only `between` and
`notBetween` require and accept `formula2`. Inline list text is quoted for
SpreadsheetML, while cell-range and defined-name sources remain formulas. An
optional leading `=` is removed from comparison and custom formulas. Valid ISO
dates from 1900 through 9999 become serial dates using the workbook's declared
1900 or 1904 date system; 1900 mode retains Excel's historical leap-day offset.
`HH:MM` and `HH:MM:SS` values become day fractions. Data-validation formula
predicates are stored but are not executed by this feature or by the cell
formula recalculation pass.

One rule carries typed blank, input-message, error-message, error-style, and
list-dropdown state. New A3S rules default `allowBlank`, `showInput`,
`showError`, and `inCellDropdown` to true and error style to `stop`.
`inCellDropdown=false` is valid only for list rules. Formula text is limited to
255 characters, prompt and error titles to 32, prompt text to 255, and error
text to 225; present values must be XML-safe and non-empty. A worksheet accepts
at most 65,534 rules and one rule accepts 1–1,024 normalized rectangular A1
ranges. Areas inside one rule must be disjoint, and no rule may overlap an area
owned by another rule. Any violation rolls back the complete editor batch.

Semantic reads expose each rule as a typed node, support selectors such as
`dataValidation[type=list]`, annotate observed cells with `dataValidation` and
`validationType`, and create a virtual annotated cell when a requested blank
cell is covered. Exact range reads carry the same rule metadata. HTML and SVG
surface it only through inert `data-validation` attributes. Ordered range
sweeps retain the sparse-cell model instead of expanding every covered cell.
Strict and transitional SpreadsheetML QNames are retained. Updates preserve
unknown rule attributes and unrelated collection data; an unknown rule child
blocks replacement, and removal of the final rule fails with
`use.office.spreadsheet_validation_unknown_content` if deleting its collection
would discard unknown attributes or children. Exact replay emits typed rules
after cell and merge mutations. Tests cover rule normalization, all lifecycle
surfaces, strict OOXML, unknown-data preservation, overlap rollback, sparse
semantic/HTML/SVG readback, replay, CLI atomic batches, and a complete standard
MCP lifecycle with an unusable OfficeCLI provider.

This milestone is cell data-validation structure, not complete rich
Spreadsheet or OfficeCLI parity. Table calculated columns and totals functions,
unsupported imported sort-state variants and date-group/color/icon filter
families, charts, pivot tables, slicers,
sparklines, data-validation predicate execution, and Excel layout fidelity remain
separate work.

Native `add-conditional-format` and `set-conditional-format` form a separate
closed typed Rust, versioned batch, CLI, and standard MCP contract. Ordinary
typed `remove` deletes one rule through its stable `/SheetName/cf[N]` path. The
CLI adds a rule below a worksheet with `--type conditional-format`, one
`--rule-type`, and repeated `--range` values. CLI `set` reads the current
semantic node and carries forward omitted fields; Rust, batch, and MCP set use
one complete `conditionalFormat` value. Conditional-format updates cannot be
mixed with ordinary cell values, cell/text formatting outside the differential
fill/font-color/bold subset, hyperlinks, merges, comments, data-validation-only
options, or defined-name options.

The classic families are `cellIs`, expression formula, contains/not-contains/
begins-with/ends-with text, top/bottom rank or percentage, above/below average,
duplicate/unique values, contains/not-contains blanks, contains/not-contains
errors, and ten date windows. Cell comparison uses the same eight closed
comparison operators as data validation. Between and not-between require a
second formula; every other operator rejects it. Formula bodies contain 1–8,192
characters without surrounding whitespace or a leading `=`. Text predicates
contain 1–255 XML-safe characters. Rank is 1–1,000, or 1–100 for percentages;
average rules accept zero through three standard deviations. Classic rules
carry only a differential solid RGB fill, RGB font color, and optional bold
state. This is intentionally not a generic style property bag.

The visual families are data bars, two- or three-color scales, and the standard
legacy three-, four-, and five-icon sets. Thresholds are closed values:
`min`, `max`, finite number, 0–100 percent, 0–100 percentile, or formula. A data
bar owns its RGB color, minimum and maximum thresholds, value visibility, and
optional 0–100 minimum/maximum lengths. A three-color scale requires both its
midpoint and midpoint color; omitting both creates a two-color scale. An icon
set accepts exactly its icon count in thresholds, or generates evenly spaced
percent thresholds when none are supplied. Visual rules reject differential
fill/font/bold fields instead of ignoring them. Advanced x14-only negative
data-bar colors, axes, borders, directions, and custom icon sets are not native.

One rule accepts 1–1,024 normalized, internally disjoint rectangular A1 areas,
and one worksheet accepts at most 65,534 rules. Different rules may overlap so
priority and `stopIfTrue` retain Excel semantics. A3S allocates a collision-free
priority and emits one `conditionalFormatting` carrier per new rule; canonical
A3S-created sequences are sequential. Imported carriers containing multiple
rules remain readable. Their shared
`sqref` cannot be changed through one rule because that would implicitly alter
its siblings; content replacement with the same ranges and precise rule
removal remain supported when the carrier can be retained losslessly.

Semantic reads expose a typed `ConditionalFormatting` node at
`/SheetName/cf[N]`, including normalized ranges, priority, `stopIfTrue`, rule
type, family-specific fields, differential colors/font state, and
`nativeMutable`. Query supports `conditionalFormatting` and selectors such as
`conditionalFormatting[type=iconSet]`. Unsupported or extension-only rules are
still visible with `nativeMutable=false`; mutation fails closed rather than
converting or dropping their content.

The loss-preserving writer retains strict or transitional SpreadsheetML,
unknown rule/container attributes, schema child order, and unrelated worksheet
content. It creates or deduplicates differential formats in `styles.xml` and
keeps existing unknown style data. Unknown rule children or container content
that cannot survive set/remove produce
`use.office.spreadsheet_conditional_format_unknown_content`. Canonical replay
emits typed rules with sequential priorities. Tests cover all classic families,
all four broad visual families, threshold and range validation, add/set/query/
remove/reopen, strict OOXML, unknown-content preservation, shared `sqref`,
atomic rollback, exact canonical replay, CLI lifecycle and atomic batch, MCP
schema conversion, and a complete standard MCP lifecycle with an unusable
OfficeCLI provider path.

This milestone stores conditional-format formulas but does not evaluate those
rule predicates or render Excel's visual result. It does not provide x14
advanced visual options, table/chart/pivot formatting, or full Excel rendering
and layout fidelity, and therefore is not complete OfficeCLI or Spreadsheet
parity.

Native `add-named-range` and `set-named-range` form a separate closed typed
Rust, versioned batch, CLI, and standard MCP contract for Spreadsheet defined
names. Ordinary typed `remove` deletes one name through its stable scoped path.
The CLI adds a name with `--type named-range`, `--name`, and `--ref`; `--scope`
accepts `workbook` or an existing worksheet. A worksheet parent becomes the
default local scope and automatically qualifies a bare A1 range. CLI `set`
carries forward omitted fields, while Rust, batch, and MCP replacements use one
complete value containing `name`, `ref`, `scope`, optional `comment`, and a
typed `volatile` flag.

Workbook-global and worksheet-local names are distinct identities. The latter
store the worksheet's zero-based `localSheetId`. Canonical paths include both
dimensions, for example `/namedrange[@name=Revenue][@scope=workbook]` and
`/namedrange[@name=Revenue][@scope=Sheet1]`; selector values are percent-encoded
when required. A local name on a worksheet literally named `workbook` uses the
escaped scope `worksheet:workbook`, keeping it distinct from the global scope
in JSON, CLI, MCP, semantic readback, and canonical paths. Name-only, `@name`, and one-based positional compatibility paths
remain available. If the same name exists in more than one scope, an unscoped
lookup fails with `use.office.spreadsheet_named_range_ambiguous` and requires an
explicit scope instead of choosing one silently. Semantic `get` and `query`
return the normalized name, ref, scope, comment, volatile state, and protection
metadata.

Defined-name identifiers follow Excel's 255-character Unicode letter/digit,
underscore, period, and backslash grammar and may not resemble A1 or R1C1
notation. Refs contain the formula body without a leading `=`, are bounded to
8,192 characters, and may be a qualified range, constant, or formula body.
Workbook-scoped bare A1 refs are rejected; local bare refs are sheet-qualified.
Simple sheet-qualified A1 refs must target an existing worksheet, and
cross-workbook refs that require unsupported external-link parts are rejected.
Comments are limited to 255 characters and each workbook to 65,536 names.
Identity uniqueness is case-insensitive by `(name, scope)`, and the shared
Excel namespace also rejects collisions with ListObject `name` or
`displayName`. `_xlnm.*` print/filter definitions and `Slicer_*` sentinels are
protected because their owning typed features must manage them.

Every defined-name mutation marks the workbook for recalculation but does not
calculate by itself. An explicit native recalculation pass resolves supported
names referenced by cell formulas. The loss-preserving writer keeps workbook
child order, strict or transitional SpreadsheetML QNames, untouched defined
names, and unknown attributes. Unknown collection children or non-text name
content fail closed; removing the final name also fails if deleting
`definedNames` would discard unknown collection attributes. Exact replay emits
named ranges after worksheet creation and reproduces the supported part map
byte-for-byte. Tests cover
scoped identity and ambiguity, validation and rollback, reserved names, unknown
data, strict OOXML, table-name collisions, replay, native CLI batch atomicity,
and a complete standard MCP unsaved/save/close lifecycle with an unusable
OfficeCLI provider.

This milestone is defined-name lifecycle and storage, not external-link
authoring or complete rich Spreadsheet parity.

Native `add-spreadsheet-auto-filter` and `set-spreadsheet-auto-filter` form a
closed typed Rust, versioned batch, CLI, and standard MCP contract. Ordinary
typed `remove` deletes the single worksheet filter through its stable
`/SheetName/autofilter` path. The same filter-column values live inside
ListObject table replacements and project below
`/SheetName/table[N]/autofilter`. CLI add uses `--type auto-filter`, one
`--range`, and repeated strict JSON `--filter` objects. CLI set preserves an
omitted range, replaces the full criterion list when `--filter` is present, and
requires `--clear-filters` to clear criteria explicitly. Rust, batch, and MCP
set always supply the complete filter value.

One value owns a rectangular normalized A1 range and zero or more unique
zero-based column offsets. The closed criteria are value sets with optional
blanks; equals/not-equals; contains/does-not-contain; begins/ends-with;
greater/less-than with inclusive variants; between/not-between;
blanks/non-blanks; top/bottom count or percentage; and dynamic average,
relative-day/week/month/quarter/year, year-to-date, numbered-quarter, and
numbered-month families. Value, column, percentage, count, aggregate text, and
range-width limits are enforced before mutation. Comparison and wildcard text
is XML-escaped, literal `~`, `*`, and `?` characters are preserved, and column
elements are serialized in deterministic index order.

A worksheet accepts one AutoFilter. Its range cannot intersect a ListObject
table or merged range. A table filter requires a header and uses the table
range minus an enabled totals row. Table and worksheet filter regions cannot
overlap. Semantic reads expose `AutoFilter`, `FilterColumn`, and `FilterValue`
nodes, including `/Sheet/autofilter/filterColumn[N]`, normalized refs, typed
criterion metadata, and `nativeMutable`. Exact replay reconstructs worksheet
and table filters from those nodes. The writer retains strict or transitional
SpreadsheetML and fails closed for unknown attributes or comments, extensions,
date-group items, color/icon filters, and embedded sort state. Those imported
forms remain readable but are not flattened into the supported type.

Native `sort-spreadsheet-range` is the separate closed typed Rust, versioned
batch/replay, CLI, and standard MCP contract for physical row rearrangement and
worksheet SpreadsheetML `sortState` authoring. The CLI is
`office native sort <file.xlsx> </Sheet|/Sheet/A1:D100>` with one to 64 ordered
`--key <A:XFD[:asc|desc]>` values, explicit `--header true|false`, explicit
`--case-sensitive true|false`, and optional `--output`. A worksheet path
auto-detects the smallest used cell rectangle; an explicit range is normalized
and limited to 100,000 cells. Key columns are absolute, unique, and inside the
selected range. The first key has highest precedence.

Sorting is stable: records equal across all keys keep source order. Numeric and
boolean values sort numerically before text, text is case-insensitive by
default, and blanks remain last for ascending and descending keys. The header
row stays fixed only when requested. An explicit partial-column range moves
only the selected cell fragments; cells outside the range and destination row
properties remain fixed. Sparse source records can materialize sparse
destination rows while retaining prefixed strict/transitional SpreadsheetML.

Semantic reads expose `/Sheet/sort` plus ordered `/Sheet/sort/key[N]` children,
the normalized range, header/case flags, key directions, and `nativeMutable`.
An ordinary typed `remove` of `/Sheet/sort` deletes only persisted sort
metadata and never reconstructs the old row order. Supported existing simple
sort state may be replaced or removed; imported column sorts, extended
conditions, unknown attributes/children, inconsistent condition ranges, and
other non-lossless state remain readable with `nativeMutable=false` and block
mutation.

The sort may cover an exact mutable ListObject or worksheet-AutoFilter range
with matching header behavior, or its exact data range. Partial intersection
fails closed, and a table with a totals row is rejected. Cell/table formulas
anywhere in the workbook, an intersecting merge, an owned pivot table, or a
drawing anchor that cannot follow one record losslessly also rejects the whole
atomic mutation. Hyperlinks, classic/threaded comment refs, VML note anchors,
data validations, conditional formats, protected ranges, ignored errors, and
supported one-/two-cell drawing anchors follow records. Chart caches are
cleared, the worksheet used dimension is recomputed, and the workbook is marked
for recalculation. Tests cover ordering and case behavior, sparse and
partial-column movement, strict/prefixed XML,
metadata lifecycle, exact table/AutoFilter ranges, rollback and all fail-closed
boundaries, drawing/comment/format metadata movement, replay, CLI, and a full
standard MCP unsaved/save/close/reopen lifecycle.

Native `add-spreadsheet-table` and `set-spreadsheet-table` form a separate
closed typed Rust, versioned batch, CLI, and standard MCP contract for
Spreadsheet ListObject tables. Ordinary typed `remove` deletes one table
through its stable `/SheetName/table[N]` path. The CLI adds a table below an
existing worksheet with `--type table`, `--name`, one final `--range`, and one
repeated `--table-column` per range column. CLI `set` carries forward omitted
fields from a mutable semantic node; Rust, batch, and MCP replacements use one
complete `table` value.

The value owns a workbook-wide `name`, an optional distinct `displayName`, the
final rectangular A1 range, exact ordered column names, header and totals-row
flags, typed filter criteria, one closed built-in style, and first/last-column
plus row/column-stripe flags. New tables default to a header, no totals row,
`TableStyleMedium2`, and row stripes. Styles are `none`, light 1–21, medium
1–28, or dark 1–11; `none`
omits `tableStyleInfo` and therefore requires every style flag to be false. The
range includes enabled header and totals rows and must retain at least one data
row. When a header is enabled, column names are stamped into its cells and the
table owns an AutoFilter range that excludes an enabled totals row.

Table and display names are limited to 255 characters, follow Excel's Unicode
identifier grammar, and may not resemble A1 or R1C1 references. Their shared
workbook namespace is case-insensitively collision-checked against every table
`name`/`displayName` and every defined name. Column names are 1–255 XML-safe
characters without surrounding whitespace and are unique case-insensitively.
The editor accepts at most 65,536 tables per workbook, requires exactly one
column identity per range column, and atomically rejects table overlap, merged
cell overlap, or overlap with a worksheet-level AutoFilter. A `set` excludes
the target table from these identity and geometry checks.

The OPC writer allocates a collision-free table ID and `xl/tables/tableN.xml`
part, content-type override, worksheet relationship, and ordered `tableParts`
entry. Removal first proves that the table has no owned or unexpected inbound
relationship graph, then removes only its worksheet reference, relationship,
content type, and part. Strict and transitional SpreadsheetML and relationship
dialects are retained. Replacement preserves supported unknown table-root,
style, and extension content. Unknown table-column metadata, formulas, totals
metadata, custom styles, or final collection data that cannot be retained make
the semantic node non-mutable or fail with
`use.office.spreadsheet_table_unknown_content` instead of being flattened.

Replacement also preserves formula identity for common ListObject structured
references. It maps the old table `name` to the new `name`, the old effective
`displayName` to the new effective `displayName`, and old columns to new
columns by physical position. When the old aliases coincide, the new display
name is preferred. The editor applies those rewrites to worksheet cell
formulas, workbook defined names, conditional-format and data-validation
formulas, chart formulas, and calculated-column/totals-row formulas in table
parts. Formula string literals and external-workbook structured references are
not changed.

Table-local forms such as `[@Qty]` are rewritten only for a formula cell inside
the old table range or another carrier whose ListObject ownership is provable.
An affected local reference with unknown ownership, or a local reference
across a table range/header/totals-row geometry change, fails atomically with
`use.office.spreadsheet_table_formula_rewrite_unsupported`. Geometry changes
with matching explicit references clear formula and chart caches before the
workbook is marked for full recalculation. Removal fails atomically with
`use.office.spreadsheet_table_referenced` while an explicit target or a
provably applicable/unknown local structured reference remains. A qualified
reference to another worksheet and a local reference provably owned by another
table do not block the target lifecycle.

Semantic reads expose table nodes and stable child column nodes, including
name/display name, normalized range, header/totals state, filter children,
built-in style, display flags, table ID, and `nativeMutable`. Query supports
selectors such as `table[name=Sales]` and
`filtercolumn[criteriaType=between]`. Exact replay emits tables after worksheet
creation and cell values so header stamping and final supported part maps
reproduce canonically. Tests cover typed lifecycle, multi-worksheet path identity,
validation and rollback, namespace and range collisions, strict OOXML, unknown
content and relationship safety, replay after table replacement, native CLI
lifecycle and atomic batch, and a complete standard MCP unsaved/save/reopen/
remove lifecycle with an unusable OfficeCLI provider.

This milestone does not own calculated-column formulas, totals-row labels or
functions, date-group/color/icon filters, unsupported embedded/imported sort
state, custom table styles, query tables, external data, slicers, or pivot-table
integration. Those remain separate closed contracts rather than generic table
properties. Exact mutable tables without totals rows can be physically sorted
through the separate sort contract.

Native `merge-cells` and `unmerge-cells` form a separate typed Rust, batch,
CLI, and standard MCP contract. The CLI projects them as
`office native set <file> <range> --merge-cells true|false`, so merge state can
compose atomically with one content write, text formatting, cell presentation,
and a hyperlink. The path must identify a Spreadsheet cell or rectangular A1
range and is normalized before mutation. Repeating an exact merge is
idempotent. Any non-identical geometric overlap fails with
`use.office.spreadsheet_merge_overlap`; a range intersecting a ListObject table
fails with `use.office.spreadsheet_merge_table_overlap` rather than producing
an invalid workbook.

Unmerge is intentionally exact and non-destructive. An absent exact range is an
unchanged success only when it is disjoint from every merge. If the requested
range intersects but does not exactly equal an existing merge,
`use.office.spreadsheet_merge_not_exact` returns the intersecting `validRanges`
so the caller can remove each one explicitly. It never performs a sweep. The XML
writer retains strict or transitional SpreadsheetML, schema child order,
unknown `mergeCells` attributes, extension children, and unrelated worksheet
data. Removing the last merge removes the collection only when its remaining
bytes are known to be whitespace and its only attribute is `count`; otherwise
`use.office.spreadsheet_merge_unknown_content` fails closed.

Semantic reads expose each merge as `/SheetName/mergeCell[N]`, annotate an
observed or virtual covered cell with `merge=<normalized-ref>` and
`mergeAnchor=true|false`, and report `merge=true|false` on exact range reads.
Blank covered cells are virtual and do not expand sparse `sheetData`. HTML and
SVG project the same metadata only as inert `data-merge` and
`data-merge-anchor` attributes. Merge collections are bounded to 100,000
ranges; overlap validation and observed-cell annotation use ordered sweeps
rather than a cell-by-range product. Versioned replay dump emits exact merge
mutations after cell values. Tests cover strict OOXML, unknown data, table and
range conflicts, idempotence, exact unmerge, rollback, replay, semantic
readback, CLI, and a complete standard MCP lifecycle with an unusable
OfficeCLI provider.

Native `set-hyperlink` is implemented through one typed Rust, batch, CLI, and
standard MCP contract. Word adds an external HTTP/HTTPS/mailto relationship or
internal bookmark anchor to a body, header, or footer paragraph, updates an
existing hyperlink path, and supports display text and tooltip in each owning
part. Spreadsheet adds or updates an external relationship or internal
workbook location on one cell or a bounded rectangular range, supports display
text and tooltip, and auto-creates a missing single cell with the display or
target text. Range links preserve existing cell contents, expose a stable
worksheet hyperlink path, and reject overlaps with a typed conflict error.
Presentation attaches an external shape-wide click or an internal jump to an
existing `slide[N]` target and optional tooltip to a shape; separate display
text remains unsupported. All three formats expose
stable semantic hyperlink nodes to `get`, `query`, annotated views, CLI, and
MCP, and remove them through the normal typed `remove` operation. External URI
validation rejects active or relative schemes, embedded credentials, controls,
and malformed targets. Relationship IDs are allocated or reused safely and are
garbage-collected only when unused, including when an owning paragraph, cell,
shape, or slide is removed. Atomic batches roll back every XML and relationship
change on failure, and both strict and transitional OOXML dialects are retained.

Native `add-comment` and `set-comment` are implemented through one typed Rust,
batch, CLI, and standard MCP contract. Word anchors a plain legacy comment to a
main-document paragraph or run, creates `word/comments.xml` and the
range/reference markers, and returns `/comments/comment[N]`. Spreadsheet
creates classic cell notes, including on blank cells, with an author table,
comments part, VML note drawing, worksheet `legacyDrawing` reference, and
stable `/SheetName/A1/comment` paths. Presentation creates legacy per-slide
comment parts plus the shared presentation author list, maintains monotonically
increasing indexes per author, accepts optional signed-32-bit EMU coordinates,
and returns `/slide[N]/comment[M]`.

All three formats expose comments through semantic `get`, `query`, and bounded
annotated views. Partial typed updates preserve omitted properties. Removal
uses the ordinary `remove` mutation; removing an owning Word paragraph/run,
Spreadsheet cell/range, or Presentation slide also removes its owned comments
and unreferenced relationship, content-type, and VML resources. Mutation is
atomic, unknown OOXML attributes and extension nodes survive updates, and
strict/transitional root and relationship dialects are preserved. Format-only
properties fail explicitly: Spreadsheet rejects initials and slide positions,
while Word rejects slide positions.

This milestone is intentionally legacy-comment scope, not full OfficeCLI or
modern Office comment parity. PowerPoint modern threaded comments and replies,
Word replies/resolved state and `commentsExtended.xml`, writable comment dates,
rich comment bodies, Word header/footer comment anchoring, and Spreadsheet
threaded comments remain unimplemented.

Row and column edits update cell and row references, dimensions, column
definitions, defined names, workbook view state, merges, filters, selections,
validation, conditional formatting, hyperlinks, sort state, ignored errors,
tables, comments, VML note anchors, drawing anchors, and chart formulas.
Supported local and cross-sheet A1 formula references, including absolute,
rectangular, whole-row, and whole-column references, are rewritten and their
cached values are invalidated; external references and string literals are
preserved. Unsafe 3D-reference and pivot-table structural or copy edits fail
closed and roll back.

Native add supports Word paragraphs and bounded table/row/cell structures,
while remove supports Word paragraphs, runs, tables, rows, and cells with
structural last-child invariants and table-grid maintenance. Spreadsheet cells
and worksheets and Presentation slides, text shapes, and basic DrawingML tables
also support native add/remove. Presentation table creation emits a real
`p:graphicFrame` and `a:tbl`; row creation follows the existing `a:tblGrid`,
blank cells accept native text replacement, and row removal updates the frame
height. Presentation columns are stable virtual `/table[N]/col[M]` nodes backed
by one `a:gridCol` and the corresponding cell in every row. Native insertion,
EMU width mutation, removal, same-table move/copy/swap, and semantic `get`
update the grid, all affected rows, and the graphic-frame width together. Slide
and worksheet removal updates their OPC relationships, content types, and owned
parts. The typed editor and `office native batch`
provide all-or-nothing in-memory rollback, bounded versioned inputs, atomic
save/save-as, revision-conflict detection, and byte preservation for untouched
package parts and XML subtrees. Safe raw XML inspection and replacement are now
implemented for existing parts. Raw reads validate XML before returning
normalized UTF-8 text and can export the original bytes. Raw replacement accepts
bounded UTF-8 input, rejects content-type and relationship parts, requires the
root local name and namespace to remain unchanged, and participates in the same
semantic/OPC validation and atomic rollback as every other native mutation.
Known part creation is implemented for Word chart/header/footer carriers and
Spreadsheet/Presentation chart carriers. It allocates collision-free part
names, writes content-type overrides, creates owner relationships, returns typed
relationship receipts, handles transitional and strict OOXML namespaces, and
rolls back every package change on failure. It does not yet insert a visible
chart frame or Word section reference.

Bounded raster image add/read/remove is implemented natively for PNG, JPEG, and
GIF across all three formats. Input bytes are base64 only at the typed batch
boundary and are never returned by normal CLI output. The decoder validates the
declared format against the decoded signature and basic image structure, reads
source dimensions, enforces byte/dimension/pixel bounds, and preserves aspect
ratio when only width or height is requested. Word inserts a real inline
DrawingML picture under `/body`, a paragraph, or a table cell. Spreadsheet
requires an anchor cell such as `/Sheet1/A1`, creates or reuses the worksheet
drawing, and inserts a one-cell anchor. Presentation inserts a real picture in
the selected slide shape tree. Every operation allocates a collision-free media
part, owner relationship, non-visual identity, name/alternative text metadata,
and final pixel dimensions.

Picture removal is reference-aware. It removes the owning XML subtree and an
unused image relationship first, then removes the media part and content-type
override only when no relationship anywhere in the package still targets that
part. All XML, relationship, content-type, and media changes participate in the
existing atomic batch rollback. Process-level tests cover Word, Spreadsheet,
and Presentation with an unusable OfficeCLI provider path; separate format
tests cover PNG, JPEG, GIF, dimension inference, invalid data, rollback, and
batch `createdImages` receipts. OOXML SVG image embedding is deferred because
interoperable SVG parts require a raster fallback representation. Replacement,
crop, effects, floating/advanced anchors, and rich image layout are not
implemented yet.

Native semantic rendering is now implemented as a separate read-only layer.
`NativeOfficeDocument::html_view` and `svg_view` produce standalone artifacts
for Word, Spreadsheet, and Presentation. Word HTML retains body, header/footer,
paragraph/run, table, picture, style, and stable-path semantics. Word SVG stacks
regions, paragraphs, tables, and validated pictures while retaining escaped
text and stable block paths. Spreadsheet output groups only observed rows and
cells, so a workbook containing both `A1` and `XFD1048576` cannot force a dense
grid allocation. Its SVG is likewise a sparse vertical semantic projection
rather than a dense worksheet canvas.
Presentation HTML and SVG use bounded semantic transforms for slide cards,
text shapes, tables, pictures, groups, charts, and connectors. Every SVG is
well-formed XML with accessible title/description metadata and no script or
external URL surface. The output does not claim theme, font, pagination, print,
or layout fidelity.

Artifacts are deterministic and contain no time or source filename. All text
and attributes are escaped, HTML declares a restrictive CSP with scripts and
network access disabled, external relationships remain inert, and only
internally related, structurally validated PNG/JPEG/GIF bytes may become
`data:` URLs. Composition stops at 16 MiB. CLI `view ... html|svg --output`
publishes atomically without replacing an existing path; inline CLI output uses
the same render bound and MCP retains its stricter 8 MiB structured-result
bound. Unit and process tests cover hostile markup, deterministic hashes,
sparse cells, invalid raster parts, all-format HTML and SVG,
no-clobber output, standard MCP, and an unusable OfficeCLI path.

Native live watch is available through `office native watch <file>`. The typed
`NativeOfficeWatchServer` renders before binding, listens only on IPv4
loopback, uses an ephemeral port by default, and issues a fresh 256-bit token.
The fixed wrapper, preview, status JSON, and standard SSE stream all require
that token or an HttpOnly same-site cookie and reject a non-matching Host. The
semantic document is isolated in a sandboxed iframe; its own CSP continues to
disable script and network access. A 50–10,000 ms bounded poller (250 ms by
default) tracks length/mtime and Unix device/inode/ctime, then reopens and
fully renders changed saved revisions. A failed revision never replaces the
last valid preview; a typed error is emitted and the poller retries until the
file recovers. The foreground CLI accepts `--port`, `--poll-ms`, and an optional
24-hour-bounded `--timeout-ms`, and prints one machine-readable startup receipt
with `--json` before serving.

The watch surface is deliberately read-only. It has no mutation endpoint,
private resident pipe, or custom RPC envelope, never invokes OfficeCLI or
LibreOffice, and sees an MCP session only after `office_save`. Full-page saved
refresh is implemented for Word, Spreadsheet, and Presentation. Inline
Spreadsheet edits, drag interactions, selection/mark/goto overlays,
slide-scoped patches, automatic browser launching, and layout goldens remain
outside this milestone. Unit and process tests cover token and Host rejection,
CSP/cookie headers, SSE, last-good retention, corrupt-file recovery, separate
CLI mutation, graceful shutdown, and an unusable OfficeCLI provider.
Runtime evidence is currently macOS/Linux-first. Windows compiles the same
CLI/server contracts and uses the portable length/mtime stamp, but remains a
preview target under the repository's separate Windows promotion gate.

Native issue analysis is implemented as a bounded, read-only pass over the
semantic tree and OPC relationship graph. `NativeOfficeDocument::issues`
defaults to 200 returned records and accepts a hard maximum of 1,000. Filtering
by the broad `format`, `content`, or `structure` category, or by an exact stable
subtype, occurs before the window is applied. The report always distinguishes
the total matching `count`, `returned` records, and `truncated` state.

The initial conservative rules are `missing_alt_text`, `broken_part_ref`,
`formula_not_evaluated`, `formula_ref_missing_sheet`, `formula_eval_error`, and
`low_contrast`. Missing-sheet detection scans direct quoted or ASCII-unquoted
worksheet-qualified formula references while excluding formula string
literals and external-workbook references. Low-contrast detection compares
only explicit RGB run text against
the explicit fill of its owning shape; scheme, inherited, transformed, or
translucent colors are skipped. Broken references are checked against the
typed relationship graph and expected relationship kind. The scanner does not
guess at text overflow, object overlap, theme resolution, pagination, or
application layout. Consequently, an empty report is useful evidence for the
implemented rules, not a complete Office validity or visual-fidelity claim.

CLI `office native view <file> issues` and MCP `office_view` with
`view=issues` expose the same typed report. CLI accepts `--type` and `--limit`;
MCP accepts `issueType` and `limit`. Unit tests cover all three formats,
filtering, limits, string-literal exclusion, missing-sheet discrimination,
explicit low contrast, broken relationships, and clean blank documents.
Process tests run the CLI and a complete standard MCP lifecycle with an
unusable OfficeCLI path, proving the view is native and provider-independent.

Browser-injected PNG screenshot output is implemented for all three formats at
the root facade. It stages the deterministic HTML in a private temporary
directory, converts the local path to a `file://` URL, and passes that URL plus
a temporary PNG destination to the existing object-safe `PageRenderer`.
`a3s-use-office` has no Browser dependency. The facade validates that the
provider returned exactly one expected regular, non-symlink PNG artifact,
checks its decoded dimensions, size, and SHA-256 receipt, then publishes the
final destination atomically without overwriting an existing entry. The PNG is
limited to 64 MiB; the rendering deadline defaults to 30 seconds and must be
between 1 and 120 seconds. External relationships are never fetched.

The typed `NativeOfficeScreenshotRenderer` accepts an injected
`Arc<dyn PageRenderer>`; `capture_native_office_screenshot` performs normal
Browser discovery for convenience. CLI
`office native view <file> screenshot --output <file.png>` and MCP
`office_view` with `view=screenshot` return the same typed receipt. MCP requires
`output` and accepts optional `timeoutMs`; the session lock is released before
Browser work starts. Process tests cover DOCX, XLSX, and PPTX CLI screenshots,
an MCP screenshot lifecycle, PNG hashes, invalid arguments, Browser-disabled
builds, and no-clobber behavior while setting an unusable OfficeCLI path.
Screenshots are raster captures of the semantic preview, not Office layout
fidelity. Layout goldens remain open.

Basic Presentation table structure is deliberately bounded. Table dimensions
must be positive, no mutation may exceed 5,000 rows, 5,000 columns, or 100,000
cells, and an explicit row width must equal the parent grid. `add --type cell`
only fills an underfull row; a full row rejects the append because PowerPoint
would silently discard a cell beyond `a:tblGrid`. Direct cell removal is
similarly limited to repairing an overflow row. Removing the final row is
rejected. Column insertion accepts a zero-based slot or appends, uses the
average existing grid width, and creates a cell in every row. Column removal
retains at least one column; column move/copy/swap remains within one table.
These structural column operations require a rectangular unmerged table and
fail closed when merged-cell spans would need rewriting. Explicit column width
mutation uses a positive signed-64-bit EMU value and keeps frame width equal to
the grid-width sum. Merged-cell editing, custom row heights, table styles,
fills, borders, and non-text cell styling remain later Presentation work. Run
text formatting and paragraph alignment inside table cells use the shared
typed format mutation. None of these operations invokes OfficeCLI or
LibreOffice.

Typed move/copy/swap is implemented as a bounded arrangement layer. `Index` is
zero-based and is evaluated after source removal for a move; `Before` and
`After` resolve stable semantic paths before mutation. A copy with no position
is inserted immediately after its source, while a move with no position moves
to the end of its supported sibling set. Every operation participates in the
same atomic batch rollback and semantic post-validation as add/set/remove.

- Word moves and swaps paragraphs/tables inside a body or table cell, rows
  inside a table, cells inside a row, and runs inside a paragraph. Copies cover
  identity-free paragraphs, tables, rows, and runs. Cross-parent movement,
  relationship or document-identity copies, and table-cell copy fail closed;
  table-cell copy remains blocked until table-grid resizing is defined.
- Spreadsheet moves, copies, and swaps worksheets. Worksheet copy requires a
  distinct name and retains the existing loss-preserving owned-subgraph clone.
  Dense plain rows can also move, copy, and swap with row/cell reference
  renumbering. Row arrangement rejects sparse rows, formulas anywhere in the
  workbook, defined names, row-addressed metadata, worksheet relationships, and
  unsafe shared-string or identity copies before mutation.
- Presentation moves, copies, and swaps slides. Slide copy currently accepts
  only a layout-only relationship graph. Top-level shapes, pictures, tables,
  charts, connectors, and groups can move or swap within one slide. Copy is
  limited to a plain relationship-free shape without placeholders, extension
  identities, or relationship attributes; the copy receives a fresh `cNvPr`
  ID and name. Cross-slide object movement and relationship-owning copies fail
  closed.

Root-scoped replay dump is implemented for the canonical subset that current
typed mutations can reproduce exactly: plain Word paragraphs and rectangular
tables, Spreadsheet worksheets, typed defined names, typed cells, typed
worksheet/table AutoFilters, ListObject tables, stable physical row order with
supported typed sort state, canonical frozen panes and import date styles,
merged ranges, typed data-validation rules, and
canonical typed conditional-format rules, natively recalculable formula caches
and canonical cached dynamic-array spills, and Presentation slides with plain
one-run text shapes and canonical basic tables. The versioned
artifact records document kind, `/` scope, blank-template part-map SHA-256,
ordered mutations, and expected result part-map SHA-256. Native `batch` checks
both fingerprints and restores the original package on a failed result check.
Unsupported rich or non-canonical content fails with
`use.office.dump_unsupported`; no element or resource is skipped. Inputs and
file output are limited to 8 MiB and 10,000 mutations, inline output is limited
to 1 MiB, and dump refuses to overwrite an existing path.

Native template merge is implemented across Word, Spreadsheet, and
Presentation. JSON data must be an object; literal top-level keys override
flattened dot/bracket paths, and replacements are single pass. Word processes
the main document, headers, footers, footnotes, endnotes, and comments.
Presentation processes slides and notes. Spreadsheet processes inline strings,
direct string values, and referenced shared rich strings while retaining run
ownership and skipping phonetic text. Shared-string replacements are counted
per referencing cell. Resolved placeholders in unsupported non-string cells
fail closed instead of changing their value type.

The editor restores the original package on any XML, type, or semantic failure.
The CLI accepts inline JSON, `@file.json`, or an existing `.json` path. File
inputs are regular, non-symlink files bounded to 8 MiB; flattened data also has
entry, key, depth, and total-byte limits. Template/output identity, including
Unix hard links, is rejected. Output creation is atomic and no-clobber by
default, with explicit `--force` replacement. Process-level tests cover all
three formats with an unusable OfficeCLI path and verify template bytes remain
unchanged.

Cross-parent/reference-graph arrangement beyond the bounded move/copy/swap
coverage above, advanced image mutation and OOXML SVG fallback, complex/custom
part carriers, Presentation table merges/rich styles, subtree and
rich-structure dump, advanced rich-format operations, modern threaded comments
and legacy-comment replies/resolution/rich bodies, and
complete Excel function breadth, structured-reference forms beyond common row
items and contiguous columns, qualified functions, and external-workbook
formula calculation remain before their respective gates can be promoted.
Creation and structural mutation remain under the interoperability gate until
Microsoft Office and optional CI LibreOffice checks confirm that no repair
dialog is required.

### Gate 3 — Rich Word

- Complete the Word scope above, including revisions, fields, forms, charts,
  equations, international text, and RTL.

### Gate 4 — Rich Spreadsheet

- Complete the Spreadsheet scope above.
- Formula conformance corpus, reference rewrite, dynamic arrays, charts, and
  pivot-table interoperability gates.

### Gate 5 — Rich Presentation

- Complete the Presentation scope above.
- Layout, theme, animation, transition, media, diagram, and chart fidelity gates.

### Gate 6 — Native product promotion

- HTML/SVG/screenshot rendering and live watch.
- Native standard MCP server and a packaged first-party Office Skill.
- CLI compatibility corpus for every core command.
- Fuzzing for ZIP, XML, selector, formula, and mutation inputs.
- macOS and Linux release evidence; Windows remains preview until its separate
  platform gate is promoted.

Status: native bounded annotated and issue analysis, semantic rendering,
Browser-injected screenshots, the explicit `a3s-office mcp`
target, and the packaged `a3s-use-office` Skill are available for evidence
gathering. Annotated views, issue reports, HTML, SVG, semantic-preview PNG
screenshots, and saved-revision live watch cover all three formats; PNG requires
a ready Browser provider. They are available through typed Rust APIs, `office
native view|watch`, `office_view`, and progressive
Word/Spreadsheet/Presentation/MCP Skill references.
The MCP target's 12 typed tools use bounded in-process sessions for validate,
create/open/list, semantic reads, annotations, and issues, constrained raw XML, atomic
mutation batches, immutable-template merge, save, and close. It limits open
sessions to 64, batch and result JSON to 8 MiB, a batch to 10,000 mutations,
query output to 1,000 nodes, annotated and issue output to 1,000 records, and
raw XML output to 1 MiB. Mutations are not persisted until `office_save`, and a
dirty session cannot close without save or explicit discard. Process-level
tests complete a standard MCP initialize/list/call lifecycle, verify annotated
reads against unsaved typed session state, capture a real PNG when Chrome is
available, and use an unusable OfficeCLI path. Skill process tests exercise
bounded `list`, `get --full`, and `path` discovery with the same unusable
provider, and release archives smoke-check the packaged `SKILL.md`. This preview
does not complete Gate 6: richer issue parity, interactive-watch parity,
compatibility corpus, fuzzing, advanced rich-format coverage, layout goldens,
and release evidence remain open, and the default Office target is not
promoted.

At Gate 6, native becomes the default and `a3s install use/office` no longer
downloads an engine. The OfficeCLI backend moves to an explicitly named
compatibility component for one deprecation cycle, then is removed.

## Verification strategy

- Unit tests cover package, XML, selectors, models, relationships, formulas,
  and mutations without spawning another process.
- Golden fixtures are small, synthetic, and checked for untouched-part hashes.
- Differential tests run the same compatibility corpus against the native
  engine and the pinned OfficeCLI binary, then compare normalized semantic
  results rather than ZIP byte layout.
- Interoperability CI may open and save outputs with Microsoft Open XML
  validation and LibreOffice; release candidates also undergo Microsoft Office
  repair-dialog checks. These are external acceptance oracles, never runtime
  dependencies.
- Rendering uses DOM and image golden tests with explicit tolerances.
- Fuzz targets retain every reproducer as a regression fixture.

## Current migration boundary

The `0.1.x` CLI exposes native blank creation, reads, typed
add/set/remove/move/copy/swap, scoped cross-format literal/regex replacement,
cross-format text formatting, typed Spreadsheet number/fill/border/alignment
and cell-presentation formatting, exact Spreadsheet merged-cell editing, typed
Spreadsheet physical row sorting with persisted sort state, Spreadsheet
CSV/TSV import with typed inference and header filter/freeze behavior,
worksheet/table AutoFilters, data-validation,
conditional-formatting, and scoped defined-name editing, typed ListObject table
lifecycle, inert hyperlinks,
typed cross-format
legacy comments, Spreadsheet range and
row/column structure edits,
worksheet rename/reorder and loss-preserving worksheet copy, safe
`raw`/`raw-set`, known `add-part` carriers, exact root replay dump for the
canonical typed subset, native PNG/JPEG/GIF add/read/remove, cross-format
template merge with `merge`, all-format semantic HTML and SVG, bounded
all-format annotated and issue reports, all-format Browser-injected semantic
PNG screenshots, plus atomic batches under the native Office route.
The distribution also packages `a3s-use-office`, with bounded
`office skills list|get|path` access and a content SHA-256 in the unified
capability snapshot. Loading the Skill never discovers or starts OfficeCLI.
The explicit `mcp serve office-native` target now exposes the current typed
subset without OfficeCLI; only its optional screenshot view requires a ready
A3S Browser provider. Other Office commands and the default `mcp serve office`
target still delegate to the pinned OfficeCLI provider. This keeps existing
users functional while native coverage grows. The native APIs are deliberately
not advertised as full Office readiness, and `doctor` continues to report
compatibility-provider readiness until the native read and mutation gates are
met.
