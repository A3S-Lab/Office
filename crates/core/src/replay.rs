use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use crate::discovery::office_error;
use crate::editor::{
    NativeOfficeEditor, NativeOfficeMutation, NativeSpreadsheetAutoFilter,
    NativeSpreadsheetCellFormat, NativeSpreadsheetConditionalFormat,
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
    NativeSpreadsheetFrozenPane, NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope,
    NativeSpreadsheetSort, NativeSpreadsheetTable, SpreadsheetCellValue,
};
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::{DocumentKind, NativeOfficePackage};

pub const NATIVE_OFFICE_REPLAY_FORMAT: &str = "a3s.office.native-replay";
pub const NATIVE_OFFICE_REPLAY_SCHEMA_VERSION: u32 = 1;
pub const MAX_NATIVE_OFFICE_REPLAY_MUTATIONS: usize = 10_000;

const EMPTY_PRESENTATION_TITLE_SEED: &str = "a3s-replay-title";

/// Required starting state for a native replay artifact.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeReplayBase {
    Blank,
}

/// A versioned, non-RPC batch artifact that exactly recreates a supported
/// native OOXML package from the A3S blank template named by `baseSha256`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeReplayArtifact {
    pub format: String,
    pub schema_version: u32,
    pub document_kind: DocumentKind,
    pub scope: String,
    pub base: NativeOfficeReplayBase,
    pub base_sha256: String,
    pub result_sha256: String,
    pub mutations: Vec<NativeOfficeMutation>,
}

impl NativeOfficeReplayArtifact {
    /// Emits a replay artifact only when current typed mutations reproduce the
    /// complete uncompressed OOXML part map byte-for-byte.
    pub fn dump(document: &NativeOfficeDocument, scope: &str) -> UseResult<Self> {
        if scope != "/" {
            return Err(replay_error(
                "use.office.dump_scope_unsupported",
                "Native replay dump currently supports only the complete document scope '/'.",
            )
            .with_detail("scope", scope));
        }

        let mutations = emit_mutations(document)?;
        if mutations.len() > MAX_NATIVE_OFFICE_REPLAY_MUTATIONS {
            return Err(replay_error(
                "use.office.dump_mutation_limit",
                format!(
                    "Native replay dump requires {} mutations; the limit is {MAX_NATIVE_OFFICE_REPLAY_MUTATIONS}.",
                    mutations.len()
                ),
            )
            .with_detail("mutations", mutations.len()));
        }

        let blank =
            NativeOfficePackage::blank_in_memory(document.kind(), document.package().limits())?;
        let base_sha256 = blank.content_sha256();
        let mut candidate = NativeOfficeEditor::from_package(blank)?;
        if let Err(error) = candidate.apply_batch(&mutations) {
            return Err(dump_unsupported(
                "/",
                format!("The current typed mutations could not reconstruct this document: {error}"),
            )
            .with_detail("causeCode", error.code));
        }

        let source_sha256 = document.package().content_sha256();
        let result_sha256 = candidate.package().content_sha256();
        if result_sha256 != source_sha256 {
            let part = first_differing_part(document.package(), candidate.package())
                .unwrap_or_else(|| "<unknown>".to_string());
            return Err(dump_unsupported(
                "/",
                format!(
                    "OOXML part '{part}' cannot yet be recreated exactly by native replay mutations."
                ),
            )
            .with_detail("part", part)
            .with_detail("sourceSha256", source_sha256)
            .with_detail("replaySha256", result_sha256));
        }

        Ok(Self {
            format: NATIVE_OFFICE_REPLAY_FORMAT.to_string(),
            schema_version: NATIVE_OFFICE_REPLAY_SCHEMA_VERSION,
            document_kind: document.kind(),
            scope: scope.to_string(),
            base: NativeOfficeReplayBase::Blank,
            base_sha256,
            result_sha256,
            mutations,
        })
    }

    /// Validates artifact identity, schema, scope, fingerprints, and limits
    /// without applying any mutation.
    pub fn validate(&self) -> UseResult<()> {
        if self.format != NATIVE_OFFICE_REPLAY_FORMAT {
            return Err(replay_error(
                "use.office.replay_format_unsupported",
                format!(
                    "Native Office replay format '{}' is not supported; expected '{NATIVE_OFFICE_REPLAY_FORMAT}'.",
                    self.format
                ),
            ));
        }
        if self.schema_version != NATIVE_OFFICE_REPLAY_SCHEMA_VERSION {
            return Err(replay_error(
                "use.office.replay_schema_unsupported",
                format!(
                    "Native Office replay schema version {} is not supported; expected {NATIVE_OFFICE_REPLAY_SCHEMA_VERSION}.",
                    self.schema_version
                ),
            ));
        }
        if self.scope != "/" {
            return Err(replay_error(
                "use.office.replay_scope_unsupported",
                "Native Office replay currently supports only the complete document scope '/'.",
            )
            .with_detail("scope", self.scope.clone()));
        }
        if !valid_sha256(&self.base_sha256) || !valid_sha256(&self.result_sha256) {
            return Err(replay_error(
                "use.office.replay_fingerprint_invalid",
                "Native Office replay fingerprints must be lowercase SHA-256 values.",
            ));
        }
        if self.mutations.len() > MAX_NATIVE_OFFICE_REPLAY_MUTATIONS {
            return Err(replay_error(
                "use.office.replay_mutation_limit",
                format!(
                    "Native Office replay contains {} mutations; the limit is {MAX_NATIVE_OFFICE_REPLAY_MUTATIONS}.",
                    self.mutations.len()
                ),
            )
            .with_detail("mutations", self.mutations.len()));
        }
        Ok(())
    }
}

fn emit_mutations(document: &NativeOfficeDocument) -> UseResult<Vec<NativeOfficeMutation>> {
    match document.kind() {
        DocumentKind::Word => emit_word(document.root()),
        DocumentKind::Spreadsheet => emit_spreadsheet(document.root()),
        DocumentKind::Presentation => emit_presentation(document.root()),
    }
}

fn emit_word(root: &DocumentNode) -> UseResult<Vec<NativeOfficeMutation>> {
    let Some(body) = root
        .children
        .iter()
        .find(|node| node.node_type == OfficeNodeType::Body)
    else {
        return Err(dump_unsupported(
            "/",
            "Word document has no replayable body.",
        ));
    };
    if root.children.len() != 1 {
        let unsupported = root
            .children
            .iter()
            .find(|node| node.node_type != OfficeNodeType::Body)
            .map_or("/", |node| node.path.as_str());
        return Err(dump_unsupported(
            unsupported,
            "Word headers, footers, and other sibling resources are not replayable yet.",
        ));
    }

    let preserve_blank_paragraph = body
        .children
        .first()
        .is_some_and(|node| node.node_type == OfficeNodeType::Paragraph);
    let mut mutations = Vec::new();
    if !preserve_blank_paragraph {
        mutations.push(NativeOfficeMutation::Remove {
            path: "/body/p[1]".to_string(),
        });
    }

    for (offset, block) in body.children.iter().enumerate() {
        match block.node_type {
            OfficeNodeType::Paragraph => {
                let needs_text = validate_word_paragraph(block)?;
                if offset == 0 && preserve_blank_paragraph {
                    if needs_text {
                        mutations.push(NativeOfficeMutation::SetText {
                            path: block.path.clone(),
                            text: block.text.clone(),
                        });
                    }
                } else {
                    if !needs_text {
                        return Err(dump_unsupported(
                            &block.path,
                            "A later empty Word paragraph without a run cannot be recreated by the current add-paragraph mutation.",
                        ));
                    }
                    mutations.push(NativeOfficeMutation::AddParagraph {
                        parent: "/body".to_string(),
                        text: block.text.clone(),
                    });
                }
            }
            OfficeNodeType::Table => emit_word_table(block, &mut mutations)?,
            _ => {
                return Err(dump_unsupported(
                    &block.path,
                    format!(
                        "Word node type '{}' is not exactly replayable yet.",
                        block.node_type.label()
                    ),
                ))
            }
        }
    }
    Ok(mutations)
}

fn validate_word_paragraph(paragraph: &DocumentNode) -> UseResult<bool> {
    require_plain_node(paragraph, OfficeNodeType::Paragraph)?;
    match paragraph.children.as_slice() {
        [] if paragraph.text.is_empty() => Ok(false),
        [run] if run.node_type == OfficeNodeType::Run => {
            require_plain_node(run, OfficeNodeType::Run)?;
            if !run.children.is_empty() || run.text != paragraph.text {
                return Err(dump_unsupported(
                    &paragraph.path,
                    "Word paragraph text does not map to one plain run.",
                ));
            }
            Ok(true)
        }
        _ => Err(dump_unsupported(
            &paragraph.path,
            "Word replay currently requires each paragraph to contain zero or one plain run.",
        )),
    }
}

fn emit_word_table(
    table: &DocumentNode,
    mutations: &mut Vec<NativeOfficeMutation>,
) -> UseResult<()> {
    require_plain_node(table, OfficeNodeType::Table)?;
    if table.children.is_empty() {
        return Err(dump_unsupported(
            &table.path,
            "Word replay does not support a table without rows.",
        ));
    }
    let columns = table.children[0].children.len();
    if columns == 0
        || table
            .children
            .iter()
            .any(|row| row.node_type != OfficeNodeType::TableRow || row.children.len() != columns)
    {
        return Err(dump_unsupported(
            &table.path,
            "Word replay currently requires a non-empty rectangular table.",
        ));
    }

    mutations.push(NativeOfficeMutation::AddTable {
        parent: "/body".to_string(),
        rows: table.children.len(),
        columns,
    });
    for row in &table.children {
        require_plain_node(row, OfficeNodeType::TableRow)?;
        for cell in &row.children {
            require_plain_node(cell, OfficeNodeType::TableCell)?;
            let [paragraph] = cell.children.as_slice() else {
                return Err(dump_unsupported(
                    &cell.path,
                    "Word replay requires each table cell to contain exactly one paragraph.",
                ));
            };
            let needs_text = validate_word_paragraph(paragraph)?;
            if paragraph.text != cell.text {
                return Err(dump_unsupported(
                    &cell.path,
                    "Word table-cell text does not map to its single paragraph.",
                ));
            }
            if needs_text && !cell.text.is_empty() {
                mutations.push(NativeOfficeMutation::SetText {
                    path: cell.path.clone(),
                    text: cell.text.clone(),
                });
            }
        }
    }
    Ok(())
}

fn emit_spreadsheet(root: &DocumentNode) -> UseResult<Vec<NativeOfficeMutation>> {
    let sheets = root
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .collect::<Vec<_>>();
    if sheets.is_empty() {
        return Err(dump_unsupported(
            "/",
            "Spreadsheet replay requires at least one worksheet.",
        ));
    }
    let mut mutations = Vec::new();
    let mut requires_recalculation = false;
    for (sheet_offset, sheet) in sheets.into_iter().enumerate() {
        let name = sheet
            .path
            .strip_prefix('/')
            .filter(|name| !name.is_empty())
            .ok_or_else(|| {
                dump_unsupported(&sheet.path, "Spreadsheet worksheet path has no name.")
            })?;
        if sheet_offset == 0 {
            if name != "Sheet1" {
                mutations.push(NativeOfficeMutation::RenameWorksheet {
                    path: "/Sheet1".to_string(),
                    name: name.to_string(),
                });
            }
        } else {
            mutations.push(NativeOfficeMutation::AddWorksheet {
                name: name.to_string(),
            });
        }

        let mut merged_ranges = Vec::new();
        let mut auto_filter = None;
        let mut sort_state = None;
        let mut frozen_pane = None;
        let mut cell_formats = Vec::new();
        let mut conditional_formats = Vec::new();
        let mut tables = Vec::new();
        let mut validations = Vec::new();
        let spill_owners = spreadsheet_spill_owners(sheet)?;
        for child in &sheet.children {
            match child.node_type {
                OfficeNodeType::Row => {
                    require_plain_node(child, OfficeNodeType::Row)?;
                    for cell in &child.children {
                        let reference = spreadsheet_cell_reference(cell)?;
                        if spill_owners
                            .get(&reference)
                            .is_some_and(|anchor| *anchor != reference)
                        {
                            if cell.format.contains_key("formula") {
                                return Err(dump_unsupported(
                                    &cell.path,
                                    "Legacy multi-cell array formula storage is not replayable.",
                                ));
                            }
                            continue;
                        }
                        if let Some(format) = spreadsheet_cell_format(cell)? {
                            cell_formats.push((cell.path.clone(), format));
                        }
                        requires_recalculation |= cell
                            .format
                            .get("formulaCached")
                            .is_some_and(|cached| cached == "true");
                        mutations.push(NativeOfficeMutation::SetCellValue {
                            path: cell.path.clone(),
                            value: spreadsheet_value(cell)?,
                        });
                    }
                }
                OfficeNodeType::Range if child.tag == "mergeCell" => {
                    if child.style.is_some()
                        || !child.children.is_empty()
                        || child
                            .format
                            .keys()
                            .any(|key| !matches!(key.as_str(), "ref" | "merge"))
                    {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet merged range contains unsupported semantic data.",
                        ));
                    }
                    let reference = child.format.get("ref").ok_or_else(|| {
                        dump_unsupported(&child.path, "Spreadsheet merged range has no reference.")
                    })?;
                    merged_ranges.push(format!("{}/{}", sheet.path, reference));
                }
                OfficeNodeType::DataValidation => {
                    validations.push(spreadsheet_data_validation(child)?);
                }
                OfficeNodeType::ConditionalFormatting => {
                    let expected_priority = conditional_formats.len() + 1;
                    if child
                        .format
                        .get("priority")
                        .and_then(|value| value.parse::<usize>().ok())
                        != Some(expected_priority)
                    {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet replay requires canonical sequential conditional-format priorities.",
                        ));
                    }
                    conditional_formats.push(
                        NativeSpreadsheetConditionalFormat::from_semantic_node(child).map_err(
                            |error| {
                                dump_unsupported(
                                    &child.path,
                                    format!(
                                        "Spreadsheet conditional format is not replayable: {error}"
                                    ),
                                )
                            },
                        )?,
                    );
                }
                OfficeNodeType::Table => {
                    tables.push(NativeSpreadsheetTable::from_semantic_node(child).map_err(
                        |error| {
                            dump_unsupported(
                                &child.path,
                                format!("Spreadsheet table is not replayable: {error}"),
                            )
                        },
                    )?);
                }
                OfficeNodeType::AutoFilter => {
                    if auto_filter.is_some() {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet replay found multiple worksheet AutoFilters.",
                        ));
                    }
                    auto_filter = Some(
                        NativeSpreadsheetAutoFilter::from_semantic_node(child).map_err(
                            |error| {
                                dump_unsupported(
                                    &child.path,
                                    format!("Spreadsheet AutoFilter is not replayable: {error}"),
                                )
                            },
                        )?,
                    );
                }
                OfficeNodeType::SortState => {
                    if sort_state.is_some() {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet replay found multiple worksheet sort states.",
                        ));
                    }
                    let reference = child.format.get("ref").ok_or_else(|| {
                        dump_unsupported(&child.path, "Spreadsheet sort state has no source range.")
                    })?;
                    let sort =
                        NativeSpreadsheetSort::from_semantic_node(child).map_err(|error| {
                            dump_unsupported(
                                &child.path,
                                format!("Spreadsheet sort state is not replayable: {error}"),
                            )
                        })?;
                    sort_state = Some((format!("{}/{}", sheet.path, reference), sort));
                }
                OfficeNodeType::FrozenPane => {
                    if frozen_pane.is_some() {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet replay found multiple frozen panes.",
                        ));
                    }
                    if child.format.get("nativeMutable").map(String::as_str) != Some("true") {
                        return Err(dump_unsupported(
                            &child.path,
                            "Spreadsheet frozen pane contains unsupported view state.",
                        ));
                    }
                    frozen_pane = Some(
                        NativeSpreadsheetFrozenPane::from_semantic_node(child).map_err(
                            |error| {
                                dump_unsupported(
                                    &child.path,
                                    format!("Spreadsheet frozen pane is not replayable: {error}"),
                                )
                            },
                        )?,
                    );
                }
                _ => {
                    return Err(dump_unsupported(
                        &child.path,
                        "Spreadsheet worksheet contains an unsupported semantic node.",
                    ));
                }
            }
        }
        mutations.extend(
            merged_ranges
                .into_iter()
                .map(|path| NativeOfficeMutation::MergeCells { path }),
        );
        mutations.extend(
            cell_formats
                .into_iter()
                .map(|(path, format)| NativeOfficeMutation::SetCellFormat { path, format }),
        );
        mutations.extend(conditional_formats.into_iter().map(|conditional_format| {
            NativeOfficeMutation::AddConditionalFormat {
                sheet: sheet.path.clone(),
                conditional_format,
            }
        }));
        mutations.extend(validations.into_iter().map(|validation| {
            NativeOfficeMutation::AddDataValidation {
                sheet: sheet.path.clone(),
                validation,
            }
        }));
        if let Some(filter) = auto_filter {
            mutations.push(NativeOfficeMutation::AddSpreadsheetAutoFilter {
                sheet: sheet.path.clone(),
                filter,
            });
        }
        if let Some(pane) = frozen_pane {
            mutations.push(NativeOfficeMutation::SetSpreadsheetFrozenPane {
                sheet: sheet.path.clone(),
                pane,
            });
        }
        mutations.extend(tables.into_iter().map(|table| {
            NativeOfficeMutation::AddSpreadsheetTable {
                sheet: sheet.path.clone(),
                table,
            }
        }));
        if let Some((path, sort)) = sort_state {
            mutations.push(NativeOfficeMutation::SortSpreadsheetRange { path, sort });
        }
    }
    let name_collections = root
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::NamedRangeCollection)
        .collect::<Vec<_>>();
    if name_collections.len() > 1 {
        return Err(dump_unsupported(
            "/namedrange",
            "Spreadsheet replay found multiple named-range collections.",
        ));
    }
    if let Some(collection) = name_collections.first() {
        if collection.path != "/namedrange"
            || collection.tag != "namedranges"
            || collection.style.is_some()
            || !collection.text.is_empty()
        {
            return Err(dump_unsupported(
                &collection.path,
                "Spreadsheet named-range collection contains unsupported semantic data.",
            ));
        }
        for node in &collection.children {
            mutations.push(NativeOfficeMutation::AddNamedRange {
                named_range: spreadsheet_named_range(node)?,
            });
        }
    }
    if let Some(unsupported) = root.children.iter().find(|node| {
        !matches!(
            node.node_type,
            OfficeNodeType::Worksheet | OfficeNodeType::NamedRangeCollection
        )
    }) {
        return Err(dump_unsupported(
            &unsupported.path,
            "Spreadsheet root contains an unsupported semantic node.",
        ));
    }
    if requires_recalculation {
        mutations.push(NativeOfficeMutation::RecalculateSpreadsheetFormulas);
    }
    Ok(mutations)
}

fn spreadsheet_spill_owners(
    sheet: &DocumentNode,
) -> UseResult<std::collections::BTreeMap<CellReference, CellReference>> {
    let mut owners = std::collections::BTreeMap::new();
    for cell in sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .flat_map(|row| &row.children)
        .filter(|node| node.node_type == OfficeNodeType::Cell)
    {
        let formula_type = cell.format.get("formulaType").map(String::as_str);
        let formula_reference = cell.format.get("formulaRef");
        match (formula_type, formula_reference) {
            (None, None) => continue,
            (Some(kind), None) => {
                return Err(dump_unsupported(
                    &cell.path,
                    format!(
                        "Spreadsheet formula storage type '{kind}' is not canonical replay input."
                    ),
                ));
            }
            (Some(kind), Some(_)) if !kind.eq_ignore_ascii_case("array") => {
                return Err(dump_unsupported(
                    &cell.path,
                    format!(
                        "Spreadsheet formula storage type '{kind}' with a spill range is not replayable."
                    ),
                ));
            }
            (None, Some(_)) => {
                return Err(dump_unsupported(
                    &cell.path,
                    "Spreadsheet formula spill range has no array storage type.",
                ));
            }
            (Some(_), Some(_)) => {}
        }
        if cell.format.get("formulaCached").map(String::as_str) != Some("true") {
            return Err(dump_unsupported(
                &cell.path,
                "Spreadsheet array formula has no cached native result.",
            ));
        }
        let reference = formula_reference.ok_or_else(|| {
            dump_unsupported(&cell.path, "Spreadsheet array formula has no spill range.")
        })?;
        let anchor = spreadsheet_cell_reference(cell)?;
        let range = CellRange::parse(reference).map_err(|error| {
            dump_unsupported(
                &cell.path,
                format!("Spreadsheet formula spill range '{reference}' is invalid: {error}"),
            )
        })?;
        if !range.contains(anchor) {
            return Err(dump_unsupported(
                &cell.path,
                "Spreadsheet formula spill range does not contain its anchor.",
            ));
        }
        let cells = range.cell_count()?;
        if cells > crate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS
            || owners.len().saturating_add(cells) > crate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS
        {
            return Err(dump_unsupported(
                &cell.path,
                "Spreadsheet formula spills exceed the native replay cell limit.",
            )
            .with_detail("cells", owners.len().saturating_add(cells)));
        }
        for row in range.start.row..=range.end.row {
            for column in range.start.column..=range.end.column {
                let reference = CellReference { column, row };
                if owners.insert(reference, anchor).is_some() {
                    return Err(dump_unsupported(
                        &cell.path,
                        "Spreadsheet formula spill ranges overlap.",
                    ));
                }
            }
        }
    }
    Ok(owners)
}

fn spreadsheet_cell_reference(cell: &DocumentNode) -> UseResult<CellReference> {
    cell.path
        .rsplit_once('/')
        .and_then(|(_, reference)| CellReference::parse(reference).ok())
        .ok_or_else(|| {
            dump_unsupported(
                &cell.path,
                "Spreadsheet cell has an invalid semantic coordinate.",
            )
        })
}

fn spreadsheet_named_range(node: &DocumentNode) -> UseResult<NativeSpreadsheetNamedRange> {
    if node.node_type != OfficeNodeType::NamedRange
        || node.tag != "namedrange"
        || node.style.is_some()
        || !node.children.is_empty()
    {
        return Err(dump_unsupported(
            &node.path,
            "Spreadsheet named range contains unsupported semantic children or style data.",
        ));
    }
    let allowed = ["name", "ref", "scope", "comment", "volatile"];
    if let Some(key) = node
        .format
        .keys()
        .find(|key| !allowed.contains(&key.as_str()))
    {
        return Err(dump_unsupported(
            &node.path,
            format!("Spreadsheet named-range property '{key}' is not replayable."),
        ));
    }
    let name = required_format(node, "name")?.to_string();
    let reference = required_format(node, "ref")?.to_string();
    if node.text != reference {
        return Err(dump_unsupported(
            &node.path,
            "Spreadsheet named-range text does not match its ref property.",
        ));
    }
    let scope_value = required_format(node, "scope")?;
    let scope =
        NativeSpreadsheetNamedRangeScope::try_from(scope_value.to_string()).map_err(|error| {
            dump_unsupported(
                &node.path,
                format!("Spreadsheet named-range scope is not replayable: {error}"),
            )
        })?;
    let volatile = parse_bool_format(node, "volatile", false)?;
    Ok(NativeSpreadsheetNamedRange {
        name,
        reference,
        scope,
        comment: node.format.get("comment").cloned(),
        volatile,
    })
}

fn spreadsheet_data_validation(node: &DocumentNode) -> UseResult<NativeSpreadsheetDataValidation> {
    if node.style.is_some() || !node.children.is_empty() {
        return Err(dump_unsupported(
            &node.path,
            "Spreadsheet data validation contains unsupported semantic children or style data.",
        ));
    }
    let allowed = [
        "ref",
        "type",
        "operator",
        "formula1",
        "formula2",
        "allowBlank",
        "showInput",
        "showError",
        "promptTitle",
        "prompt",
        "errorTitle",
        "error",
        "errorStyle",
        "inCellDropdown",
    ];
    if let Some(key) = node
        .format
        .keys()
        .find(|key| !allowed.contains(&key.as_str()))
    {
        return Err(dump_unsupported(
            &node.path,
            format!("Spreadsheet data-validation property '{key}' is not replayable yet."),
        ));
    }
    let validation_type = match required_format(node, "type")? {
        "list" => NativeSpreadsheetDataValidationType::List,
        "whole" => NativeSpreadsheetDataValidationType::Whole,
        "decimal" => NativeSpreadsheetDataValidationType::Decimal,
        "date" => NativeSpreadsheetDataValidationType::Date,
        "time" => NativeSpreadsheetDataValidationType::Time,
        "textLength" => NativeSpreadsheetDataValidationType::TextLength,
        "custom" => NativeSpreadsheetDataValidationType::Custom,
        value => {
            return Err(dump_unsupported(
                &node.path,
                format!("Spreadsheet data-validation type '{value}' is not replayable."),
            ))
        }
    };
    let operator = node
        .format
        .get("operator")
        .map(|value| match value.as_str() {
            "between" => Ok(NativeSpreadsheetDataValidationOperator::Between),
            "notBetween" => Ok(NativeSpreadsheetDataValidationOperator::NotBetween),
            "equal" => Ok(NativeSpreadsheetDataValidationOperator::Equal),
            "notEqual" => Ok(NativeSpreadsheetDataValidationOperator::NotEqual),
            "greaterThan" => Ok(NativeSpreadsheetDataValidationOperator::GreaterThan),
            "greaterThanOrEqual" => Ok(NativeSpreadsheetDataValidationOperator::GreaterThanOrEqual),
            "lessThan" => Ok(NativeSpreadsheetDataValidationOperator::LessThan),
            "lessThanOrEqual" => Ok(NativeSpreadsheetDataValidationOperator::LessThanOrEqual),
            value => Err(dump_unsupported(
                &node.path,
                format!("Spreadsheet data-validation operator '{value}' is not replayable."),
            )),
        })
        .transpose()?;
    let error_style = match node.format.get("errorStyle").map(String::as_str) {
        None | Some("stop") => NativeSpreadsheetDataValidationErrorStyle::Stop,
        Some("warning") => NativeSpreadsheetDataValidationErrorStyle::Warning,
        Some("information") => NativeSpreadsheetDataValidationErrorStyle::Information,
        Some(value) => {
            return Err(dump_unsupported(
                &node.path,
                format!("Spreadsheet data-validation error style '{value}' is not replayable."),
            ))
        }
    };
    let ranges = required_format(node, "ref")?
        .split_ascii_whitespace()
        .map(str::to_string)
        .collect::<Vec<_>>();
    let formula1 = required_format(node, "formula1")?.to_string();
    Ok(NativeSpreadsheetDataValidation {
        validation_type,
        ranges,
        operator,
        formula1,
        formula2: node.format.get("formula2").cloned(),
        allow_blank: replay_bool(node, "allowBlank")?,
        show_input: replay_bool(node, "showInput")?,
        show_error: replay_bool(node, "showError")?,
        prompt_title: node.format.get("promptTitle").cloned(),
        prompt: node.format.get("prompt").cloned(),
        error_title: node.format.get("errorTitle").cloned(),
        error: node.format.get("error").cloned(),
        error_style,
        in_cell_dropdown: replay_bool(node, "inCellDropdown")?,
    })
}

fn required_format<'a>(node: &'a DocumentNode, key: &str) -> UseResult<&'a str> {
    node.format.get(key).map(String::as_str).ok_or_else(|| {
        dump_unsupported(
            &node.path,
            format!("Spreadsheet data validation has no '{key}' property."),
        )
    })
}

fn replay_bool(node: &DocumentNode, key: &str) -> UseResult<bool> {
    match required_format(node, key)? {
        "true" => Ok(true),
        "false" => Ok(false),
        value => Err(dump_unsupported(
            &node.path,
            format!("Spreadsheet data-validation property '{key}' is not boolean: '{value}'."),
        )),
    }
}

fn parse_bool_format(node: &DocumentNode, key: &str, default: bool) -> UseResult<bool> {
    match node.format.get(key).map(String::as_str) {
        None => Ok(default),
        Some("true") => Ok(true),
        Some("false") => Ok(false),
        Some(value) => Err(dump_unsupported(
            &node.path,
            format!("Spreadsheet property '{key}' is not boolean: '{value}'."),
        )),
    }
}

fn spreadsheet_cell_format(cell: &DocumentNode) -> UseResult<Option<NativeSpreadsheetCellFormat>> {
    if cell.node_type != OfficeNodeType::Cell || cell.style.is_some() || !cell.children.is_empty() {
        return Err(dump_unsupported(
            &cell.path,
            "Spreadsheet replay requires a leaf cell without semantic child nodes.",
        ));
    }
    let plain = [
        "column",
        "row",
        "valueType",
        "valuePresent",
        "empty",
        "formula",
        "formulaCached",
        "formulaRef",
        "formulaType",
        "merge",
        "mergeAnchor",
        "dataValidation",
        "validationType",
    ];
    let styled = [
        "styleIndex",
        "baseStyleId",
        "fontId",
        "font",
        "size",
        "fillId",
        "borderId",
        "numberFormatId",
        "numberFormat",
    ];
    if let Some(key) = cell
        .format
        .keys()
        .find(|key| !plain.contains(&key.as_str()) && !styled.contains(&key.as_str()))
    {
        return Err(dump_unsupported(
            &cell.path,
            format!("Spreadsheet cell property '{key}' is not replayable yet."),
        ));
    }
    if !cell.format.contains_key("styleIndex") {
        if let Some(key) = styled.iter().find(|key| cell.format.contains_key(**key)) {
            return Err(dump_unsupported(
                &cell.path,
                format!("Spreadsheet cell has style property '{key}' without a styleIndex."),
            ));
        }
        return Ok(None);
    }
    for (key, expected) in [
        ("baseStyleId", "0"),
        ("fontId", "0"),
        ("font", "Aptos"),
        ("size", "11pt"),
        ("fillId", "0"),
        ("borderId", "0"),
    ] {
        if cell.format.get(key).map(String::as_str) != Some(expected) {
            return Err(dump_unsupported(
                &cell.path,
                format!(
                    "Spreadsheet replay supports only the native default date style; '{key}' is not '{expected}'."
                ),
            ));
        }
    }
    let number_format = cell.format.get("numberFormat").ok_or_else(|| {
        dump_unsupported(
            &cell.path,
            "Spreadsheet styled cell has no numberFormat value.",
        )
    })?;
    let valid_number_format_id = cell
        .format
        .get("numberFormatId")
        .and_then(|value| value.parse::<u32>().ok())
        .is_some_and(|value| value >= 164);
    let valid_style_index = cell
        .format
        .get("styleIndex")
        .and_then(|value| value.parse::<usize>().ok())
        .is_some();
    if number_format != "yyyy-mm-dd" || !valid_number_format_id || !valid_style_index {
        return Err(dump_unsupported(
            &cell.path,
            "Spreadsheet replay supports only the canonical native yyyy-mm-dd import style.",
        ));
    }
    Ok(Some(NativeSpreadsheetCellFormat {
        number_format: Some(number_format.clone()),
        ..NativeSpreadsheetCellFormat::default()
    }))
}

fn spreadsheet_value(cell: &DocumentNode) -> UseResult<SpreadsheetCellValue> {
    if let Some(expression) = cell.format.get("formula") {
        return Ok(SpreadsheetCellValue::Formula {
            expression: expression.clone(),
        });
    }
    match cell.format.get("valueType").map(String::as_str) {
        Some("String") => Ok(SpreadsheetCellValue::Text {
            value: cell.text.clone(),
        }),
        Some("Number") if !cell.text.is_empty() => Ok(SpreadsheetCellValue::Number {
            value: cell.text.clone(),
        }),
        Some("Boolean") => match cell.text.as_str() {
            "true" => Ok(SpreadsheetCellValue::Boolean { value: true }),
            "false" => Ok(SpreadsheetCellValue::Boolean { value: false }),
            _ => Err(dump_unsupported(
                &cell.path,
                "Spreadsheet boolean cell is not normalized to true or false.",
            )),
        },
        Some(value_type) => Err(dump_unsupported(
            &cell.path,
            format!("Spreadsheet value type '{value_type}' is not replayable yet."),
        )),
        None => Err(dump_unsupported(
            &cell.path,
            "Spreadsheet cell has no value type.",
        )),
    }
}

fn emit_presentation(root: &DocumentNode) -> UseResult<Vec<NativeOfficeMutation>> {
    let mut mutations = Vec::new();
    for slide in &root.children {
        if slide.node_type != OfficeNodeType::Slide {
            return Err(dump_unsupported(
                &slide.path,
                "Presentation root contains a non-slide node.",
            ));
        }
        for child in &slide.children {
            match child.node_type {
                OfficeNodeType::Shape => validate_presentation_shape(child)?,
                OfficeNodeType::Table => validate_presentation_table(child)?,
                _ => {
                    return Err(dump_unsupported(
                        &child.path,
                        format!(
                            "Presentation node type '{}' is not exactly replayable yet.",
                            child.node_type.label()
                        ),
                    ))
                }
            }
        }

        let title = slide.children.first().filter(|shape| {
            shape.format.get("id").is_some_and(|value| value == "2")
                && shape
                    .format
                    .get("name")
                    .is_some_and(|value| value == "Title 1")
        });
        let title_text = title.map_or("", |shape| shape.text.as_str());
        let seeded_empty_title = title.is_some() && title_text.is_empty();
        mutations.push(NativeOfficeMutation::AddSlide {
            parent: "/".to_string(),
            title: if seeded_empty_title {
                EMPTY_PRESENTATION_TITLE_SEED.to_string()
            } else {
                title_text.to_string()
            },
        });
        if seeded_empty_title {
            mutations.push(NativeOfficeMutation::SetText {
                path: format!("{}/shape[1]", slide.path),
                text: String::new(),
            });
        }
        for child in slide.children.iter().skip(usize::from(title.is_some())) {
            match child.node_type {
                OfficeNodeType::Shape => mutations.push(NativeOfficeMutation::AddShape {
                    parent: slide.path.clone(),
                    text: child.text.clone(),
                }),
                OfficeNodeType::Table => {
                    emit_presentation_table(child, &slide.path, &mut mutations)?
                }
                _ => {
                    return Err(dump_unsupported(
                        &child.path,
                        "Presentation replay encountered an unsupported node after validation.",
                    ))
                }
            }
        }
    }
    Ok(mutations)
}

fn emit_presentation_table(
    table: &DocumentNode,
    slide_path: &str,
    mutations: &mut Vec<NativeOfficeMutation>,
) -> UseResult<()> {
    let rows = table.children.len();
    let columns = table.children.first().map_or(0, |row| row.children.len());
    mutations.push(NativeOfficeMutation::AddTable {
        parent: slide_path.to_string(),
        rows,
        columns,
    });
    for row in &table.children {
        for cell in &row.children {
            if !cell.text.is_empty() {
                mutations.push(NativeOfficeMutation::SetText {
                    path: cell.path.clone(),
                    text: cell.text.clone(),
                });
            }
        }
    }
    Ok(())
}

fn validate_presentation_table(table: &DocumentNode) -> UseResult<()> {
    if table.style.is_some() || table.children.is_empty() {
        return Err(dump_unsupported(
            &table.path,
            "Presentation replay requires a non-empty basic table.",
        ));
    }
    let columns = table.children[0].children.len();
    if columns == 0 {
        return Err(dump_unsupported(
            &table.path,
            "Presentation replay requires at least one table column.",
        ));
    }
    for row in &table.children {
        if row.node_type != OfficeNodeType::TableRow
            || row.style.is_some()
            || row.children.len() != columns
        {
            return Err(dump_unsupported(
                &row.path,
                "Presentation replay requires a rectangular table with plain rows.",
            ));
        }
        for cell in &row.children {
            if cell.node_type != OfficeNodeType::TableCell || cell.style.is_some() {
                return Err(dump_unsupported(
                    &cell.path,
                    "Presentation replay requires plain table cells.",
                ));
            }
        }
    }
    Ok(())
}

fn validate_presentation_shape(shape: &DocumentNode) -> UseResult<()> {
    if shape.node_type != OfficeNodeType::Shape || shape.style.is_some() {
        return Err(dump_unsupported(
            &shape.path,
            format!(
                "Presentation node type '{}' is not exactly replayable yet.",
                shape.node_type.label()
            ),
        ));
    }
    let [paragraph] = shape.children.as_slice() else {
        return Err(dump_unsupported(
            &shape.path,
            "Presentation replay requires each text shape to contain one paragraph.",
        ));
    };
    if paragraph.node_type != OfficeNodeType::Paragraph
        || paragraph.style.is_some()
        || !paragraph.format.is_empty()
    {
        return Err(dump_unsupported(
            &paragraph.path,
            "Presentation replay requires a plain shape paragraph.",
        ));
    }
    let [run] = paragraph.children.as_slice() else {
        return Err(dump_unsupported(
            &paragraph.path,
            "Presentation replay requires each shape paragraph to contain one run.",
        ));
    };
    let allowed_run_format = ["language", "size"];
    if run.node_type != OfficeNodeType::Run
        || run.style.is_some()
        || !run.children.is_empty()
        || run
            .format
            .keys()
            .any(|key| !allowed_run_format.contains(&key.as_str()))
        || run.text != paragraph.text
        || paragraph.text != shape.text
    {
        return Err(dump_unsupported(
            &run.path,
            "Presentation shape text does not map to one plain run.",
        ));
    }
    Ok(())
}

fn require_plain_node(node: &DocumentNode, expected: OfficeNodeType) -> UseResult<()> {
    if node.node_type != expected || node.style.is_some() || !node.format.is_empty() {
        return Err(dump_unsupported(
            &node.path,
            format!(
                "{} formatting or node metadata is not exactly replayable yet.",
                expected.label()
            ),
        ));
    }
    Ok(())
}

fn first_differing_part(
    source: &NativeOfficePackage,
    replay: &NativeOfficePackage,
) -> Option<String> {
    let names = source
        .part_names()
        .chain(replay.part_names())
        .map(str::to_string)
        .collect::<BTreeSet<_>>();
    names
        .into_iter()
        .find(|name| source.part(name).ok() != replay.part(name).ok())
}

fn valid_sha256(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn dump_unsupported(path: &str, message: impl Into<String>) -> UseError {
    replay_error("use.office.dump_unsupported", message).with_detail("path", path)
}

fn replay_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}
