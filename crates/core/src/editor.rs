use std::path::Path;

use a3s_use_core::UseResult;

use crate::semantic::NativeOfficeDocument;
use crate::{
    template_merge, DocumentKind, NativeOfficePackage, NativeOfficeReplayArtifact,
    NativeOfficeTemplateMergeResult,
};

mod comment;
mod hyperlink;
mod image;
mod part;
mod presentation;
mod raw;
mod spreadsheet;
mod support;
mod text_replace;
mod types;
mod word;

pub(crate) use image::inspect_image;

use support::{
    editor_error, escape_attribute, node_not_found, parse_segments, prefix,
    preserve_space_attribute, qualified, validate_mutation_path,
};

#[cfg(test)]
mod part_tests;

pub use part::{NativeCreatedPart, NativeOfficePartType};
pub use raw::NativeRawXmlPart;
pub use types::{
    NativeBatchResult, NativeCreatedImage, NativeOfficeComment, NativeOfficeCommentPosition,
    NativeOfficeCommentUpdate, NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment,
    NativeOfficeHyperlink, NativeOfficeHyperlinkTarget, NativeOfficeImage, NativeOfficeImageFormat,
    NativeOfficeImageMetadata, NativeOfficeInsertPosition, NativeOfficeMutation,
    NativeOfficeRgbColor, NativeOfficeSwapResult, NativeOfficeTextCase, NativeOfficeTextFormat,
    NativeOfficeTextMatchMode, NativeOfficeTextReplacement, NativeOfficeTextReplacementResult,
    NativeOfficeTextScript, NativeOfficeUnderline, NativeSpreadsheetAutoFilter,
    NativeSpreadsheetBorder, NativeSpreadsheetBorderLine, NativeSpreadsheetBorderStyle,
    NativeSpreadsheetCellFormat, NativeSpreadsheetConditionalFormat,
    NativeSpreadsheetConditionalFormatIconSet, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind, NativeSpreadsheetConditionalFormatTimePeriod,
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
    NativeSpreadsheetDelimitedFormat, NativeSpreadsheetDelimitedImport,
    NativeSpreadsheetDifferentialFormat, NativeSpreadsheetDynamicFilter, NativeSpreadsheetFill,
    NativeSpreadsheetFilterColumn, NativeSpreadsheetFilterCriteria, NativeSpreadsheetFrozenPane,
    NativeSpreadsheetImportResult, NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope,
    NativeSpreadsheetReadingOrder, NativeSpreadsheetSort, NativeSpreadsheetSortDirection,
    NativeSpreadsheetSortKey, NativeSpreadsheetTable, NativeSpreadsheetTableColumn,
    NativeSpreadsheetTableStyle, NativeSpreadsheetVerticalAlignment, SpreadsheetCellValue,
    MAX_NATIVE_OFFICE_FIND_BYTES, MAX_NATIVE_OFFICE_REPLACEMENT_BYTES,
    MAX_NATIVE_OFFICE_TEXT_MATCHES, MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES,
    MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS, MAX_NATIVE_SPREADSHEET_IMPORT_BYTES,
    MAX_NATIVE_SPREADSHEET_IMPORT_CELLS,
};

/// Loss-preserving OOXML editor with transactional in-memory batches.
#[derive(Debug, Clone)]
pub struct NativeOfficeEditor {
    package: NativeOfficePackage,
}

impl NativeOfficeEditor {
    pub async fn create(path: impl AsRef<Path>) -> UseResult<Self> {
        Self::from_package(NativeOfficePackage::create(path).await?)
    }

    pub async fn open(path: impl AsRef<Path>) -> UseResult<Self> {
        let package = NativeOfficePackage::open(path).await?;
        NativeOfficeDocument::from_package(package.clone())?;
        Ok(Self { package })
    }

    pub fn from_package(package: NativeOfficePackage) -> UseResult<Self> {
        NativeOfficeDocument::from_package(package.clone())?;
        Ok(Self { package })
    }

    pub fn package(&self) -> &NativeOfficePackage {
        &self.package
    }

    /// Safely parses and returns an existing OOXML XML part.
    pub fn raw_xml_part(&self, part: &str) -> UseResult<NativeRawXmlPart> {
        raw::inspect(&self.package, part)
    }

    pub fn snapshot(&self) -> UseResult<NativeOfficeDocument> {
        NativeOfficeDocument::from_package(self.package.clone())
    }

    pub fn is_dirty(&self) -> bool {
        self.package.is_dirty()
    }

    pub fn set_text(&mut self, path: impl Into<String>, text: impl Into<String>) -> UseResult<()> {
        self.apply_batch(&[NativeOfficeMutation::SetText {
            path: path.into(),
            text: text.into(),
        }])?;
        Ok(())
    }

    /// Replaces bounded text matches within one semantic document scope.
    ///
    /// Matches may span rich-text runs, while replacement text inherits the
    /// first matched run's formatting. A zero-match operation succeeds and
    /// returns an unchanged receipt.
    pub fn replace_text(
        &mut self,
        path: impl Into<String>,
        replacement: NativeOfficeTextReplacement,
    ) -> UseResult<NativeOfficeTextReplacementResult> {
        let result = self.apply_batch(&[NativeOfficeMutation::ReplaceText {
            path: path.into(),
            replacement,
        }])?;
        result.text_replacements.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office text replacement returned no receipt.",
            )
        })
    }

    /// Applies typed rich-text properties without changing document content.
    pub fn set_text_format(
        &mut self,
        path: impl Into<String>,
        format: NativeOfficeTextFormat,
    ) -> UseResult<()> {
        self.apply_batch(&[NativeOfficeMutation::SetTextFormat {
            path: path.into(),
            format,
        }])?;
        Ok(())
    }

    /// Applies typed Spreadsheet cell presentation properties without changing
    /// cell contents.
    pub fn set_cell_format(
        &mut self,
        path: impl Into<String>,
        format: NativeSpreadsheetCellFormat,
    ) -> UseResult<()> {
        self.apply_batch(&[NativeOfficeMutation::SetCellFormat {
            path: path.into(),
            format,
        }])?;
        Ok(())
    }

    /// Creates or replaces a typed hyperlink on a supported semantic owner.
    pub fn set_hyperlink(
        &mut self,
        path: impl Into<String>,
        hyperlink: NativeOfficeHyperlink,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetHyperlink {
            path: path.into(),
            hyperlink,
        })
    }

    /// Applies a partial typed update to an existing legacy Office comment.
    pub fn set_comment(
        &mut self,
        path: impl Into<String>,
        update: NativeOfficeCommentUpdate,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetComment {
            path: path.into(),
            update,
        })
    }

    /// Sets one Presentation table-grid column width in English Metric Units.
    pub fn set_table_column_width(
        &mut self,
        path: impl Into<String>,
        width_emu: u64,
    ) -> UseResult<()> {
        self.apply_batch(&[NativeOfficeMutation::SetTableColumnWidth {
            path: path.into(),
            width_emu,
        }])?;
        Ok(())
    }

    pub fn set_cell_value(
        &mut self,
        path: impl Into<String>,
        value: SpreadsheetCellValue,
    ) -> UseResult<()> {
        self.apply_batch(&[NativeOfficeMutation::SetCellValue {
            path: path.into(),
            value,
        }])?;
        Ok(())
    }

    /// Calculates every supported Spreadsheet formula and atomically writes
    /// typed cached values and dynamic-array spill cells into the package.
    pub fn recalculate_spreadsheet_formulas(
        &mut self,
    ) -> UseResult<crate::SpreadsheetFormulaCalculation> {
        let result = self.apply_batch(&[NativeOfficeMutation::RecalculateSpreadsheetFormulas])?;
        result
            .spreadsheet_calculations
            .into_iter()
            .next()
            .ok_or_else(|| {
                editor_error(
                    "use.office.batch_validation_failed",
                    "Native Spreadsheet recalculation returned no calculation receipt.",
                )
            })
    }

    /// Adds one complete typed Spreadsheet ListObject table.
    pub fn add_spreadsheet_table(
        &mut self,
        sheet: impl Into<String>,
        table: NativeSpreadsheetTable,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddSpreadsheetTable {
            sheet: sheet.into(),
            table,
        })
    }

    /// Replaces the typed structure and style of one Spreadsheet table.
    pub fn set_spreadsheet_table(
        &mut self,
        path: impl Into<String>,
        table: NativeSpreadsheetTable,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetSpreadsheetTable {
            path: path.into(),
            table,
        })
    }

    /// Adds the single typed worksheet AutoFilter.
    pub fn add_spreadsheet_auto_filter(
        &mut self,
        sheet: impl Into<String>,
        filter: NativeSpreadsheetAutoFilter,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddSpreadsheetAutoFilter {
            sheet: sheet.into(),
            filter,
        })
    }

    /// Replaces one worksheet AutoFilter and all supported column criteria.
    pub fn set_spreadsheet_auto_filter(
        &mut self,
        path: impl Into<String>,
        filter: NativeSpreadsheetAutoFilter,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetSpreadsheetAutoFilter {
            path: path.into(),
            filter,
        })
    }

    /// Stably reorders physical Spreadsheet records and persists typed sort state.
    pub fn sort_spreadsheet_range(
        &mut self,
        path: impl Into<String>,
        sort: NativeSpreadsheetSort,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SortSpreadsheetRange {
            path: path.into(),
            sort,
        })
    }

    /// Imports bounded CSV or TSV content into one Spreadsheet worksheet.
    pub fn import_spreadsheet_delimited(
        &mut self,
        sheet: impl Into<String>,
        import: NativeSpreadsheetDelimitedImport,
    ) -> UseResult<NativeSpreadsheetImportResult> {
        let result = self.apply_batch(&[NativeOfficeMutation::ImportSpreadsheetDelimited {
            sheet: sheet.into(),
            import,
        }])?;
        result
            .spreadsheet_imports
            .into_iter()
            .next()
            .ok_or_else(|| {
                editor_error(
                    "use.office.batch_validation_failed",
                    "Native Spreadsheet import returned no receipt.",
                )
            })
    }

    /// Creates or replaces one canonical frozen pane on a Spreadsheet sheet.
    pub fn set_spreadsheet_frozen_pane(
        &mut self,
        sheet: impl Into<String>,
        pane: NativeSpreadsheetFrozenPane,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetSpreadsheetFrozenPane {
            sheet: sheet.into(),
            pane,
        })
    }

    /// Adds one complete typed Spreadsheet defined name.
    pub fn add_named_range(
        &mut self,
        named_range: NativeSpreadsheetNamedRange,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddNamedRange { named_range })
    }

    /// Replaces one existing Spreadsheet defined name completely.
    pub fn set_named_range(
        &mut self,
        path: impl Into<String>,
        named_range: NativeSpreadsheetNamedRange,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetNamedRange {
            path: path.into(),
            named_range,
        })
    }

    /// Adds one complete typed Spreadsheet conditional-formatting rule.
    pub fn add_conditional_format(
        &mut self,
        sheet: impl Into<String>,
        conditional_format: NativeSpreadsheetConditionalFormat,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddConditionalFormat {
            sheet: sheet.into(),
            conditional_format,
        })
    }

    /// Replaces one existing Spreadsheet conditional-formatting rule while
    /// retaining its worksheet priority.
    pub fn set_conditional_format(
        &mut self,
        path: impl Into<String>,
        conditional_format: NativeSpreadsheetConditionalFormat,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetConditionalFormat {
            path: path.into(),
            conditional_format,
        })
    }

    /// Adds one complete typed Spreadsheet data-validation rule.
    pub fn add_data_validation(
        &mut self,
        sheet: impl Into<String>,
        validation: NativeSpreadsheetDataValidation,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddDataValidation {
            sheet: sheet.into(),
            validation,
        })
    }

    /// Replaces one existing Spreadsheet data-validation rule completely.
    pub fn set_data_validation(
        &mut self,
        path: impl Into<String>,
        validation: NativeSpreadsheetDataValidation,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::SetDataValidation {
            path: path.into(),
            validation,
        })
    }

    /// Adds one normalized SpreadsheetML merged-cell range.
    ///
    /// Repeating the exact range is idempotent. A geometrically overlapping
    /// range or a range intersecting a Spreadsheet table fails closed.
    pub fn merge_cells(&mut self, path: impl Into<String>) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::MergeCells { path: path.into() })
    }

    /// Removes one exact SpreadsheetML merged-cell range.
    ///
    /// A non-exact range intersecting any merge is rejected with the exact
    /// references instead of destructively sweeping or partially changing it.
    pub fn unmerge_cells(&mut self, path: impl Into<String>) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::UnmergeCells { path: path.into() })
    }

    pub fn add_paragraph(
        &mut self,
        parent: impl Into<String>,
        text: impl Into<String>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddParagraph {
            parent: parent.into(),
            text: text.into(),
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office paragraph mutation returned no path.",
            )
        })
    }

    pub fn add_slide(
        &mut self,
        parent: impl Into<String>,
        title: impl Into<String>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddSlide {
            parent: parent.into(),
            title: title.into(),
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office slide mutation returned no path.",
            )
        })
    }

    pub fn add_table(
        &mut self,
        parent: impl Into<String>,
        rows: usize,
        columns: usize,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddTable {
            parent: parent.into(),
            rows,
            columns,
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office table mutation returned no path.",
            )
        })
    }

    pub fn add_table_row(
        &mut self,
        parent: impl Into<String>,
        columns: Option<usize>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddTableRow {
            parent: parent.into(),
            columns,
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office table-row mutation returned no path.",
            )
        })
    }

    /// Inserts one Presentation table column at a zero-based grid position.
    ///
    /// Omitting `index` appends the column. Every row receives one new cell so
    /// the DrawingML table grid remains rectangular.
    pub fn add_table_column(
        &mut self,
        parent: impl Into<String>,
        index: Option<usize>,
        text: impl Into<String>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddTableColumn {
            parent: parent.into(),
            index,
            text: text.into(),
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office table-column mutation returned no path.",
            )
        })
    }

    pub fn add_table_cell(
        &mut self,
        parent: impl Into<String>,
        text: impl Into<String>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddTableCell {
            parent: parent.into(),
            text: text.into(),
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office table-cell mutation returned no path.",
            )
        })
    }

    pub fn add_shape(
        &mut self,
        parent: impl Into<String>,
        text: impl Into<String>,
    ) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddShape {
            parent: parent.into(),
            text: text.into(),
        }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office shape mutation returned no path.",
            )
        })
    }

    /// Embeds a validated PNG, JPEG, or GIF as a native DrawingML picture.
    pub fn add_image(
        &mut self,
        parent: impl Into<String>,
        image: NativeOfficeImage,
    ) -> UseResult<NativeCreatedImage> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddImage {
            parent: parent.into(),
            image,
        }])?;
        result.created_images.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office image mutation returned no creation receipt.",
            )
        })
    }

    /// Adds a typed legacy comment to a Word paragraph/run, Spreadsheet cell,
    /// or Presentation slide.
    pub fn add_comment(
        &mut self,
        parent: impl Into<String>,
        comment: NativeOfficeComment,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::AddComment {
            parent: parent.into(),
            comment,
        })
    }

    /// Creates a known XML part with its content type and owner relationship.
    pub fn add_part(
        &mut self,
        parent: impl Into<String>,
        part_type: NativeOfficePartType,
    ) -> UseResult<NativeCreatedPart> {
        let result = self.apply_batch(&[NativeOfficeMutation::AddPart {
            parent: parent.into(),
            part_type,
        }])?;
        result.created_parts.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office part mutation returned no creation receipt.",
            )
        })
    }

    pub fn remove(&mut self, path: impl Into<String>) -> UseResult<String> {
        let result = self.apply_batch(&[NativeOfficeMutation::Remove { path: path.into() }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office remove mutation returned no path.",
            )
        })
    }

    pub fn add_worksheet(&mut self, name: impl Into<String>) -> UseResult<String> {
        let result =
            self.apply_batch(&[NativeOfficeMutation::AddWorksheet { name: name.into() }])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office worksheet mutation returned no path.",
            )
        })
    }

    pub fn insert_rows(
        &mut self,
        sheet: impl Into<String>,
        start: u32,
        count: u32,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::InsertRows {
            sheet: sheet.into(),
            start,
            count,
        })
    }

    pub fn delete_rows(
        &mut self,
        sheet: impl Into<String>,
        start: u32,
        count: u32,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::DeleteRows {
            sheet: sheet.into(),
            start,
            count,
        })
    }

    pub fn insert_columns(
        &mut self,
        sheet: impl Into<String>,
        start: impl Into<String>,
        count: u32,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::InsertColumns {
            sheet: sheet.into(),
            start: start.into(),
            count,
        })
    }

    pub fn delete_columns(
        &mut self,
        sheet: impl Into<String>,
        start: impl Into<String>,
        count: u32,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::DeleteColumns {
            sheet: sheet.into(),
            start: start.into(),
            count,
        })
    }

    pub fn rename_worksheet(
        &mut self,
        path: impl Into<String>,
        name: impl Into<String>,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::RenameWorksheet {
            path: path.into(),
            name: name.into(),
        })
    }

    pub fn move_worksheet(
        &mut self,
        path: impl Into<String>,
        position: usize,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::MoveWorksheet {
            path: path.into(),
            position,
        })
    }

    pub fn copy_worksheet(
        &mut self,
        path: impl Into<String>,
        name: impl Into<String>,
        position: Option<usize>,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::CopyWorksheet {
            path: path.into(),
            name: name.into(),
            position,
        })
    }

    pub fn move_node(
        &mut self,
        path: impl Into<String>,
        target_parent: Option<String>,
        position: Option<NativeOfficeInsertPosition>,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::Move {
            path: path.into(),
            target_parent,
            position,
        })
    }

    pub fn copy_node(
        &mut self,
        path: impl Into<String>,
        target_parent: Option<String>,
        position: Option<NativeOfficeInsertPosition>,
        name: Option<String>,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::Copy {
            path: path.into(),
            target_parent,
            position,
            name,
        })
    }

    pub fn swap_nodes(
        &mut self,
        path: impl Into<String>,
        with: impl Into<String>,
    ) -> UseResult<NativeOfficeSwapResult> {
        let result = self.apply_batch(&[NativeOfficeMutation::Swap {
            path: path.into(),
            with: with.into(),
        }])?;
        result.swaps.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office swap mutation returned no swap receipt.",
            )
        })
    }

    /// Replaces an existing, non-OPC-metadata XML part transactionally.
    pub fn replace_xml_part(
        &mut self,
        part: impl Into<String>,
        xml: impl Into<String>,
    ) -> UseResult<String> {
        self.single_path(NativeOfficeMutation::ReplaceXmlPart {
            part: part.into(),
            xml: xml.into(),
        })
    }

    fn single_path(&mut self, mutation: NativeOfficeMutation) -> UseResult<String> {
        let result = self.apply_batch(&[mutation])?;
        result.paths.into_iter().next().ok_or_else(|| {
            editor_error(
                "use.office.batch_validation_failed",
                "Native Office mutation returned no path.",
            )
        })
    }

    pub fn apply_batch(
        &mut self,
        mutations: &[NativeOfficeMutation],
    ) -> UseResult<NativeBatchResult> {
        if mutations.is_empty() {
            return Ok(NativeBatchResult {
                applied: 0,
                paths: Vec::new(),
                swaps: Vec::new(),
                created_parts: Vec::new(),
                created_images: Vec::new(),
                text_replacements: Vec::new(),
                spreadsheet_imports: Vec::new(),
                spreadsheet_calculations: Vec::new(),
            });
        }
        let original = self.package.clone();
        let mut paths = Vec::with_capacity(mutations.len());
        let mut swaps = Vec::new();
        let mut created_parts = Vec::new();
        let mut created_images = Vec::new();
        let mut text_replacements = Vec::new();
        let mut spreadsheet_imports = Vec::new();
        let mut spreadsheet_calculations = Vec::new();
        for mutation in mutations {
            let mut created_part = None;
            let mut created_image = None;
            let mut swap = None;
            let mut text_replacement = None;
            let mut spreadsheet_import = None;
            let mut spreadsheet_calculation = None;
            let result = match mutation {
                NativeOfficeMutation::ReplaceText { path, replacement } => {
                    text_replace::replace(&mut self.package, path, replacement).map(|receipt| {
                        let path = receipt.path.clone();
                        text_replacement = Some(receipt);
                        path
                    })
                }
                NativeOfficeMutation::SetText { path, text } => {
                    set_text(&mut self.package, path, text).map(|()| path.clone())
                }
                NativeOfficeMutation::SetTextFormat { path, format } => format
                    .validate()
                    .and_then(|()| match self.package.kind() {
                        DocumentKind::Word => {
                            word::set_text_format(&mut self.package, path, format)
                        }
                        DocumentKind::Spreadsheet => {
                            spreadsheet::set_text_format(&mut self.package, path, format)
                        }
                        DocumentKind::Presentation => {
                            presentation::set_text_format(&mut self.package, path, format)
                        }
                    })
                    .map(|()| path.clone()),
                NativeOfficeMutation::SetCellFormat { path, format } => format
                    .validate()
                    .and_then(|()| spreadsheet::set_cell_format(&mut self.package, path, format))
                    .map(|()| path.clone()),
                NativeOfficeMutation::SetHyperlink { path, hyperlink } => hyperlink
                    .validate()
                    .and_then(|()| hyperlink::set(&mut self.package, path, hyperlink)),
                NativeOfficeMutation::SetComment { path, update } => update
                    .validate()
                    .and_then(|()| comment::set(&mut self.package, path, update)),
                NativeOfficeMutation::SetTableColumnWidth { path, width_emu } => {
                    presentation::set_table_column_width(&mut self.package, path, *width_emu)
                        .map(|()| path.clone())
                }
                NativeOfficeMutation::SetCellValue { path, value } => {
                    spreadsheet::set_cell_value(&mut self.package, path, value)
                        .map(|()| path.clone())
                }
                NativeOfficeMutation::RecalculateSpreadsheetFormulas => {
                    spreadsheet::recalculate_formulas(&mut self.package).map(|receipt| {
                        spreadsheet_calculation = Some(receipt);
                        "/".to_string()
                    })
                }
                NativeOfficeMutation::AddSpreadsheetTable { sheet, table } => {
                    spreadsheet::add_table(&mut self.package, sheet, table)
                }
                NativeOfficeMutation::SetSpreadsheetTable { path, table } => {
                    spreadsheet::set_table(&mut self.package, path, table)
                }
                NativeOfficeMutation::AddSpreadsheetAutoFilter { sheet, filter } => {
                    spreadsheet::add_auto_filter(&mut self.package, sheet, filter)
                }
                NativeOfficeMutation::SetSpreadsheetAutoFilter { path, filter } => {
                    spreadsheet::set_auto_filter(&mut self.package, path, filter)
                }
                NativeOfficeMutation::SortSpreadsheetRange { path, sort } => {
                    spreadsheet::sort_range(&mut self.package, path, sort)
                }
                NativeOfficeMutation::ImportSpreadsheetDelimited { sheet, import } => {
                    spreadsheet::import_delimited(&mut self.package, sheet, import).map(|receipt| {
                        let path = receipt.path.clone();
                        spreadsheet_import = Some(receipt);
                        path
                    })
                }
                NativeOfficeMutation::SetSpreadsheetFrozenPane { sheet, pane } => {
                    spreadsheet::set_frozen_pane(&mut self.package, sheet, pane)
                }
                NativeOfficeMutation::AddNamedRange { named_range } => {
                    spreadsheet::add_named_range(&mut self.package, named_range)
                }
                NativeOfficeMutation::SetNamedRange { path, named_range } => {
                    spreadsheet::set_named_range(&mut self.package, path, named_range)
                }
                NativeOfficeMutation::AddConditionalFormat {
                    sheet,
                    conditional_format,
                } => spreadsheet::add_conditional_format(
                    &mut self.package,
                    sheet,
                    conditional_format,
                ),
                NativeOfficeMutation::SetConditionalFormat {
                    path,
                    conditional_format,
                } => {
                    spreadsheet::set_conditional_format(&mut self.package, path, conditional_format)
                }
                NativeOfficeMutation::AddDataValidation { sheet, validation } => {
                    spreadsheet::add_data_validation(&mut self.package, sheet, validation)
                }
                NativeOfficeMutation::SetDataValidation { path, validation } => {
                    spreadsheet::set_data_validation(&mut self.package, path, validation)
                }
                NativeOfficeMutation::MergeCells { path } => {
                    spreadsheet::merge_cells(&mut self.package, path)
                }
                NativeOfficeMutation::UnmergeCells { path } => {
                    spreadsheet::unmerge_cells(&mut self.package, path)
                }
                NativeOfficeMutation::AddParagraph { parent, text } => {
                    word::add_paragraph(&mut self.package, parent, text)
                }
                NativeOfficeMutation::AddTable {
                    parent,
                    rows,
                    columns,
                } => match self.package.kind() {
                    DocumentKind::Presentation => {
                        presentation::add_table(&mut self.package, parent, *rows, *columns)
                    }
                    _ => word::add_table(&mut self.package, parent, *rows, *columns),
                },
                NativeOfficeMutation::AddTableRow { parent, columns } => {
                    match self.package.kind() {
                        DocumentKind::Presentation => {
                            presentation::add_table_row(&mut self.package, parent, *columns)
                        }
                        _ => word::add_table_row(&mut self.package, parent, *columns),
                    }
                }
                NativeOfficeMutation::AddTableColumn {
                    parent,
                    index,
                    text,
                } => presentation::add_table_column(&mut self.package, parent, *index, text),
                NativeOfficeMutation::AddTableCell { parent, text } => match self.package.kind() {
                    DocumentKind::Presentation => {
                        presentation::add_table_cell(&mut self.package, parent, text)
                    }
                    _ => word::add_table_cell(&mut self.package, parent, text),
                },
                NativeOfficeMutation::AddSlide { parent, title } => {
                    presentation::add_slide(&mut self.package, parent, title)
                }
                NativeOfficeMutation::AddShape { parent, text } => {
                    presentation::add_shape(&mut self.package, parent, text)
                }
                NativeOfficeMutation::AddImage { parent, image } => {
                    image::add(&mut self.package, parent, image).map(|created| {
                        let path = created.path.clone();
                        created_image = Some(created);
                        path
                    })
                }
                NativeOfficeMutation::AddComment { parent, comment } => comment
                    .validate()
                    .and_then(|()| comment::add(&mut self.package, parent, comment)),
                NativeOfficeMutation::AddPart { parent, part_type } => {
                    part::add(&mut self.package, parent, *part_type).map(|created| {
                        let path = created.path.clone();
                        created_part = Some(created);
                        path
                    })
                }
                NativeOfficeMutation::AddWorksheet { name } => {
                    spreadsheet::add_worksheet(&mut self.package, name)
                }
                NativeOfficeMutation::InsertRows {
                    sheet,
                    start,
                    count,
                } => spreadsheet::insert_rows(&mut self.package, sheet, *start, *count),
                NativeOfficeMutation::DeleteRows {
                    sheet,
                    start,
                    count,
                } => spreadsheet::delete_rows(&mut self.package, sheet, *start, *count),
                NativeOfficeMutation::InsertColumns {
                    sheet,
                    start,
                    count,
                } => spreadsheet::insert_columns(&mut self.package, sheet, start, *count),
                NativeOfficeMutation::DeleteColumns {
                    sheet,
                    start,
                    count,
                } => spreadsheet::delete_columns(&mut self.package, sheet, start, *count),
                NativeOfficeMutation::RenameWorksheet { path, name } => {
                    spreadsheet::rename_worksheet(&mut self.package, path, name)
                }
                NativeOfficeMutation::MoveWorksheet { path, position } => {
                    spreadsheet::move_worksheet(&mut self.package, path, *position)
                }
                NativeOfficeMutation::CopyWorksheet {
                    path,
                    name,
                    position,
                } => spreadsheet::copy_worksheet(&mut self.package, path, name, *position),
                NativeOfficeMutation::Move {
                    path,
                    target_parent,
                    position,
                } => move_node(
                    &mut self.package,
                    path,
                    target_parent.as_deref(),
                    position.as_ref(),
                ),
                NativeOfficeMutation::Copy {
                    path,
                    target_parent,
                    position,
                    name,
                } => copy_node(
                    &mut self.package,
                    path,
                    target_parent.as_deref(),
                    position.as_ref(),
                    name.as_deref(),
                ),
                NativeOfficeMutation::Swap { path, with } => {
                    swap_nodes(&mut self.package, path, with).map(|receipt| {
                        let primary = receipt.first.clone();
                        swap = Some(receipt);
                        primary
                    })
                }
                NativeOfficeMutation::ReplaceXmlPart { part, xml } => {
                    raw::replace(&mut self.package, part, xml)
                }
                NativeOfficeMutation::Remove { path } => {
                    remove_node(&mut self.package, path).map(|()| path.clone())
                }
            };
            match result {
                Ok(path) => {
                    paths.push(path);
                    if let Some(created) = created_part {
                        created_parts.push(created);
                    }
                    if let Some(created) = created_image {
                        created_images.push(created);
                    }
                    if let Some(receipt) = swap {
                        swaps.push(receipt);
                    }
                    if let Some(receipt) = text_replacement {
                        text_replacements.push(receipt);
                    }
                    if let Some(receipt) = spreadsheet_import {
                        spreadsheet_imports.push(receipt);
                    }
                    if let Some(receipt) = spreadsheet_calculation {
                        spreadsheet_calculations.push(receipt);
                    }
                }
                Err(error) => {
                    self.package = original;
                    return Err(error);
                }
            }
        }
        if let Err(error) = NativeOfficeDocument::from_package(self.package.clone()) {
            self.package = original;
            return Err(editor_error(
                "use.office.batch_validation_failed",
                format!("Native Office batch failed post-mutation validation: {error}"),
            ));
        }
        Ok(NativeBatchResult {
            applied: paths.len(),
            paths,
            swaps,
            created_parts,
            created_images,
            text_replacements,
            spreadsheet_imports,
            spreadsheet_calculations,
        })
    }

    /// Applies a dump-produced replay artifact with blank-base and final-state
    /// fingerprint checks around the existing atomic mutation boundary.
    pub fn apply_replay(
        &mut self,
        artifact: &NativeOfficeReplayArtifact,
    ) -> UseResult<NativeBatchResult> {
        artifact.validate()?;
        if artifact.document_kind != self.package.kind() {
            return Err(editor_error(
                "use.office.replay_kind_mismatch",
                format!(
                    "Native Office replay targets {:?}, but the document is {:?}.",
                    artifact.document_kind,
                    self.package.kind()
                ),
            ));
        }

        let observed_base = self.package.content_sha256();
        if observed_base != artifact.base_sha256 {
            return Err(editor_error(
                "use.office.replay_base_mismatch",
                "Native Office replay requires the exact blank package recorded by the artifact.",
            )
            .with_detail("expectedSha256", artifact.base_sha256.clone())
            .with_detail("observedSha256", observed_base));
        }

        let original = self.package.clone();
        let result = self.apply_batch(&artifact.mutations)?;
        let observed_result = self.package.content_sha256();
        if observed_result != artifact.result_sha256 {
            self.package = original;
            return Err(editor_error(
                "use.office.replay_result_mismatch",
                "Native Office replay did not produce the package fingerprint recorded by the artifact; all mutations were rolled back.",
            )
            .with_detail("expectedSha256", artifact.result_sha256.clone())
            .with_detail("observedSha256", observed_result));
        }
        Ok(result)
    }

    /// Replaces `{{key}}` placeholders across the document's native OOXML text
    /// surfaces. The in-memory package is restored if any part cannot be
    /// edited losslessly or if the resulting package fails semantic validation.
    pub fn merge_template(
        &mut self,
        data: &serde_json::Value,
    ) -> UseResult<NativeOfficeTemplateMergeResult> {
        let original = self.package.clone();
        let result = match template_merge::merge(&mut self.package, data) {
            Ok(result) => result,
            Err(error) => {
                self.package = original;
                return Err(error);
            }
        };
        if let Err(error) = NativeOfficeDocument::from_package(self.package.clone()) {
            self.package = original;
            return Err(editor_error(
                "use.office.template_validation_failed",
                format!("Native Office template merge failed post-mutation validation: {error}"),
            ));
        }
        Ok(result)
    }

    pub async fn save(&mut self) -> UseResult<()> {
        self.package.save().await
    }

    pub async fn save_as(&mut self, path: impl AsRef<Path>) -> UseResult<()> {
        self.package.save_as(path).await
    }

    pub async fn save_as_new(&mut self, path: impl AsRef<Path>) -> UseResult<()> {
        self.package.save_as_new(path).await
    }
}

fn set_text(package: &mut NativeOfficePackage, path: &str, text: &str) -> UseResult<()> {
    validate_mutation_path(path)?;
    match package.kind() {
        DocumentKind::Word => word::set_text(package, path, text),
        DocumentKind::Spreadsheet => spreadsheet::set_text(package, path, text),
        DocumentKind::Presentation => presentation::set_text(package, path, text),
    }
}

fn remove_node(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    validate_mutation_path(path)?;
    let node_type = NativeOfficeDocument::from_package(package.clone())?
        .get(path, 0)?
        .node_type;
    if node_type == crate::OfficeNodeType::Picture {
        return image::remove(package, path);
    }
    if node_type == crate::OfficeNodeType::Hyperlink {
        return hyperlink::remove(package, path);
    }
    if node_type == crate::OfficeNodeType::Comment {
        return comment::remove(package, path);
    }
    match package.kind() {
        DocumentKind::Word => word::remove(package, path),
        DocumentKind::Spreadsheet => spreadsheet::remove(package, path),
        DocumentKind::Presentation => presentation::remove(package, path),
    }
}

fn move_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    validate_mutation_path(path)?;
    match package.kind() {
        DocumentKind::Word => word::move_node(package, path, target_parent, position),
        DocumentKind::Spreadsheet => spreadsheet::move_node(package, path, target_parent, position),
        DocumentKind::Presentation => {
            presentation::move_node(package, path, target_parent, position)
        }
    }
}

fn copy_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
    name: Option<&str>,
) -> UseResult<String> {
    validate_mutation_path(path)?;
    match package.kind() {
        DocumentKind::Word => word::copy_node(package, path, target_parent, position, name),
        DocumentKind::Spreadsheet => {
            spreadsheet::copy_node(package, path, target_parent, position, name)
        }
        DocumentKind::Presentation => {
            presentation::copy_node(package, path, target_parent, position, name)
        }
    }
}

fn swap_nodes(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    validate_mutation_path(path)?;
    validate_mutation_path(with)?;
    match package.kind() {
        DocumentKind::Word => word::swap_nodes(package, path, with),
        DocumentKind::Spreadsheet => spreadsheet::swap_nodes(package, path, with),
        DocumentKind::Presentation => presentation::swap_nodes(package, path, with),
    }
}
