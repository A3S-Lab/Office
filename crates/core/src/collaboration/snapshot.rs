//! Import a package into one replica, and export a package from that replica.
//!
//! OOXML bytes are not a second document. Import runs once. Later edits are
//! replica frames. A host save writes those frames back into the package.
//! While that replica is live, `office_open` must not hold another session
//! for the same file.

use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use a3s_use_core::UseResult;
use yrs::{
    Array, ClientID, Doc, Map, OffsetKind, Options, ReadTxn, StateVector, Transact, Xml,
    XmlElementPrelim, XmlFragment, XmlTextPrelim,
};

use super::{
    collaboration_error, NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationCreateRequest, NativeOfficeCollaborationMode,
    NativeOfficeCollaborationStore, NATIVE_OFFICE_COLLABORATION_NAMESPACE,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

const SOURCE_PACKAGE_FILE: &str = "source-package";

const DOCUMENT_PART: &str = "word/document.xml";

/// A document replica created from a package, or the replica that package
/// already owned.
pub struct DocumentSnapshotImport {
    pub store: NativeOfficeCollaborationStore,
    pub created: bool,
}

/// Read `word/document.xml` once and create the document replica.
///
/// A second call for the same package returns the existing replica and does
/// not import the package text again.
pub fn import_document_snapshot(
    package_path: &Path,
    parts: &BTreeMap<String, Vec<u8>>,
    store_path: &Path,
    artifact_id: &str,
) -> UseResult<DocumentSnapshotImport> {
    if let Some(bound) = bound_store_path(package_path)? {
        if !bound.join("manifest.json").is_file() {
            return Err(collaboration_error(
                "office.collaboration.replica_missing",
                format!(
                    "Package '{}' is bound to replica '{}', which is not live.",
                    package_path.display(),
                    bound.display()
                ),
            ));
        }
        return Ok(DocumentSnapshotImport {
            store: NativeOfficeCollaborationStore::open(bound)?,
            created: false,
        });
    }
    let xml = document_xml(parts)?;
    let paragraphs = read_document_paragraphs(&xml)?;
    let update = bootstrap_document_update(artifact_id, &paragraphs)?;
    let store = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: store_path.to_path_buf(),
        artifact_id: artifact_id.to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "snapshot-import".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "import-document-snapshot".to_owned(),
        namespace: None,
        client_id: Some(900_070),
        initial_update: Some(update),
    })?;
    remember_binding(package_path, store.path())?;
    Ok(DocumentSnapshotImport {
        store,
        created: true,
    })
}

/// Refuse `office_open` while this package's replica is the live document.
pub fn live_replica_blocks_office_session(package_path: &Path) -> UseResult<()> {
    let Some(bound) = bound_store_path(package_path)? else {
        return Ok(());
    };
    if !bound.join("manifest.json").is_file() {
        return Ok(());
    }
    let recorded = fs::read_to_string(bound.join(SOURCE_PACKAGE_FILE)).unwrap_or_default();
    if !paths_match(recorded.trim(), package_path) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.replica_live",
        format!(
            "Package '{}' is a snapshot of live replica '{}'.",
            package_path.display(),
            bound.display()
        ),
    )
    .with_suggestion(
        "Edit the collaboration replica. Save exports a package snapshot; office_open does not open a second session.",
    ))
}

/// Copy `parts` and write replica paragraph text into `word/document.xml`.
/// Every other part is returned unchanged. A paragraph with no `paraId`, or
/// with more than one text run, fails closed so unsupported markup is not
/// rewritten.
pub fn write_document_snapshot(
    mut parts: BTreeMap<String, Vec<u8>>,
    paragraphs: &[(&str, &str)],
) -> UseResult<BTreeMap<String, Vec<u8>>> {
    let original = parts.get(DOCUMENT_PART).cloned().ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A document snapshot requires word/document.xml.",
        )
    })?;
    let mut xml = String::from_utf8(original).map_err(|_| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            "word/document.xml is not UTF-8.",
        )
    })?;
    for (paragraph_id, text) in paragraphs {
        xml = replace_paragraph_text(&xml, paragraph_id, text)?;
    }
    parts.insert(DOCUMENT_PART.to_owned(), xml.into_bytes());
    Ok(parts)
}

/// Write plain-text cell values into worksheet parts.
///
/// Only an `inlineStr` cell with one text run, or a `t="str"` cell with one
/// value, is rewritten. Shared-string indexes, formulas, and numeric cells
/// fail closed so one cell cannot change another.
pub fn write_spreadsheet_snapshot(
    mut parts: BTreeMap<String, Vec<u8>>,
    cells: &[(&str, u32, u32, &str)],
) -> UseResult<BTreeMap<String, Vec<u8>>> {
    let mut updated = BTreeMap::<String, String>::new();
    for (part, row, column, text) in cells {
        let current = updated.get(*part).cloned().or_else(|| {
            parts
                .get(*part)
                .and_then(|bytes| String::from_utf8(bytes.clone()).ok())
        });
        let Some(xml) = current else {
            return Err(collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Spreadsheet part '{part}' is missing or not UTF-8."),
            ));
        };
        updated.insert(
            (*part).to_owned(),
            replace_inline_cell(&xml, *row, *column, text)?,
        );
    }
    for (part, xml) in updated {
        parts.insert(part, xml.into_bytes());
    }
    Ok(parts)
}

/// Write one shape's text into a slide part.
///
/// The shape is the `p:sp` whose `cNvPr` id matches. It must contain one
/// `a:t` run. A second run fails closed.
pub fn write_presentation_snapshot(
    mut parts: BTreeMap<String, Vec<u8>>,
    shapes: &[(&str, &str, &str)],
) -> UseResult<BTreeMap<String, Vec<u8>>> {
    let mut updated = BTreeMap::<String, String>::new();
    for (part, shape_id, text) in shapes {
        let current = updated.get(*part).cloned().or_else(|| {
            parts
                .get(*part)
                .and_then(|bytes| String::from_utf8(bytes.clone()).ok())
        });
        let Some(xml) = current else {
            return Err(collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Presentation part '{part}' is missing or not UTF-8."),
            ));
        };
        updated.insert(
            (*part).to_owned(),
            replace_shape_text(&xml, shape_id, text)?,
        );
    }
    for (part, xml) in updated {
        parts.insert(part, xml.into_bytes());
    }
    Ok(parts)
}

/// Write form-field and FreeText annotation literals into an uncompressed PDF.
///
/// A field is `/T (id)` followed by `/V (value)`. An annotation is `/NM (id)`
/// followed by `/Contents (value)`. Indirect objects, hex strings, and files
/// that are not UTF-8 fail closed. This does not paint glyphs onto a page.
pub fn write_pdf_snapshot(
    pdf: &[u8],
    fields: &[(&str, &str)],
    annotations: &[(&str, &str)],
) -> UseResult<Vec<u8>> {
    let mut text = String::from_utf8(pdf.to_vec()).map_err(|_| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A PDF snapshot requires UTF-8 bytes with literal field and annotation strings.",
        )
    })?;
    for (field_id, value) in fields {
        text = replace_pdf_literal(&text, "/T", field_id, "/V", value)?;
    }
    for (annotation_id, value) in annotations {
        text = replace_pdf_literal(&text, "/NM", annotation_id, "/Contents", value)?;
    }
    Ok(text.into_bytes())
}

fn replace_paragraph_text(xml: &str, paragraph_id: &str, text: &str) -> UseResult<String> {
    if paragraph_id.is_empty() || paragraph_id.contains('"') {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A document snapshot paragraph id is invalid.",
        ));
    }
    let marker = format!("paraId=\"{paragraph_id}\"");
    let Some(marker_at) = xml.find(&marker) else {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("word/document.xml has no paragraph '{paragraph_id}'."),
        ));
    };
    let start = paragraph_start(xml, marker_at).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Paragraph '{paragraph_id}' is not inside a w:p element."),
        )
    })?;
    let end_rel = xml[start..].find("</w:p>").ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Paragraph '{paragraph_id}' is missing its end tag."),
        )
    })?;
    let end = start + end_rel + "</w:p>".len();
    let paragraph = &xml[start..end];
    let open = paragraph
        .find("<w:t")
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let open_end = paragraph[open..]
        .find('>')
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let close = paragraph[open + open_end..]
        .find("</w:t>")
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let text_start = open + open_end + 1;
    let text_end = open + open_end + close;
    if paragraph[text_end + "</w:t>".len()..].contains("<w:t") {
        return Err(text_run_error(paragraph_id));
    }
    let mut next = String::with_capacity(xml.len() + text.len());
    next.push_str(&xml[..start]);
    next.push_str(&paragraph[..text_start]);
    next.push_str(&escape_xml_text(text));
    next.push_str(&paragraph[text_end..]);
    next.push_str(&xml[end..]);
    Ok(next)
}

fn paragraph_start(xml: &str, marker_at: usize) -> Option<usize> {
    let mut cursor = marker_at;
    loop {
        let start = xml[..cursor].rfind("<w:p")?;
        let rest = &xml[start..];
        if rest.starts_with("<w:p>") || rest.starts_with("<w:p ") {
            return Some(start);
        }
        if start == 0 {
            return None;
        }
        cursor = start;
    }
}

fn escape_xml_text(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

struct ImportedParagraph {
    id: String,
    text_id: String,
    text: String,
}

fn document_xml(parts: &BTreeMap<String, Vec<u8>>) -> UseResult<String> {
    let bytes = parts.get(DOCUMENT_PART).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A document snapshot requires word/document.xml.",
        )
    })?;
    String::from_utf8(bytes.clone()).map_err(|_| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            "word/document.xml is not UTF-8.",
        )
    })
}

fn read_document_paragraphs(xml: &str) -> UseResult<Vec<ImportedParagraph>> {
    let mut paragraphs = Vec::new();
    let mut search_from = 0;
    while let Some(relative) = xml[search_from..].find("paraId=\"") {
        let marker_at = search_from + relative;
        let id_start = marker_at + "paraId=\"".len();
        let id_end = xml[id_start..].find('"').ok_or_else(|| {
            collaboration_error(
                "office.collaboration.snapshot_invalid",
                "A document paragraph id is missing its closing quote.",
            )
        })? + id_start;
        let id = canonical_word_id(&xml[id_start..id_end])?;
        let start = paragraph_start(xml, marker_at).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Paragraph '{id}' is not inside a w:p element."),
            )
        })?;
        let end_rel = xml[start..].find("</w:p>").ok_or_else(|| {
            collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Paragraph '{id}' is missing its end tag."),
            )
        })?;
        let end = start + end_rel + "</w:p>".len();
        let paragraph = &xml[start..end];
        let open_tag_end = paragraph.find('>').ok_or_else(|| text_run_error(&id))?;
        let text_id = match attribute_value(&paragraph[..open_tag_end], "textId=") {
            Some(value) => canonical_word_id(&value)?,
            None => next_word_id(&id)?,
        };
        if text_id == id {
            return Err(collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Paragraph '{id}' cannot use the same identity for its text."),
            ));
        }
        paragraphs.push(ImportedParagraph {
            id: id.clone(),
            text_id,
            text: unescape_xml_text(&single_text_run(paragraph, &id)?),
        });
        search_from = end;
    }
    if paragraphs.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A document snapshot import requires at least one addressed paragraph.",
        ));
    }
    Ok(paragraphs)
}

fn canonical_word_id(value: &str) -> UseResult<String> {
    let invalid = || {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!(
                "Word identity '{value}' must be eight hex digits from 00000001 through 7FFFFFFF."
            ),
        )
    };
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid());
    }
    let number = u32::from_str_radix(value, 16).map_err(|_| invalid())?;
    if number == 0 || number > 0x7fff_ffff {
        return Err(invalid());
    }
    Ok(format!("{number:08X}"))
}

fn next_word_id(paragraph_id: &str) -> UseResult<String> {
    let number = u32::from_str_radix(paragraph_id, 16).unwrap_or(0);
    let next = number.saturating_add(1);
    if next == 0 || next > 0x7fff_ffff {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Paragraph '{paragraph_id}' has no free text identity."),
        ));
    }
    Ok(format!("{next:08X}"))
}

fn single_text_run(paragraph: &str, paragraph_id: &str) -> UseResult<String> {
    let open = paragraph
        .find("<w:t")
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let open_end = paragraph[open..]
        .find('>')
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let close = paragraph[open + open_end..]
        .find("</w:t>")
        .ok_or_else(|| text_run_error(paragraph_id))?;
    let text_start = open + open_end + 1;
    let text_end = open + open_end + close;
    if paragraph[text_end + "</w:t>".len()..].contains("<w:t") {
        return Err(text_run_error(paragraph_id));
    }
    Ok(paragraph[text_start..text_end].to_owned())
}

fn attribute_value(tag: &str, name: &str) -> Option<String> {
    let start = tag.find(name)? + name.len();
    let rest = tag.get(start..)?;
    let quoted = rest.strip_prefix('"')?;
    let end = quoted.find('"')?;
    Some(quoted[..end].to_owned())
}

fn unescape_xml_text(value: &str) -> String {
    value
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
}

fn bootstrap_document_update(
    artifact_id: &str,
    paragraphs: &[ImportedParagraph],
) -> UseResult<Vec<u8>> {
    let mut options = Options::with_client_id(ClientID::new(424_242));
    options.offset_kind = OffsetKind::Utf16;
    let document = Doc::with_options(options);
    let metadata =
        document.get_or_insert_map(format!("{NATIVE_OFFICE_COLLABORATION_NAMESPACE}.metadata"));
    let initializers = document.get_or_insert_array(format!(
        "{NATIVE_OFFICE_COLLABORATION_NAMESPACE}.bootstrap.initializers"
    ));
    let fragment = document.get_or_insert_xml_fragment(format!(
        "{NATIVE_OFFICE_COLLABORATION_NAMESPACE}.document.content"
    ));
    let mut transaction = document.transact_mut();
    metadata.insert(
        &mut transaction,
        "protocol",
        NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    );
    metadata.insert(
        &mut transaction,
        "version",
        i64::from(NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION),
    );
    metadata.insert(&mut transaction, "artifactId", artifact_id);
    metadata.insert(&mut transaction, "kind", "document");
    metadata.insert(&mut transaction, "initialized", true);
    initializers.push_back(&mut transaction, "424242:snapshot-import");
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    section.insert_attribute(&mut transaction, "id", "document-section-1");
    for paragraph in paragraphs {
        let element = section.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
        element.insert_attribute(&mut transaction, "paragraphId", paragraph.id.as_str());
        element.insert_attribute(&mut transaction, "textId", paragraph.text_id.as_str());
        element.push_back(
            &mut transaction,
            XmlTextPrelim::new(paragraph.text.as_str()),
        );
    }
    drop(transaction);
    let update = document
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    Ok(update)
}

fn remember_binding(package_path: &Path, store_path: &Path) -> UseResult<()> {
    fs::write(
        store_path.join(SOURCE_PACKAGE_FILE),
        package_path.to_string_lossy().as_bytes(),
    )
    .map_err(|error| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Failed to record the snapshot package path: {error}"),
        )
    })?;
    fs::write(
        binding_path(package_path),
        store_path.to_string_lossy().as_bytes(),
    )
    .map_err(|error| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Failed to bind the package to its replica: {error}"),
        )
    })?;
    Ok(())
}

fn bound_store_path(package_path: &Path) -> UseResult<Option<PathBuf>> {
    let path = binding_path(package_path);
    if !path.is_file() {
        return Ok(None);
    }
    let stored = fs::read_to_string(&path).map_err(|error| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Failed to read the package replica binding: {error}"),
        )
    })?;
    let stored = stored.trim();
    if stored.is_empty() {
        return Ok(None);
    }
    Ok(Some(PathBuf::from(stored)))
}

fn binding_path(package_path: &Path) -> PathBuf {
    let mut name = package_path.file_name().unwrap_or_default().to_os_string();
    name.push(".a3s-replica");
    package_path.with_file_name(name)
}

fn paths_match(stored: &str, requested: &Path) -> bool {
    if stored == requested.to_string_lossy() {
        return true;
    }
    match (fs::canonicalize(stored), fs::canonicalize(requested)) {
        (Ok(left), Ok(right)) => left == right,
        _ => false,
    }
}

fn replace_inline_cell(xml: &str, row: u32, column: u32, text: &str) -> UseResult<String> {
    let reference = a1_reference(column, row);
    let marker = format!("r=\"{reference}\"");
    let marker_at = unique_marker(xml, &marker).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("The worksheet has no single cell '{reference}'."),
        )
    })?;
    let start = xml[..marker_at]
        .rfind("<c ")
        .or_else(|| xml[..marker_at].rfind("<c>"))
        .ok_or_else(|| {
            collaboration_error(
                "office.collaboration.snapshot_invalid",
                format!("Cell '{reference}' is not a worksheet cell."),
            )
        })?;
    let end_rel = xml[start..].find("</c>").ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Cell '{reference}' is missing its end tag."),
        )
    })?;
    let end = start + end_rel + "</c>".len();
    let cell = &xml[start..end];
    if cell.contains("<f") || cell.contains(" t=\"s\"") {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Cell '{reference}' is a formula or shared string and is not rewritten."),
        ));
    }
    let (open_tag, close_tag) = if cell.contains("t=\"inlineStr\"") {
        ("<t", "</t>")
    } else if cell.contains("t=\"str\"") {
        ("<v", "</v>")
    } else {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Cell '{reference}' has no single plain-text value."),
        ));
    };
    replace_single_run(xml, start, end, open_tag, close_tag, text, &reference)
}

fn replace_shape_text(xml: &str, shape_id: &str, text: &str) -> UseResult<String> {
    if shape_id.is_empty() || shape_id.contains('"') {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A presentation snapshot shape id is invalid.",
        ));
    }
    let marker = format!("id=\"{shape_id}\"");
    let marker_at = unique_marker(xml, &marker).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("The slide has no single shape '{shape_id}'."),
        )
    })?;
    let start = xml[..marker_at].rfind("<p:sp").ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Shape '{shape_id}' is not inside a p:sp element."),
        )
    })?;
    let end_rel = xml[start..].find("</p:sp>").ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("Shape '{shape_id}' is missing its end tag."),
        )
    })?;
    let end = start + end_rel + "</p:sp>".len();
    replace_single_run(xml, start, end, "<a:t", "</a:t>", text, shape_id)
}

fn replace_single_run(
    xml: &str,
    start: usize,
    end: usize,
    open_tag: &str,
    close_tag: &str,
    text: &str,
    label: &str,
) -> UseResult<String> {
    let element = &xml[start..end];
    let open = element
        .find(open_tag)
        .ok_or_else(|| text_run_error(label))?;
    let open_end = element[open..]
        .find('>')
        .ok_or_else(|| text_run_error(label))?
        + open;
    let close = element[open_end..]
        .find(close_tag)
        .ok_or_else(|| text_run_error(label))?
        + open_end;
    let text_start = open_end + 1;
    let text_end = close;
    if element[text_end + close_tag.len()..].contains(open_tag) {
        return Err(text_run_error(label));
    }
    let mut next = String::with_capacity(xml.len() + text.len());
    next.push_str(&xml[..start]);
    next.push_str(&element[..text_start]);
    next.push_str(&escape_xml_text(text));
    next.push_str(&element[text_end..]);
    next.push_str(&xml[end..]);
    Ok(next)
}

fn replace_pdf_literal(
    pdf: &str,
    anchor_key: &str,
    identifier: &str,
    value_key: &str,
    text: &str,
) -> UseResult<String> {
    if identifier.is_empty() || identifier.contains(['(', ')', '\\']) {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            "A PDF snapshot identifier cannot contain literal-string escapes.",
        ));
    }
    let anchor = format!("{anchor_key} ({identifier})");
    let anchor_at = unique_marker(pdf, &anchor).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("The PDF has no single {anchor_key} ({identifier})."),
        )
    })?;
    let region_end = pdf[anchor_at..]
        .find("endobj")
        .map(|offset| anchor_at + offset)
        .unwrap_or(pdf.len());
    let region = &pdf[anchor_at..region_end];
    let value_at = region.find(value_key).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("PDF {anchor_key} ({identifier}) has no {value_key} literal."),
        )
    })?;
    let after_key = &region[value_at + value_key.len()..];
    let literal_rel = after_key.find('(').ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("PDF {value_key} for '{identifier}' is not a literal string."),
        )
    })?;
    let between = &after_key[..literal_rel];
    if between.chars().any(|character| !character.is_whitespace()) {
        return Err(collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("PDF {value_key} for '{identifier}' is not a literal string."),
        ));
    }
    let literal_start = anchor_at + value_at + value_key.len() + literal_rel + 1;
    let mut escaped = false;
    let mut literal_end = None;
    for (offset, character) in pdf[literal_start..].char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        match character {
            '\\' => escaped = true,
            ')' => {
                literal_end = Some(literal_start + offset);
                break;
            }
            _ => {}
        }
    }
    let literal_end = literal_end.ok_or_else(|| {
        collaboration_error(
            "office.collaboration.snapshot_invalid",
            format!("PDF {value_key} for '{identifier}' has no closing parenthesis."),
        )
    })?;
    let mut next = String::with_capacity(pdf.len() + text.len());
    next.push_str(&pdf[..literal_start]);
    next.push_str(&escape_pdf_literal(text));
    next.push_str(&pdf[literal_end..]);
    Ok(next)
}

fn unique_marker(xml: &str, marker: &str) -> Option<usize> {
    let first = xml.find(marker)?;
    if xml[first + marker.len()..].contains(marker) {
        return None;
    }
    Some(first)
}

fn a1_reference(column: u32, row: u32) -> String {
    let mut name = Vec::new();
    let mut index = column;
    loop {
        name.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    name.reverse();
    format!("{}{}", String::from_utf8(name).unwrap_or_default(), row + 1)
}

fn escape_pdf_literal(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '\\' | '(' | ')' => {
                escaped.push('\\');
                escaped.push(character);
            }
            _ => escaped.push(character),
        }
    }
    escaped
}

fn text_run_error(paragraph_id: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.snapshot_invalid",
        format!(
            "Paragraph '{paragraph_id}' must contain one text run before a snapshot can write it."
        ),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn snapshot_writes_spliced_text_and_leaves_other_parts_untouched() {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
        let mut parts = BTreeMap::new();
        parts.insert("word/document.xml".to_owned(), document.to_vec());
        parts.insert("word/styles.xml".to_owned(), b"<styles/>".to_vec());
        let written =
            write_document_snapshot(parts.clone(), &[("00000001", "Hello 🦀 <ok>")]).unwrap();
        let xml = String::from_utf8(written["word/document.xml"].clone()).unwrap();
        assert!(xml.contains("Hello 🦀 &lt;ok&gt;"));
        assert!(xml.contains("w14:paraId=\"00000001\""));
        assert_eq!(written["word/styles.xml"], parts["word/styles.xml"]);
        assert!(write_document_snapshot(parts, &[("00000001", "A"), ("00000099", "B")]).is_err());
    }

    #[test]
    fn spreadsheet_presentation_and_pdf_snapshots_write_one_target() {
        let worksheet = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Hello</t></is></c><c r=\"B1\"><v>42</v></c><c r=\"C1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>";
        let mut sheet_parts = BTreeMap::new();
        sheet_parts.insert(
            "xl/worksheets/sheet1.xml".to_owned(),
            worksheet.as_bytes().to_vec(),
        );
        sheet_parts.insert("xl/styles.xml".to_owned(), b"<styles/>".to_vec());
        let written = write_spreadsheet_snapshot(
            sheet_parts.clone(),
            &[("xl/worksheets/sheet1.xml", 0, 0, "Hello 中")],
        )
        .unwrap();
        let xml = String::from_utf8(written["xl/worksheets/sheet1.xml"].clone()).unwrap();
        assert!(xml.contains("<t>Hello 中</t>"));
        assert!(xml.contains("<v>42</v>"));
        assert!(xml.contains("t=\"s\""));
        assert_eq!(written["xl/styles.xml"], sheet_parts["xl/styles.xml"]);
        assert!(write_spreadsheet_snapshot(
            sheet_parts,
            &[("xl/worksheets/sheet1.xml", 0, 2, "nope")],
        )
        .is_err());

        let slide = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"2\" name=\"Title\"/><p:txBody><a:p><a:r><a:t>Shared</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
        let mut slide_parts = BTreeMap::new();
        slide_parts.insert(
            "ppt/slides/slide1.xml".to_owned(),
            slide.as_bytes().to_vec(),
        );
        slide_parts.insert("ppt/presentation.xml".to_owned(), b"<kept/>".to_vec());
        let written = write_presentation_snapshot(
            slide_parts.clone(),
            &[("ppt/slides/slide1.xml", "2", "Shared中文")],
        )
        .unwrap();
        let xml = String::from_utf8(written["ppt/slides/slide1.xml"].clone()).unwrap();
        assert!(xml.contains("<a:t>Shared中文</a:t>"));
        assert!(xml.contains("name=\"Title\""));
        assert_eq!(
            written["ppt/presentation.xml"],
            slide_parts["ppt/presentation.xml"]
        );

        let pdf = b"1 0 obj << /T (Applicant.Name) /V (Ada) >> endobj 2 0 obj << /T (Applicant.Email) /V (ada@example.com) >> endobj 3 0 obj << /NM (annotation-native-1) /Contents (Initial note) >> endobj trailer".to_vec();
        let saved = write_pdf_snapshot(
            &pdf,
            &[("Applicant.Name", "Grace (ok)")],
            &[("annotation-native-1", "Updated")],
        )
        .unwrap();
        let text = String::from_utf8(saved).unwrap();
        assert!(text.contains("/V (Grace \\(ok\\))"));
        assert!(text.contains("/V (ada@example.com)"));
        assert!(text.contains("/Contents (Updated)"));
        assert!(text.ends_with("trailer"));
        assert!(write_pdf_snapshot(&pdf, &[("Missing", "x")], &[]).is_err());
    }

    #[test]
    fn import_once_then_a_remote_splice_round_trips_without_a_second_session() {
        let temp = tempfile::tempdir().unwrap();
        let package = temp.path().join("letter.docx");
        std::fs::write(&package, b"snapshot").unwrap();
        let document = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body><w:p w14:paraId=\"00000001\" w14:textId=\"00000002\"><w:r><w:t>Hello 😀 world</w:t></w:r></w:p></w:body></w:document>";
        let mut parts = BTreeMap::new();
        parts.insert("word/document.xml".to_owned(), document.as_bytes().to_vec());
        parts.insert("customXml/item1.xml".to_owned(), b"<kept/>".to_vec());

        let imported = import_document_snapshot(
            &package,
            &parts,
            &temp.path().join("replica"),
            "letter-artifact",
        )
        .unwrap();
        assert!(imported.created);
        assert_eq!(paragraph_text(&imported.store), "Hello 😀 world");

        let again = import_document_snapshot(
            &package,
            &parts,
            &temp.path().join("replica-other"),
            "letter-artifact",
        )
        .unwrap();
        assert!(!again.created);

        let second = crate::NativeOfficeCollaborationStore::create(
            crate::NativeOfficeCollaborationCreateRequest {
                store: temp.path().join("replica-peer"),
                artifact_id: "letter-artifact".to_owned(),
                kind: crate::NativeOfficeCollaborationArtifactKind::Document,
                actor_id: "agent-peer".to_owned(),
                actor_kind: crate::NativeOfficeCollaborationActorKind::Agent,
                mode: crate::NativeOfficeCollaborationMode::Edit,
                operation_id: "join-letter".to_owned(),
                namespace: None,
                client_id: Some(900_071),
                initial_update: Some(imported.store.synchronize(None).unwrap().update),
            },
        )
        .unwrap();

        imported
            .store
            .mutate(crate::NativeOfficeCollaborationMutationRequest {
                operation_id: "document-splice-1".to_owned(),
                actor_id: "snapshot-import".to_owned(),
                mode: crate::NativeOfficeCollaborationMode::Edit,
                expected_artifact_id: "letter-artifact".to_owned(),
                expected_kind: crate::NativeOfficeCollaborationArtifactKind::Document,
                mutation: crate::NativeOfficeCollaborationMutation::DocumentSplice {
                    paragraph_id: "00000001".to_owned(),
                    text_id: "00000002".to_owned(),
                    index_utf16: 6,
                    delete_utf16: 2,
                    expected_slice: "😀".to_owned(),
                    insert: "🦀".to_owned(),
                },
                if_state_vector: None,
            })
            .unwrap();
        second
            .apply(crate::NativeOfficeCollaborationApplyRequest {
                operation_id: "apply-splice".to_owned(),
                actor_id: "agent-peer".to_owned(),
                mode: crate::NativeOfficeCollaborationMode::Edit,
                expected_artifact_id: "letter-artifact".to_owned(),
                expected_kind: crate::NativeOfficeCollaborationArtifactKind::Document,
                update: imported
                    .store
                    .synchronize(Some(&second.synchronize(None).unwrap().state_vector))
                    .unwrap()
                    .update,
                if_state_vector: None,
                origin: None,
            })
            .unwrap();
        assert_eq!(paragraph_text(&imported.store), "Hello 🦀 world");
        assert_eq!(paragraph_text(&second), "Hello 🦀 world");

        let saved = write_document_snapshot(
            parts.clone(),
            &[("00000001", &paragraph_text(&imported.store))],
        )
        .unwrap();
        let xml = String::from_utf8(saved["word/document.xml"].clone()).unwrap();
        assert!(xml.contains("Hello 🦀 world"));
        assert!(!xml.contains('😀'));
        assert_eq!(saved["customXml/item1.xml"], parts["customXml/item1.xml"]);

        let reopened = crate::NativeOfficeCollaborationStore::open(imported.store.path()).unwrap();
        assert_eq!(paragraph_text(&reopened), "Hello 🦀 world");
        let repeated = import_document_snapshot(
            &package,
            &parts,
            &temp.path().join("replica"),
            "letter-artifact",
        )
        .unwrap();
        assert!(!repeated.created);
        assert_eq!(paragraph_text(&repeated.store), "Hello 🦀 world");
        assert_eq!(
            live_replica_blocks_office_session(&package)
                .unwrap_err()
                .code,
            "office.collaboration.replica_live"
        );

        drop(imported);
        drop(again);
        drop(repeated);
        drop(reopened);
        std::fs::remove_dir_all(temp.path().join("replica")).unwrap();
        assert!(live_replica_blocks_office_session(&package).is_ok());
    }

    fn paragraph_text(store: &crate::NativeOfficeCollaborationStore) -> String {
        match store.project().unwrap().content {
            crate::NativeOfficeCollaborationProjectedContent::Document { paragraphs, .. } => {
                paragraphs
                    .into_iter()
                    .next()
                    .expect("imported paragraph")
                    .text
            }
            _ => panic!("imported replica must project a document"),
        }
    }
}
