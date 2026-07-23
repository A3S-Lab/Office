mod presentation;
mod spreadsheet;
mod word;

use a3s_use_core::UseResult;

use super::output::BoundedOutput;
use crate::{DocumentKind, DocumentNode, NativeOfficeDocument};

const STYLES: &str = r#"
:root{color-scheme:light;font-family:ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#eef1f5;color:#172033}
*{box-sizing:border-box}body{margin:0;padding:2rem;line-height:1.45}.preview-header,main{max-width:1200px;margin:0 auto}.preview-header{margin-bottom:1.5rem}.preview-header h1{margin:0 0 .25rem;font-size:1.5rem}.preview-header p{margin:0;color:#526079}
.semantic-note{font-size:.875rem}.word-region,.sheet,.slide-card{background:#fff;border:1px solid #d8dee9;border-radius:.75rem;box-shadow:0 .25rem 1rem rgba(24,36,56,.08);margin:0 0 1.5rem;padding:1.25rem}.word-region>h2,.sheet>h2,.slide-card>h2{font-size:1rem;margin:0 0 1rem;color:#43506a}
.word-region p{min-height:1em;margin:.5rem 0;white-space:pre-wrap}.run.is-bold{font-weight:700}.run.is-italic{font-style:italic}.run.is-strike{text-decoration:line-through}.run.is-hidden{opacity:.55}.hyperlink{color:#315ca8;text-decoration:underline dotted}.picture-inline{display:inline-block;vertical-align:middle}.picture-inline img,.sheet-picture img{max-width:100%;height:auto}.image-placeholder{display:inline-flex;align-items:center;justify-content:center;min-width:6rem;min-height:3rem;padding:.5rem;border:1px dashed #9aa6b8;background:#f8fafc;color:#667085;font-size:.75rem}
table{border-collapse:collapse;width:100%;margin:.75rem 0}th,td{border:1px solid #ccd4df;padding:.4rem .55rem;text-align:left;vertical-align:top}th{background:#f3f6fa;font-weight:600}.sheet-row-label,.cell-reference{white-space:nowrap;color:#526079}.cell-value{white-space:pre-wrap}.cell-formula{display:block;color:#6b4ca5;font-size:.8rem;overflow-wrap:anywhere}.sheet-picture{margin:1rem 0}.sheet-picture figcaption{color:#526079;font-size:.8rem}
.slide-card{padding:1rem}.slide-canvas{position:relative;width:100%;overflow:hidden;background:#fff;border:1px solid #b9c2d0;box-shadow:inset 0 0 0 1px #fff}.slide-object{position:absolute;overflow:hidden;padding:.25rem}.slide-object.unpositioned{position:relative;display:block;width:auto;height:auto;margin:.5rem}.slide-object.shape,.slide-object.placeholder{border:1px solid #8491a5;background:#f8fafc}.slide-object.placeholder{border-style:dashed}.slide-object.picture{padding:0}.slide-object.picture img{display:block;width:100%;height:100%;object-fit:contain}.slide-object.chart,.slide-object.connector{display:flex;align-items:center;justify-content:center;border:1px dashed #8491a5;color:#526079;background:#f8fafc}.slide-object.group{border:1px dotted #9aa6b8}.slide-object p{margin:.15rem;white-space:pre-wrap}.slide-object table{font-size:.75rem;margin:0}.slide-notes{margin:.75rem 0 0;padding:.75rem;background:#fffbe8;border-left:.25rem solid #e6b94f;white-space:pre-wrap}.slide-empty{display:flex;align-items:center;justify-content:center;color:#7a8699;height:100%}
@media(max-width:700px){body{padding:1rem}.word-region,.sheet,.slide-card{padding:.75rem}}
"#;

pub(super) fn render(document: &NativeOfficeDocument, limit: usize) -> UseResult<String> {
    let mut output = BoundedOutput::new(limit);
    output.push("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>A3S Native Office semantic preview</title><style>")?;
    output.push(STYLES)?;
    output.push(
        "</style></head><body data-renderer=\"a3s-office-semantic-v1\" data-document-kind=\"",
    )?;
    output.push(document_kind(document.kind()))?;
    output.push("\"><header class=\"preview-header\"><h1>")?;
    output.push(document_title(document.kind()))?;
    output.push("</h1><p class=\"semantic-note\">Deterministic semantic preview; it does not claim Microsoft Office layout fidelity.</p></header><main>")?;
    match document.kind() {
        DocumentKind::Word => word::render(document, &mut output)?,
        DocumentKind::Spreadsheet => spreadsheet::render(document, &mut output)?,
        DocumentKind::Presentation => presentation::render(document, &mut output)?,
    }
    output.push("</main></body></html>")?;
    Ok(output.into_string())
}

pub(super) fn write_node_attributes(
    output: &mut BoundedOutput,
    node: &DocumentNode,
) -> UseResult<()> {
    output.push(" data-path=\"")?;
    output.attribute(&node.path)?;
    output.push("\" data-node-type=\"")?;
    output.push(node.node_type.label())?;
    output.push("\"")?;
    if let Some(style) = &node.style {
        output.push(" data-style=\"")?;
        output.attribute(style)?;
        output.push("\"")?;
    }
    Ok(())
}

pub(super) fn write_optional_attribute(
    output: &mut BoundedOutput,
    name: &str,
    value: Option<&str>,
) -> UseResult<()> {
    let Some(value) = value else {
        return Ok(());
    };
    output.push(" ")?;
    output.push(name)?;
    output.push("=\"")?;
    output.attribute(value)?;
    output.push("\"")
}

pub(super) fn flag(node: &DocumentNode, key: &str) -> bool {
    node.format.get(key).is_some_and(|value| {
        matches!(
            value.to_ascii_lowercase().as_str(),
            "1" | "true" | "on" | "yes"
        )
    })
}

pub(super) fn safe_color(value: Option<&str>) -> Option<String> {
    let value = value?.trim().trim_start_matches('#');
    if matches!(value.len(), 6 | 8) && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Some(format!("#{}", value.to_ascii_uppercase()))
    } else {
        None
    }
}

pub(super) fn point_size(value: Option<&str>) -> Option<f64> {
    let value = value?.trim().strip_suffix("pt")?.parse::<f64>().ok()?;
    value
        .is_finite()
        .then_some(value)
        .filter(|value| (1.0..=400.0).contains(value))
}

fn document_kind(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "word",
        DocumentKind::Spreadsheet => "spreadsheet",
        DocumentKind::Presentation => "presentation",
    }
}

fn document_title(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "Word semantic preview",
        DocumentKind::Spreadsheet => "Spreadsheet semantic preview",
        DocumentKind::Presentation => "Presentation semantic preview",
    }
}
