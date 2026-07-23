use std::fmt::{self, Write as _};

use a3s_use_core::UseResult;

use super::render_error;

pub(super) struct BoundedOutput {
    content: String,
    limit: usize,
}

impl BoundedOutput {
    pub(super) fn new(limit: usize) -> Self {
        Self {
            content: String::with_capacity(limit.min(16 * 1024)),
            limit,
        }
    }

    pub(super) fn push(&mut self, value: &str) -> UseResult<()> {
        self.ensure_additional(value.len())?;
        self.content.push_str(value);
        Ok(())
    }

    pub(super) fn push_fmt(&mut self, arguments: fmt::Arguments<'_>) -> UseResult<()> {
        self.write_fmt(arguments)
            .map_err(|_| output_too_large(self.limit))
    }

    pub(super) fn text(&mut self, value: &str) -> UseResult<()> {
        self.escaped(value, false)
    }

    pub(super) fn attribute(&mut self, value: &str) -> UseResult<()> {
        self.escaped(value, true)
    }

    pub(super) fn ensure_additional(&self, bytes: usize) -> UseResult<()> {
        if self
            .content
            .len()
            .checked_add(bytes)
            .is_some_and(|length| length <= self.limit)
        {
            return Ok(());
        }
        Err(output_too_large(self.limit))
    }

    pub(super) fn into_string(self) -> String {
        self.content
    }

    fn escaped(&mut self, value: &str, attribute: bool) -> UseResult<()> {
        let mut plain_start = 0;
        for (offset, character) in value.char_indices() {
            let replacement = match character {
                '&' => Some("&amp;"),
                '<' => Some("&lt;"),
                '>' => Some("&gt;"),
                '"' if attribute => Some("&quot;"),
                '\'' if attribute => Some("&#39;"),
                '\n' if attribute => Some("&#10;"),
                '\r' if attribute => Some("&#13;"),
                '\t' if attribute => Some("&#9;"),
                value if !is_xml_character(value) => Some("\u{fffd}"),
                _ => None,
            };
            let Some(replacement) = replacement else {
                continue;
            };
            self.push(&value[plain_start..offset])?;
            self.push(replacement)?;
            plain_start = offset + character.len_utf8();
        }
        self.push(&value[plain_start..])
    }
}

impl fmt::Write for BoundedOutput {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.ensure_additional(value.len())
            .map_err(|_| fmt::Error)?;
        self.content.push_str(value);
        Ok(())
    }
}

fn is_xml_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{a}' | '\u{d}') || value >= '\u{20}'
}

fn output_too_large(limit: usize) -> a3s_use_core::UseError {
    render_error(
        "use.office.render_output_too_large",
        format!("Native Office semantic render exceeds the {limit}-byte output limit."),
    )
    .with_suggestion("Render a smaller document or remove large embedded images.")
    .with_detail("limitBytes", limit)
}
