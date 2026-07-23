use serde::{Deserialize, Serialize};

/// Workbook-global or worksheet-local scope for one Spreadsheet defined name.
///
/// JSON uses the string `"workbook"` or a worksheet name while Rust callers
/// retain a closed typed distinction. The explicit `worksheet:` prefix escapes
/// a worksheet literally named `workbook`.
#[derive(Debug, Default, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(try_from = "String", into = "String")]
pub enum NativeSpreadsheetNamedRangeScope {
    #[default]
    Workbook,
    Worksheet(String),
}

impl NativeSpreadsheetNamedRangeScope {
    pub fn worksheet(name: impl Into<String>) -> Self {
        Self::Worksheet(name.into())
    }

    pub fn as_worksheet(&self) -> Option<&str> {
        match self {
            Self::Workbook => None,
            Self::Worksheet(name) => Some(name),
        }
    }

    pub fn label(&self) -> String {
        self.clone().into()
    }
}

impl From<NativeSpreadsheetNamedRangeScope> for String {
    fn from(value: NativeSpreadsheetNamedRangeScope) -> Self {
        match value {
            NativeSpreadsheetNamedRangeScope::Workbook => "workbook".to_string(),
            NativeSpreadsheetNamedRangeScope::Worksheet(name)
                if name.eq_ignore_ascii_case("workbook") =>
            {
                format!("worksheet:{name}")
            }
            NativeSpreadsheetNamedRangeScope::Worksheet(name) => name,
        }
    }
}

impl TryFrom<String> for NativeSpreadsheetNamedRangeScope {
    type Error = String;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        const WORKSHEET_PREFIX: &str = "worksheet:";
        if value.eq_ignore_ascii_case("workbook") {
            Ok(Self::Workbook)
        } else if value
            .get(..WORKSHEET_PREFIX.len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(WORKSHEET_PREFIX))
        {
            let name = value[WORKSHEET_PREFIX.len()..].to_string();
            if name.is_empty() {
                Err("Spreadsheet named-range worksheet scope cannot be empty.".to_string())
            } else {
                Ok(Self::Worksheet(name))
            }
        } else if value.is_empty() {
            Err("Spreadsheet named-range scope cannot be empty.".to_string())
        } else {
            Ok(Self::Worksheet(value))
        }
    }
}

/// One complete typed Spreadsheet defined name (commonly called a named
/// range). `reference` may be a range, constant, or formula body, but never
/// includes the formula-bar leading `=`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetNamedRange {
    pub name: String,
    #[serde(rename = "ref")]
    pub reference: String,
    #[serde(default, skip_serializing_if = "is_workbook_scope")]
    pub scope: NativeSpreadsheetNamedRangeScope,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub volatile: bool,
}

impl NativeSpreadsheetNamedRange {
    pub fn new(name: impl Into<String>, reference: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            reference: reference.into(),
            scope: NativeSpreadsheetNamedRangeScope::Workbook,
            comment: None,
            volatile: false,
        }
    }

    pub fn with_scope(mut self, scope: NativeSpreadsheetNamedRangeScope) -> Self {
        self.scope = scope;
        self
    }

    pub fn with_comment(mut self, comment: impl Into<String>) -> Self {
        self.comment = Some(comment.into());
        self
    }

    pub fn with_volatile(mut self, volatile: bool) -> Self {
        self.volatile = volatile;
        self
    }
}

fn is_workbook_scope(scope: &NativeSpreadsheetNamedRangeScope) -> bool {
    matches!(scope, NativeSpreadsheetNamedRangeScope::Workbook)
}

const fn is_false(value: &bool) -> bool {
    !*value
}
