/// Maximum number of Unicode scalar values accepted in one cell formula.
pub const MAX_SPREADSHEET_FORMULA_CHARACTERS: usize = 8_192;

/// Maximum nesting depth accepted by the native formula parser.
pub const MAX_SPREADSHEET_FORMULA_DEPTH: usize = 128;

/// Maximum number of AST nodes accepted by the native formula parser.
pub const MAX_SPREADSHEET_FORMULA_NODES: usize = 8_192;

/// Maximum disjoint reference areas retained while graphing or calculating
/// one Spreadsheet formula value.
pub const MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS: usize = 100_000;

/// A parsed Spreadsheet formula body.
///
/// The source may include one formula-bar leading `=`. The parser removes that
/// marker before assigning spans, so all spans address the normalized formula
/// body that is stored in SpreadsheetML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SpreadsheetFormula {
    pub root: SpreadsheetFormulaExpression,
}

/// A half-open UTF-8 byte range in the normalized formula body.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct SpreadsheetFormulaSpan {
    pub start: usize,
    pub end: usize,
}

impl SpreadsheetFormulaSpan {
    pub(crate) const fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }

    pub(crate) const fn through(self, other: Self) -> Self {
        Self {
            start: self.start,
            end: other.end,
        }
    }
}

/// One source-spanned formula expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SpreadsheetFormulaExpression {
    pub span: SpreadsheetFormulaSpan,
    pub kind: SpreadsheetFormulaExpressionKind,
}

/// Closed expression variants produced by the native Spreadsheet parser.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SpreadsheetFormulaExpressionKind {
    Literal(SpreadsheetFormulaLiteral),
    Reference(SpreadsheetFormulaReference),
    Name {
        qualifier: Option<SpreadsheetFormulaQualifier>,
        name: String,
    },
    StructuredReference {
        qualifier: Option<SpreadsheetFormulaQualifier>,
        reference: String,
    },
    Unary {
        operator: SpreadsheetFormulaUnaryOperator,
        operand: Box<SpreadsheetFormulaExpression>,
    },
    Postfix {
        operator: SpreadsheetFormulaPostfixOperator,
        operand: Box<SpreadsheetFormulaExpression>,
    },
    Binary {
        operator: SpreadsheetFormulaBinaryOperator,
        left: Box<SpreadsheetFormulaExpression>,
        right: Box<SpreadsheetFormulaExpression>,
    },
    FunctionCall {
        qualifier: Option<SpreadsheetFormulaQualifier>,
        name: String,
        arguments: Vec<Option<SpreadsheetFormulaExpression>>,
    },
    Parenthesized(Box<SpreadsheetFormulaExpression>),
    Array {
        rows: Vec<Vec<SpreadsheetFormulaExpression>>,
    },
}

/// Scalar literal values represented directly in formula source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SpreadsheetFormulaLiteral {
    Number(String),
    Text(String),
    Boolean(bool),
    Error(SpreadsheetFormulaErrorLiteral),
}

/// Error literals supported by current Spreadsheet applications.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SpreadsheetFormulaErrorLiteral {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData,
    Spill,
    Calculation,
    Field,
    Blocked,
    Unknown,
    Busy,
    Connect,
    Python,
}

impl SpreadsheetFormulaErrorLiteral {
    pub fn parse(value: &str) -> Option<Self> {
        [
            Self::GettingData,
            Self::DivisionByZero,
            Self::NotAvailable,
            Self::Calculation,
            Self::Reference,
            Self::Blocked,
            Self::Unknown,
            Self::Connect,
            Self::Python,
            Self::Value,
            Self::Field,
            Self::Spill,
            Self::Number,
            Self::Name,
            Self::Busy,
            Self::Null,
        ]
        .into_iter()
        .find(|literal| value.eq_ignore_ascii_case(literal.as_str()))
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Self::Null => "#NULL!",
            Self::DivisionByZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Reference => "#REF!",
            Self::Name => "#NAME?",
            Self::Number => "#NUM!",
            Self::NotAvailable => "#N/A",
            Self::GettingData => "#GETTING_DATA",
            Self::Spill => "#SPILL!",
            Self::Calculation => "#CALC!",
            Self::Field => "#FIELD!",
            Self::Blocked => "#BLOCKED!",
            Self::Unknown => "#UNKNOWN!",
            Self::Busy => "#BUSY!",
            Self::Connect => "#CONNECT!",
            Self::Python => "#PYTHON!",
        }
    }
}

/// Workbook and worksheet qualification attached to a reference or name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SpreadsheetFormulaQualifier {
    /// External workbook prefix, including its square brackets when present.
    pub workbook: Option<String>,
    /// First worksheet in the qualifier.
    pub worksheet: String,
    /// Last worksheet for a three-dimensional worksheet qualifier.
    pub worksheet_end: Option<String>,
}

impl SpreadsheetFormulaQualifier {
    pub fn is_external(&self) -> bool {
        self.workbook.is_some()
    }

    pub fn is_three_dimensional(&self) -> bool {
        self.worksheet_end.is_some()
    }
}

/// One absolute/relative A1 reference endpoint.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SpreadsheetFormulaReference {
    pub qualifier: Option<SpreadsheetFormulaQualifier>,
    pub kind: SpreadsheetFormulaReferenceKind,
}

/// A cell, whole-column, or whole-row A1 reference endpoint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpreadsheetFormulaReferenceKind {
    Cell {
        column: u32,
        row: u32,
        absolute_column: bool,
        absolute_row: bool,
    },
    Column {
        column: u32,
        absolute: bool,
    },
    Row {
        row: u32,
        absolute: bool,
    },
}

/// Prefix operators in Spreadsheet formula order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpreadsheetFormulaUnaryOperator {
    Positive,
    Negative,
    ImplicitIntersection,
}

/// Postfix operators in Spreadsheet formula order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpreadsheetFormulaPostfixOperator {
    Percent,
    Spill,
}

/// Binary arithmetic, comparison, concatenation, and reference operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpreadsheetFormulaBinaryOperator {
    Range,
    Intersection,
    Union,
    Power,
    Multiply,
    Divide,
    Add,
    Subtract,
    Concatenate,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}
