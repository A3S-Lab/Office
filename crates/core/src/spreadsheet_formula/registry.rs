use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

/// Native function result family used for registry inspection and validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SpreadsheetFormulaFunctionReturnKind {
    Scalar,
    ScalarOrArray,
    Array,
}

/// Whether a registered function can change without precedent changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SpreadsheetFormulaFunctionVolatility {
    NonVolatile,
    Volatile,
}

/// Closed typed metadata for one natively implemented Spreadsheet function.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaFunctionDefinition {
    pub name: String,
    pub minimum_arguments: usize,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub maximum_arguments: Option<usize>,
    pub return_kind: SpreadsheetFormulaFunctionReturnKind,
    pub volatility: SpreadsheetFormulaFunctionVolatility,
}

/// Deterministic built-in function registry used by native recalculation.
#[derive(Debug, Clone)]
pub struct SpreadsheetFormulaFunctionRegistry {
    entries: BTreeMap<String, FunctionEntry>,
}

#[derive(Debug, Clone)]
struct FunctionEntry {
    definition: SpreadsheetFormulaFunctionDefinition,
    function: BuiltinFunction,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum BuiltinFunction {
    Sum,
    Average,
    Minimum,
    Maximum,
    Count,
    CountA,
    Absolute,
    SquareRoot,
    Power,
    Modulo,
    Round,
    If,
    IfError,
    And,
    Or,
    Not,
    Concatenate,
    Row,
    Column,
    Sequence,
    Transpose,
    Pi,
    NotAvailable,
}

impl Default for SpreadsheetFormulaFunctionRegistry {
    fn default() -> Self {
        let mut registry = Self {
            entries: BTreeMap::new(),
        };
        for (name, minimum, maximum, return_kind, function) in [
            (
                "SUM",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Sum,
            ),
            (
                "AVERAGE",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Average,
            ),
            (
                "MIN",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Minimum,
            ),
            (
                "MAX",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Maximum,
            ),
            (
                "COUNT",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Count,
            ),
            (
                "COUNTA",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::CountA,
            ),
            (
                "ABS",
                1,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Absolute,
            ),
            (
                "SQRT",
                1,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::SquareRoot,
            ),
            (
                "POWER",
                2,
                Some(2),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Power,
            ),
            (
                "MOD",
                2,
                Some(2),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Modulo,
            ),
            (
                "ROUND",
                2,
                Some(2),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Round,
            ),
            (
                "IF",
                2,
                Some(3),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::If,
            ),
            (
                "IFERROR",
                2,
                Some(2),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::IfError,
            ),
            (
                "AND",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::And,
            ),
            (
                "OR",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Or,
            ),
            (
                "NOT",
                1,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Not,
            ),
            (
                "CONCAT",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Concatenate,
            ),
            (
                "CONCATENATE",
                1,
                Some(255),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Concatenate,
            ),
            (
                "ROW",
                0,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Row,
            ),
            (
                "COLUMN",
                0,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::ScalarOrArray,
                BuiltinFunction::Column,
            ),
            (
                "SEQUENCE",
                1,
                Some(4),
                SpreadsheetFormulaFunctionReturnKind::Array,
                BuiltinFunction::Sequence,
            ),
            (
                "TRANSPOSE",
                1,
                Some(1),
                SpreadsheetFormulaFunctionReturnKind::Array,
                BuiltinFunction::Transpose,
            ),
            (
                "PI",
                0,
                Some(0),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::Pi,
            ),
            (
                "NA",
                0,
                Some(0),
                SpreadsheetFormulaFunctionReturnKind::Scalar,
                BuiltinFunction::NotAvailable,
            ),
        ] {
            registry.insert(
                name,
                minimum,
                maximum,
                return_kind,
                SpreadsheetFormulaFunctionVolatility::NonVolatile,
                function,
            );
        }
        registry
    }
}

impl SpreadsheetFormulaFunctionRegistry {
    pub fn definitions(
        &self,
    ) -> impl ExactSizeIterator<Item = &SpreadsheetFormulaFunctionDefinition> {
        self.entries.values().map(|entry| &entry.definition)
    }

    pub fn get(&self, name: &str) -> Option<&SpreadsheetFormulaFunctionDefinition> {
        self.resolve(name).map(|entry| &entry.definition)
    }

    pub fn contains(&self, name: &str) -> bool {
        self.resolve(name).is_some()
    }

    pub(super) fn function(&self, name: &str) -> Option<BuiltinFunction> {
        self.resolve(name).map(|entry| entry.function)
    }

    fn resolve(&self, name: &str) -> Option<&FunctionEntry> {
        self.entries.get(&normalize_function_name(name))
    }

    fn insert(
        &mut self,
        name: &str,
        minimum_arguments: usize,
        maximum_arguments: Option<usize>,
        return_kind: SpreadsheetFormulaFunctionReturnKind,
        volatility: SpreadsheetFormulaFunctionVolatility,
        function: BuiltinFunction,
    ) {
        self.entries.insert(
            name.to_string(),
            FunctionEntry {
                definition: SpreadsheetFormulaFunctionDefinition {
                    name: name.to_string(),
                    minimum_arguments,
                    maximum_arguments,
                    return_kind,
                    volatility,
                },
                function,
            },
        );
    }
}

pub(super) fn normalize_function_name(name: &str) -> String {
    let mut normalized = name.to_ascii_uppercase();
    loop {
        let stripped = ["_XLFN.", "_XLWS."]
            .into_iter()
            .find_map(|prefix| normalized.strip_prefix(prefix).map(ToOwned::to_owned));
        let Some(stripped) = stripped else {
            return normalized;
        };
        normalized = stripped;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn built_in_registry_is_typed_bounded_and_namespace_aware() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<SpreadsheetFormulaFunctionRegistry>();

        let registry = SpreadsheetFormulaFunctionRegistry::default();
        assert!(registry.contains("sum"));
        assert!(registry.contains("_xlfn._xlws.SEQUENCE"));
        assert!(!registry.contains("SHELL"));
        let sequence = registry.get("SEQUENCE").unwrap();
        assert_eq!(sequence.minimum_arguments, 1);
        assert_eq!(sequence.maximum_arguments, Some(4));
        assert_eq!(
            sequence.return_kind,
            SpreadsheetFormulaFunctionReturnKind::Array
        );
    }
}
