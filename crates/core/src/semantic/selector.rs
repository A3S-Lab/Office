use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operator {
    Equal,
    NotEqual,
    Contains,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
    Exists,
}

#[derive(Debug, Clone, PartialEq)]
struct Predicate {
    key: String,
    operator: Operator,
    value: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
enum FilterExpression {
    Predicate(Predicate),
    And(Vec<FilterExpression>),
    Or(Vec<FilterExpression>),
    Not(Box<FilterExpression>),
}

#[derive(Debug, Clone, PartialEq)]
enum Pseudo {
    Contains(String),
    Empty,
    Text,
    NoAlt,
    Has(String),
}

#[derive(Debug, Clone, PartialEq)]
struct SimpleSelector {
    kind: String,
    position: Option<usize>,
    last: bool,
    filters: Vec<FilterExpression>,
    pseudos: Vec<Pseudo>,
}

#[derive(Debug, Clone, PartialEq)]
struct Selector {
    chain: Vec<SimpleSelector>,
    path_prefix: Option<String>,
}

pub(super) fn query<'a>(
    root: &'a DocumentNode,
    expression: &str,
) -> UseResult<Vec<&'a DocumentNode>> {
    let expressions = split_top_level(expression, ',')?;
    if expressions.is_empty() {
        return Err(selector_error("Selector cannot be empty."));
    }
    let selectors = expressions
        .into_iter()
        .map(parse_selector)
        .collect::<UseResult<Vec<_>>>()?;
    let mut matches = Vec::new();
    let mut paths = BTreeSet::new();
    walk(root, &[], &selectors, &mut paths, &mut matches);
    Ok(matches)
}

fn walk<'a>(
    node: &'a DocumentNode,
    ancestors: &[&'a DocumentNode],
    selectors: &[Selector],
    paths: &mut BTreeSet<String>,
    output: &mut Vec<&'a DocumentNode>,
) {
    if selectors
        .iter()
        .any(|selector| selector_matches(selector, node, ancestors))
        && paths.insert(node.path.clone())
    {
        output.push(node);
    }
    let mut child_ancestors = ancestors.to_vec();
    child_ancestors.push(node);
    for child in &node.children {
        walk(child, &child_ancestors, selectors, paths, output);
    }
}

fn selector_matches(selector: &Selector, node: &DocumentNode, ancestors: &[&DocumentNode]) -> bool {
    if selector
        .path_prefix
        .as_deref()
        .is_some_and(|prefix| !node.path.to_ascii_lowercase().starts_with(prefix))
    {
        return false;
    }
    let Some(last) = selector.chain.last() else {
        return false;
    };
    let siblings = ancestors.last().map(|parent| parent.children.as_slice());
    if !simple_matches(last, node, siblings) {
        return false;
    }
    if selector.chain.len() == 1 {
        return true;
    }
    if ancestors.len() < selector.chain.len() - 1 {
        return false;
    }
    let start = ancestors.len() - (selector.chain.len() - 1);
    selector.chain[..selector.chain.len() - 1]
        .iter()
        .zip(&ancestors[start..])
        .enumerate()
        .all(|(index, (simple, ancestor))| {
            let ancestor_siblings = if start + index == 0 {
                None
            } else {
                Some(ancestors[start + index - 1].children.as_slice())
            };
            simple_matches(simple, ancestor, ancestor_siblings)
        })
}

fn simple_matches(
    selector: &SimpleSelector,
    node: &DocumentNode,
    siblings: Option<&[DocumentNode]>,
) -> bool {
    if !kind_matches(&selector.kind, node) {
        return false;
    }
    if let Some(position) = selector.position {
        if node_position(node) != Some(position) {
            return false;
        }
    }
    if selector.last
        && siblings.is_some_and(|siblings| {
            siblings
                .iter()
                .rev()
                .find(|sibling| kind_matches(&selector.kind, sibling))
                .is_some_and(|last| last.path != node.path)
        })
    {
        return false;
    }
    selector
        .filters
        .iter()
        .all(|filter| filter_matches(filter, node))
        && selector.pseudos.iter().all(|pseudo| match pseudo {
            Pseudo::Contains(text) => contains_case_insensitive(&node.text, text),
            Pseudo::Empty => node.text.trim().is_empty(),
            Pseudo::Text => !node.text.trim().is_empty(),
            Pseudo::NoAlt => node
                .format
                .get("alt")
                .is_none_or(|alt| alt.trim().is_empty()),
            Pseudo::Has(key) => node_value(node, key).is_some_and(|value| !value.is_empty()),
        })
}

fn filter_matches(expression: &FilterExpression, node: &DocumentNode) -> bool {
    match expression {
        FilterExpression::Predicate(predicate) => predicate_matches(predicate, node),
        FilterExpression::And(expressions) => expressions
            .iter()
            .all(|expression| filter_matches(expression, node)),
        FilterExpression::Or(expressions) => expressions
            .iter()
            .any(|expression| filter_matches(expression, node)),
        FilterExpression::Not(expression) => !filter_matches(expression, node),
    }
}

fn predicate_matches(predicate: &Predicate, node: &DocumentNode) -> bool {
    let actual = node_value(node, &predicate.key);
    if predicate.operator == Operator::Exists {
        return actual.is_some_and(|value| !value.is_empty() && value != "false");
    }
    let Some(actual) = actual else {
        return predicate.operator == Operator::NotEqual;
    };
    let expected = predicate.value.as_deref().unwrap_or_default();
    match predicate.operator {
        Operator::Equal => values_equal(actual, expected),
        Operator::NotEqual => !values_equal(actual, expected),
        Operator::Contains => contains_case_insensitive(actual, expected),
        Operator::Greater => compare_numbers(actual, expected).is_some_and(|order| order > 0.0),
        Operator::GreaterOrEqual => {
            compare_numbers(actual, expected).is_some_and(|order| order >= 0.0)
        }
        Operator::Less => compare_numbers(actual, expected).is_some_and(|order| order < 0.0),
        Operator::LessOrEqual => {
            compare_numbers(actual, expected).is_some_and(|order| order <= 0.0)
        }
        Operator::Exists => unreachable!(),
    }
}

fn node_value<'a>(node: &'a DocumentNode, key: &str) -> Option<&'a str> {
    let key = strip_quotes(key.trim());
    match key.to_ascii_lowercase().as_str() {
        "text" | "value" => Some(&node.text),
        "type" => node
            .format
            .get("valueType")
            .map(String::as_str)
            .or_else(|| {
                (matches!(
                    node.node_type,
                    OfficeNodeType::DataValidation | OfficeNodeType::ConditionalFormatting
                ))
                .then(|| node.format.get("type").map(String::as_str))
                .flatten()
            })
            .or_else(|| Some(node.node_type.label())),
        "style" => node.style.as_deref(),
        "path" => Some(&node.path),
        _ => node
            .format
            .iter()
            .find(|(candidate, _)| candidate.eq_ignore_ascii_case(key))
            .map(|(_, value)| value.as_str()),
    }
}

fn kind_matches(kind: &str, node: &DocumentNode) -> bool {
    let kind = kind.to_ascii_lowercase();
    match kind.as_str() {
        "*" => true,
        "p" | "paragraph" => node.node_type == OfficeNodeType::Paragraph,
        "r" | "run" => node.node_type == OfficeNodeType::Run,
        "tbl" | "table" => node.node_type == OfficeNodeType::Table,
        "tr" => node.node_type == OfficeNodeType::TableRow,
        "tc" => node.node_type == OfficeNodeType::TableCell,
        "row" => matches!(
            node.node_type,
            OfficeNodeType::Row | OfficeNodeType::TableRow
        ),
        "col" | "column" => node.node_type == OfficeNodeType::Column,
        "range" => node.node_type == OfficeNodeType::Range,
        "cell" => matches!(
            node.node_type,
            OfficeNodeType::Cell | OfficeNodeType::TableCell
        ),
        "datavalidation" | "data-validation" | "validation" => {
            node.node_type == OfficeNodeType::DataValidation
        }
        "namedrange" | "named-range" | "definedname" | "defined-name" => {
            node.node_type == OfficeNodeType::NamedRange
        }
        "namedranges" | "named-range-list" | "definednames" => {
            node.node_type == OfficeNodeType::NamedRangeCollection
        }
        "sheet" | "worksheet" => node.node_type == OfficeNodeType::Worksheet,
        "freeze" | "frozenpane" | "frozen-pane" => node.node_type == OfficeNodeType::FrozenPane,
        "slide" => node.node_type == OfficeNodeType::Slide,
        "shape" | "textbox" => matches!(
            node.node_type,
            OfficeNodeType::Shape | OfficeNodeType::Placeholder
        ),
        "title" => {
            matches!(
                node.node_type,
                OfficeNodeType::Shape | OfficeNodeType::Placeholder
            ) && node
                .format
                .get("title")
                .is_some_and(|value| value == "true")
        }
        "placeholder" => node.node_type == OfficeNodeType::Placeholder,
        "picture" | "pic" | "image" | "img" => node.node_type == OfficeNodeType::Picture,
        "chart" => node.node_type == OfficeNodeType::Chart,
        "connector" | "connection" => node.node_type == OfficeNodeType::Connector,
        "group" => node.node_type == OfficeNodeType::Group,
        "hyperlink" => node.node_type == OfficeNodeType::Hyperlink,
        "comment" | "note-comment" => node.node_type == OfficeNodeType::Comment,
        "notes" => node.node_type == OfficeNodeType::Notes,
        "header" => node.node_type == OfficeNodeType::Header,
        "footer" => node.node_type == OfficeNodeType::Footer,
        _ if kind
            .chars()
            .all(|character| character.is_ascii_alphabetic())
            && node.node_type == OfficeNodeType::Cell =>
        {
            node.format
                .get("column")
                .is_some_and(|column| column.eq_ignore_ascii_case(&kind))
        }
        _ => node.tag.eq_ignore_ascii_case(&kind),
    }
}

fn parse_selector(expression: &str) -> UseResult<Selector> {
    let mut expression = expression.trim();
    if expression.is_empty() {
        return Err(selector_error("Selector cannot be empty."));
    }
    let mut path_prefix = None;
    if let Some((sheet, rest)) = split_sheet_prefix(expression) {
        path_prefix = Some(format!("/{}/", sheet.trim_matches('/')).to_ascii_lowercase());
        expression = rest;
    } else if let Some((sheet, rest)) = split_sheet_path_prefix(expression) {
        path_prefix = Some(format!("/{sheet}/").to_ascii_lowercase());
        expression = rest;
    }
    let chain = split_top_level(expression, '>')?
        .into_iter()
        .map(parse_simple_selector)
        .collect::<UseResult<Vec<_>>>()?;
    if chain.is_empty() {
        return Err(selector_error("Selector has no element name."));
    }
    Ok(Selector { chain, path_prefix })
}

fn parse_simple_selector(expression: &str) -> UseResult<SimpleSelector> {
    let expression = expression.trim();
    let kind_end = expression
        .char_indices()
        .find(|(_, character)| matches!(character, '[' | ':'))
        .map_or(expression.len(), |(index, _)| index);
    let kind = expression[..kind_end].trim();
    if kind.is_empty()
        || !kind
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || matches!(character, '_' | '-'))
    {
        return Err(selector_error(format!(
            "Selector element name '{kind}' is invalid."
        )));
    }
    let mut selector = SimpleSelector {
        kind: kind.to_string(),
        position: None,
        last: false,
        filters: Vec::new(),
        pseudos: Vec::new(),
    };
    let mut offset = kind_end;
    while offset < expression.len() {
        match expression.as_bytes()[offset] {
            b'[' => {
                let end = matching_delimiter(expression, offset, '[', ']')?;
                let content = expression[offset + 1..end].trim();
                if content.is_empty() {
                    return Err(selector_error("Selector filter cannot be empty."));
                }
                if content == "last()" {
                    selector.last = true;
                } else if content.chars().all(|character| character.is_ascii_digit()) {
                    let position = content.parse::<usize>().map_err(|error| {
                        selector_error(format!("Invalid selector position: {error}"))
                    })?;
                    if position == 0 {
                        return Err(selector_error("Selector positions are one-based."));
                    }
                    selector.position = Some(position);
                } else {
                    selector.filters.push(parse_filter_expression(content)?);
                }
                offset = end + 1;
            }
            b':' => {
                let (pseudo, consumed) = parse_pseudo(&expression[offset + 1..])?;
                selector.pseudos.push(pseudo);
                offset += consumed + 1;
            }
            byte if byte.is_ascii_whitespace() => offset += 1,
            _ => {
                return Err(selector_error(format!(
                    "Unexpected selector syntax near '{}'.",
                    &expression[offset..]
                )));
            }
        }
    }
    Ok(selector)
}

fn parse_filter_expression(expression: &str) -> UseResult<FilterExpression> {
    let expression = trim_outer_parentheses(expression.trim())?;
    let alternatives = split_keyword(expression, "or")?;
    if alternatives.len() > 1 {
        return alternatives
            .into_iter()
            .map(parse_filter_expression)
            .collect::<UseResult<Vec<_>>>()
            .map(FilterExpression::Or);
    }
    let conjunctions = split_keyword(expression, "and")?;
    if conjunctions.len() > 1 {
        return conjunctions
            .into_iter()
            .map(parse_filter_expression)
            .collect::<UseResult<Vec<_>>>()
            .map(FilterExpression::And);
    }
    if expression.len() >= 5
        && expression[..4].eq_ignore_ascii_case("not(")
        && matching_delimiter(expression, 3, '(', ')')? == expression.len() - 1
    {
        return Ok(FilterExpression::Not(Box::new(parse_filter_expression(
            &expression[4..expression.len() - 1],
        )?)));
    }
    parse_predicate(expression).map(FilterExpression::Predicate)
}

fn parse_predicate(expression: &str) -> UseResult<Predicate> {
    let expression = expression.replace("\\!=", "!=");
    for (token, operator) in [
        (">=", Operator::GreaterOrEqual),
        ("<=", Operator::LessOrEqual),
        ("!=", Operator::NotEqual),
        ("~=", Operator::Contains),
        ("=", Operator::Equal),
        (">", Operator::Greater),
        ("<", Operator::Less),
    ] {
        if let Some(index) = find_operator(&expression, token) {
            let key = strip_quotes(expression[..index].trim());
            let value = strip_quotes(expression[index + token.len()..].trim());
            if key.is_empty() || value.is_empty() {
                return Err(selector_error(format!(
                    "Selector predicate '{expression}' is incomplete."
                )));
            }
            return Ok(Predicate {
                key: key.to_string(),
                operator,
                value: Some(value.to_string()),
            });
        }
    }
    let key = strip_quotes(expression.trim());
    if key.is_empty() {
        return Err(selector_error("Selector predicate cannot be empty."));
    }
    Ok(Predicate {
        key: key.to_string(),
        operator: Operator::Exists,
        value: None,
    })
}

fn parse_pseudo(expression: &str) -> UseResult<(Pseudo, usize)> {
    for (name, constructor) in [("contains", 0_u8), ("has", 1_u8)] {
        if expression.len() > name.len()
            && expression[..name.len()].eq_ignore_ascii_case(name)
            && expression.as_bytes()[name.len()] == b'('
        {
            let end = matching_delimiter(expression, name.len(), '(', ')')?;
            let value = strip_quotes(expression[name.len() + 1..end].trim());
            if value.is_empty() {
                return Err(selector_error(format!(":{name} requires a value.")));
            }
            let pseudo = if constructor == 0 {
                Pseudo::Contains(value.to_string())
            } else {
                Pseudo::Has(value.to_string())
            };
            return Ok((pseudo, end + 1));
        }
    }
    let end = expression
        .char_indices()
        .find(|(_, character)| {
            matches!(character, '[' | ':' | '>' | ',') || character.is_whitespace()
        })
        .map_or(expression.len(), |(index, _)| index);
    let name = &expression[..end];
    let pseudo = match name.to_ascii_lowercase().as_str() {
        "empty" => Pseudo::Empty,
        "text" => Pseudo::Text,
        "no-alt" => Pseudo::NoAlt,
        _ if !name.is_empty() => Pseudo::Contains(name.to_string()),
        _ => return Err(selector_error("Pseudo-selector cannot be empty.")),
    };
    Ok((pseudo, end))
}

fn split_top_level(expression: &str, delimiter: char) -> UseResult<Vec<&str>> {
    let mut output = Vec::new();
    let mut start = 0;
    let mut brackets = 0_i32;
    let mut parentheses = 0_i32;
    let mut quote = None;
    let mut escaped = false;
    for (index, character) in expression.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        if character == '\\' {
            escaped = true;
            continue;
        }
        if let Some(active) = quote {
            if character == active {
                quote = None;
            }
            continue;
        }
        if matches!(character, '\'' | '"') {
            quote = Some(character);
            continue;
        }
        match character {
            '[' => brackets += 1,
            ']' => brackets -= 1,
            '(' => parentheses += 1,
            ')' => parentheses -= 1,
            _ => {}
        }
        if brackets < 0 || parentheses < 0 {
            return Err(selector_error(
                "Selector has an unmatched closing delimiter.",
            ));
        }
        if character == delimiter && brackets == 0 && parentheses == 0 {
            let value = expression[start..index].trim();
            if value.is_empty() {
                return Err(selector_error("Selector contains an empty branch."));
            }
            output.push(value);
            start = index + character.len_utf8();
        }
    }
    if quote.is_some() || brackets != 0 || parentheses != 0 {
        return Err(selector_error(
            "Selector has an unclosed quote or delimiter.",
        ));
    }
    let value = expression[start..].trim();
    if !value.is_empty() {
        output.push(value);
    }
    Ok(output)
}

fn split_keyword<'a>(expression: &'a str, keyword: &str) -> UseResult<Vec<&'a str>> {
    let mut output = Vec::new();
    let mut start = 0;
    let mut parentheses = 0_i32;
    let mut quote = None;
    let bytes = expression.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        let character = expression[index..]
            .chars()
            .next()
            .ok_or_else(|| selector_error("Selector contains invalid text."))?;
        if let Some(active) = quote {
            if character == active {
                quote = None;
            }
            index += character.len_utf8();
            continue;
        }
        if matches!(character, '\'' | '"') {
            quote = Some(character);
            index += character.len_utf8();
            continue;
        }
        match character {
            '(' => parentheses += 1,
            ')' => parentheses -= 1,
            _ => {}
        }
        if parentheses == 0
            && expression[index..]
                .get(..keyword.len())
                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(keyword))
            && word_boundary(expression, index)
            && word_boundary(expression, index + keyword.len())
        {
            output.push(expression[start..index].trim());
            index += keyword.len();
            start = index;
            continue;
        }
        index += character.len_utf8();
    }
    if parentheses != 0 || quote.is_some() {
        return Err(selector_error("Selector predicate has unclosed grouping."));
    }
    output.push(expression[start..].trim());
    Ok(output)
}

fn matching_delimiter(expression: &str, start: usize, open: char, close: char) -> UseResult<usize> {
    let mut depth = 0_i32;
    let mut quote = None;
    let mut escaped = false;
    for (offset, character) in expression[start..].char_indices() {
        let index = start + offset;
        if escaped {
            escaped = false;
            continue;
        }
        if character == '\\' {
            escaped = true;
            continue;
        }
        if let Some(active) = quote {
            if character == active {
                quote = None;
            }
            continue;
        }
        if matches!(character, '\'' | '"') {
            quote = Some(character);
        } else if character == open {
            depth += 1;
        } else if character == close {
            depth -= 1;
            if depth == 0 {
                return Ok(index);
            }
        }
    }
    Err(selector_error(format!(
        "Selector is missing closing '{close}'."
    )))
}

fn trim_outer_parentheses(expression: &str) -> UseResult<&str> {
    if expression.starts_with('(')
        && matching_delimiter(expression, 0, '(', ')')? == expression.len() - 1
    {
        Ok(expression[1..expression.len() - 1].trim())
    } else {
        Ok(expression)
    }
}

fn find_operator(expression: &str, operator: &str) -> Option<usize> {
    let mut quote = None;
    for (index, character) in expression.char_indices() {
        if let Some(active) = quote {
            if character == active {
                quote = None;
            }
            continue;
        }
        if matches!(character, '\'' | '"') {
            quote = Some(character);
            continue;
        }
        if expression[index..].starts_with(operator) {
            return Some(index);
        }
    }
    None
}

fn split_sheet_prefix(expression: &str) -> Option<(&str, &str)> {
    let index = expression.find('!')?;
    let sheet = expression[..index].trim();
    let rest = expression[index + 1..].trim();
    (!sheet.is_empty() && !rest.is_empty()).then_some((sheet, rest))
}

fn split_sheet_path_prefix(expression: &str) -> Option<(&str, &str)> {
    let expression = expression.strip_prefix('/')?;
    let (sheet, rest) = expression.split_once('/')?;
    (!sheet.is_empty() && !rest.is_empty()).then_some((sheet, rest))
}

fn node_position(node: &DocumentNode) -> Option<usize> {
    let segment = node.path.rsplit('/').next()?;
    let start = segment.rfind('[')?;
    let end = segment.rfind(']')?;
    segment[start + 1..end].parse().ok()
}

fn values_equal(left: &str, right: &str) -> bool {
    let left = left.trim_start_matches('#');
    let right = right.trim_start_matches('#');
    left.eq_ignore_ascii_case(right)
}

fn contains_case_insensitive(haystack: &str, needle: &str) -> bool {
    haystack
        .to_lowercase()
        .contains(needle.to_lowercase().as_str())
}

fn compare_numbers(left: &str, right: &str) -> Option<f64> {
    let (left, left_unit) = number_with_unit(left)?;
    let (right, right_unit) = number_with_unit(right)?;
    let left = normalize_dimension(left, left_unit)?;
    let right = normalize_dimension(right, right_unit)?;
    Some(left - right)
}

fn number_with_unit(value: &str) -> Option<(f64, &str)> {
    let value = value.trim();
    let split = value
        .char_indices()
        .find(|(_, character)| !character.is_ascii_digit() && !matches!(character, '.' | '-' | '+'))
        .map_or(value.len(), |(index, _)| index);
    Some((value[..split].parse().ok()?, value[split..].trim()))
}

fn normalize_dimension(value: f64, unit: &str) -> Option<f64> {
    match unit.to_ascii_lowercase().as_str() {
        "" | "pt" => Some(value),
        "in" => Some(value * 72.0),
        "cm" => Some(value * 72.0 / 2.54),
        "mm" => Some(value * 72.0 / 25.4),
        "px" => Some(value * 72.0 / 96.0),
        _ => None,
    }
}

fn strip_quotes(value: &str) -> &str {
    if value.len() >= 2
        && ((value.starts_with('"') && value.ends_with('"'))
            || (value.starts_with('\'') && value.ends_with('\'')))
    {
        &value[1..value.len() - 1]
    } else {
        value
    }
}

fn word_boundary(expression: &str, index: usize) -> bool {
    if index == 0 || index == expression.len() {
        return true;
    }
    expression[..index]
        .chars()
        .next_back()
        .is_some_and(char::is_whitespace)
        || expression[index..]
            .chars()
            .next()
            .is_some_and(char::is_whitespace)
}

fn selector_error(message: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error("use.office.selector_invalid", message)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn tree() -> DocumentNode {
        let mut root = DocumentNode::new("/", "document", OfficeNodeType::Document);
        let mut body = DocumentNode::new("/body", "body", OfficeNodeType::Body);
        for (index, style, text, bold) in [
            (1, "Heading1", "Quarterly Results", "true"),
            (2, "Normal", "Summary", "false"),
        ] {
            let mut paragraph =
                DocumentNode::new(format!("/body/p[{index}]"), "p", OfficeNodeType::Paragraph);
            paragraph.style = Some(style.to_string());
            paragraph.text = text.to_string();
            let mut run =
                DocumentNode::new(format!("/body/p[{index}]/r[1]"), "r", OfficeNodeType::Run);
            run.text = text.to_string();
            run.format.insert("bold".to_string(), bold.to_string());
            paragraph.children.push(run);
            body.children.push(paragraph);
        }
        root.children.push(body);
        root.normalize();
        root
    }

    #[test]
    fn supports_alias_filters_children_pseudos_and_unions() {
        let tree = tree();
        assert_eq!(query(&tree, "p[style=Heading1]").unwrap().len(), 1);
        assert_eq!(
            query(&tree, "p[text~=quarterly] > r[bold=true]")
                .unwrap()
                .len(),
            1
        );
        assert_eq!(query(&tree, "p:contains(RESULTS)").unwrap().len(), 1);
        assert_eq!(query(&tree, "p[1],p[2]").unwrap().len(), 2);
        assert_eq!(query(&tree, "p[last()]").unwrap()[0].path, "/body/p[2]");
    }

    #[test]
    fn supports_boolean_filter_expressions() {
        let tree = tree();
        assert_eq!(
            query(
                &tree,
                "p[(style=Heading1 or style=Title) and not(text=Missing)]"
            )
            .unwrap()
            .len(),
            1
        );
    }

    #[test]
    fn rejects_malformed_selectors() {
        assert_eq!(
            query(&tree(), "p[style=Heading1").unwrap_err().code,
            "use.office.selector_invalid"
        );
        assert_eq!(
            query(&tree(), "p[0]").unwrap_err().code,
            "use.office.selector_invalid"
        );
    }
}
