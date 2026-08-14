use std::collections::BTreeMap;
use std::str::FromStr;

use a3s_use_core::{UseError, UseResult};

pub(super) fn parse_value<T: FromStr<Err = UseError>>(value: &str, option: &str) -> UseResult<T> {
    value.parse::<T>().map_err(|error| {
        usage_error(format!("{option} is invalid: {}", error.message))
            .with_detail("causeCode", error.code)
    })
}

#[derive(Debug, Default)]
pub(super) struct ParsedOptions {
    pub(super) positionals: Vec<String>,
    values: BTreeMap<String, Vec<String>>,
}

impl ParsedOptions {
    pub(super) fn parse(args: &[String]) -> UseResult<Self> {
        let mut parsed = Self::default();
        let mut index = 1;
        while index < args.len() {
            let argument = &args[index];
            if argument == "--" {
                parsed.positionals.extend_from_slice(&args[index + 1..]);
                break;
            }
            if let Some(option) = argument.strip_prefix("--") {
                let value = args
                    .get(index + 1)
                    .ok_or_else(|| usage_error(format!("--{option} requires a value")))?;
                if value.starts_with("--") {
                    return Err(usage_error(format!("--{option} requires a value")));
                }
                parsed
                    .values
                    .entry(option.to_owned())
                    .or_default()
                    .push(value.clone());
                index += 2;
                continue;
            }
            parsed.positionals.push(argument.clone());
            index += 1;
        }
        Ok(parsed)
    }

    pub(super) fn reject_unknown(&self, allowed: &[&str]) -> UseResult<()> {
        if let Some(option) = self
            .values
            .keys()
            .find(|option| !allowed.contains(&option.as_str()))
        {
            return Err(usage_error(format!("Unknown option '--{option}'")));
        }
        Ok(())
    }

    pub(super) fn value(&self, option: &str) -> UseResult<Option<&str>> {
        match self.values.get(option).map(Vec::as_slice) {
            None => Ok(None),
            Some([value]) => Ok(Some(value)),
            Some(_) => Err(usage_error(format!(
                "--{option} may be specified only once"
            ))),
        }
    }

    pub(super) fn required(&self, option: &str) -> UseResult<&str> {
        self.value(option)?
            .ok_or_else(|| usage_error(format!("--{option} is required")))
    }

    pub(super) fn one_positional(&self, label: &str) -> UseResult<String> {
        match self.positionals.as_slice() {
            [value] => Ok(value.clone()),
            _ => Err(usage_error(format!("Expected exactly one {label}."))),
        }
    }
}

pub(super) fn usage_error(message: impl Into<String>) -> UseError {
    collaboration_cli_error("office.collaboration.usage", message)
        .with_suggestion("Run 'a3s-office collab --help'.")
}

pub(super) fn collaboration_cli_error(code: &'static str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}
