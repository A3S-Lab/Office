use a3s_use_core::UseResult;

use super::{set_string_option, usage_error};

#[derive(Debug, Default)]
pub(in crate::office_native_cli) struct BorderArguments {
    pub all: Option<String>,
    pub color: Option<String>,
    pub left: Option<String>,
    pub left_color: Option<String>,
    pub right: Option<String>,
    pub right_color: Option<String>,
    pub top: Option<String>,
    pub top_color: Option<String>,
    pub bottom: Option<String>,
    pub bottom_color: Option<String>,
    pub diagonal: Option<String>,
    pub diagonal_color: Option<String>,
    pub diagonal_up: Option<String>,
    pub diagonal_down: Option<String>,
}

impl BorderArguments {
    pub(super) fn parse(&mut self, option: &str, args: &[String], index: usize) -> UseResult<()> {
        let (target, canonical) = match option {
            "--border" | "--border-all" => (&mut self.all, "--border-all"),
            "--border-color" => (&mut self.color, "--border-color"),
            "--border-left" => (&mut self.left, "--border-left"),
            "--border-left-color" => (&mut self.left_color, "--border-left-color"),
            "--border-right" => (&mut self.right, "--border-right"),
            "--border-right-color" => (&mut self.right_color, "--border-right-color"),
            "--border-top" => (&mut self.top, "--border-top"),
            "--border-top-color" => (&mut self.top_color, "--border-top-color"),
            "--border-bottom" => (&mut self.bottom, "--border-bottom"),
            "--border-bottom-color" => (&mut self.bottom_color, "--border-bottom-color"),
            "--border-diagonal" => (&mut self.diagonal, "--border-diagonal"),
            "--border-diagonal-color" => (&mut self.diagonal_color, "--border-diagonal-color"),
            "--border-diagonal-up" => (&mut self.diagonal_up, "--border-diagonal-up"),
            "--border-diagonal-down" => (&mut self.diagonal_down, "--border-diagonal-down"),
            _ => {
                return Err(usage_error(format!(
                    "unsupported native Spreadsheet border option '{option}'"
                )))
            }
        };
        set_string_option(target, args, index, canonical)
    }

    pub(in crate::office_native_cli) fn is_present(&self) -> bool {
        self.all.is_some()
            || self.color.is_some()
            || self.left.is_some()
            || self.left_color.is_some()
            || self.right.is_some()
            || self.right_color.is_some()
            || self.top.is_some()
            || self.top_color.is_some()
            || self.bottom.is_some()
            || self.bottom_color.is_some()
            || self.diagonal.is_some()
            || self.diagonal_color.is_some()
            || self.diagonal_up.is_some()
            || self.diagonal_down.is_some()
    }
}
