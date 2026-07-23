pub(super) fn wrap(value: &str, max_characters: usize) -> Vec<String> {
    let max_characters = max_characters.max(1);
    if value.is_empty() {
        return Vec::new();
    }

    let mut output = Vec::new();
    for logical_line in value.split('\n') {
        if logical_line.is_empty() {
            output.push(String::new());
            continue;
        }

        let mut line = String::new();
        let mut characters = 0_usize;
        for character in logical_line.chars() {
            if characters == max_characters {
                output.push(line);
                line = String::new();
                characters = 0;
            }
            line.push(character);
            characters += 1;
        }
        output.push(line);
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wrapping_preserves_unicode_and_explicit_empty_lines() {
        assert_eq!(wrap("A😀BC\n\nD", 2), ["A😀", "BC", "", "D"]);
        assert!(wrap("", 2).is_empty());
    }
}
