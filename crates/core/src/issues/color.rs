pub(super) fn parse_rgb(value: &str) -> Option<[u8; 3]> {
    let value = value.trim().trim_start_matches('#');
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    Some([
        u8::from_str_radix(&value[0..2], 16).ok()?,
        u8::from_str_radix(&value[2..4], 16).ok()?,
        u8::from_str_radix(&value[4..6], 16).ok()?,
    ])
}

pub(super) fn contrast_ratio(left: [u8; 3], right: [u8; 3]) -> f64 {
    let left = relative_luminance(left);
    let right = relative_luminance(right);
    (left.max(right) + 0.05) / (left.min(right) + 0.05)
}

fn relative_luminance(color: [u8; 3]) -> f64 {
    let channel = |value: u8| {
        let value = f64::from(value) / 255.0;
        if value <= 0.04045 {
            value / 12.92
        } else {
            ((value + 0.055) / 1.055).powf(2.4)
        }
    };
    0.2126 * channel(color[0]) + 0.7152 * channel(color[1]) + 0.0722 * channel(color[2])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_only_explicit_rgb_and_computes_wcag_contrast() {
        assert_eq!(parse_rgb("#Aa00fF"), Some([0xaa, 0x00, 0xff]));
        assert_eq!(parse_rgb("scheme:accent1"), None);
        assert!(contrast_ratio([0, 0, 0], [255, 255, 255]) > 20.9);
        assert!(contrast_ratio([20, 20, 20], [30, 30, 30]) < 1.2);
    }
}
