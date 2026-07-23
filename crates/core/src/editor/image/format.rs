use a3s_use_core::UseResult;

use super::{image_invalid, MAX_IMAGE_BYTES};
use crate::editor::NativeOfficeImageFormat;

const MAX_IMAGE_DIMENSION_PX: u32 = 100_000;
const MAX_IMAGE_PIXELS: u64 = 500_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ImageMetadata {
    pub format: NativeOfficeImageFormat,
    pub width_px: u32,
    pub height_px: u32,
}

pub(crate) fn inspect_image(
    bytes: &[u8],
    expected: Option<NativeOfficeImageFormat>,
) -> UseResult<ImageMetadata> {
    if bytes.is_empty() || bytes.len() > MAX_IMAGE_BYTES {
        return Err(image_invalid(format!(
            "Native Office images must contain between 1 and {MAX_IMAGE_BYTES} bytes."
        )));
    }
    let metadata = if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        inspect_png(bytes)?
    } else if bytes.starts_with(&[0xff, 0xd8]) {
        inspect_jpeg(bytes)?
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        inspect_gif(bytes)?
    } else {
        return Err(image_invalid(
            "Native Office images must be valid PNG, JPEG, or GIF data.",
        ));
    };
    if expected.is_some_and(|expected| expected != metadata.format) {
        return Err(image_invalid(
            "The declared image format does not match the decoded image data.",
        ));
    }
    validate_pixel_bounds(metadata.width_px, metadata.height_px)?;
    Ok(metadata)
}

pub(in crate::editor) fn validate_pixel_bounds(width: u32, height: u32) -> UseResult<()> {
    let pixels = u64::from(width) * u64::from(height);
    if width == 0
        || height == 0
        || width > MAX_IMAGE_DIMENSION_PX
        || height > MAX_IMAGE_DIMENSION_PX
        || pixels > MAX_IMAGE_PIXELS
    {
        return Err(image_invalid(format!(
            "Image dimensions must be positive, at most {MAX_IMAGE_DIMENSION_PX}px per axis, and at most {MAX_IMAGE_PIXELS} pixels."
        )));
    }
    Ok(())
}

fn inspect_png(bytes: &[u8]) -> UseResult<ImageMetadata> {
    let mut cursor = 8_usize;
    let mut dimensions = None;
    let mut saw_data = false;
    let mut saw_end = false;
    while cursor < bytes.len() {
        let header = bytes
            .get(cursor..cursor.saturating_add(8))
            .ok_or_else(|| image_invalid("PNG data contains a truncated chunk header."))?;
        let length = u32::from_be_bytes(header[..4].try_into().expect("four-byte PNG length"));
        let length = usize::try_from(length)
            .map_err(|_| image_invalid("PNG chunk length does not fit this platform."))?;
        let chunk_end = cursor
            .checked_add(12)
            .and_then(|value| value.checked_add(length))
            .ok_or_else(|| image_invalid("PNG chunk length overflowed."))?;
        let chunk = bytes
            .get(cursor..chunk_end)
            .ok_or_else(|| image_invalid("PNG data contains a truncated chunk."))?;
        let kind = &chunk[4..8];
        let expected_crc = u32::from_be_bytes(
            chunk[8 + length..12 + length]
                .try_into()
                .expect("four-byte PNG CRC"),
        );
        if png_crc(&chunk[4..8 + length]) != expected_crc {
            return Err(image_invalid(
                "PNG data contains a chunk with an invalid CRC.",
            ));
        }
        if cursor == 8 && (kind != b"IHDR" || length != 13) {
            return Err(image_invalid(
                "PNG data must begin with a 13-byte IHDR chunk.",
            ));
        }
        match kind {
            b"IHDR" => {
                if dimensions.is_some() || length != 13 {
                    return Err(image_invalid("PNG data contains an invalid IHDR chunk."));
                }
                dimensions = Some((
                    u32::from_be_bytes(chunk[8..12].try_into().expect("PNG width")),
                    u32::from_be_bytes(chunk[12..16].try_into().expect("PNG height")),
                ));
            }
            b"IDAT" => saw_data = true,
            b"IEND" => {
                if length != 0 || chunk_end != bytes.len() {
                    return Err(image_invalid(
                        "PNG IEND must be empty and terminate the image.",
                    ));
                }
                saw_end = true;
            }
            _ => {}
        }
        cursor = chunk_end;
        if saw_end {
            break;
        }
    }
    let (width_px, height_px) = dimensions
        .filter(|_| saw_data && saw_end)
        .ok_or_else(|| image_invalid("PNG data is missing IHDR, IDAT, or IEND."))?;
    Ok(ImageMetadata {
        format: NativeOfficeImageFormat::Png,
        width_px,
        height_px,
    })
}

fn png_crc(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                0xedb8_8320 ^ (crc >> 1)
            };
        }
    }
    !crc
}

fn inspect_jpeg(bytes: &[u8]) -> UseResult<ImageMetadata> {
    if bytes.len() < 6 || !bytes.ends_with(&[0xff, 0xd9]) {
        return Err(image_invalid(
            "JPEG data is missing its end-of-image marker.",
        ));
    }
    let mut cursor = 2_usize;
    let mut dimensions = None;
    let mut saw_scan = false;
    while cursor + 1 < bytes.len() - 2 {
        if bytes[cursor] != 0xff {
            if saw_scan {
                cursor += 1;
                continue;
            }
            return Err(image_invalid(
                "JPEG data contains an invalid marker boundary.",
            ));
        }
        while cursor < bytes.len() && bytes[cursor] == 0xff {
            cursor += 1;
        }
        let marker = *bytes
            .get(cursor)
            .ok_or_else(|| image_invalid("JPEG data ends inside a marker."))?;
        cursor += 1;
        if marker == 0x00 || (saw_scan && (0xd0..=0xd7).contains(&marker)) {
            continue;
        }
        if marker == 0x01 || marker == 0xd8 {
            continue;
        }
        let length_bytes = bytes
            .get(cursor..cursor.saturating_add(2))
            .ok_or_else(|| image_invalid("JPEG data contains a truncated segment length."))?;
        let length = usize::from(u16::from_be_bytes(
            length_bytes.try_into().expect("JPEG segment length"),
        ));
        if length < 2 {
            return Err(image_invalid(
                "JPEG segment length must include its length field.",
            ));
        }
        let segment_end = cursor
            .checked_add(length)
            .ok_or_else(|| image_invalid("JPEG segment length overflowed."))?;
        let segment = bytes
            .get(cursor..segment_end)
            .ok_or_else(|| image_invalid("JPEG data contains a truncated segment."))?;
        if is_start_of_frame(marker) {
            if segment.len() < 7 {
                return Err(image_invalid("JPEG start-of-frame data is truncated."));
            }
            dimensions = Some((
                u32::from(u16::from_be_bytes([segment[5], segment[6]])),
                u32::from(u16::from_be_bytes([segment[3], segment[4]])),
            ));
        }
        saw_scan |= marker == 0xda;
        cursor = segment_end;
    }
    let (width_px, height_px) = dimensions
        .filter(|_| saw_scan)
        .ok_or_else(|| image_invalid("JPEG data must contain supported frame and scan headers."))?;
    Ok(ImageMetadata {
        format: NativeOfficeImageFormat::Jpeg,
        width_px,
        height_px,
    })
}

fn is_start_of_frame(marker: u8) -> bool {
    matches!(
        marker,
        0xc0..=0xc3 | 0xc5..=0xc7 | 0xc9..=0xcb | 0xcd..=0xcf
    )
}

fn inspect_gif(bytes: &[u8]) -> UseResult<ImageMetadata> {
    if bytes.len() < 14 {
        return Err(image_invalid("GIF data is truncated."));
    }
    let width_px = u32::from(u16::from_le_bytes([bytes[6], bytes[7]]));
    let height_px = u32::from(u16::from_le_bytes([bytes[8], bytes[9]]));
    let packed = bytes[10];
    let table_bytes = if packed & 0x80 == 0 {
        0
    } else {
        3_usize << (usize::from(packed & 0x07) + 1)
    };
    let mut cursor = 13_usize
        .checked_add(table_bytes)
        .filter(|cursor| *cursor <= bytes.len())
        .ok_or_else(|| image_invalid("GIF global color table is truncated."))?;
    let mut saw_image = false;
    loop {
        match bytes.get(cursor).copied() {
            Some(0x2c) => {
                let descriptor = bytes
                    .get(cursor..cursor.saturating_add(10))
                    .ok_or_else(|| image_invalid("GIF image descriptor is truncated."))?;
                cursor += 10;
                if descriptor[9] & 0x80 != 0 {
                    let local_bytes = 3_usize << (usize::from(descriptor[9] & 0x07) + 1);
                    cursor = cursor
                        .checked_add(local_bytes)
                        .filter(|cursor| *cursor <= bytes.len())
                        .ok_or_else(|| image_invalid("GIF local color table is truncated."))?;
                }
                cursor = cursor
                    .checked_add(1)
                    .filter(|cursor| *cursor <= bytes.len())
                    .ok_or_else(|| image_invalid("GIF LZW code size is missing."))?;
                cursor = skip_gif_sub_blocks(bytes, cursor)?;
                saw_image = true;
            }
            Some(0x21) => {
                cursor = cursor
                    .checked_add(2)
                    .filter(|cursor| *cursor <= bytes.len())
                    .ok_or_else(|| image_invalid("GIF extension header is truncated."))?;
                cursor = skip_gif_sub_blocks(bytes, cursor)?;
            }
            Some(0x3b) if saw_image && cursor + 1 == bytes.len() => break,
            Some(0x3b) => {
                return Err(image_invalid(
                    "GIF trailer must follow image data and terminate the file.",
                ));
            }
            Some(_) => return Err(image_invalid("GIF data contains an unknown block type.")),
            None => return Err(image_invalid("GIF data has no terminating trailer.")),
        }
    }
    Ok(ImageMetadata {
        format: NativeOfficeImageFormat::Gif,
        width_px,
        height_px,
    })
}

fn skip_gif_sub_blocks(bytes: &[u8], mut cursor: usize) -> UseResult<usize> {
    loop {
        let length = usize::from(
            *bytes
                .get(cursor)
                .ok_or_else(|| image_invalid("GIF data sub-block is truncated."))?,
        );
        cursor += 1;
        if length == 0 {
            return Ok(cursor);
        }
        cursor = cursor
            .checked_add(length)
            .filter(|cursor| *cursor <= bytes.len())
            .ok_or_else(|| image_invalid("GIF data sub-block is truncated."))?;
    }
}
