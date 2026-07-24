use crate::KernelError;

pub(super) struct Utf16Offsets {
    offsets: Vec<(usize, u32)>,
}

impl Utf16Offsets {
    pub(super) fn new(text: &str) -> Self {
        let mut offsets = Vec::with_capacity(text.chars().count() + 1);
        let mut utf16 = 0_u32;
        for (byte, character) in text.char_indices() {
            offsets.push((byte, utf16));
            utf16 = utf16.saturating_add(character.len_utf16() as u32);
        }
        offsets.push((text.len(), utf16));
        Self { offsets }
    }

    pub(super) fn at(&self, byte: usize) -> Result<u32, KernelError> {
        self.offsets
            .binary_search_by_key(&byte, |(candidate, _)| *candidate)
            .ok()
            .map(|index| self.offsets[index].1)
            .ok_or_else(|| {
                KernelError::invalid(
                    "office.kernel.text_offset_invalid",
                    "Text layout produced an invalid character boundary.",
                )
            })
    }

    pub(super) fn byte_at(&self, utf16: u32) -> Result<usize, KernelError> {
        self.offsets
            .binary_search_by_key(&utf16, |(_, candidate)| *candidate)
            .ok()
            .map(|index| self.offsets[index].0)
            .ok_or_else(|| {
                KernelError::invalid(
                    "office.kernel.text_offset_invalid",
                    "A text layout run does not end on a Unicode character boundary.",
                )
            })
    }

    pub(super) fn len(&self) -> u32 {
        self.offsets.last().map_or(0, |(_, utf16)| *utf16)
    }
}
