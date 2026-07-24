export const OFFICE_DOCUMENT_LAYOUT_FONT_ID = 'noto-sans-hans-regular';
export const OFFICE_DOCUMENT_LAYOUT_FONT_FAMILY = 'A3S Office Noto Sans Hans';
export const OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID = 'noto-sans-regular';
export const OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_FAMILY = 'A3S Office Noto Sans';
export const OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID =
  'noto-naskh-arabic-regular';
export const OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_FAMILY =
  'A3S Office Noto Naskh Arabic';
export const OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID = 'noto-sans-hebrew-regular';
export const OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_FAMILY =
  'A3S Office Noto Sans Hebrew';

export interface WorkDocumentLayoutFont {
  id: string;
  family: string;
  url: string;
  weight?: number;
  style?: 'normal' | 'italic';
}

export function documentLayoutFontKey(
  fonts: readonly WorkDocumentLayoutFont[],
): string {
  return fonts
    .map(
      (font) =>
        `${font.id}\u0000${font.family}\u0000${font.url}\u0000${
          font.weight ?? 400
        }\u0000${font.style ?? 'normal'}`,
    )
    .join('\u0001');
}
