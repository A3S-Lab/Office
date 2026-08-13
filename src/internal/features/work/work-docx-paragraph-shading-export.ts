import type { IParagraphPropertiesOptions } from 'docx';
import {
  type DocumentParagraphShadingColor,
  parseDocumentParagraphShadingElement,
} from './work-document-paragraph-shading';
import type { DocxThemePatchCollector } from './work-docx-theme-reference';

type ParagraphShadingOptions = Partial<
  Pick<IParagraphPropertiesOptions, 'shading'>
>;

export function documentParagraphShadingDocxOptions(
  element: HTMLElement,
  themePatches?: DocxThemePatchCollector,
): ParagraphShadingOptions {
  const shading = parseDocumentParagraphShadingElement(element);
  if (!shading) return {};
  const color = docxShadingColor(shading.color, 'shadingColor', themePatches);
  const fill = docxShadingColor(shading.fill, 'fill', themePatches);
  return {
    shading: {
      type: shading.pattern as NonNullable<
        IParagraphPropertiesOptions['shading']
      >['type'],
      ...(color ? { color } : {}),
      ...(fill ? { fill } : {}),
    },
  };
}

function docxShadingColor(
  color: DocumentParagraphShadingColor | undefined,
  kind: 'fill' | 'shadingColor',
  themePatches: DocxThemePatchCollector | undefined,
): string | undefined {
  if (!color) return undefined;
  if (color.value === 'auto') return 'auto';
  return (
    themePatches?.marker(kind, color.theme ?? null, color.value) ??
    color.value.slice(1).toUpperCase()
  );
}
