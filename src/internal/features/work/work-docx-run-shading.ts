import type { DocumentRunShading } from './work-document-run-shading';
import { parseDocxParagraphShadingElement } from './work-docx-paragraph-shading-import';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  type DocxThemeResolver,
  type DocxThemeSource,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import { directChildren } from './work-ooxml-package';

export type DocxRunShadingInspection =
  | { status: 'absent'; spoofedCount: number }
  | { status: 'invalid'; spoofedCount: number }
  | {
      status: 'valid';
      value: DocumentRunShading;
      spoofedCount: number;
    };

export interface ResolvedDocxRunShading {
  shading: DocumentRunShading | undefined;
  invalidCount: number;
  spoofedCount: number;
}

export function inspectDocxRunShading(
  properties: Element | null | undefined,
  themeSource?: DocxThemeSource,
): DocxRunShadingInspection {
  return inspectDocxRunShadingWithTheme(
    properties,
    resolveDocxThemeResolver(themeSource),
  );
}

export function resolveDocxRunShading(
  propertySources: readonly Element[],
  themeSource?: DocxThemeSource,
): ResolvedDocxRunShading {
  const theme = resolveDocxThemeResolver(themeSource);
  let shading: DocumentRunShading | undefined;
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const inspection = inspectDocxRunShadingWithTheme(properties, theme);
    spoofedCount += inspection.spoofedCount;
    if (inspection.status === 'valid') shading = inspection.value;
    else if (inspection.status === 'invalid') {
      invalidCount += 1;
      shading = { pattern: 'nil' };
    }
  }
  return { shading, invalidCount, spoofedCount };
}

function inspectDocxRunShadingWithTheme(
  properties: Element | null | undefined,
  theme: DocxThemeResolver,
): DocxRunShadingInspection {
  if (!properties) return { status: 'absent', spoofedCount: 0 };
  const matches = directChildren(properties, 'shd');
  if (!matches.length) return { status: 'absent', spoofedCount: 0 };
  const native = matches.filter(
    (element) =>
      element.namespaceURI === properties.namespaceURI &&
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const spoofedCount = matches.length - native.length;
  if (matches.length !== 1 || native.length !== 1) {
    return { status: 'invalid', spoofedCount };
  }
  const parsed = parseDocxParagraphShadingElement(native[0], theme);
  return parsed
    ? { status: 'valid', value: parsed, spoofedCount }
    : { status: 'invalid', spoofedCount };
}
