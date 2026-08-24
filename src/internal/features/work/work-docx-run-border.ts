import type { DocumentRunBorder } from './work-document-run-border';
import { normalizeDocumentRunBorder } from './work-document-run-border';
import { parseDocxBorderElement } from './work-docx-paragraph-borders-import';
import {
  type DocxThemeResolver,
  type DocxThemeSource,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import { directChildren } from './work-ooxml-package';

export type DocxRunBorderInspection =
  | { status: 'absent'; spoofedCount: number }
  | { status: 'invalid'; spoofedCount: number }
  | {
      status: 'valid';
      value: DocumentRunBorder;
      spoofedCount: number;
    };

export interface ResolvedDocxRunBorder {
  border: DocumentRunBorder | undefined;
  invalidCount: number;
  spoofedCount: number;
}

export function inspectDocxRunBorder(
  properties: Element | null | undefined,
  themeSource?: DocxThemeSource,
): DocxRunBorderInspection {
  return inspectDocxRunBorderWithTheme(
    properties,
    resolveDocxThemeResolver(themeSource),
  );
}

export function resolveDocxRunBorder(
  propertySources: readonly Element[],
  themeSource?: DocxThemeSource,
): ResolvedDocxRunBorder {
  const theme = resolveDocxThemeResolver(themeSource);
  let border: DocumentRunBorder | undefined;
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const inspection = inspectDocxRunBorderWithTheme(properties, theme);
    spoofedCount += inspection.spoofedCount;
    if (inspection.status === 'valid') border = inspection.value;
    else if (inspection.status === 'invalid') {
      invalidCount += 1;
      border = { style: 'nil' };
    }
  }
  return { border, invalidCount, spoofedCount };
}

function inspectDocxRunBorderWithTheme(
  properties: Element | null | undefined,
  theme: DocxThemeResolver,
): DocxRunBorderInspection {
  if (!properties) return { status: 'absent', spoofedCount: 0 };
  const matches = directChildren(properties, 'bdr');
  if (!matches.length) return { status: 'absent', spoofedCount: 0 };
  const native = matches.filter(
    (element) => element.namespaceURI === properties.namespaceURI,
  );
  const spoofedCount = matches.length - native.length;
  if (matches.length !== 1 || native.length !== 1) {
    return { status: 'invalid', spoofedCount };
  }
  const parsed = normalizeDocumentRunBorder(
    parseDocxBorderElement(native[0], theme),
  );
  return parsed
    ? { status: 'valid', value: parsed, spoofedCount }
    : { status: 'invalid', spoofedCount };
}
