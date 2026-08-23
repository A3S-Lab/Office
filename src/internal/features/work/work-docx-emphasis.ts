import type { IRunOptions } from 'docx';
import {
  normalizeDocumentEmphasisMark,
  type WorkDocumentEmphasisMark,
} from './work-document-emphasis';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';

export type DocxEmphasisMarkInspection =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: WorkDocumentEmphasisMark };

export function docxEmphasisMarkFromProperties(
  properties: Element | null | undefined,
): WorkDocumentEmphasisMark | undefined {
  const inspection = inspectDocxEmphasisMark(properties);
  return inspection.status === 'valid' ? inspection.value : undefined;
}

export function resolveDocxEmphasisMark(
  propertySources: readonly Element[],
): WorkDocumentEmphasisMark | undefined {
  let emphasisMark: WorkDocumentEmphasisMark | undefined;
  for (const properties of propertySources) {
    const candidate = docxEmphasisMarkFromProperties(properties);
    if (candidate !== undefined) emphasisMark = candidate;
  }
  return emphasisMark;
}

export function inspectDocxEmphasisMark(
  properties: Element | null | undefined,
): DocxEmphasisMarkInspection {
  if (!properties) return { status: 'absent' };
  const localMatches = directChildren(properties, 'em');
  if (!localMatches.length) return { status: 'absent' };
  const nativeMatches = localMatches.filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (localMatches.length !== 1 || nativeMatches.length !== 1) {
    return { status: 'invalid' };
  }
  const emphasis = nativeMatches[0];
  if (
    directChildren(emphasis).length ||
    Array.from(emphasis.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return { status: 'invalid' };
  }
  const attributes = Array.from(emphasis.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(emphasis, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(emphasis, attributes[0]) ?? '',
    )
  ) {
    return { status: 'invalid' };
  }
  const value = normalizeDocumentEmphasisMark(attributes[0].value);
  return value === null ? { status: 'invalid' } : { status: 'valid', value };
}

export function docxEmphasisMarkRunOptions(
  value: unknown,
): IRunOptions['emphasisMark'] | undefined {
  const emphasisMark = normalizeDocumentEmphasisMark(value);
  if (!emphasisMark) return undefined;
  // docx 9.5 emits every ST_Em value but its declaration currently narrows
  // the public enum to `dot`. Keep the cast at this boundary so the rest of
  // the editor remains closed over the five schema-defined values.
  return {
    type: emphasisMark as NonNullable<IRunOptions['emphasisMark']>['type'],
  };
}
