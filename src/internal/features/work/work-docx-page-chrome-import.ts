import { normalizeDocumentImageIdentity } from './work-document-image-identity';
import { documentCharacterScaleDomAttributes } from './work-document-character-scale';
import { documentCharacterPositionDomAttributes } from './work-document-character-position';
import { documentCharacterSpacingDomAttributes } from './work-document-character-spacing';
import { docxCharacterScalePercentFromProperties } from './work-docx-character-scale';
import { docxCharacterPositionHalfPointsFromProperties } from './work-docx-character-position';
import { docxCharacterSpacingTwipsFromProperties } from './work-docx-character-spacing';
import { documentParagraphBordersDomAttributes } from './work-document-paragraph-borders';
import {
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
} from './work-document-page-chrome';
import { normalizeDocumentParagraphIdentity } from './work-document-paragraph-identity';
import { documentParagraphShadingDomAttributes } from './work-document-paragraph-shading';
import {
  DOCUMENT_TABLE_ROW_ID_ATTRIBUTE,
  DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
  normalizeDocumentTableRowIdentity,
} from './work-document-table-row-identity';
import { docxEquationHtml } from './work-docx-equation-import';
import { isDocxEquationLikeRoot } from './work-docx-equation-story';
import { parseDocxParagraphDefaultCollapsed } from './work-docx-paragraph-default-collapsed';
import {
  type DocxParagraphStyleResolver,
  createDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import { resolveDocxParagraphBordersForParagraph } from './work-docx-paragraph-borders-import';
import { parseDirectDocxParagraphShading } from './work-docx-paragraph-shading-import';
import {
  type DocxTableStyleResolver,
  createDocxTableStyleResolver,
} from './work-docx-table-styles';
import { createDocxThemeResolver } from './work-docx-theme';
import { documentUnderlineDomAttributes } from './work-document-underline';
import { importedDocxUnderline } from './work-docx-underline';
import { documentStrikeDomAttributes } from './work-document-strike';
import { importedDocxStrike } from './work-docx-strike';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  attribute,
  bytesToDataUrl,
  contentTypeForPart,
  descendants,
  directChild,
  directChildren,
  firstDescendant,
  type OoxmlPackage,
  type OoxmlRelationship,
} from './work-ooxml-package';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
  WorkDocumentSectionLayout,
} from './work-types';

type Relationships = Map<string, OoxmlRelationship>;
interface ImportedPageChromePart {
  html: string;
  showPageNumber: boolean;
}

const WORDPROCESSING_DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
]);
const WORDPROCESSING_DRAWING_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';

export async function importSectionPageChrome(
  section: Element,
  archive: OoxmlPackage,
  documentRelationships: Relationships,
  previous: WorkDocumentSectionLayout,
  oddEvenEnabled: boolean,
): Promise<{
  pageChrome: WorkDocumentPageChrome;
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
}> {
  const fallback = normalizeDocumentPageChrome(previous.pageChrome, previous);
  const variants = await Promise.all(
    (['default', 'first', 'even'] as const).map(async (variant) => {
      const content = await importVariant(
        section,
        variant,
        archive,
        documentRelationships,
        fallback[variant],
      );
      return [variant, content] as const;
    }),
  );
  const hasFirstReferences = hasVariantReference(section, 'first');
  const hasEvenReferences = hasVariantReference(section, 'even');
  const pageChrome = normalizeDocumentPageChrome({
    differentFirstPage:
      enabledElement(firstDescendant(section, 'titlePg')) || hasFirstReferences,
    differentOddEvenPages: oddEvenEnabled || hasEvenReferences,
    ...Object.fromEntries(variants),
  });
  return { pageChrome, ...documentPageChromeLegacyFields(pageChrome) };
}

export function documentUsesOddEvenPageChrome(
  settings: Document | null,
): boolean {
  return Boolean(
    settings && descendants(settings, 'evenAndOddHeaders').some(enabledElement),
  );
}

async function importVariant(
  section: Element,
  variant: WorkDocumentPageChromeVariant,
  archive: OoxmlPackage,
  documentRelationships: Relationships,
  fallback: WorkDocumentPageChromeContent,
): Promise<WorkDocumentPageChromeContent> {
  const header = await importReferencedPart(
    section,
    'headerReference',
    variant,
    archive,
    documentRelationships,
  );
  const footer = await importReferencedPart(
    section,
    'footerReference',
    variant,
    archive,
    documentRelationships,
  );
  const hasImportedPart = Boolean(header || footer);
  return {
    headerHtml: header?.html ?? fallback.headerHtml,
    footerHtml: footer?.html ?? fallback.footerHtml,
    showPageNumber: hasImportedPart
      ? Boolean(header?.showPageNumber || footer?.showPageNumber)
      : fallback.showPageNumber,
  };
}

async function importReferencedPart(
  section: Element,
  referenceName: 'headerReference' | 'footerReference',
  variant: WorkDocumentPageChromeVariant,
  archive: OoxmlPackage,
  documentRelationships: Relationships,
): Promise<ImportedPageChromePart | null> {
  const reference = directChildren(section, referenceName).find(
    (element) => (attribute(element, 'type') ?? 'default') === variant,
  );
  const relationship = documentRelationships.get(
    attribute(reference ?? section, 'r:id') ?? '',
  );
  if (
    !relationship ||
    relationship.targetMode === 'External' ||
    !archive.has(relationship.target)
  )
    return null;
  const document = await archive.xml(relationship.target);
  return {
    html: await pageChromePartHtml(document, archive, relationship.target),
    showPageNumber: containsPageNumber(document),
  };
}

async function pageChromePartHtml(
  document: Document,
  archive: OoxmlPackage,
  partPath: string,
): Promise<string> {
  const relationships = await archive.relationships(partPath);
  const themePath = archive
    .paths('word/theme/')
    .find((path) => /\/theme\d*\.xml$/i.test(path));
  const themeDocument = themePath ? await archive.xml(themePath) : null;
  const settingsDocument = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const stylesDocument = archive.has('word/styles.xml')
    ? await archive.xml('word/styles.xml')
    : null;
  const theme = createDocxThemeResolver(themeDocument, settingsDocument);
  const paragraphStyles = createDocxParagraphStyleResolver(stylesDocument);
  const tableStyles = createDocxTableStyleResolver(stylesDocument);
  const root = document.documentElement;
  const blocks: string[] = [];
  for (const child of directChildren(root)) {
    if (child.localName === 'p')
      blocks.push(
        await paragraphHtml(
          child,
          archive,
          relationships,
          theme,
          paragraphStyles,
          tableStyles,
        ),
      );
    if (child.localName === 'tbl')
      blocks.push(
        await tableHtml(
          child,
          archive,
          relationships,
          theme,
          paragraphStyles,
          tableStyles,
        ),
      );
    if (isDocxEquationLikeRoot(child))
      blocks.push(`<p>${docxEquationHtml(child)}</p>`);
  }
  return blocks.join('');
}

async function paragraphHtml(
  paragraph: Element,
  archive: OoxmlPackage,
  relationships: Relationships,
  theme?: ReturnType<typeof createDocxThemeResolver>,
  paragraphStyles?: DocxParagraphStyleResolver,
  tableStyles?: DocxTableStyleResolver,
): Promise<string> {
  const properties = directChild(paragraph, 'pPr');
  const alignmentValue = attribute(
    firstDescendant(properties, 'jc') ?? paragraph,
    'val',
  );
  const alignment =
    alignmentValue === 'center'
      ? 'center'
      : alignmentValue === 'right' || alignmentValue === 'end'
        ? 'right'
        : alignmentValue === 'both' || alignmentValue === 'distribute'
          ? 'justify'
          : alignmentValue === 'left' || alignmentValue === 'start'
            ? 'left'
            : '';
  const field: FieldState = {
    active: false,
    instruction: '',
    separated: false,
    skipResult: false,
  };
  let html = '';
  for (const child of directChildren(paragraph)) {
    if (child.localName === 'pPr') continue;
    if (isDocxEquationLikeRoot(child)) {
      html += docxEquationHtml(child);
      continue;
    }
    if (child.localName === 'hyperlink') {
      const content = await containerRunsHtml(
        child,
        field,
        archive,
        relationships,
        theme,
      );
      const relationship = relationships.get(attribute(child, 'r:id') ?? '');
      const anchor = attribute(child, 'anchor');
      const href =
        relationship?.targetMode === 'External'
          ? relationship.target
          : anchor
            ? `#${anchor}`
            : '';
      html += href ? `<a href="${escapeHtml(href)}">${content}</a>` : content;
      continue;
    }
    if (child.localName === 'fldSimple') {
      const instruction = attribute(child, 'instr') ?? '';
      if (!/\bPAGE\b/i.test(instruction)) {
        html += await containerRunsHtml(
          child,
          field,
          archive,
          relationships,
          theme,
        );
      }
      continue;
    }
    html += await containerRunsHtml(
      child,
      field,
      archive,
      relationships,
      theme,
    );
  }
  const identityAttributes = paragraphIdentityAttributes(paragraph);
  const defaultCollapsed = properties
    ? parseDocxParagraphDefaultCollapsed(properties)
    : { status: 'absent' as const };
  const defaultCollapsedAttribute =
    defaultCollapsed.status === 'valid'
      ? ` data-office-default-collapsed="${defaultCollapsed.value}"`
      : '';
  const parsedShading = properties
    ? parseDirectDocxParagraphShading(properties, theme)
    : undefined;
  const shading =
    parsedShading === null ? { pattern: 'nil' as const } : parsedShading;
  const shadingAttributes = documentParagraphShadingDomAttributes(shading);
  const borders = resolveDocxParagraphBordersForParagraph(
    paragraph,
    paragraphStyles,
    theme,
    tableStyles,
  ).borders;
  const borderAttributes = documentParagraphBordersDomAttributes(borders);
  const styles = [
    alignment ? `text-align: ${alignment}` : '',
    borderAttributes.style ?? '',
    shadingAttributes.style ?? '',
  ].filter(Boolean);
  const borderAttribute = borderAttributes['data-office-paragraph-borders'];
  const shadingAttribute = shadingAttributes['data-office-paragraph-shading'];
  return `<p${styles.length ? ` style="${escapeHtml(styles.join('; '))}"` : ''}${identityAttributes}${defaultCollapsedAttribute}${borderAttribute ? ` data-office-paragraph-borders="${escapeHtml(borderAttribute)}"` : ''}${shadingAttribute ? ` data-office-paragraph-shading="${escapeHtml(shadingAttribute)}"` : ''}>${html}</p>`;
}

function paragraphIdentityAttributes(paragraph: Element): string {
  const identity = normalizeDocumentParagraphIdentity({
    paragraphId: namespacedAttribute(paragraph, WORD_2010_NAMESPACE, 'paraId'),
    textId: namespacedAttribute(paragraph, WORD_2010_NAMESPACE, 'textId'),
  });
  return identity
    ? ` data-office-paragraph-id="${identity.paragraphId}" data-office-paragraph-text-id="${identity.textId}"`
    : '';
}

async function containerRunsHtml(
  container: Element,
  field: FieldState,
  archive: OoxmlPackage,
  relationships: Relationships,
  theme?: ReturnType<typeof createDocxThemeResolver>,
): Promise<string> {
  if (container.localName === 'r')
    return runHtml(container, field, archive, relationships, theme);
  let html = '';
  for (const child of directChildren(container)) {
    if (isDocxEquationLikeRoot(child)) {
      html += docxEquationHtml(child);
      continue;
    }
    html += await containerRunsHtml(
      child,
      field,
      archive,
      relationships,
      theme,
    );
  }
  return html;
}

async function runHtml(
  run: Element,
  field: FieldState,
  archive: OoxmlPackage,
  relationships: Relationships,
  theme?: ReturnType<typeof createDocxThemeResolver>,
): Promise<string> {
  let content = '';
  for (const child of directChildren(run)) {
    if (child.localName === 'fldChar') {
      updateFieldState(field, attribute(child, 'fldCharType'));
      continue;
    }
    if (child.localName === 'instrText') {
      if (field.active && !field.separated)
        field.instruction += child.textContent ?? '';
      continue;
    }
    if (!fieldContentVisible(field)) continue;
    if (isDocxEquationLikeRoot(child)) {
      content += docxEquationHtml(child);
      continue;
    }
    if (child.localName === 't') content += escapeHtml(child.textContent ?? '');
    if (child.localName === 'tab') content += '&#9;';
    if (child.localName === 'br' || child.localName === 'cr') content += '<br>';
    if (child.localName === 'drawing')
      content += await drawingHtml(child, archive, relationships);
  }
  if (!content) return '';

  const properties = directChild(run, 'rPr');
  if (enabledElement(directChild(properties ?? run, 'b')))
    content = `<strong>${content}</strong>`;
  if (enabledElement(directChild(properties ?? run, 'i')))
    content = `<em>${content}</em>`;
  const underline = directChild(properties ?? run, 'u');
  if (underline) {
    const attributes = documentUnderlineDomAttributes(
      importedDocxUnderline(underline, theme),
    );
    content = `<u${htmlAttributes(attributes)}>${content}</u>`;
  }
  const strike = importedDocxStrike(properties ?? run);
  if (strike) {
    content = `<s${htmlAttributes(
      documentStrikeDomAttributes(strike),
    )}>${content}</s>`;
  }
  const verticalAlign = attribute(
    directChild(properties ?? run, 'vertAlign') ?? run,
    'val',
  );
  if (verticalAlign === 'superscript') content = `<sup>${content}</sup>`;
  if (verticalAlign === 'subscript') content = `<sub>${content}</sub>`;
  const characterScale = docxCharacterScalePercentFromProperties(
    properties ?? run,
  );
  if (characterScale !== undefined) {
    content = `<span${htmlAttributes(
      documentCharacterScaleDomAttributes(characterScale),
    )}>${content}</span>`;
  }
  const characterPosition = docxCharacterPositionHalfPointsFromProperties(
    properties ?? run,
  );
  if (characterPosition !== undefined) {
    content = `<span${htmlAttributes(
      documentCharacterPositionDomAttributes(characterPosition),
    )}>${content}</span>`;
  }
  const characterSpacing = docxCharacterSpacingTwipsFromProperties(
    properties ?? run,
  );
  if (characterSpacing !== undefined) {
    content = `<span${htmlAttributes(
      documentCharacterSpacingDomAttributes(characterSpacing),
    )}>${content}</span>`;
  }
  const color = attribute(
    directChild(properties ?? run, 'color') ?? run,
    'val',
  );
  if (color && /^[0-9a-f]{6}$/i.test(color))
    content = `<span style="color: #${color}">${content}</span>`;
  return content;
}

async function drawingHtml(
  drawing: Element,
  archive: OoxmlPackage,
  relationships: Relationships,
): Promise<string> {
  const relationship = relationships.get(
    attribute(firstDescendant(drawing, 'blip') ?? drawing, 'embed') ?? '',
  );
  if (
    !relationship ||
    relationship.targetMode === 'External' ||
    !archive.has(relationship.target)
  )
    return '';
  const source = bytesToDataUrl(
    await archive.bytes(relationship.target),
    contentTypeForPart(relationship.target),
  );
  const extent = firstDescendant(drawing, 'extent');
  const width = Math.max(
    1,
    Math.round(Number(attribute(extent ?? drawing, 'cx')) / 9525),
  );
  const height = Math.max(
    1,
    Math.round(Number(attribute(extent ?? drawing, 'cy')) / 9525),
  );
  const identityAttributes = drawingIdentityAttributes(drawing);
  return `<img src="${source}" alt="Header or footer image"${width > 1 ? ` width="${width}"` : ''}${
    height > 1 ? ` height="${height}"` : ''
  }${identityAttributes}>`;
}

function drawingIdentityAttributes(drawing: Element): string {
  const anchor = Array.from(drawing.querySelectorAll('*')).find(
    (element) =>
      (element.localName === 'anchor' || element.localName === 'inline') &&
      WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const properties = anchor
    ? directChildren(anchor, 'docPr').find((element) =>
        WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
      )
    : undefined;
  const identity = normalizeDocumentImageIdentity({
    docPropertiesId: properties?.getAttribute('id'),
    anchorId: anchor
      ? namespacedAttribute(
          anchor,
          WORDPROCESSING_DRAWING_2010_NAMESPACE,
          'anchorId',
        )
      : null,
    editId: anchor
      ? namespacedAttribute(
          anchor,
          WORDPROCESSING_DRAWING_2010_NAMESPACE,
          'editId',
        )
      : null,
  });
  if (!identity) return '';
  return [
    ` data-office-image-object-id="${identity.objectId}"`,
    ` data-office-image-doc-properties-id="${identity.docPropertiesId}"`,
    ` data-office-image-anchor-id="${identity.anchorId}"`,
    ` data-office-image-edit-id="${identity.editId}"`,
  ].join('');
}

function namespacedAttribute(
  element: Element,
  namespace: string,
  localName: string,
): string | null {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === namespace
    ) {
      return item.value;
    }
  }
  return null;
}

async function tableHtml(
  table: Element,
  archive: OoxmlPackage,
  relationships: Relationships,
  theme?: ReturnType<typeof createDocxThemeResolver>,
  paragraphStyles?: DocxParagraphStyleResolver,
  tableStyles?: DocxTableStyleResolver,
): Promise<string> {
  const rows: string[] = [];
  for (const row of directChildren(table, 'tr')) {
    const cells: string[] = [];
    for (const cell of directChildren(row, 'tc')) {
      const paragraphs = await Promise.all(
        directChildren(cell, 'p').map((item) =>
          paragraphHtml(
            item,
            archive,
            relationships,
            theme,
            paragraphStyles,
            tableStyles,
          ),
        ),
      );
      cells.push(`<td>${paragraphs.join('')}</td>`);
    }
    rows.push(`<tr${tableRowIdentityAttributes(row)}>${cells.join('')}</tr>`);
  }
  return `<table><tbody>${rows.join('')}</tbody></table>`;
}

function tableRowIdentityAttributes(row: Element): string {
  const identity = normalizeDocumentTableRowIdentity({
    rowId: namespacedAttribute(row, WORD_2010_NAMESPACE, 'paraId'),
    rowTextId: namespacedAttribute(row, WORD_2010_NAMESPACE, 'textId'),
  });
  return identity
    ? ` ${DOCUMENT_TABLE_ROW_ID_ATTRIBUTE}="${identity.rowId}" ${DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE}="${identity.rowTextId}"`
    : '';
}

function containsPageNumber(document: Document): boolean {
  return (
    descendants(document, 'instrText').some((element) =>
      /\bPAGE\b/i.test(element.textContent ?? ''),
    ) ||
    descendants(document, 'fldSimple').some((element) =>
      /\bPAGE\b/i.test(attribute(element, 'instr') ?? ''),
    )
  );
}

function hasVariantReference(
  section: Element,
  variant: WorkDocumentPageChromeVariant,
): boolean {
  return [
    ...directChildren(section, 'headerReference'),
    ...directChildren(section, 'footerReference'),
  ].some((element) => (attribute(element, 'type') ?? 'default') === variant);
}

function enabledElement(element: Element | undefined): boolean {
  if (!element) return false;
  const value = attribute(element, 'val')?.toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function htmlAttributes(attributes: Record<string, string>): string {
  return Object.entries(attributes)
    .map(([name, value]) => ` ${name}="${escapeHtml(value)}"`)
    .join('');
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function updateFieldState(field: FieldState, fieldType: string | null): void {
  if (fieldType === 'begin') {
    field.active = true;
    field.instruction = '';
    field.separated = false;
    field.skipResult = false;
    return;
  }
  if (fieldType === 'separate' && field.active) {
    field.separated = true;
    field.skipResult = /\bPAGE\b/i.test(field.instruction);
    return;
  }
  if (fieldType === 'end') {
    field.active = false;
    field.instruction = '';
    field.separated = false;
    field.skipResult = false;
  }
}

function fieldContentVisible(field: FieldState): boolean {
  return !field.active || (field.separated && !field.skipResult);
}

interface FieldState {
  active: boolean;
  instruction: string;
  separated: boolean;
  skipResult: boolean;
}
