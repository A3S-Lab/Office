import {
  attribute,
  childPath,
  descendants,
  directChild,
} from './work-ooxml-package';

interface DocxStyle {
  basedOn?: string;
  paragraphProperties?: Element;
  runProperties?: Element;
}

export interface DocxParagraphStyleResolver {
  defaultProperties?: Element;
  defaultRunProperties?: Element;
  defaultStyleId?: string;
  styles: ReadonlyMap<string, DocxStyle>;
  characterStyles: ReadonlyMap<string, DocxStyle>;
}

export type DocxParagraphStyleSource =
  | Document
  | DocxParagraphStyleResolver
  | null
  | undefined;

const MAX_STYLE_INHERITANCE_DEPTH = 64;

export function createDocxParagraphStyleResolver(
  stylesDocument?: Document | null,
): DocxParagraphStyleResolver {
  if (!stylesDocument) return { styles: new Map(), characterStyles: new Map() };

  const styles = new Map<string, DocxStyle>();
  const characterStyles = new Map<string, DocxStyle>();
  let defaultStyleId: string | undefined;
  for (const style of descendants(stylesDocument, 'style')) {
    const type = wordAttribute(style, 'type');
    if (type !== 'paragraph' && type !== 'character') continue;
    const styleId = wordAttribute(style, 'styleId')?.trim();
    if (!styleId) continue;
    const basedOn = directChild(style, 'basedOn');
    const value: DocxStyle = {
      basedOn: basedOn
        ? wordAttribute(basedOn, 'val')?.trim() || undefined
        : undefined,
      paragraphProperties: directChild(style, 'pPr'),
      runProperties: directChild(style, 'rPr'),
    };
    if (type === 'paragraph') styles.set(styleId, value);
    else characterStyles.set(styleId, value);
    if (
      type === 'paragraph' &&
      !defaultStyleId &&
      onOffAttribute(style, 'default')
    ) {
      defaultStyleId = styleId;
    }
  }

  return {
    defaultProperties: childPath(
      stylesDocument.documentElement,
      'docDefaults',
      'pPrDefault',
      'pPr',
    ),
    defaultRunProperties: childPath(
      stylesDocument.documentElement,
      'docDefaults',
      'rPrDefault',
      'rPr',
    ),
    defaultStyleId,
    styles,
    characterStyles,
  };
}

export function resolveDocxParagraphStyleResolver(
  source: DocxParagraphStyleSource,
): DocxParagraphStyleResolver {
  return source && 'styles' in source
    ? source
    : createDocxParagraphStyleResolver(source);
}

export function docxParagraphPropertySources(
  directProperties: Element | undefined,
  resolver: DocxParagraphStyleResolver,
): Element[] {
  const sources: Element[] = [];
  if (resolver.defaultProperties) sources.push(resolver.defaultProperties);

  const styleReference = directProperties
    ? directChild(directProperties, 'pStyle')
    : undefined;
  const styleId =
    (styleReference
      ? wordAttribute(styleReference, 'val')?.trim()
      : undefined) || resolver.defaultStyleId;
  if (styleId)
    sources.push(
      ...stylePropertySources(resolver.styles, styleId, 'paragraphProperties'),
    );

  if (directProperties) sources.push(directProperties);
  return sources;
}

export function docxRunPropertySources(
  paragraphProperties: Element | undefined,
  runProperties: Element | undefined,
  resolver: DocxParagraphStyleResolver,
): Element[] {
  const sources: Element[] = [];
  if (resolver.defaultRunProperties)
    sources.push(resolver.defaultRunProperties);

  const paragraphStyleReference = paragraphProperties
    ? directChild(paragraphProperties, 'pStyle')
    : undefined;
  const paragraphStyleId =
    (paragraphStyleReference
      ? wordAttribute(paragraphStyleReference, 'val')?.trim()
      : undefined) || resolver.defaultStyleId;
  if (paragraphStyleId)
    sources.push(
      ...stylePropertySources(
        resolver.styles,
        paragraphStyleId,
        'runProperties',
      ),
    );

  const paragraphRunProperties = paragraphProperties
    ? directChild(paragraphProperties, 'rPr')
    : undefined;
  if (paragraphRunProperties) sources.push(paragraphRunProperties);

  const characterStyleReference = runProperties
    ? directChild(runProperties, 'rStyle')
    : undefined;
  const characterStyleId = characterStyleReference
    ? wordAttribute(characterStyleReference, 'val')?.trim()
    : undefined;
  if (characterStyleId)
    sources.push(
      ...stylePropertySources(
        resolver.characterStyles,
        characterStyleId,
        'runProperties',
      ),
    );

  if (runProperties) sources.push(runProperties);
  return sources;
}

function stylePropertySources(
  styles: ReadonlyMap<string, DocxStyle>,
  styleId: string,
  property: 'paragraphProperties' | 'runProperties',
): Element[] {
  const inherited: Element[] = [];
  const visited = new Set<string>();
  let currentStyleId: string | undefined = styleId;
  let depth = 0;
  while (
    currentStyleId &&
    depth < MAX_STYLE_INHERITANCE_DEPTH &&
    !visited.has(currentStyleId)
  ) {
    depth += 1;
    visited.add(currentStyleId);
    const style = styles.get(currentStyleId);
    if (!style) break;
    const properties = style[property];
    if (properties) inherited.push(properties);
    currentStyleId = style.basedOn;
  }
  return inherited.reverse();
}

function onOffAttribute(element: Element, name: string): boolean {
  const value = wordAttribute(element, name)?.trim().toLowerCase();
  if (value === null || value === undefined) return false;
  return value !== '0' && value !== 'false' && value !== 'off';
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}
