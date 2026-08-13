import type {
  IIndentAttributesProperties,
  INumberingOptions,
  ParagraphChild,
} from 'docx';
import {
  paragraphAlignment,
  paragraphDirectionOptions,
  paragraphIndent,
  paragraphPaginationOptions,
  paragraphSpacingOptions,
  paragraphTabStops,
} from './work-docx-export-formatting';
import type { DocxSourceNumberingIdentity } from './work-docx-numbering-extension-preservation';
import { documentParagraphShadingDocxOptions } from './work-docx-paragraph-shading-export';
import type { DocxThemePatchCollector } from './work-docx-theme-reference';

export interface DocxListExportContext {
  numbering: Array<INumberingOptions['config'][number]>;
  nextNumberingReference: number;
  numberingReferences: Map<string, string>;
  numberingLevels: Map<
    string,
    Array<INumberingOptions['config'][number]['levels'][number]>
  >;
  numberingRestartRules: Array<Map<number, number>>;
  numberingRestartRulesByIdentity: Map<string, Map<number, number>>;
  numberingSourceIdentities: Array<DocxSourceNumberingIdentity | null>;
  themePatches?: DocxThemePatchCollector;
}

type InlineRuns = (element: HTMLElement) => Promise<ParagraphChild[]>;

export async function listToDocxParagraphs(
  list: HTMLElement,
  docx: typeof import('docx'),
  context: DocxListExportContext,
  inlineRuns: InlineRuns,
  depth = 0,
): Promise<Array<InstanceType<typeof docx.Paragraph>>> {
  const paragraphs: Array<InstanceType<typeof docx.Paragraph>> = [];
  const level = listNumberingLevel(list, depth);
  const ordered = list.tagName.toLowerCase() === 'ol';
  const numbering = ordered
    ? registerOrderedListNumbering(list, level, docx, context)
    : registerBulletListNumbering(list, level, docx, context);
  const items = Array.from(list.children).filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement && child.tagName.toLowerCase() === 'li',
  );
  for (const item of items) {
    const directBlocks = Array.from(item.children).filter(
      (child): child is HTMLElement =>
        child instanceof HTMLElement &&
        child.tagName.toLowerCase() !== 'ol' &&
        child.tagName.toLowerCase() !== 'ul',
    );
    const contentBlocks = directBlocks.length
      ? directBlocks
      : [directListItemContentRoot(item)];
    for (const [blockIndex, block] of contentBlocks.entries()) {
      const tag = block.tagName.toLowerCase();
      const heading =
        tag === 'h1'
          ? docx.HeadingLevel.HEADING_1
          : tag === 'h2'
            ? docx.HeadingLevel.HEADING_2
            : tag === 'h3'
              ? docx.HeadingLevel.HEADING_3
              : undefined;
      const runs = await inlineRuns(block);
      paragraphs.push(
        new docx.Paragraph({
          children: runs.length ? runs : [new docx.TextRun('')],
          heading,
          ...(blockIndex === 0
            ? { numbering }
            : {
                indent: paragraphIndent(block, tag) ?? {
                  left: (level + 1) * 720,
                },
              }),
          alignment: paragraphAlignment(block, docx),
          spacing: paragraphSpacingOptions(block, Boolean(heading), docx),
          ...(blockIndex === 0 ? { indent: paragraphIndent(block, tag) } : {}),
          tabStops: paragraphTabStops(block, docx),
          ...paragraphPaginationOptions(block),
          ...paragraphDirectionOptions(block),
          ...documentParagraphShadingDocxOptions(block, context.themePatches),
        }),
      );
    }
    for (const nested of Array.from(item.children).filter(
      (child): child is HTMLElement =>
        child instanceof HTMLElement &&
        (child.tagName.toLowerCase() === 'ol' ||
          child.tagName.toLowerCase() === 'ul'),
    )) {
      paragraphs.push(
        ...(await listToDocxParagraphs(
          nested,
          docx,
          context,
          inlineRuns,
          depth + 1,
        )),
      );
    }
  }
  return paragraphs;
}

function registerOrderedListNumbering(
  list: HTMLElement,
  level: number,
  docx: typeof import('docx'),
  context: DocxListExportContext,
): { reference: string; level: number } {
  const sourceIdentity = importedListSourceIdentity(list);
  const identity = sourceIdentity
    ? sourceNumberingIdentityKey(sourceIdentity)
    : null;
  const existingReference = identity
    ? context.numberingReferences.get(identity)
    : undefined;
  if (existingReference && identity) {
    const levels = context.numberingLevels.get(identity);
    if (levels) levels[level] = orderedListLevelOptions(list, level, docx);
    updateNumberingRestartRule(context, identity, list, level);
    return { reference: existingReference, level };
  }
  const reference = `a3s-office-list-${context.nextNumberingReference}`;
  context.nextNumberingReference += 1;
  const levels = Array.from({ length: 9 }, (_, currentLevel) =>
    currentLevel === level
      ? orderedListLevelOptions(list, currentLevel, docx)
      : defaultOrderedListLevelOptions(currentLevel, docx),
  );
  const configuration = { reference, levels };
  const restartRules = new Map<number, number>();
  setNumberingRestartRule(restartRules, list, level);
  context.numberingRestartRules.push(restartRules);
  context.numberingSourceIdentities.push(sourceIdentity);
  if (identity) {
    context.numberingReferences.set(identity, reference);
    context.numberingLevels.set(identity, levels);
    context.numberingRestartRulesByIdentity.set(identity, restartRules);
  }
  context.numbering.push(configuration);
  return { reference, level };
}

function registerBulletListNumbering(
  list: HTMLElement,
  level: number,
  docx: typeof import('docx'),
  context: DocxListExportContext,
): { reference: string; level: number } {
  const sourceIdentity = importedListSourceIdentity(list);
  const identity = sourceIdentity
    ? sourceNumberingIdentityKey(sourceIdentity)
    : null;
  const existingReference = identity
    ? context.numberingReferences.get(identity)
    : undefined;
  if (existingReference && identity) {
    const levels = context.numberingLevels.get(identity);
    if (levels) levels[level] = bulletListLevelOptions(list, level, docx);
    updateNumberingRestartRule(context, identity, list, level);
    return { reference: existingReference, level };
  }
  const reference = `a3s-office-list-${context.nextNumberingReference}`;
  context.nextNumberingReference += 1;
  const levels = Array.from({ length: 9 }, (_, currentLevel) =>
    currentLevel === level
      ? bulletListLevelOptions(list, currentLevel, docx)
      : defaultBulletListLevelOptions(currentLevel, docx),
  );
  const configuration = { reference, levels };
  const restartRules = new Map<number, number>();
  setNumberingRestartRule(restartRules, list, level);
  context.numberingRestartRules.push(restartRules);
  context.numberingSourceIdentities.push(sourceIdentity);
  if (identity) {
    context.numberingReferences.set(identity, reference);
    context.numberingLevels.set(identity, levels);
    context.numberingRestartRulesByIdentity.set(identity, restartRules);
  }
  context.numbering.push(configuration);
  return { reference, level };
}

function orderedListLevelFormat(
  list: HTMLElement,
  level: number,
  docx: typeof import('docx'),
) {
  const imported = importedNumberingFormat(list, docx);
  if (imported && imported !== docx.LevelFormat.BULLET) return imported;
  const type = list.getAttribute('type');
  if (type === 'A') return docx.LevelFormat.UPPER_LETTER;
  if (type === 'a') return docx.LevelFormat.LOWER_LETTER;
  if (type === 'I') return docx.LevelFormat.UPPER_ROMAN;
  if (type === 'i') return docx.LevelFormat.LOWER_ROMAN;
  return defaultOrderedListLevelFormat(level, docx);
}

function defaultOrderedListLevelFormat(
  level: number,
  docx: typeof import('docx'),
) {
  if (level % 3 === 1) return docx.LevelFormat.LOWER_LETTER;
  if (level % 3 === 2) return docx.LevelFormat.LOWER_ROMAN;
  return docx.LevelFormat.DECIMAL;
}

function orderedListLevelOptions(
  list: HTMLElement,
  level: number,
  docx: typeof import('docx'),
) {
  return applyImportedListLevelLayout(
    list,
    {
      ...defaultOrderedListLevelOptions(level, docx),
      format: orderedListLevelFormat(list, level, docx),
      text: importedNumberingText(list) ?? `%${level + 1}.`,
      start: orderedListStart(list),
    },
    docx,
  );
}

function defaultOrderedListLevelOptions(
  level: number,
  docx: typeof import('docx'),
) {
  return listLevelOptions(
    level,
    defaultOrderedListLevelFormat(level, docx),
    `%${level + 1}.`,
    1,
    docx,
  );
}

function bulletListLevelOptions(
  list: HTMLElement,
  level: number,
  docx: typeof import('docx'),
) {
  return applyImportedListLevelLayout(
    list,
    listLevelOptions(
      level,
      importedNumberingFormat(list, docx) ?? docx.LevelFormat.BULLET,
      importedNumberingText(list) ?? bulletListMarker(list),
      1,
      docx,
    ),
    docx,
  );
}

function defaultBulletListLevelOptions(
  level: number,
  docx: typeof import('docx'),
) {
  return listLevelOptions(
    level,
    docx.LevelFormat.BULLET,
    defaultBulletListMarker(level),
    1,
    docx,
  );
}

function listLevelOptions(
  level: number,
  format: typeof import('docx')['LevelFormat'][keyof typeof import('docx')['LevelFormat']],
  text: string,
  start: number,
  docx: typeof import('docx'),
) {
  return {
    level,
    format,
    text,
    start,
    suffix: docx.LevelSuffix.TAB,
    style: {
      paragraph: {
        indent: { left: (level + 1) * 720, hanging: 360 },
      },
    },
  };
}

function orderedListStart(list: HTMLElement): number {
  const value = Number(list.getAttribute('start'));
  return Number.isSafeInteger(value) && value > 0
    ? Math.min(value, 2_147_483_647)
    : 1;
}

function bulletListMarker(list: HTMLElement): string {
  const style =
    list.dataset.officeBulletStyle || list.style.listStyleType || 'disc';
  if (style === 'circle') return '\u25cb';
  if (style === 'square') return '\u25a0';
  return '\u25cf';
}

function defaultBulletListMarker(level: number): string {
  if (level % 3 === 1) return '\u25cb';
  if (level % 3 === 2) return '\u25a0';
  return '\u25cf';
}

function importedListSourceIdentity(
  list: HTMLElement,
): DocxSourceNumberingIdentity | null {
  const numId = importedNumberingId(list.dataset.officeNumberingId);
  if (numId === null) return null;
  const abstractNumId = importedNumberingId(
    list.dataset.officeAbstractNumberingId,
  );
  return {
    numId,
    ...(abstractNumId === null ? {} : { abstractNumId }),
  };
}

function sourceNumberingIdentityKey(
  identity: DocxSourceNumberingIdentity,
): string {
  return `${identity.numId}:${identity.abstractNumId ?? ''}`;
}

function importedNumberingId(value: string | undefined): number | null {
  if (!value || !/^\d{1,10}$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= 2_147_483_647
    ? parsed
    : null;
}

function listNumberingLevel(list: HTMLElement, fallback: number): number {
  const imported = Number(list.dataset.officeNumberingLevel);
  return Number.isSafeInteger(imported) && imported >= 0 && imported <= 8
    ? imported
    : Math.min(8, Math.max(0, fallback));
}

function importedNumberingFormat(
  list: HTMLElement,
  docx: typeof import('docx'),
): (typeof docx.LevelFormat)[keyof typeof docx.LevelFormat] | null {
  const imported = list.dataset.officeNumberingFormat;
  if (!imported) return null;
  const supported = new Set<string>(Object.values(docx.LevelFormat));
  return supported.has(imported)
    ? (imported as (typeof docx.LevelFormat)[keyof typeof docx.LevelFormat])
    : null;
}

function importedNumberingText(list: HTMLElement): string | null {
  const imported = list.dataset.officeNumberingText;
  return imported && imported.length <= 255 ? imported : null;
}

function applyImportedListLevelLayout<
  T extends ReturnType<typeof listLevelOptions>,
>(list: HTMLElement, options: T, docx: typeof import('docx')): T {
  const suffix = list.dataset.officeNumberingSuffix;
  const supportedSuffixes = new Set<string>(Object.values(docx.LevelSuffix));
  const alignment = list.dataset.officeNumberingAlignment;
  const supportedAlignments = new Set<string>(
    Object.values(docx.AlignmentType),
  );
  const left = importedTwips(list.dataset.officeNumberingIndentLeft, true);
  const right = importedTwips(list.dataset.officeNumberingIndentRight, true);
  const start = importedTwips(list.dataset.officeNumberingIndentStart, true);
  const end = importedTwips(list.dataset.officeNumberingIndentEnd, true);
  const hanging = importedTwips(
    list.dataset.officeNumberingIndentHanging,
    false,
  );
  const firstLine = importedTwips(
    list.dataset.officeNumberingIndentFirstLine,
    false,
  );
  const importedIndentMode = hanging !== null || firstLine !== null;
  const importedPositionMode =
    left !== null || right !== null || start !== null || end !== null;
  const sourceIndent: IIndentAttributesProperties =
    options.style.paragraph.indent;
  const baseIndent: Record<string, unknown> = { ...sourceIndent };
  if (importedPositionMode) {
    delete baseIndent.left;
    delete baseIndent.right;
    delete baseIndent.start;
    delete baseIndent.end;
  }
  if (importedIndentMode) {
    delete baseIndent.hanging;
    delete baseIndent.firstLine;
  }
  return {
    ...options,
    ...(suffix && supportedSuffixes.has(suffix)
      ? {
          suffix:
            suffix as (typeof docx.LevelSuffix)[keyof typeof docx.LevelSuffix],
        }
      : {}),
    ...(alignment && supportedAlignments.has(alignment)
      ? {
          alignment:
            alignment as (typeof docx.AlignmentType)[keyof typeof docx.AlignmentType],
        }
      : {}),
    style: {
      ...options.style,
      paragraph: {
        ...options.style.paragraph,
        indent: {
          ...baseIndent,
          ...(left === null ? {} : { left }),
          ...(right === null ? {} : { right }),
          ...(start === null ? {} : { start }),
          ...(end === null ? {} : { end }),
          ...(hanging === null ? {} : { hanging }),
          ...(firstLine === null ? {} : { firstLine }),
        },
      },
    },
  };
}

function importedTwips(
  value: string | undefined,
  signed: boolean,
): number | null {
  if (value === undefined || value === '') return null;
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) return null;
  if (!signed && parsed < 0) return null;
  return parsed;
}

function updateNumberingRestartRule(
  context: DocxListExportContext,
  identity: string,
  list: HTMLElement,
  level: number,
): void {
  const rules = context.numberingRestartRulesByIdentity.get(identity);
  if (rules) setNumberingRestartRule(rules, list, level);
}

function setNumberingRestartRule(
  rules: Map<number, number>,
  list: HTMLElement,
  level: number,
): void {
  const value = list.dataset.officeNumberingRestartAfterLevel;
  if (value === undefined || value === '') return;
  const parsed = Number(value);
  if (Number.isSafeInteger(parsed) && parsed >= 0 && parsed <= 9) {
    rules.set(level, parsed);
  }
}

function directListItemContentRoot(item: HTMLElement): HTMLElement {
  const root = item.ownerDocument.createElement('span');
  const direction = item.getAttribute('dir');
  if (direction) root.setAttribute('dir', direction);
  for (const child of item.childNodes) {
    if (
      child instanceof HTMLElement &&
      (child.tagName.toLowerCase() === 'ol' ||
        child.tagName.toLowerCase() === 'ul')
    ) {
      continue;
    }
    root.append(child.cloneNode(true));
  }
  return root;
}
