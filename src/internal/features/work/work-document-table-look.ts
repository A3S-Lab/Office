import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute } from './work-ooxml-package';

export interface DocumentTableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHorizontalBand: boolean;
  noVerticalBand: boolean;
}

const LOOK_KEYS = [
  'firstRow',
  'lastRow',
  'firstColumn',
  'lastColumn',
  'noHorizontalBand',
  'noVerticalBand',
] as const;

const TABLE_LOOK_BITS = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHorizontalBand: 0x0200,
  noVerticalBand: 0x0400,
} as const;

const TBL_LOOK_ATTRIBUTES = new Set([
  'val',
  'firstRow',
  'lastRow',
  'firstColumn',
  'lastColumn',
  'noHBand',
  'noVBand',
]);

export function normalizeDocumentTableLook(
  value: unknown,
): DocumentTableLook | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some((key) => !(LOOK_KEYS as readonly string[]).includes(key))
  ) {
    return null;
  }
  const look: DocumentTableLook = {
    firstRow: false,
    lastRow: false,
    firstColumn: false,
    lastColumn: false,
    noHorizontalBand: false,
    noVerticalBand: false,
  };
  for (const key of LOOK_KEYS) {
    if (!(key in record)) continue;
    if (typeof record[key] !== 'boolean') return null;
    look[key] = record[key];
  }
  return orderedDocumentTableLook(look);
}

export function orderedDocumentTableLook(
  look: DocumentTableLook,
): DocumentTableLook {
  return {
    firstRow: look.firstRow,
    lastRow: look.lastRow,
    firstColumn: look.firstColumn,
    lastColumn: look.lastColumn,
    noHorizontalBand: look.noHorizontalBand,
    noVerticalBand: look.noVerticalBand,
  };
}

export function documentTableLookBitmask(look: DocumentTableLook): string {
  let value = 0;
  if (look.firstRow) value |= TABLE_LOOK_BITS.firstRow;
  if (look.lastRow) value |= TABLE_LOOK_BITS.lastRow;
  if (look.firstColumn) value |= TABLE_LOOK_BITS.firstColumn;
  if (look.lastColumn) value |= TABLE_LOOK_BITS.lastColumn;
  if (look.noHorizontalBand) value |= TABLE_LOOK_BITS.noHorizontalBand;
  if (look.noVerticalBand) value |= TABLE_LOOK_BITS.noVerticalBand;
  return value.toString(16).toUpperCase().padStart(4, '0');
}

export function serializeDocumentTableLookDataset(
  look: DocumentTableLook,
): string {
  return JSON.stringify(orderedDocumentTableLook(look));
}

export function parseDocumentTableLookDataset(
  value: string | undefined | null,
): DocumentTableLook | null {
  if (typeof value !== 'string' || !value.trim()) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  return normalizeDocumentTableLook(parsed);
}

export function parseDocxTblLookElement(
  element: Element,
): DocumentTableLook | null {
  if (element.localName !== 'tblLook' || element.children.length > 0) {
    return null;
  }
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) => xmlAttributeNamespace(element, candidate) === namespace,
  );
  if (!attributes.length) return null;
  if (
    attributes.some(
      (candidate) =>
        !TBL_LOOK_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
    )
  ) {
    return null;
  }
  const look: DocumentTableLook = {
    firstRow: false,
    lastRow: false,
    firstColumn: false,
    lastColumn: false,
    noHorizontalBand: false,
    noVerticalBand: false,
  };
  let sawFlag = false;
  const encoded = attribute(element, 'val')?.trim();
  if (encoded) {
    if (!/^[0-9a-f]{1,4}$/i.test(encoded)) return null;
    const value = Number.parseInt(encoded, 16);
    look.firstRow = Boolean(value & TABLE_LOOK_BITS.firstRow);
    look.lastRow = Boolean(value & TABLE_LOOK_BITS.lastRow);
    look.firstColumn = Boolean(value & TABLE_LOOK_BITS.firstColumn);
    look.lastColumn = Boolean(value & TABLE_LOOK_BITS.lastColumn);
    look.noHorizontalBand = Boolean(value & TABLE_LOOK_BITS.noHorizontalBand);
    look.noVerticalBand = Boolean(value & TABLE_LOOK_BITS.noVerticalBand);
    sawFlag = true;
  }
  const flagAssignments: Array<[keyof DocumentTableLook, string]> = [
    ['firstRow', 'firstRow'],
    ['lastRow', 'lastRow'],
    ['firstColumn', 'firstColumn'],
    ['lastColumn', 'lastColumn'],
    ['noHorizontalBand', 'noHBand'],
    ['noVerticalBand', 'noVBand'],
  ];
  for (const [key, attributeName] of flagAssignments) {
    const raw = attribute(element, attributeName);
    if (raw === null) continue;
    const parsed = onOffAttribute(raw);
    if (parsed === undefined) return null;
    look[key] = parsed;
    sawFlag = true;
  }
  return sawFlag ? orderedDocumentTableLook(look) : null;
}

function onOffAttribute(value: string): boolean | undefined {
  const normalized = value.trim().toLowerCase();
  if (
    normalized === '' ||
    normalized === '1' ||
    normalized === 'true' ||
    normalized === 'on' ||
    normalized === 'yes'
  ) {
    return true;
  }
  if (
    normalized === '0' ||
    normalized === 'false' ||
    normalized === 'off' ||
    normalized === 'no'
  ) {
    return false;
  }
  return undefined;
}
