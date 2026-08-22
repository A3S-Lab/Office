import type {
  WorkDocumentStrikeFormatting,
  WorkDocumentStrikeStyle,
} from './work-document-strike';
import { attribute, directChild } from './work-ooxml-package';

export interface ImportedDocxStrikeFlags {
  double?: boolean;
  single?: boolean;
}

export function importedDocxStrikeFlags(
  properties: Element,
): ImportedDocxStrikeFlags {
  return {
    ...importedFlag(directChild(properties, 'strike'), 'single'),
    ...importedFlag(directChild(properties, 'dstrike'), 'double'),
  };
}

export function importedDocxStrike(
  properties: Element,
): WorkDocumentStrikeFormatting | undefined {
  const flags = importedDocxStrikeFlags(properties);
  return resolvedDocxStrikeFormatting(flags.single, flags.double);
}

export function resolvedDocxStrikeFormatting(
  single: boolean | undefined,
  double: boolean | undefined,
  materializeNone = false,
): WorkDocumentStrikeFormatting | undefined {
  if (single === undefined && double === undefined && !materializeNone) {
    return undefined;
  }
  const style: WorkDocumentStrikeStyle = double
    ? 'double'
    : single
      ? 'single'
      : 'none';
  return { style };
}

function importedFlag(
  element: Element | null | undefined,
  name: keyof ImportedDocxStrikeFlags,
): ImportedDocxStrikeFlags {
  if (!element) return {};
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return {
    [name]: value !== '0' && value !== 'false' && value !== 'off',
  };
}
