import {
  normalizeDocumentIndexEntryDraft,
  normalizeDocumentIndexOptions,
  type WorkDocumentIndexEntryDraft,
  type WorkDocumentIndexOptions,
} from './work-document-index';

export type WorkDocumentIndexEntryInstructionResult =
  | { supported: true; value: WorkDocumentIndexEntryDraft }
  | {
      supported: false;
      reason:
        | 'not-index-entry'
        | 'invalid-instruction'
        | 'invalid-entry'
        | 'unsupported-switch';
    };

export type WorkDocumentIndexInstructionResult =
  | { supported: true; options: WorkDocumentIndexOptions }
  | {
      supported: false;
      reason:
        | 'not-index'
        | 'invalid-instruction'
        | 'invalid-columns'
        | 'unsupported-switch';
    };

export function parseDocumentIndexEntryInstruction(
  instruction: string,
): WorkDocumentIndexEntryInstructionResult {
  const command = /^\s*XE\b/i.exec(instruction);
  if (!command) return { supported: false, reason: 'not-index-entry' };
  const parsed = parseLeadingFieldValue(instruction.slice(command[0].length));
  if (!parsed) return { supported: false, reason: 'invalid-instruction' };
  const switches = parseFieldSwitches(parsed.rest);
  if (!switches) return { supported: false, reason: 'invalid-instruction' };
  for (const name of switches.keys()) {
    if (!['b', 'i', 't'].includes(name)) {
      return { supported: false, reason: 'unsupported-switch' };
    }
  }
  const boldValue = valueSwitch(switches, 'b');
  const italicValue = valueSwitch(switches, 'i');
  if (
    (boldValue !== undefined && boldValue !== null) ||
    (italicValue !== undefined && italicValue !== null)
  ) {
    return { supported: false, reason: 'invalid-instruction' };
  }
  const crossReferenceValue = valueSwitch(switches, 't');
  if (switches.has('t') && !crossReferenceValue) {
    return { supported: false, reason: 'invalid-instruction' };
  }
  const terms = splitIndexEntryTerms(parsed.value);
  const value = normalizeDocumentIndexEntryDraft({
    mainEntry: terms.mainEntry,
    subEntry: terms.subEntry,
    crossReference: crossReferenceValue?.replace(/^see\s+/i, '') ?? '',
    pageBold: switches.has('b'),
    pageItalic: switches.has('i'),
  });
  return value
    ? { supported: true, value }
    : { supported: false, reason: 'invalid-entry' };
}

export function documentIndexEntryInstruction(
  source: WorkDocumentIndexEntryDraft,
): string {
  const value = normalizeDocumentIndexEntryDraft(source);
  if (!value) throw new Error('Document index entry is invalid.');
  const term = value.subEntry
    ? `${escapeFieldValue(value.mainEntry)}:${escapeFieldValue(value.subEntry)}`
    : escapeFieldValue(value.mainEntry);
  const switches: string[] = [];
  if (value.crossReference) {
    switches.push(`\\t "See ${escapeFieldValue(value.crossReference)}"`);
  } else {
    if (value.pageBold) switches.push('\\b');
    if (value.pageItalic) switches.push('\\i');
  }
  return `XE "${term}"${switches.length ? ` ${switches.join(' ')}` : ''}`;
}

export function parseDocumentIndexInstruction(
  instruction: string,
): WorkDocumentIndexInstructionResult {
  const command = /^\s*INDEX\b/i.exec(instruction);
  if (!command) return { supported: false, reason: 'not-index' };
  const switches = parseFieldSwitches(instruction.slice(command[0].length));
  if (!switches) return { supported: false, reason: 'invalid-instruction' };
  for (const name of switches.keys()) {
    if (!['c', 'e', 'r'].includes(name)) {
      return { supported: false, reason: 'unsupported-switch' };
    }
  }
  const columnsValue = valueSwitch(switches, 'c');
  if (switches.has('c') && !/^[1-4]$/.test(columnsValue ?? '')) {
    return { supported: false, reason: 'invalid-columns' };
  }
  const separator = valueSwitch(switches, 'e');
  if (switches.has('e') && separator === undefined) {
    return { supported: false, reason: 'invalid-instruction' };
  }
  if (switches.has('r') && valueSwitch(switches, 'r') !== null) {
    return { supported: false, reason: 'invalid-instruction' };
  }
  return {
    supported: true,
    options: normalizeDocumentIndexOptions({
      columns: columnsValue ? Number(columnsValue) : 1,
      format: switches.has('r') ? 'run-in' : 'indented',
      rightAlignPageNumbers: !switches.has('e'),
      leader: switches.has('e') ? 'none' : 'dot',
    }),
  };
}

export function documentIndexInstruction(
  source: WorkDocumentIndexOptions,
): string {
  const options = normalizeDocumentIndexOptions(source);
  const switches = options.columns > 1 ? [`\\c "${options.columns}"`] : [];
  if (options.format === 'run-in') switches.push('\\r');
  if (!options.rightAlignPageNumbers) switches.push('\\e ", "');
  return `INDEX${switches.length ? ` ${switches.join(' ')}` : ''}`;
}

function splitIndexEntryTerms(source: string): {
  mainEntry: string;
  subEntry: string;
} {
  let escaped = false;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === '\\') {
      escaped = true;
      continue;
    }
    if (character === ':') {
      return {
        mainEntry: unescapeFieldValue(source.slice(0, index)),
        subEntry: unescapeFieldValue(source.slice(index + 1)),
      };
    }
  }
  return { mainEntry: unescapeFieldValue(source), subEntry: '' };
}

function parseLeadingFieldValue(
  source: string,
): { value: string; rest: string } | null {
  const trimmed = source.trimStart();
  if (!trimmed) return null;
  if (trimmed[0] !== '"') {
    const match = /^([^\s\\]+)([\s\S]*)$/.exec(trimmed);
    return match?.[1] ? { value: match[1], rest: match[2] ?? '' } : null;
  }
  let escaped = false;
  for (let index = 1; index < trimmed.length; index += 1) {
    const character = trimmed[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === '\\') {
      escaped = true;
      continue;
    }
    if (character === '"') {
      return {
        value: trimmed.slice(1, index),
        rest: trimmed.slice(index + 1),
      };
    }
  }
  return null;
}

function parseFieldSwitches(source: string): Map<string, string | null> | null {
  const switches = new Map<string, string | null>();
  const matcher = /\\([a-z])(?:\s+("([^"\\]*(?:\\.[^"\\]*)*)"|([^\\\s]+)))?/gi;
  let cursor = 0;
  for (let match = matcher.exec(source); match; match = matcher.exec(source)) {
    if (source.slice(cursor, match.index).trim()) return null;
    const name = match[1]?.toLowerCase();
    if (!name || switches.has(name)) return null;
    switches.set(name, unescapeFieldValue(match[3] ?? match[4] ?? '') || null);
    cursor = matcher.lastIndex;
  }
  return source.slice(cursor).trim() ? null : switches;
}

function valueSwitch(
  switches: Map<string, string | null>,
  name: string,
): string | null | undefined {
  return switches.has(name) ? switches.get(name) : undefined;
}

function escapeFieldValue(value: string): string {
  return value
    .replaceAll('\\', '\\\\')
    .replaceAll('"', '\\"')
    .replaceAll(':', '\\:');
}

function unescapeFieldValue(value: string): string {
  return value.replace(/\\([\\":])/g, '$1');
}
