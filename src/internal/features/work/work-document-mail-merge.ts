import type { Editor } from '@tiptap/core';
import {
  docxDocumentFieldKind,
  docxDocumentFieldTarget,
  documentFieldKind,
  type WorkDocumentFieldContext,
  type WorkDocumentFieldContextResolver,
  type WorkDocumentFieldRefreshOptions,
} from './work-document-fields';
import type { WorkDocumentContent } from './work-types';

/**
 * One mail-merge data row. Keys are MERGEFIELD names; values are plain text
 * substitutions for preview/refresh. Hosts own the data source.
 */
export type WorkDocumentMailMergeRecord = Readonly<Record<string, string>>;

/** Bounded comparison operators for recipient-filter rules (AND-combined). */
export type WorkDocumentMailMergeFilterOperator =
  | 'equals'
  | 'notEquals'
  | 'contains'
  | 'notContains'
  | 'startsWith'
  | 'endsWith'
  | 'isBlank'
  | 'isNotBlank';

/** One recipient-filter predicate against a MERGEFIELD / column name. */
export interface WorkDocumentMailMergeFilterRule {
  field: string;
  operator: WorkDocumentMailMergeFilterOperator;
  /** Compared text for operators other than isBlank / isNotBlank. */
  value?: string;
}

/** AND-combined recipient filter over host records. */
export interface WorkDocumentMailMergeRecipientFilter {
  rules: readonly WorkDocumentMailMergeFilterRule[];
}

/**
 * Explicit host contract for bounded mail-merge preview and batch generation.
 *
 * Unsupported MERGEFIELD switches remain outside this slice.
 */
export interface WorkDocumentMailMergeSource {
  /** Ordered recipient/data rows (unfiltered host data). */
  records: readonly WorkDocumentMailMergeRecord[];
  /**
   * Zero-based active preview row within {@link filteredMailMergeRecords}.
   */
  activeIndex: number;
  /** Optional AND recipient filter; preview and batch use the filtered view. */
  filter?: WorkDocumentMailMergeRecipientFilter | null;
}

export const DOCUMENT_MAIL_MERGE_MAX_RECORDS = 10_000;
export const DOCUMENT_MAIL_MERGE_MAX_BATCH_DOCUMENTS = 256;
export const DOCUMENT_MAIL_MERGE_MAX_FIELDS = 256;
export const DOCUMENT_MAIL_MERGE_MAX_FIELD_NAME = 64;
export const DOCUMENT_MAIL_MERGE_MAX_FIELD_VALUE = 10_000;
export const DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES = 8;

const FIELD_NAME_PATTERN = /^[\p{L}_][\p{L}\p{N}_]*$/u;

const FILTER_OPERATORS = new Set<WorkDocumentMailMergeFilterOperator>([
  'equals',
  'notEquals',
  'contains',
  'notContains',
  'startsWith',
  'endsWith',
  'isBlank',
  'isNotBlank',
]);

export function emptyMailMergeSource(): WorkDocumentMailMergeSource {
  return { records: [], activeIndex: 0, filter: null };
}

export function normalizeMailMergeRecipientFilter(
  value: Partial<WorkDocumentMailMergeRecipientFilter> | null | undefined,
): WorkDocumentMailMergeRecipientFilter | null {
  if (!value || !Array.isArray(value.rules) || !value.rules.length) return null;
  const rules: WorkDocumentMailMergeFilterRule[] = [];
  for (const entry of value.rules) {
    if (rules.length >= DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES) break;
    const rule = normalizeMailMergeFilterRule(entry);
    if (rule) rules.push(rule);
  }
  return rules.length ? { rules } : null;
}

export function normalizeMailMergeSource(
  value: Partial<WorkDocumentMailMergeSource> | null | undefined,
): WorkDocumentMailMergeSource {
  const records: WorkDocumentMailMergeRecord[] = [];
  const rawRecords = Array.isArray(value?.records) ? value.records : [];
  for (const entry of rawRecords) {
    if (records.length >= DOCUMENT_MAIL_MERGE_MAX_RECORDS) break;
    const normalized = normalizeMailMergeRecord(entry);
    if (normalized) records.push(normalized);
  }
  const filter = normalizeMailMergeRecipientFilter(value?.filter ?? null);
  const visible = applyMailMergeRecipientFilter(records, filter);
  const activeIndex = Number.isSafeInteger(value?.activeIndex)
    ? Math.max(
        0,
        Math.min(
          visible.length === 0 ? 0 : visible.length - 1,
          Number(value?.activeIndex),
        ),
      )
    : 0;
  return { records, activeIndex, filter };
}

export function normalizeMailMergeRecord(
  value: unknown,
): WorkDocumentMailMergeRecord | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record: Record<string, string> = {};
  for (const [rawName, rawValue] of Object.entries(
    value as Record<string, unknown>,
  )) {
    if (Object.keys(record).length >= DOCUMENT_MAIL_MERGE_MAX_FIELDS) break;
    const name = normalizeMailMergeFieldName(rawName);
    if (!name || Object.prototype.hasOwnProperty.call(record, name)) continue;
    if (typeof rawValue !== 'string') continue;
    record[name] = normalizeMailMergeFieldValue(rawValue);
  }
  return Object.keys(record).length ? record : null;
}

/** Recipient rows visible after applying the optional AND filter. */
export function filteredMailMergeRecords(
  source: WorkDocumentMailMergeSource | null | undefined,
): readonly WorkDocumentMailMergeRecord[] {
  const normalized = normalizeMailMergeSource(source);
  return applyMailMergeRecipientFilter(normalized.records, normalized.filter);
}

export function mailMergeRecordMatchesFilter(
  record: WorkDocumentMailMergeRecord,
  filter: WorkDocumentMailMergeRecipientFilter | null | undefined,
): boolean {
  const normalized = normalizeMailMergeRecipientFilter(filter);
  if (!normalized) return true;
  return normalized.rules.every((rule) =>
    mailMergeRecordMatchesRule(record, rule),
  );
}

export function setMailMergeRecipientFilter(
  source: WorkDocumentMailMergeSource | null | undefined,
  filter: WorkDocumentMailMergeRecipientFilter | null | undefined,
): WorkDocumentMailMergeSource {
  const normalized = normalizeMailMergeSource(source);
  return normalizeMailMergeSource({
    ...normalized,
    filter: normalizeMailMergeRecipientFilter(filter),
    activeIndex: 0,
  });
}

export function mailMergeFieldNamesFromRecords(
  records: readonly WorkDocumentMailMergeRecord[],
): string[] {
  const names: string[] = [];
  for (const record of records) {
    for (const name of Object.keys(record)) {
      names.push(name);
    }
  }
  return uniqueMailMergeFieldNames(names);
}

export function activeMailMergeRecord(
  source: WorkDocumentMailMergeSource | null | undefined,
): WorkDocumentMailMergeRecord | null {
  const normalized = normalizeMailMergeSource(source);
  const visible = applyMailMergeRecipientFilter(
    normalized.records,
    normalized.filter,
  );
  if (!visible.length) return null;
  return visible[normalized.activeIndex] ?? null;
}

export function stepMailMergeSource(
  source: WorkDocumentMailMergeSource | null | undefined,
  delta: number,
): WorkDocumentMailMergeSource {
  const normalized = normalizeMailMergeSource(source);
  const visible = applyMailMergeRecipientFilter(
    normalized.records,
    normalized.filter,
  );
  if (!visible.length || !Number.isSafeInteger(delta) || delta === 0) {
    return normalized;
  }
  return {
    ...normalized,
    activeIndex: Math.max(
      0,
      Math.min(visible.length - 1, normalized.activeIndex + delta),
    ),
  };
}

export function mailMergeFieldNamesFromHtml(html: string): string[] {
  const document = new DOMParser().parseFromString(html, 'text/html');
  return uniqueMailMergeFieldNames(
    Array.from(
      document.body.querySelectorAll<HTMLElement>('[data-document-field]'),
    ).flatMap((element) => {
      const kind =
        documentFieldKind(element.dataset.fieldKind) ??
        docxDocumentFieldKind(element.dataset.fieldInstruction ?? '');
      if (kind !== 'mergeField') return [];
      const name =
        normalizeMailMergeFieldName(element.dataset.fieldTargetName) ??
        normalizeMailMergeFieldName(
          docxDocumentFieldTarget(element.dataset.fieldInstruction ?? '') ?? '',
        );
      return name ? [name] : [];
    }),
  );
}

export function mailMergeFieldNamesFromEditor(editor: Editor): string[] {
  const names: string[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'documentField') return;
    const kind =
      documentFieldKind(node.attrs.kind) ??
      docxDocumentFieldKind(String(node.attrs.instruction ?? ''));
    if (kind !== 'mergeField') return;
    const name =
      normalizeMailMergeFieldName(node.attrs.targetName) ??
      normalizeMailMergeFieldName(
        docxDocumentFieldTarget(String(node.attrs.instruction ?? '')) ?? '',
      );
    if (name) names.push(name);
  });
  return uniqueMailMergeFieldNames(names);
}

/**
 * Layers the active mail-merge record onto a base field-context resolver
 * (typically pagination). When no base resolver exists, still returns a
 * minimal context so MERGEFIELD preview works in web/continuous views.
 */
export function createMailMergeFieldContextResolver(
  base: WorkDocumentFieldContextResolver | null | undefined,
  source: WorkDocumentMailMergeSource | null | undefined,
): WorkDocumentFieldContextResolver | null {
  const record = activeMailMergeRecord(source);
  if (!base && !record) return null;
  return (position): WorkDocumentFieldContext | null => {
    const context =
      base?.(position) ??
      ({
        pageNumber: 1,
        totalPages: 1,
        sectionNumber: 1,
        sectionPages: 1,
      } satisfies WorkDocumentFieldContext);
    if (!context) return null;
    if (!record) return context;
    return {
      ...context,
      mergeRecord: record,
    };
  };
}

/** Refreshes live fields using the active row from a host mail-merge source. */
export function previewMailMergeSource(
  editor: Editor,
  content: WorkDocumentContent,
  source: WorkDocumentMailMergeSource | null | undefined,
  options: Omit<WorkDocumentFieldRefreshOptions, 'resolveContext'> & {
    resolveContext?: WorkDocumentFieldContextResolver | null;
  } = {},
): boolean {
  const resolveContext = createMailMergeFieldContextResolver(
    options.resolveContext ?? null,
    source,
  );
  if (!resolveContext) return false;
  return editor.commands.refreshDocumentFields(content, {
    ...options,
    resolveContext,
  });
}

/** One generated DOCX for a single mail-merge recipient row. */
export interface WorkDocumentMailMergeGeneratedDocument {
  index: number;
  record: WorkDocumentMailMergeRecord;
  blob: Blob;
}

/**
 * Generates one DOCX blob per filtered recipient row for the current editor
 * template.
 *
 * Caps at `DOCUMENT_MAIL_MERGE_MAX_BATCH_DOCUMENTS`. Restores the editor to the
 * source's active preview row after generation.
 */
export async function generateMailMergeDocuments(
  editor: Editor,
  content: WorkDocumentContent,
  source: WorkDocumentMailMergeSource | null | undefined,
  options: Omit<WorkDocumentFieldRefreshOptions, 'resolveContext'> & {
    resolveContext?: WorkDocumentFieldContextResolver | null;
    exportDocument?: (next: WorkDocumentContent) => Promise<Blob>;
    maxDocuments?: number;
  } = {},
): Promise<readonly WorkDocumentMailMergeGeneratedDocument[]> {
  const normalized = normalizeMailMergeSource(source);
  const visible = applyMailMergeRecipientFilter(
    normalized.records,
    normalized.filter,
  );
  if (!visible.length) return [];
  const maxDocuments = Math.max(
    1,
    Math.min(
      DOCUMENT_MAIL_MERGE_MAX_BATCH_DOCUMENTS,
      Number.isSafeInteger(options.maxDocuments)
        ? Number(options.maxDocuments)
        : DOCUMENT_MAIL_MERGE_MAX_BATCH_DOCUMENTS,
    ),
  );
  const exportDocument =
    options.exportDocument ??
    (async (next) => {
      const { createDocxBlob } = await import('./work-docx-export');
      return createDocxBlob(next);
    });
  const generated: WorkDocumentMailMergeGeneratedDocument[] = [];
  const templateHtml = content.html;
  const restoreHtml = editor.getHTML();
  try {
    for (let index = 0; index < visible.length; index += 1) {
      if (generated.length >= maxDocuments) break;
      const record = visible[index]!;
      const rowSource: WorkDocumentMailMergeSource = {
        ...normalized,
        activeIndex: index,
      };
      editor.commands.setContent(templateHtml, { emitUpdate: false });
      const previewContent: WorkDocumentContent = {
        ...content,
        html: templateHtml,
      };
      if (
        !previewMailMergeSource(editor, previewContent, rowSource, {
          resolveContext: options.resolveContext,
        })
      ) {
        throw new Error(
          `Mail-merge batch generation failed while refreshing recipient row ${index}.`,
        );
      }
      const nextContent: WorkDocumentContent = {
        ...content,
        html: editor.getHTML(),
      };
      generated.push({
        index,
        record,
        blob: await exportDocument(nextContent),
      });
    }
  } finally {
    editor.commands.setContent(restoreHtml, { emitUpdate: false });
    previewMailMergeSource(
      editor,
      { ...content, html: restoreHtml },
      normalized,
      { resolveContext: options.resolveContext },
    );
  }
  return generated;
}

function normalizeMailMergeFilterRule(
  value: unknown,
): WorkDocumentMailMergeFilterRule | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const entry = value as Partial<WorkDocumentMailMergeFilterRule>;
  const field = normalizeMailMergeFieldName(entry.field);
  const operator = mailMergeFilterOperator(entry.operator);
  if (!field || !operator) return null;
  if (operator === 'isBlank' || operator === 'isNotBlank') {
    return { field, operator };
  }
  if (typeof entry.value !== 'string') return null;
  return {
    field,
    operator,
    value: normalizeMailMergeFieldValue(entry.value),
  };
}

function mailMergeFilterOperator(
  value: unknown,
): WorkDocumentMailMergeFilterOperator | null {
  return typeof value === 'string' &&
    FILTER_OPERATORS.has(value as WorkDocumentMailMergeFilterOperator)
    ? (value as WorkDocumentMailMergeFilterOperator)
    : null;
}

function applyMailMergeRecipientFilter(
  records: readonly WorkDocumentMailMergeRecord[],
  filter: WorkDocumentMailMergeRecipientFilter | null | undefined,
): readonly WorkDocumentMailMergeRecord[] {
  const normalized = normalizeMailMergeRecipientFilter(filter);
  if (!normalized) return records;
  return records.filter((record) =>
    mailMergeRecordMatchesFilter(record, normalized),
  );
}

function mailMergeRecordMatchesRule(
  record: WorkDocumentMailMergeRecord,
  rule: WorkDocumentMailMergeFilterRule,
): boolean {
  const raw = record[rule.field] ?? '';
  const value = rule.value ?? '';
  switch (rule.operator) {
    case 'equals':
      return raw === value;
    case 'notEquals':
      return raw !== value;
    case 'contains':
      return raw.includes(value);
    case 'notContains':
      return !raw.includes(value);
    case 'startsWith':
      return raw.startsWith(value);
    case 'endsWith':
      return raw.endsWith(value);
    case 'isBlank':
      return raw.trim().length === 0;
    case 'isNotBlank':
      return raw.trim().length > 0;
    default:
      return false;
  }
}

function normalizeMailMergeFieldName(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const name = value.trim();
  if (
    !name ||
    name.length > DOCUMENT_MAIL_MERGE_MAX_FIELD_NAME ||
    !FIELD_NAME_PATTERN.test(name)
  ) {
    return null;
  }
  return name;
}

function normalizeMailMergeFieldValue(value: string): string {
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/gu, '')
    .slice(0, DOCUMENT_MAIL_MERGE_MAX_FIELD_VALUE);
}

function uniqueMailMergeFieldNames(names: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const name of names) {
    if (seen.has(name)) continue;
    seen.add(name);
    result.push(name);
  }
  return result;
}
