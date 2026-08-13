import * as Y from 'yjs';
import type {
  WorkDocumentBibliography,
  WorkDocumentCitationSource,
} from '../features/work/work-types';
import {
  assertNoAddedCollision,
  changedRecordMissing,
  cloneJsonRecord,
  cloneJsonValue,
  insertIntoOrder,
  invalidInputSidecars,
  invalidSharedSidecars,
  isRecord,
  jsonEqual,
  optionalSharedString,
  patchOptionalScalar,
  patchRequiredScalar,
  removeFromOrder,
  requiredIdentifier,
  requiredSharedMap,
  requiredString,
  validatedOrder,
} from './office-document-collaboration-sidecar-utils';

export function validatedWorkOfficeDocumentBibliography(
  value: unknown,
): WorkDocumentBibliography {
  if (!isRecord(value) || !isCitationStyle(value.style)) {
    invalidInputSidecars('a supported bibliography style');
  }
  if (!Array.isArray(value.sources)) {
    invalidInputSidecars('an array of bibliography sources');
  }
  const ids = new Set<string>();
  const bibliography: WorkDocumentBibliography = {
    style: value.style,
    sources: value.sources.map((source) => {
      const validated = validatedSource(source);
      if (ids.has(validated.id)) {
        invalidInputSidecars(
          `a unique bibliography source ID; '${validated.id}' is repeated`,
        );
      }
      ids.add(validated.id);
      return validated;
    }),
  };
  if (value.styleName !== undefined) {
    bibliography.styleName = requiredString(
      value.styleName,
      'bibliography style name',
    );
  }
  if (value.selectedStyle !== undefined) {
    bibliography.selectedStyle = requiredString(
      value.selectedStyle,
      'selected bibliography style',
    );
  }
  return bibliography;
}

export function initializeWorkOfficeDocumentBibliography(
  settings: Y.Map<unknown>,
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  bibliography: WorkDocumentBibliography,
): void {
  setBibliographySettings(settings, bibliography);
  for (const source of bibliography.sources) {
    createSourceRecord(records, source);
    order.push([source.id]);
  }
}

export function readWorkOfficeDocumentBibliography(
  settings: Y.Map<unknown>,
  records: Y.Map<unknown>,
  order: Y.Array<string>,
): WorkDocumentBibliography {
  const sourceOrder = validatedOrder(order, records, 'bibliography source');
  const style = settings.get('style');
  if (style !== undefined && !isCitationStyle(style)) {
    invalidSharedSidecars('bibliography style');
  }
  const bibliography: WorkDocumentBibliography = {
    style: style ?? 'apa',
    sources: sourceOrder.map((id) =>
      readSourceRecord(
        requiredSharedMap(records, id, 'bibliography source'),
        id,
      ),
    ),
  };
  const styleName = optionalSharedString(
    settings.get('styleName'),
    'bibliography style name',
  );
  const selectedStyle = optionalSharedString(
    settings.get('selectedStyle'),
    'selected bibliography style',
  );
  if (styleName !== undefined) bibliography.styleName = styleName;
  if (selectedStyle !== undefined) bibliography.selectedStyle = selectedStyle;
  return bibliography;
}

export function patchWorkOfficeDocumentBibliography(
  settings: Y.Map<unknown>,
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: WorkDocumentBibliography | undefined,
  next: WorkDocumentBibliography | undefined,
): void {
  if (jsonEqual(previous, next)) return;
  if (!previous && next) {
    setBibliographySettings(settings, next);
    for (const source of next.sources) {
      if (!records.has(source.id)) createSourceRecord(records, source);
      insertIntoOrder(
        order,
        next.sources.map((item) => item.id),
        source.id,
      );
    }
    return;
  }
  if (previous && !next) {
    for (const source of previous.sources) {
      records.delete(source.id);
      removeFromOrder(order, source.id);
    }
    return;
  }
  if (!previous || !next) return;
  patchRequiredScalar(settings, 'style', previous.style, next.style);
  patchOptionalScalar(
    settings,
    'styleName',
    previous.styleName,
    next.styleName,
  );
  patchOptionalScalar(
    settings,
    'selectedStyle',
    previous.selectedStyle,
    next.selectedStyle,
  );
  const beforeById = new Map(
    previous.sources.map((source) => [source.id, source]),
  );
  const afterById = new Map(next.sources.map((source) => [source.id, source]));
  for (const source of previous.sources) {
    if (afterById.has(source.id)) continue;
    records.delete(source.id);
    removeFromOrder(order, source.id);
  }
  for (const source of next.sources) {
    const before = beforeById.get(source.id);
    if (!before) {
      if (!records.has(source.id)) createSourceRecord(records, source);
      insertIntoOrder(
        order,
        next.sources.map((item) => item.id),
        source.id,
      );
      continue;
    }
    if (jsonEqual(before, source)) continue;
    patchSourceRecord(
      requiredSharedMap(records, source.id, 'bibliography source'),
      before,
      source,
    );
  }
}

export function assertWorkOfficeDocumentBibliographyConflicts(
  previous: WorkDocumentBibliography | undefined,
  next: WorkDocumentBibliography | undefined,
  shared: WorkDocumentBibliography | undefined,
): void {
  const beforeSources = previous?.sources ?? [];
  const afterSources = next?.sources ?? [];
  const sharedSources = shared?.sources ?? [];
  assertNoAddedCollision(
    beforeSources,
    afterSources,
    sharedSources,
    'bibliography source',
  );
  if (previous && next && !jsonEqual(previous, next) && !shared) {
    changedRecordMissing('bibliography', 'settings');
  }
  const beforeById = new Map(
    beforeSources.map((source) => [source.id, source]),
  );
  const sharedIds = new Set(sharedSources.map((source) => source.id));
  for (const source of afterSources) {
    const before = beforeById.get(source.id);
    if (before && !jsonEqual(before, source) && !sharedIds.has(source.id)) {
      changedRecordMissing('bibliography source', source.id);
    }
  }
}

function setBibliographySettings(
  target: Y.Map<unknown>,
  bibliography: WorkDocumentBibliography,
): void {
  target.set('style', bibliography.style);
  if (bibliography.styleName !== undefined) {
    target.set('styleName', bibliography.styleName);
  } else {
    target.delete('styleName');
  }
  if (bibliography.selectedStyle !== undefined) {
    target.set('selectedStyle', bibliography.selectedStyle);
  } else {
    target.delete('selectedStyle');
  }
}

function validatedSource(value: unknown): WorkDocumentCitationSource {
  if (!isRecord(value)) invalidInputSidecars('valid bibliography sources');
  const clone = cloneJsonRecord(value) as unknown as WorkDocumentCitationSource;
  clone.id = requiredIdentifier(value.id, 'bibliography source');
  clone.tag = requiredString(value.tag, 'bibliography source tag');
  clone.sourceType = requiredString(
    value.sourceType,
    'bibliography source type',
  );
  clone.title = requiredString(value.title, 'bibliography source title');
  return clone;
}

function createSourceRecord(
  records: Y.Map<unknown>,
  source: WorkDocumentCitationSource,
): void {
  const record = new Y.Map<unknown>();
  records.set(source.id, record);
  for (const [key, value] of Object.entries(source)) {
    if (value !== undefined) record.set(key, cloneJsonValue(value));
  }
}

function readSourceRecord(
  record: Y.Map<unknown>,
  expectedId: string,
): WorkDocumentCitationSource {
  const raw: Record<string, unknown> = {};
  for (const [key, value] of record.entries()) {
    if (value instanceof Y.AbstractType) {
      invalidSharedSidecars(`bibliography source field '${key}'`);
    }
    raw[key] = cloneJsonValue(value);
  }
  const source = validatedSource(raw);
  if (source.id !== expectedId) {
    invalidSharedSidecars('bibliography source identity');
  }
  return source;
}

function patchSourceRecord(
  record: Y.Map<unknown>,
  previous: WorkDocumentCitationSource,
  next: WorkDocumentCitationSource,
): void {
  const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
  keys.delete('id');
  for (const key of keys) {
    const before = (previous as unknown as Record<string, unknown>)[key];
    const after = (next as unknown as Record<string, unknown>)[key];
    if (jsonEqual(before, after)) continue;
    if (after === undefined) record.delete(key);
    else record.set(key, cloneJsonValue(after));
  }
}

function isCitationStyle(
  value: unknown,
): value is WorkDocumentBibliography['style'] {
  return (
    value === 'apa' ||
    value === 'mla' ||
    value === 'chicago' ||
    value === 'ieee'
  );
}
