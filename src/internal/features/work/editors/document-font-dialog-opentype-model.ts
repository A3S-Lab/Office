import {
  normalizeDocumentOpenTypeFeatures,
  normalizeDocumentOpenTypeLigatures,
  normalizeDocumentOpenTypeNumberForm,
  normalizeDocumentOpenTypeNumberSpacing,
  parseDocumentOpenTypeFeatures,
  type WorkDocumentOpenTypeFeaturePatch,
  type WorkDocumentOpenTypeFeatures,
  type WorkDocumentOpenTypeLigatures,
  type WorkDocumentOpenTypeNumberForm,
  type WorkDocumentOpenTypeNumberSpacing,
} from '../work-document-opentype';

interface DocumentFontDialogOpenTypeValue<T> {
  mixed: boolean;
  value: T | null;
}

export interface DocumentFontDialogOpenTypeSource {
  openTypeLigatures: DocumentFontDialogOpenTypeValue<WorkDocumentOpenTypeLigatures>;
  openTypeNumberForm: DocumentFontDialogOpenTypeValue<WorkDocumentOpenTypeNumberForm>;
  openTypeNumberSpacing: DocumentFontDialogOpenTypeValue<WorkDocumentOpenTypeNumberSpacing>;
  openTypeStylisticSets: DocumentFontDialogOpenTypeValue<readonly number[]>;
  openTypeContextualAlternates: DocumentFontDialogOpenTypeValue<boolean>;
}

export type DocumentOpenTypeLigaturesMode =
  | 'inherit'
  | 'mixed'
  | WorkDocumentOpenTypeLigatures;
export type DocumentOpenTypeNumberFormMode =
  | 'inherit'
  | 'mixed'
  | WorkDocumentOpenTypeNumberForm;
export type DocumentOpenTypeNumberSpacingMode =
  | 'inherit'
  | 'mixed'
  | WorkDocumentOpenTypeNumberSpacing;
export type DocumentOpenTypeStylisticSetsMode =
  | 'inherit'
  | 'mixed'
  | 'multiple'
  | 'none'
  | `set-${number}`;
export type DocumentOpenTypeContextualAlternatesMode =
  | 'disabled'
  | 'enabled'
  | 'inherit'
  | 'mixed';

export interface DocumentFontDialogOpenTypeDraft {
  openTypeLigatures: DocumentOpenTypeLigaturesMode;
  openTypeNumberForm: DocumentOpenTypeNumberFormMode;
  openTypeNumberSpacing: DocumentOpenTypeNumberSpacingMode;
  openTypeStylisticSets: DocumentOpenTypeStylisticSetsMode;
  openTypeContextualAlternates: DocumentOpenTypeContextualAlternatesMode;
}

export interface DocumentFontDialogOpenTypePatch {
  openTypeLigatures?: WorkDocumentOpenTypeLigatures | null;
  openTypeNumberForm?: WorkDocumentOpenTypeNumberForm | null;
  openTypeNumberSpacing?: WorkDocumentOpenTypeNumberSpacing | null;
  openTypeStylisticSets?: readonly number[] | null;
  openTypeContextualAlternates?: boolean | null;
}

export interface DocumentFontDialogOpenTypeTouched {
  openTypeLigatures: boolean;
  openTypeNumberForm: boolean;
  openTypeNumberSpacing: boolean;
  openTypeStylisticSets: boolean;
  openTypeContextualAlternates: boolean;
}

export interface DocumentFontDialogOpenTypeValueMaps {
  ligatures: Map<string, WorkDocumentOpenTypeLigatures | null>;
  numberForm: Map<string, WorkDocumentOpenTypeNumberForm | null>;
  numberSpacing: Map<string, WorkDocumentOpenTypeNumberSpacing | null>;
  stylisticSets: Map<string, readonly number[] | null>;
  contextualAlternates: Map<string, boolean | null>;
}

export function createDocumentFontDialogOpenTypeValueMaps(): DocumentFontDialogOpenTypeValueMaps {
  return {
    ligatures: new Map(),
    numberForm: new Map(),
    numberSpacing: new Map(),
    stylisticSets: new Map(),
    contextualAlternates: new Map(),
  };
}

export function addDocumentFontDialogOpenTypeValues(
  values: DocumentFontDialogOpenTypeValueMaps,
  serialized: unknown,
): void {
  const features = parseDocumentOpenTypeFeatures(serialized);
  addValue(values.ligatures, features?.ligatures ?? null);
  addValue(values.numberForm, features?.numberForm ?? null);
  addValue(values.numberSpacing, features?.numberSpacing ?? null);
  const stylisticSets = features?.stylisticSets ?? null;
  values.stylisticSets.set(
    stylisticSets === null ? 'inherited' : JSON.stringify(stylisticSets),
    stylisticSets,
  );
  addValue(values.contextualAlternates, features?.contextualAlternates ?? null);
}

export function selectedDocumentFontDialogOpenTypeSource(
  values: DocumentFontDialogOpenTypeValueMaps,
): DocumentFontDialogOpenTypeSource {
  return {
    openTypeLigatures: selectedValue(values.ligatures),
    openTypeNumberForm: selectedValue(values.numberForm),
    openTypeNumberSpacing: selectedValue(values.numberSpacing),
    openTypeStylisticSets: selectedValue(values.stylisticSets),
    openTypeContextualAlternates: selectedValue(values.contextualAlternates),
  };
}

export function createDocumentFontDialogOpenTypeDraft(
  source: DocumentFontDialogOpenTypeSource,
): DocumentFontDialogOpenTypeDraft {
  return {
    openTypeLigatures: source.openTypeLigatures.mixed
      ? 'mixed'
      : (source.openTypeLigatures.value ?? 'inherit'),
    openTypeNumberForm: source.openTypeNumberForm.mixed
      ? 'mixed'
      : (source.openTypeNumberForm.value ?? 'inherit'),
    openTypeNumberSpacing: source.openTypeNumberSpacing.mixed
      ? 'mixed'
      : (source.openTypeNumberSpacing.value ?? 'inherit'),
    openTypeStylisticSets: stylisticSetsMode(source.openTypeStylisticSets),
    openTypeContextualAlternates: source.openTypeContextualAlternates.mixed
      ? 'mixed'
      : source.openTypeContextualAlternates.value === null
        ? 'inherit'
        : source.openTypeContextualAlternates.value
          ? 'enabled'
          : 'disabled',
  };
}

export function appendDocumentFontDialogOpenTypePatch(
  patch: DocumentFontDialogOpenTypePatch,
  source: DocumentFontDialogOpenTypeSource,
  draft: DocumentFontDialogOpenTypeDraft,
  touched: DocumentFontDialogOpenTypeTouched,
): void {
  if (touched.openTypeLigatures && draft.openTypeLigatures !== 'mixed') {
    const value =
      draft.openTypeLigatures === 'inherit'
        ? null
        : normalizeDocumentOpenTypeLigatures(draft.openTypeLigatures);
    appendScalarPatch(
      patch,
      'openTypeLigatures',
      source.openTypeLigatures,
      value,
    );
  }
  if (touched.openTypeNumberForm && draft.openTypeNumberForm !== 'mixed') {
    const value =
      draft.openTypeNumberForm === 'inherit'
        ? null
        : normalizeDocumentOpenTypeNumberForm(draft.openTypeNumberForm);
    appendScalarPatch(
      patch,
      'openTypeNumberForm',
      source.openTypeNumberForm,
      value,
    );
  }
  if (
    touched.openTypeNumberSpacing &&
    draft.openTypeNumberSpacing !== 'mixed'
  ) {
    const value =
      draft.openTypeNumberSpacing === 'inherit'
        ? null
        : normalizeDocumentOpenTypeNumberSpacing(draft.openTypeNumberSpacing);
    appendScalarPatch(
      patch,
      'openTypeNumberSpacing',
      source.openTypeNumberSpacing,
      value,
    );
  }
  if (
    touched.openTypeStylisticSets &&
    draft.openTypeStylisticSets !== 'mixed' &&
    draft.openTypeStylisticSets !== 'multiple'
  ) {
    const value = stylisticSetsFromMode(draft.openTypeStylisticSets);
    if (
      source.openTypeStylisticSets.mixed ||
      !sameStylisticSets(source.openTypeStylisticSets.value, value)
    ) {
      patch.openTypeStylisticSets = value;
    }
  }
  if (
    touched.openTypeContextualAlternates &&
    draft.openTypeContextualAlternates !== 'mixed'
  ) {
    const value =
      draft.openTypeContextualAlternates === 'inherit'
        ? null
        : draft.openTypeContextualAlternates === 'enabled';
    appendScalarPatch(
      patch,
      'openTypeContextualAlternates',
      source.openTypeContextualAlternates,
      value,
    );
  }
}

export function documentFontDialogOpenTypeFeaturePatch(
  patch: DocumentFontDialogOpenTypePatch,
): WorkDocumentOpenTypeFeaturePatch {
  return {
    ...(patch.openTypeLigatures !== undefined
      ? { ligatures: patch.openTypeLigatures }
      : {}),
    ...(patch.openTypeNumberForm !== undefined
      ? { numberForm: patch.openTypeNumberForm }
      : {}),
    ...(patch.openTypeNumberSpacing !== undefined
      ? { numberSpacing: patch.openTypeNumberSpacing }
      : {}),
    ...(patch.openTypeStylisticSets !== undefined
      ? { stylisticSets: patch.openTypeStylisticSets }
      : {}),
    ...(patch.openTypeContextualAlternates !== undefined
      ? { contextualAlternates: patch.openTypeContextualAlternates }
      : {}),
  };
}

export function documentFontDialogOpenTypePreviewFeatures(
  source: DocumentFontDialogOpenTypeSource,
  draft: DocumentFontDialogOpenTypeDraft,
): WorkDocumentOpenTypeFeatures | null {
  const features: WorkDocumentOpenTypeFeatures = {};
  const ligatures = draftValue(
    draft.openTypeLigatures,
    source.openTypeLigatures,
  );
  const numberForm = draftValue(
    draft.openTypeNumberForm,
    source.openTypeNumberForm,
  );
  const numberSpacing = draftValue(
    draft.openTypeNumberSpacing,
    source.openTypeNumberSpacing,
  );
  const stylisticSets =
    draft.openTypeStylisticSets === 'mixed' ||
    draft.openTypeStylisticSets === 'multiple'
      ? source.openTypeStylisticSets.value
      : stylisticSetsFromMode(draft.openTypeStylisticSets);
  const contextualAlternates =
    draft.openTypeContextualAlternates === 'mixed'
      ? source.openTypeContextualAlternates.value
      : draft.openTypeContextualAlternates === 'inherit'
        ? null
        : draft.openTypeContextualAlternates === 'enabled';
  if (ligatures !== null) features.ligatures = ligatures;
  if (numberForm !== null) features.numberForm = numberForm;
  if (numberSpacing !== null) features.numberSpacing = numberSpacing;
  if (stylisticSets !== null) features.stylisticSets = [...stylisticSets];
  if (contextualAlternates !== null) {
    features.contextualAlternates = contextualAlternates;
  }
  return normalizeDocumentOpenTypeFeatures(features);
}

function stylisticSetsMode(
  source: DocumentFontDialogOpenTypeSource['openTypeStylisticSets'],
): DocumentOpenTypeStylisticSetsMode {
  if (source.mixed) return 'mixed';
  if (source.value === null) return 'inherit';
  if (!source.value.length) return 'none';
  return source.value.length === 1 ? `set-${source.value[0]}` : 'multiple';
}

function stylisticSetsFromMode(
  mode: Exclude<DocumentOpenTypeStylisticSetsMode, 'mixed' | 'multiple'>,
): readonly number[] | null {
  if (mode === 'inherit') return null;
  if (mode === 'none') return [];
  const id = Number(mode.slice('set-'.length));
  return Number.isSafeInteger(id) && id >= 1 && id <= 20 ? [id] : null;
}

function sameStylisticSets(
  left: readonly number[] | null,
  right: readonly number[] | null,
): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function appendScalarPatch<
  Key extends
    | 'openTypeContextualAlternates'
    | 'openTypeLigatures'
    | 'openTypeNumberForm'
    | 'openTypeNumberSpacing',
>(
  patch: DocumentFontDialogOpenTypePatch,
  key: Key,
  source: DocumentFontDialogOpenTypeSource[Key],
  value: DocumentFontDialogOpenTypePatch[Key],
): void {
  if (source.mixed || source.value !== value) {
    Object.assign(patch, { [key]: value });
  }
}

function draftValue<T extends string>(
  value: 'inherit' | 'mixed' | T,
  source: DocumentFontDialogOpenTypeValue<T>,
): T | null {
  if (value === 'mixed') return source.value;
  return value === 'inherit' ? null : value;
}

function addValue<T extends boolean | string>(
  values: Map<string, T | null>,
  value: T | null,
): void {
  values.set(value === null ? 'inherited' : String(value), value);
}

function selectedValue<T>(values: Map<string, T | null>): {
  mixed: boolean;
  value: T | null;
} {
  const mixed = values.size > 1;
  return {
    mixed,
    value: mixed ? null : (values.values().next().value ?? null),
  };
}
