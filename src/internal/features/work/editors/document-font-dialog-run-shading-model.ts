import type { CSSProperties } from 'react';
import {
  type DocumentRunShading,
  type DocumentRunShadingColor,
  type DocumentRunShadingPattern,
  documentRunShadingDomAttributes,
  normalizeDocumentRunShading,
  parseDocumentRunShading,
  serializeDocumentRunShading,
} from '../work-document-run-shading';

export type DocumentFontDialogRunShadingMode =
  | 'inherit'
  | 'mixed'
  | 'none'
  | 'value';

export interface DocumentFontDialogRunShadingSource {
  mixed: boolean;
  value: DocumentRunShading | null;
}

export interface DocumentFontDialogRunShadingDraft {
  runShadingMode: DocumentFontDialogRunShadingMode;
  runShadingPattern: Exclude<DocumentRunShadingPattern, 'nil'>;
  runShadingColor: 'auto' | `#${string}`;
  runShadingFill: 'auto' | `#${string}`;
}

export interface DocumentFontDialogRunShadingPatch {
  runShading?: DocumentRunShading | null;
}

export function addDocumentFontDialogRunShadingValue(
  values: Map<string, DocumentRunShading | null>,
  source: unknown,
): void {
  const shading = parseDocumentRunShading(source);
  const serialized = serializeDocumentRunShading(shading);
  values.set(serialized ?? 'inherited', shading);
}

export function selectedDocumentFontDialogRunShadingValue(
  values: ReadonlyMap<string, DocumentRunShading | null>,
): DocumentFontDialogRunShadingSource {
  return values.size === 1
    ? { mixed: false, value: values.values().next().value ?? null }
    : { mixed: true, value: null };
}

export function createDocumentFontDialogRunShadingDraft(
  source: DocumentFontDialogRunShadingSource,
): DocumentFontDialogRunShadingDraft {
  const shading = source.value;
  return {
    runShadingMode: source.mixed
      ? 'mixed'
      : !shading
        ? 'inherit'
        : shading.pattern === 'nil'
          ? 'none'
          : 'value',
    runShadingPattern:
      shading && shading.pattern !== 'nil' ? shading.pattern : 'clear',
    runShadingColor: shading?.color?.value ?? 'auto',
    runShadingFill: shading?.fill?.value ?? 'auto',
  };
}

export function documentFontDialogRunShadingDraftError(
  draft: DocumentFontDialogRunShadingDraft,
): string | null {
  if (draft.runShadingMode !== 'value') return null;
  if (!validColor(draft.runShadingColor)) {
    return '请选择有效的字符底纹前景色。';
  }
  if (!validColor(draft.runShadingFill)) {
    return '请选择有效的字符底纹背景色。';
  }
  return null;
}

export function documentFontDialogRunShadingPatch(
  source: DocumentFontDialogRunShadingSource,
  draft: DocumentFontDialogRunShadingDraft,
  touched: boolean,
): DocumentFontDialogRunShadingPatch {
  if (!touched || draft.runShadingMode === 'mixed') return {};
  if (draft.runShadingMode === 'inherit') {
    return source.mixed || source.value ? { runShading: null } : {};
  }
  const shading =
    draft.runShadingMode === 'none'
      ? ({ pattern: 'nil' } as const)
      : documentFontDialogRunShadingFromDraft(source, draft);
  if (!shading) return {};
  return serializeDocumentRunShading(source.value) ===
    serializeDocumentRunShading(shading)
    ? {}
    : { runShading: shading };
}

export function documentFontDialogRunShadingFromDraft(
  source: DocumentFontDialogRunShadingSource,
  draft: DocumentFontDialogRunShadingDraft,
): DocumentRunShading | null {
  if (draft.runShadingMode === 'none') return { pattern: 'nil' };
  if (draft.runShadingMode !== 'value') return source.value;
  return normalizeDocumentRunShading({
    pattern: draft.runShadingPattern,
    ...shadingColorPatch('color', draft.runShadingColor, source.value?.color),
    ...shadingColorPatch('fill', draft.runShadingFill, source.value?.fill),
  });
}

export function documentFontDialogRunShadingPreviewStyle(
  source: DocumentFontDialogRunShadingSource,
  draft: DocumentFontDialogRunShadingDraft,
): CSSProperties {
  const shading = documentFontDialogRunShadingFromDraft(source, draft);
  const cssText = documentRunShadingDomAttributes(shading).style;
  if (!cssText) return {};
  const probe = document.createElement('span');
  probe.style.cssText = cssText;
  return {
    backgroundColor: probe.style.backgroundColor || undefined,
    backgroundImage: probe.style.backgroundImage || undefined,
    backgroundSize: probe.style.backgroundSize || undefined,
    boxDecorationBreak: 'clone',
    WebkitBoxDecorationBreak: 'clone',
  };
}

function shadingColorPatch(
  name: 'color' | 'fill',
  value: DocumentFontDialogRunShadingDraft[
    | 'runShadingColor'
    | 'runShadingFill'],
  source: DocumentRunShadingColor | undefined,
): Partial<Record<'color' | 'fill', DocumentRunShadingColor>> {
  if (!source && value === 'auto') return {};
  return {
    [name]: {
      value,
      ...(source?.theme && source.value === value
        ? { theme: source.theme }
        : {}),
    },
  };
}

function validColor(value: string): boolean {
  return value === 'auto' || /^#[0-9a-f]{6}$/i.test(value);
}
