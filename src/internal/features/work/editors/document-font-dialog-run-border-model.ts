import {
  DOCUMENT_RUN_BORDER_STYLES,
  type DocumentRunBorder,
  type DocumentRunBorderStyle,
  normalizeDocumentRunBorder,
  parseDocumentRunBorder,
  serializeDocumentRunBorder,
} from '../work-document-run-border';

export type DocumentFontDialogRunBorderMode =
  | 'inherit'
  | 'mixed'
  | 'none'
  | 'value';

export interface DocumentFontDialogRunBorderSource {
  mixed: boolean;
  value: DocumentRunBorder | null;
}

export interface DocumentFontDialogRunBorderDraft {
  runBorderMode: DocumentFontDialogRunBorderMode;
  runBorderStyle: DocumentRunBorderStyle;
  runBorderColor: 'auto' | `#${string}`;
  runBorderWidthPoints: string;
  runBorderSpacingPoints: string;
  runBorderShadow: boolean;
  runBorderFrame: boolean;
}

export interface DocumentFontDialogRunBorderPatch {
  runBorder?: DocumentRunBorder | null;
}

export function addDocumentFontDialogRunBorderValue(
  values: Map<string, DocumentRunBorder | null>,
  source: unknown,
): void {
  const border = parseDocumentRunBorder(source);
  const serialized = serializeDocumentRunBorder(border);
  values.set(serialized ?? 'inherited', border);
}

export function selectedDocumentFontDialogRunBorderValue(
  values: ReadonlyMap<string, DocumentRunBorder | null>,
): DocumentFontDialogRunBorderSource {
  return values.size === 1
    ? { mixed: false, value: values.values().next().value ?? null }
    : { mixed: true, value: null };
}

export function createDocumentFontDialogRunBorderDraft(
  source: DocumentFontDialogRunBorderSource,
): DocumentFontDialogRunBorderDraft {
  const border = source.value;
  const visible = border && border.style !== 'nil' && border.style !== 'none';
  return {
    runBorderMode: source.mixed
      ? 'mixed'
      : !border
        ? 'inherit'
        : visible
          ? 'value'
          : 'none',
    runBorderStyle: visible ? border.style : 'single',
    runBorderColor: border?.color?.value ?? 'auto',
    runBorderWidthPoints: formatPoints((border?.size ?? 4) / 8),
    runBorderSpacingPoints: String(border?.space ?? 1),
    runBorderShadow: border?.shadow ?? false,
    runBorderFrame: border?.frame ?? false,
  };
}

export function documentFontDialogRunBorderDraftError(
  draft: DocumentFontDialogRunBorderDraft,
): string | null {
  if (draft.runBorderMode !== 'value') return null;
  if (
    !DOCUMENT_RUN_BORDER_STYLES.includes(draft.runBorderStyle) ||
    draft.runBorderStyle === 'nil' ||
    draft.runBorderStyle === 'none'
  ) {
    return '请选择有效的字符边框线型。';
  }
  if (
    draft.runBorderColor !== 'auto' &&
    !/^#[0-9a-f]{6}$/i.test(draft.runBorderColor)
  ) {
    return '请选择有效的字符边框颜色。';
  }
  if (runBorderSizeFromDraft(draft) === null) {
    return '请输入 0.25 至 12 磅、以 0.125 磅递增的字符边框宽度。';
  }
  if (runBorderSpaceFromDraft(draft) === null) {
    return '请输入 0 至 31 磅的整数字符边框间距。';
  }
  return null;
}

export function documentFontDialogRunBorderPatch(
  source: DocumentFontDialogRunBorderSource,
  draft: DocumentFontDialogRunBorderDraft,
  touched: boolean,
): DocumentFontDialogRunBorderPatch {
  if (!touched || draft.runBorderMode === 'mixed') return {};
  if (draft.runBorderMode === 'inherit') {
    return source.mixed || source.value ? { runBorder: null } : {};
  }
  const border =
    draft.runBorderMode === 'none'
      ? ({ style: 'nil' } as const)
      : documentFontDialogRunBorderFromDraft(source, draft);
  if (!border) return {};
  return serializeDocumentRunBorder(source.value) ===
    serializeDocumentRunBorder(border)
    ? {}
    : { runBorder: border };
}

export function documentFontDialogRunBorderFromDraft(
  source: DocumentFontDialogRunBorderSource,
  draft: DocumentFontDialogRunBorderDraft,
): DocumentRunBorder | null {
  if (draft.runBorderMode === 'none') return { style: 'nil' };
  if (draft.runBorderMode !== 'value') return source.value;
  const size = runBorderSizeFromDraft(draft);
  const space = runBorderSpaceFromDraft(draft);
  if (size === null || space === null) return null;
  const sourceColor = source.value?.color;
  return normalizeDocumentRunBorder({
    style: draft.runBorderStyle,
    color: {
      value: draft.runBorderColor,
      ...(sourceColor?.theme && sourceColor.value === draft.runBorderColor
        ? { theme: sourceColor.theme }
        : {}),
    },
    size,
    space,
    shadow: draft.runBorderShadow,
    frame: draft.runBorderFrame,
  });
}

function runBorderSizeFromDraft(
  draft: DocumentFontDialogRunBorderDraft,
): number | null {
  const points = Number(draft.runBorderWidthPoints);
  const size = Math.round(points * 8);
  return Number.isFinite(points) &&
    size >= 2 &&
    size <= 96 &&
    points * 8 === size
    ? size
    : null;
}

function runBorderSpaceFromDraft(
  draft: DocumentFontDialogRunBorderDraft,
): number | null {
  const space = Number(draft.runBorderSpacingPoints);
  return Number.isSafeInteger(space) && space >= 0 && space <= 31
    ? space
    : null;
}

function formatPoints(value: number): string {
  return Number(value.toFixed(3)).toString();
}
