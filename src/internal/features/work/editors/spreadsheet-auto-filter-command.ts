import {
  normalizedWorkSpreadsheetAutoFilterRange,
  normalizeWorkSpreadsheetFilterCriteria,
} from '../work-spreadsheet-auto-filter';
import type { WorkSpreadsheetCustomFilterCondition } from '../work-types';
import {
  applySpreadsheetAutoFilterCriteria,
  clearSpreadsheetAutoFilterCriteria,
  type SpreadsheetAutoFilterCriteriaRequest,
  type SpreadsheetAutoFilterTarget,
} from './spreadsheet-auto-filter';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetAutoFilterExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetAutoFilter',
    addCommands: () => ({
      applyAutoFilterCriteria: {
        canExecute: canApplySpreadsheetAutoFilterCriteria,
        execute: applySpreadsheetAutoFilterCriteriaCommand,
      },
      clearAutoFilterCriteria: {
        canExecute: canClearSpreadsheetAutoFilterCriteria,
        execute: clearSpreadsheetAutoFilterCriteriaCommand,
      },
      openAutoFilterMenu: {
        canExecute: ({ autoFilter }) => autoFilter.canOpenMenu,
        execute: ({ autoFilter }) =>
          autoFilter.canOpenMenu && autoFilter.openMenu(),
      },
      toggleAutoFilter: {
        canExecute: ({ autoFilter }) => autoFilter.canToggle,
        execute: ({ autoFilter }) =>
          autoFilter.canToggle && autoFilter.toggle(),
      },
    }),
  });
}

function canApplySpreadsheetAutoFilterCriteria(
  context: SpreadsheetCommandContext,
  request: SpreadsheetAutoFilterCriteriaRequest,
): boolean {
  return Boolean(
    context.editable &&
      supportedSpreadsheetAutoFilterCriteria(request?.criteria) &&
      currentSpreadsheetAutoFilterTarget(context, request),
  );
}

function applySpreadsheetAutoFilterCriteriaCommand(
  context: SpreadsheetCommandContext,
  request: SpreadsheetAutoFilterCriteriaRequest,
): boolean {
  if (!canApplySpreadsheetAutoFilterCriteria(context, request)) return false;
  const next = applySpreadsheetAutoFilterCriteria(
    context.content,
    request.sheetId,
    request.column,
    request.criteria,
  );
  if (!next) return false;
  context.onChange(next);
  return true;
}

function canClearSpreadsheetAutoFilterCriteria(
  context: SpreadsheetCommandContext,
  target: SpreadsheetAutoFilterTarget,
): boolean {
  const current = currentSpreadsheetAutoFilterTarget(context, target);
  if (!context.editable || !current) return false;
  const relativeColumn = target.column - current.range.column[0];
  return Object.hasOwn(current.sheet.filter ?? {}, String(relativeColumn));
}

function clearSpreadsheetAutoFilterCriteriaCommand(
  context: SpreadsheetCommandContext,
  target: SpreadsheetAutoFilterTarget,
): boolean {
  if (!canClearSpreadsheetAutoFilterCriteria(context, target)) return false;
  const next = clearSpreadsheetAutoFilterCriteria(
    context.content,
    target.sheetId,
    target.column,
  );
  if (!next) return false;
  context.onChange(next);
  return true;
}

function currentSpreadsheetAutoFilterTarget(
  context: SpreadsheetCommandContext,
  target: SpreadsheetAutoFilterTarget | null | undefined,
) {
  if (!target || target.sheetId !== context.activeSheetId) return null;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === target.sheetId,
  );
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet?.filter_select);
  if (
    !sheet ||
    !range ||
    !sameSpreadsheetAutoFilterRange(range, target.filterRange) ||
    !Number.isSafeInteger(target.column) ||
    target.column < range.column[0] ||
    target.column > range.column[1]
  ) {
    return null;
  }
  return { range, sheet };
}

function sameSpreadsheetAutoFilterRange(
  left: SpreadsheetAutoFilterTarget['filterRange'],
  right: SpreadsheetAutoFilterTarget['filterRange'],
): boolean {
  return (
    left.row[0] === right.row[0] &&
    left.row[1] === right.row[1] &&
    left.column[0] === right.column[0] &&
    left.column[1] === right.column[1]
  );
}

function supportedSpreadsheetAutoFilterCriteria(
  criteria: SpreadsheetAutoFilterCriteriaRequest['criteria'] | undefined,
): boolean {
  const normalized = normalizeWorkSpreadsheetFilterCriteria(criteria);
  if (!normalized) return false;
  if (
    ['bottom', 'bottom-percent', 'dynamic', 'top', 'top-percent'].includes(
      normalized.type,
    )
  ) {
    return false;
  }
  if (normalized.type === 'between' || normalized.type === 'not-between') {
    const lower = finiteSpreadsheetAutoFilterNumber(normalized.lower);
    const upper = finiteSpreadsheetAutoFilterNumber(normalized.upper);
    return lower !== null && upper !== null && lower <= upper;
  }
  if (normalized.type === 'compound') {
    return normalized.conditions.every(supportedSpreadsheetAutoFilterCondition);
  }
  if ('value' in normalized) {
    return supportedSpreadsheetAutoFilterCondition(normalized);
  }
  return true;
}

function supportedSpreadsheetAutoFilterCondition(
  condition: WorkSpreadsheetCustomFilterCondition,
): boolean {
  if (
    [
      'greater-than',
      'greater-than-or-equal',
      'less-than',
      'less-than-or-equal',
    ].includes(condition.type)
  ) {
    return finiteSpreadsheetAutoFilterNumber(condition.value) !== null;
  }
  return Boolean(condition.value.trim());
}

function finiteSpreadsheetAutoFilterNumber(value: string): number | null {
  if (!value.trim()) return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}
