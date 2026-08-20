import { normalizeSheetProtectionAuthority } from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { normalizeSpreadsheetCellRange } from './spreadsheet-cell-range';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';
import {
  numericSpreadsheetCellValue,
  pasteSpreadsheetSpecialCell,
  spreadsheetClipboardCellIsBlank,
  spreadsheetPasteCellInvalid,
  spreadsheetPasteContentSupportsOperation,
} from './spreadsheet-paste-special-cell';
import {
  columnsContainReadOnlyState,
  pasteContentCopiesBorders,
  pasteContentCopiesFormulas,
  pasteContentCopiesMerges,
} from './spreadsheet-paste-special-mode';
import {
  planSpreadsheetPaste,
  spreadsheetPasteSourceAt,
} from './spreadsheet-paste-special-plan';
import {
  applySpreadsheetPasteBorders,
  applySpreadsheetPasteMerges,
  applySpreadsheetPasteSheetMetadata,
  createSpreadsheetCellWriter,
  pasteSpreadsheetColumnWidths,
  withSpreadsheetPasteMerges,
  withSpreadsheetPasteSelection,
} from './spreadsheet-paste-special-sheet';
import {
  createSpreadsheetCellReader,
  spreadsheetPasteSpecialModeAvailable,
  validSpreadsheetClipboardSnapshot,
} from './spreadsheet-paste-special-snapshot';
import {
  spreadsheetPasteOperationOptions,
  type SpreadsheetClipboardSnapshot,
  type SpreadsheetPastePlan,
  type SpreadsheetPasteSpecialOptions,
  type SpreadsheetPasteSpecialRequest,
  type SpreadsheetPasteSpecialResult,
} from './spreadsheet-paste-special-types';

export {
  captureSpreadsheetClipboardSnapshot,
  createSpreadsheetTextClipboardSnapshot,
  normalizeSpreadsheetClipboardText,
  spreadsheetPasteSpecialModeAvailable,
} from './spreadsheet-paste-special-snapshot';
export {
  MAX_SPREADSHEET_COLUMNS,
  MAX_SPREADSHEET_PASTE_SPECIAL_CELLS,
  MAX_SPREADSHEET_ROWS,
  spreadsheetPasteContentOptions,
  spreadsheetPasteOperationOptions,
} from './spreadsheet-paste-special-types';
export type {
  SpreadsheetClipboardCell,
  SpreadsheetClipboardMerge,
  SpreadsheetClipboardSnapshot,
  SpreadsheetPasteContent,
  SpreadsheetPasteOperation,
  SpreadsheetPasteSpecialOptions,
  SpreadsheetPasteSpecialRequest,
  SpreadsheetPasteSpecialResult,
} from './spreadsheet-paste-special-types';

export function spreadsheetPasteSpecialValidationError(
  content: WorkSpreadsheetContent,
  request: SpreadsheetPasteSpecialRequest,
): string | null {
  const { snapshot, options } = request;
  if (!validSpreadsheetClipboardSnapshot(snapshot)) {
    return '剪贴板内容无效或超过 50,000 个单元格。';
  }
  if (!spreadsheetPasteSpecialModeAvailable(snapshot, options.content)) {
    return '当前剪贴板不包含此粘贴方式需要的格式信息。';
  }
  if (
    !spreadsheetPasteOperationOptions.some(
      (option) => option.value === options.operation,
    )
  ) {
    return '粘贴运算无效。';
  }

  const targetSelection = normalizeSpreadsheetCellRange(
    request.targetSelection,
  );
  const plan = targetSelection
    ? planSpreadsheetPaste(snapshot, targetSelection, options.transpose)
    : null;
  if (!plan) return '粘贴区域超出了工作表边界。';
  const sheet = content.sheets.find(
    (candidate) => candidate.id === request.targetSheetId,
  );
  if (!sheet) return '目标工作表不存在。';

  if (options.content === 'column-widths') {
    if (
      options.operation !== 'none' ||
      options.skipBlanks ||
      options.transpose ||
      !snapshot.columnWidths?.length
    ) {
      return '列宽粘贴不能与运算、跳过空白或转置组合。';
    }
    const authority = normalizeSheetProtectionAuthority(
      sheet.config?.authority,
    );
    if (
      (authority.sheet === 1 && authority.formatColumns !== 1) ||
      columnsContainReadOnlyState(sheet, plan.targetRange)
    ) {
      return '目标列受保护或为只读状态，无法粘贴列宽。';
    }
    return null;
  }

  if (sheet.isPivotTable || sheet.pivotTable || sheet.pivotTables?.length) {
    return '数据透视表区域不支持选择性粘贴。';
  }
  if (!canMutateSpreadsheetCellRange(sheet, plan.targetRange)) {
    return '目标区域包含合并单元格、保护或只读单元格，无法粘贴。';
  }
  if (
    snapshot.containsUnsupportedFormulaState &&
    pasteContentCopiesFormulas(options.content)
  ) {
    return '当前选区包含不支持移动的数组、共享或外部公式。';
  }
  if (
    snapshot.merges.length > 0 &&
    options.skipBlanks &&
    pasteContentCopiesMerges(options.content)
  ) {
    return '包含合并单元格时不能同时跳过空白。';
  }
  if (
    options.operation !== 'none' &&
    !spreadsheetPasteContentSupportsOperation(options.content)
  ) {
    return '所选粘贴内容不能执行运算。';
  }
  if (options.operation !== 'none') {
    const operationError = spreadsheetPasteOperationError(
      sheet,
      snapshot,
      plan,
      options,
    );
    if (operationError) return operationError;
  }
  return null;
}

export function applySpreadsheetPasteSpecial(
  content: WorkSpreadsheetContent,
  request: SpreadsheetPasteSpecialRequest,
): SpreadsheetPasteSpecialResult | null {
  if (spreadsheetPasteSpecialValidationError(content, request)) return null;
  const targetSelection = normalizeSpreadsheetCellRange(
    request.targetSelection,
  );
  if (!targetSelection) return null;
  const plan = planSpreadsheetPaste(
    request.snapshot,
    targetSelection,
    request.options.transpose,
  );
  const sheetIndex = content.sheets.findIndex(
    (sheet) => sheet.id === request.targetSheetId,
  );
  const sheet = content.sheets[sheetIndex];
  if (!plan || !sheet || sheetIndex < 0) return null;

  if (request.options.content === 'column-widths') {
    const nextSheet = pasteSpreadsheetColumnWidths(
      sheet,
      request.snapshot,
      plan,
    );
    const sheets = [...content.sheets];
    sheets[sheetIndex] = withSpreadsheetPasteSelection(
      nextSheet,
      plan.targetRange,
    );
    return {
      content: { ...content, sheets },
      targetRange: plan.targetRange,
      firstCellValue: createSpreadsheetCellReader(nextSheet)(
        plan.targetRange.row[0],
        plan.targetRange.column[0],
      )?.v,
    };
  }

  const writer = createSpreadsheetCellWriter(sheet);
  const processed: Array<{ row: number; column: number }> = [];
  for (
    let row = plan.targetRange.row[0];
    row <= plan.targetRange.row[1];
    row += 1
  ) {
    for (
      let column = plan.targetRange.column[0];
      column <= plan.targetRange.column[1];
      column += 1
    ) {
      const source = spreadsheetPasteSourceAt(
        request.snapshot,
        plan,
        row,
        column,
        request.options.transpose,
      );
      if (
        request.options.skipBlanks &&
        spreadsheetClipboardCellIsBlank(source.cell, request.options.content)
      ) {
        continue;
      }
      processed.push({ row, column });
      if (request.options.content === 'validation') continue;

      const destination = writer.get(row, column);
      const next = pasteSpreadsheetSpecialCell({
        source: source.cell,
        destination,
        content: request.options.content,
        operation: request.options.operation,
        rowOffset:
          request.snapshot.kind === 'rich' ? row - source.sourceRow : 0,
        columnOffset:
          request.snapshot.kind === 'rich' ? column - source.sourceColumn : 0,
      });
      if (next === spreadsheetPasteCellInvalid) return null;
      writer.set(row, column, next);
    }
  }

  const pastedMerges = pasteContentCopiesMerges(request.options.content)
    ? applySpreadsheetPasteMerges(
        writer,
        request.snapshot,
        plan,
        request.options.transpose,
      )
    : {};
  let nextSheet = withSpreadsheetPasteMerges(
    writer.finish(sheet),
    pastedMerges,
  );
  nextSheet = applySpreadsheetPasteSheetMetadata(
    nextSheet,
    request.snapshot,
    plan,
    processed,
    request.options,
  );
  nextSheet = withSpreadsheetPasteSelection(nextSheet, plan.targetRange);

  let nextContent: WorkSpreadsheetContent = {
    ...content,
    sheets: content.sheets.map((candidate, index) =>
      index === sheetIndex ? nextSheet : candidate,
    ),
  };
  if (pasteContentCopiesBorders(request.options.content)) {
    nextContent = applySpreadsheetPasteBorders(
      nextContent,
      request.targetSheetId,
      request.snapshot,
      plan,
      processed,
      request.options.transpose,
    );
  }
  const finalSheet = nextContent.sheets[sheetIndex];
  return {
    content: nextContent,
    targetRange: plan.targetRange,
    firstCellValue: finalSheet
      ? createSpreadsheetCellReader(finalSheet)(
          plan.targetRange.row[0],
          plan.targetRange.column[0],
        )?.v
      : undefined,
  };
}

function spreadsheetPasteOperationError(
  sheet: WorkSpreadsheetSheet,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  options: SpreadsheetPasteSpecialOptions,
): string | null {
  const cellAt = createSpreadsheetCellReader(sheet);
  for (
    let row = plan.targetRange.row[0];
    row <= plan.targetRange.row[1];
    row += 1
  ) {
    for (
      let column = plan.targetRange.column[0];
      column <= plan.targetRange.column[1];
      column += 1
    ) {
      const source = spreadsheetPasteSourceAt(
        snapshot,
        plan,
        row,
        column,
        options.transpose,
      ).cell;
      if (
        options.skipBlanks &&
        spreadsheetClipboardCellIsBlank(source, options.content)
      ) {
        continue;
      }
      const sourceValue = numericSpreadsheetCellValue(source.cell);
      if (sourceValue === null) return '运算只能应用于数值单元格。';
      if (options.operation === 'divide' && sourceValue === 0) {
        return '除数不能为 0。';
      }
      const destination = cellAt(row, column);
      if (
        destination &&
        numericSpreadsheetCellValue(destination) === null &&
        (destination.v !== undefined ||
          destination.f ||
          destination.m !== undefined)
      ) {
        return '目标区域包含不能参与运算的非数值单元格。';
      }
    }
  }
  return null;
}
