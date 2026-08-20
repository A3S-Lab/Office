import type { Cell, Sheet } from '@fortune-sheet/core';
import type { CellObject, WorkSheet } from 'xlsx';
import {
  sparseArrayEntries,
  sparseMatrixColumnCount,
} from './spreadsheet-sparse';
import { exportXlsxCellComment } from './work-spreadsheet-comments';
import { refreshSpreadsheetPivotTables } from './work-spreadsheet-pivots';
import type { WorkArtifact, WorkSpreadsheetContent } from './work-types';
import { exportXlsxDefinedNames } from './work-xlsx-defined-names';
import {
  createXlsxErrorCell,
  createXlsxFormulaCell,
  patchXlsxFormulaFeatures,
} from './work-xlsx-formulas';
import { patchXlsxWorksheetDrawings } from './work-xlsx-images';
import { patchXlsxSheetFeatures } from './work-xlsx-interop';
import { patchXlsxPivotTables } from './work-xlsx-pivots';

export async function createWorkSpreadsheetBlob(
  artifact: WorkArtifact,
): Promise<Blob> {
  if (artifact.content.type !== 'spreadsheet')
    throw new Error('当前文件不是表格。');
  const content = refreshSpreadsheetPivotTables(artifact.content);
  const XLSX = await import('xlsx');
  const workbook = XLSX.utils.book_new();
  for (const sheet of content.sheets) {
    const worksheet = createSparseXlsxWorksheet(sheet, XLSX);
    applySpreadsheetLayout(worksheet, sheet);
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      sheet.name.slice(0, 31) || '工作表',
    );
  }
  workbook.Workbook = {
    ...workbook.Workbook,
    Sheets: content.sheets.map((sheet) => ({
      name: sheet.name.slice(0, 31) || '工作表',
      Hidden: sheet.hide ? 1 : 0,
    })),
    Names: exportXlsxDefinedNames(content),
  };
  const bytes = XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    compression: true,
  }) as ArrayBuffer;
  const withFeatures = await patchXlsxSheetFeatures(bytes, content);
  const withDrawings = await patchXlsxWorksheetDrawings(withFeatures, content);
  const withFormulas = await patchXlsxFormulaFeatures(withDrawings, content);
  const output = await patchXlsxPivotTables(withFormulas, content);
  return new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

type SparseXlsxWorksheet = WorkSheet & { '!a3sSparse'?: true };

function createSparseXlsxWorksheet(
  sheet: WorkSpreadsheetContent['sheets'][number],
  XLSX: typeof import('xlsx'),
): SparseXlsxWorksheet {
  const data = sheet.data ?? [];
  const worksheet: SparseXlsxWorksheet = { '!a3sSparse': true };
  let rowCount = Math.max(1, sheet.row ?? 0, data.length);
  let columnCount = Math.max(
    1,
    sheet.column ?? 0,
    sparseMatrixColumnCount(data),
  );
  for (const [rowIndex, row] of sparseArrayEntries(data)) {
    rowCount = Math.max(rowCount, rowIndex + 1);
    for (const [columnIndex, cell] of sparseArrayEntries(row)) {
      if (!cell) continue;
      columnCount = Math.max(columnCount, columnIndex + 1);
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const style = xlsxCellStyle(cell);
      const comment = exportXlsxCellComment(cell.ps);
      const hyperlink = sheet.hyperlink?.[`${rowIndex}_${columnIndex}`];
      const target =
        xlsxCellObject(cell, rowIndex, columnIndex, sheet) ??
        (style || comment || hyperlink ? { t: 's', v: '' } : null);
      if (!target) continue;
      if (cell.ct?.fa) target.z = cell.ct.fa;
      if (style) target.s = style;
      if (comment) target.c = comment;
      if (hyperlink) target.l = { Target: xlsxHyperlinkTarget(hyperlink) };
      worksheet[address] = target;
    }
  }
  worksheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rowCount - 1, c: columnCount - 1 },
  });
  return worksheet;
}

function xlsxCellObject(
  cell: Cell,
  row: number,
  column: number,
  sheet: WorkSpreadsheetContent['sheets'][number],
): CellObject | null {
  if (cell.f)
    return createXlsxFormulaCell(cell, row, column, sheet) as CellObject;
  if (cell.ct?.t === 'e') return createXlsxErrorCell(cell) as CellObject;
  const value = cell.v ?? cell.m;
  if (value === undefined || value === null) {
    return cell.ps ? { t: 's', v: '' } : null;
  }
  if (typeof value === 'boolean') return { t: 'b', v: value };
  if (typeof value === 'number') return { t: 'n', v: value };
  return { t: 's', v: String(value) };
}

function applySpreadsheetLayout(worksheet: WorkSheet, sheet: Sheet) {
  const merges = Object.values(sheet.config?.merge ?? {});
  if (merges.length) {
    worksheet['!merges'] = merges.map((merge) => ({
      s: { r: merge.r, c: merge.c },
      e: { r: merge.r + merge.rs - 1, c: merge.c + merge.cs - 1 },
    }));
  }
  const columnIndexes = new Set([
    ...Object.keys(sheet.config?.columnlen ?? {}),
    ...Object.keys(sheet.config?.colhidden ?? {}),
  ]);
  if (columnIndexes.size) {
    worksheet['!cols'] = [];
    for (const value of columnIndexes) {
      const index = Number(value);
      worksheet['!cols'][index] = {
        wpx: sheet.config?.columnlen?.[value],
        hidden: value in (sheet.config?.colhidden ?? {}),
      };
    }
  }
  const rowIndexes = new Set([
    ...Object.keys(sheet.config?.rowlen ?? {}),
    ...Object.keys(sheet.config?.rowhidden ?? {}),
  ]);
  if (rowIndexes.size) {
    worksheet['!rows'] = [];
    for (const value of rowIndexes) {
      const index = Number(value);
      worksheet['!rows'][index] = {
        hpx: sheet.config?.rowlen?.[value],
        hidden: value in (sheet.config?.rowhidden ?? {}),
      };
    }
  }
  if (sheet.filter_select) {
    worksheet['!autofilter'] = {
      ref: encodeSpreadsheetRange(sheet.filter_select),
    };
  }
}

function xlsxHyperlinkTarget(
  link: NonNullable<Sheet['hyperlink']>[string],
): string {
  if (link.linkType === 'webpage') return link.linkAddress;
  if (link.linkType === 'sheet')
    return `#${quoteSheetName(link.linkAddress)}!A1`;
  return `#${link.linkAddress}`;
}

function quoteSheetName(value: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(value)
    ? value
    : `'${value.replaceAll("'", "''")}'`;
}

function encodeSpreadsheetRange(
  range: NonNullable<Sheet['filter_select']>,
): string {
  return `${encodeSpreadsheetCell(range.row[0], range.column[0])}:${encodeSpreadsheetCell(
    range.row[1],
    range.column[1],
  )}`;
}

function encodeSpreadsheetCell(row: number, column: number): string {
  let value = Math.max(0, column) + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${Math.max(0, row) + 1}`;
}

function xlsxCellStyle(cell: Cell): CellObject['s'] | undefined {
  const color = (value: string | undefined) =>
    value ? { rgb: `FF${value.replace('#', '').toUpperCase()}` } : undefined;
  if (
    !cell.bl &&
    !cell.it &&
    !cell.un &&
    !cell.ff &&
    !cell.fs &&
    !cell.fc &&
    !cell.bg &&
    cell.ht === undefined &&
    cell.vt === undefined &&
    !cell.tb
  ) {
    return undefined;
  }
  return {
    font: {
      bold: Boolean(cell.bl),
      italic: Boolean(cell.it),
      strike: Boolean(cell.cl),
      underline: Boolean(cell.un),
      name: cell.ff,
      sz: cell.fs,
      color: color(cell.fc),
    },
    fill: cell.bg
      ? { patternType: 'solid', fgColor: color(cell.bg) }
      : undefined,
    alignment: {
      horizontal: cell.ht === 0 ? 'center' : cell.ht === 2 ? 'right' : 'left',
      vertical: cell.vt === 0 ? 'center' : cell.vt === 1 ? 'top' : 'bottom',
      wrapText: cell.tb === '2',
    },
  };
}
