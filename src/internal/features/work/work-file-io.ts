import type { Cell, CellMatrix, Sheet } from '@fortune-sheet/core';
import type { CellObject, WorkSheet } from 'xlsx';
import {
  sparseArrayEntries,
  sparseMatrixColumnCount,
} from './spreadsheet-sparse';
import {
  createWorkDocumentBlob,
  importWorkDocumentFile,
} from './work-document-file-io';
import {
  createWorkMarkdownBlob,
  importWorkMarkdownFile,
} from './work-markdown-file-io';
import {
  downloadBlob,
  fileNameWithoutExtension,
  safeFileName,
} from './work-file-download';
export { WORK_IMPORT_ACCEPT } from './work-file-contract';
import { materializeWorkFileSource } from './work-file-data';
import {
  WorkFileImportController,
  type WorkFileImportContext,
  type WorkFileImportOptions,
} from './work-file-import';
import { OoxmlPackage } from './work-ooxml-package';
import {
  createWorkPresentationBlob,
  importWorkPresentationFile,
  type WorkPresentationExportOptions,
} from './work-presentation-file-io';
import { readWorkSourceBlob, rememberWorkSourceBlob } from './work-repository';
import { createWorkArtifact, createWorkId } from './work-templates';
import {
  type WorkArtifact,
  type WorkArtifactKind,
  type WorkSpreadsheetContent,
  type WorkSpreadsheetDataValidationRange,
  workArtifactExtension,
} from './work-types';
import {
  exportXlsxCellComment,
  importXlsxCellComment,
} from './work-spreadsheet-comments';
import { importedSheetProtectionAuthority } from './work-spreadsheet-protection';
import { refreshSpreadsheetPivotTables } from './work-spreadsheet-pivots';
import { xlsxWorksheetChartsToSheet } from './work-xlsx-charts';
import {
  exportXlsxDefinedNames,
  importXlsxDefinedNames,
} from './work-xlsx-defined-names';
import {
  createSpreadsheetFormulaMetadata,
  createXlsxErrorCell,
  createXlsxFormulaCell,
  patchXlsxFormulaFeatures,
  readXlsxFormulaFeaturesFromPackage,
} from './work-xlsx-formulas';
import {
  patchXlsxWorksheetDrawings,
  xlsxWorksheetImagesToSheet,
} from './work-xlsx-images';
import {
  patchXlsxSheetFeatures,
  readXlsxSheetFeaturesFromPackage,
  type XlsxDataValidation,
} from './work-xlsx-interop';
import {
  applyImportedXlsxPivotTables,
  inspectXlsxPivotTables,
  patchXlsxPivotTables,
} from './work-xlsx-pivots';
import { editableSpreadsheetFormula } from './work-spreadsheet-formulas';

const DOCUMENT_EXTENSIONS = new Set(['docx', 'html', 'htm', 'txt']);
const MARKDOWN_EXTENSIONS = new Set(['md', 'markdown']);
const SPREADSHEET_EXTENSIONS = new Set(['xlsx', 'xls', 'csv', 'ods']);
const PRESENTATION_EXTENSIONS = new Set(['pptx']);
const PDF_EXTENSIONS = new Set(['pdf']);
const MAX_INLINE_DATA_VALIDATION_CELLS = 10_000;

export type WorkArtifactExportOptions = WorkPresentationExportOptions;

export async function importWorkFile(
  file: File,
  options: WorkFileImportOptions = {},
): Promise<WorkArtifact> {
  const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
  if (
    !MARKDOWN_EXTENSIONS.has(extension) &&
    !DOCUMENT_EXTENSIONS.has(extension) &&
    !SPREADSHEET_EXTENSIONS.has(extension) &&
    !PRESENTATION_EXTENSIONS.has(extension) &&
    !PDF_EXTENSIONS.has(extension)
  ) {
    throw new Error(
      '目前可导入 DOCX、XLSX、XLS、ODS、CSV、PPTX、PDF、HTML、Markdown 和文本文件。',
    );
  }
  const controller = new WorkFileImportController(options, file.size);
  const source = await materializeWorkFileSource(file, controller);
  const context: WorkFileImportContext = {
    bytes: source.bytes,
    controller,
  };
  controller.report('parsing', 0);
  let artifact: WorkArtifact;
  if (MARKDOWN_EXTENSIONS.has(extension)) {
    artifact = await importWorkMarkdownFile(source.file, context);
  } else if (DOCUMENT_EXTENSIONS.has(extension)) {
    artifact = await importWorkDocumentFile(source.file, extension, context);
  } else if (SPREADSHEET_EXTENSIONS.has(extension)) {
    artifact = await importSpreadsheet(source.file, extension, context);
  } else if (PRESENTATION_EXTENSIONS.has(extension)) {
    artifact = await importWorkPresentationFile(source.file, context);
  } else {
    artifact = await importPdf(source.file);
  }
  controller.report('analyzing', 1);
  controller.report('finalizing', 0);
  controller.complete();
  return artifact;
}

export async function exportWorkArtifact(
  artifact: WorkArtifact,
  options?: WorkArtifactExportOptions,
): Promise<void> {
  const blob = await createWorkArtifactBlob(artifact, options);
  downloadBlob(
    blob,
    `${safeFileName(artifact.title)}.${workArtifactExtension(artifact.kind)}`,
  );
}

export async function createWorkArtifactBlob(
  artifact: WorkArtifact,
  options?: WorkArtifactExportOptions,
): Promise<Blob> {
  if (artifact.kind === 'document') return createWorkDocumentBlob(artifact);
  if (artifact.kind === 'markdown') return createWorkMarkdownBlob(artifact);
  if (artifact.kind === 'spreadsheet') return createSpreadsheetBlob(artifact);
  if (artifact.kind === 'presentation')
    return createWorkPresentationBlob(artifact, options);
  return readWorkSourceBlob(artifact);
}

async function importSpreadsheet(
  file: File,
  extension: string,
  context: WorkFileImportContext,
): Promise<WorkArtifact> {
  const XLSX = await import('xlsx');
  context.controller.report('parsing', 0.05);
  const arrayBuffer = context.bytes;
  const workbook = XLSX.read(arrayBuffer, {
    type: 'array',
    cellDates: true,
    cellFormula: true,
    cellStyles: true,
    xlfn: true,
  });
  await context.controller.checkpoint('parsing', 0.2);
  const archive =
    extension === 'xlsx'
      ? await OoxmlPackage.load(arrayBuffer).catch(() => null)
      : null;
  const sheetFeatures = archive
    ? await readXlsxSheetFeaturesFromPackage(archive).catch(() => new Map())
    : new Map();
  const formulaFeatures = archive
    ? await readXlsxFormulaFeaturesFromPackage(archive).catch(() => null)
    : null;
  const pivotFeatures = archive
    ? await inspectXlsxPivotTables(archive).catch(() => null)
    : null;
  await context.controller.checkpoint('parsing', 0.4);
  const sheets: WorkSpreadsheetContent['sheets'] = [];
  for (const [index, name] of workbook.SheetNames.entries()) {
    const worksheet = workbook.Sheets[name];
    if (!worksheet) continue;
    const features = sheetFeatures.get(name);
    const range = safeSpreadsheetRange(worksheet, XLSX);
    let rowCount = Math.max(range.e.r + 1, 40);
    let columnCount = Math.max(range.e.c + 1, 12);
    const id = createWorkId('sheet');
    const hyperlinks: NonNullable<Sheet['hyperlink']> = {};
    const data: CellMatrix = [];
    const entries = Object.entries(worksheet).filter(
      ([address, value]) =>
        !address.startsWith('!') && Boolean(value) && typeof value === 'object',
    );
    const sheetProgressStart =
      0.4 + (index / Math.max(1, workbook.SheetNames.length)) * 0.45;
    const sheetProgressSize = 0.45 / Math.max(1, workbook.SheetNames.length);
    for (const [entryIndex, [address, value]] of entries.entries()) {
      let position: { r: number; c: number };
      try {
        position = XLSX.utils.decode_cell(address);
      } catch {
        continue;
      }
      if (position.r < 0 || position.c < 0) continue;
      const source = value as CellObject;
      const row = position.r;
      const column = position.c;
      rowCount = Math.max(rowCount, row + 1);
      columnCount = Math.max(columnCount, column + 1);
      const hyperlink = fortuneSheetHyperlink(source.l?.Target);
      const comment = importXlsxCellComment(source.c);
      const style = fortuneCellStyle(source);
      if (hyperlink) hyperlinks[`${row}_${column}`] = hyperlink;
      data[row] ??= [];
      data[row][column] = {
        v: source.v as Cell['v'],
        m: source.w ?? String(source.v ?? ''),
        f: source.f ? `=${editableSpreadsheetFormula(source.f)}` : undefined,
        ps: comment,
        hl: hyperlink ? { r: row, c: column, id } : undefined,
        ...style,
        fc: hyperlink ? (style.fc ?? '#0563c1') : style.fc,
        un: hyperlink ? (style.un ?? 1) : style.un,
      };
      if ((entryIndex + 1) % 2_048 === 0) {
        await context.controller.checkpoint(
          'parsing',
          sheetProgressStart +
            sheetProgressSize *
              ((entryIndex + 1) / Math.max(1, entries.length)),
        );
      }
    }
    data.length = Math.max(data.length, rowCount);
    const config = fortuneSheetConfig(worksheet);
    const protectionAuthority = importedSheetProtectionAuthority(
      features?.protection.authority,
      features?.protection.cellProtectionRanges ?? [],
    );
    if (protectionAuthority) config.authority = protectionAuthority;
    const filterSelect = fortuneSheetFilter(worksheet, XLSX);
    const dataVerification = fortuneSheetDataVerification(
      features?.validations ?? [],
      rowCount,
      columnCount,
      XLSX,
    );
    const dataValidationRanges = fortuneSheetDataValidationRanges(
      features?.validations ?? [],
      rowCount,
      columnCount,
      XLSX,
    );
    sheets.push({
      id,
      name,
      order: index,
      status: index === 0 ? 1 : 0,
      hide: workbook.Workbook?.Sheets?.[index]?.Hidden ? 1 : 0,
      row: rowCount,
      column: columnCount,
      data,
      config,
      filter: filterSelect ? {} : undefined,
      filter_select: filterSelect,
      frozen: features?.frozen,
      hyperlink: Object.keys(hyperlinks).length ? hyperlinks : undefined,
      dataVerification: Object.keys(dataVerification).length
        ? dataVerification
        : undefined,
      dataValidationRanges: dataValidationRanges.length
        ? dataValidationRanges
        : undefined,
      luckysheet_conditionformat_save: features?.conditionalFormats.length
        ? features.conditionalFormats
        : undefined,
      images: features?.images.length
        ? xlsxWorksheetImagesToSheet(features.images, config)
        : undefined,
      charts: features?.charts.length
        ? xlsxWorksheetChartsToSheet(features.charts, config)
        : undefined,
      formulaMetadata: createSpreadsheetFormulaMetadata(
        worksheet,
        formulaFeatures?.sheets.get(name),
      ),
    });
    await context.controller.checkpoint(
      'parsing',
      sheetProgressStart + sheetProgressSize,
    );
  }
  const workbookMetadata = importXlsxDefinedNames(workbook, sheets);
  const pageBreaks = sheets.flatMap((sheet) => {
    const imported = sheetFeatures.get(sheet.name)?.pageBreaks;
    if (!sheet.id || (!imported?.rows.length && !imported?.columns.length))
      return [];
    return [
      {
        sheetId: sheet.id,
        rows: imported.rows.length ? imported.rows : undefined,
        columns: imported.columns.length ? imported.columns : undefined,
      },
    ];
  });
  const pageSetups = sheets.flatMap((sheet) => {
    const imported = sheetFeatures.get(sheet.name)?.pageSetup;
    return sheet.id && imported ? [{ sheetId: sheet.id, ...imported }] : [];
  });
  const artifact = createWorkArtifact('blank-spreadsheet');
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = applyImportedXlsxPivotTables(
    {
      type: 'spreadsheet',
      sheets,
      calculation: formulaFeatures?.calculation,
      ...workbookMetadata,
      pageBreaks: pageBreaks.length ? pageBreaks : undefined,
      pageSetups: pageSetups.length ? pageSetups : undefined,
    },
    pivotFeatures,
  );
  const { analyzeSpreadsheetCompatibility } = await import(
    './work-office-diagnostics'
  );
  context.controller.report('analyzing', 0);
  artifact.compatibility = await analyzeSpreadsheetCompatibility(
    file,
    extension,
    workbook,
    archive,
  );
  context.controller.report('analyzing', 1);
  return artifact;
}

async function importPdf(file: File): Promise<WorkArtifact> {
  const contentType = file.type || 'application/pdf';
  const source = new Blob([file], { type: contentType });
  const artifact = createWorkArtifact('blank-document');
  artifact.kind = 'pdf';
  artifact.title = fileNameWithoutExtension(file.name);
  artifact.content = { type: 'pdf' };
  artifact.source = {
    name: file.name,
    contentType,
    size: file.size,
    updatedAt: file.lastModified || Date.now(),
  };
  rememberWorkSourceBlob(artifact.id, source);
  return artifact;
}

function safeSpreadsheetRange(
  worksheet: WorkSheet,
  XLSX: typeof import('xlsx'),
): ReturnType<typeof XLSX.utils.decode_range> {
  if (worksheet['!ref']) {
    try {
      return XLSX.utils.decode_range(worksheet['!ref']);
    } catch {
      // Fall back to the minimum editable worksheet dimensions below.
    }
  }
  return { s: { r: 0, c: 0 }, e: { r: 39, c: 11 } };
}

function fortuneSheetConfig(
  worksheet: WorkSheet,
): NonNullable<Sheet['config']> {
  const config: NonNullable<Sheet['config']> = {};
  for (const range of worksheet['!merges'] ?? []) {
    config.merge ??= {};
    config.merge[`${range.s.r}_${range.s.c}`] = {
      r: range.s.r,
      c: range.s.c,
      rs: range.e.r - range.s.r + 1,
      cs: range.e.c - range.s.c + 1,
    };
  }
  for (const [index, column] of (worksheet['!cols'] ?? []).entries()) {
    if (!column) continue;
    if (column.wpx || column.wch) {
      config.columnlen ??= {};
      config.columnlen[index] = Math.round(
        column.wpx ?? (column.wch ?? 8.43) * 8,
      );
    }
    if (column.hidden) {
      config.colhidden ??= {};
      config.colhidden[index] = 0;
    }
  }
  for (const [index, row] of (worksheet['!rows'] ?? []).entries()) {
    if (!row) continue;
    if (row.hpx || row.hpt) {
      config.rowlen ??= {};
      config.rowlen[index] = Math.round(row.hpx ?? ((row.hpt ?? 15) * 96) / 72);
    }
    if (row.hidden) {
      config.rowhidden ??= {};
      config.rowhidden[index] = 0;
    }
  }
  return config;
}

function fortuneCellStyle(source: CellObject): Partial<Cell> {
  const style = source.s;
  if (!style || typeof style !== 'object') {
    return source.z || source.t === 'e'
      ? { ct: { fa: source.z ? String(source.z) : undefined, t: source.t } }
      : {};
  }
  const font = style.font as
    | {
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        name?: string;
        sz?: number;
        color?: { rgb?: string };
      }
    | undefined;
  const fill = style.fill as { fgColor?: { rgb?: string } } | undefined;
  const alignment = style.alignment as
    | {
        horizontal?: string;
        vertical?: string;
        wrapText?: boolean;
        textRotation?: number;
      }
    | undefined;
  return {
    bl: font?.bold ? 1 : undefined,
    it: font?.italic ? 1 : undefined,
    un: font?.underline ? 1 : undefined,
    ff: font?.name,
    fs: font?.sz,
    fc: spreadsheetColor(font?.color?.rgb),
    bg: spreadsheetColor(fill?.fgColor?.rgb),
    ht:
      alignment?.horizontal === 'center'
        ? 0
        : alignment?.horizontal === 'right'
          ? 2
          : 1,
    vt:
      alignment?.vertical === 'center'
        ? 0
        : alignment?.vertical === 'top'
          ? 1
          : 2,
    tb: alignment?.wrapText ? '2' : undefined,
    tr: alignment?.textRotation ? String(alignment.textRotation) : undefined,
    ct:
      source.z || source.t === 'e'
        ? { fa: source.z ? String(source.z) : undefined, t: source.t }
        : undefined,
  };
}

async function createSpreadsheetBlob(artifact: WorkArtifact): Promise<Blob> {
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

function fortuneSheetFilter(
  worksheet: WorkSheet,
  XLSX: typeof import('xlsx'),
): Sheet['filter_select'] | undefined {
  const reference = worksheet['!autofilter']?.ref;
  if (!reference) return undefined;
  try {
    const range = XLSX.utils.decode_range(reference);
    return {
      row: [range.s.r, range.e.r],
      column: [range.s.c, range.e.c],
    };
  } catch {
    return undefined;
  }
}

function fortuneSheetHyperlink(
  target: string | undefined,
): NonNullable<Sheet['hyperlink']>[string] | undefined {
  if (!target) return undefined;
  if (!target.startsWith('#')) {
    return { linkType: 'webpage', linkAddress: target };
  }
  const address = target.slice(1);
  return {
    linkType: address.includes('!') ? 'cellrange' : 'sheet',
    linkAddress: address,
  };
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

function fortuneSheetDataVerification(
  validations: XlsxDataValidation[],
  rowCount: number,
  columnCount: number,
  XLSX: typeof import('xlsx'),
): Record<string, XlsxDataValidation['item']> {
  const result: Record<string, XlsxDataValidation['item']> = {};
  let remaining = MAX_INLINE_DATA_VALIDATION_CELLS;
  for (const validation of validations) {
    for (const reference of validation.references) {
      try {
        const range = XLSX.utils.decode_range(reference);
        const lastRow = Math.min(range.e.r, rowCount - 1);
        const lastColumn = Math.min(range.e.c, columnCount - 1);
        const cells =
          (lastRow - Math.max(0, range.s.r) + 1) *
          (lastColumn - Math.max(0, range.s.c) + 1);
        if (cells <= 0 || cells > remaining) continue;
        remaining -= cells;
        for (let row = Math.max(0, range.s.r); row <= lastRow; row += 1) {
          for (
            let column = Math.max(0, range.s.c);
            column <= lastColumn;
            column += 1
          ) {
            result[`${row}_${column}`] = {
              ...validation.item,
              rangeTxt: reference,
            };
          }
        }
      } catch {
        // Keep the rest of the worksheet editable when one validation reference is malformed.
      }
    }
  }
  return result;
}

function fortuneSheetDataValidationRanges(
  validations: XlsxDataValidation[],
  rowCount: number,
  columnCount: number,
  XLSX: typeof import('xlsx'),
): WorkSpreadsheetDataValidationRange[] {
  return validations.flatMap((validation) => {
    const ranges = validation.references.flatMap((reference) => {
      try {
        const range = XLSX.utils.decode_range(reference);
        const startRow = Math.max(0, range.s.r);
        const endRow = Math.min(rowCount - 1, range.e.r);
        const startColumn = Math.max(0, range.s.c);
        const endColumn = Math.min(columnCount - 1, range.e.c);
        return startRow <= endRow && startColumn <= endColumn
          ? [
              {
                row: [startRow, endRow] as [number, number],
                column: [startColumn, endColumn] as [number, number],
              },
            ]
          : [];
      } catch {
        return [];
      }
    });
    return ranges.length
      ? [
          {
            ranges,
            item: { ...validation.item },
          },
        ]
      : [];
  });
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

function spreadsheetColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const normalized = value.replace(/^#/, '').slice(-6);
  return /^[0-9a-f]{6}$/i.test(normalized)
    ? `#${normalized.toLowerCase()}`
    : undefined;
}

export function workKindForFile(file: File): WorkArtifactKind | null {
  const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
  if (MARKDOWN_EXTENSIONS.has(extension)) return 'markdown';
  if (DOCUMENT_EXTENSIONS.has(extension)) return 'document';
  if (SPREADSHEET_EXTENSIONS.has(extension)) return 'spreadsheet';
  if (PRESENTATION_EXTENSIONS.has(extension)) return 'presentation';
  if (PDF_EXTENSIONS.has(extension)) return 'pdf';
  return null;
}
