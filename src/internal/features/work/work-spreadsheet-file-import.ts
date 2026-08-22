import type { Cell, CellMatrix, Sheet } from '@fortune-sheet/core';
import type { CellObject, WorkSheet } from 'xlsx';
import { fileNameWithoutExtension } from './work-file-download';
import type { WorkFileImportContext } from './work-file-import';
import { OoxmlPackage } from './work-ooxml-package';
import { importXlsxCellComment } from './work-spreadsheet-comments';
import {
  emptySpreadsheetWorksheetCompatibilitySummary,
  type SpreadsheetWorksheetCompatibilitySummary,
  updateSpreadsheetWorksheetCompatibilitySummary,
} from './work-spreadsheet-compatibility-summary';
import { editableSpreadsheetFormula } from './work-spreadsheet-formulas';
import {
  recordSpreadsheetImportMeasure,
  spreadsheetImportNow,
} from './work-spreadsheet-import-diagnostics';
import { readSpreadsheetWorkbookInWorker } from './work-spreadsheet-import-worker-client';
import {
  freezeImportedSpreadsheetCell,
  registerImportedSpreadsheetMatrix,
} from './work-spreadsheet-matrix-profile';
import { scanSpreadsheetPackageInWorker } from './work-spreadsheet-package-scan-worker-client';
import { importedSheetProtectionAuthority } from './work-spreadsheet-protection';
import { createWorkArtifact, createWorkId } from './work-templates';
import type {
  WorkArtifact,
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationRange,
} from './work-types';
import { xlsxWorksheetChartsToSheet } from './work-xlsx-charts';
import { fortuneBorderInfoFromXlsxCells } from './work-xlsx-cell-borders';
import { withXlsxCellStyleOrigin } from './work-xlsx-cell-style-origin';
import { importXlsxDefinedNames } from './work-xlsx-defined-names';
import {
  createSpreadsheetFormulaMetadata,
  readXlsxFormulaFeaturesFromPackage,
} from './work-xlsx-formulas';
import { xlsxWorksheetImagesToSheet } from './work-xlsx-images';
import { spreadsheetUnderlineCellValueFromSheetJs } from './work-spreadsheet-underline';
import {
  spreadsheetExplicitTextOrientationFromCell,
  spreadsheetTextOrientationCellStyle,
  spreadsheetTextOrientationFromXlsx,
} from './work-spreadsheet-text-orientation';
import {
  readXlsxSheetFeaturesFromPackage,
  type XlsxDataValidation,
  type XlsxSheetFeatures,
} from './work-xlsx-interop';
import {
  applyImportedXlsxPivotTables,
  inspectXlsxPivotTables,
} from './work-xlsx-pivots';
import { xlsxWorksheetRequiresSheetJsCellStyles } from './work-xlsx-style-gate';
import { xlsxWorksheetCellEntries } from './work-xlsx-worksheet';
import { applyImportedXlsxRichText } from './work-xlsx-rich-text';

const MAX_INLINE_DATA_VALIDATION_CELLS = 10_000;

export async function importWorkSpreadsheetFile(
  file: File,
  extension: string,
  context: WorkFileImportContext,
): Promise<WorkArtifact> {
  const importStartedAt = spreadsheetImportNow();
  context.controller.report('parsing', 0.05);
  const arrayBuffer = context.bytes;
  const packageStartedAt = spreadsheetImportNow();
  const archive =
    extension === 'xlsx'
      ? await OoxmlPackage.load(arrayBuffer).catch(() => null)
      : null;
  recordSpreadsheetImportMeasure(
    'a3s-office.spreadsheet.package',
    packageStartedAt,
    spreadsheetImportNow(),
    { loaded: archive !== null },
  );
  const initialRequiresCellStyles = archive
    ? archive.has('xl/styles.xml')
    : true;
  let readOptions = {
    type: 'array',
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellStyles: initialRequiresCellStyles,
    cellText: false,
    dense: true,
    xlfn: true,
  } satisfies import('xlsx').ParsingOptions;
  const speculativeWorkerAbort = new AbortController();
  let speculativeWorkerCancelledForFastPath = false;
  const speculativeWorkerSignal = context.controller.signal
    ? AbortSignal.any([
        context.controller.signal,
        speculativeWorkerAbort.signal,
      ])
    : speculativeWorkerAbort.signal;
  let workerWorkbook = readSpreadsheetWorkbookInWorker(
    arrayBuffer.slice(0),
    readOptions,
    speculativeWorkerSignal,
  ).catch((error: unknown) => {
    if (speculativeWorkerAbort.signal.aborted) return null;
    throw error;
  });
  const styleGateStartedAt = spreadsheetImportNow();
  const packageScan = archive
    ? scanSpreadsheetPackageInWorker(
        arrayBuffer.slice(0),
        context.controller.signal,
        () => {
          if (speculativeWorkerAbort.signal.aborted) return;
          speculativeWorkerCancelledForFastPath = true;
          speculativeWorkerAbort.abort();
        },
      )
    : Promise.resolve(null);
  const packageScanResult = await packageScan;
  let usesPlainXlsxFastPath = false;
  const requiresCellStyles = archive
    ? initialRequiresCellStyles ||
      (packageScanResult
        ? Object.values(packageScanResult.worksheets).some(
            (scan) => scan.requiresSheetJsCellStyles,
          )
        : await xlsxPackageRequiresSheetJsCellStyles(archive))
    : true;
  recordSpreadsheetImportMeasure(
    'a3s-office.spreadsheet.style-gate',
    styleGateStartedAt,
    spreadsheetImportNow(),
    { required: requiresCellStyles, worker: packageScanResult !== null },
  );
  if (context.controller.signal?.aborted) {
    speculativeWorkerAbort.abort();
    await workerWorkbook;
    context.controller.throwIfAborted();
  }
  if (packageScanResult?.workbook && !requiresCellStyles) {
    speculativeWorkerAbort.abort();
    await workerWorkbook;
    workerWorkbook = Promise.resolve(packageScanResult.workbook);
    usesPlainXlsxFastPath = true;
    recordSpreadsheetImportMeasure(
      'a3s-office.spreadsheet.plain-xlsx-fast-path',
      styleGateStartedAt,
      spreadsheetImportNow(),
      { sheetCount: packageScanResult.workbook.SheetNames.length },
    );
  } else if (
    speculativeWorkerCancelledForFastPath ||
    requiresCellStyles !== initialRequiresCellStyles
  ) {
    speculativeWorkerAbort.abort();
    await workerWorkbook;
    readOptions = { ...readOptions, cellStyles: requiresCellStyles };
    workerWorkbook = readSpreadsheetWorkbookInWorker(
      arrayBuffer.slice(0),
      readOptions,
      context.controller.signal,
    );
  }
  let sheetJsModule: Promise<typeof import('xlsx')> | null = null;
  const loadSheetJs = () => {
    if (sheetJsModule) return sheetJsModule;
    const sheetJsStartedAt = spreadsheetImportNow();
    sheetJsModule = import('xlsx').then((module) => {
      recordSpreadsheetImportMeasure(
        'a3s-office.spreadsheet.sheetjs-module',
        sheetJsStartedAt,
        spreadsheetImportNow(),
      );
      return module;
    });
    return sheetJsModule;
  };
  if (!usesPlainXlsxFastPath) loadSheetJs();
  const featuresStartedAt = spreadsheetImportNow();
  const spreadsheetFeatures = Promise.all([
    archive
      ? readXlsxSheetFeaturesFromPackage(
          archive,
          packageScanResult?.worksheets,
        ).catch(() => new Map<string, XlsxSheetFeatures>())
      : Promise.resolve(new Map<string, XlsxSheetFeatures>()),
    archive
      ? readXlsxFormulaFeaturesFromPackage(
          archive,
          packageScanResult?.worksheets,
        ).catch(() => null)
      : Promise.resolve(null),
    archive
      ? inspectXlsxPivotTables(archive).catch(() => null)
      : Promise.resolve(null),
  ]).then((features) => {
    recordSpreadsheetImportMeasure(
      'a3s-office.spreadsheet.features',
      featuresStartedAt,
      spreadsheetImportNow(),
    );
    return features;
  });
  const diagnosticsModule = import('./work-office-diagnostics');
  const workbookResult = workerWorkbook.then(async (workerResult) => {
    if (workerResult) return workerResult;
    return (await loadSheetJs()).read(arrayBuffer, readOptions);
  });
  const [workbook, importedFeatures] = await Promise.all([
    workbookResult,
    spreadsheetFeatures,
  ]);
  const XLSX = usesPlainXlsxFastPath ? null : await loadSheetJs();
  const [sheetFeatures, formulaFeatures, pivotFeatures] = importedFeatures;
  await context.controller.checkpoint('parsing', 0.2);
  await context.controller.checkpoint('parsing', 0.4);
  const conversionStartedAt = spreadsheetImportNow();
  const sheets: WorkSpreadsheetContent['sheets'] = [];
  let populatedCellCount = 0;
  const worksheetCompatibility = new Map<
    string,
    SpreadsheetWorksheetCompatibilitySummary
  >();
  for (const [index, name] of workbook.SheetNames.entries()) {
    const worksheet = workbook.Sheets[name];
    if (!worksheet) continue;
    const plainWorksheet = packageScanResult?.plainWorksheets?.[name];
    const features = sheetFeatures.get(name);
    const directCellStyles = new Map(
      (features?.directCellStyles ?? []).map((entry) => [
        spreadsheetCellKey(entry.row, entry.column),
        entry,
      ]),
    );
    const richTextCells = new Map(
      (features?.richTextCells ?? []).map((entry) => [
        spreadsheetCellKey(entry.row, entry.column),
        entry,
      ]),
    );
    const range = safeSpreadsheetRange(worksheet, XLSX);
    let rowCount = Math.max(plainWorksheet?.rowCount ?? range.e.r + 1, 40);
    let columnCount = Math.max(
      plainWorksheet?.columnCount ?? range.e.c + 1,
      12,
    );
    const id = context.spreadsheetSheetIds?.[index] ?? createWorkId('sheet');
    const hyperlinks: NonNullable<Sheet['hyperlink']> = {};
    const data: CellMatrix = plainWorksheet?.data ?? [];
    const formulaCells: Array<{ column: number; row: number }> = [];
    const shownCommentCells: Array<{ c: number; r: number }> = [];
    const compatibilitySummary =
      emptySpreadsheetWorksheetCompatibilitySummary();
    const importedFormulaFeatures = formulaFeatures?.sheets.get(name);
    const formulaMetadata =
      importedFormulaFeatures?.formulas.length === 0
        ? undefined
        : createSpreadsheetFormulaMetadata(worksheet, importedFormulaFeatures);
    const estimatedCellCount = Math.max(
      1,
      (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1),
    );
    const sheetProgressStart =
      0.4 + (index / Math.max(1, workbook.SheetNames.length)) * 0.45;
    const sheetProgressSize = 0.45 / Math.max(1, workbook.SheetNames.length);
    let entryIndex = plainWorksheet?.populatedCellCount ?? 0;
    if (!plainWorksheet) {
      if (!XLSX) {
        throw new Error('The spreadsheet compatibility parser is unavailable.');
      }
      for (const { cell: source, column, row } of xlsxWorksheetCellEntries(
        worksheet,
      )) {
        entryIndex += 1;
        rowCount = Math.max(rowCount, row + 1);
        columnCount = Math.max(columnCount, column + 1);
        const hyperlink = fortuneSheetHyperlink(source.l?.Target);
        const comment = importXlsxCellComment(source.c);
        updateSpreadsheetWorksheetCompatibilitySummary(
          compatibilitySummary,
          source,
        );
        if (hyperlink) hyperlinks[`${row}_${column}`] = hyperlink;
        const directCellStyle = directCellStyles.get(
          spreadsheetCellKey(row, column),
        );
        directCellStyles.delete(spreadsheetCellKey(row, column));
        data[row] ??= [];
        data[row][column] = freezeImportedSpreadsheetCell(
          withXlsxCellStyleOrigin(
            applyImportedXlsxRichText(
              fortuneCellFromXlsx(
                source,
                row,
                column,
                id,
                hyperlink,
                comment,
                XLSX,
                directCellStyle?.style,
              ),
              richTextCells.get(spreadsheetCellKey(row, column)),
            ),
            directCellStyle?.origin,
          ),
        );
        if (source.f) formulaCells.push({ column, row });
        if (comment?.isShow) shownCommentCells.push({ c: column, r: row });
        if (entryIndex % 2_048 === 0) {
          await context.controller.checkpoint(
            'parsing',
            sheetProgressStart +
              sheetProgressSize * Math.min(1, entryIndex / estimatedCellCount),
          );
        }
      }
      releaseImportedWorksheetCells(worksheet);
    } else {
      await context.controller.checkpoint(
        'parsing',
        sheetProgressStart + sheetProgressSize * 0.95,
      );
    }
    for (const { column, origin, row, style } of directCellStyles.values()) {
      rowCount = Math.max(rowCount, row + 1);
      columnCount = Math.max(columnCount, column + 1);
      data[row] ??= [];
      const existing = data[row][column];
      data[row][column] = freezeImportedSpreadsheetCell(
        withXlsxCellStyleOrigin(
          {
            ...(existing ?? {}),
            ...style,
          },
          origin,
        ),
      );
      if (!existing) entryIndex += 1;
    }
    populatedCellCount += entryIndex;
    worksheetCompatibility.set(name, compatibilitySummary);
    data.length = Math.max(data.length, rowCount);
    const config = fortuneSheetConfig(worksheet);
    const importedBorderInfo = fortuneBorderInfoFromXlsxCells(
      features?.directCellStyles ?? [],
    );
    if (importedBorderInfo.length) {
      config.borderInfo = [
        ...(Array.isArray(config.borderInfo) ? config.borderInfo : []),
        ...importedBorderInfo,
      ];
    }
    registerImportedSpreadsheetMatrix(data, {
      columnCount,
      formulaCells,
      fortuneReady: Object.keys(config.merge ?? {}).length === 0,
      populatedCellCount: entryIndex,
      protectionCellKey: '',
      rowCount,
      shownCommentCells,
    });
    const protectionAuthority = importedSheetProtectionAuthority(
      features?.protection.authority,
      features?.protection.cellProtectionRanges ?? [],
    );
    if (protectionAuthority) config.authority = protectionAuthority;
    const filterSelect = XLSX ? fortuneSheetFilter(worksheet, XLSX) : undefined;
    const dataVerification = XLSX
      ? fortuneSheetDataVerification(
          features?.validations ?? [],
          rowCount,
          columnCount,
          XLSX,
        )
      : {};
    const dataValidationRanges = XLSX
      ? fortuneSheetDataValidationRanges(
          features?.validations ?? [],
          rowCount,
          columnCount,
          XLSX,
        )
      : [];
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
      tables: features?.tables.length ? features.tables : undefined,
      formulaMetadata,
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
  recordSpreadsheetImportMeasure(
    'a3s-office.spreadsheet.canonical-conversion',
    conversionStartedAt,
    spreadsheetImportNow(),
    {
      populatedCellCount,
      sheetCount: sheets.length,
    },
  );
  const { analyzeSpreadsheetCompatibility } = await diagnosticsModule;
  context.controller.report('analyzing', 0);
  const compatibilityStartedAt = spreadsheetImportNow();
  artifact.compatibility = await analyzeSpreadsheetCompatibility(
    file,
    extension,
    workbook,
    archive,
    worksheetCompatibility,
    {
      formulaFeatures,
      pivotFeatures,
      worksheetScans: packageScanResult?.worksheets,
    },
  );
  recordSpreadsheetImportMeasure(
    'a3s-office.spreadsheet.compatibility',
    compatibilityStartedAt,
    spreadsheetImportNow(),
  );
  context.controller.report('analyzing', 1);
  recordSpreadsheetImportMeasure(
    'a3s-office.spreadsheet.import-total',
    importStartedAt,
    spreadsheetImportNow(),
    { sheetCount: sheets.length },
  );
  return artifact;
}
function safeSpreadsheetRange(
  worksheet: WorkSheet,
  XLSX: typeof import('xlsx') | null,
): SpreadsheetRange {
  if (worksheet['!ref']) {
    if (!XLSX) {
      const range = decodeSpreadsheetRange(worksheet['!ref']);
      if (range) return range;
    }
    try {
      if (XLSX) return XLSX.utils.decode_range(worksheet['!ref']);
    } catch {
      // Fall back to the minimum editable worksheet dimensions below.
    }
  }
  return { s: { r: 0, c: 0 }, e: { r: 39, c: 11 } };
}

function decodeSpreadsheetRange(reference: string): SpreadsheetRange | null {
  const separator = reference.indexOf(':');
  if (separator >= 0 && reference.indexOf(':', separator + 1) >= 0) return null;
  const start = decodeSpreadsheetCell(
    separator < 0 ? reference : reference.slice(0, separator),
  );
  const end = decodeSpreadsheetCell(
    separator < 0 ? reference : reference.slice(separator + 1),
  );
  return start && end && start.r <= end.r && start.c <= end.c
    ? { s: start, e: end }
    : null;
}

interface SpreadsheetRange {
  e: { c: number; r: number };
  s: { c: number; r: number };
}

function decodeSpreadsheetCell(
  reference: string,
): { c: number; r: number } | null {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(reference);
  if (!match) return null;
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  const row = Number(match[2]);
  return column <= 16_384 && row <= 1_048_576
    ? { c: column - 1, r: row - 1 }
    : null;
}

async function xlsxPackageRequiresSheetJsCellStyles(
  archive: OoxmlPackage,
): Promise<boolean> {
  if (archive.has('xl/styles.xml')) return true;
  for (const partPath of archive.paths('xl/worksheets/')) {
    if (!partPath.endsWith('.xml')) continue;
    const source = await archive.text(partPath);
    if (xlsxWorksheetRequiresSheetJsCellStyles(source)) {
      return true;
    }
  }
  return false;
}

function releaseImportedWorksheetCells(worksheet: WorkSheet): void {
  if (Array.isArray(worksheet)) {
    worksheet.length = 0;
    return;
  }
  for (const address in worksheet) {
    if (!address.startsWith('!')) Reflect.deleteProperty(worksheet, address);
  }
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

/*
 * SheetJS attaches a default style object and `z: "General"` to every cell
 * when `cellStyles` is enabled. Keeping those defaults in the controlled model
 * multiplies memory and clone cost without changing what Fortune Sheet draws.
 */
function fortuneCellFromXlsx(
  source: CellObject,
  row: number,
  column: number,
  sheetId: string,
  hyperlink: ReturnType<typeof fortuneSheetHyperlink>,
  comment: ReturnType<typeof importXlsxCellComment>,
  XLSX: typeof import('xlsx'),
  directStyle?: Partial<Cell>,
): Cell {
  const cell: Cell = { ...fortuneCellStyle(source), ...directStyle };
  const orientation = spreadsheetExplicitTextOrientationFromCell(cell);
  if (orientation) {
    const orientationStyle = spreadsheetTextOrientationCellStyle(orientation);
    delete cell.rt;
    delete cell.tr;
    if (orientationStyle) Object.assign(cell, orientationStyle);
  }
  if (source.v !== undefined) cell.v = source.v as Cell['v'];
  const displayText = fortuneCellDisplayText(source, XLSX);
  if (displayText !== undefined) cell.m = displayText;
  if (source.f) cell.f = `=${editableSpreadsheetFormula(source.f)}`;
  if (comment) cell.ps = comment;
  if (hyperlink) {
    cell.hl = { r: row, c: column, id: sheetId };
    cell.fc ??= '#0563c1';
    cell.un ??= 1;
  }
  return cell;
}

function spreadsheetCellKey(row: number, column: number): string {
  return `${row}_${column}`;
}

function fortuneCellStyle(source: CellObject): Partial<Cell> {
  const style = source.s;
  const font =
    style && typeof style === 'object'
      ? (style.font as
          | {
              bold?: boolean;
              italic?: boolean;
              strike?: boolean;
              underline?: boolean | number | string;
              name?: string;
              sz?: number;
              color?: { rgb?: string };
            }
          | undefined)
      : undefined;
  const fill =
    style && typeof style === 'object'
      ? (style.fill as { fgColor?: { rgb?: string } } | undefined)
      : undefined;
  const alignment =
    style && typeof style === 'object'
      ? (style.alignment as
          | {
              horizontal?: string;
              vertical?: string;
              wrapText?: boolean;
              textRotation?: number;
            }
          | undefined)
      : undefined;
  const target: Partial<Cell> = {};
  if (font?.bold) target.bl = 1;
  if (font?.italic) target.it = 1;
  if (font?.strike) target.cl = 1;
  if (font?.underline) {
    target.un = spreadsheetUnderlineCellValueFromSheetJs(font.underline);
  }
  if (font?.name) target.ff = font.name;
  if (font?.sz !== undefined) target.fs = font.sz;
  const fontColor = spreadsheetColor(font?.color?.rgb);
  if (fontColor) target.fc = fontColor;
  const fillColor = spreadsheetColor(fill?.fgColor?.rgb);
  if (fillColor) target.bg = fillColor;
  if (alignment?.horizontal === 'center') target.ht = 0;
  else if (alignment?.horizontal === 'right') target.ht = 2;
  else if (alignment?.horizontal === 'left') target.ht = 1;
  if (alignment?.vertical === 'center') target.vt = 0;
  else if (alignment?.vertical === 'top') target.vt = 1;
  else if (alignment?.vertical === 'bottom') target.vt = 2;
  if (alignment?.wrapText) target.tb = '2';
  const orientation = spreadsheetTextOrientationFromXlsx(
    alignment?.textRotation,
  );
  const orientationStyle = spreadsheetTextOrientationCellStyle(orientation);
  if (orientationStyle) Object.assign(target, orientationStyle);
  if ((source.z && source.z !== 'General') || source.t === 'e') {
    target.ct = {
      fa: source.z && source.z !== 'General' ? String(source.z) : undefined,
      t: source.t,
    };
  }
  return target;
}

function fortuneCellDisplayText(
  source: CellObject,
  XLSX: typeof import('xlsx'),
): string | undefined {
  if (source.w !== undefined) return source.w;
  if (source.v === undefined || source.v === null) return undefined;
  if (
    source.t === 'd' ||
    source.t === 'e' ||
    (source.z !== undefined && source.z !== 'General')
  ) {
    try {
      return XLSX.utils.format_cell(source);
    } catch {
      return source.v instanceof Date
        ? source.v.toISOString()
        : String(source.v);
    }
  }
  return undefined;
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

function spreadsheetColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const normalized = value.replace(/^#/, '').slice(-6);
  return /^[0-9a-f]{6}$/i.test(normalized)
    ? `#${normalized.toLowerCase()}`
    : undefined;
}
