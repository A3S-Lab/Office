import JSZip from 'jszip';
import { decodeXmlBytes } from './work-ooxml-xml';
import type {
  SpreadsheetPackageScanResult,
  SpreadsheetPackageScanWorkerRequest,
  SpreadsheetPackageScanWorkerResponse,
} from './work-spreadsheet-package-scan-worker-protocol';
import {
  createPlainXlsxWorkbookPlan,
  streamPlainXlsxWorksheet,
} from './work-xlsx-plain-fast-path';
import { scanXlsxWorksheetXml } from './work-xlsx-worksheet-scan';

interface SpreadsheetPackageScanWorkerScope {
  onmessage:
    | ((event: MessageEvent<SpreadsheetPackageScanWorkerRequest>) => void)
    | null;
  postMessage: (
    message: SpreadsheetPackageScanWorkerResponse,
    transfer?: Transferable[],
  ) => void;
}

const scope = globalThis as unknown as SpreadsheetPackageScanWorkerScope;

scope.onmessage = async (event) => {
  if (event.data.kind !== 'scan') return;
  try {
    const zip = await JSZip.loadAsync(event.data.bytes);
    const packagePaths = Object.values(zip.files)
      .filter((entry) => !entry.dir)
      .map((entry) => entry.name);
    const plan = await plainWorkbookPlan(zip, packagePaths);
    let fastPath = plan !== null;
    let fastPathRejected = false;
    if (plan) scope.postMessage({ kind: 'workbook', workbook: plan.workbook });
    const plannedSheetByPath = new Map(
      plan?.sheets.map((sheet) => [sheet.partPath, sheet]),
    );
    const parsedSheetPaths = new Set<string>();
    const worksheets: SpreadsheetPackageScanResult['worksheets'] = {};
    for (const [partPath, entry] of Object.entries(zip.files)) {
      if (
        entry.dir ||
        !partPath.startsWith('xl/worksheets/') ||
        !partPath.endsWith('.xml')
      ) {
        continue;
      }
      const bytes = await entry.async('uint8array');
      const source = decodeXmlBytes(bytes, partPath);
      const plannedSheet = fastPath ? plannedSheetByPath.get(partPath) : null;
      if (plannedSheet) {
        const parsed = streamPlainXlsxWorksheet(
          source,
          (chunk) => {
            scope.postMessage(
              {
                chunk,
                kind: 'plain-cells',
                name: plannedSheet.name,
              },
              [
                chunk.coordinates.buffer,
                chunk.kinds.buffer,
                chunk.numericValues.buffer,
              ],
            );
          },
          ({ columnCount, rowCount }) => {
            scope.postMessage({
              columnCount,
              kind: 'plain-worksheet-start',
              name: plannedSheet.name,
              rowCount,
            });
          },
        );
        if (parsed) {
          parsedSheetPaths.add(partPath);
          worksheets[partPath] = PLAIN_WORKSHEET_SCAN;
          scope.postMessage({
            columnCount: parsed.columnCount,
            dense: true,
            kind: 'worksheet',
            name: plannedSheet.name,
            populatedCellCount: parsed.populatedCellCount,
            properties: parsed.properties,
            rowCount: parsed.rowCount,
          });
          continue;
        }
        fastPath = false;
        if (!fastPathRejected) {
          fastPathRejected = true;
          scope.postMessage({ kind: 'fast-path-rejected' });
        }
      } else if (fastPath) {
        fastPath = false;
        if (!fastPathRejected) {
          fastPathRejected = true;
          scope.postMessage({ kind: 'fast-path-rejected' });
        }
      }
      worksheets[partPath] = scanXlsxWorksheetXml(source);
    }
    fastPath =
      fastPath && plan !== null && parsedSheetPaths.size === plan.sheets.length;
    scope.postMessage({ fastPath, kind: 'success', worksheets });
  } catch {
    scope.postMessage({ kind: 'failure' });
  }
};

const PLAIN_WORKSHEET_SCAN = {
  hasDiagnosticFeatures: false,
  hasFormulaFeatures: false,
  hasImportedFeatures: false,
  requiresSheetJsCellStyles: false,
} as const;

async function plainWorkbookPlan(zip: JSZip, packagePaths: readonly string[]) {
  try {
    const [contentTypes, rootRelationships, workbook, workbookRelationships] =
      await Promise.all([
        zipXml(zip, '[Content_Types].xml'),
        zipXml(zip, '_rels/.rels'),
        zipXml(zip, 'xl/workbook.xml'),
        zipXml(zip, 'xl/_rels/workbook.xml.rels'),
      ]);
    if (
      contentTypes === null ||
      rootRelationships === null ||
      workbook === null ||
      workbookRelationships === null
    ) {
      return null;
    }
    return createPlainXlsxWorkbookPlan({
      contentTypes,
      packagePaths,
      rootRelationships,
      workbook,
      workbookRelationships,
    });
  } catch {
    return null;
  }
}

async function zipXml(zip: JSZip, partPath: string): Promise<string | null> {
  const entry = zip.file(partPath);
  if (!entry || entry.dir) return null;
  return decodeXmlBytes(await entry.async('uint8array'), partPath);
}
