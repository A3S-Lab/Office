import { documentSectionDomAttributes } from './work-document-section';
import {
  streamLargeSimpleDocxDocumentXml,
  type LargeSimpleDocxParseOptions,
  type LargeSimpleDocxStreamResult,
  type LargeSimpleDocxTableRow,
} from './work-docx-large-document-parser';
import { windowDocumentModel } from './work-document-windowing';
import type { OoxmlPackage } from './work-ooxml-package';
import type {
  WorkCompatibilityReport,
  WorkDocumentContent,
  WorkDocumentNode,
  WorkDocumentSectionLayout,
} from './work-types';

const HTML_SPECIAL_CHARACTER = /[&<>"]/;

export {
  LARGE_SIMPLE_DOCX_MINIMUM_LOGICAL_BLOCKS,
  simpleDocxXmlIsEligibleForLargeImport,
  streamLargeSimpleDocxDocumentXml,
} from './work-docx-large-document-parser';
export type {
  LargeSimpleDocxParseOptions,
  LargeSimpleDocxStreamResult,
  LargeSimpleDocxStreamSink,
} from './work-docx-large-document-parser';

export interface LargeSimpleDocxParseResult {
  html: string;
  layout: WorkDocumentSectionLayout;
  logicalBlockCount: number;
  paragraphCount: number;
  root: WorkDocumentNode;
  tableRowCount: number;
}

export async function parseLargeSimpleDocxPackage(
  archive: OoxmlPackage,
  options: LargeSimpleDocxParseOptions = {},
): Promise<LargeSimpleDocxParseResult | null> {
  if (!largeSimpleDocxPackageIsCandidate(archive)) return null;
  return parseLargeSimpleDocxDocumentXml(
    await archive.text('word/document.xml'),
    options,
  );
}

export function largeSimpleDocxPackageIsCandidate(
  archive: OoxmlPackage,
): boolean {
  const wordParts = archive.paths('word/');
  return (
    wordParts.length === 1 &&
    wordParts[0]?.toLowerCase() === 'word/document.xml'
  );
}

/**
 * A deliberately narrow fast path for huge, plain WordprocessingML bodies.
 * Any rich feature routes to the complete Mammoth plus OOXML marker pipeline.
 */
export function parseLargeSimpleDocxDocumentXml(
  source: string,
  options: LargeSimpleDocxParseOptions = {},
): LargeSimpleDocxParseResult | null {
  const blocks: WorkDocumentNode[] = [];
  const html: string[] = [];
  let tableRows: WorkDocumentNode[] | null = null;
  let tableHtml: string[] | null = null;
  const streamed = streamLargeSimpleDocxDocumentXml(source, options, {
    paragraph: (text) => {
      if (tableRows || tableHtml) {
        throw new Error('A table stream cannot contain a top-level paragraph.');
      }
      const paragraph = createLargeSimpleDocxParagraph(text);
      blocks.push(paragraph.node);
      html.push(paragraph.html);
    },
    tableEnd: () => {
      if (!tableRows || !tableHtml) {
        throw new Error('A simple DOCX table stream was not started.');
      }
      const table = createLargeSimpleDocxTable(tableRows, tableHtml);
      blocks.push(table.node);
      html.push(table.html);
      tableRows = null;
      tableHtml = null;
    },
    tableRow: (row) => {
      if (!tableRows || !tableHtml) {
        throw new Error('A simple DOCX table row has no owning table.');
      }
      const projected = createLargeSimpleDocxTableRow(row);
      tableRows.push(projected.node);
      tableHtml.push(projected.html);
    },
    tableStart: () => {
      if (tableRows || tableHtml) {
        throw new Error('Nested simple DOCX tables are not supported.');
      }
      tableRows = [];
      tableHtml = [];
    },
  });
  if (!streamed || tableRows || tableHtml || !blocks.length) return null;
  return createLargeSimpleDocxParseResult(streamed, blocks, html, options);
}

export function createLargeSimpleDocxParseResult(
  streamed: LargeSimpleDocxStreamResult,
  blocks: WorkDocumentNode[],
  html: string[],
  options: LargeSimpleDocxParseOptions = {},
): LargeSimpleDocxParseResult {
  const sectionId = 'document-section-1';
  const projectionStartedAt = largeDocxImportNow();
  const sectionAttributes = documentSectionDomAttributes(
    streamed.layout,
    sectionId,
  );
  const canonicalHtml = `<section${htmlAttributes(sectionAttributes)}>${html.join('')}</section>`;
  recordLargeDocxImportMeasure(
    'a3s-office.document.large-simple-html',
    projectionStartedAt,
    largeDocxImportNow(),
  );
  const windowingStartedAt = largeDocxImportNow();
  const root = windowDocumentModel(
    {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: documentSectionModelAttributes(streamed.layout, sectionId),
          content: blocks,
        },
      ],
    },
    {
      ...options.windowing,
      trustedIntegrityFeatures: 0,
    },
  );
  recordLargeDocxImportMeasure(
    'a3s-office.document.large-simple-windowing',
    windowingStartedAt,
    largeDocxImportNow(),
  );
  return {
    html: canonicalHtml,
    layout: streamed.layout,
    logicalBlockCount: streamed.logicalBlockCount,
    paragraphCount: streamed.paragraphCount,
    root,
    tableRowCount: streamed.tableRowCount,
  };
}

export function createLargeSimpleDocxTable(
  rows: WorkDocumentNode[],
  htmlRows: string[],
): { html: string; node: WorkDocumentNode } {
  return {
    html: `<table data-office-table-imported="true"><tbody>${htmlRows.join('')}</tbody></table>`,
    node: {
      type: 'table',
      attrs: { officeImported: true },
      content: rows,
    },
  };
}

export function createLargeSimpleDocxParagraph(text: string): {
  html: string;
  node: WorkDocumentNode;
} {
  return {
    html: `<p>${escapeHtml(text)}</p>`,
    node: {
      type: 'paragraph',
      ...(text ? { content: [{ type: 'text', text }] } : {}),
    },
  };
}

export function createLargeSimpleDocxTableRow(row: LargeSimpleDocxTableRow): {
  html: string;
  node: WorkDocumentNode;
} {
  const cellNodes: WorkDocumentNode[] = [];
  const cellHtml: string[] = [];
  let textIndex = 0;
  for (const paragraphCount of row.cellParagraphCounts) {
    const paragraphNodes: WorkDocumentNode[] = [];
    let paragraphsHtml = '';
    for (let index = 0; index < paragraphCount; index += 1) {
      const text = row.texts[textIndex] ?? '';
      const projected = createLargeSimpleDocxParagraph(text);
      paragraphNodes.push(projected.node);
      paragraphsHtml += projected.html;
      textIndex += 1;
    }
    cellNodes.push({
      type: 'tableCell',
      content: paragraphNodes,
    });
    cellHtml.push(`<td>${paragraphsHtml}</td>`);
  }
  return {
    html: `<tr>${cellHtml.join('')}</tr>`,
    node: { type: 'tableRow', content: cellNodes },
  };
}

export function largeSimpleDocxContent(
  result: LargeSimpleDocxParseResult,
): WorkDocumentContent {
  return {
    type: 'document',
    html: result.html,
    pageSize: result.layout.pageSize,
    orientation: result.layout.orientation,
    margins: result.layout.margins,
    columns: result.layout.columns,
    headerText: result.layout.headerText,
    footerText: result.layout.footerText,
    showPageNumbers: result.layout.showPageNumbers,
    pageNumberStart: result.layout.pageNumberStart,
    pageChrome: result.layout.pageChrome,
    pageBorders: result.layout.pageBorders,
    pageMargins: result.layout.pageMargins,
    pageGeometry: result.layout.pageGeometry,
    paperSource: result.layout.paperSource,
  };
}

export function largeSimpleDocxCompatibilityReport(
  file: File,
  result: LargeSimpleDocxParseResult,
): WorkCompatibilityReport {
  return {
    sourceFormat: 'DOCX',
    sourceName: file.name,
    assessedAt: Date.now(),
    issues: [
      {
        code: 'docx.page-layout',
        feature: 'Page layout',
        message:
          'Paper size and margins are preserved; exact pagination and line wrapping may normalize.',
        severity: 'warning',
      },
      {
        code: 'docx.package-state',
        feature: 'OOXML package state',
        message:
          'The plain source package remains attached for source-backed export.',
        severity: 'info',
      },
      {
        code: 'docx.large-document-windowing',
        feature: 'Large document rendering',
        message: `${result.logicalBlockCount} logical block(s) use bounded viewport rendering while the complete structured model remains editable.`,
        severity: 'info',
      },
      ...(result.tableRowCount
        ? [
            {
              code: 'docx.tables',
              feature: 'Tables',
              message: `${result.tableRowCount} table row(s) remain editable and are joined into canonical DOCX tables on export.`,
              severity: 'info' as const,
            },
          ]
        : []),
    ],
  };
}

function documentSectionModelAttributes(
  layout: WorkDocumentSectionLayout,
  id: string,
): Record<string, string | number | boolean | null> {
  const attributes = documentSectionDomAttributes(layout, id);
  return {
    id,
    pageSize: layout.pageSize,
    orientation: layout.orientation,
    marginTop: layout.margins.top,
    marginRight: layout.margins.right,
    marginBottom: layout.margins.bottom,
    marginLeft: layout.margins.left,
    columnCount: layout.columns.count,
    columnSpacing: layout.columns.spacing,
    columnSeparator: layout.columns.separator,
    columnLayout: attributes['data-section-column-layout'] ?? '',
    breakAfter: layout.breakAfter,
    headerText: layout.headerText ?? '',
    footerText: layout.footerText ?? '',
    showPageNumbers: Boolean(layout.showPageNumbers),
    pageNumberStart: layout.pageNumberStart ?? null,
    pageChrome: attributes['data-section-page-chrome'] ?? '',
    pageBorders: attributes['data-section-page-borders'] ?? '',
    pageMargins: attributes['data-section-page-margins'] ?? '',
    pageGeometry: attributes['data-section-page-geometry'] ?? '',
    paperSource: attributes['data-section-paper-source'] ?? '',
    documentGridType: attributes['data-section-document-grid-type'] ?? '',
    documentGridLinePitch: null,
  };
}

function htmlAttributes(attributes: Record<string, string>): string {
  return Object.entries(attributes)
    .map(([name, value]) => ` ${name}="${escapeHtml(value)}"`)
    .join('');
}

function escapeHtml(value: string): string {
  if (!HTML_SPECIAL_CHARACTER.test(value)) return value;
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function largeDocxImportNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function recordLargeDocxImportMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, number | boolean>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing is diagnostic only and must never affect file import.
  }
}
