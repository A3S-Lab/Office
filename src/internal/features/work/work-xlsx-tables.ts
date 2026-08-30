import JSZip from 'jszip';
import { normalizeSpreadsheetTableCalculatedFormula } from './editors/spreadsheet-table-calculated-columns';
import {
  normalizeSpreadsheetTableTotalsFormula,
  normalizeSpreadsheetTableTotalsLabel,
  spreadsheetTableTotalsFunctionFromOoxml,
  spreadsheetTableTotalsFunctionToOoxml,
} from './editors/spreadsheet-table-totals';
import {
  attribute,
  directChild,
  directChildren,
  firstDescendant,
  OoxmlPackage,
  parseXml,
} from './work-ooxml-package';
import {
  formatSpreadsheetCellRanges,
  parseSpreadsheetCellRanges,
} from './work-spreadsheet-ranges';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from './work-types';
import {
  readXlsxTableFilters,
  xlsxTableAutoFilterXml,
} from './work-xlsx-table-filters';

const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const TRANSITIONAL_DOCUMENT_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_DOCUMENT_RELATIONSHIP_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

export async function readXlsxWorksheetTables(
  archive: OoxmlPackage,
  worksheetPart: string,
): Promise<WorkSpreadsheetTable[]> {
  if (!archive.has(worksheetPart)) return [];
  const worksheet = await archive.xml(worksheetPart);
  const relationships = await archive.relationships(worksheetPart);
  const tableParts = directChild(worksheet.documentElement, 'tableParts');
  if (!tableParts) return [];
  const tables: WorkSpreadsheetTable[] = [];
  for (const tablePart of directChildren(tableParts, 'tablePart')) {
    const relationship = relationships.get(attribute(tablePart, 'r:id') ?? '');
    if (
      !relationship ||
      relationship.targetMode === 'External' ||
      !relationship.type.endsWith('/table') ||
      !archive.has(relationship.target)
    ) {
      continue;
    }
    const parsed = parseXlsxTable(
      await archive.xml(relationship.target),
      relationship.target,
    );
    if (parsed) tables.push(parsed);
  }
  return tables;
}

export async function patchXlsxSpreadsheetTables(
  buffer: ArrayBuffer,
  content: WorkSpreadsheetContent,
): Promise<ArrayBuffer> {
  const archive = await OoxmlPackage.load(buffer);
  const hasExistingTables = archive
    .paths('xl/tables/')
    .some((path) => path.endsWith('.xml'));
  if (
    !hasExistingTables &&
    !content.sheets.some((sheet) => sheet.tables?.length)
  ) {
    return buffer;
  }
  const worksheetParts = await workbookWorksheetParts(archive);
  const zip = await JSZip.loadAsync(buffer);
  for (const path of Object.keys(zip.files)) {
    if (path.startsWith('xl/tables/')) zip.remove(path);
  }

  const contentTypeParts: string[] = [];
  const usedTableIds = new Set<number>();
  let partIndex = 0;
  for (const sheet of content.sheets) {
    const exportedName = sheet.name.slice(0, 31) || '工作表';
    const worksheetPart = worksheetParts.get(exportedName);
    const worksheetEntry = worksheetPart ? zip.file(worksheetPart) : null;
    if (!worksheetPart || !worksheetEntry) continue;
    const worksheet = parseXml(
      await worksheetEntry.async('text'),
      worksheetPart,
    );
    for (const tableParts of directChildren(
      worksheet.documentElement,
      'tableParts',
    )) {
      tableParts.remove();
    }
    const relationshipPath = relationshipsPartPath(worksheetPart);
    const relationships = await readRelationshipsDocument(
      zip,
      relationshipPath,
    );
    removeTableRelationships(relationships);
    const relationshipNamespace = documentRelationshipNamespace(worksheet);
    const validTables = (sheet.tables ?? []).filter(validExportTable);
    const tableParts = validTables.length
      ? worksheet.createElementNS(
          worksheet.documentElement.namespaceURI,
          'tableParts',
        )
      : null;
    if (tableParts)
      tableParts.setAttribute('count', String(validTables.length));

    for (const table of validTables) {
      partIndex += 1;
      const tablePath = `xl/tables/table${partIndex}.xml`;
      const relationshipId = appendRelationship(
        relationships,
        `${relationshipNamespace}/table`,
        `../tables/table${partIndex}.xml`,
      );
      const tablePart = worksheet.createElementNS(
        worksheet.documentElement.namespaceURI,
        'tablePart',
      );
      tablePart.setAttributeNS(relationshipNamespace, 'r:id', relationshipId);
      tableParts?.append(tablePart);
      const tableId = allocateTableId(table.ooxmlId, usedTableIds);
      zip.file(
        tablePath,
        xlsxTableXml(
          table,
          tableId,
          worksheet.documentElement.namespaceURI ??
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        ),
      );
      contentTypeParts.push(tablePath);
    }
    if (tableParts) {
      const anchor = directChildren(worksheet.documentElement).find(
        (child) => child.localName === 'extLst',
      );
      worksheet.documentElement.insertBefore(tableParts, anchor ?? null);
    }
    zip.file(worksheetPart, new XMLSerializer().serializeToString(worksheet));
    if (
      directChildren(relationships.documentElement, 'Relationship').length >
        0 ||
      zip.file(relationshipPath)
    ) {
      zip.file(
        relationshipPath,
        new XMLSerializer().serializeToString(relationships),
      );
    }
  }
  await updateContentTypes(zip, contentTypeParts);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function parseXlsxTable(
  document: Document,
  partPath: string,
): WorkSpreadsheetTable | null {
  const root = document.documentElement;
  if (root.localName !== 'table') return null;
  const name = attribute(root, 'name')?.trim();
  const range = parseSpreadsheetCellRanges(attribute(root, 'ref') ?? '')?.[0];
  const columnsElement = directChild(root, 'tableColumns');
  if (!name || !range || !columnsElement) return null;
  const columns = directChildren(columnsElement, 'tableColumn').flatMap(
    (column) => {
      const columnName = attribute(column, 'name');
      if (columnName === null) return [];
      const calculatedFormula = normalizeSpreadsheetTableCalculatedFormula(
        directChild(column, 'calculatedColumnFormula')?.textContent,
      );
      const totalsFunction = spreadsheetTableTotalsFunctionFromOoxml(
        attribute(column, 'totalsRowFunction'),
      );
      const totalsLabel = normalizeSpreadsheetTableTotalsLabel(
        attribute(column, 'totalsRowLabel'),
      );
      const totalsFormula = normalizeSpreadsheetTableTotalsFormula(
        directChild(column, 'totalsRowFormula')?.textContent,
      );
      const normalizedTotalsFunction = totalsFormula
        ? 'custom'
        : totalsFunction;
      return [
        {
          name: columnName,
          ...(calculatedFormula ? { calculatedFormula } : {}),
          ...(normalizedTotalsFunction
            ? { totalsFunction: normalizedTotalsFunction }
            : {}),
          ...(totalsLabel && !totalsFormula ? { totalsLabel } : {}),
          ...(totalsFormula ? { totalsFormula } : {}),
        },
      ];
    },
  );
  const width = range.column[1] - range.column[0] + 1;
  if (columns.length !== width) return null;
  const ooxmlId = positiveIntegerAttribute(root, 'id');
  const displayName = attribute(root, 'displayName')?.trim();
  const headerRow = attribute(root, 'headerRowCount') !== '0';
  const totalsRow =
    positiveIntegerAttribute(root, 'totalsRowCount') !== null ||
    booleanAttribute(root, 'totalsRowShown');
  const styleInfo = directChild(root, 'tableStyleInfo');
  const style = parseTableStyle(attribute(styleInfo ?? root, 'name'));
  const stablePart = partPath.replace(/[^A-Za-z0-9]+/g, '-');
  return {
    id: `spreadsheet-table-${stablePart}-${ooxmlId ?? name}`,
    ...(ooxmlId === null ? {} : { ooxmlId }),
    name,
    ...(displayName && displayName !== name ? { displayName } : {}),
    range,
    columns,
    filters: headerRow ? readXlsxTableFilters(document, width) : [],
    headerRow,
    totalsRow,
    style,
    showFirstColumn: booleanAttribute(styleInfo, 'showFirstColumn'),
    showLastColumn: booleanAttribute(styleInfo, 'showLastColumn'),
    showRowStripes: booleanAttribute(styleInfo, 'showRowStripes'),
    showColumnStripes: booleanAttribute(styleInfo, 'showColumnStripes'),
  };
}

function xlsxTableXml(
  table: WorkSpreadsheetTable,
  tableId: number,
  namespace: string,
): string {
  const reference = formatSpreadsheetCellRanges([table.range]);
  const displayName = table.displayName?.trim() || table.name;
  const filter = table.headerRow
    ? xlsxTableAutoFilterXml(table.filters, reference, table.totalsRow)
    : '';
  const columns = table.columns
    .map((column, index) => {
      const formula = normalizeSpreadsheetTableCalculatedFormula(
        column.calculatedFormula,
      );
      const formulaElement = formula
        ? `<calculatedColumnFormula>${escapeXml(
            formula.replace(/^=/, ''),
          )}</calculatedColumnFormula>`
        : '';
      const totalsFunction = spreadsheetTableTotalsFunctionToOoxml(
        column.totalsFunction,
      );
      const totalsLabel = normalizeSpreadsheetTableTotalsLabel(
        column.totalsLabel,
      );
      const totalsFormula = normalizeSpreadsheetTableTotalsFormula(
        column.totalsFormula,
      );
      const attributes = [
        `id="${index + 1}"`,
        `name="${escapeXml(column.name)}"`,
        ...(totalsFunction
          ? [`totalsRowFunction="${escapeXml(totalsFunction)}"`]
          : []),
        ...(totalsLabel && !totalsFormula
          ? [`totalsRowLabel="${escapeXml(totalsLabel)}"`]
          : []),
      ].join(' ');
      const totalsFormulaElement = totalsFormula
        ? `<totalsRowFormula>${escapeXml(
            totalsFormula.replace(/^=/, ''),
          )}</totalsRowFormula>`
        : '';
      const children = `${formulaElement}${totalsFormulaElement}`;
      return children
        ? `<tableColumn ${attributes}>${children}</tableColumn>`
        : `<tableColumn ${attributes}/>`;
    })
    .join('');
  const styleName = tableStyleName(table.style);
  const style = styleName
    ? `<tableStyleInfo name="${styleName}" showColumnStripes="${booleanToken(
        table.showColumnStripes,
      )}" showFirstColumn="${booleanToken(
        table.showFirstColumn,
      )}" showLastColumn="${booleanToken(
        table.showLastColumn,
      )}" showRowStripes="${booleanToken(table.showRowStripes)}"/>`
    : '';
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<table xmlns="${escapeXml(namespace)}" id="${tableId}" name="${escapeXml(
      table.name,
    )}" displayName="${escapeXml(displayName)}" ref="${reference}"`,
    ` headerRowCount="${booleanToken(table.headerRow)}"`,
    ` totalsRowCount="${booleanToken(table.totalsRow)}"`,
    ` totalsRowShown="${booleanToken(table.totalsRow)}">`,
    filter,
    `<tableColumns count="${table.columns.length}">${columns}</tableColumns>`,
    style,
    '</table>',
  ].join('');
}

function validExportTable(table: WorkSpreadsheetTable): boolean {
  const range = table.range;
  const width = range.column[1] - range.column[0] + 1;
  const height = range.row[1] - range.row[0] + 1;
  return (
    Number.isSafeInteger(range.row[0]) &&
    Number.isSafeInteger(range.row[1]) &&
    Number.isSafeInteger(range.column[0]) &&
    Number.isSafeInteger(range.column[1]) &&
    range.row[0] >= 0 &&
    range.column[0] >= 0 &&
    width > 0 &&
    height > Number(table.headerRow) + Number(table.totalsRow) &&
    table.columns.length === width &&
    Boolean(table.name.trim()) &&
    table.columns.every((column) => Boolean(column.name.trim())) &&
    tableStyleName(table.style) !== null
  );
}

function parseTableStyle(value: string | null): WorkSpreadsheetTableStyle {
  if (!value) return { family: 'none' };
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(value);
  if (!match) return { family: 'none' };
  const family = match[1]?.toLocaleLowerCase() as 'dark' | 'light' | 'medium';
  const number = Number(match[2]);
  const maximum = family === 'light' ? 21 : family === 'medium' ? 28 : 11;
  return number >= 1 && number <= maximum
    ? { family, number }
    : { family: 'none' };
}

function tableStyleName(style: WorkSpreadsheetTableStyle): string | null {
  if (style.family === 'none') return '';
  const maximum =
    style.family === 'light' ? 21 : style.family === 'medium' ? 28 : 11;
  if (
    !Number.isInteger(style.number) ||
    style.number < 1 ||
    style.number > maximum
  ) {
    return null;
  }
  const family = `${style.family.charAt(0).toUpperCase()}${style.family.slice(1)}`;
  return `TableStyle${family}${style.number}`;
}

async function workbookWorksheetParts(
  archive: OoxmlPackage,
): Promise<Map<string, string>> {
  if (!archive.has('xl/workbook.xml')) return new Map();
  const workbook = await archive.xml('xl/workbook.xml');
  const relationships = await archive.relationships('xl/workbook.xml');
  const parts = new Map<string, string>();
  for (const sheet of firstDescendant(workbook, 'sheets')?.children ?? []) {
    if (!(sheet instanceof Element) || sheet.localName !== 'sheet') continue;
    const name = attribute(sheet, 'name');
    const relationship = relationships.get(attribute(sheet, 'r:id') ?? '');
    if (name && relationship?.type.endsWith('/worksheet')) {
      parts.set(name, relationship.target);
    }
  }
  return parts;
}

async function readRelationshipsDocument(
  zip: JSZip,
  partPath: string,
): Promise<Document> {
  const entry = zip.file(partPath);
  return entry
    ? parseXml(await entry.async('text'), partPath)
    : parseXml(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
        partPath,
      );
}

function appendRelationship(
  document: Document,
  type: string,
  target: string,
): string {
  const id = nextRelationshipId(document);
  const relationship = document.createElementNS(
    document.documentElement.namespaceURI ?? PACKAGE_RELATIONSHIP_NAMESPACE,
    'Relationship',
  );
  relationship.setAttribute('Id', id);
  relationship.setAttribute('Type', type);
  relationship.setAttribute('Target', target);
  document.documentElement.append(relationship);
  return id;
}

function removeTableRelationships(document: Document): void {
  for (const relationship of directChildren(
    document.documentElement,
    'Relationship',
  )) {
    if ((attribute(relationship, 'Type') ?? '').endsWith('/table')) {
      relationship.remove();
    }
  }
}

function nextRelationshipId(document: Document): string {
  const used = new Set(
    directChildren(document.documentElement, 'Relationship').map((item) =>
      attribute(item, 'Id'),
    ),
  );
  for (let index = 1; index < 1_000_000; index += 1) {
    const id = `rId${index}`;
    if (!used.has(id)) return id;
  }
  return `rId${Date.now()}`;
}

async function updateContentTypes(
  zip: JSZip,
  tableParts: readonly string[],
): Promise<void> {
  const entry = zip.file('[Content_Types].xml');
  if (!entry) return;
  const document = parseXml(await entry.async('text'), '[Content_Types].xml');
  for (const override of directChildren(document.documentElement, 'Override')) {
    if (attribute(override, 'ContentType') === TABLE_CONTENT_TYPE) {
      override.remove();
    }
  }
  const namespace = document.documentElement.namespaceURI;
  for (const path of tableParts) {
    const override = document.createElementNS(namespace, 'Override');
    override.setAttribute('PartName', `/${path}`);
    override.setAttribute('ContentType', TABLE_CONTENT_TYPE);
    document.documentElement.append(override);
  }
  zip.file(
    '[Content_Types].xml',
    new XMLSerializer().serializeToString(document),
  );
}

function relationshipsPartPath(sourcePart: string): string {
  const segments = sourcePart.split('/');
  const name = segments.pop() ?? '';
  return [...segments, '_rels', `${name}.rels`].join('/');
}

function documentRelationshipNamespace(document: Document): string {
  return document.documentElement.namespaceURI?.includes('purl.oclc.org')
    ? STRICT_DOCUMENT_RELATIONSHIP_NAMESPACE
    : TRANSITIONAL_DOCUMENT_RELATIONSHIP_NAMESPACE;
}

function allocateTableId(
  requested: number | undefined,
  used: Set<number>,
): number {
  if (
    requested &&
    Number.isSafeInteger(requested) &&
    requested > 0 &&
    !used.has(requested)
  ) {
    used.add(requested);
    return requested;
  }
  let candidate = 1;
  while (used.has(candidate)) candidate += 1;
  used.add(candidate);
  return candidate;
}

function positiveIntegerAttribute(
  element: Element,
  name: string,
): number | null {
  const value = Number(attribute(element, name));
  return Number.isSafeInteger(value) && value > 0 ? value : null;
}

function booleanAttribute(element: Element | undefined, name: string): boolean {
  return element
    ? ['1', 'true', 'on'].includes(
        (attribute(element, name) ?? '').toLocaleLowerCase(),
      )
    : false;
}

function booleanToken(value: boolean): string {
  return value ? '1' : '0';
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
