import JSZip from 'jszip';
import {
  documentCellPropertyRevisionOmmlFromElement,
  documentRowPropertyRevisionOmmlFromElement,
  documentTablePropertyRevisionOmmlFromElement,
  importDocxScopedPropertyRevisionElement,
  type DocxTableScopedPropertyRevisionKind,
} from './work-document-table-property-revision';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const MAX_TABLE_PROPERTY_REVISION_PATCHES = 4_096;

export class DocxTablePropertyRevisionPatchCollector {
  readonly tablePatches: Array<string | null> = [];
  readonly rowPatches: Array<string | null> = [];
  readonly cellPatches: Array<string | null> = [];

  /** @deprecated Prefer {@link recordTable}. */
  record(element: HTMLTableElement): void {
    this.recordTable(element);
  }

  recordTable(element: HTMLTableElement): void {
    if (
      element.getAttribute('data-document-change') === 'true' &&
      element.dataset.changeKind === 'table-formatting'
    ) {
      pushPatch(this.tablePatches, null);
      return;
    }
    pushPatch(
      this.tablePatches,
      documentTablePropertyRevisionOmmlFromElement(element),
    );
  }

  recordRow(element: HTMLTableRowElement): void {
    if (
      element.getAttribute('data-document-change') === 'true' &&
      element.dataset.changeKind === 'row-formatting'
    ) {
      pushPatch(this.rowPatches, null);
      return;
    }
    pushPatch(
      this.rowPatches,
      documentRowPropertyRevisionOmmlFromElement(element),
    );
  }

  recordCell(element: HTMLTableCellElement): void {
    if (
      element.getAttribute('data-document-change') === 'true' &&
      element.dataset.changeKind === 'cell-formatting'
    ) {
      pushPatch(this.cellPatches, null);
      return;
    }
    pushPatch(
      this.cellPatches,
      documentCellPropertyRevisionOmmlFromElement(element),
    );
  }

  get patches(): Array<string | null> {
    return this.tablePatches;
  }
}

export async function patchDocxTablePropertyRevisions(
  buffer: ArrayBuffer,
  collector:
    | DocxTablePropertyRevisionPatchCollector
    | readonly (string | null)[],
): Promise<ArrayBuffer> {
  const tablePatches =
    collector instanceof DocxTablePropertyRevisionPatchCollector
      ? collector.tablePatches
      : collector;
  const rowPatches =
    collector instanceof DocxTablePropertyRevisionPatchCollector
      ? collector.rowPatches
      : [];
  const cellPatches =
    collector instanceof DocxTablePropertyRevisionPatchCollector
      ? collector.cellPatches
      : [];
  if (
    !tablePatches.some(Boolean) &&
    !rowPatches.some(Boolean) &&
    !cellPatches.some(Boolean)
  ) {
    return buffer;
  }
  assertPatchLimit(tablePatches);
  assertPatchLimit(rowPatches);
  assertPatchLimit(cellPatches);

  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      'generated DOCX word/document.xml',
    ),
    'generated DOCX word/document.xml',
  );
  let changed = false;
  changed =
    applyScopedPatches(document, 'tbl', 'tblPr', 'tblPrChange', tablePatches) ||
    changed;
  changed =
    applyScopedPatches(document, 'tr', 'trPr', 'trPrChange', rowPatches) ||
    changed;
  changed =
    applyScopedPatches(document, 'tc', 'tcPr', 'tcPrChange', cellPatches) ||
    changed;
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function applyScopedPatches(
  document: Document,
  scopeLocalName: 'tbl' | 'tr' | 'tc',
  propertiesLocalName: 'tblPr' | 'trPr' | 'tcPr',
  changeKind: DocxTableScopedPropertyRevisionKind,
  patches: readonly (string | null)[],
): boolean {
  let patchIndex = 0;
  let changed = false;
  for (const scope of descendants(document, scopeLocalName)) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(scope.namespaceURI ?? '')) {
      continue;
    }
    const omml = patches[patchIndex] ?? null;
    patchIndex += 1;
    if (!omml) continue;
    const revision = importDocxScopedPropertyRevisionElement(
      document,
      omml,
      changeKind,
    );
    if (!revision) continue;
    let properties = directChild(scope, propertiesLocalName);
    if (!properties) {
      const namespace =
        scope.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
      const prefix = scope.prefix || 'w';
      properties = document.createElementNS(
        namespace,
        `${prefix}:${propertiesLocalName}`,
      );
      scope.insertBefore(properties, scope.firstChild);
    }
    const existing = directChild(properties, changeKind);
    if (existing) existing.replaceWith(revision);
    else properties.append(revision);
    changed = true;
  }
  return changed;
}

function pushPatch(patches: Array<string | null>, omml: string | null): void {
  if (patches.length >= MAX_TABLE_PROPERTY_REVISION_PATCHES) {
    throw new Error('Document exceeds the table property-revision limit.');
  }
  patches.push(omml);
}

function assertPatchLimit(patches: readonly (string | null)[]): void {
  if (patches.length > MAX_TABLE_PROPERTY_REVISION_PATCHES) {
    throw new Error('Document exceeds the table property-revision limit.');
  }
}
