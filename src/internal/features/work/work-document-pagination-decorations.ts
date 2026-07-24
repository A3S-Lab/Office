import { Decoration } from '@tiptap/pm/view';
import type {
  DocumentPaginationVisualBreak,
  DocumentPaginationVisualPageChrome,
  DocumentTableCellPageBreak,
  DocumentTablePaginationBreak,
} from './work-document-pagination-types';

export function pageBreakDecorations(
  pageBreak: DocumentPaginationVisualBreak,
): Decoration[] {
  const tableBreak = pageBreak.tableBreak;
  if (tableBreak?.cellBreaks?.length) {
    const leadingCellIndex = Math.min(
      ...tableBreak.cellBreaks.map((cellBreak) => cellBreak.cellIndex),
    );
    return tableBreak.cellBreaks.map((cellBreak) =>
      Decoration.widget(
        cellBreak.position,
        () =>
          documentTableCellPageBreakElement(
            pageBreak,
            cellBreak,
            cellBreak.cellIndex === leadingCellIndex,
          ),
        {
          key: [
            'document-table-cell-page',
            pageBreak.pageIndex,
            pageBreak.beforeBlockId,
            cellBreak.cellIndex,
          ].join('-'),
          side: -1,
        },
      ),
    );
  }

  return [
    Decoration.widget(
      pageBreak.position,
      () => {
        if (tableBreak) {
          return documentTablePageBreakElement(pageBreak);
        }
        const element = document.createElement('span');
        element.className = 'work-document-auto-page-break';
        element.contentEditable = 'false';
        element.dataset.pageIndex = String(pageBreak.pageIndex + 1);
        element.setAttribute('aria-hidden', 'true');
        applyDocumentPageBreakVariables(element, pageBreak);
        appendDocumentPageBreakChrome(element, pageBreak);
        return element;
      },
      {
        key: `document-page-${pageBreak.pageIndex}-${pageBreak.beforeBlockId}`,
        side: -1,
      },
    ),
  ];
}

function documentTablePageBreakElement(
  pageBreak: DocumentPaginationVisualBreak,
): HTMLTableRowElement {
  const tableBreak = pageBreak.tableBreak as DocumentTablePaginationBreak;
  const row = document.createElement('tr');
  row.className = 'work-document-table-page-break';
  row.contentEditable = 'false';
  row.dataset.pageIndex = String(pageBreak.pageIndex + 1);
  row.dataset.tableId = tableBreak.tableId;
  row.setAttribute('aria-hidden', 'true');
  applyDocumentPageBreakVariables(row, pageBreak);

  const cell = document.createElement('td');
  cell.colSpan = Math.max(1, tableBreak.columnCount);
  const spacer = document.createElement('span');
  spacer.className = 'work-document-table-page-spacer';
  appendDocumentPageBreakChrome(spacer, pageBreak);
  cell.append(spacer);

  if (tableBreak.repeatedHeaderOverlayHtml) {
    cell.append(documentTableRepeatedHeaderElement(tableBreak));
  }
  row.append(cell);
  return row;
}

function documentTableCellPageBreakElement(
  pageBreak: DocumentPaginationVisualBreak,
  cellBreak: DocumentTableCellPageBreak,
  leading: boolean,
): HTMLDivElement {
  const tableBreak = pageBreak.tableBreak as DocumentTablePaginationBreak;
  const element = document.createElement('div');
  element.className = 'work-document-table-cell-page-break';
  element.contentEditable = 'false';
  element.dataset.cellIndex = String(cellBreak.cellIndex);
  element.dataset.pageIndex = String(pageBreak.pageIndex + 1);
  element.dataset.tableId = tableBreak.tableId;
  element.setAttribute('aria-hidden', 'true');
  applyDocumentPageBreakVariables(element, pageBreak);
  element.style.setProperty(
    '--work-document-table-cell-alignment-offset',
    `${cellBreak.alignmentOffset}px`,
  );
  element.style.setProperty(
    '--work-document-table-repeat-header-height',
    `${tableBreak.repeatHeaderHeight}px`,
  );
  element.style.setProperty(
    '--work-document-table-width',
    `${tableBreak.tableWidth}px`,
  );
  element.style.setProperty(
    '--work-document-table-leading-cell-offset-left',
    `${tableBreak.leadingCellOffsetLeft}px`,
  );
  const fallbackWidth =
    tableBreak.tableWidth / Math.max(1, tableBreak.columnCount);
  element.style.setProperty(
    '--work-document-table-cell-offset-left',
    `${cellBreak.tableOffsetLeft ?? cellBreak.cellIndex * fallbackWidth}px`,
  );
  element.style.setProperty(
    '--work-document-table-cell-outer-width',
    `${cellBreak.outerWidth ?? fallbackWidth}px`,
  );
  element.style.setProperty(
    '--work-document-table-cell-content-offset-left',
    `${cellBreak.contentOffsetLeft ?? 0}px`,
  );

  const localLayer = document.createElement('div');
  localLayer.className = 'work-document-table-cell-page-layer';
  const localSpacer = document.createElement('span');
  localSpacer.className = 'work-document-table-cell-local-page-spacer';
  localLayer.append(localSpacer);

  if (tableBreak.repeatedHeaderOverlayHtml) {
    const viewport = document.createElement('div');
    viewport.className = 'work-document-table-cell-repeated-header-viewport';
    viewport.append(
      documentTableRepeatedHeaderElement(
        tableBreak,
        'work-document-table-cell-repeated-header',
      ),
    );
    localLayer.append(viewport);
  }
  element.append(localLayer);

  if (leading) {
    element.classList.add('is-leading');
    const spacer = document.createElement('span');
    spacer.className =
      'work-document-table-page-spacer work-document-table-cell-page-spacer';
    appendDocumentPageBreakChrome(spacer, pageBreak);
    element.prepend(spacer);
  }
  scheduleDocumentTableCellPageBreakPaint(element);
  return element;
}

function scheduleDocumentTableCellPageBreakPaint(element: HTMLElement): void {
  if (typeof requestAnimationFrame !== 'function') return;
  // Blink needs one post-insertion paint invalidation for overflow content in
  // table-cell widget decorations. Opacity avoids another layout pass.
  element.style.opacity = '0.999';
  requestAnimationFrame(() => {
    if (element.isConnected) element.style.opacity = '1';
  });
}

function documentTableRepeatedHeaderElement(
  tableBreak: DocumentTablePaginationBreak,
  extraClassName = '',
): HTMLDivElement {
  const header = document.createElement('div');
  header.className = ['work-document-table-repeated-header', extraClassName]
    .filter(Boolean)
    .join(' ');
  header.contentEditable = 'false';
  header.setAttribute('aria-hidden', 'true');
  header.style.width = `${tableBreak.tableWidth}px`;
  header.style.height = `${tableBreak.repeatHeaderHeight}px`;
  header.insertAdjacentHTML('beforeend', tableBreak.repeatedHeaderOverlayHtml);
  return header;
}

function applyDocumentPageBreakVariables(
  element: HTMLElement,
  pageBreak: DocumentPaginationVisualBreak,
): void {
  element.style.setProperty(
    '--work-document-page-spacer-height',
    `${pageBreak.spacerHeight}px`,
  );
  element.style.setProperty(
    '--work-document-page-gap-offset',
    `${
      pageBreak.remainingBodyHeight +
      pageBreak.page.marginBottom +
      pageBreak.page.footerHeight
    }px`,
  );
  element.style.setProperty(
    '--work-document-page-gap-height',
    `${pageBreak.page.pageGap}px`,
  );
  element.style.setProperty(
    '--work-document-previous-page-footer-top',
    `${pageBreak.remainingBodyHeight + pageBreak.page.marginBottom}px`,
  );
  element.style.setProperty(
    '--work-document-next-page-header-top',
    `${
      pageBreak.remainingBodyHeight +
      pageBreak.page.marginBottom +
      pageBreak.page.footerHeight +
      pageBreak.page.pageGap +
      pageBreak.page.marginTop
    }px`,
  );
  element.style.setProperty(
    '--work-document-page-header-height',
    `${pageBreak.page.headerHeight}px`,
  );
  element.style.setProperty(
    '--work-document-page-footer-height',
    `${pageBreak.page.footerHeight}px`,
  );
  element.style.setProperty(
    '--work-document-page-margin-left',
    `${pageBreak.page.marginLeft}px`,
  );
  element.style.setProperty(
    '--work-document-page-margin-right',
    `${pageBreak.page.marginRight}px`,
  );
  element.style.setProperty(
    '--work-document-break-inline-offset-left',
    `${pageBreak.inlineOffsetLeft}px`,
  );
  element.style.setProperty(
    '--work-document-break-inline-offset-right',
    `${pageBreak.inlineOffsetRight}px`,
  );
}

function appendDocumentPageBreakChrome(
  container: HTMLElement,
  pageBreak: DocumentPaginationVisualBreak,
): void {
  const previous = pageBreak.previousPageChrome;
  if (previous && (previous.footerHtml || previous.showPageNumber)) {
    container.append(documentPageBreakChromeElement('footer', previous));
  }
  const next = pageBreak.nextPageChrome;
  if (next?.headerHtml) {
    container.append(documentPageBreakChromeElement('header', next));
  }
}

function documentPageBreakChromeElement(
  part: 'footer' | 'header',
  chrome: DocumentPaginationVisualPageChrome,
): HTMLDivElement {
  const element = document.createElement('div');
  element.className = `work-document-page-break-chrome ${part}`;
  element.contentEditable = 'false';
  element.dataset.pageChromeVariant = chrome.variant;
  element.setAttribute('aria-hidden', 'true');

  const content = document.createElement('div');
  content.className = 'work-document-page-chrome-html';
  if (part === 'header') {
    content.insertAdjacentHTML('beforeend', chrome.headerHtml);
  } else if (chrome.footerHtml) {
    content.insertAdjacentHTML('beforeend', chrome.footerHtml);
  }
  element.append(content);

  if (part === 'footer' && chrome.showPageNumber) {
    const pageNumber = document.createElement('span');
    pageNumber.className = 'work-document-page-number';
    pageNumber.textContent = String(chrome.pageNumber);
    element.append(pageNumber);
  }
  return element;
}
