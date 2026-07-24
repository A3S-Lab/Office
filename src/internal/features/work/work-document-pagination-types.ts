import type {
  OfficeKernelLayoutBlock,
  OfficeKernelLayoutBreak,
  OfficeKernelPageMetrics,
  OfficeKernelTextLayoutParagraph,
} from '../../kernel/office-kernel-protocol';
import type { WorkDocumentSectionLayout } from './work-types';

export interface DocumentPaginationVisualBreak extends OfficeKernelLayoutBreak {
  position: number;
  page: OfficeKernelPageMetrics;
  inlineOffsetLeft: number;
  inlineOffsetRight: number;
  previousPageChrome?: DocumentPaginationVisualPageChrome;
  nextPageChrome?: DocumentPaginationVisualPageChrome;
  tableBreak?: DocumentTablePaginationBreak;
}

export interface DocumentPaginationVisualPageChrome {
  variant: 'default' | 'even' | 'first';
  headerHtml: string;
  footerHtml: string;
  showPageNumber: boolean;
  pageNumber: number;
}

export interface DocumentPaginationSection {
  id: string;
  index: number;
  position: number;
  layout: WorkDocumentSectionLayout;
}

export interface DocumentTablePaginationBreak {
  tableId: string;
  columnCount: number;
  colgroupHtml: string;
  repeatedHeaderRowsHtml: string[];
  repeatedHeaderOverlayHtml: string;
  repeatHeaderHeight: number;
  tableWidth: number;
  leadingCellOffsetLeft: number;
  cellBreaks?: DocumentTableCellPageBreak[];
}

export interface DocumentTableCellBoundary {
  position: number;
  y: number;
}

export interface DocumentTableCellFragmentMeasurement {
  cellIndex: number;
  from: number;
  to: number;
  boundaries: DocumentTableCellBoundary[];
  tableOffsetLeft?: number;
  outerWidth?: number;
  contentOffsetLeft?: number;
}

export interface DocumentTableCellPageBreak {
  cellIndex: number;
  position: number;
  alignmentOffset: number;
  tableOffsetLeft?: number;
  outerWidth?: number;
  contentOffsetLeft?: number;
}

export interface DocumentTableRowFragmentPlan {
  height: number;
  cellRanges: Array<{ from: number; to: number }>;
  cellBreaks?: DocumentTableCellPageBreak[];
}

export interface MeasuredDocumentLayoutBlock {
  block: OfficeKernelLayoutBlock;
  element: HTMLElement;
  from: number;
  to: number;
  inlineOffsetLeft: number;
  inlineOffsetRight: number;
  observeResize: boolean;
  section?: DocumentPaginationSection;
  tableBreak?: DocumentTablePaginationBreak;
  selectionRanges?: Array<{ from: number; to: number }>;
}

export interface DocumentPaginationSnapshot {
  blocks: MeasuredDocumentLayoutBlock[];
  measuredBlockCount: number;
  reusedBlockCount: number;
  unsupportedLayout: boolean;
}

export interface DocumentTextLayoutCollection {
  paragraphs: OfficeKernelTextLayoutParagraph[];
}
