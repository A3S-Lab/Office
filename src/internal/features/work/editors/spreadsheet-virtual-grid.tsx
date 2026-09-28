import type { CellMatrix } from '@fortune-sheet/core';
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react';
import {
  spreadsheetVirtualGridActiveCellId,
  spreadsheetVirtualGridAria,
} from './spreadsheet-virtual-grid-aria';
import {
  spreadsheetVirtualGridCellText,
  spreadsheetVirtualGridPaintCells,
  type SpreadsheetVirtualGridMerge,
} from './spreadsheet-virtual-grid-paint';
import {
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
  spreadsheetVirtualGridViewport,
} from './spreadsheet-virtual-grid-viewport';

/** Browser layout trees choke on multi-million-pixel spacers; map scroll. */
const SPREADSHEET_VIRTUAL_GRID_SCROLL_CAP = 5_000_000;

export interface SpreadsheetVirtualGridProps {
  sheetName: string;
  data: CellMatrix | undefined;
  rowCount: number;
  columnCount: number;
  activeRow: number;
  activeColumn: number;
  rowHeight?: number;
  columnWidth?: number;
  editable?: boolean;
  merges?: readonly SpreadsheetVirtualGridMerge[];
  onActiveCellChange: (row: number, column: number) => void;
  onCommitCell?: (row: number, column: number, raw: string) => void;
}

function mapScrollerOffset(
  scrollerOffset: number,
  scrollerExtent: number,
  clientSize: number,
  contentExtent: number,
): number {
  const scrollerMax = Math.max(0, scrollerExtent - clientSize);
  const contentMax = Math.max(0, contentExtent - clientSize);
  if (scrollerMax <= 0 || contentMax <= 0) return 0;
  if (contentMax <= scrollerMax) return Math.min(scrollerOffset, contentMax);
  return (scrollerOffset / scrollerMax) * contentMax;
}

function scrollerOffsetForContent(
  contentOffset: number,
  scrollerExtent: number,
  clientSize: number,
  contentExtent: number,
): number {
  const scrollerMax = Math.max(0, scrollerExtent - clientSize);
  const contentMax = Math.max(0, contentExtent - clientSize);
  if (scrollerMax <= 0 || contentMax <= 0) return 0;
  if (contentMax <= scrollerMax) return Math.min(contentOffset, scrollerMax);
  return (contentOffset / contentMax) * scrollerMax;
}

/**
 * A3S-owned spreadsheet viewport: canvas paint for the visible sparse window
 * plus a real `role="grid"` focus surface. Fortune remains optional underneath
 * for command-port continuity during migration.
 */
export function SpreadsheetVirtualGrid({
  sheetName,
  data,
  rowCount,
  columnCount,
  activeRow,
  activeColumn,
  rowHeight = SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
  columnWidth = SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  editable = true,
  merges,
  onActiveCellChange,
  onCommitCell,
}: SpreadsheetVirtualGridProps) {
  const scrollerRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const editorRef = useRef<HTMLInputElement>(null);
  const editingRef = useRef(false);
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [clientSize, setClientSize] = useState({ width: 0, height: 0 });
  const [draft, setDraft] = useState<string | null>(null);

  const aria = spreadsheetVirtualGridAria({
    sheetName,
    rowCount,
    columnCount,
    activeRow,
    activeColumn,
  });
  const activeCellId = spreadsheetVirtualGridActiveCellId(
    Math.min(Math.max(0, activeRow), Math.max(0, rowCount - 1)),
    Math.min(Math.max(0, activeColumn), Math.max(0, columnCount - 1)),
  );

  useLayoutEffect(() => {
    const scroller = scrollerRef.current;
    if (!scroller) return;
    const sync = () => {
      setScrollTop(scroller.scrollTop);
      setScrollLeft(scroller.scrollLeft);
      setClientSize({
        width: scroller.clientWidth,
        height: scroller.clientHeight,
      });
    };
    sync();
    const observer = new ResizeObserver(sync);
    observer.observe(scroller);
    return () => observer.disconnect();
  }, []);

  const contentHeight = rowCount * rowHeight;
  const contentWidth = columnCount * columnWidth;
  const spacerHeight = Math.min(contentHeight, SPREADSHEET_VIRTUAL_GRID_SCROLL_CAP);
  const spacerWidth = Math.min(contentWidth, SPREADSHEET_VIRTUAL_GRID_SCROLL_CAP);

  const paint = useCallback(() => {
    const canvas = canvasRef.current;
    const scroller = scrollerRef.current;
    if (!canvas || !scroller) return;
    const width = scroller.clientWidth;
    const height = scroller.clientHeight;
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1;
    if (canvas.width !== Math.floor(width * dpr) || canvas.height !== Math.floor(height * dpr)) {
      canvas.width = Math.floor(width * dpr);
      canvas.height = Math.floor(height * dpr);
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
    }
    const context = canvas.getContext('2d');
    if (!context) return;
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);

    const mappedScrollTop = mapScrollerOffset(
      scroller.scrollTop,
      spacerHeight,
      height,
      contentHeight,
    );
    const mappedScrollLeft = mapScrollerOffset(
      scroller.scrollLeft,
      spacerWidth,
      width,
      contentWidth,
    );
    const viewport = spreadsheetVirtualGridViewport({
      scrollTop: mappedScrollTop,
      scrollLeft: mappedScrollLeft,
      clientHeight: height,
      clientWidth: width,
      rowCount,
      columnCount,
      rowHeight,
      columnWidth,
      overscan: 1,
    });

    context.strokeStyle = '#d0d0d0';
    context.lineWidth = 1;
    for (let row = viewport.rowStart; row < viewport.rowEnd; row += 1) {
      const y = (row - viewport.rowStart) * rowHeight - viewport.offsetTop;
      context.beginPath();
      context.moveTo(0, y + rowHeight);
      context.lineTo(width, y + rowHeight);
      context.stroke();
    }
    for (let column = viewport.columnStart; column < viewport.columnEnd; column += 1) {
      const x =
        (column - viewport.columnStart) * columnWidth - viewport.offsetLeft;
      context.beginPath();
      context.moveTo(x + columnWidth, 0);
      context.lineTo(x + columnWidth, height);
      context.stroke();
    }

    const activeX =
      (activeColumn - viewport.columnStart) * columnWidth - viewport.offsetLeft;
    const activeY =
      (activeRow - viewport.rowStart) * rowHeight - viewport.offsetTop;
    if (
      activeRow >= viewport.rowStart &&
      activeRow < viewport.rowEnd &&
      activeColumn >= viewport.columnStart &&
      activeColumn < viewport.columnEnd
    ) {
      context.fillStyle = 'rgba(26, 115, 232, 0.12)';
      context.fillRect(activeX, activeY, columnWidth, rowHeight);
      context.strokeStyle = '#1a73e8';
      context.lineWidth = 2;
      context.strokeRect(activeX + 1, activeY + 1, columnWidth - 2, rowHeight - 2);
    }

    context.fillStyle = '#222222';
    context.font = '12px var(--a3s-ui-font, system-ui, sans-serif)';
    context.textBaseline = 'middle';
    for (const cell of spreadsheetVirtualGridPaintCells({
      data,
      viewport,
      rowHeight,
      columnWidth,
      merges,
    })) {
      if (cell.background) {
        context.fillStyle = cell.background;
        context.fillRect(cell.x, cell.y, cell.width, cell.height);
        context.fillStyle = cell.color ?? '#222222';
      } else if (cell.color) {
        context.fillStyle = cell.color;
      } else {
        context.fillStyle = '#222222';
      }
      context.fillText(
        cell.text,
        cell.x + 6,
        cell.y + cell.height / 2,
        cell.width - 10,
      );
    }
  }, [
    activeColumn,
    activeRow,
    columnCount,
    columnWidth,
    contentHeight,
    contentWidth,
    data,
    merges,
    rowCount,
    rowHeight,
    spacerHeight,
    spacerWidth,
  ]);

  useLayoutEffect(() => {
    paint();
  }, [paint, scrollTop, scrollLeft, clientSize.height, clientSize.width]);

  useEffect(() => {
    const scroller = scrollerRef.current;
    if (!scroller) return;
    const mappedTop = mapScrollerOffset(
      scroller.scrollTop,
      spacerHeight,
      scroller.clientHeight,
      contentHeight,
    );
    const mappedLeft = mapScrollerOffset(
      scroller.scrollLeft,
      spacerWidth,
      scroller.clientWidth,
      contentWidth,
    );
    const targetTop = activeRow * rowHeight;
    const targetLeft = activeColumn * columnWidth;
    const viewBottom = mappedTop + scroller.clientHeight;
    const viewRight = mappedLeft + scroller.clientWidth;
    let nextContentTop = mappedTop;
    let nextContentLeft = mappedLeft;
    if (targetTop < mappedTop) nextContentTop = targetTop;
    else if (targetTop + rowHeight > viewBottom) {
      nextContentTop = targetTop + rowHeight - scroller.clientHeight;
    }
    if (targetLeft < mappedLeft) nextContentLeft = targetLeft;
    else if (targetLeft + columnWidth > viewRight) {
      nextContentLeft = targetLeft + columnWidth - scroller.clientWidth;
    }
    const nextTop = scrollerOffsetForContent(
      nextContentTop,
      spacerHeight,
      scroller.clientHeight,
      contentHeight,
    );
    const nextLeft = scrollerOffsetForContent(
      nextContentLeft,
      spacerWidth,
      scroller.clientWidth,
      contentWidth,
    );
    if (nextTop !== scroller.scrollTop || nextLeft !== scroller.scrollLeft) {
      scroller.scrollTo({ top: Math.max(0, nextTop), left: Math.max(0, nextLeft) });
    }
  }, [
    activeColumn,
    activeRow,
    columnWidth,
    contentHeight,
    contentWidth,
    rowHeight,
    spacerHeight,
    spacerWidth,
  ]);

  const beginEdit = (seed: string | null) => {
    if (!editable || !onCommitCell) return;
    const existing = spreadsheetVirtualGridCellText(
      data?.[activeRow]?.[activeColumn],
    );
    editingRef.current = true;
    setDraft(seed ?? existing);
  };

  const cancelEdit = () => {
    editingRef.current = false;
    setDraft(null);
    scrollerRef.current?.focus();
  };

  const commitEdit = () => {
    if (!editingRef.current || draft === null || !onCommitCell) return;
    editingRef.current = false;
    const value = draft;
    setDraft(null);
    onCommitCell(activeRow, activeColumn, value);
    scrollerRef.current?.focus();
  };

  useLayoutEffect(() => {
    if (draft === null) return;
    editorRef.current?.focus();
    editorRef.current?.select();
  }, [draft]);

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if ((event.nativeEvent as { isComposing?: boolean }).isComposing) return;
    if (draft !== null) return;
    if (editable && onCommitCell && (event.key === 'F2' || event.key === 'Enter')) {
      event.preventDefault();
      event.stopPropagation();
      beginEdit(null);
      return;
    }
    if (
      editable &&
      onCommitCell &&
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey
    ) {
      event.preventDefault();
      event.stopPropagation();
      beginEdit(event.key);
      return;
    }
    const lastRow = Math.max(0, rowCount - 1);
    const lastColumn = Math.max(0, columnCount - 1);
    let nextRow = activeRow;
    let nextColumn = activeColumn;
    switch (event.key) {
      case 'ArrowDown':
        nextRow = Math.min(lastRow, activeRow + 1);
        break;
      case 'ArrowUp':
        nextRow = Math.max(0, activeRow - 1);
        break;
      case 'ArrowRight':
        nextColumn = Math.min(lastColumn, activeColumn + 1);
        break;
      case 'ArrowLeft':
        nextColumn = Math.max(0, activeColumn - 1);
        break;
      case 'Home':
        nextColumn = 0;
        if (event.ctrlKey || event.metaKey) nextRow = 0;
        break;
      case 'End':
        nextColumn = lastColumn;
        if (event.ctrlKey || event.metaKey) nextRow = lastRow;
        break;
      default:
        return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (nextRow !== activeRow || nextColumn !== activeColumn) {
      onActiveCellChange(nextRow, nextColumn);
    }
  };

  const handleScroll = () => {
    const scroller = scrollerRef.current;
    if (!scroller) return;
    setScrollTop(scroller.scrollTop);
    setScrollLeft(scroller.scrollLeft);
  };

  const handleCanvasPointerDown = (
    event: React.PointerEvent<HTMLCanvasElement>,
  ) => {
    const scroller = scrollerRef.current;
    if (!scroller) return;
    const bounds = scroller.getBoundingClientRect();
    const mappedLeft = mapScrollerOffset(
      scroller.scrollLeft,
      spacerWidth,
      scroller.clientWidth,
      contentWidth,
    );
    const mappedTop = mapScrollerOffset(
      scroller.scrollTop,
      spacerHeight,
      scroller.clientHeight,
      contentHeight,
    );
    const x = event.clientX - bounds.left + mappedLeft;
    const y = event.clientY - bounds.top + mappedTop;
    const column = Math.min(
      Math.max(0, columnCount - 1),
      Math.floor(x / columnWidth),
    );
    const row = Math.min(
      Math.max(0, rowCount - 1),
      Math.floor(y / rowHeight),
    );
    onActiveCellChange(row, column);
    scrollerRef.current?.focus();
  };

  const handleCanvasDoubleClick = (
    event: React.PointerEvent<HTMLCanvasElement>,
  ) => {
    handleCanvasPointerDown(event);
    beginEdit(null);
  };

  const editorLeft =
    (activeColumn * columnWidth) -
    mapScrollerOffset(scrollLeft, spacerWidth, clientSize.width, contentWidth);
  const editorTop =
    (activeRow * rowHeight) -
    mapScrollerOffset(scrollTop, spacerHeight, clientSize.height, contentHeight);

  return (
    <div
      ref={scrollerRef}
      className="work-spreadsheet-virtual-grid"
      tabIndex={0}
      {...aria}
      onKeyDown={handleKeyDown}
      onScroll={handleScroll}
    >
      <div
        className="work-spreadsheet-virtual-grid-spacer"
        style={{
          width: spacerWidth,
          height: spacerHeight,
          position: 'relative',
        }}
        aria-hidden="true"
      >
        <canvas
          ref={canvasRef}
          className="work-spreadsheet-virtual-grid-canvas"
          aria-hidden="true"
          onPointerDown={handleCanvasPointerDown}
          onDoubleClick={handleCanvasDoubleClick}
          style={{
            position: 'absolute',
            top: scrollTop,
            left: scrollLeft,
            width: clientSize.width || '100%',
            height: clientSize.height || '100%',
          }}
        />
      </div>
      {draft !== null ? (
        <input
          ref={editorRef}
          className="work-spreadsheet-virtual-grid-editor"
          aria-label={`${sheetName} cell editor`}
          value={draft}
          onChange={(event) => setDraft(event.target.value)}
          onBlur={commitEdit}
          onKeyDown={(event) => {
            if ((event.nativeEvent as { isComposing?: boolean }).isComposing) {
              return;
            }
            if (event.key === 'Enter') {
              event.preventDefault();
              commitEdit();
              return;
            }
            if (event.key === 'Escape') {
              event.preventDefault();
              cancelEdit();
            }
          }}
          style={{
            left: Math.max(0, editorLeft),
            top: Math.max(0, editorTop),
            width: columnWidth,
            height: rowHeight,
          }}
        />
      ) : null}
      <div
        role="row"
        className="sr-only"
        aria-rowindex={Math.min(activeRow, Math.max(0, rowCount - 1)) + 1}
      >
        <div
          id={activeCellId}
          role="gridcell"
          aria-colindex={
            Math.min(activeColumn, Math.max(0, columnCount - 1)) + 1
          }
        >
          {spreadsheetVirtualGridCellText(data?.[activeRow]?.[activeColumn]) ||
            'blank'}
        </div>
      </div>
    </div>
  );
}
