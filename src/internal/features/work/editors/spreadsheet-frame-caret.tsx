import { useLayoutEffect, useRef, useState } from 'react';
import { spreadsheetCellAt } from './spreadsheet-editor-support';
import {
  spreadsheetCellLayoutBox,
  spreadsheetFrameCaretPaint,
  type SpreadsheetFrameCell,
} from './spreadsheet-collaboration-presence';
import type { WorkSpreadsheetContent } from '../work-types';

export function SpreadsheetFrameCaretMark({
  cell,
  columnHeaderHeight,
  content,
  rowHeaderWidth,
}: {
  cell: SpreadsheetFrameCell;
  columnHeaderHeight: number;
  content: WorkSpreadsheetContent;
  rowHeaderWidth: number;
}) {
  const probeRef = useRef<HTMLElement>(null);
  const [scroll, setScroll] = useState({ left: 0, top: 0 });
  useLayoutEffect(() => {
    const host = probeRef.current?.parentElement;
    if (!host) return;
    const horizontal = host.querySelector('.luckysheet-scrollbar-x');
    const vertical = host.querySelector('.luckysheet-scrollbar-y');
    const read = () => {
      const left = horizontal instanceof HTMLElement ? horizontal.scrollLeft : 0;
      const top = vertical instanceof HTMLElement ? vertical.scrollTop : 0;
      setScroll((current) =>
        current.left === left && current.top === top ? current : { left, top },
      );
    };
    read();
    horizontal?.addEventListener('scroll', read);
    vertical?.addEventListener('scroll', read);
    return () => {
      horizontal?.removeEventListener('scroll', read);
      vertical?.removeEventListener('scroll', read);
    };
  }, [cell.column, cell.indexUtf16, cell.row, cell.sheetId]);
  const sheet = content.sheets.find((candidate) => candidate.id === cell.sheetId);
  const box =
    cell.indexUtf16 === undefined
      ? null
      : spreadsheetCellLayoutBox({
          column: cell.column,
          columnHeaderHeight,
          row: cell.row,
          rowHeaderWidth,
          scrollLeft: scroll.left,
          scrollTop: scroll.top,
          sheet,
        });
  const source = spreadsheetCellAt(sheet, cell.row, cell.column);
  const fontSize = typeof source?.fs === 'number' && source.fs > 0 ? source.fs : 11;
  const paint =
    box &&
    box.left >= rowHeaderWidth &&
    box.top >= columnHeaderHeight
      ? spreadsheetFrameCaretPaint({
          box,
          cell,
          content,
          layoutSettled: true,
          measure: (slice) => measureSpreadsheetText(slice, fontSize),
        })
      : null;
  return (
    <i
      className="work-spreadsheet-frame-caret"
      data-frame-caret-index={paint ? paint.head : undefined}
      ref={probeRef}
      style={
        paint
          ? {
              height: paint.height,
              left: paint.left,
              top: paint.top,
              width: paint.width,
            }
          : { display: 'none' }
      }
    />
  );
}

function measureSpreadsheetText(slice: string, fontSize: number): number {
  if (typeof document === 'undefined') return slice.length * fontSize;
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  if (!context) return slice.length * fontSize;
  context.font = `${fontSize}px sans-serif`;
  return context.measureText(slice).width;
}
