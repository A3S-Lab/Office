import { Table2 } from 'lucide-react';
import { type KeyboardEvent, useEffect, useRef, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import type { OfficeTableDimensions } from './office-table-dimensions';

const TABLE_PICKER_ROWS = 8;
const TABLE_PICKER_COLUMNS = 10;

interface TablePickerPosition {
  row: number;
  column: number;
}

export function OfficeTableInsertPopover({
  className = '',
  disabled = false,
  label,
  onInsert,
}: {
  className?: string;
  disabled?: boolean;
  label: string;
  onInsert: (dimensions: OfficeTableDimensions) => void;
}) {
  const [open, setOpen] = useState(false);
  const [position, setPosition] = useState<TablePickerPosition>({
    row: 1,
    column: 1,
  });
  const cellRefs = useRef<Array<HTMLButtonElement | null>>([]);

  useEffect(() => {
    if (!open) return;
    const frame = requestAnimationFrame(() => {
      cellRefs.current[0]?.focus({ preventScroll: true });
    });
    return () => cancelAnimationFrame(frame);
  }, [open]);

  const focusCell = (next: TablePickerPosition): void => {
    setPosition(next);
    cellRefs.current[cellIndex(next)]?.focus({ preventScroll: true });
  };
  const insert = (dimensions: TablePickerPosition, close: () => void): void => {
    close();
    onInsert({ rows: dimensions.row, columns: dimensions.column });
  };
  const handleCellKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    current: TablePickerPosition,
    close: () => void,
  ): void => {
    let next = current;
    if (event.key === 'ArrowLeft') {
      next = { ...current, column: Math.max(1, current.column - 1) };
    } else if (event.key === 'ArrowRight') {
      next = {
        ...current,
        column: Math.min(TABLE_PICKER_COLUMNS, current.column + 1),
      };
    } else if (event.key === 'ArrowUp') {
      next = { ...current, row: Math.max(1, current.row - 1) };
    } else if (event.key === 'ArrowDown') {
      next = {
        ...current,
        row: Math.min(TABLE_PICKER_ROWS, current.row + 1),
      };
    } else if (event.key === 'Home') {
      next = { ...current, column: 1 };
    } else if (event.key === 'End') {
      next = { ...current, column: TABLE_PICKER_COLUMNS };
    } else if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      insert(current, close);
      return;
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    focusCell(next);
  };

  return (
    <Popover
      label={label}
      panelLabel="选择表格大小"
      panelRole="dialog"
      portal
      open={open}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (nextOpen) setPosition({ row: 1, column: 1 });
      }}
      className={`work-office-table-insert-popover ${className}`.trim()}
      panelClassName="work-office-table-picker-panel"
      trigger={(triggerProps, { open: triggerOpen }) => (
        <button
          {...triggerProps}
          className={`with-label${triggerOpen ? ' active' : ''}`}
          aria-pressed={triggerOpen}
          disabled={disabled}
        >
          <Table2 size={19} />
          <span>{label}</span>
        </button>
      )}
    >
      {(close) => (
        <>
          <table className="work-office-table-picker-grid">
            <caption className="sr-only">选择表格行列</caption>
            <tbody>
              {Array.from({ length: TABLE_PICKER_ROWS }, (_, rowIndex) => (
                <tr key={rowIndex}>
                  {Array.from(
                    { length: TABLE_PICKER_COLUMNS },
                    (_, columnIndex) => {
                      const current = {
                        row: rowIndex + 1,
                        column: columnIndex + 1,
                      };
                      const selected =
                        current.row === position.row &&
                        current.column === position.column;
                      const highlighted =
                        current.row <= position.row &&
                        current.column <= position.column;
                      const index = cellIndex(current);
                      return (
                        <td key={`${current.row}-${current.column}`}>
                          <button
                            ref={(element) => {
                              cellRefs.current[index] = element;
                            }}
                            type="button"
                            aria-label={`${current.row} 行 ${current.column} 列`}
                            aria-pressed={selected}
                            data-highlighted={highlighted ? 'true' : undefined}
                            tabIndex={selected ? 0 : -1}
                            onFocus={() => setPosition(current)}
                            onMouseEnter={() => setPosition(current)}
                            onKeyDown={(event) =>
                              handleCellKeyDown(event, current, close)
                            }
                            onClick={() => insert(current, close)}
                          />
                        </td>
                      );
                    },
                  )}
                </tr>
              ))}
            </tbody>
          </table>
          <output className="work-office-table-picker-size" aria-live="polite">
            {position.row} × {position.column} 表格
          </output>
        </>
      )}
    </Popover>
  );
}

function cellIndex(position: TablePickerPosition): number {
  return (position.row - 1) * TABLE_PICKER_COLUMNS + (position.column - 1);
}
