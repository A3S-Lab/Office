import { Minus, Plus, Table2 } from 'lucide-react';
import { type KeyboardEvent, useEffect, useRef, useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import type { OfficeTableDimensions } from './office-table-dimensions';

const TABLE_PICKER_ROWS = 8;
const TABLE_PICKER_COLUMNS = 10;
const TABLE_PICKER_COMPACT_QUERY = '(max-width: 520px)';

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
  const compact = useCompactTablePicker();
  const cellRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const compactRowInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!open) return;
    const frame = requestAnimationFrame(() => {
      if (compact) {
        compactRowInputRef.current?.focus({ preventScroll: true });
        compactRowInputRef.current?.select();
      } else {
        cellRefs.current[0]?.focus({ preventScroll: true });
      }
    });
    return () => cancelAnimationFrame(frame);
  }, [compact, open]);

  const focusCell = (next: TablePickerPosition): void => {
    setPosition(next);
    cellRefs.current[cellIndex(next)]?.focus({ preventScroll: true });
  };
  const insert = (dimensions: TablePickerPosition, close: () => void): void => {
    close();
    onInsert({ rows: dimensions.row, columns: dimensions.column });
  };
  const updateDimension = (
    dimension: keyof TablePickerPosition,
    value: number,
  ): void => {
    if (!Number.isFinite(value)) return;
    const maximum =
      dimension === 'row' ? TABLE_PICKER_ROWS : TABLE_PICKER_COLUMNS;
    setPosition((current) => ({
      ...current,
      [dimension]: clampDimension(value, maximum),
    }));
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
      {(close) =>
        compact ? (
          <form
            className="work-office-table-picker-compact"
            onSubmit={(event) => {
              event.preventDefault();
              insert(position, close);
            }}
          >
            <TableDimensionField
              inputRef={compactRowInputRef}
              label="行数"
              maximum={TABLE_PICKER_ROWS}
              value={position.row}
              onChange={(value) => updateDimension('row', value)}
            />
            <TableDimensionField
              label="列数"
              maximum={TABLE_PICKER_COLUMNS}
              value={position.column}
              onChange={(value) => updateDimension('column', value)}
            />
            <button type="submit" className="work-office-table-picker-submit">
              插入 {position.row} × {position.column} 表格
            </button>
          </form>
        ) : (
          <>
            <fieldset className="work-office-table-picker-grid">
              <legend className="sr-only">选择表格行列</legend>
              {Array.from(
                { length: TABLE_PICKER_ROWS * TABLE_PICKER_COLUMNS },
                (_, index) => {
                  const current = {
                    row: Math.floor(index / TABLE_PICKER_COLUMNS) + 1,
                    column: (index % TABLE_PICKER_COLUMNS) + 1,
                  };
                  const selected =
                    current.row === position.row &&
                    current.column === position.column;
                  const highlighted =
                    current.row <= position.row &&
                    current.column <= position.column;
                  return (
                    <button
                      key={`${current.row}-${current.column}`}
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
                  );
                },
              )}
            </fieldset>
            <output
              className="work-office-table-picker-size"
              aria-live="polite"
            >
              {position.row} × {position.column} 表格
            </output>
          </>
        )
      }
    </Popover>
  );
}

function TableDimensionField({
  inputRef,
  label,
  maximum,
  value,
  onChange,
}: {
  inputRef?: React.RefObject<HTMLInputElement | null>;
  label: string;
  maximum: number;
  value: number;
  onChange: (value: number) => void;
}) {
  return (
    <fieldset className="work-office-table-picker-dimension">
      <legend>{label}</legend>
      <div className="work-office-table-picker-stepper">
        <button
          type="button"
          aria-label={`减少${label}`}
          disabled={value <= 1}
          onClick={() => onChange(value - 1)}
        >
          <Minus size={16} aria-hidden="true" />
        </button>
        <input
          ref={inputRef}
          type="number"
          aria-label={label}
          inputMode="numeric"
          min={1}
          max={maximum}
          value={value}
          onFocus={(event) => event.currentTarget.select()}
          onChange={(event) => onChange(event.currentTarget.valueAsNumber)}
        />
        <button
          type="button"
          aria-label={`增加${label}`}
          disabled={value >= maximum}
          onClick={() => onChange(value + 1)}
        >
          <Plus size={16} aria-hidden="true" />
        </button>
      </div>
    </fieldset>
  );
}

function cellIndex(position: TablePickerPosition): number {
  return (position.row - 1) * TABLE_PICKER_COLUMNS + (position.column - 1);
}

function clampDimension(value: number, maximum: number): number {
  return Math.min(maximum, Math.max(1, Math.round(value)));
}

function useCompactTablePicker(): boolean {
  const [compact, setCompact] = useState(compactTablePickerMatches);

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      typeof window.matchMedia !== 'function'
    ) {
      return;
    }
    const mediaQuery = window.matchMedia(TABLE_PICKER_COMPACT_QUERY);
    const update = () => setCompact(mediaQuery.matches);
    update();
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', update);
      return () => mediaQuery.removeEventListener('change', update);
    }
    mediaQuery.addListener(update);
    return () => mediaQuery.removeListener(update);
  }, []);

  return compact;
}

function compactTablePickerMatches(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia(TABLE_PICKER_COMPACT_QUERY).matches
  );
}
