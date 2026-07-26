import { Check, ChevronDown } from 'lucide-react';
import {
  type KeyboardEvent,
  type ReactNode,
  useEffect,
  useRef,
  useState,
} from 'react';
import { Button, Popover } from '../../../design-system/primitives';

const THEME_COLORS = [
  '#111827',
  '#1f2937',
  '#374151',
  '#6b7280',
  '#9ca3af',
  '#d1d5db',
  '#e5e7eb',
  '#f3f4f6',
  '#f9fafb',
  '#ffffff',
  '#f4cccc',
  '#fce5cd',
  '#fff2cc',
  '#d9ead3',
  '#d0e0e3',
  '#c9daf8',
  '#cfe2f3',
  '#d9d2e9',
  '#ead1dc',
  '#fce4ec',
  '#ea9999',
  '#f9cb9c',
  '#ffe599',
  '#b6d7a8',
  '#a2c4c9',
  '#a4c2f4',
  '#9fc5e8',
  '#b4a7d6',
  '#d5a6bd',
  '#f4b6c2',
  '#e06666',
  '#f6b26b',
  '#ffd966',
  '#93c47d',
  '#76a5af',
  '#6d9eeb',
  '#6fa8dc',
  '#8e7cc3',
  '#c27ba0',
  '#e67c99',
  '#990000',
  '#b45f06',
  '#bf9000',
  '#38761d',
  '#134f5c',
  '#1155cc',
  '#0b5394',
  '#351c75',
  '#741b47',
  '#a61c4c',
] as const;

const STANDARD_COLORS = [
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#7030a0',
  '#ff00ff',
] as const;

const OFFICE_COLORS: readonly string[] = [...THEME_COLORS, ...STANDARD_COLORS];
const COLOR_GRID_COLUMNS = 10;

export function OfficeColorPicker({
  ariaLabel,
  value,
  onValueChange,
  disabled = false,
  compact = false,
  className = '',
  triggerLabel,
  triggerIcon,
}: {
  ariaLabel: string;
  value: string;
  onValueChange: (value: string) => void;
  disabled?: boolean;
  compact?: boolean;
  className?: string;
  triggerLabel?: string;
  triggerIcon?: ReactNode;
}) {
  const [draft, setDraft] = useState(value);
  const [open, setOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(() =>
    Math.max(0, OFFICE_COLORS.indexOf(value.toLowerCase())),
  );
  const colorRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const normalizedDraft = normalizeCssColor(draft);

  useEffect(() => setDraft(value), [value]);
  useEffect(() => {
    if (!open) return;
    const selectedIndex = OFFICE_COLORS.indexOf(value.toLowerCase());
    const nextIndex = Math.max(0, selectedIndex);
    setActiveIndex(nextIndex);
    const frame = requestAnimationFrame(() =>
      colorRefs.current[nextIndex]?.focus({ preventScroll: true }),
    );
    return () => cancelAnimationFrame(frame);
  }, [open, value]);

  const choose = (color: string, close: () => void) => {
    setDraft(color);
    close();
    onValueChange(color);
  };

  const applyDraft = (close: () => void) => {
    if (normalizedDraft) choose(normalizedDraft, close);
  };

  const moveColorFocus = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    let nextIndex = index;
    if (event.key === 'ArrowLeft') nextIndex = Math.max(0, index - 1);
    else if (event.key === 'ArrowRight')
      nextIndex = Math.min(OFFICE_COLORS.length - 1, index + 1);
    else if (event.key === 'ArrowUp')
      nextIndex = Math.max(0, index - COLOR_GRID_COLUMNS);
    else if (event.key === 'ArrowDown')
      nextIndex = Math.min(
        OFFICE_COLORS.length - 1,
        index + COLOR_GRID_COLUMNS,
      );
    else if (event.key === 'Home') nextIndex = 0;
    else if (event.key === 'End') nextIndex = OFFICE_COLORS.length - 1;
    else return;
    event.preventDefault();
    event.stopPropagation();
    setActiveIndex(nextIndex);
    colorRefs.current[nextIndex]?.focus({ preventScroll: true });
  };

  const colorOption = (color: string, index: number, close: () => void) => (
    <button
      ref={(element) => {
        colorRefs.current[index] = element;
      }}
      type="button"
      role="option"
      aria-label={`颜色 ${color}`}
      aria-selected={color.toLowerCase() === value.toLowerCase()}
      tabIndex={index === activeIndex ? 0 : -1}
      key={color}
      style={{ backgroundColor: color, color: swatchForeground(color) }}
      onFocus={() => setActiveIndex(index)}
      onKeyDown={(event) => moveColorFocus(event, index)}
      onClick={() => choose(color, close)}
    >
      {color.toLowerCase() === value.toLowerCase() && (
        <Check size={13} aria-hidden="true" />
      )}
    </button>
  );

  return (
    <Popover
      label={ariaLabel}
      panelLabel={ariaLabel}
      panelRole="dialog"
      placement="bottom-end"
      portal
      className={`work-office-color-picker${compact ? ' compact' : ''}${triggerLabel ? ' labeled' : ''}${className ? ` ${className}` : ''}`}
      panelClassName="work-office-color-menu"
      disabled={disabled}
      open={open}
      onOpenChange={setOpen}
      trigger={(triggerProps) => (
        <button
          {...triggerProps}
          className="work-office-color-trigger"
          title={ariaLabel}
        >
          {triggerIcon ? (
            <span className="work-office-color-trigger-icon" aria-hidden="true">
              {triggerIcon}
            </span>
          ) : null}
          <span
            className="work-office-color-swatch"
            style={{ backgroundColor: value }}
            aria-hidden="true"
          />
          {triggerLabel ? (
            <span className="work-office-color-trigger-label">
              {triggerLabel}
            </span>
          ) : !compact ? (
            <span className="work-office-color-value">
              {value.toUpperCase()}
            </span>
          ) : null}
          <ChevronDown size={12} aria-hidden="true" />
        </button>
      )}
    >
      {(close) => (
        <>
          <div
            className="work-office-color-palette"
            role="listbox"
            aria-label="颜色"
          >
            <span className="work-office-color-section-title">主题颜色</span>
            <div className="work-office-color-grid theme">
              {THEME_COLORS.map((color, index) =>
                colorOption(color, index, close),
              )}
            </div>
            <span className="work-office-color-section-title">标准色</span>
            <div className="work-office-color-grid standard">
              {STANDARD_COLORS.map((color, index) =>
                colorOption(color, THEME_COLORS.length + index, close),
              )}
            </div>
          </div>
          <div className="work-office-color-custom">
            <span
              className="work-office-color-custom-preview"
              data-testid="custom-color-preview"
              style={{ backgroundColor: normalizedDraft ?? value }}
              aria-hidden="true"
            />
            <label>
              <span>自定义颜色</span>
              <input
                type="text"
                aria-label="自定义颜色值"
                aria-invalid={Boolean(draft.trim() && !normalizedDraft)}
                value={draft}
                spellCheck={false}
                onChange={(event) => setDraft(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    applyDraft(close);
                  }
                }}
              />
            </label>
            <Button
              size="compact"
              tone="primary"
              aria-label="应用自定义颜色"
              disabled={!normalizedDraft}
              onClick={() => applyDraft(close)}
            >
              应用
            </Button>
          </div>
        </>
      )}
    </Popover>
  );
}

function normalizeCssColor(value: string): string | null {
  const trimmed = value.trim();
  if (/^#[0-9a-f]{3}$/i.test(trimmed)) {
    return `#${trimmed
      .slice(1)
      .split('')
      .map((part) => `${part}${part}`)
      .join('')}`.toLowerCase();
  }
  if (/^#[0-9a-f]{6}$/i.test(trimmed)) return trimmed.toLowerCase();
  return null;
}

function swatchForeground(color: string): '#111827' | '#ffffff' {
  const red = Number.parseInt(color.slice(1, 3), 16);
  const green = Number.parseInt(color.slice(3, 5), 16);
  const blue = Number.parseInt(color.slice(5, 7), 16);
  return red * 0.299 + green * 0.587 + blue * 0.114 > 176
    ? '#111827'
    : '#ffffff';
}
