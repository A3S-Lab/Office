import { ChevronDown, ChevronUp } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

export function OfficeNumberField({
  ariaLabel,
  value,
  onValueChange,
  onCommit,
  onCancel,
  min,
  max,
  step = 1,
  disabled = false,
  className = '',
  placeholder,
  escapeConsumer = false,
  validationInvalid = false,
}: {
  ariaLabel: string;
  value: number | string;
  onValueChange: (value: string) => void;
  onCommit?: (value: string) => void;
  onCancel?: () => void;
  min?: number;
  max?: number;
  step?: number;
  disabled?: boolean;
  className?: string;
  placeholder?: string;
  escapeConsumer?: boolean;
  validationInvalid?: boolean;
}) {
  const commitInProgressRef = useRef(false);
  const numericValue = value === '' ? null : Number(value);
  const invalid =
    validationInvalid ||
    (numericValue !== null &&
      (!Number.isFinite(numericValue) ||
        (min !== undefined && numericValue < min) ||
        (max !== undefined && numericValue > max)));
  const changeBy = (direction: -1 | 1) => {
    const current = Number(value);
    const fallback = min ?? 0;
    const next = clampNumber(
      Number.isFinite(current) ? current + step * direction : fallback,
      min,
      max,
    );
    const formatted = formatSteppedNumber(next, step);
    onValueChange(formatted);
    onCommit?.(formatted);
  };

  return (
    <div className={`work-office-number-field ${className}`.trim()}>
      <input
        type="text"
        inputMode="decimal"
        aria-label={ariaLabel}
        aria-invalid={invalid || undefined}
        data-office-escape-consumer={escapeConsumer || undefined}
        value={value}
        disabled={disabled}
        placeholder={placeholder}
        onChange={(event) => onValueChange(event.target.value)}
        onBlur={(event) => {
          if (commitInProgressRef.current) return;
          onCommit?.(event.currentTarget.value);
        }}
        onKeyDown={(event) => {
          if (event.key === 'ArrowUp') {
            event.preventDefault();
            changeBy(1);
          } else if (event.key === 'ArrowDown') {
            event.preventDefault();
            changeBy(-1);
          } else if (
            event.key === 'Enter' &&
            onCommit &&
            !event.nativeEvent.isComposing
          ) {
            event.preventDefault();
            event.stopPropagation();
            commitInProgressRef.current = true;
            try {
              onCommit(event.currentTarget.value);
            } finally {
              commitInProgressRef.current = false;
            }
          } else if (event.key === 'Escape' && escapeConsumer && onCancel) {
            event.preventDefault();
            event.stopPropagation();
            onCancel();
          }
        }}
      />
      <span className="work-office-number-steppers">
        <button
          type="button"
          aria-label={`增加${ariaLabel}`}
          disabled={disabled}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => changeBy(1)}
        >
          <ChevronUp size={10} />
        </button>
        <button
          type="button"
          aria-label={`减少${ariaLabel}`}
          disabled={disabled}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => changeBy(-1)}
        >
          <ChevronDown size={10} />
        </button>
      </span>
    </div>
  );
}

export function CommittedOfficeNumberField<
  Value extends number | undefined = number,
>({
  value,
  normalizeValue,
  onValueCommit,
  formatValue = defaultNumberFormat,
  ...fieldProps
}: Omit<
  Parameters<typeof OfficeNumberField>[0],
  | 'escapeConsumer'
  | 'onCancel'
  | 'onCommit'
  | 'onValueChange'
  | 'validationInvalid'
  | 'value'
> & {
  value: Value;
  normalizeValue: (value: string) => Value | null;
  onValueCommit: (value: Value) => void;
  formatValue?: (value: number) => string;
}) {
  const controlledDraft =
    value === undefined ? '' : formatValue(value as number);
  const [draft, setDraft] = useState(controlledDraft);
  const normalizedDraft = safelyNormalizeOfficeNumberDraft(
    normalizeValue,
    draft,
  );

  useEffect(() => setDraft(controlledDraft), [controlledDraft]);

  const commit = (rawValue: string) => {
    const normalized = safelyNormalizeOfficeNumberDraft(
      normalizeValue,
      rawValue,
    );
    if (
      normalized === null ||
      (normalized !== undefined && !Number.isFinite(normalized))
    ) {
      setDraft(controlledDraft);
      return;
    }
    const normalizedDraft =
      normalized === undefined ? '' : formatValue(normalized as number);
    setDraft(normalizedDraft);
    if (!Object.is(normalized, value)) onValueCommit(normalized);
  };

  return (
    <OfficeNumberField
      {...fieldProps}
      escapeConsumer={draft !== controlledDraft}
      validationInvalid={normalizedDraft === null}
      value={draft}
      onValueChange={setDraft}
      onCommit={commit}
      onCancel={() => setDraft(controlledDraft)}
    />
  );
}

function clampNumber(value: number, min?: number, max?: number): number {
  return Math.min(
    max ?? Number.POSITIVE_INFINITY,
    Math.max(min ?? Number.NEGATIVE_INFINITY, value),
  );
}

function formatSteppedNumber(value: number, step: number): string {
  const decimalPlaces = Math.max(0, (String(step).split('.')[1] ?? '').length);
  return decimalPlaces === 0
    ? String(Math.round(value))
    : value.toFixed(decimalPlaces).replace(/\.?0+$/, '');
}

function defaultNumberFormat(value: number): string {
  return String(value);
}

function safelyNormalizeOfficeNumberDraft<Value>(
  normalizeValue: (value: string) => Value | null,
  draft: string,
): Value | null {
  try {
    return normalizeValue(draft);
  } catch {
    return null;
  }
}
