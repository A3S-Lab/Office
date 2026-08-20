import { Check, Minus } from 'lucide-react';
import { type ReactNode, useEffect, useRef } from 'react';

export function OfficeCheckbox({
  ariaLabel,
  checked,
  onCheckedChange,
  children,
  disabled = false,
  className = '',
  indeterminate = false,
}: {
  ariaLabel: string;
  checked: boolean;
  onCheckedChange: (checked: boolean) => void;
  children: ReactNode;
  disabled?: boolean;
  className?: string;
  indeterminate?: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (inputRef.current) inputRef.current.indeterminate = indeterminate;
  }, [indeterminate]);
  return (
    <label
      className={`work-office-checkbox ${checked ? 'checked' : ''} ${indeterminate ? 'indeterminate' : ''} ${disabled ? 'disabled' : ''} ${className}`.trim()}
    >
      <input
        ref={inputRef}
        type="checkbox"
        aria-label={ariaLabel}
        aria-checked={indeterminate ? 'mixed' : checked}
        checked={checked}
        disabled={disabled}
        onChange={(event) => onCheckedChange(event.target.checked)}
      />
      <span className="work-office-checkbox-box" aria-hidden="true">
        {indeterminate ? <Minus size={11} /> : checked && <Check size={11} />}
      </span>
      <span className="work-office-checkbox-label">{children}</span>
    </label>
  );
}
