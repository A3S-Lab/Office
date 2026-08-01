import {
  forwardRef,
  type InputHTMLAttributes,
  type TextareaHTMLAttributes,
  useEffect,
  useState,
} from 'react';

export const OfficeTextField = forwardRef<
  HTMLInputElement,
  Omit<InputHTMLAttributes<HTMLInputElement>, 'type'> & {
    type?: 'text' | 'search' | 'password';
  }
>(function OfficeTextField({ className = '', type = 'text', ...props }, ref) {
  return (
    <input
      ref={ref}
      type={type}
      className={`work-office-text-field ${className}`.trim()}
      {...props}
    />
  );
});

export const OfficeTextArea = forwardRef<
  HTMLTextAreaElement,
  TextareaHTMLAttributes<HTMLTextAreaElement>
>(function OfficeTextArea({ className = '', ...props }, ref) {
  return (
    <textarea
      ref={ref}
      className={`work-office-text-area ${className}`.trim()}
      {...props}
    />
  );
});

export function CommittedOfficeTextField<Value>({
  value,
  formatValue,
  parseValue,
  onValueCommit,
  ...fieldProps
}: Omit<
  InputHTMLAttributes<HTMLInputElement>,
  | 'aria-invalid'
  | 'defaultValue'
  | 'onBlur'
  | 'onChange'
  | 'onKeyDown'
  | 'type'
  | 'value'
> & {
  value: Value;
  formatValue: (value: Value) => string;
  parseValue: (draft: string) => Value | null;
  onValueCommit: (value: Value) => void;
}) {
  const { dirty, draft, invalid, commit, reset, setDraft } =
    useCommittedOfficeTextDraft({
      value,
      formatValue,
      parseValue,
      onValueCommit,
    });

  return (
    <OfficeTextField
      {...fieldProps}
      aria-invalid={invalid || undefined}
      data-office-escape-consumer={dirty || undefined}
      value={draft}
      onChange={(event) => setDraft(event.currentTarget.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Escape' && dirty) {
          event.preventDefault();
          event.stopPropagation();
          reset();
        } else if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
          event.preventDefault();
          event.stopPropagation();
          commit();
        }
      }}
    />
  );
}

export function CommittedOfficeTextArea<Value>({
  value,
  formatValue,
  parseValue,
  onValueCommit,
  ...textAreaProps
}: Omit<
  TextareaHTMLAttributes<HTMLTextAreaElement>,
  | 'aria-invalid'
  | 'defaultValue'
  | 'onBlur'
  | 'onChange'
  | 'onKeyDown'
  | 'value'
> & {
  value: Value;
  formatValue: (value: Value) => string;
  parseValue: (draft: string) => Value | null;
  onValueCommit: (value: Value) => void;
}) {
  const { dirty, draft, invalid, commit, reset, setDraft } =
    useCommittedOfficeTextDraft({
      value,
      formatValue,
      parseValue,
      onValueCommit,
    });

  return (
    <OfficeTextArea
      {...textAreaProps}
      aria-invalid={invalid || undefined}
      data-office-escape-consumer={dirty || undefined}
      value={draft}
      onChange={(event) => setDraft(event.currentTarget.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Escape' && dirty) {
          event.preventDefault();
          event.stopPropagation();
          reset();
        } else if (
          event.key === 'Enter' &&
          (event.metaKey || event.ctrlKey) &&
          !event.nativeEvent.isComposing
        ) {
          event.preventDefault();
          event.stopPropagation();
          commit();
        }
      }}
    />
  );
}

function useCommittedOfficeTextDraft<Value>({
  value,
  formatValue,
  parseValue,
  onValueCommit,
}: {
  value: Value;
  formatValue: (value: Value) => string;
  parseValue: (draft: string) => Value | null;
  onValueCommit: (value: Value) => void;
}) {
  const controlledDraft = formatValue(value);
  const [draft, setDraft] = useState(controlledDraft);

  useEffect(() => setDraft(controlledDraft), [controlledDraft]);

  const parsed = safelyParseOfficeTextDraft(parseValue, draft);
  const commit = () => {
    if (parsed === null) {
      setDraft(controlledDraft);
      return;
    }
    const normalizedDraft = formatValue(parsed);
    setDraft(normalizedDraft);
    if (normalizedDraft !== controlledDraft) onValueCommit(parsed);
  };

  return {
    dirty: draft !== controlledDraft,
    draft,
    invalid: parsed === null,
    commit,
    reset: () => setDraft(controlledDraft),
    setDraft,
  };
}

function safelyParseOfficeTextDraft<Value>(
  parseValue: (draft: string) => Value | null,
  draft: string,
): Value | null {
  try {
    return parseValue(draft);
  } catch {
    return null;
  }
}

export const OfficeFileInput = forwardRef<
  HTMLInputElement,
  Omit<InputHTMLAttributes<HTMLInputElement>, 'type' | 'onChange'> & {
    onFileSelect: (file: File) => unknown;
  }
>(function OfficeFileInput({ className = '', onFileSelect, ...props }, ref) {
  return (
    <input
      ref={ref}
      type="file"
      className={`work-file-input ${className}`.trim()}
      onChange={(event) => {
        const input = event.currentTarget;
        const file = input.files?.[0];
        try {
          if (file) onFileSelect(file);
        } finally {
          input.value = '';
        }
      }}
      {...props}
    />
  );
});
