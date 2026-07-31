import {
  type InputHTMLAttributes,
  type ReactNode,
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
} from 'react';
import {
  Button,
  type ButtonTone,
  Dialog,
} from '../../../design-system/primitives';
import { OfficeTextArea, OfficeTextField } from './office-text-field';

interface OfficePromptRequest {
  id: number;
  kind: 'prompt';
  title: string;
  description?: string;
  value: string;
  fieldLabel: string;
  placeholder?: string;
  multiline?: boolean;
  inputMode?: InputHTMLAttributes<HTMLInputElement>['inputMode'];
  confirmLabel: string;
  requiredMessage?: string;
  validate?: (value: string) => string | null;
  restoreFocusTarget?: () => HTMLElement | null;
  touched: boolean;
}

interface OfficeNoticeRequest {
  id: number;
  kind: 'notice';
  title: string;
  description?: string;
  confirmLabel: string;
  restoreFocusTarget?: () => HTMLElement | null;
}

interface OfficeConfirmRequest {
  id: number;
  kind: 'confirm';
  title: string;
  description?: string;
  confirmLabel: string;
  cancelLabel: string;
  confirmTone: ButtonTone;
  restoreFocusTarget?: () => HTMLElement | null;
}

type OfficeDialogRequest =
  | OfficePromptRequest
  | OfficeNoticeRequest
  | OfficeConfirmRequest;

export interface OfficePromptOptions {
  title: string;
  description?: string;
  initialValue?: string;
  fieldLabel?: string;
  placeholder?: string;
  multiline?: boolean;
  inputMode?: InputHTMLAttributes<HTMLInputElement>['inputMode'];
  confirmLabel?: string;
  required?: boolean | string;
  validate?: (value: string) => string | null;
  restoreFocusTarget?: () => HTMLElement | null;
}

export interface OfficeNoticeOptions {
  title: string;
  description?: string;
  confirmLabel?: string;
  restoreFocusTarget?: () => HTMLElement | null;
}

export interface OfficeConfirmOptions {
  title: string;
  description?: string;
  confirmLabel?: string;
  cancelLabel?: string;
  confirmTone?: ButtonTone;
  restoreFocusTarget?: () => HTMLElement | null;
}

export function useOfficeDialog(): {
  prompt: (options: OfficePromptOptions) => Promise<string | null>;
  notice: (options: OfficeNoticeOptions) => Promise<void>;
  confirm: (options: OfficeConfirmOptions) => Promise<boolean>;
  dialog: ReactNode;
} {
  const [request, setRequest] = useState<OfficeDialogRequest | null>(null);
  const sequence = useRef(0);
  const promptResolver = useRef<((value: string | null) => void) | null>(null);
  const noticeResolver = useRef<(() => void) | null>(null);
  const confirmResolver = useRef<((confirmed: boolean) => void) | null>(null);
  const invokerRef = useRef<HTMLElement | null>(null);
  const releaseInvokerTimer = useRef<number | null>(null);
  const promptFieldId = useId();

  const retainInvoker = useCallback(() => {
    if (releaseInvokerTimer.current !== null)
      window.clearTimeout(releaseInvokerTimer.current);
    releaseInvokerTimer.current = null;
    if (!invokerRef.current?.isConnected) {
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : null;
    }
  }, []);
  const releaseInvoker = useCallback(() => {
    if (releaseInvokerTimer.current !== null)
      window.clearTimeout(releaseInvokerTimer.current);
    releaseInvokerTimer.current = window.setTimeout(() => {
      invokerRef.current = null;
      releaseInvokerTimer.current = null;
    }, 0);
  }, []);

  const closePrompt = useCallback(
    (value: string | null) => {
      promptResolver.current?.(value);
      promptResolver.current = null;
      setRequest(null);
      releaseInvoker();
    },
    [releaseInvoker],
  );
  const closeNotice = useCallback(() => {
    noticeResolver.current?.();
    noticeResolver.current = null;
    setRequest(null);
    releaseInvoker();
  }, [releaseInvoker]);
  const closeConfirm = useCallback(
    (confirmed: boolean) => {
      confirmResolver.current?.(confirmed);
      confirmResolver.current = null;
      setRequest(null);
      releaseInvoker();
    },
    [releaseInvoker],
  );

  const prompt = useCallback(
    (options: OfficePromptOptions) =>
      new Promise<string | null>((resolve) => {
        retainInvoker();
        promptResolver.current?.(null);
        noticeResolver.current?.();
        confirmResolver.current?.(false);
        promptResolver.current = resolve;
        noticeResolver.current = null;
        confirmResolver.current = null;
        setRequest({
          id: ++sequence.current,
          kind: 'prompt',
          title: options.title,
          description: options.description,
          value: options.initialValue ?? '',
          fieldLabel: options.fieldLabel ?? options.title,
          placeholder: options.placeholder,
          multiline: options.multiline,
          inputMode: options.inputMode,
          confirmLabel: options.confirmLabel ?? '确定',
          requiredMessage:
            options.required === true
              ? '请填写此项。'
              : typeof options.required === 'string'
                ? options.required
                : undefined,
          validate: options.validate,
          restoreFocusTarget: options.restoreFocusTarget,
          touched: false,
        });
      }),
    [retainInvoker],
  );

  const notice = useCallback(
    (options: OfficeNoticeOptions) =>
      new Promise<void>((resolve) => {
        retainInvoker();
        promptResolver.current?.(null);
        noticeResolver.current?.();
        confirmResolver.current?.(false);
        promptResolver.current = null;
        noticeResolver.current = resolve;
        confirmResolver.current = null;
        setRequest({
          id: ++sequence.current,
          kind: 'notice',
          title: options.title,
          description: options.description,
          confirmLabel: options.confirmLabel ?? '知道了',
          restoreFocusTarget: options.restoreFocusTarget,
        });
      }),
    [retainInvoker],
  );
  const confirm = useCallback(
    (options: OfficeConfirmOptions) =>
      new Promise<boolean>((resolve) => {
        retainInvoker();
        promptResolver.current?.(null);
        noticeResolver.current?.();
        confirmResolver.current?.(false);
        promptResolver.current = null;
        noticeResolver.current = null;
        confirmResolver.current = resolve;
        setRequest({
          id: ++sequence.current,
          kind: 'confirm',
          title: options.title,
          description: options.description,
          confirmLabel: options.confirmLabel ?? '继续',
          cancelLabel: options.cancelLabel ?? '取消',
          confirmTone: options.confirmTone ?? 'primary',
          restoreFocusTarget: options.restoreFocusTarget,
        });
      }),
    [retainInvoker],
  );

  useEffect(
    () => () => {
      promptResolver.current?.(null);
      noticeResolver.current?.();
      confirmResolver.current?.(false);
      if (releaseInvokerTimer.current !== null)
        window.clearTimeout(releaseInvokerTimer.current);
    },
    [],
  );

  const promptError =
    request?.kind === 'prompt' ? promptValidationMessage(request) : null;
  const visiblePromptError =
    request?.kind === 'prompt' && request.touched ? promptError : null;
  const submitPrompt = () => {
    if (!request || request.kind !== 'prompt') return;
    if (promptValidationMessage(request)) {
      setRequest({ ...request, touched: true });
      return;
    }
    closePrompt(request.value);
  };

  const dialog = request ? (
    <Dialog
      key={request.id}
      title={request.title}
      description={request.description}
      className="work-office-dialog"
      focusKey={request.id}
      restoreFocusTarget={() =>
        request.restoreFocusTarget?.() ?? invokerRef.current
      }
      onClose={() =>
        request.kind === 'prompt'
          ? closePrompt(null)
          : request.kind === 'confirm'
            ? closeConfirm(false)
            : closeNotice()
      }
      footer={
        request.kind === 'prompt' ? (
          <>
            <Button tone="quiet" onClick={() => closePrompt(null)}>
              取消
            </Button>
            <Button
              tone="primary"
              disabled={Boolean(promptError)}
              onClick={submitPrompt}
            >
              {request.confirmLabel}
            </Button>
          </>
        ) : request.kind === 'confirm' ? (
          <>
            <Button tone="quiet" onClick={() => closeConfirm(false)}>
              {request.cancelLabel}
            </Button>
            <Button
              tone={request.confirmTone}
              onClick={() => closeConfirm(true)}
            >
              {request.confirmLabel}
            </Button>
          </>
        ) : (
          <Button tone="primary" onClick={closeNotice}>
            {request.confirmLabel}
          </Button>
        )
      }
    >
      {request.kind === 'prompt' && (
        <label className="work-office-dialog-field" htmlFor={promptFieldId}>
          <span className="work-office-dialog-field-label">
            {request.fieldLabel}
          </span>
          {request.multiline ? (
            <OfficeTextArea
              id={promptFieldId}
              aria-label={request.fieldLabel}
              aria-invalid={Boolean(visiblePromptError) || undefined}
              value={request.value}
              placeholder={request.placeholder}
              onChange={(event) =>
                setRequest({
                  ...request,
                  value: event.target.value,
                  touched: true,
                })
              }
              onBlur={() => setRequest({ ...request, touched: true })}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
                  event.preventDefault();
                  submitPrompt();
                }
              }}
            />
          ) : (
            <OfficeTextField
              id={promptFieldId}
              aria-label={request.fieldLabel}
              aria-invalid={Boolean(visiblePromptError) || undefined}
              value={request.value}
              inputMode={request.inputMode}
              placeholder={request.placeholder}
              onChange={(event) =>
                setRequest({
                  ...request,
                  value: event.target.value,
                  touched: true,
                })
              }
              onBlur={() => setRequest({ ...request, touched: true })}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  submitPrompt();
                }
              }}
            />
          )}
          {visiblePromptError && (
            <span className="work-office-dialog-field-error" role="alert">
              {visiblePromptError}
            </span>
          )}
        </label>
      )}
    </Dialog>
  ) : null;

  return { prompt, notice, confirm, dialog };
}

function promptValidationMessage(request: OfficePromptRequest): string | null {
  if (!request.value.trim() && request.requiredMessage)
    return request.requiredMessage;
  return request.validate?.(request.value) ?? null;
}
