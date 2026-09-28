import { Component, type ErrorInfo, type ReactNode } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';

export type OfficeEditorKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

export type OfficeEditorHostErrorPhase = 'render' | 'chunk-load';

export type OfficeEditorHostError = {
  editor: OfficeEditorKind;
  error: unknown;
  phase: OfficeEditorHostErrorPhase;
};

export type OfficeEditorDiagnostic = {
  editor: OfficeEditorKind;
  source: OfficeEditorHostErrorPhase | 'compatibility' | 'performance';
  message: string;
  error?: unknown;
};

type OfficeEditorErrorBoundaryProps = {
  children: ReactNode;
  editor: OfficeEditorKind;
  messages: OfficeMessageCatalog;
  onError?: (error: OfficeEditorHostError) => void;
  onDiagnostic?: (diagnostic: OfficeEditorDiagnostic) => void;
  title: string;
};

type OfficeEditorErrorBoundaryState = {
  error: Error | null;
};

function isChunkLoadError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  const message = error.message.toLowerCase();
  return (
    error.name === 'ChunkLoadError' ||
    message.includes('loading chunk') ||
    message.includes('failed to fetch dynamically imported module') ||
    message.includes('importing a module script failed')
  );
}

/**
 * Recoverable host boundary around lazy editor chunks. A render or chunk-load
 * failure must not unmount the host page; the boundary keeps a retry surface
 * and notifies `onError` / `onDiagnostic` once per failure.
 */
export class OfficeEditorErrorBoundary extends Component<
  OfficeEditorErrorBoundaryProps,
  OfficeEditorErrorBoundaryState
> {
  state: OfficeEditorErrorBoundaryState = { error: null };
  private notifiedError: Error | null = null;

  static getDerivedStateFromError(error: Error): OfficeEditorErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, _info: ErrorInfo): void {
    if (this.notifiedError === error) return;
    this.notifiedError = error;
    const phase: OfficeEditorHostErrorPhase = isChunkLoadError(error)
      ? 'chunk-load'
      : 'render';
    const hostError: OfficeEditorHostError = {
      editor: this.props.editor,
      error,
      phase,
    };
    this.props.onError?.(hostError);
    this.props.onDiagnostic?.({
      editor: this.props.editor,
      source: phase,
      message: error.message,
      error,
    });
  }

  private retry = () => {
    this.notifiedError = null;
    this.setState({ error: null });
  };

  render(): ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;

    const phase = isChunkLoadError(error) ? 'chunk-load' : 'render';
    return (
      <div
        className="a3s-office-editor-error"
        data-office-editor-error={phase}
        data-office-editor-kind={this.props.editor}
        role="alert"
      >
        <p>
          {officeMessage(this.props.messages, 'editor.error.failed', {
            title: this.props.title,
          })}
        </p>
        <p>{error.message}</p>
        <button type="button" onClick={this.retry}>
          {officeMessage(this.props.messages, 'editor.error.retry')}
        </button>
      </div>
    );
  }
}
