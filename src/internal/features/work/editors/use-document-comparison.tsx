import type { Content, Editor } from '@tiptap/core';
import { useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import {
  applyDocumentComparison,
  type DocumentComparisonApplyResult,
  type DocumentComparisonMode,
} from '../work-document-compare';
import { importWorkDocumentFile } from '../work-document-file-io';
import {
  documentModelForContent,
  resolveWorkDocumentEditorInput,
} from '../work-document-model';
import { normalizeDocumentHtml } from '../work-document-section';
import { forgetWorkSourceBlob } from '../work-repository';
import {
  DocumentCompareDialog,
  type DocumentCompareDialogRequest,
} from './document-compare-dialog';

interface DocumentComparisonDialogState {
  mode: DocumentComparisonMode;
  restoreFocusTarget: HTMLElement | null;
}

export function useDocumentComparison({
  editor,
  onApplied,
}: {
  editor: Editor | null;
  onApplied: () => void;
}) {
  const [dialog, setDialog] = useState<DocumentComparisonDialogState | null>(
    null,
  );
  const editorRef = useRef(editor);
  editorRef.current = editor;

  const open = (mode: DocumentComparisonMode) => {
    const current = editorRef.current;
    if (!current || current.isDestroyed) return;
    const active = document.activeElement;
    setDialog({
      mode,
      restoreFocusTarget:
        active instanceof HTMLElement && active.isConnected
          ? active
          : current.view.dom,
    });
  };
  const close = () => setDialog(null);
  const submit = async (
    request: DocumentCompareDialogRequest,
  ): Promise<DocumentComparisonApplyResult> => {
    const current = editorRef.current;
    if (!current || current.isDestroyed) return invalidComparisonResult();
    const extension = comparisonFileExtension(request.file.name);
    if (!extension) return invalidComparisonResult();
    const artifact = await importWorkDocumentFile(request.file, extension);
    try {
      if (artifact.content.type !== 'document')
        return invalidComparisonResult();
      const content = artifact.content;
      const model = documentModelForContent(content);
      const fallbackHtml = model
        ? content.html
        : normalizeDocumentHtml(content);
      const input = resolveWorkDocumentEditorInput(
        content,
        fallbackHtml,
        model,
      );
      return applyDocumentComparison(current, input.source as Content, {
        author: request.author,
        date: new Date().toISOString(),
        mode: request.mode,
        sourceName: request.file.name,
      });
    } finally {
      forgetWorkSourceBlob(artifact.id);
    }
  };
  const applied = (result: DocumentComparisonApplyResult) => {
    const count =
      result.summary.insertions +
      result.summary.deletions +
      result.summary.formatting +
      result.summary.paragraphFormatting;
    close();
    onApplied();
    showToast(`已生成 ${count} 项可审阅修订`, 'success');
  };

  return {
    dialog: dialog ? (
      <DocumentCompareDialog
        initialMode={dialog.mode}
        restoreFocusTarget={() =>
          dialog.restoreFocusTarget?.isConnected
            ? dialog.restoreFocusTarget
            : (editorRef.current?.view.dom ?? null)
        }
        onApplied={applied}
        onClose={close}
        onSubmit={submit}
      />
    ) : null,
    open,
  };
}

function comparisonFileExtension(name: string): string | null {
  const extension = name.split('.').at(-1)?.toLocaleLowerCase() ?? '';
  return extension === 'docx' ||
    extension === 'html' ||
    extension === 'htm' ||
    extension === 'txt'
    ? extension
    : null;
}

function invalidComparisonResult(): DocumentComparisonApplyResult {
  return {
    status: 'unsupported',
    summary: {
      deletions: 0,
      formatting: 0,
      insertions: 0,
      paragraphFormatting: 0,
    },
    diagnostics: [
      {
        code: 'invalid-revised-content',
        message: 'The selected file cannot be imported as a Writer document.',
      },
    ],
  };
}
