import type { Editor } from '@tiptap/core';
import { TriangleAlert } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { Button, InlineNotice } from '../../../design-system/primitives';
import type { WorkDocumentEditorInput } from '../work-document-model';
import {
  collectDocumentReviewSnapshot,
  documentReviewConflictKey,
  reconcileDocumentReviewConflicts,
  type WorkDocumentReviewConflict,
  type WorkDocumentReviewConflictEvent,
} from '../work-document-review-conflicts';
import { normalizeDocumentHtml } from '../work-document-section';
import type { WorkDocumentContent } from '../work-types';

interface MutableValue<T> {
  current: T;
}

export interface UseDocumentReviewConflictsOptions {
  activeConflictsRef: MutableValue<WorkDocumentReviewConflict[]>;
  appliedSourceKeyRef: MutableValue<string>;
  artifactId?: string;
  content: WorkDocumentContent;
  editor: Editor | null;
  editorInput: WorkDocumentEditorInput;
  normalizedContent: string;
  onReviewConflict?: (event: WorkDocumentReviewConflictEvent) => void;
  reconcileControlledUpdates?: boolean;
}

export function useDocumentReviewConflicts({
  activeConflictsRef,
  appliedSourceKeyRef,
  artifactId,
  content,
  editor,
  editorInput,
  normalizedContent,
  onReviewConflict,
  reconcileControlledUpdates = true,
}: UseDocumentReviewConflictsOptions) {
  const appliedArtifactIdRef = useRef(artifactId);
  const onReviewConflictRef = useRef(onReviewConflict);
  const [visibleConflicts, setVisibleConflicts] = useState<
    WorkDocumentReviewConflict[]
  >([]);
  onReviewConflictRef.current = onReviewConflict;

  useEffect(() => {
    if (!editor || !reconcileControlledUpdates) return;
    const artifactChanged = appliedArtifactIdRef.current !== artifactId;
    const sourceChanged = appliedSourceKeyRef.current !== editorInput.sourceKey;
    appliedArtifactIdRef.current = artifactId;
    if (artifactChanged) {
      activeConflictsRef.current = [];
      setVisibleConflicts([]);
    }
    const before = collectDocumentReviewSnapshot(editor.state.doc);
    if (sourceChanged) {
      const currentContent = normalizeDocumentHtml({
        ...content,
        html: editor.getHTML(),
      });
      const canReuseCurrentDocument =
        typeof editorInput.source === 'string' &&
        currentContent === normalizedContent;
      appliedSourceKeyRef.current = editorInput.sourceKey;
      if (!canReuseCurrentDocument) {
        editor
          .chain()
          .setMeta('addToHistory', false)
          .setContent(editorInput.source, { emitUpdate: false })
          .run();
      }
    }
    if (
      artifactChanged ||
      (!sourceChanged && activeConflictsRef.current.length === 0)
    ) {
      return;
    }
    const reconciled = reconcileDocumentReviewConflicts({
      active: activeConflictsRef.current,
      before,
      after: collectDocumentReviewSnapshot(editor.state.doc),
      comments: content.comments ?? [],
    });
    activeConflictsRef.current = reconciled.active;
    setVisibleConflicts((visible) => {
      if (reconciled.detected.length) return reconciled.active;
      const visibleKeys = new Set(visible.map(documentReviewConflictKey));
      return reconciled.active.filter((conflict) =>
        visibleKeys.has(documentReviewConflictKey(conflict)),
      );
    });
    if (reconciled.detected.length) {
      onReviewConflictRef.current?.({
        artifactId,
        conflicts: reconciled.detected.map((conflict) => ({ ...conflict })),
      });
    }
  }, [
    activeConflictsRef,
    appliedSourceKeyRef,
    artifactId,
    content,
    editor,
    editorInput,
    normalizedContent,
    reconcileControlledUpdates,
  ]);

  return {
    dismiss: () => setVisibleConflicts([]),
    visibleConflicts,
  };
}

export function DocumentReviewConflictNotice({
  conflicts,
  onDismiss,
}: {
  conflicts: readonly WorkDocumentReviewConflict[];
  onDismiss: () => void;
}) {
  if (!conflicts.length) return null;
  return (
    <InlineNotice
      className="work-document-review-conflict"
      icon={<TriangleAlert />}
      role="alert"
      title="审阅内容与外部更新冲突"
      tone="warning"
      actions={
        <Button
          aria-label="关闭审阅冲突提示"
          size="compact"
          tone="quiet"
          onClick={onDismiss}
        >
          关闭
        </Button>
      }
    >
      外部版本已应用，其中 {conflicts.length}{' '}
      个批注或修订范围发生变化。请检查后再继续审阅。
    </InlineNotice>
  );
}
