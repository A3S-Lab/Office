import type { Editor } from '@tiptap/core';
import { useCallback, useEffect, useState } from 'react';
import {
  canInsertDocumentComment,
  collectDocumentCommentAnchors,
  documentCommentDraftRange,
  documentCommentViews,
  type WorkDocumentCommentView,
} from '../work-document-comments';
import { createWorkId } from '../work-templates';
import type { WorkDocumentContent } from '../work-types';
import type { DocumentCommentDraft } from './document-comment-composer';

const CURRENT_COMMENT_AUTHOR = '我';

export interface DocumentCommentsController {
  canInsert: boolean;
  close: () => void;
  closeDraft: (restoreEditorFocus?: boolean) => void;
  comments: WorkDocumentCommentView[];
  deleteComment: (id: string) => void;
  draft: DocumentCommentDraft | null;
  open: boolean;
  reply: (id: string, text: string) => void;
  setOpen: (open: boolean) => void;
  startDraft: () => void;
  submitDraft: (text: string) => string | null;
  toggleOpen: () => void;
  toggleResolved: (id: string) => void;
}

export function useDocumentComments({
  contentRef,
  editor,
  onBeforeDraft,
}: {
  contentRef: { current: WorkDocumentContent };
  editor: Editor | null;
  onBeforeDraft?: () => void;
}): DocumentCommentsController {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState<DocumentCommentDraft | null>(null);

  useEffect(
    () => () => {
      if (editor && !editor.isDestroyed) {
        editor.commands.clearDocumentCommentDraft();
      }
    },
    [editor],
  );

  const reply = useCallback(
    (id: string, text: string) => {
      editor?.commands.addDocumentCommentReply(id, {
        id: createWorkId('comment-reply'),
        author: CURRENT_COMMENT_AUTHOR,
        date: new Date().toISOString(),
        text,
      });
    },
    [editor],
  );

  const toggleResolved = useCallback(
    (id: string) => {
      editor?.commands.toggleDocumentCommentResolved(id);
    },
    [editor],
  );

  const deleteComment = useCallback(
    (id: string) => {
      editor?.commands.deleteDocumentComment(id);
    },
    [editor],
  );

  const closeDraft = useCallback(
    (restoreEditorFocus = true) => {
      editor?.commands.clearDocumentCommentDraft();
      setDraft(null);
      if (restoreEditorFocus && editor) {
        requestAnimationFrame(() => {
          if (!editor.isDestroyed) editor.commands.focus();
        });
      }
    },
    [editor],
  );

  const close = useCallback(() => {
    closeDraft();
    setOpen(false);
  }, [closeDraft]);

  const startDraft = useCallback(() => {
    if (!editor) return;
    const { from, to } = editor.state.selection;
    const range = { from, to };
    if (!canInsertDocumentComment(editor, range)) return;
    const nextDraft: DocumentCommentDraft = {
      id: createWorkId('comment'),
      from,
      to,
      anchorText: editor.state.doc.textBetween(from, to, '\n'),
    };
    if (!editor.commands.showDocumentCommentDraft(range)) return;
    onBeforeDraft?.();
    setDraft(nextDraft);
    setOpen(true);
  }, [editor, onBeforeDraft]);

  const submitDraft = useCallback(
    (text: string): string | null => {
      if (!editor || !draft) return '批注草稿已经关闭，请重新选择文字。';
      const range = documentCommentDraftRange(editor);
      if (!range) return '所选文字已变化，请重新选择。';
      const anchorText = editor.state.doc.textBetween(
        range.from,
        range.to,
        '\n',
      );
      if (anchorText !== draft.anchorText)
        return '所选文字已变化，请重新选择。';
      const comment = {
        id: draft.id,
        author: CURRENT_COMMENT_AUTHOR,
        date: new Date().toISOString(),
        text,
        resolved: false,
      };
      if (!editor.commands.insertDocumentCommentThread(comment, range)) {
        return '所选文字已经包含批注，请重新选择。';
      }
      setDraft(null);
      setOpen(true);
      return null;
    },
    [draft, editor],
  );

  const comments = editor
    ? documentCommentViews(
        contentRef.current.comments ?? [],
        collectDocumentCommentAnchors(editor.state.doc),
      )
    : [];
  const canInsert = Boolean(editor && canInsertDocumentComment(editor));
  const toggleOpen = useCallback(() => {
    if (open) close();
    else setOpen(true);
  }, [close, open]);

  return {
    canInsert,
    close,
    closeDraft,
    comments,
    deleteComment,
    draft,
    open,
    reply,
    setOpen,
    startDraft,
    submitDraft,
    toggleOpen,
    toggleResolved,
  };
}
