import type { Editor } from '@tiptap/core';
import { useCallback, useEffect, useState } from 'react';
import type { WorkOfficeCollaborationActor } from '../../../collaboration/office-collaboration';
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
  canDeleteComment: (id: string) => boolean;
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
  actor,
  contentRef,
  deleteOwnOnly = false,
  editor,
  enabled = true,
  onBeforeDraft,
}: {
  actor?: WorkOfficeCollaborationActor;
  contentRef: { current: WorkDocumentContent };
  deleteOwnOnly?: boolean;
  editor: Editor | null;
  enabled?: boolean;
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
      if (!enabled) return;
      editor?.commands.addDocumentCommentReply(id, {
        id: createWorkId('comment-reply'),
        ...commentActor(actor),
        date: new Date().toISOString(),
        text,
      });
    },
    [actor, editor, enabled],
  );

  const toggleResolved = useCallback(
    (id: string) => {
      if (!enabled) return;
      editor?.commands.toggleDocumentCommentResolved(id);
    },
    [editor, enabled],
  );

  const canDeleteComment = useCallback(
    (id: string) => {
      if (!enabled) return false;
      if (!deleteOwnOnly) return true;
      return Boolean(
        actor &&
          contentRef.current.comments?.some(
            (comment) => comment.id === id && comment.actorId === actor.id,
          ),
      );
    },
    [actor, contentRef, deleteOwnOnly, enabled],
  );

  const deleteComment = useCallback(
    (id: string) => {
      if (!canDeleteComment(id)) return;
      editor?.commands.deleteDocumentComment(id);
    },
    [canDeleteComment, editor],
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
    if (!editor || !enabled) return;
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
  }, [editor, enabled, onBeforeDraft]);

  const submitDraft = useCallback(
    (text: string): string | null => {
      if (!enabled || !editor || !draft)
        return '批注草稿已经关闭，请重新选择文字。';
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
        ...commentActor(actor),
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
    [actor, draft, editor, enabled],
  );

  const comments = editor
    ? documentCommentViews(
        contentRef.current.comments ?? [],
        collectDocumentCommentAnchors(editor.state.doc),
      )
    : [];
  const canInsert = Boolean(
    enabled && editor && canInsertDocumentComment(editor),
  );
  const toggleOpen = useCallback(() => {
    if (open) close();
    else setOpen(true);
  }, [close, open]);

  return {
    canDeleteComment,
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

function commentActor(actor: WorkOfficeCollaborationActor | undefined): {
  actorId?: string;
  author: string;
} {
  return actor
    ? { actorId: actor.id, author: actor.name }
    : { author: CURRENT_COMMENT_AUTHOR };
}
