import { useCallback, type ReactNode } from 'react';
import { createWorkId } from '../work-templates';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { useOfficeDialog } from './office-controls';
import { clamp, updateSlide } from './presentation-editor-operations';

export interface PresentationReviewCommands {
  addComment: () => void;
  closeComments: () => void;
  deleteComment: (slideId: string, commentId: string) => void;
  dialog: ReactNode;
  locateComment: (slideId: string, commentId: string) => void;
  openComment: (commentId: string) => void;
  updateComment: (slideId: string, commentId: string, text: string) => void;
}

export function usePresentationReviewCommands({
  content,
  onChange,
  onClearSelection,
  onCloseComments,
  onOpenComments,
  onSelectComment,
  onSelectSlide,
  selectedElement,
  selectedSlide,
}: {
  content: WorkPresentationContent;
  onChange: (content: WorkPresentationContent) => void;
  onClearSelection: () => void;
  onCloseComments: () => void;
  onOpenComments: () => void;
  onSelectComment: (id: string | null) => void;
  onSelectSlide: (id: string) => void;
  selectedElement: WorkSlideElement | null;
  selectedSlide: WorkSlide;
}): PresentationReviewCommands {
  const officeDialog = useOfficeDialog();
  const addComment = useCallback(() => {
    void officeDialog
      .prompt({ title: '批注内容', multiline: true, confirmLabel: '添加批注' })
      .then((text) => {
        if (!text?.trim()) return;
        const comment = {
          id: createWorkId('slide-comment'),
          author: 'A3S Work 用户',
          initials: 'AW',
          date: new Date().toISOString(),
          text: text.trim(),
          x: clamp(
            selectedElement ? selectedElement.x + selectedElement.width : 50,
            2,
            98,
          ),
          y: clamp(selectedElement ? selectedElement.y : 50, 2, 98),
        };
        updateSlide(
          content,
          selectedSlide.id,
          (slide) => ({
            ...slide,
            comments: [...(slide.comments ?? []), comment],
          }),
          onChange,
        );
        onSelectComment(comment.id);
        onOpenComments();
      });
  }, [
    content,
    officeDialog,
    onChange,
    onOpenComments,
    onSelectComment,
    selectedElement,
    selectedSlide.id,
  ]);

  const openComment = useCallback(
    (commentId: string) => {
      onSelectComment(commentId);
      onOpenComments();
    },
    [onOpenComments, onSelectComment],
  );

  const locateComment = useCallback(
    (slideId: string, commentId: string) => {
      onSelectSlide(slideId);
      onClearSelection();
      onSelectComment(commentId);
    },
    [onClearSelection, onSelectComment, onSelectSlide],
  );

  const updateComment = useCallback(
    (slideId: string, commentId: string, text: string) => {
      updateSlide(
        content,
        slideId,
        (slide) => ({
          ...slide,
          comments: slide.comments?.map((comment) =>
            comment.id === commentId ? { ...comment, text } : comment,
          ),
        }),
        onChange,
      );
    },
    [content, onChange],
  );

  const deleteComment = useCallback(
    (slideId: string, commentId: string) => {
      updateSlide(
        content,
        slideId,
        (slide) => ({
          ...slide,
          comments: slide.comments?.filter(
            (comment) => comment.id !== commentId,
          ),
        }),
        onChange,
      );
      onSelectComment(null);
    },
    [content, onChange, onSelectComment],
  );

  return {
    addComment,
    closeComments: onCloseComments,
    deleteComment,
    dialog: officeDialog.dialog,
    locateComment,
    openComment,
    updateComment,
  };
}
