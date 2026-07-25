import type { Editor } from '@tiptap/core';
import { useCallback, useEffect, useState, type RefObject } from 'react';
import {
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
  updateDocumentPageChromeVariant,
} from '../work-document-page-chrome';
import { documentSectionById } from '../work-document-section-editor';
import type {
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
  WorkDocumentSectionLayout,
} from '../work-types';
import type { DocumentPageChromeEditingPart } from './document-page-chrome-ribbon';
import type { DocumentPaginationPageDescriptor } from './use-document-pagination';

export interface DocumentPageChromeEditingState {
  part: DocumentPageChromeEditingPart;
  sectionId: string;
  variant: WorkDocumentPageChromeVariant;
}

export function useDocumentPageChrome({
  editor,
  enabled,
  firstPage,
  footerRef,
  headerRef,
  lastPage,
  layout,
  onBeforeEdit,
  visiblePage,
}: {
  editor: Editor | null;
  enabled: boolean;
  firstPage: DocumentPaginationPageDescriptor;
  footerRef: RefObject<HTMLElement | null>;
  headerRef: RefObject<HTMLElement | null>;
  lastPage: DocumentPaginationPageDescriptor;
  layout: WorkDocumentSectionLayout;
  onBeforeEdit: () => boolean | undefined | Promise<boolean | undefined>;
  visiblePage: DocumentPaginationPageDescriptor;
}) {
  const [editing, setEditing] = useState<DocumentPageChromeEditingState | null>(
    null,
  );
  const [chromeEditor, setChromeEditor] = useState<Editor | null>(null);

  const reset = useCallback(() => {
    setEditing(null);
    setChromeEditor(null);
  }, []);

  useEffect(() => {
    if (!enabled || !editor) reset();
  }, [editor, enabled, reset]);

  const editingSection =
    editing && editor ? documentSectionById(editor, editing.sectionId) : null;
  const editingChrome = editing
    ? normalizeDocumentPageChrome(
        editingSection?.layout.pageChrome,
        editingSection?.layout,
      )[editing.variant]
    : null;
  const visibleChrome = editingChrome ?? visiblePage.pageChrome;
  const headerChrome =
    editing?.part === 'header' ? visibleChrome : firstPage.pageChrome;
  const footerChrome =
    editing?.part === 'footer' ? visibleChrome : lastPage.pageChrome;

  const update = useCallback(
    (patch: Partial<WorkDocumentPageChromeContent>) => {
      if (!editor || editor.isDestroyed) return;
      const target =
        editing ??
        ({
          part: 'header',
          sectionId: visiblePage.sectionId,
          variant: visiblePage.pageChrome.variant,
        } satisfies DocumentPageChromeEditingState);
      const targetSection = documentSectionById(editor, target.sectionId);
      const targetLayout = targetSection?.layout ?? layout;
      const targetPageChrome = normalizeDocumentPageChrome(
        targetLayout.pageChrome,
        targetLayout,
      );
      const nextPageChrome = updateDocumentPageChromeVariant(
        targetPageChrome,
        target.variant,
        patch,
      );
      editor.commands.updateDocumentSection(target.sectionId, {
        ...targetLayout,
        pageChrome: nextPageChrome,
        ...documentPageChromeLegacyFields(nextPageChrome),
      });
    },
    [editing, editor, layout, visiblePage],
  );

  const edit = useCallback(
    async (part: DocumentPageChromeEditingPart) => {
      const target = part === 'header' ? firstPage : lastPage;
      if ((await onBeforeEdit()) === false) return;
      setChromeEditor(null);
      setEditing({
        part,
        sectionId: target.sectionId,
        variant: target.pageChrome.variant,
      });
      requestAnimationFrame(() => {
        const element =
          part === 'header' ? headerRef.current : footerRef.current;
        element?.scrollIntoView?.({ block: 'center' });
      });
    },
    [firstPage, footerRef, headerRef, lastPage, onBeforeEdit],
  );

  const close = useCallback(() => {
    reset();
    requestAnimationFrame(() => {
      if (editor && !editor.isDestroyed) editor.commands.focus();
    });
  }, [editor, reset]);

  const togglePageNumber = useCallback(() => {
    update({ showPageNumber: !visibleChrome.showPageNumber });
  }, [update, visibleChrome.showPageNumber]);

  return {
    chromeEditor,
    close,
    edit,
    editing,
    footerChrome,
    headerChrome,
    reset,
    setChromeEditor,
    togglePageNumber,
    update,
    visibleChrome,
  };
}
