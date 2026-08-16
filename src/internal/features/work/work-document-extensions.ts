import type { Extensions } from '@tiptap/core';
import Color from '@tiptap/extension-color';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import FontSize from '@tiptap/extension-text-style/font-size';
import Underline from '@tiptap/extension-underline';
import StarterKit from '@tiptap/starter-kit';
import { DocumentBookmarkBoundary } from './work-document-bookmarks';
import {
  DocumentCaption,
  DocumentCrossReference,
} from './work-document-caption-nodes';
import {
  DocumentChange,
  type WorkDocumentChangeIdentity,
  type WorkDocumentChangeKind,
} from './work-document-changes';
import {
  DocumentSubscript,
  DocumentSuperscript,
} from './work-document-character-formatting';
import {
  DocumentBibliography,
  DocumentCitation,
} from './work-document-citation-nodes';
import { DocumentComment } from './work-document-comments';
import { DocumentEquation } from './work-document-equations';
import { DocumentField } from './work-document-field-node';
import { DocumentImage } from './work-document-image-layout';
import {
  DocumentBulletList,
  DocumentListCommands,
  DocumentOrderedList,
} from './work-document-lists';
import {
  DocumentNote,
  DocumentNoteReference,
} from './work-document-note-nodes';
import { DocumentPageBreak } from './work-document-page-break';
import { DocumentParagraphFormatting } from './work-document-paragraph-formatting';
import { DocumentParagraphIdentity } from './work-document-paragraph-identity';
import { DocumentSection } from './work-document-section-node';
import { DocumentTab } from './work-document-tab-node';
import { DocumentParagraphTabStops } from './work-document-tab-stops';
import {
  DocumentTableCell,
  DocumentTableFormatting,
  DocumentTableHeader,
} from './work-document-table-cell-formatting';
import { DocumentTableCommands } from './work-document-table-commands';
import { DocumentTableRow } from './work-document-table-row';
import { DocumentTableRowIdentity } from './work-document-table-row-identity';
import {
  DocumentTable,
  DocumentTableSizing,
} from './work-document-table-sizing';
import {
  DocumentFontFamily,
  DocumentHighlight,
  DocumentTextStyle,
} from './work-document-word-line-metrics';
import type { WorkDocumentContent } from './work-types';

export interface WorkDocumentExtensionOptions {
  collaborative?: boolean;
  getContent?: () => WorkDocumentContent | null;
  isTracking?: () => boolean;
  createChange?: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity;
  onContentChange?: (content: WorkDocumentContent) => void;
  onTrackingChange?: (enabled: boolean) => void;
  rotateTrackedTextIdentities?: () => boolean;
}

export function createWorkDocumentExtensions(
  options: WorkDocumentExtensionOptions = {},
): Extensions {
  const changeExtension = DocumentChange.configure({
    ...(options.isTracking ? { isTracking: options.isTracking } : {}),
    ...(options.createChange ? { createChange: options.createChange } : {}),
    ...(options.onTrackingChange
      ? { onTrackingChange: options.onTrackingChange }
      : {}),
  });
  const citationExtension = DocumentCitation.configure({
    ...(options.getContent ? { getContent: options.getContent } : {}),
    ...(options.onContentChange
      ? { onContentChange: options.onContentChange }
      : {}),
  });
  const commentExtension = DocumentComment.configure({
    ...(options.getContent ? { getContent: options.getContent } : {}),
    ...(options.onContentChange
      ? { onContentChange: options.onContentChange }
      : {}),
  });
  return [
    StarterKit.configure({
      link: {
        autolink: true,
        defaultProtocol: 'https',
        openOnClick: false,
      },
      underline: false,
      bulletList: false,
      orderedList: false,
      trailingNode: {
        notAfter: ['documentSection'],
      },
      ...(options.collaborative ? { undoRedo: false } : {}),
    }),
    DocumentBulletList,
    DocumentOrderedList,
    DocumentListCommands,
    DocumentSection,
    DocumentBookmarkBoundary,
    DocumentCaption,
    DocumentCrossReference,
    citationExtension,
    DocumentBibliography,
    DocumentField,
    DocumentEquation,
    commentExtension,
    DocumentNoteReference,
    DocumentNote,
    Underline,
    DocumentSubscript,
    DocumentSuperscript,
    DocumentImage.configure({
      allowBase64: true,
      resize: {
        enabled: true,
        alwaysPreserveAspectRatio: true,
        minWidth: 60,
        minHeight: 40,
      },
    }),
    TableKit.configure({
      table: false,
      tableCell: false,
      tableHeader: false,
      tableRow: false,
    }),
    DocumentTable,
    DocumentTableCell,
    DocumentTableHeader,
    DocumentTableRow,
    DocumentTableRowIdentity.configure({
      rotateTextId: options.rotateTrackedTextIdentities ?? (() => true),
    }),
    DocumentTableCommands,
    DocumentTableFormatting,
    DocumentTableSizing,
    DocumentTextStyle,
    DocumentFontFamily,
    FontSize,
    Color,
    DocumentHighlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),
    DocumentParagraphIdentity.configure({
      rotateTextId: options.rotateTrackedTextIdentities ?? (() => true),
    }),
    DocumentParagraphFormatting,
    DocumentParagraphTabStops,
    DocumentTab,
    DocumentPageBreak,
    changeExtension,
  ];
}
