import type { Extensions } from '@tiptap/core';
import Color from '@tiptap/extension-color';
import Highlight from '@tiptap/extension-highlight';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import FontFamily from '@tiptap/extension-text-style/font-family';
import FontSize from '@tiptap/extension-text-style/font-size';
import Underline from '@tiptap/extension-underline';
import StarterKit from '@tiptap/starter-kit';
import {
  DocumentCaption,
  DocumentCrossReference,
} from './work-document-caption-nodes';
import {
  DocumentSubscript,
  DocumentSuperscript,
} from './work-document-character-formatting';
import {
  DocumentChange,
  type WorkDocumentChangeIdentity,
  type WorkDocumentChangeKind,
} from './work-document-changes';
import {
  DocumentBibliography,
  DocumentCitation,
} from './work-document-citation-nodes';
import { DocumentComment } from './work-document-comments';
import { DocumentField } from './work-document-field-node';
import { DocumentImage } from './work-document-image-layout';
import {
  DocumentNote,
  DocumentNoteReference,
} from './work-document-note-nodes';
import { DocumentPageBreak } from './work-document-page-break';
import { DocumentParagraphFormatting } from './work-document-paragraph-formatting';
import { DocumentSection } from './work-document-section-node';
import { DocumentTab } from './work-document-tab-node';
import { DocumentParagraphTabStops } from './work-document-tab-stops';
import { DocumentTableCommands } from './work-document-table-commands';
import { DocumentTableRow } from './work-document-table-row';
import type { WorkDocumentContent } from './work-types';

export interface WorkDocumentExtensionOptions {
  getContent?: () => WorkDocumentContent | null;
  isTracking?: () => boolean;
  createChange?: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity;
  onContentChange?: (content: WorkDocumentContent) => void;
  onTrackingChange?: (enabled: boolean) => void;
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
      trailingNode: {
        notAfter: ['documentSection'],
      },
    }),
    DocumentSection,
    DocumentCaption,
    DocumentCrossReference,
    citationExtension,
    DocumentBibliography,
    DocumentField,
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
      table: { resizable: true, allowTableNodeSelection: true },
      tableRow: false,
    }),
    DocumentTableRow,
    DocumentTableCommands,
    TextStyle,
    FontFamily,
    FontSize,
    Color,
    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),
    DocumentParagraphFormatting,
    DocumentParagraphTabStops,
    DocumentTab,
    DocumentPageBreak,
    changeExtension,
  ];
}
