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
import { DocumentTableRow } from './work-document-table-row';

export interface WorkDocumentExtensionOptions {
  isTracking?: () => boolean;
  createChange?: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity;
}

export function createWorkDocumentExtensions(
  options: WorkDocumentExtensionOptions = {},
): Extensions {
  const changeExtension =
    options.isTracking && options.createChange
      ? DocumentChange.configure({
          isTracking: options.isTracking,
          createChange: options.createChange,
        })
      : DocumentChange;
  return [
    StarterKit.configure({
      link: {
        autolink: true,
        defaultProtocol: 'https',
        openOnClick: false,
      },
      underline: false,
    }),
    DocumentSection,
    DocumentCaption,
    DocumentCrossReference,
    DocumentCitation,
    DocumentBibliography,
    DocumentField,
    DocumentComment,
    DocumentNoteReference,
    DocumentNote,
    Underline,
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
