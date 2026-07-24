import {
  generateHTML,
  getSchema,
  type Extensions,
  type JSONContent,
} from '@tiptap/core';
import { DOMParser as ProseMirrorDOMParser } from '@tiptap/pm/model';
import { createWorkDocumentExtensions } from './work-document-extensions';
import {
  createWorkDocumentModel,
  documentHtmlFingerprint,
  documentModelForContent,
} from './work-document-model';
import { syncDocumentContentFromHtml } from './work-document-section';
import type { WorkDocumentContent, WorkDocumentNode } from './work-types';

let schema: ReturnType<typeof getSchema> | null = null;
let extensions: Extensions | null = null;

export function createWorkDocumentModelFromContent(
  content: WorkDocumentContent,
): WorkDocumentContent {
  const previous = content.model;
  const synchronized = syncDocumentContentFromHtml(content, content.html);
  const document = new DOMParser().parseFromString(
    synchronized.html,
    'text/html',
  );
  schema ??= getSchema(documentExtensions());
  const root = ProseMirrorDOMParser.fromSchema(schema)
    .parse(document.body)
    .toJSON() as unknown as WorkDocumentNode;
  return {
    ...synchronized,
    model: createWorkDocumentModel(synchronized.html, root, previous),
  };
}

export function materializeWorkDocumentContent(
  content: WorkDocumentContent,
): WorkDocumentContent {
  const model = documentModelForContent(content);
  if (!model) return content;
  let html: string;
  try {
    html = generateHTML(
      model.root as unknown as JSONContent,
      documentExtensions(),
    );
  } catch (error) {
    throw new Error(
      'The structured document model cannot be serialized by this Office schema.',
      { cause: error },
    );
  }
  const synchronized = syncDocumentContentFromHtml(content, html);
  return {
    ...synchronized,
    model: {
      ...model,
      htmlFingerprint: documentHtmlFingerprint(synchronized.html),
    },
  };
}

function documentExtensions(): Extensions {
  extensions ??= createWorkDocumentExtensions();
  return extensions;
}
