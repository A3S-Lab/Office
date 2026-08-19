import {
  type Extensions,
  generateHTML,
  getSchema,
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
import {
  materializeWindowedDocumentModel,
  windowDocumentModel,
} from './work-document-windowing';
import type { WorkDocumentContent, WorkDocumentNode } from './work-types';

let schema: ReturnType<typeof getSchema> | null = null;
let extensions: Extensions | null = null;

export function workDocumentSchema(): ReturnType<typeof getSchema> {
  schema ??= getSchema(documentExtensions());
  return schema;
}

export function serializeWorkDocumentNode(root: WorkDocumentNode): string {
  return generateHTML(
    materializeWindowedDocumentModel(root) as unknown as JSONContent,
    documentExtensions(),
  );
}

export function createWorkDocumentModelFromContent(
  content: WorkDocumentContent,
): WorkDocumentContent {
  const previous = content.model;
  const synchronized = syncDocumentContentFromHtml(content, content.html);
  const document = new DOMParser().parseFromString(
    synchronized.html,
    'text/html',
  );
  const root = windowDocumentModel(
    ProseMirrorDOMParser.fromSchema(workDocumentSchema())
      .parse(document.body)
      .toJSON() as unknown as WorkDocumentNode,
  );
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
    html = serializeWorkDocumentNode(model.root);
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
