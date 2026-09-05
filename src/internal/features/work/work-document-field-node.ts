import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { createDocumentFieldIdentityPlugin } from './work-document-field-identity';
import {
  documentFieldDisplay,
  documentFieldInstruction,
  documentFieldKind,
  documentFieldLabel,
  documentFieldStatisticsFromText,
  documentPageReferenceInstruction,
  docxDocumentFieldKind,
  docxDocumentFieldTarget,
  type WorkDocumentFieldContext,
  type WorkDocumentFieldContextResolver,
  type WorkDocumentFieldInsertOptions,
  type WorkDocumentFieldKind,
  type WorkDocumentFieldRefreshOptions,
} from './work-document-fields';
import { documentPageDescriptors } from './work-document-pages';
import { syncDocumentContentFromHtml } from './work-document-section';
import { createWorkId } from './work-templates';
import type { WorkDocumentContent } from './work-types';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentField: {
      insertDocumentField: (
        kind: WorkDocumentFieldKind,
        options?: WorkDocumentFieldInsertOptions,
      ) => ReturnType;
      refreshDocumentFields: (
        content: WorkDocumentContent,
        options?: WorkDocumentFieldRefreshOptions,
      ) => ReturnType;
    };
  }
}

export const DocumentField = Node.create({
  name: 'documentField',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addProseMirrorPlugins() {
    return [createDocumentFieldIdentityPlugin(this.name)];
  },

  addCommands() {
    return {
      insertDocumentField:
        (kind, options = {}) =>
        (props) =>
          insertDocumentFieldCommand(props, kind, options),
      refreshDocumentFields: (content, options) => (props) =>
        refreshDocumentFieldsCommand(props, content, options),
    };
  },

  addAttributes() {
    return {
      id: hiddenAttribute(''),
      kind: hiddenAttribute('page'),
      instruction: hiddenAttribute('PAGE'),
      display: hiddenAttribute('1'),
      targetId: hiddenAttribute(''),
      targetName: hiddenAttribute(''),
      orphaned: hiddenAttribute(false),
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-field]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const instruction = node.dataset.fieldInstruction?.trim() ?? '';
          const kind =
            documentFieldKind(node.dataset.fieldKind) ??
            docxDocumentFieldKind(instruction) ??
            'page';
          const targetName =
            node.dataset.fieldTargetName ??
            docxDocumentFieldTarget(instruction) ??
            '';
          return {
            id: node.dataset.fieldId ?? '',
            kind,
            instruction:
              kind === 'pageReference' && targetName
                ? documentPageReferenceInstruction(targetName, instruction)
                : instruction || documentFieldInstruction(kind),
            display:
              node.dataset.fieldDisplay?.trim() ||
              node.textContent?.trim() ||
              documentFieldLabel(kind),
            targetId: node.dataset.fieldTargetId ?? '',
            targetName,
            orphaned: node.dataset.fieldOrphaned === 'true',
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const instruction =
      typeof node.attrs.instruction === 'string'
        ? node.attrs.instruction.trim()
        : '';
    const kind =
      documentFieldKind(node.attrs.kind) ??
      docxDocumentFieldKind(instruction) ??
      'page';
    const display =
      typeof node.attrs.display === 'string' && node.attrs.display.trim()
        ? node.attrs.display.trim()
        : documentFieldLabel(kind);
    const targetName =
      typeof node.attrs.targetName === 'string' ? node.attrs.targetName : '';
    const normalizedInstruction =
      kind === 'pageReference' && targetName
        ? documentPageReferenceInstruction(targetName, instruction)
        : instruction || documentFieldInstruction(kind);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-field': 'true',
        'data-field-id': typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-field-kind': kind,
        'data-field-instruction': normalizedInstruction,
        'data-field-display': display,
        ...(kind === 'pageReference' &&
        typeof node.attrs.targetId === 'string' &&
        node.attrs.targetId
          ? { 'data-field-target-id': node.attrs.targetId }
          : {}),
        ...(kind === 'pageReference' &&
        typeof node.attrs.targetName === 'string' &&
        node.attrs.targetName
          ? { 'data-field-target-name': node.attrs.targetName }
          : {}),
        'data-field-orphaned': node.attrs.orphaned ? 'true' : undefined,
        class: 'work-document-field',
        'aria-label': documentFieldLabel(kind),
        title: documentFieldLabel(kind),
      }),
      display,
    ];
  },

  renderText({ node }) {
    const kind = documentFieldKind(node.attrs.kind) ?? 'page';
    return typeof node.attrs.display === 'string' && node.attrs.display.trim()
      ? node.attrs.display.trim()
      : documentFieldLabel(kind);
  },
});

function insertDocumentFieldCommand(
  { dispatch, editor, state, tr }: CommandProps,
  kind: WorkDocumentFieldKind,
  options: WorkDocumentFieldInsertOptions,
): boolean {
  const fieldType = editor.schema.nodes.documentField;
  if (!fieldType) return false;
  const targetName =
    typeof options.targetName === 'string' ? options.targetName.trim() : '';
  if (kind === 'pageReference' && !targetName) return false;
  if (!dispatch) return true;
  const instruction = documentFieldInstruction(kind, options);
  const statistics = documentFieldStatisticsFromText(
    state.doc.textBetween(0, state.doc.content.size, '\n', '\uFFFC'),
  );
  const context = {
    ...fallbackContext(state),
    ...statistics,
  } satisfies WorkDocumentFieldContext;
  tr.replaceSelectionWith(
    fieldType.create({
      id: createWorkId('field'),
      kind,
      instruction,
      display: documentFieldDisplay(kind, context, instruction),
      targetId: options.targetId?.trim() ?? '',
      targetName,
      orphaned:
        kind === 'pageReference' && !documentBookmarkExists(state, options),
    }),
    false,
  );
  tr.scrollIntoView();
  return true;
}

function refreshDocumentFieldsCommand(
  { editor, state, tr }: CommandProps,
  content: WorkDocumentContent,
  options: WorkDocumentFieldRefreshOptions = {},
): boolean {
  const fallbackFields = options.resolveContext
    ? []
    : documentPageDescriptors(
        syncDocumentContentFromHtml(content, editor.getHTML()),
      ).flatMap((page) =>
        page.segments.flatMap((segment) => {
          const document = new DOMParser().parseFromString(
            segment.html,
            'text/html',
          );
          return Array.from(
            document.body.querySelectorAll<HTMLElement>(
              '[data-document-field]',
            ),
          ).map((element) => ({
            id: element.dataset.fieldId?.trim() ?? '',
            display:
              element.dataset.fieldDisplay?.trim() ||
              element.textContent?.trim() ||
              '',
          }));
        }),
      );
  const fallbackById = new Map(
    fallbackFields.flatMap(({ id, display }) => (id ? [[id, display]] : [])),
  );
  const now = validDate(options.now) ?? new Date();
  const statistics = documentFieldStatisticsFromText(
    state.doc.textBetween(0, state.doc.content.size, '\n', '\uFFFC'),
  );
  let fieldIndex = 0;
  state.doc.descendants((node, position) => {
    if (node.type.name !== 'documentField') return;
    const kind =
      documentFieldKind(node.attrs.kind) ??
      docxDocumentFieldKind(stringAttribute(node.attrs.instruction));
    const fallback = fallbackFields[fieldIndex]?.display;
    fieldIndex += 1;
    if (!kind) return;
    if (options.updateClock === false && (kind === 'date' || kind === 'time')) {
      return;
    }
    const context = options.resolveContext?.(position);
    const display = context
      ? documentFieldDisplay(
          kind,
          {
            ...context,
            wordCount: context.wordCount ?? statistics.wordCount,
            characterCount: context.characterCount ?? statistics.characterCount,
            referencePageNumber:
              kind === 'pageReference'
                ? resolveBookmarkPageNumber(state, node, options.resolveContext)
                : context.referencePageNumber,
            now: validDate(context.now) ?? now,
          },
          stringAttribute(node.attrs.instruction),
          stringAttribute(node.attrs.display),
        )
      : (fallbackById.get(stringAttribute(node.attrs.id)) ?? fallback);
    if (!display) return;
    const targetName =
      stringAttribute(node.attrs.targetName) ||
      docxDocumentFieldTarget(stringAttribute(node.attrs.instruction)) ||
      '';
    const orphaned =
      kind === 'pageReference'
        ? !bookmarkBoundaryForField(state, node, targetName)
        : false;
    if (
      node.attrs.display === display &&
      node.attrs.targetName === targetName &&
      node.attrs.orphaned === orphaned
    ) {
      return;
    }
    tr.setNodeMarkup(position, undefined, {
      ...node.attrs,
      display,
      targetName,
      orphaned,
    });
  });
  if (tr.docChanged && options.addToHistory === false) {
    tr.setMeta('addToHistory', false);
  }
  return tr.docChanged;
}

function fallbackContext(state: Editor['state']): WorkDocumentFieldContext {
  let totalPages = 0;
  let sectionNumber = 1;
  let sectionPages = 1;
  let pageNumber = 1;
  let pagesBeforeSelection = 0;
  let foundSelection = false;
  state.doc.forEach((section, offset, index) => {
    if (section.type.name !== 'documentSection') return;
    const pages = countPageBreaks(section) + 1;
    totalPages += pages;
    if (!foundSelection && state.selection.from <= offset + section.nodeSize) {
      sectionNumber = index + 1;
      sectionPages = pages;
      pageNumber =
        pagesBeforeSelection +
        1 +
        countPageBreaksBefore(section, state.selection.from - offset - 1);
      foundSelection = true;
    }
    if (!foundSelection) pagesBeforeSelection += pages;
  });
  return {
    pageNumber,
    totalPages: Math.max(1, totalPages),
    sectionNumber,
    sectionPages,
  };
}

function countPageBreaks(node: ProseMirrorNode): number {
  let count = 0;
  node.descendants((child) => {
    if (child.type.name === 'pageBreak') count += 1;
  });
  return count;
}

function countPageBreaksBefore(
  node: ProseMirrorNode,
  position: number,
): number {
  let count = 0;
  node.descendants((child, offset) => {
    if (offset >= position) return false;
    if (child.type.name === 'pageBreak') count += 1;
    return true;
  });
  return count;
}

function hiddenAttribute(defaultValue: unknown) {
  return {
    default: defaultValue,
    rendered: false,
  };
}

function validDate(value: Date | undefined): Date | null {
  return value && Number.isFinite(value.getTime()) ? value : null;
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function documentBookmarkExists(
  state: Editor['state'],
  options: WorkDocumentFieldInsertOptions,
): boolean {
  const targetId = stringAttribute(options.targetId);
  const targetName = stringAttribute(options.targetName).toLowerCase();
  let found = false;
  state.doc.descendants((node) => {
    if (
      node.type.name === 'documentBookmarkBoundary' &&
      node.attrs.kind === 'start' &&
      ((targetId && stringAttribute(node.attrs.id) === targetId) ||
        (targetName &&
          stringAttribute(node.attrs.name).toLowerCase() === targetName))
    ) {
      found = true;
      return false;
    }
    return !found;
  });
  return found;
}

function bookmarkBoundaryForField(
  state: Editor['state'],
  node: ProseMirrorNode,
  targetName: string,
): ProseMirrorNode | null {
  const targetId = stringAttribute(node.attrs.targetId);
  const normalizedName = targetName.toLowerCase();
  let found: ProseMirrorNode | null = null;
  state.doc.descendants((candidate) => {
    if (
      found ||
      candidate.type.name !== 'documentBookmarkBoundary' ||
      candidate.attrs.kind !== 'start'
    ) {
      return found === null;
    }
    if (
      (targetId && stringAttribute(candidate.attrs.id) === targetId) ||
      (normalizedName &&
        stringAttribute(candidate.attrs.name).toLowerCase() === normalizedName)
    ) {
      found = candidate;
      return false;
    }
    return true;
  });
  return found;
}

function resolveBookmarkPageNumber(
  state: Editor['state'],
  node: ProseMirrorNode,
  resolveContext: WorkDocumentFieldContextResolver | undefined,
): number | null {
  if (!resolveContext) return null;
  const targetId = stringAttribute(node.attrs.targetId);
  const targetName =
    stringAttribute(node.attrs.targetName).toLowerCase() ||
    (
      docxDocumentFieldTarget(stringAttribute(node.attrs.instruction)) ?? ''
    ).toLowerCase();
  let page: number | null = null;
  state.doc.descendants((candidate, position) => {
    if (
      page !== null ||
      candidate.type.name !== 'documentBookmarkBoundary' ||
      candidate.attrs.kind !== 'start'
    ) {
      return page === null;
    }
    if (
      (targetId && stringAttribute(candidate.attrs.id) === targetId) ||
      (targetName &&
        stringAttribute(candidate.attrs.name).toLowerCase() === targetName)
    ) {
      page = resolveContext(position)?.pageNumber ?? null;
      return false;
    }
    return true;
  });
  return page;
}
