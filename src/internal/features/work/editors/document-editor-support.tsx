import type { Editor } from '@tiptap/core';
import {
  Copy,
  Languages,
  MessageSquareText,
  Sparkles,
  TextQuote,
  WandSparkles,
} from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { showToast } from '../../../state/app-state';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import {
  createWorkAgentProposalRequest,
  type WorkAgentProposalRequest,
  type WorkAgentProposalTarget,
} from '../work-agent-proposal';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import type {
  WorkDocumentSelectionAction,
  WorkDocumentSelectionMenuIcon,
  WorkDocumentSelectionMenuItem,
} from '../work-document-selection-menu';
import { DOCUMENT_LAZY_POSITION_BOUNDARY } from '../work-document-lazy-model';
export {
  type DocumentTextStatistics,
  documentTextStatistics,
  documentWordCount,
} from '../work-document-text-statistics';

export function documentEditorSelectionText(
  editor: Pick<Editor, 'state'>,
): string {
  const { from, to, empty } = editor.state.selection;
  if (empty) return '';
  return editor.state.doc
    .textBetween(from, to, '\n')
    .replaceAll(DOCUMENT_LAZY_POSITION_BOUNDARY, ' ')
    .trim();
}

export function documentAgentMenuItems(
  selection: string,
  onAgentRequest: (request: WorkEditorAgentRequest) => void | Promise<void>,
  messages: OfficeMessageCatalog,
  proposalOptions?: {
    target: WorkAgentProposalTarget;
    apply: WorkAgentProposalRequest['apply'];
  },
): WorkspaceContextMenuItem[] {
  return [
    {
      id: 'copy',
      label: officeMessage(messages, 'document.agent.copy'),
      icon: <Copy size={14} />,
      onSelect: () => {
        void copyDocumentSelection(selection, messages);
      },
    },
    {
      id: 'ask',
      label: officeMessage(messages, 'document.agent.ask'),
      icon: <MessageSquareText size={14} />,
      separatorBefore: true,
      onSelect: () =>
        void onAgentRequest({
          instruction: officeMessage(
            messages,
            'document.agent.askInstruction',
          ),
          selection,
        }),
    },
    {
      id: 'summarize',
      label: officeMessage(messages, 'document.agent.summarize'),
      icon: <TextQuote size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction: officeMessage(
            messages,
            'document.agent.summarizeInstruction',
          ),
          selection,
        }),
    },
    {
      id: 'rewrite',
      label: officeMessage(messages, 'document.agent.rewrite'),
      icon: <Sparkles size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction: officeMessage(
            messages,
            'document.agent.rewriteInstruction',
          ),
          selection,
          proposal: proposalOptions
            ? createWorkAgentProposalRequest({
                title: officeMessage(
                  messages,
                  'document.agent.rewriteProposalTitle',
                ),
                description: officeMessage(
                  messages,
                  'document.agent.selectionDescription',
                  { count: String(selection.length) },
                ),
                targets: [proposalOptions.target],
                apply: proposalOptions.apply,
              })
            : undefined,
        }),
    },
    {
      id: 'translate',
      label: officeMessage(messages, 'document.agent.translate'),
      icon: <Languages size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction: officeMessage(
            messages,
            'document.agent.translateInstruction',
          ),
          selection,
          proposal: proposalOptions
            ? createWorkAgentProposalRequest({
                title: officeMessage(
                  messages,
                  'document.agent.translateProposalTitle',
                ),
                description: officeMessage(
                  messages,
                  'document.agent.selectionDescription',
                  { count: String(selection.length) },
                ),
                targets: [proposalOptions.target],
                apply: proposalOptions.apply,
              })
            : undefined,
        }),
    },
  ];
}

export function documentCustomSelectionMenuItems(
  items: readonly WorkDocumentSelectionMenuItem[],
  createAction: () => WorkDocumentSelectionAction,
): WorkspaceContextMenuItem[] {
  return customSelectionMenuItems(items, createAction);
}

interface CustomSelectionAction<Context> {
  context: Context;
  dispose(): void;
}

interface CustomSelectionMenuItem<Context> {
  id: string;
  label: string;
  icon?: WorkDocumentSelectionMenuIcon;
  shortcut?: string;
  ariaKeyShortcut?: string;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
  onSelect(context: Context): void | Promise<void>;
}

export function customSelectionMenuItems<Context>(
  items: readonly CustomSelectionMenuItem<Context>[],
  createAction: () => CustomSelectionAction<Context>,
): WorkspaceContextMenuItem[] {
  const ids = new Set<string>();
  return items.map((item) => {
    const id = item.id.trim();
    if (!id) {
      throw new Error('Document selection menu item IDs cannot be empty.');
    }
    if (ids.has(id)) {
      throw new Error(`Document selection menu item "${id}" is duplicated.`);
    }
    ids.add(id);
    return {
      id,
      label: item.label,
      icon: documentSelectionMenuIcon(item.icon),
      shortcut: item.shortcut,
      ariaKeyShortcut: item.ariaKeyShortcut,
      danger: item.danger,
      disabled: item.disabled,
      separatorBefore: item.separatorBefore,
      onSelect: () => {
        const action = createAction();
        try {
          const pending = item.onSelect(action.context);
          if (pending) {
            void Promise.resolve(pending).finally(action.dispose);
          } else {
            action.dispose();
          }
        } catch (error) {
          action.dispose();
          throw error;
        }
      },
    };
  });
}

function documentSelectionMenuIcon(
  icon: WorkDocumentSelectionMenuIcon | undefined,
) {
  switch (icon) {
    case 'copy':
      return <Copy size={14} />;
    case 'language':
      return <Languages size={14} />;
    case 'message':
      return <MessageSquareText size={14} />;
    case 'quote':
      return <TextQuote size={14} />;
    case 'wand':
      return <WandSparkles size={14} />;
    default:
      return <Sparkles size={14} />;
  }
}

export function documentPageCount(editor: Editor): number {
  let pages = 1;
  const sections: Array<{ breakAfter?: string }> = [];
  editor.state.doc.forEach((node) => {
    if (node.type.name !== 'documentSection') return;
    sections.push(node.attrs);
    node.descendants((child) => {
      if (child.type.name === 'documentNote') return false;
      if (child.type.name === 'pageBreak') pages += 1;
    });
  });
  for (let index = 0; index < sections.length - 1; index += 1) {
    if (
      sections[index].breakAfter !== 'continuous' &&
      sections[index].breakAfter !== 'nextColumn'
    )
      pages += 1;
  }
  return pages;
}

export function documentCurrentPage(editor: Editor): number {
  const selectionPosition = editor.state.selection.from;
  let page = 1;
  let previousBreakAfter: string | undefined;
  let sectionIndex = 0;
  editor.state.doc.forEach((node, position) => {
    if (node.type.name !== 'documentSection') return;
    if (
      sectionIndex > 0 &&
      position < selectionPosition &&
      previousBreakAfter !== 'continuous' &&
      previousBreakAfter !== 'nextColumn'
    ) {
      page += 1;
    }
    if (position < selectionPosition) {
      node.descendants((child, childPosition) => {
        if (child.type.name === 'documentNote') return false;
        if (
          child.type.name === 'pageBreak' &&
          position + childPosition + 1 < selectionPosition
        )
          page += 1;
      });
    }
    previousBreakAfter = node.attrs.breakAfter;
    sectionIndex += 1;
  });
  return page;
}

export function documentHasRefreshableFields(editor: Editor): boolean {
  let found = false;
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'documentField') return !found;
    found = true;
    return false;
  });
  return found;
}

async function copyDocumentSelection(
  selection: string,
  messages: OfficeMessageCatalog,
): Promise<void> {
  try {
    if (!navigator.clipboard?.writeText)
      throw new Error('Clipboard API is unavailable');
    await navigator.clipboard.writeText(selection);
    showToast(officeMessage(messages, 'document.agent.copySuccess'), 'success');
  } catch {
    showToast(
      officeMessage(messages, 'document.agent.copyUnavailable'),
      'error',
    );
  }
}
