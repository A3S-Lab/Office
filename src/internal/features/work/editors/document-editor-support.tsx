import type { Editor } from '@tiptap/core';
import {
  Copy,
  Languages,
  MessageSquareText,
  Sparkles,
  TextQuote,
  WandSparkles,
} from 'lucide-react';
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
  proposalOptions?: {
    target: WorkAgentProposalTarget;
    apply: WorkAgentProposalRequest['apply'];
  },
): WorkspaceContextMenuItem[] {
  return [
    {
      id: 'copy',
      label: '复制',
      icon: <Copy size={14} />,
      onSelect: () => {
        void copyDocumentSelection(selection);
      },
    },
    {
      id: 'ask',
      label: '询问 AI 助手',
      icon: <MessageSquareText size={14} />,
      separatorBefore: true,
      onSelect: () =>
        void onAgentRequest({
          instruction: '请围绕这段选中文本回答我的问题：\n\n问题：',
          selection,
        }),
    },
    {
      id: 'summarize',
      label: '总结选中内容',
      icon: <TextQuote size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction:
            '请用简洁、准确的语言总结这段选中文本，保留关键事实和结论。',
          selection,
        }),
    },
    {
      id: 'rewrite',
      label: '改写得更清晰',
      icon: <Sparkles size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction:
            '请改写这段选中文本，使表达更清晰、自然、专业，并说明主要改动。先提供建议稿，不要直接修改文档。',
          selection,
          proposal: proposalOptions
            ? createWorkAgentProposalRequest({
                title: '审阅文字改写',
                description: `选中文本 · ${selection.length} 个字符`,
                targets: [proposalOptions.target],
                apply: proposalOptions.apply,
              })
            : undefined,
        }),
    },
    {
      id: 'translate',
      label: '翻译选中内容',
      icon: <Languages size={14} />,
      onSelect: () =>
        void onAgentRequest({
          instruction:
            '请翻译这段选中文本。请先判断原语言，并询问或根据上下文确定目标语言；先提供译文，不要直接修改文档。',
          selection,
          proposal: proposalOptions
            ? createWorkAgentProposalRequest({
                title: '审阅翻译建议',
                description: `选中文本 · ${selection.length} 个字符`,
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

async function copyDocumentSelection(selection: string): Promise<void> {
  try {
    if (!navigator.clipboard?.writeText)
      throw new Error('Clipboard API is unavailable');
    await navigator.clipboard.writeText(selection);
    showToast('选中文本已复制', 'success');
  } catch {
    showToast('无法访问剪贴板，请使用系统复制快捷键。', 'error');
  }
}
