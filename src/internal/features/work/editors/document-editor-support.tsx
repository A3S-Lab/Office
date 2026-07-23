import type { Editor } from '@tiptap/core';
import {
  Copy,
  Languages,
  MessageSquareText,
  Sparkles,
  TextQuote,
} from 'lucide-react';
import { showToast } from '../../../state/app-state';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import {
  createWorkAgentProposalRequest,
  type WorkAgentProposalRequest,
  type WorkAgentProposalTarget,
} from '../work-agent-proposal';
import type { WorkEditorAgentRequest } from '../work-agent-request';

export const MIN_DOCUMENT_ZOOM = 50;
export const MAX_DOCUMENT_ZOOM = 200;

export function documentEditorSelectionText(
  editor: Pick<Editor, 'state'>,
): string {
  const { from, to, empty } = editor.state.selection;
  if (empty) return '';
  return editor.state.doc.textBetween(from, to, '\n').trim();
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

export function plainTextAsHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replace(/\r?\n/g, '<br>');
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

export function documentWordCount(value: string): number {
  return Array.from(
    value.matchAll(
      /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]|[\p{L}\p{N}]+/gu,
    ),
  ).length;
}

export function clampDocumentZoom(zoom: number): number {
  return Math.min(
    MAX_DOCUMENT_ZOOM,
    Math.max(MIN_DOCUMENT_ZOOM, Math.round(zoom)),
  );
}

export function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(String(reader.result)));
    reader.addEventListener('error', () =>
      reject(reader.error ?? new Error('Image could not be read')),
    );
    reader.readAsDataURL(file);
  });
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
