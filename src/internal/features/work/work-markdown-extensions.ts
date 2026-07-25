import { Extension, type Extensions } from '@tiptap/core';
import { TaskItem, TaskList } from '@tiptap/extension-list';
import Image from '@tiptap/extension-image';
import Placeholder from '@tiptap/extension-placeholder';
import { TableKit } from '@tiptap/extension-table';
import { Markdown } from '@tiptap/markdown';
import StarterKit from '@tiptap/starter-kit';

export interface SetWorkMarkdownOptions {
  emitUpdate?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    workMarkdown: {
      setWorkMarkdown: (
        markdown: string,
        options?: SetWorkMarkdownOptions,
      ) => ReturnType;
    };
  }
}

export const WorkMarkdown = Extension.create({
  name: 'workMarkdown',

  addCommands() {
    return {
      setWorkMarkdown:
        (markdown, options = {}) =>
        ({ commands }) =>
          commands.setContent(markdown, {
            contentType: 'markdown',
            emitUpdate: options.emitUpdate ?? true,
          }),
    };
  },
});

export function createWorkMarkdownExtensions(): Extensions {
  return [
    StarterKit.configure({
      link: {
        autolink: true,
        defaultProtocol: 'https',
        openOnClick: false,
      },
      underline: false,
    }),
    TaskList,
    TaskItem.configure({
      nested: true,
      HTMLAttributes: {
        'data-type': 'taskItem',
      },
      a11y: {
        checkboxLabel: markdownTaskCheckboxLabel,
      },
    }),
    Image.configure({
      allowBase64: false,
      inline: true,
    }),
    TableKit.configure({
      table: {
        allowTableNodeSelection: true,
        resizable: false,
      },
    }),
    Placeholder.configure({
      placeholder: '开始写 Markdown…',
    }),
    Markdown.configure({
      indentation: { style: 'space', size: 2 },
      markedOptions: {
        gfm: true,
        breaks: false,
        pedantic: false,
      },
    }),
    WorkMarkdown,
  ];
}

export function markdownTaskCheckboxLabel(node: {
  attrs: { checked?: boolean };
  textContent: string;
}): string {
  const label = node.textContent.trim() || '任务';
  return `${node.attrs.checked ? '已完成' : '未完成'}：${label}`;
}
