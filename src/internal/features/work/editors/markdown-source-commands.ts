export interface MarkdownSourceSelection {
  start: number;
  end: number;
  direction: 'backward' | 'forward' | 'none';
}

export interface MarkdownSourceEdit {
  markdown: string;
  selection: MarkdownSourceSelection;
}

export interface MarkdownSourceSelectionState {
  markdown: string;
  selection: MarkdownSourceSelection;
  text: string;
}

export type MarkdownSourceCommand =
  | 'blockquote'
  | 'bold'
  | 'bullet-list'
  | 'code'
  | 'code-block'
  | 'heading-1'
  | 'heading-2'
  | 'heading-3'
  | 'horizontal-rule'
  | 'italic'
  | 'ordered-list'
  | 'paragraph'
  | 'strike'
  | 'table'
  | 'task-list';

const INLINE_WRAPPERS = {
  bold: '**',
  code: '`',
  italic: '*',
  strike: '~~',
} as const;

const MARKDOWN_TABLE = [
  '| Column 1 | Column 2 | Column 3 |',
  '| --- | --- | --- |',
  '|  |  |  |',
  '|  |  |  |',
].join('\n');

export function applyMarkdownSourceCommand(
  markdown: string,
  requestedSelection: MarkdownSourceSelection,
  command: MarkdownSourceCommand,
): MarkdownSourceEdit {
  const selection = normalizeMarkdownSourceSelection(
    requestedSelection,
    markdown.length,
  );
  if (isInlineMarkdownSourceCommand(command)) {
    return toggleInlineWrapper(
      markdown,
      selection,
      INLINE_WRAPPERS[command],
      command,
    );
  }
  if (command === 'horizontal-rule') {
    return insertMarkdownBlock(markdown, selection, '---', undefined, true);
  }
  if (command === 'table') {
    return insertMarkdownBlock(markdown, selection, MARKDOWN_TABLE, {
      start: 2,
      end: 10,
    });
  }
  if (command === 'code-block') {
    return toggleCodeBlock(markdown, selection);
  }
  return transformSelectedLines(markdown, selection, command);
}

function isInlineMarkdownSourceCommand(
  command: MarkdownSourceCommand,
): command is keyof typeof INLINE_WRAPPERS {
  return Object.hasOwn(INLINE_WRAPPERS, command);
}

export function replaceMarkdownSourceSelection(
  markdown: string,
  requestedSelection: MarkdownSourceSelection,
  replacement: string,
  selectedRange: { start: number; end: number } = {
    start: 0,
    end: replacement.length,
  },
): MarkdownSourceEdit {
  const selection = normalizeMarkdownSourceSelection(
    requestedSelection,
    markdown.length,
  );
  const nextMarkdown = `${markdown.slice(0, selection.start)}${replacement}${markdown.slice(selection.end)}`;
  return {
    markdown: nextMarkdown,
    selection: {
      start: selection.start + selectedRange.start,
      end: selection.start + selectedRange.end,
      direction: 'none',
    },
  };
}

function toggleInlineWrapper(
  markdown: string,
  selection: MarkdownSourceSelection,
  wrapper: string,
  command: MarkdownSourceCommand,
): MarkdownSourceEdit {
  const selected = markdown.slice(selection.start, selection.end);
  const selectedHasWrapper =
    selected.length >= wrapper.length * 2 &&
    selected.startsWith(wrapper) &&
    selected.endsWith(wrapper);
  if (selectedHasWrapper) {
    const replacement = selected.slice(wrapper.length, -wrapper.length);
    return replaceMarkdownSourceSelection(markdown, selection, replacement);
  }

  const before = markdown.slice(0, selection.start);
  const after = markdown.slice(selection.end);
  const wrapperOutsideSelection =
    before.endsWith(wrapper) &&
    after.startsWith(wrapper) &&
    !(command === 'italic' && before.endsWith('**') && after.startsWith('**'));
  if (wrapperOutsideSelection) {
    return {
      markdown: `${before.slice(0, -wrapper.length)}${selected}${after.slice(wrapper.length)}`,
      selection: {
        start: selection.start - wrapper.length,
        end: selection.end - wrapper.length,
        direction: selection.direction,
      },
    };
  }

  const replacement = `${wrapper}${selected}${wrapper}`;
  return replaceMarkdownSourceSelection(markdown, selection, replacement, {
    start: wrapper.length,
    end: wrapper.length + selected.length,
  });
}

function transformSelectedLines(
  markdown: string,
  selection: MarkdownSourceSelection,
  command: Exclude<
    MarkdownSourceCommand,
    | 'bold'
    | 'code'
    | 'code-block'
    | 'horizontal-rule'
    | 'italic'
    | 'strike'
    | 'table'
  >,
): MarkdownSourceEdit {
  const range = selectedLineRange(markdown, selection);
  const lines = markdown.slice(range.start, range.end).split('\n');
  const nonEmptyLines = lines.filter((line) => line.length > 0);
  const everyNonEmptyLine = (pattern: RegExp) =>
    nonEmptyLines.length > 0 &&
    nonEmptyLines.every((line) => pattern.test(line));
  let transformed: string[];

  if (command.startsWith('heading-')) {
    const level = Number(command.at(-1));
    const prefix = `${'#'.repeat(level)} `;
    transformed = lines.map((line) =>
      line ? `${prefix}${stripBlockPrefix(line)}` : prefix,
    );
  } else if (command === 'paragraph') {
    transformed = lines.map((line) => line.replace(/^#{1,6}\s+/, ''));
  } else if (command === 'blockquote') {
    const remove = everyNonEmptyLine(/^>\s+/);
    transformed = lines.map((line) =>
      remove ? line.replace(/^>\s+/, '') : line ? `> ${line}` : '> ',
    );
  } else {
    const pattern =
      command === 'bullet-list'
        ? /^[-+*]\s+(?!\[[ xX]\]\s+)/
        : command === 'ordered-list'
          ? /^\d+[.)]\s+/
          : /^[-+*]\s+\[[ xX]\]\s+/;
    const remove = everyNonEmptyLine(pattern);
    transformed = lines.map((line, index) => {
      const text = stripListPrefix(line);
      if (remove) return text;
      if (command === 'ordered-list') return `${index + 1}. ${text}`;
      if (command === 'task-list') return `- [ ] ${text}`;
      return `- ${text}`;
    });
  }

  const replacement = transformed.join('\n');
  return {
    markdown: `${markdown.slice(0, range.start)}${replacement}${markdown.slice(range.end)}`,
    selection: {
      start: range.start,
      end: range.start + replacement.length,
      direction: selection.direction,
    },
  };
}

function toggleCodeBlock(
  markdown: string,
  selection: MarkdownSourceSelection,
): MarkdownSourceEdit {
  const selected = markdown.slice(selection.start, selection.end);
  if (selected.startsWith('```\n') && selected.endsWith('\n```')) {
    const replacement = selected.slice(4, -4);
    return replaceMarkdownSourceSelection(markdown, selection, replacement);
  }
  const before = markdown.slice(0, selection.start);
  const after = markdown.slice(selection.end);
  if (before.endsWith('```\n') && after.startsWith('\n```')) {
    return {
      markdown: `${before.slice(0, -4)}${selected}${after.slice(4)}`,
      selection: {
        start: selection.start - 4,
        end: selection.end - 4,
        direction: selection.direction,
      },
    };
  }
  return replaceMarkdownSourceSelection(
    markdown,
    selection,
    `\`\`\`\n${selected}\n\`\`\``,
    { start: 4, end: 4 + selected.length },
  );
}

function insertMarkdownBlock(
  markdown: string,
  selection: MarkdownSourceSelection,
  block: string,
  selectedRange?: { start: number; end: number },
  appendBlank = false,
): MarkdownSourceEdit {
  const before = markdown.slice(0, selection.start);
  const after = markdown.slice(selection.end);
  const leading = before
    ? before.endsWith('\n\n')
      ? ''
      : before.endsWith('\n')
        ? '\n'
        : '\n\n'
    : '';
  const trailing = after
    ? after.startsWith('\n\n')
      ? ''
      : after.startsWith('\n')
        ? '\n'
        : '\n\n'
    : appendBlank
      ? '\n\n'
      : '';
  const replacement = `${leading}${block}${trailing}`;
  const blockStart = selection.start + leading.length;
  return {
    markdown: `${before}${replacement}${after}`,
    selection: selectedRange
      ? {
          start: blockStart + selectedRange.start,
          end: blockStart + selectedRange.end,
          direction: 'none',
        }
      : {
          start: selection.start + replacement.length,
          end: selection.start + replacement.length,
          direction: 'none',
        },
  };
}

function selectedLineRange(
  markdown: string,
  selection: MarkdownSourceSelection,
): { start: number; end: number } {
  const start =
    selection.start === 0
      ? 0
      : markdown.lastIndexOf('\n', selection.start - 1) + 1;
  const inclusiveEnd =
    selection.end > selection.start && markdown[selection.end - 1] === '\n'
      ? selection.end - 1
      : selection.end;
  const newline = markdown.indexOf('\n', inclusiveEnd);
  return { start, end: newline < 0 ? markdown.length : newline };
}

function stripBlockPrefix(line: string): string {
  return line.replace(
    /^(?:#{1,6}\s+|[-+*]\s+(?:\[[ xX]\]\s+)?|\d+[.)]\s+|>\s+)/,
    '',
  );
}

function stripListPrefix(line: string): string {
  return line.replace(/^(?:[-+*]\s+(?:\[[ xX]\]\s+)?|\d+[.)]\s+)/, '');
}

function normalizeMarkdownSourceSelection(
  selection: MarkdownSourceSelection,
  markdownLength: number,
): MarkdownSourceSelection {
  const requestedStart = Math.min(selection.start, selection.end);
  const requestedEnd = Math.max(selection.start, selection.end);
  return {
    start: Math.max(0, Math.min(markdownLength, requestedStart)),
    end: Math.max(0, Math.min(markdownLength, requestedEnd)),
    direction: selection.direction,
  };
}
