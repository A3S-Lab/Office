import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import {
  highlightMarkdownCode,
  presentMarkdownDiagram,
  renderMarkdownDiagram,
  resolveMarkdownCodeLanguage,
} from '../src/internal/features/work/editors/markdown-code-render';
import { createWorkMarkdownExtensions } from '../src/internal/features/work/work-markdown-extensions';

describe('Markdown code rendering', () => {
  test('resolves aliases and keeps mermaid off the code highlighter', () => {
    expect(resolveMarkdownCodeLanguage('ts')).toBe('typescript');
    expect(resolveMarkdownCodeLanguage('mermaid')).toBe('mermaid');
    expect(resolveMarkdownCodeLanguage('plaintext')).toBe('text');
  });

  test('highlights fenced code with both theme colors', async () => {
    const tokens = await highlightMarkdownCode('const n = 1;', 'ts');
    expect(tokens.some((token) => token.light.startsWith('#'))).toBe(true);
    expect(tokens.some((token) => token.dark.startsWith('#'))).toBe(true);
    expect(await highlightMarkdownCode('notes', 'text')).toEqual([]);
  });

  test('renders a supported diagram without a remote font', () => {
    const svg = renderMarkdownDiagram('flowchart LR\n  A[Start] --> B[End]');
    expect(svg).toContain('--fg:var(--a3s-ink)');
    expect(svg).not.toContain('fonts.googleapis.com');
    expect(renderMarkdownDiagram('pie title mix\n  "A" : 1')).toBeUndefined();
  });

  test('keeps diagram theme variables while dropping the canvas background', () => {
    const svg = presentMarkdownDiagram(
      '<svg viewBox="0 0 10 12" width="100%" style="--bg:transparent;--fg:var(--a3s-ink);--shiki-light-bg:#fff"></svg>',
    );
    expect(svg).toContain('width="10"');
    expect(svg).toContain('--fg:var(--a3s-ink)');
  });

  test('renders Shiki tokens and a mermaid diagram without changing the source', async () => {
    const editor = new Editor({
      extensions: createWorkMarkdownExtensions(),
      content: [
        '```ts',
        'const n = 1;',
        '```',
        '',
        '```mermaid',
        'flowchart LR',
        '  A[Start] --> B[End]',
        '```',
      ].join('\n'),
      contentType: 'markdown',
    });

    await new Promise((resolve) => setTimeout(resolve, 400));

    expect(editor.view.dom.querySelector('.work-markdown-token')).toBeTruthy();
    expect(
      editor.view.dom.querySelector(
        '[data-markdown-diagram="rendered"] .work-markdown-mermaid svg',
      ),
    ).toBeTruthy();
    expect(editor.getMarkdown()).toContain('```ts');
    expect(editor.getMarkdown()).toContain('```mermaid');
    expect(editor.getMarkdown()).toContain('const n = 1;');
    editor.destroy();
  });
});
