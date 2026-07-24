import type { Editor, Extensions } from '@tiptap/core';
import Color from '@tiptap/extension-color';
import Placeholder from '@tiptap/extension-placeholder';
import { TextStyle } from '@tiptap/extension-text-style';
import FontFamily from '@tiptap/extension-text-style/font-family';
import FontSize from '@tiptap/extension-text-style/font-size';
import Underline from '@tiptap/extension-underline';
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { useEffect, useMemo, useRef } from 'react';
import type { WorkSlideElement, WorkSlideTextRun } from '../work-types';

export interface PresentationTextValue {
  text: string;
  textRuns?: WorkSlideTextRun[];
}

export interface PresentationTextEditorProps {
  element: WorkSlideElement;
  onChange: (value: PresentationTextValue) => void;
  onEditorChange?: (editor: Editor | null) => void;
  onSelectionChange?: () => void;
}

export function PresentationTextEditor({
  element,
  onChange,
  onEditorChange,
  onSelectionChange,
}: PresentationTextEditorProps) {
  const elementRef = useRef(element);
  const onChangeRef = useRef(onChange);
  const onEditorChangeRef = useRef(onEditorChange);
  const onSelectionChangeRef = useRef(onSelectionChange);
  const appliedSignatureRef = useRef(presentationTextElementSignature(element));
  const initialContentRef = useRef(presentationTextElementHtml(element));
  elementRef.current = element;
  onChangeRef.current = onChange;
  onEditorChangeRef.current = onEditorChange;
  onSelectionChangeRef.current = onSelectionChange;
  const extensions = useMemo(
    () =>
      createPresentationTextEditorExtensions(
        element.placeholder?.prompt ?? '输入文字',
      ),
    [element.placeholder?.prompt],
  );
  const editor = useEditor({
    extensions,
    content: initialContentRef.current,
    editorProps: {
      attributes: {
        'aria-label': '幻灯片文本',
        'aria-multiline': 'true',
        'data-presentation-text-engine': 'tiptap',
        'data-slide-editor': 'true',
        role: 'textbox',
        spellcheck: 'true',
      },
    },
    onCreate: ({ editor: current }) => {
      applyPresentationTextStoredMarks(current, elementRef.current);
    },
    onSelectionUpdate: () => onSelectionChangeRef.current?.(),
    onUpdate: ({ editor: current }) => {
      const value = presentationTextValue(current, elementRef.current);
      const signature = presentationTextElementSignature({
        ...elementRef.current,
        ...value,
      });
      if (signature === appliedSignatureRef.current) return;
      appliedSignatureRef.current = signature;
      onChangeRef.current(value);
    },
  });

  useEffect(() => {
    if (!editor) return;
    onEditorChangeRef.current?.(editor);
    return () => onEditorChangeRef.current?.(null);
  }, [editor]);

  useEffect(() => {
    if (!editor || editor.isDestroyed) return;
    const signature = presentationTextElementSignature(element);
    if (signature === appliedSignatureRef.current) return;
    appliedSignatureRef.current = signature;
    editor.commands.setContent(presentationTextElementHtml(element), {
      emitUpdate: false,
    });
    applyPresentationTextStoredMarks(editor, element);
  }, [editor, element]);

  if (!editor) return null;
  return (
    <EditorContent
      className="work-slide-rich-editor"
      data-slide-editor
      editor={editor}
      style={presentationTextBaseStyle(element)}
    />
  );
}

export function createPresentationTextEditorExtensions(
  placeholder = '输入文字',
): Extensions {
  return [
    StarterKit.configure({
      link: {
        autolink: false,
        defaultProtocol: 'https',
        openOnClick: false,
      },
      underline: false,
    }),
    TextStyle,
    FontFamily,
    FontSize,
    Color,
    Underline,
    Placeholder.configure({ placeholder }),
  ];
}

export function presentationTextElementHtml(element: WorkSlideElement): string {
  const runs = element.textRuns?.length
    ? element.textRuns
    : [{ text: element.text } satisfies WorkSlideTextRun];
  const content = runs
    .map((run) => presentationTextRunHtml(run, element))
    .join('');
  return `<p>${content}</p>`;
}

export function presentationTextValue(
  editor: Editor,
  element: WorkSlideElement,
): PresentationTextValue {
  const runs: WorkSlideTextRun[] = [];
  editor.state.doc.forEach((block, _offset, blockIndex) => {
    if (blockIndex > 0) appendPresentationTextRun(runs, '\n', defaultRun());
    block.descendants((node) => {
      if (node.isText && node.text) {
        appendPresentationTextRun(
          runs,
          node.text,
          presentationTextRunFromMarks(node.marks, element),
        );
      } else if (node.type.name === 'hardBreak') {
        appendPresentationTextRun(
          runs,
          '\n',
          presentationTextRunFromMarks(node.marks, element),
        );
      }
    });
  });
  return {
    text: runs.map((run) => run.text).join(''),
    textRuns: runs.length ? runs : undefined,
  };
}

export function presentationTextToolbarState(
  editor: Editor,
  element: WorkSlideElement,
): WorkSlideElement {
  const textStyle = editor.getAttributes('textStyle');
  const link = editor.getAttributes('link');
  return {
    ...element,
    bold: editor.isActive('bold'),
    color:
      typeof textStyle.color === 'string' ? textStyle.color : element.color,
    fontFamily:
      typeof textStyle.fontFamily === 'string'
        ? textStyle.fontFamily
        : element.fontFamily,
    fontSize: presentationFontSize(textStyle.fontSize) ?? element.fontSize,
    href: typeof link.href === 'string' ? link.href : element.href,
    italic: editor.isActive('italic'),
    underline: editor.isActive('underline'),
  };
}

export function applyPresentationTextFormatting(
  editor: Editor,
  patch: Partial<WorkSlideElement>,
  options: { restoreFocus?: boolean } = {},
): boolean {
  let handled = false;
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  if ('fontSize' in patch && Number.isFinite(patch.fontSize)) {
    chain.setFontSize(presentationFontSizeCss(Number(patch.fontSize)));
    handled = true;
  }
  if ('fontFamily' in patch) {
    if (patch.fontFamily) chain.setFontFamily(patch.fontFamily);
    else chain.unsetFontFamily();
    handled = true;
  }
  if ('color' in patch) {
    if (patch.color) chain.setColor(patch.color);
    else chain.unsetColor();
    handled = true;
  }
  if ('bold' in patch) {
    if (patch.bold) chain.setBold();
    else chain.unsetBold();
    handled = true;
  }
  if ('italic' in patch) {
    if (patch.italic) chain.setItalic();
    else chain.unsetItalic();
    handled = true;
  }
  if ('underline' in patch) {
    if (patch.underline) chain.setUnderline();
    else chain.unsetUnderline();
    handled = true;
  }
  if ('href' in patch) {
    if (patch.href) chain.setLink({ href: patch.href });
    else chain.unsetLink();
    handled = true;
  }
  return handled ? chain.run() : false;
}

function presentationTextRunHtml(
  run: WorkSlideTextRun,
  element: WorkSlideElement,
): string {
  const fontSize = run.fontSize ?? element.fontSize;
  const fontFamily = run.fontFamily ?? element.fontFamily ?? 'Aptos';
  const color = run.color ?? element.color;
  const bold = run.bold ?? element.bold;
  const italic = run.italic ?? Boolean(element.italic);
  const underline = run.underline ?? Boolean(element.underline);
  const style = [
    `color: ${color}`,
    `font-family: ${fontFamily}`,
    `font-size: ${presentationFontSizeCss(fontSize)}`,
  ].join('; ');
  let content = `<span style="${escapeHtmlAttribute(style)}">${escapeHtml(
    run.text,
  ).replaceAll('\n', '<br>')}</span>`;
  if (bold) content = `<strong>${content}</strong>`;
  if (italic) content = `<em>${content}</em>`;
  if (underline) content = `<u>${content}</u>`;
  if (run.href)
    content = `<a href="${escapeHtmlAttribute(run.href)}">${content}</a>`;
  return content;
}

function presentationTextRunFromMarks(
  marks: readonly {
    type: { name: string };
    attrs: Record<string, unknown>;
  }[],
  element: WorkSlideElement,
): Omit<WorkSlideTextRun, 'text'> {
  const textStyle = marks.find((mark) => mark.type.name === 'textStyle');
  const link = marks.find((mark) => mark.type.name === 'link');
  return {
    bold: marks.some((mark) => mark.type.name === 'bold'),
    color:
      typeof textStyle?.attrs.color === 'string'
        ? textStyle.attrs.color
        : element.color,
    fontFamily:
      typeof textStyle?.attrs.fontFamily === 'string'
        ? textStyle.attrs.fontFamily
        : (element.fontFamily ?? 'Aptos'),
    fontSize:
      presentationFontSize(textStyle?.attrs.fontSize) ?? element.fontSize,
    ...(typeof link?.attrs.href === 'string' ? { href: link.attrs.href } : {}),
    italic: marks.some((mark) => mark.type.name === 'italic'),
    underline: marks.some((mark) => mark.type.name === 'underline'),
  };
}

function appendPresentationTextRun(
  runs: WorkSlideTextRun[],
  text: string,
  formatting: Omit<WorkSlideTextRun, 'text'>,
): void {
  if (!text) return;
  const previous = runs.at(-1);
  if (
    previous &&
    presentationTextRunKey(previous) === presentationTextRunKey(formatting)
  ) {
    previous.text += text;
    return;
  }
  runs.push({ text, ...formatting });
}

function presentationTextRunKey(run: Omit<WorkSlideTextRun, 'text'>): string {
  return JSON.stringify({
    bold: run.bold,
    color: run.color,
    fontFamily: run.fontFamily,
    fontSize: run.fontSize,
    href: run.href,
    italic: run.italic,
    underline: run.underline,
  });
}

function defaultRun(): Omit<WorkSlideTextRun, 'text'> {
  return {
    bold: false,
    color: '#172033',
    fontFamily: 'Aptos',
    fontSize: 24,
    italic: false,
    underline: false,
  };
}

function applyPresentationTextStoredMarks(
  editor: Editor,
  element: WorkSlideElement,
): void {
  if (editor.state.doc.textContent) return;
  const chain = editor
    .chain()
    .setColor(element.color)
    .setFontFamily(element.fontFamily ?? 'Aptos')
    .setFontSize(presentationFontSizeCss(element.fontSize));
  if (element.bold) chain.setBold();
  if (element.italic) chain.setItalic();
  if (element.underline) chain.setUnderline();
  chain.run();
}

function presentationTextElementSignature(element: WorkSlideElement): string {
  return JSON.stringify({
    bold: element.bold,
    color: element.color,
    fontFamily: element.fontFamily,
    fontSize: element.fontSize,
    href: element.href,
    italic: element.italic,
    text: element.text,
    textRuns: element.textRuns,
    underline: element.underline,
  });
}

function presentationTextBaseStyle(
  element: WorkSlideElement,
): React.CSSProperties {
  return {
    color: element.color,
    fontFamily: element.fontFamily,
    fontSize: presentationFontSizeCss(element.fontSize),
    fontStyle: element.italic ? 'italic' : undefined,
    fontWeight: element.bold ? 700 : 400,
    textAlign: element.align,
    textDecoration: element.underline ? 'underline' : undefined,
  };
}

function presentationFontSizeCss(value: number): string {
  const size = Math.min(400, Math.max(1, value));
  return `clamp(6px, ${size / 10}cqw, ${size}px)`;
}

function presentationFontSize(value: unknown): number | undefined {
  if (typeof value !== 'string') return undefined;
  const match =
    /,\s*(\d+(?:\.\d+)?)px\)\s*$/i.exec(value.trim()) ??
    /^(\d+(?:\.\d+)?)px$/i.exec(value.trim());
  if (!match) return undefined;
  const size = Number(match[1]);
  return Number.isFinite(size) ? size : undefined;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replace(/\r\n?/g, '\n');
}

function escapeHtmlAttribute(value: string): string {
  return escapeHtml(value).replaceAll('\n', '&#10;');
}
