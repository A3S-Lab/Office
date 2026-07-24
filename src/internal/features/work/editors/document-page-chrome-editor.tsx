import type { Editor, Extensions } from '@tiptap/core';
import Color from '@tiptap/extension-color';
import Image from '@tiptap/extension-image';
import Placeholder from '@tiptap/extension-placeholder';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import Underline from '@tiptap/extension-underline';
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  Image as ImageIcon,
  Italic,
  Link2,
  Redo2,
  Underline as UnderlineIcon,
  Undo2,
} from 'lucide-react';
import { type ReactNode, useEffect, useMemo, useRef, useState } from 'react';
import { sanitizeDocumentPageChromeHtml } from '../work-document-page-chrome';
import {
  OfficeColorPicker,
  OfficeFileInput,
  useOfficeDialog,
} from './office-controls';

export type DocumentPageChromeAlignment =
  | 'center'
  | 'justify'
  | 'left'
  | 'right';

export type DocumentPageChromeEditorCommand =
  | { type: 'insertImage'; alt: string; source: string }
  | { type: 'redo' }
  | { type: 'setAlignment'; alignment: DocumentPageChromeAlignment }
  | { type: 'setColor'; color: string }
  | { type: 'setLink'; href: string | null }
  | { type: 'toggleBold' }
  | { type: 'toggleItalic' }
  | { type: 'toggleUnderline' }
  | { type: 'undo' };

export interface DocumentPageChromeEditorState {
  alignment: DocumentPageChromeAlignment;
  bold: boolean;
  canRedo: boolean;
  canUndo: boolean;
  color: string;
  italic: boolean;
  link: string | null;
  underline: boolean;
}

export interface DocumentPageChromeRichTextEditorProps {
  autoFocus?: boolean;
  className?: string;
  label: string;
  value: string;
  onChange: (html: string) => void;
  onEditorChange?: (editor: Editor | null) => void;
  onExit?: () => void;
  showToolbar?: boolean;
}

const MAX_PAGE_CHROME_IMAGE_BYTES = 4 * 1024 * 1024;
const DEFAULT_PAGE_CHROME_COLOR = '#4d5668';
const PAGE_CHROME_IMAGE_TYPES = new Set([
  'image/bmp',
  'image/gif',
  'image/jpeg',
  'image/png',
  'image/webp',
]);

export function DocumentPageChromeRichTextEditor({
  autoFocus = false,
  className = '',
  label,
  value,
  onChange,
  onEditorChange,
  onExit,
  showToolbar = true,
}: DocumentPageChromeRichTextEditorProps) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const onChangeRef = useRef(onChange);
  const onEditorChangeRef = useRef(onEditorChange);
  const onExitRef = useRef(onExit);
  const initialHtmlRef = useRef(sanitizeDocumentPageChromeHtml(value));
  const appliedHtmlRef = useRef(initialHtmlRef.current);
  const [, setEditorRevision] = useState(0);
  const officeDialog = useOfficeDialog();
  onChangeRef.current = onChange;
  onEditorChangeRef.current = onEditorChange;
  onExitRef.current = onExit;
  const extensions = useMemo(
    () => createDocumentPageChromeEditorExtensions(`输入${label}`),
    [label],
  );
  const editor = useEditor({
    extensions,
    content: initialHtmlRef.current,
    editorProps: {
      attributes: {
        'aria-label': label,
        'aria-multiline': 'true',
        class: 'work-document-page-chrome-content',
        'data-document-page-chrome-engine': 'tiptap',
        role: 'textbox',
        spellcheck: 'true',
      },
      handleKeyDown: (_view, event) => {
        if (event.key !== 'Escape' || !onExitRef.current) return false;
        event.preventDefault();
        onExitRef.current();
        return true;
      },
    },
    onTransaction: () => setEditorRevision((revision) => revision + 1),
    onUpdate: ({ editor: current }) => {
      const html = sanitizeDocumentPageChromeHtml(current.getHTML());
      if (html === appliedHtmlRef.current) return;
      appliedHtmlRef.current = html;
      onChangeRef.current(html);
    },
  });

  useEffect(() => {
    if (!editor) return;
    onEditorChangeRef.current?.(editor);
    return () => onEditorChangeRef.current?.(null);
  }, [editor]);

  useEffect(() => {
    if (!editor || editor.isDestroyed) return;
    const html = sanitizeDocumentPageChromeHtml(value);
    if (html === appliedHtmlRef.current) return;
    appliedHtmlRef.current = html;
    editor.commands.setContent(html, { emitUpdate: false });
  }, [editor, value]);

  useEffect(() => {
    if (!autoFocus || !editor || editor.isDestroyed) return;
    editor.commands.focus('end');
  }, [autoFocus, editor]);

  const state = editor ? documentPageChromeEditorState(editor) : null;
  const run = (command: DocumentPageChromeEditorCommand) => {
    if (editor) applyDocumentPageChromeEditorCommand(editor, command);
  };
  const editLink = async () => {
    if (!editor) return;
    if (documentPageChromeEditorState(editor).link) {
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'setLink',
        href: null,
      });
      return;
    }
    const href = await officeDialog.prompt({
      title: '链接地址',
      initialValue: 'https://',
      placeholder: 'https://',
      confirmLabel: '添加链接',
    });
    if (href === null) return;
    const normalized = normalizeDocumentPageChromeHref(href);
    if (!normalized) {
      await officeDialog.notice({
        title: '无法添加链接',
        description: '请输入 http、https、mailto 或文档内 # 锚点地址。',
      });
      return;
    }
    if (!editor.isDestroyed) {
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'setLink',
        href: normalized,
      });
    }
  };
  const insertImage = async (file: File | undefined) => {
    if (!file || !editor) return;
    const image = await loadDocumentPageChromeImage(file);
    if (!image.ok) {
      await officeDialog.notice({
        title: image.title,
        description: image.description,
      });
      return;
    }
    if (!editor.isDestroyed) {
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'insertImage',
        alt: image.alt,
        source: image.source,
      });
    }
  };

  return (
    <>
      <section
        className={`work-document-page-chrome-editor ${className}`.trim()}
        aria-label={label}
      >
        {showToolbar && (
          <>
            <div
              className="work-document-page-chrome-toolbar"
              role="toolbar"
              aria-label={`${label}格式`}
            >
              <PageChromeButton
                label={`${label}撤销`}
                disabled={!state?.canUndo}
                onClick={() => run({ type: 'undo' })}
              >
                <Undo2 size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}重做`}
                disabled={!state?.canRedo}
                onClick={() => run({ type: 'redo' })}
              >
                <Redo2 size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}加粗`}
                active={state?.bold}
                onClick={() => run({ type: 'toggleBold' })}
              >
                <Bold size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}斜体`}
                active={state?.italic}
                onClick={() => run({ type: 'toggleItalic' })}
              >
                <Italic size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}下划线`}
                active={state?.underline}
                onClick={() => run({ type: 'toggleUnderline' })}
              >
                <UnderlineIcon size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}左对齐`}
                active={state?.alignment === 'left'}
                onClick={() => run({ type: 'setAlignment', alignment: 'left' })}
              >
                <AlignLeft size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}居中`}
                active={state?.alignment === 'center'}
                onClick={() =>
                  run({ type: 'setAlignment', alignment: 'center' })
                }
              >
                <AlignCenter size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}右对齐`}
                active={state?.alignment === 'right'}
                onClick={() =>
                  run({ type: 'setAlignment', alignment: 'right' })
                }
              >
                <AlignRight size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={`${label}两端对齐`}
                active={state?.alignment === 'justify'}
                onClick={() =>
                  run({ type: 'setAlignment', alignment: 'justify' })
                }
              >
                <AlignJustify size={14} />
              </PageChromeButton>
              <PageChromeButton
                label={state?.link ? `${label}移除链接` : `${label}添加链接`}
                active={Boolean(state?.link)}
                onClick={() => void editLink()}
              >
                <Link2 size={14} />
              </PageChromeButton>
              <OfficeColorPicker
                compact
                className="work-document-page-chrome-color"
                ariaLabel={`${label}文字颜色`}
                value={pageChromePickerColor(state?.color)}
                onValueChange={(color) => run({ type: 'setColor', color })}
              />
              <PageChromeButton
                label={`${label}插入图片`}
                onClick={() => imageInputRef.current?.click()}
              >
                <ImageIcon size={14} />
              </PageChromeButton>
            </div>
            <OfficeFileInput
              ref={imageInputRef}
              accept="image/bmp,image/gif,image/jpeg,image/png,image/webp"
              aria-label={`${label}图片文件`}
              onChange={(event) => {
                const file = event.target.files?.[0];
                event.target.value = '';
                void insertImage(file);
              }}
            />
          </>
        )}
        {editor ? (
          <EditorContent editor={editor} />
        ) : (
          <div
            className="work-document-page-chrome-content"
            role="status"
            aria-label={`${label}正在准备`}
          />
        )}
      </section>
      {officeDialog.dialog}
    </>
  );
}

export function createDocumentPageChromeEditorExtensions(
  placeholder = '输入页眉或页脚',
): Extensions {
  return [
    StarterKit.configure({
      code: false,
      codeBlock: false,
      heading: false,
      horizontalRule: false,
      link: {
        autolink: false,
        defaultProtocol: 'https',
        openOnClick: false,
      },
      underline: false,
    }),
    TableKit.configure({
      table: {
        allowTableNodeSelection: true,
        resizable: false,
      },
    }),
    Image.configure({ allowBase64: true, inline: true }),
    TextStyle,
    Color,
    Underline,
    TextAlign.configure({ types: ['paragraph'] }),
    Placeholder.configure({ placeholder }),
  ];
}

export function documentPageChromeEditorState(
  editor: Editor,
): DocumentPageChromeEditorState {
  const textStyle = editor.getAttributes('textStyle') as Record<
    string,
    unknown
  >;
  const link = editor.getAttributes('link') as Record<string, unknown>;
  return {
    alignment: editor.isActive('paragraph', { textAlign: 'center' })
      ? 'center'
      : editor.isActive('paragraph', { textAlign: 'right' })
        ? 'right'
        : editor.isActive('paragraph', { textAlign: 'justify' })
          ? 'justify'
          : 'left',
    bold: editor.isActive('bold'),
    canRedo: editor.can().chain().redo().run(),
    canUndo: editor.can().chain().undo().run(),
    color:
      typeof textStyle.color === 'string'
        ? textStyle.color
        : DEFAULT_PAGE_CHROME_COLOR,
    italic: editor.isActive('italic'),
    link: typeof link.href === 'string' ? link.href : null,
    underline: editor.isActive('underline'),
  };
}

export function applyDocumentPageChromeEditorCommand(
  editor: Editor,
  command: DocumentPageChromeEditorCommand,
): boolean {
  if (editor.isDestroyed) return false;
  switch (command.type) {
    case 'insertImage':
      if (!isDocumentPageChromeImageSource(command.source)) return false;
      return editor
        .chain()
        .focus()
        .setImage({
          src: command.source,
          alt: command.alt.trim() || 'Image',
          title: command.alt.trim() || undefined,
        })
        .run();
    case 'redo':
      return editor.chain().focus().redo().run();
    case 'setAlignment':
      return editor.chain().focus().setTextAlign(command.alignment).run();
    case 'setColor':
      return editor.chain().focus().setColor(command.color).run();
    case 'setLink': {
      const chain = editor.chain().focus().extendMarkRange('link');
      if (command.href === null) return chain.unsetLink().run();
      const href = normalizeDocumentPageChromeHref(command.href);
      return href ? chain.setLink({ href }).run() : false;
    }
    case 'toggleBold':
      return editor.chain().focus().toggleBold().run();
    case 'toggleItalic':
      return editor.chain().focus().toggleItalic().run();
    case 'toggleUnderline':
      return editor.chain().focus().toggleUnderline().run();
    case 'undo':
      return editor.chain().focus().undo().run();
  }
}

export function normalizeDocumentPageChromeHref(value: string): string | null {
  const href = value.trim();
  return /^(?:https?:|mailto:|#)/i.test(href) ? href : null;
}

export type DocumentPageChromeImageLoadResult =
  | { ok: true; alt: string; source: string }
  | { ok: false; title: string; description: string };

export async function loadDocumentPageChromeImage(
  file: File,
): Promise<DocumentPageChromeImageLoadResult> {
  if (
    file.size > MAX_PAGE_CHROME_IMAGE_BYTES ||
    !PAGE_CHROME_IMAGE_TYPES.has(file.type)
  ) {
    return {
      ok: false,
      title: '无法插入图片',
      description: '请选择不超过 4 MB 的 PNG、JPEG、GIF、WebP 或 BMP 图片。',
    };
  }
  try {
    return {
      ok: true,
      alt: file.name,
      source: await fileToDataUrl(file),
    };
  } catch {
    return {
      ok: false,
      title: '图片读取失败',
      description: '请重新选择本机图片文件。',
    };
  }
}

function isDocumentPageChromeImageSource(source: string): boolean {
  return /^(?:https?:|blob:|data:image\/)/i.test(source.trim());
}

function pageChromePickerColor(color: string | undefined): string {
  return color && /^#[\da-f]{6}$/i.test(color)
    ? color
    : DEFAULT_PAGE_CHROME_COLOR;
}

function PageChromeButton({
  label,
  active,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      className={active ? 'active' : undefined}
      aria-label={label}
      aria-pressed={active === undefined ? undefined : active}
      title={label}
      disabled={disabled}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(String(reader.result)));
    reader.addEventListener('error', () =>
      reject(reader.error ?? new Error('Image could not be read')),
    );
    reader.readAsDataURL(file);
  });
}
