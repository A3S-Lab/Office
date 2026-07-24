import { TaskItem, TaskList } from '@tiptap/extension-list';
import Image from '@tiptap/extension-image';
import Placeholder from '@tiptap/extension-placeholder';
import { TableKit } from '@tiptap/extension-table';
import { Markdown } from '@tiptap/markdown';
import { useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import {
  type CSSProperties,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import type { WorkMarkdownContent } from '../work-types';
import { MarkdownStatus } from './markdown-status';
import { MarkdownToolbar } from './markdown-toolbar';
import { type MarkdownViewMode, MarkdownWorkspace } from './markdown-workspace';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
} from './work-office-chrome';

export interface MarkdownEditorProps {
  content: WorkMarkdownContent;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  onChange: (content: WorkMarkdownContent) => void;
}

const MARKDOWN_PREVIEW_SYNC_DELAY = 160;

export function MarkdownEditor({
  content,
  preview,
  saveStatus,
  fileActions,
  onChange,
}: MarkdownEditorProps) {
  const contentRef = useRef(content);
  const onChangeRef = useRef(onChange);
  const receivedContentRef = useRef(content);
  const appliedMarkdownRef = useRef(content.markdown);
  const emittedMarkdownRef = useRef<string | null>(null);
  const sourceMarkdownRef = useRef(content.markdown);
  const initialMarkdownRef = useRef(content.markdown);
  const previewTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [sourceMarkdown, setSourceMarkdown] = useState(content.markdown);
  const [viewMode, setViewMode] = useState<MarkdownViewMode>('split');
  const [zoom, setZoom] = useState(100);
  const [, setSelectionVersion] = useState(0);
  contentRef.current = content;
  onChangeRef.current = onChange;

  const cancelPreviewSync = useCallback(() => {
    if (previewTimerRef.current === null) return;
    clearTimeout(previewTimerRef.current);
    previewTimerRef.current = null;
  }, []);

  const extensions = useMemo(
    () => [
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
          checkboxLabel: (node, checked) => {
            const label = node.textContent.trim() || '任务';
            return `${checked ? '已完成' : '未完成'}：${label}`;
          },
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
    ],
    [],
  );
  const editorProps = useMemo(
    () => ({
      attributes: {
        'aria-label': 'Markdown 编辑区',
        'aria-multiline': 'true',
        role: 'textbox',
        spellcheck: 'true',
      },
    }),
    [],
  );
  const editor = useEditor({
    extensions,
    content: initialMarkdownRef.current,
    contentType: 'markdown',
    editable: !preview,
    editorProps,
    onUpdate: ({ editor: current }) => {
      const markdown = current.getMarkdown();
      if (markdown === appliedMarkdownRef.current) return;
      cancelPreviewSync();
      const next = { ...contentRef.current, markdown };
      appliedMarkdownRef.current = markdown;
      emittedMarkdownRef.current = markdown;
      sourceMarkdownRef.current = markdown;
      contentRef.current = next;
      setSourceMarkdown(markdown);
      onChangeRef.current(next);
    },
    onSelectionUpdate: () => setSelectionVersion((value) => value + 1),
  });

  const applyMarkdownToEditor = useCallback(
    (markdown: string) => {
      cancelPreviewSync();
      if (
        !editor ||
        editor.isDestroyed ||
        appliedMarkdownRef.current === markdown
      ) {
        return;
      }
      appliedMarkdownRef.current = markdown;
      editor.commands.setContent(markdown, {
        contentType: 'markdown',
        emitUpdate: false,
      });
    },
    [cancelPreviewSync, editor],
  );

  const queueMarkdownPreview = useCallback(
    (markdown: string, immediate = false) => {
      cancelPreviewSync();
      if (
        !editor ||
        editor.isDestroyed ||
        appliedMarkdownRef.current === markdown
      ) {
        return;
      }
      if (immediate) {
        applyMarkdownToEditor(markdown);
        return;
      }
      previewTimerRef.current = setTimeout(() => {
        previewTimerRef.current = null;
        applyMarkdownToEditor(markdown);
      }, MARKDOWN_PREVIEW_SYNC_DELAY);
    },
    [applyMarkdownToEditor, cancelPreviewSync, editor],
  );

  useEffect(() => cancelPreviewSync, [cancelPreviewSync]);

  useEffect(() => {
    if (!editor) return;
    const applyViewState = () => {
      if (editor.isDestroyed) return;
      editor.setEditable(!preview);
      editor.view.dom.setAttribute(
        'aria-label',
        preview ? 'Markdown 预览' : 'Markdown 编辑区',
      );
      editor.view.dom.setAttribute('aria-readonly', String(preview));
      editor.view.dom.setAttribute('role', preview ? 'document' : 'textbox');
      if (preview) editor.view.dom.removeAttribute('aria-multiline');
      else editor.view.dom.setAttribute('aria-multiline', 'true');
    };
    applyViewState();
    editor.on('mount', applyViewState);
    return () => {
      editor.off('mount', applyViewState);
    };
  }, [editor, preview]);

  useEffect(() => {
    if (!editor || receivedContentRef.current === content) return;
    receivedContentRef.current = content;
    const markdown = content.markdown;
    if (sourceMarkdownRef.current !== markdown) {
      sourceMarkdownRef.current = markdown;
      setSourceMarkdown(markdown);
    }
    if (emittedMarkdownRef.current === markdown) {
      emittedMarkdownRef.current = null;
      return;
    }
    emittedMarkdownRef.current = null;
    if (preview || viewMode !== 'source') {
      queueMarkdownPreview(markdown, true);
    }
  }, [content, editor, preview, queueMarkdownPreview, viewMode]);

  const updateSource = useCallback(
    (markdown: string) => {
      if (markdown === sourceMarkdownRef.current) return;
      sourceMarkdownRef.current = markdown;
      setSourceMarkdown(markdown);
      const next = { ...contentRef.current, markdown };
      contentRef.current = next;
      emittedMarkdownRef.current = markdown;
      onChangeRef.current(next);
      if (viewMode !== 'source') {
        queueMarkdownPreview(markdown);
      }
    },
    [queueMarkdownPreview, viewMode],
  );

  const changeViewMode = useCallback(
    (mode: MarkdownViewMode) => {
      if (mode !== 'source') {
        applyMarkdownToEditor(sourceMarkdownRef.current);
      }
      setViewMode(mode);
    },
    [applyMarkdownToEditor],
  );

  const deferredMarkdown = useDeferredValue(sourceMarkdown);
  const metrics = useMemo(
    () => markdownMetrics(deferredMarkdown),
    [deferredMarkdown],
  );
  const editorStyle = {
    '--work-markdown-zoom': zoom,
  } as CSSProperties;

  if (!editor) {
    return <WorkEditorLoadingState title="正在准备 Markdown 编辑器" />;
  }

  if (preview) {
    return (
      <section className="work-markdown-editor preview" style={editorStyle}>
        <WorkOfficePreviewBar
          ariaLabel="Markdown 预览工具"
          label="只读预览"
          detail={`${metrics.lineCount} 行`}
          fileActions={fileActions}
          className="work-markdown-ribbon"
        />
        <MarkdownWorkspace
          editor={editor}
          markdown={sourceMarkdown}
          mode="visual"
          readOnly
          onSourceChange={updateSource}
          onVisualIntent={() =>
            applyMarkdownToEditor(sourceMarkdownRef.current)
          }
        />
        <MarkdownStatus
          characterCount={metrics.characterCount}
          lineCount={metrics.lineCount}
          saveStatus={saveStatus}
          zoom={zoom}
          onZoomChange={setZoom}
        />
      </section>
    );
  }

  return (
    <section className="work-markdown-editor" style={editorStyle}>
      <MarkdownToolbar
        editor={editor}
        fileActions={fileActions}
        viewMode={viewMode}
        onViewModeChange={changeViewMode}
      />
      <MarkdownWorkspace
        editor={editor}
        markdown={sourceMarkdown}
        mode={viewMode}
        onSourceChange={updateSource}
        onVisualIntent={() => applyMarkdownToEditor(sourceMarkdownRef.current)}
      />
      <MarkdownStatus
        characterCount={metrics.characterCount}
        lineCount={metrics.lineCount}
        saveStatus={saveStatus}
        zoom={zoom}
        onZoomChange={setZoom}
      />
    </section>
  );
}

function markdownMetrics(markdown: string): {
  lineCount: number;
  characterCount: number;
} {
  return {
    lineCount: markdown
      ? markdown.replace(/\r\n?/g, '\n').split('\n').length
      : 1,
    characterCount: Array.from(markdown).length,
  };
}
