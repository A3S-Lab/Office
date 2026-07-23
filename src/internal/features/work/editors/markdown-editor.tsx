import Image from '@tiptap/extension-image';
import Placeholder from '@tiptap/extension-placeholder';
import { TableKit } from '@tiptap/extension-table';
import { Markdown } from '@tiptap/markdown';
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import {
  Bold,
  Code2,
  Columns2,
  Image as ImageIcon,
  Italic,
  Link2,
  List,
  ListOrdered,
  Minus,
  Quote,
  Redo2,
  PencilLine,
  Strikethrough,
  Table2,
  Undo2,
} from 'lucide-react';
import {
  type ButtonHTMLAttributes,
  type CSSProperties,
  type ReactNode,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import type { WorkMarkdownContent } from '../work-types';
import { OfficeSelect, useOfficeDialog } from './office-controls';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

export interface MarkdownEditorProps {
  content: WorkMarkdownContent;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  onChange: (content: WorkMarkdownContent) => void;
}

type MarkdownViewMode = 'visual' | 'source' | 'split';
type MarkdownRibbonTab = 'home' | 'insert' | 'view';

const MARKDOWN_MIN_ZOOM = 60;
const MARKDOWN_MAX_ZOOM = 180;
const markdownRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'view', label: '视图' },
] as const;

export function MarkdownEditor({
  content,
  preview,
  saveStatus,
  fileActions,
  onChange,
}: MarkdownEditorProps) {
  const contentRef = useRef(content);
  const receivedContentRef = useRef(content);
  const appliedMarkdownRef = useRef(content.markdown);
  const initialMarkdownRef = useRef(content.markdown);
  const [viewMode, setViewMode] = useState<MarkdownViewMode>('visual');
  const [zoom, setZoom] = useState(100);
  const [, setSelectionVersion] = useState(0);
  contentRef.current = content;

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
      const next = { ...contentRef.current, markdown };
      appliedMarkdownRef.current = markdown;
      contentRef.current = next;
      onChange(next);
    },
    onSelectionUpdate: () => setSelectionVersion((value) => value + 1),
  });

  useEffect(() => {
    if (!editor) return;
    editor.setEditable(!preview);
    editor.view.dom.setAttribute(
      'aria-label',
      preview ? 'Markdown 预览' : 'Markdown 编辑区',
    );
    editor.view.dom.setAttribute('aria-readonly', String(preview));
    editor.view.dom.setAttribute('role', preview ? 'document' : 'textbox');
    if (preview) editor.view.dom.removeAttribute('aria-multiline');
    else editor.view.dom.setAttribute('aria-multiline', 'true');
  }, [editor, preview]);

  useEffect(() => {
    if (!editor || receivedContentRef.current === content) return;
    receivedContentRef.current = content;
    if (appliedMarkdownRef.current === content.markdown) return;
    appliedMarkdownRef.current = content.markdown;
    editor.commands.setContent(content.markdown, {
      contentType: 'markdown',
      emitUpdate: false,
    });
  }, [content, editor]);

  const updateSource = useCallback(
    (markdown: string) => {
      if (!editor) return;
      const next = { ...contentRef.current, markdown };
      appliedMarkdownRef.current = markdown;
      contentRef.current = next;
      editor.commands.setContent(markdown, {
        contentType: 'markdown',
        emitUpdate: false,
      });
      onChange(next);
    },
    [editor, onChange],
  );

  if (!editor) {
    return <WorkEditorLoadingState title="正在准备 Markdown 编辑器" />;
  }

  const lineCount = content.markdown
    ? content.markdown.replace(/\r\n?/g, '\n').split('\n').length
    : 1;
  const characterCount = Array.from(content.markdown).length;
  const editorStyle = {
    '--work-markdown-zoom': zoom,
  } as CSSProperties;

  if (preview) {
    return (
      <section className="work-markdown-editor preview" style={editorStyle}>
        <WorkOfficePreviewBar
          ariaLabel="Markdown 预览工具"
          label="只读预览"
          detail={`${lineCount} 行`}
          fileActions={fileActions}
          className="work-markdown-ribbon"
        />
        <MarkdownWorkspace
          editor={editor}
          markdown={content.markdown}
          mode="visual"
          readOnly
          onSourceChange={updateSource}
        />
        <MarkdownStatus
          characterCount={characterCount}
          lineCount={lineCount}
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
        onViewModeChange={setViewMode}
      />
      <MarkdownWorkspace
        editor={editor}
        markdown={content.markdown}
        mode={viewMode}
        onSourceChange={updateSource}
      />
      <MarkdownStatus
        characterCount={characterCount}
        lineCount={lineCount}
        saveStatus={saveStatus}
        zoom={zoom}
        onZoomChange={setZoom}
      />
    </section>
  );
}

function MarkdownWorkspace({
  editor,
  markdown,
  mode,
  readOnly = false,
  onSourceChange,
}: {
  editor: NonNullable<ReturnType<typeof useEditor>>;
  markdown: string;
  mode: MarkdownViewMode;
  readOnly?: boolean;
  onSourceChange: (markdown: string) => void;
}) {
  const showSource = !readOnly && mode !== 'visual';
  const showVisual = readOnly || mode !== 'source';
  return (
    <div className={`work-markdown-workspace ${mode}`}>
      {showSource && (
        <section className="work-markdown-pane source">
          {mode === 'split' && (
            <header className="work-markdown-pane-label">源码</header>
          )}
          <textarea
            aria-label="Markdown 源码"
            value={markdown}
            spellCheck
            onChange={(event) => onSourceChange(event.target.value)}
          />
        </section>
      )}
      {showVisual && (
        <section className="work-markdown-pane visual">
          {mode === 'split' && (
            <header className="work-markdown-pane-label">编辑结果</header>
          )}
          <div className="work-markdown-canvas">
            <EditorContent editor={editor} />
          </div>
        </section>
      )}
    </div>
  );
}

function MarkdownToolbar({
  editor,
  fileActions,
  viewMode,
  onViewModeChange,
}: {
  editor: NonNullable<ReturnType<typeof useEditor>>;
  fileActions?: readonly WorkOfficeFileAction[];
  viewMode: MarkdownViewMode;
  onViewModeChange: (mode: MarkdownViewMode) => void;
}) {
  const [activeTab, setActiveTab] = useState<MarkdownRibbonTab>('home');
  const officeDialog = useOfficeDialog();
  const toggleLink = useCallback(async () => {
    if (editor.isActive('link')) {
      editor.chain().focus().unsetLink().run();
      return;
    }
    const href = await officeDialog.prompt({
      title: '链接地址',
      initialValue: editor.getAttributes('link').href ?? 'https://',
      placeholder: 'https://',
      confirmLabel: '添加链接',
    });
    if (href?.trim())
      editor.chain().focus().setLink({ href: href.trim() }).run();
  }, [editor, officeDialog.prompt]);
  const insertImage = useCallback(async () => {
    const source = await officeDialog.prompt({
      title: '图片地址',
      placeholder: 'https://',
      confirmLabel: '插入图片',
    });
    if (source?.trim())
      editor.chain().focus().setImage({ src: source.trim() }).run();
  }, [editor, officeDialog.prompt]);

  return (
    <>
      <WorkOfficeRibbon
        ariaLabel="Markdown 功能区"
        tabs={markdownRibbonTabs}
        defaultTab="home"
        activeTab={activeTab}
        onTabChange={setActiveTab}
        fileActions={fileActions}
        className="work-markdown-ribbon"
        toolbarClassName="markdown-toolbar"
        panels={{
          home: (
            <>
              <WorkOfficeRibbonGroup label="撤销">
                <MarkdownToolbarButton
                  label="撤销"
                  shortcut="Cmd/Ctrl+Z"
                  disabled={!editor.can().chain().focus().undo().run()}
                  onClick={() => editor.chain().focus().undo().run()}
                >
                  <Undo2 size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="重做"
                  shortcut="Cmd/Ctrl+Shift+Z"
                  disabled={!editor.can().chain().focus().redo().run()}
                  onClick={() => editor.chain().focus().redo().run()}
                >
                  <Redo2 size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="样式">
                <OfficeSelect
                  ariaLabel="段落样式"
                  value={
                    editor.isActive('heading', { level: 1 })
                      ? 'h1'
                      : editor.isActive('heading', { level: 2 })
                        ? 'h2'
                        : editor.isActive('heading', { level: 3 })
                          ? 'h3'
                          : 'paragraph'
                  }
                  options={[
                    { value: 'paragraph', label: '正文' },
                    { value: 'h1', label: '标题 1' },
                    { value: 'h2', label: '标题 2' },
                    { value: 'h3', label: '标题 3' },
                  ]}
                  onValueChange={(value) => {
                    if (value === 'paragraph')
                      editor.chain().focus().setParagraph().run();
                    else
                      editor
                        .chain()
                        .focus()
                        .toggleHeading({
                          level: Number(value.slice(1)) as 1 | 2 | 3,
                        })
                        .run();
                  }}
                />
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="文字">
                <MarkdownToolbarButton
                  label="加粗"
                  shortcut="Cmd/Ctrl+B"
                  active={editor.isActive('bold')}
                  onClick={() => editor.chain().focus().toggleBold().run()}
                >
                  <Bold size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="斜体"
                  shortcut="Cmd/Ctrl+I"
                  active={editor.isActive('italic')}
                  onClick={() => editor.chain().focus().toggleItalic().run()}
                >
                  <Italic size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="删除线"
                  active={editor.isActive('strike')}
                  onClick={() => editor.chain().focus().toggleStrike().run()}
                >
                  <Strikethrough size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="行内代码"
                  active={editor.isActive('code')}
                  onClick={() => editor.chain().focus().toggleCode().run()}
                >
                  <Code2 size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="段落">
                <MarkdownToolbarButton
                  label="项目列表"
                  active={editor.isActive('bulletList')}
                  onClick={() =>
                    editor.chain().focus().toggleBulletList().run()
                  }
                >
                  <List size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="编号列表"
                  active={editor.isActive('orderedList')}
                  onClick={() =>
                    editor.chain().focus().toggleOrderedList().run()
                  }
                >
                  <ListOrdered size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="引用"
                  active={editor.isActive('blockquote')}
                  onClick={() =>
                    editor.chain().focus().toggleBlockquote().run()
                  }
                >
                  <Quote size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
            </>
          ),
          insert: (
            <>
              <WorkOfficeRibbonGroup label="链接">
                <MarkdownToolbarButton
                  label={editor.isActive('link') ? '移除链接' : '添加链接'}
                  displayLabel
                  active={editor.isActive('link')}
                  onClick={() => void toggleLink()}
                >
                  <Link2 size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="插入图片"
                  displayLabel
                  onClick={() => void insertImage()}
                >
                  <ImageIcon size={19} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="内容">
                <MarkdownToolbarButton
                  label="代码块"
                  displayLabel
                  active={editor.isActive('codeBlock')}
                  onClick={() => editor.chain().focus().toggleCodeBlock().run()}
                >
                  <Code2 size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="分隔线"
                  displayLabel
                  onClick={() =>
                    editor.chain().focus().setHorizontalRule().run()
                  }
                >
                  <Minus size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="表格"
                  displayLabel
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertTable({
                        rows: 3,
                        cols: 3,
                        withHeaderRow: true,
                      })
                      .run()
                  }
                >
                  <Table2 size={19} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
            </>
          ),
          view: (
            <WorkOfficeRibbonGroup label="编辑方式">
              <MarkdownToolbarButton
                label="编辑"
                displayLabel
                active={viewMode === 'visual'}
                onClick={() => onViewModeChange('visual')}
              >
                <PencilLine size={19} />
              </MarkdownToolbarButton>
              <MarkdownToolbarButton
                label="源码"
                displayLabel
                active={viewMode === 'source'}
                onClick={() => onViewModeChange('source')}
              >
                <Code2 size={19} />
              </MarkdownToolbarButton>
              <MarkdownToolbarButton
                label="分屏"
                displayLabel
                active={viewMode === 'split'}
                onClick={() => onViewModeChange('split')}
              >
                <Columns2 size={19} />
              </MarkdownToolbarButton>
            </WorkOfficeRibbonGroup>
          ),
        }}
      />
      {officeDialog.dialog}
    </>
  );
}

function MarkdownToolbarButton({
  label,
  shortcut,
  active = false,
  displayLabel = false,
  children,
  ...props
}: {
  label: string;
  shortcut?: string;
  active?: boolean;
  displayLabel?: boolean;
  children: ReactNode;
} & Omit<
  ButtonHTMLAttributes<HTMLButtonElement>,
  'aria-label' | 'children' | 'title' | 'type'
>) {
  return (
    <WorkOfficeRibbonButton
      {...props}
      label={label}
      title={shortcut ? `${label}（${shortcut}）` : label}
      active={active}
      displayLabel={displayLabel}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

function MarkdownStatus({
  characterCount,
  lineCount,
  saveStatus,
  zoom,
  onZoomChange,
}: {
  characterCount: number;
  lineCount: number;
  saveStatus?: string;
  zoom: number;
  onZoomChange: (zoom: number) => void;
}) {
  return (
    <WorkOfficeStatusBar
      className="work-markdown-status"
      controls={
        <WorkOfficeZoomControls
          zoom={zoom}
          minimum={MARKDOWN_MIN_ZOOM}
          maximum={MARKDOWN_MAX_ZOOM}
          decreaseLabel="缩小内容"
          increaseLabel="放大内容"
          outputLabel="Markdown 缩放比例"
          sliderLabel="调整 Markdown 缩放比例"
          onChange={onZoomChange}
        />
      }
    >
      <output>{lineCount} 行</output>
      <output>{characterCount} 字符</output>
      {saveStatus && (
        <span className="work-office-save-status">{saveStatus}</span>
      )}
    </WorkOfficeStatusBar>
  );
}
