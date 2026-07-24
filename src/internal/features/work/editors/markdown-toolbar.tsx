import type { Editor } from '@tiptap/core';
import {
  Bold,
  Code2,
  Columns2,
  Image as ImageIcon,
  Italic,
  Link2,
  List,
  ListChecks,
  ListOrdered,
  Minus,
  PencilLine,
  Quote,
  Redo2,
  Strikethrough,
  Table2,
  Undo2,
} from 'lucide-react';
import {
  type ButtonHTMLAttributes,
  type ReactNode,
  useCallback,
  useState,
} from 'react';
import { OfficeSelect, useOfficeDialog } from './office-controls';
import type { MarkdownViewMode } from './markdown-workspace';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

type MarkdownRibbonTab = 'home' | 'insert' | 'view';

const markdownRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'view', label: '视图' },
] as const;

export function MarkdownToolbar({
  editor,
  fileActions,
  viewMode,
  onViewModeChange,
}: {
  editor: Editor;
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
    if (href?.trim()) {
      editor.chain().focus().setLink({ href: href.trim() }).run();
    }
  }, [editor, officeDialog.prompt]);
  const insertImage = useCallback(async () => {
    const source = await officeDialog.prompt({
      title: '图片地址',
      placeholder: 'https://',
      confirmLabel: '插入图片',
    });
    if (source?.trim()) {
      editor.chain().focus().setImage({ src: source.trim() }).run();
    }
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
                    if (value === 'paragraph') {
                      editor.chain().focus().setParagraph().run();
                    } else {
                      editor
                        .chain()
                        .focus()
                        .toggleHeading({
                          level: Number(value.slice(1)) as 1 | 2 | 3,
                        })
                        .run();
                    }
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
                  label="任务列表"
                  active={editor.isActive('taskList')}
                  onClick={() => editor.chain().focus().toggleTaskList().run()}
                >
                  <ListChecks size={16} />
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
