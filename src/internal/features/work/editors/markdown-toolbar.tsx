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
import { type ButtonHTMLAttributes, type ReactNode, useState } from 'react';
import { OfficeSelect } from './office-controls';
import {
  MarkdownInsertDialog,
  type MarkdownInsertDialogRequest,
  type MarkdownInsertDialogResult,
} from './markdown-insert-dialog';
import type {
  MarkdownSourceCommand,
  MarkdownSourceSelectionState,
} from './markdown-source-commands';
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
  sourceEditing,
  canSourceRedo,
  canSourceUndo,
  viewMode,
  getSourceSelection,
  onSourceCommand,
  onSourceRedo,
  onSourceReplace,
  onSourceUndo,
  onViewModeChange,
}: {
  editor: Editor;
  fileActions?: readonly WorkOfficeFileAction[];
  sourceEditing: boolean;
  canSourceRedo: boolean;
  canSourceUndo: boolean;
  viewMode: MarkdownViewMode;
  getSourceSelection: () => MarkdownSourceSelectionState | null;
  onSourceCommand: (command: MarkdownSourceCommand) => boolean;
  onSourceRedo: () => boolean;
  onSourceReplace: (
    replacement: string,
    selectedRange?: { start: number; end: number },
  ) => boolean;
  onSourceUndo: () => boolean;
  onViewModeChange: (mode: MarkdownViewMode) => void;
}) {
  const [activeTab, setActiveTab] = useState<MarkdownRibbonTab>('home');
  const [insertDialog, setInsertDialog] =
    useState<MarkdownInsertDialogRequest | null>(null);
  const [insertDialogSurface, setInsertDialogSurface] = useState<
    'source' | 'visual'
  >('visual');
  const canUndo = sourceEditing
    ? canSourceUndo
    : canRunVisualEditorCommand(editor, () =>
        editor.can().chain().focus().undo().run(),
      );
  const canRedo = sourceEditing
    ? canSourceRedo
    : canRunVisualEditorCommand(editor, () =>
        editor.can().chain().focus().redo().run(),
      );
  const runCommand = (
    sourceCommand: MarkdownSourceCommand,
    visualCommand: () => void,
  ) => {
    if (sourceEditing && onSourceCommand(sourceCommand)) return;
    visualCommand();
  };
  const toggleLink = () => {
    if (sourceEditing) {
      setInsertDialogSurface('source');
      setInsertDialog({
        kind: 'link',
        label: getSourceSelection()?.text ?? '',
        source: 'https://',
      });
      return;
    }
    if (editor.isActive('link')) {
      editor.chain().focus().unsetLink().run();
      return;
    }
    const { from, to } = editor.state.selection;
    setInsertDialog({
      kind: 'link',
      label: editor.state.selection.empty
        ? ''
        : editor.state.doc.textBetween(from, to, ' '),
      source: 'https://',
    });
    setInsertDialogSurface('visual');
  };
  const openImageDialog = () => {
    setInsertDialogSurface(sourceEditing ? 'source' : 'visual');
    setInsertDialog({
      kind: 'image',
      altText: '',
      source: 'https://',
    });
  };
  const commitInsert = (result: MarkdownInsertDialogResult) => {
    setInsertDialog(null);
    if (insertDialogSurface === 'source') {
      if (result.kind === 'image') {
        const replacement = `![${result.altText}](${result.source})`;
        onSourceReplace(replacement, {
          start: 2,
          end: 2 + result.altText.length,
        });
      } else {
        const replacement = `[${result.label}](${result.source})`;
        onSourceReplace(replacement, {
          start: 1,
          end: 1 + result.label.length,
        });
      }
      return;
    }
    if (result.kind === 'image') {
      editor
        .chain()
        .focus()
        .setImage({
          src: result.source,
          alt: result.altText || undefined,
        })
        .run();
      return;
    }
    const { from, to } = editor.state.selection;
    const selectedText = editor.state.selection.empty
      ? ''
      : editor.state.doc.textBetween(from, to, ' ');
    if (selectedText && selectedText === result.label) {
      editor.chain().focus().setLink({ href: result.source }).run();
      return;
    }
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'text',
        text: result.label,
        marks: [{ type: 'link', attrs: { href: result.source } }],
      })
      .run();
  };

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
                  disabled={!canUndo}
                  onClick={() => {
                    if (sourceEditing) onSourceUndo();
                    else editor.chain().focus().undo().run();
                  }}
                >
                  <Undo2 size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="重做"
                  shortcut="Cmd/Ctrl+Shift+Z"
                  disabled={!canRedo}
                  onClick={() => {
                    if (sourceEditing) onSourceRedo();
                    else editor.chain().focus().redo().run();
                  }}
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
                      runCommand('paragraph', () => {
                        editor.chain().focus().setParagraph().run();
                      });
                    } else {
                      const level = Number(value.slice(1)) as 1 | 2 | 3;
                      runCommand(`heading-${level}`, () => {
                        editor.chain().focus().toggleHeading({ level }).run();
                      });
                    }
                  }}
                />
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="文字">
                <MarkdownToolbarButton
                  label="加粗"
                  shortcut="Cmd/Ctrl+B"
                  active={editor.isActive('bold')}
                  onClick={() =>
                    runCommand('bold', () => {
                      editor.chain().focus().toggleBold().run();
                    })
                  }
                >
                  <Bold size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="斜体"
                  shortcut="Cmd/Ctrl+I"
                  active={editor.isActive('italic')}
                  onClick={() =>
                    runCommand('italic', () => {
                      editor.chain().focus().toggleItalic().run();
                    })
                  }
                >
                  <Italic size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="删除线"
                  active={editor.isActive('strike')}
                  onClick={() =>
                    runCommand('strike', () => {
                      editor.chain().focus().toggleStrike().run();
                    })
                  }
                >
                  <Strikethrough size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="行内代码"
                  active={editor.isActive('code')}
                  onClick={() =>
                    runCommand('code', () => {
                      editor.chain().focus().toggleCode().run();
                    })
                  }
                >
                  <Code2 size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="段落">
                <MarkdownToolbarButton
                  label="项目列表"
                  active={editor.isActive('bulletList')}
                  onClick={() =>
                    runCommand('bullet-list', () => {
                      editor.chain().focus().toggleBulletList().run();
                    })
                  }
                >
                  <List size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="编号列表"
                  active={editor.isActive('orderedList')}
                  onClick={() =>
                    runCommand('ordered-list', () => {
                      editor.chain().focus().toggleOrderedList().run();
                    })
                  }
                >
                  <ListOrdered size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="任务列表"
                  active={editor.isActive('taskList')}
                  onClick={() =>
                    runCommand('task-list', () => {
                      editor.chain().focus().toggleTaskList().run();
                    })
                  }
                >
                  <ListChecks size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="引用"
                  active={editor.isActive('blockquote')}
                  onClick={() =>
                    runCommand('blockquote', () => {
                      editor.chain().focus().toggleBlockquote().run();
                    })
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
                  onClick={toggleLink}
                >
                  <Link2 size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="插入图片"
                  displayLabel
                  onClick={openImageDialog}
                >
                  <ImageIcon size={19} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="内容">
                <MarkdownToolbarButton
                  label="代码块"
                  displayLabel
                  active={editor.isActive('codeBlock')}
                  onClick={() =>
                    runCommand('code-block', () => {
                      editor.chain().focus().toggleCodeBlock().run();
                    })
                  }
                >
                  <Code2 size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="分隔线"
                  displayLabel
                  onClick={() =>
                    runCommand('horizontal-rule', () => {
                      editor.chain().focus().setHorizontalRule().run();
                    })
                  }
                >
                  <Minus size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="表格"
                  displayLabel
                  onClick={() =>
                    runCommand('table', () => {
                      editor
                        .chain()
                        .focus()
                        .insertTable({
                          rows: 3,
                          cols: 3,
                          withHeaderRow: true,
                        })
                        .run();
                    })
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
      {insertDialog && (
        <MarkdownInsertDialog
          request={insertDialog}
          restoreFocusTarget={() => editor.view.dom}
          onClose={() => setInsertDialog(null)}
          onSubmit={commitInsert}
        />
      )}
    </>
  );
}

function canRunVisualEditorCommand(
  editor: Editor,
  command: () => boolean,
): boolean {
  if (editor.isDestroyed || !editor.view) return false;
  return command();
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
