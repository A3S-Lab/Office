import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
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
  Unlink,
} from 'lucide-react';
import {
  type ButtonHTMLAttributes,
  type ReactNode,
  useEffect,
  useRef,
  useState,
} from 'react';
import {
  MarkdownInsertDialog,
  type MarkdownInsertDialogRequest,
  type MarkdownInsertDialogResult,
} from './markdown-insert-dialog';
import type {
  MarkdownSourceCommand,
  MarkdownSourceSelectionState,
} from './markdown-source-commands';
import {
  createMarkdownImageSourceInsert,
  createMarkdownLinkSourceInsert,
  createMarkdownSourceLinkRemoval,
  findMarkdownSourceLink,
  isMarkdownSourceCommandActive,
} from './markdown-source-commands';
import type { MarkdownViewMode } from './markdown-workspace';
import { OfficeSelect } from './office-controls';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

type MarkdownRibbonTab = 'home' | 'insert' | 'view';

type MarkdownInsertDialogState =
  | {
      request: MarkdownInsertDialogRequest;
      surface: 'source';
      target: MarkdownSourceSelectionState;
    }
  | {
      request: MarkdownInsertDialogRequest;
      surface: 'visual';
      target: {
        document: ProseMirrorNode;
        from: number;
        text: string;
        to: number;
      };
    };

const markdownRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'view', label: '视图' },
] as const;

interface MarkdownToolbarShortcut {
  ariaKeyShortcuts: string;
  label: string;
}

const markdownToolbarShortcuts = {
  bold: {
    ariaKeyShortcuts: 'Control+B Meta+B',
    label: 'Cmd/Ctrl+B',
  },
  italic: {
    ariaKeyShortcuts: 'Control+I Meta+I',
    label: 'Cmd/Ctrl+I',
  },
  link: {
    ariaKeyShortcuts: 'Control+K Meta+K',
    label: 'Cmd/Ctrl+K',
  },
  redo: {
    ariaKeyShortcuts: 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
    label: 'Cmd/Ctrl+Shift+Z 或 Cmd/Ctrl+Y',
  },
  undo: {
    ariaKeyShortcuts: 'Control+Z Meta+Z',
    label: 'Cmd/Ctrl+Z',
  },
} satisfies Record<string, MarkdownToolbarShortcut>;

export function MarkdownToolbar({
  editor,
  fileActions,
  collaborative,
  sourceEditing,
  canSourceRedo,
  canSourceUndo,
  viewMode,
  getSourceFocusTarget,
  getSourceSelection,
  onSourceCommand,
  onSourceRedo,
  onSourceReplace,
  onSourceUndo,
  onViewModeChange,
}: {
  editor: Editor;
  fileActions?: readonly WorkOfficeFileAction[];
  collaborative: boolean;
  sourceEditing: boolean;
  canSourceRedo: boolean;
  canSourceUndo: boolean;
  viewMode: MarkdownViewMode;
  getSourceFocusTarget: () => HTMLElement | null;
  getSourceSelection: () => MarkdownSourceSelectionState | null;
  onSourceCommand: (command: MarkdownSourceCommand) => boolean;
  onSourceRedo: () => boolean;
  onSourceReplace: (
    replacement: string,
    selectedRange?: { start: number; end: number },
    target?: MarkdownSourceSelectionState,
  ) => boolean;
  onSourceUndo: () => boolean;
  onViewModeChange: (mode: MarkdownViewMode) => void;
}) {
  const [activeTab, setActiveTab] = useState<MarkdownRibbonTab>('home');
  const [insertDialog, setInsertDialog] =
    useState<MarkdownInsertDialogState | null>(null);
  const openLinkDialogRef = useRef<() => void>(() => undefined);
  const sourceSelection = sourceEditing ? getSourceSelection() : null;
  const sourceLink = sourceSelection
    ? findMarkdownSourceLink(
        sourceSelection.markdown,
        sourceSelection.selection,
      )
    : null;
  const commandIsActive = (command: MarkdownSourceCommand) =>
    sourceSelection
      ? isMarkdownSourceCommandActive(
          sourceSelection.markdown,
          sourceSelection.selection,
          command,
        )
      : false;
  const paragraphStyle = markdownParagraphStyleValue(
    sourceEditing,
    commandIsActive,
    editor,
  );
  const linkActive = sourceEditing
    ? Boolean(sourceLink)
    : editor.isActive('link');
  const usesSharedHistory = sourceEditing || collaborative;
  const canUndo = usesSharedHistory
    ? canSourceUndo
    : canRunVisualEditorCommand(editor, () =>
        editor.can().chain().focus().undo().run(),
      );
  const canRedo = usesSharedHistory
    ? canSourceRedo
    : canRunVisualEditorCommand(editor, () =>
        editor.can().chain().focus().redo().run(),
      );
  const runCommand = (
    sourceCommand: MarkdownSourceCommand,
    visualCommand: () => void,
  ) => {
    if (sourceEditing) {
      onSourceCommand(sourceCommand);
      return;
    }
    visualCommand();
  };
  const openLinkDialog = () => {
    if (sourceEditing) {
      const current = getSourceSelection();
      if (!current) return;
      const link = findMarkdownSourceLink(current.markdown, current.selection);
      setInsertDialog({
        surface: 'source',
        target: link
          ? {
              markdown: current.markdown,
              selection: {
                start: link.range.start,
                end: link.range.end,
                direction: 'none',
              },
              text: current.markdown.slice(link.range.start, link.range.end),
            }
          : current,
        request: {
          kind: 'link',
          action: link ? 'edit' : 'insert',
          label: link?.label ?? current.text,
          source: link?.source ?? 'https://',
        },
      });
      return;
    }
    const editingLink = editor.isActive('link');
    if (editingLink) editor.commands.extendMarkRange('link');
    const { from, to } = editor.state.selection;
    setInsertDialog({
      surface: 'visual',
      target: {
        document: editor.state.doc,
        from,
        text: editor.state.doc.textBetween(from, to, ' '),
        to,
      },
      request: {
        kind: 'link',
        action: editingLink ? 'edit' : 'insert',
        label: editor.state.selection.empty
          ? ''
          : editor.state.doc.textBetween(from, to, ' '),
        source: editingLink
          ? String(editor.getAttributes('link').href ?? '')
          : 'https://',
      },
    });
  };

  useEffect(() => {
    openLinkDialogRef.current = openLinkDialog;
  });

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (
        event.repeat ||
        event.altKey ||
        event.shiftKey ||
        !(event.metaKey || event.ctrlKey) ||
        event.key.toLocaleLowerCase() !== 'k' ||
        isOfficeShortcutBlocked(event.target)
      ) {
        return;
      }
      const target = event.target;
      if (
        !(target instanceof Element) ||
        !target.closest('.work-markdown-editor')
      ) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      openLinkDialogRef.current();
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, []);

  const removeLink = () => {
    if (sourceEditing) {
      const current = getSourceSelection();
      if (!current) return;
      const link = findMarkdownSourceLink(current.markdown, current.selection);
      if (!link) return;
      const removal = createMarkdownSourceLinkRemoval(link);
      onSourceReplace(removal.replacement, removal.selectedRange, {
        markdown: current.markdown,
        selection: {
          start: link.range.start,
          end: link.range.end,
          direction: 'none',
        },
        text: current.markdown.slice(link.range.start, link.range.end),
      });
      return;
    }
    if (!editor.isActive('link')) return;
    // Keep ribbon focus on remove-link. chain().focus() schedules into the
    // editor and breaks L2 loops.
    editor.chain().extendMarkRange('link').unsetLink().run();
  };
  const openImageDialog = () => {
    if (sourceEditing) {
      const current = getSourceSelection();
      if (!current) return;
      setInsertDialog({
        surface: 'source',
        target: current,
        request: {
          kind: 'image',
          altText: '',
          source: 'https://',
        },
      });
      return;
    }
    const { from, to } = editor.state.selection;
    setInsertDialog({
      surface: 'visual',
      target: {
        document: editor.state.doc,
        from,
        text: editor.state.doc.textBetween(from, to, ' '),
        to,
      },
      request: {
        kind: 'image',
        altText: '',
        source: 'https://',
      },
    });
  };
  const commitInsert = (result: MarkdownInsertDialogResult) => {
    const dialog = insertDialog;
    if (!dialog) return;
    if (dialog.surface === 'source') {
      const insert =
        result.kind === 'image'
          ? createMarkdownImageSourceInsert(result.altText, result.source)
          : createMarkdownLinkSourceInsert(result.label, result.source);
      onSourceReplace(insert.replacement, insert.selectedRange, dialog.target);
      setInsertDialog(null);
      return;
    }
    if (editor.isDestroyed || !editor.state.doc.eq(dialog.target.document)) {
      setInsertDialog(null);
      return;
    }
    const chain = editor
      .chain()
      .focus()
      .setTextSelection({ from: dialog.target.from, to: dialog.target.to });
    if (result.kind === 'image') {
      chain
        .setImage({
          src: result.source,
          alt: result.altText || undefined,
        })
        .run();
      setInsertDialog(null);
      return;
    }
    if (dialog.target.text && dialog.target.text === result.label) {
      chain.setLink({ href: result.source }).run();
      setInsertDialog(null);
      return;
    }
    chain
      .insertContent({
        type: 'text',
        text: result.label,
        marks: [{ type: 'link', attrs: { href: result.source } }],
      })
      .run();
    setInsertDialog(null);
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
        collapsible
        className="work-markdown-ribbon"
        toolbarClassName="markdown-toolbar"
        panels={{
          home: (
            <>
              <WorkOfficeRibbonGroup label="撤销">
                <MarkdownToolbarButton
                  label="撤销"
                  shortcut={markdownToolbarShortcuts.undo}
                  disabled={!canUndo}
                  onClick={() => {
                    if (usesSharedHistory) onSourceUndo();
                    // Keep ribbon focus — chain().focus() steals into the editor.
                    else editor.commands.undo();
                  }}
                >
                  <Undo2 size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="重做"
                  shortcut={markdownToolbarShortcuts.redo}
                  disabled={!canRedo}
                  onClick={() => {
                    if (usesSharedHistory) onSourceRedo();
                    // Keep ribbon focus — chain().focus() steals into the editor.
                    else editor.commands.redo();
                  }}
                >
                  <Redo2 size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="样式">
                <OfficeSelect
                  ariaLabel="段落样式"
                  className="markdown-paragraph-style-select"
                  value={paragraphStyle}
                  options={[
                    { value: 'paragraph', label: '正文' },
                    { value: 'h1', label: '标题 1' },
                    { value: 'h2', label: '标题 2' },
                    { value: 'h3', label: '标题 3' },
                    { value: 'h4', label: '标题 4' },
                    { value: 'h5', label: '标题 5' },
                    { value: 'h6', label: '标题 6' },
                  ]}
                  onValueChange={(value) => {
                    if (value === 'paragraph') {
                      runCommand('paragraph', () => {
                        // Keep ribbon focus on the style combobox.
                        // chain().focus() schedules into the editor and breaks L2 loops.
                        editor.commands.setParagraph();
                      });
                    } else {
                      const level = Number(value.slice(1)) as
                        | 1
                        | 2
                        | 3
                        | 4
                        | 5
                        | 6;
                      runCommand(`heading-${level}`, () => {
                        // Keep ribbon focus on the style combobox.
                        // chain().focus() schedules into the editor and breaks L2 loops.
                        editor.commands.toggleHeading({ level });
                      });
                    }
                  }}
                />
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="文字">
                <MarkdownToolbarButton
                  label="加粗"
                  shortcut={markdownToolbarShortcuts.bold}
                  active={
                    sourceEditing
                      ? commandIsActive('bold')
                      : editor.isActive('bold')
                  }
                  onClick={() =>
                    runCommand('bold', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleBold();
                    })
                  }
                >
                  <Bold size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="斜体"
                  shortcut={markdownToolbarShortcuts.italic}
                  active={
                    sourceEditing
                      ? commandIsActive('italic')
                      : editor.isActive('italic')
                  }
                  onClick={() =>
                    runCommand('italic', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleItalic();
                    })
                  }
                >
                  <Italic size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="删除线"
                  active={
                    sourceEditing
                      ? commandIsActive('strike')
                      : editor.isActive('strike')
                  }
                  onClick={() =>
                    runCommand('strike', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleStrike();
                    })
                  }
                >
                  <Strikethrough size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="行内代码"
                  active={
                    sourceEditing
                      ? commandIsActive('code')
                      : editor.isActive('code')
                  }
                  onClick={() =>
                    runCommand('code', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleCode();
                    })
                  }
                >
                  <Code2 size={16} />
                </MarkdownToolbarButton>
              </WorkOfficeRibbonGroup>
              <WorkOfficeRibbonGroup label="段落">
                <MarkdownToolbarButton
                  label="项目列表"
                  active={
                    sourceEditing
                      ? commandIsActive('bullet-list')
                      : editor.isActive('bulletList')
                  }
                  onClick={() =>
                    runCommand('bullet-list', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleBulletList();
                    })
                  }
                >
                  <List size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="编号列表"
                  active={
                    sourceEditing
                      ? commandIsActive('ordered-list')
                      : editor.isActive('orderedList')
                  }
                  onClick={() =>
                    runCommand('ordered-list', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleOrderedList();
                    })
                  }
                >
                  <ListOrdered size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="任务列表"
                  active={
                    sourceEditing
                      ? commandIsActive('task-list')
                      : editor.isActive('taskList')
                  }
                  onClick={() =>
                    runCommand('task-list', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleTaskList();
                    })
                  }
                >
                  <ListChecks size={16} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="引用"
                  active={
                    sourceEditing
                      ? commandIsActive('blockquote')
                      : editor.isActive('blockquote')
                  }
                  onClick={() =>
                    runCommand('blockquote', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleBlockquote();
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
                  label={linkActive ? '编辑链接' : '添加链接'}
                  displayLabel
                  shortcut={markdownToolbarShortcuts.link}
                  active={linkActive}
                  onClick={openLinkDialog}
                >
                  <Link2 size={19} />
                </MarkdownToolbarButton>
                <MarkdownToolbarButton
                  label="移除链接"
                  displayLabel
                  disabled={!linkActive}
                  onClick={removeLink}
                >
                  <Unlink size={19} />
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
                  active={
                    sourceEditing
                      ? commandIsActive('code-block')
                      : editor.isActive('codeBlock')
                  }
                  onClick={() =>
                    runCommand('code-block', () => {
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.toggleCodeBlock();
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
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.setHorizontalRule();
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
                      // Keep ribbon focus — chain().focus() steals into the editor.
                      editor.commands.insertTable({
                        rows: 3,
                        cols: 3,
                        withHeaderRow: true,
                      });
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
                label="可视化编辑"
                visibleLabel="可视化"
                displayLabel
                active={viewMode === 'visual'}
                onClick={() => onViewModeChange('visual')}
              >
                <PencilLine size={19} />
              </MarkdownToolbarButton>
              <MarkdownToolbarButton
                label="源码编辑"
                visibleLabel="源码"
                displayLabel
                active={viewMode === 'source'}
                onClick={() => onViewModeChange('source')}
              >
                <Code2 size={19} />
              </MarkdownToolbarButton>
              <MarkdownToolbarButton
                label="分屏编辑"
                visibleLabel="分屏"
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
          request={insertDialog.request}
          restoreFocusTarget={() =>
            insertDialog.surface === 'source'
              ? (getSourceFocusTarget() ?? editor.view.dom)
              : editor.view.dom
          }
          onClose={() => setInsertDialog(null)}
          onSubmit={commitInsert}
        />
      )}
    </>
  );
}

function markdownParagraphStyleValue(
  sourceEditing: boolean,
  commandIsActive: (command: MarkdownSourceCommand) => boolean,
  editor: Editor,
): 'paragraph' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6' {
  for (const level of [1, 2, 3, 4, 5, 6] as const) {
    const active = sourceEditing
      ? commandIsActive(`heading-${level}`)
      : editor.isActive('heading', { level });
    if (active) return `h${level}`;
  }
  return 'paragraph';
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
  visibleLabel,
  shortcut,
  active = false,
  displayLabel = false,
  children,
  onMouseDown,
  ...props
}: {
  label: string;
  visibleLabel?: string;
  shortcut?: MarkdownToolbarShortcut;
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
      visibleLabel={visibleLabel}
      title={shortcut ? `${label}（${shortcut.label}）` : label}
      aria-keyshortcuts={shortcut?.ariaKeyShortcuts}
      active={active}
      displayLabel={displayLabel}
      onMouseDown={(event) => {
        // Keep focus on stay-mounted ribbon triggers for L2 loops.
        event.preventDefault();
        onMouseDown?.(event);
      }}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}
