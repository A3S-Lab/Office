import type { Editor } from '@tiptap/core';
import type {
  WorkMarkdownSelectionAction,
  WorkMarkdownSelectionContext,
  WorkMarkdownSelectionMenuItem,
  WorkMarkdownSelectionSnapshot,
} from '../work-markdown-selection-menu';
import { WorkspaceContextMenu } from '../../workspace/components/workspace-context-menu';
import { customSelectionMenuItems } from './document-editor-support';
import { createWorkMarkdownVisualSelectionAction } from '../work-markdown-selection-menu';

export interface MarkdownSelectionMenuState {
  x: number;
  y: number;
  snapshot: WorkMarkdownSelectionSnapshot;
  items: readonly WorkMarkdownSelectionMenuItem[];
}

export function MarkdownSelectionContextMenu({
  editor,
  menu,
  createSourceAction,
  onClose,
}: {
  editor: Editor;
  menu: MarkdownSelectionMenuState;
  createSourceAction: (
    snapshot: WorkMarkdownSelectionSnapshot,
  ) => WorkMarkdownSelectionAction;
  onClose: () => void;
}) {
  const createAction = (): {
    context: WorkMarkdownSelectionContext;
    dispose(): void;
  } =>
    menu.snapshot.selection.surface === 'source'
      ? createSourceAction(menu.snapshot)
      : createWorkMarkdownVisualSelectionAction(editor, menu.snapshot);

  return (
    <WorkspaceContextMenu
      label="选中文本操作"
      x={menu.x}
      y={menu.y}
      items={customSelectionMenuItems(menu.items, createAction)}
      onClose={onClose}
    />
  );
}
