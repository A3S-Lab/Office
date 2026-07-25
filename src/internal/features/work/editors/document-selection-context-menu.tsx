import type { Editor } from '@tiptap/core';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import {
  createWorkDocumentSelectionAction,
  documentPlainTextAsHtml,
  type WorkDocumentSelectionMenuItem,
  type WorkDocumentSelectionSnapshot,
} from '../work-document-selection-menu';
import { WorkspaceContextMenu } from '../../workspace/components/workspace-context-menu';
import {
  documentAgentMenuItems,
  documentCustomSelectionMenuItems,
} from './document-editor-support';

export type DocumentSelectionMenuState =
  | {
      kind: 'agent';
      x: number;
      y: number;
      snapshot: WorkDocumentSelectionSnapshot;
    }
  | {
      kind: 'custom';
      x: number;
      y: number;
      snapshot: WorkDocumentSelectionSnapshot;
      items: readonly WorkDocumentSelectionMenuItem[];
    };

export function DocumentSelectionContextMenu({
  editor,
  menu,
  getTrackChanges,
  onAgentRequest,
  onClose,
}: {
  editor: Editor;
  menu: DocumentSelectionMenuState;
  getTrackChanges: () => boolean;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
  onClose: () => void;
}) {
  const items =
    menu.kind === 'custom'
      ? documentCustomSelectionMenuItems(menu.items, () =>
          createWorkDocumentSelectionAction(
            editor,
            menu.snapshot,
            getTrackChanges,
          ),
        )
      : onAgentRequest
        ? documentAgentMenuItems(menu.snapshot.selection.text, onAgentRequest, {
            target: {
              id: 'document-selection',
              label: '选中文本',
              before: menu.snapshot.selection.rawText,
            },
            apply: (changes) => {
              const change = changes.find(
                (candidate) => candidate.id === 'document-selection',
              );
              if (!change) {
                return { appliedTargetIds: [], conflicts: [] };
              }
              const { from, to, rawText } = menu.snapshot.selection;
              const current = editor.state.doc.textBetween(from, to, '\n');
              if (current !== rawText) {
                return {
                  appliedTargetIds: [],
                  conflicts: [
                    {
                      targetId: change.id,
                      label: change.label,
                      message: '选中文本在建议生成后已发生变化。',
                    },
                  ],
                };
              }
              const applied = getTrackChanges()
                ? editor.commands.replaceDocumentTextWithTrackedChange(
                    from,
                    to,
                    change.after,
                  )
                : editor
                    .chain()
                    .focus()
                    .setTextSelection({ from, to })
                    .insertContent(documentPlainTextAsHtml(change.after))
                    .run();
              return applied
                ? { appliedTargetIds: [change.id], conflicts: [] }
                : {
                    appliedTargetIds: [],
                    conflicts: [
                      {
                        targetId: change.id,
                        label: change.label,
                        message: '编辑器无法替换当前选区。',
                      },
                    ],
                  };
            },
          })
        : null;

  if (!items) return null;

  return (
    <WorkspaceContextMenu
      label={menu.kind === 'custom' ? '选中文本操作' : '选中文本 AI 操作'}
      x={menu.x}
      y={menu.y}
      items={items}
      onClose={onClose}
    />
  );
}
