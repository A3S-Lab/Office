import { expect, test } from '@rstest/core';
import { createRef } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import {
  PresentationWorkspace,
  type PresentationWorkspaceCommands,
} from '../src/internal/features/work/editors/presentation-workspace';
import type {
  WorkPresentationContent,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('separates additive object selection from content editing', () => {
  const elements = [
    presentationElement('accent', 'Accent'),
    presentationElement('title', 'Title'),
  ];
  const selections: Array<{ id: string | null; additive: boolean }> = [];
  const edits: string[] = [];
  const drags: string[] = [];
  const content: WorkPresentationContent = {
    type: 'presentation',
    slides: [
      {
        id: 'slide',
        name: 'Slide',
        background: '#ffffff',
        elements,
      },
    ],
  };

  const view = render(
    <PresentationWorkspace
      activeBackground="#ffffff"
      activeCommentId={null}
      activeElements={elements}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands({
        editElement: (elementId) => edits.push(elementId),
        selectElement: (id, additive = false) =>
          selections.push({ id, additive }),
      })}
      content={content}
      designContent={content}
      designMode="slide"
      editingElementId={null}
      inheritedElements={[]}
      placeholderGuides={[]}
      selectedElementIds={[]}
      selectedLayout={undefined}
      selectedMaster={undefined}
      selectedSlide={content.slides[0]}
      snapGuides={[]}
      viewMode="normal"
      zoom={90}
      onBeginDrag={(_event, element, mode) =>
        drags.push(`${element.id}:${mode}`)
      }
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenContextMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />,
  );

  fireEvent.pointerDown(screen.getByRole('group', { name: 'Accent' }));
  expect(selections).toEqual([]);
  expect(drags).toEqual(['accent:move']);

  fireEvent.pointerDown(screen.getByRole('group', { name: 'Title' }), {
    shiftKey: true,
  });
  expect(selections).toEqual([{ id: 'title', additive: true }]);
  expect(drags).toEqual(['accent:move']);
  expect(screen.queryByRole('textbox', { name: '幻灯片文本' })).toBeNull();

  fireEvent.doubleClick(screen.getByRole('group', { name: 'Title' }));
  expect(edits).toEqual(['title']);

  view.rerender(
    <PresentationWorkspace
      activeBackground="#ffffff"
      activeCommentId={null}
      activeElements={elements}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands()}
      content={content}
      designContent={content}
      designMode="slide"
      editingElementId={null}
      inheritedElements={[]}
      placeholderGuides={[]}
      selectedElementIds={['accent', 'title']}
      selectedLayout={undefined}
      selectedMaster={undefined}
      selectedSlide={content.slides[0]}
      snapGuides={[]}
      viewMode="normal"
      zoom={90}
      onBeginDrag={(_event, element, mode) =>
        drags.push(`${element.id}:${mode}`)
      }
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenContextMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />,
  );
  expect(screen.getByText('已选择 2 个对象')).toBeVisible();
  expect(screen.getByRole('group', { name: 'Accent' })).toHaveClass('selected');
  expect(screen.getByRole('group', { name: 'Title' })).toHaveClass('selected');
  fireEvent.pointerDown(screen.getByRole('button', { name: '缩放所选对象' }));
  expect(drags).toEqual(['accent:move', 'title:resize']);
});

function presentationElement(id: string, text: string): WorkSlideElement {
  return {
    id,
    type: 'shape',
    x: id === 'accent' ? 8 : 20,
    y: id === 'accent' ? 8 : 20,
    width: 20,
    height: 12,
    text,
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
  };
}

function workspaceCommands(
  overrides: Partial<PresentationWorkspaceCommands> = {},
): PresentationWorkspaceCommands {
  return {
    addSlide: () => undefined,
    deleteSlideById: () => false,
    editElement: () => undefined,
    exitEditing: () => undefined,
    instantiatePlaceholder: () => undefined,
    openComment: () => undefined,
    selectElement: () => undefined,
    selectSlide: () => undefined,
    setViewMode: () => undefined,
    updateElement: () => undefined,
    updateNotes: () => undefined,
    updateTextElement: () => undefined,
    ...overrides,
  };
}
