import { expect, test } from '@rstest/core';
import { fireEvent, render } from '@testing-library/react';
import { createRef, useState } from 'react';
import { movePresentationSlide } from '../src/internal/features/work/editors/presentation-editor-operations';
import {
  PresentationWorkspace,
  type PresentationWorkspaceCommands,
} from '../src/internal/features/work/editors/presentation-workspace';
import { usePresentationSlideCommands } from '../src/internal/features/work/editors/use-presentation-slide-commands';
import { withPresentationDesign } from '../src/internal/features/work/work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkSlide,
} from '../src/internal/features/work/work-types';

test('movePresentationSlide drops a later slide in front of an earlier one', () => {
  expect(movePresentationSlide(['a', 'b', 'c'], 2, 0)).toEqual(['c', 'a', 'b']);
  expect(movePresentationSlide(['a', 'b', 'c'], 0, 2)).toEqual(['b', 'a', 'c']);
  expect(movePresentationSlide(['a', 'b', 'c'], 0, 3)).toEqual(['b', 'c', 'a']);
  expect(movePresentationSlide(['a', 'b', 'c'], 1, 2)).toEqual(['a', 'b', 'c']);
});

test('dragging slide 3 above slide 1 reorders thumbnails and shows an insertion cue', () => {
  const view = render(<ReorderHarness />);
  const restoreRects = mockThumbnailRects();
  try {
    const third = thumbnailAt(view.container, 2);
    const first = thumbnailAt(view.container, 0);
    fireEvent.pointerDown(third, { button: 0, clientY: 220, pointerId: 1 });
    fireEvent.pointerMove(first, { button: 0, clientY: 10, pointerId: 1 });
    expect(thumbnailAt(view.container, 0)).toHaveAttribute(
      'data-slide-drop-position',
      'before',
    );
    fireEvent.pointerUp(thumbnailAt(view.container, 0), {
      button: 0,
      clientY: 10,
      pointerId: 1,
    });
    expect(thumbnailIds(view.container)).toEqual([
      'slide-3',
      'slide-1',
      'slide-2',
    ]);
  } finally {
    restoreRects();
  }
});

test('dragging slide 1 below slide 3 inserts it after that slide', () => {
  const view = render(<ReorderHarness />);
  const restoreRects = mockThumbnailRects();
  try {
    fireEvent.pointerDown(thumbnailAt(view.container, 0), {
      button: 0,
      clientY: 20,
      pointerId: 1,
    });
    fireEvent.pointerMove(thumbnailAt(view.container, 2), {
      button: 0,
      clientY: 260,
      pointerId: 1,
    });
    expect(thumbnailAt(view.container, 2)).toHaveAttribute(
      'data-slide-drop-position',
      'after',
    );
    fireEvent.pointerUp(thumbnailAt(view.container, 2), {
      button: 0,
      clientY: 260,
      pointerId: 1,
    });
    expect(thumbnailIds(view.container)).toEqual([
      'slide-2',
      'slide-3',
      'slide-1',
    ]);
  } finally {
    restoreRects();
  }
});

test('a click on a thumbnail does not reorder the deck', () => {
  const view = render(<ReorderHarness />);
  fireEvent.pointerDown(thumbnailAt(view.container, 2), {
    button: 0,
    clientY: 20,
    pointerId: 1,
  });
  fireEvent.pointerUp(thumbnailAt(view.container, 2), {
    button: 0,
    clientY: 22,
    pointerId: 1,
  });
  expect(thumbnailIds(view.container)).toEqual([
    'slide-1',
    'slide-2',
    'slide-3',
  ]);
});

function ReorderHarness() {
  const [content, setContent] = useState(() => presentationContent());
  const [selectedSlideId, setSelectedSlideId] = useState('slide-1');
  const selectedSlide =
    content.slides.find((slide) => slide.id === selectedSlideId) ??
    content.slides[0];
  if (!selectedSlide) throw new Error('A selected slide is required.');
  const slideCommands = usePresentationSlideCommands({
    content,
    onChange: setContent,
    onClearSelection: () => undefined,
    onSelectSlide: setSelectedSlideId,
    selectedSlide,
  });
  return presentationWorkspace(content, selectedSlide, {
    moveSlide: slideCommands.moveSlide,
    selectSlide: (slideId) => setSelectedSlideId(slideId),
  });
}

function presentationWorkspace(
  content: WorkPresentationContent,
  selectedSlide: WorkSlide,
  commands: Partial<PresentationWorkspaceCommands>,
) {
  return (
    <PresentationWorkspace
      activeBackground={selectedSlide.background}
      activeCommentId={null}
      activeElements={selectedSlide.elements}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands(commands)}
      content={content}
      designContent={withPresentationDesign(content)}
      designMode="slide"
      editingElementId={null}
      inheritedElements={[]}
      placeholderGuides={[]}
      selectedElementIds={[]}
      selectedLayout={undefined}
      selectedMaster={undefined}
      selectedSlide={selectedSlide}
      snapGuides={[]}
      notesVisible
      viewMode="normal"
      zoom={100}
      onBeginDrag={() => undefined}
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenContextMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />
  );
}

function workspaceCommands(
  overrides: Partial<PresentationWorkspaceCommands>,
): PresentationWorkspaceCommands {
  return {
    addSlide: () => undefined,
    deleteSlideById: () => false,
    editElement: () => undefined,
    exitEditing: () => undefined,
    instantiatePlaceholder: () => undefined,
    moveSlide: () => false,
    rotateSelection: () => false,
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

function presentationContent(): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: [0, 1, 2].map((index) => presentationSlide(index)),
  };
}

function presentationSlide(index: number): WorkSlide {
  return {
    id: `slide-${index + 1}`,
    name: `Slide ${index + 1}`,
    background: '#ffffff',
    elements: [
      {
        id: `element-${index + 1}`,
        type: 'text',
        x: 10,
        y: 10,
        width: 80,
        height: 20,
        text: `Slide content ${index + 1}`,
        fontSize: 24,
        color: '#172033',
        fill: 'transparent',
        bold: true,
        align: 'center',
      },
    ],
  };
}

function thumbnailAt(container: HTMLElement, index: number): HTMLButtonElement {
  const element = container.querySelector<HTMLButtonElement>(
    `[data-slide-thumbnail][data-slide-index="${index}"]`,
  );
  if (!element) throw new Error(`Missing thumbnail ${index}.`);
  return element;
}

function thumbnailIds(container: HTMLElement): string[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>('[data-slide-thumbnail]'),
  ).map((element) => element.getAttribute('data-slide-id') ?? '');
}

function mockThumbnailRects(): () => void {
  const original = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function mockRect() {
    if (!this.hasAttribute('data-slide-thumbnail')) return original.call(this);
    const index = Number(this.getAttribute('data-slide-index'));
    const top = index * 100;
    const height = 80;
    return {
      x: 0,
      y: top,
      top,
      left: 0,
      right: 140,
      bottom: top + height,
      width: 140,
      height,
      toJSON() {
        return {};
      },
    } as DOMRect;
  };
  return () => {
    HTMLElement.prototype.getBoundingClientRect = original;
  };
}
