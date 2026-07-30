import { expect, test } from '@rstest/core';
import { createRef } from 'react';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
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

test('keeps phone slide navigation dismissible and restores focus', async () => {
  const mediaQuery = installMatchMedia(true);
  const selections: string[] = [];
  const content: WorkPresentationContent = {
    type: 'presentation',
    slides: [
      {
        id: 'slide-1',
        name: 'Opening',
        background: '#ffffff',
        elements: [],
      },
      {
        id: 'slide-2',
        name: 'Details',
        background: '#ffffff',
        elements: [],
      },
    ],
  };

  const view = render(
    <PresentationWorkspace
      activeBackground="#ffffff"
      activeCommentId={null}
      activeElements={[]}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands({
        selectSlide: (slideId) => selections.push(slideId),
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
      onBeginDrag={() => undefined}
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenContextMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />,
  );

  const layout = view.container.querySelector('.work-presentation-layout');
  const toggle = screen.getByRole('button', {
    name: '打开幻灯片导航',
  });
  expect(layout).toHaveAttribute('data-mobile-slide-navigation', 'closed');
  expect(toggle).toHaveAttribute('aria-expanded', 'false');

  fireEvent.click(toggle);
  expect(layout).toHaveAttribute('data-mobile-slide-navigation', 'open');
  expect(toggle).toHaveAttribute('aria-expanded', 'true');
  const drawer = screen.getByRole('dialog', { name: '幻灯片' });
  const close = screen.getByRole('button', { name: '关闭幻灯片导航' });
  expect(drawer).toHaveAttribute('aria-modal', 'true');
  expect(toggle).toHaveAttribute('inert');
  await waitFor(() => expect(close).toHaveFocus());

  fireEvent.keyDown(close, { key: 'Tab' });
  expect(
    screen.getByRole('button', { name: '幻灯片 1 / 2：Opening' }),
  ).toHaveFocus();
  fireEvent.keyDown(
    screen.getByRole('button', { name: '幻灯片 1 / 2：Opening' }),
    { key: 'Tab', shiftKey: true },
  );
  expect(close).toHaveFocus();

  mediaQuery.setMatches(false);
  await waitFor(() => {
    expect(drawer).not.toHaveAttribute('role');
    expect(drawer).not.toHaveAttribute('aria-modal');
    expect(toggle).not.toHaveAttribute('inert');
  });
  await waitFor(() =>
    expect(
      screen.getByRole('button', { name: '幻灯片 1 / 2：Opening' }),
    ).toHaveFocus(),
  );

  mediaQuery.setMatches(true);
  await waitFor(() => expect(drawer).toHaveAttribute('role', 'dialog'));
  expect(drawer).toHaveAttribute('aria-modal', 'true');
  await waitFor(() => expect(close).toHaveFocus());

  fireEvent.click(
    screen.getByRole('button', { name: '幻灯片 2 / 2：Details' }),
  );
  expect(selections.at(-1)).toBe('slide-2');
  expect(layout).toHaveAttribute('data-mobile-slide-navigation', 'closed');
  await waitFor(() => expect(toggle).toHaveFocus());

  fireEvent.click(toggle);
  await waitFor(() => expect(close).toHaveFocus());
  fireEvent.keyDown(close, { key: 'Escape' });
  expect(layout).toHaveAttribute('data-mobile-slide-navigation', 'closed');
  await waitFor(() => expect(toggle).toHaveFocus());
  mediaQuery.restore();
});

function installMatchMedia(initialMatches: boolean): {
  restore(): void;
  setMatches(matches: boolean): void;
} {
  const originalDescriptor = Object.getOwnPropertyDescriptor(
    window,
    'matchMedia',
  );
  const listeners = new Set<(event: MediaQueryListEvent) => void>();
  let matches = initialMatches;
  const media = '(max-width: 640px)';
  const mediaQuery = {
    get matches() {
      return matches;
    },
    media,
    onchange: null,
    addEventListener: (
      _type: 'change',
      listener: (event: MediaQueryListEvent) => void,
    ) => listeners.add(listener),
    removeEventListener: (
      _type: 'change',
      listener: (event: MediaQueryListEvent) => void,
    ) => listeners.delete(listener),
    addListener: (listener: (event: MediaQueryListEvent) => void) =>
      listeners.add(listener),
    removeListener: (listener: (event: MediaQueryListEvent) => void) =>
      listeners.delete(listener),
    dispatchEvent: () => true,
  } as MediaQueryList;

  Object.defineProperty(window, 'matchMedia', {
    configurable: true,
    writable: true,
    value: () => mediaQuery,
  });

  return {
    restore: () => {
      if (originalDescriptor) {
        Object.defineProperty(window, 'matchMedia', originalDescriptor);
      } else {
        Reflect.deleteProperty(window, 'matchMedia');
      }
    },
    setMatches: (nextMatches) => {
      matches = nextMatches;
      const event = { matches, media } as MediaQueryListEvent;
      mediaQuery.onchange?.(event);
      for (const listener of listeners) listener(event);
    },
  };
}

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
