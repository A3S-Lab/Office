import { expect, test } from '@rstest/core';
import { act, fireEvent, render } from '@testing-library/react';
import { createRef } from 'react';
import { PresentationWorkspace } from '../src/internal/features/work/editors/presentation-workspace';
import type {
  WorkPresentationContent,
  WorkSlide,
} from '../src/internal/features/work/work-types';

test('keeps every slide keyboard reachable while mounting only nearby thumbnail scenes', () => {
  const observers: MockIntersectionObserver[] = [];
  const OriginalIntersectionObserver = globalThis.IntersectionObserver;
  globalThis.IntersectionObserver = class extends MockIntersectionObserver {
    constructor(
      callback: IntersectionObserverCallback,
      options?: IntersectionObserverInit,
    ) {
      super(callback, options);
      observers.push(this);
    }
  } as typeof IntersectionObserver;

  try {
    const content = presentationContent(40);
    const view = render(presentationWorkspace(content, 'normal'));

    const thumbnails = Array.from(
      view.container.querySelectorAll<HTMLButtonElement>(
        '[data-slide-thumbnail]',
      ),
    );
    expect(thumbnails).toHaveLength(40);
    expect(observers).toHaveLength(1);
    expect(observers[0]?.observed.size).toBe(40);
    expect(renderedThumbnailCount(view.container)).toBe(1);

    act(() => {
      observers[0]?.emit([
        { target: thumbnails[10], isIntersecting: true },
        { target: thumbnails[11], isIntersecting: true },
      ]);
    });
    expect(renderedThumbnailCount(view.container)).toBe(3);

    act(() => {
      observers[0]?.emit([{ target: thumbnails[10], isIntersecting: false }]);
    });
    expect(renderedThumbnailCount(view.container)).toBe(2);

    fireEvent.keyDown(thumbnails[0], { key: 'End' });
    expect(document.activeElement).toBe(thumbnails[39]);

    view.rerender(presentationWorkspace(content, 'sorter'));
    expect(observers).toHaveLength(2);
    expect(observers[0]?.observed.size).toBe(0);
    expect(observers[1]?.observed.size).toBe(40);
    expect(observers[1]?.rootMargin).toBe('480px 240px');
  } finally {
    globalThis.IntersectionObserver = OriginalIntersectionObserver;
  }
});

test('keeps complete thumbnail rendering when intersection observation is unavailable', () => {
  const OriginalIntersectionObserver = globalThis.IntersectionObserver;
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: undefined,
    writable: true,
  });
  try {
    const content = presentationContent(4);
    const view = render(presentationWorkspace(content, 'normal'));
    expect(renderedThumbnailCount(view.container)).toBe(4);
  } finally {
    Object.defineProperty(globalThis, 'IntersectionObserver', {
      configurable: true,
      value: OriginalIntersectionObserver,
      writable: true,
    });
  }
});

class MockIntersectionObserver implements IntersectionObserver {
  readonly observed = new Set<Element>();
  readonly root: Element | Document | null;
  readonly rootMargin: string;
  readonly thresholds = [0];

  constructor(
    private readonly callback: IntersectionObserverCallback,
    options?: IntersectionObserverInit,
  ) {
    this.root = options?.root ?? null;
    this.rootMargin = options?.rootMargin ?? '0px';
  }

  disconnect(): void {
    this.observed.clear();
  }

  observe(target: Element): void {
    this.observed.add(target);
  }

  takeRecords(): IntersectionObserverEntry[] {
    return [];
  }

  unobserve(target: Element): void {
    this.observed.delete(target);
  }

  emit(entries: Array<{ target: Element; isIntersecting: boolean }>): void {
    this.callback(
      entries.map(({ target, isIntersecting }) => ({
        boundingClientRect: target.getBoundingClientRect(),
        intersectionRatio: isIntersecting ? 1 : 0,
        intersectionRect: target.getBoundingClientRect(),
        isIntersecting,
        rootBounds: null,
        target,
        time: 0,
      })),
      this,
    );
  }
}

function renderedThumbnailCount(container: HTMLElement): number {
  return container.querySelectorAll(
    '[data-slide-thumbnail][data-slide-thumbnail-rendered="true"]',
  ).length;
}

function presentationWorkspace(
  content: WorkPresentationContent,
  viewMode: 'normal' | 'sorter',
) {
  const selectedSlide = content.slides[0];
  if (!selectedSlide) throw new Error('A selected slide is required.');
  return (
    <PresentationWorkspace
      activeBackground={selectedSlide.background}
      activeCommentId={null}
      activeElements={selectedSlide.elements}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      content={content}
      designContent={content}
      designMode="slide"
      editingElementId={null}
      inheritedElements={[]}
      placeholderGuides={[]}
      selectedElementIds={[]}
      selectedLayout={undefined}
      selectedMaster={undefined}
      selectedSlide={selectedSlide}
      snapGuides={[]}
      viewMode={viewMode}
      zoom={100}
      onAddSlide={() => undefined}
      onBeginDrag={() => undefined}
      onContinueDrag={() => undefined}
      onDeleteSlide={() => false}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onEditElement={() => undefined}
      onExitEditing={() => undefined}
      onInstantiatePlaceholder={() => undefined}
      onOpenAgentMenu={() => undefined}
      onOpenComment={() => undefined}
      onSelectElement={() => undefined}
      onSelectSlide={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
      onUpdateElement={() => undefined}
      onUpdateNotes={() => undefined}
      onUpdateTextElement={() => undefined}
      onViewModeChange={() => undefined}
    />
  );
}

function presentationContent(slideCount: number): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: Array.from({ length: slideCount }, (_, index) =>
      presentationSlide(index),
    ),
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
