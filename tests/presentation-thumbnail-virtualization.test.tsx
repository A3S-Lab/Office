import { expect, test } from '@rstest/core';
import { act, fireEvent, render, waitFor } from '@testing-library/react';
import { createRef, useState } from 'react';
import {
  PresentationWorkspace,
  type PresentationWorkspaceCommands,
} from '../src/internal/features/work/editors/presentation-workspace';
import type {
  WorkPresentationContent,
  WorkSlide,
} from '../src/internal/features/work/work-types';

test('windows long decks while keeping every slide keyboard reachable', async () => {
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
    const content = presentationContent(160);
    const view = render(presentationWorkspace(content, 'normal'));

    let thumbnails = Array.from(
      view.container.querySelectorAll<HTMLButtonElement>(
        '[data-slide-thumbnail]',
      ),
    );
    expect(view.container.querySelector('[data-slide-count]')).toHaveAttribute(
      'data-slide-count',
      '160',
    );
    expect(thumbnails.length).toBeGreaterThan(1);
    expect(thumbnails.length).toBeLessThanOrEqual(40);
    expect(thumbnails[0]).toHaveAttribute(
      'aria-label',
      '幻灯片 1 / 160：Slide 1',
    );
    expect(observers).toHaveLength(1);
    expect(observers[0]?.observed.size).toBe(thumbnails.length);
    expect(renderedThumbnailCount(view.container)).toBe(1);

    act(() => {
      observers[0]?.emit([
        { target: thumbnails[10] as HTMLButtonElement, isIntersecting: true },
        { target: thumbnails[11] as HTMLButtonElement, isIntersecting: true },
      ]);
    });
    expect(renderedThumbnailCount(view.container)).toBe(3);

    act(() => {
      observers[0]?.emit([
        {
          target: thumbnails[10] as HTMLButtonElement,
          isIntersecting: false,
        },
      ]);
    });
    expect(renderedThumbnailCount(view.container)).toBe(2);

    fireEvent.keyDown(thumbnails[0] as HTMLButtonElement, { key: 'End' });
    await waitFor(() => {
      thumbnails = Array.from(
        view.container.querySelectorAll<HTMLButtonElement>(
          '[data-slide-thumbnail]',
        ),
      );
      expect(document.activeElement).toHaveAttribute('data-slide-index', '159');
      expect(thumbnails.length).toBeLessThanOrEqual(40);
    });

    view.rerender(presentationWorkspace(content, 'sorter'));
    await waitFor(() => expect(observers).toHaveLength(2));
    expect(observers[0]?.observed.size).toBe(0);
    expect(observers[1]?.rootMargin).toBe('480px 240px');
    thumbnails = Array.from(
      view.container.querySelectorAll<HTMLButtonElement>(
        '[data-slide-thumbnail]',
      ),
    );
    expect(observers[1]?.observed.size).toBe(thumbnails.length);
    expect(thumbnails.length).toBeLessThanOrEqual(48);
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
    const content = presentationContent(160);
    const view = render(presentationWorkspace(content, 'normal'));
    const mountedThumbnailCount = view.container.querySelectorAll(
      '[data-slide-thumbnail]',
    ).length;
    expect(mountedThumbnailCount).toBeLessThanOrEqual(40);
    expect(renderedThumbnailCount(view.container)).toBe(mountedThumbnailCount);
  } finally {
    Object.defineProperty(globalThis, 'IntersectionObserver', {
      configurable: true,
      value: OriginalIntersectionObserver,
      writable: true,
    });
  }
});

test('deletes a windowed slide and moves focus to the adjacent slide', async () => {
  const view = render(<StatefulPresentationWorkspace slideCount={160} />);
  const first = view.container.querySelector<HTMLButtonElement>(
    '[data-slide-thumbnail]',
  );
  if (!first) throw new Error('Missing the first thumbnail.');

  fireEvent.keyDown(first, { key: 'End' });
  await waitFor(() =>
    expect(document.activeElement).toHaveAttribute('data-slide-index', '159'),
  );

  fireEvent.keyDown(document.activeElement as HTMLButtonElement, {
    key: 'Delete',
  });
  await waitFor(() => {
    expect(view.container.querySelector('[data-slide-count]')).toHaveAttribute(
      'data-slide-count',
      '159',
    );
    expect(document.activeElement).toHaveAttribute(
      'data-slide-id',
      'slide-159',
    );
    expect(document.activeElement).toHaveAttribute('data-slide-index', '158');
  });
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
  options: {
    selectedSlideId?: string;
    onDeleteSlide?: (slideId: string) => boolean;
    onSelectSlide?: (slideId: string) => void;
  } = {},
) {
  const selectedSlide =
    content.slides.find((slide) => slide.id === options.selectedSlideId) ??
    content.slides[0];
  if (!selectedSlide) throw new Error('A selected slide is required.');
  return (
    <PresentationWorkspace
      activeBackground={selectedSlide.background}
      activeCommentId={null}
      activeElements={selectedSlide.elements}
      aspectRatio="16 / 9"
      canvasName="Slide canvas"
      canvasRef={createRef<HTMLElement>()}
      commands={workspaceCommands({
        deleteSlideById: options.onDeleteSlide ?? (() => false),
        selectSlide: (slideId) => options.onSelectSlide?.(slideId),
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
      selectedSlide={selectedSlide}
      snapGuides={[]}
      viewMode={viewMode}
      zoom={100}
      onBeginDrag={() => undefined}
      onContinueDrag={() => undefined}
      onDragCancel={() => undefined}
      onDragEnd={() => undefined}
      onOpenAgentMenu={() => undefined}
      onTextEditorChange={() => undefined}
      onTextSelectionChange={() => undefined}
    />
  );
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

function StatefulPresentationWorkspace({ slideCount }: { slideCount: number }) {
  const [content, setContent] = useState(() => presentationContent(slideCount));
  const [selectedSlideId, setSelectedSlideId] = useState(
    content.slides[0]?.id ?? '',
  );
  return presentationWorkspace(content, 'normal', {
    selectedSlideId,
    onSelectSlide: setSelectedSlideId,
    onDeleteSlide: (slideId) => {
      if (content.slides.length <= 1) return false;
      const index = content.slides.findIndex((slide) => slide.id === slideId);
      if (index < 0) return false;
      const slides = content.slides.filter((slide) => slide.id !== slideId);
      setContent({ ...content, slides });
      setSelectedSlideId(slides[Math.min(index, slides.length - 1)]?.id ?? '');
      return true;
    },
  });
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
