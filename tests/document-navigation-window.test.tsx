import { Editor } from '@tiptap/core';
import { expect, rstest, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentNavigationPanel } from '../src/internal/features/work/editors/document-navigation-panel';
import { DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT } from '../src/internal/features/work/editors/document-navigation-window';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('keeps a long outline bounded and keyboard reachable', async () => {
  const { editor, element } = createLongNavigationEditor();
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = (() => 1) as typeof requestAnimationFrame;

  try {
    const view = render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    const navigation = screen.getByRole('navigation', { name: '文档标题' });

    expect(navigation).toHaveAttribute(
      'data-document-navigation-collection',
      'outline',
    );
    expect(navigation).toHaveAttribute(
      'data-document-navigation-item-count',
      '120',
    );
    expect(navigation).toHaveAttribute(
      'data-document-navigation-windowed',
      'true',
    );
    expect(
      view.container.querySelectorAll('.work-document-outline-item').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);

    fireEvent.scroll(navigation, { target: { scrollTop: 1_000_000 } });
    await waitFor(() =>
      expect(
        Number(
          navigation.getAttribute('data-document-navigation-window-start'),
        ),
      ).toBeGreaterThan(0),
    );

    const firstHeading = within(navigation).getByRole('button', {
      name: 'Section 001',
    });
    firstHeading.focus();
    fireEvent.keyDown(firstHeading, { key: 'End' });

    const lastHeading = await within(navigation).findByRole('button', {
      name: 'Section 120',
    });
    await waitFor(() => expect(lastHeading).toHaveFocus());
    expect(firstHeading).toBeInTheDocument();
    expect(firstHeading).toHaveAttribute('aria-current', 'location');
    expect(
      view.container.querySelectorAll('.work-document-outline-item').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);
    expect(
      view.container.querySelector(
        '[data-document-navigation-spacer="between"]',
      ),
    ).toBeInTheDocument();

    fireEvent.keyDown(lastHeading, { key: 'Home' });
    await waitFor(() => expect(firstHeading).toHaveFocus());
  } finally {
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
    editor.destroy();
    element.remove();
  }
});

test('keeps long search results bounded, selectable, and keyboard reachable', async () => {
  const { editor, element } = createLongNavigationEditor();
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = (() => 1) as typeof requestAnimationFrame;

  try {
    const view = render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    const search = screen.getByRole('searchbox', { name: '搜索文档' });
    fireEvent.change(search, { target: { value: 'Shared marker' } });

    const navigation = screen.getByRole('navigation', {
      name: '文档搜索结果',
    });
    expect(screen.getByText('120 个匹配')).toBeVisible();
    expect(navigation).toHaveAttribute(
      'data-document-navigation-collection',
      'search',
    );
    expect(navigation).toHaveAttribute(
      'data-document-navigation-item-count',
      '120',
    );
    expect(navigation).toHaveAttribute(
      'data-document-navigation-windowed',
      'true',
    );
    expect(
      view.container.querySelectorAll('.work-document-search-result').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);

    fireEvent.keyDown(search, { key: 'ArrowDown' });
    const firstResult = within(navigation).getByRole('button', {
      name: '第 1 个匹配：Shared marker',
    });
    await waitFor(() => expect(firstResult).toHaveFocus());
    fireEvent.keyDown(firstResult, { key: 'End' });

    const lastResult = await within(navigation).findByRole('button', {
      name: '第 120 个匹配：Shared marker',
    });
    await waitFor(() => expect(lastResult).toHaveFocus());
    expect(
      view.container.querySelectorAll('.work-document-search-result').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);
    expect(
      view.container.querySelector(
        '[data-document-navigation-spacer="before"]',
      ),
    ).toBeInTheDocument();

    fireEvent.click(lastResult);
    expect(lastResult).toHaveAttribute('aria-current', 'location');
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
      ),
    ).toBe('Shared marker');
  } finally {
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
    editor.destroy();
    element.remove();
  }
});

test('uses an instant document scroll while selecting a navigation result', () => {
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  const restoreCallbacks: FrameRequestCallback[] = [];
  globalThis.requestAnimationFrame = ((callback: FrameRequestCallback) => {
    restoreCallbacks.push(callback);
    return restoreCallbacks.length;
  }) as typeof requestAnimationFrame;
  const scrollSurface = document.createElement('div');
  scrollSurface.className = 'work-document-scroll';
  scrollSurface.style.scrollBehavior = 'smooth';
  const element = document.createElement('div');
  scrollSurface.append(element);
  document.body.append(scrollSurface);
  let behaviorDuringSelection: string | null = null;
  const editor = new Editor({
    element,
    extensions: createWorkDocumentExtensions(),
    content: '<h1>Section</h1><p>Shared marker.</p>',
  });
  const createChain = editor.chain.bind(editor);
  const chain = rstest.spyOn(editor, 'chain').mockImplementation(() => {
    const commands = createChain();
    const run = commands.run.bind(commands);
    commands.run = () => {
      behaviorDuringSelection = scrollSurface.style.scrollBehavior;
      return run();
    };
    return commands;
  });

  try {
    render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    fireEvent.change(screen.getByRole('searchbox', { name: '搜索文档' }), {
      target: { value: 'Shared marker' },
    });
    fireEvent.click(
      screen.getByRole('button', {
        name: '第 1 个匹配：Shared marker',
      }),
    );

    expect(behaviorDuringSelection).toBe('auto');
    expect(scrollSurface.style.scrollBehavior).toBe('auto');
    for (const callback of restoreCallbacks) callback(0);
    expect(scrollSurface.style.scrollBehavior).toBe('smooth');
  } finally {
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
    chain.mockRestore();
    editor.destroy();
    scrollSurface.remove();
  }
});

function createLongNavigationEditor(): {
  editor: Editor;
  element: HTMLDivElement;
} {
  const element = document.createElement('div');
  document.body.append(element);
  const content = Array.from({ length: 120 }, (_, index) => {
    const number = String(index + 1).padStart(3, '0');
    return `<h1>Section ${number}</h1><p>Shared marker ${number}.</p>`;
  }).join('');
  return {
    editor: new Editor({
      element,
      extensions: createWorkDocumentExtensions(),
      content,
    }),
    element,
  };
}
