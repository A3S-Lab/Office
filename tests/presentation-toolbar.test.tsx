import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from '../src/internal/features/work/editors/presentation-command-types';
import { PresentationToolbar } from '../src/internal/features/work/editors/presentation-toolbar';
import type {
  WorkSlide,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('routes character-format buttons through selection-aware commands', () => {
  const calls: string[] = [];
  const commands = new Proxy(
    {},
    {
      get: (_target, property) => () => {
        calls.push(String(property));
        return true;
      },
    },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '加粗' }));
  fireEvent.click(screen.getByRole('button', { name: '斜体' }));
  fireEvent.click(screen.getByRole('button', { name: '下划线' }));

  expect(calls).toEqual(['toggleBold', 'toggleItalic', 'toggleUnderline']);
});

test('advertises the implemented history and clipboard shortcuts', () => {
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  const shortcuts = [
    ['撤销', 'Control+Z Meta+Z'],
    ['重做', 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y'],
    ['复制', 'Control+C Meta+C'],
    ['剪切', 'Control+X Meta+X'],
    ['粘贴', 'Control+V Meta+V'],
  ] as const;
  for (const [name, shortcut] of shortcuts) {
    expect(screen.getByRole('button', { name })).toHaveAttribute(
      'aria-keyshortcuts',
      shortcut,
    );
  }
});

test('disables character-format buttons from command capabilities', () => {
  const can = new Proxy(
    {},
    {
      get: (_target, property) => () => property !== 'toggleItalic',
    },
  ) as PresentationEditorCanCommands;
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  expect(screen.getByRole('button', { name: '加粗' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '斜体' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '下划线' })).toBeEnabled();
});

test('disables new-slide actions when the command surface cannot add slides', () => {
  const can = new Proxy(
    {},
    {
      get: (_target, property) => () => property !== 'addSlide',
    },
  ) as PresentationEditorCanCommands;
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={null}
      selectedUnitCount={0}
      can={can}
      textFormattingAvailable={false}
      commentsOpen={false}
      commentCount={0}
      designOpen
      editingDesign
      transition={slide.transition}
      commands={commands}
    />,
  );

  expect(screen.getByRole('button', { name: '新建幻灯片' })).toBeDisabled();
});

test('disables transition mutations when slide editing is unavailable', () => {
  const can = new Proxy(
    {},
    {
      get: (_target, property) => () =>
        property !== 'setTransition' && property !== 'applyTransitionToAll',
    },
  ) as PresentationEditorCanCommands;
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;
  const transition = {
    type: 'fade',
    speed: 'medium',
    advanceOnClick: true,
    advanceAfterMs: 5000,
  } as const;

  render(
    <PresentationToolbar
      selectedSlide={{ ...slide, transition }}
      selectedElement={null}
      selectedUnitCount={0}
      can={can}
      textFormattingAvailable={false}
      commentsOpen={false}
      commentCount={0}
      designOpen
      editingDesign
      transition={transition}
      commands={commands}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '切换' }));
  expect(
    screen.getByRole('combobox', { name: '幻灯片切换效果' }),
  ).toBeDisabled();
  expect(
    screen.getByRole('button', {
      name: '应用切换效果到全部幻灯片',
    }),
  ).toBeDisabled();
});

test('commits a validated presentation font size instead of editing intermediates', () => {
  const updates: Array<Partial<WorkSlideElement>> = [];
  const commands = new Proxy(
    {},
    {
      get: (_target, property) =>
        property === 'updateElement'
          ? (patch: Partial<WorkSlideElement>) => updates.push(patch)
          : () => true,
    },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  const fontSize = screen.getByRole('textbox', { name: '演示字号' });
  fireEvent.change(fontSize, { target: { value: '' } });
  expect(updates).toEqual([]);
  expect(fontSize).toHaveValue('');

  fireEvent.change(fontSize, { target: { value: '120' } });
  expect(fontSize).toHaveAttribute('aria-invalid', 'true');
  expect(updates).toEqual([]);
  fireEvent.blur(fontSize);

  expect(fontSize).toHaveValue('96');
  expect(updates).toEqual([{ fontSize: 96 }]);

  fireEvent.change(fontSize, { target: { value: '72' } });
  fireEvent.keyDown(fontSize, { key: 'Escape' });
  expect(fontSize).toHaveValue('24');
  expect(updates).toEqual([{ fontSize: 96 }]);
});

test('preserves a font size draft while the active text selection settles', () => {
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;
  const renderToolbar = (selectedElement: WorkSlideElement) => (
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={selectedElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />
  );
  const { rerender } = render(renderToolbar(textElement));
  const fontSize = screen.getByRole('textbox', { name: '演示字号' });

  fireEvent.change(fontSize, { target: { value: '' } });
  rerender(renderToolbar({ ...textElement, fontSize: 38 }));
  expect(fontSize).toHaveValue('');

  fireEvent.change(fontSize, { target: { value: '72' } });
  rerender(renderToolbar({ ...textElement, id: 'element-2', fontSize: 18 }));
  expect(fontSize).toHaveValue('18');
});

test('renders presentation font size as one standard ribbon control', () => {
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  const fontGroup = screen.getByRole('region', { name: '字体' });
  const fontSize = screen.getByRole('textbox', { name: '演示字号' });
  expect(fontSize.parentElement).toHaveClass('work-office-number-field');
  expect(fontGroup.querySelectorAll('.work-office-number-field')).toHaveLength(
    1,
  );
  expect(fontGroup.querySelector('.presentation-number-field')).toBeNull();
  expect(fontGroup.textContent).not.toContain('字号');
});

test('validates presentation links and restores the invoking menu control', async () => {
  const updates: Array<Partial<WorkSlideElement>> = [];
  const commands = new Proxy(
    {},
    {
      get: (_target, property) =>
        property === 'updateElement'
          ? (patch: Partial<WorkSlideElement>) => updates.push(patch)
          : () => true,
    },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;

  render(
    <PresentationToolbar
      selectedSlide={slide}
      selectedElement={textElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  const trigger = screen.getByRole('button', { name: '链接' });
  trigger.focus();
  fireEvent.click(trigger);
  const input = screen.getByRole('textbox', { name: '链接地址' });
  const apply = screen.getByRole('button', { name: '应用链接' });

  fireEvent.change(input, { target: { value: 'javascript:alert(1)' } });
  expect(apply).toBeDisabled();
  expect(
    screen.getByText('请输入完整的 http、https、mailto 或 # 文档内地址。'),
  ).toBeInTheDocument();

  fireEvent.change(input, { target: { value: ' https://a3s.dev/office ' } });
  expect(apply).toBeEnabled();
  fireEvent.click(apply);

  await waitFor(() => {
    expect(updates).toEqual([{ href: 'https://a3s.dev/office' }]);
    expect(trigger).toHaveFocus();
  });
});

test('shows the actual imported font in the presentation font menu', async () => {
  const importedElement: WorkSlideElement = {
    ...textElement,
    fontFamily: 'Calibri',
  };
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;
  const commands = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCommands;

  render(
    <PresentationToolbar
      selectedSlide={{ ...slide, elements: [importedElement] }}
      selectedElement={importedElement}
      selectedUnitCount={1}
      can={can}
      textFormattingAvailable
      commentsOpen={false}
      commentCount={0}
      designOpen={false}
      editingDesign={false}
      transition={slide.transition}
      commands={commands}
    />,
  );

  const font = screen.getByRole('combobox', { name: '演示字体' });
  expect(font).toHaveTextContent('Calibri');
  expect(font.querySelector('span')).toHaveStyle({ fontFamily: 'Calibri' });
  fireEvent.click(font);
  const importedOption = screen.getByRole('option', { name: 'Calibri' });
  expect(importedOption).toHaveAttribute('aria-selected', 'true');
  await waitFor(() => expect(importedOption).toHaveFocus());
});

const textElement: WorkSlideElement = {
  id: 'element-1',
  type: 'text',
  x: 0,
  y: 0,
  width: 320,
  height: 80,
  text: 'Selection-aware formatting',
  fontSize: 24,
  color: '#111827',
  fill: 'transparent',
  bold: false,
  align: 'left',
};

const slide: WorkSlide = {
  id: 'slide-1',
  name: 'Slide 1',
  background: '#ffffff',
  elements: [textElement],
};
