import { expect, test } from '@rstest/core';
import {
  documentCommandCatalog,
  getDocumentCommandDefinition,
} from '../src/internal/features/work/editors/document-command-catalog';

test('keeps Writer command ids and WPS locations unique', () => {
  const commands = Object.values(documentCommandCatalog);
  expect(new Set(commands.map((command) => command.id)).size).toBe(
    commands.length,
  );

  expect(getDocumentCommandDefinition('bold').location).toEqual({
    area: 'ribbon',
    tab: 'home',
    group: 'font',
  });
  expect(getDocumentCommandDefinition('insertPageBreak').location).toEqual({
    area: 'ribbon',
    tab: 'insert',
    group: 'pages',
  });
  expect(getDocumentCommandDefinition('trackChanges').location).toEqual({
    area: 'ribbon',
    tab: 'review',
    group: 'tracking',
  });
  expect(getDocumentCommandDefinition('navigationPane').location).toEqual({
    area: 'ribbon',
    tab: 'view',
    group: 'show',
  });
  expect(getDocumentCommandDefinition('showHiddenText').location).toEqual({
    area: 'ribbon',
    tab: 'view',
    group: 'show',
  });
  expect(getDocumentCommandDefinition('wordCount').location).toEqual({
    area: 'status',
  });
});

test('defines the WPS Writer shortcut contract in one catalog', () => {
  expect('allCaps' in documentCommandCatalog).toBe(true);
  expect('smallCaps' in documentCommandCatalog).toBe(true);
  expect(getDocumentCommandDefinition('allCaps').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+A',
    aria: 'Control+Shift+A Meta+Shift+A',
    editor: ['Mod-Shift-a'],
  });
  expect(getDocumentCommandDefinition('smallCaps').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+K',
    aria: 'Control+Shift+K Meta+Shift+K',
    editor: ['Mod-Shift-k'],
  });
  expect(getDocumentCommandDefinition('fontDialog').shortcut).toEqual({
    label: 'Cmd/Ctrl+D',
    aria: 'Control+D Meta+D',
    editor: ['Mod-d'],
  });
  expect(getDocumentCommandDefinition('doubleUnderline').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+D',
    aria: 'Control+Shift+D Meta+Shift+D',
    editor: ['Mod-Shift-d'],
  });
  expect(getDocumentCommandDefinition('wordsUnderline').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+W',
    aria: 'Control+Shift+W Meta+Shift+W',
    editor: ['Mod-Shift-w'],
  });
  expect(getDocumentCommandDefinition('hiddenText').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+H',
    aria: 'Control+Shift+H Meta+Shift+H',
    editor: ['Mod-Shift-h'],
  });
  expect(getDocumentCommandDefinition('strike').shortcut).toBeUndefined();
  expect(getDocumentCommandDefinition('doubleStrike')).toMatchObject({
    id: 'font.doubleStrike',
    label: '双删除线',
  });
  expect(getDocumentCommandDefinition('doubleStrike').shortcut).toBeUndefined();
  expect(
    Object.values(documentCommandCatalog).some((command) =>
      command.shortcut?.editor?.includes('Mod-Shift-s'),
    ),
  ).toBe(false);
  expect(getDocumentCommandDefinition('growFont').shortcut?.editor).toEqual([
    'Mod-Shift-.',
    'Mod-]',
  ]);
  expect(getDocumentCommandDefinition('shrinkFont').shortcut?.editor).toEqual([
    'Mod-Shift-,',
    'Mod-[',
  ]);
  expect(getDocumentCommandDefinition('subscript').shortcut).toEqual({
    label: 'Cmd/Ctrl+=',
    aria: 'Control+= Meta+=',
    editor: ['Mod-='],
  });
  expect(getDocumentCommandDefinition('superscript').shortcut).toEqual({
    label: 'Cmd/Ctrl+Shift+=',
    aria: 'Control+Shift+= Meta+Shift+=',
    editor: ['Mod-Shift-='],
  });
  expect(getDocumentCommandDefinition('alignCenter').shortcut?.editor).toEqual([
    'Mod-e',
  ]);
  expect(getDocumentCommandDefinition('trackChanges').shortcut?.editor).toEqual(
    ['Mod-Shift-e'],
  );
  expect(
    getDocumentCommandDefinition('lineSpacingSingle').shortcut?.editor,
  ).toEqual(['Mod-1']);
  expect(getDocumentCommandDefinition('heading2').shortcut?.editor).toEqual([
    'Mod-Alt-2',
  ]);
  expect(
    getDocumentCommandDefinition('insertComment').shortcut?.editor,
  ).toEqual(['Mod-Alt-m']);
  expect(getDocumentCommandDefinition('copyFormat').shortcut?.editor).toEqual([
    'Mod-Shift-c',
  ]);
  expect(getDocumentCommandDefinition('pasteFormat').shortcut?.editor).toEqual([
    'Mod-Shift-v',
  ]);
  expect(getDocumentCommandDefinition('wordCount').shortcut?.editor).toEqual([
    'Mod-Shift-g',
  ]);
});
