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
});

test('defines the WPS Writer shortcut contract in one catalog', () => {
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
});
