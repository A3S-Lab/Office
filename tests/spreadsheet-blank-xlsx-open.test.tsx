import { expect, test } from '@rstest/core';
import { render, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
  type SpreadsheetContent,
} from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('opens a round-tripped blank workbook without React 185', async () => {
  const artifact = createArtifact('blank-spreadsheet');
  const blob = await createArtifactBlob(artifact);
  const imported = await importOfficeFile(
    new File([blob], 'Excel-blank.xlsx', { type: blob.type }),
  );
  if (imported.content.type !== 'spreadsheet') {
    throw new Error('Expected a spreadsheet.');
  }
  const sheet = imported.content.sheets[0];
  const changes: SpreadsheetContent[] = [];

  function Harness() {
    const [content, setContent] = useState(imported.content as SpreadsheetContent);
    return (
      <SpreadsheetEditor
        content={content}
        onChange={(next) => {
          changes.push(next);
          setContent(next);
        }}
        theme="light"
      />
    );
  }

  const view = render(<Harness />);
  await waitFor(() => {
    const failed = view.container.querySelector('[data-office-editor-error]');
    const grid = view.container.querySelector('.fortune-sheet-overlay');
    expect(failed ?? grid).not.toBeNull();
  });

  const failure = view.container.querySelector('[data-office-editor-error]');
  expect(failure?.textContent ?? '', JSON.stringify(sheetSummary(sheet))).not.toMatch(
    /185|Maximum update depth/,
  );
  expect(failure).toBeNull();
  expect(changes.length).toBeLessThan(8);
});

function sheetSummary(sheet: SpreadsheetContent['sheets'][number] | undefined) {
  const cell = sheet?.data?.[0]?.[0];
  return {
    keys: sheet ? Object.keys(sheet) : [],
    row: sheet?.row,
    column: sheet?.column,
    dataRows: sheet?.data?.length,
    dataColumns: sheet?.data?.[0]?.length,
    cell,
    config: sheet?.config ?? null,
    status: sheet?.status,
  };
}
