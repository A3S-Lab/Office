import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useRef, useState } from 'react';
import {
  DocumentCompareDialog,
  type DocumentCompareDialogRequest,
} from '../src/internal/features/work/editors/document-compare-dialog';
import type { DocumentComparisonApplyResult } from '../src/internal/features/work/work-document-compare';

test('selects a revised document, attributes the comparison, and submits accessibly', async () => {
  const requests: DocumentCompareDialogRequest[] = [];
  const applied: DocumentComparisonApplyResult[] = [];
  render(
    <DocumentCompareDialog
      initialMode="compare"
      restoreFocusTarget={() => null}
      onApplied={(result) => applied.push(result)}
      onClose={() => undefined}
      onSubmit={async (request) => {
        requests.push(request);
        return appliedResult();
      }}
    />,
  );

  const dialog = screen.getByRole('dialog', { name: '比较与合并文档' });
  const choose = within(dialog).getByRole('button', {
    name: /^选择修订版本/,
  });
  await waitFor(() => expect(choose).toHaveFocus());
  expect(within(dialog).getByRole('radio', { name: /比较/ })).toBeChecked();

  const file = new File(['<p>Revised</p>'], 'contract-review.html', {
    type: 'text/html',
  });
  fireEvent.change(within(dialog).getByLabelText('选择修订版本文件'), {
    target: { files: [file] },
  });
  expect(choose).toHaveTextContent('contract-review.html');
  expect(
    within(dialog).getByRole('textbox', { name: '比较结果修订者名称' }),
  ).toHaveValue('contract-review');

  fireEvent.click(within(dialog).getByRole('button', { name: '生成比较结果' }));
  await waitFor(() => expect(requests).toHaveLength(1));
  expect(requests[0]).toMatchObject({
    author: 'contract-review',
    file,
    mode: 'compare',
  });
  expect(applied).toEqual([appliedResult()]);
});

test('explains a fail-closed reviewed-copy combine without closing the dialog', async () => {
  render(
    <DocumentCompareDialog
      initialMode="combine"
      restoreFocusTarget={() => null}
      onApplied={() => {
        throw new Error('Unsupported results must not apply.');
      }}
      onClose={() => undefined}
      onSubmit={async () => ({
        status: 'unsupported',
        summary: emptySummary(),
        diagnostics: [
          {
            code: 'combine-baseline-mismatch',
            message: 'Baseline mismatch.',
          },
        ],
      })}
    />,
  );

  const dialog = screen.getByRole('dialog', { name: '比较与合并文档' });
  expect(within(dialog).getByRole('radio', { name: /合并/ })).toBeChecked();
  expect(
    within(dialog).queryByRole('textbox', { name: '比较结果修订者名称' }),
  ).toBeNull();
  const file = new File(['review'], 'reviewed-copy.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  fireEvent.change(within(dialog).getByLabelText('选择审阅副本文件'), {
    target: { files: [file] },
  });
  fireEvent.click(within(dialog).getByRole('button', { name: '合并修订' }));

  const alert = await within(dialog).findByRole('alert');
  expect(alert).toHaveTextContent('审阅副本的原始基线与当前文档不一致');
  expect(screen.getByRole('dialog', { name: '比较与合并文档' })).toBeVisible();
});

test('restores focus to the Review-ribbon invoker after cancellation', async () => {
  render(<DialogFocusHarness />);
  const invoker = screen.getByRole('button', { name: '打开比较文档' });
  invoker.focus();
  fireEvent.click(invoker);
  const dialog = await screen.findByRole('dialog', {
    name: '比较与合并文档',
  });
  fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '比较与合并文档' })).toBeNull(),
  );
  expect(invoker).toHaveFocus();
});

function DialogFocusHarness() {
  const [open, setOpen] = useState(false);
  const invokerRef = useRef<HTMLButtonElement>(null);
  return (
    <>
      <button ref={invokerRef} type="button" onClick={() => setOpen(true)}>
        打开比较文档
      </button>
      {open && (
        <DocumentCompareDialog
          initialMode="compare"
          restoreFocusTarget={() => invokerRef.current}
          onApplied={() => setOpen(false)}
          onClose={() => setOpen(false)}
          onSubmit={async () => appliedResult()}
        />
      )}
    </>
  );
}

function appliedResult(): DocumentComparisonApplyResult {
  return {
    status: 'applied',
    summary: {
      deletions: 1,
      formatting: 0,
      insertions: 1,
      paragraphFormatting: 0,
    },
    diagnostics: [],
  };
}

function emptySummary() {
  return {
    deletions: 0,
    formatting: 0,
    insertions: 0,
    paragraphFormatting: 0,
  };
}
