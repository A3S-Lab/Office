import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { createArtifact } from '../src/core';
import {
  AssistantPanel,
  createPlaygroundAssistantQuestionRequest,
  createPlaygroundSelectionMenuItems,
  EditorExportButton,
} from '../playground/src/editor-workspace';

test('keeps compact export actions named when their text is hidden', () => {
  const view = render(
    <EditorExportButton
      kind="markdown"
      exporting={false}
      onExport={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '导出' })).toHaveAttribute(
    'aria-label',
    '导出',
  );

  view.rerender(
    <EditorExportButton
      kind="pdf"
      exporting={false}
      onExport={() => undefined}
    />,
  );
  expect(screen.getByRole('button', { name: '下载 PDF' })).toHaveAttribute(
    'title',
    '下载 PDF',
  );
});

test('keeps Word export formats behind one compact header action', () => {
  let docxExports = 0;
  let pdfExports = 0;
  render(
    <EditorExportButton
      kind="document"
      exporting={false}
      onExport={() => {
        docxExports += 1;
      }}
      onExportPdf={() => {
        pdfExports += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  expect(screen.getByRole('menu', { name: '导出格式' })).toBeVisible();
  fireEvent.click(screen.getByRole('menuitem', { name: '导出 PDF' }));
  expect(pdfExports).toBe(1);
  expect(docxExports).toBe(0);

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  fireEvent.click(screen.getByRole('menuitem', { name: '下载 DOCX' }));
  expect(docxExports).toBe(1);
});

test('contains compact assistant focus and restores its exact trigger', () => {
  render(<ModalAssistantHarness />);

  const trigger = screen.getByRole('button', { name: '打开 AI 助手' });
  trigger.focus();
  fireEvent.click(trigger);

  const assistant = screen.getByRole('dialog', { name: 'AI 助手' });
  const close = screen.getByRole('button', { name: '关闭 AI 助手' });
  const background = screen.getByRole('button', { name: '编辑器操作' });
  expect(assistant).toHaveAttribute('aria-modal', 'true');
  expect(background.closest<HTMLElement>('[inert]')).toBeInTheDocument();
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Tab' });
  expect(
    screen.getByRole('separator', { name: '调整 AI 助手宽度' }),
  ).toHaveFocus();
  fireEvent.keyDown(
    screen.getByRole('separator', { name: '调整 AI 助手宽度' }),
    { key: 'Tab', shiftKey: true },
  );
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: 'AI 助手' })).toBeNull();
  expect(trigger).toHaveFocus();
});

test('keeps assistant empty states contextual and free of integration code', () => {
  const document = createArtifact('blank-document');
  const view = render(
    <AssistantPanel
      artifact={document}
      lastRequest={null}
      modal={false}
      questionDraft={null}
      width={460}
      onCancelQuestion={() => undefined}
      onClose={() => undefined}
      onQuestionChange={() => undefined}
      onSubmitQuestion={() => undefined}
      onWidthChange={() => undefined}
    />,
  );

  expect(screen.getByRole('heading', { name: '从选中文本开始' })).toBeVisible();
  expect(
    screen.getByText('选中文本后，可从右键菜单发起扩写、润色或提问。'),
  ).toBeVisible();
  expect(screen.queryByText(/onAgentRequest|DocumentEditor/)).toBeNull();

  const pdf = createArtifact('blank-document');
  pdf.kind = 'pdf';
  pdf.title = 'sample';
  pdf.content = { type: 'pdf' };
  view.rerender(
    <AssistantPanel
      artifact={pdf}
      lastRequest={null}
      modal={false}
      questionDraft={null}
      width={460}
      onCancelQuestion={() => undefined}
      onClose={() => undefined}
      onQuestionChange={() => undefined}
      onSubmitQuestion={() => undefined}
      onWidthChange={() => undefined}
    />,
  );

  expect(
    screen.getByRole('heading', { name: '当前 PDF 暂无 AI 请求' }),
  ).toBeVisible();
  expect(
    screen.getByText('PDF 阅读与批注可直接使用；摘要与问答接入仍在完善。'),
  ).toBeVisible();
});

test('collects an explicit question before creating an AI request', async () => {
  const requests: string[] = [];
  const drafts: Array<{
    question: string;
    selectionContext: string;
    selectionPreview: string;
  }> = [];
  const items = createPlaygroundSelectionMenuItems(
    (request) => requests.push(request.instruction),
    (draft) => drafts.push(draft),
    () => undefined,
  );
  const ask = items.find((item) => item.id === 'ask');

  await ask?.onSelect({
    selection: {
      text: '关键结论',
      beforeText: '前文',
      afterText: '后文',
    },
    document: { text: '完整文档' },
    commands: { copyText: async () => true },
  });

  expect(requests).toEqual([]);
  expect(drafts).toHaveLength(1);
  expect(drafts[0]?.selectionPreview).toBe('关键结论');
  expect(drafts[0]?.selectionContext).toContain('完整文档：\n完整文档');
  const draft = drafts[0];
  if (!draft) throw new Error('Expected an assistant question draft.');

  expect(
    createPlaygroundAssistantQuestionRequest({
      ...draft,
      question: '  这段结论有哪些依据？  ',
    }),
  ).toEqual({
    instruction: '请结合已附带的文档上下文回答：\n\n这段结论有哪些依据？',
    selection: draft.selectionContext,
  });
  expect(
    createPlaygroundAssistantQuestionRequest({
      ...draft,
      question: '   ',
    }),
  ).toBeNull();
});

test('renders a focused assistant question composer before submission', () => {
  const questionChanges: string[] = [];
  let submissions = 0;
  const artifact = createArtifact('blank-document');
  const draft = {
    question: '',
    selectionContext: '选中文本和完整文档上下文',
    selectionPreview: '关键结论',
  };
  const view = render(
    <AssistantPanel
      artifact={artifact}
      lastRequest={null}
      modal={false}
      questionDraft={draft}
      width={460}
      onCancelQuestion={() => undefined}
      onClose={() => undefined}
      onQuestionChange={(question) => questionChanges.push(question)}
      onSubmitQuestion={() => {
        submissions += 1;
      }}
      onWidthChange={() => undefined}
    />,
  );

  const question = screen.getByRole('textbox', { name: '向 AI 助手提问' });
  expect(question).toHaveFocus();
  expect(screen.getByRole('button', { name: '发送问题' })).toBeDisabled();
  expect(screen.getByText('关键结论')).toBeVisible();

  fireEvent.change(question, {
    target: { value: '这段结论有哪些依据？' },
  });
  expect(questionChanges).toEqual(['这段结论有哪些依据？']);

  view.rerender(
    <AssistantPanel
      artifact={artifact}
      lastRequest={null}
      modal={false}
      questionDraft={{ ...draft, question: '这段结论有哪些依据？' }}
      width={460}
      onCancelQuestion={() => undefined}
      onClose={() => undefined}
      onQuestionChange={() => undefined}
      onSubmitQuestion={() => {
        submissions += 1;
      }}
      onWidthChange={() => undefined}
    />,
  );
  fireEvent.click(screen.getByRole('button', { name: '发送问题' }));
  expect(submissions).toBe(1);
});

test('keeps prepared request context collapsed until the user asks for it', () => {
  const artifact = createArtifact('blank-document');
  render(
    <AssistantPanel
      artifact={artifact}
      lastRequest={{
        instruction: '润色这段内容',
        selection: '完整文档上下文',
      }}
      modal={false}
      questionDraft={null}
      width={460}
      onCancelQuestion={() => undefined}
      onClose={() => undefined}
      onQuestionChange={() => undefined}
      onSubmitQuestion={() => undefined}
      onWidthChange={() => undefined}
    />,
  );

  const details = screen.getByText('查看附带上下文').closest('details');
  expect(details).not.toHaveAttribute('open');
  expect(screen.getByText('完整文档上下文')).not.toBeVisible();
});

function ModalAssistantHarness() {
  const [open, setOpen] = useState(false);
  const artifact = createArtifact('blank-document');
  return (
    <div>
      <button type="button" onClick={() => setOpen(true)}>
        打开 AI 助手
      </button>
      {open && (
        <AssistantPanel
          artifact={artifact}
          lastRequest={null}
          modal
          questionDraft={null}
          width={460}
          onCancelQuestion={() => undefined}
          onClose={() => setOpen(false)}
          onQuestionChange={() => undefined}
          onSubmitQuestion={() => undefined}
          onWidthChange={() => undefined}
        />
      )}
      <section>
        <button type="button">编辑器操作</button>
      </section>
    </div>
  );
}
