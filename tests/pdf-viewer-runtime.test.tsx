import { expect, rs, test } from '@rstest/core';
import { render, waitFor } from '@testing-library/react';
import type { ComponentProps } from 'react';

const viewerCapture = rs.hoisted(() => ({
  props: null as ComponentProps<
    typeof import('@embedpdf/react-pdf-viewer').PDFViewer
  > | null,
}));

rs.mock('@embedpdf/react-pdf-viewer', () => ({
  PdfAnnotationSubtype: {
    FREETEXT: 'FreeText',
    HIGHLIGHT: 'Highlight',
    INK: 'Ink',
    STRIKEOUT: 'StrikeOut',
    UNDERLINE: 'Underline',
  },
  PDFViewer: (
    props: ComponentProps<
      typeof import('@embedpdf/react-pdf-viewer').PDFViewer
    >,
  ) => {
    viewerCapture.props = props;
    return <div data-testid="embedpdf-viewer" />;
  },
}));

import { PdfViewer } from '../src/internal/features/work/editors/pdf-viewer';

test('forwards the explicit PDF worker mode to EmbedPDF', async () => {
  render(
    <PdfViewer
      loadSource={async () =>
        new Blob(['%PDF-1.7'], { type: 'application/pdf' })
      }
      wasmUrl="/assets/pdfium.wasm"
      worker={false}
    />,
  );

  await waitFor(() => {
    expect(viewerCapture.props?.config).toMatchObject({
      wasmUrl: '/assets/pdfium.wasm',
      worker: false,
    });
  });
  viewerCapture.props = null;
});

test('opens PDF bytes in the worker instead of fetching a blob URL', async () => {
  const originalCreateObjectUrl = URL.createObjectURL;
  const createObjectUrl = rs.fn(() => 'blob:a3s-pdf-test');
  URL.createObjectURL = createObjectUrl;

  try {
    render(
      <PdfViewer
        fileName="note.pdf"
        loadSource={async () =>
          new Blob(['%PDF-1.3'], { type: 'application/pdf' })
        }
        wasmUrl="/assets/pdfium.wasm"
      />,
    );

    await waitFor(() => {
      expect(viewerCapture.props?.config?.worker).toBe(true);
      expect(viewerCapture.props?.config?.src).toBeUndefined();
      const document =
        viewerCapture.props?.config?.documentManager?.initialDocuments?.[0];
      expect(document && 'buffer' in document).toBe(true);
      if (document && 'buffer' in document) {
        expect(new TextDecoder().decode(document.buffer)).toBe('%PDF-1.3');
        expect(document.name).toBe('note.pdf');
      }
    });
    expect(createObjectUrl).not.toHaveBeenCalled();
  } finally {
    URL.createObjectURL = originalCreateObjectUrl;
    viewerCapture.props = null;
  }
});
