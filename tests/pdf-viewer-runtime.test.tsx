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
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  URL.createObjectURL = () => 'blob:a3s-pdf-test';
  URL.revokeObjectURL = () => undefined;

  try {
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
  } finally {
    URL.createObjectURL = originalCreateObjectUrl;
    URL.revokeObjectURL = originalRevokeObjectUrl;
    viewerCapture.props = null;
  }
});
