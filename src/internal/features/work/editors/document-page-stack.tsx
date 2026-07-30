interface DocumentPageStackProps {
  pageColor: string;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
}

export function DocumentPageStack({
  pageColor,
  pageCount,
  pageGap,
  pageHeight,
}: DocumentPageStackProps) {
  const count = Math.max(1, Math.trunc(pageCount));
  const gap = Math.max(0, pageGap);
  const height = Math.max(1, pageHeight);

  return (
    <div
      className="work-document-page-stack"
      data-page-count={count}
      aria-hidden="true"
    >
      {Array.from({ length: count }, (_, pageIndex) => (
        <div
          className="work-document-page-sheet"
          data-page-index={pageIndex + 1}
          key={pageIndex}
          style={{
            backgroundColor: pageColor,
            height,
            top: pageIndex * (height + gap),
          }}
        />
      ))}
    </div>
  );
}
