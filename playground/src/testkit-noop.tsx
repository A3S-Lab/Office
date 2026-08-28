import type { CSSProperties, PropsWithChildren } from 'react';

type TestKitProps = PropsWithChildren<{
  enabled: boolean;
  page?: { id: string };
}>;

type ReviewOverlayProps = {
  enabled: boolean;
  locale?: 'auto' | 'en' | 'zh-CN';
  messages?: Record<string, string>;
};

type TestBoundaryProps = PropsWithChildren<{
  id: string;
  name: string;
  source?: { file: string; line?: number; column?: number };
  generated?: { file: string; line?: number; column?: number };
  ready?: () => boolean;
  facts?: () => Record<string, unknown>;
  roots?: () => readonly Element[];
  as?:
    | 'div'
    | 'section'
    | 'main'
    | 'nav'
    | 'article'
    | 'aside'
    | 'header'
    | 'footer'
    | 'span';
  className?: string;
  style?: CSSProperties;
}>;

export function A3STestKit({ children }: TestKitProps) {
  return children;
}

export function A3STestBoundary({
  as: Tag = 'div',
  children,
  className,
  style,
}: TestBoundaryProps) {
  return (
    <Tag className={className} style={style}>
      {children}
    </Tag>
  );
}

export function A3SReviewOverlay(_props: ReviewOverlayProps) {
  return null;
}
