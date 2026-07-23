import type { CSSProperties, ReactNode } from 'react';

export type OfficeTheme = 'light' | 'dark' | 'system';

export interface OfficeSurfaceProps {
  className?: string;
  style?: CSSProperties;
  theme?: OfficeTheme;
}

export function OfficeSurface({
  children,
  className = '',
  style,
  theme = 'system',
}: OfficeSurfaceProps & { children: ReactNode }) {
  return (
    <div
      className={`a3s-office${className ? ` ${className}` : ''}`}
      data-a3s-office=""
      data-theme={theme}
      style={style}
    >
      {children}
    </div>
  );
}
