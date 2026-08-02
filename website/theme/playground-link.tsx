import { removeBase, useLocation } from '@rspress/core/runtime';
import type { AnchorHTMLAttributes } from 'react';
import { playgroundAssetHrefFromDocsRoute } from './site-navigation';

export interface PlaygroundLinkProps
  extends Omit<AnchorHTMLAttributes<HTMLAnchorElement>, 'href'> {
  path?: string;
}

export function PlaygroundLink({ path = '', ...props }: PlaygroundLinkProps) {
  const { pathname } = useLocation();
  const href = playgroundAssetHrefFromDocsRoute(removeBase(pathname), path);

  return <a {...props} href={href} />;
}
