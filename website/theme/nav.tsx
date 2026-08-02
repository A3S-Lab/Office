import { removeBase, useLocation, usePage } from '@rspress/core/runtime';
import {
  Nav as OriginalNav,
  type NavProps,
} from '@rspress/core/theme-original';
import { playgroundHrefFromDocsRoute } from './site-navigation';

export function Nav(props: NavProps) {
  const { pathname } = useLocation();
  const { page } = usePage();
  const playgroundHref = playgroundHrefFromDocsRoute(removeBase(pathname));
  const isChinese = page.lang === 'zh';

  return (
    <OriginalNav
      {...props}
      afterNavTitle={
        <>
          {props.afterNavTitle}
          <a
            className="office-docs-playground-link"
            href={playgroundHref}
            aria-label={
              isChinese ? '返回 Playground 首页' : 'Open Playground homepage'
            }
          >
            <span>{isChinese ? '在线体验' : 'Playground'}</span>
            <svg aria-hidden="true" viewBox="0 0 16 16">
              <path d="m6 3 5 5-5 5M11 8H2.5" />
            </svg>
          </a>
        </>
      }
    />
  );
}
