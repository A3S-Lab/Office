import {
  removeBase,
  useLang,
  useLocation,
  useNav,
  usePage,
  useSite,
  useVersion,
  withBase,
} from '@rspress/core/runtime';
import type { NavItem } from '@rspress/core';
import {
  IconSmallMenu,
  Search,
  SocialLinks,
  SwitchAppearance,
  SvgWrapper,
  type NavProps,
} from '@rspress/core/theme-original';
import { isDarkModeSwitchEnabled } from '@rspress/shared';
import {
  NavLangs,
  NavMenu,
  NavMenuDivider,
  NavMenuItemWithChildren,
} from '@rspress/core/dist/theme/components/Nav/NavMenu.js';
import { NavScreenAppearance } from '@rspress/core/dist/theme/components/NavScreen/NavScreenAppearance.js';
import { NavScreenLangs } from '@rspress/core/dist/theme/components/NavScreen/NavScreenLangs.js';
import { NavScreenMenu } from '@rspress/core/dist/theme/components/NavScreen/NavScreenMenu.js';
import { NavScreenVersions } from '@rspress/core/dist/theme/components/NavScreen/NavScreenVersions.js';
import { NavScreenDivider } from '@rspress/core/dist/theme/components/NavScreen/index.js';
import { useNavScreen } from '@rspress/core/dist/theme/components/NavHamburger/useNavScreen.js';
import '@rspress/core/dist/theme/components/Nav/index.css';
import '@rspress/core/dist/theme/components/NavHamburger/index.css';
import '@rspress/core/dist/theme/components/NavScreen/index.css';
import { createPortal } from 'react-dom';
import { useEffect, useMemo, useRef } from 'react';
import {
  officeSiteNavigationActiveKey,
  isProductHomeRoute,
  normalizeNavigationPath,
  officeSiteNavigationItems,
  siteNavigationHref,
} from '../site-navigation';

function OfficeNavTitle({ pathname }: { pathname: string }) {
  const { site } = useSite();
  const language = useLang();
  const homeRoute = officeSiteNavigationItems(
    language === 'en' ? 'en' : 'zh',
  )[0].route;
  const homeHref = siteNavigationHref(pathname, site.base, homeRoute);
  return (
    <div className="rp-nav__title" data-site-nav-title="office">
      <a
        className="rp-nav__title__link"
        href={homeHref}
        aria-label="A3S Office home"
      >
        <div className="rp-nav__title__logo">
          <img
            src={withBase('/a3s-logo.png')}
            alt=""
            className="rspress-logo rp-nav__title__logo-image"
          />
        </div>
        <span>{site.logoText ?? 'A3S Office'}</span>
      </a>
    </div>
  );
}

function labelSocialLinks(root: ParentNode, language: string) {
  root
    .querySelectorAll<HTMLAnchorElement>(
      '.rp-social-links__item[href="https://github.com/A3S-Lab/Office"]',
    )
    .forEach((link) => {
      link.setAttribute(
        'aria-label',
        language === 'zh'
          ? '在 GitHub 上查看 A3S Office'
          : 'View A3S Office on GitHub',
      );
    });
}

function versionHref(
  pathname: string,
  currentVersion: string,
  targetVersion: string,
  defaultVersion: string,
  cleanUrls: boolean,
) {
  const parts = removeBase(pathname).split('/').filter(Boolean);

  if (currentVersion !== defaultVersion && parts[0] === currentVersion) {
    parts.shift();
  }
  if (targetVersion !== defaultVersion) {
    parts.unshift(targetVersion);
  }
  if (parts.length === 0) return '/';
  if (parts.length === 1 && targetVersion !== defaultVersion) {
    parts.push(cleanUrls ? 'index' : 'index.html');
  }
  return `/${parts.join('/')}`;
}

function dedupeNavigationItems(pathname: string, items: NavItem[]): NavItem[] {
  const seen = new Set<string>();
  return items.filter((item) => {
    if (!('link' in item)) return true;
    const normalized = normalizeNavigationPath(pathname, item.link);
    if (!normalized) return true;
    if (seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
}

function NavVersions() {
  const { pathname } = useLocation();
  const { page } = usePage();
  const { site } = useSite();
  const currentVersion = useVersion();
  const defaultVersion = site.multiVersion.default ?? '';
  const versions = site.multiVersion.versions ?? [];
  const items = versions.map((version) => ({
    text: version,
    link: versionHref(
      page.pageType === '404' ? '/' : pathname,
      currentVersion,
      version,
      defaultVersion,
      site.route?.cleanUrls ?? false,
    ),
  }));

  return items.length > 1 ? (
    <NavMenuItemWithChildren
      activeMatcher={(item) => item.text === currentVersion}
      menuItem={{ text: currentVersion, items }}
    />
  ) : null;
}

function isVisibleControl(element: HTMLElement, screen: HTMLElement) {
  const clippedGroup = element.closest<HTMLElement>(
    '.rp-nav-screen-menu-item__group, .rp-nav-screen-langs-group, .rp-nav-screen-versions-group',
  );
  if (clippedGroup && clippedGroup.getBoundingClientRect().height < 2) {
    return false;
  }

  const style = window.getComputedStyle(element);
  const bounds = element.getBoundingClientRect();
  return (
    screen.contains(element) &&
    style.display !== 'none' &&
    style.visibility !== 'hidden' &&
    bounds.width > 0 &&
    bounds.height > 0
  );
}

function OfficeNavScreen({
  isScreenOpen,
  menuItems,
}: {
  isScreenOpen: boolean;
  menuItems: NavItem[];
}) {
  useEffect(() => {
    if (!isScreenOpen) return;

    const previousOverflow = document.body.style.overflow;
    const previousPaddingRight = document.body.style.paddingRight;
    const scrollbarWidth =
      window.innerWidth - document.documentElement.clientWidth;
    document.body.style.overflow = 'hidden';
    if (scrollbarWidth > 0) {
      document.body.style.paddingRight = `${scrollbarWidth}px`;
    }

    return () => {
      document.body.style.overflow = previousOverflow;
      document.body.style.paddingRight = previousPaddingRight;
    };
  }, [isScreenOpen]);

  return (
    <div
      className={`rp-nav-screen${isScreenOpen ? ' rp-nav-screen--open' : ''}`}
    >
      <div className="rp-nav-screen__container">
        <NavScreenMenu menuItems={menuItems} />
        <NavScreenDivider />
        <NavScreenAppearance />
        <NavScreenLangs />
        <NavScreenVersions />
        <NavScreenDivider />
        <SocialLinks />
      </div>
    </div>
  );
}

function NavHamburger({ menuItems }: { menuItems: NavItem[] }) {
  const { closeScreen, isScreenOpen, toggleScreen } = useNavScreen();
  const language = useLang();
  const trigger = useRef<HTMLButtonElement>(null);
  const closeScreenRef = useRef(closeScreen);
  closeScreenRef.current = closeScreen;
  const activeClass = isScreenOpen ? ' rp-nav-hamburger--active' : '';
  const openLabel = language === 'zh' ? '打开导航' : 'Open navigation';
  const closeLabel = language === 'zh' ? '关闭导航' : 'Close navigation';

  useEffect(() => {
    if (!isScreenOpen) return;

    let dispose = () => {};
    const frame = window.requestAnimationFrame(() => {
      const screen = document.querySelector<HTMLElement>(
        '.rp-nav-screen--open',
      );
      if (!screen) return;

      screen.id = 'office-mobile-navigation';
      screen.setAttribute(
        'aria-label',
        language === 'zh' ? '站点导航' : 'Site navigation',
      );
      screen.setAttribute('aria-modal', 'true');
      screen.setAttribute('role', 'dialog');

      const enhancedControls = new Map<HTMLElement, () => void>();
      let controlIndex = 0;
      const enhanceControls = () => {
        labelSocialLinks(screen, language);

        const controls = screen.querySelectorAll<HTMLElement>(
          '.rp-nav-screen-menu-item:not(a), .rp-nav-screen-langs, .rp-nav-screen-versions, .rp-switch-appearance',
        );

        controls.forEach((control) => {
          if (enhancedControls.has(control)) return;

          const cleanups: Array<() => void> = [];
          const isThemeControl = control.matches('.rp-switch-appearance');
          control.setAttribute('role', 'button');
          control.tabIndex = 0;

          if (isThemeControl) {
            control.setAttribute(
              'aria-label',
              language === 'zh' ? '切换主题' : 'Toggle theme',
            );
          } else {
            const group = control.nextElementSibling as HTMLElement | null;
            const groupId = `office-mobile-navigation-group-${controlIndex++}`;
            if (group) {
              group.id = groupId;
              control.setAttribute('aria-controls', groupId);
            }

            const syncExpanded = () => {
              const expanded =
                control.classList.contains('rp-nav-screen-menu-item--open') ||
                control.querySelector('[class*="__icon--open"]') !== null ||
                group?.classList.contains(
                  'rp-nav-screen-versions-group--open',
                ) ||
                group?.style.gridTemplateRows === '1fr';
              control.setAttribute('aria-expanded', String(Boolean(expanded)));
              group?.setAttribute('aria-hidden', String(!expanded));
              group?.toggleAttribute('inert', !expanded);
              group
                ?.querySelectorAll<HTMLAnchorElement>('a[href]')
                .forEach((link) => {
                  link.tabIndex = expanded ? 0 : -1;
                });
            };
            syncExpanded();

            const syncAfterClick = () =>
              window.requestAnimationFrame(syncExpanded);
            control.addEventListener('click', syncAfterClick);
            cleanups.push(() => {
              control.removeEventListener('click', syncAfterClick);
              group?.removeAttribute('aria-hidden');
              group?.removeAttribute('inert');
              group
                ?.querySelectorAll<HTMLAnchorElement>('a[href]')
                .forEach((link) => {
                  link.removeAttribute('tabindex');
                });
            });
          }

          const activateWithKeyboard = (event: KeyboardEvent) => {
            if (event.key !== 'Enter' && event.key !== ' ') return;
            event.preventDefault();
            control.click();
          };
          control.addEventListener('keydown', activateWithKeyboard);
          cleanups.push(() =>
            control.removeEventListener('keydown', activateWithKeyboard),
          );
          enhancedControls.set(control, () => {
            cleanups.forEach((cleanup) => {
              cleanup();
            });
          });
        });
      };

      enhanceControls();
      const observer = new MutationObserver(enhanceControls);
      observer.observe(screen, { childList: true, subtree: true });

      const focusableControls = () =>
        Array.from(
          screen.querySelectorAll<HTMLElement>(
            'a[href], button:not([disabled]), [tabindex="0"]',
          ),
        ).filter((element) => isVisibleControl(element, screen));

      const handleKeyDown = (event: KeyboardEvent) => {
        if (event.key === 'Escape') {
          event.preventDefault();
          closeScreenRef.current();
          return;
        }
        if (event.key !== 'Tab') return;

        const focusable = focusableControls();
        if (focusable.length === 0) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      };

      screen.addEventListener('keydown', handleKeyDown);
      focusableControls()[0]?.focus();
      dispose = () => {
        observer.disconnect();
        screen.removeEventListener('keydown', handleKeyDown);
        enhancedControls.forEach((cleanup) => {
          cleanup();
        });
      };
    });

    return () => {
      window.cancelAnimationFrame(frame);
      dispose();
      if (trigger.current && document.contains(trigger.current)) {
        trigger.current.focus();
      }
    };
  }, [isScreenOpen, language]);

  const modalContainer = isScreenOpen
    ? document.getElementById('__rspress_modal_container')
    : null;

  return (
    <>
      {isScreenOpen &&
        modalContainer &&
        createPortal(
          <OfficeNavScreen isScreenOpen={isScreenOpen} menuItems={menuItems} />,
          modalContainer,
        )}
      <button
        aria-controls="office-mobile-navigation"
        aria-expanded={isScreenOpen}
        aria-label={isScreenOpen ? closeLabel : openLabel}
        className={`rp-nav-hamburger rp-nav-hamburger__sm rp-nav-hamburger__md${activeClass}`}
        onClick={toggleScreen}
        ref={trigger}
        type="button"
      >
        <SvgWrapper icon={IconSmallMenu} />
      </button>
    </>
  );
}

export function Nav({
  beforeNavTitle,
  afterNavTitle,
  beforeNavMenu,
  afterNavMenu,
  navTitle,
}: NavProps) {
  const navList = useNav();
  const language = useLang();
  const { pathname } = useLocation();
  const { site } = useSite();
  const navigationLists = useMemo(() => {
    const playgroundRoute = site.base.endsWith('/docs/')
      ? '/playground/'
      : '/playground/index.html';
    const activeKey = officeSiteNavigationActiveKey(pathname, site.base);
    const onProductHome = isProductHomeRoute(pathname, site.base);
    const routeItems: NavItem[] = officeSiteNavigationItems(
      language === 'en' ? 'en' : 'zh',
      playgroundRoute,
    )
      .filter((item) => item.key !== 'home' || onProductHome)
      .map((item) => ({
        text: item.label,
        link: siteNavigationHref(pathname, site.base, item.route),
        // `matchNavbar` only receives the router pathname.  An explicit active
        // key keeps SSR and hydration identical even when Rspress temporarily
        // reports a path relative to its `/docs/` mount.
        activeMatch: item.key === activeKey ? '.*' : '(?!)',
        // The home route is represented by the logo, just like A3S Flow. Keep
        // the two reading/product routes on the left and the Playground action
        // on the right so the header never changes shape between surfaces.
        position: item.key === 'playground' ? 'right' : 'left',
      }));
    // Keep only explicit utility menus (Resources today). Rspress may expose
    // an auto-generated documentation tree through `useNav`; that tree
    // belongs in the sidebar, not in the global product navigation.
    const utilityItems = navList.filter(
      (item) =>
        item.position === 'right' ||
        ('items' in item &&
          item.items.some(
            (child) => 'link' in child && /^https?:\/\//.test(child.link),
          )),
    );
    const deduped = dedupeNavigationItems(pathname, [
      ...routeItems,
      ...utilityItems,
    ]);
    const left = deduped.filter(
      (item) =>
        item.position === 'left' &&
        'link' in item &&
        !normalizeNavigationPath(pathname, item.link)?.match(/^(?:\/$)/),
    );
    const right = deduped.filter(
      (item) => item.position === 'right' && 'link' in item,
    );
    // Reuse the same page-aware route contract in the mobile sheet. On docs
    // pages the logo remains the home destination, while the redundant
    // "Product home" text item is intentionally omitted on every surface.
    const mobile = dedupeNavigationItems(pathname, [
      ...routeItems,
      ...utilityItems,
    ]);
    return { left, right, mobile };
  }, [language, navList, pathname, site.base]);
  const leftNavigationList = navigationLists.left;
  const navigationList = navigationLists.right;
  const mobileNavigationList = navigationLists.mobile;
  const hasAppearanceSwitch = isDarkModeSwitchEnabled(
    site.themeConfig.darkMode,
  );

  useEffect(() => {
    labelSocialLinks(document, language);
  }, [language]);

  useEffect(() => {
    const mobileSearch = document.querySelector<HTMLElement>(
      '.rp-search-button--mobile',
    );
    if (!mobileSearch) return;

    mobileSearch.setAttribute(
      'aria-label',
      language === 'zh' ? '搜索文档' : 'Search documentation',
    );
    mobileSearch.setAttribute('role', 'button');
    mobileSearch.tabIndex = 0;

    const activateWithKeyboard = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      event.stopPropagation();
      mobileSearch.click();
    };
    mobileSearch.addEventListener('keydown', activateWithKeyboard);
    return () =>
      mobileSearch.removeEventListener('keydown', activateWithKeyboard);
  }, [language]);

  useEffect(() => {
    const themeSwitch = document.querySelector<HTMLElement>(
      '.rp-nav__others .rp-switch-appearance',
    );
    if (!themeSwitch) return;

    themeSwitch.setAttribute(
      'aria-label',
      language === 'zh' ? '切换主题' : 'Toggle theme',
    );
    themeSwitch.setAttribute('role', 'button');
    themeSwitch.tabIndex = 0;

    const activateWithKeyboard = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      themeSwitch.click();
    };
    themeSwitch.addEventListener('keydown', activateWithKeyboard);
    return () =>
      themeSwitch.removeEventListener('keydown', activateWithKeyboard);
  }, [language]);

  useEffect(() => {
    const stopClosedSearchEnter = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' || !(event.target instanceof HTMLElement)) {
        return;
      }
      if (event.target.closest('.rp-search-panel__modal')) return;

      const interactive = event.target.closest(
        'a[href], button, input, select, textarea, [role="button"], [role="tab"], [contenteditable="true"]',
      );
      if (interactive) event.stopPropagation();
    };

    document.body.addEventListener('keydown', stopClosedSearchEnter);
    return () =>
      document.body.removeEventListener('keydown', stopClosedSearchEnter);
  }, []);

  return (
    <header className="rp-nav" data-site-navigation="office">
      <div className="rp-nav__left">
        {beforeNavTitle}
        {navTitle ?? <OfficeNavTitle pathname={pathname} />}
        <NavMenu menuItems={leftNavigationList} position="left" />
        {afterNavTitle}
      </div>
      <div className="rp-nav__right">
        {beforeNavMenu}
        <Search />
        <NavMenu menuItems={navigationList} position="right" />
        <div className="rp-nav__others">
          <NavMenuDivider />
          <NavLangs />
          <NavVersions />
          {hasAppearanceSwitch && <SwitchAppearance />}
          <SocialLinks />
        </div>
        <NavHamburger menuItems={mobileNavigationList} />
        {afterNavMenu}
      </div>
    </header>
  );
}
