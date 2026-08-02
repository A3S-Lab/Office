import { expect, test } from '@rstest/core';
import {
  playgroundAssetHrefFromDocsRoute,
  playgroundHrefFromDocsRoute,
} from '../website/theme/site-navigation';

test('derives a deployment-relative Playground link from every docs route depth', () => {
  expect(playgroundHrefFromDocsRoute('/')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/index.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/native-office-engine.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/guide/')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/guide/index.html')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/components/react.html?tab=usage')).toBe(
    '../../',
  );
  expect(playgroundHrefFromDocsRoute('/en/components/react.html')).toBe(
    '../../../',
  );
  expect(playgroundHrefFromDocsRoute('/0.1.0/en/components/react.html')).toBe(
    '../../../../',
  );
});

test('derives deployment-relative Playground assets from localized version routes', () => {
  expect(
    playgroundAssetHrefFromDocsRoute(
      '/0.1.0/en/automation/',
      '/downloads/a3s-office-skill.tar.gz',
    ),
  ).toBe('../../../../downloads/a3s-office-skill.tar.gz');
  expect(playgroundAssetHrefFromDocsRoute('/', '')).toBe('../');
});
