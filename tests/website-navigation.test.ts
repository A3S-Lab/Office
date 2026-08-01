import { expect, test } from '@rstest/core';
import { playgroundHrefFromDocsRoute } from '../website/theme/site-navigation';

test('derives a deployment-relative Playground link from every docs route depth', () => {
  expect(playgroundHrefFromDocsRoute('/')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/index.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/native-office-engine.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/guide/')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/guide/index.html')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/components/react.html?tab=usage')).toBe(
    '../../',
  );
});
