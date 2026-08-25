import { expect, test } from '@rstest/core';
import {
  normalizeSiteBase,
  playgroundBaseFromSiteBase,
  siteBaseFromEnvironment,
} from '../website/site-paths';

test('keeps documentation at the deployment root and Playground below it', () => {
  expect(normalizeSiteBase('/Office')).toBe('/Office/');
  expect(playgroundBaseFromSiteBase('/Office/')).toBe('/Office/playground/');
  expect(siteBaseFromEnvironment({ A3S_OFFICE_SITE_BASE: '/Preview' })).toBe(
    '/Preview/',
  );
});

test('accepts the previous deployment variable as a compatibility fallback', () => {
  expect(
    siteBaseFromEnvironment({ A3S_OFFICE_PLAYGROUND_BASE: '/Legacy/' }),
  ).toBe('/Legacy/');
  expect(siteBaseFromEnvironment({})).toBe('/');
});
