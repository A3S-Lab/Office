import { expect, test } from '@rstest/core';
import { isSupportedA3sTestVersion } from '../scripts/office-ui-ops-core';

test('accepts the supported A3S Test 1.x line', () => {
  expect(isSupportedA3sTestVersion('a3s-test 1.0.1')).toBe(true);
  expect(isSupportedA3sTestVersion('a3s-test 1.0.0')).toBe(true);
});

test('rejects stale or malformed A3S Test versions', () => {
  expect(isSupportedA3sTestVersion('a3s-test 0.5.1')).toBe(false);
  expect(isSupportedA3sTestVersion('agent-browser 0.26.0')).toBe(false);
  expect(isSupportedA3sTestVersion(undefined)).toBe(false);
});
