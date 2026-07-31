import { afterEach, expect } from '@rstest/core';
import { configure } from '@testing-library/dom';
import * as jestDomMatchers from '@testing-library/jest-dom/matchers';
import { cleanup } from '@testing-library/react';

expect.extend(jestDomMatchers);
configure({ asyncUtilTimeout: 5_000 });

afterEach(() => cleanup());
