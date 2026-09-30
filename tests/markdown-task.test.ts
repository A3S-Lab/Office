import { describe, expect, test } from '@rstest/core';
import { toggleMarkdownTask } from '../src/internal/features/work/work-markdown-task';

describe('toggleMarkdownTask', () => {
  test('checks the matching preview task and leaves the others', () => {
    const source = '# 标题\n\n- [ ] 待办\n- [x] 已做';
    expect(toggleMarkdownTask(source, 0)).toBe(
      '# 标题\n\n- [x] 待办\n- [x] 已做',
    );
    expect(toggleMarkdownTask(source, 1)).toBe(
      '# 标题\n\n- [ ] 待办\n- [ ] 已做',
    );
  });
});
