import { expect, type Locator, type Page, test } from '@playwright/test';
import {
  openDocumentFixture,
  stabilizeVisualSurface,
  waitForDocumentFixture,
} from './visual-test-support';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';

type VisualEditorKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

interface VisualFixture {
  kind: VisualEditorKind;
  open: (page: Page) => Promise<void>;
  ready: (page: Page) => Promise<void>;
}

const fixtures: VisualFixture[] = [
  {
    kind: 'document',
    open: openDocumentFixture,
    ready: waitForDocumentFixture,
  },
  {
    kind: 'markdown',
    open: (page) =>
      page
        .getByRole('button', {
          name: '# 产品说明 MD · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page.locator('.work-markdown-editor .ProseMirror').waitFor();
    },
  },
  {
    kind: 'spreadsheet',
    open: (page) =>
      page
        .getByRole('button', {
          name: '季度执行计划 XLSX · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page
        .locator('.work-spreadsheet-canvas > .fortune-container')
        .waitFor();
    },
  },
  {
    kind: 'presentation',
    open: (page) =>
      page
        .getByRole('button', {
          name: '业务策略汇报 PPTX · 本次会话',
        })
        .click(),
    ready: async (page) => {
      await page.locator('.work-slide-canvas.interactive').waitFor();
      await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
        'data-presentation-geometry-state',
        'idle',
      );
    },
  },
  {
    kind: 'pdf',
    open: openPdfFixture,
    ready: waitForPdfFixture,
  },
];

test.describe('Office editor visual contracts', () => {
  for (const fixture of fixtures) {
    test(`${fixture.kind} editor`, async ({ page }, testInfo) => {
      await page.goto('/');
      await fixture.open(page);
      await fixture.ready(page);
      await stabilizeVisualSurface(page);
      await verifySharedEditorGeometry(
        page,
        fixture.kind,
        testInfo.project.name,
      );
      await expect(page).toHaveScreenshot(`${fixture.kind}.png`);
    });
  }
});

test('shared command bar remains compact when editors enter preview', async ({
  page,
}) => {
  for (const fixture of fixtures.filter(({ kind }) => kind !== 'pdf')) {
    await page.goto('/');
    await fixture.open(page);
    await fixture.ready(page);
    await page.getByRole('button', { name: '预览' }).click();

    const shell = page.locator(`.work-editor-shell.${fixture.kind}`);
    const header = shell.getByRole('toolbar', { name: '文件命令栏' });
    const previewBar = shell.getByRole('region', {
      name: previewToolbarLabel(fixture.kind),
    });
    await expect(header).toBeVisible();
    await expect(previewBar).toBeVisible();

    const geometry = await shell.evaluate((element) => {
      const header = element.querySelector<HTMLElement>(
        '.playground-editor-header',
      );
      const previewBar = element.querySelector<HTMLElement>(
        '.work-office-preview-bar',
      );
      const host = element.querySelector<HTMLElement>(
        '.playground-editor-host',
      );
      if (!(header && previewBar && host)) {
        throw new Error('Preview command-bar geometry is unavailable.');
      }
      const headerRect = header.getBoundingClientRect();
      const previewRect = previewBar.getBoundingClientRect();
      const hostRect = host.getBoundingClientRect();
      return {
        headerHeight: headerRect.height,
        headerTop: headerRect.top,
        previewHeight: previewRect.height,
        previewTop: previewRect.top,
        hostTop: hostRect.top,
      };
    });
    expect(geometry.headerHeight).toBe(36);
    expect(geometry.previewHeight).toBe(36);
    expect(geometry.headerTop).toBeCloseTo(geometry.previewTop, 0);
    expect(geometry.hostTop).toBeCloseTo(geometry.previewTop, 0);
  }
});

test('PDF keeps its single compact command row at phone width', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'pdf');
  if (!fixture) throw new Error('Missing PDF visual fixture.');

  await page.setViewportSize({ width: 390, height: 700 });
  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const geometry = await page
    .locator('.work-editor-shell.pdf')
    .evaluate((shell) => {
      const header = shell.querySelector<HTMLElement>(
        '.playground-editor-header',
      );
      const toolbar = shell.querySelector<HTMLElement>('.work-pdf-toolbar');
      if (!(header && toolbar)) {
        throw new Error('Compact PDF command bar is incomplete.');
      }
      const headerRect = header.getBoundingClientRect();
      const toolbarRect = toolbar.getBoundingClientRect();
      return {
        viewportWidth: document.documentElement.clientWidth,
        documentScrollWidth: document.documentElement.scrollWidth,
        headerTop: headerRect.top,
        headerHeight: headerRect.height,
        toolbarTop: toolbarRect.top,
        toolbarHeight: toolbarRect.height,
        toolbarOverflowX: getComputedStyle(toolbar).overflowX,
      };
    });

  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.viewportWidth + 1,
  );
  expect(geometry.headerHeight).toBe(42);
  expect(geometry.toolbarHeight).toBe(42);
  expect(geometry.toolbarTop).toBeCloseTo(geometry.headerTop, 0);
  expect(geometry.toolbarOverflowX).toBe('auto');
});

test('PDF prioritizes page and zoom controls at compact workspace width', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'pdf');
  if (!fixture) throw new Error('Missing PDF visual fixture.');

  await page.setViewportSize({ width: 768, height: 800 });
  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const toolbar = page.getByRole('toolbar', { name: 'PDF 工具栏' });
  await expect(toolbar.locator('.work-pdf-save')).toBeHidden();
  await expect(toolbar.locator('.work-pdf-page-controls')).toBeVisible();
  await expect(toolbar.getByRole('button', { name: '上一页' })).toBeHidden();
  await expect(toolbar.getByRole('button', { name: '下一页' })).toBeHidden();
  await expect(toolbar.getByRole('button', { name: '缩小' })).toBeVisible();
  await expect(toolbar.getByRole('button', { name: '放大' })).toBeVisible();
  await expect(toolbar.getByLabel('PDF 缩放比例')).toBeHidden();
});

test('PDF keeps page and zoom controls inside the desktop command row', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'This contract targets the 1280px editor command row.',
  );
  const fixture = fixtures.find((candidate) => candidate.kind === 'pdf');
  if (!fixture) throw new Error('Missing PDF visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const toolbar = page.getByRole('toolbar', { name: 'PDF 工具栏' });
  const geometry = await toolbar.evaluate((element) => {
    const toolbarRect = element.getBoundingClientRect();
    const pageControls = element.querySelector<HTMLElement>(
      '.work-pdf-page-controls',
    );
    const zoomControls = element.querySelector<HTMLElement>(
      '.work-pdf-zoom-controls',
    );
    if (!(pageControls && zoomControls)) {
      throw new Error('PDF navigation controls are unavailable.');
    }
    const pageRect = pageControls.getBoundingClientRect();
    const zoomRect = zoomControls.getBoundingClientRect();
    return {
      clientWidth: element.clientWidth,
      pageRight: pageRect.right,
      scrollWidth: element.scrollWidth,
      toolbarRight: toolbarRect.right,
      zoomRight: zoomRect.right,
    };
  });

  expect(geometry.scrollWidth).toBeLessThanOrEqual(geometry.clientWidth + 1);
  expect(geometry.pageRight).toBeLessThanOrEqual(geometry.toolbarRight + 1);
  expect(geometry.zoomRight).toBeLessThanOrEqual(geometry.toolbarRight + 1);
});

test('Markdown keeps source and preview side by side at compact width', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'compact-768',
    'This contract targets the 768px split workspace.',
  );
  const fixture = fixtures.find((candidate) => candidate.kind === 'markdown');
  if (!fixture) throw new Error('Missing Markdown visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const geometry = await page
    .locator('.work-markdown-workspace.split')
    .evaluate((workspace) => {
      const source = workspace.querySelector<HTMLElement>(
        '.work-markdown-pane.source',
      );
      const preview = workspace.querySelector<HTMLElement>(
        '.work-markdown-pane.visual',
      );
      const splitter = workspace.querySelector<HTMLElement>(
        '.work-markdown-splitter',
      );
      if (!(source && preview && splitter)) {
        throw new Error('Markdown split workspace is incomplete.');
      }
      const sourceRect = source.getBoundingClientRect();
      const previewRect = preview.getBoundingClientRect();
      return {
        previewLeft: previewRect.left,
        previewTop: previewRect.top,
        sourceRight: sourceRect.right,
        sourceTop: sourceRect.top,
        splitterDisplay: getComputedStyle(splitter).display,
      };
    });

  expect(geometry.sourceTop).toBeCloseTo(geometry.previewTop, 0);
  expect(geometry.sourceRight).toBeLessThanOrEqual(geometry.previewLeft);
  expect(geometry.splitterDisplay).not.toBe('none');
});

test('presentation transition controls keep standard ribbon geometry', async ({
  page,
}) => {
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await page.getByRole('tab', { name: '切换', exact: true }).click();

  const toolbar = page.getByRole('toolbar', { name: '切换工具栏' });
  await expect(toolbar).toBeVisible();
  for (const label of ['切换效果', '换片方式', '应用']) {
    await expect(page.getByRole('region', { name: label })).toBeVisible();
  }
  const geometry = await toolbar.evaluate((element) => {
    const toolbarRect = element.getBoundingClientRect();
    return {
      toolbar: {
        height: toolbarRect.height,
        top: toolbarRect.top,
        bottom: toolbarRect.bottom,
      },
      groups: [
        ...element.querySelectorAll<HTMLElement>(
          ':scope > .work-office-ribbon-group',
        ),
      ].map((group) => {
        const rect = group.getBoundingClientRect();
        return {
          label: group.getAttribute('aria-label'),
          width: rect.width,
          height: rect.height,
          top: rect.top,
          bottom: rect.bottom,
        };
      }),
    };
  });

  expect(geometry.toolbar.height).toBe(74);
  expect(geometry.groups.map(({ label }) => label)).toEqual([
    '切换效果',
    '换片方式',
    '应用',
  ]);
  const minimumWidths = new Map([
    ['切换效果', 175],
    ['换片方式', 150],
    ['应用', 56],
  ]);
  for (const group of geometry.groups) {
    expect(group.height).toBe(65);
    expect(group.top).toBeGreaterThanOrEqual(geometry.toolbar.top);
    expect(group.bottom).toBeLessThanOrEqual(geometry.toolbar.bottom + 1);
    expect(group.width).toBeGreaterThanOrEqual(
      minimumWidths.get(group.label ?? '') ?? 56,
    );
  }
});

test('compact spreadsheet ribbon advances to a complete group', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'compact-768',
    'This contract targets the compact ribbon navigation.',
  );
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'spreadsheet',
  );
  if (!fixture) throw new Error('Missing spreadsheet visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  const toolbar = page.getByRole('toolbar', { name: '开始工具栏' });
  const forward = page.getByRole('button', {
    name: '向右查看更多开始工具',
  });
  await expect(forward).toBeVisible();
  await forward.click();
  const backward = page.getByRole('button', {
    name: '向左查看更多开始工具',
  });
  await expect(backward).toBeVisible();
  await expect
    .poll(() => toolbar.evaluate((element) => element.scrollLeft))
    .toBeGreaterThan(0);

  const geometry = await toolbar.evaluate((element) => {
    const toolbarRect = element.getBoundingClientRect();
    const panel = element.closest('.work-office-ribbon-panel');
    const backwardButton = panel?.querySelector<HTMLElement>(
      '.work-office-ribbon-scroll.previous',
    );
    if (!backwardButton) {
      throw new Error('Compact ribbon backward navigation is unavailable.');
    }
    const backwardRect = backwardButton.getBoundingClientRect();
    const inset = Number.parseFloat(
      getComputedStyle(element).scrollPaddingLeft,
    );
    const expectedLeft = toolbarRect.left + inset;
    const groups = [
      ...element.querySelectorAll<HTMLElement>(
        ':scope > .work-office-ribbon-group',
      ),
    ].map((group) => {
      const rect = group.getBoundingClientRect();
      return {
        label: group.getAttribute('aria-label'),
        left: rect.left,
        right: rect.right,
      };
    });
    const aligned = groups.reduce((closest, group) =>
      Math.abs(group.left - expectedLeft) <
      Math.abs(closest.left - expectedLeft)
        ? group
        : closest,
    );
    return {
      aligned,
      backwardRight: backwardRect.right,
      expectedLeft,
      scrollLeft: element.scrollLeft,
      toolbarRight: toolbarRect.right,
    };
  });

  expect(geometry.scrollLeft).toBeGreaterThan(0);
  expect(geometry.aligned.left).toBeCloseTo(geometry.expectedLeft, 0);
  expect(geometry.aligned.left).toBeGreaterThanOrEqual(
    geometry.backwardRight + 2,
  );
  expect(geometry.aligned.right).toBeLessThanOrEqual(geometry.toolbarRight - 2);

  await page.setViewportSize({ width: 1280, height: 800 });
  await expect(backward).toBeHidden();
  await expect(forward).toBeHidden();
  await expect
    .poll(() => toolbar.evaluate((element) => element.scrollLeft))
    .toBe(0);
});

test('compact shared ribbon tabs remain readable without overlapping labels', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'compact-768',
    'This contract targets compact command-bar tabs.',
  );
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const tabList = page.getByRole('tablist', { name: '演示功能区' });
  const labelsOverlap = await tabList.evaluate((element) => {
    const labels = [
      ...element.querySelectorAll<HTMLElement>(':scope > button > span'),
    ]
      .filter((label) => getComputedStyle(label).display !== 'none')
      .map((label) => label.getBoundingClientRect());
    return labels.some(
      (label, index) => index > 0 && label.left < labels[index - 1].right - 0.5,
    );
  });
  expect(labelsOverlap).toBe(false);
});

test('compact spreadsheet task panels stay contained and keyboard dismissible', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'compact-768',
    'This contract targets compact workbook task panels.',
  );
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'spreadsheet',
  );
  if (!fixture) throw new Error('Missing spreadsheet visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const verifyPanel = async (panelName: string, bodyName: string) => {
    const panel = page.getByRole('region', { name: panelName, exact: true });
    const body = page.getByRole('region', { name: bodyName, exact: true });
    const close = panel.getByRole('button', { name: '关闭工作簿设置' });
    await expect(panel).toBeVisible();
    await expect(body).toBeVisible();
    await expect(close).toBeVisible();

    const before = await panel.evaluate((element) => {
      const header = element.querySelector<HTMLElement>(':scope > header');
      const body = element.querySelector<HTMLElement>(
        ':scope > .work-spreadsheet-workbook-panel-body',
      );
      if (!(header && body)) {
        throw new Error('Workbook task-panel geometry is incomplete.');
      }
      const panelRect = element.getBoundingClientRect();
      const headerRect = header.getBoundingClientRect();
      const bodyRect = body.getBoundingClientRect();
      return {
        panelTop: panelRect.top,
        panelBottom: panelRect.bottom,
        headerTop: headerRect.top,
        headerBottom: headerRect.bottom,
        bodyTop: bodyRect.top,
        bodyBottom: bodyRect.bottom,
        bodyClientWidth: body.clientWidth,
        bodyScrollWidth: body.scrollWidth,
        bodyOverflowY: getComputedStyle(body).overflowY,
      };
    });
    expect(before.headerTop).toBeCloseTo(before.panelTop, 0);
    expect(before.bodyTop).toBeCloseTo(before.headerBottom, 0);
    expect(before.bodyBottom).toBeLessThanOrEqual(before.panelBottom + 1);
    expect(before.bodyScrollWidth).toBeLessThanOrEqual(
      before.bodyClientWidth + 1,
    );
    expect(before.bodyOverflowY).toBe('auto');

    const actions = body.locator('.actions').last();
    if ((await actions.count()) > 0) {
      await expect(actions).toBeVisible();
      const actionBottom = await actions.evaluate(
        (element) => element.getBoundingClientRect().bottom,
      );
      expect(actionBottom).toBeLessThanOrEqual(before.bodyBottom + 1);
    }

    await body.evaluate((element) => {
      element.scrollTop = element.scrollHeight;
    });
    await expect(close).toBeVisible();
    await body.locator('button, input, [tabindex]').first().focus();
    await page.keyboard.press('Escape');
    await expect(panel).toBeHidden();
  };

  await page.getByRole('tab', { name: '公式', exact: true }).click();
  await page.getByRole('button', { name: /公式与计算/ }).click();
  await verifyPanel('公式与计算', '公式与计算内容');

  await page.getByRole('button', { name: '名称管理器' }).click();
  await verifyPanel('名称管理器', '名称管理器内容');

  await page.getByRole('tab', { name: '数据', exact: true }).click();
  await page.getByRole('button', { name: '数据透视表' }).click();
  await verifyPanel('数据透视表管理器', '数据透视表内容');

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: /^插入图表/ }).click();
  await page.getByRole('button', { name: '根据当前选区新建' }).click();
  await verifyPanel('图表管理器', '工作簿图表内容');

  await page.getByRole('button', { name: '条件格式' }).click();
  await verifyPanel('条件格式管理器', '条件格式内容');

  await page.getByRole('tab', { name: '页面布局', exact: true }).click();
  await page.getByRole('button', { name: '打印设置' }).click();
  await verifyPanel('打印设置', '打印设置内容');

  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '工作表保护' }).click();
  await verifyPanel('工作表保护', '工作表保护内容');
});

test('closing PDF annotation style keeps the active pen', async ({ page }) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'pdf');
  if (!fixture) throw new Error('Missing PDF visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  const pen = page.getByRole('button', { name: '画笔' });
  await pen.click();
  await expect(pen).toHaveAttribute('aria-pressed', 'true');

  await page.getByRole('button', { name: '批注样式' }).click();
  const style = page.getByRole('dialog', { name: '批注样式' });
  await expect(style).toBeVisible();
  const fullOpacity = style.getByRole('radio', { name: '透明度 100%' });
  await expect(fullOpacity).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  const reducedOpacity = style.getByRole('radio', { name: '透明度 75%' });
  await expect(reducedOpacity).toBeFocused();
  await expect(reducedOpacity).toBeChecked();
  await fullOpacity.click({ timeout: 2_000 });
  await expect(fullOpacity).toBeChecked();
  await page.keyboard.press('Escape');

  await expect(style).toBeHidden();
  await expect(pen).toHaveAttribute('aria-pressed', 'true');
});

test.describe('Office editor context menu contracts', () => {
  test.describe.configure({ mode: 'serial' });

  for (const kind of ['document', 'spreadsheet', 'presentation'] as const) {
    test(`${kind} context menu`, async ({ page }) => {
      const fixture = fixtures.find((candidate) => candidate.kind === kind);
      if (!fixture) throw new Error(`Missing ${kind} visual fixture.`);

      await page.goto('/');
      await fixture.open(page);
      await fixture.ready(page);
      await openEditorContextMenu(page, kind);

      const menu = page.getByRole('menu');
      await expect(menu).toBeVisible();
      await expect(menu.getByRole('menuitem').first()).toBeFocused();

      const geometry = await menu.evaluate((element) => {
        const rect = element.getBoundingClientRect();
        return {
          position: getComputedStyle(element).position,
          viewportWidth: document.documentElement.clientWidth,
          viewportHeight: document.documentElement.clientHeight,
          left: rect.left,
          top: rect.top,
          right: rect.right,
          bottom: rect.bottom,
          width: rect.width,
          height: rect.height,
        };
      });
      expect(geometry.position).toBe('fixed');
      expect(geometry.left).toBeGreaterThanOrEqual(8);
      expect(geometry.top).toBeGreaterThanOrEqual(8);
      expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
      expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
      expect(geometry.width).toBeGreaterThanOrEqual(200);
      expect(geometry.width).toBeLessThanOrEqual(320);
      expect(geometry.height).toBeGreaterThan(40);

      if (kind === 'document') {
        await expect(
          menu.getByRole('menuitem', { name: '扩写选中内容' }),
        ).toBeVisible();
        await expect(
          menu.getByRole('menuitem', { name: '润色表达' }),
        ).toBeVisible();
        await expect(
          menu.getByRole('menuitem', { name: '总结选中内容' }),
        ).toHaveCount(0);
        await expect(menu).toHaveScreenshot(
          'document-selection-context-menu.png',
        );
        await menu.getByRole('menuitem', { name: '扩写选中内容' }).click();
        const assistant = page.getByRole('complementary', {
          name: 'AI 助手',
        });
        await expect(assistant).toBeVisible();
        await expect(assistant).toContainText('完整文档：');
        await expect(assistant).toContainText('新项目方案');
        return;
      }

      await page.keyboard.press('Escape');
      await expect(menu).toBeHidden();
    });
  }
});

test('component guide highlights framework examples by language', async ({
  page,
}) => {
  await page.goto('/#guide');
  await expect(page.getByRole('heading', { name: '接入文档' })).toBeVisible();

  const example = page.locator(
    '.playground-framework-example .playground-code-block pre',
  );
  await expect(example).toHaveAttribute('data-code-language', 'tsx');
  await expect(example.locator('.token.keyword').first()).toBeVisible();

  await page.getByRole('tab', { name: 'Vue' }).click();
  await expect(example).toHaveAttribute('data-code-language', 'markup');
  await expect(example.locator('.token.tag').first()).toBeVisible();

  await page.getByRole('tab', { name: 'Web Component' }).click();
  await expect(example).toHaveAttribute('data-code-language', 'typescript');
  await expect(example.locator('.token.keyword').first()).toBeVisible();
});

test('component guide provides framework-specific examples', async ({
  page,
}) => {
  await page.goto('/#guide');
  await expect(
    page.getByRole('heading', { name: '接入文档', level: 1 }),
  ).toBeVisible();
  await expect(page.getByRole('tab', { name: 'React' })).toBeVisible();

  await page.getByRole('tab', { name: 'Vue' }).click();
  await expect(page.locator('.playground-framework-example pre')).toContainText(
    'v-model:content',
  );

  await page.getByRole('tab', { name: 'Web Component' }).click();
  await expect(page.locator('.playground-framework-example pre')).toContainText(
    'defineA3SOfficeElements',
  );
});

test('unified guide keeps CLI and Skill setup in one document', async ({
  page,
}) => {
  await page.goto('/#cli');
  await expect(
    page.getByRole('heading', { name: '接入文档', level: 1 }),
  ).toBeVisible();
  await expect(page).toHaveURL(/#guide\/automation$/);
  await expect(
    page.getByRole('heading', { name: '命令行与 AI', level: 2 }),
  ).toBeVisible();
  await expect(
    page.getByRole('link', { name: '下载 CLI Skill' }),
  ).toBeVisible();
  await expect(page.getByRole('tablist', { name: '接入内容' })).toHaveCount(0);

  const guideNavigation = page.getByRole('navigation', {
    name: '接入方式',
  });
  await guideNavigation.getByRole('link', { name: '前端组件' }).click();
  await expect(page).toHaveURL(/#guide\/components$/);
  await expect(
    page.getByRole('heading', { name: '前端组件', level: 2 }),
  ).toBeVisible();

  await guideNavigation.getByRole('link', { name: '组件 API' }).click();
  await expect(page).toHaveURL(/#guide\/api$/);
  await expect(
    page.getByRole('heading', { name: '组件 API', level: 2 }),
  ).toBeVisible();

  await guideNavigation.getByRole('link', { name: '命令行与 AI' }).click();
  await expect(page).toHaveURL(/#guide\/automation$/);
  await expect(
    page.getByRole('heading', { name: '安装 CLI Skill', level: 3 }),
  ).toBeVisible();

  const openSidebar = page.getByRole('button', {
    name: '展开办公侧边栏',
  });
  if (await openSidebar.isVisible()) await openSidebar.click();
  await page
    .getByRole('navigation', { name: '产品页面' })
    .getByRole('button', { name: '接入文档' })
    .click();
  await expect(page).toHaveURL(/#guide$/);
  await expect(guideNavigation).toBeVisible();
  await expect(
    guideNavigation.getByRole('link', { name: '前端组件' }),
  ).toHaveAttribute('href', '#guide/components');
  await expect(
    guideNavigation.getByRole('link', { name: '组件 API' }),
  ).toHaveAttribute('href', '#guide/api');
  await expect(
    guideNavigation.getByRole('link', { name: '命令行与 AI' }),
  ).toHaveAttribute('href', '#guide/automation');
});

test('component API documents every editor and remains usable on compact screens', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/#guide/api');

  const api = page.locator('section[id="guide/api"]');
  await expect(
    api.getByRole('heading', { name: '组件 API', level: 2 }),
  ).toBeVisible();
  const tabs = api.getByRole('tablist', { name: '编辑器 API' });
  await expect(tabs.getByRole('tab')).toHaveCount(5);
  await expect(
    api.getByRole('rowheader', { name: 'extensions' }),
  ).toBeVisible();
  await expect(
    api.getByRole('rowheader', { name: 'getSelectionMenuItems' }),
  ).toBeVisible();
  await expect(
    api.getByRole('heading', { name: '选区右键菜单', level: 3 }),
  ).toBeVisible();
  await expect(
    api.getByText('支持 TipTap Extensions', { exact: true }),
  ).toBeVisible();
  const example = api.locator('pre[data-code-language="tsx"]');
  await expect(example.locator('.token.keyword').first()).toBeVisible();

  const propsTable = api.locator('.playground-api-table-wrap').nth(1);
  expect(
    await propsTable.evaluate(
      (element) => element.scrollWidth > element.clientWidth,
    ),
  ).toBe(true);

  await tabs.getByRole('tab', { name: /PDF.*PdfViewer/ }).click();
  await expect(
    api.getByRole('heading', { name: '文件生命周期', level: 3 }),
  ).toBeVisible();
  await expect(
    api.getByRole('rowheader', { name: 'loadSource' }),
  ).toBeVisible();
});

test('Markdown GFM source and visual panes stay synchronized', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'markdown');
  if (!fixture) throw new Error('Missing Markdown visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await expect(
    page.getByRole('checkbox', {
      name: '已完成：安装 @a3s-lab/office',
    }),
  ).toBeChecked();
  await expect(
    page.getByRole('checkbox', {
      name: '未完成：接入宿主应用的持久化',
    }),
  ).not.toBeChecked();
  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await source.fill('# Intermediate title');
  const longDocument = [
    '# Synchronized Markdown',
    '',
    '- [ ] Review the synchronized preview',
    '',
    ...Array.from({ length: 80 }, (_, index) => [
      `## Section ${index + 1}`,
      '',
      `Paragraph ${index + 1} keeps both panes independently scrollable.`,
      '',
    ]).flat(),
  ].join('\n');
  await source.fill(longDocument);

  await expect(
    page.getByRole('heading', { name: 'Synchronized Markdown' }),
  ).toBeVisible();
  await expect(
    page.getByRole('heading', { name: 'Intermediate title' }),
  ).toHaveCount(0);
  const task = page.getByRole('checkbox', {
    name: '未完成：Review the synchronized preview',
  });
  await expect(task).toBeDisabled();

  await source.evaluate((element) => {
    element.scrollTop = (element.scrollHeight - element.clientHeight) * 0.55;
    element.dispatchEvent(new Event('scroll', { bubbles: true }));
  });
  await expect
    .poll(() =>
      page
        .getByRole('region', { name: 'Markdown 预览窗格' })
        .evaluate((element) => element.scrollTop),
    )
    .toBeGreaterThan(0);

  await page.getByRole('tab', { name: '视图' }).click();
  const viewControls = page.getByRole('region', { name: '编辑方式' });
  await viewControls.getByRole('button', { name: '编辑' }).click();
  const editableTask = page.getByRole('checkbox', {
    name: '未完成：Review the synchronized preview',
  });
  await expect(editableTask).toBeEnabled();
  await editableTask.click();
  await viewControls.getByRole('button', { name: '分屏' }).click();
  await expect(
    page.getByRole('textbox', { name: 'Markdown 源码' }),
  ).toHaveValue(/- \[x\] Review the synchronized preview/);
});

test('presentation transforms snap visually and commit one undo step', async ({
  page,
}) => {
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const editor = page.locator('.work-presentation-editor');
  const canvas = page.locator('.work-slide-canvas.interactive');
  const title = canvas.locator(':scope > .work-slide-element').nth(1);
  const undo = page.getByRole('button', { name: '撤销', exact: true });
  const redo = page.getByRole('button', { name: '重做', exact: true });
  const canvasBox = await canvas.boundingBox();
  const titleBox = await title.boundingBox();
  if (!canvasBox || !titleBox) {
    throw new Error('Presentation transform geometry is unavailable.');
  }

  await page.mouse.move(
    titleBox.x + titleBox.width / 2,
    titleBox.y + titleBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    canvasBox.x + canvasBox.width / 2,
    titleBox.y + titleBox.height / 2,
  );

  await expect(editor).toHaveAttribute(
    'data-presentation-transform-state',
    'dragging',
  );
  await expect(editor).toHaveAttribute(
    'data-presentation-geometry-engine',
    'wasm',
  );
  const guide = canvas.locator('[data-presentation-snap-guide="x"]');
  await expect(guide).toHaveAttribute('data-presentation-snap-source', 'slide');
  await expect
    .poll(() =>
      title.evaluate((element) =>
        Number.parseFloat((element as HTMLElement).style.left),
      ),
    )
    .toBe(14);
  await expect(undo).toBeDisabled();

  const guideGeometry = await guide.evaluate((element) => {
    const guideRect = element.getBoundingClientRect();
    const canvasRect = element.parentElement?.getBoundingClientRect();
    if (!canvasRect) throw new Error('Presentation guide canvas is missing.');
    return {
      centerOffset:
        guideRect.left +
        guideRect.width / 2 -
        (canvasRect.left + canvasRect.width / 2),
      heightDifference: Math.abs(guideRect.height - canvasRect.height),
    };
  });
  expect(Math.abs(guideGeometry.centerOffset)).toBeLessThanOrEqual(1);
  expect(guideGeometry.heightDifference).toBeLessThanOrEqual(2);

  await page.mouse.up();
  await expect(editor).toHaveAttribute(
    'data-presentation-transform-state',
    'idle',
  );
  await expect(guide).toHaveCount(0);
  await expect(undo).toBeEnabled();
  await expect(redo).toBeDisabled();

  await undo.click();
  await expect
    .poll(() =>
      title.evaluate((element) =>
        Number.parseFloat((element as HTMLElement).style.left),
      ),
    )
    .toBe(8);
  await expect(undo).toBeDisabled();
  await expect(redo).toBeEnabled();
});

test('presentation keeps object selection separate from text editing', async ({
  page,
}, testInfo) => {
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const editor = page.locator('.work-presentation-editor');
  const canvas = page.locator('.work-slide-canvas.interactive');
  const elements = canvas.locator(':scope > .work-slide-element');
  const accent = elements.nth(0);
  const title = elements.nth(1);
  const undo = page.getByRole('button', { name: '撤销', exact: true });

  await accent.click();
  await expect(accent).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(canvas.getByRole('textbox', { name: '幻灯片文本' })).toHaveCount(
    0,
  );

  await title.click({ modifiers: ['Shift'] });
  await expect(accent).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(title).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(
    canvas.locator('[data-presentation-selection-frame]'),
  ).toBeVisible();
  await expect(page.getByText('已选择 2 个对象')).toBeVisible();
  await expect(canvas.getByRole('textbox', { name: '幻灯片文本' })).toHaveCount(
    0,
  );

  const accentBefore = await presentationElementPosition(accent);
  const titleBefore = await presentationElementPosition(title);
  const titleBox = await title.boundingBox();
  if (!titleBox) throw new Error('Presentation title geometry is unavailable.');
  await page.mouse.move(
    titleBox.x + titleBox.width / 2,
    titleBox.y + titleBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    titleBox.x + titleBox.width / 2 + 28,
    titleBox.y + titleBox.height / 2 + 18,
  );
  await page.mouse.up();
  await expect(editor).toHaveAttribute(
    'data-presentation-transform-state',
    'idle',
  );
  const accentAfter = await presentationElementPosition(accent);
  const titleAfter = await presentationElementPosition(title);
  expect(accentAfter.x - accentBefore.x).toBeCloseTo(
    titleAfter.x - titleBefore.x,
    4,
  );
  expect(accentAfter.y - accentBefore.y).toBeCloseTo(
    titleAfter.y - titleBefore.y,
    4,
  );

  await undo.click();
  await expect
    .poll(() => presentationElementPosition(accent))
    .toEqual(accentBefore);
  await expect
    .poll(() => presentationElementPosition(title))
    .toEqual(titleBefore);

  await title.dblclick();
  await expect(accent).toHaveAttribute('data-slide-element-selected', 'false');
  await expect(title).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(
    canvas.getByRole('textbox', { name: '幻灯片文本' }),
  ).toBeFocused();
  if (testInfo.project.name === 'desktop-1280') {
    await expect(
      page.getByRole('button', { name: '向右查看更多开始工具' }),
    ).toHaveCount(0);
  }

  await page.keyboard.press('Escape');
  await expect(canvas.getByRole('textbox', { name: '幻灯片文本' })).toHaveCount(
    0,
  );
  await expect(title).toBeFocused();
});

test('presentation groups remain atomic across selection and history', async ({
  page,
}) => {
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const editor = page.locator('.work-presentation-editor');
  const canvas = page.locator('.work-slide-canvas.interactive');
  const elements = canvas.locator(':scope > .work-slide-element');
  const accent = elements.nth(0);
  const title = elements.nth(1);
  const group = page.getByRole('button', { name: '组合', exact: true });
  const ungroup = page.getByRole('button', { name: '取消组合', exact: true });
  const undo = page.getByRole('button', { name: '撤销', exact: true });
  const redo = page.getByRole('button', { name: '重做', exact: true });

  await accent.click();
  await title.click({ modifiers: ['Shift'] });
  await expect(group).toBeEnabled();
  await expect(ungroup).toBeDisabled();
  await group.click();

  const groupedPaths = await elements.evaluateAll((nodes) =>
    nodes
      .slice(0, 2)
      .map((node) => node.getAttribute('data-slide-element-group-path')),
  );
  expect(groupedPaths[0]).toBeTruthy();
  expect(groupedPaths[1]).toBe(groupedPaths[0]);
  const groupedStatus = page.getByText('已选择 1 组，共 2 个对象');
  await expect(groupedStatus).toBeVisible();
  await expect(groupedStatus).toHaveAttribute('aria-live', 'polite');
  await expect(group).toBeDisabled();
  await expect(ungroup).toBeEnabled();

  await undo.click();
  await expect(accent).not.toHaveAttribute('data-slide-element-group-path');
  await expect(title).not.toHaveAttribute('data-slide-element-group-path');
  await expect(redo).toBeEnabled();
  await redo.click();
  await expect(accent).toHaveAttribute(
    'data-slide-element-group-path',
    groupedPaths[0] ?? '',
  );
  await expect(title).toHaveAttribute(
    'data-slide-element-group-path',
    groupedPaths[0] ?? '',
  );

  const resizeHandle = page.getByRole('button', {
    name: '缩放所选组合',
    exact: true,
  });
  await expect(resizeHandle).toBeVisible();
  const accentBeforeResize = await presentationElementGeometry(accent);
  const titleBeforeResize = await presentationElementGeometry(title);
  const titleFontBeforeResize = await presentationElementFontSize(title);
  const resizeHandleBox = await resizeHandle.boundingBox();
  if (!resizeHandleBox) {
    throw new Error('Presentation selection resize handle is unavailable.');
  }
  await page.mouse.move(
    resizeHandleBox.x + resizeHandleBox.width / 2,
    resizeHandleBox.y + resizeHandleBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    resizeHandleBox.x + resizeHandleBox.width / 2 + 36,
    resizeHandleBox.y + resizeHandleBox.height / 2 + 24,
  );
  await page.mouse.up();
  await expect(editor).toHaveAttribute(
    'data-presentation-transform-state',
    'idle',
  );
  const accentAfterResize = await presentationElementGeometry(accent);
  const titleAfterResize = await presentationElementGeometry(title);
  const titleFontAfterResize = await presentationElementFontSize(title);
  expect(accentAfterResize.width).toBeGreaterThan(accentBeforeResize.width);
  expect(accentAfterResize.height).toBeGreaterThan(accentBeforeResize.height);
  expect(titleAfterResize.width).toBeGreaterThan(titleBeforeResize.width);
  expect(titleAfterResize.height).toBeGreaterThan(titleBeforeResize.height);
  expect(titleFontAfterResize).toBeGreaterThan(titleFontBeforeResize);
  expect(titleAfterResize.x - accentAfterResize.x).toBeCloseTo(
    (titleBeforeResize.x - accentBeforeResize.x) *
      (accentAfterResize.width / accentBeforeResize.width),
    3,
  );
  expect(titleAfterResize.y - accentAfterResize.y).toBeCloseTo(
    (titleBeforeResize.y - accentBeforeResize.y) *
      (accentAfterResize.height / accentBeforeResize.height),
    3,
  );
  await undo.click();
  await expect
    .poll(() => presentationElementGeometry(accent))
    .toEqual(accentBeforeResize);
  await expect
    .poll(() => presentationElementGeometry(title))
    .toEqual(titleBeforeResize);
  await expect
    .poll(() => presentationElementFontSize(title))
    .toBe(titleFontBeforeResize);

  await canvas.click({ position: { x: 4, y: 4 } });
  await expect(accent).toHaveAttribute('data-slide-element-selected', 'false');
  await expect(title).toHaveAttribute('data-slide-element-selected', 'false');
  await title.click();
  await expect(accent).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(title).toHaveAttribute('data-slide-element-selected', 'true');

  await title.press('Control+Shift+G');
  await expect(accent).not.toHaveAttribute('data-slide-element-group-path');
  await expect(title).not.toHaveAttribute('data-slide-element-group-path');
  await expect(page.getByText('已选择 2 个对象')).toBeVisible();

  await undo.click();
  await expect(accent).toHaveAttribute(
    'data-slide-element-group-path',
    groupedPaths[0] ?? '',
  );
  await expect(title).toHaveAttribute(
    'data-slide-element-group-path',
    groupedPaths[0] ?? '',
  );
  await expect(page.getByText('已选择 1 组，共 2 个对象')).toBeVisible();
});

test('presentation keeps long-deck thumbnail scenes inside a viewport window', async ({
  page,
}) => {
  const fixture = fixtures.find(
    (candidate) => candidate.kind === 'presentation',
  );
  if (!fixture) throw new Error('Missing presentation visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const duplicate = page.getByRole('button', {
    name: '复制幻灯片',
    exact: true,
  });
  const thumbnails = page.locator('[data-slide-thumbnail]');
  const initialCount = await thumbnails.count();
  for (let index = 0; index < 72; index += 1) {
    await duplicate.click();
  }
  const slideCount = initialCount + 72;
  const thumbnailViewport = page.locator('[data-slide-count]');
  await expect(thumbnailViewport).toHaveAttribute(
    'data-slide-count',
    String(slideCount),
  );
  await expect(thumbnailViewport).toHaveAttribute(
    'data-slide-windowed',
    'true',
  );
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1);
  await expect.poll(() => thumbnails.count()).toBeLessThan(slideCount);
  expect(await thumbnails.count()).toBeLessThanOrEqual(40);
  await thumbnailViewport.evaluate((viewport) => {
    viewport.scrollTop = viewport.scrollHeight / 2;
    viewport.dispatchEvent(new Event('scroll'));
  });
  await expect
    .poll(() =>
      thumbnailViewport.evaluate((viewport) =>
        Number(viewport.getAttribute('data-slide-window-start')),
      ),
    )
    .toBeGreaterThan(0);
  await expect
    .poll(() =>
      thumbnailViewport.evaluate((viewport) =>
        Number(viewport.getAttribute('data-slide-window-end')),
      ),
    )
    .toBeLessThan(slideCount);

  const rendered = page.locator('[data-slide-thumbnail-rendered="true"]');
  await expect.poll(() => rendered.count()).toBeGreaterThan(1);
  await expect.poll(() => rendered.count()).toBeLessThan(slideCount);
  expect(await rendered.count()).toBeLessThanOrEqual(16);

  await thumbnails.first().focus();
  await page.keyboard.press('Home');
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.getAttribute('data-slide-index'),
      ),
    )
    .toBe('0');
  await page.keyboard.press('End');
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.getAttribute('data-slide-index'),
      ),
    )
    .toBe(String(slideCount - 1));
  await expect(page.locator('[data-slide-thumbnail]:focus')).toHaveAttribute(
    'data-slide-thumbnail-rendered',
    'true',
  );

  await page
    .getByRole('button', { name: '幻灯片浏览视图', exact: true })
    .click();
  await expect(page.locator('.work-presentation-sorter')).toBeVisible();
  await page
    .getByRole('slider', { name: '演示缩放', exact: true })
    .press('End');
  await expect(page.getByLabel('演示缩放比例')).toHaveText('200%');
  await expect(page.locator('[data-slide-thumbnail].active')).toHaveAttribute(
    'data-slide-index',
    String(slideCount - 1),
  );
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1);
  await expect.poll(() => thumbnails.count()).toBeLessThan(slideCount);
  expect(await thumbnails.count()).toBeLessThanOrEqual(48);
  await expect.poll(() => rendered.count()).toBeGreaterThan(1);
  await expect.poll(() => rendered.count()).toBeLessThan(slideCount);
  expect(await rendered.count()).toBeLessThanOrEqual(16);
});

test('PDF workspace card uses a single, legible file mark', async ({
  page,
}) => {
  await page.goto('/');
  const card = page.getByRole('button', {
    name: 'PDF 编辑器 查看、批注并保存 PDF',
  });
  const sheet = card.locator('.template-document-sheet');
  const mark = sheet.locator('.template-pdf-mark');

  await expect(card).toBeVisible();
  await expect(mark).toHaveText('PDF');
  await expect(sheet.locator('svg')).toHaveCount(0);
  await expect(mark).toHaveCSS('color', 'rgb(255, 255, 255)');
  await expect(mark).toHaveCSS('background-color', 'rgb(200, 64, 64)');
});

async function presentationElementPosition(
  element: Locator,
): Promise<{ x: number; y: number }> {
  return element.evaluate((node) => ({
    x: Number.parseFloat((node as HTMLElement).style.left),
    y: Number.parseFloat((node as HTMLElement).style.top),
  }));
}

async function presentationElementGeometry(
  element: Locator,
): Promise<{ height: number; width: number; x: number; y: number }> {
  return element.evaluate((node) => {
    const style = (node as HTMLElement).style;
    return {
      height: Number.parseFloat(style.height),
      width: Number.parseFloat(style.width),
      x: Number.parseFloat(style.left),
      y: Number.parseFloat(style.top),
    };
  });
}

async function presentationElementFontSize(element: Locator): Promise<number> {
  return element.evaluate((node) => {
    const text = node.querySelector<HTMLElement>('[style*="font-size"]');
    if (!text) throw new Error('Presentation text style is unavailable.');
    return Number.parseFloat(getComputedStyle(text).fontSize);
  });
}

async function verifySharedEditorGeometry(
  page: Page,
  kind: VisualEditorKind,
  projectName: string,
): Promise<void> {
  const shell = page.locator(`.work-editor-shell.${kind}`);
  const header = shell.locator('.work-editor-header');
  const editorHost = shell.locator('.playground-editor-host');
  await expect(shell).toBeVisible();
  await expect(header).toBeVisible();
  await expect(editorHost).toBeVisible();
  await expect(
    page.getByRole('button', { name: '返回办公首页' }),
  ).toBeVisible();
  await expect(
    page.getByRole('button', { name: '打开 AI 助手' }),
  ).toBeVisible();

  const geometry = await page.evaluate(() => {
    const shell = document.querySelector<HTMLElement>('.work-editor-shell');
    const header = shell?.querySelector<HTMLElement>('.work-editor-header');
    const host = shell?.querySelector<HTMLElement>('.playground-editor-host');
    const commandStart = header?.querySelector<HTMLElement>(
      '.playground-editor-command-start',
    );
    const commandEnd = header?.querySelector<HTMLElement>(
      '.work-editor-header-actions',
    );
    const integratedRow = host?.querySelector<HTMLElement>(
      shell?.classList.contains('pdf')
        ? '.work-pdf-toolbar'
        : '.work-office-ribbon-tabs-row',
    );
    const center = shell?.classList.contains('pdf')
      ? [...(host?.querySelectorAll<HTMLElement>('.work-pdf-toolbar > *') ?? [])]
          .find((element) => getComputedStyle(element).display !== 'none')
      : host?.querySelector<HTMLElement>('.work-office-ribbon-tabs');
    const pdfPageControls = host?.querySelector<HTMLElement>(
      '.work-pdf-page-controls',
    );
    const pdfPageTotal = host?.querySelector<HTMLElement>(
      '.work-pdf-page-total',
    );
    const pdfNextPage = host?.querySelector<HTMLElement>(
      '[aria-label="下一页"]',
    );
    if (
      !shell ||
      !header ||
      !host ||
      !commandStart ||
      !commandEnd ||
      !integratedRow ||
      !center
    ) {
      throw new Error('The shared Office editor shell is incomplete.');
    }
    const shellRect = shell.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const hostRect = host.getBoundingClientRect();
    const startRect = commandStart.getBoundingClientRect();
    const endRect = commandEnd.getBoundingClientRect();
    const rowRect = integratedRow.getBoundingClientRect();
    const centerRect = center.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      documentScrollWidth: document.documentElement.scrollWidth,
      shell: {
        left: shellRect.left,
        top: shellRect.top,
        width: shellRect.width,
        height: shellRect.height,
      },
      header: {
        left: headerRect.left,
        top: headerRect.top,
        width: headerRect.width,
        height: headerRect.height,
        position: getComputedStyle(header).position,
        paddingLeft: Number.parseFloat(getComputedStyle(header).paddingLeft),
      },
      host: {
        left: hostRect.left,
        top: hostRect.top,
        width: hostRect.width,
        height: hostRect.height,
      },
      commandStart: {
        left: startRect.left,
        right: startRect.right,
      },
      commandEnd: {
        left: endRect.left,
        right: endRect.right,
      },
      integratedRow: {
        top: rowRect.top,
        height: rowRect.height,
      },
      center: {
        left: centerRect.left,
        right: centerRect.right,
      },
      pdf:
        pdfPageControls && pdfPageTotal && pdfNextPage
          ? {
              pageControlsRight: pdfPageControls.getBoundingClientRect().right,
              pageTotalWidth: pdfPageTotal.getBoundingClientRect().width,
              nextPageWidth: pdfNextPage.getBoundingClientRect().width,
            }
          : null,
    };
  });

  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.viewportWidth + 1,
  );
  expect(geometry.shell.left).toBeCloseTo(0, 0);
  expect(geometry.shell.top).toBeCloseTo(0, 0);
  expect(geometry.shell.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.header.left).toBeCloseTo(0, 0);
  expect(geometry.header.top).toBeCloseTo(0, 0);
  expect(geometry.header.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.header.height).toBe(kind === 'pdf' ? 42 : 36);
  expect(geometry.header.position).toBe('absolute');
  expect(geometry.host.left).toBeCloseTo(0, 0);
  expect(geometry.host.top).toBeCloseTo(geometry.shell.top, 0);
  expect(geometry.host.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.host.height).toBeCloseTo(geometry.shell.height, 0);
  expect(geometry.integratedRow.top).toBeCloseTo(geometry.header.top, 0);
  expect(geometry.integratedRow.height).toBe(geometry.header.height);
  expect(geometry.commandStart.left).toBeCloseTo(
    geometry.header.left + geometry.header.paddingLeft,
    0,
  );
  expect(geometry.commandStart.right).toBeLessThanOrEqual(
    geometry.center.left + 12,
  );
  expect(geometry.commandEnd.left).toBeGreaterThanOrEqual(
    geometry.center.right - 12,
  );
  expect(geometry.commandEnd.right).toBeCloseTo(geometry.viewportWidth - 8, 0);

  if (projectName === 'compact-768') {
    await expect(
      page.getByRole('button', { name: '展开办公侧边栏' }),
    ).toBeVisible();
  }

  if (kind === 'pdf' && projectName === 'compact-768') {
    expect(geometry.pdf).not.toBeNull();
    if (!geometry.pdf) throw new Error('PDF compact controls are missing.');
    expect(geometry.pdf.pageControlsRight).toBeLessThanOrEqual(
      geometry.commandEnd.left - 2,
    );
    expect(geometry.pdf.pageTotalWidth).toBeGreaterThan(0);
    expect(geometry.pdf.nextPageWidth).toBe(0);
    await expect(page.locator('.work-pdf-page-total')).toBeVisible();
    await expect(page.getByRole('button', { name: '上一页' })).toBeHidden();
    await expect(page.getByRole('button', { name: '下一页' })).toBeHidden();
    await expect(page.locator('.work-pdf-history')).toBeHidden();
    await expect(page.locator('.work-pdf-save')).toBeHidden();
    await expect(page.locator('.work-pdf-zoom-controls')).toBeVisible();
    await expect(page.getByLabel('PDF 缩放比例')).toBeHidden();

    await page.getByRole('button', { name: '更多 PDF 工具' }).click();
    const overflow = page.getByRole('menu', { name: '更多 PDF 工具' });
    for (const action of ['撤销', '上一页', '下一页', '缩小', '放大']) {
      await expect(
        overflow.getByRole('menuitem', { name: action }),
      ).toBeVisible();
    }
    await page.keyboard.press('Escape');
    await expect(overflow).toBeHidden();
    await page.evaluate(() => {
      if (document.activeElement instanceof HTMLElement) {
        document.activeElement.blur();
      }
    });
  }

  if (kind !== 'pdf') {
    const ribbonTabs = page.locator('.work-office-ribbon-tabs');
    await expect(ribbonTabs).toBeVisible();
    const tabInset = await ribbonTabs.evaluate((element) =>
      Number.parseFloat(getComputedStyle(element).paddingLeft),
    );
    expect(tabInset).toBeGreaterThanOrEqual(10);
    const startTab = page.getByRole('tab', { name: '开始' });
    await expect(startTab).toHaveAttribute('aria-selected', 'true');
    await expect(startTab).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
    await expect(
      page.getByRole('button', { name: '文件', exact: true }),
    ).toHaveCount(0);
  }
}

async function openEditorContextMenu(
  page: Page,
  kind: 'document' | 'spreadsheet' | 'presentation',
): Promise<void> {
  if (kind === 'document') {
    await page
      .locator('.work-document-editable .ProseMirror')
      .evaluate((editor) => {
        const text = editor.querySelector('h1')?.firstChild;
        if (!(text instanceof Text)) {
          throw new Error('Document context-menu text is unavailable.');
        }
        const range = document.createRange();
        range.setStart(text, 0);
        range.setEnd(text, Math.min(4, text.length));
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        (editor as HTMLElement).focus();
        document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
        editor.dispatchEvent(
          new MouseEvent('contextmenu', {
            bubbles: true,
            cancelable: true,
            clientX: 420,
            clientY: 520,
          }),
        );
      });
    return;
  }

  const target =
    kind === 'spreadsheet'
      ? page.locator('.work-spreadsheet-canvas')
      : page.locator('.work-slide-canvas.interactive');
  await target.click({
    button: 'right',
    position: kind === 'spreadsheet' ? { x: 280, y: 180 } : undefined,
  });
}

function previewToolbarLabel(kind: VisualEditorKind): string {
  switch (kind) {
    case 'document':
      return '文字预览工具';
    case 'markdown':
      return 'Markdown 预览工具';
    case 'spreadsheet':
      return '表格预览工具';
    case 'presentation':
      return '演示预览工具';
    default:
      throw new Error(`The ${kind} editor does not have a preview toolbar.`);
  }
}
