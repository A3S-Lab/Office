import { expect, test, type Locator, type Page } from '@playwright/test';
import { jsPDF } from 'jspdf';
import {
  openDocumentFixture,
  stabilizeVisualSurface,
  waitForDocumentFixture,
} from './visual-test-support';

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
    open: async (page) => {
      await page
        .locator('input[aria-label="打开 Office 或 PDF 文件"]')
        .setInputFiles({
          name: 'visual-fixture.pdf',
          mimeType: 'application/pdf',
          buffer: visualPdf(),
        });
    },
    ready: async (page) => {
      await page.locator('.work-pdf-embed[data-ready="true"]').waitFor({
        timeout: 30_000,
      });
    },
  },
];

test.describe('Office editor visual contracts', () => {
  test.describe.configure({ mode: 'serial' });

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
  await task.click();
  await expect(source).toHaveValue(/- \[x\] Review the synchronized preview/);

  await source.evaluate((element) => {
    element.scrollTop = (element.scrollHeight - element.clientHeight) * 0.55;
    element.dispatchEvent(new Event('scroll', { bubbles: true }));
  });
  await expect
    .poll(() =>
      page
        .getByRole('region', { name: 'Markdown 编辑结果窗格' })
        .evaluate((element) => element.scrollTop),
    )
    .toBeGreaterThan(0);
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
    if (!shell || !header || !host) {
      throw new Error('The shared Office editor shell is incomplete.');
    }
    const shellRect = shell.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const hostRect = host.getBoundingClientRect();
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
      },
      host: {
        left: hostRect.left,
        top: hostRect.top,
        width: hostRect.width,
        height: hostRect.height,
      },
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
  expect(geometry.header.height).toBeGreaterThanOrEqual(48);
  expect(geometry.header.height).toBeLessThanOrEqual(56);
  expect(geometry.host.left).toBeCloseTo(0, 0);
  expect(geometry.host.top).toBeCloseTo(geometry.header.height, 0);
  expect(geometry.host.width).toBeCloseTo(geometry.viewportWidth, 0);
  expect(geometry.host.height).toBeGreaterThan(600);

  if (projectName === 'compact-768') {
    await expect(
      page.getByRole('button', { name: '展开办公侧边栏' }),
    ).toBeVisible();
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

function visualPdf(): Buffer {
  const pdf = new jsPDF({
    compress: true,
    format: 'a4',
    orientation: 'portrait',
    unit: 'pt',
  });
  pdf.setCreationDate(new Date('2026-01-01T00:00:00.000Z'));
  pdf.setFileId('A3S0FF1CE00000000000000000000001');
  pdf.setProperties({
    author: 'A3S Lab',
    creator: 'A3S Office visual tests',
    subject: 'Deterministic PDF editor fixture',
    title: 'A3S Office',
  });
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(24);
  pdf.text('A3S Office', 72, 96);
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.text('PDF editor visual fixture', 72, 124);
  pdf.setDrawColor(40, 103, 216);
  pdf.setFillColor(238, 244, 255);
  pdf.roundedRect(72, 158, 451, 92, 8, 8, 'FD');
  pdf.setTextColor(34, 52, 82);
  pdf.text(
    'Typed toolbar, PDFium canvas, annotations, search, and save.',
    92,
    194,
  );
  pdf.text(
    'This page is generated in memory by the visual regression test.',
    92,
    218,
  );
  return Buffer.from(pdf.output('arraybuffer'));
}
