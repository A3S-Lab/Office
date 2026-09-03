import type {
  ContextDiffRequest,
  ContextSnapshotRequest,
  PageContextBridge,
  PageContextSnapshot,
  TestKitHandshake,
} from '@a3s-lab/testkit';
import { expect, type Page, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';

const pageContextSymbol = 'a3s.test.page-context';
const snapshotLimits = {
  nodes: 220,
  uiNodes: 220,
  uiStateSamples: 160,
  uiDurationMs: 32,
  uiEncodedBytes: 450_000,
};

test('Test Kit handshake publishes bounded home UI evidence', async ({
  page,
}, testInfo) => {
  await page.goto('/playground/');
  await expect(
    page.getByRole('heading', { name: '我的文档', level: 1 }),
  ).toBeVisible();

  const handshake = await readHandshake(page);
  expect(handshake.protocol).toBe('a3s.test.testkit-handshake/1');
  expect(handshake.packageName).toBe('@a3s-lab/testkit');
  expect(handshake.sdkVersion).toBe('0.6.2');
  expect(handshake.pageContextProtocol).toBe('a3s.test.page-context/1');
  expect(handshake.capabilities).toEqual(
    expect.arrayContaining([
      'component_boundaries',
      'ui_layout_graph',
      'ui_state_diffs',
      'ui_style_profile',
      'ui_motion_profile',
    ]),
  );

  const snapshot = await readSnapshot(page);
  expect(snapshot.page.id).toBe('office-playground');
  expect(snapshot.page.ready).toBe(true);
  expect(snapshot.page.route).toBe('/playground/');
  expect(snapshot.components.map((item) => item.id)).toContain(
    'office-playground',
  );
  assertUiEvidence(snapshot);

  await page.screenshot({
    path: testInfo.outputPath('home-testkit-ui.png'),
    fullPage: false,
  });
});

test('Document Test Kit evidence follows toolbar state and ownership', async ({
  page,
}, testInfo) => {
  await openTemplate(page, 'project-brief', '.work-document-editable');
  const baseline = await readSnapshotForComponent(
    page,
    'office-editor-document',
  );
  const documentComponent = requireComponent(
    baseline,
    'office-editor-document',
  );
  expect(documentComponent.facts.editorKind).toBe('document');
  expect(documentComponent.facts.preview).toBe(false);

  const layoutTab = page.getByRole('tab', { name: '页面布局' });
  await layoutTab.click();
  await expect(layoutTab).toHaveAttribute('aria-selected', 'true');
  const pageColor = page.getByRole('button', { name: '页面颜色' });
  await pageColor.click();
  const colorDialog = page.getByRole('dialog', { name: '页面颜色' });
  await expect(colorDialog).toBeVisible();

  const changed = await waitForDiff(page, baseline.revision);
  expect(changed).not.toBeNull();
  if (!changed) throw new Error('Document Test Kit diff was not returned.');
  expect(changed.revision).toBeGreaterThan(baseline.revision);
  expect(['complete', 'reset_required']).toContain(changed.delta?.status);
  const changedEvidence = await readSnapshotForComponent(
    page,
    'office-editor-document',
  );
  requireComponent(changedEvidence, 'office-editor-document');
  assertUiEvidence(changedEvidence);
  expect(changedEvidence.ui?.stateDiffs.length ?? 0).toBeGreaterThan(0);
  expect(
    changedEvidence.ui?.stateDiffs.some(
      (item) =>
        item.to === 'expanded' || item.to === 'selected' || item.to === 'focus',
    ),
  ).toBe(true);

  const dialogBox = await colorDialog.boundingBox();
  expect(dialogBox).not.toBeNull();
  const viewport = page.viewportSize();
  expect(viewport).not.toBeNull();
  expect((dialogBox?.x ?? 0) + (dialogBox?.width ?? 0)).toBeLessThanOrEqual(
    (viewport?.width ?? 0) + 1,
  );
  expect((dialogBox?.y ?? 0) + (dialogBox?.height ?? 0)).toBeLessThanOrEqual(
    (viewport?.height ?? 0) + 1,
  );
  await page.screenshot({
    path: testInfo.outputPath('document-page-color-testkit-ui.png'),
    fullPage: false,
  });

  await page.keyboard.press('Escape');
  await expect(colorDialog).toBeHidden();
});

test('Spreadsheet Test Kit evidence captures ribbon and grid interaction', async ({
  page,
}, testInfo) => {
  await openTemplate(page, 'quarterly-plan', '.work-spreadsheet-editor');
  const grid = page.locator('.fortune-sheet-overlay');
  await expect(grid).toBeVisible();
  const baseline = await readSnapshotForComponent(
    page,
    'office-editor-spreadsheet',
  );
  const spreadsheetComponent = requireComponent(
    baseline,
    'office-editor-spreadsheet',
  );
  expect(spreadsheetComponent.facts.editorKind).toBe('spreadsheet');

  const dataTab = page.getByRole('tab', { name: '数据' });
  await dataTab.click();
  await expect(dataTab).toHaveAttribute('aria-selected', 'true');
  await expect(
    page.locator('.work-spreadsheet-ribbon .work-office-ribbon-panel'),
  ).toBeVisible();
  await grid.focus();
  await expect(grid).toBeFocused();

  const changed = await waitForDiff(page, baseline.revision);
  expect(changed).not.toBeNull();
  if (!changed) throw new Error('Spreadsheet Test Kit diff was not returned.');
  expect(changed.revision).toBeGreaterThan(baseline.revision);
  const changedEvidence = await readSnapshotForComponent(
    page,
    'office-editor-spreadsheet',
  );
  assertUiEvidence(changedEvidence);
  expect(changedEvidence.ui?.stateDiffs.length ?? 0).toBeGreaterThan(0);

  const ribbonBox = await page
    .locator('.work-spreadsheet-ribbon')
    .boundingBox();
  expect(ribbonBox).not.toBeNull();
  const viewport = page.viewportSize();
  expect(viewport).not.toBeNull();
  expect((ribbonBox?.x ?? 0) + (ribbonBox?.width ?? 0)).toBeLessThanOrEqual(
    (viewport?.width ?? 0) + 1,
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-ribbon-testkit-ui.png'),
    fullPage: false,
  });
});

test('Presentation Test Kit evidence tracks interaction and responsive layout', async ({
  page,
}, testInfo) => {
  await openTemplate(page, 'strategy-deck', '.work-presentation-editor');
  const baseline = await readSnapshotForComponent(
    page,
    'office-editor-presentation',
  );
  requireComponent(baseline, 'office-editor-presentation');

  const designTab = page.getByRole('tab', { name: '设计', exact: true });
  await designTab.click();
  await expect(designTab).toHaveAttribute('aria-selected', 'true');
  const slide = page.locator('[data-slide-thumbnail]').first();
  await expect(slide).toBeVisible();
  await slide.click();
  await expect(slide).toBeFocused();
  const changed = await waitForDiff(page, baseline.revision);
  expect(changed).not.toBeNull();
  if (!changed) throw new Error('Presentation Test Kit diff was not returned.');
  const changedEvidence = await readSnapshotForComponent(
    page,
    'office-editor-presentation',
  );
  assertUiEvidence(changedEvidence);
  expect(changedEvidence.ui?.stateDiffs.length ?? 0).toBeGreaterThan(0);

  await page.setViewportSize({ width: 390, height: 844 });
  await expect(page.locator('.work-presentation-editor')).toBeVisible();
  const compact = await readSnapshot(page);
  requireComponent(compact, 'office-editor-presentation');
  expect(compact.page.viewport.width).toBe(390);
  expect(compact.page.viewport.height).toBe(844);
  assertUiEvidence(compact);
  expect(
    compact.ui?.style.responsiveConditions.some((item) => item.matches),
  ).toBe(true);
  await page.screenshot({
    path: testInfo.outputPath('presentation-responsive-testkit-ui.png'),
    fullPage: false,
  });
});

test('Markdown Test Kit evidence follows source and visual mode changes', async ({
  page,
}, testInfo) => {
  await openTemplate(
    page,
    'blank-markdown',
    'textarea[aria-label="Markdown 源码"]',
  );
  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await expect(source).toBeVisible();
  const baseline = await readSnapshotForComponent(
    page,
    'office-editor-markdown',
  );
  requireComponent(baseline, 'office-editor-markdown');

  await page.getByRole('tab', { name: '视图' }).click();
  const visualMode = page
    .getByRole('region', { name: '编辑方式' })
    .getByRole('button', { name: '可视化编辑' });
  await visualMode.click();
  await expect(visualMode).toHaveAttribute('aria-pressed', 'true');
  await expect(
    page.getByRole('textbox', { name: 'Markdown 编辑区' }),
  ).toBeVisible();

  const changed = await waitForDiff(page, baseline.revision);
  expect(changed).not.toBeNull();
  if (!changed) throw new Error('Markdown Test Kit diff was not returned.');
  const changedEvidence = await readSnapshotForComponent(
    page,
    'office-editor-markdown',
  );
  assertUiEvidence(changedEvidence);
  expect(changedEvidence.ui?.stateDiffs.length ?? 0).toBeGreaterThan(0);
  await page.screenshot({
    path: testInfo.outputPath('markdown-visual-mode-testkit-ui.png'),
    fullPage: false,
  });
});

test('PDF Test Kit evidence captures page organizer interaction', async ({
  page,
}, testInfo) => {
  await page.goto('/playground/');
  await openPdfFixture(page, { pageCount: 2 });
  await waitForPdfFixture(page);
  const baseline = await readSnapshotForComponent(page, 'office-editor-pdf');
  requireComponent(baseline, 'office-editor-pdf');

  const organizerTrigger = page.getByRole('button', { name: '组织 PDF 页面' });
  if (await organizerTrigger.isVisible()) {
    await organizerTrigger.click();
  } else {
    await page.getByRole('button', { name: '更多 PDF 工具' }).click();
    await page
      .getByRole('menu', { name: '更多 PDF 工具' })
      .getByRole('menuitem', { name: '组织页面' })
      .click();
  }
  const organizer = page.getByRole('dialog', { name: '组织 PDF 页面' });
  await expect(organizer).toBeVisible();
  const secondPage = organizer.getByRole('button', { name: '选择第 2 页' });
  await secondPage.click();
  await expect(secondPage).toHaveAttribute('aria-pressed', 'true');
  const changed = await waitForDiff(page, baseline.revision);
  expect(changed).not.toBeNull();
  if (!changed) throw new Error('PDF Test Kit diff was not returned.');
  const changedEvidence = await readSnapshotForComponent(
    page,
    'office-editor-pdf',
  );
  assertUiEvidence(changedEvidence);
  expect(changedEvidence.ui?.stateDiffs.length ?? 0).toBeGreaterThan(0);
  await page.screenshot({
    path: testInfo.outputPath('pdf-organizer-testkit-ui.png'),
    fullPage: false,
  });
  await organizer.getByRole('button', { name: '完成' }).click();
  await expect(organizer).toBeHidden();
});

async function openTemplate(
  page: Page,
  templateId: string,
  readySelector: string,
): Promise<void> {
  await page.goto('/playground/');
  const card = page.locator(`button[data-template-id="${templateId}"]`);
  const editorKind =
    templateId === 'project-brief'
      ? 'document'
      : templateId === 'quarterly-plan'
        ? 'spreadsheet'
        : templateId === 'strategy-deck'
          ? 'presentation'
          : 'markdown';
  const preloadMarker = `a3s-office.playground.${editorKind}.module-preload-ready`;
  let preloaded = false;
  for (let attempt = 0; attempt < 3 && !preloaded; attempt += 1) {
    await card.waitFor({ state: 'visible' });
    await card.hover();
    try {
      await page.waitForFunction(
        (entryName) => performance.getEntriesByName(entryName).length > 0,
        preloadMarker,
        { timeout: 20_000 },
      );
      preloaded = true;
    } catch {
      // The first on-demand dev compilation may trigger an HMR reload. Return
      // to the stable home route and warm the same card again.
      await page.goto('/playground/', { waitUntil: 'domcontentloaded' });
    }
  }
  if (!preloaded) {
    throw new Error(`Timed out preloading Playground editor: ${editorKind}`);
  }
  await card.click();
  await page.locator(readySelector).waitFor({ state: 'visible' });
}

async function readHandshake(page: Page): Promise<TestKitHandshake> {
  return page.evaluate((symbolName) => {
    const bridge = (
      window as unknown as { [key: symbol]: PageContextBridge | undefined }
    )[Symbol.for(symbolName)];
    if (!bridge)
      throw new Error('A3S Test Kit page context bridge is missing.');
    return bridge.handshake();
  }, pageContextSymbol);
}

async function readSnapshot(
  page: Page,
  request: ContextSnapshotRequest = {
    detail: 'forensic',
    ui: true,
    limits: snapshotLimits,
  },
): Promise<PageContextSnapshot> {
  return page.evaluate(
    ({ symbolName, snapshotRequest }) => {
      const bridge = (
        window as unknown as { [key: symbol]: PageContextBridge | undefined }
      )[Symbol.for(symbolName)];
      if (!bridge) {
        throw new Error('A3S Test Kit page context bridge is missing.');
      }
      return bridge.snapshot(snapshotRequest);
    },
    { symbolName: pageContextSymbol, snapshotRequest: request },
  );
}

async function waitForDiff(
  page: Page,
  sinceRevision: number,
): Promise<PageContextSnapshot | null> {
  const request: ContextDiffRequest = {
    sinceRevision,
    timeoutMs: 5_000,
    ui: true,
    limits: snapshotLimits,
  };
  return page.evaluate(
    ({ symbolName, diffRequest }) => {
      const bridge = (
        window as unknown as { [key: symbol]: PageContextBridge | undefined }
      )[Symbol.for(symbolName)];
      if (!bridge) {
        throw new Error('A3S Test Kit page context bridge is missing.');
      }
      return bridge.waitForDiff(diffRequest);
    },
    { symbolName: pageContextSymbol, diffRequest: request },
  );
}

async function readSnapshotForComponent(
  page: Page,
  componentId: string,
): Promise<PageContextSnapshot> {
  let latest: PageContextSnapshot | undefined;
  await expect
    .poll(
      async () => {
        latest = await readSnapshot(page);
        return latest.components.some((item) => item.id === componentId);
      },
      { timeout: 30_000 },
    )
    .toBe(true);
  if (!latest) {
    throw new Error(`Unable to capture Test Kit snapshot: ${componentId}`);
  }
  return latest;
}

function requireComponent(
  snapshot: PageContextSnapshot,
  componentId: string,
): PageContextSnapshot['components'][number] {
  const component = snapshot.components.find((item) => item.id === componentId);
  expect(
    component,
    `Missing Test Kit component boundary: ${componentId}`,
  ).toBeDefined();
  if (!component) throw new Error(`Missing component boundary: ${componentId}`);
  return component;
}

function assertUiEvidence(snapshot: PageContextSnapshot): void {
  expect(snapshot.ui?.protocol).toBe('a3s.test.ui-understanding/1');
  expect(snapshot.ui?.evidence.totalCandidateNodes ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.layout.nodes.length ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.layout.edges.length ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.style.colors.length ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.style.typography.length ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.style.spacing.length ?? 0).toBeGreaterThan(0);
  expect(snapshot.ui?.motion.transitions.length ?? 0).toBeGreaterThan(0);
}
