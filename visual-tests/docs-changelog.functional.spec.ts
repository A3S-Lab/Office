import { expect, test } from '@playwright/test';

test('documentation changelog stays scannable, localized, and version-aware', async ({
  page,
}, testInfo) => {
  await page.goto('/docs/changelog.html');

  await expect(
    page.getByRole('heading', { level: 1, name: '更新日志' }),
  ).toBeVisible();
  const cards = page.locator('.office-release-card');
  await expect(cards).toHaveCount(156);
  await expect(cards.first()).toHaveAttribute('data-version', '0.175.0');
  await expect(cards.first()).toContainText(
    '文档 PDF 装饰绘制改用 Layout Artifact BDC',
  );
  const releaseCard = (version: string) =>
    page.locator(`.office-release-card[data-version="${version}"]`);
  await expect(releaseCard('0.175.0')).toContainText('Layout Artifact BDC');
  await expect(releaseCard('0.175.0')).toContainText(
    '高亮、下划线与段落边框在 /Artifact << /Type /Layout >> BDC … EMC 内绘制',
  );
  await expect(releaseCard('0.170.0')).toContainText('栖息与飞行剪影');
  await expect(releaseCard('0.170.0')).toContainText(
    'birds 绘制栖息侧影；birdsFlight 绘制展开双翼飞行姿态',
  );
  await expect(releaseCard('0.169.0')).toContainText('翅膀剪影');
  await expect(releaseCard('0.169.0')).toContainText(
    'bats 沿测量边绘制闭合头部 + 双翼折线母题',
  );
  await expect(releaseCard('0.168.0')).toContainText('粗几何轨线');
  await expect(releaseCard('0.168.0')).toContainText(
    'basicWideOutline 绘制双粗轨；basicWideMidline 单粗轨；basicWideInline 粗轨加外侧细伴线',
  );
  await expect(releaseCard('0.167.0')).toContainText('平行细线');
  await expect(releaseCard('0.167.0')).toContainText(
    'basicThinLines 沿测量边绘制三条间距细线',
  );
  await expect(releaseCard('0.166.0')).toContainText('离散短划印章');
  await expect(releaseCard('0.166.0')).toContainText(
    'basicBlackDashes / basicWhiteDashes 沿测量边放置间距短划',
  );
  await expect(releaseCard('0.165.0')).toContainText('离散圆点印章');
  await expect(releaseCard('0.165.0')).toContainText(
    'basicBlackDots / basicWhiteDots 沿测量边放置间距圆点',
  );
  await expect(releaseCard('0.164.0')).toContainText('离散方块印章');
  await expect(releaseCard('0.164.0')).toContainText(
    'basicBlackSquares / basicWhiteSquares 沿测量边放置间距方块',
  );
  await expect(releaseCard('0.163.0')).toContainText('月牙母题');
  await expect(releaseCard('0.163.0')).toContainText(
    'moons 沿测量边绘制外弧 + 内弧闭合折线',
  );
  await expect(releaseCard('0.162.0')).toContainText('矩形母题');
  await expect(releaseCard('0.162.0')).toContainText(
    'marquee 沿测量边绘制矩形',
  );
  await expect(releaseCard('0.161.0')).toContainText('椭圆母题');
  await expect(releaseCard('0.161.0')).toContainText('ovals 沿测量边绘制椭圆');
  await expect(releaseCard('0.160.0')).toContainText('闭合等腰三角');
  await expect(releaseCard('0.160.0')).toContainText(
    'triangles 沿测量边绘制闭合等腰三角',
  );
  await expect(releaseCard('0.159.0')).toContainText('sawtooth 三角齿');
  await expect(releaseCard('0.159.0')).toContainText(
    'sawtooth 沿测量边绘制单侧三角齿',
  );
  await expect(releaseCard('0.158.0')).toContainText(
    'zigZag 与 zigZagStitch 锯齿线',
  );
  await expect(releaseCard('0.158.0')).toContainText(
    '这些艺术样式沿测量边绘制锯齿折线',
  );
  await expect(releaseCard('0.157.0')).toContainText('双色三维浮雕与雕刻');
  await expect(releaseCard('0.157.0')).toContainText(
    'threeDEmboss 与 threeDEngrave 沿测量边绘制高光与阴影偏移',
  );
  await expect(releaseCard('0.156.0')).toContainText('显式波形折线');
  await expect(releaseCard('0.156.0')).toContainText(
    'wave 边框沿测量边绘制正弦折线',
  );
  await expect(releaseCard('0.155.0')).toContainText(
    '矢量 ActualText Span 上的 MCID',
  );
  await expect(releaseCard('0.155.0')).toContainText(
    '每次矢量文本运行发出带 /ActualText 与页内 /MCID 的 /Span',
  );
  await expect(releaseCard('0.154.0')).toContainText(
    '目录链接的 StructTreeRoot 桩',
  );
  await expect(releaseCard('0.154.0')).toContainText(
    '导出的 PDF 发出带 Document 父节点以及由大纲派生的标题或段落子节点的 StructTreeRoot',
  );
  await expect(releaseCard('0.153.0')).toContainText(
    '目录 MarkInfo 已标记标志',
  );
  await expect(releaseCard('0.153.0')).toContainText(
    '导出的 PDF 在目录上设置 /MarkInfo << /Marked true >>，供带标签 PDF 消费方使用',
  );
  await expect(releaseCard('0.152.0')).toContainText(
    'PDF 导出中的 between 与 bar 边',
  );
  await expect(releaseCard('0.152.0')).toContainText(
    '测量得到的段落框以与上/左/下/右相同的常见样式描边 between 与 bar 边',
  );
  await expect(releaseCard('0.151.0')).toContainText(
    'PDF 导出中的段落边框矢量描边',
  );
  await expect(releaseCard('0.151.0')).toContainText(
    '测量得到的段落框在与矢量文本层相同的页面几何上描边上、左、下、右 PDF 路径',
  );
  await expect(releaseCard('0.150.0')).toContainText(
    'PDF 导出中的突出显示矢量填充',
  );
  await expect(releaseCard('0.150.0')).toContainText(
    '带突出显示的测量文本在与矢量文本层相同的页面几何上填充 PDF 矩形',
  );
  await expect(releaseCard('0.149.0')).toContainText(
    'PDF 导出中的下划线矢量路径',
  );
  await expect(releaseCard('0.149.0')).toContainText(
    '带下划线的测量文本在与矢量文本层相同的页面几何上描边 single、double 或 thick PDF 路径',
  );
  await expect(releaseCard('0.148.0')).toContainText(
    'PDF 书签中的大纲级别段落',
  );
  await expect(releaseCard('0.148.0')).toContainText(
    'Writer 大纲级别段落（p[data-office-outline-level]）与 h1–h6 一并进入导出的 PDF 大纲',
  );
  await expect(releaseCard('0.147.0')).toContainText(
    '修订正文中的空分隔符字形',
  );
  await expect(releaseCard('0.147.0')).toContainText(
    '无属性空 w:separator 与 w:continuationSeparator 字形可进入整段段落标记、段落分隔与文字移动修订正文',
  );
  await expect(releaseCard('0.146.0')).toContainText(
    '修订正文中的空 annotationRef',
  );
  await expect(releaseCard('0.146.0')).toContainText(
    '无属性空 w:annotationRef 字形可进入整段段落标记、段落分隔与文字移动修订正文',
  );
  await expect(releaseCard('0.145.0')).toContainText(
    '修订正文中的空 endnoteRef',
  );
  await expect(releaseCard('0.145.0')).toContainText(
    '无属性空 w:endnoteRef 字形可进入整段段落标记、段落分隔与文字移动修订正文',
  );
  await expect(releaseCard('0.144.0')).toContainText(
    '修订正文中的空 footnoteRef',
  );
  await expect(releaseCard('0.144.0')).toContainText(
    '无属性空 w:footnoteRef 字形可进入整段段落标记、段落分隔与文字移动修订正文',
  );
  await expect(releaseCard('0.143.0')).toContainText(
    '省略的 w:ilvl 默认为级别 0',
  );
  await expect(releaseCard('0.143.0')).toContainText(
    '无显式 w:ilvl 的 CT_NumPr 导入为可审阅的 OOXML 默认级别编号修订',
  );
  await expect(releaseCard('0.142.0')).toContainText('仅属性浮动先验');
  await expect(releaseCard('0.142.0')).toContainText(
    '已知锚点/规格、有界 twips 与可选 FromText 距离作为可审阅 table-formatting 浮动快照往返',
  );
  await expect(releaseCard('0.141.0')).toContainText(
    'w:original 中的不透明兄弟格式',
  );
  await expect(releaseCard('0.141.0')).toContainText(
    '兄弟 %[ilvl]:[start]:[nfc]:[suff] 段可将非常见 ST_NumberFormat 值作为不透明先验文本往返',
  );
  await expect(releaseCard('0.140.0')).toContainText('一层嵌套表 companion');
  await expect(releaseCard('0.140.0')).toContainText(
    '正文级 move*Range* Start/End 夹住外层表，且内层嵌套表含受支持移动时',
  );
  await expect(releaseCard('0.139.0')).toContainText('简单 SDT companion');
  await expect(releaseCard('0.139.0')).toContainText(
    '正文级 move*Range* Start/End 夹住一个简单 w:sdt 时',
  );
  await expect(releaseCard('0.138.0')).toContainText('多单元格表格 companion');
  await expect(releaseCard('0.138.0')).toContainText(
    '正文级 move*Range* Start/End 夹住多单元格表格时',
  );
  await expect(releaseCard('0.137.0')).toContainText(
    '导出时生成 companion 书签',
  );
  await expect(releaseCard('0.137.0')).toContainText(
    '准入的 Compare / 同文档纯文本推断移动会随移动包装一并导出',
  );
  await expect(releaseCard('0.136.0')).toContainText('单单元格表格 companion');
  await expect(releaseCard('0.136.0')).toContainText(
    '正文级 move*Range* Start/End 夹住单单元格表格时',
  );
  await expect(releaseCard('0.135.0')).toContainText('仅含图片正文准入');
  await expect(releaseCard('0.135.0')).toContainText(
    '仅含受支持的 wp:inline 图片嵌入即可满足整段标记与段落分隔符正文准入',
  );
  await expect(releaseCard('0.134.0')).toContainText('未跟踪行内图片兄弟');
  await expect(releaseCard('0.134.0')).toContainText(
    '已解析的 wp:inline 图片嵌入可位于匹配的标记包装旁',
  );
  await expect(releaseCard('0.133.0')).toContainText('行内 DrawingML 图片准入');
  await expect(releaseCard('0.133.0')).toContainText(
    '已解析的 wp:inline 图片嵌入',
  );
  await expect(releaseCard('0.132.0')).toContainText('安全外部超链接准入');
  await expect(releaseCard('0.132.0')).toContainText(
    '已解析的 http/https/mailto',
  );
  await expect(releaseCard('0.131.0')).toContainText('跨分节 companion 书签');
  await expect(releaseCard('0.131.0')).toContainText('未配对标记仍诊断');
  await expect(releaseCard('0.130.0')).toContainText('未跟踪纯文本兄弟 run');
  await expect(releaseCard('0.130.0')).toContainText(
    '绘图与关系绑定链接保持失败闭合',
  );
  await expect(releaseCard('0.129.0')).toContainText(
    '修订正文中的空短日期字段',
  );
  await expect(releaseCard('0.129.0')).toContainText(
    '带属性短日期字段保持失败闭合',
  );
  await expect(releaseCard('0.128.0')).toContainText(
    '修订正文中的空页码与日期字段',
  );
  await expect(releaseCard('0.128.0')).toContainText(
    '带属性页码与日期字段保持失败闭合',
  );
  await expect(releaseCard('0.127.0')).toContainText(
    '修订正文中的空最后渲染分页符',
  );
  await expect(releaseCard('0.127.0')).toContainText(
    '带属性最后渲染分页符保持失败闭合',
  );
  await expect(releaseCard('0.126.0')).toContainText('修订正文中的空回车符');
  await expect(releaseCard('0.126.0')).toContainText(
    '带属性回车符保持失败闭合',
  );
  await expect(releaseCard('0.125.0')).toContainText(
    '修订正文中的制表符与连字符',
  );
  await expect(releaseCard('0.125.0')).toContainText('带属性字形保持失败闭合');
  await expect(releaseCard('0.124.0')).toContainText('更丰富的段落分隔正文');
  await expect(releaseCard('0.124.0')).toContainText('与段落标记准入对齐');
  await expect(releaseCard('0.123.0')).toContainText('多级编号先验');
  await expect(releaseCard('0.123.0')).toContainText('原子接受/拒绝');
  await expect(releaseCard('0.122.0')).toContainText('文档网格字符间距');
  await expect(releaseCard('0.122.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.121.0')).toContainText('章节页码字段');
  await expect(releaseCard('0.121.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.120.0')).toContainText('节页边框');
  await expect(releaseCard('0.120.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.119.0')).toContainText('节分隔类型');
  await expect(releaseCard('0.119.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.118.0')).toContainText('节尾注属性');
  await expect(releaseCard('0.118.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.117.0')).toContainText('节脚注属性');
  await expect(releaseCard('0.117.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.116.0')).toContainText('节双向布局');
  await expect(releaseCard('0.116.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.115.0')).toContainText('节文字方向');
  await expect(releaseCard('0.115.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.114.0')).toContainText('节尾注抑制');
  await expect(releaseCard('0.114.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.113.0')).toContainText('节垂直对齐');
  await expect(releaseCard('0.113.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.112.0')).toContainText('节表单保护');
  await expect(releaseCard('0.112.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.111.0')).toContainText('节页码');
  await expect(releaseCard('0.111.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.110.0')).toContainText('节行号');
  await expect(releaseCard('0.110.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.109.0')).toContainText('行单元格间距');
  await expect(releaseCard('0.109.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.108.0')).toContainText('行分区 ID');
  await expect(releaseCard('0.107.0')).toContainText('网格跨度');
  await expect(releaseCard('0.106.0')).toContainText('垂直合并');
  await expect(releaseCard('0.105.0')).toContainText('水平合并');
  await expect(releaseCard('0.104.0')).toContainText('文档网格');
  await expect(releaseCard('0.103.0')).toContainText('单元格边框');
  await expect(releaseCard('0.102.0')).toContainText('单元格条件格式位掩码');
  await expect(releaseCard('0.101.0')).toContainText('行条件格式位掩码');
  await expect(releaseCard('0.100.0')).toContainText('不等宽分栏');
  await expect(releaseCard('0.99.0')).toContainText('行带大小');
  await expect(releaseCard('0.98.0')).toContainText('列带大小');
  await expect(releaseCard('0.97.0')).toContainText('表描述');
  await expect(releaseCard('0.96.0')).toContainText('表题注');
  await expect(releaseCard('0.95.0')).toContainText('表边框');
  await expect(releaseCard('0.93.0')).toContainText('表样式 ID');
  await expect(releaseCard('0.93.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.92.0')).toContainText('表重叠策略');
  await expect(releaseCard('0.92.0')).toContainText('实时修订跟踪');
  await expect(releaseCard('0.92.0')).toContainText('不透明路径保持失败闭合');
  await expect(releaseCard('0.78.0')).toContainText('可审阅的属性修订');
  await expect(releaseCard('0.78.0')).toContainText('段落标记与移动保真');
  await expect(releaseCard('0.78.0')).toContainText('可搜索 PDF 矢量文本');
  await expect(releaseCard('0.77.0')).toContainText('Writer 段落与域快捷键');
  await expect(releaseCard('0.77.0')).toContainText(
    'Presentation 与 Spreadsheet 界面',
  );
  await expect(releaseCard('0.77.0')).toContainText('Markdown 与 PDF 键盘证据');
  await expect(releaseCard('0.60.0')).toContainText(
    'WPS UI 参考升级为类型化证据',
  );
  await expect(releaseCard('0.60.0')).toContainText('三个明确 profile');
  await expect(releaseCard('0.60.0')).toContainText('COM 生命周期可控');
  await expect(releaseCard('0.60.0')).toContainText('证据边界保持诚实');
  await expect(releaseCard('0.59.0')).toContainText(
    'Writer 新增类型化字段设置',
  );
  await expect(releaseCard('0.59.0')).toContainText('一条类型化插入/编辑路径');
  await expect(releaseCard('0.59.0')).toContainText('原生 WPS 格式保持诚实');
  await expect(releaseCard('0.58.0')).toContainText(
    'Writer 字段补齐有界 WPS 数字开关对齐',
  );
  await expect(releaseCard('0.58.0')).toContainText('开关保持类型化且有界');
  await expect(releaseCard('0.58.0')).toContainText('WPS COM 输出可检查');
  await expect(releaseCard('0.56.0')).toContainText(
    'Writer 连接符补齐 WPS 箭头样式对齐',
  );
  await expect(releaseCard('0.56.0')).toContainText('一个类型化箭头样式模型');
  await expect(releaseCard('0.56.0')).toContainText('Windows COM 3/4 参考');
  await expect(releaseCard('0.55.0')).toContainText(
    'Writer 连接符补齐 WPS 线型对齐',
  );
  await expect(releaseCard('0.55.0')).toContainText('一个类型化线型模型');
  await expect(releaseCard('0.55.0')).toContainText('记录 WPS COM 证据');
  await expect(releaseCard('0.54.1')).toContainText(
    'Windows 编辑器运行保持确定性',
  );
  await expect(releaseCard('0.54.1')).toContainText('直接 CDP 生命周期');
  await expect(releaseCard('0.54.0')).toContainText(
    'A3S Test 成为编辑器交互主契约',
  );
  await expect(releaseCard('0.54.0')).toContainText('声明式五编辑器 CLI');
  await expect(releaseCard('0.53.1')).toContainText(
    'Writer 明确记录 WPS 连接符边界',
  );
  await expect(releaseCard('0.53.1')).toContainText('COM 证据决定边界');
  await expect(releaseCard('0.53.0')).toContainText(
    'Writer 文本框与 WPS 使用同一套有界形状语义',
  );
  await expect(releaseCard('0.53.0')).toContainText('五种形状，一个类型化状态');
  await expect(releaseCard('0.52.0')).toContainText(
    'Writer 整段修订保持原生与原子语义',
  );
  await expect(releaseCard('0.52.0')).toContainText('一个段落，一项决定');
  await expect(releaseCard('0.51.0')).toContainText(
    'Writer 比较配对同一分节内文字移动',
  );
  await expect(releaseCard('0.51.0')).toContainText('段落范围保持完整');
  await expect(releaseCard('0.50.0')).toContainText(
    'Writer 比较识别有界文字移动',
  );
  await expect(releaseCard('0.49.0')).toContainText(
    'Writer 移动修订保持成对并原生往返',
  );
  await expect(releaseCard('0.48.1')).toContainText(
    'Writer 选区控件恢复紧凑视觉契约',
  );
  await expect(releaseCard('0.48.1')).toContainText('浏览器原生默认样式');
  await expect(
    page.locator('.office-release-card[data-version="0.41.0"]'),
  ).toContainText('表格数据验证警告现在与 Office 决策一致');
  await expect(
    page.locator('.office-release-card[data-version="0.34.0"]'),
  ).toContainText('演示文稿入场动画');
  await expect(releaseCard('0.56.0').locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-06',
  );
  await expect(releaseCard('0.58.0').locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-06',
  );
  await expect(releaseCard('0.59.0').locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-06',
  );
  await expect(releaseCard('0.60.0').locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-06',
  );
  await expect(
    releaseCard('0.60.0').locator('.office-release-card__highlights > li'),
  ).toHaveCount(3);
  await expect(
    releaseCard('0.59.0').locator('.office-release-card__highlights > li'),
  ).toHaveCount(3);
  await expect(
    releaseCard('0.58.0').locator('.office-release-card__highlights > li'),
  ).toHaveCount(3);
  await expect(
    releaseCard('0.56.0').locator('.office-release-card__highlights > li'),
  ).toHaveCount(3);

  const geometry = await page.evaluate(() => {
    const highlights = document.querySelector<HTMLElement>(
      '.office-release-card__highlights',
    );
    if (!highlights) throw new Error('Release highlights are missing.');
    return {
      documentClientWidth: document.documentElement.clientWidth,
      documentScrollWidth: document.documentElement.scrollWidth,
      highlightColumns:
        getComputedStyle(highlights).gridTemplateColumns.split(' ').length,
    };
  });
  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.documentClientWidth + 1,
  );
  expect(geometry.highlightColumns).toBe(
    testInfo.project.name === 'compact-768' ? 1 : 3,
  );

  await page.screenshot({
    fullPage: true,
    path: testInfo.outputPath('docs-changelog.png'),
  });

  await page.getByRole('link', { name: '阅读移动修订指南' }).click();
  await expect(page.locator('h2#移动修订')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读协作合同' }).click();
  await expect(page.locator('h3#同步移动修订')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读选区工具栏说明' }).click();
  await expect(page.locator('h2#选择工具栏控件')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读内容控件指南' }).click();
  await expect(page.locator('h2#原生内容控件')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读公式条件格式指南' }).click();
  await expect(page.locator('h2#公式条件格式')).toBeInViewport();

  await page.goto('/docs/0.38.0/changelog.html');
  await expect(page.locator('.office-release-card')).toHaveCount(14);
  await expect(page.locator('.office-release-card').first()).toContainText(
    'v0.38.0',
  );
  await expect(page.getByText('v0.38.1', { exact: true })).toHaveCount(0);

  await page.goto('/docs/en/changelog.html');
  await expect(
    page.getByRole('heading', { level: 1, name: "What's new" }),
  ).toBeVisible();
  await expect(
    page.locator('.office-release-card[data-version="0.43.0"]'),
  ).toContainText('Spreadsheet conditional formatting is now formula-editable');
  await expect(
    page.locator('.office-release-card[data-version="0.42.0"]'),
  ).toContainText('Spreadsheet rules can now be local custom formulas');
});
