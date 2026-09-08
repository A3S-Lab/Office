<p align="center">
  <img src="./assets/readme/hero.svg" width="100%" alt="A3S Office brings five format-native editors into a host-owned product boundary">
</p>

<p align="center">
  <strong>Language / 语言:</strong>
  <a href="README.md">English</a> ·
  <a href="README.zh-CN.md">中文</a>
</p>

<p align="center">
  <strong>开源浏览器编辑器和真实 Office 文件的确定性本机自动化。</strong>
</p>

<p align="center">
  文档·电子表格·演示文稿·Markdown·PDF
</p>

<p align="center">
  <a href="https://a3s-lab.github.io/Office/playground/"><strong>打开Playground</strong></a>
  ·
  <a href="https://a3s-lab.github.io/Office/docs/">阅读文档</a>
  ·
  <a href="#quick-start">嵌入编辑器</a>
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI status" src="https://img.shields.io/github/actions/workflow/status/A3S-Lab/Office/ci.yml?branch=main&amp;style=flat-square&amp;label=CI"></a>
  <a href="https://www.npmjs.com/package/@a3s-lab/office"><img alt="npm version" src="https://img.shields.io/npm/v/@a3s-lab/office?style=flat-square&amp;color=1456f0"></a>
  <a href="#project-status"><img alt="Project status: pre-1.0" src="https://img.shields.io/badge/status-pre--1.0-7457c8?style=flat-square"></a>
  <a href="./LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/license-MIT-16845b?style=flat-square"></a>
</p>

<p align="center">
  <a href="#proof-not-promises">产品</a> ·
  <a href="#the-core-design">设计</a> ·
  <a href="#quick-start">快速启动</a> ·
  <a href="#collaboration-without-a-bundled-cloud">协作</a> ·
  <a href="#automation-without-ui-scraping">自动化</a> ·
  <a href="#capabilities-and-boundaries">边界</a> ·
  <a href="./ROADMAP.md">路线图</a> ·
  <a href="./CONTRIBUTING.md">贡献</a>
</p>

---

A3S Office 是一个开源 Office 引擎，供团队将编辑构建为
他们自己的产品。它为文档提供完整的浏览器界面，
电子表格、演示文稿、Markdown 和 PDF，以及单独的 Rust
文件和编码代理的自动化平面。

产品边界经过深思熟虑：A3S Office 拥有格式感知编辑功能，
布局、导入、导出和类型化突变。您的应用程序可以控制
内容、持久性、身份、授权、协作传输和人工智能
提供商。核心编辑和文件工作流程不需要 A3S 后端或捆绑
云服务。

## 证据，而非承诺

这些是与真实情况的视觉回归基线
[Playground](https://a3s-lab.github.io/Office/playground/)，不是概念图。

<p align="center">
  <a href="./visual-tests/__snapshots__/linux/desktop-1280/document.png">
    <img src="./visual-tests/__snapshots__/linux/desktop-1280/document.png" alt="A3S Office Document editor with an Office-style ribbon and paginated project brief" width="100%">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png" alt="A3S Office Spreadsheet editor showing a quarterly execution plan">
      </a>
      <br><sub><strong>电子表格</strong> — 公式、表格、格式设置、排序、筛选和打印工作流程</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png" alt="A3S Office Presentation editor with slide thumbnails and a slide canvas">
      </a>
      <br><sub><strong>演示文稿</strong> — 结构化幻灯片、对象、动画和演示者流程</sub>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png" alt="A3S Office Markdown editor in synchronized source and preview mode">
      </a>
      <br><sub><strong>Markdown</strong> — GFM 源码、可视化编辑和同步预览</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png" alt="A3S Office PDF editor with search, annotation, page organization, save, and download controls">
      </a>
      <br><sub><strong>PDF</strong> — PDFium 渲染、表单、注释、页面组织和保存</sub>
    </td>
  </tr>
</table>

## 为什么 A3S Office 存在

嵌入办公室工作不是一个问题。一个可靠的产品需要四个
同时发生的事情：

1. **完整的交互界面** — 功能区、对话框、窗格、快捷方式、
   响应式布局、可访问性和可预测的焦点。
2. **格式本机行为** — DOCX、XLSX、PPTX、Markdown 和 PDF 不能
   扁平化为一个最小公分母模型而不失去意义。
3. **明确的所有权边界** - 编辑者不应规定文件的位置
   实时、用户是谁、权限如何工作或允许哪个人工智能提供商。
4. **确定性自动化** - 代理和后端作业需要键入，
   冲突感知突变而不是 UI 抓取。

A3S Office 将这些问题分开，同时通过一个解决方案将它们暴露出来
包和一套有界合同。

## 五个表面，一个产品边界

- **文档** (`DOCX`、`HTML`、`TXT`) — 结构化创作，实时
  分页、表格、方程式、参考文献、评论和 PDF 输出。
- **电子表格** (`XLSX`、`XLS`、`ODS`、`CSV`) — 稀疏工作表，
  公式、表格、格式、排序/过滤、验证、数据透视表、图表和
  打印工作流程。
- **演示文稿** (`PPTX`) — 幻灯片、键入的场景对象、母版/布局、
  过渡、有界入口/出口动画、注释、幻灯片和
  演示者视图。
- **Markdown** (`MD`) — GFM 源、视觉模式、分割预览、直接圆形
  旅行和本机自动化。
- **PDF** — PDFium 渲染、搜索、表单、注释、历史记录、保存和
  页面组织。

表面是懒惰地要求的。大型或昂贵的工作被隔离在后面
可取消的 Workers、Rust WebAssembly、视口边界渲染和
显式 PDFium 运行时。 React、Vue 3、Web 组件和框架中立
核心 API 使用相同的受控内容模型。

## 核心设计

<p align="center">
  <a href="./assets/readme/architecture.svg">
    <img src="./assets/readme/architecture.svg" width="100%" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

两个执行平面解决不同的工作：

- **浏览器编辑**通过React提供完整的交互界面，
  Vue 3、Web 组件和核心 API。 Workers、Rust/WASM 和 PDFium 保持
  布局、计算、解析和渲染有界。
- **本机自动化**提供确定性文件读取、验证、
  突变、批处理、CLI 命令、标准 MCP 服务器和 Office 技能
  无需启动桌面 Office。

两个平面均由主机控制：

| A3S Office拥有|您的产品拥有 |
| ---| ---|
|格式本机模型和命令 |内容持久性和版本历史记录 |
|导入、导出、布局和渲染 |身份、授权和策略 |
|编辑器 UI 和响应式交互 |应用程序外壳和导航 |
|类型化协作和代理端口 |房间、交通、模型提供商和 AI 生命周期 |

这两架飞机都不需要 A3S 后端。阅读
[浏览器架构](./docs/latest/en/browser-editor-architecture.md)和
[原生引擎设计](./docs/latest/en/native-office-engine.md)精确
边界。

## 快速开始

### 尝试完整的产品

最快的第一次成功是
[现场Playground](https://a3s-lab.github.io/Office/playground/)。它暴露了
普通文档模板、最近功能、文件导入和每个编辑器
无需本地安装。

要在本地运行相同的 Playground，请使用 Node.js 20+、Bun 1.3+ 和 Rust 1.85+：

```bash
git clone https://github.com/A3S-Lab/Office.git
cd Office
bun install --frozen-lockfile
bun run playground
```

### 嵌入受控的 React 编辑器

```bash
bun add @a3s-lab/office react react-dom
```

导入样式表一次，给编辑器一个明确高度的主机，并保留
`onChange` 发出的完整值：

```tsx
import { useState } from 'react';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

const initialContent: DocumentContent = {
  type: 'document',
  html: '<h1>Project brief</h1><p>Start editing here.</p>',
  pageSize: 'a4',
  pageColor: '#ffffff',
};

export function ProjectBrief() {
  const [content, setContent] = useState(initialContent);

  return (
    <main style={{ height: '100dvh', minHeight: 0 }}>
      <DocumentEditor
        content={content}
        onChange={setContent}
        theme="system"
      />
    </main>
  );
}
```

编辑拥有编辑事务。您的应用程序决定何时、何地、
以及如何存储受控值。

### 选择一个入口点

|切入点|用它来 |
| ---| ---|
| `@a3s-lab/office/react` | Lazy React 编辑器组件和预加载助手 |
| `@a3s-lab/office/vue` | Vue 3 适配器与 `v-model:content` |
| `@a3s-lab/office/web-component` |与框架无关的自定义元素 |
| `@a3s-lab/office/core` |模型、模板、导入/导出、文件工作流程和 Yjs 绑定 |
| `@a3s-lab/office/styles.css` |共享编辑器和交互系统样式|

可复制的集成位于
[组件文档](https://a3s-lab.github.io/Office/docs/components/)。

## 文件保留文件

Core API 可以在不安装编辑器的情况下导入和导出文件：

```ts
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '@a3s-lab/office/core';

const shell = createArtifact('blank-document');
const artifact = await importOfficeFile(file, {
  artifactId: shell.id,
  onProgress: ({ stage, progress }) => {
    console.info(stage, Math.round(progress * 100));
  },
});

const output = await createArtifactBlob(artifact);
```

导入的 DOCX 工件是有源支持的。安全、未经编辑的包装部件和
稳定的本土身份在有限的规则下得以保留；不支持或
不安全的结构会产生兼容性诊断或显式失败
而不是附加到错误的内容。结构化文档值也可以
通过版本化快照编解码器跨进程边界。

请参阅
[文档](./docs/latest/en/components/document.mdx),
[电子表格](./docs/latest/en/components/spreadsheet.mdx)，
[演示文稿](./docs/latest/en/components/presentation.mdx)，以及
[PDF](./docs/latest/en/components/pdf.mdx) 文件特定参考
合同。

## 无需捆绑云的协作

每个编辑器都通过以下方式公开相同的传输中立协作边界
Yjs/Yrs 内容和意识。浏览器用户、本机副本、CLI 会话、
MCP 客户端和 A3S Code可以参与同一主机拥有的文档。

A3S Office 提供特定于格式的绑定、本地撤消、验证存在、
远程选择或位置、评论、建议和本地冲突
分型突变。主机提供房间、认证、授权、
交付、离线缓冲、持久性和 `Y.Doc`。

这种分离使得协作成为可选的基础设施，而不是一个
强制帐户或存储服务。从
【协作指南](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
或可运行的
[A3S启动示例](./examples/collaboration-server/)。

## 自动化，无需 UI 抓取

Rust CLI、标准 MCP 服务器、类型化 Rust API 和打包的 Office Skill
共享相同的有界文件合约：

```bash
# Validate and inspect a file.
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- view report.docx outline --json

# Apply an exact guarded mutation.
cargo run -p a3s-office-cli -- set report.docx /body \
  --find Draft --replace Final --json

# Expose the same contracts over standard MCP.
cargo run -p a3s-office-cli -- mcp
```

本机协作副本可以交换标准 Yjs 更新和状态
矢量，然后应用键入的文档、Markdown、电子表格、演示文稿和
PDF 突变无需解释 Office 的私有 CRDT 架构。

阅读[自动化指南](https://a3s-lab.github.io/Office/docs/automation/)
和 [CLI 参考](./docs/latest/en/cli-reference.md)。

在编码代理使用缓存的指导之前，请验证打包的技能合同
与`cargo run -p a3s-office-cli -- skills manifest a3s-office --json`（或
安装后`a3s-office skills manifest a3s-office --json`）。的
清单包括 `SKILL.md` 和每个的字节数和 SHA-256 值
捆绑参考。

对于浏览器编辑器工作，请使用存储库的本地操作符而不是
组装临时 shell 条件。它是一个基于 Commander 的 CLI，其中包含一个
Writer、电子表格、演示文稿、Markdown 和 PDF 的声明矩阵：

```bash
bun run office:ops -- plan all --json
bun run office:ops -- capabilities --json
bun run office:ops -- doctor --json
bun run office:ops -- gate writer --run \
  --browser-driver standalone \
  --cdp-port 9345
bun run office:ops -- visual spreadsheet --project compact-768
```

声明性运算符检查重点 A3S 测试 ACL； `gate --run`
将它们作为主要交互合约来执行，而 Playwright 只是一个
补充桌面/紧凑像素基线。屏幕截图和诊断保留
低于`.a3s-test/office-ops/`。它从不轮询 CI 或更改已提交的视觉效果
基线。使用 `a3s agent start/observe/act/finish` 进行有界探索
会话和 `a3s cua certification --json` 在本机 GUI 工作之前。上锁的
目前不支持 CUA 驱动程序 0.10.0 Windows 配置文件，因此 Windows
浏览器编辑器证据使用 A3S Test Web/CDP。在 Windows 上，
提供 `--cdp-port` 自动编译本机 `.exe` 适配器
`.a3s-test/office-ops/`，将交互式参数排除在 `.cmd` 解析之外。
`bun run office:ops -- wps-probe --connector` 捕获有界的 WPS COM
用于 UI/OOXML 奇偶校验的参考；该探测器是证据，而不是产品
运行时。使用`--connector-type straight|elbow|curved`选择键入的WPS
通过 Commander CLI 参考形状。 `wps-fields-probe --profile
numeric|common --json` 记录已安装的WPS数字或公共字段
指令并馈送相应的 A3S 测试夹具；它同样是一个
显式的本地引用捕获，从来都不是 CI 的先决条件。
`wps-ui-probe --profile shell|fields|all --json` 记录已安装的WPS
Writer 窗口 shell、功能区/状态栏和与字段相关的命令 ID 作为
输入 JSON UX 参考。它启动并关闭一个拥有的 COM 实例，因此
收据可以指导 UI 决策，而无需将 WPS 转变为产品运行时或
CI 依赖性。

对于探索性浏览器工作，同一个 Commander 操作员公开类型化的 A3S
测试代理操作：`click`、`hover`、`focus`、`double-click`、`context-click`、
`fill`、`type`、`check`、`uncheck`、`select`、`drag`、`press`、`wheel`、
`viewport`、等待、断言、屏幕截图、可访问性、控制台和页面
错误。使用显式目标语法（`@e7`、`css=...`、`role=role|name`、
`label=...`、`placeholder=...`、`testid=...`、`automation=...` 或 `text=...`）
并保持 A3S 测试`observe → one action → observe` 生命周期。这保持
Codex 编辑器操作可在所有五个表面上键入和重现。

## 当前版本

版本 `0.77.0` 深化了五种编辑器的 WPS UI/UX 对齐，并把共享证据矩阵扩展到
103 个 ACL 契约与 78 个视觉契约：

- Writer：域代码切换与锁定/取消链接、分页符/超链接/查找替换、下划线与双删除线、
  字号与缩进族（含 `Alt+Shift` 方向键）、段落上移下移、对齐与分散对齐、项目符号、
  行距、标题、正文、`Shift+F3` 大小写循环以及 `Ctrl+Space` 清除格式。
- Spreadsheet：工作簿视图（编辑栏 / 显示公式 / 网格线 / 标题）与
  `Ctrl+PageUp`/`Ctrl+PageDown` 工作表导航。
- Presentation：备注窗格、空白屏幕、幻灯片浏览 Enter、数字+Enter 跳转、Home/End、
  F5/Shift+F5、新建/复制/删除幻灯片、组合快捷键展示，以及 `Ctrl+F1` 功能区折叠视觉证据。
- Markdown：`Ctrl+F1` 功能区折叠，以及 `Ctrl+B` / `Ctrl+I` / `Ctrl+K` 源码格式与链接对话框证据。
- PDF：WPS 缩放、翻页与搜索焦点快捷键。

版本`0.76.0`为现有的桌面和紧凑的视觉证据添加了
Writer WPS 数字字段工作流程：

- 声明性矩阵现在公开 61 个 ACL 合约和 56 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 作者证据涵盖现场`PAGE`、`NUMPAGES`、`SECTION`和`PAGEREF`
  结果，带有`MERGEFORMAT`的本机数字指令，可访问的标签，
  单页分页和F9焦点恢复。
- 通过基于Commander的Office CLI进行聚焦视觉合约；
  本地 WPS 字段参考已针对版本 12.1.0.21541 进行了刷新。

版本`0.75.0`为现有的桌面和紧凑的视觉证据添加了
Writer WPS 布局和字体网格工作流程：

- 声明性矩阵现在公开 61 个 ACL 合约和 55 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 作者证据涵盖导入的 A4 分页、自动行度量、
  拉丁文/CJK 字体规格、文档网格线间距、混合脚本方向、
  和脚本感知的行高指标。
- 聚焦视觉合约通过基于 Commander 的 Office CLI
  两种布局都带有空控制台和页面错误证据。

版本`0.74.0`在现有的基础上添加了桌面和紧凑的视觉证据
Writer WPS快捷工作流程：

- 声明性矩阵现在公开 61 个 ACL 合约和 54 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 作者证据涵盖格式复制/粘贴、互斥的文本大小写
  效果、一步撤消、段落对齐/间距、标题样式、
  拼写、审阅、菜单发现和编辑器焦点保存。
- 本地 WPS Writer COM UI 探针针对版本 12.1.0.21541 运行； A3S测试
  1.0.1 通过了带有空控制台和页面错误证据的快捷方式 ACL。

版本 `0.73.0` 提升了现有的电子表格 WPS 复制上述工作流程
进入共享的 WPS/UI 证据矩阵：

- 声明性矩阵现在公开 61 个 ACL 合约和 53 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 桌面和紧凑证据封面`Ctrl+'`公式副本，`Ctrl+Shift+'`
  计算值复制、目标式保存、一步撤消以及
  网格焦点恢复。
- 基于 Commander 的 Office CLI 在两种布局中运行可视化合同，并且
  A3S 测试 1.0.1 通过了 ACL，具有空控制台和页面错误证据。

版本`0.72.0`提升现有电子表格WPS字体别名并直接
将颜色重置工作流程整合到共享 WPS/UI 证据矩阵中：

- 声明性矩阵现在公开 60 个 ACL 合约和 52 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 桌面和紧凑型证据盖`Ctrl+2`/`Ctrl+3`/`Ctrl+4`，自动
  颜色和无填充重置、撤消语义和网格焦点恢复。
- 基于 Commander 的 Office CLI 在两种布局中运行可视化合同，并且
  A3S Test 1.0.1 通过了带有空控制台和页面错误的 ACL 场景
  证据。

版本`0.71.0`提升电子表格超链接和高级下划线
工作流程进入本机 Windows 之上的共享 WPS/UI 证据矩阵
A3S测试操作员：

- 声明性矩阵现在公开 60 个 ACL 合约和 51 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 电子表格证据现在涵盖不安全 URL 拒绝、隐藏工作表链接
  保护、标准化 HTTPS 地址、高级下划线样式、单元格格式
  奇偶校验、键盘快捷键和撤消语义。
- A3S Test 1.0.1 验证了所有 60 个 ACL 文件；四个电子表格视觉案例通过
  在桌面和紧凑布局，以及实时超链接会话恢复网格
  焦点没有页面错误或控制台消息。生产Playground
  shell 还提供其共享徽标和网站图标资产。

版本`0.70.0`推广电子表格手机查找和工作表重命名
工作流程进入本机 Windows 之上的共享 WPS/UI 证据矩阵
A3S测试操作员：

- 声明性矩阵现在公开 58 个 ACL 合约和 49 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 电子表格电话证据现在涵盖触摸大小的查找控件、匹配单元格
  选择、无效的工作表名称反馈、取消重命名和恢复
  工作表菜单焦点。
- A3S Test 1.0.1 验证了所有 58 个 ACL 文件；四个新的电子表格手机视觉效果
  案件在桌面和紧凑布局中通过，并恢复了实时查找会话
  网格焦点，没有页面错误或控制台消息。

版本`0.69.0`推广电子表格手机任务窗格和上下文菜单
工作流程进入本机 Windows 之上的共享 WPS/UI 证据矩阵
A3S测试操作员：

- 声明性矩阵现在公开 56 个 ACL 合约和 47 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 电子表格电话证据现在涵盖数据枢轴模式焦点遏制，
  键盘遍历、上下文菜单消除和网格焦点恢复。
- A3S Test 1.0.1 验证了所有 56 个 ACL 文件；四个新的电子表格手机视觉效果
  在桌面和紧凑布局以及实时任务窗格会话中传递的案例
  恢复了功能区调用程序，没有页面错误或控制台消息。

版本`0.68.0`将演示电话审阅工作流程提升到共享
原生 Windows A3S 测试操作符之上的 WPS/UI 证据矩阵：

- 声明性矩阵现在公开 54 个 ACL 合约和 45 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 演示电话证据现在涵盖图表窗格模态焦点遏制，
  无效草稿回滚、评论审核模式惰性、草稿取消、
  和原始调用者焦点恢复。
- A3S Test 1.0.1 验证了所有 54 个 ACL 文件；四种新的演示电话视觉效果
  在桌面和紧凑布局上通过的案例以及实时评论会议
  提交评论并恢复焦点，没有页面错误。

版本`0.67.0`将电子表格外观工作流程提升为共享
原生 Windows A3S 测试操作符之上的 WPS/UI 证据矩阵：

- 声明性矩阵现在公开 52 个 ACL 合约和 43 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 电子表格现在具有明确的对角线边框、线性/路径渐变和
  具有响应式对话框和撤消覆盖的模式填充合同。
- A3S Test 1.0.1 验证了所有 52 个 ACL 文件；所有 10 个新的电子表格视觉效果
  在桌面和紧凑布局上传递的案例，以及实时路径梯度编辑
  恢复网格焦点，没有页面错误。

版本`0.66.0`将电子表格排序和自动过滤工作流程提升到
在本机 Windows A3S 测试操作符之上共享 WPS/UI 证据矩阵：

- 声明性矩阵现在公开 49 个 ACL 合约和 40 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 电子表格现在具有明确的外观/颜色、自定义列表、表格拥有、
  从左到右、部分范围和简体中文文本排序合约。
- A3S Test 1.0.1 验证了所有 49 个 ACL 文件；所有 14 个电子表格视觉案例
  在桌面和紧凑布局中通过，并显示实时排序对话框会话
  颜色排序控件没有页面错误。

版本`0.65.0`在原生基础上加深了Writer WPS/UI证据矩阵
Windows A3S 测试操作员：

- 声明性矩阵现在公开 43 个 ACL 合约和 34 个视觉合约
  跨 Writer、电子表格、演示文稿、Markdown 和 PDF。
- 作者证据现在包括字符位置、比例、间距、重点、
  隐藏文本、OpenType、内容控制、审阅冲突、配对移动
  修订版和电话跟踪更改以及现有的 WPS 套件。
- A3S Test 1.0.1 验证了所有 43 个 ACL 文件； 41 个作家视觉案例通过
  桌面和紧凑布局，跳过一种平台条件情况。

版本`0.64.0`在原生Windows之上加深了WPS/UI证据矩阵
A3S 测试操作器以 `0.63.1` 发货：

版本 `0.63.1` 在本机之上强化了 Windows A3S Test 操作员
Windows 适配器：

- 在 Windows 上，`--cdp-port` 自动编译本机 `.exe` 适配器
  `.a3s-test/office-ops`，将选择器和操作 JSON 保留在 `.cmd` 之外
  参数解析，同时保留固定的独立驱动程序。
- `doctor --json` 仅在兼容的 A3S 测试后报告 CUA 认证
  能力探针；过时的 0.x 二进制文件仍保持故障关闭状态并进行静态 ACL 检查
  仍可用于诊断。

版本`0.62.0`完成了声明性WPS证据的Writer端
`0.61.0` 中提供的类型化工作流程清单顶部的矩阵：

- 作家的共享矩阵现在包括现有的 WPS 快捷方式、布局奇偶性、
  和字体/网格 A3S 测试套件以及所有生成的 DOCX 装置。 `check all`,
  因此，`gate writer`和`plan writer --json`暴露了相同的有界
  本地 UI/UX 合约，而不是将这些套件保留为单独的脚本。

版本`0.61.0`为五个编辑器添加了声明式、类型化的工作流程清单
`0.60.0` 中附带的 WPS Writer UI/UX 参考工作流程之上的表面：

- `office:ops:plan` / `office-ui-ops plan <surface> --json` 扩展共享
  矩阵分为确定性夹具、ACL 检查、A3S 测试门、桌面和
  紧凑的视觉契约、有界代理会话和 WPS 参考探针。
  Codex 可以使用 JSON 切换，而无需发明每个表面的 shell 逻辑。

版本`0.60.0`在
WPS 字段设置和数字开关奇偶校验已发货：

- 指挥官`wps-ui-probe`捕获有界的`shell`、`fields`或`all`
  来自隔离的 WPS Writer COM 实例的配置文件。它记录了真实的WPS
  版本/构建、窗口边界、可见 shell 栏、现场命令 ID 以及
  与辅助功能相关的命令标题，无需声明 COM 是浏览器
  布局证据。
- A3S Test Web/CDP 仍然是主要的桌面/紧凑型浏览器交互
  合同。锁定的 Windows CUA 配置文件仍然明确不受支持，并且
  探测收据仅是当地参考证据。

版本`0.59.0`在WPS之上添加了类型化的Writer字段设置工作流程
数字开关和连接器奇偶校验已发货：

- 插入功能区打开一个有界的字段设置对话框，用于页面、部分、
  和书签页面引用字段。它公开了 WPS 兼容的数字，
  日期/时间、超链接和目标选择，保留 `MERGEFORMAT`，并保留
  未知的导入日期格式保持不变，直到用户更改它们。
- A3S 测试涵盖桌面和 390px 手机流程，包括响应式
  ribbon的两步横向发现、预览、F9刷新、焦点恢复、
  屏幕截图、可访问性和干净的浏览器诊断。 Windows CUA
  当锁定的配置文件未经过合同测试时，保持故障关闭状态。

版本 `0.58.0` 扩展有界 Writer 字段，带有 WPS 数字开关和
将 Writer 连接器工作流程继续到 WPS
直、肘和弯曲形状子集：

- 连接器功能区公开一种类型化的连接器类型控件。一样的
  值驱动实时线、路由折线或二次路径 SVG、紧凑型
  控件和可撤消的编辑。
- WPS VML o:spt=32、33 和 37 导入为直、弯头和曲线
  连接器。原生 DrawingML 自定义几何体保留所选类型
  DOCX 导出并重新打开；任意路径点编辑仍然会失败。
- A3S测试涵盖了三个WPS记录，包括屏幕截图、可访问性、
  控制台和页面错误证据。桌面和紧凑型剧作家检查
  断言不同的 SVG 几何形状。 Windows CUA 保持故障关闭状态，而
  锁定的配置文件未经合同测试。

同一版本保留公共页面、部分和书签页面引用
当 WPS 发出数字开关（例如 `\\* ROMAN`）时，字段处于活动状态，
`\\* ALPHABETIC`，或`\\* Ordinal`；未知的开关仍然被缓存并且
诊断而不是近似。

版本`0.56.1`使A3S测试五面编辑器矩阵成为日常UI/UX
合同：

- 电子表格、演示文稿、Markdown、模板发现和 PDF 工作流程
  通过在桌面和手机宽度上键入 A3S 测试操作进行检查。稳定
  模板控件在激活之前聚焦，以便键盘和代理流动
  具有相同的可发现性行为。
- PDF Blob 导出为速度较慢的 Web/CDP 观察者保留了其对象 URL；
  ACL 证明提取保留了剩余的页面操作，而
  确切的 PDF 字节保留在补充剧作家合同中。
- 屏幕截图、辅助功能快照、空控制台/页面错误诊断、
  并保留明确的独立驱动程序分类作为证据。
  Windows CUA 保持故障关闭状态，直到其锁定的配置文件经过合同测试。

版本 `0.56.0` 扩展了有界 Writer 直连接器工作流程
原生箭头式奇偶校验：

- 连接器功能区不暴露、三角形、隐形、菱形、椭圆形和
  打开箭头样式。一种类型的值驱动实时 SVG 标记，紧凑
  控件、可撤消的更新和本机 DOCX `a:headEnd`/`a:tailEnd` 导出。
- WPS VML `startarrow`/`endarrow`值导入到同一模型中，包括
  `classic`、`open`、`diamond` 和 `oval`；未知值归一化为无。
  Windows WPS 12.0 COM 探针记录有界 3/4 引用及其
  发出`open`/`classic` VML 对。
- A3S 测试 ACL 涵盖创作和真正的 WPS 夹具导入，而桌面和
  紧凑的视觉运行保留了屏幕截图、可访问性、控制台和
  页面错误证据。

版本 `0.55.0` 扩展了有界 Writer 直连接器工作流程
原生行式奇偶校验：

- 连接器功能区显示实线、虚线、点线和点划线样式。一
  输入的值驱动实时 SVG、紧凑的控件、可撤消的更新，以及
  原生 DOCX `a:prstDash` 导出。
- WPS VML `dashstyle` 和 DrawingML `a:prstDash` 导入到同一模型中。
  Windows WPS 12.0 COM 探针记录本机虚线 `DashStyle=4`
  参考；不受支持的长尾样式保持标准化而不是
  以精确编辑的形式呈现。
- A3S 测试 ACL 涵盖创作和真正的 WPS 夹具导入，而桌面和
  紧凑的视觉运行保留了屏幕截图、可访问性、控制台和
  页面错误证据。

版本`0.54.1`强化了有界周围的本地编辑器测试循环
指挥官操作员和 A3S 测试合同：

- Windows CDP 适配器直接调用固定的独立浏览器
  在进程退出时解析，因此完成的 ACL 操作不会等待分离
  代理浏览器会话句柄或端口文件轮询。
- Writer 文本框矩阵之前明确发现响应式溢出
  打开填充/边框控件并在之前滚动手机模板表面
  打开一个文档。四种桌面/手机场景保留截图，
  没有产品 UI 的可访问性、控制台和页面错误证据
  解决方法。

版本`0.54.0`添加了有界可编辑直连接器工作流程和
所有五个编辑器界面的单个本地运算符：

- 隔离的 WPS VML 直连接器导入到类型化的 Writer 模型中
  具有端点百分比、内联/浮动布局、线条颜色/宽度以及
  三角形箭头。本机 DOCX 导出写入有界的 DrawingML 线并
  保持稳定的绘图特性；混合段落和路由/不受支持
  连接器保持故障关闭诊断。
- 上下文 Writer 功能区显示连接器插入和编辑，同时
  基于 Commander 的 `office:ops` CLI 路由装置、主要 A3S 测试 ACL
  运行、补充视觉证据、WPS COM 参考捕获以及
  通过一个声明性编辑器矩阵观察-决定-行动代理生命周期。
  本地门覆盖桌面和紧凑视口，无需等待 CI。

版本`0.53.1`记录了可编辑切片之前的WPS绘图边界：

- Windows WPS 12.0 COM 探针将 `Shapes.AddConnector` 记录为 VML
  `v:shape`（`o:spt="32"`，`#_x0000_t32`）。专用`docx.connectors`
  诊断报告，证明并保留端点、路由、箭头和
  兼容性路径上的浮动锚语义。
- 确定性连接器 DOCX 夹具、重点 Rstest 和本地 A3S 测试
  ACL 验证导入使 Writer 编辑器保持可用，不会创建
  `data-document-text-box` 节点，并捕获可访问性和干净的浏览器
  诊断。实时浏览器运行仍然依赖于固定的本地
  适配器；静态 ACL 验证是发布合同的一部分。

版本 `0.53.0` 扩展了 Writer 的有界本机文本框工作流程
WPS 引用的形状几何体并关闭导入到编辑器模型循环：

- 上下文文本框功能区显示矩形、圆角矩形、椭圆形、
  菱形和三角形预设。一种类型的形状值驱动实时页面，
  预览、PDF 捕获、本机 DOCX 导出以及一步撤消/重做。
- 独立的 WPS `mc:AlternateContent` 文本承载形状保留其原生形状
  几何、放置、填充、轮廓、填充、垂直锚点和绘图
  身份。混合段落、连接符、任意形状、畸形体，
  和不受支持的 DrawingML 分支保留诊断的兼容性边界。
- 修复了结构化模型解析回归，因此导入了形状属性
  不会被默认文本框值替换。桌面和 390/768 px 浏览器
  流程涵盖所有五种形状控制、WPS 导入、上下文发现、
  可访问性、视口遏制和干净的诊断；本地A3S
  根据固定的发布合同检查测试 ACL。

版本`0.52.0`添加了有界的、WPS引用的整段修订
对作者的忠诚：

- 精确的纯文本段落插入和删除现在作为一个原子导入
  审查项目，保留作者和日期，接受或拒绝整个块，
  并通过本机`w:pPr/w:rPr`段落标记记录进行往返。
- 进口商接受使用单独的段落标记和正文修订ID
  当作者、时间戳和结构一致时通过 WPS。孤立的段落中断
  合并/拆分、混合或关系绑定内容、格式错误的元数据，以及
  超限输入保持明确的故障关闭诊断。
- WPS Office 12.1.0.22215 COM 探针确认 `Bravo\r` 删除并
  `Delta\r` 参考文档的插入范围。桌面和 390 像素
  审核流程添加了真正的 DOCX 导入、原子决策、键盘和焦点
  检查、可访问性证据、干净的诊断和触摸尺寸的紧凑型
  行动。

版本 `0.51.0` 扩展了 Writer Compare 与有界、WPS 引用
跨段落文本移动：

- 独特的词汇范围可以在对齐的简单文本段落之间移动或
  同一部分的标题并成为一对 `move` 评论项目。
  分离器随范围移动；接受、拒绝、撤销或重新打开
  DOCX 保留准确的源文本和修订文本。
-“审阅”功能区和“更改”窗格保留一张本地化的**移动**卡，其中包含
  源/目的地标记、目的地导航、作者指南，以及
  移动感知吐司很重要。候选者部分重复或标记不匹配
  边界、丰富或关系绑定的内容、表格和超限输入
  保持故障关闭的普通修改或诊断。
- WPS 12.0 COM/UIA 探针观察到 `CompareDocuments` 返回普通
  删除/插入测试重新排序的记录。因此，A3S 记录了这一点
  配对推理作为有界局部增强，没有声称 WPS
  通过该 API 公开了本机移动类型。

版本 `0.50.0` 将 Writer Compare 扩展为有界的、WPS 引用的文本
移动工作流程：

- 在一个简单段落或标题内移动的确定性词汇范围
  成为一对`move`评论项目。分离器随范围移动，并且
  接受、拒绝、撤消或重新打开 DOCX 保留确切的来源
  并修订案文。
-“审阅”功能区和“更改”窗格显示一张本地化的 **移动** 卡，其中包含
  源/目的地标记、目的地导航、更新的作者指南、
  移动感知吐司很重要。模棱两可的重复、丰富或关系密切
  内容、跨段落移动和超限输入仍会失败关闭
  边界。
- WPS 12.0 COM/UIA 探针观察到 `CompareDocuments` 返回普通
  删除/插入测试重新排序的记录。因此，A3S 记录了其
  配对推理作为有界局部增强，没有声称 WPS
  通过该 API 公开了本机移动类型。

版本 `0.49.0` 添加了有界本机 Writer 移动修订并强化了
本地 PDF 工作流程：

- 严格和过渡性 DOCX `w:moveFrom`/`w:moveTo` 文本对作为一个导入
  原子`move`审查项目。接受/拒绝解析源和目标
  双方齐心协力，不可篡改的协作审核记录了最终的结果
  决定，导出/重新打开恢复本机移动记录，没有私有的
  标记泄漏。
- 丰富的、范围标记的、关系绑定的、畸形的和不成对的动作仍然存在
  可见作为诊断的兼容性边界。审阅窗格标签移动
  清晰地导航到目标范围而不选择文本
  两侧之间。
- Playground 现在从稳定处提供固定的 PDFium WebAssembly 资源
  原始相对路径，使直接 PDF 导入在桌面和紧凑中保持就绪状态
  布局。

版本`0.48.1`修复了Writer浮动选择工具栏的分割强调
无需更改文档模型或公共命令合约即可进行控制：

- 下划线和删除线控件现在重置浏览器本机按钮
  外观并使用与周围相同的紧凑主题状态
  操作，包括单像素分组分隔符和键盘焦点处理。
- 桌面版和紧凑版浏览器视觉合约都涵盖了该修复
  视口；格式化命令、辅助功能名称和拆分菜单仍然保留
  不变。

版本 `0.48.0` 添加了有界本机 Writer 内容控件，无需
引入第二个文档模型或远程服务：

- 插入功能区和响应式对话框作者内联纯文本或富文本
  具有别名、程序标签、多行文本、边框/标签/隐藏的控件
  外观和可选的主题颜色。每个接受的更改都是一次键入
  更新和一个撤消步骤，并且控件公开可访问的文本框名称。
- 在事务边界强制执行内容锁和外壳锁，包括
  直接编辑交易。锁定的内容不会被意外更改；
  输入的解锁/删除命令保持明确且可审核。
- 直接段落DOCX`w:sdt`通过严格和控制来回
  具有无冲突本机 ID 的过渡 WordprocessingML，丰富的运行
  格式、锁定/多行元数据和 Word 2012 外观/颜色。活跃
  绑定、占位符、重复区域、表单控件、嵌套或
  关系绑定结构保持安全且具有兼容性的可编辑文本
  诊断而不是不精确的可编辑承诺。

版本 `0.47.0` 为 Writer 添加了有界公共字段工作流程，无需
引入第二个文档模型或远程服务：

- 插入功能区显示实时字数和字符数。页数，页数，
  SECTION、SECTIONPAGES、DATE 和 TIME 持续刷新测量值
  页面模型，而生成的字段结果不包含在文档统计中。
- 交叉引用对话框可以插入带有书签的 `PAGEREF`
  稳定的目标身份。支持原生 DOCX 交换机作为字段往返；
  不支持的、格式错误的、嵌套的或丢失目标的指令仍被缓存
  具有明确兼容性诊断的文本。

版本 `0.46.0` 添加了本机、有界的 Writer 文本框，而无需引入
第二个文档模型或远程服务：

- 插入功能区添加了可编辑的 **文本框** 块。它的上下文功能区
  控制内联/浮动放置、毫米几何形状、相对偏移、
  通过一次受控更新来填充、轮廓、填充和垂直对齐
  每个意图一个撤消记录。
- 隔离的 WordprocessingML 文本框通过本机 DOCX 几何图形往返，
  主体属性和填充/轮廓状态并共享相同的实时、预览和
  PDF 投影。混合或格式错误的图纸仍然是明确的诊断
  边界，而不是默默地被压平。

版本 `0.45.0` 添加了有界 Writer 图片变换，而无需引入
第二个文档模型或远程服务：

- 上下文**图片**功能区现在显示四分之一圈旋转加上
  水平和垂直反射与可访问的 Lucide 图标。 **图片
  属性** 在响应式对话框中保留相同的控件，并且每个
  已确认的更改是一种可撤消的、受控的更新。
- 转换一致地投射到编辑、预览、PDF 捕获、
  以及本机 DOCX `a:xfrm` 字段（`rot`、`flipH` 和 `flipV`）。进口
  90 度/反射值重新打开为可编辑状态，而任意角度
  或格式错误的值会产生兼容性诊断并安全地标准化。

版本`0.44.0`向电子表格数据验证添加了依赖的本地下拉列表
无需引入远程服务：

- **数据→数据工具→数据验证**接受有界的
  `=INDIRECT(...)` 源由引用文本和单单元格引用组成。
  为每个目标单元格重新评估相对驱动因素，因此区域列
  可以通过工作簿本地命名范围驱动区域所有者列表或
  一行/一列区域。
- 空驱动程序显示空列表。外部书籍，整行/列范围，
  缺少工作表、未缓存的公式和格式错误的表达式无法关闭；
  公式保持在 255 个 Unicode 字符以内，并且本地范围读取保持在
  有界。
- 受控模型使编写的公式保持紧凑，同时网格得到
  有界运行时间投影。本机 XLSX 导入/导出/重新打开保留
  列出公式和名称，公共**新建→数据验证**模板包括
  区域 → 区域所有者示例。

版本`0.43.0`使电子表格公式条件格式成为本地，
可编辑的工作流程而不是仅文件的保存路径：

- **主页 → 条件格式 → 自定义公式** 作者规则
  相对/绝对参考、有限跨表参考、独立
  文本/填充颜色、有序优先级和 Stop-if-true 行为。进口
  XLSX 表达式规则在同一编辑器中重新打开。
- 共享同步求值器仅读取缓存的工作簿值、上限
  255 个 Unicode 字符的公式和 1,024 个单元格的每个决策，扫描
  有界空白范围，没有致密工作簿，并且失败关闭
  外部、整行/列、缺失工作表或未缓存的公式引用。
- 原生 XLSX 差分样式、`sqref`、优先级和公式往返
  通过导出/重新打开。公共 **新建 → 公式条件格式** 模板显示
  阻止规则和跨表阈值，无需任何远程服务。

版本 `0.42.0` 使用本地自定义公式扩展了电子表格数据验证
保持有界、可检查和本机文件兼容的工作流程：

- 在同一个可访问的对话框中使用可选的 `=` 前缀编写公式；
  相对参考根据选定范围和建议值进行锚定
  在进入受控工作簿之前进行评估。
- 常用局部函数、单元格/范围引用和工作表限定引用
  使用 255 个字符的公式和 1,024 个单元格读取同步评估
  预算。外部、整行/列、缺失表和未缓存公式
  引用失败关闭而不是默默地接受不安全的编辑。
- 自定义规则通过 XLSX 往返并公开演示
  **新建→数据验证**模板，以及现有列表、日期和警报
  分支机构。粘贴和对象级批量写入在预检中保持独立
  边界。

版本 `0.41.0` 将电子表格数据验证警报与本地、
传统 Office 使用的可测试交互模型：

- Stop 通过可访问的通知阻止无效编辑。警告和
  信息询问是否应保留无效值，使用作者
  标题和消息加上当前输入，因此决定是明确的。
- 保留警告或信息值通过受控提交一次
  工作簿API；取消则恢复原来的值。选择、聚焦、
  撤消和协作保持在相同的有界编辑路径上，并且本机
  XLSX `errorStyle` 元数据在重新打开时保持无损。

版本 `0.40.0` 使用原子扩展了 Writer 的本地、可测试审阅工作流程
有序列表编号修订：

- 跟踪的编号样式和起始编号更改现在显示为一个
  每个有序列表范围对审核卡进行编号。 Accept 保留当前列表；
  拒绝恢复精确的有界基线；撤消、Yjs/Yrs 协作，以及
  不可变的决策具有相同的意图标识。普通单层
  十进制、字母和罗马`w:numberingChange`记录往返
  DOCX，而格式错误或结构不明确的记录则无法关闭。

- 幻灯片对象现在可以拥有一个入口和一个出口动画。所共享的
  以对象为中心的动画选项卡作者出现/消失、淡入/淡出、飞入/飞出、
  并通过一次有序触发、定时、方向、预览来放大/缩小效果，
  协作、剪贴板和撤消模型。浏览器播放和本机 PPTX
  时序树导入/导出保留两个类，但格式错误或重叠
  序列无法通过诊断关闭。

- 文档现在包括一流的双语版本感知
  [新增内容](https://a3s-lab.github.io/Office/docs/changelog.html)时间线。
  它提供了编辑器范围、用户结果、兼容性证据和详细信息
  引用而冻结的文档隐藏了尚不存在的版本。

- 作家的共享高级字体对话框现在可以编辑完整的有界 Office
  2010 OpenType 版式集：16 种连字组合、数字形式和
  间距、样式集 1-20 和上下文替代。混合选择，
  跟踪格式、撤消、正文/页面镶边故事、精确的 DOCX 重新打开以及
  格式错误的输入诊断与结构化方程共享一个模型。
- 电子表格 XLSX 导入/导出现在保留工作簿的本机 1900 或
  1904 日期系统、精确键入的日期序列、动态日期过滤器以及
  跨受控和协作的时代正确的当前日期创作
  重新开放流动。
- 电子表格自定义排序现在可以识别本机表和自动过滤器拥有的表
  范围，保持标题和总计结构安全，并重新应用活动
  经过一种公式安全、不可撤销的排序后过滤可见性。
- 电子表格自定义排序现在包括响应式、可键盘访问的
  用于创建、编辑、删除和重新排序有界的首选项管理器
  用户序列，同时保持七个内置序列只读。
- 文档、视觉 Markdown 和演示文稿文本合成对于
  中文和其他输入法：预编辑文本保留在本地，受控替换
  等待结算，只有承诺的值到达主机。
- 电子表格排序和自动筛选工作流程在本地运行，包括有界
  多键排序、本机通配符过滤器、顶部/底部过滤器、自定义列表、
  和简体中文排序规则选择。- 结构化参考计算、计算列填充和本机总计
  行共享有界的 Rust/WASM 和 JavaScript 路径。
- 共享文件菜单现在具有明确的操作图标、可读标签、有界
  键盘导航，以及独特的破坏性动作处理。
- PDF页面组织执行插入、删除、旋转、重新排序、提取、
  合并，并在具有 Blob 级别撤消/重做的专用 Worker 中拆分。

浏览以产品为中心的【最新动态](https://a3s-lab.github.io/Office/docs/changelog.html)
页面或详尽的工程[变更日志](./CHANGELOG.md)。Playground
从带有版本标签的模板网格公开这些工作流程。

## 能力和边界

A3S Office 更倾向于明确的保真度边界而不是无声的近似。
`Supported` 表示路径具有确定性行为或本机往返
证据。 `Partial` 表示存在有记录边界的有用路径。

### 文档

**强路径：**结构化创作、分页、表格、参考文献、评论、
大型纯文档窗口和源感知 DOCX 往返。

**边界：**长尾DrawingML、字段、精确布局奇偶校验和可搜索
矢量 PDF 仍然是部分的。

### 电子表格

**强路径：**稀疏编辑、依赖关系感知公式、表格、
排序/过滤、验证、图表、数据透视和 XLSX 往返。

**边界：**公式广度、外部数据、宏、加载项和专家
分析仍然不完整。

### 推介会

**强路径：**场景编辑、母版/布局、过渡、有界入口
动画、演示者工作流程和 PPTX 往返。

**边界：**更广泛的形状、效果、媒体、动画系列和精确的
打印/视频输出仍然部分或不受支持。

### 降价

**强路径：** GFM 源代码和可视化编辑、同步预览、
协作、直接往返和本机自动化。

**边界：** Markdown 保留格式本机而不是充当转换
每个丰富文档结构的层。

### PDF

**强路径：** PDFium 渲染、搜索、表单、注释、保存和
有界页面组织。

**边界：**现有内容流编辑、OCR、签名、优化、
值得信赖的编辑需要明确的提供者或未来的引擎。

完整的能力清单、优先级和退出证据位于
[能力路线图](./ROADMAP.md)。协作交付有自己的
[路线图](./COLLABORATION_ROADMAP.md)。

## 项目状态

A3S Office 是 1.0 之前的版本。公共模型和组件道具是可用的，但是
次要版本可能仍包含重大模型更改。所需的迁移
记录在[变更日志](./CHANGELOG.md)中。

目标是可预测的浏览器编辑和诚实的文件保存，而不是
与每个桌面套件功能的像素奇偶校验。仍然存在不支持的语义
可见的兼容性边界。的
【编辑质量路线图](./docs/latest/en/editor-quality-roadmap.md)记录
移动每个边界所需的证据。

## 开发

从存储库根运行标准检查：

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

重点本地 A3S 测试门涵盖首次开放焦点和 IME 组成：

```bash
bun run test:e2e:initial-focus:check
bun run test:e2e:initial-focus
bun run test:e2e:presentation-chinese-ime:check
bun run test:e2e:presentation-chinese-ime
```

受控文档 IME 生命周期也针对 Playwright 的固定运行
CI 中的 WebKit 引擎：

```bash
bun run playground:ime:webkit
```

GitHub Actions 不需要 A3S 测试。参见
[贡献.md](./CONTRIBUTING.md)，
[E2E指南](./tests/e2e/README.md)，以及
[完整验证的视觉测试指南](./visual-tests/README.md)
矩阵。

## 文档

- [现场Playground](https://a3s-lab.github.io/Office/playground/)
- [文档中心](https://a3s-lab.github.io/Office/docs/)
- [新增内容](https://a3s-lab.github.io/Office/docs/changelog.html)
- [React、Vue、Web 组件和核心 API](https://a3s-lab.github.io/Office/docs/components/)
- [实时协作](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
- [CLI、MCP 和 Office 技能](https://a3s-lab.github.io/Office/docs/automation/)
- [浏览器编辑器架构](./docs/latest/en/browser-editor-architecture.md)
- [原生Office引擎](./docs/latest/en/native-office-engine.md)
- [编辑质量路线图](./docs/latest/en/editor-quality-roadmap.md)
- [工程变更日志](./CHANGELOG.md)

## 社区和安全

欢迎贡献。阅读
【投稿指南](./CONTRIBUTING.md)以及
[打开更改之前的行为准则](./CODE_OF_CONDUCT.md)。报告
通过私有进程怀疑存在漏洞
[SECURITY.md](./SECURITY.md)，不是通过公开问题。

## 许可证

A3S Office 可在 [MIT 许可证](./LICENSE) 下使用。捆绑 PDFium 和
其他第三方资产带有附加通知
[THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)。