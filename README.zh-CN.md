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

`office-ui-ops plan <surface> --json` 是该矩阵的类型化清单入口
（`bun run office:ops -- plan <surface> --json`）。

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

版本 `0.79.0` 深化 Writer Phase 0 可审阅的单元格与行属性修订：

- **单元格格式** — 无关系的 `tcW`（auto / dxa / pct）与 `noWrap` 先验可作为
  `cell-formatting` 审阅，并覆盖实时修订跟踪与原生 DOCX 导出。
- **行格式** — 无关系的 `hidden` 与 `jc`（left / center / right）先验可作为
  `row-formatting` 审阅，并覆盖当前行属性导出。
- **边界** — 更广的单元格/行属性集仍为不透明元数据或失败闭合。

更早的版本请查看产品向
[更新日志](https://a3s-lab.github.io/Office/docs/changelog.html) 时间线，以及完整工程
[CHANGELOG](./CHANGELOG.md)。[在线 Playground](https://a3s-lab.github.io/Office/playground/)
从带版本标签的模板网格暴露这些工作流。

## 能力和边界

A3S Office 优先显式保真边界，而不是静默近似。
`Supported` 表示该路径具备确定性行为或原生往返证据。
`Partial` 表示已有可用路径，但边界已文档化。

### 文档

**强路径：** 结构化写作、分页、表格、引用、审阅、
大篇幅纯文窗口，以及源感知的 DOCX 往返。

**边界：** 长尾 DrawingML、域、精确版式对齐，以及可搜索
矢量 PDF 仍为部分支持。

### 电子表格

**强路径：** 稀疏编辑、依赖感知公式、表格、
排序/筛选、验证、图表、数据透视、XLSX 往返，以及保留工作簿
原生的 1900 或 1904 日期系统。

**边界：** 公式覆盖面、外部数据、宏、加载项和专业
分析仍不完整。

### 演示文稿

**强路径：** 场景编辑、母版/版式、切换、有界进入
动画、演示者流程，以及 PPTX 往返。

**边界：** 更广的形状、效果、媒体、动画族，以及精确
打印/视频输出仍为部分支持或未支持。

### Markdown

**强路径：** GFM 源码与可视化编辑、同步预览、
协作、直接往返，以及原生自动化。

**边界：** Markdown 保持格式原生，而不是充当一切
富文档结构的转换层。

### PDF

**强路径：** PDFium 渲染、搜索、表单、批注、保存，以及
有界页面组织。

**边界：** 现有内容流编辑、OCR、签名、优化，以及
可信密文处理需要明确的 Provider 或后续引擎。

完整能力清单、优先级与退出证据见
[能力路线图](./ROADMAP.md)。协作交付另有
[路线图](./COLLABORATION_ROADMAP.md)。

## 项目状态

A3S Office 仍处于 1.0 之前。公共模型与组件 props 已可用，但
次要版本仍可能包含破坏性模型变更。必要迁移
记录在 [变更日志](./CHANGELOG.md)。

目标是可预期的浏览器编辑与诚实的文件保留，而不是
与每个桌面套件功能像素级对齐。不受支持的语义保持为
可见的兼容性边界。
[编辑器质量路线图](./docs/latest/en/editor-quality-roadmap.md) 记录
推进每个边界所需的证据。

## 开发

在仓库根目录运行标准检查：

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

聚焦的本地 A3S Test 门禁覆盖首次打开焦点与 IME 组合：

```bash
bun run test:e2e:initial-focus:check
bun run test:e2e:initial-focus
bun run test:e2e:presentation-chinese-ime:check
bun run test:e2e:presentation-chinese-ime
```

受控 Document IME 生命周期也会在 CI 中对 Playwright 固定的
WebKit 引擎运行：

```bash
bun run playground:ime:webkit
```

GitHub Actions 不要求安装 A3S Test。完整验证矩阵见
[CONTRIBUTING.md](./CONTRIBUTING.md)、
[E2E 指南](./tests/e2e/README.md) 与
[视觉测试指南](./visual-tests/README.md)。

## 文档

- [在线 Playground](https://a3s-lab.github.io/Office/playground/)
- [文档中心](https://a3s-lab.github.io/Office/docs/)
- [更新日志](https://a3s-lab.github.io/Office/docs/changelog.html)
- [React、Vue、Web Component 与 Core API](https://a3s-lab.github.io/Office/docs/components/)
- [实时协作](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
- [CLI、MCP 与 Office Skill](https://a3s-lab.github.io/Office/docs/automation/)
- [浏览器编辑器架构](./docs/latest/en/browser-editor-architecture.md)
- [原生 Office 引擎](./docs/latest/en/native-office-engine.md)
- [编辑器质量路线图](./docs/latest/en/editor-quality-roadmap.md)
- [工程变更日志](./CHANGELOG.md)

## 社区与安全

欢迎贡献。发起变更前请先阅读
[贡献指南](./CONTRIBUTING.md) 与
[行为准则](./CODE_OF_CONDUCT.md)。疑似漏洞请通过
[SECURITY.md](./SECURITY.md) 中的私密流程报告，不要公开议题。

## 许可证

A3S Office 以 [MIT License](./LICENSE) 提供。捆绑的 PDFium 与其他第三方
资产的附加声明见 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)。
