---
title: CLI 参考
description: a3s-office 的安装、结构化读取、类型化修改、校验、批处理与 MCP 用法。
---

# CLI 参考

`a3s-office` 在本机读取、校验和修改 Office 文件。自动化场景应默认使用 `--json`，
并在写入后执行独立的读取或校验命令。

## 安装

```bash
cargo install \
  --git https://github.com/A3S-Lab/Office.git \
  --locked a3s-office-cli
```

```bash
a3s-office --version
a3s-office --help
```

## 基本命令

| 命令 | 用途 | 是否修改文件 |
| --- | --- | --- |
| `validate <file>` | 校验文件包、关系、结构与受支持语义 | 否 |
| `view <file> <view>` | 返回目录、工作表、幻灯片、表格或其他语义视图 | 否 |
| `get <file> <selector>` | 读取选择器对应的结构化值 | 否 |
| `set <file> <selector> ...` | 修改文本、格式或结构化属性 | 是 |
| `remove <file> <selector>` | 删除受支持目标 | 是 |
| `batch <file> ...` | 原子执行带版本的多项修改 | 是 |
| `watch <file> ...` | 监听文件并输出有界变化事件 | 否 |
| `mcp` | 启动标准输入输出 MCP 服务 | 否 |

运行具体命令的 `--help` 可以查看当前版本支持的精确参数。

```bash
a3s-office view --help
a3s-office set --help
a3s-office batch --help
```

## 读取与校验

```bash
a3s-office validate report.docx --json
a3s-office view report.docx outline --json
a3s-office view workbook.xlsx sheets --json
a3s-office view deck.pptx slides --json
```

结构化输出用于确认文件身份、目标是否存在、修改前状态和修改后状态。不要解析面向人工
阅读的终端句子来驱动智能体判断。

## 选择器

选择器使用与文件结构对应的稳定路径。索引以命令帮助和当前协议为准，常见形式如下：

```text
/body/p[1]
/body/p[1]/r[1]
/Sheet1/cell[A1]
/slide[1]/shape[1]
```

在写入前先用 `get` 或 `view` 验证目标。目标不存在、命中多个对象或属性组合无效时，
命令应明确失败，不会选择“最像”的对象继续执行。

## 文本与格式修改

```bash
a3s-office set report.docx /body \
  --find Draft \
  --replace Final \
  --json

a3s-office set report.docx '/body/p[1]/r[1]' \
  --bold true \
  --font-family Aptos \
  --font-size 14 \
  --text-color 123456 \
  --language zh-CN \
  --json
```

布尔值、颜色、单位、语言和枚举会在提交前校验。不支持的格式组合返回类型化错误，
不会写入一半状态。

## 原子批处理

需要一次修改多个目标时，应使用带协议版本的批处理。引擎先验证整个输入，再在内存或
临时文件中应用；全部操作成功后才发布输出。

批处理适合编码智能体保存计划、审计修改和精确回放。它不是自然语言指令集合，每个操作
都必须属于当前协议定义的类型化命令。

## 写后验证

```bash
a3s-office set report.docx /body \
  --find Draft \
  --replace Final \
  --json

a3s-office validate report.docx --json
a3s-office view report.docx outline --json
```

零退出码表示命令执行完成，不等于业务目标已经满足。验证命令应读取用户真正关心的结构，
并确认输出路径、源文件保留策略和最终文件状态。

## 原生实时协作

`collab` 使用与浏览器相同的标准 Yjs v1 更新和状态向量。宿主仍负责房间、认证、授权与
传输；本地副本负责带校验和的更新日志、稳定操作回执、断点事件与重启恢复。

```bash
a3s-office collab join .a3s/application.replica \
  --artifact-id application --kind pdf --actor-id agent-7 \
  --actor-kind agent --mode edit --operation-id join-1 \
  --input browser-bootstrap.update --json

a3s-office collab mutate .a3s/application.replica \
  --actor-id agent-7 --artifact-id application --kind pdf --mode edit \
  --operation-id annotation-create-1 \
  --mutation '{"type":"pdf-create-annotation","annotationId":"annotation-1","pageIndex":0,"annotation":{"id":"annotation-1","pageIndex":0,"type":9,"rect":{"origin":{"x":68,"y":78},"size":{"width":300,"height":28}},"segmentRects":[{"origin":{"x":68,"y":78},"size":{"width":300,"height":28}}],"strokeColor":"#f59e0b","color":"#f59e0b","opacity":0.48,"contents":"Review this heading"}}' \
  --json
```

Document 的持久选区评论可以使用独立、绑定 Actor 的 `comment` 副本。先用
`collab read` 读取投影版本 2，再使用其中完全一致的段落/文字 ID、锚点文字和 UTF-16
偏移；新评论和回复的 `author` 必须与协作后端认证的 Actor 显示名称一致：

```bash
a3s-office collab join .a3s/report-review.replica \
  --artifact-id report --kind document --actor-id agent-7 \
  --actor-kind agent --mode comment --operation-id comment-join-1 \
  --input browser.update --json

a3s-office collab read .a3s/report-review.replica --json

a3s-office collab mutate .a3s/report-review.replica \
  --actor-id agent-7 --artifact-id report --kind document --mode comment \
  --operation-id comment-create-1 \
  --mutation '{"type":"document-comment-create","commentId":"comment-1","paragraphId":"00000001","expectedTextId":"00000002","startUtf16":6,"endUtf16":12,"expectedText":"review","author":"Ada Reviewer","createdAt":"2026-08-17T00:00:00.000Z","text":"Clarify this review point."}' \
  --json

a3s-office collab mutate .a3s/report-review.replica \
  --actor-id agent-7 --artifact-id report --kind document --mode comment \
  --operation-id comment-reply-1 \
  --mutation '{"type":"document-comment-reply","commentId":"comment-1","replyId":"reply-1","author":"Ada Reviewer","createdAt":"2026-08-17T00:01:00.000Z","text":"Suggested wording is ready."}' \
  --json

a3s-office collab mutate .a3s/report-review.replica \
  --actor-id agent-7 --artifact-id report --kind document --mode comment \
  --operation-id comment-resolve-1 \
  --mutation '{"type":"document-comment-set-resolved","commentId":"comment-1","resolved":true}' \
  --json

# resolved:false 用于重新打开。省略 replyId 会删除当前 Actor 自己的整个线程。
a3s-office collab mutate .a3s/report-review.replica \
  --actor-id agent-7 --artifact-id report --kind document --mode comment \
  --operation-id comment-delete-reply-1 \
  --mutation '{"type":"document-comment-delete","commentId":"comment-1","replyId":"reply-1"}' \
  --json
```

投影 v2 返回评论、回复、解决状态、`detached`，以及锚点的 `paragraphId`、`textId`、
`startUtf16`、`endUtf16` 和当前 `text`。正文、格式、结构和其他 Document 选项在
`comment` 模式中保持只读；删除仅限当前 Actor 自己创建的评论或回复。相同稳定 ID 的
完全一致重试保持幂等，陈旧选区、冲突 ID 和越权删除不会写入日志。MCP 的
`office_collaboration_read` 与 `office_collaboration_mutate` 使用同一投影和 Mutation
对象；需要乐观并发保护时，把读取结果的 `stateVectorBase64` 作为
`ifStateVectorBase64` 传回。

经过认证、由浏览器产生的 Document `suggest` 更新可以先由 A3S Boot/Yrs 服务端执行
语义授权，再同步进原生副本。但当前 `NativeOfficeCollaborationMutation` 尚无
`document-suggestion-*` 或 `document-change-decision-*` 变体，因此 `collab mutate`、
`office_collaboration_mutate` 与投影 v2 暂时不能主动创建、决定或列出这些建议记录。
智能体不应通过手写私有 ProseMirror/Yjs Mark 绕过封闭的 Mutation Schema。非 Document
的 `suggest` 和 `comment` 模式仍没有本地类型化修改。

Spreadsheet 原生修改使用稳定的 `sheetId` 与从零开始的 `row`、`column`。
`spreadsheet-set-cell` 会递归比较 `expectedCell`、当前共享单元格与 `nextCell`，只写入
实际变化的叶子。这样并发的批注或样式修改可以与值、公式修改合并，同一叶子的过期修改
则不会追加日志。创建空坐标时必须显式传入 `expectedCell: null`；
`spreadsheet-delete-cell` 要求完整单元格精确匹配，避免用过期的局部观察执行删除。

```bash
a3s-office collab mutate .a3s/plan.replica \
  --actor-id agent-7 --artifact-id plan --kind spreadsheet --mode edit \
  --operation-id spreadsheet-edit-1 \
  --mutation '{"type":"spreadsheet-set-cell","sheetId":"sheet-data","row":1,"column":0,"expectedCell":{"v":10,"m":"10"},"nextCell":{"v":12,"m":"12","f":"=6*2"}}' \
  --json
```

密集工作表保留并按限制扩展矩阵行长度；稀疏工作表与首次写入的空工作表继续使用
`celldata`。坐标、共享根结构、单元格 JSON 大小和深度、危险对象键以及孤立字段都会在
写入前校验。

Presentation 原生修改使用稳定的 `containerKind`（`slide`、`master` 或 `layout`）与
`containerId`。`presentation-create-element` 会写入完整对象和规范创建声明，并可通过
`afterElementId` 放在一个仍然有效的对象之后；相同内容的同 ID 重试保持幂等，不同内容、
已删除 ID 或不存在的锚点会失败。`presentation-update-element` 比较完整的
`expectedElement`、当前对象和 `nextElement`，只写入发生变化的顶层字段。无关字段的
并发修改可以合并，同一字段的过期修改不会追加日志。`presentation-delete-element` 要求
完整对象精确匹配，成功后从可见顺序移除对象并写入不可复用 ID 的持久墓碑。对象的
`id` 与 `type` 始终不可变。`presentation-move-element` 不使用数组下标，而是通过
`expectedAfterElementId` 声明调用方观察到的稳定前驱，并通过 `afterElementId` 声明目标
前驱；`null` 表示对象顺序数组的第一项。对象已经位于目标位置时操作保持幂等；否则源
前驱过期、目标锚点不存在或已删除、对象已删除以及对象以自身为锚点都会在持久更新前
失败。成功移动只移除并重新插入该对象自己的顺序项，不覆盖对象字段、其他对象或容器。

```bash
a3s-office collab mutate .a3s/deck.replica \
  --actor-id agent-7 --artifact-id deck --kind presentation --mode edit \
  --operation-id presentation-edit-1 \
  --mutation '{"type":"presentation-update-element","containerKind":"slide","containerId":"slide-1","elementId":"title-1","expectedElement":{"id":"title-1","type":"text","x":10,"y":10,"width":80,"height":20,"text":"Draft"},"nextElement":{"id":"title-1","type":"text","x":16,"y":10,"width":80,"height":20,"text":"Final"}}' \
  --json

a3s-office collab mutate .a3s/deck.replica \
  --actor-id agent-7 --artifact-id deck --kind presentation --mode edit \
  --operation-id presentation-move-1 \
  --mutation '{"type":"presentation-move-element","containerKind":"slide","containerId":"slide-1","elementId":"title-1","expectedAfterElementId":"background-1","afterElementId":null}' \
  --json
```

PDF 原生操作包括表单值，FreeText、Highlight、Underline、StrikeOut、Ink 的创建、
递归乐观更新和不可逆删除，以及只追加的脱敏提议、页面旋转/删除/重排提议与唯一最终
决定。批注 ID、源页、类型和来源身份不可漂移；无关 JSON 叶子可以并发合并，同一叶子
的过期修改会在写入前失败。源 PDF 与签名字节始终由宿主持有。完整命令与错误契约见
[English 完整 CLI 参考](/en/cli-reference.md)。

## MCP

```bash
a3s-office mcp
```

MCP 使用与 CLI 相同的准入、读取、修改、限制和错误协议。客户端应完成标准初始化，
读取工具 schema，再发送类型化请求；不要把 CLI 的显示文字当作 MCP 参数。

## 退出与清理

普通命令结束时会关闭自己创建的文件、监听器和临时资源。长时间运行的监听或 MCP 服务
收到第一次中断后会开始有界清理；调用方仍应等待进程退出，并把超时与业务失败区分开。

完整命令参数、全部选择器、批处理 schema、PDF 证据字段和错误代码见
[English 完整 CLI 参考](/en/cli-reference.md)。
