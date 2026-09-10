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

a3s-office collab session .a3s/application.replica \
  --poll-ms 100 --actor-name "A3S Agent" --actor-color "#2563eb" --json

a3s-office collab mutate .a3s/application.replica \
  --actor-id agent-7 --artifact-id application --kind pdf --mode edit \
  --operation-id annotation-create-1 \
  --mutation '{"type":"pdf-create-annotation","annotationId":"annotation-1","pageIndex":0,"annotation":{"id":"annotation-1","pageIndex":0,"type":9,"rect":{"origin":{"x":68,"y":78},"size":{"width":300,"height":28}},"segmentRects":[{"origin":{"x":68,"y":78},"size":{"width":300,"height":28}}],"strokeColor":"#f59e0b","color":"#f59e0b","opacity":0.48,"contents":"Review this heading"}}' \
  --json
```

`collab session` 通过 JSONL 接入宿主拥有的 WebSocket 或 IPC。`outbound` 与
`receive` 承载持久 Yjs 文档同步。提供 `--actor-name` 后，会话还会创建独立的内存
Yrs Awareness 参与者。宿主应把 `outbound-awareness` 转发为房间 Awareness，把远端
载荷写成 `receive-awareness`，并把离线通知写成 `peer-left`。`set-presence` 可以发布
active、idle 或 away 状态以及当前格式位置；有效远端变化会输出排序后的 `presence`
快照。重连会清理旧远端成员并重新发布本地状态，正常关闭会先输出 tombstone。
Presence 内容与时钟不会进入副本、Checkpoint、持久更新或操作回执。
`ready.clientId` 是本次宿主连接和所有房间 Envelope 的发送 ID；启用 Presence 时每次
启动都会重新生成，避免新 Awareness 被旧逻辑时钟拒绝。`ready.replicaClientId` 仍是
持久副本中稳定的 Yrs 作者 ID，宿主不能用它替换连接发送 ID。

Document 的持久选区评论可以使用独立、绑定 Actor 的 `comment` 副本。先用
`collab read` 读取投影版本 3，再使用其中完全一致的段落/文字 ID、锚点文字和 UTF-16
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

投影 v3 返回评论、回复、解决状态、`detached`，以及锚点的 `paragraphId`、`textId`、
`startUtf16`、`endUtf16` 和当前 `text`。正文、格式、结构和其他 Document 选项在
`comment` 模式中保持只读；删除仅限当前 Actor 自己创建的评论或回复。相同稳定 ID 的
完全一致重试保持幂等，陈旧选区、冲突 ID 和越权删除不会写入日志。MCP 的
`office_collaboration_read` 与 `office_collaboration_mutate` 使用同一投影和 Mutation
对象；需要乐观并发保护时，把读取结果的 `stateVectorBase64` 作为
`ifStateVectorBase64` 传回。

Document 建议使用独立、绑定 Actor 的 `suggest` 副本。创建前重新读取投影 v3，并提交
完全一致的段落/文字 ID、UTF-16 选区与原文。插入使用空选区和一个 `insertionId`，删除
使用非空选区、空 `replacement` 和一个 `deletionId`，替换同时携带两个 ID。Actor ID 始终
来自副本清单，`author` 必须与后端认证的显示名称一致：

```bash
a3s-office collab join .a3s/report-suggest.replica \
  --artifact-id report --kind document --actor-id agent-7 \
  --actor-kind agent --mode suggest --operation-id suggestion-join-1 \
  --input browser.update --json

a3s-office collab read .a3s/report-suggest.replica --json

a3s-office collab mutate .a3s/report-suggest.replica \
  --actor-id agent-7 --artifact-id report --kind document --mode suggest \
  --operation-id suggestion-create-1 \
  --mutation '{"type":"document-suggestion-create","paragraphId":"00000001","expectedTextId":"00000002","startUtf16":6,"endUtf16":8,"expectedText":"😀","replacement":"reviewed","insertionId":"agent-7-insertion-1","deletionId":"agent-7-deletion-1","author":"A3S Agent","createdAt":"2026-08-17T11:00:00.000Z"}' \
  --json

a3s-office collab diff .a3s/report-suggest.replica \
  --output agent-suggestion.update --json

a3s-office collab join .a3s/report-editor.replica \
  --artifact-id report --kind document --actor-id editor-1 \
  --actor-kind human --mode edit --operation-id editor-join-1 \
  --input agent-suggestion.update --json

a3s-office collab read .a3s/report-editor.replica --json

a3s-office collab mutate .a3s/report-editor.replica \
  --actor-id editor-1 --artifact-id report --kind document --mode edit \
  --operation-id suggestion-accept-1 \
  --mutation '{"type":"document-suggestion-decide","suggestions":[{"id":"agent-7-deletion-1","kind":"deletion","expectedActorId":"agent-7","expectedAuthor":"A3S Agent","expectedCreatedAt":"2026-08-17T11:00:00.000Z","expectedText":"😀"},{"id":"agent-7-insertion-1","kind":"insertion","expectedActorId":"agent-7","expectedAuthor":"A3S Agent","expectedCreatedAt":"2026-08-17T11:00:00.000Z","expectedText":"reviewed"}],"decision":"accept","decidedBy":"Grace Editor","decidedAt":"2026-08-17T11:01:00.000Z"}' \
  --json
```

拒绝使用 `decision: "reject"`。决定只能由 `edit` 副本执行，并会原子匹配整个建议批次。
接受会旋转受影响段落和已识别祖先表格行的文字身份；接受与拒绝都会移除实时 Mark，并为
每条建议写入不可变、带 Actor 的最终决定。相同稳定 ID 的一致重试保持幂等；陈旧身份或
文字、切开 UTF-16 代理对、重叠建议、复用 ID、伪造署名、不完整的替换批次和冲突决定都
不会追加持久日志。投影 v3 会列出实时 `suggestions`、精确位置和不可变
`changeDecisions`。智能体不应手写私有 ProseMirror/Yjs Mark。非 Document 的 `suggest`
和 `comment` 模式仍没有本地类型化修改。

Spreadsheet 原生修改使用稳定的 `sheetId` 与从零开始的 `row`、`column`。
`spreadsheet-set-cell` 会递归比较 `expectedCell`、当前共享单元格与 `nextCell`，只写入
实际变化的叶子。这样并发的批注或样式修改可以与值、公式修改合并，同一叶子的过期修改
则不会追加日志。创建空坐标时必须显式传入 `expectedCell: null`；
`spreadsheet-delete-cell` 要求完整单元格精确匹配，避免用过期的局部观察执行删除。
`spreadsheet-batch-cells` 可把一次粘贴、填充或其他有界手势中的 1 至 4096 个不同坐标
写进同一个事务。每项都使用相同的 set/create 递归保护或完整 delete 保护；所有保护先
基于同一份共享快照检查。每项必须显式提供 `nextCell`，其中 `nextCell: null` 表示完整
精确匹配后删除；任一冲突都会拒绝整个批次，不产生部分单元格或持久日志。

```bash
a3s-office collab mutate .a3s/plan.replica \
  --actor-id agent-7 --artifact-id plan --kind spreadsheet --mode edit \
  --operation-id spreadsheet-edit-1 \
  --mutation '{"type":"spreadsheet-set-cell","sheetId":"sheet-data","row":1,"column":0,"expectedCell":{"v":10,"m":"10"},"nextCell":{"v":12,"m":"12","f":"=6*2"}}' \
  --json

a3s-office collab mutate .a3s/plan.replica \
  --actor-id agent-7 --artifact-id plan --kind spreadsheet --mode edit \
  --operation-id spreadsheet-batch-2 \
  --mutation '{"type":"spreadsheet-batch-cells","sheetId":"sheet-data","changes":[{"row":1,"column":0,"expectedCell":{"v":12,"m":"12","f":"=6*2"},"nextCell":{"v":14,"m":"14","f":"=7*2"}},{"row":1,"column":1,"expectedCell":null,"nextCell":{"v":20,"m":"20"}},{"row":2,"column":0,"expectedCell":{"v":"obsolete","m":"obsolete"},"nextCell":null}]}' \
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
