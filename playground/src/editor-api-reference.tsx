import {
  Blocks,
  CodeXml,
  Info,
  MousePointer2,
  Puzzle,
  ShieldCheck,
} from 'lucide-react';
import { type ReactNode, useState } from 'react';
import { CodeBlock } from './code-block';

type EditorId =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

interface PropReference {
  name: string;
  type: string;
  required?: boolean;
  defaultValue?: string;
  frameworkBinding?: string;
  description: string;
}

interface EventReference {
  react: string;
  vue: string;
  webComponent: string;
  description: string;
}

interface EditorReference {
  component: string;
  label: string;
  summary: string;
  contentContract: string;
  contentTitle?: string;
  contentFields: readonly PropReference[];
  props: readonly PropReference[];
  events: readonly EventReference[];
  selectionMenu?: {
    description: string;
    notes: readonly string[];
    code: string;
  };
  extension: {
    status: 'available' | 'host';
    title: string;
    description: string;
    notes: readonly string[];
    code: string;
  };
}

const surfaceProps: readonly PropReference[] = [
  {
    name: 'theme',
    type: "'light' | 'dark' | 'system'",
    defaultValue: "'system'",
    frameworkBinding: 'theme',
    description: '编辑器颜色模式。',
  },
  {
    name: 'className',
    type: 'string',
    frameworkBinding: 'Vue / Web Component: class',
    description: 'React 宿主容器类名。',
  },
  {
    name: 'style',
    type: 'CSSProperties',
    frameworkBinding: 'Vue / Web Component: style',
    description: 'React 宿主容器内联样式；编辑器会填满该容器。',
  },
];

const controlledContentProps = (
  contentType: string,
  saveStatusDefault?: string,
): readonly PropReference[] => [
  {
    name: 'content',
    type: contentType,
    required: true,
    frameworkBinding: 'Vue: v-model:content · Element: .content',
    description: '受控内容模型；修改后由宿主保存并重新传入。',
  },
  {
    name: 'onChange',
    type: `(content: ${contentType}) => void`,
    required: true,
    frameworkBinding: 'Vue: @change · Element: change',
    description: '编辑内容发生变化时触发。',
  },
  {
    name: 'preview',
    type: 'boolean',
    defaultValue: 'false',
    frameworkBinding: 'preview',
    description: '只读预览模式。',
  },
  {
    name: 'saveStatus',
    type: 'string',
    defaultValue: saveStatusDefault,
    frameworkBinding: 'save-status',
    description: '显示宿主应用的保存状态。',
  },
  {
    name: 'fileActions',
    type: 'readonly OfficeFileAction[]',
    frameworkBinding: 'Vue: :file-actions · Element: .fileActions',
    description: '向文件菜单加入宿主操作，例如保存、另存为或版本历史。',
  },
];

const contentEvents = (contentType: string): readonly EventReference[] => [
  {
    react: 'onChange',
    vue: 'v-model:content / @change',
    webComponent: 'change',
    description: `返回最新 ${contentType}。`,
  },
];

const agentEvent: EventReference = {
  react: 'onAgentRequest',
  vue: '@agent-request',
  webComponent: 'agent-request',
  description: '把用户选区和操作意图交给宿主 AI 流程。',
};

const editorReferences: Record<EditorId, EditorReference> = {
  document: {
    component: 'DocumentEditor',
    label: '文字',
    summary: '分页文字编辑器，适合 DOCX、报告和长文档。',
    contentContract:
      'DocumentContent 同时保存结构化文档模型、HTML、页面、分节、批注、修订和文献数据。',
    contentFields: [
      {
        name: 'type',
        type: "'document'",
        required: true,
        description: '内容类型判别字段，固定为 document。',
      },
      {
        name: 'html',
        type: 'string',
        required: true,
        description: '兼容 HTML；应与编辑器输出的结构化 model 一起保存。',
      },
      {
        name: 'model',
        type: 'WorkDocumentModel',
        description: '带版本和修订号的 TipTap 文档树，是精确编辑的结构化模型。',
      },
      {
        name: 'pageSize',
        type: "'a4' | 'letter'",
        required: true,
        description: '默认页面尺寸；分节节点可覆盖页面设置。',
      },
      {
        name: 'orientation / margins / columns',
        type: 'Document layout fields',
        description: '默认方向、页边距和分栏设置。',
      },
      {
        name: 'pageChrome',
        type: 'WorkDocumentPageChrome',
        description: '首页、奇偶页的页眉、页脚和页码设置。',
      },
      {
        name: 'trackChanges / comments',
        type: 'boolean / WorkDocumentComment[]',
        description: '修订开关以及包含回复、解决状态的批注数据。',
      },
      {
        name: 'bibliography',
        type: 'WorkDocumentBibliography',
        description: '引用样式和文献来源。',
      },
    ],
    props: [
      ...controlledContentProps('DocumentContent', "'已自动保存'"),
      {
        name: 'extensions',
        type: 'Extensions',
        defaultValue: '[]',
        frameworkBinding: 'Vue: :extensions · Element: .extensions',
        description: '追加 TipTap 扩展；扩展名称不能与内置扩展重复。',
      },
      {
        name: 'kernelWasmUrl',
        type: 'string',
        defaultValue: '包内 office-kernel.wasm',
        frameworkBinding: 'kernel-wasm-url',
        description: '分页和几何计算使用的 WASM 资源地址。',
      },
      {
        name: 'layoutFonts',
        type: 'readonly DocumentLayoutFont[]',
        defaultValue: '包内中西文、阿拉伯文和希伯来文字体',
        frameworkBinding: 'Vue: :layout-fonts · Element: .layoutFonts',
        description: '排版测量字体；应与最终导出和显示字体匹配。',
      },
      {
        name: 'onAgentRequest',
        type: '(request: EditorAgentRequest) => void | Promise<void>',
        frameworkBinding: 'Vue: @agent-request · Element: agent-request',
        description: '接收文档选区发起的 AI 请求。',
      },
      {
        name: 'getSelectionMenuItems',
        type: 'GetDocumentSelectionMenuItems',
        frameworkBinding:
          'Vue: :get-selection-menu-items · Element: .getSelectionMenuItems',
        description:
          '完全接管选中文本的右键菜单，并为每个动作提供选区、全文和安全编辑命令。',
      },
      ...surfaceProps,
    ],
    events: [...contentEvents('DocumentContent'), agentEvent],
    selectionMenu: {
      description:
        '宿主决定菜单项、文案、可用状态和执行逻辑。回调会得到当前受控内容、全文、选区前后文、结构化选区以及可跟踪异步变化的编辑命令。',
      notes: [
        '传入 getSelectionMenuItems 后会完整替换内置选区菜单；库不会按关键词猜测业务动作。',
        '异步动作必须返回或 await Promise，编辑器会在模型响应期间跟踪原选区；原文被修改时 replaceText 会返回 stale-selection。',
        'replaceText、insertBefore 和 insertAfter 各自产生一次受控更新和一次撤销记录，并遵守修订模式。',
      ],
      code: documentSelectionMenuExample(),
    },
    extension: {
      status: 'available',
      title: '支持 TipTap Extensions',
      description:
        '追加快捷键、插件状态、Node、Mark 或 ProseMirror Plugin。扩展数组应保持引用稳定。',
      notes: [
        '使用唯一的 extension.name；与内置名称重复会直接报错。',
        '自定义 Node 或 Mark 必须同时考虑 DOCX 导入导出，否则只会保存在浏览器模型中。',
        '不要从 @a3s-lab/office/internal 路径导入实现。',
      ],
      code: tiptapExtensionExample('DocumentEditor', 'DocumentContent'),
    },
  },
  markdown: {
    component: 'MarkdownEditor',
    label: 'Markdown',
    summary: '默认左侧源码、右侧可视编辑与预览。',
    contentContract:
      'MarkdownContent 以 markdown 字符串为唯一内容源，源码和可视编辑器保持同步。',
    contentFields: [
      {
        name: 'type',
        type: "'markdown'",
        required: true,
        description: '内容类型判别字段，固定为 markdown。',
      },
      {
        name: 'markdown',
        type: 'string',
        required: true,
        description: '唯一持久化内容源；源码与可视编辑区都从该值同步。',
      },
    ],
    props: [
      ...controlledContentProps('MarkdownContent'),
      {
        name: 'extensions',
        type: 'Extensions',
        defaultValue: '[]',
        frameworkBinding: 'Vue: :extensions · Element: .extensions',
        description: '追加 TipTap 扩展；扩展名称不能与内置扩展重复。',
      },
      ...surfaceProps,
    ],
    events: contentEvents('MarkdownContent'),
    extension: {
      status: 'available',
      title: '支持 TipTap Extensions',
      description:
        '可加入快捷键和编辑行为。新增内容节点时，需要同时提供 Markdown 解析与序列化规则。',
      notes: [
        '不改变文档结构的快捷键、Plugin 和存储扩展最安全。',
        '自定义 Node 或 Mark 没有 Markdown 规则时，切换源码可能丢失该结构。',
        '扩展数组应使用 useMemo 或模块级常量。',
      ],
      code: tiptapExtensionExample('MarkdownEditor', 'MarkdownContent'),
    },
  },
  spreadsheet: {
    component: 'SpreadsheetEditor',
    label: '表格',
    summary: '受控工作簿编辑器，支持公式、工作表、图表和数据能力。',
    contentContract:
      'SpreadsheetContent 保存工作表、单元格、公式、图表、条件格式、透视表和页面设置。',
    contentFields: [
      {
        name: 'type',
        type: "'spreadsheet'",
        required: true,
        description: '内容类型判别字段，固定为 spreadsheet。',
      },
      {
        name: 'sheets',
        type: 'WorkSpreadsheetSheet[]',
        required: true,
        description: '工作表、单元格、公式、样式、图片、图表和透视表。',
      },
      {
        name: 'calculation',
        type: 'WorkSpreadsheetCalculationSettings',
        description: '计算模式、迭代次数、精度和加载时重算策略。',
      },
      {
        name: 'namedRanges',
        type: 'WorkSpreadsheetNamedRange[]',
        description: '工作簿级或工作表级名称。',
      },
      {
        name: 'printAreas / printTitles',
        type: 'Spreadsheet print fields',
        description: '打印区域以及重复标题行、标题列。',
      },
      {
        name: 'pageBreaks / pageSetups',
        type: 'Spreadsheet page fields',
        description: '分页符、纸张、方向、缩放、页边距及页眉页脚。',
      },
    ],
    props: [
      ...controlledContentProps('SpreadsheetContent', "'已自动保存'"),
      {
        name: 'kernelWasmUrl',
        type: 'string',
        defaultValue: '包内 office-kernel.wasm',
        frameworkBinding: 'kernel-wasm-url',
        description: '公式计算和工作簿内核的 WASM 资源地址。',
      },
      {
        name: 'onAgentRequest',
        type: '(request: EditorAgentRequest) => void | Promise<void>',
        frameworkBinding: 'Vue: @agent-request · Element: agent-request',
        description: '接收当前单元格或区域发起的 AI 请求。',
      },
      ...surfaceProps,
    ],
    events: [...contentEvents('SpreadsheetContent'), agentEvent],
    extension: {
      status: 'host',
      title: '当前使用宿主扩展点',
      description:
        '稳定的工作簿 Extension Context 尚未公开。当前通过文件菜单和 AI 回调接入宿主能力。',
      notes: [
        'fileActions 适合保存、导出、版本历史和业务审批。',
        'onAgentRequest 提供结构化选区，不需要解析界面文字。',
        '内部命令运行时不是公共 API，请勿从 internal 路径导入。',
      ],
      code: hostExtensionExample('SpreadsheetEditor', 'SpreadsheetContent'),
    },
  },
  presentation: {
    component: 'PresentationEditor',
    label: '演示',
    summary: '基于场景图的幻灯片编辑器，文本编辑按需使用 TipTap。',
    contentContract:
      'PresentationContent 保存幻灯片、母版、版式、元素、批注和切换效果。',
    contentFields: [
      {
        name: 'type',
        type: "'presentation'",
        required: true,
        description: '内容类型判别字段，固定为 presentation。',
      },
      {
        name: 'slides',
        type: 'WorkSlide[]',
        required: true,
        description: '幻灯片及其元素、备注、批注和切换效果。',
      },
      {
        name: 'width / height',
        type: 'number',
        description: '演示画布尺寸；未传时使用默认宽高比。',
      },
      {
        name: 'masters',
        type: 'WorkPresentationMaster[]',
        description: '母版背景和母版元素。',
      },
      {
        name: 'layouts',
        type: 'WorkPresentationLayout[]',
        description: '关联母版的版式、占位符和继承元素。',
      },
    ],
    props: [
      ...controlledContentProps('PresentationContent', "'已自动保存'"),
      {
        name: 'kernelWasmUrl',
        type: 'string',
        defaultValue: '包内 office-kernel.wasm',
        frameworkBinding: 'kernel-wasm-url',
        description: '吸附、对齐和几何计算使用的 WASM 资源地址。',
      },
      {
        name: 'onAgentRequest',
        type: '(request: EditorAgentRequest) => void | Promise<void>',
        frameworkBinding: 'Vue: @agent-request · Element: agent-request',
        description: '接收幻灯片或元素发起的 AI 请求。',
      },
      {
        name: 'onStartSlideshow',
        type: '() => void',
        frameworkBinding: 'Vue: @start-slideshow · Element: start-slideshow',
        description: '用户开始放映时交给宿主处理。',
      },
      ...surfaceProps,
    ],
    events: [
      ...contentEvents('PresentationContent'),
      agentEvent,
      {
        react: 'onStartSlideshow',
        vue: '@start-slideshow',
        webComponent: 'start-slideshow',
        description: '请求宿主开始演示。',
      },
    ],
    extension: {
      status: 'host',
      title: '当前使用宿主扩展点',
      description:
        '演示命令已经模块化，但稳定的场景图 Extension Context 尚未作为公共 API 发布。',
      notes: [
        '使用 fileActions 接入保存、导出和协作入口。',
        '使用 onAgentRequest 处理幻灯片或元素级 AI 操作。',
        '使用 onStartSlideshow 接管路由、窗口或演示设备。',
      ],
      code: hostExtensionExample('PresentationEditor', 'PresentationContent'),
    },
  },
  pdf: {
    component: 'PdfViewer',
    label: 'PDF',
    summary: '按需加载 PDF，支持导航、搜索、批注和保存副本。',
    contentContract:
      'PDF 不使用受控 content；宿主通过 loadSource 提供 Blob，并通过 onSave 接收修改后的 Blob。',
    contentTitle: '文件生命周期',
    contentFields: [
      {
        name: '输入',
        type: 'Promise<Blob>',
        required: true,
        description: 'loadSource 异步返回 PDF Blob，不要求编辑器访问文件系统。',
      },
      {
        name: '版本',
        type: 'sourceKey?: string',
        description: '源文件版本变化时更新，触发释放旧 URL 并重新加载。',
      },
      {
        name: '输出',
        type: 'onSave(pdf: Blob)',
        description: '编辑后返回完整 PDF Blob，由宿主持久化。',
      },
      {
        name: '保存结果',
        type: 'Promise<boolean>',
        description: 'true 显示保存成功；false 或异常显示失败。',
      },
    ],
    props: [
      {
        name: 'loadSource',
        type: '() => Promise<Blob>',
        required: true,
        frameworkBinding: 'Vue: :load-source · Element: .loadSource',
        description: '异步读取 PDF；函数引用变化会重新加载。',
      },
      {
        name: 'onSave',
        type: '(pdf: Blob) => Promise<boolean>',
        frameworkBinding: 'Vue: :on-save · Element: .onSave',
        description: '存在时开启编辑保存；返回 true 表示保存成功。',
      },
      {
        name: 'fileName',
        type: 'string',
        defaultValue: "'document.pdf'",
        frameworkBinding: 'file-name',
        description: '编辑器和下载流程使用的文件名。',
      },
      {
        name: 'saveLabel',
        type: 'string',
        defaultValue: "'保存'",
        frameworkBinding: 'save-label',
        description: '保存按钮文字。',
      },
      {
        name: 'sourceKey',
        type: 'string',
        frameworkBinding: 'source-key',
        description: '源文件版本键；改变后强制重新加载。',
      },
      {
        name: 'wasmUrl',
        type: 'string',
        defaultValue: '包内 pdfium.wasm',
        frameworkBinding: 'wasm-url',
        description: 'PDFium WASM 资源地址。',
      },
      ...surfaceProps,
    ],
    events: [
      {
        react: 'onSave',
        vue: ':on-save',
        webComponent: '.onSave',
        description: '宿主持久化修改后的 PDF Blob。',
      },
    ],
    extension: {
      status: 'host',
      title: '当前使用加载与保存端口',
      description:
        'PDF 命令层已模块化，但 PDFium 与批注控制器仍属于内部实现，暂不接受自定义 Extensions。',
      notes: [
        '用 loadSource 对接对象存储、IndexedDB 或本地文件。',
        '用 onSave 控制权限、上传进度、版本号和失败处理。',
        'sourceKey 应对应文件版本，避免显示旧缓存。',
      ],
      code: pdfHostExample(),
    },
  },
};

const editorIds = Object.keys(editorReferences) as EditorId[];

export function EditorApiReference() {
  const [selectedId, setSelectedId] = useState<EditorId>('document');
  const reference = editorReferences[selectedId];

  return (
    <section
      id="guide/api"
      className="playground-integration-panel playground-guide-section"
      aria-labelledby="editor-api-title"
    >
      <header className="playground-guide-section-heading">
        <span>
          <Blocks size={18} />
        </span>
        <div>
          <h2 id="editor-api-title">组件 API</h2>
          <p>按编辑器查看完整属性、事件、内容契约和扩展方式。</p>
        </div>
      </header>

      <div
        className="playground-api-tabs"
        role="tablist"
        aria-label="编辑器 API"
      >
        {editorIds.map((id) => {
          const item = editorReferences[id];
          return (
            <button
              id={`editor-api-tab-${id}`}
              type="button"
              role="tab"
              key={id}
              aria-controls={`editor-api-panel-${id}`}
              aria-selected={selectedId === id}
              onClick={() => setSelectedId(id)}
            >
              <span>{item.label}</span>
              <code>{item.component}</code>
            </button>
          );
        })}
      </div>

      <article
        id={`editor-api-panel-${selectedId}`}
        className="playground-api-panel"
        role="tabpanel"
        aria-labelledby={`editor-api-tab-${selectedId}`}
      >
        <header className="playground-api-summary">
          <span>
            <CodeXml size={17} />
          </span>
          <div>
            <h3>{reference.component}</h3>
            <p>{reference.summary}</p>
          </div>
        </header>

        <div className="playground-api-contract">
          <Info size={15} />
          <span>{reference.contentContract}</span>
        </div>

        <ApiHeading
          icon={<Blocks size={15} />}
          title={reference.contentTitle ?? '内容模型'}
        />
        <div className="playground-api-table-wrap">
          <table className="playground-api-table content">
            <thead>
              <tr>
                <th scope="col">字段</th>
                <th scope="col">类型</th>
                <th scope="col">必填</th>
                <th scope="col">说明</th>
              </tr>
            </thead>
            <tbody>
              {reference.contentFields.map((field) => (
                <tr key={field.name}>
                  <th scope="row">
                    <code>{field.name}</code>
                  </th>
                  <td>
                    <code>{field.type}</code>
                  </td>
                  <td>{field.required ? '是' : '否'}</td>
                  <td>{field.description}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <ApiHeading icon={<CodeXml size={15} />} title="属性参数" />
        <div className="playground-api-table-wrap">
          <table className="playground-api-table">
            <thead>
              <tr>
                <th scope="col">属性</th>
                <th scope="col">类型</th>
                <th scope="col">必填</th>
                <th scope="col">默认值</th>
                <th scope="col">Vue / Web Component</th>
                <th scope="col">说明</th>
              </tr>
            </thead>
            <tbody>
              {reference.props.map((prop) => (
                <tr key={prop.name}>
                  <th scope="row">
                    <code>{prop.name}</code>
                  </th>
                  <td>
                    <code>{prop.type}</code>
                  </td>
                  <td>{prop.required ? '是' : '否'}</td>
                  <td>
                    <code>{prop.defaultValue ?? '—'}</code>
                  </td>
                  <td>
                    <code>{prop.frameworkBinding ?? prop.name}</code>
                  </td>
                  <td>{prop.description}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <ApiHeading icon={<ShieldCheck size={15} />} title="回调与事件" />
        <div className="playground-api-table-wrap">
          <table className="playground-api-table events">
            <thead>
              <tr>
                <th scope="col">React</th>
                <th scope="col">Vue</th>
                <th scope="col">Web Component</th>
                <th scope="col">说明</th>
              </tr>
            </thead>
            <tbody>
              {reference.events.map((event) => (
                <tr key={`${event.react}-${event.vue}`}>
                  <td>
                    <code>{event.react}</code>
                  </td>
                  <td>
                    <code>{event.vue}</code>
                  </td>
                  <td>
                    <code>{event.webComponent}</code>
                  </td>
                  <td>{event.description}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {reference.selectionMenu && (
          <>
            <ApiHeading
              icon={<MousePointer2 size={15} />}
              title="选区右键菜单"
            />
            <div className="playground-extension-status available">
              <div>
                <strong>宿主完全控制</strong>
                <span>{reference.selectionMenu.description}</span>
              </div>
              <ul>
                {reference.selectionMenu.notes.map((note) => (
                  <li key={note}>{note}</li>
                ))}
              </ul>
            </div>
            <CodeBlock
              code={reference.selectionMenu.code}
              label="自定义选区菜单示例"
              language="tsx"
            />
          </>
        )}

        <ApiHeading icon={<Puzzle size={15} />} title="Extensions" />
        <div
          className={`playground-extension-status ${reference.extension.status}`}
        >
          <div>
            <strong>{reference.extension.title}</strong>
            <span>{reference.extension.description}</span>
          </div>
          <ul>
            {reference.extension.notes.map((note) => (
              <li key={note}>{note}</li>
            ))}
          </ul>
        </div>
        <CodeBlock
          code={reference.extension.code}
          label={
            reference.extension.status === 'available'
              ? 'Extension 示例'
              : '宿主扩展示例'
          }
          language="tsx"
        />
      </article>
    </section>
  );
}

function ApiHeading({ icon, title }: { icon: ReactNode; title: string }) {
  return (
    <div className="playground-api-heading">
      {icon}
      <h3>{title}</h3>
    </div>
  );
}

function tiptapExtensionExample(
  component: 'DocumentEditor' | 'MarkdownEditor',
  contentType: 'DocumentContent' | 'MarkdownContent',
): string {
  return `import { Extension } from '@tiptap/core';
import { useMemo } from 'react';
import type { ${contentType} } from '@a3s-lab/office/core';
import { ${component} } from '@a3s-lab/office/react';

const hostShortcuts = Extension.create({
  name: 'hostShortcuts',
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-s': () => {
        window.dispatchEvent(new Event('office-save-request'));
        return true;
      },
    };
  },
});

export function Editor({ content, onChange }: {
  content: ${contentType};
  onChange: (content: ${contentType}) => void;
}) {
  const extensions = useMemo(() => [hostShortcuts], []);

  return (
    <${component}
      content={content}
      extensions={extensions}
      onChange={onChange}
    />
  );
}`;
}

function documentSelectionMenuExample(): string {
  return `import type {
  DocumentContent,
  GetDocumentSelectionMenuItems,
} from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';

const getSelectionMenuItems: GetDocumentSelectionMenuItems = () => [
  {
    id: 'expand',
    label: '扩写',
    icon: 'sparkles',
    onSelect: async (context) => {
      const result = await llm.rewrite({
        task: 'expand',
        selectedText: context.selection.text,
        before: context.selection.beforeText,
        after: context.selection.afterText,
        document: context.document.content,
        documentText: context.document.text,
      });

      const applied = context.commands.replaceText(result.text);
      if (!applied.applied) handleStaleSelection(applied.reason);
    },
  },
  {
    id: 'polish',
    label: '润色',
    icon: 'wand',
    onSelect: async (context) => {
      const result = await llm.rewrite({
        task: 'polish',
        selectedText: context.selection.text,
        document: context.document.content,
      });
      context.commands.replaceText(result.text);
    },
  },
];

<DocumentEditor
  content={content as DocumentContent}
  getSelectionMenuItems={getSelectionMenuItems}
  onChange={setContent}
/>;`;
}

function hostExtensionExample(
  component: 'PresentationEditor' | 'SpreadsheetEditor',
  contentType: 'PresentationContent' | 'SpreadsheetContent',
): string {
  return `import type {
  EditorAgentRequest,
  ${contentType},
} from '@a3s-lab/office/core';
import type { OfficeFileAction } from '@a3s-lab/office/react';
import { ${component} } from '@a3s-lab/office/react';

const fileActions: OfficeFileAction[] = [{
  id: 'save-version',
  label: '保存版本',
  onSelect: async () => saveCurrentVersion(),
}];

<${component}
  content={content}
  fileActions={fileActions}
  onAgentRequest={(request: EditorAgentRequest) =>
    runHostAgent(request)
  }
  onChange={setContent}
/>;
`;
}

function pdfHostExample(): string {
  return `import { PdfViewer } from '@a3s-lab/office/react';

<PdfViewer
  fileName="report.pdf"
  sourceKey={fileVersion}
  loadSource={() => storage.download(fileId)}
  onSave={async (pdf) => {
    await storage.upload(fileId, pdf);
    return true;
  }}
/>;
`;
}
