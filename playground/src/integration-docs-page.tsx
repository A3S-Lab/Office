import {
  Check,
  CodeXml,
  Github,
  PackageCheck,
  PanelsTopLeft,
} from 'lucide-react';
import { useState } from 'react';
import { CodeBlock } from './code-block';
import { PageHeader } from './page-header';

type Framework = 'react' | 'vue' | 'web-component';

const installCommand = 'bun add @a3s-lab/office';

const frameworkExamples: Record<
  Framework,
  { label: string; fileName: string; code: string }
> = {
  react: {
    label: 'React',
    fileName: 'DocumentPage.tsx',
    code: `import { useState } from 'react';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

const initialDocument: DocumentContent = {
  type: 'document',
  html: '<h1>项目方案</h1><p>从这里开始编辑。</p>',
  pageSize: 'a4',
};

export function DocumentPage() {
  const [content, setContent] = useState(initialDocument);

  return (
    <div style={{ height: '100vh', minHeight: 0 }}>
      <DocumentEditor
        content={content}
        onChange={setContent}
        theme="light"
      />
    </div>
  );
}`,
  },
  vue: {
    label: 'Vue',
    fileName: 'DocumentPage.vue',
    code: `<script setup lang="ts">
import { ref } from 'vue';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/vue';
import '@a3s-lab/office/styles.css';

const content = ref<DocumentContent>({
  type: 'document',
  html: '<h1>项目方案</h1><p>从这里开始编辑。</p>',
  pageSize: 'a4',
});
</script>

<template>
  <div class="editor-host">
    <DocumentEditor v-model:content="content" theme="light" />
  </div>
</template>

<style scoped>
.editor-host {
  height: 100vh;
  min-height: 0;
}
</style>`,
  },
  'web-component': {
    label: 'Web Component',
    fileName: 'document-editor.ts',
    code: `import {
  defineA3SOfficeElements,
} from '@a3s-lab/office/web-component';
import '@a3s-lab/office/styles.css';

defineA3SOfficeElements();

const editor = document.querySelector('a3s-document-editor');

editor.content = {
  type: 'document',
  html: '<h1>项目方案</h1><p>从这里开始编辑。</p>',
  pageSize: 'a4',
};

editor.addEventListener('change', (event) => {
  console.log(event.detail);
});`,
  },
};

const editors = [
  ['DocumentEditor', '文字', '@a3s-lab/office/react'],
  ['MarkdownEditor', 'Markdown', '@a3s-lab/office/react'],
  ['SpreadsheetEditor', '表格', '@a3s-lab/office/react'],
  ['PresentationEditor', '演示', '@a3s-lab/office/react'],
  ['PdfViewer', 'PDF', '@a3s-lab/office/react'],
] as const;

export function IntegrationDocsPage({
  sidebarOpen,
  onOpenSidebar,
}: {
  sidebarOpen: boolean;
  onOpenSidebar: () => void;
}) {
  const [framework, setFramework] = useState<Framework>('react');
  const example = frameworkExamples[framework];

  return (
    <article className="playground-doc-page">
      <PageHeader
        eyebrow="A3S Office"
        title="组件接入"
        sidebarOpen={sidebarOpen}
        onOpenSidebar={onOpenSidebar}
        actions={
          <>
            <a
              className="playground-secondary-button"
              href="https://github.com/A3S-Lab/Office"
              target="_blank"
              rel="noreferrer"
            >
              <Github size={15} />
              源代码
            </a>
            <a
              className="playground-primary-button"
              href="https://www.npmjs.com/package/@a3s-lab/office"
              target="_blank"
              rel="noreferrer"
            >
              <PackageCheck size={15} />
              npm 包
            </a>
          </>
        }
      />

      <div className="playground-doc-content">
        <section
          className="playground-doc-group"
          aria-labelledby="integration-start-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="integration-start-title">安装</h2>
              <span>一个包提供五类编辑器和三种前端入口</span>
            </div>
          </div>
          <div className="playground-integration-start">
            <div className="playground-doc-card playground-install-card">
              <CodeBlock code={installCommand} label="项目目录" />
            </div>
            <div className="playground-doc-card playground-integration-notes">
              <span>
                <Check size={14} />
                引入一次全局样式
              </span>
              <span>
                <Check size={14} />
                编辑器容器必须有明确高度
              </span>
              <span>
                <Check size={14} />
                内容由宿主项目持有
              </span>
            </div>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="integration-example-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="integration-example-title">最小示例</h2>
              <span>复制后即可放进现有页面</span>
            </div>
          </div>
          <div className="playground-doc-card playground-framework-example">
            <div
              className="playground-framework-tabs"
              role="tablist"
              aria-label="前端框架"
            >
              {(Object.keys(frameworkExamples) as Framework[]).map((id) => (
                <button
                  type="button"
                  role="tab"
                  key={id}
                  aria-selected={framework === id}
                  className={framework === id ? 'active' : ''}
                  onClick={() => setFramework(id)}
                >
                  {frameworkExamples[id].label}
                </button>
              ))}
              <span>{example.fileName}</span>
            </div>
            <CodeBlock code={example.code} />
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="integration-editors-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="integration-editors-title">编辑器入口</h2>
              <span>按文件类型只加载需要的编辑器</span>
            </div>
          </div>
          <div className="playground-editor-entry-grid">
            {editors.map(([component, label, entry]) => (
              <article className="playground-doc-card" key={component}>
                <span>
                  {component === 'PdfViewer' ? (
                    <PanelsTopLeft size={16} />
                  ) : (
                    <CodeXml size={16} />
                  )}
                </span>
                <div>
                  <strong>{label}</strong>
                  <code>{component}</code>
                  <small>{entry}</small>
                </div>
              </article>
            ))}
          </div>
        </section>
      </div>
    </article>
  );
}
