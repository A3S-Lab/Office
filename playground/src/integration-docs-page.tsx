import {
  Check,
  CodeXml,
  Download,
  ExternalLink,
  FileText,
  Github,
  Package,
  PackageCheck,
  PanelsTopLeft,
  SquareTerminal,
} from 'lucide-react';
import { useEffect, useState } from 'react';
import { CodeBlock, type PlaygroundCodeLanguage } from './code-block';
import { PageHeader } from './page-header';

type Framework = 'react' | 'vue' | 'web-component';

const installCommand = 'bun add @a3s-lab/office';
const cliInstallCommand = `cargo install \\
  --git https://github.com/A3S-Lab/Office.git \\
  --locked a3s-office-cli`;
const cliQuickStartCommand = `a3s-office validate report.docx --json
a3s-office view report.docx outline --json
a3s-office set report.docx /body --find Draft --replace Final --json`;
const installSkillCommand = `mkdir -p "\${CODEX_HOME:-$HOME/.codex}/skills"
tar -xzf a3s-office-skill.tar.gz \\
  -C "\${CODEX_HOME:-$HOME/.codex}/skills"`;
const skillRequestExample =
  '使用 $a3s-office 检查这份季度报告，修正文档中的年份，并验证输出。';

const frameworkExamples: Record<
  Framework,
  {
    label: string;
    fileName: string;
    language: PlaygroundCodeLanguage;
    code: string;
  }
> = {
  react: {
    label: 'React',
    fileName: 'DocumentPage.tsx',
    language: 'tsx',
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
    language: 'markup',
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
    language: 'typescript',
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
  ['DocumentEditor', '文字'],
  ['MarkdownEditor', 'Markdown'],
  ['SpreadsheetEditor', '表格'],
  ['PresentationEditor', '演示'],
  ['PdfViewer', 'PDF'],
] as const;

export function IntegrationDocsPage({
  sidebarOpen,
  skillDownloadUrl,
  rawSkillUrl,
  onOpenSidebar,
}: {
  sidebarOpen: boolean;
  skillDownloadUrl: string;
  rawSkillUrl: string;
  onOpenSidebar: () => void;
}) {
  useEffect(() => {
    const syncGuidePosition = () => {
      let hash = window.location.hash;
      if (isLegacyAutomationHash(hash)) {
        hash = '#guide/automation';
        window.history.replaceState(null, '', hash);
      }
      if (hash === '#guide') {
        const page = document.querySelector<HTMLElement>(
          '.playground-doc-page',
        );
        if (page) page.scrollTop = 0;
        return;
      }
      if (hash === '#guide/components' || hash === '#guide/automation') {
        document
          .getElementById(hash.slice(1))
          ?.scrollIntoView?.({ block: 'start' });
      }
    };

    window.addEventListener('hashchange', syncGuidePosition);
    window.addEventListener('popstate', syncGuidePosition);
    syncGuidePosition();

    return () => {
      window.removeEventListener('hashchange', syncGuidePosition);
      window.removeEventListener('popstate', syncGuidePosition);
    };
  }, []);

  return (
    <article id="guide" className="playground-doc-page">
      <PageHeader
        eyebrow="A3S Office"
        title="接入文档"
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
        <nav className="playground-integration-nav" aria-label="接入方式">
          <a href="#guide/components">
            <CodeXml size={15} />
            前端组件
          </a>
          <a href="#guide/automation">
            <SquareTerminal size={15} />
            命令行与 AI
          </a>
        </nav>

        <ComponentGuide />
        <AutomationGuide
          rawSkillUrl={rawSkillUrl}
          skillDownloadUrl={skillDownloadUrl}
        />
      </div>
    </article>
  );
}

function ComponentGuide() {
  const [framework, setFramework] = useState<Framework>('react');
  const example = frameworkExamples[framework];

  return (
    <section
      id="guide/components"
      className="playground-integration-panel playground-guide-section"
      aria-labelledby="integration-components-title"
    >
      <header className="playground-guide-section-heading">
        <span>
          <CodeXml size={18} />
        </span>
        <div>
          <h2 id="integration-components-title">前端组件</h2>
          <p>把编辑器嵌入现有的 React、Vue 或 Web Component 项目。</p>
        </div>
      </header>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-install-title"
      >
        <div className="playground-section-heading">
          <div>
            <h3 id="integration-install-title">安装</h3>
            <span>同一个包支持 React、Vue 和 Web Component</span>
          </div>
        </div>
        <div className="playground-integration-start">
          <div className="playground-doc-card playground-install-card">
            <CodeBlock code={installCommand} label="项目目录" language="bash" />
          </div>
          <div className="playground-doc-card playground-integration-notes">
            <span>
              <Check size={14} />
              引入一次全局样式
            </span>
            <span>
              <Check size={14} />
              容器需要明确高度
            </span>
            <span>
              <Check size={14} />
              表格计算资源默认随包加载
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
            <h3 id="integration-example-title">最小示例</h3>
            <span>内容与保存逻辑由宿主项目持有</span>
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
          <CodeBlock code={example.code} language={example.language} />
        </div>
      </section>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-editors-title"
      >
        <div className="playground-section-heading">
          <div>
            <h3 id="integration-editors-title">可用组件</h3>
            <span>从对应框架入口按需引入</span>
          </div>
        </div>
        <div className="playground-editor-entry-grid">
          {editors.map(([component, label]) => (
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
                <small>@a3s-lab/office/react</small>
              </div>
            </article>
          ))}
        </div>
      </section>
    </section>
  );
}

function AutomationGuide({
  skillDownloadUrl,
  rawSkillUrl,
}: {
  skillDownloadUrl: string;
  rawSkillUrl: string;
}) {
  return (
    <section
      id="guide/automation"
      className="playground-integration-panel playground-guide-section"
      aria-labelledby="integration-automation-title"
    >
      <header className="playground-guide-section-heading">
        <span>
          <SquareTerminal size={18} />
        </span>
        <div>
          <h2 id="integration-automation-title">命令行与 AI</h2>
          <p>用 Office CLI 处理本地文件，需要时再让 Codex 调用它。</p>
        </div>
      </header>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-tools-title"
      >
        <div className="playground-section-heading">
          <div>
            <h3 id="integration-tools-title">准备工具</h3>
            <span>先安装 CLI；只有使用 Codex 时才需要 Skill</span>
          </div>
        </div>

        <div className="playground-doc-split">
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <SquareTerminal size={17} />
              </span>
              <div>
                <h3>安装 Office CLI</h3>
                <p>需要 Rust 1.85 或更高版本</p>
              </div>
            </div>
            <CodeBlock code={cliInstallCommand} language="bash" />
            <CodeBlock
              code="a3s-office --version"
              label="确认安装"
              language="bash"
            />
          </div>

          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <Package size={17} />
              </span>
              <div>
                <h3>安装 CLI Skill</h3>
                <p>下载后解压到个人 Skills 目录</p>
              </div>
            </div>
            <div className="playground-skill-card-actions">
              <a
                className="playground-primary-button"
                href={skillDownloadUrl}
                download="a3s-office-skill.tar.gz"
              >
                <Download size={15} />
                下载 CLI Skill
              </a>
              <a href={rawSkillUrl} target="_blank" rel="noreferrer">
                <FileText size={14} />
                查看说明
                <ExternalLink size={12} />
              </a>
            </div>
            <CodeBlock code={installSkillCommand} language="bash" />
          </div>
        </div>
      </section>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-usage-title"
      >
        <div className="playground-section-heading">
          <div>
            <h3 id="integration-usage-title">开始使用</h3>
            <span>直接运行命令，或在请求中指定 Skill</span>
          </div>
        </div>

        <div className="playground-doc-split">
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <Check size={17} />
              </span>
              <div>
                <h3>处理文件</h3>
                <p>先检查和定位，再修改文件</p>
              </div>
            </div>
            <CodeBlock code={cliQuickStartCommand} language="bash" />
          </div>

          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <Package size={17} />
              </span>
              <div>
                <h3>在 Codex 中调用</h3>
                <p>写明 Skill、文件和期望结果</p>
              </div>
            </div>
            <CodeBlock
              code={skillRequestExample}
              label="示例请求"
              language="text"
            />
          </div>
        </div>
      </section>

      <div className="playground-format-strip">
        <span>支持 .docx、.xlsx、.pptx</span>
        <div className="playground-guide-links">
          <a href={rawSkillUrl} target="_blank" rel="noreferrer">
            <FileText size={14} />
            查看 SKILL.md
            <ExternalLink size={12} />
          </a>
          <a
            href="https://github.com/A3S-Lab/Office/blob/main/docs/cli-reference.md"
            target="_blank"
            rel="noreferrer"
          >
            完整命令参考
            <ExternalLink size={12} />
          </a>
        </div>
      </div>
    </section>
  );
}

function isLegacyAutomationHash(hash: string): boolean {
  return (
    hash === '#cli' ||
    hash === '#skill' ||
    hash === '#guide/cli' ||
    hash === '#guide/skill'
  );
}
