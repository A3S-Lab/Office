import {
  Braces,
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
import { type KeyboardEvent, useEffect, useState } from 'react';
import { CodeBlock, type PlaygroundCodeLanguage } from './code-block';
import { PageHeader } from './page-header';

type Framework = 'react' | 'vue' | 'web-component';
type GuideSection = 'components' | 'cli' | 'skill';

const guideSections = [
  { id: 'components', label: '前端组件', icon: CodeXml },
  { id: 'cli', label: 'Office CLI', icon: SquareTerminal },
  { id: 'skill', label: 'CLI Skill', icon: Package },
] as const;

const installCommand = 'bun add @a3s-lab/office';
const cliInstallCommand = `cargo install \\
  --git https://github.com/A3S-Lab/Office.git \\
  --locked a3s-office-cli`;
const cliQuickStartCommand = `a3s-office validate report.docx --json
a3s-office view report.docx outline --json
a3s-office set report.docx /body --find Draft --replace Final --json`;
const cliPreviewCommand = `a3s-office watch report.docx --port 0
# 保存文件后，预览页面会自动刷新`;
const cliMcpCommand = `a3s-office mcp
# 通过标准输入输出提供有类型的 Office 工具`;
const installSkillCommand = `mkdir -p "\${CODEX_HOME:-$HOME/.codex}/skills"
tar -xzf a3s-office-skill.tar.gz \\
  -C "\${CODEX_HOME:-$HOME/.codex}/skills"`;
const bundledSkillCommand = `a3s-office skills list
a3s-office skills get a3s-office`;
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
  const [section, setSection] = useState<GuideSection>(readGuideSection);

  useEffect(() => {
    const syncSection = () => {
      const nextSection = readGuideSection();
      setSection(nextSection);
      if (
        window.location.hash === '#cli' ||
        window.location.hash === '#skill'
      ) {
        window.history.replaceState(null, '', guideSectionHash(nextSection));
      }
    };
    window.addEventListener('hashchange', syncSection);
    window.addEventListener('popstate', syncSection);
    syncSection();

    return () => {
      window.removeEventListener('hashchange', syncSection);
      window.removeEventListener('popstate', syncSection);
    };
  }, []);

  const openSection = (nextSection: GuideSection) => {
    setSection(nextSection);
    const nextHash = guideSectionHash(nextSection);
    if (window.location.hash !== nextHash) {
      window.history.pushState(null, '', nextHash);
    }
  };

  return (
    <article className="playground-doc-page">
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
        <div
          className="playground-integration-nav"
          role="tablist"
          aria-label="接入内容"
        >
          {guideSections.map((item, index) => {
            const Icon = item.icon;
            return (
              <button
                id={`integration-${item.id}-tab`}
                type="button"
                role="tab"
                key={item.id}
                aria-controls={`integration-${item.id}-panel`}
                aria-selected={section === item.id}
                className={section === item.id ? 'active' : ''}
                onClick={() => openSection(item.id)}
                onKeyDown={(event) =>
                  moveGuideTabFocus(event, index, openSection)
                }
              >
                <Icon size={15} />
                {item.label}
              </button>
            );
          })}
        </div>

        {section === 'components' && <ComponentGuide />}
        {section === 'cli' && <CliGuide />}
        {section === 'skill' && (
          <SkillGuide
            rawSkillUrl={rawSkillUrl}
            skillDownloadUrl={skillDownloadUrl}
          />
        )}
      </div>
    </article>
  );
}

function ComponentGuide() {
  const [framework, setFramework] = useState<Framework>('react');
  const example = frameworkExamples[framework];

  return (
    <div
      id="integration-components-panel"
      className="playground-integration-panel"
      role="tabpanel"
      aria-labelledby="integration-components-tab"
    >
      <section
        className="playground-doc-group"
        aria-labelledby="integration-install-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="integration-install-title">安装组件</h2>
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
            <h2 id="integration-example-title">最小示例</h2>
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
            <h2 id="integration-editors-title">编辑器组件</h2>
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
    </div>
  );
}

function CliGuide() {
  return (
    <div
      id="integration-cli-panel"
      className="playground-integration-panel"
      role="tabpanel"
      aria-labelledby="integration-cli-tab"
    >
      <section
        className="playground-doc-group"
        aria-labelledby="integration-cli-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="integration-cli-title">Office CLI</h2>
            <span>在本机读取、修改和验证 Word、Excel 与 PowerPoint</span>
          </div>
        </div>
        <div className="playground-doc-split">
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <SquareTerminal size={17} />
              </span>
              <div>
                <h3>安装</h3>
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
                <Check size={17} />
              </span>
              <div>
                <h3>常用命令</h3>
                <p>先检查和定位，再修改文件</p>
              </div>
            </div>
            <CodeBlock code={cliQuickStartCommand} language="bash" />
          </div>
        </div>
      </section>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-cli-automation-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="integration-cli-automation-title">预览与自动化</h2>
            <span>需要时再启动，不增加默认流程</span>
          </div>
        </div>
        <div className="playground-doc-split">
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <PanelsTopLeft size={17} />
              </span>
              <div>
                <h3>实时预览</h3>
                <p>保存文件后自动刷新，按 Ctrl+C 停止</p>
              </div>
            </div>
            <CodeBlock code={cliPreviewCommand} language="bash" />
          </div>
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <Braces size={17} />
              </span>
              <div>
                <h3>工具接口</h3>
                <p>通过 MCP 接入兼容客户端</p>
              </div>
            </div>
            <CodeBlock code={cliMcpCommand} language="bash" />
          </div>
        </div>
      </section>

      <div className="playground-format-strip">
        <span>支持 .docx、.xlsx、.pptx</span>
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
  );
}

function SkillGuide({
  skillDownloadUrl,
  rawSkillUrl,
}: {
  skillDownloadUrl: string;
  rawSkillUrl: string;
}) {
  return (
    <div
      id="integration-skill-panel"
      className="playground-integration-panel"
      role="tabpanel"
      aria-labelledby="integration-skill-tab"
    >
      <section
        className="playground-doc-group"
        aria-labelledby="integration-skill-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="integration-skill-title">CLI Skill</h2>
            <span>让 Codex 按 Office CLI 的约定处理文件</span>
          </div>
        </div>
        <div className="playground-doc-card playground-skill-download">
          <span className="playground-doc-card-icon">
            <Package size={20} />
          </span>
          <div>
            <h3>a3s-office-skill.tar.gz</h3>
            <p>包含使用说明，以及 Word、表格、演示和工具接口参考。</p>
          </div>
          <a
            className="playground-primary-button"
            href={skillDownloadUrl}
            download="a3s-office-skill.tar.gz"
          >
            <Download size={15} />
            下载 CLI Skill
          </a>
        </div>
      </section>

      <section
        className="playground-doc-group"
        aria-labelledby="integration-skill-install-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="integration-skill-install-title">安装与使用</h2>
            <span>手动安装，或直接读取 CLI 内置版本</span>
          </div>
        </div>
        <div className="playground-doc-split">
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <Download size={17} />
              </span>
              <div>
                <h3>手动安装</h3>
                <p>解压到个人 Skills 目录</p>
              </div>
            </div>
            <CodeBlock code={installSkillCommand} language="bash" />
          </div>
          <div className="playground-doc-card playground-usage-card">
            <div className="playground-card-heading">
              <span>
                <SquareTerminal size={17} />
              </span>
              <div>
                <h3>通过 CLI 获取</h3>
                <p>Office CLI 内置同一份 Skill</p>
              </div>
            </div>
            <CodeBlock code={bundledSkillCommand} language="bash" />
          </div>
        </div>
      </section>

      <div className="playground-doc-card playground-skill-usage">
        <div className="playground-skill-usage-heading">
          <span>
            <Check size={17} />
          </span>
          <div>
            <h3>调用 Skill</h3>
            <p>在请求中写明 Skill 名称、文件和期望结果。</p>
          </div>
        </div>
        <CodeBlock
          code={skillRequestExample}
          label="示例请求"
          language="text"
        />
        <div className="playground-skill-links">
          <a href={rawSkillUrl} target="_blank" rel="noreferrer">
            <FileText size={14} />
            查看 SKILL.md
            <ExternalLink size={12} />
          </a>
          <a
            href="https://github.com/A3S-Lab/Office/tree/main/crates/cli/skills/a3s-office"
            target="_blank"
            rel="noreferrer"
          >
            <Github size={14} />
            查看源文件
            <ExternalLink size={12} />
          </a>
        </div>
      </div>
    </div>
  );
}

function readGuideSection(): GuideSection {
  const hash = window.location.hash.slice(1);
  if (hash === 'cli' || hash === 'guide/cli') return 'cli';
  if (hash === 'skill' || hash === 'guide/skill') return 'skill';
  return 'components';
}

function guideSectionHash(section: GuideSection): string {
  return section === 'components' ? '#guide' : `#guide/${section}`;
}

function moveGuideTabFocus(
  event: KeyboardEvent<HTMLButtonElement>,
  currentIndex: number,
  openSection: (section: GuideSection) => void,
) {
  const lastIndex = guideSections.length - 1;
  let nextIndex: number | null = null;
  if (event.key === 'ArrowRight')
    nextIndex = (currentIndex + 1) % guideSections.length;
  if (event.key === 'ArrowLeft')
    nextIndex = (currentIndex + lastIndex) % guideSections.length;
  if (event.key === 'Home') nextIndex = 0;
  if (event.key === 'End') nextIndex = lastIndex;
  if (nextIndex === null) return;

  event.preventDefault();
  const nextSection = guideSections[nextIndex];
  openSection(nextSection.id);
  document.getElementById(`integration-${nextSection.id}-tab`)?.focus();
}
