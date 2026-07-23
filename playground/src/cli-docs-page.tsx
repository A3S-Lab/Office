import {
  ArrowRight,
  Bot,
  Braces,
  CheckCircle2,
  ExternalLink,
  Eye,
  FileSearch,
  Github,
  Layers3,
  Package,
  Play,
  ShieldCheck,
  SquareTerminal,
  WandSparkles,
} from 'lucide-react';
import { CodeBlock } from './code-block';
import { PageHeader } from './page-header';

const installCommand =
  'cargo install --git https://github.com/A3S-Lab/Office.git --locked a3s-office-cli';

const quickStartCommand = `a3s-office validate report.docx --json
a3s-office view report.docx outline --json
a3s-office set report.docx /body --find Draft --replace Final --json
a3s-office view report.docx issues --json`;

const previewCommand = `a3s-office watch report.docx --port 0
# 打开命令输出的本机预览地址，保存文件后自动刷新`;

const mcpCommand = `a3s-office mcp
# 通过标准输入输出提供有类型的 Office 工具`;

export function CliDocsPage({
  sidebarOpen,
  onOpenSidebar,
  onOpenSkill,
}: {
  sidebarOpen: boolean;
  onOpenSidebar: () => void;
  onOpenSkill: () => void;
}) {
  return (
    <article className="playground-doc-page">
      <PageHeader
        eyebrow="开发者工具"
        title="A3S Office CLI"
        description="在本地检查、创建、修改和验证 Word、Excel 与 PowerPoint 文件。无需启动桌面 Office，也不会把文件上传到服务器。"
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
              <ExternalLink size={12} />
            </a>
            <button
              type="button"
              className="playground-primary-button"
              onClick={onOpenSkill}
            >
              <Package size={15} />
              下载 CLI Skill
            </button>
          </>
        }
      />

      <div className="playground-doc-content">
        <section className="playground-cli-hero">
          <div>
            <span className="playground-feature-badge">
              <ShieldCheck size={14} />
              Rust 原生 · 本地执行
            </span>
            <h2>让人和 AI 都能可靠地操作 Office 文件</h2>
            <p>
              同一套能力同时提供命令行、标准 MCP 和 Agent
              Skill。所有修改都经过文件格式检查，无法安全处理的操作会明确失败。
            </p>
          </div>
          <div className="playground-terminal-card">
            <header>
              <span />
              <span />
              <span />
              <strong>Terminal</strong>
            </header>
            <pre>
              <code>
                <span>$</span> a3s-office view report.docx outline --json
                {'\n'}
                <i>{'{'}</i>
                {'\n  '}
                <b>"ok"</b>: true,
                {'\n  '}
                <b>"headings"</b>: 6{'\n'}
                <i>{'}'}</i>
              </code>
            </pre>
          </div>
        </section>

        <section
          className="playground-doc-section"
          aria-labelledby="cli-install-title"
        >
          <div className="playground-doc-section-heading">
            <span>
              <SquareTerminal size={17} />
            </span>
            <div>
              <h2 id="cli-install-title">安装</h2>
              <p>需要 Rust 1.85 或更高版本。</p>
            </div>
          </div>
          <CodeBlock code={installCommand} label="从 GitHub 安装" />
          <CodeBlock code="a3s-office --version" label="确认安装" />
        </section>

        <section
          className="playground-doc-section"
          aria-labelledby="cli-capabilities-title"
        >
          <div className="playground-doc-section-heading">
            <span>
              <Layers3 size={17} />
            </span>
            <div>
              <h2 id="cli-capabilities-title">核心能力</h2>
              <p>
                从读取、修改到预览与自动化，使用一致的文件路径和 JSON 输出。
              </p>
            </div>
          </div>
          <div className="playground-capability-grid">
            <CapabilityCard
              icon={<FileSearch size={18} />}
              title="读取与检查"
              description="查看正文、结构、统计与兼容性问题，不改动原文件。"
              commands="get · query · view · validate"
            />
            <CapabilityCard
              icon={<WandSparkles size={18} />}
              title="安全修改"
              description="修改文字、格式、表格、公式、幻灯片与常用结构。"
              commands="set · add · remove · batch"
            />
            <CapabilityCard
              icon={<Eye size={18} />}
              title="预览与监听"
              description="生成 HTML、SVG、PNG，并在保存后自动刷新本机预览。"
              commands="view · screenshot · watch"
            />
            <CapabilityCard
              icon={<Bot size={18} />}
              title="Agent 接入"
              description="通过标准 MCP 和 CLI Skill 让 Agent 使用同一套能力。"
              commands="mcp · skills"
            />
          </div>
        </section>

        <section
          className="playground-doc-section"
          aria-labelledby="cli-first-workflow-title"
        >
          <div className="playground-doc-section-heading">
            <span>
              <Play size={17} />
            </span>
            <div>
              <h2 id="cli-first-workflow-title">第一个工作流</h2>
              <p>先检查，再修改，最后验证结果。</p>
            </div>
          </div>
          <div className="playground-workflow">
            <ol>
              <li>
                <span>1</span>
                <div>
                  <strong>确认文件可读</strong>
                  <small>validate 会在修改前发现损坏或不支持的结构。</small>
                </div>
              </li>
              <li>
                <span>2</span>
                <div>
                  <strong>读取目标内容</strong>
                  <small>使用 outline、get 或 query 找到准确路径。</small>
                </div>
              </li>
              <li>
                <span>3</span>
                <div>
                  <strong>执行有范围的修改</strong>
                  <small>保持路径尽可能具体，必要时使用新输出文件。</small>
                </div>
              </li>
              <li>
                <span>4</span>
                <div>
                  <strong>重新检查结果</strong>
                  <small>读取目标位置并再次运行 issues 或 validate。</small>
                </div>
              </li>
            </ol>
            <CodeBlock code={quickStartCommand} />
          </div>
        </section>

        <section className="playground-doc-split">
          <div className="playground-doc-section compact">
            <div className="playground-doc-section-heading">
              <span>
                <Eye size={17} />
              </span>
              <div>
                <h2>实时预览</h2>
                <p>前台运行，按 Ctrl+C 停止。</p>
              </div>
            </div>
            <CodeBlock code={previewCommand} />
          </div>
          <div className="playground-doc-section compact">
            <div className="playground-doc-section-heading">
              <span>
                <Braces size={17} />
              </span>
              <div>
                <h2>标准 MCP</h2>
                <p>为 Agent 提供有类型的会话工具。</p>
              </div>
            </div>
            <CodeBlock code={mcpCommand} />
          </div>
        </section>

        <section className="playground-format-strip" aria-label="支持的格式">
          <div>
            <CheckCircle2 size={17} />
            <span>
              <strong>Word</strong>
              <small>.docx</small>
            </span>
          </div>
          <div>
            <CheckCircle2 size={17} />
            <span>
              <strong>Spreadsheet</strong>
              <small>.xlsx</small>
            </span>
          </div>
          <div>
            <CheckCircle2 size={17} />
            <span>
              <strong>Presentation</strong>
              <small>.pptx</small>
            </span>
          </div>
          <a
            href="https://github.com/A3S-Lab/Office/blob/main/docs/cli-reference.md"
            target="_blank"
            rel="noreferrer"
          >
            查看完整命令参考
            <ArrowRight size={14} />
          </a>
        </section>
      </div>
    </article>
  );
}

function CapabilityCard({
  icon,
  title,
  description,
  commands,
}: {
  icon: React.ReactNode;
  title: string;
  description: string;
  commands: string;
}) {
  return (
    <section>
      <span>{icon}</span>
      <h3>{title}</h3>
      <p>{description}</p>
      <code>{commands}</code>
    </section>
  );
}
