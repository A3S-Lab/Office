import {
  ArrowRight,
  Braces,
  CheckCircle2,
  Eye,
  FileSearch,
  Github,
  Package,
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
        eyebrow="A3S Office"
        title="Office CLI"
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
            <button
              type="button"
              className="playground-primary-button"
              onClick={onOpenSkill}
            >
              <Package size={15} />
              CLI Skill
            </button>
          </>
        }
      />

      <div className="playground-doc-content">
        <section
          className="playground-doc-group"
          aria-labelledby="cli-start-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="cli-start-title">开始使用</h2>
              <span>在本地读取、修改和验证 Office 文件</span>
            </div>
          </div>
          <div className="playground-cli-start-grid">
            <div className="playground-doc-card playground-cli-summary">
              <span className="playground-doc-card-icon">
                <SquareTerminal size={20} />
              </span>
              <div>
                <h3>本地命令行工具</h3>
                <p>
                  支持 Word、Excel 和 PowerPoint。文件只在本机处理，无需启动桌面
                  Office。
                </p>
                <ul
                  className="playground-inline-meta"
                  aria-label="支持的文件格式"
                >
                  <li>Word · .docx</li>
                  <li>Excel · .xlsx</li>
                  <li>PowerPoint · .pptx</li>
                </ul>
              </div>
            </div>
            <div className="playground-doc-card playground-install-card">
              <div className="playground-card-heading">
                <div>
                  <h3>安装</h3>
                  <p>需要 Rust 1.85 或更高版本</p>
                </div>
              </div>
              <CodeBlock code={installCommand} label="从 GitHub 安装" />
              <CodeBlock code="a3s-office --version" label="确认安装" />
            </div>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="cli-capabilities-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="cli-capabilities-title">常用任务</h2>
              <span>按要完成的操作选择命令</span>
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
              title="修改内容"
              description="修改文字、格式、表格、公式、幻灯片与常用结构。"
              commands="set · add · remove · batch"
            />
            <CapabilityCard
              icon={<Eye size={18} />}
              title="预览与监听"
              description="生成预览，并在文件保存后自动刷新本机页面。"
              commands="view · screenshot · watch"
            />
            <CapabilityCard
              icon={<Braces size={18} />}
              title="自动化接入"
              description="通过 JSON 输出、MCP 和 CLI Skill 接入现有流程。"
              commands="--json · mcp · skills"
            />
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="cli-workflow-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="cli-workflow-title">推荐流程</h2>
              <span>先检查，再修改，最后验证结果</span>
            </div>
          </div>
          <div className="playground-doc-card playground-workflow">
            <ol>
              <li>
                <span>1</span>
                <div>
                  <strong>检查文件</strong>
                  <small>确认文件完整，并查看是否存在不支持的结构。</small>
                </div>
              </li>
              <li>
                <span>2</span>
                <div>
                  <strong>定位内容</strong>
                  <small>使用 outline、get 或 query 找到准确位置。</small>
                </div>
              </li>
              <li>
                <span>3</span>
                <div>
                  <strong>执行修改</strong>
                  <small>限定修改范围，重要文件建议输出到新路径。</small>
                </div>
              </li>
              <li>
                <span>4</span>
                <div>
                  <strong>验证结果</strong>
                  <small>重新读取目标位置，并再次运行检查。</small>
                </div>
              </li>
            </ol>
            <CodeBlock code={quickStartCommand} />
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="cli-more-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="cli-more-title">更多用法</h2>
              <span>需要时再启用预览或标准工具接口</span>
            </div>
          </div>
          <div className="playground-doc-split">
            <div className="playground-doc-card playground-usage-card">
              <div className="playground-card-heading">
                <span>
                  <Eye size={17} />
                </span>
                <div>
                  <h3>实时预览</h3>
                  <p>前台运行，按 Ctrl+C 停止</p>
                </div>
              </div>
              <CodeBlock code={previewCommand} />
            </div>
            <div className="playground-doc-card playground-usage-card">
              <div className="playground-card-heading">
                <span>
                  <Braces size={17} />
                </span>
                <div>
                  <h3>工具接口（MCP）</h3>
                  <p>通过标准输入输出提供 Office 工具</p>
                </div>
              </div>
              <CodeBlock code={mcpCommand} />
            </div>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="cli-formats-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="cli-formats-title">支持格式</h2>
              <span>完整能力以命令参考为准</span>
            </div>
          </div>
          <div className="playground-format-strip">
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
                <strong>Excel</strong>
                <small>.xlsx</small>
              </span>
            </div>
            <div>
              <CheckCircle2 size={17} />
              <span>
                <strong>PowerPoint</strong>
                <small>.pptx</small>
              </span>
            </div>
            <a
              href="https://github.com/A3S-Lab/Office/blob/main/docs/cli-reference.md"
              target="_blank"
              rel="noreferrer"
            >
              查看命令参考
              <ArrowRight size={14} />
            </a>
          </div>
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
    <article className="playground-capability-card">
      <span>{icon}</span>
      <div>
        <h3>{title}</h3>
        <p>{description}</p>
        <code>{commands}</code>
      </div>
    </article>
  );
}
