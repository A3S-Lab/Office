import {
  ArrowRight,
  Bot,
  CheckCircle2,
  Download,
  ExternalLink,
  FileArchive,
  FileCode2,
  FileText,
  FolderTree,
  Github,
  Package,
  ShieldCheck,
  Sparkles,
  SquareTerminal,
} from 'lucide-react';
import { CodeBlock } from './code-block';
import { PageHeader } from './page-header';

const installSkillCommand = `mkdir -p "\${CODEX_HOME:-$HOME/.codex}/skills"
tar -xzf a3s-office-skill.tar.gz \\
  -C "\${CODEX_HOME:-$HOME/.codex}/skills"`;

const verifySkillCommand = `test -f "\${CODEX_HOME:-$HOME/.codex}/skills/a3s-office/SKILL.md"
echo "A3S Office Skill installed"`;

const bundledSkillCommand = `a3s-office skills list
a3s-office skills get a3s-office
a3s-office skills path a3s-office`;

export function SkillDownloadPage({
  sidebarOpen,
  skillDownloadUrl,
  rawSkillUrl,
  onOpenSidebar,
  onOpenCli,
}: {
  sidebarOpen: boolean;
  skillDownloadUrl: string;
  rawSkillUrl: string;
  onOpenSidebar: () => void;
  onOpenCli: () => void;
}) {
  return (
    <article className="playground-doc-page skill-page">
      <PageHeader
        eyebrow="Agent 能力"
        title="A3S Office CLI Skill"
        description="让 Codex 和其它支持 Skill 的 Agent 知道如何检查、修改并验证 Office 文件，同时遵守 A3S Office CLI 的安全边界。"
        sidebarOpen={sidebarOpen}
        onOpenSidebar={onOpenSidebar}
        actions={
          <>
            <button
              type="button"
              className="playground-secondary-button"
              onClick={onOpenCli}
            >
              <SquareTerminal size={15} />
              CLI 文档
            </button>
            <a
              className="playground-primary-button"
              href={skillDownloadUrl}
              download="a3s-office-skill.tar.gz"
            >
              <Download size={15} />
              下载 Skill
            </a>
          </>
        }
      />

      <div className="playground-doc-content">
        <section className="playground-skill-hero">
          <div className="playground-skill-identity">
            <span>
              <Sparkles size={27} />
            </span>
            <div>
              <small>SKILL</small>
              <h2>a3s-office</h2>
              <p>Word · Spreadsheet · Presentation · CLI · MCP</p>
            </div>
          </div>
          <div className="playground-skill-download-card">
            <FileArchive size={24} />
            <div>
              <strong>完整 Skill 包</strong>
              <span>SKILL.md、Agent 元数据与四份格式参考</span>
            </div>
            <a href={skillDownloadUrl} download="a3s-office-skill.tar.gz">
              下载 .tar.gz
              <Download size={14} />
            </a>
          </div>
        </section>

        <section className="playground-skill-benefits">
          <div>
            <span>
              <Bot size={18} />
            </span>
            <strong>准确触发</strong>
            <p>在处理 .docx、.xlsx 或 .pptx 时自动提供正确工作流。</p>
          </div>
          <div>
            <span>
              <ShieldCheck size={18} />
            </span>
            <strong>安全修改</strong>
            <p>先检查、再修改、最后验证，不会用脚本绕过格式限制。</p>
          </div>
          <div>
            <span>
              <FolderTree size={18} />
            </span>
            <strong>按需加载</strong>
            <p>只在需要时读取 Word、表格、演示或 MCP 参考。</p>
          </div>
        </section>

        <section className="playground-doc-split skill-install-grid">
          <section
            className="playground-doc-section compact"
            aria-labelledby="skill-install-title"
          >
            <div className="playground-doc-section-heading">
              <span>
                <Download size={17} />
              </span>
              <div>
                <h2 id="skill-install-title">安装到 Codex</h2>
                <p>下载后解压到个人 Skills 目录。</p>
              </div>
            </div>
            <CodeBlock code={installSkillCommand} />
            <CodeBlock code={verifySkillCommand} label="确认安装" />
          </section>
          <section
            className="playground-doc-section compact"
            aria-labelledby="skill-bundled-title"
          >
            <div className="playground-doc-section-heading">
              <span>
                <Package size={17} />
              </span>
              <div>
                <h2 id="skill-bundled-title">CLI 已内置</h2>
                <p>安装 A3S Office CLI 后可直接读取同一份 Skill。</p>
              </div>
            </div>
            <CodeBlock code={bundledSkillCommand} />
          </section>
        </section>

        <section
          className="playground-doc-section"
          aria-labelledby="skill-package-title"
        >
          <div className="playground-doc-section-heading">
            <span>
              <FolderTree size={17} />
            </span>
            <div>
              <h2 id="skill-package-title">下载内容</h2>
              <p>下载包直接来自仓库中的 CLI Skill 源目录。</p>
            </div>
          </div>
          <div className="playground-skill-tree">
            <div>
              <Package size={16} />
              <strong>a3s-office/</strong>
            </div>
            <div>
              <FileText size={15} />
              <span>SKILL.md</span>
              <small>工作流与安全边界</small>
            </div>
            <div>
              <FileCode2 size={15} />
              <span>agents/openai.yaml</span>
              <small>Skill 列表元数据</small>
            </div>
            <div>
              <FileText size={15} />
              <span>references/word.md</span>
              <small>Word 操作参考</small>
            </div>
            <div>
              <FileText size={15} />
              <span>references/spreadsheet.md</span>
              <small>表格操作参考</small>
            </div>
            <div>
              <FileText size={15} />
              <span>references/presentation.md</span>
              <small>演示操作参考</small>
            </div>
            <div>
              <FileText size={15} />
              <span>references/mcp.md</span>
              <small>标准 MCP 会话参考</small>
            </div>
          </div>
        </section>

        <section className="playground-skill-usage">
          <div>
            <span>
              <CheckCircle2 size={17} />
              安装完成后
            </span>
            <h2>直接在请求中调用 Skill</h2>
            <code>
              使用 $a3s-office 检查这份季度报告，修正文档中的年份，并验证输出。
            </code>
          </div>
          <div className="playground-skill-links">
            <a href={rawSkillUrl} target="_blank" rel="noreferrer">
              查看 SKILL.md
              <ExternalLink size={13} />
            </a>
            <a
              href="https://github.com/A3S-Lab/Office/tree/main/crates/cli/skills/a3s-office"
              target="_blank"
              rel="noreferrer"
            >
              <Github size={14} />在 GitHub 查看
              <ArrowRight size={13} />
            </a>
          </div>
        </section>
      </div>
    </article>
  );
}
