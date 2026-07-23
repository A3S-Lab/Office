import {
  ArrowRight,
  CheckCircle2,
  Download,
  ExternalLink,
  FileArchive,
  FileCode2,
  FileText,
  Github,
  Package,
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

const requestExample =
  '使用 $a3s-office 检查这份季度报告，修正文档中的年份，并验证输出。';

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
        eyebrow="A3S Office"
        title="CLI Skill"
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
              下载
            </a>
          </>
        }
      />

      <div className="playground-doc-content">
        <section
          className="playground-doc-group"
          aria-labelledby="skill-download-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="skill-download-title">下载</h2>
              <span>Office CLI 的操作说明与格式参考</span>
            </div>
          </div>
          <div className="playground-doc-card playground-skill-download">
            <span className="playground-doc-card-icon">
              <FileArchive size={20} />
            </span>
            <div>
              <h3>a3s-office-skill.tar.gz</h3>
              <p>包含主说明、工具元数据以及 Word、表格、演示和 MCP 参考。</p>
              <div className="playground-inline-meta">
                <span>来源：Office 仓库</span>
                <span>目录：a3s-office/</span>
              </div>
            </div>
            <a
              className="playground-primary-button"
              href={skillDownloadUrl}
              download="a3s-office-skill.tar.gz"
            >
              <Download size={15} />
              下载 .tar.gz
            </a>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="skill-install-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="skill-install-title">安装</h2>
              <span>下载 Skill，或直接使用 CLI 内置版本</span>
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
              <CodeBlock code={installSkillCommand} />
              <CodeBlock code={verifySkillCommand} label="确认安装" />
            </div>
            <div className="playground-doc-card playground-usage-card">
              <div className="playground-card-heading">
                <span>
                  <Package size={17} />
                </span>
                <div>
                  <h3>随 CLI 使用</h3>
                  <p>Office CLI 已包含同一份 Skill</p>
                </div>
              </div>
              <CodeBlock code={bundledSkillCommand} />
            </div>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="skill-package-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="skill-package-title">包含内容</h2>
              <span>下载包与仓库中的 Skill 源目录保持一致</span>
            </div>
          </div>
          <div className="playground-doc-card playground-skill-tree">
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
              <small>MCP 接入参考</small>
            </div>
          </div>
        </section>

        <section
          className="playground-doc-group"
          aria-labelledby="skill-use-title"
        >
          <div className="playground-section-heading">
            <div>
              <h2 id="skill-use-title">使用</h2>
              <span>在请求中直接写明 Skill 名称</span>
            </div>
          </div>
          <div className="playground-doc-card playground-skill-usage">
            <div className="playground-skill-usage-heading">
              <span>
                <CheckCircle2 size={17} />
              </span>
              <div>
                <h3>安装完成后即可调用</h3>
                <p>给出文件、要修改的内容和期望的输出位置。</p>
              </div>
            </div>
            <CodeBlock code={requestExample} label="示例请求" />
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
                <Github size={14} />在 GitHub 查看
                <ArrowRight size={13} />
              </a>
            </div>
          </div>
        </section>
      </div>
    </article>
  );
}
