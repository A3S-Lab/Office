import { Check, Copy } from 'lucide-react';
import { Highlight, themes, type Token } from 'prism-react-renderer';
import { useEffect, useState } from 'react';

export type PlaygroundCodeLanguage =
  | 'bash'
  | 'markup'
  | 'text'
  | 'tsx'
  | 'typescript';

export function CodeBlock({
  code,
  label,
  language = 'text',
}: {
  code: string;
  label?: string;
  language?: PlaygroundCodeLanguage;
}) {
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!copied) return;
    const timeout = window.setTimeout(() => setCopied(false), 1800);
    return () => window.clearTimeout(timeout);
  }, [copied]);

  return (
    <div className="playground-code-block">
      {label && <span>{label}</span>}
      <Highlight code={code} language={language} theme={themes.oneDark}>
        {({ className, getLineProps, getTokenProps, style, tokens }) => {
          const lines = keyedCodeLines(tokens);
          return (
            <pre
              className={className}
              data-code-language={language}
              style={{ ...style, background: 'transparent' }}
            >
              <code>
                {lines.map((line, lineIndex) => (
                  <span
                    {...getLineProps({ line: line.value })}
                    className="playground-code-line"
                    key={line.key}
                  >
                    {line.tokens.map(({ key, value }) => (
                      <span {...getTokenProps({ token: value })} key={key} />
                    ))}
                    {lineIndex < lines.length - 1 ? '\n' : null}
                  </span>
                ))}
              </code>
            </pre>
          );
        }}
      </Highlight>
      <button
        type="button"
        aria-label={copied ? '已复制代码' : '复制代码'}
        title={copied ? '已复制' : '复制'}
        onClick={() => {
          void copyText(code).then(() => setCopied(true));
        }}
      >
        {copied ? <Check size={14} /> : <Copy size={14} />}
        <span>{copied ? '已复制' : '复制'}</span>
      </button>
    </div>
  );
}

function keyedCodeLines(tokens: Token[][]) {
  const lineOccurrences = new Map<string, number>();
  return tokens.map((line) => {
    const signature = line
      .map((token) => `${token.types.join('.')}:${token.content}`)
      .join('|');
    const occurrence = lineOccurrences.get(signature) ?? 0;
    lineOccurrences.set(signature, occurrence + 1);
    const tokenOccurrences = new Map<string, number>();
    return {
      key: `${signature}:${occurrence}`,
      value: line,
      tokens: line.map((token) => {
        const tokenSignature = `${token.types.join('.')}:${token.content}`;
        const tokenOccurrence = tokenOccurrences.get(tokenSignature) ?? 0;
        tokenOccurrences.set(tokenSignature, tokenOccurrence + 1);
        return {
          key: `${tokenSignature}:${tokenOccurrence}`,
          value: token,
        };
      }),
    };
  });
}

async function copyText(value: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }
  const textarea = document.createElement('textarea');
  textarea.value = value;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.append(textarea);
  textarea.select();
  document.execCommand('copy');
  textarea.remove();
}
