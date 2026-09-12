import CodeBlock from '@tiptap/extension-code-block';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import {
  highlightMarkdownCode,
  renderMarkdownDiagram,
  resolveMarkdownCodeLanguage,
} from './markdown-code-render';

const highlightPluginKey = new PluginKey<DecorationSet>(
  'workMarkdownCodeHighlight',
);

export const WorkMarkdownCodeBlock = CodeBlock.extend({
  addNodeView() {
    return ({ node }) => {
      let current = node;
      let generation = 0;
      const dom = document.createElement('div');
      dom.className = 'work-markdown-code';
      const figure = document.createElement('figure');
      figure.className = 'work-markdown-mermaid';
      figure.setAttribute('aria-label', 'Mermaid diagram');
      const pre = document.createElement('pre');
      const code = document.createElement('code');
      pre.append(code);
      dom.append(figure, pre);

      const paint = () => {
        const request = ++generation;
        const language = resolveMarkdownCodeLanguage(current.attrs.language);
        const source = current.textContent;
        if (language !== 'mermaid') {
          figure.replaceChildren();
          figure.hidden = true;
          dom.dataset.markdownDiagram = 'source';
          return;
        }
        const svg = renderMarkdownDiagram(source);
        if (request !== generation) return;
        if (!svg) {
          figure.replaceChildren();
          figure.hidden = true;
          dom.dataset.markdownDiagram = 'source';
          return;
        }
        figure.innerHTML = svg;
        figure.hidden = false;
        dom.dataset.markdownDiagram = 'rendered';
      };

      paint();
      return {
        dom,
        contentDOM: code,
        update(updated) {
          if (updated.type !== current.type) return false;
          current = updated;
          paint();
          return true;
        },
        selectNode() {
          dom.classList.add('is-selected');
        },
        deselectNode() {
          dom.classList.remove('is-selected');
        },
        ignoreMutation(mutation) {
          return figure.contains(mutation.target);
        },
        destroy() {
          generation += 1;
        },
      };
    };
  },

  addProseMirrorPlugins() {
    return [...(this.parent?.() ?? []), markdownHighlightPlugin()];
  },
});

function markdownHighlightPlugin(): Plugin<DecorationSet> {
  return new Plugin({
    key: highlightPluginKey,
    state: {
      init: () => DecorationSet.empty,
      apply(transaction, decorations) {
        const next = transaction.getMeta(highlightPluginKey) as
          | DecorationSet
          | undefined;
        if (next) return next;
        return decorations.map(transaction.mapping, transaction.doc);
      },
    },
    props: {
      decorations(state) {
        return highlightPluginKey.getState(state);
      },
    },
    view(view) {
      let generation = 0;
      const schedule = () => {
        const request = ++generation;
        const doc = view.state.doc;
        void highlightDocument(doc).then((decorations) => {
          if (request !== generation || view.isDestroyed) return;
          if (!view.state.doc.eq(doc)) return;
          const current = highlightPluginKey.getState(view.state);
          if (current?.eq(decorations)) return;
          // Highlighting must not republish Markdown or rewrite the source pane.
          view.dispatch(
            view.state.tr
              .setMeta(highlightPluginKey, decorations)
              .setMeta('addToHistory', false)
              .setMeta('preventUpdate', true),
          );
        });
      };
      schedule();
      return {
        update(next, previous) {
          if (next.state.doc.eq(previous.doc)) return;
          schedule();
        },
        destroy() {
          generation += 1;
        },
      };
    },
  });
}

async function highlightDocument(doc: ProseMirrorNode): Promise<DecorationSet> {
  const decorations: Decoration[] = [];
  const jobs: Array<Promise<void>> = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'codeBlock') return;
    const source = node.textContent;
    jobs.push(
      highlightMarkdownCode(source, node.attrs.language).then((tokens) => {
        for (const token of tokens) {
          const from = pos + 1 + token.from;
          const to = pos + 1 + token.to;
          if (from >= to || to > pos + node.nodeSize - 1) continue;
          decorations.push(
            Decoration.inline(from, to, {
              class: 'work-markdown-token',
              style: `--shiki-light:${token.light};--shiki-dark:${token.dark}`,
            }),
          );
        }
      }),
    );
  });
  await Promise.all(jobs);
  decorations.sort((left, right) => left.from - right.from);
  return DecorationSet.create(doc, decorations);
}
