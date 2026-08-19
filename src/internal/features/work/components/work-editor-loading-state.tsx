import { StateView } from '../../../design-system/primitives';

export function WorkEditorLoadingState({ title }: { title: string }) {
  return (
    <StateView
      className="work-editor-loading"
      size="compact"
      role="status"
      tone="info"
      icon={<EditorLoadingVisual />}
      title={title}
    />
  );
}

function EditorLoadingVisual() {
  return (
    <span className="work-editor-loading-visual" aria-hidden="true">
      <span className="work-editor-loading-focus" />
      <span className="work-editor-loading-page">
        <span className="work-editor-loading-line work-editor-loading-line--accent" />
        <span className="work-editor-loading-line" />
        <span className="work-editor-loading-line work-editor-loading-line--short" />
      </span>
    </span>
  );
}
