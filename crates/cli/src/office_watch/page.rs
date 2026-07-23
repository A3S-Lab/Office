pub(super) const WATCH_PAGE: &str = r#"<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>A3S Native Office Watch</title>
  <style>
    :root{color-scheme:light dark;font-family:ui-sans-serif,system-ui,sans-serif}
    *{box-sizing:border-box}body{margin:0;background:#111827;color:#f8fafc}
    header{height:44px;display:flex;align-items:center;gap:12px;padding:0 16px;background:#172033;border-bottom:1px solid #344054}
    h1{font-size:14px;margin:0;font-weight:650;flex-shrink:0}.status{margin-left:auto;max-width:60%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;color:#98a2b3}
    .dot{width:8px;height:8px;border-radius:50%;background:#f59e0b}.dot.ready{background:#22c55e}.dot.error{background:#ef4444}
    main{height:calc(100vh - 44px);padding:12px;background:#e9edf3}
    iframe{width:100%;height:100%;border:0;border-radius:8px;background:white;box-shadow:0 4px 18px #10182833}
  </style>
</head>
<body>
  <header><span id="dot" class="dot" aria-hidden="true"></span><h1>A3S Native Office semantic preview</h1><span id="status" class="status" role="status" aria-live="polite">Connecting…</span></header>
  <main><iframe id="preview" title="Native Office semantic preview" sandbox referrerpolicy="no-referrer"></iframe></main>
  <script src="/watch.js" defer></script>
</body>
</html>"#;

pub(super) const WATCH_SCRIPT: &str = r#"(() => {
  'use strict';
  const preview = document.getElementById('preview');
  const status = document.getElementById('status');
  const dot = document.getElementById('dot');
  let version = 0;

  history.replaceState(null, '', '/');

  function show(state, message) {
    dot.className = `dot ${state}`;
    status.textContent = message;
    status.title = message;
  }

  function apply(event) {
    if (!event.healthy) {
      show('error', event.error ? `${event.error.code}: ${event.error.message}` : 'Reload failed');
      return;
    }
    if (event.version >= version) {
      version = event.version;
      preview.src = `/preview?version=${encodeURIComponent(String(version))}`;
    }
    show('ready', `Revision ${version}`);
  }

  const events = new EventSource('/events');
  events.addEventListener('snapshot', message => {
    try { apply(JSON.parse(message.data)); }
    catch (_) { show('error', 'Invalid watch event'); }
  });
  events.addEventListener('render-error', message => {
    try { apply(JSON.parse(message.data)); }
    catch (_) { show('error', 'Invalid watch error'); }
  });
  events.onopen = () => { if (version === 0) show('', 'Connected'); };
  events.onerror = () => show('', 'Reconnecting…');
})();"#;
