import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';

const rawArguments = process.argv.slice(2);
const nativeExecutable =
  process.env.A3S_TEST_AGENT_BROWSER_NATIVE?.trim() ||
  path.join(
    os.homedir(),
    'AppData',
    'Roaming',
    'npm',
    'node_modules',
    'agent-browser',
    'bin',
    'agent-browser-win32-x64.exe',
  );
const cdpPort = process.env.A3S_TEST_CDP_PORT?.trim() || '9343';
const debug = process.env.A3S_TEST_CDP_DEBUG === '1';

if (rawArguments[0] === '--version') {
  console.log('agent-browser 0.26.0');
  process.exit(0);
}

const session = findOptionValue(rawArguments, '--session') || 'a3s-office-cdp';
const forwardedArguments = stripAutoLaunchOptions(rawArguments);
const childEnvironment = { ...process.env };
const runtimeNamespace = `a3s-office-${cdpPort}-${session}`.replace(
  /[^a-z0-9._-]/giu,
  '_',
);
const socketDirectory = path.join(
  os.tmpdir(),
  'a3s-office-agent-browser',
  runtimeNamespace,
);
mkdirSync(socketDirectory, { recursive: true });
for (const key of ['AGENT_BROWSER_ARGS', 'AGENT_BROWSER_IDLE_TIMEOUT_MS']) {
  delete childEnvironment[key];
}
childEnvironment.AGENT_BROWSER_SOCKET_DIR = socketDirectory;
childEnvironment.AGENT_BROWSER_NAMESPACE = runtimeNamespace;
childEnvironment.AGENT_BROWSER_ALLOWED_DOMAINS = '127.0.0.1';

try {
  await ensureCdpEndpoint(cdpPort);
  const viewport = parseViewportCommand(forwardedArguments);
  if (viewport) {
    await setViewportOverCdp(cdpPort, viewport);
    process.stdout.write(
      `${JSON.stringify({
        success: true,
        data: {
          width: viewport.width,
          height: viewport.height,
          deviceScaleFactor: viewport.scale,
        },
        error: null,
      })}\n`,
    );
    process.exit(0);
  }

  const waitCommand = parseWaitCommand(forwardedArguments);
  if (waitCommand) {
    await waitForSelectorOverCdp(cdpPort, waitCommand);
    process.stdout.write(
      `${JSON.stringify({
        success: true,
        data: {
          selector: waitCommand.selector ?? null,
          expression: waitCommand.expression ? true : null,
          timeoutMs: waitCommand.timeoutMs,
        },
        error: null,
      })}\n`,
    );
    process.exit(0);
  }

  const result = await runNative(forwardedArguments);
  if (debug) {
    process.stderr.write(
      `[a3s-test-cdp] cdp=${cdpPort} command=${forwardedArguments.join(' ')} status=${result.status ?? 'null'} error=${result.error?.message ?? ''}\n`,
    );
  }
  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
  if (result.error) {
    process.stderr.write(`${result.error.message}\n`);
    process.exit(1);
  }
  process.exit(result.status ?? 1);
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  if (debug) process.stderr.write(`[a3s-test-cdp] fatal=${message}\n`);
  process.stdout.write(
    `${JSON.stringify({ success: false, data: null, error: message })}\n`,
  );
  process.exit(1);
}

function runNative(arguments_: string[]): Promise<{
  status: number | null;
  stdout: string;
  stderr: string;
  error?: Error;
}> {
  return new Promise((resolve) => {
    const child = spawn(nativeExecutable, arguments_, {
      cwd: process.cwd(),
      env: childEnvironment,
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let stdout = '';
    let stderr = '';
    child.stdout?.on('data', (chunk: Buffer | string) => {
      stdout += chunk.toString();
    });
    child.stderr?.on('data', (chunk: Buffer | string) => {
      stderr += chunk.toString();
    });
    child.once('error', (error) => {
      resolve({ status: null, stdout, stderr, error });
    });
    // agent-browser may leave a persistent session daemon holding the child
    // stdio handles open on Windows. Resolve on process exit instead of the
    // `close` event, which waits for every inherited handle to be released.
    child.once('exit', (status) => {
      child.stdout?.destroy();
      child.stderr?.destroy();
      resolve({ status, stdout, stderr });
    });
    setTimeout(() => {
      if (!child.killed) child.kill();
    }, 60_000).unref();
  });
}

function findOptionValue(
  arguments_: string[],
  option: string,
): string | undefined {
  const index = arguments_.indexOf(option);
  if (index >= 0 && arguments_[index + 1]) return arguments_[index + 1];
  const prefix = `${option}=`;
  return arguments_
    .find((value) => value.startsWith(prefix))
    ?.slice(prefix.length);
}

function stripAutoLaunchOptions(arguments_: string[]): string[] {
  const stripped: string[] = ['--cdp', cdpPort];
  for (let index = 0; index < arguments_.length; index += 1) {
    const value = arguments_[index];
    if (value === '--headed') {
      if (arguments_[index + 1] === 'false') index += 1;
      continue;
    }
    if (value === '--allowed-domains' || value === '--engine') {
      index += 1;
      continue;
    }
    if (value === '--cdp') {
      index += 1;
      continue;
    }
    if (
      value.startsWith('--allowed-domains=') ||
      value.startsWith('--engine=') ||
      value.startsWith('--cdp=')
    ) {
      continue;
    }
    stripped.push(value);
  }
  return stripped;
}

type ViewportCommand = {
  width: number;
  height: number;
  scale: number;
};

function parseViewportCommand(
  arguments_: string[],
): ViewportCommand | undefined {
  const json = arguments_.includes('--json');
  const setIndex = arguments_.indexOf('set');
  if (setIndex < 0 || arguments_[setIndex + 1] !== 'viewport') return undefined;
  const width = Number(arguments_[setIndex + 2]);
  const height = Number(arguments_[setIndex + 3]);
  const scaleRaw = arguments_[setIndex + 4];
  const scale = scaleRaw && !scaleRaw.startsWith('-') ? Number(scaleRaw) : 1;
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0 ||
    !Number.isFinite(scale) ||
    scale <= 0
  ) {
    throw new Error(
      `Invalid viewport arguments: ${arguments_.slice(setIndex).join(' ')}`,
    );
  }
  // Keep --json present for callers that only inspect argv; the adapter
  // itself always emits JSON for this intercepted command.
  void json;
  return { width, height, scale };
}

type WaitCommand = {
  selector?: string;
  expression?: string;
  timeoutMs: number;
};

function parseWaitCommand(arguments_: string[]): WaitCommand | undefined {
  const waitIndex = arguments_.indexOf('wait');
  if (waitIndex < 0) return undefined;

  let timeoutMs = 90_000;
  let expression: string | undefined;
  const selectorParts: string[] = [];
  for (let index = waitIndex + 1; index < arguments_.length; index += 1) {
    const value = arguments_[index];
    if (value === '--json') continue;
    if (value === '--timeout') {
      const next = Number(arguments_[index + 1]);
      if (Number.isFinite(next) && next > 0) timeoutMs = next;
      index += 1;
      continue;
    }
    if (value.startsWith('--timeout=')) {
      const next = Number(value.slice('--timeout='.length));
      if (Number.isFinite(next) && next > 0) timeoutMs = next;
      continue;
    }
    if (value === '--fn') {
      expression = arguments_[index + 1];
      index += 1;
      continue;
    }
    if (value.startsWith('--fn=')) {
      expression = value.slice('--fn='.length);
      continue;
    }
    if (value.startsWith('--')) {
      // Unknown wait options fall back to native agent-browser.
      return undefined;
    }
    selectorParts.push(value);
  }

  if (expression) {
    return { expression, timeoutMs };
  }

  if (selectorParts.length === 0) return undefined;
  if (selectorParts.length === 1 && /^\d+$/u.test(selectorParts[0] ?? '')) {
    return undefined;
  }

  let selector = selectorParts.join(' ').trim();
  if (selector.startsWith('css=')) selector = selector.slice(4);
  if (!selector) return undefined;
  return { selector, timeoutMs };
}

async function waitForSelectorOverCdp(
  port: string,
  waitCommand: WaitCommand,
): Promise<void> {
  const deadline = Date.now() + waitCommand.timeoutMs;
  const expression = waitCommand.expression
    ? `(${waitCommand.expression})`
    : `!!document.querySelector(${JSON.stringify(waitCommand.selector)})`;
  while (Date.now() < deadline) {
    const present = await evaluateBooleanOverCdp(port, expression);
    if (present) return;
    await delay(200);
  }
  throw new Error(
    `Timed out after ${waitCommand.timeoutMs}ms waiting for ${
      waitCommand.expression ?? waitCommand.selector
    }`,
  );
}

async function evaluateBooleanOverCdp(
  port: string,
  expression: string,
): Promise<boolean> {
  const targets = await listCdpTargets(port);
  const page = preferPageTarget(targets);
  if (!page?.webSocketDebuggerUrl) return false;
  let result = false;
  await withCdpSocket(page.webSocketDebuggerUrl, async (send) => {
    const evaluation = (await send('Runtime.evaluate', {
      expression,
      returnByValue: true,
      awaitPromise: true,
    })) as { result?: { value?: unknown } };
    result = Boolean(evaluation.result?.value);
  });
  return result;
}

async function ensureCdpEndpoint(port: string): Promise<void> {
  if (await isCdpReady(port)) return;
  const chromePath = resolveChromeExecutable();
  if (!chromePath) {
    throw new Error(
      `CDP endpoint 127.0.0.1:${port} is unavailable and no Chrome/Edge executable was found.`,
    );
  }
  const profileDirectory = path.join(
    process.cwd(),
    '.a3s-test',
    'office-ops',
    `chrome-cdp-${port}`,
  );
  mkdirSync(profileDirectory, { recursive: true });
  const child = spawn(
    chromePath,
    [
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${profileDirectory}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-background-networking',
      'about:blank',
    ],
    {
      detached: true,
      stdio: 'ignore',
      windowsHide: true,
    },
  );
  child.unref();
  writeFileSync(
    path.join(profileDirectory, 'adapter-launch.pid'),
    String(child.pid ?? ''),
  );
  const deadline = Date.now() + 30_000;
  while (Date.now() < deadline) {
    if (await isCdpReady(port)) return;
    await delay(250);
  }
  throw new Error(
    `Timed out waiting for Chrome CDP on 127.0.0.1:${port} after adapter launch.`,
  );
}

async function isCdpReady(port: string): Promise<boolean> {
  try {
    const response = await fetch(`http://127.0.0.1:${port}/json/version`, {
      signal: AbortSignal.timeout(2_000),
    });
    return response.ok;
  } catch {
    return false;
  }
}

function resolveChromeExecutable(): string | undefined {
  const configured = process.env.A3S_TEST_CHROME_EXECUTABLE?.trim();
  if (configured && existsSync(configured)) return configured;

  // Prefer the Playwright-managed Chromium that already gates Office UI
  // evidence. System Chrome can attach over CDP but has failed to finish the
  // Writer kernel/WASM boot path in local Windows runs.
  const playwrightChromium = resolvePlaywrightChromium();
  if (playwrightChromium) return playwrightChromium;

  const home = os.homedir();
  const candidates = [
    path.join(
      process.env['ProgramFiles'] ?? 'C:\\Program Files',
      'Google',
      'Chrome',
      'Application',
      'chrome.exe',
    ),
    path.join(
      process.env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)',
      'Google',
      'Chrome',
      'Application',
      'chrome.exe',
    ),
    path.join(
      home,
      'AppData',
      'Local',
      'Google',
      'Chrome',
      'Application',
      'chrome.exe',
    ),
    path.join(
      process.env['ProgramFiles'] ?? 'C:\\Program Files',
      'Microsoft',
      'Edge',
      'Application',
      'msedge.exe',
    ),
    path.join(
      home,
      'AppData',
      'Local',
      'Microsoft',
      'Edge',
      'Application',
      'msedge.exe',
    ),
  ];
  return candidates.find((candidate) => existsSync(candidate));
}

function resolvePlaywrightChromium(): string | undefined {
  const browsersRoot = path.join(
    os.homedir(),
    'AppData',
    'Local',
    'ms-playwright',
  );
  if (!existsSync(browsersRoot)) return undefined;
  try {
    const chromiumRoots = readdirSync(browsersRoot)
      .filter((name) => name.startsWith('chromium-'))
      .sort()
      .reverse();
    for (const root of chromiumRoots) {
      const candidate = path.join(
        browsersRoot,
        root,
        'chrome-win64',
        'chrome.exe',
      );
      if (existsSync(candidate)) return candidate;
    }
  } catch {
    return undefined;
  }
  return undefined;
}

async function setViewportOverCdp(
  port: string,
  viewport: ViewportCommand,
): Promise<void> {
  const targets = await listCdpTargets(port);
  let page = preferPageTarget(targets);
  if (!page) {
    await createBlankPage(port);
    page = preferPageTarget(await listCdpTargets(port));
  }
  if (!page?.id) {
    throw new Error(`No CDP page target available on 127.0.0.1:${port}.`);
  }

  // Prefer real window bounds over Emulation.setDeviceMetricsOverride.
  // Device-metrics overrides have broken the Writer kernel/WASM boot path in
  // local Windows CDP runs (async kernel chunks transfer ~300 bytes).
  if (page.webSocketDebuggerUrl) {
    await withCdpSocket(page.webSocketDebuggerUrl, async (send) => {
      await send('Emulation.clearDeviceMetricsOverride').catch(() => undefined);
    });
  }

  const browserSocket = await getBrowserWebSocketUrl(port);
  await withCdpSocket(browserSocket, async (send) => {
    const windowInfo = (await send('Browser.getWindowForTarget', {
      targetId: page.id,
    })) as { windowId?: number };
    if (typeof windowInfo.windowId !== 'number') {
      throw new Error(
        'CDP Browser.getWindowForTarget did not return windowId.',
      );
    }
    await send('Browser.setWindowBounds', {
      windowId: windowInfo.windowId,
      bounds: {
        width: Math.round(viewport.width),
        height: Math.round(viewport.height),
        windowState: 'normal',
      },
    });
  });
}

async function getBrowserWebSocketUrl(port: string): Promise<string> {
  const response = await fetch(`http://127.0.0.1:${port}/json/version`, {
    signal: AbortSignal.timeout(5_000),
  });
  if (!response.ok) {
    throw new Error(`CDP /json/version failed with HTTP ${response.status}.`);
  }
  const version = (await response.json()) as { webSocketDebuggerUrl?: string };
  if (!version.webSocketDebuggerUrl) {
    throw new Error('CDP browser WebSocket URL is missing.');
  }
  return version.webSocketDebuggerUrl;
}

type CdpTarget = {
  id?: string;
  type?: string;
  url?: string;
  webSocketDebuggerUrl?: string;
};

function preferPageTarget(targets: CdpTarget[]): CdpTarget | undefined {
  const pages = targets.filter(
    (target) =>
      target.type === 'page' &&
      typeof target.webSocketDebuggerUrl === 'string' &&
      target.webSocketDebuggerUrl.length > 0 &&
      typeof target.id === 'string',
  );
  const playground = pages.find((target) =>
    /127\.0\.0\.1|localhost/u.test(target.url ?? ''),
  );
  return playground ?? pages[0];
}

async function listCdpTargets(port: string): Promise<CdpTarget[]> {
  const response = await fetch(`http://127.0.0.1:${port}/json/list`, {
    signal: AbortSignal.timeout(5_000),
  });
  if (!response.ok) {
    throw new Error(`CDP /json/list failed with HTTP ${response.status}.`);
  }
  return (await response.json()) as CdpTarget[];
}

async function createBlankPage(port: string): Promise<void> {
  const versionResponse = await fetch(`http://127.0.0.1:${port}/json/version`, {
    signal: AbortSignal.timeout(5_000),
  });
  if (!versionResponse.ok) {
    throw new Error(
      `CDP /json/version failed with HTTP ${versionResponse.status}.`,
    );
  }
  const version = (await versionResponse.json()) as {
    webSocketDebuggerUrl?: string;
  };
  if (!version.webSocketDebuggerUrl) {
    throw new Error('CDP browser WebSocket URL is missing.');
  }
  await withCdpSocket(version.webSocketDebuggerUrl, async (send) => {
    await send('Target.createTarget', { url: 'about:blank' });
  });
}

async function withCdpSocket(
  url: string,
  run: (
    send: (
      method: string,
      params?: Record<string, unknown>,
    ) => Promise<unknown>,
  ) => Promise<void>,
): Promise<void> {
  const socket = new WebSocket(url);
  let nextId = 1;
  const pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void }
  >();

  const ready = new Promise<void>((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error(`Timed out opening CDP socket ${url}`)),
      10_000,
    );
    socket.addEventListener('open', () => {
      clearTimeout(timer);
      resolve();
    });
    socket.addEventListener('error', () => {
      clearTimeout(timer);
      reject(new Error(`CDP socket error for ${url}`));
    });
  });

  socket.addEventListener('message', (event) => {
    try {
      const payload = JSON.parse(String(event.data)) as {
        id?: number;
        result?: unknown;
        error?: { message?: string };
      };
      if (typeof payload.id !== 'number') return;
      const waiter = pending.get(payload.id);
      if (!waiter) return;
      pending.delete(payload.id);
      if (payload.error) {
        waiter.reject(new Error(payload.error.message || 'CDP command failed'));
        return;
      }
      waiter.resolve(payload.result);
    } catch (error) {
      // Ignore malformed CDP chatter; command waiters keep their own timeouts.
      void error;
    }
  });

  await ready;
  try {
    await run(async (method, params = {}) => {
      const id = nextId;
      nextId += 1;
      const result = new Promise<unknown>((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          reject(new Error(`CDP ${method} timed out`));
        }, 10_000);
        pending.set(id, {
          resolve: (value) => {
            clearTimeout(timer);
            resolve(value);
          },
          reject: (error) => {
            clearTimeout(timer);
            reject(error);
          },
        });
      });
      socket.send(JSON.stringify({ id, method, params }));
      return result;
    });
  } finally {
    socket.close();
  }
}
