import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

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
