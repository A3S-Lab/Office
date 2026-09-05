import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { InvalidArgumentError, Option } from 'commander';

export type Surface = {
  id: string;
  label: string;
  kind: string;
  formats: string[];
  visual: string[];
  acl: string[];
  fixtures: string[];
  wpsReference?: boolean;
  agent?: {
    goal: string;
    success: string;
  };
};

export type Matrix = {
  schemaVersion: number;
  description: string;
  surfaces: Surface[];
  shared: {
    visualConfig: string;
    testkitAcl: string;
    fixtureCommand: string;
    previewUrl: string;
    developmentUrl: string;
    evidenceRoot: string;
  };
};

export type CapturedProcess = {
  status: number | null;
  stdout: string;
  stderr: string;
};

export type GuiProfile = {
  platform?: string;
  status?: string;
  endpoint?: string;
  reason?: string;
};

export const repositoryRoot = path.resolve(import.meta.dirname, '..');
export const matrix = readJson<Matrix>(
  path.join(repositoryRoot, 'scripts', 'office-editor-matrix.json'),
);

export async function startPreview(
  url: string,
): Promise<ReturnType<typeof spawn> | undefined> {
  if (await isHttpReady(url)) return undefined;
  const preview = spawn('bun', ['run', 'playground:preview'], {
    cwd: repositoryRoot,
    env: process.env,
    shell: process.platform === 'win32',
    stdio: 'ignore',
    windowsHide: true,
  });
  try {
    await waitForHttp(url, 30_000);
  } catch (error) {
    await stopPreview(preview);
    throw error;
  }
  return preview;
}

export async function stopPreview(
  preview: ReturnType<typeof spawn> | undefined,
): Promise<void> {
  if (!preview?.pid) return;
  if (process.platform === 'win32') {
    spawnSync('taskkill.exe', ['/PID', String(preview.pid), '/T', '/F'], {
      stdio: 'ignore',
      windowsHide: true,
    });
    return;
  }
  preview.kill('SIGTERM');
}

async function waitForHttp(url: string, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  let lastError = 'no response';
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(1_000) });
      if (response.ok) return;
      lastError = `HTTP ${response.status}`;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    await new Promise<void>((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`Preview did not become ready at ${url}: ${lastError}`);
}

async function isHttpReady(url: string): Promise<boolean> {
  try {
    const response = await fetch(url, { signal: AbortSignal.timeout(500) });
    return response.ok;
  } catch {
    return false;
  }
}

export function validateLoopbackUrl(value: string): string {
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    throw new InvalidArgumentError(`Invalid URL: ${value}`);
  }
  if (
    !['http:', 'https:'].includes(parsed.protocol) ||
    !['127.0.0.1', 'localhost', '[::1]'].includes(parsed.hostname)
  ) {
    throw new InvalidArgumentError(
      `Only loopback HTTP(S) URLs are allowed for local UI evidence: ${value}`,
    );
  }
  return value.replace(/\/$/u, '');
}

export function materializeManifest(suite: string, baseUrl: string): string {
  const defaultOrigin = new URL(matrix.shared.previewUrl).origin;
  const targetOrigin = new URL(baseUrl).origin;
  if (defaultOrigin === targetOrigin) return suite;
  const generatedRoot = path.join(
    repositoryRoot,
    matrix.shared.evidenceRoot,
    'generated',
  );
  mkdirSync(generatedRoot, { recursive: true });
  const source = readFileSync(path.join(repositoryRoot, suite), 'utf8');
  const generatedName = `${path.basename(suite, '.acl')}-${targetOrigin
    .replace(/[^a-z0-9]+/giu, '-')
    .replace(/^-|-$/gu, '')}.acl`;
  const generatedPath = path.join(generatedRoot, generatedName);
  writeFileSync(generatedPath, source.replaceAll(defaultOrigin, targetOrigin));
  return generatedPath;
}

export function capture(
  commandName: string,
  commandArgs: string[],
  env: NodeJS.ProcessEnv = process.env,
): CapturedProcess {
  const result = spawnSync(commandName, commandArgs, {
    cwd: repositoryRoot,
    encoding: 'utf8',
    env,
    maxBuffer: 32 * 1024 * 1024,
    windowsHide: true,
  });
  return {
    status: result.status,
    stdout: result.stdout ?? '',
    stderr: result.stderr ?? result.error?.message ?? '',
  };
}

export function parseJson<T>(value: string): T | undefined {
  try {
    return JSON.parse(value) as T;
  } catch {
    return undefined;
  }
}

export function findGuiProfile(
  certification: Record<string, unknown> | undefined,
  platform: string,
): GuiProfile | undefined {
  const profiles = certification?.profiles;
  if (!Array.isArray(profiles)) return undefined;
  return profiles.find(
    (profile): profile is GuiProfile =>
      typeof profile === 'object' &&
      profile !== null &&
      (profile as GuiProfile).platform === platform,
  );
}

export function resolveSurface(surfaceId: string): Surface {
  const surface = matrix.surfaces.find(
    (candidate) => candidate.id === surfaceId,
  );
  if (!surface) {
    throw new InvalidArgumentError(
      `Unknown editor surface '${surfaceId}'. Run capabilities to list the matrix.`,
    );
  }
  return surface;
}

export function normalizeRelative(value: string): string {
  const normalized = value.replaceAll('\\', '/').replace(/^\.\//u, '');
  const absolute = path.resolve(repositoryRoot, normalized);
  if (
    !absolute.startsWith(`${repositoryRoot}${path.sep}`) ||
    !existsSync(absolute)
  ) {
    throw new InvalidArgumentError(
      `Path does not exist inside the Office repository: ${value}`,
    );
  }
  return normalized;
}

export function resolveA3sTest(): string | undefined {
  const configured = process.env.A3S_TEST_BIN?.trim();
  const executableName =
    process.platform === 'win32' ? 'a3s-test.exe' : 'a3s-test';
  const candidates = [
    configured,
    path.join(
      repositoryRoot,
      '..',
      '..',
      'crates',
      'test',
      'target',
      'release',
      executableName,
    ),
    path.join(
      repositoryRoot,
      '..',
      '..',
      'crates',
      'test',
      'target',
      'debug',
      executableName,
    ),
  ].filter((candidate): candidate is string => Boolean(candidate));
  const pathCandidate = process.env.PATH?.split(path.delimiter)
    .map((directory) => path.join(directory, executableName))
    .find((candidate) => existsSync(candidate));
  return [...candidates, pathCandidate].find(
    (candidate): candidate is string =>
      Boolean(candidate) && existsSync(candidate),
  );
}

export function commandExists(
  commandName: string,
  commandArgs: string[],
): boolean {
  return spawnSync(commandName, commandArgs, { stdio: 'ignore' }).status === 0;
}

export function collectProject(value: string, previous: string[]): string[] {
  if (!['desktop-1280', 'compact-768'].includes(value)) {
    throw new InvalidArgumentError(
      '--project must be desktop-1280 or compact-768',
    );
  }
  return [...previous, value];
}

export function browserDriverOption(fallback?: 'a3s' | 'standalone'): Option {
  const defaultDriver =
    fallback ??
    (process.env.A3S_TEST_BROWSER_DRIVER === 'a3s' ? 'a3s' : 'standalone');
  return new Option('--browser-driver <driver>', 'a3s or standalone')
    .choices(['a3s', 'standalone'])
    .default(defaultDriver);
}

export function collectValue(value: string, previous: string[]): string[] {
  const trimmed = value.trim();
  if (!trimmed) throw new InvalidArgumentError('The value must not be empty.');
  return [...previous, trimmed];
}

export function run(
  commandName: string,
  commandArgs: string[],
  options: { env?: NodeJS.ProcessEnv } = {},
): void {
  const result = spawnSync(commandName, commandArgs, {
    cwd: repositoryRoot,
    env: options.env,
    stdio: 'inherit',
    shell:
      process.platform === 'win32' &&
      ['bun', 'bunx', 'powershell.exe'].includes(commandName),
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(
      `${commandName} ${commandArgs.join(' ')} exited with ${result.status ?? 'unknown status'}.`,
    );
  }
}

function readJson<T>(filePath: string): T {
  return JSON.parse(readFileSync(filePath, 'utf8')) as T;
}
