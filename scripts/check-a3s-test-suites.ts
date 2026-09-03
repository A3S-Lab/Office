import { spawnSync } from 'node:child_process';
import {
  accessSync,
  constants,
  existsSync,
  readdirSync,
  readFileSync,
} from 'node:fs';
import { delimiter, dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

type CheckResult = {
  scenarios?: Array<{ id?: string }>;
};

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const suitesRoot = resolve(repositoryRoot, 'tests/e2e');
const a3sTest = resolveExecutable(
  process.env.A3S_TEST_BIN?.trim() || 'a3s-test',
);
const suiteFiles = selectSuites();
const failedSuites: string[] = [];
let scenarioCount = 0;

for (const suiteFile of suiteFiles) {
  const result = spawnSync(a3sTest, ['check', suiteFile, '--json'], {
    cwd: repositoryRoot,
    encoding: 'utf8',
    maxBuffer: 8 * 1024 * 1024,
  });
  if (result.status !== 0) {
    failedSuites.push(relative(repositoryRoot, suiteFile));
    const diagnostics = [result.stdout, result.stderr]
      .filter((value) => value.trim().length > 0)
      .join('\n');
    process.stderr.write(
      `✗ ${relative(repositoryRoot, suiteFile)}\n${diagnostics}\n`,
    );
    continue;
  }

  const parsed = parseCheckResult(result.stdout, suiteFile);
  const scenarios = parsed.scenarios?.length ?? 0;
  scenarioCount += scenarios;
  console.log(
    `✓ ${relative(repositoryRoot, suiteFile)} (${scenarios} scenarios)`,
  );
}

console.log(
  `A3S Test ACL check: ${suiteFiles.length} suites, ${scenarioCount} scenarios, ${failedSuites.length} failures.`,
);

if (failedSuites.length > 0) {
  process.exitCode = 1;
}

function selectSuites(): string[] {
  const requested = process.env.A3S_TEST_SUITE?.trim();
  const suiteFileList = process.env.A3S_TEST_SUITES_FILE?.trim();
  if (suiteFileList) {
    if (!existsSync(suiteFileList)) {
      throw new Error(`A3S_TEST_SUITES_FILE does not exist: ${suiteFileList}`);
    }
    return normalizeSuiteList(readFileSync(suiteFileList, 'utf8'));
  }
  if (!requested || requested === 'all') {
    return readdirSync(suitesRoot)
      .filter((fileName) => fileName.endsWith('.acl'))
      .sort()
      .map((fileName) => resolve(suitesRoot, fileName));
  }
  return normalizeSuiteList(requested);
}

function normalizeSuiteList(rawValue: string): string[] {
  const suites = rawValue
    .split(/[\n,]/u)
    .map((value) => value.trim())
    .filter((value) => value.length > 0 && !value.startsWith('#'))
    .map((value) =>
      value.startsWith('tests/e2e/') ? value : `tests/e2e/${value}`,
    )
    .map((value) => resolve(repositoryRoot, value));
  if (suites.length === 0) {
    throw new Error('A3S_TEST_SUITE did not select any suites.');
  }
  const uniqueSuites = [...new Set(suites)];
  for (const suiteFile of uniqueSuites) {
    const suiteRelative = relative(suitesRoot, suiteFile);
    if (
      suiteRelative.startsWith('..') ||
      suiteRelative.includes('/') ||
      !suiteFile.endsWith('.acl') ||
      !existsSync(suiteFile)
    ) {
      throw new Error(
        `Invalid A3S Test suite: ${relative(repositoryRoot, suiteFile)}. ` +
          'Suites must be existing tests/e2e/*.acl files.',
      );
    }
  }
  return uniqueSuites;
}

function parseCheckResult(output: string, suiteFile: string): CheckResult {
  try {
    const parsed: unknown = JSON.parse(output);
    if (!parsed || typeof parsed !== 'object') {
      throw new Error('result is not an object');
    }
    return parsed as CheckResult;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      `A3S Test returned invalid JSON for ${relative(repositoryRoot, suiteFile)}: ${message}`,
    );
  }
}

function resolveExecutable(configuredName: string): string {
  if (configuredName.includes('/') && isExecutable(configuredName)) {
    return configuredName;
  }

  const localCandidates = [
    resolve(repositoryRoot, '../../crates/test/target/release', configuredName),
    resolve(repositoryRoot, '../../crates/test/target/debug', configuredName),
  ];
  const userHome = process.env.HOME?.trim();
  const candidates = [
    ...localCandidates,
    userHome ? join(userHome, '.local/bin', configuredName) : '',
    userHome ? join(userHome, '.bun/bin', configuredName) : '',
    `/opt/homebrew/bin/${configuredName}`,
    `/usr/local/bin/${configuredName}`,
    ...(process.env.PATH ?? '')
      .split(delimiter)
      .filter((directory) => directory.length > 0)
      .map((directory) => join(directory, configuredName)),
  ].filter((candidate) => candidate.length > 0);
  const fallback = candidates.find(isExecutable);
  if (fallback) return fallback;

  throw new Error(`Required executable not found: ${configuredName}`);
}

function isExecutable(fileName: string): boolean {
  try {
    accessSync(fileName, constants.X_OK);
    return true;
  } catch {
    return false;
  }
}
