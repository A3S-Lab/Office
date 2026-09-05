import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { Command, CommanderError, Option } from 'commander';
import {
  browserDriverOption,
  capture,
  collectProject,
  collectValue,
  commandExists,
  findGuiProfile,
  materializeManifest,
  matrix,
  normalizeRelative,
  parseJson,
  repositoryRoot,
  resolveA3sTest,
  resolveSurface,
  run,
  startPreview,
  stopPreview,
  validateLoopbackUrl,
} from './office-ui-ops-core';

type JsonOption = { json?: boolean };
type VisualOption = JsonOption & {
  project?: string[];
  baseUrl?: string;
};
type WpsProbeOption = JsonOption & {
  output?: string;
  connector?: boolean;
  connectorType?: 'straight' | 'elbow' | 'curved';
};
type WpsFieldsProbeOption = JsonOption & {
  output?: string;
};
type A3sRunOption = JsonOption & {
  browserDriver?: 'a3s' | 'standalone';
  browserExecutable?: string;
  baseUrl?: string;
  cdpPort?: string;
  headed?: boolean;
  build?: boolean;
  commandTimeoutMs?: string;
  idleTimeoutMs?: string;
};
type AgentStartOption = JsonOption & {
  session?: string;
  url?: string;
  goal?: string;
  success?: string[];
  browserDriver?: 'a3s' | 'standalone';
  browserExecutable?: string;
  cdpPort?: string;
  headed?: boolean;
};
type AgentSessionOption = JsonOption & {
  session: string;
};
type AgentActionOption = AgentSessionOption & {
  observation?: string;
  actionJson: string;
};
type AgentFinishOption = AgentSessionOption & {
  status: 'passed' | 'failed';
  summary: string;
};
type CuaCertifyOption = JsonOption & {
  guiPolicyFile?: string;
  cuaProxyExecutable?: string;
  cuaEmbeddedSocket?: string;
  guiMacosBundleId?: string;
  guiTargetMode?: 'launch' | 'attach';
  guiProfile?: 'semantic' | 'window-vision';
  guiAttachPid?: string;
  guiWindowTitle?: string;
  guiWindowAutomationId?: string;
  guiArg?: string[];
  commandTimeoutMs?: string;
  cleanupTimeoutMs?: string;
  artifactsRoot?: string;
};
type GateOption = JsonOption &
  A3sRunOption & {
    run?: boolean;
    visual?: boolean;
  };

const wpsConnectorTypeCodes = {
  straight: '1',
  elbow: '2',
  curved: '3',
} as const;

const program = new Command()
  .name('office-ui-ops')
  .description(
    'Operate A3S Office editor UI contracts with local fixtures and evidence.',
  )
  .showSuggestionAfterError()
  .showHelpAfterError();

program
  .command('capabilities')
  .description('List the five editor surfaces and their evidence contracts.')
  .argument('[surface]', 'writer, spreadsheet, presentation, markdown, or pdf')
  .option('--json', 'emit a machine-readable result')
  .action((surface: string | undefined, options: JsonOption) => {
    printCapabilities(surface, options.json === true);
  });

program
  .command('doctor')
  .description('Check local tools, the UI bundle, A3S Test, and WPS probe.')
  .option('--json', 'emit a machine-readable result')
  .action((options: JsonOption) => {
    printDoctor(options.json === true);
  });

program
  .command('fixtures')
  .description('Generate deterministic ignored fixtures for all editors.')
  .action(() => runFixtures());

program
  .command('check')
  .description('Parse focused A3S Test ACLs without starting a browser.')
  .argument('[selection]', 'surface id, ACL path, or all', 'all')
  .option('--json', 'emit a machine-readable result')
  .action((selection: string, options: JsonOption) => {
    runChecks(selection, options.json === true);
  });

program
  .command('visual')
  .description(
    'Build the UI bundle and run a supplemental Playwright pixel contract.',
  )
  .argument('<selection>', 'surface id or visual spec path')
  .option(
    '--project <project>',
    'desktop-1280 or compact-768 (repeatable)',
    collectProject,
    [],
  )
  .option('--base-url <url>', 'reuse an already-running preview or dev server')
  .option('--json', 'reserved for a future structured Playwright receipt')
  .action((selection: string, options: VisualOption) => {
    runVisual(selection, options);
  });

program
  .command('wps-probe')
  .description('Capture a bounded WPS Writer COM reference document (Windows).')
  .option('--output <docx>', 'exact output DOCX path')
  .option('--connector', 'include a connector of the selected type')
  .addOption(
    new Option('--connector-type <kind>', 'WPS connector shape kind').choices([
      'straight',
      'elbow',
      'curved',
    ]),
  )
  .option('--json', 'reserved for the probe JSON receipt')
  .action((options: WpsProbeOption) => {
    runWpsProbe(options);
  });

program
  .command('wps-fields-probe')
  .description(
    'Capture WPS Writer COM field-switch reference output (Windows).',
  )
  .option('--output <docx>', 'exact output DOCX path')
  .option('--json', 'reserved for the probe JSON receipt')
  .action((options: WpsFieldsProbeOption) => {
    runWpsFieldsProbe(options);
  });

const a3s = program
  .command('a3s')
  .description('Run the primary A3S Test ACL and agent evidence workflows.');

a3s
  .command('run')
  .description(
    'Run one surface ACL through A3S Test; Playwright is not involved.',
  )
  .argument('<selection>', 'surface id, ACL path, or all')
  .addOption(browserDriverOption())
  .option('--browser-executable <path>', 'A3S Test browser executable')
  .option(
    '--base-url <url>',
    'reuse a running loopback preview; defaults to the local 4175 preview',
  )
  .option('--cdp-port <port>', 'CDP port consumed by the standalone adapter')
  .option('--headed', 'show the browser window')
  .option('--no-build', 'reuse the existing playground bundle')
  .option('--command-timeout-ms <ms>', 'per-command browser deadline')
  .option('--idle-timeout-ms <ms>', 'browser daemon idle deadline')
  .option('--json', 'emit A3S Test JSON')
  .action(async (selection: string, options: A3sRunOption) => {
    await runA3s(selection, options);
  });

const agent = a3s
  .command('agent')
  .description('Drive the persistent A3S Test observe-decide-act protocol.');

agent
  .command('start')
  .description('Start a workspace-local Web agent session.')
  .argument('<surface>', 'writer, spreadsheet, presentation, markdown, or pdf')
  .option('--session <id>', 'stable session id')
  .option('--url <url>', 'initial loopback URL')
  .option('--goal <text>', 'agent goal; defaults to the matrix goal')
  .option(
    '--success <text>',
    'observable success criterion; repeatable',
    collectValue,
    [],
  )
  .addOption(browserDriverOption('standalone'))
  .option('--browser-executable <path>', 'A3S Test browser executable')
  .option('--cdp-port <port>', 'CDP port consumed by the standalone adapter')
  .option('--headed', 'show the browser window')
  .option('--json', 'emit machine-readable JSON')
  .action(async (surfaceId: string, options: AgentStartOption) => {
    await runAgentStart(surfaceId, options);
  });

agent
  .command('observe')
  .description('Capture the next semantic observation.')
  .requiredOption('--session <id>', 'active A3S Test agent session')
  .option('--interactive', 'include only interactive targets')
  .option('--json', 'emit machine-readable JSON')
  .action((options: AgentSessionOption & { interactive?: boolean }) => {
    runAgentPassthrough([
      'observe',
      '--session',
      options.session,
      ...(options.interactive ? ['--interactive'] : []),
      ...(options.json ? ['--json'] : []),
    ]);
  });

agent
  .command('act')
  .description('Execute one schema-validated typed action.')
  .requiredOption('--session <id>', 'active A3S Test agent session')
  .requiredOption('--action-json <json>', 'typed A3S Test action JSON')
  .option('--observation <id>', 'observation that supplied a ref target')
  .option('--json', 'emit machine-readable JSON')
  .action((options: AgentActionOption) => {
    runAgentPassthrough([
      'act',
      '--session',
      options.session,
      ...(options.observation ? ['--observation', options.observation] : []),
      '--action-json',
      options.actionJson,
      ...(options.json ? ['--json'] : []),
    ]);
  });

agent
  .command('finish')
  .description('Finish a session and write its report.')
  .requiredOption('--session <id>', 'active A3S Test agent session')
  .addOption(
    new Option('--status <status>', 'passed or failed')
      .choices(['passed', 'failed'])
      .makeOptionMandatory(),
  )
  .requiredOption('--summary <text>', 'evidence-backed session summary')
  .option('--json', 'emit machine-readable JSON')
  .action((options: AgentFinishOption) => {
    runAgentPassthrough([
      'finish',
      '--session',
      options.session,
      '--status',
      options.status,
      '--summary',
      options.summary,
      ...(options.json ? ['--json'] : []),
    ]);
  });

agent
  .command('abort')
  .description('Abort a session and clean only its owned browser surface.')
  .requiredOption('--session <id>', 'active A3S Test agent session')
  .option('--json', 'emit machine-readable JSON')
  .action((options: AgentSessionOption) => {
    runAgentPassthrough([
      'abort',
      '--session',
      options.session,
      ...(options.json ? ['--json'] : []),
    ]);
  });

const cua = a3s
  .command('cua')
  .description('Inspect the locked CUA Driver/MCP certification matrix.');

cua
  .command('certification')
  .alias('capabilities')
  .description('Report platform support without attempting an unsafe GUI run.')
  .option('--json', 'emit machine-readable JSON')
  .action((options: JsonOption) => {
    runCuaCertification(options.json === true);
  });

cua
  .command('certify')
  .description(
    'Run the real CUA certification profile when the locked platform supports it.',
  )
  .option(
    '--gui-policy-file <path>',
    'absolute or workspace-relative CUA policy',
  )
  .option('--cua-proxy-executable <path>', 'CUA MCP proxy executable')
  .option('--cua-embedded-socket <path>', 'embedded CUA socket')
  .option('--gui-macos-bundle-id <id>', 'macOS application bundle ID')
  .addOption(
    new Option('--gui-target-mode <mode>', 'launch or attach')
      .choices(['launch', 'attach'])
      .default('launch'),
  )
  .addOption(
    new Option('--gui-profile <profile>', 'semantic or window-vision')
      .choices(['semantic', 'window-vision'])
      .default('semantic'),
  )
  .option('--gui-attach-pid <pid>', 'existing application PID')
  .option('--gui-window-title <title>', 'exact top-level window title')
  .option('--gui-window-automation-id <id>', 'exact window automation ID')
  .option(
    '--gui-arg <value>',
    'application argument; repeatable',
    collectValue,
    [],
  )
  .option('--command-timeout-ms <ms>', 'per-command CUA deadline')
  .option('--cleanup-timeout-ms <ms>', 'bounded cleanup deadline')
  .option('--artifacts-root <path>', 'certification artifact root')
  .option('--json', 'emit machine-readable JSON')
  .action((options: CuaCertifyOption) => {
    runCuaCertify(options);
  });

program
  .command('gate')
  .description(
    'Run fixtures and A3S ACL checks, with optional primary A3S execution and supplemental pixels.',
  )
  .argument('<surface>', 'writer, spreadsheet, presentation, markdown, or pdf')
  .option(
    '--run',
    'execute the checked ACL through A3S Test (requires a configured browser)',
  )
  .option('--no-visual', 'skip the supplemental Playwright contract')
  .addOption(browserDriverOption('standalone'))
  .option('--browser-executable <path>', 'A3S Test browser executable')
  .option('--base-url <url>', 'reuse a running loopback preview')
  .option('--cdp-port <port>', 'CDP port consumed by the standalone adapter')
  .option('--headed', 'show the browser window')
  .option('--no-build', 'reuse the existing playground bundle')
  .option('--command-timeout-ms <ms>', 'per-command browser deadline')
  .option('--idle-timeout-ms <ms>', 'browser daemon idle deadline')
  .option('--json', 'emit machine-readable stage receipts')
  .action(async (surface: string, options: GateOption) => {
    await runGate(surface, options);
  });

program.parseAsync(process.argv).catch((error: unknown) => {
  if (error instanceof CommanderError) {
    if (error.code === 'commander.helpDisplayed') return;
    process.exitCode = error.exitCode;
    return;
  }
  const message = error instanceof Error ? error.message : String(error);
  console.error(`office-ui-ops: ${message}`);
  process.exitCode = 1;
});

function printCapabilities(
  surfaceId: string | undefined,
  asJson: boolean,
): void {
  const surfaces = surfaceId ? [resolveSurface(surfaceId)] : matrix.surfaces;
  const payload = {
    schemaVersion: matrix.schemaVersion,
    ok: true,
    surfaces,
    shared: matrix.shared,
  };
  if (asJson) {
    console.log(JSON.stringify(payload, null, 2));
    return;
  }
  for (const surface of surfaces) {
    console.log(
      `${surface.id.padEnd(13)} ${surface.label.padEnd(14)} ` +
        `visual=${surface.visual.length} acl=${surface.acl.length} ` +
        `formats=${surface.formats.join(',')}`,
    );
  }
}

function printDoctor(asJson: boolean): void {
  const a3sTest = resolveA3sTest();
  const cuaProbe = a3sTest
    ? capture(a3sTest, ['gui-certification', '--json'])
    : undefined;
  const cuaCertification =
    cuaProbe?.status === 0
      ? parseJson<Record<string, unknown>>(cuaProbe.stdout)
      : undefined;
  const windowsCua = findGuiProfile(cuaCertification, 'windows');
  const checks = {
    bun: commandExists('bun', ['--version']),
    playwright: existsSync(
      path.join(repositoryRoot, 'node_modules', '@playwright', 'test'),
    ),
    playgroundBundle: existsSync(
      path.join(repositoryRoot, 'playground-dist', 'playground', 'index.html'),
    ),
    a3sTest: Boolean(a3sTest),
    cuaCertification: cuaProbe?.status === 0,
    wpsProbe:
      process.platform === 'win32' &&
      existsSync(path.join(repositoryRoot, 'scripts', 'probe-wps-shapes.ps1')),
  };
  const payload = {
    schemaVersion: 1,
    ok: Object.values(checks).every(Boolean),
    checks,
    cua: {
      windows: windowsCua?.status ?? 'unknown',
      note:
        windowsCua?.status === 'unsupported'
          ? 'The locked CUA Driver 0.10.0 profile is not contract-tested on Windows; use A3S Test Web/CDP for browser evidence.'
          : undefined,
    },
  };
  if (asJson) {
    console.log(JSON.stringify(payload, null, 2));
  } else {
    for (const [name, available] of Object.entries(checks)) {
      console.log(`${available ? '✓' : '✗'} ${name}`);
    }
    console.log(`! cua.windows=${payload.cua.windows}`);
    if (payload.cua.note) console.log(`  ${payload.cua.note}`);
  }
  if (!payload.ok) process.exitCode = 1;
}

function runFixtures(): void {
  run('bun', ['run', 'test:e2e:fixtures']);
}

function runChecks(selection: string, asJson: boolean): void {
  const suites = resolveSuites(selection);
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  const results: Array<{ suite: string; ok: boolean }> = [];
  for (const suite of [...new Set(suites)]) {
    const result = spawnSync(a3sTest, ['check', suite, '--json'], {
      cwd: repositoryRoot,
      encoding: 'utf8',
      maxBuffer: 8 * 1024 * 1024,
    });
    const ok = result.status === 0;
    results.push({ suite, ok });
    if (!ok) {
      process.stderr.write(
        result.stderr || result.stdout || 'A3S Test check failed.\n',
      );
    } else if (!asJson) {
      console.log(`✓ ${suite}`);
    }
  }
  const payload = {
    schemaVersion: 1,
    ok: results.every((result) => result.ok),
    results,
  };
  if (asJson) console.log(JSON.stringify(payload, null, 2));
  if (!payload.ok) process.exitCode = 1;
}

async function runA3s(selection: string, options: A3sRunOption): Promise<void> {
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  const suites = resolveSuites(selection);
  const baseUrl = validateLoopbackUrl(
    options.baseUrl ??
      process.env.A3S_OFFICE_A3S_BASE_URL ??
      matrix.shared.previewUrl,
  );
  const build = options.build !== false;
  if (build) run('bun', ['run', 'playground:build:ui']);

  const preview = options.baseUrl ? undefined : await startPreview(baseUrl);
  const env: NodeJS.ProcessEnv = { ...process.env };
  if (options.cdpPort) env.A3S_TEST_CDP_PORT = options.cdpPort;
  const browserDriver = options.browserDriver ?? 'standalone';
  const browserExecutable =
    options.browserExecutable ?? process.env.A3S_TEST_AGENT_BROWSER;
  const receipts: unknown[] = [];
  try {
    for (const suite of suites) {
      const manifest = materializeManifest(suite, baseUrl);
      const args = [
        'run',
        manifest,
        '--browser-driver',
        browserDriver,
        '--max-parallel-scenarios',
        '1',
        '--infrastructure-retries',
        '1',
      ];
      if (browserExecutable)
        args.push('--browser-executable', browserExecutable);
      if (options.headed) args.push('--headed');
      if (options.commandTimeoutMs) {
        args.push('--command-timeout-ms', options.commandTimeoutMs);
      }
      if (options.idleTimeoutMs) {
        args.push('--idle-timeout-ms', options.idleTimeoutMs);
      }
      if (options.json) args.push('--json');
      if (process.env.A3S_OFFICE_OPS_DEBUG === '1') {
        console.error(
          JSON.stringify({
            a3sTest,
            browserDriver,
            browserExecutable,
            cdpPort: env.A3S_TEST_CDP_PORT,
            args,
          }),
        );
      }
      const result = capture(a3sTest, args, env);
      if (options.json) {
        const receipt = parseJson<unknown>(result.stdout);
        receipts.push(receipt ?? { suite, stdout: result.stdout });
      } else {
        if (result.stdout) process.stdout.write(result.stdout);
        if (result.stderr) process.stderr.write(result.stderr);
      }
      if (result.status !== 0) {
        if (options.json) {
          if (result.stdout) process.stdout.write(result.stdout);
          if (result.stderr) process.stderr.write(result.stderr);
        }
        throw new Error(
          `A3S Test run failed for ${suite} with status ${result.status ?? 'unknown'}.`,
        );
      }
    }
  } finally {
    await stopPreview(preview);
  }
  if (options.json) {
    console.log(
      JSON.stringify(
        { schemaVersion: 1, ok: true, results: receipts },
        null,
        2,
      ),
    );
  }
}

function runAgentPassthrough(args: string[]): void {
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  run(a3sTest, ['agent', ...args]);
}

async function runAgentStart(
  surfaceId: string,
  options: AgentStartOption,
): Promise<void> {
  const surface = resolveSurface(surfaceId);
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  const url = validateLoopbackUrl(
    options.url ??
      process.env.A3S_OFFICE_A3S_AGENT_URL ??
      matrix.shared.previewUrl,
  );
  const defaults = surface.agent ?? {
    goal: `Inspect the ${surface.label} editing surface`,
    success: `The ${surface.label} editor is loaded`,
  };
  const session =
    options.session ??
    `office-${surface.id}-${new Date()
      .toISOString()
      .replace(/[^0-9a-z]/giu, '')
      .slice(0, 16)}`;
  const success = options.success?.length
    ? options.success
    : [defaults.success];
  const browserDriver = options.browserDriver ?? 'standalone';
  const browserExecutable =
    options.browserExecutable ??
    process.env.A3S_TEST_AGENT_BROWSER ??
    (options.cdpPort
      ? path.join(repositoryRoot, 'scripts', 'a3s-test-cdp-browser.cmd')
      : undefined);
  const env: NodeJS.ProcessEnv = { ...process.env };
  if (options.cdpPort) env.A3S_TEST_CDP_PORT = options.cdpPort;
  const args = [
    'agent',
    'start',
    url,
    '--session',
    session,
    '--goal',
    options.goal ?? defaults.goal,
    ...success.flatMap((criterion) => ['--success', criterion]),
    '--browser-driver',
    browserDriver,
  ];
  if (browserExecutable) args.push('--browser-executable', browserExecutable);
  if (options.headed) args.push('--headed');
  if (options.json) args.push('--json');
  run(a3sTest, args, { env });
}

function runCuaCertification(asJson: boolean): void {
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  const result = capture(a3sTest, ['gui-certification', '--json']);
  if (result.status !== 0) {
    process.stderr.write(result.stderr || result.stdout);
    throw new Error('A3S Test CUA certification failed.');
  }
  const certification = parseJson<Record<string, unknown>>(result.stdout);
  if (asJson) {
    console.log(
      JSON.stringify(certification ?? { raw: result.stdout }, null, 2),
    );
    return;
  }
  const windows = findGuiProfile(certification, 'windows');
  const macos = findGuiProfile(certification, 'macos');
  console.log(
    `CUA Driver: ${String(certification?.cua_driver_version ?? 'unknown')}`,
  );
  console.log(`Windows: ${windows?.status ?? 'unknown'}`);
  console.log(`macOS: ${macos?.status ?? 'unknown'}`);
  if (windows?.status === 'unsupported') {
    console.log(
      'Windows GUI execution is fail-closed until the locked CUA adapter has a reviewed Windows profile.',
    );
  }
}

function runCuaCertify(options: CuaCertifyOption): void {
  const a3sTest = resolveA3sTest();
  if (!a3sTest) {
    throw new Error(
      'a3s-test was not found. Set A3S_TEST_BIN or build crates/test first.',
    );
  }
  const args = ['gui-certify'];
  const optionPairs: Array<[string, string | undefined]> = [
    ['--gui-policy-file', options.guiPolicyFile],
    ['--cua-proxy-executable', options.cuaProxyExecutable],
    ['--cua-embedded-socket', options.cuaEmbeddedSocket],
    ['--gui-macos-bundle-id', options.guiMacosBundleId],
    ['--gui-target-mode', options.guiTargetMode],
    ['--gui-profile', options.guiProfile],
    ['--gui-attach-pid', options.guiAttachPid],
    ['--gui-window-title', options.guiWindowTitle],
    ['--gui-window-automation-id', options.guiWindowAutomationId],
    ['--command-timeout-ms', options.commandTimeoutMs],
    ['--cleanup-timeout-ms', options.cleanupTimeoutMs],
    ['--artifacts-root', options.artifactsRoot],
  ];
  for (const [flag, value] of optionPairs) {
    if (value) args.push(flag, value);
  }
  for (const value of options.guiArg ?? []) args.push('--gui-arg', value);
  if (options.json) args.push('--json');
  run(a3sTest, args);
}

function runVisual(selection: string, options: VisualOption): void {
  const surface = matrix.surfaces.find(
    (candidate) => candidate.id === selection,
  );
  const specs = surface ? surface.visual : [normalizeRelative(selection)];
  const projects = options.project ?? [];
  const baseUrl = options.baseUrl ?? process.env.A3S_OFFICE_VISUAL_BASE_URL;
  const outputDir = path.join(
    repositoryRoot,
    matrix.shared.evidenceRoot,
    'visual',
  );
  const env: NodeJS.ProcessEnv = { ...process.env };
  if (baseUrl) env.A3S_OFFICE_VISUAL_BASE_URL = baseUrl;

  run('bun', ['run', 'playground:build:ui'], { env });
  const playwrightArgs = [
    'playwright',
    'test',
    ...specs,
    '--config',
    matrix.shared.visualConfig,
    '--output',
    outputDir,
  ];
  for (const project of projects) playwrightArgs.push('--project', project);
  run('bunx', playwrightArgs, { env });
}

function runWpsProbe(options: WpsProbeOption): void {
  if (process.platform !== 'win32') {
    throw new Error('WPS COM probing is only available on Windows.');
  }
  const target = options.output
    ? path.resolve(repositoryRoot, options.output)
    : path.join(
        repositoryRoot,
        matrix.shared.evidenceRoot,
        'wps',
        'probe.docx',
      );
  const commandArgs = [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    path.join(repositoryRoot, 'scripts', 'probe-wps-shapes.ps1'),
    '-OutputPath',
    target,
  ];
  if (options.connector || options.connectorType) {
    commandArgs.push('-IncludeConnector');
    commandArgs.push(
      '-ConnectorType',
      wpsConnectorTypeCodes[options.connectorType ?? 'straight'],
    );
  }
  run('powershell.exe', commandArgs);
}

function runWpsFieldsProbe(options: WpsFieldsProbeOption): void {
  if (process.platform !== 'win32') {
    throw new Error('WPS COM probing is only available on Windows.');
  }
  const target = options.output
    ? path.resolve(repositoryRoot, options.output)
    : path.join(
        repositoryRoot,
        matrix.shared.evidenceRoot,
        'wps',
        'numeric-fields.docx',
      );
  run('powershell.exe', [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    path.join(repositoryRoot, 'scripts', 'probe-wps-fields.ps1'),
    '-OutputPath',
    target,
  ]);
}

async function runGate(surfaceId: string, options: GateOption): Promise<void> {
  resolveSurface(surfaceId);
  runFixtures();
  runChecks(surfaceId, options.json === true);
  if (options.run) {
    await runA3s(surfaceId, options);
  }
  if (options.visual !== false) {
    runVisual(surfaceId, {
      baseUrl: options.baseUrl,
    });
  }
}

function resolveSuites(selection: string): string[] {
  return selection === 'all'
    ? matrix.surfaces.flatMap((surface) => surface.acl)
    : selection.endsWith('.acl')
      ? [normalizeRelative(selection)]
      : resolveSurface(selection).acl;
}
