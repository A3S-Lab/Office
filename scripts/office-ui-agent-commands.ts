import { type Command, InvalidArgumentError } from 'commander';
import {
  collectModifier,
  collectValue,
  parseAgentTarget,
  parseInteger,
  parsePositiveInteger,
  parsePositiveNumber,
  targetHelp,
  targetSyntax,
} from './office-ui-agent-targets';

export type AgentCommandRunner = (args: string[]) => void;

type AgentSessionOptions = {
  session: string;
  observation?: string;
  json?: boolean;
};

type AgentTargetOptions = AgentSessionOptions & {
  target: string;
  value?: string;
};

type AgentSelectOptions = AgentSessionOptions & {
  target: string;
  value: string[];
};

type AgentDragOptions = AgentSessionOptions & {
  source: string;
  target: string;
};

type AgentWheelOptions = AgentSessionOptions & {
  target?: string;
  deltaX?: string;
  deltaY: string;
  modifier?: string[];
};

type AgentViewportOptions = AgentSessionOptions & {
  width: string;
  height: string;
  scale?: string;
};

type AgentWaitOptions = AgentSessionOptions & {
  value: string;
};

type AgentEvidenceOptions = AgentSessionOptions & {
  path: string;
  interactive?: boolean;
  clear?: boolean;
};

const targetActionSpecs = [
  ['click', 'click', 'Click a semantic or CSS target.'],
  ['hover', 'hover', 'Hover a semantic or CSS target.'],
  ['focus', 'focus', 'Focus a semantic or CSS target.'],
  ['double-click', 'double_click', 'Double-click a target.'],
  ['context-click', 'context_click', 'Open the page-scoped context menu.'],
  ['check', 'check', 'Check a checkbox target.'],
  ['uncheck', 'uncheck', 'Uncheck a checkbox target.'],
] as const;

export function registerAgentActionCommands(
  agent: Command,
  run: AgentCommandRunner,
): void {
  for (const [name, actionType, description] of targetActionSpecs) {
    registerTargetCommand(agent, run, name, actionType, description);
  }

  registerTargetCommand(
    agent,
    run,
    'fill',
    'fill',
    'Fill a target with a value.',
    true,
  );
  registerTargetCommand(
    agent,
    run,
    'type',
    'type',
    'Type into a target.',
    true,
  );

  const select = baseCommand(
    agent,
    'select',
    'Select one or more values in a target.',
  )
    .requiredOption('--target <target>', targetHelp)
    .requiredOption(
      '--value <value>',
      'selected value; repeat for multiple values',
      collectValue,
      [],
    );
  addSessionOptions(select);
  select.action((options: AgentSelectOptions) => {
    runAgentAction(run, options, {
      type: 'select',
      target: parseAgentTarget(options.target),
      values: options.value,
    });
  });

  const drag = baseCommand(agent, 'drag', 'Drag one target to another target.')
    .requiredOption('--source <target>', `source target (${targetSyntax})`)
    .requiredOption('--target <target>', targetHelp);
  addSessionOptions(drag);
  drag.action((options: AgentDragOptions) => {
    runAgentAction(run, options, {
      type: 'drag',
      source: parseAgentTarget(options.source),
      target: parseAgentTarget(options.target),
    });
  });

  const press = baseCommand(
    agent,
    'press',
    'Press a keyboard key or chord.',
  ).requiredOption(
    '--key <key>',
    'key name or chord, for example Enter or Control+z',
  );
  addSessionOptions(press);
  press.action((options: AgentSessionOptions & { key: string }) => {
    runAgentAction(run, options, { type: 'press', key: options.key });
  });

  const wheel = baseCommand(
    agent,
    'wheel',
    'Scroll by a typed delta, optionally scoped to a target.',
  )
    .requiredOption('--delta-y <pixels>', 'vertical wheel delta')
    .option('--delta-x <pixels>', 'horizontal wheel delta', '0')
    .option('--target <target>', targetHelp)
    .option(
      '--modifier <key>',
      'modifier key; repeatable',
      collectModifier,
      [],
    );
  addSessionOptions(wheel);
  wheel.action((options: AgentWheelOptions) => {
    const deltaX = parseInteger(options.deltaX ?? '0', '--delta-x');
    const deltaY = parseInteger(options.deltaY, '--delta-y');
    if (deltaX === 0 && deltaY === 0) {
      throw new InvalidArgumentError(
        'At least one wheel delta must be non-zero.',
      );
    }
    runAgentAction(run, options, {
      type: 'wheel',
      ...(options.target ? { target: parseAgentTarget(options.target) } : {}),
      delta_x: deltaX,
      delta_y: deltaY,
      modifiers: options.modifier,
    });
  });

  const viewport = baseCommand(
    agent,
    'viewport',
    'Set the browser viewport for responsive evidence.',
  )
    .requiredOption('--width <pixels>', 'viewport width')
    .requiredOption('--height <pixels>', 'viewport height')
    .option('--scale <factor>', 'device scale factor');
  addSessionOptions(viewport);
  viewport.action((options: AgentViewportOptions) => {
    runAgentAction(run, options, {
      type: 'viewport',
      width: parsePositiveInteger(options.width, '--width'),
      height: parsePositiveInteger(options.height, '--height'),
      ...(options.scale === undefined
        ? {}
        : { scale: parsePositiveNumber(options.scale, '--scale') }),
    });
  });

  registerWaitCommands(agent, run);
  registerAssertionCommands(agent, run);
  registerEvidenceCommands(agent, run);
}

function registerTargetCommand(
  agent: Command,
  run: AgentCommandRunner,
  name: string,
  actionType: string,
  description: string,
  withValue = false,
): void {
  const command = baseCommand(agent, name, description).requiredOption(
    '--target <target>',
    targetHelp,
  );
  if (withValue)
    command.requiredOption('--value <value>', 'value to send to the target');
  addSessionOptions(command);
  command.action((options: AgentTargetOptions) => {
    runAgentAction(run, options, {
      type: actionType,
      target: parseAgentTarget(options.target),
      ...(withValue ? { value: options.value } : {}),
    });
  });
}

function registerWaitCommands(agent: Command, run: AgentCommandRunner): void {
  const text = baseCommand(
    agent,
    'wait-text',
    'Wait until text is visible.',
  ).requiredOption('--value <text>', 'text to wait for');
  addSessionOptions(text);
  text.action((options: AgentWaitOptions) => {
    runAgentAction(run, options, {
      type: 'wait',
      condition: { type: 'text', value: options.value },
    });
  });

  const url = baseCommand(
    agent,
    'wait-url',
    'Wait until the current URL matches a value.',
  ).requiredOption('--value <url>', 'URL or URL pattern');
  addSessionOptions(url);
  url.action((options: AgentWaitOptions) => {
    runAgentAction(run, options, {
      type: 'wait',
      condition: { type: 'url', value: options.value },
    });
  });

  const visible = baseCommand(
    agent,
    'wait-visible',
    'Wait until a target is visible.',
  ).requiredOption('--target <target>', targetHelp);
  addSessionOptions(visible);
  visible.action((options: AgentTargetOptions) => {
    runAgentAction(run, options, {
      type: 'wait',
      condition: { type: 'visible', value: parseAgentTarget(options.target) },
    });
  });

  for (const state of ['networkidle', 'domcontentloaded'] as const) {
    const load = baseCommand(agent, `wait-${state}`, `Wait for ${state}.`);
    addSessionOptions(load);
    load.action((options: AgentSessionOptions) => {
      runAgentAction(run, options, {
        type: 'wait',
        condition: { type: 'load', value: state },
      });
    });
  }
}

function registerAssertionCommands(
  agent: Command,
  run: AgentCommandRunner,
): void {
  const text = baseCommand(
    agent,
    'assert-text',
    'Assert that text is visible.',
  ).requiredOption('--value <text>', 'text to assert');
  addSessionOptions(text);
  text.action((options: AgentWaitOptions) => {
    runAgentAction(run, options, {
      type: 'assert',
      expectation: { type: 'text_visible', value: options.value },
    });
  });

  const url = baseCommand(
    agent,
    'assert-url',
    'Assert the current URL.',
  ).requiredOption('--value <url>', 'URL or URL pattern');
  addSessionOptions(url);
  url.action((options: AgentWaitOptions) => {
    runAgentAction(run, options, {
      type: 'assert',
      expectation: { type: 'url', value: options.value },
    });
  });

  const visible = baseCommand(
    agent,
    'assert-visible',
    'Assert that a target is visible.',
  ).requiredOption('--target <target>', targetHelp);
  addSessionOptions(visible);
  visible.action((options: AgentTargetOptions) => {
    runAgentAction(run, options, {
      type: 'assert',
      expectation: { type: 'visible', value: parseAgentTarget(options.target) },
    });
  });
}

function registerEvidenceCommands(
  agent: Command,
  run: AgentCommandRunner,
): void {
  const screenshot = baseCommand(
    agent,
    'screenshot',
    'Capture a bounded screenshot artifact.',
  ).requiredOption('--path <path>', 'session-relative PNG path');
  addSessionOptions(screenshot);
  screenshot.action((options: AgentEvidenceOptions) => {
    runAgentAction(run, options, { type: 'screenshot', path: options.path });
  });

  const accessibility = baseCommand(
    agent,
    'accessibility',
    'Capture the accessibility tree.',
  )
    .requiredOption('--path <path>', 'session-relative JSON path')
    .option('--interactive', 'include only interactive targets');
  addSessionOptions(accessibility);
  accessibility.action((options: AgentEvidenceOptions) => {
    runAgentAction(run, options, {
      type: 'accessibility',
      path: options.path,
      interactive: options.interactive === true,
    });
  });

  for (const [name, actionType] of [
    ['console', 'console'],
    ['page-errors', 'page_errors'],
  ] as const) {
    const diagnostics = baseCommand(agent, name, `Capture ${name} diagnostics.`)
      .requiredOption('--path <path>', 'session-relative JSON path')
      .option('--clear', 'clear the current buffer after capture');
    addSessionOptions(diagnostics);
    diagnostics.action((options: AgentEvidenceOptions) => {
      runAgentAction(run, options, {
        type: actionType,
        path: options.path,
        clear: options.clear === true,
      });
    });
  }
}

function baseCommand(
  agent: Command,
  name: string,
  description: string,
): Command {
  return agent.command(name).description(description);
}

function addSessionOptions(command: Command): void {
  command
    .requiredOption('--session <id>', 'active A3S Test agent session')
    .option('--observation <id>', 'latest observation for ref targets')
    .option('--json', 'emit machine-readable JSON');
}

function runAgentAction(
  run: AgentCommandRunner,
  options: AgentSessionOptions,
  action: Record<string, unknown>,
): void {
  run([
    'act',
    '--session',
    options.session,
    ...(options.observation ? ['--observation', options.observation] : []),
    '--action-json',
    JSON.stringify(action),
    ...(options.json ? ['--json'] : []),
  ]);
}
