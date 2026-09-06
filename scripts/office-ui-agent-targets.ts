import { InvalidArgumentError } from 'commander';

export type AgentTarget =
  | { type: 'ref'; value: string }
  | { type: 'css'; selector: string }
  | { type: 'role'; role: string; name: string }
  | { type: 'text'; value: string; exact: boolean }
  | { type: 'test_id'; value: string }
  | { type: 'label'; value: string }
  | { type: 'placeholder'; value: string }
  | { type: 'automation_id'; value: string };

export const targetSyntax =
  '@e7 | css=<selector> | role=<role>|<name> | label=<text> | placeholder=<text> | testid=<id> | automation=<id> | text=<text>';

export const targetHelp = `target (${targetSyntax})`;

const modifierKeys = ['alt', 'control', 'meta', 'shift'] as const;

const targetParsers: Array<{
  matches: (value: string) => boolean;
  parse: (value: string) => AgentTarget;
}> = [
  {
    matches: (value) => /^@[a-z][a-z0-9]*$/iu.test(value),
    parse: (value) => ({ type: 'ref', value }),
  },
  ...[
    [
      'css=',
      (value: string): AgentTarget => ({ type: 'css', selector: value }),
    ],
    ['label=', (value: string): AgentTarget => ({ type: 'label', value })],
    [
      'placeholder=',
      (value: string): AgentTarget => ({ type: 'placeholder', value }),
    ],
    ['testid=', (value: string): AgentTarget => ({ type: 'test_id', value })],
    [
      'automation=',
      (value: string): AgentTarget => ({ type: 'automation_id', value }),
    ],
    [
      'text=',
      (value: string): AgentTarget => ({ type: 'text', value, exact: true }),
    ],
  ].map(([prefix, create]) => ({
    matches: (value: string) =>
      value.startsWith(prefix as string) &&
      value.slice((prefix as string).length).trim().length > 0,
    parse: (value: string) =>
      (create as (value: string) => AgentTarget)(
        value.slice((prefix as string).length),
      ),
  })),
  {
    matches: (value) =>
      value.startsWith('role=') && value.slice(5).includes('|'),
    parse: (value) => {
      const [role, ...nameParts] = value.slice(5).split('|');
      const name = nameParts.join('|').trim();
      if (!role.trim() || !name) {
        throw new InvalidArgumentError('role= must use role=<role>|<name>.');
      }
      return { type: 'role', role: role.trim(), name };
    },
  },
];

export function parseAgentTarget(raw: string): AgentTarget {
  const spec = raw.trim();
  if (!spec) throw new InvalidArgumentError('Agent target must not be empty.');

  const parser = targetParsers.find(({ matches }) => matches(spec));
  if (!parser) {
    throw new InvalidArgumentError(
      `Unsupported agent target '${raw}'. Use ${targetSyntax}.`,
    );
  }
  return parser.parse(spec);
}

export function collectValue(value: string, previous: string[]): string[] {
  const trimmed = value.trim();
  if (!trimmed) throw new InvalidArgumentError('The value must not be empty.');
  return [...previous, trimmed];
}

export function collectModifier(value: string, previous: string[]): string[] {
  const normalized = value.trim().toLowerCase();
  if (!modifierKeys.includes(normalized as (typeof modifierKeys)[number])) {
    throw new InvalidArgumentError(
      `Modifier must be one of ${modifierKeys.join(', ')}.`,
    );
  }
  return [...previous, normalized];
}

export function parseInteger(value: string, flag: string): number {
  if (!/^-?\d+$/u.test(value)) {
    throw new InvalidArgumentError(`${flag} must be an integer.`);
  }
  return Number(value);
}

export function parsePositiveInteger(value: string, flag: string): number {
  const parsed = parseInteger(value, flag);
  if (parsed <= 0) throw new InvalidArgumentError(`${flag} must be positive.`);
  return parsed;
}

export function parsePositiveNumber(value: string, flag: string): number {
  if (!/^\d+(?:\.\d+)?$/u.test(value)) {
    throw new InvalidArgumentError(`${flag} must be a positive number.`);
  }
  const parsed = Number(value);
  if (!(parsed > 0))
    throw new InvalidArgumentError(`${flag} must be positive.`);
  return parsed;
}
