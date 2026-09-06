import { expect, test } from '@rstest/core';
import { Command } from 'commander';
import { registerAgentActionCommands } from '../scripts/office-ui-agent-commands';
import { parseAgentTarget } from '../scripts/office-ui-agent-targets';

test('parses the stable target grammar into A3S Test targets', () => {
  expect(parseAgentTarget('@e7')).toEqual({ type: 'ref', value: '@e7' });
  expect(parseAgentTarget('role=button|Save')).toEqual({
    type: 'role',
    role: 'button',
    name: 'Save',
  });
  expect(parseAgentTarget('css=[data-testid=save]')).toEqual({
    type: 'css',
    selector: '[data-testid=save]',
  });
  expect(parseAgentTarget('text=Saved')).toEqual({
    type: 'text',
    value: 'Saved',
    exact: true,
  });
});

test('typed agent commands forward schema-valid action JSON', async () => {
  const program = new Command();
  const calls: string[][] = [];
  registerAgentActionCommands(program, (args) => calls.push(args));

  await program.parseAsync([
    'node',
    'office-ui-ops',
    'fill',
    '--session',
    'writer-ui',
    '--observation',
    '4',
    '--target',
    'label=Title',
    '--value',
    'A3S Office',
    '--json',
  ]);

  expect(calls).toHaveLength(1);
  expect(calls[0]).toEqual([
    'act',
    '--session',
    'writer-ui',
    '--observation',
    '4',
    '--action-json',
    JSON.stringify({
      type: 'fill',
      target: { type: 'label', value: 'Title' },
      value: 'A3S Office',
    }),
    '--json',
  ]);
});
