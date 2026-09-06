import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';

const repositoryRoot = path.resolve(import.meta.dirname, '..');

test('keeps the WPS UI reference probe typed and isolated', async () => {
  const [cli, probe] = await Promise.all([
    readFile(path.join(repositoryRoot, 'scripts/office-ui-ops.ts'), 'utf8'),
    readFile(path.join(repositoryRoot, 'scripts/probe-wps-ui.ps1'), 'utf8'),
  ]);

  expect(cli).toContain(".command('wps-ui-probe')");
  expect(cli).toContain(".choices(['shell', 'fields', 'all'])");
  expect(cli).toContain("'probe-wps-ui.ps1'");
  expect(probe).toContain("[ValidateSet('shell', 'fields', 'all')]");
  expect(probe).toContain(
    "$automationClsid = [guid]'{000209FF-0000-4b30-A977-D214852036FF}'",
  );
  expect(probe).toContain('$document = $application.Documents.Add()');
  expect(probe).toContain('$document.Close(0)');
  expect(probe).toContain('$application.Quit(0)');
  expect(probe).toContain('Stop-Process -Force');
  expect(probe).toContain("probe = 'wps-writer-ui'");
  expect(probe).not.toContain('taskkill');
});
