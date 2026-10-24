import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { access, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { downloadAndUnzipVSCode, resolveCliArgsFromVSCodeExecutablePath, runTests } from '@vscode/test-electron';
import { build } from 'esbuild';

const root = resolve(import.meta.dirname, '..');
const target = `${process.platform}-${process.arch}`;
assert.ok(['linux-x64', 'win32-x64'].includes(target), `Unqualified test platform: ${target}`);
const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
const vsix = join(root, 'output', 'vsix', `${manifest.name}-${target}-${manifest.version}.vsix`);
await access(vsix).catch(() => {
  throw new Error(`Build the tested VSIX first: npm run package -- --target ${target}`);
});
const directory = await mkdtemp(join(tmpdir(), 'excel-vscode-smoke-'));
try {
  const workspace = join(directory, 'synthetic-workspace');
  const userData = join(directory, 'user-data');
  const extensions = join(directory, 'extensions');
  const screenshots = join(root, 'output', 'extension-host');
  await mkdir(workspace);
  await mkdir(join(userData, 'User'), { recursive: true });
  await mkdir(extensions);
  await mkdir(screenshots, { recursive: true });
  const exceljsFixture = join(directory, 'exceljs-fixture.cjs');
  // This browser-oriented dependency requires bundling even in the Node test host.
  await build({
    stdin: { contents: "export {default as ExcelJS} from '@cweijan/exceljs';", resolveDir: root },
    bundle: true, format: 'cjs', platform: 'node', target: 'node20',
    outfile: exceljsFixture, logLevel: 'silent',
  });
  await writeFile(join(userData, 'User', 'settings.json'), JSON.stringify({
    'telemetry.telemetryLevel': 'off',
    'update.mode': 'none',
    'extensions.autoUpdate': false,
    'extensions.autoCheckUpdates': false,
    'workbench.startupEditor': 'none',
    'window.restoreWindows': 'none',
    'workbench.editor.enablePreview': false,
    'files.simpleDialog.enable': true,
  }));
  const vscodeExecutablePath = await downloadAndUnzipVSCode({
    version: process.env.VSCODE_TEST_VERSION || '1.95.3',
    cachePath: join(root, '.vscode-test'),
  });
  const [cli, ...cliArgs] = resolveCliArgsFromVSCodeExecutablePath(vscodeExecutablePath, { reuseMachineInstall: true });
  const profileArgs = [`--user-data-dir=${userData}`, `--extensions-dir=${extensions}`];
  // The Windows CLI is a .cmd. Quote each path; all paths here are generated locally.
  const installArgs = [...cliArgs, ...profileArgs, '--install-extension', vsix, '--force'];
  const install = spawnSync(
    process.platform === 'win32' ? `"${cli}"` : cli,
    process.platform === 'win32' ? installArgs.map(value => `"${value}"`) : installArgs,
    { encoding: 'utf8', shell: process.platform === 'win32', timeout: 120_000 },
  );
  if (install.error) throw install.error;
  assert.equal(install.status, 0, `VSIX installation failed: ${install.stderr}\n${install.stdout}`);
  const listener = createServer();
  const port = await new Promise((resolvePort, reject) => {
    listener.once('error', reject);
    listener.listen(0, '127.0.0.1', () => {
      const selectedPort = listener.address().port;
      listener.close(error => error ? reject(error) : resolvePort(selectedPort));
    });
  });
  await runTests({
    vscodeExecutablePath,
    extensionDevelopmentPath: join(root, 'test', 'extension-host'),
    extensionTestsPath: join(root, 'test', 'extension-host', 'suite.cjs'),
    extensionTestsEnv: {
      EXCEL_SMOKE_WORKSPACE: workspace,
      EXCEL_SMOKE_CDP_ENDPOINT: `http://127.0.0.1:${port}`,
      EXCEL_SMOKE_EXCELJS_FIXTURE: exceljsFixture,
      EXCEL_SMOKE_SCREENSHOTS: screenshots,
    },
    launchArgs: [
      workspace, ...profileArgs, `--remote-debugging-port=${port}`,
      '--remote-debugging-address=127.0.0.1', '--disable-telemetry', '--locale=en',
      '--password-store=basic', '--use-mock-keychain',
    ],
  });
} finally {
  await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 250 });
}
