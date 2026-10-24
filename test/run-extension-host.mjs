import { downloadAndUnzipVSCode, runTests } from '@vscode/test-electron';
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { resolveWindowsCliPath } from './helpers/vscode-cli.mjs';

const root = resolve(import.meta.dirname, '..');
const temporary = await mkdtemp(join(tmpdir(), 'excel-host-smoke-'));
try {
  const version = process.argv[2] || 'stable';
  let extensionDevelopmentPath = root;
  let vscodeExecutablePath;
  if (process.argv[3]) {
    if (process.argv[3] !== '--packaged' || process.platform !== 'win32' || process.arch !== 'x64') {
      throw new Error('Usage: node test/run-extension-host.mjs [version] [--packaged (Windows x64 only)]');
    }
    const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
    const vsix = process.argv[4] ? resolve(process.argv[4])
      : join(root, 'output/vsix', `${manifest.name}-win32-x64-${manifest.version}.vsix`);
    vscodeExecutablePath = await downloadAndUnzipVSCode(version);
    // Invoke the official CLI through its Electron runtime with argument boundaries
    // preserved, even when the VSIX/profile path contains spaces. Never install into
    // the user's normal VS Code profile.
    const install = spawnSync(vscodeExecutablePath, [
      await resolveWindowsCliPath(vscodeExecutablePath),
      '--install-extension', vsix, '--force',
      '--user-data-dir', join(temporary, 'profile'), '--extensions-dir', join(temporary, 'extensions'),
    ], { encoding: 'utf8', windowsHide: true, timeout: 120_000,
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' } });
    if (install.error) throw install.error;
    if (install.status !== 0) throw new Error(`VSIX installation failed: ${install.stderr}\n${install.stdout}`);
    console.log(install.stdout);
    const installed = (await readdir(join(temporary, 'extensions'), { withFileTypes: true }))
      .filter(entry => entry.isDirectory() && entry.name.startsWith(`${manifest.publisher}.${manifest.name}-`));
    if (installed.length !== 1) throw new Error('Expected one isolated installed Excel extension.');
    extensionDevelopmentPath = join(temporary, 'extensions', installed[0].name);
    const packaged = JSON.parse(await readFile(join(extensionDevelopmentPath, 'package.json'), 'utf8'));
    if (packaged.version !== manifest.version || packaged.name !== manifest.name || packaged.publisher !== manifest.publisher) {
      throw new Error('Installed VSIX identity differs from the requested release.');
    }
    console.log(`PACKAGED_INSTALL_PASSED ${manifest.version} in isolated profile.`);
  }
  await runTests({
    version, vscodeExecutablePath,
    extensionDevelopmentPath,
    extensionTestsPath: join(root, 'test/extension-host-smoke.cjs'),
    extensionTestsEnv: { EXCEL_AI_SMOKE_WORKSPACE: temporary,
      EXCEL_AI_SMOKE_EXTENSION: extensionDevelopmentPath, ELECTRON_RUN_AS_NODE: undefined },
    launchArgs: [temporary, '--disable-extensions', '--disable-gpu', '--skip-welcome', '--skip-release-notes',
      '--disable-workspace-trust', `--user-data-dir=${join(temporary, 'profile')}`, `--extensions-dir=${join(temporary, 'extensions')}`],
  });
} finally {
  await rm(temporary, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
}
