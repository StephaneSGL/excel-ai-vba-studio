import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import process from 'node:process';

const root = fileURLToPath(new URL('..', import.meta.url));
const vsceEntrypoint = resolve(root, 'node_modules', '@vscode', 'vsce', 'vsce');
const result = spawnSync(process.execPath, [vsceEntrypoint, 'ls', '--no-dependencies'], {
  cwd: root,
  encoding: 'utf8',
  shell: false,
});

if (result.status !== 0) {
  if (result.error) {
    process.stderr.write(`${result.error.message}\n`);
  }
  process.stderr.write(result.stdout ?? '');
  process.stderr.write(result.stderr ?? '');
  process.exit(result.status ?? 1);
}

const files = (result.stdout ?? '')
  .split(/\r?\n/)
  .map((line) => line.trim().replaceAll('\\', '/'))
  .filter(Boolean);

const errors = [];
const required = [
  'package.json',
  'README.md',
  'CHANGELOG.md',
  'SECURITY.md',
  'SUPPORT.md',
  'LICENSE',
  'LICENSING.md',
  'LICENSES/POLYFORM-NONCOMMERCIAL-1.0.0.md',
  'LICENSES/OFFICE-VIEWER-MIT.txt',
  'LICENSES/X-DATA-SPREADSHEET-MIT.txt',
  'LICENSES/VDITOR-MIT.txt',
  'LICENSES/PYOPENVBA-MIT.txt',
  'LICENSES/PYINSTALLER-GPL2-EXCEPTION.txt',
  'LICENSES/PYTHON-3.11-PSF.txt',
  'PRIVACY.md',
  'NOTICE.md',
  'THIRD_PARTY_NOTICES.md',
  'THIRD_PARTY_LICENSES.txt',
  'image/marketplace-icon.png',
  'out/extension.js',
  'out/webview/index.html',
  'scripts/office-ai-export.ps1',
  'scripts/office-ai-apply-edits.ps1',
  'scripts/apply-vba-designer.ps1',
  'scripts/inspect-office-security.ps1',
  'scripts/ooxml-package-signature.ps1',
  'scripts/open-excel-developer.ps1',
  'scripts/prepare-macro-workbook.ps1',
  'bin/win32-x64/excel-ai-vba-writeback.exe',
];
for (const file of required) {
  if (!files.includes(file)) {
    errors.push(`required packaged file is missing: ${file}`);
  }
}

const forbiddenPrefixes = [
  '.git/',
  '.github/',
  '.vscode/',
  'AI-slop/',
  'docs/',
  'icons/',
  'native/',
  'node_modules/',
  'output/',
  'outputs/',
  'src/',
  'syntaxes/',
  'test/',
  'theme/',
  'tools/',
  'vditor/',
  'work/',
];
const allowedFiles = new Set(required);
const allowedGeneratedAsset = /^out\/webview\/assets\/[A-Za-z0-9._-]+\.(?:js|css|ttf|woff2?)$/;
for (const file of files) {
  if (
    !allowedFiles.has(file) &&
    !allowedGeneratedAsset.test(file)
  ) {
    errors.push(`unexpected file is not on the package allowlist: ${file}`);
  }
  if (forbiddenPrefixes.some((prefix) => file.startsWith(prefix))) {
    errors.push(`development or unrelated file would be packaged: ${file}`);
  }
  if (file === 'telemetry.json' || file.startsWith('package.nls.')) {
    errors.push(`upstream/telemetry file would be packaged: ${file}`);
  }
}

if (errors.length > 0) {
  console.error(`Package validation failed (${errors.length}):`);
  for (const error of errors) {
    console.error(`- ${error}`);
  }
  process.exit(1);
}

console.log(`Package contents valid: ${files.length} files`);
