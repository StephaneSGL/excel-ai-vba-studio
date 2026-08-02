import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import {
  copyFileSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

if (process.platform !== 'win32') {
  console.log('VBA designer NTFS ADS security test skipped outside Windows.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const designerScript = join(root, 'scripts', 'apply-vba-designer.ps1');
const probeScript = join(root, 'test', 'vba-designer-ads-probe.ps1');
const helperPath = join(
  root,
  'bin',
  'win32-x64',
  'excel-ai-vba-writeback.exe',
);
const fixturePath = join(root, 'test', 'fixtures', 'DemoExcelUserForm.xlsm');
const powerShell = join(
  process.env.SystemRoot || 'C:\\Windows',
  'System32',
  'WindowsPowerShell',
  'v1.0',
  'powershell.exe',
);

const designerSource = readFileSync(designerScript, 'utf8');
assert.match(
  designerSource,
  /\$commitGuard\s*=\s*\[IO\.File\]::Open\([\s\S]*?\[IO\.FileShare\]::Read\s+-bor\s+\[IO\.FileShare\]::Delete\)/,
  'VBA designer commit must hold a read/delete-share writer guard',
);
assert.match(
  designerSource,
  /\[IO\.File\]::Replace\(\$stagingPath,\s*\$workbookPath,\s*\$backupPath\)[\s\S]*?\$commitGuard\.Dispose\(\)/,
  'VBA designer must keep the writer guard through atomic replacement verification',
);
assert.match(
  designerSource,
  /Stop-OwnedExcelProcess \$excel2Identity[\s\S]*?if \(\$null -ne \$commitGuard\)[\s\S]*?\$commitGuard\.Dispose\(\)[\s\S]*?# A failure after File\.Replace must restore/,
  'VBA designer must release the writer guard before post-commit rollback',
);

for (const requiredPath of [
  designerScript,
  probeScript,
  helperPath,
  fixturePath,
]) {
  assert.ok(existsSync(requiredPath), `Required ADS test input is missing: ${requiredPath}`);
}

const probeResult = spawnSync(
  powerShell,
  [
    '-NoLogo',
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    probeScript,
    '-DesignerScriptPath',
    designerScript,
  ],
  {
    cwd: root,
    encoding: 'utf8',
    shell: false,
    windowsHide: true,
    timeout: 30_000,
  },
);
assert.equal(
  probeResult.status,
  0,
  `NTFS ADS transaction probe failed:\n${probeResult.stdout}\n${probeResult.stderr}`,
);

const base64 = (value) => Buffer.from(value, 'utf8').toString('base64');
const testRoot = mkdtempSync(join(tmpdir(), 'excel-vba-designer-motw-'));
try {
  const workbookPath = join(testRoot, 'motw.xlsm');
  const requestPath = join(testRoot, 'request.json');
  copyFileSync(fixturePath, workbookPath);
  const expectedWorkbookSha256 = crypto
    .createHash('sha256')
    .update(readFileSync(workbookPath))
    .digest('hex');
  writeFileSync(
    requestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath,
      expectedWorkbookSha256,
      operations: [
        {
          kind: 'createUserForm',
          name: 'frmAdsPreflightProbe',
          caption: 'ADS preflight probe',
          width: 300,
          height: 200,
        },
      ],
    }),
    'utf8',
  );

  const cases = [
    {
      bytes: Buffer.from('[ZoneTransfer]\r\nZoneId=3\r\n', 'utf8'),
      expected: /Internet or Restricted Zone \(ZoneId=3\)/,
    },
    {
      bytes: Buffer.from(
        '[ZoneTransfer]\r\nZoneId=2\r\nZoneId=3\r\n',
        'utf8',
      ),
      expected: /Zone\.Identifier is missing or ambiguous/,
    },
    {
      bytes: Buffer.alloc(64 * 1024 + 1, 0x41),
      expected: /Zone\.Identifier exceeds the 65536-byte safety limit/,
    },
  ];

  for (const testCase of cases) {
    writeFileSync(`${workbookPath}:Zone.Identifier`, testCase.bytes);
    const result = spawnSync(
      powerShell,
      [
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        designerScript,
        '-RequestPathBase64',
        base64(requestPath),
        '-HelperPathBase64',
        base64(helperPath),
      ],
      {
        cwd: root,
        encoding: 'utf8',
        shell: false,
        windowsHide: true,
        timeout: 30_000,
        maxBuffer: 1024 * 1024,
      },
    );
    const output = `${result.stdout ?? ''}\n${result.stderr ?? ''}`;
    assert.notEqual(result.status, 0, 'Unsafe Zone.Identifier was accepted');
    assert.match(output, testCase.expected);
    assert.doesNotMatch(
      output,
      /OWNED_EXCEL_PID\|/,
      'Zone.Identifier must be refused before Excel COM is created',
    );
  }
} finally {
  rmSync(testRoot, { recursive: true, force: true });
}

console.log(
  'VBA designer ADS security passed: NTFS streams are bounded and preserved, and unsafe MotW is refused before COM.',
);
