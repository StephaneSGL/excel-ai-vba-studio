import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { build } from 'esbuild';

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const probePath = path.join(root, 'scripts', 'inspect-office-security.ps1');
const probe = await readFile(probePath, 'utf8');
const modelSource = await readFile(path.join(root, 'src', 'excelAiVbaStudio', 'officeSecurity.ts'), 'utf8');
const panelSource = await readFile(path.join(root, 'src', 'excelAiVbaStudio', 'securityCenterPanel.ts'), 'utf8');
const workbookServiceSource = await readFile(path.join(root, 'src', 'excelAiVbaStudio', 'workbookService.ts'), 'utf8');
const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));

const forbiddenWritePatterns = [
  /\b(?:Set|New|Remove|Copy|Move|Rename)-Item(?:Property)?\b/i,
  /\b(?:Set|Add|Clear)-Content\b/i,
  /\bOut-File\b/i,
  /\bExport-(?:Csv|Clixml)\b/i,
  /\b(?:reg\.exe|reg)\s+(?:add|delete|copy|restore|import)\b/i,
  /\[IO\.File\]::(?:WriteAll|AppendAll|Create|OpenWrite|Replace|Move|Copy|Delete)/i,
  /\[Microsoft\.Win32\.RegistryKey\]::OpenRemoteBaseKey/i,
  /\.OpenSubKey\([^\r\n]+,\s*\$true\s*\)/i,
  /\.CreateSubKey\s*\(/i,
  /\.SetValue\s*\(/i,
  /\.Delete(?:SubKey|Value)/i,
];
for (const pattern of forbiddenWritePatterns) {
  assert.doesNotMatch(probe, pattern, `probe contains a write primitive: ${pattern}`);
}

for (const pattern of [
  /Excel\.Application/i,
  /New-Object\s+-ComObject/i,
  /\bStart-Process\b/i,
  /\bGet-Process\b/i,
  /\bStop-Process\b/i,
  /WScript\.Shell/i,
]) {
  assert.doesNotMatch(probe, pattern, `probe must not automate or manage processes: ${pattern}`);
}

for (const token of [
  'WorkbookPathBase64',
  'schemaVersion = 1',
  "version = '16.0'",
  'hasVbaProject',
  'hasVbaSignature',
  'hasPackageSignature',
  'sensitivityLabelIds',
  'VBAWarnings',
  'AccessVBOM',
  'BlockContentExecutionFromInternet',
  'DisableAllActiveX',
  'UFIControls',
  'DisableInternetFilesInPV',
  'DisableUnsafeLocationsInPV',
  'DisableAttachmentsInPV',
  'DisableAllTrustedLocations',
  'AllowNetworkLocations',
  'Allow User Locations',
  'Software\\Policies\\Microsoft\\Cloud\\Office\\16.0',
  'MaxTrustedLocations = 64',
  'MaxOutputBytes = 262144',
  'MaxWorkbookBytes = 536870912',
  'zoneStatus',
  'officePackageEncrypted',
  'vbaSignatureStatus',
  'packageSignatureStatus',
  'packageSignatureVerificationStatus',
  'vbaProjectProtectionStatus',
  'cloudPolicyServiceDetected',
  'EncryptionInfo',
  'EncryptedPackage',
  'Get-Sha256',
  'digital-signature/origin',
  'digital-signature/signature',
  'application/vnd.openxmlformats-package.digital-signature-origin',
  'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml',
  'MaxCharactersInDocument = 1048576',
]) {
  assert.ok(probe.includes(token), `probe is missing expected token: ${token}`);
}

assert.match(probe, /Get-Content[\s\S]+-Stream\s+'Zone\.Identifier'/, 'MOTW must be read as an alternate stream');
assert.doesNotMatch(probe, /HostUrl|ReferrerUrl/i, 'MOTW URLs must never be returned');
assert.match(probe, /DtdProcessing\s*=\s*\[Xml\.DtdProcessing\]::Prohibit/, 'custom XML must prohibit DTD processing');
assert.doesNotMatch(probe, /GetSubKeyNames\s*\(|GetValueNames\s*\(/, 'registry reads must not enumerate arbitrary names');
assert.match(probe, /for \(\$locationIndex = 0; \$locationIndex -lt 64; \$locationIndex\+\+\)/, 'trusted locations must use the exact bounded Location0..Location63 range');
assert.doesNotMatch(probe, /FileShare\][^\r\n]*Delete|FileShare\]::Delete/, 'workbook handles must not allow deletion during inspection');
assert.doesNotMatch(probe, /HKLM[^\r\n]+Microsoft\\Cloud\\Office\\16\.0/, 'Cloud Policy security detection must not use HKLM Cloud Update keys');
assert.match(probe, /_Enabled\(\?:V2\)\?\$/, 'Purview detection must require an Enabled or EnabledV2 property');
assert.ok(
  manifest.contributes.commands.some(({ command }) => command === 'excelAiVbaStudio.openSecurityCenter'),
  'manifest must expose the security center command',
);
assert.match(workbookServiceSource, /async openSecurityCenter[\s\S]+securityCenterPanel\.open/, 'workbook service must open the security panel');
assert.match(panelSource, /Content-Security-Policy[^\r\n]+default-src 'none'/, 'security panel must use a deny-by-default CSP');
assert.match(panelSource, /localResourceRoots:\s*\[\]/, 'security panel must not expose local resource roots');
assert.match(panelSource, /function escapeHtml[\s\S]+replace\(\/\[&<>"'\]\//, 'security panel must escape dynamic HTML');
assert.match(panelSource, /https:\/\/config\.office\.com/, 'security panel must use the official Microsoft 365 Apps admin portal');
assert.match(panelSource, /rôle administrateur autorisé/, 'security panel must explain the enterprise portal authorization boundary');
for (const action of [
  'refresh',
  'copyReport',
  'openExcelSecurity',
  'openExtensionSettings',
  'openEnterpriseAdmin',
  'openAdminDocs',
]) {
  assert.ok(panelSource.includes(`'${action}'`) || panelSource.includes(`"${action}"`), `security panel is missing action ${action}`);
}
for (const source of [modelSource, panelSource, workbookServiceSource]) {
  assert.doesNotMatch(
    source,
    /(?:Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|Unblock-File|AccessVBOM\s*=|VBAWarnings\s*=)/i,
    'TypeScript integration must not contain Office policy mutation primitives',
  );
}

const modelEntry = path.join(root, 'src', 'excelAiVbaStudio', 'officeSecurity.ts');
const modelBundle = await build({
  entryPoints: [modelEntry],
  bundle: true,
  format: 'cjs',
  platform: 'node',
  target: 'node22',
  write: false,
  logLevel: 'silent',
  plugins: [{
    name: 'vscode-stub',
    setup(buildApi) {
      buildApi.onResolve({ filter: /^vscode$/ }, () => ({ path: 'vscode', namespace: 'stub' }));
      buildApi.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: 'module.exports = {};', loader: 'js' }));
    },
  }],
});
const modelModule = { exports: {} };
new Function('module', 'exports', 'require', '__filename', '__dirname', modelBundle.outputFiles[0].text)(
  modelModule,
  modelModule.exports,
  require,
  modelEntry,
  path.dirname(modelEntry),
);
const { buildEnterpriseSecurityReport, parseOfficeSecurityProbe } = modelModule.exports;
assert.equal(typeof buildEnterpriseSecurityReport, 'function', 'security model must export its report builder');
assert.equal(typeof parseOfficeSecurityProbe, 'function', 'security model must export its bounded probe parser');

function syntheticProbe(overrides = {}) {
  return {
    schemaVersion: 1,
    inspectedAtUtc: '2026-08-01T00:00:00.000Z',
    workbook: {
      path: 'C:\\Work\\book.xlsm',
      name: 'book.xlsm',
      extension: '.xlsm',
      sizeBytes: 1024,
      sha256: '0'.repeat(64),
      readOnly: false,
      efsEncrypted: false,
      officePackageEncrypted: false,
      zoneId: null,
      zoneStatus: 'absent',
      containerKind: 'zip',
      hasVbaProject: true,
      hasVbaSignature: false,
      hasPackageSignature: false,
      vbaSignatureStatus: 'absent',
      packageSignatureStatus: 'absent',
      vbaProjectProtectionStatus: 'unknown',
      sensitivityLabelIds: [],
      ...(overrides.workbook ?? {}),
    },
    office: {
      version: '16.0',
      settings: [],
      trustedLocations: [],
      cloudPolicyDetected: false,
      cloudPolicyServiceDetected: false,
      ...(overrides.office ?? {}),
    },
  };
}

const sourceSetting = (source, id, value, managed = source.endsWith('Policy')) => ({
  id,
  category: 'test',
  source,
  managed,
  registryPath: `HKCU\\Synthetic\\${source}`,
  name: id,
  value,
});

const internetReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { zoneId: 3, zoneStatus: 'read' },
  office: {
    settings: [
      sourceSetting('userPreference', 'vbaWarnings', 1, false),
      sourceSetting('userPreference', 'blockInternetMacros', 1, false),
    ],
  },
}));
assert.equal(internetReport.level, 'restricted');
assert.equal(internetReport.findings.find(({ id }) => id === 'macros').status, 'blocked');
assert.equal(internetReport.capabilities.find(({ id }) => id === 'macroExecution').status, 'blocked');

const signedInternetReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    zoneId: 3,
    zoneStatus: 'read',
    hasVbaSignature: true,
    vbaSignatureStatus: 'present',
  },
  office: {
    settings: [sourceSetting('userPolicy', 'blockInternetMacros', 1)],
  },
}));
assert.equal(
  signedInternetReport.findings.find(({ id }) => id === 'macros').status,
  'unknown',
  'a VBA signature requires a trusted-publisher decision that the static probe cannot prove',
);

const cloudPolicyReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: {
    cloudPolicyDetected: true,
    settings: [
      sourceSetting('machinePolicy', 'vbaWarnings', 1),
      sourceSetting('cloudPolicy', 'vbaWarnings', 4),
    ],
  },
}));
const cloudMacroFinding = cloudPolicyReport.findings.find(({ id }) => id === 'macros');
assert.equal(cloudMacroFinding.status, 'blocked', 'Cloud Policy must take precedence over local Group Policy');
assert.match(cloudMacroFinding.source, /Cloud Policy/);

const signedReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { hasVbaSignature: true, vbaSignatureStatus: 'present' },
  office: { settings: [sourceSetting('userPreference', 'accessVbom', 1, false)] },
}));
assert.equal(signedReport.capabilities.find(({ id }) => id === 'vbaWrite').status, 'blocked');
assert.equal(signedReport.findings.find(({ id }) => id === 'signatures').status, 'protected');

const internetWithoutExplicitPolicyReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { zoneId: 3, zoneStatus: 'read' },
  office: { settings: [sourceSetting('userPreference', 'vbaWarnings', 1, false)] },
}));
assert.equal(
  internetWithoutExplicitPolicyReport.findings.find(({ id }) => id === 'macros').status,
  'unknown',
  'MOTW without an explicit Internet macro policy must not be reported as a proven block or permission',
);

const efsReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    path: 'C:\\Work\\protected.xlsx',
    name: 'protected.xlsx',
    extension: '.xlsx',
    hasVbaProject: false,
    vbaSignatureStatus: 'absent',
    packageSignatureStatus: 'absent',
    vbaProjectProtectionStatus: 'absent',
    efsEncrypted: true,
  },
}));
assert.equal(efsReport.capabilities.find(({ id }) => id === 'grid').status, 'allowed');
assert.equal(efsReport.findings.find(({ id }) => id === 'classification').status, 'protected');

const encryptedReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    containerKind: 'compound',
    officePackageEncrypted: true,
    vbaSignatureStatus: 'unknown',
    packageSignatureStatus: 'unknown',
  },
}));
assert.equal(encryptedReport.level, 'restricted');
assert.equal(encryptedReport.capabilities.find(({ id }) => id === 'grid').status, 'blocked');
assert.equal(encryptedReport.capabilities.find(({ id }) => id === 'vbaWrite').status, 'blocked');

const conflictingPolicyReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: {
    settings: [
      { ...sourceSetting('machinePolicy', 'vbaWarnings', 1), registryView: '64' },
      { ...sourceSetting('machinePolicy', 'vbaWarnings', 4), registryView: '32' },
    ],
  },
}));
assert.equal(
  conflictingPolicyReport.findings.find(({ id }) => id === 'macros').status,
  'unknown',
  'conflicting registry views must not be collapsed into a fabricated effective value',
);

const protectedProjectReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { vbaProjectProtectionStatus: 'present' },
  office: { settings: [sourceSetting('userPreference', 'accessVbom', 1, false)] },
}));
assert.equal(protectedProjectReport.capabilities.find(({ id }) => id === 'vbaWrite').status, 'blocked');
assert.equal(protectedProjectReport.capabilities.find(({ id }) => id === 'designer').status, 'blocked');

const packageSignedReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    path: 'C:\\Work\\signed.xlsx',
    name: 'signed.xlsx',
    extension: '.xlsx',
    hasVbaProject: false,
    packageSignatureStatus: 'present',
    hasPackageSignature: true,
    vbaSignatureStatus: 'absent',
    vbaProjectProtectionStatus: 'absent',
  },
}));
assert.equal(packageSignedReport.level, 'restricted');
assert.equal(packageSignedReport.capabilities.find(({ id }) => id === 'grid').status, 'protected');
assert.equal(packageSignedReport.capabilities.find(({ id }) => id === 'vbaWrite').status, 'blocked');

const trustedLocationReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { path: 'C:\\Work\\Department\\book.xlsm' },
  office: {
    trustedLocations: [{
      source: 'userPreference',
      managed: false,
      registryPath: 'HKCU\\Synthetic\\Location0',
      path: 'C:\\Work',
      allowSubfolders: true,
    }],
  },
}));
assert.equal(trustedLocationReport.workbookInTrustedLocation, true);

const userTrustedLocationsBlockedReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { path: 'C:\\Work\\Department\\book.xlsm', zoneId: 3, zoneStatus: 'read' },
  office: {
    settings: [
      sourceSetting('userPolicy', 'allowUserTrustedLocations', 0),
      sourceSetting('userPolicy', 'blockInternetMacros', 1),
    ],
    trustedLocations: [{
      source: 'userPreference',
      managed: false,
      registryPath: 'HKCU\\Synthetic\\Location0',
      path: 'C:\\Work',
      allowSubfolders: true,
    }],
  },
}));
assert.equal(userTrustedLocationsBlockedReport.workbookInTrustedLocation, false);
assert.equal(userTrustedLocationsBlockedReport.findings.find(({ id }) => id === 'trustedLocations').status, 'blocked');
assert.equal(userTrustedLocationsBlockedReport.findings.find(({ id }) => id === 'macros').status, 'blocked');

const accessVbomBlockedReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: { settings: [sourceSetting('userPolicy', 'accessVbom', 0)] },
}));
assert.equal(accessVbomBlockedReport.capabilities.find(({ id }) => id === 'designer').status, 'blocked');
assert.equal(accessVbomBlockedReport.capabilities.find(({ id }) => id === 'activeXWrite').status, 'blocked');

const readOnlyReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { readOnly: true },
}));
assert.equal(readOnlyReport.capabilities.find(({ id }) => id === 'grid').status, 'protected');

const purviewOnlyReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { sensitivityLabelIds: ['7e4bc8d6-6897-4f7e-861d-25bf6f908374'] },
}));
assert.notEqual(purviewOnlyReport.level, 'managed', 'file metadata alone must not imply managed Office policy');
assert.equal(purviewOnlyReport.findings.find(({ id }) => id === 'classification').managed, false);

const legacyReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    path: 'C:\\Work\\legacy.xls',
    name: 'legacy.xls',
    extension: '.xls',
    containerKind: 'compound',
  },
}));
assert.equal(legacyReport.capabilities.find(({ id }) => id === 'grid').status, 'protected');
assert.notEqual(legacyReport.findings.find(({ id }) => id === 'classification').status, 'blocked');

const xlsbReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    path: 'C:\\Work\\binary.xlsb',
    name: 'binary.xlsb',
    extension: '.xlsb',
  },
}));
assert.equal(xlsbReport.capabilities.find(({ id }) => id === 'grid').status, 'blocked');

if (process.platform === 'win32') {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'excel-security-probe-'));
  const workbookPath = path.join(temporaryDirectory, 'synthetic-security.xlsx');
  const labelId = '7e4bc8d6-6897-4f7e-861d-25bf6f908374';
  const disabledLabelId = '82c03a0e-daad-4878-b50a-a1d1a42f16c7';
  try {
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '</Types>');
    zip.file('xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    zip.file('docProps/custom.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">' +
      `<property name="MSIP_Label_${labelId}_Enabled" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2">` +
      '<value xmlns="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">true</value>' +
      '</property>' +
      `<property name="MSIP_Label_${disabledLabelId}_EnabledV2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3">` +
      '<value xmlns="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">false</value>' +
      '</property></Properties>');
    zip.file('_xmlsignatures/origin.sigs', 'origin-without-a-signature-part');
    const workbookBytes = await zip.generateAsync({ type: 'nodebuffer' });
    await writeFile(workbookPath, workbookBytes);
    const sha256Before = createHash('sha256').update(workbookBytes).digest('hex');
    const neighborsBefore = (await readdir(temporaryDirectory)).sort();

    const encodedPath = Buffer.from(workbookPath, 'utf8').toString('base64');
    const powershell = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
      : 'powershell.exe';
    const { stdout, stderr } = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', encodedPath],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    assert.equal(stderr.trim(), '', 'probe must not emit non-JSON diagnostics');
    const lines = stdout.trim().split(/\r?\n/);
    assert.equal(lines.length, 1, 'probe must emit exactly one JSON object');
    assert.ok(Buffer.byteLength(lines[0], 'utf8') <= 256 * 1024, 'probe JSON must stay within 256 KiB');
    const result = JSON.parse(lines[0]);
	const parsedProbe = parseOfficeSecurityProbe(result);
    assert.equal(result.schemaVersion, 1);
    assert.match(result.inspectedAtUtc, /^\d{4}-\d{2}-\d{2}T/);
    assert.equal(result.workbook.path, workbookPath);
    assert.equal(result.workbook.name, path.basename(workbookPath));
    assert.equal(result.workbook.extension, '.xlsx');
    assert.equal(result.workbook.containerKind, 'zip');
    assert.equal(result.workbook.hasVbaProject, false);
    assert.equal(result.workbook.hasVbaSignature, false);
    assert.equal(result.workbook.hasPackageSignature, false);
    assert.equal(result.workbook.vbaSignatureStatus, 'absent');
    assert.equal(result.workbook.packageSignatureStatus, 'unknown');
    assert.equal(result.workbook.packageSignatureVerificationStatus, 'unverifiable');
    assert.equal(result.workbook.officePackageEncrypted, false);
    assert.equal(result.workbook.vbaProjectProtectionStatus, 'absent');
    assert.deepEqual(result.workbook.sensitivityLabelIds, [labelId]);
    assert.equal(result.workbook.zoneId, null);
    assert.ok(['absent', 'unsupported'].includes(result.workbook.zoneStatus));
    assert.equal(result.workbook.sha256, sha256Before);
	assert.equal(parsedProbe.workbook.sha256, sha256Before, 'the TypeScript boundary must accept the live probe schema');
	assert.equal(buildEnterpriseSecurityReport(parsedProbe).probe.workbook.path, workbookPath);
    assert.equal(
      createHash('sha256').update(await readFile(workbookPath)).digest('hex'),
      sha256Before,
      'probe must leave workbook bytes unchanged',
    );
    assert.deepEqual(
      (await readdir(temporaryDirectory)).sort(),
      neighborsBefore,
      'probe must not create neighboring files',
    );
    assert.equal(result.office.version, '16.0');
    assert.ok(Array.isArray(result.office.settings));
    assert.ok(Array.isArray(result.office.trustedLocations));
    assert.ok(result.office.trustedLocations.length <= 64);
    assert.equal(typeof result.office.cloudPolicyDetected, 'boolean');
    assert.equal(typeof result.office.cloudPolicyServiceDetected, 'boolean');
    for (const setting of result.office.settings) {
      assert.ok(['machinePolicy', 'userPolicy', 'cloudPolicy', 'userPreference', 'machinePreference'].includes(setting.source));
      assert.equal(typeof setting.managed, 'boolean');
      assert.equal(typeof setting.registryPath, 'string');
      assert.equal(typeof setting.name, 'string');
      assert.ok(Object.hasOwn(setting, 'value'));
      assert.ok(
        setting.value === null || ['string', 'number', 'boolean'].includes(typeof setting.value),
        'registry values must remain bounded JSON scalars',
      );
    }

    await writeFile(`${workbookPath}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=3\r\n');
    const zoneExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', encodedPath],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const zoneResult = JSON.parse(zoneExecution.stdout.trim());
    assert.equal(zoneResult.workbook.zoneStatus, 'read');
    assert.equal(zoneResult.workbook.zoneId, 3);
    assert.equal(zoneResult.workbook.sha256, sha256Before, 'MOTW inspection must not alter workbook content');
    assert.deepEqual((await readdir(temporaryDirectory)).sort(), neighborsBefore, 'an ADS must not be treated as a neighboring file');

    const signedPath = path.join(temporaryDirectory, 'synthetic-signed.xlsm');
    const signedZip = new JSZip();
    signedZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/custom/security/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>' +
      '<Override PartName="/custom/security/sig-alpha.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>' +
      '</Types>');
    signedZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    signedZip.file('xl/vbaProject.bin', Buffer.from([1]));
    signedZip.file('xl/vbaProjectSignatureAgile.bin', Buffer.from([2]));
    signedZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="custom/security/origin.sigs"/>' +
      '</Relationships>');
    signedZip.file('custom/security/origin.sigs', '');
    signedZip.file('custom/security/_rels/origin.sigs.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig-alpha.xml"/>' +
      '</Relationships>');
    signedZip.file('custom/security/sig-alpha.xml', '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
    await writeFile(signedPath, await signedZip.generateAsync({ type: 'nodebuffer' }));
    const signedEncodedPath = Buffer.from(signedPath, 'utf8').toString('base64');
    const signedExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', signedEncodedPath],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const signedResult = JSON.parse(signedExecution.stdout.trim());
    assert.equal(signedResult.workbook.hasVbaProject, true);
    assert.equal(signedResult.workbook.hasVbaSignature, true);
    assert.equal(signedResult.workbook.vbaSignatureStatus, 'present');
    assert.equal(signedResult.workbook.hasPackageSignature, true);
    assert.equal(signedResult.workbook.packageSignatureStatus, 'present');
    assert.equal(signedResult.workbook.packageSignatureVerificationStatus, 'verified');
    assert.equal(signedResult.workbook.vbaProjectProtectionStatus, 'unknown');

    const badContentTypePath = path.join(temporaryDirectory, 'bad-signature-content-type.xlsx');
    const badContentTypeZip = new JSZip();
    badContentTypeZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/security/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>' +
      '<Override PartName="/security/signature.xml" ContentType="application/xml"/>' +
      '</Types>');
    badContentTypeZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    badContentTypeZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="origin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="security/origin.sigs"/>' +
      '</Relationships>');
    badContentTypeZip.file('security/origin.sigs', '');
    badContentTypeZip.file('security/_rels/origin.sigs.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="signature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="signature.xml"/>' +
      '</Relationships>');
    badContentTypeZip.file('security/signature.xml', '<Signature/>');
    await writeFile(badContentTypePath, await badContentTypeZip.generateAsync({ type: 'nodebuffer' }));
    const badContentTypeExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(badContentTypePath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const badContentTypeResult = JSON.parse(badContentTypeExecution.stdout.trim());
    assert.equal(badContentTypeResult.workbook.hasPackageSignature, false);
    assert.equal(badContentTypeResult.workbook.packageSignatureStatus, 'unknown');
    assert.equal(badContentTypeResult.workbook.packageSignatureVerificationStatus, 'unverifiable');

    const externalPath = path.join(temporaryDirectory, 'external-signature.xlsx');
    const externalZip = new JSZip();
    externalZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '</Types>');
    externalZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    externalZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="externalOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="https://example.invalid/origin.sigs" TargetMode="External"/>' +
      '</Relationships>');
    await writeFile(externalPath, await externalZip.generateAsync({ type: 'nodebuffer' }));
    const externalExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(externalPath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const externalResult = JSON.parse(externalExecution.stdout.trim());
    assert.equal(externalResult.workbook.hasPackageSignature, false);
    assert.equal(externalResult.workbook.packageSignatureStatus, 'unknown');
    assert.equal(externalResult.workbook.packageSignatureVerificationStatus, 'unverifiable');

    const invalidPath = path.join(temporaryDirectory, 'renamed-garbage.xlsx');
    await writeFile(invalidPath, Buffer.from('not an Office package', 'utf8'));
    const invalidEncodedPath = Buffer.from(invalidPath, 'utf8').toString('base64');
    let invalidFailure;
    try {
      await execFileAsync(
        powershell,
        ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', invalidEncodedPath],
        { windowsHide: true, maxBuffer: 264 * 1024 },
      );
    } catch (error) {
      invalidFailure = error;
    }
    assert.ok(invalidFailure, 'a renamed non-Office file must fail closed');
    assert.equal(String(invalidFailure.stderr ?? '').trim(), '', 'failure must not emit non-JSON stderr');
    const invalidLines = String(invalidFailure.stdout ?? '').trim().split(/\r?\n/);
    assert.equal(invalidLines.length, 1, 'failure must emit exactly one JSON object');
    const invalidResult = JSON.parse(invalidLines[0]);
    assert.equal(invalidResult.schemaVersion, 1);
    assert.equal(invalidResult.error?.code, 'inspection_failed');
    assert.match(invalidResult.error?.message ?? '', /container is incompatible/i);

    const invalidZipPath = path.join(temporaryDirectory, 'renamed-zip.xlsx');
    const invalidZip = new JSZip();
    invalidZip.file('unrelated.txt', 'not an Office workbook');
    await writeFile(invalidZipPath, await invalidZip.generateAsync({ type: 'nodebuffer' }));
    const invalidZipEncodedPath = Buffer.from(invalidZipPath, 'utf8').toString('base64');
    let invalidZipFailure;
    try {
      await execFileAsync(
        powershell,
        ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', invalidZipEncodedPath],
        { windowsHide: true, maxBuffer: 264 * 1024 },
      );
    } catch (error) {
      invalidZipFailure = error;
    }
    assert.ok(invalidZipFailure, 'a renamed non-Office ZIP must fail closed');
    const invalidZipResult = JSON.parse(String(invalidZipFailure.stdout ?? '').trim());
    assert.equal(invalidZipResult.error?.code, 'inspection_failed');
    assert.match(invalidZipResult.error?.message ?? '', /container is incompatible/i);
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}

console.log('Enterprise security probe validation passed.');
