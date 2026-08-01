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
const XLSX = require('xlsx');
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
  /\bInvoke-WebRequest\b/i,
  /\bInvoke-RestMethod\b/i,
  /\bHttpClient\b/i,
  /\bWebClient\b/i,
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
  'sensitivityLabels',
  'sensitivityMetadataStatus',
  'sensitivityMetadataSource',
  'VBAWarnings',
  'AccessVBOM',
  'BlockContentExecutionFromInternet',
  'DisableAllActiveX',
  'UFIControls',
  'DisableInternetFilesInPV',
  'DisableUnsafeLocationsInPV',
  'DisableAttachmentsInPV',
  'AllLocationsDisabled',
  'AllowNetworkLocations',
  'Allow User Locations',
  'Software\\Policies\\Microsoft\\Cloud\\Office\\16.0',
  'MaxTrustedLocations = 64',
  'MaxZoneIdentifierBytes = 65536',
  'MaxOutputBytes = 262144',
  'MaxWorkbookBytes = 536870912',
  'zoneStatus',
  'officePackageEncrypted',
  'vbaSignatureStatus',
  'packageSignatureStatus',
  'packageSignatureVerificationStatus',
  'vbaProjectProtectionStatus',
  'irmProtected',
  'cloudPolicyServiceDetected',
  'intuneManagementExtensionDetected',
  'mdmEnrollmentArtifactsDetected',
  'groupPolicyHistoryDetected',
  'unreadableSettings',
  'cloudPolicyDetectionStatus',
  'cloudPolicyServiceStatus',
  'windowsPolicyRegistryStatus',
  'intuneManagementExtensionStatus',
  'mdmEnrollmentStatus',
  'groupPolicyHistoryStatus',
  'mdmProvider',
  'trustedLocationInspectionPartial',
  'registryInspectionPartial',
  'officeArchitecture',
  'EncryptionInfo',
  'EncryptedPackage',
  'Get-Sha256',
  'digital-signature/origin',
  'digital-signature/signature',
  'application/vnd.openxmlformats-package.digital-signature-origin',
  'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml',
  'MaxCharactersInDocument = 1048576',
  'MaxSensitivityLabels = 32',
  'MaxCompoundDirectoryEntries = 16384',
  'MaxCompoundHierarchyDepth = 128',
  'MaxCompoundPathChars = 2048',
  "source = 'officeDefault'",
  'Add-DefaultTrustedLocations',
  'ExpandEnvironmentVariables',
  'http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels',
  'http://schemas.microsoft.com/office/2020/mipLabelMetadata',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
  'application/vnd.openxmlformats-officedocument.custom-properties+xml',
  '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
]) {
  assert.ok(probe.includes(token), `probe is missing expected token: ${token}`);
}

assert.match(probe, /\$streamPath\s*=\s*'\{0\}:Zone\.Identifier'\s+-f\s+\$Path/, 'MOTW must be opened as an alternate data stream');
assert.match(
  probe,
  /BoundedAlternateDataStreamReader[\s\S]+while \(\(read = stream\.Read[\s\S]{0,240}total > maximumBytes/,
  'MOTW reads must enforce a streaming byte cap compatible with Windows PowerShell 5.1',
);
assert.match(probe, /CreateFile\([\s\S]+OpenExisting/, 'MOTW must use the native Windows ADS path rather than an unsupported FileStream path');
assert.doesNotMatch(probe, /Get-Content[\s\S]{0,160}-Stream\s+'Zone\.Identifier'/, 'MOTW must not rely on Get-Content dynamic stream parameters');
assert.doesNotMatch(probe, /HostUrl|ReferrerUrl/i, 'MOTW URLs must never be returned');
assert.match(probe, /DtdProcessing\s*=\s*\[Xml\.DtdProcessing\]::Prohibit/, 'custom XML must prohibit DTD processing');
assert.match(
  probe,
  /\$directoryEntryIndex\s+-ge\s+\$script:MaxCompoundDirectoryEntries[\s\S]{0,160}throw/,
  'CFB directory entry enumeration must fail closed at its explicit cap',
);
assert.match(
  probe,
  /\$childDepth\s+-gt\s+\$script:MaxCompoundHierarchyDepth[\s\S]{0,160}\$childPathLength\s+-gt\s+\$script:MaxCompoundPathChars[\s\S]{0,160}throw/,
  'CFB hierarchy reconstruction must fail closed on depth or path growth',
);
assert.doesNotMatch(probe, /GetValueNames\s*\(/, 'registry reads must not enumerate arbitrary value names');
assert.doesNotMatch(probe, /\.GetSubKeyNames\s*\(/, 'registry subkeys must not be materialized by an unbounded managed API');
assert.match(probe, /RegEnumKeyEx/, 'registry subkeys must use the bounded native enumerator');
assert.match(
  probe,
  /if \(\$null -eq \('ExcelAiVbaStudio\.Security\.BoundedAlternateDataStreamReader' -as \[type\]\)\)[\s\S]{0,100}Add-Type -TypeDefinition/,
  'native helper types must not be redefined when the script is loaded repeatedly in one PowerShell host',
);
assert.match(probe, /names\.Count\s*>=\s*maximumNames[\s\S]{0,180}throw/, 'the native registry enumerator must fail before exceeding its cap');
assert.match(probe, /\[string\[\]\]\$enrollmentNames\s*=\s*@\(\)/, 'MDM enrollment names must always be an array under StrictMode');
assert.match(probe, /\[string\[\]\]\$accountNames\s*=\s*@\(\)/, 'MDM account names must always be an array under StrictMode');
assert.match(probe, /Get-BoundedRegistrySubKeyNames\s+\$enrollmentsKey\s+64/, 'MDM enrollment enumeration must be capped at 64');
assert.match(probe, /Get-BoundedRegistrySubKeyNames\s+\$accountsKey\s+64/, 'MDM account enumeration must be capped at 64');
assert.match(probe, /Get-BoundedRegistrySubKeyNames\s+\$locationsKey\s+256/, 'trusted-location registry enumeration must be capped before filtering');
assert.match(probe, /\^Location\(\?:0\|\[1-9\]\[0-9\]\{0,8\}\)\$/, 'trusted locations must accept only explicitly numbered Location keys');
assert.match(probe, /\$locationNames\.Count\s+-gt\s+\$script:MaxTrustedLocations/, 'trusted-location results must be capped');
assert.match(probe, /\[Environment\]::ExpandEnvironmentVariables\(\$rawLocationPath\)/, 'trusted-location environment variables must be expanded before path matching');
assert.match(probe, /\$locationPath\s+-match\s+'%\[\^%\]\+%'/, 'unresolved trusted-location variables must fail closed');
assert.match(probe, /if \(\$locationReadFailure\)[\s\S]{0,180}\$trustedLocationInspectionPartial\s*=\s*\$true/, 'truncated trusted-location reads must be reported explicitly');
assert.match(probe, /\$inspectionLock[\s\S]+\[IO\.FileShare\]::Read[\s\S]+\$sha256Before/, 'inspection must hold a non-write-sharing workbook handle across all reads');
assert.doesNotMatch(probe, /FileShare\][^\r\n]*Delete|FileShare\]::Delete/, 'workbook handles must not allow deletion during inspection');
assert.doesNotMatch(probe, /HKLM[^\r\n]+Microsoft\\Cloud\\Office\\16\.0/, 'Cloud Policy security detection must not use HKLM Cloud Update keys');
assert.match(probe, /enabledV2Present[\s\S]+enabledV2/, 'Purview detection must version EnabledV2 separately');
assert.match(probe, /\$enabledState\s*=\s*if \(\[bool\]\$item\.enabledV2Present\)/, 'EnabledV2 must take precedence over the legacy Enabled value');
assert.match(probe, /-not \[bool\]\$removed\s+-and\s+\$method\s+-cnotin\s+@\('Standard', 'Privileged'\)/, 'LabelInfo method values must be validated with exact case');
assert.match(probe, /\[bool\]\$removed\s+-and\s+\$method\.Length\s+-ne\s+0/, 'a LabelInfo tombstone must have the empty method required by the specification');
assert.match(probe, /propertyName -ieq 'Sensitivity'/, 'legacy Purview metadata must corroborate the active label with Sensitivity');
assert.match(probe, /GetValueKind\(\$Name\)/, 'registry evidence must retain its native value kind');
assert.ok(
  manifest.contributes.commands.some(({ command }) => command === 'excelAiVbaStudio.openSecurityCenter'),
  'manifest must expose the security center command',
);
assert.match(workbookServiceSource, /async openSecurityCenter[\s\S]+securityCenterPanel\.open/, 'workbook service must open the security panel');
assert.match(panelSource, /Content-Security-Policy[^\r\n]+default-src 'none'/, 'security panel must use a deny-by-default CSP');
assert.match(panelSource, /localResourceRoots:\s*\[\]/, 'security panel must not expose local resource roots');
assert.match(panelSource, /function escapeHtml[\s\S]+replace\(\/\[&<>"'\]\//, 'security panel must escape dynamic HTML');
assert.match(panelSource, /https:\/\/config\.office\.com/, 'security panel must use the official Microsoft 365 Apps admin portal');
assert.match(panelSource, /https:\/\/intune\.microsoft\.com/, 'security panel must expose the official Intune admin portal');
assert.match(panelSource, /https:\/\/purview\.microsoft\.com/, 'security panel must expose the official Purview portal');
assert.match(panelSource, /rôle administrateur autorisé/, 'security panel must explain the enterprise portal authorization boundary');
for (const action of [
  'refresh',
  'copyReport',
  'openExtensionSettings',
  'openEnterpriseAdmin',
  'openIntuneAdmin',
  'openPurviewAdmin',
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
const {
  buildEnterpriseSecurityReport,
  formatEnterpriseSecurityReport,
  parseOfficeSecurityProbe,
} = modelModule.exports;
assert.equal(typeof buildEnterpriseSecurityReport, 'function', 'security model must export its report builder');
assert.equal(typeof formatEnterpriseSecurityReport, 'function', 'security model must export its safe Markdown formatter');
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
      irmProtected: false,
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
      sensitivityLabels: [],
      sensitivityMetadataStatus: 'absent',
      sensitivityMetadataSource: 'none',
      ...(overrides.workbook ?? {}),
    },
    office: {
      version: '16.0',
      architecture: 'x64',
      settings: [],
      unreadableSettings: [],
      trustedLocations: [],
      cloudPolicyDetected: false,
      cloudPolicyServiceDetected: false,
      intuneManagementExtensionDetected: false,
      mdmEnrollmentArtifactsDetected: false,
      groupPolicyHistoryDetected: false,
      cloudPolicyDetectionStatus: 'notDetected',
      cloudPolicyServiceStatus: 'notDetected',
      windowsPolicyRegistryStatus: 'notDetected',
      intuneManagementExtensionStatus: 'notDetected',
      mdmEnrollmentStatus: 'notDetected',
      mdmProvider: 'none',
      groupPolicyHistoryStatus: 'notDetected',
      trustedLocationInspectionPartial: false,
      registryInspectionPartial: false,
      ...(overrides.office ?? {}),
    },
  };
}

const sourceSetting = (source, id, value, managed = source.endsWith('Policy'), extras = {}) => ({
  id,
  category: 'test',
  source,
  managed,
  registryPath: `HKCU\\Synthetic\\${source}`,
  name: id,
  value,
  valueKind: 'DWord',
  ...extras,
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
const cloudMacroDecision = cloudPolicyReport.policyDecisions.find(({ id }) => id === 'vbaWarnings');
assert.equal(cloudMacroDecision.state, 'effective');
assert.equal(cloudMacroDecision.value, 4);
assert.equal(cloudMacroDecision.managed, true);
assert.equal(cloudMacroDecision.shadowedEvidenceCount, 1);
assert.match(cloudMacroDecision.source, /Cloud Policy/);

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

const invalidZoneProbe = parseOfficeSecurityProbe(syntheticProbe({
  workbook: { zoneId: 999, zoneStatus: 'read' },
}));
assert.equal(invalidZoneProbe.workbook.zoneId, null);
assert.equal(invalidZoneProbe.workbook.zoneStatus, 'unreadable');
assert.equal(
  buildEnterpriseSecurityReport(invalidZoneProbe).findings.find(({ id }) => id === 'origin').status,
  'unknown',
  'an out-of-range ZoneId must fail closed',
);
assert.equal(
  parseOfficeSecurityProbe(syntheticProbe({
    workbook: { zoneId: 999, zoneStatus: 'absent' },
  })).workbook.zoneStatus,
  'unreadable',
  'an invalid ZoneId must fail closed even when paired with a contradictory status',
);

const unconfirmedMacroProbe = parseOfficeSecurityProbe(syntheticProbe({
  workbook: {
    hasVbaProject: false,
    hasVbaSignature: false,
    vbaSignatureStatus: 'absent',
    vbaProjectProtectionStatus: 'absent',
  },
}));
const unconfirmedMacroReport = buildEnterpriseSecurityReport(unconfirmedMacroProbe);
for (const findingId of ['macros', 'xlmMacros', 'signatures', 'vbaProtection']) {
  assert.equal(
    unconfirmedMacroReport.findings.find(({ id }) => id === findingId).status,
    'unknown',
    `${findingId} must remain unknown for a macro-capable container without conclusive OPC inventory`,
  );
}
assert.equal(
  unconfirmedMacroReport.capabilities.find(({ id }) => id === 'vbaRead').status,
  'unknown',
  'VBA inspection must remain unknown until the native engine confirms project absence',
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
    architecture: 'unknown',
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

const wrongRegistryViewLocationReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { path: 'C:\\Work\\Department\\book.xlsm' },
  office: {
    architecture: 'x86',
    trustedLocations: [{
      source: 'machinePolicy',
      managed: true,
      registryPath: 'HKLM\\Synthetic\\Location0',
      path: 'C:\\Work',
      allowSubfolders: true,
      registryView: '64',
    }],
  },
}));
assert.equal(wrongRegistryViewLocationReport.workbookInTrustedLocation, false, 'an x64-only location must not apply to x86 Office');

const managedLocationOnlyReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { path: 'C:\\Work\\Department\\book.xlsm' },
  office: {
    architecture: 'x64',
    trustedLocations: [{
      source: 'machinePolicy',
      managed: true,
      registryPath: 'HKLM\\Synthetic\\Location0',
      path: 'C:\\Work',
      allowSubfolders: true,
      registryView: '64',
    }],
  },
}));
assert.equal(managedLocationOnlyReport.workbookInTrustedLocation, true);
assert.equal(managedLocationOnlyReport.level, 'managed', 'a matching managed location must affect the global level');

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

const defaultTrustedLocationProbe = parseOfficeSecurityProbe(syntheticProbe({
  workbook: { path: 'C:\\Users\\Tester\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\book.xlsm' },
  office: {
    settings: [sourceSetting('userPolicy', 'allowUserTrustedLocations', 0)],
    trustedLocations: [{
      source: 'officeDefault',
      managed: false,
      registryPath: 'Excel default trusted location',
      path: 'C:\\Users\\Tester\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART',
      allowSubfolders: false,
      description: 'Excel startup',
    }],
  },
}));
const defaultTrustedLocationReport = buildEnterpriseSecurityReport(defaultTrustedLocationProbe);
assert.equal(
  defaultTrustedLocationReport.workbookInTrustedLocation,
  false,
  'built-in Excel trusted locations are disabled when policy permits only policy-defined locations',
);
assert.equal(
  defaultTrustedLocationReport.findings.find(({ id }) => id === 'trustedLocations').status,
  'blocked',
);

const effectiveDefaultTrustedLocationReport = buildEnterpriseSecurityReport(parseOfficeSecurityProbe(syntheticProbe({
  workbook: { path: 'C:\\Users\\Tester\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\book.xlsm' },
  office: {
    trustedLocations: [{
      source: 'officeDefault',
      managed: false,
      registryPath: 'Excel default trusted location',
      path: 'C:\\Users\\Tester\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART',
      allowSubfolders: false,
    }],
  },
})));
assert.equal(
  effectiveDefaultTrustedLocationReport.workbookInTrustedLocation,
  true,
  'a documented built-in Excel location remains effective when policy permits non-policy locations',
);
assert.match(
  effectiveDefaultTrustedLocationReport.findings.find(({ id }) => id === 'trustedLocations').source,
  /défaut d.Excel/,
);

const partialTrustedLocationReport = buildEnterpriseSecurityReport(parseOfficeSecurityProbe(syntheticProbe({
  office: { trustedLocationInspectionPartial: true },
})));
assert.equal(
  partialTrustedLocationReport.findings.find(({ id }) => id === 'trustedLocations').status,
  'unknown',
  'a truncated trusted-location inventory without a known match must fail closed',
);

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
  workbook: {
    sensitivityLabelIds: ['7e4bc8d6-6897-4f7e-861d-25bf6f908374'],
    sensitivityLabels: [{
      id: '7e4bc8d6-6897-4f7e-861d-25bf6f908374',
      enabled: true,
      removed: false,
      name: 'Confidential',
      method: 'Privileged',
      contentBits: 8,
      siteId: '7d13cf8d-180e-4aa5-8ee3-717fcc85480f',
      source: 'customProperties',
      confidence: 'localDeclaration',
    }],
    sensitivityMetadataStatus: 'present',
    sensitivityMetadataSource: 'customProperties',
  },
}));
assert.notEqual(purviewOnlyReport.level, 'managed', 'file metadata alone must not imply managed Office policy');
assert.equal(purviewOnlyReport.findings.find(({ id }) => id === 'classification').managed, false);
assert.equal(purviewOnlyReport.findings.find(({ id }) => id === 'classification').status, 'warning');
assert.equal(purviewOnlyReport.managementServices.find(({ id }) => id === 'purview').status, 'detected');

const intuneSignalOnlyReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: {
    intuneManagementExtensionDetected: true,
    intuneManagementExtensionStatus: 'detected',
  },
}));
assert.notEqual(intuneSignalOnlyReport.level, 'managed', 'an Intune agent signal must not invent a managed Office rule');
assert.equal(intuneSignalOnlyReport.managementServices.find(({ id }) => id === 'intune').status, 'detected');
assert.match(
  intuneSignalOnlyReport.managementServices.find(({ id }) => id === 'intune').limitation,
  /ne prouve pas/i,
);

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

const x86CloudProbe = parseOfficeSecurityProbe(syntheticProbe({
  office: {
    architecture: 'x86',
    cloudPolicyDetected: true,
    cloudPolicyDetectionStatus: 'detected',
    settings: [sourceSetting('cloudPolicy', 'vbaWarnings', 4, true, { registryView: '64' })],
  },
}));
assert.equal(
  x86CloudProbe.office.settings[0].registryView,
  undefined,
  'HKCU Cloud Policy is a shared registry view and must not inherit a synthetic x64 marker',
);
const x86CloudDecision = buildEnterpriseSecurityReport(x86CloudProbe)
  .policyDecisions.find(({ id }) => id === 'vbaWarnings');
assert.equal(x86CloudDecision.state, 'effective', 'x86 Office must consume the shared HKCU Cloud Policy value');
assert.equal(x86CloudDecision.value, 4);

const unavailableCloudReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: {
    settings: [
      sourceSetting('machinePolicy', 'vbaWarnings', 1),
      sourceSetting('userPreference', 'vbaWarnings', 3, false),
    ],
    unreadableSettings: [{ id: 'vbaWarnings', source: 'cloudPolicy' }],
    cloudPolicyDetectionStatus: 'unknown',
    registryInspectionPartial: true,
  },
}));
const unavailableCloudDecision = unavailableCloudReport.policyDecisions.find(({ id }) => id === 'vbaWarnings');
assert.equal(unavailableCloudDecision.state, 'unknown', 'an unreadable higher-priority Cloud Policy source must fail closed');
assert.equal(unavailableCloudDecision.value, undefined);
assert.equal(unavailableCloudDecision.shadowedEvidenceCount, 2, 'lower-priority values remain evidence, never effective values');

const wrongRegistryTypesReport = buildEnterpriseSecurityReport(syntheticProbe({
  office: {
    settings: [
      sourceSetting('cloudPolicy', 'vbaWarnings', '4', true, { valueKind: 'String' }),
      sourceSetting('cloudPolicy', 'accessVbom', 1, true, { valueKind: 'QWord' }),
    ],
  },
}));
for (const id of ['vbaWarnings', 'accessVbom']) {
  const decision = wrongRegistryTypesReport.policyDecisions.find(candidate => candidate.id === id);
  assert.equal(decision.state, 'unknown', `${id} must reject REG_SZ and REG_QWORD policy values`);
  assert.match(decision.source, /type de registre invalide/i);
}

const dualViewTrustedLocationReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: { path: 'C:\\Work\\Department\\book.xlsm' },
  office: {
    architecture: 'unknown',
    trustedLocations: ['32', '64'].map(registryView => ({
      source: 'machinePolicy',
      managed: true,
      registryPath: 'HKLM\\Software\\Policies\\Microsoft\\Office\\16.0\\Excel\\Security\\Trusted Locations\\Location0',
      path: 'C:\\Work',
      allowSubfolders: true,
      registryView,
    })),
  },
}));
assert.equal(
  dualViewTrustedLocationReport.workbookInTrustedLocation,
  true,
  'identical x86/x64 trusted-location evidence is applicable even when Office architecture is unknown',
);

const markdownInjection = '<script>alert(1)</script>|[lien](javascript:alert(2))\n## injecté';
const injectionReport = buildEnterpriseSecurityReport(syntheticProbe({
  workbook: {
    path: `C:\\Work\\${markdownInjection}.xlsm`,
    name: `${markdownInjection}.xlsm`,
    sensitivityLabelIds: ['7e4bc8d6-6897-4f7e-861d-25bf6f908374'],
    sensitivityLabels: [{
      id: '7e4bc8d6-6897-4f7e-861d-25bf6f908374',
      enabled: true,
      removed: false,
      name: markdownInjection,
      siteId: '7d13cf8d-180e-4aa5-8ee3-717fcc85480f',
      source: 'customProperties',
      confidence: 'localDeclaration',
    }],
    sensitivityMetadataStatus: 'present',
    sensitivityMetadataSource: 'customProperties',
  },
}));
const safeMarkdown = formatEnterpriseSecurityReport(injectionReport);
assert.doesNotMatch(safeMarkdown, /<script>|javascript:|\]\(/i, 'copied Markdown must neutralize HTML and link injection');
assert.doesNotMatch(safeMarkdown, /\n## injecté/, 'embedded newlines must not create attacker-controlled headings');
assert.match(safeMarkdown, /&lt;script&gt;/, 'neutralized data should remain visible in the report');

if (process.platform === 'win32') {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'excel-security-probe-'));
  const workbookPath = path.join(temporaryDirectory, 'synthetic-security.xlsx');
  const labelId = '7e4bc8d6-6897-4f7e-861d-25bf6f908374';
  const disabledLabelId = '82c03a0e-daad-4878-b50a-a1d1a42f16c7';
  const siteId = '7d13cf8d-180e-4aa5-8ee3-717fcc85480f';
  try {
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      '</Types>');
    zip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      '</Relationships>');
    zip.file('xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    zip.file('docProps/custom.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      `<property name="MSIP_Label_${labelId}_Enabled" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2">` +
      '<vt:lpwstr>true</vt:lpwstr>' +
      '</property>' +
      `<property name="MSIP_Label_${labelId}_Name" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3"><vt:lpwstr>Confidential &lt;Finance&gt;</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${labelId}_Method" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4"><vt:lpwstr>Privileged</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${labelId}_SetDate" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5"><vt:lpwstr>2026-08-01T10:11:12Z</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${labelId}_ContentBits" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6"><vt:lpwstr>9</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${labelId}_SiteId" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7"><vt:lpwstr>${siteId}</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${disabledLabelId}_EnabledV2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="8">` +
      '<vt:lpwstr>false</vt:lpwstr>' +
      '</property>' +
      `<property name="Sensitivity" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9"><vt:lpwstr>${labelId}</vt:lpwstr></property>` +
      '</Properties>');
    zip.file('_xmlsignatures/origin.sigs', 'origin-without-a-signature-part');
    const workbookBytes = await zip.generateAsync({ type: 'nodebuffer' });
    await writeFile(workbookPath, workbookBytes);
    const sha256Before = createHash('sha256').update(workbookBytes).digest('hex');
    const neighborsBefore = (await readdir(temporaryDirectory)).sort();

    const encodedPath = Buffer.from(workbookPath, 'utf8').toString('base64');
    const powershell = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
      : 'powershell.exe';
    const nativeBootstrapStart = probe.indexOf('Add-Type -AssemblyName System.IO.Compression.FileSystem');
    const nativeBootstrapEnd = probe.indexOf('$script:MaxOutputBytes');
    assert.ok(nativeBootstrapStart >= 0 && nativeBootstrapEnd > nativeBootstrapStart, 'native helper bootstrap must be extractable');
    const nativeBootstrap = probe.slice(nativeBootstrapStart, nativeBootstrapEnd);
    const repeatedBootstrap = await execFileAsync(
      powershell,
      [
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        `$ErrorActionPreference = 'Stop'\n${nativeBootstrap}\n${nativeBootstrap}\n[Console]::Out.Write('native-bootstrap-ok')`,
      ],
      { windowsHide: true, maxBuffer: 64 * 1024 },
    );
    assert.equal(repeatedBootstrap.stderr.trim(), '', 'native helper bootstrap must not emit diagnostics under Windows PowerShell 5.1');
    assert.equal(repeatedBootstrap.stdout.trim(), 'native-bootstrap-ok', 'native helper bootstrap must be repeat-safe in one PowerShell 5.1 host');
    const inspectWorkbook = async targetPath => {
      const execution = await execFileAsync(
        powershell,
        ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(targetPath, 'utf8').toString('base64')],
        { windowsHide: true, maxBuffer: 264 * 1024 },
      );
      assert.equal(execution.stderr.trim(), '', 'probe must not emit non-JSON diagnostics');
      return JSON.parse(execution.stdout.trim());
    };
    const inspectWorkbookFailure = async targetPath => {
      let failure;
      try {
        await execFileAsync(
          powershell,
          ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(targetPath, 'utf8').toString('base64')],
          { windowsHide: true, maxBuffer: 264 * 1024 },
        );
      } catch (error) {
        failure = error;
      }
      assert.ok(failure, 'the adversarial workbook must fail closed');
      assert.equal(String(failure.stderr ?? '').trim(), '', 'failure must not emit non-JSON stderr');
      const lines = String(failure.stdout ?? '').trim().split(/\r?\n/);
      assert.equal(lines.length, 1, 'failure must emit exactly one JSON object');
      return JSON.parse(lines[0]);
    };
    const writePurviewOpc = async (fileName, { labelInfoXml, customXml }) => {
      const targetPath = path.join(temporaryDirectory, fileName);
      const packageZip = new JSZip();
      packageZip.file('[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        (customXml === undefined
          ? ''
          : '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>') +
        '</Types>');
      packageZip.file('_rels/.rels',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        (labelInfoXml === undefined
          ? ''
          : '<Relationship Id="rLabel" Type="http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels" Target="docMetadata/LabelInfo.xml"/>') +
        (customXml === undefined
          ? ''
          : '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>') +
        '</Relationships>');
      packageZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
      if (labelInfoXml !== undefined) {
        packageZip.file('docMetadata/LabelInfo.xml', labelInfoXml);
      }
      if (customXml !== undefined) {
        packageZip.file('docProps/custom.xml', customXml);
      }
      await writeFile(targetPath, await packageZip.generateAsync({ type: 'nodebuffer' }));
      return targetPath;
    };
    const legacyCustomXml = ({ id, tenantId, enabled = 'true', enabledV2 }) => {
      let pid = 2;
      const property = (name, value) =>
        `<property name="${name}" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid++}"><vt:lpwstr>${value}</vt:lpwstr></property>`;
      return '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        property(`MSIP_Label_${id}_Enabled`, enabled) +
        (enabledV2 === undefined ? '' : property(`MSIP_Label_${id}_EnabledV2`, enabledV2)) +
        property(`MSIP_Label_${id}_SiteId`, tenantId) +
        property('Sensitivity', id) +
        '</Properties>';
    };
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
    assert.equal(result.workbook.sensitivityMetadataStatus, 'present');
    assert.equal(result.workbook.sensitivityMetadataSource, 'customProperties');
    assert.equal(result.workbook.sensitivityLabels.length, 1);
    assert.deepEqual(result.workbook.sensitivityLabels[0], {
      id: labelId,
      enabled: true,
      removed: false,
      name: 'Confidential <Finance>',
      method: 'Privileged',
      setDate: '2026-08-01T10:11:12Z',
      contentBits: 9,
      siteId,
      source: 'customProperties',
      confidence: 'localDeclaration',
    });
    assert.equal(result.workbook.zoneId, null);
    assert.ok(['absent', 'unsupported'].includes(result.workbook.zoneStatus));
    assert.equal(result.workbook.sha256, sha256Before);
	assert.equal(parsedProbe.workbook.sha256, sha256Before, 'the TypeScript boundary must accept the live probe schema');
	assert.equal(parsedProbe.workbook.sensitivityLabels[0].name, 'Confidential <Finance>');
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
    assert.ok(['x86', 'x64', 'unknown'].includes(result.office.architecture));
    assert.ok(Array.isArray(result.office.settings));
    assert.ok(Array.isArray(result.office.unreadableSettings));
    assert.ok(result.office.unreadableSettings.length <= 256);
    assert.ok(Array.isArray(result.office.trustedLocations));
    assert.ok(result.office.trustedLocations.length <= 64);
    assert.equal(typeof result.office.cloudPolicyDetected, 'boolean');
    assert.equal(typeof result.office.cloudPolicyServiceDetected, 'boolean');
    assert.equal(typeof result.office.intuneManagementExtensionDetected, 'boolean');
    assert.equal(typeof result.office.mdmEnrollmentArtifactsDetected, 'boolean');
    assert.equal(typeof result.office.groupPolicyHistoryDetected, 'boolean');
    for (const field of [
      'cloudPolicyDetectionStatus',
      'cloudPolicyServiceStatus',
      'windowsPolicyRegistryStatus',
      'intuneManagementExtensionStatus',
      'mdmEnrollmentStatus',
      'groupPolicyHistoryStatus',
    ]) {
      assert.ok(['detected', 'notDetected', 'unknown'].includes(result.office[field]), `${field} must use the bounded detection status vocabulary`);
    }
    assert.ok(['microsoftIntune', 'unknown', 'none'].includes(result.office.mdmProvider));
    assert.equal(typeof result.office.trustedLocationInspectionPartial, 'boolean');
    assert.equal(typeof result.office.registryInspectionPartial, 'boolean');
    for (const setting of result.office.settings) {
      assert.ok(['machinePolicy', 'userPolicy', 'cloudPolicy', 'userPreference', 'machinePreference'].includes(setting.source));
      assert.equal(typeof setting.managed, 'boolean');
      assert.equal(typeof setting.registryPath, 'string');
      assert.equal(typeof setting.name, 'string');
      assert.equal(typeof setting.valueKind, 'string');
      assert.ok(Object.hasOwn(setting, 'value'));
      assert.ok(
        setting.value === null || ['string', 'number', 'boolean'].includes(typeof setting.value),
        'registry values must remain bounded JSON scalars',
      );
    }

    const modernLabelId = '1422e70b-b185-4d4b-88cc-60f8467cda56';
    const legacyShadowedId = '2933f81c-c296-4e5c-99dd-71f9578deb67';
    const modernLabelPath = path.join(temporaryDirectory, 'modern-label.xlsx');
    const modernLabelZip = new JSZip();
    modernLabelZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      '</Types>');
    modernLabelZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rLabel" Type="http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels" Target="docMetadata/LabelInfo.xml"/>' +
      '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      '</Relationships>');
    modernLabelZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    modernLabelZip.file('docMetadata/LabelInfo.xml',
      '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
      `<label id="${modernLabelId}" enabled="true" method="Standard" siteId="${siteId}" contentBits="3" removed="false"/>` +
      '</labelList>');
    modernLabelZip.file('docProps/custom.xml',
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      `<property name="MSIP_Label_${legacyShadowedId}_Enabled" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2"><vt:lpwstr>true</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${legacyShadowedId}_SiteId" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3"><vt:lpwstr>${siteId}</vt:lpwstr></property>` +
      `<property name="Sensitivity" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4"><vt:lpwstr>${legacyShadowedId}</vt:lpwstr></property>` +
      '</Properties>');
    await writeFile(modernLabelPath, await modernLabelZip.generateAsync({ type: 'nodebuffer' }));
    const modernExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(modernLabelPath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const modernResult = JSON.parse(modernExecution.stdout.trim());
    assert.equal(modernResult.workbook.sensitivityMetadataStatus, 'present');
    assert.equal(modernResult.workbook.sensitivityMetadataSource, 'labelInfoPart');
    assert.deepEqual(modernResult.workbook.sensitivityLabelIds, [modernLabelId], 'LabelInfo must take precedence over legacy custom properties');
    assert.equal(modernResult.workbook.sensitivityLabels[0].method, 'Standard');
    assert.equal(modernResult.workbook.sensitivityLabels[0].contentBits, 3);
    assert.equal(modernResult.workbook.sensitivityLabels[0].siteId, siteId);
    assert.equal(modernResult.workbook.sensitivityLabels[0].source, 'labelInfoPart');

    const secondTenantId = '28b97407-b95d-4f79-8af7-9d36f4769536';
    const mixedLegacyId = 'd2bd77af-6e65-4cec-9a61-156490752d1b';
    const mixedPath = await writePurviewOpc('mixed-tenants.xlsx', {
      labelInfoXml:
        '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
        `<label id="${modernLabelId}" enabled="true" method="Standard" siteId="${siteId}" contentBits="1" removed="false"/>` +
        '</labelList>',
      customXml: legacyCustomXml({ id: mixedLegacyId, tenantId: secondTenantId }),
    });
    const mixedResult = await inspectWorkbook(mixedPath);
    assert.equal(mixedResult.workbook.sensitivityMetadataStatus, 'present');
    assert.equal(mixedResult.workbook.sensitivityMetadataSource, 'mixed');
    assert.deepEqual(
      new Set(mixedResult.workbook.sensitivityLabelIds),
      new Set([modernLabelId, mixedLegacyId]),
      'LabelInfo must shadow legacy metadata only for the tenant represented by LabelInfo',
    );
    assert.deepEqual(
      new Set(mixedResult.workbook.sensitivityLabels.map(label => label.source)),
      new Set(['labelInfoPart', 'customProperties']),
    );

    const emptyLabelInfoFallbackPath = await writePurviewOpc('empty-labelinfo-fallback.xlsx', {
      labelInfoXml: '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata"/>',
      customXml: legacyCustomXml({ id: mixedLegacyId, tenantId: secondTenantId }),
    });
    const emptyLabelInfoFallbackResult = await inspectWorkbook(emptyLabelInfoFallbackPath);
    assert.equal(emptyLabelInfoFallbackResult.workbook.sensitivityMetadataStatus, 'present');
    assert.equal(emptyLabelInfoFallbackResult.workbook.sensitivityMetadataSource, 'customProperties');
    assert.deepEqual(emptyLabelInfoFallbackResult.workbook.sensitivityLabelIds, [mixedLegacyId]);

    const tombstoneId = '28fd4aa2-b799-4019-828b-2bf88c281631';
    const tombstonePath = await writePurviewOpc('labelinfo-tombstone.xlsx', {
      labelInfoXml:
        '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
        `<label id="${tombstoneId}" enabled="false" method="" siteId="${siteId}" removed="true"/>` +
        '</labelList>',
      customXml: legacyCustomXml({ id: legacyShadowedId, tenantId: siteId }),
    });
    const tombstoneResult = await inspectWorkbook(tombstonePath);
    assert.equal(tombstoneResult.workbook.sensitivityMetadataStatus, 'absent');
    assert.equal(tombstoneResult.workbook.sensitivityMetadataSource, 'labelInfoPart');
    assert.deepEqual(
      tombstoneResult.workbook.sensitivityLabelIds,
      [],
      'a LabelInfo tombstone must suppress stale legacy metadata for the same tenant',
    );

    const wrongCaseMethodPath = await writePurviewOpc('labelinfo-wrong-method-case.xlsx', {
      labelInfoXml:
        '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
        `<label id="${modernLabelId}" enabled="true" method="standard" siteId="${siteId}" removed="false"/>` +
        '</labelList>',
      customXml: undefined,
    });
    const wrongCaseMethodResult = await inspectWorkbook(wrongCaseMethodPath);
    assert.equal(
      wrongCaseMethodResult.workbook.sensitivityMetadataStatus,
      'unknown',
      'LabelInfo method values are case-sensitive and lowercase standard must fail closed',
    );
    assert.equal(wrongCaseMethodResult.workbook.sensitivityMetadataSource, 'ambiguous');

    const contradictoryTombstonePath = await writePurviewOpc('labelinfo-enabled-tombstone.xlsx', {
      labelInfoXml:
        '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
        `<label id="${tombstoneId}" enabled="true" method="" siteId="${siteId}" removed="true"/>` +
        '</labelList>',
      customXml: legacyCustomXml({ id: legacyShadowedId, tenantId: siteId }),
    });
    const contradictoryTombstoneResult = await inspectWorkbook(contradictoryTombstonePath);
    assert.equal(
      contradictoryTombstoneResult.workbook.sensitivityMetadataStatus,
      'absent',
      'removed=true remains a tombstone even when enabled=true violates only the specification recommendation',
    );
    assert.equal(contradictoryTombstoneResult.workbook.sensitivityMetadataSource, 'labelInfoPart');
    assert.deepEqual(
      contradictoryTombstoneResult.workbook.sensitivityLabelIds,
      [],
      'removed=true must take precedence and suppress stale legacy metadata for that tenant',
    );

    const enabledV2Id = 'd8f11135-857d-4128-aafe-79115d26be2d';
    const enabledV2TruePath = await writePurviewOpc('enabled-v2-true.xlsx', {
      labelInfoXml: undefined,
      customXml: legacyCustomXml({
        id: enabledV2Id,
        tenantId: secondTenantId,
        enabled: 'false',
        enabledV2: 'true',
      }),
    });
    const enabledV2TrueResult = await inspectWorkbook(enabledV2TruePath);
    assert.equal(enabledV2TrueResult.workbook.sensitivityMetadataStatus, 'present');
    assert.deepEqual(enabledV2TrueResult.workbook.sensitivityLabelIds, [enabledV2Id]);

    const adversarialMetadataPath = await writePurviewOpc('adversarial-metadata.xlsx', {
      labelInfoXml: undefined,
      customXml:
        '<!DOCTYPE Properties [<!ENTITY xxe SYSTEM "file:///C:/Windows/win.ini">]>' +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        `<property name="MSIP_Label_${labelId}_Enabled" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2"><vt:lpwstr>&xxe;</vt:lpwstr></property>` +
        '</Properties>',
    });
    const adversarialMetadataResult = await inspectWorkbook(adversarialMetadataPath);
    assert.equal(adversarialMetadataResult.workbook.sensitivityMetadataStatus, 'unknown');
    assert.equal(adversarialMetadataResult.workbook.sensitivityMetadataSource, 'ambiguous');
    assert.equal(adversarialMetadataResult.workbook.name, path.basename(adversarialMetadataPath));

    const irmCompoundPath = path.join(temporaryDirectory, 'irm-labelinfo.xls');
    const irmCompound = XLSX.CFB.utils.cfb_new();
    for (const [streamName, bytes] of [
      ['Workbook', [0x09, 0x08]],
      ['\u0006DataSpaces/DataSpaceMap', [1]],
      ['\u0006DataSpaces/TransformInfo/EUL-test', [2]],
      ['\u0006DataSpaces/TransformInfo/LabelInfo', [3]],
      ['EncryptedPackage', [4]],
    ]) {
      XLSX.CFB.utils.cfb_add(irmCompound, streamName, Buffer.from(bytes));
    }
    const irmCompoundBytes = XLSX.CFB.write(irmCompound, { type: 'buffer' });
    await writeFile(irmCompoundPath, irmCompoundBytes);
    const irmCompoundResult = await inspectWorkbook(irmCompoundPath);
    assert.equal(irmCompoundResult.workbook.containerKind, 'compound');
    assert.equal(irmCompoundResult.workbook.officePackageEncrypted, true);
    assert.equal(irmCompoundResult.workbook.irmProtected, true);
    assert.equal(irmCompoundResult.workbook.sensitivityMetadataStatus, 'unknown');
    assert.equal(irmCompoundResult.workbook.sensitivityMetadataSource, 'labelInfoStream');
    assert.deepEqual(irmCompoundResult.workbook.sensitivityLabels, []);
    assert.equal(
      irmCompoundResult.workbook.sha256,
      createHash('sha256').update(irmCompoundBytes).digest('hex'),
      'CFB/IRM detection must remain read-only',
    );

    const deepCompoundPath = path.join(temporaryDirectory, 'deep-hierarchy.xls');
    const deepCompound = XLSX.CFB.utils.cfb_new();
    XLSX.CFB.utils.cfb_add(deepCompound, 'Workbook', Buffer.from([0x09, 0x08]));
    const deepStoragePath = Array.from(
      { length: 130 },
      (_, index) => `storage${index}`,
    ).join('/');
    XLSX.CFB.utils.cfb_add(
      deepCompound,
      `${deepStoragePath}/payload`,
      Buffer.from([1]),
    );
    await writeFile(
      deepCompoundPath,
      XLSX.CFB.write(deepCompound, { type: 'buffer' }),
    );
    const deepCompoundResult = await inspectWorkbookFailure(deepCompoundPath);
    assert.equal(deepCompoundResult.schemaVersion, 1);
    assert.equal(deepCompoundResult.error?.code, 'inspection_failed');
    assert.match(
      deepCompoundResult.error?.message ?? '',
      /directory hierarchy exceeds the inspection limit/i,
      'an over-deep CFB hierarchy must be rejected before unbounded path construction',
    );

    const orphanLabelPath = path.join(temporaryDirectory, 'orphan-label.xlsx');
    const orphanLabelZip = new JSZip();
    orphanLabelZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '</Types>');
    orphanLabelZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    orphanLabelZip.file('docMetadata/LabelInfo.xml',
      '<labelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata"/>');
    await writeFile(orphanLabelPath, await orphanLabelZip.generateAsync({ type: 'nodebuffer' }));
    const orphanExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(orphanLabelPath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const orphanResult = JSON.parse(orphanExecution.stdout.trim());
    assert.equal(orphanResult.workbook.sensitivityMetadataStatus, 'unknown');
    assert.equal(orphanResult.workbook.sensitivityMetadataSource, 'ambiguous');
    assert.deepEqual(orphanResult.workbook.sensitivityLabels, []);

    const v2DisabledPath = path.join(temporaryDirectory, 'v2-disabled.xlsx');
    const v2DisabledZip = new JSZip();
    v2DisabledZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      '</Types>');
    v2DisabledZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      '</Relationships>');
    v2DisabledZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    v2DisabledZip.file('docProps/custom.xml',
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      `<property name="MSIP_Label_${labelId}_Enabled" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2"><vt:lpwstr>true</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${labelId}_EnabledV2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3"><vt:lpwstr>false</vt:lpwstr></property>` +
      `<property name="Sensitivity" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4"><vt:lpwstr>${labelId}</vt:lpwstr></property>` +
      '</Properties>');
    await writeFile(v2DisabledPath, await v2DisabledZip.generateAsync({ type: 'nodebuffer' }));
    const v2DisabledExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(v2DisabledPath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const v2DisabledResult = JSON.parse(v2DisabledExecution.stdout.trim());
    assert.equal(v2DisabledResult.workbook.sensitivityMetadataStatus, 'unknown', 'EnabledV2=false must override legacy Enabled=true');
    assert.deepEqual(v2DisabledResult.workbook.sensitivityLabelIds, []);

    const malformedMetadataPath = path.join(temporaryDirectory, 'malformed-metadata.xlsx');
    const malformedMetadataZip = new JSZip();
    malformedMetadataZip.file('[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      '</Types>');
    malformedMetadataZip.file('_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      '</Relationships>');
    malformedMetadataZip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    malformedMetadataZip.file('docProps/custom.xml',
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property');
    await writeFile(malformedMetadataPath, await malformedMetadataZip.generateAsync({ type: 'nodebuffer' }));
    const malformedMetadataExecution = await execFileAsync(
      powershell,
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', probePath, '-WorkbookPathBase64', Buffer.from(malformedMetadataPath, 'utf8').toString('base64')],
      { windowsHide: true, maxBuffer: 264 * 1024 },
    );
    const malformedMetadataResult = JSON.parse(malformedMetadataExecution.stdout.trim());
    assert.equal(malformedMetadataResult.workbook.sensitivityMetadataStatus, 'unknown', 'malformed Purview metadata must degrade only that signal');
    assert.equal(malformedMetadataResult.workbook.name, path.basename(malformedMetadataPath), 'the rest of the workbook report must remain available');

    const neighborsBeforeZone = (await readdir(temporaryDirectory)).sort();
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
    assert.deepEqual((await readdir(temporaryDirectory)).sort(), neighborsBeforeZone, 'an ADS must not be treated as a neighboring file');

    await writeFile(
      `${workbookPath}:Zone.Identifier`,
      '[Unrelated]\r\nZoneId=0\r\n[ZoneTransfer]\r\nZoneId=3\r\n',
    );
    const sectionAwareZoneResult = await inspectWorkbook(workbookPath);
    assert.equal(sectionAwareZoneResult.workbook.zoneStatus, 'read');
    assert.equal(sectionAwareZoneResult.workbook.zoneId, 3, 'only ZoneId in the ZoneTransfer section may be trusted');

    await writeFile(`${workbookPath}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=999\r\n');
    const invalidZoneResult = await inspectWorkbook(workbookPath);
    assert.equal(invalidZoneResult.workbook.zoneStatus, 'unreadable');
    assert.equal(invalidZoneResult.workbook.zoneId, null, 'out-of-range MOTW values must fail closed');

    await writeFile(
      `${workbookPath}:Zone.Identifier`,
      '[ZoneTransfer]\r\nZoneId=3\r\n[ZoneTransfer]\r\nZoneId=3\r\n',
    );
    const duplicateZoneResult = await inspectWorkbook(workbookPath);
    assert.equal(duplicateZoneResult.workbook.zoneStatus, 'unreadable');
    assert.equal(duplicateZoneResult.workbook.zoneId, null, 'ambiguous ZoneTransfer sections must fail closed');

    await writeFile(`${workbookPath}:Zone.Identifier`, Buffer.alloc(65537, 0x41));
    const oversizedZoneResult = await inspectWorkbook(workbookPath);
    assert.equal(oversizedZoneResult.workbook.zoneStatus, 'unreadable');
    assert.equal(oversizedZoneResult.workbook.zoneId, null, 'MOTW larger than 64 KiB must be rejected by the bounded native reader');

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
    assert.equal(signedResult.workbook.packageSignatureVerificationStatus, 'structureVerified');
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
