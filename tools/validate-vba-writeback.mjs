import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const cli = readFileSync(resolve(root, 'native/vba-writeback/cli.py'), 'utf8');
const service = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/vbaWritebackService.ts'),
  'utf8',
);
const languageTool = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/languageModelTool.ts'),
  'utf8',
);
const buildScript = readFileSync(resolve(root, 'tools/build-vba-writeback.ps1'), 'utf8');
const buildRequirements = readFileSync(
  resolve(root, 'native/vba-writeback/requirements.lock'),
  'utf8',
);
const helperPath = resolve(
  root,
  'bin/win32-x64/excel-ai-vba-writeback.exe',
);
const provenancePath = resolve(
  root,
  'bin/win32-x64/excel-ai-vba-writeback.provenance.json',
);

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function canonicalTextSha256(path) {
  const canonicalText = readFileSync(path, 'utf8').replace(/\r\n?/g, '\n');
  return createHash('sha256').update(canonicalText, 'utf8').digest('hex');
}

assert.match(cli, /expectedWorkbookSha256/, 'stale workbook hash gate is missing');
assert.match(cli, /detect_signature/, 'signed project gate is missing');
assert.match(cli, /has_password/, 'protected project gate is missing');
assert.match(cli, /designer_stream_hashes/, 'UserForm designer preservation is missing');
assert.match(cli, /project_stream_fingerprint/, 'complete VBA stream fingerprint is missing');
assert.match(cli, /fingerprint_request/, 'VBA fingerprint operation is missing');
assert.match(cli, /zip_payload_hashes/, 'non-VBA OOXML preservation is missing');
assert.match(cli, /MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES/, 'per-entry ZIP inflation limit is missing');
assert.match(cli, /MAX_ZIP_TOTAL_UNCOMPRESSED_BYTES/, 'aggregate ZIP inflation limit is missing');
assert.match(cli, /_create_private_workbook_snapshot/, 'private bounded workbook snapshots are missing');
assert.match(cli, /_preflight_zip_payloads\(snapshot_path\)/, 'native ZIP preflight is missing');
assert.doesNotMatch(
  cli,
  /archive[.]read\(/,
  'native ZIP payloads must be streamed instead of eagerly inflated',
);
assert.match(cli, /assert_no_reparse_point_chain/, 'native reparse-point gate is missing');
assert.match(cli, /project\.add_module/, 'new standard/class module support is missing');
assert.match(cli, /replace_file_with_backup/, 'atomic workbook replacement is missing');
assert.match(cli, /ReplaceFileW/, 'Windows displaced-file capture is missing');
assert.match(cli, /restore_displaced_workbook/, 'conflict-safe restoration is missing');
assert.match(cli, /restore_missing_workbook/, 'partial replace recovery is missing');
assert.match(cli, /handle_failed_atomic_replace/, 'partial ReplaceFile failure handling is missing');
assert.match(cli, /MoveFileExW/, 'non-overwriting missing-path recovery is missing');
assert.doesNotMatch(
  cli,
  /import\s+winreg|Excel\.Application|Dispatch\(|RunAutoMacros|\.Run\(/i,
  'write-back helper must not touch Office automation, registry, or macros',
);
assert.match(buildScript, /--require-hashes/, 'native helper dependencies must be installed with hash verification');
assert.equal(
  [...buildRequirements.matchAll(/--hash=sha256:[a-f0-9]{64}/g)].length,
  7,
  'every pinned native helper build dependency must include its wheel SHA-256',
);
assert.doesNotMatch(
  buildRequirements,
  /^[A-Za-z0-9_.-]+==[^\r\n]+(?<!\\)$/gm,
  'every native helper requirement must continue to an explicit hash',
);
assert.match(
  service,
  /expectedDesignerSha256/,
  'extension write-back contract is incomplete',
);
assert.match(service, /ensureOwnedDirectory/, 'owned request directory gate is missing');
assert.match(
  languageTool,
  /EXCEL_AI_VBA_WRITE_TOOL[\s\S]+?writeVbaFromTool/,
  'Copilot VBA write tool is not connected to the shared service',
);
assert.ok(
  statSync(helperPath).size > 1_000_000,
  'bundled Windows x64 helper is missing or implausibly small',
);
const provenance = JSON.parse(readFileSync(provenancePath, 'utf8'));
assert.deepEqual(
  Object.keys(provenance).sort(),
  ['executableSha256', 'pythonVersion', 'schemaVersion', 'sourceSha256'].sort(),
  'native helper provenance has unexpected fields',
);
assert.equal(provenance.schemaVersion, 1, 'native helper provenance schema is unsupported');
assert.equal(provenance.pythonVersion, '3.11.9', 'native helper Python version is inconsistent');
assert.equal(
  provenance.sourceSha256,
  canonicalTextSha256(resolve(root, 'native/vba-writeback/cli.py')),
  'bundled helper was not built from the current native source',
);
assert.equal(
  provenance.executableSha256,
  sha256(helperPath),
  'bundled helper bytes do not match their build provenance',
);

console.log('VBA write-back validation passed: transactional, bounded, no COM/registry/macro execution.');
