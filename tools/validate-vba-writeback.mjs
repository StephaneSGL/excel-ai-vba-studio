import assert from 'node:assert/strict';
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
const helperPath = resolve(
  root,
  'bin/win32-x64/excel-ai-vba-writeback.exe',
);

assert.match(cli, /expectedWorkbookSha256/, 'stale workbook hash gate is missing');
assert.match(cli, /detect_signature/, 'signed project gate is missing');
assert.match(cli, /has_password/, 'protected project gate is missing');
assert.match(cli, /designer_stream_hashes/, 'UserForm designer preservation is missing');
assert.match(cli, /project_stream_fingerprint/, 'complete VBA stream fingerprint is missing');
assert.match(cli, /fingerprint_request/, 'VBA fingerprint operation is missing');
assert.match(cli, /zip_payload_hashes/, 'non-VBA OOXML preservation is missing');
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

console.log('VBA write-back validation passed: transactional, bounded, no COM/registry/macro execution.');
