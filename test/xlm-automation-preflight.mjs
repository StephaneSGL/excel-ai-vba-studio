import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { promisify } from 'node:util';
import { build } from 'esbuild';
import JSZip from 'jszip';

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);
const root = resolve(import.meta.dirname, '..');
const helperPath = join(root, 'src', 'common', 'ooxmlPackageSignature.ts');
const sharedPowerShellPath = join(root, 'scripts', 'ooxml-package-signature.ps1');
const nativeEditScriptPath = join(root, 'scripts', 'office-ai-apply-edits.ps1');
const powerShell = process.platform === 'win32'
  ? join(
      process.env.SystemRoot || 'C:\\Windows',
      'System32',
      'WindowsPowerShell',
      'v1.0',
      'powershell.exe',
    )
  : 'pwsh';

const contentTypesNamespace =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const relationshipsNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const officeDocumentRelationship =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const xlmRelationship =
  'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet';
const xlmContentType = 'application/vnd.ms-excel.macrosheet+xml';

async function loadHelper() {
  const result = await build({
    entryPoints: [helperPath],
    bundle: true,
    format: 'cjs',
    platform: 'node',
    target: 'node20',
    write: false,
    logLevel: 'silent',
  });
  const module = { exports: {} };
  new Function(
    'module',
    'exports',
    'require',
    '__filename',
    '__dirname',
    result.outputFiles[0].text,
  )(
    module,
    module.exports,
    require,
    helperPath,
    join(root, 'src', 'common'),
  );
  return module.exports;
}

async function workbookBytes({ xlm = false, pathOnly = false } = {}) {
  const zip = new JSZip();
  const macroPart = pathOnly ? '/xl/macrosheets/sheet1.xml' : '/custom/macro.xml';
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="${contentTypesNamespace}">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + (xlm && !pathOnly
        ? `<Override PartName="${macroPart}" ContentType="${xlmContentType}"/>`
        : '')
      + '</Types>',
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="${relationshipsNamespace}">`
      + `<Relationship Id="rIdWorkbook" Type="${officeDocumentRelationship}" Target="xl/workbook.xml"/>`
      + '</Relationships>',
  );
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
  );
  if (xlm) {
    zip.file(macroPart.slice(1), '<macrosheet xmlns="http://schemas.microsoft.com/office/excel/2006/main"/>');
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="${relationshipsNamespace}">`
        + `<Relationship Id="rIdXlm" Type="${xlmRelationship}" Target="${pathOnly ? 'macrosheets/sheet1.xml' : '../custom/macro.xml'}"/>`
        + '</Relationships>',
    );
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function encoded(value) {
  return Buffer.from(value, 'utf8').toString('base64');
}

async function runPowerShellPreflight(filePath) {
  const command =
    "Add-Type -AssemblyName System.IO.Compression.FileSystem; "
    + `. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${encoded(sharedPowerShellPath)}'))); `
    + `Assert-OoxmlPackageHasNoXlmMacroSheets ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${encoded(filePath)}'))); `
    + "[Console]::Out.Write('SAFE')";
  return execFileAsync(powerShell, [
    '-NoLogo',
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-Command',
    command,
  ], { cwd: root, windowsHide: true, timeout: 30_000, maxBuffer: 1024 * 1024 });
}

const helper = await loadHelper();
assert.equal(typeof helper.hasOoxmlXlmMacroSheetsBytes, 'function');
assert.equal(typeof helper.assertOoxmlHasNoXlmMacroSheetsForAutomation, 'function');

const testRoot = await mkdtemp(join(tmpdir(), 'excel-xlm-preflight-'));
try {
  const safeBytes = await workbookBytes();
  const xlmBytes = await workbookBytes({ xlm: true });
  const pathOnlyBytes = await workbookBytes({ xlm: true, pathOnly: true });
  assert.equal(await helper.hasOoxmlXlmMacroSheetsBytes(safeBytes), false);
  assert.equal(await helper.hasOoxmlXlmMacroSheetsBytes(xlmBytes), true);
  assert.equal(await helper.hasOoxmlXlmMacroSheetsBytes(pathOnlyBytes), true);

  const safePath = join(testRoot, 'safe.xlsx');
  const xlmPath = join(testRoot, 'xlm.xlsm');
  const operationsPath = join(testRoot, 'operations.json');
  await writeFile(safePath, safeBytes);
  await writeFile(xlmPath, xlmBytes);
  await writeFile(operationsPath, '{}', 'utf8');
  await helper.assertOoxmlHasNoXlmMacroSheetsForAutomation(safePath);
  await assert.rejects(
    helper.assertOoxmlHasNoXlmMacroSheetsForAutomation(xlmPath),
    /feuille macro Excel 4\.0 \(XLM\)/,
  );

  const safePowerShell = await runPowerShellPreflight(safePath);
  assert.equal(String(safePowerShell.stdout), 'SAFE');
  await assert.rejects(
    runPowerShellPreflight(xlmPath),
    error => /Excel 4\.0 macro sheet detected/.test(
      String(error.stdout ?? '') + String(error.stderr ?? ''),
    ),
  );

  await assert.rejects(
    execFileAsync(powerShell, [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      nativeEditScriptPath,
      '-WorkbookPath',
      xlmPath,
      '-OperationsPath',
      operationsPath,
    ], { cwd: root, windowsHide: true, timeout: 30_000, maxBuffer: 1024 * 1024 }),
    error => {
      const output = String(error.stdout ?? '') + String(error.stderr ?? '');
      return /Excel 4\.0 macro sheet detected/.test(output)
        && !/OWNED_EXCEL_PID\|/.test(output);
    },
  );

  console.log('XLM automation preflight passed: safe OOXML accepted, macro sheets refused before Excel COM.');
} finally {
  await rm(testRoot, { recursive: true, force: true });
}
