import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdtemp, readFile, writeFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve, dirname } from 'node:path';
import { build } from 'esbuild';
import JSZip from 'jszip';

// All workbook bytes below are generated here; no user workbook is opened.
const root = resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);
const temporary = await mkdtemp(join(tmpdir(), 'portable-editor-'));
const state = { destination: undefined, commands: [], errors: [] };
const uri = path => ({ scheme: 'file', fsPath: path, path, toString: () => `file://${path}` });
globalThis.__portableTest = { state, uri, readFile, writeFile, stat };

const vscodeStub = `
const test = globalThis.__portableTest;
export const Uri = { file: test.uri, joinPath: () => { throw new Error('Unexpected virtual path'); } };
export const FilePermission = { Readonly: 1 };
export const workspace = { fs: {
  readFile: uri => test.readFile(uri.fsPath),
  writeFile: (uri, bytes) => test.writeFile(uri.fsPath, bytes),
  stat: async uri => ({ size: (await test.stat(uri.fsPath)).size, permissions: 0 }),
} };
export const commands = { executeCommand: async (...args) => test.state.commands.push(args) };
export const window = {
  showSaveDialog: async () => test.state.destination,
  showWarningMessage: async message => test.state.errors.push(message),
  showErrorMessage: async message => test.state.errors.push(message),
};
export const env = { openExternal: async () => {} };
export const ViewColumn = { Active: 1, Beside: 2 };
`;

async function bundle(relative, contents) {
  const filename = join(root, relative || 'portable-test-entry.ts');
  const result = await build({
    ...(contents ? { stdin: { contents, resolveDir: root, loader: 'ts' } } : { entryPoints: [filename] }),
    bundle: true, write: false, format: 'cjs', platform: 'node',
    target: 'node20', logLevel: 'silent',
    plugins: [{ name: 'portable-test-stubs', setup(api) {
      api.onResolve({ filter: /^vscode$/ }, () => ({ path: 'vscode', namespace: 'test' }));
      api.onResolve({ filter: /util\/vscode$/ }, () => ({ path: 'webview', namespace: 'test' }));
      api.onLoad({ filter: /.*/, namespace: 'test' }, args => ({
        contents: args.path === 'vscode' ? vscodeStub : 'export const handler = { emit() {} };',
      }));
    } }],
  });
  const module = { exports: {} };
  new Function('module', 'exports', 'require', '__filename', '__dirname', result.outputFiles[0].text)(
    module, module.exports, require, filename, dirname(filename),
  );
  return module.exports;
}

function createHandler() {
  const callbacks = new Map();
  return {
    panel: { title: '' }, events: [],
    on(name, fn) { callbacks.set(name, fn); return this; },
    emit(name, payload) { this.events.push({ name, payload }); },
    call(name, payload) { return callbacks.get(name)(payload); },
  };
}

try {
  const { ExcelJS } = await bundle(undefined, "export {default as ExcelJS} from '@cweijan/exceljs';");
  const { getPlatformCapabilities } = await bundle('src/common/platformCapabilities.ts');
  assert.equal(getPlatformCapabilities('linux', 'x64').nativeExcel, false);
  assert.equal(getPlatformCapabilities('linux', 'arm64').portableSpreadsheet, true);
  assert.equal(getPlatformCapabilities('win32', 'x64').nativeExcel, true);
  assert.equal(getPlatformCapabilities('win32', 'arm64').nativeExcel, false);

  const { createBlankWorkbookBytes } = await bundle('src/provider/newWorkbook.ts');
  const { getPortableXlsxPreservationIssue: issue } = await bundle('src/common/ooxmlPackageSignature.ts');
  const blank = await createBlankWorkbookBytes();
  assert.equal(await issue(blank), undefined, 'new workbook must be editable');
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Synthetic');
  sheet.getCell('A1').value = 7;
  sheet.getCell('A2').value = true;
  sheet.getCell('A3').value = false;
  sheet.getCell('A4').value = '42';
  sheet.getCell('A5').value = '=not-a-formula';
  sheet.getCell('A6').value = new Date('2024-02-29T15:30:00.000Z');
  sheet.getCell('A6').numFmt = 'yyyy-mm-dd hh:mm';
  sheet.getCell('A7').value = { error: '#N/A' };
  sheet.getCell('B1').value = { formula: 'A1*2', result: 14 };
  sheet.getCell('B2').value = 'editable';
  sheet.getCell('B2').font = { bold: true };
  const simple = await workbook.xlsx.writeBuffer();
  assert.equal(await issue(simple), undefined, 'simple ExcelJS workbook must be editable');

  const reader = await bundle('src/react/view/excel/excel_reader.ts');
  const writer = await bundle('src/react/view/excel/excel_writer.ts');
  const loaded = await reader.readExcel(Uint8Array.from(simple).buffer);
  loaded.sheets[0].rows[1].cells[1].text = 'edited';
  const roundtrip = await writer.buildExcelWorkbookBuffer(loaded.sheets);
  assert.equal(await issue(roundtrip), undefined, 'portable writer output must remain editable');
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(roundtrip);
  for (const address of ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'B1']) {
    assert.deepEqual(reloaded.worksheets[0].getCell(address).value, sheet.getCell(address).value, `${address} preserves its scalar/formula type`);
  }
  assert.equal(reloaded.worksheets[0].getCell('B2').value, 'edited');
  assert.equal(reloaded.worksheets[0].getCell('B2').font.bold, true);

  const richBook = new ExcelJS.Workbook();
  richBook.addWorksheet('Rich').getCell('A1').value = {
    richText: [{ font: { bold: true }, text: '=1+1' }],
  };
  const richBytes = await richBook.xlsx.writeBuffer();
  assert.match(await issue(richBytes), /preservation.*r/i, 'rich text must not be silently flattened');
  const richLoaded = await reader.readExcel(Uint8Array.from(richBytes).buffer);
  const richExport = new ExcelJS.Workbook();
  await richExport.xlsx.load(await writer.buildExcelWorkbookBuffer(richLoaded.sheets));
  assert.equal(richExport.worksheets[0].getCell('A1').value, '=1+1', 'standalone export must never promote a literal label into a formula');
  const printedBook = new ExcelJS.Workbook();
  const printedSheet = printedBook.addWorksheet('Printed');
  printedSheet.getCell('A1').value = 'Synthetic';
  printedSheet.headerFooter = { oddHeader: 'Synthetic header', oddFooter: 'Synthetic footer' };
  const printedBytes = await printedBook.xlsx.writeBuffer();
  assert.match(await issue(printedBytes), /headerFooter/, 'print content requires preservation');

  const mutate = async (path, value) => {
    const zip = await JSZip.loadAsync(simple);
    zip.file(path, typeof value === 'function' ? value(await zip.file(path).async('string')) : value);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };
  const opaque = await mutate('xl/charts/chart1.xml', '<chart/>');
  for (const [label, bytes] of [
    ['chart', opaque],
    ['custom XML', await mutate('customXml/item1.xml', '<data/>')],
    ['orphaned VBA', await mutate('xl/vbaProject.bin', 'opaque')],
    ['date1904', await mutate('xl/workbook.xml', xml => xml.replace('<workbookPr', '<workbookPr date1904="1"'))],
    ['hidden sheet', await mutate('xl/workbook.xml', xml => xml.replace('state="visible"', 'state="hidden"'))],
    ['defined name', await mutate('xl/workbook.xml', xml => xml.replace('</workbook>', '<definedNames><definedName name="Secret">Synthetic!A1</definedName></definedNames></workbook>'))],
    ['unsupported extension', await mutate('xl/worksheets/sheet1.xml', xml => xml.replace('</worksheet>', '<extLst><ext uri="future"/></extLst></worksheet>'))],
  ]) assert.ok(await issue(bytes), `${label} requires preservation mode`);

  if (process.platform !== 'win32') {
    const { handleCommonEvent } = await bundle('src/provider/compress/commonHandler.ts');
    const open = async (filename, bytes) => {
      const file = join(temporary, filename);
      await writeFile(file, bytes);
      const handler = createHandler();
      handleCommonEvent(uri(file), handler);
      await handler.call('init');
      return { handler, file, payload: handler.events.find(event => event.name === 'open').payload };
    };
    const valid = await open('editable.xlsx', simple);
    assert.equal(valid.payload.nativeExcel, false);
    assert.equal(valid.payload.readOnly, false);
    await valid.handler.call('save', Array.from(roundtrip));
    assert.deepEqual(await readFile(valid.file), Buffer.from(roundtrip));
    state.destination = uri(join(temporary, 'copy.xlsx'));
    await valid.handler.call('saveAs', { content: Array.from(roundtrip), ext: 'xlsx' });
    assert.deepEqual(await readFile(state.destination.fsPath), Buffer.from(roundtrip));
    await assert.rejects(valid.handler.call('saveAs', { content: Array.from(roundtrip), ext: 'xlsx' }), /exist|new name|choose/i);
    await writeFile(valid.file, blank);
    await assert.rejects(valid.handler.call('save', Array.from(roundtrip)), /changed on disk/);
    assert.deepEqual(await readFile(valid.file), Buffer.from(blank));

    for (const [filename, bytes] of [['opaque.xlsx', opaque], ['rich.xlsx', richBytes], ['printed.xlsx', printedBytes], ['macros.xlsm', simple], ['legacy.xls', Buffer.from('synthetic legacy')]]) {
      const restricted = await open(filename, bytes);
      assert.equal(restricted.payload.readOnly, true, filename);
      assert.equal(restricted.payload.readOnlyReason, 'portable-preservation', filename);
      await assert.rejects(restricted.handler.call('save', Array.from(blank)), /portable/i);
      state.destination = uri(join(temporary, `${filename}-forged.xlsx`));
      await assert.rejects(restricted.handler.call('saveAs', { content: Array.from(blank), ext: 'xlsx' }), /portable/i);
      await restricted.handler.call('saveNative', { operations: [], expectedWorkbookSha256: '0'.repeat(64) });
      assert.ok(restricted.handler.events.some(event => event.name === 'writeBlocked' && /Windows x64/.test(event.payload.message)));
      assert.deepEqual(await readFile(restricted.file), Buffer.from(bytes));
    }
    const forged = await open('forged.xlsx', simple);
    await assert.rejects(forged.handler.call('save', Array.from(opaque)), /portable/i);
    state.destination = uri(join(temporary, 'forged-output.xlsx'));
    await assert.rejects(forged.handler.call('saveAs', { content: Array.from(opaque), ext: 'xlsx' }), /portable/i);
    assert.throws(() => forged.handler.call('openExcel'), /Windows x64/);
    assert.equal(state.commands.some(([name]) => name === 'excelAiVbaStudio.openWorkbookInExcel'), false);
  }
  console.log('Portable capability, preservation, scalar roundtrip and host save regressions passed.');
} finally {
  delete globalThis.__portableTest;
  await rm(temporary, { recursive: true, force: true });
}
