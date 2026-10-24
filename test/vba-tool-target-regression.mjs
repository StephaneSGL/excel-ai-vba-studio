import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { pathToFileURL } from 'node:url';
import { withModule } from './helpers/load-ts.mjs';

const targetA = resolve(tmpdir(), 'BudgetA.xlsm');
const targetB = resolve(tmpdir(), 'BudgetB.xlsm');
const registered = new Map();
globalThis.__excelTestTools = registered;
try {
  // Mock filesystem/COM boundaries only: exercise real registration, parsers,
  // explicit target resolution, confirmation and invocation orchestration.
  await withModule('src/excelAiVbaStudio/languageModelTool.ts', async ({ registerExcelAiVbaLanguageModelTool }) => {
    let active = targetA;
    const writes = [];
    registerExcelAiVbaLanguageModelTool({ subscriptions: [] }, {
      getOutputChannel: () => ({ appendLine() {} }),
      resolveToolWorkbookUri: () => { throw new Error(`Unexpected active-workbook fallback: ${active}`); },
      writeVbaFromTool: async uri => { writes.push(uri.fsPath); return { targetWorkbookPath: uri.fsPath }; },
      designVbaFromTool: async uri => { writes.push(uri.fsPath); return { targetWorkbookPath: uri.fsPath }; },
    });
    for (const [name, input] of [
      ['excel_ai_vba_writeModule', { componentFile: 'Module1.bas', source: 'Option Explicit' }],
      ['excel_ai_vba_designWorkbook', { operations: [{ kind: 'createUserForm', name: 'Form1' }] }],
    ]) {
      const tool = registered.get(name);
      assert.ok(tool, name);
      for (const workbookPath of [undefined, '', '  ', 'Budget.xlsm', 123, 'C:relative.xlsm', 'https://example.test/file.xlsm']) {
        await assert.rejects(tool.prepareInvocation({ input: { ...input, workbookPath } }), /workbookPath/);
        await assert.rejects(tool.invoke({ input: { ...input, workbookPath } }), /workbookPath/);
      }
      const args = { input: { ...input, workbookPath: targetA } };
      const prepared = await tool.prepareInvocation(args);
      assert.ok(prepared.confirmationMessages.message.includes(targetA));
      active = targetB;
      await tool.invoke(args);
      assert.equal(writes.at(-1), targetA);
      await tool.invoke({ input: { ...input, workbookPath: pathToFileURL(targetA).href } });
      assert.equal(writes.at(-1), targetA);
      await assert.rejects(tool.invoke(args, { isCancellationRequested: true }), /cancel/i);
    }
    assert.equal(writes.length, 4);
  }, {
    vscode: `import { fileURLToPath } from 'node:url';
      export const Uri = { file: p => ({ scheme: 'file', authority: '', fsPath: p }), parse: p => Uri.file(fileURLToPath(p)) };
      export const workspace = { workspaceFolders: [] };
      export const lm = { registerTool(name, tool) { globalThis.__excelTestTools.set(name, tool); return { dispose() {} }; } };
      export class CancellationError extends Error { constructor() { super('Cancelled'); } }
      export class LanguageModelToolResult { constructor(parts) { this.parts = parts; } }
      export class LanguageModelTextPart { constructor(value) { this.value = value; } }`,
    './security': `import { Uri } from 'vscode';
      export const assertNoReparsePointChain = async () => {};
      export const canonicalizeWorkbookUri = async uri => uri;
      export const pathIsInside = () => false;
      export const workbookUriFromPathInput = p => /^file:/i.test(p) ? Uri.parse(p) : Uri.file(p);`,
  });
} finally {
  delete globalThis.__excelTestTools;
}
console.log('VBA mutation targeting regressions passed (COM/filesystem boundaries mocked).');
