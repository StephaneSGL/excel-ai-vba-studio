const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const vscode = require('vscode');

exports.run = async function () {
  const extension = vscode.extensions.getExtension('steph-tools.excel-ai-vba-studio');
  assert.ok(extension, 'extension discovered');
  await extension.activate();
  assert.equal(extension.isActive, true, 'extension activated');
  const commands = new Set(await vscode.commands.getCommands(true));
  for (const { command } of extension.packageJSON.contributes.commands) {
    assert.ok(commands.has(command), `command registered: ${command}`);
  }
  assert.ok(vscode.lm?.tools, 'Language Model tools API available');
  const tools = new Set(vscode.lm.tools.map(tool => tool.name));
  for (const { name } of extension.packageJSON.contributes.languageModelTools) {
    assert.ok(tools.has(name), `AI tool registered: ${name}`);
  }
  const fixture = path.join(process.env.EXCEL_AI_SMOKE_WORKSPACE, 'synthetic.csv');
  await fs.writeFile(fixture, 'Name,Value\nSynthetic,42\n', 'utf8');
  try {
    await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(fixture), 'excelAiVbaStudio.officeViewer');
    assert.ok(vscode.window.tabGroups.all.flatMap(group => group.tabs).some(tab =>
      tab.input instanceof vscode.TabInputCustom && tab.input.viewType === 'excelAiVbaStudio.officeViewer'
    ), 'CSV custom editor opened');
    // This verifies host registration/opening, not webview rendering or Excel COM.
    console.log(`HOST_SMOKE_PASSED VS Code ${vscode.version}: activation, ${extension.packageJSON.contributes.commands.length} commands, ${tools.size} tools visible, CSV custom editor.`);
  } finally {
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  }
};
