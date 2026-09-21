const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const vscode = require('vscode');

exports.run = async function () {
  const extension = vscode.extensions.getExtension('steph-tools.excel-ai-vba-studio');
  assert.ok(extension, 'extension discovered');
  if (process.env.EXCEL_AI_SMOKE_EXTENSION) {
    assert.equal(await fs.realpath(extension.extensionPath), await fs.realpath(process.env.EXCEL_AI_SMOKE_EXTENSION),
      'host must load the requested installed package, not another development copy');
  }
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
    await vscode.commands.executeCommand('excelAiVbaStudio.openStart');
    assert.ok(vscode.window.tabGroups.all.flatMap(group => group.tabs).some(tab =>
      tab.input instanceof vscode.TabInputWebview && tab.label === 'Excel & VBA · Accueil'
    ), 'start page and local diagnostics opened');
    await vscode.commands.executeCommand('excelAiVbaStudio.openStart');
    assert.equal(vscode.window.tabGroups.all.flatMap(group => group.tabs).filter(tab =>
      tab.input instanceof vscode.TabInputWebview && tab.label === 'Excel & VBA · Accueil'
    ).length, 1, 'start page reuses its panel');
    await vscode.commands.executeCommand('excelAiVbaStudio.openWorkbook', vscode.Uri.file(fixture));
    assert.ok(vscode.window.tabGroups.all.flatMap(group => group.tabs).some(tab =>
      tab.input instanceof vscode.TabInputCustom && tab.input.viewType === 'excelAiVbaStudio.officeViewer'
    ), 'CSV custom editor opened');
    // This verifies host registration/opening, not webview rendering or Excel COM.
    console.log(`HOST_SMOKE_PASSED VS Code ${vscode.version}: activation, ${extension.packageJSON.contributes.commands.length} commands, ${tools.size} tools visible, start page, diagnostics, guided CSV opening.`);
  } finally {
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  }
};
