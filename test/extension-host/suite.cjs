const assert = require('node:assert/strict');
const { readFile, writeFile, realpath } = require('node:fs/promises');
const { join, relative, isAbsolute, sep } = require('node:path');
const vscode = require('vscode');
const JSZip = require('jszip');
const { chromium } = require('playwright-core');

async function step(label, action, timeout = 30_000) {
  const started = Date.now();
  console.log(`[smoke] START ${label}`);
  let timer;
  try {
    const result = await Promise.race([
      Promise.resolve().then(action),
      new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(`Timed out after ${timeout}ms: ${label}`)), timeout); }),
    ]);
    console.log(`[smoke] PASS ${label} (${Date.now() - started}ms)`);
    return result;
  } catch (error) {
    console.error(`[smoke] FAIL ${label}: ${error.message}`);
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function until(check, label, timeout = 60_000) {
  const deadline = Date.now() + timeout;
  let lastError;
  while (Date.now() < deadline) {
    try {
      const result = await check();
      if (result) return result;
    } catch (error) { lastError = error; }
    await new Promise(resolve => setTimeout(resolve, 150));
  }
  throw new Error(`Timed out: ${label}${lastError ? ` (${lastError.message})` : ''}`);
}

async function spreadsheetFrame(browser) {
  return until(async () => {
    for (const context of browser.contexts()) {
      for (const page of context.pages()) {
        for (const frame of page.frames()) {
          if (await frame.locator('.excel-viewer .x-spreadsheet-formula-bar-name').isVisible()) {
            return frame;
          }
        }
      }
    }
    return null;
  }, 'spreadsheet webview and formula bar');
}

async function selectCell(frame, address) {
  const cellAddress = frame.getByRole('textbox', { name: 'Cell address', exact: true });
  await cellAddress.fill(address);
  await cellAddress.press('Enter');
}

async function setCell(frame, address, value) {
  await selectCell(frame, address);
  const content = frame.getByRole('textbox', { name: 'Cell content', exact: true });
  await content.fill(value);
  await content.press('Enter');
}

async function openWorkbook(browser, filePath) {
  await step('open workbook custom editor', () => vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(filePath), 'excelAiVbaStudio.officeViewer'));
  return spreadsheetFrame(browser);
}

async function closeEditors() {
  assert.equal(vscode.window.tabGroups.all.some(group => group.tabs.some(tab => tab.isDirty)), false,
    'an expected save did not finish; do not open an interactive save prompt during cleanup');
  await step('close saved editors', () => vscode.commands.executeCommand('workbench.action.closeAllEditors'), 10_000);
}

function isSyntheticTab(tab, workspace) {
  const uri = tab.input?.uri;
  if (uri?.scheme !== 'file') return false;
  const local = relative(workspace, uri.fsPath);
  return local !== '' && local !== '..' && !local.startsWith(`..${sep}`) && !isAbsolute(local);
}

async function visibleNotifications(browser) {
  const notifications = [];
  for (const context of browser.contexts()) {
    for (const page of context.pages()) {
      for (const item of await page.locator('.notification-list-item').all()) {
        if (!await item.isVisible()) continue;
        notifications.push({
          message: await item.locator('.notification-list-item-message').innerText({ timeout: 1_000 }),
          error: await item.locator('.codicon-error').count() > 0,
        });
      }
    }
  }
  return notifications;
}

async function quickInputState(page) {
  const input = page.locator('.quick-input-widget .quick-input-box input').first();
  const present = await input.count() > 0;
  const visible = present && await input.isVisible();
  const value = present ? await input.inputValue({ timeout: 1_000 }).catch(() => undefined) : undefined;
  return { present, visible, visibleValue: visible ? value : undefined, hiddenDomValue: visible ? undefined : value };
}

async function captureFailure(browser, error, contextInfo) {
  const prefix = join(process.env.EXCEL_SMOKE_SCREENSHOTS, `${process.platform}-${process.arch}-vscode-${vscode.version}-failure`);
  const diagnostic = { error: error.stack || String(error), ...contextInfo, tabs: vscode.window.tabGroups.all.flatMap(group => group.tabs.map(tab => ({ label: tab.label, dirty: tab.isDirty, uri: tab.input?.uri?.toString() }))), pages: [] };
  for (const context of browser.contexts()) {
    for (const page of context.pages()) {
      if (!await page.locator('.monaco-workbench').isVisible()) continue;
      // Showing notification history can change focus: record visibility first.
      const quickInputBeforeNotificationCenter = await quickInputState(page);
      await step('show notification history', () => vscode.commands.executeCommand('notifications.showList'), 2_000)
        .catch(error => { diagnostic.notificationCenterError = error.message; });
      diagnostic.pages.push({
        notifications: await page.locator('.notification-list-item-message').allTextContents(),
        dialogs: await page.locator('.monaco-dialog-box').allTextContents(),
        quickInputBeforeNotificationCenter,
        quickInputAfterNotificationCenter: await quickInputState(page),
      });
      await page.screenshot({ path: `${prefix}.png`, timeout: 3_000 }).catch(error => { diagnostic.screenshotError = error.message; });
      break;
    }
  }
  await writeFile(`${prefix}.json`, JSON.stringify(diagnostic, null, 2));
  console.error(`[smoke] Failure diagnostics: ${JSON.stringify(diagnostic)}`);
}

async function cleanupSyntheticEditors(workspace) {
  await step('dismiss test dialogs and notifications', async () => {
    await vscode.commands.executeCommand('workbench.action.closeQuickOpen');
    await vscode.commands.executeCommand('notifications.clearAll');
  }, 5_000);
  const tabs = vscode.window.tabGroups.all.flatMap(group => group.tabs).filter(tab => isSyntheticTab(tab, workspace));
  for (const tab of tabs) {
    await step(`discard synthetic editor ${tab.label}`, async () => {
      if (tab.isDirty) {
        await vscode.commands.executeCommand('vscode.openWith', tab.input.uri, 'excelAiVbaStudio.officeViewer');
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
      } else {
        await vscode.window.tabGroups.close(tab);
      }
    }, 5_000);
  }
}

exports.run = async function run() {
  const { ExcelJS } = require(process.env.EXCEL_SMOKE_EXCELJS_FIXTURE);
  const workspace = process.env.EXCEL_SMOKE_WORKSPACE;
  assert.ok(workspace, 'synthetic workspace must be provided');
  const actualWorkspace = vscode.workspace.workspaceFolders[0].uri.fsPath;
  const contextInfo = {
    workspace: { lexical: workspace, canonical: await realpath(workspace), uri: vscode.workspace.workspaceFolders[0].uri.toString(), uriFsPath: actualWorkspace },
    notificationsDuringCreation: [],
  };
  console.log(`[smoke] Synthetic workspace paths: ${JSON.stringify(contextInfo.workspace)}`);
  // VS Code normalizes Windows drive-letter casing in URI.fsPath.
  assert.equal(
    process.platform === 'win32' ? actualWorkspace.toLowerCase() : actualWorkspace,
    process.platform === 'win32' ? workspace.toLowerCase() : workspace,
  );
  const extension = vscode.extensions.getExtension('steph-tools.excel-ai-vba-studio');
  assert.ok(extension, 'the VSIX must be installed in the isolated profile');
  assert.ok(!extension.extensionPath.includes(join('test', 'extension-host')), 'test the installed product');
  await step('activate installed extension', () => extension.activate());
  assert.equal(extension.isActive, true);
  const commands = await vscode.commands.getCommands(true);
  assert.ok(commands.includes('excelAiVbaStudio.newWorkbook'), 'new-workbook command must be registered');

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Synthetic');
  sheet.getCell('A1').value = 'original';
  sheet.getCell('B1').value = 2;
  const path = join(workspace, 'portable-smoke.xlsx');
  await writeFile(path, Buffer.from(await workbook.xlsx.writeBuffer()));
  const original = await readFile(path);
  const browser = await step('connect to VS Code UI', () => chromium.connectOverCDP(process.env.EXCEL_SMOKE_CDP_ENDPOINT, { timeout: 15_000 }));
  let failure;
  try {
    let frame = await openWorkbook(browser, path);
    assert.equal(await frame.locator('.excel-load-error').count(), 0);
    await selectCell(frame, 'A1');
    await until(async () => await frame.getByRole('textbox', { name: 'Cell content', exact: true }).inputValue() === 'original', 'initial cell content');
    await setCell(frame, 'A1', 'portable-edited');
    await setCell(frame, 'B1', '=SUM(2,3)');
    await until(() => vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'custom editor dirty event');
    await step('save edited workbook', () => vscode.commands.executeCommand('workbench.action.files.save'));
    await until(async () => {
      if (vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty) return false;
      return !(await readFile(path)).equals(original);
    }, 'custom editor save');
    const saved = new ExcelJS.Workbook();
    await saved.xlsx.load(await readFile(path));
    assert.equal(saved.getWorksheet('Synthetic').getCell('A1').value, 'portable-edited');
    assert.equal(saved.getWorksheet('Synthetic').getCell('B1').formula, 'SUM(2,3)');
    await closeEditors();
    frame = await openWorkbook(browser, path);
    await selectCell(frame, 'A1');
    assert.equal(await frame.getByRole('textbox', { name: 'Cell content', exact: true }).inputValue(), 'portable-edited');
    const screenshot = join(process.env.EXCEL_SMOKE_SCREENSHOTS, `${process.platform}-${process.arch}-vscode-${vscode.version}.png`);
    await frame.locator('.excel-viewer').screenshot({ path: screenshot });
    console.log(`Synthetic workbook screenshot: ${screenshot}`);
    await closeEditors();

    const createdPath = join(workspace, 'created-through-command.xlsx');
    contextInfo.expectedTarget = { path: createdPath, uri: vscode.Uri.file(createdPath).toString() };
    const oldErrors = new Set((await visibleNotifications(browser)).filter(item => item.error).map(item => item.message));
    let rejectNotification;
    const notificationFailure = new Promise((_, reject) => { rejectNotification = reject; });
    notificationFailure.catch(() => {});
    let inspectingNotifications = false;
    const notificationTimer = setInterval(async () => {
      if (inspectingNotifications) return;
      inspectingNotifications = true;
      try {
        for (const notice of await visibleNotifications(browser)) {
          if (!contextInfo.notificationsDuringCreation.some(item => item.message === notice.message)) {
            contextInfo.notificationsDuringCreation.push(notice);
            console.log(`[smoke] Notification during creation: ${JSON.stringify(notice)}`);
          }
          if (notice.error && !oldErrors.has(notice.message)) rejectNotification(new Error(`New workbook error notification: ${notice.message}`));
        }
      } catch (error) { contextInfo.notificationInspectionError = error.message; }
      finally { inspectingNotifications = false; }
    }, 100);
    try {
      await step('create workbook through save dialog', () => Promise.race([notificationFailure, (async () => {
      const creation = Promise.resolve(vscode.commands.executeCommand('excelAiVbaStudio.newWorkbook'));
      // Observe an early rejection while the save-dialog locator is still waiting.
      creation.catch(() => {});
      const filename = await until(async () => {
        for (const context of browser.contexts()) {
          for (const page of context.pages()) {
            const input = page.locator('.quick-input-widget .quick-input-box input');
            if (await input.isVisible()) return input;
          }
        }
        return null;
      }, 'new workbook save dialog', 25_000);
      await filename.fill(createdPath, { timeout: 5_000 });
      contextInfo.dialogInputAtSubmit = { visible: await filename.isVisible(), value: await filename.inputValue() };
      console.log(`[smoke] New workbook dialog input: ${JSON.stringify(contextInfo.dialogInputAtSubmit)}`);
      await filename.press('Enter', { timeout: 5_000 });
      await creation;
      })()]));
    } finally { clearInterval(notificationTimer); }
    frame = await spreadsheetFrame(browser);
    await setCell(frame, 'A1', 'created-in-vscode');
    await until(() => vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'new workbook dirty event');
    await step('save newly created workbook', () => vscode.commands.executeCommand('workbench.action.files.save'));
    await until(() => !vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'new workbook save');
    const created = new ExcelJS.Workbook();
    await created.xlsx.load(await readFile(createdPath));
    assert.equal(created.getWorksheet('Sheet1').getCell('A1').value, 'created-in-vscode');
    await closeEditors();

    if (process.platform !== 'win32') {
      for (const tool of extension.packageJSON.contributes.languageModelTools) {
        assert.equal(tool.when, 'excelAiVbaStudio.nativeExcel', 'native tool UI must require the Windows capability');
      }
      // Newer VS Code exposes disabled contributions in lm.tools as well.
      // Verify runtime unavailability instead of mistaking metadata for an implementation.
      const nativeReadTool = vscode.lm?.tools.find(tool => tool.name === 'excel_ai_vba_readWorkbook');
      if (nativeReadTool) {
        await assert.rejects(
          vscode.lm.invokeTool(nativeReadTool.name, {
            input: { workbookPath: path, includeVba: false, format: 'json' },
            toolInvocationToken: undefined,
          }),
          /not.*implementation.*registered/i,
          'native tool invocation must fail before any workbook automation starts',
        );
      }
      const archive = await JSZip.loadAsync(original);
      const types = await archive.file('[Content_Types].xml').async('string');
      archive.file('[Content_Types].xml', types.replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      ));
      const macroPath = join(workspace, 'read-only.xlsm');
      await writeFile(macroPath, await archive.generateAsync({ type: 'nodebuffer' }));
      const before = await readFile(macroPath);
      frame = await openWorkbook(browser, macroPath);
      assert.equal(await frame.getByRole('textbox', { name: 'Cell content', exact: true }).isEditable(), false);
      assert.ok(await frame.locator('.excel-readonly-banner').isVisible(), 'macro workbook must explain read-only mode');
      await frame.locator('.excel-viewer').screenshot({
        path: join(process.env.EXCEL_SMOKE_SCREENSHOTS, `${process.platform}-${process.arch}-vscode-${vscode.version}-read-only.png`),
      });
      await closeEditors();
      assert.deepEqual(await readFile(macroPath), before, 'read-only workbook bytes must be unchanged');
    }
  } catch (error) {
    failure = error;
    await step('capture failure diagnostics', () => captureFailure(browser, error, contextInfo), 8_000).catch(() => {});
  } finally {
    try { await cleanupSyntheticEditors(workspace); } catch (error) { failure ||= error; }
    try { await step('disconnect test browser', () => browser.close(), 5_000); } catch (error) { failure ||= error; }
  }
  if (failure) throw failure;
  console.log('Installed VSIX smoke passed: activation, custom editor rendering, cell/formula edit, host save, disk verification, reopen, new-workbook command, and portable macro protection.');
};
