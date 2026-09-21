import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { withModule } from './helpers/load-ts.mjs';
import { resolveWindowsCliPath } from './helpers/vscode-cli.mjs';

const cliFixture = await mkdtemp(join(tmpdir(), 'excel-cli-layout-'));
try {
  const executable = join(cliFixture, 'Code.exe');
  await assert.rejects(resolveWindowsCliPath(executable), /found 0/);
  const classic = join(cliFixture, 'resources/app/out/cli.js');
  const versioned = join(cliFixture, '7debcd0e2a/resources/app/out/cli.js');
  await mkdir(dirname(classic), { recursive: true });
  await writeFile(classic, '// synthetic CLI fixture');
  assert.equal(await resolveWindowsCliPath(executable), classic);
  await mkdir(dirname(versioned), { recursive: true });
  await writeFile(versioned, '// synthetic CLI fixture');
  await assert.rejects(resolveWindowsCliPath(executable), /found 2/);
  await rm(classic);
  assert.equal(await resolveWindowsCliPath(executable), versioned);
} finally {
  await rm(cliFixture, { recursive: true, force: true });
}

const snapshot = {
  checkedAt: '2026-09-21T12:00:00.000Z', extensionVersion: '0.6.1', vscodeVersion: '1.95.0',
  platform: 'win32', arch: 'x64', trusted: true, powershell: true, nativeHelper: true,
  excel: 'missing', aiTools: true,
};

await withModule('src/excelAiVbaStudio/setupDiagnostics.ts', async m => {
  assert.equal(m.parseExcelRegistration('detected\r\n'), 'detected');
  assert.equal(m.parseExcelRegistration('missing'), 'missing');
  for (const invalid of ['', 'access denied', 'true', 'detected\nerror']) {
    assert.equal(m.parseExcelRegistration(invalid), 'unknown');
  }
  assert.equal(m.buildSetupChecks(snapshot).find(c => c.id === 'excel').state, 'warning');
  assert.match(m.buildSetupChecks(snapshot).find(c => c.id === 'excel').detail, /sans Excel/);
  assert.equal(m.buildSetupChecks({ ...snapshot, excel: 'unknown' }).find(c => c.id === 'excel').state, 'unknown');
  assert.match(m.buildSetupChecks({ ...snapshot, excel: 'detected' }).find(c => c.id === 'excel').detail, /Signal du registre uniquement/);
  assert.equal(m.buildSetupChecks({ ...snapshot, platform: 'linux' }).find(c => c.id === 'platform').state, 'blocked');
  assert.equal(m.buildSetupChecks({ ...snapshot, arch: 'arm64' }).find(c => c.id === 'platform').state, 'blocked');
  assert.equal(m.buildSetupChecks({ ...snapshot, trusted: false }).find(c => c.id === 'trust').state, 'blocked');
  assert.equal(m.buildSetupChecks({ ...snapshot, nativeHelper: false }).find(c => c.id === 'helper').state, 'warning');
  const report = m.formatSetupReport({ ...snapshot, filePath: 'C:\\secret.xlsx', token: 'PRIVATE' });
  assert.doesNotMatch(report, /secret|PRIVATE/);
  assert.match(report, /pas une certification/);
  assert.match(m.EXCEL_REGISTRATION_PROBE, /Registry64/);
  assert.match(m.EXCEL_REGISTRATION_PROBE, /Registry32/);
  assert.doesNotMatch(m.EXCEL_REGISTRATION_PROBE, /CreateInstance|ComObject|CreateSubKey|SetValue|AccessVBOM/i);
});

await withModule('src/excelAiVbaStudio/startPageHtml.ts', async ({ renderStartPage }) => {
  const html = renderStartPage({ ...snapshot, extensionVersion: '<script>alert(1)</script>' }, 'nonce123');
  assert.match(html, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.doesNotMatch(html, /<script>alert/);
  assert.match(html, /default-src 'none'/);
  assert.match(html, /script-src 'nonce-nonce123'/);
  assert.doesNotMatch(html, /unsafe-inline|https?:\/\/|command:|onclick=/);
  assert.match(html, /lang="fr"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /\.textContent = event\.data\.text/);
  assert.match(renderStartPage(undefined, 'nonce123'), /data-action="copyReport" disabled/);
  assert.throws(() => renderStartPage(snapshot, '\"><script>'), /nonce/);
}, { vscode: 'export const Uri = {};' });

const state = {
  commands: [], dialogs: 0, choice: undefined, clipboard: [], panels: [], feedback: [],
  trusted: true, probeCalls: 0, releases: [], info: undefined,
};
globalThis.__setupStartTest = state;
class Uri {
  constructor(fsPath, scheme = 'file', authority = '') { Object.assign(this, { fsPath, scheme, authority }); }
  static file(fsPath) { return new Uri(fsPath); }
}
state.Uri = Uri;
try {
  await withModule('src/excelAiVbaStudio/startPage.ts', async m => {
    for (const invalid of [null, [], 'openWorkbook', {}, { action: 'workbench.action.terminal.new' }, { action: 12 }]) {
      assert.equal(m.parseStartPageAction(invalid), undefined);
    }
    assert.equal(m.parseStartPageAction({ action: 'openWorkbook', path: 'untrusted.xlsx' }), 'openWorkbook');
    await m.openWorkbookInGrid(new Uri('C:\\Test\\Budget.XLSX'));
    assert.deepEqual(state.commands[0], ['vscode.openWith', new Uri('C:\\Test\\Budget.XLSX'), 'excelAiVbaStudio.officeViewer']);
    for (const uri of [new Uri('C:\\bad.xlsb'), new Uri('C:\\bad.exe'), new Uri('/bad.xlsx', 'https'), new Uri('\\\\server\\bad.xlsx', 'file', 'server')]) {
      await assert.rejects(m.openWorkbookInGrid(uri), /local/);
    }
    await m.openWorkbookInGrid({ fsPath: 'C:\\injected.xlsx', scheme: 'file' });
    assert.equal(state.dialogs, 1, 'non-URI input must not bypass the picker');
    assert.equal(state.commands.length, 1, 'cancellation opens nothing');
    state.choice = [Uri.file('C:\\Test\\data.csv')];
    await m.openWorkbookInGrid();
    assert.equal(state.commands.length, 2);
    state.trusted = false;
    await assert.rejects(m.openWorkbookInGrid(state.choice[0]), /confiance/);
    state.trusted = true;

    const page = new m.ExcelStartPage({ extensionPath: 'C:\\Extension', extension: { packageJSON: {
      version: '0.6.1', contributes: { languageModelTools: [{ name: 'tool-a' }, { name: 'tool-b' }] },
    } } });
    const first = page.show();
    const repeated = page.show();
    assert.equal(state.panels.length, 1);
    assert.equal(state.probeCalls, 1, 'repeated refreshes share one bounded probe');
    assert.equal(state.info.aiTools, false, 'partial tool registration is not reported as ready');
    state.releases.shift()(snapshot);
    await Promise.all([first, repeated]);
    const panel = state.panels[0];
    assert.deepEqual(panel.options.localResourceRoots, []);
    assert.equal(panel.options.enableCommandUris, false);
    panel.message({ action: 'copyReport' });
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(state.clipboard.length, 1);
    assert.match(state.feedback.at(-1).text, /Diagnostic copié/);
    const commandCount = state.commands.length;
    panel.message({ action: 'openWorkbook', path: 'C:\\injected.xlsx' });
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(state.commands.at(-1), ['excelAiVbaStudio.openWorkbook'], 'webview payload cannot supply a path');
    panel.message({ action: 'workbench.action.terminal.new' });
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(state.commands.length, commandCount + 1);
    const closing = page.show();
    page.dispose();
    state.releases.shift()(snapshot);
    await closing;
    assert.equal(panel.disposed, true, 'probe finishing after close does not reuse a disposed panel');
  }, {
    vscode: `const s = globalThis.__setupStartTest;
      export const Uri = s.Uri;
      export const workspace = { get isTrusted() { return s.trusted; } };
      export const version = '1.95.0';
      export const lm = { tools: [{ name: 'tool-a' }] };
      export const ViewColumn = { One: 1 };
      export const commands = { async executeCommand(...args) { s.commands.push(args); } };
      export const env = { clipboard: { async writeText(value) { s.clipboard.push(value); } } };
      export const window = {
        async showOpenDialog() { s.dialogs++; return s.choice; },
        createWebviewPanel(type, title, column, options) {
          const panel = { options, disposed: false, reveal() {},
            webview: { html: '', onDidReceiveMessage(fn) { panel.message = fn; return { dispose() {} }; },
              async postMessage(message) { s.feedback.push(message); return true; } },
            onDidDispose(fn) { panel.onClose = fn; return { dispose() {} }; },
            dispose() { panel.disposed = true; panel.onClose?.(); } };
          s.panels.push(panel); return panel;
        }
      };`,
    './setupDiagnostics': `const s = globalThis.__setupStartTest;
      export function collectSetupSnapshot(path, info) { s.probeCalls++; s.info = info; return new Promise(resolve => s.releases.push(resolve)); }
      export function formatSetupReport(snapshot) { return 'Test diagnostic ' + snapshot.extensionVersion; }
      export function buildSetupChecks() { return []; }`,
  });
} finally {
  delete globalThis.__setupStartTest;
}
console.log('Setup/start-page regressions passed: capability states, privacy, CSP, action allowlist, picker, trust, panel lifecycle.');
