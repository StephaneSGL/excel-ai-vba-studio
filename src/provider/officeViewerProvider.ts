import { getReactWebviewResourceRoots } from '@/common/extensionResource';
import { getFileSuffix } from '@/common/fileSuffix';
import { Handler } from '@/common/handler';
import { ReactApp } from '@/common/reactApp';
import * as vscode from 'vscode';
import { handleCommonEvent } from './compress/commonHandler';

const SUPPORTED_SPREADSHEET_SUFFIXES = new Set([
	'.xlsx',
	'.xlsm',
	'.xls',
	'.csv',
	'.tsv',
]);

/**
 * Hosts the Excel/CSV webview. Other legacy Office viewers are intentionally
 * excluded from the published runtime.
 */
export class OfficeViewerProvider implements vscode.CustomReadonlyEditorProvider {
	constructor(private readonly context: vscode.ExtensionContext) {}

	bindCustomEditor(
		viewOption: { webviewOptions: vscode.WebviewPanelOptions }
	): vscode.Disposable {
		return vscode.window.registerCustomEditorProvider(
			'excelAiVbaStudio.officeViewer',
			this,
			viewOption
		);
	}

	openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
		return { uri, dispose: (): void => {} };
	}

	resolveCustomEditor(
		document: vscode.CustomDocument,
		webviewPanel: vscode.WebviewPanel
	): void | Thenable<void> {
		const uri = document.uri;
		const suffix = getFileSuffix(uri.fsPath);
		if (!SUPPORTED_SPREADSHEET_SUFFIXES.has(suffix)) {
			void vscode.commands.executeCommand('vscode.openWith', uri, 'default');
			return;
		}

		const webview = webviewPanel.webview;
		const folderPath = vscode.Uri.joinPath(uri, '..');
		webview.options = {
			enableScripts: true,
			localResourceRoots: [
				...getReactWebviewResourceRoots(this.context),
				folderPath,
			],
		};

		const handler = Handler.bind(webviewPanel, uri);
		handleCommonEvent(uri, handler);
		return ReactApp.view(webview, { route: 'excel' });
	}
}
