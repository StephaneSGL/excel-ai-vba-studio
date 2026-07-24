import { handleClass } from '@/provider/handlers/classHandler';
import * as vscode from 'vscode';

/**
 * Desktop-only Java decompiler custom editor (`excelAiVbaStudio.classViewer`).
 */
export class ClassViewerProvider implements vscode.CustomReadonlyEditorProvider {

	constructor(private context: vscode.ExtensionContext) { }

	bindCustomEditor(viewOption: { webviewOptions: vscode.WebviewPanelOptions }) {
		return vscode.window.registerCustomEditorProvider('excelAiVbaStudio.classViewer', this, viewOption);
	}

	openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
		return { uri, dispose: (): void => { } };
	}

	resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel): void {
		webviewPanel.webview.options = {
			enableScripts: false,
			localResourceRoots: [],
		};
		handleClass(document.uri, webviewPanel);
	}
}
