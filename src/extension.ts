import * as vscode from 'vscode';
import { ReactApp } from './common/reactApp';
import { registerExcelAiVbaStudio } from './excelAiVbaStudio';
import { OfficeViewerProvider } from './provider/officeViewerProvider';

export function activate(context: vscode.ExtensionContext): void {
	registerExcelAiVbaStudio(context);
	ReactApp.init(context);

	const viewer = new OfficeViewerProvider(context);
	context.subscriptions.push(
		viewer.bindCustomEditor({
			webviewOptions: { retainContextWhenHidden: true }
		})
	);
}

export function deactivate(): void {}
