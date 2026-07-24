import * as path from 'path';
import * as vscode from 'vscode';
import { ExcelAiVbaExplorerProvider } from './explorer';
import { registerExcelAiVbaLanguageModelTool } from './languageModelTool';
import { EXCEL_AI_COMMANDS, EXCEL_AI_EXPLORER_VIEW } from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

export function registerExcelAiVbaStudio(
	context: vscode.ExtensionContext
): void {
	const service = new ExcelAiVbaWorkbookService(context);
	const explorer = new ExcelAiVbaExplorerProvider(service);
	context.subscriptions.push(
		service,
		explorer,
		service.onDidChangeContext(() => explorer.refresh()),
		vscode.window.createTreeView(EXCEL_AI_EXPLORER_VIEW, {
			treeDataProvider: explorer,
			showCollapseAll: true
		}),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.exportWorkbook,
			async (candidate?: unknown) => {
				const result = await service.exportWorkbook(candidate, {
					open: true
				});
				if (result) {
					await vscode.window.showInformationMessage(
						`Contexte IA créé dans le stockage de l’extension (${path.basename(
							result.paths.markdownPath
						)}).`
					);
				}
			}
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.copyWorkbookContext,
			async (candidate?: unknown) => {
				const result = await service.exportWorkbook(candidate, {
					open: false
				});
				if (result) {
					await service.copyExportResult(result);
				}
			}
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openWorkbookContext,
			(candidate?: unknown) => service.openWorkbookContext(candidate)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.copyGeneratedContext,
			() => service.copyGeneratedContext()
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openFullExcel,
			(candidate?: unknown) => service.openExcel(candidate, false)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openVbaDeveloper,
			(candidate?: unknown) => service.openExcel(candidate, true)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openVbaExplorer,
			(candidate?: unknown) => service.openVbaExplorer(candidate)
		),
		vscode.commands.registerCommand(EXCEL_AI_COMMANDS.refreshExplorer, () =>
			explorer.refresh()
		),
		vscode.commands.registerCommand(EXCEL_AI_COMMANDS.cleanExports, () =>
			service.cleanExports()
		),
		vscode.window.onDidChangeActiveTextEditor(() => explorer.refresh())
	);

	const vscodeWindow = vscode.window as typeof vscode.window & {
		tabGroups?: {
			onDidChangeTabs?: (listener: () => unknown) => vscode.Disposable;
		};
	};
	if (vscodeWindow.tabGroups?.onDidChangeTabs) {
		context.subscriptions.push(
			vscodeWindow.tabGroups.onDidChangeTabs(() => explorer.refresh())
		);
	}

	const settings = service.getSettings();
	service
		.getOutputChannel()
		.appendLine(
			`[config] maxRows=${settings.maxRows}, maxColumns=${settings.maxColumns}, includeVba=${settings.includeVba}`
		);
	registerExcelAiVbaLanguageModelTool(context, service);
}

export {
	EXCEL_AI_COMMANDS,
	EXCEL_AI_EXPLORER_VIEW,
	EXCEL_AI_LANGUAGE_MODEL_TOOL
} from './types';
