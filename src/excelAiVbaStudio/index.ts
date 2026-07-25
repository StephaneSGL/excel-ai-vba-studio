import * as path from 'path';
import * as vscode from 'vscode';
import {
	ExcelAiVbaExplorerProvider,
	ExcelAiVbaPropertiesProvider
} from './explorer';
import { registerExcelAiVbaLanguageModelTool } from './languageModelTool';
import {
	EXCEL_AI_COMMANDS,
	EXCEL_AI_EXPLORER_VIEW,
	EXCEL_AI_PROPERTIES_VIEW
} from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

export function registerExcelAiVbaStudio(
	context: vscode.ExtensionContext
): void {
	const service = new ExcelAiVbaWorkbookService(context);
	const explorer = new ExcelAiVbaExplorerProvider(service);
	const properties = new ExcelAiVbaPropertiesProvider();
	const explorerTree = vscode.window.createTreeView(EXCEL_AI_EXPLORER_VIEW, {
		treeDataProvider: explorer,
		showCollapseAll: true
	});
	context.subscriptions.push(
		service,
		explorer,
		properties,
		service.onDidChangeContext(() => explorer.refresh()),
		explorerTree,
		explorerTree.onDidChangeSelection(event => {
			properties.setSelected(event.selection[0]);
		}),
		vscode.window.createTreeView(EXCEL_AI_PROPERTIES_VIEW, {
			treeDataProvider: properties,
			showCollapseAll: false
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
			EXCEL_AI_COMMANDS.openVbaDeveloper,
			(candidate?: unknown) => service.openVbaExplorer(candidate)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openVbaExplorer,
			(candidate?: unknown) => service.openVbaExplorer(candidate)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.openVbaComponent,
			(candidate?: unknown) => service.openVbaComponent(candidate)
		),
		vscode.commands.registerCommand(
			EXCEL_AI_COMMANDS.askCopilotAboutWorkbook,
			(candidate?: unknown) => service.askCopilotAboutWorkbook(candidate)
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
	EXCEL_AI_PROPERTIES_VIEW,
	EXCEL_AI_LANGUAGE_MODEL_TOOL
} from './types';
