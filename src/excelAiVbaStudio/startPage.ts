import { randomBytes } from 'crypto';
import * as path from 'path';
import * as vscode from 'vscode';
import { collectSetupSnapshot, formatSetupReport, type SetupSnapshot } from './setupDiagnostics';
import { renderStartPage } from './startPageHtml';
import { EXCEL_AI_COMMANDS } from './types';

const GRID_EXTENSIONS = new Set(['.xlsx', '.xlsm', '.xls', '.csv', '.tsv']);
const PAGE_ACTIONS = new Set(['openWorkbook', 'refresh', 'copyReport']);

export function parseStartPageAction(message: unknown): string | undefined {
	if (!message || typeof message !== 'object' || Array.isArray(message)) return undefined;
	const action = (message as { action?: unknown }).action;
	return typeof action === 'string' && PAGE_ACTIONS.has(action) ? action : undefined;
}

export async function openWorkbookInGrid(candidate?: unknown): Promise<void> {
	if (!vscode.workspace.isTrusted) throw new Error('Ouvrez un espace de travail de confiance avant d’utiliser le tableur.');
	// Only explicit VS Code URIs are accepted. No command or path comes from the webview.
	const uri = candidate instanceof vscode.Uri ? candidate : (await vscode.window.showOpenDialog({
		canSelectMany: false, canSelectFiles: true, canSelectFolders: false,
		openLabel: 'Ouvrir dans le tableur',
		filters: { 'Classeurs et données': ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'] }
	}))?.[0];
	if (!uri) return;
	if (uri.scheme !== 'file' || uri.authority || !GRID_EXTENSIONS.has(path.extname(uri.fsPath).toLowerCase())) {
		throw new Error('Choisissez un fichier local XLSX, XLSM, XLS, CSV ou TSV. XLSB n’a pas de grille intégrée.');
	}
	await vscode.commands.executeCommand('vscode.openWith', uri, 'excelAiVbaStudio.officeViewer');
}

export class ExcelStartPage implements vscode.Disposable {
	private panel: vscode.WebviewPanel | undefined;
	private snapshot: SetupSnapshot | undefined;
	private refreshing: Promise<void> | undefined;
	private opening = false;
	private readonly subscriptions: vscode.Disposable[] = [];

	constructor(private readonly context: vscode.ExtensionContext) {}

	async show(): Promise<void> {
		if (!this.panel) {
			const panel = vscode.window.createWebviewPanel('excelAiVbaStudio.start', 'Excel & VBA · Accueil', vscode.ViewColumn.One,
				{ enableScripts: true, localResourceRoots: [], enableCommandUris: false });
			this.panel = panel;
			panel.webview.html = renderStartPage(undefined, randomBytes(18).toString('hex'));
			panel.onDidDispose(() => {
				if (this.panel === panel) this.panel = undefined;
				this.subscriptions.splice(0).forEach(subscription => subscription.dispose());
			}, undefined, this.subscriptions);
			panel.webview.onDidReceiveMessage(message => {
				void this.handleMessage(message).catch(error => this.feedback(error instanceof Error ? error.message : 'Action impossible.'));
			}, undefined, this.subscriptions);
		} else this.panel.reveal();
		await this.refresh();
	}

	private refresh(): Promise<void> {
		if (this.refreshing) return this.refreshing;
		this.refreshing = this.updateSnapshot().finally(() => { this.refreshing = undefined; });
		return this.refreshing;
	}

	private async updateSnapshot(): Promise<void> {
		const toolNames = new Set(vscode.lm?.tools?.map(tool => tool.name) ?? []);
		this.snapshot = await collectSetupSnapshot(this.context.extensionPath, {
			extensionVersion: String(this.context.extension.packageJSON.version),
			vscodeVersion: vscode.version, trusted: vscode.workspace.isTrusted,
			aiTools: this.context.extension.packageJSON.contributes.languageModelTools.every((tool: { name: string }) => toolNames.has(tool.name))
		});
		if (this.panel) this.panel.webview.html = renderStartPage(this.snapshot, randomBytes(18).toString('hex'));
	}

	private async feedback(text: string): Promise<void> {
		await this.panel?.webview.postMessage({ type: 'feedback', text });
	}

	private async handleMessage(message: unknown): Promise<void> {
		switch (parseStartPageAction(message)) {
			case 'refresh': await this.refresh(); break;
			case 'copyReport':
				if (this.snapshot) {
					await vscode.env.clipboard.writeText(formatSetupReport(this.snapshot));
					await this.feedback('Diagnostic copié. Aucun contenu de classeur ni chemin de fichier inclus.');
				}
				break;
			case 'openWorkbook':
				if (this.opening) return;
				this.opening = true;
				try { await vscode.commands.executeCommand(EXCEL_AI_COMMANDS.openWorkbook); }
				finally { this.opening = false; }
		}
	}

	dispose(): void {
		this.panel?.dispose();
		this.subscriptions.splice(0).forEach(subscription => subscription.dispose());
	}
}
