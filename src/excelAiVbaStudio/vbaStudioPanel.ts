import { randomBytes } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import {
	assertNoReparsePointChain,
	assertOwnedDirectory,
	pathIsInside
} from './security';
import { ExportContext } from './types';

type ComponentKind =
	| 'document'
	| 'module'
	| 'class'
	| 'userform'
	| 'other'
	| 'reference';

interface ManifestModule {
	name?: string;
	type?: string;
	lineCount?: number;
	file?: string;
	resourceFile?: string;
}

interface VbaManifest {
	status?: string;
	message?: string;
	modules?: ManifestModule[];
	references?: Array<Record<string, unknown>>;
}

interface StudioComponent {
	id: string;
	name: string;
	type: string;
	kind: ComponentKind;
	file?: string;
	source: string;
	editable: boolean;
	properties: Array<{ name: string; value: string }>;
}

interface StudioCategory {
	id: string;
	label: string;
	icon: string;
	components: StudioComponent[];
}

interface StudioProject {
	workbookName: string;
	workbookPath: string;
	projectName: string;
	status: string;
	statusMessage: string;
	embeddedVba: boolean;
	sourceDirectory: string;
	categories: StudioCategory[];
}

interface WebviewMessage {
	type?: unknown;
	file?: unknown;
	source?: unknown;
	kind?: unknown;
}

const MAX_SOURCE_CHARACTERS = 2_000_000;
const SOURCE_EXTENSIONS = new Set(['.bas', '.cls', '.frm', '.txt']);

function asRecord(value: unknown): Record<string, unknown> | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: undefined;
}

function asString(value: unknown): string {
	return value === undefined || value === null ? '' : String(value);
}

function safeFileStem(value: string, fallback: string): string {
	const stem = value
		.normalize('NFKD')
		.replace(/[\u0300-\u036f]/g, '')
		.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-')
		.replace(/\s+/g, '-')
		.replace(/-+/g, '-')
		.replace(/^[.\-\s]+|[.\-\s]+$/g, '')
		.slice(0, 80);
	return stem || fallback;
}

function normalizeType(type: string, file: string): ComponentKind {
	const normalized = type.toLocaleLowerCase('en-US');
	if (normalized.includes('document')) {
		return 'document';
	}
	if (normalized.includes('userform')) {
		return 'userform';
	}
	if (normalized.includes('class')) {
		return 'class';
	}
	if (normalized.includes('standard')) {
		return 'module';
	}
	switch (path.extname(file).toLocaleLowerCase('en-US')) {
		case '.bas':
			return 'module';
		case '.frm':
			return 'userform';
		case '.cls':
			return 'class';
		default:
			return 'other';
	}
}

function componentTemplate(kind: ComponentKind, name: string): string {
	switch (kind) {
		case 'class':
			return [
				'VERSION 1.0 CLASS',
				'BEGIN',
				'  MultiUse = -1  \'True',
				'END',
				`Attribute VB_Name = "${name}"`,
				'Attribute VB_GlobalNameSpace = False',
				'Attribute VB_Creatable = False',
				'Attribute VB_PredeclaredId = False',
				'Attribute VB_Exposed = False',
				'Option Explicit',
				''
			].join('\r\n');
		case 'userform':
			return [
				'VERSION 5.00',
				`Begin VB.UserForm ${name}`,
				`   Caption         =   "${name}"`,
				'   ClientHeight    =   3000',
				'   ClientLeft      =   120',
				'   ClientTop       =   465',
				'   ClientWidth     =   4800',
				'End',
				`Attribute VB_Name = "${name}"`,
				'Attribute VB_GlobalNameSpace = False',
				'Attribute VB_Creatable = False',
				'Attribute VB_PredeclaredId = True',
				'Attribute VB_Exposed = False',
				'Option Explicit',
				''
			].join('\r\n');
		case 'module':
		default:
			return [
				`Attribute VB_Name = "${name}"`,
				'Option Explicit',
				'',
				'Public Sub TestVBA()',
				'    Range("A1").Value = "Test réussi"',
				'    Range("A1").Interior.Color = RGB(0, 176, 80)',
				'    MsgBox "Le code VBA fonctionne !", vbInformation',
				'End Sub',
				''
			].join('\r\n');
	}
}

export class VbaStudioPanel implements vscode.Disposable {
	private readonly disposables: vscode.Disposable[] = [];
	private panel: vscode.WebviewPanel | undefined;
	private currentContext: ExportContext | undefined;
	private componentFiles = new Map<string, string>();
	private savingFile: string | undefined;

	constructor(private readonly outputChannel: vscode.OutputChannel) {
		this.disposables.push(
			vscode.workspace.onDidChangeTextDocument(event => {
				void this.handleExternalDocumentChange(event.document);
			})
		);
	}

	dispose(): void {
		this.panel?.dispose();
		for (const disposable of this.disposables) {
			disposable.dispose();
		}
		this.disposables.length = 0;
	}

	async prepare(context: ExportContext): Promise<void> {
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const manifest = await this.readManifest(context);
		const sourceFiles = await this.listSourceFiles(context);
		if (
			sourceFiles.length === 0 &&
			(manifest.status === 'none' || manifest.status === 'extracted')
		) {
			const starterPath = path.join(context.paths.vbaDirectory, 'Module1.bas');
			await assertNoReparsePointChain(
				starterPath,
				context.paths.vbaDirectory
			);
			await fs.promises.writeFile(
				starterPath,
				componentTemplate('module', 'Module1'),
				{
					encoding: 'utf8',
					flag: 'wx'
				}
			);
		}
	}

	async open(context: ExportContext): Promise<void> {
		this.currentContext = context;
		if (!this.panel) {
			this.panel = vscode.window.createWebviewPanel(
				'excelAiVbaStudio.vbaStudio',
				`VBA Studio · ${path.basename(context.workbookUri.fsPath)}`,
				vscode.ViewColumn.Active,
				{
					enableScripts: true,
					retainContextWhenHidden: true,
					localResourceRoots: []
				}
			);
			this.panel.onDidDispose(
				() => {
					this.panel = undefined;
					this.currentContext = undefined;
					this.componentFiles.clear();
				},
				undefined,
				this.disposables
			);
			// Register the receiver before loading HTML: the webview posts `ready`
			// immediately and fast machines could otherwise lose that first message.
			this.panel.webview.onDidReceiveMessage(
				message => this.handleMessage(message as WebviewMessage),
				undefined,
				this.disposables
			);
			this.panel.webview.html = this.getHtml(this.panel.webview);
		} else {
			this.panel.title = `VBA Studio · ${path.basename(
				context.workbookUri.fsPath
			)}`;
			this.panel.reveal(vscode.ViewColumn.Active, false);
		}
		await this.postProject();
	}

	private async readManifest(context: ExportContext): Promise<VbaManifest> {
		const manifestPath = path.join(context.paths.vbaDirectory, 'manifest.json');
		await assertNoReparsePointChain(
			manifestPath,
			context.paths.vbaDirectory
		);
		try {
			const text = await fs.promises.readFile(manifestPath, 'utf8');
			return JSON.parse(text.replace(/^\uFEFF/, '')) as VbaManifest;
		} catch {
			return {
				status: 'unavailable',
				message: 'Le manifeste VBA n’est pas encore disponible.',
				modules: [],
				references: []
			};
		}
	}

	private async listSourceFiles(context: ExportContext): Promise<string[]> {
		const entries = await fs.promises.readdir(context.paths.vbaDirectory, {
			withFileTypes: true
		});
		return entries
			.filter(
				entry =>
					entry.isFile() &&
					SOURCE_EXTENSIONS.has(
						path.extname(entry.name).toLocaleLowerCase('en-US')
					)
			)
			.map(entry => path.join(context.paths.vbaDirectory, entry.name))
			.filter(file => pathIsInside(file, context.paths.vbaDirectory))
			.sort((left, right) =>
				path.basename(left).localeCompare(path.basename(right))
			);
	}

	private async readWorkbookJson(
		context: ExportContext
	): Promise<Record<string, unknown>> {
		await assertNoReparsePointChain(
			context.paths.jsonPath,
			context.paths.outputDirectory
		);
		try {
			const text = await fs.promises.readFile(context.paths.jsonPath, 'utf8');
			return JSON.parse(text.replace(/^\uFEFF/, '')) as Record<
				string,
				unknown
			>;
		} catch {
			return {};
		}
	}

	private async buildProject(): Promise<StudioProject | undefined> {
		const context = this.currentContext;
		if (!context) {
			return undefined;
		}
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const [manifest, sourceFiles, workbookData] = await Promise.all([
			this.readManifest(context),
			this.listSourceFiles(context),
			this.readWorkbookJson(context)
		]);
		const manifestByFile = new Map(
			(manifest.modules || [])
				.filter(module => module.file)
				.map(module => [
					asString(module.file).toLocaleLowerCase('en-US'),
					module
				])
		);
		const components: StudioComponent[] = [];
		this.componentFiles.clear();
		for (const sourcePath of sourceFiles) {
			await assertNoReparsePointChain(
				sourcePath,
				context.paths.vbaDirectory
			);
			const file = path.basename(sourcePath);
			const record = manifestByFile.get(file.toLocaleLowerCase('en-US'));
			const name =
				asString(record?.name) ||
				path.basename(file, path.extname(file));
			const type = asString(record?.type) || 'Module de travail';
			const source = await fs.promises.readFile(sourcePath, 'utf8');
			const kind = normalizeType(type, file);
			const component: StudioComponent = {
				id: file,
				name,
				type,
				kind,
				file,
				source,
				editable: true,
				properties: [
					{ name: '(Name)', value: name },
					{ name: 'Type', value: type },
					{ name: 'Fichier', value: file },
					{
						name: 'Lignes',
						value: String(source.split(/\r?\n/).length)
					},
					{ name: 'Chemin', value: sourcePath }
				]
			};
			components.push(component);
			this.componentFiles.set(file, sourcePath);
		}

		const workbookRecord = asRecord(workbookData.workbook) || {};
		const worksheets = Array.isArray(workbookData.worksheets)
			? workbookData.worksheets
			: [];
		const documentComponents = components.filter(
			component => component.kind === 'document'
		);
		const virtualDocuments: StudioComponent[] = [];
		const existingNames = new Set(
			documentComponents.map(component =>
				component.name.toLocaleLowerCase('en-US')
			)
		);
		if (!existingNames.has('thisworkbook')) {
			virtualDocuments.push({
				id: 'virtual:ThisWorkbook',
				name: 'ThisWorkbook',
				type: 'Objet Excel',
				kind: 'document',
				source: [
					"' Aucun module ThisWorkbook n’a été extrait du classeur.",
					"' Ajoutez un module standard pour créer du code accessible à Copilot."
				].join('\r\n'),
				editable: false,
				properties: [
					{ name: '(Name)', value: 'ThisWorkbook' },
					{ name: 'Type', value: 'Objet Excel' },
					{ name: 'Classeur', value: context.workbookUri.fsPath }
				]
			});
		}
		for (const [index, rawSheet] of worksheets.entries()) {
			const sheet = asRecord(rawSheet) || {};
			const name =
				asString(sheet.name || sheet.sheetName) || `Feuille${index + 1}`;
			if (existingNames.has(name.toLocaleLowerCase('en-US'))) {
				continue;
			}
			virtualDocuments.push({
				id: `virtual:sheet:${index}`,
				name,
				type: 'Worksheet',
				kind: 'document',
				source: `' Aucun code de feuille n’a été extrait pour ${name}.`,
				editable: false,
				properties: [
					{ name: '(Name)', value: name },
					{ name: 'Type', value: 'Worksheet' },
					{
						name: 'Plage utilisée',
						value: asString(
							sheet.range ||
								sheet.address ||
								asRecord(sheet.usedRange)?.address
						)
					}
				].filter(property => property.value)
			});
		}

		const references = (manifest.references || []).map(
			(reference, index): StudioComponent => {
				const name = asString(reference.name) || `Référence ${index + 1}`;
				return {
					id: `reference:${index}`,
					name,
					type: 'Référence',
					kind: 'reference',
					source: asString(reference.fullPath || reference.description),
					editable: false,
					properties: Object.entries(reference)
						.filter(([, value]) => value !== undefined && value !== null)
						.map(([key, value]) => ({
							name: key,
							value: String(value)
						}))
				};
			}
		);
		const categories: StudioCategory[] = [
			{
				id: 'objects',
				label: 'Microsoft Excel Objects',
				icon: '▦',
				components: [...documentComponents, ...virtualDocuments]
			},
			{
				id: 'forms',
				label: 'UserForms',
				icon: '▣',
				components: components.filter(
					component => component.kind === 'userform'
				)
			},
			{
				id: 'modules',
				label: 'Modules',
				icon: '◆',
				components: components.filter(
					component => component.kind === 'module'
				)
			},
			{
				id: 'classes',
				label: 'Modules de classe',
				icon: '◇',
				components: components.filter(
					component => component.kind === 'class'
				)
			},
			{
				id: 'other',
				label: 'Autres composants',
				icon: '◫',
				components: components.filter(
					component => component.kind === 'other'
				)
			},
			{
				id: 'references',
				label: 'Références',
				icon: '◉',
				components: references
			}
		].filter(category => category.components.length > 0);

		const workbookName =
			asString(workbookRecord.name) ||
			path.basename(context.workbookUri.fsPath);
		return {
			workbookName,
			workbookPath: context.workbookUri.fsPath,
			projectName: `VBAProject (${workbookName})`,
			status: asString(manifest.status) || 'inconnu',
			statusMessage: asString(manifest.message),
			embeddedVba: manifest.status === 'extracted',
			sourceDirectory: context.paths.vbaDirectory,
			categories
		};
	}

	private async postProject(selectFile?: string): Promise<void> {
		if (!this.panel) {
			return;
		}
		const project = await this.buildProject();
		if (project) {
			await this.panel.webview.postMessage({
				type: 'project',
				project,
				selectFile
			});
		}
	}

	private async handleMessage(message: WebviewMessage): Promise<void> {
		const type = typeof message.type === 'string' ? message.type : '';
		switch (type) {
			case 'ready':
			case 'reload':
				await this.postProject();
				return;
			case 'save':
				await this.saveSource(message.file, message.source);
				return;
			case 'create':
				await this.createComponent(message.kind);
				return;
			case 'openFile':
				await this.openSourceFile(message.file);
				return;
			case 'askCopilot':
				await this.askCopilot();
				return;
		}
	}

	private async saveSource(fileValue: unknown, sourceValue: unknown): Promise<void> {
		const file = typeof fileValue === 'string' ? fileValue : '';
		const source = typeof sourceValue === 'string' ? sourceValue : '';
		const sourcePath = this.componentFiles.get(file);
		if (!sourcePath || !this.currentContext) {
			await this.postStatus('error', 'Ce composant ne peut pas être enregistré.');
			return;
		}
		if (source.length > MAX_SOURCE_CHARACTERS) {
			await this.postStatus(
				'error',
				'Le module dépasse la limite de 2 000 000 de caractères.'
			);
			return;
		}
		await assertNoReparsePointChain(
			sourcePath,
			this.currentContext.paths.vbaDirectory
		);
		if (!pathIsInside(sourcePath, this.currentContext.paths.vbaDirectory)) {
			throw new Error('Le module demandé sort du projet VBA.');
		}
		const uri = vscode.Uri.file(sourcePath);
		const document = await vscode.workspace.openTextDocument(uri);
		if (document.isDirty && document.getText() !== source) {
			await this.postStatus(
				'error',
				'Ce fichier contient déjà des modifications non enregistrées dans VS Code. Enregistrez-les ou rechargez le studio.'
			);
			return;
		}
		const lastLine = document.lineAt(Math.max(0, document.lineCount - 1));
		const edit = new vscode.WorkspaceEdit();
		edit.replace(
			uri,
			new vscode.Range(0, 0, lastLine.lineNumber, lastLine.text.length),
			source
		);
		this.savingFile = file;
		try {
			const applied = await vscode.workspace.applyEdit(edit);
			if (!applied || !(await document.save())) {
				throw new Error('VS Code a refusé l’enregistrement du module.');
			}
			await this.postStatus('saved', `${file} enregistré — Copilot voit ce fichier.`);
		} finally {
			this.savingFile = undefined;
		}
	}

	private async createComponent(kindValue: unknown): Promise<void> {
		const context = this.currentContext;
		if (!context) {
			return;
		}
		const kind =
			kindValue === 'class' || kindValue === 'userform'
				? kindValue
				: 'module';
		const settings: Record<
			'module' | 'class' | 'userform',
			{ prefix: string; extension: string }
		> = {
			module: { prefix: 'Module', extension: '.bas' },
			class: { prefix: 'Classe', extension: '.cls' },
			userform: { prefix: 'UserForm', extension: '.frm' }
		};
		const setting = settings[kind];
		const existing = new Set(
			(await this.listSourceFiles(context)).map(file =>
				path.basename(file).toLocaleLowerCase('en-US')
			)
		);
		let index = 1;
		let name = `${setting.prefix}${index}`;
		let file = `${safeFileStem(name, setting.prefix)}${setting.extension}`;
		while (existing.has(file.toLocaleLowerCase('en-US'))) {
			index++;
			name = `${setting.prefix}${index}`;
			file = `${safeFileStem(name, setting.prefix)}${setting.extension}`;
		}
		const sourcePath = path.join(context.paths.vbaDirectory, file);
		await assertNoReparsePointChain(sourcePath, context.paths.vbaDirectory);
		await fs.promises.writeFile(
			sourcePath,
			componentTemplate(kind, name),
			{
				encoding: 'utf8',
				flag: 'wx'
			}
		);
		await this.postProject(file);
		await this.postStatus('saved', `${file} créé — accessible à Copilot.`);
	}

	private async openSourceFile(fileValue: unknown): Promise<void> {
		const file = typeof fileValue === 'string' ? fileValue : '';
		const sourcePath = this.componentFiles.get(file);
		if (!sourcePath) {
			return;
		}
		const document = await vscode.workspace.openTextDocument(sourcePath);
		await vscode.window.showTextDocument(document, {
			preview: false,
			preserveFocus: false,
			viewColumn: vscode.ViewColumn.Beside
		});
	}

	private async askCopilot(): Promise<void> {
		const context = this.currentContext;
		if (!context) {
			return;
		}
		const prompt = [
			'Analyse le projet VBA actuellement ouvert dans Excel AI & VBA Studio.',
			`Classeur : ${context.workbookUri.fsPath}`,
			`Sources VBA : ${context.paths.vbaDirectory}`,
			'Lis les fichiers .bas, .cls et .frm de cette racine VS Code.',
			'Utilise aussi #excelVbaWorkbook avec includeVba: true pour le contexte du classeur.',
			'N’exécute aucune macro. Modifie uniquement les fichiers source lorsque je te le demande.'
		].join('\n');
		try {
			await vscode.commands.executeCommand('workbench.action.chat.open', {
				query: prompt
			});
		} catch {
			await vscode.env.clipboard.writeText(prompt);
			await vscode.commands.executeCommand('workbench.action.chat.open');
		}
	}

	private async handleExternalDocumentChange(
		document: vscode.TextDocument
	): Promise<void> {
		const context = this.currentContext;
		if (
			!this.panel ||
			!context ||
			document.uri.scheme !== 'file' ||
			!pathIsInside(document.uri.fsPath, context.paths.vbaDirectory)
		) {
			return;
		}
		const file = path.basename(document.uri.fsPath);
		if (!this.componentFiles.has(file) || this.savingFile === file) {
			return;
		}
		await this.panel.webview.postMessage({
			type: 'sourceChanged',
			file,
			source: document.getText()
		});
	}

	private async postStatus(
		status: 'saved' | 'error' | 'info',
		message: string
	): Promise<void> {
		await this.panel?.webview.postMessage({
			type: 'status',
			status,
			message
		});
	}

	private getHtml(webview: vscode.Webview): string {
		const nonce = randomBytes(18).toString('base64');
		return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
<title>VBA Studio</title>
<style nonce="${nonce}">
:root{color-scheme:light dark}
*{box-sizing:border-box}
html,body{height:100%;margin:0;overflow:hidden;font-family:var(--vscode-font-family);font-size:13px;color:var(--vscode-foreground);background:var(--vscode-editor-background)}
button,select,textarea{font:inherit}
.app{height:100%;display:grid;grid-template-rows:38px 1fr 24px}
.toolbar{display:flex;align-items:center;gap:4px;padding:4px 7px;border-bottom:1px solid var(--vscode-panel-border);background:var(--vscode-editorGroupHeader-tabsBackground)}
.brand{display:flex;align-items:center;gap:7px;font-weight:650;margin-right:8px;white-space:nowrap}
.brand-icon{width:22px;height:22px;display:grid;place-items:center;color:white;background:#107c41;border-radius:3px}
.tool{height:27px;padding:0 9px;border:1px solid transparent;color:var(--vscode-foreground);background:transparent;border-radius:3px;cursor:pointer}
.tool:hover{background:var(--vscode-toolbar-hoverBackground);border-color:var(--vscode-panel-border)}
.tool:disabled{opacity:.45;cursor:default}
.separator{height:20px;border-left:1px solid var(--vscode-panel-border);margin:0 3px}
.spacer{flex:1}
.ai{color:var(--vscode-button-foreground);background:var(--vscode-button-background);border-color:var(--vscode-button-border,transparent)}
.ai:hover{background:var(--vscode-button-hoverBackground)}
.workspace{min-height:0;display:grid;grid-template-columns:minmax(260px,32%) 1fr}
.left{min-width:0;min-height:0;display:grid;grid-template-rows:minmax(180px,58%) minmax(130px,42%);border-right:1px solid var(--vscode-panel-border)}
.pane{min-height:0;display:flex;flex-direction:column}
.pane+.pane{border-top:1px solid var(--vscode-panel-border)}
.pane-title{height:27px;display:flex;align-items:center;padding:0 8px;font-weight:650;background:var(--vscode-sideBarSectionHeader-background);border-bottom:1px solid var(--vscode-panel-border)}
.pane-body{min-height:0;overflow:auto;padding:4px 0}
.project-root,.category,.component{display:flex;align-items:center;height:23px;gap:5px;white-space:nowrap;cursor:default}
.project-root{padding-left:7px;font-weight:600}
.category{padding-left:20px}
.component{padding-left:42px}
.component:hover,.component.selected{background:var(--vscode-list-hoverBackground);color:var(--vscode-list-activeSelectionForeground)}
.component.selected{background:var(--vscode-list-activeSelectionBackground)}
.twisty{width:12px;text-align:center;opacity:.8}
.node-icon{width:17px;text-align:center;color:var(--vscode-symbolIcon-methodForeground,#b180d7)}
.node-name{overflow:hidden;text-overflow:ellipsis}
.count{margin-left:auto;padding-right:8px;opacity:.55}
.props{width:100%;border-collapse:collapse;font-size:12px}
.props td{padding:4px 6px;border-bottom:1px solid var(--vscode-panel-border);vertical-align:top}
.props td:first-child{width:38%;font-weight:600;background:var(--vscode-editorWidget-background)}
.props td:last-child{word-break:break-word}
.editor-area{min-width:0;min-height:0;display:grid;grid-template-rows:30px auto 1fr}
.editor-head{display:grid;grid-template-columns:1fr 1fr;border-bottom:1px solid var(--vscode-panel-border);background:var(--vscode-editorGroupHeader-tabsBackground)}
.editor-head select{min-width:0;border:0;border-right:1px solid var(--vscode-panel-border);padding:0 8px;color:var(--vscode-dropdown-foreground);background:var(--vscode-dropdown-background)}
.notice{display:none;padding:6px 10px;border-bottom:1px solid var(--vscode-panel-border);background:var(--vscode-inputValidation-warningBackground);color:var(--vscode-inputValidation-warningForeground)}
.notice.visible{display:block}
.code-wrap{min-height:0;position:relative;display:grid;grid-template-columns:48px 1fr;background:var(--vscode-editor-background)}
.lines{margin:0;padding:10px 8px 10px 0;overflow:hidden;text-align:right;white-space:pre;color:var(--vscode-editorLineNumber-foreground);background:var(--vscode-editorGutter-background);border-right:1px solid var(--vscode-panel-border);font-family:var(--vscode-editor-font-family);font-size:var(--vscode-editor-font-size);line-height:var(--vscode-editor-line-height)}
.code{width:100%;height:100%;resize:none;border:0;outline:0;padding:10px 12px;tab-size:4;white-space:pre;overflow:auto;color:var(--vscode-editor-foreground);background:var(--vscode-editor-background);font-family:var(--vscode-editor-font-family);font-size:var(--vscode-editor-font-size);line-height:var(--vscode-editor-line-height)}
.code:disabled{opacity:.72}
.statusbar{display:flex;align-items:center;gap:16px;padding:0 8px;color:var(--vscode-statusBar-foreground);background:var(--vscode-statusBar-background)}
.statusbar .right{margin-left:auto}
.status-ok{color:var(--vscode-testing-iconPassed)}
.status-error{color:var(--vscode-errorForeground)}
@media(max-width:760px){.workspace{grid-template-columns:230px 1fr}.tool .label{display:none}}
</style>
</head>
<body>
<div class="app">
  <div class="toolbar">
    <div class="brand"><span class="brand-icon">VB</span><span>VBA Studio</span></div>
    <button class="tool" id="save" title="Enregistrer (Ctrl+S)">💾 <span class="label">Enregistrer</span></button>
    <span class="separator"></span>
    <button class="tool create" data-kind="module">＋ Module</button>
    <button class="tool create" data-kind="class">＋ Classe</button>
    <button class="tool create" data-kind="userform">＋ UserForm</button>
    <span class="separator"></span>
    <button class="tool" id="open-file">↗ <span class="label">Éditeur VS Code</span></button>
    <button class="tool" id="reload">⟳ <span class="label">Recharger</span></button>
    <div class="spacer"></div>
    <button class="tool ai" id="copilot">✦ Analyser avec Copilot</button>
  </div>
  <div class="workspace">
    <aside class="left">
      <section class="pane">
        <div class="pane-title" id="project-title">Projet - VBAProject</div>
        <div class="pane-body" id="tree"></div>
      </section>
      <section class="pane">
        <div class="pane-title" id="properties-title">Propriétés</div>
        <div class="pane-body"><table class="props"><tbody id="properties"></tbody></table></div>
      </section>
    </aside>
    <main class="editor-area">
      <div class="editor-head">
        <select id="object-select" aria-label="Objet"></select>
        <select id="procedure-select" aria-label="Procédure"><option>(Général)</option></select>
      </div>
      <div class="notice" id="notice"></div>
      <div class="code-wrap">
        <pre class="lines" id="lines">1</pre>
        <textarea class="code" id="code" spellcheck="false" aria-label="Code VBA"></textarea>
      </div>
    </main>
  </div>
  <div class="statusbar">
    <span id="status">Prêt</span>
    <span id="position">Ln 1, Col 1</span>
    <span class="right" id="ai-status">Copilot : sources réelles du workspace</span>
  </div>
</div>
<script nonce="${nonce}">
(() => {
  const vscode = acquireVsCodeApi();
  const state = { project: null, selected: null, dirty: false };
  const tree = document.getElementById('tree');
  const properties = document.getElementById('properties');
  const code = document.getElementById('code');
  const lines = document.getElementById('lines');
  const status = document.getElementById('status');
  const position = document.getElementById('position');
  const notice = document.getElementById('notice');
  const objectSelect = document.getElementById('object-select');
  const procedureSelect = document.getElementById('procedure-select');
  const openFile = document.getElementById('open-file');
  const save = document.getElementById('save');

  const allComponents = () => state.project
    ? state.project.categories.flatMap(category => category.components)
    : [];

  function setStatus(message, kind = 'info') {
    status.textContent = message;
    status.className = kind === 'saved' ? 'status-ok' : kind === 'error' ? 'status-error' : '';
  }

  function showNotice(message) {
    notice.textContent = message || '';
    notice.classList.toggle('visible', Boolean(message));
  }

  function updateLines() {
    const count = Math.max(1, code.value.split('\\n').length);
    lines.textContent = Array.from({ length: count }, (_, index) => index + 1).join('\\n');
    lines.scrollTop = code.scrollTop;
  }

  function updatePosition() {
    const before = code.value.slice(0, code.selectionStart);
    const parts = before.split('\\n');
    position.textContent = 'Ln ' + parts.length + ', Col ' + (parts[parts.length - 1].length + 1);
  }

  function renderProcedures() {
    const selectedValue = procedureSelect.value;
    const procedures = [];
    const matcher = /^\\s*(?:Public\\s+|Private\\s+|Friend\\s+|Static\\s+)?(?:Sub|Function|Property\\s+(?:Get|Let|Set))\\s+([A-Za-z_]\\w*)/gim;
    let match;
    while ((match = matcher.exec(code.value))) procedures.push(match[1]);
    procedureSelect.innerHTML = '<option value="">(Général)</option>' +
      procedures.map(name => '<option>' + escapeText(name) + '</option>').join('');
    if (procedures.includes(selectedValue)) procedureSelect.value = selectedValue;
  }

  function escapeText(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function selectComponent(id) {
    const component = allComponents().find(item => item.id === id);
    if (!component) return;
    state.selected = component;
    state.dirty = false;
    code.value = component.source || '';
    code.disabled = !component.editable;
    save.disabled = !component.editable;
    openFile.disabled = !component.file;
    objectSelect.value = component.id;
    document.getElementById('properties-title').textContent = 'Propriétés - ' + component.name;
    properties.innerHTML = component.properties.map(property =>
      '<tr><td>' + escapeText(property.name) + '</td><td>' + escapeText(property.value) + '</td></tr>'
    ).join('');
    document.querySelectorAll('.component').forEach(element =>
      element.classList.toggle('selected', element.dataset.id === id)
    );
    showNotice(component.editable ? '' : 'Cet objet est affiché à titre de contexte. Ajoutez un module pour écrire du code.');
    renderProcedures();
    updateLines();
    updatePosition();
    setStatus(component.file ? component.file : component.type);
  }

  function renderProject(project, selectFile) {
    state.project = project;
    document.getElementById('project-title').textContent = 'Projet - ' + project.projectName;
    document.getElementById('ai-status').textContent = 'Copilot : ' + project.sourceDirectory;
    tree.innerHTML = '<div class="project-root"><span class="twisty">▾</span><span class="node-icon">▤</span><span class="node-name">' +
      escapeText(project.projectName) + '</span></div>' +
      project.categories.map(category =>
        '<div class="category"><span class="twisty">▾</span><span class="node-icon">' + category.icon +
        '</span><span class="node-name">' + escapeText(category.label) + '</span><span class="count">' +
        category.components.length + '</span></div>' +
        category.components.map(component =>
          '<div class="component" data-id="' + escapeText(component.id) + '"><span class="twisty"></span><span class="node-icon">' +
          (component.kind === 'userform' ? '▣' : component.kind === 'module' ? '◆' : component.kind === 'class' ? '◇' : '▦') +
          '</span><span class="node-name">' + escapeText(component.name) + '</span></div>'
        ).join('')
      ).join('');
    objectSelect.innerHTML = allComponents().map(component =>
      '<option value="' + escapeText(component.id) + '">' + escapeText(component.name) + '</option>'
    ).join('');
    tree.querySelectorAll('.component').forEach(element =>
      element.addEventListener('click', () => selectComponent(element.dataset.id))
    );
    objectSelect.onchange = () => selectComponent(objectSelect.value);
    const target = selectFile && allComponents().find(component => component.file === selectFile);
    const firstEditable = allComponents().find(component => component.editable);
    const first = target || firstEditable || allComponents()[0];
    if (first) selectComponent(first.id);
    showNotice(project.embeddedVba
      ? ''
      : 'Copie de travail VS Code : ce classeur ne contient pas encore de projet VBA incorporé. Les fichiers restent accessibles à Copilot.');
  }

  code.addEventListener('input', () => {
    state.dirty = true;
    if (state.selected) state.selected.source = code.value;
    updateLines();
    updatePosition();
    renderProcedures();
    setStatus('Modifié — Ctrl+S');
  });
  code.addEventListener('scroll', () => { lines.scrollTop = code.scrollTop; });
  code.addEventListener('click', updatePosition);
  code.addEventListener('keyup', updatePosition);
  code.addEventListener('keydown', event => {
    if (event.key === 'Tab') {
      event.preventDefault();
      const start = code.selectionStart;
      code.setRangeText('    ', start, code.selectionEnd, 'end');
      code.dispatchEvent(new Event('input'));
    }
  });
  procedureSelect.onchange = () => {
    if (!procedureSelect.value) return;
    const matcher = new RegExp('^\\\\s*(?:Public\\\\s+|Private\\\\s+|Friend\\\\s+|Static\\\\s+)?(?:Sub|Function|Property\\\\s+(?:Get|Let|Set))\\\\s+' + procedureSelect.value + '\\\\b', 'im');
    const match = matcher.exec(code.value);
    if (match) {
      code.focus();
      code.setSelectionRange(match.index, match.index + match[0].length);
      updatePosition();
    }
  };
  save.onclick = () => {
    if (!state.selected || !state.selected.file || !state.selected.editable) return;
    vscode.postMessage({ type: 'save', file: state.selected.file, source: code.value });
  };
  openFile.onclick = () => state.selected?.file &&
    vscode.postMessage({ type: 'openFile', file: state.selected.file });
  document.getElementById('reload').onclick = () => vscode.postMessage({ type: 'reload' });
  document.getElementById('copilot').onclick = () => vscode.postMessage({ type: 'askCopilot' });
  document.querySelectorAll('.create').forEach(button =>
    button.addEventListener('click', () => vscode.postMessage({ type: 'create', kind: button.dataset.kind }))
  );
  window.addEventListener('keydown', event => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
      event.preventDefault();
      save.click();
    }
  });
  window.addEventListener('message', event => {
    const message = event.data;
    if (message.type === 'project') {
      renderProject(message.project, message.selectFile);
    } else if (message.type === 'status') {
      if (message.status === 'saved') state.dirty = false;
      setStatus(message.message, message.status);
    } else if (message.type === 'sourceChanged' && state.selected?.file === message.file) {
      if (state.dirty) {
        showNotice('Copilot ou l’éditeur VS Code a aussi modifié ce fichier. Enregistrez votre version ou cliquez sur Recharger.');
      } else {
        state.selected.source = message.source;
        code.value = message.source;
        updateLines();
        renderProcedures();
        setStatus(message.file + ' actualisé par VS Code');
      }
    }
  });
  vscode.postMessage({ type: 'ready' });
})();
</script>
</body>
</html>`;
	}
}
