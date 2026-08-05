import { randomBytes } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { errorMessage } from '../common/errorMessage';
import {
	assertNoReparsePointChain,
	assertOwnedDirectory,
	pathIsInside
} from './security';
import {
	ExportContext,
	VbaDesignOperation,
	VbaDesignToolResult,
	VbaUserFormControl,
	VbaUserFormControlChanges,
	VbaUserFormControlType
} from './types';
import {
	buildVbaInteractionGraph,
	VbaInteractionGraph
} from './vbaInteractionGraph';
import { VbaWritebackService } from './vbaWritebackService';

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

interface StudioUserFormControl {
	type: string;
	typeName: string;
	name: string;
	caption: string;
	left: number;
	top: number;
	width: number;
	height: number;
	enabled: boolean;
	visible: boolean;
	tabIndex?: number;
	controlTipText: string;
}

interface StudioUserForm {
	name: string;
	caption: string;
	width: number;
	height: number;
	controls: StudioUserFormControl[];
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
	userForms: StudioUserForm[];
	worksheetNames: string[];
	interactions: VbaInteractionGraph;
}

interface WebviewMessage {
	type?: unknown;
	file?: unknown;
	source?: unknown;
	kind?: unknown;
	sheetName?: unknown;
	controlName?: unknown;
	macroName?: unknown;
	formName?: unknown;
	objectName?: unknown;
	eventName?: unknown;
	procedureSource?: unknown;
	replaceExisting?: unknown;
	changes?: unknown;
	control?: unknown;
}

type DesignWorkbookHandler = (
	workbookUri: vscode.Uri,
	operations: VbaDesignOperation[]
) => Promise<VbaDesignToolResult>;

type OpenExcelHandler = (
	workbookUri: vscode.Uri,
	showVbe: boolean
) => Promise<void>;

const MAX_SOURCE_CHARACTERS = 2_000_000;
const MAX_EVENT_SOURCE_CHARACTERS = 200_000;
const VBA_IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_]{0,30}$/;
const SOURCE_EXTENSIONS = new Set(['.bas', '.cls', '.frm', '.txt']);
const VISUAL_CONTROL_TYPES = new Set<VbaUserFormControlType>([
	'label',
	'textBox',
	'commandButton',
	'comboBox',
	'listBox',
	'checkBox',
	'optionButton',
	'toggleButton',
	'frame',
	'image',
	'spinButton',
	'scrollBar'
]);

function asRecord(value: unknown): Record<string, unknown> | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: undefined;
}

function asString(value: unknown): string {
	return value === undefined || value === null ? '' : String(value);
}

function asFiniteNumber(value: unknown, fallback = 0): number {
	return typeof value === 'number' && Number.isFinite(value)
		? value
		: fallback;
}

function asBoolean(value: unknown, fallback: boolean): boolean {
	return typeof value === 'boolean' ? value : fallback;
}

function designerIdentifier(value: unknown, label: string): string {
	const identifier = typeof value === 'string' ? value : '';
	if (!VBA_IDENTIFIER_PATTERN.test(identifier)) {
		throw new Error(`${label} n’est pas un identifiant VBA valide.`);
	}
	return identifier;
}

function designerNumber(
	value: unknown,
	label: string,
	options: { minimum: number; exclusiveMinimum?: boolean }
): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw new Error(`${label} doit être un nombre fini.`);
	}
	if (
		(options.exclusiveMinimum
			? value <= options.minimum
			: value < options.minimum) ||
		value > 10_000
	) {
		throw new Error(`${label} sort des limites autorisées.`);
	}
	return value;
}

function designerText(value: unknown, label: string): string {
	if (typeof value !== 'string' || value.length > 1_000 || value.includes('\0')) {
		throw new Error(`${label} doit être un texte valide de 1 000 caractères maximum.`);
	}
	return value;
}

function validateEventProcedureSource(
	source: unknown,
	objectName: string,
	eventName: string
): string {
	if (
		typeof source !== 'string' ||
		source.length < 1 ||
		source.length > MAX_EVENT_SOURCE_CHARACTERS ||
		source.includes('\0')
	) {
		throw new Error('Le gestionnaire VBA est vide ou dépasse la limite autorisée.');
	}
	const procedureName = `${objectName}_${eventName}`;
	const escapedName = procedureName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const lines = source
		.replace(/\r\n?/g, '\n')
		.split('\n')
		.map(line => line.replace(/[ \t]+$/g, ''));
	while (lines.length && !lines[0].trim()) {
		lines.shift();
	}
	while (lines.length && !lines[lines.length - 1].trim()) {
		lines.pop();
	}
	const headerPattern = new RegExp(
		`^[\\t ]*Private[\\t ]+Sub[\\t ]+${escapedName}[\\t ]*\\([^\\r\\n]*\\)[\\t ]*$`,
		'i'
	);
	if (
		lines.length < 2 ||
		!headerPattern.test(lines[0]) ||
		!/^[\t ]*End[\t ]+Sub[\t ]*$/i.test(lines[lines.length - 1]) ||
		lines
			.slice(1, -1)
			.some(line =>
				/^[\t ]*(?:(?:Private|Public|Friend|Static)[\t ]+)?(?:Sub|Function|Property[\t ]+(?:Get|Let|Set))\b/i.test(
					line
				) || /^[\t ]*End[\t ]+Sub\b/i.test(line)
			)
	) {
		throw new Error(
			`Le code doit contenir uniquement Private Sub ${procedureName}(...) ... End Sub.`
		);
	}
	return source;
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

	constructor(
		private readonly outputChannel: vscode.OutputChannel,
		private readonly writebackService: VbaWritebackService,
		private readonly designWorkbook: DesignWorkbookHandler,
		private readonly openExcel: OpenExcelHandler
	) {
		this.disposables.push(
			vscode.workspace.onDidChangeTextDocument(event => {
				this.runInBackground(
					'Synchronisation de l’éditeur VBA',
					this.handleExternalDocumentChange(event.document)
				);
			}),
			vscode.workspace.onDidSaveTextDocument(document => {
				this.runInBackground(
					'Réinjection automatique VBA',
					this.handleExternalDocumentSave(document)
				);
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
		await this.writebackService.prepare(context);
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
		const worksheetNames: string[] = [];
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
			worksheetNames.push(name);
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
		const vbaRecord = asRecord(workbookRecord.vba) || {};
		const userForms: StudioUserForm[] = (
			Array.isArray(vbaRecord.userForms) ? vbaRecord.userForms : []
		)
			.slice(0, 200)
			.map(rawForm => {
				const form = asRecord(rawForm) || {};
				const controls: StudioUserFormControl[] = (
					Array.isArray(form.controls) ? form.controls : []
				)
					.slice(0, 1_000)
					.map(rawControl => {
						const control = asRecord(rawControl) || {};
						const tabIndex = asFiniteNumber(control.tabIndex, -1);
						return {
							type: asString(control.type) || 'customActiveX',
							typeName: asString(control.typeName),
							name: asString(control.name),
							caption: asString(control.caption),
							left: asFiniteNumber(control.left),
							top: asFiniteNumber(control.top),
							width: Math.max(1, asFiniteNumber(control.width, 72)),
							height: Math.max(1, asFiniteNumber(control.height, 24)),
							enabled: asBoolean(control.enabled, true),
							visible: asBoolean(control.visible, true),
							...(Number.isInteger(tabIndex) && tabIndex >= 0
								? { tabIndex }
								: {}),
							controlTipText: asString(control.controlTipText)
						};
					})
					.filter(control => control.name);
				return {
					name: asString(form.name),
					caption: asString(form.caption),
					width: Math.max(120, asFiniteNumber(form.width, 400)),
					height: Math.max(80, asFiniteNumber(form.height, 300)),
					controls
				};
			})
			.filter(form => form.name);
		const interactions = buildVbaInteractionGraph(
			components.map(component => ({
				name: component.name,
				type: component.type,
				source: component.source
			})),
			workbookRecord.vba
		);
		return {
			workbookName,
			workbookPath: context.workbookUri.fsPath,
			projectName: `VBAProject (${workbookName})`,
			status: asString(manifest.status) || 'inconnu',
			statusMessage: asString(manifest.message),
			embeddedVba: manifest.status === 'extracted',
			sourceDirectory: context.paths.vbaDirectory,
			categories,
			userForms,
			worksheetNames,
			interactions
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
			case 'assignFormButton':
				await this.assignMacro(message, false);
				return;
			case 'bindActiveX':
				await this.assignMacro(message, true);
				return;
			case 'createActiveX':
				await this.createWorksheetActiveXControl();
				return;
			case 'addVisualControl':
				await this.addVisualControl(message);
				return;
			case 'updateVisualControl':
				await this.updateVisualControl(message);
				return;
			case 'setUserFormEvent':
				await this.setUserFormEvent(message);
				return;
			case 'openExcel':
				if (this.currentContext) {
					await this.openExcel(this.currentContext.workbookUri, false);
				}
				return;
			case 'openVbe':
				if (this.currentContext) {
					await this.openExcel(this.currentContext.workbookUri, true);
				}
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
		} catch (error) {
			await this.postStatus(
				'error',
				`Enregistrement du fichier refusé : ${(error as Error).message}`
			);
			this.savingFile = undefined;
			return;
		}
		try {
			const result = await this.writebackService.applySource(
				this.currentContext,
				file,
				source
			);
			await this.postStatus(
				'saved',
				result.changed
					? `${file} enregistré et réinjecté dans le classeur.`
					: `${file} déjà synchronisé avec le classeur.`
			);
		} catch (error) {
			await this.postStatus(
				'error',
				`Fichier enregistré, mais réinjection refusée : ${(error as Error).message}`
			);
		} finally {
			this.savingFile = undefined;
		}
	}

	private async applyDesignerOperations(
		operations: VbaDesignOperation[],
		successMessage: string
	): Promise<void> {
		const context = this.currentContext;
		if (!context) {
			return;
		}
		try {
			await this.postStatus('info', 'Transaction VBA Designer en cours…');
			await this.designWorkbook(context.workbookUri, operations);
			await this.postProject();
			await this.postStatus('saved', successMessage);
		} catch (error) {
			await this.postStatus(
				'error',
				`Transaction refusée : ${(error as Error).message}`
			);
		}
	}

	private async addVisualControl(message: WebviewMessage): Promise<void> {
		try {
			const formName = designerIdentifier(message.formName, 'formName');
			const raw = asRecord(message.control);
			if (!raw) {
				throw new Error('Le contrôle visuel est absent.');
			}
			const type = asString(raw.type) as VbaUserFormControlType;
			if (!VISUAL_CONTROL_TYPES.has(type)) {
				throw new Error('Ce type de contrôle n’est pas disponible dans la palette.');
			}
			const control: VbaUserFormControl = {
				type,
				name: designerIdentifier(raw.name, 'control.name'),
				left: designerNumber(raw.left, 'control.left', { minimum: 0 }),
				top: designerNumber(raw.top, 'control.top', { minimum: 0 }),
				width: designerNumber(raw.width, 'control.width', {
					minimum: 0,
					exclusiveMinimum: true
				}),
				height: designerNumber(raw.height, 'control.height', {
					minimum: 0,
					exclusiveMinimum: true
				}),
				...(raw.caption !== undefined
					? { caption: designerText(raw.caption, 'control.caption') }
					: {})
			};
			await this.applyDesignerOperations(
				[{ kind: 'addUserFormControl', formName, control }],
				`Contrôle ${formName}.${control.name} créé et vérifié.`
			);
		} catch (error) {
			await this.postStatus('error', (error as Error).message);
		}
	}

	private async updateVisualControl(message: WebviewMessage): Promise<void> {
		try {
			const formName = designerIdentifier(message.formName, 'formName');
			const name = designerIdentifier(message.controlName, 'controlName');
			const raw = asRecord(message.changes);
			if (!raw) {
				throw new Error('Les propriétés à modifier sont absentes.');
			}
			const changes: VbaUserFormControlChanges = {};
			for (const property of ['left', 'top', 'width', 'height'] as const) {
				if (raw[property] === undefined) {
					continue;
				}
				changes[property] = designerNumber(
					raw[property],
					`changes.${property}`,
					{
						minimum: 0,
						exclusiveMinimum: property === 'width' || property === 'height'
					}
				);
			}
			for (const property of ['caption', 'controlTipText'] as const) {
				if (raw[property] !== undefined) {
					changes[property] = designerText(
						raw[property],
						`changes.${property}`
					);
				}
			}
			for (const property of ['enabled', 'visible'] as const) {
				if (raw[property] === undefined) {
					continue;
				}
				if (typeof raw[property] !== 'boolean') {
					throw new Error(`changes.${property} doit être un booléen.`);
				}
				changes[property] = raw[property] as boolean;
			}
			if (raw.tabIndex !== undefined) {
				const tabIndex = designerNumber(raw.tabIndex, 'changes.tabIndex', {
					minimum: 0
				});
				if (!Number.isInteger(tabIndex) || tabIndex > 32_767) {
					throw new Error('changes.tabIndex doit être un entier entre 0 et 32767.');
				}
				changes.tabIndex = tabIndex;
			}
			if (Object.keys(changes).length === 0) {
				throw new Error('Aucune propriété de contrôle à modifier.');
			}
			await this.applyDesignerOperations(
				[{ kind: 'updateUserFormControl', formName, name, changes }],
				`Contrôle ${formName}.${name} mis à jour et vérifié.`
			);
		} catch (error) {
			await this.postStatus('error', (error as Error).message);
		}
	}

	private async setUserFormEvent(message: WebviewMessage): Promise<void> {
		try {
			const formName = designerIdentifier(message.formName, 'formName');
			const objectName = designerIdentifier(message.objectName, 'objectName');
			const eventName = designerIdentifier(message.eventName, 'eventName');
			const procedureSource = validateEventProcedureSource(
				message.procedureSource,
				objectName,
				eventName
			);
			const replaceExisting = message.replaceExisting === true;
			if (message.replaceExisting !== undefined && typeof message.replaceExisting !== 'boolean') {
				throw new Error('replaceExisting doit être un booléen.');
			}
			if (replaceExisting) {
				const confirmation = await vscode.window.showWarningMessage(
					`Remplacer le gestionnaire ${formName}.${objectName}_${eventName} ? Une sauvegarde du classeur sera créée.`,
					{ modal: true },
					'Remplacer'
				);
				if (confirmation !== 'Remplacer') {
					await this.postStatus('info', 'Remplacement du gestionnaire annulé.');
					return;
				}
			}
			await this.applyDesignerOperations(
				[
					{
						kind: 'setUserFormEventHandler',
						formName,
						objectName,
						eventName,
						procedureSource,
						replaceExisting
					}
				],
				`Gestionnaire ${formName}.${objectName}_${eventName} affecté et vérifié.`
			);
		} catch (error) {
			await this.postStatus('error', (error as Error).message);
		}
	}

	private async assignMacro(
		message: WebviewMessage,
		activeX: boolean
	): Promise<void> {
		const sheetName =
			typeof message.sheetName === 'string' ? message.sheetName : '';
		const controlName =
			typeof message.controlName === 'string' ? message.controlName : '';
		const macroName =
			typeof message.macroName === 'string' ? message.macroName : '';
		const project = await this.buildProject();
		if (!project || !sheetName || !controlName || !macroName) {
			await this.postStatus('error', 'Affectation de macro incomplète.');
			return;
		}
		if (
			!project.interactions.macros.some(
				macro =>
					macro.qualifiedName.toLocaleLowerCase('en-US') ===
					macroName.toLocaleLowerCase('en-US')
			)
		) {
			await this.postStatus(
				'error',
				`Macro publique sans argument introuvable : ${macroName}.`
			);
			return;
		}
		const controls = activeX
			? project.interactions.worksheetActiveXControls
			: project.interactions.worksheetButtons;
		if (
			!controls.some(
				control =>
					control.sheetName === sheetName && control.name === controlName
			)
		) {
			await this.postStatus(
				'error',
				`Contrôle introuvable : ${sheetName}.${controlName}.`
			);
			return;
		}
		const confirmed = await vscode.window.showWarningMessage(
			activeX
				? `Créer ${controlName}_Click et appeler ${macroName} ?`
				: `Affecter ${macroName} au bouton ${controlName} ?`,
			{ modal: true },
			'Appliquer'
		);
		if (confirmed !== 'Appliquer') {
			return;
		}
		await this.applyDesignerOperations(
			[
				activeX
					? {
							kind: 'bindWorksheetActiveXMacro',
							sheetName,
							name: controlName,
							macroName
					  }
					: {
							kind: 'assignWorksheetButtonMacro',
							sheetName,
							name: controlName,
							macroName
					  }
			],
			activeX
				? `${sheetName}.${controlName} lié à ${macroName}.`
				: `${sheetName}.${controlName} affecté à ${macroName}.`
		);
	}

	private async createWorksheetActiveXControl(): Promise<void> {
		const project = await this.buildProject();
		if (!project || !project.worksheetNames.length) {
			await this.postStatus('error', 'Aucune feuille Excel disponible.');
			return;
		}
		const sheetName = await vscode.window.showQuickPick(project.worksheetNames, {
			placeHolder: 'Feuille recevant le contrôle ActiveX'
		});
		if (!sheetName) {
			return;
		}
		const standardTypes: Array<{
			label: string;
			type: Exclude<VbaUserFormControlType, 'customActiveX'>;
		}> = [
			{ label: 'CommandButton', type: 'commandButton' },
			{ label: 'ToggleButton', type: 'toggleButton' },
			{ label: 'Label', type: 'label' },
			{ label: 'TextBox', type: 'textBox' },
			{ label: 'ComboBox', type: 'comboBox' },
			{ label: 'ListBox', type: 'listBox' },
			{ label: 'CheckBox', type: 'checkBox' },
			{ label: 'OptionButton', type: 'optionButton' },
			{ label: 'Frame', type: 'frame' },
			{ label: 'Image', type: 'image' },
			{ label: 'SpinButton', type: 'spinButton' },
			{ label: 'ScrollBar', type: 'scrollBar' }
		];
		const customProgIds = vscode.workspace
			.getConfiguration(
				'excelAiVbaStudio',
				this.currentContext?.workbookUri
			)
			.get<string[]>('allowedCustomActiveXProgIds', []);
		const controlChoice = await vscode.window.showQuickPick(
			[
				...standardTypes.map(item => ({
					label: item.label,
					type: item.type as VbaUserFormControlType,
					progId: undefined as string | undefined
				})),
				...customProgIds.map(progId => ({
					label: `Personnalisé : ${progId}`,
					type: 'customActiveX' as VbaUserFormControlType,
					progId
				}))
			],
			{ placeHolder: 'Type de contrôle ActiveX' }
		);
		if (!controlChoice) {
			return;
		}
		const suggestedPrefix =
			controlChoice.type === 'customActiveX'
				? 'CustomControl'
				: controlChoice.type;
		const name = await vscode.window.showInputBox({
			title: 'Nom VBA du contrôle',
			value: `${suggestedPrefix}1`,
			validateInput: value =>
				/^[A-Za-z_][A-Za-z0-9_]{0,30}$/.test(value)
					? undefined
					: 'Identifiant VBA attendu, 31 caractères maximum.'
		});
		if (!name) {
			return;
		}
		const caption = await vscode.window.showInputBox({
			title: 'Texte visible du contrôle',
			value:
				controlChoice.type === 'commandButton' ||
				controlChoice.type === 'toggleButton'
					? 'Ouvrir'
					: ''
		});
		if (caption === undefined) {
			return;
		}
		const controlsOnSheet =
			project.interactions.worksheetActiveXControls.filter(
				control => control.sheetName === sheetName
			).length;
		const top = Math.min(9_000, 20 + controlsOnSheet * 36);
		const control = {
			type: controlChoice.type,
			name,
			left: 20,
			top,
			width: 110,
			height: 28,
			...(caption ? { caption } : {}),
			...(controlChoice.progId ? { progId: controlChoice.progId } : {})
		};
		const operations: VbaDesignOperation[] = [
			{
				kind: 'createWorksheetActiveXControl',
				sheetName,
				control
			}
		];
		if (
			controlChoice.type === 'commandButton' ||
			controlChoice.type === 'toggleButton'
		) {
			const macroChoice = await vscode.window.showQuickPick(
				[
					{ label: 'Sans liaison', macroName: '' },
					...project.interactions.macros.map(macro => ({
						label: macro.qualifiedName,
						description: macro.userFormsOpened.length
							? `Ouvre ${macro.userFormsOpened.join(', ')}`
							: undefined,
						macroName: macro.qualifiedName
					}))
				],
				{ placeHolder: 'Macro appelée au clic' }
			);
			if (!macroChoice) {
				return;
			}
			if (macroChoice.macroName) {
				operations.push({
					kind: 'bindWorksheetActiveXMacro',
					sheetName,
					name,
					macroName: macroChoice.macroName
				});
			}
		}
		const confirmed = await vscode.window.showWarningMessage(
			`Créer le contrôle ActiveX ${sheetName}.${name} ?`,
			{ modal: true },
			'Créer'
		);
		if (confirmed !== 'Créer') {
			return;
		}
		await this.applyDesignerOperations(
			operations,
			`Contrôle ActiveX ${sheetName}.${name} créé et vérifié.`
		);
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
		if (kind === 'userform') {
			const project = await this.buildProject();
			const existingNames = new Set(
				(project?.categories.flatMap(category => category.components) || []).map(
					component => component.name.toLocaleLowerCase('en-US')
				)
			);
			let index = 1;
			while (existingNames.has(`userform${index}`)) {
				index++;
			}
			const name = await vscode.window.showInputBox({
				title: 'Nom du nouveau UserForm',
				value: `UserForm${index}`,
				validateInput: value =>
					/^[A-Za-z_][A-Za-z0-9_]{0,30}$/.test(value)
						? undefined
						: 'Identifiant VBA attendu, 31 caractères maximum.'
			});
			if (!name) {
				return;
			}
			const caption = await vscode.window.showInputBox({
				title: 'Titre du UserForm',
				value: name
			});
			if (caption === undefined) {
				return;
			}
			await this.applyDesignerOperations(
				[
					{
						kind: 'createUserForm',
						name,
						caption,
						width: 400,
						height: 300,
						source: 'Option Explicit'
					}
				],
				`${name} créé avec son designer et son flux .frx.`
			);
			return;
		}
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
		try {
			await this.writebackService.prepare(context);
			const source = await fs.promises.readFile(sourcePath, 'utf8');
			const result = await this.writebackService.applySource(context, file, source);
			await this.postProject(file);
			await this.postStatus(
				'saved',
				result.changed
					? `${file} créé et réinjecté dans le classeur.`
					: `${file} créé — déjà synchronisé.`
			);
		} catch (error) {
			await this.postProject(file);
			await this.postStatus(
				'error',
				`${file} créé dans VS Code, mais non réinjecté : ${(error as Error).message}`
			);
		}
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
			'Pour créer un UserForm réel, ajouter ou repositionner ses contrôles, affecter une procédure événementielle complète ou créer un bouton de feuille, utilise #excelVbaDesignWorkbook.',
			'Pour un événement, setUserFormEventHandler accepte une procédure Private Sub complète, y compris une signature complexe avec paramètres ; replaceExisting doit être explicitement vrai pour remplacer un gestionnaire.',
			'N’exécute aucune macro. Le test réel des événements reste une action explicite dans Excel.'
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

	private async handleExternalDocumentSave(
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
		const extension = path.extname(file).toLocaleLowerCase('en-US');
		if (
			this.savingFile === file ||
			!SOURCE_EXTENSIONS.has(extension) ||
			extension === '.txt'
		) {
			return;
		}
		try {
			if (!this.componentFiles.has(file)) {
				if (extension === '.frm') {
					throw new Error(
						'Un nouveau UserForm exige un designer .frx ; utilisez un formulaire modèle existant.'
					);
				}
				await this.writebackService.prepare(context);
			}
			const result = await this.writebackService.applySource(
				context,
				file,
				document.getText()
			);
			if (!this.componentFiles.has(file)) {
				await this.postProject(file);
			}
			await this.postStatus(
				'saved',
				result.changed
					? `${file} réinjecté automatiquement dans le classeur.`
					: `${file} déjà synchronisé avec le classeur.`
			);
		} catch (error) {
			await this.postStatus(
				'error',
				`Réinjection automatique refusée : ${errorMessage(error)}`
			);
		}
	}

	private runInBackground(operation: string, promise: Promise<void>): void {
		void promise.catch(error => {
			const message = errorMessage(error);
			this.outputChannel.appendLine(`[vba studio] ${operation} : ${message}`);
			void this.postStatus('error', `${operation} : ${message}`).catch(
				reportingError => {
					this.outputChannel.appendLine(
						`[vba studio] Affichage de l’erreur impossible : ${errorMessage(reportingError)}`
					);
				}
			);
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
:root{color-scheme:light}
*{box-sizing:border-box}
body{
  --studio-bg:var(--vscode-editor-background,#ffffff);
  --studio-fg:var(--vscode-editor-foreground,var(--vscode-foreground,#1f2328));
  --studio-muted:var(--vscode-descriptionForeground,#61666d);
  --studio-surface:var(--vscode-sideBar-background,var(--vscode-editor-background,#ffffff));
  --studio-toolbar:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));
  --studio-section:var(--vscode-sideBarSectionHeader-background,var(--studio-toolbar));
  --studio-border:var(--vscode-panel-border,var(--vscode-editorWidget-border,#d0d7de));
  --studio-hover:var(--vscode-list-hoverBackground,rgba(0,0,0,.06));
  --studio-hover-fg:var(--vscode-list-hoverForeground,var(--studio-fg));
  --studio-selected:var(--vscode-list-activeSelectionBackground,#0969da);
  --studio-selected-fg:var(--vscode-list-activeSelectionForeground,#ffffff);
  --studio-focus:var(--vscode-focusBorder,#0969da);
  --studio-brand:#107c41;
  --studio-brand-fg:#ffffff;
}
body.vscode-light,body.vscode-high-contrast-light{color-scheme:light}
body.vscode-dark,body.vscode-high-contrast{
  color-scheme:dark;
  --studio-bg:var(--vscode-editor-background,#1e1e1e);
  --studio-fg:var(--vscode-editor-foreground,var(--vscode-foreground,#cccccc));
  --studio-muted:var(--vscode-descriptionForeground,#9b9b9b);
  --studio-surface:var(--vscode-sideBar-background,var(--vscode-editor-background,#1e1e1e));
  --studio-toolbar:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#252526));
  --studio-border:var(--vscode-panel-border,var(--vscode-editorWidget-border,#454545));
  --studio-hover:var(--vscode-list-hoverBackground,rgba(255,255,255,.08));
  --studio-selected:var(--vscode-list-activeSelectionBackground,#094771);
  --studio-selected-fg:var(--vscode-list-activeSelectionForeground,#ffffff);
  --studio-focus:var(--vscode-focusBorder,#007fd4);
  --studio-brand:var(--vscode-charts-green,#16825d);
}
body.vscode-high-contrast,body.vscode-high-contrast-light{
  --studio-border:var(--vscode-contrastBorder,var(--vscode-panel-border,#6fc3df));
  --studio-focus:var(--vscode-focusBorder,var(--vscode-contrastActiveBorder,#f38518));
}
html,body{height:100%;margin:0;overflow:hidden;font-family:var(--vscode-font-family);font-size:13px;color:var(--studio-fg);background:var(--studio-bg)}
button,select,textarea{font:inherit}
button,select,textarea{color:inherit}
button:focus-visible,select:focus-visible,textarea:focus-visible{outline:1px solid var(--studio-focus);outline-offset:-1px}
.app{height:100%;display:grid;grid-template-rows:38px 1fr 24px}
.toolbar{display:flex;align-items:center;gap:4px;padding:4px 7px;border-bottom:1px solid var(--studio-border);background:var(--studio-toolbar)}
.brand{display:flex;align-items:center;gap:7px;font-weight:650;margin-right:8px;white-space:nowrap}
.brand-icon{width:22px;height:22px;display:grid;place-items:center;color:var(--studio-brand-fg);background:var(--studio-brand);border-radius:3px}
.tool{height:27px;padding:0 9px;border:1px solid transparent;color:var(--studio-fg);background:transparent;border-radius:3px;cursor:pointer}
.tool:hover{color:var(--studio-hover-fg);background:var(--vscode-toolbar-hoverBackground,var(--studio-hover));border-color:var(--studio-border)}
.tool:disabled{opacity:.45;cursor:default}
.separator{height:20px;border-left:1px solid var(--studio-border);margin:0 3px}
.spacer{flex:1}
.ai{color:var(--vscode-button-foreground,#ffffff);background:var(--vscode-button-background,#0e639c);border-color:var(--vscode-button-border,transparent)}
.ai:hover{color:var(--vscode-button-foreground,#ffffff);background:var(--vscode-button-hoverBackground,#1177bb)}
.workspace{min-height:0;display:grid;grid-template-columns:minmax(260px,32%) 1fr}
.left{min-width:0;min-height:0;display:grid;grid-template-rows:minmax(180px,58%) minmax(130px,42%);border-right:1px solid var(--studio-border);background:var(--studio-surface)}
.pane{min-height:0;display:flex;flex-direction:column}
.pane+.pane{border-top:1px solid var(--studio-border)}
.pane-title{height:27px;display:flex;align-items:center;padding:0 8px;font-weight:650;background:var(--studio-section);border-bottom:1px solid var(--studio-border)}
.pane-body{min-height:0;overflow:auto;padding:4px 0}
.project-root,.category,.component{display:flex;align-items:center;height:23px;gap:5px;white-space:nowrap;cursor:default}
.project-root{padding-left:7px;font-weight:600}
.category{padding-left:20px}
.component{padding-left:42px}
.component:hover{color:var(--studio-hover-fg);background:var(--studio-hover)}
.component.selected{color:var(--studio-selected-fg);background:var(--studio-selected)}
.twisty{width:12px;text-align:center;opacity:.8}
.node-icon{width:17px;text-align:center;color:var(--vscode-symbolIcon-methodForeground,#b180d7)}
.node-name{overflow:hidden;text-overflow:ellipsis}
.count{margin-left:auto;padding-right:8px;opacity:.55}
.props{width:100%;border-collapse:collapse;font-size:12px}
.props td{padding:4px 6px;border-bottom:1px solid var(--studio-border);vertical-align:top}
.props td:first-child{width:38%;font-weight:600;background:var(--vscode-editorWidget-background,var(--studio-toolbar))}
.props td:last-child{word-break:break-word}
.editor-area{min-width:0;min-height:0;display:flex;flex-direction:column}
.code-view{min-width:0;min-height:0;flex:1;display:grid;grid-template-rows:30px auto 1fr}
.code-view.hidden{display:none}
.editor-head{display:grid;grid-template-columns:1fr 1fr;border-bottom:1px solid var(--studio-border);background:var(--studio-toolbar)}
.editor-head select{min-width:0;border:0;border-right:1px solid var(--studio-border);padding:0 8px;color:var(--vscode-dropdown-foreground,var(--studio-fg));background:var(--vscode-dropdown-background,var(--studio-toolbar))}
.notice{display:none;padding:6px 10px;border-bottom:1px solid var(--vscode-inputValidation-warningBorder,var(--studio-border));background:var(--vscode-inputValidation-warningBackground,#fff4ce);color:var(--vscode-inputValidation-warningForeground,#5f4500)}
.notice.visible{display:block}
.code-wrap{min-height:0;position:relative;display:grid;grid-template-columns:48px 1fr;background:var(--studio-bg)}
.lines{margin:0;padding:10px 8px 10px 0;overflow:hidden;text-align:right;white-space:pre;color:var(--vscode-editorLineNumber-foreground,var(--studio-muted));background:var(--vscode-editorGutter-background,var(--studio-bg));border-right:1px solid var(--studio-border);font-family:var(--vscode-editor-font-family);font-size:var(--vscode-editor-font-size);line-height:var(--vscode-editor-line-height)}
.code{width:100%;height:100%;resize:none;border:0;outline:0;padding:10px 12px;tab-size:4;white-space:pre;overflow:auto;color:var(--studio-fg);background:var(--studio-bg);caret-color:var(--studio-fg);font-family:var(--vscode-editor-font-family);font-size:var(--vscode-editor-font-size);line-height:var(--vscode-editor-line-height)}
.code:disabled{opacity:.72}
.controls-view{display:none;min-height:0;overflow:auto;padding:16px;background:var(--studio-bg)}
.controls-view.visible{display:block;flex:1}
.designer-view{display:none;min-height:0;flex:1;background:var(--studio-bg)}
.designer-view.visible{display:block}
.designer-layout{height:100%;display:grid;grid-template-columns:170px minmax(300px,1fr) 290px}
.designer-palette,.designer-inspector{min-width:0;overflow:auto;padding:10px;background:var(--studio-surface)}
.designer-palette{border-right:1px solid var(--studio-border)}
.designer-inspector{border-left:1px solid var(--studio-border)}
.designer-palette h3,.designer-inspector h3{margin:3px 0 9px;font-size:12px}
.palette-grid{display:grid;grid-template-columns:1fr 1fr;gap:5px}
.palette-control{min-height:34px;padding:4px;border:1px solid var(--studio-border);border-radius:3px;color:var(--studio-fg);background:var(--studio-toolbar);cursor:pointer;font-size:11px}
.palette-control:hover{background:var(--studio-hover)}
.designer-stage{min-width:0;min-height:0;overflow:auto;padding:24px;background:var(--vscode-editor-background,var(--studio-bg))}
.userform-canvas{position:relative;min-width:220px;min-height:140px;border:1px solid var(--studio-border);box-shadow:0 8px 28px rgba(0,0,0,.18);background-color:var(--vscode-input-background,#fff);background-image:radial-gradient(var(--studio-border) .8px,transparent .8px);background-size:8px 8px;color:var(--vscode-input-foreground,#111)}
.userform-caption{position:absolute;left:0;right:0;top:-24px;height:23px;display:flex;align-items:center;padding:0 7px;border:1px solid var(--studio-border);background:var(--studio-toolbar);color:var(--studio-fg);font-weight:600}
.design-control{position:absolute;display:flex;align-items:center;justify-content:center;box-sizing:border-box;overflow:hidden;border:1px solid #7a7a7a;background:#efefef;color:#111;cursor:move;user-select:none;font-family:Segoe UI,sans-serif;font-size:12px}
.design-control.input,.design-control.combo,.design-control.list{justify-content:flex-start;padding:2px 5px;background:#fff}
.design-control.label{justify-content:flex-start;border-color:transparent;background:transparent}
.design-control.frame{justify-content:flex-start;align-items:flex-start;padding:7px;border-style:groove;background:transparent}
.design-control.image{border-style:dashed;background:#ddd}
.design-control.hidden-control{opacity:.45}
.design-control.selected{outline:2px solid var(--studio-focus);outline-offset:1px;z-index:20}
.resize-handle{position:absolute;right:-1px;bottom:-1px;width:9px;height:9px;border:1px solid #fff;background:var(--studio-focus);cursor:nwse-resize}
.design-control:not(.selected) .resize-handle{display:none}
.inspector-grid{display:grid;grid-template-columns:92px 1fr;gap:6px;align-items:center}
.inspector-grid label{font-size:11px;color:var(--studio-muted)}
.inspector-grid input,.inspector-grid select,.event-editor textarea{min-width:0;box-sizing:border-box;border:1px solid var(--studio-border);padding:5px;color:var(--vscode-input-foreground,var(--studio-fg));background:var(--vscode-input-background,var(--studio-bg))}
.inspector-actions{display:flex;gap:6px;margin:10px 0 14px}
.event-editor{border-top:1px solid var(--studio-border);padding-top:10px}
.event-editor textarea{width:100%;height:180px;resize:vertical;font-family:var(--vscode-editor-font-family);font-size:12px}
.event-existing{margin:6px 0;color:var(--vscode-inputValidation-warningForeground,#cca700)}
.controls-summary{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.metric{padding:7px 10px;border:1px solid var(--studio-border);border-radius:4px;background:var(--studio-surface)}
.interaction-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:10px}
.interaction-card{padding:12px;border:1px solid var(--studio-border);border-radius:5px;background:var(--studio-surface)}
.interaction-card h3{margin:0 0 8px;font-size:13px}
.flow{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:8px 0;color:var(--studio-muted)}
.flow-node{padding:4px 7px;border:1px solid var(--studio-border);border-radius:3px;color:var(--studio-fg);background:var(--studio-toolbar)}
.flow-arrow{font-weight:700}
.assignment{display:grid;grid-template-columns:1fr auto;gap:6px;margin-top:9px}
.assignment select{min-width:0;border:1px solid var(--studio-border);padding:5px;color:var(--vscode-dropdown-foreground,var(--studio-fg));background:var(--vscode-dropdown-background,var(--studio-toolbar))}
.empty-state{padding:18px;border:1px dashed var(--studio-border);color:var(--studio-muted)}
.statusbar{display:flex;align-items:center;gap:16px;padding:0 8px;color:var(--vscode-statusBar-foreground,#ffffff);background:var(--vscode-statusBar-background,#007acc);border-top:1px solid var(--vscode-statusBar-border,transparent)}
.statusbar .right{margin-left:auto}
.status-ok{color:var(--vscode-testing-iconPassed,#73c991)}
.status-error{color:var(--vscode-errorForeground,#f14c4c)}
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
    <button class="tool" id="controls-tab">Contrôles</button>
    <button class="tool" id="designer-tab" disabled>Designer</button>
    <button class="tool" id="create-activex">＋ ActiveX</button>
    <button class="tool" id="open-excel">Excel</button>
    <button class="tool" id="open-vbe">VBE</button>
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
      <div class="code-view" id="code-view">
        <div class="editor-head">
          <select id="object-select" aria-label="Objet"></select>
          <select id="procedure-select" aria-label="Procédure"><option>(Général)</option></select>
        </div>
        <div class="notice" id="notice"></div>
        <div class="code-wrap">
          <pre class="lines" id="lines">1</pre>
          <textarea class="code" id="code" spellcheck="false" aria-label="Code VBA"></textarea>
        </div>
      </div>
      <div class="controls-view" id="controls-view"></div>
      <div class="designer-view" id="designer-view"></div>
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
  const state = {
    project: null,
    selected: null,
    dirty: false,
    viewMode: 'code',
    selectedDesignControl: null
  };
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
  const codeView = document.getElementById('code-view');
  const controlsView = document.getElementById('controls-view');
  const controlsTab = document.getElementById('controls-tab');
  const designerView = document.getElementById('designer-view');
  const designerTab = document.getElementById('designer-tab');

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

  function setViewMode(mode) {
    state.viewMode = mode;
    codeView.classList.toggle('hidden', mode !== 'code');
    controlsView.classList.toggle('visible', mode === 'controls');
    designerView.classList.toggle('visible', mode === 'designer');
    controlsTab.textContent = mode === 'controls' ? 'Code' : 'Contrôles';
    designerTab.textContent = mode === 'designer' ? 'Code' : 'Designer';
  }

  function renderControls(project) {
    const graph = project.interactions;
    const macroOptions = graph.macros.map(macro =>
      '<option value="' + escapeText(macro.qualifiedName) + '">' +
      escapeText(macro.qualifiedName) +
      (macro.userFormsOpened.length ? ' · ouvre ' + escapeText(macro.userFormsOpened.join(', ')) : '') +
      '</option>'
    ).join('');
    const cards = graph.relationships.map((relationship, index) => {
      const activeX = relationship.kind === 'activeX';
      const activeXRecord = activeX
        ? graph.worksheetActiveXControls.find(control =>
            control.sheetName === relationship.sheetName &&
            control.name === relationship.controlName
          )
        : null;
      const bindable = !activeX ||
        activeXRecord?.progId === 'Forms.CommandButton.1' ||
        activeXRecord?.progId === 'Forms.ToggleButton.1';
      const macroNode = relationship.macroName || 'macro non affectée';
      const formNodes = relationship.userFormsOpened.length
        ? relationship.userFormsOpened.map(name =>
            '<span class="flow-arrow">›</span><span class="flow-node">' +
            escapeText(name) + '</span>'
          ).join('')
        : '';
      const assignment = bindable && graph.macros.length
        ? '<div class="assignment"><select class="macro-assignment" data-index="' + index + '">' +
          '<option value="">Choisir une macro…</option>' + macroOptions +
          '</select><button class="tool apply-assignment" data-index="' + index + '">Affecter</button></div>'
        : '<div class="empty-state">' +
          (bindable ? 'Aucune macro publique sans argument.' : 'Liaison Click limitée aux CommandButton et ToggleButton MSForms.') +
          '</div>';
      const simulation = relationship.userFormsOpened.length
        ? '<button class="tool simulate-flow" data-index="' + index + '">Simuler le clic</button>'
        : '';
      return '<article class="interaction-card">' +
        '<h3>' + (activeX ? 'ActiveX' : 'Bouton formulaire') + ' · ' +
        escapeText(relationship.sheetName + '.' + relationship.controlName) + '</h3>' +
        '<div class="flow"><span class="flow-node">' + escapeText(relationship.controlCaption) +
        '</span><span class="flow-arrow">›</span><span class="flow-node">' +
        escapeText(macroNode) + '</span>' + formNodes + '</div>' +
        '<div>' + escapeText(activeXRecord?.progId || relationship.resolution) + '</div>' +
        assignment + simulation +
        '</article>';
    }).join('');
    controlsView.innerHTML =
      '<div class="controls-summary">' +
      '<div class="metric">Macros : ' + graph.macros.length + '</div>' +
      '<div class="metric">UserForms : ' + graph.userForms.length + '</div>' +
      '<div class="metric">Boutons : ' + graph.worksheetButtons.length + '</div>' +
      '<div class="metric">ActiveX : ' + graph.worksheetActiveXControls.length + '</div>' +
      '</div>' +
      '<p>Aperçu statique : aucun code VBA n’est exécuté dans VS Code. Le clic réel reste dans Excel.</p>' +
      (cards ? '<div class="interaction-grid">' + cards + '</div>' :
        '<div class="empty-state">Aucun bouton de formulaire ou contrôle ActiveX détecté.</div>');
    controlsView.querySelectorAll('.apply-assignment').forEach(button =>
      button.addEventListener('click', () => {
        const index = Number(button.dataset.index);
        const relationship = graph.relationships[index];
        const select = controlsView.querySelector('.macro-assignment[data-index="' + index + '"]');
        const macroName = select?.value || '';
        if (!relationship || !macroName) {
          setStatus('Choisissez une macro.', 'error');
          return;
        }
        vscode.postMessage({
          type: relationship.kind === 'activeX' ? 'bindActiveX' : 'assignFormButton',
          sheetName: relationship.sheetName,
          controlName: relationship.controlName,
          macroName
        });
      })
    );
    controlsView.querySelectorAll('.simulate-flow').forEach(button =>
      button.addEventListener('click', () => {
        const relationship = graph.relationships[Number(button.dataset.index)];
        if (!relationship) return;
        setStatus(
          'Simulation uniquement : ' + relationship.controlName + ' afficherait ' +
          relationship.userFormsOpened.join(', ') + '. Aucune macro exécutée.'
        );
      })
    );
  }

  const visualControlPalette = [
    ['label', 'Label'],
    ['textBox', 'TextBox'],
    ['commandButton', 'Bouton'],
    ['comboBox', 'ComboBox'],
    ['listBox', 'ListBox'],
    ['checkBox', 'CheckBox'],
    ['optionButton', 'Option'],
    ['toggleButton', 'Toggle'],
    ['frame', 'Frame'],
    ['image', 'Image'],
    ['spinButton', 'Spin'],
    ['scrollBar', 'Scroll']
  ];

  const controlPrefixes = {
    label: 'lbl',
    textBox: 'txt',
    commandButton: 'cmd',
    comboBox: 'cbo',
    listBox: 'lst',
    checkBox: 'chk',
    optionButton: 'opt',
    toggleButton: 'tgl',
    frame: 'fra',
    image: 'img',
    spinButton: 'spn',
    scrollBar: 'scr'
  };

  const captionControlTypes = new Set([
    'label', 'commandButton', 'checkBox', 'optionButton', 'toggleButton', 'frame'
  ]);

  function designControlClass(type) {
    if (type === 'textBox') return 'input';
    if (type === 'comboBox') return 'combo';
    if (type === 'listBox') return 'list';
    if (type === 'label') return 'label';
    if (type === 'frame') return 'frame';
    if (type === 'image') return 'image';
    return 'button';
  }

  function escapeRegex(value) {
    return String(value).replace(/[-/\\\\^$*+?.()|[\\]{}]/g, '\\\\$&');
  }

  function findEventProcedure(component, objectName, eventName) {
    const procedureName = objectName + '_' + eventName;
    const matcher = new RegExp(
      '^[\\\\t ]*Private[\\\\t ]+Sub[\\\\t ]+' + escapeRegex(procedureName) +
      '[\\\\t ]*\\\\([^\\\\r\\\\n]*\\\\)[^\\\\r\\\\n]*(?:\\\\r?\\\\n)[\\\\s\\\\S]*?' +
      '^[\\\\t ]*End[\\\\t ]+Sub\\\\b',
      'im'
    );
    return matcher.exec(component.source || '')?.[0] || '';
  }

  function eventArguments(eventName) {
    const normalized = String(eventName).toLowerCase();
    if (normalized === 'queryclose') return 'Cancel As Integer, CloseMode As Integer';
    if (normalized === 'beforeupdate') return 'ByVal Cancel As MSForms.ReturnBoolean';
    if (normalized === 'keydown' || normalized === 'keyup') {
      return 'ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer';
    }
    if (normalized === 'keypress') return 'ByVal KeyAscii As MSForms.ReturnInteger';
    if (normalized === 'mousedown' || normalized === 'mousemove' || normalized === 'mouseup') {
      return 'ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single';
    }
    return '';
  }

  function eventSkeleton(objectName, eventName) {
    return 'Private Sub ' + objectName + '_' + eventName + '(' +
      eventArguments(eventName) + ')\\r\\n' +
      "    ' Code VBA ici\\r\\n" +
      'End Sub';
  }

  function nextControlName(type, controls) {
    const prefix = controlPrefixes[type] || 'ctl';
    const used = new Set(controls.map(control => control.name.toLowerCase()));
    let index = 1;
    while (used.has((prefix + index).toLowerCase())) index++;
    return prefix + index;
  }

  function defaultControlSize(type) {
    if (type === 'label') return [110, 20];
    if (type === 'textBox' || type === 'comboBox') return [130, 24];
    if (type === 'listBox') return [140, 70];
    if (type === 'frame') return [180, 100];
    if (type === 'image') return [90, 70];
    if (type === 'spinButton') return [20, 40];
    if (type === 'scrollBar') return [140, 20];
    return [100, 28];
  }

  function renderFormEventInspector(form, component) {
    const host = document.getElementById('designer-inspector-body');
    if (!host) return;
    host.innerHTML =
      '<h3>' + escapeText(form.name) + '</h3>' +
      '<p>Code événementiel du UserForm. Aucune macro n’est exécutée dans VS Code.</p>' +
      '<div class="event-editor"><div class="inspector-grid">' +
      '<label>Objet</label><input value="UserForm" disabled>' +
      '<label>Événement</label><input id="d-event" list="form-event-list" value="Initialize">' +
      '<datalist id="form-event-list"><option>Initialize</option><option>Activate</option>' +
      '<option>Deactivate</option><option>Terminate</option><option>QueryClose</option>' +
      '<option>Resize</option><option>Click</option><option>DblClick</option>' +
      '<option>KeyDown</option><option>KeyPress</option><option>KeyUp</option></datalist></div>' +
      '<div class="event-existing" id="event-existing"></div>' +
      '<textarea id="event-source" spellcheck="false"></textarea>' +
      '<label><input id="replace-event" type="checkbox"> Remplacer explicitement si présent</label>' +
      '<div class="inspector-actions"><button class="tool ai" id="apply-event">Affecter le code</button></div></div>';
    const eventNameInput = document.getElementById('d-event');
    const eventSource = document.getElementById('event-source');
    const existingStatus = document.getElementById('event-existing');
    const replaceEvent = document.getElementById('replace-event');
    const refresh = () => {
      const eventName = eventNameInput.value.trim() || 'Initialize';
      const existing = findEventProcedure(component, 'UserForm', eventName);
      eventSource.value = existing || eventSkeleton('UserForm', eventName);
      existingStatus.textContent = existing
        ? 'Gestionnaire existant détecté. Cochez le remplacement pour le modifier.'
        : 'Nouveau gestionnaire du formulaire.';
      replaceEvent.checked = false;
    };
    eventNameInput.onchange = refresh;
    document.getElementById('apply-event').onclick = () =>
      vscode.postMessage({
        type: 'setUserFormEvent',
        formName: form.name,
        objectName: 'UserForm',
        eventName: eventNameInput.value.trim(),
        procedureSource: eventSource.value,
        replaceExisting: replaceEvent.checked
      });
    refresh();
  }

  function renderDesignerInspector(form, control, component) {
    const host = document.getElementById('designer-inspector-body');
    if (!host || !control) return;
    const hasCaption = captionControlTypes.has(control.type);
    host.innerHTML =
      '<h3>' + escapeText(form.name + '.' + control.name) + '</h3>' +
      '<div class="inspector-grid">' +
      '<label>Type</label><input value="' + escapeText(control.typeName || control.type) + '" disabled>' +
      '<label>Caption</label><input id="d-caption" value="' + escapeText(control.caption || '') + '"' +
        (hasCaption ? '' : ' disabled') + '>' +
      '<label>Left</label><input id="d-left" type="number" min="0" max="10000" step="1" value="' + control.left + '">' +
      '<label>Top</label><input id="d-top" type="number" min="0" max="10000" step="1" value="' + control.top + '">' +
      '<label>Width</label><input id="d-width" type="number" min="1" max="10000" step="1" value="' + control.width + '">' +
      '<label>Height</label><input id="d-height" type="number" min="1" max="10000" step="1" value="' + control.height + '">' +
      '<label>Visible</label><input id="d-visible" type="checkbox"' + (control.visible ? ' checked' : '') + '>' +
      '<label>Enabled</label><input id="d-enabled" type="checkbox"' + (control.enabled ? ' checked' : '') + '>' +
      '<label>TabIndex</label><input id="d-tabindex" type="number" min="0" max="32767" step="1" value="' +
        (control.tabIndex ?? '') + '">' +
      '<label>Info-bulle</label><input id="d-tip" value="' + escapeText(control.controlTipText || '') + '">' +
      '</div>' +
      '<div class="inspector-actions"><button class="tool ai" id="apply-control-props">Appliquer</button></div>' +
      '<div class="event-editor"><h3>Événement VBA</h3>' +
      '<div class="inspector-grid"><label>Objet</label><input value="' + escapeText(control.name) + '" disabled>' +
      '<label>Événement</label><input id="d-event" list="event-list" value="Click">' +
      '<datalist id="event-list"><option>Click</option><option>Change</option><option>DblClick</option>' +
      '<option>Enter</option><option>Exit</option><option>BeforeUpdate</option><option>AfterUpdate</option>' +
      '<option>KeyDown</option><option>KeyPress</option><option>KeyUp</option>' +
      '<option>MouseDown</option><option>MouseMove</option><option>MouseUp</option></datalist></div>' +
      '<div class="event-existing" id="event-existing"></div>' +
      '<textarea id="event-source" spellcheck="false"></textarea>' +
      '<label><input id="replace-event" type="checkbox"> Remplacer explicitement si présent</label>' +
      '<div class="inspector-actions"><button class="tool ai" id="apply-event">Affecter le code</button></div>' +
      '</div>';

    document.getElementById('apply-control-props').onclick = () => {
      const changes = {
        left: Number(document.getElementById('d-left').value),
        top: Number(document.getElementById('d-top').value),
        width: Number(document.getElementById('d-width').value),
        height: Number(document.getElementById('d-height').value),
        visible: document.getElementById('d-visible').checked,
        enabled: document.getElementById('d-enabled').checked,
        controlTipText: document.getElementById('d-tip').value
      };
      if (hasCaption) changes.caption = document.getElementById('d-caption').value;
      const tabIndex = document.getElementById('d-tabindex').value;
      if (tabIndex !== '') changes.tabIndex = Number(tabIndex);
      vscode.postMessage({
        type: 'updateVisualControl',
        formName: form.name,
        controlName: control.name,
        changes
      });
    };

    const eventNameInput = document.getElementById('d-event');
    const eventSource = document.getElementById('event-source');
    const existingStatus = document.getElementById('event-existing');
    const replaceEvent = document.getElementById('replace-event');
    const refreshEventEditor = () => {
      const eventName = eventNameInput.value.trim() || 'Click';
      const existing = findEventProcedure(component, control.name, eventName);
      eventSource.value = existing || eventSkeleton(control.name, eventName);
      existingStatus.textContent = existing
        ? 'Gestionnaire existant détecté. Cochez le remplacement pour le modifier.'
        : 'Nouveau gestionnaire. Il sera ajouté sans réécrire les autres procédures.';
      replaceEvent.checked = false;
    };
    eventNameInput.onchange = refreshEventEditor;
    document.getElementById('apply-event').onclick = () => {
      const eventName = eventNameInput.value.trim();
      vscode.postMessage({
        type: 'setUserFormEvent',
        formName: form.name,
        objectName: control.name,
        eventName,
        procedureSource: eventSource.value,
        replaceExisting: replaceEvent.checked
      });
    };
    refreshEventEditor();
  }

  function selectDesignControl(form, control, component) {
    state.selectedDesignControl = control.name;
    designerView.querySelectorAll('.design-control').forEach(element =>
      element.classList.toggle('selected', element.dataset.name === control.name)
    );
    renderDesignerInspector(form, control, component);
  }

  function bindDesignerDrag(form, component) {
    designerView.querySelectorAll('.design-control').forEach(element => {
      const control = form.controls.find(item => item.name === element.dataset.name);
      if (!control) return;
      element.addEventListener('pointerdown', event => {
        event.preventDefault();
        selectDesignControl(form, control, component);
        const resizing = event.target.classList.contains('resize-handle');
        const startX = event.clientX;
        const startY = event.clientY;
        const startLeft = control.left;
        const startTop = control.top;
        const startWidth = control.width;
        const startHeight = control.height;
        const move = moveEvent => {
          const deltaX = moveEvent.clientX - startX;
          const deltaY = moveEvent.clientY - startY;
          if (resizing) {
            const width = Math.max(12, Math.round((startWidth + deltaX) / 4) * 4);
            const height = Math.max(12, Math.round((startHeight + deltaY) / 4) * 4);
            element.style.width = Math.min(width, form.width - startLeft) + 'px';
            element.style.height = Math.min(height, form.height - startTop) + 'px';
          } else {
            const left = Math.max(0, Math.round((startLeft + deltaX) / 4) * 4);
            const top = Math.max(0, Math.round((startTop + deltaY) / 4) * 4);
            element.style.left = Math.min(left, form.width - startWidth) + 'px';
            element.style.top = Math.min(top, form.height - startHeight) + 'px';
          }
        };
        const up = () => {
          window.removeEventListener('pointermove', move);
          window.removeEventListener('pointerup', up);
          const changes = resizing
            ? {
                width: Number.parseFloat(element.style.width),
                height: Number.parseFloat(element.style.height)
              }
            : {
                left: Number.parseFloat(element.style.left),
                top: Number.parseFloat(element.style.top)
              };
          vscode.postMessage({
            type: 'updateVisualControl',
            formName: form.name,
            controlName: control.name,
            changes
          });
        };
        window.addEventListener('pointermove', move);
        window.addEventListener('pointerup', up, { once: true });
      });
    });
  }

  function renderDesigner(component) {
    const form = state.project?.userForms?.find(
      candidate => candidate.name.toLowerCase() === component.name.toLowerCase()
    );
    if (!form) {
      designerView.innerHTML =
        '<div class="empty-state">Designer COM indisponible pour ce UserForm. Rechargez après export VBA autorisé.</div>';
      return;
    }
    const palette = visualControlPalette.map(item =>
      '<button class="palette-control" data-control-type="' + item[0] + '">' +
      escapeText(item[1]) + '</button>'
    ).join('');
    const controls = form.controls.map(control =>
      '<div class="design-control ' + designControlClass(control.type) +
      (control.visible ? '' : ' hidden-control') + '" data-name="' + escapeText(control.name) +
      '" title="' + escapeText((control.typeName || control.type) + ': ' + control.name) +
      '" style="left:' + control.left + 'px;top:' + control.top + 'px;width:' +
      control.width + 'px;height:' + control.height + 'px">' +
      '<span>' + escapeText(control.caption || control.name) + '</span>' +
      '<span class="resize-handle"></span></div>'
    ).join('');
    designerView.innerHTML =
      '<div class="designer-layout">' +
      '<aside class="designer-palette"><h3>Boîte à outils</h3><div class="palette-grid">' +
      palette + '</div><div class="inspector-actions"><button class="tool" id="form-events">Événements UserForm</button></div>' +
      '<p>Ajout transactionnel. Déplacement et taille aimantés sur 4 points.</p></aside>' +
      '<section class="designer-stage"><div class="userform-canvas" style="width:' +
      form.width + 'px;height:' + form.height + 'px"><div class="userform-caption">' +
      escapeText(form.caption || form.name) + '</div>' + controls + '</div></section>' +
      '<aside class="designer-inspector" id="designer-inspector-body">' +
      '<div class="empty-state">Sélectionnez un contrôle.</div></aside></div>';

    designerView.querySelectorAll('.palette-control').forEach(button =>
      button.addEventListener('click', () => {
        const type = button.dataset.controlType;
        const name = nextControlName(type, form.controls);
        const size = defaultControlSize(type);
        const offset = 12 + (form.controls.length % 8) * 8;
        vscode.postMessage({
          type: 'addVisualControl',
          formName: form.name,
          control: {
            type,
            name,
            left: Math.min(offset, Math.max(0, form.width - size[0])),
            top: Math.min(offset, Math.max(0, form.height - size[1])),
            width: size[0],
            height: size[1],
            caption: captionControlTypes.has(type) ? name : ''
          }
        });
      })
    );
    document.getElementById('form-events').onclick = () => {
      state.selectedDesignControl = null;
      designerView.querySelectorAll('.design-control').forEach(element =>
        element.classList.remove('selected')
      );
      renderFormEventInspector(form, component);
    };
    bindDesignerDrag(form, component);
    const selected = form.controls.find(
      control => control.name === state.selectedDesignControl
    ) || form.controls[0];
    if (selected) selectDesignControl(form, selected, component);
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
    designerTab.disabled = component.kind !== 'userform';
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
    setViewMode('code');
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
    renderControls(project);
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
  controlsTab.onclick = () => setViewMode(
    state.viewMode === 'controls' ? 'code' : 'controls'
  );
  designerTab.onclick = () => {
    if (!state.selected || state.selected.kind !== 'userform') return;
    if (state.viewMode === 'designer') {
      setViewMode('code');
      return;
    }
    renderDesigner(state.selected);
    setViewMode('designer');
  };
  document.getElementById('create-activex').onclick = () =>
    vscode.postMessage({ type: 'createActiveX' });
  document.getElementById('open-excel').onclick = () =>
    vscode.postMessage({ type: 'openExcel' });
  document.getElementById('open-vbe').onclick = () =>
    vscode.postMessage({ type: 'openVbe' });
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
