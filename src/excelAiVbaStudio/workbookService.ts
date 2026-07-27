import { createHash } from 'crypto';
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import {
	assertNoReparsePointChain,
	assertNotManagedBackupPath,
	assertOwnedDirectory,
	assertLocalPath,
	canonicalizeWorkbookUri,
	ensureLocalDirectory,
	ensureOwnedDirectory,
	pathIsInside,
	removeOwnedDirectory,
	workbookUriFromPathInput
} from './security';
import {
	EXCEL_EXTENSIONS,
	ExcelAiSettings,
	ExportContext,
	ExportOptions,
	ExportPaths,
	ProcessResult,
	ToolInput,
	VbaToolWriteResult
} from './types';
import { showUserFormPreview } from './userFormPreview';
import { VbaStudioPanel } from './vbaStudioPanel';
import {
	VbaWritebackResult,
	VbaWritebackService
} from './vbaWritebackService';

const DEFAULT_MAX_ROWS = 200;
const DEFAULT_MAX_COLUMNS = 50;
const MAX_CONFIGURED_ROWS = 5_000;
const MAX_CONFIGURED_COLUMNS = 256;
const MAX_GENERATED_FILE_BYTES = 16 * 1024 * 1024;
const EXPORT_TIMEOUT_MS = 180_000;
const EXCEL_LAUNCH_TIMEOUT_MS = 45_000;
const VBA_BOOTSTRAP_TIMEOUT_MS = 90_000;

interface PowerShellRunOptions {
	progress?: vscode.Progress<{ message?: string }>;
	cancellationToken?: vscode.CancellationToken;
	timeoutMs: number;
	cleanupOwnedExcel: boolean;
}

interface MacroBootstrapResult {
	ok?: unknown;
	targetWorkbookPath?: unknown;
	sourceWorkbookPath?: unknown;
	convertedToXlsm?: unknown;
	changed?: unknown;
	modifiedModules?: unknown;
	workbookSha256?: unknown;
	macrosExecuted?: unknown;
	accessVbomChanged?: unknown;
}

function isUri(value: unknown): value is vscode.Uri {
	const candidate = value as vscode.Uri | undefined;
	return Boolean(
		candidate &&
			typeof candidate === 'object' &&
			typeof candidate.scheme === 'string' &&
			typeof candidate.fsPath === 'string'
	);
}

function getUriFromTabInput(input: unknown): vscode.Uri | undefined {
	if (!input || typeof input !== 'object') {
		return undefined;
	}
	const record = input as Record<string, unknown>;
	if (isUri(record.uri)) {
		return record.uri;
	}
	for (const key of ['modified', 'original', 'result', 'base', 'input1', 'input2']) {
		if (isUri(record[key])) {
			return record[key] as vscode.Uri;
		}
	}
	return undefined;
}

function getActiveResourceUri(): vscode.Uri | undefined {
	const vscodeWindow = vscode.window as typeof vscode.window & {
		tabGroups?: {
			activeTabGroup?: { activeTab?: { input?: unknown } };
		};
	};
	const tabUri = getUriFromTabInput(
		vscodeWindow.tabGroups?.activeTabGroup?.activeTab?.input
	);
	return tabUri || vscode.window.activeTextEditor?.document.uri;
}

function isExcelUri(uri: vscode.Uri | undefined): boolean {
	return Boolean(
		uri &&
			uri.scheme === 'file' &&
			!uri.authority &&
			EXCEL_EXTENSIONS.has(path.extname(uri.fsPath).toLocaleLowerCase('en-US'))
	);
}

function sameFile(left: string, right: string): boolean {
	return (
		path.resolve(left).toLocaleLowerCase('en-US') ===
		path.resolve(right).toLocaleLowerCase('en-US')
	);
}

function boundedInteger(
	value: unknown,
	fallback: number,
	minimum: number,
	maximum: number
): number {
	const parsed = Number(value);
	if (!Number.isFinite(parsed)) {
		return fallback;
	}
	return Math.min(maximum, Math.max(minimum, Math.trunc(parsed)));
}

function getPowerShellPath(): string {
	if (process.env.SystemRoot) {
		return path.join(
			process.env.SystemRoot,
			'System32',
			'WindowsPowerShell',
			'v1.0',
			'powershell.exe'
		);
	}
	return 'powershell.exe';
}

function safeFileStem(value: string): string {
	const normalized = value
		.normalize('NFKD')
		.replace(/[\u0300-\u036f]/g, '')
		.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-')
		.replace(/\s+/g, '-')
		.replace(/-+/g, '-')
		.replace(/^[.\-\s]+|[.\-\s]+$/g, '');
	return normalized.slice(0, 80) || 'workbook';
}

function processError(stderr: string, stdout: string): string {
	if (stderr.trim()) {
		return stderr.trim();
	}
	const lines = stdout.trim().split(/\r?\n/).filter(Boolean);
	for (let index = lines.length - 1; index >= 0; index--) {
		try {
			const payload = JSON.parse(lines[index]) as {
				error?: string;
				message?: string;
			};
			if (payload.error?.trim()) {
				return payload.error.trim();
			}
			if (payload.message?.trim()) {
				return payload.message.trim();
			}
		} catch {
			// Progress lines are intentionally not all JSON.
		}
	}
	return stdout.trim();
}

async function pathExists(candidatePath: string): Promise<boolean> {
	try {
		await fs.promises.lstat(candidatePath);
		return true;
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
			return false;
		}
		throw error;
	}
}

async function hashFileSha256(filePath: string): Promise<string> {
	const digest = createHash('sha256');
	const stream = fs.createReadStream(filePath);
	for await (const chunk of stream) {
		digest.update(chunk as Buffer);
	}
	return digest.digest('hex');
}

async function terminateExactProcess(processId: number): Promise<void> {
	if (
		process.platform !== 'win32' ||
		!Number.isSafeInteger(processId) ||
		processId <= 0
	) {
		return;
	}
	await new Promise<void>(resolve => {
		const taskkillPath = process.env.SystemRoot
			? path.join(process.env.SystemRoot, 'System32', 'taskkill.exe')
			: 'taskkill.exe';
		const killer = spawn(
			taskkillPath,
			['/PID', String(processId), '/F'],
			{
				windowsHide: true,
				shell: false,
				stdio: 'ignore'
			}
		);
		killer.once('error', () => resolve());
		killer.once('close', () => resolve());
	});
}

export class ExcelAiVbaWorkbookService implements vscode.Disposable {
	private readonly outputChannel = vscode.window.createOutputChannel('Excel AI & VBA Studio');
	private readonly contextChangeEmitter = new vscode.EventEmitter<void>();
	private readonly runningExports = new Map<string, Promise<ExportContext | undefined>>();
	private readonly vbaStudioPanel: VbaStudioPanel;
	private readonly vbaWritebackService: VbaWritebackService;
	private lastContext: ExportContext | undefined;
	private storageRoot: string | undefined;
	private exportsRoot: string | undefined;

	readonly onDidChangeContext = this.contextChangeEmitter.event;

	constructor(private readonly extensionContext: vscode.ExtensionContext) {
		this.vbaWritebackService = new VbaWritebackService(
			this.extensionContext,
			this.outputChannel
		);
		this.vbaStudioPanel = new VbaStudioPanel(
			this.outputChannel,
			this.vbaWritebackService
		);
	}

	dispose(): void {
		this.vbaStudioPanel.dispose();
		this.vbaWritebackService.dispose();
		this.contextChangeEmitter.dispose();
		this.outputChannel.dispose();
	}

	getLastContext(): ExportContext | undefined {
		return this.lastContext;
	}

	getOutputChannel(): vscode.OutputChannel {
		return this.outputChannel;
	}

	async applyVbaSource(
		context: ExportContext,
		file: string,
		source: string,
		persistSourceFile = false
	): Promise<VbaWritebackResult> {
		const sourcePath = path.join(context.paths.vbaDirectory, path.basename(file));
		if (
			path.basename(file) !== file ||
			!pathIsInside(sourcePath, context.paths.vbaDirectory)
		) {
			throw new Error('Le composant demandé sort du projet VBA.');
		}
		await assertNoReparsePointChain(
			sourcePath,
			context.paths.vbaDirectory
		);
		const sourceExists = await fs.promises
			.lstat(sourcePath)
			.then(stat => stat.isFile() && !stat.isSymbolicLink())
			.catch(error => {
				if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
					return false;
				}
				throw error;
			});
		if (persistSourceFile && !sourceExists) {
			if (path.extname(file).toLocaleLowerCase('en-US') === '.frm') {
				throw new Error(
					'La création automatique d’un nouveau UserForm est refusée sans designer .frx sûr.'
				);
			}
			await fs.promises.writeFile(sourcePath, source, {
				encoding: 'utf8',
				flag: 'wx'
			});
			await this.vbaWritebackService.prepare(context);
		}
		const result = await this.vbaWritebackService.applySource(
			context,
			file,
			source
		);
		if (persistSourceFile && sourceExists) {
			await fs.promises.writeFile(sourcePath, source, 'utf8');
		}
		return result;
	}

	async writeVbaFromTool(
		workbookUri: vscode.Uri,
		file: string,
		source: string,
		cancellationToken?: vscode.CancellationToken
	): Promise<VbaToolWriteResult> {
		if (cancellationToken?.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const canonicalUri = await canonicalizeWorkbookUri(workbookUri);
		assertNotManagedBackupPath(canonicalUri.fsPath);
		if (!(await this.ensureActiveWorkbookIsSaved(canonicalUri))) {
			throw new Error(
				'Le classeur contient des modifications non enregistrées et ne peut pas recevoir de code VBA.'
			);
		}

		const workbookExtension = path
			.extname(canonicalUri.fsPath)
			.toLocaleLowerCase('en-US');
		if (workbookExtension === '.xlsx') {
			return await this.bootstrapMacroWorkbook(
				canonicalUri,
				file,
				source,
				cancellationToken
			);
		}
		if (workbookExtension !== '.xlsm' && workbookExtension !== '.xlam') {
			throw new Error(
				'L’écriture VBA accepte un fichier .xlsm ou .xlam, ou un fichier .xlsx à convertir en copie .xlsm. Les formats .xls et .xlsb restent protégés.'
			);
		}

		const contextResult = await this.exportWorkbook(canonicalUri, {
			open: false,
			includeVba: true,
			requestedByTool: true,
			cancellationToken
		});
		if (!contextResult) {
			throw new Error('Le projet VBA n’a pas pu être préparé.');
		}
		const writeResult = await this.applyVbaSource(
			contextResult,
			file,
			source,
			true
		);
		return {
			targetWorkbookPath: canonicalUri.fsPath,
			sourceWorkbookPath: canonicalUri.fsPath,
			convertedToXlsm: false,
			changed: writeResult.changed,
			modifiedModules: writeResult.modifiedModules,
			workbookSha256: writeResult.workbookSha256,
			backupPath: writeResult.backupPath
		};
	}

	private async bootstrapMacroWorkbook(
		sourceUri: vscode.Uri,
		file: string,
		source: string,
		cancellationToken?: vscode.CancellationToken
	): Promise<VbaToolWriteResult> {
		const componentExtension = path
			.extname(file)
			.toLocaleLowerCase('en-US');
		if (componentExtension === '.frm') {
			throw new Error(
				'La création d’un nouveau UserForm est refusée : un vrai designer, ses contrôles et son fichier .frx ne peuvent pas être remplacés par un faux fichier .frm. Créez d’abord le UserForm dans le VBE natif.'
			);
		}
		if (
			path.basename(file) !== file ||
			(componentExtension !== '.bas' && componentExtension !== '.cls')
		) {
			throw new Error(
				'La première écriture dans un fichier .xlsx accepte uniquement un module .bas ou une classe .cls sans chemin.'
			);
		}

		const sourcePath = sourceUri.fsPath;
		const expectedTargetPath = path.join(
			path.dirname(sourcePath),
			`${path.basename(sourcePath, path.extname(sourcePath))}.xlsm`
		);
		assertNotManagedBackupPath(expectedTargetPath);
		await assertNoReparsePointChain(expectedTargetPath);

		const scriptPath = this.extensionContext.asAbsolutePath(
			path.join('scripts', 'prepare-macro-workbook.ps1')
		);
		const result = await this.runPowerShell(
			scriptPath,
			[
				'-WorkbookPathBase64',
				Buffer.from(sourcePath, 'utf8').toString('base64'),
				'-ComponentFileBase64',
				Buffer.from(file, 'utf8').toString('base64'),
				'-SourceBase64',
				Buffer.from(source, 'utf8').toString('base64')
			],
			path.dirname(sourcePath),
			{
				cancellationToken,
				timeoutMs: VBA_BOOTSTRAP_TIMEOUT_MS,
				cleanupOwnedExcel: true
			}
		);
		if (result.code !== 0) {
			throw new Error(
				processError(result.stderr, result.stdout) ||
					`La préparation XLSM a échoué avec le code ${result.code}.`
			);
		}

		const resultLine = result.stdout
			.replace(/\r/g, '')
			.split('\n')
			.map(line => line.trim())
			.filter(Boolean)
			.pop();
		if (!resultLine) {
			throw new Error('Le préparateur XLSM n’a renvoyé aucun résultat.');
		}
		let parsed: MacroBootstrapResult;
		try {
			parsed = JSON.parse(resultLine) as MacroBootstrapResult;
		} catch {
			throw new Error('Le préparateur XLSM a renvoyé un JSON invalide.');
		}
		if (
			parsed.ok !== true ||
			parsed.convertedToXlsm !== true ||
			parsed.changed !== true ||
			parsed.macrosExecuted !== false ||
			parsed.accessVbomChanged !== false ||
			typeof parsed.sourceWorkbookPath !== 'string' ||
			typeof parsed.targetWorkbookPath !== 'string' ||
			typeof parsed.workbookSha256 !== 'string' ||
			!/^[0-9a-f]{64}$/.test(parsed.workbookSha256) ||
			!Array.isArray(parsed.modifiedModules) ||
			parsed.modifiedModules.length !== 1 ||
			typeof parsed.modifiedModules[0] !== 'string' ||
			!parsed.modifiedModules[0]
		) {
			throw new Error('Le préparateur XLSM a renvoyé un résultat incomplet.');
		}
		if (!sameFile(parsed.sourceWorkbookPath, sourcePath)) {
			throw new Error('Le préparateur XLSM a confirmé un classeur source inattendu.');
		}
		if (!sameFile(parsed.targetWorkbookPath, expectedTargetPath)) {
			throw new Error('Le préparateur XLSM a confirmé un classeur cible inattendu.');
		}

		assertNotManagedBackupPath(parsed.targetWorkbookPath);
		await assertNoReparsePointChain(parsed.targetWorkbookPath);
		const targetUri = await canonicalizeWorkbookUri(
			vscode.Uri.file(parsed.targetWorkbookPath)
		);
		if (!sameFile(targetUri.fsPath, expectedTargetPath)) {
			throw new Error('Le chemin canonique du classeur XLSM est inattendu.');
		}
		const verifiedHash = await hashFileSha256(targetUri.fsPath);
		if (verifiedHash !== parsed.workbookSha256) {
			throw new Error('Le hash du classeur XLSM créé ne correspond pas au résultat.');
		}

		let contextResult: ExportContext | undefined;
		try {
			contextResult = await this.exportWorkbook(targetUri, {
				open: false,
				includeVba: true,
				requestedByTool: true,
				cancellationToken
			});
			if (!contextResult) {
				throw new Error('Le nouveau projet VBA n’a pas pu être relu.');
			}
			await this.vbaWritebackService.prepare(contextResult);
		} catch (error) {
			this.outputChannel.appendLine(
				`[vba bootstrap] Le fichier créé est conservé pour récupération : ${targetUri.fsPath}`
			);
			throw error;
		}

		return {
			targetWorkbookPath: targetUri.fsPath,
			sourceWorkbookPath: sourcePath,
			convertedToXlsm: true,
			changed: true,
			modifiedModules: parsed.modifiedModules as string[],
			workbookSha256: parsed.workbookSha256
		};
	}

	getSettings(): ExcelAiSettings {
		const configuration = vscode.workspace.getConfiguration('excelAiVbaStudio');
		return {
			maxRows: boundedInteger(
				configuration.get('maxRows', DEFAULT_MAX_ROWS),
				DEFAULT_MAX_ROWS,
				1,
				MAX_CONFIGURED_ROWS
			),
			maxColumns: boundedInteger(
				configuration.get('maxColumns', DEFAULT_MAX_COLUMNS),
				DEFAULT_MAX_COLUMNS,
				1,
				MAX_CONFIGURED_COLUMNS
			),
			includeVba: configuration.get<boolean>('includeVba', false) === true
		};
	}

	private candidateUri(candidate: unknown): vscode.Uri | undefined {
		if (isUri(candidate)) {
			return candidate;
		}
		if (candidate && typeof candidate === 'object') {
			const resourceUri = (candidate as { resourceUri?: unknown }).resourceUri;
			if (isUri(resourceUri)) {
				return resourceUri;
			}
		}
		return undefined;
	}

	private async findWorkspaceDefault(): Promise<vscode.Uri | undefined> {
		const matches: vscode.Uri[] = [];
		for (const folder of vscode.workspace.workspaceFolders || []) {
			if (folder.uri.scheme !== 'file' || folder.uri.authority) {
				continue;
			}
			try {
				// Reject mapped network workspaces before findFiles can enumerate
				// their contents.
				await assertLocalPath(folder.uri.fsPath);
			} catch {
				continue;
			}
			const remaining = 2 - matches.length;
			if (remaining <= 0) {
				break;
			}
			const folderMatches = await vscode.workspace.findFiles(
				new vscode.RelativePattern(folder, '**/*.{xlsx,xlsm,xls,xlsb}'),
				'**/{.git,node_modules}/**',
				remaining
			);
			matches.push(...folderMatches.filter(uri => isExcelUri(uri)));
		}
		return matches.length === 1 ? matches[0] : undefined;
	}

	async resolveWorkbookUri(candidate?: unknown): Promise<vscode.Uri | undefined> {
		const explicitUri = this.candidateUri(candidate);
		if (explicitUri && isExcelUri(explicitUri)) {
			return explicitUri;
		}

		const activeUri = getActiveResourceUri();
		if (activeUri && isExcelUri(activeUri)) {
			return activeUri;
		}

		if (
			this.lastContext &&
			activeUri?.scheme === 'file' &&
			(sameFile(activeUri.fsPath, this.lastContext.markdownUri.fsPath) ||
				sameFile(activeUri.fsPath, this.lastContext.jsonUri.fsPath))
		) {
			return this.lastContext.workbookUri;
		}

		return await this.findWorkspaceDefault();
	}

	async resolveToolWorkbookUri(input: ToolInput): Promise<vscode.Uri | undefined> {
		if (typeof input.workbookPath === 'string' && input.workbookPath.trim()) {
			return workbookUriFromPathInput(input.workbookPath);
		}
		return await this.resolveWorkbookUri();
	}

	private async ensureStorage(): Promise<{ storageRoot: string; exportsRoot: string }> {
		if (this.storageRoot && this.exportsRoot) {
			await assertNoReparsePointChain(this.storageRoot);
			await assertOwnedDirectory(this.exportsRoot, this.storageRoot);
			return { storageRoot: this.storageRoot, exportsRoot: this.exportsRoot };
		}

		const storageUri = this.extensionContext.globalStorageUri;
		let storagePath =
			storageUri.scheme === 'file' && !storageUri.authority
				? storageUri.fsPath
				: undefined;
		if (!storagePath && path.isAbsolute(this.extensionContext.globalStoragePath)) {
			storagePath = this.extensionContext.globalStoragePath;
			this.outputChannel.appendLine(
				`[stockage] URI ${storageUri.scheme} remplacée par le chemin local fourni par VS Code.`
			);
		}
		if (!storagePath) {
			const localAppData =
				process.env.LOCALAPPDATA?.trim() || process.env.APPDATA?.trim();
			if (localAppData) {
				storagePath = path.join(
					localAppData,
					'ExcelAiVbaStudio',
					'extension-storage'
				);
				this.outputChannel.appendLine(
					`[stockage] Repli local utilisé : ${storagePath}`
				);
			}
		}
		if (!storagePath) {
			throw new Error(
				'VS Code ne fournit aucun dossier de stockage local utilisable.'
			);
		}

		this.storageRoot = await ensureLocalDirectory(storagePath);
		this.exportsRoot = await ensureOwnedDirectory(
			path.join(this.storageRoot, 'workbook-exports'),
			this.storageRoot
		);
		if (!pathIsInside(this.exportsRoot, this.storageRoot)) {
			throw new Error('Le dossier d’export sort du stockage global de l’extension.');
		}
		return { storageRoot: this.storageRoot, exportsRoot: this.exportsRoot };
	}

	private async listVbaSourceUris(
		context: ExportContext
	): Promise<vscode.Uri[]> {
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const entries = await fs.promises.readdir(context.paths.vbaDirectory, {
			withFileTypes: true
		});
		return entries
			.filter(
				entry =>
					entry.isFile() &&
					['.bas', '.cls', '.frm', '.txt'].includes(
						path.extname(entry.name).toLocaleLowerCase('en-US')
					)
			)
			.sort((left, right) => left.name.localeCompare(right.name))
			.map(entry =>
				vscode.Uri.file(path.join(context.paths.vbaDirectory, entry.name))
			);
	}

	private async writeCopilotWorkspaceFiles(
		context: ExportContext
	): Promise<vscode.Uri> {
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const sourceUris = await this.listVbaSourceUris(context);
		const githubDirectory = await ensureOwnedDirectory(
			path.join(context.paths.vbaDirectory, '.github'),
			context.paths.vbaDirectory
		);
		const instructionsPath = path.join(
			githubDirectory,
			'copilot-instructions.md'
		);
		const projectPath = path.join(context.paths.vbaDirectory, 'VBA-PROJECT.md');
		await assertNoReparsePointChain(instructionsPath, githubDirectory);
		await assertNoReparsePointChain(projectPath, context.paths.vbaDirectory);

		const moduleList = sourceUris.length
			? sourceUris
					.map(
						uri =>
							`- [${path.basename(uri.fsPath)}](./${path.basename(uri.fsPath)})`
					)
					.join('\n')
			: '- Aucun module n’a pu être extrait.';
		const commonInstructions = [
			'Ce dossier est un projet VBA extrait localement par Excel AI & VBA Studio.',
			'Les fichiers .bas, .cls et .frm sont les sources à analyser et modifier.',
			'Ne jamais exécuter une macro.',
			'Préserver les déclarations Attribute, les signatures Public/Private et les événements Excel.',
			'Pour relire les données du classeur, utiliser l’outil #excelVbaWorkbook avec includeVba: true.',
			`Classeur source : ${context.workbookUri.fsPath}`,
			'Ne jamais utiliser un chemin contenant un composant exact .excel-ai-vba-backups : ce dossier contient uniquement des sauvegardes de récupération.',
			'Pour écrire un module .bas ou une classe .cls, appeler #excelVbaWriteModule. Si le classeur est .xlsx, la première écriture crée une nouvelle copie .xlsm voisine ; utiliser ensuite uniquement le targetWorkbookPath renvoyé.',
			'Un .frm existant dans un .xlsm peut recevoir du code, mais ne jamais inventer un nouveau .frm, un designer, des contrôles ou un .frx. MsgBox et InputBox ne sont pas des UserForms.',
			'Ne déclarer une écriture réussie qu’après le résultat de #excelVbaWriteModule, puis indiquer exactement targetWorkbookPath.',
			'Les fichiers de ce dossier restent une copie de travail tant qu’aucun outil d’écriture ou enregistrement synchronisé n’a confirmé la modification du classeur.'
		];
		await fs.promises.writeFile(
			instructionsPath,
			`${commonInstructions.join('\n')}\n`,
			'utf8'
		);
		await fs.promises.writeFile(
			projectPath,
			[
				`# VBAProject (${path.basename(context.workbookUri.fsPath)})`,
				'',
				'Ce dossier est ouvert comme une racine VS Code afin que GitHub Copilot puisse indexer les modules, classes, objets Excel et UserForms.',
				'',
				`- Classeur source : \`${context.workbookUri.fsPath}\``,
				'- Macros exécutées pendant l’extraction : **non**',
				'- Outils Copilot : `#excelVbaWorkbook` pour lire, `#excelVbaWriteModule` pour écrire',
				'',
				'## Composants',
				'',
				moduleList,
				'',
				'> Les modifications locales portent sur la copie de travail. Pour les appliquer au classeur, appeler `#excelVbaWriteModule` et reprendre le `targetWorkbookPath` confirmé par l’outil. Ne jamais cibler `.excel-ai-vba-backups`.'
			].join('\n'),
			'utf8'
		);
		return vscode.Uri.file(projectPath);
	}

	private async exposeVbaFolderToWorkspace(
		context: ExportContext
	): Promise<void> {
		const vbaUri = context.paths.vbaDirectoryUri;
		const alreadyVisible = (vscode.workspace.workspaceFolders || []).some(
			folder =>
				folder.uri.scheme === 'file' &&
				!folder.uri.authority &&
				(pathIsInside(vbaUri.fsPath, folder.uri.fsPath) ||
					sameFile(folder.uri.fsPath, vbaUri.fsPath))
		);
		if (alreadyVisible) {
			return;
		}
		const added = vscode.workspace.updateWorkspaceFolders(
			vscode.workspace.workspaceFolders?.length || 0,
			0,
			{
				uri: vbaUri,
				name: `VBA · ${context.paths.baseName}`
			}
		);
		if (!added) {
			this.outputChannel.appendLine(
				`[vba] Le dossier n’a pas pu être ajouté à l’espace de travail : ${vbaUri.fsPath}`
			);
		}
	}

	async getContextPaths(canonicalWorkbookUri: vscode.Uri): Promise<ExportPaths> {
		const { exportsRoot } = await this.ensureStorage();
		const extension = path.extname(canonicalWorkbookUri.fsPath);
		const baseName = safeFileStem(
			path.basename(canonicalWorkbookUri.fsPath, extension)
		);
		const pathHash = createHash('sha256')
			.update(canonicalWorkbookUri.fsPath.toLocaleLowerCase('en-US'))
			.digest('hex')
			.slice(0, 20);
		const outputDirectory = path.join(exportsRoot, `${baseName}-${pathHash}`);
		if (!pathIsInside(outputDirectory, exportsRoot)) {
			throw new Error('Le dossier calculé sort de la zone d’export autorisée.');
		}
		const markdownPath = path.join(outputDirectory, `${baseName}.md`);
		const jsonPath = path.join(outputDirectory, `${baseName}.json`);
		const vbaDirectory = path.join(outputDirectory, 'vba');
		return {
			workbookPath: canonicalWorkbookUri.fsPath,
			canonicalWorkbookPath: canonicalWorkbookUri.fsPath,
			baseName,
			outputDirectory,
			markdownPath,
			markdownUri: vscode.Uri.file(markdownPath),
			jsonPath,
			jsonUri: vscode.Uri.file(jsonPath),
			vbaDirectory,
			vbaDirectoryUri: vscode.Uri.file(vbaDirectory)
		};
	}

	private findOpenTabForUri(uri: vscode.Uri): { isDirty?: boolean } | undefined {
		const vscodeWindow = vscode.window as typeof vscode.window & {
			tabGroups?: {
				all?: Array<{ tabs?: Array<{ input?: unknown; isDirty?: boolean }> }>;
				activeTabGroup?: { activeTab?: { input?: unknown; isDirty?: boolean } };
			};
		};
		for (const group of vscodeWindow.tabGroups?.all || []) {
			for (const tab of group.tabs || []) {
				const tabUri = getUriFromTabInput(tab.input);
				if (
					tabUri?.scheme === 'file' &&
					sameFile(tabUri.fsPath, uri.fsPath)
				) {
					return tab;
				}
			}
		}
		return undefined;
	}

	private async ensureActiveWorkbookIsSaved(workbookUri: vscode.Uri): Promise<boolean> {
		const tab = this.findOpenTabForUri(workbookUri);
		if (!tab?.isDirty) {
			return true;
		}

		try {
			await vscode.commands.executeCommand('workbench.action.files.save');
		} catch (error) {
			this.outputChannel.appendLine(
				`[export] Échec de la sauvegarde préalable : ${String(error)}`
			);
		}

		for (let attempt = 0; attempt < 20; attempt++) {
			if (!this.findOpenTabForUri(workbookUri)?.isDirty) {
				return true;
			}
			await new Promise(resolve => setTimeout(resolve, 100));
		}

		await vscode.window.showWarningMessage(
			'Le classeur contient encore des modifications non enregistrées. La lecture est annulée pour éviter d’exporter une ancienne version.'
		);
		return false;
	}

	private appendProcessOutput(label: string, chunk: unknown): void {
		for (const line of String(chunk).replace(/\r/g, '').split('\n')) {
			if (line.trim()) {
				this.outputChannel.appendLine(`[${label}] ${line}`);
			}
		}
	}

	private async runPowerShell(
		scriptPath: string,
		scriptArguments: string[],
		cwd: string,
		options: PowerShellRunOptions
	): Promise<ProcessResult> {
		await fs.promises.access(scriptPath, fs.constants.R_OK);
		if (options.cancellationToken?.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const args = [
			'-NoLogo',
			'-NoProfile',
			'-NonInteractive',
			'-STA',
			'-ExecutionPolicy',
			'Bypass',
			'-File',
			scriptPath,
			...scriptArguments
		];

		return await new Promise<ProcessResult>((resolve, reject) => {
			const child = spawn(getPowerShellPath(), args, {
				cwd,
				windowsHide: true,
				shell: false,
				stdio: ['ignore', 'pipe', 'pipe']
			});
			let stdout = '';
			let stderr = '';
			let settled = false;
			let ownedExcelProcessId: number | undefined;
			let timeout: NodeJS.Timeout | undefined;
			let cancellationSubscription: vscode.Disposable | undefined;

			const cleanupListeners = () => {
				if (timeout) {
					clearTimeout(timeout);
					timeout = undefined;
				}
				cancellationSubscription?.dispose();
				cancellationSubscription = undefined;
			};
			const stopOwnedProcesses = async () => {
				const terminations: Promise<void>[] = [];
				if (typeof child.pid === 'number') {
					terminations.push(terminateExactProcess(child.pid));
				}
				if (options.cleanupOwnedExcel && ownedExcelProcessId) {
					terminations.push(terminateExactProcess(ownedExcelProcessId));
				}
				await Promise.all(terminations);
			};
			const abort = (error: Error) => {
				if (settled) {
					return;
				}
				settled = true;
				cleanupListeners();
				void stopOwnedProcesses().finally(() => reject(error));
			};

			child.stdout.setEncoding('utf8');
			child.stderr.setEncoding('utf8');
			child.stdout.on('data', chunk => {
				this.appendProcessOutput('PowerShell', chunk);
				const chunkText = String(chunk);
				stdout = (stdout + chunkText).slice(-32_000);
				const ownedProcessMatch = /OWNED_EXCEL_PID\|(\d+)/.exec(stdout);
				if (ownedProcessMatch) {
					const parsedProcessId = Number.parseInt(ownedProcessMatch[1], 10);
					if (Number.isSafeInteger(parsedProcessId) && parsedProcessId > 0) {
						ownedExcelProcessId = parsedProcessId;
					}
				}
				const message = chunkText
					.replace(/\r/g, '')
					.split('\n')
					.map(line => line.trim())
					.filter(line => Boolean(line) && !line.startsWith('OWNED_EXCEL_PID|'))
					.pop();
				if (message) {
					options.progress?.report({ message });
				}
			});
			child.stderr.on('data', chunk => {
				this.appendProcessOutput('PowerShell erreur', chunk);
				stderr = (stderr + String(chunk)).slice(-32_000);
			});
			child.once('error', error => {
				if (!settled) {
					settled = true;
					cleanupListeners();
					reject(
						new Error(`Impossible de démarrer Windows PowerShell : ${error.message}`)
					);
				}
			});
			child.once('close', (code, signal) => {
				if (!settled) {
					settled = true;
					cleanupListeners();
					resolve({
						code: typeof code === 'number' ? code : 1,
						signal,
						stdout: stdout.trim(),
						stderr: stderr.trim()
					});
				}
			});

			timeout = setTimeout(
				() =>
					abort(
						new Error(
							`L’opération PowerShell a dépassé le délai de ${Math.round(
								options.timeoutMs / 1000
							)} secondes et a été arrêtée.`
						)
					),
				options.timeoutMs
			);
			cancellationSubscription =
				options.cancellationToken?.onCancellationRequested(() =>
					abort(new vscode.CancellationError())
				);
			if (options.cancellationToken?.isCancellationRequested) {
				abort(new vscode.CancellationError());
			}
		});
	}

	private async runExporter(
		paths: ExportPaths,
		settings: ExcelAiSettings,
		includeVba: boolean,
		progress: vscode.Progress<{ message?: string }>,
		cancellationToken: vscode.CancellationToken
	): Promise<void> {
		const { exportsRoot } = await this.ensureStorage();
		await assertOwnedDirectory(paths.outputDirectory, exportsRoot);
		await assertOwnedDirectory(paths.vbaDirectory, paths.outputDirectory);
		const scriptPath = this.extensionContext.asAbsolutePath(
			path.join('scripts', 'office-ai-export.ps1')
		);
		this.outputChannel.appendLine('');
		this.outputChannel.appendLine(
			`[export] ${paths.workbookPath} -> ${paths.outputDirectory}`
		);
		const result = await this.runPowerShell(
			scriptPath,
			[
				'-WorkbookPath',
				paths.workbookPath,
				'-StorageRoot',
				exportsRoot,
				'-OutputPath',
				paths.markdownPath,
				'-JsonOutputPath',
				paths.jsonPath,
				'-VbaOutputDirectory',
				paths.vbaDirectory,
				'-MaxRows',
				String(settings.maxRows),
				'-MaxColumns',
				String(settings.maxColumns),
				'-IncludeVba',
				includeVba ? 'true' : 'false'
			],
			paths.outputDirectory,
			{
				progress,
				cancellationToken,
				timeoutMs: EXPORT_TIMEOUT_MS,
				cleanupOwnedExcel: true
			}
		);
		if (result.code !== 0) {
			throw new Error(
				processError(result.stderr, result.stdout) ||
					`L’export PowerShell a échoué avec le code ${result.code}.`
			);
		}
	}

	private async validateContextOutputs(paths: ExportPaths): Promise<void> {
		const { exportsRoot } = await this.ensureStorage();
		await assertOwnedDirectory(paths.outputDirectory, exportsRoot);
		await assertNoReparsePointChain(paths.markdownPath, paths.outputDirectory);
		await assertNoReparsePointChain(paths.jsonPath, paths.outputDirectory);
		const [markdownText, jsonText] = await Promise.all([
			fs.promises.readFile(paths.markdownPath, 'utf8'),
			fs.promises.readFile(paths.jsonPath, 'utf8')
		]);
		if (!markdownText.trim()) {
			throw new Error('Le contexte Markdown généré est vide.');
		}
		if (
			Buffer.byteLength(markdownText, 'utf8') > MAX_GENERATED_FILE_BYTES ||
			Buffer.byteLength(jsonText, 'utf8') > MAX_GENERATED_FILE_BYTES
		) {
			throw new Error(
				'Le contexte généré dépasse la limite de sécurité de 16 Mio par fichier. Réduisez maxRows ou maxColumns.'
			);
		}
		JSON.parse(jsonText.replace(/^\uFEFF/, ''));
	}

	async exportWorkbook(
		candidate?: unknown,
		options: ExportOptions = {}
	): Promise<ExportContext | undefined> {
		if (options.cancellationToken?.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const includeVba =
			typeof options.includeVba === 'boolean'
				? options.includeVba
				: this.getSettings().includeVba;
		const requestedUri = await this.resolveWorkbookUri(candidate);
		if (!requestedUri) {
			const message =
				'Aucun classeur Excel local n’est actif. Ouvrez un classeur ou placez un seul classeur dans l’espace de travail.';
			if (options.requestedByTool) {
				throw new Error(message);
			}
			await vscode.window.showWarningMessage(message);
			return undefined;
		}

		let canonicalUri: vscode.Uri | undefined;
		try {
			canonicalUri = await canonicalizeWorkbookUri(requestedUri);
		} catch (error) {
			if (options.requestedByTool) {
				throw error;
			}
			await vscode.window.showErrorMessage(
				`Lecture du classeur refusée : ${(error as Error).message}`
			);
			return undefined;
		}
		if (!canonicalUri || !(await this.ensureActiveWorkbookIsSaved(canonicalUri))) {
			return undefined;
		}

		const paths = await this.getContextPaths(canonicalUri);
		const exportKey = `${paths.outputDirectory}|${includeVba ? 'vba' : 'data'}`.toLocaleLowerCase(
			'en-US'
		);
		let runningExport = this.runningExports.get(exportKey);
		if (!runningExport) {
			runningExport = (async () => {
				const { exportsRoot } = await this.ensureStorage();
				await ensureOwnedDirectory(paths.outputDirectory, exportsRoot);
				// Keep this directory stable once VS Code/Copilot indexes it.
				// The exporter replaces only its managed artifacts in place.
				await ensureOwnedDirectory(paths.vbaDirectory, paths.outputDirectory);
				const settings = this.getSettings();
				await vscode.window.withProgress(
					{
						location: vscode.ProgressLocation.Notification,
						title: `Excel AI & VBA : export de ${path.basename(paths.workbookPath)}`,
						cancellable: true
					},
					async (progress, progressCancellationToken) => {
						const linkedCancellation =
							new vscode.CancellationTokenSource();
						const subscriptions: vscode.Disposable[] = [
							progressCancellationToken.onCancellationRequested(() =>
								linkedCancellation.cancel()
							)
						];
						if (options.cancellationToken) {
							subscriptions.push(
								options.cancellationToken.onCancellationRequested(() =>
									linkedCancellation.cancel()
								)
							);
							if (options.cancellationToken.isCancellationRequested) {
								linkedCancellation.cancel();
							}
						}
						progress.report({
							message: includeVba
								? 'Lecture des données et du projet VBA autorisé…'
								: 'Lecture des valeurs, formules et métadonnées…'
						});
						try {
							await this.runExporter(
								paths,
								settings,
								includeVba,
								progress,
								linkedCancellation.token
							);
						} finally {
							for (const subscription of subscriptions) {
								subscription.dispose();
							}
							linkedCancellation.dispose();
						}
					}
				);
				await this.validateContextOutputs(paths);
				return {
					workbookUri: canonicalUri as vscode.Uri,
					markdownUri: paths.markdownUri,
					jsonUri: paths.jsonUri,
					paths,
					includeVba
				};
			})();
			this.runningExports.set(exportKey, runningExport);
			void runningExport
				.finally(() => this.runningExports.delete(exportKey))
				.catch(() => undefined);
		}

		try {
			const result = await runningExport;
			if (!result) {
				return undefined;
			}
			this.lastContext = result;
			this.contextChangeEmitter.fire();
			if (options.open) {
				await this.openMarkdown(result.markdownUri);
			}
			return result;
		} catch (error) {
			const typedError = error as Error;
			if (error instanceof vscode.CancellationError) {
				if (options.requestedByTool) {
					throw error;
				}
				await vscode.window.showInformationMessage(
					'L’export Excel a été annulé.'
				);
				return undefined;
			}
			this.outputChannel.appendLine(
				`[export] ERREUR : ${typedError.stack || typedError.message}`
			);
			if (options.requestedByTool) {
				throw typedError;
			}
			const choice = await vscode.window.showErrorMessage(
				`L’export du classeur a échoué : ${typedError.message}`,
				'Voir le journal'
			);
			if (choice === 'Voir le journal') {
				this.outputChannel.show(true);
			}
			return undefined;
		}
	}

	private async openMarkdown(markdownUri: vscode.Uri): Promise<void> {
		const document = await vscode.workspace.openTextDocument(markdownUri);
		await vscode.window.showTextDocument(document, {
			preview: false,
			preserveFocus: false
		});
	}

	async copyGeneratedContext(silent = false): Promise<boolean> {
		const context = this.lastContext;
		if (!context) {
			if (!silent) {
				await vscode.window.showWarningMessage(
					'Aucun contexte Excel généré par cette session n’est disponible.'
				);
			}
			return false;
		}

		await canonicalizeWorkbookUri(context.workbookUri);
		const { exportsRoot } = await this.ensureStorage();
		await assertOwnedDirectory(context.paths.outputDirectory, exportsRoot);
		await assertNoReparsePointChain(
			context.markdownUri.fsPath,
			context.paths.outputDirectory
		);
		const markdownPath = await fs.promises.realpath(context.markdownUri.fsPath);
		if (!pathIsInside(markdownPath, exportsRoot)) {
			throw new Error('Le contexte généré sort du stockage autorisé.');
		}
		const markdown = (await fs.promises.readFile(markdownPath, 'utf8')).replace(
			/^\uFEFF/,
			''
		);
		if (!markdown.trim()) {
			throw new Error('Le contexte Markdown est vide.');
		}
		await vscode.env.clipboard.writeText(markdown);
		if (!silent) {
			await vscode.window.showInformationMessage(
				'Le contexte Excel a été copié dans le presse-papiers.'
			);
		}
		return true;
	}

	async copyExportResult(
		context: ExportContext,
		silent = false
	): Promise<boolean> {
		const markdown = await this.readExportedContext(context, 'markdown');
		if (!markdown.trim()) {
			throw new Error('Le contexte Markdown est vide.');
		}
		await vscode.env.clipboard.writeText(markdown);
		if (!silent) {
			await vscode.window.showInformationMessage(
				'Le contexte Excel a été copié dans le presse-papiers.'
			);
		}
		return true;
	}

	async openWorkbookContext(candidate?: unknown): Promise<void> {
		const requestedUri = await this.resolveWorkbookUri(candidate);
		if (
			this.lastContext &&
			requestedUri &&
			sameFile(requestedUri.fsPath, this.lastContext.workbookUri.fsPath)
		) {
			await canonicalizeWorkbookUri(this.lastContext.workbookUri);
			await this.validateContextOutputs(this.lastContext.paths);
			await this.openMarkdown(this.lastContext.markdownUri);
			return;
		}
		await this.exportWorkbook(requestedUri, { open: true, includeVba: false });
	}

	private async runExcelLauncher(
		workbookPath: string,
		showVbe: boolean,
		cancellationToken: vscode.CancellationToken
	): Promise<ProcessResult> {
		const scriptPath = this.extensionContext.asAbsolutePath(
			path.join('scripts', 'open-excel-developer.ps1')
		);
		return await this.runPowerShell(
			scriptPath,
			['-WorkbookPath', workbookPath, ...(showVbe ? ['-ShowVbe'] : [])],
			path.dirname(workbookPath),
			{
				cancellationToken,
				timeoutMs: EXCEL_LAUNCH_TIMEOUT_MS,
				cleanupOwnedExcel: false
			}
		);
	}

	async openExcel(candidate?: unknown, showVbe = false): Promise<boolean> {
		const requestedUri = await this.resolveWorkbookUri(candidate);
		if (!requestedUri) {
			await vscode.window.showWarningMessage('Ouvrez d’abord un classeur Excel local.');
			return false;
		}

		let canonicalUri: vscode.Uri | undefined;
		try {
			canonicalUri = await canonicalizeWorkbookUri(requestedUri);
		} catch (error) {
			await vscode.window.showErrorMessage((error as Error).message);
			return false;
		}
		if (!canonicalUri || !(await this.ensureActiveWorkbookIsSaved(canonicalUri))) {
			return false;
		}

		try {
			const result = await vscode.window.withProgress(
				{
					location: vscode.ProgressLocation.Notification,
					title: showVbe
						? `Ouverture du mode Développeur pour ${path.basename(
								canonicalUri.fsPath
						  )}`
						: `Ouverture de ${path.basename(canonicalUri.fsPath)} dans Excel`,
					cancellable: true
				},
				(_progress, cancellationToken) =>
					this.runExcelLauncher(
						(canonicalUri as vscode.Uri).fsPath,
						showVbe,
						cancellationToken
					)
			);
			if (result.code === 0) {
				if (showVbe) {
					await vscode.window.showInformationMessage(
						'Microsoft Excel et l’éditeur VBA sont ouverts.'
					);
				}
				return true;
			}
			if (showVbe && result.code === 3) {
				await vscode.window.showWarningMessage(
					'Le classeur est ouvert dans Excel, mais le VBE n’a pas pu être affiché automatiquement. Utilisez Alt+F11 dans Excel.'
				);
				return false;
			}
			throw new Error(
				processError(result.stderr, result.stdout) ||
					`Le lanceur Excel a échoué avec le code ${result.code}.`
			);
		} catch (error) {
			if (error instanceof vscode.CancellationError) {
				await vscode.window.showInformationMessage(
					'L’attente d’Excel a été annulée. Excel reste ouvert s’il a déjà démarré.'
				);
				return false;
			}
			const typedError = error as Error;
			this.outputChannel.appendLine(`[excel] ERREUR : ${typedError.message}`);
			const choice = await vscode.window.showErrorMessage(
				`Impossible d’ouvrir Microsoft Excel : ${typedError.message}`,
				'Voir le journal'
			);
			if (choice === 'Voir le journal') {
				this.outputChannel.show(true);
			}
			return false;
		}
	}

	async openVbaExplorer(candidate?: unknown): Promise<void> {
		const result = await this.exportWorkbook(candidate, {
			open: false,
			includeVba: true
		});
		if (!result) {
			return;
		}
		let projectUri: vscode.Uri | undefined;
		try {
			await this.vbaStudioPanel.prepare(result);
			projectUri = await this.writeCopilotWorkspaceFiles(result);
			await this.exposeVbaFolderToWorkspace(result);
			await this.vbaStudioPanel.open(result);
		} catch (error) {
			this.outputChannel.appendLine(
				`[vba] Préparation de l’espace VS Code incomplète : ${(error as Error).message}`
			);
			if (projectUri) {
				const document = await vscode.workspace.openTextDocument(projectUri);
				await vscode.window.showTextDocument(document, {
					preview: false,
					preserveFocus: false
				});
			}
		}
	}

	async openVbaComponent(candidate?: unknown): Promise<void> {
		const uri = isUri(candidate)
			? candidate
			: this.candidateUri(candidate);
		const context = this.lastContext;
		if (!uri || uri.scheme !== 'file' || uri.authority || !context) {
			return;
		}
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		await assertNoReparsePointChain(uri.fsPath, context.paths.vbaDirectory);
		if (!pathIsInside(uri.fsPath, context.paths.vbaDirectory)) {
			throw new Error('Le composant VBA demandé sort du projet extrait.');
		}
		const document = await vscode.workspace.openTextDocument(uri);
		await vscode.window.showTextDocument(document, {
			preview: false,
			preserveFocus: false,
			viewColumn: vscode.ViewColumn.Active
		});
		if (
			path.extname(uri.fsPath).toLocaleLowerCase('en-US') === '.frm'
		) {
			await showUserFormPreview(uri, document.getText());
		}
	}

	async askCopilotAboutWorkbook(candidate?: unknown): Promise<void> {
		const workbookUri = await this.resolveWorkbookUri(candidate);
		if (!workbookUri) {
			await vscode.window.showWarningMessage(
				'Aucun classeur Excel local actif n’a été trouvé.'
			);
			return;
		}
		const result = await this.exportWorkbook(workbookUri, {
			open: false,
			includeVba: true
		});
		if (!result) {
			return;
		}
		await this.vbaStudioPanel.prepare(result);
		await this.writeCopilotWorkspaceFiles(result);
		await this.exposeVbaFolderToWorkspace(result);
		const requestedTask =
			candidate &&
			typeof candidate === 'object' &&
			typeof (candidate as { request?: unknown }).request === 'string'
				? (candidate as { request: string }).request.trim().slice(0, 4000)
				: '';
		const prompt = [
			'Utilise #excelVbaWorkbook pour analyser ce classeur local.',
			`workbookPath: ${workbookUri.fsPath}`,
			'includeVba: true',
			`Le projet VBA est aussi disponible comme dossier VS Code : ${result.paths.vbaDirectory}`,
			'Lis les fichiers .bas, .cls et .frm du dossier VBA avant de proposer des modifications.',
			'Commence par résumer les objets Excel, les modules, les classes et les UserForms.',
			'Ne cible jamais un chemin contenant le composant exact .excel-ai-vba-backups.',
			'Pour appliquer un .bas ou .cls, utilise #excelVbaWriteModule. Sur un .xlsx, reprends ensuite le targetWorkbookPath .xlsm renvoyé pour toutes les écritures suivantes.',
			'Ne crée jamais de faux UserForm .frm ou .frx. Un nouveau UserForm avec designer doit être créé dans le VBE natif ; MsgBox et InputBox ne sont pas des UserForms.',
			'Ne confirme une modification du classeur qu’après le succès de #excelVbaWriteModule et donne le targetWorkbookPath exact.',
			...(requestedTask ? [`Tâche demandée depuis le ruban : ${requestedTask}`] : []),
			'N’exécute aucune macro : analyse uniquement le code et les données.'
		].join('\n');
		try {
			await vscode.commands.executeCommand('workbench.action.chat.open', {
				query: prompt
			});
		} catch {
			await vscode.env.clipboard.writeText(prompt);
			await vscode.commands.executeCommand('workbench.action.chat.open');
			await vscode.window.showInformationMessage(
				'La demande Excel/VBA a été copiée. Collez-la dans Copilot Chat.'
			);
		}
	}

	async readExportedContext(
		context: ExportContext,
		format: 'markdown' | 'json'
	): Promise<string> {
		const requestedPath =
			format === 'json' ? context.paths.jsonPath : context.paths.markdownPath;
		const { exportsRoot } = await this.ensureStorage();
		await assertOwnedDirectory(context.paths.outputDirectory, exportsRoot);
		await assertNoReparsePointChain(
			requestedPath,
			context.paths.outputDirectory
		);
		const canonicalPath = await fs.promises.realpath(requestedPath);
		if (!pathIsInside(canonicalPath, exportsRoot)) {
			throw new Error('Le contexte demandé sort du stockage global autorisé.');
		}
		return (await fs.promises.readFile(canonicalPath, 'utf8')).replace(
			/^\uFEFF/,
			''
		);
	}

	async cleanExports(): Promise<void> {
		const { storageRoot, exportsRoot } = await this.ensureStorage();
		if (!pathIsInside(exportsRoot, storageRoot) || exportsRoot === storageRoot) {
			throw new Error('Le dossier d’export à nettoyer n’est pas sûr.');
		}
		const choice = await vscode.window.showWarningMessage(
			`Supprimer tous les contextes Excel générés ?\n\n${exportsRoot}`,
			{
				modal: true,
				detail: `Seul ce dossier du stockage global de l’extension sera supprimé : ${exportsRoot}`
			},
			'Supprimer les exports'
		);
		if (choice !== 'Supprimer les exports') {
			return;
		}
		await removeOwnedDirectory(exportsRoot, storageRoot);
		this.exportsRoot = await ensureOwnedDirectory(exportsRoot, storageRoot);
		this.lastContext = undefined;
		this.contextChangeEmitter.fire();
		await vscode.window.showInformationMessage('Les contextes Excel générés ont été supprimés.');
	}
}
