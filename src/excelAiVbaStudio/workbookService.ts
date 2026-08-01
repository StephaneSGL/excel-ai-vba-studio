import { createHash, randomUUID } from 'crypto';
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { assertOoxmlPackageUnsignedForMutation } from '../common/ooxmlPackageSignature';
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
import { OfficeSecurityService } from './officeSecurity';
import { SecurityCenterPanel } from './securityCenterPanel';
import {
	EXCEL_EXTENSIONS,
	ExcelAiSettings,
	ExportContext,
	ExportOptions,
	ExportPaths,
	ProcessResult,
	ToolInput,
	VbaDesignToolInput,
	VbaDesignToolResult,
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
const VBA_DESIGN_TIMEOUT_MS = 120_000;
const MAX_VBA_DESIGN_REQUEST_BYTES = 1024 * 1024;
const MAX_CUSTOM_ACTIVEX_PROGIDS = 32;
const ACTIVEX_PROGID_PATTERN = /^[A-Za-z][A-Za-z0-9_.-]{1,127}$/;
const DEVELOPER_MARKER_NAME = '.excel-ai-vba-studio-project.json';
const MAX_DEVELOPER_MARKER_BYTES = 64 * 1024;

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

interface VbaDesignerProcessResult {
	ok?: unknown;
	targetWorkbookPath?: unknown;
	sourceWorkbookPath?: unknown;
	convertedToXlsm?: unknown;
	changed?: unknown;
	createdUserForms?: unknown;
	addedControls?: unknown;
	updatedControls?: unknown;
	updatedEventHandlers?: unknown;
	createdButtons?: unknown;
	assignedButtons?: unknown;
	createdActiveXControls?: unknown;
	boundActiveXControls?: unknown;
	workbookSha256?: unknown;
	backupPath?: unknown;
	macrosExecuted?: unknown;
	accessVbomChanged?: unknown;
	designerVerified?: unknown;
}

interface DeveloperWorkspaceMarker {
	schemaVersion: 1;
	workbookPath: string;
	workbookSha256: string;
	baseName: string;
	outputDirectory: string;
	markdownPath: string;
	jsonPath: string;
	vbaDirectory: string;
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

function isStringArray(value: unknown): value is string[] {
	return (
		Array.isArray(value) &&
		value.every(item => typeof item === 'string' && item.length > 0)
	);
}

function allowedCustomActiveXProgIds(resource: vscode.Uri): string[] {
	const configured = vscode.workspace
		.getConfiguration('excelAiVbaStudio', resource)
		.get<unknown>('allowedCustomActiveXProgIds', []);
	if (!Array.isArray(configured) || configured.length > MAX_CUSTOM_ACTIVEX_PROGIDS) {
		throw new Error(
			`excelAiVbaStudio.allowedCustomActiveXProgIds doit contenir au maximum ${MAX_CUSTOM_ACTIVEX_PROGIDS} ProgID.`
		);
	}
	const unique = new Set<string>();
	for (const value of configured) {
		if (typeof value !== 'string' || !ACTIVEX_PROGID_PATTERN.test(value)) {
			throw new Error(
				'excelAiVbaStudio.allowedCustomActiveXProgIds contient un ProgID invalide.'
			);
		}
		const key = value.toLocaleLowerCase('en-US');
		if (unique.has(key)) {
			throw new Error(
				'excelAiVbaStudio.allowedCustomActiveXProgIds contient un doublon.'
			);
		}
		unique.add(key);
	}
	return configured as string[];
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
	private readonly officeSecurityService: OfficeSecurityService;
	private readonly securityCenterPanel: SecurityCenterPanel;
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
			this.vbaWritebackService,
			(workbookUri, operations) =>
				this.designVbaFromTool(workbookUri, operations),
			async (workbookUri, showVbe) => {
				await this.openExcel(workbookUri, showVbe);
			}
		);
		this.officeSecurityService = new OfficeSecurityService(
			this.extensionContext,
			this.outputChannel
		);
		this.securityCenterPanel = new SecurityCenterPanel(
			this.extensionContext,
			this.officeSecurityService,
			this.outputChannel
		);
	}

	dispose(): void {
		this.securityCenterPanel.dispose();
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
		await assertOoxmlPackageUnsignedForMutation(canonicalUri.fsPath);
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

	async designVbaFromTool(
		workbookUri: vscode.Uri,
		operations: VbaDesignToolInput['operations'],
		cancellationToken?: vscode.CancellationToken
	): Promise<VbaDesignToolResult> {
		if (cancellationToken?.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const canonicalUri = await canonicalizeWorkbookUri(workbookUri);
		assertNotManagedBackupPath(canonicalUri.fsPath);
		await assertOoxmlPackageUnsignedForMutation(canonicalUri.fsPath);
		if (!(await this.ensureActiveWorkbookIsSaved(canonicalUri))) {
			throw new Error(
				'Le classeur contient des modifications non enregistrées et ne peut pas recevoir de composants visuels VBA.'
			);
		}
		if (
			path.extname(canonicalUri.fsPath).toLocaleLowerCase('en-US') !==
			'.xlsm'
		) {
			throw new Error(
				'La création de UserForms et de boutons accepte uniquement un classeur .xlsm existant.'
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

		const expectedWorkbookSha256 = await hashFileSha256(canonicalUri.fsPath);
		await assertOoxmlPackageUnsignedForMutation(canonicalUri.fsPath);
		const { exportsRoot } = await this.ensureStorage();
		const requestDirectory = await assertOwnedDirectory(
			contextResult.paths.outputDirectory,
			exportsRoot
		);
		const requestPath = path.join(
			requestDirectory,
			`vba-designer-request-${randomUUID()}.json`
		);
		await assertNoReparsePointChain(requestPath, requestDirectory);
		const requestJson = JSON.stringify({
			schemaVersion: 2,
			workbookPath: canonicalUri.fsPath,
			expectedWorkbookSha256,
			allowedCustomActiveXProgIds: allowedCustomActiveXProgIds(canonicalUri),
			operations
		});
		if (
			Buffer.byteLength(requestJson, 'utf8') >
			MAX_VBA_DESIGN_REQUEST_BYTES
		) {
			throw new Error('La demande de création VBA dépasse la limite de 1 Mio.');
		}

		const scriptPath = this.extensionContext.asAbsolutePath(
			path.join('scripts', 'apply-vba-designer.ps1')
		);
		const helperPath = this.extensionContext.asAbsolutePath(
			path.join('bin', 'win32-x64', 'excel-ai-vba-writeback.exe')
		);
		await fs.promises.access(helperPath, fs.constants.R_OK);
		await fs.promises.writeFile(requestPath, requestJson, {
			encoding: 'utf8',
			flag: 'wx'
		});

		try {
			const processResult = await this.runPowerShell(
				scriptPath,
				[
					'-RequestPathBase64',
					Buffer.from(requestPath, 'utf8').toString('base64'),
					'-HelperPathBase64',
					Buffer.from(helperPath, 'utf8').toString('base64')
				],
				path.dirname(canonicalUri.fsPath),
				{
					cancellationToken,
					timeoutMs: VBA_DESIGN_TIMEOUT_MS,
					cleanupOwnedExcel: true
				}
			);
			if (processResult.code !== 0) {
				throw new Error(
					processError(processResult.stderr, processResult.stdout) ||
						`La création des composants visuels VBA a échoué avec le code ${processResult.code}.`
				);
			}

			const resultLine = processResult.stdout
				.replace(/\r/g, '')
				.split('\n')
				.map(line => line.trim())
				.filter(Boolean)
				.pop();
			if (!resultLine) {
				throw new Error('Le moteur VBA Designer n’a renvoyé aucun résultat.');
			}
			let parsed: VbaDesignerProcessResult;
			try {
				parsed = JSON.parse(resultLine) as VbaDesignerProcessResult;
			} catch {
				throw new Error('Le moteur VBA Designer a renvoyé un JSON invalide.');
			}

			if (
				parsed.ok !== true ||
				parsed.convertedToXlsm !== false ||
				parsed.changed !== true ||
				parsed.macrosExecuted !== false ||
				parsed.accessVbomChanged !== false ||
				parsed.designerVerified !== true ||
				typeof parsed.targetWorkbookPath !== 'string' ||
				typeof parsed.sourceWorkbookPath !== 'string' ||
				typeof parsed.workbookSha256 !== 'string' ||
				!/^[0-9a-f]{64}$/.test(parsed.workbookSha256) ||
				typeof parsed.backupPath !== 'string' ||
				!isStringArray(parsed.createdUserForms) ||
				!isStringArray(parsed.addedControls) ||
				!isStringArray(parsed.updatedControls) ||
				!isStringArray(parsed.updatedEventHandlers) ||
				!isStringArray(parsed.createdButtons) ||
				!isStringArray(parsed.assignedButtons) ||
				!isStringArray(parsed.createdActiveXControls) ||
				!isStringArray(parsed.boundActiveXControls)
			) {
				throw new Error('Le moteur VBA Designer a renvoyé un résultat incomplet.');
			}
			if (
				!sameFile(parsed.targetWorkbookPath, canonicalUri.fsPath) ||
				!sameFile(parsed.sourceWorkbookPath, canonicalUri.fsPath)
			) {
				throw new Error('Le moteur VBA Designer a confirmé un classeur inattendu.');
			}

			const actualWorkbookSha256 = await hashFileSha256(canonicalUri.fsPath);
			if (actualWorkbookSha256 !== parsed.workbookSha256) {
				throw new Error(
					'Le hash du classeur modifié ne correspond pas au résultat du moteur VBA Designer.'
				);
			}

			const backupPath = path.resolve(parsed.backupPath);
			const expectedBackupDirectory = path.join(
				path.dirname(canonicalUri.fsPath),
				'.excel-ai-vba-backups'
			);
			if (
				!path.isAbsolute(parsed.backupPath) ||
				!sameFile(path.dirname(backupPath), expectedBackupDirectory) ||
				path.extname(backupPath).toLocaleLowerCase('en-US') !== '.xlsm'
			) {
				throw new Error('Le chemin de sauvegarde VBA Designer est inattendu.');
			}
			await assertNoReparsePointChain(backupPath);
			const backupStat = await fs.promises.lstat(backupPath);
			if (!backupStat.isFile() || backupStat.isSymbolicLink()) {
				throw new Error('La sauvegarde VBA Designer n’est pas un fichier sûr.');
			}
			if ((await hashFileSha256(backupPath)) !== expectedWorkbookSha256) {
				throw new Error(
					'La sauvegarde VBA Designer ne correspond pas au classeur d’origine.'
				);
			}

			const refreshedContext = await this.exportWorkbook(canonicalUri, {
				open: false,
				includeVba: true,
				requestedByTool: true,
				cancellationToken
			});
			if (!refreshedContext) {
				throw new Error(
					'Le classeur modifié existe, mais son contexte VBA n’a pas pu être actualisé.'
				);
			}

			return {
				targetWorkbookPath: canonicalUri.fsPath,
				sourceWorkbookPath: canonicalUri.fsPath,
				convertedToXlsm: false,
				changed: true,
				createdUserForms: parsed.createdUserForms,
				addedControls: parsed.addedControls,
				updatedControls: parsed.updatedControls,
				updatedEventHandlers: parsed.updatedEventHandlers,
				createdButtons: parsed.createdButtons,
				assignedButtons: parsed.assignedButtons,
				createdActiveXControls: parsed.createdActiveXControls,
				boundActiveXControls: parsed.boundActiveXControls,
				workbookSha256: parsed.workbookSha256,
				backupPath,
				macrosExecuted: false,
				accessVbomChanged: false,
				designerVerified: true
			};
		} finally {
			await fs.promises.unlink(requestPath).catch(error => {
				if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
					this.outputChannel.appendLine(
						`[vba designer] Nettoyage de la demande impossible : ${String(error)}`
					);
				}
			});
		}
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
		await assertOoxmlPackageUnsignedForMutation(sourcePath);
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
			'Un .frm existant peut recevoir du code via #excelVbaWriteModule. Pour créer un vrai UserForm, ajouter ou repositionner ses contrôles, affecter une procédure événementielle complète, créer un bouton de feuille ou un ActiveX autorisé, utiliser #excelVbaDesignWorkbook ; ne jamais fabriquer un faux .frm/.frx.',
			'Dans #excelVbaDesignWorkbook, setUserFormEventHandler accepte une unique procédure Private Sub objectName_eventName(...). Les signatures complexes avec paramètres sont autorisées ; replaceExisting doit être explicitement vrai pour remplacer un gestionnaire existant.',
			'Un customActiveX exige un ProgID déjà présent dans excelAiVbaStudio.allowedCustomActiveXProgIds. Ne jamais proposer de modifier cette liste sans demande explicite de l’utilisateur.',
			'Ne déclarer une écriture réussie qu’après le résultat de #excelVbaWriteModule ou #excelVbaDesignWorkbook, puis indiquer exactement targetWorkbookPath.',
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
				'- Outils Copilot : `#excelVbaWorkbook` pour lire, `#excelVbaWriteModule` pour le code et `#excelVbaDesignWorkbook` pour les UserForms/contrôles/boutons',
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
				const ownedProcessMatches = [
					...stdout.matchAll(/OWNED_EXCEL_PID\|(\d+)/g)
				];
				const latestOwnedProcessMatch = ownedProcessMatches.at(-1);
				if (latestOwnedProcessMatch) {
					const parsedProcessId = Number.parseInt(
						latestOwnedProcessMatch[1],
						10
					);
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

	async openSecurityCenter(candidate?: unknown): Promise<void> {
		const requestedUri = await this.resolveWorkbookUri(candidate);
		if (!requestedUri) {
			await vscode.window.showWarningMessage(
				'Ouvrez d’abord un classeur Excel local.'
			);
			return;
		}
		try {
			const canonicalUri = await canonicalizeWorkbookUri(requestedUri);
			await this.securityCenterPanel.open(canonicalUri);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			this.outputChannel.appendLine(`[security] ERREUR : ${message}`);
			await vscode.window.showErrorMessage(
				`Centre de sécurité Excel : ${message}`
			);
		}
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

	private async writeDeveloperWorkspaceMarker(
		context: ExportContext
	): Promise<vscode.Uri> {
		const { exportsRoot } = await this.ensureStorage();
		await assertOwnedDirectory(context.paths.outputDirectory, exportsRoot);
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const markerPath = path.join(
			context.paths.vbaDirectory,
			DEVELOPER_MARKER_NAME
		);
		await assertNoReparsePointChain(
			markerPath,
			context.paths.vbaDirectory
		);
		const marker: DeveloperWorkspaceMarker = {
			schemaVersion: 1,
			workbookPath: context.workbookUri.fsPath,
			workbookSha256: await hashFileSha256(context.workbookUri.fsPath),
			baseName: context.paths.baseName,
			outputDirectory: context.paths.outputDirectory,
			markdownPath: context.paths.markdownPath,
			jsonPath: context.paths.jsonPath,
			vbaDirectory: context.paths.vbaDirectory
		};
		const serialized = JSON.stringify(marker, null, 2);
		if (Buffer.byteLength(serialized, 'utf8') > MAX_DEVELOPER_MARKER_BYTES) {
			throw new Error('Le marqueur du mode Développeur dépasse 64 Kio.');
		}
		await fs.promises.writeFile(markerPath, serialized, {
			encoding: 'utf8',
			flag: 'w'
		});
		return vscode.Uri.file(context.paths.vbaDirectory);
	}

	async openVbaDeveloperWindow(candidate?: unknown): Promise<void> {
		const result = await this.exportWorkbook(candidate, {
			open: false,
			includeVba: true
		});
		if (!result) {
			return;
		}
		await this.writeCopilotWorkspaceFiles(result);
		const workspaceUri = await this.writeDeveloperWorkspaceMarker(result);
		await vscode.commands.executeCommand(
			'vscode.openFolder',
			workspaceUri,
			{ forceNewWindow: true, noRecentEntry: true }
		);
	}

	async openDeveloperWorkspaceIfPresent(): Promise<void> {
		const folders = vscode.workspace.workspaceFolders || [];
		if (folders.length !== 1 || folders[0].uri.scheme !== 'file') {
			return;
		}
		const workspaceDirectory = folders[0].uri.fsPath;
		const { exportsRoot } = await this.ensureStorage();
		if (!pathIsInside(workspaceDirectory, exportsRoot)) {
			return;
		}
		const markerPath = path.join(workspaceDirectory, DEVELOPER_MARKER_NAME);
		try {
			await assertNoReparsePointChain(markerPath, workspaceDirectory);
			const markerStat = await fs.promises.lstat(markerPath);
			if (
				!markerStat.isFile() ||
				markerStat.isSymbolicLink() ||
				markerStat.size > MAX_DEVELOPER_MARKER_BYTES
			) {
				return;
			}
			const parsed = JSON.parse(
				await fs.promises.readFile(markerPath, 'utf8')
			) as Partial<DeveloperWorkspaceMarker>;
			if (
				parsed.schemaVersion !== 1 ||
				typeof parsed.workbookPath !== 'string' ||
				typeof parsed.workbookSha256 !== 'string' ||
				!/^[0-9a-f]{64}$/.test(parsed.workbookSha256) ||
				typeof parsed.baseName !== 'string' ||
				typeof parsed.outputDirectory !== 'string' ||
				typeof parsed.markdownPath !== 'string' ||
				typeof parsed.jsonPath !== 'string' ||
				typeof parsed.vbaDirectory !== 'string'
			) {
				throw new Error('Marqueur Développeur invalide.');
			}
			if (!sameFile(parsed.vbaDirectory, workspaceDirectory)) {
				throw new Error('Le marqueur cible un autre dossier VBA.');
			}
			await assertOwnedDirectory(parsed.outputDirectory, exportsRoot);
			await assertOwnedDirectory(
				parsed.vbaDirectory,
				parsed.outputDirectory
			);
			for (const generatedPath of [
				parsed.markdownPath,
				parsed.jsonPath
			]) {
				if (!pathIsInside(generatedPath, parsed.outputDirectory)) {
					throw new Error('Le marqueur contient un chemin de sortie inattendu.');
				}
				await assertNoReparsePointChain(
					generatedPath,
					parsed.outputDirectory
				);
				await fs.promises.access(generatedPath, fs.constants.R_OK);
			}
			const workbookUri = await canonicalizeWorkbookUri(
				vscode.Uri.file(parsed.workbookPath)
			);
			assertNotManagedBackupPath(workbookUri.fsPath);
			if (
				(await hashFileSha256(workbookUri.fsPath)) !==
				parsed.workbookSha256
			) {
				throw new Error(
					'Le classeur a changé depuis la création de la fenêtre Développeur. Rouvrez-la depuis le classeur.'
				);
			}
			const context: ExportContext = {
				workbookUri,
				markdownUri: vscode.Uri.file(parsed.markdownPath),
				jsonUri: vscode.Uri.file(parsed.jsonPath),
				paths: {
					workbookPath: workbookUri.fsPath,
					canonicalWorkbookPath: workbookUri.fsPath,
					baseName: parsed.baseName,
					outputDirectory: parsed.outputDirectory,
					markdownPath: parsed.markdownPath,
					markdownUri: vscode.Uri.file(parsed.markdownPath),
					jsonPath: parsed.jsonPath,
					jsonUri: vscode.Uri.file(parsed.jsonPath),
					vbaDirectory: parsed.vbaDirectory,
					vbaDirectoryUri: vscode.Uri.file(parsed.vbaDirectory)
				},
				includeVba: true
			};
			this.lastContext = context;
			this.contextChangeEmitter.fire();
			await this.vbaStudioPanel.prepare(context);
			await this.vbaStudioPanel.open(context);
		} catch (error) {
			this.outputChannel.appendLine(
				`[vba developer] ${(error as Error).message}`
			);
			await vscode.window.showErrorMessage(
				`Mode Développeur non ouvert : ${(error as Error).message}`
			);
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
			await showUserFormPreview(
				uri,
				document.getText(),
				context.workbookUri
			);
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
			'Ne crée jamais de faux UserForm .frm ou .frx. Pour un nouveau UserForm réel, ajouter ou repositionner ses contrôles, affecter ses événements, créer ses boutons ou ActiveX autorisés, utilise #excelVbaDesignWorkbook ; MsgBox et InputBox ne sont pas des UserForms.',
			'Pour un événement complexe, utilise setUserFormEventHandler avec une unique procédure Private Sub complète et replaceExisting=true uniquement si le remplacement est explicitement voulu.',
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
