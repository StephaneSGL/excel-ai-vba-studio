import { createHash } from 'crypto';
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { assertOoxmlPackageUnsignedForMutation } from '../common/ooxmlPackageSignature';
import {
	assertNoReparsePointChain,
	assertOwnedDirectory,
	ensureOwnedDirectory,
	pathIsInside
} from './security';
import { ExportContext } from './types';

const MAX_SOURCE_CHARACTERS = 2_000_000;
const MAX_REQUEST_BYTES = 5 * 1024 * 1024;
const WRITEBACK_TIMEOUT_MS = 45_000;
const SUPPORTED_EXTENSIONS = new Set(['.xlsm', '.xlam']);
const SOURCE_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

export type VbaWritebackComponentKind =
	| 'module'
	| 'class'
	| 'document'
	| 'userform';

interface ManifestModule {
	file?: unknown;
	name?: unknown;
	type?: unknown;
}

interface VbaManifest {
	modules?: ManifestModule[];
}

interface ComponentBaseline {
	file: string;
	moduleName: string;
	componentKind: VbaWritebackComponentKind;
	embedded: boolean;
	sourceSha256: string;
	designerSha256?: string;
}

interface WorkbookBaseline {
	workbookSha256: string;
	components: Map<string, ComponentBaseline>;
}

interface HelperResult {
	ok: boolean;
	changed?: boolean;
	code?: string;
	message?: string;
	modifiedModules?: string[];
	workbookSha256?: string;
	backupPath?: string | null;
}

export interface VbaWritebackResult {
	changed: boolean;
	modifiedModules: string[];
	workbookSha256: string;
	backupPath?: string;
}

function normalizeNewlines(value: string): string {
	return value.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, '\r\n');
}

function hashText(value: string): string {
	return createHash('sha256').update(value, 'utf8').digest('hex');
}

async function hashFile(filePath: string): Promise<string> {
	const digest = createHash('sha256');
	const stream = fs.createReadStream(filePath);
	for await (const chunk of stream) {
		digest.update(chunk as Buffer);
	}
	return digest.digest('hex');
}

function normalizeComponentKind(
	typeValue: unknown,
	file: string
): VbaWritebackComponentKind {
	const type = typeof typeValue === 'string' ? typeValue.toLocaleLowerCase('en-US') : '';
	if (type.includes('document')) {
		return 'document';
	}
	if (type.includes('userform')) {
		return 'userform';
	}
	if (type.includes('class')) {
		return 'class';
	}
	if (type.includes('standard')) {
		return 'module';
	}
	switch (path.extname(file).toLocaleLowerCase('en-US')) {
		case '.bas':
			return 'module';
		case '.frm':
			return 'userform';
		default:
			return 'class';
	}
}

function userFormDesignerHash(source: string): string {
	const normalized = normalizeNewlines(source);
	const match = /^Attribute[ \t]+VB_Name[ \t]*=/im.exec(normalized);
	if (!match) {
		throw new Error(
			'Le fichier UserForm ne contient pas Attribute VB_Name et ne peut pas être réinjecté.'
		);
	}
	return hashText(normalized.slice(0, match.index));
}

function parseHelperResult(stdout: string): HelperResult {
	const line = stdout
		.replace(/\r/g, '')
		.split('\n')
		.map(value => value.trim())
		.filter(Boolean)
		.pop();
	if (!line) {
		throw new Error('Le moteur VBA n’a renvoyé aucun résultat.');
	}
	const parsed = JSON.parse(line) as HelperResult;
	if (!parsed || typeof parsed.ok !== 'boolean') {
		throw new Error('Le moteur VBA a renvoyé un résultat invalide.');
	}
	return parsed;
}

async function terminateOwnedProcessTree(processId: number): Promise<void> {
	if (!Number.isSafeInteger(processId) || processId <= 0) {
		return;
	}
	await new Promise<void>(resolve => {
		const taskkillPath = process.env.SystemRoot
			? path.join(process.env.SystemRoot, 'System32', 'taskkill.exe')
			: 'taskkill.exe';
		const killer = spawn(
			taskkillPath,
			['/PID', String(processId), '/T', '/F'],
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

export class VbaWritebackService implements vscode.Disposable {
	private readonly baselines = new Map<string, WorkbookBaseline>();
	private readonly queues = new Map<string, Promise<unknown>>();
	private disposed = false;

	constructor(
		private readonly extensionContext: vscode.ExtensionContext,
		private readonly outputChannel: vscode.OutputChannel
	) {}

	dispose(): void {
		this.disposed = true;
		this.baselines.clear();
		this.queues.clear();
	}

	async prepare(context: ExportContext): Promise<void> {
		const key = this.workbookKey(context.workbookUri.fsPath);
		await this.enqueue(key, () => this.prepareUnlocked(context));
	}

	private async prepareUnlocked(context: ExportContext): Promise<void> {
		const workbookPath = context.workbookUri.fsPath;
		if (!SUPPORTED_EXTENSIONS.has(path.extname(workbookPath).toLocaleLowerCase('en-US'))) {
			this.baselines.delete(this.workbookKey(workbookPath));
			return;
		}
		await assertNoReparsePointChain(workbookPath);
		await assertOwnedDirectory(
			context.paths.vbaDirectory,
			context.paths.outputDirectory
		);
		const manifestPath = path.join(context.paths.vbaDirectory, 'manifest.json');
		let manifest: VbaManifest = {};
		try {
			await assertNoReparsePointChain(manifestPath, context.paths.vbaDirectory);
			manifest = JSON.parse(
				(await fs.promises.readFile(manifestPath, 'utf8')).replace(/^\uFEFF/, '')
			) as VbaManifest;
		} catch {
			manifest = {};
		}
		const manifestByFile = new Map(
			(manifest.modules || [])
				.filter(module => typeof module.file === 'string')
				.map(module => [
					String(module.file).toLocaleLowerCase('en-US'),
					module
				])
		);
		const components = new Map<string, ComponentBaseline>();
		for (const entry of await fs.promises.readdir(context.paths.vbaDirectory, {
			withFileTypes: true
		})) {
			if (
				!entry.isFile() ||
				!SOURCE_EXTENSIONS.has(
					path.extname(entry.name).toLocaleLowerCase('en-US')
				)
			) {
				continue;
			}
			const sourcePath = path.join(context.paths.vbaDirectory, entry.name);
			await assertNoReparsePointChain(sourcePath, context.paths.vbaDirectory);
			const source = (await fs.promises.readFile(sourcePath, 'utf8')).replace(
				/^\uFEFF/,
				''
			);
			const manifestRecord = manifestByFile.get(
				entry.name.toLocaleLowerCase('en-US')
			);
			const moduleName =
				typeof manifestRecord?.name === 'string' && manifestRecord.name
					? manifestRecord.name
					: path.basename(entry.name, path.extname(entry.name));
			const componentKind = normalizeComponentKind(
				manifestRecord?.type,
				entry.name
			);
			components.set(entry.name.toLocaleLowerCase('en-US'), {
				file: entry.name,
				moduleName,
				componentKind,
				embedded: Boolean(manifestRecord),
				sourceSha256: hashText(source),
				designerSha256:
					componentKind === 'userform'
						? userFormDesignerHash(source)
						: undefined
			});
		}
		this.baselines.set(this.workbookKey(workbookPath), {
			workbookSha256: await hashFile(workbookPath),
			components
		});
	}

	async applySource(
		context: ExportContext,
		file: string,
		source: string
	): Promise<VbaWritebackResult> {
		if (this.disposed) {
			throw new Error('Le service de réinjection VBA est arrêté.');
		}
		if (source.length > MAX_SOURCE_CHARACTERS) {
			throw new Error(
				`Le module dépasse la limite de ${MAX_SOURCE_CHARACTERS.toLocaleString(
					'fr-FR'
				)} caractères.`
			);
		}
		const workbookPath = context.workbookUri.fsPath;
		if (!SUPPORTED_EXTENSIONS.has(path.extname(workbookPath).toLocaleLowerCase('en-US'))) {
			throw new Error(
				'La réinjection automatique VBA est disponible pour les fichiers .xlsm et .xlam.'
			);
		}
		const sourcePath = path.join(context.paths.vbaDirectory, path.basename(file));
		if (
			path.basename(file) !== file ||
			!pathIsInside(sourcePath, context.paths.vbaDirectory)
		) {
			throw new Error('Le composant demandé sort du projet VBA.');
		}
		await assertNoReparsePointChain(sourcePath, context.paths.vbaDirectory);
		const key = this.workbookKey(workbookPath);
		if (!this.baselines.has(key)) {
			await this.prepare(context);
		}

		return await this.enqueue(key, async () => {
			const latestBaseline = this.baselines.get(key);
			if (!latestBaseline) {
				throw new Error('Le contexte VBA doit être rechargé.');
			}
			const component = latestBaseline.components.get(
				file.toLocaleLowerCase('en-US')
			);
			if (!component) {
				throw new Error(
					'Le composant n’appartient pas au dernier export VBA. Rechargez le studio.'
				);
			}
			if (component.componentKind === 'userform') {
				const currentDesignerHash = userFormDesignerHash(source);
				if (currentDesignerHash !== component.designerSha256) {
					throw new Error(
						'Le designer du UserForm a changé. La v1 réinjecte uniquement son code ; le dessin et le .frx restent protégés.'
					);
				}
			}
			if (component.embedded && hashText(source) === component.sourceSha256) {
				return {
					changed: false,
					modifiedModules: [],
					workbookSha256: latestBaseline.workbookSha256
				};
			}
			await assertOoxmlPackageUnsignedForMutation(workbookPath);
			const result = await this.invokeHelper(context, {
				schemaVersion: 1,
				workbookPath,
				expectedWorkbookSha256: latestBaseline.workbookSha256,
				patches: [
					{
						moduleName: component.moduleName,
						componentKind: component.componentKind,
						source,
						...(component.designerSha256
							? { expectedDesignerSha256: component.designerSha256 }
							: {})
					}
				]
			});
			if (!result.ok) {
				throw new Error(result.message || result.code || 'Réinjection VBA refusée.');
			}
			if (!result.workbookSha256) {
				throw new Error('Le moteur VBA n’a pas confirmé le nouveau hash du classeur.');
			}
			latestBaseline.workbookSha256 = result.workbookSha256;
			component.embedded = true;
			component.sourceSha256 = hashText(source);
			if (component.componentKind === 'userform') {
				component.designerSha256 = userFormDesignerHash(source);
			}
			this.outputChannel.appendLine(
				`[vba write-back] ${component.moduleName}: ${
					result.changed ? 'réinjecté' : 'déjà synchronisé'
				}${result.backupPath ? ` | sauvegarde=${result.backupPath}` : ''}`
			);
			return {
				changed: result.changed === true,
				modifiedModules: result.modifiedModules || [],
				workbookSha256: result.workbookSha256,
				backupPath: result.backupPath || undefined
			};
		});
	}

	private workbookKey(workbookPath: string): string {
		return path.resolve(workbookPath).toLocaleLowerCase('en-US');
	}

	private async enqueue<T>(key: string, operation: () => Promise<T>): Promise<T> {
		const previous = this.queues.get(key) || Promise.resolve();
		const current = previous.catch(() => undefined).then(operation);
		this.queues.set(key, current);
		try {
			return await current;
		} finally {
			if (this.queues.get(key) === current) {
				this.queues.delete(key);
			}
		}
	}

	private async invokeHelper(
		context: ExportContext,
		request: Record<string, unknown>
	): Promise<HelperResult> {
		const executablePath = this.extensionContext.asAbsolutePath(
			path.join('bin', 'win32-x64', 'excel-ai-vba-writeback.exe')
		);
		await fs.promises.access(executablePath, fs.constants.R_OK);
		await assertNoReparsePointChain(executablePath);
		const requestText = JSON.stringify(request);
		if (Buffer.byteLength(requestText, 'utf8') > MAX_REQUEST_BYTES) {
			throw new Error('La demande de réinjection dépasse la limite de 5 Mio.');
		}
		const requestDirectory = path.join(
			context.paths.outputDirectory,
			'.writeback-requests'
		);
		await ensureOwnedDirectory(
			requestDirectory,
			context.paths.outputDirectory
		);
		const requestPath = path.join(
			requestDirectory,
			`request-${Date.now()}-${process.pid}-${Math.random()
				.toString(16)
				.slice(2)}.json`
		);
		await assertNoReparsePointChain(requestPath, requestDirectory);
		await fs.promises.writeFile(requestPath, requestText, {
			encoding: 'utf8',
			flag: 'wx'
		});
		try {
			return await new Promise<HelperResult>((resolve, reject) => {
				const child = spawn(executablePath, [requestPath], {
					cwd: path.dirname(context.workbookUri.fsPath),
					windowsHide: true,
					shell: false,
					stdio: ['ignore', 'pipe', 'pipe']
				});
				let stdout = '';
				let stderr = '';
				let settled = false;
				const timeout = setTimeout(() => {
					if (settled) {
						return;
					}
					settled = true;
					const timeoutError = new Error(
						'La réinjection VBA a dépassé 45 secondes et a été arrêtée.'
					);
					if (typeof child.pid === 'number') {
						void terminateOwnedProcessTree(child.pid).finally(() =>
							reject(timeoutError)
						);
					} else {
						child.kill();
						reject(timeoutError);
					}
				}, WRITEBACK_TIMEOUT_MS);
				child.stdout.setEncoding('utf8');
				child.stderr.setEncoding('utf8');
				child.stdout.on('data', chunk => {
					stdout = (stdout + String(chunk)).slice(-64_000);
				});
				child.stderr.on('data', chunk => {
					stderr = (stderr + String(chunk)).slice(-64_000);
				});
				child.once('error', error => {
					if (!settled) {
						settled = true;
						clearTimeout(timeout);
						reject(error);
					}
				});
				child.once('close', () => {
					if (settled) {
						return;
					}
					settled = true;
					clearTimeout(timeout);
					try {
						resolve(parseHelperResult(stdout));
					} catch (error) {
						reject(
							new Error(
								`${(error as Error).message}${
									stderr.trim() ? ` ${stderr.trim()}` : ''
								}`
							)
						);
					}
				});
			});
		} finally {
			await fs.promises.rm(requestPath, { force: true }).catch(() => undefined);
		}
	}
}
