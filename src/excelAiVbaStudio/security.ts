import { execFile } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { EXCEL_EXTENSIONS } from './types';

const ALLOW_ONCE = 'Autoriser une fois';
export const EXPORT_OWNER_MARKER = '.excel-ai-vba-studio-owned';
const EXPORT_OWNER_MARKER_CONTENT =
	'excel-ai-vba-studio:managed-export-directory:v1\n';

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

function isUncOrDeviceNetworkPath(value: string): boolean {
	const normalized = value.replace(/\//g, '\\');
	return (
		normalized.startsWith('\\\\') ||
		normalized.toLocaleLowerCase('en-US').startsWith('\\\\?\\unc\\')
	);
}

async function readWindowsDriveType(candidatePath: string): Promise<number> {
	if (process.platform !== 'win32') {
		throw new Error(
			'L’intégration Microsoft Excel locale est disponible uniquement sous Windows.'
		);
	}

	const root = path.win32.parse(candidatePath).root;
	if (!root || isUncOrDeviceNetworkPath(root)) {
		throw new Error('Le chemin doit désigner un disque local Windows.');
	}

	const script = [
		'$ErrorActionPreference = "Stop"',
		'$candidate = $env:EXCEL_AI_VBA_DRIVE_PATH',
		'$drive = New-Object IO.DriveInfo([IO.Path]::GetPathRoot($candidate))',
		'[Console]::Out.Write([int]$drive.DriveType)'
	].join('; ');

	return await new Promise<number>((resolve, reject) => {
		execFile(
			getPowerShellPath(),
			[
				'-NoLogo',
				'-NoProfile',
				'-NonInteractive',
				'-Command',
				script
			],
			{
				windowsHide: true,
				timeout: 10_000,
				env: {
					...process.env,
					EXCEL_AI_VBA_DRIVE_PATH: candidatePath
				}
			},
			(error, stdout, stderr) => {
				if (error) {
					reject(
						new Error(
							`Impossible de vérifier le lecteur local : ${
								String(stderr || '').trim() || error.message
							}`
						)
					);
					return;
				}
				const driveType = Number.parseInt(String(stdout).trim(), 10);
				if (!Number.isFinite(driveType)) {
					reject(new Error('Le type du lecteur local est indéterminé.'));
					return;
				}
				resolve(driveType);
			}
		);
	});
}

async function assertLocalDrive(candidatePath: string): Promise<void> {
	if (isUncOrDeviceNetworkPath(candidatePath)) {
		throw new Error(
			'Les chemins UNC et les fichiers réseau sont refusés. Copiez le classeur sur un disque local.'
		);
	}

	const driveType = await readWindowsDriveType(candidatePath);
	// System.IO.DriveType.Network === 4. Unknown/NoRootDirectory are refused too:
	// the extension must positively establish that the source is local.
	if (driveType === 4) {
		throw new Error(
			'Les lecteurs réseau mappés sont refusés. Copiez le classeur sur un disque local.'
		);
	}
	if (driveType === 0 || driveType === 1) {
		throw new Error('Le lecteur du classeur n’a pas pu être vérifié comme local.');
	}
}

export async function assertLocalPath(candidatePath: string): Promise<void> {
	await assertLocalDrive(candidatePath);
}

/**
 * Refuse every existing Windows reparse point in a path chain. Checking only
 * realpath() is insufficient for output paths: a junction can be swapped in
 * between the lexical containment check and a later recursive delete.
 */
export async function assertNoReparsePointChain(
	candidatePath: string,
	expectedRoot?: string
): Promise<string> {
	const fullPath = path.resolve(candidatePath);
	await assertLocalDrive(fullPath);
	if (expectedRoot) {
		const fullRoot = path.resolve(expectedRoot);
		await assertLocalDrive(fullRoot);
		if (!isPathInside(fullPath, fullRoot)) {
			throw new Error(
				`Le chemin géré sort du stockage autorisé : ${fullPath}`
			);
		}
	}

	const script = [
		'$ErrorActionPreference = "Stop"',
		'$full = [IO.Path]::GetFullPath($env:EXCEL_AI_VBA_REPARSE_PATH)',
		'$root = [IO.Path]::GetPathRoot($full)',
		'$current = $root',
		'$relative = $full.Substring($root.Length)',
		'foreach ($part in $relative.Split([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar), [StringSplitOptions]::RemoveEmptyEntries)) {',
		'  $current = [IO.Path]::Combine($current, $part)',
		'  if (-not (Test-Path -LiteralPath $current)) { break }',
		'  $item = Get-Item -LiteralPath $current -Force',
		'  if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {',
		'    [Console]::Out.Write([IO.Path]::GetFullPath($item.FullName))',
		'    exit 23',
		'  }',
		'}'
	].join('\n');

	await new Promise<void>((resolve, reject) => {
		execFile(
			getPowerShellPath(),
			['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script],
			{
				windowsHide: true,
				timeout: 10_000,
				env: {
					...process.env,
					EXCEL_AI_VBA_REPARSE_PATH: fullPath
				}
			},
			(error, stdout, stderr) => {
				if (error) {
					const reparsePath = String(stdout || '').trim();
					reject(
						new Error(
							reparsePath
								? `Un lien symbolique, une jonction ou un point de réanalyse est refusé dans le stockage d’export : ${reparsePath}`
								: `Impossible de vérifier la chaîne du chemin d’export : ${
										String(stderr || '').trim() || error.message
								  }`
						)
					);
					return;
				}
				resolve();
			}
		);
	});
	return fullPath;
}

export async function assertNoReparsePointsInTree(
	directoryPath: string
): Promise<void> {
	const safeDirectory = await assertNoReparsePointChain(directoryPath);
	const script = [
		'$ErrorActionPreference = "Stop"',
		'$root = [IO.Path]::GetFullPath($env:EXCEL_AI_VBA_REPARSE_TREE)',
		'$pending = New-Object "System.Collections.Generic.Stack[string]"',
		'$pending.Push($root)',
		'while ($pending.Count -gt 0) {',
		'  $current = $pending.Pop()',
		'  $item = Get-Item -LiteralPath $current -Force',
		'  if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {',
		'    [Console]::Out.Write([IO.Path]::GetFullPath($item.FullName))',
		'    exit 23',
		'  }',
		'  if ($item.PSIsContainer) {',
		'    foreach ($child in Get-ChildItem -LiteralPath $item.FullName -Force) {',
		'      if (($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {',
		'        [Console]::Out.Write([IO.Path]::GetFullPath($child.FullName))',
		'        exit 23',
		'      }',
		'      if ($child.PSIsContainer) { $pending.Push([string]$child.FullName) }',
		'    }',
		'  }',
		'}'
	].join('\n');

	await new Promise<void>((resolve, reject) => {
		execFile(
			getPowerShellPath(),
			['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script],
			{
				windowsHide: true,
				timeout: 30_000,
				env: {
					...process.env,
					EXCEL_AI_VBA_REPARSE_TREE: safeDirectory
				}
			},
			(error, stdout, stderr) => {
				if (error) {
					const reparsePath = String(stdout || '').trim();
					reject(
						new Error(
							reparsePath
								? `Suppression refusée : un point de réanalyse existe dans le dossier géré : ${reparsePath}`
								: `Impossible d’auditer le dossier avant suppression : ${
										String(stderr || '').trim() || error.message
								  }`
						)
					);
					return;
				}
				resolve();
			}
		);
	});
}

export async function ensureLocalDirectory(directoryPath: string): Promise<string> {
	await assertLocalDrive(directoryPath);
	await assertNoReparsePointChain(directoryPath);
	await fs.promises.mkdir(directoryPath, { recursive: true });
	await assertNoReparsePointChain(directoryPath);
	const canonicalPath = await fs.promises.realpath(directoryPath);
	await assertLocalDrive(canonicalPath);
	return canonicalPath;
}

async function readOwnerMarker(directoryPath: string): Promise<string> {
	await assertNoReparsePointChain(directoryPath);
	const markerPath = path.join(directoryPath, EXPORT_OWNER_MARKER);
	await assertNoReparsePointChain(markerPath, directoryPath);
	const markerStat = await fs.promises.lstat(markerPath);
	if (!markerStat.isFile() || markerStat.isSymbolicLink()) {
		throw new Error(
			`Le marqueur de propriété du dossier d’export n’est pas un fichier sûr : ${markerPath}`
		);
	}
	return await fs.promises.readFile(markerPath, 'utf8');
}

export async function assertOwnedDirectory(
	directoryPath: string,
	expectedRoot?: string
): Promise<string> {
	const safePath = await assertNoReparsePointChain(directoryPath, expectedRoot);
	const stat = await fs.promises.lstat(safePath);
	if (!stat.isDirectory() || stat.isSymbolicLink()) {
		throw new Error(`Le dossier d’export géré n’est pas sûr : ${safePath}`);
	}
	if ((await readOwnerMarker(safePath)) !== EXPORT_OWNER_MARKER_CONTENT) {
		throw new Error(
			`Le dossier ne porte pas le marqueur de propriété attendu : ${safePath}`
		);
	}
	await assertNoReparsePointChain(safePath, expectedRoot);
	return safePath;
}

export async function ensureOwnedDirectory(
	directoryPath: string,
	expectedRoot?: string
): Promise<string> {
	const fullPath = await assertNoReparsePointChain(directoryPath, expectedRoot);
	const existed = await fs.promises
		.lstat(fullPath)
		.then(() => true)
		.catch(error => {
			if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
				return false;
			}
			throw error;
		});

	if (!existed) {
		await fs.promises.mkdir(fullPath, { recursive: true });
	}
	await assertNoReparsePointChain(fullPath, expectedRoot);

	const markerPath = path.join(fullPath, EXPORT_OWNER_MARKER);
	const markerExists = await fs.promises
		.lstat(markerPath)
		.then(() => true)
		.catch(error => {
			if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
				return false;
			}
			throw error;
		});
	if (!markerExists) {
		const entries = await fs.promises.readdir(fullPath);
		if (entries.length !== 0) {
			throw new Error(
				`Refus de prendre possession d’un dossier d’export non vide : ${fullPath}`
			);
		}
		await assertNoReparsePointChain(markerPath, fullPath);
		await fs.promises.writeFile(markerPath, EXPORT_OWNER_MARKER_CONTENT, {
			encoding: 'utf8',
			flag: 'wx'
		});
	}
	return await assertOwnedDirectory(fullPath, expectedRoot);
}

export async function removeOwnedDirectory(
	directoryPath: string,
	expectedRoot: string
): Promise<void> {
	const safePath = await assertOwnedDirectory(directoryPath, expectedRoot);
	if (path.resolve(safePath) === path.resolve(expectedRoot)) {
		throw new Error('Le dossier racine ne peut pas être supprimé par cette opération.');
	}
	// Revalidate immediately before the destructive operation.
	await assertOwnedDirectory(safePath, expectedRoot);
	await assertNoReparsePointsInTree(safePath);
	await fs.promises.rm(safePath, { recursive: true, force: false });
}

function isExcelPath(candidatePath: string): boolean {
	return EXCEL_EXTENSIONS.has(path.extname(candidatePath).toLocaleLowerCase('en-US'));
}

function isPathInside(candidatePath: string, rootPath: string): boolean {
	const relative = path.relative(rootPath, candidatePath);
	return (
		relative === '' ||
		(!relative.startsWith(`..${path.sep}`) &&
			relative !== '..' &&
			!path.isAbsolute(relative))
	);
}

async function canonicalWorkspaceRoots(): Promise<string[]> {
	const roots: string[] = [];
	for (const folder of vscode.workspace.workspaceFolders || []) {
		if (folder.uri.scheme !== 'file' || folder.uri.authority) {
			continue;
		}
		try {
			await assertLocalDrive(folder.uri.fsPath);
			roots.push(await fs.promises.realpath(folder.uri.fsPath));
		} catch {
			// A remote, unavailable or unverifiable workspace must not silently
			// authorize a workbook path.
		}
	}
	return roots;
}

export async function canonicalizeWorkbookUri(uri: vscode.Uri): Promise<vscode.Uri> {
	if (!uri || uri.scheme !== 'file' || uri.authority) {
		throw new Error(
			'Le classeur doit être un fichier local, sans autorité réseau ni URI distante.'
		);
	}
	if (!isExcelPath(uri.fsPath)) {
		throw new Error('Le fichier doit être un classeur .xlsx, .xlsm, .xls ou .xlsb.');
	}

	// Check the drive before realpath/stat so a mapped network path is never read.
	await assertLocalDrive(uri.fsPath);
	const canonicalPath = await fs.promises.realpath(uri.fsPath);
	await assertLocalDrive(canonicalPath);
	const stat = await fs.promises.stat(canonicalPath);
	if (!stat.isFile()) {
		throw new Error('La ressource demandée n’est pas un fichier.');
	}
	return vscode.Uri.file(canonicalPath);
}

export async function isWorkbookInsideWorkspace(
	canonicalWorkbookPath: string
): Promise<boolean> {
	const roots = await canonicalWorkspaceRoots();
	return roots.some(root => isPathInside(canonicalWorkbookPath, root));
}

async function confirmExactPath(
	message: string,
	canonicalWorkbookPath: string
): Promise<boolean> {
	const choice = await vscode.window.showWarningMessage(
		`${message}\n\n${canonicalWorkbookPath}`,
		{
			modal: true,
			detail: `Chemin canonique exact : ${canonicalWorkbookPath}`
		},
		ALLOW_ONCE
	);
	return choice === ALLOW_ONCE;
}

export async function authorizeWorkbookRead(
	candidate: vscode.Uri,
	options: { includeVba: boolean }
): Promise<vscode.Uri | undefined> {
	const canonicalUri = await canonicalizeWorkbookUri(candidate);
	const isInsideWorkspace = await isWorkbookInsideWorkspace(canonicalUri.fsPath);

	if (
		!isInsideWorkspace &&
		!(await confirmExactPath(
			'Autoriser la lecture ponctuelle de ce classeur situé hors de l’espace de travail ?',
			canonicalUri.fsPath
		))
	) {
		return undefined;
	}

	if (
		options.includeVba &&
		!(await confirmExactPath(
			'Autoriser l’extraction ponctuelle du code et des références VBA de ce classeur ?',
			canonicalUri.fsPath
		))
	) {
		return undefined;
	}

	return canonicalUri;
}

export function workbookUriFromPathInput(requestedPath: string): vscode.Uri {
	const trimmed = requestedPath.trim();
	if (!trimmed) {
		throw new Error('workbookPath est vide.');
	}

	if (/^file:/i.test(trimmed)) {
		const parsed = vscode.Uri.parse(trimmed, true);
		if (parsed.scheme !== 'file' || parsed.authority) {
			throw new Error('workbookPath doit être un URI file local sans autorité réseau.');
		}
		return parsed;
	}

	let localPath = trimmed;
	if (!path.isAbsolute(localPath)) {
		const localWorkspace = (vscode.workspace.workspaceFolders || []).find(
			folder => folder.uri.scheme === 'file' && !folder.uri.authority
		);
		if (!localWorkspace) {
			throw new Error(
				'Un chemin relatif exige qu’un dossier local soit ouvert dans VS Code.'
			);
		}
		localPath = path.resolve(localWorkspace.uri.fsPath, localPath);
	}
	return vscode.Uri.file(localPath);
}

export function pathIsInside(candidatePath: string, rootPath: string): boolean {
	return isPathInside(path.resolve(candidatePath), path.resolve(rootPath));
}
