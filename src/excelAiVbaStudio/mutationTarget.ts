import * as path from 'path';
import * as vscode from 'vscode';
import { assertNoReparsePointChain, canonicalizeWorkbookUri, pathIsInside, workbookUriFromPathInput } from './security';

/** Mutations must never fall back to the active editor or the last export. */
export function requireExplicitWorkbookPath(value: unknown): string {
	if (typeof value !== 'string' || !value.trim()) {
		throw new Error('workbookPath est obligatoire pour une écriture.');
	}
	const requestedPath = value.trim();
	if (requestedPath.length > 32_767 || /[\u0000-\u001f]/.test(requestedPath)) {
		throw new Error('workbookPath est invalide ou trop long.');
	}
	const isFileUri = /^file:/i.test(requestedPath);
	let localPath = requestedPath;
	if (isFileUri) {
		const uri = vscode.Uri.parse(requestedPath, true);
		if (uri.scheme !== 'file' || uri.authority || uri.query || uri.fragment) {
			throw new Error('workbookPath doit être un URI file local sans autorité, requête ni fragment.');
		}
		localPath = uri.fsPath;
	}
	if (!path.isAbsolute(localPath) || (process.platform === 'win32' && !/^[A-Za-z]:[\\/]/.test(localPath))) {
		throw new Error('workbookPath doit être un chemin absolu explicite pour cette opération avec effet de bord.');
	}
	return requestedPath;
}

export async function resolveMutationTarget(workbookPath: string): Promise<vscode.Uri> {
	const uri = workbookUriFromPathInput(requireExplicitWorkbookPath(workbookPath));
	await assertNoReparsePointChain(uri.fsPath);
	return canonicalizeWorkbookUri(uri);
}

export function mutationConfirmation(uri: vscode.Uri, summary: string) {
	const insideWorkspace = (vscode.workspace.workspaceFolders ?? []).some(folder => (
		folder.uri.scheme === 'file' && !folder.uri.authority && pathIsInside(uri.fsPath, folder.uri.fsPath)
	));
	return {
		title: `Modifier ${path.basename(uri.fsPath)} ?`,
		message: [
			summary,
			`Chemin canonique complet : ${uri.fsPath}`,
			insideWorkspace ? 'La cible se trouve dans l’espace de travail ouvert.' : 'Attention : la cible se trouve hors de l’espace de travail ouvert.',
			'Écriture transactionnelle avec vérification. Aucune macro ne sera exécutée ; AccessVBOM ne sera pas modifié.'
		].join('\n\n')
	};
}
