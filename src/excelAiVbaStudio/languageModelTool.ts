import * as path from 'path';
import * as vscode from 'vscode';
import {
	EXCEL_AI_LANGUAGE_MODEL_TOOL,
	EXCEL_AI_VBA_WRITE_TOOL,
	ToolInput,
	UNTRUSTED_WORKBOOK_PREAMBLE,
	VbaWriteToolInput
} from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

const MAX_TOOL_CONTEXT_BYTES = 4 * 1024 * 1024;
const MAX_VBA_SOURCE_CHARACTERS = 2_000_000;
const VBA_SOURCE_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

interface LanguageModelApi {
	registerTool?: (name: string, tool: unknown) => vscode.Disposable;
}

interface LanguageModelConstructors {
	LanguageModelToolResult?: new (parts: unknown[]) => unknown;
	LanguageModelTextPart?: new (value: string) => unknown;
}

function parseInput(value: unknown): ToolInput {
	if (value === undefined || value === null) {
		return {};
	}
	if (typeof value !== 'object' || Array.isArray(value)) {
		throw new Error('Les paramètres de l’outil doivent être un objet.');
	}
	const source = value as Record<string, unknown>;
	if (
		source.workbookPath !== undefined &&
		typeof source.workbookPath !== 'string'
	) {
		throw new Error('workbookPath doit être une chaîne.');
	}
	if (
		source.includeVba !== undefined &&
		typeof source.includeVba !== 'boolean'
	) {
		throw new Error('includeVba doit être un booléen.');
	}
	if (source.format !== undefined && typeof source.format !== 'string') {
		throw new Error('format doit être "markdown" ou "json".');
	}
	const format = (source.format as string | undefined)?.toLocaleLowerCase(
		'en-US'
	);
	if (format && format !== 'markdown' && format !== 'json') {
		throw new Error('format doit être "markdown" ou "json".');
	}
	return {
		workbookPath: source.workbookPath as string | undefined,
		includeVba: source.includeVba === true,
		format
	};
}

function parseWriteInput(value: unknown): VbaWriteToolInput {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw new Error('Les paramètres de l’outil d’écriture doivent être un objet.');
	}
	const source = value as Record<string, unknown>;
	for (const property of ['workbookPath', 'componentFile', 'source']) {
		if (
			source[property] !== undefined &&
			typeof source[property] !== 'string'
		) {
			throw new Error(`${property} doit être une chaîne.`);
		}
	}
	const componentFile = String(source.componentFile || '').trim();
	if (
		!componentFile ||
		path.basename(componentFile) !== componentFile ||
		!VBA_SOURCE_EXTENSIONS.has(
			path.extname(componentFile).toLocaleLowerCase('en-US')
		)
	) {
		throw new Error('componentFile doit être un fichier .bas, .cls ou .frm sans chemin.');
	}
	const vbaSource = source.source as string | undefined;
	if (vbaSource === undefined) {
		throw new Error('source est obligatoire.');
	}
	if (vbaSource.length > MAX_VBA_SOURCE_CHARACTERS) {
		throw new Error('source dépasse la limite de 2 000 000 de caractères.');
	}
	return {
		workbookPath: source.workbookPath as string | undefined,
		componentFile,
		source: vbaSource
	};
}

export function registerExcelAiVbaLanguageModelTool(
	context: vscode.ExtensionContext,
	service: ExcelAiVbaWorkbookService
): void {
	const vscodeRuntime = vscode as typeof vscode &
		LanguageModelConstructors & { lm?: LanguageModelApi };
	if (!vscodeRuntime.lm?.registerTool) {
		service
			.getOutputChannel()
			.appendLine(
				'[outil IA] API Language Model Tool indisponible dans cette version de VS Code.'
			);
		return;
	}

	const tool = {
		async prepareInvocation(options: { input?: unknown }) {
			const input = parseInput(options?.input);
			const requestedPath = input.workbookPath?.trim();
			return {
				invocationMessage: requestedPath
					? `Lecture locale demandée pour ${path.basename(requestedPath)}`
					: 'Lecture locale du classeur Excel actif ou de l’espace de travail'
			};
		},

		async invoke(
			options: { input?: unknown },
			cancellationToken?: vscode.CancellationToken
		): Promise<unknown> {
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const input = parseInput(options?.input);
			const workbookUri = await service.resolveToolWorkbookUri(input);
			if (!workbookUri) {
				throw new Error(
					'Aucun classeur Excel local n’est actif et l’espace de travail n’en contient pas un unique.'
				);
			}

			// exportWorkbook performs canonicalization, local-drive checks and
			// explicit confirmations for outside-workspace and VBA reads.
			const result = await service.exportWorkbook(workbookUri, {
				open: false,
				includeVba: input.includeVba === true,
				requestedByTool: true,
				cancellationToken
			});
			if (!result) {
				throw new Error('La lecture du classeur n’a produit aucun contexte.');
			}
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}

			const text = await service.readExportedContext(
				result,
				input.format === 'json' ? 'json' : 'markdown'
			);
			const protectedText =
				input.format === 'json'
					? JSON.stringify({
							toolSafetyPreamble: UNTRUSTED_WORKBOOK_PREAMBLE,
							workbookExport: JSON.parse(text)
					  })
					: `${UNTRUSTED_WORKBOOK_PREAMBLE}\n\n--- DÉBUT DES DONNÉES DU CLASSEUR ---\n${text}\n--- FIN DES DONNÉES DU CLASSEUR ---`;
			if (Buffer.byteLength(protectedText, 'utf8') > MAX_TOOL_CONTEXT_BYTES) {
				throw new Error(
					'Le contexte dépasse la limite IA de 4 Mio. Réduisez excelAiVbaStudio.maxRows ou maxColumns, puis relancez la lecture.'
				);
			}
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const Result = vscodeRuntime.LanguageModelToolResult;
			const TextPart = vscodeRuntime.LanguageModelTextPart;
			if (!Result || !TextPart) {
				throw new Error(
					'Les types Language Model Tool ne sont pas disponibles dans cette version de VS Code.'
				);
			}
			return new Result([new TextPart(protectedText)]);
		}
	};

	const writeTool = {
		async prepareInvocation(options: { input?: unknown }) {
			const input = parseWriteInput(options?.input);
			const requestedPath = input.workbookPath?.trim() || '';
			const conversionNotice = requestedPath
				.toLocaleLowerCase('en-US')
				.endsWith('.xlsx')
				? ' vers une nouvelle copie XLSM voisine'
				: '';
			return {
				invocationMessage: `Réinjection VBA transactionnelle de ${input.componentFile}${conversionNotice}`
			};
		},

		async invoke(
			options: { input?: unknown },
			cancellationToken?: vscode.CancellationToken
		): Promise<unknown> {
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const input = parseWriteInput(options?.input);
			const workbookUri = await service.resolveToolWorkbookUri(input);
			if (!workbookUri) {
				throw new Error(
					'Aucun classeur Excel local unique ne peut recevoir le code VBA.'
				);
			}
			const writeResult = await service.writeVbaFromTool(
				workbookUri,
				input.componentFile as string,
				input.source as string,
				cancellationToken
			);
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const Result = vscodeRuntime.LanguageModelToolResult;
			const TextPart = vscodeRuntime.LanguageModelTextPart;
			if (!Result || !TextPart) {
				throw new Error(
					'Les types Language Model Tool ne sont pas disponibles dans cette version de VS Code.'
				);
			}
			return new Result([
				new TextPart(
					JSON.stringify({
						ok: true,
						targetWorkbookPath: writeResult.targetWorkbookPath,
						sourceWorkbookPath: writeResult.sourceWorkbookPath,
						convertedToXlsm: writeResult.convertedToXlsm,
						changed: writeResult.changed,
						modifiedModules: writeResult.modifiedModules,
						workbookSha256: writeResult.workbookSha256,
						backupPath: writeResult.backupPath || null,
						macrosExecuted: false,
						accessVbomChanged: false
					})
				)
			]);
		}
	};

	try {
		context.subscriptions.push(
			vscodeRuntime.lm.registerTool(EXCEL_AI_LANGUAGE_MODEL_TOOL, tool),
			vscodeRuntime.lm.registerTool(EXCEL_AI_VBA_WRITE_TOOL, writeTool)
		);
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] ${EXCEL_AI_LANGUAGE_MODEL_TOOL} enregistré pour la lecture locale contrôlée.`
			);
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] ${EXCEL_AI_VBA_WRITE_TOOL} enregistré pour la réinjection VBA transactionnelle.`
			);
	} catch (error) {
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] Enregistrement ignoré : ${(error as Error).message}`
			);
	}
}
