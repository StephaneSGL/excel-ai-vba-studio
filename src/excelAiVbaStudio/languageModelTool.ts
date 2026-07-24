import * as path from 'path';
import * as vscode from 'vscode';
import {
	EXCEL_AI_LANGUAGE_MODEL_TOOL,
	ToolInput,
	UNTRUSTED_WORKBOOK_PREAMBLE
} from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

const MAX_TOOL_CONTEXT_BYTES = 4 * 1024 * 1024;

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

	try {
		context.subscriptions.push(
			vscodeRuntime.lm.registerTool(EXCEL_AI_LANGUAGE_MODEL_TOOL, tool)
		);
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] ${EXCEL_AI_LANGUAGE_MODEL_TOOL} enregistré pour la lecture locale contrôlée.`
			);
	} catch (error) {
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] Enregistrement ignoré : ${(error as Error).message}`
			);
	}
}
