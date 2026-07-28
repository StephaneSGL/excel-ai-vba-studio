import * as path from 'path';
import * as vscode from 'vscode';
import {
	EXCEL_AI_LANGUAGE_MODEL_TOOL,
	EXCEL_AI_VBA_DESIGN_TOOL,
	EXCEL_AI_VBA_WRITE_TOOL,
	ToolInput,
	UNTRUSTED_WORKBOOK_PREAMBLE,
	VbaDesignOperation,
	VbaDesignToolInput,
	VbaUserFormControl,
	VbaUserFormControlType,
	VbaWriteToolInput
} from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

const MAX_TOOL_CONTEXT_BYTES = 4 * 1024 * 1024;
const MAX_VBA_SOURCE_CHARACTERS = 2_000_000;
const MAX_VBA_DESIGN_OPERATIONS = 100;
const MAX_DESIGN_TEXT_CHARACTERS = 1_000;
const MAX_DESIGN_COORDINATE = 10_000;
const MAX_DESIGN_TAB_INDEX = 32_767;
const VBA_SOURCE_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);
const VBA_IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_]{0,30}$/;
const VBA_MACRO_PATTERN =
	/^[A-Za-z_][A-Za-z0-9_]{0,30}(?:\.[A-Za-z_][A-Za-z0-9_]{0,30})?$/;
const ACTIVEX_PROGID_PATTERN = /^[A-Za-z][A-Za-z0-9_.-]{1,127}$/;
const VBA_USERFORM_CONTROL_TYPES = new Set<VbaUserFormControlType>([
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
	'scrollBar',
	'customActiveX'
]);

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

function designObject(
	value: unknown,
	label: string
): Record<string, unknown> {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw new Error(`${label} doit être un objet.`);
	}
	return value as Record<string, unknown>;
}

function rejectUnknownDesignProperties(
	value: Record<string, unknown>,
	allowed: readonly string[],
	label: string
): void {
	const allowedSet = new Set(allowed);
	const unexpected = Object.keys(value).find(key => !allowedSet.has(key));
	if (unexpected) {
		throw new Error(`${label}.${unexpected} n’est pas une propriété autorisée.`);
	}
}

function designString(
	value: unknown,
	label: string,
	options: { required?: boolean; maxLength?: number } = {}
): string | undefined {
	if (value === undefined) {
		if (options.required) {
			throw new Error(`${label} est obligatoire.`);
		}
		return undefined;
	}
	if (typeof value !== 'string') {
		throw new Error(`${label} doit être une chaîne.`);
	}
	if (value.includes('\0')) {
		throw new Error(`${label} contient un caractère NUL interdit.`);
	}
	if (options.required && value.length === 0) {
		throw new Error(`${label} ne peut pas être vide.`);
	}
	if (
		options.maxLength !== undefined &&
		value.length > options.maxLength
	) {
		throw new Error(
			`${label} dépasse la limite de ${options.maxLength} caractères.`
		);
	}
	return value;
}

function designNumber(
	value: unknown,
	label: string,
	minimum: number,
	maximum: number,
	required = true
): number | undefined {
	if (value === undefined && !required) {
		return undefined;
	}
	if (
		typeof value !== 'number' ||
		!Number.isFinite(value) ||
		value < minimum ||
		value > maximum
	) {
		throw new Error(
			`${label} doit être un nombre fini compris entre ${minimum} et ${maximum}.`
		);
	}
	return value;
}

function designIdentifier(value: unknown, label: string): string {
	const parsed = designString(value, label, {
		required: true,
		maxLength: 31
	}) as string;
	if (!VBA_IDENTIFIER_PATTERN.test(parsed)) {
		throw new Error(`${label} n’est pas un identifiant VBA valide.`);
	}
	return parsed;
}

function parseDesignControl(
	value: unknown,
	label: string
): VbaUserFormControl {
	const source = designObject(value, label);
	rejectUnknownDesignProperties(
		source,
		[
			'type',
			'name',
			'left',
			'top',
			'width',
			'height',
			'caption',
			'enabled',
			'visible',
			'tabIndex',
			'controlTipText',
			'progId'
		],
		label
	);
	const type = designString(source.type, `${label}.type`, {
		required: true
	}) as string;
	if (
		!VBA_USERFORM_CONTROL_TYPES.has(type as VbaUserFormControlType)
	) {
		throw new Error(`${label}.type n’est pas un type de contrôle pris en charge.`);
	}
	const progId = designString(source.progId, `${label}.progId`, {
		maxLength: 128
	});
	if (type === 'customActiveX') {
		if (!progId || !ACTIVEX_PROGID_PATTERN.test(progId)) {
			throw new Error(
				`${label}.progId est obligatoire pour customActiveX et doit être un ProgID valide.`
			);
		}
	} else if (progId !== undefined) {
		throw new Error(
			`${label}.progId est accepté uniquement avec type=customActiveX.`
		);
	}
	const name = designIdentifier(source.name, `${label}.name`);
	const left = designNumber(
		source.left,
		`${label}.left`,
		0,
		MAX_DESIGN_COORDINATE
	) as number;
	const top = designNumber(
		source.top,
		`${label}.top`,
		0,
		MAX_DESIGN_COORDINATE
	) as number;
	const width = designNumber(
		source.width,
		`${label}.width`,
		Number.MIN_VALUE,
		MAX_DESIGN_COORDINATE
	) as number;
	const height = designNumber(
		source.height,
		`${label}.height`,
		Number.MIN_VALUE,
		MAX_DESIGN_COORDINATE
	) as number;
	const caption = designString(source.caption, `${label}.caption`, {
		maxLength: MAX_DESIGN_TEXT_CHARACTERS
	});
	const controlTipText = designString(
		source.controlTipText,
		`${label}.controlTipText`,
		{ maxLength: MAX_DESIGN_TEXT_CHARACTERS }
	);
	if (source.enabled !== undefined && typeof source.enabled !== 'boolean') {
		throw new Error(`${label}.enabled doit être un booléen.`);
	}
	if (source.visible !== undefined && typeof source.visible !== 'boolean') {
		throw new Error(`${label}.visible doit être un booléen.`);
	}
	let tabIndex: number | undefined;
	if (source.tabIndex !== undefined) {
		tabIndex = designNumber(
			source.tabIndex,
			`${label}.tabIndex`,
			0,
			MAX_DESIGN_TAB_INDEX
		) as number;
		if (!Number.isInteger(tabIndex)) {
			throw new Error(`${label}.tabIndex doit être un entier.`);
		}
	}
	return {
		type: type as VbaUserFormControlType,
		name,
		left,
		top,
		width,
		height,
		...(caption !== undefined ? { caption } : {}),
		...(source.enabled !== undefined
			? { enabled: source.enabled as boolean }
			: {}),
		...(source.visible !== undefined
			? { visible: source.visible as boolean }
			: {}),
		...(tabIndex !== undefined ? { tabIndex } : {}),
		...(controlTipText !== undefined ? { controlTipText } : {}),
		...(progId !== undefined ? { progId } : {})
	};
}

function parseDesignOperation(
	value: unknown,
	index: number
): VbaDesignOperation {
	const label = `operations[${index}]`;
	const source = designObject(value, label);
	const kind = designString(source.kind, `${label}.kind`, {
		required: true
	});
	if (kind === 'createUserForm') {
		rejectUnknownDesignProperties(
			source,
			['kind', 'name', 'caption', 'width', 'height', 'source', 'controls'],
			label
		);
		const name = designIdentifier(source.name, `${label}.name`);
		const caption = designString(source.caption, `${label}.caption`, {
			maxLength: MAX_DESIGN_TEXT_CHARACTERS
		});
		const width = designNumber(
			source.width,
			`${label}.width`,
			Number.MIN_VALUE,
			MAX_DESIGN_COORDINATE,
			false
		);
		const height = designNumber(
			source.height,
			`${label}.height`,
			Number.MIN_VALUE,
			MAX_DESIGN_COORDINATE,
			false
		);
		const vbaSource = designString(source.source, `${label}.source`, {
			maxLength: MAX_VBA_SOURCE_CHARACTERS
		});
		let controls: VbaUserFormControl[] | undefined;
		if (source.controls !== undefined) {
			if (!Array.isArray(source.controls)) {
				throw new Error(`${label}.controls doit être un tableau.`);
			}
			controls = source.controls.map((control, controlIndex) =>
				parseDesignControl(
					control,
					`${label}.controls[${controlIndex}]`
				)
			);
		}
		return {
			kind,
			name,
			...(caption !== undefined ? { caption } : {}),
			...(width !== undefined ? { width } : {}),
			...(height !== undefined ? { height } : {}),
			...(vbaSource !== undefined ? { source: vbaSource } : {}),
			...(controls !== undefined ? { controls } : {})
		};
	}
	if (kind === 'addUserFormControl') {
		rejectUnknownDesignProperties(
			source,
			['kind', 'formName', 'control'],
			label
		);
		if (source.control === undefined) {
			throw new Error(`${label}.control est obligatoire.`);
		}
		return {
			kind,
			formName: designIdentifier(source.formName, `${label}.formName`),
			control: parseDesignControl(source.control, `${label}.control`)
		};
	}
	if (kind === 'createWorksheetButton') {
		rejectUnknownDesignProperties(
			source,
			[
				'kind',
				'sheetName',
				'name',
				'caption',
				'macroName',
				'left',
				'top',
				'width',
				'height'
			],
			label
		);
		const macroName = designString(
			source.macroName,
			`${label}.macroName`,
			{ required: true, maxLength: 63 }
		) as string;
		if (!VBA_MACRO_PATTERN.test(macroName)) {
			throw new Error(`${label}.macroName n’est pas un nom de macro valide.`);
		}
		return {
			kind,
			sheetName: designString(source.sheetName, `${label}.sheetName`, {
				required: true,
				maxLength: MAX_DESIGN_TEXT_CHARACTERS
			}) as string,
			name: designIdentifier(source.name, `${label}.name`),
			caption: designString(source.caption, `${label}.caption`, {
				required: true,
				maxLength: MAX_DESIGN_TEXT_CHARACTERS
			}) as string,
			macroName,
			left: designNumber(
				source.left,
				`${label}.left`,
				0,
				MAX_DESIGN_COORDINATE
			) as number,
			top: designNumber(
				source.top,
				`${label}.top`,
				0,
				MAX_DESIGN_COORDINATE
			) as number,
			width: designNumber(
				source.width,
				`${label}.width`,
				Number.MIN_VALUE,
				MAX_DESIGN_COORDINATE
			) as number,
			height: designNumber(
				source.height,
				`${label}.height`,
				Number.MIN_VALUE,
				MAX_DESIGN_COORDINATE
			) as number
		};
	}
	if (
		kind === 'assignWorksheetButtonMacro' ||
		kind === 'bindWorksheetActiveXMacro'
	) {
		rejectUnknownDesignProperties(
			source,
			['kind', 'sheetName', 'name', 'macroName'],
			label
		);
		const macroName = designString(
			source.macroName,
			`${label}.macroName`,
			{ required: true, maxLength: 63 }
		) as string;
		if (!VBA_MACRO_PATTERN.test(macroName)) {
			throw new Error(`${label}.macroName n’est pas un nom de macro valide.`);
		}
		return {
			kind,
			sheetName: designString(source.sheetName, `${label}.sheetName`, {
				required: true,
				maxLength: MAX_DESIGN_TEXT_CHARACTERS
			}) as string,
			name: designIdentifier(source.name, `${label}.name`),
			macroName
		};
	}
	if (kind === 'createWorksheetActiveXControl') {
		rejectUnknownDesignProperties(
			source,
			['kind', 'sheetName', 'control'],
			label
		);
		if (source.control === undefined) {
			throw new Error(`${label}.control est obligatoire.`);
		}
		return {
			kind,
			sheetName: designString(source.sheetName, `${label}.sheetName`, {
				required: true,
				maxLength: MAX_DESIGN_TEXT_CHARACTERS
			}) as string,
			control: parseDesignControl(source.control, `${label}.control`)
		};
	}
	throw new Error(`${label}.kind n’est pas une opération prise en charge.`);
}

function parseDesignInput(value: unknown): VbaDesignToolInput {
	const source = designObject(value, 'input');
	rejectUnknownDesignProperties(
		source,
		['workbookPath', 'operations'],
		'input'
	);
	const workbookPath = designString(source.workbookPath, 'workbookPath');
	if (
		!Array.isArray(source.operations) ||
		source.operations.length < 1 ||
		source.operations.length > MAX_VBA_DESIGN_OPERATIONS
	) {
		throw new Error(
			`operations doit contenir de 1 à ${MAX_VBA_DESIGN_OPERATIONS} opérations.`
		);
	}
	const operations = source.operations.map(parseDesignOperation);
	const formNames = new Set<string>();
	const controlKeys = new Set<string>();
	const createdButtonKeys = new Set<string>();
	const assignedButtonKeys = new Set<string>();
	const createdActiveXKeys = new Set<string>();
	const boundActiveXKeys = new Set<string>();
	for (const operation of operations) {
		if (operation.kind === 'createUserForm') {
			const formKey = operation.name.toLocaleLowerCase('en-US');
			if (formNames.has(formKey)) {
				throw new Error(`UserForm demandé plusieurs fois : ${operation.name}.`);
			}
			formNames.add(formKey);
			for (const control of operation.controls || []) {
				const controlKey = `${formKey}\0${control.name.toLocaleLowerCase(
					'en-US'
				)}`;
				if (controlKeys.has(controlKey)) {
					throw new Error(
						`Contrôle demandé plusieurs fois : ${operation.name}.${control.name}.`
					);
				}
				controlKeys.add(controlKey);
			}
		} else if (operation.kind === 'addUserFormControl') {
			const controlKey = `${operation.formName.toLocaleLowerCase(
				'en-US'
			)}\0${operation.control.name.toLocaleLowerCase('en-US')}`;
			if (controlKeys.has(controlKey)) {
				throw new Error(
					`Contrôle demandé plusieurs fois : ${operation.formName}.${operation.control.name}.`
				);
			}
			controlKeys.add(controlKey);
		} else if (
			operation.kind === 'createWorksheetButton' ||
			operation.kind === 'assignWorksheetButtonMacro'
		) {
			const buttonKey = `${operation.sheetName.toLocaleLowerCase(
				'en-US'
			)}\0${operation.name.toLocaleLowerCase('en-US')}`;
			const targetSet =
				operation.kind === 'createWorksheetButton'
					? createdButtonKeys
					: assignedButtonKeys;
			if (targetSet.has(buttonKey)) {
				throw new Error(
					`Bouton demandé plusieurs fois : ${operation.sheetName}.${operation.name}.`
				);
			}
			targetSet.add(buttonKey);
		} else {
			const control = operation.kind === 'createWorksheetActiveXControl'
				? operation.control.name
				: operation.name;
			const controlKey = `${operation.sheetName.toLocaleLowerCase(
				'en-US'
			)}\0${control.toLocaleLowerCase('en-US')}`;
			const targetSet =
				operation.kind === 'createWorksheetActiveXControl'
					? createdActiveXKeys
					: boundActiveXKeys;
			if (targetSet.has(controlKey)) {
				throw new Error(
					`Contrôle ActiveX demandé plusieurs fois : ${operation.sheetName}.${control}.`
				);
			}
			targetSet.add(controlKey);
		}
	}
	return {
		...(workbookPath?.trim() ? { workbookPath } : {}),
		operations
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

	const designTool = {
		async prepareInvocation(options: { input?: unknown }) {
			const input = parseDesignInput(options?.input);
			const requestedPath = input.workbookPath?.trim();
			return {
				invocationMessage: requestedPath
					? `Création transactionnelle de composants visuels VBA dans ${path.basename(
							requestedPath
					  )}`
					: 'Création transactionnelle de composants visuels dans le classeur XLSM actif'
			};
		},

		async invoke(
			options: { input?: unknown },
			cancellationToken?: vscode.CancellationToken
		): Promise<unknown> {
			if (cancellationToken?.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const input = parseDesignInput(options?.input);
			const workbookUri = await service.resolveToolWorkbookUri(input);
			if (!workbookUri) {
				throw new Error(
					'Aucun classeur XLSM local unique ne peut recevoir les composants visuels VBA.'
				);
			}
			const designResult = await service.designVbaFromTool(
				workbookUri,
				input.operations,
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
						...designResult
					})
				)
			]);
		}
	};

	try {
		context.subscriptions.push(
			vscodeRuntime.lm.registerTool(EXCEL_AI_LANGUAGE_MODEL_TOOL, tool),
			vscodeRuntime.lm.registerTool(EXCEL_AI_VBA_WRITE_TOOL, writeTool),
			vscodeRuntime.lm.registerTool(EXCEL_AI_VBA_DESIGN_TOOL, designTool)
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
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] ${EXCEL_AI_VBA_DESIGN_TOOL} enregistré pour les UserForms et boutons transactionnels.`
			);
	} catch (error) {
		service
			.getOutputChannel()
			.appendLine(
				`[outil IA] Enregistrement ignoré : ${(error as Error).message}`
			);
	}
}
