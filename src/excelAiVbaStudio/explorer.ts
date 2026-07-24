import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { pathIsInside } from './security';
import {
	EXCEL_AI_COMMANDS,
	ExplorerTreeItem,
	ExportContext
} from './types';
import { ExcelAiVbaWorkbookService } from './workbookService';

type UnknownRecord = Record<string, unknown>;

function getNestedValue(source: unknown, keys: string[]): unknown {
	let value = source;
	for (const key of keys) {
		if (!value || typeof value !== 'object') {
			return undefined;
		}
		value = (value as UnknownRecord)[key];
	}
	return value;
}

function firstDefined(source: unknown, paths: string[][]): unknown {
	for (const keys of paths) {
		const value = getNestedValue(source, keys);
		if (value !== undefined && value !== null) {
			return value;
		}
	}
	return undefined;
}

function toRecordList(value: unknown): unknown[] {
	if (Array.isArray(value)) {
		return value;
	}
	if (!value || typeof value !== 'object') {
		return value === undefined || value === null ? [] : [value];
	}
	return Object.entries(value as UnknownRecord).map(([key, item]) => {
		if (item && typeof item === 'object' && !Array.isArray(item)) {
			return { __key: key, ...(item as UnknownRecord) };
		}
		return { __key: key, value: item };
	});
}

function recordLabel(record: unknown, fallback: string): string {
	if (record === undefined || record === null) {
		return fallback;
	}
	if (typeof record !== 'object') {
		return String(record);
	}
	const candidate = record as UnknownRecord;
	for (const key of [
		'name',
		'Name',
		'title',
		'sheetName',
		'moduleName',
		'queryName',
		'connectionName',
		'__key'
	]) {
		if (candidate[key] !== undefined && candidate[key] !== null) {
			return String(candidate[key]);
		}
	}
	return fallback;
}

function recordDescription(
	record: unknown,
	keys: string[]
): string | undefined {
	if (!record || typeof record !== 'object') {
		return undefined;
	}
	const candidate = record as UnknownRecord;
	for (const key of keys) {
		const value = candidate[key];
		if (value !== undefined && value !== null && String(value).trim()) {
			return String(value);
		}
	}
	return undefined;
}

function createTreeItem(
	label: string,
	options: {
		children?: ExplorerTreeItem[];
		description?: string;
		tooltip?: string;
		contextValue?: string;
		icon?: string;
		command?: string;
		arguments?: unknown[];
		resourceUri?: vscode.Uri;
	} = {}
): ExplorerTreeItem {
	const item = new vscode.TreeItem(
		label,
		options.children?.length
			? vscode.TreeItemCollapsibleState.Collapsed
			: vscode.TreeItemCollapsibleState.None
	) as ExplorerTreeItem;
	item.children = options.children;
	item.description = options.description;
	item.tooltip = options.tooltip;
	item.contextValue = options.contextValue;
	item.iconPath = options.icon ? new vscode.ThemeIcon(options.icon) : undefined;
	item.resourceUri = options.resourceUri;
	if (options.command) {
		item.command = {
			command: options.command,
			title: label,
			arguments: options.arguments || []
		};
	}
	return item;
}

function actionItem(
	label: string,
	command: string,
	icon: string,
	workbookUri?: vscode.Uri
): ExplorerTreeItem {
	return createTreeItem(label, {
		command,
		arguments: workbookUri ? [workbookUri] : [],
		icon,
		contextValue: 'excelAiVbaAction'
	});
}

async function findVbaSourceFiles(context: ExportContext): Promise<string[]> {
	try {
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
			.map(entry => path.join(context.paths.vbaDirectory, entry.name))
			.filter(candidate =>
				pathIsInside(candidate, context.paths.vbaDirectory)
			);
	} catch {
		return [];
	}
}

function normalizeComparableName(value: unknown): string {
	return String(value)
		.normalize('NFKD')
		.replace(/[\u0300-\u036f]/g, '')
		.replace(/[^a-z0-9]/gi, '')
		.toLocaleLowerCase('en-US');
}

function moduleFileForRecord(
	record: unknown,
	sourceFiles: string[]
): string | undefined {
	const wantedName = normalizeComparableName(recordLabel(record, ''));
	return sourceFiles.find(
		sourcePath =>
			normalizeComparableName(
				path.basename(sourcePath, path.extname(sourcePath))
			) === wantedName
	);
}

function section(
	label: string,
	icon: string,
	children: ExplorerTreeItem[]
): ExplorerTreeItem | undefined {
	if (!children.length) {
		return undefined;
	}
	return createTreeItem(label, {
		description: String(children.length),
		children,
		icon
	});
}

async function workbookSections(
	data: UnknownRecord,
	context: ExportContext
): Promise<ExplorerTreeItem[]> {
	const sections: Array<ExplorerTreeItem | undefined> = [];
	const workbook =
		data.workbook && typeof data.workbook === 'object'
			? (data.workbook as UnknownRecord)
			: data;
	const sheets = toRecordList(
		firstDefined(data, [['worksheets'], ['sheets'], ['Worksheets']]) ||
			firstDefined(workbook, [['worksheets'], ['sheets'], ['Worksheets']])
	);

	sections.push(
		section(
			'Feuilles',
			'layers',
			sheets.map((sheet, index) => {
				const usedRange =
					sheet &&
					typeof sheet === 'object' &&
					(sheet as UnknownRecord).usedRange &&
					typeof (sheet as UnknownRecord).usedRange === 'object'
						? ((sheet as UnknownRecord).usedRange as UnknownRecord)
						: undefined;
				return createTreeItem(recordLabel(sheet, `Feuille ${index + 1}`), {
					description:
						recordDescription(sheet, ['range', 'address', 'dimensions']) ||
						recordDescription(usedRange, ['address']),
					icon: 'table',
					contextValue: 'excelAiVbaSheet'
				});
			})
		)
	);

	const directTables = toRecordList(
		firstDefined(workbook, [['tables'], ['listObjects']])
	);
	const sheetTables = sheets.flatMap(sheet => {
		const record =
			sheet && typeof sheet === 'object' ? (sheet as UnknownRecord) : undefined;
		return toRecordList(record?.tables || record?.listObjects);
	});
	sections.push(
		section(
			'Tableaux',
			'list-tree',
			[...directTables, ...sheetTables].map((table, index) =>
				createTreeItem(recordLabel(table, `Tableau ${index + 1}`), {
					description: recordDescription(table, [
						'range',
						'address',
						'sheetName'
					]),
					icon: 'layout',
					contextValue: 'excelAiVbaTable'
				})
			)
		)
	);

	const names = toRecordList(
		firstDefined(workbook, [['names'], ['namedRanges'], ['definedNames']])
	);
	sections.push(
		section(
			'Noms définis',
			'symbol-key',
			names.map((name, index) =>
				createTreeItem(recordLabel(name, `Nom ${index + 1}`), {
					description: recordDescription(name, [
						'refersTo',
						'formula',
						'address',
						'value'
					]),
					icon: 'symbol-constant',
					contextValue: 'excelAiVbaName'
				})
			)
		)
	);

	const queries = toRecordList(
		firstDefined(workbook, [['queries'], ['powerQueries']])
	);
	const connections = toRecordList(
		firstDefined(workbook, [['connections'], ['dataConnections']])
	);
	sections.push(
		section(
			'Requêtes et connexions',
			'server-process',
			[
				...queries.map((query, index) =>
					createTreeItem(recordLabel(query, `Requête ${index + 1}`), {
						description:
							recordDescription(query, [
								'type',
								'connection',
								'description'
							]) || 'Requête',
						icon: 'database',
						contextValue: 'excelAiVbaQuery'
					})
				),
				...connections.map((connection, index) =>
					createTreeItem(
						recordLabel(connection, `Connexion ${index + 1}`),
						{
							description:
								recordDescription(connection, [
									'type',
									'provider',
									'description'
								]) || 'Connexion',
							icon: 'plug',
							contextValue: 'excelAiVbaConnection'
						}
					)
				)
			]
		)
	);

	const sourceFiles = await findVbaSourceFiles(context);
	const modules = toRecordList(
		firstDefined(data, [
			['workbook', 'vba', 'modules'],
			['workbook', 'VBA', 'modules'],
			['vba', 'modules'],
			['VBA', 'modules'],
			['vbaModules'],
			['modules']
		])
	);
	const effectiveModules = modules.length
		? modules
		: sourceFiles.map(sourcePath => ({
				name: path.basename(sourcePath, path.extname(sourcePath))
		  }));
	sections.push(
		section(
			'Modules VBA',
			'symbol-class',
			effectiveModules.map((module, index) => {
				const modulePath = moduleFileForRecord(module, sourceFiles);
				return createTreeItem(recordLabel(module, `Module ${index + 1}`), {
					description: recordDescription(module, [
						'type',
						'kind',
						'moduleType'
					]),
					icon: 'symbol-method',
					contextValue: modulePath
						? 'excelAiVbaModuleFile'
						: 'excelAiVbaModule',
					command: modulePath ? 'vscode.open' : undefined,
					arguments: modulePath ? [vscode.Uri.file(modulePath)] : undefined,
					resourceUri: modulePath ? vscode.Uri.file(modulePath) : undefined,
					tooltip: modulePath
				});
			})
		)
	);

	const references = toRecordList(
		firstDefined(data, [
			['workbook', 'vba', 'references'],
			['workbook', 'VBA', 'references'],
			['vba', 'references'],
			['VBA', 'references'],
			['vbaReferences']
		])
	);
	sections.push(
		section(
			'Références VBA',
			'references',
			references.map((reference, index) =>
				createTreeItem(recordLabel(reference, `Référence ${index + 1}`), {
					description: recordDescription(reference, [
						'version',
						'fullPath',
						'path',
						'guid'
					]),
					icon: 'references',
					contextValue: 'excelAiVbaReference'
				})
			)
		)
	);

	const warnings = [
		...toRecordList(data.warnings),
		...toRecordList(getNestedValue(data, ['workbook', 'vba', 'warnings'])),
		...toRecordList(getNestedValue(data, ['vba', 'warnings']))
	];
	sections.push(
		section(
			'Avertissements',
			'warning',
			warnings.map((warning, index) =>
				createTreeItem(recordLabel(warning, `Avertissement ${index + 1}`), {
					description: recordDescription(warning, [
						'message',
						'description'
					]),
					tooltip:
						typeof warning === 'string'
							? warning
							: recordDescription(warning, ['message', 'description']),
					icon: 'warning',
					contextValue: 'excelAiVbaWarning'
				})
			)
		)
	);

	return sections.filter(
		(candidate): candidate is ExplorerTreeItem => Boolean(candidate)
	);
}

export class ExcelAiVbaExplorerProvider
	implements vscode.TreeDataProvider<ExplorerTreeItem>, vscode.Disposable
{
	private readonly changeEmitter = new vscode.EventEmitter<
		ExplorerTreeItem | undefined
	>();
	readonly onDidChangeTreeData = this.changeEmitter.event;

	constructor(private readonly service: ExcelAiVbaWorkbookService) {}

	refresh(): void {
		this.changeEmitter.fire(undefined);
	}

	dispose(): void {
		this.changeEmitter.dispose();
	}

	getTreeItem(element: ExplorerTreeItem): vscode.TreeItem {
		return element;
	}

	async getChildren(element?: ExplorerTreeItem): Promise<ExplorerTreeItem[]> {
		if (element?.children) {
			return element.children;
		}
		if (element) {
			return [];
		}

		const workbookUri = await this.service.resolveWorkbookUri();
		const actions = [
			actionItem(
				'Exporter le contexte IA',
				EXCEL_AI_COMMANDS.exportWorkbook,
				'sparkle-filled',
				workbookUri
			),
			actionItem(
				'Ouvrir dans Microsoft Excel',
				EXCEL_AI_COMMANDS.openFullExcel,
				'window',
				workbookUri
			),
			actionItem(
				'Ouvrir l’éditeur VBA',
				EXCEL_AI_COMMANDS.openVbaDeveloper,
				'code',
				workbookUri
			),
			actionItem(
				'Extraire et révéler les sources VBA',
				EXCEL_AI_COMMANDS.openVbaExplorer,
				'folder-opened',
				workbookUri
			),
			actionItem(
				'Nettoyer les exports générés',
				EXCEL_AI_COMMANDS.cleanExports,
				'trash'
			)
		];
		const rootItems = [
			createTreeItem('Actions', {
				children: actions,
				icon: 'tools'
			})
		];

		if (!workbookUri) {
			rootItems.push(
				createTreeItem('Ouvrez un classeur Excel pour l’inspecter', {
					icon: 'info',
					contextValue: 'excelAiVbaInfo'
				})
			);
			return rootItems;
		}

		rootItems.unshift(
			createTreeItem(path.basename(workbookUri.fsPath), {
				description: 'classeur actif',
				tooltip: workbookUri.fsPath,
				icon: 'file',
				resourceUri: workbookUri,
				contextValue: 'excelAiVbaWorkbook'
			})
		);

		const context = this.service.getLastContext();
		if (
			!context ||
			!sameWorkbook(context.workbookUri.fsPath, workbookUri.fsPath)
		) {
			rootItems.push(
				createTreeItem('Exportez le classeur pour afficher sa structure', {
					command: EXCEL_AI_COMMANDS.exportWorkbook,
					arguments: [workbookUri],
					icon: 'cloud-download',
					contextValue: 'excelAiVbaInfo'
				})
			);
			return rootItems;
		}

		try {
			const jsonText = await this.service.readExportedContext(context, 'json');
			const data = JSON.parse(jsonText) as UnknownRecord;
			const sections = await workbookSections(data, context);
			rootItems.push(
				...(sections.length
					? sections
					: [
							createTreeItem(
								'Le contexte exporté ne contient aucun index détaillé',
								{
									icon: 'info',
									contextValue: 'excelAiVbaInfo'
								}
							)
					  ])
			);
		} catch {
			rootItems.push(
				createTreeItem('Réexportez le classeur pour actualiser sa structure', {
					command: EXCEL_AI_COMMANDS.exportWorkbook,
					arguments: [workbookUri],
					icon: 'refresh',
					contextValue: 'excelAiVbaInfo'
				})
			);
		}
		return rootItems;
	}
}

function sameWorkbook(left: string, right: string): boolean {
	return (
		path.resolve(left).toLocaleLowerCase('en-US') ===
		path.resolve(right).toLocaleLowerCase('en-US')
	);
}
