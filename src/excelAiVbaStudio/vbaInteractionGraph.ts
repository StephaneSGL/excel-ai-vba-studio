const MAX_INTERACTIONS = 2_000;
const VBA_IDENTIFIER = '[A-Za-z_][A-Za-z0-9_]{0,30}';

export interface VbaSourceComponent {
	name: string;
	type: string;
	source: string;
}

export interface VbaMacroRecord {
	name: string;
	moduleName: string;
	qualifiedName: string;
	userFormsOpened: string[];
}

export interface WorksheetButtonRecord {
	sheetName: string;
	sheetCodeName: string;
	name: string;
	caption: string;
	onAction: string;
	left: number;
	top: number;
	width: number;
	height: number;
}

export interface WorksheetActiveXRecord {
	sheetName: string;
	sheetCodeName: string;
	name: string;
	progId: string;
	caption: string;
	enabled?: boolean;
	visible: boolean;
	left: number;
	top: number;
	width: number;
	height: number;
	macroName?: string;
}

export interface VbaInteractionRelationship {
	kind: 'formButton' | 'activeX';
	sheetName: string;
	controlName: string;
	controlCaption: string;
	macroName?: string;
	userFormsOpened: string[];
	resolution: 'resolved' | 'missing-macro' | 'complex-event' | 'unassigned';
}

export interface VbaInteractionGraph {
	macros: VbaMacroRecord[];
	userForms: string[];
	worksheetButtons: WorksheetButtonRecord[];
	worksheetActiveXControls: WorksheetActiveXRecord[];
	relationships: VbaInteractionRelationship[];
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: undefined;
}

function asString(value: unknown): string {
	return typeof value === 'string' ? value : '';
}

function asNumber(value: unknown): number {
	return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

function asBoolean(value: unknown, fallback: boolean): boolean {
	return typeof value === 'boolean' ? value : fallback;
}

function recordList(value: unknown): Record<string, unknown>[] {
	return Array.isArray(value)
		? value
				.slice(0, MAX_INTERACTIONS)
				.map(asRecord)
				.filter((record): record is Record<string, unknown> => Boolean(record))
		: [];
}

function isStandardModule(type: string): boolean {
	const normalized = type.toLocaleLowerCase('en-US');
	return normalized === 'module' || normalized.includes('standard');
}

function isUserForm(type: string): boolean {
	return type.toLocaleLowerCase('en-US').includes('userform');
}

function normalizeMacroTarget(onAction: string): string {
	const bangIndex = onAction.lastIndexOf('!');
	return (bangIndex >= 0 ? onAction.slice(bangIndex + 1) : onAction)
		.trim()
		.replace(/^'+|'+$/g, '');
}

function uniqueStrings(values: string[]): string[] {
	const seen = new Set<string>();
	return values.filter(value => {
		const key = value.toLocaleLowerCase('en-US');
		if (seen.has(key)) {
			return false;
		}
		seen.add(key);
		return true;
	});
}

function extractPublicZeroArgumentMacros(
	components: VbaSourceComponent[]
): VbaMacroRecord[] {
	const records: VbaMacroRecord[] = [];
	const procedurePattern = new RegExp(
		`^[\\t ]*(?:(Public|Private|Friend|Static)[\\t ]+)?Sub[\\t ]+(${VBA_IDENTIFIER})[\\t ]*\\([\\t ]*\\)[^\\r\\n]*(?:\\r?\\n)([\\s\\S]*?)^[\\t ]*End[\\t ]+Sub\\b`,
		'gim'
	);
	const userFormShowPattern = new RegExp(
		`^[\\t ]*(?:Call[\\t ]+)?(${VBA_IDENTIFIER})\\.Show\\b`,
		'gim'
	);
	for (const component of components.filter(item =>
		isStandardModule(item.type)
	)) {
		procedurePattern.lastIndex = 0;
		let match: RegExpExecArray | null;
		while ((match = procedurePattern.exec(component.source))) {
			const visibility = (match[1] || '').toLocaleLowerCase('en-US');
			if (visibility === 'private' || visibility === 'friend') {
				continue;
			}
			const forms: string[] = [];
			userFormShowPattern.lastIndex = 0;
			let formMatch: RegExpExecArray | null;
			while ((formMatch = userFormShowPattern.exec(match[3]))) {
				forms.push(formMatch[1]);
			}
			records.push({
				name: match[2],
				moduleName: component.name,
				qualifiedName: `${component.name}.${match[2]}`,
				userFormsOpened: uniqueStrings(forms)
			});
			if (records.length >= MAX_INTERACTIONS) {
				return records;
			}
		}
	}
	return records;
}

function resolveMacro(
	target: string,
	macros: VbaMacroRecord[]
): VbaMacroRecord | undefined {
	const normalized = normalizeMacroTarget(target).toLocaleLowerCase('en-US');
	const qualified = macros.find(
		macro => macro.qualifiedName.toLocaleLowerCase('en-US') === normalized
	);
	if (qualified) {
		return qualified;
	}
	const byName = macros.filter(
		macro => macro.name.toLocaleLowerCase('en-US') === normalized
	);
	return byName.length === 1 ? byName[0] : undefined;
}

interface ActiveXEventTarget {
	handlerExists: boolean;
	macroTarget?: string;
}

function activeXMacroTarget(
	control: WorksheetActiveXRecord,
	components: VbaSourceComponent[]
): ActiveXEventTarget {
	const document = components.find(
		component =>
			component.name.toLocaleLowerCase('en-US') ===
			control.sheetCodeName.toLocaleLowerCase('en-US')
	);
	if (!document) {
		return { handlerExists: false };
	}
	const escapedName = control.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const handlerPattern = new RegExp(
		`^[\\t ]*(?:Private|Public|Friend)?[\\t ]*Sub[\\t ]+${escapedName}_Click[\\t ]*\\([\\t ]*\\)[^\\r\\n]*(?:\\r?\\n)([\\s\\S]*?)^[\\t ]*End[\\t ]+Sub\\b`,
		'im'
	);
	const handler = handlerPattern.exec(document.source);
	if (!handler) {
		return { handlerExists: false };
	}
	const statements = handler[1]
		.split(/\r?\n/)
		.map(line => line.trim())
		.filter(line => line && !/^'/.test(line) && !/^Rem(?:\s|$)/i.test(line));
	if (statements.length !== 1) {
		return { handlerExists: true };
	}
	const directCallPattern = new RegExp(
		`^(?:Call[\\t ]+)?(${VBA_IDENTIFIER}(?:\\.${VBA_IDENTIFIER})?)[\\t ]*(?:\\(\\s*\\))?[\\t ]*(?:'.*)?$`,
		'i'
	);
	const macroTarget = directCallPattern.exec(statements[0])?.[1];
	return {
		handlerExists: true,
		...(macroTarget ? { macroTarget } : {})
	};
}

export function buildVbaInteractionGraph(
	components: VbaSourceComponent[],
	vbaData: unknown
): VbaInteractionGraph {
	const vba = asRecord(vbaData) || {};
	const macros = extractPublicZeroArgumentMacros(components);
	const userForms = uniqueStrings(
		components.filter(item => isUserForm(item.type)).map(item => item.name)
	);
	const worksheetButtons = recordList(vba.worksheetButtons).map(
		record => ({
			sheetName: asString(record.sheetName),
			sheetCodeName: asString(record.sheetCodeName),
			name: asString(record.name),
			caption: asString(record.caption),
			onAction: asString(record.onAction),
			left: asNumber(record.left),
			top: asNumber(record.top),
			width: asNumber(record.width),
			height: asNumber(record.height)
		})
	);
	const worksheetActiveXControls: WorksheetActiveXRecord[] = recordList(
		vba.worksheetActiveXControls
	).map(record => ({
		sheetName: asString(record.sheetName),
		sheetCodeName: asString(record.sheetCodeName),
		name: asString(record.name),
		progId: asString(record.progId),
		caption: asString(record.caption),
		...(typeof record.enabled === 'boolean'
			? { enabled: record.enabled }
			: {}),
		visible: asBoolean(record.visible, true),
		left: asNumber(record.left),
		top: asNumber(record.top),
		width: asNumber(record.width),
		height: asNumber(record.height)
	}));

	const relationships: VbaInteractionRelationship[] = [];
	for (const button of worksheetButtons) {
		const macro = button.onAction
			? resolveMacro(button.onAction, macros)
			: undefined;
		relationships.push({
			kind: 'formButton',
			sheetName: button.sheetName,
			controlName: button.name,
			controlCaption: button.caption || button.name,
			...(macro ? { macroName: macro.qualifiedName } : {}),
			userFormsOpened: macro?.userFormsOpened || [],
			resolution: macro
				? 'resolved'
				: button.onAction
					? 'missing-macro'
					: 'unassigned'
		});
	}
	for (const control of worksheetActiveXControls) {
		const eventTarget = activeXMacroTarget(control, components);
		const macro = eventTarget.macroTarget
			? resolveMacro(eventTarget.macroTarget, macros)
			: undefined;
		if (macro) {
			control.macroName = macro.qualifiedName;
		}
		relationships.push({
			kind: 'activeX',
			sheetName: control.sheetName,
			controlName: control.name,
			controlCaption: control.caption || control.name,
			...(macro ? { macroName: macro.qualifiedName } : {}),
			userFormsOpened: macro?.userFormsOpened || [],
			resolution: macro
				? 'resolved'
				: eventTarget.macroTarget
					? 'missing-macro'
					: eventTarget.handlerExists
						? 'complex-event'
						: 'unassigned'
		});
	}

	return {
		macros,
		userForms,
		worksheetButtons,
		worksheetActiveXControls,
		relationships
	};
}
