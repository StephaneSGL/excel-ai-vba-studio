import { execFile } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';
import { canonicalizeWorkbookUri } from './security';

export type OfficeSecuritySource =
	| 'machinePolicy'
	| 'userPolicy'
	| 'cloudPolicy'
	| 'userPreference'
	| 'machinePreference';

export type OfficeSecurityStatus =
	| 'protected'
	| 'blocked'
	| 'prompt'
	| 'allowed'
	| 'managed'
	| 'warning'
	| 'unknown'
	| 'notApplicable';

export type SecurityPresenceStatus = 'present' | 'absent' | 'unknown';
export type ZoneReadStatus = 'absent' | 'read' | 'unreadable' | 'unsupported';

export interface OfficeSecuritySetting {
	id: string;
	category: string;
	source: OfficeSecuritySource;
	managed: boolean;
	registryPath: string;
	name: string;
	value: string | number | boolean | null;
	valueKind?: string;
	registryView?: string;
}

export interface OfficeTrustedLocation {
	source: OfficeSecuritySource;
	managed: boolean;
	registryPath: string;
	path: string;
	allowSubfolders: boolean;
	description?: string;
}

export interface OfficeSecurityProbe {
	schemaVersion: 1;
	inspectedAtUtc: string;
	workbook: {
		path: string;
		name: string;
		extension: string;
		sizeBytes: number;
		sha256: string;
		readOnly: boolean;
		efsEncrypted: boolean;
		officePackageEncrypted: boolean;
		zoneId: number | null;
		zoneStatus: ZoneReadStatus;
		containerKind: 'zip' | 'compound' | 'other';
		hasVbaProject: boolean;
		hasVbaSignature: boolean;
		hasPackageSignature: boolean;
		vbaSignatureStatus: SecurityPresenceStatus;
		packageSignatureStatus: SecurityPresenceStatus;
		vbaProjectProtectionStatus: SecurityPresenceStatus;
		sensitivityLabelIds: string[];
	};
	office: {
		version: string;
		settings: OfficeSecuritySetting[];
		trustedLocations: OfficeTrustedLocation[];
		cloudPolicyDetected: boolean;
		cloudPolicyServiceDetected: boolean;
	};
}

export interface OfficeSecurityFinding {
	id: string;
	title: string;
	status: OfficeSecurityStatus;
	detail: string;
	impact: string;
	managed: boolean;
	source: string;
}

export interface OfficeSecurityCapability {
	id: string;
	title: string;
	status: OfficeSecurityStatus;
	detail: string;
}

export interface EnterpriseSecurityReport {
	probe: OfficeSecurityProbe;
	level: 'restricted' | 'managed' | 'standard' | 'unknown';
	summary: string;
	workbookInTrustedLocation: boolean;
	findings: OfficeSecurityFinding[];
	capabilities: OfficeSecurityCapability[];
}

const MACRO_EXTENSIONS = new Set(['.xlsm', '.xlsb', '.xlam', '.xla']);
const MAX_PROBE_OUTPUT_BYTES = 256 * 1024;
const SECURITY_PROBE_TRANSPORT_BYTES = MAX_PROBE_OUTPUT_BYTES + 16 * 1024;
const SECURITY_PROBE_TIMEOUT_MS = 45_000;

function getPowerShellPath(): string {
	return process.env.SystemRoot
		? path.join(
				process.env.SystemRoot,
				'System32',
				'WindowsPowerShell',
				'v1.0',
				'powershell.exe'
		  )
		: 'powershell.exe';
}

function probeFailureMessage(
	stdout: string,
	stderr: string,
	fallback: string
): string {
	const resultLine = stdout
		.replace(/\r/g, '')
		.split('\n')
		.map(line => line.trim())
		.filter(Boolean)
		.pop();
	if (resultLine) {
		try {
			const root = JSON.parse(resultLine) as {
				error?: { message?: unknown };
			};
			if (typeof root.error?.message === 'string' && root.error.message.trim()) {
				return root.error.message.trim();
			}
		} catch {
			// The generic process error below remains the safe fallback.
		}
	}
	return stderr.trim() || fallback;
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: undefined;
}

function asString(value: unknown, fallback = ''): string {
	return typeof value === 'string' ? value : fallback;
}

function asBoolean(value: unknown, fallback = false): boolean {
	return typeof value === 'boolean' ? value : fallback;
}

function asFiniteNumber(value: unknown, fallback = 0): number {
	return typeof value === 'number' && Number.isFinite(value)
		? value
		: fallback;
}

function asNullableInteger(value: unknown): number | null {
	return typeof value === 'number' && Number.isInteger(value)
		? value
		: null;
}

function asPresenceStatus(
	value: unknown,
	fallback: SecurityPresenceStatus
): SecurityPresenceStatus {
	return ['present', 'absent', 'unknown'].includes(String(value))
		? (value as SecurityPresenceStatus)
		: fallback;
}

function asZoneReadStatus(value: unknown, zoneId: number | null): ZoneReadStatus {
	return ['absent', 'read', 'unreadable', 'unsupported'].includes(String(value))
		? (value as ZoneReadStatus)
		: zoneId === null
			? 'absent'
			: 'read';
}

function asStringArray(value: unknown, maximum = 128): string[] {
	if (!Array.isArray(value)) {
		return [];
	}
	return value
		.filter(item => typeof item === 'string' && item.length > 0)
		.slice(0, maximum) as string[];
}

function uniqueBy<T>(items: T[], keyOf: (item: T) => string): T[] {
	const seen = new Set<string>();
	return items.filter(item => {
		const key = keyOf(item);
		if (seen.has(key)) {
			return false;
		}
		seen.add(key);
		return true;
	});
}

function securitySource(value: unknown): OfficeSecuritySource | undefined {
	return [
		'machinePolicy',
		'userPolicy',
		'cloudPolicy',
		'userPreference',
		'machinePreference'
	].includes(String(value))
		? (value as OfficeSecuritySource)
		: undefined;
}

function parseSetting(value: unknown): OfficeSecuritySetting | undefined {
	const record = asRecord(value);
	const source = securitySource(record?.source);
	if (!record || !source) {
		return undefined;
	}
	const id = asString(record.id).slice(0, 100);
	const name = asString(record.name).slice(0, 200);
	const registryPath = asString(record.registryPath).slice(0, 1000);
	if (!id || !name || !registryPath) {
		return undefined;
	}
	const rawValue = record.value;
	const normalizedValue: string | number | boolean | null =
		rawValue === null ||
		typeof rawValue === 'string' ||
		typeof rawValue === 'number' ||
		typeof rawValue === 'boolean'
			? (rawValue as string | number | boolean | null)
			: String(rawValue ?? '');
	return {
		id,
		category: asString(record.category, 'other').slice(0, 100),
		source,
		managed: asBoolean(record.managed, source.endsWith('Policy')),
		registryPath,
		name,
		value: normalizedValue,
		valueKind: asString(record.valueKind).slice(0, 100) || undefined,
		registryView: asString(record.registryView).slice(0, 100) || undefined
	};
}

function parseTrustedLocation(value: unknown): OfficeTrustedLocation | undefined {
	const record = asRecord(value);
	const source = securitySource(record?.source);
	if (!record || !source) {
		return undefined;
	}
	const locationPath = asString(record.path).slice(0, 4000);
	const registryPath = asString(record.registryPath).slice(0, 1000);
	if (!locationPath || !registryPath) {
		return undefined;
	}
	return {
		source,
		managed: asBoolean(record.managed, source.endsWith('Policy')),
		registryPath,
		path: locationPath,
		allowSubfolders: asBoolean(record.allowSubfolders),
		description: asString(record.description).slice(0, 1000) || undefined
	};
}

export function parseOfficeSecurityProbe(value: unknown): OfficeSecurityProbe {
	const root = asRecord(value);
	const workbook = asRecord(root?.workbook);
	const office = asRecord(root?.office);
	if (root?.schemaVersion !== 1 || !workbook || !office) {
		throw new Error('Le diagnostic de sécurité a renvoyé un schéma inattendu.');
	}
	const workbookPath = asString(workbook.path);
	const workbookName = asString(workbook.name);
	const extension = asString(workbook.extension).toLocaleLowerCase('en-US');
	const sha256 = asString(workbook.sha256).toLocaleLowerCase('en-US');
	if (
		!workbookPath ||
		!workbookName ||
		!extension ||
		!/^[a-f0-9]{64}$/.test(sha256)
	) {
		throw new Error('Le diagnostic de sécurité ne confirme pas le classeur inspecté.');
	}
	const zoneId = asNullableInteger(workbook.zoneId);
	const hasVbaSignature = asBoolean(workbook.hasVbaSignature);
	const hasPackageSignature = asBoolean(workbook.hasPackageSignature);
	const containerKind = ['zip', 'compound', 'other'].includes(
		String(workbook.containerKind)
	)
		? (workbook.containerKind as 'zip' | 'compound' | 'other')
		: 'other';
	const parsedSettings = Array.isArray(office.settings)
		? office.settings
				.map(parseSetting)
				.filter((item): item is OfficeSecuritySetting => Boolean(item))
				.slice(0, 256)
		: [];
	const parsedTrustedLocations = Array.isArray(office.trustedLocations)
		? office.trustedLocations
				.map(parseTrustedLocation)
				.filter((item): item is OfficeTrustedLocation => Boolean(item))
				.slice(0, 64)
		: [];
	return {
		schemaVersion: 1,
		inspectedAtUtc: asString(root.inspectedAtUtc),
		workbook: {
			path: workbookPath,
			name: workbookName,
			extension,
			sizeBytes: Math.max(0, asFiniteNumber(workbook.sizeBytes)),
			sha256,
			readOnly: asBoolean(workbook.readOnly),
			efsEncrypted: asBoolean(workbook.efsEncrypted),
			officePackageEncrypted: asBoolean(workbook.officePackageEncrypted),
			zoneId,
			zoneStatus: asZoneReadStatus(workbook.zoneStatus, zoneId),
			containerKind,
			hasVbaProject: asBoolean(workbook.hasVbaProject),
			hasVbaSignature,
			hasPackageSignature,
			vbaSignatureStatus: asPresenceStatus(
				workbook.vbaSignatureStatus,
				hasVbaSignature ? 'present' : 'absent'
			),
			packageSignatureStatus: asPresenceStatus(
				workbook.packageSignatureStatus,
				hasPackageSignature ? 'present' : 'absent'
			),
			vbaProjectProtectionStatus: asPresenceStatus(
				workbook.vbaProjectProtectionStatus,
				asBoolean(workbook.hasVbaProject) ? 'unknown' : 'absent'
			),
			sensitivityLabelIds: asStringArray(workbook.sensitivityLabelIds, 32)
		},
		office: {
			version: asString(office.version, '16.0').slice(0, 20),
			settings: uniqueBy(parsedSettings, setting =>
				[
					setting.source,
					setting.id,
					setting.registryPath.toLocaleLowerCase('en-US'),
					setting.name.toLocaleLowerCase('en-US'),
					JSON.stringify(setting.value)
				].join('|')
			),
			trustedLocations: uniqueBy(parsedTrustedLocations, location =>
				[
					location.source,
					location.registryPath.toLocaleLowerCase('en-US'),
					location.path.toLocaleLowerCase('en-US'),
					String(location.allowSubfolders)
				].join('|')
			),
			cloudPolicyDetected: asBoolean(office.cloudPolicyDetected),
			cloudPolicyServiceDetected: asBoolean(
				office.cloudPolicyServiceDetected
			)
		}
	};
}

function numericValue(setting: OfficeSecuritySetting | undefined): number | undefined {
	if (!setting) {
		return undefined;
	}
	const parsed = Number(setting.value);
	return Number.isFinite(parsed) ? parsed : undefined;
}

interface SettingResolution {
	setting?: OfficeSecuritySetting;
	candidates: OfficeSecuritySetting[];
	conflict: boolean;
	managed: boolean;
}

function resolveSetting(
	settings: OfficeSecuritySetting[],
	id: string
): SettingResolution {
	const candidates = settings.filter(setting => setting.id === id);
	const cloud = candidates.filter(setting => setting.source === 'cloudPolicy');
	const groupPolicy = candidates.filter(setting =>
		['machinePolicy', 'userPolicy'].includes(setting.source)
	);
	const preferences = candidates.filter(setting =>
		['userPreference', 'machinePreference'].includes(setting.source)
	);
	const applicable = cloud.length > 0
		? cloud
		: groupPolicy.length > 0
			? groupPolicy
			: preferences;
	const distinctValues = new Set(
		applicable.map(setting => JSON.stringify(setting.value))
	);
	const conflict = distinctValues.size > 1;
	return {
		setting: conflict ? undefined : applicable[0],
		candidates: applicable,
		conflict,
		managed: applicable.some(setting => setting.managed)
	};
}

function sourceLabel(resolution: SettingResolution): string {
	if (resolution.conflict) {
		return 'Valeurs contradictoires détectées · décision Excel à confirmer';
	}
	const setting = resolution.setting;
	if (!setting) {
		return 'Valeur Office par défaut ou indéterminée';
	}
	switch (setting.source) {
		case 'machinePolicy':
			return 'Stratégie gérée · ordinateur';
		case 'userPolicy':
			return 'Stratégie gérée · utilisateur';
		case 'cloudPolicy':
			return 'Stratégie Microsoft 365 Cloud Policy';
		case 'machinePreference':
			return 'Préférence locale · ordinateur';
		default:
			return 'Préférence locale · utilisateur';
	}
}

function expandWindowsEnvironment(value: string): string {
	return value.replace(/%([^%]+)%/g, (match, name: string) => {
		const found = Object.entries(process.env).find(
			([key]) => key.toLocaleLowerCase('en-US') === name.toLocaleLowerCase('en-US')
		)?.[1];
		return found || match;
	});
}

function pathInside(candidatePath: string, location: OfficeTrustedLocation): boolean {
	const rawExpanded = expandWindowsEnvironment(location.path);
	const parsedRoot = path.win32.parse(rawExpanded).root;
	const expanded = rawExpanded.length === parsedRoot.length
		? parsedRoot
		: rawExpanded.replace(/[\\/]+$/, '');
	if (!path.win32.isAbsolute(expanded)) {
		return false;
	}
	const candidate = path.win32.resolve(candidatePath).toLocaleLowerCase('en-US');
	const root = path.win32.resolve(expanded).toLocaleLowerCase('en-US');
	const descendantPrefix = root.endsWith('\\') ? root : `${root}\\`;
	return location.allowSubfolders
		? candidate === root || candidate.startsWith(descendantPrefix)
		: path.win32.dirname(candidate) === root;
}

function zoneLabel(zoneId: number | null): string {
	switch (zoneId) {
		case 0:
			return 'Ordinateur local';
		case 1:
			return 'Intranet local';
		case 2:
			return 'Sites approuvés';
		case 3:
			return 'Internet';
		case 4:
			return 'Sites sensibles';
		default:
			return 'Aucune marque d’origine détectée';
	}
}

export function buildEnterpriseSecurityReport(
	probe: OfficeSecurityProbe
): EnterpriseSecurityReport {
	const settings = probe.office.settings;
	const macroResolution = resolveSetting(settings, 'vbaWarnings');
	const accessVbomResolution = resolveSetting(settings, 'accessVbom');
	const blockInternetResolution = resolveSetting(settings, 'blockInternetMacros');
	const disableActiveXResolution = resolveSetting(settings, 'disableAllActiveX');
	const ufiControlsResolution = resolveSetting(settings, 'ufiControls');
	const disableInternetPvResolution = resolveSetting(
		settings,
		'disableInternetFilesInProtectedView'
	);
	const disableAllTrustedResolution = resolveSetting(
		settings,
		'disableAllTrustedLocations'
	);
	const allowUserTrustedResolution = resolveSetting(
		settings,
		'allowUserTrustedLocations'
	);
	const macroSetting = macroResolution.setting;
	const accessVbomSetting = accessVbomResolution.setting;
	const blockInternetSetting = blockInternetResolution.setting;
	const disableActiveXSetting = disableActiveXResolution.setting;
	const ufiControlsSetting = ufiControlsResolution.setting;
	const disableInternetPv = disableInternetPvResolution.setting;
	const disableAllTrusted = disableAllTrustedResolution.setting;
	const allowUserTrusted = allowUserTrustedResolution.setting;
	const managed = settings.some(setting => setting.managed) ||
		probe.office.cloudPolicyDetected;
	const matchingTrustedLocations = probe.office.trustedLocations.filter(location =>
		pathInside(probe.workbook.path, location)
	);
	const matchingManagedLocation = matchingTrustedLocations.find(
		location => location.managed
	);
	const matchingLocalLocation = matchingTrustedLocations.find(
		location => !location.managed
	);
	const trustedLocationsDisabled = numericValue(disableAllTrusted) === 1;
	const userTrustedLocationsDisabled =
		!allowUserTrustedResolution.conflict &&
		numericValue(allowUserTrusted) === 0;
	const trustedLocationUncertain =
		disableAllTrustedResolution.conflict ||
		Boolean(matchingLocalLocation && allowUserTrustedResolution.conflict);
	const workbookInTrustedLocation = Boolean(
		!trustedLocationsDisabled &&
			!disableAllTrustedResolution.conflict &&
			(matchingManagedLocation ||
				(!userTrustedLocationsDisabled &&
					!allowUserTrustedResolution.conflict &&
					matchingLocalLocation))
	);
	const macroCapable =
		MACRO_EXTENSIONS.has(probe.workbook.extension) ||
		probe.workbook.hasVbaProject;
	const isXlsx = probe.workbook.extension === '.xlsx';
	const isXlsm = probe.workbook.extension === '.xlsm';
	const isLegacyXls = probe.workbook.extension === '.xls';
	const gridSupported = isXlsx || isXlsm || isLegacyXls;
	const encryptedPackage = probe.workbook.officePackageEncrypted;
	const vbaSignaturePresent =
		probe.workbook.vbaSignatureStatus === 'present';
	const vbaSignatureUnknown =
		probe.workbook.vbaSignatureStatus === 'unknown';
	const internetOrigin =
		probe.workbook.zoneStatus === 'read' &&
		(probe.workbook.zoneId === 3 || probe.workbook.zoneId === 4);
	const originUnknown =
		probe.workbook.zoneStatus === 'unreadable' ||
		probe.workbook.zoneStatus === 'unsupported' ||
		(probe.workbook.zoneStatus === 'read' && probe.workbook.zoneId === null);
	const explicitInternetMacroBlock = Boolean(
		macroCapable &&
			internetOrigin &&
			!workbookInTrustedLocation &&
			!blockInternetResolution.conflict &&
			numericValue(blockInternetSetting) === 1 &&
			probe.workbook.vbaSignatureStatus === 'absent'
	);
	const internetMacroDecisionUnknown = Boolean(
		macroCapable &&
			internetOrigin &&
			!workbookInTrustedLocation &&
			!explicitInternetMacroBlock &&
			(blockInternetResolution.conflict ||
				numericValue(blockInternetSetting) === undefined ||
				vbaSignaturePresent ||
				vbaSignatureUnknown)
	);

	const macroValue = numericValue(macroSetting);
	const macroDescriptions: Record<number, [OfficeSecurityStatus, string]> = {
		1: ['warning', 'Toutes les macros VBA sont autorisées.'],
		2: ['prompt', 'Macros désactivées avec notification.'],
		3: ['protected', 'Seules les macros signées par un éditeur approuvé sont autorisées.'],
		4: ['blocked', 'Macros désactivées sans notification.']
	};
	const macroDescription = macroResolution.conflict
		? (['unknown', 'Des valeurs de sécurité macro contradictoires sont visibles.'] as const)
		: macroValue === undefined
		? (['prompt', 'Aucun réglage explicite détecté ; Excel applique sa valeur par défaut.'] as const)
		: macroDescriptions[macroValue] ||
			(['unknown', `Valeur VBAWarnings inconnue : ${macroValue}.`] as const);
	const accessVbomValue = numericValue(accessVbomSetting);
	const accessVbomAllowed = !accessVbomResolution.conflict && accessVbomValue === 1;
	const activeXValue = numericValue(disableActiveXSetting);
	const activeXBlocked =
		!disableActiveXResolution.conflict && activeXValue === 1;
	const protectedViewDisabled =
		!disableInternetPvResolution.conflict &&
		numericValue(disableInternetPv) === 1;
	const vbaProtectionPresent =
		probe.workbook.vbaProjectProtectionStatus === 'present';
	const vbaProtectionUnknown =
		probe.workbook.vbaProjectProtectionStatus === 'unknown';
	const packageSignaturePresent =
		probe.workbook.packageSignatureStatus === 'present';
	const packageSignatureUnknown =
		probe.workbook.packageSignatureStatus === 'unknown';
	const immutableWorkbook = probe.workbook.readOnly || encryptedPackage;

	const findings: OfficeSecurityFinding[] = [
		{
			id: 'origin',
			title: 'Origine du fichier',
			status: originUnknown
				? 'unknown'
				: internetOrigin
					? 'warning'
					: probe.workbook.zoneId === 2
						? 'allowed'
						: 'protected',
			detail: originUnknown
				? 'Le flux d’origine Windows est illisible ou non pris en charge.'
				: probe.workbook.zoneStatus === 'absent'
					? 'Aucune marque d’origine Windows n’est présente.'
					: `${zoneLabel(probe.workbook.zoneId)}${
							probe.workbook.zoneId === null
								? ''
								: ` · ZoneId ${probe.workbook.zoneId}`
						  }`,
			impact: explicitInternetMacroBlock
				? 'Une stratégie explicite bloque les macros de ce fichier Internet.'
				: internetOrigin
					? 'Excel doit encore évaluer les stratégies, documents approuvés et éditeurs approuvés.'
					: 'Aucun blocage Internet certain n’est déduit pour ce fichier.',
			managed: false,
			source: 'Attribut NTFS Zone.Identifier'
		},
		{
			id: 'macros',
			title: 'Macros VBA',
			status: !macroCapable
				? 'notApplicable'
				: explicitInternetMacroBlock
					? 'blocked'
					: internetMacroDecisionUnknown
						? 'unknown'
					: macroDescription[0],
			detail: !macroCapable
				? 'Aucun projet VBA n’est détecté dans ce fichier.'
				: explicitInternetMacroBlock
					? 'Macros bloquées par la stratégie Internet détectée.'
					: internetMacroDecisionUnknown
						? 'Origine Internet détectée ; le diagnostic statique ne prouve pas la décision finale d’Excel.'
						: macroDescription[1],
			impact: macroCapable
				? 'Ce réglage détermine si Excel peut exécuter le VBA du classeur.'
				: 'Le format actuel ne contient pas de projet VBA détecté.',
			managed: macroResolution.managed || blockInternetResolution.managed,
			source: explicitInternetMacroBlock
				? `${sourceLabel(blockInternetResolution)} + origine du fichier`
				: internetMacroDecisionUnknown &&
					  numericValue(blockInternetSetting) === 1
					? `${sourceLabel(blockInternetResolution)} + confiance de l’éditeur à confirmer`
				: sourceLabel(macroResolution)
		},
		{
			id: 'accessVbom',
			title: 'Accès au projet VBA (AccessVBOM)',
			status: accessVbomResolution.conflict
				? 'unknown'
				: accessVbomAllowed
					? 'allowed'
					: 'blocked',
			detail: accessVbomResolution.conflict
				? 'Des valeurs AccessVBOM contradictoires sont visibles.'
				: accessVbomAllowed
				? 'Accès programmatique autorisé.'
				: accessVbomValue === 0
					? 'Accès programmatique désactivé.'
					: 'Aucune autorisation explicite détectée ; Office refuse cet accès par défaut.',
			impact: 'Nécessaire pour exporter le projet, créer des UserForms et utiliser le designer Excel contrôlé.',
			managed: accessVbomResolution.managed,
			source: sourceLabel(accessVbomResolution)
		},
		{
			id: 'activeX',
			title: 'Contrôles ActiveX',
			status: disableActiveXResolution.conflict || ufiControlsResolution.conflict
				? 'unknown'
				: activeXBlocked
					? 'blocked'
					: activeXValue === 0
						? 'prompt'
						: 'unknown',
			detail: disableActiveXResolution.conflict || ufiControlsResolution.conflict
				? 'Des valeurs ActiveX contradictoires sont visibles.'
				: activeXBlocked
				? 'Tous les contrôles ActiveX sont désactivés.'
				: activeXValue === 0
					? `ActiveX n’est pas globalement bloqué${
						ufiControlsSetting ? ` · UFIControls=${ufiControlsSetting.value}` : ''
					  }.`
					: 'Aucune exception locale détectée ; les versions Office récentes peuvent bloquer ActiveX par défaut.',
			impact: 'La création ou l’ouverture d’un contrôle reste soumise à Excel et à la stratégie de l’organisation.',
			managed: disableActiveXResolution.managed || ufiControlsResolution.managed,
			source: sourceLabel(
				disableActiveXResolution.candidates.length > 0
					? disableActiveXResolution
					: ufiControlsResolution
			)
		},
		{
			id: 'protectedView',
			title: 'Vue protégée',
			status: disableInternetPvResolution.conflict
				? 'unknown'
				: protectedViewDisabled
				? 'warning'
				: internetOrigin && !workbookInTrustedLocation
					? 'protected'
					: 'unknown',
			detail: disableInternetPvResolution.conflict
				? 'Des valeurs de Vue protégée contradictoires sont visibles.'
				: protectedViewDisabled
				? 'La Vue protégée des fichiers Internet est désactivée par configuration.'
				: internetOrigin && !workbookInTrustedLocation
					? 'Excel devrait appliquer la Vue protégée ou un blocage équivalent.'
					: 'Aucune décision certaine pour ce fichier.',
			impact: 'La décision finale appartient à Excel, Windows et aux stratégies Microsoft 365.',
			managed: disableInternetPvResolution.managed,
			source: sourceLabel(disableInternetPvResolution)
		},
		{
			id: 'trustedLocations',
			title: 'Emplacements approuvés',
			status: trustedLocationUncertain
				? 'unknown'
				: trustedLocationsDisabled
				? 'blocked'
				: matchingLocalLocation && userTrustedLocationsDisabled && !matchingManagedLocation
					? 'blocked'
				: workbookInTrustedLocation
					? 'allowed'
					: 'protected',
			detail: trustedLocationUncertain
				? 'La politique qui autorise le mélange des emplacements utilisateur et administrateur contient des valeurs contradictoires.'
				: trustedLocationsDisabled
				? 'Tous les emplacements approuvés sont désactivés.'
				: matchingLocalLocation && userTrustedLocationsDisabled && !matchingManagedLocation
					? 'Le dossier correspond à un emplacement utilisateur, mais la politique d’entreprise autorise uniquement les emplacements définis par stratégie.'
				: workbookInTrustedLocation
					? 'Le dossier du classeur correspond à un emplacement approuvé déclaré.'
					: `${probe.office.trustedLocations.length} emplacement(s) déclaré(s) ; ce fichier n’en fait pas partie.`,
			impact: 'Un emplacement approuvé contourne plusieurs contrôles Office et doit rester rare et protégé.',
			managed:
				disableAllTrustedResolution.managed ||
				allowUserTrustedResolution.managed ||
				probe.office.trustedLocations.some(location => location.managed),
			source: sourceLabel(
				disableAllTrustedResolution.candidates.length > 0
					? disableAllTrustedResolution
					: allowUserTrustedResolution
			)
		},
		{
			id: 'classification',
			title: 'Classification et chiffrement',
			status: encryptedPackage
				? 'blocked'
				: probe.workbook.efsEncrypted
					? 'protected'
				: probe.workbook.sensitivityLabelIds.length > 0
						? 'protected'
						: 'notApplicable',
			detail: encryptedPackage
				? 'Package Office chiffré : le contenu n’est pas statiquement lisible.'
				: probe.workbook.efsEncrypted
					? 'Chiffrement EFS détecté ; le fichier reste lisible pour l’utilisateur autorisé.'
					: probe.workbook.sensitivityLabelIds.length > 0
						? `${probe.workbook.sensitivityLabelIds.length} étiquette(s) Microsoft Purview détectée(s).`
						: 'Aucune étiquette de sensibilité lisible détectée.',
			impact: 'Le Centre détecte les métadonnées locales mais ne contacte jamais Microsoft Purview pour résoudre le nom des étiquettes.',
			managed: false,
			source: 'Métadonnées du fichier'
		},
		{
			id: 'signatures',
			title: 'Signatures numériques',
			status:
				probe.workbook.vbaSignatureStatus === 'present' ||
				packageSignaturePresent
				? 'protected'
				: probe.workbook.vbaSignatureStatus === 'unknown' ||
					  probe.workbook.packageSignatureStatus === 'unknown'
					? 'unknown'
					: macroCapable
						? 'warning'
						: 'notApplicable',
			detail: probe.workbook.vbaSignatureStatus === 'present'
				? 'Signature du projet VBA détectée ; l’extension refusera de modifier ce projet.'
				: packageSignaturePresent
					? 'Signature du package Office détectée ; toute modification invaliderait cette signature.'
					: probe.workbook.vbaSignatureStatus === 'unknown' ||
						  probe.workbook.packageSignatureStatus === 'unknown'
						? 'La présence d’une signature ne peut pas être déterminée statiquement pour ce format.'
						: 'Aucune signature lisible détectée.',
			impact: 'La présence est vérifiée localement ; la confiance du certificat et sa chaîne ne sont pas évaluées.',
			managed: false,
			source: 'Structure du fichier'
		},
		{
			id: 'vbaProtection',
			title: 'Protection du projet VBA',
			status: !macroCapable
				? 'notApplicable'
				: vbaProtectionPresent
					? 'blocked'
					: vbaProtectionUnknown
						? 'unknown'
						: 'protected',
			detail: !macroCapable
				? 'Aucun projet VBA n’est détecté.'
				: vbaProtectionPresent
					? 'Projet VBA protégé par mot de passe détecté.'
					: vbaProtectionUnknown
						? 'La protection est vérifiée par le moteur natif au moment d’une opération.'
						: 'Aucune protection par mot de passe détectée.',
			impact: 'Un projet protégé n’est jamais modifié par l’extension.',
			managed: false,
			source: 'Structure du projet VBA'
		}
	];
	let designerStatus: OfficeSecurityStatus;
	let designerDetail: string;
	if (!isXlsm) {
		designerStatus = 'blocked';
		designerDetail = 'Un classeur XLSM existant est requis.';
	} else if (encryptedPackage) {
		designerStatus = 'blocked';
		designerDetail = 'Package Office chiffré : designer indisponible.';
	} else if (probe.workbook.readOnly) {
		designerStatus = 'blocked';
		designerDetail = 'Attribut lecture seule : designer indisponible.';
	} else if (vbaSignaturePresent) {
		designerStatus = 'blocked';
		designerDetail = 'Projet VBA signé : modification refusée.';
	} else if (packageSignaturePresent) {
		designerStatus = 'blocked';
		designerDetail = 'Package Office signé : modification refusée.';
	} else if (vbaProtectionPresent) {
		designerStatus = 'blocked';
		designerDetail = 'Projet VBA protégé : modification refusée.';
	} else if (accessVbomResolution.conflict) {
		designerStatus = 'unknown';
		designerDetail = 'La valeur AccessVBOM effective doit être confirmée dans Excel.';
	} else if (!accessVbomAllowed) {
		designerStatus = 'blocked';
		designerDetail = 'AccessVBOM est désactivé.';
	} else if (
		vbaSignatureUnknown ||
		packageSignatureUnknown ||
		vbaProtectionUnknown
	) {
		designerStatus = 'prompt';
		designerDetail =
			'Les signatures et protections seront confirmées avant toute modification.';
	} else {
		designerStatus = 'allowed';
		designerDetail = 'Designer contrôlé disponible ; aucune macro exécutée.';
	}

	let activeXWriteStatus: OfficeSecurityStatus;
	let activeXWriteDetail: string;
	if (activeXBlocked) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'ActiveX bloqué par Office ou stratégie.';
	} else if (!isXlsm) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'Un classeur XLSM existant est requis.';
	} else if (immutableWorkbook) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'Le classeur est chiffré ou en lecture seule.';
	} else if (
		vbaSignaturePresent ||
		packageSignaturePresent ||
		vbaProtectionPresent
	) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'Le projet ou le package signé/protégé n’est pas modifié.';
	} else if (accessVbomResolution.conflict) {
		activeXWriteStatus = 'unknown';
		activeXWriteDetail =
			'La valeur AccessVBOM effective doit être confirmée dans Excel.';
	} else if (!accessVbomAllowed) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'AccessVBOM est requis pour lier les événements.';
	} else if (
		disableActiveXResolution.conflict ||
		vbaSignatureUnknown ||
		packageSignatureUnknown ||
		vbaProtectionUnknown
	) {
		activeXWriteStatus = 'unknown';
		activeXWriteDetail =
			'Les politiques et protections doivent être confirmées avant l’insertion.';
	} else {
		activeXWriteStatus = 'prompt';
		activeXWriteDetail =
			'Excel décide au moment de l’insertion ; les ProgID tiers restent sur liste d’autorisation.';
	}

	const capabilities: OfficeSecurityCapability[] = [
		{
			id: 'grid',
			title: 'Grille intégrée',
			status: !gridSupported
				? 'blocked'
				: encryptedPackage
				? 'blocked'
				: isLegacyXls
					? 'protected'
					: probe.workbook.readOnly
						? 'protected'
						: packageSignaturePresent
							? 'protected'
							: 'allowed',
			detail: !gridSupported
				? 'Ce format ne possède pas de grille intégrée modifiable.'
				: encryptedPackage
				? 'Contenu chiffré non lisible sans Excel et les droits appropriés.'
				: isLegacyXls
					? 'Le format XLS hérité est disponible uniquement en lecture protégée.'
					: probe.workbook.readOnly
						? 'Lecture possible ; écriture refusée par l’attribut lecture seule.'
						: packageSignaturePresent
							? 'Lecture possible ; toute édition invaliderait la signature du package.'
							: 'Lecture et édition selon les capacités du format.'
		},
		{
			id: 'vbaRead',
			title: 'Inspection VBA',
			status: encryptedPackage
				? 'blocked'
				: !macroCapable
				? 'notApplicable'
				: accessVbomResolution.conflict
					? 'unknown'
				: accessVbomAllowed
					? 'allowed'
					: 'blocked',
			detail: encryptedPackage
				? 'Le projet n’est pas lisible dans un package Office chiffré.'
				: !macroCapable
				? 'Aucun projet VBA n’est détecté.'
				: accessVbomResolution.conflict
					? 'La valeur AccessVBOM effective doit être confirmée dans Excel.'
				: accessVbomAllowed
				? 'Excel autorise l’export contrôlé du projet VBA.'
				: 'AccessVBOM doit être autorisé par l’utilisateur ou l’administrateur.'
		},
		{
			id: 'vbaWrite',
			title: 'Écriture VBA',
			status: immutableWorkbook ||
				vbaSignaturePresent ||
				packageSignaturePresent ||
				vbaProtectionPresent ||
				(!isXlsm && !isXlsx)
				? 'blocked'
				: isXlsx && accessVbomResolution.conflict
					? 'unknown'
					: isXlsx && !accessVbomAllowed
						? 'blocked'
						: vbaSignatureUnknown || packageSignatureUnknown || vbaProtectionUnknown
							? 'prompt'
							: 'allowed',
			detail: encryptedPackage
				? 'Package Office chiffré : écriture refusée.'
				: probe.workbook.readOnly
					? 'Attribut lecture seule : écriture refusée.'
					: vbaSignaturePresent
						? 'Projet signé : l’extension refuse toute modification.'
						: packageSignaturePresent
							? 'Package Office signé : écriture refusée pour ne pas invalider la signature.'
					: vbaProtectionPresent
							? 'Projet protégé par mot de passe : écriture refusée.'
							: !isXlsm && !isXlsx
								? 'Ce format ne prend pas en charge l’écriture VBA transactionnelle.'
								: isXlsx && accessVbomResolution.conflict
									? 'La valeur AccessVBOM effective doit être confirmée dans Excel.'
									: isXlsx && !accessVbomAllowed
										? 'La première conversion XLSX vers XLSM exige AccessVBOM.'
										: vbaSignatureUnknown || packageSignatureUnknown || vbaProtectionUnknown
											? 'Signature et protection seront vérifiées par le moteur natif avant toute écriture.'
											: 'Écriture transactionnelle possible sous réserve des validations du projet.'
		},
		{
			id: 'designer',
			title: 'UserForms et designer',
			status: designerStatus,
			detail: designerDetail
		},
		{
			id: 'activeXWrite',
			title: 'Création ActiveX',
			status: activeXWriteStatus,
			detail: activeXWriteDetail
		},
		{
			id: 'macroExecution',
			title: 'Exécution de macros par l’extension',
			status: 'blocked',
			detail: 'Toujours interdite, quel que soit le niveau Office.'
		}
	];

	const restricted =
		encryptedPackage ||
		probe.workbook.readOnly ||
		explicitInternetMacroBlock ||
		vbaSignaturePresent ||
		packageSignaturePresent ||
		vbaProtectionPresent ||
		Boolean(
			findings.find(
				finding => finding.managed && finding.status === 'blocked'
			)
		);
	const level = restricted
		? 'restricted'
		: managed
			? 'managed'
			: findings.some(finding => finding.status === 'unknown')
				? 'unknown'
				: 'standard';
	const summaries = {
		restricted: 'Des restrictions de sécurité s’appliquent à ce classeur.',
		managed: 'Une ou plusieurs règles sont gérées par l’organisation.',
		standard: 'Aucune restriction gérée supplémentaire n’a été détectée.',
		unknown: 'Certaines décisions restent à confirmer par Microsoft Excel.'
	};
	return {
		probe,
		level,
		summary: summaries[level],
		workbookInTrustedLocation,
		findings,
		capabilities
	};
}

function markdownValue(value: unknown): string {
	return String(value ?? '').replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}

export function formatEnterpriseSecurityReport(
	report: EnterpriseSecurityReport
): string {
	const lines = [
		'# Excel AI & VBA Studio — Diagnostic de sécurité',
		'',
		`- Classeur : ${report.probe.workbook.path}`,
		`- SHA-256 : ${report.probe.workbook.sha256}`,
		`- Niveau : ${report.level}`,
		`- Résumé : ${report.summary}`,
		`- Inspection UTC : ${report.probe.inspectedAtUtc || 'indéterminée'}`,
		'',
		'## Règles',
		'',
		'| Contrôle | État | Source | Détail |',
		'| --- | --- | --- | --- |',
		...report.findings.map(
			finding =>
				`| ${markdownValue(finding.title)} | ${finding.status} | ${markdownValue(
					finding.source
				)} | ${markdownValue(finding.detail)} |`
		),
		'',
		'## Capacités de l’extension',
		'',
		'| Capacité | État | Détail |',
		'| --- | --- | --- |',
		...report.capabilities.map(
			capability =>
				`| ${markdownValue(capability.title)} | ${capability.status} | ${markdownValue(
					capability.detail
				)} |`
		),
		'',
		'## Configuration Office détectée',
		'',
		'| Contrôle | Valeur | Source | Géré | Preuve locale |',
		'| --- | --- | --- | --- | --- |',
		...(report.probe.office.settings.length > 0
			? report.probe.office.settings.map(
					setting =>
						`| ${markdownValue(setting.id)} | ${markdownValue(
							setting.value
						)} | ${markdownValue(setting.source)} | ${setting.managed ? 'oui' : 'non'} | ${markdownValue(
							`${setting.registryPath}\\${setting.name}`
						)} |`
			  )
			: ['| Aucun réglage explicite |  |  |  |  |']),
		'',
		'## Emplacements approuvés détectés',
		'',
		'| Chemin | Sous-dossiers | Source | Géré |',
		'| --- | --- | --- | --- |',
		...(report.probe.office.trustedLocations.length > 0
			? report.probe.office.trustedLocations.map(
					location =>
						`| ${markdownValue(location.path)} | ${
							location.allowSubfolders ? 'oui' : 'non'
						} | ${markdownValue(location.source)} | ${location.managed ? 'oui' : 'non'} |`
			  )
			: ['| Aucun |  |  |  |']),
		'',
		'## Étiquettes Microsoft Purview détectées',
		'',
		...(report.probe.workbook.sensitivityLabelIds.length > 0
			? report.probe.workbook.sensitivityLabelIds.map(id => `- ${id}`)
			: ['- Aucune métadonnée d’étiquette lisible']),
		'',
		'> Rapport local en lecture seule. Aucune macro exécutée et aucun paramètre Office, registre ou stratégie modifié.'
	];
	return lines.join('\n');
}

export class OfficeSecurityService {
	constructor(
		private readonly extensionContext: vscode.ExtensionContext,
		private readonly outputChannel: vscode.OutputChannel
	) {}

	async inspect(workbookUri: vscode.Uri): Promise<EnterpriseSecurityReport> {
		if (process.platform !== 'win32') {
			throw new Error('Le Centre de sécurité Office est disponible uniquement sous Windows.');
		}
		const canonicalUri = await canonicalizeWorkbookUri(workbookUri);
		const scriptPath = this.extensionContext.asAbsolutePath(
			path.join('scripts', 'inspect-office-security.ps1')
		);
		const stdout = await new Promise<string>((resolve, reject) => {
			execFile(
				getPowerShellPath(),
				[
					'-NoLogo',
					'-NoProfile',
					'-NonInteractive',
					'-ExecutionPolicy',
					'Bypass',
					'-File',
					scriptPath,
					'-WorkbookPathBase64',
					Buffer.from(canonicalUri.fsPath, 'utf8').toString('base64')
				],
				{
					cwd: path.dirname(canonicalUri.fsPath),
					windowsHide: true,
					timeout: SECURITY_PROBE_TIMEOUT_MS,
					maxBuffer: SECURITY_PROBE_TRANSPORT_BYTES,
					encoding: 'utf8'
				},
				(error, output, stderr) => {
					if (error) {
						reject(
							new Error(
								`Diagnostic Office impossible : ${probeFailureMessage(
									String(output),
									String(stderr || ''),
									error.message
								)}`
							)
						);
						return;
					}
					resolve(String(output));
				}
			);
		});
		if (Buffer.byteLength(stdout, 'utf8') > MAX_PROBE_OUTPUT_BYTES) {
			throw new Error('Le diagnostic Office dépasse la taille maximale autorisée.');
		}
		const resultLine = stdout
			.replace(/\r/g, '')
			.split('\n')
			.map(line => line.trim())
			.filter(Boolean)
			.pop();
		if (!resultLine) {
			throw new Error('Le diagnostic Office n’a renvoyé aucun résultat.');
		}
		let raw: unknown;
		try {
			raw = JSON.parse(resultLine);
		} catch {
			throw new Error('Le diagnostic Office a renvoyé un JSON invalide.');
		}
		const probe = parseOfficeSecurityProbe(raw);
		if (
			path.resolve(probe.workbook.path).toLocaleLowerCase('en-US') !==
			path.resolve(canonicalUri.fsPath).toLocaleLowerCase('en-US')
		) {
			throw new Error('Le diagnostic Office a confirmé un autre classeur.');
		}
		this.outputChannel.appendLine(
			`[security] ${probe.workbook.name} : ${probe.office.settings.length} réglage(s), ${probe.office.trustedLocations.length} emplacement(s) approuvé(s)`
		);
		return buildEnterpriseSecurityReport(probe);
	}
}
