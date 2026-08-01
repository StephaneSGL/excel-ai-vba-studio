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
export type OfficeArchitecture = 'x86' | 'x64' | 'unknown';
export type EnterpriseServiceDetectionStatus = 'detected' | 'notDetected' | 'unknown';
export type MdmProvider = 'microsoftIntune' | 'unknown' | 'none';

export interface PurviewSensitivityLabel {
	id: string;
	enabled: boolean;
	removed: boolean;
	name?: string;
	method?: string;
	setDate?: string;
	contentBits?: number;
	siteId?: string;
	source: 'customProperties' | 'labelInfoPart' | 'labelInfoStream';
	confidence: 'localDeclaration' | 'corroborated' | 'ambiguous';
}

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

export interface OfficeUnreadableSetting {
	id: string;
	source: OfficeSecuritySource;
	registryView?: '32' | '64';
}

export interface OfficeTrustedLocation {
	source: OfficeSecuritySource;
	managed: boolean;
	registryPath: string;
	path: string;
	allowSubfolders: boolean;
	description?: string;
	registryView?: '32' | '64';
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
		irmProtected: boolean;
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
		sensitivityLabels: PurviewSensitivityLabel[];
		sensitivityMetadataStatus: SecurityPresenceStatus;
		sensitivityMetadataSource:
			| 'none'
			| 'customProperties'
			| 'labelInfoPart'
			| 'labelInfoStream'
			| 'mixed'
			| 'encryptedContainer'
			| 'ambiguous'
			| 'unsupported';
	};
	office: {
		version: string;
		architecture: OfficeArchitecture;
		settings: OfficeSecuritySetting[];
		unreadableSettings: OfficeUnreadableSetting[];
		trustedLocations: OfficeTrustedLocation[];
		cloudPolicyDetected: boolean;
		cloudPolicyServiceDetected: boolean;
		intuneManagementExtensionDetected: boolean;
		mdmEnrollmentArtifactsDetected: boolean;
		groupPolicyHistoryDetected: boolean;
		cloudPolicyDetectionStatus: EnterpriseServiceDetectionStatus;
		cloudPolicyServiceStatus: EnterpriseServiceDetectionStatus;
		windowsPolicyRegistryStatus: EnterpriseServiceDetectionStatus;
		intuneManagementExtensionStatus: EnterpriseServiceDetectionStatus;
		mdmEnrollmentStatus: EnterpriseServiceDetectionStatus;
		mdmProvider: MdmProvider;
		groupPolicyHistoryStatus: EnterpriseServiceDetectionStatus;
		registryInspectionPartial: boolean;
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

export interface OfficePolicyDecision {
	id: string;
	title: string;
	state: 'effective' | 'conflict' | 'default' | 'unknown';
	value?: string | number | boolean | null;
	source: string;
	managed: boolean;
	applicableEvidenceCount: number;
	shadowedEvidenceCount: number;
}

export interface EnterpriseManagementService {
	id: 'cloudPolicy' | 'windowsPolicy' | 'intune' | 'mdm' | 'purview';
	title: string;
	status: EnterpriseServiceDetectionStatus;
	detail: string;
	limitation: string;
}

export interface EnterpriseSecurityReport {
	probe: OfficeSecurityProbe;
	level: 'restricted' | 'managed' | 'standard' | 'unknown';
	summary: string;
	workbookInTrustedLocation: boolean;
	findings: OfficeSecurityFinding[];
	capabilities: OfficeSecurityCapability[];
	policyDecisions: OfficePolicyDecision[];
	managementServices: EnterpriseManagementService[];
}

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

function asBoundedDisplayText(value: unknown, maximum: number): string {
	return asString(value)
		.replace(/[\u0000-\u001f\u007f]/g, ' ')
		.trim()
		.slice(0, maximum);
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

function asDetectionStatus(
	value: unknown,
	fallback: EnterpriseServiceDetectionStatus
): EnterpriseServiceDetectionStatus {
	return ['detected', 'notDetected', 'unknown'].includes(String(value))
		? (value as EnterpriseServiceDetectionStatus)
		: fallback;
}

function asZoneReadStatus(value: unknown, zoneId: number | null): ZoneReadStatus {
	return ['absent', 'read', 'unreadable', 'unsupported'].includes(String(value))
		? (value as ZoneReadStatus)
		: zoneId === null
			? 'absent'
			: 'read';
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
		registryView: registryPath.toLocaleUpperCase('en-US').startsWith('HKCU\\')
			? undefined
			: asString(record.registryView).slice(0, 100) || undefined
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
		description: asString(record.description).slice(0, 1000) || undefined,
		registryView: registryPath.toLocaleUpperCase('en-US').startsWith('HKCU\\')
			? undefined
			: ['32', '64'].includes(String(record.registryView))
			? (record.registryView as '32' | '64')
			: undefined
	};
}

function parseUnreadableSetting(value: unknown): OfficeUnreadableSetting | undefined {
	const record = asRecord(value);
	const source = securitySource(record?.source);
	const id = asString(record?.id).slice(0, 100);
	if (!record || !source || !id) {
		return undefined;
	}
	return {
		id,
		source,
		registryView: ['32', '64'].includes(String(record.registryView))
			? (record.registryView as '32' | '64')
			: undefined
	};
}

function parseSensitivityLabel(value: unknown): PurviewSensitivityLabel | undefined {
	const record = asRecord(value);
	const id = asString(record?.id).toLocaleLowerCase('en-US');
	const source = ['customProperties', 'labelInfoPart', 'labelInfoStream'].includes(
		String(record?.source)
	)
		? (record?.source as PurviewSensitivityLabel['source'])
		: undefined;
	const confidence = ['localDeclaration', 'corroborated', 'ambiguous'].includes(
		String(record?.confidence)
	)
		? (record?.confidence as PurviewSensitivityLabel['confidence'])
		: undefined;
	if (
		!/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(id) ||
		!source ||
		!confidence ||
		record?.enabled !== true ||
		record?.removed !== false
	) {
		return undefined;
	}
	const contentBits = asNullableInteger(record?.contentBits);
	const siteId = asString(record?.siteId).toLocaleLowerCase('en-US');
	const method = asBoundedDisplayText(record?.method, 128);
	if (
		!/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(siteId) ||
		(source === 'labelInfoPart' && !['Standard', 'Privileged'].includes(method))
	) {
		return undefined;
	}
	return {
		id,
		enabled: true,
		removed: false,
		name: asBoundedDisplayText(record?.name, 256) || undefined,
		method: method || undefined,
		setDate: asBoundedDisplayText(record?.setDate, 128) || undefined,
		contentBits:
			contentBits !== null && contentBits >= 0 && contentBits <= 0xffffffff
				? contentBits
				: undefined,
		siteId,
		source,
		confidence
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
	const parsedUnreadableSettings = Array.isArray(office.unreadableSettings)
		? office.unreadableSettings
				.map(parseUnreadableSetting)
				.filter((item): item is OfficeUnreadableSetting => Boolean(item))
				.slice(0, 256)
		: [];
	const parsedSensitivityLabels = Array.isArray(workbook.sensitivityLabels)
		? workbook.sensitivityLabels
				.map(parseSensitivityLabel)
				.filter((item): item is PurviewSensitivityLabel => Boolean(item))
				.slice(0, 32)
		: [];
	const sensitivityLabels = uniqueBy(parsedSensitivityLabels, label => label.id).slice(
		0,
		32
	);
	const reportedSensitivityStatus = asPresenceStatus(
		workbook.sensitivityMetadataStatus,
		sensitivityLabels.length > 0 ? 'present' : 'absent'
	);
	const sensitivityMetadataStatus =
		(reportedSensitivityStatus === 'present' && sensitivityLabels.length === 0) ||
		(reportedSensitivityStatus === 'absent' && sensitivityLabels.length > 0)
			? 'unknown'
			: reportedSensitivityStatus;
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
			irmProtected: asBoolean(workbook.irmProtected),
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
			sensitivityLabelIds: sensitivityLabels.map(label => label.id),
			sensitivityLabels,
			sensitivityMetadataStatus,
			sensitivityMetadataSource: [
				'none',
				'customProperties',
				'labelInfoPart',
				'labelInfoStream',
				'mixed',
				'encryptedContainer',
				'ambiguous',
				'unsupported'
			].includes(String(workbook.sensitivityMetadataSource))
				? (workbook.sensitivityMetadataSource as OfficeSecurityProbe['workbook']['sensitivityMetadataSource'])
				: sensitivityLabels.length > 0
					? 'customProperties'
					: 'none'
		},
		office: {
			version: asString(office.version, '16.0').slice(0, 20),
			architecture: ['x86', 'x64'].includes(String(office.architecture))
				? (office.architecture as OfficeArchitecture)
				: 'unknown',
			settings: uniqueBy(parsedSettings, setting =>
				[
					setting.source,
					setting.id,
					setting.registryPath.toLocaleLowerCase('en-US'),
					setting.name.toLocaleLowerCase('en-US'),
					JSON.stringify(setting.value),
					setting.registryView || 'any'
				].join('|')
			),
			unreadableSettings: uniqueBy(parsedUnreadableSettings, setting =>
				[setting.source, setting.id, setting.registryView || 'any'].join('|')
			),
			trustedLocations: uniqueBy(parsedTrustedLocations, location =>
				[
					location.source,
					location.registryPath.toLocaleLowerCase('en-US'),
					location.path.toLocaleLowerCase('en-US'),
					String(location.allowSubfolders),
					location.registryView || 'any'
				].join('|')
			),
			cloudPolicyDetected: asBoolean(office.cloudPolicyDetected),
			cloudPolicyServiceDetected: asBoolean(
				office.cloudPolicyServiceDetected
			),
			intuneManagementExtensionDetected: asBoolean(
				office.intuneManagementExtensionDetected
			),
			mdmEnrollmentArtifactsDetected: asBoolean(
				office.mdmEnrollmentArtifactsDetected
			),
			groupPolicyHistoryDetected: asBoolean(
				office.groupPolicyHistoryDetected
			),
			cloudPolicyDetectionStatus: asDetectionStatus(
				office.cloudPolicyDetectionStatus,
				asBoolean(office.cloudPolicyDetected) ? 'detected' : 'notDetected'
			),
			cloudPolicyServiceStatus: asDetectionStatus(
				office.cloudPolicyServiceStatus,
				asBoolean(office.cloudPolicyServiceDetected) ? 'detected' : 'notDetected'
			),
			windowsPolicyRegistryStatus: asDetectionStatus(
				office.windowsPolicyRegistryStatus,
				'notDetected'
			),
			intuneManagementExtensionStatus: asDetectionStatus(
				office.intuneManagementExtensionStatus,
				asBoolean(office.intuneManagementExtensionDetected)
					? 'detected'
					: 'notDetected'
			),
			mdmEnrollmentStatus: asDetectionStatus(
				office.mdmEnrollmentStatus,
				asBoolean(office.mdmEnrollmentArtifactsDetected) ? 'detected' : 'notDetected'
			),
			mdmProvider: ['microsoftIntune', 'unknown', 'none'].includes(
				String(office.mdmProvider)
			)
				? (office.mdmProvider as MdmProvider)
				: 'none',
			groupPolicyHistoryStatus: asDetectionStatus(
				office.groupPolicyHistoryStatus,
				asBoolean(office.groupPolicyHistoryDetected) ? 'detected' : 'notDetected'
			),
			registryInspectionPartial: asBoolean(office.registryInspectionPartial)
		}
	};
}

const POLICY_VALUE_DOMAINS: Record<string, ReadonlySet<number>> = {
	vbaWarnings: new Set([1, 2, 3, 4]),
	accessVbom: new Set([0, 1]),
	blockInternetMacros: new Set([0, 1]),
	xl4MacroOff: new Set([0, 1]),
	disableAllActiveX: new Set([0, 1]),
	disableInternetFilesInProtectedView: new Set([0, 1]),
	disableUnsafeLocationsInProtectedView: new Set([0, 1]),
	disableAttachmentsInProtectedView: new Set([0, 1]),
	disableAllTrustedLocations: new Set([0, 1]),
	allowNetworkTrustedLocations: new Set([0, 1]),
	allowUserTrustedLocations: new Set([0, 1])
};

function numericValue(setting: OfficeSecuritySetting | undefined): number | undefined {
	if (
		!setting ||
		(setting.valueKind || '').toLocaleLowerCase('en-US') !== 'dword'
	) {
		return undefined;
	}
	const parsed = Number(setting.value);
	if (!Number.isSafeInteger(parsed)) {
		return undefined;
	}
	const domain = POLICY_VALUE_DOMAINS[setting.id];
	return !domain || domain.has(parsed) ? parsed : undefined;
}

interface SettingResolution {
	setting?: OfficeSecuritySetting;
	candidates: OfficeSecuritySetting[];
	observedCandidates: OfficeSecuritySetting[];
	conflict: boolean;
	invalid: boolean;
	uncertain: boolean;
	unreadableSources: OfficeSecuritySource[];
	managed: boolean;
}

function resolveSetting(
	settings: OfficeSecuritySetting[],
	id: string,
	unreadableSettings: OfficeUnreadableSetting[] = []
): SettingResolution {
	const candidates = settings.filter(setting => setting.id === id);
	const unreadable = unreadableSettings.filter(setting => setting.id === id);
	const cloud = candidates.filter(setting => setting.source === 'cloudPolicy');
	const groupPolicy = candidates.filter(setting =>
		['machinePolicy', 'userPolicy'].includes(setting.source)
	);
	const preferences = candidates.filter(setting =>
		['userPreference', 'machinePreference'].includes(setting.source)
	);
	const cloudUnreadable = unreadable.some(setting => setting.source === 'cloudPolicy');
	const groupPolicyUnreadable = unreadable.some(setting =>
		['machinePolicy', 'userPolicy'].includes(setting.source)
	);
	const preferenceUnreadable = unreadable.some(setting =>
		['userPreference', 'machinePreference'].includes(setting.source)
	);
	let applicable: OfficeSecuritySetting[] = [];
	let uncertain = false;
	if (cloud.length > 0) {
		applicable = cloud;
		uncertain = cloudUnreadable;
	} else if (cloudUnreadable) {
		uncertain = true;
	} else if (groupPolicy.length > 0) {
		applicable = groupPolicy;
		uncertain = groupPolicyUnreadable;
	} else if (groupPolicyUnreadable) {
		uncertain = true;
	} else {
		applicable = preferences;
		uncertain = preferenceUnreadable;
	}
	const distinctValues = new Set(
		applicable.map(setting => JSON.stringify(setting.value))
	);
	const conflict = distinctValues.size > 1;
	const invalid = applicable.some(setting => numericValue(setting) === undefined);
	uncertain = uncertain || invalid;
	return {
		setting: conflict || uncertain ? undefined : applicable[0],
		candidates: applicable,
		observedCandidates: candidates,
		conflict,
		invalid,
		uncertain,
		unreadableSources: uniqueBy(unreadable, item => item.source).map(
			item => item.source
		),
		managed: applicable.some(setting => setting.managed)
	};
}

function sourceName(source: OfficeSecuritySource): string {
	switch (source) {
		case 'machinePolicy':
			return 'Registre de stratégie Windows · ordinateur';
		case 'userPolicy':
			return 'Registre de stratégie Windows · utilisateur';
		case 'cloudPolicy':
			return 'Stratégie Microsoft 365 Cloud Policy';
		case 'machinePreference':
			return 'Préférence locale · ordinateur';
		default:
			return 'Préférence locale · utilisateur';
	}
}

function sourceLabel(resolution: SettingResolution): string {
	if (resolution.conflict) {
		return 'Valeurs contradictoires détectées · décision Excel à confirmer';
	}
	if (resolution.invalid) {
		return 'Valeur ou type de registre invalide · décision Excel à confirmer';
	}
	if (resolution.uncertain) {
		const sources = resolution.unreadableSources.map(sourceName);
		return `Lecture refusée ou incomplète${sources.length > 0 ? ` · ${sources.join(' + ')}` : ''}`;
	}
	const setting = resolution.setting;
	if (!setting) {
		return 'Valeur Office par défaut ou indéterminée';
	}
	const sources = uniqueBy(resolution.candidates, candidate => candidate.source).map(
		candidate => sourceName(candidate.source)
	);
	return sources.length > 1 ? sources.join(' + ') : sourceName(setting.source);
}

function resolutionUnknown(resolution: SettingResolution): boolean {
	return resolution.conflict || resolution.uncertain;
}

const POLICY_CONTROLS: ReadonlyArray<{ id: string; title: string }> = [
	{ id: 'vbaWarnings', title: 'Niveau des macros VBA' },
	{ id: 'blockInternetMacros', title: 'Blocage des macros Internet' },
	{ id: 'xl4MacroOff', title: 'Désactivation des macros Excel 4.0 (XLM)' },
	{ id: 'accessVbom', title: 'Accès au modèle d’objet VBA' },
	{ id: 'disableAllActiveX', title: 'Désactivation globale ActiveX' },
	{ id: 'ufiControls', title: 'Initialisation des contrôles non sûrs' },
	{ id: 'disableInternetFilesInProtectedView', title: 'Vue protégée · Internet' },
	{ id: 'disableUnsafeLocationsInProtectedView', title: 'Vue protégée · emplacements non sûrs' },
	{ id: 'disableAttachmentsInProtectedView', title: 'Vue protégée · pièces jointes' },
	{ id: 'disableAllTrustedLocations', title: 'Désactivation des emplacements approuvés' },
	{ id: 'allowNetworkTrustedLocations', title: 'Emplacements réseau approuvés' },
	{ id: 'allowUserTrustedLocations', title: 'Emplacements approuvés utilisateur' }
];

function buildPolicyDecisions(
	settings: OfficeSecuritySetting[],
	unreadableSettings: OfficeUnreadableSetting[]
): OfficePolicyDecision[] {
	return POLICY_CONTROLS.map(control => {
		const resolution = resolveSetting(settings, control.id, unreadableSettings);
		return {
			id: control.id,
			title: control.title,
			state: resolution.conflict
				? 'conflict'
				: resolution.uncertain
					? 'unknown'
				: resolution.setting
					? 'effective'
					: 'default',
			value: resolution.setting?.value,
			source: sourceLabel(resolution),
			managed: resolution.managed,
			applicableEvidenceCount: resolution.candidates.length,
			shadowedEvidenceCount: Math.max(
				0,
				resolution.observedCandidates.length - resolution.candidates.length
			)
		};
	});
}

function buildManagementServices(
	probe: OfficeSecurityProbe
): EnterpriseManagementService[] {
	const windowsPolicySettings = uniqueBy(
		probe.office.settings.filter(setting =>
			['machinePolicy', 'userPolicy'].includes(setting.source)
		),
		setting =>
			[
				setting.source,
				setting.id,
				setting.registryPath.toLocaleLowerCase('en-US'),
				JSON.stringify(setting.value)
			].join('|')
	);
	const windowsPolicyLocations = uniqueBy(
		probe.office.trustedLocations.filter(location =>
			['machinePolicy', 'userPolicy'].includes(location.source)
		),
		location =>
			[
				location.source,
				location.registryPath.toLocaleLowerCase('en-US'),
				location.path.toLocaleLowerCase('en-US')
			].join('|')
	);
	const labelCount = probe.workbook.sensitivityLabels.length;
	const cloudStatus =
		probe.office.cloudPolicyDetectionStatus === 'detected' ||
		probe.office.cloudPolicyServiceStatus === 'detected'
			? 'detected'
			: probe.office.cloudPolicyDetectionStatus === 'unknown' ||
				  probe.office.cloudPolicyServiceStatus === 'unknown'
				? 'unknown'
				: 'notDetected';
	const windowsPolicyStatus =
		windowsPolicySettings.length > 0 ||
		windowsPolicyLocations.length > 0 ||
		probe.office.groupPolicyHistoryStatus === 'detected'
			? 'detected'
			: probe.office.windowsPolicyRegistryStatus === 'unknown' ||
				  probe.office.groupPolicyHistoryStatus === 'unknown'
				? 'unknown'
				: 'notDetected';
	const intuneStatus =
		probe.office.intuneManagementExtensionStatus === 'detected' ||
		(probe.office.mdmEnrollmentStatus === 'detected' &&
			probe.office.mdmProvider === 'microsoftIntune')
			? 'detected'
			: probe.office.intuneManagementExtensionStatus === 'unknown'
				? 'unknown'
				: 'notDetected';
	return [
		{
			id: 'cloudPolicy',
			title: 'Microsoft 365 Cloud Policy',
			status: cloudStatus,
			detail: probe.office.cloudPolicyDetected
				? 'Une ou plusieurs règles Office Cloud Policy ciblées sont présentes.'
				: probe.office.cloudPolicyServiceDetected
					? 'Le client Cloud Policy est visible, sans règle ciblée détectée.'
					: cloudStatus === 'unknown'
						? 'Au moins une lecture locale Cloud Policy a été refusée ou a échoué.'
						: 'Aucun signal Cloud Policy local n’a été trouvé.',
			limitation:
				'L’affectation au groupe Microsoft 365 et l’état du tenant ne sont pas interrogés.'
		},
		{
			id: 'windowsPolicy',
			title: 'Stratégies Windows gérées',
			status: windowsPolicyStatus,
			detail:
				windowsPolicySettings.length > 0 || windowsPolicyLocations.length > 0
					? `${windowsPolicySettings.length} réglage(s) Office et ${windowsPolicyLocations.length} emplacement(s) de stratégie observé(s).`
					: probe.office.groupPolicyHistoryStatus === 'detected'
						? 'Un historique de stratégie de groupe Windows est présent, sans règle Office ciblée.'
						: windowsPolicyStatus === 'unknown'
							? 'Une partie du registre de stratégie Windows n’a pas pu être lue.'
							: 'Aucune valeur Office ciblée ni historique GPO visible.',
			limitation:
				'Le registre prouve le magasin de stratégie, pas son canal de livraison : GPO, Intune/MDM, Configuration Manager ou script restent possibles.'
		},
		{
			id: 'intune',
			title: 'Microsoft Intune',
			status: intuneStatus,
			detail: probe.office.intuneManagementExtensionStatus === 'detected'
				? 'Le composant local Microsoft Intune Management Extension est détecté.'
				: probe.office.mdmEnrollmentStatus === 'detected' &&
					  probe.office.mdmProvider === 'microsoftIntune'
					? 'Une inscription MDM active pointe vers un service Microsoft Intune.'
					: intuneStatus === 'unknown'
						? 'La présence du composant Intune n’a pas pu être déterminée.'
						: 'Aucun signal local propre à Microsoft Intune n’a été trouvé.',
			limitation:
				'La présence du client ou du fournisseur ne prouve pas qu’Intune a livré une règle Office particulière.'
		},
		{
			id: 'mdm',
			title: 'Inscription MDM Windows',
			status: probe.office.mdmEnrollmentStatus,
			detail:
				probe.office.mdmEnrollmentStatus === 'detected'
					? probe.office.mdmProvider === 'microsoftIntune'
						? 'Une inscription MDM active corrélée est détectée ; le fournisseur local correspond à Microsoft Intune.'
						: 'Une inscription MDM active corrélée est détectée, sans fournisseur Microsoft confirmé.'
					: probe.office.mdmEnrollmentStatus === 'unknown'
						? 'Les clés d’inscription MDM n’ont pas toutes pu être lues de manière fiable.'
						: 'Aucune inscription MDM active corrélée n’a été trouvée.',
			limitation:
				'Le Centre ne lit ni identité utilisateur, ni identifiant d’appareil, ni contenu de certificat et ne contacte aucun service MDM.'
		},
		{
			id: 'purview',
			title: 'Microsoft Purview Information Protection',
			status:
				probe.workbook.sensitivityMetadataStatus === 'unknown'
					? 'unknown'
					: labelCount > 0
						? 'detected'
						: 'notDetected',
			detail:
				labelCount > 0
					? `${labelCount} métadonnée(s) locale(s) déclarant une étiquette de sensibilité ont été validées structurellement.`
					: probe.workbook.sensitivityMetadataStatus === 'unknown'
						? 'La présence ou l’état des métadonnées de sensibilité reste indéterminé.'
						: 'Aucune métadonnée d’étiquette applicable n’a été trouvée.',
			limitation:
				'Une déclaration locale n’authentifie ni le tenant, ni la politique actuelle, ni un niveau ordinal ; aucun catalogue Purview distant n’est interrogé.'
		}
	];
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

function isNetworkLocation(location: OfficeTrustedLocation): boolean {
	const expanded = expandWindowsEnvironment(location.path).trim();
	return /^\\\\/.test(expanded) || /^[a-z][a-z0-9+.-]*:\/\//i.test(expanded);
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
	const officeRegistryView = probe.office.architecture === 'x64'
		? '64'
		: probe.office.architecture === 'x86'
			? '32'
			: undefined;
	const settings = probe.office.settings.filter(
		setting =>
			!setting.registryView ||
			!officeRegistryView ||
			setting.registryView === officeRegistryView
	);
	const unreadableSettings = probe.office.unreadableSettings.filter(
		setting =>
			!setting.registryView ||
			!officeRegistryView ||
			setting.registryView === officeRegistryView
	);
	const trustedLocationsForOffice = probe.office.trustedLocations.filter(
		location =>
			!location.registryView ||
			!officeRegistryView ||
			location.registryView === officeRegistryView
	);
	const macroResolution = resolveSetting(settings, 'vbaWarnings', unreadableSettings);
	const accessVbomResolution = resolveSetting(settings, 'accessVbom', unreadableSettings);
	const blockInternetResolution = resolveSetting(
		settings,
		'blockInternetMacros',
		unreadableSettings
	);
	const xl4MacroResolution = resolveSetting(settings, 'xl4MacroOff', unreadableSettings);
	const disableActiveXResolution = resolveSetting(
		settings,
		'disableAllActiveX',
		unreadableSettings
	);
	const ufiControlsResolution = resolveSetting(settings, 'ufiControls', unreadableSettings);
	const disableInternetPvResolution = resolveSetting(
		settings,
		'disableInternetFilesInProtectedView',
		unreadableSettings
	);
	const disableAllTrustedResolution = resolveSetting(
		settings,
		'disableAllTrustedLocations',
		unreadableSettings
	);
	const allowUserTrustedResolution = resolveSetting(
		settings,
		'allowUserTrustedLocations',
		unreadableSettings
	);
	const allowNetworkTrustedResolution = resolveSetting(
		settings,
		'allowNetworkTrustedLocations',
		unreadableSettings
	);
	const macroSetting = macroResolution.setting;
	const accessVbomSetting = accessVbomResolution.setting;
	const blockInternetSetting = blockInternetResolution.setting;
	const xl4MacroSetting = xl4MacroResolution.setting;
	const disableActiveXSetting = disableActiveXResolution.setting;
	const ufiControlsSetting = ufiControlsResolution.setting;
	const disableInternetPv = disableInternetPvResolution.setting;
	const disableAllTrusted = disableAllTrustedResolution.setting;
	const allowUserTrusted = allowUserTrustedResolution.setting;
	const allowNetworkTrusted = allowNetworkTrustedResolution.setting;
	const allMatchingTrustedLocations = probe.office.trustedLocations.filter(location =>
		pathInside(probe.workbook.path, location)
	);
	const matchingNetworkLocation = allMatchingTrustedLocations.some(isNetworkLocation);
	const networkTrustedLocationsAllowed =
		!resolutionUnknown(allowNetworkTrustedResolution) &&
		numericValue(allowNetworkTrusted) === 1;
	const networkTrustedLocationUncertain =
		matchingNetworkLocation && resolutionUnknown(allowNetworkTrustedResolution);
	const networkTrustedLocationBlocked =
		matchingNetworkLocation &&
		!networkTrustedLocationUncertain &&
		!networkTrustedLocationsAllowed;
	const networkEligibleTrustedLocations = allMatchingTrustedLocations.filter(
		location => !isNetworkLocation(location) || networkTrustedLocationsAllowed
	);
	const dualViewTrustedMatch = networkEligibleTrustedLocations.some(
		location =>
			location.registryView === '32' &&
			networkEligibleTrustedLocations.some(
				candidate =>
					candidate.registryView === '64' &&
					candidate.source === location.source &&
					candidate.managed === location.managed &&
					candidate.registryPath.toLocaleLowerCase('en-US') ===
						location.registryPath.toLocaleLowerCase('en-US') &&
					candidate.path.toLocaleLowerCase('en-US') ===
						location.path.toLocaleLowerCase('en-US') &&
					candidate.allowSubfolders === location.allowSubfolders
			)
	);
	const trustedLocationArchitectureUncertain = Boolean(
		!officeRegistryView &&
			networkEligibleTrustedLocations.some(location => location.registryView) &&
			!networkEligibleTrustedLocations.some(location => !location.registryView) &&
			!dualViewTrustedMatch
	);
	const matchingTrustedLocations = trustedLocationArchitectureUncertain
		? []
		: trustedLocationsForOffice.filter(location =>
				pathInside(probe.workbook.path, location) &&
				(!isNetworkLocation(location) || networkTrustedLocationsAllowed)
		  );
	const matchingManagedLocation = matchingTrustedLocations.find(
		location => location.managed
	);
	const matchingLocalLocation = matchingTrustedLocations.find(
		location => !location.managed
	);
	const managed =
		settings.some(setting => setting.managed) ||
		Boolean(matchingManagedLocation) ||
		probe.office.cloudPolicyDetected;
	const trustedLocationsDisabled = numericValue(disableAllTrusted) === 1;
	const userTrustedLocationsDisabled =
		!resolutionUnknown(allowUserTrustedResolution) &&
		numericValue(allowUserTrusted) === 0;
	const trustedLocationUncertain =
		networkTrustedLocationUncertain ||
		trustedLocationArchitectureUncertain ||
		resolutionUnknown(disableAllTrustedResolution) ||
		Boolean(matchingLocalLocation && resolutionUnknown(allowUserTrustedResolution));
	const workbookInTrustedLocation = Boolean(
		!trustedLocationsDisabled &&
			!resolutionUnknown(disableAllTrustedResolution) &&
			(matchingManagedLocation ||
				(!userTrustedLocationsDisabled &&
					!resolutionUnknown(allowUserTrustedResolution) &&
					matchingLocalLocation))
	);
	const macroCapable = probe.workbook.hasVbaProject;
	const isXlsx = probe.workbook.extension === '.xlsx';
	const isXlsm = probe.workbook.extension === '.xlsm';
	const isLegacyXls = probe.workbook.extension === '.xls';
	const legacyXlmPotential = ['.xls', '.xlt'].includes(probe.workbook.extension);
	const legacyXlmBlocked =
		legacyXlmPotential &&
		!resolutionUnknown(xl4MacroResolution) &&
		numericValue(xl4MacroSetting) === 1;
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
			!resolutionUnknown(blockInternetResolution) &&
			numericValue(blockInternetSetting) === 1 &&
			probe.workbook.vbaSignatureStatus === 'absent'
	);
	const internetMacroDecisionUnknown = Boolean(
		macroCapable &&
			internetOrigin &&
			!workbookInTrustedLocation &&
			!explicitInternetMacroBlock &&
			(resolutionUnknown(blockInternetResolution) ||
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
	const macroDescription = resolutionUnknown(macroResolution)
		? (['unknown', 'La valeur de sécurité macro est contradictoire ou partiellement illisible.'] as const)
		: macroValue === undefined
		? (['prompt', 'Aucun réglage explicite détecté ; Excel applique sa valeur par défaut.'] as const)
		: macroDescriptions[macroValue] ||
			(['unknown', `Valeur VBAWarnings inconnue : ${macroValue}.`] as const);
	const accessVbomValue = numericValue(accessVbomSetting);
	const accessVbomAllowed = !resolutionUnknown(accessVbomResolution) && accessVbomValue === 1;
	const activeXValue = numericValue(disableActiveXSetting);
	const activeXBlocked =
		!resolutionUnknown(disableActiveXResolution) && activeXValue === 1;
	const protectedViewDisabled =
		!resolutionUnknown(disableInternetPvResolution) &&
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
				: workbookInTrustedLocation
					? 'warning'
				: explicitInternetMacroBlock
					? 'blocked'
					: internetMacroDecisionUnknown
						? 'unknown'
					: macroDescription[0],
			detail: !macroCapable
				? 'Aucun projet VBA n’est détecté dans ce fichier.'
				: workbookInTrustedLocation
					? 'Le classeur se trouve dans un emplacement approuvé effectif : Excel peut activer le contenu, indépendamment du niveau VBAWarnings.'
				: explicitInternetMacroBlock
					? 'Macros bloquées par la stratégie Internet détectée.'
					: internetMacroDecisionUnknown
						? 'Origine Internet détectée ; le diagnostic statique ne prouve pas la décision finale d’Excel.'
						: macroDescription[1],
			impact: macroCapable
				? workbookInTrustedLocation
					? 'Les emplacements approuvés contournent plusieurs contrôles Office et doivent être strictement protégés.'
					: 'Ce réglage détermine si Excel peut exécuter le VBA du classeur.'
				: 'Le format actuel ne contient pas de projet VBA détecté.',
			managed:
				macroResolution.managed ||
				(macroCapable && internetOrigin && blockInternetResolution.managed) ||
				Boolean(matchingManagedLocation),
			source: workbookInTrustedLocation && (matchingManagedLocation || matchingLocalLocation)
				? `${sourceName((matchingManagedLocation || matchingLocalLocation)!.source)} · emplacement approuvé`
				: explicitInternetMacroBlock
				? `${sourceLabel(blockInternetResolution)} + origine du fichier`
				: internetMacroDecisionUnknown &&
					  numericValue(blockInternetSetting) === 1
					? `${sourceLabel(blockInternetResolution)} + confiance de l’éditeur à confirmer`
					: sourceLabel(macroResolution)
		},
		{
			id: 'xlmMacros',
			title: 'Macros Excel 4.0 (XLM)',
			status: !legacyXlmPotential
				? 'notApplicable'
				: legacyXlmBlocked
					? 'blocked'
					: 'unknown',
			detail: !legacyXlmPotential
				? 'Le format de ce classeur ne peut pas contenir de feuille macro XLM héritée.'
				: legacyXlmBlocked
					? 'La stratégie XL4MacroOff désactive les macros Excel 4.0.'
					: resolutionUnknown(xl4MacroResolution)
						? 'La stratégie XL4MacroOff est contradictoire, invalide ou partiellement illisible.'
						: 'Le conteneur XLS/XLT peut contenir des feuilles macro XLM ; leur présence n’est pas déterminée par cette inspection statique.',
			impact: legacyXlmPotential
				? 'Ouvrez le fichier comme contenu actif potentiel tant que l’absence de feuilles macro XLM n’est pas confirmée par Excel.'
				: 'Sans objet pour ce format.',
			managed: xl4MacroResolution.managed,
			source: sourceLabel(xl4MacroResolution)
		},
		{
			id: 'accessVbom',
			title: 'Accès au projet VBA (AccessVBOM)',
			status: resolutionUnknown(accessVbomResolution)
				? 'unknown'
				: accessVbomAllowed
					? 'allowed'
					: 'blocked',
			detail: resolutionUnknown(accessVbomResolution)
				? 'La valeur AccessVBOM est contradictoire ou partiellement illisible.'
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
			status: resolutionUnknown(disableActiveXResolution) || resolutionUnknown(ufiControlsResolution)
				? 'unknown'
				: activeXBlocked
					? 'blocked'
					: activeXValue === 0
						? 'prompt'
						: 'unknown',
			detail: resolutionUnknown(disableActiveXResolution) || resolutionUnknown(ufiControlsResolution)
				? 'Les valeurs ActiveX sont contradictoires ou partiellement illisibles.'
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
			status: resolutionUnknown(disableInternetPvResolution)
				? 'unknown'
				: protectedViewDisabled
				? 'warning'
				: internetOrigin && !workbookInTrustedLocation
					? 'protected'
					: 'unknown',
			detail: resolutionUnknown(disableInternetPvResolution)
				? 'La valeur de Vue protégée est contradictoire ou partiellement illisible.'
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
				: networkTrustedLocationBlocked
					? 'blocked'
				: matchingLocalLocation && userTrustedLocationsDisabled && !matchingManagedLocation
					? 'blocked'
				: workbookInTrustedLocation
					? 'allowed'
					: 'protected',
			detail: trustedLocationUncertain
				? trustedLocationArchitectureUncertain
					? 'Le chemin correspond à un emplacement déclaré, mais l’architecture Office est inconnue ; la vue registre applicable ne peut pas être confirmée.'
					: 'La politique qui autorise le mélange des emplacements utilisateur et administrateur contient des valeurs contradictoires.'
				: trustedLocationsDisabled
				? 'Tous les emplacements approuvés sont désactivés.'
				: networkTrustedLocationBlocked
					? 'Le chemin correspond à un emplacement réseau déclaré, mais les emplacements réseau approuvés ne sont pas autorisés.'
				: matchingLocalLocation && userTrustedLocationsDisabled && !matchingManagedLocation
					? 'Le dossier correspond à un emplacement utilisateur, mais la politique d’entreprise autorise uniquement les emplacements définis par stratégie.'
				: workbookInTrustedLocation
					? 'Le dossier du classeur correspond à un emplacement approuvé déclaré.'
					: `${probe.office.trustedLocations.length} emplacement(s) déclaré(s) ; ce fichier n’en fait pas partie.`,
			impact: 'Un emplacement approuvé contourne plusieurs contrôles Office et doit rester rare et protégé.',
			managed:
				disableAllTrustedResolution.managed ||
				allowUserTrustedResolution.managed ||
				allowNetworkTrustedResolution.managed ||
				Boolean(matchingManagedLocation),
			source: matchingManagedLocation
				? `${sourceName(matchingManagedLocation.source)} · ${matchingManagedLocation.registryPath}`
				: matchingNetworkLocation
					? sourceLabel(allowNetworkTrustedResolution)
					: sourceLabel(
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
					: probe.workbook.sensitivityMetadataStatus === 'unknown'
						? 'unknown'
						: probe.workbook.sensitivityLabels.length > 0
							? 'warning'
							: 'notApplicable',
			detail: encryptedPackage
				? probe.workbook.irmProtected
					? 'Conteneur Office protégé par IRM : le contenu chiffré nécessite Excel et les droits attribués.'
					: 'Package Office chiffré : le contenu n’est pas statiquement lisible.'
				: probe.workbook.efsEncrypted
					? 'Chiffrement EFS détecté ; le fichier reste lisible pour l’utilisateur autorisé.'
					: probe.workbook.sensitivityMetadataStatus === 'unknown'
						? 'Les métadonnées de sensibilité sont ambiguës, illisibles ou non prises en charge pour ce conteneur.'
						: probe.workbook.sensitivityLabels.length > 0
							? `${probe.workbook.sensitivityLabels.length} métadonnée(s) locale(s) déclarant une étiquette Microsoft Purview.`
							: 'Aucune métadonnée d’étiquette applicable détectée.',
			impact: 'La classification déclarée est séparée du chiffrement réel. Elle ne prouve ni l’authenticité du tenant, ni la politique Purview actuellement publiée.',
			managed: false,
			source: `Métadonnées du fichier · ${probe.workbook.sensitivityMetadataSource}`
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
	} else if (resolutionUnknown(accessVbomResolution)) {
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
	} else if (resolutionUnknown(accessVbomResolution)) {
		activeXWriteStatus = 'unknown';
		activeXWriteDetail =
			'La valeur AccessVBOM effective doit être confirmée dans Excel.';
	} else if (!accessVbomAllowed) {
		activeXWriteStatus = 'blocked';
		activeXWriteDetail = 'AccessVBOM est requis pour lier les événements.';
	} else if (
		resolutionUnknown(disableActiveXResolution) ||
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
							: packageSignatureUnknown
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
								: packageSignatureUnknown
									? 'Lecture seulement : l’état de la signature du package doit être confirmé avant toute édition.'
								: 'Lecture et édition selon les capacités du format.'
		},
		{
			id: 'vbaRead',
			title: 'Inspection VBA',
			status: encryptedPackage
				? 'blocked'
				: !macroCapable
				? 'notApplicable'
				: resolutionUnknown(accessVbomResolution)
					? 'unknown'
				: accessVbomAllowed
					? 'allowed'
					: 'blocked',
			detail: encryptedPackage
				? 'Le projet n’est pas lisible dans un package Office chiffré.'
				: !macroCapable
				? 'Aucun projet VBA n’est détecté.'
				: resolutionUnknown(accessVbomResolution)
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
				: isXlsx && resolutionUnknown(accessVbomResolution)
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
								: isXlsx && resolutionUnknown(accessVbomResolution)
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
			: probe.office.registryInspectionPartial ||
				  findings.some(finding => finding.status === 'unknown')
				? 'unknown'
				: 'standard';
	const summaries = {
		restricted: 'Des restrictions de sécurité s’appliquent à ce classeur.',
		managed: 'Une ou plusieurs règles sont gérées par l’organisation.',
		standard: 'Aucune restriction gérée supplémentaire n’a été détectée.',
		unknown: 'Certaines décisions restent à confirmer par Microsoft Excel.'
	};
	const policyDecisions = buildPolicyDecisions(settings, unreadableSettings);
	const managementServices = buildManagementServices(probe);
	return {
		probe,
		level,
		summary: summaries[level],
		workbookInTrustedLocation,
		findings,
		capabilities,
		policyDecisions,
		managementServices
	};
}

function markdownValue(value: unknown): string {
	const encoded: Record<string, string> = {
		'&': '&amp;',
		'\\': '&#92;',
		'|': '&#124;',
		'`': '&#96;',
		'[': '&#91;',
		']': '&#93;',
		'<': '&lt;',
		'>': '&gt;',
		'(': '&#40;',
		')': '&#41;',
		':': '&#58;'
	};
	return String(value ?? '')
		.replace(/\r?\n/g, ' ')
		.replace(/[&\\|`\[\]<>()\:]/g, character => encoded[character]);
}

function contentBitsDescription(value: number | undefined): string {
	if (value === undefined) {
		return 'non déclaré';
	}
	const markings = [
		...(value & 1 ? ['en-tête'] : []),
		...(value & 2 ? ['pied de page'] : []),
		...(value & 4 ? ['filigrane'] : []),
		...(value & 8 ? ['chiffrement déclaré'] : [])
	];
	return markings.length > 0 ? `${value} (${markings.join(', ')})` : String(value);
}

export function formatEnterpriseSecurityReport(
	report: EnterpriseSecurityReport
): string {
	const lines = [
		'# Excel AI & VBA Studio — Diagnostic de sécurité',
		'',
		`- Classeur : ${markdownValue(report.probe.workbook.path)}`,
		`- SHA-256 : ${report.probe.workbook.sha256}`,
		`- Niveau : ${report.level}`,
		`- Résumé : ${markdownValue(report.summary)}`,
		`- Architecture Office détectée : ${report.probe.office.architecture}`,
		`- Inspection UTC : ${report.probe.inspectedAtUtc || 'indéterminée'}`,
		'',
		'## Services et canaux de gestion',
		'',
		'| Service | Détection | Preuve locale | Limite d’attribution |',
		'| --- | --- | --- | --- |',
		...report.managementServices.map(
			service =>
				`| ${markdownValue(service.title)} | ${service.status} | ${markdownValue(
					service.detail
				)} | ${markdownValue(service.limitation)} |`
		),
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
		'## Décisions Office effectives',
		'',
		'| Contrôle | Résolution | Valeur | Source effective | Verrouillé | Preuves remplacées |',
		'| --- | --- | --- | --- | --- | --- |',
		...report.policyDecisions.map(
			decision =>
				`| ${markdownValue(decision.title)} | ${decision.state} | ${markdownValue(
					decision.value === undefined ? 'défaut/indéterminé' : decision.value
				)} | ${markdownValue(decision.source)} | ${decision.managed ? 'oui' : 'non'} | ${
					decision.shadowedEvidenceCount
				} |`
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
		'| Chemin | Sous-dossiers | Vue registre | Source | Géré |',
		'| --- | --- | --- | --- | --- |',
		...(report.probe.office.trustedLocations.length > 0
			? report.probe.office.trustedLocations.map(
					location =>
						`| ${markdownValue(location.path)} | ${
							location.allowSubfolders ? 'oui' : 'non'
						} | ${location.registryView || 'toutes'} | ${markdownValue(location.source)} | ${location.managed ? 'oui' : 'non'} |`
			  )
			: ['| Aucun |  |  |  |  |']),
		'',
		'## Étiquettes Microsoft Purview détectées',
		'',
		`- État des métadonnées : ${report.probe.workbook.sensitivityMetadataStatus}`,
		`- Source : ${report.probe.workbook.sensitivityMetadataSource}`,
		'',
		...(report.probe.workbook.sensitivityLabels.length > 0
			? [
					'| Nom technique | ID étiquette | Tenant ID déclaré | Méthode déclarée | Date déclarée | Marquages déclarés | Source | Confiance |',
					'| --- | --- | --- | --- | --- | --- | --- | --- |',
					...report.probe.workbook.sensitivityLabels.map(
						label =>
							`| ${markdownValue(label.name || 'non enregistré')} | ${label.id} | ${markdownValue(
								label.siteId || 'non enregistré'
							)} | ${markdownValue(label.method || 'non enregistrée')} | ${markdownValue(
								label.setDate || 'non enregistrée'
							)} | ${markdownValue(contentBitsDescription(label.contentBits))} | ${label.source} | ${label.confidence} |`
					)
			  ]
			: ['- Aucune déclaration locale structurée d’étiquette applicable']),
		'',
		'> Une métadonnée locale n’authentifie pas le tenant et ne permet pas de déduire un classement ordinal hors ligne.',
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
