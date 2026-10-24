import { execFile } from 'child_process';
import { constants, promises as fs } from 'fs';
import * as path from 'path';

export type ExcelRegistration = 'detected' | 'missing' | 'unknown';
export interface SetupSnapshot {
	checkedAt: string;
	extensionVersion: string;
	vscodeVersion: string;
	platform: string;
	arch: string;
	trusted: boolean;
	powershell: boolean;
	nativeHelper: boolean;
	excel: ExcelRegistration;
	aiTools: boolean;
}
export interface SetupCheck {
	id: string;
	label: string;
	state: 'ok' | 'warning' | 'blocked' | 'unknown';
	status: string;
	detail: string;
}

// Read registration only. Never instantiate Excel, open a workbook or change policy.
export const EXCEL_REGISTRATION_PROBE = `
$ErrorActionPreference = 'Stop'
$found = $false
$failed = $false
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
  $root = $null
  $key = $null
  try {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, $view)
    $key = $root.OpenSubKey('Excel.Application\\CLSID', $false)
    if ($null -ne $key) {
      $guid = [Guid]::Empty
      if ([Guid]::TryParse([string]$key.GetValue(''), [ref]$guid) -and $guid -ne [Guid]::Empty) { $found = $true }
    }
  } catch { $failed = $true }
  finally {
    if ($null -ne $key) { $key.Dispose() }
    if ($null -ne $root) { $root.Dispose() }
  }
}
if ($found) { [Console]::Out.Write('detected') }
elseif ($failed) { [Console]::Out.Write('unknown') }
else { [Console]::Out.Write('missing') }
`;

export function parseExcelRegistration(output: string): ExcelRegistration {
	const value = output.trim();
	return value === 'detected' || value === 'missing' ? value : 'unknown';
}

async function readable(file: string): Promise<boolean> {
	try { await fs.access(file, constants.R_OK); return true; }
	catch { return false; }
}

export async function collectSetupSnapshot(
	extensionPath: string,
	info: Pick<SetupSnapshot, 'extensionVersion' | 'vscodeVersion' | 'trusted' | 'aiTools'>
): Promise<SetupSnapshot> {
	const powershellPath = process.env.SystemRoot
		? path.join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
		: '';
	const powershell = process.platform === 'win32' && !!powershellPath && await readable(powershellPath);
	const nativeHelper = await readable(path.join(extensionPath, 'bin', 'win32-x64', 'excel-ai-vba-writeback.exe'));
	const excel = powershell ? await new Promise<ExcelRegistration>(resolve => {
		execFile(powershellPath, [
			'-NoLogo', '-NoProfile', '-NonInteractive', '-EncodedCommand',
			Buffer.from(EXCEL_REGISTRATION_PROBE, 'utf16le').toString('base64')
		], { windowsHide: true, timeout: 5000, maxBuffer: 4096, encoding: 'utf8' }, (error, stdout) => {
			resolve(error ? 'unknown' : parseExcelRegistration(stdout));
		});
	}) : 'unknown';
	return { ...info, checkedAt: new Date().toISOString(), platform: process.platform,
		arch: process.arch, powershell, nativeHelper, excel };
}

export function buildSetupChecks(snapshot: SetupSnapshot): SetupCheck[] {
	const windows = snapshot.platform === 'win32' && snapshot.arch === 'x64';
	return [
		{ id: 'platform', label: 'Environnement', state: windows ? 'ok' : 'blocked',
			status: windows ? 'Windows x64' : 'Non pris en charge',
			detail: `VS Code ${snapshot.vscodeVersion} · ${snapshot.platform}/${snapshot.arch}. Le paquet est destiné à Windows x64, VS Code 1.95 ou ultérieur.` },
		{ id: 'trust', label: 'Espace de travail', state: snapshot.trusted ? 'ok' : 'blocked',
			status: snapshot.trusted ? 'Approuvé' : 'Mode restreint',
			detail: snapshot.trusted ? 'Les commandes peuvent être utilisées sur vos fichiers de confiance.' : 'Utilisez uniquement un dossier dont vous connaissez et approuvez le contenu.' },
		{ id: 'powershell', label: 'Windows PowerShell', state: snapshot.powershell ? 'ok' : 'warning',
			status: snapshot.powershell ? 'Présent' : 'Non détecté',
			detail: 'Nécessaire aux exports et aux opérations natives. Sa présence ne garantit pas que la politique de votre organisation autorise les scripts.' },
		{ id: 'excel', label: 'Microsoft Excel desktop', state: snapshot.excel === 'detected' ? 'ok' : snapshot.excel === 'missing' ? 'warning' : 'unknown',
			status: snapshot.excel === 'detected' ? 'Enregistrement détecté' : snapshot.excel === 'missing' ? 'Non détecté' : 'Non vérifié',
			detail: snapshot.excel === 'detected'
				? 'Signal du registre uniquement : Excel n’a pas été lancé. Le fonctionnement COM, la licence et les autorisations VBA restent à vérifier sur un classeur de test.'
				: 'La grille XLSX/CSV/TSV reste utilisable sans Excel. Les opérations natives XLSM, tableaux, graphiques et VBA nécessitent Excel desktop ; Excel web ne suffit pas.' },
		{ id: 'helper', label: 'Composant VBA embarqué', state: snapshot.nativeHelper ? 'ok' : 'warning',
			status: snapshot.nativeHelper ? 'Présent' : 'Absent',
			detail: snapshot.nativeHelper ? 'Le composant de réinjection est présent. Aucun code VBA n’a été exécuté.' : 'Réinstallez le VSIX Windows x64 complet avant d’utiliser la réinjection VBA.' },
		{ id: 'ai', label: 'Outils pour le chat IA', state: snapshot.aiTools ? 'ok' : 'warning',
			status: snapshot.aiTools ? '4 outils enregistrés' : 'Indisponibles',
			detail: 'L’édition du tableur ne nécessite pas Copilot. L’utilisation des outils IA demande un chat compatible et votre propre connexion ; celle-ci n’est pas testée ici.' }
	];
}

export function formatSetupReport(snapshot: SetupSnapshot): string {
	return [
		`Excel AI & VBA Studio ${snapshot.extensionVersion} — diagnostic local`,
		`Date : ${snapshot.checkedAt}`,
		...buildSetupChecks(snapshot).map(check => `${check.label} : ${check.status}\n  ${check.detail}`),
		'Limite : diagnostic de présence, pas une certification Excel/COM ni une recette de release.',
		'Aucun classeur lu, aucune macro exécutée, aucun réglage de sécurité modifié.',
		'Aucun chemin de fichier, contenu de classeur ou identifiant de connexion inclus.'
	].join('\n\n');
}
